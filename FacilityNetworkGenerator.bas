Attribute VB_Name = "FacilityNetworkGenerator"
Option Explicit

' ============================================================
' FacilityNetworkGenerator.bas  v4.2
' Visio VBA module – generates Network Boundary, Riser, and
' Dataflow diagrams from a linked data recordset.
' ============================================================

' ----------------------------------------------------------
' Visio API constants
' ----------------------------------------------------------
Private Const visOpenHidden       As Long = 64
Private Const visLayerLock        As Long = 7
Private Const visLORouteRightAngle As Long = 1
Private Const visCharacterStyle   As Long = 2
Private Const visBold             As Long = 1
Private Const visHorzAlign        As Long = 6

' ----------------------------------------------------------
' Column name constants
' ----------------------------------------------------------
Private Const COL_LOCATION   As String = "Location"
Private Const COL_MODEL      As String = "Model #"
Private Const COL_LEVEL      As String = "DoD UFGS Purdue Level"
Private Const COL_UPSTREAM   As String = "Upstream Device"
Private Const COL_PROTOCOL   As String = "Control Protocol"
Private Const COL_MEDIA      As String = "Network Media Type"
Private Const COL_SUBTYPE    As String = "Device Sub-Type"
Private Const COL_IDENTIFIER As String = "Identifier"
Private Const COL_PORT       As String = "Port"

' ----------------------------------------------------------
' Layout / geometry constants (inches unless noted)
' ----------------------------------------------------------
Private Const BOX_LEFT_MARGIN       As Double = 0.375
Private Const RIGHT_MARGIN_RESERVE  As Double = 2.075
Private Const PAGE_TOP_BUFFER       As Double = 0.625
Private Const PAGE_BOTTOM_BUFFER    As Double = 0.625
Private Const LEFT_MARGIN_RESERVE   As Double = 3
Private Const X_OFFSET              As Double = 2.5
Private Const SWIMLANE_INTERNAL_GAP As Double = 0.1
Private Const BOX_ROUNDING          As String = "0.125 in"
Private Const MIN_ROW_HEIGHT        As Double = 2
Private Const ICON_Y_OFFSET         As Double = -0.85
Private Const TEXT_Y_OFFSET         As Double = 0.85
Private Const BOX_VERTICAL_OFFSET   As Double = 0.5
Private Const PAGE3_ARROW_STYLE     As String = "4"
Private Const CURVE_RADIUS          As String = "0.0625 in"
Private Const CONT_BANNER_HEIGHT    As Double = 0.4
Private Const OFFPAGE_REF_SIZE      As Double = 0.3
Private Const PAGE_NAME_SEPARATOR   As String = " – "
Private Const INF                   As Double = 1E+30

' ----------------------------------------------------------
' Master / shape name constants
' ----------------------------------------------------------
Private Const MASTER_SHAPE_NAME    As String = "Title"
Private Const MASTER_SHAPE_LEVEL_2 As String = "Level 2"
Private Const MASTER_SHAPE_LEVEL_1 As String = "Level 1"

' ----------------------------------------------------------
' Connector line-style constants
' ----------------------------------------------------------
Private Const LINESTYLE_DEFAULT  As Long = 6
Private Const LINESTYLE_ETHERNET As Long = 1
Private Const LINESTYLE_FIBER    As Long = 2
Private Const LINESTYLE_WIRELESS As Long = 3
Private Const LINESTYLE_SERIAL   As Long = 4

' ----------------------------------------------------------
' Connector color constants
' ----------------------------------------------------------
Private Const COLOR_ETHERNET As String = "RGB(0,90,156)"
Private Const COLOR_FIBER    As String = "RGB(227,108,10)"
Private Const COLOR_WIRELESS As String = "RGB(80,160,60)"
Private Const COLOR_SERIAL   As String = "RGB(150,50,150)"
Private Const COLOR_DEFAULT  As String = "RGB(80,80,80)"

' ----------------------------------------------------------
' PageType enum
' ----------------------------------------------------------
Public Enum PageType
    ptBoundary = 1
    ptRiser    = 2
    ptDataflow = 3
End Enum

' ----------------------------------------------------------
' ColumnIndices type
' ----------------------------------------------------------
Private Type ColumnIndices
    Location   As Long
    Model      As Long
    Level      As Long
    Upstream   As Long
    Protocol   As Long
    Media      As Long
    SubType    As Long
    Identifier As Long
    Port       As Long
End Type

' ----------------------------------------------------------
' PageChunk type  (used by DP pagination)
' ----------------------------------------------------------
Private Type PageChunk
    StartIdx As Long
    EndIdx   As Long
    Weight   As Double
End Type

' ----------------------------------------------------------
' Module-level cache variables
' ----------------------------------------------------------
Private dictMasterCache    As Object   ' Scripting.Dictionary
Private dictSubTypeMap     As Object   ' Scripting.Dictionary
Private dictShapeToPageName As Object  ' Scripting.Dictionary

' ============================================================
'  PUBLIC ENTRY POINTS
' ============================================================

Public Sub LaunchFacilityNetworkGenerator()
    NetworkSettings.ResetDefaults
    NetworkSettings.LogActiveSettings
    GenerateFacilityNetworkDiagrams
End Sub

Public Sub GenerateFacilityNetworkDiagrams()
    LogStep "GenerateFacilityNetworkDiagrams", "Start"

    Dim oDoc   As Object
    Dim oRS    As Object
    Dim oStencil As Object
    Dim cols   As ColumnIndices
    Dim nScope As Long

    On Error GoTo ErrHandler

    Set oDoc = Visio.ActiveDocument
    If oDoc Is Nothing Then
        MsgBox "No active Visio document found.", vbExclamation
        Exit Sub
    End If

    Set oRS = GetFirstRecordset(oDoc)
    If oRS Is Nothing Then
        MsgBox "No data recordset found in document.", vbExclamation
        Exit Sub
    End If

    Set oStencil = OpenStencil()
    If oStencil Is Nothing Then
        MsgBox "Could not open stencil: " & NetworkSettings.StencilFilename, vbExclamation
        Exit Sub
    End If

    InitSubTypeMap

    cols = MapColumns(oRS)

    If NetworkSettings.CleanupOldPages Then CleanupGeneratedPages oDoc

    nScope = oDoc.BeginUndoScope("Generate Facility Network Diagrams")

    EnableBatchMode True

    Dim buildingNames As Variant
    buildingNames = GetSortedBuildingNames(oRS, cols)

    Dim i As Long
    For i = LBound(buildingNames) To UBound(buildingNames)
        Dim bldg As String
        bldg = buildingNames(i)
        UpdateProgress "Processing building " & (i + 1) & " of " & _
                       (UBound(buildingNames) - LBound(buildingNames) + 1) & _
                       ": " & bldg
        ProcessBuilding oDoc, oRS, oStencil, cols, bldg
    Next i

    EnableBatchMode False
    oDoc.EndUndoScope nScope, True

    Application.StatusBar = "Facility Network Generator complete."
    LogStep "GenerateFacilityNetworkDiagrams", "Done"
    Exit Sub

ErrHandler:
    EnableBatchMode False
    If nScope <> 0 Then oDoc.EndUndoScope nScope, False
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    LogStep "GenerateFacilityNetworkDiagrams", "ERROR " & Err.Number & " – " & Err.Description
End Sub

' ============================================================
'  BUILDING PROCESSING
' ============================================================

Private Sub ProcessBuilding(ByVal oDoc As Object, ByVal oRS As Object, _
                             ByVal oStencil As Object, ByRef cols As ColumnIndices, _
                             ByVal buildingName As String)
    LogStep "ProcessBuilding", buildingName

    Dim dictIDtoModel    As Object
    Dim dictIDtoUpstream As Object
    Dim dictNameToID     As Object
    Dim dictBuildingGroups As Object
    Dim dictLaneRows     As Object
    Dim sortedLevelKeys  As Variant
    Dim totalWeight      As Double
    Dim maxWeightPerPage As Double

    BuildChainLookups oRS, cols, buildingName, dictIDtoModel, dictIDtoUpstream, dictNameToID
    Set dictBuildingGroups = GroupDevicesByLevelModelChain(oRS, cols, buildingName, _
                                dictIDtoModel, dictIDtoUpstream, dictNameToID)
    ComputeLaneLayout dictBuildingGroups, dictLaneRows, sortedLevelKeys, totalWeight

    ' --- Auto-Typical Overflow pre-check ---
    Dim templatePageH As Double
    templatePageH = 11   ' default letter height in inches
    On Error Resume Next
    Dim oTemplatePage As Object
    Set oTemplatePage = oDoc.Pages(1)
    If Not oTemplatePage Is Nothing Then
        templatePageH = oTemplatePage.PageSheet.Cells("PageHeight").ResultIU
    End If
    On Error GoTo 0

    Dim availableHeight As Double
    availableHeight = templatePageH - PAGE_TOP_BUFFER - PAGE_BOTTOM_BUFFER

    maxWeightPerPage = Int(availableHeight / MIN_ROW_HEIGHT)
    If maxWeightPerPage < 1 Then maxWeightPerPage = 1

    If NetworkSettings.AutoTypicalOverflow Then
        AutoTypicalOverflowLevels dictBuildingGroups, dictLaneRows, sortedLevelKeys, _
                                  totalWeight, dictIDtoModel, dictIDtoUpstream, _
                                  dictNameToID, maxWeightPerPage
    End If

    ' Recompute layout after potential overflow adjustment
    ComputeLaneLayout dictBuildingGroups, dictLaneRows, sortedLevelKeys, totalWeight

    ' --- Generate page sets ---
    If NetworkSettings.ShouldGeneratePage("BOUNDARY") Then
        GeneratePageSet oDoc, oStencil, cols, oRS, buildingName, ptBoundary, _
                        dictBuildingGroups, dictLaneRows, sortedLevelKeys, totalWeight, _
                        maxWeightPerPage, dictIDtoModel, dictIDtoUpstream, dictNameToID
    End If
    If NetworkSettings.ShouldGeneratePage("RISER") Then
        GeneratePageSet oDoc, oStencil, cols, oRS, buildingName, ptRiser, _
                        dictBuildingGroups, dictLaneRows, sortedLevelKeys, totalWeight, _
                        maxWeightPerPage, dictIDtoModel, dictIDtoUpstream, dictNameToID
    End If
    If NetworkSettings.ShouldGeneratePage("DATAFLOW") Then
        GeneratePageSet oDoc, oStencil, cols, oRS, buildingName, ptDataflow, _
                        dictBuildingGroups, dictLaneRows, sortedLevelKeys, totalWeight, _
                        maxWeightPerPage, dictIDtoModel, dictIDtoUpstream, dictNameToID
    End If

    CleanupCaches
End Sub

' ----------------------------------------------------------
' AutoTypicalOverflowLevels
' When a single level's row weight exceeds maxWeightPerPage,
' automatically regroups that level using chain-head (typical)
' grouping even if UseTypicalGrouping = False globally.
' ----------------------------------------------------------
Private Sub AutoTypicalOverflowLevels(ByRef dictBuildingGroups As Object, _
                                       ByRef dictLaneRows As Object, _
                                       ByRef sortedLevelKeys As Variant, _
                                       ByRef totalWeight As Double, _
                                       ByVal dictIDtoModel As Object, _
                                       ByVal dictIDtoUpstream As Object, _
                                       ByVal dictNameToID As Object, _
                                       ByVal maxWeightPerPage As Double)
    LogStep "AutoTypicalOverflowLevels", "maxWeightPerPage=" & maxWeightPerPage

    If Not IsArray(sortedLevelKeys) Then Exit Sub

    Dim k As Long
    For k = LBound(sortedLevelKeys) To UBound(sortedLevelKeys)
        Dim lvlKey As String
        lvlKey = sortedLevelKeys(k)

        Dim rowW As Double
        rowW = 0
        If dictLaneRows.Exists(lvlKey) Then rowW = dictLaneRows(lvlKey)

        If rowW > maxWeightPerPage Then
            LogStep "AutoTypicalOverflowLevels", "Overflow in level: " & lvlKey & _
                    " weight=" & rowW

            ' Collect all device records for this level
            Dim dictDevices As Object
            Set dictDevices = CreateObject("Scripting.Dictionary")

            ' First pass: collect device records and keys to remove
            Dim grpKey As Variant
            Dim keysToRemove() As String
            Dim removeCount As Long
            removeCount = 0
            ReDim keysToRemove(0 To dictBuildingGroups.Count - 1)
            For Each grpKey In dictBuildingGroups.Keys
                If Left(CStr(grpKey), Len(lvlKey) + 1) = lvlKey & "|" Then
                    Dim devList As Object
                    Set devList = dictBuildingGroups(grpKey)
                    Dim d As Long
                    For d = 1 To devList.Count
                        Dim rec As Object
                        Set rec = devList(d)
                        Dim devID As String
                        devID = CStr(rec("ID"))
                        If Not dictDevices.Exists(devID) Then
                            dictDevices.Add devID, rec
                        End If
                    Next d
                    keysToRemove(removeCount) = CStr(grpKey)
                    removeCount = removeCount + 1
                End If
            Next grpKey

            ' Second pass: remove collected keys (cannot modify dict during For Each)
            Dim ri As Long
            For ri = 0 To removeCount - 1
                dictBuildingGroups.Remove keysToRemove(ri)
            Next ri

            ' Re-add devices using chain-head grouping for this level
            Dim devKey As Variant
            For Each devKey In dictDevices.Keys
                Dim devRec As Object
                Set devRec = dictDevices(devKey)
                Dim modelStr  As String
                Dim chainHead As String
                modelStr  = SafeCleanString(devRec("Model"))
                chainHead = GetChainHeadID(CStr(devKey), dictIDtoUpstream, dictIDtoModel)

                Dim newGrpKey As String
                newGrpKey = lvlKey & "|" & modelStr & "|" & chainHead

                If Not dictBuildingGroups.Exists(newGrpKey) Then
                    Dim newList As Object
                    Set newList = New Collection
                    dictBuildingGroups.Add newGrpKey, newList
                End If
                dictBuildingGroups(newGrpKey).Add devRec
            Next devKey
        End If
    Next k

    ' Recompute lane rows and total weight after adjustments
    ComputeLaneLayout dictBuildingGroups, dictLaneRows, sortedLevelKeys, totalWeight
End Sub

' ============================================================
'  PAGE GENERATION
' ============================================================

Private Sub GeneratePageSet(ByVal oDoc As Object, ByVal oStencil As Object, _
                             ByRef cols As ColumnIndices, ByVal oRS As Object, _
                             ByVal buildingName As String, ByVal pType As PageType, _
                             ByVal dictBuildingGroups As Object, ByVal dictLaneRows As Object, _
                             ByVal sortedLevelKeys As Variant, ByVal totalWeight As Double, _
                             ByVal maxWeightPerPage As Double, _
                             ByVal dictIDtoModel As Object, ByVal dictIDtoUpstream As Object, _
                             ByVal dictNameToID As Object)

    LogStep "GeneratePageSet", BuildPageName(pType, buildingName)

    Select Case NetworkSettings.OverflowStrategy
        Case omContinuation
            ' Smart DP-based pagination
            Dim chunks() As PageChunk
            Dim costs()  As Double

            costs = BuildInterLevelCutCosts(sortedLevelKeys, dictBuildingGroups, _
                                            dictIDtoUpstream, dictNameToID)
            chunks = PartitionLevelsIntoChunks(sortedLevelKeys, dictLaneRows, costs, _
                                               maxWeightPerPage)

            If UBound(chunks) - LBound(chunks) + 1 <= 1 Then
                ' Single page path
                Dim dictNTS As Object
                Set dictNTS = CreateObject("Scripting.Dictionary")
                Set dictShapeToPageName = CreateObject("Scripting.Dictionary")
                GenerateSinglePage oDoc, oStencil, cols, oRS, buildingName, pType, _
                                   dictBuildingGroups, dictLaneRows, sortedLevelKeys, _
                                   dictNTS, dictIDtoModel, dictIDtoUpstream, dictNameToID, _
                                   1, 1
            Else
                ' Multi-page continuation path
                Set dictShapeToPageName = CreateObject("Scripting.Dictionary")
                Dim dictAllShapes As Object
                Set dictAllShapes = CreateObject("Scripting.Dictionary")

                Dim totalPages As Long
                totalPages = UBound(chunks) - LBound(chunks) + 1

                Dim ci As Long
                Dim pages() As Object
                ReDim pages(LBound(chunks) To UBound(chunks))

                For ci = LBound(chunks) To UBound(chunks)
                    Dim chunkKeys As Variant
                    ExtractChunkKeys sortedLevelKeys, chunks(ci), chunkKeys

                    Dim dictPageShapes As Object
                    Set dictPageShapes = CreateObject("Scripting.Dictionary")

                    Set pages(ci) = GenerateSinglePage(oDoc, oStencil, cols, oRS, _
                                       buildingName, pType, dictBuildingGroups, dictLaneRows, _
                                       chunkKeys, dictPageShapes, dictIDtoModel, _
                                       dictIDtoUpstream, dictNameToID, ci + 1, totalPages)

                    TrackShapePage dictPageShapes, pages(ci)
                    MergeShapeDicts dictAllShapes, dictPageShapes

                    If pages(ci) Is Nothing Then GoTo NextChunk
                    DrawContinuationBanner pages(ci), ci + 1, totalPages
NextChunk:
                Next ci

                ' Draw off-page references
                If NetworkSettings.ShowContRefs Then
                    DrawOffPageReferences dictAllShapes, dictIDtoUpstream, dictNameToID, _
                                         oStencil, pType
                End If
            End If

        Case omAutoExpand, omCompress
            GenerateSinglePageLegacy oDoc, oStencil, cols, oRS, buildingName, pType, _
                                     dictBuildingGroups, dictLaneRows, sortedLevelKeys, _
                                     dictIDtoModel, dictIDtoUpstream, dictNameToID
    End Select
End Sub

' ----------------------------------------------------------
' GenerateSinglePage – returns the generated Page object
' ----------------------------------------------------------
Private Function GenerateSinglePage(ByVal oDoc As Object, ByVal oStencil As Object, _
                                     ByRef cols As ColumnIndices, ByVal oRS As Object, _
                                     ByVal buildingName As String, ByVal pType As PageType, _
                                     ByVal dictBuildingGroups As Object, _
                                     ByVal dictLaneRows As Object, _
                                     ByVal levelKeys As Variant, _
                                     ByRef dictNameToShape As Object, _
                                     ByVal dictIDtoModel As Object, _
                                     ByVal dictIDtoUpstream As Object, _
                                     ByVal dictNameToID As Object, _
                                     ByVal pageNum As Long, _
                                     ByVal totalPages As Long) As Object
    LogStep "GenerateSinglePage", "page " & pageNum & " of " & totalPages

    Dim pageName As String
    pageName = UniquePageName(oDoc, BuildPageName(pType, buildingName), pageNum, totalPages)

    Dim oPage As Object
    Set oPage = oDoc.Pages.Add
    oPage.Name = pageName

    UpdateTitleBlock oPage, oStencil, buildingName, pType, pageNum, totalPages

    Dim lB As Double
    lB = PAGE_BOTTOM_BUFFER

    If IsArray(levelKeys) Then
        Dim ki As Long
        For ki = UBound(levelKeys) To LBound(levelKeys) Step -1
            lB = DrawLevelLane(oPage, oStencil, levelKeys(ki), dictBuildingGroups, _
                               dictLaneRows, dictNameToShape, dictIDtoModel, _
                               dictIDtoUpstream, dictNameToID, lB, pType)
        Next ki
    End If

    DrawUpstreamConnectors oPage, dictNameToShape, dictIDtoUpstream, dictNameToID, pType

    If pType = ptDataflow And NetworkSettings.ShowLegend Then
        DrawLegendBox oPage, dictNameToShape
    End If

    Set GenerateSinglePage = oPage
End Function

' ----------------------------------------------------------
' GenerateSinglePageLegacy – single page with auto-expand or compress
' ----------------------------------------------------------
Private Sub GenerateSinglePageLegacy(ByVal oDoc As Object, ByVal oStencil As Object, _
                                      ByRef cols As ColumnIndices, ByVal oRS As Object, _
                                      ByVal buildingName As String, ByVal pType As PageType, _
                                      ByVal dictBuildingGroups As Object, _
                                      ByVal dictLaneRows As Object, _
                                      ByVal sortedLevelKeys As Variant, _
                                      ByVal dictIDtoModel As Object, _
                                      ByVal dictIDtoUpstream As Object, _
                                      ByVal dictNameToID As Object)
    LogStep "GenerateSinglePageLegacy", BuildPageName(pType, buildingName)

    Dim dictNTS As Object
    Set dictNTS = CreateObject("Scripting.Dictionary")
    Set dictShapeToPageName = CreateObject("Scripting.Dictionary")

    Dim oPage As Object
    Set oPage = GenerateSinglePage(oDoc, oStencil, cols, oRS, buildingName, pType, _
                                   dictBuildingGroups, dictLaneRows, sortedLevelKeys, _
                                   dictNTS, dictIDtoModel, dictIDtoUpstream, dictNameToID, _
                                   1, 1)

    If NetworkSettings.OverflowStrategy = omAutoExpand And Not oPage Is Nothing Then
        ' Expand the page to fit all content
        On Error Resume Next
        Dim contentH As Double
        contentH = oPage.PageSheet.Cells("PageHeight").ResultIU
        If contentH < 11 Then contentH = 11
        oPage.PageSheet.Cells("PageHeight").FormulaU = CStr(contentH) & " in"
        On Error GoTo 0
    End If
End Sub

' ============================================================
'  SMART PAGINATION (DP-BASED)
' ============================================================

Private Function BuildInterLevelCutCosts(ByVal sortedLevelKeys As Variant, _
                                          ByVal dictBuildingGroups As Object, _
                                          ByVal dictIDtoUpstream As Object, _
                                          ByVal dictNameToID As Object) As Double()
    Dim n As Long
    n = UBound(sortedLevelKeys) - LBound(sortedLevelKeys) + 1

    Dim costs() As Double

    If n < 2 Then
        BuildInterLevelCutCosts = costs
        Exit Function
    End If

    ReDim costs(0 To n - 2)   ' cost of cutting between level i and i+1

    Dim i As Long
    For i = 0 To n - 2
        Dim lvlA As String
        Dim lvlB As String
        lvlA = sortedLevelKeys(LBound(sortedLevelKeys) + i)
        lvlB = sortedLevelKeys(LBound(sortedLevelKeys) + i + 1)

        Dim cutCost As Double
        cutCost = 0

        ' Count cross-level connections between lvlA and lvlB
        Dim grpKey As Variant
        For Each grpKey In dictBuildingGroups.Keys
            If Left(CStr(grpKey), Len(lvlA) + 1) = lvlA & "|" Then
                Dim devList As Object
                Set devList = dictBuildingGroups(grpKey)
                Dim d As Long
                For d = 1 To devList.Count
                    Dim rec As Object
                    Set rec = devList(d)
                    Dim upID As String
                    upID = SafeCleanString(rec("Upstream"))
                    If dictNameToID.Exists(upID) Then
                        Dim upDevID As String
                        upDevID = CStr(dictNameToID(upID))
                        ' Check if upstream device is in lvlB
                        Dim grpKey2 As Variant
                        For Each grpKey2 In dictBuildingGroups.Keys
                            If Left(CStr(grpKey2), Len(lvlB) + 1) = lvlB & "|" Then
                                Dim devList2 As Object
                                Set devList2 = dictBuildingGroups(grpKey2)
                                Dim d2 As Long
                                For d2 = 1 To devList2.Count
                                    Dim rec2 As Object
                                    Set rec2 = devList2(d2)
                                    If CStr(rec2("ID")) = upDevID Then
                                        cutCost = cutCost + 1
                                    End If
                                Next d2
                            End If
                        Next grpKey2
                    End If
                Next d
            End If
        Next grpKey

        costs(i) = cutCost
    Next i

    BuildInterLevelCutCosts = costs
End Function

Private Function PartitionLevelsIntoChunks(ByVal sortedLevelKeys As Variant, _
                                             ByVal dictLaneRows As Object, _
                                             ByVal costs() As Double, _
                                             ByVal maxW As Double) As PageChunk()
    Dim n As Long
    n = UBound(sortedLevelKeys) - LBound(sortedLevelKeys) + 1

    If n = 0 Then
        Dim empty() As PageChunk
        ReDim empty(0 To 0)
        empty(0).StartIdx = 0
        empty(0).EndIdx   = -1
        empty(0).Weight   = 0
        PartitionLevelsIntoChunks = empty
        Exit Function
    End If

    ' weights(i) = row weight of level i
    Dim weights() As Double
    ReDim weights(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1
        Dim lk As String
        lk = sortedLevelKeys(LBound(sortedLevelKeys) + i)
        If dictLaneRows.Exists(lk) Then
            weights(i) = dictLaneRows(lk)
        Else
            weights(i) = 1
        End If
    Next i

    ' dp(i)  = minimum cost to partition levels 0..i
    ' cut(i) = best last-cut position before i
    Dim dp()  As Double
    Dim cut() As Long
    ReDim dp(0 To n)
    ReDim cut(0 To n)

    Dim j As Long
    For i = 0 To n
        dp(i) = INF
        cut(i) = -1
    Next i
    dp(0) = 0

    For j = 1 To n
        Dim pageW As Double
        pageW = 0
        For i = j To 1 Step -1
            pageW = pageW + weights(i - 1)
            If pageW > maxW And (j - i) > 0 Then Exit For ' can't fit
            Dim extraCost As Double
            extraCost = 0
            If i > 1 Then extraCost = costs(i - 2)  ' cost of cut before level i
            Dim candidate As Double
            candidate = dp(i - 1) + extraCost
            If pageW <= maxW Then
                If candidate < dp(j) Then
                    dp(j) = candidate
                    cut(j) = i - 1
                End If
            End If
        Next i
    Next j

    ' --- Fallback: if dp(n) is still INF, a single level exceeds maxW ---
    ' Fall back to 1 level per page
    If dp(n) >= INF Then
        LogStep "PartitionLevelsIntoChunks", "INF bestCost – falling back to 1 level per page"
        Dim fallback() As PageChunk
        ReDim fallback(0 To n - 1)
        For i = 0 To n - 1
            fallback(i).StartIdx = i
            fallback(i).EndIdx   = i
            fallback(i).Weight   = weights(i)
        Next i
        PartitionLevelsIntoChunks = fallback
        Exit Function
    End If

    ' Reconstruct chunks by tracing back through cut()
    Dim chunkList() As PageChunk
    ReDim chunkList(0 To n - 1)
    Dim numChunks As Long
    numChunks = 0

    j = n
    Do While j > 0
        Dim startI As Long
        startI = cut(j)
        chunkList(numChunks).StartIdx = startI
        chunkList(numChunks).EndIdx   = j - 1
        Dim w As Double
        w = 0
        Dim ci As Long
        For ci = startI To j - 1
            w = w + weights(ci)
        Next ci
        chunkList(numChunks).Weight = w
        numChunks = numChunks + 1
        j = startI
    Loop

    ' Reverse the chunk list (we built it back-to-front)
    Dim result() As PageChunk
    ReDim result(0 To numChunks - 1)
    For i = 0 To numChunks - 1
        result(i) = chunkList(numChunks - 1 - i)
    Next i

    PartitionLevelsIntoChunks = result
End Function

Private Sub ExtractChunkKeys(ByVal sortedLevelKeys As Variant, _
                              ByRef chunk As PageChunk, _
                              ByRef chunkKeys As Variant)
    Dim cnt As Long
    cnt = chunk.EndIdx - chunk.StartIdx + 1
    If cnt <= 0 Then
        chunkKeys = Array()
        Exit Sub
    End If

    ReDim chunkKeys(0 To cnt - 1)
    Dim i As Long
    For i = 0 To cnt - 1
        chunkKeys(i) = sortedLevelKeys(LBound(sortedLevelKeys) + chunk.StartIdx + i)
    Next i
End Sub

' ============================================================
'  CONTINUATION PAGES
' ============================================================

Private Sub DrawContinuationBanner(ByVal oPage As Object, _
                                    ByVal pageNum As Long, _
                                    ByVal totalPages As Long)
    If oPage Is Nothing Then Exit Sub
    On Error Resume Next

    Dim oShape As Object
    Dim pageW  As Double
    pageW = oPage.PageSheet.Cells("PageWidth").ResultIU

    Set oShape = oPage.DrawRectangle(0, 0, pageW, CONT_BANNER_HEIGHT)
    oShape.Text = "Page " & pageNum & " of " & totalPages
    oShape.CellsU("FillForegnd").FormulaU = "RGB(220,230,245)"
    oShape.CellsU("LineColor").FormulaU   = "RGB(180,180,180)"
    oShape.CellsU("VerticalAlign").FormulaU = "1"
    On Error GoTo 0
End Sub

Private Sub MergeShapeDicts(ByRef dictTarget As Object, ByVal dictSource As Object)
    If dictSource Is Nothing Then Exit Sub
    Dim k As Variant
    For Each k In dictSource.Keys
        If Not dictTarget.Exists(k) Then
            dictTarget.Add k, dictSource(k)
        End If
    Next k
End Sub

Private Sub TrackShapePage(ByVal dictPageShapes As Object, ByVal oPage As Object)
    If dictPageShapes Is Nothing Then Exit Sub
    If oPage Is Nothing Then Exit Sub
    Dim k As Variant
    For Each k In dictPageShapes.Keys
        If Not dictShapeToPageName Is Nothing Then
            If Not dictShapeToPageName.Exists(k) Then
                dictShapeToPageName.Add k, oPage.Name
            End If
        End If
    Next k
End Sub

Private Sub DrawOffPageReferences(ByVal dictAllShapes As Object, _
                                   ByVal dictIDtoUpstream As Object, _
                                   ByVal dictNameToID As Object, _
                                   ByVal oStencil As Object, _
                                   ByVal pType As PageType)
    If dictAllShapes Is Nothing Then Exit Sub
    If dictShapeToPageName Is Nothing Then Exit Sub

    Dim shapeKey As Variant
    For Each shapeKey In dictAllShapes.Keys
        Dim oShape As Object
        Set oShape = dictAllShapes(shapeKey)
        If oShape Is Nothing Then GoTo NextShape

        Dim upstreamName As String
        On Error Resume Next
        upstreamName = SafeCleanString(oShape.CellsU("Prop.UpstreamDevice").ResultStr(""))
        On Error GoTo 0
        If upstreamName = "" Then GoTo NextShape

        Dim upstreamKey As String
        upstreamKey = upstreamName

        If Not dictAllShapes.Exists(upstreamKey) Then GoTo NextShape

        Dim oUpstream As Object
        Set oUpstream = dictAllShapes(upstreamKey)
        If oUpstream Is Nothing Then GoTo NextShape

        If Not AreSamePage(oShape, oUpstream) Then
            DrawSingleOffPageRef oShape, oUpstream, oStencil, pType
        End If

NextShape:
    Next shapeKey
End Sub

Private Sub DrawSingleOffPageRef(ByVal oFrom As Object, ByVal oTo As Object, _
                                  ByVal oStencil As Object, ByVal pType As PageType)
    If oFrom Is Nothing Or oTo Is Nothing Then Exit Sub
    On Error Resume Next

    ' CRITICAL: Use ChrW() for Unicode arrows (not Chr())
    Dim arrowUp   As String
    Dim arrowDown As String
    arrowUp   = ChrW(9650)   ' ▲
    arrowDown = ChrW(9660)   ' ▼

    ' Determine direction by comparing page indices
    Dim fromPageName As String
    Dim toPageName   As String
    fromPageName = oFrom.ContainingPage.Name
    toPageName   = oTo.ContainingPage.Name

    Dim fromIdx As Long
    Dim toIdx   As Long
    fromIdx = oFrom.ContainingPage.Index
    toIdx   = oTo.ContainingPage.Index

    Dim arrowStr As String
    If toIdx < fromIdx Then
        arrowStr = arrowDown
    Else
        arrowStr = arrowUp
    End If

    ' Draw a small reference marker near the source shape
    Dim oFromPage As Object
    Set oFromPage = oFrom.ContainingPage

    Dim refX As Double
    Dim refY As Double
    refX = oFrom.CellsU("PinX").ResultIU + OFFPAGE_REF_SIZE
    refY = oFrom.CellsU("PinY").ResultIU

    Dim oRef As Object
    Set oRef = oFromPage.DrawRectangle(refX - OFFPAGE_REF_SIZE / 2, _
                                        refY - OFFPAGE_REF_SIZE / 2, _
                                        refX + OFFPAGE_REF_SIZE / 2, _
                                        refY + OFFPAGE_REF_SIZE / 2)
    oRef.Text = arrowStr & " " & toPageName
    oRef.CellsU("FillForegnd").FormulaU = "RGB(255,255,200)"
    oRef.CellsU("LineColor").FormulaU   = "RGB(180,150,0)"

    On Error GoTo 0
End Sub

' ============================================================
'  LEVEL LANE DRAWING
' ============================================================

Private Function DrawLevelLane(ByVal oPage As Object, ByVal oStencil As Object, _
                                ByVal levelKey As String, _
                                ByVal dictBuildingGroups As Object, _
                                ByVal dictLaneRows As Object, _
                                ByRef dictNameToShape As Object, _
                                ByVal dictIDtoModel As Object, _
                                ByVal dictIDtoUpstream As Object, _
                                ByVal dictNameToID As Object, _
                                ByVal lB As Double, _
                                ByVal pType As PageType) As Double
    ' Empty guard
    Dim levelGroups As Object
    Set levelGroups = CreateObject("Scripting.Dictionary")

    Dim grpKey As Variant
    For Each grpKey In dictBuildingGroups.Keys
        If Left(CStr(grpKey), Len(levelKey) + 1) = levelKey & "|" Then
            levelGroups.Add grpKey, dictBuildingGroups(grpKey)
        End If
    Next grpKey

    If levelGroups.Count = 0 Then
        DrawLevelLane = lB
        Exit Function
    End If

    Dim rowWeight As Double
    rowWeight = 1
    If dictLaneRows.Exists(levelKey) Then rowWeight = dictLaneRows(levelKey)

    Dim laneH As Double
    laneH = rowWeight * MIN_ROW_HEIGHT

    Dim lT As Double
    lT = lB + laneH

    ' Draw swimlane background
    DrawSwimLaneBox oPage, levelKey, lB, lT, pType

    ' Sort group keys by upstream gravity
    Dim optKeys As Variant
    optKeys = GetGravitySortedKeys(levelGroups, dictNameToShape)

    ' Array guard
    If Not IsArray(optKeys) Then
        DrawLevelLane = lT
        Exit Function
    End If

    Dim pageW As Double
    pageW = oPage.PageSheet.Cells("PageWidth").ResultIU

    Dim xPos As Double
    xPos = LEFT_MARGIN_RESERVE + BOX_LEFT_MARGIN

    Dim ki As Long
    For ki = LBound(optKeys) To UBound(optKeys)
        Dim gKey As String
        gKey = CStr(optKeys(ki))

        If levelGroups.Exists(gKey) Then
            Dim devList As Object
            Set devList = levelGroups(gKey)

            Dim centY As Double
            centY = lB + laneH / 2

            xPos = PlaceDeviceGroup(oPage, oStencil, gKey, devList, xPos, centY, _
                                    dictNameToShape, dictIDtoModel, dictIDtoUpstream, _
                                    dictNameToID, pType, lB, lT)
        End If
    Next ki

    DrawLevelLane = lT
End Function

' ============================================================
'  DEVICE GROUP PLACEMENT
' ============================================================

Private Function PlaceDeviceGroup(ByVal oPage As Object, ByVal oStencil As Object, _
                                   ByVal groupKey As String, ByVal devList As Object, _
                                   ByVal xPos As Double, ByVal centY As Double, _
                                   ByRef dictNameToShape As Object, _
                                   ByVal dictIDtoModel As Object, _
                                   ByVal dictIDtoUpstream As Object, _
                                   ByVal dictNameToID As Object, _
                                   ByVal pType As PageType, _
                                   ByVal lB As Double, ByVal lT As Double) As Double
    If devList Is Nothing Then
        PlaceDeviceGroup = xPos
        Exit Function
    End If

    Dim qty As Long
    qty = devList.Count
    If qty = 0 Then
        PlaceDeviceGroup = xPos
        Exit Function
    End If

    ' Get model info from first device in group
    Dim firstRec As Object
    Set firstRec = devList(1)
    Dim modelStr  As String
    Dim subType   As String
    modelStr = SafeCleanString(firstRec("Model"))
    subType  = SafeCleanString(firstRec("SubType"))

    ' Box dimensions
    Dim boxW As Double
    Dim boxH As Double
    boxW = X_OFFSET
    boxH = lT - lB - 2 * SWIMLANE_INTERNAL_GAP

    Dim boxLeft   As Double
    Dim boxBottom As Double
    boxLeft   = xPos
    boxBottom = lB + SWIMLANE_INTERNAL_GAP

    ' Draw enclosing box
    Dim oBox As Object
    Set oBox = PickBoxMaster(oPage, oStencil, pType, firstRec)
    If oBox Is Nothing Then
        Set oBox = oPage.DrawRectangle(boxLeft, boxBottom, boxLeft + boxW, boxBottom + boxH)
    Else
        Set oBox = oPage.Drop(oBox, boxLeft + boxW / 2, boxBottom + boxH / 2)
        oBox.CellsU("Width").FormulaU  = CStr(boxW) & " in"
        oBox.CellsU("Height").FormulaU = CStr(boxH) & " in"
    End If

    oBox.CellsU("Rounding").FormulaU = BOX_ROUNDING

    ' Drop icon
    Dim iconMaster As Object
    Set iconMaster = GetCachedMaster(oStencil, modelStr, subType)

    Dim iconX As Double
    Dim iconY As Double
    iconX = boxLeft + boxW / 2
    iconY = centY + ICON_Y_OFFSET

    If Not iconMaster Is Nothing Then
        Dim oIcon As Object
        Set oIcon = oPage.Drop(iconMaster, iconX, iconY)

        ' Icon auto-resize feature
        If NetworkSettings.AutoResizeIcons Then
            Dim iconSz As Double
            iconSz = GetIconSizeForCount(qty)
            oIcon.CellsU("Width").FormulaU  = CStr(iconSz) & " in"
            oIcon.CellsU("Height").FormulaU = CStr(iconSz) & " in"
        End If

        RegisterShape oIcon, groupKey & "_icon", dictNameToShape
    End If

    ' Apply label
    ApplyDeviceLabel oPage, oBox, firstRec, qty, groupKey, dictNameToShape, pType

    RegisterShape oBox, groupKey, dictNameToShape

    PlaceDeviceGroup = xPos + boxW + BOX_LEFT_MARGIN
End Function

' ----------------------------------------------------------
' GetIconSizeForCount – returns icon size based on group count
' ----------------------------------------------------------
Private Function GetIconSizeForCount(ByVal deviceCount As Long) As Double
    Select Case deviceCount
        Case 1
            GetIconSizeForCount = NetworkSettings.IconSizeSingle
        Case 2 To 5
            GetIconSizeForCount = NetworkSettings.IconSizeSmall
        Case 6 To 10
            GetIconSizeForCount = NetworkSettings.IconSizeMedium
        Case Else
            GetIconSizeForCount = NetworkSettings.IconSizeLarge
    End Select
End Function

' ============================================================
'  CONNECTOR DRAWING
' ============================================================

Private Sub DrawUpstreamConnectors(ByVal oPage As Object, _
                                    ByVal dictNameToShape As Object, _
                                    ByVal dictIDtoUpstream As Object, _
                                    ByVal dictNameToID As Object, _
                                    ByVal pType As PageType)
    If dictNameToShape Is Nothing Then Exit Sub

    Dim dictLegend As Object
    Set dictLegend = CreateObject("Scripting.Dictionary")

    Dim shapeKey As Variant
    For Each shapeKey In dictNameToShape.Keys
        Dim keyStr As String
        keyStr = CStr(shapeKey)

        ' Skip icon shapes and non-group shapes
        If Right(keyStr, 5) = "_icon" Then GoTo NextShape2

        Dim oShape As Object
        Set oShape = dictNameToShape(keyStr)
        If oShape Is Nothing Then GoTo NextShape2

        ' Find upstream name from shape property
        Dim upName As String
        upName = ""
        On Error Resume Next
        upName = SafeCleanString(oShape.CellsU("Prop.UpstreamDevice").ResultStr(""))
        On Error GoTo 0
        If upName = "" Then GoTo NextShape2

        If Not dictNameToShape.Exists(upName) Then GoTo NextShape2

        Dim oUpstream As Object
        Set oUpstream = dictNameToShape(upName)
        If oUpstream Is Nothing Then GoTo NextShape2

        ' Only draw connectors for shapes on the same page
        If Not AreSamePage(oShape, oUpstream) Then GoTo NextShape2

        DrawSingleConnector oPage, oShape, oUpstream, keyStr, pType, dictLegend

NextShape2:
    Next shapeKey

    If pType = ptDataflow Then AddToLegend oPage, dictLegend
End Sub

Private Sub DrawSingleConnector(ByVal oPage As Object, _
                                  ByVal oFrom As Object, ByVal oTo As Object, _
                                  ByVal connKey As String, ByVal pType As PageType, _
                                  ByRef dictLegend As Object)
    If oFrom Is Nothing Or oTo Is Nothing Then Exit Sub
    On Error Resume Next

    Dim oConn As Object
    Set oConn = oPage.Drop(oPage.Application.ConnectorToolDataObject, 0, 0)

    If oConn Is Nothing Then Exit Sub

    GlueConnector oConn, oFrom, oTo

    ' CRITICAL: Use CStr() when assigning integer values to FormulaU
    oConn.CellsU("ShapeRouteStyle").FormulaU = CStr(visLORouteRightAngle)

    Dim mediaStr As String
    mediaStr = ""
    On Error Resume Next
    mediaStr = SafeCleanString(oFrom.CellsU("Prop.NetworkMedia").ResultStr(""))
    On Error GoTo 0

    Dim styleVal As Long
    styleVal = GetLineStyleForMedia(mediaStr)
    oConn.CellsU("LinePattern").FormulaU = CStr(styleVal)

    If NetworkSettings.ColorByMedia Then
        oConn.CellsU("LineColor").FormulaU = GetColorForMedia(mediaStr)
    End If

    If NetworkSettings.CurvedRouting Then
        oConn.CellsU("ObjType").FormulaU   = "2"
        oConn.CellsU("RouteStyle").FormulaU = PAGE3_ARROW_STYLE
    End If

    If NetworkSettings.ShowConnectorLabels Then
        ApplyConnectorLabel oConn, oFrom, mediaStr, pType
    End If

    On Error GoTo 0
End Sub

' ============================================================
'  DRAWING HELPERS
' ============================================================

Private Sub DrawSwimLaneBox(ByVal oPage As Object, ByVal levelKey As String, _
                              ByVal lB As Double, ByVal lT As Double, _
                              ByVal pType As PageType)
    On Error Resume Next

    Dim pageW As Double
    pageW = oPage.PageSheet.Cells("PageWidth").ResultIU

    ' Lane background
    Dim oLane As Object
    Set oLane = oPage.DrawRectangle(LEFT_MARGIN_RESERVE, lB, pageW - RIGHT_MARGIN_RESERVE, lT)
    oLane.CellsU("FillForegnd").FormulaU = "RGB(245,245,250)"
    oLane.CellsU("LineColor").FormulaU   = "RGB(180,180,190)"
    oLane.SendToBack

    ' Level label on left margin
    Dim oLabel As Object
    Set oLabel = oPage.DrawRectangle(0, lB, LEFT_MARGIN_RESERVE, lT)
    oLabel.Text = GetDetailedLevelLabel(levelKey)
    oLabel.CellsU("FillForegnd").FormulaU   = "RGB(30,60,120)"
    oLabel.CellsU("FontColor").FormulaU     = "RGB(255,255,255)"
    oLabel.CellsU("VerticalAlign").FormulaU = "1"
    oLabel.CellsU("HorzAlign").FormulaU     = CStr(visHorzAlign)
    On Error GoTo 0
End Sub

Private Sub ApplyDeviceLabel(ByVal oPage As Object, ByVal oBox As Object, _
                               ByVal rec As Object, ByVal qty As Long, _
                               ByVal groupKey As String, _
                               ByRef dictNameToShape As Object, _
                               ByVal pType As PageType)
    On Error Resume Next

    Dim labelText As String
    labelText = SafeString(rec("Identifier"))
    If NetworkSettings.ShowModel Then
        labelText = labelText & vbCrLf & SafeString(rec("Model"))
    End If
    If qty > 1 Then
        labelText = labelText & vbCrLf & "(" & qty & "x)"
    End If

    If NetworkSettings.AttachLabel Then
        oBox.Text = labelText
    Else
        ' Place a separate text shape near the icon
        Dim oText As Object
        Dim pinX As Double
        Dim pinY As Double
        pinX = oBox.CellsU("PinX").ResultIU
        pinY = oBox.CellsU("PinY").ResultIU - TEXT_Y_OFFSET

        Set oText = oPage.DrawRectangle(pinX - 0.5, pinY - 0.2, pinX + 0.5, pinY + 0.2)
        oText.Text = labelText
        oText.CellsU("FillPattern").FormulaU = "0"
        oText.CellsU("LinePattern").FormulaU = "0"
        RegisterShape oText, groupKey & "_label", dictNameToShape
    End If
    On Error GoTo 0
End Sub

Private Function PickBoxMaster(ByVal oPage As Object, ByVal oStencil As Object, _
                                 ByVal pType As PageType, ByVal rec As Object) As Object
    On Error Resume Next
    Dim masterName As String
    Select Case pType
        Case ptBoundary
            masterName = MASTER_SHAPE_LEVEL_2
        Case ptRiser
            masterName = MASTER_SHAPE_LEVEL_1
        Case Else
            masterName = MASTER_SHAPE_NAME
    End Select

    Dim oMaster As Object
    Set oMaster = GetCachedMaster(oStencil, masterName, "")
    Set PickBoxMaster = oMaster
    On Error GoTo 0
End Function

Private Sub RegisterShape(ByVal oShape As Object, ByVal key As String, _
                            ByRef dictNameToShape As Object)
    If oShape Is Nothing Then Exit Sub
    If dictNameToShape Is Nothing Then Exit Sub
    If key = "" Then Exit Sub

    On Error Resume Next
    If dictNameToShape.Exists(key) Then
        dictNameToShape(key) = oShape
    Else
        dictNameToShape.Add key, oShape
    End If
    On Error GoTo 0
End Sub

' ============================================================
'  CONNECTOR HELPERS
' ============================================================

Private Function GlueConnector(ByVal oConn As Object, _
                                 ByVal oFrom As Object, _
                                 ByVal oTo As Object) As Boolean
    On Error Resume Next
    GlueConnector = False
    If oConn Is Nothing Or oFrom Is Nothing Or oTo Is Nothing Then Exit Function
    oConn.CellsU("BeginX").GlueTo oFrom.CellsU("PinX")
    oConn.CellsU("EndX").GlueTo   oTo.CellsU("PinX")
    GlueConnector = True
    On Error GoTo 0
End Function

Private Function GetLineStyleForMedia(ByVal mediaStr As String) As Long
    Dim m As String
    m = UCase(Trim(mediaStr))
    Select Case True
        Case InStr(m, "ETHERNET") > 0
            GetLineStyleForMedia = LINESTYLE_ETHERNET
        Case InStr(m, "FIBER") > 0
            GetLineStyleForMedia = LINESTYLE_FIBER
        Case InStr(m, "WIRELESS") > 0 Or InStr(m, "WIFI") > 0
            GetLineStyleForMedia = LINESTYLE_WIRELESS
        Case InStr(m, "SERIAL") > 0
            GetLineStyleForMedia = LINESTYLE_SERIAL
        Case Else
            GetLineStyleForMedia = LINESTYLE_DEFAULT
    End Select
End Function

Private Function GetColorForMedia(ByVal mediaStr As String) As String
    Dim m As String
    m = UCase(Trim(mediaStr))
    Select Case True
        Case InStr(m, "ETHERNET") > 0
            GetColorForMedia = COLOR_ETHERNET
        Case InStr(m, "FIBER") > 0
            GetColorForMedia = COLOR_FIBER
        Case InStr(m, "WIRELESS") > 0 Or InStr(m, "WIFI") > 0
            GetColorForMedia = COLOR_WIRELESS
        Case InStr(m, "SERIAL") > 0
            GetColorForMedia = COLOR_SERIAL
        Case Else
            GetColorForMedia = COLOR_DEFAULT
    End Select
End Function

Private Sub ApplyConnectorLabel(ByVal oConn As Object, ByVal oFrom As Object, _
                                  ByVal mediaStr As String, ByVal pType As PageType)
    On Error Resume Next
    Dim proto As String
    proto = SafeCleanString(oFrom.CellsU("Prop.Protocol").ResultStr(""))

    Dim labelText As String
    If proto <> "" And mediaStr <> "" Then
        labelText = proto & " / " & mediaStr
    ElseIf proto <> "" Then
        labelText = proto
    Else
        labelText = mediaStr
    End If

    If labelText <> "" Then
        oConn.Text = labelText
        oConn.CellsU("TxtPinY").FormulaU = "0.5"
    End If
    On Error GoTo 0
End Sub

Private Sub AddToLegend(ByVal oPage As Object, ByVal dictLegend As Object)
    If Not NetworkSettings.ShowLegend Then Exit Sub
    If dictLegend Is Nothing Then Exit Sub
    ' Legend entries are collected in dictLegend and rendered by DrawLegendBox
End Sub

Private Function IsValidParentName(ByVal name As String) As Boolean
    IsValidParentName = (Trim(name) <> "" And name <> "N/A" And _
                         UCase(Trim(name)) <> "NONE")
End Function

Private Function IsInternalConnection(ByVal fromKey As String, _
                                       ByVal toKey As String) As Boolean
    ' Consider connection internal if both shapes share the same level/model prefix
    Dim fromParts() As String
    Dim toParts()   As String
    fromParts = Split(fromKey, "|")
    toParts   = Split(toKey, "|")

    If UBound(fromParts) >= 1 And UBound(toParts) >= 1 Then
        IsInternalConnection = (fromParts(0) & fromParts(1) = toParts(0) & toParts(1))
    Else
        IsInternalConnection = False
    End If
End Function

Private Function ResolveShape(ByVal dictNameToShape As Object, _
                                ByVal key As String) As Object
    Set ResolveShape = Nothing
    If dictNameToShape Is Nothing Then Exit Function
    If dictNameToShape.Exists(key) Then
        Set ResolveShape = dictNameToShape(key)
    End If
End Function

Private Function AreSamePage(ByVal oShapeA As Object, ByVal oShapeB As Object) As Boolean
    AreSamePage = False
    If oShapeA Is Nothing Or oShapeB Is Nothing Then Exit Function
    On Error Resume Next
    AreSamePage = (oShapeA.ContainingPage.ID = oShapeB.ContainingPage.ID)
    On Error GoTo 0
End Function

Private Function CreateExternalCloud(ByVal oPage As Object, ByVal oStencil As Object, _
                                      ByVal parentName As String, _
                                      ByVal nearX As Double, _
                                      ByVal nearY As Double) As Object
    On Error Resume Next
    Dim oCloud As Object
    Set oCloud = oPage.DrawRectangle(nearX - 0.75, nearY - 0.4, nearX + 0.75, nearY + 0.4)
    oCloud.Text = parentName
    oCloud.CellsU("FillForegnd").FormulaU = "RGB(230,230,255)"
    oCloud.CellsU("LineColor").FormulaU   = "RGB(120,120,200)"
    oCloud.CellsU("Rounding").FormulaU    = "0.25 in"
    Set CreateExternalCloud = oCloud
    On Error GoTo 0
End Function

' ============================================================
'  LEGEND BOX
' ============================================================

Private Sub DrawLegendBox(ByVal oPage As Object, ByVal dictNameToShape As Object)
    If Not NetworkSettings.ShowLegend Then Exit Sub
    On Error Resume Next

    Dim pageW As Double
    Dim pageH As Double
    pageW = oPage.PageSheet.Cells("PageWidth").ResultIU
    pageH = oPage.PageSheet.Cells("PageHeight").ResultIU

    Dim lgLeft   As Double
    Dim lgBottom As Double
    Dim lgWidth  As Double
    Dim lgHeight As Double
    lgWidth  = 1.75
    lgHeight = 1.4
    lgLeft   = pageW - lgWidth - 0.1
    lgBottom = 0.1

    Dim oLegend As Object
    Set oLegend = oPage.DrawRectangle(lgLeft, lgBottom, lgLeft + lgWidth, lgBottom + lgHeight)
    oLegend.Text = "Legend"
    oLegend.CellsU("FillForegnd").FormulaU = "RGB(250,250,250)"
    oLegend.CellsU("LineColor").FormulaU   = "RGB(100,100,100)"

    ' Add color key entries
    Dim entries(4) As String
    Dim colors(4)  As String
    entries(0) = "Ethernet"  : colors(0) = COLOR_ETHERNET
    entries(1) = "Fiber"     : colors(1) = COLOR_FIBER
    entries(2) = "Wireless"  : colors(2) = COLOR_WIRELESS
    entries(3) = "Serial"    : colors(3) = COLOR_SERIAL
    entries(4) = "Other"     : colors(4) = COLOR_DEFAULT

    Dim entryH As Double
    entryH = 0.22

    Dim ei As Long
    For ei = 0 To 4
        Dim eY As Double
        eY = lgBottom + lgHeight - 0.25 - ei * entryH
        Dim oEntry As Object
        Set oEntry = oPage.DrawRectangle(lgLeft + 0.05, eY - entryH / 2, _
                                          lgLeft + 0.25, eY + entryH / 2)
        oEntry.CellsU("FillForegnd").FormulaU = colors(ei)
        oEntry.CellsU("LinePattern").FormulaU = "0"

        Dim oTxt As Object
        Set oTxt = oPage.DrawRectangle(lgLeft + 0.3, eY - entryH / 2, _
                                        lgLeft + lgWidth - 0.05, eY + entryH / 2)
        oTxt.Text = entries(ei)
        oTxt.CellsU("FillPattern").FormulaU = "0"
        oTxt.CellsU("LinePattern").FormulaU = "0"
    Next ei

    On Error GoTo 0
End Sub

' ============================================================
'  CHAIN HEAD
' ============================================================

Private Function GetChainHeadID(ByVal deviceID As String, _
                                  ByVal dictIDtoUpstream As Object, _
                                  ByVal dictIDtoModel As Object) As String
    Dim current   As String
    Dim visited   As Object
    Set visited   = CreateObject("Scripting.Dictionary")

    current = deviceID
    Dim currentModel As String
    If dictIDtoModel.Exists(current) Then
        currentModel = CStr(dictIDtoModel(current))
    Else
        GetChainHeadID = deviceID
        Exit Function
    End If

    Do
        If visited.Exists(current) Then Exit Do  ' cycle guard
        visited.Add current, True

        If Not dictIDtoUpstream.Exists(current) Then Exit Do

        Dim upID As String
        upID = CStr(dictIDtoUpstream(current))
        If upID = "" Then Exit Do

        Dim upModel As String
        upModel = ""
        If dictIDtoModel.Exists(upID) Then upModel = CStr(dictIDtoModel(upID))

        If UCase(Trim(upModel)) <> UCase(Trim(currentModel)) Then Exit Do

        current = upID
    Loop

    GetChainHeadID = current
End Function

' ============================================================
'  MASTER CACHE
' ============================================================

Private Function GetCachedMaster(ByVal oStencil As Object, _
                                   ByVal modelStr As String, _
                                   ByVal subType As String) As Object
    Set GetCachedMaster = Nothing
    If oStencil Is Nothing Then Exit Function

    If dictMasterCache Is Nothing Then
        Set dictMasterCache = CreateObject("Scripting.Dictionary")
    End If

    Dim cacheKey As String
    cacheKey = UCase(Trim(modelStr)) & "|" & UCase(Trim(subType))

    If dictMasterCache.Exists(cacheKey) Then
        Set GetCachedMaster = dictMasterCache(cacheKey)
        Exit Function
    End If

    ' Try exact model name first
    Dim oMaster As Object
    On Error Resume Next
    Set oMaster = oStencil.Masters(modelStr)
    On Error GoTo 0

    ' Subtype fallback
    If oMaster Is Nothing And subType <> "" Then
        If dictSubTypeMap Is Nothing Then InitSubTypeMap
        If dictSubTypeMap.Exists(UCase(Trim(subType))) Then
            Dim altName As String
            altName = CStr(dictSubTypeMap(UCase(Trim(subType))))
            On Error Resume Next
            Set oMaster = oStencil.Masters(altName)
            On Error GoTo 0
        End If
    End If

    ' Cache result (even if Nothing)
    dictMasterCache.Add cacheKey, oMaster
    Set GetCachedMaster = oMaster
End Function

' ============================================================
'  DATA ACCESS & INITIALIZATION
' ============================================================

Private Sub InitSubTypeMap()
    Set dictSubTypeMap = CreateObject("Scripting.Dictionary")
    dictSubTypeMap.Add "CONTROLLER", "PLC"
    dictSubTypeMap.Add "PLC",        "PLC"
    dictSubTypeMap.Add "RTU",        "RTU"
    dictSubTypeMap.Add "HMI",        "HMI"
    dictSubTypeMap.Add "HISTORIAN",  "Server"
    dictSubTypeMap.Add "SERVER",     "Server"
    dictSubTypeMap.Add "WORKSTATION","Workstation"
    dictSubTypeMap.Add "SWITCH",     "Network Switch"
    dictSubTypeMap.Add "ROUTER",     "Router"
    dictSubTypeMap.Add "FIREWALL",   "Firewall"
    dictSubTypeMap.Add "SENSOR",     "Field Device"
    dictSubTypeMap.Add "ACTUATOR",   "Field Device"
    dictSubTypeMap.Add "VALVE",      "Field Device"
    dictSubTypeMap.Add "PUMP",       "Field Device"
    dictSubTypeMap.Add "DRIVE",      "VFD"
    dictSubTypeMap.Add "VFD",        "VFD"
End Sub

Private Sub CleanupCaches()
    Set dictMasterCache     = Nothing
    Set dictSubTypeMap      = Nothing
    Set dictShapeToPageName = Nothing
End Sub

Private Sub EnableBatchMode(ByVal enable As Boolean)
    On Error Resume Next
    If enable Then
        Application.ScreenUpdating     = False
        Application.EventsEnabled      = False
        Application.ShowChanges        = False
    Else
        Application.ScreenUpdating     = True
        Application.EventsEnabled      = True
        Application.ShowChanges        = True
    End If
    On Error GoTo 0
End Sub

Private Sub CleanupGeneratedPages(ByVal oDoc As Object)
    If oDoc Is Nothing Then Exit Sub
    On Error Resume Next

    Dim i As Long
    For i = oDoc.Pages.Count To 1 Step -1
        Dim oPage As Object
        Set oPage = oDoc.Pages(i)
        Dim pName As String
        pName = UCase(oPage.Name)
        If InStr(pName, "BOUNDARY") > 0 Or InStr(pName, "RISER") > 0 Or _
           InStr(pName, "DATAFLOW") > 0 Then
            oPage.Delete
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub UpdateProgress(ByVal msg As String)
    On Error Resume Next
    Application.StatusBar = msg
    On Error GoTo 0
End Sub

Private Function GetFirstRecordset(ByVal oDoc As Object) As Object
    Set GetFirstRecordset = Nothing
    If oDoc Is Nothing Then Exit Function
    On Error Resume Next
    If oDoc.DataRecordsets.Count > 0 Then
        Set GetFirstRecordset = oDoc.DataRecordsets(1)
    End If
    On Error GoTo 0
End Function

Private Function MapColumns(ByVal oRS As Object) As ColumnIndices
    Dim cols As ColumnIndices
    cols.Location   = -1
    cols.Model      = -1
    cols.Level      = -1
    cols.Upstream   = -1
    cols.Protocol   = -1
    cols.Media      = -1
    cols.SubType    = -1
    cols.Identifier = -1
    cols.Port       = -1

    If oRS Is Nothing Then
        MapColumns = cols
        Exit Function
    End If

    On Error Resume Next
    Dim i As Long
    Dim colNames As Variant
    colNames = oRS.DataColumns.GetNames

    If IsArray(colNames) Then
        For i = LBound(colNames) To UBound(colNames)
            Dim cn As String
            cn = CStr(colNames(i))
            If IsMatch(cn, COL_LOCATION)   Then cols.Location   = i
            If IsMatch(cn, COL_MODEL)      Then cols.Model      = i
            If IsMatch(cn, COL_LEVEL)      Then cols.Level      = i
            If IsMatch(cn, COL_UPSTREAM)   Then cols.Upstream   = i
            If IsMatch(cn, COL_PROTOCOL)   Then cols.Protocol   = i
            If IsMatch(cn, COL_MEDIA)      Then cols.Media      = i
            If IsMatch(cn, COL_SUBTYPE)    Then cols.SubType    = i
            If IsMatch(cn, COL_IDENTIFIER) Then cols.Identifier = i
            If IsMatch(cn, COL_PORT)       Then cols.Port       = i
        Next i
    End If
    On Error GoTo 0

    MapColumns = cols
End Function

Private Function OpenStencil() As Object
    Set OpenStencil = Nothing
    On Error Resume Next

    Dim stencilPath As String
    stencilPath = Environ("USERPROFILE") & "\" & NetworkSettings.StencilSubfolder & _
                  NetworkSettings.StencilFilename

    Dim oStencil As Object
    Set oStencil = Application.Documents.OpenEx(stencilPath, visOpenHidden)
    Set OpenStencil = oStencil
    On Error GoTo 0
End Function

Private Function GetSortedBuildingNames(ByVal oRS As Object, _
                                          ByRef cols As ColumnIndices) As Variant
    Dim dictBuildings As Object
    Set dictBuildings = CreateObject("Scripting.Dictionary")

    If oRS Is Nothing Then
        GetSortedBuildingNames = Array()
        Exit Function
    End If

    On Error Resume Next
    Dim rowIDs As Variant
    rowIDs = oRS.GetDataRowIDs("")

    If IsArray(rowIDs) Then
        Dim i As Long
        For i = LBound(rowIDs) To UBound(rowIDs)
            Dim rowData As Variant
            rowData = oRS.GetRowData(rowIDs(i))
            If IsArray(rowData) And cols.Location >= 0 Then
                Dim bldgOrig  As String
                Dim bldgClean As String
                bldgOrig  = Trim(CStr(rowData(cols.Location)))
                bldgClean = CleanString(bldgOrig)
                If bldgClean <> "" And Not dictBuildings.Exists(bldgClean) Then
                    dictBuildings.Add bldgClean, bldgOrig   ' key=clean, value=original
                End If
            End If
        Next i
    End If
    On Error GoTo 0

    Dim names() As String
    If dictBuildings.Count = 0 Then
        GetSortedBuildingNames = Array()
        Exit Function
    End If
    ReDim names(0 To dictBuildings.Count - 1)
    Dim k As Long
    Dim key As Variant
    k = 0
    For Each key In dictBuildings.Keys
        names(k) = CStr(dictBuildings(key))   ' return original display name
        k = k + 1
    Next key

    BubbleSort names
    GetSortedBuildingNames = names
End Function

Private Sub BuildChainLookups(ByVal oRS As Object, ByRef cols As ColumnIndices, _
                                ByVal buildingName As String, _
                                ByRef dictIDtoModel As Object, _
                                ByRef dictIDtoUpstream As Object, _
                                ByRef dictNameToID As Object)
    Set dictIDtoModel    = CreateObject("Scripting.Dictionary")
    Set dictIDtoUpstream = CreateObject("Scripting.Dictionary")
    Set dictNameToID     = CreateObject("Scripting.Dictionary")

    If oRS Is Nothing Then Exit Sub

    On Error Resume Next
    Dim rowIDs As Variant
    rowIDs = oRS.GetDataRowIDs("")

    If Not IsArray(rowIDs) Then Exit Sub

    Dim i As Long
    For i = LBound(rowIDs) To UBound(rowIDs)
        Dim rowData As Variant
        rowData = oRS.GetRowData(rowIDs(i))
        If Not IsArray(rowData) Then GoTo NextRow

        ' Filter by building
        Dim loc As String
        loc = CleanString(CStr(rowData(cols.Location)))
        If loc <> CleanString(buildingName) Then GoTo NextRow

        Dim devID As String
        devID = CStr(rowIDs(i))

        Dim modelVal As String
        modelVal = ""
        If cols.Model >= 0 Then modelVal = CleanString(CStr(rowData(cols.Model)))

        Dim upVal As String
        upVal = ""
        If cols.Upstream >= 0 Then upVal = CleanString(CStr(rowData(cols.Upstream)))

        Dim identVal As String
        identVal = ""
        If cols.Identifier >= 0 Then identVal = CleanString(CStr(rowData(cols.Identifier)))

        If Not dictIDtoModel.Exists(devID) Then dictIDtoModel.Add devID, modelVal
        If Not dictIDtoUpstream.Exists(devID) Then dictIDtoUpstream.Add devID, upVal
        If identVal <> "" And Not dictNameToID.Exists(identVal) Then
            dictNameToID.Add identVal, devID
        End If

NextRow:
    Next i
    On Error GoTo 0
End Sub

Private Function GroupDevicesByLevelModelChain(ByVal oRS As Object, _
                                                 ByRef cols As ColumnIndices, _
                                                 ByVal buildingName As String, _
                                                 ByVal dictIDtoModel As Object, _
                                                 ByVal dictIDtoUpstream As Object, _
                                                 ByVal dictNameToID As Object) As Object
    Dim dictGroups As Object
    Set dictGroups = CreateObject("Scripting.Dictionary")

    If oRS Is Nothing Then
        Set GroupDevicesByLevelModelChain = dictGroups
        Exit Function
    End If

    On Error Resume Next
    Dim rowIDs As Variant
    rowIDs = oRS.GetDataRowIDs("")

    If Not IsArray(rowIDs) Then
        Set GroupDevicesByLevelModelChain = dictGroups
        Exit Function
    End If

    Dim i As Long
    For i = LBound(rowIDs) To UBound(rowIDs)
        Dim rowData As Variant
        rowData = oRS.GetRowData(rowIDs(i))
        If Not IsArray(rowData) Then GoTo NextDevRow

        Dim loc As String
        loc = CleanString(CStr(rowData(cols.Location)))
        If loc <> CleanString(buildingName) Then GoTo NextDevRow

        Dim devID    As String
        Dim lvlVal   As String
        Dim modelVal As String
        Dim chainKey As String

        devID    = CStr(rowIDs(i))
        lvlVal   = ""
        modelVal = ""

        If cols.Level >= 0 Then lvlVal   = CleanString(CStr(rowData(cols.Level)))
        If cols.Model >= 0 Then modelVal = CleanString(CStr(rowData(cols.Model)))

        If NetworkSettings.UseTypicalGrouping Then
            chainKey = GetChainHeadID(devID, dictIDtoUpstream, dictIDtoModel)
        Else
            chainKey = devID
        End If

        Dim grpKey As String
        grpKey = lvlVal & "|" & modelVal & "|" & chainKey

        If Not dictGroups.Exists(grpKey) Then
            Dim newColl As Object
            Set newColl = New Collection
            dictGroups.Add grpKey, newColl
        End If

        ' Build a simple record object (use Dictionary as record)
        Dim recDict As Object
        Set recDict = CreateObject("Scripting.Dictionary")
        recDict.Add "ID",         devID
        recDict.Add "Location",   loc
        recDict.Add "Model",      modelVal
        recDict.Add "Level",      lvlVal
        If cols.Upstream >= 0 Then
            recDict.Add "Upstream",   CleanString(CStr(rowData(cols.Upstream)))
        Else
            recDict.Add "Upstream", ""
        End If
        If cols.Protocol >= 0 Then
            recDict.Add "Protocol",   CleanString(CStr(rowData(cols.Protocol)))
        Else
            recDict.Add "Protocol", ""
        End If
        If cols.Media >= 0 Then
            recDict.Add "Media",      CleanString(CStr(rowData(cols.Media)))
        Else
            recDict.Add "Media", ""
        End If
        If cols.SubType >= 0 Then
            recDict.Add "SubType",    CleanString(CStr(rowData(cols.SubType)))
        Else
            recDict.Add "SubType", ""
        End If
        If cols.Identifier >= 0 Then
            recDict.Add "Identifier", CleanString(CStr(rowData(cols.Identifier)))
        Else
            recDict.Add "Identifier", ""
        End If
        If cols.Port >= 0 Then
            recDict.Add "Port",       CleanString(CStr(rowData(cols.Port)))
        Else
            recDict.Add "Port", ""
        End If

        dictGroups(grpKey).Add recDict

NextDevRow:
    Next i
    On Error GoTo 0

    Set GroupDevicesByLevelModelChain = dictGroups
End Function

Private Sub ComputeLaneLayout(ByVal dictBuildingGroups As Object, _
                               ByRef dictLaneRows As Object, _
                               ByRef sortedLevelKeys As Variant, _
                               ByRef totalWeight As Double)
    Set dictLaneRows = CreateObject("Scripting.Dictionary")
    totalWeight = 0

    If dictBuildingGroups Is Nothing Then
        sortedLevelKeys = Array()
        Exit Sub
    End If

    ' Collect all unique level keys and their max per-row counts
    Dim dictLevels As Object
    Set dictLevels = CreateObject("Scripting.Dictionary")

    Dim grpKey As Variant
    For Each grpKey In dictBuildingGroups.Keys
        Dim parts() As String
        parts = Split(CStr(grpKey), "|")
        If UBound(parts) >= 0 Then
            Dim lvlKey As String
            lvlKey = parts(0)
            If Not dictLevels.Exists(lvlKey) Then dictLevels.Add lvlKey, 0
        End If
    Next grpKey

    ' For each level, count groups and compute row weight
    Dim lvlKey2 As Variant
    For Each lvlKey2 In dictLevels.Keys
        Dim lk As String
        lk = CStr(lvlKey2)

        Dim grpCount As Long
        grpCount = 0
        For Each grpKey In dictBuildingGroups.Keys
            If Left(CStr(grpKey), Len(lk) + 1) = lk & "|" Then
                grpCount = grpCount + 1
            End If
        Next grpKey

        Dim rowW As Double
        rowW = 1
        If grpCount > NetworkSettings.MaxPerRow Then
            rowW = Int((grpCount - 1) / NetworkSettings.MaxPerRow) + 1
        End If

        dictLaneRows(lk) = rowW
        totalWeight = totalWeight + rowW
    Next lvlKey2

    sortedLevelKeys = GetSortedLevelKeys(dictLevels)
End Sub

' ============================================================
'  UTILITY FUNCTIONS
' ============================================================

Private Function SafeString(ByVal v As Variant) As String
    On Error Resume Next
    If IsNull(v) Or IsEmpty(v) Then
        SafeString = ""
    Else
        SafeString = CStr(v)
    End If
    On Error GoTo 0
End Function

Private Function SafeCleanString(ByVal v As Variant) As String
    SafeCleanString = CleanString(SafeString(v))
End Function

Private Function CleanString(ByVal s As String) As String
    CleanString = UCase(Trim(Replace(s, " ", "")))
End Function

Private Function IsMatch(ByVal a As String, ByVal b As String) As Boolean
    IsMatch = (UCase(Trim(a)) = UCase(Trim(b)))
End Function

Private Function BuildPageName(ByVal pType As PageType, _
                                ByVal buildingName As String) As String
    Dim typeName As String
    Select Case pType
        Case ptBoundary : typeName = "Boundary"
        Case ptRiser    : typeName = "Riser"
        Case ptDataflow : typeName = "Dataflow"
        Case Else       : typeName = "Network"
    End Select
    BuildPageName = typeName & PAGE_NAME_SEPARATOR & buildingName
End Function

Private Function UniquePageName(ByVal oDoc As Object, ByVal baseName As String, _
                                  ByVal pageNum As Long, ByVal totalPages As Long) As String
    Dim name As String
    If totalPages > 1 Then
        name = baseName & " (" & pageNum & " of " & totalPages & ")"
    Else
        name = baseName
    End If

    ' Ensure uniqueness
    Dim suffix As Long
    suffix = 0
    Dim candidate As String
    candidate = name

    Dim found As Boolean
    found = True
    Do While found
        found = False
        On Error Resume Next
        Dim oTest As Object
        Set oTest = Nothing
        Set oTest = oDoc.Pages(candidate)
        If Not oTest Is Nothing Then
            found = True
            suffix = suffix + 1
            candidate = name & " (" & suffix & ")"
        End If
        On Error GoTo 0
    Loop

    UniquePageName = candidate
End Function

Private Sub UpdateTitleBlock(ByVal oPage As Object, ByVal oStencil As Object, _
                               ByVal buildingName As String, ByVal pType As PageType, _
                               ByVal pageNum As Long, ByVal totalPages As Long)
    If oPage Is Nothing Then Exit Sub
    On Error Resume Next

    Dim oTitle As Object
    Set oTitle = oPage.Shapes(MASTER_SHAPE_NAME)

    If oTitle Is Nothing Then
        ' Try to drop title block from stencil
        Dim oTitleMaster As Object
        Set oTitleMaster = GetCachedMaster(oStencil, MASTER_SHAPE_NAME, "")
        If Not oTitleMaster Is Nothing Then
            Dim pageW As Double
            Dim pageH As Double
            pageW = oPage.PageSheet.Cells("PageWidth").ResultIU
            pageH = oPage.PageSheet.Cells("PageHeight").ResultIU
            Set oTitle = oPage.Drop(oTitleMaster, pageW / 2, 0.5)
        End If
    End If

    If Not oTitle Is Nothing Then
        Dim typeName As String
        Select Case pType
            Case ptBoundary : typeName = "Network Boundary"
            Case ptRiser    : typeName = "Riser Diagram"
            Case ptDataflow : typeName = "Dataflow Diagram"
        End Select
        oTitle.CellsU("Prop.Title").FormulaU    = Chr(34) & typeName & Chr(34)
        oTitle.CellsU("Prop.Building").FormulaU = Chr(34) & buildingName & Chr(34)
        oTitle.CellsU("Prop.Date").FormulaU     = Chr(34) & Format(Date, "YYYY-MM-DD") & Chr(34)
        If totalPages > 1 Then
            oTitle.CellsU("Prop.Sheet").FormulaU = Chr(34) & pageNum & " of " & totalPages & Chr(34)
        End If
    End If
    On Error GoTo 0
End Sub

Private Function GetDetailedLevelLabel(ByVal levelKey As String) As String
    Dim lk As String
    lk = UCase(Trim(levelKey))

    Select Case lk
        Case "0", "LEVEL0", "LEVEL 0"
            GetDetailedLevelLabel = "Level 0 – Field Devices"
        Case "1", "LEVEL1", "LEVEL 1"
            GetDetailedLevelLabel = "Level 1 – Basic Control"
        Case "2", "LEVEL2", "LEVEL 2"
            GetDetailedLevelLabel = "Level 2 – Supervisory Control"
        Case "3", "LEVEL3", "LEVEL 3"
            GetDetailedLevelLabel = "Level 3 – Manufacturing Operations"
        Case "3.5", "LEVEL3.5", "LEVEL 3.5", "DMZ", "IDMZ"
            GetDetailedLevelLabel = "Level 3.5 – IDMZ"
        Case "4", "LEVEL4", "LEVEL 4"
            GetDetailedLevelLabel = "Level 4 – Business Planning"
        Case "5", "LEVEL5", "LEVEL 5"
            GetDetailedLevelLabel = "Level 5 – Enterprise"
        Case Else
            GetDetailedLevelLabel = levelKey
    End Select
End Function

Private Function GetSortedLevelKeys(ByVal dictLevels As Object) As Variant
    If dictLevels Is Nothing Then
        GetSortedLevelKeys = Array()
        Exit Function
    End If

    Dim keys() As String
    If dictLevels.Count = 0 Then
        GetSortedLevelKeys = Array()
        Exit Function
    End If
    ReDim keys(0 To dictLevels.Count - 1)
    Dim i As Long
    Dim k As Variant
    i = 0
    For Each k In dictLevels.Keys
        keys(i) = CStr(k)
        i = i + 1
    Next k

    ' Sort by Purdue level rank
    Dim n As Long
    n = UBound(keys) - LBound(keys) + 1
    Dim sorted As Boolean
    sorted = False
    Do While Not sorted
        sorted = True
        Dim j As Long
        For j = LBound(keys) To UBound(keys) - 1
            If GetLevelRank(keys(j)) > GetLevelRank(keys(j + 1)) Then
                Dim tmp As String
                tmp = keys(j)
                keys(j) = keys(j + 1)
                keys(j + 1) = tmp
                sorted = False
            End If
        Next j
    Loop

    GetSortedLevelKeys = keys
End Function

Private Function GetLevelRank(ByVal levelKey As String) As Double
    Dim lk As String
    lk = UCase(Trim(levelKey))

    Select Case lk
        Case "0", "LEVEL0", "LEVEL 0"           : GetLevelRank = 0
        Case "1", "LEVEL1", "LEVEL 1"           : GetLevelRank = 1
        Case "2", "LEVEL2", "LEVEL 2"           : GetLevelRank = 2
        Case "3", "LEVEL3", "LEVEL 3"           : GetLevelRank = 3
        Case "3.5", "LEVEL3.5", "LEVEL 3.5", "DMZ", "IDMZ" : GetLevelRank = 3.5
        Case "4", "LEVEL4", "LEVEL 4"           : GetLevelRank = 4
        Case "5", "LEVEL5", "LEVEL 5"           : GetLevelRank = 5
        Case Else
            ' Try to parse numeric portion
            Dim numVal As Double
            On Error Resume Next
            numVal = CDbl(Replace(Replace(lk, "LEVEL", ""), " ", ""))
            On Error GoTo 0
            GetLevelRank = numVal
    End Select
End Function

Private Function GetGravitySortedKeys(ByVal dictGroups As Object, _
                                        ByVal dictNameToShape As Object) As Variant
    If dictGroups Is Nothing Then
        GetGravitySortedKeys = Array()
        Exit Function
    End If

    If dictGroups.Count = 0 Then
        GetGravitySortedKeys = Array()
        Exit Function
    End If

    Dim keys() As String
    Dim xVals() As Double
    ReDim keys(0 To dictGroups.Count - 1)
    ReDim xVals(0 To dictGroups.Count - 1)

    Dim i As Long
    Dim k As Variant
    i = 0
    For Each k In dictGroups.Keys
        keys(i) = CStr(k)

        ' Try to find upstream shape X position
        Dim xPos As Double
        xPos = 99999

        Dim devList As Object
        Set devList = dictGroups(k)
        If Not devList Is Nothing Then
            If devList.Count > 0 Then
                On Error Resume Next
                Dim firstRec As Object
                Set firstRec = devList(1)
                Dim upName As String
                upName = SafeCleanString(firstRec("Upstream"))
                If upName <> "" And Not dictNameToShape Is Nothing Then
                    If dictNameToShape.Exists(upName) Then
                        Dim oUp As Object
                        Set oUp = dictNameToShape(upName)
                        If Not oUp Is Nothing Then
                            xPos = oUp.CellsU("PinX").ResultIU
                        End If
                    End If
                End If
                On Error GoTo 0
            End If
        End If
        xVals(i) = xPos
        i = i + 1
    Next k

    ' Bubble sort by xVals
    Dim sorted As Boolean
    sorted = False
    Do While Not sorted
        sorted = True
        Dim j As Long
        For j = LBound(keys) To UBound(keys) - 1
            If xVals(j) > xVals(j + 1) Then
                Dim tmpKey As String
                Dim tmpX   As Double
                tmpKey    = keys(j)
                tmpX      = xVals(j)
                keys(j)   = keys(j + 1)
                xVals(j)  = xVals(j + 1)
                keys(j + 1) = tmpKey
                xVals(j + 1) = tmpX
                sorted = False
            End If
        Next j
    Loop

    GetGravitySortedKeys = keys
End Function

Private Function GetCenteredXPosition(ByVal totalItems As Long, _
                                        ByVal itemWidth As Double, _
                                        ByVal pageWidth As Double) As Double
    Dim totalW As Double
    totalW = totalItems * itemWidth
    GetCenteredXPosition = (pageWidth - totalW) / 2
End Function

Private Sub BubbleSort(ByRef arr() As String)
    Dim i As Long
    Dim j As Long
    Dim tmp As String
    Dim n As Long
    n = UBound(arr) - LBound(arr) + 1

    For i = 0 To n - 2
        For j = LBound(arr) To LBound(arr) + n - 2 - i
            If arr(j) > arr(j + 1) Then
                tmp        = arr(j)
                arr(j)     = arr(j + 1)
                arr(j + 1) = tmp
            End If
        Next j
    Next i
End Sub

Private Function GetOrAddLayer(ByVal oPage As Object, ByVal layerName As String) As Object
    Set GetOrAddLayer = Nothing
    If oPage Is Nothing Then Exit Function

    On Error Resume Next
    Dim oLayer As Object
    Set oLayer = oPage.Layers(layerName)
    If oLayer Is Nothing Then
        Set oLayer = oPage.Layers.Add(layerName)
    End If
    Set GetOrAddLayer = oLayer
    On Error GoTo 0
End Function

Private Sub LogStep(ByVal context As String, ByVal msg As String)
    Debug.Print "[" & Format(Now(), "HH:MM:SS") & "] " & context & ": " & msg
End Sub
