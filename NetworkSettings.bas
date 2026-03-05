Attribute VB_Name = "NetworkSettings"
Option Explicit

' ============================================================
' NetworkSettings.bas  v4.2
' Centralized configuration for FacilityNetworkGenerator
' ============================================================

' ----------------------------------------------------------
' Default / compile-time constants
' ----------------------------------------------------------
Private Const CFG_GENERATE_BOUNDARY     As Boolean = True
Private Const CFG_GENERATE_RISER        As Boolean = True
Private Const CFG_GENERATE_DATAFLOW     As Boolean = True
Private Const CFG_ATTACH_LABEL          As Boolean = True
Private Const CFG_SHOW_MODEL            As Boolean = False
Private Const CFG_USE_TYPICAL_GROUPING  As Boolean = True
Private Const CFG_AUTO_TYPICAL_OVERFLOW As Boolean = True
Private Const CFG_SHOW_CONNECTOR_LABELS As Boolean = True
Private Const CFG_COLOR_BY_MEDIA        As Boolean = True
Private Const CFG_SHOW_LEGEND           As Boolean = True
Private Const CFG_CURVED_ROUTING        As Boolean = True
Private Const CFG_OVERFLOW_STRATEGY     As Long    = 1
Private Const CFG_SHOW_CONT_REFS        As Boolean = True
Private Const CFG_MAX_PER_ROW           As Long    = 5
Private Const CFG_AUTO_RESIZE_ICONS     As Boolean = True
Private Const CFG_ICON_SIZE_SINGLE      As Double  = 1.25
Private Const CFG_ICON_SIZE_SMALL       As Double  = 1.0
Private Const CFG_ICON_SIZE_MEDIUM      As Double  = 0.85
Private Const CFG_ICON_SIZE_LARGE       As Double  = 0.7
Private Const CFG_STENCIL_FILENAME      As String  = "Master.vssx"
Private Const CFG_STENCIL_SUBFOLDER     As String  = "Documents\My Shapes\"
Private Const CFG_CLEANUP_OLD_PAGES     As Boolean = True

' ----------------------------------------------------------
' OverflowMode enum
' ----------------------------------------------------------
Public Enum OverflowMode
    omContinuation = 1
    omAutoExpand   = 2
    omCompress     = 3
End Enum

' ----------------------------------------------------------
' Internal state variables
' ----------------------------------------------------------
Private m_Initialized           As Boolean
Private m_GenerateBoundary      As Boolean
Private m_GenerateRiser         As Boolean
Private m_GenerateDataflow      As Boolean
Private m_AttachLabel           As Boolean
Private m_ShowModel             As Boolean
Private m_UseTypicalGrouping    As Boolean
Private m_AutoTypicalOverflow   As Boolean
Private m_ShowConnectorLabels   As Boolean
Private m_ColorByMedia          As Boolean
Private m_ShowLegend            As Boolean
Private m_CurvedRouting         As Boolean
Private m_OverflowStrategy      As OverflowMode
Private m_ShowContRefs          As Boolean
Private m_MaxPerRow             As Long
Private m_AutoResizeIcons       As Boolean
Private m_IconSizeSingle        As Double
Private m_IconSizeSmall         As Double
Private m_IconSizeMedium        As Double
Private m_IconSizeLarge         As Double
Private m_StencilFilename       As String
Private m_StencilSubfolder      As String
Private m_CleanupOldPages       As Boolean

' ----------------------------------------------------------
' EnsureInitialized
' ----------------------------------------------------------
Private Sub EnsureInitialized()
    If Not m_Initialized Then ResetDefaults
End Sub

' ----------------------------------------------------------
' ResetDefaults  – restore all settings to compile-time values
' ----------------------------------------------------------
Public Sub ResetDefaults()
    m_GenerateBoundary    = CFG_GENERATE_BOUNDARY
    m_GenerateRiser       = CFG_GENERATE_RISER
    m_GenerateDataflow    = CFG_GENERATE_DATAFLOW
    m_AttachLabel         = CFG_ATTACH_LABEL
    m_ShowModel           = CFG_SHOW_MODEL
    m_UseTypicalGrouping  = CFG_USE_TYPICAL_GROUPING
    m_AutoTypicalOverflow = CFG_AUTO_TYPICAL_OVERFLOW
    m_ShowConnectorLabels = CFG_SHOW_CONNECTOR_LABELS
    m_ColorByMedia        = CFG_COLOR_BY_MEDIA
    m_ShowLegend          = CFG_SHOW_LEGEND
    m_CurvedRouting       = CFG_CURVED_ROUTING
    m_OverflowStrategy    = CFG_OVERFLOW_STRATEGY
    m_ShowContRefs        = CFG_SHOW_CONT_REFS
    m_MaxPerRow           = CFG_MAX_PER_ROW
    m_AutoResizeIcons     = CFG_AUTO_RESIZE_ICONS
    m_IconSizeSingle      = CFG_ICON_SIZE_SINGLE
    m_IconSizeSmall       = CFG_ICON_SIZE_SMALL
    m_IconSizeMedium      = CFG_ICON_SIZE_MEDIUM
    m_IconSizeLarge       = CFG_ICON_SIZE_LARGE
    m_StencilFilename     = CFG_STENCIL_FILENAME
    m_StencilSubfolder    = CFG_STENCIL_SUBFOLDER
    m_CleanupOldPages     = CFG_CLEANUP_OLD_PAGES
    m_Initialized         = True
End Sub

' ----------------------------------------------------------
' Property Get / Let pairs
' ----------------------------------------------------------
Public Property Get GenerateBoundary() As Boolean
    EnsureInitialized
    GenerateBoundary = m_GenerateBoundary
End Property
Public Property Let GenerateBoundary(ByVal v As Boolean)
    EnsureInitialized
    m_GenerateBoundary = v
End Property

Public Property Get GenerateRiser() As Boolean
    EnsureInitialized
    GenerateRiser = m_GenerateRiser
End Property
Public Property Let GenerateRiser(ByVal v As Boolean)
    EnsureInitialized
    m_GenerateRiser = v
End Property

Public Property Get GenerateDataflow() As Boolean
    EnsureInitialized
    GenerateDataflow = m_GenerateDataflow
End Property
Public Property Let GenerateDataflow(ByVal v As Boolean)
    EnsureInitialized
    m_GenerateDataflow = v
End Property

Public Property Get AttachLabel() As Boolean
    EnsureInitialized
    AttachLabel = m_AttachLabel
End Property
Public Property Let AttachLabel(ByVal v As Boolean)
    EnsureInitialized
    m_AttachLabel = v
End Property

Public Property Get ShowModel() As Boolean
    EnsureInitialized
    ShowModel = m_ShowModel
End Property
Public Property Let ShowModel(ByVal v As Boolean)
    EnsureInitialized
    m_ShowModel = v
End Property

Public Property Get UseTypicalGrouping() As Boolean
    EnsureInitialized
    UseTypicalGrouping = m_UseTypicalGrouping
End Property
Public Property Let UseTypicalGrouping(ByVal v As Boolean)
    EnsureInitialized
    m_UseTypicalGrouping = v
End Property

Public Property Get AutoTypicalOverflow() As Boolean
    EnsureInitialized
    AutoTypicalOverflow = m_AutoTypicalOverflow
End Property
Public Property Let AutoTypicalOverflow(ByVal v As Boolean)
    EnsureInitialized
    m_AutoTypicalOverflow = v
End Property

Public Property Get ShowConnectorLabels() As Boolean
    EnsureInitialized
    ShowConnectorLabels = m_ShowConnectorLabels
End Property
Public Property Let ShowConnectorLabels(ByVal v As Boolean)
    EnsureInitialized
    m_ShowConnectorLabels = v
End Property

Public Property Get ColorByMedia() As Boolean
    EnsureInitialized
    ColorByMedia = m_ColorByMedia
End Property
Public Property Let ColorByMedia(ByVal v As Boolean)
    EnsureInitialized
    m_ColorByMedia = v
End Property

Public Property Get ShowLegend() As Boolean
    EnsureInitialized
    ShowLegend = m_ShowLegend
End Property
Public Property Let ShowLegend(ByVal v As Boolean)
    EnsureInitialized
    m_ShowLegend = v
End Property

Public Property Get CurvedRouting() As Boolean
    EnsureInitialized
    CurvedRouting = m_CurvedRouting
End Property
Public Property Let CurvedRouting(ByVal v As Boolean)
    EnsureInitialized
    m_CurvedRouting = v
End Property

Public Property Get OverflowStrategy() As OverflowMode
    EnsureInitialized
    OverflowStrategy = m_OverflowStrategy
End Property
Public Property Let OverflowStrategy(ByVal v As OverflowMode)
    EnsureInitialized
    m_OverflowStrategy = v
End Property

Public Property Get ShowContRefs() As Boolean
    EnsureInitialized
    ShowContRefs = m_ShowContRefs
End Property
Public Property Let ShowContRefs(ByVal v As Boolean)
    EnsureInitialized
    m_ShowContRefs = v
End Property

Public Property Get MaxPerRow() As Long
    EnsureInitialized
    MaxPerRow = m_MaxPerRow
End Property
Public Property Let MaxPerRow(ByVal v As Long)
    EnsureInitialized
    m_MaxPerRow = v
End Property

Public Property Get AutoResizeIcons() As Boolean
    EnsureInitialized
    AutoResizeIcons = m_AutoResizeIcons
End Property
Public Property Let AutoResizeIcons(ByVal v As Boolean)
    EnsureInitialized
    m_AutoResizeIcons = v
End Property

Public Property Get IconSizeSingle() As Double
    EnsureInitialized
    IconSizeSingle = m_IconSizeSingle
End Property
Public Property Let IconSizeSingle(ByVal v As Double)
    EnsureInitialized
    m_IconSizeSingle = v
End Property

Public Property Get IconSizeSmall() As Double
    EnsureInitialized
    IconSizeSmall = m_IconSizeSmall
End Property
Public Property Let IconSizeSmall(ByVal v As Double)
    EnsureInitialized
    m_IconSizeSmall = v
End Property

Public Property Get IconSizeMedium() As Double
    EnsureInitialized
    IconSizeMedium = m_IconSizeMedium
End Property
Public Property Let IconSizeMedium(ByVal v As Double)
    EnsureInitialized
    m_IconSizeMedium = v
End Property

Public Property Get IconSizeLarge() As Double
    EnsureInitialized
    IconSizeLarge = m_IconSizeLarge
End Property
Public Property Let IconSizeLarge(ByVal v As Double)
    EnsureInitialized
    m_IconSizeLarge = v
End Property

Public Property Get StencilFilename() As String
    EnsureInitialized
    StencilFilename = m_StencilFilename
End Property
Public Property Let StencilFilename(ByVal v As String)
    EnsureInitialized
    m_StencilFilename = v
End Property

Public Property Get StencilSubfolder() As String
    EnsureInitialized
    StencilSubfolder = m_StencilSubfolder
End Property
Public Property Let StencilSubfolder(ByVal v As String)
    EnsureInitialized
    m_StencilSubfolder = v
End Property

Public Property Get CleanupOldPages() As Boolean
    EnsureInitialized
    CleanupOldPages = m_CleanupOldPages
End Property
Public Property Let CleanupOldPages(ByVal v As Boolean)
    EnsureInitialized
    m_CleanupOldPages = v
End Property

' ----------------------------------------------------------
' ShouldGeneratePage – convenience helper
' ----------------------------------------------------------
Public Function ShouldGeneratePage(ByVal pageTypeStr As String) As Boolean
    EnsureInitialized
    Dim t As String
    t = UCase(Trim(pageTypeStr))
    Select Case t
        Case "BOUNDARY"
            ShouldGeneratePage = m_GenerateBoundary
        Case "RISER"
            ShouldGeneratePage = m_GenerateRiser
        Case "DATAFLOW"
            ShouldGeneratePage = m_GenerateDataflow
        Case Else
            ShouldGeneratePage = True
    End Select
End Function

' ----------------------------------------------------------
' LogActiveSettings – diagnostics dump
' ----------------------------------------------------------
Public Sub LogActiveSettings()
    EnsureInitialized
    Debug.Print "=================================================="
    Debug.Print " NetworkSettings v4.2  –  Active Configuration"
    Debug.Print "  " & Now()
    Debug.Print "=================================================="
    Debug.Print "  Generate Boundary      : " & m_GenerateBoundary
    Debug.Print "  Generate Riser         : " & m_GenerateRiser
    Debug.Print "  Generate Dataflow      : " & m_GenerateDataflow
    Debug.Print "  Attach Label           : " & m_AttachLabel
    Debug.Print "  Show Model             : " & m_ShowModel
    Debug.Print "  Use Typical Grouping   : " & m_UseTypicalGrouping
    Debug.Print "  Auto-Typical Overflow  : " & m_AutoTypicalOverflow
    Debug.Print "  Show Connector Labels  : " & m_ShowConnectorLabels
    Debug.Print "  Color By Media         : " & m_ColorByMedia
    Debug.Print "  Show Legend            : " & m_ShowLegend
    Debug.Print "  Curved Routing         : " & m_CurvedRouting
    Debug.Print "  Overflow Strategy      : " & m_OverflowStrategy & _
                "  (1=Continuation 2=AutoExpand 3=Compress)"
    Debug.Print "  Show Cont Refs         : " & m_ShowContRefs
    Debug.Print "  Max Per Row            : " & m_MaxPerRow
    Debug.Print "  Icon Auto-Resize       : " & m_AutoResizeIcons
    Debug.Print "    Icon Size Single     : " & m_IconSizeSingle & " in"
    Debug.Print "    Icon Size Small      : " & m_IconSizeSmall & " in"
    Debug.Print "    Icon Size Medium     : " & m_IconSizeMedium & " in"
    Debug.Print "    Icon Size Large      : " & m_IconSizeLarge & " in"
    Debug.Print "  Stencil Filename       : " & m_StencilFilename
    Debug.Print "  Stencil Subfolder      : " & m_StencilSubfolder
    Debug.Print "  Cleanup Old Pages      : " & m_CleanupOldPages
    Debug.Print "=================================================="
End Sub
