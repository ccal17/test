# Facility Network Generator

VBA modules for Visio-based facility network diagram generation.

## Modules

- **FacilityNetworkGenerator.bas** - Main generator module (v4.2)
- **NetworkSettings.bas** - Centralized configuration module (v4.2)

## Download

Download each file individually from GitHub:

1. Open the file in GitHub (e.g. click **FacilityNetworkGenerator.bas** in the file list above)
2. Click the **Download raw file** button (⬇ icon, top-right of the file view)  
   — *or* click **Raw**, then **File → Save As** in your browser
3. Repeat for **NetworkSettings.bas**

**Direct raw links** (right-click → Save link as…):

- [`FacilityNetworkGenerator.bas`](../../raw/HEAD/FacilityNetworkGenerator.bas)
- [`NetworkSettings.bas`](../../raw/HEAD/NetworkSettings.bas)

## Usage

1. Open your Visio document with data recordset
2. Import both `.bas` modules into the VBA editor (Alt+F11)
3. Edit settings in `NetworkSettings` module
4. Run `LaunchFacilityNetworkGenerator`