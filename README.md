# Better Export

**Better Export** is an Autodesk Fusion add-in that makes repeated exports faster, cleaner, and more consistent.

It brings common mesh and CAD export formats into one place, remembers the settings you actually use, and can optionally sort exported files into organized project folders automatically. It is designed for people who export often and want a smoother workflow than Fusion’s default export flow provides.

<img width="534" height="1169" alt="image" src="https://github.com/user-attachments/assets/8185c05e-436f-4098-b679-dc705aac4316" />

## Features

Better Export can:

- Export `STL`, `OBJ`, `3MF`, `F3D`, `IGES`, `SAT`, `SMT`, `STEP`, and `USDZ`
- Export multiple formats in a single pass
- Remember commonly used settings and folders
- Support either direct export or automatic post-export sorting
- Keep project-specific behavior where it matters, including export folders and auto-sort preferences
- Export either currently visible bodies or a specific selection
- Use either shared settings across formats or separate settings per format
- Send a single mesh export directly to Fusion’s print utility workflow
- Open the export destination automatically after a successful export
- Integrate naturally into Fusion with a toolbar button, Browser context-menu entry, and visible batch progress

## Why use it

Fusion’s built-in export tools work, but repeated exports can get repetitive quickly. Common tasks often involve:

- choosing the same output formats again and again
- re-entering the same folders
- exporting one format at a time
- manually cleaning up the resulting files afterward

Better Export is built to reduce that friction and make repeated export workflows feel faster and more predictable.

## Export sorting

If you choose **Sort Into Project Folders**, Better Export will:

1. Export files into a temporary staging folder
2. Process them immediately after the batch finishes
3. Move the cleaned results into your chosen output folder

The sorter currently supports:

- `.stl`
- `.3mf`
- `.obj`
- `.mtl`
- `.f3d`
- `.iges`
- `.sat`
- `.smt`
- `.step`
- `.usdz`

During sorting, Better Export can:

- deduplicate mesh exports by version
- remove version markers and spaces from final filenames
- keep OBJ and MTL files linked correctly after renaming
- archive original F3D files
- create one cleaned top-level F3D copy for the highest-version file
- organize exports into project folders with type-specific subfolders

There is also a **Replace Existing Sorted Files** option for workflows where sorted outputs should overwrite older sorted files. If you prefer to keep Fusion’s version markers in the filenames, you can disable **Strip Version Numbers From Sorted Filenames**. Better Export will still use the cleaned base name for the project folder while allowing the sorted file itself to keep its versioned name.

A useful side effect of auto-sort is that it works well with **Reload from disk** workflows in many slicers. Because Better Export writes cleaned files back to stable project folders, you can often refresh an existing model in place instead of re-importing it every time.

## Example sorted output

If **Sorted Projects Folder** is set to `~/Documents/Fusion Exports`, Better Export might create a structure like this:

```text
~/Documents/Fusion Exports/
├── DeskCableClip/
│   ├── DeskCableClip.f3d
│   ├── 3MF/
│   │   └── DeskCableClip_MainBody.3mf
│   ├── F3D/
│   │   ├── DeskCableClip v4.f3d
│   │   └── DeskCableClip v5.f3d
│   ├── OBJ/
│   │   ├── DeskCableClip_MainBody.mtl
│   │   └── DeskCableClip_MainBody.obj
│   └── STL/
│       └── DeskCableClip_MainBody.stl
└── LampArmBracket/
    ├── LampArmBracket.f3d
    ├── F3D/
    │   ├── LampArmBracket v2.f3d
    │   └── LampArmBracket v3.f3d
    └── STL/
        └── LampArmBracket_45deg.stl
```
In this structure:

- project folders are created from the cleaned filename prefix before the first underscore
- files are organized into type-specific folders such as `STL`, `3MF`, `OBJ`, `STEP`, and `USDZ`
- `.mtl` files stay beside their matching `.obj`
- original `.f3d` files are archived in `F3D`
- the highest-version `.f3d` also gets one cleaned top-level copy in the project folder

## Installation

### Option 1: Install from inside Fusion

1. Download the latest release ZIP from the GitHub Releases page
2. Unzip it
3. In Fusion, open **Utilities > Scripts and Add-Ins**
4. Open the **Add-Ins** tab
5. Click the green `+` button
6. Select the unzipped `BetterExport` folder
7. Run the add-in

### Option 2: Copy into Fusion’s Add-Ins folder

**macOS default Add-Ins folder**  
`~/Library/Application Support/Autodesk/Autodesk Fusion 360/API/AddIns/`

**Windows default Add-Ins folder**  
`%AppData%\Autodesk\Autodesk Fusion 360\API\AddIns\`

To install by copy:

1. Download and unzip the release package
2. Copy the `BetterExport` folder into your platform’s Add-Ins folder
3. Open Autodesk Fusion
4. Go to **Utilities > Scripts and Add-Ins**
5. Open the **Add-Ins** tab
6. Find **Better Export** and run it

## Where it appears in Fusion

When Fusion exposes the **Utilities** tab in the current workspace, Better Export adds a **Better Export** button there.

It also adds **Better Export** to the Browser right-click menu for exportable items such as:

- components
- occurrences
- bodies

## Notes

- Settings are stored locally in `BetterExport/settings.json`
- Most preferences save as soon as they are changed, even if the dialog is closed without exporting
- `File Name` is intentionally not persisted and refreshes from the active Fusion document each time the dialog opens
- The `Target` control lets you export only currently visible bodies or a specific selection
- Some options appear only when the installed Fusion version exposes the required API support

## Update checks

Better Export can let you know when a newer release is available, and it can also stage an update from inside Fusion when a new release is ready.

## Project structure

- `BetterExport/BetterExport.py` — main add-in logic
- `BetterExport/export_sorter.py` — post-export sorting logic
- `BetterExport/BetterExport.manifest` — Fusion add-in manifest
- `BetterExport/HOW TO INSTALL.txt` — packaged quick-install guide

## Status

Better Export is ready to use today and is intended to be practical for real day-to-day export workflows. Future updates may expand format coverage, improve UI details, and support additional Fusion export behaviors as Autodesk’s API allows. This project is already useful now, with room to keep improving over time.
