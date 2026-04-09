# Better Export

Better Export is a Fusion add-in for people who are tired of setting the same export options over and over again.

It pulls STL, OBJ, 3MF, and F3D export into one place, remembers the settings you care about, and can optionally sort the exported files into tidy project folders afterward.

## Why I made it

Fusion's built-in export tools work, but if you export the same kinds of files repeatedly, the workflow gets old fast:

- re-select the same format options every time
- re-enter the same folders
- export multiple formats one by one
- clean up the resulting files manually afterward

This add-in is meant to make that process feel much less repetitive.

## What it can do

- Export `STL`, `OBJ`, `3MF`, and `F3D` from one command instead of bouncing between separate Fusion export flows
- Export multiple output types in a single pass when you need more than one format from the same design
- Remember the settings and folders you actually use, so repeat exports take much less setup
- Let you choose between a simple direct-export workflow and an automatic sort-and-organize workflow
- Keep per-project behavior where it matters, like normal export folders and whether a given file should auto-sort
- Choose whether to export the full design, only currently visible bodies, or a specific selection
- Support either one shared settings block or separate settings per format, depending on how much control you want
- Fit into Fusion more naturally with a toolbar button, Browser context-menu entry, and visible export progress while the batch runs

## Export sorting

If you turn on `Sort Automatically After Export`, Better Export will:

1. Export files into a temporary staging folder
2. Process the exported files immediately after the batch finishes
3. Move the cleaned results into your chosen output folder

The sorter currently handles:

- `.stl`
- `.3mf`
- `.obj`
- `.mtl`
- `.f3d`

The sorting pass can:

- dedupe mesh exports by version
- strip version markers and spaces from final filenames
- keep OBJ and MTL files linked correctly after renaming
- archive original F3D files
- create one cleaned top-level F3D copy for the highest-version file
- organize outputs into per-project folders and type-specific subfolders

There is also an `Allow Overwrite` option for cases where you want sorted outputs to replace existing files.

One nice side effect of auto-sort is that it plays well with the `Reload from disk` feature in many 3D printer slicers. Because the add-in keeps writing cleaned files back to stable project folders, you can often refresh an existing model in the slicer in place instead of dragging a new copy in every time you export an updated version from Fusion.

## Example sorted output

If `Sorted Projects Folder` is set to `~/Documents/Fusion Exports`, the add-in might create a structure like this after an auto-sorted export:

```text
~/Documents/Fusion Exports/
├── DeskCableClip/
│   ├── DeskCableClip.f3d
│   ├── 3MFs/
│   │   └── DeskCableClip_MainBody.3mf
│   ├── F3Ds/
│   │   ├── DeskCableClip v4.f3d
│   │   └── DeskCableClip v5.f3d
│   ├── OBJs/
│   │   ├── DeskCableClip_MainBody.mtl
│   │   └── DeskCableClip_MainBody.obj
│   └── STLs/
│       └── DeskCableClip_MainBody.stl
└── LampArmBracket/
    ├── LampArmBracket.f3d
    ├── F3Ds/
    │   ├── LampArmBracket v2.f3d
    │   └── LampArmBracket v3.f3d
    └── STLs/
        └── LampArmBracket_45deg.stl
```

A few things to notice:

- project folders are created from the cleaned filename prefix before the first underscore
- mesh files are sorted into `STLs`, `3MFs`, or `OBJs`
- `.mtl` files are kept beside their matching `.obj`
- original `.f3d` files are archived in `F3Ds`
- the highest-version `.f3d` also gets one cleaned top-level copy in the project folder

## Install

The easiest way to install it is:

1. Download the latest zip from the GitHub releases page.
2. Unzip it.
3. In Fusion, open `Utilities > Scripts and Add-Ins`.
4. Go to the `Add-Ins` tab.
5. Click the green `+` button.
6. Select the unzipped `BetterExport` folder.
7. Run the add-in.

The release zip also includes a `HOW TO INSTALL.txt` file with the default add-in locations for macOS and Windows if you prefer to install it by copying the folder into Fusion's Add-Ins directory yourself.

## Where it shows up in Fusion

When Fusion exposes the `Utilities` tab in the current workspace, the add-in adds a `Better Export` button there.

It also adds `Better Export` to the Browser right-click menu for exportable items like:

- components
- occurrences
- bodies

If nothing is selected, the add-in exports the active root component.

## A few notes

- Settings are stored locally in `BetterExport/settings.json`.
- Most preferences save as soon as you change them, even if you close the dialog without exporting.
- `File Name` is intentionally not persisted. It refreshes from the active Fusion document each time the dialog opens.
- The `Target` control lets you export the full design, only currently visible bodies, or a specific selection. `Export Full Design` temporarily exports from the root component and then restores the previous view state afterward.
- The add-in can let you know when a newer release is available, and you can also check manually from inside Fusion whenever you want.
- Fusion's API support varies a bit by version, so some options are shown only when your installed Fusion build exposes them.

## Project files

- `BetterExport/BetterExport.py` contains the add-in itself
- `BetterExport/export_sorter.py` contains the post-export sorting logic
- `BetterExport/BetterExport.manifest` is the Fusion add-in manifest
- `BetterExport/HOW TO INSTALL.txt` is the packaged quick-install guide

## Status

This project is working and installable today, but I still think of it as something that can keep getting nicer over time. If you spot missing Fusion options, rough edges in the UI, or export cases that should be handled better, those are all fair game for future updates.
