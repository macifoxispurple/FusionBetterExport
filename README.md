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

- Export `STL`, `OBJ`, `3MF`, and `F3D` from one command
- Export multiple file types in a single run
- Remember your export preferences between Fusion launches
- Use either one shared settings block or separate settings per format
- Default the `File Name` field to the active Fusion document name each time the dialog opens
- Remember the normal export folder and auto-sort preference per Fusion project, while keeping one global sorted-output folder for all automatically sorted exports
- Add a toolbar button in Fusion's `Utilities` area when available
- Add a right-click Browser context menu entry for exportable selections
- Show export progress while the batch is running
- Optionally auto-sort exported files into project folders after export
- Hide mesh-only controls in `F3D` per-format settings, so that group only shows options Fusion actually supports

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

On a successful sorted export, the temporary staging folder is removed automatically. If something fails, the temp folder is kept so you can inspect what was exported.

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
- Fusion's API support varies a bit by version, so some options are shown only when your installed Fusion build exposes them.

## Project files

- `BetterExport/BetterExport.py` contains the add-in itself
- `BetterExport/export_sorter.py` contains the post-export sorting logic
- `BetterExport/BetterExport.manifest` is the Fusion add-in manifest
- `BetterExport/HOW TO INSTALL.txt` is the packaged quick-install guide

## Status

This project is working and installable today, but I still think of it as something that can keep getting nicer over time. If you spot missing Fusion options, rough edges in the UI, or export cases that should be handled better, those are all fair game for future updates.
