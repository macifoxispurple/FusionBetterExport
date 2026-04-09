# Better Export

`Better Export` is an Autodesk Fusion add-in that gives STL, OBJ, 3MF, and F3D export a persistent settings dialog. It can also export into a staging folder and automatically sort the resulting files into project folders. The last-used export options are stored in `settings.json` beside the add-in and are restored the next time the command runs, even after Fusion is restarted.

## What it does

- Exposes STL, OBJ, 3MF, and F3D export in one command, including multi-select export in a single run.
- Supports either one shared export settings block or per-format settings blocks, depending on the selected `Settings Scope`.
- Can export into an automatic temporary staging folder and then sort results into project folders under a separate output root.
- Remembers the last-used export folder, sorted-output folder, format, mesh refinement, unit choice, and print utility settings.
- Uses the active Fusion document name as the default `File Name` each time the dialog opens.
- Remembers the normal `Export Folder` per Fusion project by using the document base name with version tokens stripped.
- Supports custom refinement values for:
  - surface deviation
  - normal deviation
  - maximum edge length
  - aspect ratio
- Supports STL binary export when Fusion exposes that API.
- Supports one-file-per-body and print-utility export when Fusion exposes those API options.
- Resolves F3D export to a component automatically when the selection is a body or occurrence.
- Sorts top-level staging-folder files using built-in rules for `.stl`, `.3mf`, `.obj`, `.mtl`, and `.f3d`.

## Sorting Workflow

- Turn on `Sort Automatically After Export` to make Fusion write all exported files into an automatic temporary staging folder first.
- After the export batch completes, the add-in runs a sorting pass against the top level of that temporary staging folder.
- Mesh files are deduped by version, renamed, and moved into project folders under `STLs`, `3MFs`, or `OBJs`.
- F3D files are archived under each project's `F3Ds` folder, and the highest-version cleaned copy is also placed at the top level of the project folder.
- `Simulate Sort Only` lets you test the sorter without moving or deleting files.
- `Allow Overwrite` controls whether existing sorted outputs can be replaced.
- On success, the temporary staging folder is deleted automatically. On failure, it is retained and its path is shown in the error message.

## Install

1. In Fusion, open `Utilities > Scripts and Add-Ins`.
2. Go to the `Add-Ins` tab.
3. Download and unzip `dist/BetterExport-1.0.0.zip` from the GitHub repository, or clone the repo locally.
4. Click the green `+` button and select the folder:
   `/Users/alex/Documents/Codex/BetterExport/BetterExport`
5. Run the add-in.

The command appears as `Better Export` in a `Better Export` panel on the `Utilities` ribbon tab when Fusion exposes that tab. If Fusion does not expose a Utilities tab in the current workspace, the add-in falls back to the Scripts and Add-Ins area. It is also added to Fusion's right-click context menu for exportable Browser selections like components, occurrences, and bodies.

## Notes

- If no target is selected, the add-in exports the active root component.
- Settings persistence is machine-local because it is stored in the add-in folder.
- Fusion’s API surface changes over time, so the dialog hides controls that are not supported by the installed Fusion version.

## Files

- `BetterExport/BetterExport.py`: add-in implementation
- `BetterExport/export_sorter.py`: staging-folder sorter implementation
- `BetterExport/BetterExport.manifest`: Fusion add-in manifest
- `BetterExport/settings.json`: created automatically when settings are changed or after an export
