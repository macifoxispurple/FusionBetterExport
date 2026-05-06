# Better Export Maintainer Guide

## Canonical Release Packaging (Required)

Always build release zips with the repo script:

- `scripts/package_release.py`

Do not publish archives created by OS zip tools, file explorers, or ad-hoc shell zip commands.

Reason:

- Release zips must be compatible with users on older updater code that expects a real `BetterExport/` directory after extraction.
- Canonical zip entries must use `/` path separators even when packaged on Windows.

## One-Command Build

From repo root:

- macOS/Linux: `python3 scripts/package_release.py`
- Windows (PowerShell): `python scripts/package_release.py`

This reads `BetterExport/BetterExport.manifest` and writes:

- `dist/BetterExport-<version>.zip`

You can also target an explicit output path:

- macOS/Linux: `python3 scripts/package_release.py --output dist/BetterExport-custom.zip`
- Windows (PowerShell): `python scripts/package_release.py --output dist/BetterExport-custom.zip`

## Release Artifact Sanity Check

The packaging script already validates and fails fast if any rule is violated:

- no backslashes in entry names
- no absolute paths
- no `..` path traversal segments
- every entry starts with `BetterExport/`

Manual spot-check (optional):

- `python -c "import zipfile; z=zipfile.ZipFile('dist/BetterExport-<version>.zip'); print('\\n'.join(i.filename for i in z.infolist()))"`

## Validation Before Publishing

- `python -m py_compile BetterExport/BetterExport.py BetterExport/export_sorter.py BetterExport/update_state.py`
- `python -m unittest discover -s Tests -v`

## Release Flow

1. Bump `BetterExport/BetterExport.manifest` version.
2. Build zip with `scripts/package_release.py`.
3. Run validation commands above.
4. Commit source + tests + manifest + zip.
5. Tag `v<version>`.
6. Publish GitHub release with `dist/BetterExport-<version>.zip` attached.

