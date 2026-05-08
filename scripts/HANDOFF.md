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

## Release Notes Responsibility (Required At Ship Start)

- Agent drafts release notes first from the actual diff/context before final ship.
- If commit history is low-signal or changes are substantial, use curated notes via `--notes-file`.
- Release notes style is mandatory for every release:
  - Dense, high-level, end-user readable summary of actual fixes/features.
  - No internal implementation detail, test logs, or tool/process narration.
  - Never include secrets or environment-specific operational details (API keys, SSH keys, tokens, auth identities, local usernames, machine names, local filesystem paths, or private infrastructure hostnames/IPs).
  - Maximum 3 non-empty lines total (typically: 1 header + up to 2 bullets).
  - Prefer plain-language outcomes (what changed for users and why it matters).
- Recommended command when curated notes are prepared:
  - `gh release create v<version> dist/BetterExport-<version>.zip --repo macifoxispurple/FusionBetterExport --title "Better Export v<version>" --notes-file dist/release-notes-v<version>.md`

## Release Flow

1. Bump `BetterExport/BetterExport.manifest` version.
2. Build zip with `scripts/package_release.py`.
3. Draft release notes from actual diff/context and finalize curated notes file if needed.
4. Run validation commands above.
5. Commit source + tests + manifest + zip.
6. Tag `v<version>`.
7. Publish GitHub release with `dist/BetterExport-<version>.zip` attached.
