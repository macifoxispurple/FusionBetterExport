#!/usr/bin/env python3
"""Build canonical Better Export release zips.

This script is the only supported packaging path for release artifacts.
It always writes POSIX zip entry names rooted under BetterExport/.
"""

from __future__ import annotations

import argparse
import os
from pathlib import Path, PurePosixPath
import zipfile


REPO_ROOT = Path(__file__).resolve().parents[1]
ADDIN_DIR = REPO_ROOT / "BetterExport"
DIST_DIR = REPO_ROOT / "dist"

EXCLUDED_FILE_NAMES = {
    "settings.json",
    "update_helper.py",
    "update_state.json",
    ".DS_Store",
}

EXCLUDED_DIR_NAMES = {
    "__pycache__",
    "_pending_update",
}


def _is_hidden_path(relative_path: Path) -> bool:
    return any(part.startswith(".") for part in relative_path.parts if part not in (".", ".."))


def _should_include(relative_path: Path) -> bool:
    if not relative_path.parts:
        return False
    if _is_hidden_path(relative_path):
        return False
    if any(part in EXCLUDED_DIR_NAMES for part in relative_path.parts):
        return False
    if relative_path.name in EXCLUDED_FILE_NAMES:
        return False
    return True


def _manifest_version(manifest_path: Path) -> str:
    import json

    with manifest_path.open("r", encoding="utf-8") as handle:
        manifest = json.load(handle)
    version_text = str(manifest.get("version", "")).strip()
    if not version_text:
        raise ValueError(f"Manifest {manifest_path} did not contain a version.")
    return version_text


def _iter_payload_files(addin_dir: Path):
    for path in sorted(addin_dir.rglob("*")):
        if not path.is_file():
            continue
        rel = path.relative_to(addin_dir)
        if _should_include(rel):
            yield path, rel


def _canonical_arcname(relative_path: Path) -> str:
    # Force POSIX separator usage regardless of host OS.
    rel_posix = PurePosixPath(*relative_path.parts).as_posix()
    return f"BetterExport/{rel_posix}"


def validate_entry_name(name: str) -> None:
    if "\\" in name:
        raise ValueError(f"Invalid zip entry (contains backslash): {name}")
    if name.startswith("/") or (len(name) >= 3 and name[1:3] == ":/"):
        raise ValueError(f"Invalid zip entry (absolute path): {name}")
    parts = [part for part in name.split("/") if part]
    if any(part == ".." for part in parts):
        raise ValueError(f"Invalid zip entry (path traversal): {name}")
    if not name.startswith("BetterExport/"):
        raise ValueError(f"Invalid zip entry (must be rooted at BetterExport/): {name}")


def validate_release_zip(zip_path: Path) -> None:
    with zipfile.ZipFile(zip_path, "r") as archive:
        for info in archive.infolist():
            validate_entry_name(info.filename)


def build_release_zip(output_zip: Path, addin_dir: Path = ADDIN_DIR) -> Path:
    output_zip.parent.mkdir(parents=True, exist_ok=True)
    if output_zip.exists():
        output_zip.unlink()

    with zipfile.ZipFile(output_zip, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for source_path, relative_path in _iter_payload_files(addin_dir):
            archive.write(source_path, arcname=_canonical_arcname(relative_path))

    validate_release_zip(output_zip)
    return output_zip


def main() -> int:
    parser = argparse.ArgumentParser(description="Build canonical Better Export release zip.")
    parser.add_argument(
        "--version",
        default="",
        help="Version for output filename. Defaults to BetterExport.manifest version.",
    )
    parser.add_argument(
        "--output",
        default="",
        help="Optional explicit output zip path. Overrides --version naming when provided.",
    )
    args = parser.parse_args()

    if args.output:
        output_zip = Path(args.output).resolve()
    else:
        version = args.version or _manifest_version(ADDIN_DIR / "BetterExport.manifest")
        output_zip = DIST_DIR / f"BetterExport-{version}.zip"

    built = build_release_zip(output_zip)
    print(str(built))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
