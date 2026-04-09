import re
import shutil
from pathlib import Path


VERSION_RE = re.compile(r" v([0-9]{1,5})")
VERSION_TOKEN_RE = re.compile(r" v[0-9]{1,5}(_?)")
MTLLIB_RE = re.compile(r"^(\s*mtllib\s+)(.+?)(\r?\n?)$", re.IGNORECASE)

MESH_EXTS = {".stl", ".3mf", ".obj", ".mtl"}
CAD_EXTS = {".iges", ".igs", ".sat", ".smt", ".step", ".stp", ".usdz"}
F3D_EXTS = {".f3d"}


def extract_version(name):
    match = VERSION_RE.search(name)
    return int(match.group(1)) if match else 0


def has_version_token(name):
    return bool(VERSION_RE.search(name))


def normalize_keep_key(name):
    return VERSION_TOKEN_RE.sub(lambda match: match.group(1), name, count=1)


def normalize_final_name(name):
    return VERSION_TOKEN_RE.sub(lambda match: match.group(1), name).replace(" ", "")


def sorted_final_name(name, strip_version_numbers=True):
    return normalize_final_name(name) if strip_version_numbers else name


def project_name(filename):
    stem = Path(filename).stem
    part = stem.split("_", 1)[0]
    return part or "Project"


def export_dest_folder(ext):
    ext = ext.lower()
    if ext == ".stl":
        return "STL"
    if ext == ".3mf":
        return "3MF"
    if ext in {".obj", ".mtl"}:
        return "OBJ"
    if ext in {".iges", ".igs"}:
        return "IGES"
    if ext == ".sat":
        return "SAT"
    if ext == ".smt":
        return "SMT"
    if ext in {".step", ".stp"}:
        return "STEP"
    if ext == ".usdz":
        return "USDZ"
    raise ValueError(ext)


def _iter_top_level_files(path):
    return [entry for entry in path.iterdir() if entry.is_file()]


def _format_conflict_message(operation, source, target):
    return f"Conflict during {operation}: incoming '{source.name}' would overwrite '{target.name}' at '{target}'."


def _unique_conflict_target(target):
    if not target.exists():
        return target

    stem = target.stem
    suffix = target.suffix
    counter = 2
    while True:
        candidate = target.with_name(f"{stem}_copy{counter}{suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def _replace_existing(target, allow_overwrite):
    if not target.exists():
        return True
    if not allow_overwrite:
        return False
    if target.is_file():
        target.unlink()
        return True
    shutil.rmtree(target)
    return True


def _resolve_conflict_target(source, target, operation, allow_overwrite, conflict_resolver, keep_both_target=None):
    if not target.exists():
        return target, "target"

    if allow_overwrite:
        return target, "overwrite"

    if not conflict_resolver:
        raise FileExistsError(_format_conflict_message(operation, source, target))

    action = conflict_resolver(source, target, operation, keep_both_target or target)
    if action == "overwrite":
        return target, "overwrite"
    if action == "keep_both":
        candidate = keep_both_target or target
        if candidate == source:
            return source, "keep_both"
        return _unique_conflict_target(candidate), "keep_both"
    if action == "skip":
        return target, "skip"

    raise ValueError(f"Unsupported conflict action: {action}")


def _move_file(source, target, simulate_only, allow_overwrite, conflict_resolver=None, keep_both_target=None):
    target.parent.mkdir(parents=True, exist_ok=True)
    resolved_target, action = _resolve_conflict_target(
        source,
        target,
        "move",
        allow_overwrite,
        conflict_resolver,
        keep_both_target
    )
    if action == "skip":
        return False, "skip", target
    if simulate_only:
        return True, action, resolved_target
    if action == "overwrite" and not _replace_existing(resolved_target, True):
        return False, "skip", resolved_target
    shutil.move(str(source), str(resolved_target))
    return True, action, resolved_target


def _copy_file(source, target, simulate_only, allow_overwrite, conflict_resolver=None, keep_both_target=None):
    target.parent.mkdir(parents=True, exist_ok=True)
    resolved_target, action = _resolve_conflict_target(
        source,
        target,
        "copy",
        allow_overwrite,
        conflict_resolver,
        keep_both_target
    )
    if action == "skip":
        return False, "skip", target
    if simulate_only:
        return True, action, resolved_target
    if action == "overwrite" and not _replace_existing(resolved_target, True):
        return False, "skip", resolved_target
    shutil.copy2(str(source), str(resolved_target))
    return True, action, resolved_target


def _rename_file(source, target, simulate_only, allow_overwrite, conflict_resolver=None):
    if source == target:
        return source, "unchanged"
    resolved_target, action = _resolve_conflict_target(
        source,
        target,
        "rename",
        allow_overwrite,
        conflict_resolver,
        source
    )
    if action == "skip":
        return source, "skip"
    if resolved_target == source:
        return source, "keep_both"
    if simulate_only:
        return resolved_target, action
    if action == "overwrite" and not _replace_existing(resolved_target, True):
        return source, "skip"
    source.rename(resolved_target)
    return resolved_target, action


def _rewrite_obj_mtllib_reference(path, simulate_only):
    if path.suffix.lower() != ".obj" or not path.exists() or simulate_only:
        return False

    original_text = path.read_text(encoding="utf-8", errors="surrogateescape", newline="")
    changed = False
    rewritten_lines = []

    for line in original_text.splitlines(keepends=True):
        match = MTLLIB_RE.match(line)
        if not match:
            rewritten_lines.append(line)
            continue

        prefix, referenced_name, line_ending = match.groups()
        normalized_reference = normalize_final_name(referenced_name)
        if normalized_reference != referenced_name:
            rewritten_lines.append(f"{prefix}{normalized_reference}{line_ending}")
            changed = True
        else:
            rewritten_lines.append(line)

    if not changed:
        return False

    path.write_text("".join(rewritten_lines), encoding="utf-8", errors="surrogateescape", newline="")
    return True


def scan_export_conflicts(input_dir, output_dir, strip_version_numbers=True):
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)

    if not input_dir.exists() or not input_dir.is_dir():
        return []

    conflicts = []
    files = _iter_top_level_files(input_dir)
    sortable_files = [path for path in files if path.suffix.lower() in MESH_EXTS.union(CAD_EXTS)]
    f3d_files = [path for path in files if path.suffix.lower() in F3D_EXTS]

    best_sortable = {}
    for path in sortable_files:
        key = (normalize_keep_key(path.name), path.suffix.lower())
        version = extract_version(path.name)
        current = best_sortable.get(key)
        if current is None or version > current[0]:
            best_sortable[key] = (version, path)

    keep_names = {path.name for _, path in best_sortable.values()}
    planned_sortable = []
    for path in sortable_files:
        if path.name not in keep_names and has_version_token(path.name):
            continue
        planned_sortable.append({
            "original_name": path.name,
            "final_name": sorted_final_name(path.name, strip_version_numbers),
            "suffix": path.suffix
        })

    for entry in planned_sortable:
        final_name = entry["final_name"]
        proj = project_name(normalize_final_name(entry["original_name"]))
        destination = output_dir / proj / export_dest_folder(entry["suffix"]) / final_name
        if destination.exists():
            conflicts.append({
                "operation": "move",
                "incoming_name": final_name,
                "existing_name": destination.name,
                "target_path": str(destination),
                "keep_both_name": entry["original_name"]
            })

    best_f3d = {}
    for path in f3d_files:
        cleaned_name = normalize_keep_key(path.name)
        version = extract_version(path.name)
        current = best_f3d.get(cleaned_name)
        if current is None or version > current[0]:
            best_f3d[cleaned_name] = (version, path.name)

    best_clean_names = {name for _, name in best_f3d.values()}
    for path in f3d_files:
        proj = project_name(normalize_final_name(path.name))
        proj_dir = output_dir / proj

        archive_target = proj_dir / "F3D" / path.name
        if archive_target.exists():
            conflicts.append({
                "operation": "move",
                "incoming_name": path.name,
                "existing_name": archive_target.name,
                "target_path": str(archive_target),
                "keep_both_name": path.name
            })

        if path.name in best_clean_names:
            clean_name = sorted_final_name(path.name, strip_version_numbers)
            clean_target = proj_dir / clean_name
            if clean_target.exists():
                conflicts.append({
                    "operation": "copy",
                    "incoming_name": clean_name,
                    "existing_name": clean_target.name,
                    "target_path": str(clean_target),
                    "keep_both_name": path.name
                })

    return conflicts


def process_exports(input_dir, output_dir, simulate_only=False, allow_overwrite=True, conflict_resolver=None, strip_version_numbers=True):
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)

    result = {
        "deleted_mesh_duplicates": 0,
        "renamed_mesh_files": 0,
        "moved_mesh_files": 0,
        "archived_f3d_files": 0,
        "copied_f3d_files": 0,
        "skipped_overwrite": 0,
        "conflicts_overwritten": 0,
        "conflicts_kept_both": 0,
        "conflicts_skipped": 0
    }

    if not input_dir.exists() or not input_dir.is_dir():
        return result

    files = _iter_top_level_files(input_dir)
    sortable_files = [path for path in files if path.suffix.lower() in MESH_EXTS.union(CAD_EXTS)]
    f3d_files = [path for path in files if path.suffix.lower() in F3D_EXTS]

    best_sortable = {}
    for path in sortable_files:
        key = (normalize_keep_key(path.name), path.suffix.lower())
        version = extract_version(path.name)
        current = best_sortable.get(key)
        if current is None or version > current[0]:
            best_sortable[key] = (version, path)

    keep_names = {path.name for _, path in best_sortable.values()}
    for path in sortable_files:
        if path.name not in keep_names and has_version_token(path.name):
            result["deleted_mesh_duplicates"] += 1
            if not simulate_only and path.exists():
                path.unlink()

    sortable_files = [path for path in _iter_top_level_files(input_dir) if path.suffix.lower() in MESH_EXTS.union(CAD_EXTS)]
    renamed_sortable_files = []
    for path in sorted(sortable_files, key=lambda entry: entry.name.lower()):
        original_name = path.name
        new_name = sorted_final_name(path.name, strip_version_numbers)
        target = path.with_name(new_name)
        renamed_path, rename_action = _rename_file(path, target, simulate_only, allow_overwrite, conflict_resolver)
        if strip_version_numbers:
            _rewrite_obj_mtllib_reference(renamed_path, simulate_only)
        if renamed_path.name != path.name:
            result["renamed_mesh_files"] += 1
        if rename_action == "overwrite":
            result["conflicts_overwritten"] += 1
        elif rename_action == "keep_both":
            result["conflicts_kept_both"] += 1
        elif rename_action == "skip":
            result["conflicts_skipped"] += 1
        renamed_sortable_files.append((renamed_path, original_name))

    for path, original_name in renamed_sortable_files:
        proj = project_name(normalize_final_name(path.name))
        destination = output_dir / proj / export_dest_folder(path.suffix) / path.name
        keep_both_destination = destination.with_name(original_name)
        moved, move_action, _ = _move_file(
            path,
            destination,
            simulate_only,
            allow_overwrite,
            conflict_resolver,
            keep_both_destination
        )
        if moved:
            result["moved_mesh_files"] += 1
            if move_action == "overwrite":
                result["conflicts_overwritten"] += 1
            elif move_action == "keep_both":
                result["conflicts_kept_both"] += 1
        else:
            result["skipped_overwrite"] += 1
            if move_action == "skip":
                result["conflicts_skipped"] += 1

    best_f3d = {}
    for path in f3d_files:
        cleaned_name = normalize_keep_key(path.name)
        version = extract_version(path.name)
        current = best_f3d.get(cleaned_name)
        if current is None or version > current[0]:
            best_f3d[cleaned_name] = (version, path.name)

    best_clean_names = {name for _, name in best_f3d.values()}
    for path in f3d_files:
        proj = project_name(normalize_final_name(path.name))
        proj_dir = output_dir / proj
        archive_target = proj_dir / "F3D" / path.name
        moved, move_action, archived_path = _move_file(
            path,
            archive_target,
            simulate_only,
            allow_overwrite,
            conflict_resolver,
            archive_target.with_name(path.name)
        )
        if moved:
            result["archived_f3d_files"] += 1
            if move_action == "overwrite":
                result["conflicts_overwritten"] += 1
            elif move_action == "keep_both":
                result["conflicts_kept_both"] += 1
        else:
            result["skipped_overwrite"] += 1
            if move_action == "skip":
                result["conflicts_skipped"] += 1
            continue

        if path.name in best_clean_names:
            clean_target = proj_dir / sorted_final_name(path.name, strip_version_numbers)
            keep_both_clean_target = proj_dir / archived_path.name
            copied, copy_action, _ = _copy_file(
                archived_path,
                clean_target,
                simulate_only,
                allow_overwrite,
                conflict_resolver,
                keep_both_clean_target
            )
            if copied:
                result["copied_f3d_files"] += 1
                if copy_action == "overwrite":
                    result["conflicts_overwritten"] += 1
                elif copy_action == "keep_both":
                    result["conflicts_kept_both"] += 1
            else:
                result["skipped_overwrite"] += 1
                if copy_action == "skip":
                    result["conflicts_skipped"] += 1

    return result
