import re
import shutil
from pathlib import Path


VERSION_RE = re.compile(r" v([0-9]{1,5})")
VERSION_TOKEN_RE = re.compile(r" v[0-9]{1,5}(_?)")
MTLLIB_RE = re.compile(r"^(\s*mtllib\s+)(.+?)(\r?\n?)$", re.IGNORECASE)

MESH_EXTS = {".stl", ".3mf", ".obj", ".mtl"}
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


def project_name(filename):
    stem = Path(filename).stem
    part = stem.split("_", 1)[0]
    return part or "Project"


def mesh_dest_folder(ext):
    ext = ext.lower()
    if ext == ".stl":
        return "STLs"
    if ext == ".3mf":
        return "3MFs"
    if ext in {".obj", ".mtl"}:
        return "OBJs"
    raise ValueError(ext)


def _iter_top_level_files(path):
    return [entry for entry in path.iterdir() if entry.is_file()]


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


def _move_file(source, target, simulate_only, allow_overwrite):
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists() and not allow_overwrite:
        return False
    if simulate_only:
        return True
    if not _replace_existing(target, allow_overwrite):
        return False
    shutil.move(str(source), str(target))
    return True


def _copy_file(source, target, simulate_only, allow_overwrite):
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists() and not allow_overwrite:
        return False
    if simulate_only:
        return True
    if not _replace_existing(target, allow_overwrite):
        return False
    shutil.copy2(str(source), str(target))
    return True


def _rename_file(source, target, simulate_only, allow_overwrite):
    if source == target:
        return source
    if target.exists() and not allow_overwrite:
        return source
    if simulate_only:
        return target
    if not _replace_existing(target, allow_overwrite):
        return source
    source.rename(target)
    return target


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


def process_exports(input_dir, output_dir, simulate_only=False, allow_overwrite=True):
    input_dir = Path(input_dir)
    output_dir = Path(output_dir)

    result = {
        "deleted_mesh_duplicates": 0,
        "renamed_mesh_files": 0,
        "moved_mesh_files": 0,
        "archived_f3d_files": 0,
        "copied_f3d_files": 0,
        "skipped_overwrite": 0
    }

    if not input_dir.exists() or not input_dir.is_dir():
        return result

    files = _iter_top_level_files(input_dir)
    mesh_files = [path for path in files if path.suffix.lower() in MESH_EXTS]
    f3d_files = [path for path in files if path.suffix.lower() in F3D_EXTS]

    best_mesh = {}
    for path in mesh_files:
        key = (normalize_keep_key(path.name), path.suffix.lower())
        version = extract_version(path.name)
        current = best_mesh.get(key)
        if current is None or version > current[0]:
            best_mesh[key] = (version, path)

    keep_names = {path.name for _, path in best_mesh.values()}
    for path in mesh_files:
        if path.name not in keep_names and has_version_token(path.name):
            result["deleted_mesh_duplicates"] += 1
            if not simulate_only and path.exists():
                path.unlink()

    mesh_files = [path for path in _iter_top_level_files(input_dir) if path.suffix.lower() in MESH_EXTS]
    renamed_mesh_files = []
    for path in sorted(mesh_files, key=lambda entry: entry.name.lower()):
        new_name = normalize_final_name(path.name)
        target = path.with_name(new_name)
        renamed_path = _rename_file(path, target, simulate_only, allow_overwrite)
        _rewrite_obj_mtllib_reference(renamed_path, simulate_only)
        if renamed_path.name != path.name:
            result["renamed_mesh_files"] += 1
        renamed_mesh_files.append(renamed_path)

    for path in renamed_mesh_files:
        proj = project_name(path.name)
        destination = output_dir / proj / mesh_dest_folder(path.suffix) / path.name
        moved = _move_file(path, destination, simulate_only, allow_overwrite)
        if moved:
            result["moved_mesh_files"] += 1
        else:
            result["skipped_overwrite"] += 1

    best_f3d = {}
    for path in f3d_files:
        cleaned_name = normalize_final_name(path.name)
        version = extract_version(path.name)
        current = best_f3d.get(cleaned_name)
        if current is None or version > current[0]:
            best_f3d[cleaned_name] = (version, path.name)

    best_clean_names = {name for _, name in best_f3d.values()}
    for path in f3d_files:
        proj = project_name(normalize_final_name(path.name))
        proj_dir = output_dir / proj
        archive_target = proj_dir / "F3Ds" / path.name
        moved = _move_file(path, archive_target, simulate_only, allow_overwrite)
        if moved:
            result["archived_f3d_files"] += 1
        else:
            result["skipped_overwrite"] += 1
            continue

        if path.name in best_clean_names:
            clean_target = proj_dir / normalize_final_name(path.name)
            copied = _copy_file(archive_target, clean_target, simulate_only, allow_overwrite)
            if copied:
                result["copied_f3d_files"] += 1
            else:
                result["skipped_overwrite"] += 1

    return result
