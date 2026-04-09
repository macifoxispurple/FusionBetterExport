import json
import importlib.util
import os
import shutil
import sys
import tempfile
import time
import traceback
import urllib.error
import urllib.request
import zipfile

import adsk.core
import adsk.fusion

ADDIN_DIR = os.path.dirname(os.path.abspath(__file__))
if ADDIN_DIR not in sys.path:
    sys.path.insert(0, ADDIN_DIR)

from export_sorter import VERSION_TOKEN_RE, process_exports, scan_export_conflicts


COMMAND_ID = 'betterMeshExportCommand'
COMMAND_NAME = 'Better Export'
COMMAND_DESCRIPTION = 'Export STL, OBJ, 3MF, and F3D with persistent settings.'
WORKSPACE_ID = 'FusionSolidEnvironment'
FALLBACK_PANEL_ID = 'SolidScriptsAddinsPanel'
UTILITIES_PANEL_ID = 'BetterExportPanel'
UTILITIES_PANEL_NAME = 'Better Export'
UTILITIES_TAB_CANDIDATE_IDS = [
    'UtilitiesTab',
    'SolidUtilitiesTab',
    'ToolsTab'
]

SETTINGS_PATH = os.path.join(os.path.dirname(__file__), 'settings.json')
MANIFEST_PATH = os.path.join(os.path.dirname(__file__), 'BetterExport.manifest')
LATEST_RELEASE_API_URL = 'https://api.github.com/repos/macifoxispurple/FusionBetterExport/releases/latest'
LATEST_RELEASE_PAGE_URL = 'https://github.com/macifoxispurple/FusionBetterExport/releases/latest'
UPDATE_CACHE_MAX_AGE_SECONDS = 5 * 60
PENDING_UPDATE_DIR = os.path.join(ADDIN_DIR, '_pending_update')
PENDING_UPDATE_INFO_PATH = os.path.join(PENDING_UPDATE_DIR, 'update.json')
UPDATE_HELPER_PATH = os.path.join(ADDIN_DIR, 'update_helper.py')

FORMAT_LABELS = {
    'stl': 'STL',
    'obj': 'OBJ',
    '3mf': '3MF',
    'f3d': 'F3D'
}

MESH_REFINEMENT_LABELS = {
    'high': 'High',
    'medium': 'Medium',
    'low': 'Low',
    'custom': 'Custom'
}

UNIT_LABELS = {
    'default': 'Use design default',
    'mm': 'Millimeters',
    'cm': 'Centimeters',
    'm': 'Meters',
    'in': 'Inches',
    'ft': 'Feet'
}

SETTINGS_MODE_LABELS = {
    'global': 'Global',
    'per_format': 'Per Format'
}

TARGET_MODE_LABELS = {
    'full_design': 'Export Full Design',
    'visible_bodies': 'Export Only Visible Bodies',
    'selection': 'Export Selection'
}

DESTINATION_MODE_LABELS = {
    'direct': 'Direct Export',
    'sorted': 'Sort Into Project Folders'
}

OPTION_DEFAULTS = {
    'filename': '',
    'mesh_refinement': 'medium',
    'surface_deviation_cm': '0.1',
    'normal_deviation_rad': '0.523599',
    'maximum_edge_length_cm': '0.1',
    'aspect_ratio': '5.0',
    'unit_type': 'default',
    'binary_format': True,
    'one_file_per_body': False,
    'send_to_print_utility': False,
    'print_utility_mode': 'default',
    'print_utility_value': ''
}

GENERAL_DEFAULTS = {
    'folder': os.path.expanduser('~'),
    'sorted_output_folder': os.path.expanduser('~'),
    'auto_sort_after_export': False,
    'always_export_full_root': False,
    'target_mode': 'selection',
    'f3d_enabled_preference': False,
    'auto_check_updates': True,
    'allow_overwrite': True,
    'project_export_folders': {},
    'project_auto_sort_preferences': {},
    'update_check': {}
}

DEFAULT_SETTINGS = {
    'formats': ['stl'],
    'settings_mode': 'global',
    **GENERAL_DEFAULTS,
    **OPTION_DEFAULTS,
    'per_format_settings': {}
}

MESH_REFINEMENT_KEYS_BY_LABEL = {label: key for key, label in MESH_REFINEMENT_LABELS.items()}
UNIT_KEYS_BY_LABEL = {label: key for key, label in UNIT_LABELS.items()}

_app = None
_ui = None
_handlers = []
_updated_runtime_module = None


def _safe_call(fn):
    try:
        return fn()
    except Exception:
        return None


def _toolbar_tab_by_name(workspace, expected_name):
    tabs = _safe_call(lambda: workspace.toolbarTabs)
    if not tabs:
        return None

    count = _safe_call(lambda: tabs.count) or 0
    for index in range(count):
        tab = _safe_call(lambda i=index: tabs.item(i))
        if tab and _safe_call(lambda t=tab: t.name) == expected_name:
            return tab
    return None


def _target_toolbar_panel(workspace):
    if not workspace:
        return None

    toolbar_tabs = _safe_call(lambda: workspace.toolbarTabs)
    tab = None
    if toolbar_tabs:
        for candidate_id in UTILITIES_TAB_CANDIDATE_IDS:
            tab = _safe_call(lambda cid=candidate_id: toolbar_tabs.itemById(cid))
            if tab:
                break

        if not tab:
            tab = _toolbar_tab_by_name(workspace, 'Utilities')

    if tab:
        panel = _safe_call(lambda: tab.toolbarPanels.itemById(UTILITIES_PANEL_ID))
        if not panel:
            panel = _safe_call(lambda: tab.toolbarPanels.add(UTILITIES_PANEL_ID, UTILITIES_PANEL_NAME))
        return panel

    return _safe_call(lambda: workspace.toolbarPanels.itemById(FALLBACK_PANEL_ID))


def _supports_attr(obj, attr_name):
    try:
        getattr(obj, attr_name)
        return True
    except Exception:
        return False


def _supports_export_selection(entity):
    return bool(
        adsk.fusion.BRepBody.cast(entity) or
        adsk.fusion.Occurrence.cast(entity) or
        adsk.fusion.Component.cast(entity)
    )


def _merge_settings(values):
    merged = dict(DEFAULT_SETTINGS)
    merged.update(values or {})
    legacy_full_root = bool(merged.get('always_export_full_root'))
    merged['target_mode'] = merged.get('target_mode')
    if merged['target_mode'] not in TARGET_MODE_LABELS:
        merged['target_mode'] = 'full_design' if legacy_full_root else 'selection'
    merged['always_export_full_root'] = merged['target_mode'] == 'full_design'
    merged['formats'] = _normalized_formats(merged.get('formats'), merged.get('format'))
    if isinstance(values, dict) and 'f3d_enabled_preference' in values:
        merged['f3d_enabled_preference'] = bool(values.get('f3d_enabled_preference'))
    else:
        merged['f3d_enabled_preference'] = 'f3d' in merged['formats']
    merged['settings_mode'] = merged['settings_mode'] if merged.get('settings_mode') in SETTINGS_MODE_LABELS else 'global'
    merged['per_format_settings'] = _normalized_per_format_settings(merged.get('per_format_settings'))
    merged['project_export_folders'] = _normalized_project_export_folders(merged.get('project_export_folders'))
    merged['project_auto_sort_preferences'] = _normalized_project_auto_sort_preferences(merged.get('project_auto_sort_preferences'))
    merged['update_check'] = _normalized_update_check(merged.get('update_check'))
    for key, default_value in GENERAL_DEFAULTS.items():
        merged[key] = merged.get(key, default_value)
    for key, default_value in OPTION_DEFAULTS.items():
        merged[key] = merged.get(key, default_value)
    return merged


def _load_settings():
    if not os.path.exists(SETTINGS_PATH):
        settings = dict(DEFAULT_SETTINGS)
        settings['folder'] = _folder_for_current_project(settings)
        settings['auto_sort_after_export'] = _auto_sort_for_current_project(settings)
        return settings

    try:
        with open(SETTINGS_PATH, 'r', encoding='utf-8') as handle:
            settings = _merge_settings(json.load(handle))
            settings['folder'] = _folder_for_current_project(settings)
            settings['auto_sort_after_export'] = _auto_sort_for_current_project(settings)
            return settings
    except Exception:
        settings = dict(DEFAULT_SETTINGS)
        settings['folder'] = _folder_for_current_project(settings)
        settings['auto_sort_after_export'] = _auto_sort_for_current_project(settings)
        return settings


def _load_settings_for_save():
    if not os.path.exists(SETTINGS_PATH):
        return _merge_settings({})

    try:
        with open(SETTINGS_PATH, 'r', encoding='utf-8') as handle:
            return _merge_settings(json.load(handle))
    except Exception:
        return _merge_settings({})


def _save_settings(values):
    existing_settings = _load_settings_for_save()
    settings = _merge_settings(existing_settings)
    settings.update(values or {})
    settings.pop('simulate_sort_only', None)
    settings['formats'] = _normalized_formats(settings.get('formats'), settings.get('format'))
    settings['settings_mode'] = settings['settings_mode'] if settings.get('settings_mode') in SETTINGS_MODE_LABELS else 'global'
    settings['per_format_settings'] = _normalized_per_format_settings(settings.get('per_format_settings'))
    settings['project_export_folders'] = _normalized_project_export_folders(settings.get('project_export_folders'))
    settings['project_auto_sort_preferences'] = _normalized_project_auto_sort_preferences(settings.get('project_auto_sort_preferences'))
    settings['update_check'] = _normalized_update_check(settings.get('update_check'))
    project_key = _current_project_key()
    if project_key and settings.get('folder'):
        settings['project_export_folders'][project_key] = settings['folder']
    if settings.get('folder'):
        settings['project_export_folders']['recent'] = settings['folder']
    settings.pop('project_sorted_output_folders', None)
    if project_key:
        settings['project_auto_sort_preferences'][project_key] = bool(settings.get('auto_sort_after_export'))
    settings['project_auto_sort_preferences']['recent'] = bool(settings.get('auto_sort_after_export'))
    settings['filename'] = ''
    for format_key in FORMAT_LABELS:
        if format_key in settings['per_format_settings']:
            settings['per_format_settings'][format_key]['filename'] = ''
    with open(SETTINGS_PATH, 'w', encoding='utf-8') as handle:
        json.dump(settings, handle, indent=2, sort_keys=True)


def _normalized_per_format_settings(value):
    source = value if isinstance(value, dict) else {}
    normalized = {}
    for format_key in FORMAT_LABELS:
        format_values = dict(OPTION_DEFAULTS)
        candidate = source.get(format_key)
        if isinstance(candidate, dict):
            format_values.update(candidate)
        normalized[format_key] = format_values
    return normalized


def _normalized_project_export_folders(value):
    if not isinstance(value, dict):
        return {}
    normalized = {}
    for key, folder in value.items():
        if isinstance(key, str) and isinstance(folder, str) and key.strip() and folder.strip():
            normalized[key.strip()] = folder.strip()
    return normalized


def _normalized_project_auto_sort_preferences(value):
    if not isinstance(value, dict):
        return {}
    normalized = {}
    for key, enabled in value.items():
        if isinstance(key, str) and key.strip():
            normalized[key.strip()] = bool(enabled)
    return normalized


def _normalized_update_check(value):
    if not isinstance(value, dict):
        return {}
    normalized = {}
    if isinstance(value.get('checked_at'), (int, float)):
        normalized['checked_at'] = float(value['checked_at'])
    if isinstance(value.get('latest_version'), str):
        normalized['latest_version'] = value['latest_version'].strip()
    if isinstance(value.get('latest_url'), str):
        normalized['latest_url'] = value['latest_url'].strip()
    if isinstance(value.get('latest_asset_url'), str):
        normalized['latest_asset_url'] = value['latest_asset_url'].strip()
    if isinstance(value.get('latest_asset_name'), str):
        normalized['latest_asset_name'] = value['latest_asset_name'].strip()
    if isinstance(value.get('error'), str):
        normalized['error'] = value['error'].strip()
    return normalized


def _active_design():
    product = _app.activeProduct
    return adsk.fusion.Design.cast(product)


def _root_component():
    design = _active_design()
    return design.rootComponent if design else None


def _default_filename():
    design = _active_design()
    if design and design.parentDocument:
        name = design.parentDocument.name or 'mesh-export'
        return _sanitize_filename(name)
    return 'mesh-export'


def _current_project_key():
    design = _active_design()
    if not design or not design.parentDocument:
        return ''

    document_name = design.parentDocument.name or ''
    stem = os.path.splitext(document_name)[0]
    key = VERSION_TOKEN_RE.sub(lambda match: match.group(1), stem).strip()
    return key or _sanitize_filename(stem)


def _folder_for_current_project(settings):
    project_key = _current_project_key()
    project_folders = settings.get('project_export_folders', {})
    if project_key and project_key in project_folders:
        return project_folders[project_key]
    if 'recent' in project_folders:
        return project_folders['recent']
    return settings.get('folder', GENERAL_DEFAULTS['folder'])


def _auto_sort_for_current_project(settings):
    project_key = _current_project_key()
    project_preferences = settings.get('project_auto_sort_preferences', {})
    if project_key and project_key in project_preferences:
        return bool(project_preferences[project_key])
    if 'recent' in project_preferences:
        return bool(project_preferences['recent'])
    return bool(settings.get('auto_sort_after_export', GENERAL_DEFAULTS['auto_sort_after_export']))


def _short_path(path_value):
    if not path_value:
        return ''

    normalized = os.path.normpath(path_value)
    parts = normalized.split(os.sep)
    if len(parts) <= 2:
        return normalized
    return '...{}{}'.format(os.sep, os.sep.join(parts[-2:]))


def _sanitize_filename(name):
    invalid = '<>:"/\\|?*'
    sanitized = ''.join('_' if char in invalid else char for char in (name or '').strip())
    sanitized = sanitized.rstrip('. ')
    return sanitized or 'mesh-export'


def _format_extension(format_key):
    return '3mf' if format_key == '3mf' else format_key


def _current_addin_version():
    try:
        with open(MANIFEST_PATH, 'r', encoding='utf-8') as handle:
            return str(json.load(handle).get('version', '')).strip() or '0.0.0'
    except Exception:
        return '0.0.0'


def _version_parts(version_text):
    text = (version_text or '').strip().lower()
    if text.startswith('v'):
        text = text[1:]
    parts = []
    for part in text.split('.'):
        digits = ''.join(ch for ch in part if ch.isdigit())
        parts.append(int(digits or '0'))
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:3])


def _is_version_newer(candidate_version, current_version):
    return _version_parts(candidate_version) > _version_parts(current_version)


def _save_update_check(update_check):
    settings = _load_settings_for_save()
    settings['update_check'] = _normalized_update_check(update_check)
    with open(SETTINGS_PATH, 'w', encoding='utf-8') as handle:
        json.dump(_merge_settings(settings), handle, indent=2, sort_keys=True)


def _upgrade_settings_file():
    settings = _load_settings()
    _save_settings(settings)
    return _load_settings()


def _release_zip_asset(payload):
    assets = payload.get('assets') or []
    zip_assets = [asset for asset in assets if str(asset.get('name', '')).lower().endswith('.zip')]
    for asset in zip_assets:
        name = str(asset.get('name') or '')
        if name.lower().startswith('betterexport-'):
            return asset
    return zip_assets[0] if zip_assets else {}


def _fetch_latest_release_info():
    request = urllib.request.Request(
        LATEST_RELEASE_API_URL,
        headers={
            'Accept': 'application/vnd.github+json',
            'User-Agent': 'BetterExport',
            'Cache-Control': 'no-cache'
        }
    )
    with urllib.request.urlopen(request, timeout=4) as response:
        payload = json.loads(response.read().decode('utf-8'))

    latest_version = str(payload.get('tag_name') or payload.get('name') or '').strip()
    if latest_version.lower().startswith('v'):
        latest_version = latest_version[1:]

    latest_url = str(payload.get('html_url') or LATEST_RELEASE_PAGE_URL).strip() or LATEST_RELEASE_PAGE_URL
    asset = _release_zip_asset(payload)
    latest_asset_url = str(asset.get('browser_download_url') or '').strip()
    latest_asset_name = str(asset.get('name') or '').strip()
    if not latest_version:
        raise ValueError('GitHub did not return a release version.')

    return {
        'checked_at': time.time(),
        'latest_version': latest_version,
        'latest_url': latest_url,
        'latest_asset_url': latest_asset_url,
        'latest_asset_name': latest_asset_name,
        'error': ''
    }


def _latest_release_info(force_refresh=False, allow_cached_on_error=True):
    settings = _load_settings_for_save()
    cached = _normalized_update_check(settings.get('update_check'))
    checked_at = cached.get('checked_at', 0)
    is_fresh = bool(checked_at and (time.time() - checked_at) < UPDATE_CACHE_MAX_AGE_SECONDS)

    if cached and not force_refresh and is_fresh:
        return cached

    try:
        latest = _fetch_latest_release_info()
        _save_update_check(latest)
        return latest
    except Exception as exc:
        if cached and allow_cached_on_error:
            cached['error'] = str(exc)
            return cached
        return {
            'checked_at': time.time(),
            'latest_version': '',
            'latest_url': LATEST_RELEASE_PAGE_URL,
            'latest_asset_url': '',
            'latest_asset_name': '',
            'error': str(exc)
        }


def _pending_update_version():
    if not os.path.exists(PENDING_UPDATE_INFO_PATH):
        return ''
    try:
        with open(PENDING_UPDATE_INFO_PATH, 'r', encoding='utf-8') as handle:
            payload = json.load(handle)
        return str(payload.get('latest_version') or '').strip()
    except Exception:
        return ''


def _download_release_asset(asset_url, destination_path):
    request = urllib.request.Request(
        asset_url,
        headers={
            'User-Agent': 'BetterExport',
            'Cache-Control': 'no-cache'
        }
    )
    with urllib.request.urlopen(request, timeout=20) as response, open(destination_path, 'wb') as handle:
        shutil.copyfileobj(response, handle)


def _find_extracted_addin_dir(extract_root):
    direct = os.path.join(extract_root, 'BetterExport')
    if os.path.isdir(direct):
        return direct

    for entry in os.listdir(extract_root):
        candidate = os.path.join(extract_root, entry, 'BetterExport')
        if os.path.isdir(candidate):
            return candidate

    return ''


def _updater_script_contents():
    return r'''import os
import shutil


def apply_update(source_dir, target_dir, skip_names=None):
    skip_names = set(skip_names or [])
    os.makedirs(target_dir, exist_ok=True)
    for name in os.listdir(source_dir):
        if name in skip_names:
            continue
        source_path = os.path.join(source_dir, name)
        target_path = os.path.join(target_dir, name)
        if os.path.isdir(source_path):
            os.makedirs(target_path, exist_ok=True)
            apply_update(source_path, target_path, skip_names=None)
        else:
            os.makedirs(os.path.dirname(target_path), exist_ok=True)
            shutil.copy2(source_path, target_path)


if __name__ == '__main__':
    import sys
    apply_update(sys.argv[1], sys.argv[2], set(sys.argv[3:]))
'''


def _write_update_helper():
    with open(UPDATE_HELPER_PATH, 'w', encoding='utf-8') as handle:
        handle.write(_updater_script_contents())


def _script_item_for_addin():
    scripts = _safe_call(lambda: _app.scripts)
    if not scripts:
        return None
    return _safe_call(lambda: scripts.itemByPath(ADDIN_DIR))


def _set_run_on_startup(enabled):
    script_item = _script_item_for_addin()
    if not script_item or not bool(_safe_call(lambda: script_item.isAddIn)):
        raise RuntimeError('Fusion could not find Better Export as an add-in.')
    script_item.isRunOnStartup = bool(enabled)


def _stage_update_payload(release_info):
    latest_version = release_info.get('latest_version', '')
    asset_url = release_info.get('latest_asset_url', '')
    asset_name = release_info.get('latest_asset_name') or 'BetterExport-{}.zip'.format(latest_version or 'update')

    if not asset_url:
        raise ValueError('No downloadable release package was found for the latest version.')

    if os.path.isdir(PENDING_UPDATE_DIR):
        shutil.rmtree(PENDING_UPDATE_DIR, ignore_errors=True)
    os.makedirs(PENDING_UPDATE_DIR, exist_ok=True)

    zip_path = os.path.join(PENDING_UPDATE_DIR, asset_name)
    extract_root = os.path.join(PENDING_UPDATE_DIR, 'extracted')
    os.makedirs(extract_root, exist_ok=True)
    _download_release_asset(asset_url, zip_path)
    with zipfile.ZipFile(zip_path, 'r') as archive:
        archive.extractall(extract_root)

    extracted_addin_dir = _find_extracted_addin_dir(extract_root)
    if not extracted_addin_dir:
        raise ValueError('The downloaded release package did not contain a BetterExport add-in folder.')

    _write_update_helper()
    _set_run_on_startup(True)

    update_info = {
        'latest_version': latest_version,
        'staged_addin_dir': extracted_addin_dir,
        'staged_at': time.time()
    }
    with open(PENDING_UPDATE_INFO_PATH, 'w', encoding='utf-8') as handle:
        json.dump(update_info, handle, indent=2, sort_keys=True)
    return update_info


def _apply_pending_update_if_needed():
    if not os.path.exists(PENDING_UPDATE_INFO_PATH) or not os.path.exists(UPDATE_HELPER_PATH):
        return None

    try:
        with open(PENDING_UPDATE_INFO_PATH, 'r', encoding='utf-8') as handle:
            update_info = json.load(handle)
        staged_addin_dir = str(update_info.get('staged_addin_dir') or '').strip()
        latest_version = str(update_info.get('latest_version') or '').strip()
        if not staged_addin_dir or not os.path.isdir(staged_addin_dir):
            raise ValueError('The staged update files are missing.')

        spec = importlib.util.spec_from_file_location('better_export_update_helper', UPDATE_HELPER_PATH)
        if not spec or not spec.loader:
            raise RuntimeError('Could not load the update helper.')
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        module.apply_update(staged_addin_dir, ADDIN_DIR, {'settings.json', os.path.basename(UPDATE_HELPER_PATH), os.path.basename(PENDING_UPDATE_DIR)})

        pycache_dir = os.path.join(ADDIN_DIR, '__pycache__')
        if os.path.isdir(pycache_dir):
            shutil.rmtree(pycache_dir, ignore_errors=True)

        shutil.rmtree(PENDING_UPDATE_DIR, ignore_errors=True)
        return {'latest_version': latest_version or _current_addin_version(), 'error': ''}
    except Exception as exc:
        return {'latest_version': '', 'error': str(exc)}


def _launch_updated_addin_from_disk(context):
    global _updated_runtime_module
    updated_entry_path = os.path.join(ADDIN_DIR, 'BetterExport.py')
    module_name = 'better_export_updated_main'
    spec = importlib.util.spec_from_file_location(module_name, updated_entry_path)
    if not spec or not spec.loader:
        raise RuntimeError('Could not load the updated Better Export entry point.')
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    if not hasattr(module, 'run'):
        raise RuntimeError('The updated Better Export entry point did not define run(context).')
    _updated_runtime_module = module
    module.run(context)


def _normalized_formats(formats_value, legacy_format=None):
    if isinstance(formats_value, list):
        normalized = [value for value in formats_value if value in FORMAT_LABELS]
        if normalized:
            return normalized

    if legacy_format in FORMAT_LABELS:
        return [legacy_format]

    return list(DEFAULT_SETTINGS['formats'])


def _selected_formats_from_inputs(inputs):
    selected = []
    for format_key in FORMAT_LABELS:
        input_id = f'format_{format_key}'
        checkbox = adsk.core.BoolValueCommandInput.cast(inputs.itemById(input_id))
        if checkbox and checkbox.value:
            selected.append(format_key)
    return selected


def _set_format_enabled(inputs, format_key, enabled):
    checkbox = adsk.core.BoolValueCommandInput.cast(inputs.itemById(f'format_{format_key}'))
    if not checkbox:
        return
    checkbox.isEnabled = enabled
    if not enabled:
        checkbox.value = False


def _option_prefix(scope_key):
    return f'{scope_key}_'


def _option_input_id(scope_key, field_name):
    return f'{_option_prefix(scope_key)}{field_name}'


def _group_input_id(scope_key):
    return f'{scope_key}_settings_group'


def _custom_group_input_id(scope_key):
    return f'{scope_key}_custom_group'


def _read_bool_input(inputs, input_id):
    input_obj = adsk.core.BoolValueCommandInput.cast(inputs.itemById(input_id))
    return input_obj.value if input_obj else None


def _read_string_input(inputs, input_id):
    input_obj = adsk.core.StringValueCommandInput.cast(inputs.itemById(input_id))
    return input_obj.value.strip() if input_obj else None


def _settings_mode_from_inputs(inputs):
    raw_value = _dropdown_value(inputs, 'settings_mode')
    for key, label in SETTINGS_MODE_LABELS.items():
        if raw_value == label:
            return key
    return 'global'


def _target_mode_from_inputs(inputs):
    raw_value = _dropdown_value(inputs, 'target_mode')
    for key, label in TARGET_MODE_LABELS.items():
        if raw_value == label:
            return key
    full_root_enabled = _read_bool_input(inputs, 'always_export_full_root')
    return 'full_design' if full_root_enabled else 'selection'


def _destination_mode_from_inputs(inputs):
    raw_value = _dropdown_value(inputs, 'destination_mode')
    for key, label in DESTINATION_MODE_LABELS.items():
        if raw_value == label:
            return key
    auto_sort_enabled = _read_bool_input(inputs, 'auto_sort_after_export')
    return 'sorted' if auto_sort_enabled else 'direct'


def _read_option_values(inputs, scope_key):
    result = {}
    for field_name, default_value in OPTION_DEFAULTS.items():
        input_id = _option_input_id(scope_key, field_name)
        if isinstance(default_value, bool):
            value = _read_bool_input(inputs, input_id)
        elif field_name in ('mesh_refinement', 'unit_type', 'print_utility_mode'):
            value = _selected_key(inputs, input_id)
        else:
            value = _read_string_input(inputs, input_id)

        if value is None:
            return None
        result[field_name] = value
    return result


def _read_general_settings(inputs):
    folder = _read_string_input(inputs, 'folder')
    sorted_output_folder = _read_string_input(inputs, 'sorted_output_folder')
    target_mode = _target_mode_from_inputs(inputs)
    destination_mode = _destination_mode_from_inputs(inputs)
    auto_sort_after_export = destination_mode == 'sorted'
    always_export_full_root = target_mode == 'full_design'
    auto_check_updates = _read_bool_input(inputs, 'auto_check_updates')
    allow_overwrite = _read_bool_input(inputs, 'allow_overwrite')
    customize_per_format = _read_bool_input(inputs, 'customize_per_format')
    f3d_enabled_preference = _read_bool_input(inputs, 'f3d_enabled_preference')

    if None in (
        folder,
        sorted_output_folder,
        auto_check_updates,
        allow_overwrite,
        customize_per_format,
        f3d_enabled_preference
    ):
        return None

    return {
        'folder': folder,
        'sorted_output_folder': sorted_output_folder,
        'auto_sort_after_export': auto_sort_after_export,
        'always_export_full_root': always_export_full_root,
        'target_mode': target_mode,
        'f3d_enabled_preference': f3d_enabled_preference,
        'auto_check_updates': auto_check_updates,
        'allow_overwrite': allow_overwrite,
        'settings_mode': 'per_format' if customize_per_format else 'global'
    }


def _settings_for_format(settings, format_key):
    if settings.get('settings_mode') == 'per_format':
        format_settings = dict(OPTION_DEFAULTS)
        format_settings.update(settings['per_format_settings'].get(format_key, {}))
        return format_settings

    return {key: settings[key] for key in OPTION_DEFAULTS}


def _primary_format(settings):
    formats = _normalized_formats(settings.get('formats'), settings.get('format'))
    return formats[0]


def _selected_entity(inputs):
    selection_input = adsk.core.SelectionCommandInput.cast(inputs.itemById('geometry'))
    if selection_input and selection_input.selectionCount > 0:
        return selection_input.selection(0).entity
    return None


def _selected_geometry(inputs):
    return _selected_entity(inputs) or _root_component()


def _target_geometry(settings, inputs):
    if settings.get('target_mode') == 'selection':
        return _selected_entity(inputs)
    return _root_component()


def _geometry_for_format(format_key, geometry):
    if format_key != 'f3d':
        return geometry

    if not geometry:
        return None

    occurrence = adsk.fusion.Occurrence.cast(geometry)
    if occurrence:
        return occurrence.component

    body = adsk.fusion.BRepBody.cast(geometry)
    if body:
        return body.parentComponent

    component = adsk.fusion.Component.cast(geometry)
    if component:
        return component

    return None


def _component_has_bodies(component, visited=None):
    component = adsk.fusion.Component.cast(component)
    if not component:
        return False

    visited = visited or set()
    token = _safe_call(lambda: component.entityToken) or str(id(component))
    if token in visited:
        return False
    visited.add(token)

    body_count = _safe_call(lambda: component.bRepBodies.count) or 0
    if body_count > 0:
        return True

    occurrence_count = _safe_call(lambda: component.occurrences.count) or 0
    for index in range(occurrence_count):
        occurrence = _safe_call(lambda i=index: component.occurrences.item(i))
        child_component = _safe_call(lambda occ=occurrence: occ.component)
        if child_component and _component_has_bodies(child_component, visited):
            return True

    return False


def _component_has_visible_bodies(component, visited=None):
    component = adsk.fusion.Component.cast(component)
    if not component:
        return False

    visited = visited or set()
    token = _safe_call(lambda: component.entityToken) or str(id(component))
    if token in visited:
        return False
    visited.add(token)

    body_count = _safe_call(lambda: component.bRepBodies.count) or 0
    for index in range(body_count):
        body = _safe_call(lambda i=index: component.bRepBodies.item(i))
        if body and bool(_safe_call(lambda b=body: b.isLightBulbOn)):
            return True

    occurrence_count = _safe_call(lambda: component.occurrences.count) or 0
    for index in range(occurrence_count):
        occurrence = _safe_call(lambda i=index: component.occurrences.item(i))
        if not occurrence or not bool(_safe_call(lambda occ=occurrence: occ.isLightBulbOn)):
            continue
        child_component = _safe_call(lambda occ=occurrence: occ.component)
        if child_component and _component_has_visible_bodies(child_component, visited):
            return True

    return False


def _geometry_is_exportable(entity, target_mode='selection'):
    if adsk.fusion.BRepBody.cast(entity):
        return True

    occurrence = adsk.fusion.Occurrence.cast(entity)
    if occurrence:
        if target_mode == 'visible_bodies':
            return _component_has_visible_bodies(occurrence.component)
        return _component_has_bodies(occurrence.component)

    component = adsk.fusion.Component.cast(entity)
    if component:
        if target_mode == 'visible_bodies':
            return _component_has_visible_bodies(component)
        return _component_has_bodies(component)

    return False


def _body_collections_for_component(component):
    component = adsk.fusion.Component.cast(component)
    if not component:
        return []

    collections = []
    for attribute_name in ('bRepBodies', 'meshBodies'):
        collection = _safe_call(lambda name=attribute_name: getattr(component, name))
        if collection:
            collections.append(collection)
    return collections


def _collect_full_root_state(component, visited=None, state=None):
    component = adsk.fusion.Component.cast(component)
    if not component:
        return state or {'occurrences': [], 'bodies': []}

    visited = visited or set()
    state = state or {'occurrences': [], 'bodies': []}

    token = _safe_call(lambda: component.entityToken) or str(id(component))
    if token in visited:
        return state
    visited.add(token)

    for collection in _body_collections_for_component(component):
        count = _safe_call(lambda c=collection: c.count) or 0
        for index in range(count):
            body = _safe_call(lambda c=collection, i=index: c.item(i))
            if body:
                state['bodies'].append((body, bool(_safe_call(lambda b=body: b.isLightBulbOn))))

    occurrences = _safe_call(lambda: component.occurrences)
    count = _safe_call(lambda o=occurrences: o.count) or 0
    for index in range(count):
        occurrence = _safe_call(lambda o=occurrences, i=index: o.item(i))
        if not occurrence:
            continue
        state['occurrences'].append((
            occurrence,
            bool(_safe_call(lambda occ=occurrence: occ.isLightBulbOn)),
            bool(_safe_call(lambda occ=occurrence: occ.isIsolated))
        ))
        child_component = _safe_call(lambda occ=occurrence: occ.component)
        _collect_full_root_state(child_component, visited, state)

    return state


def _restore_full_root_state(design, saved_state):
    if not saved_state:
        return

    for body, was_visible in reversed(saved_state.get('bodies', [])):
        try:
            body.isLightBulbOn = was_visible
        except Exception:
            pass

    for occurrence, was_visible, was_isolated in reversed(saved_state.get('occurrences', [])):
        try:
            occurrence.isLightBulbOn = was_visible
        except Exception:
            pass
        try:
            occurrence.isIsolated = was_isolated
        except Exception:
            pass

    active_occurrence = saved_state.get('active_occurrence')
    try:
        if active_occurrence:
            active_occurrence.activate()
        elif design:
            design.activateRootComponent()
    except Exception:
        pass


def _prepare_full_root_export(design):
    root_component = _root_component()
    if not design or not root_component:
        return None

    saved_state = _collect_full_root_state(root_component)
    saved_state['active_occurrence'] = _safe_call(lambda: design.activeOccurrence)

    try:
        design.activateRootComponent()
    except Exception:
        pass

    for occurrence, _, _ in saved_state.get('occurrences', []):
        try:
            occurrence.isIsolated = False
        except Exception:
            pass
        try:
            occurrence.isLightBulbOn = True
        except Exception:
            pass

    for body, _ in saved_state.get('bodies', []):
        try:
            body.isLightBulbOn = True
        except Exception:
            pass

    return saved_state


def _prepare_visible_bodies_export(design):
    root_component = _root_component()
    if not design or not root_component:
        return None

    saved_state = _collect_full_root_state(root_component)
    saved_state['active_occurrence'] = _safe_call(lambda: design.activeOccurrence)

    try:
        design.activateRootComponent()
    except Exception:
        pass

    return saved_state


def _mesh_refinement_enum(setting_key):
    enum_type = getattr(adsk.fusion, 'MeshRefinementSettings', None)
    if enum_type:
        mapping = {
            'high': getattr(enum_type, 'MeshRefinementHigh', 0),
            'medium': getattr(enum_type, 'MeshRefinementMedium', 1),
            'low': getattr(enum_type, 'MeshRefinementLow', 2),
            'custom': getattr(enum_type, 'MeshRefinementCustom', 3)
        }
        return mapping[setting_key]

    fallback = {
        'high': 0,
        'medium': 1,
        'low': 2,
        'custom': 3
    }
    return fallback[setting_key]


def _distance_unit_enum(unit_key):
    enum_type = getattr(adsk.fusion, 'DistanceUnits', None)
    if enum_type:
        mapping = {
            'mm': getattr(enum_type, 'MillimeterDistanceUnits', 0),
            'cm': getattr(enum_type, 'CentimeterDistanceUnits', 1),
            'm': getattr(enum_type, 'MeterDistanceUnits', 2),
            'in': getattr(enum_type, 'InchDistanceUnits', 3),
            'ft': getattr(enum_type, 'FootDistanceUnits', 4)
        }
        return mapping.get(unit_key)

    fallback = {
        'mm': 0,
        'cm': 1,
        'm': 2,
        'in': 3,
        'ft': 4
    }
    return fallback.get(unit_key)


def _create_export_options(format_key, geometry, filename=''):
    design = _active_design()
    export_manager = design.exportManager

    if format_key == 'stl':
        return export_manager.createSTLExportOptions(geometry, filename) if filename else export_manager.createSTLExportOptions(geometry)
    if format_key == 'obj':
        return export_manager.createOBJExportOptions(geometry, filename)
    if format_key == '3mf':
        return export_manager.createC3MFExportOptions(geometry, filename)
    if format_key == 'f3d':
        return export_manager.createFusionArchiveExportOptions(filename, geometry)

    raise ValueError(f'Unsupported format: {format_key}')


def _capability_probe_path(format_key):
    return os.path.join(
        os.path.expanduser('~'),
        '__better_mesh_export_probe__.{}'.format(_format_extension(format_key))
    )


def _capabilities_for(format_key, geometry):
    probe_path = _capability_probe_path(format_key) if format_key in ('obj', '3mf', 'f3d') else ''
    try:
        options = _create_export_options(format_key, geometry, probe_path)
    except Exception:
        return {
            'binary_format': format_key == 'stl',
            'mesh_refinement': format_key != 'f3d',
            'surface_deviation': format_key != 'f3d',
            'normal_deviation': format_key != 'f3d',
            'maximum_edge_length': format_key != 'f3d',
            'aspect_ratio': format_key != 'f3d',
            'unit_type': format_key != 'f3d',
            'one_file_per_body': format_key != 'f3d',
            'send_to_print': False,
            'print_utility': False,
            'available_print_utilities': []
        }
    return {
        'binary_format': _supports_attr(options, 'isBinaryFormat'),
        'mesh_refinement': _supports_attr(options, 'meshRefinement'),
        'surface_deviation': _supports_attr(options, 'surfaceDeviation'),
        'normal_deviation': _supports_attr(options, 'normalDeviation'),
        'maximum_edge_length': _supports_attr(options, 'maximumEdgeLength'),
        'aspect_ratio': _supports_attr(options, 'aspectRatio'),
        'unit_type': _supports_attr(options, 'unitType'),
        'one_file_per_body': _supports_attr(options, 'isOneFilePerBody'),
        'send_to_print': _supports_attr(options, 'sendToPrintUtility'),
        'print_utility': _supports_attr(options, 'printUtility'),
        'available_print_utilities': list(_safe_call(lambda: options.availablePrintUtilities) or [])
    }


def _empty_capabilities():
    return {
        'binary_format': False,
        'mesh_refinement': False,
        'surface_deviation': False,
        'normal_deviation': False,
        'maximum_edge_length': False,
        'aspect_ratio': False,
        'unit_type': False,
        'one_file_per_body': False,
        'send_to_print': False,
        'print_utility': False,
        'available_print_utilities': []
    }


def _combined_capabilities(format_keys, geometry):
    combined = _empty_capabilities()

    seen_utilities = set()
    for format_key in format_keys:
        capabilities = _capabilities_for(format_key, geometry)
        for key in (
            'binary_format',
            'mesh_refinement',
            'surface_deviation',
            'normal_deviation',
            'maximum_edge_length',
            'aspect_ratio',
            'unit_type',
            'one_file_per_body',
            'send_to_print',
            'print_utility'
        ):
            combined[key] = combined[key] or capabilities[key]

        for utility_name in capabilities['available_print_utilities']:
            if utility_name not in seen_utilities:
                seen_utilities.add(utility_name)
                combined['available_print_utilities'].append(utility_name)

    return combined


def _dropdown_value(inputs, input_id):
    dropdown = adsk.core.DropDownCommandInput.cast(inputs.itemById(input_id))
    selected_item = dropdown.selectedItem if dropdown else None
    return selected_item.name if selected_item else ''


def _selected_key(inputs, input_id):
    raw_value = _dropdown_value(inputs, input_id)

    if input_id.endswith('mesh_refinement'):
        return MESH_REFINEMENT_KEYS_BY_LABEL.get(raw_value, 'medium')
    if input_id.endswith('unit_type'):
        return UNIT_KEYS_BY_LABEL.get(raw_value, 'default')
    if input_id.endswith('print_utility_mode'):
        if raw_value == 'Fusion Default':
            return 'default'
        if raw_value == 'Custom Path Or Name':
            return 'custom'
    return raw_value


def _parse_positive_float(text, label):
    try:
        value = float(text)
    except Exception as exc:
        raise ValueError(f'{label} must be a number.') from exc

    if value <= 0:
        raise ValueError(f'{label} must be greater than zero.')

    return value


def _current_settings_from_inputs(inputs):
    general_settings = _read_general_settings(inputs)
    global_values = _read_option_values(inputs, 'global')
    if global_values is None or general_settings is None:
        return dict(_load_settings())
    per_format_settings = {}
    for format_key in FORMAT_LABELS:
        option_values = _read_option_values(inputs, format_key)
        if option_values is None:
            return dict(_load_settings())
        per_format_settings[format_key] = option_values

    return {
        **general_settings,
        **global_values,
        'formats': _selected_formats_from_inputs(inputs),
        'settings_mode': general_settings['settings_mode'],
        'per_format_settings': per_format_settings
    }


def _sync_option_scope_ui(command_inputs, scope_key, option_values, capabilities, group_visible, auto_sort_enabled):
    group_input = adsk.core.GroupCommandInput.cast(command_inputs.itemById(_group_input_id(scope_key)))
    refinement_input = adsk.core.DropDownCommandInput.cast(command_inputs.itemById(_option_input_id(scope_key, 'mesh_refinement')))
    binary_input = adsk.core.BoolValueCommandInput.cast(command_inputs.itemById(_option_input_id(scope_key, 'binary_format')))
    one_per_body_input = adsk.core.BoolValueCommandInput.cast(command_inputs.itemById(_option_input_id(scope_key, 'one_file_per_body')))
    send_to_print_input = adsk.core.BoolValueCommandInput.cast(command_inputs.itemById(_option_input_id(scope_key, 'send_to_print_utility')))
    unit_input = adsk.core.DropDownCommandInput.cast(command_inputs.itemById(_option_input_id(scope_key, 'unit_type')))
    custom_group = adsk.core.GroupCommandInput.cast(command_inputs.itemById(_custom_group_input_id(scope_key)))
    print_selector = adsk.core.DropDownCommandInput.cast(command_inputs.itemById(_option_input_id(scope_key, 'print_utility_mode')))
    print_value = adsk.core.StringValueCommandInput.cast(command_inputs.itemById(_option_input_id(scope_key, 'print_utility_value')))

    if (
        not group_input or
        not refinement_input or
        not binary_input or
        not one_per_body_input or
        not send_to_print_input or
        not unit_input or
        not custom_group or
        not print_selector or
        not print_value
    ):
        return

    group_input.isVisible = group_visible
    if not group_visible:
        return

    refinement_input.isVisible = capabilities['mesh_refinement']
    binary_input.isVisible = capabilities['binary_format']
    one_per_body_input.isVisible = capabilities['one_file_per_body']
    unit_input.isVisible = capabilities['unit_type']
    send_to_print_input.isVisible = capabilities['send_to_print'] and not auto_sort_enabled

    if not capabilities['send_to_print'] or auto_sort_enabled:
        send_to_print_input.value = False

    refinement_visible = capabilities['surface_deviation'] or capabilities['normal_deviation'] or capabilities['maximum_edge_length'] or capabilities['aspect_ratio']
    custom_group.isVisible = refinement_visible and option_values['mesh_refinement'] == 'custom'

    utilities = capabilities['available_print_utilities']
    print_selector.isVisible = capabilities['send_to_print'] and not auto_sort_enabled and send_to_print_input.value and capabilities['print_utility']
    print_value.isVisible = capabilities['send_to_print'] and not auto_sort_enabled and send_to_print_input.value and capabilities['print_utility']

    if print_selector.isVisible:
        list_items = print_selector.listItems
        list_items.clear()
        list_items.add('Fusion Default', option_values['print_utility_mode'] == 'default', '')
        for utility_name in utilities:
            list_items.add(utility_name, option_values['print_utility_mode'] == utility_name, '')
        list_items.add('Custom Path Or Name', option_values['print_utility_mode'] == 'custom', '')

        if not print_selector.selectedItem:
            list_items.item(0).isSelected = True

        print_value.isVisible = print_selector.selectedItem.name == 'Custom Path Or Name'
        if print_selector.selectedItem.name != 'Custom Path Or Name':
            print_value.tooltip = 'Using Fusion or utility default.'
        else:
            print_value.tooltip = 'Enter a print utility executable path or a known utility name.'
    else:
        print_value.isVisible = False


def _sync_ui(command_inputs):
    settings = _merge_settings(_current_settings_from_inputs(command_inputs))
    geometry = _target_geometry(settings, command_inputs) or _root_component()
    if not geometry:
        return

    target_mode_input = adsk.core.DropDownCommandInput.cast(command_inputs.itemById('target_mode'))
    target_input = adsk.core.SelectionCommandInput.cast(command_inputs.itemById('geometry'))
    target_hint = adsk.core.TextBoxCommandInput.cast(command_inputs.itemById('target_hint'))
    destination_mode_input = adsk.core.DropDownCommandInput.cast(command_inputs.itemById('destination_mode'))
    format_note = adsk.core.TextBoxCommandInput.cast(command_inputs.itemById('format_note'))
    f3d_pref_input = adsk.core.BoolValueCommandInput.cast(command_inputs.itemById('f3d_enabled_preference'))
    f3d_checkbox = adsk.core.BoolValueCommandInput.cast(command_inputs.itemById('format_f3d'))
    last_target_mode_input = adsk.core.StringValueCommandInput.cast(command_inputs.itemById('last_target_mode'))
    folder_input = adsk.core.StringValueCommandInput.cast(command_inputs.itemById('folder'))
    folder_summary = adsk.core.TextBoxCommandInput.cast(command_inputs.itemById('folder_summary'))
    browse_folder_button = adsk.core.BoolValueCommandInput.cast(command_inputs.itemById('browse_folder'))
    sorted_output_input = adsk.core.StringValueCommandInput.cast(command_inputs.itemById('sorted_output_folder'))
    sorted_output_summary = adsk.core.TextBoxCommandInput.cast(command_inputs.itemById('sorted_output_folder_summary'))
    browse_sorted_output_button = adsk.core.BoolValueCommandInput.cast(command_inputs.itemById('browse_sorted_output_folder'))
    overwrite_input = adsk.core.BoolValueCommandInput.cast(command_inputs.itemById('allow_overwrite'))
    customize_per_format_input = adsk.core.BoolValueCommandInput.cast(command_inputs.itemById('customize_per_format'))

    if (
        not target_mode_input or
        not target_input or
        not target_hint or
        not destination_mode_input or
        not format_note or
        not f3d_pref_input or
        not f3d_checkbox or
        not last_target_mode_input or
        not folder_input or
        not folder_summary or
        not browse_folder_button or
        not sorted_output_input or
        not sorted_output_summary or
        not browse_sorted_output_button or
        not overwrite_input or
        not customize_per_format_input
    ):
        return

    target_mode = settings.get('target_mode', 'selection')
    target_input.isVisible = target_mode == 'selection'
    if target_mode == 'full_design':
        target_hint.formattedText = 'Exports the full design from the root component after temporarily showing everything.'
    elif target_mode == 'visible_bodies':
        target_hint.formattedText = 'Exports only bodies that are currently visible in the design.'
    else:
        target_hint.formattedText = 'Select a body, component, or occurrence to export.'

    visible_bodies_mode = target_mode == 'visible_bodies'
    previous_target_mode = (last_target_mode_input.value or '').strip() or target_mode
    if visible_bodies_mode:
        if previous_target_mode != 'visible_bodies':
            f3d_pref_input.value = bool(f3d_checkbox.value)
        f3d_checkbox.value = False
        f3d_checkbox.isEnabled = False
    else:
        f3d_checkbox.isEnabled = True
        if previous_target_mode == 'visible_bodies':
            f3d_checkbox.value = bool(f3d_pref_input.value)
    format_note.isVisible = visible_bodies_mode
    format_note.formattedText = 'F3D is unavailable in Export Only Visible Bodies mode because Fusion archive export works at the component level.'
    last_target_mode_input.value = target_mode

    folder_input.isVisible = False
    folder_summary.isVisible = not settings['auto_sort_after_export']
    browse_folder_button.isVisible = not settings['auto_sort_after_export']
    sorted_output_input.isVisible = False
    sorted_output_summary.isVisible = settings['auto_sort_after_export']
    browse_sorted_output_button.isVisible = settings['auto_sort_after_export']
    overwrite_input.isVisible = settings['auto_sort_after_export']
    customize_per_format_input.tooltip = 'Reveal separate STL, OBJ, 3MF, and F3D settings sections.'

    folder_summary.formattedText = _short_path(settings['folder'])
    folder_summary.tooltip = settings['folder']
    sorted_output_summary.formattedText = _short_path(settings['sorted_output_folder'])
    sorted_output_summary.tooltip = settings['sorted_output_folder']

    settings_mode = settings['settings_mode']
    global_capabilities = _combined_capabilities(settings['formats'], geometry)
    global_values = _settings_for_format(settings, _primary_format(settings))
    _sync_option_scope_ui(command_inputs, 'global', global_values, global_capabilities, settings_mode == 'global', settings['auto_sort_after_export'])

    for format_key in FORMAT_LABELS:
        format_geometry = _geometry_for_format(format_key, geometry)
        capabilities = _capabilities_for(format_key, format_geometry) if format_geometry else _empty_capabilities()
        option_values = settings['per_format_settings'][format_key]
        group_visible = settings_mode == 'per_format' and format_key in settings['formats']
        _sync_option_scope_ui(command_inputs, format_key, option_values, capabilities, group_visible, settings['auto_sort_after_export'])


def _refresh_update_ui(command_inputs, force_refresh=False, manual=False):
    status_input = adsk.core.TextBoxCommandInput.cast(command_inputs.itemById('update_status'))
    auto_check_input = adsk.core.BoolValueCommandInput.cast(command_inputs.itemById('auto_check_updates'))
    update_now_input = adsk.core.BoolValueCommandInput.cast(command_inputs.itemById('update_now'))

    if not status_input or not auto_check_input or not update_now_input:
        return

    current_version = _current_addin_version()
    pending_version = _pending_update_version()
    auto_check_enabled = bool(auto_check_input.value)

    if pending_version:
        status_input.isVisible = True
        status_input.formattedText = 'Version v{} - Restart pending for v{}'.format(current_version, pending_version)
        status_input.tooltip = ''
        update_now_input.isVisible = False
        return

    if not auto_check_enabled and not manual:
        status_input.isVisible = True
        status_input.formattedText = 'Version v{}'.format(current_version)
        status_input.tooltip = ''
        update_now_input.isVisible = False
        return

    release_info = _latest_release_info(force_refresh=force_refresh, allow_cached_on_error=not manual)
    latest_version = release_info.get('latest_version', '')
    latest_url = release_info.get('latest_url') or LATEST_RELEASE_PAGE_URL
    has_update = bool(latest_version and _is_version_newer(latest_version, current_version))

    if has_update:
        status_input.isVisible = True
        status_input.formattedText = 'Version v{} - <a href="{}">Update available: v{}</a>'.format(current_version, latest_url, latest_version)
        status_input.tooltip = latest_url
        update_now_input.isVisible = True
        return

    if manual or not auto_check_enabled:
        status_input.isVisible = True
        if release_info.get('error'):
            status_input.formattedText = 'Version v{} - Unable to check for updates right now'.format(current_version)
            status_input.tooltip = str(release_info.get('error', ''))
        else:
            status_input.formattedText = 'Version v{} - Up to date'.format(current_version)
            status_input.tooltip = ''
        update_now_input.isVisible = False
    else:
        status_input.isVisible = True
        status_input.formattedText = 'Version v{}'.format(current_version)
        status_input.tooltip = ''
        update_now_input.isVisible = False


def _show_error(message):
    if _ui:
        _ui.messageBox(message, COMMAND_NAME)


def _choose_sort_conflict_action(conflicts):
    preview_lines = []
    for conflict in conflicts[:3]:
        preview_lines.append(
            "Incoming: {}\nExisting: {}\nLocation: {}".format(
                conflict["incoming_name"],
                conflict["existing_name"],
                conflict["target_path"]
            )
        )

    extra_count = max(0, len(conflicts) - len(preview_lines))
    extra_text = "\n\nAnd {} more conflict(s).".format(extra_count) if extra_count else ""
    message = (
        '{} sorted export conflict(s) were found.\n\n'
        '{}{}'
        '\n\nChoose an action for all conflicts:\n'
        'Yes: Overwrite existing files\n'
        'No: Keep both files and save the new one with a unique name if needed\n'
        'Cancel: Keep the existing files and discard all new conflicting files'
    ).format(len(conflicts), "\n\n".join(preview_lines), extra_text)

    if not _ui:
        return "skip"

    result = _ui.messageBox(
        message,
        'Overwrite conflicting files?',
        adsk.core.MessageBoxButtonTypes.YesNoCancelButtonType,
        adsk.core.MessageBoxIconTypes.WarningIconType
    )
    if result == adsk.core.DialogResults.DialogYes:
        return "overwrite"
    if result == adsk.core.DialogResults.DialogNo:
        return "keep_both"
    return "skip"


def _choose_single_sort_conflict_action(source, target, operation, keep_both_target):
    keep_both_name = keep_both_target.name if keep_both_target else source.name
    message = (
        'A sorted export file already exists at this location. Overwrite it?\n\n'
        'Incoming file:\n{}\n\n'
        'Existing file:\n{}\n\n'
        'Location:\n{}\n\n'
        'Choose an action:\n'
        'Yes: Overwrite the existing file\n'
        'No: Keep both files and save the new one as {}\n'
        'Cancel: Keep the existing file and discard the new conflicting file'
    ).format(
        source.name,
        target.name,
        str(target),
        keep_both_name
    )

    if not _ui:
        return "skip"

    result = _ui.messageBox(
        message,
        'Overwrite conflicting file?',
        adsk.core.MessageBoxButtonTypes.YesNoCancelButtonType,
        adsk.core.MessageBoxIconTypes.WarningIconType
    )
    if result == adsk.core.DialogResults.DialogYes:
        return "overwrite"
    if result == adsk.core.DialogResults.DialogNo:
        return "keep_both"
    return "skip"


def _persist_current_preferences(inputs):
    settings = _current_settings_from_inputs(inputs)
    if settings:
        _save_settings(settings)


def _apply_options_from_settings(options, settings):
    if not _supports_attr(options, 'meshRefinement'):
        return

    if _supports_attr(options, 'meshRefinement'):
        if settings['mesh_refinement'] == 'custom':
            options.meshRefinement = _mesh_refinement_enum('custom')
        else:
            options.meshRefinement = _mesh_refinement_enum(settings['mesh_refinement'])

    if settings['mesh_refinement'] == 'custom':
        if _supports_attr(options, 'surfaceDeviation'):
            options.surfaceDeviation = _parse_positive_float(settings['surface_deviation_cm'], 'Surface deviation')
        if _supports_attr(options, 'normalDeviation'):
            options.normalDeviation = _parse_positive_float(settings['normal_deviation_rad'], 'Normal deviation')
        if _supports_attr(options, 'maximumEdgeLength'):
            options.maximumEdgeLength = _parse_positive_float(settings['maximum_edge_length_cm'], 'Maximum edge length')
        if _supports_attr(options, 'aspectRatio'):
            options.aspectRatio = _parse_positive_float(settings['aspect_ratio'], 'Aspect ratio')

    if _supports_attr(options, 'unitType') and settings['unit_type'] != 'default':
        unit_enum = _distance_unit_enum(settings['unit_type'])
        if unit_enum is not None:
            options.unitType = unit_enum

    if _supports_attr(options, 'isBinaryFormat'):
        options.isBinaryFormat = bool(settings['binary_format'])

    if _supports_attr(options, 'isOneFilePerBody'):
        options.isOneFilePerBody = bool(settings['one_file_per_body'])

    if _supports_attr(options, 'sendToPrintUtility'):
        options.sendToPrintUtility = bool(settings['send_to_print_utility'])

    if _supports_attr(options, 'printUtility') and settings['send_to_print_utility']:
        if settings['print_utility_mode'] == 'custom':
            if not settings['print_utility_value']:
                raise ValueError('Enter a print utility path or switch the print utility mode away from custom.')
            options.printUtility = settings['print_utility_value']
        elif settings['print_utility_mode'] not in ('', 'default'):
            options.printUtility = settings['print_utility_mode']


def _validate_inputs(command_inputs):
    design = _active_design()
    if not design:
        return False, 'Open a Fusion design before using this command.'

    settings = _current_settings_from_inputs(command_inputs)
    target_mode = settings.get('target_mode', 'selection')
    geometry = _target_geometry(settings, command_inputs)
    if not geometry:
        if target_mode == 'selection':
            return False, 'Select a body, component, or occurrence to export.'
        return False, 'Open a Fusion design before exporting the full design.'
    if not _geometry_is_exportable(geometry, target_mode):
        if target_mode == 'full_design':
            return False, 'Nothing exportable was found in the active root design.'
        if target_mode == 'visible_bodies':
            return False, 'No visible bodies were found in the active design.'
        return False, 'Nothing exportable was found in the current selection or active design.'

    if not settings['formats']:
        return False, 'Select at least one export format.'

    for format_key in settings['formats']:
        format_geometry = _geometry_for_format(format_key, geometry)
        if not format_geometry or not _geometry_is_exportable(format_geometry, target_mode):
            if format_key == 'f3d':
                if target_mode == 'full_design':
                    return False, 'F3D export could not find exportable geometry in the active root design.'
                return False, 'F3D export requires a component, occurrence, or body selection with actual model geometry, or an active root component with bodies.'
            return False, 'Nothing exportable was found for {} in the current selection or active design.'.format(FORMAT_LABELS[format_key])

        format_settings = _settings_for_format(settings, format_key)

        if not format_settings['send_to_print_utility']:
            effective_folder = settings['folder']
            if not effective_folder:
                return False, 'Choose an export folder for {}.'.format(FORMAT_LABELS[format_key])
            if not format_settings['filename']:
                return False, 'Enter a file name for {}.'.format(FORMAT_LABELS[format_key])

        if format_key != 'f3d' and format_settings['mesh_refinement'] == 'custom':
            try:
                _parse_positive_float(format_settings['surface_deviation_cm'], 'Surface deviation')
                _parse_positive_float(format_settings['normal_deviation_rad'], 'Normal deviation')
                _parse_positive_float(format_settings['maximum_edge_length_cm'], 'Maximum edge length')
                _parse_positive_float(format_settings['aspect_ratio'], 'Aspect ratio')
            except ValueError as exc:
                return False, '{} for {}.'.format(str(exc).rstrip('.'), FORMAT_LABELS[format_key])

        if format_settings['send_to_print_utility'] and format_settings['print_utility_mode'] == 'custom' and not format_settings['print_utility_value']:
            return False, 'Enter a print utility path for {} or switch away from the custom print utility mode.'.format(FORMAT_LABELS[format_key])

        if settings['auto_sort_after_export'] and format_settings['send_to_print_utility']:
            return False, 'Disable Send To Print Utility for {} when automatic sorting is enabled.'.format(FORMAT_LABELS[format_key])

    if settings['auto_sort_after_export']:
        if not settings['sorted_output_folder']:
            return False, 'Choose a sorted projects folder.'

    return True, ''


def _add_option_inputs(container, scope_key, option_values, label):
    group = container.addGroupCommandInput(_group_input_id(scope_key), label)
    group.isExpanded = scope_key == 'global'
    children = group.children

    filename_input = children.addStringValueInput(_option_input_id(scope_key, 'filename'), 'File Name', option_values['filename'] or _default_filename())
    filename_input.tooltip = 'File name without the extension.'

    refinement_input = children.addDropDownCommandInput(
        _option_input_id(scope_key, 'mesh_refinement'),
        'Refinement',
        adsk.core.DropDownStyles.TextListDropDownStyle
    )
    for key, label_text in MESH_REFINEMENT_LABELS.items():
        refinement_input.listItems.add(label_text, key == option_values['mesh_refinement'], '')

    unit_input = children.addDropDownCommandInput(
        _option_input_id(scope_key, 'unit_type'),
        'Units',
        adsk.core.DropDownStyles.TextListDropDownStyle
    )
    for key, label_text in UNIT_LABELS.items():
        unit_input.listItems.add(label_text, key == option_values['unit_type'], '')

    binary_input = children.addBoolValueInput(
        _option_input_id(scope_key, 'binary_format'),
        'Binary STL',
        True,
        '',
        bool(option_values['binary_format'])
    )
    binary_input.tooltip = 'For STL exports only.'

    one_per_body_input = children.addBoolValueInput(
        _option_input_id(scope_key, 'one_file_per_body'),
        'One File Per Body',
        True,
        '',
        bool(option_values['one_file_per_body'])
    )
    one_per_body_input.tooltip = 'When exporting a component or occurrence, create separate files for each body.'

    children.addBoolValueInput(
        _option_input_id(scope_key, 'send_to_print_utility'),
        'Send To Print Utility',
        True,
        '',
        bool(option_values['send_to_print_utility'])
    )

    print_selector = children.addDropDownCommandInput(
        _option_input_id(scope_key, 'print_utility_mode'),
        'Print Utility',
        adsk.core.DropDownStyles.TextListDropDownStyle
    )
    print_selector.listItems.add('Fusion Default', option_values['print_utility_mode'] == 'default', '')
    print_selector.listItems.add('Custom Path Or Name', option_values['print_utility_mode'] == 'custom', '')

    children.addStringValueInput(
        _option_input_id(scope_key, 'print_utility_value'),
        'Custom Utility',
        option_values['print_utility_value']
    )

    custom_group = children.addGroupCommandInput(_custom_group_input_id(scope_key), 'Custom Refinement')
    custom_inputs = custom_group.children
    custom_inputs.addStringValueInput(_option_input_id(scope_key, 'surface_deviation_cm'), 'Surface Deviation (cm)', option_values['surface_deviation_cm'])
    custom_inputs.addStringValueInput(_option_input_id(scope_key, 'normal_deviation_rad'), 'Normal Deviation (rad)', option_values['normal_deviation_rad'])
    custom_inputs.addStringValueInput(_option_input_id(scope_key, 'maximum_edge_length_cm'), 'Maximum Edge Length (cm)', option_values['maximum_edge_length_cm'])
    custom_inputs.addStringValueInput(_option_input_id(scope_key, 'aspect_ratio'), 'Aspect Ratio', option_values['aspect_ratio'])


class CommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def notify(self, args):
        try:
            settings = _upgrade_settings_file()
            current_project_key = _current_project_key()
            needs_project_seed = bool(
                current_project_key and (
                    current_project_key not in settings.get('project_export_folders', {}) or
                    current_project_key not in settings.get('project_auto_sort_preferences', {})
                )
            )
            if needs_project_seed:
                _save_settings(settings)
                settings = _load_settings()
            cmd = args.command
            cmd.okButtonText = 'Export'
            cmd.setDialogInitialSize(520, 660)

            inputs = cmd.commandInputs

            target_mode_input = inputs.addDropDownCommandInput(
                'target_mode',
                'Target',
                adsk.core.DropDownStyles.TextListDropDownStyle
            )
            target_mode_input.listItems.add(
                TARGET_MODE_LABELS['full_design'],
                settings.get('target_mode', 'selection') == 'full_design',
                ''
            )
            target_mode_input.listItems.add(
                TARGET_MODE_LABELS['visible_bodies'],
                settings.get('target_mode', 'selection') == 'visible_bodies',
                ''
            )
            target_mode_input.listItems.add(
                TARGET_MODE_LABELS['selection'],
                settings.get('target_mode', 'selection') == 'selection',
                ''
            )

            target_input = inputs.addSelectionInput('geometry', 'Selection', 'Select a body, component, or occurrence to export.')
            target_input.addSelectionFilter('Bodies')
            target_input.addSelectionFilter('Occurrences')
            target_input.addSelectionFilter('RootComponents')
            target_input.setSelectionLimits(0, 1)
            target_input.tooltip = 'Leave this empty to export the active root component. F3D exports use the resolved component from the selection.'

            inputs.addTextBoxCommandInput(
                'target_hint',
                '',
                'Select a body, component, or occurrence to export.',
                2,
                True
            )
            target_hint_input = adsk.core.TextBoxCommandInput.cast(inputs.itemById('target_hint'))
            target_hint_input.isFullWidth = True
            last_target_mode_input = inputs.addStringValueInput('last_target_mode', 'Last Target Mode', settings.get('target_mode', 'selection'))
            last_target_mode_input.isVisible = False

            destination_mode_input = inputs.addDropDownCommandInput(
                'destination_mode',
                'Destination',
                adsk.core.DropDownStyles.TextListDropDownStyle
            )
            destination_mode_input.listItems.add(
                DESTINATION_MODE_LABELS['direct'],
                not settings['auto_sort_after_export'],
                ''
            )
            destination_mode_input.listItems.add(
                DESTINATION_MODE_LABELS['sorted'],
                bool(settings['auto_sort_after_export']),
                ''
            )

            auto_sort_input = inputs.addBoolValueInput(
                'auto_sort_after_export',
                'Sort Automatically After Export',
                True,
                '',
                bool(settings['auto_sort_after_export'])
            )
            auto_sort_input.isVisible = False

            full_root_input = inputs.addBoolValueInput(
                'always_export_full_root',
                'Always Export Full Design',
                True,
                '',
                bool(settings['always_export_full_root'])
            )
            full_root_input.isVisible = False

            folder_input = inputs.addStringValueInput('folder', 'Export Folder', settings['folder'])
            folder_input.isVisible = False
            folder_input.tooltip = 'Used when automatic sorting is disabled.'
            folder_summary = inputs.addTextBoxCommandInput(
                'folder_summary',
                'Export Folder',
                _short_path(settings['folder']),
                1,
                True
            )
            folder_summary.tooltip = settings['folder']
            browse_button = inputs.addBoolValueInput('browse_folder', 'Browse Export Folder…', False, '', False)
            browse_button.tooltip = 'Choose an export folder.'

            sorted_output_input = inputs.addStringValueInput('sorted_output_folder', 'Sorted Projects Folder', settings['sorted_output_folder'])
            sorted_output_input.isVisible = False
            sorted_output_input.tooltip = 'Sorted project folders are created here.'
            sorted_output_summary = inputs.addTextBoxCommandInput(
                'sorted_output_folder_summary',
                'Sorted Projects Folder',
                _short_path(settings['sorted_output_folder']),
                1,
                True
            )
            sorted_output_summary.tooltip = settings['sorted_output_folder']
            browse_sorted_output_button = inputs.addBoolValueInput('browse_sorted_output_folder', 'Browse Sorted Output…', False, '', False)
            browse_sorted_output_button.tooltip = 'Choose the sorted projects folder.'
            inputs.addBoolValueInput(
                'allow_overwrite',
                'Allow Overwrite',
                True,
                '',
                bool(settings['allow_overwrite'])
            )

            format_group = inputs.addGroupCommandInput('format_group', 'Formats')
            format_inputs = format_group.children
            selected_formats = settings['formats']
            for key, label in FORMAT_LABELS.items():
                format_inputs.addBoolValueInput(
                    f'format_{key}',
                    label,
                    True,
                    '',
                    key in selected_formats
                )
                if key == 'f3d':
                    f3d_pref_input = inputs.addBoolValueInput(
                        'f3d_enabled_preference',
                        'Remember F3D Preference',
                        True,
                        '',
                        bool(settings.get('f3d_enabled_preference', 'f3d' in selected_formats))
                    )
                    f3d_pref_input.isVisible = False
            format_note = format_inputs.addTextBoxCommandInput('format_note', '', '', 2, True)
            format_note.isFullWidth = True
            format_note.isVisible = False

            customize_per_format_input = inputs.addBoolValueInput(
                'customize_per_format',
                'Customize Settings Per Format',
                True,
                '',
                settings['settings_mode'] == 'per_format'
            )
            customize_per_format_input.tooltip = 'Reveal separate STL, OBJ, 3MF, and F3D settings sections.'

            settings_mode_input = inputs.addDropDownCommandInput('settings_mode', 'Settings Scope', adsk.core.DropDownStyles.TextListDropDownStyle)
            for key, label in SETTINGS_MODE_LABELS.items():
                settings_mode_input.listItems.add(label, key == settings['settings_mode'], '')
            settings_mode_input.isVisible = False

            global_option_values = _settings_for_format(settings, _primary_format(settings))
            global_option_values['filename'] = _default_filename()
            _add_option_inputs(inputs, 'global', global_option_values, 'Shared Export Settings')
            for format_key, label in FORMAT_LABELS.items():
                option_values = dict(settings['per_format_settings'][format_key])
                option_values['filename'] = _default_filename()
                _add_option_inputs(inputs, format_key, option_values, '{} Settings'.format(label))

            update_status = inputs.addTextBoxCommandInput('update_status', '', '', 1, True)
            update_status.isFullWidth = True
            update_status.isVisible = True
            inputs.addBoolValueInput(
                'auto_check_updates',
                'Check For Updates Automatically',
                True,
                '',
                bool(settings['auto_check_updates'])
            )
            inputs.addBoolValueInput('check_updates_now', 'Check For Updates', False, '', False)
            update_now_input = inputs.addBoolValueInput('update_now', 'Update Now', False, '', False)
            update_now_input.isVisible = False

            _sync_ui(inputs)
            _refresh_update_ui(inputs, force_refresh=False, manual=False)

            on_execute = ExecuteHandler()
            cmd.execute.add(on_execute)
            _handlers.append(on_execute)

            on_input_changed = InputChangedHandler()
            cmd.inputChanged.add(on_input_changed)
            _handlers.append(on_input_changed)

            on_validate = ValidateHandler()
            cmd.validateInputs.add(on_validate)
            _handlers.append(on_validate)

            on_destroy = DestroyHandler()
            cmd.destroy.add(on_destroy)
            _handlers.append(on_destroy)
        except Exception:
            if _ui:
                _ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


class InputChangedHandler(adsk.core.InputChangedEventHandler):
    def notify(self, args):
        try:
            changed_input = args.input
            inputs = args.inputs

            if changed_input.id in ('browse_folder', 'browse_sorted_output_folder'):
                button = adsk.core.BoolValueCommandInput.cast(changed_input)
                dialog = _ui.createFolderDialog()
                dialog.title = 'Choose {}'.format(
                    'Projects Root Folder' if changed_input.id == 'browse_sorted_output_folder' else 'Export Folder'
                )
                result = dialog.showDialog()
                if result == adsk.core.DialogResults.DialogOK:
                    target_input_map = {
                        'browse_folder': 'folder',
                        'browse_sorted_output_folder': 'sorted_output_folder'
                    }
                    folder_input = adsk.core.StringValueCommandInput.cast(inputs.itemById(target_input_map[changed_input.id]))
                    folder_input.value = dialog.folder
                button.value = False
            elif changed_input.id == 'check_updates_now':
                button = adsk.core.BoolValueCommandInput.cast(changed_input)
                _refresh_update_ui(inputs, force_refresh=True, manual=True)
                button.value = False
            elif changed_input.id == 'update_now':
                button = adsk.core.BoolValueCommandInput.cast(changed_input)
                if button:
                    button.value = False
                release_info = _latest_release_info(force_refresh=True, allow_cached_on_error=False)
                latest_version = release_info.get('latest_version', '')
                current_version = _current_addin_version()
                if not latest_version or not _is_version_newer(latest_version, current_version):
                    _show_error('No newer release is available right now.')
                else:
                    _set_run_on_startup(True)
                    _stage_update_payload(release_info)
                    _ui.messageBox(
                        'Better Export v{} has been downloaded and staged.\n\n'
                        'Run on Startup has been enabled for Better Export so Fusion can finish the update on next launch.\n\n'
                        'You can finish your exports in the current session.\n\n'
                        'Restart Fusion when convenient to apply the update.'.format(
                            latest_version
                        ),
                        COMMAND_NAME
                    )
                    _refresh_update_ui(inputs, force_refresh=False, manual=False)
                return
            elif changed_input.id == 'format_f3d':
                f3d_checkbox = adsk.core.BoolValueCommandInput.cast(changed_input)
                f3d_pref_input = adsk.core.BoolValueCommandInput.cast(inputs.itemById('f3d_enabled_preference'))
                if f3d_checkbox and f3d_pref_input:
                    f3d_pref_input.value = bool(f3d_checkbox.value)

            if changed_input.id.endswith('filename'):
                filename_input = adsk.core.StringValueCommandInput.cast(changed_input)
                filename_input.value = _sanitize_filename(filename_input.value)

            _sync_ui(inputs)
            _persist_current_preferences(inputs)
        except Exception:
            _show_error(traceback.format_exc())


class ValidateHandler(adsk.core.ValidateInputsEventHandler):
    def notify(self, args):
        try:
            valid, message = _validate_inputs(args.inputs)
            args.areInputsValid = valid
            target_hint = args.inputs.itemById('target_hint')
            if not valid and message:
                target_hint.formattedText = message
            elif _target_mode_from_inputs(args.inputs) == 'full_design':
                target_hint.formattedText = 'Exports the full design from the root component after temporarily showing everything.'
            elif _target_mode_from_inputs(args.inputs) == 'visible_bodies':
                target_hint.formattedText = 'Exports only bodies that are currently visible in the design.'
            else:
                target_hint.formattedText = 'Select a body, component, or occurrence to export.'
        except Exception:
            args.areInputsValid = False


class ExecuteHandler(adsk.core.CommandEventHandler):
    def notify(self, args):
        progress_dialog = None
        export_succeeded = False
        export_cancelled = False
        temp_staging_dir = None
        full_root_state = None
        design = None
        try:
            inputs = args.command.commandInputs
            valid, message = _validate_inputs(inputs)
            if not valid:
                raise ValueError(message)

            settings = _current_settings_from_inputs(inputs)
            design = _active_design()
            target_mode = settings.get('target_mode', 'selection')
            geometry = _target_geometry(settings, inputs)
            if target_mode == 'full_design':
                full_root_state = _prepare_full_root_export(design)
            elif target_mode == 'visible_bodies':
                full_root_state = _prepare_visible_bodies_export(design)

            total_exports = len(settings['formats'])
            progress_dialog = _ui.createProgressDialog() if _ui else None
            if progress_dialog:
                progress_dialog.cancelButtonText = ''
                progress_dialog.isCancelButtonShown = False
                progress_dialog.show(
                    'Better Export',
                    'Preparing exports...',
                    0,
                    max(1, total_exports),
                    0
                )

            for index, format_key in enumerate(settings['formats'], start=1):
                format_settings = _settings_for_format(settings, format_key)
                format_settings['filename'] = _sanitize_filename(format_settings['filename'] or _default_filename())
                if settings['settings_mode'] == 'global':
                    settings['filename'] = format_settings['filename']
                else:
                    settings['per_format_settings'][format_key]['filename'] = format_settings['filename']

                export_path = ''
                if not format_settings['send_to_print_utility']:
                    export_folder = settings['folder']
                    if settings['auto_sort_after_export']:
                        if temp_staging_dir is None:
                            temp_staging_dir = tempfile.mkdtemp(prefix='better-export-')
                        export_folder = temp_staging_dir
                    os.makedirs(export_folder, exist_ok=True)
                    export_path = os.path.join(
                        export_folder,
                        '{}.{}'.format(format_settings['filename'], _format_extension(format_key))
                    )

                format_geometry = _geometry_for_format(format_key, geometry)
                if not format_geometry:
                    raise ValueError('Fusion could not resolve valid geometry for the {} export.'.format(FORMAT_LABELS[format_key]))

                if progress_dialog:
                    progress_dialog.progressValue = index - 1
                    progress_dialog.message = 'Exporting {} ({} of {})...'.format(
                        FORMAT_LABELS[format_key],
                        index,
                        total_exports
                    )

                options = _create_export_options(format_key, format_geometry, export_path)
                _apply_options_from_settings(options, format_settings)

                success = design.exportManager.execute(options)
                if not success:
                    raise RuntimeError('Fusion reported that the {} export did not complete.'.format(FORMAT_LABELS[format_key]))

                if progress_dialog:
                    progress_dialog.progressValue = index

            if settings['auto_sort_after_export']:
                if progress_dialog:
                    progress_dialog.message = 'Sorting exported files...'
                    adsk.doEvents()
                conflict_resolver = None
                if not settings['allow_overwrite']:
                    scanned_conflicts = scan_export_conflicts(temp_staging_dir, settings['sorted_output_folder'])
                    if scanned_conflicts:
                        conflict_action = _choose_sort_conflict_action(scanned_conflicts)
                        if conflict_action == 'skip':
                            export_cancelled = True
                        conflict_resolver = (lambda source, target, operation, keep_both_target, action=conflict_action: action)
                    else:
                        def _tracking_single_conflict_resolver(source, target, operation, keep_both_target):
                            nonlocal export_cancelled
                            action = _choose_single_sort_conflict_action(source, target, operation, keep_both_target)
                            if action == 'skip':
                                export_cancelled = True
                            return action
                        conflict_resolver = _tracking_single_conflict_resolver
                sort_result = process_exports(
                    temp_staging_dir,
                    settings['sorted_output_folder'],
                    allow_overwrite=settings['allow_overwrite'],
                    conflict_resolver=conflict_resolver
                )
                if sort_result.get('conflicts_skipped'):
                    export_cancelled = True

            _save_settings(settings)
            export_succeeded = True
        except Exception as exc:
            _show_error(str(exc))
        finally:
            if progress_dialog:
                try:
                    if export_succeeded:
                        progress_dialog.progressValue = len(settings['formats']) if 'settings' in locals() else progress_dialog.progressValue
                        progress_dialog.message = 'Export cancelled.' if export_cancelled else 'Export successful.'
                        adsk.doEvents()
                        time.sleep(1)
                    progress_dialog.hide()
                except Exception:
                    pass
            if temp_staging_dir and os.path.isdir(temp_staging_dir):
                try:
                    shutil.rmtree(temp_staging_dir)
                except Exception:
                    pass
            if full_root_state:
                _restore_full_root_state(design or _active_design(), full_root_state)


class DestroyHandler(adsk.core.CommandEventHandler):
    def notify(self, args):
        pass


class MarkingMenuHandler(adsk.core.MarkingMenuEventHandler):
    def notify(self, args):
        try:
            event_args = adsk.core.MarkingMenuEventArgs.cast(args)
            if not event_args:
                return

            selected_entities = event_args.selectedEntities or []
            if not selected_entities:
                return

            if not _supports_export_selection(selected_entities[0]):
                return

            linear_menu = event_args.linearMarkingMenu
            if not linear_menu:
                return

            controls = linear_menu.controls
            if not controls:
                return

            existing = controls.itemById(COMMAND_ID)
            if not existing:
                command_definition = _ui.commandDefinitions.itemById(COMMAND_ID)
                if command_definition:
                    controls.addCommand(command_definition)
        except Exception:
            pass


def run(context):
    global _app, _ui

    try:
        _app = adsk.core.Application.get()
        _ui = _app.userInterface

        update_result = _apply_pending_update_if_needed()
        if update_result:
            if update_result.get('error'):
                _ui.messageBox(
                    'Better Export could not apply the staged update:\n{}'.format(update_result['error']),
                    COMMAND_NAME
                )
            else:
                _launch_updated_addin_from_disk(context)
            return

        command_definition = _ui.commandDefinitions.itemById(COMMAND_ID)
        if not command_definition:
            command_definition = _ui.commandDefinitions.addButtonDefinition(
                COMMAND_ID,
                COMMAND_NAME,
                COMMAND_DESCRIPTION,
                os.path.join(os.path.dirname(__file__), 'resources')
            )

        on_command_created = CommandCreatedHandler()
        command_definition.commandCreated.add(on_command_created)
        _handlers.append(on_command_created)

        marking_menu_event = _ui.markingMenuDisplaying
        if marking_menu_event:
            on_marking_menu = MarkingMenuHandler()
            marking_menu_event.add(on_marking_menu)
            _handlers.append(on_marking_menu)

        workspace = _ui.workspaces.itemById(WORKSPACE_ID)
        panel = _target_toolbar_panel(workspace)
        if panel:
            control = panel.controls.itemById(COMMAND_ID)
            if not control:
                control = panel.controls.addCommand(command_definition)
            control.isPromoted = True

    except Exception:
        if _ui:
            _ui.messageBox('Add-in start failed:\n{}'.format(traceback.format_exc()))


def stop(context):
    try:
        workspace = _ui.workspaces.itemById(WORKSPACE_ID) if _ui else None
        fallback_panel = _safe_call(lambda: workspace.toolbarPanels.itemById(FALLBACK_PANEL_ID)) if workspace else None
        if fallback_panel:
            control = fallback_panel.controls.itemById(COMMAND_ID)
            if control:
                control.deleteMe()

        utilities_tab = None
        toolbar_tabs = _safe_call(lambda: workspace.toolbarTabs) if workspace else None
        if toolbar_tabs:
            for candidate_id in UTILITIES_TAB_CANDIDATE_IDS:
                utilities_tab = _safe_call(lambda cid=candidate_id: toolbar_tabs.itemById(cid))
                if utilities_tab:
                    break
            if not utilities_tab:
                utilities_tab = _toolbar_tab_by_name(workspace, 'Utilities')

        utilities_panel = _safe_call(lambda: utilities_tab.toolbarPanels.itemById(UTILITIES_PANEL_ID)) if utilities_tab else None
        if utilities_panel:
            control = utilities_panel.controls.itemById(COMMAND_ID)
            if control:
                control.deleteMe()
            if utilities_panel.controls.count == 0:
                utilities_panel.deleteMe()

        if _ui:
            marking_menu_event = _safe_call(lambda: _ui.markingMenuDisplaying)
            if marking_menu_event:
                for handler in list(_handlers):
                    if isinstance(handler, MarkingMenuHandler):
                        try:
                            marking_menu_event.remove(handler)
                        except Exception:
                            pass

            definition = _ui.commandDefinitions.itemById(COMMAND_ID)
            if definition:
                definition.deleteMe()
    except Exception:
        if _ui:
            _ui.messageBox('Add-in stop failed:\n{}'.format(traceback.format_exc()))
