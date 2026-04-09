import json
import os
import time


STATE_IDLE = 'idle'
STATE_STAGED = 'staged'
STATE_FAILED = 'failed'
STATE_APPLIED = 'applied'
VALID_STATES = {STATE_IDLE, STATE_STAGED, STATE_FAILED, STATE_APPLIED}


def _as_bool(value, default=False):
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    return bool(value)


def empty_update_state():
    return {
        'state': STATE_IDLE,
        'target_version': '',
        'installed_version': '',
        'staged_addin_dir': '',
        'staged_at': 0.0,
        'previous_run_on_startup': False,
        'failure_message': '',
        'failed_at': 0.0,
        'applied_version': '',
        'applied_at': 0.0
    }


def normalize_update_state(value):
    state = empty_update_state()
    if isinstance(value, dict):
        candidate = str(value.get('state') or '').strip().lower()
        if candidate in VALID_STATES:
            state['state'] = candidate
        for key in ('target_version', 'installed_version', 'staged_addin_dir', 'failure_message', 'applied_version'):
            if isinstance(value.get(key), str):
                state[key] = value[key].strip()
        for key in ('staged_at', 'failed_at', 'applied_at'):
            if isinstance(value.get(key), (int, float)):
                state[key] = float(value[key])
        state['previous_run_on_startup'] = _as_bool(value.get('previous_run_on_startup'), False)

    if state['state'] == STATE_IDLE:
        return empty_update_state()
    return state


def read_update_state(path):
    if not os.path.exists(path):
        return empty_update_state()
    try:
        with open(path, 'r', encoding='utf-8') as handle:
            return normalize_update_state(json.load(handle))
    except Exception:
        return empty_update_state()


def write_update_state(path, value):
    normalized = normalize_update_state(value)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as handle:
        json.dump(normalized, handle, indent=2, sort_keys=True)
    return normalized


def clear_update_state(path):
    try:
        if os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


def stage_update_state(target_version, installed_version, staged_addin_dir, previous_run_on_startup):
    return normalize_update_state({
        'state': STATE_STAGED,
        'target_version': str(target_version or '').strip(),
        'installed_version': str(installed_version or '').strip(),
        'staged_addin_dir': str(staged_addin_dir or '').strip(),
        'staged_at': time.time(),
        'previous_run_on_startup': bool(previous_run_on_startup),
        'failure_message': '',
        'failed_at': 0.0,
        'applied_version': '',
        'applied_at': 0.0
    })


def fail_update_state(current_state, message):
    state = normalize_update_state(current_state)
    state['state'] = STATE_FAILED
    state['failure_message'] = str(message or '').strip()
    state['failed_at'] = time.time()
    state['applied_version'] = ''
    state['applied_at'] = 0.0
    state['staged_addin_dir'] = ''
    return state


def applied_update_state(current_state, applied_version):
    state = normalize_update_state(current_state)
    state['state'] = STATE_APPLIED
    state['applied_version'] = str(applied_version or '').strip()
    state['applied_at'] = time.time()
    state['staged_addin_dir'] = ''
    state['failure_message'] = ''
    state['failed_at'] = 0.0
    return state


def startup_preference_after_apply(current_state):
    state = normalize_update_state(current_state)
    return bool(state.get('previous_run_on_startup'))
