import importlib
import json
import shutil
import sys
import tempfile
import types
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from BetterExport.update_state import STATE_FAILED, read_update_state, stage_update_state


def _install_adsk_stub():
    if 'adsk' in sys.modules:
        return

    adsk = types.ModuleType('adsk')
    core = types.ModuleType('adsk.core')
    fusion = types.ModuleType('adsk.fusion')

    def _cast(value):
        return value

    for name in (
        'CommandCreatedEventHandler',
        'InputChangedEventHandler',
        'ValidateInputsEventHandler',
        'CommandEventHandler',
        'CustomEventHandler',
        'MarkingMenuEventHandler',
    ):
        setattr(core, name, type(name, (), {}))

    for name in (
        'TextBoxCommandInput',
        'BoolValueCommandInput',
        'StringValueCommandInput',
        'SelectionCommandInput',
        'DropDownCommandInput',
        'GroupCommandInput',
        'MarkingMenuEventArgs',
    ):
        setattr(core, name, type(name, (), {'cast': staticmethod(_cast)}))

    core.MessageBoxButtonTypes = types.SimpleNamespace(YesNoCancelButtonType=1)
    core.MessageBoxIconTypes = types.SimpleNamespace(WarningIconType=1)
    core.DialogResults = types.SimpleNamespace(DialogYes=1, DialogNo=2, DialogOK=3)
    core.DropDownStyles = types.SimpleNamespace(TextListDropDownStyle=1)
    core.TablePresentationStyles = types.SimpleNamespace(itemBorderTablePresentationStyle=1)
    core.Application = type('Application', (), {'get': staticmethod(lambda: None)})

    for name in ('Design', 'BRepBody', 'Occurrence', 'Component'):
        setattr(fusion, name, type(name, (), {'cast': staticmethod(_cast)}))

    adsk.core = core
    adsk.fusion = fusion
    adsk.doEvents = lambda: None

    sys.modules['adsk'] = adsk
    sys.modules['adsk.core'] = core
    sys.modules['adsk.fusion'] = fusion


_install_adsk_stub()
better_export = importlib.import_module('BetterExport.BetterExport')


class BetterExportUpdateRuntimeTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix='better-export-runtime-'))
        self.addin_dir = self.temp_dir / 'BetterExport'
        self.addin_dir.mkdir()
        self.pending_dir = self.addin_dir / '_pending_update'
        self.pending_dir.mkdir()
        self.pending_info_path = self.pending_dir / 'update.json'
        self.update_helper_path = self.addin_dir / 'update_helper.py'
        self.update_state_path = self.addin_dir / 'update_state.json'
        self.manifest_path = self.addin_dir / 'BetterExport.manifest'
        self.manifest_path.write_text(json.dumps({'version': '1.4.6'}), encoding='utf-8')

        self.startup_state = {'value': False}
        self.startup_calls = []

        self.patchers = [
            mock.patch.object(better_export, 'ADDIN_DIR', str(self.addin_dir)),
            mock.patch.object(better_export, 'PENDING_UPDATE_DIR', str(self.pending_dir)),
            mock.patch.object(better_export, 'PENDING_UPDATE_INFO_PATH', str(self.pending_info_path)),
            mock.patch.object(better_export, 'UPDATE_HELPER_PATH', str(self.update_helper_path)),
            mock.patch.object(better_export, 'UPDATE_STATE_PATH', str(self.update_state_path)),
            mock.patch.object(better_export, 'MANIFEST_PATH', str(self.manifest_path)),
            mock.patch.object(
                better_export,
                '_script_item_for_addin',
                side_effect=lambda: types.SimpleNamespace(isRunOnStartup=self.startup_state['value'])
            ),
            mock.patch.object(better_export, '_set_run_on_startup', side_effect=self._set_run_on_startup),
        ]
        for patcher in self.patchers:
            patcher.start()
        self.addCleanup(self._stop_patchers)

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def _stop_patchers(self):
        for patcher in reversed(self.patchers):
            patcher.stop()

    def _set_run_on_startup(self, enabled):
        self.startup_calls.append(bool(enabled))
        self.startup_state['value'] = bool(enabled)

    def test_missing_pending_files_transition_staged_update_to_failed(self):
        staged_state = stage_update_state('1.5.0', '1.4.6', str(self.pending_dir / 'staged'), True)
        better_export._write_current_update_state(staged_state)

        result = better_export._apply_pending_update_if_needed()

        self.assertEqual(result['status'], 'failed')
        self.assertIn('missing', result['error'].lower())
        persisted = read_update_state(str(self.update_state_path))
        self.assertEqual(persisted['state'], STATE_FAILED)
        self.assertIn('missing', persisted['failure_message'].lower())
        self.assertEqual(json.loads(self.manifest_path.read_text(encoding='utf-8'))['version'], '1.4.6')
        self.assertEqual(self.startup_calls[-1], True)

    def test_stage_failure_rolls_back_manifest_and_startup_side_effects(self):
        def write_asset_zip(_asset_url, destination_path):
            with zipfile.ZipFile(destination_path, 'w') as archive:
                archive.writestr('BetterExport/BetterExport.py', '# staged update')

        release_info = {
            'latest_version': '1.5.0',
            'latest_asset_url': 'https://example.invalid/BetterExport-1.5.0.zip',
            'latest_asset_name': 'BetterExport-1.5.0.zip',
        }

        with mock.patch.object(better_export, '_download_release_asset', side_effect=write_asset_zip), \
             mock.patch.object(better_export, '_write_current_update_state', side_effect=OSError('disk full')):
            with self.assertRaises(OSError):
                better_export._stage_update_payload(release_info)

        self.assertEqual(self.startup_calls, [True, False])
        self.assertFalse(self.startup_state['value'])
        self.assertEqual(json.loads(self.manifest_path.read_text(encoding='utf-8'))['version'], '1.4.6')
        self.assertFalse(self.pending_dir.exists())


if __name__ == '__main__':
    unittest.main()
