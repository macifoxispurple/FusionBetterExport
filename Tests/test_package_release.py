import importlib.util
import tempfile
import unittest
import zipfile
from pathlib import Path


def _load_package_release_module():
    repo_root = Path(__file__).resolve().parents[1]
    script_path = repo_root / 'scripts' / 'package_release.py'
    spec = importlib.util.spec_from_file_location('package_release', script_path)
    module = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    spec.loader.exec_module(module)
    return module


package_release = _load_package_release_module()


class PackageReleaseTests(unittest.TestCase):
    def test_canonical_build_uses_forward_slashes_and_root_prefix(self):
        with tempfile.TemporaryDirectory(prefix='better-export-package-') as temp_dir:
            output_zip = Path(temp_dir) / 'BetterExport-test.zip'
            built_path = package_release.build_release_zip(output_zip)
            self.assertEqual(built_path, output_zip)

            with zipfile.ZipFile(output_zip, 'r') as archive:
                names = [info.filename for info in archive.infolist()]

            self.assertTrue(names, 'Expected release zip to contain files.')
            for name in names:
                self.assertNotIn('\\', name)
                self.assertTrue(name.startswith('BetterExport/'), name)
                self.assertNotIn('/..', '/' + name)

    def test_validator_rejects_backslash_entries(self):
        with self.assertRaises(ValueError):
            package_release.validate_entry_name('BetterExport\\BetterExport.py')


if __name__ == '__main__':
    unittest.main()
