import shutil
import tempfile
import unittest
from pathlib import Path

from BetterExport.update_state import (
    STATE_APPLIED,
    STATE_FAILED,
    STATE_IDLE,
    STATE_STAGED,
    applied_update_state,
    clear_update_state,
    fail_update_state,
    read_update_state,
    stage_update_state,
    startup_preference_after_apply,
    write_update_state,
)


class UpdateStateTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp(prefix="better-export-update-state-"))
        self.state_path = self.temp_dir / "update_state.json"

    def tearDown(self):
        shutil.rmtree(self.temp_dir)

    def test_staged_update_present_round_trips_as_restart_pending_state(self):
        state = stage_update_state("1.5.0", "1.4.6", "/tmp/staged", False)
        write_update_state(str(self.state_path), state)

        loaded = read_update_state(str(self.state_path))

        self.assertEqual(loaded["state"], STATE_STAGED)
        self.assertEqual(loaded["target_version"], "1.5.0")
        self.assertEqual(loaded["installed_version"], "1.4.6")

    def test_apply_failure_is_persisted_and_suppresses_normal_checks(self):
        state = stage_update_state("1.5.0", "1.4.6", "/tmp/staged", True)
        failed = fail_update_state(state, "The staged update files are missing.")
        write_update_state(str(self.state_path), failed)

        loaded = read_update_state(str(self.state_path))

        self.assertEqual(loaded["state"], STATE_FAILED)
        self.assertEqual(loaded["target_version"], "1.5.0")
        self.assertEqual(loaded["installed_version"], "1.4.6")
        self.assertIn("missing", loaded["failure_message"].lower())

    def test_successful_apply_clears_pending_state_when_state_file_is_removed(self):
        state = stage_update_state("1.5.0", "1.4.6", "/tmp/staged", False)
        write_update_state(str(self.state_path), state)
        clear_update_state(str(self.state_path))

        loaded = read_update_state(str(self.state_path))

        self.assertEqual(loaded["state"], STATE_IDLE)

    def test_applied_state_is_recorded_after_success(self):
        state = stage_update_state("1.5.0", "1.4.6", "/tmp/staged", False)
        applied = applied_update_state(state, "1.5.0")
        write_update_state(str(self.state_path), applied)

        loaded = read_update_state(str(self.state_path))

        self.assertEqual(loaded["state"], STATE_APPLIED)
        self.assertEqual(loaded["applied_version"], "1.5.0")

    def test_startup_preference_is_restored_after_success(self):
        state = stage_update_state("1.5.0", "1.4.6", "/tmp/staged", True)
        applied = applied_update_state(state, "1.5.0")

        self.assertTrue(startup_preference_after_apply(applied))


if __name__ == "__main__":
    unittest.main()
