#!/usr/bin/env python3
"""Tests for the night stop and its per-machine configuration.

The window is decided by night_seconds_remaining(now), which takes the
clock as an argument - so every hour of the day is tested in under a
second and nothing waits for 11pm.

    python3 -m unittest test_night_pause -v
    python3 test_night_pause.py
"""

import datetime
import importlib.util
import json
import os
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(HERE, "find_school.py")

_spec = importlib.util.spec_from_file_location("find_school_night", SCRIPT)
fs = importlib.util.module_from_spec(_spec)
sys.modules["find_school_night"] = fs
_spec.loader.exec_module(fs)


def at(hour, minute=0):
    """A timezone-aware moment on the configured clock."""
    import zoneinfo
    return datetime.datetime(2026, 11, 18, hour, minute, 0,
                             tzinfo=zoneinfo.ZoneInfo(fs.TIME_ZONE))


class NightBase(unittest.TestCase):

    def setUp(self):
        self._kept = {n: getattr(fs, n) for n in
                      ("NIGHT_PAUSE_ENABLED", "NIGHT_START_HOUR",
                       "NIGHT_END_HOUR", "TIME_ZONE", "RANK_FILE")}
        self._source = dict(fs.SCHEDULE_SOURCE)
        fs.NIGHT_PAUSE_ENABLED = True
        fs.NIGHT_START_HOUR = 23
        fs.NIGHT_END_HOUR = 5
        fs.TIME_ZONE = "Asia/Kolkata"
        self.folder = tempfile.mkdtemp(prefix="school_night_")
        fs.RANK_FILE = os.path.join(self.folder, "rank.json")

    def tearDown(self):
        for name, value in self._kept.items():
            setattr(fs, name, value)
        fs.SCHEDULE_SOURCE.clear()
        fs.SCHEDULE_SOURCE.update(self._source)
        shutil.rmtree(self.folder, ignore_errors=True)

    def write_rank_file(self, payload):
        with open(fs.RANK_FILE, "w", encoding="utf-8") as handle:
            if isinstance(payload, str):
                handle.write(payload)
            else:
                json.dump(payload, handle)


# ---------------------------------------------------------------------------
# 1. the window, every hour of the day, without waiting
# ---------------------------------------------------------------------------

class TheWindow(NightBase):

    def test_inside_the_window_just_after_11pm(self):
        self.assertGreater(fs.night_seconds_remaining(at(23, 0)), 0)
        self.assertGreater(fs.night_seconds_remaining(at(23, 30)), 0)

    def test_inside_across_midnight(self):
        for hour in (0, 1, 2, 3, 4):
            self.assertGreater(fs.night_seconds_remaining(at(hour, 30)), 0,
                               "%02d:30 must be inside the night stop" % hour)

    def test_outside_from_5am(self):
        self.assertEqual(fs.night_seconds_remaining(at(5, 0)), 0)
        self.assertEqual(fs.night_seconds_remaining(at(5, 1)), 0)

    def test_outside_during_the_working_day(self):
        for hour in (6, 9, 12, 15, 18, 21, 22):
            self.assertEqual(fs.night_seconds_remaining(at(hour, 30)), 0,
                             "%02d:30 must be outside the night stop" % hour)

    def test_every_hour_is_either_in_or_out_and_the_count_is_right(self):
        inside = [h for h in range(24)
                  if fs.night_seconds_remaining(at(h, 30)) > 0]
        self.assertEqual(inside, [0, 1, 2, 3, 4, 23],
                         "23:00-05:00 must cover exactly six hours")

    def test_seconds_remaining_counts_down_to_5am(self):
        # 23:00 -> 05:00 is six hours
        self.assertEqual(fs.night_seconds_remaining(at(23, 0)), 6 * 3600)
        # 04:00 -> 05:00 is one hour
        self.assertEqual(fs.night_seconds_remaining(at(4, 0)), 3600)
        # 04:59 -> 05:00 is a minute
        self.assertEqual(fs.night_seconds_remaining(at(4, 59)), 60)

    def test_disabled_means_never_inside(self):
        fs.NIGHT_PAUSE_ENABLED = False
        for hour in range(24):
            self.assertEqual(fs.night_seconds_remaining(at(hour, 30)), 0)

    def test_equal_hours_switch_it_off_rather_than_pausing_all_day(self):
        fs.NIGHT_START_HOUR = fs.NIGHT_END_HOUR = 23
        for hour in range(24):
            self.assertEqual(fs.night_seconds_remaining(at(hour, 30)), 0)


# ---------------------------------------------------------------------------
# 2. the run actually takes the pause
# ---------------------------------------------------------------------------

class TheRunStops(NightBase):

    def test_a_scheduled_stop_is_due_at_11pm(self):
        due = fs.due_pause(at(23, 10))
        self.assertIsNotNone(due, "11:10pm must be a due night stop")
        reason, seconds = due
        self.assertIn("night stop", reason)
        self.assertIn("23:00-05:00", reason)
        self.assertIn("Asia/Kolkata", reason)
        self.assertGreater(seconds, 0)

    def test_no_stop_is_due_in_the_afternoon(self):
        fs.WORK["started"] = None
        self.assertIsNone(fs.due_pause(at(15, 0)))

    def test_night_is_checked_before_the_work_break(self):
        """A period ending at 22:59 must not break and then stop again."""
        due = fs.due_pause(at(23, 30))
        self.assertIn("night stop", due[0])

    def test_the_status_line_says_when_it_is_inside(self):
        fs.NIGHT_PAUSE_ENABLED = True
        line = fs.night_pause_status()
        self.assertIn("23:00-05:00", line)
        self.assertIn("Asia/Kolkata", line)

    def test_the_status_line_says_off_when_disabled(self):
        fs.NIGHT_PAUSE_ENABLED = False
        self.assertIn("off", fs.night_pause_status())


# ---------------------------------------------------------------------------
# 3. the per-machine setting, which is the durable part
# ---------------------------------------------------------------------------

class RankFileConfiguration(NightBase):

    def test_rank_json_turns_the_night_stop_on(self):
        fs.NIGHT_PAUSE_ENABLED = False
        fs.NIGHT_START_HOUR, fs.NIGHT_END_HOUR = 1, 2
        self.write_rank_file({"rank": 7, "num_sys": 10,
                              "night_pause_enabled": True,
                              "night_start_hour": 23,
                              "night_end_hour": 5,
                              "time_zone": "Asia/Kolkata"})
        note = fs.apply_rank_file_schedule()
        self.assertTrue(fs.NIGHT_PAUSE_ENABLED)
        self.assertEqual(fs.NIGHT_START_HOUR, 23)
        self.assertEqual(fs.NIGHT_END_HOUR, 5)
        self.assertEqual(fs.TIME_ZONE, "Asia/Kolkata")
        self.assertEqual(fs.SCHEDULE_SOURCE["night"], "rank.json")
        self.assertIn("rank.json", note)

    def test_a_missing_file_leaves_the_builtins_alone(self):
        self.assertEqual(fs.apply_rank_file_schedule(), "")
        self.assertEqual(fs.SCHEDULE_SOURCE["night"], self._source["night"])

    def test_a_file_without_night_keys_changes_nothing(self):
        self.write_rank_file({"rank": 7, "num_sys": 10})
        self.assertEqual(fs.apply_rank_file_schedule(), "")
        self.assertEqual(fs.NIGHT_START_HOUR, 23)

    def test_an_unreadable_file_is_reported_not_fatal(self):
        self.write_rank_file("{not json")
        note = fs.apply_rank_file_schedule()
        self.assertIn("could not be read", note)
        self.assertEqual(fs.NIGHT_START_HOUR, 23)

    def test_an_out_of_range_hour_applies_NOTHING(self):
        """Half-applying a schedule is worse than ignoring the file."""
        fs.NIGHT_PAUSE_ENABLED = False
        self.write_rank_file({"night_pause_enabled": True,
                              "night_start_hour": 99,
                              "night_end_hour": 5})
        note = fs.apply_rank_file_schedule()
        self.assertIn("NOTHING", note)
        self.assertFalse(fs.NIGHT_PAUSE_ENABLED,
                         "night_pause_enabled must NOT survive a bad hour")
        self.assertEqual(fs.SCHEDULE_SOURCE["night"], self._source["night"])

    def test_a_non_numeric_hour_applies_NOTHING(self):
        fs.NIGHT_PAUSE_ENABLED = False
        self.write_rank_file({"night_pause_enabled": True,
                              "night_start_hour": "eleven"})
        self.assertIn("NOTHING", fs.apply_rank_file_schedule())
        self.assertFalse(fs.NIGHT_PAUSE_ENABLED)

    def test_an_unknown_time_zone_applies_NOTHING(self):
        fs.NIGHT_PAUSE_ENABLED = False
        self.write_rank_file({"night_pause_enabled": True,
                              "night_start_hour": 22,
                              "time_zone": "Mars/Olympus"})
        self.assertIn("NOTHING", fs.apply_rank_file_schedule())
        self.assertFalse(fs.NIGHT_PAUSE_ENABLED)
        self.assertEqual(fs.NIGHT_START_HOUR, 23,
                         "the valid hour must not survive either")

    def test_a_list_instead_of_an_object_is_ignored(self):
        self.write_rank_file([1, 2, 3])
        self.assertEqual(fs.apply_rank_file_schedule(), "")

    def test_it_can_also_turn_the_night_stop_off(self):
        self.write_rank_file({"night_pause_enabled": False})
        fs.apply_rank_file_schedule()
        self.assertFalse(fs.NIGHT_PAUSE_ENABLED)

    def test_rank_json_is_gitignored_so_it_survives_source_updates(self):
        """The whole reason the setting lives here and not at the top."""
        ignore = os.path.join(HERE, ".gitignore")
        self.assertTrue(os.path.exists(ignore))
        with open(ignore, encoding="utf-8") as handle:
            lines = [l.strip() for l in handle]
        self.assertIn("rank.json", lines,
                      "if rank.json stops being ignored, a source update "
                      "could overwrite the night stop")


if __name__ == "__main__":
    unittest.main(verbosity=2)
