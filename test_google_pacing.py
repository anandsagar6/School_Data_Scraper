#!/usr/bin/env python3
"""Tests for the Google search pacing and the CAPTCHA cooldown.

No browser, no network and no account. Every test that touches disk works
in a temporary folder, so running this never changes collected data.

    python3 -m unittest test_google_pacing -v
    python3 test_google_pacing.py
"""

import importlib.util
import json
import os
import shutil
import sys
import tempfile
import time
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(HERE, "find_school.py")

_spec = importlib.util.spec_from_file_location("find_school_pacing", SCRIPT)
fs = importlib.util.module_from_spec(_spec)
sys.modules["find_school_pacing"] = fs
_spec.loader.exec_module(fs)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def school(name="Test Public School", phone="011 2345 6789", **extra):
    record = fs.blank_record()
    record["School Name"] = name
    record["City"] = "Udupi"
    record["State"] = "Karnataka"
    record["School Address"] = "Main Road, Udupi"
    record[fs.OFFICIAL_PHONE_COLUMN] = phone
    record.update(extra)
    return record


class FakeBrowser:
    """Records every search it is asked for and answers from a script."""

    def __init__(self, answers):
        #: list of (text, links, why) returned in order
        self.answers = list(answers)
        self.asked = []

    def page(self):
        return object()


class FakeDirector:
    """Hands out the fake search browser, and counts how often."""

    def __init__(self, browser):
        self.browser = browser
        self.opened = 0

    def search(self):
        self.opened += 1
        return self.browser


class PacingBase(unittest.TestCase):
    """Isolates module state so one test cannot leak into the next."""

    def setUp(self):
        self.folder = tempfile.mkdtemp(prefix="school_pacing_")
        self._kept = {
            "state": dict(fs.GOOGLE_SEARCH_STATE),
            "cooldown": fs.GOOGLE_CHALLENGE_COOLDOWN_MINUTES,
            "between_queries": fs.GOOGLE_BETWEEN_QUERIES_DELAY_SECONDS,
            "between_schools": fs.GOOGLE_BETWEEN_SCHOOLS_DELAY_SECONDS,
            "to_gemini": fs.GOOGLE_TO_GEMINI_DELAY_SECONDS,
            "run_google_search": fs.run_google_search,
            "checkpoint_file": fs.CHECKPOINT_FILE,
            "output": fs.OUTPUT_DIR,
            "temp": fs.TEMP_DIR,
            "excel": fs.EXCEL_DIR,
        }
        # no test may ever actually wait
        fs.GOOGLE_BETWEEN_QUERIES_DELAY_SECONDS = (0, 0)
        fs.GOOGLE_BETWEEN_SCHOOLS_DELAY_SECONDS = (0, 0)
        fs.GOOGLE_TO_GEMINI_DELAY_SECONDS = (0, 0)
        fs.GOOGLE_SEARCH_STATE.update({"blocked_until": 0.0, "challenges": 0,
                                       "cooldowns": 0, "deferred": 0})
        fs.OUTPUT_DIR = self.folder
        fs.TEMP_DIR = os.path.join(self.folder, "Temp_Data")
        fs.EXCEL_DIR = os.path.join(self.folder, "Excel_Data")
        fs.CHECKPOINT_FILE = os.path.join(self.folder, "checkpoint.json")

    def tearDown(self):
        fs.GOOGLE_SEARCH_STATE.update(self._kept["state"])
        fs.GOOGLE_CHALLENGE_COOLDOWN_MINUTES = self._kept["cooldown"]
        fs.GOOGLE_BETWEEN_QUERIES_DELAY_SECONDS = self._kept["between_queries"]
        fs.GOOGLE_BETWEEN_SCHOOLS_DELAY_SECONDS = self._kept["between_schools"]
        fs.GOOGLE_TO_GEMINI_DELAY_SECONDS = self._kept["to_gemini"]
        fs.run_google_search = self._kept["run_google_search"]
        fs.CHECKPOINT_FILE = self._kept["checkpoint_file"]
        fs.OUTPUT_DIR = self._kept["output"]
        fs.TEMP_DIR = self._kept["temp"]
        fs.EXCEL_DIR = self._kept["excel"]
        shutil.rmtree(self.folder, ignore_errors=True)

    def answer_with(self, answers):
        """Replace run_google_search with a scripted one; record the queries."""
        self.queries = []
        answers = list(answers)

        def fake(browser, query):
            self.queries.append(query)
            if answers:
                return answers.pop(0)
            return "", [], ""

        fs.run_google_search = fake


# ---------------------------------------------------------------------------
# 1. the requested pacing values
# ---------------------------------------------------------------------------

class PacingSettings(unittest.TestCase):

    def test_between_schools_is_45_to_90(self):
        self.assertEqual(fs.GOOGLE_BETWEEN_SCHOOLS_DELAY_SECONDS, (45, 90))

    def test_between_queries_is_15_to_25(self):
        self.assertEqual(fs.GOOGLE_BETWEEN_QUERIES_DELAY_SECONDS, (15, 25))

    def test_google_to_gemini_is_5_to_12(self):
        self.assertEqual(fs.GOOGLE_TO_GEMINI_DELAY_SECONDS, (5, 12))

    def test_cooldown_is_15_to_30_minutes(self):
        self.assertEqual(fs.GOOGLE_CHALLENGE_COOLDOWN_MINUTES, (15, 30))

    def test_every_pacing_setting_is_a_range_not_a_fixed_number(self):
        for name in ("GOOGLE_BETWEEN_SCHOOLS_DELAY_SECONDS",
                     "GOOGLE_BETWEEN_QUERIES_DELAY_SECONDS",
                     "GOOGLE_TO_GEMINI_DELAY_SECONDS",
                     "GOOGLE_CHALLENGE_COOLDOWN_MINUTES"):
            span = getattr(fs, name)
            self.assertEqual(len(span), 2, name)
            self.assertLess(span[0], span[1],
                            "%s must be a real range so the wait is random"
                            % name)


class RandomnessWithinTheRange(unittest.TestCase):
    """Requirement 1: the delay is a fresh random draw inside the range."""

    def test_pause_draws_a_different_value_inside_the_range(self):
        """pause() itself, with a span small enough not to slow the suite."""
        span = (0.001, 0.004)
        draws = [fs.pause(span) for _ in range(80)]
        self.assertTrue(all(span[0] <= d <= span[1] for d in draws),
                        "every wait must fall inside the configured range")
        self.assertGreater(len(set(round(d, 6) for d in draws)), 60,
                           "the wait must be a fresh random draw each time, "
                           "not one value repeated")

    def test_pause_is_skipped_when_a_stop_was_requested(self):
        fs.RUNTIME["stop"] = True
        try:
            self.assertEqual(fs.pause((45, 90)), 0.0,
                             "a 90-second wait must not swallow a shutdown")
        finally:
            fs.RUNTIME["stop"] = False

    def test_cooldown_minutes_vary_and_stay_inside_the_range(self):
        starts = []
        for _ in range(60):
            fs.GOOGLE_SEARCH_STATE["blocked_until"] = 0.0
            before = time.time()
            minutes = fs.start_google_cooldown()
            starts.append(minutes)
            left = fs.GOOGLE_SEARCH_STATE["blocked_until"] - before
            self.assertGreaterEqual(left, 15 * 60 - 1)
            self.assertLessEqual(left, 30 * 60 + 1)
        self.assertGreater(len(set(round(m, 3) for m in starts)), 40,
                           "each cooldown must be a fresh random draw so ten "
                           "ranks do not resume together")


# ---------------------------------------------------------------------------
# 2. the cooldown itself
# ---------------------------------------------------------------------------

class CooldownBehaviour(PacingBase):

    def test_a_challenge_starts_a_cooldown(self):
        self.answer_with([("", [], "Google served a challenge for this search")])
        browser = FakeBrowser([])
        director = FakeDirector(browser)
        self.assertEqual(fs.google_cooldown_left(), 0.0)
        found = fs.find_official_contacts(director, school())
        self.assertGreater(fs.google_cooldown_left(), 14 * 60)
        self.assertEqual(fs.GOOGLE_SEARCH_STATE["cooldowns"], 1)
        self.assertTrue(found["deferred"])

    def test_the_challenged_page_is_never_requested_again(self):
        """Requirement 6: no bypass, and no retry of the challenged query."""
        self.answer_with([("", [], "Google served a challenge for this search")])
        fs.find_official_contacts(FakeDirector(FakeBrowser([])), school())
        self.assertEqual(len(self.queries), 1,
                         "after a challenge the run must stop asking, not ask "
                         "again: %r" % (self.queries,))

    def test_during_cooldown_no_search_is_attempted_at_all(self):
        self.answer_with([("text", [], "")])
        fs.start_google_cooldown()
        director = FakeDirector(FakeBrowser([]))
        found = fs.find_official_contacts(director, school())
        self.assertEqual(self.queries, [],
                         "no query may be sent while in cooldown")
        self.assertEqual(director.opened, 0,
                         "the search browser must not even be opened")
        self.assertTrue(found["deferred"])
        self.assertTrue(any("cooldown" in n for n in found["notes"]))

    def test_cooldown_keeps_the_google_maps_phone(self):
        """Requirement 5: the Maps fallback survives a cooldown."""
        fs.start_google_cooldown()
        record = school(phone="011 2345 6789")
        found = fs.find_official_contacts(FakeDirector(FakeBrowser([])), record)
        fs.merge_official_contacts(record, found)
        self.assertEqual(record[fs.OFFICIAL_PHONE_COLUMN], "011 2345 6789")

    def test_cooldown_never_writes_a_fabricated_contact(self):
        fs.start_google_cooldown()
        record = school(phone="")
        found = fs.find_official_contacts(FakeDirector(FakeBrowser([])), record)
        fs.merge_official_contacts(record, found)
        self.assertEqual(fs.clean(record[fs.OFFICIAL_EMAIL_COLUMN]), "")
        self.assertEqual(fs.clean(record[fs.OFFICIAL_PHONE_COLUMN]), "")

    def test_searching_resumes_when_the_cooldown_expires(self):
        fs.start_google_cooldown()
        self.assertGreater(fs.google_cooldown_left(), 0)
        fs.GOOGLE_SEARCH_STATE["blocked_until"] = time.time() - 1
        self.assertEqual(fs.google_cooldown_left(), 0.0)
        self.answer_with([
            ("Test Public School Udupi Karnataka info@testschool.in", [], ""),
            ("Test Public School Udupi Karnataka 011 2345 6789", [], ""),
        ])
        found = fs.find_official_contacts(FakeDirector(FakeBrowser([])),
                                          school(phone=""))
        self.assertEqual(len(self.queries), 2,
                         "both searches must run once the cooldown is over")
        self.assertFalse(found["deferred"])

    def test_cooldown_note_counts_down_in_whole_minutes(self):
        fs.GOOGLE_SEARCH_STATE["blocked_until"] = time.time() + 61
        note = fs.google_cooldown_note()
        self.assertIn("cooldown", note)
        self.assertIn("minute", note)
        fs.GOOGLE_SEARCH_STATE["blocked_until"] = 0.0
        self.assertEqual(fs.google_cooldown_note(), "")


# ---------------------------------------------------------------------------
# 3. both searches run independently when Google is available
# ---------------------------------------------------------------------------

class BothSearchesIndependent(PacingBase):

    def test_email_and_phone_are_searched_separately(self):
        """Requirement 4, and the phone search is not skipped for a Maps phone."""
        self.answer_with([
            ("Test Public School Udupi Karnataka contact info@testschool.in",
             [], ""),
            ("Test Public School Udupi Karnataka phone 011 2345 6789", [], ""),
        ])
        found = fs.find_official_contacts(FakeDirector(FakeBrowser([])),
                                          school(phone="011 9999 0000"))
        self.assertEqual(len(self.queries), 2)
        self.assertIn("email", self.queries[0].lower())
        self.assertIn("phone", self.queries[1].lower())
        self.assertEqual(found["queries"], ["email", "phone"])

    def test_a_browser_failure_defers_rather_than_settling_the_school(self):
        def boom(browser, query):
            raise fs.PlaywrightError("browser went away")
        fs.run_google_search = boom
        found = fs.find_official_contacts(FakeDirector(FakeBrowser([])),
                                          school())
        self.assertTrue(found["deferred"])
        self.assertEqual(fs.google_cooldown_left(), 0.0,
                         "a browser fault is not a Google challenge, so it "
                         "must not start a cooldown")


# ---------------------------------------------------------------------------
# 4. unresolved searches are preserved for a later retry
# ---------------------------------------------------------------------------

class PendingContactsAreRemembered(PacingBase):

    def test_a_deferred_school_is_recorded(self):
        checkpoint = fs.Checkpoint()
        record = school(name="Deferred School")
        added = checkpoint.note_contacts_pending("Karnataka", "Udupi", record,
                                                 reason="cooldown")
        self.assertTrue(added)
        pending = checkpoint.data["contacts_pending"]
        self.assertEqual(len(pending), 1)
        self.assertEqual(pending[0]["school"], "Deferred School")
        self.assertEqual(pending[0]["city"], "Udupi")
        self.assertEqual(pending[0]["state"], "Karnataka")
        self.assertTrue(pending[0]["when"])

    def test_the_same_school_is_not_recorded_twice(self):
        checkpoint = fs.Checkpoint()
        record = school(name="Repeated School")
        self.assertTrue(checkpoint.note_contacts_pending("Karnataka", "Udupi",
                                                         record))
        self.assertFalse(checkpoint.note_contacts_pending("Karnataka", "Udupi",
                                                          record))
        self.assertEqual(len(checkpoint.data["contacts_pending"]), 1)

    def test_a_nameless_school_is_not_recorded(self):
        checkpoint = fs.Checkpoint()
        record = school(name="")
        self.assertFalse(checkpoint.note_contacts_pending("Karnataka", "Udupi",
                                                          record))
        self.assertEqual(checkpoint.data["contacts_pending"], [])

    def test_the_list_is_capped(self):
        checkpoint = fs.Checkpoint()
        checkpoint.CONTACTS_PENDING_LIMIT = 3
        for n in range(6):
            checkpoint.note_contacts_pending("Karnataka", "Udupi",
                                             school(name="School %d" % n))
        self.assertEqual(len(checkpoint.data["contacts_pending"]), 3,
                         "a long block must not grow the checkpoint without "
                         "limit")

    def test_it_survives_a_reload(self):
        checkpoint = fs.Checkpoint()
        checkpoint.note_contacts_pending("Karnataka", "Udupi",
                                         school(name="Survives Restart"))
        again = fs.Checkpoint()
        self.assertEqual([e["school"] for e in
                          again.data.get("contacts_pending", [])],
                         ["Survives Restart"])


# ---------------------------------------------------------------------------
# 5. the new key must not disturb a checkpoint written before it existed
# ---------------------------------------------------------------------------

class OlderCheckpointsStillLoad(PacingBase):
    """The eight running ranks have checkpoints with no contacts_pending."""

    def _live_looking_checkpoint(self):
        """A checkpoint exactly as a running rank writes it today."""
        data = fs.empty_checkpoint()
        data.pop("contacts_pending")            # the key did not exist yet
        data.update({
            "rank": 10,
            "num_sys": 10,
            "current_city": ["Uttar Pradesh", "Shahjahanpur"],
            "current_keyword": "Private Schools",
            "completed_cities": ["Uttar Pradesh|Rampur"],
            "committed": {"Uttar Pradesh|Shahjahanpur": ["email:a@b.c"]},
            "enriched": {"Uttar Pradesh|Shahjahanpur": ["email:a@b.c"]},
            "saved": 46,
            "duplicates": 3,
        })
        with open(fs.CHECKPOINT_FILE, "w", encoding="utf-8") as handle:
            json.dump(data, handle)
        return data

    def test_an_existing_checkpoint_is_not_set_aside(self):
        written = self._live_looking_checkpoint()
        loaded = fs.Checkpoint()
        self.assertEqual(loaded.data["rank"], 10)
        self.assertEqual(loaded.data["num_sys"], 10)
        self.assertEqual(loaded.data["saved"], 46)
        self.assertEqual(loaded.data["duplicates"], 3)
        self.assertEqual(loaded.data["current_city"],
                         ["Uttar Pradesh", "Shahjahanpur"])
        self.assertEqual(loaded.data["completed_cities"],
                         ["Uttar Pradesh|Rampur"])
        self.assertEqual(loaded.data["committed"], written["committed"])

    def test_the_new_key_defaults_to_empty(self):
        self._live_looking_checkpoint()
        loaded = fs.Checkpoint()
        self.assertEqual(loaded.data.get("contacts_pending"), [])

    def test_the_version_was_not_bumped(self):
        """Bumping it would set aside every live rank's progress."""
        self.assertEqual(fs.CHECKPOINT_VERSION, 3)


if __name__ == "__main__":
    unittest.main(verbosity=2)
