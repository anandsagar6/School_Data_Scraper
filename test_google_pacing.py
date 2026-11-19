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
            # BACKUP_DIR matters as much as the rest: the retry takes a
            # pre-update copy of every workbook it touches, and without
            # this the suite wrote three files into the REAL Backup/ of
            # whichever install it was run in. It did that on all ten
            # ranks before this was caught.
            "backup": fs.BACKUP_DIR,
            "project": fs.PROJECT_DIR,
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
        fs.BACKUP_DIR = os.path.join(self.folder, "Backup")
        # the cooldown share file lives one level above PROJECT_DIR, so
        # this keeps the suite out of the real install base too
        fs.PROJECT_DIR = os.path.join(self.folder, "rank_test")
        os.makedirs(fs.PROJECT_DIR, exist_ok=True)

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
        fs.BACKUP_DIR = self._kept["backup"]
        fs.PROJECT_DIR = self._kept["project"]
        shutil.rmtree(self.folder, ignore_errors=True)

    def reset_google_state(self):
        """Clear BOTH the in-process cooldown and the shared file.

        Setting GOOGLE_SEARCH_STATE alone is not enough: google_cooldown_left()
        deliberately adopts the longest window any rank on the laptop
        published, so a leftover share file puts the cooldown straight back.
        """
        fs.GOOGLE_SEARCH_STATE.update({"blocked_until": 0.0, "challenges": 0,
                                       "cooldowns": 0, "deferred": 0,
                                       "streak": 0})
        path = fs.cooldown_share_path()
        if path and os.path.exists(path):
            os.remove(path)

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

    #: The pacing the senior has configured. These three were (45, 90),
    #: (15, 25) and (5, 12) when this suite was written; the senior has
    #: since widened them, and that configuration is authoritative - a
    #: source update must never quietly put the older values back. The
    #: table is here so changing them stays a deliberate one-line edit.
    SENIOR_PACING = {
        "GOOGLE_BETWEEN_SCHOOLS_DELAY_SECONDS": (30, 150),
        "GOOGLE_BETWEEN_QUERIES_DELAY_SECONDS": (20, 120),
        "GOOGLE_TO_GEMINI_DELAY_SECONDS": (25, 130),
    }

    def test_the_configured_pacing_is_what_the_senior_set(self):
        for name, expected in self.SENIOR_PACING.items():
            self.assertEqual(getattr(fs, name), expected, name)

    def test_each_delay_is_an_ordered_pair_of_positive_seconds(self):
        """What the three old value tests were really guarding: a usable
        (min, max) range, whatever numbers are in it."""
        for name in self.SENIOR_PACING:
            span = getattr(fs, name)
            self.assertIsInstance(span, tuple, name)
            self.assertEqual(len(span), 2, name)
            low, high = span
            self.assertGreater(low, 0, name)
            self.assertLess(low, high, name)

    def test_no_delay_is_faster_than_the_old_upstream_default(self):
        """A widened range must not become a FASTER one by accident."""
        was = {"GOOGLE_BETWEEN_SCHOOLS_DELAY_SECONDS": 45,
               "GOOGLE_BETWEEN_QUERIES_DELAY_SECONDS": 15,
               "GOOGLE_TO_GEMINI_DELAY_SECONDS": 5}
        for name, old_low in was.items():
            low, high = getattr(fs, name)
            self.assertGreaterEqual((low + high) / 2.0, old_low,
                                    "%s is now faster on average" % name)

    def test_cooldown_is_60_to_120_minutes(self):
        self.assertEqual(fs.GOOGLE_CHALLENGE_COOLDOWN_MINUTES, (60, 120))

    def test_cooldown_ceiling_is_four_hours(self):
        self.assertEqual(fs.GOOGLE_CHALLENGE_COOLDOWN_MAX_MINUTES, 240)

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

    def test_cooldown_minutes_vary_and_stay_inside_the_first_band(self):
        starts = []
        for _ in range(60):
            fs.GOOGLE_SEARCH_STATE["blocked_until"] = 0.0
            fs.GOOGLE_SEARCH_STATE["streak"] = 0
            minutes, low, high = fs.cooldown_window_minutes(1)
            starts.append(minutes)
            self.assertGreaterEqual(minutes, 60)
            self.assertLessEqual(minutes, 120)
            self.assertEqual((low, high), (60.0, 120.0))
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
        # the window elapses, locally and in the shared file
        self.reset_google_state()
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

class TheSuiteStaysInsideItsOwnFolder(PacingBase):
    """Regression: this suite wrote into ten production Backup/ folders.

    PacingBase isolated OUTPUT_DIR, TEMP_DIR, EXCEL_DIR and CHECKPOINT_FILE
    but NOT BACKUP_DIR, so the retry's pre-update copies landed in the real
    install's Backup/ on every rank the suite was run in. Three files per
    rank, on all ten. Nothing was damaged - they were new files beside no
    genuine backups - but recovery_paths() globs Backup/*.xlsx, so a sync
    would have carried test fixtures off as if they were real backups.
    """

    def test_every_written_path_is_under_the_temporary_folder(self):
        for name in ("OUTPUT_DIR", "TEMP_DIR", "EXCEL_DIR", "BACKUP_DIR",
                     "CHECKPOINT_FILE", "PROJECT_DIR"):
            value = getattr(fs, name)
            self.assertTrue(os.path.abspath(value).startswith(
                os.path.abspath(self.folder)),
                "%s points outside the test folder: %s" % (name, value))

    def test_a_retry_backup_lands_in_the_test_folder(self):
        os.makedirs(fs.EXCEL_DIR, exist_ok=True)
        rec = school(name="Backup Path School", phone="")
        rec["City"] = "Udupi"
        fs.write_records(fs.master_path(), [rec])
        kept = fs.backup_before_contact_update(fs.master_path())
        self.assertTrue(kept)
        self.assertTrue(os.path.abspath(kept).startswith(
            os.path.abspath(self.folder)),
            "the pre-update backup escaped the test folder: %s" % kept)

    def test_the_cooldown_share_file_lands_in_the_test_folder(self):
        fs.write_shared_cooldown(time.time() + 60, 1)
        path = fs.cooldown_share_path()
        self.assertTrue(os.path.abspath(path).startswith(
            os.path.abspath(self.folder)),
            "the shared cooldown file escaped the test folder: %s" % path)


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


# ---------------------------------------------------------------------------
# 6. bounded exponential backoff, and sharing it between ranks
# ---------------------------------------------------------------------------

class BoundedBackoff(PacingBase):

    def test_each_consecutive_challenge_doubles_the_band(self):
        self.assertEqual(fs.cooldown_window_minutes(1)[1:], (60.0, 120.0))
        self.assertEqual(fs.cooldown_window_minutes(2)[1:], (120.0, 240.0))

    def test_the_band_is_capped_at_four_hours(self):
        for streak in (3, 4, 8, 50):
            minutes, low, high = fs.cooldown_window_minutes(streak)
            self.assertEqual((low, high), (240.0, 240.0),
                             "streak %d must be capped" % streak)
            self.assertLessEqual(minutes, 240.0)

    def test_the_streak_climbs_with_each_challenge(self):
        self.answer_with([("", [], "challenge")] * 3)
        for expected in (1, 2, 3):
            # clear the window but KEEP the streak - that is what a
            # cooldown expiring looks like
            fs.GOOGLE_SEARCH_STATE["blocked_until"] = 0.0
            path = fs.cooldown_share_path()
            if path and os.path.exists(path):
                os.remove(path)
            fs.find_official_contacts(FakeDirector(FakeBrowser([])), school())
            self.assertEqual(fs.GOOGLE_SEARCH_STATE["streak"], expected)

    def test_a_served_search_resets_the_streak(self):
        fs.GOOGLE_SEARCH_STATE["streak"] = 3
        fs.note_google_search_served()
        self.assertEqual(fs.GOOGLE_SEARCH_STATE["streak"], 0)

    def test_a_normal_search_run_resets_the_streak(self):
        fs.GOOGLE_SEARCH_STATE["streak"] = 2
        self.answer_with([
            ("Test Public School Udupi Karnataka info@testschool.in", [], ""),
            ("Test Public School Udupi Karnataka 011 2345 6789", [], ""),
        ])
        fs.find_official_contacts(FakeDirector(FakeBrowser([])),
                                  school(phone=""))
        self.assertEqual(fs.GOOGLE_SEARCH_STATE["streak"], 0)

    def test_the_cooldown_is_published_for_the_other_ranks(self):
        fs.PROJECT_DIR = os.path.join(self.folder, "rank_3")
        os.makedirs(fs.PROJECT_DIR, exist_ok=True)
        fs.start_google_cooldown()
        until, streak = fs.read_shared_cooldown()
        self.assertGreater(until, time.time() + 59 * 60)
        self.assertEqual(streak, 1)
        self.assertTrue(os.path.exists(fs.cooldown_share_path()))

    def test_a_rank_honours_another_rank_s_longer_cooldown(self):
        fs.PROJECT_DIR = os.path.join(self.folder, "rank_4")
        os.makedirs(fs.PROJECT_DIR, exist_ok=True)
        far = time.time() + 3 * 3600
        fs.write_shared_cooldown(far, 3)
        fs.GOOGLE_SEARCH_STATE["blocked_until"] = 0.0
        fs.GOOGLE_SEARCH_STATE["streak"] = 0
        left = fs.google_cooldown_left()
        self.assertGreater(left, 2.9 * 3600,
                           "a rank must adopt the longer shared window")
        self.assertEqual(fs.GOOGLE_SEARCH_STATE["streak"], 3,
                         "and the shared streak, so its own backoff starts "
                         "from the right place")

    def test_a_shorter_shared_cooldown_never_shortens_our_own(self):
        fs.PROJECT_DIR = os.path.join(self.folder, "rank_5")
        os.makedirs(fs.PROJECT_DIR, exist_ok=True)
        mine = time.time() + 2 * 3600
        fs.GOOGLE_SEARCH_STATE["blocked_until"] = mine
        fs.write_shared_cooldown(time.time() + 60, 1)
        self.assertGreater(fs.google_cooldown_left(), 1.9 * 3600)

    def test_an_unusable_shared_file_is_survivable(self):
        fs.PROJECT_DIR = os.path.join(self.folder, "rank_6")
        os.makedirs(fs.PROJECT_DIR, exist_ok=True)
        with open(fs.cooldown_share_path(), "w", encoding="utf-8") as handle:
            handle.write("{not json at all")
        self.assertEqual(fs.read_shared_cooldown(), (0.0, 0))
        self.assertEqual(fs.google_cooldown_left(), 0.0)


# ---------------------------------------------------------------------------
# 7. the contact retry: filling blanks in real workbooks
# ---------------------------------------------------------------------------

class ContactRetryWorkbooks(PacingBase):

    def setUp(self):
        super().setUp()
        os.makedirs(fs.EXCEL_DIR, exist_ok=True)
        self.a = school(name="Blank Contacts School", phone="")
        self.a["City"] = "Udupi"
        self.b = school(name="Full Contacts School", phone="011 1111 2222")
        self.b["City"] = "Udupi"
        self.b[fs.OFFICIAL_EMAIL_COLUMN] = "already@there.in"
        fs.write_records(fs.master_path(), [self.a, self.b])
        fs.write_records(fs.cycle_path(0), [self.a, self.b])
        fs.write_records(fs.cycle_path(1), [self.a])

    def _master(self):
        return {fs.stable_identity(r): r
                for r in fs.read_records(fs.master_path())}

    def test_a_blank_column_is_filled_in_every_workbook(self):
        update = {"identity": fs.stable_identity(self.a),
                  "school": "Blank Contacts School",
                  "email": "found@school.in", "phone": "011 3333 4444"}
        files, fields = fs.apply_contact_update(update, backups=set())
        self.assertEqual(files, 3, "master, current and cycle1 all hold it")
        self.assertEqual(fields, 6, "two blank columns in each of three files")
        got = self._master()[fs.stable_identity(self.a)]
        self.assertEqual(got[fs.OFFICIAL_EMAIL_COLUMN], "found@school.in")
        self.assertIn("3333", got[fs.OFFICIAL_PHONE_COLUMN])

    def test_an_existing_value_is_never_overwritten(self):
        update = {"identity": fs.stable_identity(self.b),
                  "school": "Full Contacts School",
                  "email": "intruder@elsewhere.in", "phone": "099 9999 9999"}
        fs.apply_contact_update(update, backups=set())
        got = self._master()[fs.stable_identity(self.b)]
        self.assertEqual(got[fs.OFFICIAL_EMAIL_COLUMN], "already@there.in")
        self.assertIn("1111", got[fs.OFFICIAL_PHONE_COLUMN])

    def test_all_57_columns_survive_the_update(self):
        update = {"identity": fs.stable_identity(self.a),
                  "school": "Blank Contacts School",
                  "email": "found@school.in", "phone": ""}
        fs.apply_contact_update(update, backups=set())
        import openpyxl
        for path in (fs.master_path(), fs.cycle_path(0), fs.cycle_path(1)):
            book = openpyxl.load_workbook(path, read_only=True)
            self.assertEqual(book.active.max_column, 57, path)
            book.close()

    def test_no_row_is_added_or_removed(self):
        before = [len(fs.read_records(p)) for p in
                  (fs.master_path(), fs.cycle_path(0), fs.cycle_path(1))]
        fs.apply_contact_update({"identity": fs.stable_identity(self.a),
                                 "school": "Blank Contacts School",
                                 "email": "found@school.in", "phone": ""},
                                backups=set())
        after = [len(fs.read_records(p)) for p in
                 (fs.master_path(), fs.cycle_path(0), fs.cycle_path(1))]
        self.assertEqual(before, after, "a retry must never add a duplicate")

    def test_applying_twice_changes_nothing_the_second_time(self):
        update = {"identity": fs.stable_identity(self.a),
                  "school": "Blank Contacts School",
                  "email": "found@school.in", "phone": "011 3333 4444"}
        first = fs.apply_contact_update(update, backups=set())
        second = fs.apply_contact_update(update, backups=set())
        self.assertGreater(first[1], 0)
        self.assertEqual(second, (0, 0),
                         "the update must be idempotent so the journal can "
                         "be replayed safely")

    def test_an_unknown_school_changes_nothing(self):
        files, fields = fs.apply_contact_update(
            {"identity": "no-such-school|nowhere", "school": "Ghost",
             "email": "x@y.z", "phone": "1"}, backups=set())
        self.assertEqual((files, fields), (0, 0))

    def test_the_journal_is_replayed_after_a_crash(self):
        update = {"identity": fs.stable_identity(self.a),
                  "school": "Blank Contacts School",
                  "email": "recovered@school.in", "phone": ""}
        fs.write_contact_journal(update)          # crash happened here
        self.assertTrue(os.path.exists(fs.contact_journal_path()))
        fs.replay_contact_journal()
        got = self._master()[fs.stable_identity(self.a)]
        self.assertEqual(got[fs.OFFICIAL_EMAIL_COLUMN], "recovered@school.in")
        self.assertFalse(os.path.exists(fs.contact_journal_path()))

    def test_an_unreadable_journal_is_removed_not_fatal(self):
        os.makedirs(fs.EXCEL_DIR, exist_ok=True)
        with open(fs.contact_journal_path(), "w", encoding="utf-8") as handle:
            handle.write("{broken")
        self.assertFalse(fs.replay_contact_journal())
        self.assertFalse(os.path.exists(fs.contact_journal_path()))


class ContactRetryPass(PacingBase):

    def setUp(self):
        super().setUp()
        os.makedirs(fs.EXCEL_DIR, exist_ok=True)
        self.rec = school(name="Owed Search School", phone="")
        self.rec["City"] = "Udupi"
        fs.write_records(fs.master_path(), [self.rec])
        self.checkpoint = fs.Checkpoint()
        self.checkpoint.note_contacts_pending("Karnataka", "Udupi", self.rec,
                                              reason="cooldown")

    def test_a_verified_contact_is_persisted_and_the_entry_dropped(self):
        self.answer_with([
            ("Owed Search School Udupi Karnataka info@owed.in", [], ""),
            ("Owed Search School Udupi Karnataka 011 5555 6666", [], ""),
        ])
        report = fs.retry_contacts_pass(FakeDirector(FakeBrowser([])),
                                        self.checkpoint)
        self.assertEqual(report["searched"], 1)
        self.assertEqual(report["remaining"], 0,
                         "the entry goes only once the result is on disk")
        saved = fs.read_records(fs.master_path())[0]
        self.assertEqual(saved[fs.OFFICIAL_EMAIL_COLUMN], "info@owed.in")

    def test_a_captcha_leaves_the_entry_pending_and_stops_the_pass(self):
        self.answer_with([("", [], "Google served a challenge")])
        report = fs.retry_contacts_pass(FakeDirector(FakeBrowser([])),
                                        self.checkpoint)
        self.assertTrue(report["blocked"])
        self.assertEqual(report["remaining"], 1,
                         "a CAPTCHA is not a result - the entry must stay")
        self.assertGreater(fs.google_cooldown_left(), 0)

    def test_nothing_is_asked_while_a_cooldown_runs(self):
        self.answer_with([("text", [], "")])
        fs.start_google_cooldown()
        report = fs.retry_contacts_pass(FakeDirector(FakeBrowser([])),
                                        self.checkpoint)
        self.assertTrue(report["blocked"])
        self.assertEqual(report["searched"], 0)
        self.assertEqual(self.queries, [])
        self.assertEqual(report["remaining"], 1)

    def test_a_school_no_longer_in_the_master_is_dropped(self):
        fs.write_records(fs.master_path(), [])
        self.answer_with([("text", [], "")])
        report = fs.retry_contacts_pass(FakeDirector(FakeBrowser([])),
                                        self.checkpoint)
        self.assertEqual(report["missing"], 1)
        self.assertEqual(report["remaining"], 0)
        self.assertEqual(self.queries, [], "no point searching for it")

    def test_a_school_with_both_contacts_already_is_dropped_unsearched(self):
        full = dict(self.rec)
        full[fs.OFFICIAL_EMAIL_COLUMN] = "a@b.c"
        full[fs.OFFICIAL_PHONE_COLUMN] = "011 2222 3333"
        fs.write_records(fs.master_path(), [full])
        self.answer_with([("text", [], "")])
        report = fs.retry_contacts_pass(FakeDirector(FakeBrowser([])),
                                        self.checkpoint)
        self.assertEqual(report["already"], 1)
        self.assertEqual(report["remaining"], 0)
        self.assertEqual(self.queries, [])

    def test_a_genuine_blank_is_a_definitive_outcome(self):
        self.answer_with([("nothing relevant here", [], ""),
                          ("nothing relevant here either", [], "")])
        report = fs.retry_contacts_pass(FakeDirector(FakeBrowser([])),
                                        self.checkpoint)
        self.assertEqual(report["no_contact"], 1)
        self.assertEqual(report["remaining"], 0,
                         "Google answered and published nothing - that is an "
                         "answer, so the entry is finished")

    def test_the_batch_limit_is_respected(self):
        """A bounded pass must leave the rest of the backlog alone.

        Every school is in the master, so each one genuinely costs a search
        from the budget - a school that is NOT in the master is dropped
        without searching and deliberately costs nothing (covered by
        test_a_school_no_longer_in_the_master_is_dropped).
        """
        everyone = [self.rec]
        for n in range(5):
            extra = school(name="Extra School %d" % n, phone="")
            extra["City"] = "Udupi"
            everyone.append(extra)
            self.checkpoint.note_contacts_pending("Karnataka", "Udupi", extra)
        fs.write_records(fs.master_path(), everyone)
        self.answer_with([("nothing relevant", [], "")] * 40)
        report = fs.retry_contacts_pass(FakeDirector(FakeBrowser([])),
                                        self.checkpoint, limit=2)
        self.assertEqual(report["searched"], 2,
                         "exactly the budget, no more")
        self.assertEqual(report["remaining"], 4,
                         "the other four stay owed for the next pass")

    def test_an_empty_backlog_is_not_an_error(self):
        self.checkpoint.data["contacts_pending"] = []
        self.checkpoint.save()
        report = fs.retry_contacts_pass(FakeDirector(FakeBrowser([])),
                                        self.checkpoint)
        self.assertEqual(report["searched"], 0)
        self.assertEqual(report["remaining"], 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
