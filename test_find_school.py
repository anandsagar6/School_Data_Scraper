#!/usr/bin/env python3
"""Automated tests for find_school.py, with mocked browser responses.

No browser, no network and no account are needed. Every test that touches
disk works in a temporary folder, so running this never changes collected
data.

    python3 -m unittest test_find_school -v
    python3 test_find_school.py
"""

import importlib.util
import datetime
import json
import logging
import os
import shutil
import sys
import tempfile
import types
import unittest
import unittest.mock

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(HERE, "find_school.py")

_spec = importlib.util.spec_from_file_location("find_school_under_test", SCRIPT)
fs = importlib.util.module_from_spec(_spec)
sys.modules["find_school_under_test"] = fs
_spec.loader.exec_module(fs)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def school(name, **extra):
    """A record as it is after collection: it has a contact number.

    Without one the script will not save it, which is the point of
    ContactRequiredTests below.
    """
    record = fs.blank_record()
    record.update({"School Name": name, "City": "Bengaluru",
                   "State": "Karnataka", "Country": "India",
                   "School Official Phone Number": "9900112233"})
    record.update(extra)
    return record


ANSWER = """School_Name: {name}
School_Short_Name: GPS
School_Website: https://gupta.example.edu.in
School_Official_Email: office@gupta.example.edu.in
School_Official_Phone_Number: 9900112233
UDISE_Code: 29200112345
Affiliation_Board: CBSE
Affiliation_Number: 830123
Principal_Name: Mrs. R. Rao
Classes_Offered: Nursery to XII
Total_Student_Strength: 1200
City: Bengaluru
State: Karnataka
PIN_Code: 560011
School_Data_Source_URL: https://gupta.example.edu.in/about
Library_Available: Yes
"""


#: The same answer with no phone number and no email: a school that is
#: fully researched and still has no way to be contacted.
NO_CONTACT_ANSWER = "\n".join(
    line for line in ANSWER.splitlines()
    if "Email" not in line and "Phone" not in line) + "\n"


class Sandbox:
    """Point every output path at a temporary folder."""

    FIELDS = ("OUTPUT_DIR", "TEMP_DIR", "EXCEL_DIR", "LOG_FILE",
              "CHECKPOINT_FILE", "BROWSER_PROFILE_DIR", "LOG_FILE",
              "CYCLE_SIZE", "SCROLL_PAUSE_SECONDS", "PAGE_SETTLE_SECONDS",
              "SCHOOL_DELAY_SECONDS",
              "CITY_DELAY_SECONDS", "GEMINI_FIRST_READ_SECONDS",
              "GEMINI_SETTLE_SECONDS", "GEMINI_POLL_SECONDS",
              "GEMINI_FIRST_READ_SECONDS", "COPY_KEY_PAUSE_SECONDS",
              "CLICK_SETTLE_SECONDS", "TYPING_SETTLE_SECONDS",
              "INPUT_BOX_POLL_SECONDS", "WINDOW_CHECK_POLL_SECONDS",
              "COPY_CLIPBOARD_PAUSE_SECONDS", "DOWN_KEY_PAUSE_SECONDS",
              "GEMINI_ANSWER_TIMEOUT", "RETRY_BACKOFF_SECONDS",
              "KEEP_RAW_GEMINI_ANSWERS",
              "KEEP_CAPTURED_TEXT", "COMPRESS_CAPTURE_FILES",
              "LOG_KEEP_DAYS",
              # the new settings, so a test never waits for real
              "CONTACT_LOOKUP_DELAY_SECONDS",
              "GEMINI_HANDOVER_DELAY_SECONDS",
              "BACKUP_DIR", "BACKUP_COMPLETED_CYCLES",
              "GIT_SYNC_ENABLED", "NIGHT_PAUSE_ENABLED",
              "WORK_PERIOD_MIN_HOURS", "WORK_PERIOD_MAX_HOURS",
              )

    def __init__(self):
        self.folder = None
        self.kept = {}

    def start(self):
        self.folder = tempfile.mkdtemp(prefix="school_test_")
        self.kept = {name: getattr(fs, name) for name in self.FIELDS}
        fs.OUTPUT_DIR = self.folder
        fs.TEMP_DIR = os.path.join(self.folder, "Temp_Data")
        fs.EXCEL_DIR = os.path.join(self.folder, "Excel_Data")
        fs.LOG_FILE = os.path.join(self.folder, "find_school.log")
        fs.BROWSER_PROFILE_DIR = os.path.join(self.folder, "profile")
        fs.CHECKPOINT_FILE = os.path.join(self.folder, "checkpoint.json")
        for name in ("SCROLL_PAUSE_SECONDS", "PAGE_SETTLE_SECONDS",
                     "GEMINI_FIRST_READ_SECONDS", "GEMINI_SETTLE_SECONDS",
                     "GEMINI_POLL_SECONDS", "RETRY_BACKOFF_SECONDS",
                     "COPY_KEY_PAUSE_SECONDS", "COPY_CLIPBOARD_PAUSE_SECONDS",
                     "DOWN_KEY_PAUSE_SECONDS", "CLICK_SETTLE_SECONDS",
                     "TYPING_SETTLE_SECONDS", "INPUT_BOX_POLL_SECONDS",
                     "WINDOW_CHECK_POLL_SECONDS"):
            setattr(fs, name, 0)
        for name in ("SCHOOL_DELAY_SECONDS", "CITY_DELAY_SECONDS",
                     "CONTACT_LOOKUP_DELAY_SECONDS",
                     "GEMINI_HANDOVER_DELAY_SECONDS"):
            setattr(fs, name, (0, 0))
        fs.BACKUP_DIR = os.path.join(self.folder, "Backup")
        # Nothing in the suite may push to a remote or sit out a night
        # stop or a scheduled break.
        fs.GIT_SYNC_ENABLED = False
        fs.NIGHT_PAUSE_ENABLED = False
        fs.WORK_PERIOD_MIN_HOURS = 0
        fs.WORK_PERIOD_MAX_HOURS = 0
        fs.GEMINI_ANSWER_TIMEOUT = 5
        fs.KEEP_RAW_GEMINI_ANSWERS = False
        # no test may start a real second browser
        fs.RUNTIME["stop"] = False
        fs.ensure_dirs()
        fs.LOG.handlers.clear()
        fs.LOG.addHandler(__import__("logging").NullHandler())
        return self

    def stop(self):
        for name, value in self.kept.items():
            setattr(fs, name, value)
        shutil.rmtree(self.folder, ignore_errors=True)


class SandboxTest(unittest.TestCase):
    def setUp(self):
        self.sandbox = Sandbox().start()

    def tearDown(self):
        self.sandbox.stop()


# ---------------------------------------------------------------------------
# mocked browser
# ---------------------------------------------------------------------------

class FakeKeyboard:
    def __init__(self, page):
        self.page = page

    def insert_text(self, text):
        self.page.typed.append(text)

    def press(self, key):
        page = self.page
        page.pressed.append(key)
        if key == "Control+v":
            page.box_text = getattr(page, "clipboard", "")
        elif key == "Control+a":
            pass
        elif key == "Delete":
            page.box_text = ""
        elif key == "Enter":
            page.submit()

    def insert_text(self, value):
        self.page.typed.append(value)
        self.page.box_text = value


class FakeFeed:
    """The Maps results list element, as the script touches it."""

    def __init__(self, page):
        self.page = page
        self.focused = 0

    def evaluate(self, script):
        if "focus()" in script:
            self.focused += 1
            return None
        if "scrollBy" in script:
            self.page.scroll_top += 400
            self.page.reveal()
            return None
        if "scrollTop" in script:
            return self.page.scroll_top
        return None


class FakeMapsPage:
    """A Google Maps results page driven by the keyboard, with a clipboard.

    Models what a live search actually does: Ctrl+A/Ctrl+C copies the text
    of everything currently in the list, the list grows as it is scrolled,
    and the end-of-list line only appears once every result is in the page.
    """

    #: What a real Maps page copies before its first result, measured from
    #: a live capture: the query, the filter buttons, the side rail.
    HEADER = ("\n{query}\n\n\n\n\nRating\n\n\n\nHours\n\n\n\nAll filters\n"
              "\n\n\n\nSaved\n\n\nRecents\n\n\nGet app\nResults\n\n\nShare\n")

    def __init__(self, cards, per_scroll=5, visible=5, end_message=True,
                 window=None, body=""):
        self.cards = cards
        self.per_scroll = per_scroll
        self.visible = min(visible, len(cards))
        self.end_message = end_message
        self.body = body
        self.url = "https://www.google.com/maps/search/test"
        self.query = "Private Schools in Bengaluru"
        self.scroll_top = 0
        self.downs = 0
        self.copies = 0
        self.closed = False
        self.clipboard = ""
        self.keyboard = FakeMapsKeyboard(self)
        self.feed = FakeFeed(self)
        self.typed, self.pressed = [], []

    # -- plumbing ----------------------------------------------------------
    def goto(self, url, **_kwargs):
        self.url = url
        import urllib.parse as _u
        tail = url.split("/maps/search/")[-1].split("?")[0]
        if tail:
            self.query = _u.unquote_plus(tail)

    def is_closed(self):
        return self.closed

    def close(self):
        self.closed = True

    def title(self):
        return "Results"

    def wait_for_selector(self, _selector, **_kwargs):
        return True

    def query_selector(self, selector):
        if "feed" in selector:
            return self.feed
        return None

    def query_selector_all(self, _selector):
        return []

    def submit(self):
        pass

    # -- the page's behaviour ----------------------------------------------
    def reveal(self):
        self.visible = min(len(self.cards), self.visible + self.per_scroll)

    def copied_text(self):
        """What Ctrl+A / Ctrl+C would put on the clipboard right now."""
        parts = [self.HEADER.replace("{query}", self.query)]
        for card in self.cards[:self.visible]:
            parts.append(card["copied"])
        if self.end_message and self.visible >= len(self.cards):
            parts.append("\nYou've reached the end of the list.\n")
        return "\n".join(parts)

    def inner_text(self, selector="body"):
        if selector == "body":
            return self.body or "Results"
        return self.copied_text()

    def evaluate(self, script, *args):
        if "clipboard.readText" in script:
            return self.clipboard
        if "getSelection" in script:
            return self.clipboard
        if "clipboard.writeText" in script:
            self.clipboard = args[0] if args else ""
            return None
        return None


class FakeMapsKeyboard:
    def __init__(self, page):
        self.page = page

    def insert_text(self, value):
        self.page.typed.append(value)

    def press(self, key):
        page = self.page
        page.pressed.append(key)
        if key == "Control+a":
            page.selected = True
        elif key == "Control+c":
            page.clipboard = page.copied_text()
            page.copies += 1
        elif key == "ArrowDown":
            page.downs += 1
            page.scroll_top += 40
            if page.downs % 12 == 0:
                page.reveal()


class FakeDirector:
    """Hands out one fake browser at a time and records the order.

    ``max_open`` is the point of it: it must never reach 2.
    """

    def __init__(self, maps_pages=None, gemini_page=None, show=False):
        self._maps_pages = list(maps_pages or [FakeMapsPage([])])
        self._gemini_page = gemini_page or FakeGeminiPage("")
        self.show = show
        self.display = ":0" if show else ""
        self.authority = "/tmp/auth"
        self.kind = ""
        self.events = []
        self.opened = {"maps": 0, "gemini": 0, "search": 0}
        self.on_close = {}
        self.max_open = 0
        self._live = 0
        self._browser = None

    def open(self, kind):
        if self.kind == kind and self._browser is not None:
            return self._browser
        self.close()
        if kind == "maps":
            page = self._maps_pages[min(self.opened["maps"],
                                        len(self._maps_pages) - 1)]
            self._browser = FakeBrowser([page])
        else:
            self._browser = FakeBrowser([self._gemini_page,
                                         self._gemini_page])
        self.kind = kind
        self.opened[kind] += 1
        self._live = 1
        self.max_open = max(self.max_open, self._live)
        self.events.append(("open", kind))
        return self._browser

    def maps(self):
        return self.open("maps")

    def search(self):
        """The official-contact searches ask for this one."""
        return self.open("search")

    def gemini(self):
        return self.open("gemini")

    def close(self):
        if not self.kind:
            return True, "nothing was open"
        callback = self.on_close.get(self.kind)
        if callback is not None:
            callback()
        self.events.append(("close", self.kind))
        self.kind, self._browser, self._live = "", None, 0
        return True, "closed"

    def shutdown(self):
        self.close()

    def windows_open(self):
        return self._live


def maps_entry(number, name=None, pin="560011", street=None, phone=True):
    """One result as it appears in the COPIED text."""
    name = name or "School %d" % number
    street = street or "%d MG Road" % number
    hours = "Open \u00b7 Closes 6 pm"
    if phone:
        hours += " \u00b7 99001%05d" % number
    return {
        "name": name,
        "copied": "\n".join([name, name, "4.5(120)",
                             "Private school \u00b7  \u00b7 %s, Bengaluru %s"
                             % (street, pin), hours, "Website", "",
                             "Directions"]),
    }


class FakeGeminiPage:
    """A Gemini tab whose answer arrives after a few reads."""

    def __init__(self, answer, blank_reads=1, has_input=True, body="New chat"):
        self.answer = answer
        self.blank_reads = blank_reads
        self.has_input = has_input
        self.body = body
        self.url = fs.GEMINI_URL
        self.closed = False
        self.keyboard = FakeKeyboard(self)
        self.typed, self.pressed = [], []
        self.reads = 0
        self.submitted = False
        self.asked = []
        self.clipboard = ""
        self.box_text = ""

    def goto(self, url, **_kwargs):
        self.url = url
        self.reads = 0
        self.submitted = False

    def is_closed(self):
        return self.closed

    def close(self):
        self.closed = True

    def title(self):
        return "Gemini"

    def wait_for_selector(self, _selector, **_kwargs):
        return True

    def inner_text(self, _selector="body"):
        return self.body

    def evaluate(self, script, *args):
        if "clipboard.writeText" in script:
            self.clipboard = args[0] if args else ""
            return None
        if "clipboard.readText" in script:
            return self.clipboard
        return None

    def submit(self):
        self.submitted = True
        self.asked.append(self.typed[-1] if self.typed else "")

    def query_selector(self, selector):
        if self.has_input and ("contenteditable" in selector
                               or selector == "textarea"):
            return FakeInputBox(self)
        return None

    def query_selector_all(self, selector):
        if selector != "model-response" or not self.submitted:
            return []
        self.reads += 1
        if self.reads <= self.blank_reads:
            return [FakeNode(self, "Gemini said")]
        return [FakeNode(self, self.answer)]


class FakeInputBox:
    """The Gemini chat box: Ctrl+V fills it from the page's clipboard."""

    def __init__(self, page):
        self.page = page

    def click(self):
        pass

    def inner_text(self):
        return self.page.box_text


class FakeNode:
    def __init__(self, page, text=""):
        self.page = page
        self.text = text

    def click(self):
        pass

    def inner_text(self):
        return self.text


class FakeBrowser:
    """Hands out the pages the test prepared, in order."""

    def __init__(self, pages):
        self.pages = list(pages)
        self.handed = []
        self.restarts = 0
        self.spare_closed = 0
        self.brought_to_front = 0
        self.display = ""
        self.authority = ""

    def start(self):
        pass

    def close(self):
        pass

    def alive(self):
        return True

    def restart(self):
        self.restarts += 1

    def page(self):
        return self.pages[0]

    def new_page(self):
        nxt = self.pages[min(1, len(self.pages) - 1)]
        self.handed.append(nxt)
        return nxt

    def playwright(self):
        return "shared-driver"

    def close_spare_tabs(self):
        self.spare_closed += 1
        return 0

    def show_tab(self):
        self.brought_to_front += 1

    def tab_titles(self):
        return [getattr(p, "url", "?") for p in self.pages]

    def goto(self, page, url, wait="domcontentloaded"):
        page.goto(url)
        return page

    @staticmethod
    def text_of(page, selector="body"):
        try:
            return page.inner_text(selector) or ""
        except Exception:                               # noqa: BLE001
            return ""


# ===========================================================================
# 1. the embedded city dataset
# ===========================================================================

class CityDatasetTests(unittest.TestCase):
    def test_it_matches_mylib_in_order_and_content(self):
        """What the script loaded equals city_by_states in mylib.py.

        The mylib.py BESIDE THE SCRIPT is the only source, so that is what
        this compares against - not another project's copy.
        """
        import ast
        reference = os.path.join(os.path.dirname(SCRIPT), "mylib.py")
        if not os.path.exists(reference):
            self.skipTest("mylib.py is not beside the script")
        with open(reference, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        source = None
        for node in tree.body:
            if isinstance(node, ast.Assign) and any(
                    getattr(t, "id", "") == "city_by_states"
                    for t in node.targets):
                source = ast.literal_eval(node.value)
        self.assertIsNotNone(source)
        self.assertEqual(list(fs.CITY_BY_STATES), list(source),
                         "the group order differs")
        for group in source:
            self.assertEqual(fs.CITY_BY_STATES[group], source[group],
                             "the cities of %s differ" % group)

    def test_the_counts_are_the_expected_ones(self):
        self.assertEqual(len(fs.CITY_BY_STATES), 29)
        self.assertEqual(sum(len(v) for v in fs.CITY_BY_STATES.values()), 905)
        self.assertEqual(len(fs.ALL_CITIES), 866)

    def test_a_repeated_city_keeps_its_first_group(self):
        flat = [(g, c) for g, names in fs.CITY_BY_STATES.items()
                for c in names]
        first = {}
        for group, city in flat:
            first.setdefault(city, group)
        for city, group in first.items():
            self.assertEqual(fs.CITY_TO_STATE[city], group, city)

    def test_duplicates_were_removed_the_dict_fromkeys_way(self):
        flat = [c for names in fs.CITY_BY_STATES.values()
                for c in names]
        self.assertEqual(fs.ALL_CITIES, list(dict.fromkeys(flat)))


# ===========================================================================
# 2. the rank contract
# ===========================================================================

class RankTests(unittest.TestCase):
    def test_the_function_is_the_github_extractors(self):
        """Same code, verified against the reference file where it exists."""
        import ast
        reference = "/home/anand/github_profile_extractor/student_scraper.py"
        if not os.path.exists(reference):
            self.skipTest("student_scraper.py is not on this machine")
        with open(reference, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        theirs = None
        for node in ast.walk(tree):
            if isinstance(node, ast.FunctionDef) and \
                    node.name == "get_cities_by_rank":
                theirs = node
        self.assertIsNotNone(theirs)
        mine = None
        with open(SCRIPT, encoding="utf-8") as handle:
            for node in ast.walk(ast.parse(handle.read())):
                if isinstance(node, ast.FunctionDef) and \
                        node.name == "get_cities_by_rank":
                    mine = node
        self.assertIsNotNone(mine)
        strip = lambda fn: [ast.dump(n) for n in fn.body
                            if not isinstance(n, ast.Expr)
                            or not isinstance(n.value, ast.Constant)]
        self.assertEqual(strip(mine), strip(theirs),
                         "the rank algorithm is not the reference one")

    def test_every_city_is_assigned_exactly_once(self):
        for systems in (1, 2, 3, 4, 7, 10, 25, 50, 100, 866):
            rejoined = []
            for this_rank in range(1, systems + 1):
                rejoined.extend(fs.get_cities_by_rank(fs.ALL_CITIES,
                                                      this_rank, systems))
            self.assertEqual(rejoined, fs.ALL_CITIES, "num_sys=%d" % systems)

    def test_num_sys_25_covers_every_city(self):
        total = sum(len(fs.get_cities_by_rank(fs.ALL_CITIES, r, 25))
                    for r in range(1, 26))
        self.assertEqual(total, len(fs.ALL_CITIES))
        self.assertEqual(total, 866)

    def test_the_slices_are_contiguous(self):
        at = 0
        for this_rank in range(1, 26):
            mine = fs.get_cities_by_rank(fs.ALL_CITIES, this_rank, 25)
            self.assertEqual(mine, fs.ALL_CITIES[at:at + len(mine)])
            at += len(mine)

    def test_sizes_differ_by_at_most_one(self):
        sizes = [len(fs.get_cities_by_rank(fs.ALL_CITIES, r, 25))
                 for r in range(1, 26)]
        self.assertLessEqual(max(sizes) - min(sizes), 1)

    def test_an_invalid_rank_fails_clearly(self):
        for bad_rank, bad_num in ((0, 25), (26, 25), (-3, 25), (1, 0), (1, -1)):
            with self.assertRaises(fs.RankError):
                fs.validate_rank(bad_rank, bad_num)

    def test_a_rank_runs_only_its_own_cities(self):
        mine = {c for _s, c in fs.cities_for_rank(3, 25)}
        others = set()
        for this_rank in range(1, 26):
            if this_rank != 3:
                others |= {c for _s, c in fs.cities_for_rank(this_rank, 25)}
        self.assertFalse(mine & others)

    def test_rank_json_and_the_environment_override_the_defaults(self):
        folder = tempfile.mkdtemp(prefix="school_rank_")
        kept_file, kept_env = fs.RANK_FILE, dict(os.environ)
        try:
            fs.RANK_FILE = os.path.join(folder, "rank.json")
            self.assertEqual(fs.resolve_rank()[:2], (fs.rank, fs.num_sys))
            with open(fs.RANK_FILE, "w", encoding="utf-8") as handle:
                json.dump({"rank": 7, "num_sys": 25}, handle)
            self.assertEqual(fs.resolve_rank()[:2], (7, 25))
            os.environ["SCHOOL_RANK"] = "9"
            self.assertEqual(fs.resolve_rank()[:2], (9, 25))
            self.assertEqual(fs.resolve_rank(cli_rank=4)[:2], (4, 25))
        finally:
            fs.RANK_FILE = kept_file
            os.environ.clear()
            os.environ.update(kept_env)
            shutil.rmtree(folder, ignore_errors=True)

    def test_a_missing_rank_json_is_not_required(self):
        kept = fs.RANK_FILE
        try:
            fs.RANK_FILE = "/nonexistent/rank.json"
            value_rank, value_num, source = fs.resolve_rank()
            self.assertEqual((value_rank, value_num), (fs.rank, fs.num_sys))
            self.assertIn("top of this file", source)
        finally:
            fs.RANK_FILE = kept


# ===========================================================================
# 3. keywords and the search plan
# ===========================================================================

class KeywordTests(unittest.TestCase):
    def test_the_query_and_url_are_the_intended_ones(self):
        self.assertEqual(fs.build_query("Private Schools", "Bengaluru"),
                         "Private Schools in Bengaluru")
        self.assertTrue(fs.maps_search_url("private schools in bangalore")
                        .startswith("https://www.google.com/maps/search/"
                                    "private+schools+in+bangalore"))

    def test_the_plan_is_cities_then_keywords_in_order(self):
        plan = fs.search_plan(1, 866)          # exactly one city
        self.assertEqual(len({(s, c) for s, c, _k in plan}), 1)
        self.assertEqual([k for _s, _c, k in plan], fs.enabled_keywords())

    def test_the_keywords_come_from_mylib(self):
        import ast
        reference = fs.mylib_path()
        if not reference or not os.path.exists(reference):
            self.skipTest("mylib.py is not on this machine")
        with open(reference, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        theirs = None
        for node in tree.body:
            if isinstance(node, ast.Assign) and any(
                    getattr(t, "id", "") == fs.MYLIB_KEYWORD_NAME
                    for t in node.targets):
                theirs = ast.literal_eval(node.value)
        self.assertIsNotNone(theirs, "mylib.py has no keyword list")
        self.assertEqual(list(fs.SCHOOL_KEYWORDS), list(theirs))
        self.assertIn("mylib.py", fs.SHARED_SOURCE["keywords"])

    def test_the_cities_come_from_mylib(self):
        import ast
        reference = fs.mylib_path()
        if not reference or not os.path.exists(reference):
            self.skipTest("mylib.py is not on this machine")
        with open(reference, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        theirs = None
        for node in tree.body:
            if isinstance(node, ast.Assign) and any(
                    getattr(t, "id", "") == fs.MYLIB_CITY_NAME
                    for t in node.targets):
                theirs = ast.literal_eval(node.value)
        self.assertIsNotNone(theirs)
        self.assertEqual(list(fs.CITY_BY_STATES), list(theirs))
        self.assertIn("mylib.py", fs.SHARED_SOURCE["cities"])

    def test_a_missing_mylib_is_reported_not_guessed_around(self):
        """There is no second copy of the lists, so it must say so."""
        kept = os.environ.get("SCHOOL_MYLIB")
        try:
            os.environ["SCHOOL_MYLIB"] = "/nonexistent/mylib.py"
            cities, keywords, fields = fs.load_shared_lists()
            self.assertEqual(cities, {})
            self.assertEqual(keywords, [])
            self.assertEqual(fields, ())
            self.assertIn("not found", fs.SHARED_SOURCE["cities"])
        finally:
            if kept is None:
                os.environ.pop("SCHOOL_MYLIB", None)
            else:
                os.environ["SCHOOL_MYLIB"] = kept
            fs.load_shared_lists()          # put the real source back

    def test_the_message_says_how_to_fix_it(self):
        kept_cities = fs.CITY_BY_STATES
        try:
            fs.CITY_BY_STATES = {}
            problem = fs.shared_lists_problem()
            self.assertIn("mylib.py", problem)
            self.assertIn("SCHOOL_MYLIB", problem)
        finally:
            fs.CITY_BY_STATES = kept_cities
        self.assertEqual(fs.shared_lists_problem(), "",
                         "a healthy load still reported a problem")

    def test_only_the_named_lists_are_asked_for(self):
        """mylib.py is parsed, never imported, and only three names read."""
        import ast
        with open(SCRIPT, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        asked = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Call) and \
                    getattr(node.func, "id", "") == "read_list_from_file":
                self.assertEqual(len(node.args), 2)
                name = node.args[1]
                self.assertIsInstance(
                    name, ast.Name,
                    "a literal name is read out of mylib.py; it should be "
                    "one of the two settings")
                asked.add(name.id)
        self.assertEqual(asked, {"MYLIB_CITY_NAME", "MYLIB_KEYWORD_NAME",
                                 "MYLIB_FIELD_NAME"})

    def test_mylib_is_never_imported(self):
        import ast
        with open(SCRIPT, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        for node in ast.walk(tree):
            if isinstance(node, (ast.Import, ast.ImportFrom)):
                names = [a.name for a in getattr(node, "names", [])]
                names.append(getattr(node, "module", "") or "")
                for name in names:
                    self.assertNotIn("mylib", name)

    def test_a_broken_mylib_cannot_stop_the_script(self):
        folder = tempfile.mkdtemp(prefix="school_mylib_")
        try:
            broken = os.path.join(folder, "mylib.py")
            with open(broken, "w", encoding="utf-8") as handle:
                handle.write("city_by_states = { this is not python")
            self.assertIsNone(fs.read_list_from_file(broken,
                                                     fs.MYLIB_CITY_NAME))
        finally:
            shutil.rmtree(folder, ignore_errors=True)

    def test_mylib_holds_no_credentials_at_all(self):
        """mylib.py is committed to Git, so it must carry no secret.

        It used to: a `file_langs` dictionary of mail accounts with app
        passwords, copied in with the city data. Those are gone, and this
        is what keeps them gone - checking the file itself rather than
        only checking that they did not reach find_school.py, which would
        pass trivially once the file is clean.
        """
        import ast
        import re as _re
        reference = os.path.join(os.path.dirname(SCRIPT), "mylib.py")
        if not os.path.exists(reference):
            self.skipTest("mylib.py is not beside the script")
        with open(reference, encoding="utf-8") as handle:
            body = handle.read()
        for node in ast.walk(ast.parse(body)):
            if not isinstance(node, ast.Dict):
                continue
            for key in node.keys:
                if isinstance(key, ast.Constant):
                    self.assertNotIn(
                        str(key.value).lower(),
                        ("password", "passwd", "email", "token", "secret",
                         "api_key", "apikey"),
                        "a credential-shaped key is back in mylib.py")
        self.assertEqual(
            _re.findall(r"[\w.+-]+@[\w-]+\.[a-z]{2,}", body), [],
            "an email address is back in mylib.py")
        self.assertEqual(
            [n.name for n in ast.parse(body).body
             if isinstance(n, (ast.Import, ast.ImportFrom))
             and hasattr(n, "name")], [],
            "mylib.py should be data only")

    def test_no_secret_reaches_the_script(self):
        """Whatever mylib.py holds, none of it may land in find_school.py."""
        import ast
        reference = os.path.join(os.path.dirname(SCRIPT), "mylib.py")
        if not os.path.exists(reference):
            self.skipTest("mylib.py is not beside the script")
        with open(reference, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        secrets = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Dict):
                for key, value in zip(node.keys, node.values):
                    if isinstance(key, ast.Constant) and key.value in (
                            "password", "email") and isinstance(
                                value, ast.Constant):
                        secrets.add(value.value)
        with open(SCRIPT, encoding="utf-8") as handle:
            source = handle.read()
        if secrets:
            # mylib.py carries credentials again: none may be copied here.
            for secret in secrets:
                self.assertNotIn(secret, source,
                                 "a credential was copied across")
            return
        # mylib.py is clean, which is the state we want - so this instead
        # asserts the script keeps no credential of its own. Without this
        # branch the test would pass while checking nothing.
        for node in ast.walk(ast.parse(source)):
            if not isinstance(node, ast.Dict):
                continue
            for key in node.keys:
                if isinstance(key, ast.Constant):
                    self.assertNotIn(
                        str(key.value).lower(),
                        ("password", "passwd", "token", "secret", "apikey"),
                        "find_school.py holds a credential-shaped entry")

    def test_each_laptop_chooses_which_keywords_to_run(self):
        kept = list(fs.ENABLED_SCHOOL_KEYWORDS)
        try:
            fs.ENABLED_SCHOOL_KEYWORDS = ["Government Schools"]
            self.assertEqual(fs.enabled_keywords(),
                         ["Government Schools"])
            fs.ENABLED_SCHOOL_KEYWORDS = []
            self.assertEqual(fs.enabled_keywords(), list(fs.SCHOOL_KEYWORDS))
        finally:
            fs.ENABLED_SCHOOL_KEYWORDS = kept

    def test_a_misspelled_enabled_keyword_is_reported(self):
        kept = list(fs.ENABLED_SCHOOL_KEYWORDS)
        try:
            fs.ENABLED_SCHOOL_KEYWORDS = ["Private Schools", "Privat Schols"]
            self.assertEqual(fs.unknown_enabled_keywords(), ["Privat Schols"])
            self.assertEqual(fs.enabled_keywords(), ["Private Schools"])
        finally:
            fs.ENABLED_SCHOOL_KEYWORDS = kept


# ===========================================================================
# 4. the 57 fields
# ===========================================================================

class FieldTests(unittest.TestCase):
    def test_there_are_exactly_57_of_everything(self):
        self.assertEqual(len(fs.SCHOOL_COLUMNS), 57)
        self.assertEqual(len(fs.ANSWER_KEYS), 57)
        self.assertEqual(len(set(fs.SCHOOL_COLUMNS)), 57)
        self.assertEqual(len(fs.blank_record()), 57)

    def test_every_answer_key_maps_to_its_own_column(self):
        for column, key in fs.SCHOOL_FIELDS:
            self.assertEqual(fs.KEY_TO_COLUMN[key], column)
            self.assertEqual(fs.COLUMN_TO_KEY[column], key)

    def test_each_key_lands_in_its_own_column_and_no_other(self):
        for column, key in fs.SCHOOL_FIELDS:
            parsed = fs.parse_answer("%s: marker-value" % key)
            self.assertEqual(parsed, {column: "marker-value"},
                             "%s leaked out of %s" % (key, column))

    def test_source_urls_is_not_a_58th_column(self):
        self.assertNotIn("Source URLs", fs.SCHOOL_COLUMNS)
        self.assertEqual(
            fs.parse_answer("Source_URLs: https://a.example.com"),
            {"School Data Source URL": "https://a.example.com"})

    def test_the_excel_header_is_the_field_order(self):
        book = fs.new_workbook()
        header = [cell.value for cell in book.active[1]]
        self.assertEqual(header, list(fs.SCHOOL_COLUMNS))


# ===========================================================================
# 5. the parser
# ===========================================================================

class ParserTests(unittest.TestCase):
    def test_no_data_is_stored_as_blank(self):
        parsed = fs.parse_answer("School_Name: ABC\nUDISE_Code: NO_DATA\n")
        self.assertEqual(parsed.get("School Name"), "ABC")
        self.assertNotIn("UDISE Code", parsed)

    def test_extra_spaces_and_blank_lines_are_tolerated(self):
        messy = "\n\n  School_Name :   ABC School  \n\n\nAffiliation_Board:CBSE\n"
        parsed = fs.parse_answer(messy)
        self.assertEqual(parsed.get("School Name"), "ABC School")
        self.assertEqual(parsed.get("Affiliation Board"), "CBSE")

    def test_two_fields_on_one_line_are_separated(self):
        parsed = fs.parse_answer("School_Name: ABC SchoolPrincipal_Name: R. Rao")
        self.assertEqual(parsed.get("School Name"), "ABC School")
        self.assertEqual(parsed.get("Principal Name"), "R. Rao")

    def test_a_multiline_value_survives(self):
        parsed = fs.parse_answer("Additional_School_Notes: one\ntwo\n"
                                 "School_Name: ABC\n")
        self.assertEqual(parsed["Additional School Notes"], "one two")
        self.assertEqual(parsed["School Name"], "ABC")

    def test_semicolon_separated_values_are_kept_whole(self):
        parsed = fs.parse_answer("Available_Streams: Science; Commerce; Arts")
        self.assertEqual(parsed["Available Streams (Science / Commerce / Arts)"],
                         "Science; Commerce; Arts")

    def test_unrelated_text_is_ignored(self):
        self.assertEqual(fs.parse_answer("Sure, here is what I found!"), {})
        self.assertFalse(fs.looks_like_an_answer("Gemini said"))

    def test_the_echoed_prompt_never_becomes_data(self):
        echoed = fs.answer_format_text() + "\nSchool_Name: Real School\n"
        self.assertEqual(fs.parse_answer(echoed).get("School Name"),
                         "Real School")
        self.assertEqual(fs.parse_answer("School_Name: Value"), {})

    def test_a_malformed_answer_is_rejected(self):
        self.assertFalse(fs.looks_like_an_answer("School_Name: ABC"))
        self.assertTrue(fs.looks_like_an_answer(ANSWER.format(name="ABC")))


# ===========================================================================
# 6. duplicate rules
# ===========================================================================

class DuplicateTests(unittest.TestCase):
    def test_the_same_campus_is_one_row(self):
        index = fs.SchoolIndex()
        index.add(school("Gupta Public School",
                         **{"School Address": "12 MG Road, Bengaluru",
                            "PIN Code": "560001"}))
        index.add(school("Gupta Public School",
                         **{"School Address": "12 MG Road, Bengaluru",
                            "PIN Code": "560001",
                            "Principal Name": "R. Rao"}))
        self.assertEqual(len(index), 1)
        self.assertEqual(index.records[0]["Principal Name"], "R. Rao")

    def test_two_branches_of_one_chain_stay_apart(self):
        index = fs.SchoolIndex()
        for area, pin in (("Sarjapur Road", "560035"), ("Hennur Road", "560043")):
            index.add(school("ORCHIDS The International School - %s" % area,
                             **{"School Website": "https://orchids.edu.in",
                                "PIN Code": pin,
                                "School Address": "%s, Bengaluru" % area}))
        self.assertEqual(len(index), 2)

    def test_a_similar_name_alone_never_merges(self):
        index = fs.SchoolIndex()
        index.add(school("St Mary's School",
                         **{"School Address": "5 Church Street"}))
        index.add(school("St Mary's School",
                         **{"School Address": "88 Hosur Road"}))
        self.assertEqual(len(index), 2)

    def test_the_same_maps_place_is_one_school(self):
        index = fs.SchoolIndex()
        link = ("https://www.google.com/maps/place/X/data=!1s0x3bae16:0x7f00a1")
        index.add(school("Alpha School", **{"School Google Maps Link": link}))
        index.add(school("Alpha School - Main",
                         **{"School Google Maps Link": link}))
        self.assertEqual(len(index), 1)

    def test_the_most_complete_information_is_kept(self):
        index = fs.SchoolIndex()
        index.add(school("Beta School", **{"PIN Code": "560001",
                                           "School Address": "1 A Road"}))
        index.add(school("Beta School", **{"PIN Code": "560001",
                                           "School Address": "1 A Road",
                                           "School Website": "https://b.in",
                                           "Library Available": "Yes"}))
        self.assertEqual(len(index), 1)
        self.assertEqual(index.records[0]["School Website"], "https://b.in")
        self.assertEqual(index.records[0]["Library Available"], "Yes")


# ===========================================================================
# 7. Maps discovery, with a mocked page
# ===========================================================================

class DiscoveryTests(SandboxTest):
    def test_it_collects_every_card_not_just_the_first_screen(self):
        cards = [maps_entry(n) for n in range(1, 61)]
        page = FakeMapsPage(cards, per_scroll=5, visible=5)
        browser = FakeBrowser([page])
        result = fs.collect_search(browser, "Private Schools", "Karnataka",
                                   "Bengaluru")
        self.assertEqual(len(result.records), 60)
        self.assertTrue(result.complete)
        self.assertIn("end-of-list message", result.reason)

    def test_a_virtualised_feed_does_not_lose_the_cards_that_scrolled_away(self):
        """The real failure: only 20 cards are ever rendered at once."""
        cards = [maps_entry(n) for n in range(1, 101)]
        page = FakeMapsPage(cards, per_scroll=6, visible=6, window=20)
        browser = FakeBrowser([page])
        result = fs.collect_search(browser, "Private Schools", "Karnataka",
                                   "Bengaluru")
        self.assertEqual(len(result.records), 100)

    def test_a_search_with_no_end_message_is_reported_incomplete(self):
        cards = [maps_entry(n) for n in range(1, 21)]
        page = FakeMapsPage(cards, per_scroll=5, visible=5, end_message=False)
        browser = FakeBrowser([page])
        result = fs.collect_search(browser, "Private Schools", "Karnataka",
                                   "Bengaluru")
        self.assertFalse(result.complete)
        self.assertIn("no new school", result.reason)
        self.assertEqual(len(result.records), 20)

    def test_a_page_that_never_loaded_is_not_an_empty_city(self):
        class DeadPage(FakeMapsPage):
            def wait_for_selector(self, _selector, **_kwargs):
                raise fs.PlaywrightTimeout("no feed")

        page = DeadPage([], end_message=False)
        page.url = "https://www.google.com/maps/search/test"
        result = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                   "Karnataka", "Bengaluru")
        self.assertFalse(result.complete)
        self.assertIn("did not load", result.reason)

    def test_a_bot_challenge_is_reported_not_worked_around(self):
        page = FakeMapsPage([maps_entry(1)],
                            body="Our systems have detected unusual traffic")
        with self.assertRaises(fs.AccessBlocked):
            fs.collect_search(FakeBrowser([page]), "Private Schools",
                              "Karnataka", "Bengaluru")

    def test_the_limit_stops_a_search_early(self):
        cards = [maps_entry(n) for n in range(1, 61)]
        page = FakeMapsPage(cards, per_scroll=5, visible=5)
        result = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                   "Karnataka", "Bengaluru", limit=7)
        self.assertEqual(len(result.records), 7)
        self.assertFalse(result.complete)

    def test_a_card_becomes_a_record_without_inventing_anything(self):
        page = FakeMapsPage([maps_entry(1, name="Gupta Public School")])
        result = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                   "Karnataka", "Bengaluru")
        record = result.records[0]
        self.assertEqual(record["School Name"], "Gupta Public School")
        self.assertEqual(record["City"], "Bengaluru")
        self.assertEqual(record["State"], "Karnataka")
        self.assertEqual(record["PIN Code"], "560011")
        # the copied text carries no links, so this stays empty rather than
        # being invented
        self.assertEqual(record["School Google Maps Link"], "")
        self.assertEqual(record["UDISE Code"], "")
        self.assertEqual(record["Principal Name"], "")

    def test_an_affiliation_number_in_the_name_is_not_read_as_a_pin(self):
        """A live run hit this: the wrong PIN then rejected a right answer."""
        card = maps_entry(1,
                         name="ORCHIDS The International School in Horamavu "
                              "- CBSE - Affiliation No. 831330",
                         pin="560043")
        page = FakeMapsPage([card])
        result = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                   "Karnataka", "Bengaluru")
        self.assertEqual(result.records[0]["PIN Code"], "560043")

    def test_the_keyword_never_becomes_the_address(self):
        page = FakeMapsPage([maps_entry(1)])
        result = fs.collect_search(FakeBrowser([page]), "CBSE Schools",
                                   "Karnataka", "Bengaluru")
        for value in result.records[0].values():
            self.assertNotIn("CBSE Schools", str(value))


# ===========================================================================
# 8. Gemini, with a mocked page
# ===========================================================================

class GeminiTests(SandboxTest):
    def test_an_answer_is_parsed_into_the_right_columns(self):
        page = FakeGeminiPage(ANSWER.format(name="Gupta Public School"))
        gemini = fs.Gemini(FakeBrowser([page, page]), separate=False)
        record = school("Gupta Public School", **{"PIN Code": "560011"})
        values, raw = gemini.research(record)
        self.assertEqual(values["Principal Name"], "Mrs. R. Rao")
        self.assertEqual(values["Affiliation Board"], "CBSE")
        self.assertIn("School_Name", raw)

    def test_it_waits_for_the_answer_instead_of_a_fixed_delay(self):
        page = FakeGeminiPage(ANSWER.format(name="Gupta Public School"),
                              blank_reads=4)
        gemini = fs.Gemini(FakeBrowser([page, page]), separate=False)
        values, _raw = gemini.research(school("Gupta Public School"))
        self.assertTrue(values)
        self.assertGreater(page.reads, 4)

    def test_an_answer_about_another_school_is_refused(self):
        page = FakeGeminiPage(ANSWER.format(name="Delhi Model Academy"))
        gemini = fs.Gemini(FakeBrowser([page, page]), separate=False)
        with self.assertRaises(fs.GeminiUnavailable):
            gemini.research(school("Gupta Public School"))

    def test_a_sign_in_page_is_reported_not_bypassed(self):
        page = FakeGeminiPage("", has_input=False,
                              body="Sign in to continue to Gemini")
        gemini = fs.Gemini(FakeBrowser([page, page]), separate=False)
        ready, why = gemini.check_ready()
        self.assertFalse(ready)
        self.assertIn("sign-in", why.lower())
        self.assertIn("does not attempt", why)

    def test_only_the_current_response_is_read(self):
        """Not the whole page: the prompt and the interface are not data."""
        page = FakeGeminiPage(ANSWER.format(name="Gupta Public School"))
        page.body = ("Gemini\nNew chat\nSchool_Name: WRONG SCHOOL FROM THE "
                     "PAGE BODY\n")
        gemini = fs.Gemini(FakeBrowser([page, page]), separate=False)
        values, _raw = gemini.research(school("Gupta Public School"))
        self.assertEqual(values["School Name"], "Gupta Public School")

    def test_maps_evidence_is_not_replaced_by_the_answer(self):
        record = school("Gupta Public School",
                        **{"School Official Phone Number": "9876543210",
                           "School Address": "12 MG Road"})
        values = fs.parse_answer(ANSWER.format(name="Gupta Public School"))
        fs.apply_answer(record, values)
        self.assertEqual(record["School Official Phone Number"], "9876543210")
        self.assertEqual(record["Principal Name"], "Mrs. R. Rao")

    def test_the_prompt_carries_the_known_values(self):
        record = school("Gupta Public School",
                        **{"School Address": "12 MG Road",
                           "School Website": "https://gupta.example.edu.in"})
        prompt = fs.build_prompt(record)
        self.assertIn("Gupta Public School", prompt)
        self.assertIn("12 MG Road", prompt)
        self.assertIn("https://gupta.example.edu.in", prompt)
        for key in fs.ANSWER_KEYS:
            self.assertIn("%s: Value" % key, prompt)


# ===========================================================================
# 9. Excel
# ===========================================================================

class ExcelTests(SandboxTest):
    def test_rows_are_written_and_read_back(self):
        written, skipped = fs.append_records(
            [school("School %d" % n, **{"PIN Code": "56000%d" % n,
                                        "School Address": "%d Road" % n})
             for n in range(1, 5)])
        self.assertEqual((written, skipped), (4, 0))
        rows = [r for p in fs.all_cycle_paths() for r in fs.read_records(p)]
        self.assertEqual(len(rows), 4)
        self.assertEqual(list(rows[0].keys()), list(fs.SCHOOL_COLUMNS))

    def test_a_finished_cycle_holds_exactly_the_cycle_size(self):
        fs.CYCLE_SIZE = 3
        fs.append_records([school("School %d" % n,
                                  **{"PIN Code": "5600%02d" % n,
                                     "School Address": "%d Road" % n})
                           for n in range(1, 11)])
        self.assertEqual(fs.finished_cycle_numbers(), [1, 2, 3])
        for number in (1, 2, 3):
            self.assertEqual(fs.row_count(fs.cycle_path(number)), 3,
                             "cycle%d" % number)
        self.assertEqual(fs.row_count(fs.cycle_path(0)), 1)

    def test_the_cycle_files_are_numbered_in_order(self):
        fs.CYCLE_SIZE = 2
        names = []
        for n in range(1, 8):
            fs.append_records([school("School %d" % n,
                                      **{"PIN Code": "5600%02d" % n,
                                         "School Address": "%d Road" % n})])
            names = [os.path.basename(p) for p in fs.all_cycle_paths()]
        self.assertEqual(names, ["current_data_cycle1.xlsx",
                                 "current_data_cycle2.xlsx",
                                 "current_data_cycle3.xlsx",
                                 "current_data.xlsx"])

    def test_a_finished_cycle_is_never_rewritten(self):
        fs.CYCLE_SIZE = 3
        fs.append_records([school("School %d" % n,
                                  **{"PIN Code": "5600%02d" % n,
                                     "School Address": "%d Road" % n})
                           for n in range(1, 8)])
        before = {os.path.basename(p): fs.row_count(p)
                  for p in fs.all_cycle_paths() if "cycle" in p}
        fs.append_records([school("Later School",
                                  **{"PIN Code": "560099",
                                     "School Address": "99 Road"})])
        after = {os.path.basename(p): fs.row_count(p)
                 for p in fs.all_cycle_paths() if "cycle" in p}
        for name, rows in before.items():
            self.assertEqual(after[name], rows, name)

    def test_the_master_holds_every_unique_record(self):
        fs.CYCLE_SIZE = 3
        rows = [school("School %d" % n, **{"PIN Code": "5600%02d" % n,
                                           "School Address": "%d Road" % n})
                for n in range(1, 11)]
        fs.append_records(rows)
        self.assertEqual(fs.row_count(fs.master_path()), 10)
        in_cycles = sum(fs.row_count(p) for p in fs.all_cycle_paths())
        self.assertEqual(in_cycles, 10)

    def test_a_duplicate_is_caught_across_finished_cycles(self):
        fs.CYCLE_SIZE = 2
        first = school("Repeat School", **{"PIN Code": "560001",
                                           "School Address": "1 Road"})
        fs.append_records([first,
                           school("Other 1", **{"PIN Code": "560002",
                                                "School Address": "2 Road"}),
                           school("Other 2", **{"PIN Code": "560003",
                                                "School Address": "3 Road"})])
        self.assertIn(1, fs.finished_cycle_numbers())
        written, skipped = fs.append_records([first])
        self.assertEqual((written, skipped), (0, 1))
        self.assertEqual(fs.row_count(fs.master_path()), 3)

    def test_the_default_cycle_size_is_a_thousand(self):
        """1000 records per cycle file, as the production plan requires."""
        import ast
        with open(SCRIPT, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        for node in ast.walk(tree):
            if isinstance(node, ast.Assign) and any(
                    getattr(t, "id", "") == "CYCLE_SIZE"
                    for t in node.targets):
                self.assertEqual(ast.literal_eval(node.value), 1000)
                return
        self.fail("CYCLE_SIZE was not found")

    def test_a_school_is_never_written_twice(self):
        rows = [school("Repeat School", **{"PIN Code": "560001",
                                           "School Address": "1 Road"})]
        self.assertEqual(fs.append_records(rows)[0], 1)
        self.assertEqual(fs.append_records(rows), (0, 1))

    def test_a_damaged_workbook_is_kept_and_replaced(self):
        fs.append_records([school("First", **{"PIN Code": "560001",
                                              "School Address": "1 Road"})])
        path = fs.cycle_path(0)
        with open(path, "wb") as handle:
            handle.write(b"not a workbook")
        book = fs.load_workbook_safely(path)
        self.assertEqual([c.value for c in book.active[1]],
                         list(fs.SCHOOL_COLUMNS))
        self.assertTrue([n for n in os.listdir(fs.EXCEL_DIR)
                         if ".damaged_" in n])

    def test_leading_zeros_survive(self):
        fs.append_records([school("Zero School",
                                  **{"PIN Code": "560001",
                                     "School Address": "1 Road",
                                     "School Official Phone Number": "08012345678",
                                     "UDISE_Code" if False else "UDISE Code":
                                         "029200112345"})])
        row = fs.read_records(fs.cycle_path(0))[0]
        self.assertEqual(row["School Official Phone Number"], "08012345678")
        self.assertEqual(row["UDISE Code"], "029200112345")


# ===========================================================================
# 10. checkpoint and resume
# ===========================================================================

class CheckpointTests(SandboxTest):
    def test_a_crash_during_discovery_loses_nothing_already_found(self):
        records = [school("School %d" % n) for n in range(1, 6)]
        fs.save_city_progress("Karnataka", "Bengaluru", records)
        again = fs.load_city_progress("Karnataka", "Bengaluru")
        self.assertEqual(len(again), 5)

    def test_a_crash_during_enrichment_resumes_at_the_next_school(self):
        point = fs.Checkpoint()
        point.mark_enriched("Karnataka", "Bengaluru", "name:alpha")
        fresh = fs.Checkpoint()
        self.assertIn("name:alpha", fresh.enriched_keys("Karnataka",
                                                        "Bengaluru"))

    def test_a_city_is_not_complete_before_its_rows_are_saved(self):
        point = fs.Checkpoint()
        self.assertFalse(point.is_city_complete("Karnataka", "Bengaluru"))
        point.mark_city_complete("Karnataka", "Bengaluru", 3, 0)
        self.assertTrue(fs.Checkpoint().is_city_complete("Karnataka",
                                                         "Bengaluru"))

    def test_a_changed_rank_stops_a_resume(self):
        point = fs.Checkpoint()
        point.adopt_plan(1, 25, "fingerprint")
        point.save()
        self.assertTrue(fs.Checkpoint().check_plan(2, 25, "fingerprint"))
        self.assertTrue(fs.Checkpoint().check_plan(1, 10, "fingerprint"))
        # A different plan fingerprint with the SAME cities is a keyword
        # change, which is migrated. Changing the stored city fingerprint
        # is what makes this the city change the test is about.
        point = fs.Checkpoint()
        point.data["city_fingerprint"] = "a-different-city-list"
        self.assertTrue(point.check_plan(1, 25, "other"))
        self.assertEqual(fs.Checkpoint().check_plan(1, 25, "fingerprint"), "")

    def test_a_damaged_checkpoint_does_not_lose_the_excel_data(self):
        fs.append_records([school("Kept School", **{"PIN Code": "560001",
                                                    "School Address": "1 R"})])
        with open(fs.CHECKPOINT_FILE, "w", encoding="utf-8") as handle:
            handle.write("{ not json")
        point = fs.Checkpoint()
        self.assertEqual(point.data["completed_cities"], [])
        self.assertEqual(fs.row_count(fs.cycle_path(0)), 1)

    def test_a_failed_school_stays_retryable(self):
        point = fs.Checkpoint()
        point.record_failure(school("Failed School"), "no answer")
        self.assertEqual(len(fs.Checkpoint().data["failed"]), 1)
        taken = fs.Checkpoint().take_failures()
        self.assertEqual(len(taken), 1)
        self.assertEqual(taken[0]["record"]["School Name"], "Failed School")


# ===========================================================================
# 11. the whole city pipeline, mocked end to end
# ===========================================================================

class PipelineTests(SandboxTest):
    def _run_city(self, cards, answer=ANSWER, **page_kwargs):
        maps_page = FakeMapsPage(cards, per_scroll=5, visible=5, **page_kwargs)
        director = FakeDirector([maps_page], FakeGeminiPage(answer))
        gemini = fs.Gemini(director)
        point = fs.Checkpoint()
        report = fs.process_city(director, gemini, point, "Karnataka",
                                 "Bengaluru", ["Private Schools"])
        self.assertLessEqual(director.max_open, 1,
                             "two browsers were open at once")
        return report, point

    def test_a_city_runs_discovery_research_save_and_completion(self):
        cards = [maps_entry(n, name="School %d" % n) for n in range(1, 6)]
        answer = ANSWER.replace("School_Name: {name}", "School_Name: School 1")
        report, point = self._run_city(cards, answer=answer.format(name="x"))
        self.assertEqual(report["found"], 5)
        self.assertGreaterEqual(report["saved"], 1)
        self.assertEqual(fs.row_count(fs.cycle_path(0)), report["saved"])

    def test_an_incomplete_search_leaves_the_city_open(self):
        cards = [maps_entry(n) for n in range(1, 6)]
        report, point = self._run_city(cards, end_message=False)
        self.assertFalse(report["complete"])
        self.assertFalse(point.is_city_complete("Karnataka", "Bengaluru"))
        self.assertIn("unresolved", report["note"])

    def test_the_working_files_outlive_an_unfinished_city(self):
        cards = [maps_entry(n) for n in range(1, 6)]
        self._run_city(cards, end_message=False)
        self.assertTrue(os.path.exists(
            fs.temp_excel_path("Karnataka", "Bengaluru")),
            "the temporary workbook was removed from an unfinished city")

    def test_rows_reach_excel_before_the_city_is_marked_complete(self):
        order = []
        real_append = fs.append_records
        real_mark = fs.Checkpoint.mark_city_complete

        def watched_append(records):
            order.append("save")
            return real_append(records)

        def watched_mark(self, *a, **k):
            order.append("complete")
            return real_mark(self, *a, **k)

        fs.append_records = watched_append
        fs.Checkpoint.mark_city_complete = watched_mark
        try:
            cards = [maps_entry(1, name="School 1")]
            self._run_city(cards,
                           answer=ANSWER.replace("{name}", "School 1"))
        finally:
            fs.append_records = real_append
            fs.Checkpoint.mark_city_complete = real_mark
        self.assertIn("save", order)
        if "complete" in order:
            self.assertLess(order.index("save"), order.index("complete"))


# ===========================================================================
# 12. the single file really is single
# ===========================================================================

class PortabilityTests(unittest.TestCase):
    BANNED = ("new_mylib", "mylib", "config", "school_source",
              "gemini_controller", "source_file", "lib",
              "school_gemini_prompt", "school_browser", "school_excel")

    def test_it_imports_no_project_file(self):
        import ast
        with open(SCRIPT, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        names = []
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                names += [a.name for a in node.names]
            elif isinstance(node, ast.ImportFrom) and node.module:
                names.append(node.module)
        for name in names:
            self.assertNotIn(name.split(".")[0], self.BANNED, name)

    def test_the_only_outside_file_it_reads_is_mylib(self):
        """Cities and keywords are shared; nothing else is read from disk."""
        import ast
        with open(SCRIPT, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        literals = [n.value for n in ast.walk(tree)
                    if isinstance(n, ast.Constant)
                    and isinstance(n.value, str)]
        docstrings = set()
        for node in ast.walk(tree):
            if isinstance(node, (ast.Module, ast.FunctionDef, ast.ClassDef)):
                text = ast.get_docstring(node, clean=False)
                if text:
                    docstrings.add(text)
        paths = [t for t in literals if t not in docstrings
                 and ("/" in t or t.endswith((".txt", ".py", ".json")))]
        for text in paths:
            if text.endswith("mylib.py") or text in fs.MYLIB_PATHS:
                continue
            for name in ("school_gemini_prompt", "new_mylib",
                         "Evision_School_Information_Columns",
                         "approved_city_dictionary", "school_types",
                         "source_file"):
                self.assertNotIn(name, text,
                                 "%r is used as a runtime path" % name)

    def test_it_opens_only_its_own_runtime_files(self):
        import ast
        with open(SCRIPT, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        for node in ast.walk(tree):
            if isinstance(node, ast.Call) and getattr(node.func, "id", "") \
                    == "open":
                if node.args and isinstance(node.args[0], ast.Constant):
                    self.fail("open() on the hard-coded path %r"
                              % node.args[0].value)

    def test_the_lists_are_not_duplicated_in_the_script(self):
        """One source of truth: mylib.py. No second copy to drift."""
        with open(SCRIPT, encoding="utf-8") as handle:
            source = handle.read()
        self.assertNotIn("FALLBACK_CITY_BY_STATES", source)
        self.assertNotIn("FALLBACK_SCHOOL_KEYWORDS", source)
        self.assertNotIn("\nSCHOOL_FIELDS = (", source,
                         "the 57 fields are written out here as well")
        self.assertIn("GEMINI_PROMPT_TEMPLATE", source)

    def test_the_city_names_are_not_written_out_here(self):
        """A copy of the cities would be a second source of truth."""
        with open(SCRIPT, encoding="utf-8") as handle:
            source = handle.read()
        for city in ("'Gurugram'", '"Gurugram"', "'Bengaluru'",
                     "'Jorethang'"):
            self.assertNotIn(city, source,
                             "%s is written into the script" % city)

    def test_it_carries_no_credential(self):
        """A comment may mention the word; no value may be one."""
        import ast
        with open(SCRIPT, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        for node in ast.walk(tree):
            if isinstance(node, ast.Constant) and isinstance(node.value, str):
                low = node.value.lower()
                for secret in ("@gmail.com", "@evisiontechnoserve",
                               "api_key=", "smtp"):
                    self.assertNotIn(secret, low, node.value[:60])

    def test_show_is_a_real_option(self):
        args = fs.parse_args(["--show"])
        self.assertTrue(args.show)
        self.assertFalse(fs.parse_args([]).show)

    def test_the_script_and_mylib_are_enough(self):
        """Two files in an empty folder, nothing else."""
        import subprocess
        folder = tempfile.mkdtemp(prefix="school_alone_")
        try:
            shutil.copy(SCRIPT, os.path.join(folder, "find_school.py"))
            source = fs.mylib_path()
            if not source or not os.path.exists(source):
                self.skipTest("no mylib.py on this machine")
            shutil.copy(source, os.path.join(folder, "mylib.py"))
            done = subprocess.run(
                [sys.executable, os.path.join(folder, "find_school.py"),
                 "--self-test"], cwd=folder, capture_output=True, text=True,
                timeout=300)
            self.assertEqual(done.returncode, 0,
                             done.stdout[-2000:] + done.stderr[-2000:])
            self.assertIn("checks passed", done.stdout)
        finally:
            shutil.rmtree(folder, ignore_errors=True)

    def test_without_mylib_it_says_so_instead_of_running(self):
        import subprocess
        folder = tempfile.mkdtemp(prefix="school_nolib_")
        try:
            shutil.copy(SCRIPT, os.path.join(folder, "find_school.py"))
            env = dict(os.environ, SCHOOL_MYLIB="/nonexistent/mylib.py")
            done = subprocess.run(
                [sys.executable, os.path.join(folder, "find_school.py"),
                 "--status"], cwd=folder, capture_output=True, text=True,
                timeout=120, env=env)
            self.assertIn("mylib", (done.stdout + done.stderr).lower())
        finally:
            shutil.rmtree(folder, ignore_errors=True)

    def test_dry_run_opens_no_browser(self):
        import subprocess
        done = subprocess.run(
            [sys.executable, SCRIPT, "--dry-run", "--rank", "1",
             "--num-sys", "866"],
            cwd=HERE, capture_output=True, text=True, timeout=120)
        self.assertEqual(done.returncode, 0, done.stderr[-1000:])
        self.assertIn("no browser is opened", done.stdout)
        self.assertIn("maps/search/", done.stdout)


# ===========================================================================
# 13. --show: putting the window on the real desktop
# ===========================================================================

class ShowOptionTests(unittest.TestCase):
    """The decision logic. The visibility itself is checked live, below."""

    def setUp(self):
        self.kept = dict(os.environ)

    def tearDown(self):
        os.environ.clear()
        os.environ.update(self.kept)

    def test_xvfb_is_detected_from_the_authority_file(self):
        os.environ["XAUTHORITY"] = "/tmp/xvfb-run.ab12Cd/Xauthority"
        self.assertTrue(fs.running_under_xvfb())
        os.environ["XAUTHORITY"] = "/run/user/1000/.Xauthority"
        self.assertFalse(fs.running_under_xvfb())

    def test_show_forces_a_drawn_window(self):
        browser = fs.Browser(show=True)
        self.assertEqual(browser.mode, "headful")

    def test_without_show_no_window_lands_on_a_real_desktop(self):
        os.environ["DISPLAY"] = ":0"
        os.environ["XAUTHORITY"] = "/run/user/1000/.Xauthority"
        self.assertEqual(fs.Browser().mode, "headless")

    def test_without_show_the_virtual_screen_is_used_as_it_is(self):
        os.environ["DISPLAY"] = ":99"
        os.environ["XAUTHORITY"] = "/tmp/xvfb-run.ab12Cd/Xauthority"
        self.assertEqual(fs.Browser().mode, "headful")

    def test_an_unreachable_display_is_an_error_not_a_silent_fallback(self):
        os.environ["SCHOOL_DESKTOP_DISPLAY"] = ":77"
        os.environ.pop("SCHOOL_DESKTOP_XAUTHORITY", None)
        with self.assertRaises(fs.DesktopUnavailable) as caught:
            fs.find_desktop_display()
        self.assertIn(":77", str(caught.exception))

    def test_a_display_failure_stops_the_run(self):
        self.assertTrue(issubclass(fs.DesktopUnavailable,
                                   fs.BrowserUnavailable))

    def test_the_wayland_socket_is_taken_out_of_the_browser_environment(self):
        """Chromium prefers Wayland and then ignores DISPLAY entirely."""
        os.environ["WAYLAND_DISPLAY"] = "wayland-0"
        env = fs._display_env(":0", "/tmp/auth")
        self.assertEqual(env["DISPLAY"], ":0")
        self.assertEqual(env["XAUTHORITY"], "/tmp/auth")
        self.assertNotIn("WAYLAND_DISPLAY", env)

    def test_the_x11_backend_is_forced_when_showing(self):
        """Without this flag Chromium draws no window at all on Wayland."""
        import ast
        with open(SCRIPT, encoding="utf-8") as handle:
            source = handle.read()
        self.assertIn("--ozone-platform=x11", source)

    def test_a_clipboard_helper_window_is_not_mistaken_for_the_browser(self):
        real = fs.confirm_window_is_visible
        with unittest.mock.patch.object(
                fs, "windows_on_display",
                return_value=["Chromium clipboard"]):
            ok, detail = real(":0", "", wait=0.2)
        self.assertFalse(ok)
        with unittest.mock.patch.object(
                fs, "windows_on_display",
                return_value=["Chromium clipboard",
                              "Maps - Google Chrome for Testing"]):
            ok, detail = real(":0", "", wait=0.2)
        self.assertTrue(ok)
        self.assertIn("Chrome for Testing", detail)


# ===========================================================================
# 14. one visible tab, Gemini out of the way
# ===========================================================================

class FakeTab:
    def __init__(self, url="about:blank"):
        self.url = url
        self.closed = False
        self.fronted = 0
        self.keyboard = FakeKeyboard(self)
        self.typed, self.pressed = [], []

    def is_closed(self):
        return self.closed

    def close(self):
        self.closed = True

    def bring_to_front(self):
        self.fronted += 1

    def title(self):
        return self.url

    def goto(self, url, **_kwargs):
        self.url = url

    def submit(self):
        pass


class FakeContext:
    def __init__(self, pages):
        self.pages = list(pages)
        self.opened = 0

    def new_page(self):
        self.opened += 1
        tab = FakeTab()
        self.pages.append(tab)
        return tab

    def set_default_timeout(self, _ms):
        pass

    def close(self):
        pass


def browser_with(pages):
    """A real fs.Browser wired to a fake context - no Chromium involved."""
    browser = fs.Browser.__new__(fs.Browser)
    browser.show = False
    browser.mode = "headful"
    browser.profile_dir = "/tmp/unused"
    browser.label = "browser"
    browser.display = ""
    browser.authority = ""
    browser._kept_env = {}
    browser._play = None
    browser._page = None
    browser.restarts = 0
    browser._context = FakeContext(pages)
    return browser


class OneVisibleTabTests(unittest.TestCase):
    def test_the_contexts_own_blank_tab_is_used_not_a_second_one(self):
        """This is the about:blank tab that used to sit beside the map."""
        blank = FakeTab("about:blank")
        browser = browser_with([blank])
        self.assertIs(browser.page(), blank)
        self.assertEqual(browser._context.opened, 0)
        self.assertEqual(len(browser._context.pages), 1)

    def test_a_tab_is_opened_only_when_there_is_none(self):
        browser = browser_with([])
        page = browser.page()
        self.assertEqual(browser._context.opened, 1)
        self.assertIs(browser.page(), page)
        self.assertEqual(browser._context.opened, 1)

    def test_spare_tabs_are_closed_and_the_working_one_is_kept(self):
        keep, spare1, spare2 = FakeTab("maps"), FakeTab("x"), FakeTab("y")
        browser = browser_with([keep, spare1, spare2])
        browser.page()
        self.assertFalse(keep.closed)
        self.assertTrue(spare1.closed and spare2.closed)

    def test_the_working_tab_can_be_brought_to_the_front(self):
        tab = FakeTab("maps")
        browser = browser_with([tab])
        browser.page()
        browser.show_tab()
        self.assertEqual(tab.fronted, 1)


class GeminiStaysOutOfTheWayTests(SandboxTest):
    def test_gemini_gets_its_own_browser_from_the_director(self):
        own_page = FakeGeminiPage(ANSWER.format(name="Gupta Public School"))
        director = FakeDirector([FakeMapsPage([])], own_page)
        gemini = fs.Gemini(director)
        page = gemini.page()
        self.assertIs(page, own_page)
        self.assertEqual(director.kind, "gemini")
        self.assertEqual(director.max_open, 1)

    def test_the_maps_browser_is_closed_before_gemini_opens(self):
        director = FakeDirector([FakeMapsPage([])], FakeGeminiPage(ANSWER))
        director.maps()
        fs.Gemini(director).page()
        self.assertEqual(director.events,
                         [("open", "maps"), ("close", "maps"),
                          ("open", "gemini")])
        self.assertEqual(director.max_open, 1)

    def test_in_fallback_mode_the_map_comes_back_after_every_answer(self):
        page = FakeGeminiPage(ANSWER.format(name="Gupta Public School"))
        maps = FakeBrowser([FakeMapsPage([]), page])
        gemini = fs.Gemini(maps, separate=False)
        before = maps.brought_to_front
        gemini.research(school("Gupta Public School"))
        self.assertGreater(maps.brought_to_front, before,
                           "the Maps tab was left behind Gemini")

    def test_closing_the_gemini_browser_drops_its_page(self):
        director = FakeDirector([FakeMapsPage([])], FakeGeminiPage(ANSWER))
        gemini = fs.Gemini(director)
        gemini.page()
        self.assertIsNotNone(gemini._page)
        director.close()
        self.assertIsNone(gemini._page,
                          "Gemini kept a page from a browser that is gone")

class VisibleTabFollowsTheSearchTests(SandboxTest):
    def test_each_search_retunes_and_fronts_the_visible_tab(self):
        page = FakeMapsPage([maps_entry(1)])
        browser = FakeBrowser([page])
        fs.collect_search(browser, "CBSE Schools", "Karnataka", "Mysuru")
        self.assertIn("CBSE+Schools+in+Mysuru", page.url)
        self.assertGreaterEqual(browser.brought_to_front, 1)
        self.assertGreaterEqual(browser.spare_closed, 1)

    def test_a_new_city_moves_the_visible_tab_to_that_city(self):
        page = FakeMapsPage([maps_entry(1)])
        browser = FakeBrowser([page])
        for city in ("Delhi", "Noida"):
            fs.collect_search(browser, "Private Schools", "Delhi NCR", city)
            self.assertIn("in+%s" % city, page.url)


# ===========================================================================
# 15. reading a real Maps card
# ===========================================================================

class CardReadingTests(unittest.TestCase):
    """Card layouts taken from live Google Maps results."""

    @staticmethod
    def read(text, name="Example School", href=None):
        href = href or ("https://www.google.com/maps/place/X/data=!1s0x1:0x2"
                        "!3d12.9!4d77.6")
        return fs.card_to_record({"href": href, "label": name, "text": text},
                                 "Delhi NCR", "Delhi")

    def test_the_usual_card(self):
        record = self.read(
            "PP International School\nPP International School\nNo reviews\n"
            "CBSE School ·  · Metro Pillar Number 333, Lala Jagat Narayan Marg\n"
            "Open · Closes 8 pm · 011 4516 1066\nWebsite\nDirections",
            name="PP International School")
        self.assertEqual(record["School Name"], "PP International School")
        self.assertEqual(record["School Address"],
                         "Metro Pillar Number 333, Lala Jagat Narayan Marg")
        self.assertEqual(record["School Official Phone Number"], "01145161066")
        self.assertEqual(record["School Latitude"], "12.9")

    def test_a_rating_is_never_the_address(self):
        """A rating line sits above the address line on the card."""
        for rating in ("4.7 (149)", "4.7 ★★★★★ (149)",
                       "4.8", "No reviews"):
            record = self.read(
                "VSPK International School\n%s\n"
                "Educational institution ·  · Sector 12 Rohini\n"
                "Closes soon · 4 pm · 085888 22255" % rating,
                name="VSPK International School")
            self.assertEqual(record["School Address"], "Sector 12 Rohini",
                             "rating %r leaked into the address" % rating)

    def test_interface_words_are_never_the_address(self):
        record = self.read(
            "Doon Public School\n4.8\nSchool ·  · B 2\n"
            "Closed · Opens 8 am Fri · 011 4231 6891\n\nWebsite\n\nDirections",
            name="Doon Public School")
        self.assertEqual(record["School Address"], "B 2")
        self.assertNotIn(record["School Address"], ("Website", "Directions"))

    def test_a_phone_on_a_line_of_its_own_is_found(self):
        """A card with no opening hours puts the number on its own line."""
        record = self.read(
            "The Mother's International School\nNo reviews\n"
            "CBSE School ·  · junior wing, Sri Aurobindo Marg\n"
            "011 2696 4140\nWebsite\nDirections",
            name="The Mother's International School")
        self.assertEqual(record["School Official Phone Number"], "01126964140")
        self.assertEqual(record["School Address"],
                         "junior wing, Sri Aurobindo Marg")

    def test_the_address_is_never_read_as_the_phone_number(self):
        record = self.read(
            "Richmondd Global School\nNo reviews\n"
            "CBSE School ·  · M3HM+MXC, N.S. Road, Opp. Inder Enclave\n"
            "Open · Closes 5 pm · 093120 83145",
            name="Richmondd Global School")
        self.assertEqual(record["School Official Phone Number"], "09312083145")

    def test_nothing_is_invented(self):
        record = self.read("Lone School\nSchool", name="Lone School")
        self.assertEqual(record["School Name"], "Lone School")
        for column in ("UDISE Code", "Affiliation Board", "Principal Name",
                       "School Website", "Total Student Strength"):
            self.assertEqual(record[column], "", column)


# ===========================================================================
# 16. the "Restore pages?" bubble
# ===========================================================================

class RestoreBubbleTests(unittest.TestCase):
    def test_the_crash_flag_is_cleared_before_launching(self):
        folder = tempfile.mkdtemp(prefix="school_prefs_")
        try:
            default = os.path.join(folder, "Default")
            os.makedirs(default)
            path = os.path.join(default, "Preferences")
            with open(path, "w", encoding="utf-8") as handle:
                json.dump({"profile": {"exit_type": "Crashed",
                                       "exited_cleanly": False},
                           "something_else": {"keep": "me"}}, handle)
            self.assertTrue(fs.forget_the_last_crash(folder))
            with open(path, encoding="utf-8") as handle:
                prefs = json.load(handle)
            self.assertEqual(prefs["profile"]["exit_type"], "Normal")
            self.assertTrue(prefs["profile"]["exited_cleanly"])
            self.assertEqual(prefs["something_else"], {"keep": "me"},
                             "the rest of the profile was disturbed")
        finally:
            shutil.rmtree(folder, ignore_errors=True)

    def test_a_clean_profile_is_left_alone(self):
        folder = tempfile.mkdtemp(prefix="school_prefs_")
        try:
            default = os.path.join(folder, "Default")
            os.makedirs(default)
            with open(os.path.join(default, "Preferences"), "w",
                      encoding="utf-8") as handle:
                json.dump({"profile": {"exit_type": "Normal",
                                       "exited_cleanly": True}}, handle)
            self.assertFalse(fs.forget_the_last_crash(folder))
        finally:
            shutil.rmtree(folder, ignore_errors=True)

    def test_a_missing_or_broken_file_is_survivable(self):
        folder = tempfile.mkdtemp(prefix="school_prefs_")
        try:
            self.assertFalse(fs.forget_the_last_crash(folder))
            default = os.path.join(folder, "Default")
            os.makedirs(default)
            with open(os.path.join(default, "Preferences"), "w",
                      encoding="utf-8") as handle:
                handle.write("{ not json")
            self.assertFalse(fs.forget_the_last_crash(folder))
        finally:
            shutil.rmtree(folder, ignore_errors=True)

    def test_the_bubble_is_suppressed_on_the_command_line_too(self):
        with open(SCRIPT, encoding="utf-8") as handle:
            source = handle.read()
        for flag in ("--hide-crash-restore-bubble",
                     "--disable-session-crashed-bubble",
                     "--restore-last-session=false"):
            self.assertIn(flag, source, flag)


# ===========================================================================
# 17. the captured .txt files and the regex extraction
# ===========================================================================

class CaptureFileTests(SandboxTest):
    def test_every_copy_is_saved_as_its_own_snapshot(self):
        page = FakeMapsPage([maps_entry(n) for n in range(1, 41)],
                            per_scroll=5, visible=5)
        result = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                   "Karnataka", "Bengaluru")
        self.assertTrue(os.path.exists(result.txt_path))
        with open(result.txt_path, encoding="utf-8") as handle:
            text = handle.read()
        snapshots = fs.read_snapshots(text)
        self.assertEqual(len(snapshots), result.snapshots)
        self.assertGreater(len(snapshots), 1)
        self.assertEqual([s["number"] for s in snapshots],
                         list(range(1, len(snapshots) + 1)))

    def test_the_text_is_saved_before_every_down_press(self):
        """Never a Down before the text on screen has been written."""
        page = FakeMapsPage([maps_entry(n) for n in range(1, 25)],
                            per_scroll=6, visible=6)
        fs.collect_search(FakeBrowser([page]), "Private Schools",
                          "Karnataka", "Bengaluru")
        order = [k for k in page.pressed
                 if k in ("Control+a", "Control+c", "ArrowDown")]
        self.assertEqual(order[0], "Control+a")
        self.assertEqual(order[1], "Control+c")
        seen_copy = False
        for key in order:
            if key == "Control+c":
                seen_copy = True
            if key == "ArrowDown":
                self.assertTrue(seen_copy, "a Down came before any copy")

    def test_the_keyboard_is_really_used(self):
        page = FakeMapsPage([maps_entry(n) for n in range(1, 13)])
        fs.collect_search(FakeBrowser([page]), "Private Schools",
                          "Karnataka", "Bengaluru")
        self.assertIn("Control+a", page.pressed)
        self.assertIn("Control+c", page.pressed)
        self.assertIn("ArrowDown", page.pressed)
        self.assertGreater(page.copies, 1)
        self.assertGreater(page.feed.focused, 0,
                           "the results list was never focused")

    def test_no_card_is_clicked_and_no_place_page_opened(self):
        page = FakeMapsPage([maps_entry(n) for n in range(1, 13)])
        fs.collect_search(FakeBrowser([page]), "Private Schools",
                          "Karnataka", "Bengaluru")
        self.assertTrue(page.url.startswith(
            "https://www.google.com/maps/search/"))
        self.assertNotIn("/maps/place/", page.url)

    def test_overlapping_snapshots_never_duplicate_a_school(self):
        """Each snapshot repeats every entry seen so far - that is normal."""
        page = FakeMapsPage([maps_entry(n) for n in range(1, 31)],
                            per_scroll=3, visible=6)
        result = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                   "Karnataka", "Bengaluru")
        names = [r["School Name"] for r in result.records]
        self.assertEqual(len(names), len(set(names)))
        self.assertEqual(len(names), 30)

    def test_the_file_is_named_after_the_keyword_and_the_city(self):
        page = FakeMapsPage([maps_entry(1)])
        result = fs.collect_search(FakeBrowser([page]), "CBSE Schools",
                                   "Karnataka", "Mysuru")
        self.assertTrue(result.txt_path.endswith(
            "Karnataka_Mysuru_cbse_schools.txt"), result.txt_path)
        self.assertEqual(os.path.dirname(result.txt_path), fs.TEMP_DIR,
                         "a folder was made for the city")

    def test_an_interrupted_search_keeps_what_it_copied(self):
        page = FakeMapsPage([maps_entry(n) for n in range(1, 21)],
                            per_scroll=5, visible=5, end_message=False)
        result = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                   "Karnataka", "Bengaluru")
        self.assertFalse(result.complete)
        with open(result.txt_path, encoding="utf-8") as handle:
            text = handle.read()
        self.assertGreater(len(fs.read_snapshots(text)), 1)
        self.assertIn("INCOMPLETE", text)

    def test_a_failed_clipboard_leaves_the_search_incomplete(self):
        """It must not quietly read the page some other way."""
        page = FakeMapsPage([maps_entry(n) for n in range(1, 11)])
        page.copied_text = lambda: ""
        result = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                   "Karnataka", "Bengaluru")
        self.assertFalse(result.complete)
        self.assertIn("clipboard", result.reason)
        self.assertEqual(result.records, [])

    def test_a_copy_that_suddenly_shrinks_is_retried(self):
        """Seen live: a copy came back with 14 characters mid-search."""
        page = FakeMapsPage([maps_entry(n) for n in range(1, 21)],
                            per_scroll=5, visible=5)
        real = page.copied_text
        state = {"calls": 0}

        def flaky():
            state["calls"] += 1
            if state["calls"] == 3:
                return "Results\n"          # the clipboard let go
            return real()
        page.copied_text = flaky
        result = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                   "Karnataka", "Bengaluru")
        self.assertTrue(result.complete,
                        "one bad copy ended the whole search")
        self.assertEqual(len(result.records), 20)

    def test_a_tiny_first_copy_is_not_mistaken_for_an_empty_city(self):
        """Seen live: the first copy came back with 26 characters."""
        page = FakeMapsPage([maps_entry(n) for n in range(1, 11)])
        page.copied_text = lambda: "Results\nShare\n"   # 26 characters
        result = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                   "Karnataka", "Bengaluru")
        self.assertFalse(result.complete)
        self.assertIn("clipboard", result.reason,
                      "a failed clipboard was reported as an empty list")

    def test_a_search_gives_up_if_the_clipboard_keeps_failing(self):
        page = FakeMapsPage([maps_entry(n) for n in range(1, 21)],
                            per_scroll=5, visible=5)
        real = page.copied_text
        state = {"calls": 0}

        def failing():
            state["calls"] += 1
            return real() if state["calls"] < 2 else "x\n"
        page.copied_text = failing
        result = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                   "Karnataka", "Bengaluru")
        self.assertFalse(result.complete)
        self.assertIn("clipboard", result.reason)

    def test_the_file_is_added_to_not_overwritten_on_a_second_run(self):
        page = FakeMapsPage([maps_entry(1)])
        first = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                  "Karnataka", "Bengaluru")
        size = os.path.getsize(first.txt_path)
        page2 = FakeMapsPage([maps_entry(2)])
        second = fs.collect_search(FakeBrowser([page2]), "Private Schools",
                                   "Karnataka", "Bengaluru")
        self.assertEqual(first.txt_path, second.txt_path)
        self.assertGreater(os.path.getsize(second.txt_path), size)

    def test_records_are_rebuilt_from_the_saved_text_alone(self):
        page = FakeMapsPage([maps_entry(n) for n in range(1, 11)],
                            per_scroll=5, visible=5)
        result = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                   "Karnataka", "Bengaluru")
        again = fs.records_from_txt(result.txt_path, "Karnataka", "Bengaluru")
        self.assertEqual([r["School Name"] for r in again],
                         [r["School Name"] for r in result.records])
        self.assertTrue(all(r["School Official Phone Number"] for r in again))

    def test_interface_text_never_becomes_a_school(self):
        page = FakeMapsPage([maps_entry(n) for n in range(1, 11)])
        result = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                   "Karnataka", "Bengaluru")
        names = [r["School Name"].lower() for r in result.records]
        for junk in ("results", "website", "directions", "rating", "hours",
                     "all filters", "saved", "recents", "get app", "share"):
            self.assertNotIn(junk, names)

    def test_the_end_of_list_line_is_not_a_school(self):
        page = FakeMapsPage([maps_entry(n) for n in range(1, 6)])
        result = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                   "Karnataka", "Bengaluru")
        self.assertTrue(result.complete)
        for record in result.records:
            self.assertNotIn("end of the list",
                             record["School Name"].lower())
            self.assertNotIn("end of the list",
                             record["School Address"].lower())


class KeptEvidenceTests(SandboxTest):
    """KEEP_CAPTURED_TEXT = True. Off is the default and is covered by
    CleanupTests below; this class is about what the setting does when it
    is deliberately turned on."""

    def setUp(self):
        SandboxTest.setUp(self)
        self.was = fs.KEEP_CAPTURED_TEXT
        fs.KEEP_CAPTURED_TEXT = True

    def tearDown(self):
        fs.KEEP_CAPTURED_TEXT = self.was
        SandboxTest.tearDown(self)

    def test_the_captured_text_survives_the_city_cleanup(self):
        page = FakeMapsPage([maps_entry(n) for n in range(1, 6)])
        result = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                   "Karnataka", "Bengaluru")
        fs.save_temp_excel("Karnataka", "Bengaluru", result.records)
        fs.save_city_progress("Karnataka", "Bengaluru", result.records)
        fs.clear_city_temp("Karnataka", "Bengaluru")
        self.assertTrue(os.path.exists(result.txt_path),
                        "the captured text was deleted")
        self.assertFalse(os.path.exists(
            fs.temp_excel_path("Karnataka", "Bengaluru")))
        self.assertFalse(os.path.exists(os.path.join(
            fs.city_dir("Karnataka", "Bengaluru"), "schools.json")))

    def test_the_text_can_still_be_turned_back_into_records(self):
        page = FakeMapsPage([maps_entry(n) for n in range(1, 6)])
        result = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                   "Karnataka", "Bengaluru")
        fs.clear_city_temp("Karnataka", "Bengaluru")
        again = fs.records_from_txt(result.txt_path, "Karnataka", "Bengaluru")
        self.assertEqual(len(again), 5)


class TemporaryExcelTests(SandboxTest):
    def test_the_city_gets_a_57_column_temporary_workbook(self):
        rows = [school("School %d" % n, **{"PIN Code": "5600%02d" % n})
                for n in range(1, 6)]
        path = fs.save_temp_excel("Karnataka", "Bengaluru", rows)
        self.assertTrue(path.endswith("temporary_schools.xlsx"))
        back = fs.load_temp_excel("Karnataka", "Bengaluru")
        self.assertEqual(len(back), 5)
        self.assertEqual(list(back[0].keys()), list(fs.SCHOOL_COLUMNS))

    def test_the_temporary_file_is_not_the_final_file(self):
        fs.save_temp_excel("Karnataka", "Bengaluru",
                           [school("Temp Only", **{"PIN Code": "560001"})])
        self.assertEqual(fs.row_count(fs.master_path()), 0)
        self.assertEqual(fs.all_cycle_paths(), [])


class FolderLayoutTests(unittest.TestCase):
    def setUp(self):
        self.kept = (fs.OUTPUT_DIR, fs.TEMP_DIR, fs.EXCEL_DIR,
                     fs.CHECKPOINT_FILE)
        self.folder = tempfile.mkdtemp(prefix="school_layout_")
        fs.OUTPUT_DIR = self.folder

    def tearDown(self):
        (fs.OUTPUT_DIR, fs.TEMP_DIR, fs.EXCEL_DIR,
         fs.CHECKPOINT_FILE) = self.kept
        shutil.rmtree(self.folder, ignore_errors=True)

    def test_the_folders_are_flat_whatever_the_rank(self):
        for this_rank in (1, 7, 25):
            fs.set_rank_paths(this_rank)
            self.assertEqual(fs.EXCEL_DIR,
                             os.path.join(self.folder, "Excel_Data"))
            self.assertEqual(fs.TEMP_DIR,
                             os.path.join(self.folder, "Temp_Data"))
            self.assertNotIn("rank_", fs.TEMP_DIR)

    def test_the_checkpoint_sits_beside_the_script(self):
        fs.set_rank_paths(3)
        self.assertEqual(fs.CHECKPOINT_FILE,
                         os.path.join(self.folder, "checkpoint.json"))
        self.assertFalse(os.path.isdir(os.path.join(self.folder, "state")))
        self.assertFalse(os.path.isdir(os.path.join(self.folder, "logs")))

    def test_rank_still_decides_the_cities(self):
        """Only where it writes changed, not what it collects."""
        one = {c for _s, c in fs.cities_for_rank(1, 25)}
        seven = {c for _s, c in fs.cities_for_rank(7, 25)}
        self.assertTrue(one)
        self.assertFalse(one & seven)

    def test_the_folders_are_created(self):
        fs.set_rank_paths(1)
        for path in (fs.TEMP_DIR, fs.EXCEL_DIR):
            self.assertTrue(os.path.isdir(path), path)


class CrashRepairTests(SandboxTest):
    def test_a_record_only_in_the_master_reaches_the_current_cycle(self):
        fs.CYCLE_SIZE = 5
        rows = [school("School %d" % n, **{"PIN Code": "5600%02d" % n,
                                           "School Address": "%d Road" % n})
                for n in range(1, 4)]
        fs.append_records(rows)
        orphan = school("Orphan School", **{"PIN Code": "560099",
                                            "School Address": "99 Road"})
        fs.write_records(fs.master_path(),
                         fs.read_records(fs.master_path()) + [orphan])
        report = fs.reconcile_master_and_cycles()
        self.assertEqual(report["added_to_current"], 1)
        names = [r["School Name"] for p in fs.all_cycle_paths()
                 for r in fs.read_records(p)]
        self.assertIn("Orphan School", names)

    def test_a_record_only_in_a_cycle_reaches_the_master(self):
        fs.CYCLE_SIZE = 5
        fs.append_records([school("Known", **{"PIN Code": "560001",
                                              "School Address": "1 Road"})])
        stray = school("Stray School", **{"PIN Code": "560088",
                                          "School Address": "88 Road"})
        fs.write_records(fs.cycle_path(0),
                         fs.read_records(fs.cycle_path(0)) + [stray])
        report = fs.reconcile_master_and_cycles()
        self.assertEqual(report["added_to_master"], 1)
        self.assertIn("Stray School",
                      [r["School Name"]
                       for r in fs.read_records(fs.master_path())])

    def test_nothing_is_removed_by_the_repair(self):
        fs.CYCLE_SIZE = 5
        rows = [school("School %d" % n, **{"PIN Code": "5600%02d" % n,
                                           "School Address": "%d Road" % n})
                for n in range(1, 5)]
        fs.append_records(rows)
        before = fs.row_count(fs.master_path())
        fs.reconcile_master_and_cycles()
        self.assertEqual(fs.row_count(fs.master_path()), before)

    def test_the_journal_is_cleared_once_the_save_finishes(self):
        fs.append_records([school("Journal School",
                                  **{"PIN Code": "560001",
                                     "School Address": "1 Road"})])
        self.assertFalse(os.path.exists(fs.journal_path()))


# ===========================================================================
# 18. the order of work: one keyword finished before the next starts
# ===========================================================================

class WorkflowOrderTests(SandboxTest):
    def _city(self, keywords, per_keyword=3, answer=None):
        """Run a city with a fresh Maps page per keyword, recording order."""
        answer = answer or ANSWER
        pages = {k: FakeMapsPage([maps_entry(n, name="%s %d" % (k, n))
                                  for n in range(1, per_keyword + 1)])
                 for k in keywords}
        director = FakeDirector([pages[keywords[0]]], FakeGeminiPage(answer))
        events = []

        real_collect = fs.collect_search
        real_append = fs.append_records

        def pick(browser, keyword, state, city, limit=None):
            browser.pages[0] = pages[keyword]
            events.append(("search", keyword))
            return real_collect(browser, keyword, state, city, limit=limit)

        def watched_append(records):
            for record in records:
                events.append(("save", record["School Name"]))
            return real_append(records)

        gemini = fs.Gemini(director)
        gemini.research = lambda record: (
            events.append(("gemini", record["School Name"]))
            or (fs.parse_answer(answer.format(name=record["School Name"])),
                answer))
        fs.collect_search = pick
        fs.append_records = watched_append
        try:
            report = fs.process_city(director, gemini, fs.Checkpoint(),
                                     "Karnataka", "Bengaluru", keywords)
        finally:
            fs.collect_search = real_collect
            fs.append_records = real_append
        self.assertLessEqual(director.max_open, 1,
                             "two browsers were open at once")
        return report, events

    def test_a_keyword_is_finished_before_the_next_one_starts(self):
        keywords = ["Private Schools", "CBSE Schools"]
        _report, events = self._city(keywords)
        searches = [i for i, (kind, _v) in enumerate(events)
                    if kind == "search"]
        self.assertEqual(len(searches), 2)
        between = events[searches[0]:searches[1]]
        self.assertTrue(any(k == "gemini" for k, _v in between),
                        "no research happened before the second search")
        self.assertTrue(any(k == "save" for k, _v in between),
                        "nothing was saved before the second search")
        after_first = [v for k, v in between if k == "save"]
        self.assertTrue(all(v.startswith("Private Schools")
                            for v in after_first),
                        "a later keyword's school was saved too early")

    def test_each_school_is_saved_before_the_next_one_is_researched(self):
        _report, events = self._city(["Private Schools"], per_keyword=3)
        trail = [k for k, _v in events if k in ("gemini", "save")]
        self.assertEqual(trail, ["gemini", "save"] * 3,
                         "saving did not follow every single school")

    def test_a_keyword_with_no_results_still_counts_as_finished(self):
        """A small town has no ICSE school; the city must still close."""
        point = fs.Checkpoint()
        report = fs.process_keyword(
            FakeDirector([FakeMapsPage([], end_message=True)]), None, point,
            "Sikkim", "Mangan", "ICSE Schools", 1, 1)
        self.assertEqual(report["found"], 0)
        self.assertTrue(report["complete"],
                        "an empty but finished keyword was called unresolved")
        self.assertTrue(point.keyword_done("Sikkim", "Mangan",
                                           "ICSE Schools"))

    def test_an_empty_keyword_that_never_finished_is_not_complete(self):
        point = fs.Checkpoint()
        report = fs.process_keyword(
            FakeDirector([FakeMapsPage([], end_message=False)]), None, point,
            "Sikkim", "Mangan", "ICSE Schools", 1, 1)
        self.assertFalse(report["complete"])
        self.assertFalse(point.keyword_done("Sikkim", "Mangan",
                                            "ICSE Schools"))

    def test_a_completed_keyword_is_remembered_and_skipped(self):
        point = fs.Checkpoint()
        self.assertFalse(point.keyword_done("Karnataka", "Bengaluru",
                                            "CBSE Schools"))
        point.mark_keyword_done("Karnataka", "Bengaluru", "CBSE Schools")
        fresh = fs.Checkpoint()
        self.assertTrue(fresh.keyword_done("Karnataka", "Bengaluru",
                                           "CBSE Schools"))
        self.assertFalse(fresh.keyword_done("Karnataka", "Bengaluru",
                                            "ICSE Schools"))

    def test_an_unresolved_keyword_leaves_the_city_open(self):
        keywords = ["Private Schools"]
        page = FakeMapsPage([maps_entry(n) for n in range(1, 4)],
                            end_message=False)
        director = FakeDirector([page], FakeGeminiPage(ANSWER))
        gemini = fs.Gemini(director)
        point = fs.Checkpoint()
        report = fs.process_city(director, gemini, point, "Karnataka",
                                 "Bengaluru", keywords)
        self.assertFalse(report["complete"])
        self.assertFalse(point.keyword_done("Karnataka", "Bengaluru",
                                            keywords[0]))
        self.assertFalse(point.is_city_complete("Karnataka", "Bengaluru"))

    def test_a_school_already_in_the_master_is_not_researched_again(self):
        existing = school("Private Schools 1",
                          **{"School Address": "1 MG Road, Bengaluru 560011",
                             "PIN Code": "560011"})
        fs.append_records([existing])
        _report, events = self._city(["Private Schools"], per_keyword=2)
        asked = [v for k, v in events if k == "gemini"]
        self.assertNotIn("Private Schools 1", asked,
                         "a school already in the master was researched again")

    def test_the_raw_gemini_answer_is_kept_under_its_school(self):
        fs.KEEP_RAW_GEMINI_ANSWERS = True
        record = school("Kept Answer School")
        fs.save_gemini_answer("Karnataka", "Bengaluru", record,
                              ANSWER.format(name="Kept Answer School"))
        path = fs.gemini_school_path("Karnataka", "Bengaluru", record)
        with open(path, encoding="utf-8") as handle:
            body = handle.read()
        self.assertIn("SCHOOL: Kept Answer School", body)
        self.assertIn("School_Name:", body)


class GeminiPasteTests(SandboxTest):
    def test_the_prompt_is_pasted_with_control_v(self):
        page = FakeGeminiPage(ANSWER.format(name="Gupta Public School"))
        gemini = fs.Gemini(FakeBrowser([page, page]), separate=False)
        gemini.research(school("Gupta Public School",
                               **{"PIN Code": "560011"}))
        self.assertIn("Control+v", page.pressed)
        self.assertIn("Gupta Public School", page.clipboard)
        self.assertIn("Additional_School_Notes", page.clipboard,
                      "the whole prompt did not reach the clipboard")

    def test_the_box_is_checked_after_the_paste(self):
        page = FakeGeminiPage(ANSWER.format(name="Gupta Public School"))
        gemini = fs.Gemini(FakeBrowser([page, page]), separate=False)
        gemini.research(school("Gupta Public School",
                               **{"PIN Code": "560011"}))
        self.assertIn("Gupta Public School", page.box_text)

    def test_a_paste_that_does_nothing_falls_back_to_typing(self):
        page = FakeGeminiPage(ANSWER.format(name="Gupta Public School"))
        real_press = page.keyboard.press

        def swallow_paste(key):
            if key == "Control+v":
                page.pressed.append(key)
                return
            real_press(key)
        page.keyboard.press = swallow_paste
        gemini = fs.Gemini(FakeBrowser([page, page]), separate=False)
        values, _raw = gemini.research(school("Gupta Public School",
                                              **{"PIN Code": "560011"}))
        self.assertTrue(values)
        self.assertTrue(page.typed, "the prompt was never typed either")

    def test_the_wait_floor_is_fifteen_seconds(self):
        import ast
        with open(SCRIPT, encoding="utf-8") as handle:
            for node in ast.parse(handle.read()).body:
                if isinstance(node, ast.Assign) and getattr(
                        node.targets[0], "id", "") == "GEMINI_FIRST_READ_SECONDS":
                    self.assertGreaterEqual(ast.literal_eval(node.value), 15)
                    return
        self.fail("GEMINI_FIRST_READ_SECONDS is not set")

    def test_the_answer_comes_from_the_response_element_not_the_clipboard(self):
        """Ctrl+A/Ctrl+C on the Gemini page would also take the prompt,
        the interface and the previous school's answer."""
        with open(SCRIPT, encoding="utf-8") as handle:
            source = handle.read()
        gemini_part = source[source.index("class Gemini:"):]
        self.assertNotIn("Control+c", gemini_part,
                         "the Gemini page is being copied")
        self.assertNotIn("clipboard.readText", gemini_part,
                         "the Gemini answer is being read off the clipboard")
        self.assertIn("nodes[-1].inner_text()", gemini_part)

    def test_only_the_last_response_container_is_read(self):
        page = FakeGeminiPage(ANSWER.format(name="Gupta Public School"))
        page.body = ("Gemini\nNew chat\n"
                     "School_Name: WRONG SCHOOL FROM THE PAGE BODY\n")
        gemini = fs.Gemini(FakeBrowser([page, page]), separate=False)
        values, _raw = gemini.research(school("Gupta Public School",
                                              **{"PIN Code": "560011"}))
        self.assertEqual(values["School Name"], "Gupta Public School")


# ===========================================================================
# 19. the bugs a live run found
# ===========================================================================

class IdentityAndCompletionTests(SandboxTest):
    def test_a_school_with_only_a_name_still_has_a_key(self):
        """It had none, so it could never be found again."""
        bare = school("Propel International School",
                      **{"School Address": "", "PIN Code": ""})
        self.assertTrue(fs.match_keys(bare))
        fs.append_records([bare])
        written, skipped = fs.append_records([bare])
        self.assertEqual((written, skipped), (0, 1),
                         "a name-only school was saved twice")

    def test_enrichment_does_not_hide_a_saved_school(self):
        """Gemini replaces the short Maps address with a fuller one."""
        discovered = school("Father LeBlond School",
                            **{"School Address": "Bhimbhar"})
        fs.append_records([discovered])
        enriched = dict(discovered)
        enriched["School Address"] = "Bhimbhar, P.O - Madati, Darjeeling"
        self.assertEqual(fs.missing_from_master([enriched]), [],
                         "an enriched school looked missing from the master")

    def test_a_genuinely_missing_school_is_still_reported(self):
        fs.append_records([school("Known School",
                                  **{"School Address": "1 Road"})])
        stranger = school("Never Saved School", **{"School Address": "9 Road"})
        self.assertEqual(len(fs.missing_from_master([stranger])), 1)

    def test_two_branches_are_not_collapsed_by_the_weak_key(self):
        one = school("ORCHIDS The International School - Sarjapur Road",
                     **{"School Address": "Sarjapur Road", "PIN Code": "560035"})
        two = school("ORCHIDS The International School - Hennur Road",
                     **{"School Address": "Hennur Road", "PIN Code": "560043"})
        index = fs.SchoolIndex()
        index.add(one)
        index.add(two)
        self.assertEqual(len(index), 2)

    def test_a_city_closes_once_its_schools_are_in_the_master(self):
        keywords = ["Private Schools"]
        page = FakeMapsPage([maps_entry(n) for n in range(1, 4)])
        director = FakeDirector([page], FakeGeminiPage(ANSWER))
        gemini = fs.Gemini(director)
        gemini.research = lambda record: (
            fs.parse_answer(ANSWER.format(name=record["School Name"])),
            ANSWER)
        point = fs.Checkpoint()
        report = fs.process_city(director, gemini, point, "Karnataka",
                                 "Bengaluru", keywords)
        self.assertTrue(report["complete"], report["note"])
        self.assertTrue(point.is_city_complete("Karnataka", "Bengaluru"))
        self.assertLessEqual(director.max_open, 1)

    def test_schools_left_by_an_interrupted_run_are_picked_up(self):
        """The killed run discovered them but never researched them."""
        left_over = [school("Left Over %d" % n,
                            **{"School Address": "%d Old Road" % n})
                     for n in (1, 2)]
        fs.save_temp_excel("Karnataka", "Bengaluru", left_over)
        page = FakeMapsPage([maps_entry(9, name="Fresh School")])
        director = FakeDirector([page], FakeGeminiPage(ANSWER))
        asked = []
        gemini = fs.Gemini(director)

        def research(record):
            asked.append(record["School Name"])
            return (fs.parse_answer(ANSWER.format(name=record["School Name"])),
                    ANSWER)
        gemini.research = research
        fs.process_keyword(director, gemini, fs.Checkpoint(),
                           "Karnataka", "Bengaluru", "Private Schools", 1, 1)
        self.assertIn("Left Over 1", asked,
                      "a school from the interrupted run was abandoned")
        self.assertIn("Fresh School", asked)


class ClosingSweepTests(SandboxTest):
    def test_leftovers_are_finished_even_when_every_keyword_is_done(self):
        """The flags said the work was done; the sheet said otherwise."""
        left_over = [school("Stranded %d" % n,
                            **{"School Address": "%d Old Road" % n})
                     for n in (1, 2)]
        fs.save_temp_excel("Karnataka", "Bengaluru", left_over)
        point = fs.Checkpoint()
        point.mark_keyword_done("Karnataka", "Bengaluru", "Private Schools")

        asked = []
        director = FakeDirector([FakeMapsPage([])], FakeGeminiPage(ANSWER))
        gemini = fs.Gemini(director)

        def research(record):
            asked.append(record["School Name"])
            return (fs.parse_answer(ANSWER.format(name=record["School Name"])),
                    ANSWER)
        gemini.research = research

        report = fs.process_city(director, gemini, point, "Karnataka",
                                 "Bengaluru", ["Private Schools"])
        self.assertEqual(sorted(asked), ["Stranded 1", "Stranded 2"])
        self.assertTrue(report["complete"], report["note"])
        self.assertTrue(point.is_city_complete("Karnataka", "Bengaluru"))

    def test_a_city_with_nothing_left_over_needs_no_sweep(self):
        page = FakeMapsPage([maps_entry(n) for n in range(1, 3)])
        director = FakeDirector([page], FakeGeminiPage(ANSWER))
        gemini = fs.Gemini(director)
        gemini.research = lambda record: (
            fs.parse_answer(ANSWER.format(name=record["School Name"])), ANSWER)
        point = fs.Checkpoint()
        report = fs.process_city(director, gemini, point,
                                 "Karnataka", "Bengaluru", ["Private Schools"])
        self.assertTrue(report["complete"], report["note"])


class OneBrowserAtATimeTests(unittest.TestCase):
    """The director: one Chromium, never two."""

    def test_asking_for_the_other_browser_closes_the_first(self):
        director = FakeDirector()
        director.maps()
        director.gemini()
        director.maps()
        self.assertEqual(director.events,
                         [("open", "maps"), ("close", "maps"),
                          ("open", "gemini"), ("close", "gemini"),
                          ("open", "maps")])
        self.assertEqual(director.max_open, 1)

    def test_asking_twice_for_the_same_one_does_not_reopen_it(self):
        director = FakeDirector()
        first = director.maps()
        self.assertIs(director.maps(), first)
        self.assertEqual(director.opened["maps"], 1)

    def test_the_real_director_refuses_to_open_while_one_will_not_close(self):
        director = fs.BrowserDirector()
        director.current = types.SimpleNamespace(
            alive=lambda: True,
            close=lambda: (_ for _ in ()).throw(RuntimeError("stuck")))
        director.kind = "maps"
        with self.assertRaises(fs.BrowserUnavailable) as caught:
            director.open("gemini")
        self.assertIn("would not close", str(caught.exception))

    def test_the_two_browsers_keep_their_own_profiles(self):
        director = fs.BrowserDirector()
        try:
            self.assertNotEqual(director.profile_of("maps"),
                                director.profile_of("gemini"))
        finally:
            shutil.rmtree(director._scratch or "/nonexistent",
                          ignore_errors=True)

    def test_the_profiles_are_always_outside_the_project(self):
        director = fs.BrowserDirector()
        try:
            for kind in ("maps", "gemini"):
                where = os.path.abspath(director.profile_of(kind))
                self.assertFalse(where.startswith(
                    os.path.abspath(fs.OUTPUT_DIR)),
                    "%s profile is inside the project: %s" % (kind, where))
        finally:
            shutil.rmtree(director._scratch or "/nonexistent",
                          ignore_errors=True)

    def test_profiles_left_by_a_killed_run_are_swept_up(self):
        old = tempfile.mkdtemp(prefix=fs.BrowserDirector.SCRATCH_PREFIX)
        fresh = tempfile.mkdtemp(prefix=fs.BrowserDirector.SCRATCH_PREFIX)
        try:
            os.utime(old, (0, 0))          # older than the one hour threshold
            fs.sweep_old_browser_profiles()
            self.assertFalse(os.path.exists(old),
                             "a profile from a killed run was left behind")
            self.assertTrue(os.path.exists(fresh),
                            "a profile in use right now was deleted")
        finally:
            for path in (old, fresh):
                shutil.rmtree(path, ignore_errors=True)

    def test_shutdown_removes_the_temporary_profiles(self):
        if True:
            director = fs.BrowserDirector()
            folder = director.scratch()
            os.makedirs(os.path.join(folder, "maps"), exist_ok=True)
            self.assertTrue(os.path.isdir(folder))
            director.shutdown()
            self.assertFalse(os.path.exists(folder),
                             "the browser profiles were left on disk")

    def test_closing_a_browser_does_not_stop_the_driver(self):
        director = fs.BrowserDirector()
        director._play = types.SimpleNamespace(stopped=False)
        director.current = types.SimpleNamespace(alive=lambda: False,
                                                 close=lambda: None)
        director.kind = "maps"
        with unittest.mock.patch.object(fs, "profile_is_free",
                                        return_value=(True, 0)):
            ok, _detail = director.close()
        self.assertTrue(ok)
        self.assertIsNotNone(director._play,
                             "the shared driver was thrown away")

    def test_a_fresh_profile_is_never_waited_on(self):
        """It is made for this run, so no earlier run can be holding it."""
        director = fs.BrowserDirector()
        director._play = "driver"
        try:
            with unittest.mock.patch.object(fs, "profile_is_free",
                                            return_value=(False, 30)), \
                    unittest.mock.patch.object(
                        fs, "Browser",
                        return_value=FakeBrowser([FakeMapsPage([])])):
                director.open("maps")          # must not raise
        finally:
            shutil.rmtree(director._scratch or "/nonexistent",
                          ignore_errors=True)


class WindowLineTests(unittest.TestCase):
    LINE = ('     0x1e00004 "Google Maps - Google Chrome for Testing": '
            '("com.anthropic.Claude (/tmp/win/prof_maps)" "Com.anthropic")  '
            '912x1028+76+4')

    def test_a_real_xwininfo_line_is_understood(self):
        found = fs._WINDOW_LINE.match(self.LINE)
        self.assertIsNotNone(found)
        window_id, title, klass = found.groups()
        self.assertEqual(window_id, "0x1e00004")
        self.assertIn("Google Maps", title)
        self.assertIn("prof_maps", klass)

    def test_a_window_is_found_by_its_profile_directory(self):
        windows = [
            {"id": "0x100", "title": "Maps", "profile": "/tmp/prof_maps"},
            {"id": "0x200", "title": "Gemini", "profile": "/tmp/prof_gem"},
        ]
        with unittest.mock.patch.object(fs, "desktop_windows",
                                        return_value=windows):
            found = fs.window_of_profile(":0", "/tmp/auth", "/tmp/prof_gem")
            self.assertEqual(found["id"], "0x200")
            self.assertIsNone(
                fs.window_of_profile(":0", "/tmp/auth", "/tmp/nothing"))


class DiskUseTests(SandboxTest):
    def test_the_capture_text_is_packed_and_still_reads_the_same(self):
        page = FakeMapsPage([maps_entry(n) for n in range(1, 41)],
                            per_scroll=3, visible=6)
        result = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                   "Karnataka", "Bengaluru")
        before = os.path.getsize(result.txt_path)
        wanted = [r["School Name"] for r in result.records]

        fs.COMPRESS_CAPTURE_FILES = True
        packed_before, packed_after = fs.pack_capture_file(result.txt_path)
        if packed_before == packed_after:
            self.skipTest("the capture file was too small to be worth packing")
        self.assertFalse(os.path.exists(result.txt_path),
                         "the plain file was left behind as well")
        self.assertTrue(os.path.exists(result.txt_path + ".gz"))
        self.assertLess(os.path.getsize(result.txt_path + ".gz"), before)
        again = fs.records_from_txt(result.txt_path, "Karnataka", "Bengaluru")
        self.assertEqual([r["School Name"] for r in again], wanted)

    def test_packing_keeps_every_byte(self):
        folder = fs.city_dir("Karnataka", "Bengaluru")
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, "big.txt")
        body = ("the same copied line over and over\n" * 4000)
        with open(path, "w", encoding="utf-8") as handle:
            handle.write(body)
        fs.COMPRESS_CAPTURE_FILES = True
        fs.pack_capture_file(path)
        self.assertEqual(fs.read_capture(path), body)

    def test_a_resumed_keyword_keeps_both_halves(self):
        folder = fs.city_dir("Karnataka", "Bengaluru")
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, "two.txt")
        fs.COMPRESS_CAPTURE_FILES = True
        with open(path, "w", encoding="utf-8") as handle:
            handle.write("first run\n" * 8000)
        fs.pack_capture_file(path)
        with open(path, "w", encoding="utf-8") as handle:
            handle.write("second run\n" * 8000)
        fs.pack_capture_file(path)
        text = fs.read_capture(path)
        self.assertIn("first run", text)
        self.assertIn("second run", text)

    def test_a_small_file_is_left_alone(self):
        folder = fs.city_dir("Karnataka", "Bengaluru")
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, "small.txt")
        with open(path, "w", encoding="utf-8") as handle:
            handle.write("tiny")
        fs.COMPRESS_CAPTURE_FILES = True
        fs.pack_capture_file(path)
        self.assertTrue(os.path.exists(path))
        self.assertFalse(os.path.exists(path + ".gz"))

    def test_there_is_one_log_file_not_one_per_day(self):
        with open(SCRIPT, encoding="utf-8") as handle:
            source = handle.read()
        self.assertNotIn("school_%s.log", source)
        self.assertIn("LOG_FILE", source)
        self.assertTrue(fs.LOG_FILE.endswith(".log"))

    def test_lines_older_than_the_limit_are_dropped(self):
        folder = tempfile.mkdtemp(prefix="school_log_")
        try:
            path = os.path.join(folder, "find_school.log")
            today = datetime.date.today()
            old_day = (today - datetime.timedelta(days=40)).isoformat()
            recent = (today - datetime.timedelta(days=1)).isoformat()
            with open(path, "w", encoding="utf-8") as handle:
                handle.write("%s 10:00:00 INFO    ancient\n" % old_day)
                handle.write("    a wrapped line of the ancient message\n")
                handle.write("%s 10:00:00 INFO    yesterday\n" % recent)
                handle.write("%s 11:00:00 INFO    today\n" % today.isoformat())
            kept, dropped = fs.trim_log(path, keep_days=10)
            self.assertEqual((kept, dropped), (2, 2))
            with open(path, encoding="utf-8") as handle:
                body = handle.read()
            self.assertNotIn("ancient", body)
            self.assertNotIn("wrapped line", body)
            self.assertIn("yesterday", body)
            self.assertIn("today", body)
        finally:
            shutil.rmtree(folder, ignore_errors=True)

    def test_a_fresh_log_is_left_alone(self):
        folder = tempfile.mkdtemp(prefix="school_log_")
        try:
            path = os.path.join(folder, "find_school.log")
            today = datetime.date.today().isoformat()
            with open(path, "w", encoding="utf-8") as handle:
                handle.write("%s 10:00:00 INFO    hello\n" % today)
            before = open(path, encoding="utf-8").read()
            _kept, dropped = fs.trim_log(path, keep_days=10)
            self.assertEqual(dropped, 0)
            self.assertEqual(open(path, encoding="utf-8").read(), before)
        finally:
            shutil.rmtree(folder, ignore_errors=True)

    def test_keeping_the_log_for_ever_can_be_switched_on(self):
        folder = tempfile.mkdtemp(prefix="school_log_")
        try:
            path = os.path.join(folder, "find_school.log")
            with open(path, "w", encoding="utf-8") as handle:
                handle.write("1999-01-01 10:00:00 INFO    very old\n")
            self.assertEqual(fs.trim_log(path, keep_days=0), (0, 0))
            self.assertIn("very old", open(path, encoding="utf-8").read())
        finally:
            shutil.rmtree(folder, ignore_errors=True)

    def test_the_old_per_city_backup_scheme_is_still_gone(self):
        """Backup/ is back, but only for FINISHED cycle files.

        The thing that was removed was a recovery workbook per city, kept
        to a rolling count. What exists now is one permanent, verified
        copy of each completed 1000-record cycle - and the master, the
        cycle files and the journal are still the safety net underneath.
        """
        with open(SCRIPT, encoding="utf-8") as handle:
            source = handle.read()
        for name in ("BACKUP_KEEP", "prune_old_backups", "backup_city"):
            self.assertNotIn(name, source, name)
        self.assertIn("reconcile_master_and_cycles", source)
        self.assertIn("COMMIT_JOURNAL", source)
        self.assertIn("backup_completed_cycle", source)

    def test_the_browser_cache_is_capped(self):
        with open(SCRIPT, encoding="utf-8") as handle:
            source = handle.read()
        self.assertIn("--disk-cache-size=", source)
        self.assertIn("--media-cache-size=", source)
        self.assertIn("--disable-breakpad", source)
        self.assertLessEqual(fs.BROWSER_CACHE_MB, 100)


class NoCapTests(SandboxTest):
    def test_a_long_list_is_not_cut_short_by_a_cap(self):
        """200 entries, far past the old 120-round cap, reach the end."""
        page = FakeMapsPage([maps_entry(n) for n in range(1, 201)],
                            per_scroll=5, visible=5)
        result = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                   "Karnataka", "Bengaluru")
        self.assertTrue(result.complete)
        self.assertEqual(result.reason,
                         "end-of-list message in the copied text")
        self.assertEqual(len(result.records), 200)
        self.assertGreater(result.snapshots, 40)

    def test_a_list_that_never_ends_still_stops_on_its_own(self):
        """No cap must not mean an endless loop."""
        page = FakeMapsPage([maps_entry(n) for n in range(1, 11)],
                            per_scroll=5, visible=10, end_message=False)
        result = fs.collect_search(FakeBrowser([page]), "Private Schools",
                                   "Karnataka", "Bengaluru")
        self.assertFalse(result.complete)
        self.assertIn("no new school", result.reason)


class ContactRequiredTests(SandboxTest):
    """A school with no phone and no email does not earn a row."""

    @staticmethod
    def bare(name, **extra):
        record = fs.blank_record()
        record.update({"School Name": name, "City": "Bengaluru",
                       "State": "Karnataka", "Country": "India"})
        record.update(extra)
        return record

    def test_a_phone_number_is_enough(self):
        self.assertTrue(fs.has_contact(
            self.bare("A", **{"School Official Phone Number": "9900112233"})))

    def test_an_email_is_enough(self):
        self.assertTrue(fs.has_contact(
            self.bare("A", **{"School Official Email": "head@school.edu.in"})))

    def test_the_principals_details_count_too(self):
        self.assertTrue(fs.has_contact(
            self.bare("A", **{"Principal Official Email": "p@school.in"})))
        self.assertTrue(fs.has_contact(
            self.bare("A", **{"Principal Office Phone Number": "08012345678"})))

    def test_neither_is_not_enough(self):
        self.assertFalse(fs.has_contact(self.bare("A")))
        self.assertFalse(fs.has_contact(
            self.bare("A", **{"School Address": "12 MG Road",
                              "School Website": "https://a.edu.in"})))

    def test_no_data_does_not_count_as_contact(self):
        self.assertFalse(fs.has_contact(
            self.bare("A", **{"School Official Phone Number": "NO_DATA",
                              "School Official Email": "NO_DATA"})))

    def test_a_number_that_is_not_a_phone_number_does_not_count(self):
        self.assertFalse(fs.has_contact(
            self.bare("A", **{"School Official Phone Number": "560011"})))

    def test_it_is_not_written_to_the_final_files(self):
        point = fs.Checkpoint()
        outcome = fs.research_and_save(None, point, "Karnataka", "Bengaluru",
                                       self.bare("No Contact School"), 1, 1)
        self.assertEqual(outcome, "no contact")
        self.assertEqual(fs.row_count(fs.master_path()), 0)
        self.assertEqual(fs.all_cycle_paths(), [])

    def test_one_with_contact_is_written(self):
        point = fs.Checkpoint()
        outcome = fs.research_and_save(
            None, point, "Karnataka", "Bengaluru",
            self.bare("Reachable School",
                      **{"School Official Phone Number": "9900112233"}), 1, 1)
        self.assertEqual(outcome, "saved")
        self.assertEqual(fs.row_count(fs.master_path()), 1)

    def test_it_is_not_researched_again_on_the_next_run(self):
        point = fs.Checkpoint()
        record = self.bare("No Contact School")
        fs.research_and_save(None, point, "Karnataka", "Bengaluru",
                             record, 1, 1)
        self.assertIn(fs.record_identity(record),
                      fs.Checkpoint().enriched_keys("Karnataka", "Bengaluru"))

    def test_a_skipped_school_does_not_keep_its_city_open(self):
        """It is left out on purpose, so it is not "missing"."""
        kept = self.bare("Reachable",
                         **{"School Official Phone Number": "9900112233"})
        fs.append_records([kept])
        skipped = self.bare("No Contact School")
        self.assertEqual(fs.missing_from_master([kept, skipped]), [])

    def test_a_reachable_school_that_is_missing_is_still_reported(self):
        stranger = self.bare("Never Saved",
                             **{"School Official Phone Number": "9876543210"})
        self.assertEqual(len(fs.missing_from_master([stranger])), 1)

    def test_a_whole_city_of_them_still_closes(self):
        page = FakeMapsPage([maps_entry(n, phone=False)
                             for n in range(1, 4)])
        director = FakeDirector([page], FakeGeminiPage(ANSWER))
        gemini = fs.Gemini(director)
        gemini.research = lambda record: ({}, "no fields")
        point = fs.Checkpoint()
        report = fs.process_city(director, gemini, point, "Karnataka",
                                 "Bengaluru", ["Private Schools"])
        self.assertTrue(report["complete"], report["note"])
        self.assertEqual(report["saved"], 0)
        self.assertEqual(report["no_contact"], 3)
        self.assertEqual(fs.row_count(fs.master_path()), 0)


class FlatTempFolderTests(SandboxTest):
    def test_no_folder_is_made_for_a_city(self):
        fs.save_temp_excel("Karnataka", "Bengaluru", [school("A")])
        fs.save_city_progress("Karnataka", "Bengaluru", [school("A")])
        for name in os.listdir(fs.TEMP_DIR):
            self.assertFalse(
                os.path.isdir(os.path.join(fs.TEMP_DIR, name)),
                "%s is a folder inside Temp_Data" % name)

    def test_the_city_is_in_the_file_name(self):
        path = fs.temp_excel_path("Delhi NCR", "Greater Noida")
        self.assertEqual(os.path.dirname(path), fs.TEMP_DIR)
        self.assertTrue(os.path.basename(path).startswith(
            "Delhi_NCR_Greater_Noida_"), os.path.basename(path))

    def test_two_cities_do_not_collide(self):
        one = fs.temp_excel_path("Karnataka", "Bengaluru")
        two = fs.temp_excel_path("Sikkim", "Rangpo")
        self.assertNotEqual(one, two)

    def test_only_this_city_s_files_are_cleaned_up(self):
        fs.save_temp_excel("Karnataka", "Bengaluru", [school("A")])
        fs.save_temp_excel("Sikkim", "Rangpo", [school("B")])
        fs.clear_city_temp("Karnataka", "Bengaluru")
        self.assertFalse(os.path.exists(
            fs.temp_excel_path("Karnataka", "Bengaluru")))
        self.assertTrue(os.path.exists(
            fs.temp_excel_path("Sikkim", "Rangpo")),
            "another city's working file was deleted")

    def test_a_city_with_a_slash_in_its_name_is_safe(self):
        path = fs.temp_excel_path("Andaman/Nicobar", "Car Nicobar")
        self.assertEqual(os.path.dirname(path), fs.TEMP_DIR)
        self.assertNotIn("/", os.path.basename(path))


class UserSettingsTests(unittest.TestCase):
    """One small block of settings at the top; the rest is internal."""

    USER = ("rank", "num_sys", "DOWN_PRESSES_PER_ROUND",
            "SCROLL_PAUSE_SECONDS", "GEMINI_URL",
            "GEMINI_FIRST_READ_SECONDS", "GEMINI_ANSWER_TIMEOUT",
            "GEMINI_RETRIES", "CYCLE_SIZE", "KEEP_CAPTURED_TEXT",
            "LOG_KEEP_DAYS")

    TECHNICAL = ("PAGE_SETTLE_SECONDS", "BROWSER_TIMEOUT_MS",
                 "CLICK_SETTLE_SECONDS", "TYPING_SETTLE_SECONDS",
                 "COPY_KEY_PAUSE_SECONDS", "COPY_CLIPBOARD_PAUSE_SECONDS",
                 "COPY_EMPTY_RETRIES", "DOWN_KEY_PAUSE_SECONDS",
                 "SCROLL_NO_PROGRESS_PRESSES",
                 "GEMINI_POLL_SECONDS", "GEMINI_SETTLE_SECONDS",
                 "GEMINI_INPUT_TIMEOUT", "INPUT_BOX_POLL_SECONDS",
                 "SCHOOL_DELAY_SECONDS", "CITY_DELAY_SECONDS",
                 "RETRY_BACKOFF_SECONDS", "MAX_BROWSER_RESTARTS",
                 "SHOW_WINDOW_WAIT_SECONDS", "WINDOW_CHECK_POLL_SECONDS",
                 "DISPLAY_CHECK_TIMEOUT", "WINDOW_LIST_TIMEOUT",
                 "SOCKET_PROBE_TIMEOUT", "PROFILE_RELEASE_WAIT_SECONDS",
                 "SCRATCH_SWEEP_AGE_SECONDS", "BROWSER_CACHE_MB")

    GONE = ("PERSIST_BROWSER_PROFILE", "GEMINI_SEPARATE_BROWSER",
            "ROLLOVER_ROWS", "MAX_RESULTS_PER_SEARCH",
            "SEARCH_DELAY_SECONDS", "KEEP_RAW_MAPS_TEXT", "STATE_DIR",
            "LOG_DIR", "KEEP_BROWSER_PROFILES", "BACKUP_KEEP",
            "backup_city",
            "MAX_SCROLL_ROUNDS", "SEPARATE_FOLDER_PER_RANK",
            "SEARCH_TIMEOUT_SECONDS")

    def positions(self):
        import ast
        with open(SCRIPT, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        where = {}
        for node in tree.body:
            if isinstance(node, ast.Assign):
                name = getattr(node.targets[0], "id", "")
                if name:
                    where.setdefault(name, node.lineno)
        return where

    def test_every_user_setting_exists(self):
        for name in self.USER:
            self.assertTrue(hasattr(fs, name), name)

    def test_the_user_settings_sit_together_at_the_top(self):
        where = self.positions()
        lines = []
        for name in self.USER:
            self.assertIn(name, where, "%s is not a plain top-level value"
                          % name)
            lines.append(where[name])
        self.assertLess(max(lines), 120,
                        "a user setting is far down the file")
        self.assertLess(max(lines) - min(lines), 60,
                        "the user settings are not in one block")

    def test_the_technical_ones_are_not_in_that_block(self):
        where = self.positions()
        top = max(where[name] for name in self.USER)
        for name in self.TECHNICAL:
            self.assertIn(name, where, name)
            self.assertGreater(where[name], top,
                               "%s is still in the user block" % name)

    def test_the_obsolete_settings_are_gone(self):
        with open(SCRIPT, encoding="utf-8") as handle:
            source = handle.read()
        for name in self.GONE:
            self.assertFalse(hasattr(fs, name), "%s still exists" % name)
            self.assertNotIn(name, source, "%s is still referenced" % name)

    def test_no_hard_coded_wait_is_left_in_the_code(self):
        import ast
        with open(SCRIPT, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            if getattr(node.func, "attr", "") == "sleep" and node.args:
                self.assertNotIsInstance(
                    node.args[0], ast.Constant,
                    "a sleep with a number in it at line %d" % node.lineno)

    def test_each_one_is_a_number_or_a_pair(self):
        for name in self.USER + self.TECHNICAL:
            value = getattr(fs, name)
            if isinstance(value, (bool, str)):
                continue
            if isinstance(value, tuple):
                self.assertEqual(len(value), 2, name)
                self.assertLessEqual(value[0], value[1], name)
            else:
                self.assertIsInstance(value, (int, float), name)
                self.assertGreaterEqual(value, 0, name)

    def test_nothing_but_the_end_of_list_can_finish_a_keyword(self):
        """No round cap, no clock - only Google's own message."""
        import ast
        with open(SCRIPT, encoding="utf-8") as handle:
            source = handle.read()
        self.assertNotIn("MAX_SCROLL_ROUNDS", source)
        self.assertNotIn("SEARCH_TIMEOUT_SECONDS", source)
        self.assertIn("END_OF_LIST_MARKERS", source)

    def test_giving_up_is_measured_in_presses_not_copies(self):
        """Changing DOWN_PRESSES_PER_ROUND must not end searches early."""
        self.assertGreaterEqual(fs.SCROLL_NO_PROGRESS_PRESSES, 120)
        rounds = fs.SCROLL_NO_PROGRESS_PRESSES / max(
            fs.DOWN_PRESSES_PER_ROUND, 1)
        self.assertGreaterEqual(
            rounds, 15,
            "a search would give up after only %.0f copies" % rounds)


# ===========================================================================
# 27. cleaning up temporary files: success cleans, failure is preserved
# ===========================================================================

class CleanupTests(SandboxTest):
    """SUCCESS -> CLEAN UP. UNRESOLVED FAILURE -> PRESERVE FOR RESUME."""

    def _run_city(self, keywords=("Private Schools",), cards=3,
                  end_message=True, gemini_fails=False, phone=True,
                  checkpoint=None, answer=ANSWER):
        keywords = list(keywords)
        pages = {k: FakeMapsPage(
            [maps_entry(n, name="%s %d" % (k, n), phone=phone)
             for n in range(1, cards + 1)], end_message=end_message)
            for k in keywords}
        director = FakeDirector([pages[keywords[0]]],
                               FakeGeminiPage(answer))
        real_collect = fs.collect_search

        def pick(browser, keyword, state, city, limit=None):
            browser.pages[0] = pages[keyword]
            return real_collect(browser, keyword, state, city, limit=limit)

        gemini = fs.Gemini(director)
        if gemini_fails:
            def explode(record):
                raise fs.GeminiUnavailable("no answer came back")
            gemini.research = explode
        else:
            gemini.research = lambda record: (
                fs.parse_answer(answer.format(name=record["School Name"])),
                answer)
        fs.collect_search = pick
        try:
            return fs.process_city(director, gemini,
                                   checkpoint or fs.Checkpoint(),
                                   "Karnataka", "Bengaluru", keywords)
        finally:
            fs.collect_search = real_collect

    def _temp_files(self):
        return sorted(os.listdir(fs.TEMP_DIR)) \
            if os.path.isdir(fs.TEMP_DIR) else []

    # ---------------------------------------------------------- success --
    def test_temp_data_is_empty_after_a_city_finishes(self):
        report = self._run_city()
        self.assertTrue(report["complete"], report.get("note"))
        self.assertEqual(self._temp_files(), [],
                         "Temp_Data still holds %s" % self._temp_files())

    def test_a_saved_school_s_answer_file_goes_at_once(self):
        fs.KEEP_RAW_GEMINI_ANSWERS = True
        record = school("Verified School")
        fs.save_gemini_answer("Karnataka", "Bengaluru", record, "raw text")
        path = fs.gemini_school_path("Karnataka", "Bengaluru", record)
        self.assertTrue(os.path.exists(path))
        fs.clear_school_temp("Karnataka", "Bengaluru", record)
        self.assertFalse(os.path.exists(path),
                         "the raw answer outlived the saved school")

    def test_every_school_keeps_its_own_answer_file(self):
        fs.KEEP_RAW_GEMINI_ANSWERS = True
        one, two = school("First School"), school("Second School")
        fs.save_gemini_answer("Karnataka", "Bengaluru", one, "first")
        fs.save_gemini_answer("Karnataka", "Bengaluru", two, "second")
        self.assertNotEqual(
            fs.gemini_school_path("Karnataka", "Bengaluru", one),
            fs.gemini_school_path("Karnataka", "Bengaluru", two))
        fs.clear_school_temp("Karnataka", "Bengaluru", one)
        self.assertFalse(os.path.exists(
            fs.gemini_school_path("Karnataka", "Bengaluru", one)))
        self.assertTrue(os.path.exists(
            fs.gemini_school_path("Karnataka", "Bengaluru", two)),
            "clearing one school took another school's answer with it")

    def test_a_finished_keyword_s_copied_text_is_removed(self):
        self._run_city(keywords=["Private Schools", "CBSE Schools"])
        left = [n for n in self._temp_files() if n.endswith((".txt", ".gz"))]
        self.assertEqual(left, [], "copied text survived: %s" % left)

    def test_the_working_sheet_goes_when_the_city_closes(self):
        self._run_city()
        self.assertFalse(os.path.exists(
            fs.temp_excel_path("Karnataka", "Bengaluru")))
        self.assertFalse(os.path.exists(
            fs.city_file("Karnataka", "Bengaluru", "schools.json")))

    # ----------------------------------------------- the no-contact skip --
    def test_a_no_contact_school_is_a_finished_skip_not_a_failure(self):
        report = self._run_city(phone=False, answer=NO_CONTACT_ANSWER)
        self.assertTrue(report["complete"],
                        "a no-contact school kept the city open: %s"
                        % report.get("note"))
        self.assertEqual(report["failed"], 0)
        self.assertEqual(report["saved"], 0)
        self.assertEqual(report["no_contact"], 3)
        self.assertEqual(self._temp_files(), [],
                         "left behind %s" % self._temp_files())

    def test_a_no_contact_school_reaches_no_final_workbook(self):
        self._run_city(phone=False, answer=NO_CONTACT_ANSWER)
        self.assertEqual(fs.row_count(fs.master_path()), 0)
        self.assertEqual(fs.row_count(fs.cycle_path(0)), 0)

    # ---------------------------------------------------------- failures --
    def test_an_interrupted_gemini_preserves_everything(self):
        report = self._run_city(gemini_fails=True)
        self.assertFalse(report["complete"])
        self.assertTrue(os.path.exists(
            fs.temp_excel_path("Karnataka", "Bengaluru")),
            "the working sheet was deleted with schools still to research")
        text = [n for n in self._temp_files() if n.endswith((".txt", ".gz"))]
        self.assertTrue(text, "the copied text was deleted after a failure")

    def test_a_failed_school_keeps_its_raw_answer(self):
        fs.KEEP_RAW_GEMINI_ANSWERS = True
        record = school("Unresolved School")
        fs.save_gemini_answer("Karnataka", "Bengaluru", record, "half an answer")
        point = fs.Checkpoint()
        point.record_failure(record, "the answer never arrived")
        self.assertTrue(os.path.exists(
            fs.gemini_school_path("Karnataka", "Bengaluru", record)),
            "a failed school's answer was cleaned up")

    def test_an_incomplete_search_keeps_its_copied_text(self):
        report = self._run_city(end_message=False)
        self.assertFalse(report["complete"])
        text = [n for n in self._temp_files() if n.endswith((".txt", ".gz"))]
        self.assertTrue(text,
                        "the text of a search that never reached its end "
                        "was deleted")

    def test_the_keyword_sweep_is_never_reached_without_completion(self):
        """The sweep must not be in a finally: block."""
        import ast
        with open(SCRIPT, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        for node in ast.walk(tree):
            if not isinstance(node, ast.Try):
                continue
            for inner in ast.walk(ast.Module(body=node.finalbody,
                                             type_ignores=[])):
                if (isinstance(inner, ast.Call)
                        and getattr(inner.func, "id", "") in
                        ("sweep_completed_keyword", "clear_city_temp",
                         "clear_school_temp")):
                    self.fail("cleanup runs from a finally: block")

    # ------------------------------------------------------------ resume --
    def test_an_interrupted_city_resumes_without_collecting_again(self):
        self._run_city(gemini_fails=True)
        rows = fs.load_temp_excel("Karnataka", "Bengaluru")
        self.assertEqual(len(rows), 3,
                         "the discovered schools were lost")
        self.assertEqual(len(fs.missing_from_master(rows)), 3,
                         "the resume would skip schools that never saved")

    def test_the_second_run_finishes_what_the_first_left(self):
        point = fs.Checkpoint()
        self._run_city(gemini_fails=True, checkpoint=point)
        self.assertTrue(os.path.exists(
            fs.temp_excel_path("Karnataka", "Bengaluru")))
        report = self._run_city(checkpoint=fs.Checkpoint())
        self.assertTrue(report["complete"], report.get("note"))
        self.assertEqual(self._temp_files(), [],
                         "the finished city left %s" % self._temp_files())

    # ----------------------------------------------------------- startup --
    def test_startup_removes_leftovers_from_a_finished_city(self):
        fs.append_records([school("Done School",
                                  **{"PIN Code": "560001",
                                     "School Address": "1 Road"})])
        fs.save_temp_excel("Karnataka", "Bengaluru",
                           [school("Done School",
                                   **{"PIN Code": "560001",
                                      "School Address": "1 Road"})])
        fs.dump_text("Karnataka", "Bengaluru", "private_schools.txt", "old")
        point = fs.Checkpoint()
        point.mark_city_complete("Karnataka", "Bengaluru", 1, 0)
        fs.sweep_completed_cities(point)
        self.assertEqual(self._temp_files(), [],
                         "leftovers survived: %s" % self._temp_files())

    def test_startup_keeps_the_files_of_a_city_still_open(self):
        fs.save_temp_excel("Karnataka", "Bengaluru", [school("Open School")])
        fs.dump_text("Karnataka", "Bengaluru", "private_schools.txt", "text")
        before = self._temp_files()
        fs.sweep_completed_cities(fs.Checkpoint())
        self.assertEqual(self._temp_files(), before,
                         "an unfinished city's files were cleaned up")

    def test_startup_refuses_a_complete_city_whose_rows_are_not_saved(self):
        """The second guard: the checkpoint alone is not enough."""
        fs.save_temp_excel("Karnataka", "Bengaluru",
                           [school("Never Saved School")])
        fs.dump_text("Karnataka", "Bengaluru", "private_schools.txt", "text")
        point = fs.Checkpoint()
        point.mark_city_complete("Karnataka", "Bengaluru", 0, 0)
        before = self._temp_files()
        result = fs.sweep_completed_cities(point)
        self.assertEqual(self._temp_files(), before,
                         "files were removed although a school never saved")
        self.assertTrue(result["kept"], "the refusal was not reported")

    def test_startup_leaves_another_city_alone(self):
        fs.save_temp_excel("Karnataka", "Bengaluru", [school("A")])
        fs.save_temp_excel("Sikkim", "Rangpo", [school("B")])
        point = fs.Checkpoint()
        point.mark_city_complete("Karnataka", "Bengaluru", 0, 0)
        fs.sweep_completed_cities(point)
        self.assertTrue(os.path.exists(fs.temp_excel_path("Sikkim", "Rangpo")),
                        "an unrelated city's working sheet was deleted")

    # ------------------------------------------- what is never deleted --
    def test_the_permanent_files_are_never_touched(self):
        fs.append_records([school("Permanent School",
                                  **{"PIN Code": "560001",
                                     "School Address": "9 Road"})])
        point = fs.Checkpoint()
        point.mark_city_complete("Karnataka", "Bengaluru", 1, 0)
        fs.clear_city_temp("Karnataka", "Bengaluru")
        fs.sweep_completed_cities(point)
        self.assertTrue(os.path.exists(fs.master_path()))
        self.assertTrue(os.path.exists(fs.cycle_path(0)))
        self.assertTrue(os.path.exists(point.path))
        self.assertEqual(fs.row_count(fs.master_path()), 1)

    def test_the_cleanup_only_ever_touches_temp_data(self):
        import ast
        with open(SCRIPT, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        for node in ast.walk(tree):
            if not (isinstance(node, ast.FunctionDef)
                    and node.name in ("clear_school_temp",
                                      "sweep_completed_keyword",
                                      "clear_city_temp",
                                      "sweep_completed_cities",
                                      "school_temp_files")):
                continue
            for inner in ast.walk(node):
                name = getattr(inner.func, "id", "") if isinstance(
                    inner, ast.Call) else ""
                self.assertNotIn(name, ("master_path", "cycle_path",
                                        "journal_path"),
                                 "%s reaches a permanent file" % node.name)

    def test_an_older_run_s_shared_answers_file_is_still_swept(self):
        """Earlier versions wrote one gemini_answers.txt per city. The
        sweep matches on the city's file name prefix, so it goes too."""
        fs.append_records([school("Legacy School",
                                  **{"PIN Code": "560001",
                                     "School Address": "2 Road"})])
        legacy = fs.city_file("Karnataka", "Bengaluru", "gemini_answers.txt")
        with open(legacy, "w", encoding="utf-8") as handle:
            handle.write("an answer from an older run")
        point = fs.Checkpoint()
        point.mark_city_complete("Karnataka", "Bengaluru", 1, 0)
        fs.sweep_completed_cities(point)
        self.assertFalse(os.path.exists(legacy),
                         "the old shared answers file was left behind")

    def test_the_setting_is_off_by_default(self):
        with open(SCRIPT, encoding="utf-8") as handle:
            source = handle.read()
        self.assertIn("KEEP_CAPTURED_TEXT = False", source)



# ===========================================================================
# 28. a window appears ONLY with --show
# ===========================================================================

class OnlyShowIsVisibleTests(unittest.TestCase):
    """The rule the operator asked for, with no exceptions."""

    def setUp(self):
        self.kept = {name: os.environ.get(name)
                     for name in ("DISPLAY", "XAUTHORITY")}

    def tearDown(self):
        for name, value in self.kept.items():
            if value is None:
                os.environ.pop(name, None)
            else:
                os.environ[name] = value

    def _under_xvfb(self):
        os.environ["DISPLAY"] = ":99"
        os.environ["XAUTHORITY"] = "/tmp/xvfb-run.ABC123/Xauthority"

    def _on_a_desktop(self):
        os.environ["DISPLAY"] = ":0"
        os.environ["XAUTHORITY"] = "/run/user/1000/.Xauthority"

    def test_without_show_the_desktop_is_never_looked_up(self):
        """find_desktop_display() is what puts a window on :0."""
        import ast
        with open(SCRIPT, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        for node in ast.walk(tree):
            if not (isinstance(node, ast.ClassDef) and node.name == "Browser"):
                continue
            for call in ast.walk(node):
                if not (isinstance(call, ast.Call)
                        and getattr(call.func, "id", "")
                        == "find_desktop_display"):
                    continue
                # it must sit under "if self.show"
                guarded = False
                for branch in ast.walk(node):
                    if not isinstance(branch, ast.If):
                        continue
                    test = ast.dump(branch.test)
                    if "show" not in test:
                        continue
                    if any(call is inner for inner in ast.walk(branch)):
                        guarded = True
                self.assertTrue(
                    guarded,
                    "the desktop is looked up outside an 'if self.show'")

    def test_the_message_says_hidden_without_show(self):
        self._under_xvfb()
        line = fs.where_the_window_goes(False)
        self.assertIn("hidden", line.lower())
        self.assertNotIn("VISIBLE", line)

    def test_the_message_says_visible_with_show(self):
        self._on_a_desktop()
        self.assertIn("VISIBLE", fs.where_the_window_goes(True))

    def test_headless_is_reported_as_no_window(self):
        self._on_a_desktop()
        self.assertIn("no window",
                      fs.where_the_window_goes(False, headless=True))

    def test_no_show_on_a_desktop_runs_headless(self):
        self._on_a_desktop()
        browser = fs.Browser(show=False)
        self.assertEqual(browser.mode, "headless",
                         "a window would be drawn on the operator's desktop")

    def test_show_is_always_headful(self):
        self._on_a_desktop()
        self.assertEqual(fs.Browser(show=True).mode, "headful")

    def test_under_xvfb_without_show_nothing_reaches_the_desktop(self):
        """Headful is fine there - the screen is virtual and unwatched."""
        self._under_xvfb()
        browser = fs.Browser(show=False)
        self.assertEqual(browser.display, "",
                         "a display was chosen before the browser started")
        self.assertFalse(browser.show)

    def test_the_orphan_sweep_needs_both_conditions(self):
        """A browser AND one of this script's own profile folders."""
        import inspect
        source = inspect.getsource(fs.this_script_browsers)
        self.assertIn("--user-data-dir=", source,
                      "the profile must be read from the actual option")
        self.assertIn("SCRATCH_PREFIX", source)
        self.assertIn("basename", source,
                      "the program itself must be checked, not the whole "
                      "command line - otherwise a shell that mentions the "
                      "path is killed")

    def test_a_process_that_only_mentions_the_path_is_safe(self):
        """The bug this guards: a shell whose command line names a profile.

        It was found the hard way - the sweep killed the terminal it ran
        in, because `--user-data-dir=/tmp/school_browsers_x` appeared in
        that shell's own command line.
        """
        import subprocess, time
        fake = subprocess.Popen(
            ["/bin/sh", "-c",
             "sleep 20 --user-data-dir=%sfake/Default"
             % fs.BrowserDirector.SCRATCH_PREFIX])
        try:
            time.sleep(0.3)
            self.assertNotIn(fake.pid, fs.this_script_browsers(),
                             "a shell was mistaken for a browser")
        finally:
            fake.kill()
            fake.wait()

    def test_a_real_browser_on_this_script_s_profile_is_found(self):
        import subprocess, time, os, shutil
        folder = "/tmp/%sunittest" % fs.BrowserDirector.SCRATCH_PREFIX
        link = os.path.join(folder, "chromium")
        os.makedirs(folder, exist_ok=True)
        shutil.copy("/bin/sh", link)
        # sh keeps running and ignores the trailing option, so the process
        # is still alive to be found - unlike sleep, which rejects it
        fake = subprocess.Popen([link, "-c", "sleep 20", "chromium",
                                 "--user-data-dir=%s/Default" % folder])
        try:
            time.sleep(0.3)
            self.assertIn(fake.pid, fs.this_script_browsers(),
                          "a real orphan browser was not found")
        finally:
            fake.kill()
            fake.wait()
            shutil.rmtree(folder, ignore_errors=True)

    def test_a_browser_that_rewrote_its_argv_is_still_found(self):
        """Chromium puts its WHOLE command line into argv[0].

        Found live: the sweep saw nothing while a browser was running,
        because basename(argv[0]) was "maps --remote-debugging-pipe
        about:blank" rather than "chrome".
        """
        import inspect
        source = inspect.getsource(fs.this_script_browsers)
        self.assertIn(".split()", source,
                      "argv[0] is used whole, so a running browser is "
                      "never matched")

    def test_the_wayland_guard_is_on_the_quiet_path_too(self):
        """Without --show the window must be pinned to the X screen.

        xvfb-run replaces DISPLAY but leaves WAYLAND_DISPLAY, so Chromium
        can find the real compositor by itself and draw on the desktop.
        """
        import ast
        with open(SCRIPT, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        guarded = False
        for node in ast.walk(tree):
            if not isinstance(node, ast.If):
                continue
            if "show" not in ast.dump(node.test):
                continue
            for branch in node.orelse:
                body = ast.dump(branch)
                if ("WAYLAND_DISPLAY" in body
                        and "ozone-platform=x11" in body):
                    guarded = True
        self.assertTrue(
            guarded,
            "the no---show path does not remove WAYLAND_DISPLAY and force "
            "--ozone-platform=x11, so a window can reach the desktop")

    def test_the_orphan_sweep_steps_aside_for_a_second_copy(self):
        import inspect
        source = inspect.getsource(fs.sweep_orphan_browsers)
        self.assertIn("other_copies_running", source,
                      "the sweep could kill a running run's browser")



# ===========================================================================
# 29. the terminal is quiet unless you ask to watch
# ===========================================================================

class QuietTerminalTests(SandboxTest):
    """The log file always gets every step; the terminal only with --show."""

    def tearDown(self):
        fs.setup_logging()
        SandboxTest.tearDown(self)

    def _levels(self, **flags):
        args = type("Args", (), {"verbose": False, "quiet": False,
                                 "show": False})()
        for name, value in flags.items():
            setattr(args, name, value)
        fs.setup_logging(
            level="DEBUG" if args.verbose else None,
            screen=("DEBUG" if args.verbose
                    else "ERROR" if args.quiet
                    else "INFO" if args.show
                    else "WARNING"))
        screen = file_to = None
        for handler in fs.LOG.handlers:
            if isinstance(handler, logging.FileHandler):
                file_to = handler
            elif isinstance(handler, logging.StreamHandler):
                screen = handler
        return screen, file_to

    def test_the_terminal_hides_the_steps_by_default(self):
        screen, _file = self._levels()
        self.assertGreater(screen.level, logging.INFO,
                           "the step-by-step log goes to the terminal")

    def test_the_log_file_keeps_the_steps_anyway(self):
        _screen, file_to = self._levels()
        self.assertLessEqual(file_to.level, logging.INFO,
                             "the log file lost the detail as well")

    def test_show_puts_the_steps_on_the_terminal(self):
        screen, _file = self._levels(show=True)
        self.assertLessEqual(screen.level, logging.INFO)

    def test_verbose_puts_the_steps_on_the_terminal(self):
        screen, _file = self._levels(verbose=True)
        self.assertLessEqual(screen.level, logging.INFO)

    def test_quiet_hides_even_the_warnings(self):
        screen, _file = self._levels(quiet=True)
        self.assertGreater(screen.level, logging.WARNING)

    def test_a_warning_still_reaches_a_quiet_terminal(self):
        screen, _file = self._levels()
        self.assertLessEqual(screen.level, logging.WARNING,
                             "a problem would be silent")

    def test_the_result_is_printed_even_when_quiet(self):
        """print_run_summary must not vanish with the step log."""
        import ast
        with open(SCRIPT, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        for node in ast.walk(tree):
            if not (isinstance(node, ast.FunctionDef)
                    and node.name == "print_run_summary"):
                continue
            calls = {getattr(c.func, "id", "")
                     for c in ast.walk(node) if isinstance(c, ast.Call)}
            self.assertIn("announce", calls,
                          "the summary would be hidden with the step log")
            return
        self.fail("print_run_summary not found")

    def test_announce_does_not_say_things_twice(self):
        self._levels(show=True)
        self.assertFalse(fs.terminal_is_quiet())
        self._levels()
        self.assertTrue(fs.terminal_is_quiet())

    def test_the_how_to_watch_lines_are_not_shown_under_show(self):
        """They would contradict themselves: 'nothing more will be
        printed' while printing everything."""
        import ast
        with open(SCRIPT, encoding="utf-8") as handle:
            source = handle.read()
        self.assertIn("if terminal_is_quiet():", source)
        tree = ast.parse(source)
        for node in ast.walk(tree):
            if not isinstance(node, ast.If):
                continue
            if "terminal_is_quiet" not in ast.dump(node.test):
                continue
            body = " ".join(ast.dump(n) for n in node.body)
            if "to watch it" in body:
                return
        self.fail("the 'to watch it' line is not guarded")



# ===========================================================================
# 30. the live status line
# ===========================================================================

class StatusLineTests(unittest.TestCase):
    """One line that says what is happening now, rewritten in place."""

    class Fake:
        """Stands in for a terminal."""

        def __init__(self, tty=True):
            self.written = []
            self.tty = tty

        def write(self, text):
            self.written.append(text)

        def flush(self):
            pass

        def isatty(self):
            return self.tty

        def text(self):
            return "".join(self.written)

    def setUp(self):
        self.kept_out = sys.stdout
        self.kept_handler = fs.SCREEN_HANDLER
        self.status = fs.StatusLine()

    def tearDown(self):
        sys.stdout = self.kept_out
        fs.SCREEN_HANDLER = self.kept_handler

    def _quiet_terminal(self, tty=True):
        fake = self.Fake(tty)
        sys.stdout = fake
        handler = logging.StreamHandler(fake)
        handler.setLevel(logging.WARNING)       # the step log is NOT shown
        fs.SCREEN_HANDLER = handler
        return fake

    def _loud_terminal(self):
        fake = self.Fake(True)
        sys.stdout = fake
        handler = logging.StreamHandler(fake)
        handler.setLevel(logging.INFO)          # --show: steps ARE shown
        fs.SCREEN_HANDLER = handler
        return fake

    def test_it_writes_one_line_with_no_newline(self):
        fake = self._quiet_terminal()
        self.status.set("copying - 10 schools")
        self.assertIn("copying - 10 schools", fake.text())
        self.assertNotIn("\n", fake.text(),
                         "a newline would scroll instead of rewriting")

    def test_every_update_returns_to_the_start_of_the_line(self):
        fake = self._quiet_terminal()
        self.status.set("first")
        self.status.set("second")
        self.assertTrue(fake.text().startswith("\r"))
        text = fake.text()
        self.assertNotIn(chr(10), text,
                         "the line scrolled instead of rewriting")
        # one write per update; the text appears twice inside it so
        # the cursor ends on the line, not out past the padding
        self.assertEqual(text.count(chr(13) + "first"), 2)
        self.assertIn("second", text)
        self.assertLess(text.index("first"), text.index("second"))

    def test_a_shorter_line_covers_the_longer_one_before_it(self):
        fake = self._quiet_terminal()
        self.status.set("a very long status line indeed")
        fake.written = []
        self.status.set("short")
        self.assertIn(" ", fake.text(),
                      "the tail of the previous line would still be visible")

    def test_nothing_is_written_when_it_is_not_a_terminal(self):
        fake = self._quiet_terminal(tty=False)
        self.status.set("copying")
        self.assertEqual(fake.text(), "",
                         "carriage returns would litter a redirected file")

    def test_nothing_is_written_under_show(self):
        fake = self._loud_terminal()
        self.status.set("copying")
        self.assertEqual(fake.text(), "",
                         "--show prints every step, so a status line would "
                         "fight with it")

    def test_clearing_wipes_the_line(self):
        fake = self._quiet_terminal()
        self.status.set("something")
        fake.written = []
        self.status.clear()
        self.assertTrue(fake.text().startswith("\r"))
        self.assertIn(" ", fake.text())

    def test_a_log_record_does_not_scribble_over_it(self):
        """A warning must not land in the middle of the status line."""
        import inspect
        source = inspect.getsource(fs.TidyScreenHandler)
        self.assertIn("STATUS.clear()", source)
        self.assertIn("STATUS.draw()", source)

    def test_a_finished_city_leaves_a_permanent_line(self):
        """The live line says "now"; this one is the history."""
        import ast
        with open(SCRIPT, encoding="utf-8") as handle:
            tree = ast.parse(handle.read())
        for node in ast.walk(tree):
            if not (isinstance(node, ast.FunctionDef)
                    and node.name == "process_city"):
                continue
            calls = [c for c in ast.walk(node)
                     if isinstance(c, ast.Call)
                     and getattr(c.func, "id", "") == "announce"]
            self.assertTrue(
                calls, "a finished city leaves nothing on a quiet terminal")
            joined = " ".join(ast.dump(c) for c in calls)
            self.assertIn("done:", joined)
            return
        self.fail("process_city not found")

    def test_the_status_says_where_the_run_is(self):
        fake = self._quiet_terminal()
        kept = dict(fs.WHERE)
        try:
            fs.WHERE.update({"city": "Delhi", "city_at": 3, "cities": 35,
                             "keyword": "CBSE Schools", "keyword_at": 2,
                             "keywords": 7})
            fs.STATUS = self.status
            fs.say_status("copying from Maps")
            line = fake.text()
            for part in ("city 3/35", "Delhi", "keyword 2/7",
                         "CBSE Schools", "copying from Maps"):
                self.assertIn(part, line)
        finally:
            fs.WHERE.clear()
            fs.WHERE.update(kept)
            fs.STATUS = fs.StatusLine()



# ===========================================================================
# 31. the same email, website or phone number means the same school
# ===========================================================================

class ContactDuplicateTests(SandboxTest):
    """Checked against master_data.xlsx, and against the city being read."""

    def _save(self, record):
        return fs.append_records([record])

    def test_the_same_email_is_a_duplicate(self):
        first = school("Sunrise Public School",
                       **{"PIN Code": "560001", "School Address": "1 Road",
                          "School Official Email": "office@sunrise.edu.in"})
        self.assertEqual(self._save(first), (1, 0))
        again = school("Sunrise Pub. School",
                       **{"PIN Code": "560001", "School Address": "1 Road",
                          "School Official Email": "office@sunrise.edu.in"})
        self.assertEqual(self._save(again), (0, 1),
                         "the same email was saved twice")
        self.assertEqual(fs.row_count(fs.master_path()), 1)

    def test_the_same_website_is_a_duplicate(self):
        first = school("Hilltop School",
                       **{"PIN Code": "560002", "School Address": "2 Road",
                          "School Website": "https://hilltop.example.edu.in"})
        self._save(first)
        again = school("Hilltop Sr Sec School",
                       **{"PIN Code": "560002", "School Address": "2 Road",
                          "School Website": "http://www.hilltop.example.edu.in/about"})
        self.assertEqual(self._save(again), (0, 1),
                         "www and a path made the same site look different")

    def test_the_same_phone_number_is_a_duplicate(self):
        first = school("River School",
                       **{"PIN Code": "560003", "School Address": "3 Road",
                          "School Official Phone Number": "+91 98111 22333"})
        self._save(first)
        again = school("River School Annexe",
                       **{"PIN Code": "560003", "School Address": "3 Road",
                          "School Official Phone Number": "09811122333"})
        self.assertEqual(self._save(again), (0, 1),
                         "the same number written differently was saved twice")

    def test_one_number_written_four_ways_gives_one_key(self):
        """The column keeps the number as found; the key must not."""
        forms = ("+91 98111 22333", "09811122333", "9811122333",
                 "91-98111-22333")
        keys = {fs.phone_key(form) for form in forms}
        self.assertEqual(keys, {"9811122333"},
                         "the same telephone produced %d keys" % len(keys))

    def test_a_short_number_is_not_a_key(self):
        for junk in ("12345", "560011", ""):
            self.assertEqual(fs.phone_key(junk), "",
                             "%r became a phone key" % junk)

    def test_an_alternate_number_matches_an_official_one(self):
        first = school("Green Valley",
                       **{"PIN Code": "560004", "School Address": "4 Road",
                          "School Official Phone Number": "9811100001"})
        self._save(first)
        again = school("Green Valley School",
                       **{"PIN Code": "560004", "School Address": "4 Road",
                          "School Official Phone Number": "",
                          "School Alternate Phone Number": "9811100001"})
        self.assertEqual(self._save(again), (0, 1))

    def test_it_is_checked_against_the_master_file_itself(self):
        """Not just within one batch - the rule reads master_data.xlsx."""
        first = school("Master Check School",
                       **{"PIN Code": "560005", "School Address": "5 Road",
                          "School Official Email": "head@mcs.example.edu.in"})
        self._save(first)
        self.assertEqual(fs.row_count(fs.master_path()), 1)
        keymap = fs.master_key_map()
        self.assertIn(("email", "head@mcs.example.edu.in"), keymap,
                      "the master's emails are not part of the check")
        later = school("A Different Name Entirely",
                       **{"PIN Code": "560005", "School Address": "5 Road",
                          "School Official Email": "head@mcs.example.edu.in"})
        self.assertIsNotNone(fs.found_in_master(later, keymap))

    # ------------------------------------------------- what must NOT merge --
    def test_two_branches_sharing_a_website_are_both_kept(self):
        """A chain uses one domain for every branch.

        Without the contradiction guard this rule would quietly throw away
        every branch but the first - losing real schools, which is worse
        than a duplicate row.
        """
        rohini = school("DPS Rohini",
                        **{"PIN Code": "110085", "School Address": "Sector 24 Rohini",
                           "School Website": "https://dps.example.edu.in",
                           "School Official Phone Number": "9811100011"})
        dwarka = school("DPS Dwarka",
                        **{"PIN Code": "110078", "School Address": "Sector 3 Dwarka",
                           "School Website": "https://dps.example.edu.in",
                           "School Official Phone Number": "9811100022"})
        self._save(rohini)
        self.assertEqual(self._save(dwarka), (1, 0),
                         "a real branch school was thrown away as a duplicate")
        self.assertEqual(fs.row_count(fs.master_path()), 2)

    def test_two_branches_sharing_a_phone_are_both_kept(self):
        """One trust office number answers for several schools."""
        one = school("Trust School North",
                     **{"PIN Code": "110085", "School Address": "North Block Rohini",
                        "School Official Phone Number": "9811100033"})
        two = school("Trust School South",
                     **{"PIN Code": "110078", "School Address": "South Block Dwarka",
                        "School Official Phone Number": "9811100033"})
        self._save(one)
        self.assertEqual(self._save(two), (1, 0))

    def test_a_blank_contact_does_not_match_everything(self):
        one = school("Blank One", **{"PIN Code": "560011",
                                     "School Address": "11 Road",
                                     "School Official Email": "",
                                     "School Website": "",
                                     "School Official Phone Number": "9811100044"})
        two = school("Blank Two", **{"PIN Code": "560012",
                                     "School Address": "12 Lane",
                                     "School Official Email": "",
                                     "School Website": "",
                                     "School Official Phone Number": "9811100055"})
        self._save(one)
        self.assertEqual(self._save(two), (1, 0),
                         "two blanks were treated as the same contact")

    def test_a_placeholder_is_not_a_contact_key(self):
        for junk in ("N/A", "Not Available", "-", "NA"):
            record = school("Junk School",
                            **{"School Official Email": junk,
                               "School Website": junk,
                               "School Official Phone Number": junk})
            kinds = [kind for kind, _v in fs.contact_keys(record)]
            self.assertEqual(kinds, [],
                             "%r became a duplicate key" % junk)

    # ------------------------------------------------- the cycle file --
    def test_the_cycle_file_is_checked_as_well_as_the_master(self):
        """A row can be in the cycle but not in the master.

        Found by testing: an interrupted save, or a file edited by hand,
        leaves a row the master has not got. Checking only the master let
        its twin through and wrote the same school into the cycle file
        twice.
        """
        fs.append_records([school("Cycle Known",
                                  **{"PIN Code": "110001",
                                     "School Address": "1 Road",
                                     "School Official Email": "k@c.example.edu.in"})])
        stray = school("Cycle Stray",
                       **{"PIN Code": "110002", "School Address": "2 Road",
                          "School Official Phone Number": "9811100099"})
        fs.write_records(fs.cycle_path(0),
                         fs.read_records(fs.cycle_path(0)) + [stray])
        twin = school("Cycle Twin",
                      **{"PIN Code": "110002", "School Address": "2 Road",
                         "School Official Phone Number": "+91 98111 00099"})
        written, skipped = fs.append_records([twin])
        self.assertEqual((written, skipped), (0, 1),
                         "the twin of a cycle-only row was saved")
        names = [r["School Name"]
                 for r in fs.read_records(fs.cycle_path(0))]
        self.assertNotIn("Cycle Twin", names)
        self.assertEqual(names.count("Cycle Stray"), 1)

    def test_a_refused_record_reaches_neither_file(self):
        """The check happens before either write, so they never disagree."""
        stray = school("Neither Stray",
                       **{"PIN Code": "110003", "School Address": "3 Road",
                          "School Website": "https://neither.example.edu.in"})
        fs.write_records(fs.cycle_path(0),
                         fs.read_records(fs.cycle_path(0)) + [stray])
        twin = school("Neither Twin",
                      **{"PIN Code": "110003", "School Address": "3 Road",
                         "School Website": "http://www.neither.example.edu.in/x"})
        fs.append_records([twin])
        in_master = [r["School Name"]
                     for r in fs.read_records(fs.master_path())]
        in_cycle = [r["School Name"]
                    for r in fs.read_records(fs.cycle_path(0))]
        self.assertNotIn("Neither Twin", in_master)
        self.assertNotIn("Neither Twin", in_cycle,
                         "it was kept out of one file but not the other")

    def test_a_branch_is_still_added_to_the_cycle(self):
        """The cycle check uses the same guard: branches are not lost."""
        fs.append_records([school("Chain Rohini",
                                  **{"PIN Code": "110085",
                                     "School Address": "Sector 24 Rohini",
                                     "School Website": "https://chain.example.edu.in"})])
        written, _skipped = fs.append_records(
            [school("Chain Dwarka",
                    **{"PIN Code": "110078",
                       "School Address": "Sector 3 Dwarka",
                       "School Website": "https://chain.example.edu.in"})])
        self.assertEqual(written, 1)
        names = [r["School Name"]
                 for r in fs.read_records(fs.cycle_path(0))]
        self.assertIn("Chain Dwarka", names,
                      "a real branch never reached the cycle file")

    # --------------------------------------------------------- agreement --
    def test_saving_and_the_missing_check_agree(self):
        """If they disagree, a city can never close.

        A school rejected as a duplicate when saving would be reported
        missing from the master for ever, and its city would stay open.
        """
        first = school("Agreement School",
                       **{"PIN Code": "560006", "School Address": "6 Road",
                          "School Official Email": "a@agree.example.edu.in"})
        self._save(first)
        twin = school("Agreement School Two",
                      **{"PIN Code": "560006", "School Address": "6 Road",
                         "School Official Email": "a@agree.example.edu.in"})
        written, skipped = self._save(twin)
        self.assertEqual((written, skipped), (0, 1))
        self.assertEqual(fs.missing_from_master([twin]), [],
                         "a school refused as a duplicate is reported "
                         "missing - its city could never close")

    def test_the_repair_agrees_with_saving_too(self):
        one = school("Repair School",
                     **{"PIN Code": "110085", "School Address": "North Rohini",
                        "School Official Phone Number": "9811100066"})
        two = school("Repair School Other",
                     **{"PIN Code": "110078", "School Address": "South Dwarka",
                        "School Official Phone Number": "9811100066"})
        self._save(one)
        self._save(two)
        before = fs.row_count(fs.master_path())
        report = fs.reconcile_master_and_cycles()
        self.assertEqual(report["added_to_master"], 0)
        self.assertEqual(report["added_to_current"], 0)
        self.assertEqual(fs.row_count(fs.master_path()), before,
                         "the repair changed rows that were already right")


if __name__ == "__main__":
    unittest.main(verbosity=2)
