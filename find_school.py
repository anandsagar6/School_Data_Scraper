#!/usr/bin/env python3
"""School Data Extractor - one standalone file.

Collects Indian school data city by city from Google Maps, researches each
school with the Gemini web application, and writes the approved 57-column
Excel dataset.

    Rank -> Assigned Cities -> School Search Keywords -> School Discovery
    -> Duplicate Removal -> Temporary Storage -> Gemini Enrichment
    -> Parsing -> Validation -> Excel Save -> Checkpoint -> Next City

EVERYTHING IS IN THIS FILE
--------------------------
The city dataset, the search keywords, the 57 fields, both Gemini prompts, the
parser, the duplicate rules, the Excel writer, the checkpoint and the browser
driver. Copy this one file to another laptop, install the two dependencies
below, change ``rank`` at the top, and it runs. It imports nothing from this
project.

INSTALL
-------
    python3 -m pip install playwright openpyxl
    python3 -m playwright install chromium
    sudo apt install xvfb        # only to run it without a visible screen

RUN
---
    python3 find_school.py --status        # what this laptop will do
    python3 find_school.py --dry-run       # the plan, no browser, no network
    python3 find_school.py --self-test     # offline checks, no network
    python3 find_school.py --state Karnataka --city Bengaluru --limit 10 --test
    python3 find_school.py --state Karnataka --city Bengaluru
    python3 find_school.py                 # this rank's cities, in order
    python3 find_school.py --retry-failed  # the schools that failed before

WHAT IT WILL NOT DO
-------------------
No CAPTCHA solving, no proxy rotation, no fingerprint spoofing, no attempt to
get past a sign-in page or a quota. If a source will not serve the data, the
run says so and stops that step instead of inventing records.
"""

import argparse
import datetime
import glob
import gzip
import hashlib
import json
import logging
import os
import random
import re
import shutil
import signal
import subprocess
import sys
import tempfile
import time
import unicodedata
import urllib.parse

# ===========================================================================
#  SETTINGS
# ===========================================================================
#  Everything meant to be changed, most important first. Anything not here
#  is technical and lives further down, beside the code that uses it.
# ---------------------------------------------------------------------------

# --- this laptop -----------------------------------------------------------
rank = 1                            # laptop 1..num_sys; rank.json overrides
num_sys = 5                         # SAME on every laptop; splits the cities


# --- where the records go --------------------------------------------------
CYCLE_SIZE = 1000                   # rows in a finished Excel cycle file
BACKUP_COMPLETED_CYCLES = True      # keep a verified copy in Backup/

# --- GitHub sync -----------------------------------------------------------
GIT_SYNC_ENABLED = True             # push recovery files to a remote
GIT_REMOTE = "origin"               # rank.json "git_remote" overrides
GIT_SYNC_MINUTES = 5                # never sync more often than this
GIT_BRANCH_PER_RANK = True          # rank 1 -> school-data-rank-1

# --- what is collected -----------------------------------------------------
OFFICIAL_CONTACT_SEARCH = True      # official email/phone from Google Search
KEYWORD_FIRST_ORDER = True          # one keyword across every city, then next
COUNTRY = "India"                   # added to every Maps and Google query

# --- working hours ---------------------------------------------------------
WORK_PERIOD_MIN_HOURS = 2           # work this long (working time only)...
WORK_PERIOD_MAX_HOURS = 6
BREAK_MIN_MINUTES = 30              # ...then stop this long
BREAK_MAX_MINUTES = 90
NIGHT_PAUSE_ENABLED = False         # no collecting between the hours below
NIGHT_START_HOUR = 23
NIGHT_END_HOUR = 5
TIME_ZONE = "Asia/Kolkata"          # the clock the night stop is read on

# --- pacing: every wait is a fresh random draw from its range --------------
GOOGLE_BETWEEN_SCHOOLS_DELAY_SECONDS = (8, 16)   # between two schools
GOOGLE_BETWEEN_QUERIES_DELAY_SECONDS = (3, 7)    # email search -> phone search
GOOGLE_TO_GEMINI_DELAY_SECONDS = (5, 12)         # searches -> Gemini
PAUSE_CLOSES_BROWSER_SECONDS = 600  # a break this long closes the browser

# --- Google Maps scrolling -------------------------------------------------
#  Each copy takes the WHOLE list, so more presses between copies cannot
#  skip a school - it only means fewer copies. One press moves ~24 pixels.
DOWN_PRESSES_PER_ROUND = 10         # lower = copy more often, slower
SCROLL_PAUSE_SECONDS = 2.5          # wait after the presses for new results

# --- Gemini ----------------------------------------------------------------
GEMINI_URL = "https://gemini.google.com/app?hl=en-IN"
GEMINI_FIRST_READ_SECONDS = 15      # before the answer is looked at
GEMINI_ANSWER_TIMEOUT = 240         # give up on one school after this
GEMINI_RETRIES = 2                  # extra tries for an unusable answer

# --- housekeeping ----------------------------------------------------------
KEEP_CAPTURED_TEXT = False          # keep a finished keyword's copied text
LOG_KEEP_DAYS = 10                  # trim find_school.log to this; 0 keeps all


# ===========================================================================
#  WHERE THINGS ARE WRITTEN
# ===========================================================================

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = PROJECT_DIR
TEMP_DIR = os.path.join(OUTPUT_DIR, "Temp_Data")
EXCEL_DIR = os.path.join(OUTPUT_DIR, "Excel_Data")
#: Permanent copies of finished cycle files. Written once, verified, and
#: never removed by any cleanup in this script.
BACKUP_DIR = os.path.join(OUTPUT_DIR, "Backup")
LOG_FILE = os.path.join(OUTPUT_DIR, "find_school.log")
CHECKPOINT_FILE = os.path.join(OUTPUT_DIR, "checkpoint.json")
RANK_FILE = os.path.join(PROJECT_DIR, "rank.json")
BROWSER_PROFILE_DIR = os.path.join(OUTPUT_DIR, "browser_profile")
GEMINI_PROFILE_DIR = os.path.join(OUTPUT_DIR, "browser_profile_gemini")

#: The cities, the school keywords and the 57 fields all come from the
#: mylib.py BESIDE THIS SCRIPT - the only place that is looked in, so this
#: folder is self-contained and no other project's copy can be picked up
#: by accident. Nothing is copied into this file, and mylib.py is read by
#: parsing it, never by importing it, so the mail passwords in it are
#: never executed or loaded.
MYLIB_PATHS = [
    os.path.join(PROJECT_DIR, "mylib.py"),
]
MYLIB_CITY_NAME = "city_by_states"
MYLIB_KEYWORD_NAME = "SCHOOL_SEARCH_KEYWORDS"
MYLIB_FIELD_NAME = "SCHOOL_FIELDS"




def set_rank_paths(this_rank=None):
    """Settle where this run writes. One flat set of folders, always.

    rank and num_sys decide which cities this laptop takes, never where it
    writes: each laptop has its own copy of this script and its own folder.
    """
    global TEMP_DIR, EXCEL_DIR, CHECKPOINT_FILE, BACKUP_DIR
    TEMP_DIR = os.path.join(OUTPUT_DIR, "Temp_Data")
    EXCEL_DIR = os.path.join(OUTPUT_DIR, "Excel_Data")
    BACKUP_DIR = os.path.join(OUTPUT_DIR, "Backup")
    CHECKPOINT_FILE = os.path.join(OUTPUT_DIR, "checkpoint.json")
    ensure_dirs()
    return TEMP_DIR, EXCEL_DIR


# ===========================================================================
#  THE 57 SCHOOL FIELDS
# ===========================================================================
#  (Excel column heading, Gemini answer key), read from mylib.py. The Excel
#  columns, the requested field list in the prompt, the answer format and
#  the parser are all built from that one tuple, so they cannot drift apart
#  - and there is no second copy of it here.
#
#  One deliberate decision: the prompt's rules mention Source_URLs while the
#  answer format uses School_Data_Source_URL. School_Data_Source_URL is the
#  authoritative field; Source_URLs is accepted by the parser as a spelling
#  of it and never becomes a 58th column.
# ---------------------------------------------------------------------------



#: Spellings Gemini sometimes uses, mapped onto the authoritative key.
KEY_SYNONYMS = {
    "Source_URLs": "School_Data_Source_URL",
    "Source_URL": "School_Data_Source_URL",
    "Available_Streams_Science_Commerce_Arts": "Available_Streams",
    "Co_educational_Status": "Co_Educational_Status",
    "PIN_code": "PIN_Code",
    "Udise_Code": "UDISE_Code",
}

#: Written as text, never as a number, so a leading zero survives.
TEXT_COLUMNS = frozenset({
    "School Official Phone Number", "School Alternate Phone Number",
    "Principal Office Phone Number", "PIN Code", "UDISE Code",
    "Affiliation Number", "School Registration Number",
    "Year of Establishment", "School Latitude", "School Longitude",
    "Total Student Strength", "Number of Boys", "Number of Girls",
    "Number of Teaching Staff", "Number of Non-Teaching Staff",
    "Number of Computers",
})

#: Filled from Google Maps, so a Gemini answer never silently replaces them.
SOURCE_EVIDENCE_COLUMNS = (
    "School Name", "School Address", "Area / Locality", "City", "State",
    "Country", "School Website", "School Official Phone Number", "PIN Code",
    "School Google Maps Link", "School Latitude", "School Longitude",
)

#: What a blank looks like coming back from Gemini.
NO_DATA = "NO_DATA"
BLANK_ANSWERS = frozenset({
    "", "no_data", "nodata", "n/a", "na", "none", "null", "-", "--",
    "not available", "not found", "unknown", "value",
})


# ---------------------------------------------------------------------------
#  READING YOUR mylib.py
# ---------------------------------------------------------------------------
#  The cities and the school keywords live in your own mylib.py so that one
#  edit there reaches every script that uses them.
#
#  That file is NOT imported. It is parsed, and only the two lists wanted
#  are read out of it. Three reasons:
#
#    * importing it would run it, which pulls in Levenshtein, gtts and
#      psutil - packages this script has no use for and a fresh laptop may
#      not have;
#    * that file also holds mail accounts and their passwords, and reading
#      two named lists out of the syntax tree cannot touch them;
#    * a syntax error anywhere else in it cannot stop this script.
# ---------------------------------------------------------------------------

#: Filled in by load_shared_lists(); each says where its value came from.
SHARED_SOURCE = {"cities": "not loaded", "keywords": "not loaded",
                 "fields": "not loaded"}


def mylib_path():
    """The mylib.py this run reads, the nearest one first."""
    told = os.environ.get("SCHOOL_MYLIB")
    if told is not None:
        return told
    for path in MYLIB_PATHS:
        if path and os.path.exists(path):
            return path
    return MYLIB_PATHS[0] if MYLIB_PATHS else ""


def read_list_from_file(path, name):
    """One top-level list or dict literal out of a Python file.

    Nothing in the file is executed and nothing else in it is read. Returns
    None when the file, the name or a plain literal value is not there.
    """
    import ast
    if not path or not os.path.exists(path):
        return None
    try:
        with open(path, encoding="utf-8") as handle:
            tree = ast.parse(handle.read(), filename=path)
    except (OSError, SyntaxError, ValueError) as error:
        # this runs while the module is still being read, before logging
        # exists, so it goes to the screen directly
        print("WARNING %s could not be read (%s)"
              % (os.path.basename(path), str(error)[:100]))
        return None
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        for target in node.targets:
            if getattr(target, "id", "") != name:
                continue
            try:
                return ast.literal_eval(node.value)
            except (ValueError, SyntaxError):
                print("WARNING %s in %s is not a plain list"
                      % (name, os.path.basename(path)))
                return None
    return None


def load_shared_lists():
    """Take the cities and the keywords out of mylib.py.

    That file is the one place they live, so there is no second copy here
    to drift out of step with it. If it cannot be read the lists come back
    empty and SHARED_SOURCE says why; main() stops with that message rather
    than running against nothing.
    """
    path = mylib_path()
    short = os.path.basename(path) if path else "(none)"

    if not path:
        why = "no mylib.py is configured - set MYLIB_PATHS or SCHOOL_MYLIB"
        SHARED_SOURCE.update(cities=why, keywords=why, fields=why)
        return {}, [], ()
    if not os.path.exists(path):
        why = "%s was not found" % path
        SHARED_SOURCE.update(cities=why, keywords=why, fields=why)
        return {}, [], ()

    cities = read_list_from_file(path, MYLIB_CITY_NAME)
    if isinstance(cities, dict) and cities:
        SHARED_SOURCE["cities"] = "%s (%s)" % (short, MYLIB_CITY_NAME)
    else:
        cities = {}
        SHARED_SOURCE["cities"] = "%s has no usable %s" % (short,
                                                           MYLIB_CITY_NAME)

    keywords = read_list_from_file(path, MYLIB_KEYWORD_NAME)
    if isinstance(keywords, (list, tuple)) and keywords:
        keywords = [" ".join(str(k).split()) for k in keywords
                    if str(k).strip()]
        SHARED_SOURCE["keywords"] = "%s (%s)" % (short, MYLIB_KEYWORD_NAME)
    else:
        keywords = []
        SHARED_SOURCE["keywords"] = "%s has no usable %s" % (
            short, MYLIB_KEYWORD_NAME)

    fields = read_list_from_file(path, MYLIB_FIELD_NAME)
    if (isinstance(fields, (list, tuple)) and fields
            and all(isinstance(pair, (list, tuple)) and len(pair) == 2
                    for pair in fields)):
        fields = tuple((str(label), str(key)) for label, key in fields)
        SHARED_SOURCE["fields"] = "%s (%s)" % (short, MYLIB_FIELD_NAME)
    else:
        fields = ()
        SHARED_SOURCE["fields"] = "%s has no usable %s" % (short,
                                                           MYLIB_FIELD_NAME)
    return cities, keywords, fields


def shared_lists_problem():
    """A sentence explaining why the lists are missing, or ""."""
    if CITY_BY_STATES and SCHOOL_KEYWORDS and SCHOOL_FIELDS:
        return ""
    missing = []
    if not CITY_BY_STATES:
        missing.append("the cities (%s)" % SHARED_SOURCE["cities"])
    if not SCHOOL_KEYWORDS:
        missing.append("the keywords (%s)" % SHARED_SOURCE["keywords"])
    if not SCHOOL_FIELDS:
        missing.append("the 57 fields (%s)" % SHARED_SOURCE["fields"])
    return ("This script reads its city list and its school keywords from "
            "your mylib.py, and could not get %s.\n"
            "Looked in: %s\n"
            "Fix it by pointing MYLIB_PATHS at the right file, or set "
            "SCHOOL_MYLIB in the environment:\n"
            "    SCHOOL_MYLIB=/path/to/mylib.py python3 find_school.py "
            "--status" % (" and ".join(missing), mylib_path() or "nowhere"))


CITY_BY_STATES, SCHOOL_KEYWORDS, SCHOOL_FIELDS = load_shared_lists()

#: Built from the one tuple above, so the Excel columns, the prompt, the
#: answer format and the parser can never disagree about a field.
SCHOOL_COLUMNS = tuple(label for label, _key in SCHOOL_FIELDS)
ANSWER_KEYS = tuple(key for _label, key in SCHOOL_FIELDS)
KEY_TO_COLUMN = {key: label for label, key in SCHOOL_FIELDS}
COLUMN_TO_KEY = {label: key for label, key in SCHOOL_FIELDS}

#: Which of the keywords this laptop actually searches for. The full list
#: is shared in mylib.py; which ones to run is a choice per machine, so it
#: lives here. Empty means every keyword in the shared list.
ENABLED_SCHOOL_KEYWORDS = [
    "Private Schools",
    "Government Schools",
]


def enabled_keywords():
    """The search keywords this run will use, in the shared list's order."""
    if not ENABLED_SCHOOL_KEYWORDS:
        return list(SCHOOL_KEYWORDS)
    wanted = {clean(name).lower() for name in ENABLED_SCHOOL_KEYWORDS}
    return [name for name in SCHOOL_KEYWORDS if clean(name).lower() in wanted]


def unknown_enabled_keywords():
    """Enabled names that the shared list does not have - a typo guard."""
    have = {clean(name).lower() for name in SCHOOL_KEYWORDS}
    return [name for name in ENABLED_SCHOOL_KEYWORDS
            if clean(name).lower() not in have]


# ===========================================================================
#  THE GEMINI PROMPT
# ===========================================================================
#  The text of school_gemini_prompt.txt, with the field list and the answer
#  format generated from SCHOOL_FIELDS above rather than typed out again.
# ---------------------------------------------------------------------------

GEMINI_PROMPT_TEMPLATE = """\
School Name is {school_name}. It is located in {city}, {state}, India.

Already known about this school:
School Address: {school_address}
Area / Locality: {school_locality}
PIN Code: {school_pin}
School Website: {school_website}
School Phone: {existing_school_phone}
Google Maps Link: {school_maps_link}

Give the following data for {school_name}:

{requested_fields}

--------------------------------

Rules:

Search for this exact school and this exact branch or campus. Use the address,
locality, PIN code and Google Maps link above to identify it.
Do not mix in details of a school with a similar name, and do not use details of
another branch just because it shares the same website.
Do not invent, guess or estimate anything. Never build an email address or a phone
number from the school name or the website domain.
Never invent a UDISE Code, Affiliation Number or School Registration Number.
Give the source URLs that support your answer in School_Data_Source_URL, separated
by semi-colons.
Where a field can hold more than one value, separate the values with semi-colons.
Write one line per field, in the order given, using the field name exactly as shown.
Do not write any explanation, heading or markdown.
Do not say whether the information is verified; the program decides that itself.

--------------------------------

Write in this format:

{answer_format}


--------------------------------

If you do not find the "Value" of any data then you can simply write NO_DATA in \
that Value.
"""


def gemini_skips_columns():
    """Columns Gemini is NOT asked about: Google Search researches them.

    Asking as well would spend an answer on work already done, and invite
    a second, unverified opinion on the one pair of fields that has to be
    verifiable. The Excel file still has all 57 columns - only the prompt
    is shorter. Resolved when called, because the column names are
    defined further down the file beside the search that owns them.
    """
    if not OFFICIAL_CONTACT_SEARCH:
        return set()
    return {OFFICIAL_EMAIL_COLUMN, OFFICIAL_PHONE_COLUMN}


def gemini_asked_columns():
    """The columns Gemini is asked about, in the 57-column order."""
    skip = gemini_skips_columns()
    return [c for c in SCHOOL_COLUMNS if c not in skip]


def gemini_asked_keys():
    """The answer keys matching gemini_asked_columns()."""
    skip = gemini_skips_columns()
    return [key for column, key in zip(SCHOOL_COLUMNS, ANSWER_KEYS)
            if column not in skip]


def requested_fields_text():
    """The field names Gemini is asked for, one per line, in order."""
    return "\n".join(gemini_asked_columns())


def answer_format_text():
    """The ``Key: Value`` lines Gemini is asked to answer with."""
    return "\n".join("%s: Value    \\n" % key for key in gemini_asked_keys())


def build_prompt(record):
    """The finished prompt for one school record."""
    return GEMINI_PROMPT_TEMPLATE.format(
        school_name=record.get("School Name", "") or "",
        city=record.get("City", "") or "",
        state=record.get("State", "") or "",
        school_address=record.get("School Address", "") or NO_DATA,
        school_locality=record.get("Area / Locality", "") or NO_DATA,
        school_pin=record.get("PIN Code", "") or NO_DATA,
        school_website=record.get("School Website", "") or NO_DATA,
        existing_school_phone=(record.get("School Official Phone Number", "")
                               or NO_DATA),
        school_maps_link=record.get("School Google Maps Link", "") or NO_DATA,
        requested_fields=requested_fields_text(),
        answer_format=answer_format_text(),
    )

# ===========================================================================
#  WORK DIVISION - the GitHub Profile Extractor's algorithm, unchanged
# ===========================================================================

def get_cities_by_rank(all_cities, rank, num_sys):
    """This rank's equal, contiguous slice of the city list.

    Returns [] for a rank outside 1..num_sys, which the callers already
    report as "this rank has no cities to search" rather than crashing.
    """
    if num_sys < 1 or rank < 1 or rank > num_sys:
        return []
    len_city = len(all_cities)
    # base cities for every rank, one extra for the first `extra` ranks
    base, extra = divmod(len_city, num_sys)
    first_idx = base * (rank - 1) + min(rank - 1, extra)
    last_idx = first_idx + base + (1 if rank <= extra else 0)
    return all_cities[first_idx:last_idx]


def build_city_rows():
    """``[(state, city), ...]`` - the flattened, de-duplicated city list.

    The GitHub extractor's rule, kept exactly: flatten every group in file
    order, then drop exact duplicate names keeping the first occurrence,
    which is what ``dict.fromkeys`` does. A name listed under two groups
    therefore keeps the group it appears under first, so one city never
    belongs to two states.

    Every city is searched under its own name and nothing else. There is no
    canonical/alias distinction: "Bengaluru" and "Bangalore" are two separate
    cities here, exactly as they are in the GitHub extractor, and they may
    land on different ranks. That is the reference behaviour and it is chosen,
    not overlooked.
    """
    rows, seen = [], set()
    for state, names in CITY_BY_STATES.items():
        for name in names:
            if name in seen:
                continue
            seen.add(name)
            rows.append((state, name))
    return rows


CITY_ROWS = build_city_rows()
ALL_CITIES = [city for _state, city in CITY_ROWS]
CITY_TO_STATE = {city: state for state, city in CITY_ROWS}


def city_fingerprint():
    """Stable hash of the ordered city list ALONE.

    Kept apart from the keyword list so that changing which keywords are
    enabled - which costs nothing already collected - can be told apart
    from changing the cities, which would mean the saved progress refers
    to a different plan altogether.
    """
    payload = "\n".join("%s|%s" % pair for pair in CITY_ROWS)
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def dataset_fingerprint():
    """Stable hash of the ordered city list and the enabled keywords.

    Over the resolved plan, not the file bytes: editing a comment here does
    not invalidate a running collection, while adding, removing, renaming or
    reordering a city or a keyword does.
    """
    payload = "\n".join(["%s|%s" % pair for pair in CITY_ROWS]
                        + ["KEYWORD|%s" % k for k in enabled_keywords()])
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def resolve_git_remote():
    """``(remote, where_it_came_from)`` for the sync destination.

    Same ladder as the rank: the setting at the top of this file, then
    rank.json beside it, then SCHOOL_GIT_REMOTE in the environment. A
    production install therefore points at its own remote WITHOUT editing
    this file, so `git pull` of a source update never conflicts with a
    local edit and never has to touch the data branches.
    """
    value, source = GIT_REMOTE, "the top of this file"
    if os.path.exists(RANK_FILE):
        try:
            with open(RANK_FILE, encoding="utf-8") as handle:
                data = json.load(handle)
            if data.get("git_remote"):
                value = str(data["git_remote"])
                source = os.path.basename(RANK_FILE)
        except (OSError, ValueError, TypeError):
            pass
    from_env = os.environ.get("SCHOOL_GIT_REMOTE")
    if from_env:
        value, source = from_env, "the environment"
    return value, source


def resolve_rank(cli_rank=None, cli_num_sys=None):
    """``(rank, num_sys, where_it_came_from)``.

    Lowest to highest: the block at the top of this file, rank.json beside
    it, the SCHOOL_RANK / SCHOOL_NUM_SYS environment variables, the command
    line. Every layer above the first is optional - a fresh copy of this file
    runs on its built-in defaults alone.
    """
    value_rank, value_num, source = rank, num_sys, "the top of this file"

    if os.path.exists(RANK_FILE):
        try:
            with open(RANK_FILE, encoding="utf-8") as handle:
                data = json.load(handle)
            if "rank" in data or "num_sys" in data:
                value_rank = int(data.get("rank", value_rank))
                value_num = int(data.get("num_sys", value_num))
                source = os.path.basename(RANK_FILE)
        except (OSError, ValueError, TypeError):
            logging.getLogger("school").warning(
                "%s could not be read - using the values at the top of this "
                "file", RANK_FILE)

    env_rank, env_num = os.environ.get("SCHOOL_RANK"), os.environ.get("SCHOOL_NUM_SYS")
    if env_rank or env_num:
        try:
            value_rank = int(env_rank) if env_rank else value_rank
            value_num = int(env_num) if env_num else value_num
            source = "the environment"
        except ValueError:
            pass

    if cli_rank is not None:
        value_rank, source = int(cli_rank), "the command line"
    if cli_num_sys is not None:
        value_num, source = int(cli_num_sys), "the command line"
    return value_rank, value_num, source


class RankError(ValueError):
    """A rank that cannot be run, reported rather than silently ignored."""


def validate_rank(value_rank, value_num):
    """Raise RankError with a readable message for an impossible rank."""
    if value_num < 1:
        raise RankError("num_sys is %r: it must be 1 or more" % value_num)
    if value_rank < 1:
        raise RankError("rank is %r: ranks count from 1" % value_rank)
    if value_rank > value_num:
        raise RankError("rank %d is above num_sys %d: with %d laptops the "
                        "ranks are 1 to %d"
                        % (value_rank, value_num, value_num, value_num))


def cities_for_rank(value_rank, value_num):
    """``[(state, city), ...]`` for this rank, in plan order."""
    validate_rank(value_rank, value_num)
    mine = set(get_cities_by_rank(ALL_CITIES, value_rank, value_num))
    return [(state, city) for state, city in CITY_ROWS if city in mine]


def search_plan(value_rank, value_num, states=None, only_city=None):
    """``[(state, city, keyword), ...]`` - every search, in a fixed order.

        this rank's cities, in dataset order
            -> each enabled keyword, in list order
    """
    pairs = cities_for_rank(value_rank, value_num)
    if only_city:
        wanted = {c.strip().lower() for c in only_city}
        pairs = [p for p in pairs if p[1].lower() in wanted]
    if states:
        wanted = {s.strip().lower() for s in states}
        pairs = [p for p in pairs if p[0].lower() in wanted]
    keywords = enabled_keywords()
    return [(state, city, keyword)
            for state, city in pairs for keyword in keywords]


# ===========================================================================
#  LOGGING
# ===========================================================================

LOG_LEVEL = "INFO"
CITY_DELAY_SECONDS = (5.0, 10.0)    # between two cities

LOG = logging.getLogger("school")


def ensure_dirs():
    folders = [OUTPUT_DIR, TEMP_DIR, EXCEL_DIR]
    if BACKUP_COMPLETED_CYCLES:
        folders.append(BACKUP_DIR)
    for path in folders:
        os.makedirs(path, exist_ok=True)


#: The handler that writes to your terminal. Kept so that announce()
#: below can tell whether a line is already being shown there.
SCREEN_HANDLER = None


def setup_logging(level=None, screen=None):
    """Two destinations with two different levels.

    The file always gets every step. The terminal gets only what you asked
    for: by default the step-by-step log goes to find_school.log alone and
    the terminal stays quiet, so an unattended run does not fill the
    screen. --show (or --verbose) puts the steps on the terminal too.
    """
    global SCREEN_HANDLER
    ensure_dirs()
    LOG.handlers.clear()
    file_level = getattr(logging, (level or LOG_LEVEL).upper(), logging.INFO)
    screen_level = getattr(logging, (screen or "WARNING").upper(),
                           logging.WARNING)
    # the logger must pass whatever either destination wants
    LOG.setLevel(min(file_level, screen_level))
    to_screen = TidyScreenHandler(sys.stdout)
    to_screen.setFormatter(logging.Formatter("%(levelname)-7s %(message)s"))
    to_screen.setLevel(screen_level)
    LOG.addHandler(to_screen)
    SCREEN_HANDLER = to_screen
    trim_log()
    to_file = logging.FileHandler(LOG_FILE, encoding="utf-8")
    to_file.setFormatter(logging.Formatter(
        "%(asctime)s %(levelname)-7s %(message)s"))
    to_file.setLevel(file_level)
    LOG.addHandler(to_file)
    return LOG


class StatusLine:
    """One line at the bottom of the terminal saying what is happening NOW.

    It rewrites itself in place rather than scrolling, so a long run shows
    its current state without filling the screen. Only used when the
    step-by-step log is NOT on the terminal (no --show) and the output is
    a terminal rather than a file or a pipe - a file full of carriage
    returns helps nobody.
    """

    def __init__(self):
        self.text = ""
        self.width = 0

    def usable(self):
        if not terminal_is_quiet():
            return False
        try:
            return sys.stdout.isatty()
        except (AttributeError, ValueError):
            return False

    def set(self, message, *args):
        if not self.usable():
            return
        text = (message % args) if args else message
        self.text = text
        self.draw()

    def draw(self):
        if not self.usable() or not self.text:
            return
        try:
            room = max(shutil.get_terminal_size((100, 24)).columns - 1, 20)
        except Exception:                               # noqa: BLE001
            room = 99
        line = self.text[:room]
        # pad to cover whatever a longer previous line left behind, then
        # return to the start so the cursor is not parked out in space
        padding = " " * max(0, self.width - len(line))
        self.width = len(line)
        try:
            sys.stdout.write("\r" + line + padding + "\r" + line)
            sys.stdout.flush()
        except (OSError, ValueError):
            pass

    def clear(self):
        """Wipe the line so something permanent can be printed over it."""
        if not self.usable() or not self.width:
            return
        try:
            sys.stdout.write("\r" + " " * self.width + "\r")
            sys.stdout.flush()
        except (OSError, ValueError):
            pass
        self.width = 0

    def done(self):
        """Finish with the line left behind, on its own row."""
        if self.usable() and self.width:
            try:
                sys.stdout.write("\n")
                sys.stdout.flush()
            except (OSError, ValueError):
                pass
        self.text = ""
        self.width = 0


STATUS = StatusLine()


class TidyScreenHandler(logging.StreamHandler):
    """A log handler that does not scribble over the status line."""

    def emit(self, record):
        STATUS.clear()
        logging.StreamHandler.emit(self, record)
        STATUS.draw()


def terminal_is_quiet():
    """True when the step-by-step log is NOT going to the terminal."""
    return (SCREEN_HANDLER is None
            or SCREEN_HANDLER.level > logging.INFO)


def announce(message, *args):
    """Say something on the terminal even when the terminal is quiet.

    A few things must always reach you: that the run started, where to
    watch it, and how it ended. Everything else is in the log file.
    Printed only when the screen is not already showing INFO lines, so a
    --show run never says anything twice.
    """
    LOG.info(message, *args)
    if SCREEN_HANDLER is not None and SCREEN_HANDLER.level > logging.INFO:
        STATUS.clear()
        try:
            print(message % args if args else message)
            sys.stdout.flush()
        except (TypeError, ValueError):
            print(message)
        STATUS.draw()


# ===========================================================================
#  SMALL TEXT HELPERS
# ===========================================================================

_WS = re.compile(r"\s+")
_PHONE_OK = re.compile(r"[^0-9+]")
#: Six digits standing alone. The digit guards matter: without them the
#: middle of a phone number ("9900112233" -> "990011") reads as a PIN.
_PIN = re.compile(r"(?<!\d)([1-9][0-9]{5})(?!\d)")
#: A six digit number introduced by one of these is NOT a PIN code. A live
#: run read the affiliation number out of "... - Affiliation No. 831330" and
#: then rejected the school's own correct answer for having the wrong PIN.
_NOT_A_PIN = re.compile(
    r"(affiliation|affl?|udise|registration|reg|school|cbse|icse|board|code|"
    r"no\.?|number|#)\s*[:.\-]?\s*$", re.I)
_EMAIL = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
_URL = re.compile(r"https?://[^\s,;]+")


def clean(text):
    """One line, single spaces, no control characters."""
    if text is None:
        return ""
    text = unicodedata.normalize("NFKC", str(text))
    text = text.replace("\u200b", "").replace("\xa0", " ")
    return _WS.sub(" ", text).strip()


def is_blank_answer(value):
    return clean(value).lower() in BLANK_ANSWERS


def tidy_phone(text):
    """A phone number as text, or "" when it is not one.

    Indian numbers only: 10 digits, or 11-13 with a country or trunk prefix.
    Anything else - a PIN code, a house number, a year - is rejected rather
    than stored in the phone column.
    """
    raw = clean(text)
    if not raw:
        return ""
    first = re.split(r"[;,/]", raw)[0]
    kept = _PHONE_OK.sub("", first)
    digits = kept.lstrip("+")
    if len(digits) == 10 and digits[0] in "6789":
        return digits
    if 11 <= len(digits) <= 13:
        return ("+" + digits) if kept.startswith("+") else digits
    return ""


def tidy_pin(text):
    """A six digit Indian PIN code, or "".

    Rejects a number that something in front of it says is an affiliation,
    UDISE or registration number, and a number that is part of a longer run
    of digits such as a phone number.
    """
    body = clean(text)
    for found in _PIN.finditer(body):
        before = body[:found.start()].rstrip()
        if _NOT_A_PIN.search(before[-30:]):
            continue
        return found.group(1)
    return ""


def tidy_email(text):
    found = _EMAIL.search(clean(text))
    return found.group(0).lower() if found else ""


def tidy_url(text):
    raw = clean(text)
    found = _URL.search(raw)
    if found:
        return found.group(0).rstrip(".,);")
    if raw and "." in raw and " " not in raw and not raw.lower().startswith("no"):
        return "https://" + raw.lstrip("/")
    return ""


def domain_of(url):
    try:
        host = urllib.parse.urlparse(tidy_url(url)).netloc.lower()
    except ValueError:
        return ""
    return host[4:] if host.startswith("www.") else host


def blank_record():
    return {column: "" for column in SCHOOL_COLUMNS}


# ===========================================================================
#  DUPLICATE DETECTION
# ===========================================================================
#  Two records are the same school only on evidence: the same Maps place, or
#  the same website AND PIN code, or the same name AND the same address. A
#  similar name alone is never enough, and anything that says "different
#  branch" vetoes the merge even when a weaker key matched.
# ---------------------------------------------------------------------------

_BRANCH_WORDS = re.compile(
    r"\b(campus|branch|block|sector|phase|road|nagar|extension|colony|puram|"
    r"layout|north|south|east|west|main|annexe|annex|junior|senior|primary|"
    r"pre[- ]?primary|boys|girls|i{1,3}|iv|v)\b", re.I)
_NOISE_WORDS = re.compile(
    r"\b(the|school|schools|public|international|academy|vidyalaya|vidya|"
    r"mandir|high|higher|secondary|senior|sr|hr|convent|english|medium|"
    r"cbse|icse|isc|ib|igcse|state|board|trust|society|pvt|private|ltd)\b",
    re.I)


def name_key(name):
    """A school name reduced to its distinguishing words."""
    text = clean(name).lower()
    text = re.sub(r"\(.*?\)", " ", text)
    text = re.sub(r"[^a-z0-9 ]+", " ", text)
    text = _NOISE_WORDS.sub(" ", text)
    return _WS.sub(" ", text).strip()


def branch_signature(record):
    """The words that say WHICH campus this is, from the name and locality."""
    text = "%s %s" % (record.get("School Name", ""),
                      record.get("Area / Locality", ""))
    return frozenset(w.lower() for w in _BRANCH_WORDS.findall(text))


#: Columns whose value, on its own, says "this is the same school".
#: A school's own address, number or site belongs to one school, so two
#: records carrying the same one are the same school - whichever city or
#: keyword found them. Checked against master_data.xlsx as well as
#: against the city being collected.
EMAIL_COLUMNS = ("School Official Email", "School Alternate Email",
                 "Principal Official Email")
PHONE_COLUMNS = ("School Official Phone Number",
                 "School Alternate Phone Number",
                 "Principal Office Phone Number")


def phone_key(text):
    """One phone number reduced to the form two records can be compared on.

    The column keeps the number as it was found, prefix and all, so
    +91 98111 22333, 09811122333 and 9811122333 are three different
    strings there. They are one telephone, so the key is the last ten
    digits - the subscriber number - and nothing else.
    """
    number = tidy_phone(text)
    if not number:
        return ""
    digits = re.sub(r"\D", "", number)
    return digits[-10:] if len(digits) >= 10 else ""


def contact_keys(record):
    """Keys made from the ways a school can be contacted.

    One per email, one per phone number and one for the website. These
    are strong, but not infallible: a group of schools can share a
    reception number or a single domain across every branch. The
    contradiction guard is what keeps two branches apart - a different
    PIN code, a different Maps place or a different branch word stops the
    merge even when the number matches.
    """
    keys = []
    for column in EMAIL_COLUMNS:
        address = tidy_email(record.get(column, ""))
        if address:
            keys.append(("email", address))
    site = domain_of(record.get("School Website", ""))
    if site:
        keys.append(("site", site))
    for column in PHONE_COLUMNS:
        number = phone_key(record.get(column, ""))
        if number:
            keys.append(("phone", number))
    return keys


def match_keys(record):
    """Every key this record can be recognised by, strongest first."""
    keys = []
    place = clean(record.get("School Google Maps Link", ""))
    source_id = place_id_of(place)
    if source_id:
        keys.append(("place", source_id))
    # the same email, website or phone number means the same school
    keys.extend(contact_keys(record))
    site, pin = domain_of(record.get("School Website", "")), tidy_pin(
        record.get("PIN Code", ""))
    if site and pin:
        keys.append(("site+pin", "%s|%s" % (site, pin)))
    nkey = name_key(record.get("School Name", ""))
    address = clean(record.get("School Address", "")).lower()
    if nkey and address:
        keys.append(("name+address", "%s|%s" % (nkey, address[:60])))
    if nkey and pin:
        keys.append(("name+pin", "%s|%s" % (nkey, pin)))
    # The weakest key, and the reason it exists: a copied entry often has a
    # name and nothing else. Without this such a record has NO key at all,
    # so the duplicate check cannot see it and the "is it saved?" check can
    # never find it - which left a city open for ever. It is last, so every
    # stronger key is tried first, and the contradiction guards still apply.
    plain = clean(record.get("School Name", "")).lower()
    if nkey or plain:
        keys.append(("name+city", "%s|%s" % (nkey or plain,
                                             clean(record.get("City", "")).lower())))
    return keys


def place_id_of(url):
    """The stable identifier inside a Google Maps place link, if it has one."""
    text = clean(url)
    for pattern in (r"!1s(0x[0-9a-f]+:0x[0-9a-f]+)", r"\bplace_id[:=]([\w-]+)",
                    r"/maps/place/[^/]+/data=[^/]*?(0x[0-9a-f]+:0x[0-9a-f]+)"):
        found = re.search(pattern, text, re.I)
        if found:
            return found.group(1).lower()
    return ""


def contradicts(left, right):
    """A reason these two records are different schools, or ""."""
    left_id, right_id = (place_id_of(left.get("School Google Maps Link", "")),
                         place_id_of(right.get("School Google Maps Link", "")))
    if left_id and right_id and left_id != right_id:
        return "different Google Maps places"

    left_pin, right_pin = tidy_pin(left.get("PIN Code", "")), tidy_pin(
        right.get("PIN Code", ""))
    if left_pin and right_pin and left_pin != right_pin:
        return "different PIN codes (%s / %s)" % (left_pin, right_pin)

    left_branch, right_branch = branch_signature(left), branch_signature(right)
    if left_branch and right_branch and left_branch != right_branch:
        only = (left_branch ^ right_branch)
        return "different branch words (%s)" % ", ".join(sorted(only))

    left_address = clean(left.get("School Address", "")).lower()
    right_address = clean(right.get("School Address", "")).lower()
    if left_address and right_address:
        shared = set(left_address.split()) & set(right_address.split())
        if len(shared) < 2 and left_address[:25] != right_address[:25]:
            return "addresses do not overlap"
    return ""


def merge_records(kept, incoming):
    """Fill the blanks of ``kept`` from ``incoming``. Never overwrites."""
    for column in SCHOOL_COLUMNS:
        if not clean(kept.get(column, "")) and clean(incoming.get(column, "")):
            kept[column] = incoming[column]
    return kept


class SchoolIndex:
    """The unique schools of one city, with their lookup keys."""

    def __init__(self):
        self.records = []
        self._by_key = {}

    def __len__(self):
        return len(self.records)

    def find(self, record):
        """``(existing_record, why)`` when this is already known."""
        for kind, key in match_keys(record):
            existing = self._by_key.get((kind, key))
            if existing is None:
                continue
            reason = contradicts(existing, record)
            if reason:
                LOG.debug("  not merged (%s): %s", reason,
                          clean(record.get("School Name", ""))[:60])
                continue
            return existing, kind
        return None, ""

    def add(self, record):
        """``(record, "new" | <match kind>)``."""
        existing, why = self.find(record)
        if existing is not None:
            merge_records(existing, record)
            self._index(existing)
            return existing, why
        self.records.append(record)
        self._index(record)
        return record, "new"

    def _index(self, record):
        for key in match_keys(record):
            self._by_key.setdefault(key, record)


# ===========================================================================
#  GEMINI ANSWER PARSER
# ===========================================================================

def _key_pattern():
    keys = sorted(set(ANSWER_KEYS) | set(KEY_SYNONYMS), key=len, reverse=True)
    # No "start of word" guard: Gemini sometimes runs two fields together as
    # "...ABC SchoolPrincipal_Name: R. Rao", and a guard would hide the second
    # field and let its label and value be swallowed by the first one's value.
    # The labels are distinctive enough, and the longest match at a position
    # wins because the alternation is sorted longest-first.
    return re.compile(r"(" + "|".join(map(re.escape, keys))
                      + r")\s*[::]\s*", re.I)


_KEY_RE = _key_pattern()
_CANON_KEY = {k.lower(): k for k in ANSWER_KEYS}
_CANON_KEY.update({k.lower(): v for k, v in KEY_SYNONYMS.items()})


def parse_answer(text):
    """``{column: value}`` for every field the answer actually states.

    Scans for field labels anywhere in the text, not line by line: Gemini
    sometimes joins two fields onto one line, and a label glued to the end of
    a previous value would otherwise swallow the next field whole. A value
    runs up to the next label, so a genuinely multi-line value survives.
    """
    found = {}
    if not text:
        return found
    body = strip_prompt_echo(text)
    matches = list(_KEY_RE.finditer(body))
    for index, match in enumerate(matches):
        key = _CANON_KEY.get(match.group(1).lower())
        if key is None:
            continue
        stop = matches[index + 1].start() if index + 1 < len(matches) else len(body)
        value = body[match.end():stop]
        value = re.sub(r"\\n\s*$", "", value.strip()).strip()
        value = value.strip("*").strip()
        if is_blank_answer(value):
            continue
        column = KEY_TO_COLUMN.get(key)
        if column and column not in found:
            found[column] = clean(value)
    return found


def strip_prompt_echo(text):
    """Remove the prompt from a capture that contains it.

    The prompt's own format block is literally ``School_Name: Value``, so a
    capture that includes the question would otherwise parse as an answer of
    57 fields all reading "Value". Everything up to and including the last
    copy of the format block's final line is discarded.
    """
    if not text:
        return ""
    marker = "%s: Value" % ANSWER_KEYS[-1]
    cut = text.rfind(marker)
    if cut == -1:
        return text
    end = text.find("\n", cut)
    tail = text[end + 1:] if end != -1 else ""
    # only trust the tail if it still carries field labels
    return tail if _KEY_RE.search(tail) else text


def looks_like_an_answer(text, minimum=3):
    """True once the text carries at least ``minimum`` real field lines."""
    return len(parse_answer(text)) >= minimum


def answer_is_about(record, values):
    """``""`` when the answer belongs to this school, else why it does not."""
    said = name_key(values.get("School Name", ""))
    asked = name_key(record.get("School Name", ""))
    if said and asked:
        shared = set(said.split()) & set(asked.split())
        if not shared:
            return ("the answer is about %r, not %r"
                    % (clean(values.get("School Name"))[:60],
                       clean(record.get("School Name"))[:60]))
    asked_pin = tidy_pin(record.get("PIN Code", ""))
    said_pin = tidy_pin(values.get("PIN Code", ""))
    if asked_pin and said_pin and asked_pin != said_pin:
        return "the answer's PIN %s is not the school's PIN %s" % (said_pin,
                                                                   asked_pin)
    return ""


def apply_answer(record, values):
    """Write a validated answer onto a record. Returns how many fields landed.

    Google Maps evidence wins: a column this run actually observed is never
    replaced by something Gemini said, only filled when it is empty.
    """
    written = 0
    for column, value in values.items():
        if column not in record:
            continue
        value = tidy_value(column, value)
        if not value:
            continue
        # GEMINI NEVER FILLS THE OFFICIAL CONTACTS.
        # Google Search owns those two columns, and the prompt no longer
        # even asks for them. Should an answer mention one anyway, it is
        # moved to the matching ALTERNATE column - which is what it is: a
        # second-hand suggestion, kept but not presented as verified. A
        # blank official column is left blank rather than promoted.
        if OFFICIAL_CONTACT_SEARCH and column in OFFICIAL_CONTACT_COLUMNS:
            column = OFFICIAL_TO_ALTERNATE.get(column, column)
            value = tidy_value(column, value)
            if not value or column not in record:
                continue
        if column in SOURCE_EVIDENCE_COLUMNS and clean(record.get(column, "")):
            continue
        if clean(record.get(column, "")):
            continue
        record[column] = value
        written += 1
    return written


def tidy_value(column, value):
    """Column-appropriate formatting, applied once, on the way in."""
    text = clean(value)
    if not text or is_blank_answer(text):
        return ""
    if "Email" in column:
        return tidy_email(text)
    if "Phone" in column:
        return tidy_phone(text)
    if column == "PIN Code":
        return tidy_pin(text)
    if "Website" in column or "URL" in column or "Link" in column:
        return tidy_url(text) or text
    if column in ("UDISE Code", "Affiliation Number",
                  "School Registration Number"):
        kept = re.sub(r"[^0-9A-Za-z/-]", "", text)
        return kept if any(ch.isdigit() for ch in kept) else ""
    return text

# ===========================================================================
#  EXCEL
# ===========================================================================
#  current_data.xlsx fills to CYCLE_SIZE, then current_data_cycle1.xlsx,
#  _cycle2 and so on. A finished cycle file is never reopened for writing, so
#  nothing can overwrite completed work. Every save is atomic: written to a
#  temporary file in the same folder, then os.replace()d into place, so a
#  crash leaves either the old workbook or the new one, never half of one.
# ---------------------------------------------------------------------------

try:
    import openpyxl
    from openpyxl.styles import Font
    OPENPYXL_AVAILABLE = True
except ImportError:                                     # pragma: no cover
    openpyxl = None
    Font = None
    OPENPYXL_AVAILABLE = False


class ExcelUnavailable(RuntimeError):
    pass


def require_openpyxl():
    if not OPENPYXL_AVAILABLE:
        raise ExcelUnavailable(
            "openpyxl is not installed. Run:  python3 -m pip install openpyxl")


# --- the workbooks ---------------------------------------------------------
EXCEL_BASENAME = "current_data"
MASTER_BASENAME = "master_data"
#: Written before the master and the cycle file are updated, removed after.
#: Its presence on start-up means a run died between the two writes.
COMMIT_JOURNAL = "pending_commit.json"


def cycle_path(cycle):
    """0 -> current_data.xlsx, 1 -> current_data_cycle1.xlsx, ..."""
    name = EXCEL_BASENAME if cycle == 0 else "%s_cycle%d" % (EXCEL_BASENAME,
                                                             cycle)
    return os.path.join(EXCEL_DIR, name + ".xlsx")


def master_path():
    """Every unique school this rank has ever saved."""
    return os.path.join(EXCEL_DIR, MASTER_BASENAME + ".xlsx")


def journal_path():
    return os.path.join(EXCEL_DIR, COMMIT_JOURNAL)


def new_workbook():
    require_openpyxl()
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Schools"
    sheet.append(list(SCHOOL_COLUMNS))
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    sheet.freeze_panes = "A2"
    return book


def load_workbook_safely(path):
    """The workbook at ``path``, or a new one when it is missing or damaged.

    A damaged workbook is moved aside rather than deleted, so nothing the
    operator collected is ever thrown away by this program.
    """
    require_openpyxl()
    if not os.path.exists(path):
        return new_workbook()
    try:
        book = openpyxl.load_workbook(path)
        sheet = book.active
        header = [cell.value for cell in sheet[1]] if sheet.max_row else []
        if [clean(h) for h in header] != list(SCHOOL_COLUMNS):
            raise ValueError("unexpected column headings")
        return book
    except Exception as error:                          # noqa: BLE001
        spoiled = "%s.damaged_%s" % (path, time.strftime("%Y%m%d_%H%M%S"))
        shutil.move(path, spoiled)
        LOG.error("%s could not be opened (%s). It has been kept as %s and a "
                  "new workbook started.", os.path.basename(path),
                  str(error)[:120], os.path.basename(spoiled))
        return new_workbook()


def save_workbook_atomically(book, path):
    require_openpyxl()
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    handle, temporary = tempfile.mkstemp(dir=os.path.dirname(path) or ".",
                                         suffix=".xlsx.tmp")
    os.close(handle)
    try:
        book.save(temporary)
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.remove(temporary)


def row_count(path):
    """Data rows in a workbook (header excluded). 0 when it does not exist."""
    if not os.path.exists(path):
        return 0
    require_openpyxl()
    try:
        book = openpyxl.load_workbook(path, read_only=True)
        rows = max(book.active.max_row - 1, 0)
        book.close()
        return rows
    except Exception:                                   # noqa: BLE001
        return 0


def read_records(path):
    """Every saved row of one workbook, as records."""
    if not os.path.exists(path):
        return []
    require_openpyxl()
    book = openpyxl.load_workbook(path, read_only=True)
    sheet = book.active
    records = []
    for values in sheet.iter_rows(min_row=2, values_only=True):
        record = blank_record()
        for column, value in zip(SCHOOL_COLUMNS, values):
            record[column] = "" if value is None else str(value)
        if clean(record["School Name"]):
            records.append(record)
    book.close()
    return records


def write_records(path, records):
    """Replace a workbook with exactly these records. Atomic."""
    book = new_workbook()
    sheet = book.active
    for record in records:
        sheet.append([record.get(column, "") for column in SCHOOL_COLUMNS])
        for index, column in enumerate(SCHOOL_COLUMNS, start=1):
            if column in TEXT_COLUMNS:
                sheet.cell(row=sheet.max_row, column=index).number_format = "@"
    save_workbook_atomically(book, path)
    return len(records)


def finished_cycle_numbers():
    """The cycle files that already exist, in order."""
    numbers = []
    if not os.path.isdir(EXCEL_DIR):
        return numbers
    pattern = re.compile(r"^%s_cycle(\d+)\.xlsx$" % re.escape(EXCEL_BASENAME))
    for name in os.listdir(EXCEL_DIR):
        found = pattern.match(name)
        if found:
            numbers.append(int(found.group(1)))
    return sorted(numbers)


def all_cycle_paths():
    """Every finished cycle file, then the active one. Existing files only."""
    paths = [cycle_path(number) for number in finished_cycle_numbers()]
    if os.path.exists(cycle_path(0)):
        paths.append(cycle_path(0))
    return paths


def stable_identity(record):
    """Name and city: the two things research never changes.

    Gemini fills in a fuller address than the results list showed, so a key
    built on the address stops matching the moment a record is enriched.
    Asking "did this school reach the master?" with such a key reported
    schools as missing while they sat in the file, and the city never
    closed. The name is source evidence and is never overwritten, so that is
    what this check uses.
    """
    return "%s|%s" % (name_key(record.get("School Name", ""))
                      or clean(record.get("School Name", "")).lower(),
                      clean(record.get("City", "")).lower())


def master_match_keys():
    """Every key the master file already holds - the duplicate check."""
    keys = set()
    for record in read_records(master_path()):
        for key in match_keys(record):
            keys.add(key)
    return keys


def master_key_map(records=None):
    """key -> the master record holding it.

    The keys alone say "something in the master matches". They cannot say
    whether it is the SAME school, and with the contact keys that matters:
    every branch of one group shares a website, and sometimes a phone
    number. Keeping the record lets the same contradiction guard the city
    index uses run here too, so a different branch is still saved.
    """
    keymap = {}
    for record in (read_records(master_path()) if records is None
                   else records):
        for key in match_keys(record):
            keymap.setdefault(key, record)
    return keymap


def found_in_master(record, keymap):
    """The master record this one duplicates, or None.

    A key match is not enough on its own: if the two records contradict
    each other - a different PIN code, a different Google Maps place, a
    different branch word - they are different schools and this one still
    has to be saved.
    """
    for key in match_keys(record):
        existing = keymap.get(key)
        if existing is None:
            continue
        reason = contradicts(existing, record)
        if reason:
            LOG.debug("   same %s but %s: %s kept separate", key[0], reason,
                      clean(record.get("School Name", ""))[:50])
            continue
        return existing
    return None


def master_identities():
    """``(stable identities, match keys)`` for everything in the master."""
    identities, keys = set(), set()
    for record in read_records(master_path()):
        identities.add(stable_identity(record))
        keys.update(match_keys(record))
    return identities, keys


def missing_from_master(records):
    """The records that genuinely are not in the master yet.

    A school with no phone number and no email is left out on purpose, so
    it is not "missing" - counting it would keep its city open for ever.
    """
    # The SAME test the save path uses. If these two ever disagree, a
    # school rejected as a duplicate when saving is reported missing here
    # for ever, and its city can never close.
    master = read_records(master_path())
    identities = {stable_identity(r) for r in master}
    keymap = master_key_map(master)
    return [r for r in records
            if has_contact(r)
            and stable_identity(r) not in identities
            and found_in_master(r, keymap) is None]


def existing_match_keys():
    """Keys across the master and every cycle file. Used by the repair."""
    keys = set(master_match_keys())
    for path in all_cycle_paths():
        for record in read_records(path):
            for key in match_keys(record):
                keys.add(key)
    return keys


# ===========================================================================
#  SAVING RECOVERABLE PROGRESS TO GITHUB
# ===========================================================================
#  The point of this is one thing only: if a production laptop dies, the
#  work it had already done must not die with it. So the files that let a
#  fresh clone carry on - the checkpoint, the Excel files, the cycle
#  backups and the few temporary files an unfinished city needs - are
#  committed and pushed on a timer and at the natural milestones.
#
#  What it must never do is put the collection at risk. GitHub being down,
#  a rejected push, no network at all: every one of those is logged, noted
#  in the checkpoint as NOT synced, and then ignored so the run carries on.
#
#  Rules followed here, all of them deliberate:
#    * only named paths are added - never "git add ." - so a credential
#      file, a browser profile or a log cannot be committed by accident
#    * no repository is created and no git config is written
#    * no force-push, no reset, no clean, no checkout of existing branches
#    * each rank pushes its OWN branch, so 25 laptops cannot collide
#    * a push that did not happen is never recorded as if it had
# ---------------------------------------------------------------------------

GIT_STATE = {"last_attempt": 0.0, "last_ok": 0.0, "pending": False,
             "last_error": "", "branch": "", "checked": False,
             "usable": False, "why": ""}


# ===========================================================================
#  WORKING HOURS: RANDOM WORK PERIODS, BREAKS AND THE NIGHT STOP
# ===========================================================================
#  Modelled on the GitHub Profile Extractor's clock, which was read for
#  reference and not modified. Three ideas, all of them from it:
#
#    * a work period is a fresh random draw, not a fixed number, and a
#      break is drawn fresh too, so no two are the same length;
#    * the clock counts WORKING time - it stops while a break runs, so a
#      3-hour period means three hours of actual collecting;
#    * a stop is only ever taken at a safe point, between schools or
#      between cities, never in the middle of an Excel write.
#
#  The night stop is read on the TIME_ZONE clock so every laptop agrees
#  whatever the machine is set to.
# ---------------------------------------------------------------------------

WORK = {"started": None, "banked": 0.0, "mark": None, "depth": 0,
        "period_hours": 0.0, "breaks": 0, "break_until": 0.0}


def local_now():
    """Now, on the configured clock. Falls back to the machine's own."""
    try:
        import zoneinfo
        return datetime.datetime.now(zoneinfo.ZoneInfo(TIME_ZONE))
    except Exception:                                   # noqa: BLE001
        return datetime.datetime.now()


def pick_work_period_hours():
    """A fresh period length in hours; 0 when the cycle is switched off."""
    low = max(0.0, float(WORK_PERIOD_MIN_HOURS))
    high = max(low, float(WORK_PERIOD_MAX_HOURS))
    if high <= 0:
        return 0.0
    if high == low:
        return high
    return random.randint(int(low * 10), int(high * 10)) / 10.0


def pick_break_seconds():
    """A fresh break length in seconds, from the configured range."""
    low = max(0, int(BREAK_MIN_MINUTES)) * 60
    high = max(low, int(BREAK_MAX_MINUTES) * 60)
    return low if high <= low else random.randint(low, high)


def reset_work_timer():
    """Begin a fresh work period. At startup and after every break."""
    WORK["started"] = time.monotonic()
    WORK["banked"] = 0.0
    WORK["mark"] = time.monotonic()
    WORK["depth"] = 0
    WORK["period_hours"] = pick_work_period_hours()
    return WORK["period_hours"]


def work_clock_pause():
    """Stop counting working time. Safe to nest."""
    WORK["depth"] += 1
    if WORK["depth"] == 1 and WORK["mark"] is not None:
        WORK["banked"] += time.monotonic() - WORK["mark"]
        WORK["mark"] = None


def work_clock_resume():
    """Count working time again. Safe to nest."""
    if WORK["depth"] > 0:
        WORK["depth"] -= 1
    if WORK["depth"] == 0 and WORK["mark"] is None:
        WORK["mark"] = time.monotonic()


def work_seconds():
    """Seconds actually spent working since the last break."""
    if WORK["started"] is None:
        return 0.0
    total = WORK["banked"]
    if WORK["mark"] is not None:
        total += time.monotonic() - WORK["mark"]
    return total


def night_seconds_remaining(now=None):
    """Seconds until the night window ends, or 0 outside it.

    The one place the window is decided. It may cross midnight, and
    START == END switches it off rather than meaning "all day".
    """
    if not NIGHT_PAUSE_ENABLED:
        return 0
    now = now or local_now()
    start = int(NIGHT_START_HOUR) % 24
    end = int(NIGHT_END_HOUR) % 24
    if start == end:
        return 0
    inside = (start <= now.hour < end) if start < end else (
        now.hour >= start or now.hour < end)
    if not inside:
        return 0
    target = now.replace(hour=end, minute=0, second=0, microsecond=0)
    if target <= now:
        target += datetime.timedelta(days=1)
    return max(0, int((target - now).total_seconds()))


def due_pause(now=None):
    """``(reason, seconds)`` when a scheduled stop is due, else None.

    Night is checked first, so a work period ending at 22:59 does not take
    a break and then immediately take the night stop as well.
    """
    night = night_seconds_remaining(now)
    if night > 0:
        return ("night stop %02d:00-%02d:00 %s"
                % (int(NIGHT_START_HOUR) % 24, int(NIGHT_END_HOUR) % 24,
                   TIME_ZONE), night)
    period = WORK["period_hours"]
    if WORK["started"] is None or period <= 0:
        return None
    worked = work_seconds()
    if worked < period * 3600.0:
        return None
    seconds = pick_break_seconds()
    if seconds <= 0:
        return None
    return ("worked %.2f h of a %.1f h period - taking a %d min break"
            % (worked / 3600.0, period, seconds // 60), seconds)


def sleep_through(seconds, slice_seconds=5.0):
    """Wait, but notice Ctrl+C. Returns False when a stop was asked for.

    A long break must not swallow a shutdown request, so the wait is taken
    in slices and the stop flag is checked between them.
    """
    deadline = time.monotonic() + max(0.0, seconds)
    while time.monotonic() < deadline:
        if RUNTIME["stop"]:
            return False
        left = deadline - time.monotonic()
        say_status("%s - resuming in %s" % (WORK.get("why", "break"),
                                            fmt_span(left)))
        time.sleep(min(slice_seconds, max(0.1, left)))
    return not RUNTIME["stop"]


def fmt_span(seconds):
    """"1 h 05 m" / "12 m 30 s", for a countdown a human reads."""
    seconds = int(max(0, seconds))
    hours, rest = divmod(seconds, 3600)
    minutes, secs = divmod(rest, 60)
    if hours:
        return "%d h %02d m" % (hours, minutes)
    if minutes:
        return "%d m %02d s" % (minutes, secs)
    return "%d s" % secs


def take_scheduled_pause(director, checkpoint, reason, seconds):
    """A safe stop: save, sync, release the browser, wait, come back.

    Called only between schools or between cities, so nothing is half
    written when it runs. The order is deliberate - the progress is on
    disk and pushed BEFORE the wait, because a laptop that dies during a
    two-hour break must lose nothing.
    """
    seconds = int(max(0, seconds))
    if seconds <= 0:
        return True
    WORK["why"] = reason.split(" - ")[0][:40]
    resume_at = local_now() + datetime.timedelta(seconds=seconds)
    LOG.info("")
    LOG.info("[PAUSE] %s", reason)
    LOG.info("        waiting %s, resuming about %s",
             fmt_span(seconds), resume_at.strftime("%H:%M:%S"))

    if checkpoint is not None:
        WORK["break_until"] = time.time() + seconds
        try:
            checkpoint.data["work"] = {
                "period_hours": WORK["period_hours"],
                "breaks_taken": WORK["breaks"],
                "break_until": datetime.datetime.fromtimestamp(
                    WORK["break_until"]).isoformat(timespec="seconds"),
                "reason": reason[:120],
            }
            checkpoint.save()
        except Exception:                               # noqa: BLE001
            pass
    git_sync("before a %s pause" % WORK["why"], checkpoint, force=True)

    if director is not None and seconds >= PAUSE_CLOSES_BROWSER_SECONDS:
        LOG.info("        closing the browser for the wait")
        try:
            director.close()
        except Exception:                               # noqa: BLE001
            pass

    work_clock_pause()
    try:
        finished = sleep_through(seconds)
    finally:
        work_clock_resume()
    STATUS.clear()

    WORK["breaks"] += 1
    WORK["break_until"] = 0.0
    reset_work_timer()
    if finished:
        LOG.info("[RESUME] back to work - next period %.1f h",
                 WORK["period_hours"])
    else:
        LOG.warning("[PAUSE] cut short by your stop request")
    if checkpoint is not None:
        try:
            checkpoint.data["work"] = {
                "period_hours": WORK["period_hours"],
                "breaks_taken": WORK["breaks"], "break_until": "",
                "reason": "",
            }
            checkpoint.save()
        except Exception:                               # noqa: BLE001
            pass
    return finished


def schedule_gate(director, checkpoint):
    """Take a scheduled stop if one is due. A safe point calls this.

    Returns False when the run should stop altogether.
    """
    if RUNTIME["stop"]:
        return False
    stop = due_pause()
    if stop is None:
        return True
    reason, seconds = stop
    return take_scheduled_pause(director, checkpoint, reason, seconds)


def work_break_status():
    """One line for the banner and --status."""
    low = max(0.0, float(WORK_PERIOD_MIN_HOURS))
    high = max(low, float(WORK_PERIOD_MAX_HOURS))
    if high <= 0:
        return "no scheduled breaks"
    drawn = WORK["period_hours"]
    period = ("%.1f h drawn for this period" % drawn if drawn > 0
              else "%.0f-%.0f h periods" % (low, high))
    return "%s, then %d-%d min off" % (period, int(BREAK_MIN_MINUTES),
                                       int(BREAK_MAX_MINUTES))


def night_pause_status():
    """One line about the night stop."""
    if not NIGHT_PAUSE_ENABLED:
        return "off (set the hours and NIGHT_PAUSE_ENABLED = True)"
    start = int(NIGHT_START_HOUR) % 24
    end = int(NIGHT_END_HOUR) % 24
    if start == end:
        return "off (start hour == end hour)"
    left = night_seconds_remaining()
    inside = " - IN IT NOW, %s left" % fmt_span(left) if left else ""
    return "%02d:00-%02d:00 %s%s" % (start, end, TIME_ZONE, inside)


def git_remote_now():
    """The remote this run pushes to, resolved on first use and kept.

    run() resolves it at startup, but git_sync() must not depend on having
    been reached through run(): any other entry point would otherwise fall
    back to the setting at the top of this file and ignore rank.json.
    """
    if not RUNTIME.get("git_remote_from"):
        RUNTIME["git_remote"], RUNTIME["git_remote_from"] = resolve_git_remote()
    return RUNTIME["git_remote"]


def git_run(args, timeout=120, index=None):
    """One git command in the project directory. Never raises.

    Returns ``(ok, output)``. Git is given the project folder explicitly
    because the run's working directory is not guaranteed to be it.
    ``index`` points git at a throwaway index file, so staging the data
    files does not disturb whatever the operator has staged.
    """
    env = None
    if index:
        env = dict(os.environ)
        env["GIT_INDEX_FILE"] = index
    try:
        done = subprocess.run(["git", "-C", PROJECT_DIR] + list(args),
                              capture_output=True, text=True, timeout=timeout,
                              env=env)
    except (OSError, subprocess.SubprocessError) as error:
        return False, str(error)[:300]
    output = (done.stdout or "") + (done.stderr or "")
    return done.returncode == 0, output.strip()


def git_branch_name():
    """The branch this rank pushes to."""
    if not GIT_BRANCH_PER_RANK:
        ok, name = git_run(["rev-parse", "--abbrev-ref", "HEAD"])
        return name.strip() if ok and name.strip() != "HEAD" else ""
    try:
        this_rank = int(RUNTIME.get("rank") or rank)
    except (TypeError, ValueError):
        this_rank = rank
    return "school-data-rank-%d" % this_rank


def git_usable():
    """Whether syncing can work at all, with the reason when it cannot.

    Checked once per run: git present, this really is a repository, a
    remote exists, and the branch is known. Nothing is created or
    configured here - a missing piece is reported for the operator to fix.
    """
    if GIT_STATE["checked"]:
        return GIT_STATE["usable"], GIT_STATE["why"]
    GIT_STATE["checked"] = True
    GIT_STATE["usable"], GIT_STATE["why"] = False, ""

    if not GIT_SYNC_ENABLED:
        GIT_STATE["why"] = "off (GIT_SYNC_ENABLED = False)"
        return False, GIT_STATE["why"]
    if shutil.which("git") is None:
        GIT_STATE["why"] = "git is not installed"
        return False, GIT_STATE["why"]
    ok, _out = git_run(["rev-parse", "--git-dir"])
    if not ok:
        GIT_STATE["why"] = ("%s is not a git repository - create one and add "
                            "a remote yourself; this script will not do it"
                            % PROJECT_DIR)
        return False, GIT_STATE["why"]
    ok, remotes = git_run(["remote"])
    remote = git_remote_now()
    if not ok or remote not in (remotes or "").split():
        GIT_STATE["why"] = ("no %r remote - add one with: git remote add %s "
                            "<url>" % (remote, remote))
        return False, GIT_STATE["why"]
    branch = git_branch_name()
    if not branch:
        GIT_STATE["why"] = "the branch could not be worked out"
        return False, GIT_STATE["why"]
    GIT_STATE["branch"] = branch
    GIT_STATE["usable"] = True
    GIT_STATE["why"] = "on, pushing %s to %s" % (branch, git_remote_now())
    return True, GIT_STATE["why"]


def recovery_paths():
    """Exactly the files a fresh clone needs to carry on. Nothing else.

    Named one by one on purpose. The capture text and the per-school raw
    answers are left out: they are evidence, they can be large, and a
    resume does not need them. The temporary workbook and the resume file
    of an UNFINISHED city are included, because without them a half-done
    city would start again from its first keyword.
    """
    paths = [CHECKPOINT_FILE]
    for folder, pattern in ((EXCEL_DIR, "*.xlsx"), (BACKUP_DIR, "*.xlsx")):
        if os.path.isdir(folder):
            paths.extend(sorted(glob.glob(os.path.join(folder, pattern))))
    if os.path.isdir(TEMP_DIR):
        for pattern in ("*_temporary_schools.xlsx", "*_schools.json"):
            paths.extend(sorted(glob.glob(os.path.join(TEMP_DIR, pattern))))
    here = os.path.abspath(PROJECT_DIR)
    keep = []
    for path in paths:
        if not os.path.exists(path):
            continue
        # never outside the project, whatever a setting says
        if not os.path.abspath(path).startswith(here + os.sep):
            continue
        keep.append(os.path.relpath(os.path.abspath(path), here))
    return keep


def git_note_state(checkpoint=None):
    """Record the sync state where a recovery can read it."""
    if checkpoint is None:
        return
    try:
        checkpoint.data["git"] = {
            "branch": GIT_STATE["branch"],
            "pending": bool(GIT_STATE["pending"]),
            "last_ok": (datetime.datetime.fromtimestamp(
                GIT_STATE["last_ok"]).isoformat(timespec="seconds")
                if GIT_STATE["last_ok"] else ""),
            "last_error": GIT_STATE["last_error"][:300],
        }
        checkpoint.save()
    except Exception:                                   # noqa: BLE001
        pass


def git_sync(reason, checkpoint=None, force=False):
    """Commit and push the recovery files. Returns True only if PUSHED.

    Called at the milestones and on a timer. Everything that can go wrong
    is caught: the collection loop must not stop because GitHub is down.
    A failure sets "pending", which the next call retries.
    """
    if not GIT_SYNC_ENABLED:
        return False
    usable, why = git_usable()
    if not usable:
        return False

    now = time.time()
    due = (now - GIT_STATE["last_attempt"]) >= max(0, GIT_SYNC_MINUTES) * 60
    if not (force or due or GIT_STATE["pending"]):
        return False
    GIT_STATE["last_attempt"] = now

    paths = recovery_paths()
    if not paths:
        return False

    branch = GIT_STATE["branch"]
    # THE DATA NEVER TOUCHES THE SOURCE BRANCH.
    #
    # This used to commit onto whatever branch was checked out and push
    # that. When a push failed - no network, no key - the commit stayed
    # behind on master, and the next ordinary "git push origin master"
    # would have carried the collected school records into the public
    # source repository. On a public repo that is a one-way mistake.
    #
    # So the commit is built with plumbing instead: a throwaway index, a
    # tree written from it, and a commit recorded straight onto the rank
    # branch's ref. HEAD, the checked-out branch and the operator's own
    # staged changes are never touched, and nothing about the data can
    # reach the source branch even if every push fails for a week.
    ref = "refs/heads/%s" % branch
    # git_run returns (ok, output): the FIRST value says whether the ref
    # exists, the second is its hash.
    has_ref, parent_sha = git_run(["rev-parse", "--verify", "--quiet", ref])
    parent_sha = parent_sha.strip() if has_ref else ""
    if not parent_sha:
        # A FRESH CLONE OF A RANK THAT HAS SYNCED BEFORE.
        #
        # The data branch exists on the remote but not in this clone - a
        # re-imaged laptop, or a new install directory for an existing
        # rank. Building on HEAD here would make the push a
        # non-fast-forward, which is REJECTED, and the retry builds on
        # HEAD again, so the rank would never sync at all.
        #
        # So bring the branch down first and build on its tip. A remote
        # that has no such branch fails this harmlessly and the first
        # sync starts from the source commit as before.
        got, _out = git_run(["fetch", git_remote_now(),
                             "%s:%s" % (ref, ref)], timeout=180)
        if got:
            has_ref, parent_sha = git_run(["rev-parse", "--verify",
                                           "--quiet", ref])
            parent_sha = parent_sha.strip() if has_ref else ""
            if parent_sha:
                LOG.info("   [GIT] %s already existed on %s - continuing "
                         "that branch", branch, git_remote_now())
    if not parent_sha:
        # first sync of this rank anywhere: build on the source commit
        ok_head, head_sha = git_run(["rev-parse", "--verify", "--quiet",
                                     "HEAD"])
        parent_sha = head_sha.strip() if ok_head else ""
    # A repository with no commits at all has neither; then the commit is
    # a root commit built on an empty tree. A laptop that ran "git init"
    # and nothing else lands here.
    base = parent_sha
    index = os.path.join(PROJECT_DIR, ".git", "school-sync-index")
    try:
        if os.path.exists(index):
            os.remove(index)
    except OSError:
        pass

    def fail(stage, detail):
        GIT_STATE["pending"] = True
        GIT_STATE["last_error"] = "%s failed: %s" % (stage, detail[:200])
        LOG.warning("   [GIT] %s", GIT_STATE["last_error"])
        git_note_state(checkpoint)
        return False

    # the rank branch's own tree (or the source commit, the first time)
    ok, out = git_run(["read-tree", base] if base else ["read-tree", "--empty"],
                      index=index)
    if not ok:
        return fail("read-tree", out)
    # -f because the collected data is deliberately in .gitignore
    ok, out = git_run(["add", "-f", "--"] + paths, index=index)
    if not ok:
        return fail("add", out)
    ok, tree = git_run(["write-tree"], index=index)
    if not ok or not tree.strip():
        return fail("write-tree", tree)
    tree = tree.strip()

    # nothing changed since the last sync: not a failure, just nothing to do
    if base:
        ok, head_tree = git_run(["rev-parse", "%s^{tree}" % base])
        if ok and head_tree.strip() == tree and not GIT_STATE["pending"]:
            return False

    message = "school data: %s (rank %s)" % (reason,
                                             RUNTIME.get("rank_label", "?"))
    making = ["commit-tree", tree, "-m", message]
    if parent_sha:
        making = ["commit-tree", tree, "-p", parent_sha, "-m", message]
    ok, commit = git_run(making)
    if not ok or not commit.strip():
        return fail("commit-tree", commit)
    ok, out = git_run(["update-ref", ref, commit.strip()])
    if not ok:
        return fail("update-ref", out)

    ok, out = git_run(["push", git_remote_now(),
                       "%s:%s" % (ref, ref)],
                      timeout=180)
    if not ok:
        GIT_STATE["pending"] = True
        GIT_STATE["last_error"] = "push failed: %s" % out[:200]
        LOG.warning("   [GIT] not pushed - will retry: %s",
                    GIT_STATE["last_error"])
        git_note_state(checkpoint)
        return False

    GIT_STATE["pending"] = False
    GIT_STATE["last_error"] = ""
    GIT_STATE["last_ok"] = time.time()
    LOG.info("   [GIT] pushed %d file(s) to %s/%s - %s",
             len(paths), git_remote_now(), branch, reason)
    git_note_state(checkpoint)
    return True


def git_status_line():
    """One line for --status and the startup banner."""
    if not GIT_SYNC_ENABLED:
        return "off (GIT_SYNC_ENABLED = False)"
    usable, why = git_usable()
    if not usable:
        return "cannot sync: %s" % why
    if GIT_STATE["pending"]:
        return "%s - LAST PUSH FAILED: %s" % (why, GIT_STATE["last_error"][:80])
    if GIT_STATE["last_ok"]:
        return "%s - last pushed %s" % (why, time.strftime(
            "%H:%M:%S", time.localtime(GIT_STATE["last_ok"])))
    return "%s - nothing pushed yet" % why


def backup_path(path):
    """Where a finished cycle file's permanent copy lives."""
    return os.path.join(BACKUP_DIR, os.path.basename(path))


def same_workbook(left, right):
    """True when two workbooks hold the same bytes."""
    try:
        with open(left, "rb") as one, open(right, "rb") as two:
            return hashlib.sha1(one.read()).digest() == \
                hashlib.sha1(two.read()).digest()
    except OSError:
        return False


def backup_completed_cycle(path, expect=None):
    """Keep a permanent, verified copy of one FINISHED cycle file.

    Called only after finalise_current_cycle() has written and checked the
    cycle file, so this never sees a part-finished one. The copy is read
    back and its rows counted before it is called a backup.

    Nothing is ever overwritten with different data: a backup that already
    exists and differs is left alone and reported, because the file on disk
    may be the only remaining copy of those records.

    Returns the backup's path, or "" when nothing was written.
    """
    if not BACKUP_COMPLETED_CYCLES or not path or not os.path.exists(path):
        return ""
    rows = row_count(path)
    if expect is not None and rows != expect:
        LOG.error("   NOT backed up: %s holds %d row(s), not the %d expected",
                  os.path.basename(path), rows, expect)
        return ""
    if rows <= 0:
        LOG.error("   NOT backed up: %s has no rows", os.path.basename(path))
        return ""

    os.makedirs(BACKUP_DIR, exist_ok=True)
    target = backup_path(path)
    if os.path.exists(target):
        if same_workbook(path, target):
            LOG.info("   already backed up: %s", os.path.basename(target))
            return target
        LOG.error("   %s already exists in Backup/ with DIFFERENT contents - "
                  "left untouched, nothing was overwritten",
                  os.path.basename(target))
        return ""

    temporary = target + ".tmp"
    try:
        shutil.copy2(path, temporary)
        os.replace(temporary, target)
    except OSError as error:
        LOG.error("   the backup of %s failed: %s - the cycle file itself is "
                  "untouched", os.path.basename(path), error)
        try:
            os.remove(temporary)
        except OSError:
            pass
        return ""

    # verified, not assumed: read the copy back and count its rows
    copied = row_count(target)
    if copied != rows:
        LOG.error("   the backup of %s came back with %d row(s), not %d - "
                  "removing the bad copy, the cycle file is untouched",
                  os.path.basename(path), copied, rows)
        try:
            os.remove(target)
        except OSError:
            pass
        return ""
    LOG.info("   backed up: Backup/%s (%d records, verified)",
             os.path.basename(target), copied)
    return target


def finalise_current_cycle():
    """current_data.xlsx is full: name it cycleN and start an empty one.

    Returns the finished file's path, or "" when there was nothing to close.
    A cycle number already on disk is never reused, so a finished file can
    not be overwritten.
    """
    active = cycle_path(0)
    rows = row_count(active)
    if rows < CYCLE_SIZE:
        return ""
    finished = finished_cycle_numbers()
    number = (finished[-1] + 1) if finished else 1
    while os.path.exists(cycle_path(number)):
        number += 1
    target = cycle_path(number)
    records = read_records(active)
    write_records(target, records)
    if row_count(target) != rows:
        LOG.error("%s did not come back with %d rows - the active file is "
                  "left as it is", os.path.basename(target), rows)
        return ""
    write_records(active, [])                 # a fresh, empty current file
    LOG.info("")
    LOG.info("   %s finalised with %d records", os.path.basename(target), rows)
    LOG.info("   a new empty %s has been started", os.path.basename(active))
    # The permanent copy is taken AFTER the cycle file is complete and
    # checked. A failure here is reported and changes nothing: the return
    # value, and so the existing behaviour of this function, is the same
    # either way.
    backup_completed_cycle(target, expect=rows)
    git_sync("cycle %s finalised" % os.path.basename(target))
    return target


def _write_journal(records):
    """Name what is about to be written, so a crash can be repaired."""
    os.makedirs(EXCEL_DIR, exist_ok=True)
    try:
        with open(journal_path(), "w", encoding="utf-8") as handle:
            json.dump({"when": time.strftime("%Y-%m-%d %H:%M:%S"),
                       "records": records}, handle, ensure_ascii=False)
            handle.flush()
            os.fsync(handle.fileno())
    except OSError:
        pass


def _clear_journal():
    try:
        os.remove(journal_path())
    except OSError:
        pass


def reconcile_master_and_cycles():
    """Repair a run that died between the master write and the cycle write.

    The two files are written one after the other, so a crash in between can
    leave a school in one and not the other. Nothing is deleted: a record
    found in only one file is copied into the other. Returns a short report.
    """
    if not OPENPYXL_AVAILABLE or not os.path.isdir(EXCEL_DIR):
        return {"added_to_master": 0, "added_to_current": 0, "journal": False}

    had_journal = os.path.exists(journal_path())
    master = read_records(master_path())
    # The same guarded test the save path uses, so the repair agrees with
    # it: a key match that the records contradict is a different school,
    # and copying it across is right, not a duplicate.
    master_map = master_key_map(master)

    cycle_records = []
    for path in all_cycle_paths():
        cycle_records.extend(read_records(path))
    cycle_map = master_key_map(cycle_records)

    missing_from_master = [r for r in cycle_records
                           if found_in_master(r, master_map) is None]
    for record in missing_from_master:
        master.append(record)
        for key in match_keys(record):
            master_map.setdefault(key, record)
    if missing_from_master:
        write_records(master_path(), master)

    missing_from_cycles = [r for r in master
                           if found_in_master(r, cycle_map) is None]
    if missing_from_cycles:
        active = cycle_path(0)
        rows = read_records(active) + missing_from_cycles
        write_records(active, rows)

    _clear_journal()
    report = {"added_to_master": len(missing_from_master),
              "added_to_current": len(missing_from_cycles),
              "journal": had_journal}
    if missing_from_master or missing_from_cycles:
        LOG.warning("Repaired the Excel files after an interrupted save: "
                    "%d record(s) added to the master, %d to the current "
                    "cycle. Nothing was removed.",
                    len(missing_from_master), len(missing_from_cycles))
    return report


def append_records(records):
    """Save finished schools into the master and the current cycle file.

    A school already in the master is not written again - that is the
    duplicate check the whole run depends on. The master is written first,
    then the current cycle file, with a journal around the pair so an
    interrupted save can be repaired on the next start.

    Returns ``(written, skipped_as_duplicate)``.
    """
    require_openpyxl()
    if not records:
        return 0, 0

    keymap = master_key_map()
    # The cycle file is checked as well as the master, not instead of it.
    # A row can be sitting in the active cycle that the master has not got
    # - an interrupted save, a file edited by hand - and checking only the
    # master would let its twin through and write the same school into the
    # cycle file twice. Checked here, before either file is written, so
    # the two never disagree about what was saved.
    #
    # Only the active file: a finished cycle is never appended to again,
    # and a row that drifted out of the master is copied back into it by
    # reconcile_master_and_cycles() at the start of every run.
    for record in read_records(cycle_path(0)):
        for key in match_keys(record):
            keymap.setdefault(key, record)
    fresh = []
    skipped = 0
    for record in records:
        if found_in_master(record, keymap) is not None:
            skipped += 1
            continue
        fresh.append(record)
        # so two new records in the same batch cannot both go in
        for key in match_keys(record):
            keymap.setdefault(key, record)

    if not fresh:
        LOG.info("   nothing new to save: %d school(s) are already in %s",
                 skipped, os.path.basename(master_path()))
        return 0, skipped

    _write_journal(fresh)
    try:
        write_records(master_path(), read_records(master_path()) + fresh)

        remaining = list(fresh)
        while remaining:
            active = cycle_path(0)
            here = read_records(active)
            room = max(CYCLE_SIZE - len(here), 0)
            if room == 0:
                finalise_current_cycle()
                continue
            take, remaining = remaining[:room], remaining[room:]
            write_records(active, here + take)
            if row_count(active) >= CYCLE_SIZE:
                finalise_current_cycle()
    finally:
        _clear_journal()

    # verification: read it back from disk before anything is marked done
    saved = row_count(master_path())
    LOG.info("   saved %d new school(s); %d already known", len(fresh),
             skipped)
    LOG.info("   %-26s %d record(s)", os.path.basename(master_path()), saved)
    for path in all_cycle_paths():
        LOG.info("   %-26s %d record(s)", os.path.basename(path),
                 row_count(path))
    return len(fresh), skipped


# ===========================================================================
#  TEMPORARY PER-CITY STORAGE
# ===========================================================================
#  Discovered schools are written here the moment a city's searches finish,
#  before any Gemini work starts. A crash during enrichment therefore costs
#  the research, never the discovery. These files are removed only after the
#  city's rows have been read back from the final workbook.
# ---------------------------------------------------------------------------

def city_dir(state, city):
    """Temp_Data itself. Every city's files live in it side by side.

    A folder per city would be 866 folders, so the state and city go into
    the file name instead: Sikkim_Rangpo_private_schools.txt.
    """
    return TEMP_DIR


def city_prefix(state, city):
    """The start of every file name belonging to one city."""
    tidy = lambda text: re.sub(r"[^A-Za-z0-9]+", "_",
                               clean(text)).strip("_") or "unknown"
    return "%s_%s" % (tidy(state), tidy(city))


def city_file(state, city, name):
    return os.path.join(TEMP_DIR, "%s_%s" % (city_prefix(state, city), name))


def city_files(state, city):
    """Every file in Temp_Data that belongs to this city."""
    start = city_prefix(state, city) + "_"
    try:
        names = os.listdir(TEMP_DIR)
    except OSError:
        return []
    return [os.path.join(TEMP_DIR, name) for name in sorted(names)
            if name.startswith(start)]


def temp_excel_path(state, city):
    return city_file(state, city, "temporary_schools.xlsx")


def save_temp_excel(state, city, records):
    """The city's discovered schools, in the 57-column schema.

    Written after discovery and again after every school Gemini answers, so
    the file on disk always shows how far the city has got. It is working
    storage only: nothing goes from here to the master or a cycle file until
    the city's research is done.
    """
    if not OPENPYXL_AVAILABLE:
        return ""
    path = temp_excel_path(state, city)
    os.makedirs(TEMP_DIR, exist_ok=True)
    write_records(path, records)
    return path


def load_temp_excel(state, city):
    return read_records(temp_excel_path(state, city))


def save_city_progress(state, city, records):
    os.makedirs(TEMP_DIR, exist_ok=True)
    path = city_file(state, city, "schools.json")
    temporary = path + ".tmp"
    with open(temporary, "w", encoding="utf-8") as handle:
        json.dump(records, handle, ensure_ascii=False, indent=1)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temporary, path)
    return path


def load_city_progress(state, city):
    path = city_file(state, city, "schools.json")
    if not os.path.exists(path):
        return []
    try:
        with open(path, encoding="utf-8") as handle:
            return json.load(handle)
    except (OSError, ValueError):
        LOG.warning("The saved progress for %s could not be read - its "
                    "searches will run again", city)
        return []


# ===========================================================================
#  CLEANING UP TEMPORARY FILES
# ===========================================================================
#  One rule, in three stages:
#
#      SUCCESS            -> CLEAN UP
#      UNRESOLVED FAILURE -> PRESERVE FOR RESUME
#
#  Nothing below is ever called from an error path, and nothing below is in
#  a finally: block. That is deliberate. A school that failed, a keyword
#  whose list never reached its end and a city with rows still missing from
#  the master all keep every file they have, because the next run needs them
#  to carry on from the right school instead of collecting everything again.
#
#  What gets written, and when it goes:
#
#    <State>_<City>_<keyword>.txt[.gz]   the copied Maps text
#         -> removed when that keyword is complete (stage 2)
#    <State>_<City>_gemini_<school>.txt  one school's raw answer
#         -> removed when that school is saved and verified (stage 1)
#    <State>_<City>_temporary_schools.xlsx   the city's working sheet
#         -> removed when the city is complete (stage 3)
#    <State>_<City>_schools.json         the resume file
#         -> removed when the city is complete (stage 3)
# ---------------------------------------------------------------------------

def drop_files(paths):
    """Remove files that are certainly not needed any more."""
    removed = []
    for path in paths:
        try:
            if os.path.exists(path):
                os.remove(path)
                removed.append(os.path.basename(path))
        except OSError as error:
            LOG.warning("   could not remove %s: %s",
                        os.path.basename(path), error)
    return removed


def school_temp_files(state, city, record):
    """Everything written for one school alone."""
    base = gemini_school_path(state, city, record)
    return [base, base + ".gz"]


def clear_school_temp(state, city, record):
    """Stage 1. Only ever called once the school's row is safely in the
    master and has been read back out of it, or once the school is a
    finished no-contact skip. Its raw answer has done its job.
    """
    return drop_files(school_temp_files(state, city, record))


def sweep_completed_keyword(state, city, keyword, resolved=()):
    """Stage 2. Only ever called after a keyword finished with every one of
    its schools resolved.

    Out goes that keyword's copied Maps text: the schools are in the city's
    working sheet by now and the ones worth keeping are in the master, so
    nothing reads it again. Raw answers left behind by an interrupted
    earlier run go with it, but only for schools the checkpoint confirms
    are resolved - a school still waiting for research keeps everything.

    Never reached from an error path. An unresolved keyword keeps its text.
    """
    gone = []
    if not KEEP_CAPTURED_TEXT:
        base = keyword_txt_path(state, city, keyword)
        gone += drop_files([base, base + ".gz"])
    for record in resolved:
        gone += drop_files(school_temp_files(state, city, record))
    return gone


def clear_city_temp(state, city):
    """Stage 3. The city is complete: every row is saved, a finished
    no-contact skip, or already in the master - and process_city has read
    them back out of the final workbook before calling this.

    So the working sheet, the resume file, the copied text and any raw
    answer still lying about all go, and Temp_Data is left empty. With
    KEEP_CAPTURED_TEXT on, the copied text is packed and kept instead.
    An unresolved city never reaches this function.
    """
    mine = city_files(state, city)
    if not mine:
        return 0
    if KEEP_CAPTURED_TEXT:
        pack_city_text(state, city)
        mine = city_files(state, city)
    keep = []
    if KEEP_CAPTURED_TEXT:
        keep = [path for path in mine
                if path.lower().endswith((".txt", ".gz"))
                and "_gemini_" not in os.path.basename(path).lower()]
    removed = drop_files([path for path in mine if path not in keep])
    LOG.info("   %d working file(s) removed, %d captured file(s) kept",
             len(removed), len(keep))
    return len(removed)


def city_of_checkpoint_key(key):
    """"Delhi NCR|Delhi" -> ("Delhi NCR", "Delhi")."""
    state, _, city = str(key).partition("|")
    return state, city


def city_temp_is_finished(state, city):
    """True when nothing in this city's temporary files is still needed.

    The working sheet is read and every row in it must be accounted for:
    either already in the master, or a school with no phone and no email,
    which is a finished skip. One row that is neither means an interrupted
    run still has work to do here, and the files stay.
    """
    rows = load_temp_excel(state, city)
    if not rows:
        return True, 0
    missing = missing_from_master(rows)
    return (not missing), len(missing)


def sweep_completed_cities(checkpoint):
    """Startup. Leftovers from work that is already finished AND verified.

    A city's files are removed only when the checkpoint says the city is
    complete and its working sheet proves it: every row read back out of
    the master, or left out for having no contact details. A city that is
    still open, or that fails that check, is not touched at all - those
    files are exactly what the resume needs.
    """
    try:
        names = os.listdir(TEMP_DIR)
    except OSError:
        return {"cities": 0, "files": 0, "kept": []}
    if not names:
        return {"cities": 0, "files": 0, "kept": []}

    swept, files, kept = 0, 0, []
    for key in list(checkpoint.data.get("completed_cities", [])):
        state, city = city_of_checkpoint_key(key)
        if not city or not city_files(state, city):
            continue
        finished, left = city_temp_is_finished(state, city)
        if not finished:
            kept.append("%s (%d school(s) not in the master)" % (city, left))
            LOG.warning("   %s is marked complete but %d of its school(s) "
                        "are not in the master - its files are kept",
                        city, left)
            continue
        removed = clear_city_temp(state, city)
        if removed:
            swept += 1
            files += removed
    if files:
        LOG.info("   startup cleanup: %d leftover file(s) from %d finished "
                 "city(ies) removed", files, swept)
    return {"cities": swept, "files": files, "kept": kept}


def dump_text(state, city, name, text):
    if not text:
        return ""
    os.makedirs(TEMP_DIR, exist_ok=True)
    path = city_file(state, city, name)
    with open(path, "a", encoding="utf-8") as handle:
        handle.write(text if text.endswith("\n") else text + "\n")
    return path


# ===========================================================================
#  CHECKPOINT
# ===========================================================================

CHECKPOINT_VERSION = 3


def empty_checkpoint():
    return {
        "version": CHECKPOINT_VERSION,
        "rank": None,
        "num_sys": None,
        "plan_fingerprint": "",
        "city_fingerprint": "",
        "current_city": None,          # [state, city]
        "current_keyword": None,
        "discovery_done": {},          # "State|City" -> True
        "keywords_done": {},           # "State|City" -> [keyword, ...]
        "committed": {},               # "State|City" -> [identity, ...]
        "enriched": {},                # "State|City" -> [match key strings]
        "completed_cities": [],        # ["State|City", ...]
        "failed": [],                  # [{record, reason, when}]
        "duplicates": 0,
        "saved": 0,
        "cycle": 0,
        #: how the last GitHub sync went, so a recovery can see what was
        #: really pushed and what was only ever local
        "git": {"branch": "", "pending": False, "last_ok": "",
                "last_error": ""},
        #: the work/break clock, so an interrupted break resumes sanely
        "work": {"period_hours": 0.0, "breaks_taken": 0, "break_until": "",
                 "reason": ""},
        "updated": "",
    }


class Checkpoint:
    """Progress on disk. Written only after the data it describes is saved."""

    def __init__(self, path=None):
        self.path = path or CHECKPOINT_FILE
        self.data = empty_checkpoint()
        self.load()

    # ------------------------------------------------------------ storage
    def load(self):
        if not os.path.exists(self.path):
            return
        try:
            with open(self.path, encoding="utf-8") as handle:
                stored = json.load(handle)
        except (OSError, ValueError):
            LOG.warning("The checkpoint could not be read - starting a fresh "
                        "one. Nothing already saved to Excel is affected.")
            return
        if stored.get("version") != CHECKPOINT_VERSION:
            LOG.warning("The checkpoint was written by an older version and "
                        "has been set aside.")
            return
        merged = empty_checkpoint()
        merged.update(stored)
        self.data = merged

    def save(self):
        os.makedirs(os.path.dirname(self.path) or ".", exist_ok=True)
        self.data["updated"] = datetime.datetime.now().isoformat(timespec="seconds")
        temporary = self.path + ".tmp"
        with open(temporary, "w", encoding="utf-8") as handle:
            json.dump(self.data, handle, ensure_ascii=False, indent=2)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, self.path)

    # ------------------------------------------------------------- plan
    def check_plan(self, value_rank, value_num, fingerprint):
        """Report a changed plan rather than resuming against the wrong one.

        Returns a message when the saved progress was made under a different
        rank, laptop count or city/keyword list; "" when it still matches.
        """
        stored_rank = self.data.get("rank")
        stored_num = self.data.get("num_sys")
        stored_print = self.data.get("plan_fingerprint") or ""
        if stored_rank is None and not stored_print:
            return ""
        if (stored_rank, stored_num) != (value_rank, value_num):
            return ("the saved progress is for rank %s of %s, this run is "
                    "rank %s of %s" % (stored_rank, stored_num, value_rank,
                                       value_num))
        if stored_print and stored_print != fingerprint:
            if self.keywords_only_changed():
                return ""           # handled by migrate_keyword_plan()
            return ("the city list or the enabled keywords changed since the "
                    "saved progress was written")
        return ""

    def keywords_only_changed(self):
        """True when the CITIES still match and only the keywords moved.

        Changing the enabled keywords invalidates nothing already
        collected: a school saved under one keyword is the same school.
        So this case is migrated instead of refused.

        An older checkpoint has no city fingerprint to compare. It is
        accepted as a keyword change only when no city has been completed
        yet - then there is nothing a wrong guess could cost.
        """
        stored_city = self.data.get("city_fingerprint") or ""
        if stored_city:
            return stored_city == city_fingerprint()
        return not self.data.get("completed_cities")

    def migrate_keyword_plan(self, value_rank, value_num, fingerprint):
        """Adopt a new keyword list, keeping every saved record.

        Nothing is cleared: completed cities stay completed, the saved and
        duplicate counts stay, and the finished-keyword lists stay as they
        are - a name no longer in the plan is simply never asked for again.
        """
        self.data["rank"] = value_rank
        self.data["num_sys"] = value_num
        self.data["plan_fingerprint"] = fingerprint
        self.data["city_fingerprint"] = city_fingerprint()
        self.save()

    def adopt_plan(self, value_rank, value_num, fingerprint):
        self.data["rank"] = value_rank
        self.data["num_sys"] = value_num
        self.data["plan_fingerprint"] = fingerprint
        self.data["city_fingerprint"] = city_fingerprint()

    # -------------------------------------------------- changing num_sys
    def cities_with_progress(self):
        """Every city this checkpoint knows anything about.

        Splitting the work across more laptops moves cities between ranks.
        A city this rank has touched must still belong to this rank
        afterwards, or its saved progress would be stranded in a copy that
        no longer searches it. These are the places progress can hide.
        """
        keys = set()
        for field in ("committed", "enriched", "discovery_done",
                      "keywords_done"):
            keys.update(self.data.get(field) or {})
        keys.update(self.data.get("completed_cities") or [])
        current = self.data.get("current_city") or []
        if len(current) == 2:
            keys.add(self.key(current[0], current[1]))
        for item in self.data.get("failed") or []:
            if isinstance(item, dict):
                state = item.get("State") or item.get("state")
                city = item.get("City") or item.get("city")
                if state and city:
                    keys.add(self.key(state, city))
        return sorted(k for k in keys if k)

    def plan_migration(self, value_rank, value_num):
        """Can this checkpoint move to ``value_rank`` of ``value_num``?

        Returns ``(ok, lines)``. It is safe only when every city this
        checkpoint has touched is still assigned to ``value_rank`` under
        the new split - then the rank keeps exactly the work it already
        did. If any touched city moves to a different rank, migrating
        would strand that city's records, so this refuses and names the
        city and the rank it moved to. Nothing is written either way.
        """
        mine = set(get_cities_by_rank(ALL_CITIES, value_rank, value_num))
        lines, stranded = [], []
        for key in self.cities_with_progress():
            # the key is "State|City"; the split returns both, and it is the
            # city name that the rank slices are made of
            _state, city = city_of_checkpoint_key(key)
            if city in mine:
                lines.append("   %s - stays with rank %d" % (key, value_rank))
                continue
            owner = None
            for other in range(1, value_num + 1):
                if city in get_cities_by_rank(ALL_CITIES, other, value_num):
                    owner = other
                    break
            stranded.append((key, owner))
            lines.append("   %s - MOVES to rank %s" % (key, owner or "none"))
        return not stranded, lines

    def migrate_plan(self, value_rank, value_num, fingerprint):
        """Adopt the new rank/laptop count, keeping every saved record.

        Only the plan fields change: the committed and enriched keys, the
        completed cities, the counters and the retry queue are untouched.
        Call it only after :meth:`plan_migration` has approved the move.
        """
        self.adopt_plan(value_rank, value_num, fingerprint)
        self.save()

    # ------------------------------------------------------------ cities
    @staticmethod
    def key(state, city):
        return "%s|%s" % (state, city)

    def is_city_complete(self, state, city):
        return self.key(state, city) in self.data["completed_cities"]

    def mark_discovery_done(self, state, city):
        self.data["discovery_done"][self.key(state, city)] = True
        self.save()

    def discovery_done(self, state, city):
        return bool(self.data["discovery_done"].get(self.key(state, city)))

    def keyword_done(self, state, city, keyword):
        return keyword in self.data["keywords_done"].get(self.key(state, city),
                                                         [])

    def mark_keyword_done(self, state, city, keyword):
        done = self.data["keywords_done"].setdefault(self.key(state, city), [])
        if keyword not in done:
            done.append(keyword)
        self.save()

    def note_committed(self, state, city, identity):
        """A record that has reached the master and a cycle file."""
        done = self.data["committed"].setdefault(self.key(state, city), [])
        if identity not in done:
            done.append(identity)
        self.data["cycle"] = len(finished_cycle_numbers())
        self.save()

    def committed_keys(self, state, city):
        return set(self.data["committed"].get(self.key(state, city), []))

    def enriched_keys(self, state, city):
        return set(self.data["enriched"].get(self.key(state, city), []))

    def mark_enriched(self, state, city, identity):
        done = self.data["enriched"].setdefault(self.key(state, city), [])
        if identity not in done:
            done.append(identity)
        self.save()

    def add_totals(self, saved, duplicates):
        """Add one visit's counts to the running totals.

        Needed because a city can now be visited more than once - one
        keyword per pass - and only the visit that CLOSES the city reaches
        mark_city_complete(). Without this the earlier passes' saved
        schools were in the Excel files but missing from the totals that
        --status reports.
        """
        self.data["saved"] += int(saved)
        self.data["duplicates"] += int(duplicates)
        self.save()

    def mark_city_complete(self, state, city, saved, duplicates):
        key = self.key(state, city)
        if key not in self.data["completed_cities"]:
            self.data["completed_cities"].append(key)
        self.data["discovery_done"].pop(key, None)
        self.data["enriched"].pop(key, None)
        self.data["keywords_done"].pop(key, None)
        self.data["committed"].pop(key, None)
        self.data["saved"] += int(saved)
        self.data["duplicates"] += int(duplicates)
        self.data["current_city"] = None
        self.data["current_keyword"] = None
        self.save()

    def start_city(self, state, city):
        self.data["current_city"] = [state, city]
        self.save()

    def note_keyword(self, keyword):
        self.data["current_keyword"] = keyword
        self.save()

    # ------------------------------------------------------------ failures
    def record_failure(self, record, reason):
        self.data["failed"].append({
            "school": clean(record.get("School Name", "")),
            "city": clean(record.get("City", "")),
            "state": clean(record.get("State", "")),
            "reason": str(reason)[:300],
            "when": datetime.datetime.now().isoformat(timespec="seconds"),
            "record": record,
        })
        self.save()

    def take_failures(self):
        failures = list(self.data["failed"])
        self.data["failed"] = []
        self.save()
        return failures

    def forget_city(self, state, city):
        key = self.key(state, city)
        self.data["completed_cities"] = [k for k in self.data["completed_cities"]
                                         if k != key]
        self.data["discovery_done"].pop(key, None)
        self.data["enriched"].pop(key, None)
        self.data["keywords_done"].pop(key, None)
        self.data["committed"].pop(key, None)
        self.save()

# ===========================================================================
#  BROWSER
# ===========================================================================

try:
    from playwright.sync_api import sync_playwright
    from playwright.sync_api import Error as PlaywrightError
    from playwright.sync_api import TimeoutError as PlaywrightTimeout
    PLAYWRIGHT_AVAILABLE = True
except ImportError:                                     # pragma: no cover
    sync_playwright = None

    class PlaywrightError(Exception):
        pass

    class PlaywrightTimeout(PlaywrightError):
        pass

    PLAYWRIGHT_AVAILABLE = False


class BrowserUnavailable(RuntimeError):
    pass


class AccessBlocked(RuntimeError):
    """A sign-in wall, a consent wall or a bot challenge. Reported, not solved."""


class GeminiUnavailable(RuntimeError):
    pass


def gemini_input_error(message):
    return GeminiUnavailable(message)


class WrongSchoolAnswer(RuntimeError):
    """The answer is about a different school than the one that was asked."""


BLOCK_SIGNS = (
    "unusual traffic", "are you a robot", "recaptcha", "captcha",
    "verify you are human", "detected unusual activity",
    "our systems have detected",
)
SIGN_IN_SIGNS = ("sign in to continue", "sign in to gemini",
                 "to continue to gemini", "choose an account")


def looks_blocked(text):
    low = (text or "").lower()
    return any(sign in low for sign in BLOCK_SIGNS)


class DesktopUnavailable(BrowserUnavailable):
    """--show was asked for but the real desktop display cannot be reached.

    A BrowserUnavailable, so a run that cannot show the window stops with the
    explanation rather than a traceback - and never falls back to the
    invisible screen without saying so.
    """


# --- finding and checking the desktop screen -------------------------------
#: Set these only if the detection below picks the wrong screen;
#: SCHOOL_DESKTOP_DISPLAY / SCHOOL_DESKTOP_XAUTHORITY do the same.
DESKTOP_DISPLAY = None              # e.g. ":0"
DESKTOP_XAUTHORITY = None
SHOW_WINDOW_WAIT_SECONDS = 25       # for the window to appear
WINDOW_CHECK_POLL_SECONDS = 1.0     # between two looks at the screen
DISPLAY_CHECK_TIMEOUT = 10          # for xdpyinfo / xset / xdotool
WINDOW_LIST_TIMEOUT = 15            # for xwininfo
SOCKET_PROBE_TIMEOUT = 3            # for the bare X socket check


# ===========================================================================
#  SHOWING THE BROWSER ON THE REAL DESKTOP  (--show)
# ===========================================================================
#  xvfb-run runs the whole command against a private, invisible X server: it
#  replaces DISPLAY (with something like :99) AND XAUTHORITY (with a file in
#  /tmp/xvfb-run.XXXX/). A browser started inside it draws into that virtual
#  screen, which is exactly what you want for an unattended run and exactly
#  what you do NOT want when you are watching.
#
#  --show finds the display the desktop is really on, together with the
#  authority file that opens it, and starts Chromium against that instead.
#  Both are needed: the display alone gets "Authorization required, but no
#  authorization protocol specified" and the browser dies.
#
#  Nothing else changes. The rank, the city list, Maps, Gemini, Excel and the
#  checkpoint neither know nor care which screen the window is on.
# ---------------------------------------------------------------------------

X_SOCKET_DIR = "/tmp/.X11-unix"


def running_under_xvfb():
    """True when this process was started by xvfb-run."""
    return "xvfb-run" in (os.environ.get("XAUTHORITY") or "").lower()


def _desktop_from_processes():
    """Candidates read from the environment of the session's own processes.

    The most reliable source there is: the desktop's own programs still hold
    the DISPLAY and XAUTHORITY they were started with, and xvfb-run cannot
    reach into them. Only this user's processes are read.
    """
    here = os.environ.get("DISPLAY", "")
    uid = os.getuid()
    found = {}
    try:
        entries = os.listdir("/proc")
    except OSError:
        return []
    for pid in entries:
        if not pid.isdigit():
            continue
        try:
            if os.stat("/proc/" + pid).st_uid != uid:
                continue
            with open("/proc/%s/environ" % pid, "rb") as handle:
                raw = handle.read()
        except OSError:
            continue
        env = {}
        for piece in raw.split(b"\0"):
            if b"=" in piece:
                key, _sep, value = piece.partition(b"=")
                try:
                    env[key.decode()] = value.decode()
                except UnicodeDecodeError:
                    pass
        display = env.get("DISPLAY", "")
        authority = env.get("XAUTHORITY", "")
        if not display or display == here:
            continue
        if "xvfb" in authority.lower() or "xvfb" in display.lower():
            continue
        found.setdefault((display, authority), "/proc/%s" % pid)
    return [(d, a, "another session process") for d, a in found]


def _desktop_from_files():
    """Candidates built from the X sockets and the usual authority files."""
    here = os.environ.get("DISPLAY", "")
    uid = os.getuid()
    displays = []
    try:
        names = os.listdir(X_SOCKET_DIR)
    except OSError:
        names = []
    for name in sorted(names):
        if not name.startswith("X") or not name[1:].isdigit():
            continue
        path = os.path.join(X_SOCKET_DIR, name)
        try:
            if os.stat(path).st_uid != uid:      # gdm's own servers, not ours
                continue
        except OSError:
            continue
        display = ":" + name[1:]
        if display != here:
            displays.append((int(name[1:]), display))

    authorities = []
    run_dir = "/run/user/%d" % uid
    try:
        for name in sorted(os.listdir(run_dir)):
            if "xwaylandauth" in name.lower() or name == ".Xauthority":
                authorities.append(os.path.join(run_dir, name))
    except OSError:
        pass
    for extra in (os.path.join(run_dir, "gdm", "Xauthority"),
                  os.path.expanduser("~/.Xauthority")):
        if os.path.exists(extra):
            authorities.append(extra)
    authorities.append("")                        # no authority file at all

    out = []
    for _number, display in sorted(displays):
        for authority in authorities:
            out.append((display, authority, "an X socket"))
    return out


def _display_env(display, authority):
    env = dict(os.environ)
    env["DISPLAY"] = display
    if authority:
        env["XAUTHORITY"] = authority
    else:
        env.pop("XAUTHORITY", None)
    # a Wayland socket in the environment makes Chromium prefer Wayland and
    # ignore DISPLAY entirely, which is the whole point of this function
    env.pop("WAYLAND_DISPLAY", None)
    return env


def display_works(display, authority):
    """``(ok, how_it_was_checked)`` for one display and authority file."""
    import subprocess
    env = _display_env(display, authority)
    for command in (["xdpyinfo", "-display", display],
                    ["xset", "-display", display, "-q"]):
        if not shutil.which(command[0]):
            continue
        try:
            done = subprocess.run(command, env=env,
                                  timeout=DISPLAY_CHECK_TIMEOUT,
                                  stdout=subprocess.DEVNULL,
                                  stderr=subprocess.DEVNULL)
        except (OSError, subprocess.SubprocessError):
            continue
        return done.returncode == 0, command[0]
    # neither tool is installed: a socket that accepts a connection is the
    # most that can be checked here, and the caller is told so
    import socket
    path = os.path.join(X_SOCKET_DIR, "X" + display.lstrip(":").split(".")[0])
    try:
        probe = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
        probe.settimeout(SOCKET_PROBE_TIMEOUT)
        probe.connect(path)
        probe.close()
        return True, "socket only (install x11-utils to check properly)"
    except OSError:
        return False, "socket only (install x11-utils to check properly)"


def find_desktop_display():
    """``(display, authority, how)`` for the real desktop, or raise.

    Order: what the operator set, then the environment variables, then the
    session's own processes, then the X sockets. Every candidate is actually
    opened before it is returned, so --show never silently ends up drawing
    into the invisible screen.
    """
    tried = []

    forced_display = DESKTOP_DISPLAY or os.environ.get("SCHOOL_DESKTOP_DISPLAY")
    if forced_display:
        authority = (DESKTOP_XAUTHORITY
                     or os.environ.get("SCHOOL_DESKTOP_XAUTHORITY") or "")
        ok, how = display_works(forced_display, authority)
        if ok:
            return forced_display, authority, "you set it (%s)" % how
        raise DesktopUnavailable(
            "DISPLAY %s was set by hand but could not be opened (%s). Check "
            "DESKTOP_DISPLAY / DESKTOP_XAUTHORITY at the top of this file."
            % (forced_display, how))

    candidates = []
    # not inside xvfb-run: the screen this process already has IS the desktop
    if not running_under_xvfb() and os.environ.get("DISPLAY"):
        candidates.append((os.environ["DISPLAY"],
                           os.environ.get("XAUTHORITY", ""),
                           "the DISPLAY this run already has"))
    candidates += _desktop_from_processes() + _desktop_from_files()
    seen = set()
    for display, authority, source in candidates:
        if (display, authority) in seen:
            continue
        seen.add((display, authority))
        ok, how = display_works(display, authority)
        tried.append("%s with %s -> %s"
                     % (display, authority or "no authority file",
                        "opened" if ok else "refused"))
        if ok:
            return display, authority, "%s, checked with %s" % (source, how)

    raise DesktopUnavailable(
        "--show cannot find your desktop screen.\n"
        "This run's DISPLAY is %s%s.\n"
        "Tried:\n    %s\n"
        "Fix it by setting the right one yourself:\n"
        "    SCHOOL_DESKTOP_DISPLAY=:0 \\\n"
        "    SCHOOL_DESKTOP_XAUTHORITY=$XAUTHORITY \\\n"
        "    xvfb-run -a python3 find_school.py --show\n"
        "Run `echo $DISPLAY $XAUTHORITY` in a normal desktop terminal to see "
        "what those two values should be."
        % (os.environ.get("DISPLAY") or "unset",
           " (inside xvfb-run)" if running_under_xvfb() else "",
           "\n    ".join(tried) or "nothing - no X socket belongs to you"))


_WINDOW_TITLE = re.compile(r'^\s*0x[0-9a-f]+\s+"([^"]*)"')


def windows_on_display(display, authority):
    """The window titles the given display is showing. [] when none.

    ``xwininfo -root -tree`` walks the whole window tree, which is what is
    needed: a browser window is a child of the window manager's frame, not a
    direct child of the root.
    """
    import subprocess
    env = _display_env(display, authority)
    if shutil.which("xwininfo"):
        try:
            done = subprocess.run(
                ["xwininfo", "-root", "-tree", "-display", display],
                env=env, timeout=WINDOW_LIST_TIMEOUT,
                capture_output=True, text=True)
            if done.returncode == 0:
                titles = []
                for line in done.stdout.splitlines():
                    found = _WINDOW_TITLE.match(line)
                    if found and found.group(1).strip():
                        titles.append(found.group(1).strip())
                return titles
        except (OSError, subprocess.SubprocessError):
            pass
    if shutil.which("xdotool"):
        try:
            done = subprocess.run(
                ["xdotool", "search", "--onlyvisible", "--name", "."],
                env=env, timeout=DISPLAY_CHECK_TIMEOUT,
                capture_output=True, text=True)
            if done.returncode == 0:
                titles = []
                for window in done.stdout.split():
                    name = subprocess.run(
                        ["xdotool", "getwindowname", window], env=env,
                        timeout=SOCKET_PROBE_TIMEOUT,
                        capture_output=True, text=True)
                    if name.returncode == 0 and name.stdout.strip():
                        titles.append(name.stdout.strip())
                return titles
        except (OSError, subprocess.SubprocessError):
            pass
    return []


_WINDOW_LINE = re.compile(
    r'^\s*(0x[0-9a-f]+)\s+"([^"]*)":\s+\("([^"]*)"')


def desktop_windows(display, authority):
    """``[{id, title, profile}, ...]`` for the windows on that screen.

    Chromium puts its user-data-dir in the window's class, which is the only
    thing that tells two of this script's browsers apart - both are called
    "Google Chrome for Testing" and both may be showing a Google page.
    """
    import subprocess
    if not shutil.which("xwininfo"):
        return []
    env = _display_env(display, authority)
    try:
        done = subprocess.run(
            ["xwininfo", "-root", "-tree", "-display", display], env=env,
            timeout=WINDOW_LIST_TIMEOUT, capture_output=True, text=True)
    except (OSError, subprocess.SubprocessError):
        return []
    if done.returncode != 0:
        return []
    windows = []
    for line in done.stdout.splitlines():
        found = _WINDOW_LINE.match(line)
        if not found:
            continue
        window_id, title, klass = found.groups()
        profile = ""
        inside = re.search(r"\(([^()]*)\)\s*$", klass)
        if inside:
            profile = inside.group(1).strip()
        windows.append({"id": window_id, "title": title, "profile": profile})
    return windows


def window_of_profile(display, authority, profile_dir):
    """The window belonging to the browser using that profile, or ""."""
    wanted = os.path.abspath(profile_dir)
    for window in desktop_windows(display, authority):
        if window["profile"] and os.path.abspath(window["profile"]) == wanted:
            return window
    return None


def _xdotool(display, authority, *arguments):
    import subprocess
    if not shutil.which("xdotool"):
        return False
    env = _display_env(display, authority)
    try:
        done = subprocess.run(["xdotool"] + list(arguments), env=env,
                              timeout=DISPLAY_CHECK_TIMEOUT,
                              capture_output=True, text=True)
    except (OSError, subprocess.SubprocessError):
        return False
    return done.returncode == 0


def active_window(display, authority):
    import subprocess
    if not shutil.which("xdotool"):
        return ""
    env = _display_env(display, authority)
    try:
        done = subprocess.run(["xdotool", "getactivewindow"], env=env,
                              timeout=DISPLAY_CHECK_TIMEOUT,
                              capture_output=True, text=True)
    except (OSError, subprocess.SubprocessError):
        return ""
    if done.returncode != 0:
        return ""
    try:
        return hex(int(done.stdout.strip()))
    except ValueError:
        return ""


def bring_window_forward(display, authority, window_id, others=()):
    """Raise one window and send the others behind it.

    ``(True, detail)`` only when the screen really changed: the raise is
    checked against the window manager's own idea of the active window, so a
    failed switch is reported instead of leaving the operator watching the
    wrong browser.
    """
    for other in others:
        if other and other != window_id:
            _xdotool(display, authority, "windowminimize", other)
    ok = _xdotool(display, authority, "windowactivate", "--sync", window_id)
    if not ok:
        _xdotool(display, authority, "windowmap", window_id)
        ok = _xdotool(display, authority, "windowraise", window_id)
    time.sleep(WINDOW_CHECK_POLL_SECONDS)
    active = active_window(display, authority)
    if active and int(active, 16) == int(window_id, 16):
        return True, "it is the active window"
    still_there = any(w["id"] == window_id
                      for w in desktop_windows(display, authority))
    if ok and still_there:
        return True, "raised (the window manager reports %s active)" % (
            active or "nothing")
    return False, "the window could not be brought forward"


def confirm_window_is_visible(display, authority, wait=None):
    """``(True, title)`` once a Chromium window is really on that screen.

    Setting headless=False is not proof of anything: the window can be drawn
    into the virtual screen, or Chromium can fail to reach the display and
    die. This looks at the screen itself.
    """
    deadline = time.time() + float(wait if wait is not None
                                   else SHOW_WINDOW_WAIT_SECONDS)
    last = []
    while time.time() < deadline:
        last = windows_on_display(display, authority)
        for title in last:
            low = title.lower()
            if "clipboard" in low:
                continue            # Chromium's 10x10 helper, not the window
            if ("chrom" in low or "chrome for testing" in low
                    or "playwright" in low or "new tab" in low
                    or "about:blank" in low):
                return True, title
        time.sleep(WINDOW_CHECK_POLL_SECONDS)
    return False, ("; ".join(last[:6]) if last
                   else "no window at all on %s" % display)


# --- the browser -----------------------------------------------------------
BROWSER_MODE = "headful"            # "headful" | "headless"
BROWSER_LOCALE = "en-IN"
BROWSER_TIMEOUT_MS = 45000          # Playwright's own limit, milliseconds
BROWSER_CACHE_MB = 40               # page cache each browser may keep
CLIPBOARD_ENABLED = True            # Ctrl+C needs permission to read it back
MAX_BROWSER_RESTARTS = 3            # before a run stops instead of retrying
PAGE_SETTLE_SECONDS = 3.0           # after a page loads, before reading it
CLICK_SETTLE_SECONDS = 0.2          # after clicking into a box
TYPING_SETTLE_SECONDS = 0.4         # after text is put into a box
RETRY_BACKOFF_SECONDS = 5.0         # grows with each attempt
PROFILE_RELEASE_WAIT_SECONDS = 30   # for Chromium to let go of its profile
BROWSER_EXIT_POLL_SECONDS = 0.2     # between two looks for it to have gone
SCRATCH_SWEEP_AGE_SECONDS = 3600    # before a left-behind profile is swept


def forget_the_last_crash(profile_dir):
    """Tell the profile its last session ended normally.

    Chromium decides whether to offer "Restore pages?" from two values in the
    profile's Preferences file. A run killed with Ctrl+C never gets to write
    them, so the next run starts with the bubble over the map. Only those two
    values are touched; the rest of the profile, including any sign-in, is
    left exactly as it is. A missing or unreadable file is left alone - the
    command-line flags still suppress the bubble.
    """
    path = os.path.join(profile_dir, "Default", "Preferences")
    if not os.path.exists(path):
        return False
    try:
        with open(path, encoding="utf-8") as handle:
            prefs = json.load(handle)
    except (OSError, ValueError):
        return False
    profile = prefs.get("profile")
    if not isinstance(profile, dict):
        return False
    if (profile.get("exit_type") == "Normal"
            and profile.get("exited_cleanly") is True):
        return False
    profile["exit_type"] = "Normal"
    profile["exited_cleanly"] = True
    try:
        temporary = path + ".tmp"
        with open(temporary, "w", encoding="utf-8") as handle:
            json.dump(prefs, handle)
        os.replace(temporary, path)
    except OSError:
        return False
    LOG.debug("  cleared the crash flag on %s", os.path.basename(profile_dir))
    return True


class Browser:
    """One Chromium session, reused for the whole run.

    Used as a context manager. A crashed or closed browser is detected and
    restarted up to MAX_BROWSER_RESTARTS times; beyond that the run stops
    rather than retrying forever.
    """

    def __init__(self, mode=None, show=False, profile_dir=None,
                 label="browser", playwright=None, display=None,
                 authority=None):
        self.show = bool(show)
        #: a second visible browser is told which screen to use rather than
        #: working it out again - the first one already did that
        self._told_display = display
        self._told_authority = authority
        self.profile_dir = profile_dir or BROWSER_PROFILE_DIR
        self.label = label
        #: a second browser shares the first one's Playwright driver: one
        #: process may only ever start the sync driver once
        self._shared_play = playwright
        self._owns_play = playwright is None
        self.mode = (mode or BROWSER_MODE).lower()
        if self.show:
            self.mode = "headful"       # a hidden window shows nobody anything
        elif mode is None and not running_under_xvfb() and os.environ.get(
                "DISPLAY"):
            # no --show and no virtual screen: do not put a window on the
            # desktop the operator is using
            self.mode = "headless"
        self.display = ""               # the screen the window is drawn on
        self.authority = ""
        self._kept_env = {}             # DISPLAY etc. as they were before
        self._play = None
        self._context = None
        self._page = None
        self.restarts = 0

    def __enter__(self):
        self.start()
        return self

    def __exit__(self, *_exc):
        self.close()
        return False

    # ------------------------------------------------------------ lifecycle
    def start(self):
        if not PLAYWRIGHT_AVAILABLE:
            raise BrowserUnavailable(
                "Playwright is not installed. Run:\n"
                "    python3 -m pip install playwright\n"
                "    python3 -m playwright install chromium")
        if self._context is not None:
            return
        self._play = self._shared_play or sync_playwright().start()
        options = {
            "headless": self.mode == "headless",
            "locale": BROWSER_LOCALE,
            # Ctrl+C is worthless if the page may not read the clipboard back
            "permissions": (["clipboard-read", "clipboard-write"]
                            if CLIPBOARD_ENABLED else []),
            "timezone_id": "Asia/Kolkata",
            "viewport": {"width": 1440, "height": 900},
            "args": [
                "--disable-blink-features=AutomationControlled",
                # "Restore pages? Chromium didn't shut down correctly."
                # A run that is stopped with Ctrl+C or killed leaves the
                # profile marked as crashed, and the next run opens with that
                # bubble over the page - and a stray click on Restore would
                # reopen the previous run's tabs. Suppressed at both ends:
                # the flags here, and the profile repair below.
                "--hide-crash-restore-bubble",
                "--disable-session-crashed-bubble",
                "--no-first-run",
                "--no-default-browser-check",
                "--disable-infobars",
                "--restore-last-session=false",
                # this script opens each page once, so a large page cache
                # buys nothing and cost 625 MB on the first real run
                "--disk-cache-size=%d" % (BROWSER_CACHE_MB * 1024 * 1024),
                "--media-cache-size=%d" % (BROWSER_CACHE_MB * 1024 * 1024),
                "--disable-breakpad",
                "--disable-crash-reporter",
                "--disable-component-update",
                "--disable-background-networking",
                "--no-pings",
            ],
        }

        where = ("no screen (headless)" if self.mode == "headless"
                 else "this run's DISPLAY (%s)" % (os.environ.get("DISPLAY")
                                                   or "none"))
        if self.show:
            if self._told_display:
                self.display = self._told_display
                self.authority = self._told_authority or ""
                how = "handed over by the first browser"
            else:
                # raises DesktopUnavailable, with what to do about it, rather
                # than quietly drawing into the invisible screen
                self.display, self.authority, how = find_desktop_display()

            # Setting Playwright's env= is NOT enough: the browser process
            # does not inherit it on this platform. The variables have to be
            # in THIS process before the Playwright driver is started, so the
            # whole chain down to Chromium inherits them.
            self._kept_env = {name: os.environ.get(name)
                              for name in ("DISPLAY", "XAUTHORITY",
                                           "WAYLAND_DISPLAY")}
            os.environ.update(_display_env(self.display, self.authority))
            os.environ.pop("WAYLAND_DISPLAY", None)

            # And this is the flag that actually decides it. On a Wayland
            # desktop Chromium finds the compositor socket by itself, ignores
            # DISPLAY, then fails to make a surface ("'--ozone-platform=
            # wayland' is not compatible with Vulkan") and draws no window at
            # all - which looks exactly like --show being ignored.
            options["args"] = list(options["args"]) + [
                "--ozone-platform=x11", "--start-maximized"]
            options["env"] = dict(os.environ)
            where = "your desktop, %s (found by %s)" % (self.display, how)
            LOG.info("--show: putting the window on %s", where)
        elif self.mode != "headless":
            # No --show, and a window IS being made - on xvfb-run's private
            # screen, where nobody sees it. That only holds if Chromium
            # actually uses that screen, and the choice was being left to
            # it: xvfb-run replaces DISPLAY but leaves WAYLAND_DISPLAY
            # alone, so on a Wayland desktop Chromium can find the real
            # compositor by itself, ignore DISPLAY and put a window in
            # front of you. The same two things that make --show work are
            # what keep this one hidden, pointed the other way.
            self._kept_env = {name: os.environ.get(name)
                              for name in ("WAYLAND_DISPLAY",)}
            os.environ.pop("WAYLAND_DISPLAY", None)
            options["args"] = list(options["args"]) + ["--ozone-platform=x11"]
            options["env"] = dict(os.environ)

        try:
            # Always a profile directory this script knows the path of:
            # the clipboard permission, the crash-flag repair and the
            # --show window check all need one. It is made in a temporary
            # folder and removed when the browser closes.
            os.makedirs(self.profile_dir, exist_ok=True)
            forget_the_last_crash(self.profile_dir)
            self._context = self._play.chromium.launch_persistent_context(
                self.profile_dir, **options)
        except Exception as error:                      # noqa: BLE001
            if self._owns_play and self._play is not None:
                self._play.stop()
            self._play = None
            text = str(error)
            if "already in use" in text or "existing browser session" in text:
                raise BrowserUnavailable(
                    "The browser profile %s is still held by a Chromium from "
                    "an earlier run.\nWait a few seconds and start again, or "
                    "close that browser window. Nothing has been collected "
                    "and nothing is lost." % self.profile_dir)
            raise BrowserUnavailable(
                "Chromium could not be started: %s\nRun:  python3 -m "
                "playwright install chromium" % text[:200])
        self._context.set_default_timeout(BROWSER_TIMEOUT_MS)
        LOG.info("%s started (%s) on %s", self.label.capitalize(),
                 self.mode, where)

        if self.show:
            self._prove_the_window_is_visible()

    def _prove_the_window_is_visible(self):
        """Open a page and check the desktop really grew a Chromium window."""
        try:
            # the working tab itself, started on Maps - no extra tab is
            # opened for the check and none is left behind
            page = self.page()
            page.goto("https://www.google.com/maps?hl=en&gl=in",
                      wait_until="domcontentloaded",
                      timeout=BROWSER_TIMEOUT_MS)
        except Exception:                               # noqa: BLE001
            pass
        visible, detail = confirm_window_is_visible(self.display,
                                                    self.authority)
        if visible:
            LOG.info("--show confirmed: %r is on %s", detail[:60],
                     self.display)
            return
        LOG.error("--show could not confirm a browser window on %s "
                  "(%s).", self.display, detail)
        LOG.error("The run is continuing, but you may not be able to see it. "
                  "Set SCHOOL_DESKTOP_DISPLAY and SCHOOL_DESKTOP_XAUTHORITY "
                  "to the values a normal desktop terminal reports for "
                  "$DISPLAY and $XAUTHORITY.")

    def playwright(self):
        """This browser's driver, so a second browser can share it."""
        return self._play

    def close(self):
        try:
            if self._context is not None:
                self._context.close()
        except Exception:                               # noqa: BLE001
            pass
        try:
            # only the browser that started the driver may stop it
            if self._owns_play and self._play is not None:
                self._play.stop()
        except Exception:                               # noqa: BLE001
            pass
        self._context = self._page = self._play = None
        for name, value in getattr(self, "_kept_env", {}).items():
            if value is None:
                os.environ.pop(name, None)
            else:
                os.environ[name] = value
        self._kept_env = {}

    def alive(self):
        try:
            return self._context is not None and bool(self._context.pages)
        except Exception:                               # noqa: BLE001
            return False

    def restart(self):
        """Recover from a crashed browser, within a bounded number of tries."""
        if self.restarts >= MAX_BROWSER_RESTARTS:
            raise BrowserUnavailable(
                "the browser failed %d times - stopping instead of retrying "
                "indefinitely" % self.restarts)
        self.restarts += 1
        LOG.warning("Restarting the browser (attempt %d of %d)",
                    self.restarts, MAX_BROWSER_RESTARTS)
        self.close()
        time.sleep(RETRY_BACKOFF_SECONDS)
        self.start()

    # ---------------------------------------------------------------- pages
    def page(self):
        """The one working tab.

        A persistent context opens with a blank tab of its own. Taking that
        tab instead of adding another is what keeps the window down to a
        single visible tab - otherwise the operator watches an about:blank
        tab sit next to the one actually doing the work.
        """
        if self._page is not None and not self._page.is_closed():
            return self._page
        existing = [page for page in self._context.pages if not page.is_closed()]
        self._page = existing[0] if existing else self._context.new_page()
        self.close_spare_tabs()
        return self._page

    def close_spare_tabs(self):
        """Close every tab except the working one. Returns how many went."""
        closed = 0
        for page in list(self._context.pages):
            if page is self._page or page.is_closed():
                continue
            try:
                page.close()
                closed += 1
            except Exception:                           # noqa: BLE001
                pass
        return closed

    def tab_titles(self):
        titles = []
        for page in self._context.pages:
            if page.is_closed():
                continue
            try:
                titles.append(page.title() or page.url)
            except Exception:                           # noqa: BLE001
                titles.append("?")
        return titles

    def show_tab(self):
        """Put the working tab in front, so Maps is what the operator sees."""
        try:
            if self._page is not None and not self._page.is_closed():
                self._page.bring_to_front()
        except Exception:                               # noqa: BLE001
            pass

    def new_page(self):
        return self._context.new_page()

    def goto(self, page, url, wait="domcontentloaded"):
        try:
            page.goto(url, wait_until=wait, timeout=BROWSER_TIMEOUT_MS)
        except PlaywrightTimeout:
            LOG.debug("  %s was slow to load - continuing with what is there",
                      url[:80])
        return page

    @staticmethod
    def text_of(page, selector="body"):
        try:
            return page.inner_text(selector) or ""
        except Exception:                               # noqa: BLE001
            return ""


# ===========================================================================
#  GOOGLE MAPS SCHOOL DISCOVERY
# ===========================================================================
#  WHY THIS IS NOT Ctrl+A / Down / Ctrl+C
#  --------------------------------------
#  Ctrl+A selects everything the results pane has RENDERED. Maps virtualises
#  that pane: it removes cards that scroll out of view. So one Down keypress
#  may add nothing, may add several cards, and may drop as many from the top
#  as it adds at the bottom. Copying after every keypress therefore produces
#  overlapping, unordered, partly repeated text, and a run that stops as soon
#  as two copies look alike loses most of the list - which is exactly what
#  "six schools" was.
#
#  Instead the results feed is scrolled and the rendered cards are read after
#  each scroll, keyed by their place link. A card already seen is ignored, a
#  card not seen before is added. That is insensitive to how much one scroll
#  happens to move, and nothing is counted twice.
#
#  Stopping is deliberate and the reason is always reported:
#    * the end-of-list message appeared           -> complete
#    * no new school for SCROLL_NO_PROGRESS_PRESSES presses -> INCOMPLETE
#    * the clipboard stopped being readable               -> INCOMPLETE
#    * the list stopped moving                            -> INCOMPLETE
#  A city is never marked complete on an incomplete search.
# ---------------------------------------------------------------------------

FEED_SELECTOR = 'div[role="feed"]'
CARD_SELECTOR = 'div[role="feed"] a[href*="/maps/place/"]'


def maps_search_url(query):
    """https://www.google.com/maps/search/private+schools+in+bengaluru"""
    return (MAPS_SEARCH_URL + urllib.parse.quote_plus(query)
            + "?hl=en&gl=in")


def build_query(keyword, city):
    return " ".join(SEARCH_QUERY_TEMPLATE.format(
        keyword=clean(keyword), city=clean(city)).split())


# --- the Google Maps search ------------------------------------------------
MAPS_SEARCH_URL = "https://www.google.com/maps/search/"
SEARCH_QUERY_TEMPLATE = "{keyword} in {city}"
END_OF_LIST_MARKERS = (
    "you've reached the end of the list",
    "you have reached the end of the list",
    "youve reached the end of the list",
)
COPY_KEY_PAUSE_SECONDS = 0.35       # between Ctrl+A and Ctrl+C
COPY_CLIPBOARD_PAUSE_SECONDS = 0.7  # before the clipboard is read back
COPY_EMPTY_RETRIES = 3              # failed copies allowed before giving up
#: A copy smaller than this share of the biggest one so far is a failed
#: copy, not a shorter list: every Ctrl+C takes the whole list, which only
#: grows. Seen live - a copy came back with 14 characters while the list
#: held ten schools, and without this it counted as "no new school".
COPY_SHRINK_LIMIT = 0.5
#: A real copy always carries the search itself - Maps leaves the query in
#: the search box - or the end-of-list notice. A city with no result of
#: that kind still copies about 120 characters of page furniture, so length
#: alone cannot tell an empty city from a failed clipboard. Seen live: the
#: first copy of a keyword came back with 26 characters and the search ran
#: on to report no school at all.
COPY_MIN_CHARS = 60
DOWN_KEY_PAUSE_SECONDS = 0.12       # between two Down presses
#: Down presses with no new school before a search gives up. Counted in
#: PRESSES, not copies, because one press moves the list about 24 pixels
#: and a block of ten results only loads after a few hundred. Counting
#: copies made a search stop early as soon as DOWN_PRESSES_PER_ROUND was
#: lowered to 1.
SCROLL_NO_PROGRESS_PRESSES = 180


# ---------------------------------------------------------------------------
#  READING THE RESULTS LIST: Ctrl+A, Ctrl+C, the real clipboard
# ---------------------------------------------------------------------------
#  No card is clicked, no place page is opened and nothing is read out of the
#  cards' markup. The results list is focused, the page is selected and
#  copied with the keyboard, and the text comes back off the system
#  clipboard - the same thing a person does by hand.
#
#  Measured on a live Delhi search: the copied text GROWS as the list loads
#  (10 entries at 1,966 characters, 80 entries at 14,548). Maps keeps the
#  earlier results in the page, so a later snapshot contains the earlier
#  ones. Every snapshot is still kept and parsed, because that is only true
#  while Google chooses to keep them.
# ---------------------------------------------------------------------------

#: Google's own icon glyphs land in the clipboard as private-use characters.
_ICON_CHARS = re.compile(r"[\ue000-\uf8ff]")


def strip_icons(text):
    return _ICON_CHARS.sub("", text or "")


def focus_results_list(page, feed):
    """Put the keyboard on the results list, without clicking a result.

    Clicking a card would open that school's page, which is exactly what
    this script must not do, so focus is set on the feed element itself.
    """
    try:
        feed.evaluate("el => { if (!el.hasAttribute('tabindex')) "
                      "el.setAttribute('tabindex','-1'); el.focus(); }")
        return True
    except Exception as error:                          # noqa: BLE001
        LOG.debug("  could not focus the results list: %s", str(error)[:120])
        return False


def copy_results_text(page):
    """Ctrl+A, Ctrl+C, then read the clipboard. "" when the copy failed.

    The selection is the fallback, not the method: if the clipboard cannot
    be read at all the caller treats the search as incomplete rather than
    quietly reading the page some other way.
    """
    try:
        page.keyboard.press("Control+a")
        time.sleep(COPY_KEY_PAUSE_SECONDS)
        page.keyboard.press("Control+c")
        time.sleep(COPY_CLIPBOARD_PAUSE_SECONDS)
    except Exception as error:                          # noqa: BLE001
        LOG.warning("   the keyboard copy failed: %s", str(error)[:140])
        return "", "keyboard"

    for how, script in (("clipboard", "() => navigator.clipboard.readText()"),
                        ("selection",
                         "() => window.getSelection().toString()")):
        try:
            text = page.evaluate(script)
        except Exception as error:                      # noqa: BLE001
            LOG.debug("  %s unreadable: %s", how, str(error)[:100])
            continue
        if text and text.strip():
            return strip_icons(text), how
    return "", "nothing"


def copy_looks_real(text, query):
    """True when this looks like the results page rather than a failed copy.

    The query is in the search box, so every real copy carries it. An
    end-of-list notice counts too, for a city whose list is empty. Without
    this, a 26 character copy read as "this city has no schools".
    """
    low = " ".join((text or "").split()).lower()
    if not low:
        return False
    if any(marker in low for marker in END_OF_LIST_MARKERS):
        return True
    wanted = " ".join(str(query).split()).lower()
    if wanted and wanted in low:
        return True
    # the city on its own: Maps sometimes re-words the box
    tail = wanted.split(" in ")[-1]
    return bool(tail) and tail in low


def press_down(page, feed, presses=None):
    """Press Down and report whether the list actually moved.

    ``(moved, how)``. A Down that changes nothing is reported, not ignored:
    a list that will not move is a reason to stop, not to keep pressing.
    """
    presses = int(DOWN_PRESSES_PER_ROUND if presses is None else presses)
    try:
        before = feed.evaluate("el => el.scrollTop")
    except Exception:                                   # noqa: BLE001
        before = None
    for _ in range(presses):
        try:
            page.keyboard.press("ArrowDown")
        except Exception:                               # noqa: BLE001
            break
        time.sleep(DOWN_KEY_PAUSE_SECONDS)
    time.sleep(SCROLL_PAUSE_SECONDS)
    try:
        after = feed.evaluate("el => el.scrollTop")
    except Exception:                                   # noqa: BLE001
        after = None
    if before is not None and after is not None and after > before:
        return True, "down-arrow"
    # the key presses landed somewhere else: move the list directly, and say
    # so, rather than pretending the keyboard did it
    try:
        feed.evaluate("el => el.scrollBy(0, el.clientHeight * 0.9)")
        time.sleep(SCROLL_PAUSE_SECONDS)
        moved = feed.evaluate("el => el.scrollTop")
        if after is None or moved > after:
            return True, "scrolled (the Down keys did not move the list)"
    except Exception:                                   # noqa: BLE001
        pass
    return False, "stuck"


#: "4.8", "4.7 (149)", "4.4(762)", "No reviews" - never a name or address.
_RATING = re.compile(r"^(no reviews|\d([.,]\d)?\s*(\u2605|\(|$))", re.I)
#: Buttons and labels around a result, never part of the school.
_UI_WORDS = re.compile(r"^(website|directions|call|share|save|saved|open now|"
                       r"closed|sponsored|ad|results|recents|get app|"
                       r"rating|hours|all filters|menu|search)$", re.I)
#: The opening-hours line, which also carries the phone number.
_CLOSES = re.compile(r"\b(open|closed|closes|opens|temporarily|permanently)\b",
                     re.I)


def card_to_record(card, state, city):
    """One copied result entry turned into a record. Only what it shows."""
    record = blank_record()
    lines = [clean(line) for line in card["text"].splitlines()]
    lines = [line for line in lines if line]

    name = clean(card.get("label") or (lines[0] if lines else ""))
    # aria-label is the school name; the first text line repeats it
    record["School Name"] = name
    record["City"] = city
    record["State"] = state
    record["Country"] = COUNTRY
    record["School Google Maps Link"] = clean(card.get("href", ""))

    # The address is the tail of the category line:
    #     "CBSE School · ♿ · Metro Pillar 333, Lala Jagat Narayan Marg"
    # That line is looked for FIRST, on its own pass. Doing it in one pass
    # with the fallback below let a rating line that happened to come earlier
    # ("4.7 (149)") win the address slot.
    street = ""
    body = [line for line in lines[1:]
            if line != name and not _RATING.match(line)
            and not _UI_WORDS.match(line)]
    for line in body:
        if "·" not in line or _CLOSES.search(line):
            continue
        tail = line.split("·")[-1].strip()
        if tail and not _CLOSES.search(tail) and not tail.isdigit():
            street = tail
            break
    if not street:
        for line in body:
            if _CLOSES.search(line) or "·" in line:
                continue
            if len(line) > 6 and any(ch.isdigit() for ch in line):
                street = line
                break
    if street:
        record["School Address"] = street
        record["Area / Locality"] = street
    # from the address only: the NAME of a school often ends in its
    # affiliation number, which is also six digits
    pin = tidy_pin(street) or tidy_pin(
        "\n".join(line for line in lines[1:] if line != name))
    if pin:
        record["PIN Code"] = pin

    # Usually the tail of the hours line ("Open · Closes 6 pm · 011 4516
    # 1066"), but on a card with no hours the number sits on a line of its
    # own, so that case is tried too. Never the address line: a house number
    # is not a phone number.
    phone = ""
    for line in lines:
        if "·" in line and _CLOSES.search(line):
            phone = tidy_phone(line.split("·")[-1])
            if phone:
                break
    if not phone:
        for line in body:
            if line == street or "·" in line:
                continue
            phone = tidy_phone(line)
            if phone:
                break
    if phone:
        record["School Official Phone Number"] = phone

    coords = re.search(r"!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)",
                       record["School Google Maps Link"])
    if coords:
        record["School Latitude"] = coords.group(1)
        record["School Longitude"] = coords.group(2)
    return record


#: Pack a keyword's capture file once its schools have been extracted. Every
#: copy repeats the one before it, so the file is mostly the same text over
#: and over: a live Delhi keyword came to 1.1 MB and packs to 26 KB. Nothing
#: is dropped - the packed file is read straight back as before.
COMPRESS_CAPTURE_FILES = True


# ---------------------------------------------------------------------------
#  THE CAPTURED TEXT FILE
# ---------------------------------------------------------------------------
#  Every Ctrl+C is appended to the keyword's .txt as its own snapshot, with a
#  header saying which copy it was. The file is written during the search,
#  not at the end, so an interrupted keyword keeps everything it copied.
#
#  This file is then the ONLY input to the extraction step: the records are
#  produced by reading it, never by going back to the page.
# ---------------------------------------------------------------------------

SNAPSHOT_RULE = "=" * 36
_SNAPSHOT_SPLIT = re.compile(
    r"^=+\s*$\s*^KEYWORD:[^\n]*$\s*^SNAPSHOT:\s*(\d+)\s*$\s*^=+\s*$",
    re.M)

#: UI text that is never a school name.
_NOT_A_NAME = re.compile(
    r"^(results|sponsored|ads?|website|directions|share|saved|recents|"
    r"get app|rating|hours|all filters|update results when map moves|"
    r"back to top|you'?v?e? ?reached the end of the list\.?|"
    r"map data|terms|privacy|send product feedback|layers|"
    r"\d+ ?(km|m|mi|ft)|google|sign in|menu|search|nearby|"
    r"photos|reviews?|about|overview|claim this business|suggest an edit|"
    r"add missing information|add place'?s phone number|add website)$", re.I)
_RATING_LINE = re.compile(r"^(no reviews|\d[.,]\d\s*\(?[\d,]*\)?)$", re.I)


def keyword_slug(keyword):
    return re.sub(r"[^a-z0-9]+", "_", clean(keyword).lower()).strip("_")


def keyword_txt_path(state, city, keyword):
    return city_file(state, city, "%s.txt" % keyword_slug(keyword))


def gemini_school_path(state, city, record):
    """One school's raw answer. Its own file, so it can go on its own.

    The school's name makes it readable and a short digest of its identity
    keeps two schools of the same name apart.
    """
    name = re.sub(r"[^A-Za-z0-9]+", "_",
                  clean(record.get("School Name", ""))).strip("_").lower()
    digest = hashlib.sha1(
        record_identity(record).encode("utf-8", "replace")).hexdigest()[:8]
    return city_file(state, city,
                     "gemini_%s_%s.txt" % (name[:48] or "school", digest))


def start_keyword_txt(path, query, url):
    """Begin (or continue) a keyword's capture file. Never overwrites."""
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "a", encoding="utf-8") as handle:
        handle.write("\n%s\n# SEARCH : %s\n# URL    : %s\n# STARTED: %s\n%s\n"
                     % ("#" * 72, query, url,
                        time.strftime("%Y-%m-%d %H:%M:%S"), "#" * 72))
    return path


def append_snapshot(path, keyword_query, number, text, how):
    """Append one Ctrl+C exactly as it came off the clipboard."""
    with open(path, "a", encoding="utf-8") as handle:
        handle.write("\n%s\nKEYWORD: %s\nSNAPSHOT: %d\n%s\n\n"
                     % (SNAPSHOT_RULE, keyword_query, number, SNAPSHOT_RULE))
        handle.write("# read via: %s | %d characters | %s\n\n"
                     % (how, len(text), time.strftime("%H:%M:%S")))
        handle.write(text.rstrip() + "\n")
        handle.flush()
        os.fsync(handle.fileno())


def close_keyword_txt(path, snapshots, found, reason, complete):
    with open(path, "a", encoding="utf-8") as handle:
        handle.write("\n%s\n# FINISHED : %s\n# SNAPSHOTS: %d\n# SCHOOLS  : %d\n"
                     "# STOPPED  : %s\n# STATUS   : %s\n%s\n"
                     % ("#" * 72, time.strftime("%Y-%m-%d %H:%M:%S"),
                        snapshots, found, reason,
                        "COMPLETE" if complete else "INCOMPLETE", "#" * 72))


def read_snapshots(text):
    """Every snapshot in a capture file, in order."""
    if not text:
        return []
    pieces = _SNAPSHOT_SPLIT.split(text)
    # split() gives [before, number, body, number, body, ...]
    out = []
    for index in range(1, len(pieces) - 1, 2):
        out.append({"number": int(pieces[index]), "text": pieces[index + 1]})
    return out


def split_result_entries(text):
    """The result entries inside one copied snapshot.

    Google repeats a result's name on two consecutive lines - once as the
    heading and once as the link's own text. That doubled line is the only
    reliable boundary in the copied text, and it is what separates one
    school from the next.
    """
    lines = [clean(strip_icons(line)) for line in (text or "").splitlines()]
    starts = []
    for index in range(len(lines) - 1):
        name = lines[index]
        if not name or name != lines[index + 1]:
            continue
        if len(name) < 3 or _NOT_A_NAME.match(name) or _RATING_LINE.match(name):
            continue
        starts.append(index)

    entries = []
    for position, index in enumerate(starts):
        stop = starts[position + 1] if position + 1 < len(starts) else len(lines)
        body = [line for line in lines[index + 2:stop] if line]
        # the end-of-list notice belongs to the list, not to the last school
        body = [line for line in body
                if not any(marker in line.lower()
                           for marker in END_OF_LIST_MARKERS)]
        entries.append({"name": lines[index], "lines": body})
    return entries


def parse_maps_text(text, state, city):
    """Records for every school entry in one copied snapshot."""
    records = []
    for entry in split_result_entries(text):
        record = card_to_record(
            {"href": "", "label": entry["name"],
             "text": "\n".join([entry["name"]] + entry["lines"])},
            state, city)
        if clean(record["School Name"]):
            records.append(record)
    return records


def entry_identity(record):
    """What makes this entry THIS school rather than another campus.

    The copied text carries no place identifier, so identity is the name
    together with whatever the entry said about where it is. Two branches of
    one chain keep different addresses and so stay apart; the same entry
    copied in ten snapshots collapses to one.
    """
    name = name_key(record.get("School Name", "")) or clean(
        record.get("School Name", "")).lower()
    where = clean(record.get("School Address", "")).lower()[:40]
    phone = tidy_phone(record.get("School Official Phone Number", ""))
    return "%s|%s|%s" % (name, where, phone)


def capture_files(path):
    """The plain and packed halves of one capture file, oldest first."""
    found = []
    if os.path.exists(path + ".gz"):
        found.append(path + ".gz")
    if os.path.exists(path):
        found.append(path)
    return found


def read_capture(path):
    """The text of a capture file, packed or not."""
    text = []
    for name in capture_files(path):
        try:
            if name.endswith(".gz"):
                with gzip.open(name, "rt", encoding="utf-8") as handle:
                    text.append(handle.read())
            else:
                with open(name, encoding="utf-8") as handle:
                    text.append(handle.read())
        except (OSError, EOFError) as error:
            LOG.warning("   %s could not be read (%s)", os.path.basename(name),
                        str(error)[:80])
    return "\n".join(text)


def pack_capture_file(path):
    """Store a finished capture file packed. ``(before, after)`` bytes.

    The copied text is kept in full - this only stops the same text being
    stored a hundred times at full size. The packed file is unpacked and
    compared with the original before the original is removed, so a bad
    write cannot lose the text.
    """
    if not COMPRESS_CAPTURE_FILES or not os.path.exists(path):
        return 0, 0
    before = os.path.getsize(path)
    if before < 64 * 1024:
        return before, before            # not worth it

    packed = path + ".gz"
    try:
        with open(path, "rb") as handle:
            raw = handle.read()
        if os.path.exists(packed):       # a resumed keyword: join the two
            with gzip.open(packed, "rb") as handle:
                raw = handle.read() + b"\n" + raw
        temporary = packed + ".tmp"
        with gzip.open(temporary, "wb", compresslevel=9) as handle:
            handle.write(raw)
        with gzip.open(temporary, "rb") as handle:
            if handle.read() != raw:
                raise OSError("the packed file did not read back the same")
        os.replace(temporary, packed)
        os.remove(path)
    except (OSError, EOFError) as error:
        LOG.debug("   could not pack %s: %s", path, str(error)[:120])
        for leftover in (packed + ".tmp",):
            if os.path.exists(leftover):
                os.remove(leftover)
        return before, before

    after = os.path.getsize(packed)
    LOG.info("   capture text packed: %.0f KB -> %.0f KB (%.0f%% less disk, "
             "every byte still there)", before / 1024.0, after / 1024.0,
             100.0 * (before - after) / max(before, 1))
    return before, after


def pack_city_text(state, city):
    """Pack everything a finished city wrote. Returns bytes saved."""
    if not COMPRESS_CAPTURE_FILES:
        return 0
    saved = 0
    for path in city_files(state, city):
        if not path.endswith(".txt"):
            continue
        before, after = pack_capture_file(path)
        saved += max(before - after, 0)
    return saved


_LOG_DATE = re.compile(r"^(\d{4}-\d{2}-\d{2})[ T]")


def trim_log(path=None, keep_days=None):
    """Drop lines older than LOG_KEEP_DAYS from the single log file.

    Called once as the script starts, before anything is written, so the
    file always holds roughly the last ten days and nothing older. A line
    with no date of its own - a traceback, a wrapped message - belongs to
    the line above it and is kept or dropped with it.

    Returns ``(lines_kept, lines_dropped)``.
    """
    path = path or LOG_FILE
    keep_days = LOG_KEEP_DAYS if keep_days is None else keep_days
    if not keep_days or not os.path.exists(path):
        return 0, 0
    cutoff = (datetime.date.today()
              - datetime.timedelta(days=int(keep_days))).isoformat()
    try:
        with open(path, encoding="utf-8", errors="replace") as handle:
            lines = handle.readlines()
    except OSError:
        return 0, 0

    kept, dropped, keeping = [], 0, True
    for line in lines:
        found = _LOG_DATE.match(line)
        if found:
            keeping = found.group(1) >= cutoff
        if keeping:
            kept.append(line)
        else:
            dropped += 1
    if not dropped:
        return len(kept), 0

    temporary = path + ".tmp"
    try:
        with open(temporary, "w", encoding="utf-8") as handle:
            handle.writelines(kept)
        os.replace(temporary, path)
    except OSError:
        if os.path.exists(temporary):
            os.remove(temporary)
        return len(lines), 0
    return len(kept), dropped


def records_from_txt(path, state, city):
    """Read a capture file and turn its snapshots into school records.

    This is the extraction step. It runs on the saved text, so it can be
    re-run, inspected and corrected without collecting anything again. The
    snapshots overlap heavily - each one usually contains every entry seen
    so far - and the duplicates are removed here.
    """
    text = read_capture(path)
    if not text.strip():
        return []
    snapshots = read_snapshots(text)
    if not snapshots:                       # a file with no snapshot header
        snapshots = [{"number": 1, "text": text}]
    records, seen = [], {}
    for snapshot in snapshots:
        for record in parse_maps_text(snapshot["text"], state, city):
            identity = entry_identity(record)
            if identity in seen:
                merge_records(seen[identity], record)
                continue
            seen[identity] = record
            records.append(record)
    return records


class SearchResult:
    """What one keyword-in-one-city search produced, and why it stopped."""

    def __init__(self, query):
        self.query = query
        self.records = []
        self.rounds = 0
        self.complete = False
        self.reason = "not started"
        self.txt_path = ""
        self.snapshots = 0

    def __repr__(self):
        return "SearchResult(%r, %d records, complete=%s, %s)" % (
            self.query, len(self.records), self.complete, self.reason)


def collect_search(browser, keyword, state, city, limit=None):
    """Collect one keyword's results list by copying it with the keyboard.

    The loop is fixed and never varies:

        focus the list -> Ctrl+A -> Ctrl+C -> read the clipboard
        -> save the snapshot -> look for the end-of-list message
        -> Down -> wait -> repeat

    Nothing is copied after a Down without the text before it having been
    saved first, and the search is only complete when Google's own
    end-of-list message appears in the copied text.
    """
    query = build_query(keyword, city)
    url = maps_search_url(query)
    result = SearchResult(query)

    LOG.info("-" * 72)
    LOG.info("rank %s | %s, %s | keyword: %s", RUNTIME["rank_label"], city,
             state, keyword)
    LOG.info("GOOGLE MAPS QUERY: %s", query)
    LOG.info("URL: %s", url)
    LOG.info("-" * 72)

    # the one visible tab follows the run: every new city and every new
    # keyword retunes it to the search being collected right now
    page = browser.page()
    spare = browser.close_spare_tabs()
    if spare:
        LOG.debug("   closed %d stray tab(s)", spare)
    browser.goto(page, url)
    browser.show_tab()
    time.sleep(PAGE_SETTLE_SECONDS)

    body = browser.text_of(page, "body")[:4000]
    if looks_blocked(body):
        raise AccessBlocked(
            "Google Maps served a bot challenge for %r. Nothing was "
            "collected; this is reported rather than worked around." % query)

    txt_path = keyword_txt_path(state, city, keyword)
    start_keyword_txt(txt_path, query, url)
    result.txt_path = txt_path

    # confirm the right page is open: a results list, or a single place
    try:
        page.wait_for_selector(FEED_SELECTOR, timeout=BROWSER_TIMEOUT_MS)
    except (PlaywrightTimeout, PlaywrightError):
        if "/maps/place/" in page.url:
            # Google answered the SEARCH with one place instead of a list.
            # The script clicked nothing; this is the page the search URL
            # itself returned, so its one result is copied and kept.
            text, how = copy_results_text(page)
            append_snapshot(txt_path, query, 1,
                            text or browser.text_of(page, "body")[:2000], how)
            title = clean(page.title()).split(" - ")[0]
            record = blank_record()
            record.update({"School Name": title, "City": city, "State": state,
                           "Country": COUNTRY})
            if clean(record["School Name"]):
                result.records = [record]
            result.complete = True
            result.reason = "one matching place, no result list"
            result.snapshots = 1
            close_keyword_txt(txt_path, 1, len(result.records),
                              result.reason, True)
            LOG.info("   a single place page, not a list: %d record",
                     len(result.records))
            return result
        result.reason = "the results list never appeared (page did not load)"
        close_keyword_txt(txt_path, 0, 0, result.reason, False)
        LOG.warning("   %s", result.reason)
        return result

    feed = page.query_selector(FEED_SELECTOR)
    if feed is None:
        result.reason = "the results list disappeared before it could be read"
        close_keyword_txt(txt_path, 0, 0, result.reason, False)
        LOG.warning("   %s", result.reason)
        return result

    seen = set()
    empty_copies = 0
    biggest_copy = 0
    idle = 0
    snapshot = 0
    how_read = how_moved = "n/a"

    # No round cap and no clock: a keyword ends when Google says the list
    # has ended, or when something is plainly wrong - the clipboard failing
    # or the list refusing to move. Either of those leaves the keyword
    # open, to be picked up again from the checkpoint.
    round_number = 0
    while True:
        round_number += 1
        result.rounds = round_number

        # 1. focus the list   2. Ctrl+A   3. Ctrl+C   4. read the clipboard
        focus_results_list(page, feed)
        text, how_read = copy_results_text(page)

        too_small = (len(text.strip()) < COPY_MIN_CHARS
                     or not copy_looks_real(text, query))
        shrank = (biggest_copy
                  and len(text) < biggest_copy * COPY_SHRINK_LIMIT)
        if too_small or shrank:
            empty_copies += 1
            LOG.warning("   copy %d came back with %d character(s) (%s), "
                        "after %d - retry %d of %d", round_number, len(text),
                        how_read, biggest_copy, empty_copies,
                        COPY_EMPTY_RETRIES)
            if empty_copies > COPY_EMPTY_RETRIES:
                result.reason = ("the clipboard could not be read properly "
                                 "%d times - the search is left incomplete "
                                 "rather than read some other way"
                                 % empty_copies)
                break
            focus_results_list(page, feed)
            time.sleep(SCROLL_PAUSE_SECONDS)
            continue
        empty_copies = 0
        biggest_copy = max(biggest_copy, len(text))

        # 5. save the copied text BEFORE anything else happens to the page
        snapshot += 1
        append_snapshot(txt_path, query, snapshot, text, how_read)
        result.snapshots = snapshot

        before = len(seen)
        for record in parse_maps_text(text, state, city):
            seen.add(entry_identity(record))
        gained = len(seen) - before
        LOG.info("   snapshot %-3d | %-26s | %6d chars | +%-3d new | %-4d "
                 "schools", snapshot, ("%s in %s" % (keyword, city))[:26],
                 len(text), gained, len(seen))
        say_status("copying from Maps - %d school(s) so far (copy %d)"
                   % (len(seen), snapshot))

        # 6. the end-of-list message, in the copied text itself
        if any(marker in text.lower() for marker in END_OF_LIST_MARKERS):
            result.complete = True
            result.reason = "end-of-list message in the copied text"
            break
        if limit and len(seen) >= limit:
            result.reason = "the --limit of %d was reached" % limit
            break
        if gained == 0:
            idle += 1
            if idle * max(int(DOWN_PRESSES_PER_ROUND), 1) \
                    >= SCROLL_NO_PROGRESS_PRESSES:
                result.reason = ("no new school in %d copies (%d Down "
                                 "presses) and no end-of-list message"
                                 % (idle, idle * DOWN_PRESSES_PER_ROUND))
                break
        else:
            idle = 0
        # 7. only now the Down key, and only if it really moves the list
        moved, how_moved = press_down(page, feed)
        if not moved:
            result.reason = "the results list stopped moving"
            break

    # the extraction runs on the saved text, never on the page
    result.records = records_from_txt(txt_path, state, city)
    if limit:
        result.records = result.records[:limit]
    close_keyword_txt(txt_path, result.snapshots, len(result.records),
                      result.reason, result.complete)

    LOG.info("SEARCH DONE  %r", query)
    LOG.info("   snapshots     : %d  (read via %s, list moved by %s)",
             result.snapshots, how_read, how_moved)
    LOG.info("   schools found : %d", len(result.records))
    LOG.info("   with phone    : %d",
             sum(1 for r in result.records
                 if r["School Official Phone Number"]))
    LOG.info("   with address  : %d",
             sum(1 for r in result.records if r["School Address"]))
    LOG.info("   stopped on    : %s   (%s)", result.reason,
             "COMPLETE" if result.complete else "INCOMPLETE")
    LOG.info("   text saved to : %s", result.txt_path)
    if not result.complete:
        LOG.warning("   this search is INCOMPLETE - the keyword stays open "
                    "and will be collected again")
    return result


def merge_into_index(index, records):
    """Add records to a city's index. Returns (new, merged)."""
    new = merged = 0
    for record in records:
        _kept, how = index.add(record)
        if how == "new":
            new += 1
        else:
            merged += 1
    return new, merged


# --- Gemini ----------------------------------------------------------------
GEMINI_ENABLED = True
GEMINI_NEW_CHAT_EACH_SCHOOL = True  # so one school cannot colour the next
GEMINI_POLL_SECONDS = 1.5           # how often the answer is re-read
GEMINI_SETTLE_SECONDS = 4           # it must stop growing for this long
GEMINI_INPUT_TIMEOUT = 15           # for the chat box to appear
INPUT_BOX_POLL_SECONDS = 0.5        # between two looks for it
SCHOOL_DELAY_SECONDS = (2.0, 4.0)   # between two schools
KEEP_RAW_GEMINI_ANSWERS = True      # kept under their school in a .txt


# ===========================================================================
#  GEMINI RESEARCH
# ===========================================================================
#  The answer is taken from the CURRENT response container, not from a copy
#  of the whole page: a whole-page capture also contains the question, the
#  previous answers and the interface, and mixing two schools' answers is the
#  one mistake that silently corrupts the dataset.
#
#  The wait is on the answer itself, not on a fixed number of seconds: the
#  response counts as finished when it carries real field lines AND has
#  stopped growing for GEMINI_SETTLE_SECONDS.
# ---------------------------------------------------------------------------

INPUT_SELECTORS = (
    'rich-textarea div[contenteditable="true"]',
    'div[contenteditable="true"][role="textbox"]',
    'div[contenteditable="true"]',
    "textarea",
)
RESPONSE_SELECTORS = (
    "model-response",
    "message-content",
    'div[data-response-index]',
    ".model-response-text",
)


def profile_is_free(profile_dir, wait=None):
    """Wait for Chromium to let go of a profile. ``(free, waited)``.

    Chromium keeps a SingletonLock in the user-data-dir while it is running
    and removes it on a clean exit. Opening the next browser before that
    happens fails with "the profile is already in use", so each stage waits
    here before the next one starts.
    """
    lock = os.path.join(profile_dir, "SingletonLock")
    deadline = time.time() + float(PROFILE_RELEASE_WAIT_SECONDS
                                   if wait is None else wait)
    started = time.time()
    while time.time() < deadline:
        if not os.path.lexists(lock):
            return True, time.time() - started
        time.sleep(WINDOW_CHECK_POLL_SECONDS)
    return not os.path.lexists(lock), time.time() - started


def where_the_window_goes(show, headless=False):
    """One line saying, in plain words, whether you will see anything.

    The rule is simple and has no exceptions: a window appears on your
    desktop only with --show.
    """
    if show:
        return "VISIBLE on your desktop (--show)"
    if headless:
        return "hidden - no window is drawn at all (--headless)"
    if running_under_xvfb():
        return ("hidden on the virtual screen xvfb-run made (%s) - nothing "
                "appears on your desktop. Add --show to watch it."
                % (os.environ.get("DISPLAY") or "?"))
    return ("hidden - no window is drawn on your desktop. Add --show to "
            "watch it.")


#: What a Playwright-launched Chromium calls itself. The sweep below
#: requires BOTH of these: a program that really is one of these browsers,
#: AND one of this script's own profile folders on its command line.
BROWSER_PROGRAM_NAMES = ("chrome", "chromium", "headless_shell", "msedge")


def this_script_browsers():
    """PIDs of browsers holding one of THIS script's profile folders.

    Two conditions, both required:

      1. the program really is a browser - checked on the executable it is
         running, not on words anywhere in its command line. Without this
         a shell that merely mentions the path matches, and the sweep
         kills the terminal it was started from.
      2. one of this script's own profile folders is its --user-data-dir.

    Your own Chrome or Brave passes 1 and fails 2, so it is never touched.
    """
    marks = [m for m in (BrowserDirector.SCRATCH_PREFIX, BROWSER_PROFILE_DIR,
                         GEMINI_PROFILE_DIR) if m]
    found = []
    try:
        names = os.listdir("/proc")
    except OSError:
        return found
    for name in names:
        if not name.isdigit():
            continue
        try:
            with open("/proc/%s/cmdline" % name, "rb") as handle:
                raw = handle.read().decode("utf-8", "replace")
        except OSError:
            continue
        if not raw:
            continue
        parts = raw.split("\0")
        # Chromium rewrites its own argv, and a running browser turns up
        # with the WHOLE command line sitting in argv[0] instead of just
        # the program. Taking the first word of it covers both shapes;
        # without this the sweep finds nothing it is meant to find.
        first = (parts[0] or "").split()
        if not first:
            continue
        program = os.path.basename(first[0]).lower()
        if not any(word in program for word in BROWSER_PROGRAM_NAMES):
            continue
        # The profile must be given as the --user-data-dir option, not
        # merely mentioned somewhere. The check above is what keeps a
        # shell that names the path from matching: its first word is a
        # shell, not a browser.
        found_mark = False
        for piece in raw.replace("\0", " ").split():
            if not piece.startswith("--user-data-dir="):
                continue
            profile = piece.split("=", 1)[1]
            if any(mark in profile for mark in marks):
                found_mark = True
                break
        if found_mark:
            found.append(int(name))
    return found


def other_copies_running():
    """True when another find_school.py is running besides this one."""
    me, parent = os.getpid(), os.getppid()
    others = []
    try:
        names = os.listdir("/proc")
    except OSError:
        return False
    for name in names:
        if not name.isdigit() or int(name) in (me, parent):
            continue
        try:
            with open("/proc/%s/cmdline" % name, "rb") as handle:
                line = handle.read().decode("utf-8", "replace")
        except OSError:
            continue
        if "find_school.py" in line and "python" in line.lower():
            others.append(int(name))
    return bool(others)


def still_running(pid):
    """True while a process is really there.

    A killed child whose parent has not collected it stays in /proc as a
    zombie, so the state field is read rather than the folder's existence.
    """
    try:
        with open("/proc/%d/stat" % pid, "rb") as handle:
            fields = handle.read().decode("utf-8", "replace").rsplit(")", 1)
    except OSError:
        return False
    state = fields[-1].split()[0] if len(fields) > 1 else ""
    return state not in ("Z", "X", "x")


def sweep_orphan_browsers():
    """Close a browser window left behind by a run that was killed.

    A run stopped with kill -9 cannot close its own Chromium, so that
    window stays on the desktop - and the next run then looks as though it
    opened a browser when it did not. This is the only thing that can put
    one of this script's windows on your screen without --show.

    Skipped entirely while another copy of the script is running, whose
    browser is not an orphan. Only this script's own profiles are touched.
    """
    if other_copies_running():
        return 0
    pids = this_script_browsers()
    if not pids:
        return 0
    for pid in pids:
        try:
            os.kill(pid, signal.SIGTERM)
        except OSError:
            pass
    deadline = time.time() + PROFILE_RELEASE_WAIT_SECONDS
    while time.time() < deadline:
        if not [pid for pid in pids if still_running(pid)]:
            break
        time.sleep(BROWSER_EXIT_POLL_SECONDS)
    for pid in [p for p in pids if still_running(p)]:
        try:
            os.kill(pid, signal.SIGKILL)
        except OSError:
            pass
    # what actually went, not what was asked to go
    closed = [pid for pid in pids if not still_running(pid)]
    if closed:
        LOG.info("   %d browser window(s) left by a killed run were closed",
                 len(closed))
    stubborn = [pid for pid in pids if still_running(pid)]
    if stubborn:
        LOG.warning("   %d browser(s) from an earlier run would not close "
                    "(pid %s) - close them by hand if a window is in the way",
                    len(stubborn),
                    ", ".join(str(pid) for pid in stubborn[:4]))
    return len(closed)


def sweep_old_browser_profiles(older_than=None):
    """Remove scratch profiles left by a run that was killed.

    A run that ends normally clears its own. One stopped with kill -9 cannot,
    so the next run tidies up after it. Only folders older than an hour are
    touched, so a second copy of the script running right now is left alone.
    """
    root = tempfile.gettempdir()
    cutoff = time.time() - (SCRATCH_SWEEP_AGE_SECONDS if older_than is None
                            else older_than)
    removed = 0
    try:
        names = os.listdir(root)
    except OSError:
        return 0
    for name in names:
        if not name.startswith(BrowserDirector.SCRATCH_PREFIX):
            continue
        path = os.path.join(root, name)
        try:
            if os.path.isdir(path) and os.path.getmtime(path) < cutoff:
                shutil.rmtree(path, ignore_errors=True)
                removed += 1
        except OSError:
            pass
    if removed:
        LOG.debug("   %d browser profile(s) from an earlier run removed from "
                  "%s", removed, root)
    return removed


class BrowserDirector:
    """Exactly one Chromium open at a time, on one shared driver.

    Two visible windows side by side was the problem this solves: Maps and
    Gemini each get the screen to themselves, and the other one is not
    merely hidden but closed.

    The Playwright driver is started once for the whole run and handed to
    each browser, because a process may start the sync driver only once -
    closing the Maps browser must not take the driver Gemini needs with it.
    """

    KINDS = {"maps": "Google Maps browser", "gemini": "Gemini browser",
             "search": "Google Search browser"}

    def __init__(self, show=False, mode=None):
        self.show = bool(show)
        self.mode = mode
        self.current = None
        self.kind = ""
        self._play = None
        self._scratch = None
        self.display = ""
        self.authority = ""
        self.on_close = {}          # kind -> callback, so a client can let go

    # ------------------------------------------------------------- driver
    def _driver(self):
        if self._play is None:
            if not PLAYWRIGHT_AVAILABLE:
                raise BrowserUnavailable(
                    "Playwright is not installed. Run:\n"
                    "    python3 -m pip install playwright\n"
                    "    python3 -m playwright install chromium")
            self._play = sync_playwright().start()
        return self._play

    SCRATCH_PREFIX = "school_browsers_"

    def scratch(self):
        """A temporary home for the profiles, made once per run."""
        if self._scratch is None:
            # A killed run leaves both a profile folder and, worse, a live
            # Chromium window. Both go before this run opens anything.
            sweep_orphan_browsers()
            sweep_old_browser_profiles()
            self._scratch = tempfile.mkdtemp(prefix=self.SCRATCH_PREFIX)
        return self._scratch

    def profile_of(self, kind):
        """Where this browser's profile lives: a temporary folder, always.

        It is a real directory rather than Playwright's own hidden one,
        because the clipboard permission, the crash-flag repair and the
        --show window check all need a path this script knows. It is
        removed when the browser closes, so nothing is left in the project
        folder and nothing is carried between runs.
        """
        return os.path.join(self.scratch(), kind)

    # -------------------------------------------------------------- opening
    def open(self, kind):
        """The browser for this stage, opening it if it is not the one open."""
        if kind not in self.KINDS:
            raise ValueError("unknown browser %r" % kind)
        if self.kind == kind and self.current is not None \
                and self.current.alive():
            return self.current

        ok, detail = self.close()
        if not ok:
            # never two competing browsers: if the old one will not go, the
            # new one is not started
            raise BrowserUnavailable(
                "The %s would not close (%s), so the %s was not opened. "
                "Nothing has been lost; close that window and run again."
                % (self.KINDS.get(self.kind, "browser"), detail,
                   self.KINDS[kind]))

        profile = self.profile_of(kind)
        # a profile made for this run cannot be held by an earlier one
        free, waited = True, 0
        if not free:
            raise BrowserUnavailable(
                "The profile %s is still locked after %.0fs. A Chromium from "
                "an earlier run is holding it; close that window and start "
                "again." % (profile, waited))
        if waited > 1:
            LOG.debug("   waited %.0fs for %s to be released", waited,
                      os.path.basename(profile))

        browser = Browser(mode=self.mode, show=self.show, profile_dir=profile,
                          label=self.KINDS[kind], playwright=self._driver(),
                          display=self.display, authority=self.authority)
        browser.start()
        self.current, self.kind = browser, kind
        if self.show and browser.display:
            # the first visible browser works the screen out; the second is
            # told, so it cannot pick a different one
            self.display, self.authority = browser.display, browser.authority
        LOG.info("   [BROWSER] %s open%s", self.KINDS[kind],
                 " on %s" % browser.display if browser.display else "")
        return browser

    def maps(self):
        return self.open("maps")

    def gemini(self):
        return self.open("gemini")

    def search(self):
        """The browser the official-contact searches run in.

        Its own kind, so opening it closes Maps or Gemini first and the
        one-browser-at-a-time rule still holds.
        """
        return self.open("search")

    # -------------------------------------------------------------- closing
    def close(self):
        """Close whatever is open. ``(ok, detail)``; ok when nothing is left."""
        if self.current is None:
            return True, "nothing was open"
        kind, browser = self.kind, self.current
        callback = self.on_close.get(kind)
        if callback is not None:
            try:
                callback()
            except Exception:                           # noqa: BLE001
                pass
        try:
            browser.close()
        except Exception as error:                      # noqa: BLE001
            LOG.error("   [BROWSER] the %s did not close cleanly: %s",
                      self.KINDS.get(kind, kind), str(error)[:140])
            return False, str(error)[:140]
        self.current, self.kind = None, ""

        free, waited = profile_is_free(self.profile_of(kind))
        if not free:
            return False, "its profile is still locked"
        if self.show and self.display:
            left = window_of_profile(self.display, self.authority,
                                     self.profile_of(kind))
            if left is not None:
                return False, "its window is still on %s" % self.display
        LOG.info("   [BROWSER] %s closed%s", self.KINDS.get(kind, kind),
                 " after %.0fs" % waited if waited > 1 else "")
        return True, "closed"

    def shutdown(self):
        """End of the run: close the browser, stop the driver, tidy up."""
        self.close()
        try:
            if self._play is not None:
                self._play.stop()
        except Exception:                               # noqa: BLE001
            pass
        self._play = None
        if self._scratch is not None:
            size = 0
            for folder, _dirs, files in os.walk(self._scratch):
                for name in files:
                    try:
                        size += os.path.getsize(os.path.join(folder, name))
                    except OSError:
                        pass
            shutil.rmtree(self._scratch, ignore_errors=True)
            LOG.info("Browser profiles removed (%.0f MB, nothing left in the "
                     "project folder)", size / 1048576.0)
            self._scratch = None

    # ------------------------------------------------------------- checking
    def windows_open(self):
        """How many of this script's Chromium windows are on the screen."""
        if not self.show or not self.display:
            return 0
        ours = {os.path.abspath(self.profile_of(k)) for k in self.KINDS}
        return len([w for w in desktop_windows(self.display, self.authority)
                    if w["profile"]
                    and os.path.abspath(w["profile"]) in ours])


class Gemini:
    """The Gemini web application, kept out of the operator's way.

    It gets a browser of its own, asked for from the director, which closes
    the Maps browser first - only one Chromium is ever open. Its profile is
    separate because two Chromium instances cannot share one directory.
    """

    def __init__(self, source, separate=None):
        #: a BrowserDirector in a real run, a plain browser in the tests
        self.director = source if hasattr(source, "gemini") else None
        self.browser = None if self.director else source
        #: only the tests pass a plain browser and ask for a tab in it; a
        #: real run always gets a browser of its own from the director
        self.separate = True if separate is None else bool(separate)
        self._own = None
        self._page = None
        self.asked = 0
        self.where = "not started"
        if self.director is not None:
            # when the director closes the Gemini browser, this page is gone
            self.director.on_close["gemini"] = self.forget_page

    def forget_page(self):
        self._page = None
        self._own = None

    def _host(self):
        """The browser Gemini should use, opening it if it is not open.

        With a director this is the ONLY browser open: asking for it closes
        the Maps browser first. Nothing here pre-opens anything - the first
        call happens when the first school is about to be researched.
        """
        if self.director is not None:
            self._own = self.director.gemini()
            return self._own
        if not self.separate:
            return self.browser
        return self.browser

    def page(self):
        if self._page is not None and not self._page.is_closed():
            return self._page
        host = self._host()
        if self.director is None and host is self.browser:
            # tests only: a background tab of the one browser there is
            self._page = self.browser.new_page()
            self.where = "a background tab of the main window"
        else:
            self._page = host.page()
            host.close_spare_tabs()
            self.where = "its own browser"
        host.goto(self._page, GEMINI_URL)
        time.sleep(PAGE_SETTLE_SECONDS)
        if host is self.browser:
            self.browser.show_tab()
        LOG.info("Gemini is running in %s", self.where)
        return self._page

    def _give_the_screen_back(self):
        """Tests only: put the Maps tab back in front of the Gemini tab."""
        if self.director is None and not self.separate and self.browser:
            self.browser.show_tab()

    def close(self):
        """Let go of the page. The director owns the browser itself."""
        try:
            if self._page is not None and not self._page.is_closed():
                self._page.close()
        except Exception:                               # noqa: BLE001
            pass
        self._page = None
        if self.director is None and self._own is not None:
            self._own.close()
        self._own = None

    # ------------------------------------------------------------- readiness
    def input_box(self, page, timeout=None):
        deadline = time.time() + (GEMINI_INPUT_TIMEOUT if timeout is None
                                  else timeout)
        while time.time() < deadline:
            for selector in INPUT_SELECTORS:
                try:
                    node = page.query_selector(selector)
                except Exception:                       # noqa: BLE001
                    continue
                if node is not None:
                    return node, selector
            time.sleep(INPUT_BOX_POLL_SECONDS)
        return None, ""

    def check_ready(self):
        """``(ready, message)``. The chat box decides, not the page text."""
        try:
            page = self.page()
        except BrowserUnavailable as error:
            return False, str(error)
        node, selector = self.input_box(page)
        self._give_the_screen_back()
        if node is not None:
            return True, ("the Gemini chat box is available (%s), in %s"
                          % (selector, self.where))
        body = self.browser.text_of(page, "body")[:3000].lower()
        if any(sign in body for sign in SIGN_IN_SIGNS) or "sign in" in body:
            return False, ("Gemini is showing a sign-in page. Sign in by hand "
                           "in the browser window and run again - this "
                           "program does not attempt to get past it.")
        if looks_blocked(body):
            return False, "Gemini is showing a bot challenge, not a chat box."
        return False, "the Gemini chat box did not appear on the page."

    # -------------------------------------------------------------- asking
    def _new_chat(self, page):
        if not GEMINI_NEW_CHAT_EACH_SCHOOL:
            return
        self._host().goto(page, GEMINI_URL)
        time.sleep(PAGE_SETTLE_SECONDS)

    def _response_text(self, page):
        """The text of the CURRENT (last) response container only."""
        best = ""
        for selector in RESPONSE_SELECTORS:
            try:
                nodes = page.query_selector_all(selector)
            except Exception:                           # noqa: BLE001
                continue
            if not nodes:
                continue
            try:
                text = nodes[-1].inner_text() or ""
            except Exception:                           # noqa: BLE001
                continue
            if len(text) > len(best):
                best = text
        return best

    def _wait_for_response(self, page):
        """Wait the floor, then inspect the CURRENT response until it settles.

        Only the last response container is read - never a copy of the whole
        page, which would also carry the prompt, the interface and the
        previous school's answer.
        """
        deadline = time.time() + GEMINI_ANSWER_TIMEOUT
        LOG.debug("      waiting %ds before the first look",
                  GEMINI_FIRST_READ_SECONDS)
        time.sleep(GEMINI_FIRST_READ_SECONDS)
        previous, steady_since, fullest = "", None, ""
        while time.time() < deadline:
            current = self._response_text(page)
            if current and looks_like_an_answer(current):
                if len(current) > len(fullest):
                    fullest = current
                if current == previous:
                    if steady_since is None:
                        steady_since = time.time()
                    elif time.time() - steady_since >= GEMINI_SETTLE_SECONDS:
                        return current
                else:
                    previous, steady_since = current, None
            time.sleep(GEMINI_POLL_SECONDS)
        if fullest:
            LOG.warning("   the answer never settled within %ds - using the "
                        "fullest text that arrived", GEMINI_ANSWER_TIMEOUT)
            return fullest
        return ""

    def ask(self, prompt):
        """Send one prompt and return the finished answer text."""
        page = self.page()
        self._new_chat(page)
        node, _selector = self.input_box(page)
        if node is None:
            ready, why = self.check_ready()
            raise GeminiUnavailable(why if not ready else
                                    "the chat box disappeared")
        how = self._paste_prompt(page, node, prompt)
        page.keyboard.press("Enter")
        self.asked += 1
        LOG.info("      prompt %s (%d characters) and submitted; waiting at "
                 "least %ds", how, len(prompt), GEMINI_FIRST_READ_SECONDS)
        try:
            return self._wait_for_response(page)
        finally:
            # whatever happened, the operator's window goes back to the map
            self._give_the_screen_back()

    def _paste_prompt(self, page, node, prompt):
        """Put the prompt on the clipboard, Ctrl+V it, and check it arrived.

        Typing a 3,000 character prompt key by key is slow and loses
        characters when the editor reformats mid-stream. Pasting is one
        event. The box is read back afterwards, because a paste that
        silently did nothing would send Gemini an empty question.
        """
        node.click()
        time.sleep(CLICK_SETTLE_SECONDS)
        pasted = False
        try:
            page.evaluate("text => navigator.clipboard.writeText(text)",
                          prompt)
            time.sleep(COPY_CLIPBOARD_PAUSE_SECONDS)
            page.keyboard.press("Control+v")
            time.sleep(COPY_CLIPBOARD_PAUSE_SECONDS)
            pasted = True
        except Exception as error:                      # noqa: BLE001
            LOG.debug("      the paste failed (%s)", str(error)[:120])

        if pasted and self._prompt_is_in_the_box(node, prompt):
            return "pasted"

        # the clipboard was refused or the paste did not land: type it
        # instead, and say so rather than sending a half-empty prompt
        LOG.warning("      Ctrl+V did not fill the box - typing the prompt "
                    "instead")
        try:
            node.click()
            page.keyboard.press("Control+a")
            page.keyboard.press("Delete")
        except Exception:                               # noqa: BLE001
            pass
        page.keyboard.insert_text(prompt)
        time.sleep(TYPING_SETTLE_SECONDS)
        if not self._prompt_is_in_the_box(node, prompt):
            raise gemini_input_error(
                "the prompt could not be put into the Gemini box")
        return "typed"

    @staticmethod
    def _prompt_is_in_the_box(node, prompt):
        """True when the box really holds the prompt that was sent."""
        try:
            inside = clean(node.inner_text() or "")
        except Exception:                               # noqa: BLE001
            return False
        if not inside:
            return False
        head = clean(prompt.splitlines()[0])[:40]
        tail = clean(ANSWER_KEYS[-1])
        return (head in inside) and (tail in inside)

    # ------------------------------------------------------------- research
    def research(self, record):
        """``(values, raw)`` for one school. Retries, then gives up cleanly."""
        prompt = build_prompt(record)
        last_error = ""
        for attempt in range(1, GEMINI_RETRIES + 2):
            try:
                raw = self.ask(prompt)
            except (PlaywrightError, PlaywrightTimeout) as error:
                last_error = "browser error: %s" % str(error)[:160]
                self.close()
                time.sleep(RETRY_BACKOFF_SECONDS * attempt)
                continue
            if not raw:
                last_error = "no answer arrived"
                time.sleep(RETRY_BACKOFF_SECONDS * attempt)
                continue
            values = parse_answer(raw)
            if not values:
                last_error = "the answer carried no recognised field"
                time.sleep(RETRY_BACKOFF_SECONDS * attempt)
                continue
            wrong = answer_is_about(record, values)
            if wrong:
                last_error = wrong
                LOG.warning("   attempt %d rejected: %s", attempt, wrong)
                time.sleep(RETRY_BACKOFF_SECONDS * attempt)
                continue
            return values, raw
        raise GeminiUnavailable(last_error or "no usable answer")

# ===========================================================================
#  THE RUN
# ===========================================================================

#: What the status line says about where the run has got to. Set as the
#: run moves; read whenever the line is redrawn.
WHERE = {"city": "", "city_at": 0, "cities": 0,
         "keyword": "", "keyword_at": 0, "keywords": 0}


def say_status(doing):
    """Repaint the status line: where we are, and what is happening now."""
    place = WHERE["city"] or "-"
    if WHERE["cities"]:
        place = "city %d/%d %s" % (WHERE["city_at"], WHERE["cities"], place)
    if WHERE["keyword"]:
        place += " | keyword %d/%d %s" % (WHERE["keyword_at"],
                                          WHERE["keywords"], WHERE["keyword"])
    STATUS.set("%s | %s", place, doing)


# ===========================================================================
#  ONE RUN AT A TIME
# ===========================================================================
#  Two extractors sharing one checkpoint and one master workbook would
#  undo each other: both would research the same schools, both would write
#  the same Excel file, and the one that saved last would win. systemd
#  keeps a service to a single copy, but nothing stops a second run being
#  started by hand in a terminal while the service is working.
#
#  So the writing commands take an exclusive lock first. fcntl.flock() is
#  the right tool because the KERNEL holds it, not a file's contents:
#
#    * it is released the moment the process ends, however it ends - a
#      clean exit, a crash, kill -9, or the power going out. There is no
#      such thing as a stale flock, so nothing has to time out or be
#      cleared by hand after a power cut.
#    * it is advisory between processes on this machine, which is exactly
#      the scope that matters: one project directory, one run.
#
#  The lock file lives in the project directory, so two installations in
#  two different folders - or on two laptops - never collide. It is listed
#  in .gitignore and is not one of recovery_paths(), so it is never
#  committed, pushed or backed up.
#
#  --status, --dry-run, --self-test, --list-keywords and --check-display
#  read or work in a temporary folder, so they are NOT blocked: an
#  operator can always look at a running collection.
# ---------------------------------------------------------------------------

#: Where the lock lives. Named after the project folder it guards.
LOCK_FILE = os.path.join(PROJECT_DIR, ".school_extractor.lock")

#: The open file object. Kept for the whole run: closing it, or the
#: process ending, releases the lock. Never let this be garbage collected.
RUN_LOCK = {"handle": None, "path": ""}


def take_run_lock(path=None):
    """Claim the right to write this project's data. ``(ok, who)``.

    ``ok`` is False only when another live process holds the lock, and
    ``who`` then describes it, as recorded by that process. The lock is
    held until this process ends; nothing has to release it explicitly.
    """
    try:
        import fcntl
    except ImportError:                                 # not Linux
        return True, "fcntl is unavailable - no single-instance protection"

    target = path or LOCK_FILE
    try:
        handle = open(target, "a+", encoding="utf-8")
    except OSError as error:
        # An unwritable directory must not stop a run that would otherwise
        # work; it only means this protection is unavailable.
        return True, "the lock file could not be opened (%s)" % error

    try:
        fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError:
        try:
            handle.seek(0)
            who = handle.read(300).strip()
        except OSError:
            who = ""
        handle.close()
        return False, who or "another process holds the lock"

    try:
        handle.seek(0)
        handle.truncate()
        handle.write("pid %d on %s, started %s\n"
                     % (os.getpid(), PROJECT_DIR,
                        datetime.datetime.now().isoformat(timespec="seconds")))
        handle.flush()
        os.fsync(handle.fileno())
    except OSError:
        pass
    RUN_LOCK["handle"] = handle
    RUN_LOCK["path"] = target
    return True, ""


def release_run_lock():
    """Let the lock go. Rarely needed - ending the process does it too."""
    handle = RUN_LOCK.get("handle")
    if handle is None:
        return False
    try:
        handle.close()                  # closing releases the flock
    except OSError:
        pass
    RUN_LOCK["handle"] = None
    return True


RUNTIME = {"rank_label": "1/1", "stop": False,
           "git_remote": GIT_REMOTE}


def pause(span, why=""):
    """Wait a fresh random time from a (min, max) setting.

    Every delay in the run is a pair of numbers in the settings block and
    a draw from it, so no two waits are identical and all of them can be
    tuned in one place. A stop request skips the wait entirely.
    """
    try:
        low, high = float(span[0]), float(span[1])
    except (TypeError, ValueError, IndexError):
        return 0.0
    if high < low:
        low, high = high, low
    seconds = random.uniform(low, high)
    if seconds <= 0 or RUNTIME["stop"]:
        return 0.0
    if why:
        LOG.info("      waiting %.1fs (%.0f-%.0f) %s", seconds, low, high, why)
    # Taken in slices so Ctrl+C is noticed during the wait rather than
    # after it: a 16-second pause must not swallow a shutdown request.
    deadline = time.monotonic() + seconds
    while True:
        left = deadline - time.monotonic()
        if left <= 0 or RUNTIME["stop"]:
            break
        time.sleep(min(1.0, left))
    return seconds


def _install_stop_handler():
    """A stop request finishes the current step and saves, rather than
    tearing the run down where it stands.

    SIGTERM matters as much as SIGINT here. Ctrl+C sends SIGINT, but
    systemd sends SIGTERM on "systemctl stop", on reboot and on shutdown -
    and with no handler for it Python ends the process immediately,
    abandoning the school being researched. Both now set the same flag,
    so an unattended server shuts down the way an operator's Ctrl+C does.
    """
    names = {getattr(signal, "SIGINT", None): "Ctrl+C",
             getattr(signal, "SIGTERM", None): "SIGTERM (systemd stop)"}

    def handler(signum, _frame):
        label = names.get(signum, "stop request")
        if RUNTIME["stop"]:
            LOG.warning("Second stop request (%s) - stopping immediately.",
                        label)
            raise KeyboardInterrupt
        RUNTIME["stop"] = True
        LOG.warning("")
        LOG.warning("%s: finishing the current school, saving, then "
                    "stopping. Send it again to stop at once.", label)
    for sig in (signal.SIGINT, signal.SIGTERM):
        try:
            signal.signal(sig, handler)
        except (ValueError, OSError, AttributeError):   # not the main thread
            pass


def use_test_folders():
    """Send every output to School_Data/Test_Run, so --test touches no real data."""
    global OUTPUT_DIR, TEMP_DIR, EXCEL_DIR
    global CHECKPOINT_FILE, LOG_FILE, BROWSER_PROFILE_DIR, GEMINI_PROFILE_DIR
    OUTPUT_DIR = os.path.join(PROJECT_DIR, "Test_Run")
    TEMP_DIR = os.path.join(OUTPUT_DIR, "Temp_Data")
    EXCEL_DIR = os.path.join(OUTPUT_DIR, "Excel_Data")
    LOG_FILE = os.path.join(OUTPUT_DIR, "find_school.log")
    CHECKPOINT_FILE = os.path.join(OUTPUT_DIR, "checkpoint.json")
    BROWSER_PROFILE_DIR = os.path.join(OUTPUT_DIR, "browser_profile")
    GEMINI_PROFILE_DIR = os.path.join(OUTPUT_DIR, "browser_profile_gemini")
    ensure_dirs()


#: A school with no way to contact it is not worth a row. Any of these
#: counts: the main or alternate number, the school's own address, or the
#: principal's. Google Maps supplies most phone numbers and Gemini most
#: emails, so the test is applied AFTER research, never before it.
CONTACT_COLUMNS = (
    "School Official Phone Number", "School Alternate Phone Number",
    "Principal Office Phone Number", "School Official Email",
    "School Alternate Email", "Principal Official Email",
)


def has_contact(record):
    """True when the record carries a usable phone number or email."""
    for column in CONTACT_COLUMNS:
        value = clean(record.get(column, ""))
        if not value or is_blank_answer(value):
            continue
        if "Email" in column:
            if tidy_email(value):
                return True
        elif tidy_phone(value):
            return True
    return False


def record_identity(record):
    """A stable string for one school, for the checkpoint's "already done"."""
    keys = match_keys(record)
    if keys:
        return "%s:%s" % keys[0]
    return "name:%s" % entry_identity(record)


def save_gemini_answer(state, city, record, raw):
    """Keep the raw answer, under the school it belongs to."""
    if not KEEP_RAW_GEMINI_ANSWERS or not raw:
        return ""
    path = gemini_school_path(state, city, record)
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "a", encoding="utf-8") as handle:
        handle.write("\n%s\nSCHOOL: %s\nCITY  : %s\nSTATE : %s\nWHEN  : %s\n"
                     "%s\n\n%s\n"
                     % ("=" * 36, clean(record.get("School Name", "")),
                        clean(record.get("City", "")),
                        clean(record.get("State", "")),
                        time.strftime("%Y-%m-%d %H:%M:%S"), "=" * 36,
                        raw.rstrip()))
        handle.flush()
        os.fsync(handle.fileno())
    return path


# ===========================================================================
#  THE OFFICIAL EMAIL AND PHONE NUMBER, FROM GOOGLE SEARCH
# ===========================================================================
#  The senior's rule: the OFFICIAL contact columns come from a Google
#  Search of the school itself, not from Gemini. Gemini keeps the alternate
#  contacts and the other fields.
#
#  The whole difficulty is evidence. A search result page will happily show
#  a phone number belonging to a directory site, another branch, or a
#  coaching centre with a similar name. So a value is written only when it
#  can be tied to THIS school:
#
#    * it is published on a page whose host belongs to the school, or
#    * it is already on the record as Google Maps evidence (the phone), or
#    * its email domain is the school's own website domain.
#
#  A snippet on somebody else's site is not evidence. Nothing is invented,
#  nothing is guessed, and an uncertain field is left BLANK - which the
#  contact rule then treats as "no official contact", not as a failure.
# ---------------------------------------------------------------------------

GOOGLE_SEARCH_URL = "https://www.google.com/search?q=%s&hl=en&gl=in"
OFFICIAL_EMAIL_QUERY = "Official email address of %s in %s"
OFFICIAL_PHONE_QUERY = "Official contact phone number of %s in %s"

#: How firmly a found number is tied to the school. A higher tie replaces a
#: lower one, so the number the phone search itself verifies beats one that
#: merely appeared while searching for the email, which beats a repeat of
#: what Google Maps already showed.
PHONE_TIE_NONE = 0
PHONE_TIE_MAPS_CONFIRMED = 1
PHONE_TIE_EMAIL_SEARCH = 2
PHONE_TIE_PHONE_SEARCH = 3

#: Columns this feature owns. Gemini is not allowed to fill them.
OFFICIAL_EMAIL_COLUMN = "School Official Email"
OFFICIAL_PHONE_COLUMN = "School Official Phone Number"
OFFICIAL_CONTACT_COLUMNS = (OFFICIAL_EMAIL_COLUMN, OFFICIAL_PHONE_COLUMN)
#: Where a value Gemini offers for an official column is put instead.
OFFICIAL_TO_ALTERNATE = {
    OFFICIAL_EMAIL_COLUMN: "School Alternate Email",
    OFFICIAL_PHONE_COLUMN: "School Alternate Phone Number",
}

#: Hosts that list schools but are not schools. A contact found only here
#: is somebody else's data, so it is never promoted to an official field.
DIRECTORY_HOSTS = (
    "justdial", "indiamart", "sulekha", "yellowpages", "indiastudychannel",
    "schoolmykids", "edustoke", "uniapply", "icbse", "schools.org",
    "schooldekho", "collegedunia", "shiksha", "careers360", "glassdoor",
    "facebook", "instagram", "twitter", "x.com", "linkedin", "youtube",
    "wikipedia", "quora", "tripadvisor", "mapquest", "google", "yelp",
    "practo", "magicbricks", "99acres", "olx", "indeed", "naukri",
)

#: Words that say nothing about WHICH school this is, so they are ignored
#: when a host is matched against a school's name.
NAME_NOISE = {
    "the", "school", "schools", "public", "senior", "secondary", "higher",
    "international", "academy", "vidyalaya", "vidya", "convent", "college",
    "high", "primary", "english", "medium", "education", "educational",
    "institute", "institution", "trust", "society", "sr", "sec", "hr",
    "govt", "government", "private", "co", "ed", "campus", "branch",
    "cbse", "icse", "board", "mandir", "nursery", "kids", "play", "centre",
    "center", "and", "of", "for", "in", "at",
}


def name_words(text):
    """The words of a name that actually identify one school."""
    words = re.findall(r"[a-z0-9]+", clean(text).lower())
    return [w for w in words if len(w) > 2 and w not in NAME_NOISE]


def host_of(url):
    """The host of a URL, without www."""
    try:
        host = urllib.parse.urlparse(clean(url)).netloc.lower()
    except ValueError:
        return ""
    host = host.split(":")[0]
    return host[4:] if host.startswith("www.") else host


def is_directory_host(host):
    low = (host or "").lower()
    return any(mark in low for mark in DIRECTORY_HOSTS)


#: A school in India does not publish its contact details on another
#: country's domain. Found by a live run: a school in Sikkim was offered
#: an address on a NEPALI school's domain, accepted only because the two
#: names share one word. The two-letter country code is the cheapest
#: reliable way to refuse that. Generic suffixes are allowed because an
#: Indian school may well use .com, .org or .edu.
FOREIGN_TLD_ALLOWED = {"in", "com", "org", "net", "edu", "gov", "info",
                       "biz", "site", "online", "school", "academy"}


def foreign_country_host(host):
    """True when the host's last label is another country's code."""
    last = (host or "").rsplit(".", 1)[-1].lower()
    if not last or last in FOREIGN_TLD_ALLOWED:
        return False
    # a two-letter suffix that is not "in" is another country
    return len(last) == 2


def host_belongs_to_school(host, record):
    """True when this host looks like the school's OWN site.

    Decided on the distinctive words of the school's name appearing in the
    host, or on the host being the website already recorded for it. A
    directory is never the school's own site, however well its URL matches.
    """
    host = (host or "").lower()
    if not host or is_directory_host(host):
        return False
    if foreign_country_host(host):
        return False
    known = domain_of(record.get("School Website", ""))
    if known and (host == known or host.endswith("." + known)):
        return True
    label = re.sub(r"[^a-z0-9]", "", host.split(".")[0])
    if not label:
        return False
    words = name_words(record.get("School Name", ""))
    if not words:
        return False
    hits = [w for w in words if w in label]
    # one distinctive word is enough for a short host ("dpsrohini.edu.in"),
    # but it must be a real word of the name, not a two-letter fragment
    return bool(hits)


def branch_conflict(record, text):
    """True when the page is plainly about a DIFFERENT branch.

    The record's own branch words must not be contradicted: a page whose
    text carries another branch word from this school's family, and none of
    this record's, is somebody else's campus.
    """
    mine = branch_signature(record)
    if not mine:
        return False
    low = clean(text).lower()
    if any(word.lower() in low for word in mine):
        return False
    other = branch_signature({"School Name": clean(text)[:4000],
                              "School Address": ""})
    return bool(other - {w for w in mine})


def emails_in(text):
    """Every email address in some text, tidied and de-duplicated."""
    found = []
    for raw in _EMAIL.findall(clean(text)):
        address = tidy_email(raw)
        if address and address not in found:
            found.append(address)
    return found


#: A phone number as it appears in result text: spaces, dashes, brackets
#: and a +91 or leading 0 all allowed. Both phones_in() and
#: phone_evidence() use it, so what is extracted is what gets verified.
PHONE_IN_TEXT = re.compile(r"(?:\+91[\s-]*)?(?:0)?\d[\d\s().-]{8,16}\d")


def phones_in(text):
    """Every usable Indian phone number in some text, de-duplicated by key."""
    found, seen = [], set()
    for raw in PHONE_IN_TEXT.findall(clean(text)):
        number = tidy_phone(raw)
        key = phone_key(raw)
        if not number or not key or key in seen:
            continue
        seen.add(key)
        found.append(number)
    return found


def phone_evidence(text, record, phone, window=350):
    """Is THIS number, in this result text, this school's?

    search_evidence() looks for the value as a literal run of characters,
    which works for an email address but not for a phone number: the text
    says "080 4444 9999" while the number in hand has been tidied to
    "08044449999", so the literal search never matched and a number could
    only ever be accepted when Google happened to print it unseparated.

    So the text is searched the way the number was found - by the same
    pattern - and each match whose digits are this number is judged on the
    words around it, exactly as search_evidence() would.
    """
    body = clean(text)
    key = phone_key(phone)
    if not body or not key:
        return False
    words = name_words(record.get("School Name", ""))
    city = clean(record.get("City", "")).lower()
    state = clean(record.get("State", "")).lower()
    low = body.lower()
    for match in PHONE_IN_TEXT.finditer(body):
        if phone_key(match.group(0)) != key:
            continue
        start, end = match.span()
        near = low[max(0, start - window):end + window]
        has_name = any(word in near for word in words) if words else False
        has_place = (city and city in near) or (state and state in near)
        if has_name and has_place and not branch_conflict(record, near):
            return True
    return False


def school_place(record):
    """"Bengaluru Karnataka", with the locality when there is one."""
    parts = [clean(record.get("City", "")), clean(record.get("State", ""))]
    place = " ".join(p for p in parts if p)
    locality = clean(record.get("School Address", ""))
    if locality:
        # the first piece of the address, which is usually the locality
        head = locality.split(",")[0].strip()
        if head and head.lower() not in place.lower():
            place = "%s %s" % (head, place) if place else head
    return place.strip()


def contact_query(template, record):
    return template % (clean(record.get("School Name", "")),
                       school_place(record))


def run_google_search(browser, query):
    """Open one Google search and hand back its text and its result links.

    Returns ``(text, links, why)``. A blocked search returns no results and
    says so - it is reported, never worked around.
    """
    url = GOOGLE_SEARCH_URL % urllib.parse.quote_plus(query)
    page = browser.page()
    browser.goto(page, url)
    time.sleep(PAGE_SETTLE_SECONDS)
    text = browser.text_of(page, "body")
    if looks_blocked(text):
        return "", [], ("Google served a challenge for this search; "
                        "no official contact was taken from it")
    links = []
    try:
        for href in page.eval_on_selector_all(
                "a[href]", "els => els.map(e => e.href)") or []:
            host = host_of(href)
            if not host or host.startswith("google."):
                continue
            if (host, href) not in links:
                links.append((host, href))
    except (PlaywrightError, PlaywrightTimeout, AttributeError):
        # No link list is survivable: the snippet text is still read, and
        # without a verified link nothing is written anyway.
        pass
    return text, links, ""


def search_evidence(text, record, value, window=350):
    """Is this value, in this search result text, this school's?

    Nothing is opened to check - the decision is made on the text Google
    returned, which is all this workflow is allowed to use. A value is
    accepted only where it appears in a stretch of text that ALSO names
    the school and its city or state, because that is what distinguishes
    "this school's number" from one belonging to the next result down.

    A stretch that names a different branch of the same school is
    refused outright.
    """
    body = clean(text)
    if not body or not value:
        return False
    words = name_words(record.get("School Name", ""))
    city = clean(record.get("City", "")).lower()
    state = clean(record.get("State", "")).lower()
    low = body.lower()
    needle = value.lower()
    at = low.find(needle)
    while at != -1:
        near = low[max(0, at - window):at + len(needle) + window]
        has_name = any(word in near for word in words) if words else False
        has_place = (city and city in near) or (state and state in near)
        if has_name and has_place and not branch_conflict(record, near):
            return True
        at = low.find(needle, at + 1)
    return False


def find_official_contacts(director, record):
    """The school's official email and phone, from GOOGLE SEARCH alone.

    TWO SEARCHES, ONE EACH, FOR EVERY SCHOOL: one asking Google for the
    official email, one asking for the official phone number. Both run on
    their own; the phone is never left to whatever the email search
    happened to show. NO SCHOOL WEBSITE IS OPENED: not a homepage, not a
    contact page, nothing. Everything used here comes out of the search
    results text.

    PRIORITY
      1. Google Search, when the result can be tied to this school.
      2. Google Maps - the number already on the record, kept only when
         the search verified none.
    Gemini never fills these two columns. It answers the ALTERNATE email
    and the ALTERNATE phone, and the other fields.

    HOW A RESULT IS TIED TO THE SCHOOL
      email: its domain belongs to the school, OR the address appears in
             text that also names the school and its city or state.
      phone: the number appears in text that names the school and its
             city or state, or it matches what Google Maps showed.
    Either way a stretch of text about a different branch is refused, and
    an address on another country's domain is refused. Anything that
    cannot be tied to the school is left BLANK rather than guessed.

    WHICH NUMBER WINS
    Both searches are read for a phone, so one school can yield more than
    one candidate. The number the PHONE search ties to the school wins;
    below it, one the email search tied to the school; below that, a
    repeat of what Google Maps already showed. A firmer tie found later
    replaces a weaker one found earlier.

    Returns ``{"email", "phone", "email_from", "phone_from", "source",
    "notes", "queries", "phone_tie"}``, where the "_from" values are
    "search", "maps" or "", and "queries" lists the searches that ran.
    """
    found = {"email": "", "phone": "", "email_from": "", "phone_from": "",
             "source": "", "notes": [], "queries": [],
             # how firmly the chosen number is tied to this school
             "phone_tie": PHONE_TIE_NONE}
    if not OFFICIAL_CONTACT_SEARCH:
        return found

    have_email = clean(record.get(OFFICIAL_EMAIL_COLUMN, ""))
    maps_phone_text = clean(record.get(OFFICIAL_PHONE_COLUMN, ""))
    maps_phone = phone_key(maps_phone_text)
    if maps_phone_text:
        found["phone_from"] = "maps"


    opener = getattr(director, "search", None)
    if opener is None:
        found["notes"].append("no search browser available - official "
                              "fields left blank")
        return found
    browser = opener()

    # EVERY SCHOOL GETS ITS OWN PHONE SEARCH.
    #
    # Each query's results are read for both an email and a phone, so the
    # email search sometimes shows a usable number. That is a bonus, never
    # a substitute: the phone search below runs for every school either
    # way, and a number it ties to the school replaces one that merely
    # turned up while looking for the address. Only the email search is
    # skipped, and only when the record already holds an official email.
    queries = [("email", OFFICIAL_EMAIL_QUERY), ("phone", OFFICIAL_PHONE_QUERY)]
    ran = 0
    for _number, (what, template) in enumerate(queries):
        if what == "email" and (have_email or found["email"]):
            continue
        if ran:
            pause(GOOGLE_BETWEEN_QUERIES_DELAY_SECONDS,
                  "between the email and phone searches")
        ran += 1
        found["queries"].append(what)
        query = contact_query(template, record)
        LOG.info("      search: %s", query[:96])
        try:
            text, _links, why = run_google_search(browser, query)
        except (PlaywrightError, PlaywrightTimeout) as error:
            found["notes"].append("search failed: %s" % str(error)[:90])
            break
        if why:
            found["notes"].append(why)
            LOG.warning("      %s", why)
            break
        if not text or branch_conflict(record, text):
            continue

        # --- the official email, from the results text -----------------
        if not found["email"] and not have_email:
            for address in emails_in(text):
                domain = address.split("@")[-1]
                if is_directory_host(domain) or foreign_country_host(domain):
                    continue
                own_domain = host_belongs_to_school(domain, record)
                if own_domain or search_evidence(text, record, address):
                    found["email"] = address
                    found["email_from"] = "search"
                    found["notes"].append(
                        "email verified by %s" % ("the school's own mail "
                                                  "domain" if own_domain
                                                  else "the school name and "
                                                       "place beside it"))
                    break

        # --- the official phone, from the results text -----------------
        # A verified search number is PREFERRED over the Maps one, which
        # is why this can replace a value already on the record.
        # --- the official phone, from the results text -----------------
        # How firmly a number is tied to THIS school decides which wins, so
        # a better-tied one later replaces an earlier one rather than being
        # dropped because a number was already in hand.
        for phone in phones_in(text):
            key = phone_key(phone)
            if not key:
                continue
            if phone_evidence(text, record, phone):
                rank = (PHONE_TIE_PHONE_SEARCH if what == "phone"
                        else PHONE_TIE_EMAIL_SEARCH)
                why = ("phone verified by the school name and place beside "
                       "it in the %s results" % what)
            elif maps_phone and key == maps_phone:
                rank = PHONE_TIE_MAPS_CONFIRMED
                why = ("the number Google Maps showed, confirmed by the "
                       "search")
            else:
                continue
            if rank > found["phone_tie"]:
                found["phone"] = phone
                found["phone_from"] = "search"
                found["phone_tie"] = rank
                found["notes"].append(why)
            if found["phone_tie"] >= PHONE_TIE_PHONE_SEARCH:
                break           # nothing can be tied more firmly

    if not found["email"] and not have_email:
        found["notes"].append("no official email could be verified - left blank")
    if found["phone_from"] == "maps":
        found["notes"].append("no verified search number - keeping the "
                              "Google Maps phone")
    if not found["phone"] and not maps_phone_text:
        found["notes"].append("no official phone could be verified - left blank")
    found["notes"].append("source - email: %s | phone: %s"
                          % (found["email_from"] or "none",
                             found["phone_from"] or "none"))
    return found


def merge_official_contacts(record, found):
    """Write the verified official contacts onto the record, by priority.

    An empty official column is filled by whatever was verified. A column
    that already holds something is replaced ONLY by a value from the
    school's own website, which outranks everything: that is how a number
    Google Maps showed gives way to the one the school publishes itself.

    A search-grade value never replaces an existing one, so a later pass
    cannot quietly downgrade a website contact. Returns how many fields
    were written.
    """
    written = 0
    pairs = ((OFFICIAL_EMAIL_COLUMN, tidy_email(found.get("email", "")),
              found.get("email_from", "")),
             (OFFICIAL_PHONE_COLUMN, tidy_phone(found.get("phone", "")),
              found.get("phone_from", "")))
    for column, value, source in pairs:
        if not value or column not in record:
            continue
        already = clean(record.get(column, ""))
        # Google Search is the highest-priority source now, so a value it
        # verified replaces what is already in the column - which in
        # practice means the number Google Maps showed. Nothing else may
        # overwrite: Maps does not replace a search value, and Gemini
        # never reaches these two columns at all.
        if already and source != "search":
            continue
        if already == value:
            continue
        record[column] = value
        written += 1
    source = clean(found.get("source", ""))
    if source and "School Data Source URL" in record and not clean(
            record.get("School Data Source URL", "")):
        record["School Data Source URL"] = source
        written += 1
    return written


def research_and_save(gemini, checkpoint, state, city, record, position,
                      total, director=None):
    """One school: official contacts, Gemini, validate, merge, save.

    The order matters. The OFFICIAL email and phone are searched for first,
    from Google Search and the school's own website, so those two columns
    are already filled (or deliberately blank) by the time Gemini is
    asked - and apply_answer() never overwrites a column that has a value.
    Gemini then fills the alternate contacts and the other fields.

    The record reaches the master and the cycle file HERE, before the next
    school is started. Nothing waits for the keyword or the city to finish.

    Returns "saved", "duplicate", "no contact", or "failed".
    """
    name = clean(record.get("School Name", ""))[:52]
    LOG.info("")
    LOG.info("   School %d/%d - %s", position, total, name)
    say_status("researching school %d/%d - %s" % (position, total, name[:28]))

    if OFFICIAL_CONTACT_SEARCH and director is not None:
        LOG.info("      [OFFICIAL CONTACTS] searching Google")
        say_status("official contacts %d/%d - %s" % (position, total,
                                                     name[:26]))
        try:
            found = find_official_contacts(director, record)
        except AccessBlocked:
            raise
        except (PlaywrightError, PlaywrightTimeout) as error:
            # A failed search must not cost the school its research: the
            # official columns simply stay blank and Gemini still runs.
            LOG.warning("      the contact search failed: %s",
                        str(error)[:120])
            found = {"email": "", "phone": "", "source": "",
                     "notes": ["the search could not be completed"]}
        written = merge_official_contacts(record, found)
        for note in found.get("notes", [])[:3]:
            LOG.info("      %s", note)
        LOG.info("      official email %s | official phone %s | %d field(s) "
                 "written",
                 clean(record.get(OFFICIAL_EMAIL_COLUMN, "")) or "blank",
                 clean(record.get(OFFICIAL_PHONE_COLUMN, "")) or "blank",
                 written)
        pause(GOOGLE_TO_GEMINI_DELAY_SECONDS,
              "after the contact research, before Gemini")

    if gemini is not None:
        started = time.time()
        try:
            values, raw = gemini.research(record)
        except GeminiUnavailable as error:
            LOG.warning("      no usable answer: %s", str(error)[:140])
            checkpoint.record_failure(record, error)
            return "failed"
        except (PlaywrightError, PlaywrightTimeout) as error:
            LOG.warning("      browser problem: %s", str(error)[:140])
            checkpoint.record_failure(record, error)
            return "failed"
        kept = save_gemini_answer(state, city, record, raw)
        LOG.info("      answer complete after %.0fs, %d characters -> %s",
                 time.time() - started, len(raw or ""),
                 os.path.basename(kept) if kept else "not kept")
        filled = apply_answer(record, values)
        LOG.info("      %d of 57 field(s) read from the answer, %d written "
                 "into the record", len(values), filled)

    if not has_contact(record):
        # Fully researched, nothing to contact it by. That is a finished
        # school, not a failure: it is not retried, it does not keep its
        # city open, and it never reaches the final workbooks.
        LOG.info("      no phone number and no email - not saved")
        checkpoint.mark_enriched(state, city, record_identity(record))
        clear_school_temp(state, city, record)
        return "no contact"

    LOG.info("      [EXCEL SAVING]")
    written, duplicates = append_records([record])
    checkpoint.mark_enriched(state, city, record_identity(record))

    # "saved" means read back off the disk, not "the write did not raise"
    on_disk = not missing_from_master([record])
    if not on_disk:
        LOG.error("      NOT SAVED - %s is not in the master after the write",
                  name)
        checkpoint.record_failure(record, "the row did not reach the master")
        return "failed"
    if written:
        checkpoint.note_committed(state, city, record_identity(record))
        LOG.info("      saved and verified in %s",
                 os.path.basename(master_path()))
        # Verified on disk a moment ago, so its raw answer is spent.
        clear_school_temp(state, city, record)
        return "saved"
    LOG.info("      already in %s - not saved twice",
             os.path.basename(master_path()))
    clear_school_temp(state, city, record)
    return "duplicate"


def stored_discovery(state, city, keyword):
    """The schools this keyword already copied from Maps, if they survive.

    A keyword that was interrupted part way through its research keeps its
    capture file - cleanup only removes it once the keyword COMPLETES - so
    the whole discovered list is still on disk after a crash or a power
    cut. Reading it back costs a moment; re-copying it from Google Maps
    costs the whole scroll again (a Delhi keyword took 38 copies and over
    three minutes), and asks Google for the same thing twice.

    Returns ``(records, complete, path)``. ``complete`` is True only when
    the saved text contains Google's own end-of-list message, which is the
    same test the live collection uses - a half-copied list is not reused,
    because schools further down were never reached.

    A missing, empty or unparsable file returns no records, and the caller
    then collects from Maps as it always did. Nothing already saved is at
    risk either way: the master workbook and the city's working sheet are
    separate from this.
    """
    base = keyword_txt_path(state, city, keyword)
    for path in (base, base + ".gz"):
        if not os.path.exists(path):
            continue
        try:
            text = read_capture(path)
        except (OSError, EOFError, gzip.BadGzipFile) as error:
            LOG.warning("   the saved list for %r could not be read (%s) - "
                        "collecting it again", keyword, str(error)[:80])
            continue
        if not text.strip():
            continue
        low = text.lower()
        complete = any(marker in low for marker in END_OF_LIST_MARKERS)
        try:
            records = records_from_txt(path, state, city)
        except Exception as error:                      # noqa: BLE001
            LOG.warning("   the saved list for %r is unreadable (%s) - "
                        "collecting it again", keyword, str(error)[:80])
            continue
        if records and complete:
            return records, True, path
        if records and not complete:
            LOG.info("   a saved list for %r exists but never reached the "
                     "end of the list - collecting it again", keyword)
    return [], False, ""


def process_keyword(director, gemini, checkpoint, state, city, keyword,
                    position, total, limit=None):
    """One keyword, from the Maps search to the last school saved.

    The order is fixed and does not interleave with any other keyword:

        Maps search -> Ctrl+A/Ctrl+C until the end of the list -> .txt
        -> regex extraction -> temporary workbook
        -> for each school: Gemini -> raw .txt -> regex -> merge
           -> master duplicate check -> final Excel -> checkpoint
        -> keyword complete
    """
    report = {"keyword": keyword, "found": 0, "saved": 0, "duplicates": 0,
              "failed": 0, "no_contact": 0, "complete": False, "reason": ""}
    LOG.info("")
    LOG.info("[KEYWORD %d of %d] %s in %s", position, total, keyword, city)
    WHERE.update({"keyword": keyword, "keyword_at": position,
                  "keywords": total})
    say_status("opening Google Maps")

    # RESUME WITHOUT ASKING GOOGLE AGAIN.
    # If this keyword already copied its whole list and was interrupted
    # during the research, the list is still on disk. Reuse it rather than
    # scrolling Maps a second time for the same schools.
    saved_records, saved_complete, saved_path = stored_discovery(
        state, city, keyword)
    if saved_complete and saved_records:
        LOG.info("")
        LOG.info("[MAPS SKIPPED] reusing the saved list for %s in %s - %d "
                 "school(s) from %s", keyword, city, len(saved_records),
                 os.path.basename(saved_path))
        say_status("reusing %d saved school(s)" % len(saved_records))
        found = SearchResult(build_query(keyword, city))
        found.records = saved_records
        found.complete = True
        found.txt_path = saved_path
        found.reason = "reused the saved list (no Maps search needed)"
    else:
        LOG.info("")
        LOG.info("[MAPS COPYING] %s in %s", keyword, city)
        browser = director.maps()
        found = collect_search(browser, keyword, state, city, limit=limit)
    report["reason"] = found.reason
    if not found.records:
        # A keyword with no results is finished, not failed - as long as the
        # list really reached its end. Treating it as unresolved would keep
        # the city open for ever: a small town simply has no ICSE school.
        report["complete"] = bool(found.complete)
        if report["complete"]:
            checkpoint.mark_keyword_done(state, city, keyword)
            LOG.info("   KEYWORD COMPLETE: %s in %s - the list was empty",
                     keyword, city)
            sweep_completed_keyword(state, city, keyword)
        else:
            LOG.warning("   no school came out of this keyword, and the list "
                        "never reached its end - it will be tried again")
        return report

    # the city's working sheet: this keyword's schools merged into whatever
    # earlier keywords already found
    # the copying is finished and the text is on disk, so the Maps browser
    # has nothing left to do: it goes before anything else is opened
    ok, detail = director.close()
    if not ok:
        report["reason"] = "the Maps browser would not close: %s" % detail
        LOG.error("   %s - nothing else will be opened", report["reason"])
        return report

    LOG.info("")
    LOG.info("[REGEX EXTRACTION] reading %s",
             os.path.basename(found.txt_path))
    # the schools are out of the text now, so it can be stored packed
    pack_capture_file(found.txt_path)
    index = SchoolIndex()
    for record in load_temp_excel(state, city):
        index.add(record)
    before = len(index)
    new, merged = merge_into_index(index, found.records)
    save_temp_excel(state, city, index.records)
    report["found"] = len(found.records)
    LOG.info("")
    LOG.info("   extracted %d school(s) from the copied text", len(found.records))
    LOG.info("   new in this city %d | already known %d | city total %d",
             new, merged, len(index))
    LOG.info("   temporary workbook updated: %s",
             temp_excel_path(state, city))

    # only this keyword's schools are researched now, and only the ones that
    # are not already done or already in the master
    # Every school in the city's sheet that is still unresearched - not
    # only the ones this keyword added. A keyword interrupted half way
    # leaves discovered-but-unresearched rows behind, and slicing to this
    # keyword's own additions abandoned them silently.
    done = checkpoint.enriched_keys(state, city)
    identities, known = master_identities()
    todo, already = [], 0
    for record in index.records:
        if record_identity(record) in done:
            already += 1
            continue
        if (stable_identity(record) in identities
                or set(match_keys(record)) & known):
            report["duplicates"] += 1
            continue
        todo.append(record)
    left_over = len(todo) - max(len(index.records) - before, 0)
    if before and left_over > 0:
        LOG.info("   including %d school(s) left unresearched by an earlier "
                 "run", left_over)
    # --limit researches only the first few schools. Whatever it leaves
    # behind means the keyword is NOT finished, so it must not be marked
    # complete and its saved list must not be swept away. Collection used
    # to stop early under a limit, which hid this; now that a saved list
    # is reused, the limit is the only thing cutting the work short.
    cut_by_limit = False
    if limit:
        cut_by_limit = len(todo) > limit
        todo = todo[:limit]

    LOG.info("")
    LOG.info("[GEMINI RESEARCH] %s in %s", keyword, city)
    # The line an operator reads after a restart: how much of this
    # keyword is already behind us, and which school is next.
    settled = already + report["duplicates"]
    LOG.info("   %s in %s: %d discovered, %d completed, %d pending",
             keyword, city, len(index.records), settled, len(todo))
    if settled and todo:
        LOG.info("   resuming the saved %s list: %d completed, %d pending - "
                 "starting with %s", city, settled, len(todo),
                 clean(todo[0].get("School Name", ""))[:48])
    LOG.info("   %d school(s) to research, %d already done, %d already in "
             "the master", len(todo), already, report["duplicates"])

    if todo and gemini is not None:
        # the first time Gemini is needed in this run: the Maps browser is
        # already closed, so this is the only browser there is
        ready, why = gemini.check_ready()
        LOG.info("   Gemini: %s", why)
        if not ready:
            report["reason"] = "Gemini is not available: %s" % why
            LOG.error("   %s", report["reason"])
            director.close()
            return report

    for number, record in enumerate(todo, start=1):
        if RUNTIME["stop"]:
            report["reason"] = "stopped at your request"
            LOG.warning("   stopping before school %d at your request", number)
            return report
        outcome = research_and_save(gemini, checkpoint, state, city, record,
                                    number, len(todo), director=director)
        if outcome == "saved":
            report["saved"] += 1
        elif outcome == "duplicate":
            report["duplicates"] += 1
        elif outcome == "no contact":
            report["no_contact"] += 1
        else:
            report["failed"] += 1
        save_temp_excel(state, city, index.records)
        # A safe point: this school is saved and the city's sheet is
        # written, so a break here costs nothing and repeats nothing.
        if not schedule_gate(director, checkpoint):
            report["reason"] = "stopped at your request"
            return report
        git_sync("school saved in %s" % city, checkpoint)
        if number < len(todo):
            # ONE wait between two schools, not two. When the contact
            # search runs, the next thing the next school does is a Google
            # search, so its own between-schools setting is the one that
            # applies; without the search the shorter general pause does.
            if OFFICIAL_CONTACT_SEARCH:
                pause(GOOGLE_BETWEEN_SCHOOLS_DELAY_SECONDS,
                      "before the next school's Google contact research")
            else:
                pause(SCHOOL_DELAY_SECONDS, "before the next school")

    # this keyword's schools are done, so the Gemini browser goes too. The
    # next keyword opens Maps again, into an empty screen.
    if todo:
        ok, detail = director.close()
        if not ok:
            LOG.error("   [BROWSER] the Gemini browser would not close: %s",
                      detail)
    report["complete"] = (bool(found.complete) and not report["failed"]
                          and not cut_by_limit)
    if cut_by_limit:
        report["reason"] = ("--limit stopped after %d school(s); the keyword "
                            "stays open and its saved list is kept" % limit)
        LOG.info("   %s", report["reason"])
    LOG.info("")
    if report["complete"]:
        checkpoint.mark_keyword_done(state, city, keyword)
        LOG.info("   KEYWORD COMPLETE: %s in %s - %d saved, %d duplicate(s)"
                 "%s", keyword, city, report["saved"], report["duplicates"],
                 ", %d with no phone or email" % report["no_contact"]
                 if report["no_contact"] else "")
        # Every school of this keyword is resolved, so its copied text and
        # any answer left by an interrupted earlier run can go. Reached
        # only here: an unresolved keyword keeps the lot.
        git_sync("keyword %r finished in %s" % (keyword, city), checkpoint,
                 force=True)
        done = checkpoint.enriched_keys(state, city)
        gone = sweep_completed_keyword(
            state, city, keyword,
            resolved=[r for r in index.records
                      if record_identity(r) in done])
        if gone:
            LOG.info("   cleaned up: %s", ", ".join(gone[:4])
                     + (" and %d more" % (len(gone) - 4) if len(gone) > 4
                        else ""))
    else:
        # Say which of the three reasons it was. A reused complete capture
        # cut short by --limit used to print a bare dash and no reason.
        why = []
        if not found.complete:
            why.append("the list never reached its end")
        elif report.get("reason"):
            why.append(report["reason"])
        if report["failed"]:
            why.append("%d school(s) in the retry queue" % report["failed"])
        LOG.warning("   KEYWORD UNRESOLVED: %s in %s%s", keyword, city,
                    " - " + "; ".join(why) if why else "")
    return report


def process_city(director, gemini, checkpoint, state, city, keywords,
                 limit=None, all_keywords=None):
    """One visit to a city: the keywords given, each finished completely.

    ``keywords`` is what this visit works on. ``all_keywords`` is every
    keyword the city needs before it can be called complete - the same
    list when a city is done in one visit, and the full enabled set when
    the run is working one keyword at a time across many cities. A city is
    closed, and its working files removed, only when every one of
    ``all_keywords`` is finished, whichever order they were done in.
    """
    LOG.info("")
    LOG.info("=" * 72)
    LOG.info("CITY: %s, %s", city, state)
    LOG.info("=" * 72)
    checkpoint.start_city(state, city)
    report = {"state": state, "city": city, "found": 0, "saved": 0,
              "duplicates": 0, "failed": 0, "no_contact": 0,
              "complete": False, "note": ""}

    unresolved = []
    for position, keyword in enumerate(keywords, start=1):
        if RUNTIME["stop"]:
            report["note"] = "stopped at your request inside %s" % city
            break
        if checkpoint.keyword_done(state, city, keyword):
            LOG.info("")
            LOG.info("[KEYWORD %d of %d] %s - already complete, skipped",
                     position, len(keywords), keyword)
            continue
        checkpoint.note_keyword(keyword)
        try:
            done = process_keyword(director, gemini, checkpoint, state,
                                   city, keyword, position, len(keywords),
                                   limit=limit)
        except AccessBlocked:
            raise
        except (PlaywrightError, PlaywrightTimeout) as error:
            LOG.error("   the browser failed on %r: %s", keyword,
                      str(error)[:160])
            director.close()
            unresolved.append(keyword)
            continue
        report["found"] += done["found"]
        report["saved"] += done["saved"]
        report["duplicates"] += done["duplicates"]
        report["failed"] += done["failed"]
        report["no_contact"] += done["no_contact"]
        if not done["complete"]:
            unresolved.append(keyword)

    if unresolved:
        report["note"] = ("%d keyword(s) unresolved (%s) - %s stays open"
                          % (len(unresolved), ", ".join(unresolved[:3]), city))
        LOG.warning("   %s", report["note"])
        return report
    if RUNTIME["stop"]:
        return report

    # This visit's keywords are done, but the city may still owe another
    # one from a later pass. Closing it now would clear the working files
    # that the next keyword's duplicate check needs, and would mark a city
    # complete that has not been searched for Government Schools yet.
    expected = list(all_keywords or keywords)
    waiting = [k for k in expected
               if not checkpoint.keyword_done(state, city, k)]
    if waiting:
        report["note"] = ("%d of %d keyword(s) still to come for %s (%s) - "
                          "it stays open until they are done"
                          % (len(waiting), len(expected), city,
                             ", ".join(waiting[:3])))
        LOG.info("   %s", report["note"])
        # This visit's rows are already in the Excel files, so they belong
        # in the totals now - the visit that finally closes the city will
        # only count its own.
        checkpoint.add_totals(report["saved"], report["duplicates"])
        return report

    # every keyword finished: the city's rows must be in the master before
    # it is closed and its working files removed
    records = load_temp_excel(state, city)
    without_contact = [r for r in records if not has_contact(r)]
    missing = missing_from_master(records)

    # Every keyword says it is finished, yet the sheet still holds schools
    # that never reached the master - an earlier run was interrupted between
    # discovering them and researching them. The keyword flags mean those
    # rows would never be looked at again, so they are finished here before
    # the city is allowed to close.
    if missing and not RUNTIME["stop"]:
        LOG.info("")
        LOG.info("[FINISHING OFF] %d school(s) were discovered but never "
                 "researched", len(missing))
        for number, record in enumerate(missing, start=1):
            if RUNTIME["stop"]:
                break
            outcome = research_and_save(gemini, checkpoint, state, city,
                                        record, number, len(missing),
                                        director=director)
            if outcome == "saved":
                report["saved"] += 1
            elif outcome == "duplicate":
                report["duplicates"] += 1
            elif outcome == "no contact":
                report["no_contact"] += 1
            else:
                report["failed"] += 1
            save_temp_excel(state, city, records)
        missing = missing_from_master(records)

    if missing:
        report["note"] = ("%d school(s) are not in the master yet - %s stays "
                          "open and its files are kept" % (len(missing), city))
        LOG.error("   %s", report["note"])
        return report

    director.close()            # nothing is left open between cities
    clear_city_temp(state, city)
    checkpoint.mark_city_complete(state, city, report["saved"],
                                  report["duplicates"])
    git_sync("city %s complete" % city, checkpoint, force=True)
    report["complete"] = True
    LOG.info("   CITY COMPLETE: %s - %d saved, %d duplicate(s), %d with no "
             "phone or email, across %d keyword(s)", city, report["saved"],
             report["duplicates"], len(without_contact), len(keywords))
    # One permanent line per city, so a quiet run still leaves a history
    # on the terminal behind the live status line.
    if terminal_is_quiet():
        STATUS.clear()
        announce("  done: %-22s %3d saved, %3d duplicate(s), %3d with no "
                 "phone or email", "%s, %s" % (city, state),
                 report["saved"], report["duplicates"], len(without_contact))
    return report


def run(args):
    """The whole collection, city by city."""
    value_rank, value_num, source = resolve_rank(args.rank, args.num_sys)
    try:
        validate_rank(value_rank, value_num)
    except RankError as error:
        LOG.error("%s", error)
        return 2
    RUNTIME["rank_label"] = "%d/%d" % (value_rank, value_num)
    # The RESOLVED rank, not the constant at the top of the file. The
    # per-rank branch name is built from this, and a laptop that takes its
    # rank from rank.json, SCHOOL_RANK or --rank would otherwise push to
    # school-data-rank-1 and overwrite rank 1's data.
    RUNTIME["rank"] = value_rank
    RUNTIME["num_sys"] = value_num
    # resolved once per run, so every sync in this run pushes to the same
    # remote even if rank.json is edited while it works
    RUNTIME["git_remote"], RUNTIME["git_remote_from"] = resolve_git_remote()
    set_rank_paths(value_rank)
    reconcile_master_and_cycles()

    keywords = enabled_keywords()
    if not keywords:
        LOG.error("Every search keyword is disabled - nothing to search for. "
                  "Set at least one \"enabled\": True in "
                  "ENABLED_SCHOOL_KEYWORDS.")
        return 2

    plan = search_plan(value_rank, value_num,
                       [args.state] if args.state else None,
                       [args.city] if args.city else None)
    if not plan:
        LOG.error("Nothing to do: no city of rank %d of %d matched %s",
                  value_rank, value_num,
                  "--city %s" % args.city if args.city else
                  "--state %s" % args.state if args.state else "the plan")
        return 2

    pairs = list(dict.fromkeys((state, city) for state, city, _k in plan))
    checkpoint = Checkpoint()
    changed = checkpoint.check_plan(value_rank, value_num,
                                    dataset_fingerprint())
    if changed and not args.ignore_plan_change:
        LOG.error("The saved progress does not match this run: %s.", changed)
        LOG.error("Nothing has been changed. Re-run with "
                  "--ignore-plan-change to continue anyway, or restore the "
                  "previous rank / keyword settings.")
        return 3

    # The enabled keywords changed but the cities did not. Nothing already
    # collected is affected - a saved school is the same school whichever
    # keyword found it - so the new plan is adopted and every record kept.
    stored_print = checkpoint.data.get("plan_fingerprint") or ""
    if stored_print and stored_print != dataset_fingerprint() \
            and checkpoint.keywords_only_changed():
        done_cities = len(checkpoint.data.get("completed_cities") or [])
        LOG.warning("The enabled keywords changed since the saved progress "
                    "was written. The cities are unchanged, so the new list "
                    "has been adopted and nothing was cleared.")
        LOG.warning("   keywords now: %s", ", ".join(keywords))
        LOG.warning("   kept: %d saved school(s), %d completed city(ies)",
                    checkpoint.data.get("saved", 0), done_cities)
        if done_cities:
            # Said plainly rather than silently: a city closed under the
            # old list is not reopened, so a keyword added now will never
            # be searched there unless the operator asks for it.
            LOG.warning("   NOTE: a city already complete is not searched "
                        "again, so a newly added keyword will not run "
                        "there. Use --reset-city <name> for any you want "
                        "collected again.")
        checkpoint.migrate_keyword_plan(value_rank, value_num,
                                        dataset_fingerprint())

    checkpoint.adopt_plan(value_rank, value_num, dataset_fingerprint())
    checkpoint.save()

    # Leftovers from finished work, removed only after the checkpoint and
    # the master both confirm the city really is done. An open city keeps
    # every file it has.
    sweep_completed_cities(checkpoint)

    pending = [p for p in pairs if not checkpoint.is_city_complete(*p)]
    if not pending:
        LOG.info("Every selected city is already complete. Use --reset-city "
                 "to collect one again.")
        return 0

    LOG.info("")
    LOG.info("=" * 49)
    LOG.info("SCHOOL DATA EXTRACTOR")
    LOG.info("Rank: %d / %d   (from %s)", value_rank, value_num, source)
    LOG.info("Cities for this rank: %d   still to do: %d",
             len(cities_for_rank(value_rank, value_num)), len(pending))
    LOG.info("Writing to: %s", OUTPUT_DIR)
    LOG.info("=" * 49)
    LOG.info("Keywords (%d): %s", len(keywords), ", ".join(keywords))
    LOG.info("Cities to process (%d of this rank's %d): %s", len(pending),
             len(cities_for_rank(value_rank, value_num)),
             ", ".join("%s (%s)" % (c, s) for s, c in pending[:12])
             + (" ..." if len(pending) > 12 else ""))
    LOG.info("Searches queued: %d", len(pending) * len(keywords))
    LOG.info("Official contacts: %s",
             "Google Search, one email and one phone query per school"
             if OFFICIAL_CONTACT_SEARCH else "off (Gemini only)")
    # the clock starts here, so the first period is drawn before any work
    reset_work_timer()
    LOG.info("Work / break      : %s", work_break_status())
    LOG.info("Night stop        : %s", night_pause_status())
    LOG.info("GitHub sync       : %s", git_status_line())
    LOG.info("Cycle backups     : %s", "Backup/ (verified copies)"
             if BACKUP_COMPLETED_CYCLES else "off")
    # Without --show the lines above went to the log file only, so say in
    # one line that the run is going and where to watch it. With --show
    # the whole log is on the screen already and saying this would be
    # both pointless and untrue.
    if terminal_is_quiet():
        announce("Working through %d city(ies), %d search(es)%s.",
                 len(pending), len(pending) * len(keywords),
                 ", one keyword at a time across all of them"
                 if KEYWORD_FIRST_ORDER and len(keywords) > 1 else "")
        announce("   below        : what it is doing right now, on one line")
        announce("   every step   : tail -f %s", LOG_FILE)
        announce("   to watch it  : run again with --show")

    _install_stop_handler()
    exit_code = 0
    reports = []
    director = gemini = None
    try:
        director = BrowserDirector(show=getattr(args, "show", False),
                                   mode="headless" if args.headless else None)
        if GEMINI_ENABLED and not args.no_gemini:
            # built, not started: Gemini must not open while Maps is being
            # copied, and that includes any start-up readiness check. The
            # first time it opens is when the first school needs it.
            gemini = Gemini(director)
            LOG.info("Gemini research: on (its browser opens after the first "
                     "keyword's copying is finished)")
        else:
            LOG.info("Gemini research is off - Maps discovery only")
        LOG.info("The browser is %s", where_the_window_goes(
            getattr(args, "show", False), args.headless))
        LOG.info("One browser at a time: Maps while copying, Gemini while "
                 "researching")

        # The order of the work, decided in one place. Each visit is one
        # city and the keyword(s) to do there now.
        #   keyword-first: Private Schools across every city, then
        #                  Government Schools across every city
        #   city-first   : every keyword of one city, then the next city
        # Both do the same searches and save the same rows.
        if KEYWORD_FIRST_ORDER and len(keywords) > 1:
            visits = [(state, city, [keyword])
                      for keyword in keywords for state, city in pending]
        else:
            visits = [(state, city, list(keywords)) for state, city in pending]

        for position, (state, city, doing) in enumerate(visits, start=1):
            if RUNTIME["stop"]:
                LOG.warning("Stopped before %s at your request.", city)
                break
            # The safest point there is: nothing of this visit has started.
            if not schedule_gate(director, checkpoint):
                LOG.warning("Stopped before %s at your request.", city)
                break
            # A city finished by an earlier pass needs no further visit.
            if checkpoint.is_city_complete(state, city):
                continue
            LOG.info("")
            LOG.info("=" * 49)
            LOG.info("Visit %d of %d: %s, %s%s", position, len(visits), city,
                     state,
                     "" if len(doing) == len(keywords)
                     else "  [%s]" % ", ".join(doing))
            WHERE.update({"city": city, "city_at": position,
                          "cities": len(visits), "keyword": "",
                          "keyword_at": 0, "keywords": 0})
            say_status("starting")
            LOG.info("=" * 49)
            try:
                report = process_city(director, gemini, checkpoint, state,
                                      city, doing, limit=args.limit,
                                      all_keywords=keywords)
            except BrowserUnavailable as error:
                LOG.error("%s", error)
                exit_code = 4
                break
            except AccessBlocked as error:
                LOG.error("%s", error)
                exit_code = 5
                break
            reports.append(report)
            if not report["complete"]:
                exit_code = exit_code or 1
            if position < len(visits) and not RUNTIME["stop"]:
                pause(CITY_DELAY_SECONDS, "before the next city")
    except DesktopUnavailable as error:
        # --show could not reach the desktop: say so plainly instead of
        # drawing into the invisible screen or printing a traceback
        for line in str(error).splitlines():
            LOG.error("%s", line)
        exit_code = 6
    except BrowserUnavailable as error:
        LOG.error("%s", error)
        exit_code = 4
    except KeyboardInterrupt:
        LOG.warning("Interrupted. Progress is saved.")
        exit_code = exit_code or 130
    finally:
        if gemini is not None:
            gemini.close()
        if director is not None:
            director.shutdown()
        checkpoint.save()

    # Graceful shutdown: whatever has been done is pushed before the
    # process ends, so a laptop that never comes back loses nothing.
    git_sync("run finished", checkpoint, force=True)
    print_run_summary(reports, checkpoint)
    return exit_code


def print_run_summary(reports, checkpoint):
    STATUS.done()
    announce("")
    announce("=" * 72)
    announce("RUN SUMMARY")
    announce("=" * 72)
    for report in reports:
        announce("  %-24s found %-5d saved %-5d dup %-4d no-contact %-4d "
                 "failed %-4d %s",
                 "%s, %s" % (report["city"], report["state"]),
                 report["found"], report["saved"], report["duplicates"],
                 report.get("no_contact", 0), report["failed"],
                 "complete" if report["complete"] else "OPEN")
        if report["note"]:
            announce("        %s", report["note"])
    announce("  cities complete overall : %d",
             len(checkpoint.data["completed_cities"]))
    announce("  schools saved overall   : %d", checkpoint.data["saved"])
    announce("  in the retry queue      : %d", len(checkpoint.data["failed"]))
    announce("  %-28s %d record(s)", os.path.basename(master_path()),
             row_count(master_path()))
    for path in all_cycle_paths():
        announce("  %-28s %d record(s)", os.path.basename(path),
                 row_count(path))


def retry_failed(args):
    """Research the schools that failed before, then save them."""
    value_rank, value_num, _source = resolve_rank(args.rank, args.num_sys)
    RUNTIME["rank_label"] = "%d/%d" % (value_rank, value_num)
    RUNTIME["rank"] = value_rank
    RUNTIME["num_sys"] = value_num
    # resolved once per run, so every sync in this run pushes to the same
    # remote even if rank.json is edited while it works
    RUNTIME["git_remote"], RUNTIME["git_remote_from"] = resolve_git_remote()
    set_rank_paths(value_rank)
    reconcile_master_and_cycles()
    checkpoint = Checkpoint()
    failures = checkpoint.take_failures()
    if not failures:
        LOG.info("The retry queue is empty.")
        return 0
    LOG.info("Retrying %d school(s)", len(failures))
    _install_stop_handler()
    director = gemini = None
    still_failed, done = [], []
    try:
        director = BrowserDirector(show=getattr(args, "show", False),
                                   mode="headless" if args.headless else None)
        gemini = Gemini(director)
        ready, why = gemini.check_ready()
        LOG.info("Gemini: %s", why)
        if not ready:
            for item in failures:
                checkpoint.record_failure(item["record"], "not retried: " + why)
            LOG.error("Gemini is not available - the queue is unchanged.")
            return 4
        for number, item in enumerate(failures, start=1):
            if RUNTIME["stop"]:
                still_failed.extend(f["record"] for f in failures[number - 1:])
                break
            record = item["record"]
            LOG.info("[%d/%d] %s", number, len(failures),
                     clean(record.get("School Name", ""))[:60])
            try:
                values, _raw = gemini.research(record)
            except (GeminiUnavailable, PlaywrightError) as error:
                LOG.warning("   still failing: %s", str(error)[:160])
                still_failed.append(record)
                continue
            apply_answer(record, values)
            done.append(record)
    except BrowserUnavailable as error:
        for line in str(error).splitlines():
            LOG.error("%s", line)
        for item in failures:
            checkpoint.record_failure(item["record"], "not retried")
        return 6 if isinstance(error, DesktopUnavailable) else 4
    finally:
        if gemini is not None:
            gemini.close()
        if director is not None:
            director.shutdown()
    if done:
        append_records(done)
    for record in still_failed:
        checkpoint.record_failure(record, "retry did not succeed")
    LOG.info("Retried %d, saved %d, still failing %d", len(failures),
             len(done), len(still_failed))
    return 0 if not still_failed else 1


# ===========================================================================
#  REPORTING MODES
# ===========================================================================

def print_status(args):
    value_rank, value_num, source = resolve_rank(args.rank, args.num_sys)
    try:
        validate_rank(value_rank, value_num)
    except RankError as error:
        print("RANK ERROR: %s" % error)
        return 2
    set_rank_paths(value_rank)
    checkpoint = Checkpoint()
    try:
        mine = cities_for_rank(value_rank, value_num)
    except RankError as error:
        print("RANK ERROR: %s" % error)
        return 2
    keywords = enabled_keywords()
    completed = checkpoint.data["completed_cities"]
    pending = [p for p in mine if not checkpoint.is_city_complete(*p)]

    print("")
    print("School Data Extractor - status")
    print("-" * 72)
    print("  rank                : %d of %d (from %s)" % (value_rank,
                                                          value_num, source))
    print("  cities in dataset   : %d unique (%d groups, %d listed entries)"
          % (len(ALL_CITIES), len(CITY_BY_STATES),
             sum(len(v) for v in CITY_BY_STATES.values())))
    print("  cities for this rank: %d" % len(mine))
    print("  first cities        : %s"
          % ", ".join("%s (%s)" % (c, s) for s, c in mine[:6]))
    print("  cities from         : %s" % SHARED_SOURCE["cities"])
    print("  keywords from       : %s" % SHARED_SOURCE["keywords"])
    print("  57 fields from      : %s" % SHARED_SOURCE["fields"])
    print("  search keywords     : %d of %d enabled - %s"
          % (len(keywords), len(SCHOOL_KEYWORDS), ", ".join(keywords)))
    print("  searches for a run  : %d" % (len(pending) * len(keywords)))
    print("  cities completed    : %d" % len(completed))
    print("  cities still to do  : %d" % len(pending))
    print("  current city        : %s" % (checkpoint.data["current_city"]
                                          or "none"))
    print("  current keyword     : %s" % (checkpoint.data["current_keyword"]
                                          or "none"))
    keywords_done = sum(len(v) for v in
                        checkpoint.data.get("keywords_done", {}).values())
    print("  keywords completed  : %d (in cities not yet closed)"
          % keywords_done)
    print("  schools saved       : %d" % checkpoint.data["saved"])
    print("  duplicates merged   : %d" % checkpoint.data["duplicates"])
    print("  retry queue         : %d" % len(checkpoint.data["failed"]))
    print("  plan fingerprint    : %s" % (dataset_fingerprint()[:16]))
    print("  checkpoint updated  : %s" % (checkpoint.data["updated"] or "never"))
    print("  cycle size          : %d records per file" % CYCLE_SIZE)
    print("  official contacts   : %s"
          % ("Google Search - an email query and a phone query per school"
             if OFFICIAL_CONTACT_SEARCH else "off (Gemini only)"))
    print("  work / break        : %s" % work_break_status())
    print("  night stop          : %s" % night_pause_status())
    print("  github sync         : %s" % git_status_line())
    saved_git = checkpoint.data.get("git") or {}
    if saved_git.get("last_ok") or saved_git.get("last_error"):
        print("    last pushed       : %s" % (saved_git.get("last_ok")
                                              or "never"))
        if saved_git.get("pending"):
            print("    NOT yet pushed    : %s"
                  % (saved_git.get("last_error") or "a sync is outstanding"))
    saved_work = checkpoint.data.get("work") or {}
    if saved_work.get("break_until"):
        print("  a break was running : until %s (%s)"
              % (saved_work.get("break_until"),
                 saved_work.get("reason", "")[:50]))
    backups = sorted(glob.glob(os.path.join(BACKUP_DIR, "*.xlsx"))) \
        if os.path.isdir(BACKUP_DIR) else []
    print("  cycle backups       : %d in Backup/" % len(backups))
    master = row_count(master_path())
    print("  %-20s: %d record(s)" % (MASTER_BASENAME + ".xlsx", master))
    paths = all_cycle_paths()
    if paths:
        for path in paths:
            print("  %-20s: %d record(s)" % (os.path.basename(path),
                                             row_count(path)))
    else:
        print("  cycle files         : none yet")
    if os.path.exists(journal_path()):
        print("  UNFINISHED SAVE     : %s is present; the next run repairs it"
              % COMMIT_JOURNAL)
    print("-" * 72)
    if checkpoint.data["failed"]:
        print("  retry queue (first 5):")
        for item in checkpoint.data["failed"][:5]:
            print("     %-40s %s" % (item["school"][:40], item["reason"][:40]))
    return 0


def print_dry_run(args):
    """The plan, with no browser and no network request of any kind."""
    value_rank, value_num, source = resolve_rank(args.rank, args.num_sys)
    try:
        validate_rank(value_rank, value_num)
    except RankError as error:
        print("RANK ERROR: %s" % error)
        return 2
    plan = search_plan(value_rank, value_num,
                       [args.state] if args.state else None,
                       [args.city] if args.city else None)
    pairs = list(dict.fromkeys((s, c) for s, c, _k in plan))
    print("")
    print("DRY RUN - no browser is opened and no request is made")
    print("-" * 72)
    print("rank %d of %d (from %s)" % (value_rank, value_num, source))
    print("%d cities x %d keywords = %d searches"
          % (len(pairs), len(enabled_keywords()), len(plan)))
    words = enabled_keywords()
    first = KEYWORD_FIRST_ORDER and len(words) > 1
    print("order: %s" % ("one keyword across every city, then the next"
                         if first else "every keyword of one city, then the "
                                       "next city"))
    print("-" * 72)
    if first:
        # exactly the order run() will work in
        for keyword in words:
            print("%s - all %d city(ies)" % (keyword, len(pairs)))
            for state, city in pairs:
                query = build_query(keyword, city)
                print("    %-46s %s" % (query, maps_search_url(query)))
    else:
        for state, city in pairs:
            print("%s, %s" % (city, state))
            for keyword in words:
                query = build_query(keyword, city)
                print("    %-46s %s" % (query, maps_search_url(query)))
    print("-" * 72)
    print("%d searches in total" % len(plan))
    return 0


def print_display_report():
    """Say exactly which screen --show would use, and why. Opens nothing."""
    print("")
    print("Display check")
    print("-" * 72)
    print("  this run's DISPLAY   : %s" % (os.environ.get("DISPLAY")
                                           or "not set"))
    print("  this run's XAUTHORITY: %s" % (os.environ.get("XAUTHORITY")
                                           or "not set"))
    print("  started by xvfb-run  : %s" % ("yes" if running_under_xvfb()
                                           else "no"))
    try:
        display, authority, how = find_desktop_display()
    except DesktopUnavailable as error:
        print("-" * 72)
        print("--show would FAIL:")
        print("")
        print(str(error))
        return 4
    print("  --show would use     : %s" % display)
    print("  with XAUTHORITY      : %s" % (authority or "none"))
    print("  found by             : %s" % how)
    windows = windows_on_display(display, authority)
    print("  windows on it now    : %d" % len(windows))
    for title in windows[:5]:
        print("      %s" % title[:66])
    print("-" * 72)
    return 0


def print_keywords():
    on = enabled_keywords()
    print("")
    print("School search keywords (%d, %d enabled)"
          % (len(SCHOOL_KEYWORDS), len(on)))
    print("  from: %s" % SHARED_SOURCE["keywords"])
    print("-" * 72)
    for number, name in enumerate(SCHOOL_KEYWORDS, start=1):
        print("  %2d  %-3s %-34s %s" % (number, "on" if name in on else "",
                                        name, build_query(name, "Bengaluru")))
    print("-" * 72)
    print("  %d enabled x 1 city = %d Google Maps searches per city"
          % (len(on), len(on)))
    return 0


# ===========================================================================
#  SELF TEST - offline, no browser, no network
# ===========================================================================

def self_test():
    """Every check that can be made without a live website."""
    checks, failures = 0, []

    def ok(label, condition, detail=""):
        nonlocal checks
        checks += 1
        if not condition:
            failures.append("%s%s" % (label, " - " + detail if detail else ""))

    # --- 1. the embedded city dataset --------------------------------------
    ok("29 city groups", len(CITY_BY_STATES) == 29,
       "got %d" % len(CITY_BY_STATES))
    ok("905 listed city entries",
       sum(len(v) for v in CITY_BY_STATES.values()) == 905,
       "got %d" % sum(len(v) for v in CITY_BY_STATES.values()))
    ok("866 unique cities after exact de-duplication",
       len(ALL_CITIES) == 866, "got %d" % len(ALL_CITIES))
    ok("no city appears twice", len(ALL_CITIES) == len(set(ALL_CITIES)))
    ok("the first city keeps its first group",
       CITY_TO_STATE.get("Delhi") == "Delhi NCR")
    flat = [c for names in CITY_BY_STATES.values() for c in names]
    ok("dataset order is preserved",
       ALL_CITIES == list(dict.fromkeys(flat)))

    # --- 2. the rank algorithm ---------------------------------------------
    ok("rank 0 gives nothing", get_cities_by_rank(ALL_CITIES, 0, 25) == [])
    ok("rank above num_sys gives nothing",
       get_cities_by_rank(ALL_CITIES, 26, 25) == [])
    ok("num_sys below 1 gives nothing",
       get_cities_by_rank(ALL_CITIES, 1, 0) == [])
    for systems in (1, 2, 3, 7, 10, 25, 100):
        pieces = [get_cities_by_rank(ALL_CITIES, r, systems)
                  for r in range(1, systems + 1)]
        rejoined = [city for piece in pieces for city in piece]
        ok("num_sys=%d covers every city exactly once" % systems,
           rejoined == ALL_CITIES)
        sizes = {len(p) for p in pieces}
        ok("num_sys=%d slices differ by at most one" % systems,
           max(sizes) - min(sizes) <= 1)
    ok("num_sys=25 assigns every city",
       sum(len(get_cities_by_rank(ALL_CITIES, r, 25)) for r in range(1, 26))
       == len(ALL_CITIES))
    ok("each slice is contiguous",
       get_cities_by_rank(ALL_CITIES, 2, 25)[0]
       == ALL_CITIES[len(get_cities_by_rank(ALL_CITIES, 1, 25))])
    for bad_rank, bad_num in ((0, 25), (26, 25), (-1, 25), (1, 0)):
        try:
            validate_rank(bad_rank, bad_num)
            ok("rank %s of %s is rejected" % (bad_rank, bad_num), False)
        except RankError:
            ok("rank %s of %s is rejected" % (bad_rank, bad_num), True)

    # --- 3. keywords --------------------------------------------------------
    ok("the shared lists loaded from mylib.py", not shared_lists_problem(),
       SHARED_SOURCE["cities"])
    # However many keywords mylib.py lists - the senior's two, or all 38 -
    # the run needs at least one and every enabled name must exist.
    ok("the keyword list is loaded", len(SCHOOL_KEYWORDS) >= 1,
       "got %d" % len(SCHOOL_KEYWORDS))
    ok("every enabled keyword exists in the shared list",
       not unknown_enabled_keywords(), str(unknown_enabled_keywords()))
    ok("at least one keyword is enabled", len(enabled_keywords()) >= 1)
    SCHOOL_WORDS = ("school", "vidyalaya", "mandir", "academy", "gurukul",
                    "madrasa", "convent", "sainik", "navy", "boarding")
    ok("every keyword names a kind of school",
       all(any(word in name.lower() for word in SCHOOL_WORDS)
           for name in SCHOOL_KEYWORDS),
       str([n for n in SCHOOL_KEYWORDS
            if not any(w in n.lower() for w in SCHOOL_WORDS)]))
    JOB_WORDS = ("developer", "engineer", "intern", "fresher", "b.tech",
                 "btech", "mba", "bca", "mca", "hiring", "resume", "job",
                 "recruit", "student", "graduate", "analyst", "manager")
    # whole words: "intern" must not match "International Schools"
    job_pattern = re.compile(r"\b(%s)\b" % "|".join(
        w.replace(".", r"\.") for w in JOB_WORDS), re.I)
    ok("no keyword is a student or job keyword",
       not [n for n in SCHOOL_KEYWORDS if job_pattern.search(n)],
       str([n for n in SCHOOL_KEYWORDS if job_pattern.search(n)]))
    ok("keyword names are unique",
       len(set(SCHOOL_KEYWORDS)) == len(SCHOOL_KEYWORDS))
    ok("the query reads '<keyword> in <city>'",
       build_query("Private Schools", "Bengaluru")
       == "Private Schools in Bengaluru")
    ok("the URL is the typed-search URL",
       maps_search_url("Private Schools in Bengaluru").startswith(
           "https://www.google.com/maps/search/Private+Schools+in+Bengaluru"))

    # --- 4. the 57 fields ---------------------------------------------------
    ok("57 columns", len(SCHOOL_COLUMNS) == 57, "got %d" % len(SCHOOL_COLUMNS))
    ok("57 answer keys", len(ANSWER_KEYS) == 57)
    ok("column names are unique", len(set(SCHOOL_COLUMNS)) == 57)
    ok("answer keys are unique", len(set(ANSWER_KEYS)) == 57)
    ok("School_Data_Source_URL is the source field",
       "School_Data_Source_URL" in ANSWER_KEYS)
    ok("Source_URLs is a spelling of it, not a 58th column",
       "Source URLs" not in SCHOOL_COLUMNS
       and KEY_SYNONYMS.get("Source_URLs") == "School_Data_Source_URL")
    # The Excel file has 57 columns; the PROMPT asks for the ones Gemini
    # is responsible for. Google Search owns the two official contacts,
    # so those are deliberately left out of the prompt - asking as well
    # would duplicate the work and invite an unverified second opinion.
    asked = len(gemini_asked_columns())
    ok("the prompt asks for every field Gemini is responsible for",
       requested_fields_text().count("\n") == asked - 1,
       "%d asked" % asked)
    ok("the answer format lists the same fields",
       answer_format_text().count("\n") == asked - 1)
    ok("exactly the two official contacts are left out",
       sorted(gemini_skips_columns()) == ["School Official Email",
                                          "School Official Phone Number"]
       if OFFICIAL_CONTACT_SEARCH else asked == 57,
       sorted(gemini_skips_columns()))
    ok("a blank record still has all 57 columns", len(blank_record()) == 57)

    # --- 5. the parser ------------------------------------------------------
    answer = ("School_Name: Gupta Public School\n"
              "School_Official_Email: office@gupta.edu.in\n"
              "UDISE_Code: NO_DATA\n"
              "Affiliation_Board: CBSE\n"
              "PIN_Code: 560011\n"
              "Additional_School_Notes: line one\nstill the same note\n")
    parsed = parse_answer(answer)
    ok("a plain answer parses",
       parsed.get("School Name") == "Gupta Public School")
    ok("NO_DATA is not stored", "UDISE Code" not in parsed)
    ok("a multi-line value survives",
       "still the same note" in parsed.get("Additional School Notes", ""))
    glued = "School_Name: ABC SchoolPrincipal_Name: R. Rao"
    ok("two fields on one line are separated",
       parse_answer(glued).get("School Name") == "ABC School"
       and parse_answer(glued).get("Principal Name") == "R. Rao")
    ok("Source_URLs lands in School Data Source URL",
       parse_answer("Source_URLs: https://a.example.com").get(
           "School Data Source URL") == "https://a.example.com")
    ok("unrelated text parses to nothing",
       parse_answer("Sure! Here is what I found about schools.") == {})
    echoed = (answer_format_text() + "\nSchool_Name: Real School\n"
              "Affiliation_Board: ICSE\n")
    echo_parsed = parse_answer(echoed)
    ok("the echoed prompt never reaches a column",
       echo_parsed.get("School Name") == "Real School")
    ok("a value of exactly 'Value' is treated as blank",
       parse_answer("School_Name: Value") == {})

    # --- 6. duplicate rules -------------------------------------------------
    def school(name, **extra):
        record = blank_record()
        record.update({"School Name": name, "City": "Bengaluru",
                       "State": "Karnataka"})
        record.update(extra)
        return record

    index = SchoolIndex()
    index.add(school("Gupta Public School",
                     **{"School Address": "12 MG Road, Bengaluru",
                        "PIN Code": "560001"}))
    index.add(school("Gupta Public School",
                     **{"School Address": "12 MG Road, Bengaluru",
                        "PIN Code": "560001",
                        "School Official Phone Number": "9900112233"}))
    ok("the same school found twice is one row", len(index) == 1)
    ok("the merge filled the blank",
       index.records[0]["School Official Phone Number"] == "9900112233")

    branches = SchoolIndex()
    branches.add(school("ORCHIDS The International School - Sarjapur Road",
                        **{"School Website": "https://orchids.edu.in",
                           "PIN Code": "560035",
                           "School Address": "Sarjapur Road, Bengaluru"}))
    branches.add(school("ORCHIDS The International School - Hennur Road",
                        **{"School Website": "https://orchids.edu.in",
                           "PIN Code": "560043",
                           "School Address": "Hennur Road, Bengaluru"}))
    ok("two branches of one chain stay two rows", len(branches) == 2)

    similar = SchoolIndex()
    similar.add(school("St Mary's School",
                       **{"School Address": "5 Church Street, Bengaluru"}))
    similar.add(school("St Mary's School",
                       **{"School Address": "88 Hosur Road, Bengaluru"}))
    ok("a similar name with a different address is not merged",
       len(similar) == 2)

    places = SchoolIndex()
    link = "https://www.google.com/maps/place/X/data=!1s0x3bae16:0x7f00a1"
    places.add(school("Alpha School", **{"School Google Maps Link": link}))
    places.add(school("Alpha School Main Campus",
                      **{"School Google Maps Link": link}))
    ok("the same Maps place is one school", len(places) == 1)

    # --- 7. value formatting -------------------------------------------------
    ok("a 10 digit phone is kept", tidy_phone("99001 12233") == "9900112233")
    ok("a PIN is not a phone number", tidy_phone("560011") == "")
    ok("an address number is not a phone number",
       tidy_phone("PID No.29-105-415/1") == "")
    ok("a PIN is found in an address",
       tidy_pin("12 MG Road, Bengaluru 560001") == "560001")
    ok("an affiliation number is not a PIN",
       tidy_pin("ORCHIDS School - CBSE - Affiliation No. 831330") == "")
    ok("a UDISE code is not a PIN",
       tidy_pin("UDISE Code: 292001") == "")
    ok("the middle of a phone number is not a PIN",
       tidy_pin("Call 9900112233") == "")
    ok("a real PIN still wins over a labelled number",
       tidy_pin("Affiliation No. 831330, Hennur Road, Bengaluru 560043")
       == "560043")
    ok("an email is lower-cased", tidy_email("Office@School.IN")
       == "office@school.in")
    ok("a bare domain becomes a URL",
       tidy_url("school.edu.in") == "https://school.edu.in")
    ok("a leading-zero number stays text",
       "School Official Phone Number" in TEXT_COLUMNS
       and "PIN Code" in TEXT_COLUMNS)

    # --- 8. Maps evidence is not overwritten ---------------------------------
    observed = school("Real Name", **{"School Official Phone Number": "9900112233",
                                      "School Address": "12 MG Road"})
    apply_answer(observed, {"School Name": "Something Else",
                            "School Official Phone Number": "080 1111 2222",
                            "Principal Name": "Mrs R Rao"})
    ok("Gemini does not overwrite an observed phone number",
       observed["School Official Phone Number"] == "9900112233")
    ok("Gemini does not rename an observed school",
       observed["School Name"] == "Real Name")
    ok("Gemini fills a field Maps did not have",
       observed["Principal Name"] == "Mrs R Rao")

    # --- 9. the wrong school is rejected --------------------------------------
    asked = school("Gupta Public School", **{"PIN Code": "560011"})
    ok("an answer about another school is rejected",
       bool(answer_is_about(asked, {"School Name": "Delhi Model Academy"})))
    ok("an answer with the wrong PIN is rejected",
       bool(answer_is_about(asked, {"School Name": "Gupta Public School",
                                    "PIN Code": "110001"})))
    ok("the right answer is accepted",
       answer_is_about(asked, {"School Name": "Gupta Public School",
                               "PIN Code": "560011"}) == "")

    # --- 10. the prompt ------------------------------------------------------
    prompt = build_prompt(school("Gupta Public School",
                                 **{"School Address": "12 MG Road"}))
    ok("the prompt names the school", "Gupta Public School" in prompt)
    ok("the prompt carries the known address", "12 MG Road" in prompt)
    ok("the prompt carries every key Gemini is asked for",
       all(("%s: Value" % key) in prompt for key in gemini_asked_keys()))
    ok("and none of the keys Google Search owns",
       not any(("%s: Value" % key) in prompt
               for key in ("School_Official_Email",
                           "School_Official_Phone_Number"))
       if OFFICIAL_CONTACT_SEARCH else True)
    ok("the prompt forbids inventing a UDISE code",
       "Never invent a UDISE Code" in prompt)
    ok("a missing value is asked for as NO_DATA", NO_DATA in prompt)

    # --- 11. this file is self-contained -------------------------------------
    source = ""
    try:
        with open(os.path.abspath(__file__), encoding="utf-8") as handle:
            source = handle.read()
    except OSError:
        pass
    if source:
        banned = ("new_mylib", "mylib", "config", "school_source",
                  "gemini_controller", "source_file", "school_gemini_prompt")
        imports = re.findall(r"^\s*(?:from|import)\s+([A-Za-z_][\w.]*)",
                             source, re.M)
        offenders = [name for name in imports
                     if any(name.split(".")[0] == bad for bad in banned)]
        ok("no project file is imported", not offenders, str(offenders))
        ok("the city data is embedded",
           "CITY_BY_STATES = {" in source)
        ok("both prompts are embedded",
           "GEMINI_PROMPT_TEMPLATE" in source and "Write in this format"
           in source)

    # --- 12. checkpoint behaviour ---------------------------------------------
    folder = tempfile.mkdtemp(prefix="school_selftest_")
    try:
        point = Checkpoint(os.path.join(folder, "checkpoint.json"))
        ok("a fresh checkpoint has no completed city",
           point.data["completed_cities"] == [])
        point.adopt_plan(1, 25, "abc")
        point.save()
        again = Checkpoint(os.path.join(folder, "checkpoint.json"))
        ok("the checkpoint survives a restart", again.data["rank"] == 1)
        ok("a changed rank is detected",
           bool(again.check_plan(2, 25, "abc")))
        # A CHANGED CITY LIST still stops the run. Changing the stored
        # city fingerprint is what makes this a city change rather than a
        # keyword change - passing a different plan fingerprint alone is
        # a keyword change, which is migrated instead of refused.
        again.data["city_fingerprint"] = "a-different-city-list"
        ok("a changed city list is detected",
           bool(again.check_plan(1, 25, "different")))
        again.data["city_fingerprint"] = city_fingerprint()
        ok("changing only the keywords is migrated, not refused",
           again.check_plan(1, 25, "different") == ""
           and again.keywords_only_changed())
        ok("an unchanged plan resumes", again.check_plan(1, 25, "abc") == "")
        again.mark_discovery_done("Karnataka", "Bengaluru")
        ok("discovery is remembered",
           Checkpoint(again.path).discovery_done("Karnataka", "Bengaluru"))
        again.mark_city_complete("Karnataka", "Bengaluru", 5, 1)
        ok("a completed city is remembered",
           Checkpoint(again.path).is_city_complete("Karnataka", "Bengaluru"))
        failed_record = school("Failed School")
        again.record_failure(failed_record, "no answer")
        ok("a failure is retryable",
           len(Checkpoint(again.path).data["failed"]) == 1)
        ok("the retry queue can be emptied",
           len(again.take_failures()) == 1
           and Checkpoint(again.path).data["failed"] == [])
    finally:
        shutil.rmtree(folder, ignore_errors=True)

    # --- 13. Excel, in a temporary folder --------------------------------------
    if OPENPYXL_AVAILABLE:
        failures.extend(_excel_self_test())
        checks += 14
    else:
        failures.append("openpyxl is not installed, so the Excel checks did "
                        "not run")

    print("")
    print("SELF TEST")
    print("-" * 72)
    if failures:
        for line in failures:
            print("  FAILED  %s" % line)
        print("-" * 72)
        print("%d of %d checks failed" % (len(failures), checks))
        return 1
    print("  %d checks passed - city data, rank division, keywords, the 57" % checks)
    print("  fields, the parser, duplicate rules, value formatting, the")
    print("  prompt, self-containment, the checkpoint and Excel.")
    print("-" * 72)
    return 0


def _excel_self_test():
    """The master file, the 100-record cycles and the crash repair."""
    # BACKUP_DIR is redirected too, and that is not optional: this section
    # finalises cycles at CYCLE_SIZE = 3, and finalise_current_cycle() now
    # takes a permanent backup. Without this the self-test wrote three-row
    # fixture workbooks straight into the project's real Backup/ folder
    # every single time it ran.
    # GIT_SYNC_ENABLED is turned off here for the same reason BACKUP_DIR
    # is redirected: this section finalises cycles, finalise_current_cycle()
    # calls git_sync(), and an offline self-test must not commit anything -
    # it was creating a real commit on the rank branch holding the real
    # checkpoint and working sheet.
    global EXCEL_DIR, CYCLE_SIZE, BACKUP_DIR, GIT_SYNC_ENABLED
    kept_dir, kept_size = EXCEL_DIR, CYCLE_SIZE
    kept_backup = BACKUP_DIR
    kept_sync = GIT_SYNC_ENABLED
    kept_level = LOG.level
    folder = tempfile.mkdtemp(prefix="school_excel_")
    problems = []
    try:
        LOG.setLevel(logging.CRITICAL)
        EXCEL_DIR, CYCLE_SIZE = folder, 3
        BACKUP_DIR = os.path.join(folder, "Backup")
        GIT_SYNC_ENABLED = False

        def school(number):
            record = blank_record()
            record.update({"School Name": "School %d" % number,
                           "City": "Bengaluru", "State": "Karnataka",
                           "PIN Code": "5600%02d" % number,
                           "School Address": "%d Some Road" % number})
            return record

        written, _skipped = append_records([school(n) for n in range(1, 8)])
        if written != 7:
            problems.append("Excel: wrote %d of 7 records" % written)
        if row_count(master_path()) != 7:
            problems.append("Excel: the master holds %d of 7 records"
                            % row_count(master_path()))
        if finished_cycle_numbers() != [1, 2]:
            problems.append("Excel: finished cycles are %s, expected [1, 2]"
                            % finished_cycle_numbers())
        for number in (1, 2):
            if row_count(cycle_path(number)) != 3:
                problems.append("Excel: cycle%d holds %d records, not 3"
                                % (number, row_count(cycle_path(number))))
        if row_count(cycle_path(0)) != 1:
            problems.append("Excel: the active file holds %d records, not 1"
                            % row_count(cycle_path(0)))

        before = [row_count(cycle_path(n)) for n in (1, 2)]
        again, skipped = append_records([school(n) for n in range(1, 8)])
        if again or skipped != 7:
            problems.append("Excel: re-saving wrote %d and skipped %d"
                            % (again, skipped))
        if [row_count(cycle_path(n)) for n in (1, 2)] != before:
            problems.append("Excel: a finished cycle file was rewritten")

        rows = [r for p in all_cycle_paths() for r in read_records(p)]
        if len(rows) != 7:
            problems.append("Excel: the cycles hold %d records, not 7"
                            % len(rows))
        if rows and list(rows[0].keys()) != list(SCHOOL_COLUMNS):
            problems.append("Excel: the columns changed on the way to disk")

        # a crash between the master write and the cycle write
        write_records(master_path(), read_records(master_path())
                      + [school(99)])
        report = reconcile_master_and_cycles()
        if report["added_to_current"] != 1:
            problems.append("Excel: the repair added %d records to the "
                            "current cycle, not 1" % report["added_to_current"])
        if row_count(master_path()) != 8:
            problems.append("Excel: the repair changed the master count to %d"
                            % row_count(master_path()))

        # and the other way round
        write_records(cycle_path(0), read_records(cycle_path(0))
                      + [school(98)])
        report = reconcile_master_and_cycles()
        if report["added_to_master"] != 1:
            problems.append("Excel: the repair added %d records to the "
                            "master, not 1" % report["added_to_master"])

        damaged = cycle_path(0)
        with open(damaged, "wb") as handle:
            handle.write(b"this is not a workbook")
        book = load_workbook_safely(damaged)
        if [c.value for c in book.active[1]] != list(SCHOOL_COLUMNS):
            problems.append("Excel: a damaged workbook was not replaced")
        if not any(".damaged_" in name for name in os.listdir(folder)):
            problems.append("Excel: the damaged workbook was not kept")
    except Exception as error:                          # noqa: BLE001
        problems.append("Excel: %s" % str(error)[:200])
    finally:
        EXCEL_DIR, CYCLE_SIZE = kept_dir, kept_size
        BACKUP_DIR = kept_backup
        GIT_SYNC_ENABLED = kept_sync
        LOG.setLevel(kept_level)
        shutil.rmtree(folder, ignore_errors=True)
    return problems


# ===========================================================================
#  COMMAND LINE
# ===========================================================================

def master_row_count():
    """Rows in the master workbook, or None when there is no workbook yet."""
    path = master_path()
    if not os.path.exists(path):
        return None
    try:
        book = openpyxl.load_workbook(path, read_only=True)
        sheet = book.active
        rows = max(sheet.max_row - 1, 0)
        book.close()
        return rows
    except Exception:                                   # noqa: BLE001
        return None


def migrate_plan_command(args):
    """Move a checkpoint to a new laptop count, or refuse and say why.

    Splitting the work across more laptops re-cuts the city list, so a
    rank's slice changes. This writes the new rank and laptop count into
    the checkpoint ONLY when every city the checkpoint has touched still
    belongs to this rank; otherwise that city's saved rows would sit in a
    copy that no longer searches it. Nothing else in the checkpoint is
    altered, and a refusal changes nothing at all.
    """
    value_rank, value_num, source = resolve_rank(args.rank, args.num_sys)
    set_rank_paths(value_rank)
    checkpoint = Checkpoint()
    stored_rank = checkpoint.data.get("rank")
    stored_num = checkpoint.data.get("num_sys")

    LOG.info("")
    LOG.info("CHECKPOINT PLAN MIGRATION")
    LOG.info("   file          : %s", CHECKPOINT_FILE)
    LOG.info("   saved under   : rank %s of %s", stored_rank, stored_num)
    LOG.info("   this run wants: rank %s of %s (from %s)",
             value_rank, value_num, source)

    if not os.path.exists(CHECKPOINT_FILE):
        LOG.info("   there is no checkpoint here yet - nothing to migrate")
        return 0
    if (stored_rank, stored_num) == (value_rank, value_num):
        LOG.info("   already on this plan - nothing to migrate")
        return 0
    if stored_rank != value_rank:
        LOG.error("   REFUSED: this checkpoint was written by rank %s, not "
                  "rank %s. Migrating it would mix two ranks' work.",
                  stored_rank, value_rank)
        return 3

    touched = checkpoint.cities_with_progress()
    LOG.info("   cities with saved progress: %d", len(touched))
    ok, lines = checkpoint.plan_migration(value_rank, value_num)
    for line in lines:
        LOG.info("%s", line)
    if not ok:
        LOG.error("   REFUSED: a city with saved progress changes rank under "
                  "%s laptops. Its records would be stranded.", value_num)
        LOG.error("   Nothing was written. Keep this rank on %s laptops, or "
                  "move that city's rows by hand first.", stored_num)
        return 3

    kept = {
        "rows in the master workbook": master_row_count(),
        "committed keys": sum(len(v) for v in
                              (checkpoint.data.get("committed") or {}).values()),
        "enriched keys": sum(len(v) for v in
                             (checkpoint.data.get("enriched") or {}).values()),
        "completed cities": len(checkpoint.data.get("completed_cities") or []),
        "retry queue": len(checkpoint.data.get("failed") or []),
    }
    checkpoint.migrate_plan(value_rank, value_num, dataset_fingerprint())
    again = Checkpoint()
    LOG.info("   MIGRATED to rank %s of %s", again.data.get("rank"),
             again.data.get("num_sys"))
    for label, value in kept.items():
        if value is not None:
            LOG.info("      kept %-24s %s", label + ":", value)
    LOG.info("   cities this rank now searches: %d",
             len(get_cities_by_rank(ALL_CITIES, value_rank, value_num)))
    return 0


def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Collect Indian school data city by city.",
        epilog="Install:  python3 -m pip install playwright openpyxl  &&  "
               "python3 -m playwright install chromium")
    parser.add_argument("--status", action="store_true",
                        help="show rank, cities, progress and files, then exit")
    parser.add_argument("--dry-run", action="store_true",
                        help="show the search plan without opening a browser")
    parser.add_argument("--self-test", action="store_true",
                        help="run the offline checks (no network) and exit")
    parser.add_argument("--list-keywords", action="store_true",
                        help="show every school search keyword")
    parser.add_argument("--retry-failed", action="store_true",
                        help="research the schools in the retry queue")
    parser.add_argument("--state", help="only cities of this state")
    parser.add_argument("--city", help="only this city")
    parser.add_argument("--limit", type=int,
                        help="stop after this many schools per city (testing)")
    parser.add_argument("--test", action="store_true",
                        help="write everything to School_Data/Test_Run instead "
                             "of the real folders")
    parser.add_argument("--rank", type=int, help="override the rank")
    parser.add_argument("--num-sys", type=int, help="override the laptop count")
    parser.add_argument("--no-gemini", action="store_true",
                        help="Maps discovery only, no research")
    parser.add_argument("--headless", action="store_true",
                        help="never draw a window, not even under --show")
    parser.add_argument("--show", action="store_true",
                        help="put the Chromium window on your real desktop, "
                             "even when the command starts with xvfb-run")
    parser.add_argument("--check-display", action="store_true",
                        help="report which screen --show would use, then exit")
    parser.add_argument("--reset-city", metavar="CITY",
                        help="forget one city's progress so it runs again")
    parser.add_argument("--migrate-plan", action="store_true",
                        help="move saved progress to this rank's new laptop "
                             "count, only if no touched city changes rank")
    parser.add_argument("--ignore-plan-change", action="store_true",
                        help="resume although the rank or keyword list changed")
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="log every step")
    parser.add_argument("--quiet", "-q", action="store_true",
                        help="log warnings and errors only")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    if args.test:
        use_test_folders()
    # The file keeps everything; the terminal shows the steps only when
    # you ask to watch the process, with --show or --verbose.
    setup_logging(
        level="DEBUG" if args.verbose else None,
        screen=("DEBUG" if args.verbose
                else "ERROR" if args.quiet
                else "INFO" if getattr(args, "show", False)
                else "WARNING"))

    problem = shared_lists_problem()
    if problem and not args.self_test:
        for line in problem.splitlines():
            LOG.error("%s", line)
        return 2
    if args.headless and args.show:
        LOG.error("--headless and --show ask for opposite things. Use one.")
        return 2
    if args.check_display:
        return print_display_report()
    if args.self_test:
        return self_test()
    if args.list_keywords:
        return print_keywords()
    if args.dry_run:
        return print_dry_run(args)
    if args.status:
        return print_status(args)
    if args.migrate_plan:
        return migrate_plan_command(args)
    if args.reset_city:
        set_rank_paths(resolve_rank(args.rank, args.num_sys)[0])
        checkpoint = Checkpoint()
        state = CITY_TO_STATE.get(args.reset_city, args.state or "")
        checkpoint.forget_city(state, args.reset_city)
        LOG.info("%s has been reset. Nothing already saved to Excel was "
                 "removed.", args.reset_city)
        return 0
    # From here on the command WRITES this project's data, so it must be
    # the only one doing so. The read-only commands above are already
    # past, deliberately: looking at a running collection is always
    # allowed. --test writes to its own folder but shares the browser
    # profiles and the lock keeps that honest too.
    taken, who = take_run_lock()
    if not taken:
        LOG.error("Another School Data Extractor instance is already "
                  "running.")
        if who:
            LOG.error("   it is: %s", who)
        LOG.error("   nothing was read or written. Wait for it to finish, "
                  "or stop it first.")
        return 7
    if who:
        LOG.warning("   %s", who)

    if args.retry_failed:
        return retry_failed(args)
    if args.test and not (args.city or args.state):
        LOG.error("--test needs --city or --state, so a test run stays small.")
        return 2
    return run(args)


if __name__ == "__main__":
    sys.exit(main())
