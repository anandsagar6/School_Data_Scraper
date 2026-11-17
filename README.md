# School Data Extractor

One standalone Python file, [`find_school.py`](find_school.py), that collects
Indian school data city by city from Google Maps, researches each school with
the Gemini web application, and writes the approved 57-column Excel dataset.

**Copy that one file to another laptop, install two packages, change `rank`,
and it runs.** It imports nothing from this folder.

---

## Where the shared data comes from

Three things live in `mylib.py`, not in the script:

```
mylib.py
    city_by_states            29 groups, 905 entries, 866 unique
    SCHOOL_SEARCH_KEYWORDS    38 keywords
    SCHOOL_FIELDS             the 57 columns, as (heading, answer key)
```

Edit any of them there and the next run picks it up. There is **no second
copy in find_school.py**, so nothing can drift out of step.

`--status` always says where each one came from:

```text
cities from         : mylib.py (city_by_states)
keywords from       : mylib.py (SCHOOL_SEARCH_KEYWORDS)
57 fields from      : mylib.py (SCHOOL_FIELDS)
```

The `mylib.py` **beside the script** is the only file looked in, so this
folder is self-contained: a laptop needs those two files in one folder and
nothing else. No other project's copy can be picked up by accident.
`SCHOOL_MYLIB` in the environment overrides it if you ever need to point
somewhere else for one run.

That file is **parsed, never imported**, and only those three names are read
out of it. Three reasons:

* importing it would run it, which pulls in `Levenshtein`, `gtts` and
  `psutil` - packages this script has no use for and a fresh laptop may not
  have;
* it also holds mail accounts and their secrets, and reading three named
  lists out of the syntax tree cannot touch them;
* a syntax error anywhere else in it cannot stop this script.

If none of it can be read the run stops and says how to fix it, rather than
collecting against an empty list.

**Which keywords this laptop runs** is a per-machine choice, so that stays
in `find_school.py`:

```python
ENABLED_SCHOOL_KEYWORDS = [
    "Private Schools",
    "CBSE Schools",
    ...
]
```

Leave it empty to run all 38. A name that is not in the shared list is
reported rather than silently ignored.

---

## Settings

One block at the top of `find_school.py`. Nothing else there is meant to be
edited; the technical constants sit lower down beside the code that uses
them.

| Setting | Default | What it does |
|---|---|---|
| `rank` | 1 | which laptop this is |
| `num_sys` | 25 | how many laptops share the work |
| `DOWN_PRESSES_PER_ROUND` | 10 | Down presses between two copies |
| `SCROLL_PAUSE_SECONDS` | 2.5 | wait after them for new results |
| `GEMINI_URL` | the signed-out chat | where research is done |
| `GEMINI_FIRST_READ_SECONDS` | 15 | earliest the answer is read |
| `GEMINI_ANSWER_TIMEOUT` | 240 | give up on one school after this |
| `GEMINI_RETRIES` | 2 | extra attempts per school |
| `CYCLE_SIZE` | 100 | rows in a finished cycle file |
| `KEEP_CAPTURED_TEXT` | True | keep each keyword's copied text |
| `LOG_KEEP_DAYS` | 10 | days of log kept in `find_school.log` |

Each copy takes the **whole** list, not the new part, so pressing Down
several times between copies cannot skip a school - it only means fewer
copies of the same text. One press moves the list about 24 pixels, so at
one press per copy a long city takes hours; at ten it takes minutes. Lower
`DOWN_PRESSES_PER_ROUND` to copy more often, raise it to finish sooner.

### When two records are the same school

Every school about to be saved is checked against **`master_data.xlsx`**,
not just against the city being collected. A record is a duplicate if any
one of these matches:

| Key | Notes |
|---|---|
| Google Maps place id | the strongest - one place, one id |
| **email** | any of the three email columns, case-insensitive |
| **website** | the domain, so `www.` and a page path do not matter |
| **phone number** | the last ten digits, so `+91 98111 22333`, `09811122333` and `9811122333` are one number |
| website + PIN code | |
| name + address, name + PIN, name + city | weakest, tried last |

A matching key is **not** enough on its own. Two records that match on a
key but contradict each other are different schools and both are kept:

* a different Google Maps place id
* a different PIN code
* a different branch word (Rohini / Dwarka)
* addresses with nothing in common

That guard is what makes the website rule safe. Every branch of a chain
shares one domain, so without it the second branch onwards would be thrown
away — losing real schools, which is far worse than a duplicate row:

```text
DPS Rohini                 -> SAVED
DPS Dwarka                 -> SAVED          (same website, different PIN)
Sunrise Public School      -> SAVED
Sunrise Pub School         -> duplicate      (same email, different case)
Sunrise (Main Branch)      -> duplicate      (same phone, written +91…)
Hilltop School             -> SAVED
Hilltop Senior School      -> duplicate      (same site, www + a path)
```

Blank values and placeholders never become keys, so `N/A`, `-` and an
empty cell cannot make two unrelated schools look like one.

**The cycle file is checked too, not just the master.** Before a school is
written, its email, website and phone are compared against
`master_data.xlsx` *and* against the rows already in `current_data.xlsx`.
A row can be sitting in the cycle that the master has not got — an
interrupted save, or a file edited by hand — and checking only the master
let its twin through, putting the same school in the cycle file twice.
The check runs before either file is written, so a refused school reaches
neither and the two never disagree. A finished cycle file is never
appended to again, and a row that drifted out of the master is copied back
into it by the repair at the start of every run.

The same test is used in all three places that ask the question — saving a
school, deciding whether a city still has schools missing from the master,
and repairing the files after an interrupted save. If they ever disagreed,
a school refused as a duplicate would be reported missing for ever and its
city could never close; a test fails if they do.

### A school needs a way to be contacted

A school is only written to `master_data.xlsx` and the cycle file if it has
**a phone number or an email** - its own, or the principal's. One with
neither is left out:

```text
School 7/98 - Some Tuition Centre
   no phone number and no email - not saved
```

The test runs **after** Gemini research, not before, so a school with no
number on Maps still gets its chance: Gemini supplies most of the emails.
It is not counted as a failure and does not keep its city open; the city
report says how many were left out:

```text
CITY COMPLETE: Rangpo - 29 saved, 1 duplicate(s), 34 with no phone or email
```

They stay in the city's temporary workbook while the city is being worked
on, and that workbook is removed when the city completes - so a skipped
school leaves no record beyond the log line above. If you ever want to see
which schools were skipped and why, keep the log, or set
`KEEP_CAPTURED_TEXT = True` to keep the copied Maps text they came from.

### No caps on a search

A keyword is finished when, and only when, the copied text contains
Google's own message:

```text
You've reached the end of the list.
```

There is no round limit and no clock. A search can still stop early, but
only on a real failure, and then the keyword stays open and is collected
again from the checkpoint:

| Stopped on | Meaning |
|---|---|
| end-of-list message in the copied text | complete |
| no new school in 180 Down presses | **incomplete** |
| the clipboard stopped being readable | **incomplete** |
| the list stopped moving | **incomplete** |

### Folders and profiles

The output folders are always flat - `Excel_Data/`, `Temp_Data/`,
`checkpoint.json`. `rank` and `num_sys` decide which cities this
laptop takes, never where it writes.

Chromium's two profiles are made in a temporary folder outside the project
and removed when the run ends:

```text
Browser profiles removed (35 MB, nothing left in the project folder)
```

A run killed outright cannot clear its own, so the next run sweeps up any
older than an hour - never one a run in progress might be using.

There is no `Backup/` folder. The safety net is the master file, the
finished cycle files, and the journal that is written around every save and
reconciled on the next start - a crash between the two writes is repaired,
and nothing is ever deleted to make room.

---

## Install---

## Install

```bash
python3 -m pip install playwright openpyxl
python3 -m playwright install chromium
sudo apt install xvfb      # only needed to run without a visible screen
```

Python 3.8 or newer. Nothing else is required: no API key, no account, no
second file.

---

## Run

```bash
python3 find_school.py --status
```

```bash
python3 find_school.py --dry-run
```

```bash
python3 find_school.py --self-test
```

```bash
python3 find_school.py --state Karnataka --city Bengaluru --limit 10 --test
```

```bash
python3 find_school.py --state Karnataka --city Bengaluru
```

```bash
xvfb-run -a python3 find_school.py
```

| Option | What it does |
|---|---|
| `--status` | rank, assigned cities, progress, retry queue, workbook rows |
| `--dry-run` | the whole search plan — no browser, no network |
| `--self-test` | 98 offline checks — no browser, no network |
| `--list-keywords` | all 49 school categories and which are on |
| `--state` / `--city` | narrow one run |
| `--limit N` | stop after N schools per city (testing) |
| `--test` | write to `School_Data/Test_Run`, never the real folders |
| `--retry-failed` | research the schools that failed before |
| `--no-gemini` | Maps discovery only, much faster |
| `--show` | put the Chromium window on your **real desktop**, even under `xvfb-run` |
| `--check-display` | say which screen `--show` would use, then exit |
| `--headless` | never draw a window at all |
| `--reset-city CITY` | collect one city again |
| `--rank` / `--num-sys` | override the work division for one run |
| `--verbose` / `--quiet` | more or less logging |

### Watching it work

```bash
xvfb-run -a python3 find_school.py
```

```bash
xvfb-run -a python3 find_school.py --show
```

The first draws into `xvfb-run`'s private screen — nothing appears, which is
what you want for a long unattended run. The second finds the display your
desktop is actually on and puts the window there, in front of you.

`--show` controls **two** things together: whether you see the browser,
and whether the step-by-step log is printed. Without it the run is silent
apart from problems and the final result:

```text
$ xvfb-run -a python3 find_school.py
Working through 35 city(ies), 245 search(es). Nothing more will be printed here until the run ends.
   to follow it : tail -f /home/anand/School_data _extractor/find_school.log
   to watch it  : run again with --show
```

The last line is **live**: it rewrites itself in place, so you can always
see what is happening now without the screen filling up.

```text
city 3/35 Delhi | keyword 2/7 CBSE Schools | copying from Maps - 48 school(s) so far (copy 17)
```

```text
city 3/35 Delhi | keyword 2/7 CBSE Schools | researching school 12/48 - Gupta Public School
```

Each finished city then leaves one permanent line behind it, so the terminal
keeps a history as the run goes on:

```text
  done: Rangpo, Sikkim            29 saved,   1 duplicate(s),  34 with no phone or email
```

That is the whole terminal output: four lines at the start, one live line,
one line per finished city, any warnings, and the closing summary. The
status line is skipped when the output is redirected to a file or a pipe,
where carriage returns would only make a mess.

That run still wrote 1,862 lines to `find_school.log` — **nothing is lost
from the record**, it simply is not on your screen. Warnings and errors
always come through, and so does the closing RUN SUMMARY. To follow a
silent run:

```bash
tail -f find_school.log
```

| | terminal | log file |
|---|---|---|
| default | problems and the result | every step |
| `--show` | every step | every step |
| `--verbose` | every step, plus debug | every step, plus debug |
| `--quiet` | errors only | every step |

**`--show` is the only thing that puts a window on your screen.** The rule
has no exceptions, and every run says which it is doing, in its first lines:

```text
The browser is hidden on the virtual screen xvfb-run made (:102) -
nothing appears on your desktop. Add --show to watch it.
```

```text
The browser is VISIBLE on your desktop (--show)
```

Two things keep a window off your desktop without `--show`, and both are
needed on a Wayland session:

* the code never looks up your desktop outside an `if self.show` — a test
  fails if it ever does;
* the browser is pinned to the virtual screen with `--ozone-platform=x11`,
  and `WAYLAND_DISPLAY` is removed from its environment. `xvfb-run`
  replaces `DISPLAY` but **not** `WAYLAND_DISPLAY`, so without this
  Chromium can find your real compositor by itself, ignore `DISPLAY` and
  put a window in front of you. This is the same guard `--show` uses,
  pointed the other way.

Note that a window drawn that way is a native Wayland window, so `xdotool`
and other X tools cannot see it — checking with them proves nothing.

If you do see a browser without `--show`, it is left over from an earlier
run that was killed — a run stopped with `kill -9` cannot close its own
Chromium, so its window stays on screen and the next run looks as though it
opened one. Each run now closes those first:

```text
1 browser window(s) left by a killed run were closed
```

Only a process that **is** a browser *and* holds one of this script's own
profile folders is touched, so your own Chrome or Brave is never affected.

Two things have to be right for that, and both are done for you:

* **`DISPLAY` and `XAUTHORITY` together.** `xvfb-run` replaces both. The
  display alone gives "Authorization required" and the browser dies, so the
  authority file is recovered too — from the environment of your own session's
  processes, falling back to the X sockets and the usual authority files.
* **`--ozone-platform=x11`.** On a Wayland desktop Chromium finds the
  compositor socket by itself, ignores `DISPLAY`, fails to create a surface
  and draws nothing — which looks exactly like `--show` being ignored.

Every candidate display is opened with `xdpyinfo` before it is used, and after
the browser starts the script looks at the screen and confirms a Chromium
window is really there. If it cannot find one it says so loudly rather than
pretending. If no desktop can be reached at all, the run stops with exit code
6 and tells you what to set:

```bash
SCHOOL_DESKTOP_DISPLAY=:0 SCHOOL_DESKTOP_XAUTHORITY=$XAUTHORITY xvfb-run -a python3 find_school.py --show
```

Run `python3 find_school.py --check-display` to see what it would pick.

### One browser at a time

The script never has two Chromiums open. Each stage opens the browser it
needs, finishes, and **closes it** before the next one starts:

```
[MAPS COPYING]     -> open the Maps browser, copy to the end of the list
                   -> close it
[REGEX EXTRACTION] -> read the .txt that was just saved (no browser at all)
[GEMINI RESEARCH]  -> open the Gemini browser, one school at a time
[EXCEL SAVING]     -> each row written and read back
                   -> close it, and open Maps again for the next keyword
```

With `--show` you watch that happen on your desktop; without it nothing is
drawn anywhere. Either way the lifecycle is the same.

Three things make it safe:

* **The Playwright driver belongs to the run, not to a browser.** A process
  may start the sync driver only once, so closing the Maps browser must not
  take the driver Gemini needs with it.
* **Nothing is opened until the last one has really gone.** The close is
  verified against the profile's `SingletonLock` and, with `--show`, against
  the window list. If a browser will not close, the next one is **not**
  opened and the run says so.
* **Gemini is not opened during collection, not even to check it.** There is
  no start-up readiness test; the first time its browser opens is when the
  first school is about to be researched, by which time the Maps browser is
  already closed.

The two profiles (`browser_profile` and `browser_profile_gemini`) are
separate and are kept between stages, so a sign-in survives.

### The old behaviour

**One tab, always Google Maps, always the search running right now.** It
retunes itself for every new city and every new keyword, and stray tabs are
closed each time.

Gemini does its research in a **second, always-headless Chromium** with its
own profile (`School_Data/browser_profile_gemini`), sharing the first one's
Playwright driver. It has no window, so it cannot cover the map or steal
focus mid-scroll. Nothing is signed in, so nothing is lost by the separate
profile; if you want to sign in to Gemini by hand, either do it in that
profile or set `GEMINI_SEPARATE_BROWSER = False`, which puts Gemini back in a
background tab of the main window and brings the Maps tab to the front after
every answer.

If the second browser cannot start, it falls back to that background tab by
itself and says so.

---

## Deploying to 25 laptops

Each laptop gets its own copy of `find_school.py`. Only two lines differ:

```python
# Laptop 7
rank = 7
num_sys = 25
```

`num_sys` must be the same everywhere. Rank 7 of 25 then collects cities 210
to 244 of the 871, and no other laptop touches them.

Optionally — never required — a `rank.json` beside the script, or the
`SCHOOL_RANK` / `SCHOOL_NUM_SYS` environment variables, override the values at
the top of the file. Precedence, lowest first:

```
the top of find_school.py  ->  rank.json  ->  environment  ->  command line
```

---

## What it does, in order

```
rank -> assigned cities -> keywords -> Maps discovery -> duplicate removal
     -> temporary storage -> Gemini research -> parsing -> validation
     -> Excel save -> checkpoint -> next city
```

One city is finished completely before the next one starts.

### Google Maps discovery — the keyboard, not the markup

The URL is the one you would type:

```
https://www.google.com/maps/search/Private+Schools+in+Bengaluru?hl=en&gl=in
```

Then, for every round, in this order and no other:

```
focus the results list -> Ctrl+A -> Ctrl+C -> read the clipboard
-> save the snapshot -> look for the end-of-list message
-> Down x12 -> wait -> repeat
```

Nothing is ever copied after a `Down` without the text before it having been
saved first. No card is clicked and no place page is opened — the only page
the browser visits is the search results list.

The text comes off the **system clipboard** (`navigator.clipboard.readText`),
with the selection as a fallback. If neither can be read, the search is left
**incomplete** and says so; it never quietly reads the page some other way.

Measured on a live Delhi search: the copied text grows as the list loads —
10 entries at 1,966 characters, 80 at 14,548, 118 at the end — because Maps
keeps earlier results in the page. Every snapshot is still saved and parsed,
because that is only true while Google chooses to keep them.

A search stops for one of five reasons, always printed:

| Stopped on | Meaning |
|---|---|
| end-of-list message in the copied text | complete |
| no new school in 6 copies | **incomplete** |
| the clipboard could not be read | **incomplete** |
| the 120 copy cap / 900 second timeout | **incomplete** |
| the list stopped moving | **incomplete** |

A keyword is never marked complete on an incomplete search, and a city is
never complete while any keyword is unresolved.

### Reading the copied text

Google repeats a result's name on two consecutive lines. That doubled line is
the entry boundary, and the regex extraction uses it to cut one school from
the next. Ratings, review counts, `Website`, `Directions`, filter labels and
the end-of-list line are all rejected as names.

The copied text carries no place links, so `School Google Maps Link` stays
empty rather than being invented, and identity is the name together with the
address and phone the entry showed — which keeps two branches of one chain
apart.

### Gemini research

`https://gemini.google.com/app?hl=en-IN`, the signed-out chat. No API key.

Only the **current response container** is read — not a copy of the whole
page, which would also contain the question, the previous answers and the
interface, and which is how one school's answer ends up in another school's
row. The wait is on the answer itself: it counts as finished when it carries
real `Field_Key:` lines and has stopped growing for 4 seconds, with a 240
second ceiling and 2 retries. There is no fixed 15-second delay.

An answer about a different school, or with a different PIN code, is rejected
and the school goes to the retry queue.

If Gemini shows a sign-in page, the run says so and stops. It does not try to
get past it.

### What is and is not verified

Values read from Google Maps are **observed**. Values from Gemini are
**AI-reported** and may be wrong. Gemini never overwrites an observed name,
address, phone number, PIN code, website or Maps link — it only fills columns
that are empty. Nothing is labelled verified.

---

## Output

Everything the run writes sits beside the script, and every folder is
per rank so two laptops can share one directory.

```
find_school.py            the whole program
test_find_school.py       the tests
README.md                 this file
requirements.txt          the two packages it needs
rank.example.json         copy to rank.json to set this laptop's rank
checkpoint.json           where the run got to

Excel_Data/
    master_data.xlsx      every unique school - the duplicate check
    current_data.xlsx     the active file, fills to 100
    current_data_cycle1.xlsx   finished, exactly 100 records
Temp_Data/                normally EMPTY - see below
find_school.log           one log file, trimmed to 10 days
```

### Temporary files: success cleans up, failure is preserved

`Temp_Data/` is working space, not storage. One rule governs all of it:

```text
SUCCESS            -> CLEAN UP
UNRESOLVED FAILURE -> PRESERVE FOR RESUME
```

Four kinds of file are written, each removed at the moment it stops being
needed:

| File | Written | Removed |
|---|---|---|
| `<State>_<City>_<keyword>.txt[.gz]` | while Maps is copied | when that keyword completes |
| `<State>_<City>_gemini_<school>.txt` | after a school is researched | when that school is saved and verified |
| `<State>_<City>_temporary_schools.xlsx` | as schools are discovered | when the city completes |
| `<State>_<City>_schools.json` | the resume point | when the city completes |

So after a clean run **`Temp_Data/` is empty**. Nothing accumulates.

The other half of the rule matters more. None of that cleanup runs from an
error path, and none of it is in a `finally:` block - a test fails if it
ever is. A school whose research failed keeps its answer; a keyword whose
list never reached its end keeps its copied text; a city with rows still
missing from the master keeps its working sheet. That is exactly what the
next run needs to carry on from the right school instead of collecting
everything again.

At startup the script looks for leftovers from work that already finished.
It removes them only when **two** things agree: the checkpoint says the
city is complete, *and* every row in that city's working sheet is read back
out of `master_data.xlsx` (or is a school with no phone and no email, which
is a finished skip). A city that fails either test is not touched and says
so:

```text
Delhi is marked complete but 102 of its school(s) are not in the master
   - its files are kept
```

`KEEP_CAPTURED_TEXT = False` is the default. Set it to `True` to keep each
keyword's copied text, packed, as the evidence behind every row. It only
ever applies to work that succeeded: text belonging to an unresolved
keyword or an unfinished city is preserved either way.

**Never removed by any of this:** `master_data.xlsx`, `current_data.xlsx`,
a finished cycle workbook, `checkpoint.json` or `find_school.log`.

Chromium cannot run without a profile directory, and fills one with tens of
MB of cache. By default those live in a temporary folder **outside this
project** and are removed when the run ends, so nothing is left behind:

```text
Browser profiles removed (37 MB, nothing left in the project folder)
```

A run killed with `kill -9` cannot clear its own, so the next run sweeps up
anything older than an hour.

`checkpoint.json` is the resume point: the city and keyword in progress, the
keywords already finished, the schools already researched and already saved.
Delete it and the next run starts the city from the beginning; keep it and
the run carries on where it stopped.

Each laptop has its own copy of this script and so its own folder, so the
output is not split by rank.


### Disk use

The first real run filled 656 MB, almost none of it data. Four things keep
it down:

| Setting | What it stops |
|---|---|
| `BROWSER_CACHE_MB = 40` | Chromium hoarding page cache - it was 625 MB |
| `COMPRESS_CAPTURE_FILES = True` | a capture file storing the same text a hundred times |
| `KEEP_CAPTURED_TEXT = False` | a capture file per keyword per city piling up |
| `LOG_KEEP_DAYS = 10` | the log growing for ever |

There is **one** log file, `find_school.log`. Each time the script starts it
drops any line dated more than `LOG_KEEP_DAYS` ago, so the file always holds
about the last ten days and never grows without limit. A line with no date
of its own - a traceback, a wrapped message - stays with the line above it.
Set `LOG_KEEP_DAYS = 0` to keep everything.

**Nothing collected is thrown away by any of them.** The capture text is
*packed*, not trimmed: a live Delhi keyword wrote 1.1 MB and now stores as
26 KB, and `private_schools.txt.gz` is read straight back by the extraction
step, byte for byte. Every school, every snapshot, every phone number is
still in it.

If you want to read one by hand:

```bash
zless "Temp_Data/Delhi_NCR_Delhi_private_schools.txt.gz"
```

With `KEEP_CAPTURED_TEXT = False` that file is removed as soon as its
keyword completes, so there is normally nothing there to read. Turn the
setting on if you want to keep it.

The browser caches can also be cleared at any time the script is not
running; they rebuild themselves, and sign-ins live in `Cookies` and
`Login Data`, not in the cache:

```bash
rm -rf browser_profile/Default/Cache "browser_profile/Default/Code Cache"
```

### The order of work

One keyword is finished completely before the next one starts:

```
city -> keyword -> Maps search -> Ctrl+A/Ctrl+C until the end of the list
     -> .txt -> regex extraction -> temporary workbook
     -> school 1: Gemini -> raw .txt -> regex -> merge -> master check
                  -> current_data.xlsx -> checkpoint
     -> school 2: the same
     -> ... -> keyword complete -> next keyword
     -> all keywords complete -> city complete -> next city
```

Each school reaches the master and the cycle file **as soon as its own
research finishes**, before the next school is asked. Nothing waits for the
keyword or the city to end.

The captured `.txt` is the evidence for every row and is **kept** after the
city finishes; the working files beside it are removed. Each card is written
once, the moment it is first seen, inside a block that names the scroll it
came from:

```
========== CARD 0001 | SNAPSHOT 1 | 16:22:59 ==========
LINK: https://www.google.com/maps/place/NORTH+SIKKIM+ACADEMY,+MANGAN/...
NAME: NORTH SIKKIM ACADEMY, MANGAN
TEXT:
NORTH SIKKIM ACADEMY, MANGAN
No reviews
General education school · GG3M+4X8, Rinzing Namgyal Marg
Closed · Opens 9 am Fri
========== END CARD 0001 ==========
```

The records are then built from that file, not from the live page, so the
extraction can be re-run and inspected without collecting anything again.

### Cycles and the master

`current_data.xlsx` fills to `CYCLE_SIZE` (100), is finalised as
`current_data_cycle1.xlsx`, and a fresh empty `current_data.xlsx` takes over.
`master_data.xlsx` holds every unique school this rank has saved, and is what
new schools are checked against — across restarts, cycle rollovers and
checkpoint resumes.

The master is written first, then the cycle file, with a journal around the
pair. If a run dies between the two, the next start copies whatever is in one
file but not the other into the other one. Nothing is ever deleted by that
repair.

Every save is atomic, a finished cycle file is never reopened for writing, a
damaged workbook is kept aside rather than deleted, and phone numbers, PIN
codes, UDISE codes and affiliation numbers are stored as text so leading
zeros survive.

A city's working files are removed only after its rows have been read back
out of the master.

---

## Tests

```bash
python3 find_school.py --self-test
```

```bash
python3 -m unittest test_find_school -v
```

98 offline checks inside the script, and 72 tests with mocked browser
responses in [`test_find_school.py`](test_find_school.py) — the city dataset
against `new_mylib.py`, the rank algorithm against `student_scraper.py`,
duplicate rules, the 57-field mapping, crash and resume at each phase, Excel
rollover, and that the single file imports no project file.

---

## Limits, honestly

* **Google Maps `robots.txt` disallows `/maps/search/` for automated
  clients.** The script reads the public results page with a real browser and
  does nothing to disguise itself; whether to run it is your decision.
* **A Maps card rarely shows a PIN code** — about 3 in 100. Most PIN codes
  come from Gemini, and are AI-reported.
* **Gemini can be wrong.** It is a research assistant, not a register. UDISE
  codes, affiliation numbers and registration numbers are the fields to treat
  with most suspicion, even though the prompt forbids inventing them.
* **Signed-out Gemini has quotas**, and may start asking for a sign-in. The
  run reports that and stops.
* No CAPTCHA solving, no proxy rotation, no fingerprint spoofing.

---

## What else is in this folder

The previous 34-file version was deleted on 13 November 2026, once the
single file had been tested, along with the archive of the data cleared for
the fresh start. One file was kept:

* `mylib.py.before_school_keywords.backup` - your company script's
  `lib/mylib.py` as it was before the 38 school keywords were appended to
  it. The only way to undo that edit.

`School_Data/browser_profile*` are the two Chromium profiles. They are large
(several hundred MB) and are rebuilt automatically if deleted - at the cost
of any sign-in held in them.
