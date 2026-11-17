# School Data Extractor — work summary

**Date:** 12 November 2026
**Project:** `/home/anand/School_data _extractor`
**Status:** rebuilt as one standalone file, tested live, ready for a real run

---

## The short version

Yesterday's 34-file project became **one file you can copy to a laptop**:
`find_school.py`, 4,684 lines, importing nothing from this folder. Along the
way the collection method, the processing order, the Excel layout and the
browser behaviour were all replaced — each because a live test showed the old
one was wrong.

| | Yesterday | Today |
|---|---|---|
| Files needed to run | 34 | **1** |
| Maps collection | DOM card reading | **Ctrl+A / Ctrl+C, real clipboard** |
| Order of work | all keywords, then all research | **one keyword start to finish** |
| Excel | one growing workbook | **master + 100-record cycles** |
| Saving | once per city | **once per school** |
| Folders | shared | **per rank** |
| Time settings | scattered | **30 settings, one block** |
| Tests | 264 | **106 self-test + 159 mocked** |

---

## What was built, in order

### 1. The brake came off

`LARGE_RUN_CITY_LIMIT` was refusing any run over 3 cities. Set to `0`, and
`0` now means *no limit* rather than *refuse everything*, so
`xvfb-run -a python3 find_school.py` just runs.

### 2. The Gemini step was corrected

It was settling on half-written answers and reading the wrong part of the
page. Now: wait, read the finished answer, keep the raw text.

### 3. Rebuilt as one standalone file

The city dataset, the 38 search keywords, the 57 fields, both Gemini prompts,
the parser, the duplicate rules, the Excel writer, the checkpoint and the
browser driver are all embedded. Two dependencies, documented:

```bash
python3 -m pip install playwright openpyxl
python3 -m playwright install chromium
```

The old project is kept, unused, in `Old_Multifile_Version/`.

### 4. The city list moved to your company script's `mylib.py`

`city_by_states` from `/home/anand/company_extractor_script/lib/mylib.py`:
**29 groups, 905 entries, 866 unique**. Your 38 school keywords were appended
to that file — byte-for-byte additive, nothing else touched. A test fails if
any of its four Gmail passwords ever appears in `find_school.py`.

`get_cities_by_rank()` is the GitHub extractor's function, unchanged; a test
compares the two files' syntax trees.

### 5. `--show` puts the window on your real desktop

Two things had to be right, and only the first was obvious:

* `xvfb-run` replaces **`XAUTHORITY` as well as `DISPLAY`**. The authority
  file is recovered from your own session's processes.
* **`--ozone-platform=x11`.** On a Wayland desktop Chromium finds the
  compositor by itself, ignores `DISPLAY`, fails to make a surface and draws
  nothing — which looks exactly like `--show` being ignored.

After launching, the script reads the screen and confirms a window is really
there. If no desktop can be reached it stops with exit code 6 and says what
to set.

### 6. One visible tab, Gemini out of the way

The `about:blank` tab is gone (the profile's own blank tab is reused), and
Gemini moved to a **second, always-headless Chromium** with its own profile,
sharing the first one's Playwright driver. The visible window shows only the
Maps search, retuned to the current keyword.

### 7. Google Maps is now read with the keyboard

```
focus the results list -> Ctrl+A -> Ctrl+C -> read the clipboard
-> save the snapshot -> check for the end-of-list -> Down x12 -> repeat
```

Real `navigator.clipboard.readText()`. Nothing is copied after a `Down`
without the text before it being saved first. No card is clicked and no
school's own page is opened.

Measured on a live Delhi search: the copied text grows as the list loads —
**1,966 characters at 10 entries, 20,974 at 118** — and `Down` moves the list
(scrollTop 0 → 10,826). Entries arrive in blocks of ten every third round,
which is why the no-progress limit is 6 and not 3.

### 8. One keyword, start to finish

```
city -> keyword -> Maps search -> Ctrl+A/Ctrl+C to the end -> .txt
     -> regex extraction -> temporary workbook
     -> school 1: Gemini -> raw .txt -> regex -> merge -> master check
                  -> current_data.xlsx -> checkpoint
     -> school 2: the same -> ... -> keyword complete -> next keyword
```

Each school reaches the master **as soon as its own research finishes**,
before the next school is asked.

### 9. Excel: a master and 100-record cycles

```
Excel_Data/rank_1/
    current_data.xlsx          the active file, fills to 100
    current_data_cycle1.xlsx   finished, exactly 100
    current_data_cycle2.xlsx   finished, exactly 100
    master_data.xlsx           every unique school - the duplicate check
```

The master is written first, then the cycle file, with a journal around the
pair; a run that dies between them is repaired on the next start by copying
whatever is in one file but not the other. Nothing is ever deleted by that
repair.

### 10. Every time gap is a setting

30 named settings in one block directly under `rank` / `num_sys` — page
loads, the copy loop, Gemini, pacing, the window check. A test walks the
whole file and fails if any `time.sleep()` is ever given a bare number again.

---

## Live results — actually run, not mocked

| Test | Result |
|---|---|
| `Private Schools in Delhi`, clipboard | **118 schools, 38 snapshots, 202 s**, ended on the real end-of-list message |
| Same search, before the rewrite | 117 schools, 83 s (DOM reading — faster, but not what you asked for) |
| `Jorethang`, 7 keywords | 164 discovered, 131 unique, cycle1 finalised at exactly 100, master 131 |
| `Mangan`, full city | **36 schools, city COMPLETE, retry queue 0** |
| Gemini on one school | 42 of 57 columns filled, signed out, no API key |
| Crash recovery | killed mid-school; resumed skipping 2 finished keywords, re-researched nothing |

---

## Bugs found and fixed — each by a test or a live run

| # | Bug | Why it mattered |
|---|---|---|
| 1 | `--limit` capped collection **and** let the city be marked complete | the original "six schools" |
| 2 | An affiliation number in a school's name was read as its PIN | then rejected Gemini's correct answer |
| 3 | A rating line could become the address | `4.7 ★★★★★ (149)` in the address column |
| 4 | A phone on its own line was missed | cards with no opening hours |
| 5 | "Restore pages?" popup after any Ctrl+C'd run | a stray click reopened the old tabs |
| 6 | `\s*` in a regex spanned newlines | an empty `LINK:` captured the next line |
| 7 | A keyword with no results blocked the city for ever | a small town has no ICSE school |
| 8 | A name-only school had **no key at all** | invisible to the duplicate check |
| 9 | Enrichment hid saved schools | Gemini's fuller address broke the lookup |
| 10 | An interrupted keyword abandoned its schools | discovered, never researched |
| 11 | Keyword flags masked unfinished work | **permanent deadlock — the city could never close** |

Numbers 8–11 were all found by one live run of Mangan that refused to finish.

---

## Where it stands

```bash
python3 find_school.py --status
```

```text
rank                : 1 of 25
cities for this rank: 35
search keywords     : 7 of 38 enabled
cycle size          : 100 records per file
master_data.xlsx    : 0 record(s)
cities completed    : 0
```

Nothing has been collected into the real folders. The two commands:

```bash
xvfb-run -a python3 find_school.py
```

```bash
xvfb-run -a python3 find_school.py --show
```

**106 self-test checks and 159 mocked tests pass.**

---

## Honest limits

* The clipboard method is **2.4× slower** than reading the markup — 202 s
  against 83 s for the same Delhi search. That is the cost of copying 38
  times instead of reading once.
* The copied text carries **no Maps links**, so that column stays empty
  rather than being invented, and identity is name + address + phone.
* `robots.txt` disallows `/maps/search/` for automated clients. The script
  uses a real browser and disguises nothing; running it is your decision.
* Gemini can be wrong. UDISE codes, affiliation numbers and registration
  numbers deserve the most suspicion.
* A 250-school multi-cycle run has been proven in tests and once at 100 rows
  live, not at 250.
* The reference projects are untouched: 0 files changed in
  `github_profile_extractor`, `engineering_college_extractor_script`,
  `company_extractor_script` (apart from the keywords you asked to add) and
  `naukri_script`.
