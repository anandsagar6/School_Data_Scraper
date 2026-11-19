# School Data Extractor — work summary

**Date:** 18 November 2026
**Project:** `/home/anand/School_data _extractor`
**Repository:** `git@github.com:anandsagar6/School_Data_Scraper.git` (private)
**Status:** two commits pushed; three bodies of work finished in isolation and
waiting on approval; nine items blocked on access or a decision.

> Yesterday's work (17 Nov) is in commits `2d57bb1`…`0216a40` — the
> Google-Search-only contacts, the phone search, school-level resume, the run
> lock, and the ten-rank deployment that was built and then removed.

---

## The short version

Today produced only **two commits**, because most of the day went into work
that is deliberately **not deployed**: three specs arrived, each asking for
changes to be built in isolation and reported for approval first.

| | This morning | Tonight |
|---|---|---|
| `School Type` | left to Gemini to guess | **set from the keyword** (committed) |
| Allowed type values | any word Gemini chose | **only `Private` or `Government`** (isolated) |
| Duplicate classification | first keyword to find the school won | **Private always wins, either order** (isolated) |
| A row already saved wrong | stayed wrong for ever | **corrected in place across every workbook** (isolated) |
| `Modern School` vs `…Branch II` | merged into one row | **kept separate** (isolated) |
| Rank → repository | one repo, ten branches | **URL generator built, switched off** (isolated) |
| Failing tests | 2 | **0** in the isolated copy |
| Tests / self-test checks | 309 / 121 | **350 / 134** in the isolated copy |

---

## 1. Committed and pushed today

| Time | Commit | What it did |
|---|---|---|
| 08:05 | `6e012e0` | add the systemd unit for review, not installed |
| 08:33 | `d1c938e` | `School Type` comes from the keyword, not from a guess |

### 1.1 The systemd unit (`6e012e0`)

Committed after a pre-flight check: `systemd-analyze verify` clean, the
`Description` matching the source (`rank 1 of 5`), `WorkingDirectory` and
`ExecStart` paths present, `/usr/bin/xvfb-run` present, SIGTERM handled in the
script. **Not installed and not enabled** — installing needs a `sudo` password.

### 1.2 `School Type` from the keyword (`d1c938e`)

The column was never new — it is the 9th of the 57 — it was simply being left
to Gemini. It is now stamped from the keyword where the fresh-capture and
reused-capture paths meet, so a keyword resumed after a restart marks its
schools exactly as the first pass did. Gemini is no longer asked for it (54 of
57), and `apply_answer()` drops a value for it even if one arrives.

---

## 2. Verified today, no code changed

### 2.1 The phone search was already in place

You asked twice whether Google is searched for the phone as well as the email.
It is — built yesterday in `d744a06` and kept when the deployment work was
removed. Rather than assert it, I proved it on a fake browser recording every
URL:

```
search 1: Official email address of Sunrise Public School in MG Road Bengaluru Karnataka
search 2: Official contact phone number of Sunrise Public School in MG Road Bengaluru Karnataka
```

Then on a **real school**, live:

```
School 1/1 - The British School, New Delhi
  email verified by the school's own mail domain
  phone verified by the school name and place beside it in the email results
  phone verified by the school name and place beside it in the phone results
  official email thebritishschool@british-school.org | official phone +911140664166
```

Both searches ran; the number was verified in the email results and then
re-verified and upgraded by the dedicated phone search. 57 columns, 40 fields
filled, and the only host the whole run visited was `https://www.google.com` —
no school website opened, even though Gemini supplied the website address.

That check ran in a throwaway folder with sync off, so nothing was written to
production, and the folder was deleted afterwards.

### 2.2 Automatic restart after a server outage

Asked, and answered with evidence rather than assurances:

- **Today: nothing restarts by itself.** `systemctl is-enabled` →
  `not-found`, 0 units installed, nothing in `/etc/systemd/system`.
- **A hard kill leaves no stale lock.** I `SIGKILL`ed a running instance, as a
  power cut would. The lock file stays on disk but the kernel drops the lock,
  so the next run takes it with no cleanup: `lock taken = True | blocked by:
  nobody`. While the first was alive a second run was correctly refused.
- **Resume works** at the next unresearched school, reusing a completed capture
  instead of re-scrolling — covered by three committed tests.
- A **network** drop with the machine up does not stop a run: failures are
  caught, schools queue for retry, sync marks itself pending.

### 2.3 Install commands handed over

With the pre-flight results, and `enable` separated from `start` on purpose, so
arming it for the next boot does not immediately begin collecting.

---

## 3. Removed again, at your request

I began converting three of your settings to per-use ranges —
`PAUSE_CLOSES_BROWSER_SECONDS`, `DOWN_PRESSES_PER_ROUND`,
`SCROLL_PAUSE_SECONDS` — and decoupling the no-progress threshold. You asked
for that removed. It was, restored from a backup taken before my first edit, so
**your own eight edits stayed intact**:

```
BREAK_MAX_MINUTES = 120          GOOGLE_BETWEEN_SCHOOLS_DELAY_SECONDS = (30, 150)
NIGHT_PAUSE_ENABLED = True       GOOGLE_BETWEEN_QUERIES_DELAY_SECONDS = (20, 120)
                                 GOOGLE_TO_GEMINI_DELAY_SECONDS       = (25, 130)
PAUSE_CLOSES_BROWSER_SECONDS = random.randint(300, 900)
DOWN_PRESSES_PER_ROUND       = random.randint(10,100)
SCROLL_PAUSE_SECONDS         = random.randint(10,90)
```

The interrupted edit to the test file had already been written, so that was
reverted to the commit as well.

---

## 4. Built in isolation today — awaiting approval

Nothing below is deployed. Production `find_school.py` is byte-identical to the
baseline taken this evening.

### 4.1 `/home/anand/school_task_2026-11-18/working` — rank-wise repository URLs

```
rank  1 -> git@github.com:YOUR_TEST_ACCOUNT/school_extractor1.git
rank 10 -> git@github.com:YOUR_TEST_ACCOUNT/school_extractor10.git
```

The rank comes from the existing ladder, so `rank.json` saying 7 gives
repository 7 and `SCHOOL_RANK=3` gives 3. `RANK_REPO_ENABLED = False`, so live
sync is untouched. `rank_repo_ready()` proves existence **and** write access
with `ls-remote` plus `push --dry-run`, sending no objects.
`github_profile_extractor` is refused by name — it belongs to the other
project. **25 checks pass, against local bare repositories only.**

One correction to the senior's snippet: `base.replace(".git", f"{rank}.git")`
replaces *every* `.git` in a string; the implementation slices the suffix.

### 4.2 `/home/anand/school_type_task_2026-11-18/working` — the School Type rules

**350 tests, zero failures. 134 self-test checks.**

Three defects found by inspecting the nine functions before changing anything:

1. The keyword deriver returned **any** word — `"International Schools"` →
   `"International"`, which would have broken the website's filter.
2. **No Private-over-Government rule existed.** `merge_records()` only filled
   blanks, so the *first* keyword to reach a school won — order-dependent.
3. A record duplicating a **master** row was *dropped*, never merged, so a
   school saved as Government and later found as Private **kept the wrong
   classification for ever**.

What it does now:

- Only `Private` or `Government` can ever be written; any other keyword yields
  `""` rather than a third category.
- `Government + Private → Private`, `Private + Government → Private`,
  `Private + Private → Private`, `Government + Government → Government`.
- A row already saved as Government is corrected **in place**, with no
  duplicate row, across `master_data.xlsx`, `current_data.xlsx`, finished cycle
  files **and** the temporary workbook a resumed run reads back.
- Safety: every file copied and verified with `filecmp.cmp(shallow=False)`
  before any write, atomic temp-and-replace, and **full rollback** if any file
  fails — proven by a test that fails the second write and checks all three
  files come back byte-identical.

### 4.3 The campus/branch fix

`Modern School` and `Modern School Branch II` were merging. The naive fix is a
trap: the existing branch-word list contains **road, nagar, colony, main,
north/south** — address words — so a one-sided check would split one school the
moment one copy had its locality filled in.

A narrow designator list is now read from the **school name only**, and yields
to stronger evidence:

```
Modern School vs Modern School Branch II (no shared contact) -> separate
Modern School, locality blank vs "Barakhamba Road"           -> one school
Same Maps place id, one name says Branch II                  -> one school
River School vs River School Annexe, SAME phone              -> one school
Branch II vs Branch III                                      -> separate
```

The shared-contact exemption came out of a regression: the first version split
`River School` / `River School Annexe` on one phone line.

### 4.4 The two failing tests — outdated assumptions, not a defect

I measured rather than reasoned, pinning the press count and zeroing the waits:

```
WITH the end-of-list message        WITHOUT it (fallback path)
presses  complete  records          presses  complete  records
     10      True      200               10     False      200
     50      True      200               50     False      200
    100      True      200              100     False      200
```

**At every press count all 200 records arrive.** The press count changes how
*often* the list is copied, never how much — each copy takes the whole list.
Both tests now pin the count and assert the real invariant; **coverage is
stronger, not weaker**, and the suite is green.

This also corrected something I had said earlier: because `random.randint()` is
evaluated **once at import**, `DOWN_PRESSES_PER_ROUND` is constant for a run,
so the existing `idle × setting` formula *does* equal the real press count. The
no-progress logic is sound as written and I changed none of it.

### 4.5 Provenance of the two bad rows — evidence, not names

| Evidence | Result |
|---|---|
| Both schools in `Delhi_NCR_Delhi_private_schools.txt.gz` | yes |
| Any Government capture on disk | **none exists** |
| `keywords_done` in the checkpoint | `{}` — Government never completed |
| The temp workbook from that keyword | 116 rows: **114 `Private`**, these 2 odd |

Demonstrated on a copy of your real data: 6 cells corrected across 3 workbooks,
5 rows stayed 5 rows, 57 columns intact, no leftovers in `Temp_Data`.

---

## 5. Configuration preserved — proven, not promised

A baseline of **all 162 module-level assignments** was captured as full source
expressions, so `random.randint(300, 900)` is recorded as written rather than
as the number it drew. After all of today's work:

```
expressions CHANGED : NONE
settings REMOVED    : none
settings ADDED      : SCHOOL_TYPES_ALLOWED, SCHOOL_TYPE_GOVERNMENT,
                      SCHOOL_TYPE_PRIVATE, _CAMPUS_DESIGNATORS
```

All four additions are new names. Production `find_school.py` is still sha256
`fea6391e…`, identical to this evening's baseline.

---

## 6. Open items

| # | Item | Blocked on |
|---|---|---|
| 1 | Two live rows still read `International School` and `Co-Educational` | approval to edit live Excel |
| 2 | systemd service not installed — nothing restarts after a reboot | `sudo` password |
| 3 | Unattended GitHub sync cannot authenticate | all three SSH keys are passphrase-encrypted; needs a passphrase-free deploy key |
| 4 | 15–20 school email verification | only 5 records exist locally |
| 5 | Collection-rate benchmark | no access to the five laptops |
| 6 | `125 per 3 hours` vs `1,000 per day` | the two conflict — 125 × 6 blocks ≈ **750/day** with the night pause |
| 7 | Finished cycle files | correction path tested only against synthetic cycles |
| 8 | Four Gmail app passwords in `mylib_ORIGINAL_with_credentials_2026-11-17.py.bak` | never confirmed rotated |
| 9 | `school_extractor_repo_backup_2026-11-17.tar.gz` | holds school records **and** the old credentials; deletion was blocked |

### Production state tonight

- `find_school.py` 7,555 lines; **309 tests, 2 failures** — both from your own
  uncommitted timing edits, and both fixed in the isolated copy.
- Uncommitted: your eight timing edits, plus this summary.
- 5 school records, 12 branches on GitHub, no systemd service installed.

---

## 7. Suggested order for tomorrow

1. **Decide the two live rows** — provenance is settled (§4.5).
2. **Approve or reject the campus change** (§4.3) — it will split any
   `X` / `X Branch II` pair that shares no contact, so row counts change.
3. **Deploy the School Type work** if approved, then the rank-URL generator.
4. **Create a passphrase-free deploy key** (item 3) — until then no laptop can
   sync unattended, which blocks the whole fleet design.
5. **Install the systemd unit** (item 2) and schedule a real reboot test.
6. **Settle the rate target** (item 6) before any pacing is tuned.
7. **Rotate the Gmail passwords and remove the two backup files** (items 8, 9).
