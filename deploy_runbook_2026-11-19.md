# Deploying master `2ed2e61` to the ten production ranks

**Why this is a runbook and not a completed deployment:** all five laptops are
unreachable from this machine. It sits on `192.168.31.0/24`; the laptops are on
`192.168.1.0/24`, and port 22 times out on every one of
`192.168.1.121`–`192.168.1.125`. Nothing was deployed.

Run the steps below **on each laptop**, one rank at a time, in rank order.
Every step is read-only until step 4, and step 4 is reversible from the backup
step 3 makes.

Replace `<RANK>` with the rank number and `<DIR>` with that rank's
installation directory.

---

## 0. Confirm which rank this directory is — do not trust the table

```bash
cd <DIR>
cat rank.json
python3 find_school.py --status | head -20
```

**Gate.** Stop unless all of these hold:

- `rank.json` names the rank you intend to update, with `"num_sys": 10`
- `--status` reports `rank <RANK> of 10`
- the checkpoint line says `rank <RANK> of 10`

If `--status` says `of 5`, this is **not** a production rank — it is the
development machine. Leave it alone.

---

## 1. Is it running, and is the service boot-enabled?

```bash
systemctl is-enabled school-extractor-rank-<RANK>.service
systemctl is-active school-extractor-rank-<RANK>.service
pgrep -af "find_school.py"
```

Note the answers; step 6 must restore the same state.

---

## 2. Record the before state

```bash
cd <DIR>
python3 - <<'PY'
import glob, json, openpyxl, os
print("checkpoint:", json.load(open("checkpoint.json"))["rank"], "of",
      json.load(open("checkpoint.json"))["num_sys"])
c = json.load(open("checkpoint.json"))
print("completed cities:", len(c.get("completed_cities") or []))
print("enriched keys  :", sum(len(v) for v in (c.get("enriched") or {}).values()))
print("retry queue    :", len(c.get("failed") or []))
for p in sorted(glob.glob("Excel_Data/*.xlsx")):
    print(p, openpyxl.load_workbook(p, read_only=True).active.max_row - 1, "rows")
print("captures       :", len(glob.glob("Temp_Data/*")))
PY
git -C . status --porcelain
git -C . branch --show-current
git -C . remote -v
sha256sum find_school.py rank.json checkpoint.json
```

Keep this output. Step 7 compares against it.

---

## 3. Verified backup — never into an existing directory

```bash
cd <DIR>
B="$HOME/rank<RANK>_backup_$(date +%Y-%m-%d_%H%M%S)"
mkdir "$B"                     # fails if it somehow exists: good
cp -a find_school.py rank.json checkpoint.json "$B"/
for d in Excel_Data Temp_Data Backup; do [ -d "$d" ] && cp -a "$d" "$B"/; done
cp -a find_school.log "$B"/ 2>/dev/null
git -C . bundle create "$B/git_history.bundle" --all
git -C . status --porcelain > "$B/git_status.txt"
git -C . diff > "$B/uncommitted.patch"
sudo cp /etc/systemd/system/school-extractor-rank-<RANK>.service "$B"/ 2>/dev/null
( cd "$B" && find . -type f ! -name CHECKSUMS.sha256 -print0 \
    | xargs -0 sha256sum | sort -k2 > CHECKSUMS.sha256 )
```

Verify the copy before trusting it:

```bash
cd <DIR>
while read -r h f; do f="${f#./}"
  case "$f" in git_history.bundle|git_status.txt|uncommitted.patch|*.service) continue;; esac
  [ "$(sha256sum "$f" | cut -d' ' -f1)" = "$h" ] || echo "MISMATCH $f"
done < "$B/CHECKSUMS.sha256"; echo "backup verified"
```

**Gate.** Do not continue if anything prints `MISMATCH`.

---

## 4. Update the source only

Stop the service first so it is not reading a file mid-change. It finishes the
school it is on and writes its checkpoint:

```bash
sudo systemctl stop school-extractor-rank-<RANK>.service
```

If `git status` in step 2 showed **no** local modifications:

```bash
cd <DIR>
git fetch origin
git merge --ff-only origin/master
```

If it showed local modifications, do **not** discard them:

```bash
cd <DIR>
git stash push -m "local changes before 2ed2e61"
git fetch origin
git merge --ff-only origin/master
git stash pop        # resolve any conflict by hand, keeping both sides
```

Never `git reset --hard`, never `git clean`, never force anything.

Confirm the version and that `rank.json` is untouched:

```bash
git -C . rev-parse HEAD        # expect 2ed2e61c090889d0a97f40ee5530489dbccd9261
cat rank.json                  # expect this rank, num_sys 10
```

---

## 5. Timing settings, after the update

```bash
cd <DIR>
python3 - <<'PY'
import ast
src = open("find_school.py").read()
e = {n.targets[0].id: ast.get_source_segment(src, n.value)
     for n in ast.parse(src).body
     if isinstance(n, ast.Assign) and len(n.targets) == 1
     and isinstance(n.targets[0], ast.Name)}
want = {"GOOGLE_BETWEEN_SCHOOLS_DELAY_SECONDS": "(30, 150)",
        "GOOGLE_BETWEEN_QUERIES_DELAY_SECONDS": "(20, 120)",
        "GOOGLE_TO_GEMINI_DELAY_SECONDS": "(25, 130)",
        "DOWN_PRESSES_PER_ROUND": "random.randint(10,100)",
        "SCROLL_PAUSE_SECONDS": "random.randint(10,90)",
        "PAUSE_CLOSES_BROWSER_SECONDS": "random.randint(300, 900)",
        "BREAK_MAX_MINUTES": "120", "NIGHT_PAUSE_ENABLED": "True",
        "GEMINI_ANSWER_TIMEOUT": "240"}
bad = [k for k, v in want.items()
       if (e.get(k) or "").replace(" ", "") != v.replace(" ", "")]
for k, v in want.items():
    print("%-40s %-26s %s" % (k, e.get(k), "ok" if k not in bad else "MISMATCH"))
print("mismatches:", len(bad))
PY
```

**Gate.** Any mismatch means `rank.json` or a local edit is overriding the
configuration. Stop and roll back (step 8).

---

## 6. Tests, then resume

```bash
cd <DIR>
python3 find_school.py --self-test                 # expect 134 checks passed
python3 -m unittest test_find_school -q            # expect 357 OK
python3 -m unittest test_google_pacing -q          # expect 54 OK
python3 -m unittest test_night_pause -q            # expect 23 OK
python3 -m unittest test_rank_repo -q              # expect 12 OK
```

Run the tests from a **scratch copy**, not the live directory, if you can:
a suite run inside a production folder has been seen to leave a synthetic
capture in `Temp_Data`. If you must run them in place, afterwards check:

```bash
ls Temp_Data/     # every file must be a city THIS rank collects
```

Then start it again, only if it was running before:

```bash
sudo systemctl start school-extractor-rank-<RANK>.service
systemctl status school-extractor-rank-<RANK>.service --no-pager | head -12
journalctl -u school-extractor-rank-<RANK>.service -n 40 --no-pager
```

---

## 7. Verify, and compare with step 2

```bash
cd <DIR>
python3 find_school.py --status | head -24
```

**Gate.** All of these must hold:

- `rank <RANK> of 10`, night stop `23:00-05:00 Asia/Kolkata`
- Excel row counts **the same or higher** than step 2, never lower
- completed cities and enriched keys the same or higher
- the log shows `[MAPS SKIPPED] reusing the saved list` on a city with a
  complete capture — it resumed instead of recollecting
- `[GIT] pushed ... to <remote>/school-data-rank-<RANK>` — its own branch
- no new `(blank)` or non-`Private`/`Government` value in `School Type`
- the retry queue is not silently emptied

Classification check:

```bash
python3 - <<'PY'
import importlib.util, openpyxl
s = importlib.util.spec_from_file_location("fs", "find_school.py")
fs = importlib.util.module_from_spec(s); s.loader.exec_module(fs)
ws = openpyxl.load_workbook("Excel_Data/master_data.xlsx", read_only=True).active
h = [c.value for c in next(ws.iter_rows(max_row=1))]; i = h.index("School Type")
bad = [r[0] for r in ws.iter_rows(min_row=2, values_only=True)
       if not fs.valid_school_type(r[i])]
print("columns:", len(h), "| rows:", ws.max_row - 1, "| invalid School Type:", len(bad))
PY
```

Expect 57 columns and **0** invalid.

---

## 8. Rollback, if any gate fails

```bash
sudo systemctl stop school-extractor-rank-<RANK>.service
cd <DIR>
cp -a "$B/find_school.py" find_school.py
cp -a "$B/rank.json" rank.json
cp -a "$B/checkpoint.json" checkpoint.json
rm -rf Excel_Data Temp_Data Backup
cp -a "$B/Excel_Data" "$B/Temp_Data" "$B/Backup" .
python3 find_school.py --status | head -20
sudo systemctl start school-extractor-rank-<RANK>.service
```

Then stop the rollout and report which gate failed.

---

## Order

Rank 1 alone, through every gate, before any other rank. Then one laptop at a
time: ranks 1–2, then 3–4, 5–6, 7–8, 9–10. Do not start a laptop until the
previous one has passed step 7 for both of its ranks.

## Leave alone

- `RANK_REPO_ENABLED` stays `False`. The rank-wise URLs are generated and
  tested but no repository exists yet; nothing uses them.
- The existing `school-data-rank-N` data branches stay as they are. They are
  not replaced by separate repositories in this deployment.
- The development machine at `/home/anand/School_data _extractor` is **rank 1
  of 5** with its own checkpoint. It is not production rank 1. Do not deploy
  to it as part of this rollout and do not migrate its checkpoint.
