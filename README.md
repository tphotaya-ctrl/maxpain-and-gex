# Max Pain / GEX daily updater (CME Gold weekly options)

Fills one Excel workbook, `MaxPain_GEX.xlsx`, automatically from CME Group data. It holds both Max Pain and GEX (calculator GEX V.4.1), and each day is logged.

| Sheet | Filled by | Source |
|---|---|---|
| `OI Data` → `Max Pain Calc` | Call/Put OI per strike, CME totals, DTE; futures settle into `Max Pain Calc!F4` | CME "Volume & OI" JSON endpoints |
| `Log` | one row per day: Max Pain, distance, P/C, ... | |
| `Gamma Data` → `GEX Calc` | Call/Put Gamma (1 Pct) per strike (rows 2-152); futures settle into `GEX Calc!J6` | CME QuikStrike, OI Matrix view (needs CME login) |
| `GEX Log` | one row per day: totals, NET GEX, mode, Gamma Flip, peak | |
| `บันทึกโหมด` | **manual** - your own daily journal, never written by the scripts | |
| `Outcome` | created by `outcomes.py`, see below | CME futures settlements |

Max Pain and GEX results are calculated by the **Excel formulas**. Python only fills the input sheets and the logs.

`MaxPain_GEX.xlsx` was built once by `python merge_workbooks.py` from the three files it replaced: `คำนวน_max_pain_10 (2).xlsx`, `GEX V.4.1.xlsx`, and `คำนวน GEX 1.xlsx` (only its `Log`, carried over as `GEX Log`). The originals are kept as they were. The merge copies sheets with Excel itself, so charts, dropdowns and array formulas survive. It also widens V.4.1's ranges to rows 2-152:
- `GEX Calc` A:F use the blank-if-empty form on every row (rows 2-66 used to show 0 for an empty strike, which pulled MEDIAN down).
- E/F's MEDIAN, the H helper and the conditional formatting now reach row 152.
- J16 checks the last filled strike instead of row 62.
- Nothing else in V.4.1's logic changed.

The GEX `สถานะ` in `GEX Log` follows V.4.1's J5 rule ("ไม่มีโหมด" when |NET| < 5% of Σ|call−put|), so it reads differently from pre-V.4.1 rows ("แกว่งแรง"). Gamma Flip still comes from Python (`update_gex.gamma_flip`), since V.4.1 has no flip column.

## Requirements
- Windows, Google Chrome installed, Python 3.10+, Microsoft Excel (only `verify.py` uses it, via COM)
- `pip install -r requirements.txt` (or `-r requirements-dev.txt` to also get `pytest`)
- A CME Group account (free) for the gamma part
- A Google account + a service-account credential (free), only if you want the automatic mobile/web view - see **Mobile/web view**

## Tests
```
pip install -r requirements-dev.txt
python -m pytest -q
```
Covers the pure logic only (contract selection, Max Pain, Gamma Flip, retry behaviour, atomic
save, notification fallback) with no network, browser, or CME login - `tests/` for what's
covered and why. `.github/workflows/ci.yml` runs this (plus `python -m compileall`) on every
push, on `windows-latest`. The live paths - `fetch()`, `fetch_gamma()`'s actual QuikStrike
scraping, `verify.py`'s Excel-COM checks - aren't covered by CI; keep verifying those by hand
per **Daily run** below.

## Setup (once)
```
python quikstrike_login.py
```
A Chrome window opens. Log in to CME yourself, open the QuikStrike view once, close the window. The session is kept in `.chrome_profile/` (git-ignored - it holds your login cookies, never commit or share it). Without a session the gamma step is skipped and the Max Pain step still runs.

## Daily run
Runs automatically: Windows Task Scheduler task **"MaxPainGEX Daily Update"**, weekdays 08:30 Bangkok time, as the logged-in user (must stay logged in - the browser runs headed, Chrome will visibly pop up). It runs `run_daily.bat`, which appends a timestamped block to `run.log` ending in one summary line - `OK`, `WARN - ... GEX skipped` (Max Pain still updated), or `FAIL`. Re-create the task with:
```
schtasks /create /tn "MaxPainGEX Daily Update" /tr "\"<repo path>\run_daily.bat\"" /sc weekly /d MON,TUE,WED,THU,FRI /st 08:30 /rl limited /f
```
Manual run:
```
python update_workbooks.py        # or run_daily.bat (appends to run.log)
python verify.py                  # independent re-check, exit code 1 on any FAIL
```
Close both workbooks in Excel first, otherwise saving fails.

`config.json`:

| key | meaning |
|---|---|
| `target` | `today` (contract expiring today, local date; next expiry if none), `YYYY-MM-DD`, `auto` (nearest of `family`), or a label - `Week 4 - SEP 2026` for a weekly, `DEC 2026` for the standard monthly series |
| `family` | product used by `auto`/label: `MW1` Mon, `AB1` Tue, `WD1` Wed, `BB1` Thu, `E21` Fri, `AME` standard/monthly |
| `include_monthly` | default `true`: whether `today`/`YYYY-MM-DD` also considers the standard monthly series, not just weeklies |
| `price_month` | optional override, e.g. `DEC 26`: normally derived automatically from CME's own data, see below |
| `strike_min` / `strike_max` | strike window written to `OI Data` (max 200 rows) - the standard monthly series can have 600+ strikes, narrow this if you target it |
| `workbook`, `gex_workbook` | files to fill - both `MaxPain_GEX.xlsx` now (separate files still work if the GEX one has the V.4.1 sheets) |
| `google_sheet_id` | the ID from your Google Sheet's URL (`.../d/<this part>/edit`) - enables automatic sync, see **Mobile/web view** |
| `google_credentials_file` | optional, default `service_account.json` - path to the Google service-account key, resolved relative to the repo root |

Set env `MAXPAIN_CONFIG=path\to\other.json` to work on copies without touching the originals.

## Mobile/web view
Two ways to check today's numbers without opening Excel:

**Google Sheet (automatic, private)** - `sheets_sync.py` writes each workbook's latest `Log`
row into its own tab ("Max Pain Log" / "GEX Log") of a Google Sheet you own. It runs
automatically at the end of `update_workbooks.py` (a failure there is caught and never fails
the daily run - same pattern as the GEX step), so it updates on whatever schedule the Task
Scheduler job already runs on, no manual step. A same-day rerun updates that day's row in
place instead of duplicating it. One-time setup:
1. In [Google Cloud Console](https://console.cloud.google.com/), create a project, enable the
   **Google Sheets API**, then create a **Service Account** and download its JSON key as
   `service_account.json` in the repo root (git-ignored - it's a credential, never commit it).
2. Create a blank Google Sheet, then share it (Editor access) with the service account's
   `client_email` (inside the JSON key, looks like `...@...iam.gserviceaccount.com`).
3. Copy the Sheet's ID out of its URL into `config.json`'s `google_sheet_id`.
4. `pip install -r requirements.txt` (adds `gspread`/`google-auth`), then run
   `python update_workbooks.py` (or just `python sheets_sync.py` to sync without a full run).

View it any time from the Google Sheets app on your phone, or the Sheet's normal share link -
only whoever you've shared it with can see it, unlike the option below.

**GitHub Pages (manual, public)** - `publish_report.py` renders `docs/index.html`, a summary
card for each workbook plus the last 20 `Log` rows, from the same data. It also runs
automatically at the end of `update_workbooks.py`, but **publishing it is a separate, manual
step** - the daily automation does not push anything by itself, since that would mean an
unattended scheduled task pushing to a public GitHub repo every day with no one looking:
```
git add docs && git commit -m "Update report" && git push
```
One-time setup so the page is reachable: GitHub repo -> Settings -> Pages -> source = `main`
branch, `/docs` folder. After that, `https://tphotaya-ctrl.github.io/maxpain-and-gex/` works
from any phone or browser. **The repo is public, so this page is public too** - anyone with
the link can see the daily strikes/zone/GEX sign, same trust decision already made when the
repo itself was made public. Worth it only if you want a plain URL instead of the Sheets app.

## Outcomes & signal testing
The question this answers: do Max Pain / GEX actually predict anything for these contracts?
Only look at trading rules after this has a real sample.

- `outcomes.py` runs as part of every daily update. For each `Log` contract that has expired,
  it appends a row to an `Outcome` sheet in the Max Pain workbook: start price, Max Pain,
  NET GEX / Gamma Flip (joined from the GEX `Log`), the futures settle on expiry day, the
  period high/low, and the distances between them. **It has to run daily**: CME only keeps
  ~5 trading days, so a missed week loses those outcomes for good. A row with some days
  already gone is still stored, and `ข้อมูลครบ (วัน)` shows e.g. `1/2`.
- `python analyze.py` is read-only and prints the report. The key line compares Max Pain's
  error against the naive "price stays where it was" guess. **If Max Pain can't beat
  doing nothing, it isn't a signal.** It also splits realized range by NET GEX sign and
  results by the workbook's zone buckets. Below 30 samples it prints a warning, and the
  numbers are noise.
- Rows logged before the futures month was derived automatically are marked `เดา: <month>`
  in the last column. Their start price used the old OI>10,000 guess (OCT 26 rather than
  the DEC 26 those weeklies really settle against), so `analyze.py` leaves them out.
- Caveats: the "expiry price" is the futures **settle** on expiry day, not the option's exact
  expiry-time price. Exchange holidays aren't excluded from the day count. Standard monthly
  (AME) contracts aren't recorded yet.

## Paper trades
`analyze.py` also paper-trades a few rules on every recorded outcome. No orders are placed anywhere; this only builds evidence. The rules live in `rules.py` and were **fixed on 2026-10-03, before any results were seen**. Changing a rule makes its earlier results in-sample, so note the change date and count the out-of-sample record from there.

| Rule | When | Trade | Stop |
|---|---|---|---|
| R1 Range fade | Positive GEX and entry strictly between Put Wall and Call Wall | long in the lower half, short in the upper half | the wall behind the trade |
| R2 Momentum | Negative GEX | in the direction of the data day's futures change | the wall behind the trade, if any |
| R3 Max Pain magnet | any mode, Max Pain ≥ 0.5% from entry | toward Max Pain | none |
| Baseline | always | long | none |

Every rule has to beat the baseline to mean anything.

Assumptions:
- **Entry and exit:** entry is the **open of the first trading day after the data date**, since day T's CME data only exists after T closes. Exit is the futures settle on expiry day. P&L is in futures points.
- **Stops are judged worst-case.** Only the period's daily high/low are known, not the order things happened. If the stop level is anywhere inside high..low, the trade counts as stopped at the stop, even if price recovered.
- **"ไม่มีโหมด" (no mode):** conviction (V.4.1's J18) below 5% means no R1/R2 trade.
- **Where the levels come from:** `update_gex.levels()` computes Call Wall / Put Wall / Pin / conviction in Python, mirroring V.4.1's J9/J10/J30/J18 (openpyxl can't read Excel's results). `verify.py` checks each against Excel every time, and `GEX Log` stores them.

## Backfill
```
python backfill.py            # Max Pain workbook, every trade date CME still has
python backfill.py --gex      # also the GEX workbook (needs the QuikStrike login)
```
Only appends to `Log` (never touches `OI Data`/`Gamma Data`, which only ever hold the latest day) and skips dates already logged, so it's safe to run repeatedly. CME's own trade-date history is a short rolling window - currently ~5 trading days - so this catches up a few missed days, it doesn't build long-run history. Each day picks whatever contract CME's date-matching would pick for a `target` of *that* day (mirrors `target: today`, anchored to the day being backfilled); a day can still come back empty (a contract settling with zero OI that same day, or QuikStrike not carrying that historical (code, date) pair) - those are skipped with a one-line reason, not fatal.

## How it works
- `fetch_cme.py` - CME blocks plain HTTP and headless browsers (403 / HTTP2 errors), so it drives a real, visible Chrome via Playwright and calls the same JSON endpoints the CME page uses. Picks the contract, sums OI per strike, reads futures settle. `_underlying_future_month()` asks CME's options-quotes page which futures contract (e.g. `GCZ6`) each option series actually settles against - no login, no guessing - and `_settle_price()` uses that month unless `price_month` overrides it.
- `fetch_gamma.py` - drives the QuikStrike UI in the persistent Chrome profile: Metals -> Gold -> Greek "Gamma (1 Pct)" -> Strikes "(All)" -> expiration, then reads the matrix table. One column pair (C/P) per trade date, so past days can be read too. The two dropdown steps are verified and retried individually (`_step`); the whole fetch (product/expiration popups, which have no cheap readback) is retried up to 3 times from a fresh page (`fetch_gamma`) if anything fails - except a missing/expired login (`LoginRequired`), which isn't retried.
- `update_workbooks.py` - writes `OI Data`, extends formulas to 200 rows on first run, appends `Log`. Calls `update_gex.py` last; a gamma failure never loses the Max Pain update.
- `update_gex.py` - writes `Gamma Data` and the price into `GEX Calc!J6`, widens V.4.1's ranges if needed (`extend_layout`, idempotent), computes Gamma Flip and the J5-style mode (`gex_status`), appends `GEX Log`. It alerts only when the mode flips between Positive and Negative ("ไม่มีโหมด" never alerts).
- `rules.py` - the paper-trading rules and their stats, see **Paper trades**.
- `charts.py` - openpyxl drops every chart on save. `rebuild_charts()` restores the Max Pain line chart and the V.4.1 bar chart and is called before every save of the workbook (daily run, outcomes, backfill).
- `merge_workbooks.py` - the one-time build of `MaxPain_GEX.xlsx` described at the top.
- `util.py` - `save_atomic`: writes a temp file then replaces, so a crash or a workbook open in Excel never corrupts the real file.
- `verify.py` - re-fetches and compares; recomputes Max Pain by brute force and compares with Excel's own result. Also independently recomputes Gamma Flip (via `update_gex.select_rows`/`gamma_flip` against freshly re-fetched gamma) rather than trusting the stored `Log` value.
- `run_daily.bat` - Task Scheduler entry point; wraps a run in `run.log` with an `OK`/`WARN`/`FAIL` summary line (uses `setlocal enabledelayedexpansion` / `!errorlevel!` deliberately - the plain `%errorlevel%` form reads stale values inside a parenthesized `if` block).
- `notify.py` - `notify(title, message)`: a Windows popup via `msg.exe` (built in, no extra package), deliberately intrusive since these are same-day "go look" alerts, not routine status. Falls back to printing (so `run.log` still has it) if the popup can't show. Three triggers, each comparing the new `Log` row against the one it replaces so a rerun never re-alerts on an unchanged state: GEX being skipped (`update_workbooks.py`, e.g. an expired QuikStrike login), the Max Pain zone changing (`update_workbooks.py::zone()`, using the sheet's own F9/F10 thresholds), and NET GEX flipping sign (`update_gex.py`).
- `backfill.py` - loops `fetch_cme.fetch`/`fetch_gamma.fetch_gamma` over whatever trade dates CME still has, appending `Log` rows only; reuses `max_pain()`/`gamma_flip()`/`select_rows()` from the daily scripts rather than recomputing independently.
- `outcomes.py` / `analyze.py` - see **Outcomes & signal testing**. `fetch_cme.cme_session()` is the shared headed-Chrome session both `fetch()` and `outcomes.py` use; `fetch_cme.futures_row()` returns the full CME settlement row (open/high/low/settle) that `_settle_price()` also reads from.
- `publish_report.py` - reads both `Log` sheets read-only and renders `docs/index.html` for the mobile/web view (see **Mobile/web view** above). Only writes the local file; publishing it is a manual `git push`, deliberately not automated.
- `sheets_sync.py` - reuses `publish_report._read_log()` to get each workbook's latest `Log` row, then writes it into its own tab of a private Google Sheet via `gspread`, replacing that day's row in place on a rerun rather than duplicating it. Runs automatically every day (see **Mobile/web view**); skips quietly (with a notification) until `google_sheet_id`/`service_account.json` are set up.

## Known limits / ideas for next steps
- Latest trade date is usually **PRELIMINARY**; OI can change when CME publishes FINAL. Re-run next day.
- QuikStrike gamma values are rounded integers, and Net GEX = Call - Put assumes dealers long calls / short puts (a proxy, not real positioning).
- `_underlying_future_month()` only covers CME's rolling ~4-expirations-per-weekday window; older contracts (e.g. `verify.py` re-checking a past date) fall back to the "first month with OI > 10,000" heuristic, which can be wrong the same way the old default always was.
- Opening/saving with openpyxl drops existing charts; both scripts rebuild them.
- `verify.py` was tested against deliberately corrupted copies (wrong OI / price / gamma / Gamma Flip) and reported FAIL with exit code 1. An OI difference is only downgraded to WARN when CME's report changed PRELIMINARY -> FINAL since the data was stored.
- `notify.py` needs an interactive desktop session (same requirement Chrome already has); it prints instead of popping up if that's unavailable. Untested on Windows Home (`msg.exe` may not ship there).
- `backfill.py` can legitimately come back empty for a day - a contract that settled with zero OI that same day, or QuikStrike not carrying that historical (code, date) pair - this was observed live (2026-09-21/22) and isn't a bug, just what CME/QuikStrike have.
- The test suite (see **Tests**) only covers pure logic; the live scraping paths still rely on manual verification (`verify.py`, and the ad-hoc corrupted-copy / mocked-failure checks used while building each feature this session) - a colleague extending `fetch()` or `fetch_gamma()` won't get CI feedback on whether the actual scraping still works, only on the logic around it.
- Single underlying (Gold/GC) only; `underlying_product_id`/`qs_product` would need to become per-workbook to support Silver/Platinum.
- Single underlying (Gold/GC) only; `underlying_product_id`/`qs_product` would need to become per-workbook to support Silver/Platinum.
