# Max Pain / GEX daily updater (CME Gold weekly options)

Fills one Excel workbook, `MaxPain_GEX.xlsx`, automatically from CME Group data. It holds both Max Pain and GEX (calculator GEX V.4.1), and each day is logged.

| Sheet | Filled by | Source |
|---|---|---|
| `OI Data` → `Max Pain Calc` | Call/Put OI per strike, CME totals, DTE; futures settle into `Max Pain Calc!F4` | CME "Volume & OI" JSON endpoints |
| `Log` | one row per day: Max Pain, distance, P/C, ... | |
| `Gamma Data` → `GEX Calc` | Call/Put Gamma (1 Pct) per strike (rows 2-152); the reference price into `GEX Calc!J6` (live futures quote at run time, see `gex_price_source`), data date/contract into `I1`/`J1` | CME QuikStrike, OI Matrix view (needs CME login) |
| `GEX Log` | one row per day: totals, NET GEX, mode, Gamma Flip, peak | |
| `Log รวม` | one row per day: GEX over **every expiration** QuikStrike lists by default (nearest weekly + monthlies), see `update_gex_all` | QuikStrike expiry matrix |
| `บันทึกโหมด` | **manual** - your own daily journal, never written by the scripts | |
| `Outcome` | created by `outcomes.py`, see below | CME futures settlements |

Max Pain and GEX results are calculated by the **Excel formulas**. Python only fills the input sheets and the logs.

`MaxPain_GEX.xlsx` was built once by `python merge_workbooks.py` from the three files it replaced: `คำนวน_max_pain_10 (2).xlsx`, `GEX V.4.1.xlsx`, and `คำนวน GEX 1.xlsx` (only its `Log`, carried over as `GEX Log`). The originals are kept as they were. The merge copies sheets with Excel itself, so charts, dropdowns and array formulas survive. It also widens V.4.1's ranges to rows 2-152:
- `GEX Calc` A:F use the blank-if-empty form on every row (rows 2-66 used to show 0 for an empty strike, which pulled MEDIAN down).
- E/F's MEDIAN, the H helper and the conditional formatting now reach row 152.
- J16 checks the last filled strike instead of row 62.
- J12 (the median gross that E/F's strength/zone compare against) counts only strikes carrying gamma within ±J26 of the price, and E/F read J12 - over every row that median was 0 once most strikes were empty, so every row read จุดหนืด (`update_gex.fix_zone_median`, idempotent, applied on the next run).
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
Runs automatically from the Windows Task Scheduler task **"MaxPainGEX Daily Update"**, as the logged-in user. You must be logged in, because the browser runs headed and Chrome visibly pops up.

**The PC doesn't have to be on all the time.** The task fires:
- weekdays 09:00, 11:00 and 14:00 Bangkok time - CME/QuikStrike often haven't published the previous US session by 09:00, the later triggers pick it up
- if the PC was off at a trigger, as soon as it's on (StartWhenAvailable)
- at every logon, 2 minutes after

`run_state.py` keeps that to one real run per day (`last_ok.txt` = today + the trade date logged), but only once that run ended **OK** with the previous weekday's data. A stale, WARN (GEX skipped, workbook open) or FAILed day lets the next trigger run again. That retry is what rescues a day's GEX: the day's contract expires the next day, and after that QuikStrike no longer lists it, so catch-up can't get it back. Each run uses `--if-new`, so a run that finds nothing new is a quick `SKIP`. The phone card shows an orange ⏳ banner while the data is older than the previous weekday (the day after a US holiday can show it falsely).

Each run also **catches up**: `backfill.catch_up` logs any trade date CME still has that's missing from either Log, so outcomes can still be recorded later. CME only keeps ~5 trading days, so **turn the PC on at least once every ~4 trading days** or those days are gone for good. Register or re-register the task with:
```
powershell -ExecutionPolicy Bypass -File install_task.ps1
```
`run_daily.bat` appends a timestamped block to `run.log` (rotated to `run.log.1` past 1 MB), ending in one summary line: `OK`; `WARN` (finished, but a workbook was open in Excel or GEX was skipped - each step carries on and the phone report still goes out); or `FAIL` (crashed, watchdog, or non-zero exit). Every run has a watchdog (`watchdog_minutes`, default 20).

### Phone alerts (Telegram) and a "hasn't run" alarm
Optional, free. Both secrets go in **`secrets.json`** (git-ignored; the repo is public, so never put them in `config.json`):
```json
{"telegram_token": "123456:ABC...", "telegram_chat_id": "123456789",
 "healthcheck_url": "https://hc-ping.com/your-uuid"}
```
- **Telegram:**
  - In Telegram, talk to **@BotFather** → `/newbot` → copy the token, then run `python telegram_report.py --setup`: paste the token, press **Start** on the bot; it fills `telegram_token`/`telegram_chat_id` in `secrets.json` (keeping its other keys). A legacy `telegram.json` is still read.
  - Every alert also reaches your phone: GEX skipped, Max Pain zone change, GEX mode flip, a workbook left open in Excel, a watchdog abort, a failed run, any step skipped.
  - After each run that got new data you get **one album** (`telegram_report.py`): a colour summary card - mode, NET GEX/conviction, key levels sorted by strike with the price row slotted in and distances, densest gamma near the price, the all-expiration GEX, OI / Max Pain tables, and which paper rule (`rules.py`) is armed for the next session - plus the gamma chart and the Max Pain chart. The card is HTML rendered by headless Chrome (matplotlib can't shape Thai); Walls/Pins/mode come from Excel's own calculation of the saved workbook (`verify.excel_cells`). `daily_summary.py`'s text is still printed to `run.log`, and is the Telegram fallback if the card can't be rendered.
  - Test with `python notify.py Test "hello"`.
- **"Hasn't run" alarm:** create a free check at healthchecks.io and paste its ping URL.
  - Set the schedule to cron `0 9 * * 1-5`, timezone Asia/Bangkok, grace ~6 h (past the 14:00 trigger), and connect its Telegram/email integration.
  - Every run pings it (`/fail` on FAIL). If no ping arrives, *healthchecks.io* alerts you, which works even while the PC is off.

### Would it work in the cloud?
`cloud_probe.py` checks whether CME's data endpoints answer from a given machine. It tries plain HTTP, headless Chromium, and the headed Chrome the daily job uses. Run it from GitHub's servers via **Actions → "Cloud probe (CME reachability)" → Run workflow**; the result shows on the run page. It doesn't test QuikStrike, which needs the CME login session, so even a pass only means the Max Pain half could move.
Manual run:
```
python update_workbooks.py        # or run_daily.bat (appends to run.log; adds --if-new)
python update_workbooks.py --if-new   # exit early with SKIP if both Logs already hold this trade date
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
| `gex_price_source` | `live` (default): `GEX Calc!J6`, the GEX Log price and walls/pins/distances use CME's 10-min delayed quote for the price future at run time, so they match a live chart (it can sit tens of dollars off the previous settle); `settle` uses the previous session's settle. The Max Pain sheet always uses the settle |
| `gex_fill_price` | default `true`; `false` leaves `GEX Calc!J6` to be typed by hand |
| `watchdog_minutes` | default 20: a run longer than this aborts as FAIL |
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
- Each row also carries the **all-expiration GEX** for the same data date, from `Log รวม`:
  `โหมด (รวม)`, `ความชัดเจน (รวม)`, `Call Wall (รวม)`, `Put Wall (รวม)`. `analyze.py` shows
  realized range by mode for both readings, single contract and รวม, to see which one
  tells volatility apart better.
- Holidays: the `ข้อมูลครบ` day count leaves out a weekday that falls inside CME's own
  trade-date list but has no trade date there, i.e. an exchange holiday. It comes from CME's
  data, so there's no holiday table to maintain.
- Standard monthly (AME) contracts are recorded too. Their expiry comes from CME's
  options-quotes page, which only lists upcoming contracts, so only rows from the last ~60
  days are tried. There's no single-contract GEX for them (no QuikStrike code mapping).
- Caveat: the "expiry price" is the futures **settle** on expiry day, not the option's exact
  expiry-time price.

### Weekly report
`weekly_report.py` sends one short Telegram message a week, from Friday's run. If Friday
never ran, it goes out with the next week's first run, covering the week before. It contains:
- the sample so far against the 30 needed
- Max Pain's error against the naive guess
- each paper rule's running total against the baseline
- how many days that week actually got GEX

It's the same numbers as `analyze.py`. `weekly_sent.txt` (git-ignored) keeps it to once a
week; it's only written after Telegram confirms delivery. `python weekly_report.py --force`
sends it now.

## Paper trades
`analyze.py` also paper-trades a few rules on every recorded outcome. No orders are placed anywhere; this only builds evidence. The rules live in `rules.py` and were **fixed on 2026-10-03, before any results were seen**. Changing a rule makes its earlier results in-sample, so note the change date and count the out-of-sample record from there.

| Rule | When | Trade | Stop |
|---|---|---|---|
| R1 Range fade | Positive GEX and entry strictly between Put Wall and Call Wall | long in the lower half, short in the upper half | the wall behind the trade |
| R2 Momentum | Negative GEX | in the direction of the data day's futures change | the wall behind the trade, if any |
| R3 Max Pain magnet | any mode, Max Pain ≥ 0.5% from entry | toward Max Pain | none |
| R1-all / R2-all | as R1 / R2, but judged on the **all-expiration** GEX (`Log รวม`) | same | same |
| Baseline | always | long | none |

Every rule has to beat the baseline to mean anything. R1-all/R2-all were **added on
2026-10-10**, so their record starts there; they reuse R1/R2's logic unchanged, fed the
rolled-up mode and walls.

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
- `fetch_cme.py` - CME blocks plain HTTP and headless browsers (403 / HTTP2 errors), so it drives a real, visible Chrome via Playwright and calls the same JSON endpoints the CME page uses. Picks the contract, sums OI per strike, reads futures settle. `_underlying_future_month()` asks CME's options-quotes page which futures contract (e.g. `GCZ6`) each option series actually settles against - no login, no guessing - and `_settle_price()` uses that month unless `price_month` overrides it. `cme_session()` uses a throwaway *persistent* context (`util.launch_persistent`, which also retries a Chrome that exits at launch) because `Browser.close()` hangs forever on the Windows 11 26200 PC, and races every JSON call against a 30 s timeout. A trade date CME has only just started publishing is often incomplete (no expirations, zero OI, today's weekly missing), so `fetch()` tries the latest two dates, exact-expiry match first, and takes the *next* expiry only as a last resort. `_live_quote()` adds CME's 10-min delayed quote for the price future (`quotes/v2`).
- `fetch_gamma.py` - drives the QuikStrike UI in the persistent Chrome profile: Metals -> Gold -> Greek "Gamma (1 Pct)" -> Strikes "(All)" -> expiration, then reads the matrix table. One column pair (C/P) per trade date, so past days can be read too. The two dropdown steps are verified and retried individually (`_step`); the whole fetch (product/expiration popups, which have no cheap readback) is retried up to 3 times from a fresh page (`fetch_gamma`) if anything fails - except a missing/expired login (`LoginRequired`), which isn't retried. When the per-expiration history lags (it can sit a day behind), the column is read from the default expiry matrix instead, but only if that matrix shows the wanted trade date. `fetch_gamma_all()` reads every column of that matrix for the all-expiration GEX. Waits poll for the rendered table instead of fixed sleeps. When the per-expiration path fails, `save_diag()` writes a full-page screenshot and the frame's text to `diag/` (git-ignored, newest 20 kept), so the failure can be diagnosed from what QuikStrike actually showed.
- `update_workbooks.py` - writes `OI Data`, extends formulas to 200 rows on first run, appends `Log`. Calls `update_gex.py` last; a gamma failure never loses the Max Pain update.
- `update_gex.py` - writes `Gamma Data` and the price into `GEX Calc!J6`, widens V.4.1's ranges if needed (`extend_layout`, idempotent), computes Gamma Flip and the J5-style mode (`gex_status`), appends `GEX Log`. It alerts only when the mode flips between Positive and Negative ("ไม่มีโหมด" never alerts). The price it writes and measures from is `ref_price()` (see `gex_price_source`); a workbook open in Excel no longer aborts the run. `update_gex_all()` sums gamma over every expiration and applies the same level rules (`agg_levels`), logging to `Log รวม`. The daily run does this **first**, and when the default matrix already contains the day's contract, `update_gex(cfg, d, g=...)` takes its table from there without clicking the EXPIRATION popup, the step that kept timing out (the run log says `source=matrix` / `source=per-expiration`). The default matrix only lists the nearest Friday weekly (OG...) and the monthlies, so this covers the Friday contract; Mon-Thu weeklies still use the popup; its Gamma Flip uses the cumulative-by-strike definition, so it is often absent - a true flip (net GEX re-priced at hypothetical spots) would need IV per strike.
- `rules.py` - the paper-trading rules and their stats, see **Paper trades**.
- `charts.py` - openpyxl drops every chart on save. `rebuild_charts()` restores the Max Pain line chart and the V.4.1 bar chart and is called before every save of the workbook (daily run, outcomes, backfill).
- `merge_workbooks.py` - the one-time build of `MaxPain_GEX.xlsx` described at the top.
- `util.py` - `save_atomic`: writes a temp file then replaces, so a crash or a workbook open in Excel never corrupts the real file. `launch_persistent`: Chrome launch with retries.
- `verify.py` - re-fetches and compares; recomputes Max Pain by brute force and compares with Excel's own result. Also independently recomputes Gamma Flip (via `update_gex.select_rows`/`gamma_flip` against freshly re-fetched gamma) rather than trusting the stored `Log` value. `excel_cells()` retries a `Workbooks.Open` that returns nothing (about 1 run in 6) and kills exactly the EXCEL.EXE it started, in `finally`.
- `run_daily.bat` - Task Scheduler entry point; wraps a run in `run.log` with an `OK`/`WARN`/`FAIL` summary line (uses `setlocal enabledelayedexpansion` / `!errorlevel!` deliberately - the plain `%errorlevel%` form reads stale values inside a parenthesized `if` block).
- `notify.py` - `notify(title, message)`: a Windows popup via `msg.exe`, or a PowerShell `WScript.Shell` popup where `msg.exe` doesn't exist (Windows Home), plus Telegram when `secrets.json` is set up; falls back to printing so `run.log` keeps it.
- `telegram_report.py` - the end-of-run album (card + charts), its `--setup`, and the stale-data check (`expected_trade_date`).
- `run_state.py` - one real run per day once the data is fresh, healthchecks.io ping, Telegram on FAIL.
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
- `notify.py`'s popup needs an interactive desktop session (same requirement Chrome already has); it prints instead of popping up if that's unavailable.
- `backfill.py` can legitimately come back empty for a day - a contract that settled with zero OI that same day, or QuikStrike not carrying that historical (code, date) pair - this was observed live (2026-09-21/22) and isn't a bug, just what CME/QuikStrike have.
- The test suite (see **Tests**) only covers pure logic; the live scraping paths still rely on manual verification (`verify.py`, and the ad-hoc corrupted-copy / mocked-failure checks used while building each feature this session) - a colleague extending `fetch()` or `fetch_gamma()` won't get CI feedback on whether the actual scraping still works, only on the logic around it.
- Single underlying (Gold/GC) only; `underlying_product_id`/`qs_product` would need to become per-workbook to support Silver/Platinum.
