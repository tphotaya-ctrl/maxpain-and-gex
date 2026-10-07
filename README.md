# Max Pain / GEX daily updater (CME Gold weekly options)

Fills two Excel workbooks automatically from CME Group data, then logs each day in a `Log` sheet.

| Workbook | Data written | Source |
|---|---|---|
| `คำนวน_max_pain_10 (2).xlsx` | Call/Put OI per strike, CME totals, DTE, futures settle price | CME "Volume & OI" JSON endpoints |
| `GEX V.4.1.xlsx` | Call/Put Gamma (1 Pct) per strike | CME QuikStrike, OI Matrix view (needs CME login) |

`GEX V.4.1.xlsx` replaces the older `คำนวน GEX 1.xlsx` (kept in the repo for reference only,
no longer wired into `config.json`). Its `GEX Calc` sheet is a richer, hand-built layout
(Call Wall/Put Wall, zones, Pin score, up to column AB) whose formulas already cover a fixed
`Gamma Data` range (`$2:$152`) and guard against blanks - so `update_gex.py` only ever
overwrites `Gamma Data`'s raw values, never `GEX Calc`'s formulas/conditional
formatting/chart. The script owns only a few `GEX Calc` cells: `I1`/`J1` (data date + contract),
`J6` (price: the price future's live quote at run time - CME, 10 min delayed - so walls,
pins and distances match the live chart; `"gex_price_source": "settle"` uses the previous
session's settle instead, `"gex_fill_price": false` leaves J6 to be typed by hand), and the chart's row range.

Max Pain and GEX results are calculated by the **Excel formulas** in the workbooks. Python only fills the input sheets (`OI Data`, `Gamma Data`) and appends to `Log`.

## Requirements
- Windows, Google Chrome installed, Python 3.10+, Microsoft Excel (only `verify.py` uses it, via COM)
- `pip install -r requirements.txt` (or `-r requirements-dev.txt` to also get `pytest`)
- A CME Group account (free) for the gamma part

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

## Phone (Telegram) - once
1. In Telegram open **@BotFather**, send `/newbot`, follow the steps, copy the token.
2. `python telegram_report.py --setup` - paste the token, then press **Start** on the bot in Telegram.

This writes `telegram.json` (token + chat id, git-ignored - never commit or share it). From then on
every run that gets new data sends one album: a colour summary card (mode, NET GEX, key levels
sorted by strike with distance from price, OI / Max Pain tables; HTML rendered by headless Chrome
so Thai text shapes correctly), the gamma chart and the Max Pain chart, with a one-line caption, and every `notify()` alert (GEX skipped,
zone change, NET GEX sign flip, a workbook left open, watchdog) is copied to the chat. Walls/Pins/mode
are read back from Excel's own calculation of the saved `GEX V.4.1.xlsx`; if that file couldn't be
saved (open in Excel) the report still goes out with Python totals only. `SKIP` runs send nothing.

## Daily run
Runs automatically: two Windows Task Scheduler tasks, **"MaxPainGEX Daily Update"** (weekdays 09:00 Bangkok) and **"MaxPainGEX Afternoon Update"** (14:00 - CME/QuikStrike often haven't published the previous day by 09:00; this run exits early with `SKIP` via `--if-new` when the morning one already logged the trade date), as the logged-in user (must stay logged in - the browser runs headed, Chrome will visibly pop up). It runs `run_daily.bat`, which appends a timestamped block to `run.log` ending in one summary line - `OK`, `WARN - ... GEX skipped` (Max Pain still updated), or `FAIL`. Re-create the task with:
```
schtasks /create /tn "MaxPainGEX Daily Update" /tr "\"<repo path>\run_daily.bat\"" /sc weekly /d MON,TUE,WED,THU,FRI /st 09:00 /rl limited /f
schtasks /create /tn "MaxPainGEX Afternoon Update" /tr "\"<repo path>\run_daily.bat\"" /sc weekly /d MON,TUE,WED,THU,FRI /st 14:00 /rl limited /f
```
Manual run:
```
python update_workbooks.py        # or run_daily.bat (appends to run.log; adds --if-new)
python update_workbooks.py --if-new   # exit early if both Logs already hold this trade date
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
| `workbook`, `gex_workbook` | files to fill |

Set env `MAXPAIN_CONFIG=path\to\other.json` to work on copies without touching the originals.

## Backfill
```
python backfill.py            # Max Pain workbook, every trade date CME still has
python backfill.py --gex      # also the GEX workbook (needs the QuikStrike login)
```
Only appends to `Log` (never touches `OI Data`/`Gamma Data`, which only ever hold the latest day) and skips dates already logged, so it's safe to run repeatedly. CME's own trade-date history is a short rolling window - currently ~5 trading days - so this catches up a few missed days, it doesn't build long-run history. Each day picks whatever contract CME's date-matching would pick for a `target` of *that* day (mirrors `target: today`, anchored to the day being backfilled); a day can still come back empty (a contract settling with zero OI that same day, or QuikStrike not carrying that historical (code, date) pair) - those are skipped with a one-line reason, not fatal.

## How it works
- `fetch_cme.py` - CME blocks plain HTTP and headless browsers (403 / HTTP2 errors), so it drives a real, visible Chrome via Playwright and calls the same JSON endpoints the CME page uses. Picks the contract, sums OI per strike, reads futures settle. `_underlying_future_month()` asks CME's options-quotes page which futures contract (e.g. `GCZ6`) each option series actually settles against - no login, no guessing - and `_settle_price()` uses that month unless `price_month` overrides it.
- `fetch_gamma.py` - drives the QuikStrike UI in the persistent Chrome profile: Metals -> Gold -> Greek "Gamma (1 Pct)" -> Strikes "(All)" -> expiration, then reads the matrix table. One column pair (C/P) per trade date, so past days can be read too. The two dropdown steps are verified and retried individually (`_step`); the whole fetch (product/expiration popups, which have no cheap readback) is retried up to 3 times from a fresh page (`fetch_gamma`) if anything fails - except a missing/expired login (`LoginRequired`), which isn't retried.
- `update_workbooks.py` - writes `OI Data`, extends formulas to 200 rows on first run, rebuilds the chart, appends `Log`. Calls `update_gex.py` last; a gamma failure never loses the Max Pain update.
- `update_gex.py` - writes `Gamma Data` raw values only (`GEX Calc`'s formulas/chart/conditional formatting are fixed already, see above), computes Gamma Flip, appends `Log`.
- `util.py` - `save_atomic`: writes a temp file then replaces, so a crash or a workbook open in Excel never corrupts the real file.
- `verify.py` - re-fetches and compares; recomputes Max Pain by brute force and compares with Excel's own result. Also independently recomputes Gamma Flip (via `update_gex.select_rows`/`gamma_flip` against freshly re-fetched gamma) rather than trusting the stored `Log` value.
- `run_daily.bat` - Task Scheduler entry point; wraps a run in `run.log` with an `OK`/`WARN`/`FAIL` summary line (uses `setlocal enabledelayedexpansion` / `!errorlevel!` deliberately - the plain `%errorlevel%` form reads stale values inside a parenthesized `if` block).
- `notify.py` - `notify(title, message)`: a Windows popup via `msg.exe`, or a PowerShell `WScript.Shell` popup where `msg.exe` doesn't exist (Windows Home), deliberately intrusive since these are same-day "go look" alerts, not routine status. Falls back to printing (so `run.log` still has it) if the popup can't show. Three triggers, each comparing the new `Log` row against the one it replaces so a rerun never re-alerts on an unchanged state: GEX being skipped (`update_workbooks.py`, e.g. an expired QuikStrike login), the Max Pain zone changing (`update_workbooks.py::zone()`, using the sheet's own F9/F10 thresholds), and NET GEX flipping sign (`update_gex.py`).
- `backfill.py` - loops `fetch_cme.fetch`/`fetch_gamma.fetch_gamma` over whatever trade dates CME still has, appending `Log` rows only; reuses `max_pain()`/`gamma_flip()`/`select_rows()` from the daily scripts rather than recomputing independently.

## Known limits / ideas for next steps
- Latest trade date is usually **PRELIMINARY**; OI can change when CME publishes FINAL. Re-run next day.
- QuikStrike gamma values are rounded integers, and Net GEX = Call - Put assumes dealers long calls / short puts (a proxy, not real positioning).
- `_underlying_future_month()` only covers CME's rolling ~4-expirations-per-weekday window; older contracts (e.g. `verify.py` re-checking a past date) fall back to the "first month with OI > 10,000" heuristic, which can be wrong the same way the old default always was.
- Opening/saving with openpyxl drops existing charts; both scripts rebuild them.
- `verify.py` was tested against deliberately corrupted copies (wrong OI / price / gamma / Gamma Flip) and reported FAIL with exit code 1. An OI difference is only downgraded to WARN when CME's report changed PRELIMINARY -> FINAL since the data was stored.
- `notify.py` needs an interactive desktop session (same requirement Chrome already has); it prints instead of popping up if that's unavailable.
- Every run has a watchdog (`watchdog_minutes` in `config.json`, default 15) that aborts with a `FAIL` line, and every CME JSON call times out after 30s - added after `Browser.close()` was seen hanging forever on Playwright 1.61 / Windows 11 26200 (`fetch_cme.py` now uses a throwaway persistent context instead).
- `backfill.py` can legitimately come back empty for a day - a contract that settled with zero OI that same day, or QuikStrike not carrying that historical (code, date) pair - this was observed live (2026-09-21/22) and isn't a bug, just what CME/QuikStrike have.
- The test suite (see **Tests**) only covers pure logic; the live scraping paths still rely on manual verification (`verify.py`, and the ad-hoc corrupted-copy / mocked-failure checks used while building each feature this session) - a colleague extending `fetch()` or `fetch_gamma()` won't get CI feedback on whether the actual scraping still works, only on the logic around it.
- Single underlying (Gold/GC) only; `underlying_product_id`/`qs_product` would need to become per-workbook to support Silver/Platinum.
- Single underlying (Gold/GC) only; `underlying_product_id`/`qs_product` would need to become per-workbook to support Silver/Platinum.
