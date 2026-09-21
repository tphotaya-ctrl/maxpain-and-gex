# Max Pain / GEX daily updater (CME Gold weekly options)

Fills two Excel workbooks automatically from CME Group data, then logs each day in a `Log` sheet.

| Workbook | Data written | Source |
|---|---|---|
| `คำนวน_max_pain_10 (2).xlsx` | Call/Put OI per strike, CME totals, DTE, futures settle price | CME "Volume & OI" JSON endpoints |
| `คำนวน GEX 1.xlsx` | Call/Put Gamma (1 Pct) per strike | CME QuikStrike, OI Matrix view (needs CME login) |

Max Pain and GEX results are calculated by the **Excel formulas** in the workbooks. Python only fills the input sheets (`OI Data`, `Gamma Data`) and appends to `Log`.

## Requirements
- Windows, Google Chrome installed, Python 3.10+, Microsoft Excel (only `verify.py` uses it, via COM)
- `pip install -r requirements.txt`
- A CME Group account (free) for the gamma part

## Setup (once)
```
python quikstrike_login.py
```
A Chrome window opens. Log in to CME yourself, open the QuikStrike view once, close the window. The session is kept in `.chrome_profile/` (git-ignored - it holds your login cookies, never commit or share it). Without a session the gamma step is skipped and the Max Pain step still runs.

## Daily run
```
python update_workbooks.py        # or run_daily.bat (appends to run.log)
python verify.py                  # independent re-check, exit code 1 on any FAIL
```
Close both workbooks in Excel first, otherwise saving fails.

`config.json`:

| key | meaning |
|---|---|
| `target` | `today` (contract expiring today, local date; next expiry if none), `YYYY-MM-DD`, `auto` (nearest of `family`), or a label like `Week 4 - SEP 2026` |
| `family` | weekly product used by `auto`/label: `MW1` Mon, `AB1` Tue, `WD1` Wed, `BB1` Thu, `E21` Fri |
| `price_month` | optional, e.g. `DEC 26`: futures month used as the price (default: first month with OI > 10,000). See limits below |
| `strike_min` / `strike_max` | strike window written to `OI Data` (max 200 rows) |
| `workbook`, `gex_workbook` | files to fill |

Set env `MAXPAIN_CONFIG=path\to\other.json` to work on copies without touching the originals.

## How it works
- `fetch_cme.py` - CME blocks plain HTTP and headless browsers (403 / HTTP2 errors), so it drives a real, visible Chrome via Playwright and calls the same JSON endpoints the CME page uses. Picks the contract, sums OI per strike, reads futures settle.
- `fetch_gamma.py` - drives the QuikStrike UI in the persistent Chrome profile: Metals -> Gold -> Greek "Gamma (1 Pct)" -> Strikes "(All)" -> expiration, then reads the matrix table. One column pair (C/P) per trade date, so past days can be read too.
- `update_workbooks.py` - writes `OI Data`, extends formulas to 200 rows on first run, rebuilds the chart, appends `Log`. Calls `update_gex.py` last; a gamma failure never loses the Max Pain update.
- `update_gex.py` - writes `Gamma Data`, fixes `GEX Calc` ranges, rebuilds the chart, computes Gamma Flip, appends `Log`.
- `util.py` - `save_atomic`: writes a temp file then replaces, so a crash or a workbook open in Excel never corrupts the real file.
- `verify.py` - re-fetches and compares; recomputes Max Pain by brute force and compares with Excel's own result.

## Known limits / ideas for next steps
- Latest trade date is usually **PRELIMINARY**; OI can change when CME publishes FINAL. Re-run next day.
- QuikStrike gamma values are rounded integers, and Net GEX = Call - Put assumes dealers long calls / short puts (a proxy, not real positioning).
- QuikStrike is driven by fixed waits and element ids (ASP.NET postbacks) - it can time out or break if CME changes the page. Adding retries would help.
- Monthly (non-weekly) expirations are not supported in `target: today`.
- Task Scheduler entry for a daily run is not set up yet (suggested: ~08:30 Bangkok time, machine logged in, Chrome window visible).
- Opening/saving with openpyxl drops existing charts; both scripts rebuild them.
- **Price = futures month may not be the option's real underlying.** In QuikStrike, weeklies expiring before the Oct standard option use GCV6 (Oct) but later ones use GCZ6 (Dec). The default price uses the first liquid month; the `Log` sheet records which month (`Futures ที่ใช้เป็นราคา`). Set `price_month` when a weekly is past the standard option expiry. The exact CME rule is not derived in code yet.
- `verify.py` was tested against deliberately corrupted copies (wrong OI / price / gamma) and reported FAIL with exit code 1. An OI difference is only downgraded to WARN when CME's report changed PRELIMINARY -> FINAL since the data was stored.
