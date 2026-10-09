@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8

rem the task fires at 09:00, 11:00, 14:00 and at logon (or late after a missed trigger);
rem run_state.py skips once today has run OK/WARN with the previous weekday's data
python run_state.py check
if !errorlevel! neq 0 (
    echo [%date% %time%] skip - already ran today with fresh data >> run.log
    endlocal
    exit /b 0
)

rem keep run.log bounded: past 1 MB, the old log becomes run.log.1 (one generation kept)
if exist run.log for %%A in (run.log) do if %%~zA gtr 1000000 move /y run.log run.log.1 >nul

set OUT=%TEMP%\maxpain_run_%RANDOM%.txt
echo ============================== >> run.log
echo [%date% %time%] start >> run.log
rem -u: unbuffered, so a run killed by the watchdog still leaves its progress in run.log
rem --if-new: a re-run exits early with SKIP when this trade date is already in both Logs
python -u update_workbooks.py --if-new > "%OUT%" 2>&1
set EXITCODE=%errorlevel%
type "%OUT%" >> run.log

rem one-line summary so a failure is visible without reading the whole log.
rem FAIL = crashed (Traceback, incl. the watchdog) or exited non-zero; WARN = finished but a
rem workbook was open in Excel or GEX was skipped - the phone report still went out.
rem uses delayed expansion (!errorlevel!) - plain %errorlevel% inside a
rem parenthesized block would keep the value from when the block was parsed,
rem not the findstr that just ran, and silently print the wrong banner.
set STATUS=OK
findstr /i /c:"Traceback" "%OUT%" >nul
if !errorlevel! equ 0 (
    set STATUS=FAIL
) else if !EXITCODE! neq 0 (
    set STATUS=FAIL
) else (
    findstr /i /c:"Cannot save" /c:"GEX skipped" "%OUT%" >nul
    if !errorlevel! equ 0 set STATUS=WARN
)
if "!STATUS!"=="WARN" (
    echo [%date% %time%] WARN - workbook open in Excel or GEX skipped, see above >> run.log
) else (
    echo [%date% %time%] !STATUS! - exit !EXITCODE! >> run.log
)
python run_state.py finish !STATUS! "%OUT%" >> run.log 2>&1
del "%OUT%" 2>nul
endlocal
