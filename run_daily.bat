@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set OUT=%TEMP%\maxpain_run_%RANDOM%.txt

rem keep run.log bounded: past 1 MB, the old log becomes run.log.1 (one generation kept)
if exist run.log for %%A in (run.log) do if %%~zA gtr 1000000 move /y run.log run.log.1 >nul

echo ============================== >> run.log
echo [%date% %time%] start >> run.log
rem -u: unbuffered, so a run killed by the watchdog still leaves its progress in run.log
rem --if-new: the 11:00/14:00 re-runs exit early when an earlier run already logged this trade date
python -u update_workbooks.py --if-new > "%OUT%" 2>&1
set EXITCODE=%errorlevel%
type "%OUT%" >> run.log

rem one-line summary so a failure is visible without reading the whole log.
rem FAIL = the run crashed (Traceback, incl. the watchdog) or exited non-zero;
rem WARN = it finished but a workbook was open in Excel or GEX was skipped - the
rem phone report still went out, so that isn't a failure of the run.
rem uses delayed expansion (!errorlevel!) - plain %errorlevel% inside a
rem parenthesized block would keep the value from when the block was parsed,
rem not the findstr that just ran, and silently print the wrong banner.
findstr /i /c:"Traceback" "%OUT%" >nul
if !errorlevel! equ 0 (
    echo [%date% %time%] FAIL - see above, exit !EXITCODE! >> run.log
) else if !EXITCODE! neq 0 (
    echo [%date% %time%] FAIL - exit !EXITCODE! >> run.log
) else (
    findstr /i /c:"Cannot save" /c:"GEX skipped" "%OUT%" >nul
    if !errorlevel! equ 0 (
        echo [%date% %time%] WARN - workbook open in Excel or GEX skipped, see above >> run.log
    ) else (
        echo [%date% %time%] OK - exit !EXITCODE! >> run.log
    )
)
del "%OUT%" 2>nul
endlocal
