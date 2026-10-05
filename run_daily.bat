@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set OUT=%TEMP%\maxpain_run_%RANDOM%.txt

echo ============================== >> run.log
echo [%date% %time%] start >> run.log
rem -u: unbuffered, so a run killed by the watchdog still leaves its progress in run.log
rem --if-new: the 14:00 re-run exits early when the 09:00 run already logged this trade date
python -u update_workbooks.py --if-new > "%OUT%" 2>&1
set EXITCODE=%errorlevel%
type "%OUT%" >> run.log

rem one-line summary so a failure is visible without reading the whole log.
rem uses delayed expansion (!errorlevel!) - plain %errorlevel% inside a
rem parenthesized block would keep the value from when the block was parsed,
rem not the findstr that just ran, and silently print the wrong banner.
findstr /i /c:"Traceback" /c:"Cannot save" "%OUT%" >nul
if !errorlevel! equ 0 (
    echo [%date% %time%] FAIL - see above, exit !EXITCODE! >> run.log
) else (
    findstr /i /c:"GEX skipped" "%OUT%" >nul
    if !errorlevel! equ 0 (
        echo [%date% %time%] WARN - Max Pain OK, GEX skipped, exit !EXITCODE! >> run.log
    ) else if !EXITCODE! neq 0 (
        echo [%date% %time%] FAIL - exit !EXITCODE! >> run.log
    ) else (
        echo [%date% %time%] OK - exit !EXITCODE! >> run.log
    )
)
del "%OUT%" 2>nul
endlocal
