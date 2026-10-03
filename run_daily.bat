@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8

rem the task fires at 08:30 and at logon (or late after a missed 08:30) - one real run a day
python run_state.py check
if !errorlevel! neq 0 (
    echo [%date% %time%] skip - already ran today >> run.log
    endlocal
    exit /b 0
)

set OUT=%TEMP%\maxpain_run_%RANDOM%.txt
echo ============================== >> run.log
echo [%date% %time%] start >> run.log
python update_workbooks.py > "%OUT%" 2>&1
set EXITCODE=%errorlevel%
type "%OUT%" >> run.log

rem one-line summary so a failure is visible without reading the whole log.
rem uses delayed expansion (!errorlevel!) - plain %errorlevel% inside a
rem parenthesized block would keep the value from when the block was parsed,
rem not the findstr that just ran, and silently print the wrong banner.
set STATUS=OK
findstr /i /c:"Traceback" /c:"Cannot save" "%OUT%" >nul
if !errorlevel! equ 0 (
    set STATUS=FAIL
) else (
    findstr /i /c:"GEX skipped" "%OUT%" >nul
    if !errorlevel! equ 0 (
        set STATUS=WARN
    ) else if !EXITCODE! neq 0 (
        set STATUS=FAIL
    )
)
if "!STATUS!"=="WARN" (
    echo [%date% %time%] WARN - Max Pain OK, GEX skipped, exit !EXITCODE! >> run.log
) else (
    echo [%date% %time%] !STATUS! - exit !EXITCODE! >> run.log
)
python run_state.py finish !STATUS! "%OUT%" >> run.log 2>&1
del "%OUT%" 2>nul
endlocal
