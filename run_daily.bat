@echo off
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
echo [%date% %time%] start >> run.log
python update_workbooks.py >> run.log 2>&1
echo [%date% %time%] exit %errorlevel% >> run.log
