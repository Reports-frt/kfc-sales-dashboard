@echo off
REM ============================================================
REM KFC Dashboard Daily Auto-Update
REM Trigger: Windows Task Scheduler (1x daily)
REM Runs hidden via pythonw.exe; logs to _work\update.log
REM ============================================================

cd /d "C:\Users\IT\Documents\GitHub\kfc-sales-dashboard"

REM Use pythonw.exe to suppress console window
pythonw.exe "C:\Users\IT\Documents\GitHub\kfc-sales-dashboard\automation\update_dashboard.py"

exit /b %ERRORLEVEL%
