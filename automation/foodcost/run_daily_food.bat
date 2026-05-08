@echo off
REM =====================================================================
REM KFC Food Cost — Daily Update (v2: strict mode)
REM =====================================================================
REM Steps:
REM   1. Pull 3 latest emails from Outlook (TODAY's only)
REM      - If any missing → STOP, send notification, no build/push
REM   2. If all pulled → build food_data.json
REM   3. Push to GitHub
REM
REM Run manually: double-click this file
REM Schedule via setup_food_task.ps1 to run daily at 11:15
REM =====================================================================

cd /d "%~dp0"
echo.
echo ============================================
echo KFC Food Cost - Daily Update (STRICT)
echo ============================================
echo.

REM Step 1: Pull emails (strict — today only)
echo === STEP 1: Pull today's emails from Outlook ===
"C:\Program Files\Python312\python.exe" pull_food_emails.py
set PULL_RESULT=%ERRORLEVEL%

if %PULL_RESULT% NEQ 0 (
    echo.
    echo ============================================
    echo PIPELINE STOPPED - emails missing
    echo ============================================
    echo.
    echo Check Outlook for notification email.
    echo Run manually when emails arrive:
    echo   .\run_daily_food.bat
    echo.
    echo Or run in lenient mode to use yesterday's files:
    echo   "C:\Program Files\Python312\python.exe" pull_food_emails.py --lenient
    echo   "C:\Program Files\Python312\python.exe" build_pipeline.py
    echo.
    REM exit with error so Task Scheduler logs the failure
    exit /b 1
)

REM Step 2 + 3: Build + Push
echo.
echo === STEP 2-3: Build food_data.json + Push to GitHub ===
"C:\Program Files\Python312\python.exe" build_pipeline.py
set BUILD_RESULT=%ERRORLEVEL%

echo.
if %BUILD_RESULT% NEQ 0 (
    echo ============================================
    echo BUILD/PUSH FAILED - Check errors above
    echo ============================================
    exit /b 1
)

echo ============================================
echo DAILY UPDATE COMPLETE
echo ============================================
echo.
exit /b 0
