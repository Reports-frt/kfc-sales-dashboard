@echo off
REM =====================================================================
REM KFC Food Cost — Daily Update (Pull + Build + Push)
REM =====================================================================
REM Steps it performs:
REM   1. Pulls 3 latest emails from Outlook → saves to _work\
REM   2. Reads xlsx files + builds food_data.json
REM   3. Pushes to GitHub
REM
REM Run manually: double-click this file
REM Or schedule via Task Scheduler to run daily (e.g., 09:30)
REM =====================================================================

cd /d "%~dp0"
echo.
echo ============================================
echo KFC Food Cost - Daily Update
echo ============================================
echo.

REM Step 1: Pull emails from Outlook
echo === STEP 1: Pulling emails from Outlook ===
"C:\Program Files\Python312\python.exe" pull_food_emails.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo WARNING: Email pull had issues. Continuing with existing files...
    echo.
)

REM Step 2 + 3: Build + Push
echo.
echo === STEP 2-3: Build food_data.json + Push to GitHub ===
"C:\Program Files\Python312\python.exe" build_pipeline.py

echo.
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ============================================
    echo BUILD/PUSH FAILED - Check errors above
    echo ============================================
    pause
    exit /b 1
)

echo.
echo ============================================
echo DAILY UPDATE COMPLETE
echo ============================================
echo.
echo Press any key to close...
pause >nul
