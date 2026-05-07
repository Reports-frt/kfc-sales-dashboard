@echo off
REM =====================================================================
REM KFC Food Cost Hub — Manual Update
REM =====================================================================
REM Run this whenever you place fresh FoodCost.xlsx in _work\
REM
REM Steps it performs:
REM   1. Reads _work\FoodCost.xlsx + CategoriesFC.xlsx
REM   2. Builds kfc-dashboard\food\food_data.json
REM   3. Pushes to GitHub
REM =====================================================================

cd /d "%~dp0"
echo.
echo ============================================
echo KFC Food Cost Hub - Manual Update
echo ============================================
echo.

"C:\Program Files\Python312\python.exe" build_pipeline.py

echo.
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ============================================
    echo BUILD FAILED - Check errors above
    echo ============================================
    pause
    exit /b 1
)

echo.
echo Press any key to close...
pause >nul
