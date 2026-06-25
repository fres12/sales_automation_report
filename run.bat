@echo off
setlocal enabledelayedexpansion

cd /d "%~dp0"

echo.
echo ========================================
echo Starting Application...
echo ========================================
echo.

python main.py

if %errorlevel% neq 0 (
    echo.
    echo ========================================
    echo Error occurred! Closing in 5 seconds...
    echo ========================================
    timeout /t 5 /nobreak >nul
    exit /b %errorlevel%
) else (
    echo.
    echo ========================================
    echo Application completed successfully! Closing in 5 seconds...
    echo ========================================
    timeout /t 5 /nobreak >nul
)

endlocal