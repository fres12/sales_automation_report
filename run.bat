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
    echo Error occurred! Press any key to exit...
    echo ========================================
    pause
    exit /b %errorlevel%
) else (
    echo.
    echo ========================================
    echo Application completed successfully!
    echo ========================================
    pause
)

endlocal
