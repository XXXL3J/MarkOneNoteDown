@echo off
setlocal

:: ============================================================
:: MarkOneNoteDown Launcher
:: Double-click this file to start the conversion.
:: ============================================================

cd /d "%~dp0"

echo ========================================
echo   MarkOneNoteDown — OneNote to Markdown
echo ========================================
echo.

:: Check for PowerShell
where powershell >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] PowerShell not found in PATH.
    pause
    exit /b 1
)

:: Check for config file
if exist "config.ps1" (
    echo [INFO]  config.ps1 found — loading your saved settings.
) else (
    echo [INFO]  No config.ps1 found — running interactive setup.
    echo         Copy config.example.ps1 to config.ps1 to skip prompts next time.
)
echo.

:: Run (ExecutionPolicy bypass for this session only)
powershell -NoProfile -ExecutionPolicy Bypass -File ".\MarkOneNoteDown.ps1" %*

if %errorlevel% neq 0 (
    echo.
    echo [FAIL] Conversion ended with errors (code %errorlevel%).
    echo         See README.md #errors--solutions for help.
) else (
    echo.
    echo [OK]   Conversion complete!
)
echo.
pause
