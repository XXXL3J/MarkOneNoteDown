@echo off
setlocal EnableDelayedExpansion

:: ============================================================
:: MarkOneNoteDown Launcher
:: Double-click to run with automatic dependency check.
:: ============================================================

cd /d "%~dp0"

echo =========================================
echo   MarkOneNoteDown - OneNote to Markdown
echo =========================================
echo.

:: --- 1. PowerShell ---
echo [1/3] Checking PowerShell...
where powershell >nul 2>&1
if %errorlevel% neq 0 (
    echo   [FAIL] PowerShell not found.
    pause & exit /b 1
)
for /f "tokens=*" %%v in ('powershell -NoProfile -Command "$PSVersionTable.PSVersion.ToString()"') do set PSVER=%%v
echo   [OK]   PowerShell !PSVER!

:: --- 2. Pandoc ---
echo [2/3] Checking Pandoc...
where pandoc >nul 2>&1
if %errorlevel% neq 0 (
    echo   [MISS] Pandoc not found - will prompt during run.
    echo          Install: https://pandoc.org/installing.html
) else (
    for /f "tokens=*" %%v in ('pandoc --version 2^>nul ^| findstr /r "^pandoc"') do set PANDOCVER=%%v
    echo   [OK]   !PANDOCVER!
)

:: --- 3. Config ---
echo [3/3] Checking config...
if exist "config.ps1" (
    echo   [OK]   config.ps1 found
) else (
    echo   [INFO] No config.ps1 - interactive mode will run.
    echo          Copy config.example.ps1 ^> config.ps1 to skip prompts.
)

echo.
echo Starting conversion...
echo.

:: --- Run ---
powershell -NoProfile -ExecutionPolicy Bypass -File ".\MarkOneNoteDown.ps1" %*

if %errorlevel% neq 0 (
    echo.
    echo [FAIL] Errors occurred (code %errorlevel%).
    echo        See README "Common Errors" section.
) else (
    echo.
    echo [OK]   Finished.
)
echo.
pause
