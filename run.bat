@echo off
setlocal

title MarkOneNoteDown
cd /d "%~dp0"

echo =========================================
echo   MarkOneNoteDown - OneNote to Markdown
echo =========================================
echo.

:: --- 1. PowerShell ---
echo [1/3] Checking PowerShell...
where powershell >nul 2>&1
if errorlevel 1 (
    echo   [FAIL] PowerShell not found.
    goto :end
)
for /f "tokens=*" %%v in ('powershell -NoProfile -Command "$PSVersionTable.PSVersion.ToString()"') do set PSVER=%%v
echo   [OK]   PowerShell %PSVER%

:: --- 2. Pandoc ---
echo [2/3] Checking Pandoc...
where pandoc >nul 2>&1
if errorlevel 1 goto :pandoc_miss
for /f "tokens=*" %%v in ('pandoc --version 2^>nul ^| findstr /r /c:"^pandoc"') do echo   [OK]   %%v
goto :pandoc_done
:pandoc_miss
echo   [MISS] Pandoc not found in PATH - will prompt during run.
echo          Install: https://pandoc.org/installing.html
:pandoc_done

:: --- 3. Config ---
echo [3/3] Checking config...
if exist "config.ps1" goto :config_found
echo   [INFO] No config.ps1 - interactive mode will run.
echo          Copy config.example.ps1 ^> config.ps1 to skip prompts.
goto :config_done
:config_found
echo   [OK]   config.ps1 found
:config_done

echo.
echo Starting conversion...
echo.

:: --- Run ---
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0MarkOneNoteDown.ps1" %*
set EXITCODE=%ERRORLEVEL%

echo.
if %EXITCODE% neq 0 goto :run_fail
echo [OK]   Finished.
goto :end
:run_fail
echo [FAIL] Errors occurred (code %EXITCODE%).
echo        See README.md "Common Errors" section.

:end
echo.
pause
