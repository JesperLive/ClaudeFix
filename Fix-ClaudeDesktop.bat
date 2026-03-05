@echo off
:: Fix-ClaudeDesktop -- launcher
:: Place this file alongside Fix-ClaudeDesktop.ps1 and double-click to run.

set "SCRIPT=%~dp0Fix-ClaudeDesktop.ps1"

if not exist "%SCRIPT%" (
    echo.
    echo   [!] Fix-ClaudeDesktop.ps1 not found.
    echo       It must be in the same folder as this .bat file.
    echo       Expected: %SCRIPT%
    echo.
    pause
    exit /b 1
)

echo.
echo   Starting Fix-ClaudeDesktop...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*

:: PS1 already shows "Press any key to close..." so we just exit.
:: Only pause on PS1 launch failure (errorlevel from powershell itself).
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo   [!] PowerShell exited with error code %ERRORLEVEL%
    echo       If you see a red error above, please screenshot it.
    echo.
    pause
)
