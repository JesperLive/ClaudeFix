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

:: Exit code 1 means the script RAN and reported a problem. From v6.0.0 the
:: script sets that deliberately when retries are exhausted or an unhandled
:: error occurs. It has already explained why and waited for a keypress, so
:: pausing again here would double-prompt and imply the launcher failed.
::
:: Anything other than 0 or 1 means PowerShell itself could not run the script,
:: which is worth stopping for.
if %ERRORLEVEL% EQU 0 exit /b 0
if %ERRORLEVEL% EQU 1 exit /b 1

echo.
echo   [!] PowerShell could not run the script (exit code %ERRORLEVEL%)
echo       If you see a red error above, please screenshot it.
echo.
pause
exit /b %ERRORLEVEL%
