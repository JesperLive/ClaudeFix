@echo off
:: Stop-ClaudeDesktop -- clean shutdown launcher
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
echo   Stopping Claude Desktop cleanly...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Close %*

:: Exit code 1 means the script ran and reported a problem, and it has already
:: said so and waited for a keypress. Only stop here when PowerShell itself
:: could not run it.
if %ERRORLEVEL% EQU 0 exit /b 0
if %ERRORLEVEL% EQU 1 exit /b 1

echo.
echo   [!] PowerShell could not run the script (exit code %ERRORLEVEL%)
echo.
pause
exit /b %ERRORLEVEL%
