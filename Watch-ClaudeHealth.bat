@echo off
:: Watch-ClaudeHealth -- launcher
:: Usually started automatically by the Prevent-ClaudeIssues scheduled task.
:: Run this manually to monitor in the foreground and see real-time
:: health check output. Close this window to stop monitoring.

title Claude Desktop - Health Monitor

set "SCRIPT=%~dp0Watch-ClaudeHealth.ps1"

if not exist "%SCRIPT%" (
    echo.
    echo   [!] Watch-ClaudeHealth.ps1 not found.
    echo       It must be in the same folder as this .bat file.
    echo       Expected: %SCRIPT%
    echo.
    pause
    exit /b 1
)

echo.
echo   Starting Claude Health Monitor...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*

:: Exit code 1 means the monitor refused to start and wrote the reason to the
:: watch log, for example because Fix-ClaudeDesktop.ps1 was not found next to
:: it. Pause so a person who double-clicked this can read it.
if %ERRORLEVEL% EQU 0 exit /b 0

echo.
if %ERRORLEVEL% EQU 1 (
    echo   [!] The health monitor stopped on startup. The reason was written to:
    echo       %APPDATA%\Claude\watch-logs\
) else (
    echo   [!] PowerShell could not run the script (exit code %ERRORLEVEL%)
)
echo.
pause
exit /b %ERRORLEVEL%
