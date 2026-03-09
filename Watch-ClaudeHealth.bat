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

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo   [!] PowerShell exited with error code %ERRORLEVEL%
    echo.
    pause
)
