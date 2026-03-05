@echo off
:: Prevent-ClaudeIssues -- launcher
:: Run once to configure Windows for stable Claude Desktop/Cowork.

set "SCRIPT=%~dp0Prevent-ClaudeIssues.ps1"

if not exist "%SCRIPT%" (
    echo.
    echo   [!] Prevent-ClaudeIssues.ps1 not found.
    echo       It must be in the same folder as this .bat file.
    echo.
    pause
    exit /b 1
)

echo.
echo   Starting Prevent-ClaudeIssues...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*

:: PS1 already shows "Press any key to close..." so we just exit.
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo   [!] PowerShell exited with error code %ERRORLEVEL%
    echo.
    pause
)
