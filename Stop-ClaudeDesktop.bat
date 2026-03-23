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
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo   [!] PowerShell exited with error code %ERRORLEVEL%
    echo.
    pause
)
