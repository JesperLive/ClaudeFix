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
