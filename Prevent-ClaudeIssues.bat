@echo off
:: Prevent-ClaudeIssues -- launcher
:: Run once to configure Windows for stable Claude Desktop/Cowork.

set "SCRIPT=%~dp0Prevent-ClaudeIssues.ps1"

:: The script is looked for by its name, through its attributes. Until
:: 6.3.1 "if exist" was asked, which lists the folder: it answered "no" for a
:: script that is there in a folder that may be read but not listed, and the
:: block in parentheses that echoed the path ended early at a path holding a
:: parenthesis, "... was unexpected at this time." (measured 2026-11-23).
:: When the attributes cannot be read and the folder cannot be listed either,
:: whether the script is there cannot be told, and that is said.
set "SCRIPT_ATTR="
for %%I in ("%SCRIPT%") do set "SCRIPT_ATTR=%%~aI"
if defined SCRIPT_ATTR goto :script_there
dir /b "%~dp0" >nul 2>&1
if errorlevel 1 goto :script_unread
echo.
echo   [!] Prevent-ClaudeIssues.ps1 not found.
echo       It must be in the same folder as this .bat file.
echo.
pause
exit /b 1

:script_unread
echo.
echo   [!] Prevent-ClaudeIssues.ps1 could not be looked for: this folder could not be read.
echo       Expected: %SCRIPT%
echo.
pause
exit /b 1

:script_there

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
