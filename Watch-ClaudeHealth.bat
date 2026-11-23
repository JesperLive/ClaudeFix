@echo off
:: Watch-ClaudeHealth -- launcher
:: Usually started automatically by the Prevent-ClaudeIssues scheduled task.
:: Run this manually to monitor in the foreground and see real-time
:: health check output. Close this window to stop monitoring.

title Claude Desktop - Health Monitor

set "SCRIPT=%~dp0Watch-ClaudeHealth.ps1"

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
echo   [!] Watch-ClaudeHealth.ps1 not found.
echo       It must be in the same folder as this .bat file.
echo       Expected: %SCRIPT%
echo.
pause
exit /b 1

:script_unread
echo.
echo   [!] Watch-ClaudeHealth.ps1 could not be looked for: this folder could not be read.
echo       Expected: %SCRIPT%
echo.
pause
exit /b 1

:script_there

echo.
echo   Starting Claude Health Monitor...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*

:: Exit code 1 means the monitor refused to start and wrote the reason to the
:: watch log, for example because Fix-ClaudeDesktop.ps1 was not found next to
:: it. Pause so a person who double-clicked this can read it.
if %ERRORLEVEL% EQU 0 exit /b 0

echo.
:: Not in a block in parentheses: a profile path holding a parenthesis
:: would end the block early (measured 2026-11-23, as for the script's).
if %ERRORLEVEL% NEQ 1 goto :ps_failed
echo   [!] The health monitor stopped on startup. The reason was written to:
echo       %APPDATA%\Claude\watch-logs\
goto :said
:ps_failed
echo   [!] PowerShell could not run the script (exit code %ERRORLEVEL%)
:said
echo.
pause
exit /b %ERRORLEVEL%
