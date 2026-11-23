@echo off
:: Fix-ClaudeDesktop -- launcher
:: Place this file alongside Fix-ClaudeDesktop.ps1 and double-click to run.

set "SCRIPT=%~dp0Fix-ClaudeDesktop.ps1"

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
echo   [!] Fix-ClaudeDesktop.ps1 not found.
echo       It must be in the same folder as this .bat file.
echo       Expected: %SCRIPT%
echo.
pause
exit /b 1

:script_unread
echo.
echo   [!] Fix-ClaudeDesktop.ps1 could not be looked for: this folder could not be read.
echo       Expected: %SCRIPT%
echo.
pause
exit /b 1

:script_there

echo.
echo   Starting Fix-ClaudeDesktop...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*

:: Exit code 1 means the script RAN and reported a problem. From v6.0.0 the
:: script sets that deliberately when retries are exhausted or an unhandled
:: error occurs. It has already explained why and waited for a keypress, so
:: pausing again here would double-prompt and imply the launcher failed.
::
:: 3 and 4 are outcomes too, each with its own line. 3: the run was held back
:: before it stopped the app or the service, because another repair is
:: running or, in a -Quiet run, the PC or Claude was in use or a Claude Code
:: session could not be ruled out. 4: a -Quiet repair ran part of the way,
:: then stopped because a Claude Code session was found or could not be ruled
:: out. The script has said why, and in a manual run waited for a keypress,
:: so there is no pause for them either.
:: Until 6.3.0 nothing set 3 or 4, and this reported them as PowerShell
:: failing to run the script.
::
:: Anything other than 0, 1, 3 or 4 means PowerShell itself could not run the
:: script, which is worth stopping for.
if %ERRORLEVEL% EQU 0 exit /b 0
if %ERRORLEVEL% EQU 1 exit /b 1
if %ERRORLEVEL% EQU 3 (
    echo.
    echo   Held back before the app or the service was stopped: another repair is running,
    echo   or the PC or Claude is in use, or a Claude Code session could not be ruled out.
    echo   The lines above say which.
    echo.
    exit /b 3
)
if %ERRORLEVEL% EQU 4 (
    echo.
    echo   The repair ran part of the way, then stopped before going further: a Claude Code
    echo   session was found, or could not be ruled out. The lines above say which.
    echo.
    exit /b 4
)

echo.
echo   [!] PowerShell could not run the script (exit code %ERRORLEVEL%)
echo       If you see a red error above, please screenshot it.
echo.
pause
exit /b %ERRORLEVEL%
