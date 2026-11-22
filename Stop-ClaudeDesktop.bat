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
:: said so and waited for a keypress. 3 and 4 are outcomes too, each with its
:: own line: 3, the run was held back before it stopped the app or the
:: service; 4, it ran part of the way and stopped because a Claude Code
:: session was found or could not be ruled out. The script has said why, and
:: in a manual run waited for a keypress, so there is no pause for them
:: either. Until 6.3.0 nothing set 3 or 4, and this reported them as
:: PowerShell failing to run the script. Only stop here when PowerShell itself
:: could not run it.
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
echo.
pause
exit /b %ERRORLEVEL%
