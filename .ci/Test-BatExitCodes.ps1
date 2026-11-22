# Test-BatExitCodes.ps1
#
# The four .bat launchers are the only entry point most users ever touch, and
# until v6.0.0 nothing tested them. They all carried:
#
#     if %ERRORLEVEL% NEQ 0 ( echo please screenshot it & pause )
#
# which was correct while the scripts only ever exited 0. Finding 19 made
# Fix-ClaudeDesktop exit 1 deliberately when retries are exhausted. That turned
# a legitimate, already-explained outcome into a second scary prompt telling the
# user to screenshot an error the launcher did not produce.
#
# This gate runs each real .bat with the PowerShell invocation swapped for a
# stub that exits with a chosen code, and asserts what the user sees.
#
# The stub copy is written into the repo root, not %TEMP%, because every
# launcher starts with an "is the .ps1 next to me" guard keyed on %~dp0. Running
# from %TEMP% would take that branch and test nothing.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$script:pass = 0
$script:fail = 0

function Assert {
    param([string]$Label, [bool]$Condition)
    if ($Condition) {
        Write-Host "  PASS  $Label" -ForegroundColor Green
        $script:pass++
    } else {
        Write-Host "  FAIL  $Label" -ForegroundColor Red
        $script:fail++
    }
}

function Invoke-Launcher {
    param([string]$BatName, [int]$StubExit)

    $batPath = Join-Path $root $BatName
    if (-not (Test-Path $batPath)) { throw "launcher not found: $batPath" }

    $lines = [System.IO.File]::ReadAllLines($batPath)
    $hits = 0
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*powershell(\.exe)?\s') {
            $lines[$i] = "cmd /c exit $StubExit"
            $hits++
        }
    }
    if ($hits -ne 1) {
        throw "expected exactly one powershell invocation in ${BatName}, found $hits"
    }

    $stub = Join-Path $root ("_battest_" + [Guid]::NewGuid().ToString('N') + ".bat")
    $out  = "$stub.out"
    # Every launcher ends in "pause". An empty file as stdin gives it EOF at
    # once, which is what a redirected console does anyway. Start-Process will
    # not take NUL here, it resolves the name against the working directory.
    $in   = "$stub.in"
    [System.IO.File]::WriteAllLines($stub, $lines, (New-Object System.Text.ASCIIEncoding))
    [System.IO.File]::WriteAllText($in, '')

    try {
        $p = Start-Process -FilePath $env:ComSpec `
                           -ArgumentList @('/c', "`"$stub`"") `
                           -NoNewWindow -Wait -PassThru `
                           -RedirectStandardOutput $out `
                           -RedirectStandardInput $in
        $text = ''
        if (Test-Path $out) { $text = [System.IO.File]::ReadAllText($out) }
        return [pscustomobject]@{ ExitCode = $p.ExitCode; Output = $text }
    } finally {
        Remove-Item $stub, $out, $in -Force -ErrorAction SilentlyContinue
    }
}

# ------------------------------------------------------------------
# The three launchers that treat 0 and 1 as "the script spoke for itself"
# ------------------------------------------------------------------
foreach ($bat in @('Fix-ClaudeDesktop.bat', 'Prevent-ClaudeIssues.bat', 'Stop-ClaudeDesktop.bat')) {

    Write-Host ""
    Write-Host "--- $bat" -ForegroundColor Cyan

    $r0 = Invoke-Launcher -BatName $bat -StubExit 0
    Assert "exit 0 propagates"                    ($r0.ExitCode -eq 0)
    Assert "exit 0 prints no launcher complaint"  ($r0.Output -notmatch '\[!\]')

    # This is the regression the gate exists for.
    $r1 = Invoke-Launcher -BatName $bat -StubExit 1
    Assert "exit 1 propagates"                    ($r1.ExitCode -eq 1)
    Assert "exit 1 prints no launcher complaint"  ($r1.Output -notmatch '\[!\]')
    Assert "exit 1 does not ask for a screenshot" ($r1.Output -notmatch 'screenshot')

    # A code the script never sets means PowerShell itself failed to start it,
    # and that is still worth stopping for.
    $r9 = Invoke-Launcher -BatName $bat -StubExit 9
    Assert "exit 9 propagates"                    ($r9.ExitCode -eq 9)
    Assert "exit 9 does complain"                 ($r9.Output -match '\[!\]')
    Assert "exit 9 names the code"                ($r9.Output -match '9')
}

# ------------------------------------------------------------------
# 3 and 4 are outcomes of a Fix run, through Fix-ClaudeDesktop.bat and
# Stop-ClaudeDesktop.bat alike: 3, held back before the app or the service
# was stopped; 4, stopped part-way because a Claude Code session was found or
# could not be ruled out. Each has its own lines, no launcher complaint and
# no pause of the launcher's own, as for 0 and 1. Until 6.3.0 Fix set
# neither, and the launchers read both as PowerShell failing to run the
# script. Until 6.3.0 too, 3 said nothing was changed and 4 that a session
# had started, and neither is always so.
# ------------------------------------------------------------------
foreach ($bat in @('Fix-ClaudeDesktop.bat', 'Stop-ClaudeDesktop.bat')) {

    Write-Host ""
    Write-Host "--- $bat, exits 3 and 4" -ForegroundColor Cyan

    # The pause the launcher prints its prompt for, seen where it must be,
    # so its absence below means something.
    $r9 = Invoke-Launcher -BatName $bat -StubExit 9
    Assert "the harness sees the launcher's pause (exit 9)" ($r9.Output -match 'Press any key')

    foreach ($c in @(
        @{ Code = 3; Rx = @(
            '(?m)^\s*Held back before the app or the service was stopped: another repair is running,\s*$',
            '(?m)^\s*or the PC or Claude is in use, or a Claude Code session could not be ruled out\.\s*$',
            '(?m)^\s*The lines above say which\.\s*$') },
        @{ Code = 4; Rx = @(
            '(?m)^\s*The repair ran part of the way, then stopped before going further: a Claude Code\s*$',
            '(?m)^\s*session was found, or could not be ruled out\. The lines above say which\.\s*$') }
    )) {
        $r = Invoke-Launcher -BatName $bat -StubExit $c.Code
        $unsaid = @($c.Rx | Where-Object { $r.Output -notmatch $_ })
        Assert "exit $($c.Code) propagates"                    ($r.ExitCode -eq $c.Code)
        Assert "exit $($c.Code) says what happened"            ($unsaid.Count -eq 0)
        Assert "exit $($c.Code) does not say nothing was changed, or that a session started" `
            ($r.Output -notmatch 'nothing was changed|a Claude Code session started')
        Assert "exit $($c.Code) prints no launcher complaint"  ($r.Output -notmatch '\[!\]')
        Assert "exit $($c.Code) does not ask for a screenshot" ($r.Output -notmatch 'screenshot')
        Assert "exit $($c.Code) does not pause, as 0 and 1 do not" ($r.Output -notmatch 'Press any key')
    }
}

# Prevent never sets 3 or 4, so there they are still PowerShell failing.
$p3 = Invoke-Launcher -BatName 'Prevent-ClaudeIssues.bat' -StubExit 3
Assert "Prevent-ClaudeIssues.bat: exit 3 is still a complaint" ($p3.ExitCode -eq 3 -and $p3.Output -match '\[!\]')

# ------------------------------------------------------------------
# Watch-ClaudeHealth is the exception. Exit 1 there means the monitor refused
# to start, so the launcher SHOULD pause and point at the log folder.
# ------------------------------------------------------------------
Write-Host ""
Write-Host "--- Watch-ClaudeHealth.bat" -ForegroundColor Cyan

$w0 = Invoke-Launcher -BatName 'Watch-ClaudeHealth.bat' -StubExit 0
Assert "exit 0 propagates"                   ($w0.ExitCode -eq 0)
Assert "exit 0 prints no launcher complaint" ($w0.Output -notmatch '\[!\]')

$w1 = Invoke-Launcher -BatName 'Watch-ClaudeHealth.bat' -StubExit 1
Assert "exit 1 propagates"                   ($w1.ExitCode -eq 1)
Assert "exit 1 explains the monitor stopped" ($w1.Output -match 'monitor stopped')
Assert "exit 1 points at the log folder"     ($w1.Output -match 'watch-logs')
Assert "exit 1 resolves APPDATA, not literal" ($w1.Output -notmatch '%APPDATA%')

$w9 = Invoke-Launcher -BatName 'Watch-ClaudeHealth.bat' -StubExit 9
Assert "exit 9 propagates"                   ($w9.ExitCode -eq 9)
Assert "exit 9 blames PowerShell, not the monitor" ($w9.Output -match 'could not run')
Assert "exit 9 does not point at the log folder"   ($w9.Output -notmatch 'watch-logs')

# ------------------------------------------------------------------
# Positive control. If the launchers ever regress to "NEQ 0", the exit-1
# assertions above must fail. Prove the harness can actually see that by
# checking the old pattern is gone from source.
# ------------------------------------------------------------------
Write-Host ""
Write-Host "--- old pattern is gone" -ForegroundColor Cyan
foreach ($bat in @('Fix-ClaudeDesktop.bat', 'Prevent-ClaudeIssues.bat',
                   'Stop-ClaudeDesktop.bat', 'Watch-ClaudeHealth.bat')) {
    $src = [System.IO.File]::ReadAllText((Join-Path $root $bat))
    Assert "$bat has no blanket NEQ 0 handler" ($src -notmatch '%ERRORLEVEL%\s+NEQ\s+0')
}

Write-Host ""
Write-Host "Launcher exit codes: $script:pass passed, $script:fail failed"
if ($script:fail -gt 0) {
    Write-Host "Launcher exit code gate FAILED" -ForegroundColor Red
    exit 1
}
Write-Host "Launcher exit code gate PASSED" -ForegroundColor Green
exit 0
