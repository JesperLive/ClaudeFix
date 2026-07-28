# Test-Elevation.ps1
#
# Both self-elevating scripts relaunch themselves through
# Start-Process -Verb RunAs. That relaunch is a new process with a new
# parameter set, so anything the user typed has to be forwarded by hand.
#
# Two things were being dropped, in one script each:
#
#   -WhatIf          Prevent-ClaudeIssues did not forward it. Running the
#                    script unelevated with -WhatIf relaunched it elevated
#                    without -WhatIf and applied every change for real.
#
#   -Wait -PassThru  Without these the parent returns the instant UAC is
#                    accepted, always with code 0, so the .bat launcher and
#                    any caller saw success no matter how the elevated run
#                    turned out.
#
# Both are invisible in normal use: the script does something, and it looks
# like it worked. This gate reads the source rather than running anything,
# because the failure only shows up on an elevation the CI runner cannot do.

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

$targets = @(
    'Fix-ClaudeDesktop.ps1',
    'Watch-ClaudeHealth.ps1',
    'Prevent-ClaudeIssues.ps1'
)

foreach ($name in $targets) {

    $path = Join-Path $root $name
    $src  = [System.IO.File]::ReadAllText($path)

    # Find the relaunch, if there is one. Watch-ClaudeHealth does not elevate
    # itself, so it legitimately has none and is skipped.
    $relaunch = [regex]::Matches($src, 'Start-Process\s+PowerShell[^\r\n]*-Verb\s+RunAs[^\r\n]*')
    if ($relaunch.Count -eq 0) {
        Write-Host ""
        Write-Host "--- $name (does not self-elevate, skipped)" -ForegroundColor DarkGray
        continue
    }

    Write-Host ""
    Write-Host "--- $name" -ForegroundColor Cyan

    Assert "exactly one RunAs relaunch" ($relaunch.Count -eq 1)
    $line = $relaunch[0].Value

    Assert "relaunch waits for the child"   ($line -match '-Wait')
    Assert "relaunch captures the child"    ($line -match '-PassThru')

    # The parameters have to be assembled somewhere above the relaunch. Look
    # for the -WhatIf forward specifically, since that is the one that silently
    # turns a dry run into a real one.
    Assert "forwards -WhatIf across elevation" `
        ($src -match '\$WhatIfPreference[^\r\n]*-WhatIf')

    # And the captured exit code has to actually be used.
    Assert "propagates the child exit code" `
        ($src -match 'exit\s+\$\w+\.ExitCode')

    # A hardcoded success right after the relaunch defeats all of the above.
    # Allowed: "exit 0" as the fallback when Start-Process returned nothing.
    # Not allowed: "exit 0" as the only thing following the relaunch.
    $tail = $src.Substring($relaunch[0].Index)
    $firstExit = [regex]::Match($tail, 'exit\s+\S+')
    Assert "first exit after relaunch is not a hardcoded 0" `
        ($firstExit.Success -and $firstExit.Value -notmatch '^exit\s+0$')
}

# ------------------------------------------------------------------
# Every self-elevating script must also declare SupportsShouldProcess,
# otherwise $WhatIfPreference is never set and the forward above is dead code.
# ------------------------------------------------------------------
Write-Host ""
Write-Host "--- ShouldProcess declared" -ForegroundColor Cyan
foreach ($name in $targets) {
    $src = [System.IO.File]::ReadAllText((Join-Path $root $name))
    # The same pattern as the loop above, not a loose "-Verb RunAs". A bare
    # search matches the comment in Watch-ClaudeHealth that explains how
    # Fix-ClaudeDesktop elevates, which made this loop assert against a script
    # that does not elevate at all.
    if ($src -notmatch 'Start-Process\s+PowerShell[^\r\n]*-Verb\s+RunAs') { continue }
    Assert "$name declares SupportsShouldProcess" `
        ($src -match '\[CmdletBinding\([^\)]*SupportsShouldProcess')
}

Write-Host ""
Write-Host "Elevation forwarding: $script:pass passed, $script:fail failed"
if ($script:fail -gt 0) {
    Write-Host "Elevation gate FAILED" -ForegroundColor Red
    exit 1
}
Write-Host "Elevation gate PASSED" -ForegroundColor Green
exit 0
