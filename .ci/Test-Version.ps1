<#
.SYNOPSIS
    Fails if the three scripts disagree about the toolkit version.
.DESCRIPTION
    Before v6.0.0 the scripts carried independent versions: Fix 5.4.0, Watch
    5.0.0, Prevent 2.0.1. Fix also disagreed with itself, declaring 5.4.0 in
    the constants and 5.3.1 in its own help block, which is what made the drift
    obvious in the first place.

    One version now covers the whole toolkit, and this check keeps it that way
    by comparing every place the number appears:

      - the $ToolkitVersion assignment
      - the "Version :" line in the comment-based help

    Exits 0 when all of them agree, 1 otherwise.
#>
[CmdletBinding()]
param(
    [string]$Root
)

if (-not $Root) {
    $here = $PSScriptRoot
    if (-not $here) { $here = Split-Path $MyInvocation.MyCommand.Path -Parent }
    $Root = Split-Path $here -Parent
}

$found = @()
$failed = 0

foreach ($name in @('Fix-ClaudeDesktop.ps1', 'Watch-ClaudeHealth.ps1', 'Prevent-ClaudeIssues.ps1')) {
    $path = Join-Path $Root $name
    if (-not (Test-Path $path)) {
        Write-Host ("MISSING : {0}" -f $name)
        $failed++
        continue
    }

    $text = Get-Content $path -Raw

    $declared = $null
    if ($text -match '\$ToolkitVersion\s*=\s*"([^"]+)"') { $declared = $Matches[1] }

    $helpVersion = $null
    if ($text -match '(?m)^\s*Version\s*:\s*(\S+)\s*$') { $helpVersion = $Matches[1] }

    if (-not $declared) {
        Write-Host ("FAIL    : {0} has no `$ToolkitVersion assignment" -f $name)
        $failed++
        continue
    }
    if (-not $helpVersion) {
        Write-Host ("FAIL    : {0} has no 'Version :' line in its help block" -f $name)
        $failed++
        continue
    }
    if ($declared -ne $helpVersion) {
        Write-Host ("FAIL    : {0} disagrees with itself (constant {1}, help {2})" -f $name, $declared, $helpVersion)
        $failed++
        continue
    }

    Write-Host ("OK      : {0,-26} {1}" -f $name, $declared)
    $found += $declared
}

$distinct = @($found | Select-Object -Unique)
if ($distinct.Count -gt 1) {
    Write-Host ("FAIL    : scripts disagree with each other: {0}" -f ($distinct -join ', '))
    $failed++
}

Write-Host ""
if ($failed -eq 0) {
    Write-Host ("Version gate PASSED ({0})" -f $distinct[0])
    exit 0
}
Write-Host "Version gate FAILED"
exit 1
