<#
.SYNOPSIS
    Fails if any toolkit script contains a non-ASCII byte.
.DESCRIPTION
    These scripts run under Windows PowerShell 5.1 on machines whose console
    code page is frequently cp1252. A stray non-ASCII character either renders
    as mojibake in the console and the log, or breaks the line outright.

    The check is byte-level rather than character-level, because the failure it
    guards against is an encoding failure.

    A UTF-8 BOM at the start of a file is reported but does not fail: it is
    deliberate, and Windows PowerShell uses it to pick the right encoding.

    Written after a full-width parenthesis reached the source of
    Watch-ClaudeHealth.ps1. It parsed as an ordinary token, so the syntax gate
    was perfectly happy with it.

    Exits 0 when clean, 1 otherwise.
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

$targets = @(
    'Fix-ClaudeDesktop.ps1',
    'Watch-ClaudeHealth.ps1',
    'Prevent-ClaudeIssues.ps1',
    '.ci\ClaudeEnv.region.ps1'
)

$failed = 0
foreach ($name in $targets) {
    $path = Join-Path $Root $name
    if (-not (Test-Path $path)) {
        Write-Host ("MISSING : {0}" -f $name)
        $failed++
        continue
    }

    $bytes = [System.IO.File]::ReadAllBytes($path)
    $start = 0
    $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    if ($hasBom) { $start = 3 }

    # Map byte offsets to line numbers so a hit is actually findable.
    $line = 1
    $bad  = @()
    for ($i = $start; $i -lt $bytes.Length; $i++) {
        $b = $bytes[$i]
        if ($b -eq 0x0A) { $line++; continue }
        if ($b -gt 0x7F) {
            $bad += [pscustomobject]@{ Line = $line; Offset = $i; Byte = $b }
            if ($bad.Count -ge 20) { break }
        }
    }

    $bomNote = if ($hasBom) { " (UTF-8 BOM present, allowed)" } else { "" }
    if ($bad.Count -eq 0) {
        Write-Host ("OK      : {0}{1}" -f $name, $bomNote)
    } else {
        Write-Host ("FAIL    : {0} ({1} non-ASCII byte(s)){2}" -f $name, $bad.Count, $bomNote)
        foreach ($b in $bad) {
            Write-Host ("          line {0,-6} offset {1,-8} 0x{2:X2}" -f $b.Line, $b.Offset, $b.Byte)
        }
        $failed++
    }
}

Write-Host ""
if ($failed -eq 0) { Write-Host "ASCII gate PASSED"; exit 0 }
Write-Host ("ASCII gate FAILED ({0} file(s))" -f $failed)
exit 1
