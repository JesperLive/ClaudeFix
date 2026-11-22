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

    It reads every .ps1 and .bat file in the repo root and in .ci, so a new
    script, test or launcher is checked without being added to a list. The
    three main scripts and the ClaudeEnv region are named as well, so that one
    gone missing fails. Until 6.3.0 only those four were read: a test or a
    batch file could carry a non-ASCII byte that no gate saw, and cmd.exe reads
    a .bat in the console's code page too.

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

$required = @(
    'Fix-ClaudeDesktop.ps1',
    'Watch-ClaudeHealth.ps1',
    'Prevent-ClaudeIssues.ps1',
    '.ci\ClaudeEnv.region.ps1'
)

# The named four first, in that order, then every other script found, by name.
$targets = [System.Collections.Generic.List[string]]::new()
foreach ($name in $required) { $targets.Add($name) }
foreach ($dir in @('', '.ci')) {
    $full = if ($dir) { Join-Path $Root $dir } else { $Root }
    if (-not (Test-Path -LiteralPath $full -PathType Container)) { continue }
    $files = @(Get-ChildItem -LiteralPath $full -File |
        Where-Object { $_.Extension -eq '.ps1' -or $_.Extension -eq '.bat' } |
        Sort-Object Name)
    foreach ($file in $files) {
        $name = if ($dir) { Join-Path $dir $file.Name } else { $file.Name }
        if (-not ($targets -contains $name)) { $targets.Add($name) }
    }
}

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
