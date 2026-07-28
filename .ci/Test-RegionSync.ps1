<#
.SYNOPSIS
    Fails if the embedded ClaudeEnv discovery region has drifted between the
    three toolkit scripts.
.DESCRIPTION
    Each script carries its own copy of the region so that every script stays
    runnable standalone, with no shared file to go missing. The cost of that
    choice is three copies, and three hand-maintained copies is precisely how
    the divergent-duplicate bugs in this codebase happened.

    This gate removes the discipline requirement. Edit
    .ci/ClaudeEnv.region.ps1, resync the three scripts, and CI proves it.

    Compares the span from "#region ClaudeEnv" through "#endregion ClaudeEnv"
    inclusive, against the same span in the canonical file. Exits 0 on match.
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

function Get-RegionSpan {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    $lines = @(Get-Content $Path)
    $start = -1
    $end   = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $t = $lines[$i].Trim()
        if ($start -lt 0 -and $t -eq '#region ClaudeEnv') { $start = $i; continue }
        if ($start -ge 0 -and $t -eq '#endregion ClaudeEnv') { $end = $i; break }
    }
    if ($start -lt 0 -or $end -lt 0) { return $null }
    return ($lines[$start..$end] -join "`n")
}

function Get-SpanHash {
    param([Parameter(Mandatory)][string]$Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        return ([BitConverter]::ToString($sha.ComputeHash($bytes)) -replace '-', '').Substring(0, 16)
    } finally { $sha.Dispose() }
}

$canonicalPath = Join-Path $Root '.ci\ClaudeEnv.region.ps1'
$canonical = Get-RegionSpan -Path $canonicalPath
if (-not $canonical) {
    Write-Host "FAIL: canonical region not found at .ci/ClaudeEnv.region.ps1"
    exit 1
}
$canonicalHash = Get-SpanHash -Text $canonical
Write-Host ("canonical : {0}  ({1} chars)" -f $canonicalHash, $canonical.Length)

$failed = 0
foreach ($name in @('Fix-ClaudeDesktop.ps1', 'Watch-ClaudeHealth.ps1', 'Prevent-ClaudeIssues.ps1')) {
    $span = Get-RegionSpan -Path (Join-Path $Root $name)
    if (-not $span) {
        Write-Host ("FAIL      : {0} has no ClaudeEnv region" -f $name)
        $failed++
        continue
    }
    $h = Get-SpanHash -Text $span
    if ($span -ceq $canonical) {
        Write-Host ("OK        : {0}  {1}" -f $h, $name)
    } else {
        Write-Host ("DRIFTED   : {0}  {1}  ({2} chars)" -f $h, $name, $span.Length)
        $failed++
    }
}

Write-Host ""
if ($failed -eq 0) {
    Write-Host "Region sync gate PASSED"
    exit 0
}
Write-Host ("Region sync gate FAILED ({0} file(s) drifted)" -f $failed)
exit 1
