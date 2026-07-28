<#
.SYNOPSIS
    Parser-only syntax gate for the toolkit scripts.
.DESCRIPTION
    Parses each script with the PowerShell language parser and reports every
    error with its position. Parsing does not execute anything, so this is safe
    to run anywhere, including CI.

    Exits 0 when every file parses clean, 1 otherwise.
#>
[CmdletBinding()]
param(
    [string]$Root
)

# $PSScriptRoot is not reliably populated inside a param() default expression,
# so resolve it in the body instead.
if (-not $Root) {
    $here = $PSScriptRoot
    if (-not $here) { $here = Split-Path $MyInvocation.MyCommand.Path -Parent }
    $Root = Split-Path $here -Parent
}

$targets = @(
    'Fix-ClaudeDesktop.ps1',
    'Watch-ClaudeHealth.ps1',
    'Prevent-ClaudeIssues.ps1'
)

$failed = 0
foreach ($name in $targets) {
    $path = Join-Path $Root $name
    if (-not (Test-Path $path)) {
        Write-Host ("MISSING : {0}" -f $name)
        $failed++
        continue
    }

    $tokens = $null
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile(
        $path, [ref]$tokens, [ref]$errors)

    $errCount = @($errors).Count
    if ($errCount -eq 0) {
        $lines = @(Get-Content $path).Count
        Write-Host ("OK      : {0} ({1} lines, {2} tokens)" -f $name, $lines, @($tokens).Count)
    } else {
        Write-Host ("FAIL    : {0} ({1} parse error(s))" -f $name, $errCount)
        foreach ($e in @($errors)) {
            Write-Host ("          line {0} col {1}: {2}" -f `
                $e.Extent.StartLineNumber, $e.Extent.StartColumnNumber, $e.Message)
        }
        $failed++
    }
}

Write-Host ""
if ($failed -eq 0) {
    Write-Host "Syntax gate PASSED"
    exit 0
}
Write-Host ("Syntax gate FAILED ({0} file(s))" -f $failed)
exit 1
