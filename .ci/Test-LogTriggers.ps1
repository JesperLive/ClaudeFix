<#
.SYNOPSIS
    Feeds real log lines to the health monitor's log trigger and checks which
    ones would start a repair.
.DESCRIPTION
    Test-LogsForErrors is the only check in Watch that starts a repair on a
    single hit, with no second confirming poll. Until 6.2.1 it matched its
    error strings against every new byte of every *.log file in the app's logs
    folder, at any log level.

    App 6.18355 writes "RPC error -1" in an [info] [remote-bash] line whenever a
    remote session's probe finds its worker already running and carries on.
    Measured 2026-11-21: 14 such lines in main.log on 2026-10-31, 11-07, 11-16
    and 11-18, and every one would have queued a repair that restarts the app.

    The real VM failures in the same logs are [error] [VM:start] lines in
    cowork_vm_node.log, for example the Plan9 mount failures of 2026-09-11.

    The cases below are those captured lines, with the remote worker names and
    process ids replaced by placeholders of the same shape. Test-LogsForErrors
    is lifted out of Watch-ClaudeHealth.ps1 with the language parser, together
    with the pattern list and the helpers it calls, and runs against a
    throwaway logs folder, so what is tested is the shipped code.

    Exits 0 when every case passes, 1 otherwise.
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

$watchPath = Join-Path $Root 'Watch-ClaudeHealth.ps1'

$tokens = $null
$errors = $null
$watchAst = [System.Management.Automation.Language.Parser]::ParseFile($watchPath, [ref]$tokens, [ref]$errors)
if (@($errors).Count -gt 0) { Write-Host "FAIL  parse errors in $watchPath"; exit 1 }

function Get-FunctionText {
    param([string]$Name)
    $fn = $watchAst.Find({
        param($n)
        $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name
    }, $true)
    if ($fn) { return $fn.Extent.Text }
    return $null
}

# A top-level assignment such as $LegacyErrorPatterns = @(...), found by the
# variable it assigns, so the test uses the shipped list rather than a copy.
function Get-AssignmentText {
    param([string]$Name)
    $node = $watchAst.Find({
        param($n)
        $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
        $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
        $n.Left.VariablePath.UserPath -eq $Name
    }, $false)
    if ($node) { return $node.Extent.Text }
    return $null
}

$pass = 0
$fail = 0
function Assert-Case {
    param([string]$Name, [bool]$Condition, [string]$Detail = '')
    if ($Condition) { Write-Host ("  PASS  {0}" -f $Name); $script:pass++ }
    else { Write-Host ("  FAIL  {0}  {1}" -f $Name, $Detail); $script:fail++ }
}

$patternsText = Get-AssignmentText -Name 'LegacyErrorPatterns'
if (-not $patternsText) { Write-Host "FAIL  `$LegacyErrorPatterns not found in $watchPath"; exit 1 }
Invoke-Expression $patternsText

# The file-name filter and the line matcher arrived in 6.2.1. They are lifted
# when present so this file also runs, and fails for the right reason, against
# the code from before them.
$filterText = Get-AssignmentText -Name 'LogTriggerFilePattern'
if ($filterText) { Invoke-Expression $filterText }
$matcherText = Get-FunctionText -Name 'Find-LogErrorPattern'
if ($matcherText) { Invoke-Expression $matcherText }

$scanText = Get-FunctionText -Name 'Test-LogsForErrors'
if (-not $scanText) { Write-Host "FAIL  Test-LogsForErrors not found in $watchPath"; exit 1 }
Invoke-Expression $scanText

$tempBase = [System.IO.Path]::GetTempPath()
$ClaudeLogDir = Join-Path $tempBase ('cf-logtrig-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $ClaudeLogDir -Force
$script:LogBaselines = @{}

# Appends to one log and runs one scan. The first scan of a file only records
# its size, as the monitor does, so each file is seeded with a line first.
$utf8 = New-Object System.Text.UTF8Encoding($false)
function Invoke-Scan {
    param([string]$File, [string[]]$Lines)
    $path = Join-Path $ClaudeLogDir $File
    if (-not (Test-Path -LiteralPath $path)) {
        [System.IO.File]::WriteAllText($path, "2026-11-21 18:00:00 [info] seed`r`n", $utf8)
        $null = Test-LogsForErrors
    }
    [System.IO.File]::AppendAllText($path, (($Lines -join "`r`n") + "`r`n"), $utf8)
    return (Test-LogsForErrors)
}

try {
    # -----------------------------------------------------------------
    Write-Host ""
    Write-Host "--- harmless lines must not start a repair"

    # main.log, 2026-11-18 15:26:11. Worker name and process id replaced.
    $nov18 = @(
        '2026-11-18 15:26:11 [info] [remote-bash] resume failed (RPC error -1: process with name "rcw-0000aaaabbbbccccddddeeee" already running (id: oneshot-00000000-0000-4000-8000-000000000001)), continuing probe',
        '2026-11-18 15:26:11 [info] [remote-bash] create failed (RPC error -1: ensure user: user rcw-0000aaaabbbbccccddddeeee already exists unexpectedly: uid=1295 gid=1295), continuing probe'
    )
    $r = Invoke-Scan -File 'main.log' -Lines $nov18
    Assert-Case "11-18 [info] [remote-bash] RPC error -1 pair" ($null -eq $r) "returned '$r'"

    # main1.log, 2026-11-16, four probes in five minutes.
    $nov16 = @()
    foreach ($t in @('20:16:02', '20:17:07', '20:20:43', '20:21:46')) {
        $nov16 += "2026-11-16 $t [info] [remote-bash] resume failed (RPC error -1: process with name `"rcw-0000ffffgggghhhhiiiijjjj`" already running (id: oneshot-00000000-0000-4000-8000-000000000002)), continuing probe"
        $nov16 += "2026-11-16 $t [info] [remote-bash] create failed (RPC error -1: ensure user: user rcw-0000ffffgggghhhhiiiijjjj already exists unexpectedly: uid=1281 gid=1281), continuing probe"
    }
    $r = Invoke-Scan -File 'main1.log' -Lines $nov16
    Assert-Case "11-16 [info] [remote-bash] RPC error -1, eight lines" ($null -eq $r) "returned '$r'"

    # A remote session's probe is never the local VM failing, at any level.
    $r = Invoke-Scan -File 'main.log' -Lines @(
        '2026-11-21 18:01:00 [warn] [remote-bash] resume failed (RPC error -1: process with name "rcw-0000aaaabbbbccccddddeeee" already running), continuing probe')
    Assert-Case "[warn] [remote-bash] RPC error -1" ($null -eq $r) "returned '$r'"
    $r = Invoke-Scan -File 'main.log' -Lines @(
        '2026-11-21 18:02:00 [error] [remote-bash] probe gave up (RPC error -1: Plan9 mount failed: invalid argument)')
    Assert-Case "[error] [remote-bash] is not the local VM either" ($null -eq $r) "returned '$r'"

    # cowork_vm_node1.log, 2026-09-11 19:35:17. The app retried on its own
    # fast path; the same incident also logged [error] [VM:start] lines, which
    # the positive cases below show still start a repair.
    $r = Invoke-Scan -File 'cowork_vm_node.log' -Lines @(
        '2026-09-11 19:35:17 [warn] [startVM] Post-connect setup failed on already-connected fast path: Error: RPC error -1: host drive share "c" not mounted: Plan9 mount failed: invalid argument',
        '    at Socket.emit (node:events:514:28)',
        "  [message]: 'RPC error -1: host drive share `"c`" not mounted: Plan9 mount failed: invalid argument'")
    Assert-Case "[warn] [startVM] with its stack and message lines" ($null -eq $r) "returned '$r'"

    # "bad address" is a bare errno string. It is not a VM signature on its own.
    $r = Invoke-Scan -File 'main.log' -Lines @(
        '2026-11-21 18:03:00 [error] [mcp] connect EFAULT: bad address')
    Assert-Case "[error] line with only 'bad address'" ($null -eq $r) "returned '$r'"

    # A chunk that starts inside an entry has no level to go by.
    $r = Invoke-Scan -File 'cowork_vm_node.log' -Lines @(
        "  [message]: 'RPC error -1: host drive share `"c`" not mounted: Plan9 mount failed: invalid argument'",
        '    at Socket.emit (node:events:514:28)')
    Assert-Case "continuation lines with no entry header" ($null -eq $r) "returned '$r'"

    # The in-app browser logs what web pages write to their consoles.
    $r = Invoke-Scan -File 'in-app-browser.log' -Lines @(
        '2026-11-21 18:04:00 [error] Uncaught Error: HCS operation failed')
    Assert-Case "an [error] line in a log outside the app's VM and main logs" ($null -eq $r) "returned '$r'"

    # -----------------------------------------------------------------
    Write-Host ""
    Write-Host "--- real VM failures must still start a repair"

    # cowork_vm_node1.log, 2026-09-11 18:59:18, header and stack as captured.
    $r = Invoke-Scan -File 'cowork_vm_node.log' -Lines @(
        '2026-09-11 18:59:18 [error] [VM:start] Startup failed: Error: RPC error -1: host drive share "c" not mounted: Plan9 mount failed: invalid argument',
        '    at Socket.<anonymous> (C:\Program Files\WindowsApps\Claude_1.52386.0.0_x64__pzs8sxrjxfjjc\app\resources\app.asar\.vite\build\index.chunk-b0PhZOZp.js:13:1422973)',
        '    at Socket.emit (node:events:514:28)')
    Assert-Case "[error] [VM:start] Plan9 mount failure" ($r -eq 'Plan9 mount failed') "returned '$r'"

    # The signature can sit on a continuation line of an [error] entry.
    $r = Invoke-Scan -File 'cowork_vm_node1.log' -Lines @(
        '2026-09-11 18:48:26 [error] [VM:start] Startup failed: Error: see message',
        "  [message]: 'RPC error -1: failed to ensure virtiofs mount'")
    Assert-Case "signature on a continuation line of an [error] entry" ($r -eq 'failed to ensure virtiofs mount') "returned '$r'"

    # main.log stays in scope for an [error] entry from the app itself.
    $r = Invoke-Scan -File 'main.log' -Lines @(
        '2026-11-21 18:05:00 [error] Failed to start workspace: HCS operation failed: 0x80370109')
    Assert-Case "[error] HCS failure in main.log" ($r -eq 'HCS operation failed') "returned '$r'"

    # A harmless [info] line in front of a real error in the same chunk.
    $r = Invoke-Scan -File 'main.log' -Lines @(
        $nov18[0],
        '2026-11-21 18:06:00 [error] [VM:start] Startup failed: Error: RPC error -1: guest connection failed')
    Assert-Case "a real error after a harmless line in one chunk" ($r -eq 'RPC error -1') "returned '$r'"
}
finally {
    # Remove only the folder this script made: under the temp root, named by
    # this script, created above. Best effort: every case is counted by
    # now, and a folder left in %TEMP% changes nothing.
    $leaf = Split-Path $ClaudeLogDir -Leaf
    if ($leaf -like 'cf-logtrig-*' -and $ClaudeLogDir.StartsWith($tempBase, [System.StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $ClaudeLogDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ""
Write-Host ("Log triggers: {0} passed, {1} failed" -f $pass, $fail)
if ($fail -eq 0) { Write-Host "Log trigger gate PASSED"; exit 0 }
Write-Host "Log trigger gate FAILED"
exit 1
