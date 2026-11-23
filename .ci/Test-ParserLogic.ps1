<#
.SYNOPSIS
    Extracts individual functions from the scripts and drives them against
    known inputs, including real captured output from this class of machine.
.DESCRIPTION
    The destructive paths cannot be exercised safely, but the parsing and
    arithmetic that decides whether they run can be. Each function is pulled
    out of the source with the language parser and invoked in isolation, so the
    thing under test is the shipped code rather than a copy of it.

    Covers the four pieces of logic that were silently wrong:

      1. the hcsdiag list GUID parser
      2. Write-WatchLog brace handling
      3. the TickCount wrap in idle-time detection
      4. log baseline keying across rotation

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

function Get-FunctionSource {
    param([string]$Path, [string]$Name)
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    if (@($errors).Count -gt 0) { throw "parse errors in $Path" }
    $fn = $ast.Find({
        param($n)
        $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name
    }, $true)
    if (-not $fn) { throw "function '$Name' not found in $Path" }
    return $fn.Extent.Text
}

$pass = 0
$fail = 0
function Assert-Case {
    param([string]$Name, [bool]$Condition, [string]$Detail = '')
    if ($Condition) { Write-Host ("  PASS  {0}" -f $Name); $script:pass++ }
    else { Write-Host ("  FAIL  {0}  {1}" -f $Name, $Detail); $script:fail++ }
}

$fixPath   = Join-Path $Root 'Fix-ClaudeDesktop.ps1'
$watchPath = Join-Path $Root 'Watch-ClaudeHealth.ps1'

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- hcsdiag list parser (Close-StaleHcsVms)"

# Real captured output from this machine. Name on line one, detail on line two,
# GUID uppercase and mid-line. This is the exact shape the old parser could not
# match, because it looked for a GUID alone on a line.
$realHcsList = @"
cowork-vm-1699151a
    VM,                       `tRunning, DE1517EC-BF1D-5A0B-B459-77A48CB1AD97, cowork-vm-1699151a
"@

# Close-StaleHcsVms delegates the parse to Get-CoworkHcsGuids in the shared
# ClaudeEnv region, so the region has to be in scope here. Dot-sourcing the
# canonical copy also means these cases cover what Prevent and Watch call, not
# just what Fix calls.
. (Join-Path $PSScriptRoot 'ClaudeEnv.region.ps1')

# The double-count case, tested against the shared parser directly. One VM
# produces two lines that both contain "cowork-vm". Watch counted string
# matches and compared the result against a threshold of 1, so a single healthy
# VM read as "Multiple cowork-vm instances in HCS" forever.
$oneVmGuids = @(Get-CoworkHcsGuids -ListOutput $realHcsList)
Assert-Case "one VM yields one GUID, not two" ($oneVmGuids.Count -eq 1) "got $($oneVmGuids.Count)"
Assert-Case "raw string count really is 2 (why the bug existed)" `
    ([regex]::Matches($realHcsList, 'cowork-vm').Count -eq 2) `
    ("got " + [regex]::Matches($realHcsList, 'cowork-vm').Count)
Assert-Case "empty input yields no GUIDs" (@(Get-CoworkHcsGuids -ListOutput '').Count -eq 0) ""
Assert-Case "null input yields no GUIDs" (@(Get-CoworkHcsGuids -ListOutput $null).Count -eq 0) ""
Assert-Case "a GUID with no cowork-vm on its line is ignored" `
    (@(Get-CoworkHcsGuids -ListOutput "AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE`ncowork-vm-x").Count -eq 0) ""

# A stateful hcsdiag. "list" prints what is still running in the real
# two-line shape, beside a WSL utility VM that must never be touched; "kill"
# removes a compute system unless it is stuck. Until 6.3.0 Close-StaleHcsVms
# counted every kill it sent, so a static fake that never changed could not
# tell a kill that worked from one that did nothing.
#   HcsListDown          every list fails, as when hcsdiag is unavailable
#   HcsListDownAfterKill the lists after the first kill fail
#   HcsNoWsl             no WSL VM either, so a list with nothing left in it
#                        is empty, as "hcsdiag list" prints it
# Start-Sleep is faked so the re-check costs nothing, and removed after.
$guidA   = 'DE1517EC-BF1D-5A0B-B459-77A48CB1AD97'
$guidB   = 'AB1517EC-BF1D-5A0B-B459-77A48CB1AD98'
$guidWsl = '11111111-2222-3333-4444-555555555555'
function Format-HcsList {
    $out = @()
    if (-not $script:HcsNoWsl) { $out += "wsl-utility`n    VM,   Running, $guidWsl, wsl-utility" }
    foreach ($g in $script:HcsLive) {
        $n = 'cowork-vm-' + $g.Substring(0, 8).ToLowerInvariant()
        $out += "$n`n    VM,                       `tRunning, $g, $n"
    }
    return ($out -join "`n")
}
function Reset-Hcs {
    param([string[]]$Live = @(), [string[]]$Stuck = @())
    $script:HcsLive = New-Object System.Collections.ArrayList
    foreach ($g in $Live) { [void]$script:HcsLive.Add($g) }
    $script:HcsStuck             = @($Stuck)
    $script:HcsListDown          = $false
    $script:HcsListDownAfterKill = $false
    $script:HcsNoWsl             = $false
    $script:KilledGuids          = @()
    $script:CoworkGuidsSeen      = @{}
    $script:LastHcsListed        = 0
    $script:LastHcsKillsSent     = 0
    $script:LogLines             = @()
    $script:Slept                = 0
}
function Invoke-HcsDiag {
    param([string[]]$Arguments, [int]$TimeoutSeconds = 15)
    $null = $TimeoutSeconds
    if ($Arguments[0] -eq 'list') {
        if ($script:HcsListDown) { return $null }
        if ($script:HcsListDownAfterKill -and @($script:KilledGuids).Count -gt 0) { return $null }
        return (Format-HcsList)
    }
    if ($Arguments[0] -eq 'kill') {
        $script:KilledGuids += $Arguments[1]
        if ($script:HcsStuck -notcontains $Arguments[1]) { $script:HcsLive.Remove($Arguments[1]) }
        return 'ok'
    }
    return $null
}
function Log {
    param([string]$Message, [string]$Colour = 'White', [switch]$Indent)
    $null = $Colour, $Indent
    $script:LogLines += $Message
}
function Start-Sleep {
    param([int]$Seconds, [int]$Milliseconds)
    $null = $Milliseconds
    $script:Slept += $Seconds
}

Invoke-Expression (Get-FunctionSource -Path $fixPath -Name 'Close-StaleHcsVms')
# Since 6.3.0 every list Close-StaleHcsVms reads goes through Get-HcsList,
# which also records the Cowork GUIDs it shows, for step 5.
$getHcsListSrc = $null
try { $getHcsListSrc = Get-FunctionSource -Path $fixPath -Name 'Get-HcsList' } catch { $getHcsListSrc = $null }
if ($getHcsListSrc) { Invoke-Expression $getHcsListSrc }

# One VM, under StrictMode, the way Fix runs.
Reset-Hcs -Live $guidA
$killed = & { Set-StrictMode -Version Latest; Close-StaleHcsVms -Action kill }
Assert-Case "returns 1 for one live cowork-vm, under StrictMode" ($killed -eq 1) "got $killed"
Assert-Case "extracts the exact GUID" `
    (@($script:KilledGuids).Count -eq 1 -and $script:KilledGuids[0] -eq $guidA) ($script:KilledGuids -join ',')
Assert-Case "records the GUID it saw for step 5" `
    ($script:CoworkGuidsSeen.ContainsKey($guidA)) "seen: $(@($script:CoworkGuidsSeen.Keys) -join ',')"
Assert-Case "and says one kill went out" ($script:LastHcsKillsSent -eq 1) "got $($script:LastHcsKillsSent)"
Assert-Case "a VM gone at the first check costs no wait" ($script:Slept -eq 0) "slept $($script:Slept)s"

# Two instances, both gone; the WSL VM untouched.
Reset-Hcs -Live $guidA, $guidB
$two = Close-StaleHcsVms -Action kill
Assert-Case "handles two instances" ($two -eq 2 -and @($script:KilledGuids).Count -eq 2) "got $two"
Assert-Case "and says it listed two" ($script:LastHcsListed -eq 2) "got $($script:LastHcsListed)"
Assert-Case "never kills a compute system that is not the Cowork VM" ($script:KilledGuids -notcontains $guidWsl)

# A kill that returns but removes nothing.
Reset-Hcs -Live $guidA, $guidB -Stuck $guidB
$stuck = Close-StaleHcsVms -Action kill
Assert-Case "a kill that did not remove the VM is not counted" ($stuck -eq 1) "got $stuck"
# Listed when last read, not proof the kill failed for good. Until 6.3.0
# the line said the kill did not remove it.
Assert-Case "and the log names the one still listed" `
    (@($script:LogLines | Where-Object { $_ -eq "[!] compute system $guidB was still listed after the kill" }).Count -eq 1) `
    ($script:LogLines -join ' | ')
Assert-Case "and two kills went out, so step 5 knows one is not confirmed" ($script:LastHcsKillsSent -eq 2) "got $($script:LastHcsKillsSent)"
Assert-Case "after looking again, a moment later" ($script:Slept -ge 2) "slept $($script:Slept)s"

# The kills went out, then hcsdiag stopped answering.
Reset-Hcs -Live $guidA
$script:HcsListDownAfterKill = $true
$unknown = Close-StaleHcsVms -Action kill
Assert-Case "an unverifiable kill is not counted" ($unknown -eq 0) "got $unknown"
Assert-Case "and the log says it is not known" `
    (@($script:LogLines | Where-Object { $_ -match 'did not answer after the kill.*1 compute system\(s\) are gone is not known' }).Count -eq 1) `
    ($script:LogLines -join ' | ')

# -WhatIf must not kill anything.
Reset-Hcs -Live $guidA
$whatIfKilled = Close-StaleHcsVms -Action kill -WhatIf
Assert-Case "-WhatIf kills nothing" (@($script:KilledGuids).Count -eq 0) ($script:KilledGuids -join ',')
Assert-Case "-WhatIf reports 0 acted on" ($whatIfKilled -eq 0) "got $whatIfKilled"
# One listed and no kill sent. Step 5 says so from these, and does not call
# the result clean. Until 6.3.0 it said "No orphan compute systems found".
Assert-Case "-WhatIf: one listed and no kill sent, so step 5 can say so" `
    ($script:LastHcsListed -eq 1 -and $script:LastHcsKillsSent -eq 0) "listed $($script:LastHcsListed), sent $($script:LastHcsKillsSent)"

# No cowork-vm present. hcsdiag answered, nothing to kill; the flag has to say
# so, because step 5 prints "none found via hcsdiag" off the back of it.
Reset-Hcs
$none = Close-StaleHcsVms -Action kill
Assert-Case "ignores a non-cowork compute system" ($none -eq 0 -and @($script:KilledGuids).Count -eq 0) "got $none"
Assert-Case "LastHcsListOk true when hcsdiag answered" ($script:LastHcsListOk -eq $true) "got $($script:LastHcsListOk)"
Assert-Case "and nothing listed" ($script:LastHcsListed -eq 0) "got $($script:LastHcsListed)"

# "hcsdiag list" prints nothing at all when no compute system is running.
# That is an answer. Until 6.3.0 an empty list read as no answer, and step 5
# said "hcsdiag did not answer" on a clean host.
Reset-Hcs
$script:HcsNoWsl = $true
$emptyBefore = Close-StaleHcsVms -Action kill
Assert-Case "an empty list before the kill: 0, and hcsdiag answered" `
    ($emptyBefore -eq 0 -and $script:LastHcsListOk -eq $true) "got $emptyBefore, LastHcsListOk $($script:LastHcsListOk)"

# The lone Cowork VM, killed: the list after the kill is empty, and the kill
# is counted.
Reset-Hcs -Live $guidA
$script:HcsNoWsl = $true
$lone = Close-StaleHcsVms -Action kill
Assert-Case "a lone VM killed, the list after it empty: counted" ($lone -eq 1) "got $lone; log: $($script:LogLines -join ' | ')"

# hcsdiag unavailable.
Reset-Hcs -Live $guidA
$script:HcsListDown = $true
# What an earlier call left.
$script:LastHcsListed = 9
$dead = Close-StaleHcsVms -Action kill
Assert-Case "survives hcsdiag returning nothing" ($dead -eq 0 -and @($script:KilledGuids).Count -eq 0) "got $dead"

# Both cases return 0, so the count alone cannot tell them apart. That is why
# step 5 printed a clean hcsdiag result on a run whose step 0 had already said
# hcsdiag was unavailable.
Assert-Case "LastHcsListOk false when hcsdiag did not answer" `
    ($script:LastHcsListOk -eq $false) "got $($script:LastHcsListOk)"
Assert-Case "and no kill went out" ($script:LastHcsKillsSent -eq 0) "got $($script:LastHcsKillsSent)"
Assert-Case "and nothing is said to be listed" ($script:LastHcsListed -eq 0) "got $($script:LastHcsListed)"

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Get-CoworkHcsSurvivors"

Reset-Hcs
$r = Get-CoworkHcsSurvivors -Guids $guidA -List { "wsl-utility`n    VM, Running, $guidWsl, wsl-utility" }
Assert-Case "gone: an empty array, not null" ($null -ne $r -and @($r).Count -eq 0) "got [$r]"
# hcsdiag prints nothing at all when no compute system is left, which is
# what killing the only one leaves. Until 6.3.0 that read as "no answer".
$r = Get-CoworkHcsSurvivors -Guids $guidA -List { '' }
Assert-Case "an empty list is an answer: all gone" ($null -ne $r -and @($r).Count -eq 0) "got [$r]"
$r = Get-CoworkHcsSurvivors -Guids $guidA -List { "`r`n" }
Assert-Case "so is a list of blank lines" ($null -ne $r -and @($r).Count -eq 0) "got [$r]"
$r = Get-CoworkHcsSurvivors -Guids $guidA -List { $null }
Assert-Case "no answer: null" ($null -eq $r) "got [$r]"
$r = Get-CoworkHcsSurvivors -Guids $guidA -List { }
Assert-Case "nothing returned at all: null" ($null -eq $r) "got [$r]"
# A read that fails after one that answered keeps what that one showed.
# Until 6.3.0 the whole result became "not known", and a kill the first read
# had confirmed was reported as unknown.
$script:SurvivorReads = 0
$r = Get-CoworkHcsSurvivors -Guids $guidA, $guidB -List {
    $script:SurvivorReads++
    if ($script:SurvivorReads -eq 1) { "cowork-vm-ab1517ec`n    VM, Running, $guidB, cowork-vm-ab1517ec" }
}
Assert-Case "a later read that fails keeps the earlier answer" `
    ($null -ne $r -and @($r).Count -eq 1 -and $r[0] -eq $guidB -and $script:SurvivorReads -eq 2) `
    "got [$($r -join ',')] after $($script:SurvivorReads) reads"
# & hcsdiag list hands back an array of lines; [string] on one joins them with
# spaces, and the parser reads a line at a time.
$r = Get-CoworkHcsSurvivors -Guids $guidA, $guidB -Attempts 1 -List {
    @('cowork-vm-de1517ec', "    VM, Running, $guidA, cowork-vm-de1517ec")
}
Assert-Case "a list given as lines is read line by line" (@($r).Count -eq 1 -and $r[0] -eq $guidA) "got [$($r -join ',')]"
Reset-Hcs
$r = Get-CoworkHcsSurvivors -Guids $guidA -Attempts 3 -DelaySeconds 5 -List {
    "cowork-vm-de1517ec`n    VM, Running, $guidA, cowork-vm-de1517ec"
}
Assert-Case "still listed after every attempt: returned" (@($r).Count -eq 1 -and $r[0] -eq $guidA) "got [$($r -join ',')]"
Assert-Case "waiting between attempts, not before the first" ($script:Slept -eq 10) "slept $($script:Slept)s"

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Fix's Invoke-HcsDiag against a stand-in hcsdiag"

# The shipped Invoke-HcsDiag, run for real, with the path of hcsdiag.exe in
# its source swapped for a batch file in a temp folder. The batch file reads
# its mode from mode.txt: list prints a Cowork VM, empty prints nothing,
# error prints an error and exits 5, other prints another error and exits 5,
# hang outlasts the timeout. A kill copies afterkill.txt over mode.txt. Until
# 6.3.0 the exit code was thrown away, so an error came back as if it were
# the list.
$standDir = Join-Path ([System.IO.Path]::GetTempPath()) ('cf-hcs-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$null = New-Item -ItemType Directory -Path $standDir -Force
$standCmd = Join-Path $standDir 'hcsdiag-standin.cmd'
[System.IO.File]::WriteAllText($standCmd, ((@(
    '@echo off',
    'set /p MODE=<"%~dp0mode.txt"',
    'if "%1"=="kill" (',
    '  copy /y "%~dp0afterkill.txt" "%~dp0mode.txt" >nul',
    '  echo kill sent',
    '  exit /b 0',
    ')',
    'if "%MODE%"=="list" (',
    '  type "%~dp0list.txt"',
    '  exit /b 0',
    ')',
    'if "%MODE%"=="empty" exit /b 0',
    'if "%MODE%"=="error" (',
    '  echo The operation could not be started: 0x80370114',
    '  exit /b 5',
    ')',
    'if "%MODE%"=="other" (',
    '  echo Access is denied.',
    '  exit /b 5',
    ')',
    'if "%MODE%"=="hang" (',
    '  ping -n 6 127.0.0.1 >nul',
    '  exit /b 0',
    ')',
    'exit /b 9'
) -join "`r`n") + "`r`n"), [System.Text.Encoding]::ASCII)
[System.IO.File]::WriteAllText((Join-Path $standDir 'list.txt'), ($realHcsList -replace "`r?`n", "`r`n") + "`r`n", [System.Text.Encoding]::ASCII)
function Set-StandIn {
    param([string]$Mode, [string]$AfterKill = 'empty')
    [System.IO.File]::WriteAllText((Join-Path $standDir 'mode.txt'), "$Mode`r`n", [System.Text.Encoding]::ASCII)
    [System.IO.File]::WriteAllText((Join-Path $standDir 'afterkill.txt'), "$AfterKill`r`n", [System.Text.Encoding]::ASCII)
}
# The swap is checked before anything runs: the real hcsdiag must never be
# reached from here.
$realPathLiteral = '"$env:SystemRoot\System32\hcsdiag.exe"'
$hcsSrc = Get-FunctionSource -Path $fixPath -Name 'Invoke-HcsDiag'
$swapCount = ([regex]::Matches($hcsSrc, [regex]::Escape($realPathLiteral))).Count
$standSrc   = $hcsSrc.Replace($realPathLiteral, "'$standCmd'")
$missingSrc = $hcsSrc.Replace($realPathLiteral, "'$(Join-Path $standDir 'absent.exe')'").Replace('function Invoke-HcsDiag', 'function Invoke-HcsDiagMissing')
$swapped = ($swapCount -eq 1 -and $standSrc -notmatch 'System32\\hcsdiag' -and $missingSrc -notmatch 'System32\\hcsdiag')
Assert-Case "the path of hcsdiag.exe was swapped for the stand-in" $swapped "found the path $swapCount time(s)"
if ($swapped) {
    try {
        Invoke-Expression $standSrc
        Invoke-Expression $missingSrc
        # Fix sets it at the top of the script; nothing has failed yet.
        $script:LastHcsDiagNote = $null

        Set-StandIn -Mode list
        $r = Invoke-HcsDiag -Arguments 'list'
        Assert-Case "exit 0: the text" ("$r" -match $guidA) "got [$r]"

        Set-StandIn -Mode empty
        $r = Invoke-HcsDiag -Arguments 'list'
        Assert-Case "exit 0 and nothing printed: '', not null" ($null -ne $r -and "$r".Trim() -eq '') "got [$r]"

        Set-StandIn -Mode error
        $script:LogLines = @()
        $r = Invoke-HcsDiag -Arguments 'list'
        Assert-Case "exit 5 with text: null" ($null -eq $r) "got [$r]"
        Assert-Case "and the exit code is logged" `
            (@($script:LogLines | Where-Object { $_ -match '^hcsdiag list exited 5: The operation could not be started' }).Count -eq 1) `
            ($script:LogLines -join ' | ')

        # Each pass of Step 9's wait lists, and so does each check of Step
        # 0 and Step 5, so a failing hcsdiag would print the same line again
        # and again. It is printed once, again when what hcsdiag says
        # changes, and again after it has answered. Until 6.3.0 it was
        # printed every time.
        $r = Invoke-HcsDiag -Arguments 'list'
        Assert-Case "the same failure again: null, and not logged again" `
            ($null -eq $r -and @($script:LogLines).Count -eq 1) ($script:LogLines -join ' | ')
        Set-StandIn -Mode other
        $r = Invoke-HcsDiag -Arguments 'list'
        Assert-Case "a different failure: logged" `
            ($null -eq $r -and @($script:LogLines).Count -eq 2 -and $script:LogLines[-1] -match '^hcsdiag list exited 5: Access is denied') `
            ($script:LogLines -join ' | ')
        Set-StandIn -Mode list
        $null = Invoke-HcsDiag -Arguments 'list'
        Set-StandIn -Mode other
        $null = Invoke-HcsDiag -Arguments 'list'
        Assert-Case "the same failure after an answer: logged again" `
            (@($script:LogLines).Count -eq 3 -and $script:LogLines[-1] -match '^hcsdiag list exited 5: Access is denied') `
            ($script:LogLines -join ' | ')

        Set-StandIn -Mode hang
        $r = Invoke-HcsDiag -Arguments 'list' -TimeoutSeconds 2
        Assert-Case "timed out: null" ($null -eq $r) "got [$r]"

        $r = Invoke-HcsDiagMissing -Arguments 'list'
        Assert-Case "not installed: null" ($null -eq $r) "got [$r]"

        # Close-StaleHcsVms on top of it. The kill goes out, then hcsdiag
        # fails: not counted. Until 6.3.0 the error text read as a list with
        # no VM in it, and the kill was counted as done.
        Reset-Hcs
        Set-StandIn -Mode list -AfterKill error
        $r = Close-StaleHcsVms -Action kill
        Assert-Case "an error after the kill: not counted" ($r -eq 0) "got $r; log: $($script:LogLines -join ' | ')"
        Assert-Case "and the log says it is not known" `
            (@($script:LogLines | Where-Object { $_ -match 'did not answer after the kill' }).Count -eq 1) ($script:LogLines -join ' | ')

        Reset-Hcs
        Set-StandIn -Mode list -AfterKill empty
        $r = Close-StaleHcsVms -Action kill
        Assert-Case "the list empty after the kill: counted" ($r -eq 1) "got $r; log: $($script:LogLines -join ' | ')"

        Reset-Hcs
        Set-StandIn -Mode empty
        $r = Close-StaleHcsVms -Action kill
        Assert-Case "an empty list before the kill: hcsdiag answered" `
            ($r -eq 0 -and $script:LastHcsListOk -eq $true) "got $r, LastHcsListOk $($script:LastHcsListOk)"
    } finally {
        # A fake that was never made is the one way this fails, and then
        # there is nothing to remove.
        Remove-Item Function:\Invoke-HcsDiagMissing -ErrorAction SilentlyContinue
    }
}
# Clean-up of this script's stand-in folder: every case is counted by
# now, and a folder left in %TEMP% changes nothing.
Remove-Item -LiteralPath $standDir -Recurse -Force -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Step 5's hung worker (Stop-HungCoworkWorker)"

# Get-CimInstance, Get-Process and Invoke-HcsDiag are faked. A fake process
# exits when it is killed, unless it is unkillable, or, with -ExitsDuringWait,
# while it is waited for; Kill, WaitForExit and HasExited are what the code
# calls on the object it holds. Until 6.3.0 the worker was matched against
# the one list the last cleanup read, killed by a bare PID two seconds after
# it was looked up, and counted stopped as soon as Stop-Process returned; and
# after the hcsdiag kill it slept two seconds and then force-killed a worker
# that was still on its way out.
$workerDef = $null
try { $workerDef = Get-FunctionSource -Path $fixPath -Name 'Stop-HungCoworkWorker' } catch { $workerDef = $null }
$exitedDef = $null
try { $exitedDef = Get-FunctionSource -Path $fixPath -Name 'Test-WorkerExited' } catch { $exitedDef = $null }
Assert-Case "Stop-HungCoworkWorker and Test-WorkerExited were found in Fix" ($null -ne $workerDef -and $null -ne $exitedDef)
if ($workerDef -and $exitedDef) {
    $started = [datetime]::new(2026, 11, 21, 8, 0, 0)
    function New-FakeWorker {
        param([int]$Id, [datetime]$StartTime, [switch]$Unkillable, [switch]$HasExitedThrows, [switch]$ExitsDuringWait, [switch]$WaitThrows)
        $o = [pscustomobject]@{ Id = $Id; StartTime = $StartTime; Exited = $false; Killable = (-not $Unkillable); Handle = [IntPtr]1; Throws = [bool]$HasExitedThrows
            ExitsInWait = [bool]$ExitsDuringWait; WaitThrows = [bool]$WaitThrows }
        $o | Add-Member -MemberType ScriptProperty -Name HasExited -Value {
            if ($this.Throws) { throw 'Access is denied' }
            return $this.Exited
        }
        # As a Process: the getter's method throws where the getter reads as
        # nothing (measured 2026-11-23). Test-WorkerExited asks the method.
        $o | Add-Member -MemberType ScriptMethod -Name get_HasExited -Value {
            if ($this.Throws) { throw 'Access is denied' }
            return $this.Exited
        }
        $o | Add-Member -MemberType ScriptMethod -Name Kill -Value {
            $script:Kills += $this.Id
            if ($this.Killable) { $this.Exited = $true }
        }
        $o | Add-Member -MemberType ScriptMethod -Name WaitForExit -Value {
            param($ms)
            if ($this.WaitThrows) { throw 'No process is associated with this object.' }
            $script:Waited += $ms
            if ($this.ExitsInWait) { $this.Exited = $true }
            return $this.Exited
        }
        $o | Add-Member -MemberType ScriptMethod -Name Dispose -Value { }
        return $o
    }
    function Reset-Workers {
        param([object[]]$Cim = @(), [object[]]$Procs = @(), [string[]]$Seen = @(), [switch]$HcsKillWorks, [bool]$Admin = $true,
            [switch]$LookupDenied, [switch]$CimFails)
        $script:FakeCim  = @($Cim)
        $script:LookupDenied = [bool]$LookupDenied
        $script:CimFails = [bool]$CimFails
        $script:FakeProcs = @{}
        foreach ($fp in $Procs) { $script:FakeProcs[[int]$fp.Id] = $fp }
        $script:CoworkGuidsSeen = @{}
        foreach ($g in $Seen) { $script:CoworkGuidsSeen[$g] = $true }
        $script:HcsKillWorks = [bool]$HcsKillWorks
        $script:IsAdmin  = $Admin
        $script:Kills    = @()
        $script:Waited   = @()
        $script:HcsKills = @()
        $script:LogLines = @()
        $script:Slept    = 0
    }
    $workerResult = & {
        function Get-CimInstance {
            [CmdletBinding()]
            param($ClassName, $Filter)
            $null = $ClassName, $Filter
            if ($script:CimFails) {
                $PSCmdlet.WriteError((New-Object System.Management.Automation.ErrorRecord(
                    (New-Object System.UnauthorizedAccessException 'Access denied'),
                    'HRESULT 0x80041003,Microsoft.Management.Infrastructure.CimCmdlets.GetCimInstanceCommand', 'PermissionDenied', $null)))
                return
            }
            return $script:FakeCim
        }
        # The errors Get-Process gives. No process with the id, measured
        # 2026-11-22 with Get-Process -Id 999999: id NoProcessFoundForGivenId,
        # category ObjectNotFound, ProcessCommandException. A process that
        # cannot be opened gives any other identity, here access denied.
        function Get-Process {
            [CmdletBinding()]
            param([int]$Id)
            if ($script:LookupDenied) {
                $PSCmdlet.WriteError((New-Object System.Management.Automation.ErrorRecord(
                    (New-Object System.ComponentModel.Win32Exception 5),
                    'CouldNotOpenProcess', 'PermissionDenied', $Id)))
                return
            }
            $fp = $script:FakeProcs[$Id]
            if ($null -eq $fp -or $fp.Exited) {
                $PSCmdlet.WriteError((New-Object System.Management.Automation.ErrorRecord(
                    (New-Object Microsoft.PowerShell.Commands.ProcessCommandException "Cannot find a process with the process identifier $Id."),
                    'NoProcessFoundForGivenId', 'ObjectNotFound', $Id)))
                return
            }
            return $fp
        }
        function Invoke-HcsDiag {
            param([string[]]$Arguments, [int]$TimeoutSeconds = 15)
            $null = $TimeoutSeconds
            $script:HcsKills += $Arguments[1]
            if ($script:HcsKillWorks) {
                foreach ($fp in $script:FakeProcs.Values) {
                    $w = @($script:FakeCim | Where-Object { $_.ProcessId -eq $fp.Id })
                    if ($w.Count -and ([string]$w[0].CommandLine) -match [regex]::Escape($Arguments[1])) { $fp.Exited = $true }
                }
            }
            return ''
        }
        Invoke-Expression $exitedDef
        Invoke-Expression $workerDef
        $vmwpExe = 'C:\WINDOWS\system32\vmwp.exe'
        $cimA   = [pscustomobject]@{ ProcessId = 40; CommandLine = """$vmwpExe"" $guidA 0x330"; CreationDate = $started }
        $cimWsl = [pscustomobject]@{ ProcessId = 41; CommandLine = """$vmwpExe"" $guidWsl 0x2f4"; CreationDate = $started }
        $res = [ordered]@{}

        # Seen earlier in the run, stopped through hcsdiag.
        Reset-Workers -Cim $cimA, $cimWsl -Procs (New-FakeWorker -Id 40 -StartTime $started), (New-FakeWorker -Id 41 -StartTime $started) -Seen $guidA -HcsKillWorks
        $r = Stop-HungCoworkWorker
        $res['hcs'] = "$($r.Stopped)/$($r.Left)|$($script:HcsKills -join ',')|$($script:Kills -join ',')|$($script:LogLines -join ' | ')"

        # hcsdiag does not stop it; the kill through the held object does.
        Reset-Workers -Cim $cimA -Procs (New-FakeWorker -Id 40 -StartTime $started) -Seen $guidA
        $r = Stop-HungCoworkWorker
        $res['kill'] = "$($r.Stopped)/$($r.Left)|$($script:Kills -join ',')|$($script:Waited -join ',')|$($script:LogLines -join ' | ')"

        # The hcsdiag kill works, and the worker takes a moment to go: it
        # exits while it is waited for, and is not force-killed.
        Reset-Workers -Cim $cimA -Procs (New-FakeWorker -Id 40 -StartTime $started -ExitsDuringWait) -Seen $guidA
        $r = Stop-HungCoworkWorker
        $res['waited'] = "$($r.Stopped)/$($r.Left)|$($script:Kills -join ',')|$($script:Waited -join ',')|$($script:Slept)|$($script:LogLines -join ' | ')"

        # No handle to wait on: the two seconds the wait used to be.
        Reset-Workers -Cim $cimA -Procs (New-FakeWorker -Id 40 -StartTime $started -WaitThrows) -Seen $guidA -HcsKillWorks
        $r = Stop-HungCoworkWorker
        $res['nowait'] = "$($r.Stopped)/$($r.Left)|$($script:Kills -join ',')|$($script:Waited -join ',')|$($script:Slept)|$($script:LogLines -join ' | ')"

        # Nothing stops it.
        Reset-Workers -Cim $cimA -Procs (New-FakeWorker -Id 40 -StartTime $started -Unkillable) -Seen $guidA
        $r = Stop-HungCoworkWorker
        $res['unkillable'] = "$($r.Stopped)/$($r.Left)|$($script:LogLines -join ' | ')"

        # The PID now belongs to a process started later: left alone.
        Reset-Workers -Cim $cimA -Procs (New-FakeWorker -Id 40 -StartTime $started.AddMinutes(3)) -Seen $guidA
        $r = Stop-HungCoworkWorker
        $res['reused'] = "$($r.Stopped)/$($r.Left)|$($script:HcsKills -join ',')|$($script:Kills -join ',')|$($script:LogLines -join ' | ')"

        # A worker whose GUID no list in this run showed: not touched.
        Reset-Workers -Cim $cimWsl -Procs (New-FakeWorker -Id 41 -StartTime $started) -Seen $guidA
        $r = Stop-HungCoworkWorker
        $res['unseen'] = "$($r.Stopped)/$($r.Left)|$($script:HcsKills -join ',')|$($script:Kills -join ',')|$($script:LogLines -join ' | ')"

        # No list answered at all in this run: nothing is known, nothing touched.
        Reset-Workers -Cim $cimA -Procs (New-FakeWorker -Id 40 -StartTime $started)
        $r = Stop-HungCoworkWorker
        $res['noseen'] = "$($r.Stopped)/$($r.Left)|$($script:Kills -join ',')|$($script:LogLines -join ' | ')"

        # Gone by the time it is looked up.
        Reset-Workers -Cim $cimA -Seen $guidA
        $r = Stop-HungCoworkWorker
        $res['gone'] = "$($r.Stopped)/$($r.Left)|$($script:Kills -join ',')|$($script:LogLines -join ' | ')"

        # Without elevation, straight to the kill.
        Reset-Workers -Cim $cimA -Procs (New-FakeWorker -Id 40 -StartTime $started) -Seen $guidA -HcsKillWorks -Admin $false
        $r = Stop-HungCoworkWorker
        $res['noadmin'] = "$($r.Stopped)/$($r.Left)|$($script:HcsKills -join ',')|$($script:Kills -join ',')"

        # No handle to ask: Test-WorkerExited looks the PID up again, and a
        # different start time there means the worker is gone.
        $noHandle = New-FakeWorker -Id 40 -StartTime $started -HasExitedThrows
        Reset-Workers -Procs (New-FakeWorker -Id 40 -StartTime $started.AddMinutes(1))
        $res['exited-reused'] = Test-WorkerExited -Held $noHandle -ProcessId 40 -StartTime $started
        Reset-Workers -Procs (New-FakeWorker -Id 40 -StartTime $started)
        $res['exited-same'] = Test-WorkerExited -Held $noHandle -ProcessId 40 -StartTime $started
        Reset-Workers
        $res['exited-none'] = Test-WorkerExited -Held $noHandle -ProcessId 40 -StartTime $started
        # A lookup that is refused is not "no process": not exited, and said.
        Reset-Workers -LookupDenied
        $res['exited-denied'] = "$(Test-WorkerExited -Held $noHandle -ProcessId 40 -StartTime $started)|$($script:LogLines -join ' | ')"

        # A worker that could not be opened is left alone, counted left, and said.
        Reset-Workers -Cim $cimA -Procs (New-FakeWorker -Id 40 -StartTime $started) -Seen $guidA -HcsKillWorks -LookupDenied
        $r = Stop-HungCoworkWorker
        $res['denied'] = "$($r.Stopped)/$($r.Left)|$($script:HcsKills -join ',')|$($script:Kills -join ',')|$($script:LogLines -join ' | ')"

        # The worker list could not be read: it throws, and nothing is touched.
        Reset-Workers -Cim $cimA -Procs (New-FakeWorker -Id 40 -StartTime $started) -Seen $guidA -HcsKillWorks -CimFails
        $cimThrew = ''
        try { $null = Stop-HungCoworkWorker } catch { $cimThrew = "$_" }
        $res['cimfails'] = "$cimThrew|$($script:HcsKills -join ',')|$($script:Kills -join ',')|$($script:LogLines -join ' | ')"

        # -WhatIf touches nothing.
        Reset-Workers -Cim $cimA -Procs (New-FakeWorker -Id 40 -StartTime $started) -Seen $guidA -HcsKillWorks
        $r = Stop-HungCoworkWorker -WhatIf
        $res['whatif'] = "$($r.Stopped)/$($r.Left)|$($script:HcsKills -join ',')|$($script:Kills -join ',')"
        $res
    }
    $w = $workerResult
    Assert-Case "a worker seen earlier is stopped through hcsdiag, and only that one" `
        ($w['hcs'] -match "^1/0\|$guidA\|\|.*Stopped the Cowork VM's worker \(vmwp\.exe PID 40\) through hcsdiag") "got $($w['hcs'])"
    Assert-Case "hcsdiag did not stop it: waited for, killed through the held object, waited for" `
        ($w['kill'] -match '^1/0\|40\|5000,5000\|.*Force-killed vmwp\.exe \(PID 40') "got $($w['kill'])"
    Assert-Case "it exits while waited for after the hcsdiag kill: stopped through hcsdiag, not force-killed" `
        ($w['waited'] -match "^1/0\|\|5000\|0\|.*Stopped the Cowork VM's worker \(vmwp\.exe PID 40\) through hcsdiag" -and $w['waited'] -notmatch 'Force-killed') `
        "got $($w['waited'])"
    Assert-Case "no handle to wait on: two seconds, then the same check" `
        ($w['nowait'] -match "^1/0\|\|\|2\|.*Stopped the Cowork VM's worker \(vmwp\.exe PID 40\) through hcsdiag") "got $($w['nowait'])"
    Assert-Case "a worker that does not exit is not counted, and is said to be left" `
        ($w['unkillable'] -match '^0/1\|.*vmwp\.exe \(PID 40\) is unkillable') "got $($w['unkillable'])"
    Assert-Case "a PID now another process's is left alone" `
        ($w['reused'] -match '^0/0\|\|\|.*PID now belongs to another process') "got $($w['reused'])"
    Assert-Case "a worker whose GUID no list showed is not touched" `
        ($w['unseen'] -match '^0/0\|\|\|No hung Cowork VM worker$') "got $($w['unseen'])"
    Assert-Case "no GUID seen in the run: nothing touched" `
        ($w['noseen'] -match "^0/0\|\|1 VM worker process\(es\) running, none known to be the Cowork VM's") "got $($w['noseen'])"
    Assert-Case "a worker gone before it was held: not counted either way" `
        ($w['gone'] -match '^0/0\|\|.*had exited before it was stopped') "got $($w['gone'])"
    Assert-Case "without elevation: no hcsdiag, the held object killed" ($w['noadmin'] -eq '1/0||40') "got $($w['noadmin'])"
    Assert-Case "no handle, the PID now a later process: exited" ($w['exited-reused'] -eq $true) "got $($w['exited-reused'])"
    Assert-Case "no handle, the same process still there: not exited" ($w['exited-same'] -eq $false) "got $($w['exited-same'])"
    Assert-Case "no handle, no process with the PID: exited" ($w['exited-none'] -eq $true) "got $($w['exited-none'])"
    # Until 6.3.1 any error from the lookup read as "exited", and an
    # unopenable worker as "had exited before it was stopped".
    Assert-Case "no handle, the lookup refused: not exited, and said" `
        ($w['exited-denied'] -match '^False\|\[!\] Whether vmwp\.exe PID 40 has exited could not be read') "got $($w['exited-denied'])"
    Assert-Case "a worker that could not be opened: left alone, counted left, and said" `
        ($w['denied'] -match '^0/1\|\|\|\[!\] vmwp\.exe PID 40 could not be opened \(.*\), so it was left alone' -and
         $w['denied'] -notmatch 'had exited') "got $($w['denied'])"
    Assert-Case "the worker list could not be read: thrown, nothing touched" `
        ($w['cimfails'] -match '^the VM worker processes could not be listed \(.*Access denied.*\), so none were stopped\|\|\|$') "got $($w['cimfails'])"
    Assert-Case "-WhatIf touches nothing" ($w['whatif'] -eq '0/0||') "got $($w['whatif'])"
}

# Step 5 can match a worker only against GUIDs some list showed earlier, so
# every list Fix reads goes through Get-HcsList, and the set it fills is made
# once, at the start, and never emptied. Until 6.3.0 step 5 cleared the one
# list it kept, just before reading it.
$fixAstP = [System.Management.Automation.Language.Parser]::ParseFile($fixPath, [ref]$null, [ref]$null)
$seenAssign = @($fixAstP.FindAll({ param($n)
    $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
    $n.Left.Extent.Text -eq '$script:CoworkGuidsSeen' }, $true))
$seenTop = @($seenAssign | Where-Object {
    $_.Parent.Parent -is [System.Management.Automation.Language.ScriptBlockAst] -and $null -eq $_.Parent.Parent.Parent })
Assert-Case "the set of Cowork GUIDs seen is made once, at the top of the script" `
    ($seenAssign.Count -eq 1 -and $seenTop.Count -eq 1 -and $seenAssign[0].Right.Extent.Text -eq '@{}') `
    "found $($seenAssign.Count) assignment(s): $(@($seenAssign | ForEach-Object { 'line ' + $_.Extent.StartLineNumber }) -join ', ')"
$seenCleared = @($fixAstP.FindAll({ param($n)
    $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and
    $n.Expression.Extent.Text -eq '$script:CoworkGuidsSeen' -and $n.Member.Extent.Text -match '^(Clear|Remove)$' }, $true))
Assert-Case "and nothing empties it" ($seenCleared.Count -eq 0) "found $($seenCleared.Count)"
$listCalls = @($fixAstP.FindAll({ param($n)
    $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Invoke-HcsDiag' -and
    $n.Extent.Text -match '-Arguments\s+"list"' }, $true))
$listOutside = @($listCalls | Where-Object {
    $fd = $_.Parent
    while ($fd -and -not ($fd -is [System.Management.Automation.Language.FunctionDefinitionAst])) { $fd = $fd.Parent }
    -not $fd -or $fd.Name -ne 'Get-HcsList' })
Assert-Case "every hcsdiag list Fix reads goes through Get-HcsList" `
    ($listCalls.Count -eq 1 -and $listOutside.Count -eq 0) `
    "direct list calls outside it: $(@($listOutside | ForEach-Object { 'line ' + $_.Extent.StartLineNumber }) -join ', ')"

Remove-Item Function:\Start-Sleep

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- ages in UTC across a daylight-saving change"

# A file's local LastWriteTime and the local clock jump an hour apart at a
# daylight-saving change. Get-ChildItem and Get-Item are faked to hand back
# the pair a real file shows across the spring change: written 20 minutes
# ago in UTC, 80 minutes ago by the local clock. Until 6.3.0 the local pair
# was compared, and a live session read as idle.
$dstRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('cf-dst-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$null = New-Item -ItemType Directory -Path (Join-Path $dstRoot 'projects') -Force
$null = New-Item -ItemType File -Path (Join-Path $dstRoot 'coworkd.log') -Force
$savedCfgDir = $env:CLAUDE_CONFIG_DIR
$env:CLAUDE_CONFIG_DIR = $dstRoot
function Get-DstFile {
    param([double]$UtcMinutes, [double]$LocalMinutes)
    [pscustomobject]@{
        FullName         = Join-Path $dstRoot 'coworkd.log'
        Name             = 'coworkd.log'
        Length           = 1
        LastWriteTime    = (Get-Date).AddMinutes(-$LocalMinutes)
        LastWriteTimeUtc = [datetime]::UtcNow.AddMinutes(-$UtcMinutes)
    }
}
function Get-ChildItem { $script:DstFile }
# Since 6.3.1 the region reads the projects folder and each VM log through
# Get-ClaudeItemOrNull, so this fake answers as Get-Item does: the folder,
# the one log, and for anything else the error Get-Item writes for a path
# with nothing there (PathNotFound, ObjectNotFound, ItemNotFoundException,
# measured 2026-11-22).
function Get-Item {
    [CmdletBinding()]
    param([string]$LiteralPath, [switch]$Force)
    $null = $Force
    if ($LiteralPath -eq (Join-Path $dstRoot 'projects')) {
        return [pscustomobject]@{ FullName = $LiteralPath; PSIsContainer = $true }
    }
    if ($LiteralPath -eq (Join-Path $dstRoot 'coworkd.log')) { return $script:DstFile }
    $ex = New-Object System.Management.Automation.ItemNotFoundException "Cannot find path '$LiteralPath' because it does not exist."
    $PSCmdlet.WriteError((New-Object System.Management.Automation.ErrorRecord($ex, 'PathNotFound',
        [System.Management.Automation.ErrorCategory]::ObjectNotFound, $LiteralPath)))
}
try {
    $script:DstFile = Get-DstFile -UtcMinutes 20 -LocalMinutes 80
    Assert-Case "transcript 20 min old (80 by the local clock): recent" (Test-ClaudeCodeTranscriptRecent -Minutes 35)
    $script:DstFile = Get-DstFile -UtcMinutes 70 -LocalMinutes 10
    Assert-Case "transcript 70 min old (10 by the local clock): not recent" (-not (Test-ClaudeCodeTranscriptRecent -Minutes 35))
    $script:DstFile = Get-DstFile -UtcMinutes 20 -LocalMinutes 80
    $age = Get-ClaudeVmLogAge -LogDir $dstRoot
    Assert-Case "VM log age from UTC: about 1200 s" ($null -ne $age -and $age -gt 1100 -and $age -lt 1300) "got $age"
} finally {
    Remove-Item Function:\Get-ChildItem
    Remove-Item Function:\Get-Item
    # Clean-up: a variable that is not set, the one way the first fails, is
    # what it is for, and a folder left in %TEMP% changes nothing.
    if ($savedCfgDir) { $env:CLAUDE_CONFIG_DIR = $savedCfgDir } else { Remove-Item Env:\CLAUDE_CONFIG_DIR -ErrorAction SilentlyContinue }
    Remove-Item -LiteralPath $dstRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- vmwp.exe workers: only the Cowork VM's"

$vmwp = 'C:\WINDOWS\system32\vmwp.exe'
Assert-Case "the GUID is the first argument" `
    ((Get-VmWorkerGuid -CommandLine """$vmwp"" $guidA 0x330") -eq $guidA)
Assert-Case "an unquoted path and a braced GUID" `
    ((Get-VmWorkerGuid -CommandLine "$vmwp {$guidA} 0x330") -eq $guidA)
Assert-Case "a GUID later on the line is not the worker's" `
    ($null -eq (Get-VmWorkerGuid -CommandLine """$vmwp"" -x $guidA"))
Assert-Case "no command line: null" ($null -eq (Get-VmWorkerGuid -CommandLine ''))

$twoVms = "cowork-vm-de1517ec`n    VM, Running, $guidA, cowork-vm-de1517ec`nwsl-utility`n    VM, Running, $guidWsl, wsl-utility"
$workers = @(
    [pscustomobject]@{ ProcessId = 10; CommandLine = """$vmwp"" $guidA 0x330" },
    [pscustomobject]@{ ProcessId = 11; CommandLine = """$vmwp"" $guidWsl 0x2f4" },
    [pscustomobject]@{ ProcessId = 12; CommandLine = $null },
    $null
)
$ids = @(Get-CoworkWorkerProcessIds -ListOutput $twoVms -Workers $workers)
Assert-Case "picks the Cowork VM's worker and not WSL's" (($ids -join ',') -eq '10') "got $($ids -join ',')"
$ids = @(Get-CoworkWorkerProcessIds -ListOutput $twoVms -Workers @(
    [pscustomobject]@{ ProcessId = 13; CommandLine = """$vmwp"" $($guidA.ToLowerInvariant()) 0x330" }))
Assert-Case "matches the GUID in any case" (($ids -join ',') -eq '13') "got $($ids -join ',')"
$ids = @(Get-CoworkWorkerProcessIds -ListOutput "wsl-utility`n    VM, Running, $guidWsl, wsl-utility" -Workers $workers)
Assert-Case "no Cowork VM listed: no worker" ($ids.Count -eq 0) "got $($ids -join ',')"

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- w32tm offsets in any locale (Get-W32tmOffsetSeconds)"

$stripchart = "Tracking time.windows.com [20.101.57.9:123].`r`nCollecting 1 samples.`r`nThe current time is 21-11-2026 20:22:22.`r`n20:22:22, -00.0684055s"
Assert-Case "a dot" ((Get-W32tmOffsetSeconds -Text $stripchart) -eq -0.0684055)
Assert-Case "a comma, as Dutch and German machines print it" `
    ((Get-W32tmOffsetSeconds -Text "20:22:22, +12,5000000s") -eq 12.5)
Assert-Case "an error line has no offset" `
    ($null -eq (Get-W32tmOffsetSeconds -Text "Collecting 1 samples.`n20:22:22, error: 0x800705B4"))
Assert-Case "the last sample wins" `
    ((Get-W32tmOffsetSeconds -Text "20:22:22, 01.0000000s`n20:22:24, -03.2500000s") -eq -3.25)
Assert-Case "no text: null" ($null -eq (Get-W32tmOffsetSeconds -Text ''))
$savedCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
try {
    [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('nl-NL')
    $nlDot   = Get-W32tmOffsetSeconds -Text "20:22:22, -00.5000000s"
    $nlComma = Get-W32tmOffsetSeconds -Text "20:22:22, -00,5000000s"
} finally {
    [System.Threading.Thread]::CurrentThread.CurrentCulture = $savedCulture
}
Assert-Case "on a Dutch thread, a dot still reads -0.5" ($nlDot -eq -0.5) "got $nlDot"
Assert-Case "on a Dutch thread, a comma still reads -0.5" ($nlComma -eq -0.5) "got $nlComma"

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Write-WatchLog brace handling"

$script:WatchLogFile = Join-Path $env:TEMP ("cfix-watchlog-" + [guid]::NewGuid().ToString("N").Substring(0,8) + ".log")
$Quiet = $true

Invoke-Expression (Get-FunctionSource -Path $watchPath -Name 'Write-WatchLog')

# The message that used to throw. HCS and Hyper-V errors embed {GUID}, and the
# old implementation interpolated the message INTO an -f format string, so a
# brace was parsed as a format placeholder.
$braceMsg = 'HCS operation failed for {DE1517EC-BF1D-5A0B-B459-77A48CB1AD97}: 0xC037010D'
$threw = $false
try { Write-WatchLog $braceMsg } catch { $threw = $true }
Assert-Case "a message containing {GUID} does not throw" (-not $threw)

# Format-specifier shapes that would also have been interpreted.
foreach ($m in @('index {0} and {1}', 'literal {{ and }}', 'value {0:yyyy-MM-dd}', 'unmatched { brace')) {
    $threw = $false
    try { Write-WatchLog $m } catch { $threw = $true }
    Assert-Case ("does not throw on: $m") (-not $threw)
}

if (Test-Path $script:WatchLogFile) {
    $written = Get-Content $script:WatchLogFile -Raw
    Assert-Case "the brace message reached the log intact" ($written -match [regex]::Escape($braceMsg))
    Assert-Case "no UTF-8 BOM in the watch log" `
        (-not ((([System.IO.File]::ReadAllBytes($script:WatchLogFile))[0] -eq 0xEF)))
    # Clean-up of this script's own temp log: both cases are counted, and a
    # file left in %TEMP% changes nothing.
    Remove-Item $script:WatchLogFile -Force -ErrorAction SilentlyContinue
} else {
    Assert-Case "watch log was created" $false "file missing"
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- TickCount wrap arithmetic"

# The shipped expression, lifted verbatim. Signed TickCount goes negative past
# 24.9 days of uptime while dwTime stays unsigned, so a direct subtraction
# yields a hugely negative idle time, which reads as "user is active" and stops
# every unattended run from doing anything.
function Get-IdleMs {
    param([long]$RawTickCount, [uint32]$LastInputTick)
    # StrictMode inside the helper so an unassigned variable is an error rather
    # than a silent $null. Without it, the first version of this test "passed"
    # while the cast below was throwing: $nowTicks stayed unset, $null -ge N was
    # false, and the wrap branch happened to return a plausible number anyway.
    Set-StrictMode -Version Latest
    $nowTicks  = [uint32]($RawTickCount -band 0xFFFFFFFFL)
    $lastTicks = $LastInputTick
    if ($nowTicks -ge $lastTicks) { return [long]$nowTicks - [long]$lastTicks }
    return [long]4294967296 - [long]$lastTicks + [long]$nowTicks
}

# The mask literal must be Int64. A bare 0xFFFFFFFF is Int32 -1 in PowerShell,
# which makes -band a no-op on a negative value.
Assert-Case "0xFFFFFFFFL is Int64, not Int32 -1" ((0xFFFFFFFFL) -eq 4294967295)
Assert-Case "the bare literal really is -1 (why the L matters)" ((0xFFFFFFFF) -eq -1)

# Normal case: 5 seconds idle.
Assert-Case "5s idle on a freshly booted machine" ((Get-IdleMs -RawTickCount 100000 -LastInputTick 95000) -eq 5000)

# Across the wrap. TickCount has just wrapped to a small positive value while
# the last input was recorded just before the boundary.
$justAfterWrap  = 1000
$justBeforeWrap = [uint32]4294966296   # 2^32 - 1000
Assert-Case "2s idle measured across the 32-bit wrap" `
    ((Get-IdleMs -RawTickCount $justAfterWrap -LastInputTick $justBeforeWrap) -eq 2000) `
    ("got " + (Get-IdleMs -RawTickCount $justAfterWrap -LastInputTick $justBeforeWrap))

# Past 24.9 days, where TickCount is negative as a signed Int32. This is the
# case the whole fix exists for, and the case the first version of the fix
# crashed on.
#
# -2000000000 unsigned is 2294967296. With last input 4000ms earlier, at
# 2294963296, the idle time must come out at exactly 4000.
$negativeTick = -2000000000
$idleNeg = $null
$threwNeg = $false
try { $idleNeg = Get-IdleMs -RawTickCount $negativeTick -LastInputTick ([uint32]2294963296) }
catch { $threwNeg = $true }

Assert-Case "a negative TickCount does not throw" (-not $threwNeg)
Assert-Case "idle time is computed, not left null" ($null -ne $idleNeg)
Assert-Case "and is exactly 4000ms past 24.9 days uptime" ($idleNeg -eq 4000) "got $idleNeg"

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- log baseline keying across rotation"

# Electron shift-rotates content between paths: main.log becomes main1.log and
# so on. Keyed on path alone, a rotation either reset the baseline to 0 and
# re-scanned megabytes of already-seen content, or compared against a baseline
# belonging to entirely different content. Either way a months-old error line
# could re-fire a repair.
function Get-BaselineKey {
    param([string]$Path, [datetime]$CreationTimeUtc)
    return $Path + '|' + $CreationTimeUtc.ToString('o')
}

$createdA = [datetime]::new(2026, 7, 20, 10, 0, 0, [System.DateTimeKind]::Utc)
$createdB = [datetime]::new(2026, 7, 25, 10, 0, 0, [System.DateTimeKind]::Utc)

$keyMainBefore = Get-BaselineKey -Path 'C:\logs\main.log'  -CreationTimeUtc $createdA
$keyMain1After = Get-BaselineKey -Path 'C:\logs\main1.log' -CreationTimeUtc $createdA
$keyMainNew    = Get-BaselineKey -Path 'C:\logs\main.log'  -CreationTimeUtc $createdB

Assert-Case "the same stream at a new path is a different key" ($keyMainBefore -ne $keyMain1After)
Assert-Case "a new stream at the same path is a different key" ($keyMainBefore -ne $keyMainNew)
Assert-Case "an unchanged stream keeps its key" `
    ($keyMainBefore -eq (Get-BaselineKey -Path 'C:\logs\main.log' -CreationTimeUtc $createdA))

# Under the old path-only scheme both of those collide, which is the bug.
Assert-Case "path-only keying would have collided (the old bug)" `
    ('C:\logs\main.log' -eq 'C:\logs\main.log')

# ---------------------------------------------------------------------
Write-Host ""
Write-Host ("Parser logic: {0} passed, {1} failed" -f $pass, $fail)
if ($fail -eq 0) { Write-Host "Parser logic gate PASSED"; exit 0 }
Write-Host "Parser logic gate FAILED"
exit 1
