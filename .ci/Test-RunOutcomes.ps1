<#
.SYNOPSIS
    How a Fix run ends and what the monitor makes of it, the monitor's clock
    check, its hcsdiag and w32tm helpers, and the ages both scripts measure.
.DESCRIPTION
    A run Fix holds back before it stops the app or the service exits 3, and
    the monitor does not count it as an attempt at a repair. Until 6.3.0 it
    exited 0, the same as a finished repair, so three runs that had repaired
    nothing set off the "Auto-fix failed 3 times" escalation, and each one
    dropped what the monitor had detected. A -Quiet run that stops part-way,
    after Steps 1 and 2 have run, because it found a Claude Code session or
    could not rule one out, exits 4, and the monitor does count that one.

    The pieces of Fix that end a run with exit run in a child PowerShell,
    lifted out of the source with the language parser, with every command
    they call faked; the exit code is the outcome. Before any lifted code
    runs, the commands it calls are checked against the fakes, and it is not
    run when one is missing, so this machine's services and VMs are never
    within reach. hcsdiag and w32tm are batch-file stand-ins in a temp
    folder, swapped into the source in place of the real path, and the swap
    is checked before anything runs.

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

$pass = 0
$fail = 0
function Assert-Case {
    param([string]$Name, [bool]$Condition, [string]$Detail = '')
    if ($Condition) { Write-Host ("  PASS  {0}" -f $Name); $script:pass++ }
    else { Write-Host ("  FAIL  {0}  {1}" -f $Name, $Detail); $script:fail++ }
}

# Get-W32tmOffsetSeconds, which the clock check reads w32tm's output with,
# lives in the shared region.
. (Join-Path $PSScriptRoot 'ClaudeEnv.region.ps1')

$fixPath   = Join-Path $Root 'Fix-ClaudeDesktop.ps1'
$watchPath = Join-Path $Root 'Watch-ClaudeHealth.ps1'
$fixSrc    = [System.IO.File]::ReadAllText($fixPath)
$watchSrc  = [System.IO.File]::ReadAllText($watchPath)
$fixAst    = [System.Management.Automation.Language.Parser]::ParseFile($fixPath, [ref]$null, [ref]$null)
$watchAst  = [System.Management.Automation.Language.Parser]::ParseFile($watchPath, [ref]$null, [ref]$null)

function Get-AstFunction {
    param($Ast, [string]$Name)
    return $Ast.Find({ param($n)
        $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
}

# Fix's main body sits inside one try, so its statements are found at any
# depth, not only at the top level.
function Find-FixStatement {
    param([scriptblock]$Predicate)
    return @($fixAst.FindAll($Predicate, $true))
}

# Every command a piece of lifted code calls that is neither faked nor one of
# the harmless built-ins below, and every command called through a variable
# or a string (& $x), unless named in -AllowDynamic. Empty means it is safe
# to run.
$harmless = @('Get-Date', 'Where-Object', 'Select-Object', 'ForEach-Object', 'Sort-Object',
    'Measure-Object', 'Join-Path', 'Out-Null', 'Out-String', 'Write-Output')
function Get-UnfakedCommand {
    param([object[]]$Asts, [string[]]$Faked, [string[]]$AllowDynamic = @())
    $found = @()
    foreach ($a in $Asts) {
        foreach ($c in @($a.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true))) {
            $name = $c.GetCommandName()
            if ($name) {
                if ($Faked -notcontains $name -and $harmless -notcontains $name) { $found += $name }
            } else {
                $first = $c.CommandElements[0].Extent.Text
                if ($AllowDynamic -notcontains $first) { $found += "<called through $first>" }
            }
        }
    }
    return @($found | Sort-Object -Unique)
}

# Runs a block in this process and hands back the last thing it wrote, so a
# stray line of output cannot take the place of the result; an error in it is
# a failed case, not the end of the run.
function Invoke-Isolated {
    param([scriptblock]$Block, [string]$Name)
    try {
        $r = @(& $Block)
        if ($r.Count -gt 0) { return $r[-1] }
        Assert-Case "$Name wrote a result" $false
    } catch {
        Assert-Case "$Name ran without an error" $false $_.Exception.Message
    }
    return $null
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('cf-outcomes-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$null = New-Item -ItemType Directory -Path $tempRoot -Force

# Runs a script body in a child PowerShell. Returns the exit code and output.
function Invoke-Child {
    param([string]$Body, [string]$Name)
    $path = Join-Path $tempRoot "$Name.ps1"
    [System.IO.File]::WriteAllText($path, $Body, [System.Text.Encoding]::ASCII)
    $out = (& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $path 2>&1 | Out-String)
    return [pscustomobject]@{ Exit = $LASTEXITCODE; Out = $out }
}

# The line that sets the held-back code, lifted so a child uses Fix's own.
$fixCodeMatch   = [regex]::Match($fixSrc, '(?m)^\$script:BlockedExitCode = (\d+)\s*$')
$watchCodeMatch = [regex]::Match($watchSrc, '(?m)^\$FixHeldBackExitCode = (\d+)\s*$')
$fixCodeLine    = $fixCodeMatch.Value.Trim()
$fixPartMatch   = [regex]::Match($fixSrc, '(?m)^\$script:StoppedPartWayExitCode = (\d+)\s*$')
$watchPartMatch = [regex]::Match($watchSrc, '(?m)^\$FixStoppedPartWayExitCode = (\d+)\s*$')

$childFakes = @'
function Log { param([string]$Message, [string]$Colour, [switch]$Indent) $null = $Colour, $Indent; Write-Output "LOG: $Message" }
function Save-Log { Write-Output 'SAVED' }
function Wait-ForAnyKey { Write-Output 'WAITED FOR A KEY' }
'@

# The lines of Fix outside the shared region.
$fixAllLines = $fixSrc -split "\r?\n"
$fixRegionFrom = [array]::IndexOf($fixAllLines, '#region ClaudeEnv') + 1
$fixRegionTo   = [array]::IndexOf($fixAllLines, '#endregion ClaudeEnv') + 1
function Test-OutsideFixRegion {
    param([int]$Line)
    return ($fixRegionFrom -le 0 -or $Line -lt $fixRegionFrom -or $Line -gt $fixRegionTo)
}

$fixLockLiteral = '@("Global\ClaudeDesktopFix", "Local\ClaudeDesktopFix")'

try {

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Fix's exit codes: 3 held back, 4 stopped part-way"

Assert-Case "Fix sets `$script:BlockedExitCode to 3" ($fixCodeMatch.Success -and $fixCodeMatch.Groups[1].Value -eq '3') "found [$($fixCodeMatch.Value)]"
Assert-Case "the monitor reads the same number" `
    ($watchCodeMatch.Success -and $fixCodeMatch.Success -and $watchCodeMatch.Groups[1].Value -eq $fixCodeMatch.Groups[1].Value) `
    "Watch [$($watchCodeMatch.Value)], Fix [$($fixCodeMatch.Value)]"
Assert-Case "Fix sets `$script:StoppedPartWayExitCode to 4" ($fixPartMatch.Success -and $fixPartMatch.Groups[1].Value -eq '4') "found [$($fixPartMatch.Value)]"
Assert-Case "the monitor reads the same number for it" `
    ($watchPartMatch.Success -and $fixPartMatch.Success -and $watchPartMatch.Groups[1].Value -eq $fixPartMatch.Groups[1].Value) `
    "Watch [$($watchPartMatch.Value)], Fix [$($fixPartMatch.Value)]"

# Every exit in Fix outside the shared region, by what it exits with. 3 and
# 4 leave only through the two names above: 3 straight from the -Quiet gate
# and the lock, and both through Assert-NoNewAppSession, which exits with
# the code its caller passes.
$fixExits = @($fixAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.ExitStatementAst] }, $true) |
    Where-Object { Test-OutsideFixRegion -Line $_.Extent.StartLineNumber })
$exitText = @($fixExits | ForEach-Object { if ($_.Pipeline) { $_.Pipeline.Extent.Text } else { '' } })
$heldExits = @($fixExits | Where-Object { $_.Pipeline -and $_.Pipeline.Extent.Text -eq '$script:BlockedExitCode' })
Assert-Case "no exit 3 or exit 4 written as a number" (@($exitText | Where-Object { $_ -in @('3', '4') }).Count -eq 0) ($exitText -join ', ')
Assert-Case "nothing exits with `$script:StoppedPartWayExitCode directly" `
    (@($exitText | Where-Object { $_ -eq '$script:StoppedPartWayExitCode' }).Count -eq 0) ($exitText -join ', ')
Assert-Case "two exits with `$script:BlockedExitCode: the gate and the lock" ($heldExits.Count -eq 2) "found $($heldExits.Count)"
foreach ($e in $heldExits) {
    $k = $e.Extent.StartLineNumber - 1
    $said = $null
    for ($m = $k - 1; $m -ge [math]::Max(0, $k - 10); $m--) {
        if ($fixAllLines[$m] -match '^\s*Log "(BLOCKED: \$blocker|Another Fix instance is already running)') { $said = $Matches[1]; break }
    }
    Assert-Case "line $($k + 1) follows the gate's or the lock's line" ($null -ne $said) $fixAllLines[$k].Trim()
}
$codeExits = @($fixExits | Where-Object { $_.Pipeline -and $_.Pipeline.Extent.Text -eq '$ExitCode' })
$codeExitIn = @($codeExits | ForEach-Object {
    $p = $_.Parent
    while ($p -and $p -isnot [System.Management.Automation.Language.FunctionDefinitionAst]) { $p = $p.Parent }
    if ($p) { $p.Name } else { '<script>' }
})
Assert-Case "one exit with `$ExitCode, in Assert-NoNewAppSession" ($codeExits.Count -eq 1 -and $codeExitIn[0] -eq 'Assert-NoNewAppSession') `
    "found $($codeExits.Count) in [$($codeExitIn -join ', ')]"
$otherExits = @($exitText | Where-Object { $_ -notin @('0', '1', '$elevated.ExitCode', '$script:ExitCode', '$script:BlockedExitCode', '$ExitCode') })
Assert-Case "every other exit is 0, 1, the elevated run's or `$script:ExitCode" ($otherExits.Count -eq 0) ($otherExits -join ', ')
$setCodes = @([regex]::Matches($fixSrc, '(?m)^\s*\$script:ExitCode = (\S+)\s*$') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
Assert-Case "and `$script:ExitCode is only ever 0 or 1" (($setCodes -join ',') -eq '0,1') "set to [$($setCodes -join ',')]"

# Who asks Assert-NoNewAppSession, with which code. 3 before Step 1, when
# nothing that stops the app or the service has run; 4 after it, when Steps
# 1 and 2 have run.
$guardAsks = @($fixAst.FindAll({ param($n)
    $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Assert-NoNewAppSession' }, $true))
$asked = @($guardAsks | ForEach-Object {
    $els = @($_.CommandElements)
    $before = ''; $code = ''
    for ($j = 1; $j -lt $els.Count - 1; $j++) {
        if ($els[$j].Extent.Text -eq '-Before')   { $before = $els[$j + 1].Extent.Text.Trim('"') }
        if ($els[$j].Extent.Text -eq '-ExitCode') { $code = $els[$j + 1].Extent.Text }
    }
    "$before=$code"
} | Sort-Object)
# Step 1 asks twice: before its kill, and in its branch for a list of
# processes that could not be read (6.3.1).
$wantAsked = @('Step 1=$script:BlockedExitCode', 'Step 1=$script:BlockedExitCode', 'Step 4=$script:StoppedPartWayExitCode',
    'the escalation to Deep=$script:StoppedPartWayExitCode', 'the no-progress escalation=$script:StoppedPartWayExitCode') | Sort-Object
Assert-Case "asked before Step 1 with 3, and before Step 4 and both escalations with 4" `
    (($asked -join ' | ') -eq ($wantAsked -join ' | ')) "asked: $($asked -join ' | ')"

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Fix's line that says a run is the repair"

# The monitor ends a repair it keeps pending for a Fix run since its hold only
# when that run's log holds this line, so a run that changed nothing, a
# Diagnostic run, a Cancel in the menu or a No at "Proceed?", does not end
# it. Until 6.3.0 any Fix log newer than the hold ended it. The line is
# logged in one place, in Fix's main body, after every exit there but the
# last and every check that exits 3, before every check that exits 4, and
# not in a -WhatIf run. The if that logs it runs here alone, with Log faked.
$repairLineMatch  = [regex]::Match($fixSrc, '(?m)^\$script:RepairUnderWayLine = ''([^'']+)''\s*$')
$watchRepairMatch = [regex]::Match($watchSrc, '(?m)^\$FixRepairLine\s*=\s*''([^'']+)''\s*$')
Assert-Case "Fix names the line, and the monitor looks for the same one" `
    ($repairLineMatch.Success -and $watchRepairMatch.Success -and $repairLineMatch.Groups[1].Value -ceq $watchRepairMatch.Groups[1].Value) `
    "Fix [$($repairLineMatch.Value)], Watch [$($watchRepairMatch.Value)]"
$repairLogs = @(Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Log' -and
    @($n.CommandElements).Count -ge 2 -and $n.CommandElements[1].Extent.Text -eq '$script:RepairUnderWayLine' })
Assert-Case "it is logged in one place" ($repairLogs.Count -eq 1) "found $($repairLogs.Count)"
if ($repairLogs.Count -eq 1) {
    $repairLog = $repairLogs[0]
    $repairInFunction = $false
    $repairGuard = $null
    $p = $repairLog.Parent
    while ($p) {
        if ($p -is [System.Management.Automation.Language.FunctionDefinitionAst]) { $repairInFunction = $true }
        if ($null -eq $repairGuard -and $p -is [System.Management.Automation.Language.IfStatementAst]) { $repairGuard = $p }
        $p = $p.Parent
    }
    Assert-Case "in Fix's main body, not in a function" (-not $repairInFunction)
    $guardText = if ($repairGuard) { $repairGuard.Clauses[0].Item1.Extent.Text } else { '<none>' }
    Assert-Case "inside an if that leaves out a -WhatIf run, and nothing else" `
        ($null -ne $repairGuard -and @($repairGuard.Clauses).Count -eq 1 -and $null -eq $repairGuard.ElseClause -and
         $guardText -eq '-not $WhatIfPreference') "if [$guardText]"
    $repairAt = $repairLog.Extent.StartOffset
    function Test-InFixFunction {
        param($Ast)
        $q = $Ast.Parent
        while ($q) { if ($q -is [System.Management.Automation.Language.FunctionDefinitionAst]) { return $true }; $q = $q.Parent }
        return $false
    }
    # The last exit ends a run that went on to the repair.
    $bodyExits = @($fixExits | Where-Object { -not (Test-InFixFunction $_) })
    $lastExits = @($bodyExits | Where-Object { $_.Pipeline -and $_.Pipeline.Extent.Text -eq '$script:ExitCode' })
    $earlyExits = @($bodyExits | Where-Object { -not ($_.Pipeline -and $_.Pipeline.Extent.Text -eq '$script:ExitCode') })
    $lateExits = @($earlyExits | Where-Object { $_.Extent.StartOffset -gt $repairAt } | ForEach-Object { "line $($_.Extent.StartLineNumber): $($_.Extent.Text)" })
    Assert-Case "after every exit in the main body but the last" `
        ($earlyExits.Count -ge 9 -and $lateExits.Count -eq 0 -and $lastExits.Count -eq 1 -and $lastExits[0].Extent.StartOffset -gt $repairAt) `
        "$($earlyExits.Count) exits before the last; after the line: $($lateExits -join '; ')"
    $heldAsks = @($guardAsks | Where-Object { $_.Extent.Text -match '-ExitCode \$script:BlockedExitCode' })
    $partAsks = @($guardAsks | Where-Object { $_.Extent.Text -match '-ExitCode \$script:StoppedPartWayExitCode' })
    Assert-Case "after the check right before Step 1's kill, which exits 3, and before the three that exit 4" `
        ($heldAsks.Count -ge 1 -and @($heldAsks | Where-Object { $_.Extent.StartOffset -gt $repairAt }).Count -eq 0 -and
         $partAsks.Count -ge 3 -and @($partAsks | Where-Object { $_.Extent.StartOffset -lt $repairAt }).Count -eq 0) `
        "exits 3: $($heldAsks.Count), exits 4: $($partAsks.Count)"
    $stopLogs = @(Find-FixStatement { param($n)
        $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Log' -and
        @($n.CommandElements).Count -ge 2 -and $n.CommandElements[1].Extent.Text -match '^"(Cancelled by user|Diagnostic complete, no changes were made)"$' })
    Assert-Case "after the Cancel, the No and the Diagnostic run's lines" `
        ($stopLogs.Count -ge 4 -and @($stopLogs | Where-Object { $_.Extent.StartOffset -gt $repairAt }).Count -eq 0) "found $($stopLogs.Count)"
    $unfaked = @(if ($repairGuard) { Get-UnfakedCommand -Asts @($repairGuard) -Faked @('Log') } else { '<none>' })
    Assert-Case "every command the if calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
    if ($unfaked.Count -eq 0) {
        $guardRuns = @(foreach ($w in @($true, $false)) {
            $said = [System.Collections.ArrayList]::new()
            & {
                $WhatIfPreference = $w
                $script:RepairUnderWayLine = $repairLineMatch.Groups[1].Value
                function Log { param([string]$Message, [string]$Colour, [switch]$Indent) $null = $Colour, $Indent; [void]$said.Add($Message) }
                . ([scriptblock]::Create($repairGuard.Extent.Text))
            }
            "$w=$(@($said) -join '/')"
        })
        Assert-Case "a -WhatIf run does not log it, and any other run logs it once" `
            (($guardRuns -join ' | ') -eq "True= | False=$($repairLineMatch.Groups[1].Value)") ($guardRuns -join ' | ')
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Fix's -Quiet gate, run"

$gateIf = Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and
    $n.Clauses[0].Item1.Extent.Text -eq '$Quiet' -and $n.Extent.Text -match 'Get-AutomatedFixBlocker' }
Assert-Case "the gate was found" ($gateIf.Count -eq 1) "found $($gateIf.Count)"
if ($gateIf.Count -eq 1) {
    $unfaked = @(Get-UnfakedCommand -Asts $gateIf -Faked @('Log', 'Save-Log', 'Get-AutomatedFixBlocker'))
    Assert-Case "every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
    if ($unfaked.Count -eq 0) {
        foreach ($gc in @(@{ Blocker = "'a test blocker'"; Exit = 3; Rx = 'LOG: BLOCKED: a test blocker, skipping automated fix' },
                          @{ Blocker = '$null';            Exit = 0; Rx = 'PASSED THE GATE' })) {
            $body = @(
                'Set-StrictMode -Version Latest',
                $fixCodeLine,
                $childFakes,
                "function Get-AutomatedFixBlocker { return $($gc.Blocker) }",
                '$Quiet = $true',
                $gateIf[0].Extent.Text,
                "Write-Output 'PASSED THE GATE'",
                'exit 0'
            ) -join "`r`n"
            $r = Invoke-Child -Body $body -Name 'gate'
            Assert-Case "blocker $($gc.Blocker): exit $($gc.Exit)" ($r.Exit -eq $gc.Exit -and $r.Out -match $gc.Rx) "exit $($r.Exit); $($r.Out.Trim())"
        }
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Fix's lock, run"

# The lock's loop, with both lock names swapped for one only this test uses.
# This process holds it; the child must find it held and end as held back.
$lockLoop = Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.ForEachStatementAst] -and $n.Condition.Extent.Text -match 'ClaudeDesktopFix' }
Assert-Case "the lock was found" ($lockLoop.Count -eq 1) "found $($lockLoop.Count)"
if ($lockLoop.Count -eq 1) {
    $testLock = 'Local\ClaudeFixTest_' + [guid]::NewGuid().ToString('N')
    $lockText = $lockLoop[0].Extent.Text
    $swaps = ([regex]::Matches($lockText, [regex]::Escape($fixLockLiteral))).Count
    $lockText = $lockText.Replace($fixLockLiteral, "@(""$testLock"")")
    $lockSafe = ($swaps -eq 1 -and $lockText -notmatch 'ClaudeDesktopFix')
    Assert-Case "the lock names were swapped for a test name" $lockSafe "swapped $swaps time(s)"
    $lockAst = [System.Management.Automation.Language.Parser]::ParseInput($lockText, [ref]$null, [ref]$null)
    $unfaked = @(Get-UnfakedCommand -Asts @($lockAst) -Faked @('Log', 'Save-Log', 'Wait-ForAnyKey'))
    Assert-Case "every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
    if ($lockSafe -and $unfaked.Count -eq 0) {
        $lockBody = {
            param([bool]$Quiet)
            @(
                'Set-StrictMode -Version Latest',
                $fixCodeLine,
                $childFakes,
                "`$Quiet = `$$Quiet",
                '$fixMutex = $null',
                '$fixMutexHeld = $false',
                $lockText,
                'Write-Output "LOCK TAKEN $fixMutexHeld"',
                'exit 0'
            ) -join "`r`n"
        }
        $body = & $lockBody $true
        $held = [System.Threading.Mutex]::new($true, $testLock)
        try {
            $r = Invoke-Child -Body $body -Name 'lock'
            $rManual = Invoke-Child -Body (& $lockBody $false) -Name 'lock-manual'
        } finally {
            $held.ReleaseMutex()
            $held.Dispose()
        }
        Assert-Case "another run holds the lock: exit 3" `
            ($r.Exit -eq 3 -and $r.Out -match 'LOG: Another Fix instance is already running') "exit $($r.Exit); $($r.Out.Trim())"
        Assert-Case "a -Quiet run does not wait for a key" ($r.Out -notmatch 'WAITED FOR A KEY') $r.Out.Trim()
        # A manual run ends as every other manual run does, so the window does
        # not close before the line can be read: the launcher does not pause
        # after a 3. Until 6.3.0 it closed at once.
        Assert-Case "a manual run: exit 3, after waiting for a key" `
            ($rManual.Exit -eq 3 -and $rManual.Out -match '(?s)Another Fix instance is already running.*WAITED FOR A KEY') "exit $($rManual.Exit); $($rManual.Out.Trim())"
        $r = Invoke-Child -Body $body -Name 'lock'
        Assert-Case "nobody holds it: taken, the run goes on" ($r.Exit -eq 0 -and $r.Out -match 'LOCK TAKEN True') "exit $($r.Exit); $($r.Out.Trim())"
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- -BootPrep: retired"

# The switch still binds, so a boot task an older Prevent left behind does
# not fail on it, and the run does nothing but say so. Until 6.3.0 it
# restarted vmcompute.
$pb = $fixAst.ParamBlock
$bindText = ((@($pb.Attributes | ForEach-Object { $_.Extent.Text }) -join "`r`n") + "`r`n" + $pb.Extent.Text + "`r`n" + '"$BootPrep|$Quiet"')
$bound = '<nothing>'
try { $bound = & ([scriptblock]::Create($bindText)) -BootPrep -Quiet } catch { $bound = "<threw: $($_.Exception.Message)>" }
Assert-Case "-BootPrep -Quiet binds" ($bound -eq 'True|True') "got $bound"

# And it is the first thing the run does: before the elevation, which would
# raise a UAC prompt at logon, and before the environment is saved, old logs
# and backups are pruned and a log is written. Until 6.3.0 all of that ran
# first, for a run that was to do nothing, and an unelevated run passed
# -BootPrep on to an elevated one.
$bootFirst = @($fixAst.EndBlock.Statements)[0]
$bootIsFirst = ($bootFirst -is [System.Management.Automation.Language.IfStatementAst] -and
    @($bootFirst.Clauses).Count -eq 1 -and $bootFirst.Clauses[0].Item1.Extent.Text -eq '$BootPrep' -and
    $null -eq $bootFirst.ElseClause)
Assert-Case "the -BootPrep branch is the first statement after the parameters" $bootIsFirst `
    "first: $(@($bootFirst.Extent.Text -split "\r?\n")[0])"
$bootReads = @($fixAst.FindAll({ param($n)
    $n -is [System.Management.Automation.Language.VariableExpressionAst] -and $n.VariablePath.UserPath -eq 'BootPrep' -and
    $n.Parent -isnot [System.Management.Automation.Language.ParameterAst] }, $true))
Assert-Case "and nothing else reads -BootPrep: no relaunch passes it on" ($bootReads.Count -eq 1) "read $($bootReads.Count) time(s)"
if ($bootIsFirst) {
    $unfaked = @(Get-UnfakedCommand -Asts @($bootFirst) -Faked @('Write-Host'))
    Assert-Case "it calls nothing but Write-Host" ($unfaked.Count -eq 0) ($unfaked -join ', ')
    if ($unfaked.Count -eq 0) {
        # Run alone in a child, with a recorder for each command the run
        # called before 6.3.0, once with -BootPrep and once without.
        $bootRecorders = @('Log', 'Save-Log', 'Get-Service', 'Start-Service', 'Stop-Service', 'Restart-Service', 'Start-Sleep',
            'Invoke-HcsDiag', 'Close-StaleHcsVms', 'Get-Process', 'Stop-Process', 'Start-Process') |
            ForEach-Object { "function $_ { Write-Output 'CALLED $_' }" }
        $bootRuns = [ordered]@{}
        foreach ($on in @('true', 'false')) {
            $body = @(
                'Set-StrictMode -Version Latest',
                'function Write-Host { param([object]$Object, $ForegroundColor) $null = $ForegroundColor; Write-Output "HOST: $Object" }',
                ($bootRecorders -join "`r`n"),
                ('$BootPrep = $' + $on),
                $bootFirst.Extent.Text,
                "Write-Output 'FELL THROUGH'",
                'exit 7'
            ) -join "`r`n"
            $bootRuns[$on] = Invoke-Child -Body $body -Name "bootprep-$on"
        }
        $r = $bootRuns['true']
        $calls = @(($r.Out -split "\r?\n") | Where-Object { $_ -match '^CALLED ' })
        Assert-Case "-BootPrep: exits 0" ($r.Exit -eq 0 -and $r.Out -notmatch 'FELL THROUGH') "exit $($r.Exit); $($r.Out.Trim())"
        Assert-Case "calls nothing: no service, no hcsdiag, no process, no log" ($calls.Count -eq 0) ($calls -join ', ')
        Assert-Case "and says it is retired" `
            ($r.Out -match 'HOST:   \[BootPrep\] -BootPrep is retired since 6\.3\.0 and does nothing; vmcompute and the VMs were not touched') $r.Out.Trim()
        $r = $bootRuns['false']
        Assert-Case "without it: the run goes on, and nothing is said" ($r.Exit -eq 7 -and $r.Out -match 'FELL THROUGH' -and $r.Out -notmatch 'HOST:') `
            "exit $($r.Exit); $($r.Out.Trim())"
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- the monitor: a run Fix held back is not an attempt"

# Invoke-AutoFix, lifted from Watch, with every check it makes faked to say
# "go ahead". $fixScript is a stand-in that only exits with a chosen code,
# run in this process the way the monitor runs Fix. Three held-back runs
# must leave the escalation unarmed and what was detected in place; three
# finished runs must still arm it. Fix's lock names are swapped for a test
# name so a real Fix run cannot change the outcome, and the elevation check
# is taken out so the gates give the same answer in any shell.
#
# Import-WatchState reads watch-state.json from a temp file this test
# writes; Get-Item, Test-Path and Get-Content are the real ones unless a case
# makes the read fail, and ConvertFrom-Json is the real one.
$autoFixDef = Get-AstFunction -Ast $watchAst -Name 'Invoke-AutoFix'
$persistDef = Get-AstFunction -Ast $watchAst -Name 'Test-PersistentFailure'
$importDef  = Get-AstFunction -Ast $watchAst -Name 'Import-WatchState'
$stateTimeDef = Get-AstFunction -Ast $watchAst -Name 'ConvertFrom-WatchStateTime'
$clearDef   = Get-AstFunction -Ast $watchAst -Name 'Clear-HeldBackRepair'
$closeDef   = Get-AstFunction -Ast $watchAst -Name 'Close-HeldBackRepair'
$registerDef = Get-AstFunction -Ast $watchAst -Name 'Register-HeldBackRepair'
# The helpers the 6.3.0 rules live in, lifted when they are there, so on a
# tree without them this section still runs, and each rule it lacks fails as
# a case of its own.
$autoHelperDefs = @(@('Test-FixLockHeld', 'Write-FixRunningSkip', 'Merge-HeldBackTrigger', 'Get-ClaudeCodeSessionHold', 'Test-LineDue',
        'Get-HeldBackRetryWait', 'Get-CheckErrorText', 'Test-CheckErrorDue', 'Write-CheckError', 'Get-ProcessStartUtc',
        'Test-ClaudeNotFoundError', 'Get-ClaudeItemOrNull', 'Get-ClaudeErrorText') |
        ForEach-Object { Get-AstFunction -Ast $watchAst -Name $_ } | Where-Object { $_ })
Assert-Case "Invoke-AutoFix, Test-PersistentFailure and Import-WatchState were found" `
    ($null -ne $autoFixDef -and $null -ne $persistDef -and $null -ne $importDef)
if ($autoFixDef -and $persistDef -and $importDef) {
    $autoText = $autoFixDef.Extent.Text
    # Fix's lock names: in Invoke-AutoFix until 6.3.0, in Test-FixLockHeld
    # since. Swapped once, wherever they are.
    $autoTestLock   = '@("Local\ClaudeFixTest_' + [guid]::NewGuid().ToString('N') + '")'
    $autoLockSwaps  = ([regex]::Matches($autoText, [regex]::Escape($fixLockLiteral))).Count
    foreach ($d in $autoHelperDefs) { $autoLockSwaps += ([regex]::Matches($d.Extent.Text, [regex]::Escape($fixLockLiteral))).Count }
    $autoAdminSwaps = ([regex]::Matches($autoText, [regex]::Escape('if (-not $watchIsAdmin) {'))).Count
    $autoText = $autoText.Replace($fixLockLiteral, $autoTestLock)
    $autoText = $autoText.Replace('if (-not $watchIsAdmin) {', 'if ($false) {')
    $autoHelperTexts = @($autoHelperDefs | ForEach-Object { $_.Extent.Text.Replace($fixLockLiteral, $autoTestLock) })
    $autoSafe = ($autoLockSwaps -eq 1 -and $autoAdminSwaps -eq 1 -and
        @(@($autoText) + $autoHelperTexts | Where-Object { $_ -match '"(Global|Local)\\ClaudeDesktopFix"' }).Count -eq 0)
    Assert-Case "the lock names and the elevation check were swapped" $autoSafe "lock $autoLockSwaps, elevation $autoAdminSwaps"
    $autoFakes = @('Write-WatchLog', 'Save-WatchState', 'Add-Type', 'New-Object', 'Start-Sleep', 'Get-Process',
        'Get-CimInstance', 'Get-ClaudeDesktopProcess', 'Get-ClaudeVmLogAge', 'Test-UserActivity', 'Get-ClaudeProcessByName',
        'Test-ClaudeCodeActive', 'Test-CoworkSessionActive', 'Test-PersistentFailure',
        'ConvertFrom-WatchStateTime', 'Test-Path', 'Get-Content', 'Get-Item', 'ConvertFrom-Json', 'Clear-HeldBackRepair',
        'Close-HeldBackRepair', 'Register-HeldBackRepair', 'Close-StaleHeldBackRepair', 'Get-ClaudeAppCliProcess',
        'Test-ClaudeCodeTranscriptRecent', 'Get-FixRunAfter') + @($autoHelperDefs | ForEach-Object { $_.Name })
    $autoAsts = @($autoFixDef, $persistDef, $importDef) + @($stateTimeDef, $clearDef, $closeDef, $registerDef | Where-Object { $_ }) +
        $autoHelperDefs
    $unfaked = @(Get-UnfakedCommand -Asts $autoAsts -Faked $autoFakes -AllowDynamic @('$fixScript'))
    Assert-Case "every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
    if ($autoSafe -and $unfaked.Count -eq 0) {
        $fakeFix = Join-Path $tempRoot 'fake-fix.ps1'
        $auto = Invoke-Isolated -Name 'Invoke-AutoFix' -Block {
            Set-StrictMode -Version Latest
            function Write-WatchLog { param([string]$Message) [void]$script:WatchLines.Add($Message) }
            function Save-WatchState { $script:StateSaves++ }
            function Add-Type { }
            function New-Object { throw 'no desktop here' }
            function Start-Sleep { param([int]$Seconds, [int]$Milliseconds) $null = $Seconds, $Milliseconds }
            # $script:ReadFails is the read a case makes fail: 'desktop' the
            # app's processes, 'procs' every process named claude, 'vmlog'
            # the VM logs; 'cpu' and 'cpuGone' list one claude process whose
            # CPU time cannot be read again, running or exited, and 'cpuEnd'
            # one whose end cannot be read either. As a Process: a getter
            # that throws reads as nothing, and its get_ method throws
            # (measured 2026-11-23). Get-Process is what a tree before 6.3.1
            # asks: with SilentlyContinue, a list that cannot be read is
            # empty there.
            function New-CpuProcess {
                param([bool]$Gone, [bool]$EndUnread)
                $p = [pscustomobject]@{ Id = 77; TotalProcessorTime = [timespan]::Zero; Gone = $Gone; EndUnread = $EndUnread }
                $p | Add-Member -MemberType ScriptMethod -Name Refresh -Value { throw 'Access is denied' }
                $p | Add-Member -MemberType ScriptProperty -Name HasExited -Value {
                    if ($this.EndUnread) { throw 'Access is denied' }
                    return $this.Gone
                }
                $p | Add-Member -MemberType ScriptMethod -Name get_HasExited -Value {
                    if ($this.EndUnread) { throw 'Access is denied' }
                    return $this.Gone
                }
                return $p
            }
            function Get-Process {
                if ($script:ReadFails -eq 'cpu') { return (New-CpuProcess -Gone $false) }
                if ($script:ReadFails -eq 'cpuGone') { return (New-CpuProcess -Gone $true) }
                if ($script:ReadFails -eq 'cpuEnd') { return (New-CpuProcess -Gone $false -EndUnread $true) }
                return @()
            }
            function Get-ClaudeProcessByName {
                param([string]$Name)
                $null = $Name
                if ($script:ReadFails -eq 'procs') { throw 'Access is denied' }
                if ($script:ReadFails -eq 'cpu') { return , @(New-CpuProcess -Gone $false) }
                if ($script:ReadFails -eq 'cpuGone') { return , @(New-CpuProcess -Gone $true) }
                if ($script:ReadFails -eq 'cpuEnd') { return , @(New-CpuProcess -Gone $false -EndUnread $true) }
                return
            }
            function Get-CimInstance { return @() }
            function Get-ClaudeDesktopProcess {
                param($Environment)
                $null = $Environment
                if ($script:ReadFails -eq 'desktop') { throw 'Access is denied' }
                # 'startTime': the app's one process, whose start time cannot
                # be read, still running; 'startGone': exited since listed.
                if ($script:ReadFails -eq 'startTime' -or $script:ReadFails -eq 'startGone') {
                    $p = [pscustomobject]@{ Id = 88; HasExited = ($script:ReadFails -eq 'startGone') }
                    $p | Add-Member -MemberType ScriptProperty -Name StartTime -Value { throw 'Access is denied' }
                    $p | Add-Member -MemberType ScriptMethod -Name get_StartTime -Value { throw 'Access is denied' }
                    return @($p)
                }
                return @()
            }
            function Get-ClaudeVmLogAge {
                param($LogDir)
                $null = $LogDir
                if ($script:ReadFails -eq 'vmlog') { throw 'The VM log folder could not be read' }
                return $null
            }
            function Test-UserActivity { return [bool]$script:UserActive }
            function Test-ClaudeCodeActive { return [bool]$script:CodeActive }
            function Test-CoworkSessionActive { return $false }
            # The Claude Code sessions the app started, as many as the case
            # sets; 'THROW' is a list that cannot be read.
            function Get-ClaudeAppCliProcess {
                param($Environment)
                $null = $Environment
                if ("$script:AppCli" -eq 'THROW') { throw 'the parents could not be read' }
                $n = [int]$script:AppCli
                if ($n -eq 0) { return @() }
                return @(1..$n | ForEach-Object { [pscustomobject]@{ Id = $_ } })
            }
            # 'THROW' is a transcript folder that cannot be read, with the
            # error the region's function throws for it.
            function Test-ClaudeCodeTranscriptRecent {
                param([int]$Minutes)
                $null = $Minutes
                if ("$script:TranscriptRecent" -eq 'THROW') {
                    throw 'The Claude Code transcript folder C:\fake\.claude\projects could not be read: Access is denied'
                }
                return [bool]$script:TranscriptRecent
            }
            # The state file's reads: the real ones on the temp file, unless
            # $script:StateReadFails makes one of them fail as Windows does,
            # for the state file only. Measured 2026-11-23 with an explicit
            # Deny ACE (S\h-fixwatch\acl-probe.ps1, S the scratchpad):
            # 'item', its folder denies the read: Test-Path answers $false
            # and writes a non-terminating ItemExistsUnauthorizedAccessError,
            # PermissionDenied, UnauthorizedAccessException, "Access is
            # denied", and Get-Item writes the same error; 'content', the
            # file denies Read Data: Get-Item reads it, and Get-Content
            # writes GetContentReaderUnauthorizedAccessError,
            # PermissionDenied, UnauthorizedAccessException, naming the path.
            $script:StateReadFails = ''
            function New-StateDeniedError {
                param([string]$Id, [string]$Message, [string]$Target)
                return [System.Management.Automation.ErrorRecord]::new([System.UnauthorizedAccessException]::new($Message),
                    $Id, [System.Management.Automation.ErrorCategory]::PermissionDenied, $Target)
            }
            function Test-Path {
                [CmdletBinding()]
                param([Parameter(Position = 0)][string]$Path, [string]$LiteralPath)
                $p = if ($LiteralPath) { $LiteralPath } else { $Path }
                if ($script:StateReadFails -eq 'item' -and $p -eq $script:StatePath) {
                    $PSCmdlet.WriteError((New-StateDeniedError 'ItemExistsUnauthorizedAccessError' 'Access is denied' $p))
                    return $false
                }
                return (Microsoft.PowerShell.Management\Test-Path -LiteralPath $p)
            }
            function Get-Item {
                [CmdletBinding()]
                param([Parameter(Position = 0)][string]$Path, [string]$LiteralPath, [switch]$Force)
                $p = if ($LiteralPath) { $LiteralPath } else { $Path }
                if ($script:StateReadFails -eq 'item' -and $p -eq $script:StatePath) {
                    $PSCmdlet.WriteError((New-StateDeniedError 'ItemExistsUnauthorizedAccessError' 'Access is denied' $p))
                    return
                }
                return (Microsoft.PowerShell.Management\Get-Item -LiteralPath $p -Force:$Force -ErrorAction Stop)
            }
            function Get-Content {
                [CmdletBinding()]
                param([Parameter(Position = 0)][string]$Path, [string]$LiteralPath, [switch]$Raw)
                $p = if ($LiteralPath) { $LiteralPath } else { $Path }
                if ($script:StateReadFails -eq 'content' -and $p -eq $script:StatePath) {
                    $PSCmdlet.WriteError((New-StateDeniedError 'GetContentReaderUnauthorizedAccessError' "Access to the path '$p' is denied." $p))
                    return
                }
                return (Microsoft.PowerShell.Management\Get-Content -LiteralPath $p -Raw:$Raw -ErrorAction Stop)
            }
            . ([scriptblock]::Create($persistDef.Extent.Text))
            . ([scriptblock]::Create($importDef.Extent.Text))
            if ($stateTimeDef) { . ([scriptblock]::Create($stateTimeDef.Extent.Text)) }
            if ($clearDef) { . ([scriptblock]::Create($clearDef.Extent.Text)) }
            if ($closeDef) { . ([scriptblock]::Create($closeDef.Extent.Text)) }
            if ($registerDef) { . ([scriptblock]::Create($registerDef.Extent.Text)) }
            foreach ($ht in $autoHelperTexts) { . ([scriptblock]::Create($ht)) }
            # No repair here is ever over: no app start, no Fix log.
            function Close-StaleHeldBackRepair { param([switch]$AgeOnly) $null = $AgeOnly; return $false }
            function Get-FixRunAfter { param([datetime]$After) $null = $After; return $null }
            . ([scriptblock]::Create($autoText))

            if ($watchCodeMatch.Success) { $FixHeldBackExitCode = [int]$watchCodeMatch.Groups[1].Value }
            if ($watchPartMatch.Success) { $FixStoppedPartWayExitCode = [int]$watchPartMatch.Groups[1].Value }
            $Cooldown    = 5
            $ProcessName = 'claude'
            $fixScript   = $fakeFix
            $script:ClaudeEnv            = [pscustomobject]@{ ProcessName = 'claude' }
            $script:ClaudeLogDir         = 'C:\nowhere'
            $script:IsInteractiveSession = $false
            $script:Win32ActivityLoaded  = $false
            $script:FixHistory           = [System.Collections.ArrayList]::new()
            $script:FixCount             = 0
            $script:StateSaves           = 0
            $script:WatchLines           = [System.Collections.ArrayList]::new()
            $script:CodeActive           = $false
            $script:PendingFixReason     = $null
            $script:PendingFixSince      = [datetime]::MinValue
            $script:PendingFixNextCheck  = [datetime]::MinValue
            $script:PendingFixWaitLogged = [datetime]::MinValue
            $script:PendingFixHeldAt     = [datetime]::MinValue
            $script:PendingFixTries      = 0
            $script:PendingFixRetriesHeld = 0
            $script:PendingFixLockLogged  = $false
            $script:FixHeldInARow        = 0
            $script:FixHeldLastTime      = [datetime]::MinValue
            $script:CooldownFromHold     = $false
            $script:RepeatLinesSaid      = @{}
            $script:CheckErrorsSaid      = @{}
            $script:UserActive           = $false
            $script:AppCli               = 0
            $script:TranscriptRecent     = $false
            $script:ReadFails            = ''
            function Set-Detected {
                $script:LogBaselines       = @{ 'detected' = 1 }
                $script:VmLogStaleCount    = 2
                $script:VmLogEverActive    = $true
                $script:ServiceDownCount   = 2
                $script:EventLogHitCount   = 2
                $script:VmcomputeLeakCount = 2
            }
            $o = [ordered]@{}

            [System.IO.File]::WriteAllText($fakeFix, "param([string]`$Mode, [switch]`$Quiet)`r`n`$null = `$Mode, `$Quiet`r`nexit 3`r`n")
            Set-Detected
            $retry = @()
            $pendingAfter = ''
            for ($run = 1; $run -le 3; $run++) {
                $script:LastFixTime = [datetime]::MinValue
                $null = Invoke-AutoFix -Reason 'test trigger'
                $retry += [math]::Round(([datetime]::UtcNow - $script:LastFixTime).TotalMinutes, 1)
                $pendingAfter += "[$($script:PendingFixReason)]"
            }
            $o['heldHistory']  = $script:FixHistory.Count
            $o['heldKept']     = ($script:LogBaselines.ContainsKey('detected') -and $script:VmLogStaleCount -eq 2 -and
                                  $script:ServiceDownCount -eq 2 -and $script:EventLogHitCount -eq 2 -and $script:VmcomputeLeakCount -eq 2)
            $o['heldRetry']    = ($retry -join ',')
            $o['heldLines']    = @($script:WatchLines | Where-Object { $_ -match 'HELD BACK by Fix \(exit 3\)' }).Count
            $o['heldComplete'] = @($script:WatchLines | Where-Object { $_ -match 'COMPLETE' }).Count
            # 'test trigger' is a re-checked trigger, as the service check
            # is: it fires again while the fault lasts, so a hold keeps no
            # repair pending.
            $o['heldText']     = @($script:WatchLines | Where-Object { $_ -match '^>>> AUTO-FIX #\d+ HELD BACK by Fix \(exit 3\): not counted as an attempt; the check fires again while the fault lasts$' }).Count
            # After each hold: the third try ends a kept repair too.
            $o['heldPending']  = $pendingAfter
            # A one-shot trigger, as a log error is, does not fire again, so
            # its repair is kept. -OneShot is asked about rather than passed,
            # so on a tree without it this case still runs and fails alone.
            $script:WatchLines.Clear()
            $script:LastFixTime = [datetime]::MinValue
            if ((Get-Command Invoke-AutoFix).Parameters.ContainsKey('OneShot')) {
                $null = Invoke-AutoFix -Reason 'one-shot trigger' -OneShot
            } else {
                $null = Invoke-AutoFix -Reason 'one-shot trigger'
            }
            $o['oneShotText']    = @($script:WatchLines | Where-Object { $_ -match '^>>> AUTO-FIX #\d+ HELD BACK by Fix \(exit 3\): not counted as an attempt; it is retried when the hold clears, for up to 4 hours$' }).Count
            $o['oneShotPending'] = "$($script:PendingFixReason)|$($script:FixHistory.Count)"
            $script:PendingFixReason = $null
            $script:PendingFixSince  = [datetime]::MinValue
            $script:LastFixTime = [datetime]::MinValue
            $null = Invoke-AutoFix -Reason 'test trigger'
            $o['heldEscalated'] = @($script:WatchLines | Where-Object { $_ -match '^ESCALATION' }).Count

            # A Fix run that takes 3 seconds to hold back: the 2 minutes are
            # counted from when it returned, so 3 minutes of the 5-minute
            # cooldown have gone, not 3 minutes and 3 seconds. The first run
            # held back in a row, so the wait is 2 minutes.
            $script:FixHeldInARow   = 0
            $script:FixHeldLastTime = [datetime]::MinValue
            [System.IO.File]::WriteAllText($fakeFix, "param([string]`$Mode, [switch]`$Quiet)`r`n`$null = `$Mode, `$Quiet`r`n[System.Threading.Thread]::Sleep(3000)`r`nexit 3`r`n")
            $script:LastFixTime = [datetime]::MinValue
            $null = Invoke-AutoFix -Reason 'test trigger'
            $o['heldFromReturn'] = [math]::Round(([datetime]::UtcNow - $script:LastFixTime).TotalSeconds)

            # A run the 30-second warning cancels, because a Claude Code
            # session is active by then: checked again about 2 minutes after
            # the cancel. The warning takes 3 seconds here. Until 6.3.0 the 2
            # minutes were counted from before the warning, so 3 minutes and
            # 3 seconds of the 5-minute cooldown had gone.
            $script:FixHistory.Clear()
            $script:WatchLines.Clear()
            $script:CodeActive  = $true
            $script:LastFixTime = [datetime]::MinValue
            & {
                function Start-Sleep {
                    param([int]$Seconds, [int]$Milliseconds)
                    $null = $Milliseconds
                    if ($Seconds -eq 30) { [System.Threading.Thread]::Sleep(3000) }
                }
                $null = Invoke-AutoFix -Reason 'test trigger'
            }
            $o['cancelFromCancel'] = [math]::Round(([datetime]::UtcNow - $script:LastFixTime).TotalSeconds)
            $o['cancelLines']   = @($script:WatchLines | Where-Object { $_ -match '^>>> CANCELLED: Claude Code session active during 30s grace period$' }).Count
            $o['cancelHistory'] = $script:FixHistory.Count
            $script:CodeActive  = $false

            # Each read before the repair that fails: the repair is deferred
            # or cancelled, and the line says why. Until 6.3.1 each read as
            # nothing, and the repair went ahead. Fix holds back at once here.
            [System.IO.File]::WriteAllText($fakeFix, "param([string]`$Mode, [switch]`$Quiet)`r`n`$null = `$Mode, `$Quiet`r`nexit 3`r`n")
            foreach ($rf in @('desktop', 'procs', 'cpu', 'cpuGone', 'cpuEnd', 'vmlog', 'startTime', 'startGone')) {
                $script:FixHistory.Clear()
                $script:WatchLines.Clear()
                $script:ReadFails = $rf
                $script:LastFixTime = [datetime]::MinValue
                try { $null = Invoke-AutoFix -Reason 'test trigger' } catch { [void]$script:WatchLines.Add("THREW: $($_.Exception.Message)") }
                $o["readFail_$rf"] = "$($script:FixHistory.Count)|" +
                    (@($script:WatchLines | Where-Object { $_ -match '^(>>> (DEFERRED|CANCELLED|AUTO-FIX)|THREW)' }) -join ' / ')
            }
            $script:ReadFails        = ''
            $script:FixHeldInARow    = 0
            $script:FixHeldLastTime  = [datetime]::MinValue
            $script:CooldownFromHold = $false

            # Stopped part-way three times: each one an attempt.
            [System.IO.File]::WriteAllText($fakeFix, "param([string]`$Mode, [switch]`$Quiet)`r`n`$null = `$Mode, `$Quiet`r`nexit 4`r`n")
            $script:FixHistory.Clear()
            $script:WatchLines.Clear()
            for ($run = 1; $run -le 3; $run++) {
                Set-Detected
                $script:LastFixTime = [datetime]::MinValue
                $null = Invoke-AutoFix -Reason 'test trigger'
            }
            $o['partHistory']  = $script:FixHistory.Count
            $o['partReset']    = (-not $script:LogBaselines.ContainsKey('detected') -and $script:VmLogStaleCount -eq 0 -and
                                  $script:ServiceDownCount -eq 0 -and $script:EventLogHitCount -eq 0 -and $script:VmcomputeLeakCount -eq 0)
            $o['partLines']    = @($script:WatchLines | Where-Object { $_ -match 'STOPPED PART-WAY by Fix \(exit 4\): the repair ran part of the way, then Fix stopped before going further: a Claude Code session was found, or could not be ruled out \(see the Fix log\); counted as an attempt$' }).Count
            $o['partComplete'] = @($script:WatchLines | Where-Object { $_ -match 'COMPLETE|HELD BACK' }).Count
            # The full cooldown, from when the fix began, not the 2-minute
            # retry of a run held back.
            $null = Invoke-AutoFix -Reason 'test trigger'
            $o['partCooldown'] = @($script:WatchLines | Where-Object { $_ -match '^COOLDOWN: Skipping fix \(test trigger\), a fix is allowed again in 5 min$' }).Count
            $script:LastFixTime = [datetime]::MinValue
            $null = Invoke-AutoFix -Reason 'test trigger'
            $o['partEscalated'] = @($script:WatchLines | Where-Object { $_ -match '^ESCALATION' }).Count

            # The cooldown, with the local clock an hour ahead of UTC, as it
            # is from the moment summer time starts: a fix 2 minutes ago is
            # 2 minutes ago.
            $script:FixHistory.Clear()
            $script:WatchLines.Clear()
            $script:RepeatLinesSaid = @{}
            $script:LastFixTime = [datetime]::UtcNow.AddMinutes(-2)
            & {
                function Get-Date { return [datetime]::UtcNow.AddHours(1) }
                $null = Invoke-AutoFix -Reason 'test trigger'
            }
            $o['dstCooldown'] = ($script:WatchLines -join ' | ')

            # Three fixes in the last 30 minutes escalate whatever the local
            # clock says, and three older ones do not.
            $script:FixHistory.Clear()
            foreach ($ago in 20, 10, 0) { $null = $script:FixHistory.Add([datetime]::UtcNow.AddMinutes(-$ago)) }
            $o['dstRecent'] = & { function Get-Date { return [datetime]::UtcNow.AddHours(1) }; Test-PersistentFailure }
            $script:FixHistory.Clear()
            foreach ($ago in 40, 35, 31) { $null = $script:FixHistory.Add([datetime]::UtcNow.AddMinutes(-$ago)) }
            $o['dstOld'] = & { function Get-Date { return [datetime]::UtcNow.AddHours(-1) }; Test-PersistentFailure }

            # watch-state.json as an older version wrote it: local times with
            # their offset. Read back as UTC, the same instants.
            $script:StatePath = Join-Path $tempRoot 'watch-state.json'
            function Read-State {
                param([string]$Json)
                [System.IO.File]::WriteAllText($script:StatePath, $Json)
                $script:LastFixTime = [datetime]::MinValue
                $script:FixHistory.Clear()
                Import-WatchState
                $inv = [System.Globalization.CultureInfo]::InvariantCulture
                return "$($script:LastFixTime.Kind) $($script:LastFixTime.ToString('yyyy-MM-ddTHH:mm:ss', $inv)) | " +
                    (@($script:FixHistory | ForEach-Object { "$($_.Kind) $($_.ToString('yyyy-MM-ddTHH:mm:ss', $inv))" }) -join ', ')
            }
            $o['stateOld']  = Read-State '{ "LastFixTime": "2026-07-01T10:00:00.0000000+01:00", "FixCount": 3, "FixHistory": [ "2026-07-01T09:50:00.0000000+01:00" ] }'
            $o['stateNew']  = Read-State '{ "LastFixTime": "2026-11-22T10:00:00.0000000Z", "FixCount": 3, "FixHistory": [ "2026-11-22T09:50:00.0000000Z" ] }'
            $o['stateNone'] = Read-State '{ "LastFixTime": "0001-01-01T00:00:00.0000000", "FixCount": 0, "FixHistory": [ ] }'
            # A file that is not JSON: what it held starts empty, a cooldown
            # starts now, and that is said, once in the hour for the same
            # error. Until 6.3.0 it was not said, and until 6.3.1 it started
            # with no cooldown, as on a first start.
            $script:WatchLines.Clear()
            $script:CheckErrorsSaid = @{}
            $null = Read-State '{ "LastFixTime": '
            $null = Read-State '{ "LastFixTime": '
            $o['stateBroken'] = "$(@($script:WatchLines | Where-Object { $_ -match '^STATE: The state saved in .+watch-state\.json could not be restored in full \(.+\); what was not restored starts empty, and a cooldown starts now, as after a repair: no automatic repair runs for 5 min; the same error is logged at most once an hour$' }).Count)|$($script:WatchLines.Count)"
            # A held-back repair whose time cannot be read: not restored, and
            # that is said, once in the hour for the same error. Until 6.3.0
            # it was dropped without a word.
            $script:WatchLines.Clear()
            $script:CheckErrorsSaid = @{}
            $script:PendingFixReason = $null
            $script:PendingFixSince = [datetime]::MinValue
            $pendingJson = '{ "LastFixTime": "0001-01-01T00:00:00.0000000", "FixCount": 0, "FixHistory": [ ], ' +
                '"PendingFixReason": "log trigger", "PendingFixSince": "not a time" }'
            $null = Read-State $pendingJson
            $null = Read-State $pendingJson
            $o['statePending'] = "$(@($script:WatchLines | Where-Object { $_ -match '^STATE: The held-back repair saved in .+watch-state\.json could not be restored \(.+\), so it is not tried again; the same error is logged at most once an hour$' }).Count)|$($script:WatchLines.Count)|$([string]$script:PendingFixReason)|$($script:LastFixTime -eq [datetime]::MinValue)"
            # And the cooldown from such a file: a fix 2 minutes ago, written
            # in local time an hour ahead of UTC, is still 2 minutes ago.
            $script:WatchLines.Clear()
            $script:RepeatLinesSaid = @{}
            $twoAgo = [datetimeoffset]::UtcNow.AddMinutes(-2).ToOffset([timespan]::FromHours(1)).ToString('o', [System.Globalization.CultureInfo]::InvariantCulture)
            $null = Read-State ('{ "LastFixTime": "' + $twoAgo + '", "FixCount": 1, "FixHistory": [ "' + $twoAgo + '" ] }')
            $null = Invoke-AutoFix -Reason 'test trigger'
            $o['stateCooldown'] = ($script:WatchLines -join ' | ')

            # A file that is there but could not be tested, read or parsed:
            # what was not restored starts empty, a cooldown starts now, as
            # after a repair, and whether a hold left it is taken as not;
            # said once in the hour. A wait already longer is kept, and no
            # repair is added to the history. A file that is not there is a
            # first start: no cooldown and no line. Since 6.3.1: until then a
            # file Test-Path could not test, and an empty one, read as no
            # state at all without a word, and a file that could not be read
            # or parsed started with no cooldown. Each is read twice: the
            # second, in the same hour, says nothing. The history is the
            # first read's.
            function Read-StateClosed {
                param([string]$Json, [string]$Fails = '', [switch]$NoFile)
                $script:WatchLines.Clear()
                $script:CheckErrorsSaid = @{}
                $script:RepeatLinesSaid = @{}
                if ($NoFile) { [System.IO.File]::Delete($script:StatePath) } else { [System.IO.File]::WriteAllText($script:StatePath, $Json) }
                $script:LastFixTime = [datetime]::MinValue
                $script:FixHistory.Clear()
                $script:CooldownFromHold = $true
                $before = [datetime]::UtcNow
                $script:StateReadFails = $Fails
                try {
                    Import-WatchState
                    $history = $script:FixHistory.Count
                    $script:FixHistory.Clear()
                    Import-WatchState
                } finally { $script:StateReadFails = '' }
                $after = [datetime]::UtcNow
                $last = $script:LastFixTime
                $when = if ($last -eq [datetime]::MinValue) { 'none' }
                    elseif ($last -ge $before -and $last -le $after) { 'now' }
                    else { '{0} min from now' -f [math]::Round(($last - $after).TotalMinutes).ToString([System.Globalization.CultureInfo]::InvariantCulture) }
                return "$when|$($script:CooldownFromHold)|$history|$($script:WatchLines.Count)|$($script:WatchLines -join ' / ')"
            }
            $closedGood = '{ "LastFixTime": "0001-01-01T00:00:00.0000000", "FixCount": 0, "FixHistory": [ ] }'
            $inv = [System.Globalization.CultureInfo]::InvariantCulture
            $o['closedNoFile']  = Read-StateClosed -NoFile
            $o['closedNoFileUnread'] = [string]$script:StateUnread
            $o['closedItem']    = Read-StateClosed -Json $closedGood -Fails 'item'
            $o['closedItemUnread'] = [string]$script:StateUnread
            # And no automatic repair runs in it.
            $script:WatchLines.Clear()
            $script:RepeatLinesSaid = @{}
            $null = Invoke-AutoFix -Reason 'test trigger'
            $o['closedItemFix'] = ($script:WatchLines -join ' | ')
            $o['closedContent'] = Read-StateClosed -Json $closedGood -Fails 'content'
            $o['closedNotJson'] = Read-StateClosed -Json '{ "LastFixTime": '
            $o['closedEmpty']   = Read-StateClosed -Json ''
            $o['closedNull']    = Read-StateClosed -Json 'null'
            $o['closedHistory'] = Read-StateClosed -Json ('{ "LastFixTime": "' + [datetime]::UtcNow.AddMinutes(-10).ToString('o', $inv) +
                '", "FixCount": 1, "FixHistory": [ "not a time" ] }')
            $o['closedLonger']  = Read-StateClosed -Json ('{ "LastFixTime": "' + [datetime]::UtcNow.AddMinutes(25).ToString('o', $inv) +
                '", "FixCount": 1, "FixHistory": [ "not a time" ] }')
            $o['closedGood']    = Read-StateClosed -Json ('{ "LastFixTime": "2026-11-22T10:00:00.0000000Z", "FixCount": 1, ' +
                '"FixHistory": [ "2026-11-22T09:50:00.0000000Z" ], "CooldownFromHold": true }')
            $o['closedGoodUnread'] = [string]$script:StateUnread

            [System.IO.File]::WriteAllText($fakeFix, "param([string]`$Mode, [switch]`$Quiet)`r`n`$null = `$Mode, `$Quiet`r`nexit 0`r`n")
            $script:FixHistory.Clear()
            $script:WatchLines.Clear()
            for ($run = 1; $run -le 3; $run++) {
                Set-Detected
                $script:LastFixTime = [datetime]::MinValue
                $null = Invoke-AutoFix -Reason 'test trigger'
            }
            $o['doneHistory']  = $script:FixHistory.Count
            $o['doneReset']    = (-not $script:LogBaselines.ContainsKey('detected') -and $script:VmLogStaleCount -eq 0)
            $o['doneComplete'] = @($script:WatchLines | Where-Object { $_ -match 'COMPLETE \(exit 0\)' }).Count
            $script:LastFixTime = [datetime]::MinValue
            $null = Invoke-AutoFix -Reason 'test trigger'
            $o['doneEscalated'] = @($script:WatchLines | Where-Object { $_ -match '^ESCALATION' }).Count
            $o['lines'] = ($script:WatchLines -join ' | ')

            # What Fix's -Quiet gate holds a run back for, asked before the
            # 30-second warning: a Claude Code session the app started, a
            # transcript written in the last 35 minutes, and a list that
            # cannot be read. BLOCKED, and Fix does not run. Until 6.3.0 the
            # monitor asked none of them when no process named claude ran,
            # so it showed the warning, and Fix held the run back.
            foreach ($cs in @(@{ Name = 'codeCli'; Cli = 1; Transcript = $false }, @{ Name = 'codeTranscript'; Cli = 0; Transcript = $true },
                              @{ Name = 'codeCliThrow'; Cli = 'THROW'; Transcript = $false },
                              @{ Name = 'codeTranscriptThrow'; Cli = 0; Transcript = 'THROW' })) {
                $script:FixHistory.Clear()
                $script:WatchLines.Clear()
                $script:RepeatLinesSaid  = @{}
                $script:AppCli           = $cs.Cli
                $script:TranscriptRecent = $cs.Transcript
                $script:LastFixTime      = [datetime]::MinValue
                $null = Invoke-AutoFix -Reason 'test trigger'
                $o[$cs.Name] = "$(@($script:WatchLines | Where-Object { $_ -match '^>>> BLOCKED: test trigger, ' }) -join ' / ')|" +
                    "$(@($script:WatchLines | Where-Object { $_ -match 'PRE-FIX WARNING|COMPLETE|HELD BACK' }).Count)|$($script:FixHistory.Count)"
            }
            # A one-shot trigger it holds is kept, not as a try, and tried
            # again once that is over.
            $script:WatchLines.Clear()
            $script:RepeatLinesSaid  = @{}
            $script:AppCli           = 1
            $script:TranscriptRecent = $false
            $script:LastFixTime      = [datetime]::MinValue
            $script:PendingFixReason = $null
            $script:PendingFixSince  = [datetime]::MinValue
            if ((Get-Command Invoke-AutoFix).Parameters.ContainsKey('OneShot')) {
                $null = Invoke-AutoFix -Reason 'one-shot trigger' -OneShot
            } else {
                $null = Invoke-AutoFix -Reason 'one-shot trigger'
            }
            $o['codeOneShot'] = "$($script:PendingFixReason)|$($script:PendingFixTries)|" +
                "$(@($script:WatchLines | Where-Object { $_ -match '^    It is tried again once that is over, for up to 4 hours, or run Fix-ClaudeDesktop\.ps1 manually when ready$' }).Count)"
            $script:AppCli           = 0
            $script:PendingFixReason = $null
            $script:PendingFixSince  = [datetime]::MinValue

            # BLOCKED for the same reason: its lines at most once every 10
            # minutes. Another reason has lines of its own, and 10 minutes on
            # they come again.
            $script:WatchLines.Clear()
            $script:RepeatLinesSaid = @{}
            $script:UserActive = $true
            foreach ($r in @('test trigger', 'test trigger', 'other trigger')) {
                $script:LastFixTime = [datetime]::MinValue
                $null = Invoke-AutoFix -Reason $r
            }
            if ($script:RepeatLinesSaid.ContainsKey('BLOCKED|test trigger')) {
                $script:RepeatLinesSaid['BLOCKED|test trigger'] = [datetime]::UtcNow.AddMinutes(-10.5)
            }
            $script:LastFixTime = [datetime]::MinValue
            $null = Invoke-AutoFix -Reason 'test trigger'
            $o['blockedOnce'] = "$(@($script:WatchLines | Where-Object { $_ -match '^>>> BLOCKED: test trigger, user is active' }).Count)|" +
                "$(@($script:WatchLines | Where-Object { $_ -match '^>>> BLOCKED: other trigger, user is active' }).Count)|" +
                "$(@($script:WatchLines | Where-Object { $_ -match '^    Run Fix-ClaudeDesktop\.ps1 manually when ready$' }).Count)"
            $script:UserActive = $false
            # COOLDOWN the same: once in 10 minutes for each reason.
            $script:WatchLines.Clear()
            $script:RepeatLinesSaid = @{}
            $script:LastFixTime = [datetime]::UtcNow.AddMinutes(-1)
            foreach ($r in @('test trigger', 'test trigger', 'test trigger', 'other trigger')) { $null = Invoke-AutoFix -Reason $r }
            $o['cooldownOnce'] = "$(@($script:WatchLines | Where-Object { $_ -match '^COOLDOWN: Skipping fix \(test trigger\)' }).Count)|" +
                "$(@($script:WatchLines | Where-Object { $_ -match '^COOLDOWN: Skipping fix \(other trigger\)' }).Count)"
            $o
        }
        if ($auto) {
            Assert-Case "held back three times: nothing recorded toward the escalation" ($auto['heldHistory'] -eq 0) "FixHistory $($auto['heldHistory'])"
            Assert-Case "what was detected is kept" ($auto['heldKept'] -eq $true)
            Assert-Case "logged as held back, not as complete" ($auto['heldLines'] -eq 3 -and $auto['heldComplete'] -eq 0) `
                "held $($auto['heldLines']), complete $($auto['heldComplete'])"
            # Until 6.3.0 each hold left 2 minutes: '3,3,3'.
            Assert-Case "and each run held back in a row doubles the wait before the next repair: 2, 4, then 8 minutes" ($auto['heldRetry'] -eq '3,1,-3') `
                "minutes since the last fix, with a 5-minute cooldown: $($auto['heldRetry'])"
            Assert-Case "a re-checked trigger: the line says the check fires again while the fault lasts" ($auto['heldText'] -eq 3) "held lines in that wording: $($auto['heldText'])"
            Assert-Case "and its repair is not kept pending, after any of the three" ($auto['heldPending'] -eq '[][][]') "pending after each: $($auto['heldPending'])"
            Assert-Case "a one-shot trigger: the line says it is retried when the hold clears, for up to 4 hours" ($auto['oneShotText'] -eq 1) `
                "held lines in that wording: $($auto['oneShotText'])"
            Assert-Case "and its repair is pending, under its own reason, and is not an attempt" ($auto['oneShotPending'] -eq 'one-shot trigger|0') `
                "pending|history: [$($auto['oneShotPending'])]"
            Assert-Case "a fourth trigger is not an escalation" ($auto['heldEscalated'] -eq 0) "escalated $($auto['heldEscalated'])"
            Assert-Case "the 2 minutes run from when Fix returned" ($auto['heldFromReturn'] -ge 179 -and $auto['heldFromReturn'] -le 181) `
                "seconds of the cooldown gone: $($auto['heldFromReturn']), not 180"
            Assert-Case "a run the warning cancels is logged as cancelled, and is not an attempt" `
                ($auto['cancelLines'] -eq 1 -and $auto['cancelHistory'] -eq 0) "cancelled $($auto['cancelLines']), FixHistory $($auto['cancelHistory'])"
            Assert-Case "and is checked again 2 minutes after the cancel, not after the warning began" `
                ($auto['cancelFromCancel'] -ge 179 -and $auto['cancelFromCancel'] -le 181) `
                "seconds of the cooldown gone: $($auto['cancelFromCancel']), not 180"
            Assert-Case "the app's processes that cannot be listed at the startup check: deferred, saying why, and Fix does not run" `
                ($auto['readFail_desktop'] -eq ("0|>>> DEFERRED: test trigger, Claude's processes could not be listed (Access is denied), " +
                    'so whether it has just started is not known')) "got $($auto['readFail_desktop'])"
            Assert-Case "the processes named claude that cannot be listed after the warning: cancelled, saying why" `
                ($auto['readFail_procs'] -eq "0|>>> CANCELLED: Claude's processes could not be listed (Access is denied) during 30s grace period") `
                "got $($auto['readFail_procs'])"
            Assert-Case "a claude process whose CPU time cannot be read: cancelled, naming it" `
                ($auto['readFail_cpu'] -match '^0\|>>> CANCELLED: whether claude process 77 is using CPU could not be read \(.*Access is denied.*\) during 30s grace period$') `
                "got $($auto['readFail_cpu'])"
            Assert-Case "an app process whose start time cannot be read, still running: deferred, saying why, and Fix does not run" `
                ($auto['readFail_startTime'] -match ("^0\|>>> DEFERRED: test trigger, the start time of one of Claude's processes could not be read " +
                    '\(.*Access is denied.*\), so whether it has just started is not known$')) "got $($auto['readFail_startTime'])"
            Assert-Case "an app process that has exited since it was listed: passed over, and the repair goes ahead" `
                ($auto['readFail_startGone'] -notmatch 'DEFERRED|CANCELLED|THREW' -and $auto['readFail_startGone'] -match 'AUTO-FIX #\d+ HELD BACK by Fix \(exit 3\)') `
                "got $($auto['readFail_startGone'])"
            # Until 6.3.1 the end was read through the getter, which reads as
            # nothing when it throws, so an end that could not be read was
            # not said.
            Assert-Case "one whose end cannot be read either: cancelled, naming both" `
                ($auto['readFail_cpuEnd'] -match ('^0\|>>> CANCELLED: whether claude process 77 is using CPU could not be read \(.*Access is denied.*\), ' +
                    'nor whether it has ended \(.*Access is denied.*\) during 30s grace period$')) `
                "got $($auto['readFail_cpuEnd'])"
            Assert-Case "one that has exited: not busy, and the repair goes ahead" `
                ($auto['readFail_cpuGone'] -notmatch 'CANCELLED|THREW' -and $auto['readFail_cpuGone'] -match 'AUTO-FIX #\d+ HELD BACK by Fix \(exit 3\)') `
                "got $($auto['readFail_cpuGone'])"
            Assert-Case "VM logs that cannot be read after the warning: cancelled, saying why" `
                ($auto['readFail_vmlog'] -eq '0|>>> CANCELLED: the VM logs could not be read (The VM log folder could not be read) during 30s grace period') `
                "got $($auto['readFail_vmlog'])"
            Assert-Case "stopped part-way three times: each recorded as an attempt" ($auto['partHistory'] -eq 3) "FixHistory $($auto['partHistory'])"
            Assert-Case "and what was detected is reset" ($auto['partReset'] -eq $true)
            Assert-Case "logged as stopped part-way, not as complete or held back" ($auto['partLines'] -eq 3 -and $auto['partComplete'] -eq 0) `
                "stopped $($auto['partLines']), complete or held $($auto['partComplete'])"
            Assert-Case "the normal cooldown follows, from when the fix began" ($auto['partCooldown'] -eq 1) "cooldown lines $($auto['partCooldown'])"
            Assert-Case "and the fourth trigger escalates" ($auto['partEscalated'] -eq 1) "escalated $($auto['partEscalated'])"
            Assert-Case "the local clock an hour ahead: a fix 2 min ago is still in its cooldown" `
                ($auto['dstCooldown'] -match '^COOLDOWN: Skipping fix \(test trigger\), a fix is allowed again in (3|2\.9) min$') $auto['dstCooldown']
            Assert-Case "three fixes in 30 min, the local clock an hour ahead: an escalation" ($auto['dstRecent'] -eq $true) "got $($auto['dstRecent'])"
            Assert-Case "three fixes 31 to 40 min ago, the local clock an hour behind: none" ($auto['dstOld'] -eq $false) "got $($auto['dstOld'])"
            Assert-Case "an older watch-state.json, local with an offset: read as the same instants in UTC" `
                ($auto['stateOld'] -eq 'Utc 2026-07-01T09:00:00 | Utc 2026-07-01T08:50:00') "got $($auto['stateOld'])"
            Assert-Case "a 6.3.0 watch-state.json, in UTC: read as written" `
                ($auto['stateNew'] -eq 'Utc 2026-11-22T10:00:00 | Utc 2026-11-22T09:50:00') "got $($auto['stateNew'])"
            Assert-Case "no fix yet: MinValue" ($auto['stateNone'] -match '^\w+ 0001-01-01T00:00:00 \| $') "got $($auto['stateNone'])"
            Assert-Case "a file that is not JSON: a cooldown from now, said, naming the file and the error, once in the hour" ($auto['stateBroken'] -eq '1|1') `
                "that line|lines: $($auto['stateBroken'])"
            Assert-Case "a held-back repair whose time cannot be read: not restored, and said, naming the file and the error, once in the hour" `
                ($auto['statePending'] -match '^1\|1\|\|') "that line|lines|pending: $($auto['statePending'])"
            Assert-Case "and the rest is restored as it was, with no cooldown from now" ($auto['statePending'] -match '\|True$') `
                "that line|lines|pending|no fix yet: $($auto['statePending'])"
            $closedLine = '^STATE: The state saved in .+watch-state\.json could not be restored in full \({0}\); what was not restored starts empty, ' +
                'and a cooldown starts now, as after a repair: no automatic repair runs for {1} min; the same error is logged at most once an hour$'
            $closedRx = { param($e, $m) '^now\|False\|0\|1\|' + ($closedLine -f $e, $m).Substring(1) }
            Assert-Case "no state file: a first start, with no cooldown and no line" ($auto['closedNoFile'] -eq 'none|True|0|0|') `
                "when|from a hold|history|lines|lines: $($auto['closedNoFile'])"
            Assert-Case "a state file whose folder denies the read: a cooldown from now, not from a hold, no repair made up, said once in the hour" `
                ($auto['closedItem'] -match (& $closedRx 'Access is denied' '5')) "when|from a hold|history|lines|lines: $($auto['closedItem'])"
            Assert-Case "and no automatic repair runs in it" `
                ($auto['closedItemFix'] -match '^COOLDOWN: Skipping fix \(test trigger\), a fix is allowed again in 5 min$') $auto['closedItemFix']
            Assert-Case "a state file whose content cannot be read: the same, naming the error" `
                ($auto['closedContent'] -match (& $closedRx "Access to the path '.+watch-state\.json' is denied\." '5')) "got: $($auto['closedContent'])"
            Assert-Case "a state file that is not JSON: the same" `
                ($auto['closedNotJson'] -match (& $closedRx '.+' '5')) "got: $($auto['closedNotJson'])"
            Assert-Case "an empty state file: the same, saying it is empty" `
                ($auto['closedEmpty'] -match (& $closedRx 'the file is empty' '5')) "got: $($auto['closedEmpty'])"
            Assert-Case "a state file that holds no object: the same" `
                ($auto['closedNull'] -match (& $closedRx 'the file does not hold a saved state' '5')) "got: $($auto['closedNull'])"
            Assert-Case "a fix history that cannot be read: a cooldown from now, not from the fix 10 minutes ago it names" `
                ($auto['closedHistory'] -match (& $closedRx '.+' '5')) "got: $($auto['closedHistory'])"
            Assert-Case "and a wait already longer than the cooldown is kept, and said" `
                ($auto['closedLonger'] -match ('^25 min from now\|False\|0\|1\|' + ($closedLine -f '.+', '30').Substring(1))) "got: $($auto['closedLonger'])"
            Assert-Case "a good state file: restored as it was, with no line" `
                ($auto['closedGood'] -match '^-\d+ min from now\|True\|1\|0\|$') `
                "when|from a hold|history|lines|lines: $($auto['closedGood'])"
            # Kept for the stealth relaunch (Invoke-StealthRelaunch), since
            # 6.3.1: the markers already acted on are in that state.
            $unread = "$($auto['closedNoFileUnread'])|$($auto['closedItemUnread'])|$($auto['closedGoodUnread'])"
            Assert-Case "a state that could not be restored is kept as such for the stealth relaunch; no file and a good one are not" `
                ($unread -eq 'False|True|False') "no file|could not be read|good: $unread"
            Assert-Case "an older file's fix 2 min ago is still in its cooldown" `
                ($auto['stateCooldown'] -match '^COOLDOWN: Skipping fix \(test trigger\), a fix is allowed again in (3|2\.9) min$') $auto['stateCooldown']
            Assert-Case "three finished runs are recorded" ($auto['doneHistory'] -eq 3 -and $auto['doneComplete'] -eq 3) `
                "FixHistory $($auto['doneHistory']), complete $($auto['doneComplete'])"
            Assert-Case "and reset what was detected" ($auto['doneReset'] -eq $true)
            Assert-Case "and the fourth trigger escalates" ($auto['doneEscalated'] -eq 1) "escalated $($auto['doneEscalated']); $($auto['lines'])"
            Assert-Case "a Claude Code session the app started: BLOCKED before the warning, and Fix does not run" `
                ($auto['codeCli'] -eq '>>> BLOCKED: test trigger, 1 Claude Code session(s) the app started are running|0|0') "BLOCKED|warnings and runs|history: $($auto['codeCli'])"
            Assert-Case "a Claude Code transcript in the last 35 minutes: the same" `
                ($auto['codeTranscript'] -eq '>>> BLOCKED: test trigger, a Claude Code transcript was written in the last 35 minutes|0|0') "BLOCKED|warnings and runs|history: $($auto['codeTranscript'])"
            Assert-Case "the app's Claude Code sessions that cannot be listed: the same" `
                ($auto['codeCliThrow'] -eq ('>>> BLOCKED: test trigger, the Claude Code sessions the app started could not be listed ' +
                    '(the parents could not be read)|0|0')) "BLOCKED|warnings and runs|history: $($auto['codeCliThrow'])"
            # Until 6.3.1 neither line said why, and the transcript line had
            # no case.
            Assert-Case "Claude Code transcripts that cannot be read: the same, saying why" `
                ($auto['codeTranscriptThrow'] -eq ('>>> BLOCKED: test trigger, whether a Claude Code transcript was written in the last 35 ' +
                    'minutes could not be read (The Claude Code transcript folder C:\fake\.claude\projects could not be read: ' +
                    'Access is denied)|0|0')) "BLOCKED|warnings and runs|history: $($auto['codeTranscriptThrow'])"
            Assert-Case "a one-shot trigger held by one: kept, not a try, and tried again once it is over" ($auto['codeOneShot'] -eq 'one-shot trigger|0|1') `
                "pending|tries|retried line: $($auto['codeOneShot'])"
            Assert-Case "BLOCKED for the same reason: said once in 10 minutes, another reason has its own" ($auto['blockedOnce'] -eq '2|1|3') `
                "test trigger|other trigger|run-by-hand lines: $($auto['blockedOnce'])"
            Assert-Case "COOLDOWN for the same reason: said once in 10 minutes, another reason has its own" ($auto['cooldownOnce'] -eq '1|1') `
                "test trigger|other trigger: $($auto['cooldownOnce'])"
        }
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- the monitor's clock check"

# w32tm is faked: Invoke-W32tm hands back the offsets queued here, 'TIMEOUT'
# for no answer, and the resync's exit code. Start-Job and its kin and a
# w32tm function stand in for what the check called until 6.3.0, so the
# older code is driven by the same cases and never reaches the real w32tm.
# Get-CheckErrorText, Test-CheckErrorDue and Write-CheckError write and
# throttle the line a check that fails writes, and the line for a
# measurement that does not answer; .ci\Test-PreventSteps.ps1 tests both.
$clockNames = @('Test-TimeSyncHealth', 'Measure-ClockOffset', 'Get-TimeSourceName', 'Get-CheckErrorText', 'Test-CheckErrorDue',
    'Write-CheckError')
$clockDefs  = @($clockNames | ForEach-Object { Get-AstFunction -Ast $watchAst -Name $_ } | Where-Object { $_ })
# Get-ClaudeServiceOrNull: the shared region's service read, which the
# check calls from 6.3.1; until then it called Get-Service.
$clockFakes = @('Write-WatchLog', 'Get-Service', 'Get-ClaudeServiceOrNull', 'Start-Service', 'Invoke-W32tm', 'Get-W32tmOffsetSeconds',
    'Start-Job', 'Wait-Job', 'Receive-Job', 'Stop-Job', 'Remove-Job', 'w32tm') + $clockNames
Assert-Case "Test-TimeSyncHealth was found" (@($clockDefs | Where-Object { $_.Name -eq 'Test-TimeSyncHealth' }).Count -eq 1)
$unfaked = @(Get-UnfakedCommand -Asts $clockDefs -Faked $clockFakes)
Assert-Case "every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($clockDefs.Count -gt 0 -and $unfaked.Count -eq 0) {
    $clock = Invoke-Isolated -Name 'the clock check' -Block {
        Set-StrictMode -Version Latest
        $script:Strip   = [System.Collections.Queue]::new()
        $script:Resync  = 0
        $script:Pending = $null
        $script:Calls   = [System.Collections.ArrayList]::new()
        $script:WatchLines = [System.Collections.ArrayList]::new()
        function Format-Offset { param($v) return ([double]$v).ToString('+00.0000000;-00.0000000', [Globalization.CultureInfo]::InvariantCulture) }
        function Write-WatchLog { param([string]$Message) [void]$script:WatchLines.Add($Message) }
        function Get-Service { return [pscustomobject]@{ Status = 'Running' } }
        function Get-ClaudeServiceOrNull { param([string]$Name) $null = $Name; return [pscustomobject]@{ Status = 'Running' } }
        function Start-Service { [void]$script:Calls.Add('start-service') }
        function Invoke-W32tm {
            param([string[]]$Arguments, [int]$TimeoutSeconds = 15)
            $null = $TimeoutSeconds
            switch ($Arguments[0]) {
                '/stripchart' {
                    [void]$script:Calls.Add('stripchart')
                    $v = $script:Strip.Dequeue()
                    if ("$v" -eq 'TIMEOUT') { return $null }
                    return [pscustomobject]@{ Text = "Tracking time.windows.com.`r`n20:22:22, $(Format-Offset $v)s`r`n"; ExitCode = 0 }
                }
                '/resync' {
                    [void]$script:Calls.Add('resync')
                    if ("$script:Resync" -eq 'TIMEOUT') { return $null }
                    return [pscustomobject]@{ Text = "Sending resync command to local computer`r`n"; ExitCode = [int]$script:Resync }
                }
                '/query' {
                    [void]$script:Calls.Add('source')
                    return [pscustomobject]@{ Text = "time.example.test,0x9`r`n"; ExitCode = 0 }
                }
            }
        }
        # What the check called until 6.3.0.
        function Start-Job { [void]$script:Calls.Add('stripchart'); $script:Pending = $script:Strip.Dequeue(); return 'job' }
        function Wait-Job { if ("$script:Pending" -eq 'TIMEOUT') { return $null }; return 'job' }
        function Receive-Job { return "20:22:22, $(Format-Offset $script:Pending)s" }
        function Stop-Job { }
        function Remove-Job { }
        function w32tm {
            [void]$script:Calls.Add('resync')
            if ("$script:Resync" -eq 'TIMEOUT') { $global:LASTEXITCODE = 1460 } else { $global:LASTEXITCODE = [int]$script:Resync }
            return 'resync output'
        }
        foreach ($d in $clockDefs) { . ([scriptblock]::Create($d.Extent.Text)) }

        $script:ResyncBackoffUntil = [datetime]::MinValue
        $script:CheckErrorsSaid = @{}
        function Invoke-Clock {
            param([object[]]$Strip, $Resync = 0, [switch]$KeepGate)
            $script:Strip.Clear()
            foreach ($s in $Strip) { $script:Strip.Enqueue($s) }
            $script:Resync = $Resync
            $script:Calls.Clear()
            $script:WatchLines.Clear()
            if (-not $KeepGate) { $script:LastTimeSyncCheck = [datetime]::MinValue }
            $ret = Test-TimeSyncHealth
            return [pscustomobject]@{
                Ret   = $ret
                Calls = ($script:Calls -join ',')
                Log   = ($script:WatchLines -join ' | ')
                Hold  = $script:ResyncBackoffUntil
            }
        }
        $c = [ordered]@{}
        $c['fixed']    = Invoke-Clock -Strip 10, 2 -Resync 0
        $script:ResyncBackoffUntil = [datetime]::MinValue
        $c['stays']    = Invoke-Clock -Strip 10, 10 -Resync 0
        $c['held']     = Invoke-Clock -Strip 10 -Resync 0
        $c['heldFar']  = Invoke-Clock -Strip 40 -Resync 0
        $c['gate']     = Invoke-Clock -Strip 10 -Resync 0 -KeepGate
        $script:ResyncBackoffUntil = [datetime]::UtcNow.AddMinutes(-1)
        $c['expired']  = Invoke-Clock -Strip 10, 3 -Resync 0
        $script:ResyncBackoffUntil = [datetime]::UtcNow.AddMinutes(30)
        $c['back']     = Invoke-Clock -Strip 1 -Resync 0
        $script:ResyncBackoffUntil = [datetime]::MinValue
        $c['exit5']    = Invoke-Clock -Strip 10 -Resync 5
        $script:ResyncBackoffUntil = [datetime]::MinValue
        $c['resyncTo'] = Invoke-Clock -Strip 10 -Resync 'TIMEOUT'
        $script:ResyncBackoffUntil = [datetime]::MinValue
        $c['noAfter']  = Invoke-Clock -Strip 10, 'TIMEOUT' -Resync 0
        $script:ResyncBackoffUntil = [datetime]::MinValue
        $c['stripTo']  = Invoke-Clock -Strip 'TIMEOUT' -Resync 0
        # The last check 90 minutes ahead, as after a resync that set the
        # clock back by that much.
        $script:ResyncBackoffUntil = [datetime]::MinValue
        $script:LastTimeSyncCheck  = [datetime]::UtcNow.AddMinutes(90)
        $c['gateAhead'] = Invoke-Clock -Strip 1 -Resync 0 -KeepGate
        $c
    }
    if ($clock) {
        $now = [datetime]::UtcNow
        $x = $clock['fixed']
        Assert-Case "drift 10s, 2s after the resync: REPAIRED, measured again" `
            ($null -eq $x.Ret -and $x.Calls -eq 'stripchart,resync,stripchart' -and $x.Log -match 'REPAIRED: Forced NTP resync \(drift was 10s, now 2s\)') `
            "calls $($x.Calls); log $($x.Log); returned $($x.Ret)"
        $x = $clock['stays']
        Assert-Case "still 10s after a resync that exited 0: not REPAIRED" ($x.Log -notmatch 'REPAIRED') "log $($x.Log)"
        Assert-Case "a warning that names the time source" `
            ("$($x.Ret)" -match '^Clock drift 10s and it is still 10s after w32tm /resync; time source: time\.example\.test,0x9\. No resync is forced for 60 min') `
            "returned $($x.Ret)"
        Assert-Case "and no resync for 60 minutes" ($x.Hold -gt $now.AddMinutes(59) -and $x.Hold -lt $now.AddMinutes(61)) "hold until $($x.Hold)"
        $x = $clock['held']
        Assert-Case "held: measured, no resync, no warning under 30s" ($x.Calls -eq 'stripchart' -and $null -eq $x.Ret) "calls $($x.Calls); returned $($x.Ret)"
        $x = $clock['heldFar']
        Assert-Case "held: over 30s still warns, without a resync" `
            ($x.Calls -eq 'stripchart' -and "$($x.Ret)" -eq 'Clock drift 40s, may cause VM connectivity issues') "calls $($x.Calls); returned $($x.Ret)"
        $x = $clock['gate']
        Assert-Case "within 5 minutes of the last check: nothing runs" ($x.Calls -eq '' -and $null -eq $x.Ret) "calls $($x.Calls)"
        $x = $clock['expired']
        Assert-Case "the hold over: a resync again" `
            ($x.Calls -eq 'stripchart,resync,stripchart' -and $x.Log -match 'REPAIRED' -and $x.Hold -eq [datetime]::MinValue) "calls $($x.Calls); log $($x.Log)"
        $x = $clock['back']
        Assert-Case "back within 5s during a hold: the hold ends, and says so" `
            ($x.Calls -eq 'stripchart' -and $x.Hold -eq [datetime]::MinValue -and $x.Log -match 'Clock drift back to 1s; forced resyncs allowed again') `
            "calls $($x.Calls); hold $($x.Hold); log $($x.Log)"
        $x = $clock['exit5']
        Assert-Case "a resync that exits 5: a warning naming the exit and the source" `
            ("$($x.Ret)" -match 'w32tm /resync failed \(exit 5\); time source: time\.example\.test' -and $x.Hold -gt $now) "returned $($x.Ret)"
        $x = $clock['resyncTo']
        Assert-Case "a resync that does not answer: said so" ("$($x.Ret)" -match 'w32tm /resync did not answer within 15s; time source') "returned $($x.Ret)"
        $x = $clock['noAfter']
        Assert-Case "no measurement after the resync: not REPAIRED" `
            ($x.Log -notmatch 'REPAIRED' -and "$($x.Ret)" -match 'the clock could not be measured after w32tm /resync') "returned $($x.Ret); log $($x.Log)"
        # The warning says the clock could not be measured after the resync,
        # so the log does not also say the check was skipped. Until 6.3.0 it
        # did.
        Assert-Case "and the log says the measurement after the resync did not answer, not that the check was skipped" `
            ($x.Log -eq 'w32tm stripchart did not answer within 15s after the resync') "log $($x.Log)"
        $x = $clock['stripTo']
        Assert-Case "the measurement does not answer: skipped, no resync" `
            ($null -eq $x.Ret -and $x.Calls -eq 'stripchart' -and $x.Log -match 'skipping the clock-drift check') "calls $($x.Calls); log $($x.Log)"
        Assert-Case "and that line is the one it always was" `
            ($x.Log -eq 'w32tm stripchart did not answer within 15s; skipping the clock-drift check') "log $($x.Log)"
        $x = $clock['gateAhead']
        Assert-Case "a last check in the future, after the clock was set back: the check runs, rather than waiting until then" `
            ($x.Calls -eq 'stripchart' -and $null -eq $x.Ret -and $x.Log -eq '') "calls $($x.Calls); log $($x.Log); returned $($x.Ret)"
        # A check that fails part-way logs it, so no case above may have.
        $broke = @($clock.Keys | Where-Object { $clock[$_].Log -match 'The clock check failed' })
        Assert-Case "no case failed part-way" ($broke.Count -eq 0) "failed: $($broke -join ', ')"
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- the monitor's Invoke-W32tm and Invoke-HcsDiag against stand-ins"

# Batch files in a temp folder stand in for w32tm.exe and hcsdiag.exe; the
# real paths in the lifted source are swapped for them, and the swap is
# checked before anything runs.
$standW32 = Join-Path $tempRoot 'w32tm-standin.cmd'
[System.IO.File]::WriteAllText($standW32, ((@(
    '@echo off',
    'if "%1"=="/hang" (',
    '  ping -n 5 127.0.0.1 >nul',
    '  exit /b 0',
    ')',
    'echo 20:22:22, +01.5000000s',
    'exit /b %2'
) -join "`r`n") + "`r`n"), [System.Text.Encoding]::ASCII)
$standHcs = Join-Path $tempRoot 'hcsdiag-standin.cmd'
[System.IO.File]::WriteAllText($standHcs, ((@(
    '@echo off',
    'set /p MODE=<"%~dp0hcsmode.txt"',
    'if "%MODE%"=="list" (',
    '  echo cowork-vm-de1517ec',
    '  echo     VM, Running, DE1517EC-BF1D-5A0B-B459-77A48CB1AD97, cowork-vm-de1517ec',
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
    'exit /b 9'
) -join "`r`n") + "`r`n"), [System.Text.Encoding]::ASCII)
function Set-HcsMode { param([string]$Mode) [System.IO.File]::WriteAllText((Join-Path $tempRoot 'hcsmode.txt'), "$Mode`r`n", [System.Text.Encoding]::ASCII) }

$w32Def = Get-AstFunction -Ast $watchAst -Name 'Invoke-W32tm'
Assert-Case "Invoke-W32tm was found" ($null -ne $w32Def)
if ($w32Def) {
    $w32Literal = '"$env:SystemRoot\System32\w32tm.exe"'
    $w32Swaps = ([regex]::Matches($w32Def.Extent.Text, [regex]::Escape($w32Literal))).Count
    $w32Text = $w32Def.Extent.Text.Replace($w32Literal, "'$standW32'")
    $w32Safe = ($w32Swaps -eq 1 -and $w32Text -notmatch 'System32\\w32tm')
    Assert-Case "the path of w32tm.exe was swapped for the stand-in" $w32Safe "swapped $w32Swaps time(s)"
    if ($w32Safe) {
        $w = Invoke-Isolated -Name 'Invoke-W32tm' -Block {
            Set-StrictMode -Version Latest
            . ([scriptblock]::Create($w32Text))
            $o = [ordered]@{}
            $o['seven'] = Invoke-W32tm -Arguments '/code', '7'
            $o['zero']  = Invoke-W32tm -Arguments '/code', '0'
            $o['hang']  = Invoke-W32tm -Arguments '/hang' -TimeoutSeconds 2
            $o
        }
        if ($w) {
            Assert-Case "the exit code comes back with the text" `
                ($null -ne $w['seven'] -and $w['seven'].ExitCode -eq 7 -and "$($w['seven'].Text)" -match '\+01\.5000000s') "got $($w['seven'] | Out-String)"
            Assert-Case "exit 0 comes back as 0" ($null -ne $w['zero'] -and $w['zero'].ExitCode -eq 0) "got $($w['zero'] | Out-String)"
            Assert-Case "no answer in time: null" ($null -eq $w['hang']) "got $($w['hang'] | Out-String)"
        }
    }
}

$hcsDef = Get-AstFunction -Ast $watchAst -Name 'Invoke-HcsDiag'
Assert-Case "the monitor's Invoke-HcsDiag was found" ($null -ne $hcsDef)
if ($hcsDef) {
    $hcsLiteral = '"$env:SystemRoot\System32\hcsdiag.exe"'
    $hcsSwaps = ([regex]::Matches($hcsDef.Extent.Text, [regex]::Escape($hcsLiteral))).Count
    $hcsText = $hcsDef.Extent.Text.Replace($hcsLiteral, "'$standHcs'")
    $hcsSafe = ($hcsSwaps -eq 1 -and $hcsText -notmatch 'System32\\hcsdiag')
    Assert-Case "the path of hcsdiag.exe was swapped for the stand-in" $hcsSafe "swapped $hcsSwaps time(s)"
    if ($hcsSafe) {
        $h = Invoke-Isolated -Name "the monitor's Invoke-HcsDiag" -Block {
            Set-StrictMode -Version Latest
            $script:WatchLines = [System.Collections.ArrayList]::new()
            $script:LastHcsDiagNote = $null
            function Write-WatchLog { param([string]$Message) [void]$script:WatchLines.Add($Message) }
            . ([scriptblock]::Create($hcsText))
            $o = [ordered]@{}
            Set-HcsMode list
            $o['list'] = Invoke-HcsDiag -Arguments 'list'
            Set-HcsMode empty
            $o['empty'] = Invoke-HcsDiag -Arguments 'list'
            Set-HcsMode error
            $o['error1'] = Invoke-HcsDiag -Arguments 'list'
            $o['log1'] = $script:WatchLines.Count
            $o['error2'] = Invoke-HcsDiag -Arguments 'list'
            $o['log2'] = $script:WatchLines.Count
            Set-HcsMode other
            $o['other'] = Invoke-HcsDiag -Arguments 'list'
            $o['log3'] = $script:WatchLines.Count
            Set-HcsMode list
            $null = Invoke-HcsDiag -Arguments 'list'
            Set-HcsMode other
            $null = Invoke-HcsDiag -Arguments 'list'
            $o['log4'] = $script:WatchLines.Count
            $o['lines'] = ($script:WatchLines -join ' | ')
            # A job that could not run hcsdiag: no exit code comes back, only
            # the job's error. Said once an hour, naming it; until 6.3.1 it
            # was $null without a word. The real job runs and is drained.
            foreach ($n in @('Get-CheckErrorText', 'Test-CheckErrorDue')) {
                $fd = Get-AstFunction -Ast $watchAst -Name $n
                if ($fd) { . ([scriptblock]::Create($fd.Extent.Text)) }
            }
            $script:CheckErrorsSaid = @{}
            function Receive-Job {
                [CmdletBinding()]
                param([Parameter(Position = 0)]$Job)
                $null = Microsoft.PowerShell.Core\Receive-Job -Job $Job
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [System.Management.Automation.CommandNotFoundException]::new("The term 'hcsdiag.exe' is not recognized as the name of a cmdlet"),
                        'CommandNotFoundException', [System.Management.Automation.ErrorCategory]::ObjectNotFound, $null))
            }
            $script:WatchLines.Clear()
            Set-HcsMode list
            $o['jobFail'] = Invoke-HcsDiag -Arguments 'list'
            $o['jobFail2'] = Invoke-HcsDiag -Arguments 'list'
            $o['jobFailLines'] = ($script:WatchLines -join ' | ')
            $o
        }
        if ($h) {
            Assert-Case "exit 0: the text" ("$($h['list'])" -match 'DE1517EC-BF1D-5A0B-B459-77A48CB1AD97') "got [$($h['list'])]"
            Assert-Case "exit 0 and nothing printed: '', not null" ($null -ne $h['empty'] -and "$($h['empty'])".Trim() -eq '') "got [$($h['empty'])]"
            Assert-Case "exit 5 with text: null" ($null -eq $h['error1']) "got [$($h['error1'])]"
            Assert-Case "and it is logged with its exit code" `
                ($h['log1'] -eq 1 -and $h['lines'] -match '^hcsdiag list exited 5: The operation could not be started') $h['lines']
            Assert-Case "the same failure again: not logged again" ($null -eq $h['error2'] -and $h['log2'] -eq 1) "log lines $($h['log2'])"
            Assert-Case "a different failure: logged" ($null -eq $h['other'] -and $h['log3'] -eq 2) "log lines $($h['log3'])"
            Assert-Case "a failure after an answer: logged again" ($h['log4'] -eq 3) "log lines $($h['log4']); $($h['lines'])"
            Assert-Case "a job that could not run hcsdiag: null, and said once an hour, naming why" ($null -eq $h['jobFail'] -and $null -eq $h['jobFail2'] -and
                $h['jobFailLines'] -eq ("hcsdiag list could not be run (The term 'hcsdiag.exe' is not recognized as the name of a cmdlet); " +
                    'the same error is logged at most once an hour')) $h['jobFailLines']
        }
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- ages in UTC"

# Local time jumps an hour when summer time starts or ends. Until 6.3.0
# the "how long ago" checks outside the shared region, the monitor's
# cooldown, its rate limits and its count of recent fixes compared local
# times, and for that hour read a fresh file as old or an old one as fresh,
# or a fix a minute ago as an hour ago.
#
# Each line outside the region that reads the local clock in a sum or a
# comparison, keeps it for one, or reads a file's local write time, is a
# hit, unless the list below names it. The clock read only to be shown, with
# -Format, -f or ToString, is not. Each entry must match as many lines as it
# says, so an entry that no longer matches goes red too.
function Get-OutsideRegion {
    param([string]$Text)
    $lines = $Text -split "\r?\n"
    $out = @()
    $inRegion = $false
    for ($k = 0; $k -lt $lines.Count; $k++) {
        if ($lines[$k] -match '^#region ClaudeEnv') { $inRegion = $true; continue }
        if ($lines[$k] -match '^#endregion ClaudeEnv') { $inRegion = $false; continue }
        if (-not $inRegion) { $out += [pscustomobject]@{ No = $k + 1; Text = $lines[$k] } }
    }
    return $out
}
function Get-EnclosingFunction {
    param($Ast, [int]$Line)
    $best = $null
    foreach ($f in @($Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true))) {
        if ($f.Extent.StartLineNumber -le $Line -and $f.Extent.EndLineNumber -ge $Line) {
            if ($null -eq $best -or $f.Extent.StartLineNumber -gt $best.Extent.StartLineNumber) { $best = $f }
        }
    }
    if ($best) { return $best.Name }
    return '<script>'
}
$localAgeRx = @(
    '\.LastWriteTime\b',
    '\.TimeCreated\s+-(gt|lt|ge|le)\b',
    '-\s*\$youngest\)',
    '-\s*\$ld\.Newest\)',
    # The local clock in a sum or a comparison, either side.
    '\(Get-Date\)(\.Add\w+\([^)]*\))?\s*-(?!f\b)\w*',
    '(\s-|-(lt|le|gt|ge|eq|ne))\s*\(Get-Date\)',
    # The local clock kept, for a sum or a comparison later.
    '=\s*\(?Get-Date\b(?!\s+-Format)(?!\)\.ToString\()',
    '\.Add\(\(?Get-Date\b',
    '\[datetime\]::(Now|Today)\b'
)
$stayLocal = @(
    @{ File = 'Fix'; Function = 'Test-CoworkServiceLog'; Rx = '^\s*\$now = Get-Date$'; Count = 1
       Why = 'compared with stamps parsed from log text the service writes in local time, with no offset' },
    @{ File = 'Watch'; Function = 'Test-CoworkSessionActive'; Rx = '^\s*\$now\s+= Get-Date$'; Count = 1
       Why = 'compared with stamps parsed from log text the app writes in local time, with no offset' }
)
foreach ($s in @(@{ Name = 'Fix'; Text = $fixSrc; Ast = $fixAst }, @{ Name = 'Watch'; Text = $watchSrc; Ast = $watchAst })) {
    $hits = @()
    $allowedSeen = @{}
    foreach ($l in (Get-OutsideRegion -Text $s.Text)) {
        if ($l.Text.TrimStart().StartsWith('#')) { continue }
        $hit = $false
        foreach ($rx in $localAgeRx) { if ($l.Text -match $rx) { $hit = $true; break } }
        if (-not $hit) { continue }
        $fn = Get-EnclosingFunction -Ast $s.Ast -Line $l.No
        $allowed = $null
        for ($a = 0; $a -lt $stayLocal.Count; $a++) {
            $e = $stayLocal[$a]
            if ($e.File -eq $s.Name -and ($e.Function -eq '*' -or $e.Function -eq $fn) -and $l.Text -match $e.Rx) { $allowed = $a; break }
        }
        if ($null -ne $allowed) { $allowedSeen[$allowed] = 1 + [int]$allowedSeen[$allowed]; continue }
        $hits += "$($l.No) ($fn): $($l.Text.Trim())"
    }
    Assert-Case "$($s.Name): no age measured in local time outside the region" ($hits.Count -eq 0) ($hits -join ' | ')
    for ($a = 0; $a -lt $stayLocal.Count; $a++) {
        $e = $stayLocal[$a]
        if ($e.File -ne $s.Name) { continue }
        $n = [int]$allowedSeen[$a]
        Assert-Case "$($s.Name) keeps local time in $($e.Function): $($e.Why)" ($n -eq $e.Count) "matched $n line(s), not $($e.Count)"
    }
}

# Get-WinEvent's StartTime, outside the region. Given a [datetime], it reads
# it as local time whatever its Kind; the UTC time as round-trip text, which
# ends in Z, it reads as the instant it names (measured on Windows
# PowerShell 5.1, in summer time and across the hour the clocks went back).
# Until 6.3.0 every window was the local time, (Get-Date).Add...(-N), which
# in that hour is an hour off. Each is the UTC text, or in Watch's
# Test-EventLogErrors $since, which is. The counts are the windows each
# script has, so one that is no longer found goes red too.
$utcWindowRx = '^\s*StartTime\s*=\s*\[datetime\]::UtcNow\.Add(Seconds|Minutes|Hours|Days)\(-\d+\)\.ToString\(''o''\)$'
foreach ($s in @(@{ Name = 'Fix'; Text = $fixSrc; Ast = $fixAst; Utc = 6; Since = 0 },
                 @{ Name = 'Watch'; Text = $watchSrc; Ast = $watchAst; Utc = 1; Since = 4 })) {
    $windows = @(Get-OutsideRegion -Text $s.Text | Where-Object { $_.Text -match '^\s*StartTime\s*=' })
    $utcWindows = @($windows | Where-Object { $_.Text -match $utcWindowRx })
    $sinceWindows = @($windows | Where-Object { $_.Text -match '^\s*StartTime\s*=\s*\$since$' -and
        (Get-EnclosingFunction -Ast $s.Ast -Line $_.No) -eq 'Test-EventLogErrors' })
    $known = @(@($utcWindows) + @($sinceWindows) | ForEach-Object { $_.No })
    $other = @($windows | Where-Object { $known -notcontains $_.No } | ForEach-Object { "$($_.No): $($_.Text.Trim())" })
    Assert-Case "$($s.Name): every Event Log window outside the region is the UTC time as round-trip text" `
        ($other.Count -eq 0 -and $utcWindows.Count -eq $s.Utc -and $sinceWindows.Count -eq $s.Since) `
        "UTC $($utcWindows.Count) of $($s.Utc), `$since $($sinceWindows.Count) of $($s.Since); other: $($other -join ' | ')"
}
$sinceLines = @(Get-OutsideRegion -Text $watchSrc | Where-Object { $_.Text -match '^\s*\$since\s*=' -and
    (Get-EnclosingFunction -Ast $watchAst -Line $_.No) -eq 'Test-EventLogErrors' })
Assert-Case "Watch: Test-EventLogErrors' `$since is the UTC time as round-trip text" `
    ($sinceLines.Count -eq 1 -and $sinceLines[0].Text -match '^\s*\$since = \[datetime\]::UtcNow\.AddSeconds\(-\$lookback\)\.ToString\(''o''\)$') `
    (@($sinceLines | ForEach-Object { $_.Text.Trim() }) -join ' | ')

# The scan finds the shapes it is there for: each line below is one this
# toolkit had, or a regression could bring back.
$scanMust = @(
    '    $now = Get-Date',
    '    $elapsed = ((Get-Date) - $script:LastFixTime).TotalMinutes',
    '        $script:LastFixTime = (Get-Date).AddMinutes(30 - $Cooldown)',
    '    $null = $script:FixHistory.Add((Get-Date))',
    '    if (-not $Force -and (Get-Date) -lt $script:NextDataPathSync) { return }',
    '        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |',
    '    $t = [datetime]::Now'
)
$scanMustNot = @(
    '$stamp = (Get-Date).ToString(''yyyy-MM-dd HH:mm:ss'')',
    '$stamp     = Get-Date -Format "yyyyMMdd_HHmmss"',
    '$LogFile = "{0:yyyyMMdd}" -f (Get-Date)',
    '    $now = [datetime]::UtcNow'
)
$missed = @($scanMust | Where-Object { $t = $_; -not @($localAgeRx | Where-Object { $t -match $_ }).Count })
$wrong  = @($scanMustNot | Where-Object { $t = $_; @($localAgeRx | Where-Object { $t -match $_ }).Count })
Assert-Case "the scan finds every local-clock shape it is given" ($missed.Count -eq 0) ($missed -join ' | ')
Assert-Case "and none of the display-only or UTC ones" ($wrong.Count -eq 0) ($wrong -join ' | ')

# Test-ClaudeCodeActive with the pair a real file shows across the change to
# summer time: its local time and its UTC time an hour apart.
$codeDef = Get-AstFunction -Ast $watchAst -Name 'Test-ClaudeCodeActive'
# The region's error helpers, lifted: the Code check tells a folder that is
# not there from one that could not be read with them. Test-Path is what a
# tree before 6.3.1 asks.
$regionErrNames = @('Test-ClaudeNotFoundError', 'Get-ClaudeErrorText')
$regionErrDefs = @($regionErrNames | ForEach-Object { Get-AstFunction -Ast $watchAst -Name $_ } | Where-Object { $_ })
$codeFakes = @('Get-ClaudeAppCliProcess', 'Get-ClaudeDataCandidates', 'Test-Path', 'Get-ChildItem', 'Get-ClaudeItemOrNull',
    'Test-ClaudeCodeTranscriptRecent', 'Write-WatchLog', 'Write-CheckError') + $regionErrNames
$unfaked = @(if ($codeDef) { Get-UnfakedCommand -Asts (@($codeDef) + $regionErrDefs) -Faked $codeFakes } else { '<not found>' })
Assert-Case "Test-ClaudeCodeActive was found, and every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $ages = Invoke-Isolated -Name 'Test-ClaudeCodeActive' -Block {
        Set-StrictMode -Version Latest
        $script:ClaudeEnv = [pscustomobject]@{ PackageFamilyName = 'Test_pfn'; ProcessName = 'claude' }
        $script:NoCodeSessionDirLogged = $true
        function Get-ClaudeAppCliProcess { param($Environment) $null = $Environment; return @() }
        function Get-ClaudeDataCandidates {
            param($PackageFamilyName)
            $null = $PackageFamilyName
            return [pscustomobject]@{ DataRoots = @([pscustomobject]@{ Path = 'C:\fake-root' }); LogDirs = @(); Errors = @() }
        }
        function Test-Path { return $true }
        function Get-ClaudeItemOrNull { param([string]$LiteralPath) $null = $LiteralPath; return [pscustomobject]@{ PSIsContainer = $true } }
        function Get-ChildItem { return $script:DstFile }
        function Write-CheckError { param([string]$Check, [string]$Text, $ErrorRecord) $null = $Check, $Text, $ErrorRecord }
        foreach ($d in $regionErrDefs) { . ([scriptblock]::Create($d.Extent.Text)) }
        function Test-ClaudeCodeTranscriptRecent { param([int]$Minutes) $null = $Minutes; return $false }
        function Write-WatchLog { param([string]$Message) $null = $Message }
        . ([scriptblock]::Create($codeDef.Extent.Text))
        $o = [ordered]@{}
        $script:DstFile = [pscustomobject]@{
            LastWriteTime = (Get-Date).AddMinutes(-65); LastWriteTimeUtc = [datetime]::UtcNow.AddMinutes(-5) }
        $o['fresh'] = Test-ClaudeCodeActive
        $script:DstFile = [pscustomobject]@{
            LastWriteTime = (Get-Date).AddMinutes(-5); LastWriteTimeUtc = [datetime]::UtcNow.AddMinutes(-70) }
        $o['old'] = Test-ClaudeCodeActive
        $o
    }
    if ($ages) {
        Assert-Case "a session file 5 min old (65 by the local clock): a Code session" ($ages['fresh'] -eq $true) "got $($ages['fresh'])"
        Assert-Case "a session file 70 min old (5 by the local clock): none" ($ages['old'] -eq $false) "got $($ages['old'])"
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- what is said after an HCS kill"

# The check after the kill reads the list once or twice, a few seconds
# apart. Still listed then is not proof the kill failed for good.
Assert-Case "the monitor says a compute system was still listed after the kill" `
    ($watchSrc -match 'Write-WatchLog "HCS: stale cowork-vm \(\$guid\) was still listed after the kill"' -and
     $watchSrc -notmatch 'hcsdiag kill did not remove')
Assert-Case "and so does Fix" `
    ($fixSrc -match 'Log "\[!\] compute system \$g was still listed after the kill"' -and $fixSrc -notmatch 'hcsdiag kill did not remove')

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- what -Close says at the end"

# The last check reads the app's processes, the service and the hcsdiag
# list. When hcsdiag does not answer, whether a compute system is left is
# not known, and the line does not call the shutdown complete. Until 6.3.0
# it did. The same for a process list that could not be read, since 6.3.1.
$closeIf = Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and
    $n.Clauses[0].Item1.Extent.Text -eq '-not $procsUnknown -and $remainingProcs.Count -eq 0 -and -not $remainingHcs' }
Assert-Case "the closing check was found" ($closeIf.Count -eq 1) "found $($closeIf.Count)"
if ($closeIf.Count -eq 1) {
    $unfaked = @(Get-UnfakedCommand -Asts $closeIf -Faked @('Log'))
    Assert-Case "every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
    if ($unfaked.Count -eq 0) {
        $closeSaid = Invoke-Isolated -Name 'the closing check' -Block {
            Set-StrictMode -Version Latest
            function Log { param([string]$Message, [string]$Colour, [switch]$Indent) $null = $Colour, $Indent; [void]$script:CloseLines.Add($Message) }
            $closeBlock = [scriptblock]::Create($closeIf[0].Extent.Text)
            $o = [ordered]@{}
            foreach ($c in @(@{ Name = 'unknownUp'; Unknown = $true; Up = $true }, @{ Name = 'unknownDown'; Unknown = $true; Up = $false },
                             @{ Name = 'knownUp'; Unknown = $false; Up = $true },
                             @{ Name = 'procsUnread'; Unknown = $false; Up = $true; ProcsUnread = $true },
                             @{ Name = 'svcUnread'; Unknown = $false; Up = $false; SvcUnknown = $true },
                             @{ Name = 'hcsSvcUnread'; Unknown = $true; Up = $false; SvcUnknown = $true })) {
                $script:CloseLines = [System.Collections.ArrayList]::new()
                $procsUnknown   = [bool]$c['ProcsUnread']
                $remainingProcs = @()
                $remainingHcs   = $false
                $hcsUnknown     = $c.Unknown
                $svcRunning     = $c.Up
                $svcUnknown     = [bool]$c['SvcUnknown']
                . $closeBlock
                $o[$c.Name] = ($script:CloseLines -join ' | ')
            }
            $o
        }
        if ($closeSaid) {
            Assert-Case "hcsdiag did not answer, the service running: not called complete" `
                ($closeSaid['unknownUp'] -notmatch 'complete' -and $closeSaid['unknownUp'] -match 'the service is running; whether a compute system is left is not known') $closeSaid['unknownUp']
            Assert-Case "hcsdiag did not answer, the service stopped: not called complete" `
                ($closeSaid['unknownDown'] -notmatch 'complete' -and $closeSaid['unknownDown'] -match 'the service is not running; whether a compute system is left is not known') $closeSaid['unknownDown']
            Assert-Case "hcsdiag answered, nothing left: complete" ($closeSaid['knownUp'] -eq 'Clean shutdown complete (service idle, ready for relaunch)') $closeSaid['knownUp']
            Assert-Case "a process list that could not be read: not called complete" ($closeSaid['procsUnread'] -notmatch 'complete') $closeSaid['procsUnread']
            # Since 6.3.1 a service that could not be read is not known to
            # run or not; until 6.3.1 it read as not running.
            Assert-Case "nothing left, the service could not be read: complete, and whether it runs is not known" `
                ($closeSaid['svcUnread'] -eq 'Clean shutdown complete (whether the service is running is not known)') $closeSaid['svcUnread']
            Assert-Case "hcsdiag did not answer, the service could not be read: both not known" `
                ($closeSaid['hcsSvcUnread'] -match 'No app processes are left and whether the service is running is not known; whether a compute system is left is not known') `
                $closeSaid['hcsSvcUnread']
        }
    }
}

# What it says on the console after that. The green lines, that the app is
# shut down and that a relaunch should work at once, follow only the branch
# that calls the shutdown clean. Until 6.3.0 they followed every outcome, so
# a warning that processes or a compute system were left, or that hcsdiag
# had not answered, was followed by "Claude Desktop is shut down." in green.
# Lifted from the end of the line that reads the service's state to the
# Save-Log after the summary, so the closing check runs with it.
$svcAssign = Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$svcRunning' }
$closeSave = @()
if ($svcAssign.Count -eq 1) {
    $closeSave = @(Find-FixStatement { param($n)
        $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Save-Log' -and
        $n.Extent.StartOffset -gt $svcAssign[0].Extent.EndOffset } | Sort-Object { $_.Extent.StartOffset } | Select-Object -First 1)
}
Assert-Case "the summary after the closing check was found" ($svcAssign.Count -eq 1 -and $closeSave.Count -eq 1) `
    "found $($svcAssign.Count) `$svcRunning line(s), $($closeSave.Count) Save-Log"
if ($svcAssign.Count -eq 1 -and $closeSave.Count -eq 1) {
    $summaryFrom = $svcAssign[0].Extent.EndOffset
    $summaryText = $fixSrc.Substring($summaryFrom, $closeSave[0].Extent.StartOffset - $summaryFrom)
    $summaryAst  = [System.Management.Automation.Language.Parser]::ParseInput($summaryText, [ref]$null, [ref]$null)
    $unfaked = @(Get-UnfakedCommand -Asts @($summaryAst) -Faked @('Log', 'Write-Host'))
    Assert-Case "every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
    Assert-Case "and it holds the closing check and the lines after it" `
        ($summaryText -match 'Clean shutdown complete' -and $summaryText -match 'Claude Desktop is shut down\.')
    if ($unfaked.Count -eq 0) {
        $summarySaid = Invoke-Isolated -Name 'the summary after the closing check' -Block {
            Set-StrictMode -Version Latest
            function Log { param([string]$Message, [string]$Colour, [switch]$Indent) $null = $Colour, $Indent; [void]$script:CloseLines.Add("LOG: $Message") }
            function Write-Host {
                param([object]$Object, [string]$ForegroundColor)
                $colour = if ($ForegroundColor) { $ForegroundColor } else { 'Plain' }
                if ("$Object") { [void]$script:CloseLines.Add("[$colour] $Object") }
            }
            $summaryBlock = [scriptblock]::Create($summaryText)
            $o = [ordered]@{}
            foreach ($c in @(
                @{ Name = 'knownUp';      Procs = 0; Hcs = $false; Unknown = $false; Up = $true },
                @{ Name = 'knownDown';    Procs = 0; Hcs = $false; Unknown = $false; Up = $false },
                @{ Name = 'unknownUp';    Procs = 0; Hcs = $false; Unknown = $true;  Up = $true },
                @{ Name = 'unknownDown';  Procs = 0; Hcs = $false; Unknown = $true;  Up = $false },
                @{ Name = 'procsLeft';    Procs = 2; Hcs = $false; Unknown = $false; Up = $true },
                @{ Name = 'hcsLeft';      Procs = 0; Hcs = $true;  Unknown = $false; Up = $true },
                @{ Name = 'procsUnknown'; Procs = 2; Hcs = $false; Unknown = $true;  Up = $true },
                @{ Name = 'procsUnread';  Procs = 0; Hcs = $false; Unknown = $false; Up = $true; ProcsUnread = $true },
                @{ Name = 'procsUnreadHcs'; Procs = 0; Hcs = $false; Unknown = $true; Up = $true; ProcsUnread = $true },
                @{ Name = 'svcUnread';    Procs = 0; Hcs = $false; Unknown = $false; Up = $false; SvcUnknown = $true })) {
                $script:CloseLines = [System.Collections.ArrayList]::new()
                $procsUnknown   = [bool]$c['ProcsUnread']
                $remainingProcs = @()
                for ($k = 0; $k -lt $c.Procs; $k++) { $remainingProcs += $k }
                $remainingHcs   = $c.Hcs
                $hcsUnknown     = $c.Unknown
                $svcRunning     = $c.Up
                $svcUnknown     = [bool]$c['SvcUnknown']
                . $summaryBlock
                $o[$c.Name] = ($script:CloseLines -join ' | ')
            }
            $o
        }
        if ($summarySaid) {
            Assert-Case "nothing left, the service running: shut down, and a relaunch should work at once" `
                ($summarySaid['knownUp'] -match '\[Green\]\s+Claude Desktop is shut down\.' -and
                 $summarySaid['knownUp'] -match '\[Green\]\s+Service is idle and ready, relaunch should work immediately\.') $summarySaid['knownUp']
            Assert-Case "nothing left, the service stopped: shut down, and Fix may be needed" `
                ($summarySaid['knownDown'] -match '\[Green\]\s+Claude Desktop is shut down\.' -and
                 $summarySaid['knownDown'] -match '\[DarkYellow\]\s+Service is not running, you may need to run Fix after relaunch\.' -and
                 $summarySaid['knownDown'] -notmatch 'relaunch should work immediately') $summarySaid['knownDown']
            foreach ($k in 'unknownUp', 'unknownDown') {
                Assert-Case "$k, hcsdiag not answering: no green line, and whether a compute system is left is not known" `
                    ($summarySaid[$k] -notmatch '\[Green\]' -and $summarySaid[$k] -notmatch 'is shut down\.|relaunch should work immediately' -and
                     $summarySaid[$k] -match "\[DarkYellow\]\s+Claude Desktop's processes are stopped, but whether a compute system is left is not known\." -and
                     $summarySaid[$k] -match '\[DarkYellow\]\s+If the relaunch fails, run Fix\.') $summarySaid[$k]
            }
            foreach ($k in 'procsLeft', 'hcsLeft', 'procsUnknown') {
                Assert-Case "$k, after a warning: no green line, and not fully shut down" `
                    ($summarySaid[$k] -match 'LOG: Warning: ' -and $summarySaid[$k] -notmatch '\[Green\]' -and
                     $summarySaid[$k] -notmatch 'is shut down\.|relaunch should work immediately' -and
                     $summarySaid[$k] -match '\[DarkYellow\]\s+Claude Desktop is not fully shut down: see the warnings above\.' -and
                     $summarySaid[$k] -match '\[DarkYellow\]\s+If the relaunch fails, run Fix\.') $summarySaid[$k]
            }
            # A process list that could not be read, since 6.3.1: neither
            # shut down nor "processes are stopped", whatever hcsdiag said.
            foreach ($k in 'procsUnread', 'procsUnreadHcs') {
                Assert-Case "$k, the process list not read: no green line, and whether the processes are stopped is not known" `
                    ($summarySaid[$k] -notmatch '\[Green\]' -and $summarySaid[$k] -notmatch 'complete|is shut down\.|processes are stopped, but' -and
                     $summarySaid[$k] -match "\[DarkYellow\]\s+Whether Claude Desktop's processes are stopped is not known: see the warning above\." -and
                     $summarySaid[$k] -match '\[DarkYellow\]\s+If the relaunch fails, run Fix\.') $summarySaid[$k]
            }
            Assert-Case "nothing left, the service could not be read: shut down, and whether it runs is not known" `
                ($summarySaid['svcUnread'] -match '\[Green\]\s+Claude Desktop is shut down\.' -and
                 $summarySaid['svcUnread'] -match '\[DarkYellow\]\s+Whether the service is running is not known: see the warning above\. If the relaunch fails, run Fix\.' -and
                 $summarySaid['svcUnread'] -notmatch 'Service is not running, you may need') $summarySaid['svcUnread']
        }
    }
}

# The read before the closing check, lifted and run with the process list
# failing. Since 6.3.1 it says so and marks the list as not known; until
# 6.3.1 the list could not fail, and read as empty.
$closeRead = Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.TryStatementAst] -and @($n.Body.Statements).Count -eq 1 -and
    $n.Body.Statements[0].Extent.Text -eq '$remainingProcs = @(Get-ClaudeDesktopProcess -Environment $script:ClaudeEnv)' }
Assert-Case "the closing check's read of the processes was found" ($closeRead.Count -eq 1) "found $($closeRead.Count)"
if ($closeRead.Count -eq 1) {
    $unfaked = @(Get-UnfakedCommand -Asts $closeRead -Faked @('Log', 'Get-ClaudeDesktopProcess', 'Get-ClaudeErrorText'))
    Assert-Case "every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
    if ($unfaked.Count -eq 0) {
        $readSaid = Invoke-Isolated -Name 'the closing read' -Block {
            Set-StrictMode -Version Latest
            $lines = [System.Collections.ArrayList]::new()
            function Log { param([string]$Message, [string]$Colour, [switch]$Indent) $null = $Colour, $Indent; [void]$lines.Add($Message) }
            function Get-ClaudeDesktopProcess { param($Environment) $null = $Environment; throw 'Access is denied' }
            $script:ClaudeEnv = [pscustomobject]@{ ProcessName = 'claude' }
            $procsUnknown   = $false
            $remainingProcs = @()
            . ([scriptblock]::Create($closeRead[0].Extent.Text))
            [pscustomobject]@{ Unknown = $procsUnknown; Count = @($remainingProcs).Count; Lines = ($lines -join ' | ') }
        }
        if ($readSaid) {
            Assert-Case "a process list that could not be read is marked not known, and said" `
                ($readSaid.Unknown -and $readSaid.Count -eq 0 -and
                 $readSaid.Lines -eq 'Whether app processes are left is not known: the process list could not be read (Access is denied)') `
                "unknown $($readSaid.Unknown); said: $($readSaid.Lines)"
        }
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- what runs after Step 9 when the workspace is not ready"

# Since 6.3.1 a VM log that could not be read marks the run
# ($script:VmLogUnread), and neither the escalation to Deep, which deletes
# the VM cache, nor the retry cycle, which restarts the service, runs on it;
# the run says so and ends as a failure. Until 6.3.1 such a log read as one
# with nothing new in it, and both ran. Since 6.3.1 the escalation does not
# run with -KeepCache either, and neither runs on a service the prerequisite
# check found cannot start, when it was down as Step 9's wait began or a
# restart in the wait failed on it; each is said. Step 9 marks that with
# $prereqStillBlocked, which it does not set for a cause the check found
# with the service up or not read as the wait began and no restart in it
# failing. The flag is set only when Step 9 runs, which -SkipLaunch never
# enters, so each condition must read it after
# -not $SkipLaunch: the -SkipLaunch cases run with it not set, under
# StrictMode. The pieces are found by what they say, so a change to a
# condition is a case that fails, not a piece not found. The three that say
# something are run, and the conditions of the escalation and of the retry
# cycle are evaluated, with Log faked.
$prereqIf = Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and
    $n.Clauses[0].Item2.Extent.Text -match 'Log "Workspace not ready, but not retrying: ' }
$keepIf = Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and
    $n.Clauses[0].Item2.Extent.Text -match 'Log "Smart mode: workspace not ready, but not escalating to Deep: -KeepCache' }
$smartIf = Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and
    $n.Clauses[0].Item2.Extent.Text -match 'Log "Smart mode: workspace not ready after quick fix, escalating to Deep"' }
$retryIf = Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and
    $n.Clauses[0].Item2.Extent.Text -match 'Log "Workspace not ready, attempting retry cycle' }
$unknownIf = Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and
    $n.Clauses[0].Item2.Extent.Text -match 'Log "\[!\] Whether the workspace came up is not known' }
Assert-Case "after Step 9, the escalation to Deep, the retry cycle and the lines around them were found" `
    ($prereqIf.Count -eq 1 -and $keepIf.Count -eq 1 -and $smartIf.Count -eq 1 -and $retryIf.Count -eq 1 -and $unknownIf.Count -eq 1) `
    "found $($prereqIf.Count), $($keepIf.Count), $($smartIf.Count), $($retryIf.Count), $($unknownIf.Count)"
if ($prereqIf.Count -eq 1 -and $keepIf.Count -eq 1 -and $smartIf.Count -eq 1 -and $retryIf.Count -eq 1 -and $unknownIf.Count -eq 1) {
    $unfaked = @(Get-UnfakedCommand -Asts @($prereqIf[0], $keepIf[0], $smartIf[0].Clauses[0].Item1, $retryIf[0].Clauses[0].Item1, $unknownIf[0]) -Faked @('Log'))
    Assert-Case "after Step 9, every command the pieces call is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
    if ($unfaked.Count -eq 0) {
        # Each case differs from this by its keys: the workspace not ready,
        # launched, admin, Smart mode, not yet escalated, every VM log read,
        # no -KeepCache, and a service that can start.
        $escCases = [ordered]@{
            'True'             = @{ Unread = $true }
            'False'            = @{}
            'ready'            = @{ Ready = $true }
            'readyUnread'      = @{ Ready = $true; Unread = $true }
            'notAdmin'         = @{ NotAdmin = $true }
            'deep'             = @{ Deep = $true }
            'escalated'        = @{ Escalated = $true }
            'flagged'          = @{ Flagged = $true }
            'keep'             = @{ Keep = $true }
            'keepUnread'       = @{ Keep = $true; Unread = $true }
            'keepReady'        = @{ Keep = $true; Ready = $true }
            'keepNotAdmin'     = @{ Keep = $true; NotAdmin = $true }
            'keepDeep'         = @{ Keep = $true; Deep = $true }
            'keepEscalated'    = @{ Keep = $true; Escalated = $true }
            'blocked'          = @{ Blocked = $true }
            'blockedKeep'      = @{ Blocked = $true; Keep = $true }
            'blockedUnread'    = @{ Blocked = $true; Unread = $true }
            'blockedDeep'      = @{ Blocked = $true; Deep = $true }
            'blockedNotAdmin'  = @{ Blocked = $true; NotAdmin = $true }
            'blockedEscalated' = @{ Blocked = $true; Escalated = $true }
            'blockedReady'     = @{ Blocked = $true; Ready = $true }
            'skip'             = @{ Skip = $true; Keep = $true }
            'skipUnread'       = @{ Skip = $true; Unread = $true }
        }
        $escSaid = Invoke-Isolated -Name 'what runs after Step 9' -Block {
            Set-StrictMode -Version Latest
            $lines = [System.Collections.ArrayList]::new()
            function Log { param([string]$Message, [string]$Colour, [switch]$Indent) $null = $Colour, $Indent; if ($Message) { [void]$lines.Add($Message) } }
            $prereqRun = [scriptblock]::Create($prereqIf[0].Extent.Text)
            $keepRun = [scriptblock]::Create($keepIf[0].Extent.Text)
            $smartCond = [scriptblock]::Create($smartIf[0].Clauses[0].Item1.Extent.Text)
            $retryCond = [scriptblock]::Create($retryIf[0].Clauses[0].Item1.Extent.Text)
            $unknownRun = [scriptblock]::Create($unknownIf[0].Extent.Text)
            $o = [ordered]@{}
            foreach ($case in $escCases.Keys) {
                $c = $escCases[$case]
                $vmReady = [bool]$c['Ready']
                $SkipLaunch = [bool]$c['Skip']
                $KeepCache = [bool]$c['Keep']
                $script:IsAdmin = -not $c['NotAdmin']
                $script:SelectedMode = if ($c['Deep']) { 'Deep' } else { 'Smart' }
                $script:DeepEscalated = [bool]$c['Escalated']
                $script:VmLogUnread = [bool]$c['Unread']
                $script:ServicePrereq = $null
                # Flagged: the check found a cause, but the service was up or
                # not read as the wait began, and no restart in it failed:
                # Step 9 does not take that for a service that cannot start.
                if ($c['Blocked'] -or $c['Flagged']) { $script:ServicePrereq = [pscustomobject]@{ Blocked = $true; Reason = 'CoworkVMService start type is Disabled.' } }
                # Set only when Step 9 runs, as in Fix: not set with -SkipLaunch.
                Remove-Variable -Name prereqStillBlocked -Scope Local -ErrorAction SilentlyContinue
                if (-not $SkipLaunch) { $prereqStillBlocked = [bool]$c['Blocked'] }
                $flag = 'set'
                try { $null = $prereqStillBlocked } catch { $flag = 'not set' }
                $script:ExitCode = 0
                $lines.Clear()
                $smartGo = $null
                $retryGo = $null
                $err = ''
                try {
                    . $prereqRun
                    . $keepRun
                    $smartGo = [bool](. $smartCond)
                    $retryGo = [bool](. $retryCond)
                    . $unknownRun
                } catch {
                    $err = $_.Exception.Message
                }
                $o[$case] = "smart $smartGo, retry $retryGo, exit $($script:ExitCode), $flag, error [$err] | $($lines -join ' | ')"
            }
            $script:ServicePrereq = $null
            $o
        }
        if ($escSaid) {
            $unknownSaid = ('[!] Whether the workspace came up is not known: a VM log could not be read | ' +
                'Nothing more was escalated or retried on that. Open a Cowork session in Claude, and run Fix again if it does not start.')
            $keepSaid = 'Smart mode: workspace not ready, but not escalating to Deep: -KeepCache keeps the VM cache. Rerun without it to purge the cache.'
            $blockedSmartSaid = ('Smart mode: workspace not ready, but not escalating to Deep or retrying: CoworkVMService start type is Disabled. | ' +
                'Purging the VM cache or restarting the service cannot fix that. Resolve it and rerun.')
            $blockedSaid = ('Workspace not ready, but not retrying: CoworkVMService start type is Disabled. | ' +
                'Restarting the service cannot fix that. Resolve it and rerun.')
            # case, escalation, retry cycle, exit code, flag, lines, what it shows
            foreach ($w in @(
                    @('True', 'False', 'False', 1, 'set', $unknownSaid,
                      'a VM log not read: no escalation to Deep, no retry cycle; said, and the run ends as a failure'),
                    @('False', 'True', 'True', 0, 'set', '',
                      'logs that were read: both still run, and nothing is said'),
                    @('keep', 'False', 'True', 0, 'set', $keepSaid,
                      'with -KeepCache: no escalation to Deep, the retry cycle still runs, and it is said'),
                    @('keepUnread', 'False', 'False', 1, 'set', $unknownSaid,
                      'with -KeepCache and a VM log not read: only the line for the log not read'),
                    @('keepReady', 'False', 'False', 0, 'set', '',
                      'with -KeepCache and the workspace ready: nothing runs, and nothing is said'),
                    @('keepNotAdmin', 'False', 'True', 0, 'set', '',
                      'with -KeepCache, not admin: the retry cycle runs, and the line for -KeepCache is not said'),
                    @('keepDeep', 'False', 'True', 0, 'set', '',
                      'with -KeepCache in Deep mode: the retry cycle runs, and the line for -KeepCache is not said'),
                    @('keepEscalated', 'False', 'True', 0, 'set', '',
                      'with -KeepCache, already escalated to Deep: the retry cycle runs, and the line for -KeepCache is not said'),
                    @('ready', 'False', 'False', 0, 'set', '',
                      'the workspace ready: nothing runs, and nothing is said'),
                    @('readyUnread', 'False', 'False', 0, 'set', '',
                      'the workspace ready and a VM log not read: nothing runs, and nothing is said'),
                    @('notAdmin', 'False', 'True', 0, 'set', '',
                      'not admin: no escalation to Deep, the retry cycle runs, and nothing is said'),
                    @('deep', 'False', 'True', 0, 'set', '',
                      'Deep mode: no escalation to Deep, the retry cycle runs, and nothing is said'),
                    @('escalated', 'False', 'True', 0, 'set', '',
                      'already escalated to Deep: not again, the retry cycle runs, and nothing is said'),
                    @('flagged', 'True', 'True', 0, 'set', '',
                      'a cause the check found, the service up or not read as the wait began: both still run, and nothing is said'),
                    @('blocked', 'False', 'False', 1, 'set', $blockedSmartSaid,
                      'a service that cannot start, Smart: no escalation to Deep, no retry cycle; said with the cause, and the run ends as a failure'),
                    @('blockedKeep', 'False', 'False', 1, 'set', $blockedSmartSaid,
                      'a service that cannot start, with -KeepCache: only the line for the service'),
                    @('blockedUnread', 'False', 'False', 1, 'set', $blockedSmartSaid,
                      'a service that cannot start and a VM log not read: only the line for the service'),
                    @('blockedDeep', 'False', 'False', 1, 'set', $blockedSaid,
                      'a service that cannot start, Deep mode: no retry cycle; said with the cause, not naming the escalation'),
                    @('blockedNotAdmin', 'False', 'False', 1, 'set', $blockedSaid,
                      'a service that cannot start, not admin: said with the cause, not naming the escalation'),
                    @('blockedEscalated', 'False', 'False', 1, 'set', $blockedSaid,
                      'a service that cannot start, already escalated to Deep: said with the cause, not naming the escalation'),
                    @('blockedReady', 'False', 'False', 0, 'set', '',
                      'a service that cannot start and the workspace ready: nothing runs, and nothing is said'),
                    @('skip', 'False', 'False', 0, 'not set', '',
                      '-SkipLaunch with -KeepCache: nothing runs or is said, and the flag Step 9 sets is not read'),
                    @('skipUnread', 'False', 'False', 0, 'not set', '',
                      '-SkipLaunch with a VM log not read: nothing runs or is said, and the flag Step 9 sets is not read'))) {
                $want = "smart $($w[1]), retry $($w[2]), exit $($w[3]), $($w[4]), error [] | $($w[5])"
                Assert-Case "after Step 9, $($w[6])" ($escSaid[$w[0]] -eq $want) "got: $($escSaid[$w[0]]); want: $want"
            }
        }
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Step 9's last word on the wait"

# For a service that cannot start, when the wait did not run or stopped on
# it, Step 9 does not say the workspace was not confirmed ready, nor to open
# a Cowork session to trigger setup: the run says the cause after Step 9.
# Until 6.3.1 it said both, "after 0s" when the wait did not run.
$finalIf = @(Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and
    $n.Clauses[0].Item2.Extent.Text -match 'Log "Workspace partially ready \(log: ' })
Assert-Case "Step 9's last word on the wait was found" ($finalIf.Count -eq 1) "found $($finalIf.Count)"
if ($finalIf.Count -eq 1) {
    $unfaked = @(Get-UnfakedCommand -Asts @($finalIf[0]) -Faked @('Log', 'Test-VmLogReady'))
    Assert-Case "Step 9's last word: every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
    if ($unfaked.Count -eq 0) {
        $finalSaid = Invoke-Isolated -Name "Step 9's last word" -Block {
            Set-StrictMode -Version Latest
            $lines = [System.Collections.ArrayList]::new()
            function Log { param([string]$Message, [string]$Colour, [switch]$Indent) $null = $Colour, $Indent; [void]$lines.Add($Message) }
            function Test-VmLogReady { param($Baseline) $null = $Baseline; return $script:FrFinalLog }
            $finalRun = [scriptblock]::Create($finalIf[0].Extent.Text)
            $logBaseline = @{}
            $vmElapsed = 40
            $o = [ordered]@{}
            foreach ($c in @(
                    @{ N = 'blocked'; Blocked = $true },
                    @{ N = 'blockedUnread'; Blocked = $true; Unread = $true },
                    @{ N = 'blockedPartial'; Blocked = $true; Log = 'vsock-connected' },
                    @{ N = 'notReady' },
                    @{ N = 'unread'; Unread = $true },
                    @{ N = 'partial'; Log = 'vsock-connected' },
                    @{ N = 'ready'; Ready = $true })) {
                $vmReady = [bool]$c['Ready']
                $prereqStillBlocked = [bool]$c['Blocked']
                $script:VmLogUnread = [bool]$c['Unread']
                $script:FrFinalLog = $c['Log']
                $lines.Clear()
                $err = ''
                try { . $finalRun } catch { $err = $_.Exception.Message }
                $o[$c.N] = "error [$err] | $($lines -join ' | ')"
            }
            $script:VmLogUnread = $false
            $o
        }
        if ($finalSaid) {
            $none = 'error [] | '
            Assert-Case "Step 9, a service that cannot start: the wait's last word is not said, with every VM log read, one not read, or the log partly ready" `
                ($finalSaid['blocked'] -eq $none -and $finalSaid['blockedUnread'] -eq $none -and $finalSaid['blockedPartial'] -eq $none) `
                "[$($finalSaid['blocked'])] [$($finalSaid['blockedUnread'])] [$($finalSaid['blockedPartial'])]"
            Assert-Case "Step 9, the workspace not ready: not confirmed ready, and to open a Cowork session" `
                ($finalSaid['notReady'] -eq "$($none)[!] Workspace not confirmed ready after 40s | Open a Cowork session in Claude to trigger setup") $finalSaid['notReady']
            Assert-Case "Step 9, a VM log not read: not confirmed ready, and why" `
                ($finalSaid['unread'] -eq "$($none)[!] Workspace not confirmed ready after 40s: a VM log could not be read (see above)") $finalSaid['unread']
            Assert-Case "Step 9, the log partly ready: said" `
                ($finalSaid['partial'] -eq "$($none)Workspace partially ready (log: vsock-connected), may still be loading") $finalSaid['partial']
            Assert-Case "Step 9, the workspace ready: nothing said" ($finalSaid['ready'] -eq $none) $finalSaid['ready']
        }
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- what Step 9's comment says the wait reads"

$step9 = [regex]::Match($fixSrc, '(?s)# STEP 9 -- Wait for Cowork workspace readiness.*?Log "\[9/10\]').Value
Assert-Case "Step 9's comment was found" ($step9.Length -gt 0)
Assert-Case "it does not list Get-VM, its heartbeat or a file size fallback as a signal" `
    ($step9 -notmatch '(?m)^\s*#\s+[A-E]\d?\.\s.*(Get-VM|heartbeat)' -and $step9 -notmatch 'File size stability fallback')
Assert-Case "it lists the log tail and the hcsdiag list, noted once" `
    ($step9 -match '(?m)^\s*#\s+A2\. The tail of any VM log' -and $step9 -match '(?m)^\s*#\s+B\. Whether "hcsdiag list" shows a cowork-vm compute system, noted')
# Until 6.3.0 it said a timeout or an error starts a recovery. Only a
# timeout does; an error is only logged.
Assert-Case "it says a guest timeout starts a recovery and a guest error is only logged" `
    ($step9 -notmatch 'a timeout or an error starts a recovery' -and
     $step9 -match 'A timeout, 90 s or more in, starts a recovery; an\s+(#\s+)?error is only logged\.')
# The wait's own branches, which also read how far into the wait it is, not
# the diagnostic's check of the same two states.
$guestIfs = @(Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and
    @($n.Clauses | Where-Object { $_.Item1.Extent.Text -match '^\$guestState -eq "guest-(timeout|error)" -and \$vmElapsed ' }).Count -gt 0 })
$guestCalls = @{}
foreach ($cl in @($guestIfs | ForEach-Object { $_.Clauses })) {
    if ($cl.Item1.Extent.Text -match '^\$guestState -eq "(guest-timeout|guest-error)" -and \$vmElapsed ') {
        $guestKind = $Matches[1]
        $guestCalls[$guestKind] = @($cl.Item2.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) |
            ForEach-Object { $_.GetCommandName() } | Sort-Object -Unique) -join ','
    }
}
Assert-Case "and the code agrees: a guest timeout stops the service, a guest error calls only Log" `
    ($guestIfs.Count -eq 1 -and $guestCalls['guest-timeout'] -match 'Stop-CoworkServiceWithTimeout' -and $guestCalls['guest-error'] -eq 'Log') `
    "found $($guestIfs.Count); timeout [$($guestCalls['guest-timeout'])], error [$($guestCalls['guest-error'])]"

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Step 9's note that the VM may be stuck"

# The no-progress detector, lifted from Step 9's wait and run for a few
# passes, 5 seconds apart, with everything it calls faked. The note that the
# VM may be stuck is logged once for each stretch with no log or guest
# activity, and not at all while either shows. Until 6.3.0 it was logged on
# every pass after 60 s whenever hcsdiag listed the VM, so a slow boot that
# was making progress filled the log with it.
$detectIf = Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and
    $n.Clauses[0].Item1.Extent.Text -eq '$vmElapsed -ge 60 -and -not $vmReady' }
Assert-Case "the no-progress detector was found" ($detectIf.Count -eq 1) "found $($detectIf.Count)"
# What it notes starts false once, before the wait, not on each pass.
$waitLoop = @(Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.WhileStatementAst] -and $n.Condition.Extent.Text -match '^-not \$prereqStillBlocked' })
$notedInit = @(Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
    $n.Left.Extent.Text -eq '$hcsStuckNoted' -and $n.Right.Extent.Text -eq '$false' })
Assert-Case "Step 9 sets `$hcsStuckNoted false once, before its wait" `
    ($waitLoop.Count -eq 1 -and @($notedInit | Where-Object { $_.Extent.EndOffset -lt $waitLoop[0].Extent.StartOffset }).Count -eq 1) `
    "found $($waitLoop.Count) wait loop(s), $($notedInit.Count) assignment(s)"
$detectFaked = @('Test-VmLogReady', 'Get-HcsList', 'Log', 'Assert-NoNewAppSession', 'Stop-ClaudeDesktopApp', 'Start-Sleep',
    'Close-StaleHcsVms', 'Stop-CoworkServiceWithTimeout', 'Start-Service', 'Start-ClaudeDesktop', 'Get-VmLogSizes',
    'Start-FixService', 'Get-ClaudeErrorText')
if ($detectIf.Count -eq 1) {
    $unfaked = @(Get-UnfakedCommand -Asts $detectIf -Faked $detectFaked)
    Assert-Case "every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
    if ($unfaked.Count -eq 0) {
        # Each pass: the seconds into the wait, what the VM logs show since
        # the baseline, the guest state and what hcsdiag lists.
        $passLoop = [scriptblock]::Create(@(
            'foreach ($pass in $passes) {',
            '    $vmElapsed  = $pass.At',
            '    $vmReady    = $false',
            '    $guestState = $pass.Guest',
            '    $script:DetLog = $pass.Log',
            '    $script:DetHcs = $pass.Hcs',
            $detectIf[0].Extent.Text,
            '}'
        ) -join "`r`n")
        $detected = Invoke-Isolated -Name 'the no-progress detector' -Block {
            Set-StrictMode -Version Latest
            function Log { param([string]$Message, [string]$Colour, [switch]$Indent) $null = $Colour, $Indent; [void]$script:DetLines.Add($Message) }
            function Test-VmLogReady { param($Baseline) $null = $Baseline; return $script:DetLog }
            function Get-HcsList {
                if ($script:DetHcs -eq 'THROWS') { throw 'hcsdiag did not answer' }
                return $script:DetHcs
            }
            function Assert-NoNewAppSession { param([string]$Before, [int]$ExitCode) $null = $ExitCode; [void]$script:DetLines.Add("ASKED before $Before") }
            function Stop-ClaudeDesktopApp { return 0 }
            function Start-Sleep { param([int]$Seconds) $null = $Seconds }
            function Close-StaleHcsVms { param([string]$Action) $null = $Action; return 0 }
            function Stop-CoworkServiceWithTimeout { return $true }
            function Start-Service { param([string]$Name, [string]$ErrorAction) $null = $Name, $ErrorAction }
            function Start-ClaudeDesktop { return $true }
            function Get-VmLogSizes { return @{} }
            function Start-FixService { return $true }
            function Get-ClaudeErrorText { param($ErrorRecord) return [string]$ErrorRecord.Exception.Message }
            $ServiceName = 'CoworkVMService'
            $script:StoppedPartWayExitCode = 4
            $listed = "cowork-vm-1699151a`r`n    VM, Running, DE1517EC-0000-4000-8000-77A48CB1AD97, cowork-vm-1699151a"
            $quiet  = @{ Log = $null;             Guest = 'no-polling';    Hcs = $listed }
            $busy   = @{ Log = 'vsock-connected'; Guest = 'no-polling';    Hcs = $listed }
            $guest  = @{ Log = $null;             Guest = 'guest-polling'; Hcs = $listed }
            $noVm   = @{ Log = $null;             Guest = 'no-polling';    Hcs = '' }
            $noHcs  = @{ Log = $null;             Guest = 'no-polling';    Hcs = 'THROWS' }
            # A VM log that could not be read (6.3.1): no progress is not known.
            $unread = @{ Log = 'unreadable';      Guest = 'no-polling';    Hcs = $listed }
            $o = [ordered]@{}
            foreach ($c in @(
                @{ Name = 'progressing';  Escalated = $true;  Passes = @($busy, $busy, $busy, $busy) },
                @{ Name = 'guest';        Escalated = $true;  Passes = @($guest, $guest, $guest, $guest) },
                @{ Name = 'stuck';        Escalated = $true;  Passes = @($quiet, $quiet, $quiet, $quiet) },
                @{ Name = 'stuckFirst';   Escalated = $false; Passes = @($quiet, $quiet, $quiet, $quiet) },
                @{ Name = 'twoStretches'; Escalated = $true;  Passes = @($quiet, $quiet, $busy, $quiet, $quiet) },
                @{ Name = 'noVm';         Escalated = $true;  Passes = @($noVm, $noVm, $noVm) },
                @{ Name = 'noHcs';        Escalated = $true;  Passes = @($noHcs, $noHcs, $noHcs) },
                @{ Name = 'unread';       Escalated = $false; Passes = @($unread, $unread, $unread) })) {
                $script:DetLines = [System.Collections.ArrayList]::new()
                $script:NoProgressEscalated = $c.Escalated
                $hcsStuckNoted = $false
                $logBaseline = @{}
                $passes = @()
                $at = 60
                foreach ($p in $c.Passes) { $at += 5; $passes += [pscustomobject]@{ At = $at; Log = $p.Log; Guest = $p.Guest; Hcs = $p.Hcs } }
                . $passLoop
                $o[$c.Name] = [pscustomobject]@{
                    Noted     = @($script:DetLines | Where-Object { $_ -eq 'HCS VM present but no log/guest activity, VM may be stuck' }).Count
                    Escalated = @($script:DetLines | Where-Object { $_ -match '^\[!\] No progress after' }).Count
                    Lines     = ($script:DetLines -join ' | ')
                }
            }
            $o
        }
        if ($detected) {
            Assert-Case "a boot the VM logs show progress for: not noted" ($detected['progressing'].Noted -eq 0) $detected['progressing'].Lines
            Assert-Case "one the guest polls for: not noted" ($detected['guest'].Noted -eq 0) $detected['guest'].Lines
            Assert-Case "a stuck one: noted once, not on every pass" ($detected['stuck'].Noted -eq 1) $detected['stuck'].Lines
            Assert-Case "a stuck one before the escalation: noted once, escalated once" `
                ($detected['stuckFirst'].Noted -eq 1 -and $detected['stuckFirst'].Escalated -eq 1) $detected['stuckFirst'].Lines
            Assert-Case "stuck, then activity, then stuck again: noted once for each stretch" ($detected['twoStretches'].Noted -eq 2) $detected['twoStretches'].Lines
            Assert-Case "no compute system listed: not noted" ($detected['noVm'].Noted -eq 0) $detected['noVm'].Lines
            Assert-Case "hcsdiag not answering: not noted" ($detected['noHcs'].Noted -eq 0) $detected['noHcs'].Lines
            Assert-Case "VM logs that could not be read: neither noted nor escalated" `
                ($detected['unread'].Noted -eq 0 -and $detected['unread'].Escalated -eq 0) $detected['unread'].Lines
        }
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Step 9's check when the guest says it is connected"

# When the guest says it is connected, Step 9 reads the VM logs once more
# and takes any marker since the baseline as ready. A VM log that could not
# be read answers 'unreadable' (6.3.1), which is not a marker: until 6.3.1
# the check took any answer at all, so a log it could not read would have
# ended the wait as ready. The clause is lifted from Step 9's wait and run in
# a loop of one pass, which its break leaves, with Log, Start-Sleep and the
# log read faked.
$connClause = @($guestIfs | ForEach-Object { $_.Clauses } |
    Where-Object { $_.Item1.Extent.Text -eq '$guestState -eq "guest-connected"' })
Assert-Case "the guest-connected clause was found" ($connClause.Count -eq 1) "found $($connClause.Count)"
if ($connClause.Count -eq 1) {
    $unfaked = @(Get-UnfakedCommand -Asts @($connClause[0].Item2) -Faked @('Log', 'Start-Sleep', 'Test-VmLogReady'))
    Assert-Case "every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
    if ($unfaked.Count -eq 0) {
        # Item2 is the clause's block, braces and all: the loop's body.
        $connPass = [scriptblock]::Create('foreach ($pass in @(1)) ' + $connClause[0].Item2.Extent.Text)
        $conn = Invoke-Isolated -Name "the guest-connected check" -Block {
            Set-StrictMode -Version Latest
            function Log { param([string]$Message, [string]$Colour, [switch]$Indent) $null = $Colour, $Indent; [void]$script:ConnLines.Add($Message) }
            function Start-Sleep { param([int]$Seconds) $null = $Seconds }
            function Test-VmLogReady { param($Baseline) $null = $Baseline; return $script:ConnLog }
            $o = [ordered]@{}
            foreach ($c in @(
                @{ Name = 'unread'; Log = 'unreadable' },
                @{ Name = 'marker'; Log = 'startup-complete' },
                @{ Name = 'none';   Log = $null })) {
                $script:ConnLines = [System.Collections.ArrayList]::new()
                $script:ConnLog = $c.Log
                $vmReady = $false
                $logBaseline = @{}
                . $connPass
                $o[$c.Name] = [pscustomobject]@{ Ready = $vmReady; Lines = ($script:ConnLines -join ' | ') }
            }
            $o
        }
        if ($conn) {
            Assert-Case "a VM log that could not be read: not ready, and not said to be" `
                (-not $conn['unread'].Ready -and $conn['unread'].Lines -eq '') "ready $($conn['unread'].Ready); said: $($conn['unread'].Lines)"
            Assert-Case "a marker in the log: ready, and said, naming it" `
                ($conn['marker'].Ready -and $conn['marker'].Lines -eq 'Workspace ready (guest connected + log: startup-complete)') `
                "ready $($conn['marker'].Ready); said: $($conn['marker'].Lines)"
            Assert-Case "no marker yet: not ready, and nothing said" `
                (-not $conn['none'].Ready -and $conn['none'].Lines -eq '') "ready $($conn['none'].Ready); said: $($conn['none'].Lines)"
        }
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- which cowork-vm the monitor keeps"

# With more than one cowork-vm compute system the monitor keeps the one
# whose worker process started last and kills the others. The start times
# are compared in UTC: in the hour the clocks go back, a worker started
# later can show an earlier local time, and until 6.3.0 the monitor then
# kept the older VM and killed the live one. Stop-StaleCoworkVm is lifted
# from Watch, where until 6.3.0 this was a block in the main loop, and run
# with the list, the worker list, hcsdiag and the log faked.
$keepDef = Get-AstFunction -Ast $watchAst -Name 'Stop-StaleCoworkVm'
Assert-Case "Stop-StaleCoworkVm, which picks the live VM, was found" ($null -ne $keepDef)
Assert-Case "it reads each worker's start time in UTC" `
    ($watchSrc -match '\$started\[\$wg\] = \(\[datetime\]\$w\.CreationDate\)\.ToUniversalTime\(\)')
# Get-VmWorkerGuid is the shared region's parser of a command line, dot-
# sourced above. It runs no command.
# Test-CheckErrorDue throttles the line for a worker that could not be
# read, from 6.3.1; here every worker is read, so a call to it is a fault.
$keepFaked = @('Get-HcsListCached', 'Get-CoworkHcsGuids', 'Get-CimInstance', 'Invoke-HcsDiag', 'Get-CoworkHcsSurvivors',
    'Write-WatchLog', 'Write-CheckError', 'Test-CheckErrorDue', 'Get-VmWorkerGuid')
if ($keepDef) {
    $unfaked = @(Get-UnfakedCommand -Asts @($keepDef) -Faked $keepFaked)
    Assert-Case "every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
    $guidOld = 'AAAAAAAA-1111-4222-8333-444444444444'
    $guidNew = 'BBBBBBBB-1111-4222-8333-444444444444'
    function New-TestWorker {
        param([string]$Guid, [datetime]$Utc)
        # CIM hands CreationDate back in local time.
        return [pscustomobject]@{ CommandLine = "`"C:\Windows\System32\vmwp.exe`" $Guid 0x330"; CreationDate = $Utc.ToLocalTime() }
    }
    $utc = [System.DateTimeKind]::Utc
    $keepCases = @(@{ Name = 'plain'; Workers = @(
        (New-TestWorker -Guid $guidOld -Utc ([datetime]::SpecifyKind([datetime]'2026-07-01T12:00:00', $utc))),
        (New-TestWorker -Guid $guidNew -Utc ([datetime]::SpecifyKind([datetime]'2026-07-01T12:10:00', $utc)))) })
    # The first minute in the coming year at which this zone's clocks go
    # back. A zone without summer time has none, and there local and UTC
    # order always agree.
    $tz = [System.TimeZoneInfo]::Local
    $back = $null
    $scanFrom = [datetime]::SpecifyKind([datetime]::UtcNow.Date, $utc)
    for ($h = 1; $h -le 24 * 366 -and $null -eq $back; $h++) {
        $t = $scanFrom.AddHours($h)
        if ($tz.GetUtcOffset($t) -lt $tz.GetUtcOffset($t.AddHours(-1))) { $back = $t }
    }
    if ($back) {
        $before = $tz.GetUtcOffset($back.AddHours(-1))
        $m = $back.AddHours(-1)
        while ($tz.GetUtcOffset($m) -ge $before) { $m = $m.AddMinutes(1) }
        $back = $m
        $stepBack = ($before - $tz.GetUtcOffset($back)).TotalMinutes
        # One worker a third of the step before the change, one a third
        # after it: the later one shows the earlier local time.
        $wOld = New-TestWorker -Guid $guidOld -Utc $back.AddMinutes(-$stepBack / 3)
        $wNew = New-TestWorker -Guid $guidNew -Utc $back.AddMinutes($stepBack / 3)
        Assert-Case "the pair is in reverse order by the local clock, and in order in UTC" `
            ($wOld.CreationDate -gt $wNew.CreationDate -and $wOld.CreationDate.ToUniversalTime() -lt $wNew.CreationDate.ToUniversalTime()) `
            "local $($wOld.CreationDate.ToString('s')) and $($wNew.CreationDate.ToString('s'))"
        $keepCases += @{ Name = 'clockBack'; Workers = @($wOld, $wNew) }
    } else {
        Write-Host "  NOTE  this zone's clocks never go back, so the reversed pair cannot be built here; the UTC check above still runs"
    }
    if ($unfaked.Count -eq 0) {
        $kept = Invoke-Isolated -Name 'Stop-StaleCoworkVm' -Block {
            Set-StrictMode -Version Latest
            function Get-HcsListCached { return 'the list' }
            function Get-CoworkHcsGuids { param([string]$ListOutput) $null = $ListOutput; return @($guidOld, $guidNew) }
            function Get-CimInstance { param([string]$ClassName, [string]$Filter, [string]$ErrorAction) $null = $ClassName, $Filter, $ErrorAction; return $script:KeepWorkers }
            function Invoke-HcsDiag { param([string[]]$Arguments) if ($Arguments[0] -eq 'kill') { [void]$script:KeepKills.Add($Arguments[1]) }; return '' }
            function Get-CoworkHcsSurvivors { param([string[]]$Guids, [scriptblock]$List) $null = $Guids, $List; return ,@() }
            function Write-WatchLog { param([string]$Message) $null = $Message }
            function Write-CheckError { param([string]$Check, [string]$Text, $ErrorRecord) $null = $Check, $Text, $ErrorRecord; [void]$script:KeepKills.Add('CHECK ERROR') }
            function Test-CheckErrorDue { param([string]$Check, [string]$Message, $Now) $null = $Check, $Message, $Now; [void]$script:KeepKills.Add('UNREAD'); return $true }
            . ([scriptblock]::Create($keepDef.Extent.Text))
            $o = [ordered]@{}
            foreach ($c in $keepCases) {
                $script:KeepWorkers = @($c.Workers)
                $script:KeepKills = [System.Collections.ArrayList]::new()
                $script:LastHcsListTime = [datetime]::UtcNow
                Stop-StaleCoworkVm
                $o[$c.Name] = ($script:KeepKills -join ',')
            }
            $o
        }
        if ($kept) {
            Assert-Case "two workers in the usual order: the older VM is killed" ($kept['plain'] -eq $guidOld) "killed [$($kept['plain'])]"
            if ($kept.Contains('clockBack')) {
                Assert-Case "either side of the hour the clocks go back: the older VM is killed, not the live one" `
                    ($kept['clockBack'] -eq $guidOld) "killed [$($kept['clockBack'])]"
            }
        }
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- the monitor's heartbeat check, which could never fire"

# It read the Hyper-V heartbeat through Get-VM, which never lists an HCS
# compute system, and from 6.0.0 it was a function that returned nothing.
# It is gone, with its state, and the startup banner still says the
# heartbeat is not monitored.
Assert-Case "no Test-VmHeartbeat, and nothing calls it" `
    ($null -eq (Get-AstFunction -Ast $watchAst -Name 'Test-VmHeartbeat') -and $watchSrc -notmatch 'Test-VmHeartbeat')
Assert-Case "no heartbeat state" ($watchSrc -notmatch 'HeartbeatFailCount|LastHeartbeatStatus')
$watchHelp = [regex]::Match($watchSrc, '(?s)<#(.*?)#>').Groups[1].Value
Assert-Case "the help does not list it as monitored" ($watchHelp -match 'VM log staleness' -and $watchHelp -notmatch 'heartbeat')
Assert-Case "the startup banner says it is not monitored" `
    ($watchSrc -match '(?m)^Write-WatchLog "  Not monitored : Hyper-V heartbeat \(no such thing for an HCS compute system\)"')

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- what Step 5 says it found"

# Step 5's body, lifted from Fix and run with Close-StaleHcsVms, the worker
# check and the log faked. "No orphan compute systems found" is green, and is
# printed only when hcsdiag's list answered and the worker check finished,
# with nothing found. Until 6.3.0 it was printed with no admin, with hcsdiag
# not answering, when either check failed, and when compute systems were
# listed and no kill went out. -Close's step 3, run the same way, says that
# last case too. When kills went out for some of those listed and not for
# the rest, both say how many are left, and Step 5 counts them as left; until
# 6.3.0 neither said so.
$step5Try = Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.TryStatementAst] -and
    $n.Body.Statements.Count -gt 0 -and $n.Body.Statements[0].Extent.Text -eq '$orphanKilled = $false' }
$close3If = Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and
    $n.Clauses[0].Item1.Extent.Text -eq '$script:IsAdmin' -and $n.Extent.Text -match 'remaining compute system\(s\)' }
Assert-Case "Step 5's body and -Close's step 3 were found" ($step5Try.Count -eq 1 -and $close3If.Count -eq 1) `
    "found $($step5Try.Count) and $($close3If.Count)"
$step5 = $null
if ($step5Try.Count -eq 1 -and $close3If.Count -eq 1) {
    $unfaked = @(Get-UnfakedCommand -Asts (@($step5Try) + @($close3If)) -Faked @('Log', 'Close-StaleHcsVms', 'Stop-HungCoworkWorker', 'Start-Sleep'))
    Assert-Case "every command they call is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
    if ($unfaked.Count -eq 0) {
        $step5 = Invoke-Isolated -Name "Step 5's body" -Block {
            Set-StrictMode -Version Latest
            function Log { param([string]$Message, [string]$Colour, [switch]$Indent) $null = $Indent; [void]$script:Step5Lines.Add("${Colour}: $Message") }
            function Start-Sleep { param([int]$Seconds) $null = $Seconds }
            # Sets what the real one records: whether the list answered, how
            # many Cowork compute systems it showed, and how many kills went
            # out. One that throws leaves what an earlier call set.
            function Close-StaleHcsVms {
                param([string]$Action)
                $null = $Action
                $h = $script:Hcs
                if ($h.Throws) { throw 'hcsdiag blew up' }
                $script:LastHcsListOk    = $h.Ok
                $script:LastHcsListed    = $h.Listed
                $script:LastHcsKillsSent = $h.Sent
                return $h.Killed
            }
            function Stop-HungCoworkWorker {
                if ($script:Worker.Throws) { throw 'no worker list' }
                return [pscustomobject]@{ Stopped = $script:Worker.Stopped; Left = $script:Worker.Left }
            }
            $step5Block  = [scriptblock]::Create($step5Try[0].Extent.Text)
            $close3Block = [scriptblock]::Create($close3If[0].Extent.Text)
            $clean = @{ Ok = $true; Listed = 0; Sent = 0; Killed = 0; Throws = $false }
            $idle  = @{ Stopped = 0; Left = 0; Throws = $false }
            $cases = @(
                @{ Name = 'adminClean';    Admin = $true;  Hcs = $clean; Worker = $idle },
                @{ Name = 'adminNoAnswer'; Admin = $true;  Hcs = @{ Ok = $false; Listed = 0; Sent = 0; Killed = 0; Throws = $false }; Worker = $idle },
                @{ Name = 'notAdmin';      Admin = $false; Hcs = $clean; Worker = $idle },
                @{ Name = 'hcsThrows';     Admin = $true;  Hcs = @{ Ok = $true; Listed = 0; Sent = 0; Killed = 0; Throws = $true }; Worker = $idle },
                @{ Name = 'workerThrows';  Admin = $true;  Hcs = $clean; Worker = @{ Stopped = 0; Left = 0; Throws = $true } },
                @{ Name = 'listedNoKill';  Admin = $true;  Hcs = @{ Ok = $true; Listed = 2; Sent = 0; Killed = 0; Throws = $false }; Worker = $idle },
                @{ Name = 'killed';        Admin = $true;  Hcs = @{ Ok = $true; Listed = 1; Sent = 1; Killed = 1; Throws = $false }; Worker = $idle },
                @{ Name = 'workerLeft';    Admin = $true;  Hcs = $clean; Worker = @{ Stopped = 0; Left = 1; Throws = $false } },
                @{ Name = 'partial';       Admin = $true;  Hcs = @{ Ok = $true; Listed = 2; Sent = 1; Killed = 1; Throws = $false }; Worker = $idle },
                @{ Name = 'partialNone';   Admin = $true;  Hcs = @{ Ok = $true; Listed = 3; Sent = 2; Killed = 0; Throws = $false }; Worker = $idle }
            )
            $o = [ordered]@{}
            foreach ($c in $cases) {
                $script:IsAdmin = $c.Admin
                $script:Hcs     = $c.Hcs
                $script:Worker  = $c.Worker
                # What Step 0's list left: it answered.
                $script:LastHcsListOk = $true; $script:LastHcsListed = 0; $script:LastHcsKillsSent = 0
                $script:Step5Lines = [System.Collections.ArrayList]::new()
                . $step5Block
                $o[$c.Name] = ($script:Step5Lines -join ' | ')
                $o["$($c.Name)_left"] = "$orphanLeft"
                $script:LastHcsListOk = $true; $script:LastHcsListed = 0; $script:LastHcsKillsSent = 0
                $script:Step5Lines = [System.Collections.ArrayList]::new()
                . $close3Block
                $o["close_$($c.Name)"] = ($script:Step5Lines -join ' | ')
            }
            $o
        }
    }
}
if ($null -eq $step5) { $step5 = [ordered]@{} }
Assert-Case "admin, the list answered, nothing found: the green line" `
    ($step5['adminClean'] -eq 'DarkGray: No orphan compute systems via hcsdiag | Green: No orphan compute systems found') "said: $($step5['adminClean'])"
Assert-Case "hcsdiag did not answer: no green line; the compute systems were not checked" `
    ($step5['adminNoAnswer'] -eq 'DarkGray: hcsdiag did not answer, so this check was not run | DarkYellow: Nothing was stopped, but compute systems were not checked (hcsdiag did not answer)') `
    "said: $($step5['adminNoAnswer'])"
Assert-Case "no admin: no green line; not checked, and why" `
    ($step5['notAdmin'] -eq 'DarkYellow: Nothing was stopped, but compute systems were not checked (needs admin)') "said: $($step5['notAdmin'])"
Assert-Case "the hcsdiag query failed: no green line" `
    ("$($step5['hcsThrows'])" -notmatch 'Green: No orphan' -and "$($step5['hcsThrows'])" -match 'DarkYellow: Nothing was stopped, but compute systems were not checked \(the hcsdiag query failed\)$') `
    "said: $($step5['hcsThrows'])"
Assert-Case "the worker check failed: no green line, and says so" `
    ("$($step5['workerThrows'])" -notmatch 'Green: No orphan' -and "$($step5['workerThrows'])" -match '\| DarkYellow: Nothing was stopped, but the VM worker check did not finish$') `
    "said: $($step5['workerThrows'])"
Assert-Case "compute systems listed and no kill sent: no green line, and says so" `
    ($step5['listedNoKill'] -eq 'DarkYellow: 2 orphan compute system(s) listed, none killed') "said: $($step5['listedNoKill'])"
Assert-Case "one killed: says so, and no clean-result line" `
    ($step5['killed'] -eq 'Green: Killed 1 orphan compute system(s)') "said: $($step5['killed'])"
Assert-Case "a worker left running: no clean-result line" `
    ("$($step5['workerLeft'])" -eq 'DarkGray: No orphan compute systems via hcsdiag') "said: $($step5['workerLeft'])"
Assert-Case "-Close: listed and no kill sent says so, not 'No remaining compute systems'" `
    ($step5['close_listedNoKill'] -eq 'DarkYellow: 2 compute system(s) listed, none killed') "said: $($step5['close_listedNoKill'])"
Assert-Case "-Close: an empty list is still 'No remaining compute systems'" `
    ($step5['close_adminClean'] -eq 'DarkGray: No remaining compute systems') "said: $($step5['close_adminClean'])"
Assert-Case "one of two killed, and no kill sent for the other: says the other is left" `
    ($step5['partial'] -eq 'Green: Killed 1 orphan compute system(s) | DarkYellow: 1 orphan compute system(s) listed, not killed') `
    "said: $($step5['partial'])"
Assert-Case "and counts it as left; one of one killed is not" ($step5['partial_left'] -eq 'True' -and $step5['killed_left'] -eq 'False') `
    "left: $($step5['partial_left']), one of one killed: $($step5['killed_left'])"
Assert-Case "two kills of three sent, none confirmed: says so, and that the third is left" `
    ($step5['partialNone'] -eq 'DarkYellow: No orphan compute system confirmed gone | DarkYellow: 1 orphan compute system(s) listed, not killed') `
    "said: $($step5['partialNone'])"
Assert-Case "-Close: one of two killed says the other is left" `
    ($step5['close_partial'] -eq 'Green: Killed 1 remaining compute system(s) | DarkYellow: 1 compute system(s) listed, not killed') `
    "said: $($step5['close_partial'])"
Assert-Case "-Close: two kills of three sent, none confirmed, says the third is left" `
    ($step5['close_partialNone'] -eq 'DarkYellow: No compute system confirmed gone | DarkYellow: 1 compute system(s) listed, not killed') `
    "said: $($step5['close_partialNone'])"
Assert-Case "-Close: every one listed killed says nothing more" ($step5['close_killed'] -eq 'Green: Killed 1 remaining compute system(s)') `
    "said: $($step5['close_killed'])"

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- the monitor: a repair Fix held back is retried when the hold clears"

# Invoke-HeldBackRetry and Invoke-AutoFix, lifted from Watch as above, with
# the real Save-WatchState and Import-WatchState on a temp file. Time passing
# is every time they keep moved back by that many minutes, the app's start
# times and the Fix logs' times included. The Claude Code checks, the
# user-activity check and the Cowork check answer what each case sets, and
# the stand-in Fix counts its runs and exits with the code the case sets. A
# one-shot trigger's repair that was held back is retried once the hold has
# cleared, for up to 4 hours; until 6.3.0 nothing retried it.
#
# The helpers the 6.3.0 rules live in are lifted when they are there, so on
# a tree without them this section still runs, and each rule it lacks fails
# as a case of its own.
$retryCore = @('Invoke-HeldBackRetry', 'Clear-HeldBackRepair', 'Invoke-AutoFix', 'Test-PersistentFailure',
    'Import-WatchState', 'Save-WatchState', 'ConvertFrom-WatchStateTime')
$retryHelpers = @('Close-HeldBackRepair', 'Register-HeldBackRepair', 'Get-HeldBackRetryWait', 'Get-ClaudeAppStartTime',
    'Get-FixRunAfter', 'Close-StaleHeldBackRepair', 'Test-FixLockHeld', 'Write-FixRunningSkip', 'Merge-HeldBackTrigger',
    'Get-ClaudeCodeSessionHold', 'Test-LineDue', 'Get-CheckErrorText', 'Test-CheckErrorDue', 'Write-CheckError', 'Get-ProcessStartUtc',
    'Test-ClaudeNotFoundError', 'Get-ClaudeItemOrNull', 'Get-ClaudeErrorText')
$retryNames = $retryCore + $retryHelpers
$retryDefs = @($retryNames | ForEach-Object { Get-AstFunction -Ast $watchAst -Name $_ } | Where-Object { $_ })
$retryDefNames = @($retryDefs | ForEach-Object { $_.Name })
$retryFound = (@($retryCore | Where-Object { $retryDefNames -notcontains $_ }).Count -eq 0)
Assert-Case "Invoke-HeldBackRetry and what it calls were found" ($retryDefs.Count -eq $retryNames.Count) `
    "missing: $(@($retryNames | Where-Object { $retryDefNames -notcontains $_ }) -join ', ')"
$retry = $null
$retryLockName = 'Local\ClaudeFixTest_' + [guid]::NewGuid().ToString('N')
if ($retryFound) {
    $retryTexts = @{}
    foreach ($d in $retryDefs) { $retryTexts[$d.Name] = $d.Extent.Text }
    # Fix's lock names: in Invoke-AutoFix until 6.3.0, in Test-FixLockHeld
    # since. Swapped once, wherever they are.
    $lockSwaps = 0
    foreach ($k in @($retryTexts.Keys)) {
        $lockSwaps += ([regex]::Matches($retryTexts[$k], [regex]::Escape($fixLockLiteral))).Count
        $retryTexts[$k] = $retryTexts[$k].Replace($fixLockLiteral, '@("' + $retryLockName + '")')
    }
    $t = $retryTexts['Invoke-AutoFix']
    $adminSwaps = ([regex]::Matches($t, [regex]::Escape('if (-not $watchIsAdmin) {'))).Count
    $t = $t.Replace('if (-not $watchIsAdmin) {', 'if ($script:NotElevated) {')
    $retryTexts['Invoke-AutoFix'] = $t
    $retrySafe = ($lockSwaps -eq 1 -and $adminSwaps -eq 1 -and
        @($retryTexts.Values | Where-Object { $_ -match '"(Global|Local)\\ClaudeDesktopFix"' }).Count -eq 0)
    Assert-Case "the lock names and the elevation check were swapped" $retrySafe "lock $lockSwaps, elevation $adminSwaps"
    # Get-Item, Test-Path, Get-Content, ConvertFrom-Json and ConvertTo-Json
    # are the real ones, on the temp file.
    $retryFakes = @('Write-WatchLog', 'Add-Type', 'New-Object', 'Start-Sleep', 'Get-Process', 'Get-CimInstance', 'Get-ClaudeProcessByName',
        'Get-ClaudeDesktopProcess', 'Get-ClaudeVmLogAge', 'Test-UserActivity', 'Test-ClaudeCodeActive',
        'Test-CoworkSessionActive', 'Get-ClaudeAppCliProcess', 'Get-ChildItem', 'Test-Path', 'Get-Content', 'Get-Item',
        'ConvertFrom-Json', 'ConvertTo-Json', 'Test-ClaudeCodeTranscriptRecent', 'Test-FixLogRepaired') + $retryNames
    $unfaked = @(Get-UnfakedCommand -Asts $retryDefs -Faked $retryFakes -AllowDynamic @('$fixScript'))
    Assert-Case "every command they call is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
    if ($retrySafe -and $unfaked.Count -eq 0) {
        $retryFix = Join-Path $tempRoot 'fake-fix-retry.ps1'
        # Holds the run back, as Fix's -Quiet gate does, when the case says
        # input within 3 minutes does. With RetryFixWritesLog it writes a log
        # first, created 20 seconds before it returns, as a real run does. The
        # log holds the line Fix logs once it begins the repair, which a run
        # held back would not, so that only its times keep it from counting.
        $retryFixText = @(
            'param([string]$Mode, [switch]$Quiet)'
            '$null = $Mode, $Quiet'
            '$global:RetryFixRuns++'
            'if ($global:RetryFixWritesLog) { $n = [datetime]::UtcNow; $null = $global:RetryFixLogs.Add(@{ C = $n.AddSeconds(-20); W = $n; Text = $global:RetryRepairLine }) }'
            "if (`$global:RetryFixThrows) { throw 'could not start' }"
            'if ($global:RetryHoldOnInput -and $global:RetryInputAgoMin -lt 3) { exit 3 }'
            'exit $global:RetryFixExit'
        ) -join "`r`n"
        [System.IO.File]::WriteAllText($retryFix, $retryFixText + "`r`n")
        $retry = Invoke-Isolated -Name 'the held-back retry' -Block {
            Set-StrictMode -Version Latest
            function Write-WatchLog { param([string]$Message) [void]$script:WatchLines.Add($Message) }
            function Add-Type { }
            # Save-WatchState's encoding is real; a balloon is not.
            function New-Object {
                param([string]$TypeName, [object[]]$ArgumentList)
                $null = $ArgumentList
                if ($TypeName -eq 'System.Text.UTF8Encoding') { return [System.Text.UTF8Encoding]::new($false) }
                throw 'no desktop here'
            }
            # The 30-second warning: noted, and the minutes the case sets
            # pass. A user it disturbs (RetryTouch) has just used the mouse.
            function Start-Sleep {
                param([int]$Seconds, [int]$Milliseconds)
                $null = $Milliseconds
                if ($Seconds -eq 30) {
                    [void]$script:WarnAt.Add([math]::Round($script:Elapsed, 1))
                    if ($global:RetryTouch) { $global:RetryInputAgoMin = 0 }
                    if ($script:WarnMinutes -gt 0) { Step-Clock $script:WarnMinutes }
                }
            }
            function Get-Process { return @() }
            function Get-ClaudeProcessByName { param([string]$Name) $null = $Name; return }
            function Get-CimInstance { return @() }
            # The app's processes, started at the UTC times the case sets;
            # 'THROW' is one whose start time cannot be read, as for one that
            # has just exited.
            function Get-ClaudeDesktopProcess {
                param($Environment)
                $null = $Environment
                $list = @()
                foreach ($s in @($script:DesktopStarts)) {
                    $p = [pscustomobject]@{ Id = $list.Count + 1 }
                    if ($s -is [datetime]) {
                        $p | Add-Member -MemberType NoteProperty -Name StartTime -Value $s.ToLocalTime()
                    } else {
                        $p | Add-Member -MemberType ScriptProperty -Name StartTime -Value { throw 'the process has exited' }
                        $p | Add-Member -MemberType NoteProperty -Name HasExited -Value $true
                    }
                    $list += $p
                }
                return $list
            }
            # Fix's logs, created and written at the UTC times the case sets,
            # each one's text standing in for its path. LogStep: the minutes
            # that pass while they are listed, once.
            function Get-ChildItem {
                param($LiteralPath, $Filter, [switch]$File, $ErrorAction)
                $null = $LiteralPath, $Filter, $File, $ErrorAction
                if ($script:LogStep -gt 0) {
                    $s = $script:LogStep
                    $script:LogStep = 0
                    Step-Clock $s
                }
                return @($script:FixLogs | ForEach-Object { [pscustomobject]@{ CreationTimeUtc = $_.C; LastWriteTimeUtc = $_.W; FullName = [string]$_['Text'] } })
            }
            # Whether a Fix log holds the line Fix logs once it begins the
            # repair, read from the text the case sets. The real function,
            # which opens the file, is tested on real files below.
            function Test-FixLogRepaired { param([string]$Path) return $Path.Contains($FixRepairLine) }
            function Test-ClaudeCodeTranscriptRecent {
                param([int]$Minutes)
                $null = $Minutes
                $script:TranscriptAsks++
                return [bool]$script:TranscriptRecent
            }
            function Get-ClaudeVmLogAge { param($LogDir) $null = $LogDir; return $null }
            # Asked with -Silent by the retry's quiet check, and without it by
            # BLOCKED. UserRace: the user arrives between the two. With
            # RetryUseInput, input within 3 minutes is activity.
            function Test-UserActivity {
                param([switch]$Silent)
                if ($Silent) { $script:SilentUserAsks++ } else { $script:LoudUserAsks++ }
                if ($script:UserRace) { return (-not $Silent) }
                if ($global:RetryUseInput) { return ($global:RetryInputAgoMin -lt 3) }
                return [bool]$script:UserActive
            }
            function Test-ClaudeCodeActive {
                param([switch]$Silent)
                if ($Silent) { $script:SilentCodeAsks++ } else { $script:LoudCodeAsks++ }
                return [bool]$script:CodeActive
            }
            function Test-CoworkSessionActive { return [bool]$script:CoworkActive }
            function Get-ClaudeAppCliProcess {
                param($Environment)
                $null = $Environment
                $script:CliAsks++
                if ("$script:AppCli" -eq 'THROW') { throw 'the parents could not be read' }
                $n = [int]$script:AppCli
                if ($n -eq 0) { return @() }
                return @(1..$n | ForEach-Object { [pscustomobject]@{ Id = $_ } })
            }
            foreach ($n in $retryNames) {
                if ($retryTexts.ContainsKey($n)) { . ([scriptblock]::Create($retryTexts[$n])) }
            }

            if ($watchCodeMatch.Success) { $FixHeldBackExitCode = [int]$watchCodeMatch.Groups[1].Value }
            if ($watchPartMatch.Success) { $FixStoppedPartWayExitCode = [int]$watchPartMatch.Groups[1].Value }
            $Cooldown    = 5
            $ProcessName = 'claude'
            $fixScript   = $retryFix
            $FixLogDir   = Join-Path $tempRoot 'no-fix-logs'
            $FixRepairLine = if ($watchRepairMatch.Success) { $watchRepairMatch.Groups[1].Value } else { 'Repair under way' }
            $script:ClaudeEnv            = [pscustomobject]@{ ProcessName = 'claude' }
            $script:ClaudeLogDir         = 'C:\nowhere'
            $script:IsInteractiveSession = $false
            $script:Win32ActivityLoaded  = $false
            $script:StatePath            = Join-Path $tempRoot 'watch-state-retry.json'
            $script:LastStealthMarkerTs  = [long]0
            $script:LogBaselines         = @{}
            $script:VmLogStaleCount      = 0
            $script:VmLogEverActive      = $false
            $script:ServiceDownCount     = 0
            $script:EventLogHitCount     = 0
            $script:VmcomputeLeakCount   = 0
            function Reset-Retry {
                $script:FixHistory   = [System.Collections.ArrayList]::new()
                $script:FixCount     = 0
                $script:LastFixTime  = [datetime]::MinValue
                $script:WatchLines   = [System.Collections.ArrayList]::new()
                $script:CodeActive   = $false
                $script:AppCli       = 0
                $script:CliAsks      = 0
                $script:UserActive   = $false
                $script:UserRace     = $false
                $script:CoworkActive = $false
                $script:NotElevated  = $false
                $script:SilentUserAsks = 0
                $script:LoudUserAsks   = 0
                $script:SilentCodeAsks = 0
                $script:LoudCodeAsks   = 0
                $script:DesktopStarts  = [System.Collections.ArrayList]::new()
                $script:FixLogs        = [System.Collections.ArrayList]::new()
                $global:RetryFixLogs   = $script:FixLogs
                $script:WarnAt         = [System.Collections.ArrayList]::new()
                $script:WarnMinutes    = 0
                $script:Elapsed        = 0.0
                $global:RetryFixRuns     = 0
                $global:RetryFixExit     = 3
                $global:RetryFixThrows   = $false
                $global:RetryTouch       = $false
                $global:RetryUseInput    = $false
                $global:RetryHoldOnInput = $false
                $global:RetryFixWritesLog = $false
                $global:RetryInputAgoMin = 60.0
                $script:PendingFixReason      = $null
                $script:PendingFixSince       = [datetime]::MinValue
                $script:PendingFixHeldAt      = [datetime]::MinValue
                $script:PendingFixTries       = 0
                $script:PendingFixRetriesHeld = 0
                $script:PendingFixNextCheck   = [datetime]::MinValue
                $script:PendingFixWaitLogged  = [datetime]::MinValue
                $script:PendingFixLockLogged  = $false
                $script:FixHeldInARow         = 0
                $script:FixHeldLastTime       = [datetime]::MinValue
                $script:CooldownFromHold      = $false
                $script:RepeatLinesSaid       = @{}
                $script:CheckErrorsSaid       = @{}
                $script:TranscriptRecent      = $false
                $script:TranscriptAsks        = 0
                $script:LogStep               = 0
                $global:RetryRepairLine       = "[10:00:05] $FixRepairLine"
                # No case reads what an earlier one saved.
                [System.IO.File]::Delete($script:StatePath)
            }
            function Step-Clock {
                param([double]$Minutes)
                if ($script:LastFixTime -ne [datetime]::MinValue) { $script:LastFixTime = $script:LastFixTime.AddMinutes(-$Minutes) }
                if ($script:PendingFixSince -ne [datetime]::MinValue) { $script:PendingFixSince = $script:PendingFixSince.AddMinutes(-$Minutes) }
                if ($script:PendingFixHeldAt -ne [datetime]::MinValue) { $script:PendingFixHeldAt = $script:PendingFixHeldAt.AddMinutes(-$Minutes) }
                if ($script:PendingFixNextCheck -ne [datetime]::MinValue) { $script:PendingFixNextCheck = $script:PendingFixNextCheck.AddMinutes(-$Minutes) }
                if ($script:PendingFixWaitLogged -ne [datetime]::MinValue) { $script:PendingFixWaitLogged = $script:PendingFixWaitLogged.AddMinutes(-$Minutes) }
                if ($script:FixHeldLastTime -ne [datetime]::MinValue) { $script:FixHeldLastTime = $script:FixHeldLastTime.AddMinutes(-$Minutes) }
                foreach ($k in @($script:RepeatLinesSaid.Keys)) { $script:RepeatLinesSaid[$k] = $script:RepeatLinesSaid[$k].AddMinutes(-$Minutes) }
                for ($i = 0; $i -lt $script:DesktopStarts.Count; $i++) {
                    if ($script:DesktopStarts[$i] -is [datetime]) { $script:DesktopStarts[$i] = $script:DesktopStarts[$i].AddMinutes(-$Minutes) }
                }
                foreach ($f in $script:FixLogs) { $f.C = $f.C.AddMinutes(-$Minutes); $f.W = $f.W.AddMinutes(-$Minutes) }
                for ($i = 0; $i -lt $script:FixHistory.Count; $i++) { $script:FixHistory[$i] = $script:FixHistory[$i].AddMinutes(-$Minutes) }
                $global:RetryInputAgoMin += $Minutes
                $script:Elapsed += $Minutes
            }
            function Get-Lines { param([string]$Pattern) return @($script:WatchLines | Where-Object { $_ -match $Pattern }).Count }
            # The last line written, or '' when there is none. Indexed, an
            # empty list throws under StrictMode and ends every case after it.
            function Get-LastLine { return "$(@($script:WatchLines) | Select-Object -Last 1)" }
            function Get-SavedField {
                param([string]$Name)
                if (-not [System.IO.File]::Exists($script:StatePath)) { return '<no file>' }
                $saved = [System.IO.File]::ReadAllText($script:StatePath) | ConvertFrom-Json
                if (-not $saved.PSObject.Properties[$Name]) { return '<absent>' }
                return "$($saved.$Name)"
            }
            function Get-SavedPending { return (Get-SavedField -Name 'PendingFixReason') }
            # A one-shot trigger, as the log-error and Event Log checks call
            # it. A tree without -OneShot gets the call it has, and fails on
            # what it does with it.
            $oneShot = (Get-Command Invoke-AutoFix).Parameters.ContainsKey('OneShot')
            function Invoke-OneShot {
                param([string]$Reason)
                if ($oneShot) { Invoke-AutoFix -Reason $Reason -OneShot } else { Invoke-AutoFix -Reason $Reason }
            }
            function Test-JustNow { param([datetime]$Time) return (([datetime]::UtcNow - $Time).TotalMinutes -lt 1) }
            # Whole minutes since the pending repair's 4 hours started.
            function Get-SinceMinutes { return [math]::Round(([datetime]::UtcNow - $script:PendingFixSince).TotalMinutes) }
            $waitLine = '^a held-back repair for test trigger is still waiting: '
            $o = [ordered]@{}
            $o['takesOneShot'] = $oneShot

            # Held back: pending, and not asked about before the cooldown it
            # leaves is over.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $o['heldPending'] = "$($script:PendingFixReason)|$($global:RetryFixRuns)|$(Get-SavedPending)"
            $script:WatchLines.Clear()
            $script:LoudCodeAsks = 0
            # The trigger's own BLOCKED check asked once.
            $script:CliAsks = 0
            $script:CodeActive = $true
            $null = Invoke-HeldBackRetry
            $o['inCooldownAsks'] = $script:CliAsks

            # Still waiting for a Code session: no warning, no Fix, one line
            # in 10 minutes, asked no more than every 2 minutes, and asked
            # silently.
            Step-Clock 2.1
            $null = Invoke-HeldBackRetry
            $asks = $script:CliAsks
            $null = Invoke-HeldBackRetry
            $o['askedAgainAtOnce'] = $script:CliAsks - $asks
            for ($i = 0; $i -lt 4; $i++) { Step-Clock 2.1; $null = Invoke-HeldBackRetry }
            $o['waitAsks']    = $script:CliAsks
            $o['waitLines10'] = Get-Lines $waitLine
            Step-Clock 2.1
            $null = Invoke-HeldBackRetry
            $o['waitLines12'] = Get-Lines $waitLine
            $o['waitText']    = "$(@($script:WatchLines | Where-Object { $_ -match $waitLine }) | Select-Object -First 1)"
            $script:AppCli = 2
            Step-Clock 10.5
            $null = Invoke-HeldBackRetry
            $o['waitCli'] = Get-LastLine
            $script:AppCli = 'THROW'
            Step-Clock 10.5
            $null = Invoke-HeldBackRetry
            $o['waitCliThrow'] = Get-LastLine
            $script:AppCli = 0
            $o['waitQuiet'] = "$(Get-Lines 'PRE-FIX WARNING')|$($global:RetryFixRuns)|$($script:PendingFixReason)|$($script:LoudCodeAsks)"

            # The session gone: retried once, under the reason it was held
            # back for, and the repair that ran ends it.
            $script:CodeActive = $false
            $global:RetryFixExit = 0
            Step-Clock 2.1
            $null = Invoke-HeldBackRetry
            $o['retryRuns']      = $global:RetryFixRuns
            $o['retryWarning']   = Get-Lines '^>>> PRE-FIX WARNING: retry of held-back repair: test trigger, auto-fix in 30s'
            $o['retryTriggered'] = Get-Lines '^>>> AUTO-FIX #\d+ TRIGGERED: retry of held-back repair: test trigger$'
            $o['retryEnded']     = "$($null -eq $script:PendingFixReason)|$($script:FixHistory.Count)|[$(Get-SavedPending)]"
            Step-Clock 10
            $null = Invoke-HeldBackRetry
            $o['retryOnce'] = $global:RetryFixRuns

            # Past 4 hours: dropped, said once, and no Fix, though the
            # session has gone.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $script:WatchLines.Clear()
            Step-Clock 241
            $null = Invoke-HeldBackRetry
            Step-Clock 10
            $null = Invoke-HeldBackRetry
            $o['dropLines'] = Get-Lines '^>>> held-back repair for test trigger dropped after 4 hours; run Fix-ClaudeDesktop\.ps1 when you are ready$'
            $o['dropEnded'] = "$($null -eq $script:PendingFixReason)|$($global:RetryFixRuns)|$(Get-Lines 'PRE-FIX WARNING')|[$(Get-SavedPending)]"
            # Under 4 hours it is not.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $script:CodeActive = $true
            Step-Clock 239
            $null = Invoke-HeldBackRetry
            $o['notDropped'] = "$($script:PendingFixReason)|$(Get-Lines 'dropped after 4 hours')"
            # The 4 hours end while it waits out the doubled wait after a
            # hold: dropped on the first poll past them, not once the wait is
            # over.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            Step-Clock 239.5
            $script:PendingFixHeldAt      = [datetime]::UtcNow
            $script:PendingFixRetriesHeld = 3
            Step-Clock 1
            $null = Invoke-HeldBackRetry
            $o['dropInWait'] = "$(Get-Lines '^>>> held-back repair for test trigger dropped after 4 hours')|[$($script:PendingFixReason)]|$($global:RetryFixRuns)"
            # Past 4 hours, and a one-shot trigger of its own fires: the old
            # repair is over, and is said to be, and the trigger's own hold
            # starts 4 hours of its own, so it is not dropped on the next poll.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            Step-Clock 241
            $null = Invoke-OneShot 'second trigger'
            $script:CodeActive = $true
            $null = Invoke-HeldBackRetry
            $o['expiredHeld'] = "$($script:PendingFixReason)|$(Test-JustNow $script:PendingFixSince)|" +
                "$(Get-Lines '^>>> held-back repair for test trigger dropped after 4 hours')|$($global:RetryFixRuns)"

            # Held back again on the retry: the first-held time and the
            # reason are kept, with no prefix stacked, and it is still not an
            # attempt. It is a try, and a retry held back, and the latest hold.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            Step-Clock 2.1
            $first = $script:PendingFixSince
            $null = Invoke-HeldBackRetry
            $o['repeat'] = "$($global:RetryFixRuns)|$($script:PendingFixReason)|$($script:PendingFixSince -eq $first)|$($script:FixHistory.Count)|" +
                "$($script:PendingFixTries)|$($script:PendingFixRetriesHeld)|$(Test-JustNow $script:PendingFixHeldAt)"

            # A one-shot trigger that fires on its own takes the repair over:
            # in the cooldown, and when Fix holds it back again. It is new
            # evidence, so the 4 hours start again from it. The retry is then
            # for it.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            Step-Clock 1
            $null = Invoke-OneShot 'second trigger'
            $o['freshCooldown'] = "$($script:PendingFixReason)|$(Get-SinceMinutes)|$($global:RetryFixRuns)|$(Get-SavedPending)"
            # And it says so, after COOLDOWN, naming the repair it took over,
            # as a trigger kept as a new one says it is kept.
            $o['freshCooldownLines'] = "$(Get-Lines '^COOLDOWN: Skipping fix \(second trigger\)')|" +
                "$(Get-Lines '^    It takes over the held-back repair for test trigger, and is tried again once the wait is over, for up to 4 hours$')|" +
                "$(Get-Lines '^    It is tried again once the wait is over')"
            Step-Clock 3
            $null = Invoke-OneShot 'third trigger'
            $o['freshHeld'] = "$($script:PendingFixReason)|$(Get-SinceMinutes)|$($global:RetryFixRuns)"
            # The second run held back in a row: the next repair waits 4
            # minutes.
            Step-Clock 4.1
            $script:WatchLines.Clear()
            $global:RetryFixExit = 0
            $null = Invoke-HeldBackRetry
            $o['freshRetry'] = Get-Lines '^>>> AUTO-FIX #\d+ TRIGGERED: retry of held-back repair: third trigger$'
            # That line is written with COOLDOWN or not at all, so at most
            # once in 10 minutes for the same trigger. One that fires again
            # in the same wait takes nothing over and is not said again.
            # Another trigger is said. And one that takes the repair back
            # from that other trigger in the same wait is a takeover, but its
            # COOLDOWN is not due, so neither is its line.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            Step-Clock 0.5
            $null = Invoke-OneShot 'second trigger'
            Step-Clock 0.5
            $null = Invoke-OneShot 'second trigger'
            Step-Clock 0.5
            $null = Invoke-OneShot 'third trigger'
            Step-Clock 0.3
            $null = Invoke-OneShot 'second trigger'
            $o['freshCooldownAgain'] = "$(Get-Lines '^COOLDOWN: Skipping fix \((second|third) trigger\)')|$(Get-Lines '^    It takes over the held-back repair for ')|" +
                "$(Get-Lines '^    It takes over the held-back repair for second trigger, ')|$($script:PendingFixReason)|$($global:RetryFixRuns)"
            # A pending repair past its 4 hours is dropped instead, and the
            # trigger is kept as a new one: said so, and not as a takeover.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            Step-Clock 1
            $script:PendingFixSince = $script:PendingFixSince.AddMinutes(-240)
            $null = Invoke-OneShot 'second trigger'
            $o['freshCooldownStale'] = "$(Get-Lines '^>>> held-back repair for test trigger dropped after 4 hours')|$(Get-Lines '^    It takes over the held-back repair')|" +
                "$(Get-Lines '^    It is tried again once the wait is over, for up to 4 hours$')|$($script:PendingFixReason)|$(Get-SinceMinutes)"

            # A retry that stops part-way, one that finishes and one that
            # cannot start each end it; so does a trigger of its own that runs.
            foreach ($end in @(@{ Name = 'end4'; Exit = 4; Throws = $false }, @{ Name = 'end0'; Exit = 0; Throws = $false },
                               @{ Name = 'endThrow'; Exit = 0; Throws = $true })) {
                Reset-Retry
                $null = Invoke-OneShot 'test trigger'
                Step-Clock 2.1
                $global:RetryFixExit   = $end.Exit
                $global:RetryFixThrows = $end.Throws
                $null = Invoke-HeldBackRetry
                $o[$end.Name] = "$($null -eq $script:PendingFixReason)|$($global:RetryFixRuns)|$($script:FixHistory.Count)"
            }
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            Step-Clock 3
            $global:RetryFixExit = 0
            $null = Invoke-OneShot 'second trigger'
            $o['endFresh'] = "$($null -eq $script:PendingFixReason)|$($global:RetryFixRuns)"

            # The warning cancelled: it ends, and Fix does not run.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            Step-Clock 2.1
            $script:CoworkActive = $true
            $null = Invoke-HeldBackRetry
            $o['cancel'] = "$($null -eq $script:PendingFixReason)|$($global:RetryFixRuns)|$(Get-Lines '^>>> CANCELLED: Cowork session active during 30s grace period$')|[$(Get-SavedPending)]"

            # The user arrives between the quiet check and BLOCKED: it stays
            # pending with its first-held time, and waits from this hold.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            Step-Clock 2.1
            $first = $script:PendingFixSince
            $script:UserRace = $true
            $script:WatchLines.Clear()
            $null = Invoke-HeldBackRetry
            $o['blockedRace'] = "$($script:PendingFixReason)|$($global:RetryFixRuns)|$(Get-Lines '^>>> BLOCKED: retry of held-back repair: test trigger, user is active')|" +
                "$($script:PendingFixSince -eq $first)|$(Test-JustNow $script:PendingFixHeldAt)|" +
                "$(Get-Lines '^    It is tried again once the user is away, for up to 4 hours, or run Fix-ClaudeDesktop\.ps1 manually when ready$')"

            # The user at the PC: the quiet check asks what BLOCKED asks,
            # silently, so the repair waits with one line in 10 minutes, and
            # no warning and no BLOCKED pair every 2 minutes.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $script:UserActive = $true
            $script:SilentUserAsks = 0
            $script:LoudUserAsks   = 0
            $script:WatchLines.Clear()
            for ($i = 0; $i -lt 5; $i++) { Step-Clock 2.1; $null = Invoke-HeldBackRetry }
            $o['userWait'] = "$(Get-Lines '^a held-back repair for test trigger is still waiting: the user is active \(Claude in focus, recent input, or VM busy\)$')|" +
                "$(Get-Lines 'PRE-FIX WARNING')|$(Get-Lines 'BLOCKED')|$($global:RetryFixRuns)|$($script:SilentUserAsks)|$($script:LoudUserAsks)|$($script:PendingFixReason)"

            # Not elevated: Fix cannot be started from here, so it ends.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            Step-Clock 2.1
            $script:NotElevated = $true
            $null = Invoke-HeldBackRetry
            $o['notElevated'] = "$($null -eq $script:PendingFixReason)|$($global:RetryFixRuns)"

            # Three fixes in the last 30 minutes: the retry escalates like any
            # trigger, Fix does not run, and it stays pending.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            foreach ($ago in 20, 10, 5) { $null = $script:FixHistory.Add([datetime]::UtcNow.AddMinutes(-$ago)) }
            Step-Clock 2.1
            $null = Invoke-HeldBackRetry
            $o['escalated'] = "$(Get-Lines '^ESCALATION')|$($global:RetryFixRuns)|$($script:PendingFixReason)"

            # A re-checked trigger, as the service and vmcompute checks are:
            # held back, BLOCKED or DEFERRED, its repair is not kept, because
            # the check fires again while the fault lasts.
            Reset-Retry
            $null = Invoke-AutoFix -Reason 'test trigger'
            Step-Clock 2.1
            $null = Invoke-HeldBackRetry
            $o['recheckedHeld'] = "[$($script:PendingFixReason)]|$($global:RetryFixRuns)|$(Get-Lines '; the check fires again while the fault lasts$')|[$(Get-SavedPending)]"
            Reset-Retry
            $script:UserActive = $true
            $null = Invoke-AutoFix -Reason 'test trigger'
            $o['recheckedBlocked'] = "[$($script:PendingFixReason)]|$($global:RetryFixRuns)|$(Get-Lines '^>>> BLOCKED: test trigger, user is active')|" +
                "$(Get-Lines '^    Run Fix-ClaudeDesktop\.ps1 manually when ready$')"
            Reset-Retry
            $null = $script:DesktopStarts.Add([datetime]::UtcNow.AddSeconds(-30))
            $null = Invoke-AutoFix -Reason 'test trigger'
            $o['recheckedDeferred'] = "[$($script:PendingFixReason)]|$($global:RetryFixRuns)|$(Get-Lines '^>>> DEFERRED: test trigger, Claude launched \d+s ago')|" +
                "$(Get-Lines 'It is tried again once startup is over')"

            # A one-shot trigger BLOCKED or DEFERRED: kept pending as a hold
            # by Fix is, not counted as a try, and retried once that is over.
            Reset-Retry
            $script:UserActive = $true
            $null = Invoke-OneShot 'test trigger'
            $o['oneShotBlocked'] = "$($script:PendingFixReason)|$($global:RetryFixRuns)|$($script:PendingFixTries)|$(Get-Lines '^>>> BLOCKED: test trigger, user is active')|" +
                "$(Get-Lines '^    It is tried again once the user is away, for up to 4 hours, or run Fix-ClaudeDesktop\.ps1 manually when ready$')|$(Get-SavedPending)"
            $script:UserActive = $false
            $global:RetryFixExit = 0
            Step-Clock 2.1
            $null = Invoke-HeldBackRetry
            $o['oneShotBlockedRetry'] = "$($null -eq $script:PendingFixReason)|$($global:RetryFixRuns)|" +
                "$(Get-Lines '^>>> AUTO-FIX #\d+ TRIGGERED: retry of held-back repair: test trigger$')"
            Reset-Retry
            $null = $script:DesktopStarts.Add([datetime]::UtcNow.AddSeconds(-30))
            $null = Invoke-OneShot 'test trigger'
            $o['oneShotDeferred'] = "$($script:PendingFixReason)|$($global:RetryFixRuns)|$($script:PendingFixTries)|$(Get-Lines '^>>> DEFERRED: test trigger, Claude launched \d+s ago')|" +
                "$(Get-Lines '^    It is tried again once startup is over, for up to 4 hours$')|$(Get-SavedPending)"
            $global:RetryFixExit = 0
            Step-Clock 2.1
            $null = Invoke-HeldBackRetry
            $o['oneShotDeferredRetry'] = "$($null -eq $script:PendingFixReason)|$($global:RetryFixRuns)|" +
                "$(Get-Lines '^>>> AUTO-FIX #\d+ TRIGGERED: retry of held-back repair: test trigger$')"

            # A re-checked trigger held back while a one-shot repair is
            # pending: the pending repair keeps its reason, its time and its
            # tries, and waits from this hold.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            Step-Clock 5.1
            $first = $script:PendingFixSince
            $null = Invoke-AutoFix -Reason 'service down'
            $o['recheckedWhilePending'] = "$($script:PendingFixReason)|$($script:PendingFixSince -eq $first)|$(Test-JustNow $script:PendingFixHeldAt)|" +
                "$($script:PendingFixTries)|$($global:RetryFixRuns)"

            # The 4 hours are hard. A retry that starts just before they end
            # and that Fix holds back after them is dropped, once, and the
            # repair does not start 4 hours again.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $script:CodeActive = $true
            while ($script:Elapsed -lt 239.5) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $o['straddleBefore'] = "$($script:PendingFixReason)|$($global:RetryFixRuns)"
            $script:WatchLines.Clear()
            $script:CodeActive = $false
            $script:PendingFixNextCheck = [datetime]::MinValue
            $script:WarnMinutes = 1
            Step-Clock 0.3
            $null = Invoke-HeldBackRetry
            while ($script:Elapsed -lt 470) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $o['straddle'] = "$(Get-Lines '^>>> held-back repair for test trigger dropped after 4 hours; run Fix-ClaudeDesktop\.ps1 when you are ready$')|" +
                "[$($script:PendingFixReason)]|$($global:RetryFixRuns)|$(Get-Lines '^>>> PRE-FIX WARNING: retry of held-back repair: test trigger')|" +
                "$(Get-Lines 'dropped after 3 tries')|[$(Get-SavedPending)]|$(Get-Lines 'it is retried when the hold clears')"

            # A one-shot trigger of its own just before the 4 hours end, held
            # back by Fix after them: it takes the repair over, and its 4
            # hours start again from when it fired, so that hold does not
            # drop it. The hold is the repair's second try: its tries are kept.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $script:CodeActive = $true
            while ($script:Elapsed -lt 238) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $script:CodeActive = $false
            $script:WatchLines.Clear()
            $script:WarnMinutes = 4
            $null = Invoke-OneShot 'second trigger'
            $o['freshStraddle'] = "$(Get-Lines 'dropped after')|$($script:PendingFixReason)|$(Get-SinceMinutes)|" +
                "$($script:PendingFixTries)|$($global:RetryFixRuns)|$(Get-SavedPending)|$(Get-Lines 'it is retried when the hold clears')|$($script:WarnAt -join ',')"
            # The same, BLOCKED after them: kept, and not a try, as BLOCKED is
            # not, and not dropped on the poll after. LogStep: the 4 minutes
            # pass while the merge reads the Fix logs, after the trigger's own
            # time was taken.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $script:CodeActive = $true
            while ($script:Elapsed -lt 238) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $script:WatchLines.Clear()
            $script:UserActive = $true
            $script:LogStep = 4
            $null = Invoke-OneShot 'second trigger'
            $o['blockedStraddle'] = "$(Get-Lines 'dropped after')|" +
                "$(Get-Lines '^>>> BLOCKED: second trigger, user is active')|$($script:PendingFixReason)|" +
                "$($script:PendingFixTries)|$($global:RetryFixRuns)|$(Get-SavedPending)|$([math]::Round($script:Elapsed))"
            Step-Clock 0.5
            $null = Invoke-HeldBackRetry
            $o['blockedStraddle'] += "|$(Get-Lines 'dropped after')|[$($script:PendingFixReason)]"
            # The same, DEFERRED after them: the app has run all along, and a
            # process of its own started 30 seconds before. Kept, with no try,
            # and not dropped on the poll after.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $null = $script:DesktopStarts.Add([datetime]::UtcNow.AddMinutes(-300))
            $script:CodeActive = $true
            while ($script:Elapsed -lt 238) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $script:CodeActive = $false
            $script:WatchLines.Clear()
            $null = $script:DesktopStarts.Add([datetime]::UtcNow.AddMinutes(3.5))
            $script:LogStep = 4
            $null = Invoke-OneShot 'second trigger'
            $o['deferredStraddle'] = "$(Get-Lines 'dropped after')|" +
                "$(Get-Lines '^>>> DEFERRED: second trigger, Claude launched \d+s ago')|$($script:PendingFixReason)|" +
                "$($script:PendingFixTries)|$($global:RetryFixRuns)|$(Get-SavedPending)|$(Get-Lines '^    It is tried again once startup is over, for up to 4 hours$')"
            Step-Clock 0.5
            $null = Invoke-HeldBackRetry
            $o['deferredStraddle'] += "|$(Get-Lines 'dropped after')|[$($script:PendingFixReason)]"
            # A trigger of its own at 238, BLOCKED there, before any warning:
            # kept past the first hold's 4 hours, as its own run from when it
            # fired, and dropped once those end, at 478.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $script:CodeActive = $true
            while ($script:Elapsed -lt 238) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $script:WatchLines.Clear()
            $script:UserActive = $true
            $null = Invoke-OneShot 'second trigger'
            while ($script:Elapsed -lt 245) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $at245 = "$(Get-Lines 'dropped after')|$($script:PendingFixReason)|$(Get-SinceMinutes)|$($script:PendingFixTries)|$($global:RetryFixRuns)|$(Get-SavedPending)"
            $dropAt = 'never'
            while ($script:Elapsed -lt 485) {
                Step-Clock 0.5
                $null = Invoke-HeldBackRetry
                if ($dropAt -eq 'never' -and (Get-Lines 'dropped after') -gt 0) { $dropAt = $script:Elapsed }
            }
            $o['mergedKept'] = "$at245|$(Get-Lines '^>>> held-back repair for second trigger dropped after 4 hours; run Fix-ClaudeDesktop\.ps1 when you are ready$')|" +
                "$dropAt|[$($script:PendingFixReason)]"
            # A trigger of its own at 200, tried with the warning and held
            # back: its 4 hours run from 200 too, so it is not dropped at the
            # first hold's 4 hours, and is dropped once, at 440.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $script:CodeActive = $true
            while ($script:Elapsed -lt 200) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $script:CodeActive = $false
            $null = Invoke-OneShot 'second trigger'
            $script:CodeActive = $true
            $script:WatchLines.Clear()
            while ($script:Elapsed -lt 241) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $at241 = "$(Get-Lines 'dropped after')|[$($script:PendingFixReason)]|$($global:RetryFixRuns)"
            $dropAt = 'never'
            while ($script:Elapsed -lt 445) {
                Step-Clock 0.5
                $null = Invoke-HeldBackRetry
                if ($dropAt -eq 'never' -and (Get-Lines 'dropped after') -gt 0) { $dropAt = $script:Elapsed }
            }
            $o['mergedTried'] = "$at241|$(Get-Lines '^>>> held-back repair for second trigger dropped after 4 hours; run Fix-ClaudeDesktop\.ps1 when you are ready$')|" +
                "$dropAt|[$($script:PendingFixReason)]|$($global:RetryFixRuns)"
            # The monitor restarts right after a trigger of its own took the
            # repair over at 238 and was BLOCKED. The new start of its 4 hours
            # is in watch-state.json, so after the restart it is kept past the
            # first hold's 4 hours, and tried once the user is away.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $script:CodeActive = $true
            while ($script:Elapsed -lt 238) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $script:UserActive = $true
            $null = Invoke-OneShot 'second trigger'
            # The restart: what was in memory is gone, as at a new start, and
            # watch-state.json is read.
            $script:PendingFixReason      = $null
            $script:PendingFixSince       = [datetime]::MinValue
            $script:PendingFixHeldAt      = [datetime]::MinValue
            $script:PendingFixTries       = 0
            $script:PendingFixRetriesHeld = 0
            $script:PendingFixNextCheck   = [datetime]::MinValue
            $script:PendingFixWaitLogged  = [datetime]::MinValue
            $script:PendingFixLockLogged  = $false
            $script:FixHeldInARow         = 0
            $script:FixHeldLastTime       = [datetime]::MinValue
            $script:CooldownFromHold      = $false
            $script:LastFixTime           = [datetime]::MinValue
            $script:RepeatLinesSaid       = @{}
            $script:FixHistory.Clear()
            Import-WatchState
            $read = "$($script:PendingFixReason)|$(Get-SinceMinutes)|$($script:PendingFixTries)"
            $script:WatchLines.Clear()
            while ($script:Elapsed -lt 250) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $script:CodeActive = $false
            $script:UserActive = $false
            $global:RetryFixExit = 0
            while ($script:Elapsed -lt 300) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $o['restartMerged'] = "$read|$(Get-Lines 'dropped after')|" +
                "$(Get-Lines '^>>> AUTO-FIX #\d+ TRIGGERED: retry of held-back repair: second trigger$')|[$($script:PendingFixReason)]"

            # A trigger checked again on every poll, the service down, that
            # Fix holds back on every run: each run held back in a row
            # doubles the wait before the next repair, 2, 4, 8, 16, 32, then
            # 60 minutes, so in 4 hours the warning comes 8 times. Until 6.3.0
            # it came every 2 minutes. The COOLDOWN line it meets meanwhile is
            # written at most once every 10 minutes.
            Reset-Retry
            $cooldownAt = [System.Collections.ArrayList]::new()
            while ($script:Elapsed -lt 240) {
                $before = @($script:WatchLines).Count
                $null = Invoke-AutoFix -Reason 'service down'
                if (@(@($script:WatchLines) | Select-Object -Skip $before | Where-Object { $_ -match '^COOLDOWN: ' }).Count -gt 0) { [void]$cooldownAt.Add($script:Elapsed) }
                Step-Clock 1
            }
            $gaps = @(for ($i = 1; $i -lt $cooldownAt.Count; $i++) { $cooldownAt[$i] - $cooldownAt[$i - 1] })
            $minGap = if ($gaps.Count -gt 0) { ($gaps | Measure-Object -Minimum).Minimum } else { 'none' }
            $o['streak'] = "$($script:WarnAt -join ',')|$($global:RetryFixRuns)|" +
                "$(Get-Lines '^    Fix has held back 8 runs in a row, so the next repair waits 60 minutes$')|$minGap|[$($script:PendingFixReason)]"
            # A run that is not held back starts the count again: the next
            # hold waits 2 minutes.
            # Both the count and the time of the latest hold, each of which
            # alone would start it again.
            $global:RetryFixExit = 0
            Step-Clock 60
            $null = Invoke-AutoFix -Reason 'service down'
            $afterRun = "$($script:FixHeldInARow)|$($script:FixHeldLastTime -eq [datetime]::MinValue)"
            $global:RetryFixExit = 3
            Step-Clock 5.1
            $null = Invoke-AutoFix -Reason 'service down'
            $o['streakReset'] = "$($global:RetryFixRuns)|$afterRun|$($script:FixHeldInARow)|$([math]::Round(([datetime]::UtcNow - $script:LastFixTime).TotalMinutes, 1))"
            # So does a hold 4 hours or more after the one before.
            Reset-Retry
            $null = Invoke-AutoFix -Reason 'service down'
            Step-Clock 2.1
            $null = Invoke-AutoFix -Reason 'service down'
            $afterTwo = [math]::Round(([datetime]::UtcNow - $script:LastFixTime).TotalMinutes, 1)
            Step-Clock 240
            $null = Invoke-AutoFix -Reason 'service down'
            $o['streakDecay'] = "$($global:RetryFixRuns)|$afterTwo|$($script:FixHeldInARow)|$([math]::Round(([datetime]::UtcNow - $script:LastFixTime).TotalMinutes, 1))"
            # A one-shot trigger with nothing pending, during the 32 minutes
            # the fifth run Fix held back in a row leaves: kept and saved,
            # and tried once the wait is over, at 62. It does not fire again,
            # and the COOLDOWN gate alone lost it.
            Reset-Retry
            while ($script:FixHeldInARow -lt 5) { $null = Invoke-AutoFix -Reason 'service down'; Step-Clock 1 }
            Step-Clock 5
            $script:WatchLines.Clear()
            $null = Invoke-OneShot 'log trigger'
            $kept = "$($script:PendingFixReason)|$(Get-SavedPending)|$(Get-Lines '^COOLDOWN: Skipping fix \(log trigger\)')|" +
                "$(Get-Lines '^    It is tried again once the wait is over, for up to 4 hours$')|$($global:RetryFixRuns)"
            $global:RetryFixExit = 0
            $triedAt = 'never'
            while ($script:Elapsed -lt 90) {
                Step-Clock 0.5
                $runs = $global:RetryFixRuns
                $null = Invoke-HeldBackRetry
                if ($triedAt -eq 'never' -and $global:RetryFixRuns -gt $runs) { $triedAt = $script:Elapsed }
            }
            $o['oneShotInWait'] = "$kept|$triedAt|$(Get-Lines '^>>> AUTO-FIX #\d+ TRIGGERED: retry of held-back repair: log trigger$')|[$($script:PendingFixReason)]"
            # The trigger the repair is pending for, fired again in the same
            # wait once COOLDOWN's 10 minutes are over: said to be tried
            # again, its 4 hours from now, and not said to take over the
            # repair it is. Until 6.3.0 it said it took over the held-back
            # repair for itself.
            Reset-Retry
            while ($script:FixHeldInARow -lt 5) { $null = Invoke-AutoFix -Reason 'service down'; Step-Clock 1 }
            Step-Clock 5
            $null = Invoke-OneShot 'log trigger'
            Step-Clock 11
            $script:WatchLines.Clear()
            $null = Invoke-OneShot 'log trigger'
            $o['sameAgain'] = "$(Get-Lines '^COOLDOWN: Skipping fix \(log trigger\)')|" +
                "$(Get-Lines '^    It is tried again once the wait is over, for up to 4 hours$')|$(Get-Lines 'takes over')|" +
                "$($script:PendingFixReason)|$(Get-SinceMinutes)|$($global:RetryFixRuns)"
            # The same in the 2 minutes BLOCKED leaves for a re-checked
            # trigger.
            Reset-Retry
            $script:UserActive = $true
            $null = Invoke-AutoFix -Reason 'service down'
            $script:UserActive = $false
            Step-Clock 1
            $null = Invoke-OneShot 'log trigger'
            $kept = "$($script:PendingFixReason)|$(Get-SavedPending)"
            $global:RetryFixExit = 0
            for ($i = 0; $i -lt 10; $i++) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $o['oneShotInBlockedWait'] = "$kept|$(Get-Lines '^>>> AUTO-FIX #\d+ TRIGGERED: retry of held-back repair: log trigger$')|[$($script:PendingFixReason)]"
            # And in the minute DEFERRED leaves, the app started 30 seconds
            # before.
            Reset-Retry
            $null = $script:DesktopStarts.Add([datetime]::UtcNow.AddSeconds(-30))
            $null = Invoke-AutoFix -Reason 'service down'
            Step-Clock 0.5
            $null = Invoke-OneShot 'log trigger'
            $kept = "$($script:PendingFixReason)|$(Get-SavedPending)"
            $global:RetryFixExit = 0
            for ($i = 0; $i -lt 10; $i++) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $o['oneShotInDeferredWait'] = "$kept|$(Get-Lines '^>>> AUTO-FIX #\d+ TRIGGERED: retry of held-back repair: log trigger$')|[$($script:PendingFixReason)]"
            # In the cooldown after a repair that ran, as before, though a hold
            # came before it: skipped, not kept, and not tried. That repair has
            # just restarted the app and the service.
            Reset-Retry
            $script:UserActive = $true
            $null = Invoke-AutoFix -Reason 'service down'
            $script:UserActive = $false
            Step-Clock 2.1
            $global:RetryFixExit = 0
            $null = Invoke-AutoFix -Reason 'service down'
            Step-Clock 1
            $script:WatchLines.Clear()
            $null = Invoke-OneShot 'log trigger'
            $after = "[$($script:PendingFixReason)]|[$(Get-SavedPending)]|$(Get-Lines '^COOLDOWN: Skipping fix \(log trigger\)')|$(Get-Lines 'It is tried again')"
            for ($i = 0; $i -lt 20; $i++) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $o['oneShotAfterRun'] = "$after|$($global:RetryFixRuns)"
            # After an escalation and after a cancel too, though a hold came
            # before them.
            Reset-Retry
            $script:UserActive = $true
            $null = Invoke-AutoFix -Reason 'service down'
            $script:UserActive = $false
            Step-Clock 2.1
            foreach ($ago in 20, 10, 5) { $null = $script:FixHistory.Add([datetime]::UtcNow.AddMinutes(-$ago)) }
            $null = Invoke-AutoFix -Reason 'service down'
            Step-Clock 1
            $null = Invoke-OneShot 'log trigger'
            $o['oneShotAfterEscalation'] = "$(Get-Lines '^ESCALATION')|[$($script:PendingFixReason)]|$(Get-Lines '^COOLDOWN: Skipping fix \(log trigger\)')|$($global:RetryFixRuns)"
            Reset-Retry
            $script:UserActive = $true
            $null = Invoke-AutoFix -Reason 'service down'
            $script:UserActive = $false
            Step-Clock 2.1
            $script:CoworkActive = $true
            $null = Invoke-AutoFix -Reason 'service down'
            $script:CoworkActive = $false
            Step-Clock 1
            $null = Invoke-OneShot 'log trigger'
            $o['oneShotAfterCancel'] = "$(Get-Lines '^>>> CANCELLED: ')|[$($script:PendingFixReason)]|$(Get-Lines '^COOLDOWN: Skipping fix \(log trigger\)')|$($global:RetryFixRuns)"
            # A Fix run by hand that started the repair, from 31 to 34, in
            # the 32 minutes the fifth run Fix held back in a row left: its
            # log ends those runs. At 41 the service is down again: the
            # trigger gets through, and Fix's hold is the first in a row,
            # with a wait of 2 minutes, not 60. One that changed nothing, a
            # Diagnostic run, does not: COOLDOWN until 62.
            foreach ($mc in @(@{ Name = 'manualAfterStreak'; Text = $global:RetryRepairLine },
                              @{ Name = 'diagAfterStreak'; Text = '[10:00:09] Diagnostic complete, no changes were made' })) {
                Reset-Retry
                while ($script:FixHeldInARow -lt 5) { $null = Invoke-AutoFix -Reason 'service down'; Step-Clock 1 }
                $nowM = [datetime]::UtcNow
                $null = $script:FixLogs.Add(@{ C = $nowM; W = $nowM.AddMinutes(3); Text = $mc.Text })
                Step-Clock 10
                $null = Invoke-AutoFix -Reason 'service down'
                $o[$mc.Name] = "$($global:RetryFixRuns)|$($script:FixHeldInARow)|$([math]::Round(([datetime]::UtcNow - $script:LastFixTime).TotalMinutes, 1))"
            }
            # The same run by hand, from 31 to 32, and the service down again
            # at 34: the count is 0 and saved, and the wait ends 5 minutes
            # after that run, at 37, the cooldown after a repair. A log error
            # in it is skipped, as after a repair this monitor ran.
            Reset-Retry
            while ($script:FixHeldInARow -lt 5) { $null = Invoke-AutoFix -Reason 'service down'; Step-Clock 1 }
            $nowM = [datetime]::UtcNow
            $null = $script:FixLogs.Add(@{ C = $nowM; W = $nowM.AddMinutes(1); Text = $global:RetryRepairLine })
            Step-Clock 3
            $null = Invoke-AutoFix -Reason 'service down'
            $null = Invoke-OneShot 'log trigger'
            $o['manualCooldown'] = "$($global:RetryFixRuns)|$($script:FixHeldInARow)|$(Get-SavedField -Name 'FixHeldInARow')|" +
                "$([math]::Round(($script:LastFixTime.AddMinutes(5) - [datetime]::UtcNow).TotalMinutes, 1))|[$($script:PendingFixReason)]"
            # It never makes a wait longer: after one hold, whose wait ends at
            # 2, a run by hand from 1 to 4, and the service down at 5. The
            # trigger gets through, as it would have without that run.
            Reset-Retry
            $null = Invoke-AutoFix -Reason 'service down'
            $nowM = [datetime]::UtcNow
            $null = $script:FixLogs.Add(@{ C = $nowM.AddMinutes(1); W = $nowM.AddMinutes(4); Text = $global:RetryRepairLine })
            Step-Clock 5
            $null = Invoke-AutoFix -Reason 'service down'
            $o['manualNoLonger'] = "$($global:RetryFixRuns)|$($script:FixHeldInARow)"

            # BLOCKED for a trigger checked again on every poll while the
            # user is active: it comes back every 2 minutes, and its lines are
            # written at most once every 10 minutes, as is COOLDOWN between.
            # Until 6.3.0 each was written every time.
            Reset-Retry
            $script:UserActive = $true
            while ($script:Elapsed -lt 30) { $null = Invoke-AutoFix -Reason 'service down'; Step-Clock 1 }
            $o['blockedEvery'] = "$(Get-Lines '^>>> BLOCKED: service down, user is active')|$(Get-Lines '^    Run Fix-ClaudeDesktop\.ps1 manually when ready$')|" +
                "$(Get-Lines '^COOLDOWN: Skipping fix \(service down\)')|$($global:RetryFixRuns)"

            # Fix run by hand, and holding its lock when the retry comes. The
            # run may be the repair, or one that changes nothing, such as a
            # Diagnostic run: the repair is kept while it lasts, said once,
            # and decided from its log once it is over. Until 6.3.0 the lock
            # ended it as "that run is the repair", so a Diagnostic run going
            # on at a retry dropped it.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $script:WatchLines.Clear()
            Step-Clock 2.1
            $manualMutex = [System.Threading.Mutex]::new($true, $retryLockName)
            try {
                for ($i = 0; $i -lt 12; $i++) { $null = Invoke-HeldBackRetry; Step-Clock 0.5 }
                $o['manualLock'] = "$(Get-Lines '^AUTO-FIX SKIPPED: Fix-ClaudeDesktop is already running$')|" +
                    "$(Get-Lines '^    The held-back repair for test trigger is kept while that run lasts: it ends if the run gets as far as the repair, and is tried again if the run stops before it, as a Diagnostic run does$')|" +
                    "$(Get-Lines 'ended')|[$($script:PendingFixReason)]|[$(Get-SavedPending)]"
                # It is a Diagnostic run: its log does not hold the line.
                $nowM = [datetime]::UtcNow
                $null = $script:FixLogs.Add(@{ C = $nowM.AddMinutes(-5); W = $nowM; Text = '[10:00:09] Diagnostic complete, no changes were made' })
            } finally {
                $manualMutex.ReleaseMutex()
                $manualMutex.Dispose()
            }
            $global:RetryFixExit = 0
            for ($i = 0; $i -lt 12; $i++) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $o['manualLockAfter'] = "$(Get-Lines 'SKIPPED')|$(Get-Lines 'ended')|$(Get-Lines '^>>> PRE-FIX WARNING: retry of held-back repair: test trigger')|" +
                "$($global:RetryFixRuns)|[$($script:PendingFixReason)]"
            # The same, and the run gets as far as the repair: its log holds
            # the line, and the repair ends once the run is over, said once.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $script:WatchLines.Clear()
            Step-Clock 2.1
            $manualMutex = [System.Threading.Mutex]::new($true, $retryLockName)
            try {
                for ($i = 0; $i -lt 6; $i++) { $null = Invoke-HeldBackRetry; Step-Clock 0.5 }
                $nowM = [datetime]::UtcNow
                $null = $script:FixLogs.Add(@{ C = $nowM.AddMinutes(-2); W = $nowM; Text = $global:RetryRepairLine })
            } finally {
                $manualMutex.ReleaseMutex()
                $manualMutex.Dispose()
            }
            $global:RetryFixExit = 0
            for ($i = 0; $i -lt 12; $i++) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $o['manualLockRepaired'] = "$(Get-Lines 'SKIPPED')|" +
                "$(Get-Lines '^>>> held-back repair for test trigger ended: Fix-ClaudeDesktop ran after it was held back, so that run was the repair$')|" +
                "$(Get-Lines 'PRE-FIX WARNING')|$($global:RetryFixRuns)|[$($script:PendingFixReason)]|[$(Get-SavedPending)]"
            # A second Fix run that takes the lock while the repair is still
            # pending is said again, once: after the first is over, whether a
            # retry or a trigger was the first to find the lock gone.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $script:WatchLines.Clear()
            Step-Clock 2.1
            $manualMutex = [System.Threading.Mutex]::new($true, $retryLockName)
            try { for ($i = 0; $i -lt 6; $i++) { $null = Invoke-HeldBackRetry; Step-Clock 0.5 } }
            finally { $manualMutex.ReleaseMutex(); $manualMutex.Dispose() }
            # The retry finds the lock gone, and waits on a Code session.
            $script:CodeActive = $true
            for ($i = 0; $i -lt 6; $i++) { $null = Invoke-HeldBackRetry; Step-Clock 0.5 }
            $manualMutex = [System.Threading.Mutex]::new($true, $retryLockName)
            try { for ($i = 0; $i -lt 6; $i++) { $null = Invoke-HeldBackRetry; Step-Clock 0.5 } }
            finally { $manualMutex.ReleaseMutex(); $manualMutex.Dispose() }
            $o['lockTwice'] = "$(Get-Lines '^AUTO-FIX SKIPPED: Fix-ClaudeDesktop is already running$')|$(Get-Lines 'is kept while that run lasts')|" +
                "[$($script:PendingFixReason)]|$($global:RetryFixRuns)"
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $script:WatchLines.Clear()
            $manualMutex = [System.Threading.Mutex]::new($true, $retryLockName)
            try { Step-Clock 0.2; $null = Invoke-OneShot 'second trigger' }
            finally { $manualMutex.ReleaseMutex(); $manualMutex.Dispose() }
            # A re-checked trigger finds the lock gone, in the cooldown.
            Step-Clock 0.2
            $null = Invoke-AutoFix -Reason 'service down'
            $manualMutex = [System.Threading.Mutex]::new($true, $retryLockName)
            try { Step-Clock 0.2; $null = Invoke-OneShot 'third trigger' }
            finally { $manualMutex.ReleaseMutex(); $manualMutex.Dispose() }
            $o['lockTwiceAuto'] = "$(Get-Lines '^AUTO-FIX SKIPPED: Fix-ClaudeDesktop is already running$')|$(Get-Lines 'is kept while that run lasts')|" +
                "$($script:PendingFixReason)|$($global:RetryFixRuns)"
            # A one-shot trigger of its own while a Fix run holds the lock:
            # merged into the pending repair as at any other time, and the
            # lock still said once.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $script:WatchLines.Clear()
            $manualMutex = [System.Threading.Mutex]::new($true, $retryLockName)
            try {
                Step-Clock 1
                $first = $script:PendingFixSince
                $null = Invoke-OneShot 'second trigger'
                $restarted = ($script:PendingFixSince -gt $first)
                Step-Clock 2.1
                $null = Invoke-HeldBackRetry
                $o['lockMerge'] = "$($script:PendingFixReason)|$restarted|$(Get-Lines 'SKIPPED')|" +
                    "$(Get-Lines '^    The held-back repair for second trigger is kept while that run lasts')|$(Get-SavedPending)|$($global:RetryFixRuns)"
            } finally {
                $manualMutex.ReleaseMutex()
                $manualMutex.Dispose()
            }
            # A one-shot trigger while a Fix run holds the lock, with nothing
            # pending: kept, as Fix's own hold at its lock keeps it, said
            # once while the lock lasts, and tried again once that run is
            # over, as it changed nothing. A re-checked one is skipped each
            # time it fires, and not kept.
            Reset-Retry
            $manualMutex = [System.Threading.Mutex]::new($true, $retryLockName)
            try {
                $null = Invoke-OneShot 'test trigger'
                Step-Clock 1
                $null = Invoke-OneShot 'test trigger'
                for ($i = 0; $i -lt 6; $i++) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
                $o['lockFresh'] = "$(Get-Lines 'SKIPPED')|$(Get-Lines '^    The held-back repair for test trigger is kept while that run lasts')|" +
                    "[$($script:PendingFixReason)]|[$(Get-SavedPending)]|$($global:RetryFixRuns)"
            } finally {
                $manualMutex.ReleaseMutex()
                $manualMutex.Dispose()
            }
            $global:RetryFixExit = 0
            for ($i = 0; $i -lt 8; $i++) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $o['lockFreshAfter'] = "$(Get-Lines '^>>> PRE-FIX WARNING: retry of held-back repair: test trigger')|$($global:RetryFixRuns)|[$($script:PendingFixReason)]"
            Reset-Retry
            $manualMutex = [System.Threading.Mutex]::new($true, $retryLockName)
            try {
                for ($i = 0; $i -lt 3; $i++) { $null = Invoke-AutoFix -Reason 'service down'; Step-Clock 1 }
                $o['lockRechecked'] = "$(Get-Lines '^AUTO-FIX SKIPPED: Fix-ClaudeDesktop is already running$')|$(Get-Lines 'is kept while that run lasts')|" +
                    "[$($script:PendingFixReason)]"
            } finally {
                $manualMutex.ReleaseMutex()
                $manualMutex.Dispose()
            }
            # Nothing pending, and a Fix run by hand waits in its menu for an
            # hour while a re-checked trigger fires on every poll: SKIPPED at
            # most once every 10 minutes. Until 6.3.0 it was written each
            # time, 120 times here.
            Reset-Retry
            $manualMutex = [System.Threading.Mutex]::new($true, $retryLockName)
            try {
                for ($i = 0; $i -lt 120; $i++) { Step-Clock 0.5; $null = Invoke-AutoFix -Reason 'service down'; $null = Invoke-HeldBackRetry }
                $o['lockNoPending'] = "$(Get-Lines '^AUTO-FIX SKIPPED: Fix-ClaudeDesktop is already running$')|[$($script:PendingFixReason)]|$($global:RetryFixRuns)"
            } finally {
                $manualMutex.ReleaseMutex()
                $manualMutex.Dispose()
            }

            # Fix run by hand between the hold and the retry: its log was
            # created after the hold. Said once and saved.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $script:WatchLines.Clear()
            Step-Clock 2.1
            $nowM = [datetime]::UtcNow
            $null = $script:FixLogs.Add(@{ C = $nowM.AddMinutes(-1); W = $nowM.AddMinutes(-0.5); Text = $global:RetryRepairLine })
            $null = Invoke-HeldBackRetry
            for ($i = 0; $i -lt 6; $i++) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $o['manualLog'] = "$(Get-Lines '^>>> held-back repair for test trigger ended: Fix-ClaudeDesktop ran after it was held back, so that run was the repair$')|" +
                "[$($script:PendingFixReason)]|$($global:RetryFixRuns)|$(@($script:WatchLines).Count)|[$(Get-SavedPending)]"
            # Running when the hold came, and written after it.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $nowM = [datetime]::UtcNow
            $null = $script:FixLogs.Add(@{ C = $nowM.AddMinutes(-3); W = $nowM.AddMinutes(1); Text = $global:RetryRepairLine })
            Step-Clock 2.1
            $null = Invoke-HeldBackRetry
            $o['manualOverlap'] = "$(Get-Lines 'ran after it was held back')|[$($script:PendingFixReason)]|$($global:RetryFixRuns)"
            # Fix run by hand after the hold, and it changed nothing: a
            # Diagnostic run, a Cancel in the menu, a No at "Proceed?". Its
            # log does not hold the line, so the repair is kept, and the retry
            # goes ahead. Until 6.3.0 any Fix log since the hold ended it.
            foreach ($nc in @(@{ Name = 'diagLog'; Text = '[10:00:01] Diagnostic complete, no changes were made' },
                              @{ Name = 'cancelLog'; Text = '[10:00:01] Cancelled by user' },
                              @{ Name = 'noLog'; Text = "  Proceed? (Y/n): n`r`n[10:00:02] Cancelled by user" })) {
                Reset-Retry
                $null = Invoke-OneShot 'test trigger'
                $script:WatchLines.Clear()
                Step-Clock 2.1
                $nowM = [datetime]::UtcNow
                $null = $script:FixLogs.Add(@{ C = $nowM.AddMinutes(-1); W = $nowM.AddMinutes(-0.5); Text = $nc.Text })
                $global:RetryFixExit = 0
                $null = Invoke-HeldBackRetry
                $o[$nc.Name] = "$(Get-Lines 'ran after it was held back')|$(Get-Lines '^>>> PRE-FIX WARNING: retry of held-back repair: test trigger')|$($global:RetryFixRuns)"
            }
            # The held-back run's own logs, written before Fix returned: not a
            # Fix run since, and the retry goes ahead. It holds the line, so
            # that only its times keep it from counting.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $nowM = [datetime]::UtcNow
            $null = $script:FixLogs.Add(@{ C = $nowM.AddMinutes(-1); W = $nowM.AddSeconds(-5); Text = $global:RetryRepairLine })
            Step-Clock 2.1
            $null = Invoke-HeldBackRetry
            $o['ownLog'] = "$(Get-Lines 'ran after it was held back')|$($null -ne $script:PendingFixReason)|$($global:RetryFixRuns)|$($script:PendingFixRetriesHeld)"
            # Each run Fix holds back writes its log, the retries' included:
            # none is a Fix run since the latest hold, though the second is
            # newer than the first hold.
            Reset-Retry
            $global:RetryFixWritesLog = $true
            $null = Invoke-OneShot 'test trigger'
            Step-Clock 2.1
            $null = Invoke-HeldBackRetry
            Step-Clock 4.1
            $null = Invoke-HeldBackRetry
            $o['ownRetryLogs'] = "$(Get-Lines 'ran after it was held back')|$($global:RetryFixRuns)|$(Get-Lines 'dropped after 3 tries')|$($script:FixLogs.Count)"

            # Claude Desktop started again since the hold: what was detected
            # came from the session that ended. Said once and saved.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            $script:WatchLines.Clear()
            Step-Clock 3
            $nowA = [datetime]::UtcNow
            $null = $script:DesktopStarts.Add($nowA.AddMinutes(-2.5))
            $null = $script:DesktopStarts.Add($nowA.AddMinutes(-1))
            $null = Invoke-HeldBackRetry
            for ($i = 0; $i -lt 6; $i++) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $o['appRestart'] = "$(Get-Lines '^>>> held-back repair for test trigger ended: Claude Desktop was started again since it was held back$')|" +
                "[$($script:PendingFixReason)]|$($global:RetryFixRuns)|[$(Get-SavedPending)]"
            # A helper process started since, with the main process older
            # than the hold, and one that has just exited: not a restart.
            Reset-Retry
            $null = $script:DesktopStarts.Add([datetime]::UtcNow.AddMinutes(-30))
            $null = Invoke-OneShot 'test trigger'
            Step-Clock 4
            $null = $script:DesktopStarts.Add([datetime]::UtcNow.AddMinutes(-2.5))
            $null = $script:DesktopStarts.Add('THROW')
            $null = Invoke-HeldBackRetry
            $o['helperRestart'] = "$(Get-Lines 'started again')|$($global:RetryFixRuns)|$($script:PendingFixRetriesHeld)"

            # A user the warning disturbs, who moves the mouse each time it
            # comes up, so that Fix holds every run back: three warnings in
            # all, the first included, then dropped, said once.
            Reset-Retry
            $global:RetryUseInput    = $true
            $global:RetryHoldOnInput = $true
            $global:RetryTouch       = $true
            $global:RetryInputAgoMin = 10
            $global:RetryFixExit     = 0
            $script:WarnMinutes      = 0.5
            $null = Invoke-OneShot 'test trigger'
            while ($script:Elapsed -lt 240) { Step-Clock 0.5; $null = Invoke-HeldBackRetry }
            $o['touch'] = "$(Get-Lines 'PRE-FIX WARNING')|$(Get-Lines '^>>> held-back repair for test trigger dropped after 3 tries; run Fix-ClaudeDesktop\.ps1 when you are ready$')|" +
                "$(Get-Lines 'dropped after 4 hours')|[$($script:PendingFixReason)]|$($global:RetryFixRuns)|[$(Get-SavedPending)]"
            $o['touchAt'] = ($script:WarnAt -join ',')

            # Each retry Fix holds back doubles the wait from that hold: 2
            # minutes after a hold that was not a retry, then 4, then 8.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            Step-Clock 2.1
            $null = Invoke-HeldBackRetry
            $b1 = $global:RetryFixRuns
            Step-Clock 3.9
            $null = Invoke-HeldBackRetry
            $b2 = $global:RetryFixRuns
            Step-Clock 0.2
            $null = Invoke-HeldBackRetry
            $o['backoff'] = "$b1|$b2|$($global:RetryFixRuns)|$(Get-Lines 'dropped after 3 tries')|[$($script:PendingFixReason)]"
            # From a BLOCKED origin, which is not a try: three warnings too.
            Reset-Retry
            $script:UserActive = $true
            $null = Invoke-OneShot 'test trigger'
            $script:UserActive = $false
            $runs = @()
            foreach ($m in 2.1, 3.9, 0.2, 7.9, 0.2) { Step-Clock $m; $null = Invoke-HeldBackRetry; $runs += $global:RetryFixRuns }
            $o['backoffBlocked'] = "$($runs -join ',')|$(Get-Lines 'PRE-FIX WARNING')|$(Get-Lines 'dropped after 3 tries')|[$($script:PendingFixReason)]"

            # watch-state.json: written and read back, and older files.
            Reset-Retry
            $null = Invoke-OneShot 'test trigger'
            Step-Clock 2.1
            $null = Invoke-HeldBackRetry
            $savedSince = $script:PendingFixSince.ToString('o')
            $savedHeld  = $script:PendingFixHeldAt.ToString('o')
            $script:PendingFixReason      = $null
            $script:PendingFixSince       = [datetime]::MinValue
            $script:PendingFixHeldAt      = [datetime]::MinValue
            $script:PendingFixTries       = 0
            $script:PendingFixRetriesHeld = 0
            Import-WatchState
            $o['roundTrip'] = "$($script:PendingFixReason)|$($script:PendingFixSince.ToString('o') -eq $savedSince -and $script:PendingFixHeldAt.ToString('o') -eq $savedHeld)|" +
                "$($script:PendingFixSince.Kind)|$($script:PendingFixHeldAt.Kind)|$($script:PendingFixTries)|$($script:PendingFixRetriesHeld)"
            $o['savedFields'] = (@(([System.IO.File]::ReadAllText($script:StatePath) | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $_.Name }) -join ',')
            function Read-Pending {
                param([string]$Json)
                [System.IO.File]::WriteAllText($script:StatePath, $Json)
                $script:PendingFixReason = $null
                $script:PendingFixSince  = [datetime]::MinValue
                # Left from an earlier repair: a file that names none, or
                # none that can be read, leaves none of them.
                $script:PendingFixHeldAt      = [datetime]::new(2026, 1, 1, 0, 0, 0, [System.DateTimeKind]::Utc)
                $script:PendingFixTries       = 99
                $script:PendingFixRetriesHeld = 99
                $script:LastFixTime      = [datetime]::MinValue
                $script:FixHistory.Clear()
                Import-WatchState
                $inv = [System.Globalization.CultureInfo]::InvariantCulture
                return "[$($script:PendingFixReason)] $($script:PendingFixSince.ToString('yyyy-MM-ddTHH:mm:ss', $inv)) " +
                    "$($script:PendingFixHeldAt.ToString('yyyy-MM-ddTHH:mm:ss', $inv)) $($script:PendingFixTries) $($script:PendingFixRetriesHeld) | " +
                    "$($script:LastFixTime.ToString('yyyy-MM-ddTHH:mm:ss', $inv)) $($script:FixHistory.Count)"
            }
            $base = '"LastFixTime": "2026-11-22T10:00:00.0000000Z", "FixCount": 3, "FixHistory": [ "2026-11-22T09:50:00.0000000Z" ]'
            $o['stateOld']    = Read-Pending "{ $base, ""LastStealthMarkerTs"": 0 }"
            $o['stateNone']   = Read-Pending "{ $base, ""PendingFixReason"": null, ""PendingFixSince"": ""0001-01-01T00:00:00.0000000"" }"
            $o['stateBad']    = Read-Pending "{ $base, ""PendingFixReason"": ""test trigger"", ""PendingFixSince"": ""not a time"" }"
            $o['stateSet']    = Read-Pending "{ $base, ""PendingFixReason"": ""test trigger"", ""PendingFixSince"": ""2026-11-22T09:55:00.0000000Z"" }"
            $o['stateNoTime'] = Read-Pending "{ $base, ""PendingFixReason"": ""test trigger"", ""PendingFixSince"": ""0001-01-01T00:00:00.0000000"" }"
            $o['stateFull']   = Read-Pending ("{ $base, ""PendingFixReason"": ""test trigger"", ""PendingFixSince"": ""2026-11-22T09:55:00.0000000Z"", " +
                """PendingFixHeldAt"": ""2026-11-22T10:05:00.0000000Z"", ""PendingFixTries"": 2, ""PendingFixRetriesHeld"": 1 }")
            $o['stateBadNew'] = Read-Pending ("{ $base, ""PendingFixReason"": ""test trigger"", ""PendingFixSince"": ""2026-11-22T09:55:00.0000000Z"", " +
                """PendingFixHeldAt"": ""not a time"", ""PendingFixTries"": ""x"", ""PendingFixRetriesHeld"": ""y"" }")
            $o['stateNegTries'] = Read-Pending ("{ $base, ""PendingFixReason"": ""test trigger"", ""PendingFixSince"": ""2026-11-22T09:55:00.0000000Z"", " +
                """PendingFixHeldAt"": ""2026-11-22T10:05:00.0000000Z"", ""PendingFixTries"": -2, ""PendingFixRetriesHeld"": -5 }")
            $o['stateNewNoPending'] = Read-Pending ("{ $base, ""PendingFixReason"": null, ""PendingFixSince"": ""0001-01-01T00:00:00.0000000"", " +
                """PendingFixHeldAt"": ""2026-11-22T10:05:00.0000000Z"", ""PendingFixTries"": 2, ""PendingFixRetriesHeld"": 1 }")
            # The runs Fix held back in a row, and when the latest was.
            function Read-Held {
                param([string]$Json)
                [System.IO.File]::WriteAllText($script:StatePath, $Json)
                # Left from earlier holds: a file that has none, or none that
                # can be read, leaves none of them.
                $script:FixHeldInARow   = 99
                $script:FixHeldLastTime = [datetime]::new(2026, 1, 1, 0, 0, 0, [System.DateTimeKind]::Utc)
                Import-WatchState
                $inv = [System.Globalization.CultureInfo]::InvariantCulture
                $heldTime = $script:FixHeldLastTime
                if ($heldTime -is [datetime]) { $heldTime = "$($heldTime.ToString('yyyy-MM-ddTHH:mm:ss', $inv)) $(if ($heldTime.Ticks -eq 0) { 'none' } else { $heldTime.Kind })" }
                return "$($script:FixHeldInARow) $heldTime"
            }
            $o['heldOld']  = Read-Held "{ $base, ""LastStealthMarkerTs"": 0, ""PendingFixReason"": null, ""PendingFixSince"": ""0001-01-01T00:00:00.0000000"", ""PendingFixHeldAt"": ""0001-01-01T00:00:00.0000000"", ""PendingFixTries"": 0, ""PendingFixRetriesHeld"": 0 }"
            $o['heldSet']  = Read-Held "{ $base, ""FixHeldInARow"": 3, ""FixHeldLastTime"": ""2026-11-22T10:05:00.0000000Z"" }"
            $o['heldBad']  = Read-Held "{ $base, ""FixHeldInARow"": ""x"", ""FixHeldLastTime"": ""not a time"" }"
            $o['heldNeg']  = Read-Held "{ $base, ""FixHeldInARow"": -4, ""FixHeldLastTime"": ""0001-01-01T00:00:00.0000000"" }"
            Reset-Retry
            $null = Invoke-AutoFix -Reason 'service down'
            Step-Clock 2.1
            $null = Invoke-AutoFix -Reason 'service down'
            $savedHeldTime = $script:FixHeldLastTime.ToString('o')
            $script:FixHeldInARow   = 0
            $script:FixHeldLastTime = [datetime]::MinValue
            Import-WatchState
            $o['heldRoundTrip'] = "$($script:FixHeldInARow)|$($script:FixHeldLastTime.ToString('o') -eq $savedHeldTime)|$($script:FixHeldLastTime.Kind)"
            # And a file written now, read the way a monitor from before
            # these fields reads it: by name, the rest ignored. Two runs,
            # both held back, none in the history.
            $savedNow = [System.IO.File]::ReadAllText($script:StatePath) | ConvertFrom-Json
            $o['heldOlderReader'] = "$($savedNow.FixCount)|$(@($savedNow.FixHistory).Count)|[$($savedNow.PendingFixReason)]"
            # Whether a hold left the cooldown: those two holds did, and it
            # reads back as written.
            $script:CooldownFromHold = $false
            Import-WatchState
            # Read by name, so a file without it gives a result, not an error.
            $written = $savedNow.PSObject.Properties['CooldownFromHold']
            $o['cooldownRoundTrip'] = "$($script:CooldownFromHold)|$(if ($written) { $written.Value } else { 'absent' })"
            # A file from before it was kept, and one that cannot be read:
            # not, whatever was in memory.
            function Read-CooldownSource {
                param([string]$Json)
                [System.IO.File]::WriteAllText($script:StatePath, $Json)
                $script:CooldownFromHold = $true
                Import-WatchState
                return "$($script:CooldownFromHold)"
            }
            $o['cooldownOld'] = Read-CooldownSource "{ $base, ""FixHeldInARow"": 3, ""FixHeldLastTime"": ""2026-11-22T10:05:00.0000000Z"" }"
            $o['cooldownSet'] = Read-CooldownSource "{ $base, ""CooldownFromHold"": true }"
            $o['cooldownNo']  = Read-CooldownSource "{ $base, ""CooldownFromHold"": false }"
            $o['cooldownBad'] = Read-CooldownSource "{ $base, ""CooldownFromHold"": ""yes"" }"
            $o
        }
    }
}
if ($null -eq $retry) { $retry = [ordered]@{} }
Assert-Case "Invoke-AutoFix takes -OneShot, for a trigger that does not fire again" ($retry['takesOneShot'] -eq $true) "got $($retry['takesOneShot'])"
Assert-Case "held back: the repair is pending, and saved" ($retry['heldPending'] -eq 'test trigger|1|test trigger') "got $($retry['heldPending'])"
Assert-Case "not asked about before the cooldown the hold leaves is over" ($retry['inCooldownAsks'] -eq 0) "asked $($retry['inCooldownAsks']) time(s)"
Assert-Case "still waiting: asked no more than every 2 minutes" ($retry['askedAgainAtOnce'] -eq 0 -and $retry['waitAsks'] -eq 5) `
    "asked again at once $($retry['askedAgainAtOnce']), asked $($retry['waitAsks']) time(s) in 5 checks"
Assert-Case "still waiting: one line in the first 10 minutes, a second after them" ($retry['waitLines10'] -eq 1 -and $retry['waitLines12'] -eq 2) `
    "lines $($retry['waitLines10']) and $($retry['waitLines12'])"
Assert-Case "and the line says why" `
    ($retry['waitText'] -eq 'a held-back repair for test trigger is still waiting: a Claude Code session is active, or a transcript was written in the last 35 minutes') `
    "got $($retry['waitText'])"
Assert-Case "the app's own Claude Code sessions hold it too" `
    ($retry['waitCli'] -eq 'a held-back repair for test trigger is still waiting: 2 Claude Code session(s) the app started are running') "got $($retry['waitCli'])"
Assert-Case "and so does not being able to list them" `
    ($retry['waitCliThrow'] -eq ('a held-back repair for test trigger is still waiting: the Claude Code sessions the app started ' +
        'could not be listed (the parents could not be read)')) "got $($retry['waitCliThrow'])"
Assert-Case "while it waits: no warning and no Fix, it stays pending, and the Code check is asked silently" ($retry['waitQuiet'] -eq '0|1|test trigger|0') `
    "warnings|Fix runs|pending|asked aloud: $($retry['waitQuiet'])"
Assert-Case "the session gone: Fix runs once, for the retry" `
    ($retry['retryRuns'] -eq 2 -and $retry['retryWarning'] -eq 1 -and $retry['retryTriggered'] -eq 1) `
    "Fix runs $($retry['retryRuns']), warnings $($retry['retryWarning']), triggered $($retry['retryTriggered'])"
Assert-Case "and the repair that ran ends it, counted and saved" ($retry['retryEnded'] -eq 'True|1|[]') "ended|history|saved: $($retry['retryEnded'])"
Assert-Case "and it is not retried again" ($retry['retryOnce'] -eq 2) "Fix runs $($retry['retryOnce'])"
Assert-Case "past 4 hours: dropped, and said once" ($retry['dropLines'] -eq 1) "lines $($retry['dropLines'])"
Assert-Case "and Fix does not run, and the drop is saved" ($retry['dropEnded'] -eq 'True|1|0|[]') "ended|Fix runs|warnings|saved: $($retry['dropEnded'])"
Assert-Case "under 4 hours: kept" ($retry['notDropped'] -eq 'test trigger|0') "got $($retry['notDropped'])"
Assert-Case "past 4 hours during the doubled wait: dropped on that poll" ($retry['dropInWait'] -eq '1|[]|1') `
    "dropped|pending|Fix runs: $($retry['dropInWait'])"
Assert-Case "past 4 hours, a one-shot trigger of its own: the old repair is dropped, once, and the new hold has 4 hours of its own" `
    ($retry['expiredHeld'] -eq 'second trigger|True|1|2') "reason|held just now|dropped|Fix runs: $($retry['expiredHeld'])"
Assert-Case "held back again on the retry: first-held time and reason kept, not an attempt, a try, a retry held back, the latest hold" `
    ($retry['repeat'] -eq '2|test trigger|True|0|2|1|True') "Fix runs|reason|first time kept|history|tries|retries held|latest hold now: $($retry['repeat'])"
Assert-Case "a trigger of its own in the cooldown: takes the repair over, its 4 hours start again, it runs nothing" `
    ($retry['freshCooldown'] -eq 'second trigger|0|1|second trigger') "reason|minutes since the 4 hours started|Fix runs|saved: $($retry['freshCooldown'])"
Assert-Case "and after COOLDOWN it says it takes over the held-back repair, naming it, and is tried again once the wait is over" `
    ($retry['freshCooldownLines'] -eq '1|1|0') "COOLDOWN|takes over|kept as new: $($retry['freshCooldownLines'])"
Assert-Case "a trigger of its own held back again: the same" ($retry['freshHeld'] -eq 'third trigger|0|2') `
    "reason|minutes since the 4 hours started|Fix runs: $($retry['freshHeld'])"
Assert-Case "and the retry is for that trigger, 4 minutes after the second run held back in a row" ($retry['freshRetry'] -eq 1) "got $($retry['freshRetry'])"
Assert-Case "the takeover line comes with COOLDOWN: the same trigger again in the wait is not said again, another trigger is, and taking the repair back is not" `
    ($retry['freshCooldownAgain'] -eq '2|2|1|second trigger|1') "COOLDOWN|takes over|for second trigger|pending|Fix runs: $($retry['freshCooldownAgain'])"
Assert-Case "a pending repair past its 4 hours: dropped, and the trigger in the wait is kept as a new one, not said to take it over" `
    ($retry['freshCooldownStale'] -eq '1|0|1|second trigger|0') "dropped|takes over|kept|pending|minutes since: $($retry['freshCooldownStale'])"
Assert-Case "a retry that stops part-way (exit 4) ends it, and is an attempt" ($retry['end4'] -eq 'True|2|1') "ended|Fix runs|history: $($retry['end4'])"
Assert-Case "a retry that finishes (exit 0) ends it" ($retry['end0'] -eq 'True|2|1') "ended|Fix runs|history: $($retry['end0'])"
Assert-Case "a retry Fix could not be started for ends it" ($retry['endThrow'] -eq 'True|2|1') "ended|Fix runs|history: $($retry['endThrow'])"
Assert-Case "a trigger of its own that runs ends it" ($retry['endFresh'] -eq 'True|2') "got $($retry['endFresh'])"
Assert-Case "the warning cancelled: it ends, Fix does not run, and that is saved" ($retry['cancel'] -eq 'True|1|1|[]') "ended|Fix runs|cancelled|saved: $($retry['cancel'])"
Assert-Case "BLOCKED on the retry: it stays pending with its first-held time, and waits from this hold" `
    ($retry['blockedRace'] -eq 'test trigger|1|1|True|True|1') "pending|Fix runs|blocked|first time kept|latest hold now|retried line: $($retry['blockedRace'])"
Assert-Case "the user at the PC: the quiet check asks silently, and the repair waits with one line, no warning and no BLOCKED" `
    ($retry['userWait'] -eq '1|0|0|1|5|0|test trigger') "lines|warnings|blocked|Fix runs|asked silently|asked aloud|pending: $($retry['userWait'])"
Assert-Case "this monitor not elevated: it ends" ($retry['notElevated'] -eq 'True|1') "got $($retry['notElevated'])"
Assert-Case "three fixes in 30 minutes: the retry escalates, Fix does not run, it stays pending" ($retry['escalated'] -eq '1|1|test trigger') `
    "escalations|Fix runs|pending: $($retry['escalated'])"
Assert-Case "a re-checked trigger held back: not kept, and not retried" ($retry['recheckedHeld'] -eq '[]|1|1|[]') `
    "pending|Fix runs|held lines|saved: $($retry['recheckedHeld'])"
Assert-Case "a re-checked trigger BLOCKED: not kept" ($retry['recheckedBlocked'] -eq '[]|0|1|1') "pending|Fix runs|blocked|run-by-hand line: $($retry['recheckedBlocked'])"
Assert-Case "a re-checked trigger DEFERRED: not kept" ($retry['recheckedDeferred'] -eq '[]|0|1|0') "pending|Fix runs|deferred|retried line: $($retry['recheckedDeferred'])"
Assert-Case "a one-shot trigger BLOCKED: kept and saved, not a try, and the line says it is tried again" `
    ($retry['oneShotBlocked'] -eq 'test trigger|0|0|1|1|test trigger') "pending|Fix runs|tries|blocked|retried line|saved: $($retry['oneShotBlocked'])"
Assert-Case "and retried once the user is away" ($retry['oneShotBlockedRetry'] -eq 'True|1|1') "ended|Fix runs|triggered: $($retry['oneShotBlockedRetry'])"
Assert-Case "a one-shot trigger DEFERRED: kept and saved, not a try, and the line says it is tried again" `
    ($retry['oneShotDeferred'] -eq 'test trigger|0|0|1|1|test trigger') "pending|Fix runs|tries|deferred|retried line|saved: $($retry['oneShotDeferred'])"
Assert-Case "and retried once startup is over" ($retry['oneShotDeferredRetry'] -eq 'True|1|1') "ended|Fix runs|triggered: $($retry['oneShotDeferredRetry'])"
Assert-Case "a re-checked trigger held back while a one-shot repair is pending: reason, time and tries kept, the latest hold moved" `
    ($retry['recheckedWhilePending'] -eq 'test trigger|True|True|1|2') "reason|first time kept|latest hold now|tries|Fix runs: $($retry['recheckedWhilePending'])"
Assert-Case "a retry started just before the 4 hours end: pending until then" ($retry['straddleBefore'] -eq 'test trigger|1') "got $($retry['straddleBefore'])"
Assert-Case "and held back after them: dropped once, at that hold, and not started again for another 4 hours" ($retry['straddle'] -eq '1|[]|2|1|0|[]|0') `
    "dropped|pending|Fix runs|retry warnings|dropped after 3 tries|saved|said to be retried: $($retry['straddle'])"
Assert-Case "a trigger of its own at 238 minutes, held back by Fix at 242: takes the repair over, not dropped, its 4 hours from 238, the second try" `
    ($retry['freshStraddle'] -eq '0|second trigger|4|2|2|second trigger|1|0,238') `
    "dropped|pending|minutes since the 4 hours started|tries|Fix runs|saved|retried line|warnings at: $($retry['freshStraddle'])"
Assert-Case "at 238 and BLOCKED at 242: the same, with no try" ($retry['blockedStraddle'] -eq '0|1|second trigger|1|1|second trigger|242|0|[second trigger]') `
    "dropped|blocked|pending|tries|Fix runs|saved|minute, then on the next poll dropped|pending: $($retry['blockedStraddle'])"
Assert-Case "at 238 and DEFERRED at 242: the same, with no try" ($retry['deferredStraddle'] -eq '0|1|second trigger|1|1|second trigger|1|0|[second trigger]') `
    "dropped|deferred|pending|tries|Fix runs|saved|tried-again line, then on the next poll dropped|pending: $($retry['deferredStraddle'])"
Assert-Case "at 238 and BLOCKED there: kept past the first hold's 4 hours, then dropped once, at 478" `
    ($retry['mergedKept'] -eq '0|second trigger|7|1|1|second trigger|1|478|[]') `
    "at 245 dropped|pending|minutes since the 4 hours started|tries|Fix runs|saved, then dropped|minute|pending: $($retry['mergedKept'])"
Assert-Case "at 200, tried with the warning and held back: kept past the first hold's 4 hours, then dropped once, at 440" `
    ($retry['mergedTried'] -eq '0|[second trigger]|2|1|440|[]|2') "at 241 dropped|pending|Fix runs, then dropped|minute|pending|Fix runs: $($retry['mergedTried'])"
Assert-Case "a restart right after a trigger of its own took the repair over: its new 4 hours are read back, it is kept and tried" `
    ($retry['restartMerged'] -eq 'second trigger|0|1|0|1|[]') "read|minutes since the 4 hours started|tries, then dropped|retried|pending: $($retry['restartMerged'])"
Assert-Case "a re-checked trigger Fix holds back on every run: 8 warnings in 4 hours, the wait doubled to 60 minutes" `
    ($retry['streak'] -eq '0,2,6,14,30,62,122,182|8|1|10|[]') "warnings at|Fix runs|60-minute line|fewest minutes between COOLDOWN lines|pending: $($retry['streak'])"
Assert-Case "and a run that is not held back starts the count again: the next hold waits 2 minutes" ($retry['streakReset'] -eq '10|0|True|1|3') `
    "Fix runs|held in a row after the run|latest hold none|held in a row|minutes of the cooldown gone: $($retry['streakReset'])"
Assert-Case "a hold 4 hours after the one before starts it again too" ($retry['streakDecay'] -eq '3|1|1|3') `
    "Fix runs|minutes gone after the second|held in a row|minutes gone after the third: $($retry['streakDecay'])"
Assert-Case "a one-shot trigger in the 32 minutes after the fifth run held back: kept and saved, and tried once the wait is over, at 62" `
    ($retry['oneShotInWait'] -eq 'log trigger|log trigger|1|1|5|62|1|[]') `
    "pending|saved|cooldown line|tried-again line|Fix runs|tried at|retried|pending after: $($retry['oneShotInWait'])"
Assert-Case "the same one-shot trigger again in that wait, after 10 minutes: COOLDOWN and tried again, its 4 hours from now, and it takes nothing over" `
    ($retry['sameAgain'] -eq '1|1|0|log trigger|0|5') `
    "COOLDOWN|tried-again line|takes over|pending|minutes since the 4 hours started|Fix runs: $($retry['sameAgain'])"
Assert-Case "one in the 2 minutes BLOCKED left for a re-checked trigger: kept, saved and tried" ($retry['oneShotInBlockedWait'] -eq 'log trigger|log trigger|1|[]') `
    "pending|saved|retried|pending after: $($retry['oneShotInBlockedWait'])"
Assert-Case "one in the minute DEFERRED left: kept, saved and tried" ($retry['oneShotInDeferredWait'] -eq 'log trigger|log trigger|1|[]') `
    "pending|saved|retried|pending after: $($retry['oneShotInDeferredWait'])"
Assert-Case "one in the cooldown after a repair that ran, a hold before it: skipped as before, not kept, not tried" ($retry['oneShotAfterRun'] -eq '[]|[]|1|0|1') `
    "pending|saved|cooldown line|tried-again line|Fix runs: $($retry['oneShotAfterRun'])"
Assert-Case "one after an escalation, though a hold came before it: not kept" ($retry['oneShotAfterEscalation'] -eq '1|[]|1|0') `
    "escalations|pending|cooldown line|Fix runs: $($retry['oneShotAfterEscalation'])"
Assert-Case "one after a cancel, though a hold came before it: not kept" ($retry['oneShotAfterCancel'] -eq '1|[]|1|0') `
    "cancelled|pending|cooldown line|Fix runs: $($retry['oneShotAfterCancel'])"
Assert-Case "a Fix run by hand that started the repair, in the wait five runs held back left: that wait ends, the next hold is the first and waits 2 minutes" `
    ($retry['manualAfterStreak'] -eq '6|1|3') "Fix runs|held in a row|minutes of the cooldown gone: $($retry['manualAfterStreak'])"
Assert-Case "a Diagnostic run by hand: the count and the 32-minute wait go on" ($retry['diagAfterStreak'] -eq '5|5|-16') `
    "Fix runs|held in a row|minutes of the cooldown gone: $($retry['diagAfterStreak'])"
Assert-Case "the run by hand 2 minutes before: the count is 0 and saved, the wait ends 5 minutes after that run, a log error in it is not kept" `
    ($retry['manualCooldown'] -eq '5|0|0|3|[]') "Fix runs|held in a row|saved|minutes left|pending: $($retry['manualCooldown'])"
Assert-Case "and it never makes a wait longer: a run by hand after a 2-minute wait, a trigger 1 minute after it gets through" `
    ($retry['manualNoLonger'] -eq '2|1') "Fix runs|held in a row: $($retry['manualNoLonger'])"
Assert-Case "BLOCKED every 2 minutes for 30 minutes: its lines, and COOLDOWN's, 3 times each" ($retry['blockedEvery'] -eq '3|3|3|0') `
    "blocked|run-by-hand|cooldown|Fix runs: $($retry['blockedEvery'])"
Assert-Case "Fix run by hand, holding its lock at the retry: SKIPPED once, with what becomes of the repair, which is kept and not ended" `
    ($retry['manualLock'] -eq '1|1|0|[test trigger]|[test trigger]') "skipped|kept line|ended|pending|saved: $($retry['manualLock'])"
Assert-Case "a run that holds the lock and turns out to be Diagnostic does not end it: the retry runs once the lock is gone" `
    ($retry['manualLockAfter'] -eq '1|0|1|2|[]') "skipped|ended|retry warnings|Fix runs|pending: $($retry['manualLockAfter'])"
Assert-Case "one that gets as far as the repair ends it, once the lock is gone, said once and saved" `
    ($retry['manualLockRepaired'] -eq '1|1|0|1|[]|[]') "skipped|ended|warnings|Fix runs|pending|saved: $($retry['manualLockRepaired'])"
Assert-Case "a second Fix run that takes the lock after a retry found it gone: said again, once" ($retry['lockTwice'] -eq '2|2|[test trigger]|1') `
    "skipped|kept line|pending|Fix runs: $($retry['lockTwice'])"
Assert-Case "the same after a trigger found it gone" ($retry['lockTwiceAuto'] -eq '2|2|third trigger|1') `
    "skipped|kept line|pending|Fix runs: $($retry['lockTwiceAuto'])"
Assert-Case "a one-shot trigger of its own while the lock is held: merged, its 4 hours started again, SKIPPED still once" `
    ($retry['lockMerge'] -eq 'second trigger|True|1|1|second trigger|1') "reason|4 hours started again|skipped|kept line|saved|Fix runs: $($retry['lockMerge'])"
Assert-Case "a one-shot trigger while the lock is held, with nothing pending: kept, SKIPPED once, Fix not run" `
    ($retry['lockFresh'] -eq '1|1|[test trigger]|[test trigger]|0') "skipped|kept line|pending|saved|Fix runs: $($retry['lockFresh'])"
Assert-Case "and tried once the lock is gone, as that run left no log that it repaired" ($retry['lockFreshAfter'] -eq '1|1|[]') `
    "retry warnings|Fix runs|pending: $($retry['lockFreshAfter'])"
Assert-Case "a re-checked trigger while the lock is held: SKIPPED once in 3 minutes, and not kept" ($retry['lockRechecked'] -eq '1|0|[]') `
    "skipped|kept line|pending: $($retry['lockRechecked'])"
Assert-Case "nothing pending and the lock held for 60 minutes, a trigger on every poll: SKIPPED 6 times, once in 10 minutes" `
    ($retry['lockNoPending'] -eq '6|[]|0') "skipped|pending|Fix runs: $($retry['lockNoPending'])"
Assert-Case "a Fix log created after the hold: ended once, before the retry, and saved" ($retry['manualLog'] -eq '1|[]|1|1|[]') `
    "ended|pending|Fix runs|lines|saved: $($retry['manualLog'])"
Assert-Case "a Fix log written after the hold, from a run that started before it: ended" ($retry['manualOverlap'] -eq '1|[]|1') `
    "ended|pending|Fix runs: $($retry['manualOverlap'])"
Assert-Case "a Diagnostic run's log after the hold: not ended, and the retry runs" ($retry['diagLog'] -eq '0|1|2') "ended|retry warnings|Fix runs: $($retry['diagLog'])"
Assert-Case "a Cancel's log after the hold: the same" ($retry['cancelLog'] -eq '0|1|2') "ended|retry warnings|Fix runs: $($retry['cancelLog'])"
Assert-Case "a No at Proceed? after the hold: the same" ($retry['noLog'] -eq '0|1|2') "ended|retry warnings|Fix runs: $($retry['noLog'])"
Assert-Case "the held-back run's own log, written before the hold: not a Fix run since" ($retry['ownLog'] -eq '0|True|2|1') `
    "ended|pending|Fix runs|retries held: $($retry['ownLog'])"
Assert-Case "and the retries' own logs: none is a Fix run since the latest hold" ($retry['ownRetryLogs'] -eq '0|3|1|3') `
    "ended|Fix runs|dropped after 3 tries|logs: $($retry['ownRetryLogs'])"
Assert-Case "Claude Desktop started again since the hold: ended once, before the retry, and saved" ($retry['appRestart'] -eq '1|[]|1|[]') `
    "ended|pending|Fix runs|saved: $($retry['appRestart'])"
Assert-Case "a helper process started since, and one that has exited: not a restart" ($retry['helperRestart'] -eq '0|2|1') `
    "ended|Fix runs|retries held: $($retry['helperRestart'])"
Assert-Case "a user the warning disturbs: three warnings in 4 hours, then dropped once, and saved" ($retry['touch'] -eq '3|1|0|[]|3|[]') `
    "warnings|dropped after 3 tries|dropped after 4 hours|pending|Fix runs|saved: $($retry['touch']) (warnings at minute $($retry['touchAt']))"
Assert-Case "each retry Fix holds back doubles the wait: 2 minutes, then 4" ($retry['backoff'] -eq '2|2|3|1|[]') `
    "Fix runs after 2.1, 6.0 and 6.2 minutes|dropped|pending: $($retry['backoff'])"
Assert-Case "from a BLOCKED origin: 2, 4 and 8 minutes, three warnings, then dropped" ($retry['backoffBlocked'] -eq '1,1,2,2,3|3|1|[]') `
    "Fix runs after each step|warnings|dropped|pending: $($retry['backoffBlocked'])"
Assert-Case "watch-state.json: the pending repair reads back as written, in UTC" ($retry['roundTrip'] -eq 'test trigger|True|Utc|Utc|2|1') `
    "reason|times|kinds|tries|retries held: $($retry['roundTrip'])"
# An older monitor reads the fields it knows by name and ignores the rest,
# so the new ones go after them.
Assert-Case "and the new fields are written after the ones an older monitor reads" `
    ($retry['savedFields'] -eq 'LastFixTime,FixCount,FixHistory,LastStealthMarkerTs,PendingFixReason,PendingFixSince,PendingFixHeldAt,PendingFixTries,PendingFixRetriesHeld,FixHeldInARow,FixHeldLastTime,CooldownFromHold') `
    "got $($retry['savedFields'])"
Assert-Case "a file from before the runs held in a row were kept: none" ($retry['heldOld'] -eq '0 0001-01-01T00:00:00 none') "got $($retry['heldOld'])"
Assert-Case "a file with them: read, in UTC" ($retry['heldSet'] -eq '3 2026-11-22T10:05:00 Utc') "got $($retry['heldSet'])"
Assert-Case "ones that cannot be read: none" ($retry['heldBad'] -eq '0 0001-01-01T00:00:00 none') "got $($retry['heldBad'])"
Assert-Case "a negative count: none" ($retry['heldNeg'] -eq '0 0001-01-01T00:00:00 none') "got $($retry['heldNeg'])"
Assert-Case "written and read back as they were" ($retry['heldRoundTrip'] -eq '2|True|Utc') "count|time|kind: $($retry['heldRoundTrip'])"
Assert-Case "and the fields an older monitor reads are still there" ($retry['heldOlderReader'] -eq '2|0|[]') "count|history|pending: $($retry['heldOlderReader'])"
Assert-Case "whether a hold left the cooldown: written and read back" ($retry['cooldownRoundTrip'] -eq 'True|True') "read|written: $($retry['cooldownRoundTrip'])"
Assert-Case "a file from before it was kept: not" ($retry['cooldownOld'] -eq 'False') "got $($retry['cooldownOld'])"
Assert-Case "a file with it: read" ($retry['cooldownSet'] -eq 'True' -and $retry['cooldownNo'] -eq 'False') "true|false: $($retry['cooldownSet'])|$($retry['cooldownNo'])"
Assert-Case "one that cannot be read: not" ($retry['cooldownBad'] -eq 'False') "got $($retry['cooldownBad'])"
$stateNone = '[] 0001-01-01T00:00:00 0001-01-01T00:00:00 0 0 | 2026-11-22T10:00:00 1'
Assert-Case "a file from before 6.3.0: none pending, and the rest read as before" ($retry['stateOld'] -eq $stateNone) "got $($retry['stateOld'])"
Assert-Case "a file with none pending: none" ($retry['stateNone'] -eq $stateNone) "got $($retry['stateNone'])"
Assert-Case "a pending time that cannot be read: none pending, and the rest read" ($retry['stateBad'] -eq $stateNone) "got $($retry['stateBad'])"
Assert-Case "a file with one pending and no newer fields: read, with no latest hold and no tries" `
    ($retry['stateSet'] -eq '[test trigger] 2026-11-22T09:55:00 0001-01-01T00:00:00 0 0 | 2026-11-22T10:00:00 1') "got $($retry['stateSet'])"
# A reason with no time could never be dropped after 4 hours.
Assert-Case "a reason with no time: none pending" ($retry['stateNoTime'] -eq $stateNone) "got $($retry['stateNoTime'])"
Assert-Case "a file with every field: read" `
    ($retry['stateFull'] -eq '[test trigger] 2026-11-22T09:55:00 2026-11-22T10:05:00 2 1 | 2026-11-22T10:00:00 1') "got $($retry['stateFull'])"
Assert-Case "newer fields that cannot be read: none of them, and the repair still pending" `
    ($retry['stateBadNew'] -eq '[test trigger] 2026-11-22T09:55:00 0001-01-01T00:00:00 0 0 | 2026-11-22T10:00:00 1') "got $($retry['stateBadNew'])"
Assert-Case "negative counts read as none" `
    ($retry['stateNegTries'] -eq '[test trigger] 2026-11-22T09:55:00 2026-11-22T10:05:00 0 0 | 2026-11-22T10:00:00 1') "got $($retry['stateNegTries'])"
Assert-Case "newer fields with none pending: none" ($retry['stateNewNoPending'] -eq $stateNone) "got $($retry['stateNewNoPending'])"

# Which trigger is which kind. A one-shot one does not fire again for the
# same fault, so a hold keeps its repair; a re-checked one fires again while
# the fault lasts, so it is not kept.
$autoFixSites = [System.Collections.Generic.List[string]]::new()
foreach ($c in $watchAst.FindAll({ param($n)
            $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Invoke-AutoFix' }, $true)) {
    $autoFixSites.Add($c.Extent.Text)
}
$autoFixSites.Sort([System.StringComparer]::Ordinal)
$autoFixExpected = @(
    'Invoke-AutoFix -Reason "CoworkVMService stopped while Claude is running (confirmed)"'
    'Invoke-AutoFix -Reason "Event Log: $eventError" -OneShot'
    'Invoke-AutoFix -Reason "Log error: $logError" -OneShot'
    'Invoke-AutoFix -Reason $pending -Retry'
    'Invoke-AutoFix -Reason $vmcomputeIssue'
) -join ' | '
Assert-Case "the log error and Event Log triggers are one-shot; the service and vmcompute ones are re-checked" `
    (($autoFixSites -join ' | ') -eq $autoFixExpected) "calls: $($autoFixSites -join ' | ')"
Assert-Case "the main loop calls it on a poll where no check fired" `
    ($watchSrc -match '(?s)Write-WatchLog "TIME WARNING: \$timeIssue"\s*\}\s*(#[^\r\n]*\s*)*Invoke-HeldBackRetry\s*\}\s*\} catch \{')

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- the monitor: the quiet checks, and the held-back repair's helpers"

# Test-ClaudeCodeActive and Test-UserActivity, lifted from Watch with what
# they call faked. With -Silent, which the retry's quiet check passes every 2
# minutes while a repair waits, neither writes a line; without it they write
# as before. Until 6.3.0 there was no -Silent, and the retry asked only the
# Code check, so a user at the PC met BLOCKED on every retry.
$quietCodeDef = Get-AstFunction -Ast $watchAst -Name 'Test-ClaudeCodeActive'
$quietUserDef = Get-AstFunction -Ast $watchAst -Name 'Test-UserActivity'
$userFakes = @('Get-Process', 'Get-CimInstance', 'New-Object', 'Get-ClaudeVmLogAge', 'Start-Sleep', 'Test-ClaudeCodeActive',
    'Test-CoworkSessionActive', 'Write-WatchLog', 'Get-ClaudeProcessByName', 'Write-CheckError')
# The throttled error line, lifted: both checks say a failure through it.
$quietCheckDefs = @(@('Get-CheckErrorText', 'Test-CheckErrorDue', 'Write-CheckError') |
        ForEach-Object { Get-AstFunction -Ast $watchAst -Name $_ } | Where-Object { $_ })
$unfaked = @(if ($quietCodeDef -and $quietUserDef) {
        Get-UnfakedCommand -Asts @($quietCodeDef) -Faked $codeFakes
        Get-UnfakedCommand -Asts @($quietUserDef) -Faked $userFakes
    } else { '<not found>' })
Assert-Case "Test-ClaudeCodeActive and Test-UserActivity were found, and every command they call is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
$quiet = [ordered]@{}
if ($unfaked.Count -eq 0) {
    $quietCode = Invoke-Isolated -Name 'the quiet Code check' -Block {
        Set-StrictMode -Version Latest
        $script:ClaudeEnv  = [pscustomobject]@{ PackageFamilyName = 'Test_pfn'; ProcessName = 'claude' }
        $script:QuietLines = [System.Collections.ArrayList]::new()
        function Write-WatchLog { param([string]$Message) [void]$script:QuietLines.Add($Message) }
        function Get-ClaudeAppCliProcess {
            param($Environment)
            $null = $Environment
            if ($script:CliThrows) { throw 'the parents could not be read' }
            return @()
        }
        # One data root; a case sets the candidates' errors and log folders.
        function Get-ClaudeDataCandidates {
            param($PackageFamilyName)
            $null = $PackageFamilyName
            return [pscustomobject]@{ DataRoots = @([pscustomobject]@{ Path = 'C:\fake-root' }); LogDirs = @($script:CandLogDirs)
                Errors = @($script:CandErrors) }
        }
        # $script:SessionDir: 'none' is no claude-code-sessions folder
        # anywhere, 'denied' one that cannot be read, 'there' one with
        # nothing recent in it. $script:WalkError is the error a listing
        # writes ('denied' or 'gone'), as the real one does: not thrown,
        # written. Test-Path is what a tree before 6.3.1 asks; it answers
        # $false for a folder it could not read, as the cmdlet does
        # (measured 2026-11-23).
        function Test-Path { return ($script:SessionDir -eq 'there') }
        function Get-ClaudeItemOrNull {
            param([string]$LiteralPath)
            # The error Get-Item gives for a folder it may not open, as
            # Get-ClaudeItemOrNull passes it on (measured 2026-11-23).
            if ($script:SessionDir -eq 'denied') {
                throw [System.Management.Automation.ErrorRecord]::new(
                    [System.UnauthorizedAccessException]::new('Access is denied'),
                    'ItemExistsUnauthorizedAccessError,Microsoft.PowerShell.Commands.GetItemCommand',
                    [System.Management.Automation.ErrorCategory]::PermissionDenied, $LiteralPath)
            }
            if ($script:SessionDir -eq 'none') { return }
            return [pscustomobject]@{ PSIsContainer = $true }
        }
        function Get-ChildItem {
            [CmdletBinding()]
            param([string]$LiteralPath, [string]$Filter, [switch]$Recurse, [switch]$File)
            $null = $Filter, $Recurse, $File
            if ($script:WalkError -eq 'gone') {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [System.Management.Automation.ItemNotFoundException]::new("Cannot find path '$LiteralPath\x' because it does not exist."),
                        'PathNotFound', [System.Management.Automation.ErrorCategory]::ObjectNotFound, $LiteralPath))
            } elseif ($script:WalkError -eq 'denied') {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [System.UnauthorizedAccessException]::new("Access to the path '$LiteralPath' is denied."),
                        'DirUnauthorizedAccessError', [System.Management.Automation.ErrorCategory]::PermissionDenied, $LiteralPath))
            }
        }
        function Test-ClaudeCodeTranscriptRecent { param([int]$Minutes) $null = $Minutes; return $false }
        foreach ($d in (@($quietCheckDefs) + @($regionErrDefs))) { . ([scriptblock]::Create($d.Extent.Text)) }
        $script:CheckErrorsSaid = @{}
        $script:CandErrors  = @()
        $script:CandLogDirs = @()
        $script:SessionDir  = 'none'
        $script:WalkError   = ''
        . ([scriptblock]::Create($quietCodeDef.Extent.Text))
        $codeSilent = (Get-Command Test-ClaudeCodeActive).Parameters.ContainsKey('Silent')
        function Invoke-CodeCheck {
            param([switch]$Quiet)
            if ($Quiet -and $codeSilent) { return (Test-ClaudeCodeActive -Silent) }
            return (Test-ClaudeCodeActive)
        }
        $o = [ordered]@{}
        $script:CliThrows = $false
        $script:NoCodeSessionDirLogged = $false
        $r = Invoke-CodeCheck -Quiet
        $o['silentNoDir'] = "$r|$($script:QuietLines.Count)|$($script:NoCodeSessionDirLogged)"
        $r = Invoke-CodeCheck
        $r = Invoke-CodeCheck
        $o['loudNoDir'] = "$r|$($script:QuietLines.Count)|$($script:NoCodeSessionDirLogged)"
        $script:QuietLines.Clear()
        $script:CliThrows = $true
        $r = Invoke-CodeCheck -Quiet
        $o['silentCodeThrow'] = "$r|$($script:QuietLines.Count)"
        $r = Invoke-CodeCheck
        $o['loudCodeThrow'] = "$r|$($script:QuietLines.Count)"

        # Reads that fail, each alone, aloud: yes, and the line names what
        # could not be read. Until 6.3.1 each read as a folder with nothing
        # recent in it, and the answer was no.
        $script:CliThrows = $false
        function Invoke-UnreadCase {
            param([string[]]$CandErrors = @(), [string]$SessionDir = 'none', [string]$WalkError = '', [object[]]$LogDirs = @())
            $script:CandErrors  = $CandErrors
            $script:SessionDir  = $SessionDir
            $script:WalkError   = $WalkError
            $script:CandLogDirs = $LogDirs
            $script:CheckErrorsSaid = @{}
            $script:QuietLines.Clear()
            $r = Invoke-CodeCheck
            $said = @($script:QuietLines | Where-Object { $_ -match 'assuming a Code session IS active' }) -join ' / '
            $script:CandErrors  = @()
            $script:SessionDir  = 'none'
            $script:WalkError   = ''
            $script:CandLogDirs = @()
            return "$r|$said"
        }
        $recentLogs = @([pscustomobject]@{ Path = 'C:\fake-logs'; Newest = [datetime]::UtcNow })
        $o['unreadCand'] = Invoke-UnreadCase -CandErrors @("C:\fake-other could not be read: Access to the path 'C:\fake-other' is denied.")
        $o['unreadDir']  = Invoke-UnreadCase -SessionDir 'denied'
        $o['unreadWalk'] = Invoke-UnreadCase -SessionDir 'there' -WalkError 'denied'
        $o['goneWalk']   = Invoke-UnreadCase -SessionDir 'there' -WalkError 'gone'
        $o['unreadLogs'] = Invoke-UnreadCase -LogDirs $recentLogs -WalkError 'denied'
        $o['goneLogs']   = Invoke-UnreadCase -LogDirs $recentLogs -WalkError 'gone'
        $o
    }
    $quietUser = Invoke-Isolated -Name 'the quiet user check' -Block {
        Set-StrictMode -Version Latest
        $ProcessName = 'claude'
        $script:ClaudeLogDir         = 'C:\nowhere'
        $script:IsInteractiveSession = $false
        $script:Win32ActivityLoaded  = $false
        $script:QuietLines = [System.Collections.ArrayList]::new()
        function Write-WatchLog { param([string]$Message) [void]$script:QuietLines.Add($Message) }
        # One process, idle: the CPU check reads no change. $script:ProcRead
        # 'throws' is a list that cannot be read; 'cpuDenied' and 'cpuGone'
        # a process whose CPU time cannot be read again, running or exited.
        # Get-Process is what a tree before 6.3.1 asks: with
        # SilentlyContinue, a list that cannot be read is empty there.
        function New-UserProcess {
            $p = [pscustomobject]@{ Id = 4242; TotalProcessorTime = [timespan]::Zero; HasExited = ($script:ProcRead -eq 'cpuGone') }
            if ($script:ProcRead -like 'cpu*') { $p | Add-Member -MemberType ScriptMethod -Name Refresh -Value { throw 'Access is denied' } }
            else { $p | Add-Member -MemberType ScriptMethod -Name Refresh -Value { } }
            return $p
        }
        function Get-Process { if ($script:ProcRead -eq 'throws') { return }; return @(New-UserProcess) }
        function Get-ClaudeProcessByName {
            param([string]$Name)
            $null = $Name
            if ($script:ProcRead -eq 'throws') { throw 'Access is denied' }
            return , @(New-UserProcess)
        }
        function Get-CimInstance { return $null }
        function New-Object { throw 'no desktop here' }
        function Get-ClaudeVmLogAge {
            param($LogDir)
            $null = $LogDir
            if ($script:VmThrows) { throw 'the VM logs could not be read' }
            return $null
        }
        function Start-Sleep { param([int]$Seconds, [int]$Milliseconds) $null = $Seconds, $Milliseconds }
        function Test-ClaudeCodeActive { param([switch]$Silent) $script:CodeAskedSilent = [bool]$Silent; return [bool]$script:CodeAnswer }
        function Test-CoworkSessionActive { return $false }
        foreach ($d in $quietCheckDefs) { . ([scriptblock]::Create($d.Extent.Text)) }
        $script:CheckErrorsSaid = @{}
        $script:ProcRead   = 'ok'
        $script:CodeAnswer = $true
        . ([scriptblock]::Create($quietUserDef.Extent.Text))
        $userSilent = (Get-Command Test-UserActivity).Parameters.ContainsKey('Silent')
        function Invoke-UserCheck {
            param([switch]$Quiet)
            if ($Quiet -and $userSilent) { return (Test-UserActivity -Silent) }
            return (Test-UserActivity)
        }
        $o = [ordered]@{}
        $script:VmThrows = $false
        $script:CodeAskedSilent = $null
        $r = Invoke-UserCheck -Quiet
        $o['silentCode'] = "$r|$($script:CodeAskedSilent)"
        $r = Invoke-UserCheck
        $o['loudCode'] = "$r|$($script:CodeAskedSilent)"
        $script:VmThrows = $true
        $r = Invoke-UserCheck -Quiet
        $o['silentUserThrow'] = "$r|$($script:QuietLines.Count)"
        $r = Invoke-UserCheck
        $o['loudUserThrow'] = "$r|$($script:QuietLines.Count)"

        # Reads that fail, aloud, with the Code check saying no: yes, and
        # the line names the error. Until 6.3.1 a list that could not be
        # read was no claude running, and a CPU time that could not be read
        # again was an idle process: no.
        $script:VmThrows   = $false
        $script:CodeAnswer = $false
        foreach ($pr in @('ok', 'throws', 'cpuDenied', 'cpuGone')) {
            $script:ProcRead = $pr
            $script:CheckErrorsSaid = @{}
            $script:QuietLines.Clear()
            $r = Invoke-UserCheck
            $o["userRead_$pr"] = "$r|$(@($script:QuietLines) -join ' / ')"
        }
        $script:ProcRead   = 'ok'
        $script:CodeAnswer = $true
        $o
    }
    foreach ($part in @($quietCode, $quietUser)) {
        if ($null -ne $part) { foreach ($k in @($part.Keys)) { $quiet[$k] = $part[$k] } }
    }
}
Assert-Case "Test-ClaudeCodeActive -Silent: no line about the missing folder, and it is left for a call without it" ($quiet['silentNoDir'] -eq 'False|0|False') `
    "result|lines|said: $($quiet['silentNoDir'])"
Assert-Case "Test-ClaudeCodeActive without it: the line, once" ($quiet['loudNoDir'] -eq 'False|1|True') "result|lines|said: $($quiet['loudNoDir'])"
Assert-Case "Test-ClaudeCodeActive -Silent that fails: still yes, and no line" ($quiet['silentCodeThrow'] -eq 'True|0') "result|lines: $($quiet['silentCodeThrow'])"
Assert-Case "Test-ClaudeCodeActive without it that fails: yes, and the line" ($quiet['loudCodeThrow'] -eq 'True|1') "result|lines: $($quiet['loudCodeThrow'])"
Assert-Case "Test-UserActivity -Silent asks the Code check silently" ($quiet['silentCode'] -eq 'True|True') "result|asked silently: $($quiet['silentCode'])"
Assert-Case "Test-UserActivity without it asks it aloud" ($quiet['loudCode'] -eq 'True|False') "result|asked silently: $($quiet['loudCode'])"
Assert-Case "Test-UserActivity -Silent that fails: still yes, and no line" ($quiet['silentUserThrow'] -eq 'True|0') "result|lines: $($quiet['silentUserThrow'])"
Assert-Case "Test-UserActivity without it that fails: yes, and the line" ($quiet['loudUserThrow'] -eq 'True|1') "result|lines: $($quiet['loudUserThrow'])"
$quietTail = '; the same error is logged at most once an hour'
$codeFailed = 'True|Test-ClaudeCodeActive failed (no recent sign of a session, and what could not be read may hold one: '
Assert-Case "Test-ClaudeCodeActive, a candidate folder that could not be read: yes, and said, naming it" `
    ($quiet['unreadCand'] -eq ($codeFailed + "C:\fake-other could not be read: Access to the path 'C:\fake-other' is denied.); " +
        "assuming a Code session IS active$quietTail")) "got $($quiet['unreadCand'])"
Assert-Case "a claude-code-sessions folder that could not be read: yes, and said, naming it" `
    ($quiet['unreadDir'] -eq ($codeFailed + "C:\fake-root\claude-code-sessions could not be read: Access is denied); " +
        "assuming a Code session IS active$quietTail")) "got $($quiet['unreadDir'])"
Assert-Case "one whose walk could not list a folder: yes, and said, naming it" `
    ($quiet['unreadWalk'] -eq ($codeFailed + "C:\fake-root\claude-code-sessions could not be read in full: Access to the path " +
        "'C:\fake-root\claude-code-sessions' is denied.); assuming a Code session IS active$quietTail")) "got $($quiet['unreadWalk'])"
Assert-Case "one whose walk found a folder gone: no, and nothing said" ($quiet['goneWalk'] -eq 'False|') "got $($quiet['goneWalk'])"
Assert-Case "a recent log folder that could not be listed: yes, and said once, naming it" `
    ($quiet['unreadLogs'] -eq ($codeFailed + "C:\fake-logs could not be read: Access to the path 'C:\fake-logs' is denied.); " +
        "assuming a Code session IS active$quietTail")) "got $($quiet['unreadLogs'])"
Assert-Case "one gone when listed: no, and nothing said" ($quiet['goneLogs'] -eq 'False|') "got $($quiet['goneLogs'])"
Assert-Case "Test-UserActivity, every read made, nothing in use: no" ($quiet['userRead_ok'] -eq 'False|') "got $($quiet['userRead_ok'])"
Assert-Case "a process list that cannot be read: yes, and said, naming the error" `
    ($quiet['userRead_throws'] -eq "True|Test-UserActivity failed (Access is denied); assuming the user IS active$quietTail") `
    "got $($quiet['userRead_throws'])"
Assert-Case "a CPU time that cannot be read again: yes, and said, naming the error" `
    ($quiet['userRead_cpuDenied'] -match '^True\|Test-UserActivity failed \(.*Access is denied.*\); assuming the user IS active;') `
    "got $($quiet['userRead_cpuDenied'])"
Assert-Case "a process that has exited during the sample: not busy, no, and nothing said" ($quiet['userRead_cpuGone'] -eq 'False|') `
    "got $($quiet['userRead_cpuGone'])"

# Get-HeldBackRetryWait, Get-ClaudeAppStartTime and Get-FixRunAfter. The
# last reads a temp folder with the real Get-ChildItem, its files' times set
# by hand: only fix_*.log files that hold the line Fix logs once it begins
# the repair count, by the later of when each was created and when it was
# last written. Test-FixLogRepaired, which reads that line, is lifted when
# it is there: until 6.3.0 it was not, and any Fix log counted.
$helperNames = @('Get-HeldBackRetryWait', 'Get-ClaudeAppStartTime', 'Get-FixRunAfter')
$helperDefs = @($helperNames | ForEach-Object { Get-AstFunction -Ast $watchAst -Name $_ } | Where-Object { $_ })
$repairedDefs = @(Get-AstFunction -Ast $watchAst -Name 'Test-FixLogRepaired' | Where-Object { $_ })
# Get-ProcessStartUtc arrived in 6.3.1; lifted when it is there.
$repairedDefs += @(Get-AstFunction -Ast $watchAst -Name 'Get-ProcessStartUtc' | Where-Object { $_ })
$unfaked = @(if ($helperDefs.Count -eq $helperNames.Count) {
        Get-UnfakedCommand -Asts (@($helperDefs) + @($repairedDefs)) -Faked @('Get-ClaudeDesktopProcess', 'Get-ChildItem', 'Test-FixLogRepaired',
            'Write-CheckError', 'Get-ProcessStartUtc')
    } else { '<not found>' })
Assert-Case "the held-back repair's helpers were found, and every command they call is faked or reads a temp folder" ($unfaked.Count -eq 0) ($unfaked -join ', ')
$helpers = $null
if ($unfaked.Count -eq 0) {
    $helpers = Invoke-Isolated -Name "the held-back repair's helpers" -Block {
        Set-StrictMode -Version Latest
        $script:ClaudeEnv = [pscustomobject]@{ ProcessName = 'claude' }
        function Get-ClaudeDesktopProcess {
            param($Environment)
            $null = $Environment
            if ($script:ListThrows) { throw 'the processes could not be listed' }
            return @($script:Procs)
        }
        # What is said, kept: only the list that fails is said to have
        # failed here; no case reads a folder that is there and cannot be
        # read.
        $script:HelperSaid = [System.Collections.ArrayList]::new()
        function Write-CheckError {
            param([string]$Check, [string]$Text, $ErrorRecord)
            [void]$script:HelperSaid.Add("$Check|$($Text.Replace('{0}', [string]$ErrorRecord.Exception.Message))")
        }
        foreach ($d in (@($helperDefs) + @($repairedDefs))) { . ([scriptblock]::Create($d.Extent.Text)) }
        $FixRepairLine = if ($watchRepairMatch.Success) { $watchRepairMatch.Groups[1].Value } else { 'Repair under way' }
        $o = [ordered]@{}
        $o['waits'] = (@(0..7 | ForEach-Object { Get-HeldBackRetryWait -RetriesHeld $_ }) -join ',')

        $nowUtc = [datetime]::UtcNow
        $script:ListThrows = $false
        $script:Procs = @()
        $none = Get-ClaudeAppStartTime
        $exited = [pscustomobject]@{ Id = 3; HasExited = $true }
        $exited | Add-Member -MemberType ScriptProperty -Name StartTime -Value { throw 'the process has exited' }
        $script:Procs = @(
            [pscustomobject]@{ Id = 1; StartTime = $nowUtc.AddMinutes(-10).ToLocalTime() },
            [pscustomobject]@{ Id = 2; StartTime = $nowUtc.AddMinutes(-1).ToLocalTime() },
            $exited)
        $main = Get-ClaudeAppStartTime
        $script:ListThrows = $true
        $thrown = Get-ClaudeAppStartTime
        $mainAge  = if ($main) { [math]::Round(($nowUtc - $main).TotalMinutes, 1) } else { 'none' }
        $mainKind = if ($main) { $main.Kind } else { 'none' }
        $o['appStart'] = "$($null -eq $none)|$mainAge|$mainKind|$($null -eq $thrown)"
        $o['appStartSaid'] = @($script:HelperSaid) -join ' / '
        $script:HelperSaid.Clear()
        # A process still running whose start time cannot be read, beside a
        # helper started a minute ago: no answer, and said. Until 6.3.1 it
        # was passed over, and the helper's start was taken for the app's.
        $script:ListThrows = $false
        # HasExited reads as nothing for a process this monitor may not open,
        # as its getter throws Access is denied (measured on 5.1 for System,
        # csrss and lsass, 2026-11-22).
        $locked = [pscustomobject]@{ Id = 4 }
        $locked | Add-Member -MemberType ScriptProperty -Name HasExited -Value { throw 'Access is denied' }
        $locked | Add-Member -MemberType ScriptProperty -Name StartTime -Value { throw 'Access is denied' }
        $locked | Add-Member -MemberType ScriptMethod -Name get_StartTime -Value { throw 'Access is denied' }
        $script:Procs = @([pscustomobject]@{ Id = 2; StartTime = $nowUtc.AddMinutes(-1).ToLocalTime() }, $locked)
        $lockedStart = Get-ClaudeAppStartTime
        $o['appStartLocked'] = "$($null -eq $lockedStart)|$(@($script:HelperSaid) -join ' / ')"
        $script:HelperSaid.Clear()

        $FixLogDir = Join-Path $tempRoot 'fix-logs'
        $null = [System.IO.Directory]::CreateDirectory($FixLogDir)
        $t0 = [datetime]::new(2026, 11, 22, 10, 0, 0, [System.DateTimeKind]::Utc)
        # A log that holds the line, unless the case writes other text.
        function New-FixLogFile {
            param([string]$Name, [datetime]$Created, [datetime]$Written, [string]$Text = "[10:00:01] $FixRepairLine")
            $p = Join-Path $FixLogDir $Name
            [System.IO.File]::WriteAllText($p, $Text)
            [System.IO.File]::SetCreationTimeUtc($p, $Created)
            [System.IO.File]::SetLastWriteTimeUtc($p, $Written)
        }
        New-FixLogFile -Name 'fix_a.log' -Created $t0.AddMinutes(-10) -Written $t0.AddMinutes(-5)
        New-FixLogFile -Name 'fix_b_transcript.log' -Created $t0.AddMinutes(-20) -Written $t0.AddMinutes(3)
        New-FixLogFile -Name 'claude-env.json' -Created $t0.AddMinutes(10) -Written $t0.AddMinutes(10)
        $dir = Join-Path $FixLogDir 'fix_dir.log'
        $null = [System.IO.Directory]::CreateDirectory($dir)
        [System.IO.Directory]::SetCreationTimeUtc($dir, $t0.AddMinutes(20))
        [System.IO.Directory]::SetLastWriteTimeUtc($dir, $t0.AddMinutes(20))
        $inv = [System.Globalization.CultureInfo]::InvariantCulture
        function Format-RunAfter {
            param($Time)
            if ($null -eq $Time) { return 'none' }
            return "$($Time.ToString('HH:mm:ss', $inv)) $($Time.Kind)"
        }
        $o['runAfter0'] = Format-RunAfter (Get-FixRunAfter -After $t0)
        $o['runAfter5'] = Format-RunAfter (Get-FixRunAfter -After $t0.AddMinutes(5))
        New-FixLogFile -Name 'fix_d.log' -Created $t0.AddMinutes(8) -Written $t0.AddMinutes(4)
        $o['runAfterCreated'] = Format-RunAfter (Get-FixRunAfter -After $t0.AddMinutes(5))
        # Runs that changed nothing: a Diagnostic run's log, and a Cancel's
        # transcript. Neither holds the line.
        New-FixLogFile -Name 'fix_e.log' -Created $t0.AddMinutes(12) -Written $t0.AddMinutes(15) -Text '[10:12:00] Diagnostic complete, no changes were made'
        New-FixLogFile -Name 'fix_f_transcript.log' -Created $t0.AddMinutes(13) -Written $t0.AddMinutes(16) -Text "  Proceed? (Y/n): n`r`n[10:13:00] Cancelled by user"
        $o['runAfterNoRepair'] = Format-RunAfter (Get-FixRunAfter -After $t0.AddMinutes(9))
        New-FixLogFile -Name 'fix_g_transcript.log' -Created $t0.AddMinutes(20) -Written $t0.AddMinutes(25)
        $o['runAfterRepair'] = Format-RunAfter (Get-FixRunAfter -After $t0.AddMinutes(9))
        New-FixLogFile -Name 'fix_h.log' -Created $t0.AddMinutes(28) -Written $t0.AddMinutes(30) -Text '[10:28:00] Cancelled by user'
        $o['runAfterLater'] = Format-RunAfter (Get-FixRunAfter -After $t0.AddMinutes(9))
        # Test-FixLogRepaired on real files: one with the line, one without,
        # none at all, one in UTF-16 with its byte order mark, one Fix still
        # has open for writing, and one opened to share nothing.
        if (Get-Command -Name 'Test-FixLogRepaired' -CommandType Function -ErrorAction SilentlyContinue) {
            $u16 = Join-Path $FixLogDir 'fix_u16_transcript.log'
            [System.IO.File]::WriteAllText($u16, "Transcript started`r`n[10:00:01] $FixRepairLine`r`n", [System.Text.Encoding]::Unicode)
            $open = Join-Path $FixLogDir 'fix_open_transcript.log'
            $writer = [System.IO.File]::Open($open, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
            try {
                $bytes = [System.Text.Encoding]::ASCII.GetBytes("[10:00:01] $FixRepairLine`r`n")
                $writer.Write($bytes, 0, $bytes.Length)
                $writer.Flush()
                $whileOpen = Test-FixLogRepaired -Path $open
            } finally { $writer.Dispose() }
            $shut = Join-Path $FixLogDir 'fix_shut.log'
            [System.IO.File]::WriteAllText($shut, "[10:00:01] $FixRepairLine")
            $locker = [System.IO.File]::Open($shut, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
            try { $whileShut = Test-FixLogRepaired -Path $shut } finally { $locker.Dispose() }
            $o['repairedFiles'] = "$(Test-FixLogRepaired -Path (Join-Path $FixLogDir 'fix_g_transcript.log'))|" +
                "$(Test-FixLogRepaired -Path (Join-Path $FixLogDir 'fix_e.log'))|$(Test-FixLogRepaired -Path (Join-Path $FixLogDir 'fix_none.log'))|" +
                "$(Test-FixLogRepaired -Path $u16)|$whileOpen|$whileShut"
        } else {
            $o['repairedFiles'] = 'Test-FixLogRepaired not found'
        }
        $FixLogDir = Join-Path $tempRoot 'no-such-folder'
        $o['runAfterMissing'] = Format-RunAfter (Get-FixRunAfter -After $t0)
        $o['helperSaid'] = @($script:HelperSaid) -join ' / '
        $o
    }
}
if ($null -eq $helpers) { $helpers = [ordered]@{} }
Assert-Case "the wait after a hold: 2 minutes, doubled for each retry held back, 60 at most" ($helpers['waits'] -eq '2,4,8,16,32,60,60,60') "got $($helpers['waits'])"
Assert-Case "the app's start: none running is none; the oldest process, in UTC, past one that has exited; a list that fails is none" `
    ($helpers['appStart'] -eq 'True|10|Utc|True') "none|minutes ago|kind|failed: $($helpers['appStart'])"
Assert-Case "and the list that fails is said, naming the error; until 6.3.1 it was not" `
    ($helpers['appStartSaid'] -eq ('app start time|When Claude Desktop started could not be read (the processes could not be listed), ' +
        'so a held-back repair is not ended on it')) "got $($helpers['appStartSaid'])"
Assert-Case "the app's start, a process still running whose start time cannot be read: no answer, and said" `
    ($helpers['appStartLocked'] -like 'True|app start time|When Claude Desktop started could not be read (*Access is denied*), so a held-back repair is not ended on it') `
    "got $($helpers['appStartLocked'])"
Assert-Case "nothing else here is said to have failed" ($helpers['helperSaid'] -eq '') "got $($helpers['helperSaid'])"
Assert-Case "a Fix log written after the time, a transcript included: its write time" ($helpers['runAfter0'] -eq '10:03:00 Utc') "got $($helpers['runAfter0'])"
Assert-Case "none after it, a JSON file and a folder named like a log not counted: none" ($helpers['runAfter5'] -eq 'none') "got $($helpers['runAfter5'])"
Assert-Case "a Fix log created after the time: its creation time" ($helpers['runAfterCreated'] -eq '10:08:00 Utc') "got $($helpers['runAfterCreated'])"
Assert-Case "a Diagnostic run's log and a Cancel's transcript, neither with the line that the run began the repair: none" `
    ($helpers['runAfterNoRepair'] -eq 'none') "got $($helpers['runAfterNoRepair'])"
Assert-Case "a transcript with that line: its write time" ($helpers['runAfterRepair'] -eq '10:25:00 Utc') "got $($helpers['runAfterRepair'])"
Assert-Case "a later log without it does not count over it" ($helpers['runAfterLater'] -eq '10:25:00 Utc') "got $($helpers['runAfterLater'])"
Assert-Case "Test-FixLogRepaired on real files: with the line, without, missing, UTF-16, open for writing, opened to share nothing" `
    ($helpers['repairedFiles'] -eq 'True|False|False|True|True|False') "got $($helpers['repairedFiles'])"
Assert-Case "no log folder: none" ($helpers['runAfterMissing'] -eq 'none') "got $($helpers['runAfterMissing'])"

# A Fix log folder that is there and cannot be read: no run, and said, once
# in the hour for the same error. No folder is no run yet, and not said.
# Until 6.3.0 neither was said. Get-ChildItem is faked, writing the error the
# real one writes, as a non-terminating error.
$runAfterDef = Get-AstFunction -Ast $watchAst -Name 'Get-FixRunAfter'
$checkErrorNames = @('Get-CheckErrorText', 'Test-CheckErrorDue', 'Write-CheckError')
$checkErrorDefs = @($checkErrorNames | ForEach-Object { Get-AstFunction -Ast $watchAst -Name $_ } | Where-Object { $_ })
$unfaked = @(if ($runAfterDef) {
        Get-UnfakedCommand -Asts (@($runAfterDef) + $checkErrorDefs) -Faked (@('Get-ChildItem', 'Test-FixLogRepaired', 'Write-WatchLog') + $checkErrorNames)
    } else { '<not found>' })
Assert-Case "Get-FixRunAfter was found, and every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $unread = Invoke-Isolated -Name 'Get-FixRunAfter on a folder that cannot be read' -Block {
        Set-StrictMode -Version Latest
        function Get-ChildItem {
            [CmdletBinding()]
            param([string]$LiteralPath, [string]$Filter, [switch]$File)
            $null = $Filter, $File
            if ($script:DirAnswer -eq 'missing') {
                $ex = [System.Management.Automation.ItemNotFoundException]::new("Cannot find path '$LiteralPath' because it does not exist.")
                $id = 'PathNotFound'
                $cat = [System.Management.Automation.ErrorCategory]::ObjectNotFound
            } else {
                $ex = [System.UnauthorizedAccessException]::new("Access to the path '$LiteralPath' is denied.")
                $id = 'DirUnauthorizedAccessError'
                $cat = [System.Management.Automation.ErrorCategory]::PermissionDenied
            }
            $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new($ex, $id, $cat, $LiteralPath))
        }
        function Test-FixLogRepaired { param([string]$Path) $null = $Path; return $true }
        function Write-WatchLog { param([string]$Message) [void]$script:RunAfterLines.Add($Message) }
        foreach ($d in (@($runAfterDef) + $checkErrorDefs)) { . ([scriptblock]::Create($d.Extent.Text)) }
        $FixLogDir = 'C:\fix-logs-test'
        $script:CheckErrorsSaid = @{}
        $script:RunAfterLines = [System.Collections.ArrayList]::new()
        $o = [ordered]@{}
        $script:DirAnswer = 'missing'
        $missing = Get-FixRunAfter -After ([datetime]::UtcNow.AddHours(-1))
        $o['missing'] = "$($null -eq $missing)|$($script:RunAfterLines.Count)"
        $script:DirAnswer = 'denied'
        $first = Get-FixRunAfter -After ([datetime]::UtcNow.AddHours(-1))
        $second = Get-FixRunAfter -After ([datetime]::UtcNow.AddHours(-1))
        $o['denied'] = "$($null -eq $first)|$($null -eq $second)|$($script:RunAfterLines.Count)|$(@($script:RunAfterLines) -join ' / ')"
        $o
    }
    if ($unread) {
        Assert-Case "no Fix log folder: no run, and nothing said" ($unread['missing'] -eq 'True|0') "none|lines: $($unread['missing'])"
        Assert-Case "a Fix log folder that cannot be read: no run, said once in the hour, naming the folder and the error" `
            ($unread['denied'] -eq ("True|True|1|The Fix logs in C:\fix-logs-test could not be read (Access to the path 'C:\fix-logs-test' is denied.), " +
                'so a Fix run is not seen there, and a held-back repair stays pending; the same error is logged at most once an hour')) `
            "none|none|lines|line: $($unread['denied'])"
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- the monitor: whether Fix's lock is held, when it cannot be opened"

# Only a lock that does not exist is not held. One this monitor may not open,
# as when Fix runs elevated and the monitor does not, is held; one that
# cannot be asked at all is taken as held and said, once in the hour for the
# same error. Until 6.3.0 any error read as not held, and a repair could
# start on top of a running one. The Mutex call is swapped for a stand-in
# that answers as the case sets.
$lockDef = Get-AstFunction -Ast $watchAst -Name 'Test-FixLockHeld'
$openCall = '[System.Threading.Mutex]::OpenExisting($fixLockName)'
$lockText = if ($lockDef) { $lockDef.Extent.Text } else { '' }
$openSwaps = ([regex]::Matches($lockText, [regex]::Escape($openCall))).Count
Assert-Case "Test-FixLockHeld was found, and its one Mutex call swapped" ($null -ne $lockDef -and $openSwaps -eq 1) "swaps $openSwaps"
if ($lockDef -and $openSwaps -eq 1) {
    $lockText = $lockText.Replace($openCall, 'Open-TestMutex $fixLockName')
    $lockAst = [System.Management.Automation.Language.Parser]::ParseInput($lockText, [ref]$null, [ref]$null)
    $unfaked = @(Get-UnfakedCommand -Asts (@($lockAst) + $checkErrorDefs) -Faked (@('Open-TestMutex', 'Write-WatchLog') + $checkErrorNames))
    Assert-Case "every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
    if ($unfaked.Count -eq 0) {
        $lock = Invoke-Isolated -Name 'Test-FixLockHeld' -Block {
            Set-StrictMode -Version Latest
            function Open-TestMutex {
                param([string]$Name)
                [void]$script:LockAsked.Add($Name)
                switch ($script:LockAnswer) {
                    'held' { return ([pscustomobject]@{ Name = $Name } | Add-Member -MemberType ScriptMethod -Name Dispose -Value { } -PassThru) }
                    'none' { throw [System.Threading.WaitHandleCannotBeOpenedException]::new('No handle of the given name exists.') }
                    'denied' { throw [System.UnauthorizedAccessException]::new("Access to the path '$Name' is denied.") }
                    'invalid' { throw [System.IO.IOException]::new('The handle is invalid.') }
                }
            }
            function Write-WatchLog { param([string]$Message) [void]$script:LockLines.Add($Message) }
            foreach ($d in $checkErrorDefs) { . ([scriptblock]::Create($d.Extent.Text)) }
            . ([scriptblock]::Create($lockText))
            $script:CheckErrorsSaid = @{}
            $script:LockLines = [System.Collections.ArrayList]::new()
            $o = [ordered]@{}
            foreach ($c in @(@('held', 'held'), @('none', 'none'), @('denied', 'denied'), @('invalid', 'invalid'), @('invalidAgain', 'invalid'))) {
                $script:LockAnswer = $c[1]
                $script:LockAsked = [System.Collections.ArrayList]::new()
                $before = $script:LockLines.Count
                $held = Test-FixLockHeld
                $o[$c[0]] = "$held|$($script:LockAsked.Count)|$($script:LockLines.Count - $before)"
            }
            $o['line'] = (@($script:LockLines) -join ' / ')
            $o
        }
        if ($lock) {
            Assert-Case "a lock Fix holds: held" ($lock['held'] -eq 'True|1|0') "held|names asked|lines: $($lock['held'])"
            Assert-Case "no lock under either name: not held, and nothing said" ($lock['none'] -eq 'False|2|0') "held|names asked|lines: $($lock['none'])"
            Assert-Case "a lock this monitor may not open: held, and nothing said" ($lock['denied'] -eq 'True|1|0') "held|names asked|lines: $($lock['denied'])"
            Assert-Case "a lock that cannot be asked: taken as held, and said" ($lock['invalid'] -eq 'True|1|1') "held|names asked|lines: $($lock['invalid'])"
            Assert-Case "and the same error again in the hour: held, not said again" ($lock['invalidAgain'] -eq 'True|1|0') `
                "held|names asked|lines: $($lock['invalidAgain'])"
            Assert-Case "the line names the lock and the error" `
                ($lock['line'] -eq ('Whether Fix-ClaudeDesktop is running could not be told from its lock Global\ClaudeDesktopFix (The handle is invalid.); ' +
                    'taken as running; the same error is logged at most once an hour')) "got $($lock['line'])"
        }
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- the monitor: an elevation check that fails"

# Whether this monitor is elevated, read before it starts Fix. A check that
# fails says so, naming why, and Fix is not started. Until 6.3.0 it read as
# "not elevated", and sent the user to reinstall a task that was fine. The
# lines from "$watchIsAdmin = $false" to the end of the if that reads it are
# lifted from Invoke-AutoFix, with the identity call swapped for one that
# throws.
$adminFn = Get-AstFunction -Ast $watchAst -Name 'Invoke-AutoFix'
$adminFrom = $null
$adminIf = $null
if ($adminFn) {
    $adminFrom = $adminFn.Find({ param($n)
        $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Extent.Text -eq '$watchIsAdmin = $false' }, $true)
    $adminIf = $adminFn.Find({ param($n)
        $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -eq '-not $watchIsAdmin' }, $true)
}
$identityCall = '[Security.Principal.WindowsIdentity]::GetCurrent()'
$adminText = ''
if ($adminFrom -and $adminIf) {
    $adminText = $adminFrom.Extent.StartScriptPosition.GetFullScript().Substring($adminFrom.Extent.StartOffset,
        $adminIf.Extent.EndOffset - $adminFrom.Extent.StartOffset)
}
$identitySwaps = ([regex]::Matches($adminText, [regex]::Escape($identityCall))).Count
Assert-Case "the elevation check was found in Invoke-AutoFix, and its identity call swapped" ($identitySwaps -eq 1) "swaps $identitySwaps"
if ($identitySwaps -eq 1) {
    $adminText = $adminText.Replace($identityCall, '(Get-TestIdentity)')
    $adminAst = [System.Management.Automation.Language.Parser]::ParseInput($adminText, [ref]$null, [ref]$null)
    $unfaked = @(Get-UnfakedCommand -Asts (@($adminAst) + $checkErrorDefs) `
        -Faked (@('Get-TestIdentity', 'Write-WatchLog', 'Clear-HeldBackRepair', 'Save-WatchState') + $checkErrorNames))
    Assert-Case "every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
    if ($unfaked.Count -eq 0) {
        $admin = Invoke-Isolated -Name 'the elevation check' -Block {
            Set-StrictMode -Version Latest
            function Get-TestIdentity { throw 'Either a required impersonation level was not provided, or the provided impersonation level is invalid.' }
            function Write-WatchLog { param([string]$Message) [void]$script:AdminLines.Add($Message) }
            function Clear-HeldBackRepair { $script:PendingFixReason = $null; $script:AdminCleared++ }
            function Save-WatchState { $script:AdminSaves++ }
            foreach ($d in $checkErrorDefs) { . ([scriptblock]::Create($d.Extent.Text)) }
            $script:AdminLines = [System.Collections.ArrayList]::new()
            $script:AdminCleared = 0
            $script:AdminSaves = 0
            $script:FixCount = 2
            $script:PendingFixReason = 'test trigger'
            $after = & ([scriptblock]::Create($adminText + "`r`n'FELL THROUGH'"))
            "$($null -eq $after)|$($script:AdminLines.Count)|$(@($script:AdminLines) -join ' / ')|$($script:AdminCleared)|$($script:AdminSaves)"
        }
        Assert-Case "a check that fails: Fix is not started, the pending repair ends, and the line names the error" `
            ($admin -eq ('True|1|>>> AUTO-FIX #2 SKIPPED: whether this monitor is elevated could not be read (Either a required impersonation level ' +
                'was not provided, or the provided impersonation level is invalid.), so Fix was not started: unelevated, it would raise a UAC ' +
                'prompt with nobody to answer it. Run Fix-ClaudeDesktop manually.|1|1')) "stopped|lines|line|ended|saved: $admin"
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- the Event Log reads: their window, and a log that cannot be read"

# Get-WinEvent is faked: it records the filter it is given and answers with
# nothing, as the real one does for a log with nothing in the window (a
# NoMatchingEventsFound error), or with another error, both as the
# non-terminating errors the real one writes. The window is the UTC time as
# round-trip text, which ends in Z: Get-WinEvent reads a [datetime] as local
# time whatever its Kind, and until 6.3.0 both scripts gave it the local
# time, an hour off in the hour the clocks go back. A log with nothing in the
# window is no error; any other error is said, the monitor's once in the hour
# for the same error. Until 6.3.0 every read here used SilentlyContinue, and
# a log that could not be read looked like one with nothing in it.
$evWatchNames = @('Test-EventLogErrors', 'Get-MatchingWinEvent') + $checkErrorNames
$evWatchDefs = @($evWatchNames | ForEach-Object { Get-AstFunction -Ast $watchAst -Name $_ } | Where-Object { $_ })
$evFixNames = @('Test-RecentHcsErrors', 'Get-FixWinEvent')
$evFixDefs = @($evFixNames | ForEach-Object { Get-AstFunction -Ast $fixAst -Name $_ } | Where-Object { $_ })
$evFound = (@($evWatchDefs | Where-Object { $_.Name -eq 'Test-EventLogErrors' }).Count -eq 1 -and
    @($evFixDefs | Where-Object { $_.Name -eq 'Test-RecentHcsErrors' }).Count -eq 1)
$evFakes = @('Get-WinEvent', 'Write-WatchLog', 'Test-StartupGracePeriod', 'Log', 'Test-Path', 'Get-ChildItem', 'Get-Content',
    'Test-CoworkServiceLog', 'Get-Item', 'Get-ClaudeItemOrNull') + $evWatchNames + $evFixNames
$unfaked = @(if ($evFound) { Get-UnfakedCommand -Asts (@($evWatchDefs) + @($evFixDefs)) -Faked $evFakes -AllowDynamic @('$firstLine') } else { '<not found>' })
Assert-Case "Test-EventLogErrors and Fix's Test-RecentHcsErrors were found, and every command they call is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $ev = Invoke-Isolated -Name 'the Event Log reads' -Block {
        Set-StrictMode -Version Latest
        function Get-WinEvent {
            [CmdletBinding()]
            param([hashtable]$FilterHashtable, [int]$MaxEvents)
            $null = $MaxEvents
            [void]$script:EvFilters.Add($FilterHashtable)
            if ($script:EvAnswer -eq 'none') {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [System.Exception]::new('No events were found that match the specified selection criteria.'),
                        'NoMatchingEventsFound', [System.Management.Automation.ErrorCategory]::ObjectNotFound, $null))
            } else {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [System.Exception]::new('The RPC server is unavailable'),
                        'System.Diagnostics.Eventing.Reader.EventLogException', [System.Management.Automation.ErrorCategory]::NotSpecified, $null))
            }
        }
        function Write-WatchLog { param([string]$Message) [void]$script:EvLines.Add($Message) }
        function Log { param([string]$Message, [string]$Colour, [switch]$Indent) $null = $Colour, $Indent; [void]$script:EvLines.Add($Message) }
        function Test-StartupGracePeriod { return $false }
        # The Claude log folder: none in the Event Log cases. In the Claude
        # log cases it is X:\claude\logs, and $script:EvLogs holds its logs,
        # each a text or $null for one that cannot be read, or is 'denied'
        # for a folder that cannot be listed, or 'folderDenied' for one that
        # cannot be read, with the error Get-Item gives for a folder in one
        # that denies reading (measured 2026-11-23). Each fails the way the
        # cmdlet does, with an error that does not end the command. Fix asks
        # Get-Item since 6.3.1 (the region's Get-ClaudeItemOrNull), Test-Path
        # before, which answers $false for a folder it could not read.
        function Test-EvFolderDenied { return ($script:EvLogs -is [string] -and $script:EvLogs -eq 'folderDenied') }
        function Test-Path {
            param([string]$LiteralPath)
            return ($LiteralPath -eq 'X:\claude\logs' -and -not (Test-EvFolderDenied))
        }
        function Get-Item {
            [CmdletBinding()]
            param([string]$LiteralPath, [switch]$Force)
            $null = $Force
            if (Test-EvFolderDenied) {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [System.UnauthorizedAccessException]::new('Access is denied'),
                        'ItemExistsUnauthorizedAccessError,Microsoft.PowerShell.Commands.GetItemCommand',
                        [System.Management.Automation.ErrorCategory]::PermissionDenied, $LiteralPath))
                return
            }
            if ($LiteralPath -eq 'X:\claude\logs') { return [pscustomobject]@{ FullName = $LiteralPath; PSIsContainer = $true } }
            $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                    [System.Management.Automation.ItemNotFoundException]::new("Cannot find path '$LiteralPath' because it does not exist."),
                    'PathNotFound,Microsoft.PowerShell.Commands.GetItemCommand',
                    [System.Management.Automation.ErrorCategory]::ObjectNotFound, $LiteralPath))
        }
        function Get-ChildItem {
            [CmdletBinding()]
            param([string]$LiteralPath, [string]$Filter)
            $null = $Filter
            if ($script:EvLogs -is [string]) {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [System.UnauthorizedAccessException]::new("Access to the path '$LiteralPath' is denied."),
                        'DirUnauthorizedAccessError,Microsoft.PowerShell.Commands.GetChildItemCommand',
                        [System.Management.Automation.ErrorCategory]::PermissionDenied, $LiteralPath))
                return
            }
            foreach ($k in @($script:EvLogs.Keys)) {
                [pscustomobject]@{ Name = $k; FullName = "$LiteralPath\$k"; LastWriteTimeUtc = [datetime]::UtcNow.AddMinutes(-1) }
            }
        }
        function Get-Content {
            [CmdletBinding()]
            param([string]$LiteralPath, [int]$Tail)
            $null = $Tail
            $text = $script:EvLogs[$LiteralPath.Substring($LiteralPath.LastIndexOf('\') + 1)]
            if ($null -eq $text) {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [System.IO.IOException]::new("The process cannot access the file '$LiteralPath' because it is being used by another process."),
                        'GetContentReaderIOError,Microsoft.PowerShell.Commands.GetContentCommand',
                        [System.Management.Automation.ErrorCategory]::ReadError, $LiteralPath))
                return
            }
            return $text
        }
        function Test-CoworkServiceLog { param([int]$WindowSeconds, [switch]$Brief) $null = $WindowSeconds, $Brief; return $null }
        foreach ($d in (@($evWatchDefs) + @($evFixDefs))) { . ([scriptblock]::Create($d.Extent.Text)) }
        $PollInterval = 30
        $LegacyErrorPatterns = @('VirtioFS')
        $script:EventLogHitCount = 0
        $script:CheckErrorsSaid = @{}
        $script:ClaudeEnv = [pscustomobject]@{ LogDir = $null }
        $script:EvLogs = $null
        $inv = [System.Globalization.CultureInfo]::InvariantCulture
        # Each filter's window: Z and the seconds back it starts, to the
        # nearest 10, for round-trip UTC text; the type's name for anything
        # else.
        function Get-EvWindows {
            $now = [datetime]::UtcNow
            return (@($script:EvFilters | ForEach-Object {
                        $st = $_['StartTime']
                        if ($null -eq $st) { 'none' }
                        elseif ($st -is [string] -and $st.EndsWith('Z')) {
                            $t = [datetime]::Parse($st, $inv, [System.Globalization.DateTimeStyles]::RoundtripKind)
                            "Z$([math]::Round(($now - $t).TotalSeconds / 10) * 10)"
                        } else { $st.GetType().Name }
                    }) -join ',')
        }
        function Invoke-Ev {
            param([string]$Answer, [scriptblock]$Call)
            $script:EvAnswer = $Answer
            $script:EvFilters = [System.Collections.ArrayList]::new()
            $script:EvLines = [System.Collections.ArrayList]::new()
            $ret = & $Call
            return "$($null -eq $ret)|$(Get-EvWindows)|$($script:EvLines.Count)"
        }
        $o = [ordered]@{}
        $o['watchNone'] = Invoke-Ev 'none' { Test-EventLogErrors }
        $o['watchFail'] = Invoke-Ev 'fail' { Test-EventLogErrors }
        $o['watchFailLines'] = @($script:EvLines)
        $o['watchFailAgain'] = Invoke-Ev 'fail' { Test-EventLogErrors }
        $o['fixNone'] = Invoke-Ev 'none' { Test-RecentHcsErrors }
        $o['fixFail'] = Invoke-Ev 'fail' { Test-RecentHcsErrors }
        $o['fixFailLines'] = (@($script:EvLines) -join ' / ')
        # The Claude logs, with nothing in the Event Logs: one that cannot be
        # read, alone and before one with an HCS error in it, and a folder
        # that cannot be listed. Until 6.3.0 each was passed over without a
        # word, as if it held no HCS error.
        $script:ClaudeEnv = [pscustomobject]@{ LogDir = 'X:\claude\logs' }
        $script:EvLogs = [ordered]@{ 'main.log' = $null }
        $o['fixLogLocked'] = Invoke-Ev 'none' { Test-RecentHcsErrors }
        $o['fixLogLockedLines'] = (@($script:EvLines) -join ' / ')
        $script:EvLogs = [ordered]@{ 'main.log' = $null; 'cowork.log' = "a line`nHCS operation failed: 0x80370114" }
        $o['fixLogNext'] = Invoke-Ev 'none' { Test-RecentHcsErrors }
        $script:EvLogs = 'denied'
        $o['fixLogDir'] = Invoke-Ev 'none' { Test-RecentHcsErrors }
        $o['fixLogDirLines'] = (@($script:EvLines) -join ' / ')
        # A folder that could not be read, and one that is not there, whose
        # listing would find an HCS error. Until 6.3.1 Test-Path answered
        # $false for both, and the first was passed over without a word.
        $script:EvLogs = 'folderDenied'
        $o['fixLogFolder'] = Invoke-Ev 'none' { Test-RecentHcsErrors }
        $o['fixLogFolderLines'] = (@($script:EvLines) -join ' / ')
        $script:ClaudeEnv = [pscustomobject]@{ LogDir = 'X:\claude\gone' }
        $script:EvLogs = [ordered]@{ 'cowork.log' = 'HCS operation failed: 0x80370114' }
        $o['fixLogGone'] = Invoke-Ev 'none' { Test-RecentHcsErrors }
        $o
    }
    if ($ev) {
        Assert-Case "the monitor's four logs, nothing in the window: nothing found, each window the last 60 seconds as UTC text, nothing said" `
            ($ev['watchNone'] -eq 'True|Z60,Z60,Z60,Z60|0') "clean|windows|lines: $($ev['watchNone'])"
        Assert-Case "the four logs that cannot be read: nothing found, and one line for each" `
            ($ev['watchFail'] -eq 'True|Z60,Z60,Z60,Z60|4') "clean|windows|lines: $($ev['watchFail'])"
        $named = @('The Application log could not be checked for CoworkVMService errors', 'The Microsoft-Windows-Hyper-V-Worker-Admin log could not be checked',
            'The Microsoft-Windows-Hyper-V-VMMS-Admin log could not be checked', 'The Microsoft-Windows-Hyper-V-Compute-Admin log could not be checked')
        $lines = @($ev['watchFailLines'])
        $wrong = @(for ($k = 0; $k -lt $named.Count; $k++) {
                $want = "EVENT LOG: $($named[$k]) (The RPC server is unavailable); the same error is logged at most once an hour"
                if ($lines.Count -le $k -or $lines[$k] -ne $want) { "$k" }
            })
        Assert-Case "each line names its log and the error" ($wrong.Count -eq 0) ($lines -join ' / ')
        Assert-Case "and the same errors again in the hour: not said again" ($ev['watchFailAgain'] -eq 'True|Z60,Z60,Z60,Z60|0') `
            "clean|windows|lines: $($ev['watchFailAgain'])"
        Assert-Case "Fix's HCS read, nothing in the window: clean, both windows the last 5 minutes as UTC text, nothing said" `
            ($ev['fixNone'] -eq 'True|Z300,Z300|0') "clean|windows|lines: $($ev['fixNone'])"
        Assert-Case "a log Fix cannot read: clean, and said, naming the error" `
            ($ev['fixFail'] -eq 'True|Z300|1' -and
                $ev['fixFailLines'] -eq '  HCS: the Hyper-V Compute event logs could not be read (The RPC server is unavailable), so recent HCS errors are not known') `
            "clean|windows|lines: $($ev['fixFail']); $($ev['fixFailLines'])"
        Assert-Case "a Claude log Fix cannot read: clean, and said, naming the log and the error" `
            ($ev['fixLogLocked'] -eq 'True|Z300,Z300|1' -and
                $ev['fixLogLockedLines'] -eq ("  HCS: the Claude log main.log could not be read (The process cannot access the file 'X:\claude\logs\main.log' " +
                    'because it is being used by another process.), so HCS errors in it are not known')) `
            "clean|windows|lines: $($ev['fixLogLocked']); $($ev['fixLogLockedLines'])"
        Assert-Case "and the next log is still read: the HCS error in it is found" ($ev['fixLogNext'] -eq 'False|Z300,Z300|1') `
            "clean|windows|lines: $($ev['fixLogNext'])"
        Assert-Case "a Claude log folder Fix cannot list: clean, and said, naming the folder and the error" `
            ($ev['fixLogDir'] -eq 'True|Z300,Z300|1' -and
                $ev['fixLogDirLines'] -eq ("  HCS: the Claude log folder X:\claude\logs could not be listed (Access to the path 'X:\claude\logs' is denied.), " +
                    'so HCS errors in its logs are not known')) `
            "clean|windows|lines: $($ev['fixLogDir']); $($ev['fixLogDirLines'])"
        Assert-Case "a Claude log folder Fix cannot read: clean, and said, naming the folder and the error" `
            ($ev['fixLogFolder'] -eq 'True|Z300,Z300|1' -and
                $ev['fixLogFolderLines'] -eq ("  HCS: the Claude log folder X:\claude\logs could not be listed (Access is denied), " +
                    'so HCS errors in its logs are not known')) `
            "clean|windows|lines: $($ev['fixLogFolder']); $($ev['fixLogFolderLines'])"
        Assert-Case "a Claude log folder that is not there: clean, not listed, and nothing said" ($ev['fixLogGone'] -eq 'True|Z300,Z300|0') `
            "clean|windows|lines: $($ev['fixLogGone'])"
    }
}

Write-Host "--- the monitor's checks: a check that fails is said, once in the hour"

# Each check is lifted with fakes and run twice with the same error. The
# first run says so, naming the error, and the second, in the same hour, does
# not; and the check returns what it returns when it finds nothing wrong, so
# the poll goes on. Until 6.3.0 each of these passed its failure over without
# a word, and a check that could not run looked like one that found nothing
# wrong. Each line is the check's own, so each case also says which check
# failed.
$ckNames = @('Test-WorkspaceStorage', 'Test-VmLogStaleness', 'Test-VmcomputeHealth', 'Test-HcsStateHealth',
    'Get-MatchingWinEvent', 'Save-WatchState', 'Import-WatchState', 'Test-LogsForErrors', 'Find-LogErrorPattern')
$ckDefs = @($ckNames | ForEach-Object { Get-AstFunction -Ast $watchAst -Name $_ } | Where-Object { $_ })
$ckDefNames = @($ckDefs | ForEach-Object { $_.Name })
# The shared region's helpers the checks read through since 6.3.1, lifted
# when they are there, so this also runs, and fails for the right reason,
# against the code from before them.
$ckRegionNames = @('Test-ClaudeNotFoundError', 'Get-ClaudeServiceOrNull', 'Get-ClaudeProcessByName', 'Get-ClaudeItemOrNull')
$ckRegionDefs = @($ckRegionNames | ForEach-Object { Get-AstFunction -Ast $watchAst -Name $_ } | Where-Object { $_ })
# Check 4b's own throttle since 6.3.1, lifted when it is there, so this also
# runs, and fails for the right reason, against the code from before it.
$ck4bNames = @('Test-StorageWarningDue')
$ck4bDefs = @($ck4bNames | ForEach-Object { Get-AstFunction -Ast $watchAst -Name $_ } | Where-Object { $_ })
$ckPatternMatch = [regex]::Match($watchSrc, '(?m)^\$LogTriggerFilePattern = ''([^'']+)''\s*$')
# Check 4b, the main loop's storage warning: the statement that asks
# Test-WorkspaceStorage and the if after it that says what it found, lifted
# from the main loop and run as each poll runs them. Until 6.3.1 the if wrote
# the warning on every poll while it lasted.
$ck4bAsts = @()
$ck4bCalls = @($watchAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and
            $n.GetCommandName() -eq 'Test-WorkspaceStorage' }, $true))
if ($ck4bCalls.Count -eq 1) {
    $ck4bStmt = $ck4bCalls[0]
    while ($ck4bStmt -and -not ($ck4bStmt.Parent -is [System.Management.Automation.Language.StatementBlockAst] -or
            $ck4bStmt.Parent -is [System.Management.Automation.Language.NamedBlockAst])) { $ck4bStmt = $ck4bStmt.Parent }
    if ($ck4bStmt) {
        $ck4bSiblings = @($ck4bStmt.Parent.Statements)
        $ck4bAt = [array]::IndexOf($ck4bSiblings, $ck4bStmt)
        if ($ck4bAt -ge 0 -and $ck4bAt + 1 -lt $ck4bSiblings.Count -and
            $ck4bSiblings[$ck4bAt + 1] -is [System.Management.Automation.Language.IfStatementAst]) {
            $ck4bAsts = @($ck4bStmt, $ck4bSiblings[$ck4bAt + 1])
        }
    }
}
$ck4bText = (@($ck4bAsts | ForEach-Object { $_.Extent.Text }) -join "`r`n")
$ck4bWhere = "Check 4b was not found in the main loop: $($ck4bCalls.Count) call(s) of Test-WorkspaceStorage"
$ckFound = ($ckDefs.Count -eq $ckNames.Count -and $ckPatternMatch.Success)
Assert-Case "the checks, and the Claude logs' file pattern, were found" $ckFound `
    "missing: $(@($ckNames | Where-Object { $ckDefNames -notcontains $_ }) -join ', '); pattern $($ckPatternMatch.Success)"
# ConvertTo-Json and New-Object are the real ones: Save-WatchState writes its
# file in a temp folder, and Test-LogsForErrors makes a set with New-Object.
$ckFakes = @('Write-WatchLog', 'Test-StartupGracePeriod', 'Test-Path', 'Get-ChildItem', 'Get-Item', 'Get-PSDrive',
    'Get-ClaudeVmLogAge', 'Test-CoworkSessionActive', 'Get-Process', 'Get-HcsListCached', 'Get-CoworkHcsGuids',
    'Get-WinEvent', 'Get-Service', 'Test-FixLockHeld', 'Restart-Service', 'Start-Sleep', 'Get-Content',
    'ConvertFrom-Json', 'ConvertFrom-WatchStateTime', 'ConvertTo-Json', 'New-Object') + $ckNames + $checkErrorNames
$unfaked = @(if ($ckFound) { Get-UnfakedCommand -Asts (@($ckDefs) + $checkErrorDefs + $ckRegionDefs + $ck4bDefs + $ck4bAsts) -Faked ($ckFakes + $ckRegionNames + $ck4bNames) } else { '<not found>' })
Assert-Case "every command they call is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $ckTriggerPattern = $ckPatternMatch.Groups[1].Value
    $ckLogDir = Join-Path $tempRoot 'check-logs'
    $null = New-Item -ItemType Directory -Path $ckLogDir -Force
    $ckBadState = Join-Path $ckLogDir 'no-such-folder\watch-state.json'
    $ckLogPath = Join-Path $ckLogDir 'main.log'
    $ck = Invoke-Isolated -Name "the monitor's checks" -Block {
        Set-StrictMode -Version Latest
        $script:CkTripped = [System.Collections.ArrayList]::new()
        function Write-WatchLog { param([string]$Message) [void]$script:CkLines.Add($Message) }
        function Test-StartupGracePeriod { return $false }
        function Test-CoworkSessionActive { return $script:CkActive }
        # The VM cache folder is there unless CkCache says 'denied' (it
        # could not be read) or 'missing'; Test-Path answers $false for both,
        # as the cmdlet does.
        function Test-Path {
            param([string]$Path)
            if ($Path -eq 'V:\ClaudeVmCache' -and $script:CkCache -ne 'there') { return $false }
            return $script:CkPathThere
        }
        # The Claude log folder's files, from .NET; the VM cache folder is
        # empty. CkLogList 'denied': the folder cannot be listed; 'missing':
        # it is not there, with the error Get-ChildItem gives for that
        # (measured 2026-11-22). The same for the VM cache folder, CkCache
        # 'missing'.
        function Get-ChildItem {
            [CmdletBinding()]
            param([Parameter(Position = 0)][string]$Path, [string]$LiteralPath, [string]$Filter, [switch]$Force, [switch]$File)
            $null = $Filter, $Force, $File
            $p = if ($LiteralPath) { $LiteralPath } else { $Path }
            # CkCache 'unlisted': the VM cache folder is there and cannot be
            # listed, with the error Get-ChildItem gives for that (measured
            # 2026-11-23, Case A), which names the folder.
            if ($p -eq 'V:\ClaudeVmCache' -and $script:CkCache -eq 'unlisted') {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [System.UnauthorizedAccessException]::new("Access to the path '$p' is denied."),
                        'DirUnauthorizedAccessError,Microsoft.PowerShell.Commands.GetChildItemCommand',
                        [System.Management.Automation.ErrorCategory]::PermissionDenied, $p))
                return
            }
            if ($p -eq 'V:\ClaudeVmCache' -and $script:CkCache -eq 'missing') {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [System.Management.Automation.ItemNotFoundException]::new("Cannot find path '$p' because it does not exist."),
                        'PathNotFound', [System.Management.Automation.ErrorCategory]::ObjectNotFound, $p))
                return
            }
            if ($p -eq $ckLogDir) {
                if ($script:CkLogList -eq 'denied') {
                    $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                            [System.UnauthorizedAccessException]::new("Access to the path '$p' is denied."),
                            'DirUnauthorizedAccessError', [System.Management.Automation.ErrorCategory]::PermissionDenied, $p))
                    return
                }
                if ($script:CkLogList -eq 'missing') {
                    $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                            [System.Management.Automation.ItemNotFoundException]::new("Cannot find path '$p' because it does not exist."),
                            'PathNotFound', [System.Management.Automation.ErrorCategory]::ObjectNotFound, $p))
                    return
                }
                return @([System.IO.DirectoryInfo]::new($p).GetFiles('*.log'))
            }
            return @()
        }
        # The state file, asked for by -LiteralPath since 6.3.1
        # (Import-WatchState): not there, with the error Get-Item gives for
        # that (measured 2026-11-23), unless CkPathThere; this section reads
        # no state file that is there.
        function Get-Item {
            [CmdletBinding()]
            param([Parameter(Position = 0)][string]$Path, [string]$LiteralPath, [switch]$Force)
            $null = $Force
            # The VM cache folder, asked for by -LiteralPath since 6.3.1
            # (Test-WorkspaceStorage), with the errors Get-Item gives
            # (measured 2026-11-23).
            if ($LiteralPath -eq 'V:\ClaudeVmCache') {
                if ($script:CkCache -eq 'denied') {
                    $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                            [System.UnauthorizedAccessException]::new('Access is denied'),
                            'ItemExistsUnauthorizedAccessError,Microsoft.PowerShell.Commands.GetItemCommand',
                            [System.Management.Automation.ErrorCategory]::PermissionDenied, $LiteralPath))
                    return
                }
                if ($script:CkCache -eq 'missing') {
                    $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                            [System.Management.Automation.ItemNotFoundException]::new("Cannot find path '$LiteralPath' because it does not exist."),
                            'PathNotFound,Microsoft.PowerShell.Commands.GetItemCommand',
                            [System.Management.Automation.ErrorCategory]::ObjectNotFound, $LiteralPath))
                    return
                }
                return [pscustomobject]@{ PSDrive = [pscustomobject]@{ Name = 'V' } }
            }
            if ($LiteralPath) {
                if ($script:CkPathThere) { [void]$script:CkTripped.Add('Get-Item -LiteralPath'); throw 'TRIPWIRE: Get-Item -LiteralPath' }
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [System.Management.Automation.ItemNotFoundException]::new("Cannot find path '$LiteralPath' because it does not exist."),
                        'PathNotFound', [System.Management.Automation.ErrorCategory]::ObjectNotFound, $LiteralPath))
                return
            }
            if ($script:CkItem -eq 'denied') {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [System.UnauthorizedAccessException]::new('Access is denied'),
                        'ItemExistsUnauthorizedAccessError,Microsoft.PowerShell.Commands.GetItemCommand',
                        [System.Management.Automation.ErrorCategory]::PermissionDenied, $Path))
                return
            }
            return [pscustomobject]@{ PSDrive = [pscustomobject]@{ Name = 'V' } }
        }
        function Get-PSDrive {
            [CmdletBinding()]
            param([string]$Name)
            $null = $Name
            return [pscustomobject]@{ Free = $script:CkFree }
        }
        function Get-ClaudeVmLogAge {
            param([string]$LogDir)
            $null = $LogDir
            $script:CkAgeCalls++
            if ($script:CkAgeThrowsOn -contains $script:CkAgeCalls) { throw 'The VM log folder could not be read' }
            return $script:CkAge
        }
        # vmcompute, or claude, as Get-Process gives it. A count that cannot
        # be read reads as nothing through the property, as a real Process's
        # does (PowerShell passes a getter's error over), and throws through
        # the getter method. CkProcs 'throws': the list cannot be read, with
        # an error that is not the one for "not there"; 'none': not running,
        # with the error Get-Process gives for that (measured 2026-11-22).
        function Get-Process {
            [CmdletBinding()]
            param([string]$Name)
            if ($script:CkProcs -eq 'tripwire') { [void]$script:CkTripped.Add('Get-Process'); throw 'TRIPWIRE: Get-Process' }
            if ($script:CkProcs -eq 'throws') {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [System.ComponentModel.Win32Exception]::new(5, 'Access is denied'),
                        'CouldNotEnumerateProcesses', [System.Management.Automation.ErrorCategory]::NotSpecified, $Name))
                return
            }
            if ($script:CkProcs -eq 'none') {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [Microsoft.PowerShell.Commands.ProcessCommandException]::new("Cannot find a process with the name `"$Name`". Verify the process name and call the cmdlet again."),
                        'NoProcessFoundForGivenName', [System.Management.Automation.ErrorCategory]::ObjectNotFound, $Name))
                return
            }
            $p = [pscustomobject]@{ Id = 4321 }
            if ($script:CkProcs -eq 'unreadable') {
                $p | Add-Member -MemberType ScriptProperty -Name HandleCount -Value { throw 'Access is denied' }
                $p | Add-Member -MemberType ScriptMethod -Name get_HandleCount -Value { throw 'Access is denied' }
            } else {
                $p | Add-Member -MemberType ScriptProperty -Name HandleCount -Value { 500 }
                $p | Add-Member -MemberType ScriptMethod -Name get_HandleCount -Value { 500 }
            }
            return $p
        }
        function Get-HcsListCached {
            if ($script:CkHcs -eq 'throws') { throw 'hcsdiag could not be run' }
            return $null
        }
        function Get-WinEvent {
            [CmdletBinding()]
            param([hashtable]$FilterHashtable, [int]$MaxEvents)
            $null = $MaxEvents
            [void]$script:CkWindows.Add($FilterHashtable['StartTime'])
            # 'spike': 16 shutdown failures (0xC037010D) this minute.
            if ($script:CkEvents -eq 'spike') {
                return @(1..16 | ForEach-Object { [pscustomobject]@{ Message = 'The operation failed with 0xC037010D'; TimeCreated = [datetime]::Now } })
            }
            if ($script:CkEvents -eq 'none') {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [System.Exception]::new('No events were found that match the specified selection criteria.'),
                        'NoMatchingEventsFound', [System.Management.Automation.ErrorCategory]::ObjectNotFound, $null))
            } else {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [System.Exception]::new('The RPC server is unavailable'),
                        'System.Diagnostics.Eventing.Reader.EventLogException', [System.Management.Automation.ErrorCategory]::NotSpecified, $null))
            }
        }
        function Get-CoworkHcsGuids { [void]$script:CkTripped.Add('Get-CoworkHcsGuids'); throw 'TRIPWIRE: Get-CoworkHcsGuids' }
        # The Cowork service, as CkSvc sets it: 'denied' cannot be read,
        # with an error that is not the one for "not there"; 'none' is not
        # installed, with the error Get-Service gives for that (measured
        # 2026-11-22).
        function Get-Service {
            [CmdletBinding()]
            param([string]$Name)
            if ($script:CkSvc -eq 'tripwire') { [void]$script:CkTripped.Add('Get-Service'); throw 'TRIPWIRE: Get-Service' }
            if ($script:CkSvc -eq 'denied') {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [System.ComponentModel.Win32Exception]::new(5, 'Access is denied'),
                        'CouldNotGetServiceInfo', [System.Management.Automation.ErrorCategory]::PermissionDenied, $Name))
                return
            }
            if ($script:CkSvc -eq 'none') {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [Microsoft.PowerShell.Commands.ServiceCommandException]::new("Cannot find any service with service name '$Name'."),
                        'NoServiceFoundForGivenName', [System.Management.Automation.ErrorCategory]::ObjectNotFound, $Name))
                return
            }
            return [pscustomobject]@{ Name = $Name; Status = 'Running' }
        }
        function Test-FixLockHeld {
            if ($script:CkFix -eq 'no') { return $false }
            [void]$script:CkTripped.Add('Test-FixLockHeld'); throw 'TRIPWIRE: Test-FixLockHeld'
        }
        # A restart is counted, never made; refused, so the check says so.
        function Restart-Service {
            param([string]$Name, [switch]$Force)
            $null = $Force
            [void]$script:CkRestarts.Add($Name)
            throw 'restart refused here'
        }
        function Start-Sleep { [void]$script:CkTripped.Add('Start-Sleep'); throw 'TRIPWIRE: Start-Sleep' }
        function Get-Content { [void]$script:CkTripped.Add('Get-Content'); throw 'TRIPWIRE: Get-Content' }
        function ConvertFrom-Json { [void]$script:CkTripped.Add('ConvertFrom-Json'); throw 'TRIPWIRE: ConvertFrom-Json' }
        function ConvertFrom-WatchStateTime { [void]$script:CkTripped.Add('ConvertFrom-WatchStateTime'); throw 'TRIPWIRE: ConvertFrom-WatchStateTime' }
        foreach ($d in (@($ckDefs) + $checkErrorDefs + $ckRegionDefs + $ck4bDefs)) { . ([scriptblock]::Create($d.Extent.Text)) }
        $script:CheckErrorsSaid = @{}
        $script:StorageWarningsSaid = @{}
        $script:CkSvc = 'tripwire'
        $script:CkFix = 'tripwire'
        $script:CkLogList = 'ok'
        $script:CkCache = 'there'
        $script:CkRestarts = [System.Collections.ArrayList]::new()
        $script:ClaudeEnv = [pscustomobject]@{ VmCachePath = 'V:\ClaudeVmCache' }
        $script:ClaudeLogDir = $ckLogDir
        $ClaudeLogDir = $ckLogDir
        $LogTriggerFilePattern = $ckTriggerPattern
        $LegacyErrorPatterns = @('HCS operation failed')
        $ServiceName = 'CoworkVMService'
        $ProcessName = 'claude'
        $script:LogBaselines = @{}
        $script:VmLogStaleCount = 0
        $script:VmLogEverActive = $true
        $script:LastVmcomputeCheck = [datetime]::MinValue
        $script:VmcomputeLeakCount = 0
        $script:LastShutdownFailLogTime = [datetime]::MinValue
        $script:FixHistory = [System.Collections.ArrayList]::new()
        $script:LastFixTime = [datetime]::MinValue
        $script:FixCount = 0
        $script:LastStealthMarkerTs = [long]0
        $script:PendingFixReason = $null
        $script:PendingFixSince = [datetime]::MinValue
        $script:PendingFixHeldAt = [datetime]::MinValue
        $script:PendingFixTries = 0
        $script:PendingFixRetriesHeld = 0
        $script:FixHeldInARow = 0
        $script:FixHeldLastTime = [datetime]::MinValue
        $script:CooldownFromHold = $false
        $script:CkActive = $false
        $script:CkPathThere = $true
        $script:CkItem = 'ok'
        $script:CkFree = 100GB
        $script:CkAge = 400
        $script:CkAgeCalls = 0
        $script:CkAgeThrowsOn = @()
        $script:CkProcs = 'tripwire'
        $script:CkHcs = 'none'
        $script:CkEvents = 'none'
        $script:CkWindows = [System.Collections.ArrayList]::new()
        $script:CkLines = [System.Collections.ArrayList]::new()
        $inv = [System.Globalization.CultureInfo]::InvariantCulture
        # One call: what it returned ('null' for nothing) and how many lines
        # it said.
        function Invoke-Ck {
            param([scriptblock]$Call)
            $script:CkLines = [System.Collections.ArrayList]::new()
            $script:CkAgeCalls = 0
            $ret = & $Call
            $r = if ($null -eq $ret) { 'null' } else { [string]$ret }
            return "$r|$($script:CkLines.Count)"
        }
        function Get-CkLines { return (@($script:CkLines) -join ' / ') }
        $o = [ordered]@{}

        # The free space of the VM image's volume cannot be read.
        $script:CkItem = 'denied'
        $o['storage'] = Invoke-Ck { Test-WorkspaceStorage }
        $o['storageLine'] = Get-CkLines
        $o['storageAgain'] = Invoke-Ck { Test-WorkspaceStorage }
        $script:CkItem = 'ok'
        $o['storageOk'] = Invoke-Ck { Test-WorkspaceStorage }
        # The VM cache folder: one that could not be read, and one that is
        # not there. Until 6.3.1 Test-Path answered $false for the first,
        # which read as no VM cache, and nothing was said.
        $script:CkCache = 'denied'
        $o['storageCacheDenied'] = Invoke-Ck { Test-WorkspaceStorage }
        $script:CkCache = 'missing'
        $o['storageCacheGone'] = Invoke-Ck { Test-WorkspaceStorage }
        $script:CkCache = 'there'

        # Check 4b, as the main loop runs it on each poll: a storage warning
        # is said when it is first found or its text changes, and again at
        # most once an hour while the same text lasts. An hour passes here by
        # moving each time said an hour and a minute back. Until 6.3.1 it was
        # written on every poll, about 120 identical lines an hour, and the
        # line for a VM cache folder that could not be read did not name it.
        if ($ck4bText) {
            $ck4b = [scriptblock]::Create($ck4bText)
            function Step-CkHour {
                param([int]$Minutes = -61)
                foreach ($k in @($script:CheckErrorsSaid.Keys)) { $script:CheckErrorsSaid[$k] = $script:CheckErrorsSaid[$k].AddMinutes($Minutes) }
                foreach ($k in @($script:StorageWarningsSaid.Keys)) { $script:StorageWarningsSaid[$k] = $script:StorageWarningsSaid[$k].AddMinutes($Minutes) }
            }
            $ckSaidBefore = $script:CheckErrorsSaid.Clone()
            $script:CheckErrorsSaid = @{}
            $script:StorageWarningsSaid = @{}
            $script:CkCache = 'denied'
            $o['4bDenied'] = Invoke-Ck $ck4b
            $o['4bDeniedLine'] = Get-CkLines
            $o['4bDeniedAgain'] = Invoke-Ck $ck4b
            Step-CkHour
            $o['4bDeniedHour'] = Invoke-Ck $ck4b
            $o['4bDeniedHourLine'] = Get-CkLines
            # Its text changes: the folder can now be read but not listed.
            $script:CkCache = 'unlisted'
            $o['4bUnlisted'] = Invoke-Ck $ck4b
            $o['4bUnlistedLine'] = Get-CkLines
            $script:CkCache = 'there'
            # A volume low on space: 1.5 GB free, and then 1 GB.
            $script:CheckErrorsSaid = @{}
            $script:StorageWarningsSaid = @{}
            $script:CkFree = 1.5GB
            $o['4bLow'] = Invoke-Ck $ck4b
            $o['4bLowLine'] = Get-CkLines
            $o['4bLowAgain'] = Invoke-Ck $ck4b
            Step-CkHour
            $o['4bLowHour'] = Invoke-Ck $ck4b
            $o['4bLowHourLine'] = Get-CkLines
            $script:CkFree = 1GB
            $o['4bLowChanged'] = Invoke-Ck $ck4b
            $o['4bLowChangedLine'] = Get-CkLines
            # A volume that fills, from 1.99 GB free to none: each of its 21
            # texts (2 GB, 1.9 GB and so on down to 0 GB) is said on its first
            # poll, all in the same hour; then a fault of another kind, and a
            # free space that cannot be read, each said on its first poll too;
            # and a text said in this hour that comes back is not said again.
            # When the warnings shared the storage check's throttle
            # (Test-CheckErrorDue, key 'storage'), its cap of 3 different
            # lines an hour held back the fourth text and every text after it,
            # the ones nearest to full, the new fault and the read error, for
            # up to an hour.
            $script:CheckErrorsSaid = @{}
            $script:StorageWarningsSaid = @{}
            $ckFill = @(foreach ($free in (@(1.99GB) + @(foreach ($n in 19..0) { $n * 0.1GB }))) {
                    $script:CkFree = $free
                    $ckPoll = Invoke-Ck $ck4b
                    "$ckPoll $(Get-CkLines)"
                })
            $o['4bFill'] = $ckFill -join ' // '
            $script:CkCache = 'unlisted'
            $o['4bFillNewKind'] = Invoke-Ck $ck4b
            $o['4bFillNewKindLine'] = Get-CkLines
            # The 23rd different text in the hour: the VM cache folder not
            # read, with another error.
            $script:CkCache = 'denied'
            $o['4bFillThirdKind'] = Invoke-Ck $ck4b
            $o['4bFillThirdKindLine'] = Get-CkLines
            $script:CkCache = 'there'
            $script:CkItem = 'denied'
            $o['4bFillReadError'] = Invoke-Ck $ck4b
            $o['4bFillReadErrorLine'] = Get-CkLines
            $script:CkItem = 'ok'
            $script:CkFree = 1.9GB
            $o['4bFillBack'] = Invoke-Ck $ck4b
            # The hour runs from when a text was said, not from when it was
            # last seen: polls 31 and 59 minutes after it was said do not say
            # it again, and one 61 minutes after it does, though it was
            # polled in between.
            Step-CkHour -Minutes -31
            $o['4bAt31'] = Invoke-Ck $ck4b
            Step-CkHour -Minutes -28
            $o['4bAt59'] = Invoke-Ck $ck4b
            Step-CkHour -Minutes -2
            $o['4bAt61'] = Invoke-Ck $ck4b
            $o['4bAt61Line'] = Get-CkLines
            # A time said that is in the future, after the clock was set
            # back, counts as old: the warning is said again, here with the
            # time said 5 minutes ahead.
            Step-CkHour -Minutes 5
            $o['4bFuture'] = Invoke-Ck $ck4b
            $o['4bFutureLine'] = Get-CkLines
            # More free space, a text not said in this hour, is said at once,
            # as less is.
            $script:StorageWarningsSaid = @{}
            $script:CkFree = 1GB
            $null = Invoke-Ck $ck4b
            $script:CkFree = 1.5GB
            $o['4bUp'] = Invoke-Ck $ck4b
            $o['4bUpLine'] = Get-CkLines
            # Each text has its own hour: when the hour of one said 61
            # minutes ago runs out, one said 31 minutes ago is not said again,
            # and the one said 61 minutes ago is.
            $script:StorageWarningsSaid = @{}
            $script:CkFree = 1.5GB
            $null = Invoke-Ck $ck4b
            Step-CkHour -Minutes -30
            $script:CkFree = 1GB
            $null = Invoke-Ck $ck4b
            Step-CkHour -Minutes -31
            $o['4bOwnHour'] = Invoke-Ck $ck4b
            $script:CkFree = 1.5GB
            $o['4bOwnHourOld'] = Invoke-Ck $ck4b
            $o['4bOwnHourOldLine'] = Get-CkLines
            $script:CkFree = 100GB
            $script:CheckErrorsSaid = $ckSaidBefore
            $script:StorageWarningsSaid = @{}
        }

        # The VM log's age is read, and cannot be read the second time.
        $script:CkAgeThrowsOn = @(2)
        $o['stale'] = Invoke-Ck { Test-VmLogStaleness }
        $o['staleLine'] = Get-CkLines
        $o['staleAgain'] = Invoke-Ck { Test-VmLogStaleness }
        $script:CkAgeThrowsOn = @()
        # The first read fails: the count stays as it was. Until 6.3.1 that
        # read was outside the try, and the error ended the poll.
        $script:CheckErrorsSaid = @{}
        $script:VmLogStaleCount = 1
        $script:CkAgeThrowsOn = @(1)
        $o['staleFirst'] = try { Invoke-Ck { Test-VmLogStaleness } } catch { "threw: $($_.Exception.Message)" }
        $o['staleFirstLine'] = Get-CkLines
        $o['staleFirstCount'] = $script:VmLogStaleCount
        $o['staleFirstAgain'] = try { Invoke-Ck { Test-VmLogStaleness } } catch { "threw: $($_.Exception.Message)" }
        $script:CkAgeThrowsOn = @()
        $script:VmLogStaleCount = 0

        # vmcompute's handle count cannot be read; then the process list
        # cannot be read at all. The check runs once a minute, so its time
        # is put back before each call.
        # A run of one high count is kept through a check that reads no
        # count: until 6.3.1 that check set it back to 0.
        $script:CkProcs = 'unreadable'
        $script:VmcomputeLeakCount = 1
        $o['handles'] = Invoke-Ck { Test-VmcomputeHealth }
        $o['handlesLine'] = Get-CkLines
        $o['handlesLeak'] = $script:VmcomputeLeakCount
        $script:VmcomputeLeakCount = 0
        $script:LastVmcomputeCheck = [datetime]::MinValue
        $o['handlesAgain'] = Invoke-Ck { Test-VmcomputeHealth }
        $script:LastVmcomputeCheck = [datetime]::MinValue
        $script:CkProcs = 'throws'
        $o['vmcompute'] = Invoke-Ck { Test-VmcomputeHealth }
        $o['vmcomputeLine'] = Get-CkLines
        $script:LastVmcomputeCheck = [datetime]::MinValue
        $o['vmcomputeAgain'] = Invoke-Ck { Test-VmcomputeHealth }
        $script:LastVmcomputeCheck = [datetime]::MinValue
        $script:CkProcs = 'ok'
        $o['vmcomputeOk'] = Invoke-Ck { Test-VmcomputeHealth }
        $script:LastVmcomputeCheck = [datetime]::MinValue
        $script:CkProcs = 'none'
        $o['vmcomputeNone'] = Invoke-Ck { Test-VmcomputeHealth }
        $script:LastVmcomputeCheck = [datetime]::MinValue
        $script:CkProcs = 'tripwire'

        # A spike of shutdown failures, 16 in 5 minutes, restarts vmcompute
        # only when no session is active. Whether one is, read from the
        # Cowork service and the claude processes, that cannot be read holds
        # the restart back, said once an hour. Until 6.3.1 both were read
        # with SilentlyContinue, and a read that failed let the restart
        # through. The hourly line about the count is kept out of the way.
        $script:CkEvents = 'spike'
        $script:CkFix = 'no'
        foreach ($sc in @(@('svcDenied', 'denied', 'ok'), @('procsDenied', 'running', 'throws'), @('active', 'running', 'ok'), @('idle', 'none', 'none'))) {
            $script:CheckErrorsSaid = @{}
            $script:CkRestarts.Clear()
            $script:CkSvc = $sc[1]
            $script:CkProcs = $sc[2]
            $script:LastShutdownFailLogTime = [datetime]::UtcNow
            $o["spike_$($sc[0])"] = "$(Invoke-Ck { Test-HcsStateHealth })|$($script:CkRestarts.Count)|$(Get-CkLines)"
            $script:LastShutdownFailLogTime = [datetime]::UtcNow
            $o["spike_$($sc[0])_again"] = "$(Invoke-Ck { Test-HcsStateHealth })|$($script:CkRestarts.Count)"
        }
        $script:CkRestarts.Clear()
        $script:CheckErrorsSaid = @{}
        $script:CkEvents = 'none'
        $script:CkFix = 'tripwire'
        $script:CkSvc = 'tripwire'
        $script:CkProcs = 'tripwire'
        $script:LastShutdownFailLogTime = [datetime]::MinValue

        # The HCS list cannot be read, with a Cowork session on, so the
        # shutdown failures are not read; then, with none on, the
        # Hyper-V Compute log cannot be read, and then has nothing in it.
        $script:CkActive = $true
        $script:CkHcs = 'throws'
        $o['hcsCount'] = Invoke-Ck { Test-HcsStateHealth }
        $o['hcsCountLine'] = Get-CkLines
        $o['hcsCountAgain'] = Invoke-Ck { Test-HcsStateHealth }
        $script:CkActive = $false
        $script:CkHcs = 'none'
        $script:CkEvents = 'fail'
        $script:CkWindows.Clear()
        $o['hcsShut'] = Invoke-Ck { Test-HcsStateHealth }
        $o['hcsShutLine'] = Get-CkLines
        $o['hcsShutAgain'] = Invoke-Ck { Test-HcsStateHealth }
        $script:CkEvents = 'none'
        $o['hcsShutNone'] = Invoke-Ck { Test-HcsStateHealth }
        $now = [datetime]::UtcNow
        $o['hcsWindows'] = (@($script:CkWindows | ForEach-Object {
                    if ($_ -is [string] -and $_.EndsWith('Z')) {
                        $t = [datetime]::Parse($_, $inv, [System.Globalization.DateTimeStyles]::RoundtripKind)
                        "Z$([math]::Round(($now - $t).TotalSeconds / 10) * 10)"
                    } elseif ($null -eq $_) { 'none' } else { $_.GetType().Name }
                }) -join ',')

        # The state cannot be saved: its folder is not there. Then it can.
        $script:StatePath = $ckBadState
        $o['save'] = Invoke-Ck { Save-WatchState }
        $o['saveLine'] = Get-CkLines
        $o['saveAgain'] = Invoke-Ck { Save-WatchState }
        $script:StatePath = Join-Path $ckLogDir 'watch-state.json'
        $o['saveOk'] = "$(Invoke-Ck { Save-WatchState })|$([System.IO.File]::Exists($script:StatePath))"

        # No state file, and the VM logs' age cannot be read at start; then
        # it can, and the VM log was written to a minute and a half ago.
        $script:StatePath = Join-Path $ckLogDir 'none.json'
        $script:CkPathThere = $false
        $script:VmLogEverActive = $false
        $script:CkAgeThrowsOn = @(1)
        $o['seed'] = "$(Invoke-Ck { Import-WatchState })|$($script:VmLogEverActive)"
        $o['seedLine'] = Get-CkLines
        $o['seedAgain'] = Invoke-Ck { Import-WatchState }
        $script:CkAgeThrowsOn = @()
        $script:CkAge = 90
        $o['seedOk'] = "$(Invoke-Ck { Import-WatchState })|$($script:VmLogEverActive)"
        $script:CkPathThere = $true

        # main.log is read once, which sets where the next read starts. It
        # gains an error line and is then held open by another, which lets
        # nothing else open it; once that lets go, the line is read.
        [System.IO.File]::WriteAllText($ckLogPath, "2026-11-22 10:00:00 [info] [app] started`r`n")
        $o['logsFirst'] = Invoke-Ck { Test-LogsForErrors }
        [System.IO.File]::AppendAllText($ckLogPath, "2026-11-22 10:00:30 [error] [vm] HCS operation failed`r`n")
        $lock = [System.IO.File]::Open($ckLogPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        try {
            $o['logs'] = Invoke-Ck { Test-LogsForErrors }
            $o['logsLine'] = Get-CkLines
            $o['logsAgain'] = Invoke-Ck { Test-LogsForErrors }
        } finally { $lock.Close() }
        $o['logsAfter'] = Invoke-Ck { Test-LogsForErrors }

        # The folder cannot be listed: nothing found, said once, and the
        # baselines kept, so what a log gains meanwhile is read once it can
        # be listed. A folder that is not there: nothing found or said.
        # Until 6.3.1 the folder was listed with SilentlyContinue, and one
        # that could not be listed read as a folder with no logs, which
        # dropped every baseline.
        [System.IO.File]::AppendAllText($ckLogPath, "2026-11-22 10:01:00 [info] [app] quiet`r`n")
        $null = Invoke-Ck { Test-LogsForErrors }
        $kept = $script:LogBaselines.Count
        $script:CkLogList = 'denied'
        $o['listDenied'] = Invoke-Ck { Test-LogsForErrors }
        $o['listDeniedLine'] = Get-CkLines
        $o['listDeniedKept'] = "$kept|$($script:LogBaselines.Count)"
        [System.IO.File]::AppendAllText($ckLogPath, "2026-11-22 10:01:30 [error] [vm] HCS operation failed`r`n")
        $o['listDeniedAgain'] = Invoke-Ck { Test-LogsForErrors }
        $script:CkLogList = 'missing'
        $o['listMissing'] = "$(Invoke-Ck { Test-LogsForErrors })|$($script:LogBaselines.Count)"
        $script:CkLogList = 'ok'
        $o['listAfter'] = Invoke-Ck { Test-LogsForErrors }

        $o['tripped'] = @($script:CkTripped) -join ','
        $o
    }
    if ($ck) {
        $tail = '; the same error is logged at most once an hour'
        Assert-Case "storage, a free space that cannot be read: nothing found, and said, naming the error" `
            ($ck['storage'] -eq 'null|1' -and
                $ck['storageLine'] -eq ("STORAGE WARNING: The free space on the VM image's volume could not be read " +
                    "(Access is denied), so a full volume would not be seen$tail")) `
            "returned|lines: $($ck['storage']); $($ck['storageLine'])"
        Assert-Case "storage, the same error again in the hour: not said again" ($ck['storageAgain'] -eq 'null|0') $ck['storageAgain']
        Assert-Case "storage, read: nothing said" ($ck['storageOk'] -eq 'null|0') $ck['storageOk']
        Assert-Case "storage, a VM cache folder that cannot be read: reported, naming the folder and the error" `
            ($ck['storageCacheDenied'] -eq "The VM cache directory V:\ClaudeVmCache could not be read (Access is denied)|0") $ck['storageCacheDenied']
        Assert-Case "storage, no VM cache folder: nothing found, nothing said" ($ck['storageCacheGone'] -eq 'null|0') $ck['storageCacheGone']
        $warnTail = '; the same warning is logged at most once an hour'
        $deniedLine = "STORAGE WARNING: The VM cache directory V:\ClaudeVmCache could not be read (Access is denied)$warnTail"
        $unlistedLine = ("STORAGE WARNING: The VM cache directory V:\ClaudeVmCache could not be read " +
            "(Access to the path 'V:\ClaudeVmCache' is denied.)$warnTail")
        $lowLine = "STORAGE WARNING: Only 1.5 GB free on V: the VM image needs room to grow and mounts fail without it$warnTail"
        $lowerLine = "STORAGE WARNING: Only 1 GB free on V: the VM image needs room to grow and mounts fail without it$warnTail"
        $no4b = if ($ck4bText) { '' } else { "$ck4bWhere; " }
        Assert-Case "storage, Check 4b: a VM cache folder that cannot be read is said once, naming the folder" `
            ($ck['4bDenied'] -eq 'null|1' -and $ck['4bDeniedLine'] -eq $deniedLine) "$no4b$($ck['4bDenied']); $($ck['4bDeniedLine'])"
        Assert-Case "storage, Check 4b: and not again on a poll in the same hour" ($ck['4bDeniedAgain'] -eq 'null|0') "$no4b$($ck['4bDeniedAgain'])"
        Assert-Case "storage, Check 4b: and again an hour later, the same line" `
            ($ck['4bDeniedHour'] -eq 'null|1' -and $ck['4bDeniedHourLine'] -eq $deniedLine) "$no4b$($ck['4bDeniedHour']); $($ck['4bDeniedHourLine'])"
        Assert-Case "storage, Check 4b: a text that changes is said at once: the folder cannot be listed, naming it" `
            ($ck['4bUnlisted'] -eq 'null|1' -and $ck['4bUnlistedLine'] -eq $unlistedLine) "$no4b$($ck['4bUnlisted']); $($ck['4bUnlistedLine'])"
        Assert-Case "storage, Check 4b: a volume low on space is said once" `
            ($ck['4bLow'] -eq 'null|1' -and $ck['4bLowLine'] -eq $lowLine) "$no4b$($ck['4bLow']); $($ck['4bLowLine'])"
        Assert-Case "storage, Check 4b: and not again on a poll in the same hour, low on space" ($ck['4bLowAgain'] -eq 'null|0') "$no4b$($ck['4bLowAgain'])"
        Assert-Case "storage, Check 4b: and again an hour later, low on space, the same line" `
            ($ck['4bLowHour'] -eq 'null|1' -and $ck['4bLowHourLine'] -eq $lowLine) "$no4b$($ck['4bLowHour']); $($ck['4bLowHourLine'])"
        Assert-Case "storage, Check 4b: less space, a text that changes, is said at once" `
            ($ck['4bLowChanged'] -eq 'null|1' -and $ck['4bLowChangedLine'] -eq $lowerLine) "$no4b$($ck['4bLowChanged']); $($ck['4bLowChangedLine'])"
        # The 21 texts of a volume that fills, written out, not computed.
        $fillTexts = @('2', '1.9', '1.8', '1.7', '1.6', '1.5', '1.4', '1.3', '1.2', '1.1', '1',
            '0.9', '0.8', '0.7', '0.6', '0.5', '0.4', '0.3', '0.2', '0.1', '0')
        $fillWant = (@($fillTexts | ForEach-Object {
                    "null|1 STORAGE WARNING: Only $_ GB free on V: the VM image needs room to grow and mounts fail without it$warnTail" }) -join ' // ')
        $fillFirstLine = "STORAGE WARNING: Only 1.9 GB free on V: the VM image needs room to grow and mounts fail without it$warnTail"
        $readErrorLine = ("STORAGE WARNING: The free space on the VM image's volume could not be read " +
            "(Access is denied), so a full volume would not be seen$tail")
        Assert-Case "storage, Check 4b: a volume that fills is said at each new text, all 21 from 2 GB free to none, in the same hour" `
            ($fillTexts.Count -eq 21 -and $ck['4bFill'] -eq $fillWant) "$no4b$($ck['4bFill'])"
        Assert-Case "storage, Check 4b: after those 21, a fault of another kind is said on its first poll" `
            ($ck['4bFillNewKind'] -eq 'null|1' -and $ck['4bFillNewKindLine'] -eq $unlistedLine) "$no4b$($ck['4bFillNewKind']); $($ck['4bFillNewKindLine'])"
        Assert-Case "storage, Check 4b: and a 23rd different text in the hour, the VM cache folder not read with another error, is said on its first poll too" `
            ($ck['4bFillThirdKind'] -eq 'null|1' -and $ck['4bFillThirdKindLine'] -eq $deniedLine) "$no4b$($ck['4bFillThirdKind']); $($ck['4bFillThirdKindLine'])"
        Assert-Case "storage, Check 4b: and a free space that cannot be read is still said, not held back by the warnings" `
            ($ck['4bFillReadError'] -eq 'null|1' -and $ck['4bFillReadErrorLine'] -eq $readErrorLine) "$no4b$($ck['4bFillReadError']); $($ck['4bFillReadErrorLine'])"
        Assert-Case "storage, Check 4b: a text said in this hour that comes back is not said again" ($ck['4bFillBack'] -eq 'null|0') "$no4b$($ck['4bFillBack'])"
        Assert-Case "storage, Check 4b: nor 31 minutes after it was said" ($ck['4bAt31'] -eq 'null|0') "$no4b$($ck['4bAt31'])"
        Assert-Case "storage, Check 4b: nor 59 minutes after it was said" ($ck['4bAt59'] -eq 'null|0') "$no4b$($ck['4bAt59'])"
        Assert-Case "storage, Check 4b: and 61 minutes after it was said it is said again, though polled in between" `
            ($ck['4bAt61'] -eq 'null|1' -and $ck['4bAt61Line'] -eq $fillFirstLine) "$no4b$($ck['4bAt61']); $($ck['4bAt61Line'])"
        Assert-Case "storage, Check 4b: a time said that is in the future, after the clock was set back, counts as old, 5 minutes ahead" `
            ($ck['4bFuture'] -eq 'null|1' -and $ck['4bFutureLine'] -eq $fillFirstLine) "$no4b$($ck['4bFuture']); $($ck['4bFutureLine'])"
        Assert-Case "storage, Check 4b: more free space, a text not said in this hour, is said at once" `
            ($ck['4bUp'] -eq 'null|1' -and $ck['4bUpLine'] -eq $lowLine) "$no4b$($ck['4bUp']); $($ck['4bUpLine'])"
        Assert-Case "storage, Check 4b: each text has its own hour: one said 31 minutes ago is not said again when an older one's runs out" `
            ($ck['4bOwnHour'] -eq 'null|0') "$no4b$($ck['4bOwnHour'])"
        Assert-Case "storage, Check 4b: and the one said 61 minutes ago, its own hour over, is said again" `
            ($ck['4bOwnHourOld'] -eq 'null|1' -and $ck['4bOwnHourOldLine'] -eq $lowLine) "$no4b$($ck['4bOwnHourOld']); $($ck['4bOwnHourOldLine'])"
        Assert-Case "VM log staleness, an age that cannot be read: nothing found, and said, naming the error" `
            ($ck['stale'] -eq 'null|1' -and
                $ck['staleLine'] -eq "STALENESS WARNING: The VM log staleness check failed (The VM log folder could not be read)$tail") `
            "returned|lines: $($ck['stale']); $($ck['staleLine'])"
        Assert-Case "VM log staleness, the same error again in the hour: not said again" ($ck['staleAgain'] -eq 'null|0') $ck['staleAgain']
        Assert-Case "VM log staleness, the first read fails: nothing found, the count kept, and said, naming the error" `
            ($ck['staleFirst'] -eq 'null|1' -and $ck['staleFirstCount'] -eq 1 -and
                $ck['staleFirstLine'] -eq ('STALENESS WARNING: The VM logs could not be read (The VM log folder could not be read), ' +
                    "so whether one has gone stale is not known and nothing is done about it$tail")) `
            "returned|lines: $($ck['staleFirst']); count $($ck['staleFirstCount']); $($ck['staleFirstLine'])"
        Assert-Case "and not said again in the hour" ($ck['staleFirstAgain'] -eq 'null|0') $ck['staleFirstAgain']
        Assert-Case "vmcompute, a handle count that cannot be read: nothing found, the run of high counts kept, and said, naming the process and the error" `
            ($ck['handles'] -eq 'null|1' -and $ck['handlesLeak'] -eq 1 -and
                $ck['handlesLine'] -like "The handle count of vmcompute (PID 4321) could not be read (*Access is denied*), so a handle leak in it would not be seen$tail") `
            "returned|lines: $($ck['handles']); $($ck['handlesLine'])"
        Assert-Case "vmcompute, the same error again in the hour: not said again" ($ck['handlesAgain'] -eq 'null|0') $ck['handlesAgain']
        Assert-Case "vmcompute, a check that fails: nothing found, and said, naming the error" `
            ($ck['vmcompute'] -eq 'null|1' -and
                $ck['vmcomputeLine'] -eq "The vmcompute handle check failed (Access is denied)$tail") `
            "returned|lines: $($ck['vmcompute']); $($ck['vmcomputeLine'])"
        Assert-Case "vmcompute, the same error again in the hour: not said again" ($ck['vmcomputeAgain'] -eq 'null|0') $ck['vmcomputeAgain']
        Assert-Case "vmcompute, read: nothing said" ($ck['vmcomputeOk'] -eq 'null|0') $ck['vmcomputeOk']
        Assert-Case "vmcompute, not running: nothing found, nothing said" ($ck['vmcomputeNone'] -eq 'null|0') $ck['vmcomputeNone']
        $spikeIssue = 'HCS shutdown failure spike: 16 events in last hour'
        $spikeUnread = "HCS SPIKE: 16 failures in 5 min, but whether a session is active could not be read (Access is denied); skipping vmcompute restart$tail"
        Assert-Case "HCS spike, the Cowork service cannot be read: vmcompute not restarted, said once" `
            ($ck['spike_svcDenied'] -eq "$spikeIssue|1|0|$spikeUnread" -and $ck['spike_svcDenied_again'] -eq "$spikeIssue|0|0") `
            "$($ck['spike_svcDenied']); again $($ck['spike_svcDenied_again'])"
        Assert-Case "HCS spike, the claude processes cannot be listed: vmcompute not restarted, said once" `
            ($ck['spike_procsDenied'] -eq "$spikeIssue|1|0|$spikeUnread" -and $ck['spike_procsDenied_again'] -eq "$spikeIssue|0|0") `
            "$($ck['spike_procsDenied']); again $($ck['spike_procsDenied_again'])"
        Assert-Case "HCS spike, a session active: vmcompute not restarted, and said" `
            ($ck['spike_active'] -eq "$spikeIssue|1|0|HCS SPIKE: 16 failures in 5 min but active session detected, skipping vmcompute restart") $ck['spike_active']
        Assert-Case "HCS spike, no service and no app: vmcompute restarted" `
            ($ck['spike_idle'] -eq ("$spikeIssue|2|1|HCS SPIKE: 16 shutdown failures in 5 min, restarting vmcompute / " +
                'vmcompute restart failed: restart refused here')) $ck['spike_idle']
        Assert-Case "HCS, a list that cannot be read: nothing found, and said, naming the error" `
            ($ck['hcsCount'] -eq 'null|1' -and
                $ck['hcsCountLine'] -eq "HCS: The check for more than one cowork-vm compute system failed (hcsdiag could not be run)$tail") `
            "returned|lines: $($ck['hcsCount']); $($ck['hcsCountLine'])"
        Assert-Case "HCS, the same error again in the hour: not said again" ($ck['hcsCountAgain'] -eq 'null|0') $ck['hcsCountAgain']
        Assert-Case "HCS, a Hyper-V Compute log that cannot be read: nothing found, and said, naming the error" `
            ($ck['hcsShut'] -eq 'null|1' -and
                $ck['hcsShutLine'] -eq "HCS: The check for shutdown failures (0xC037010D) failed (The RPC server is unavailable)$tail") `
            "returned|lines: $($ck['hcsShut']); $($ck['hcsShutLine'])"
        Assert-Case "HCS, the same error again in the hour: not said again; nothing in the log: nothing said" `
            ($ck['hcsShutAgain'] -eq 'null|0' -and $ck['hcsShutNone'] -eq 'null|0') "again $($ck['hcsShutAgain']), none $($ck['hcsShutNone'])"
        Assert-Case "HCS, each read's window is the last hour as UTC text" ($ck['hcsWindows'] -eq 'Z3600,Z3600,Z3600') $ck['hcsWindows']
        $saveWant = '^STATE: The state could not be saved to ' + [regex]::Escape($ckBadState) +
            ' \(.*Could not find a part of the path.*\); a restart would lose the cooldown, the fix history and a held-back repair' +
            [regex]::Escape($tail) + '$'
        Assert-Case "the state, a save that fails: said, naming the file and the error" `
            ($ck['save'] -eq 'null|1' -and $ck['saveLine'] -match $saveWant) "returned|lines: $($ck['save']); $($ck['saveLine'])"
        Assert-Case "the state, the same error again in the hour: not said again; a save that works: written, nothing said" `
            ($ck['saveAgain'] -eq 'null|0' -and $ck['saveOk'] -eq 'null|0|True') "again $($ck['saveAgain']), ok $($ck['saveOk'])"
        Assert-Case "the start, the VM logs' age that cannot be read: said, naming the error, and staleness waits" `
            ($ck['seed'] -eq 'null|1|False' -and
                $ck['seedLine'] -eq ("The age of the VM logs could not be read at start (The VM log folder could not be read), " +
                    "so the VM log staleness check waits until it sees a VM log written to$tail")) `
            "returned|lines|ever active: $($ck['seed']); $($ck['seedLine'])"
        Assert-Case "the start, the same error again in the hour: not said again; read: the VM log was active" `
            ($ck['seedAgain'] -eq 'null|0' -and $ck['seedOk'] -eq 'null|0|True') "again $($ck['seedAgain']), ok $($ck['seedOk'])"
        $logsWant = '^LOG WARNING: The Claude log ' + [regex]::Escape($ckLogPath) +
            ' could not be read \(.*being used by another process.*\); what it gained since the last read is read again at the next poll' +
            [regex]::Escape($tail) + '$'
        Assert-Case "the Claude logs, a log that cannot be opened: nothing found, and said, naming the log and the error" `
            ($ck['logsFirst'] -eq 'null|0' -and $ck['logs'] -eq 'null|1' -and $ck['logsLine'] -match $logsWant) `
            "first $($ck['logsFirst']); returned|lines: $($ck['logs']); $($ck['logsLine'])"
        Assert-Case "the Claude logs, the same error again in the hour: not said again; once it opens, what it gained is read" `
            ($ck['logsAgain'] -eq 'null|0' -and $ck['logsAfter'] -eq 'HCS operation failed|0') "again $($ck['logsAgain']), after $($ck['logsAfter'])"
        Assert-Case "the Claude log folder that cannot be listed: nothing found, said once, and the baselines kept" `
            ($ck['listDenied'] -eq 'null|1' -and $ck['listDeniedKept'] -eq '1|1' -and $ck['listDeniedAgain'] -eq 'null|0' -and
                $ck['listDeniedLine'] -eq ("LOG WARNING: The Claude log folder $ckLogDir could not be listed (Access to the path '$ckLogDir' is denied.), " +
                    "so no log trigger can fire until it can$tail")) `
            "returned|lines: $($ck['listDenied']); kept $($ck['listDeniedKept']); again $($ck['listDeniedAgain']); $($ck['listDeniedLine'])"
        Assert-Case "a folder that is not there: nothing found, nothing said; once listed, what the log gained meanwhile is read" `
            ($ck['listMissing'] -eq 'null|0|1' -and $ck['listAfter'] -eq 'HCS operation failed|0') "missing $($ck['listMissing']); after $($ck['listAfter'])"
        Assert-Case "and nothing a check must not do here was done" ($ck['tripped'] -eq '') $ck['tripped']
    }
}

# The main script's own: the lock that keeps a second monitor out, the
# session read and what is said of it, and the catches of the poll and of the
# focus check before a repair, each lifted and run the same way. And the
# stale VM kill, which is a function since 6.3.0, is called from the main
# loop, once.
$topStatements = @($watchAst.EndBlock.Statements)
$mutexTry = @($topStatements | Where-Object { $_ -is [System.Management.Automation.Language.TryStatementAst] -and
        $_.Extent.Text.Contains('[System.Threading.Mutex]::new($false, $mutexName)') })
$sessionTry = @($topStatements | Where-Object { $_ -is [System.Management.Automation.Language.TryStatementAst] -and
        $_.Extent.Text.Contains('(Get-Process -Id $PID -ErrorAction Stop).SessionId') })
$sessionIf = @($topStatements | Where-Object { $_ -is [System.Management.Automation.Language.IfStatementAst] -and
        $_.Clauses[0].Item1.Extent.Text -eq '$script:SessionReadError' })
$pollCatch = @($watchAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.CatchClauseAst] -and
            $n.Body.Extent.Text.Contains("-Check 'poll'") }, $true))
$focusCatch = @($watchAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.CatchClauseAst] -and
            $n.Body.Extent.Text.Contains("-Check 'focus before repair'") }, $true))
$staleCalls = @($watchAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and
            $n.GetCommandName() -eq 'Stop-StaleCoworkVm' }, $true))
$staleInLoop = @($staleCalls | Where-Object {
        $up = $_.Parent
        $inFunction = $false
        $inLoop = $false
        while ($up) {
            if ($up -is [System.Management.Automation.Language.FunctionDefinitionAst]) { $inFunction = $true }
            if ($up -is [System.Management.Automation.Language.WhileStatementAst]) { $inLoop = $true }
            $up = $up.Parent
        }
        $inLoop -and -not $inFunction
    })
Assert-Case "the main loop calls Stop-StaleCoworkVm, once, and nothing else does" ($staleCalls.Count -eq 1 -and $staleInLoop.Count -eq 1) `
    "calls $($staleCalls.Count), from the main loop $($staleInLoop.Count)"

# The lock: its mutex, its wait and its exit are swapped for stand-ins. The
# wait is where a lock left by a monitor that ended without letting go of it
# is taken, which is no error.
$mutexCtor = '[System.Threading.Mutex]::new($false, $mutexName)'
$mutexWait = '$script:Mutex.WaitOne(0)'
$mutexText = if ($mutexTry.Count -eq 1) { $mutexTry[0].Extent.Text } else { '' }
$mutexSwaps = "$([regex]::Matches($mutexText, [regex]::Escape($mutexCtor)).Count)|" +
    "$([regex]::Matches($mutexText, [regex]::Escape($mutexWait)).Count)|$([regex]::Matches($mutexText, '\bexit 0\b').Count)"
$mutexText = $mutexText.Replace($mutexCtor, '(New-TestMutex)').Replace($mutexWait, '(Wait-TestMutex)').Replace('exit 0', "throw 'TRIPWIRE: exit'")
$mutexAst = [System.Management.Automation.Language.Parser]::ParseInput($mutexText, [ref]$null, [ref]$null)
$unfaked = @(Get-UnfakedCommand -Asts (@($mutexAst) + $checkErrorDefs) -Faked (@('New-TestMutex', 'Wait-TestMutex', 'Write-Host', 'Write-WatchLog') + $checkErrorNames))
$mutexReady = ($mutexTry.Count -eq 1 -and $mutexSwaps -eq '1|1|1' -and $unfaked.Count -eq 0)
Assert-Case "the single-instance lock was found, its mutex, its wait and its exit swapped, and every command it calls faked" $mutexReady `
    "found $($mutexTry.Count), swaps $mutexSwaps; $($unfaked -join ', ')"
if ($mutexReady) {
    $mx = Invoke-Isolated -Name 'the single-instance lock' -Block {
        Set-StrictMode -Version Latest
        function Write-WatchLog { param([string]$Message) [void]$script:MxLines.Add($Message) }
        function Write-Host { [void]$script:MxLines.Add("HOST: $args") }
        function New-TestMutex {
            if ($script:MxAnswer -eq 'io') { throw [System.IO.IOException]::new('The handle is invalid.') }
            return 'a mutex'
        }
        function Wait-TestMutex {
            if ($script:MxAnswer -eq 'abandoned') { throw [System.Threading.AbandonedMutexException]::new() }
            return $true
        }
        foreach ($d in $checkErrorDefs) { . ([scriptblock]::Create($d.Extent.Text)) }
        $script:CheckErrorsSaid = @{}
        $mutexName = 'Global\ClaudeHealthMonitor'
        $Quiet = $true
        $run = [scriptblock]::Create($mutexText)
        function Invoke-Mx {
            param([string]$Answer)
            $script:MxAnswer = $Answer
            $script:MxLines = [System.Collections.ArrayList]::new()
            $script:Mutex = $null
            . $run
            return "$($null -ne $script:Mutex)|$($script:MxLines.Count)"
        }
        $o = [ordered]@{}
        $o['io'] = Invoke-Mx 'io'
        $o['ioLine'] = @($script:MxLines) -join ' / '
        $o['ioAgain'] = Invoke-Mx 'io'
        $o['abandoned'] = Invoke-Mx 'abandoned'
        $o['taken'] = Invoke-Mx 'taken'
        $o
    }
    if ($mx) {
        Assert-Case "a lock that cannot be taken: monitoring goes on, said, naming the lock and the error" `
            ($mx['io'] -eq 'False|1' -and
                $mx['ioLine'] -eq ('The lock that keeps a second monitor from running (Global\ClaudeHealthMonitor) could not be taken ' +
                    '(The handle is invalid.); monitoring without it; the same error is logged at most once an hour')) `
            "lock|lines: $($mx['io']); $($mx['ioLine'])"
        Assert-Case "the same error again in the hour: not said again" ($mx['ioAgain'] -eq 'False|0') $mx['ioAgain']
        Assert-Case "a lock left by a monitor that ended without letting go of it: taken, and nothing said" ($mx['abandoned'] -eq 'True|0') $mx['abandoned']
        Assert-Case "a lock taken: nothing said" ($mx['taken'] -eq 'True|0') $mx['taken']
    }
}

# The session: the read, and the line after the banner that says it failed.
$sessionText = if ($sessionTry.Count -eq 1 -and $sessionIf.Count -eq 1) { $sessionTry[0].Extent.Text + "`r`n" + $sessionIf[0].Extent.Text } else { '' }
$sessionAst = [System.Management.Automation.Language.Parser]::ParseInput($sessionText, [ref]$null, [ref]$null)
$unfaked = @(Get-UnfakedCommand -Asts (@($sessionAst) + $checkErrorDefs) -Faked (@('Get-Process', 'Write-WatchLog') + $checkErrorNames))
$sessionReady = ($sessionTry.Count -eq 1 -and $sessionIf.Count -eq 1 -and $unfaked.Count -eq 0)
Assert-Case "the session read and its line were found, and every command they call is faked" $sessionReady `
    "read $($sessionTry.Count), line $($sessionIf.Count); $($unfaked -join ', ')"
if ($sessionReady) {
    $ss = Invoke-Isolated -Name 'the session read' -Block {
        Set-StrictMode -Version Latest
        function Write-WatchLog { param([string]$Message) [void]$script:SsLines.Add($Message) }
        function Get-Process {
            [CmdletBinding()]
            param([int]$Id)
            if ($script:SsAnswer -eq 'gone') {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [System.ArgumentException]::new("Cannot find a process with the process identifier $Id."),
                        'NoProcessFoundForGivenId,Microsoft.PowerShell.Commands.GetProcessCommand',
                        [System.Management.Automation.ErrorCategory]::ObjectNotFound, $Id))
                return
            }
            return [pscustomobject]@{ Id = $Id; SessionId = 1 }
        }
        foreach ($d in $checkErrorDefs) { . ([scriptblock]::Create($d.Extent.Text)) }
        $script:CheckErrorsSaid = @{}
        $run = [scriptblock]::Create($sessionText)
        function Invoke-Ss {
            param([string]$Answer)
            $script:SsAnswer = $Answer
            $script:SsLines = [System.Collections.ArrayList]::new()
            $script:SessionReadError = $null
            $script:IsInteractiveSession = $false
            . $run
            return "$($script:IsInteractiveSession)|$($script:SsLines.Count)"
        }
        $o = [ordered]@{}
        $o['gone'] = Invoke-Ss 'gone'
        $o['goneLine'] = @($script:SsLines) -join ' / '
        $o['goneAgain'] = Invoke-Ss 'gone'
        $o['read'] = Invoke-Ss 'read'
        $o
    }
    if ($ss) {
        Assert-Case "a session that cannot be read: taken as Session 0, and said, naming the error" `
            ($ss['gone'] -eq 'False|1' -and
                $ss['goneLine'] -like ('The session this monitor runs in could not be read (Cannot find a process with the process identifier *.), ' +
                    'so the foreground-window and input checks are skipped, as in Session 0; the same error is logged at most once an hour')) `
            "interactive|lines: $($ss['gone']); $($ss['goneLine'])"
        Assert-Case "the same error again in the hour: not said again; read: interactive, nothing said" `
            ($ss['goneAgain'] -eq 'False|0' -and $ss['read'] -eq 'True|0') "again $($ss['goneAgain']), read $($ss['read'])"
    }
}

# Test-ServiceHealth, each read failing in turn. A read that fails is said,
# at most once an hour, and the check passes with its count as it was: no
# repair is started on it. Until 6.3.1 a process list that could not be read
# was no Claude, and a service that could not be read was "not installed":
# either reset the count, and nothing was said.
$svcHealthDef = Get-AstFunction -Ast $watchAst -Name 'Test-ServiceHealth'
$unfaked = @(if ($svcHealthDef) {
        Get-UnfakedCommand -Asts (@($svcHealthDef) + $checkErrorDefs) `
            -Faked (@('Get-ClaudeDesktopProcess', 'Get-ClaudeServiceOrNull', 'Get-Service', 'Write-WatchLog') + $checkErrorNames)
    } else { '<not found>' })
Assert-Case "Test-ServiceHealth was found, and every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $sh = Invoke-Isolated -Name 'the service check' -Block {
        Set-StrictMode -Version Latest
        $ServiceName = 'CoworkVMService'
        $script:ClaudeEnv = [pscustomobject]@{ ProcessName = 'claude' }
        function Write-WatchLog { param([string]$Message) [void]$script:ShLines.Add($Message) }
        # $script:ShProcs: 'one' app process, 'none', or 'throws'.
        # $script:ShSvc: 'Running', 'Stopped', 'missing' or 'throws'.
        # Get-Service is what a tree before 6.3.1 asks, with
        # SilentlyContinue: nothing for a service not there or not read.
        function Get-ClaudeDesktopProcess {
            param($Environment)
            $null = $Environment
            if ($script:ShProcs -eq 'throws') { throw 'Access is denied' }
            if ($script:ShProcs -eq 'none') { return }
            return [pscustomobject]@{ Id = 11 }
        }
        function Get-ClaudeServiceOrNull {
            param([string]$Name)
            if ($script:ShSvc -eq 'throws') { throw "Cannot open $Name service on computer '.'." }
            if ($script:ShSvc -eq 'missing') { return $null }
            return [pscustomobject]@{ Name = $Name; Status = $script:ShSvc }
        }
        function Get-Service {
            param([string]$Name, [object]$ErrorAction)
            $null = $ErrorAction
            if (@('throws', 'missing') -contains $script:ShSvc) { return }
            return [pscustomobject]@{ Name = $Name; Status = $script:ShSvc }
        }
        foreach ($d in $checkErrorDefs) { . ([scriptblock]::Create($d.Extent.Text)) }
        . ([scriptblock]::Create($svcHealthDef.Extent.Text))
        $script:CheckErrorsSaid = @{}
        function Invoke-Sh {
            param([string]$Procs, [string]$Svc, [int]$CountBefore)
            $script:ShProcs = $Procs
            $script:ShSvc = $Svc
            $script:ServiceDownCount = $CountBefore
            $script:ShLines = [System.Collections.ArrayList]::new()
            $r = try { [string](Test-ServiceHealth) } catch { "threw: $($_.Exception.Message)" }
            return "$r|$($script:ServiceDownCount)|$(@($script:ShLines) -join ' / ')"
        }
        $o = [ordered]@{}
        $o['procsUnread']      = Invoke-Sh 'throws' 'Running' 1
        $o['procsUnreadAgain'] = Invoke-Sh 'throws' 'Running' 1
        $o['svcUnread']        = Invoke-Sh 'one' 'throws' 1
        $o['svcUnreadAgain']   = Invoke-Sh 'one' 'throws' 1
        $o['svcMissing']       = Invoke-Sh 'one' 'missing' 1
        $o['noClaude']         = Invoke-Sh 'none' 'Stopped' 1
        $o['stoppedFirst']     = Invoke-Sh 'one' 'Stopped' 0
        $o['stoppedSecond']    = Invoke-Sh 'one' 'Stopped' 1
        $o['running']          = Invoke-Sh 'one' 'Running' 1
        $o
    }
    if ($sh) {
        $tail = '; the same error is logged at most once an hour'
        Assert-Case "Claude's processes that cannot be listed: passes, the count kept, and said, naming the error" `
            ($sh['procsUnread'] -eq ("True|1|Claude's processes could not be listed (Access is denied), so whether CoworkVMService is needed " +
                "is not known and nothing is done about it$tail")) $sh['procsUnread']
        Assert-Case "and not said again in the hour" ($sh['procsUnreadAgain'] -eq 'True|1|') $sh['procsUnreadAgain']
        Assert-Case "a service that cannot be read: passes, the count kept, and said, naming the error" `
            ($sh['svcUnread'] -eq ("True|1|CoworkVMService could not be read (Cannot open CoworkVMService service on computer '.'.), " +
                "so whether it is running is not known and nothing is done about it$tail")) $sh['svcUnread']
        Assert-Case "and not said again in the hour" ($sh['svcUnreadAgain'] -eq 'True|1|') $sh['svcUnreadAgain']
        Assert-Case "a service not installed: passes, the count reset, nothing said" ($sh['svcMissing'] -eq 'True|0|') $sh['svcMissing']
        Assert-Case "no Claude: passes, the count reset" ($sh['noClaude'] -eq 'True|0|') $sh['noClaude']
        Assert-Case "stopped: the first check waits to confirm, the second fails" `
            ($sh['stoppedFirst'] -eq 'True|1|Service not running (check 1/2, waiting to confirm)' -and $sh['stoppedSecond'] -eq 'False|0|') `
            "first $($sh['stoppedFirst']); second $($sh['stoppedSecond'])"
        Assert-Case "running: passes, the count reset" ($sh['running'] -eq 'True|0|') $sh['running']
    }
}

# The start and each poll: whether Claude is running. A list that cannot be
# read at start leaves it unknown, said after the banner with the reads the
# discovery could not make, and the first poll that reads it seeds it, with
# no transition to say; a poll that cannot read it checks nothing and does
# nothing, said at most once an hour. Until 6.3.1 such a list was empty:
# "Claude is not running".
$seedTry = @($topStatements | Where-Object { $_ -is [System.Management.Automation.Language.TryStatementAst] -and
        @($_.Body.Statements).Count -eq 1 -and
        $_.Body.Statements[0].Extent.Text -eq '$script:ClaudeWasRunning = (@(Get-ClaudeDesktopProcess -Environment $script:ClaudeEnv).Count -gt 0)' })
$readErrorsLoop = @($topStatements | Where-Object { $_ -is [System.Management.Automation.Language.ForEachStatementAst] -and
        $_.Condition.Extent.Text -eq '@($script:ClaudeEnv.ReadErrors)' })
$seedIf = @($topStatements | Where-Object { $_ -is [System.Management.Automation.Language.IfStatementAst] -and
        $_.Clauses[0].Item1.Extent.Text -eq '$script:SeedReadError' })
$loopReadTry = @($watchAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.TryStatementAst] -and
            @($n.Body.Statements).Count -eq 1 -and
            $n.Body.Statements[0].Extent.Text -eq '$claudeRunning = @(Get-ClaudeDesktopProcess -Environment $script:ClaudeEnv).Count -gt 0' }, $true))
$transitionIf = @($watchAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.IfStatementAst] -and
            $n.Clauses[0].Item1.Extent.Text -eq '$null -eq $script:ClaudeWasRunning' }, $true))
$startFound = ($seedTry.Count -eq 1 -and $readErrorsLoop.Count -eq 1 -and $seedIf.Count -eq 1 -and $loopReadTry.Count -eq 1 -and
    $transitionIf.Count -eq 1)
$unfaked = @(if ($startFound) {
        Get-UnfakedCommand -Asts (@($seedTry) + @($readErrorsLoop) + @($seedIf) + @($loopReadTry) + @($transitionIf) + $checkErrorDefs) `
            -Faked (@('Get-ClaudeDesktopProcess', 'Get-ClaudeServiceOrNull', 'Get-StealthRelaunchMarker', 'Write-WatchLog') + $checkErrorNames)
    } else { '<not found>' })
Assert-Case "the start's read of Claude's processes, its lines, the poll's read and the transition were found, and every command they call is faked" `
    ($startFound -and $unfaked.Count -eq 0) ("seed $($seedTry.Count), read errors $($readErrorsLoop.Count), seed line $($seedIf.Count), " +
        "poll $($loopReadTry.Count), transition $($transitionIf.Count); $($unfaked -join ', ')")
if ($startFound -and $unfaked.Count -eq 0) {
    $sp = Invoke-Isolated -Name "the start's and the poll's read of Claude's processes" -Block {
        Set-StrictMode -Version Latest
        function Write-WatchLog { param([string]$Message) [void]$script:SpLines.Add($Message) }
        function Get-ClaudeDesktopProcess {
            param($Environment)
            $null = $Environment
            if ($script:SpProcs -eq 'throws') { throw 'Access is denied' }
            if ($script:SpProcs -eq 'none') { return }
            return [pscustomobject]@{ Id = 11 }
        }
        function Get-ClaudeServiceOrNull {
            param([string]$Name)
            if ($script:SpSvc -eq 'throws') { throw "Cannot open $Name service on computer '.'." }
            if ($script:SpSvc -eq 'missing') { return $null }
            return [pscustomobject]@{ Name = $Name; Status = 'Running' }
        }
        function Get-StealthRelaunchMarker { return $null }
        foreach ($d in $checkErrorDefs) { . ([scriptblock]::Create($d.Extent.Text)) }
        $script:CheckErrorsSaid = @{}
        $ServiceName = 'CoworkVMService'
        $NoStealthRelaunch = $false
        $seedRun = [scriptblock]::Create($seedTry[0].Extent.Text + "`r`n" + $readErrorsLoop[0].Extent.Text + "`r`n" + $seedIf[0].Extent.Text)
        # The poll's read in a loop of one pass, as in the main loop: a
        # continue ends the pass before the transition.
        $pollRun = [scriptblock]::Create("foreach (`$pass in 1) {`r`n" + $loopReadTry[0].Extent.Text + "`r`n`$script:SpReached = `$true`r`n" +
            $transitionIf[0].Extent.Text + "`r`n}")
        function Get-SpWas { if ($null -eq $script:ClaudeWasRunning) { return 'unknown' } return [string]$script:ClaudeWasRunning }
        function Invoke-Seed {
            param([string]$Procs, [string[]]$ReadErrors)
            $script:SpProcs = $Procs
            $script:SpLines = [System.Collections.ArrayList]::new()
            $script:ClaudeEnv = [pscustomobject]@{ ProcessName = 'claude'; ReadErrors = @($ReadErrors) }
            # The two lines before the read, as at start.
            $script:ClaudeWasRunning = $null
            $script:SeedReadError = $null
            . $seedRun
            return "$(Get-SpWas)|$(@($script:SpLines) -join ' / ')"
        }
        function Invoke-Poll {
            param([string]$Procs, [string]$Svc = 'Running')
            $script:SpProcs = $Procs
            $script:SpSvc = $Svc
            $script:SpReached = $false
            $script:SpLines = [System.Collections.ArrayList]::new()
            . $pollRun
            return "$($script:SpReached)|$(Get-SpWas)|$(@($script:SpLines) -join ' / ')"
        }
        $o = [ordered]@{}
        $o['seedUnread']         = Invoke-Seed 'throws' @('Service CoworkVMService could not be read: Access is denied')
        $o['pollUnread']         = Invoke-Poll 'throws'
        $o['pollUnreadAgain']    = Invoke-Poll 'throws'
        $o['pollFirst']          = Invoke-Poll 'none'
        $o['pollSame']           = Invoke-Poll 'none'
        $script:ClaudeWasRunning = $true
        $o['pollGoneSvcUnread']  = Invoke-Poll 'none' 'throws'
        $script:ClaudeWasRunning = $true
        $o['pollGoneSvcMissing'] = Invoke-Poll 'none' 'missing'
        $o['pollBack']           = Invoke-Poll 'one'
        $o['seedRead']           = Invoke-Seed 'one' @()
        $o
    }
    if ($sp) {
        $tail = '; the same error is logged at most once an hour'
        Assert-Case "a list that cannot be read at start: unknown, and the discovery's read errors and the seed's said after the banner" `
            ($sp['seedUnread'] -eq ('unknown|READ ERROR at start: Service CoworkVMService could not be read: Access is denied / ' +
                "Whether Claude is running could not be read at start (Access is denied); the first poll that reads it starts from there$tail")) `
            $sp['seedUnread']
        Assert-Case "a poll that cannot read it: goes no further, still unknown, and said" `
            ($sp['pollUnread'] -eq "False|unknown|Whether Claude is running could not be read (Access is denied), so this poll checked nothing and did nothing$tail") `
            $sp['pollUnread']
        Assert-Case "and not said again in the hour" ($sp['pollUnreadAgain'] -eq 'False|unknown|') $sp['pollUnreadAgain']
        Assert-Case "the first poll that reads it: seeds it, with no transition said" ($sp['pollFirst'] -eq 'True|False|') $sp['pollFirst']
        Assert-Case "unchanged on the next: nothing said" ($sp['pollSame'] -eq 'True|False|') $sp['pollSame']
        Assert-Case "Claude gone and the service not read: said as not known, not as absent" `
            ($sp['pollGoneSvcUnread'] -eq ("True|False|Claude is no longer running (service: not known, it could not be read: " +
                "Cannot open CoworkVMService service on computer '.'.). Not acting: a crash and a deliberate quit look the same from here.")) `
            $sp['pollGoneSvcUnread']
        Assert-Case "Claude gone and the service not installed: absent" `
            ($sp['pollGoneSvcMissing'] -eq ('True|False|Claude is no longer running (service: absent). ' +
                'Not acting: a crash and a deliberate quit look the same from here.')) $sp['pollGoneSvcMissing']
        Assert-Case "Claude back: said" ($sp['pollBack'] -eq 'True|True|Claude is running again') $sp['pollBack']
        Assert-Case "a list read at start: seeds it, nothing said" ($sp['seedRead'] -eq 'True|') $sp['seedRead']
    }
}

# The catches of the poll and of the focus check before a repair, each run
# after an error.
$catchAsts = @($pollCatch) + @($focusCatch)
$unfaked = @(Get-UnfakedCommand -Asts ($catchAsts + $checkErrorDefs) -Faked (@('Write-WatchLog') + $checkErrorNames))
$catchReady = ($pollCatch.Count -eq 1 -and $focusCatch.Count -eq 1 -and $unfaked.Count -eq 0)
Assert-Case "the poll's catch and the focus check's were found, and every command they call is faked" $catchReady `
    "poll $($pollCatch.Count), focus $($focusCatch.Count); $($unfaked -join ', ')"
if ($catchReady) {
    $cc = Invoke-Isolated -Name "the poll's and the focus check's catches" -Block {
        Set-StrictMode -Version Latest
        function Write-WatchLog { param([string]$Message) [void]$script:CcLines.Add($Message) }
        foreach ($d in $checkErrorDefs) { . ([scriptblock]::Create($d.Extent.Text)) }
        $script:CheckErrorsSaid = @{}
        $pollRun = [scriptblock]::Create("try { throw 'The poll broke' } catch " + $pollCatch[0].Body.Extent.Text)
        $focusRun = [scriptblock]::Create("try { throw 'The foreground window could not be read' } catch " + $focusCatch[0].Body.Extent.Text)
        function Invoke-Cc {
            param([scriptblock]$Run)
            $script:CcLines = [System.Collections.ArrayList]::new()
            . $Run
            return (@($script:CcLines) -join ' / ')
        }
        $o = [ordered]@{}
        $o['poll'] = Invoke-Cc $pollRun
        $o['pollAgain'] = Invoke-Cc $pollRun
        $o['focus'] = Invoke-Cc $focusRun
        $o['focusAgain'] = Invoke-Cc $focusRun
        $o
    }
    if ($cc) {
        Assert-Case "a poll that fails: said, naming the error, once in the hour" `
            ($cc['poll'] -eq 'MONITOR ERROR: The poll broke; the same error is logged at most once an hour' -and $cc['pollAgain'] -eq '') `
            "first: $($cc['poll']); again: $($cc['pollAgain'])"
        Assert-Case "a focus check that fails before a repair: said, naming the error, once in the hour" `
            ($cc['focus'] -eq ('Whether Claude is in focus could not be read before the repair (The foreground window could not be read); ' +
                    'the other checks still decide; the same error is logged at most once an hour') -and $cc['focusAgain'] -eq '') `
            "first: $($cc['focus']); again: $($cc['focusAgain'])"
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- the monitor: its Fix script and the stealth marker, when they cannot be read"

# The Fix script the monitor starts: beside it, or in a fallback place. One
# that could not be read stops the monitor, saying why; it is not passed over
# for the next place, which would start a different Fix. Until 6.3.1
# Test-Path answered $false for it, so the next copy was started, or the
# monitor stopped saying the script was not found. The search and the fatal
# check are lifted and run in a child PowerShell, since the check exits,
# against fakes of what they read. Every path is under C:\nx, which is never
# created, or a fallback place the fake answers for.
$fsFrom = $watchSrc.IndexOf('$fixScript = Join-Path $myDir "Fix-ClaudeDesktop.ps1"', [System.StringComparison]::Ordinal)
$fsTo = if ($fsFrom -ge 0) { $watchSrc.IndexOf('# The fatal check for a missing Fix script is deliberately NOT here', $fsFrom, [System.StringComparison]::Ordinal) } else { -1 }
$fsFind = if ($fsTo -gt $fsFrom) { $watchSrc.Substring($fsFrom, $fsTo - $fsFrom) } else { $null }
$fsFatal = @($watchAst.EndBlock.Statements | Where-Object {
        $_ -is [System.Management.Automation.Language.IfStatementAst] -and $_.Extent.Text -match 'FATAL: Fix-ClaudeDesktop\.ps1' })
$fsFaked = @('Get-Item', 'Test-Path', 'Write-WatchLog', 'Write-Host', 'Get-ClaudeItemOrNull', 'Get-ClaudeErrorText', 'Test-ClaudeNotFoundError')
$unfaked = @(if ($fsFind -and $fsFatal.Count -eq 1) {
        Get-UnfakedCommand -Asts @([System.Management.Automation.Language.Parser]::ParseInput($fsFind + "`r`n" + $fsFatal[0].Extent.Text, [ref]$null, [ref]$null)) -Faked $fsFaked
    } else { "<found search $($null -ne $fsFind), fatal check $($fsFatal.Count)>" })
Assert-Case "the monitor's search for its Fix script and the fatal check were found, and every command they call is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $fsRegion = (Join-Path $PSScriptRoot 'ClaudeEnv.region.ps1').Replace("'", "''")
    $fsBody = @'
Set-StrictMode -Version Latest
. '<REGION>'
# FsItems maps a path to There or Denied; a path not listed is not there.
# Get-Item writes the errors the cmdlet writes (measured 2026-11-22 and
# 2026-11-23), and Test-Path answers $false for both, as the cmdlet does.
function Get-Item {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string]$Path, [string]$LiteralPath, [switch]$Force)
    $null = $Force
    $p = if ($LiteralPath) { $LiteralPath } else { $Path }
    $v = $FsItems[$p]
    if ($null -eq $v) {
        $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                [System.Management.Automation.ItemNotFoundException]::new("Cannot find path '$p' because it does not exist."),
                'PathNotFound', [System.Management.Automation.ErrorCategory]::ObjectNotFound, $p))
        return
    }
    if ($v -eq 'Denied') {
        $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                [System.UnauthorizedAccessException]::new('Access is denied'),
                'ItemExistsUnauthorizedAccessError', [System.Management.Automation.ErrorCategory]::PermissionDenied, $p))
        return
    }
    return [pscustomobject]@{ FullName = $p }
}
function Test-Path { param([Parameter(Position = 0)][string]$Path) return ($FsItems[$Path] -eq 'There') }
# Straight to stdout, so no line is wrapped to a console width.
function Write-WatchLog { param([string]$Message) [Console]::Out.WriteLine("WATCHLOG: $Message") }
function Write-Host { param([object]$Object, [string]$ForegroundColor) $null = $ForegroundColor; [Console]::Out.WriteLine("HOST: $Object") }
$myDir = 'C:\nx\watch'
$Quiet = $false
$script:WatchLogFile = 'C:\nx\watch.log'
'@
    $fsBody = $fsBody.Replace('<REGION>', $fsRegion)
    $fsMine = 'C:\nx\watch\Fix-ClaudeDesktop.ps1'
    $fsC = 'C:\ClaudeFix\Fix-ClaudeDesktop.ps1'
    $fsDesk = Join-Path $env:USERPROFILE 'Desktop\Fix-ClaudeDesktop.ps1'
    $fs = [ordered]@{}
    foreach ($c in @(@{ N = 'mineDenied'; I = @{ $fsMine = 'Denied'; $fsC = 'There' } },
                     @{ N = 'fallback'; I = @{ $fsC = 'There' } },
                     @{ N = 'fallbackDenied'; I = @{ $fsC = 'Denied'; $fsDesk = 'There' } },
                     @{ N = 'none'; I = @{} })) {
        $items = @($c.I.Keys | ForEach-Object { "'$($_.Replace("'", "''"))' = '$($c.I[$_])'" }) -join '; '
        $child = Invoke-Child -Name "watch-fixscript-$($c.N)" -Body ("`$FsItems = @{ $items }`r`n" + $fsBody + "`r`n" + $fsFind + "`r`n" +
            "[Console]::Out.WriteLine(`"CHOSEN: `$fixScript`")`r`n" + $fsFatal[0].Extent.Text + "`r`n[Console]::Out.WriteLine('RAN ON')`r`n")
        $fs[$c.N] = "exit $($child.Exit) | " + ((@($child.Out -split "`r?`n") | Where-Object { $_ }) -join ' | ')
    }
    Assert-Case "a Fix script beside the monitor that could not be read: the monitor stops, saying why, and no other copy is started" `
        ($fs['mineDenied'] -eq ("exit 1 | CHOSEN: $fsMine | WATCHLOG: FATAL: Fix-ClaudeDesktop.ps1 at '$fsMine' could not be read " +
            "(Access is denied). The health monitor cannot start. | HOST:   [!] Fix-ClaudeDesktop.ps1 at '$fsMine' " +
            'could not be read. Health monitor cannot start. | HOST:       Reason written to: C:\nx\watch.log')) $fs['mineDenied']
    Assert-Case "none beside it: the first fallback that is there is started" ($fs['fallback'] -eq "exit 0 | CHOSEN: $fsC | RAN ON") $fs['fallback']
    Assert-Case "a fallback that could not be read: the monitor stops, saying why, and the next one is not started" `
        ($fs['fallbackDenied'] -match ("^exit 1 \| CHOSEN: " + [regex]::Escape($fsC) + ' \| WATCHLOG: FATAL: Fix-ClaudeDesktop\.ps1 at ' +
            [regex]::Escape("'$fsC' could not be read (Access is denied).")) -and $fs['fallbackDenied'] -notmatch 'RAN ON') `
        $fs['fallbackDenied']
    Assert-Case "none anywhere: the monitor stops, saying it was not found" `
        ($fs['none'] -match "^exit 1 \| CHOSEN: .* \| WATCHLOG: FATAL: Fix-ClaudeDesktop\.ps1 not found\. Looked next to this script at 'C:\\nx\\watch' " -and
         $fs['none'] -notmatch 'RAN ON') $fs['none']
}

# The stealth marker: one that could not be looked for is said, as one that
# could not be read is, once while it lasts; one that is not there is not.
# Until 6.3.1 Test-Path answered $false for it, and it was passed over as
# absent without a word.
$smNames = @('Get-StealthRelaunchMarker', 'Write-StealthNotice')
$smDefs = @($smNames | ForEach-Object { Get-AstFunction -Ast $watchAst -Name $_ } | Where-Object { $_ })
$smFaked = @('Write-WatchLog', 'Test-Path', 'Get-Item', 'Get-Content', 'ConvertFrom-Json', 'Get-ClaudeItemOrNull', 'Get-ClaudeErrorText',
    'Test-ClaudeNotFoundError') + $smNames
$unfaked = @(if ($smDefs.Count -eq 2) { Get-UnfakedCommand -Asts $smDefs -Faked $smFaked } else { '<not found>' })
Assert-Case "Get-StealthRelaunchMarker was found, and every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $sm = Invoke-Isolated -Name 'the stealth marker reads' -Block {
        Set-StrictMode -Version Latest
        function Write-WatchLog { param([string]$Message) [void]$script:SmLines.Add($Message) }
        function Get-Item {
            [CmdletBinding()]
            param([Parameter(Position = 0)][string]$Path, [string]$LiteralPath, [switch]$Force)
            $null = $Force
            $p = if ($LiteralPath) { $LiteralPath } else { $Path }
            if ($script:SmMarker -eq 'denied') {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [System.UnauthorizedAccessException]::new('Access is denied'),
                        'ItemExistsUnauthorizedAccessError', [System.Management.Automation.ErrorCategory]::PermissionDenied, $p))
                return
            }
            $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                    [System.Management.Automation.ItemNotFoundException]::new("Cannot find path '$p' because it does not exist."),
                    'PathNotFound', [System.Management.Automation.ErrorCategory]::ObjectNotFound, $p))
        }
        function Test-Path { param([string]$LiteralPath) $null = $LiteralPath; return $false }
        # No case here has a marker that is there: reading one is a tripwire.
        function Get-Content { param([string]$LiteralPath, [switch]$Raw) $null = $LiteralPath, $Raw; throw 'TRIPWIRE: Get-Content' }
        foreach ($d in $smDefs) { . ([scriptblock]::Create($d.Extent.Text)) }
        $script:LastStealthNotice = $null
        $o = [ordered]@{}
        foreach ($c in 'denied', 'deniedAgain', 'gone', 'deniedAfter') {
            $script:SmLines = [System.Collections.ArrayList]::new()
            $script:SmMarker = if ($c -eq 'gone') { 'gone' } else { 'denied' }
            $ret = Get-StealthRelaunchMarker -Path 'C:\nx\marker.json'
            $o[$c] = "$($null -eq $ret)|$((@($script:SmLines)) -join ' / ')"
        }
        $o
    }
    if ($sm) {
        $smLine = "STEALTH WARNING: marker at C:\nx\marker.json could not be read: Access is denied"
        Assert-Case "a stealth marker that could not be looked for: not acted on, and said" ($sm['denied'] -eq "True|$smLine") $sm['denied']
        Assert-Case "and again while it lasts: not said again" ($sm['deniedAgain'] -eq 'True|') $sm['deniedAgain']
        Assert-Case "a marker that is not there: nothing said" ($sm['gone'] -eq 'True|') $sm['gone']
        Assert-Case "and one that could not be looked for after it: said again" ($sm['deniedAfter'] -eq "True|$smLine") $sm['deniedAfter']
    }
}

# The stealth relaunch after a start whose saved state could not be
# restored. The ts of the last marker decided about is in that state, so a
# marker written before this monitor started is not acted on: it is recorded
# as decided, and said. One written after the start is acted on, as is any
# when the state was restored. Until 6.3.1 the ts read as 0, and Claude was
# started again for a marker the monitor before may already have relaunched
# for. Start-Process is a fake that records.
$srDef = Get-AstFunction -Ast $watchAst -Name 'Invoke-StealthRelaunch'
$srFaked = @('Get-StealthRelaunchMarker', 'Write-WatchLog', 'Save-WatchState', 'Start-Process')
$unfaked = @(if ($srDef) { Get-UnfakedCommand -Asts @($srDef) -Faked $srFaked } else { '<not found>' })
Assert-Case "Invoke-StealthRelaunch was found, and every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $sr = Invoke-Isolated -Name 'the stealth relaunch after a state that could not be restored' -Block {
        Set-StrictMode -Version Latest
        function Get-StealthRelaunchMarker { return $script:SrMarker }
        function Write-WatchLog { param([string]$Message) [void]$script:SrLines.Add($Message) }
        function Save-WatchState { [void]$script:SrLines.Add('SAVED') }
        function Start-Process {
            [CmdletBinding()]
            param([Parameter(Position = 0)][string]$FilePath)
            [void]$script:SrLines.Add("START $FilePath")
        }
        . ([scriptblock]::Create($srDef.Extent.Text))
        $NoStealthRelaunch = $false
        $MaxStealthMarkerAgeMinutes = 10
        $script:ClaudeEnv = [pscustomobject]@{ Aumid = 'Claude_nx!Claude' }
        $script:StealthNoAumidLogged = $false
        $script:LastStealthMarkerTs = [long]0
        # Started a minute ago; each marker is as old as Age seconds.
        $script:StartTime = [datetime]::UtcNow.AddSeconds(-60)
        $written = [datetime]::new(2026, 11, 23, 10, 0, 0, [System.DateTimeKind]::Local)
        $o = [ordered]@{}
        foreach ($c in @(@{ N = 'unreadBefore'; U = $true; Age = 120; Ts = 1001 },
                         @{ N = 'unreadBeforeAgain'; U = $true; Age = 125; Ts = 1001 },
                         @{ N = 'unreadAfter'; U = $true; Age = 30; Ts = 1002 },
                         @{ N = 'readBefore'; U = $false; Age = 120; Ts = 1003 })) {
            $script:SrLines = [System.Collections.ArrayList]::new()
            $script:StateUnread = $c.U
            $script:SrMarker = [pscustomobject]@{ Path = 'C:\nx\marker.json'; TsRaw = [long]$c.Ts; Written = $written
                AgeMinutes = $c.Age / 60.0; AgeSeconds = [double]$c.Age }
            $ret = Invoke-StealthRelaunch
            $o[$c.N] = "$ret|$($script:LastStealthMarkerTs)|$((@($script:SrLines)) -join ' / ')"
        }
        $script:StateUnread = $false
        $o
    }
    if ($sr) {
        Assert-Case "a state that could not be restored, and a marker written before the start: Claude not started, the marker recorded, and said" `
            ($sr['unreadBefore'] -eq ('False|1001|Stealth marker written 2026-11-23 10:00:00, before this monitor started, not acted on: ' +
                'the saved state that records whether it was could not be read. / SAVED')) $sr['unreadBefore']
        Assert-Case "and that marker again: passed over, nothing said" ($sr['unreadBeforeAgain'] -eq 'False|1001|') $sr['unreadBeforeAgain']
        Assert-Case "a marker written after the start: Claude started, as ever" `
            ($sr['unreadAfter'] -match '^True\|1002\|SAVED / Stealth-update relaunch: marker 30s old, starting Claude_nx!Claude / START Claude_nx!Claude / ') `
            $sr['unreadAfter']
        Assert-Case "a state that was restored: a marker from before the start is acted on, as ever" `
            ($sr['readBefore'] -match '^True\|1003\|SAVED / Stealth-update relaunch: marker 120s old, starting Claude_nx!Claude / START Claude_nx!Claude / ') `
            $sr['readBefore']
    }
}

Write-Host "--- Fix: a read that fails is not 'nothing there'"

# Since 6.3.1 every read of a service, a process list, a scheduled task or a
# path in Fix tells "not there" from "could not be read" by the error's
# identity, and a read that could not be made is said and decides nothing
# an empty read would have decided: no kill, stop, restart, delete or
# escalation is taken on it. Until 6.3.1 those reads ran with their errors
# silenced, and a read that failed looked like nothing there.
#
# Each piece below is lifted from Fix and run against fakes of the cmdlets
# it calls. The fakes write the errors the cmdlets write: "not there" with
# the identity measured 2026-11-22 on Windows 11, Windows PowerShell 5.1
# (the region's helpers carry the measurements), and "could not be read"
# with access denied, any other identity. They are advanced functions, so
# -ErrorAction and -ErrorVariable work on them as on the cmdlets, and the
# code from before 6.3.1, which silenced them, meets the same errors. Every
# path is under C:\nx, which is never created: Join-Path needs a real drive,
# and the fakes answer for everything under it.
Add-Type -AssemblyName 'Microsoft.PowerShell.Commands.Management'

$frFakeText = @'
Set-StrictMode -Version Latest
$script:FrLog   = [System.Collections.ArrayList]::new()
$script:FrCalls = [System.Collections.ArrayList]::new()
$script:FrFiles = [System.Collections.ArrayList]::new()
function Reset-Fr {
    $script:FrLog.Clear(); $script:FrCalls.Clear(); $script:FrFiles.Clear()
    $script:FrSvc = @{}; $script:FrProcs = @{}; $script:FrFails = @{}; $script:FrItems = @{}
    $script:FrDirs = @(); $script:FrDeniedDirs = @(); $script:FrRefuse = @()
    $script:FrHcs = ''; $script:FrMpDenied = $false; $script:FrVal = $null
}
Reset-Fr
# What was said, then what was done, in order; the service reads left out.
function Get-FrSaid {
    return ((@($script:FrLog | ForEach-Object { "LOG: $_" }) +
             @($script:FrCalls | Where-Object { $_ -notmatch '^READ ' })) -join ' | ')
}
function New-FrRecord {
    param([string]$TypeName, [string]$Message, [string]$Id, [string]$Category, $Target)
    $ex = Microsoft.PowerShell.Utility\New-Object $TypeName $Message
    return (Microsoft.PowerShell.Utility\New-Object System.Management.Automation.ErrorRecord($ex, $Id,
            [System.Management.Automation.ErrorCategory]$Category, $Target))
}
function New-FrDenied {
    param($Target)
    return (New-FrRecord 'System.UnauthorizedAccessException' 'Access is denied' 'AccessDenied' 'PermissionDenied' $Target)
}
function New-FrPathDenied {
    # The message is the one each cmdlet really gives (measured 2026-11-23):
    # Get-ChildItem's DirUnauthorizedAccessError names the path, and
    # Get-Item's ItemExistsUnauthorizedAccessError says "Access is denied",
    # which does not.
    param([string]$Target, [string]$Id)
    $msg = "Access to the path '$Target' is denied."
    if ($Id -like 'ItemExists*') { $msg = 'Access is denied' }
    return (New-FrRecord 'System.UnauthorizedAccessException' $msg $Id 'PermissionDenied' $Target)
}
function Get-FrParent { param([string]$Path) return $Path.Substring(0, $Path.LastIndexOf('\')) }
function Get-FrLeaf { param([string]$Path) return $Path.Substring($Path.LastIndexOf('\') + 1) }
function Log { param([string]$Message, [string]$Colour, [switch]$Indent) $null = $Colour, $Indent; [void]$script:FrLog.Add($Message) }
function Write-Host {
    param([object]$Object, [string]$ForegroundColor, [switch]$NoNewline)
    $null = $NoNewline
    $colour = if ($ForegroundColor) { $ForegroundColor } else { 'Plain' }
    if ("$Object") { [void]$script:FrLog.Add("[$colour] $Object") }
}
function Start-Sleep { param([int]$Seconds, [int]$Milliseconds) $null = $Seconds, $Milliseconds; [System.Threading.Thread]::Sleep(10) }
# A service's state: the next in its queue, the last one staying. Running,
# Stopped, Denied (could not be read), NoStatus (its status could not be
# read); a service not listed is not there.
function Get-FrSvcState {
    param([string]$Name)
    if (-not $script:FrSvc.ContainsKey($Name)) { return 'Absent' }
    $q = $script:FrSvc[$Name]
    if ($q -is [System.Collections.ArrayList]) {
        if ($q.Count -gt 1) { $s = $q[0]; $q.RemoveAt(0); return [string]$s }
        return [string]$q[0]
    }
    return [string]$q
}
function Get-Service {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string[]]$Name)
    foreach ($n in $Name) {
        [void]$script:FrCalls.Add("READ $n")
        $s = Get-FrSvcState $n
        if ($s -eq 'Absent') {
            $PSCmdlet.WriteError((New-FrRecord 'Microsoft.PowerShell.Commands.ServiceCommandException' `
                "Cannot find any service with service name '$n'." 'NoServiceFoundForGivenName' 'ObjectNotFound' $n))
        } elseif ($s -eq 'Denied') {
            $PSCmdlet.WriteError((New-FrDenied $n))
        } else {
            $o = [pscustomobject]@{ Name = $n; StartType = 'Automatic' }
            if ($s -eq 'NoStatus') {
                $o | Add-Member -MemberType ScriptProperty -Name Status -Value { throw 'Access is denied' }
            } else {
                $o | Add-Member -MemberType NoteProperty -Name Status -Value $s
            }
            $o | Add-Member -MemberType ScriptMethod -Name Stop -Value { [void]$script:FrCalls.Add("STOP $($this.Name)") }
            $o
        }
    }
}
function Stop-Service {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string]$Name, [switch]$Force)
    $null = $Force
    [void]$script:FrCalls.Add("STOP $Name")
    if ($script:FrFails.ContainsKey("STOP $Name")) { $PSCmdlet.WriteError((New-FrDenied $Name)) }
}
function Start-Service {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string]$Name)
    [void]$script:FrCalls.Add("START $Name")
    if ($script:FrFails.ContainsKey("START $Name")) { $PSCmdlet.WriteError((New-FrDenied $Name)) }
}
function Restart-Service {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string]$Name, [switch]$Force)
    $null = $Force
    [void]$script:FrCalls.Add("RESTART $Name")
    if ($script:FrFails.ContainsKey("RESTART $Name")) { $PSCmdlet.WriteError((New-FrDenied $Name)) }
}
# The processes of a name: a list of ids, or Denied for a list that could
# not be read; a name not listed has none running.
function Get-Process {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string[]]$Name, [int[]]$Id)
    $null = $Id
    foreach ($n in $Name) {
        [void]$script:FrCalls.Add("LIST $n")
        $p = $script:FrProcs[$n]
        if ($p -is [string] -and $p -eq 'Denied') {
            $PSCmdlet.WriteError((New-FrRecord 'System.ComponentModel.Win32Exception' 'Access is denied' `
                'CouldNotEnumerateProcesses' 'NotSpecified' $n))
        } elseif ($null -eq $p -or @($p).Count -eq 0) {
            $PSCmdlet.WriteError((New-FrRecord 'Microsoft.PowerShell.Commands.ProcessCommandException' `
                "Cannot find a process with the name `"$n`". Verify the process name and call the cmdlet again." `
                'NoProcessFoundForGivenName' 'ObjectNotFound' $n))
        } else {
            foreach ($i in @($p)) { [pscustomobject]@{ Name = $n; ProcessName = $n; Id = [int]$i } }
        }
    }
}
function Stop-Process {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline = $true)]$InputObject, [string[]]$Name, [int[]]$Id, [switch]$Force)
    process {
        $null = $Force, $Id
        $what = if ($null -ne $InputObject) { "$($InputObject.Name) $($InputObject.Id)" } else { "$Name" }
        if ($script:FrFails.ContainsKey("KILL $what")) { $PSCmdlet.WriteError((New-FrDenied $what)); return }
        [void]$script:FrCalls.Add("KILL $what")
    }
}
function Start-Process {
    [CmdletBinding()]
    param([Parameter(Position = 0)]$FilePath, [Parameter(Position = 1)]$ArgumentList, [switch]$NoNewWindow, [switch]$Wait, [switch]$PassThru)
    $null = $NoNewWindow, $Wait, $PassThru
    [void]$script:FrCalls.Add(("START-PROCESS $FilePath $ArgumentList").Trim())
}
function Get-HcsList {
    if ($script:FrHcs -eq 'THROWS') { throw 'hcsdiag did not answer' }
    return $script:FrHcs
}
function Close-StaleHcsVms { param([string]$Action) $null = $Action; [void]$script:FrCalls.Add('HCS KILL'); return 1 }
# Paths: FrItems maps a path to its item, or to Denied for one that could
# not be read; a path not listed is not there. FrFiles holds the files,
# FrDirs the folders, FrDeniedDirs the folders that cannot be listed and
# FrRefuse the paths a delete is refused for.
function Get-Item {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string]$Path, [string]$LiteralPath, [switch]$Force)
    $null = $Force
    $p = if ($LiteralPath) { $LiteralPath } else { $Path }
    if (-not $script:FrItems.ContainsKey($p)) {
        $PSCmdlet.WriteError((New-FrRecord 'System.Management.Automation.ItemNotFoundException' `
            "Cannot find path '$p' because it does not exist." 'PathNotFound' 'ObjectNotFound' $p))
        return
    }
    $v = $script:FrItems[$p]
    if ($v -is [string] -and $v -eq 'Denied') { $PSCmdlet.WriteError((New-FrPathDenied $p 'ItemExistsUnauthorizedAccessError')); return }
    return $v
}
# A path that is $null or empty does not bind, as with the cmdlet, which
# ends the statement (measured 2026-11-23:
# ParameterArgumentValidationErrorNullNotAllowed).
function Test-Path {
    [CmdletBinding(DefaultParameterSetName = 'Path')]
    param([Parameter(Position = 0, Mandatory, ParameterSetName = 'Path')][string]$Path,
          [Parameter(Mandatory, ParameterSetName = 'LiteralPath')][string]$LiteralPath, [string]$PathType)
    $null = $PathType
    $p = if ($LiteralPath) { $LiteralPath } else { $Path }
    $v = $script:FrItems[$p]
    return ($null -ne $v -and -not ($v -is [string] -and $v -eq 'Denied'))
}
function Get-ChildItem {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string]$Path, [string]$LiteralPath, [string]$Filter, [switch]$Recurse,
          [switch]$File, [switch]$Directory, [switch]$Force)
    $null = $Force, $File
    $p = if ($LiteralPath) { $LiteralPath } else { $Path }
    if ($script:FrDeniedDirs -contains $p) { $PSCmdlet.WriteError((New-FrPathDenied $p 'DirUnauthorizedAccessError')); return }
    if ($Directory) {
        foreach ($d in @($script:FrDirs)) {
            if ((Get-FrParent $d) -eq $p) { [pscustomobject]@{ Name = (Get-FrLeaf $d); FullName = $d } }
        }
        return
    }
    foreach ($f in @($script:FrFiles)) {
        $under = if ($Recurse) { $f.FullName.StartsWith("$p\") } else { (Get-FrParent $f.FullName) -eq $p }
        if ($under -and (-not $Filter -or $f.Name -like $Filter)) { $f }
    }
    if ($Recurse) {
        foreach ($d in @($script:FrDeniedDirs)) {
            if ($d.StartsWith("$p\")) { $PSCmdlet.WriteError((New-FrPathDenied $d 'DirUnauthorizedAccessError')) }
        }
    }
}
function Remove-Item {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline = $true)]$InputObject, [Parameter(Position = 0)][string]$Path, [string]$LiteralPath,
          [switch]$Recurse, [switch]$Force)
    process {
        $null = $Recurse, $Force
        $t = if ($null -ne $InputObject) { [string]$InputObject.FullName } elseif ($LiteralPath) { $LiteralPath } else { $Path }
        if ($script:FrRefuse -contains $t) {
            $PSCmdlet.WriteError((New-FrRecord 'System.IO.IOException' `
                "Cannot remove item ${t}: The process cannot access the file because it is being used by another process." `
                'RemoveFileSystemItemIOError' 'WriteError' $t))
            return
        }
        [void]$script:FrCalls.Add("REMOVE $t")
        foreach ($f in @($script:FrFiles)) {
            if ($f.FullName -eq $t -or $f.FullName.StartsWith("$t\")) { [void]$script:FrFiles.Remove($f) }
        }
    }
}
function Copy-Item {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string]$Path, [Parameter(Position = 1)][string]$Destination, [switch]$Force)
    $null = $Force
    if ($script:FrRefuse -contains $Path) {
        $PSCmdlet.WriteError((New-FrRecord 'System.IO.IOException' `
            "The process cannot access the file '$Path' because it is being used by another process." 'CopyFileInfoItemIOError' 'WriteError' $Path))
        return
    }
    [void]$script:FrCalls.Add("COPY $Path -> $Destination")
}
function New-FrFile {
    param([string]$Path, [int]$Days)
    [void]$script:FrFiles.Add([pscustomobject]@{ Name = (Get-FrLeaf $Path); FullName = $Path; Length = 1MB
                                                LastWriteTimeUtc = [datetime]::UtcNow.AddDays(-$Days) })
}
# The main body's $PSCmdlet: every action is allowed.
$PSCmdlet = [pscustomobject]@{ Fake = $true } |
    Add-Member -MemberType ScriptMethod -Name ShouldProcess -Value { param($a, $b) $null = $a, $b; return $true } -PassThru
$ServiceName = 'CoworkVMService'
$ServiceExe  = 'cowork-svc'
$ProcessName = 'claude'
'@

# Fix's helpers, lifted where they exist; each block defines them after the
# fakes. Until 6.3.1 none of them existed.
$frHelperNames = @('Get-FixFirstError', 'Get-FixServiceState', 'Stop-FixServiceProcess', 'Start-FixService',
    'Remove-FixCacheFolder', 'Test-FixPathPresent', 'Write-SvcLogUnread')
$frDefs = @($frHelperNames | ForEach-Object { Get-AstFunction -Ast $fixAst -Name $_ } | Where-Object { $_ })
$frSetup = [scriptblock]::Create($frFakeText + "`r`n" + 'foreach ($d in $frDefs) { . ([scriptblock]::Create($d.Extent.Text)) }')
$frFaked = @('Log', 'Write-Host', 'Start-Sleep', 'Get-Service', 'Stop-Service', 'Start-Service', 'Restart-Service',
    'Get-Process', 'Stop-Process', 'Start-Process', 'Get-HcsList', 'Close-StaleHcsVms', 'Get-Item', 'Test-Path',
    'Get-ChildItem', 'Remove-Item', 'Copy-Item', 'Get-ClaudeServiceOrNull', 'Get-ClaudeProcessByName', 'Get-ClaudeScheduledTaskOrNull',
    'Get-ClaudeItemOrNull', 'Get-ClaudeErrorText', 'Test-ClaudeNotFoundError') + $frHelperNames

# The text of Fix from the first $From to the next $To after it: up to its
# start, its end (-ToEnd) or the end of its line (-ToLineEnd).
function Get-FixSegment {
    param([string]$From, [string]$To, [switch]$ToEnd, [switch]$ToLineEnd)
    $i = $fixSrc.IndexOf($From, [System.StringComparison]::Ordinal)
    if ($i -lt 0) { return $null }
    $j = $fixSrc.IndexOf($To, $i + $From.Length, [System.StringComparison]::Ordinal)
    if ($j -lt 0) { return $null }
    if ($ToEnd) { $j += $To.Length }
    if ($ToLineEnd) { $j = $fixSrc.IndexOf("`n", $j, [System.StringComparison]::Ordinal); if ($j -lt 0) { return $null } }
    return $fixSrc.Substring($i, $j - $i)
}

# Every command a lifted piece and the helpers call that is not faked.
# Nothing runs unless this is empty: the shell may be elevated. A piece not
# found, with no definitions, is '<not found>'. $Text has no type: a
# [string] turns the $null of a segment not found into '', and until 6.3.1
# the check for $null here never fired, so "was found" passed for a piece
# that was not.
function Get-FrUnsafe {
    param($Text = $null, [object[]]$Defs = @(), [string[]]$Also = @())
    $asts = @($frDefs) + @($Defs | Where-Object { $_ })
    if ($Text) { $asts += [System.Management.Automation.Language.Parser]::ParseInput($Text, [ref]$null, [ref]$null) }
    elseif (@($Defs | Where-Object { $_ }).Count -eq 0) { return @('<not found>') }
    return @(Get-UnfakedCommand -Asts $asts -Faked ($frFaked + $Also))
}

# --- the helpers -----------------------------------------------------
$fssDef = Get-AstFunction -Ast $fixAst -Name 'Get-FixServiceState'
$unfaked = @(if ($fssDef) { Get-FrUnsafe -Defs @($fssDef) } else { '<not found>' })
Assert-Case "Get-FixServiceState was found, and every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $fss = Invoke-Isolated -Name 'Get-FixServiceState' -Block {
        . $frSetup
        $o = [ordered]@{}
        foreach ($c in @(@{ N = 'up'; S = 'Running' }, @{ N = 'down'; S = 'Stopped' }, @{ N = 'gone'; S = $null },
                         @{ N = 'denied'; S = 'Denied' }, @{ N = 'nostatus'; S = 'NoStatus' })) {
            Reset-Fr
            if ($c.S) { $script:FrSvc['Svc1'] = $c.S }
            $s = Get-FixServiceState -Name 'Svc1'
            $o[$c.N] = "$($s.Status)|$($s.Error)|$($null -ne $s.Service)"
        }
        $o
    }
    if ($fss) {
        Assert-Case "a service that runs: Running" ($fss['up'] -eq 'Running||True') $fss['up']
        Assert-Case "a stopped service: Stopped" ($fss['down'] -eq 'Stopped||True') $fss['down']
        Assert-Case "no service of that name: Absent" ($fss['gone'] -eq 'Absent||False') $fss['gone']
        Assert-Case "a service that could not be read: Unread, with why" ($fss['denied'] -eq 'Unread|Access is denied|False') $fss['denied']
        Assert-Case "a status that could not be read: Unread" ($fss['nostatus'] -eq 'Unread|its status could not be read|True') $fss['nostatus']
    }
}

$sfpDef = Get-AstFunction -Ast $fixAst -Name 'Stop-FixServiceProcess'
$sfsDef = Get-AstFunction -Ast $fixAst -Name 'Start-FixService'
$unfaked = @(if ($sfpDef -and $sfsDef) { Get-FrUnsafe -Defs @($sfpDef, $sfsDef) } else { '<not found>' })
Assert-Case "Stop-FixServiceProcess and Start-FixService were found, and every command they call is faked" `
    ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $sfp = Invoke-Isolated -Name 'Stop-FixServiceProcess' -Block {
        . $frSetup
        $o = [ordered]@{}
        foreach ($c in 'none', 'two', 'denied', 'refused') {
            Reset-Fr
            switch ($c) {
                'two'     { $script:FrProcs['cowork-svc'] = @(301, 302) }
                'denied'  { $script:FrProcs['cowork-svc'] = 'Denied' }
                'refused' { $script:FrProcs['cowork-svc'] = @(301); $script:FrFails['KILL cowork-svc 301'] = $true }
            }
            $n = Stop-FixServiceProcess
            $o[$c] = "returned $(if ($null -eq $n) { 'null' } else { $n }) | $(Get-FrSaid)"
        }
        foreach ($c in 'started', 'refused') {
            Reset-Fr
            if ($c -eq 'refused') { $script:FrFails['START CoworkVMService'] = $true }
            $ok = Start-FixService
            $o["start-$c"] = "returned $ok | $(Get-FrSaid)"
        }
        # Under -WhatIf no kill and no start goes out, and neither is
        # returned as done.
        Reset-Fr
        $script:FrProcs['cowork-svc'] = @(301)
        $n = Stop-FixServiceProcess -WhatIf
        $o['whatif-stop'] = "returned $(if ($null -eq $n) { 'null' } else { $n }) | $(Get-FrSaid)"
        Reset-Fr
        $ok = Start-FixService -WhatIf
        $o['whatif-start'] = "returned $ok | $(Get-FrSaid)"
        $o
    }
    if ($sfp) {
        Assert-Case "no service process running: 0, nothing killed" ($sfp['none'] -eq 'returned 0 | LIST cowork-svc') $sfp['none']
        Assert-Case "two running: both killed, 2" `
            ($sfp['two'] -eq 'returned 2 | LIST cowork-svc | KILL cowork-svc 301 | KILL cowork-svc 302') $sfp['two']
        Assert-Case "a process list that could not be read: null, nothing killed, and said" `
            ($sfp['denied'] -eq "returned null | LOG: [!] cowork-svc's processes could not be listed (Access is denied), so none were force-killed | LIST cowork-svc") `
            $sfp['denied']
        Assert-Case "a kill that was refused: null, and said" `
            ($sfp['refused'] -eq 'returned null | LOG: [!] cowork-svc could not be force-killed (Access is denied) | LIST cowork-svc') $sfp['refused']
        Assert-Case "a start that works: true" ($sfp['start-started'] -eq 'returned True | START CoworkVMService') $sfp['start-started']
        Assert-Case "a start that fails: false, and said" `
            ($sfp['start-refused'] -eq 'returned False | LOG: [!] CoworkVMService could not be started (Access is denied) | START CoworkVMService') `
            $sfp['start-refused']
        Assert-Case "under -WhatIf: nothing killed or started, and neither returned as done" `
            ($sfp['whatif-stop'] -eq 'returned null | LIST cowork-svc' -and $sfp['whatif-start'] -eq 'returned False | ') `
            "[$($sfp['whatif-stop'])] [$($sfp['whatif-start'])]"
    }
}

# The three purges, Step 6 and the two escalations to Deep, call one
# helper. Until 6.3.1 each said the folder was removed whether or not the
# delete worked, and a folder that could not be read was "not present".
$rcfDef = Get-AstFunction -Ast $fixAst -Name 'Remove-FixCacheFolder'
$rcfCalls = @(Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Remove-FixCacheFolder' })
Assert-Case "the three purges go through Remove-FixCacheFolder" ($null -ne $rcfDef -and $rcfCalls.Count -eq 3) "found $($rcfCalls.Count) call(s)"
$unfaked = @(if ($rcfDef) { Get-FrUnsafe -Defs @($rcfDef) } else { '<not found>' })
Assert-Case "every command Remove-FixCacheFolder calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $rcf = Invoke-Isolated -Name 'Remove-FixCacheFolder' -Block {
        . $frSetup
        $o = [ordered]@{}
        foreach ($c in 'denied', 'absent', 'absentSaid', 'refused', 'removed', 'unknown', 'unknownSaid') {
            Reset-Fr
            if ($c -eq 'denied') { $script:FrItems['C:\nx\cache'] = 'Denied' }
            if ($c -in @('refused', 'removed')) {
                $script:FrItems['C:\nx\cache'] = [pscustomobject]@{ FullName = 'C:\nx\cache' }
                New-FrFile 'C:\nx\cache\a.vhdx' 1
                New-FrFile 'C:\nx\cache\b.vhdx' 1
            }
            if ($c -eq 'refused') { $script:FrRefuse = @('C:\nx\cache') }
            try {
                if ($c -eq 'unknown') { Remove-FixCacheFolder -Path $null -Label 'VM bundle' }
                elseif ($c -eq 'unknownSaid') { Remove-FixCacheFolder -Path $null -Label 'VM bundle' -SayAbsent }
                elseif ($c -eq 'absentSaid') { Remove-FixCacheFolder -Path 'C:\nx\cache' -Label 'VM bundle' -SayAbsent }
                else { Remove-FixCacheFolder -Path 'C:\nx\cache' -Label 'VM bundle' }
                $o[$c] = Get-FrSaid
            } catch { $o[$c] = "THREW $($_.FullyQualifiedErrorId)" }
        }
        $o
    }
    if ($rcf) {
        Assert-Case "a cache folder that could not be read: not purged, and said" `
            ($rcf['denied'] -eq "LOG: [!] VM bundle could not be read (Access is denied), so it was not purged") $rcf['denied']
        Assert-Case "a cache folder that is not there: nothing, or 'not present' when asked" `
            ($rcf['absent'] -eq '' -and $rcf['absentSaid'] -eq 'LOG: VM bundle not present') "[$($rcf['absent'])] [$($rcf['absentSaid'])]"
        Assert-Case "a delete that failed: not called removed, and said" `
            ($rcf['refused'] -match '^LOG: \[!\] VM bundle could not be removed in full \(Cannot remove item C:\\nx\\cache: ' -and
             $rcf['refused'] -notmatch 'removed \(') $rcf['refused']
        Assert-Case "a delete that worked: removed, with its size" ($rcf['removed'] -eq 'LOG: VM bundle removed (2 MB freed) | REMOVE C:\nx\cache') $rcf['removed']
        # No Claude data folder was found. Until 6.3.1 the empty path did
        # not bind, which ended the run with UNEXPECTED ERROR.
        Assert-Case "a cache folder that is not known: nothing, or 'not known' when asked, and no error" `
            ($rcf['unknown'] -eq '' -and $rcf['unknownSaid'] -eq 'LOG: VM bundle not known (no Claude data folder was found)') `
            "[$($rcf['unknown'])] [$($rcf['unknownSaid'])]"
    }
}

# Test-FixPathPresent: a file at the path is there; nothing, or a folder, is
# not; a path that could not be read is $null, and said. Since 6.3.1 a folder
# is not there: until then it answered $true, as Test-Path does, and a folder
# at a Claude.exe path was started, which opens it in Explorer. The items are
# made in memory only: nothing at C:\nx is created.
$fppDef = Get-AstFunction -Ast $fixAst -Name 'Test-FixPathPresent'
$unfaked = @(if ($fppDef) { Get-FrUnsafe -Defs @($fppDef) } else { '<not found>' })
Assert-Case "every command Test-FixPathPresent calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $fpp = Invoke-Isolated -Name 'Test-FixPathPresent' -Block {
        . $frSetup
        Reset-Fr
        $script:FrItems['C:\nx\app\Claude.exe']    = [System.IO.FileInfo]::new('C:\nx\app\Claude.exe')
        $script:FrItems['C:\nx\dir\Claude.exe']    = [System.IO.DirectoryInfo]::new('C:\nx\dir\Claude.exe')
        $script:FrItems['C:\nx\denied\Claude.exe'] = 'Denied'
        $o = [ordered]@{}
        foreach ($c in @(@{ N = 'file'; P = 'C:\nx\app\Claude.exe' }, @{ N = 'folder'; P = 'C:\nx\dir\Claude.exe' },
                         @{ N = 'none'; P = 'C:\nx\none\Claude.exe' }, @{ N = 'denied'; P = 'C:\nx\denied\Claude.exe' })) {
            $script:FrLog.Clear()
            $r = Test-FixPathPresent -Path $c.P -What 'The Claude.exe'
            $o[$c.N] = "$(if ($null -eq $r) { 'null' } else { $r })|$(Get-FrSaid)"
        }
        $o
    }
    if ($fpp) {
        Assert-Case "a launch candidate that is a file: there" ($fpp['file'] -eq 'True|') $fpp['file']
        Assert-Case "a launch candidate that is a folder: not there, so it is not started" ($fpp['folder'] -eq 'False|') $fpp['folder']
        Assert-Case "a launch candidate that is not there: not there, nothing said" ($fpp['none'] -eq 'False|') $fpp['none']
        Assert-Case "a launch candidate that could not be read: not known, and said" `
            ($fpp['denied'] -eq ("null|LOG: [!] The Claude.exe could not be read (Access is denied), " +
                'so it was passed over: C:\nx\denied\Claude.exe')) $fpp['denied']
    }
}

# --- Step 6 with the cache kept, and the Diagnostic run's VM cache lines ----
# A folder that is not known (no Claude data folder was found) is passed
# over, and one that could not be read is said, with why. Until 6.3.1
# Test-Path answered $false for the second, "Not present" in the Diagnostic
# run, and threw for the first, which ended the run with UNEXPECTED ERROR.
$keepLoop = @(Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.ForEachStatementAst] -and $n.Body.Extent.Text -match 'preserved \(\$sizeMB MB\)' })
$diagLoop = @(Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.ForEachStatementAst] -and $n.Body.Extent.Text -match ' : Not present"' })
$unfaked = @(if ($keepLoop.Count -eq 1 -and $diagLoop.Count -eq 1) { Get-FrUnsafe -Text ($keepLoop[0].Extent.Text + "`r`n" + $diagLoop[0].Extent.Text) }
    else { "<found $($keepLoop.Count) and $($diagLoop.Count)>" })
Assert-Case "Step 6's kept-cache lines and the Diagnostic run's were found, and every command they call is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $vc = Invoke-Isolated -Name 'the VM cache lines' -Block {
        . $frSetup
        $o = [ordered]@{}
        foreach ($k in @(@{ N = 'keep'; S = $keepLoop[0] }, @{ N = 'diag'; S = $diagLoop[0] })) {
            $seg = [scriptblock]::Create($k.S.Extent.Text)
            foreach ($c in 'unknownDenied', 'thereGone') {
                Reset-Fr
                if ($c -eq 'unknownDenied') {
                    $VmCachePath = $null
                    $BundlePath = 'C:\nx\bundle'
                    $script:FrItems['C:\nx\bundle'] = 'Denied'
                } else {
                    $VmCachePath = 'C:\nx\vm'
                    $BundlePath = 'C:\nx\bundle'
                    $script:FrItems['C:\nx\vm'] = [pscustomobject]@{ FullName = 'C:\nx\vm' }
                    New-FrFile 'C:\nx\vm\a.vhdx' 1
                }
                try { & { . $seg }; $o["$($k.N)_$c"] = Get-FrSaid }
                catch { $o["$($k.N)_$c"] = "$(Get-FrSaid) | THREW $($_.FullyQualifiedErrorId)" }
            }
        }
        $o
    }
    if ($vc) {
        Assert-Case "Step 6, kept: a folder that is not known is passed over, and one that could not be read is said" `
            ($vc['keep_unknownDenied'] -eq "LOG: [!] vm_bundles could not be read (Access is denied); it was kept") $vc['keep_unknownDenied']
        Assert-Case "Step 6, kept: a folder that is there is said with its size, one that is not there is not" `
            ($vc['keep_thereGone'] -eq 'LOG: claude-code-vm preserved (1 MB)') $vc['keep_thereGone']
        Assert-Case "Diagnostic: a folder that is not known is said so, and one that could not be read is said, with why" `
            ($vc['diag_unknownDenied'] -eq ('LOG:   claude-code-vm : Not known (no Claude data folder was found) | ' +
                "LOG:   vm_bundles : could not be read (Access is denied)")) $vc['diag_unknownDenied']
        Assert-Case "Diagnostic: a folder that is there is said with its size, one that is not there is 'Not present'" `
            ($vc['diag_thereGone'] -eq 'LOG:   claude-code-vm : 1 MB | LOG:   vm_bundles : Not present') $vc['diag_thereGone']
    }
}

# --- the support bundle ------------------------------------------------
# A log that could not be read or copied is said, with why, apart from the
# ones that are not there, and a VM log that could not be read gives no note
# that the VM may never have started. Until 6.3.1 Test-Path answered $false
# for a log it could not read, and both were "Skipped" with the absent ones.
$bundleText = Get-FixSegment -From '        $bundleBudget = 4MB' -To '        # The Service Control Manager entries'
$unfaked = @(Get-FrUnsafe -Text $bundleText -Also @('New-Object'))
Assert-Case "the support bundle's log copy was found, and every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $bun = Invoke-Isolated -Name 'the support bundle log copy' -Block {
        . $frSetup
        $o = [ordered]@{}
        $seg = [scriptblock]::Create($bundleText)
        $bundleDir = 'C:\nx\bundle'
        $script:ClaudeEnv = [pscustomobject]@{ LogDir = 'C:\nx\logs' }
        foreach ($c in 'mixed', 'noVmLog') {
            Reset-Fr
            if ($c -eq 'mixed') {
                $wanted = @('main.log', 'cowork_vm_node.log', 'gone.log', 'big.log')
                $script:FrItems['C:\nx\logs\main.log'] = [pscustomobject]@{ Length = 1KB }
                $script:FrItems['C:\nx\logs\cowork_vm_node.log'] = 'Denied'
                $script:FrItems['C:\nx\logs\big.log'] = [pscustomobject]@{ Length = 1KB }
                $script:FrRefuse = @('C:\nx\logs\big.log')
            } else {
                $wanted = @('main.log')
                $script:FrItems['C:\nx\logs\main.log'] = [pscustomobject]@{ Length = 1KB }
            }
            try { & { . $seg }; $o[$c] = Get-FrSaid }
            catch { $o[$c] = "$(Get-FrSaid) | THREW $($_.Exception.Message)" }
        }
        $o
    }
    if ($bun) {
        $mixed = " | $($bun['mixed']) | "
        Assert-Case "the bundle: what was copied and what is not there are said as before" `
            ($mixed -match ' \| LOG:   Copied  : 1 of 4 log files \| ' -and $mixed -match ' \| LOG:   Skipped : gone\.log \| ' -and
             $mixed -match ' \| COPY C:\\nx\\logs\\main\.log -> C:\\nx\\bundle\\main\.log \| ') $bun['mixed']
        Assert-Case "the bundle: a log that could not be read is said, with why, and not skipped as absent" `
            ($mixed -match (" \| LOG:   \[!\] cowork_vm_node\.log could not be read \(Access is denied\), " +
                'so it is not in the bundle \| ')) $bun['mixed']
        Assert-Case "the bundle: a log that could not be copied is said, with why" `
            ($mixed -match (" \| LOG:   \[!\] big\.log could not be copied \(The process cannot access the file 'C:\\nx\\logs\\big\.log' " +
                'because it is being used by another process\.\), so it is not in the bundle \| ')) $bun['mixed']
        Assert-Case "the bundle: a VM log that could not be read gives no note that the VM never started" `
            ($mixed -notmatch 'no VM log in the folder') $bun['mixed']
        Assert-Case "the bundle: a VM log that is not there still gives the note" `
            ($bun['noVmLog'] -match '\| LOG:   Note    : no VM log in the folder, so the Cowork VM may never have started here( \||$)') $bun['noVmLog']
    }
}

# --- the start-up clean-ups ------------------------------------------
# What they could not do is noted and said once Log exists. Until 6.3.1 it
# was passed over without a word. Nothing is decided on it.
$startText = Get-FixSegment -From '# Clean up logs older than 30 days.' -To '$script:LogLines = New-Object'
$saveIf = @(Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -match '^-not \(Save-ClaudeEnvironment ' })
$notesLoop = @(Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.ForEachStatementAst] -and $n.Condition.Extent.Text -eq '@($script:StartupNotes)' })
$unfaked = @(Get-FrUnsafe -Text $startText)
if ($saveIf.Count -ne 1 -or $notesLoop.Count -ne 1) { $unfaked += "<found $($saveIf.Count) save check(s), $($notesLoop.Count) notes loop(s)>" }
else { $unfaked += @(Get-UnfakedCommand -Asts @($saveIf[0], $notesLoop[0]) -Faked ($frFaked + 'Save-ClaudeEnvironment')) }
Assert-Case "the start-up clean-ups and their notes were found, and every command they call is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $su = Invoke-Isolated -Name 'the start-up clean-ups' -Block {
        . $frSetup
        function Save-ClaudeEnvironment { param($Environment, $Path) $null = $Environment, $Path; return $false }
        $seg = [scriptblock]::Create($startText)
        $LogDir = 'C:\nx\logs'
        $EnvArtifact = 'C:\nx\logs\claude-env.json'
        $script:ClaudeEnv = [pscustomobject]@{ LogDir = $LogDir }
        $script:StartupNotes = @()
        New-FrFile 'C:\nx\logs\fix_old1.log' 40
        New-FrFile 'C:\nx\logs\fix_old2.log' 40
        New-FrFile 'C:\nx\logs\fix_new.log' 1
        $script:FrRefuse = @('C:\nx\logs\fix_old2.log')
        $script:FrItems['C:\nx\logs\vhdx-backup'] = 'Denied'
        . ([scriptblock]::Create($saveIf[0].Extent.Text))
        & { . $seg }
        $notes = @($script:StartupNotes)
        $done = Get-FrSaid
        . ([scriptblock]::Create($notesLoop[0].Extent.Text))
        [ordered]@{ Notes = ($notes -join ' | '); Done = $done; Said = ((@($script:FrLog) | Where-Object { $_ -like '`[!`]*' }) -join ' | ') }
    }
    if ($su) {
        Assert-Case "an old log that could not be deleted is noted, and the others are deleted" `
            ($su.Notes -match "Old logs in C:\\nx\\logs could not all be cleaned up \(Cannot remove item C:\\nx\\logs\\fix_old2\.log: " -and
             $su.Done -match 'REMOVE C:\\nx\\logs\\fix_old1\.log' -and $su.Done -notmatch 'fix_new') "notes: $($su.Notes); done: $($su.Done)"
        Assert-Case "a backup folder that could not be read is noted, and nothing in it is touched" `
            ($su.Notes -match "Old VHDX backups in C:\\nx\\logs could not be cleaned up \(Access is denied\)" -and
             $su.Done -notmatch 'vhdx-backup') "notes: $($su.Notes); done: $($su.Done)"
        Assert-Case "a discovery record that could not be saved is noted" `
            ($su.Notes -match '^The discovery record could not be saved to C:\\nx\\logs\\claude-env\.json') $su.Notes
        Assert-Case "and each note is said once Log exists" `
            (@($su.Said -split ' \| ').Count -eq 3 -and $su.Said -match '^\[!\] The discovery record could not be saved') $su.Said
    }
}

# --- Step 0 ------------------------------------------------------------
# A cowork-vm compute system is killed only when Claude is known not to
# run. Until 6.3.1 a process list that could not be read read as Claude not
# running, and the compute system was killed under it.
$step0Text = Get-FixSegment -From 'Log "[0/10] Checking for stale HCS compute systems..."' -To 'Start-Sleep -Seconds 1' -ToEnd
$unfaked = @(Get-FrUnsafe -Text $step0Text)
Assert-Case "Step 0 was found, and every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $s0 = Invoke-Isolated -Name 'Step 0' -Block {
        . $frSetup
        $seg = [scriptblock]::Create($step0Text)
        $script:IsAdmin = $true
        $o = [ordered]@{}
        foreach ($c in 'denied', 'none', 'running') {
            Reset-Fr
            $script:FrHcs = "cowork-vm-1699151a`r`n    VM, Running, DE1517EC-0000-4000-8000-77A48CB1AD97, cowork-vm-1699151a"
            if ($c -eq 'denied') { $script:FrProcs['claude'] = 'Denied' }
            if ($c -eq 'running') { $script:FrProcs['claude'] = @(11) }
            & { . $seg }
            $o[$c] = Get-FrSaid
        }
        $o
    }
    if ($s0) {
        Assert-Case "Step 0, whether Claude runs could not be read: the compute system is left, and that is said" `
            ($s0['denied'] -notmatch 'HCS KILL' -and
             $s0['denied'] -match 'LOG: cowork-vm present, and whether Claude is running could not be read \(Access is denied\); leaving it for Step 5') $s0['denied']
        Assert-Case "Step 0, Claude not running: the compute system is killed" ($s0['none'] -match 'HCS KILL') $s0['none']
        Assert-Case "Step 0, Claude running: the compute system is left" `
            ($s0['running'] -notmatch 'HCS KILL' -and $s0['running'] -match 'Claude is running; leaving it for Step 5') $s0['running']
    }
}

# --- -PurgeSessions ----------------------------------------------------
# A session folder whose files could not all be listed is kept: until
# 6.3.1 it looked empty and was deleted with everything in it. A file that
# could not be deleted is not counted as deleted.
$purgeIf = @(Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -eq '$PurgeSessions' -and
    $n.Extent.Text -match 'SessionsDir' })
$unfaked = @(if ($purgeIf.Count -eq 1) { Get-FrUnsafe -Text $purgeIf[0].Extent.Text } else { "<found $($purgeIf.Count)>" })
Assert-Case "the session purge was found, and every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $ps = Invoke-Isolated -Name 'the session purge' -Block {
        . $frSetup
        $seg = [scriptblock]::Create($purgeIf[0].Extent.Text)
        $PurgeSessions = $true
        $script:ClaudeEnv = [pscustomobject]@{ SessionsDir = 'C:\nx\sess' }
        $script:FrItems['C:\nx\sess'] = [pscustomobject]@{ FullName = 'C:\nx\sess' }
        $script:FrDirs = @('C:\nx\sess\A', 'C:\nx\sess\B', 'C:\nx\sess\C')
        $script:FrDeniedDirs = @('C:\nx\sess\B')
        New-FrFile 'C:\nx\sess\A\f1' 30
        New-FrFile 'C:\nx\sess\C\f3' 30
        $script:FrRefuse = @('C:\nx\sess\C\f3')
        & { . $seg }
        Get-FrSaid
    }
    if ($ps) {
        Assert-Case "a session folder that could not be read is kept" ($ps -notmatch 'REMOVE C:\\nx\\sess\\B') $ps
        Assert-Case "an emptied one is removed, and one with a file left is kept" `
            ($ps -match 'REMOVE C:\\nx\\sess\\A( |$)' -and $ps -notmatch 'REMOVE C:\\nx\\sess\\C( |$)') $ps
        Assert-Case "a file that could not be deleted is not counted, and is said" `
            ($ps -match 'LOG: Deleted 1 of 2 session files older than 7 days' -and $ps -match 'LOG: \[!\] 1 could not be deleted \(Cannot remove item C:\\nx\\sess\\C\\f3: ') $ps
        Assert-Case "the part that could not be read is said, and the folders kept for it" `
            ($ps -match "LOG: \[!\] Part of C:\\nx\\sess could not be read \(Access to the path 'C:\\nx\\sess\\B' is denied\.\); files there were not considered" -and
             $ps -match 'LOG: \[!\] 1 session folder\(s\) could not be read in full, so they were kept') $ps
    }
    # The sessions folder itself could not be read. Until 6.3.1 Test-Path
    # answered $false for it, and the purge asked for was skipped without a
    # word.
    $psDenied = Invoke-Isolated -Name 'the session purge, a folder that could not be read' -Block {
        . $frSetup
        $seg = [scriptblock]::Create($purgeIf[0].Extent.Text)
        $PurgeSessions = $true
        $script:ClaudeEnv = [pscustomobject]@{ SessionsDir = 'C:\nx\sess' }
        $script:FrItems['C:\nx\sess'] = 'Denied'
        $script:FrDirs = @('C:\nx\sess\A')
        New-FrFile 'C:\nx\sess\A\f1' 30
        & { . $seg }
        "said: $(Get-FrSaid)"
    }
    Assert-Case "a sessions folder that could not be read: nothing purged, and said" `
        ($psDenied -eq "said: LOG: [!] The sessions folder C:\nx\sess could not be read (Access is denied), so no session files were purged") `
        "$psDenied"
}

# --- Step 1 and -Close: the kill counts --------------------------------
# Counted from what Stop-Process did not refuse. Until 6.3.1 every process
# listed was counted as killed.
$killIfs = [ordered]@{}
foreach ($k in @(@{ N = 'Step 1'; P = '$claudeProcs | Stop-Process' }, @{ N = '-Close'; P = '$procs | Stop-Process' })) {
    $pipe = @(Find-FixStatement { param($n)
        if ($n -isnot [System.Management.Automation.Language.PipelineAst] -or -not $n.Extent.Text.StartsWith($k.P)) { return $false }
        for ($up = $n.Parent; $null -ne $up; $up = $up.Parent) {
            if ($up -is [System.Management.Automation.Language.FunctionDefinitionAst]) { return $false }
        }
        return $true })
    $node = if ($pipe.Count -eq 1) { $pipe[0].Parent } else { $null }
    while ($null -ne $node -and $node -isnot [System.Management.Automation.Language.IfStatementAst]) { $node = $node.Parent }
    $killIfs[$k.N] = $node
}
$unfaked = @(foreach ($k in $killIfs.Keys) { if ($killIfs[$k]) { Get-FrUnsafe -Text $killIfs[$k].Extent.Text } else { "<$k not found>" } })
Assert-Case "Step 1's kill and -Close's were found, and every command they call is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $kc = Invoke-Isolated -Name 'the kill counts' -Block {
        . $frSetup
        $o = [ordered]@{}
        foreach ($k in @($killIfs.Keys)) {
            Reset-Fr
            $script:FrFails['KILL claude 22'] = $true
            $two = @([pscustomobject]@{ Name = 'claude'; Id = 21 }, [pscustomobject]@{ Name = 'claude'; Id = 22 })
            $claudeProcs = $two
            $procs = $two
            $procsUnread = $null
            $seg = [scriptblock]::Create($killIfs[$k].Extent.Text)
            & { . $seg }
            $o[$k] = Get-FrSaid
        }
        $o
    }
    if ($kc) {
        Assert-Case "Step 1, one kill of two refused: one counted, and the other said" `
            ($kc['Step 1'] -match 'LOG: Killed 1 Claude process\(es\)' -and $kc['Step 1'] -match 'LOG: \[!\] 1 could not be stopped \(Access is denied\); Step 4 looks again') $kc['Step 1']
        Assert-Case "-Close, one kill of two refused: one counted, and the other said" `
            ($kc['-Close'] -match 'LOG: Killed 1 Claude process\(es\)' -and $kc['-Close'] -match 'LOG: \[!\] 1 could not be stopped \(Access is denied\)') $kc['-Close']
    }
}

# --- Step 2 ------------------------------------------------------------
# A service that could not be read is not stopped, and a stop that could
# not be confirmed is not followed by a force-kill. Until 6.3.1 the first
# was "Service not found", and the second "Service stopped gracefully".
$step2Text = Get-FixSegment -From 'Log "[2/10] Stopping $ServiceName..."' -To 'Log "[3/10] Checking HCS service health..."'
$unfaked = @(Get-FrUnsafe -Text $step2Text)
Assert-Case "Step 2 was found, and every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $s2 = Invoke-Isolated -Name 'Step 2' -Block {
        . $frSetup
        $seg = [scriptblock]::Create($step2Text)
        $script:IsAdmin = $true
        $ServiceTimeout = 1
        $o = [ordered]@{}
        foreach ($c in 'unreadStart', 'unreadPoll', 'stuck', 'killUnread') {
            Reset-Fr
            switch ($c) {
                'unreadStart' { $script:FrSvc['CoworkVMService'] = 'Denied' }
                'unreadPoll'  { $script:FrSvc['CoworkVMService'] = [System.Collections.ArrayList]@('Running', 'Denied') }
                'stuck'       { $script:FrSvc['CoworkVMService'] = 'Running' }
                'killUnread'  { $script:FrSvc['CoworkVMService'] = 'Running'; $script:FrProcs['cowork-svc'] = 'Denied' }
            }
            & { . $seg }
            $o[$c] = Get-FrSaid
        }
        $o
    }
    if ($s2) {
        Assert-Case "Step 2, the service could not be read: not stopped, and said" `
            ($s2['unreadStart'] -match 'LOG: \[!\] CoworkVMService could not be read \(Access is denied\), so it was not stopped' -and
             $s2['unreadStart'] -notmatch 'STOP |KILL |not found') $s2['unreadStart']
        Assert-Case "Step 2, whether it stopped could not be read: not force-killed, not called stopped, and said" `
            ($s2['unreadPoll'] -match 'STOP CoworkVMService' -and $s2['unreadPoll'] -notmatch 'stopped gracefully|KILL ' -and
             $s2['unreadPoll'] -match 'LOG: \[!\] Whether CoworkVMService stopped could not be read \(Access is denied\), so cowork-svc was not force-killed') $s2['unreadPoll']
        Assert-Case "Step 2, a service that does not stop: force-killed" `
            ($s2['stuck'] -match 'KILL cowork-svc' -and $s2['stuck'] -match 'LOG: Force-killed') $s2['stuck']
        Assert-Case "Step 2, whether the force-kill took could not be read: not called force-killed, and said" `
            ($s2['killUnread'] -match 'KILL cowork-svc' -and $s2['killUnread'] -notmatch 'LOG: Force-killed' -and
             $s2['killUnread'] -match 'LOG: \[!\] Whether cowork-svc is still running could not be read \(Access is denied\)') $s2['killUnread']
    }
}

# --- Step 3 ------------------------------------------------------------
# A service that could not be read is not restarted, and nothing further
# is escalated on it: HvHost's restart touches every Hyper-V VM. Until
# 6.3.1 vmms that could not be read was "not found", and a vmcompute that
# could not be read was down, so the escalation went on to HvHost.
$step3Text = Get-FixSegment -From 'Log "[3/10] Checking HCS service health..."' -To 'Log "[4/10] Checking for orphan processes..."'
$unfaked = @(Get-FrUnsafe -Text $step3Text -Also @('Test-RecentHcsErrors'))
Assert-Case "Step 3 was found, and every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $s3 = Invoke-Isolated -Name 'Step 3' -Block {
        . $frSetup
        function Test-RecentHcsErrors { return 'construct_failure' }
        $seg = [scriptblock]::Create($step3Text)
        $script:IsAdmin = $true
        $o = [ordered]@{}
        $five = @('Stopped', 'Stopped', 'Stopped', 'Stopped', 'Stopped')
        foreach ($c in @(
                @{ N = 'vmcUnread';      Mode = 'Deep';  Vmc = @('Denied');           Vmms = 'Running'; Hv = 'Running' },
                @{ N = 'vmmsUnread';     Mode = 'Deep';  Vmc = @('Stopped');          Vmms = 'Denied';  Hv = 'Running' },
                @{ N = 'vmcAfterUnread'; Mode = 'Deep';  Vmc = ($five + 'Denied');   Vmms = 'Running'; Hv = 'Running' },
                @{ N = 'hvUnread';       Mode = 'Deep';  Vmc = @('Stopped');          Vmms = 'Running'; Hv = 'Denied' },
                @{ N = 'vmmsError';      Mode = 'Smart'; Vmc = @('Stopped');          Vmms = 'Running'; Hv = 'Running'; Fail = 'STOP vmms' },
                @{ N = 'deep';           Mode = 'Deep';  Vmc = @('Stopped');          Vmms = 'Running'; Hv = 'Running' })) {
            Reset-Fr
            $script:SelectedMode = $c.Mode
            $script:FrSvc['vmcompute'] = [System.Collections.ArrayList]@($c.Vmc)
            $script:FrSvc['vmms'] = $c.Vmms
            $script:FrSvc['HvHost'] = $c.Hv
            if ($c['Fail']) { $script:FrFails[$c.Fail] = $true }
            & { . $seg }
            $o[$c.N] = Get-FrSaid
        }
        $o
    }
    if ($s3) {
        Assert-Case "Step 3, whether vmcompute came back could not be read: nothing escalated, and said" `
            ($s3['vmcUnread'] -match 'LOG: \[!\] Whether vmcompute came back could not be read \(Access is denied\), so nothing was escalated' -and
             $s3['vmcUnread'] -notmatch 'vmms|HvHost') $s3['vmcUnread']
        Assert-Case "Step 3, vmms could not be read: neither it nor HvHost restarted, and said" `
            ($s3['vmmsUnread'] -match 'LOG: \[!\] vmms could not be read \(Access is denied\), so it was not restarted and nothing further was escalated' -and
             $s3['vmmsUnread'] -notmatch '(STOP|START|RESTART) (vmms|HvHost)|not found') $s3['vmmsUnread']
        Assert-Case "Step 3, vmcompute after the vmms restart could not be read: HvHost not restarted, and said" `
            ($s3['vmcAfterUnread'] -match 'STOP vmms \| START vmms' -and $s3['vmcAfterUnread'] -notmatch 'RESTART HvHost' -and
             $s3['vmcAfterUnread'] -match 'LOG: \[!\] Whether vmcompute came back after the vmms restart could not be read \(Access is denied\)') $s3['vmcAfterUnread']
        Assert-Case "Step 3, HvHost could not be read: not restarted, and said" `
            ($s3['hvUnread'] -match 'START vmms' -and $s3['hvUnread'] -notmatch 'RESTART HvHost|not found' -and
             $s3['hvUnread'] -match 'LOG: \[!\] HvHost could not be read \(Access is denied\), so it was not restarted') $s3['hvUnread']
        Assert-Case "Step 3, a vmms restart that reported an error: not called restarted, and said" `
            ($s3['vmmsError'] -match 'LOG: \[!\] The vmms restart reported an error \(Access is denied\)' -and
             $s3['vmmsError'] -notmatch 'LOG: vmms service restarted') $s3['vmmsError']
        Assert-Case "Step 3, Deep, vmcompute still down after vmms: HvHost restarted" `
            ($s3['deep'] -match 'LOG: vmms service restarted' -and $s3['deep'] -match 'RESTART HvHost' -and $s3['deep'] -match 'LOG: HvHost service restarted') $s3['deep']
    }
}

# --- Smart mode's escalation to Deep ------------------------------------
# A service that could not be read is not "not running": no purge of the
# VM cache on it. Until 6.3.1 it escalated.
# With -KeepCache the escalation, a purge of the VM cache, does not run, and
# that is said (6.3.1), but not for a service the prerequisite check found
# cannot start, whose cause is said instead. The two conditions after the
# check, the escalation's and the line asking for admin, are lifted and
# evaluated as admin and not.
$smartText = Get-FixSegment -From '$smartSvc = ' -To 'if (-not $prereqBlocked -and $smartSvcDown'
$smartEscIf = @(Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and
    $n.Clauses[0].Item1.Extent.Text -match '\$smartSvcDown' -and $n.Clauses[0].Item1.Extent.Text -match '\$script:IsAdmin -and -not \$script:DeepEscalated$' })
$smartAdminIf = @(Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and
    $n.Clauses[0].Item1.Extent.Text -match '\$smartSvcDown' -and $n.Clauses[0].Item1.Extent.Text -match '-not \$script:IsAdmin$' })
$unfaked = @(Get-FrUnsafe -Text $smartText)
Assert-Case "Smart mode's escalation check was found, and every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
Assert-Case "Smart mode's escalation and its line asking for admin were found" ($smartEscIf.Count -eq 1 -and $smartAdminIf.Count -eq 1) `
    "found $($smartEscIf.Count), $($smartAdminIf.Count)"
if ($unfaked.Count -eq 0 -and $smartEscIf.Count -eq 1 -and $smartAdminIf.Count -eq 1) {
    $smartEscCond = $smartEscIf[0].Clauses[0].Item1.Extent.Text
    $smartAdminCond = $smartAdminIf[0].Clauses[0].Item1.Extent.Text
    $sm = Invoke-Isolated -Name "Smart mode's escalation check" -Block {
        . $frSetup
        $seg = [scriptblock]::Create($smartText)
        $escCond = [scriptblock]::Create($smartEscCond)
        $adminCond = [scriptblock]::Create($smartAdminCond)
        $o = [ordered]@{}
        foreach ($c in 'Denied', 'Stopped', 'Running', 'Absent', 'DeniedKeep', 'StoppedKeep', 'RunningKeep', 'AbsentKeep',
                 'StoppedBlocked', 'StoppedBlockedKeep') {
            Reset-Fr
            $svc = $c -replace 'Keep$', '' -replace 'Blocked$', ''
            if ($svc -ne 'Absent') { $script:FrSvc['CoworkVMService'] = $svc }
            $KeepCache = $c.EndsWith('Keep')
            # A service the prerequisite check found cannot start.
            $script:ServicePrereq = $null
            if ($c -match 'Blocked') { $script:ServicePrereq = [pscustomobject]@{ Blocked = $true; Reason = 'CoworkVMService start type is Disabled.' } }
            & {
                . $seg
                $script:FrVal = "$smartSvcDown"
                $script:DeepEscalated = $false
                $script:IsAdmin = $true
                $script:FrEsc = [bool](. $escCond)
                $script:IsAdmin = $false
                $script:FrAdmin = [bool](. $adminCond)
            }
            $o[$c] = "$($script:FrVal) | $(Get-FrSaid)"
            $o["$c|go"] = "escalate $($script:FrEsc), ask for admin $($script:FrAdmin)"
        }
        $script:ServicePrereq = $null
        $o
    }
    if ($sm) {
        Assert-Case "Smart, the service could not be read: not escalated, and said" `
            ($sm['Denied'] -eq "False | LOG: Smart mode: not escalating to Deep, the service's state could not be read (Access is denied)" -and
             $sm['Denied|go'] -eq 'escalate False, ask for admin False') "$($sm['Denied']); $($sm['Denied|go'])"
        Assert-Case "Smart, the service stopped or not there: escalated" `
            ($sm['Stopped'] -eq 'True | ' -and $sm['Absent'] -eq 'True | ' -and
             $sm['Stopped|go'] -eq 'escalate True, ask for admin True' -and $sm['Absent|go'] -eq 'escalate True, ask for admin True') `
            "[$($sm['Stopped'])] [$($sm['Absent'])] $($sm['Stopped|go']); $($sm['Absent|go'])"
        Assert-Case "Smart, the service running: not escalated" `
            ($sm['Running'] -eq 'False | ' -and $sm['Running|go'] -eq 'escalate False, ask for admin False') "$($sm['Running']); $($sm['Running|go'])"
        $keepSaid = 'True | LOG: Smart mode: service not running, but not escalating to Deep: -KeepCache keeps the VM cache. Rerun without it to purge the cache.'
        Assert-Case "Smart with -KeepCache, the service stopped or not there: not escalated, admin not asked for, and said" `
            ($sm['StoppedKeep'] -eq $keepSaid -and $sm['AbsentKeep'] -eq $keepSaid -and
             $sm['StoppedKeep|go'] -eq 'escalate False, ask for admin False' -and $sm['AbsentKeep|go'] -eq 'escalate False, ask for admin False') `
            "[$($sm['StoppedKeep'])] [$($sm['AbsentKeep'])] $($sm['StoppedKeep|go']); $($sm['AbsentKeep|go'])"
        Assert-Case "Smart with -KeepCache, the service running or not read: nothing more said" `
            ($sm['RunningKeep'] -eq 'False | ' -and $sm['DeniedKeep'] -eq $sm['Denied'] -and
             $sm['RunningKeep|go'] -eq 'escalate False, ask for admin False' -and $sm['DeniedKeep|go'] -eq 'escalate False, ask for admin False') `
            "[$($sm['RunningKeep'])] [$($sm['DeniedKeep'])] $($sm['RunningKeep|go']); $($sm['DeniedKeep|go'])"
        $blockedSaid = ('True | LOG: Smart mode: not escalating to Deep, CoworkVMService start type is Disabled. | ' +
            'LOG: Purging the VM cache cannot fix that. Resolve it and rerun.')
        Assert-Case "Smart, a service that cannot start: not escalated, admin not asked for, and the cause said; with -KeepCache too, without the line for -KeepCache" `
            ($sm['StoppedBlocked'] -eq $blockedSaid -and $sm['StoppedBlockedKeep'] -eq $blockedSaid -and
             $sm['StoppedBlocked|go'] -eq 'escalate False, ask for admin False' -and $sm['StoppedBlockedKeep|go'] -eq 'escalate False, ask for admin False') `
            "[$($sm['StoppedBlocked'])] [$($sm['StoppedBlockedKeep'])] $($sm['StoppedBlocked|go']); $($sm['StoppedBlockedKeep|go'])"
    }
}

# --- Step 9 ------------------------------------------------------------
# The wait is skipped for a blocked prerequisite only when the service is
# known to be down, and the service is restarted in the wait only when it
# is known not to run. Until 6.3.1 a service that could not be read was
# down on both counts, and was restarted on every pass. A restart in the
# wait that fails on a blocked prerequisite stops the wait and sets the
# flag, as a skipped wait does; until 6.3.1 it stopped the wait only.
$waitStart = @(Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$svcAtWaitStart' })
# The flag as the wait begins, beside $svcAtWaitStart; a restart in the wait
# that fails sets it too, further in.
$stillBlocked = @(Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$prereqStillBlocked' } |
    Where-Object { $waitStart.Count -eq 1 -and $_.Parent -eq $waitStart[0].Parent })
$svcNowAssign = @()
if ($waitStart.Count -eq 1) {
    $svcNowAssign = @(Find-FixStatement { param($n)
        $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$svcNow' -and
        $n.Extent.StartOffset -gt $waitStart[0].Extent.StartOffset } | Sort-Object { $_.Extent.StartOffset } | Select-Object -First 1)
}
$waitText = $null
$svcNowLoop = $null
if ($waitStart.Count -eq 1 -and $stillBlocked.Count -eq 1 -and $svcNowAssign.Count -eq 1) {
    $waitText = $fixSrc.Substring($waitStart[0].Extent.StartOffset, $stillBlocked[0].Extent.EndOffset - $waitStart[0].Extent.StartOffset)
    $nowParent = @($svcNowAssign[0].Parent.Statements)
    $nowAt = [array]::IndexOf($nowParent, $svcNowAssign[0])
    $svcNowLoop = (@('$svcRestartAttempts = 0', '$svcRestartLimit = 3', '$svcNowUnreadSaid = $false', '$prereqStillBlocked = $false', 'foreach ($frPass in 1..3) {',
            $svcNowAssign[0].Extent.Text, $nowParent[$nowAt + 1].Extent.Text, '    [void]$script:FrCalls.Add(''PASSED'')', '}') -join "`r`n")
}
$unfaked = @(if ($waitText) { Get-FrUnsafe -Text ($waitText + "`r`n" + $svcNowLoop) -Also @('Restart-CoworkService') } else { '<not found>' })
Assert-Case "Step 9's service reads were found, and every command they call is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $s9 = Invoke-Isolated -Name "Step 9's service reads" -Block {
        . $frSetup
        # A restart fails with $script:FrRestartFails set.
        function Restart-CoworkService { [void]$script:FrCalls.Add('RESTART'); return (-not $script:FrRestartFails) }
        $script:FrRestartFails = $false
        $waitSeg = [scriptblock]::Create($waitText)
        $loopSeg = [scriptblock]::Create($svcNowLoop)
        $o = [ordered]@{}
        foreach ($c in 'Denied', 'Stopped', 'Running') {
            Reset-Fr
            $script:ServicePrereq = [pscustomobject]@{ Blocked = $true; Reason = 'blocked' }
            $script:FrSvc['CoworkVMService'] = $c
            & { . $waitSeg; $script:FrVal = "$prereqStillBlocked" }
            $o["wait-$c"] = $script:FrVal
            Reset-Fr
            $script:ServicePrereq = $null
            $script:FrSvc['CoworkVMService'] = $c
            & { . $loopSeg; $script:FrVal = "$prereqStillBlocked" }
            $o["loop-$c"] = Get-FrSaid
            $o["loopFlag-$c"] = $script:FrVal
        }
        # A restart in the wait that fails on a cause the check found.
        Reset-Fr
        $script:ServicePrereq = [pscustomobject]@{ Blocked = $true; Reason = 'blocked' }
        $script:FrSvc['CoworkVMService'] = 'Stopped'
        $script:FrRestartFails = $true
        & { . $loopSeg; $script:FrVal = "$prereqStillBlocked" }
        $o['loop-Blocked'] = "$($script:FrVal) | $(Get-FrSaid)"
        # A restart in the wait that fails with no cause found, the check not
        # run or its flag clear: the wait goes on, and the flag stays clear.
        foreach ($frCause in 'None', 'Clear') {
            Reset-Fr
            $script:ServicePrereq = if ($frCause -eq 'Clear') { [pscustomobject]@{ Blocked = $false; Reason = $null } } else { $null }
            $script:FrSvc['CoworkVMService'] = 'Stopped'
            $script:FrRestartFails = $true
            & { . $loopSeg; $script:FrVal = "$prereqStillBlocked" }
            $o["loopFail-$frCause"] = "$($script:FrVal) | $(Get-FrSaid)"
        }
        $script:FrRestartFails = $false
        $script:ServicePrereq = $null
        $o
    }
    if ($s9) {
        Assert-Case "Step 9, a blocked prerequisite and a service that could not be read: the wait runs" ($s9['wait-Denied'] -eq 'False') $s9['wait-Denied']
        Assert-Case "Step 9, a blocked prerequisite and the service down: no wait" ($s9['wait-Stopped'] -eq 'True') $s9['wait-Stopped']
        Assert-Case "Step 9, a blocked prerequisite and the service running: the wait runs" ($s9['wait-Running'] -eq 'False') $s9['wait-Running']
        Assert-Case "Step 9, the service could not be read on each pass: not restarted, and said once" `
            ($s9['loop-Denied'] -eq ("LOG: [!] The service's state could not be read (Access is denied); it is not restarted on that | " +
                'PASSED | PASSED | PASSED')) $s9['loop-Denied']
        Assert-Case "Step 9, the service down: restarted" ($s9['loop-Stopped'] -match '^(LOG: Service died, restarting \(\d of 3\)\.\.\. \| ){3}RESTART \| RESTART \| RESTART$') $s9['loop-Stopped']
        Assert-Case "Step 9, the service running: not restarted" ($s9['loop-Running'] -eq 'PASSED | PASSED | PASSED') $s9['loop-Running']
        Assert-Case "Step 9, the service restarted in the wait, running or not read: the flag that holds back what follows Step 9 is not set" `
            ($s9['loopFlag-Denied'] -eq 'False' -and $s9['loopFlag-Stopped'] -eq 'False' -and $s9['loopFlag-Running'] -eq 'False') `
            "$($s9['loopFlag-Denied']), $($s9['loopFlag-Stopped']), $($s9['loopFlag-Running'])"
        Assert-Case "Step 9, a restart in the wait that fails on a cause the check found: the wait stops, and the flag that holds back what follows Step 9 is set" `
            ($s9['loop-Blocked'] -eq ('True | LOG: Service died, restarting (1 of 3)... | ' +
                'LOG: This needs to be fixed on the machine first; not retrying. | RESTART')) $s9['loop-Blocked']
        $failRx = '^False \| (LOG: Service died, restarting \(\d of 3\)\.\.\. \| ){3}RESTART \| RESTART \| RESTART$'
        Assert-Case "Step 9, a restart in the wait that fails with no cause found, the check not run or its flag clear: the wait goes on, and the flag that holds back what follows Step 9 is not set" `
            ($s9['loopFail-None'] -match $failRx -and $s9['loopFail-Clear'] -match $failRx) `
            "$($s9['loopFail-None']); $($s9['loopFail-Clear'])"
    }
}

# --- -Close --------------------------------------------------------------
# Step 1: a service that could not be read is not stopped; without admin,
# "Killed" is said only when a kill went out. Step 4: whether the service
# came back could not be read is said as that. Step 5: whether the service
# runs is not known, rather than not running. Until 6.3.1 each read as the
# service not there or not running, and Step 1 said "Killed" whatever
# happened.
$close1Text = Get-FixSegment -From 'Log "[1/5] Stopping CoworkVMService (graceful)..."' -To '# 2) Kill the Claude UI'
$close4Text = Get-FixSegment -From 'Log "[4/5] Restarting service (idle, ready for next launch)..."' -To '# 5) Verify clean state'
$close5Text = Get-FixSegment -From '$remainingSvc = ' -To '$svcRunning = ' -ToLineEnd
$unfaked = @(@($close1Text, $close4Text, $close5Text) | ForEach-Object { if ($_) { Get-FrUnsafe -Text $_ -Also @('New-Object') } else { '<not found>' } })
Assert-Case "-Close's steps 1, 4 and 5 were found, and every command they call is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $cl = Invoke-Isolated -Name "-Close's service reads" -Block {
        . $frSetup
        # A tripwire: the stop through a ServiceController is recorded.
        function New-Object { param([string]$TypeName, [object[]]$ArgumentList) [void]$script:FrCalls.Add("NEW $TypeName $ArgumentList"); return $null }
        $o = [ordered]@{}
        foreach ($c in @(
                @{ N = 'step1-unread';    Seg = $close1Text; Admin = $true;  Svc = 'Denied' },
                @{ N = 'step1-none';      Seg = $close1Text; Admin = $false; Svc = 'Running' },
                @{ N = 'step1-one';       Seg = $close1Text; Admin = $false; Svc = 'Running'; Procs = @(301) },
                @{ N = 'step1-listUnread'; Seg = $close1Text; Admin = $false; Svc = 'Running'; Procs = 'Denied' },
                @{ N = 'step4-unread';    Seg = $close4Text; Admin = $true;  Svc = 'Denied' },
                @{ N = 'step4-running';   Seg = $close4Text; Admin = $true;  Svc = 'Running' },
                @{ N = 'step5-unread';    Seg = $close5Text; Admin = $true;  Svc = 'Denied' },
                @{ N = 'step5-running';   Seg = $close5Text; Admin = $true;  Svc = 'Running' })) {
            Reset-Fr
            $script:IsAdmin = $c.Admin
            $script:FrSvc['CoworkVMService'] = $c.Svc
            if ($null -ne $c['Procs']) { $script:FrProcs['cowork-svc'] = $c.Procs }
            $seg = [scriptblock]::Create($c.Seg)
            & { . $seg; $script:FrVal = "$($ExecutionContext.SessionState.PSVariable.GetValue('svcRunning', 'unset'))|$($ExecutionContext.SessionState.PSVariable.GetValue('svcUnknown', 'unset'))" }
            $o[$c.N] = Get-FrSaid
            if ($c.N -like 'step5-*') { $o[$c.N] = "$($script:FrVal) | $($o[$c.N])" }
        }
        $o
    }
    if ($cl) {
        Assert-Case "-Close step 1, the service could not be read: not stopped, and said" `
            ($cl['step1-unread'] -eq ('LOG: [1/5] Stopping CoworkVMService (graceful)... | ' +
                'LOG: [!] CoworkVMService could not be read (Access is denied), so it was not stopped')) $cl['step1-unread']
        Assert-Case "-Close step 1 without admin, no service process: not called killed" `
            ($cl['step1-none'] -match 'LOG: No service process was running' -and $cl['step1-none'] -notmatch 'Killed') $cl['step1-none']
        Assert-Case "-Close step 1 without admin, one service process: killed" `
            ($cl['step1-one'] -match 'KILL cowork-svc 301' -and $cl['step1-one'] -match 'LOG: Killed service process \(no admin\)') $cl['step1-one']
        Assert-Case "-Close step 1 without admin, the list could not be read: not called killed, and said" `
            ($cl['step1-listUnread'] -cnotmatch 'Killed service|KILL ' -and
             $cl['step1-listUnread'] -match "LOG: \[!\] cowork-svc's processes could not be listed \(Access is denied\), so none were force-killed") $cl['step1-listUnread']
        Assert-Case "-Close step 4, whether the service came back could not be read: said as that" `
            ($cl['step4-unread'] -match 'START CoworkVMService' -and
             $cl['step4-unread'] -match 'LOG: Whether the service reached Running could not be read \(Access is denied\), may need Fix on relaunch' -and
             $cl['step4-unread'] -notmatch 'did not reach Running') $cl['step4-unread']
        Assert-Case "-Close step 4, the service came back: running" ($cl['step4-running'] -match 'LOG: Service running \(idle\)') $cl['step4-running']
        Assert-Case "-Close step 5, the service could not be read: not known, not 'not running', and said" `
            ($cl['step5-unread'] -match '^False\|True \| ' -and
             $cl['step5-unread'] -match 'LOG: Whether the service is running is not known: it could not be read \(Access is denied\)') $cl['step5-unread']
        Assert-Case "-Close step 5, the service running: running" ($cl['step5-running'] -match '^True\|False \| ') $cl['step5-running']
    }
}

# --- Diagnostics -----------------------------------------------------------
# A service that could not be read says so. Until 6.3.1 it was "Not
# installed", and vmcompute's line was left out. Defender's exclusions that
# could not be read are said as that; until 6.3.1 every one was MISSING.
$diagText = Get-FixSegment -From '$diagSvc = ' -To '# HCS errors'
$defText  = Get-FixSegment -From '# Defender exclusion completeness (v4.8.0)' -To '# -- Support bundle'
$unfaked = @(@($diagText, $defText) | ForEach-Object { if ($_) { Get-FrUnsafe -Text $_ -Also @('Get-MpPreference') } else { '<not found>' } })
Assert-Case "Diagnostics' service and Defender reads were found, and every command they call is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $dg = Invoke-Isolated -Name "Diagnostics' reads" -Block {
        . $frSetup
        function Get-MpPreference {
            [CmdletBinding()]
            param()
            if ($script:FrMpDenied) { $PSCmdlet.WriteError((New-FrDenied 'MSFT_MpPreference')); return }
            return [pscustomobject]@{ ExclusionProcess = @('vmwp.exe', 'vmms.exe', 'vmcompute.exe', 'cowork-svc.exe') }
        }
        $o = [ordered]@{}
        foreach ($c in @(@{ N = 'svcDenied'; Seg = $diagText; Svc = 'Denied' }, @{ N = 'svcAbsent'; Seg = $diagText; Svc = $null },
                         @{ N = 'defDenied'; Seg = $defText; Admin = $true; Mp = $true },
                         @{ N = 'defUser'; Seg = $defText; Admin = $false; Mp = $true },
                         @{ N = 'defOk'; Seg = $defText; Admin = $true; Mp = $false })) {
            Reset-Fr
            if ($c['Svc']) { $script:FrSvc['CoworkVMService'] = $c.Svc; $script:FrSvc['vmcompute'] = $c.Svc }
            $script:IsAdmin = [bool]$c['Admin']
            $script:FrMpDenied = [bool]$c['Mp']
            $seg = [scriptblock]::Create($c.Seg)
            & { . $seg }
            $o[$c.N] = Get-FrSaid
        }
        $o
    }
    if ($dg) {
        Assert-Case "Diagnostics, the services could not be read: said, not 'Not installed'" `
            ($dg['svcDenied'] -eq 'LOG:   CoworkVMService : could not be read (Access is denied) | LOG:   vmcompute     : could not be read (Access is denied)') $dg['svcDenied']
        Assert-Case "Diagnostics, no service: Not installed" ($dg['svcAbsent'] -eq 'LOG:   CoworkVMService : Not installed') $dg['svcAbsent']
        Assert-Case "Diagnostics, the exclusions could not be read: said, not MISSING" `
            ($dg['defDenied'] -eq 'LOG:   Defender procs: could not be read (Access is denied)') $dg['defDenied']
        Assert-Case "Diagnostics without admin: the exclusions are not checked" `
            ($dg['defUser'] -eq 'LOG:   Defender procs: not checked (needs admin to read the exclusions)') $dg['defUser']
        Assert-Case "Diagnostics, the exclusions read: all present" ($dg['defOk'] -eq 'LOG:   Defender procs: All exclusions present') $dg['defOk']
    }
}

# --- Step 8 and Start-ClaudeDesktop: the LaunchClaudeAdmin task ------------
# A task that could not be read is said, and the launch falls back. Until
# 6.3.1 it was "not found", or nothing was said.
$elevAt = $fixSrc.IndexOf('$elevTaskPath = "\Claude\LaunchClaudeAdmin"', [System.StringComparison]::Ordinal)
$step8Try = @()
if ($elevAt -ge 0) {
    $step8Try = @(Find-FixStatement { param($n)
        $n -is [System.Management.Automation.Language.TryStatementAst] -and $n.Extent.StartOffset -gt $elevAt -and
        $n.Body.Extent.Text -match 'Start-ScheduledTask -TaskName "LaunchClaudeAdmin"' } | Sort-Object { $_.Extent.StartOffset } | Select-Object -First 1)
}
$step8Text = if ($step8Try.Count -eq 1) { $fixSrc.Substring($elevAt, $step8Try[0].Extent.EndOffset - $elevAt) } else { $null }
$scdDef = Get-AstFunction -Ast $fixAst -Name 'Start-ClaudeDesktop'
$taskFakes = @('Get-ScheduledTask', 'Start-ScheduledTask', 'Get-ClaudeDesktopProcess', 'Find-ClaudeExe')
$unfaked = @(if ($step8Text -and $scdDef) { Get-FrUnsafe -Text $step8Text -Defs @($scdDef) -Also $taskFakes } else { '<not found>' })
Assert-Case "Step 8's task read and Start-ClaudeDesktop were found, and every command they call is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $tk = Invoke-Isolated -Name 'the LaunchClaudeAdmin task reads' -Block {
        . $frSetup
        function Get-ScheduledTask {
            [CmdletBinding()]
            param([string]$TaskName, [string]$TaskPath)
            [void]$script:FrCalls.Add("TASK READ $TaskPath$TaskName")
            $PSCmdlet.WriteError((New-FrDenied "$TaskPath$TaskName"))
        }
        function Start-ScheduledTask { [CmdletBinding()] param([string]$TaskName, [string]$TaskPath) [void]$script:FrCalls.Add("START-TASK $TaskPath$TaskName") }
        function Get-ClaudeDesktopProcess { param($Environment) $null = $Environment; return @() }
        function Find-ClaudeExe { return $null }
        . ([scriptblock]::Create($scdDef.Extent.Text))
        $script:ClaudeEnv = [pscustomobject]@{ IsMsix = $true; Aumid = 'shell:AppsFolder\Claude_x!Claude'; PackageFamilyName = 'Claude_x' }
        $o = [ordered]@{}
        Reset-Fr
        $launched = $false
        $seg = [scriptblock]::Create($step8Text)
        & { . $seg }
        $o['step8'] = Get-FrSaid
        Reset-Fr
        $ok = Start-ClaudeDesktop
        $o['start'] = "returned $ok | $(Get-FrSaid)"
        $o
    }
    if ($tk) {
        Assert-Case "Step 8, the task could not be read: said, not used, and the launch falls back" `
            ($tk['step8'] -match 'LOG: Whether the LaunchClaudeAdmin task exists could not be read \(Access is denied\), falling back to standard launch' -and
             $tk['step8'] -notmatch 'START-TASK|not found') $tk['step8']
        Assert-Case "Start-ClaudeDesktop, the task could not be read: said, and the app started another way" `
            ($tk['start'] -match '^returned True \| LOG: Whether the LaunchClaudeAdmin task exists could not be read \(Access is denied\), so it was not used' -and
             $tk['start'] -notmatch 'START-TASK' -and $tk['start'] -match 'START-PROCESS shell:AppsFolder\\Claude_x!Claude') $tk['start']
    }
}

# --- the service helpers ---------------------------------------------------
# Stop-CoworkServiceWithTimeout: a stop that could not be confirmed is not
# followed by a force-kill; its job runs here in-process, against the same
# fakes. Wait-ForServiceProcessExit: a list that could not be read is not a
# process that exited. Test-CoworkServicePrereq: a service that could not
# be read is not "not registered". Until 6.3.1 each read as nothing there.
$fnNames = @('Stop-CoworkServiceWithTimeout', 'Wait-ForServiceProcessExit', 'Test-CoworkServicePrereq')
$fnDefs = @($fnNames | ForEach-Object { Get-AstFunction -Ast $fixAst -Name $_ } | Where-Object { $_ })
$fnFakes = @('Start-Job', 'Wait-Job', 'Receive-Job', 'Stop-Job', 'Remove-Job', 'Get-CimInstance', 'Get-AppxPackage', 'Set-Service',
    'Get-WindowsOptionalFeature', 'bcdedit')
$unfaked = @(if ($fnDefs.Count -eq 3) { Get-FrUnsafe -Defs $fnDefs -Also $fnFakes } else { '<not found>' })
Assert-Case "Stop-CoworkServiceWithTimeout, Wait-ForServiceProcessExit and Test-CoworkServicePrereq were found, and every command they call is faked" `
    ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $sh = Invoke-Isolated -Name 'the service helpers' -Block {
        . $frSetup
        function Start-Job {
            [CmdletBinding()]
            param([scriptblock]$ScriptBlock, [object[]]$ArgumentList)
            return [pscustomobject]@{ Out = @(& $ScriptBlock @ArgumentList) }
        }
        function Wait-Job { [CmdletBinding()] param([Parameter(Position = 0)]$Job, [int]$Timeout) $null = $Timeout; return $Job }
        function Receive-Job { [CmdletBinding()] param([Parameter(Position = 0)]$Job) return $Job.Out }
        function Stop-Job { [CmdletBinding()] param([Parameter(Position = 0)]$Job) $null = $Job }
        function Remove-Job { [CmdletBinding()] param([Parameter(Position = 0)]$Job, [switch]$Force) $null = $Job, $Force }
        function Get-CimInstance { [CmdletBinding()] param([Parameter(Position = 0)]$ClassName, $Filter) $null = $ClassName, $Filter; return [pscustomobject]@{ StartMode = 'Auto' } }
        function Get-AppxPackage { [CmdletBinding()] param($Name) $null = $Name; return [pscustomobject]@{ Name = 'Claude'; PackageFullName = 'Claude_1' } }
        function Set-Service { [CmdletBinding()] param($Name, $StartupType) [void]$script:FrCalls.Add("SET $Name $StartupType") }
        function Get-WindowsOptionalFeature { [CmdletBinding()] param([switch]$Online, $FeatureName) $null = $Online; [void]$script:FrCalls.Add("FEATURE $FeatureName") }
        function bcdedit { [void]$script:FrCalls.Add('BCDEDIT') }
        foreach ($d in $fnDefs) { . ([scriptblock]::Create($d.Extent.Text)) }
        $script:IsAdmin = $false
        $o = [ordered]@{}
        foreach ($c in @(@{ N = 'jobUnread'; Svc = @('Denied') }, @{ N = 'stuck'; Svc = @('Running') },
                         @{ N = 'killed'; Svc = @('Running', 'Stopped') }, @{ N = 'afterUnread'; Svc = @('Running', 'Denied') })) {
            Reset-Fr
            $script:FrSvc['CoworkVMService'] = [System.Collections.ArrayList]@($c.Svc)
            $script:FrProcs['cowork-svc'] = @(301)
            $r = Stop-CoworkServiceWithTimeout -TimeoutSeconds 5
            $o["stop-$($c.N)"] = "returned $r | $(Get-FrSaid)"
        }
        foreach ($c in @(@{ N = 'denied'; P = 'Denied' }, @{ N = 'none'; P = $null }, @{ N = 'stays'; P = @(301) })) {
            Reset-Fr
            if ($null -ne $c.P) { $script:FrProcs['cowork-svc'] = $c.P }
            $r = Wait-ForServiceProcessExit -TimeoutSeconds 2
            $o["wait-$($c.N)"] = "returned $r | $((@($script:FrLog) | ForEach-Object { "LOG: $_" }) -join ' | ')"
        }
        foreach ($c in @(@{ N = 'hcsUnread'; Vmc = 'Denied'; Svc = 'Running' }, @{ N = 'svcUnread'; Vmc = 'Running'; Svc = 'Denied' },
                         @{ N = 'svcAbsent'; Vmc = 'Running'; Svc = $null }, @{ N = 'ok'; Vmc = 'Running'; Svc = 'Running' })) {
            Reset-Fr
            $script:FrSvc['vmcompute'] = $c.Vmc
            $script:FrSvc['hns'] = 'Running'
            if ($c.Svc) { $script:FrSvc['CoworkVMService'] = $c.Svc }
            $r = Test-CoworkServicePrereq
            $o["prereq-$($c.N)"] = "$($r.Blocked)|$($r.Found)|$($r.Reason)"
        }
        $o
    }
    if ($sh) {
        Assert-Case "Stop-CoworkServiceWithTimeout, whether it stopped could not be read: not force-killed, false, and said" `
            ($sh['stop-jobUnread'] -match '^returned False \| LOG: \[!\] Whether CoworkVMService stopped could not be read \(Access is denied\), so cowork-svc was not force-killed' -and
             $sh['stop-jobUnread'] -cnotmatch 'LIST cowork-svc|KILL ') $sh['stop-jobUnread']
        Assert-Case "Stop-CoworkServiceWithTimeout, a service that does not stop: force-killed, and still running is said" `
            ($sh['stop-stuck'] -match '^returned False \|' -and $sh['stop-stuck'] -match 'KILL cowork-svc 301' -and
             $sh['stop-stuck'] -match 'is still running after a force-kill attempt') $sh['stop-stuck']
        Assert-Case "Stop-CoworkServiceWithTimeout, stopped by the force-kill: true" `
            ($sh['stop-killed'] -match '^returned True \|' -and $sh['stop-killed'] -match 'KILL cowork-svc 301') $sh['stop-killed']
        Assert-Case "Stop-CoworkServiceWithTimeout, after the force-kill the service could not be read: false, and said" `
            ($sh['stop-afterUnread'] -match '^returned False \|' -and
             $sh['stop-afterUnread'] -match 'LOG: \[!\] Whether CoworkVMService stopped after the force-kill could not be read \(Access is denied\)') $sh['stop-afterUnread']
        Assert-Case "Wait-ForServiceProcessExit, the list could not be read: false, and said" `
            ($sh['wait-denied'] -eq 'returned False | LOG: [!] Whether the service process has exited could not be read (Access is denied); the VHDX files may be locked') $sh['wait-denied']
        Assert-Case "Wait-ForServiceProcessExit, no process: true" ($sh['wait-none'] -eq 'returned True | ') $sh['wait-none']
        Assert-Case "Wait-ForServiceProcessExit, a process that stays: false" `
            ($sh['wait-stays'] -eq 'returned False | LOG: Service process still running after 2s; the VHDX files may be locked') $sh['wait-stays']
        Assert-Case "Test-CoworkServicePrereq, vmcompute could not be read: blocked, as not known" `
            ($sh['prereq-hcsUnread'] -eq 'True|False|Whether the Host Compute Service stack is registered could not be read: vmcompute (Access is denied).') $sh['prereq-hcsUnread']
        Assert-Case "Test-CoworkServicePrereq, the service could not be read: blocked, not 'not registered'" `
            ($sh['prereq-svcUnread'] -eq 'True|False|Whether CoworkVMService is registered could not be read (Access is denied).') $sh['prereq-svcUnread']
        Assert-Case "Test-CoworkServicePrereq, no service: not registered" ($sh['prereq-svcAbsent'] -match '^True\|False\|CoworkVMService is not registered\.') $sh['prereq-svcAbsent']
        Assert-Case "Test-CoworkServicePrereq, all there: not blocked" ($sh['prereq-ok'] -eq 'False|True|') $sh['prereq-ok']
    }
}

# --- Test-CoworkServiceLog and Close-StaleHcsVms -----------------------------
# A service log that could not be read is said once a run; until 6.3.1 it
# was not said. A kill that could not be sent, and an HCS clean-up that
# stopped on an error, are said; until 6.3.1 both were silent, and the 0
# returned read as nothing to kill.
$logDefs = @(@('Test-CoworkServiceLog', 'Close-StaleHcsVms') | ForEach-Object { Get-AstFunction -Ast $fixAst -Name $_ } | Where-Object { $_ })
$unfaked = @(if ($logDefs.Count -eq 2) {
        Get-FrUnsafe -Defs $logDefs -Also @('Get-Content', 'New-Object', 'Invoke-HcsDiag', 'Get-CoworkHcsGuids', 'Get-CoworkHcsSurvivors')
    } else { '<not found>' })
Assert-Case "Test-CoworkServiceLog and Close-StaleHcsVms were found, and every command they call is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $lg = Invoke-Isolated -Name 'Test-CoworkServiceLog and Close-StaleHcsVms' -Block {
        . $frSetup
        function Get-Content {
            [CmdletBinding()]
            param([Parameter(Position = 0)][string]$Path, [string]$LiteralPath, [int]$Tail)
            $null = $Tail
            $p = if ($LiteralPath) { $LiteralPath } else { $Path }
            $PSCmdlet.WriteError((New-FrRecord 'System.IO.IOException' "The process cannot access the file '$p' because it is being used by another process." `
                'GetContentReaderIOError' 'ReadError' $p))
        }
        function New-Object { param([string]$TypeName, [object[]]$ArgumentList) $null = $ArgumentList; throw "The process cannot access the file because it is being used by another process ($TypeName)." }
        function Invoke-HcsDiag { param([string[]]$Arguments) $null = $Arguments; throw 'The job could not be started' }
        foreach ($d in $logDefs) { . ([scriptblock]::Create($d.Extent.Text)) }
        $script:ClaudeEnv = [pscustomobject]@{ LogDir = 'C:\nx\logs' }
        $o = [ordered]@{}
        foreach ($c in 'denied', 'locked') {
            Reset-Fr
            $script:SvcLogUnreadSaid = $false
            $script:FrItems['C:\nx\logs\cowork-service.log'] = if ($c -eq 'denied') { 'Denied' } else { [pscustomobject]@{ FullName = 'C:\nx\logs\cowork-service.log' } }
            $first = Test-CoworkServiceLog
            $second = Test-CoworkServiceLog
            $o["log-$c"] = "$($null -eq $first)|$($null -eq $second) | $(Get-FrSaid)"
        }
        Reset-Fr
        $script:FrHcs = "cowork-vm-1699151a`r`n    VM, Running, DE1517EC-0000-4000-8000-77A48CB1AD97, cowork-vm-1699151a"
        $o['hcs-send'] = "returned $(Close-StaleHcsVms -Action kill) | $(Get-FrSaid)"
        Reset-Fr
        $script:FrHcs = 'THROWS'
        function Get-HcsList { throw 'hcsdiag did not answer' }
        $o['hcs-stop'] = "returned $(Close-StaleHcsVms -Action kill) | $(Get-FrSaid)"
        $o
    }
    if ($lg) {
        Assert-Case "Test-CoworkServiceLog, a log that could not be read: no state, said once a run" `
            ($lg['log-denied'] -eq "True|True | LOG: [!] C:\nx\logs\cowork-service.log could not be read (Access is denied), so the guest's state in it is not known") `
            $lg['log-denied']
        Assert-Case "Test-CoworkServiceLog, a log that is there but locked: no state, said once a run" `
            ($lg['log-locked'] -match '^True\|True \| LOG: \[!\] C:\\nx\\logs\\cowork-service\.log could not be read \(The process cannot access the file' -and
             @($lg['log-locked'] -split 'LOG: ').Count -eq 2) $lg['log-locked']
        Assert-Case "Close-StaleHcsVms, a kill that could not be sent: not counted, and said" `
            ($lg['hcs-send'] -match '^returned 0 \| LOG: \[!\] The kill of compute system \S+ could not be sent \(The job could not be started\)') $lg['hcs-send']
        Assert-Case "Close-StaleHcsVms, an error that stops the clean-up: said" `
            ($lg['hcs-stop'] -eq 'returned 0 | LOG: [!] The HCS clean-up stopped on an error (hcsdiag did not answer), so what it removed is not known') $lg['hcs-stop']
    }
}

# --- the final summary -------------------------------------------------------
# A service that could not be read says so. Until 6.3.1 it said "Not found".
$fsumText = Get-FixSegment -From '$fSvc' -To '# If the service can never start, say so here'
$unfaked = @(Get-FrUnsafe -Text $fsumText -Also @('Get-ClaudeDesktopProcess'))
Assert-Case "the final summary's service line was found, and every command it calls is faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $fs = Invoke-Isolated -Name "the final summary's service line" -Block {
        . $frSetup
        function Get-ClaudeDesktopProcess { param($Environment) $null = $Environment; return @() }
        $seg = [scriptblock]::Create($fsumText)
        $script:ClaudeEnv = [pscustomobject]@{ LogDir = 'C:\nx\logs' }
        $o = [ordered]@{}
        foreach ($c in 'Denied', 'Running', 'Absent') {
            Reset-Fr
            if ($c -ne 'Absent') { $script:FrSvc['CoworkVMService'] = $c }
            & { . $seg }
            $o[$c] = Get-FrSaid
        }
        $o
    }
    if ($fs) {
        Assert-Case "the summary, the service could not be read: said, not 'Not found'" `
            ($fs['Denied'] -match '\[Yellow\]\s+VM Service:\s+could not be read \(Access is denied\)' -and $fs['Denied'] -notmatch 'Not found') $fs['Denied']
        Assert-Case "the summary, the service running: Running, in green" ($fs['Running'] -match '\[Green\]\s+VM Service:\s+Running') $fs['Running']
        Assert-Case "the summary, no service: Not found" ($fs['Absent'] -match '\[Yellow\]\s+VM Service:\s+Not found') $fs['Absent']
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- the monitor: the old watch logs, and the activity checks' type"

# Remove-OldWatchLog deletes the watch logs older than 30 days, at startup
# and at each rollover. Until 6.3.1 the list and the deletes ran with
# SilentlyContinue, and a folder that could not be listed, or a log that
# could not be deleted, was left without a word. It is lifted and run on a
# temp folder, with Get-ChildItem and Remove-Item handing on to the real
# ones; a path outside that folder is a tripwire. A log held open makes a
# real delete fail. The listing that is refused is faked, with the error id
# and category a folder that cannot be read gives (DirUnauthorizedAccessError,
# PermissionDenied).
$oldLogDef = Get-AstFunction -Ast $watchAst -Name 'Remove-OldWatchLog'
$rollDef = Get-AstFunction -Ast $watchAst -Name 'Update-LogFile'
$oldLogRegion = @(Get-AstFunction -Ast $watchAst -Name 'Test-ClaudeNotFoundError' | Where-Object { $_ })
$oldLogSayNames = @('Get-CheckErrorText', 'Test-CheckErrorDue', 'Write-CheckError')
$oldLogSay = @($oldLogSayNames | ForEach-Object { Get-AstFunction -Ast $watchAst -Name $_ } | Where-Object { $_ })
$unfaked = @(if ($oldLogDef -and $rollDef -and $oldLogRegion.Count -eq 1 -and $oldLogSay.Count -eq 3) {
        Get-UnfakedCommand -Asts (@($oldLogDef) + $oldLogRegion) -Faked @('Get-ChildItem', 'Remove-Item', 'Test-ClaudeNotFoundError')
        Get-UnfakedCommand -Asts (@($rollDef) + $oldLogSay) -Faked (@('Remove-OldWatchLog', 'Write-WatchLog') + $oldLogSayNames)
    } else { '<not found>' })
Assert-Case "Remove-OldWatchLog and Update-LogFile were found, and every command they call is faked or reads a temp folder" `
    ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $oldLogs = Invoke-Isolated -Name 'the old watch logs' -Block {
        Set-StrictMode -Version Latest
        $logRoot = Join-Path $tempRoot 'watch-logs'
        $script:LogTrips = [System.Collections.ArrayList]::new()
        $script:Deletes = [System.Collections.ArrayList]::new()
        $script:ListFault = $null
        $script:GoneFirst = @()
        function Test-InLogRoot { param([string]$Path) return $Path.StartsWith($logRoot + '\', [StringComparison]::OrdinalIgnoreCase) }
        function Get-ChildItem {
            [CmdletBinding()]
            param([string]$LiteralPath, [string]$Filter, [switch]$File)
            if (-not (Test-InLogRoot $LiteralPath)) {
                [void]$script:LogTrips.Add("list $LiteralPath")
                throw "tripwire: a listing outside the temp folder: $LiteralPath"
                return
            }
            if ($script:ListFault -eq 'denied') {
                $ex = New-Object System.UnauthorizedAccessException "Access to the path '$LiteralPath' is denied."
                $PSCmdlet.WriteError((New-Object System.Management.Automation.ErrorRecord($ex, 'DirUnauthorizedAccessError',
                    [System.Management.Automation.ErrorCategory]::PermissionDenied, $LiteralPath)))
                return
            }
            Microsoft.PowerShell.Management\Get-ChildItem @PSBoundParameters
        }
        # $script:GoneFirst: logs deleted by someone else just before this
        # delete, as a second monitor at its own rollover would.
        function Remove-Item {
            [CmdletBinding()]
            param([string]$LiteralPath, [switch]$Force)
            if (-not (Test-InLogRoot $LiteralPath)) {
                [void]$script:LogTrips.Add("delete $LiteralPath")
                throw "tripwire: a delete outside the temp folder: $LiteralPath"
                return
            }
            $leaf = [System.IO.Path]::GetFileName($LiteralPath)
            [void]$script:Deletes.Add($leaf)
            if ($script:GoneFirst -contains $leaf) { [System.IO.File]::Delete($LiteralPath) }
            Microsoft.PowerShell.Management\Remove-Item @PSBoundParameters
        }
        # A folder with the logs named in -Ages, each last written that many
        # days ago, and a folder named like a log, 40 days old, which is not
        # a log.
        function New-LogDir {
            param([string]$Name, [hashtable]$Ages)
            $d = Join-Path $logRoot $Name
            $null = [System.IO.Directory]::CreateDirectory($d)
            foreach ($k in $Ages.Keys) {
                $p = Join-Path $d $k
                [System.IO.File]::WriteAllText($p, 'a line')
                [System.IO.File]::SetLastWriteTimeUtc($p, [datetime]::UtcNow.AddDays(-$Ages[$k]))
            }
            $sub = Join-Path $d 'watch_folder.log'
            $null = [System.IO.Directory]::CreateDirectory($sub)
            [System.IO.Directory]::SetLastWriteTimeUtc($sub, [datetime]::UtcNow.AddDays(-40))
            return $d
        }
        function Get-LogLeft {
            param([string]$Dir)
            return (@([System.IO.Directory]::GetFileSystemEntries($Dir) | ForEach-Object { [System.IO.Path]::GetFileName($_) } |
                        Sort-Object) -join ',')
        }
        function Get-ErrorSaid {
            param($Result)
            if ($null -eq $Result) { return 'none' }
            if ($Result -isnot [System.Management.Automation.ErrorRecord]) { return "not an error: $Result" }
            return "$(@("$($Result.FullyQualifiedErrorId)" -split ',')[0]): $($Result.Exception.Message)"
        }
        foreach ($d in $oldLogRegion) { . ([scriptblock]::Create($d.Extent.Text)) }
        . ([scriptblock]::Create($oldLogDef.Extent.Text))
        $ages = @{ 'watch_20261001.log' = 40; 'watch_20261020.log' = 31; 'watch_20261121.log' = 1; 'other.log' = 40 }
        $o = [ordered]@{}

        $WatchLogDir = New-LogDir -Name 'plain' -Ages $ages
        $script:Deletes.Clear()
        $r = Remove-OldWatchLog
        $o['plain'] = "$(Get-ErrorSaid $r)|$(Get-LogLeft $WatchLogDir)|$($script:Deletes -join ',')"

        $WatchLogDir = Join-Path $logRoot 'none'
        $script:Deletes.Clear()
        $r = Remove-OldWatchLog
        $o['missing'] = "$(Get-ErrorSaid $r)|$($script:Deletes.Count)"

        $WatchLogDir = New-LogDir -Name 'denied' -Ages $ages
        $script:Deletes.Clear()
        $script:ListFault = 'denied'
        $r = Remove-OldWatchLog
        $script:ListFault = $null
        $o['denied'] = "$(Get-ErrorSaid $r)|$(Get-LogLeft $WatchLogDir)|$($script:Deletes.Count)"
        $o['deniedDir'] = $WatchLogDir

        $WatchLogDir = New-LogDir -Name 'locked' -Ages $ages
        $script:Deletes.Clear()
        $lock = [System.IO.File]::Open((Join-Path $WatchLogDir 'watch_20261001.log'), 'Open', 'Read', 'None')
        try { $r = Remove-OldWatchLog } finally { $lock.Dispose() }
        $o['locked'] = "$(@((Get-ErrorSaid $r) -split ':')[0])|$(Get-LogLeft $WatchLogDir)|$($script:Deletes -join ',')"

        $WatchLogDir = New-LogDir -Name 'gone' -Ages $ages
        $script:Deletes.Clear()
        $script:GoneFirst = @('watch_20261001.log')
        $r = Remove-OldWatchLog
        $script:GoneFirst = @()
        $o['gone'] = "$(Get-ErrorSaid $r)|$(Get-LogLeft $WatchLogDir)|$($script:Deletes -join ',')"

        $o['trips'] = ($script:LogTrips -join ', ')
        $o
    }
    if ($oldLogs) {
        $kept = 'other.log,watch_20261121.log,watch_folder.log'
        Assert-Case "the old watch logs: those over 30 days old are deleted; a newer one, another file and a folder are kept" `
            ($oldLogs['plain'] -eq "none|$kept|watch_20261001.log,watch_20261020.log") $oldLogs['plain']
        Assert-Case "the old watch logs: a folder that is not there is no error, and nothing is deleted" ($oldLogs['missing'] -eq 'none|0') $oldLogs['missing']
        Assert-Case "the old watch logs: a folder that cannot be listed is the error returned, and nothing is deleted" `
            ($oldLogs['denied'] -eq "DirUnauthorizedAccessError: Access to the path '$($oldLogs['deniedDir'])' is denied.|other.log,watch_20261001.log,watch_20261020.log,watch_20261121.log,watch_folder.log|0") `
            $oldLogs['denied']
        Assert-Case "the old watch logs: a log that cannot be deleted is the error returned, and the next old one is still deleted" `
            ($oldLogs['locked'] -eq "RemoveFileSystemItemIOError|other.log,watch_20261001.log,watch_20261121.log,watch_folder.log|watch_20261001.log,watch_20261020.log") `
            $oldLogs['locked']
        Assert-Case "the old watch logs: a log already gone is no error, and the next old one is still deleted" `
            ($oldLogs['gone'] -eq "none|$kept|watch_20261001.log,watch_20261020.log") $oldLogs['gone']
        Assert-Case "the old watch logs: nothing outside the temp folder was listed or deleted" ($oldLogs['trips'] -eq '') $oldLogs['trips']
    }

    # Update-LogFile at the day's rollover: the error Remove-OldWatchLog
    # returns is said through the hourly throttle, naming the folder and the
    # error. Remove-OldWatchLog is faked here, and the throttle is lifted.
    $roll = Invoke-Isolated -Name 'the rollover' -Block {
        Set-StrictMode -Version Latest
        $WatchLogDir = 'C:\nx\watch-logs'
        $script:WatchLines = [System.Collections.ArrayList]::new()
        $script:CheckErrorsSaid = @{}
        $script:CleanupCalls = 0
        $script:RollError = $null
        function Write-WatchLog { param([string]$Message) [void]$script:WatchLines.Add($Message) }
        function Remove-OldWatchLog { $script:CleanupCalls++; return $script:RollError }
        foreach ($d in $oldLogSay) { . ([scriptblock]::Create($d.Extent.Text)) }
        . ([scriptblock]::Create($rollDef.Extent.Text))
        $ex = New-Object System.UnauthorizedAccessException "Access to the path '$WatchLogDir' is denied."
        $script:RollError = New-Object System.Management.Automation.ErrorRecord($ex, 'DirUnauthorizedAccessError',
            [System.Management.Automation.ErrorCategory]::PermissionDenied, $WatchLogDir)
        $o = [ordered]@{}
        $script:WatchLogFile = 'C:\nx\watch-logs\watch_19990101.log'
        Update-LogFile
        $script:WatchLogFile = 'C:\nx\watch-logs\watch_19990102.log'
        Update-LogFile
        $o['denied'] = "$($script:CleanupCalls)|$($script:WatchLines -join ' | ')"
        $script:WatchLines.Clear()
        $script:CheckErrorsSaid = @{}
        $script:RollError = $null
        $script:WatchLogFile = 'C:\nx\watch-logs\watch_19990103.log'
        Update-LogFile
        $o['clean'] = "$($script:CleanupCalls)|$($script:WatchLines.Count)"
        Update-LogFile
        $o['sameDay'] = "$($script:CleanupCalls)"
        $o
    }
    if ($roll) {
        $rollLine = "The watch logs older than 30 days in C:\nx\watch-logs could not all be deleted (Access to the path 'C:\nx\watch-logs' is denied.); " +
            'they are tried again at the next rollover; the same error is logged at most once an hour'
        Assert-Case "the rollover: old logs that could not be deleted are said, naming the folder and the error, once in the hour" `
            ($roll['denied'] -eq "2|$rollLine") $roll['denied']
        Assert-Case "the rollover: old logs all deleted, nothing said; and no clean-up again the same day" `
            ($roll['clean'] -eq '3|0' -and $roll['sameDay'] -eq '3') "$($roll['clean']); $($roll['sameDay'])"
    }
}

# The startup clean-up and the activity checks' type are tried before there
# is a log, so their errors are kept and said after the banner. The two
# statements that say them are lifted from the top level of Watch and run.
$bannerIfs = @($watchAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.IfStatementAst] -and
            @('$script:Win32ActivityError', '$script:LogCleanupError') -contains $n.Clauses[0].Item1.Extent.Text }, $false))
Assert-Case "Watch keeps the startup clean-up's error and the type's error, and says each after the banner" `
    ($bannerIfs.Count -eq 2 -and $watchSrc -match '(?m)^\$script:LogCleanupError = Remove-OldWatchLog\s*$' -and
    $watchSrc -match '(?m)^\$script:Win32ActivityError = Import-Win32Activity\s*$') "$($bannerIfs.Count) statement(s)"
$unfaked = @(Get-UnfakedCommand -Asts (@($bannerIfs) + @($oldLogSay)) -Faked (@('Write-WatchLog') + $oldLogSayNames))
if ($bannerIfs.Count -eq 2 -and $oldLogSay.Count -eq 3 -and $unfaked.Count -eq 0) {
    $banner = Invoke-Isolated -Name "the banner's lines" -Block {
        Set-StrictMode -Version Latest
        $WatchLogDir = 'C:\nx\watch-logs'
        $script:WatchLines = [System.Collections.ArrayList]::new()
        $script:CheckErrorsSaid = @{}
        function Write-WatchLog { param([string]$Message) [void]$script:WatchLines.Add($Message) }
        foreach ($d in $oldLogSay) { . ([scriptblock]::Create($d.Extent.Text)) }
        $ex = New-Object System.Exception "c:\Temp\x\x.0.cs(1) : } expected`r`n`r`nc:\Temp\x\x.0.cs(1) : >>> public static class Win32Activity {"
        $script:Win32ActivityError = New-Object System.Management.Automation.ErrorRecord($ex, 'SOURCE_CODE_ERROR',
            [System.Management.Automation.ErrorCategory]::InvalidData, $null)
        $ex = New-Object System.IO.IOException "The process cannot access the file 'C:\nx\watch-logs\watch_19990101.log' because it is being used by another process."
        $script:LogCleanupError = New-Object System.Management.Automation.ErrorRecord($ex, 'RemoveFileSystemItemIOError',
            [System.Management.Automation.ErrorCategory]::WriteError, $null)
        foreach ($i in $bannerIfs) { . ([scriptblock]::Create($i.Extent.Text)) }
        ($script:WatchLines -join ' | ')
    }
    $wantBanner = 'The foreground-window and input checks could not be loaded (c:\Temp\x\x.0.cs(1) : } expected c:\Temp\x\x.0.cs(1) : >>> ' +
        'public static class Win32Activity {), so they are skipped, as in Session 0; the same error is logged at most once an hour | ' +
        "The watch logs older than 30 days in C:\nx\watch-logs could not all be deleted (The process cannot access the file " +
        "'C:\nx\watch-logs\watch_19990101.log' because it is being used by another process.); they are tried again at the next rollover; " +
        'the same error is logged at most once an hour'
    Assert-Case "after the banner: the type that did not load and the old logs not deleted are each said, naming the error" `
        ($banner -eq $wantBanner) $banner
}

# Import-Win32Activity, which loads the type the foreground-window and input
# checks call. Until 6.3.1 its Add-Type ran with SilentlyContinue and why the
# checks were off was never said. It runs in a child PowerShell, which has
# no Win32Activity type yet, with Add-Type faked: the compile error and the
# type that already exists carry the ids and categories measured 2026-11-22
# (SOURCE_CODE_ERROR, InvalidData; TYPE_ALREADY_EXISTS, InvalidOperation),
# and "stub" hands the real Add-Type a type with no native calls in it.
$win32Def = Get-AstFunction -Ast $watchAst -Name 'Import-Win32Activity'
$unfaked = @(if ($win32Def) { Get-UnfakedCommand -Asts @($win32Def) -Faked @('Add-Type') } else { '<not found>' })
Assert-Case "Import-Win32Activity was found, and the only command it calls is Add-Type, faked" ($unfaked.Count -eq 0) ($unfaked -join ', ')
if ($unfaked.Count -eq 0) {
    $importFakes = @'
Set-StrictMode -Version Latest
$script:Win32ActivityLoaded = $false
$script:AddTypeMode = 'compile'
function Add-Type {
    [CmdletBinding()]
    param([string]$TypeDefinition)
    $null = $TypeDefinition
    switch ($script:AddTypeMode) {
        'compile' {
            $ex = New-Object System.Exception 'c:\Temp\x\x.0.cs(1) : } expected'
            $PSCmdlet.WriteError((New-Object System.Management.Automation.ErrorRecord($ex, 'SOURCE_CODE_ERROR',
                [System.Management.Automation.ErrorCategory]::InvalidData, $null)))
        }
        'exists' {
            $ex = New-Object System.Exception "Cannot add type. The type name 'Win32Activity' already exists."
            $PSCmdlet.WriteError((New-Object System.Management.Automation.ErrorRecord($ex, 'TYPE_ALREADY_EXISTS',
                [System.Management.Automation.ErrorCategory]::InvalidOperation, $null)))
        }
        'stub' { Microsoft.PowerShell.Utility\Add-Type -TypeDefinition 'public static class Win32Activity { public static int Stub() { return 1; } }' }
        default { }
    }
}
'@
    $importCases = @'
foreach ($m in 'compile', 'quiet', 'stub', 'exists') {
    $script:AddTypeMode = $m
    $r = Import-Win32Activity
    $id = if ($null -eq $r) { 'none' } elseif ($r -is [System.Management.Automation.ErrorRecord]) { @("$($r.FullyQualifiedErrorId)" -split ',')[0] } else { "other: $r" }
    Write-Output "RESULT $m loaded=$($script:Win32ActivityLoaded) error=$id"
}
'@
    $child = Invoke-Child -Body ($importFakes + "`r`n" + $win32Def.Extent.Text + "`r`n" + $importCases) -Name 'import-win32activity'
    $got = @(($child.Out -split "`r?`n") | Where-Object { $_ -like 'RESULT *' })
    $saidAll = "exit $($child.Exit); $($got -join ' / ')"
    Assert-Case "the type, a compile error: not loaded, and the error returned" ($got -contains 'RESULT compile loaded=False error=SOURCE_CODE_ERROR') $saidAll
    Assert-Case "the type, an Add-Type that says nothing and adds nothing: not loaded, and the missing type returned" `
        ($got -contains 'RESULT quiet loaded=False error=TypeNotFound') $saidAll
    Assert-Case "the type, added: loaded, and nothing returned" ($got -contains 'RESULT stub loaded=True error=none') $saidAll
    Assert-Case "the type, already there and Add-Type refuses it: loaded, and nothing returned" `
        ($child.Exit -eq 0 -and $got.Count -eq 4 -and $got -contains 'RESULT exists loaded=True error=none') $saidAll
}

} catch {
    # An error that escapes a section skips every section after it. Until
    # 6.3.0 nothing caught it here, and the gate passed on the cases that had
    # run before it.
    Assert-Case "every section ran to its end" $false "$($_.Exception.Message) (line $($_.InvocationInfo.ScriptLineNumber))"
} finally {
    # Clean-up of this script's own temp folder: every case is counted by
    # now, and a folder left in %TEMP% changes nothing.
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host ("Run outcomes: {0} passed, {1} failed" -f $pass, $fail)
if ($fail -gt 0) {
    Write-Host "Run outcomes gate FAILED" -ForegroundColor Red
    exit 1
}
Write-Host "Run outcomes gate PASSED" -ForegroundColor Green
exit 0
