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
$wantAsked = @('Step 1=$script:BlockedExitCode', 'Step 4=$script:StoppedPartWayExitCode',
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
# writes; Test-Path, Get-Content and ConvertFrom-Json are the real ones.
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
        'Get-HeldBackRetryWait', 'Get-CheckErrorText', 'Test-CheckErrorDue', 'Write-CheckError') |
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
        'Get-CimInstance', 'Get-ClaudeDesktopProcess', 'Get-ClaudeVmLogAge', 'Test-UserActivity',
        'Test-ClaudeCodeActive', 'Test-CoworkSessionActive', 'Test-PersistentFailure',
        'ConvertFrom-WatchStateTime', 'Test-Path', 'Get-Content', 'ConvertFrom-Json', 'Clear-HeldBackRepair',
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
            function Get-Process { return @() }
            function Get-CimInstance { return @() }
            function Get-ClaudeDesktopProcess { param($Environment) $null = $Environment; return @() }
            function Get-ClaudeVmLogAge { param($LogDir) $null = $LogDir; return $null }
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
            function Test-ClaudeCodeTranscriptRecent { param([int]$Minutes) $null = $Minutes; return [bool]$script:TranscriptRecent }
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
            # A file that is not JSON: what it held starts empty, and that is
            # said, once in the hour for the same error. Until 6.3.0 it was
            # not said.
            $script:WatchLines.Clear()
            $script:CheckErrorsSaid = @{}
            $null = Read-State '{ "LastFixTime": '
            $null = Read-State '{ "LastFixTime": '
            $o['stateBroken'] = "$(@($script:WatchLines | Where-Object { $_ -match '^STATE: The state saved in .+watch-state\.json could not be restored in full \(.+\); what was not restored starts empty, as on a first start; the same error is logged at most once an hour$' }).Count)|$($script:WatchLines.Count)"
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
            $o['statePending'] = "$(@($script:WatchLines | Where-Object { $_ -match '^STATE: The held-back repair saved in .+watch-state\.json could not be restored \(.+\), so it is not tried again; the same error is logged at most once an hour$' }).Count)|$($script:WatchLines.Count)|$([string]$script:PendingFixReason)"
            # And the cooldown from such a file: a fix 2 minutes ago, written
            # in local time an hour ahead of UTC, is still 2 minutes ago.
            $script:WatchLines.Clear()
            $script:RepeatLinesSaid = @{}
            $twoAgo = [datetimeoffset]::UtcNow.AddMinutes(-2).ToOffset([timespan]::FromHours(1)).ToString('o', [System.Globalization.CultureInfo]::InvariantCulture)
            $null = Read-State ('{ "LastFixTime": "' + $twoAgo + '", "FixCount": 1, "FixHistory": [ "' + $twoAgo + '" ] }')
            $null = Invoke-AutoFix -Reason 'test trigger'
            $o['stateCooldown'] = ($script:WatchLines -join ' | ')

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
                              @{ Name = 'codeCliThrow'; Cli = 'THROW'; Transcript = $false })) {
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
            Assert-Case "a file that is not JSON: said, naming the file and the error, once in the hour" ($auto['stateBroken'] -eq '1|1') `
                "that line|lines: $($auto['stateBroken'])"
            Assert-Case "a held-back repair whose time cannot be read: not restored, and said, naming the file and the error, once in the hour" `
                ($auto['statePending'] -eq '1|1|') "that line|lines|pending: $($auto['statePending'])"
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
                ($auto['codeCliThrow'] -eq '>>> BLOCKED: test trigger, the Claude Code sessions the app started could not be listed|0|0') "BLOCKED|warnings and runs|history: $($auto['codeCliThrow'])"
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
$clockFakes = @('Write-WatchLog', 'Get-Service', 'Start-Service', 'Invoke-W32tm', 'Get-W32tmOffsetSeconds',
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
$codeFakes = @('Get-ClaudeAppCliProcess', 'Get-ClaudeDataCandidates', 'Test-Path', 'Get-ChildItem',
    'Test-ClaudeCodeTranscriptRecent', 'Write-WatchLog')
$unfaked = @(if ($codeDef) { Get-UnfakedCommand -Asts @($codeDef) -Faked $codeFakes } else { '<not found>' })
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
            return [pscustomobject]@{ DataRoots = @([pscustomobject]@{ Path = 'C:\fake-root' }); LogDirs = @() }
        }
        function Test-Path { return $true }
        function Get-ChildItem { return $script:DstFile }
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
# it did.
$closeIf = Find-FixStatement { param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and
    $n.Clauses[0].Item1.Extent.Text -eq '$remainingProcs.Count -eq 0 -and -not $remainingHcs' }
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
                             @{ Name = 'knownUp'; Unknown = $false; Up = $true })) {
                $script:CloseLines = [System.Collections.ArrayList]::new()
                $remainingProcs = @()
                $remainingHcs   = $false
                $hcsUnknown     = $c.Unknown
                $svcRunning     = $c.Up
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
                @{ Name = 'procsUnknown'; Procs = 2; Hcs = $false; Unknown = $true;  Up = $true })) {
                $script:CloseLines = [System.Collections.ArrayList]::new()
                $remainingProcs = @()
                for ($k = 0; $k -lt $c.Procs; $k++) { $remainingProcs += $k }
                $remainingHcs   = $c.Hcs
                $hcsUnknown     = $c.Unknown
                $svcRunning     = $c.Up
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
    'Close-StaleHcsVms', 'Stop-CoworkServiceWithTimeout', 'Start-Service', 'Start-ClaudeDesktop', 'Get-VmLogSizes')
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
            $ServiceName = 'CoworkVMService'
            $script:StoppedPartWayExitCode = 4
            $listed = "cowork-vm-1699151a`r`n    VM, Running, DE1517EC-0000-4000-8000-77A48CB1AD97, cowork-vm-1699151a"
            $quiet  = @{ Log = $null;             Guest = 'no-polling';    Hcs = $listed }
            $busy   = @{ Log = 'vsock-connected'; Guest = 'no-polling';    Hcs = $listed }
            $guest  = @{ Log = $null;             Guest = 'guest-polling'; Hcs = $listed }
            $noVm   = @{ Log = $null;             Guest = 'no-polling';    Hcs = '' }
            $noHcs  = @{ Log = $null;             Guest = 'no-polling';    Hcs = 'THROWS' }
            $o = [ordered]@{}
            foreach ($c in @(
                @{ Name = 'progressing';  Escalated = $true;  Passes = @($busy, $busy, $busy, $busy) },
                @{ Name = 'guest';        Escalated = $true;  Passes = @($guest, $guest, $guest, $guest) },
                @{ Name = 'stuck';        Escalated = $true;  Passes = @($quiet, $quiet, $quiet, $quiet) },
                @{ Name = 'stuckFirst';   Escalated = $false; Passes = @($quiet, $quiet, $quiet, $quiet) },
                @{ Name = 'twoStretches'; Escalated = $true;  Passes = @($quiet, $quiet, $busy, $quiet, $quiet) },
                @{ Name = 'noVm';         Escalated = $true;  Passes = @($noVm, $noVm, $noVm) },
                @{ Name = 'noHcs';        Escalated = $true;  Passes = @($noHcs, $noHcs, $noHcs) })) {
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
$keepFaked = @('Get-HcsListCached', 'Get-CoworkHcsGuids', 'Get-CimInstance', 'Invoke-HcsDiag', 'Get-CoworkHcsSurvivors',
    'Write-WatchLog', 'Write-CheckError', 'Get-VmWorkerGuid')
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
    'Get-ClaudeCodeSessionHold', 'Test-LineDue', 'Get-CheckErrorText', 'Test-CheckErrorDue', 'Write-CheckError')
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
    # Test-Path, Get-Content, ConvertFrom-Json and ConvertTo-Json are the real
    # ones, on the temp file.
    $retryFakes = @('Write-WatchLog', 'Add-Type', 'New-Object', 'Start-Sleep', 'Get-Process', 'Get-CimInstance',
        'Get-ClaudeDesktopProcess', 'Get-ClaudeVmLogAge', 'Test-UserActivity', 'Test-ClaudeCodeActive',
        'Test-CoworkSessionActive', 'Get-ClaudeAppCliProcess', 'Get-ChildItem', 'Test-Path', 'Get-Content',
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
    ($retry['waitCliThrow'] -eq 'a held-back repair for test trigger is still waiting: the Claude Code sessions the app started could not be listed') "got $($retry['waitCliThrow'])"
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
    'Test-CoworkSessionActive', 'Write-WatchLog')
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
        function Get-ClaudeDataCandidates {
            param($PackageFamilyName)
            $null = $PackageFamilyName
            return [pscustomobject]@{ DataRoots = @([pscustomobject]@{ Path = 'C:\fake-root' }); LogDirs = @() }
        }
        # No claude-code-sessions folder anywhere.
        function Test-Path { return $false }
        function Get-ChildItem { return $null }
        function Test-ClaudeCodeTranscriptRecent { param([int]$Minutes) $null = $Minutes; return $false }
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
        # One process, idle; it has no Refresh, so the CPU check reads no
        # change.
        function Get-Process { return @([pscustomobject]@{ Id = 4242; TotalProcessorTime = [timespan]::Zero }) }
        function Get-CimInstance { return $null }
        function New-Object { throw 'no desktop here' }
        function Get-ClaudeVmLogAge {
            param($LogDir)
            $null = $LogDir
            if ($script:VmThrows) { throw 'the VM logs could not be read' }
            return $null
        }
        function Start-Sleep { param([int]$Seconds, [int]$Milliseconds) $null = $Seconds, $Milliseconds }
        function Test-ClaudeCodeActive { param([switch]$Silent) $script:CodeAskedSilent = [bool]$Silent; return $true }
        function Test-CoworkSessionActive { return $false }
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

# Get-HeldBackRetryWait, Get-ClaudeAppStartTime and Get-FixRunAfter. The
# last reads a temp folder with the real Get-ChildItem, its files' times set
# by hand: only fix_*.log files that hold the line Fix logs once it begins
# the repair count, by the later of when each was created and when it was
# last written. Test-FixLogRepaired, which reads that line, is lifted when
# it is there: until 6.3.0 it was not, and any Fix log counted.
$helperNames = @('Get-HeldBackRetryWait', 'Get-ClaudeAppStartTime', 'Get-FixRunAfter')
$helperDefs = @($helperNames | ForEach-Object { Get-AstFunction -Ast $watchAst -Name $_ } | Where-Object { $_ })
$repairedDefs = @(Get-AstFunction -Ast $watchAst -Name 'Test-FixLogRepaired' | Where-Object { $_ })
$unfaked = @(if ($helperDefs.Count -eq $helperNames.Count) {
        Get-UnfakedCommand -Asts (@($helperDefs) + @($repairedDefs)) -Faked @('Get-ClaudeDesktopProcess', 'Get-ChildItem', 'Test-FixLogRepaired',
            'Write-CheckError')
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
        # No case here reads a folder that is there and cannot be read.
        function Write-CheckError { param([string]$Check, [string]$Text, $ErrorRecord) $null = $Check, $ErrorRecord; throw "unexpected: $Text" }
        foreach ($d in (@($helperDefs) + @($repairedDefs))) { . ([scriptblock]::Create($d.Extent.Text)) }
        $FixRepairLine = if ($watchRepairMatch.Success) { $watchRepairMatch.Groups[1].Value } else { 'Repair under way' }
        $o = [ordered]@{}
        $o['waits'] = (@(0..7 | ForEach-Object { Get-HeldBackRetryWait -RetriesHeld $_ }) -join ',')

        $nowUtc = [datetime]::UtcNow
        $script:ListThrows = $false
        $script:Procs = @()
        $none = Get-ClaudeAppStartTime
        $exited = [pscustomobject]@{ Id = 3 }
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
        $o
    }
}
if ($null -eq $helpers) { $helpers = [ordered]@{} }
Assert-Case "the wait after a hold: 2 minutes, doubled for each retry held back, 60 at most" ($helpers['waits'] -eq '2,4,8,16,32,60,60,60') "got $($helpers['waits'])"
Assert-Case "the app's start: none running is none; the oldest process, in UTC, past one that has exited; a list that fails is none" `
    ($helpers['appStart'] -eq 'True|10|Utc|True') "none|minutes ago|kind|failed: $($helpers['appStart'])"
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
    'Test-CoworkServiceLog') + $evWatchNames + $evFixNames
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
        # for a folder that cannot be listed. Each fails the way the cmdlet
        # does, with an error that does not end the command.
        function Test-Path {
            param([string]$LiteralPath)
            return ($LiteralPath -eq 'X:\claude\logs')
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
$ckPatternMatch = [regex]::Match($watchSrc, '(?m)^\$LogTriggerFilePattern = ''([^'']+)''\s*$')
$ckFound = ($ckDefs.Count -eq $ckNames.Count -and $ckPatternMatch.Success)
Assert-Case "the checks, and the Claude logs' file pattern, were found" $ckFound `
    "missing: $(@($ckNames | Where-Object { $ckDefNames -notcontains $_ }) -join ', '); pattern $($ckPatternMatch.Success)"
# ConvertTo-Json and New-Object are the real ones: Save-WatchState writes its
# file in a temp folder, and Test-LogsForErrors makes a set with New-Object.
$ckFakes = @('Write-WatchLog', 'Test-StartupGracePeriod', 'Test-Path', 'Get-ChildItem', 'Get-Item', 'Get-PSDrive',
    'Get-ClaudeVmLogAge', 'Test-CoworkSessionActive', 'Get-Process', 'Get-HcsListCached', 'Get-CoworkHcsGuids',
    'Get-WinEvent', 'Get-Service', 'Test-FixLockHeld', 'Restart-Service', 'Start-Sleep', 'Get-Content',
    'ConvertFrom-Json', 'ConvertFrom-WatchStateTime', 'ConvertTo-Json', 'New-Object') + $ckNames + $checkErrorNames
$unfaked = @(if ($ckFound) { Get-UnfakedCommand -Asts (@($ckDefs) + $checkErrorDefs) -Faked $ckFakes } else { '<not found>' })
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
        function Test-Path { param([string]$Path) $null = $Path; return $script:CkPathThere }
        # The Claude log folder's files, from .NET; the VM cache folder is
        # empty.
        function Get-ChildItem {
            param([string]$Path, [string]$Filter, [switch]$Force, [object]$ErrorAction)
            $null = $Filter, $Force, $ErrorAction
            if ($Path -eq $ckLogDir) { return @([System.IO.DirectoryInfo]::new($Path).GetFiles('*.log')) }
            return @()
        }
        function Get-Item {
            [CmdletBinding()]
            param([string]$Path)
            if ($script:CkItem -eq 'denied') {
                $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new(
                        [System.UnauthorizedAccessException]::new("Access to the path '$Path' is denied."),
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
            return [pscustomobject]@{ Free = 100GB }
        }
        function Get-ClaudeVmLogAge {
            param([string]$LogDir)
            $null = $LogDir
            $script:CkAgeCalls++
            if ($script:CkAgeThrowsOn -contains $script:CkAgeCalls) { throw 'The VM log folder could not be read' }
            return $script:CkAge
        }
        # vmcompute as Get-Process gives it. A count that cannot be read
        # reads as nothing through the property, as a real Process's does
        # (PowerShell passes a getter's error over), and throws through the
        # getter method.
        function Get-Process {
            param([string]$Name, [object]$ErrorAction)
            $null = $Name, $ErrorAction
            if ($script:CkProcs -eq 'tripwire') { [void]$script:CkTripped.Add('Get-Process'); throw 'TRIPWIRE: Get-Process' }
            if ($script:CkProcs -eq 'throws') { throw 'The process list could not be read' }
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
        function Get-Service { [void]$script:CkTripped.Add('Get-Service'); throw 'TRIPWIRE: Get-Service' }
        function Test-FixLockHeld { [void]$script:CkTripped.Add('Test-FixLockHeld'); throw 'TRIPWIRE: Test-FixLockHeld' }
        function Restart-Service { [void]$script:CkTripped.Add('Restart-Service'); throw 'TRIPWIRE: Restart-Service' }
        function Start-Sleep { [void]$script:CkTripped.Add('Start-Sleep'); throw 'TRIPWIRE: Start-Sleep' }
        function Get-Content { [void]$script:CkTripped.Add('Get-Content'); throw 'TRIPWIRE: Get-Content' }
        function ConvertFrom-Json { [void]$script:CkTripped.Add('ConvertFrom-Json'); throw 'TRIPWIRE: ConvertFrom-Json' }
        function ConvertFrom-WatchStateTime { [void]$script:CkTripped.Add('ConvertFrom-WatchStateTime'); throw 'TRIPWIRE: ConvertFrom-WatchStateTime' }
        foreach ($d in (@($ckDefs) + $checkErrorDefs)) { . ([scriptblock]::Create($d.Extent.Text)) }
        $script:CheckErrorsSaid = @{}
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

        # The VM log's age is read, and cannot be read the second time.
        $script:CkAgeThrowsOn = @(2)
        $o['stale'] = Invoke-Ck { Test-VmLogStaleness }
        $o['staleLine'] = Get-CkLines
        $o['staleAgain'] = Invoke-Ck { Test-VmLogStaleness }
        $script:CkAgeThrowsOn = @()

        # vmcompute's handle count cannot be read; then the process list
        # cannot be read at all. The check runs once a minute, so its time
        # is put back before each call.
        $script:CkProcs = 'unreadable'
        $o['handles'] = Invoke-Ck { Test-VmcomputeHealth }
        $o['handlesLine'] = Get-CkLines
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
        $script:CkProcs = 'tripwire'

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

        $o['tripped'] = @($script:CkTripped) -join ','
        $o
    }
    if ($ck) {
        $tail = '; the same error is logged at most once an hour'
        Assert-Case "storage, a free space that cannot be read: nothing found, and said, naming the error" `
            ($ck['storage'] -eq 'null|1' -and
                $ck['storageLine'] -eq ("STORAGE WARNING: The free space on the VM image's volume could not be read " +
                    "(Access to the path 'V:\ClaudeVmCache' is denied.), so a full volume would not be seen$tail")) `
            "returned|lines: $($ck['storage']); $($ck['storageLine'])"
        Assert-Case "storage, the same error again in the hour: not said again" ($ck['storageAgain'] -eq 'null|0') $ck['storageAgain']
        Assert-Case "storage, read: nothing said" ($ck['storageOk'] -eq 'null|0') $ck['storageOk']
        Assert-Case "VM log staleness, an age that cannot be read: nothing found, and said, naming the error" `
            ($ck['stale'] -eq 'null|1' -and
                $ck['staleLine'] -eq "STALENESS WARNING: The VM log staleness check failed (The VM log folder could not be read)$tail") `
            "returned|lines: $($ck['stale']); $($ck['staleLine'])"
        Assert-Case "VM log staleness, the same error again in the hour: not said again" ($ck['staleAgain'] -eq 'null|0') $ck['staleAgain']
        Assert-Case "vmcompute, a handle count that cannot be read: nothing found, and said, naming the process and the error" `
            ($ck['handles'] -eq 'null|1' -and
                $ck['handlesLine'] -like "The handle count of vmcompute (PID 4321) could not be read (*Access is denied*), so a handle leak in it would not be seen$tail") `
            "returned|lines: $($ck['handles']); $($ck['handlesLine'])"
        Assert-Case "vmcompute, the same error again in the hour: not said again" ($ck['handlesAgain'] -eq 'null|0') $ck['handlesAgain']
        Assert-Case "vmcompute, a check that fails: nothing found, and said, naming the error" `
            ($ck['vmcompute'] -eq 'null|1' -and
                $ck['vmcomputeLine'] -eq "The vmcompute handle check failed (The process list could not be read)$tail") `
            "returned|lines: $($ck['vmcompute']); $($ck['vmcomputeLine'])"
        Assert-Case "vmcompute, the same error again in the hour: not said again" ($ck['vmcomputeAgain'] -eq 'null|0') $ck['vmcomputeAgain']
        Assert-Case "vmcompute, read: nothing said" ($ck['vmcomputeOk'] -eq 'null|0') $ck['vmcomputeOk']
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

} catch {
    # An error that escapes a section skips every section after it. Until
    # 6.3.0 nothing caught it here, and the gate passed on the cases that had
    # run before it.
    Assert-Case "every section ran to its end" $false "$($_.Exception.Message) (line $($_.InvocationInfo.ScriptLineNumber))"
} finally {
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
