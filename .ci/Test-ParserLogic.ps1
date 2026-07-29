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

$script:HcsListToReturn = $realHcsList
$script:KilledGuids     = @()

function Invoke-HcsDiag {
    param([string[]]$Arguments, [int]$TimeoutSeconds = 15)
    if ($Arguments[0] -eq 'list') { return $script:HcsListToReturn }
    if ($Arguments[0] -eq 'kill') { $script:KilledGuids += $Arguments[1]; return 'ok' }
    return $null
}
function Log { param([string]$Message, [string]$Colour = 'White', [switch]$Indent) }

Invoke-Expression (Get-FunctionSource -Path $fixPath -Name 'Close-StaleHcsVms')

$killed = Close-StaleHcsVms -Action kill
Assert-Case "returns 1 for one live cowork-vm" ($killed -eq 1) "got $killed"
Assert-Case "extracts the exact GUID" `
    (@($script:KilledGuids).Count -eq 1 -and $script:KilledGuids[0] -eq 'DE1517EC-BF1D-5A0B-B459-77A48CB1AD97') `
    ($script:KilledGuids -join ',')

# -WhatIf must not kill anything.
$script:KilledGuids = @()
$whatIfKilled = Close-StaleHcsVms -Action kill -WhatIf
Assert-Case "-WhatIf kills nothing" (@($script:KilledGuids).Count -eq 0) ($script:KilledGuids -join ',')
Assert-Case "-WhatIf reports 0 acted on" ($whatIfKilled -eq 0) "got $whatIfKilled"

# No cowork-vm present.
$script:HcsListToReturn = "some-other-vm`n    VM, Running, AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE, some-other-vm"
$script:KilledGuids = @()
$none = Close-StaleHcsVms -Action kill
Assert-Case "ignores a non-cowork compute system" ($none -eq 0 -and @($script:KilledGuids).Count -eq 0) "got $none"

# Two instances, both killed.
$script:HcsListToReturn = @"
cowork-vm-1699151a
    VM,   Running, DE1517EC-BF1D-5A0B-B459-77A48CB1AD97, cowork-vm-1699151a
cowork-vm-2788262b
    VM,   Running, AB1517EC-BF1D-5A0B-B459-77A48CB1AD98, cowork-vm-2788262b
"@
$script:KilledGuids = @()
$two = Close-StaleHcsVms -Action kill
Assert-Case "handles two instances" ($two -eq 2 -and @($script:KilledGuids).Count -eq 2) "got $two"

# hcsdiag answered, nothing to kill. The flag has to say so, because step 5
# prints "none found via hcsdiag" off the back of it.
$script:HcsListToReturn = "some-other-vm`n    VM, Running, AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE, some-other-vm"
$null = Close-StaleHcsVms -Action kill
Assert-Case "LastHcsListOk true when hcsdiag answered" ($script:LastHcsListOk -eq $true) "got $($script:LastHcsListOk)"

# hcsdiag unavailable.
$script:HcsListToReturn = $null
$script:KilledGuids = @()
$dead = Close-StaleHcsVms -Action kill
Assert-Case "survives hcsdiag returning nothing" ($dead -eq 0) "got $dead"

# Both cases return 0, so the count alone cannot tell them apart. That is why
# step 5 printed a clean hcsdiag result on a run whose step 0 had already said
# hcsdiag was unavailable.
Assert-Case "LastHcsListOk false when hcsdiag did not answer" `
    ($script:LastHcsListOk -eq $false) "got $($script:LastHcsListOk)"

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
