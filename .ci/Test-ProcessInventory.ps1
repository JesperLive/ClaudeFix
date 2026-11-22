<#
.SYNOPSIS
    Drives the desktop-or-CLI process classifier in the shared ClaudeEnv region
    against fake processes, and checks no script kills by name any more.
.DESCRIPTION
    "claude" is not a unique process name. The Claude Code CLI runs under it
    too: the copies Claude Desktop starts for Code sessions and project
    threads, and any terminal install. Until 6.2.0 every kill in this toolkit
    stopped all of them. Measured 2026-11-21 on app 6.18355: 18 processes
    named claude, 14 the app, 3 app-started CLI, 1 terminal CLI.

    Get-Process, Get-Item and Get-CimInstance are replaced by functions here,
    which PowerShell resolves before the cmdlets, so the region's own code runs
    against processes this script describes. Nothing real is touched.

    The second half parses the three scripts, and the scripts they generate,
    and fails on a kill that takes processes chosen by the name claude without
    narrowing them to the app. Before trusting a clean result it proves the
    scan flags every broad kill the toolkit shipped until 6.2.0, and passes the
    narrowed kills that replaced them. The section comment says what the scan
    follows and what it does not.

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

. (Join-Path $PSScriptRoot 'ClaudeEnv.region.ps1')

# ---------------------------------------------------------------------
# Fakes. Each fake process is Id, Path (what MainModule.FileName returns, or
# $null for a process whose module list cannot be read), CimPath (what
# Win32_Process.ExecutablePath returns for it) and Parent (its
# Win32_Process.ParentProcessId). $script:CimThrows makes the process query
# fail, as it does when WMI is broken or the caller lacks the right.
# ---------------------------------------------------------------------
$script:FakeProcs    = @()
$script:FakeProducts = @{}
$script:CimThrows    = $false

function New-FakeProc {
    param([int]$Id, [string]$Path, [string]$CimPath, [int]$Parent = 0)
    $o = [pscustomobject]@{ Id = $Id; Name = 'claude'; CimPath = $CimPath; FakePath = $Path; Parent = $Parent }
    if ($Path) {
        $o | Add-Member -MemberType NoteProperty -Name MainModule -Value ([pscustomobject]@{ FileName = $Path })
    } else {
        $o | Add-Member -MemberType ScriptProperty -Name MainModule -Value { throw 'Access is denied' }
    }
    return $o
}

function Get-Process {
    param([string]$Name, [string]$ErrorAction)
    $null = $ErrorAction
    if ($Name -eq 'claude') { return $script:FakeProcs }
    return @()
}

function Get-Item {
    param([string]$LiteralPath, [string]$ErrorAction)
    $null = $ErrorAction
    if (-not $script:FakeProducts.ContainsKey($LiteralPath)) { throw "Cannot find path '$LiteralPath'" }
    return [pscustomobject]@{ VersionInfo = [pscustomobject]@{ ProductName = $script:FakeProducts[$LiteralPath] } }
}

function Get-CimInstance {
    param([string]$ClassName, [string]$Filter, [string]$ErrorAction)
    $null = $ClassName, $Filter, $ErrorAction
    if ($script:CimThrows) { throw 'Invalid class' }
    return @($script:FakeProcs | ForEach-Object {
        [pscustomobject]@{ ProcessId = $_.Id; ExecutablePath = $_.CimPath; ParentProcessId = $_.Parent }
    })
}

function Get-KindById {
    param($Environment)
    $map = @{}
    foreach ($r in @(Get-ClaudeProcessInventory -Environment $Environment)) { $map[[int]$r.Id] = $r.Kind }
    return $map
}

$wa      = 'C:\Program Files\WindowsApps'
$install = "$wa\Claude_1.0.4.0_x64__pzs8sxrjxfjjc"
$msixEnv = [pscustomobject]@{
    ProcessName       = 'claude'
    InstallLocation   = $install
    PackageFamilyName = 'Claude_pzs8sxrjxfjjc'
    ClaudeExe         = "$install\app\Claude.exe"
}

$appCli   = 'C:\Users\u\AppData\Local\Claude-Data\claude_pzs8sxrjxfjjc\claude-code\2.1.327\claude.exe'
$termCli  = 'C:\Users\u\.local\bin\claude.exe'
$stranger = 'C:\tools\claude.exe'
$script:FakeProducts[$termCli]  = 'Claude Code'
$script:FakeProducts[$stranger] = 'Some Other Tool'

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- MSIX install: the 6.18355 population"

$script:FakeProcs = @(
    (New-FakeProc 101 "$install\app\Claude.exe"),
    (New-FakeProc 102 "$install\app\Claude.exe"),
    (New-FakeProc 103 "$wa\Claude_1.0.3.0_x64__pzs8sxrjxfjjc\app\Claude.exe"),
    (New-FakeProc 201 $appCli),
    (New-FakeProc 202 $termCli),
    (New-FakeProc 301 $stranger),
    (New-FakeProc 302 "$wa\Claude_1.0.4.0_x64__someoneelse\app\Claude.exe"),
    (New-FakeProc 303 "${install}X\app\Claude.exe"),
    (New-FakeProc 401 $null "$install\app\Claude.exe"),
    (New-FakeProc 402 $null $null)
)
$k = Get-KindById $msixEnv
Assert-Case "app under InstallLocation is desktop"              ($k[101] -eq 'desktop') "got $($k[101])"
Assert-Case "every app process is desktop, not just the first"  ($k[102] -eq 'desktop') "got $($k[102])"
Assert-Case "app from the previous package version is desktop"  ($k[103] -eq 'desktop') "got $($k[103])"
Assert-Case "app-started CLI under claude-code is cli"          ($k[201] -eq 'cli')     "got $($k[201])"
Assert-Case "terminal CLI is cli by its product name"           ($k[202] -eq 'cli')     "got $($k[202])"
Assert-Case "an unrelated claude.exe is other"                  ($k[301] -eq 'other')   "got $($k[301])"
Assert-Case "same package name, other publisher is other"       ($k[302] -eq 'other')   "got $($k[302])"
Assert-Case "a folder that only starts like InstallLocation is other" ($k[303] -eq 'other') "got $($k[303])"
Assert-Case "unreadable module, path from CIM, is desktop"      ($k[401] -eq 'desktop') "got $($k[401])"
Assert-Case "no path anywhere is unknown"                       ($k[402] -eq 'unknown') "got $($k[402])"

$desk = @(Get-ClaudeDesktopProcess -Environment $msixEnv | ForEach-Object { [int]$_.Id } | Sort-Object)
Assert-Case "Get-ClaudeDesktopProcess returns the app and nothing else" `
    (($desk -join ',') -eq '101,102,103,401') "got $($desk -join ',')"

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Squirrel install: no package, ClaudeExe stands in"

$sqRoot = 'C:\Users\u\AppData\Local\AnthropicClaude'
$sqEnv = [pscustomobject]@{
    ProcessName       = 'claude'
    InstallLocation   = $null
    PackageFamilyName = $null
    ClaudeExe         = "$sqRoot\app-1.2.3\Claude.exe"
}
$script:FakeProcs = @(
    (New-FakeProc 501 "$sqRoot\app-1.2.3\Claude.exe"),
    (New-FakeProc 502 "$sqRoot\app-1.2.4\Claude.exe"),
    (New-FakeProc 503 $appCli),
    (New-FakeProc 504 $termCli)
)
$k = Get-KindById $sqEnv
Assert-Case "app in the discovered app-<version> is desktop"   ($k[501] -eq 'desktop') "got $($k[501])"
Assert-Case "app in a newer app-<version> is desktop"          ($k[502] -eq 'desktop') "got $($k[502])"
Assert-Case "app-started CLI is cli"                           ($k[503] -eq 'cli')     "got $($k[503])"
Assert-Case "terminal CLI is cli"                              ($k[504] -eq 'cli')     "got $($k[504])"

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- nothing discovered: the product name decides"

$bareEnv = [pscustomobject]@{
    ProcessName       = 'claude'
    InstallLocation   = $null
    PackageFamilyName = $null
    ClaudeExe         = $termCli
}
$someApp = 'D:\Apps\Claude\Claude.exe'
$script:FakeProducts[$someApp] = 'Claude'
$script:FakeProcs = @(
    (New-FakeProc 601 $someApp),
    (New-FakeProc 602 $termCli),
    (New-FakeProc 603 $stranger)
)
$k = Get-KindById $bareEnv
Assert-Case "a ClaudeExe that is the CLI is not used as the app's root" ($k[602] -eq 'cli') "got $($k[602])"
Assert-Case "ProductName Claude is desktop"                    ($k[601] -eq 'desktop') "got $($k[601])"
Assert-Case "anything else is other"                           ($k[603] -eq 'other')   "got $($k[603])"

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Test-ClaudeCodeCliPath"

Assert-Case "claude-code folder is the CLI"    (Test-ClaudeCodeCliPath $appCli)
Assert-Case "product name Claude Code is the CLI" (Test-ClaudeCodeCliPath $termCli)
Assert-Case "the app is not the CLI"           (-not (Test-ClaudeCodeCliPath "$install\app\Claude.exe"))
Assert-Case "an empty path is not the CLI"     (-not (Test-ClaudeCodeCliPath ''))

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Get-ClaudeAppCliProcess: the sessions the app started"

# The CLI copies the app starts are its children; a terminal CLI's parent is a
# shell or python. Measured 2026-11-21: four app-started CLIs, parent the app's
# main process, and one terminal CLI under python.exe. A repair ends the
# first kind and not the second, so only the first holds an automated repair
# back.
function Get-AppCliIds {
    param($Environment)
    return (@(Get-ClaudeAppCliProcess -Environment $Environment | ForEach-Object { [int]$_.Id } | Sort-Object) -join ',')
}
$script:CimThrows = $false
$script:FakeProcs = @(
    (New-FakeProc 101 "$install\app\Claude.exe" '' 4000),
    (New-FakeProc 102 "$install\app\Claude.exe" '' 101),
    (New-FakeProc 201 $appCli '' 101),
    (New-FakeProc 203 $appCli '' 102),
    (New-FakeProc 202 $termCli '' 900),
    (New-FakeProc 204 $appCli '' 999),
    (New-FakeProc 205 $termCli '' 101)
)
$got = '<threw>'
try { $got = Get-AppCliIds $msixEnv } catch { $got = "<threw: $($_.Exception.Message)>" }
Assert-Case "a CLI whose parent is the app's main process counts"       ($got -match '(^|,)201(,|$)') "got $got"
Assert-Case "a CLI whose parent is another of the app's processes counts" ($got -match '(^|,)203(,|$)') "got $got"
Assert-Case "a terminal CLI under python does not count"                ($got -notmatch '(^|,)202(,|$)') "got $got"
Assert-Case "a CLI whose parent has exited does not count"              ($got -notmatch '(^|,)204(,|$)') "got $got"
Assert-Case "a CLI by product name, started by the app, counts"         ($got -match '(^|,)205(,|$)') "got $got"
Assert-Case "exactly the app's sessions, nothing else"                  ($got -eq '201,203,205') "got $got"

$script:FakeProcs = @(
    (New-FakeProc 201 $appCli '' 101),
    (New-FakeProc 202 $termCli '' 900)
)
$got = '<threw>'
try { $got = Get-AppCliIds $msixEnv } catch { $got = "<threw: $($_.Exception.Message)>" }
Assert-Case "no app running: no session counts" ($got -eq '') "got $got"

$script:FakeProcs = @((New-FakeProc 101 "$install\app\Claude.exe" '' 4000))
$script:CimThrows = $true
$got = '<threw>'
try { $got = Get-AppCliIds $msixEnv } catch { $got = "<threw: $($_.Exception.Message)>" }
Assert-Case "no CLI at all: empty, without asking WMI" ($got -eq '') "got $got"

$script:FakeProcs = @(
    (New-FakeProc 101 "$install\app\Claude.exe" '' 4000),
    (New-FakeProc 201 $appCli '' 101)
)
$threwWith = ''
try { $null = Get-ClaudeAppCliProcess -Environment $msixEnv } catch { $threwWith = $_.Exception.Message }
Assert-Case "a failed parent query throws, so a caller can fail closed" ($threwWith -eq 'Invalid class') "threw: $threwWith"
$script:CimThrows = $false

# ---------------------------------------------------------------------
# Fix's own functions, lifted out by the parser and run here against the same
# fakes.
# ---------------------------------------------------------------------
$fixPath = Join-Path $Root 'Fix-ClaudeDesktop.ps1'
$fixAst  = [System.Management.Automation.Language.Parser]::ParseFile($fixPath, [ref]$null, [ref]$null)
function Get-FixFunction {
    param([string]$Name)
    return $fixAst.Find({ param($n)
        $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Stop-ClaudeDesktopApp stops the app and counts the rest"

# The two escalations after a relaunch kill through this. It has to stop the
# app's processes, leave the CLI, and say how many it left.
$stopDef = Get-FixFunction 'Stop-ClaudeDesktopApp'
Assert-Case "Stop-ClaudeDesktopApp was found in Fix" ($null -ne $stopDef)
if ($stopDef) {
    $script:ClaudeEnv = $msixEnv
    $script:FakeProcs = @(
        (New-FakeProc 101 "$install\app\Claude.exe"),
        (New-FakeProc 102 "$install\app\Claude.exe"),
        (New-FakeProc 201 $appCli),
        (New-FakeProc 202 $termCli),
        (New-FakeProc 301 $stranger),
        (New-FakeProc 401 $null "$install\app\Claude.exe"),
        (New-FakeProc 402 $null $null)
    )
    $outcome = & {
        $stoppedIds = New-Object System.Collections.ArrayList
        $logged     = New-Object System.Collections.ArrayList
        function Stop-Process {
            param([Parameter(ValueFromPipeline = $true)]$InputObject, [switch]$Force)
            process { $null = $Force; [void]$stoppedIds.Add([int]$InputObject.Id) }
        }
        function Log {
            param([string]$Message, [string]$Colour, [switch]$Indent)
            $null = $Colour, $Indent
            [void]$logged.Add($Message)
        }
        . ([scriptblock]::Create($stopDef.Extent.Text))
        $dryReturned = Stop-ClaudeDesktopApp -WhatIf
        $dryStopped  = $stoppedIds.Count
        $returned    = Stop-ClaudeDesktopApp
        [pscustomobject]@{
            DryReturned = $dryReturned
            DryStopped  = $dryStopped
            Returned    = $returned
            Stopped     = (@($stoppedIds | Sort-Object) -join ',')
            Logged      = ($logged -join ' | ')
        }
    }
    Assert-Case "under -WhatIf it stops nothing and says so"   ($outcome.DryStopped -eq 0 -and $outcome.DryReturned -eq 0) `
        "stopped $($outcome.DryStopped), returned $($outcome.DryReturned)"
    Assert-Case "it stops the app's processes and nothing else" ($outcome.Stopped -eq '101,102,401') "stopped $($outcome.Stopped)"
    Assert-Case "it returns how many it stopped"                ($outcome.Returned -eq 3) "returned $($outcome.Returned)"
    Assert-Case "it says how many it left alone" `
        ($outcome.Logged -match 'Left 4 other process\(es\) named claude alone') "logged: $($outcome.Logged)"
}

# Every kill site in Fix says how many processes named claude it spared, so a
# log reader can tell a spared CLI from one the kill missed.
$fixSrc = [System.IO.File]::ReadAllText($fixPath)
foreach ($site in @(
    @{ Name = '-Close';                 Rx = 'Log "Left \$procsSpared other process\(es\) named claude alone' }
    @{ Name = 'Step 1';                 Rx = 'Log "Leaving \$claudeSpared other process\(es\) named claude alone' }
    @{ Name = 'Step 4 orphan cleanup';  Rx = 'Log "Leaving \$orphanSpared other process\(es\) named claude alone' }
    @{ Name = 'Stop-ClaudeDesktopApp';  Rx = 'Log "Left \$spared other process\(es\) named claude alone' }
)) {
    Assert-Case "Fix $($site.Name) logs the processes it spared" ($fixSrc -match $site.Rx)
}
$helperCalls = [regex]::Matches($fixSrc, '(?m)^\s*\$null = Stop-ClaudeDesktopApp\s*$').Count
Assert-Case "Fix's two escalation kills go through Stop-ClaudeDesktopApp" ($helperCalls -eq 2) "found $helperCalls"
Assert-Case "Fix -Close stops only the desktop set" `
    ($fixSrc -match '\$procs = @\(\$procInventory \| Where-Object \{ \$_\.IsDesktop \}')
Assert-Case "Fix Step 1 stops only the desktop set" `
    ($fixSrc -match '\$claudeProcs\s+= @\(\$claudeInventory \| Where-Object \{ \$_\.IsDesktop \}')
Assert-Case "Fix Step 4 stops only the desktop set and the service" `
    ($fixSrc -match '\$remaining = @\(@\(\$orphanInventory \| Where-Object \{ \$_\.IsDesktop \}')

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Fix -Quiet holds back for the app's sessions"

# Until 6.3.0 the -Quiet gate asked only about CPU, a VM log under 120 s old
# and recent input. A project thread waiting on a 25-minute command uses no
# CPU and writes nothing, so an unattended run passed all three and ended the
# thread. Get-AutomatedFixBlocker asks about the app's sessions first, fails
# closed when they cannot be listed, then asks the monitor's transcript
# question. Everything it calls is faked inside the scope that runs it.
$blockDef = Get-FixFunction 'Get-AutomatedFixBlocker'
Assert-Case "Get-AutomatedFixBlocker was found in Fix" ($null -ne $blockDef)
# Since 6.3.0 the question about the app's sessions is its own function, so
# the check before each kill of the app asks the same one.
$sessionDef = Get-FixFunction 'Get-AppSessionBlocker'
if ($blockDef) {
    $blocked = & {
        Set-StrictMode -Version Latest
        $ProcessName      = 'claude'
        $script:ClaudeEnv = [pscustomobject]@{ ProcessName = 'claude'; LogDir = 'C:\nowhere' }
        $appCliNow        = @()
        $appCliThrows     = $false
        $transcriptRecent = $false
        $busyProc         = $null
        $vmLogAge         = $null
        $transcriptAsked  = New-Object System.Collections.ArrayList
        function Get-ClaudeAppCliProcess {
            param($Environment)
            $null = $Environment
            if ($appCliThrows) { throw 'Access denied' }
            return @($appCliNow)
        }
        function Test-ClaudeCodeTranscriptRecent {
            param([int]$Minutes)
            [void]$transcriptAsked.Add($Minutes)
            return $transcriptRecent
        }
        function Get-ClaudeVmLogAge { param($LogDir) $null = $LogDir; return $vmLogAge }
        function Start-Sleep { param([int]$Milliseconds, [int]$Seconds) $null = $Milliseconds, $Seconds }
        function Get-Process {
            param([string]$Name, [int]$Id, [string]$ErrorAction)
            $null = $ErrorAction
            if ($Id) { return [pscustomobject]@{ Id = $Id; SessionId = 0 } }
            if ($Name -eq 'claude' -and $busyProc) { return @($busyProc) }
            return @()
        }
        if ($sessionDef) { . ([scriptblock]::Create($sessionDef.Extent.Text)) }
        . ([scriptblock]::Create($blockDef.Extent.Text))

        $r = [ordered]@{}
        $appCliNow = @([pscustomobject]@{ Id = 201 }, [pscustomobject]@{ Id = 203 })
        $transcriptRecent = $true
        $r['cli'] = Get-AutomatedFixBlocker
        $appCliNow = @()
        $appCliThrows = $true
        $r['throws'] = Get-AutomatedFixBlocker
        $appCliThrows = $false
        $r['transcript'] = Get-AutomatedFixBlocker
        $r['minutes'] = (@($transcriptAsked) -join ',')
        $transcriptRecent = $false
        $busyProc = [pscustomobject]@{ Id = 77; Ms = 0 }
        $busyProc | Add-Member -MemberType ScriptProperty -Name TotalProcessorTime -Value { [timespan]::FromMilliseconds($this.Ms) }
        $busyProc | Add-Member -MemberType ScriptMethod -Name Refresh -Value { $this.Ms += 400 }
        $r['cpu'] = Get-AutomatedFixBlocker
        $busyProc = $null
        $vmLogAge = 30.4
        $r['vmlog'] = Get-AutomatedFixBlocker
        $vmLogAge = 600
        $r['clear'] = Get-AutomatedFixBlocker
        $r
    }
    Assert-Case "a running app session blocks, naming its PIDs" `
        ("$($blocked['cli'])" -match '^2 Claude Code session\(s\) the app started are running \(PID 201, 203\)$') "got: $($blocked['cli'])"
    Assert-Case "sessions that cannot be listed block (fail closed)" `
        ("$($blocked['throws'])" -match '^the Claude Code sessions the app started could not be listed \(Access denied\)$') "got: $($blocked['throws'])"
    Assert-Case "a transcript written in the last 35 minutes blocks" `
        ("$($blocked['transcript'])" -match 'transcript was written in the last 35 minutes') "got: $($blocked['transcript'])"
    Assert-Case "the transcript question is asked with 35 minutes" `
        ($blocked['minutes'] -match '^35(,35)*$') "asked with: $($blocked['minutes'])"
    Assert-Case "a busy claude process still blocks" `
        ("$($blocked['cpu'])" -eq 'claude process 77 is using CPU') "got: $($blocked['cpu'])"
    Assert-Case "a fresh VM log still blocks" `
        ("$($blocked['vmlog'])" -eq 'a VM log was written 30s ago') "got: $($blocked['vmlog'])"
    Assert-Case "nothing going on: no blocker" ($null -eq $blocked['clear']) "got: $($blocked['clear'])"
}
# Held back, not done: exit $script:BlockedExitCode. Until 6.3.0 it was exit 0,
# the same as a finished repair. Test-RunOutcomes runs the gate.
$gateRx = '(?ms)^if \(\$Quiet\) \{\s*\r?\n\s*\$blocker = Get-AutomatedFixBlocker\s*\r?\n\s*if \(\$blocker\) \{\s*\r?\n\s*Log "BLOCKED: \$blocker, skipping automated fix"[^\r\n]*\r?\n\s*Save-Log\s*\r?\n(\s*#[^\r\n]*\r?\n)*\s*exit \$script:BlockedExitCode\s*\r?\n'
Assert-Case "Fix's -Quiet gate exits through Get-AutomatedFixBlocker, as held back" ($fixSrc -match $gateRx)
$blockerCalls = [regex]::Matches($fixSrc, '(?m)^\s*\$blocker = Get-AutomatedFixBlocker\s*$').Count
Assert-Case "and the gate at the start asks it once" ($blockerCalls -eq 1) "found $blockerCalls call(s)"

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Fix -Quiet asks again before each kill of the app"

# The gate asks at the start, and a Smart run takes minutes: the app can
# start a Code session or a project thread in that time. Until 6.3.0 nothing
# asked again, and the next kill of the app ended it. Assert-NoNewAppSession
# asks right before Step 1's kill, Step 4's kill of the app's leftover
# processes and each escalation that stops the app. Before Step 1 nothing
# that stops the app or the service has run, so a session found there ends
# the run as held back, exit 3. After it Steps 1 and 2 have run, so a session
# found ends the run as stopped part-way, exit 4, and the service is started
# again first when this run stopped it. Until 6.3.0 Step 4 was not asked, and
# every guard ended the run as held back.
$fixLines = $fixSrc -split "\r?\n"
$guardLineRx = '^(if \(\$orphanAppCount -gt 0\) \{ )?Assert-NoNewAppSession -Before "([^"]+)" -ExitCode \$script:(BlockedExitCode|StoppedPartWayExitCode)( \})?$'
function Get-GuardBefore {
    # The guard call on the nearest code line above line $At, or $null when
    # that line is anything else; comment and blank lines in between are
    # skipped.
    param([int]$At)
    for ($k = $At - 1; $k -ge 0 -and $k -ge $At - 8; $k--) {
        $t = $fixLines[$k].Trim()
        if ($t -eq '' -or $t.StartsWith('#')) { continue }
        if ($t -match $guardLineRx) { return [pscustomobject]@{ Before = $Matches[2]; Code = $Matches[3]; Line = $t } }
        return $null
    }
    return $null
}
$killSites = @()
for ($k = 0; $k -lt $fixLines.Count; $k++) {
    if ($fixLines[$k] -match '^\s*\$null = Stop-ClaudeDesktopApp\s*$') { $killSites += $k }
    if ($fixLines[$k] -match '^\s*if \(\$PSCmdlet\.ShouldProcess\("\$\(\$claudeProcs\.Count\) Claude process\(es\)", "Stop"\)\) \{\s*$') { $killSites += $k }
    if ($fixLines[$k] -match '^\s*if \(\$PSCmdlet\.ShouldProcess\("\$\(\$remaining\.Count\) orphan process\(es\)", "Force-kill"\)\) \{\s*$') { $killSites += $k }
}
Assert-Case "four kills of the app found (Step 1, Step 4 and the two escalations)" ($killSites.Count -eq 4) "found $($killSites.Count)"
$guardedBy = @()
foreach ($k in $killSites) {
    $g = Get-GuardBefore -At $k
    Assert-Case "line $($k + 1) asks Assert-NoNewAppSession right before it" ($null -ne $g) $fixLines[$k].Trim()
    if ($g) {
        $want = if ($g.Before -eq 'Step 1') { 'BlockedExitCode' } else { 'StoppedPartWayExitCode' }
        Assert-Case "  and a session found before $($g.Before) ends the run with `$script:$want" ($g.Code -eq $want) $g.Line
        $guardedBy += $g.Before
    }
}
$guardedSeen = (@($guardedBy | Sort-Object) -join ' | ')
$guardedWant = (@('Step 1', 'Step 4', 'the escalation to Deep', 'the no-progress escalation') | Sort-Object) -join ' | '
Assert-Case "the guards stand before Step 1, Step 4 and both escalations" ($guardedSeen -eq $guardedWant) "before: $guardedSeen"
$guardCalls = [regex]::Matches($fixSrc, '(?m)^\s*(if \(\$orphanAppCount -gt 0\) \{ )?Assert-NoNewAppSession -Before "').Count
Assert-Case "and nowhere else" ($guardCalls -eq 4) "found $guardCalls call(s)"
# Step 4 asks only when a process of the app is left over, not for the
# service alone, which no Claude Code session runs in.
Assert-Case "Step 4 counts the app's own processes it is about to stop" `
    ($fixSrc -match '(?m)^\$orphanAppCount\s+= @\(\$orphanInventory \| Where-Object \{ \$_\.IsDesktop \}\)\.Count\s*$')

# Run in a child process, because it ends the run with exit. Everything it
# calls is faked there; the exit code is the outcome.
$guardDef   = Get-FixFunction 'Assert-NoNewAppSession'
$codeLine   = [regex]::Match($fixSrc, '(?m)^\$script:BlockedExitCode = \d+\s*$').Value
$partLine   = [regex]::Match($fixSrc, '(?m)^\$script:StoppedPartWayExitCode = \d+\s*$').Value
Assert-Case "Fix sets both codes the guard ends a run with" ($codeLine -ne '' -and $partLine -ne '') "[$($codeLine.Trim())] [$($partLine.Trim())]"
Assert-Case "Assert-NoNewAppSession and Get-AppSessionBlocker were found in Fix" ($null -ne $guardDef -and $null -ne $sessionDef)
# The service as Step 2 left it: stopped by this run, unless the case says
# 'notstopped' (it was not running then), or running again ('svcup'), as at
# an escalation after Step 7. Start-Service only says what it was handed,
# and throws for 'startfails'.
$restartDef = Get-FixFunction 'Start-ServiceStoppedByRun'
Assert-Case "Start-ServiceStoppedByRun was found in Fix" ($null -ne $restartDef)
$restartText = if ($restartDef) { $restartDef.Extent.Text } else { '' }
$svcFakes = @(
    '$ServiceName = "CoworkVMService"',
    '$script:ServiceStoppedByRun = ($Case -notmatch ''notstopped'')',
    'function Get-Service {',
    '    param([string]$Name, [string]$ErrorAction)',
    '    $null = $ErrorAction',
    '    $st = if ($Case -match ''svcup'') { ''Running'' } else { ''Stopped'' }',
    '    return [pscustomobject]@{ Name = $Name; Status = $st }',
    '}',
    'function Start-Service {',
    '    param([string]$Name, [string]$ErrorAction)',
    '    $null = $ErrorAction',
    '    if ($Case -match ''startfails'') { throw "Cannot start service $Name" }',
    '    [Console]::Out.WriteLine("STARTED $Name")',
    '}',
    $restartText
) -join "`r`n"
# Every place Fix stops the service notes it, so the guard knows the stop
# was this run's: Step 2, inside its ShouldProcess, Restart-CoworkService
# and Stop-CoworkServiceWithTimeout.
Assert-Case "Fix starts each run with `$script:ServiceStoppedByRun false" `
    ($fixSrc -match '(?m)^\$script:ServiceStoppedByRun = \$false\s*$')
$s2Stop = @($fixAst.FindAll({ param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and
    $n.Clauses[0].Item1.Extent.Text -eq '$PSCmdlet.ShouldProcess($ServiceName, "Stop")' }, $true))
Assert-Case "Step 2 notes that it stopped the service" `
    ($s2Stop.Count -eq 1 -and $s2Stop[0].Clauses[0].Item2.Extent.Text -match '(?m)^\s*\$script:ServiceStoppedByRun = \$true\s*$') `
    "found $($s2Stop.Count) stop block(s)"
foreach ($fn in 'Restart-CoworkService', 'Stop-CoworkServiceWithTimeout') {
    $fnDef = Get-FixFunction $fn
    Assert-Case "and so does $fn" ($null -ne $fnDef -and $fnDef.Extent.Text -match '(?m)^\s*\$script:ServiceStoppedByRun = \$true\s*$')
}
$guardFakes = @(
    $svcFakes,
    '$script:ClaudeEnv = [pscustomobject]@{ ProcessName = ''claude'' }',
    'function Log { param([string]$Message, [string]$Colour, [switch]$Indent) $null = $Colour, $Indent; Write-Output "LOG: $Message" }',
    'function Save-Log { Write-Output ''SAVED'' }',
    'function Get-ClaudeAppCliProcess {',
    '    param($Environment)',
    '    $null = $Environment',
    '    if ($Case -match ''throws'') { throw ''Access denied'' }',
    '    if ($Case -match ''session'') { return @([pscustomobject]@{ Id = 301 }) }',
    '    return @()',
    '}'
) -join "`r`n"
if ($guardDef -and $sessionDef) {
    $guardScript = Join-Path ([System.IO.Path]::GetTempPath()) ('cf-guard-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.ps1')
    $guardBody = @(
        'param([string]$Case, [string]$Before, [string]$Code, [switch]$Quiet)',
        'Set-StrictMode -Version Latest',
        $codeLine,
        $partLine,
        $guardFakes,
        $sessionDef.Extent.Text,
        $guardDef.Extent.Text,
        '$exitWith = if ($Code -eq ''Part'') { $script:StoppedPartWayExitCode } else { $script:BlockedExitCode }',
        'Assert-NoNewAppSession -Before $Before -ExitCode $exitWith',
        "Write-Output 'RETURNED'",
        'exit 0'
    ) -join "`r`n"
    [System.IO.File]::WriteAllText($guardScript, $guardBody, [System.Text.Encoding]::ASCII)
    try {
        foreach ($gc in @(
            @{ Case = 'session'; Before = 'Step 1'; Code = 'Blocked'; Quiet = $true; Exit = 3; Not = 'STOPPED PART-WAY|STARTED CoworkVMService'
               Rx = 'LOG: BLOCKED: 1 Claude Code session\(s\) the app started are running \(PID 301\); Step 1 would stop the app, so the automated fix ends here' },
            @{ Case = 'throws';  Before = 'Step 1'; Code = 'Blocked'; Quiet = $true; Exit = 3; Not = 'STOPPED PART-WAY|STARTED CoworkVMService'
               Rx = 'LOG: BLOCKED: the Claude Code sessions the app started could not be listed \(Access denied\)' },
            @{ Case = 'session'; Before = 'Step 4'; Code = 'Part'; Quiet = $true; Exit = 4; Not = 'BLOCKED'
               Rx = 'LOG: STOPPED PART-WAY: 1 Claude Code session\(s\) the app started are running \(PID 301\); Step 4 would stop the app, so the automated fix ends here, after the steps it has run' },
            @{ Case = 'throws';  Before = 'the escalation to Deep'; Code = 'Part'; Quiet = $true; Exit = 4; Not = 'BLOCKED'
               Rx = 'LOG: STOPPED PART-WAY: the Claude Code sessions the app started could not be listed \(Access denied\); the escalation to Deep would stop the app' },
            @{ Case = 'none';    Before = 'Step 4'; Code = 'Part'; Quiet = $true;  Exit = 0; Not = 'BLOCKED|STOPPED|STARTED CoworkVMService'; Rx = 'RETURNED' },
            @{ Case = 'session'; Before = 'Step 4'; Code = 'Part'; Quiet = $false; Exit = 0; Not = 'BLOCKED|STOPPED|STARTED CoworkVMService'; Rx = 'RETURNED' }
        )) {
            $gArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $guardScript, '-Case', $gc.Case, '-Before', $gc.Before, '-Code', $gc.Code)
            if ($gc.Quiet) { $gArgs += '-Quiet' }
            $gOut = (& powershell.exe @gArgs 2>&1 | Out-String)
            $gExit = $LASTEXITCODE
            $label = "$($gc.Case) before $($gc.Before)$(if ($gc.Quiet) { ', -Quiet' } else { ', manual' })"
            Assert-Case "$label exits $($gc.Exit)" ($gExit -eq $gc.Exit -and $gOut -match $gc.Rx -and $gOut -notmatch $gc.Not) "exit $gExit; output: $($gOut.Trim())"
        }
    } finally {
        Remove-Item -LiteralPath $guardScript -Force -ErrorAction SilentlyContinue
    }
}

# Step 4 itself, lifted from its inventory of the app's processes to the end
# of its kill and run in a child process. Every command it calls is faked
# there, and the kill is a fake Stop-Process that only says which Id it was
# handed; ShouldProcess is a stand-in that says it was asked. Step 2 ran
# before it, so the service is stopped and this run stopped it, unless the
# case says otherwise. Until 6.3.0 an exit 4 here left it stopped, and
# Windows does not start a manually stopped service through its trigger.
$s4First = @($fixAst.FindAll({ param($n)
    $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$orphanInventory' }, $true))
$s4Last  = @($fixAst.FindAll({ param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -eq '$remaining.Count -gt 0' }, $true))
Assert-Case "Step 4's inventory and its kill were found" ($s4First.Count -eq 1 -and $s4Last.Count -eq 1) "found $($s4First.Count), $($s4Last.Count)"
if ($s4First.Count -eq 1 -and $s4Last.Count -eq 1 -and $guardDef -and $sessionDef) {
    $s4Text = $fixAst.Extent.Text.Substring($s4First[0].Extent.StartOffset, $s4Last[0].Extent.EndOffset - $s4First[0].Extent.StartOffset)
    $s4Swaps = ([regex]::Matches($s4Text, [regex]::Escape('$PSCmdlet.ShouldProcess('))).Count
    $s4Text = $s4Text.Replace('$PSCmdlet.ShouldProcess(', '$fakeCmdlet.ShouldProcess(')
    $s4Faked = @('Get-ClaudeProcessInventory', 'Get-Process', 'Log', 'Save-Log', 'Stop-Process', 'Start-Sleep',
        'Get-ClaudeDesktopProcess', 'Get-ClaudeAppCliProcess', 'Assert-NoNewAppSession', 'Get-AppSessionBlocker',
        'Where-Object', 'ForEach-Object', 'Start-ServiceStoppedByRun', 'Get-Service', 'Start-Service')
    $s4Parsed = [System.Management.Automation.Language.Parser]::ParseInput($s4Text, [ref]$null, [ref]$null)
    $s4Unfaked = @(@($s4Parsed, $guardDef, $sessionDef, $restartDef) | Where-Object { $_ } | ForEach-Object {
        $_.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) } | ForEach-Object {
        $name = $_.GetCommandName()
        if (-not $name) { "<called through $($_.CommandElements[0].Extent.Text)>" } elseif ($s4Faked -notcontains $name) { $name }
    } | Sort-Object -Unique)
    Assert-Case "Step 4's ShouldProcess was swapped for the stand-in, and every command it calls is faked" `
        ($s4Swaps -eq 1 -and $s4Unfaked.Count -eq 0) "swapped $s4Swaps; not faked: $($s4Unfaked -join ', ')"
    if ($s4Swaps -eq 1 -and $s4Unfaked.Count -eq 0) {
        $s4Script = Join-Path ([System.IO.Path]::GetTempPath()) ('cf-step4-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.ps1')
        $s4Body = @(
            'param([string]$Case, [switch]$Quiet)',
            'Set-StrictMode -Version Latest',
            $codeLine,
            $partLine,
            $guardFakes,
            '$ServiceExe = "cowork-svc"',
            'function Start-Sleep { param([int]$Seconds, [int]$Milliseconds) $null = $Seconds, $Milliseconds }',
            'function Get-ClaudeProcessInventory {',
            '    param($Environment)',
            '    $null = $Environment',
            '    if ($Case -match ''desktop'') { return @([pscustomobject]@{ IsDesktop = $true; Process = [pscustomobject]@{ Id = 101 } }) }',
            '    return @()',
            '}',
            'function Get-Process {',
            '    param([string]$Name, [string]$ErrorAction)',
            '    $null = $ErrorAction',
            '    if ($Case -match ''service'' -and $Name -eq $ServiceExe) { return @([pscustomobject]@{ Id = 900 }) }',
            '    return @()',
            '}',
            'function Stop-Process { param([switch]$Force, [string]$ErrorAction) $null = $Force, $ErrorAction; foreach ($p in $input) { [Console]::Out.WriteLine("KILLED $($p.Id)") } }',
            'function Get-ClaudeDesktopProcess { param($Environment) $null = $Environment; return @() }',
            '$fakeCmdlet = New-Object psobject',
            '$fakeCmdlet | Add-Member -MemberType ScriptMethod -Name ShouldProcess -Value { param($Target, $Action) [Console]::Out.WriteLine("ASKED $Action"); return $true }',
            $sessionDef.Extent.Text,
            $guardDef.Extent.Text,
            $s4Text,
            "Write-Output 'STEP 4 DONE'",
            'exit 0'
        ) -join "`r`n"
        [System.IO.File]::WriteAllText($s4Script, $s4Body, [System.Text.Encoding]::ASCII)
        try {
            foreach ($sc in @(
                # Until 6.3.0 this one went straight on to the kill. It
                # starts the service Step 2 stopped again before it exits.
                @{ Case = 'desktop-session'; Quiet = $true;  Exit = 4
                   Has = @('STOPPED PART-WAY: 1 Claude Code session\(s\) the app started are running \(PID 301\); Step 4 would stop the app',
                           'STARTED CoworkVMService', 'LOG: Started CoworkVMService again, which this run had stopped')
                   Not = 'ASKED|KILLED|STEP 4 DONE|Could not start' },
                # The other paths start nothing.
                @{ Case = 'desktop';         Quiet = $true;  Exit = 0; Has = @('ASKED Force-kill', 'KILLED 101', 'STEP 4 DONE'); Not = 'STOPPED|BLOCKED|STARTED CoworkVMService' },
                @{ Case = 'service-session'; Quiet = $true;  Exit = 0; Has = @('ASKED Force-kill', 'KILLED 900', 'STEP 4 DONE'); Not = 'STOPPED|BLOCKED|KILLED 101|STARTED CoworkVMService' },
                @{ Case = 'desktop-session'; Quiet = $false; Exit = 0; Has = @('ASKED Force-kill', 'KILLED 101', 'STEP 4 DONE'); Not = 'STOPPED|BLOCKED|STARTED CoworkVMService' },
                # Not this run's stop, or running already: nothing to start.
                @{ Case = 'desktop-session-notstopped'; Quiet = $true; Exit = 4; Has = @('STOPPED PART-WAY: 1 Claude Code session')
                   Not = 'ASKED|KILLED|STEP 4 DONE|STARTED CoworkVMService|Could not start' },
                @{ Case = 'desktop-session-svcup'; Quiet = $true; Exit = 4; Has = @('STOPPED PART-WAY: 1 Claude Code session')
                   Not = 'ASKED|KILLED|STEP 4 DONE|STARTED CoworkVMService|Could not start' },
                # A start that fails is logged, and the run still ends with 4.
                @{ Case = 'desktop-session-startfails'; Quiet = $true; Exit = 4
                   Has = @('STOPPED PART-WAY: 1 Claude Code session', 'LOG: \[!\] Could not start CoworkVMService again \(Cannot start service CoworkVMService\); run Fix when Cowork is needed')
                   Not = 'ASKED|KILLED|STEP 4 DONE|STARTED CoworkVMService' }
            )) {
                $sArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $s4Script, '-Case', $sc.Case)
                if ($sc.Quiet) { $sArgs += '-Quiet' }
                $sOut = (& powershell.exe @sArgs 2>&1 | Out-String)
                $sExit = $LASTEXITCODE
                $missing = @($sc.Has | Where-Object { $sOut -notmatch $_ })
                $label = "Step 4, $($sc.Case)$(if ($sc.Quiet) { ', -Quiet' } else { ', manual' })"
                Assert-Case "$label exits $($sc.Exit)" ($sExit -eq $sc.Exit -and $missing.Count -eq 0 -and $sOut -notmatch $sc.Not) `
                    "exit $sExit; missing [$($missing -join '; ')]; output: $($sOut.Trim())"
            }
        } finally {
            Remove-Item -LiteralPath $s4Script -Force -ErrorAction SilentlyContinue
        }
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Find-ClaudeExe never hands back the CLI"

# A cache written before 6.2.0 can hold the terminal CLI's path: the old Step 1
# capture took the first process named claude, and the cache check only knew
# the claude-code folder. Every relaunch would then start the CLI. The App
# Paths and Start Menu tiers get the same check. Test-Path, Get-ItemProperty,
# Get-ChildItem and New-Object are faked inside the scope that runs it, so
# neither the registry nor the real Start Menu is read; the cache is a real
# file in a sandbox, because the function deletes a cache it rejects.
$findDef = Get-FixFunction 'Find-ClaudeExe'
Assert-Case "Find-ClaudeExe was found in Fix" ($null -ne $findDef)
if ($findDef) {
    $sandbox = Join-Path ([System.IO.Path]::GetTempPath()) ('cf-findexe-' + [guid]::NewGuid().ToString('N'))
    $realTermCli = Join-Path $sandbox '.local\bin\claude.exe'
    $realAppCli  = Join-Path $sandbox 'Claude-Data\claude_x\claude-code\2.1.327\claude.exe'
    $realApp     = Join-Path $sandbox 'Claude\Claude.exe'
    foreach ($f in $realTermCli, $realAppCli, $realApp) {
        $null = New-Item -ItemType Directory -Path (Split-Path $f -Parent) -Force
        [System.IO.File]::WriteAllText($f, '')
    }
    $script:FakeProducts[$realTermCli] = 'Claude Code'
    $script:FakeProducts[$realAppCli]  = 'Claude Code'
    $script:FakeProducts[$realApp]     = 'Claude'
    $script:ClaudeEnv         = [pscustomobject]@{ ClaudeExe = $null; ClaudeExeSource = $null }
    $script:CapturedClaudeExe = $null
    try {
        $found = & {
            $ExePathCache   = Join-Path $sandbox '.claude-exe-path'
            $appPathsTarget = $null
            $lnkTarget      = $null
            function Log { param([string]$Message, [string]$Colour, [switch]$Indent) $null = $Message, $Colour, $Indent }
            function Test-Path {
                param($Path)
                if ("$Path" -like 'HKLM:*')      { return [bool]$appPathsTarget }
                if ("$Path" -like '*Start Menu') { return [bool]$lnkTarget }
                return (Microsoft.PowerShell.Management\Test-Path -LiteralPath $Path)
            }
            function Get-ItemProperty {
                param($Path, $ErrorAction)
                $null = $Path, $ErrorAction
                return [pscustomobject]@{ '(default)' = $appPathsTarget }
            }
            function Get-ChildItem {
                param($Path, [switch]$Recurse, $Filter, $ErrorAction)
                $null = $Path, $Recurse, $Filter, $ErrorAction
                if ($lnkTarget) { return [pscustomobject]@{ FullName = 'Claude.lnk' } }
            }
            function New-Object {
                param($ComObject)
                $null = $ComObject
                $shell = [pscustomobject]@{ Target = $lnkTarget }
                $shell | Add-Member -MemberType ScriptMethod -Name CreateShortcut -Value {
                    param($LinkPath)
                    $null = $LinkPath
                    [pscustomobject]@{ TargetPath = $this.Target }
                }
                return $shell
            }
            . ([scriptblock]::Create($findDef.Extent.Text))

            $r = [ordered]@{}
            foreach ($c in @(
                @{ Key = 'cacheTerm'; Cache = $realTermCli }
                @{ Key = 'cacheCli';  Cache = $realAppCli }
                @{ Key = 'cacheApp';  Cache = $realApp }
                @{ Key = 'regTerm';   Reg = $realTermCli }
                @{ Key = 'regApp';    Reg = $realApp }
                @{ Key = 'lnkTerm';   Lnk = $realTermCli }
                @{ Key = 'lnkApp';    Lnk = $realApp }
            )) {
                Microsoft.PowerShell.Management\Remove-Item -LiteralPath $ExePathCache -Force -ErrorAction SilentlyContinue
                if ($c['Cache']) { [System.IO.File]::WriteAllText($ExePathCache, $c['Cache']) }
                $appPathsTarget = $c['Reg']
                $lnkTarget      = $c['Lnk']
                $got = Find-ClaudeExe
                $r[$c.Key] = [pscustomobject]@{
                    Got       = "$got"
                    CacheKept = [System.IO.File]::Exists($ExePathCache)
                }
            }
            $r
        }
    } finally {
        Remove-Item -LiteralPath $sandbox -Recurse -Force -ErrorAction SilentlyContinue
        $script:ClaudeEnv = $null
    }
    Assert-Case "a cached terminal CLI path is refused"      ($found['cacheTerm'].Got -eq '') "got $($found['cacheTerm'].Got)"
    Assert-Case "and the cache holding it is removed"        (-not $found['cacheTerm'].CacheKept)
    Assert-Case "a cached app-started CLI path is refused"   ($found['cacheCli'].Got -eq '')  "got $($found['cacheCli'].Got)"
    Assert-Case "and the cache holding it is removed"        (-not $found['cacheCli'].CacheKept)
    Assert-Case "a cached app path is used"                  ($found['cacheApp'].Got -eq $realApp) "got $($found['cacheApp'].Got)"
    Assert-Case "and its cache is kept"                      ($found['cacheApp'].CacheKept)
    Assert-Case "App Paths naming the terminal CLI is refused" ($found['regTerm'].Got -eq '') "got $($found['regTerm'].Got)"
    Assert-Case "App Paths naming the app is used"           ($found['regApp'].Got -eq $realApp) "got $($found['regApp'].Got)"
    Assert-Case "a Start Menu shortcut to the terminal CLI is refused" ($found['lnkTerm'].Got -eq '') "got $($found['lnkTerm'].Got)"
    Assert-Case "a Start Menu shortcut to the app is used"   ($found['lnkApp'].Got -eq $realApp) "got $($found['lnkApp'].Got)"
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- no script kills every process named claude"

# The scan reads the parse tree, not lines, so a pipe split over lines with a
# backtick, a here-string whose variables are escaped, or a lookup assigned on
# one line and stopped thirty lines later cannot hide a kill from it. It flags:
#   - processes chosen by the name claude (Get-Process claude, Get-Process
#     -Name $ProcessName, Get-Process or Get-CimInstance piped into a
#     Where-Object on the name, Get-CimInstance filtered on claude.exe) that
#     reach Stop-Process, .Kill() or a Terminate in the same pipeline;
#   - the same lookup held in a variable, when the variable's latest
#     assignment above the kill is that lookup (or a pipeline from a variable
#     that was), and the variable is piped to a kill, passed to Stop-Process,
#     has .Kill() called on it, or is walked by a foreach that kills;
#   - Stop-Process -Name claude, taskkill /IM claude (direct or through
#     Start-Process), and wmic ending claude.exe;
#   - all of that inside a script the toolkit generates from a string, which
#     is parsed as code in its turn.
# A lookup stops being broad once a later stage of its pipeline filters on
# IsDesktop, the claude-code folder, the product name or
# Test-ClaudeCodeCliPath. What it does not follow: a lookup returned from a
# function, a variable across function scopes, or a variable reassigned
# further down inside a loop. Assignments are read in source order.

$script:AstT = @{
    Cmd     = ([System.Management.Automation.Language.CommandAst])
    CmdExpr = ([System.Management.Automation.Language.CommandExpressionAst])
    Param   = ([System.Management.Automation.Language.CommandParameterAst])
    Pipe    = ([System.Management.Automation.Language.PipelineAst])
    Var     = ([System.Management.Automation.Language.VariableExpressionAst])
    ArrExpr = ([System.Management.Automation.Language.ArrayExpressionAst])
    SubExpr = ([System.Management.Automation.Language.SubExpressionAst])
    Paren   = ([System.Management.Automation.Language.ParenExpressionAst])
    Binary  = ([System.Management.Automation.Language.BinaryExpressionAst])
    ArrLit  = ([System.Management.Automation.Language.ArrayLiteralAst])
    Convert = ([System.Management.Automation.Language.ConvertExpressionAst])
    ForEach = ([System.Management.Automation.Language.ForEachStatementAst])
    Invoke  = ([System.Management.Automation.Language.InvokeMemberExpressionAst])
    Member  = ([System.Management.Automation.Language.MemberExpressionAst])
    Str     = ([System.Management.Automation.Language.StringConstantExpressionAst])
    XStr    = ([System.Management.Automation.Language.ExpandableStringExpressionAst])
}
$script:ClaudeNameRx = '(?i)^["'']?(claude(\.exe)?\*?|\$(script:)?ProcessName)["'']?$'
$script:NameFilterRx = '(?i)(\$(script:)?ProcessName\b|["'']claude(\.exe)?["'']|-(eq|like|match)\s+claude\b)'
$script:NarrowRx     = '(?i)IsDesktop|claude-code|ProductName|Test-ClaudeCodeCliPath'
$script:KillRx       = '(?i)\bStop-Process\b|\.Kill\(|\bTerminate\b'
$script:ScanAssign   = @{}

function Get-ScanVarKey {
    param($VarAst)
    return ($VarAst.VariablePath.UserPath -replace '^(script|global|local|private):', '').ToLowerInvariant()
}

function Get-ScanCmdName {
    param($Cmd)
    $n = $Cmd.GetCommandName()
    if ($n) { return [string]$n }
    return ''
}

function Test-ScanNamesClaude {
    param($Node)
    if ($null -eq $Node) { return $false }
    if ($Node -is $script:AstT.ArrLit) {
        foreach ($e in $Node.Elements) { if (Test-ScanNamesClaude $e) { return $true } }
        return $false
    }
    return ($Node.Extent.Text -match $script:ClaudeNameRx)
}

function Test-ScanNameLookup {
    # Get-Process given the name claude, by -Name or by position.
    param($Cmd)
    if ((Get-ScanCmdName $Cmd) -notmatch '(?i)^(Get-Process|gps|ps)$') { return $false }
    $els = $Cmd.CommandElements
    for ($i = 1; $i -lt $els.Count; $i++) {
        $e = $els[$i]
        if ($e -is $script:AstT.Param) {
            if ($e.Argument -and (Test-ScanNamesClaude $e.Argument)) { return $true }
        } elseif (Test-ScanNamesClaude $e) {
            return $true
        }
    }
    return $false
}

function Test-ScanBroadPipeline {
    # True when the pipeline starts from processes chosen by the name claude
    # and no later stage narrows them to the app.
    param($Pipe)
    $els = $Pipe.PipelineElements
    for ($i = 1; $i -lt $els.Count; $i++) {
        if ($els[$i].Extent.Text -match $script:NarrowRx) { return $false }
    }
    $head = $els[0]
    if ($head -is $script:AstT.CmdExpr) { return (Test-ScanBroadExpr $head.Expression) }
    if ($head -isnot $script:AstT.Cmd) { return $false }
    if (Test-ScanNameLookup $head) { return $true }
    $name = Get-ScanCmdName $head
    if ($name -match '(?i)^(Get-CimInstance|gcim|Get-WmiObject|gwmi)$' -and $head.Extent.Text -match '(?i)claude\.exe') {
        return $true
    }
    if ($name -match '(?i)^(Get-Process|gps|ps|Get-CimInstance|gcim|Get-WmiObject|gwmi)$') {
        for ($i = 1; $i -lt $els.Count; $i++) {
            $e = $els[$i]
            if ($e -is $script:AstT.Cmd -and (Get-ScanCmdName $e) -match '^(Where-Object|where|\?)$' -and
                $e.Extent.Text -match '(?i)\bName\b' -and $e.Extent.Text -match $script:NameFilterRx) {
                return $true
            }
        }
    }
    return $false
}

function Test-ScanBroadStatement {
    param($Statement)
    if ($Statement -is $script:AstT.Pipe)    { return (Test-ScanBroadPipeline $Statement) }
    if ($Statement -is $script:AstT.CmdExpr) { return (Test-ScanBroadExpr $Statement.Expression) }
    return $false
}

function Test-ScanBroadExpr {
    param($Expr)
    if ($null -eq $Expr) { return $false }
    if ($Expr -is $script:AstT.Var) {
        return (Test-ScanTainted (Get-ScanVarKey $Expr) $Expr.Extent.StartOffset)
    }
    if ($Expr -is $script:AstT.Invoke)  { return $false }
    if ($Expr -is $script:AstT.Member)  { return (Test-ScanBroadExpr $Expr.Expression) }
    if ($Expr -is $script:AstT.Convert) { return (Test-ScanBroadExpr $Expr.Child) }
    if ($Expr -is $script:AstT.Paren)   { return (Test-ScanBroadStatement $Expr.Pipeline) }
    if ($Expr -is $script:AstT.ArrExpr -or $Expr -is $script:AstT.SubExpr) {
        foreach ($s in $Expr.SubExpression.Statements) { if (Test-ScanBroadStatement $s) { return $true } }
        return $false
    }
    if ($Expr -is $script:AstT.Binary) {
        return ((Test-ScanBroadExpr $Expr.Left) -or (Test-ScanBroadExpr $Expr.Right))
    }
    if ($Expr -is $script:AstT.ArrLit) {
        foreach ($e in $Expr.Elements) { if (Test-ScanBroadExpr $e) { return $true } }
    }
    return $false
}

function Test-ScanTainted {
    # True when the variable's latest assignment above this point is a broad
    # lookup. A += keeps looking further up; a plain = ends the search.
    param([string]$Key, [int]$Offset)
    $list = $script:ScanAssign[$Key]
    if (-not $list) { return $false }
    for ($i = $list.Count - 1; $i -ge 0; $i--) {
        $a = $list[$i]
        if ($a.Extent.EndOffset -gt $Offset) { continue }
        if (Test-ScanBroadStatement $a.Right) { return $true }
        if ($a.Operator -eq [System.Management.Automation.Language.TokenKind]::Equals) { return $false }
    }
    return $false
}

function Get-ScanLabel {
    param([string]$Where, $Node)
    return "$Where line $($Node.Extent.StartLineNumber): " + (($Node.Extent.Text -split "`r?`n")[0].Trim())
}

function Find-BroadKill {
    # Parses $Code and returns the broad kills in it, how many kill sites it
    # examined, and how many parse errors it met. $Where labels each hit.
    param([string]$Code, [string]$Where, [int]$Depth = 0)
    $errs = $null
    $ast  = [System.Management.Automation.Language.Parser]::ParseInput($Code, [ref]$null, [ref]$errs)
    $saved = $script:ScanAssign
    $assign = @{}
    $assignNodes = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true) |
        Sort-Object { $_.Extent.StartOffset })
    foreach ($a in $assignNodes) {
        $left = $a.Left
        if ($left -is $script:AstT.Convert) { $left = $left.Child }
        if ($left -isnot $script:AstT.Var) { continue }
        $k = Get-ScanVarKey $left
        if (-not $assign.ContainsKey($k)) { $assign[$k] = New-Object System.Collections.ArrayList }
        [void]$assign[$k].Add($a)
    }
    $script:ScanAssign = $assign
    $hits  = New-Object System.Collections.ArrayList
    $kills = 0
    try {
        foreach ($n in $ast.FindAll({ param($x) $null = $x; $true }, $true)) {
            if ($n -is $script:AstT.Pipe) {
                $els = $n.PipelineElements
                $killing = $false
                for ($i = 1; $i -lt $els.Count; $i++) { if ($els[$i].Extent.Text -match $script:KillRx) { $killing = $true } }
                if ($killing -and (Test-ScanBroadPipeline $n)) { [void]$hits.Add((Get-ScanLabel $Where $n)) }
            } elseif ($n -is $script:AstT.Cmd) {
                $name = Get-ScanCmdName $n
                $text = $n.Extent.Text
                if ($name -match '(?i)^(Stop-Process|spps|kill)$') {
                    $kills++
                    $els = $n.CommandElements
                    for ($i = 1; $i -lt $els.Count; $i++) {
                        $e = $els[$i]
                        if ($e -is $script:AstT.Param) {
                            $pn = $e.ParameterName.ToLowerInvariant()
                            if ($pn -and ('name'.StartsWith($pn) -or 'processname'.StartsWith($pn))) {
                                $val = $e.Argument
                                if (-not $val -and $i + 1 -lt $els.Count) { $val = $els[$i + 1] }
                                if (Test-ScanNamesClaude $val) { [void]$hits.Add((Get-ScanLabel $Where $n)) }
                            }
                            continue
                        }
                        foreach ($v in $e.FindAll({ param($x) $x -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)) {
                            if (Test-ScanTainted (Get-ScanVarKey $v) $v.Extent.StartOffset) { [void]$hits.Add((Get-ScanLabel $Where $n)) }
                        }
                    }
                } elseif ($name -match '(?i)^taskkill(\.exe)?$' -or
                          ($name -match '(?i)^(Start-Process|saps|start)$' -and $text -match '(?i)taskkill')) {
                    $kills++
                    if ($text -match '(?i)/IM\s+["'']?(claude|\$(script:)?ProcessName)\b') { [void]$hits.Add((Get-ScanLabel $Where $n)) }
                } elseif ($name -match '(?i)^wmic(\.exe)?$') {
                    $kills++
                    if ($text -match '(?i)claude\.exe' -and $text -match '(?i)\b(delete|terminate)\b') { [void]$hits.Add((Get-ScanLabel $Where $n)) }
                }
            } elseif ($n -is $script:AstT.ForEach) {
                if ($n.Body.Extent.Text -match $script:KillRx -and (Test-ScanBroadStatement $n.Condition)) {
                    [void]$hits.Add((Get-ScanLabel $Where $n))
                }
            } elseif ($n -is $script:AstT.Invoke) {
                if ($n.Member.Extent.Text -eq 'Kill') {
                    $kills++
                    if (Test-ScanBroadExpr $n.Expression) { [void]$hits.Add((Get-ScanLabel $Where $n)) }
                }
            } elseif ($Depth -lt 2 -and (($n -is $script:AstT.XStr) -or
                      ($n -is $script:AstT.Str -and "$($n.StringConstantType)" -ne 'BareWord'))) {
                $v = [string]$n.Value
                if ($v -match '(?i)Stop-Process|taskkill|wmic|\.Kill\(') {
                    $inner = Find-BroadKill -Code $v -Where "$Where line $($n.Extent.StartLineNumber), generated script," -Depth ($Depth + 1)
                    foreach ($h in $inner.Hits) { [void]$hits.Add($h) }
                    $kills += $inner.Kills
                }
            }
        }
    } finally {
        $script:ScanAssign = $saved
    }
    return [pscustomobject]@{
        Hits   = @($hits | Select-Object -Unique)
        Kills  = $kills
        Errors = @($errs).Count
    }
}

# The scan has to flag every broad kill this toolkit shipped until 6.2.0, or a
# clean result below means nothing. The first five are copied from 0aa8d07,
# the last commit before the fix, with only enough around them to parse; the
# rest are shapes a regression could take.
$mustFlag = [ordered]@{
    'Fix -Close (0aa8d07)' = @'
$procs = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
if ($procs.Count -gt 0) {
    $procs | Stop-Process -Force -ErrorAction SilentlyContinue
}
'@
    'Fix Step 1 (0aa8d07)' = @'
$claudeProcs = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
if ($claudeProcs.Count -gt 0) {
    if ($PSCmdlet.ShouldProcess("$($claudeProcs.Count) Claude process(es)", "Stop")) {
        $claudeProcs | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}
'@
    'Fix Step 4 orphans (0aa8d07)' = @'
$remaining = @(Get-Process -ErrorAction SilentlyContinue | Where-Object {
    ($_.Name -eq $ProcessName) -or ($_.Name -eq $ServiceExe)
})
if ($remaining.Count -gt 0) {
    foreach ($proc in $remaining) {
        try { $proc | Stop-Process -Force -ErrorAction Stop } catch { $null = $_ }
    }
}
'@
    'Fix escalations (0aa8d07)' = @'
Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Stop-Process -Force
'@
    'Prevent NoGPU launcher (0aa8d07)' = @'
$ps1Content = @"
`$ErrorActionPreference = 'Stop'
`$running = @(Get-Process claude -ErrorAction SilentlyContinue)
if (`$running.Count -gt 0) {
    `$running | Stop-Process -Force -ErrorAction SilentlyContinue
}
"@
'@
    'Stop-Process -Name claude'          = 'Stop-Process -Name claude -Force'
    'Stop-Process -Name, quoted variable' = 'Stop-Process -Name "$ProcessName" -Force'
    'taskkill'                           = 'taskkill /F /IM claude.exe'
    'taskkill through Start-Process'     = 'Start-Process "taskkill" -ArgumentList "/F /IM claude.exe" -NoNewWindow -Wait'
    'a double-quoted name'               = 'Get-Process "claude" | Stop-Process -Force'
    'a pipe split by a backtick' = @'
Get-Process -Name $ProcessName -ErrorAction SilentlyContinue `
    | Stop-Process -Force
'@
    'a reassignment, then .Kill() in a foreach' = @'
$all = @(Get-Process -Name claude)
$all = @($all | Sort-Object Id)
foreach ($p in $all) { $p.Kill() }
'@
    'a variable passed to Stop-Process' = @'
$c = Get-Process claude
Stop-Process -InputObject $c -Force
'@
    'a CIM lookup by name' = @'
Get-CimInstance Win32_Process -Filter "Name = 'claude.exe'" | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
'@
}
foreach ($k in $mustFlag.Keys) {
    $r = Find-BroadKill -Code $mustFlag[$k] -Where 'control'
    Assert-Case "control flagged: $k" ($r.Hits.Count -gt 0 -and $r.Errors -eq 0) "hits $($r.Hits.Count), parse errors $($r.Errors)"
}

# And it has to pass the kills that replaced them, and the lookups that never
# kill, or it would be red for good and get switched off.
$mustPass = [ordered]@{
    'CPU sampling across every claude (Fix today)' = @'
$claudeCheck = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
foreach ($cp in $claudeCheck) { $cpu1 = $cp.TotalProcessorTime.TotalMilliseconds }
'@
    'the service, by name' = 'Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue | Stop-Process -Force'
    'taskkill on the service' = 'Start-Process "taskkill" -ArgumentList "/F /IM $ServiceExe.exe" -NoNewWindow -Wait'
    'Step 4 orphans (Fix today)' = @'
$orphanInventory = @(Get-ClaudeProcessInventory -Environment $script:ClaudeEnv)
$remaining = @(@($orphanInventory | Where-Object { $_.IsDesktop } | ForEach-Object { $_.Process }) +
               @(Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue))
foreach ($proc in $remaining) {
    try { $proc | Stop-Process -Force -ErrorAction Stop } catch { $null = $_ }
}
'@
    'the NoGPU launcher filter (Prevent today)' = @'
$ps1Content = @"
function Get-ClaudeAppProcess {
    @(Get-Process claude -ErrorAction SilentlyContinue | Where-Object {
        `$exe = `$null
        try { `$exe = `$_.MainModule.FileName } catch { `$null = `$_ }
        if (-not `$exe -or `$exe -match '[\\/]claude-code[\\/]') { return `$false }
        `$product = ''
        try { `$product = [string](Get-Item -LiteralPath `$exe).VersionInfo.ProductName } catch { `$null = `$_ }
        `$product -ne 'Claude Code'
    })
}
`$running = @(Get-ClaudeAppProcess)
`$running | Stop-Process -Force -ErrorAction SilentlyContinue
"@
'@
    'a broad lookup narrowed before the kill' = @'
$p = @(Get-Process -Name claude)
$p = @($p | Where-Object { -not (Test-ClaudeCodeCliPath $_.Path) })
$p | Stop-Process -Force
'@
    'a comment' = '# Get-Process claude | Stop-Process'
}
foreach ($k in $mustPass.Keys) {
    $r = Find-BroadKill -Code $mustPass[$k] -Where 'control'
    Assert-Case "control passed: $k" ($r.Hits.Count -eq 0 -and $r.Errors -eq 0) "hits: $($r.Hits -join ' | ')"
}

foreach ($n in 'Fix-ClaudeDesktop.ps1', 'Watch-ClaudeHealth.ps1', 'Prevent-ClaudeIssues.ps1') {
    $r = Find-BroadKill -Code ([System.IO.File]::ReadAllText((Join-Path $Root $n))) -Where $n
    Write-Host ("  ....  {0}: {1} kill site(s) examined" -f $n, $r.Kills)
    Assert-Case "$n parses"                                      ($r.Errors -eq 0) "$($r.Errors) parse error(s)"
    Assert-Case "$n kills no process chosen by the name claude"  ($r.Hits.Count -eq 0) ($r.Hits -join ' | ')
}

Write-Host ""
Write-Host ("Process inventory: {0} passed, {1} failed" -f $pass, $fail)
if ($fail -gt 0) {
    Write-Host "Process inventory gate FAILED" -ForegroundColor Red
    exit 1
}
Write-Host "Process inventory gate PASSED" -ForegroundColor Green
exit 0
