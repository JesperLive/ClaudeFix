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

    The second half reads the three scripts and fails on any kill that takes
    every process named claude, so the broad shape cannot come back unnoticed.

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
# $null for a process whose module list cannot be read), and CimPath (what
# Win32_Process.ExecutablePath returns for it).
# ---------------------------------------------------------------------
$script:FakeProcs    = @()
$script:FakeProducts = @{}

function New-FakeProc {
    param([int]$Id, [string]$Path, [string]$CimPath)
    $o = [pscustomobject]@{ Id = $Id; Name = 'claude'; CimPath = $CimPath; FakePath = $Path }
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
    return @($script:FakeProcs | ForEach-Object {
        [pscustomobject]@{ ProcessId = $_.Id; ExecutablePath = $_.CimPath }
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
Write-Host "--- no script kills every process named claude"

# The shapes every kill had until 6.2.0: a name lookup for claude piped
# straight into Stop-Process, or assigned and stopped a line later.
$broadPipe = '(?i)Get-Process\s+(-Name\s+)?(\$ProcessName|''?claude''?)\b[^|\r\n]*\|\s*Stop-Process'
foreach ($n in 'Fix-ClaudeDesktop.ps1', 'Watch-ClaudeHealth.ps1', 'Prevent-ClaudeIssues.ps1') {
    $lines = @([System.IO.File]::ReadAllLines((Join-Path $Root $n)))
    $hits = @()
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*#') { continue }
        if ($lines[$i] -match $broadPipe) { $hits += "line $($i + 1)" }
    }
    Assert-Case "$n pipes no name lookup for claude into Stop-Process" ($hits.Count -eq 0) ($hits -join ', ')
}

# Each kill site that used to take every claude now names the desktop filter.
$fixSrc = [System.IO.File]::ReadAllText((Join-Path $Root 'Fix-ClaudeDesktop.ps1'))
$desktopKills = [regex]::Matches($fixSrc, 'Get-ClaudeDesktopProcess -Environment \$script:ClaudeEnv \| Stop-Process').Count
Assert-Case "Fix's two escalation kills go through the desktop filter" ($desktopKills -ge 2) "found $desktopKills"
Assert-Case "Fix Step 1 stops only the desktop set" `
    ($fixSrc -match '\$claudeProcs\s+= @\(\$claudeInventory \| Where-Object \{ \$_\.IsDesktop \}')

Write-Host ""
Write-Host ("Process inventory: {0} passed, {1} failed" -f $pass, $fail)
if ($fail -gt 0) {
    Write-Host "Process inventory gate FAILED" -ForegroundColor Red
    exit 1
}
Write-Host "Process inventory gate PASSED" -ForegroundColor Green
exit 0
