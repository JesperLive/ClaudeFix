<#
.SYNOPSIS
    Builds fake AppData trees and checks that discovery finds the app's real
    data and logs in them, and that Watch's Code-session guard reads them.
.DESCRIPTION
    Until 6.2.0 every path the toolkit read was %APPDATA%\Claude\<something>.
    Measured 2026-11-21 on app 6.18355: %APPDATA%\Claude holds only this
    toolkit's files, the data root is %LOCALAPPDATA%\Claude-Data\claude_<id>,
    whose own logs folder went quiet on 2026-08-21, and the live logs are in
    %LOCALAPPDATA%\Claude\logs. Test-ClaudeCodeActive therefore returned false
    with a Code session running.

    Each case builds a tree under a fresh folder in %TEMP%, points LOCALAPPDATA
    and APPDATA at it for this process only, sets write times, and runs the
    shipped functions against it. Test-ClaudeCodeActive is pulled out of
    Watch-ClaudeHealth.ps1 with the language parser, so the code under test is
    the shipped code, not a copy.

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

. (Join-Path $PSScriptRoot 'ClaudeEnv.region.ps1')
. ([scriptblock]::Create((Get-FunctionSource (Join-Path $Root 'Watch-ClaudeHealth.ps1') 'Test-ClaudeCodeActive')))

# What Test-ClaudeCodeActive reads from its script.
$script:WatchLogLines          = @()
$script:NoCodeSessionDirLogged = $false
$script:ClaudeEnv              = [pscustomobject]@{ PackageFamilyName = 'Claude_testpub' }
function Write-WatchLog { param([string]$Message) $script:WatchLogLines += $Message }

# ---------------------------------------------------------------------
# Fixture helpers. Everything lives under one folder this script creates and
# removes; the removal checks it is that folder before it deletes anything.
# ---------------------------------------------------------------------
$fixtureRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("claudefix-ci-" + [guid]::NewGuid().ToString('N'))
$savedLocal   = $env:LOCALAPPDATA
$savedRoaming = $env:APPDATA
$savedUdd     = $env:CLAUDE_USER_DATA_DIR

function New-FixtureFile {
    param([string]$Path, [double]$MinutesAgo, [switch]$Directory)
    $parent = Split-Path $Path -Parent
    if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    if ($Directory) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
    else { [System.IO.File]::WriteAllText($Path, 'x') }
    (Get-Item -LiteralPath $Path -Force).LastWriteTime = (Get-Date).AddMinutes(-$MinutesAgo)
}

function Reset-Fixture {
    param([string]$Name)
    $case = Join-Path $fixtureRoot $Name
    $env:LOCALAPPDATA = Join-Path $case 'Local'
    $env:APPDATA      = Join-Path $case 'Roaming'
    Remove-Item Env:\CLAUDE_USER_DATA_DIR -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path $env:LOCALAPPDATA -Force | Out-Null
    New-Item -ItemType Directory -Path $env:APPDATA -Force | Out-Null
    $script:NoCodeSessionDirLogged = $false
    $script:WatchLogLines = @()
}

$week = 7 * 24 * 60
try {
    # -----------------------------------------------------------------
    Write-Host ""
    Write-Host "--- app 6.18355 layout"

    Reset-Fixture 'current'
    $L = $env:LOCALAPPDATA
    $R = $env:APPDATA
    $dataRoot = Join-Path $L 'Claude-Data\claude_testpub'
    # %APPDATA%\Claude: the toolkit's own files, freshly written, plus the
    # config the app left there long ago.
    New-FixtureFile (Join-Path $R 'Claude\watch-logs\watch_today.log') 1
    New-FixtureFile (Join-Path $R 'Claude\fix-logs\fix_old.log') (11 * $week)
    New-FixtureFile (Join-Path $R 'Claude\claude_desktop_config.json') (11 * $week)
    # The live data root, and the logs folder it stopped writing in August.
    New-FixtureFile (Join-Path $dataRoot 'Preferences') 2
    New-FixtureFile (Join-Path $dataRoot 'Local State') (13 * $week)
    New-FixtureFile (Join-Path $dataRoot 'config.json') 20
    New-FixtureFile (Join-Path $dataRoot 'logs\main.log') (13 * $week)
    New-FixtureFile (Join-Path $dataRoot 'logs\claude.ai-web.log') (13 * $week)
    New-FixtureFile (Join-Path $dataRoot 'claude-code-sessions\acct\org\local_old.json') (6 * 60)
    # A stale copy of the toolkit's folders inside the data root, which a
    # migration left there on this machine. Not a marker.
    New-FixtureFile (Join-Path $dataRoot 'watch-logs\watch_old.log') 1
    # The live logs.
    New-FixtureFile (Join-Path $L 'Claude\logs\main.log') 1
    New-FixtureFile (Join-Path $L 'Claude\logs\claude.ai-web.log') 45

    $c = Get-ClaudeDataCandidates -PackageFamilyName 'Claude_testpub'
    Assert-Case "the Claude-Data root is the first data root" `
        (@($c.DataRoots).Count -gt 0 -and $c.DataRoots[0].Path -eq $dataRoot) "got $(@($c.DataRoots | ForEach-Object Path) -join ' | ')"
    Assert-Case "the Claude-Data root is app data" ($c.DataRoots[0].IsAppData)
    $roam = @($c.DataRoots | Where-Object { $_.Path -eq (Join-Path $R 'Claude') })
    Assert-Case "%APPDATA%\Claude is listed" ($roam.Count -eq 1)
    Assert-Case "%APPDATA%\Claude is not app data (toolkit files only)" ($roam.Count -eq 1 -and -not $roam[0].IsAppData)
    Assert-Case "%LOCALAPPDATA%\Claude\logs is the first log folder" `
        (@($c.LogDirs).Count -gt 0 -and $c.LogDirs[0].Path -eq (Join-Path $L 'Claude\logs')) "got $(@($c.LogDirs | ForEach-Object Path) -join ' | ')"
    $stale = @($c.LogDirs | Where-Object { $_.Path -eq (Join-Path $dataRoot 'logs') })
    Assert-Case "the data root's August logs folder is listed behind it" ($stale.Count -eq 1)

    Assert-Case "no recent session file and no recent app log: not active" (-not (Test-ClaudeCodeActive))

    New-FixtureFile (Join-Path $dataRoot 'claude-code-sessions\acct\org\local_live.json') 2
    Assert-Case "a session file written 2 min ago under Claude-Data: active" (Test-ClaudeCodeActive)
    (Get-Item -LiteralPath (Join-Path $dataRoot 'claude-code-sessions\acct\org\local_live.json')).LastWriteTime = (Get-Date).AddMinutes(-30)

    New-FixtureFile (Join-Path $L 'Claude\logs\claude.ai-web.log') 3
    Assert-Case "claude.ai-web.log written 3 min ago in %LOCALAPPDATA%\Claude\logs: active" (Test-ClaudeCodeActive)
    New-FixtureFile (Join-Path $L 'Claude\logs\claude.ai-web.log') 45

    New-FixtureFile (Join-Path $L 'Claude\logs\remote-control.log') 4
    Assert-Case "remote-control.log written 4 min ago: active" (Test-ClaudeCodeActive)
    New-FixtureFile (Join-Path $L 'Claude\logs\remote-control.log') 45
    Assert-Case "and inactive again once every signal is old" (-not (Test-ClaudeCodeActive))

    # -----------------------------------------------------------------
    Write-Host ""
    Write-Host "--- the layout before 6.18355: everything in %APPDATA%\Claude"

    Reset-Fixture 'legacy'
    $R = $env:APPDATA
    New-FixtureFile (Join-Path $R 'Claude\Preferences') 1
    New-FixtureFile (Join-Path $R 'Claude\Local State') 60
    New-FixtureFile (Join-Path $R 'Claude\logs\cowork_vm_node.log') 2
    New-FixtureFile (Join-Path $R 'Claude\logs\unknown-window.log') 30
    New-FixtureFile (Join-Path $R 'Claude\claude-code-sessions\acct\org\local_a.json') 30

    $c = Get-ClaudeDataCandidates -PackageFamilyName 'Claude_testpub'
    Assert-Case "%APPDATA%\Claude is the data root" `
        (@($c.DataRoots).Count -gt 0 -and $c.DataRoots[0].Path -eq (Join-Path $R 'Claude') -and $c.DataRoots[0].IsAppData)
    Assert-Case "%APPDATA%\Claude\logs is the log folder" `
        (@($c.LogDirs).Count -eq 1 -and $c.LogDirs[0].Path -eq (Join-Path $R 'Claude\logs'))
    Assert-Case "old session file and old renderer log: not active" (-not (Test-ClaudeCodeActive))
    New-FixtureFile (Join-Path $R 'Claude\logs\unknown-window.log') 1
    Assert-Case "unknown-window.log written 1 min ago: active" (Test-ClaudeCodeActive)

    # -----------------------------------------------------------------
    Write-Host ""
    Write-Host "--- MSIX LocalCache and CLAUDE_USER_DATA_DIR are candidates"

    Reset-Fixture 'redirected'
    $L = $env:LOCALAPPDATA
    $lc = Join-Path $L 'Packages\Claude_testpub\LocalCache\Roaming\Claude'
    New-FixtureFile (Join-Path $lc 'Preferences') 1
    New-FixtureFile (Join-Path $lc 'claude-code-sessions\a\b\local_x.json') 1
    $udd = Join-Path $fixtureRoot 'redirected\custom-data'
    New-FixtureFile (Join-Path $udd 'Local State') 5
    $env:CLAUDE_USER_DATA_DIR = $udd

    $c = Get-ClaudeDataCandidates -PackageFamilyName 'Claude_testpub'
    Assert-Case "the LocalCache\Roaming\Claude folder is a data root" `
        (@($c.DataRoots | Where-Object { $_.Path -eq $lc -and $_.IsAppData }).Count -eq 1)
    Assert-Case "CLAUDE_USER_DATA_DIR is a data root" `
        (@($c.DataRoots | Where-Object { $_.Path -eq $udd -and $_.Source -eq 'CLAUDE_USER_DATA_DIR' }).Count -eq 1)
    Assert-Case "the fresher of the two comes first" ($c.DataRoots[0].Path -eq $lc) "got $($c.DataRoots[0].Path)"
    Assert-Case "a session file in the redirected folder: active" (Test-ClaudeCodeActive)

    # -----------------------------------------------------------------
    Write-Host ""
    Write-Host "--- nothing installed"

    Reset-Fixture 'empty'
    $c = Get-ClaudeDataCandidates -PackageFamilyName 'Claude_testpub'
    Assert-Case "no data roots" (@($c.DataRoots).Count -eq 0)
    Assert-Case "no log folders" (@($c.LogDirs).Count -eq 0)
    Assert-Case "not active" (-not (Test-ClaudeCodeActive))
    Assert-Case "says once that it found no session folder" `
        (@($script:WatchLogLines | Where-Object { $_ -match 'no claude-code-sessions folder' }).Count -eq 1)
    $null = Test-ClaudeCodeActive
    Assert-Case "and does not say it again" `
        (@($script:WatchLogLines | Where-Object { $_ -match 'no claude-code-sessions folder' }).Count -eq 1)
} finally {
    $env:LOCALAPPDATA = $savedLocal
    $env:APPDATA      = $savedRoaming
    if ($savedUdd) { $env:CLAUDE_USER_DATA_DIR = $savedUdd } else { Remove-Item Env:\CLAUDE_USER_DATA_DIR -ErrorAction SilentlyContinue }
    $tempRoot = [System.IO.Path]::GetTempPath()
    if ((Split-Path $fixtureRoot -Leaf) -like 'claudefix-ci-*' -and
        $fixtureRoot.StartsWith($tempRoot, [System.StringComparison]::OrdinalIgnoreCase) -and
        (Test-Path -LiteralPath $fixtureRoot)) {
        Remove-Item -LiteralPath $fixtureRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ""
Write-Host ("Data discovery: {0} passed, {1} failed" -f $pass, $fail)
if ($fail -gt 0) {
    Write-Host "Data discovery gate FAILED" -ForegroundColor Red
    exit 1
}
Write-Host "Data discovery gate PASSED" -ForegroundColor Green
exit 0
