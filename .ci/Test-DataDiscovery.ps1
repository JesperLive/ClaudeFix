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

    Each case builds a tree under a fresh folder in %TEMP%, points LOCALAPPDATA,
    APPDATA and USERPROFILE at it for this process only, sets write times, and
    runs the shipped functions against it: the candidate list, the choice
    Resolve-ClaudeDataPaths makes from it, Get-ClaudeEnvironment end to end,
    the Code-session guard, and Watch's Sync-ClaudeDataPaths following a move
    and refusing a fallback. Then the VM logs: their age, Watch's Cowork
    guard, and Fix's wait for a booted VM, each reading every VM log in its
    own format. Watch's and Fix's functions are pulled out of their scripts
    with the language parser, so the code under test is the shipped code, not
    a copy.

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
$watchPath = Join-Path $Root 'Watch-ClaudeHealth.ps1'
foreach ($fn in @('Test-ClaudeCodeActive', 'Sync-ClaudeDataPaths', 'Get-StealthMarkerPath', 'Test-CoworkSessionActive')) {
    . ([scriptblock]::Create((Get-FunctionSource $watchPath $fn)))
}

function Test-StartsWith {
    param([string[]]$Lines, [string]$Prefix)
    return (@($Lines | Where-Object { $_ -and $_.StartsWith($Prefix, [System.StringComparison]::Ordinal) }).Count -eq 1)
}

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
$savedProfile = $env:USERPROFILE
$savedCfg     = $env:CLAUDE_CONFIG_DIR

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
    # The guard also reads the CLI's transcripts under the user profile. On a
    # machine where Claude Code is working, the real ones are minutes old and
    # every "not active" case would fail, so the profile is a fixture too.
    $env:USERPROFILE  = Join-Path $case 'Profile'
    Remove-Item Env:\CLAUDE_USER_DATA_DIR -ErrorAction SilentlyContinue
    Remove-Item Env:\CLAUDE_CONFIG_DIR -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path $env:LOCALAPPDATA -Force | Out-Null
    New-Item -ItemType Directory -Path $env:APPDATA -Force | Out-Null
    New-Item -ItemType Directory -Path $env:USERPROFILE -Force | Out-Null
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

    $p = Resolve-ClaudeDataPaths -PackageFamilyName 'Claude_testpub' -AppDataDir (Join-Path $R 'Claude')
    Assert-Case "resolved data root is the Claude-Data root" ($p.DataRoot -eq $dataRoot -and $p.DataRootSource -ne 'fallback') "got $($p.DataRoot) ($($p.DataRootSource))"
    Assert-Case "resolved log folder is %LOCALAPPDATA%\Claude\logs" ($p.LogDir -eq (Join-Path $L 'Claude\logs') -and $p.LogDirSource -ne 'fallback') "got $($p.LogDir) ($($p.LogDirSource))"
    Assert-Case "the VM cache, bundles and session folders are under the data root" (
        $p.VmCachePath -eq (Join-Path $dataRoot 'claude-code-vm') -and
        $p.BundlePath -eq (Join-Path $dataRoot 'vm_bundles') -and
        $p.SessionsDir -eq (Join-Path $dataRoot 'local-agent-mode-sessions') -and
        $p.CodeSessionsDir -eq (Join-Path $dataRoot 'claude-code-sessions'))
    Assert-Case "the notes name the data root chosen" (Test-StartsWith $p.Notes "Data root: $dataRoot (")
    Assert-Case "the notes name the log folder chosen" (Test-StartsWith $p.Notes "Log folder: $(Join-Path $L 'Claude\logs') (")
    Assert-Case "the notes name the August logs folder passed over" (Test-StartsWith $p.Notes "Log folder passed over: $(Join-Path $dataRoot 'logs') (")

    # The whole discovery, as Watch and Fix call it. Read-only: it also asks
    # the real machine about its package and services, which only fills
    # fields this case does not look at.
    $e = Get-ClaudeEnvironment -SkipCacheInventory
    Assert-Case "Get-ClaudeEnvironment: schema version 2" ($e.SchemaVersion -eq 2)
    Assert-Case "Get-ClaudeEnvironment: AppDataDir stays the toolkit's folder" ($e.AppDataDir -eq (Join-Path $R 'Claude')) "got $($e.AppDataDir)"
    Assert-Case "Get-ClaudeEnvironment: DataRoot and LogDir come from discovery" (
        $e.DataRoot -eq $dataRoot -and $e.LogDir -eq (Join-Path $L 'Claude\logs')) "got $($e.DataRoot) | $($e.LogDir)"
    Assert-Case "Get-ClaudeEnvironment: the choice is in Warnings" (Test-StartsWith $e.Warnings "Log folder: $(Join-Path $L 'Claude\logs') (")

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

    # Claude Code's own transcripts. Measured 2026-11-21 17:25: two working
    # project threads had written theirs that minute while the newest session
    # file was 16 minutes old and remote-control.log 13.
    $P = $env:USERPROFILE
    New-FixtureFile (Join-Path $P '.claude\projects\C--Dev\old.jsonl') 30
    Assert-Case "a transcript written 30 min ago: not active" (-not (Test-ClaudeCodeActive))
    New-FixtureFile (Join-Path $P '.claude\projects\C--Dev\live.jsonl') 2
    Assert-Case "a transcript written 2 min ago, every app signal old: active" (Test-ClaudeCodeActive)
    New-FixtureFile (Join-Path $P '.claude\projects\C--Dev\live.jsonl') 30
    New-FixtureFile (Join-Path $P '.claude\projects\C--Dev\live\subagents\workflows\wf_1\agent-a.jsonl') 1
    Assert-Case "a subagent transcript written 1 min ago: active" (Test-ClaudeCodeActive)
    New-FixtureFile (Join-Path $P '.claude\projects\C--Dev\live\subagents\workflows\wf_1\agent-a.jsonl') 30
    New-FixtureFile (Join-Path $P '.claude\projects\C--Dev\memory\notes.md') 1
    New-FixtureFile (Join-Path $P '.claude\settings.json') 1
    Assert-Case "fresh files that are not transcripts: not active" (-not (Test-ClaudeCodeActive))
    $cfg = Join-Path $fixtureRoot 'current\custom-config'
    New-FixtureFile (Join-Path $cfg 'projects\D--Work\s.jsonl') 1
    Assert-Case "a transcript in a folder nothing points at: not active" (-not (Test-ClaudeCodeActive))
    $env:CLAUDE_CONFIG_DIR = $cfg
    Assert-Case "the same transcript under CLAUDE_CONFIG_DIR: active" (Test-ClaudeCodeActive)
    Remove-Item Env:\CLAUDE_CONFIG_DIR

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
    $p = Resolve-ClaudeDataPaths -PackageFamilyName 'Claude_testpub' -AppDataDir (Join-Path $R 'Claude')
    Assert-Case "resolved: the pre-6.2.0 paths, found rather than assumed" (
        $p.DataRoot -eq (Join-Path $R 'Claude') -and $p.DataRootSource -ne 'fallback' -and
        $p.LogDir -eq (Join-Path $R 'Claude\logs') -and $p.LogDirSource -ne 'fallback')
    $vm = @(Get-ClaudeVmLogCandidates -LogDir $p.LogDir)
    Assert-Case "the VM log in it is found" ($vm.Count -eq 1 -and $vm[0].Name -eq 'cowork_vm_node.log')
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
    $p = Resolve-ClaudeDataPaths -PackageFamilyName 'Claude_testpub' -AppDataDir (Join-Path $env:APPDATA 'Claude')
    Assert-Case "the data root falls back to %APPDATA%\Claude and says so" (
        $p.DataRoot -eq (Join-Path $env:APPDATA 'Claude') -and $p.DataRootSource -eq 'fallback' -and
        (Test-StartsWith $p.Notes 'No folder holding Preferences or Local State was found'))
    Assert-Case "the log folder falls back to %APPDATA%\Claude\logs and says so" (
        $p.LogDir -eq (Join-Path $env:APPDATA 'Claude\logs') -and $p.LogDirSource -eq 'fallback' -and
        (Test-StartsWith $p.Notes 'No logs folder holding a *.log was found'))
    Assert-Case "no VM log in a folder that does not exist" (@(Get-ClaudeVmLogCandidates -LogDir $p.LogDir).Count -eq 0)
    Assert-Case "not active" (-not (Test-ClaudeCodeActive))
    Assert-Case "says once that it found no session folder" `
        (@($script:WatchLogLines | Where-Object { $_ -match 'no claude-code-sessions folder' }).Count -eq 1)
    $null = Test-ClaudeCodeActive
    Assert-Case "and does not say it again" `
        (@($script:WatchLogLines | Where-Object { $_ -match 'no claude-code-sessions folder' }).Count -eq 1)

    # -----------------------------------------------------------------
    Write-Host ""
    Write-Host "--- Watch follows the folders when the app moves them, never to a fallback"

    Reset-Fixture 'moving'
    $L = $env:LOCALAPPDATA
    $R = $env:APPDATA
    # A monitor started before the update, on the old layout.
    New-FixtureFile (Join-Path $R 'Claude\Preferences') 30
    New-FixtureFile (Join-Path $R 'Claude\logs\cowork_vm_node.log') 30
    # What Sync-ClaudeDataPaths and Get-StealthMarkerPath read from Watch.
    $ClaudeAppData     = Join-Path $R 'Claude'
    $StealthMarkerName = 'stealth-relaunch'
    $start = Resolve-ClaudeDataPaths -PackageFamilyName 'Claude_testpub' -AppDataDir $ClaudeAppData
    $script:ClaudeEnv = [ordered]@{
        PackageFamilyName = 'Claude_testpub'
        VmLogFile         = (Join-Path $R 'Claude\logs\cowork_vm_node.log')
        VmLogCandidates   = @()
    }
    foreach ($k in @('DataRoot', 'DataRootSource', 'LogDir', 'LogDirSource', 'VmCachePath', 'BundlePath',
                     'SessionsDir', 'CodeSessionsDir', 'DataRootCandidates', 'LogDirCandidates')) {
        $script:ClaudeEnv[$k] = $start.$k
    }
    $script:ClaudeLogDir      = $start.LogDir
    $script:StealthMarkerPath = Get-StealthMarkerPath
    $script:NextDataPathSync  = (Get-Date).AddMinutes(10)
    Assert-Case "starts on %APPDATA%\Claude" (
        $script:ClaudeEnv.DataRoot -eq $ClaudeAppData -and $script:ClaudeLogDir -eq (Join-Path $ClaudeAppData 'logs'))

    # The app updates and moves everything.
    $dataRoot = Join-Path $L 'Claude-Data\claude_testpub'
    New-FixtureFile (Join-Path $dataRoot 'Preferences') 1
    New-FixtureFile (Join-Path $L 'Claude\logs\main.log') 1
    New-FixtureFile (Join-Path $L 'Claude\logs\cowork_vm_node.log') 1

    Sync-ClaudeDataPaths
    Assert-Case "between ten-minute marks it does not look" ($script:ClaudeEnv.DataRoot -eq $ClaudeAppData)

    Sync-ClaudeDataPaths -Force
    Assert-Case "forced, it follows the data root" ($script:ClaudeEnv.DataRoot -eq $dataRoot) "got $($script:ClaudeEnv.DataRoot)"
    Assert-Case "and the VM cache with it" ($script:ClaudeEnv.VmCachePath -eq (Join-Path $dataRoot 'claude-code-vm'))
    Assert-Case "it follows the log folder" ($script:ClaudeLogDir -eq (Join-Path $L 'Claude\logs')) "got $($script:ClaudeLogDir)"
    Assert-Case "and picks the VM log in it" ($script:ClaudeEnv.VmLogFile -eq (Join-Path $L 'Claude\logs\cowork_vm_node.log')) "got $($script:ClaudeEnv.VmLogFile)"
    Assert-Case "the stealth marker moves to the data root" ($script:StealthMarkerPath -eq (Join-Path $dataRoot 'stealth-relaunch')) "got $($script:StealthMarkerPath)"
    Assert-Case "each move is logged once" (
        (Test-StartsWith $script:WatchLogLines 'Claude data root moved: ') -and
        (Test-StartsWith $script:WatchLogLines 'Claude log folder moved: ') -and
        (Test-StartsWith $script:WatchLogLines 'VM log moved: '))
    $before = @($script:WatchLogLines).Count
    Sync-ClaudeDataPaths -Force
    Assert-Case "nothing moved, nothing logged" (@($script:WatchLogLines).Count -eq $before)

    # Every candidate vanishes for a poll, so discovery falls back to the
    # folder the app abandoned.
    Remove-Item -LiteralPath (Join-Path $L 'Claude-Data') -Recurse -Force
    Remove-Item -LiteralPath (Join-Path $L 'Claude') -Recurse -Force
    Remove-Item -LiteralPath (Join-Path $R 'Claude\Preferences') -Force
    Remove-Item -LiteralPath (Join-Path $R 'Claude\logs') -Recurse -Force
    $probe = Resolve-ClaudeDataPaths -PackageFamilyName 'Claude_testpub' -AppDataDir $ClaudeAppData
    Assert-Case "(discovery now answers with fallbacks)" ($probe.DataRootSource -eq 'fallback' -and $probe.LogDirSource -eq 'fallback')
    Sync-ClaudeDataPaths -Force
    Assert-Case "a fallback answer is not followed" (
        $script:ClaudeEnv.DataRoot -eq $dataRoot -and $script:ClaudeLogDir -eq (Join-Path $L 'Claude\logs'))

    # -----------------------------------------------------------------
    Write-Host ""
    Write-Host "--- every VM log is read, each by its own format"

    # Measured 2026-11-21 on app 6.18355: a boot is written to two logs in the
    # same second. cowork_vm_node.log gets "[VM:start] Startup complete" and
    # "[Keepalive] Starting" with dashed dates; coworkd.log gets "[coworkd]
    # connected, waiting for commands" with slashed dates, plus a clock line
    # every few hours, and is usually the newer of the two. Until this was
    # fixed each check read the one log that was newest when the script
    # started, which was coworkd.log, and found none of its markers there.
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    function Get-NodeStamp { param([double]$SecondsAgo) (Get-Date).AddSeconds(-$SecondsAgo).ToString('yyyy-MM-dd HH:mm:ss', $inv) }
    function Get-CoworkdStamp { param([double]$SecondsAgo) (Get-Date).AddSeconds(-$SecondsAgo).ToString('yyyy/MM/dd HH:mm:ss', $inv) }
    function Add-LogLine { param([string]$Path, [string]$Line) [System.IO.File]::AppendAllText($Path, $Line + "`n") }
    function Set-WriteTime { param([string]$Path, [double]$SecondsAgo) (Get-Item -LiteralPath $Path).LastWriteTime = (Get-Date).AddSeconds(-$SecondsAgo) }

    Reset-Fixture 'vmlogs'
    $vmDir   = Join-Path $env:LOCALAPPDATA 'Claude\logs'
    $nodeLog = Join-Path $vmDir 'cowork_vm_node.log'
    $cwdLog  = Join-Path $vmDir 'coworkd.log'
    Assert-Case "Get-ClaudeVmLogAge: no folder, no age" ($null -eq (Get-ClaudeVmLogAge -LogDir $vmDir))
    New-FixtureFile (Join-Path $vmDir 'main.log') 1
    Assert-Case "Get-ClaudeVmLogAge: main.log alone is no VM log" ($null -eq (Get-ClaudeVmLogAge -LogDir $vmDir))

    Add-LogLine $nodeLog "$(Get-NodeStamp 1800) [info] [Keepalive] Starting"
    Set-WriteTime $nodeLog 1800
    Add-LogLine $cwdLog "$(Get-CoworkdStamp 30) [coworkd] connected, waiting for commands"
    Set-WriteTime $cwdLog 30
    $age = Get-ClaudeVmLogAge -LogDir $vmDir
    Assert-Case "Get-ClaudeVmLogAge: coworkd.log written 30 s ago counts" ($null -ne $age -and $age -lt 120) "got $age"
    Set-WriteTime $cwdLog 1800
    Set-WriteTime $nodeLog 30
    $age = Get-ClaudeVmLogAge -LogDir $vmDir
    Assert-Case "Get-ClaudeVmLogAge: so does the node log, read now rather than at start" ($null -ne $age -and $age -lt 120) "got $age"
    Set-WriteTime $nodeLog 1800
    $age = Get-ClaudeVmLogAge -LogDir $vmDir
    Assert-Case "Get-ClaudeVmLogAge: both 30 min old" ($null -ne $age -and $age -gt 1700) "got $age"

    # Watch's Cowork guard. The old code read $script:ClaudeEnv.VmLogFile,
    # which on 6.18355 is coworkd.log, so it is pointed there here.
    $script:ClaudeLogDir = $vmDir
    $script:ClaudeEnv = [pscustomobject]@{ PackageFamilyName = 'Claude_testpub'; VmLogFile = $cwdLog }
    Add-LogLine $nodeLog "$(Get-NodeStamp 30) [info] [Keepalive] Ping ok"
    Set-WriteTime $nodeLog 30
    Add-LogLine $cwdLog "$(Get-CoworkdStamp 5) [process:oneshot-0a1b2c3d-4e5f-4a6b-8c7d-9e0f1a2b3c4d] started PID 4242"
    Set-WriteTime $cwdLog 5
    Assert-Case "Cowork guard: a [Keepalive] 30 s ago in the node log, coworkd.log newer: active" (Test-CoworkSessionActive)
    Add-LogLine $nodeLog "$(Get-NodeStamp 600) [info] [startVM] old entry written late"
    Set-WriteTime $nodeLog 1
    Assert-Case "Cowork guard: the newest tagged line decides, and it is 10 min old: not active" (-not (Test-CoworkSessionActive))
    Add-LogLine $nodeLog "$(Get-NodeStamp 60) [info] [vmOneShot] spawn"
    Assert-Case "Cowork guard: a [vmOneShot] 60 s ago: active" (Test-CoworkSessionActive)
    Remove-Item -LiteralPath $nodeLog -Force
    Assert-Case "Cowork guard: coworkd.log alone, however fresh, is not parsed: not active" (-not (Test-CoworkSessionActive))
    $script:ClaudeLogDir = $null
    Assert-Case "Cowork guard: no log folder: not active" (-not (Test-CoworkSessionActive))

    # Watch names a VM log that appears after it started.
    Reset-Fixture 'vmlater'
    $L = $env:LOCALAPPDATA
    $ClaudeAppData = Join-Path $env:APPDATA 'Claude'
    New-FixtureFile (Join-Path $L 'Claude-Data\claude_testpub\Preferences') 1
    New-FixtureFile (Join-Path $L 'Claude\logs\main.log') 1
    $start = Resolve-ClaudeDataPaths -PackageFamilyName 'Claude_testpub' -AppDataDir $ClaudeAppData
    $script:ClaudeEnv = [ordered]@{
        PackageFamilyName = 'Claude_testpub'
        VmLogFile         = $null
        VmLogCandidates   = @()
    }
    foreach ($k in @('DataRoot', 'DataRootSource', 'LogDir', 'LogDirSource', 'VmCachePath', 'BundlePath',
                     'SessionsDir', 'CodeSessionsDir', 'DataRootCandidates', 'LogDirCandidates')) {
        $script:ClaudeEnv[$k] = $start.$k
    }
    $script:ClaudeLogDir      = $start.LogDir
    $script:StealthMarkerPath = Get-StealthMarkerPath
    Sync-ClaudeDataPaths -Force
    Assert-Case "no VM log yet: none named, nothing logged" (
        $null -eq $script:ClaudeEnv.VmLogFile -and @($script:WatchLogLines).Count -eq 0) "got $($script:ClaudeEnv.VmLogFile) | $($script:WatchLogLines -join ' | ')"
    $later = Join-Path $L 'Claude\logs\coworkd.log'
    Add-LogLine $later "$(Get-CoworkdStamp 1) [coworkd] connected, waiting for commands"
    Sync-ClaudeDataPaths -Force
    Assert-Case "a VM log that appears later is named" ($script:ClaudeEnv.VmLogFile -eq $later) "got $($script:ClaudeEnv.VmLogFile)"
    Assert-Case "and that is logged once" (
        (Test-StartsWith $script:WatchLogLines "VM log found: $later") -and @($script:WatchLogLines).Count -eq 1) "got $($script:WatchLogLines -join ' | ')"
    Add-LogLine (Join-Path $L 'Claude\logs\cowork_vm_node.log') "$(Get-NodeStamp 0) [info] [Keepalive] Starting"
    Sync-ClaudeDataPaths -Force
    Assert-Case "a second log in the same folder is no move" (
        $script:ClaudeEnv.VmLogFile -eq $later -and @($script:WatchLogLines).Count -eq 1) "got $($script:WatchLogLines -join ' | ')"

    # Fix's wait for the VM, which decides whether a Smart repair escalates to
    # Deep, and Deep deletes the VM cache.
    $fixPath = Join-Path $Root 'Fix-ClaudeDesktop.ps1'
    foreach ($fn in @('Get-VmLogSizes', 'Test-VmLogReady')) {
        . ([scriptblock]::Create((Get-FunctionSource $fixPath $fn)))
    }
    Reset-Fixture 'vmready'
    $vmLogDir = Join-Path $env:LOCALAPPDATA 'Claude\logs'
    $nodeLog  = Join-Path $vmLogDir 'cowork_vm_node.log'
    $cwdLog   = Join-Path $vmLogDir 'coworkd.log'
    New-Item -ItemType Directory -Path $vmLogDir -Force | Out-Null
    # The previous boot, already in both logs when the wait begins.
    Add-LogLine $nodeLog "$(Get-NodeStamp 7200) [info] [VM:start] Startup complete"
    Add-LogLine $nodeLog "$(Get-NodeStamp 7200) [info] [Keepalive] Starting"
    Add-LogLine $cwdLog "$(Get-CoworkdStamp 7200) [coworkd] connected, waiting for commands"
    Add-LogLine $cwdLog "$(Get-CoworkdStamp 7100) [process:oneshot-11111111-2222-4333-8444-555555555555] started PID 17"
    $b = Get-VmLogSizes
    Assert-Case "Fix: the baseline holds both logs" ($b.Count -eq 2 -and $b.ContainsKey($nodeLog) -and $b.ContainsKey($cwdLog))
    Assert-Case "Fix: the previous boot's lines are not this boot" ($null -eq (Test-VmLogReady -Baseline $b))

    Add-LogLine $cwdLog "$(Get-CoworkdStamp 0) [coworkd] connected, waiting for commands"
    $r = Test-VmLogReady -Baseline $b
    Assert-Case "Fix: coworkd's 'connected, waiting for commands' is a finished boot" ($r -eq 'startup-complete') "got $r"

    $b = Get-VmLogSizes
    Add-LogLine $cwdLog "$(Get-CoworkdStamp 0) [process:oneshot-0a1b2c3d-4e5f-4a6b-8c7d-9e0f1a2b3c4d] started PID 4242"
    $r = Test-VmLogReady -Baseline $b
    Assert-Case "Fix: a 'process:oneshot-<uuid>' start is seen" ($r -eq 'process-started') "got $r"

    $b = Get-VmLogSizes
    Add-LogLine $nodeLog "$(Get-NodeStamp 0) [info] [VM:start] Startup complete"
    Add-LogLine $cwdLog "$(Get-CoworkdStamp 0) [coworkd] timesync host=2026-11-21T10:00:00Z"
    Set-WriteTime $nodeLog 5
    Set-WriteTime $cwdLog 0
    $r = Test-VmLogReady -Baseline $b
    Assert-Case "Fix: the node log's Startup complete is seen while coworkd.log is the newer" ($r -eq 'startup-complete') "got $r"

    $b = Get-VmLogSizes
    [System.IO.File]::WriteAllText($nodeLog, "$(Get-NodeStamp 0) [info] [Keepalive] Starting`n")
    $r = Test-VmLogReady -Baseline $b
    Assert-Case "Fix: a log shorter than at the baseline was rotated and is read from its start" ($r -eq 'keepalive') "got $r"

    Remove-Item -LiteralPath $cwdLog -Force
    $b = Get-VmLogSizes
    Assert-Case "Fix: a missing log is not in the baseline" ($b.Count -eq 1 -and -not $b.ContainsKey($cwdLog))
    Add-LogLine $cwdLog "$(Get-CoworkdStamp 0) [coworkd] connected, waiting for commands"
    $r = Test-VmLogReady -Baseline $b
    Assert-Case "Fix: a log that appears during the wait is read from its start" ($r -eq 'startup-complete') "got $r"
    $b = Get-VmLogSizes
    Assert-Case "Fix: nothing written since the baseline: not ready" ($null -eq (Test-VmLogReady -Baseline $b))
    Remove-Item -LiteralPath $vmLogDir -Recurse -Force
    Assert-Case "Fix: the log folder gone: an empty baseline, not ready, no throw" (
        (Get-VmLogSizes).Count -eq 0 -and $null -eq (Test-VmLogReady -Baseline $b))
} finally {
    $env:LOCALAPPDATA = $savedLocal
    $env:APPDATA      = $savedRoaming
    $env:USERPROFILE  = $savedProfile
    if ($savedUdd) { $env:CLAUDE_USER_DATA_DIR = $savedUdd } else { Remove-Item Env:\CLAUDE_USER_DATA_DIR -ErrorAction SilentlyContinue }
    if ($savedCfg) { $env:CLAUDE_CONFIG_DIR = $savedCfg } else { Remove-Item Env:\CLAUDE_CONFIG_DIR -ErrorAction SilentlyContinue }
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
