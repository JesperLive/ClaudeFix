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
    own format. Last, which Defender exclusions Prevent counts as covering
    the data root, and its Defender step, with Defender faked. Watch's,
    Fix's and Prevent's functions are pulled out of
    their scripts with the language parser, so the code under test is the
    shipped code, not a copy.

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
foreach ($fn in @('Test-ClaudeCodeActive', 'Sync-ClaudeDataPaths', 'Get-StealthMarkerPath', 'Test-CoworkSessionActive',
                  'Get-CheckErrorText', 'Test-CheckErrorDue', 'Write-CheckError')) {
    . ([scriptblock]::Create((Get-FunctionSource $watchPath $fn)))
}

# The sessions the app started are read from live processes, which a fixture
# cannot make, and the real ones on the machine running this would answer
# every case below. Test-ProcessInventory drives the real function against
# fake processes; here it answers from $script:AppCli, or throws when
# $script:AppCliThrows is set. Defined after the region, so it wins.
$script:AppCli       = @()
$script:AppCliThrows = $false
function Get-ClaudeAppCliProcess {
    param($Environment)
    $null = $Environment
    if ($script:AppCliThrows) { throw 'WMI is not answering' }
    return @($script:AppCli)
}

# Stand-ins for reads that fail. Each is dot-sourced inside a & { } block,
# so it answers for the region's functions called in that block and for
# nothing after it.
#
# Get-Item, for the paths in $script:DeniedPaths: the error Get-Item writes
# for a path it may not open, id ItemExistsUnauthorizedAccessError,
# PermissionDenied, UnauthorizedAccessException. Measured 2026-11-23 with a
# Deny ACE on a temp folder, from a shell whose backup and restore
# privileges were not enabled; a shell that has them enabled, as the
# children of Git Bash do, is not refused at all. The measured message is
# "Access is denied", which does not name the path, so a line that names
# what could not be read has to name it itself. For a registry key it is
# SecurityException, "Requested registry access is not allowed." (measured
# the same day, h-prevent reg-probe). What the
# region decides on is the not-there identity, PathNotFound, ObjectNotFound,
# ItemNotFoundException (measured 2026-11-22), and any other error is a
# read that failed. Every other path is read for real, and its error is
# passed on as it came.
$script:DeniedPaths = @()
$denyGetItem = {
    function Get-Item {
        [CmdletBinding()]
        param([string]$LiteralPath, [switch]$Force)
        if ($script:DeniedPaths -contains $LiteralPath) {
            if ($LiteralPath -like 'HK*:*') {
                $ex = New-Object System.Security.SecurityException 'Requested registry access is not allowed.'
                $PSCmdlet.WriteError((New-Object System.Management.Automation.ErrorRecord($ex, 'System.Security.SecurityException',
                    [System.Management.Automation.ErrorCategory]::PermissionDenied, $LiteralPath)))
                return
            }
            $ex = New-Object System.UnauthorizedAccessException 'Access is denied'
            $PSCmdlet.WriteError((New-Object System.Management.Automation.ErrorRecord($ex, 'ItemExistsUnauthorizedAccessError',
                [System.Management.Automation.ErrorCategory]::PermissionDenied, $LiteralPath)))
            return
        }
        try {
            Microsoft.PowerShell.Management\Get-Item -LiteralPath $LiteralPath -Force:$Force -ErrorAction Stop
        } catch {
            $PSCmdlet.WriteError($_)
        }
    }
}
# Get-ChildItem, for the transcript walk: one error for a folder in the walk,
# then the real listing. $script:WalkError 'denied' is a folder that could
# not be listed (DirUnauthorizedAccessError, PermissionDenied,
# UnauthorizedAccessException, the error Get-ChildItem -Recurse writes for
# one); 'gone' is a folder that went away during the walk (PathNotFound,
# ObjectNotFound, ItemNotFoundException, measured).
$script:WalkError = 'denied'
$walkGetChildItem = {
    function Get-ChildItem {
        [CmdletBinding()]
        param([string]$LiteralPath, [string]$Filter, [switch]$Recurse, [switch]$File, [switch]$Directory, [switch]$Force)
        $sub = Join-Path $LiteralPath 'p-walk'
        if ($script:WalkError -eq 'gone') {
            $ex = New-Object System.Management.Automation.ItemNotFoundException "Cannot find path '$sub' because it does not exist."
            $rec = New-Object System.Management.Automation.ErrorRecord($ex, 'PathNotFound',
                [System.Management.Automation.ErrorCategory]::ObjectNotFound, $sub)
        } else {
            $ex = New-Object System.UnauthorizedAccessException "Access to the path '$sub' is denied."
            $rec = New-Object System.Management.Automation.ErrorRecord($ex, 'DirUnauthorizedAccessError',
                [System.Management.Automation.ErrorCategory]::PermissionDenied, $sub)
        }
        $PSCmdlet.WriteError($rec)
        Microsoft.PowerShell.Management\Get-ChildItem -LiteralPath $LiteralPath -Filter $Filter -Recurse:$Recurse `
            -File:$File -Directory:$Directory -Force:$Force
    }
}

function Test-StartsWith {
    param([string[]]$Lines, [string]$Prefix)
    return (@($Lines | Where-Object { $_ -and $_.StartsWith($Prefix, [System.StringComparison]::Ordinal) }).Count -eq 1)
}

# What Test-ClaudeCodeActive reads from its script.
$script:WatchLogLines          = @()
$script:CheckErrorsSaid        = @{}
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
    # A variable that is not set is the one way these two fail, and not
    # set is what they are for.
    Remove-Item Env:\CLAUDE_USER_DATA_DIR -ErrorAction SilentlyContinue
    Remove-Item Env:\CLAUDE_CONFIG_DIR -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path $env:LOCALAPPDATA -Force | Out-Null
    New-Item -ItemType Directory -Path $env:APPDATA -Force | Out-Null
    New-Item -ItemType Directory -Path $env:USERPROFILE -Force | Out-Null
    $script:NoCodeSessionDirLogged = $false
    $script:WatchLogLines = @()
}

$week = 7 * 24 * 60
$fixtureCasesEnded = $false
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
    New-FixtureFile (Join-Path $P '.claude\projects\C--Dev\old.jsonl') 40
    Assert-Case "a transcript written 40 min ago: not active" (-not (Test-ClaudeCodeActive))
    New-FixtureFile (Join-Path $P '.claude\projects\C--Dev\live.jsonl') 2
    Assert-Case "a transcript written 2 min ago, every app signal old: active" (Test-ClaudeCodeActive)
    # A tool call writes nothing to the transcript until it returns, and a
    # foreground command may run 30 minutes. Measured 2026-11-21: three Bash
    # calls of 25 minutes each in one project thread, nothing written during
    # the first two.
    New-FixtureFile (Join-Path $P '.claude\projects\C--Dev\live.jsonl') 25
    Assert-Case "a transcript last written 25 min ago, inside a long tool call: active" (Test-ClaudeCodeActive)
    New-FixtureFile (Join-Path $P '.claude\projects\C--Dev\live.jsonl') 34
    Assert-Case "34 min ago, past the 30-minute command limit: still active" (Test-ClaudeCodeActive)
    New-FixtureFile (Join-Path $P '.claude\projects\C--Dev\live.jsonl') 36
    Assert-Case "36 min ago: not active" (-not (Test-ClaudeCodeActive))
    New-FixtureFile (Join-Path $P '.claude\projects\C--Dev\live\subagents\workflows\wf_1\agent-a.jsonl') 1
    Assert-Case "a subagent transcript written 1 min ago: active" (Test-ClaudeCodeActive)
    New-FixtureFile (Join-Path $P '.claude\projects\C--Dev\live\subagents\workflows\wf_1\agent-a.jsonl') 40
    New-FixtureFile (Join-Path $P '.claude\projects\C--Dev\memory\notes.md') 1
    New-FixtureFile (Join-Path $P '.claude\settings.json') 1
    Assert-Case "fresh files that are not transcripts: not active" (-not (Test-ClaudeCodeActive))
    $cfg = Join-Path $fixtureRoot 'current\custom-config'
    New-FixtureFile (Join-Path $cfg 'projects\D--Work\s.jsonl') 1
    Assert-Case "a transcript in a folder nothing points at: not active" (-not (Test-ClaudeCodeActive))
    $env:CLAUDE_CONFIG_DIR = $cfg
    Assert-Case "the same transcript under CLAUDE_CONFIG_DIR: active" (Test-ClaudeCodeActive)
    Remove-Item Env:\CLAUDE_CONFIG_DIR

    # A session the app started counts by its process alone. One waiting on a
    # long command uses no CPU and writes nothing anywhere this reads.
    $script:AppCli = @([pscustomobject]@{ Id = 201 })
    Assert-Case "an app-started CLI, every other signal old: active" (Test-ClaudeCodeActive)
    $script:AppCli = @()
    $script:AppCliThrows = $true
    $logBefore = @($script:WatchLogLines).Count
    Assert-Case "the app's sessions cannot be listed: active (fails closed)" (Test-ClaudeCodeActive)
    Assert-Case "and the log says why" (@(@($script:WatchLogLines) | Select-Object -Skip $logBefore |
        Where-Object { $_ -match 'WMI is not answering.*assuming a Code session IS active' }).Count -eq 1)
    $script:AppCliThrows = $false
    Assert-Case "no app session and every signal old: not active" (-not (Test-ClaudeCodeActive))

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
    # A node log that cannot be read is not one with no session in it.
    # Until 6.3.1 it was read with SilentlyContinue, and a log that could
    # not be read answered "no session", which let a repair through. Here
    # Get-Content refuses it with an error that is not the one for a file
    # that is not there.
    $res = & {
        function Get-Content {
            [CmdletBinding()]
            param([string]$LiteralPath, [int]$Tail)
            $null = $Tail
            $ex = New-Object System.UnauthorizedAccessException "Access to the path '$LiteralPath' is denied."
            $PSCmdlet.WriteError((New-Object System.Management.Automation.ErrorRecord($ex, 'GetContentReaderUnauthorizedAccessError',
                [System.Management.Automation.ErrorCategory]::PermissionDenied, $LiteralPath)))
        }
        $script:WatchLogLines = @()
        $script:CheckErrorsSaid = @{}
        $first = Test-CoworkSessionActive
        $second = Test-CoworkSessionActive
        [pscustomobject]@{ First = $first; Second = $second; Lines = @($script:WatchLogLines) }
    }
    $want = "The Cowork VM log $nodeLog could not be read (Access to the path '$nodeLog' is denied.), so a Cowork session is " +
        'taken to be active and nothing that would end one is done; the same error is logged at most once an hour'
    Assert-Case "Cowork guard: a node log that cannot be read: taken as active, and said once" ($res.First -eq $true -and $res.Second -eq $true -and
        @($res.Lines).Count -eq 1 -and @($res.Lines)[0] -eq $want) "got $($res.First), $($res.Second); $(@($res.Lines) -join ' | ')"
    # A log whose lines cannot be read through is not one with no session
    # in it either. Until 6.3.1 that answered "no session" without a word.
    # Here the one line is an object whose text cannot be had: PowerShell
    # then throws "Cannot convert value to type System.String." (measured
    # on 5.1, 2026-11-22).
    $res = & {
        function Get-Content {
            [CmdletBinding()]
            param([string]$LiteralPath, [int]$Tail)
            $null = $LiteralPath, $Tail
            $bad = New-Object psobject
            $bad | Add-Member -MemberType ScriptMethod -Name ToString -Value { throw 'the line could not be decoded' } -Force
            return $bad
        }
        $script:WatchLogLines = @()
        $script:CheckErrorsSaid = @{}
        $r = Test-CoworkSessionActive
        [pscustomobject]@{ Result = $r; Lines = @($script:WatchLogLines) }
    }
    $want = "The Cowork VM log $nodeLog could not be read through (Cannot convert value to type System.String.), so a Cowork session is " +
        'taken to be active and nothing that would end one is done; the same error is logged at most once an hour'
    Assert-Case "Cowork guard: a node log whose lines cannot be read through: taken as active, and said" ($res.Result -eq $true -and
        @($res.Lines).Count -eq 1 -and @($res.Lines)[0] -eq $want) "got $($res.Result); $(@($res.Lines) -join ' | ')"
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
    foreach ($fn in @('Get-VmLogSizes', 'Test-VmLogReady', 'Write-VmLogUnread')) {
        $src = $null
        try { $src = Get-FunctionSource $fixPath $fn } catch { Assert-Case "Fix defines $fn" $false "$_" }
        if ($src) { . ([scriptblock]::Create($src)) }
    }
    $script:FixLogLines     = @()
    $script:VmLogUnread     = $false
    $script:VmLogUnreadSaid = $false
    function Log { param([string]$Message, [string]$Colour, [switch]$Indent) $null = $Colour, $Indent; $script:FixLogLines += $Message }
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
    Assert-Case "Fix: and nothing was said to be unreadable" (-not $script:VmLogUnread -and @($script:FixLogLines).Count -eq 0) ($script:FixLogLines -join ' | ')

    # -----------------------------------------------------------------
    Write-Host ""
    Write-Host "--- Fix's wait: a VM log that could not be read"

    # Until 6.3.1 a VM log that could not be read was passed over as one with
    # nothing new in it, and a wait that saw nothing new escalated Smart to
    # Deep, which deletes the VM cache. Since 6.3.1 the answer is
    # 'unreadable', said once. A log held open with no sharing is a read that
    # really fails, made without touching an access rule.
    Reset-Fixture 'vmunread'
    $vmLogDir = Join-Path $env:LOCALAPPDATA 'Claude\logs'
    $nodeLog  = Join-Path $vmLogDir 'cowork_vm_node.log'
    $cwdLog   = Join-Path $vmLogDir 'coworkd.log'
    New-Item -ItemType Directory -Path $vmLogDir -Force | Out-Null
    Add-LogLine $nodeLog "$(Get-NodeStamp 7200) [info] [Keepalive] Starting"
    Add-LogLine $cwdLog "$(Get-CoworkdStamp 7200) [coworkd] connected, waiting for commands"
    $script:FixLogLines     = @()
    $script:VmLogUnread     = $false
    $script:VmLogUnreadSaid = $false
    $b = Get-VmLogSizes
    Add-LogLine $cwdLog "$(Get-CoworkdStamp 0) [coworkd] connected, waiting for commands"
    $lock = [System.IO.File]::Open($cwdLog, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    try {
        $r1 = Test-VmLogReady -Baseline $b
        $r2 = Test-VmLogReady -Baseline $b
    } finally { $lock.Dispose() }
    Assert-Case "Fix: new lines in a log held open by another process: unreadable, not nothing new" (
        $r1 -eq 'unreadable' -and $r2 -eq 'unreadable') "got $r1, $r2"
    Assert-Case "Fix: that is said once, with the error" (
        $script:VmLogUnread -and @($script:FixLogLines).Count -eq 1 -and
        $script:FixLogLines[0] -match ('^\[!\] The VM log ' + [regex]::Escape($cwdLog) +
            ' could not be read \(.*being used by another process.*\); whether the workspace is ready cannot be read from the logs$')) `
        ($script:FixLogLines -join ' | ')
    $r = Test-VmLogReady -Baseline $b
    Assert-Case "Fix: read again once it is let go, the boot in it is seen" ($r -eq 'startup-complete') "got $r"

    $b = Get-VmLogSizes
    Add-LogLine $nodeLog "$(Get-NodeStamp 0) [info] [VM:start] Startup complete"
    Add-LogLine $cwdLog "$(Get-CoworkdStamp 0) [coworkd] timesync host=2026-11-21T10:00:00Z"
    $lock = [System.IO.File]::Open($cwdLog, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    try { $r = Test-VmLogReady -Baseline $b } finally { $lock.Dispose() }
    Assert-Case "Fix: a boot marker in a log that could be read still answers" ($r -eq 'startup-complete') "got $r"

    # A log whose entry could not be read at all: the list of logs fails, so
    # the wait has no baseline, and every answer is 'unreadable'.
    $script:FixLogLines     = @()
    $script:VmLogUnread     = $false
    $script:VmLogUnreadSaid = $false
    $res = & {
        . $denyGetItem
        $script:DeniedPaths = @($cwdLog)
        $sizes = Get-VmLogSizes
        $o = [pscustomobject]@{
            NoSizes  = ($null -eq $sizes)
            NoBase   = (Test-VmLogReady -Baseline $sizes)
            WithBase = (Test-VmLogReady -Baseline @{})
        }
        $script:DeniedPaths = @()
        $o
    }
    Assert-Case "Fix: a VM log whose entry could not be read: no baseline" ($res.NoSizes)
    Assert-Case "Fix: and every answer is unreadable" ($res.NoBase -eq 'unreadable' -and $res.WithBase -eq 'unreadable') "got $($res.NoBase), $($res.WithBase)"
    # The error Get-Item gives says only "Access is denied", so the line
    # names the log from the error's target. Until 6.3.1 it was "A VM log
    # could not be read (Access is denied)", which named none.
    Assert-Case "Fix: said once, naming the log and the error" (
        $script:VmLogUnread -and @($script:FixLogLines).Count -eq 1 -and
        $script:FixLogLines[0] -eq "[!] The VM log $cwdLog could not be read (Access is denied); whether the workspace is ready cannot be read from the logs") `
        ($script:FixLogLines -join ' | ')

    # An error that names no path, from a caller that knows none: the log
    # folder is named.
    $script:FixLogLines     = @()
    $script:VmLogUnread     = $false
    $script:VmLogUnreadSaid = $false
    Write-VmLogUnread (New-Object System.Management.Automation.ErrorRecord((New-Object System.IO.IOException 'The device is not ready.'),
        'IOException', 'ReadError', $null))
    Assert-Case "Fix: an error with no path in it names the log folder" (
        $script:VmLogUnread -and @($script:FixLogLines).Count -eq 1 -and
        $script:FixLogLines[0] -eq "[!] A VM log in $vmLogDir could not be read (The device is not ready.); whether the workspace is ready cannot be read from the logs") `
        ($script:FixLogLines -join ' | ')

    # -----------------------------------------------------------------
    Write-Host ""
    Write-Host "--- the region tells a folder that could not be read from one not there"

    # Until 6.3.1 each read below went through Test-Path or SilentlyContinue,
    # and a folder or log that could not be read looked like one that was
    # not there.
    Reset-Fixture 'unread'
    $L = $env:LOCALAPPDATA
    $R = $env:APPDATA
    $dataRoot = Join-Path $L 'Claude-Data\claude_testpub'
    New-FixtureFile (Join-Path $dataRoot 'Preferences') 1
    New-FixtureFile (Join-Path $R 'Claude\Preferences') 600
    New-FixtureFile (Join-Path $L 'Claude\logs\main.log') 1
    $vmDirU = Join-Path $L 'Claude\logs'
    New-FixtureFile (Join-Path $vmDirU 'cowork_vm_node.log') 1
    $projects = Join-Path $env:USERPROFILE '.claude\projects'
    New-FixtureFile (Join-Path $projects 'p1\old.jsonl') 600
    $res = & {
        . $denyGetItem
        $o = [ordered]@{}
        $script:DeniedPaths = @()
        $o.Open = Get-ClaudeDataCandidates -PackageFamilyName 'Claude_testpub'
        $script:DeniedPaths = @((Join-Path $L 'Claude-Data'))
        $o.Shut = Get-ClaudeDataCandidates -PackageFamilyName 'Claude_testpub'
        $o.Resolved = Resolve-ClaudeDataPaths -PackageFamilyName 'Claude_testpub' -AppDataDir (Join-Path $R 'Claude')
        $script:DeniedPaths = @((Join-Path $vmDirU 'coworkd.log'))
        $errOf = { param($e) "$(([string]$e.FullyQualifiedErrorId).Split(',')[0])|$($e.TargetObject)|$(Get-ClaudeErrorText $e)" }
        try { $null = Get-ClaudeVmLogCandidates -LogDir $vmDirU; $o.Cands = 'no throw' } catch { $o.Cands = & $errOf $_ }
        try { $null = Get-ClaudeVmLogAge -LogDir $vmDirU; $o.Age = 'no throw' } catch { $o.Age = & $errOf $_ }
        $script:DeniedPaths = @($projects)
        try { $o.Projects = [string](Test-ClaudeCodeTranscriptRecent -Minutes 35) } catch { $o.Projects = "THREW: $(Get-ClaudeErrorText $_)" }
        $script:DeniedPaths = @()
        $o.OpenCands = @(Get-ClaudeVmLogCandidates -LogDir $vmDirU).Count
        $o.OpenRecent = [string](Test-ClaudeCodeTranscriptRecent -Minutes 35)
        $o
    }
    Assert-Case "through the stand-in, folders that are not there are no error" (
        @($res.Open.Errors).Count -eq 0 -and $res.Open.DataRoots[0].Path -eq $dataRoot) "errors: $(@($res.Open.Errors) -join ' | ')"
    Assert-Case "a candidate folder that could not be read is an error that names it" (
        @($res.Shut.Errors).Count -eq 1 -and
        $res.Shut.Errors[0] -eq "$(Join-Path $L 'Claude-Data') could not be read: Access is denied") `
        "errors: $(@($res.Shut.Errors) -join ' | ')"
    Assert-Case "and the data root behind it is not offered" (@($res.Shut.DataRoots | Where-Object { $_.Path -eq $dataRoot }).Count -eq 0)
    Assert-Case "Resolve-ClaudeDataPaths passes it on, and its notes say the choice may be wrong" (
        @($res.Resolved.Errors).Count -eq 1 -and
        (Test-StartsWith $res.Resolved.Notes "Candidate not read, so the choice above may be wrong: $(Join-Path $L 'Claude-Data') could not be read: ")) `
        "notes: $(@($res.Resolved.Notes) -join ' | ')"
    # The error passed on is the one Get-Item gave, whose target is the log:
    # Fix names the log from it (Write-VmLogUnread).
    $cwdU = "ItemExistsUnauthorizedAccessError|$(Join-Path $vmDirU 'coworkd.log')|Access is denied"
    Assert-Case "Get-ClaudeVmLogCandidates throws for a VM log that could not be read, with the log as the error's target" (
        $res.Cands -eq $cwdU) "got $($res.Cands)"
    Assert-Case "Get-ClaudeVmLogAge throws for it too" ($res.Age -eq $cwdU) "got $($res.Age)"
    Assert-Case "and, read, the one log there is listed" ($res.OpenCands -eq 1) "got $($res.OpenCands)"
    # Named: the error Get-Item gives says only "Access is denied", and the
    # monitor's line was "Test-ClaudeCodeActive failed (Access is denied)"
    # until 6.3.1.
    Assert-Case "Test-ClaudeCodeTranscriptRecent throws when the projects folder could not be read, naming it" (
        $res.Projects -eq "THREW: The Claude Code transcript folder $projects could not be read: Access is denied") "got $($res.Projects)"
    Assert-Case "and, read, an old transcript is not recent" ($res.OpenRecent -eq 'False') "got $($res.OpenRecent)"

    # The walk of the projects folder: a folder in it that could not be
    # listed leaves the answer unknown, unless a recent transcript answers
    # it anyway. A folder that went away during the walk is a session that
    # ended.
    $res = & {
        . $walkGetChildItem
        $o = [ordered]@{}
        $script:WalkError = 'denied'
        try { $o.DeniedOld = [string](Test-ClaudeCodeTranscriptRecent -Minutes 35) } catch { $o.DeniedOld = "THREW: $(Get-ClaudeErrorText $_)" }
        $script:WalkError = 'gone'
        try { $o.GoneOld = [string](Test-ClaudeCodeTranscriptRecent -Minutes 35) } catch { $o.GoneOld = "THREW: $(Get-ClaudeErrorText $_)" }
        New-FixtureFile (Join-Path $projects 'p2\new.jsonl') 1
        $script:WalkError = 'denied'
        try { $o.DeniedNew = [string](Test-ClaudeCodeTranscriptRecent -Minutes 35) } catch { $o.DeniedNew = "THREW: $(Get-ClaudeErrorText $_)" }
        $script:WalkError = 'denied'
        $o
    }
    Assert-Case "a folder in the walk that could not be listed, no recent transcript: throws" (
        $res.DeniedOld -like "THREW: Access to the path*p-walk*is denied.") "got $($res.DeniedOld)"
    Assert-Case "a folder that went away during the walk: not recent" ($res.GoneOld -eq 'False') "got $($res.GoneOld)"
    Assert-Case "a recent transcript answers whatever else could not be listed" ($res.DeniedNew -eq 'True') "got $($res.DeniedNew)"

    # -----------------------------------------------------------------
    Write-Host ""
    Write-Host "--- Watch's Code-session guard and folder sync, with a folder that could not be read"

    # A folder that could not be read may hold a session, or be the one the
    # app moved to. Until 6.3.1 the guard passed it over as a folder with
    # nothing recent in it and answered "no session", and the sync followed
    # whichever candidate was left. The transcript written above is made old.
    New-FixtureFile (Join-Path $projects 'p2\new.jsonl') 600
    $codeSessions = Join-Path $dataRoot 'claude-code-sessions'
    New-FixtureFile (Join-Path $codeSessions 'acct\org\old.json') 600
    $logsU = Join-Path $L 'Claude\logs'
    $script:AppCli = @()
    $script:CheckErrorsSaid = @{}
    function Get-GuardLine {
        param([int]$From)
        return (@(@($script:WatchLogLines) | Select-Object -Skip $From |
                Where-Object { $_ -match 'assuming a Code session IS active' }) -join ' / ')
    }
    $res = & {
        . $denyGetItem
        $o = [ordered]@{}
        $script:DeniedPaths = @()
        $o.Open = [string](Test-ClaudeCodeActive)
        $script:DeniedPaths = @((Join-Path $L 'Claude-Data'))
        $n = @($script:WatchLogLines).Count
        $script:CheckErrorsSaid = @{}
        $o.Cand = [string](Test-ClaudeCodeActive)
        $o.CandLine = Get-GuardLine $n
        $script:DeniedPaths = @($codeSessions)
        $n = @($script:WatchLogLines).Count
        $script:CheckErrorsSaid = @{}
        $o.Dir = [string](Test-ClaudeCodeActive)
        $o.DirLine = Get-GuardLine $n
        $script:DeniedPaths = @()
        $o
    }
    Assert-Case "every folder read, every signal old: not active" ($res.Open -eq 'False') "got $($res.Open)"
    Assert-Case "a candidate folder that could not be read: active (fails closed), and said, naming it" (
        $res.Cand -eq 'True' -and
        $res.CandLine -eq ('Test-ClaudeCodeActive failed (no recent sign of a session, and what could not be read may hold one: ' +
            "$(Join-Path $L 'Claude-Data') could not be read: Access is denied); " +
            'assuming a Code session IS active; the same error is logged at most once an hour')) "got $($res.Cand); $($res.CandLine)"
    Assert-Case "a claude-code-sessions folder that could not be read: active, and said, naming it" (
        $res.Dir -eq 'True' -and $res.DirLine -like "*$codeSessions could not be read: Access is denied*") `
        "got $($res.Dir); $($res.DirLine)"
    $res = & {
        . $walkGetChildItem
        $o = [ordered]@{}
        $script:WalkError = 'denied'
        $n = @($script:WatchLogLines).Count
        $script:CheckErrorsSaid = @{}
        $o.Walk = [string](Test-ClaudeCodeActive)
        $o.WalkLine = Get-GuardLine $n
        $script:WalkError = 'gone'
        $n = @($script:WatchLogLines).Count
        $script:CheckErrorsSaid = @{}
        $o.Gone = [string](Test-ClaudeCodeActive)
        $o.GoneLines = @(@($script:WatchLogLines) | Select-Object -Skip $n | Where-Object { $_ -match 'could not be read|IS active' })
        # The candidates' own listings, of Claude-Data and of each logs
        # folder, through the same stand-in.
        $o.GoneCands = @((Get-ClaudeDataCandidates -PackageFamilyName 'Claude_testpub').Errors)
        $script:WalkError = 'denied'
        $o.DeniedCands = @((Get-ClaudeDataCandidates -PackageFamilyName 'Claude_testpub').Errors)
        $o
    }
    Assert-Case "a folder in the session walk that could not be listed: active, and said" (
        $res.Walk -eq 'True' -and $res.WalkLine -like '*p-walk*is denied.*') "got $($res.Walk); $($res.WalkLine)"
    Assert-Case "a folder that went away during the walk: not active, and no read said to have failed" (
        $res.Gone -eq 'False' -and @($res.GoneLines).Count -eq 0) "got $($res.Gone); $(@($res.GoneLines) -join ' | ')"
    Assert-Case "a candidate folder gone when listed is no error" (@($res.GoneCands).Count -eq 0) "errors: $(@($res.GoneCands) -join ' | ')"
    Assert-Case "one that could not be listed is, naming it" (
        @($res.DeniedCands | Where-Object { $_ -like "$(Join-Path $L 'Claude-Data') could not be read: Access to the path*p-walk*is denied." }).Count -eq 1) `
        "errors: $(@($res.DeniedCands) -join ' | ')"
    # The app's log folder, listed for its renderer and remote-control logs,
    # could not be: the only read that fails.
    $res = & {
        function Get-ChildItem {
            [CmdletBinding()]
            param([string]$LiteralPath, [string]$Filter, [switch]$Recurse, [switch]$File, [switch]$Directory, [switch]$Force)
            if ($LiteralPath -eq $logsU -and @('unknown-window*.log', 'claude.ai-web*.log', 'remote-control*.log') -contains $Filter) {
                $ex = New-Object System.UnauthorizedAccessException "Access to the path '$LiteralPath' is denied."
                $PSCmdlet.WriteError((New-Object System.Management.Automation.ErrorRecord($ex, 'DirUnauthorizedAccessError',
                    [System.Management.Automation.ErrorCategory]::PermissionDenied, $LiteralPath)))
                return
            }
            Microsoft.PowerShell.Management\Get-ChildItem -LiteralPath $LiteralPath -Filter $Filter -Recurse:$Recurse `
                -File:$File -Directory:$Directory -Force:$Force
        }
        $o = [ordered]@{}
        $n = @($script:WatchLogLines).Count
        $script:CheckErrorsSaid = @{}
        $o.Logs = [string](Test-ClaudeCodeActive)
        $o.LogsLine = Get-GuardLine $n
        $o
    }
    Assert-Case "the app's log folder that could not be listed: active, and said, naming it" (
        $res.Logs -eq 'True' -and $res.LogsLine -like "*$logsU could not be read: Access to the path*is denied.*") `
        "got $($res.Logs); $($res.LogsLine)"

    # The sync, started on the data root under Claude-Data, which then could
    # not be read. %APPDATA%\Claude, with old Preferences of its own, is the
    # candidate left.
    $ClaudeAppData     = Join-Path $R 'Claude'
    $StealthMarkerName = 'stealth-relaunch'
    $start = Resolve-ClaudeDataPaths -PackageFamilyName 'Claude_testpub' -AppDataDir $ClaudeAppData
    $script:ClaudeEnv = [ordered]@{
        PackageFamilyName = 'Claude_testpub'
        VmLogFile         = (Join-Path $vmDirU 'cowork_vm_node.log')
        VmLogCandidates   = @()
    }
    foreach ($k in @('DataRoot', 'DataRootSource', 'LogDir', 'LogDirSource', 'VmCachePath', 'BundlePath',
                     'SessionsDir', 'CodeSessionsDir', 'DataRootCandidates', 'LogDirCandidates')) {
        $script:ClaudeEnv[$k] = $start.$k
    }
    $script:ClaudeLogDir      = $start.LogDir
    $script:StealthMarkerPath = Get-StealthMarkerPath
    $script:NextDataPathSync  = (Get-Date).AddMinutes(10)
    $script:CheckErrorsSaid   = @{}
    $res = & {
        . $denyGetItem
        $o = [ordered]@{}
        $script:DeniedPaths = @((Join-Path $L 'Claude-Data'))
        $n = @($script:WatchLogLines).Count
        Sync-ClaudeDataPaths -Force
        $o.Root  = $script:ClaudeEnv.DataRoot
        $o.Lines = @(@($script:WatchLogLines) | Select-Object -Skip $n)
        $n = @($script:WatchLogLines).Count
        Sync-ClaudeDataPaths -Force
        $o.Again = @(@($script:WatchLogLines) | Select-Object -Skip $n).Count
        $script:DeniedPaths = @()
        $n = @($script:WatchLogLines).Count
        Sync-ClaudeDataPaths -Force
        $o.ReadRoot  = $script:ClaudeEnv.DataRoot
        $o.ReadLines = @(@($script:WatchLogLines) | Select-Object -Skip $n).Count
        $o
    }
    Assert-Case "(the sync starts on the data root under Claude-Data)" ($start.DataRoot -eq $dataRoot) "got $($start.DataRoot)"
    Assert-Case "a candidate that could not be read: nothing is moved" ($res.Root -eq $dataRoot) "got $($res.Root)"
    Assert-Case "and it is said once, naming the folder and the error" (
        @($res.Lines).Count -eq 1 -and
        $res.Lines[0] -eq ("Looking for the app's folders again failed (1 candidate folder(s) could not be read: " +
            "$(Join-Path $L 'Claude-Data') could not be read: Access is denied); " +
            'nothing was moved; the same error is logged at most once an hour')) "lines: $(@($res.Lines) -join ' | ')"
    Assert-Case "the same error again in the hour: not said again" ($res.Again -eq 0) "got $($res.Again)"
    Assert-Case "read again: it stays, and nothing is said" ($res.ReadRoot -eq $dataRoot -and $res.ReadLines -eq 0) `
        "got $($res.ReadRoot), $($res.ReadLines) line(s)"

    # The service, told apart by Get-ClaudeEnvironment: not registered is
    # $false and a note; could not be read is $null and a read error. The
    # recovery settings the same. Until 6.3.1 both failures read as "not
    # registered" and "not configured", and Fix's report said so. The rest
    # of the discovery reads this machine, read-only.
    Add-Type -AssemblyName 'Microsoft.PowerShell.Commands.Management'
    $res = & {
        . $denyGetItem
        # Not there: the measured identity (NoServiceFoundForGivenName,
        # ObjectNotFound, ServiceCommandException). Could not be read: any
        # other; this one is made up, since an elevated shell is not refused.
        function Get-Service {
            [CmdletBinding()]
            param([string]$Name)
            if ($script:SvcCase -eq 'missing') {
                $ex = New-Object Microsoft.PowerShell.Commands.ServiceCommandException "Cannot find any service with service name '$Name'."
                $PSCmdlet.WriteError((New-Object System.Management.Automation.ErrorRecord($ex, 'NoServiceFoundForGivenName',
                    [System.Management.Automation.ErrorCategory]::ObjectNotFound, $Name)))
                return
            }
            $ex = New-Object System.InvalidOperationException "Cannot open $Name service on computer '.'."
            $PSCmdlet.WriteError((New-Object System.Management.Automation.ErrorRecord($ex, 'CouldNotGetService',
                [System.Management.Automation.ErrorCategory]::PermissionDenied, $Name)))
        }
        $o = [ordered]@{}
        $script:SvcCase = 'missing'
        $script:DeniedPaths = @()
        $o.Missing = Get-ClaudeEnvironment -SkipCacheInventory
        $script:SvcCase = 'unread'
        $script:DeniedPaths = @('HKLM:\SYSTEM\CurrentControlSet\Services\CoworkVMService')
        $o.Unread = Get-ClaudeEnvironment -SkipCacheInventory
        $script:DeniedPaths = @()
        $o
    }
    Assert-Case "Get-ClaudeEnvironment: a service that is not registered is false, and noted" (
        $res.Missing.ServiceFound -is [bool] -and -not $res.Missing.ServiceFound -and
        @($res.Missing.Warnings | Where-Object { $_ -eq 'CoworkVMService is not registered' }).Count -eq 1 -and
        @($res.Missing.ReadErrors | Where-Object { $_ -like 'Service *' }).Count -eq 0) "found $($res.Missing.ServiceFound); read errors: $(@($res.Missing.ReadErrors) -join ' | ')"
    Assert-Case "Get-ClaudeEnvironment: a service that could not be read is null, and a read error" (
        $null -eq $res.Unread.ServiceFound -and
        @($res.Unread.Warnings | Where-Object { $_ -like '*is not registered' }).Count -eq 0 -and
        (Test-StartsWith $res.Unread.ReadErrors "Service CoworkVMService could not be read: Cannot open CoworkVMService service")) `
        "found $($res.Unread.ServiceFound); read errors: $(@($res.Unread.ReadErrors) -join ' | ')"
    Assert-Case "Get-ClaudeEnvironment: recovery settings that could not be read are null, and a read error" (
        $null -eq $res.Unread.ServiceRecoveryConfigured -and
        (Test-StartsWith $res.Unread.ReadErrors "The recovery settings of service CoworkVMService could not be read: Requested registry access is not allowed.")) `
        "recovery $($res.Unread.ServiceRecoveryConfigured); read errors: $(@($res.Unread.ReadErrors) -join ' | ')"

    # -----------------------------------------------------------------
    Write-Host ""
    Write-Host "--- Get-ClaudeEnvironment: a path that could not be read is not one with nothing there"

    # These reads went through Test-Path until 6.3.1. Test-Path answers $false
    # for nothing there and for a path it could not read alike (measured
    # 2026-11-23 with an explicit Deny ACE: it writes an error and answers
    # $false). So hcsdiag.exe that could not be read was "not present", the
    # package's Claude.exe that could not be read was passed over without a
    # word, a log folder or a cache folder that could not be read was passed
    # over as if it held nothing, and a folder named claude.exe was taken for
    # the app. Get-AppxPackage answers for a package in the fixture,
    # Get-Process and Get-CimInstance for a claude process whose image is a
    # fixture path, or for none, Resolve-ClaudeDataPaths names fixture
    # folders, and LOCALAPPDATA and ProgramFiles name empty fixture folders,
    # so the cases do not depend on what this machine has installed; the
    # service reads are this machine's, read-only. A denied path gives the
    # message Get-Item gives (measured 2026-11-23), which does not name the
    # path, so each line is checked for naming the path itself.
    Reset-Fixture 'envreads'
    $erBase    = Join-Path $fixtureRoot 'envreads'
    $erPf      = Join-Path $erBase 'ProgramFiles'
    $hcsPath   = Join-Path $env:SystemRoot 'System32\hcsdiag.exe'
    New-FixtureFile (Join-Path $erBase 'Install\app\claude.exe') 5
    New-FixtureFile (Join-Path $erBase 'Logs\main.log') 5
    New-FixtureFile (Join-Path $erBase 'Cache\rootfs.vhdx') 5
    New-FixtureFile (Join-Path $erBase 'FolderInstall\app\claude.exe') 5 -Directory
    New-FixtureFile (Join-Path $erBase 'Proc\claude.exe') 5
    New-FixtureFile (Join-Path $erBase 'Cim\claude.exe') 5
    New-FixtureFile (Join-Path $erBase 'FolderProc\claude.exe') 5 -Directory
    New-FixtureFile (Join-Path $erBase 'FolderCim\claude.exe') 5 -Directory
    New-Item -ItemType Directory -Path $erPf -Force | Out-Null
    $erLa      = Join-Path $erBase 'LocalAppData'
    New-Item -ItemType Directory -Path $erLa -Force | Out-Null
    $script:ErProcExe = $null
    $script:ErCimExe  = $null
    $savedPf   = $env:ProgramFiles
    $savedLa   = $env:LOCALAPPDATA
    $er = $null
    try {
        $env:ProgramFiles = $erPf
        $env:LOCALAPPDATA = $erLa
        $er = & {
            . $denyGetItem
            function Get-AppxPackage {
                [CmdletBinding()]
                param([string]$Name)
                $null = $Name
                return [pscustomobject]@{ Name = 'Claude'; PackageFullName = 'Claude_1.0.0.0_x64__testpub'
                    PackageFamilyName = 'Claude_testpub'; InstallLocation = $script:ErInstall; Version = '1.0.0.0' }
            }
            # A claude process whose image is $script:ErProcExe; with none
            # set, none running, the identity measured for Get-Process.
            function Get-Process {
                [CmdletBinding()]
                param([string]$Name)
                if ($script:ErProcExe) {
                    return [pscustomobject]@{ Name = $Name; Id = 4242; MainModule = [pscustomobject]@{ FileName = $script:ErProcExe } }
                }
                $ex = New-Object Microsoft.PowerShell.Commands.ProcessCommandException "Cannot find a process with the name `"$Name`". Verify the process name and call the cmdlet again."
                $PSCmdlet.WriteError((New-Object System.Management.Automation.ErrorRecord($ex, 'NoProcessFoundForGivenName',
                    [System.Management.Automation.ErrorCategory]::ObjectNotFound, $Name)))
            }
            # Win32_Process: a claude process whose image is $script:ErCimExe.
            # Anything else, or with none set: nothing matches, an empty
            # answer with no error, as measured.
            function Get-CimInstance {
                [CmdletBinding()]
                param([Parameter(Position = 0)][string]$ClassName, [string]$Filter)
                $null = $Filter
                if ($ClassName -eq 'Win32_Process' -and $script:ErCimExe) {
                    return [pscustomobject]@{ Name = 'claude.exe'; ProcessId = 4243; ExecutablePath = $script:ErCimExe }
                }
            }
            function Resolve-ClaudeDataPaths {
                param([string]$PackageFamilyName, [string]$AppDataDir)
                $null = $PackageFamilyName, $AppDataDir
                return [pscustomobject]@{
                    DataRoot = $null; DataRootSource = $null; LogDir = $script:ErLogs; LogDirSource = 'test'
                    VmCachePath = $script:ErCache; BundlePath = $null; SessionsDir = $null; CodeSessionsDir = $null
                    DataRootCandidates = @(); LogDirCandidates = @(); Notes = @(); Errors = @()
                }
            }
            $o = [ordered]@{ Unread = $null; UnreadError = ''; Folder = $null; FolderError = ''; NoLocal = $null; NoLocalError = '' }
            Set-StrictMode -Version Latest

            # 1. Each one there, and denied.
            $script:ErInstall = Join-Path $erBase 'Install'
            $script:ErLogs    = Join-Path $erBase 'Logs'
            $script:ErCache   = Join-Path $erBase 'Cache'
            $script:ErProcExe = Join-Path $erBase 'Proc\claude.exe'
            $script:ErCimExe  = Join-Path $erBase 'Cim\claude.exe'
            $script:DeniedPaths = @($hcsPath, (Join-Path $script:ErInstall 'app\claude.exe'), $script:ErProcExe, $script:ErCimExe,
                $script:ErLogs, $script:ErCache)
            try { $o.Unread = Get-ClaudeEnvironment } catch { $o.UnreadError = "$($_.Exception.Message)" }
            $script:DeniedPaths = @()

            # 2. The package's app\claude.exe and each process's image are
            # folders; the log and cache folders are not there.
            $script:ErProcExe = Join-Path $erBase 'FolderProc\claude.exe'
            $script:ErCimExe  = Join-Path $erBase 'FolderCim\claude.exe'
            $script:ErInstall = Join-Path $erBase 'FolderInstall'
            $script:ErLogs    = Join-Path $erBase 'NoLogs'
            $script:ErCache   = Join-Path $erBase 'NoCache'
            try { $o.Folder = Get-ClaudeEnvironment } catch { $o.FolderError = "$($_.Exception.Message)" }

            # 3. Nothing in the package, no process, and LOCALAPPDATA not set.
            $script:ErProcExe = $null
            $script:ErCimExe  = $null
            $script:ErInstall = Join-Path $erBase 'NoInstall'
            Remove-Item Env:\LOCALAPPDATA -ErrorAction SilentlyContinue
            try { $o.NoLocal = Get-ClaudeEnvironment -SkipCacheInventory } catch { $o.NoLocalError = "$($_.Exception.Message)" }
            $o
        }
    } finally {
        $script:DeniedPaths = @()
        $script:ErProcExe = $null
        $script:ErCimExe  = $null
        $env:ProgramFiles = $savedPf
        $env:LOCALAPPDATA = $savedLa
    }
    $u = $er.Unread
    $f = $er.Folder
    $n = $er.NoLocal
    $erInstallExe = Join-Path $erBase 'Install\app\claude.exe'
    Assert-Case "Get-ClaudeEnvironment: the denied run ran to the end" ($null -ne $u -and -not $er.UnreadError) "error: $($er.UnreadError)"
    Assert-Case "Get-ClaudeEnvironment: the folder run ran to the end" ($null -ne $f -and -not $er.FolderError) "error: $($er.FolderError)"
    if ($null -ne $u) {
        Assert-Case "hcsdiag.exe that could not be read is a read error, not 'not present'" (
            $null -eq $u.HcsDiagPath -and
            @($u.Warnings | Where-Object { $_ -like 'hcsdiag.exe not present*' }).Count -eq 0 -and
            @($u.ReadErrors) -contains "hcsdiag.exe ($hcsPath) could not be read: Access is denied") `
            "path $($u.HcsDiagPath); read errors: $(@($u.ReadErrors) -join ' | ')"
        $uExeLines = @($u.ReadErrors | Where-Object { $_ -like 'Claude executable candidate*' })
        Assert-Case "the package's Claude.exe that could not be read is not taken, and is said" (
            $null -eq $u.ClaudeExe -and
            @($u.Warnings | Where-Object { $_ -eq 'Claude executable not located' }).Count -eq 1 -and
            $uExeLines -contains "Claude executable candidate $erInstallExe could not be read: Access is denied") `
            "exe $($u.ClaudeExe) ($($u.ClaudeExeSource)); read errors: $(@($u.ReadErrors) -join ' | ')"
        Assert-Case "a running claude.exe that could not be read is said" (
            $uExeLines -contains "Claude executable candidate $(Join-Path $erBase 'Proc\claude.exe') could not be read: Access is denied") `
            "read errors: $(@($u.ReadErrors) -join ' | ')"
        Assert-Case "a claude.exe the process list names that could not be read is said" (
            $uExeLines -contains "Claude executable candidate $(Join-Path $erBase 'Cim\claude.exe') could not be read: Access is denied") `
            "read errors: $(@($u.ReadErrors) -join ' | ')"
        Assert-Case "each Claude.exe candidate that could not be read is said once" ($uExeLines.Count -eq 3) `
            "read errors: $(@($u.ReadErrors) -join ' | ')"
        Assert-Case "a log folder that could not be read is a read error, not 'no VM log'" (
            @($u.ReadErrors) -contains "The VM logs in $(Join-Path $erBase 'Logs') could not be read: Access is denied" -and
            @($u.Warnings | Where-Object { $_ -like 'No VM log present*' }).Count -eq 0) `
            "warnings: $(@($u.Warnings) -join ' | ')"
        Assert-Case "a cache folder that could not be read makes the size a lower bound, and is named" (
            @($u.Warnings) -contains ("Cache inventory: 1 item(s) could not be read, so CacheSizeBytes is a lower bound " +
                "(first: $(Join-Path $erBase 'Cache'): Access is denied)") -and
            $u.CacheSizeBytes -eq 0 -and @($u.VhdxFiles).Count -eq 0) `
            "size $($u.CacheSizeBytes); warnings: $(@($u.Warnings) -join ' | ')"
    }
    if ($null -ne $f) {
        Assert-Case "a folder named claude.exe, in the package or as a process's image, is not taken for the app" (
            $null -eq $f.ClaudeExe -and
            @($f.Warnings | Where-Object { $_ -eq 'Claude executable not located' }).Count -eq 1 -and
            @($f.ReadErrors | Where-Object { $_ -like 'Claude executable candidate*' -or $_ -like 'The process*' }).Count -eq 0) `
            "exe $($f.ClaudeExe) ($($f.ClaudeExeSource)); read errors: $(@($f.ReadErrors) -join ' | ')"
        Assert-Case "a log folder and a cache folder with nothing there are passed over without a word" (
            @($f.ReadErrors | Where-Object { $_ -like 'The VM logs in*' }).Count -eq 0 -and
            @($f.Warnings | Where-Object { $_ -like 'No VM log present*' -or $_ -like 'Cache inventory:*' }).Count -eq 0 -and
            $f.CacheSizeBytes -eq 0) `
            "size $($f.CacheSizeBytes); warnings: $(@($f.Warnings) -join ' | ')"
    }
    Assert-Case "Get-ClaudeEnvironment with LOCALAPPDATA not set runs, and locates nothing" (
        $null -ne $n -and -not $er.NoLocalError -and $null -eq $n.ClaudeExe -and
        @($n.Warnings | Where-Object { $_ -eq 'Claude executable not located' }).Count -eq 1) `
        "error: $($er.NoLocalError)"
    $fixtureCasesEnded = $true
} catch {
    Assert-Case "the fixture cases ran to the end" $false `
        "$($_.Exception.Message) (line $($_.InvocationInfo.ScriptLineNumber))"
} finally {
    $env:LOCALAPPDATA = $savedLocal
    $env:APPDATA      = $savedRoaming
    $env:USERPROFILE  = $savedProfile
    # A variable that is not set is the one way the two removals fail, and
    # not set is what they are for.
    if ($savedUdd) { $env:CLAUDE_USER_DATA_DIR = $savedUdd } else { Remove-Item Env:\CLAUDE_USER_DATA_DIR -ErrorAction SilentlyContinue }
    if ($savedCfg) { $env:CLAUDE_CONFIG_DIR = $savedCfg } else { Remove-Item Env:\CLAUDE_CONFIG_DIR -ErrorAction SilentlyContinue }
    # Clean-up of this script's own fixtures and nothing else: every case
    # is counted by now, and a folder left in %TEMP% changes nothing.
    $tempRoot = [System.IO.Path]::GetTempPath()
    if ((Split-Path $fixtureRoot -Leaf) -like 'claudefix-ci-*' -and
        $fixtureRoot.StartsWith($tempRoot, [System.StringComparison]::OrdinalIgnoreCase) -and
        (Test-Path -LiteralPath $fixtureRoot)) {
        Remove-Item -LiteralPath $fixtureRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------
# Prevent's Defender exclusion check, against the data root. The data root
# sits beside the app's own %LOCALAPPDATA%\Claude, and the old prefix test
# counted an exclusion for that folder as covering it. Measured 2026-11-21:
# a machine with exactly that exclusion, so Prevent never excluded the data
# root there.
# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Prevent: which Defender exclusions cover a path"
$preventPath = Join-Path $Root 'Prevent-ClaudeIssues.ps1'
# Until 6.3.1 a lift that failed lost its reason; each case below now says it.
$coversSrc = $null
$coversWhy = ''
try { $coversSrc = Get-FunctionSource $preventPath 'Test-DefenderExclusionCovers' } catch { $coversWhy = $_.Exception.Message }
Assert-Case "Test-DefenderExclusionCovers was found in Prevent" ($null -ne $coversSrc) $coversWhy
# The step asks Test-DefenderListHidden of each list before it reads it. It
# looks only at the list it is given, so it is lifted as it is, not faked.
$hiddenSrc = $null
$hiddenWhy = ''
try { $hiddenSrc = Get-FunctionSource $preventPath 'Test-DefenderListHidden' } catch { $hiddenWhy = $_.Exception.Message }
Assert-Case "Test-DefenderListHidden was found in Prevent" ($null -ne $hiddenSrc) $hiddenWhy
if ($hiddenSrc) { . ([scriptblock]::Create($hiddenSrc)) }
$preventText = [System.IO.File]::ReadAllText($preventPath)
Assert-Case "the needed-paths loop asks Test-DefenderExclusionCovers" `
    ($preventText -match '(?s)foreach \(\$np in \$neededPaths\) \{\s*if \(-not \(Test-DefenderExclusionCovers -Path \$np -Exclusions \$exclusions\)\)')
Assert-Case "the data root is the first needed path" `
    ($preventText -match '(?s)\$neededPaths = @\(\s*\$ClaudeDataRoot,')
$U  = 'C:\Users\u\AppData\Local'
$dr = "$U\Claude-Data\claude_abc123"
if ($coversSrc) {
    . ([scriptblock]::Create($coversSrc))
    $coverCases = @(
        @{ Name = "the app's own %LOCALAPPDATA%\Claude does not cover the data root beside it"; Path = $dr; Ex = @("$U\Claude"); Want = $false }
        @{ Name = "nor with a trailing backslash";                       Path = $dr; Ex = @("$U\Claude\");              Want = $false }
        @{ Name = "Claude-Data covers it";                               Path = $dr; Ex = @("$U\Claude-Data");          Want = $true }
        @{ Name = "%LOCALAPPDATA% covers it";                            Path = $dr; Ex = @($U);                        Want = $true }
        @{ Name = "the data root itself covers it";                      Path = $dr; Ex = @($dr);                       Want = $true }
        @{ Name = "with a trailing backslash too";                       Path = $dr; Ex = @("$dr\");                    Want = $true }
        @{ Name = "case does not matter";                                Path = $dr; Ex = @($dr.ToLowerInvariant());    Want = $true }
        @{ Name = "a drive root covers it";                              Path = $dr; Ex = @('C:\');                     Want = $true }
        @{ Name = "a folder whose name only starts the same does not";   Path = $dr; Ex = @("$U\Claude-Data\claude_abc"); Want = $false }
        @{ Name = "an empty entry covers nothing";                       Path = $dr; Ex = @('', '  ');                  Want = $false }
        @{ Name = "the N/A placeholder covers nothing";                  Path = $dr; Ex = @('N/A: Must be an administrator to view exclusions'); Want = $false }
        @{ Name = "no exclusions at all";                                Path = $dr; Ex = $null;                        Want = $false }
        @{ Name = "the right entry is found among others";               Path = $dr; Ex = @("$U\Claude", 'D:\Games', "$U\Claude-Data"); Want = $true }
        @{ Name = "a wildcard entry covers what it matches";             Path = 'C:\Program Files\WindowsApps\Claude_1.0.0.0_x64__x\app\claude.exe'
           Ex = @('C:\Program Files\WindowsApps\Claude_*'); Want = $true }
        @{ Name = "an environment variable in an entry is expanded";     Path = "$env:SystemRoot\System32\vmwp.exe"
           Ex = @('%SystemRoot%\System32\vmwp.exe'); Want = $true }
        @{ Name = "brackets in an entry are literal";                    Path = 'C:\Games [x]\data'; Ex = @('C:\Games [x]'); Want = $true }
        @{ Name = "and are not a character class";                       Path = 'C:\Games x\data';   Ex = @('C:\Games [x]'); Want = $false }
        @{ Name = "an empty path is never covered";                      Path = '';  Ex = @('C:\');                     Want = $false }
    )
    foreach ($cc in $coverCases) {
        $got = Test-DefenderExclusionCovers -Path $cc.Path -Exclusions $cc.Ex
        Assert-Case $cc.Name ($got -eq $cc.Want) "got $got"
    }
}

# Then the step itself: its try block, lifted out of Prevent and run with
# Get-MpPreference and Add-MpPreference faked, so nothing real is read or
# changed. Since 6.3.0 every exclusion is written to the undo record before it
# is added, and taken out again when adding it fails, so -Undo removes the
# ones Prevent added and no others; the fakes below record both.
Write-Host ""
Write-Host "--- Prevent: the Defender exclusion step"
$preventAst = [System.Management.Automation.Language.Parser]::ParseFile($preventPath, [ref]$null, [ref]$null)
$defTry = @($preventAst.FindAll({ param($n)
    $n -is [System.Management.Automation.Language.TryStatementAst] -and
    $n.Body.Extent.Text -match '\$neededPaths = @\(' }, $true) |
    Sort-Object { $_.Extent.EndOffset - $_.Extent.StartOffset } | Select-Object -First 1)
Assert-Case "the Defender exclusion step was found in Prevent" ($defTry.Count -eq 1)
if ($defTry.Count -eq 1) {
    $stepText = $defTry[0].Extent.Text
    $hyperV = Join-Path $env:ProgramFiles 'Hyper-V'
    $vmwp   = "$env:SystemRoot\System32\vmwp.exe"
    $vmms   = "$env:SystemRoot\System32\vmms.exe"
    $allProcs = @('vmwp.exe', 'vmms.exe', 'vmcompute.exe', 'CoworkVMService.exe')
    $runStep = {
        param([string[]]$Paths, [string[]]$Procs, [string[]]$FailOn = @(), [switch]$ReadFails)
        & {
            $state = @{ Paths = [System.Collections.ArrayList]@($Paths); Procs = [System.Collections.ArrayList]@($Procs)
                        AddedPaths = @(); AddedProcs = @(); Logged = @()
                        Records = (New-Object System.Collections.ArrayList); Statuses = @() }
            function Get-MpPreference {
                param($ErrorAction)
                $null = $ErrorAction
                if ($ReadFails) { throw "Get-MpPreference: access denied" }
                [pscustomobject]@{ ExclusionPath = @($state.Paths); ExclusionProcess = @($state.Procs) }
            }
            function Add-MpPreference {
                param([string[]]$ExclusionPath, [string[]]$ExclusionProcess, $ErrorAction)
                $null = $ErrorAction
                foreach ($x in @($ExclusionPath) + @($ExclusionProcess)) {
                    if ($x -and $FailOn -contains $x) { throw "Add-MpPreference: refused $x" }
                }
                foreach ($x in @($ExclusionPath))    { if ($x) { $state.AddedPaths += $x; $null = $state.Paths.Add($x) } }
                foreach ($x in @($ExclusionProcess)) { if ($x) { $state.AddedProcs += $x; $null = $state.Procs.Add($x) } }
            }
            function Log { param([string]$Message, [string]$Colour, [switch]$Indent) $null = $Colour, $Indent; $state.Logged += $Message }
            function Add-PreventUndoItem {
                param([string]$Kind, [string]$Key, $Original)
                $null = $Original
                $null = $state.Records.Add("${Kind}|${Key}")
            }
            function Clear-PreventUndoItem {
                param([string]$Kind, [string]$Key)
                $state.Records.Remove("${Kind}|${Key}")
            }
            function Add-StepResult {
                param([string]$Status, [string]$Detail, [switch]$Reboot)
                $null = $Detail, $Reboot
                $state.Statuses += $Status
            }
            $ClaudeDataRoot = $dr
            $ServiceExe     = 'CoworkVMService'
            . ([scriptblock]::Create($stepText))
            [pscustomobject]@{ AddedPaths = ($state.AddedPaths -join ','); AddedProcs = ($state.AddedProcs -join ',')
                               Logged = ($state.Logged -join ' | '); Records = (@($state.Records) -join ',')
                               Statuses = (@($state.Statuses | Sort-Object -Unique) -join ',') }
        }
    }

    # This machine on 2026-11-21: the app's own folder excluded, the data
    # root beside it not.
    $r = & $runStep @("$U\Claude", $hyperV, $vmwp, $vmms) $allProcs
    Assert-Case "with only %LOCALAPPDATA%\Claude excluded, the data root is added" ($r.AddedPaths -eq $dr) "added $($r.AddedPaths)"
    Assert-Case "and nothing else" ($r.AddedProcs -eq '') "added processes $($r.AddedProcs)"
    Assert-Case "and it is in the undo record" ($r.Records -eq "DefenderPath|$dr") "recorded $($r.Records)"
    Assert-Case "and the step reports it done" ($r.Statuses -eq 'Done,Unchanged') "statuses $($r.Statuses)"

    $r = & $runStep @("$U\Claude-Data", $hyperV, $vmwp, $vmms) $allProcs
    Assert-Case "with everything excluded, nothing is added" ($r.AddedPaths -eq '' -and $r.AddedProcs -eq '') `
        "added $($r.AddedPaths) / $($r.AddedProcs)"
    Assert-Case "and it says so" ($r.Logged -match 'path exclusions already configured' -and
                                  $r.Logged -match 'process exclusions already configured') "logged: $($r.Logged)"
    Assert-Case "and records nothing, so -Undo leaves exclusions it did not add" ($r.Records -eq '') "recorded $($r.Records)"
    Assert-Case "and reports nothing changed" ($r.Statuses -eq 'Unchanged') "statuses $($r.Statuses)"

    $r = & $runStep @("$U\Claude-Data", $hyperV, $vmwp, $vmms) @('vmwp.exe')
    Assert-Case "with every path excluded, missing process exclusions are still added" `
        ($r.AddedPaths -eq '' -and $r.AddedProcs -eq 'vmms.exe,vmcompute.exe,CoworkVMService.exe') "added $($r.AddedPaths) / $($r.AddedProcs)"
    Assert-Case "and verified" ($r.Logged -match 'All process exclusions verified') "logged: $($r.Logged)"
    Assert-Case "and those three are in the undo record" (
        $r.Records -eq 'DefenderProcess|vmms.exe,DefenderProcess|vmcompute.exe,DefenderProcess|CoworkVMService.exe') "recorded $($r.Records)"

    $r = & $runStep @() @()
    Assert-Case "with no exclusions, all four paths and four processes are added" (
        $r.AddedPaths -eq (@($dr, $hyperV, $vmwp, $vmms) -join ',') -and
        $r.AddedProcs -eq ($allProcs -join ',')) "added $($r.AddedPaths) / $($r.AddedProcs)"
    Assert-Case "and all eight are in the undo record" (@($r.Records -split ',').Count -eq 8) "recorded $($r.Records)"

    # An exclusion Defender refuses is taken back out of the record, so -Undo
    # does not try to remove what was never added, and the step fails.
    $r = & $runStep @() @() @($vmwp, 'vmcompute.exe')
    Assert-Case "a refused path is not added" ($r.AddedPaths -eq (@($dr, $hyperV, $vmms) -join ',')) "added $($r.AddedPaths)"
    Assert-Case "a refused path and process are not left in the undo record" (
        $r.Records -notmatch [regex]::Escape($vmwp) -and $r.Records -notmatch 'vmcompute' -and
        @($r.Records -split ',').Count -eq 6) "recorded $($r.Records)"
    Assert-Case "and the step fails" ($r.Statuses -match 'Failed') "statuses $($r.Statuses)"

    # Until 6.3.0 Get-MpPreference was read with SilentlyContinue, so a Defender
    # that could not be read looked like one with nothing excluded.
    $r = & $runStep @() @() -ReadFails
    Assert-Case "a Defender that cannot be read adds nothing" ($r.AddedPaths -eq '' -and $r.AddedProcs -eq '') "added $($r.AddedPaths) / $($r.AddedProcs)"
    Assert-Case "records nothing" ($r.Records -eq '') "recorded $($r.Records)"
    Assert-Case "and fails the step" ($r.Statuses -eq 'Failed') "statuses $($r.Statuses)"

    # A list Defender hides from administrators holds only its "N/A"
    # placeholder, which says nothing of what is excluded.
    $na = 'N/A: Must be an administrator to view exclusions'
    $r = & $runStep @($na) @($na)
    Assert-Case "lists Defender hides: nothing is added" ($r.AddedPaths -eq '' -and $r.AddedProcs -eq '') "added $($r.AddedPaths) / $($r.AddedProcs)"
    Assert-Case "nothing is recorded, so -Undo removes nothing" ($r.Records -eq '') "recorded $($r.Records)"
    Assert-Case "and the step warns, once for each list" ($r.Statuses -eq 'Unchanged,Warning' -and
        $r.Logged -match 'which path exclusions exist; none were added' -and
        $r.Logged -match 'which process exclusions exist; none were added') "statuses $($r.Statuses); logged: $($r.Logged)"
}

# A case that ended the try early without an error (a return, a break) is
# caught here.
Assert-Case "the fixture cases ran to the end, with no case skipped" $fixtureCasesEnded

Write-Host ""
Write-Host ("Data discovery: {0} passed, {1} failed" -f $pass, $fail)
if ($fail -gt 0) {
    Write-Host "Data discovery gate FAILED" -ForegroundColor Red
    exit 1
}
Write-Host "Data discovery gate PASSED" -ForegroundColor Green
exit 0
