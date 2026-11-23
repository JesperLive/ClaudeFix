<#
.SYNOPSIS
    Claude Desktop / Cowork, Health Monitor

.DESCRIPTION
    Persistent background monitor that detects VirtioFS/Plan9 mount failures
    and HCS/vmcompute errors in Claude Desktop and automatically runs the fix
    script.

    Monitors:
    - [error] entries in Claude's VM and main logs that carry a VirtioFS mount
      or HCS failure signature
    - CoworkVMService status (stopped while Claude is running)
    - Windows Event Log for service, Hyper-V, and HCS Compute errors/warnings
    - VM log staleness (hung VM detection)
    - vmcompute handle leak detection (HCS service health)
    - Host clock drift (NTP/time sync). A clock check that fails part-way,
      as when w32tm cannot be started or answers with an error and no
      offset, is logged as a TIME WARNING naming the error, and a
      measurement that does not answer within 15s as skipped, each at most
      once an hour for the same error (v6.3.0+)
    - A check or a read that fails, such as an Event Log read, the free
      space of the VM's volume or the list of VM worker processes, is
      logged naming the check and the error, at most once an hour for the
      same error and for 3 errors of one check (v6.3.0+)

    SAFETY FEATURES (v3.3+):
    - Auto-fix is BLOCKED when user is active (Claude in focus, recent input, VM busy, CPU active),
      and, before the 30-second warning, for a Claude Code session the app started or a Code
      transcript written in the last 35 minutes, which Fix's own gate holds a run back for (v6.3.0+)
    - Claude Code session awareness (detects active Code sessions via session files, app logs + CLI transcripts)
    - Cowork session awareness (detects active Cowork VM processes via log parsing)
    - Claude startup grace period (defers auto-fix for 120s after Claude launch)
    - Extended CPU sampling (3 samples over 3s to catch bursty API-wait patterns)
    - Electron-aware window detection (GetWindowThreadProcessId, not MainWindowHandle)
    - Session 0 safe (falls back to process heuristics when Win32 APIs are unavailable)
    - Startup grace period (skips first 180s to avoid matching pre-existing events)
    - Consecutive-check requirements on all heuristic triggers
    - Event log matching tightened to Claude-specific messages only
    - Cooldown between fixes (default 5 min)
    - A repair for a log error or an Event Log error that Fix held back (exit
      3), that was BLOCKED or DEFERRED, or that fired during the wait a hold
      leaves, is tried again once the hold has cleared: with the 30-second
      warning at most 3 times, for up to 4 hours, and not after an app
      restart or a Fix run that started the repair. A Fix run that stopped
      before the repair, such as a Diagnostic run, a Cancel or a No at
      "Proceed?", does not end it, and while a Fix run holds Fix's lock the
      repair waits for it to end. A log error or Event Log error that fires
      while a repair is pending takes it over, and its 4 hours start again
      from then (v6.3.0+)
    - Each run Fix holds back in a row doubles the wait before the next
      repair: 2, 4, 8, 16, 32, then 60 minutes, until a run Fix does not hold
      back, a Fix run by hand that started the repair, which also ends the
      wait 5 minutes after it at the latest, or a hold 4 hours or more after
      the one before. BLOCKED and COOLDOWN lines for the same
      reason, and SKIPPED lines while a Fix run holds Fix's lock, are written
      at most once every 10 minutes; with a repair pending, SKIPPED is
      written once for each such run (v6.3.0+)
    - VM log window extended to 120s (covers Code thinking phases)

    When a failure is detected AND the user is idle, automatically runs
    Fix-ClaudeDesktop.ps1 to reset the VM and relaunch Claude.

    Designed to run as a hidden scheduled task (installed by Prevent-ClaudeIssues.ps1)
    but can also be started manually for foreground monitoring.

    STEALTH-UPDATE RELAUNCH (v6.1.0+):
    Claude Desktop applies background updates by quitting itself once idle and
    handing off to the Electron updater. On the MSIX path that handoff calls the
    AppX deployment API, which registers the new package but does not start the
    app afterwards, so the app stays down until a human starts it. Before
    quitting it writes a marker file recording window position and navigation
    history, and it restores from that marker if the next launch happens within
    five minutes. This monitor reads that marker and performs the launch the
    update path never does. The marker exists if and only if the app went down
    through the stealth path: a deliberate user quit writes nothing, so an app
    the user closed on purpose still stays closed.

.PARAMETER PollInterval
    Seconds between health checks (default: 30).

.PARAMETER Cooldown
    Minutes to wait between auto-fix runs (default: 5).

.PARAMETER NoStealthRelaunch
    Disable the stealth-update relaunch described above. The feature is ON by
    default; this switch turns it off. Nothing else changes: a missing marker
    is still reported and still not acted on.

.PARAMETER MaxStealthMarkerAgeMinutes
    Refuse to relaunch from a stealth marker older than this (default: 10).
    Deliberately wider than the app's own five-minute restore window, so a poll
    that lands just after that window still brings the app back, just without
    the window and tab restore. A marker much older than that means the machine
    was off or the user chose to stay closed, and neither is ours to overrule.

.PARAMETER Quiet
    Suppress console output (for scheduled task use).

.NOTES
    Version : 6.3.2
    Author  : Jesper Driessen
    Licence : MIT
#>

[CmdletBinding()]
param(
    [int]$PollInterval = 30,
    [int]$Cooldown = 5,
    [switch]$NoStealthRelaunch,
    [int]$MaxStealthMarkerAgeMinutes = 10,
    [switch]$Quiet
)

Set-StrictMode -Version Latest
# =====================================================================
# region ClaudeEnv
#
# Runtime discovery. This block is byte-identical in Fix-ClaudeDesktop.ps1,
# Watch-ClaudeHealth.ps1 and Prevent-ClaudeIssues.ps1, and CI fails the build
# if they drift. The canonical copy lives at .ci/ClaudeEnv.region.ps1.
#
# Do not hand-edit a single copy. Edit the canonical file, then resync all
# three. Three hand-maintained copies is exactly how findings 15, 28, 32 and
# 53 happened.
# =====================================================================
#region ClaudeEnv
# -- Reads that tell "not there" from "could not read" ----------------------
# Until 6.3.1 most reads in this toolkit passed -ErrorAction SilentlyContinue,
# so a read that failed (access denied, a broken WMI, a service manager that
# did not answer) came back empty and was taken as "nothing there": no such
# service, no Claude process running. The helpers below answer "not there"
# only for the one error that says so, told by its identity, and throw any
# other error unchanged, so the caller can say what it could not read and
# hold back what that read would have justified.

function Test-ClaudeNotFoundError {
    <#
    .SYNOPSIS
        Whether an error says the thing asked for is not there, rather than
        that it could not be read.
    .DESCRIPTION
        Told by the error's identity, never by its message, which Windows
        translates into the display language. All three must match: the
        error id, the category, which must be ObjectNotFound, and the full
        name of the exception type.

        The error id is the FullyQualifiedErrorId up to its first comma.
        PowerShell adds the name of the command that wrote the error after
        the comma, and a test's stand-in for a cmdlet has a different name
        from the cmdlet: measured 2026-11-22, a function named Get-FakeA that
        writes id NoServiceFoundForGivenName arrives as
        "NoServiceFoundForGivenName,Get-FakeA".
    .PARAMETER ErrorId
        The ids that mean "not there" for the command that was run.
    .PARAMETER ExceptionType
        The full name of the exception type that command raises for it.
    #>
    param(
        [System.Management.Automation.ErrorRecord]$ErrorRecord,
        [string[]]$ErrorId,
        [string]$ExceptionType
    )
    if ($null -eq $ErrorRecord -or $null -eq $ErrorRecord.Exception) { return $false }
    $id = @("$($ErrorRecord.FullyQualifiedErrorId)" -split ',')[0]
    return (($ErrorId -contains $id) -and
            $ErrorRecord.CategoryInfo.Category -eq [System.Management.Automation.ErrorCategory]::ObjectNotFound -and
            $ErrorRecord.Exception.GetType().FullName -eq $ExceptionType)
}

function Get-ClaudeErrorText {
    <#
    .SYNOPSIS
        An error's message on one line, or the name of its type when it has
        no message, for a line that says what could not be read and why.
    #>
    param($ErrorRecord)
    $ex = $ErrorRecord
    if ($ErrorRecord -is [System.Management.Automation.ErrorRecord]) { $ex = $ErrorRecord.Exception }
    if ($null -eq $ex) { return 'unknown error' }
    $text = ([string]$ex.Message -replace '\s*[\r\n]+\s*', ' ').Trim()
    if (-not $text) { $text = $ex.GetType().FullName }
    return $text
}

function Get-ClaudeServiceOrNull {
    <#
    .SYNOPSIS
        The service of that name, $null when no service has that name, and
        throws, with the error Get-Service gave, when it could not be read.
    .DESCRIPTION
        Measured 2026-11-22 on Windows 11, Windows PowerShell 5.1, with a
        read-only probe (Get-Service -Name NoSuchSvc_cf): error id
        NoServiceFoundForGivenName, category ObjectNotFound, exception
        Microsoft.PowerShell.Commands.ServiceCommandException. Any other
        error means the service could not be read.
    #>
    param([Parameter(Mandatory)][string]$Name)
    try {
        return (Get-Service -Name $Name -ErrorAction Stop)
    } catch {
        if (Test-ClaudeNotFoundError -ErrorRecord $_ -ErrorId 'NoServiceFoundForGivenName' `
                -ExceptionType 'Microsoft.PowerShell.Commands.ServiceCommandException') { return $null }
        throw
    }
}

function Get-ClaudeProcessByName {
    <#
    .SYNOPSIS
        Every running process of that name, nothing when none runs, and
        throws, with the error Get-Process gave, when the list could not be
        read.
    .DESCRIPTION
        Measured 2026-11-22 on Windows 11, Windows PowerShell 5.1, with a
        read-only probe (Get-Process -Name NoSuchProc_cf): error id
        NoProcessFoundForGivenName, category ObjectNotFound, exception
        Microsoft.PowerShell.Commands.ProcessCommandException. Any other
        error means the list could not be read. Wrap a call in @() for an
        array: none running outputs nothing.
    #>
    param([Parameter(Mandatory)][string]$Name)
    try {
        return @(Get-Process -Name $Name -ErrorAction Stop)
    } catch {
        if (Test-ClaudeNotFoundError -ErrorRecord $_ -ErrorId 'NoProcessFoundForGivenName' `
                -ExceptionType 'Microsoft.PowerShell.Commands.ProcessCommandException') { return }
        throw
    }
}

function Get-ClaudeScheduledTaskOrNull {
    <#
    .SYNOPSIS
        The scheduled task of that name in that folder, $null when there is
        no such task, and throws, with the error Get-ScheduledTask gave, when
        it could not be read.
    .DESCRIPTION
        Measured 2026-11-22 on Windows 11, Windows PowerShell 5.1, with a
        read-only probe (Get-ScheduledTask -TaskName NoSuch_cf -TaskPath
        \NoSuch_cf\, and again with -TaskPath \): error id
        CmdletizationQuery_NotFound, category ObjectNotFound, exception
        Microsoft.PowerShell.Cmdletization.Cim.CimJobException, whether or
        not the folder exists. Without -TaskPath the id is
        CmdletizationQuery_NotFound_TaskName, which is why the path is
        required here. Any other error means the task could not be read.
    .PARAMETER TaskPath
        The task's folder, with a backslash at each end: \Claude\.
    #>
    param(
        [Parameter(Mandatory)][string]$TaskName,
        [Parameter(Mandatory)][string]$TaskPath
    )
    try {
        return (Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction Stop)
    } catch {
        if (Test-ClaudeNotFoundError -ErrorRecord $_ -ErrorId 'CmdletizationQuery_NotFound' `
                -ExceptionType 'Microsoft.PowerShell.Cmdletization.Cim.CimJobException') { return $null }
        throw
    }
}

function Get-ClaudeItemOrNull {
    <#
    .SYNOPSIS
        The file, folder or registry key at that path, $null when there is
        nothing there, and throws, with the error Get-Item gave, when it
        could not be read.
    .DESCRIPTION
        Measured 2026-11-22 on Windows 11, Windows PowerShell 5.1, with
        read-only probes of a missing file, a file under a missing folder
        and a missing registry key: error id PathNotFound, category
        ObjectNotFound, exception
        System.Management.Automation.ItemNotFoundException, all three. Any
        other error means the path could not be read. Hidden items count
        as there (-Force).

        Test-Path cannot be used for this: it answers $false for a path it
        could not read, which is the confusion this exists to end.
    #>
    param([Parameter(Mandatory)][string]$LiteralPath)
    try {
        return (Get-Item -LiteralPath $LiteralPath -Force -ErrorAction Stop)
    } catch {
        if (Test-ClaudeNotFoundError -ErrorRecord $_ -ErrorId 'PathNotFound' `
                -ExceptionType 'System.Management.Automation.ItemNotFoundException') { return $null }
        throw
    }
}

function Get-ClaudeEnvironment {
    <#
    .SYNOPSIS
        Discovers where this machine's Claude Desktop install actually lives.
    .DESCRIPTION
        Every location the toolkit acts on is queried at run time rather than
        guessed at authoring time. Nothing here throws. Anything that cannot be
        determined stays $null and is recorded in Warnings, so a partial result
        is still usable by a caller that only needs some of the fields.

        A read that failed, as opposed to one that found nothing, is also
        listed in ReadErrors, one line each naming the read and the error,
        for the scripts to show: Warnings holds notes about the choices made
        as well. ServiceFound and ServiceRecoveryConfigured are $null when
        they could not be read. Until 6.3.1 a failed read of the package, the
        processes, the service or its recovery settings looked the same as
        finding nothing.

        Locations are discoverable and are discovered. Thresholds and log
        signatures are NOT in scope here: those are not location problems and
        cannot be fixed by looking harder at the filesystem.
    .PARAMETER SkipCacheInventory
        Skip the recursive walk of the VM cache. That walk is the expensive half
        of discovery and only the backup and purge paths consume its results, so
        anything that polls should pass this. CacheSizeBytes is left at -1 when
        skipped, which is distinguishable from a genuine zero, and is -1 as
        well when the walk failed.
    #>
    param([switch]$SkipCacheInventory)

    # Version 2 (6.2.0): the app's folders come from Resolve-ClaudeDataPaths
    # instead of %APPDATA%\Claude. The version is written into claude-env.json
    # so that a person reading the record knows which fields it has; nothing in
    # this toolkit reads the record back, since every run discovers afresh.
    $info = [ordered]@{
        SchemaVersion = 2
        DiscoveredUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')

        PackageFound      = $false
        PackageFullName   = $null
        PackageFamilyName = $null
        PackageVersion    = $null
        InstallLocation   = $null
        ApplicationId     = $null
        Aumid             = $null
        IsMsix            = $false

        ClaudeExe       = $null
        ClaudeExeSource = $null

        ServiceName               = 'CoworkVMService'
        ServiceFound              = $false
        ServiceStartMode          = $null
        ServiceBinaryPath         = $null
        ServiceProcessName        = 'cowork-svc'
        ServiceRecoveryConfigured = $false

        ProcessName = 'claude'

        # AppDataDir is this toolkit's own folder: fix-logs, watch-logs,
        # watch-state.json, the generated launchers. The app kept its data
        # there too until 6.18355; DataRoot is where it keeps it now.
        AppDataDir         = $null
        DataRoot           = $null
        DataRootSource     = $null
        LogDir             = $null
        LogDirSource       = $null
        VmCachePath        = $null
        BundlePath         = $null
        SessionsDir        = $null
        CodeSessionsDir    = $null
        DataRootCandidates = @()
        LogDirCandidates   = @()

        VmLogFile       = $null
        VmLogCandidates = @()

        VhdxFiles      = @()
        CacheSizeBytes = 0

        HcsDiagPath    = $null
        PsVersion      = $PSVersionTable.PSVersion.ToString()
        HasTickCount64 = $false
        IsAdmin        = $false

        Warnings   = @()
        ReadErrors = @()
    }

    # A read that failed: said in Warnings and in ReadErrors.
    $readFailed = {
        param([string]$What, $ErrorRecord)
        $line = "$What could not be read: $(Get-ClaudeErrorText $ErrorRecord)"
        $info.Warnings   += $line
        $info.ReadErrors += $line
    }

    # -- Host capabilities ------------------------------------------------
    try {
        $info.IsAdmin = ([Security.Principal.WindowsPrincipal] `
            [Security.Principal.WindowsIdentity]::GetCurrent()
            ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { & $readFailed 'Whether this shell is elevated' $_ }

    # TickCount64 is .NET Core and up. Windows PowerShell 5.1 runs on .NET
    # Framework and does not have it, which is why the idle-time maths has to
    # survive the 24.9 day TickCount wrap instead of just using the 64-bit call.
    # A probe: the error is the answer "not there".
    try { $null = [Environment]::TickCount64; $info.HasTickCount64 = $true }
    catch { $info.HasTickCount64 = $false }

    # hcsdiag.exe ships with Windows. Not there is said as not present, and
    # one that could not be read is a read error (Get-ClaudeItemOrNull); HCS
    # cleanup is unavailable either way, since a run of it would fail the same
    # way. Until 6.3.1 Test-Path said "not present" for both. The outer catch
    # is for a SystemRoot that is not set, and says so.
    try {
        $hcs = Join-Path $env:SystemRoot "System32\hcsdiag.exe"
        try {
            if (Get-ClaudeItemOrNull -LiteralPath $hcs) { $info.HcsDiagPath = $hcs }
            else { $info.Warnings += "hcsdiag.exe not present; HCS cleanup unavailable" }
        } catch { & $readFailed "hcsdiag.exe ($hcs)" $_ }
    } catch { $info.Warnings += "hcsdiag probe failed: $($_.Exception.Message)" }

    # -- This toolkit's own folder -----------------------------------------
    # %APPDATA%\Claude, where the toolkit keeps its logs, state and launchers.
    # Where the app keeps its own data and logs is found further down, once the
    # package is known; see Resolve-ClaudeDataPaths. The old
    # %ProgramData%\Claude tree is gone, so it is not probed at all.
    try {
        if ($env:APPDATA) {
            $info.AppDataDir = Join-Path $env:APPDATA "Claude"
        } else {
            $info.Warnings += "APPDATA is not set; per-user paths undiscoverable"
        }
    } catch { $info.Warnings += "AppData probe failed: $($_.Exception.Message)" }

    # -- MSIX package -------------------------------------------------------
    # No match is an empty answer with no error (measured 2026-11-22 with
    # Get-AppxPackage -Name NoSuch_cf*), so any error is a failed read. Until
    # 6.3.1 it was passed over and read as "no package installed".
    $pkgs = @()
    try {
        $pkgs = @(Get-AppxPackage -Name "Claude*" -ErrorAction Stop |
                  Where-Object { $_.Name -like "Claude*" })
        if ($pkgs.Count -gt 1) {
            $info.Warnings += "$($pkgs.Count) Claude packages registered; duplicates make the service start fail with error 87"
            $pkgs = @($pkgs | Sort-Object -Property @{ Expression = {
                try { [version]"$($_.Version)" } catch { [version]"0.0.0.0" }
            }} -Descending)
        }
        if ($pkgs.Count -gt 0) {
            $p = $pkgs[0]
            $info.PackageFound     = $true
            $info.IsMsix           = $true
            $info.PackageFullName  = "$($p.PackageFullName)"
            $info.PackageFamilyName = "$($p.PackageFamilyName)"
            $info.InstallLocation  = "$($p.InstallLocation)"
            # Get-AppxPackage hands back Version as a String on Windows
            # PowerShell 5.1, so the PackageVersion struct accessors are not
            # there. PackageFullName carries the same version as a second source,
            # which is why a failure of the first is not said: the second is
            # read next.
            try { $info.PackageVersion = "$($p.Version)" } catch { $null = $_ }
            if (-not $info.PackageVersion -and $info.PackageFullName) {
                # String work on text already read; nothing here can fail to
                # read, and a version not found stays $null.
                try {
                    $parts = @($info.PackageFullName -split '_')
                    if ($parts.Count -ge 2) { $info.PackageVersion = $parts[1] }
                } catch { $null = $_ }
            }
        }
    } catch { & $readFailed 'The Claude package (Get-AppxPackage)' $_ }

    # Application Id comes from the manifest. The old hardcoded "App" fallback
    # produced an invalid AUMID, and it only ever ran when the manifest read had
    # already failed, which is precisely when a wrong guess does damage.
    if ($info.InstallLocation) {
        # A manifest that is not there leaves the fallback below, which says
        # so. One that could not be read is said as well: until 6.3.1
        # Test-Path answered "not there" for it.
        try {
            $manifestPath = Join-Path $info.InstallLocation "AppxManifest.xml"
            if (Get-ClaudeItemOrNull -LiteralPath $manifestPath) {
                [xml]$mx = Get-Content -LiteralPath $manifestPath -ErrorAction Stop
                $ns = New-Object Xml.XmlNamespaceManager($mx.NameTable)
                $ns.AddNamespace("x", "http://schemas.microsoft.com/appx/manifest/foundation/windows10")
                $appNode = $mx.SelectSingleNode("//x:Application", $ns)
                if ($appNode) { $info.ApplicationId = $appNode.GetAttribute("Id") }
            }
        } catch { & $readFailed 'The package manifest' $_ }
    }
    if (-not $info.ApplicationId -and $info.PackageFound) {
        # Every shipped Claude manifest to date uses Id="Claude", matching the
        # package Name. Better than a literal guess, and it is recorded as one.
        # The name was read with the package; a failure here leaves no Id,
        # which the line below does not claim.
        try { $info.ApplicationId = "$($pkgs[0].Name)" } catch { $null = $_ }
        if ($info.ApplicationId) {
            $info.Warnings += "Application Id fell back to package name '$($info.ApplicationId)'"
        }
    }
    if ($info.PackageFamilyName -and $info.ApplicationId) {
        $info.Aumid = "shell:AppsFolder\$($info.PackageFamilyName)!$($info.ApplicationId)"
    }

    # -- The app's data and log folders --------------------------------------
    # After the package query because the package's LocalCache is one of the
    # candidates. Every choice and every candidate passed over goes into
    # Warnings, so claude-env.json says why a path was picked.
    try {
        $paths = Resolve-ClaudeDataPaths -PackageFamilyName $info.PackageFamilyName `
                                         -AppDataDir $info.AppDataDir
        foreach ($k in @('DataRoot', 'DataRootSource', 'LogDir', 'LogDirSource', 'VmCachePath',
                         'BundlePath', 'SessionsDir', 'CodeSessionsDir',
                         'DataRootCandidates', 'LogDirCandidates')) {
            $info[$k] = $paths.$k
        }
        $info.Warnings += @($paths.Notes)
        # Each candidate folder that could not be read is in Notes already.
        foreach ($e in @($paths.Errors)) { $info.ReadErrors += "Data folder candidate $e" }
    } catch { & $readFailed "The app's data folders" $_ }

    # -- Claude executable ---------------------------------------------------
    # Ordered by authority, not by convenience. The package branch runs first
    # because it is the only source that reliably identifies the DESKTOP app,
    # and it is also what makes a WindowsApps install reachable at all: no
    # filesystem search path covers WindowsApps and a brute scan cannot read its
    # TrustedInstaller ACL.
    #
    # Reading the path off a running process looks more reliable but is not.
    # The Claude Code CLI also runs as a process named "claude": the copies the
    # app starts sit in claude-code\<version>\ under its data folder, and a
    # terminal install sits in %USERPROFILE%\.local\bin. Measured 2026-11-21: 4
    # of 18 "claude" processes were the CLI. Taking the first readable one can
    # therefore cache the CLI path and turn every later relaunch into a CLI
    # start instead of an app start. Test-ClaudeCodeCliPath rejects both kinds;
    # the path pattern it replaced here only knew the first.
    # Each source below that cannot be read is said in ReadErrors, and the
    # next one is tried. A candidate path that is there but could not be read
    # is passed over too, since the search is for a file that is there and
    # the next source may have it; when no source has one, each such path is
    # said in ReadErrors, so that "not located" is not taken for "not
    # installed". Until 6.3.1 Test-Path answered "not there" for such a path,
    # without a word, and answered "there" for a folder.
    $exeUnread = New-Object 'System.Collections.Generic.List[string]'
    # $true for a file at that path. Nothing there, or a folder, is $false; a
    # path that could not be read is $false and noted in $exeUnread.
    $exeAt = {
        param([string]$Path)
        if (-not $Path) { return $false }
        try {
            $it = Get-ClaudeItemOrNull -LiteralPath $Path
        } catch {
            $exeUnread.Add("Claude executable candidate $Path could not be read: $(Get-ClaudeErrorText $_)")
            return $false
        }
        return ($null -ne $it -and -not $it.PSIsContainer)
    }
    if ($info.InstallLocation) {
        foreach ($rel in @('app\claude.exe', 'claude.exe')) {
            # Join-Path fails only for a drive that is not there; such a
            # candidate is passed over like one with nothing there.
            try {
                $cand = Join-Path $info.InstallLocation $rel
                if (& $exeAt $cand) {
                    $info.ClaudeExe = $cand; $info.ClaudeExeSource = 'package'; break
                }
            } catch { $null = $_ }
        }
    }
    if (-not $info.ClaudeExe) {
        try {
            foreach ($rp in @(Get-ClaudeProcessByName -Name $info.ProcessName)) {
                try {
                    $cand = $rp.MainModule.FileName
                    if ($cand -and (& $exeAt $cand) -and -not (Test-ClaudeCodeCliPath $cand)) {
                        $info.ClaudeExe = $cand; $info.ClaudeExeSource = 'process'; break
                    }
                } catch {
                    # MainModule needs more access than the image name does,
                    # and the process may have exited since the list was
                    # read. The CIM query below answers for such a process.
                    $null = $_
                }
            }
        } catch { & $readFailed 'The processes named claude' $_ }
    }
    if (-not $info.ClaudeExe) {
        # No match is an empty answer with no error (measured 2026-11-22), so
        # any error is a failed read.
        try {
            foreach ($wp in @(Get-CimInstance Win32_Process -Filter "Name LIKE '%claude%'" `
                              -ErrorAction Stop)) {
                $cand = "$($wp.ExecutablePath)"
                if ($cand -and (& $exeAt $cand) -and -not (Test-ClaudeCodeCliPath $cand)) {
                    $info.ClaudeExe = $cand; $info.ClaudeExeSource = 'wmi'; break
                }
            }
        } catch { & $readFailed 'The process list (CIM Win32_Process)' $_ }
    }
    if (-not $info.ClaudeExe) {
        # Guesses at the usual install folders. A folder variable that is not
        # set has no guess: until 6.3.1 Join-Path of $null threw out of the
        # whole discovery when LOCALAPPDATA was not set. Join-Path fails only
        # for a drive that is not there, which is passed over as well.
        $guesses = @()
        foreach ($g in @(
            @($env:LOCALAPPDATA, 'Programs\claude\Claude.exe'),
            @($env:LOCALAPPDATA, 'AnthropicClaude\Claude.exe'),
            @($env:LOCALAPPDATA, 'Anthropic\Claude\Claude.exe'),
            @($env:ProgramFiles, 'Claude\Claude.exe')
        )) {
            if ($g[0]) { try { $guesses += (Join-Path $g[0] $g[1]) } catch { $null = $_ } }
        }
        foreach ($cand in $guesses) {
            if (& $exeAt $cand) { $info.ClaudeExe = $cand; $info.ClaudeExeSource = 'commonpath'; break }
        }
    }
    if (-not $info.ClaudeExe) {
        $info.Warnings += "Claude executable not located"
        foreach ($u in $exeUnread) { $info.Warnings += $u; $info.ReadErrors += $u }
    }

    # -- Service -------------------------------------------------------------
    # Not registered and could not be read are told apart
    # (Get-ClaudeServiceOrNull). Until 6.3.1 a failed read said "is not
    # registered".
    $svc = $null
    try {
        $svc = Get-ClaudeServiceOrNull -Name $info.ServiceName
        if (-not $svc) { $info.Warnings += "$($info.ServiceName) is not registered" }
    } catch {
        $info.ServiceFound = $null
        & $readFailed "Service $($info.ServiceName)" $_
    }
    if ($svc) {
        $info.ServiceFound = $true
        try {
            $cim = Get-CimInstance Win32_Service `
                   -Filter "Name='$($info.ServiceName)'" -ErrorAction Stop
            if ($cim) {
                $info.ServiceStartMode = "$($cim.StartMode)"
                $raw = "$($cim.PathName)"
                if ($raw) {
                    # PathName may be quoted and may carry arguments.
                    if ($raw -match '^\s*"([^"]+)"') { $exe = $Matches[1] }
                    elseif ($raw -match '^\s*(\S+\.exe)') { $exe = $Matches[1] }
                    else { $exe = $raw.Trim() }
                    $info.ServiceBinaryPath = $exe
                    try {
                        $leaf = [System.IO.Path]::GetFileNameWithoutExtension($exe)
                        if ($leaf) { $info.ServiceProcessName = $leaf }
                    } catch {
                        $info.Warnings += "The service's process name could not be taken from '$exe'; it stays $($info.ServiceProcessName)"
                    }
                }
            }
        } catch {
            # The start mode is on the service object as well. The binary
            # path is not, and the process name stays the one measured on
            # app 6.18355.
            & $readFailed "The details of service $($info.ServiceName) (CIM Win32_Service)" $_
            if ($null -ne $svc.StartType) { $info.ServiceStartMode = "$($svc.StartType)" }
        }
    }

    # Recovery actions are read from the registry, not from parsed sc.exe output.
    # Matching the literal string RESTART in console text fails on any Windows
    # whose display language is not English.
    #
    # No key means no such service, and GetValue answers $null for a value
    # that is not there, so neither is an error. Any error is a failed read,
    # and the answer is then $null, not "not configured": until 6.3.1 an
    # error here read as not configured, and Fix's report said so.
    try {
        $svcKey = Get-ClaudeItemOrNull -LiteralPath "HKLM:\SYSTEM\CurrentControlSet\Services\$($info.ServiceName)"
        if ($svcKey) {
            $fa = $svcKey.GetValue('FailureActions')
            if ($fa) { $info.ServiceRecoveryConfigured = $true }
        }
    } catch {
        $info.ServiceRecoveryConfigured = $null
        & $readFailed "The recovery settings of service $($info.ServiceName)" $_
    }

    # -- VM log selection ----------------------------------------------------
    # See Get-ClaudeVmLogCandidates for which logs count and why. A log folder
    # with nothing there holds no VM log and is passed over, as before; one
    # that could not be read is a read error (Get-ClaudeItemOrNull), where
    # until 6.3.1 Test-Path passed over it without a word.
    try {
        if ($info.LogDir -and (Get-ClaudeItemOrNull -LiteralPath $info.LogDir)) {
            $cands = @(Get-ClaudeVmLogCandidates -LogDir $info.LogDir)
            $info.VmLogCandidates = $cands
            if ($cands.Count -gt 0) { $info.VmLogFile = $cands[0].Path }
            else { $info.Warnings += "No VM log present in $($info.LogDir)" }
        }
    } catch { & $readFailed "The VM logs in $($info.LogDir)" $_ }

    # -- VM cache inventory --------------------------------------------------
    # One walk yields both the real cache size and the real VHDX sizes, so the
    # backup space check can be computed from what is actually on disk instead
    # of a literal, and the help text can quote a true figure.
    if ($SkipCacheInventory) {
        $info.CacheSizeBytes = [long](-1)
    } else {
        try {
            $vhdx  = @()
            $total = [long]0
            # A file or folder the walk cannot read is left out of the total
            # and counted, and Warnings says the total is then a lower bound.
            # An estimate, not a decision: the backup measures what it copies.
            $unread = 0
            $firstUnread = $null
            foreach ($dir in @($info.VmCachePath, $info.BundlePath)) {
                if (-not $dir) { continue }
                # A folder with nothing there holds nothing to count. One that
                # could not be read is counted as unread, so the total is said
                # to be a lower bound; until 6.3.1 Test-Path passed over it and
                # the total read as complete.
                try {
                    if ($null -eq (Get-ClaudeItemOrNull -LiteralPath $dir)) { continue }
                } catch {
                    # Get-Item's message for a path it may not read does not
                    # name the path (measured 2026-11-23: "Access is denied"),
                    # so the folder is named here.
                    $unread++
                    if (-not $firstUnread) { $firstUnread = $dir + ': ' + (Get-ClaudeErrorText $_) }
                    continue
                }
                $walkErrors = $null
                foreach ($f in @(Get-ChildItem $dir -Recurse -File -ErrorAction SilentlyContinue -ErrorVariable walkErrors)) {
                    $total += $f.Length
                    if ($f.Extension -eq '.vhdx') {
                        $vhdx += [pscustomobject]@{
                            Name   = $f.Name
                            Path   = $f.FullName
                            Length = $f.Length
                        }
                    }
                }
                foreach ($we in @($walkErrors)) {
                    if ($null -eq $we) { continue }
                    $unread++
                    if (-not $firstUnread) { $firstUnread = Get-ClaudeErrorText $we }
                }
            }
            $info.VhdxFiles      = @($vhdx)
            $info.CacheSizeBytes = $total
            if ($unread -gt 0) {
                $info.Warnings += "Cache inventory: $unread item(s) could not be read, so CacheSizeBytes is a lower bound (first: $firstUnread)"
            }
        } catch {
            # -1 as when skipped: not known, which a 0 would not say.
            $info.CacheSizeBytes = [long](-1)
            & $readFailed 'The VM cache inventory' $_
        }
    }

    return $info
}

function Save-ClaudeEnvironment {
    <#
    .SYNOPSIS
        Writes a discovery result to a file, claude-env.json, for a person
        reading what discovery found and for bug reports.
    .DESCRIPTION
        Nothing in this toolkit reads the file back: every run discovers
        afresh. Written as UTF-8 without a BOM. Out-File -Encoding utf8 emits a BOM on
        Windows PowerShell 5.1, which breaks anything reading the file as plain
        JSON. Never throws: persistence is a convenience, not a dependency.
    #>
    param(
        [Parameter(Mandatory)]$Environment,
        [Parameter(Mandatory)][string]$Path
    )
    try {
        $dir = Split-Path $Path -Parent
        # A folder Test-Path could not read is made again here, which is
        # harmless: New-Item -Force hands back a folder that is there as it
        # is (measured 2026-11-23), and one that cannot be written to fails
        # the write below, which answers $false.
        if ($dir -and -not (Test-Path $dir)) {
            New-Item $dir -ItemType Directory -Force | Out-Null
        }
        $json = $Environment | ConvertTo-Json -Depth 4
        [System.IO.File]::WriteAllText($Path, $json, `
            (New-Object System.Text.UTF8Encoding($false)))
        return $true
    } catch { return $false }
}

function Get-CoworkHcsGuids {
    <#
    .SYNOPSIS
        Pulls the distinct compute system GUIDs for cowork-vm out of the text
        that "hcsdiag list" prints.
    .DESCRIPTION
        This parse lived in four places: Close-StaleHcsVms in Fix, the HCS
        cleanup step in Prevent, and both the counting check and the kill path
        in Watch. All four had at least one of these two bugs, and fixing one
        copy did nothing for the other three. It lives here now.

        The output looks like this, with the name repeated on both lines and an
        uppercase GUID mid-line:

            cowork-vm-1699151a
                VM,           Running, DE1517EC-...-77A48CB1AD97, cowork-vm-1699151a

        Bug one: the GUID never appears on a line by itself, so a parser that
        looks for a leading GUID and then a following name finds nothing, or
        pairs a GUID with the wrong name.

        Bug two: counting occurrences of the string "cowork-vm" returns 2 for a
        single VM, because the name is on both lines. Watch used that count
        against a threshold of 1, so "Multiple cowork-vm instances in HCS" was a
        standing false positive on every healthy machine.

        Matching the name first and then extracting the GUID from that same
        line, then taking distinct values, fixes both. Order matters: the second
        -match is what leaves the capture in $Matches.

        Takes the text rather than running hcsdiag itself. The three scripts
        invoke it differently (Fix wraps it in a job with a timeout, Watch
        caches the result for 25 seconds) and none of that belongs in a parser.
        It also means the parser is testable without a Hyper-V host.
    #>
    param([string]$ListOutput)

    if (-not $ListOutput) { return @() }
    if ($ListOutput -notmatch 'cowork-vm') { return @() }

    $guidPattern = '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})'
    $found = @()
    foreach ($line in ($ListOutput -split "`r?`n")) {
        if ($line -match 'cowork-vm') {
            if ($line -match $guidPattern) { $found += $Matches[1] }
        }
    }
    return @($found | Select-Object -Unique)
}

function Test-ClaudeCodeCliPath {
    <#
    .SYNOPSIS
        True when an executable path is the Claude Code CLI, not the Claude
        Desktop app.
    .DESCRIPTION
        Both run as a process named "claude". The CLI copies Claude Desktop
        starts for Code sessions and project threads live in
        claude-code\<version>\ under the app's data folder, which the path
        pattern catches. A terminal install lives wherever its installer put
        it, %USERPROFILE%\.local\bin on this machine, so the file's ProductName
        decides the rest. Measured 2026-11-21 on app 6.18355: the app's
        Claude.exe says "Claude", both CLI copies say "Claude Code".

        Every caller asks so that it never takes the CLI for the app. The
        answers:
          $false  an empty path.
          $true   a path with a claude-code folder in it.
          $false  a path with nothing there; the caller's own check of
                  whether it is there passes it over.
          $false  a folder: it is not the CLI. Whether a folder may be used
                  as an exe is the caller's question.
          $true   a path that is there but could not be read, or whose
                  product name could not be read: not one to start or
                  record as the app.
          $true   a file whose product name is "Claude Code".
          $false  any other file.
        Until 6.3.1 any failed read answered $false, "not the CLI".
    #>
    param([string]$Path)
    if (-not $Path) { return $false }
    if ($Path -match '[\\/]claude-code[\\/]') { return $true }
    try {
        $item = Get-ClaudeItemOrNull -LiteralPath $Path
    } catch {
        return $true
    }
    if (-not $item) { return $false }
    # A folder is not the CLI, and it has no VersionInfo: reading a property
    # an object does not have throws under StrictMode Latest.
    if ($item.PSIsContainer) { return $false }
    # VersionInfo is read when asked for, and a getter that throws reads as
    # $null, so $null here is a product name that could not be read.
    $vi = $item.VersionInfo
    if ($null -eq $vi) { return $true }
    return ("$($vi.ProductName)" -eq 'Claude Code')
}

function Get-ClaudeProcessInventory {
    <#
    .SYNOPSIS
        Sorts every running process named "claude" into the desktop app and
        everything else.
    .DESCRIPTION
        "claude" is not a unique process name. Measured 2026-11-21 on app
        6.18355: 18 processes answered Get-Process -Name claude. 14 were the
        desktop app, under the package InstallLocation. 3 were Claude Code CLI
        processes the app had started for project threads, under
        %LOCALAPPDATA%\Claude-Data\claude_<id>\claude-code\2.1.327\. 1 was a
        terminal CLI in %USERPROFILE%\.local\bin. Every kill in this toolkit
        stopped all 18.

        A process is the desktop app when its executable sits under the package
        InstallLocation, or under another version folder of the same package:
        an update registers the new version while processes started from the
        old folder are still running. With no package, the folder of the
        discovered ClaudeExe stands in for InstallLocation. With neither, a
        ProductName of "Claude" decides.

        A process whose path cannot be read is not counted as the desktop app.
        A process this function cannot place is not one the toolkit may kill.
        Kind is unknown for it, and for a process whose product name was
        needed and could not be read.

        Throws, with the original error, when the processes could not be
        listed, and when a process's path could not be read and the one CIM
        query that answers for it failed. Until 6.3.1 a list that could not
        be read came back empty, which every caller took for "Claude is not
        running".

        Each result carries Process, Id, Path, Kind and IsDesktop. Kind is
        desktop, cli, other or unknown.
    .PARAMETER Environment
        A Get-ClaudeEnvironment result.
    #>
    param([Parameter(Mandatory)]$Environment)

    $roots = @()
    $siblingPattern = $null
    try {
        if ($Environment.InstallLocation) {
            $loc = "$($Environment.InstallLocation)".TrimEnd('\')
            $roots += $loc + '\'
            # InstallLocation is <WindowsApps>\<Name>_<Version>_<Arch>_<ResourceId>_<PublisherId>
            # and the family name is <Name>_<PublisherId>, so any version of
            # this package matches <WindowsApps>\<Name>_*_<PublisherId>\.
            $fam = "$($Environment.PackageFamilyName)"
            $cut = $fam.LastIndexOf('_')
            if ($cut -gt 0) {
                $parent = Split-Path $loc -Parent
                $siblingPattern = '^' + [regex]::Escape($parent + '\' + $fam.Substring(0, $cut) + '_') +
                    '[^\\]*_' + [regex]::Escape($fam.Substring($cut + 1)) + '\\'
            }
        } elseif ($Environment.ClaudeExe -and -not (Test-ClaudeCodeCliPath "$($Environment.ClaudeExe)")) {
            $dir = Split-Path "$($Environment.ClaudeExe)" -Parent
            # A Squirrel install runs from <root>\app-<version>\ beside a
            # <root>\Claude.exe stub, and an update adds a new app-<version>.
            if ((Split-Path $dir -Leaf) -match '^app-\d') { $dir = Split-Path $dir -Parent }
            $roots += $dir.TrimEnd('\') + '\'
        }
    } catch {
        # String work on fields discovery already read. With no roots the
        # product name decides below, so a failure here places nothing wrong.
        $null = $_
    }

    $result = @()
    $cimPaths = $null
    $products = @{}
    # Throws when the list cannot be read (Get-ClaudeProcessByName).
    foreach ($p in @(Get-ClaudeProcessByName -Name $Environment.ProcessName)) {
        $path = $null
        # MainModule needs more access than the image name does, and the
        # process may have exited since the list was read. The CIM query
        # below answers for either.
        try { $path = $p.MainModule.FileName } catch { $null = $_ }
        if (-not $path) {
            # One CIM query answers for every process MainModule could not
            # read. When it fails, this process cannot be placed, so the
            # inventory cannot be told: the error is thrown as it came.
            if ($null -eq $cimPaths) {
                $cimPaths = @{}
                $flt = "Name='$($Environment.ProcessName).exe'"
                foreach ($w in @(Get-CimInstance Win32_Process -Filter $flt -ErrorAction Stop)) {
                    if ($w.ExecutablePath) { $cimPaths[[int]$w.ProcessId] = "$($w.ExecutablePath)" }
                }
            }
            if ($cimPaths.ContainsKey([int]$p.Id)) { $path = $cimPaths[[int]$p.Id] }
        }

        $kind = 'unknown'
        if ($path) {
            $kind = 'other'
            if ($path -match '[\\/]claude-code[\\/]') { $kind = 'cli' }
            else {
                foreach ($r in $roots) {
                    if ($path.StartsWith($r, [System.StringComparison]::OrdinalIgnoreCase)) { $kind = 'desktop'; break }
                }
                if ($kind -eq 'other' -and $siblingPattern -and $path -match $siblingPattern) { $kind = 'desktop' }
                if ($kind -eq 'other') {
                    # Read once per distinct path: fourteen desktop processes
                    # share one executable.
                    # A product name that could not be read is $null, and the
                    # process is then unknown, not other: it may be the CLI,
                    # or with no roots the app. Until 6.3.1 it was other.
                    if (-not $products.ContainsKey($path)) {
                        $products[$path] = $null
                        try {
                            $vi = (Get-Item -LiteralPath $path -Force -ErrorAction Stop).VersionInfo
                            if ($null -ne $vi) { $products[$path] = "$($vi.ProductName)" }
                        } catch {
                            # Said through Kind, which every caller reads.
                            $null = $_
                        }
                    }
                    if ($null -eq $products[$path]) { $kind = 'unknown' }
                    elseif ($products[$path] -eq 'Claude Code') { $kind = 'cli' }
                    elseif ($roots.Count -eq 0 -and $products[$path] -eq 'Claude') { $kind = 'desktop' }
                }
            }
        }
        $result += [pscustomobject]@{
            Process   = $p
            Id        = $p.Id
            Path      = $path
            Kind      = $kind
            IsDesktop = ($kind -eq 'desktop')
        }
    }
    return $result
}

function Get-ClaudeDesktopProcess {
    <#
    .SYNOPSIS
        The running Claude Desktop app processes, and no Claude Code CLI.
    .DESCRIPTION
        Use this for anything that kills Claude, and for anything that asks
        whether the app is running before it does something. A check that only
        holds a repair back because Claude looks busy may keep counting every
        "claude" process: a busy CLI is a reason to wait as well.
        Get-ClaudeProcessInventory says how the two are told apart, and when
        it throws, which this passes on: a caller must not read a list that
        could not be read as "not running".
    .PARAMETER Environment
        A Get-ClaudeEnvironment result.
    #>
    param([Parameter(Mandatory)]$Environment)
    return @(Get-ClaudeProcessInventory -Environment $Environment |
             Where-Object { $_.IsDesktop } | ForEach-Object { $_.Process })
}

function Get-ClaudeAppCliProcess {
    <#
    .SYNOPSIS
        The Claude Code CLI processes the desktop app started, one for each
        project thread or Code session it runs.
    .DESCRIPTION
        Measured 2026-11-21 on app 6.18355: four CLI processes, each a child
        of the app's main process, and one terminal CLI whose parent was
        python.exe. A repair stops the app, and the sessions it started stop
        with it, busy or waiting. A session waiting on a long command uses no
        CPU and writes nothing, so a check of CPU, logs or transcripts can
        miss it. Its process is there to be seen.

        A CLI counts when its parent is any desktop process, so which of the
        app's processes started it does not matter. A terminal CLI is not
        counted: a repair does not end it. A child of the app whose kind
        could not be told (unknown) counts as well, since it may be a
        session; until 6.3.1 one whose product name could not be read was
        not counted.

        Throws when the processes or their parent ids cannot be read, so a
        caller deciding whether a repair may run can fail closed.
    .PARAMETER Environment
        A Get-ClaudeEnvironment result.
    #>
    param([Parameter(Mandatory)]$Environment)

    $inventory = @(Get-ClaudeProcessInventory -Environment $Environment)
    $cli = @($inventory | Where-Object { $_.Kind -eq 'cli' -or $_.Kind -eq 'unknown' })
    if ($cli.Count -eq 0) { return @() }
    $desktop = @{}
    foreach ($e in $inventory) { if ($e.IsDesktop) { $desktop[[int]$e.Id] = $true } }
    if ($desktop.Count -eq 0) { return @() }
    $parentOf = @{}
    foreach ($w in @(Get-CimInstance Win32_Process -Filter "Name='$($Environment.ProcessName).exe'" -ErrorAction Stop)) {
        $parentOf[[int]$w.ProcessId] = [int]$w.ParentProcessId
    }
    return @($cli | Where-Object { $parentOf.ContainsKey([int]$_.Id) -and $desktop.ContainsKey($parentOf[[int]$_.Id]) })
}

function Test-ClaudeCodeTranscriptRecent {
    <#
    .SYNOPSIS
        Whether any Claude Code transcript was written in the last $Minutes
        minutes.
    .DESCRIPTION
        <config dir>\projects\<project>\<session>.jsonl, and a subagent's
        <session>\subagents\**\*.jsonl beside it. The config dir is
        CLAUDE_CONFIG_DIR when set, else %USERPROFILE%\.claude. The CLI
        appends on every message and tool result. Measured 2026-11-21: 3,904
        files, listed in 160 ms.

        The monitor and Fix -Quiet both ask with 35 minutes. A tool call
        writes nothing here until it returns, and a foreground command may
        run for 30 minutes: measured 2026-11-21, three Bash calls of 25
        minutes each in a row in one project thread, nothing written during
        the first two.

        Ages are taken in UTC. Local time jumps an hour at a daylight-saving
        change, and until 6.3.0 a transcript written 20 minutes before the
        spring change read as 80 minutes old, so an unattended repair could
        start while a session was working.

        Throws when it cannot tell: no config dir to look in, a projects
        folder that could not be read, or a folder in it that could not be
        listed and no recent transcript elsewhere. A caller deciding whether
        a repair may run holds back on the throw and says why. Until 6.3.1
        each of these answered $false, "no session busy", and let the repair
        run. No projects folder at all is $false: nothing has been written.
    #>
    param([int]$Minutes = 35)

    $cfgDir = $env:CLAUDE_CONFIG_DIR
    if (-not $cfgDir -and $env:USERPROFILE) { $cfgDir = Join-Path $env:USERPROFILE ".claude" }
    if (-not $cfgDir) { throw 'Neither CLAUDE_CONFIG_DIR nor USERPROFILE is set, so there is no transcript folder to read' }
    $projects = Join-Path $cfgDir "projects"
    # Named here: the error Get-Item gives for a folder it may not open says
    # only "Access is denied" (measured 2026-11-23), and a caller says this
    # error as why it could not tell. Until 6.3.1 the monitor's line was
    # "Test-ClaudeCodeActive failed (Access is denied)".
    try {
        $folder = Get-ClaudeItemOrNull -LiteralPath $projects
    } catch {
        throw "The Claude Code transcript folder $projects could not be read: $(Get-ClaudeErrorText $_)"
    }
    if (-not $folder -or -not $folder.PSIsContainer) { return $false }
    $now = [datetime]::UtcNow
    # Errors are gathered, not shown: a recent transcript answers the
    # question whatever else could not be listed. Without one, the first
    # error is thrown, unless it is a file or folder that went away during
    # the walk (PathNotFound), which a session does when it ends.
    $walkErrors = $null
    $recent = Get-ChildItem -LiteralPath $projects -Filter "*.jsonl" -Recurse -File -Force `
            -ErrorAction SilentlyContinue -ErrorVariable walkErrors |
        Where-Object { ($now - $_.LastWriteTimeUtc).TotalMinutes -lt $Minutes } |
        Select-Object -First 1
    if ($recent) { return $true }
    # @($null) holds one element, so a $null is passed over.
    foreach ($we in @($walkErrors)) {
        if ($null -eq $we) { continue }
        if (-not (Test-ClaudeNotFoundError -ErrorRecord $we -ErrorId 'PathNotFound' `
                -ExceptionType 'System.Management.Automation.ItemNotFoundException')) { throw $we }
    }
    return $false
}

function Get-ClaudeDataCandidates {
    <#
    .SYNOPSIS
        Every folder that may hold the Claude Desktop app's data or its logs,
        freshest first.
    .DESCRIPTION
        The app has moved its data before. Until app 6.18355 it kept everything
        in %APPDATA%\Claude. Measured 2026-11-21 on 6.18355: %APPDATA%\Claude
        holds only this toolkit's own files, the data root is
        %LOCALAPPDATA%\Claude-Data\claude_<publisher id> (Preferences written
        that minute), whose logs folder the app last wrote on 2026-08-21, and
        the live logs are in %LOCALAPPDATA%\Claude\logs. A fixed path is wrong
        on one build or the other, so every place the app is known to use is
        listed and the caller takes the freshest.

        DataRoots: each existing candidate folder. Newest is the latest write
        among the app's own files in it (Preferences, Local State, config.json,
        Network, Session Storage, window-state.json, lockfile). IsAppData is
        true when it holds Preferences or Local State, which every Electron data
        folder has. The toolkit's own files (watch-logs, fix-logs, launchers)
        are not markers, so a folder does not look fresh because the toolkit
        writes to it.

        LogDirs: the logs folder of each candidate, where it exists. Newest is
        the latest write to a *.log directly in it.

        Cheap enough to call on every poll: one item read and one directory
        listing per candidate, nothing recursive. Never throws.

        Errors: one line for each folder or file that is there but could not
        be read, naming it and the error. A candidate behind such an error
        may be missing from DataRoots or LogDirs, or carry a Newest that is
        too old, so a caller that would act on the choice (move a folder,
        call the app idle) must not while Errors holds anything. Until 6.3.1
        a folder that could not be read looked the same as one that was not
        there.
    .PARAMETER PackageFamilyName
        The MSIX package family. Its LocalCache, where Windows redirects a
        packaged app's AppData writes, is probed as well.
    #>
    param([string]$PackageFamilyName)

    $bases = @()
    if ($env:LOCALAPPDATA) { $bases += @{ Path = $env:LOCALAPPDATA; Tag = 'LOCALAPPDATA' } }
    if ($env:APPDATA)      { $bases += @{ Path = $env:APPDATA;      Tag = 'APPDATA' } }
    if ($PackageFamilyName -and $env:LOCALAPPDATA) {
        $lc = Join-Path $env:LOCALAPPDATA "Packages\$PackageFamilyName\LocalCache"
        $bases += @{ Path = (Join-Path $lc 'Local');   Tag = 'LocalCache\Local' }
        $bases += @{ Path = (Join-Path $lc 'Roaming'); Tag = 'LocalCache\Roaming' }
    }

    $wanted = @()
    if ($env:CLAUDE_USER_DATA_DIR) {
        $wanted += @{ Path = $env:CLAUDE_USER_DATA_DIR; Source = 'CLAUDE_USER_DATA_DIR' }
    }
    $errors = @()
    $readError = { param($What, $ErrorRecord) "$What could not be read: $(Get-ClaudeErrorText $ErrorRecord)" }
    # A folder found and then gone when listed (PathNotFound, the identity
    # Get-ChildItem gives for a path not there, measured 2026-11-22) holds
    # nothing, as one never found does: an update that removes the folder it
    # moved from must not read as a folder that could not be read.
    $goneError = { param($ErrorRecord) Test-ClaudeNotFoundError -ErrorRecord $ErrorRecord -ErrorId 'PathNotFound' `
            -ExceptionType 'System.Management.Automation.ItemNotFoundException' }
    foreach ($b in $bases) {
        foreach ($n in @('Claude-Data', 'Claude-3p')) {
            $parent = Join-Path $b.Path $n
            try {
                $pi = Get-ClaudeItemOrNull -LiteralPath $parent
                if ($pi -and $pi.PSIsContainer) {
                    foreach ($d in @(Get-ChildItem -LiteralPath $parent -Directory -Force -ErrorAction Stop)) {
                        $wanted += @{ Path = $d.FullName; Source = "$($b.Tag)\$n" }
                    }
                }
            } catch { if (-not (& $goneError $_)) { $errors += & $readError $parent $_ } }
        }
        $wanted += @{ Path = (Join-Path $b.Path 'Claude'); Source = "$($b.Tag)\Claude" }
    }

    $markers = @('Preferences', 'Local State', 'config.json', 'Network', 'Session Storage',
                 'window-state.json', 'lockfile')
    $roots = @()
    $logs  = @()
    $seen  = @{}
    foreach ($w in $wanted) {
        $full = "$($w.Path)".TrimEnd('\')
        if (-not $full -or $seen.ContainsKey($full.ToLowerInvariant())) { continue }
        $seen[$full.ToLowerInvariant()] = $true
        try {
            $fi = Get-ClaudeItemOrNull -LiteralPath $full
        } catch { $errors += & $readError $full $_; continue }
        if (-not $fi -or -not $fi.PSIsContainer) { continue }

        $newest = $null
        $isApp  = $false
        foreach ($m in $markers) {
            $mp = Join-Path $full $m
            try {
                $mi = Get-ClaudeItemOrNull -LiteralPath $mp
            } catch { $errors += & $readError $mp $_; continue }
            if (-not $mi) { continue }
            if ($m -eq 'Preferences' -or $m -eq 'Local State') { $isApp = $true }
            if ($null -eq $newest -or $mi.LastWriteTime -gt $newest) { $newest = $mi.LastWriteTime }
        }
        $roots += [pscustomobject]@{ Path = $full; Source = $w.Source; IsAppData = $isApp; Newest = $newest }

        $ld = Join-Path $full 'logs'
        try {
            $li = Get-ClaudeItemOrNull -LiteralPath $ld
            if ($li -and $li.PSIsContainer) {
                $lnewest = $null
                foreach ($f in @(Get-ChildItem -LiteralPath $ld -Filter '*.log' -File -Force -ErrorAction Stop)) {
                    if ($null -eq $lnewest -or $f.LastWriteTime -gt $lnewest) { $lnewest = $f.LastWriteTime }
                }
                $logs += [pscustomobject]@{ Path = $ld; Source = $w.Source; Newest = $lnewest }
            }
        } catch { if (-not (& $goneError $_)) { $errors += & $readError $ld $_ } }
    }

    $byNewest = @{ Expression = { if ($_.Newest) { $_.Newest } else { [datetime]::MinValue } }; Descending = $true }
    return [pscustomobject]@{
        DataRoots = @($roots | Sort-Object -Property @{ Expression = { [int]$_.IsAppData }; Descending = $true }, $byNewest)
        LogDirs   = @($logs | Sort-Object -Property $byNewest)
        Errors    = @($errors)
    }
}

function Resolve-ClaudeDataPaths {
    <#
    .SYNOPSIS
        Picks the app's data root and log folder out of
        Get-ClaudeDataCandidates and names the folders the repairs act on.
    .DESCRIPTION
        DataRoot is the freshest candidate holding Preferences or Local State.
        LogDir is the logs folder whose newest *.log was written last. The two
        are chosen separately because they are not nested: on app 6.18355 the
        data root is %LOCALAPPDATA%\Claude-Data\claude_<id> and the live logs
        are in %LOCALAPPDATA%\Claude\logs.

        VmCachePath, BundlePath, SessionsDir and CodeSessionsDir are named
        folders under DataRoot. Measured 2026-11-21 on 6.18355, all four are
        there: claude-code-vm, vm_bundles, local-agent-mode-sessions and
        claude-code-sessions.

        When no candidate qualifies, a path falls back to where it was before
        6.2.0, under AppDataDir, and says so. Notes holds one line per choice
        and one per candidate passed over; Get-ClaudeEnvironment adds them to
        Warnings.

        This toolkit's own state (fix-logs, watch-logs, watch-state.json, the
        generated launchers) stays in AppDataDir whatever is chosen here.
        Never throws: Get-ClaudeDataCandidates does not, and the rest is string
        work.

        Errors passes on Get-ClaudeDataCandidates' Errors, and Notes has a
        line for each: a choice made while one of them stands may be the
        wrong one, and a caller that would act on it must not.
    .PARAMETER PackageFamilyName
        Passed to Get-ClaudeDataCandidates for the MSIX LocalCache candidates.
    .PARAMETER AppDataDir
        The toolkit's folder, %APPDATA%\Claude, used only as the fallback.
    #>
    param(
        [string]$PackageFamilyName,
        [string]$AppDataDir
    )

    $stamp = { param($t) if ($t) { $t.ToString('yyyy-MM-dd HH:mm') } else { 'never' } }
    $notes = @()
    $c = Get-ClaudeDataCandidates -PackageFamilyName $PackageFamilyName

    $dataRoot   = $null
    $dataSource = $null
    $appRoots   = @($c.DataRoots | Where-Object { $_.IsAppData })
    if ($appRoots.Count -gt 0) {
        $dataRoot   = $appRoots[0].Path
        $dataSource = $appRoots[0].Source
        $notes += "Data root: $dataRoot ($dataSource, app files last written $(& $stamp $appRoots[0].Newest))"
        foreach ($o in @($appRoots | Select-Object -Skip 1)) {
            $notes += "Data root passed over: $($o.Path) ($($o.Source), app files last written $(& $stamp $o.Newest))"
        }
    } elseif ($AppDataDir) {
        $dataRoot   = $AppDataDir
        $dataSource = 'fallback'
        $notes += "No folder holding Preferences or Local State was found; the data root falls back to $AppDataDir"
    }

    $logDir    = $null
    $logSource = $null
    $liveLogs  = @($c.LogDirs | Where-Object { $_.Newest })
    if ($liveLogs.Count -gt 0) {
        $logDir    = $liveLogs[0].Path
        $logSource = $liveLogs[0].Source
        $notes += "Log folder: $logDir ($logSource, newest log $(& $stamp $liveLogs[0].Newest))"
        foreach ($o in @($liveLogs | Select-Object -Skip 1)) {
            $notes += "Log folder passed over: $($o.Path) ($($o.Source), newest log $(& $stamp $o.Newest))"
        }
    } elseif ($AppDataDir) {
        $logDir    = Join-Path $AppDataDir 'logs'
        $logSource = 'fallback'
        $notes += "No logs folder holding a *.log was found; the log folder falls back to $logDir"
    }
    foreach ($e in @($c.Errors)) {
        $notes += "Candidate not read, so the choice above may be wrong: $e"
    }

    $under = { param($leaf) if ($dataRoot) { Join-Path $dataRoot $leaf } else { $null } }
    return [pscustomobject]@{
        DataRoot           = $dataRoot
        DataRootSource     = $dataSource
        LogDir             = $logDir
        LogDirSource       = $logSource
        VmCachePath        = (& $under 'claude-code-vm')
        BundlePath         = (& $under 'vm_bundles')
        SessionsDir        = (& $under 'local-agent-mode-sessions')
        CodeSessionsDir    = (& $under 'claude-code-sessions')
        DataRootCandidates = @($c.DataRoots | ForEach-Object {
            '{0} ({1}, app data {2}, app files last written {3})' -f $_.Path, $_.Source, $_.IsAppData, (& $stamp $_.Newest)
        })
        LogDirCandidates   = @($c.LogDirs | ForEach-Object {
            '{0} ({1}, newest log {2})' -f $_.Path, $_.Source, (& $stamp $_.Newest)
        })
        Notes              = @($notes)
        Errors             = @($c.Errors)
    }
}

function Get-ClaudeVmLogCandidates {
    <#
    .SYNOPSIS
        The VM logs present in a Claude log folder, newest first.
    .DESCRIPTION
        Deliberately restricted to VM-relevant logs. main.log is excluded: it is
        the Electron main process log and it can sit frozen for a day at a time
        on a perfectly healthy machine, so keying staleness on it is a
        guaranteed false positive. Newest wins, and the age is recorded so a
        caller can reject a candidate that predates its own lookback window.

        A function of its own since 6.2.0 so Watch can pick the VM log again
        when the log folder moves under it.

        Throws, with the original error, when a VM log, or the folder, could
        not be read. A log that is not there is left out. Until 6.3.1 a log
        that could not be read was left out as well, so the newest log could
        be one that had stopped while the live one went unseen.
    #>
    param([string]$LogDir)

    $cands = @()
    if (-not $LogDir) { return $cands }
    foreach ($n in @('cowork_vm_node.log', 'coworkd.log', 'cowork-service.log')) {
        $fi = Get-ClaudeItemOrNull -LiteralPath (Join-Path $LogDir $n)
        if ($fi) {
            $cands += [pscustomobject]@{
                Name          = $n
                Path          = $fi.FullName
                LastWriteTime = $fi.LastWriteTime
                LastWriteTimeUtc = $fi.LastWriteTimeUtc
                AgeHours      = [math]::Round( `
                    ([datetime]::UtcNow - $fi.LastWriteTimeUtc).TotalHours, 2)
                Length        = $fi.Length
            }
        }
    }
    return @($cands | Sort-Object LastWriteTimeUtc -Descending)
}

function Get-ClaudeVmLogAge {
    <#
    .SYNOPSIS
        Seconds since any VM log in a Claude log folder was last written, or
        $null when the folder holds none.
    .DESCRIPTION
        Reads every VM log, now. App 6.18355 writes two of them, and which one
        is newest changes through the day: cowork_vm_node.log is the host side
        (each boot, each Cowork command) and coworkd.log is the daemon inside
        the VM (each command again, and a clock line every few hours). A
        freshness test that read only the log that happened to be newest when
        the script started could be reading the quiet one. Throws when a VM
        log could not be read (Get-ClaudeVmLogCandidates): until 6.3.1 the
        age was then that of the logs that could be read, or $null, "no VM
        log", and an age is what the staleness repair and the repair gate
        act on. Measured in UTC, so a daylight-saving change does not move
        it an hour.
    #>
    param([string]$LogDir)

    $c = @(Get-ClaudeVmLogCandidates -LogDir $LogDir)
    if ($c.Count -eq 0) { return $null }
    return ([datetime]::UtcNow - $c[0].LastWriteTimeUtc).TotalSeconds
}

function Get-CoworkWorkerProcessIds {
    <#
    .SYNOPSIS
        The process ids of the vmwp.exe workers that run the Cowork VM.
    .DESCRIPTION
        Every Hyper-V VM, and every WSL2 or Windows Sandbox utility VM, has a
        vmwp.exe of its own. A worker's first command-line argument is the
        GUID of the compute system it runs:

            "C:\Windows\System32\vmwp.exe" ECDA1B8B-...-0A4035733BF3 0x330

        Matching that against the GUIDs Get-CoworkHcsGuids reads out of
        "hcsdiag list" picks the Cowork VM's worker and nothing else. Until
        6.3.0 Prevent and the monitor raised the priority of every vmwp.exe
        on the machine.

        Takes the list text and the worker objects (anything with ProcessId
        and CommandLine, as Win32_Process gives) rather than querying either,
        so it can be tested without Hyper-V.
    #>
    param([string]$ListOutput, [object[]]$Workers)

    $guids = @(Get-CoworkHcsGuids -ListOutput $ListOutput)
    if ($guids.Count -eq 0) { return @() }
    $ids = @()
    foreach ($w in @($Workers)) {
        if ($null -eq $w) { continue }
        $g = Get-VmWorkerGuid -CommandLine ([string]$w.CommandLine)
        if ($g -and $guids -contains $g) { $ids += [int]$w.ProcessId }
    }
    return @($ids)
}

function Get-VmWorkerGuid {
    <#
    .SYNOPSIS
        The GUID of the compute system a vmwp.exe runs, read from its command
        line, or $null when the line holds none.
    .DESCRIPTION
        The worker takes that GUID as its first argument:

            "C:\Windows\System32\vmwp.exe" ECDA1B8B-...-0A4035733BF3 0x330

        Only the first argument is read. Until 6.3.0 Fix took the first GUID
        anywhere on the line.
    #>
    param([string]$CommandLine)

    if ($CommandLine -match '^\s*(?:"[^"]*"|\S+)\s+\{?([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})\b') {
        return $Matches[1]
    }
    return $null
}

function Get-CoworkHcsSurvivors {
    <#
    .SYNOPSIS
        Which of the given Cowork compute systems "hcsdiag list" still shows,
        or $null when the list could not be read.
    .DESCRIPTION
        A kill that returned is no proof the compute system is gone. Until
        6.3.0 Fix counted every kill it sent as done, and the monitor logged
        "killed" for each one, whether hcsdiag had removed anything or not.

        $List is a script block that returns the list text when hcsdiag
        answered, and nothing (or $null) when it did not: a timeout or an
        exit code other than 0. Each script keeps its own way of running
        hcsdiag: Fix and the monitor through a job with a timeout, Prevent
        directly. The list is read up to $Attempts times, $DelaySeconds
        apart, because a compute system can take a moment to leave it.

        An empty answer is an answer: "hcsdiag list" prints nothing at all
        when no compute system is left, which is what a kill of the only one
        leaves. Until 6.3.0 that read as "did not answer", so every kill of a
        lone Cowork VM was reported as unconfirmed.

        Returns an array, empty when every one is gone. The leading comma
        keeps an empty array from reaching the caller as $null, which here
        means "not known": the first read did not answer. A later read that
        does not answer keeps what the read before it showed, so a kill that
        read confirmed is not thrown away; the ones returned were still
        listed when the list was last read.
    #>
    param([string[]]$Guids, [scriptblock]$List, [int]$Attempts = 2, [int]$DelaySeconds = 2)

    $left = @($Guids)
    for ($i = 0; $i -lt $Attempts; $i++) {
        if ($i -gt 0) { Start-Sleep -Seconds $DelaySeconds }
        $out = & $List
        if ($null -eq $out) {
            if ($i -eq 0) { return $null }
            break
        }
        # Joined with newlines: [string] on an array joins with spaces, and the
        # parser reads the list one line at a time.
        $text = @($out) -join "`n"
        $listed = @(Get-CoworkHcsGuids -ListOutput $text)
        $left = @($left | Where-Object { $listed -contains $_ })
        if ($left.Count -eq 0) { break }
    }
    return ,$left
}

function Get-W32tmOffsetSeconds {
    <#
    .SYNOPSIS
        The clock offset in seconds from what "w32tm /stripchart /dataonly"
        prints, or $null when the text holds none.
    .DESCRIPTION
        Each sample line ends in the offset:

            20:22:22, -00.0684055s

        w32tm writes the number in the system locale, so a machine set to
        Dutch or German prints "-00,0684055s". The pattern every copy of this
        check used, a dot only, never matched there, and the drift check read
        "could not measure" on every such machine. Both separators are taken,
        and the number is parsed with the invariant culture so the result does
        not depend on the locale either. Error lines ("error: 0x800705B4")
        carry no offset and are skipped; the last sample wins.
    #>
    param([string]$Text)

    if (-not $Text) { return $null }
    $value = $null
    foreach ($line in ($Text -split "`r?`n")) {
        if ($line -match '([+-]?\d+[.,]\d+)s\s*$') {
            $value = [double]::Parse(($Matches[1] -replace ',', '.'), [Globalization.CultureInfo]::InvariantCulture)
        }
    }
    return $value
}
#endregion ClaudeEnv

# -- Runtime discovery ---------------------------------------------------
# -SkipCacheInventory because this polls: the cache walk is the expensive half
# of discovery (measured 590ms of a 775ms total) and nothing here needs the
# VHDX inventory.
$script:ClaudeEnv = Get-ClaudeEnvironment -SkipCacheInventory

# -- Constants -----------------------------------------------------------
$ToolkitVersion = "6.3.2"
$ServiceName    = $script:ClaudeEnv.ServiceName
$ProcessName    = $script:ClaudeEnv.ProcessName
# This toolkit's folder, %APPDATA%\Claude: watch-logs and watch-state.json
# live here whatever the app does with its own data.
$ClaudeAppData  = $script:ClaudeEnv.AppDataDir

# Fix's exit code for a run it held back before it stopped the app or the
# service: another Fix run held the lock; or Fix's -Quiet gate at its start
# found Claude or the PC in use (a Claude Code session or a recent transcript,
# claude using CPU, a recent VM log, or keyboard or mouse input in the last 3
# minutes); or, right before Step 1's kill, Fix found a Claude Code session or
# could not rule one out. It must match $script:BlockedExitCode
# in Fix-ClaudeDesktop.ps1, which a test checks. Such a run is not counted as
# an attempt at a repair, and a repair for a one-shot trigger is kept pending
# and tried again once the hold has cleared, as is one for a one-shot trigger
# that fires during the wait the hold leaves: see Invoke-HeldBackRetry. Each
# run Fix holds back in a row doubles the wait before the next repair, from 2
# minutes up to 60, until a run it does not hold back, a Fix run by hand that
# started the repair, or a hold 4 hours or more after the one before: see
# Invoke-AutoFix.
# Until 6.3.0 Fix exited 0 for it, so it was logged as COMPLETE, counted
# toward the "failed 3 times" escalation, and what had been detected was
# dropped.
$FixHeldBackExitCode = 3

# Fix's exit code for a run that stopped part-way: Steps 1 and 2 had run, so
# the app and the service were stopped if they were running, and then, before
# Step 4 or an escalation would stop the app again, Fix found a Claude Code
# session or could not rule one out. If this run stopped the service, Fix
# tries to start it again before it exits, and logs a warning if that start
# fails. It must match
# $script:StoppedPartWayExitCode in Fix-ClaudeDesktop.ps1, which a test checks.
# The run changed things, so unlike a held-back run it is an attempt at a
# repair: it counts toward the escalation, what was detected is reset, the
# normal cooldown applies, and the longer wait that runs Fix held back in a
# row build up starts again from 2 minutes.
$FixStoppedPartWayExitCode = 4

# The app's freshest logs folder, from Resolve-ClaudeDataPaths. Until 6.2.0
# this was %APPDATA%\Claude\logs, which app 6.18355 no longer writes: its logs
# are in %LOCALAPPDATA%\Claude\logs. Sync-ClaudeDataPaths looks again while
# this runs, so an app update that moves the folder is followed.
$ClaudeLogDir   = $script:ClaudeEnv.LogDir

if (-not $ClaudeAppData -or -not $ClaudeLogDir) {
    Write-Host "  [!] Could not locate the per-user Claude folder. Nothing to monitor." -ForegroundColor Red
    exit 1
}
$WatchLogDir    = Join-Path $ClaudeAppData "watch-logs"
# Fix-ClaudeDesktop's own log folder, its $LogDir. A Fix log there newer than
# the latest hold of a pending repair, and holding $FixRepairLine, is a Fix
# run since that started the repair, which ends the pending one: see
# Get-FixRunAfter.
$FixLogDir      = Join-Path $ClaudeAppData "fix-logs"
# The line Fix logs once a run is past every check that holds it back and
# has begun the repair. A Fix log without it is a run that changed nothing,
# such as a Diagnostic run, a Cancel or a No at "Proceed?", and does not end
# a pending repair. It must match $script:RepairUnderWayLine in
# Fix-ClaudeDesktop.ps1, which a test checks.
$FixRepairLine  = 'Repair under way: this run is past every check that holds it back'

# -- Stealth-update marker -----------------------------------------------
# The app writes this file immediately before it quits itself to install a
# background update, and consumes it on the next launch to put the window and
# the tab history back. Derived, never hardcoded: the package family name and
# the data root are whatever discovery found, not the Claude_pzs8sxrjxfjjc seen
# on one machine.
#
# The path is app.getPath('userData'). Before app 6.18355 a packaged MSIX app
# had that redirected into the package's own LocalCache\Roaming\Claude. On
# 6.18355 userData is the data root, %LOCALAPPDATA%\Claude-Data\claude_<id>
# (measured 2026-11-21: Preferences written that minute there, and the
# LocalCache folder holding nothing but a .user-data-left.claude_<id> note).
# So the marker is looked for in the discovered data root, and in the
# LocalCache folder only when no data root was found. No stealth update was
# seen on 6.18355 itself; the location follows from where userData now is.
#
# Only an MSIX install is watched, deliberately. The defect this feature works
# around is specific to the MSIX update path, where quitAndInstall hands off to
# the AppX deployment API and a package registration does not start the app
# afterwards. The NSIS path relaunches on its own, so watching for a marker
# there would only add a second launch to a restart that already worked.
$StealthMarkerName = "stealth-relaunch"
function Get-StealthMarkerPath {
    if (-not $script:ClaudeEnv.PackageFamilyName) { return $null }
    if ($script:ClaudeEnv.DataRoot -and $script:ClaudeEnv.DataRootSource -ne 'fallback') {
        return (Join-Path $script:ClaudeEnv.DataRoot $StealthMarkerName)
    }
    if (-not $env:LOCALAPPDATA) { return $null }
    return (Join-Path $env:LOCALAPPDATA (
        "Packages\{0}\LocalCache\Roaming\Claude\{1}" -f `
            $script:ClaudeEnv.PackageFamilyName, $StealthMarkerName))
}
$script:StealthMarkerPath = Get-StealthMarkerPath

# -- Log error signatures ------------------------------------------------
# Matched literally, and only inside an [error] entry of the app's VM and main
# logs: see Find-LogErrorPattern and $LogTriggerFilePattern below. One hit in
# a log starts a repair with no confirming second poll, so a line counts only
# when the app itself says it is an error. Test-EventLogErrors matches the same
# list against CoworkVMService events of level Error.
#
# Measured 2026-11-21 across every log in the app's logs folder on app 6.18355,
# rotations included. The strings occur in two places:
#
#   - cowork_vm_node1.log, 2026-09-11: 31 [error] [VM:start] "Startup failed"
#     entries carrying "RPC error -1", 30 of them "Plan9 mount failed", and 71
#     [warn] [startVM] retries of the same mount. A real failure; the [error]
#     entries still start a repair.
#   - main.log and its rotations, 2026-10-31 to 2026-11-18: 14 [info]
#     [remote-bash] lines carrying "RPC error -1". A remote session's probe
#     finding its worker already running and carrying on. Harmless. Until
#     6.2.1 each one started the repair path, which restarts the app.
#
# "bad address" left the list in 6.2.1. It is the bare text of EFAULT, which any
# component can log, and the VirtioFS failure it described also carries
# "Plan9 mount failed" or "failed to ensure virtiofs mount".
#
# The old set mixed a regex, "isGuestConnected.*timeout", into a list matched
# literally, so that entry could never match anything; the literal
# "Request timed out: isGuestConnected" below covers the real message.
$LegacyErrorPatterns = @(
    "Plan9 mount failed",
    "failed to ensure virtiofs mount",
    "RPC error -1",
    "HCS operation failed",
    "failed to create compute system",
    "HcsWaitForOperationResult",
    "VM is already running",
    "Request timed out: isGuestConnected",
    "guest connection failed"
)

# The logs Test-LogsForErrors reads: cowork_vm_node*.log, where the VM's start
# and its mounts are logged, and main*.log, the app's own process. The rest of
# the folder is not the VM: in-app-browser, claude.ai-web and unknown-window
# hold what web pages write to their consoles, the mcp logs hold MCP servers'
# output, and coworkd.log, the daemon inside the VM, has no levels. A web page
# that logs "HCS operation failed" is not the VM failing.
$LogTriggerFilePattern = '^(cowork_vm_node|main)\d*\.log$'

# -- Win32 APIs for user activity detection ------------------------------
# NOTE: These APIs only work in interactive sessions (Session 1+).
# In Session 0 (SYSTEM scheduled tasks), they return zero/stale data.
# We detect this and fall back to process-only heuristics.
$script:IsInteractiveSession = $false
# Why the session could not be read, logged below the banner, once the log
# exists; $null when it was read.
$script:SessionReadError = $null
try {
    $sessionId = (Get-Process -Id $PID -ErrorAction Stop).SessionId
    $script:IsInteractiveSession = ($sessionId -gt 0)
} catch { $script:SessionReadError = $_ }

function Import-Win32Activity {
    <#
    .SYNOPSIS
        Loads the Win32Activity type the foreground-window and input checks
        call, and sets $script:Win32ActivityLoaded once the type is there.
        Returns why it is not, or nothing.
    .DESCRIPTION
        The type is looked for after the Add-Type, whatever the Add-Type
        said: one already loaded in this session is a type that works, and
        the error Add-Type gives for it is not worth a word.

        If it is not there, the first [Win32Activity] reference threw inside
        the single try wrapping the user-activity checks, so checks 2
        through 6 were all skipped, the empty catch swallowed the error, and
        the function returned $false, which means "user is not active". The
        auto-fix then proceeded while someone was working. Hence the check
        that it loaded. Until 6.3.1 the Add-Type ran with SilentlyContinue
        and nothing said why the checks were off; the caller now says it.
    #>
    $script:Win32ActivityLoaded = $false
    $why = $null
    try {
        Add-Type -ErrorAction Stop -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class Win32Activity {
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int processId);

    [StructLayout(LayoutKind.Sequential)]
    public struct LASTINPUTINFO {
        public uint cbSize;
        public uint dwTime;
    }

    [DllImport("user32.dll")] public static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);
}
'@
    } catch {
        $why = $_
    }
    try {
        $null = [Win32Activity]
        $script:Win32ActivityLoaded = $true
        return
    } catch {
        if (-not $why) { $why = $_ }
    }
    return $why
}
$script:Win32ActivityError = Import-Win32Activity

# -- Find Fix script -----------------------------------------------------
# Whether each is there is asked by the error's identity. One that could not
# be read is not passed over for the next place: that would start a
# different Fix. The monitor stops on it instead, below, saying why. Until
# 6.3.1 Test-Path answered $false for it, so the next copy was started, or
# the monitor stopped saying the script was not found.
$myDir = if ($PSCommandPath) { Split-Path $PSCommandPath -Parent } else { $PWD.Path }
$fixScript = Join-Path $myDir "Fix-ClaudeDesktop.ps1"
$fixScriptThere = $false
$fixScriptError = $null
try {
    $fixScriptThere = ($null -ne (Get-ClaudeItemOrNull -LiteralPath $fixScript))
} catch {
    $fixScriptError = $_
}

if (-not $fixScriptThere -and -not $fixScriptError) {
    $fallbacks = @(
        "C:\ClaudeFix\Fix-ClaudeDesktop.ps1",
        (Join-Path $env:USERPROFILE "Desktop\Fix-ClaudeDesktop.ps1"),
        (Join-Path $env:USERPROFILE "Documents\Fix-ClaudeDesktop.ps1")
    )
    foreach ($fb in $fallbacks) {
        try {
            if ($null -ne (Get-ClaudeItemOrNull -LiteralPath $fb)) { $fixScript = $fb; $fixScriptThere = $true; break }
        } catch {
            $fixScript = $fb
            $fixScriptError = $_
            break
        }
    }
}

# The fatal check for a missing Fix script is deliberately NOT here. It used
# to be, which meant it ran before $script:WatchLogFile existed, so its only
# output was a Write-Host with no console to write to under
# -WindowStyle Hidden. Task Scheduler retried three times and gave up in
# silence. See the check after Write-WatchLog is defined, below.

# -- Logging -------------------------------------------------------------
if (-not (Test-Path $WatchLogDir)) { New-Item $WatchLogDir -ItemType Directory -Force | Out-Null }

function Remove-OldWatchLog {
    <#
    .SYNOPSIS
        Deletes the watch logs last written more than 30 days ago. Returns
        the first error, or nothing.
    .DESCRIPTION
        Measured in UTC, like every age this monitor keeps: local time jumps
        an hour when summer time starts or ends.

        A folder that is not there has no logs, and a log gone before it is
        deleted needs no deleting: both give PathNotFound with ObjectNotFound
        and ItemNotFoundException (measured 2026-11-22). Any other error is
        returned, and the caller says it. Until 6.3.1 the list and the
        deletes ran with SilentlyContinue, and a folder that could not be
        listed, or a log that could not be deleted, was left without a word
        while the logs grew.
    #>
    # SupportsShouldProcess: this deletes files.
    [CmdletBinding(SupportsShouldProcess)]
    param()
    $notThere = 'System.Management.Automation.ItemNotFoundException'
    try {
        $logs = @(Get-ChildItem -LiteralPath $WatchLogDir -Filter 'watch_*.log' -File -ErrorAction Stop)
    } catch {
        if (Test-ClaudeNotFoundError -ErrorRecord $_ -ErrorId 'PathNotFound' -ExceptionType $notThere) { return }
        return $_
    }
    $firstError = $null
    $cutoff = [datetime]::UtcNow.AddDays(-30)
    foreach ($f in $logs) {
        if ($f.LastWriteTimeUtc -ge $cutoff) { continue }
        if (-not $PSCmdlet.ShouldProcess($f.FullName, 'Delete watch log older than 30 days')) { continue }
        try {
            Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop
        } catch {
            if (Test-ClaudeNotFoundError -ErrorRecord $_ -ErrorId 'PathNotFound' -ExceptionType $notThere) { continue }
            if (-not $firstError) { $firstError = $_ }
        }
    }
    return $firstError
}

# Before there is a log to say it in: kept, and said after the banner.
$script:LogCleanupError = Remove-OldWatchLog

$script:WatchLogFile = Join-Path $WatchLogDir ("watch_{0:yyyyMMdd}.log" -f (Get-Date))

function Write-WatchLog {
    param([string]$Message)
    # Built by concatenation, deliberately.
    #
    # The old form was "[{0:yyyy-MM-dd HH:mm:ss}] $Message" -f (Get-Date), which
    # interpolates $Message INTO the format string before -f runs. Any message
    # containing a brace was then parsed as a format placeholder and threw, and
    # HCS and Hyper-V errors routinely embed {GUID}. This line also sat outside
    # the try below, so inside Invoke-AutoFix it aborted the repair mid-flight,
    # in exactly the situation the repair existed to handle.
    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line  = '[' + $stamp + '] ' + $Message
    try {
        # AppendAllText rather than Out-File -Encoding utf8, which prefixes a
        # BOM on Windows PowerShell 5.1.
        [System.IO.File]::AppendAllText($script:WatchLogFile, $line + "`r`n", `
            (New-Object System.Text.UTF8Encoding($false)))
    } catch {
        # The log cannot say that it cannot be written. The line still goes
        # to the console below when there is one, and each line is tried.
        $null = $_
    }
    if (-not $Quiet) { Write-Host $line -ForegroundColor DarkGray }
}

# Rotate to a new daily log file if the date rolls over
function Update-LogFile {
    # SupportsShouldProcess because the rollover deletes files. It only ever
    # removes watch logs older than 30 days, but "deletes files" is a state
    # change and should be gateable.
    [CmdletBinding(SupportsShouldProcess)]
    param()
    $newPath = Join-Path $WatchLogDir ("watch_{0:yyyyMMdd}.log" -f (Get-Date))
    if ($newPath -ne $script:WatchLogFile) {
        $script:WatchLogFile = $newPath
        # The 30-day watch-log cleanup used to run only at startup, but this is
        # a daemon with ExecutionTimeLimit P365D, so in practice it never ran
        # again. Doing it on rollover is the only point at which it recurs.
        $cleanupError = Remove-OldWatchLog
        if ($cleanupError) {
            Write-CheckError -Check 'old logs' -ErrorRecord $cleanupError `
                -Text "The watch logs older than 30 days in $WatchLogDir could not all be deleted ({0}); they are tried again at the next rollover"
        }
    }
}

# Fatal check for the Fix script, placed AFTER Write-WatchLog exists so the
# reason lands in the watch log rather than on a console that is not there.
# One that could not be read is said with why (see Find Fix script).
if ($fixScriptError -or -not $fixScriptThere) {
    if ($fixScriptError) {
        Write-WatchLog "FATAL: Fix-ClaudeDesktop.ps1 at '$fixScript' could not be read ($(Get-ClaudeErrorText $fixScriptError)). The health monitor cannot start."
    } else {
        Write-WatchLog "FATAL: Fix-ClaudeDesktop.ps1 not found. Looked next to this script at '$myDir' and in the standard fallback locations. The health monitor cannot start."
    }
    if (-not $Quiet) {
        if ($fixScriptError) {
            Write-Host "  [!] Fix-ClaudeDesktop.ps1 at '$fixScript' could not be read. Health monitor cannot start." -ForegroundColor Red
        } else {
            Write-Host "  [!] Fix-ClaudeDesktop.ps1 not found. Health monitor cannot start." -ForegroundColor Red
        }
        Write-Host "      Reason written to: $script:WatchLogFile" -ForegroundColor DarkGray
    }
    exit 1
}

# -- hcsdiag helper (timeout-protected) ----------------------------------
function Invoke-HcsDiag {
    <#
    .SYNOPSIS
        Runs hcsdiag.exe with a timeout. Returns its output when it exited 0,
        '' when it exited 0 and printed nothing, and $null when it exited with
        any other code, timed out or is not installed.
    .DESCRIPTION
        The same contract as Fix's copy. Until 6.3.0 the exit code was thrown
        away: an error hcsdiag printed came back as if it were the list, so
        the check after a kill read the error, found no compute system in it
        and logged the kill as done. The exit code is read inside the job,
        where hcsdiag runs.

        A non-zero exit is logged once, and again only when what hcsdiag says
        changes or it has answered in between: this runs on every poll, and a
        host where it keeps failing would fill the log with the same line.

        -HcsdiagPath is there for the tests, which point it at a stand-in.
    #>
    param(
        [string[]]$Arguments,
        [int]$TimeoutSeconds = 15,
        [string]$HcsdiagPath = "$env:SystemRoot\System32\hcsdiag.exe"
    )
    if (-not (Test-Path -LiteralPath $HcsdiagPath)) { return $null }
    $job = Start-Job -ScriptBlock {
        param($p, $a)
        $text = & $p @a 2>&1 | Out-String
        [pscustomobject]@{ Text = $text; ExitCode = $LASTEXITCODE }
    } -ArgumentList $HcsdiagPath, $Arguments
    $completed = Wait-Job $job -Timeout $TimeoutSeconds
    if (-not $completed) {
        # Clean-up of a job that has not answered: one that cannot be
        # stopped or removed now ends with the monitor, and the timeout is
        # what is said.
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        Write-WatchLog "hcsdiag timed out after ${TimeoutSeconds}s (args: $($Arguments -join ' '))"
        return $null
    }
    # The job's errors are kept and said: a job that could not run hcsdiag
    # gives no exit code, and until 6.3.1 that returned $null without a
    # word, read by the callers as hcsdiag not answering.
    $jobErrors = $null
    $result = @(Receive-Job $job -ErrorAction SilentlyContinue -ErrorVariable jobErrors) |
        Where-Object { $null -ne $_ -and $_.PSObject.Properties['ExitCode'] } | Select-Object -Last 1
    $jobState = $job.State
    # As above: the result is in hand.
    Remove-Job $job -Force -ErrorAction SilentlyContinue
    if ($null -eq $result -or $null -eq $result.ExitCode) {
        $first = @($jobErrors) | Where-Object { $null -ne $_ } | Select-Object -First 1
        $why = if ($first) { Get-CheckErrorText -ErrorRecord $first } else { "the job ended $jobState with no exit code" }
        if (Test-CheckErrorDue -Check 'hcsdiag' -Message $why -Now ([datetime]::UtcNow)) {
            Write-WatchLog "hcsdiag $(@($Arguments)[0]) could not be run ($why); the same error is logged at most once an hour"
        }
        return $null
    }
    if ($result.ExitCode -ne 0) {
        $firstLine = @(([string]$result.Text) -split "`r?`n" | Where-Object { $_.Trim() }) | Select-Object -First 1
        $said = if ($firstLine) { ": $($firstLine.Trim())" } else { '' }
        # The verb only, so a kill of each GUID in turn is one note, not one each.
        $note = "hcsdiag $(@($Arguments)[0]) exited $($result.ExitCode)$said"
        if ($note -ne $script:LastHcsDiagNote) {
            Write-WatchLog $note
            $script:LastHcsDiagNote = $note
        }
        return $null
    }
    $script:LastHcsDiagNote = $null
    return [string]$result.Text
}

function Get-HcsListCached {
    <#
    .SYNOPSIS
        Returns "hcsdiag list" output, cached for a short interval.
    .DESCRIPTION
        Invoke-HcsDiag runs its call inside Start-Job, which means a whole
        child PowerShell runspace every time. It was called on every poll from
        more than one check: at least 2,880 runspaces a day at the default 30
        second interval, and one more per GUID inside the cleanup loop.

        HCS instance state does not change meaningfully within a few seconds,
        so a single call per poll serves every consumer in that poll.

        A failed or timed-out call is cached too, deliberately. Retrying a
        hanging hcsdiag several times inside one poll is worse than waiting for
        the next one.
    #>
    param([int]$MaxAgeSeconds = 25)

    # UTC, so a change of the clock to or from summer time does not make the
    # cache look an hour old or an hour from now.
    $now = [datetime]::UtcNow
    if (($now - $script:LastHcsListTime).TotalSeconds -lt $MaxAgeSeconds) {
        return $script:LastHcsListText
    }
    $script:LastHcsListTime = $now
    $script:LastHcsListText = Invoke-HcsDiag -Arguments "list"
    return $script:LastHcsListText
}

# -- State ---------------------------------------------------------------
# Every time below is UTC. Local time jumps an hour when summer time starts or
# ends, and for that hour a cooldown, a rate limit or the count of fixes in the
# last 30 minutes, measured in local time, read an hour too long or too short.
# Until 6.3.0 they were local.
$script:StartTime            = [datetime]::UtcNow
$script:LastHcsListTime      = [datetime]::MinValue
$script:LastHcsListText      = $null
# The last non-zero exit Invoke-HcsDiag logged, so it is not logged every poll.
$script:LastHcsDiagNote      = $null
$script:LastFixTime          = [datetime]::MinValue
$script:LogBaselines         = @{}   # Track file sizes per log file
$script:FixCount             = 0
$script:LastTimeSyncCheck    = [datetime]::MinValue
# UTC time before which Test-TimeSyncHealth forces no resync, set when one
# failed or did not bring the clock back. MinValue: no hold.
$script:ResyncBackoffUntil   = [datetime]::MinValue
# When each error a check failed with was last logged, by the check and the
# error's message, UTC (Test-CheckErrorDue). In memory only.
$script:CheckErrorsSaid      = @{}
# When each STORAGE WARNING Check 4b logged was last logged, by its text, UTC
# (Test-StorageWarningDue). In memory only.
$script:StorageWarningsSaid  = @{}
$script:VmLogStaleCount      = 0     # Consecutive stale checks
$script:VmLogEverActive      = $false  # Has VM log ever been active this session?
$script:ServiceDownCount     = 0     # Consecutive service-down checks
$script:EventLogHitCount     = 0     # Consecutive event log hits
$script:LastVmcomputeCheck   = [datetime]::MinValue
$script:VmcomputeLeakCount   = 0     # Consecutive vmcompute handle leak checks
$script:LastHcsStateLogTime  = [datetime]::MinValue
$script:LastShutdownFailLogTime = [datetime]::MinValue
$script:FixHistory           = New-Object System.Collections.ArrayList
# A repair held back for a one-shot trigger (Fix's exit 3, BLOCKED or
# DEFERRED in Invoke-AutoFix, or the wait one of them leaves), which
# Invoke-HeldBackRetry tries again once the hold has cleared: the reason it
# was triggered for, when its 4 hours started (its first hold, or the latest
# one-shot trigger that took it over), when it was last held back, how many
# times Fix held it back after the 30-second warning, and how many of those
# were retries, all persisted; when the hold is next asked about, and when
# the last "still waiting" line was written, in memory only. $null, MinValue
# and 0: none.
$script:PendingFixReason     = $null
$script:PendingFixSince      = [datetime]::MinValue
$script:PendingFixHeldAt     = [datetime]::MinValue
$script:PendingFixTries      = 0
$script:PendingFixRetriesHeld = 0
$script:PendingFixNextCheck  = [datetime]::MinValue
$script:PendingFixWaitLogged = [datetime]::MinValue
# In memory only, for the repair pending now: whether the line that a Fix
# run holds Fix's lock has been written while it holds it
# (Write-FixRunningSkip). $false: not yet.
$script:PendingFixLockLogged  = $false
# How many runs in a row Fix has held back (exit 3), and when it held back
# the latest, UTC, both persisted. Each one doubles the wait before the next
# repair (Get-HeldBackRetryWait); a run Fix does not hold back, a Fix log
# since the latest that started the repair (a run by hand), or a hold 4
# hours or more after the one before, starts the count again. 0 and
# MinValue: none.
$script:FixHeldInARow         = 0
$script:FixHeldLastTime       = [datetime]::MinValue
# Whether the cooldown in $script:LastFixTime was left by a hold, BLOCKED,
# DEFERRED or Fix's exit 3, rather than by a repair that ran (this monitor's,
# or a Fix run by hand after runs Fix held back), an escalation or a cancel.
# Persisted with it. A one-shot trigger that fires in a cooldown a hold left
# is kept as a held-back repair (Invoke-AutoFix). $false: not.
$script:CooldownFromHold      = $false
# When each line Test-LineDue keeps from repeating was last written, by key,
# UTC. In memory only.
$script:RepeatLinesSaid       = @{}
# Epoch ms of the last stealth marker this monitor made a decision about, 0 for
# none. Persisted, because the whole point of it is to survive the restart that
# a failed launch attempt could otherwise turn into a 30-second relaunch loop.
$script:LastStealthMarkerTs  = [long]0
# $true when the state saved at start was there and could not be restored
# (Import-WatchState), so which stealth markers the monitor before this one
# acted on is not known. In memory only.
$script:StateUnread          = $false
# Last line the stealth path wrote about a marker it could not use. In memory
# only, and declared here rather than on first use because reading an unset
# variable is an error under Set-StrictMode -Version Latest.
$script:LastStealthNotice    = $null
# Separate from the line above on purpose. Discovery runs once at startup, so
# "no AUMID" cannot change for the life of this process, and it is reported
# after a marker that read perfectly well -- which is exactly when the
# text-keyed suppression above has just been cleared.
$script:StealthNoAumidLogged = $false
# Test-ClaudeCodeActive says once when it finds no claude-code-sessions folder.
$script:NoCodeSessionDirLogged = $false
# When Sync-ClaudeDataPaths next looks for the app's folders again.
$script:NextDataPathSync = [datetime]::UtcNow.AddMinutes(10)
$script:StatePath           = Join-Path $WatchLogDir "watch-state.json"
# Seeded from the current state so the first poll does not report a transition
# that did not happen. The app only, as in the main loop. $null when the
# processes could not be listed: the first poll that lists them seeds it, and
# the error is said after the banner. Until 6.3.1 a list that could not be
# read was empty, and seeded "not running".
$script:ClaudeWasRunning     = $null
$script:SeedReadError        = $null
try {
    $script:ClaudeWasRunning = (@(Get-ClaudeDesktopProcess -Environment $script:ClaudeEnv).Count -gt 0)
} catch {
    $script:SeedReadError = $_
}

function ConvertFrom-WatchStateTime {
    <#
    .SYNOPSIS
        A time read back from watch-state.json, as UTC.
    .DESCRIPTION
        From 6.3.0 the file holds UTC times, written with ToString('o'), which
        ends them in Z. A file an older version wrote holds local times, with
        the offset they had then. Read with RoundtripKind a Z time comes back
        as UTC, and any other is read as local and converted, so a file from
        before the update does not move the cooldown or the escalation's
        window by the UTC offset. MinValue, which has no offset, stays
        MinValue. Windows PowerShell's ConvertFrom-Json hands the times back
        as strings; PowerShell 7's makes them DateTime already, and those are
        converted the same way.
    #>
    param($Value)
    if ($Value -is [datetime]) {
        $t = $Value
    } else {
        $t = [datetime]::Parse([string]$Value, [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::RoundtripKind)
    }
    if ($t.Ticks -eq 0) { return [datetime]::MinValue }
    if ($t.Kind -ne [System.DateTimeKind]::Utc) { $t = $t.ToUniversalTime() }
    return $t
}

function Import-WatchState {
    <#
    .SYNOPSIS
        Restores the cooldown, the fix history, a held-back repair and the
        runs Fix held back in a row from disk.
    .DESCRIPTION
        All of this state used to live only in memory. On restart, and this is
        a scheduled task that restarts on logon and after any crash, the
        cooldown, the three-fixes-in-thirty-minutes backoff and VmLogEverActive
        all reset to zero.

        That made the backoff defeatable by the very thing it guards against: a
        repair loop severe enough to take the machine down also cleared the
        record that would have stopped it running again.

        VmLogEverActive is seeded from the VM log's own write time rather than
        persisted, because the file is the authority. Without that, a machine
        that boots with an already-hung VM never sets the flag and therefore
        never runs staleness detection at all.

        The times are UTC, whatever version wrote the file: see
        ConvertFrom-WatchStateTime.

        A file that is not there is a first start, and is silent. One that
        is there but could not be read, or whose state could not be, fails
        closed: what was not restored starts empty, and a cooldown starts
        now, as after a repair that ran, so no automatic repair runs until
        it is over. That is said, once in the hour (Write-CheckError). A
        field of its own that cannot be read loses only that field, as
        before. Until 6.3.1 such a file read as no state at all, and one
        that could not be tested (Test-Path), or was empty, without a word.
        Such a file also sets $script:StateUnread: see
        Invoke-StealthRelaunch.
    #>
    $script:StateUnread = $false
    try {
        # Asked through Get-ClaudeItemOrNull: $null only when nothing is
        # there, which is a first start. Until 6.3.1 this was Test-Path,
        # which answers $false, with no error a caller sees, for a file
        # that is there but could not be tested: access denied on it or on
        # its folder, or a drive that is gone (measured 2026-11-23). Such a
        # file read as no state at all.
        if ($null -ne (Get-ClaudeItemOrNull -LiteralPath $script:StatePath)) {
            $raw = Get-Content -LiteralPath $script:StatePath -Raw -ErrorAction Stop
            # Save-WatchState always writes an object, so a file that is
            # empty, or holds anything else, is one whose write did not
            # finish, as when the machine went down during it, and nothing
            # in it can be trusted. Until 6.3.1 an empty file read as no
            # state at all.
            if ([string]::IsNullOrWhiteSpace($raw)) { throw 'the file is empty' }
            $st = $raw | ConvertFrom-Json -ErrorAction Stop
            if ($st -isnot [System.Management.Automation.PSCustomObject]) { throw 'the file does not hold a saved state' }
            if ($st.LastFixTime) {
                $script:LastFixTime = ConvertFrom-WatchStateTime -Value $st.LastFixTime
            }
            if ($st.FixCount) { $script:FixCount = [int]$st.FixCount }
            # Probed through PSObject.Properties, not read directly. Every
            # state file written before v6.1.0 lacks this field, and under
            # Set-StrictMode -Version Latest a bare $st.LastStealthMarkerTs
            # on such a file throws PropertyNotFound. That throw would go
            # to the catch below, abandoning the fix history restore that
            # follows it.
            if ($st.PSObject.Properties['LastStealthMarkerTs']) {
                try { $script:LastStealthMarkerTs = [long]$st.LastStealthMarkerTs }
                catch { $script:LastStealthMarkerTs = [long]0 }
            }
            foreach ($h in @($st.FixHistory)) {
                if (-not $h) { continue }
                $null = $script:FixHistory.Add((ConvertFrom-WatchStateTime -Value $h))
            }
            # Probed like LastStealthMarkerTs: a file written before 6.3.0
            # has neither field, and loads as no repair pending. In a try
            # of its own, so a bad value loses only this.
            $script:PendingFixHeldAt = [datetime]::MinValue
            $script:PendingFixTries  = 0
            $script:PendingFixRetriesHeld = 0
            if ($st.PSObject.Properties['PendingFixReason'] -and $st.PSObject.Properties['PendingFixSince'] -and
                $st.PendingFixReason) {
                try {
                    $since = ConvertFrom-WatchStateTime -Value $st.PendingFixSince
                    if ($since -ne [datetime]::MinValue) {
                        $script:PendingFixReason = [string]$st.PendingFixReason
                        $script:PendingFixSince  = $since
                    }
                } catch {
                    Write-CheckError -Check 'state' -ErrorRecord $_ `
                        -Text "STATE: The held-back repair saved in $($script:StatePath) could not be restored ({0}), so it is not tried again"
                }
                # The latest hold, the tries and the retries held back,
                # probed the same way and each in a try of its own. A
                # file written before they were kept has none of them,
                # and they load as none: the latest hold is then the
                # first, and no try is counted.
                if ($script:PendingFixReason) {
                    if ($st.PSObject.Properties['PendingFixHeldAt']) {
                        try { $script:PendingFixHeldAt = ConvertFrom-WatchStateTime -Value $st.PendingFixHeldAt }
                        catch { $script:PendingFixHeldAt = [datetime]::MinValue }
                    }
                    if ($st.PSObject.Properties['PendingFixTries']) {
                        try { $script:PendingFixTries = [math]::Max(0, [int]$st.PendingFixTries) }
                        catch { $script:PendingFixTries = 0 }
                    }
                    if ($st.PSObject.Properties['PendingFixRetriesHeld']) {
                        try { $script:PendingFixRetriesHeld = [math]::Max(0, [int]$st.PendingFixRetriesHeld) }
                        catch { $script:PendingFixRetriesHeld = 0 }
                    }
                }
            }
            # The runs Fix held back in a row, probed the same way and
            # each in a try of its own. A file written before they were
            # kept has neither, and loads as none held back.
            $script:FixHeldInARow   = 0
            $script:FixHeldLastTime = [datetime]::MinValue
            if ($st.PSObject.Properties['FixHeldInARow']) {
                try { $script:FixHeldInARow = [math]::Max(0, [int]$st.FixHeldInARow) }
                catch { $script:FixHeldInARow = 0 }
            }
            if ($st.PSObject.Properties['FixHeldLastTime']) {
                try { $script:FixHeldLastTime = ConvertFrom-WatchStateTime -Value $st.FixHeldLastTime }
                catch { $script:FixHeldLastTime = [datetime]::MinValue }
            }
            # Whether the cooldown was left by a hold, probed the same
            # way. A file written before it was kept has none, and loads
            # as not: a one-shot trigger in that cooldown is skipped, as
            # it was then.
            $script:CooldownFromHold = $false
            if ($st.PSObject.Properties['CooldownFromHold']) {
                try { $script:CooldownFromHold = ($st.CooldownFromHold -is [bool] -and $st.CooldownFromHold) }
                catch { $script:CooldownFromHold = $false }
            }
        }
    } catch {
        # Fail closed. The cooldown, the fix history and a held-back repair
        # are not all known. The history is what matters: three repairs in
        # thirty minutes stop a fourth (Test-PersistentFailure), and a
        # repair loop bad enough to take the machine down is the very thing
        # that leaves such a file behind. Starting as on a first start let
        # that loop run again at once. So a cooldown starts now, as after a
        # repair that ran, and no automatic repair runs until it is over,
        # Invoke-HeldBackRetry's included. A wait already longer, as an
        # escalation leaves, is kept. Whether a hold left the wait is not
        # known either, and is taken as not: a one-shot trigger in it is
        # then skipped, not kept to be tried once it is over. The history
        # itself is not made up: repairs that did not run here are not
        # counted toward the escalation.
        $now = [datetime]::UtcNow
        if ($script:LastFixTime -lt $now) { $script:LastFixTime = $now }
        $script:CooldownFromHold = $false
        # The stealth markers already acted on are not known either.
        $script:StateUnread = $true
        $waitMin = [math]::Round(($script:LastFixTime - $now).TotalMinutes + $Cooldown, 1)
        Write-CheckError -Check 'state' -ErrorRecord $_ `
            -Text "STATE: The state saved in $($script:StatePath) could not be restored in full ({0}); what was not restored starts empty, and a cooldown starts now, as after a repair: no automatic repair runs for $waitMin min"
    }

    # Seed from the logs themselves, not from the saved file.
    try {
        $seedAge = Get-ClaudeVmLogAge -LogDir $script:ClaudeLogDir
        if ($null -ne $seedAge -and $seedAge -lt 3600) { $script:VmLogEverActive = $true }
    } catch {
        Write-CheckError -Check 'VM log seed' -ErrorRecord $_ `
            -Text "The age of the VM logs could not be read at start ({0}), so the VM log staleness check waits until it sees a VM log written to"
    }
}

function Save-WatchState {
    <#
    .SYNOPSIS
        Persists the cooldown and whether a hold left it, the fix history, a
        held-back repair and the runs Fix held back in a row. Never throws;
        this is bookkeeping. A save that fails is logged (Write-CheckError).
    .DESCRIPTION
        Fields are only ever added, after the ones before them, so a monitor
        older than a field reads the file as it always did.
    #>
    try {
        $cutoff = [datetime]::UtcNow.AddHours(-6)
        # Trimmed on write. $script:FixHistory was appended to and never
        # trimmed, so it grew for the life of the daemon and every
        # Test-PersistentFailure call filtered a list that only got longer.
        $recent = @($script:FixHistory | Where-Object { $_ -gt $cutoff })
        $script:FixHistory.Clear()
        foreach ($r in $recent) { $null = $script:FixHistory.Add($r) }

        $payload = [ordered]@{
            LastFixTime         = $script:LastFixTime.ToString('o')
            FixCount            = $script:FixCount
            FixHistory          = @($recent | ForEach-Object { $_.ToString('o') })
            LastStealthMarkerTs = $script:LastStealthMarkerTs
            PendingFixReason    = $script:PendingFixReason
            PendingFixSince     = $script:PendingFixSince.ToString('o')
            PendingFixHeldAt    = $script:PendingFixHeldAt.ToString('o')
            PendingFixTries     = $script:PendingFixTries
            PendingFixRetriesHeld = $script:PendingFixRetriesHeld
            FixHeldInARow       = $script:FixHeldInARow
            FixHeldLastTime     = $script:FixHeldLastTime.ToString('o')
            CooldownFromHold    = [bool]$script:CooldownFromHold
        }
        [System.IO.File]::WriteAllText($script:StatePath,
            ($payload | ConvertTo-Json -Depth 3),
            (New-Object System.Text.UTF8Encoding($false)))
    } catch {
        # Still never throws: said, at most once an hour for the same error.
        Write-CheckError -Check 'state save' -ErrorRecord $_ `
            -Text "STATE: The state could not be saved to $($script:StatePath) ({0}); a restart would lose the cooldown, the fix history and a held-back repair"
    }
}

# Baselines are seeded lazily by Test-LogsForErrors: any stream it has not seen
# before starts at that file's current size. The eager loop that used to live
# here keyed on path alone, which no longer matches the (path, creation time)
# key the detector uses, and it would have had to be kept in sync by hand.

# -- Stealth-update relaunch ---------------------------------------------

function Write-StealthNotice {
    <#
    .SYNOPSIS
        Writes a line about an unusable marker, but only when it is not the
        line that was written last.
    .DESCRIPTION
        Everything on this path is asked once per poll for as long as Claude
        stays down, so an unreadable marker sitting on disk beside a closed app
        would write the same sentence 2,880 times a day. That is not a louder
        signal, it is a buried one: the watch log is also where the VirtioFS
        and HCS evidence lives.

        The message text is the key, so a DIFFERENT problem still gets its own
        line immediately. The record is cleared whenever the marker reads
        cleanly or goes away, so the same fault on a later marker is reported
        again rather than silently swallowed for the life of the daemon.
    #>
    param([string]$Message)

    if ($Message -eq $script:LastStealthNotice) { return }
    $script:LastStealthNotice = $Message
    Write-WatchLog $Message
}

function Get-StealthRelaunchMarker {
    <#
    .SYNOPSIS
        Reads the app's stealth-update marker. Returns $null when there is
        nothing to act on.
    .DESCRIPTION
        The marker is the only signal on this machine that distinguishes an
        update quit from a user quit. Both leave the same absence of a process,
        and only one of them should be undone.

        Written by the app at quit time, JSON:

            ts                epoch milliseconds
            windowVisible     bool
            windowMinimized   bool
            zOrderAnchorHwnd  optional
            navEntries        optional array
            navIndex          optional int

        Only ts is read here. The rest is the app's business: it consumes the
        file on next launch and restores the window and tab history from it.

        This function NEVER deletes the marker. Deleting it would leave the app
        with nothing to restore from, which throws away the entire point of
        launching it inside the honour window.

        Returns $null for absent, unreadable or unparseable. An unparseable
        marker is logged, because a marker that exists and cannot be read is
        the one case where the app went down through the stealth path and this
        monitor still will not act.
    .PARAMETER Path
        Marker file to read. Defaults to the discovered package path;
        overridable so the logic can be exercised without touching the real one.
    #>
    param([string]$Path = $script:StealthMarkerPath)

    # No path and no file are the ordinary cases, not faults. Both also clear
    # the repeat-suppression record: a marker that is gone cannot still be the
    # one whose complaint is being held down.
    if (-not $Path) { $script:LastStealthNotice = $null; return $null }
    # Whether it is there is asked by the error's identity: a marker that
    # could not be looked for is said, as one that could not be read is.
    # Until 6.3.1 Test-Path answered $false for it, and it was passed over
    # as absent without a word.
    $markerItem = $null
    try {
        $markerItem = Get-ClaudeItemOrNull -LiteralPath $Path
    } catch {
        Write-StealthNotice "STEALTH WARNING: marker at $Path could not be read: $(Get-ClaudeErrorText $_)"
        return $null
    }
    if ($null -eq $markerItem) { $script:LastStealthNotice = $null; return $null }

    $raw = $null
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    } catch {
        Write-StealthNotice "STEALTH WARNING: marker at $Path could not be read: $($_.Exception.Message)"
        return $null
    }
    if (-not $raw) {
        Write-StealthNotice "STEALTH WARNING: marker at $Path is empty"
        return $null
    }

    $obj = $null
    try {
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
    } catch {
        Write-StealthNotice "STEALTH WARNING: marker at $Path is not valid JSON: $($_.Exception.Message)"
        return $null
    }
    if (-not $obj) {
        Write-StealthNotice "STEALTH WARNING: marker at $Path parsed to nothing"
        return $null
    }

    # PSObject.Properties rather than $obj.ts. Under Set-StrictMode -Version
    # Latest a missing property throws, and "the app changed its payload" is
    # exactly the case this branch exists to report rather than crash on.
    $tsProp = $null
    try { $tsProp = $obj.PSObject.Properties['ts'] } catch { $tsProp = $null }
    if (-not $tsProp) {
        Write-StealthNotice "STEALTH WARNING: marker at $Path has no 'ts' field"
        return $null
    }

    $tsRaw = [long]0
    try { $tsRaw = [long]$tsProp.Value } catch {
        Write-StealthNotice "STEALTH WARNING: marker at $Path has a non-numeric 'ts': $($tsProp.Value)"
        return $null
    }

    # Epoch arithmetic done by hand rather than through
    # DateTimeOffset::FromUnixTimeMilliseconds, which throws on out-of-range
    # input. Here a nonsense ts should be reported, not raised.
    $epochUtc = [datetime]::new(1970, 1, 1, 0, 0, 0, [System.DateTimeKind]::Utc)
    $writtenUtc = $null
    try { $writtenUtc = $epochUtc.AddMilliseconds([double]$tsRaw) } catch {
        Write-StealthNotice "STEALTH WARNING: marker at $Path has an out-of-range 'ts': $tsRaw"
        return $null
    }

    # A ts in seconds instead of milliseconds lands in 1970 and would otherwise
    # read as a very stale marker, which is the same verdict for the wrong
    # reason. Say so instead.
    if ($writtenUtc.Year -lt 2000) {
        Write-StealthNotice "STEALTH WARNING: marker at $Path has an implausible 'ts' ($tsRaw resolves to $($writtenUtc.ToString('u')))"
        return $null
    }

    $ageSeconds = ([datetime]::UtcNow - $writtenUtc).TotalSeconds

    # It read cleanly, so nothing is being held down any more.
    $script:LastStealthNotice = $null

    return [pscustomobject]@{
        Path       = $Path
        TsRaw      = $tsRaw
        Written    = $writtenUtc.ToLocalTime()
        AgeMinutes = ($ageSeconds / 60.0)
        AgeSeconds = $ageSeconds
    }
}

function Invoke-StealthRelaunch {
    <#
    .SYNOPSIS
        Starts Claude again if, and only if, it went down through the
        stealth-update path. Returns $true when a launch was attempted.
    .DESCRIPTION
        Called on every poll on which Claude is not running, not only on the
        moment it disappears. A monitor that restarts, or that starts at logon
        after the update already happened, sees no transition at all, and the
        act-once guard makes re-evaluating each poll free.

        Five ways this declines:

          - the operator passed -NoStealthRelaunch
          - there is no marker, so this was a crash or a deliberate quit, and
            neither is ours to undo
          - the marker is older than -MaxStealthMarkerAgeMinutes, so the
            machine was off or the user chose to stay closed
          - the marker's ts is one already decided about
          - the state saved at start could not be restored, and the marker
            was written before this monitor started: the monitor before it
            may have decided about it already (since 6.3.1)

        The ts is recorded BEFORE the launch, not after. If Start-Process
        throws, the record is already down and the next poll will not retry the
        same marker thirty seconds later, forever. A genuine second update
        writes a new marker with a new ts and is acted on normally.

        The stale branch records the ts too. It is still "do nothing" as far as
        the app is concerned; the record is only there so a stale marker sitting
        on disk beside a closed app does not write the same line to the watch
        log 2,880 times a day.
    #>
    param()

    if ($NoStealthRelaunch) { return $false }

    $marker = Get-StealthRelaunchMarker
    if (-not $marker) { return $false }

    if ($marker.TsRaw -eq $script:LastStealthMarkerTs) { return $false }

    if ($marker.AgeMinutes -gt $MaxStealthMarkerAgeMinutes) {
        Write-WatchLog ("Stale stealth marker, not resurrecting: written {0}, {1:N1} min old, limit {2} min." -f `
            $marker.Written.ToString('yyyy-MM-dd HH:mm:ss'), $marker.AgeMinutes, $MaxStealthMarkerAgeMinutes)
        $script:LastStealthMarkerTs = $marker.TsRaw
        Save-WatchState
        return $false
    }

    # The ts of the last marker decided about is kept in the state file. When
    # that could not be restored, a marker written before this monitor
    # started may be one the monitor before it already relaunched for, and
    # starting Claude again could undo a quit that came after. So it is not
    # acted on: it is recorded as decided, and said. A marker written after
    # the start is a new one, and is acted on as ever. Until 6.3.1 the ts read
    # as 0, and such a marker as new.
    if ($script:StateUnread -and $marker.AgeSeconds -gt ([datetime]::UtcNow - $script:StartTime).TotalSeconds) {
        Write-WatchLog ("Stealth marker written {0}, before this monitor started, not acted on: the saved state that records whether it was could not be read." -f `
            $marker.Written.ToString('yyyy-MM-dd HH:mm:ss'))
        $script:LastStealthMarkerTs = $marker.TsRaw
        Save-WatchState
        return $false
    }

    # Checked BEFORE the ts is recorded, not after. Nothing was decided here:
    # discovery runs once at startup, so a monitor restarted after whatever hid
    # the package should get to try again on a marker that is still fresh.
    #
    # Said once per process rather than through Write-StealthNotice, because
    # the marker itself read cleanly, and a clean read is precisely what clears
    # that suppression. A fresh marker is re-read every poll for ten minutes,
    # so without a flag of its own this would be twenty identical lines.
    if (-not $script:ClaudeEnv.Aumid) {
        if (-not $script:StealthNoAumidLogged) {
            $script:StealthNoAumidLogged = $true
            Write-WatchLog "Stealth marker is fresh but no AUMID was discovered, so there is nothing to launch."
        }
        return $false
    }

    # Record first, launch second. See the description.
    $script:LastStealthMarkerTs = $marker.TsRaw
    Save-WatchState

    # shell:AppsFolder activation, which is what Fix-ClaudeDesktop uses for the
    # same job. Launching app\Claude.exe out of WindowsApps directly produces a
    # loose instance with a duplicate taskbar icon, and the elevated
    # \Claude\LaunchClaudeAdmin task does exactly that on purpose, to keep its
    # admin token. Neither cost is worth paying for a restart the user is meant
    # not to notice, and coming back elevated is not the state the app was in.
    #
    # Both values below are format arguments. Interpolating the AUMID into the
    # format string instead would be the same defect Write-WatchLog documents:
    # any brace in an interpolated value is parsed as a placeholder and throws.
    Write-WatchLog ("Stealth-update relaunch: marker {0:N0}s old, starting {1}" -f `
        $marker.AgeSeconds, $script:ClaudeEnv.Aumid)
    try {
        Start-Process $script:ClaudeEnv.Aumid -ErrorAction Stop
        Write-WatchLog "Stealth-update relaunch: launch requested."
    } catch {
        Write-WatchLog "Stealth-update relaunch FAILED: $($_.Exception.Message)"
    }
    return $true
}

# -- Detection functions -------------------------------------------------

function Test-StartupGracePeriod {
    <#
    .SYNOPSIS
        Returns $true if the monitor just started and should skip heuristic checks.
        Grace period: 180 seconds. This prevents matching pre-existing event log
        entries or stale logs from before the monitor was running.
    #>
    return ([datetime]::UtcNow - $script:StartTime).TotalSeconds -lt 180
}

function Test-PersistentFailure {
    <#
    .SYNOPSIS
        Returns $true if the fix has been attempted 3+ times in 30 minutes,
        indicating an unrecoverable error (e.g., HCS JSON corruption).
    #>
    # UTC, as FixHistory is.
    $cutoff = [datetime]::UtcNow.AddMinutes(-30)
    $recentFixes = @($script:FixHistory | Where-Object { $_ -gt $cutoff })
    return $recentFixes.Count -ge 3
}

function Find-LogErrorPattern {
    <#
    .SYNOPSIS
        Returns the first pattern found inside an [error] entry of a chunk of
        log text, or $null.
    .DESCRIPTION
        The app writes one entry per line, such as
        "2026-11-18 15:26:11 [info] [remote-bash] text", and puts an error's
        stack trace and message object on the lines after it, with no
        timestamp. A line that starts with a timestamp opens a new entry; any
        other line belongs to the entry above it and shares its level.

        A pattern counts only in an entry whose level is error and whose
        component does not start with "remote". The remote-bash component runs
        remote sessions' commands in the VM, and its "RPC error -1" lines are a
        probe's normal answer. Lines before the chunk's first timestamp belong
        to an entry whose header was read on an earlier poll, so they have no
        level here and never count.
    #>
    param(
        [AllowEmptyString()][string]$Text,
        [string[]]$Patterns
    )
    if (-not $Text) { return $null }
    $level = ''
    $component = ''
    foreach ($line in ($Text -split "`r?`n")) {
        $entry = [regex]::Match($line,
            '^\d{4}[-/]\d{2}[-/]\d{2}[ T][0-9:.]+Z?(?:\s+\[(?<level>[A-Za-z]+)\])?(?:\s*\[(?<component>[^\]]*)\])?')
        if ($entry.Success) {
            $level = $entry.Groups['level'].Value
            $component = $entry.Groups['component'].Value
        }
        if ($level -ne 'error') { continue }
        if ($component -like 'remote*') { continue }
        foreach ($pattern in $Patterns) {
            if ($line.IndexOf($pattern, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                return $pattern
            }
        }
    }
    return $null
}

function Test-LogsForErrors {
    <#
    .SYNOPSIS
        Scans what the app's VM and main logs gained since the last poll for an
        [error] entry carrying a VM failure signature. Returns the pattern
        found, or $null if clean.
    .DESCRIPTION
        One hit starts a repair with no confirming second poll, so the files
        are limited by $LogTriggerFilePattern and the lines by
        Find-LogErrorPattern. $LegacyErrorPatterns records what the logs hold.
    #>
    # A folder that is not there gives PathNotFound with ObjectNotFound and
    # ItemNotFoundException, for a missing parent too (measured 2026-11-22):
    # no logs yet. Any other error is said through the hourly throttle and
    # the scan ends before the baselines are pruned. Until 6.3.1 the folder
    # was listed with SilentlyContinue after a Test-Path that answers false
    # for a folder it may not read, and a listing that failed read as a
    # folder with no logs: no trigger could fire, and every baseline was
    # dropped, so the errors written meanwhile were never scanned.
    try {
        $listed = @(Get-ChildItem -LiteralPath $ClaudeLogDir -Filter '*.log' -File -ErrorAction Stop)
    } catch {
        if (Test-ClaudeNotFoundError -ErrorRecord $_ -ErrorId 'PathNotFound' -ExceptionType 'System.Management.Automation.ItemNotFoundException') {
            return $null
        }
        Write-CheckError -Check 'Claude logs' -ErrorRecord $_ `
            -Text "LOG WARNING: The Claude log folder $ClaudeLogDir could not be listed ({0}), so no log trigger can fire until it can"
        return $null
    }
    $logFiles = @($listed | Where-Object { $_.Name -match $LogTriggerFilePattern })
    $seenKeys = New-Object 'System.Collections.Generic.HashSet[string]'

    foreach ($logFile in $logFiles) {
        $path = $logFile.FullName
        $currentSize = $logFile.Length

        # Keyed on path AND creation time, not path alone.
        #
        # Electron shift-rotates content between paths: main.log becomes
        # main1.log, main1 becomes main2, and so on across four families and
        # 36 files. Keyed on path alone, a rotation either reset the baseline
        # to 0 and re-scanned multi-MB of already-seen content, or compared
        # against a baseline belonging to entirely different content. Either
        # way any historical error line could re-fire an auto-fix.
        #
        # Creation time changes when the content moves, so the pair identifies
        # the actual stream rather than the filename it currently occupies.
        $key = $path + '|' + $logFile.CreationTimeUtc.ToString('o')
        $null = $seenKeys.Add($key)

        $isKnown  = $script:LogBaselines.ContainsKey($key)
        $baseline = 0
        if ($isKnown) { $baseline = $script:LogBaselines[$key] }

        # A file this monitor has never seen starts at its CURRENT size, not at
        # zero. Starting at zero would scan the entire backlog on first sight
        # and treat months-old errors as new.
        if (-not $isKnown) {
            $script:LogBaselines[$key] = $currentSize
            continue
        }

        # Truncated in place. Resync to the current size rather than zero.
        if ($currentSize -lt $baseline) {
            $script:LogBaselines[$key] = $currentSize
            continue
        }

        # Skip if no new content
        if ($currentSize -le $baseline) { continue }

        # Read only new content
        try {
            $stream = $null
            $reader = $null
            try {
                $stream = [System.IO.FileStream]::new(
                    $path,
                    [System.IO.FileMode]::Open,
                    [System.IO.FileAccess]::Read,
                    [System.IO.FileShare]::ReadWrite)
                $stream.Position = $baseline
                $reader = [System.IO.StreamReader]::new($stream)
                $newContent = $reader.ReadToEnd()
            } finally {
                # A close that fails has nothing left to release, and what
                # was read stands.
                if ($reader) { try { $reader.Close() } catch { $null = $_ } }
                if ($stream) { try { $stream.Close() } catch { $null = $_ } }
            }

            # Advance the baseline only after a SUCCESSFUL read.
            #
            # It used to advance before the read, so a read that failed for any
            # reason (a locked file, a mid-rotation moment) permanently
            # discarded that window of log content without ever examining it,
            # and the failure was swallowed by the catch below.
            $script:LogBaselines[$key] = $currentSize

            $hit = Find-LogErrorPattern -Text $newContent -Patterns $LegacyErrorPatterns
            if ($hit) { return $hit }
        } catch {
            # Read failed. The baseline stays put, so this window is retried on
            # the next poll rather than silently skipped. Said, at most once an
            # hour for the same error: a log that can never be read is a check
            # that never runs, and until 6.3.0 nothing said so.
            Write-CheckError -Check 'Claude logs' -ErrorRecord $_ `
                -Text "LOG WARNING: The Claude log $path could not be read ({0}); what it gained since the last read is read again at the next poll"
        }
    }

    # Drop baselines for streams that no longer exist. This hashtable was never
    # pruned, so it grew for the entire life of the daemon, and with the key now
    # including creation time every rotation would otherwise add a fresh entry
    # and keep the old one forever.
    foreach ($staleKey in @($script:LogBaselines.Keys | Where-Object { -not $seenKeys.Contains($_) })) {
        $script:LogBaselines.Remove($staleKey)
    }
    return $null
}

function Test-ServiceHealth {
    <#
    .SYNOPSIS
        Checks if CoworkVMService is running while Claude is active.
        Returns $true if healthy, $false if unhealthy.
        Now requires 2 CONSECUTIVE failures to avoid transient blips.
    #>
    # The app, not any "claude". The service backs the app's workspace, and a
    # Claude Code CLI process alone neither needs it nor means the app is up.
    #
    # A read that failed is said, at most once an hour, and the check passes
    # this poll with its count as it was: a repair is never started on it.
    # Until 6.3.1 a process list that could not be read was empty, and a
    # service that could not be read was "not installed".
    try {
        $claude = @(Get-ClaudeDesktopProcess -Environment $script:ClaudeEnv)
    } catch {
        Write-CheckError -Check 'service health' -ErrorRecord $_ `
            -Text "Claude's processes could not be listed ({0}), so whether $ServiceName is needed is not known and nothing is done about it"
        return $true
    }
    if ($claude.Count -eq 0) {
        $script:ServiceDownCount = 0
        return $true
    }

    try {
        $svc = Get-ClaudeServiceOrNull -Name $ServiceName
    } catch {
        Write-CheckError -Check 'service health' -ErrorRecord $_ `
            -Text "$ServiceName could not be read ({0}), so whether it is running is not known and nothing is done about it"
        return $true
    }
    if (-not $svc) {
        $script:ServiceDownCount = 0
        return $true   # Service not installed -- not a Cowork setup
    }

    if ($svc.Status -ne "Running") {
        $script:ServiceDownCount++
        if ($script:ServiceDownCount -ge 2) {
            $script:ServiceDownCount = 0
            return $false
        }
        Write-WatchLog "Service not running (check $($script:ServiceDownCount)/2, waiting to confirm)"
        return $true   # First failure -- wait one more cycle
    }

    $script:ServiceDownCount = 0
    return $true
}

function Get-MatchingWinEvent {
    <#
    .SYNOPSIS
        Get-WinEvent -FilterHashtable with -ErrorAction Stop, where nothing to
        read is no error: returns the events, or nothing when none matches,
        the log does not exist or its provider is not installed. Any other
        error is thrown, for the check to name.
    .DESCRIPTION
        Get-WinEvent says there is nothing to read with an ObjectNotFound
        error whose id starts NoMatchingEventsFound, NoMatchingLogsFound or
        NoMatchingProvidersFound (measured on Windows PowerShell 5.1). On a
        machine without Hyper-V, or before the Cowork service has logged
        anything, that is the usual answer. Until 6.3.0 every read here used
        SilentlyContinue, which hid every other error with them: a check that
        could not read its log looked like one that found nothing.
    #>
    param([hashtable]$Filter, [int]$MaxEvents = 0)
    try {
        if ($MaxEvents -gt 0) { return @(Get-WinEvent -FilterHashtable $Filter -MaxEvents $MaxEvents -ErrorAction Stop) }
        return @(Get-WinEvent -FilterHashtable $Filter -ErrorAction Stop)
    } catch {
        if ([string]$_.FullyQualifiedErrorId -match '^NoMatching(Events|Logs|Providers)Found(,|$)') { return }
        throw
    }
}

function Test-EventLogErrors {
    <#
    .SYNOPSIS
        Checks Windows Event Log for Claude-specific errors.
        TIGHTENED in v3.0:
        - VMMS check now requires "claude" in the message (no more "failed"/"unexpected" wildcards)
        - Requires 2 consecutive hits before triggering
        - Skipped during startup grace period
        A log that cannot be read is logged, naming the log and the error, at
        most once an hour for the same error (Write-CheckError); a log that
        does not exist, or has nothing in the window, is not an error
        (Get-MatchingWinEvent).
    #>
    # Skip during startup grace period (avoids pre-existing events)
    if (Test-StartupGracePeriod) { return $null }

    $lookback = [math]::Max($PollInterval, 60)
    # Used only as Get-WinEvent's StartTime below: the UTC time as round-trip
    # text, which ends in Z. Measured on Windows PowerShell 5.1, on this
    # machine's log, in summer time and across the hour the clocks went back:
    # given a [datetime], Get-WinEvent reads it as local time whatever its
    # Kind, so a UTC time moves the window by the UTC offset, and a local
    # time is wrong in the hour the clocks go back, which happens twice. The
    # text was exact in both. Until 6.3.0 this was the local time.
    $since = [datetime]::UtcNow.AddSeconds(-$lookback).ToString('o')
    $foundIssue = $null

    # The four checks below examine the WHOLE event message.
    #
    # They used to take only the first line, via (.Message -split "`n")[0].
    # Hyper-V puts the summary on line one and the operation detail and error
    # code on the lines after it, so matching "Plan9", "virtio", "cowork" or an
    # HRESULT against line one alone missed the part of the message that
    # actually identifies the fault.
    #
    # Only the first line is used when reporting, to keep the watch log
    # readable.
    $firstLine = {
        param([string]$Text)
        $parts = @($Text -split "`r?`n" | Where-Object { $_.Trim() })
        if ($parts.Count -gt 0) { return $parts[0].Trim() }
        return $Text
    }

    # Check 1: CoworkVMService errors -- only match VirtioFS patterns
    try {
        $evFilter = @{
            LogName      = "Application"
            ProviderName = "CoworkVMService"
            Level        = 2  # Error
            StartTime    = $since
        }
        $events = Get-MatchingWinEvent -Filter $evFilter -MaxEvents 1
        if ($events) {
            $msg = "$($events[0].Message)"
            foreach ($pattern in $LegacyErrorPatterns) {
                if ($msg -match [regex]::Escape($pattern)) {
                    $foundIssue = "CoworkVMService: $(& $firstLine $msg)"
                    break
                }
            }
        }
    } catch {
        Write-CheckError -Check 'Event Log: Application' -ErrorRecord $_ `
            -Text "EVENT LOG: The Application log could not be checked for CoworkVMService errors ({0})"
    }

    # Check 2: Hyper-V Worker -- must mention claude/Plan9/virtio/shared memory
    if (-not $foundIssue) {
        try {
            $hvWorkerFilter = @{
                LogName   = "Microsoft-Windows-Hyper-V-Worker-Admin"
                Level     = @(1, 2)  # Critical, Error only (dropped warnings)
                StartTime = $since
            }
            $hvEvents = Get-MatchingWinEvent -Filter $hvWorkerFilter -MaxEvents 1
            if ($hvEvents) {
                $msg = "$($hvEvents[0].Message)"
                if ($msg -match "claude" -or $msg -match "Plan9" -or $msg -match "virtio" -or $msg -match "shared memory") {
                    $foundIssue = "Hyper-V Worker: $(& $firstLine $msg)"
                }
            }
        } catch {
            Write-CheckError -Check 'Event Log: Hyper-V Worker' -ErrorRecord $_ `
                -Text "EVENT LOG: The Microsoft-Windows-Hyper-V-Worker-Admin log could not be checked ({0})"
        }
    }

    # Check 3: Hyper-V VMMS -- MUST mention "claude" (no more generic "failed" matching)
    if (-not $foundIssue) {
        try {
            $vmmsFilter = @{
                LogName   = "Microsoft-Windows-Hyper-V-VMMS-Admin"
                Level     = @(1, 2)  # Critical, Error
                StartTime = $since
            }
            $vmmsEvents = Get-MatchingWinEvent -Filter $vmmsFilter -MaxEvents 1
            if ($vmmsEvents) {
                $msg = "$($vmmsEvents[0].Message)"
                # TIGHTENED: must specifically mention claude or cowork
                if ($msg -match "claude" -or $msg -match "cowork") {
                    $foundIssue = "Hyper-V VMMS: $(& $firstLine $msg)"
                }
            }
        } catch {
            Write-CheckError -Check 'Event Log: Hyper-V VMMS' -ErrorRecord $_ `
                -Text "EVENT LOG: The Microsoft-Windows-Hyper-V-VMMS-Admin log could not be checked ({0})"
        }
    }

    # Check 4: HCS Compute log -- dedicated channel for Host Compute Service
    if (-not $foundIssue) {
        try {
            $hcsFilter = @{
                LogName   = "Microsoft-Windows-Hyper-V-Compute-Admin"
                Level     = @(1, 2)  # Critical, Error
                StartTime = $since
            }
            $hcsEvents = Get-MatchingWinEvent -Filter $hcsFilter -MaxEvents 1
            if ($hcsEvents) {
                $msg = "$($hcsEvents[0].Message)"
                if ($msg -match "claude" -or $msg -match "cowork" -or $msg -match "failed to create" -or $msg -match "HcsWaitForOperationResult") {
                    $foundIssue = "HCS Compute: $(& $firstLine $msg)"
                }
            }
        } catch {
            Write-CheckError -Check 'Event Log: HCS Compute' -ErrorRecord $_ `
                -Text "EVENT LOG: The Microsoft-Windows-Hyper-V-Compute-Admin log could not be checked ({0})"
        }
    }

    # Require 2 consecutive event log hits before triggering
    if ($foundIssue) {
        $script:EventLogHitCount++
        if ($script:EventLogHitCount -ge 2) {
            $script:EventLogHitCount = 0
            return $foundIssue
        }
        Write-WatchLog "Event log hit (check $($script:EventLogHitCount)/2): $foundIssue"
        return $null
    }

    $script:EventLogHitCount = 0
    return $null
}

function Test-WorkspaceStorage {
    <#
    .SYNOPSIS
        Checks that the VM's own storage is readable and that its volume has
        room. Returns $null when healthy, or a description when not.
    .DESCRIPTION
        Service state was the only thing this monitor checked, and a hung
        cowork-svc still reports Running to the SCM. That is precisely the
        failure this toolkit is named for, and nothing was looking at it.

        Two things ARE observable from the host:

        1. Whether the VM cache directory and the VHDX inside it can still be
           enumerated and opened. A directory that has become unreadable, or a
           sessiondata.vhdx that has vanished while the service claims to be
           running, is a concrete fault rather than an inference.

        2. Whether the volume holding the VM image has space. The image is
           roughly 13 GB and grows; a full volume produces mount failures that
           look like everything else.

        The virtiofs mounts themselves live inside the guest at /sessions/...
        and are not reachable from the host, so this deliberately does not
        claim to test them. It tests what it can actually see.

        Reports only. It does not trigger an auto-fix, because neither a full
        disk nor a missing file is something a VM restart repairs. What it
        returns, a VM cache folder that could not be read (naming the folder
        and the error) or a volume with less than 2 GB free, the main loop's
        Check 4b logs as a STORAGE WARNING, each different text at most once
        an hour and with no cap on how many (Test-StorageWarningDue). Until
        6.3.1 that line was written on every poll, about 120 times an hour.
        A free space that cannot be read is logged here, at most once an hour
        for the same error and no more than 3 different errors in an hour
        (Write-CheckError, key 'storage'); until 6.3.0 it was passed over.
    #>
    if (Test-StartupGracePeriod) { return $null }

    $cachePath = $script:ClaudeEnv.VmCachePath
    if (-not $cachePath) { return $null }
    # Whether it is there is asked by the error's identity: a folder that
    # could not be read is the fault this check is here to report. Until
    # 6.3.1 Test-Path answered $false for it, which read as no VM cache.
    # Both lines name the folder: the error Get-Item gives for a Deny,
    # "Access is denied", does not.
    $cacheItem = $null
    try {
        $cacheItem = Get-ClaudeItemOrNull -LiteralPath $cachePath
    } catch {
        return "The VM cache directory $cachePath could not be read ($(Get-CheckErrorText -ErrorRecord $_))"
    }
    if ($null -eq $cacheItem) { return $null }

    # 1. Readability.
    try {
        $null = @(Get-ChildItem $cachePath -Force -ErrorAction Stop | Select-Object -First 1)
    } catch {
        return "The VM cache directory $cachePath could not be read ($(Get-CheckErrorText -ErrorRecord $_))"
    }

    # 2. Free space on the volume that holds it.
    try {
        $driveLetter = (Get-Item $cachePath -ErrorAction Stop).PSDrive.Name
        if ($driveLetter) {
            $psd = Get-PSDrive $driveLetter -ErrorAction Stop
            if ($psd -and $null -ne $psd.Free) {
                $freeGb = [math]::Round($psd.Free / 1GB, 1)
                if ($psd.Free -lt 2GB) {
                    return "Only ${freeGb} GB free on ${driveLetter}: the VM image needs room to grow and mounts fail without it"
                }
            }
        }
    } catch {
        Write-CheckError -Check 'storage' -ErrorRecord $_ `
            -Text "STORAGE WARNING: The free space on the VM image's volume could not be read ({0}), so a full volume would not be seen"
    }

    return $null
}

function Test-VmLogStaleness {
    <#
    .SYNOPSIS
        Checks if the VM's node log has gone stale while Cowork should be active.
        TIGHTENED in v3.0:
        - Stale threshold: 300s (was 120s) -- 5 minutes of silence
        - Consecutive checks: 5 (was 3) -- 150s of confirmed stale
        - Only triggers if the log was PREVIOUSLY active (VmLogEverActive)
          This prevents false positives when user is in Chat mode (no Cowork)
    #>
    # Staleness of the freshest VM log, read now (Get-ClaudeVmLogAge). Reading
    # one chosen file, as this used to, measured whichever log was newest when
    # the monitor started: once a file last written in March, and on app
    # 6.18355 possibly coworkd.log while the boot and command lines go to
    # cowork_vm_node.log.
    #
    # main.log is deliberately NOT a candidate. It is the Electron main process
    # log and it can sit frozen for a day at a time on a completely healthy
    # machine, which makes staleness there a guaranteed false positive.
    #
    # A VM log that could not be read is said, at most once an hour, and the
    # count stays as it was. Until 6.3.1 it read as no VM log, and the count
    # started again from 0.
    try {
        if ($null -eq (Get-ClaudeVmLogAge -LogDir $script:ClaudeLogDir)) {
            $script:VmLogStaleCount = 0
            return $null
        }
    } catch {
        Write-CheckError -Check 'VM log staleness' -ErrorRecord $_ `
            -Text "STALENESS WARNING: The VM logs could not be read ({0}), so whether one has gone stale is not known and nothing is done about it"
        return $null
    }

    # Skip during startup grace period
    if (Test-StartupGracePeriod) { return $null }

    # Skip staleness check when a Cowork session is active
    # During API-wait periods, the VM log stops updating but the session is alive.
    # Without this, Watch would flag "VM log stale" and eventually auto-fix,
    # killing the active Cowork research mid-conversation.
    if (Test-CoworkSessionActive) {
        $script:VmLogStaleCount = 0
        return $null
    }

    try {
        $staleSec = Get-ClaudeVmLogAge -LogDir $script:ClaudeLogDir
        if ($null -eq $staleSec) { return $null }

        # Track if the log has ever been active during this monitor session
        if ($staleSec -lt 60) {
            $script:VmLogEverActive = $true
            $script:VmLogStaleCount = 0
            return $null
        }

        # Only check staleness if the log was previously active
        # (Prevents false triggers when user is in Chat mode, not Cowork)
        if (-not $script:VmLogEverActive) {
            return $null
        }

        if ($staleSec -gt 300) {
            $script:VmLogStaleCount++
            if ($script:VmLogStaleCount -ge 5) {
                $script:VmLogStaleCount = 0
                return "VM log stale for $([math]::Round($staleSec))s, VM may be hung"
            }
            if ($script:VmLogStaleCount -eq 1) {
                Write-WatchLog "VM log stale ($([math]::Round($staleSec))s), monitoring (check 1/5)"
            }
        } else {
            $script:VmLogStaleCount = 0
        }
    } catch {
        Write-CheckError -Check 'VM log staleness' -ErrorRecord $_ `
            -Text "STALENESS WARNING: The VM log staleness check failed ({0})"
    }
    return $null
}

function Invoke-W32tm {
    <#
    .SYNOPSIS
        Runs w32tm.exe in a job with a timeout. Returns an object with its
        output (Text) and exit code (ExitCode), or $null when it did not
        finish in time. Throws, naming why, when it could not be run.
    .DESCRIPTION
        w32tm makes network calls, and this runs inside the monitor's
        single-threaded poll loop: unbounded, a slow time server stalls every
        check for as long as the network takes to give up. Until 6.3.0 only
        the measurement ran with a timeout, and the forced resync ran inline.

        A job whose process cannot be started does not throw at Start-Job: it
        ends Failed, with no output, and why comes only from Receive-Job, as
        an error. A program that cannot be found ends the job with no exit
        code and the error beside it. Both throw here, naming the first
        error, or the job's state when there is none. Until 6.3.0 both
        returned $null, as a timeout does, and the clock check logged that
        w32tm "did not answer within 15s" each time it ran.

        -W32tmPath is there for the tests, which point it at a stand-in.
    #>
    param(
        [string[]]$Arguments,
        [int]$TimeoutSeconds = 15,
        [string]$W32tmPath = "$env:SystemRoot\System32\w32tm.exe"
    )
    # What an error names: the command and its first argument.
    $what = "w32tm $(@($Arguments)[0])"
    try {
        $job = Start-Job -ScriptBlock {
            param($p, $a)
            $text = & $p @a 2>&1 | Out-String
            [pscustomobject]@{ Text = $text; ExitCode = $LASTEXITCODE }
        } -ArgumentList $W32tmPath, $Arguments -ErrorAction Stop
    } catch {
        throw "$what could not be run: $(Get-CheckErrorText -ErrorRecord $_)"
    }
    try {
        if (-not (Wait-Job $job -Timeout $TimeoutSeconds)) {
            # Clean-up of a job that has not answered: one that cannot be
            # stopped ends with the monitor. The caller says the timeout.
            Stop-Job $job -ErrorAction SilentlyContinue
            return $null
        }
        $jobErrors = $null
        $result = @(Receive-Job $job -ErrorAction SilentlyContinue -ErrorVariable jobErrors) |
            Where-Object { $null -ne $_ -and $_.PSObject.Properties['ExitCode'] } | Select-Object -Last 1
        if ($null -eq $result -or $null -eq $result.ExitCode) {
            $first = @($jobErrors) | Where-Object { $null -ne $_ } | Select-Object -First 1
            if ($first) {
                $why = Get-CheckErrorText -ErrorRecord $first
            } else {
                $why = "the job ended $($job.State) with no exit code"
            }
            throw "$what could not be run: $why"
        }
        return $result
    } finally {
        # Clean-up: the result, or why there is none, is in hand, and a job
        # that cannot be removed ends with the monitor.
        Remove-Job $job -Force -ErrorAction SilentlyContinue
    }
}

function Measure-ClockOffset {
    <#
    .SYNOPSIS
        The host clock's offset from time.windows.com in seconds, or $null
        when it could not be measured.
    .DESCRIPTION
        Get-W32tmOffsetSeconds reads the text as one string. Until a fix
        before 6.3.0 the raw output, an array of lines, was matched directly,
        which filters rather than fills $Matches, and the drift check never
        ran; and until 6.3.0 the pattern took a decimal point only, so on a
        machine whose locale writes a comma it never measured anything.

        $null means only that w32tm did not answer within 15s. That is
        logged as a skipped check at most once an hour (Test-CheckErrorDue):
        the check runs every 5 minutes, and a time server that does not
        answer once most often does not answer the next time either. Until
        6.3.0 it was logged on every check.

        Text with no offset in it, as when the time server answers with an
        error ("12:00:00, error: 0x800705B4"), throws, naming that line
        without the time it starts with, so the same error reads the same
        each time and Test-TimeSyncHealth logs it as a TIME WARNING. Until
        6.3.0 it returned $null with no line, as if the check were skipped.
        So does a w32tm that could not be run at all (Invoke-W32tm).
    #>
    param(
        # The measurement after a forced resync. When it does not answer the
        # check is not skipped: Test-TimeSyncHealth warns that the clock could
        # not be measured after the resync. Until 6.3.0 this logged "skipping
        # the clock-drift check" then too, beside that warning.
        [switch]$AfterResync
    )
    $r = Invoke-W32tm -Arguments '/stripchart', '/computer:time.windows.com', '/dataonly', '/samples:1'
    if ($null -eq $r) {
        if ($AfterResync) {
            Write-WatchLog "w32tm stripchart did not answer within 15s after the resync"
        } elseif (Test-CheckErrorDue -Check 'clock timeout' -Message 'stripchart' -Now ([datetime]::UtcNow)) {
            Write-WatchLog "w32tm stripchart did not answer within 15s; skipping the clock-drift check"
        }
        return $null
    }
    $offset = Get-W32tmOffsetSeconds -Text ([string]$r.Text)
    if ($null -eq $offset) {
        $lines = @(([string]$r.Text) -split "`r?`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        $said = @($lines | Where-Object { $_ -match 'error' }) | Select-Object -First 1
        if (-not $said) { $said = $lines | Select-Object -Last 1 }
        if (-not $said) { $said = "no output, exit $($r.ExitCode)" }
        throw "w32tm /stripchart gave no offset: $($said -replace '^\d[^,]*,\s*', '')"
    }
    return $offset
}

function Get-TimeSourceName {
    <#
    .SYNOPSIS
        The time source Windows syncs from, as "w32tm /query /source" names
        it, or 'unknown'.
    #>
    $r = Invoke-W32tm -Arguments '/query', '/source'
    if ($null -ne $r -and $r.ExitCode -eq 0) {
        $line = @(([string]$r.Text) -split "`r?`n" | Where-Object { $_.Trim() }) | Select-Object -First 1
        if ($line) { return $line.Trim() }
    }
    return 'unknown'
}

function Test-TimeSyncHealth {
    <#
    .SYNOPSIS
        Checks the host clock every 5 minutes and forces an NTP resync when it
        is more than 5 seconds out. Returns a warning, or $null. Does NOT
        trigger auto-fix.
    .DESCRIPTION
        A resync is logged REPAIRED only when the clock, measured again after
        it, is back within 5 seconds. Until 6.3.0 an exit code of 0 was
        logged REPAIRED without measuring again, and a resync that changed
        nothing was forced again every 5 minutes for as long as the drift
        lasted.

        When a resync fails, or the clock stays out, the warning names the
        time source Windows uses, and no resync is forced for 60 minutes
        while the drift lasts. The clock is still measured every 5 minutes,
        and the hold ends as soon as it is back within 5 seconds.

        A check that fails part-way, when w32tm cannot be run or answers the
        measurement with an error and no offset, logs a TIME WARNING naming
        the error, once for each error and at most once an hour while it
        lasts (Write-CheckError), and returns $null, so the poll goes on as
        before. A measurement that does not answer within 15s is logged as
        a skipped check instead, at most once an hour (Measure-ClockOffset).
        Until 6.3.0 such an error was passed over without a word, and a
        check that could not run looked like a clock that was right.

        Once the drift is known, a resync, a measurement after it or a query
        of the time source that fails is named in the warning, and the
        warning is still returned. Until 6.3.0 any of them that threw lost
        the warning, and only the failure was logged.

        A hold more than 60 minutes ahead is cleared: no hold is set longer
        than that, so the clock was set back since, as a resync can do, and
        the hold would otherwise keep resyncs off for as long again. Until
        6.3.0 it kept them off.
    #>
    $now = [datetime]::UtcNow
    # A last check in the future, after the clock was set back, counts as
    # old, as it does in Test-LineDue: a resync that sets the clock back by
    # hours would otherwise hold the next check off for as long. Until
    # 6.3.0 it did.
    $sinceCheck = ($now - $script:LastTimeSyncCheck).TotalMinutes
    if ($sinceCheck -ge 0 -and $sinceCheck -lt 5) { return $null }
    $script:LastTimeSyncCheck = $now
    if ($script:ResyncBackoffUntil -gt $now.AddMinutes(60)) { $script:ResyncBackoffUntil = [datetime]::MinValue }

    try {
        # Until 6.3.1 W32Time was read with SilentlyContinue, so a read that
        # failed looked like a service that is not installed, and nobody was
        # told. Get-ClaudeServiceOrNull tells the two apart by the error's
        # identity (NoServiceFoundForGivenName, ServiceCommandException).
        # A read that fails is said through the hourly throttle, and the
        # service is not started on it, since whether it runs is not known;
        # the clock is still measured.
        $w32svc = $null
        try {
            $w32svc = Get-ClaudeServiceOrNull -Name 'W32Time'
        } catch {
            Write-CheckError -Check 'W32Time' -ErrorRecord $_ -Now $now `
                -Text 'TIME WARNING: The Windows Time service could not be read ({0}), so it is not started if it is stopped; the clock is still measured every 5 min'
        }
        if ($w32svc -and $w32svc.Status -ne "Running") {
            try {
                Start-Service -Name "W32Time" -ErrorAction Stop
                Write-WatchLog "REPAIRED: Started W32Time service"
            } catch {
                return "W32Time service stopped, clock may drift"
            }
        }

        $offset = Measure-ClockOffset
        if ($null -eq $offset) { return $null }
        $drift = [math]::Round([math]::Abs($offset), 2)
        if ($drift -le 5.0) {
            if ($script:ResyncBackoffUntil -ne [datetime]::MinValue) {
                Write-WatchLog "Clock drift back to ${drift}s; forced resyncs allowed again"
                $script:ResyncBackoffUntil = [datetime]::MinValue
            }
            return $null
        }
        if ($now -lt $script:ResyncBackoffUntil) {
            # The warning that started the hold said why. Past 30 seconds it
            # is said again: that much drift breaks the VM's connections.
            if ($drift -gt 30.0) { return "Clock drift ${drift}s, may cause VM connectivity issues" }
            return $null
        }

        # Each step from here in a try of its own: one that fails is named in
        # the warning, and the drift is still reported.
        $resync = $null
        $resyncError = $null
        try {
            $resync = Invoke-W32tm -Arguments '/resync', '/force'
        } catch {
            # Invoke-W32tm's own error names the command already.
            $resyncError = Get-CheckErrorText -ErrorRecord $_
            if ($resyncError -notlike 'w32tm /resync *') { $resyncError = "w32tm /resync could not be run: $resyncError" }
        }
        $after = $null
        $afterError = $null
        if ($null -ne $resync -and $resync.ExitCode -eq 0) {
            $afterOffset = $null
            try {
                $afterOffset = Measure-ClockOffset -AfterResync
            } catch {
                $afterError = Get-CheckErrorText -ErrorRecord $_
            }
            if ($null -ne $afterOffset) {
                $after = [math]::Round([math]::Abs($afterOffset), 2)
                if ($after -le 5.0) {
                    Write-WatchLog "REPAIRED: Forced NTP resync (drift was ${drift}s, now ${after}s)"
                    $script:ResyncBackoffUntil = [datetime]::MinValue
                    return $null
                }
            }
        }
        if ($resyncError) {
            $what = $resyncError
        } elseif ($null -eq $resync) {
            $what = "w32tm /resync did not answer within 15s"
        } elseif ($resync.ExitCode -ne 0) {
            $what = "w32tm /resync failed (exit $($resync.ExitCode))"
        } elseif ($afterError) {
            $what = "the clock could not be measured after w32tm /resync ($afterError)"
        } elseif ($null -eq $after) {
            $what = "the clock could not be measured after w32tm /resync"
        } else {
            $what = "it is still ${after}s after w32tm /resync"
        }
        try {
            $source = Get-TimeSourceName
        } catch {
            $source = "unknown ($(Get-CheckErrorText -ErrorRecord $_))"
        }
        # From the time after the resync, not $now, which was read before
        # it. Until 6.3.1 it ran from $now: a resync that set the clock
        # forward left a hold that had already ended, or nearly, and the
        # next check forced another resync at once.
        $script:ResyncBackoffUntil = [datetime]::UtcNow.AddMinutes(60)
        return "Clock drift ${drift}s and $what; time source: $source. No resync is forced for 60 min while the drift lasts; it is still measured every 5 min"
    } catch {
        Write-CheckError -Check 'clock' -ErrorRecord $_ -Now $now `
            -Text 'TIME WARNING: The clock check failed ({0}), so whether the clock is right is not known. It runs again every 5 min'
    }
    return $null
}

function Get-CheckErrorText {
    <#
    .SYNOPSIS
        An error's message on one line, or the name of its type when it has
        no message.
    #>
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)
    $text = ([string]$ErrorRecord.Exception.Message -replace '\s*[\r\n]+\s*', ' ').Trim()
    if (-not $text) { $text = $ErrorRecord.Exception.GetType().FullName }
    return $text
}

function Test-CheckErrorDue {
    <#
    .SYNOPSIS
        Whether a line saying that a check failed may be written now: $true
        once for each error of a check, and again at most once an hour while
        it lasts, for at most 3 different errors of one check in an hour.
        Check 4b's STORAGE WARNING has its own (Test-StorageWarningDue).
    .DESCRIPTION
        The checks run on every poll or every few minutes, and what breaks a
        check once most often breaks it every time: unthrottled, the same
        line would come with each poll, and Test-LineDue's 10 minutes would
        still give 6 an hour. Keyed on the check and the error's message, so
        a different error is said at once; and no more than 3 for one check
        in an hour, so an error whose message changes each time, as one that
        names a time or a process does, cannot fill the log either.

        Entries an hour old are dropped first, so the table stays small. A
        time in the future, after the clock was set back, counts as old,
        which matters in the clock check, which sets the clock. UTC, and in
        memory only: after a restart each error is logged again once.

        Until 6.3.0 the checks passed a failure over without a word, and a
        check that could not run looked like one that found nothing wrong.
    #>
    param([string]$Check, [string]$Message, [datetime]$Now)
    foreach ($k in @($script:CheckErrorsSaid.Keys)) {
        $age = ($Now - $script:CheckErrorsSaid[$k]).TotalMinutes
        if ($age -ge 60 -or $age -lt 0) { $script:CheckErrorsSaid.Remove($k) }
    }
    $key = "$Check|$Message"
    if ($script:CheckErrorsSaid.ContainsKey($key)) { return $false }
    $ofCheck = @($script:CheckErrorsSaid.Keys | Where-Object { $_.StartsWith("$Check|", [StringComparison]::OrdinalIgnoreCase) })
    if ($ofCheck.Count -ge 3) { return $false }
    $script:CheckErrorsSaid[$key] = $Now
    return $true
}

function Write-CheckError {
    <#
    .SYNOPSIS
        Logs that a check or a read failed, naming the error, when
        Test-CheckErrorDue lets it.
    .DESCRIPTION
        -Text is the line, with {0} where the error goes (Get-CheckErrorText);
        "; the same error is logged at most once an hour" is added to it.
        -Check names the check, which with the error is what the throttle
        keys on. Returns nothing: the caller goes on as it would without the
        check, and returns what it returned before.
    #>
    param(
        [string]$Check,
        [string]$Text,
        [System.Management.Automation.ErrorRecord]$ErrorRecord,
        [datetime]$Now = [datetime]::UtcNow
    )
    $err = Get-CheckErrorText -ErrorRecord $ErrorRecord
    if (Test-CheckErrorDue -Check $Check -Message $err -Now $Now) {
        Write-WatchLog ($Text.Replace('{0}', $err) + '; the same error is logged at most once an hour')
    }
}

function Test-StorageWarningDue {
    <#
    .SYNOPSIS
        Whether Check 4b may log a STORAGE WARNING now: $true for a text not
        logged in the last hour, so a new warning, or one whose text changes
        to one not logged in that hour, is logged at once, and the same text
        again at most once an hour while it lasts.
    .DESCRIPTION
        Keyed on the text alone, with no cap on how many different texts are
        logged in an hour. The texts are few, so the log stays small: the
        low-space text moves in steps of 0.1 GB below 2 GB, so a volume that
        fills gives at most 21 different texts, from 2 GB free to 0 GB; a
        free space that goes back and forth across one step gives 2 an hour,
        and a VM cache folder that cannot be read gives one for each error.

        Not Test-CheckErrorDue: its cap of 3 different errors of one check in
        an hour held back the fourth low-space text as the volume filled, the
        one nearest to full, and a new fault after three, for up to an hour;
        and the storage check's own free-space read errors, which go through
        it under the key 'storage', were held back by the warnings too.

        Entries an hour old are dropped first; a time in the future, after
        the clock was set back, counts as old. UTC, and in memory only: after
        a restart a warning that lasts is logged again once. Until 6.3.1 the
        warning was written on every poll, about 120 times an hour.
    #>
    param([string]$Message, [datetime]$Now)
    foreach ($k in @($script:StorageWarningsSaid.Keys)) {
        $age = ($Now - $script:StorageWarningsSaid[$k]).TotalMinutes
        if ($age -ge 60 -or $age -lt 0) { $script:StorageWarningsSaid.Remove($k) }
    }
    if ($script:StorageWarningsSaid.ContainsKey($Message)) { return $false }
    $script:StorageWarningsSaid[$Message] = $Now
    return $true
}

function Test-VmcomputeHealth {
    <#
    .SYNOPSIS
        Checks the vmcompute process for handle leaks that indicate HCS instability.
        Returns $null if healthy, or a trigger string if a serious leak is detected.
        Requires 2 consecutive checks above threshold before triggering.
        Only checks every 60 seconds.
    #>
    # Skip during startup grace period
    if (Test-StartupGracePeriod) { return $null }

    # Only check every 60 seconds
    $now = [datetime]::UtcNow
    if (($now - $script:LastVmcomputeCheck).TotalSeconds -lt 60) { return $null }
    $script:LastVmcomputeCheck = $now

    try {
        # @()-wrapped. Get-Process returns a scalar for a single match, and
        # .HandleCount on a scalar is fine, but the collection form is what the
        # rest of this file assumes and it also survives zero matches.
        #
        # Through Get-ClaudeProcessByName: a vmcompute that is not running
        # (NoProcessFoundForGivenName) gives nothing, and any other error is
        # thrown to the catch below, which says it. Until 6.3.1 it was read
        # with SilentlyContinue, and a list that could not be read was taken
        # for a vmcompute that is not running.
        $vmcomputeProcs = @(Get-ClaudeProcessByName -Name 'vmcompute')
        if ($vmcomputeProcs.Count -eq 0) { return $null }

        $handleCount = 0
        $countsRead = 0
        foreach ($vp in $vmcomputeProcs) {
            # Through the getter method: a property whose getter throws reads
            # as nothing in PowerShell, with no error (measured on 5.1), so
            # until 6.3.0 a count that could not be read was taken as 0 and
            # the catch below was never reached.
            try {
                $hc = $vp.get_HandleCount()
                $countsRead++
                if ($hc -gt $handleCount) { $handleCount = $hc }
            } catch {
                Write-CheckError -Check 'vmcompute handles' -ErrorRecord $_ `
                    -Text "The handle count of vmcompute (PID $($vp.Id)) could not be read ({0}), so a handle leak in it would not be seen"
            }
        }
        # No count read is not a count of 0. Until 6.3.1 it was taken for
        # one and set the run of high counts back to 0, so a count that
        # could be read only now and then could never reach the second
        # high check that starts a repair.
        if ($countsRead -eq 0) { return $null }

        # Thresholds are 20000 and 12000, not 10000 and 5000.
        #
        # Measured on a healthy machine: vmcompute holds 215 handles. The old
        # 5000 warning level was chosen without a baseline, and the real memory
        # pressure on this box is in vmwp (645 handles), which nothing measures.
        # These numbers are still a guess, but they are now a guess anchored to
        # an observation, and they no longer sit two orders of magnitude below
        # a level that would indicate an actual leak.
        if ($handleCount -gt 20000) {
            $script:VmcomputeLeakCount++
            if ($script:VmcomputeLeakCount -ge 2) {
                $script:VmcomputeLeakCount = 0
                return "vmcompute handle leak critical ($handleCount handles), service needs restart"
            }
            Write-WatchLog "vmcompute handle count high ($handleCount), check $($script:VmcomputeLeakCount)/2"
        } elseif ($handleCount -gt 12000) {
            Write-WatchLog "vmcompute handle count elevated ($handleCount)"
            $script:VmcomputeLeakCount = 0
        } else {
            $script:VmcomputeLeakCount = 0
        }
    } catch {
        Write-CheckError -Check 'vmcompute handles' -ErrorRecord $_ -Text "The vmcompute handle check failed ({0})"
    }
    return $null
}

function Test-HcsStateHealth {
    <#
    .SYNOPSIS
        Monitors HCS state for stale cowork-vm and 0xC037010D shutdown failure
        frequency. Returns $null if healthy, or a trigger string if
        intervention is needed.

        It used to also count session transcript files and treat a high count
        as a critical issue. That check is gone: see the note where Check 3
        used to be.

        A check that fails is logged, naming the error, at most once an hour
        for the same error (Write-CheckError). Until 6.3.0 it was passed over.
    #>
    # Skip during startup grace period
    if (Test-StartupGracePeriod) { return $null }

    $issues = @()

    # Check 1: Stale cowork-vm in hcsdiag (should not exist while service is healthy)
    try {
        $hcsList = Get-HcsListCached
        if ($hcsList) {
            # Counts DISTINCT GUIDs, not occurrences of the name.
            #
            # "hcsdiag list" prints the compute system name twice for a single
            # VM: once on its own line and again at the end of the detail line.
            # Counting matches of "cowork-vm" therefore returned 2 for one
            # healthy VM, which is above the threshold of 1 used when no Cowork
            # session is active. That made "Multiple cowork-vm instances in HCS"
            # a standing false positive on a completely normal machine.
            $vmCount = @(Get-CoworkHcsGuids -ListOutput ([string]$hcsList)).Count
            # During active Cowork sessions, 1-2 HCS instances is normal
            # (VM runtime + network/storage component). Only flag 3+.
            # When no Cowork session is active, flag 2+ (stale leftovers).
            $coworkActive = Test-CoworkSessionActive
            $threshold = if ($coworkActive) { 2 } else { 1 }
            if ($vmCount -gt $threshold) {
                $issues += "Multiple cowork-vm instances in HCS ($vmCount found, threshold $threshold)"
            }
        }
    } catch {
        Write-CheckError -Check 'HCS cowork-vm count' -ErrorRecord $_ `
            -Text "HCS: The check for more than one cowork-vm compute system failed ({0})"
    }

    # Check 2: 0xC037010D shutdown failure frequency
    # These are caused by a Claude Desktop property query bug (literal '$') and happen
    # on every VM shutdown. Rate-limit logging and skip during active sessions.
    try {
        if (-not (Test-CoworkSessionActive)) {
            # -MaxEvents added as a ceiling. The query itself is left alone.
            #
            # This was flagged as an unbounded scan of a channel holding 19,491
            # records, rendering every message every 30 seconds. Measured on
            # this machine, it is not: StartTime is pushed down into the event
            # log service, so the query returns 68 events in 68 ms and
            # rendering all 68 messages costs a further 8 ms. The 19,491 is the
            # size of the whole channel, not of this query.
            #
            # A -FilterXPath rewrite was tried and reverted. It is invalid
            # syntax for this cmdlet unless the comparison is written raw, it
            # cannot address the error code (which lives in the rendered
            # message, not a named EventData field), and it bought nothing over
            # a filter that was already fast.
            #
            # What -MaxEvents does buy is a bound for a machine having a very
            # bad hour, where the same query could return thousands.
            $shutdownFilter = @{
                LogName   = "Microsoft-Windows-Hyper-V-Compute-Operational"
                # The UTC time as round-trip text: see Test-EventLogErrors.
                StartTime = [datetime]::UtcNow.AddHours(-1).ToString('o')
            }
            $allShutdownEvents = @(Get-MatchingWinEvent -Filter $shutdownFilter -MaxEvents 500 |
                Where-Object { $_.Message -match "0xC037010D" })
            $shutdownCount = $allShutdownEvents.Count
            if ($shutdownCount -gt 0) {
                $shutdownLogElapsed = ([datetime]::UtcNow - $script:LastShutdownFailLogTime).TotalMinutes
                if ($shutdownLogElapsed -ge 10) {
                    Write-WatchLog "HCS: $shutdownCount shutdown failures (0xC037010D) in last hour (property query bug, informational)"
                    $script:LastShutdownFailLogTime = [datetime]::UtcNow
                }
            } else {
                # Reset rate-limit timer when count drops to zero
                $script:LastShutdownFailLogTime = [datetime]::MinValue
            }
            if ($shutdownCount -gt 10) {
                $issues += "HCS shutdown failure spike: $shutdownCount events in last hour"
            }
            # Spike detection: excessive shutdown failures may indicate real corruption
            # In UTC, for the same reason as Test-ClaudeCodeActive's ages.
            $fiveMinAgoUtc = [datetime]::UtcNow.AddMinutes(-5)
            $recentSpike = @($allShutdownEvents | Where-Object { $_.TimeCreated -and $_.TimeCreated.ToUniversalTime() -gt $fiveMinAgoUtc })
            if ($recentSpike.Count -gt 15) {
                # Before restarting vmcompute, check if there is an active session
                #
                # Every process named claude counts, the Claude Code CLI
                # included: this only holds a vmcompute restart back, and
                # erring that way is safe.
                #
                # Until 6.3.1 both were read with SilentlyContinue, and a read
                # that failed read as no service and no app, which let the
                # restart through under a session that may be live. The
                # helpers give nothing only for the error that means not
                # there (NoServiceFoundForGivenName, NoProcessFoundForGivenName)
                # and throw any other; a read that fails holds the restart
                # back, said through the hourly throttle.
                $sessionReadError = $null
                $coworkSvc = $null
                $claudeRunning = $false
                try {
                    $coworkSvc = Get-ClaudeServiceOrNull -Name $ServiceName
                    $claudeRunning = @(Get-ClaudeProcessByName -Name $ProcessName).Count -gt 0
                } catch {
                    $sessionReadError = $_
                }
                # Fix's lock, as Invoke-AutoFix asks it. This was pinned to
                # "Global\ClaudeDesktopFix_v4.8", so it never found the lock and
                # would have restarted vmcompute out from under a running
                # repair; and until 6.3.0 it kept a copy of the check of its
                # own, which read a lock it may not open as not held.
                if ($sessionReadError) {
                    Write-CheckError -Check 'HCS spike' -ErrorRecord $sessionReadError `
                        -Text "HCS SPIKE: $($recentSpike.Count) failures in 5 min, but whether a session is active could not be read ({0}); skipping vmcompute restart"
                } elseif ($claudeRunning -and $coworkSvc -and $coworkSvc.Status -eq "Running") {
                    Write-WatchLog "HCS SPIKE: $($recentSpike.Count) failures in 5 min but active session detected, skipping vmcompute restart"
                } elseif (Test-FixLockHeld) {
                    Write-WatchLog "HCS SPIKE: $($recentSpike.Count) failures in 5 min but Fix is running, skipping vmcompute restart"
                } else {
                    Write-WatchLog "HCS SPIKE: $($recentSpike.Count) shutdown failures in 5 min, restarting vmcompute"
                    try {
                        Restart-Service -Name "vmcompute" -Force -ErrorAction Stop
                        Start-Sleep -Seconds 3
                        Write-WatchLog "vmcompute restarted (pre-emptive, cleared stale HCS state)"
                    } catch {
                        Write-WatchLog "vmcompute restart failed: $($_.Exception.Message)"
                    }
                }
            }
        }
    } catch {
        Write-CheckError -Check 'HCS shutdown failures' -ErrorRecord $_ `
            -Text "HCS: The check for shutdown failures (0xC037010D) failed ({0})"
    }

    # Check 3 has been removed entirely.
    #
    # It counted files under local-agent-mode-sessions and added
    # "Session file accumulation critical" to $issues past 1000. On this
    # machine that count is 10,433, so it fired on EVERY poll: today's watch log
    # is 97 lines and every one of them is that message. Because $issues was
    # non-empty it also held the HCS trigger below permanently open, and it cost
    # 1,841 ms of recursive directory enumeration every 30 seconds.
    #
    # A session file count has no relationship to HCS or VM health. Transcripts
    # accumulate because people use the product. Fix-ClaudeDesktop now offers
    # -PurgeSessions for anyone who wants to trim them, and that is a disk-space
    # decision, not a health signal.

    if ($issues.Count -gt 0) {
        return ($issues -join "; ")
    }
    return $null
}

# -- VM maintenance functions ---------------------------------------------

function Set-VmWorkerPriority {
    <#
    .SYNOPSIS
        Raises the Cowork VM's worker process to AboveNormal.
    .DESCRIPTION
        Only the vmwp.exe that runs the Cowork VM's compute system, picked by
        Get-CoworkWorkerProcessIds. Until 6.3.0 this raised every vmwp.exe on
        the machine: every Hyper-V VM, WSL2 and Windows Sandbox. A priority
        does not outlive its process, so this runs on every poll and catches
        each new VM.

        The worker processes are listed with Stop. A list that cannot be
        read, or a worker that cannot be raised, is logged, naming why, at
        most once an hour for the same error (Write-CheckError); a worker
        that exited since the list was read has nothing to raise, and is
        passed over. Until 6.3.0 the list was read with SilentlyContinue,
        and a list that could not be read, or a raise that failed, left the
        worker where it was without a word.
    #>
    # SupportsShouldProcess: this genuinely changes system state, raising the
    # scheduling priority of a live process.
    [CmdletBinding(SupportsShouldProcess)]
    param()
    try {
        $hcsList = Get-HcsListCached
        if (-not $hcsList) { return }
        # The process query only when a Cowork VM is listed at all.
        if (@(Get-CoworkHcsGuids -ListOutput ([string]$hcsList)).Count -eq 0) { return }
        try {
            $workers = @(Get-CimInstance Win32_Process -Filter "Name = 'vmwp.exe'" -ErrorAction Stop)
        } catch {
            Write-CheckError -Check 'VM worker priority' -ErrorRecord $_ `
                -Text "The VM worker processes (vmwp.exe) could not be listed ({0}), so the Cowork VM's worker was not raised to AboveNormal"
            return
        }
        foreach ($id in @(Get-CoworkWorkerProcessIds -ListOutput ([string]$hcsList) -Workers $workers)) {
            try {
                $p = Get-Process -Id $id -ErrorAction Stop
                if ($p.PriorityClass -ne 'AboveNormal') {
                    if ($PSCmdlet.ShouldProcess("vmwp.exe PID $id", "Set priority to AboveNormal")) {
                        $p.PriorityClass = 'AboveNormal'
                        Write-WatchLog "Raised the Cowork VM's worker (vmwp.exe PID $id) to AboveNormal"
                    }
                }
            } catch {
                if ([string]$_.FullyQualifiedErrorId -notlike 'NoProcessFoundForGivenId*') {
                    Write-CheckError -Check 'VM worker priority' -ErrorRecord $_ `
                        -Text "The Cowork VM's worker (vmwp.exe PID $id) could not be raised to AboveNormal ({0})"
                }
            }
        }
    } catch {
        Write-CheckError -Check 'VM worker priority' -ErrorRecord $_ `
            -Text "The Cowork VM's worker could not be raised to AboveNormal ({0})"
    }
}

function Stop-StaleCoworkVm {
    <#
    .SYNOPSIS
        Kills the Cowork compute systems left behind beside the live one, and
        logs what became of each.
    .DESCRIPTION
        Run from the main loop with the HCS STATE line, at most once every 10
        minutes. The live VM is the one whose worker process (vmwp.exe)
        started last; every other Cowork compute system "hcsdiag list" shows
        is killed, and the list is read again to see which are gone. Until
        6.3.0 this kept the last entry hcsdiag printed, on the belief that
        the list runs oldest to newest, which nothing documents, and it
        logged "killed" for every kill it sent.

        The worker processes are listed with Stop. A list that cannot be read
        is logged, naming why, at most once an hour for the same error
        (Write-CheckError), and nothing is killed. Until 6.3.0 it was read
        with SilentlyContinue, and logged as if no compute system had a
        worker process: "none has a worker process to tell the live one by".
        That line is now only for a list that was read.

        Rewritten before 6.3.0. The block this replaced carried the same two
        bugs as Fix-ClaudeDesktop findings 1 and 2: it looked for the GUID on
        a line of its own, which "hcsdiag list" never prints, so it never ran;
        and it called "hcsdiag close", a verb hcsdiag does not have. The parse
        lives in Get-CoworkHcsGuids, in the shared ClaudeEnv region. Until
        6.3.0 this was part of the main loop.
    #>
    # SupportsShouldProcess: this kills compute systems.
    [CmdletBinding(SupportsShouldProcess)]
    param()
    try {
        $hcsList = Get-HcsListCached
        if (-not $hcsList) { return }
        $vmGuids = @(Get-CoworkHcsGuids -ListOutput ([string]$hcsList))
        if ($vmGuids.Count -le 1) { return }
        try {
            $workers = @(Get-CimInstance Win32_Process -Filter "Name = 'vmwp.exe'" -ErrorAction Stop)
        } catch {
            Write-CheckError -Check 'stale VMs' -ErrorRecord $_ `
                -Text "HCS: $($vmGuids.Count) cowork-vm compute systems, and the VM worker processes (vmwp.exe) could not be listed to tell the live one by ({0}); none killed"
            return
        }
        $started = @{}
        $unread = 0
        foreach ($w in $workers) {
            # A worker whose command line could not be read cannot be told
            # to be a Cowork one or not, and a Cowork worker with no start
            # time cannot be told to be the live one or not. Until 6.3.1
            # either was passed over: if it was the live VM's worker, the
            # live VM was taken for a stale one and killed. Now nothing is
            # killed while any worker is unread.
            $cmd = [string]$w.CommandLine
            if (-not $cmd) { $unread++; continue }
            $wg = Get-VmWorkerGuid -CommandLine $cmd
            if (-not $wg) { continue }
            if (-not $w.CreationDate) {
                if ($vmGuids -contains $wg) { $unread++ }
                continue
            }
            # In UTC: in the hour the clocks go back, a worker started later
            # can show an earlier local time. Measured 2026-11-22 on 5.1:
            # Win32_Process.CreationDate comes from CIM as a local DateTime
            # holding the same instant as Process.StartTime, and one in the
            # repeated hour keeps which of the two it is, so this UTC time is
            # exact (2026-10-25 00:30Z read as 01:30 and went back to 00:30Z).
            $started[$wg] = ([datetime]$w.CreationDate).ToUniversalTime()
        }
        if ($unread -gt 0) {
            if (Test-CheckErrorDue -Check 'stale VMs' -Message 'a worker could not be read' -Now ([datetime]::UtcNow)) {
                Write-WatchLog ("HCS: $($vmGuids.Count) cowork-vm compute systems, and the command line or start time of $unread VM worker " +
                    "process(es) (vmwp.exe) could not be read to tell the live one by; none killed; the same error is logged at most once an hour")
            }
            return
        }
        $live = $null
        foreach ($guid in $vmGuids) {
            if ($started.ContainsKey($guid) -and ($null -eq $live -or $started[$guid] -gt $started[$live])) {
                $live = $guid
            }
        }
        if ($null -eq $live) {
            Write-WatchLog "HCS: $($vmGuids.Count) cowork-vm compute systems and none has a worker process to tell the live one by; none killed"
            return
        }
        $sent = @()
        foreach ($guid in $vmGuids) {
            if ($guid -eq $live) { continue }
            if ($PSCmdlet.ShouldProcess("cowork-vm $guid", "hcsdiag kill")) {
                $null = Invoke-HcsDiag -Arguments "kill", $guid
                $sent += $guid
            }
        }
        if ($sent.Count -eq 0) { return }
        $left = Get-CoworkHcsSurvivors -Guids $sent -List { Invoke-HcsDiag -Arguments "list" }
        # The cached list predates the kill.
        $script:LastHcsListTime = [datetime]::MinValue
        foreach ($guid in $sent) {
            if ($null -eq $left) {
                Write-WatchLog "HCS: sent a kill to stale cowork-vm ($guid); hcsdiag did not answer the check after it"
            } elseif ($left -contains $guid) {
                # Listed when last read, which is not proof the kill failed
                # for good: it can take longer to leave the list than the
                # check waits.
                Write-WatchLog "HCS: stale cowork-vm ($guid) was still listed after the kill"
            } else {
                Write-WatchLog "HCS: killed stale cowork-vm ($guid)"
            }
        }
    } catch {
        Write-CheckError -Check 'stale VMs' -ErrorRecord $_ -Text 'HCS: the check for stale cowork-vm compute systems failed ({0})'
    }
}

# -- User activity detection ---------------------------------------------

function Test-ClaudeCodeActive {
    <#
    .SYNOPSIS
        Returns $true if a Claude Code session appears to be active.
        Claude Code runs INSIDE Claude Desktop but does NOT use the Cowork VM.
        It communicates directly with the Anthropic API, so VM staleness is
        expected and normal during Code usage.

        CHECKS:
        0. A Claude Code CLI the app started is running: a project thread or
           Code session (Get-ClaudeAppCliProcess). Added in 6.3.0.
        1. Session files: claude-code-sessions folder for recently-modified files (10 min)
        2. App logs: unknown-window*.log, claude.ai-web*.log or remote-control*.log
           recently written (10 min). The first two show LOCAL_SESSION activity
           when Code is streaming/processing; the third is written for Remote
           Control sessions, which is how project threads run on a PC.
        3. Claude Code's own transcripts: any *.jsonl under the CLI's
           projects folder written in the last 35 min.

        WHY CHECK 3. Checks 1 and 2 are things the app writes about a session,
        and it does not write them while the session works. Measured
        2026-11-21 17:25, with two project threads working that minute: their
        transcripts were written at 17:25, the newest session file at 17:09,
        remote-control.log at 17:12 and claude.ai-web.log at 16:47. Checks 1
        and 2 both answered no. The CLI appends to its transcript on every
        message and tool result, and a subagent to its own file beside it.
        A terminal session counts too: holding a repair back for one costs a
        late repair, or one the user is told to run by hand, which is the
        cheap mistake.

        WHY 35 MINUTES. A tool call writes nothing to the transcript while it
        runs, and a foreground command may run for 30 minutes. Measured
        2026-11-21 in one project thread's transcript: three Bash calls of 25
        minutes each in a row, with no line written during the first two and
        one queue line during the third. With a 10-minute window, and no
        other session writing, the guard answered "no session" from minute 10
        of each, and a repair could close the app under the running command.
        The cost is that a repair waits up to 35 minutes after the last Code
        activity.

        WHERE. Both checks used to read %APPDATA%\Claude only, through
        $ClaudeAppData and $ClaudeLogDir. On app 6.18355 (measured 2026-11-21)
        neither folder is written any more: the sessions live under
        %LOCALAPPDATA%\Claude-Data\claude_<id>\claude-code-sessions\ and the
        logs in %LOCALAPPDATA%\Claude\logs, so this returned false with a
        Code session running and stopped holding repairs back. It now asks
        Get-ClaudeDataCandidates on every call and reads every candidate, so a
        later move is covered without a restart. Reading a stale folder as
        well costs a directory listing; missing the live one cost the guard.
    .PARAMETER Silent
        Writes no line to the watch log: for Invoke-HeldBackRetry's quiet
        check, which asks this every 2 minutes while a repair waits. The
        one-time line about the missing folder is then left for a call
        without it to write.
    #>
    param([switch]$Silent)
    try {
        # Check 0: a Claude Code CLI the app started. A session waiting on a
        # long command uses no CPU and writes nothing, so checks 1 to 3 can
        # all miss it once their windows pass, and the repair ends it. This
        # holds every repair back for as long as the app runs a session. It
        # throws when it cannot tell, and the catch below then answers yes.
        if (@(Get-ClaudeAppCliProcess -Environment $script:ClaudeEnv).Count -gt 0) { return $true }

        $cands = Get-ClaudeDataCandidates -PackageFamilyName $script:ClaudeEnv.PackageFamilyName
        # Ages in UTC. Local time jumps an hour when summer time starts or
        # ends, and for that hour a file written 70 minutes ago read as 10
        # minutes old, or the other way round. Until 6.3.0 these were local.
        $nowUtc = [datetime]::UtcNow

        # What could not be read is gathered. A recent sign anywhere answers
        # yes whatever else could not be read; without one, a read that
        # failed answers yes too, through the catch below. Until 6.3.1 a
        # folder that could not be read was passed over as one with nothing
        # recent in it, and the answer was "no session".
        $unread = @($cands.Errors)

        # Check 1: Claude Code session persistence files, written as Code works,
        # under <data root>\claude-code-sessions\{accountId}\{orgId}\.
        $sessionDirsSeen = 0
        foreach ($root in @($cands.DataRoots)) {
            $codeSessionDir = Join-Path $root.Path "claude-code-sessions"
            try {
                $sessionDir = Get-ClaudeItemOrNull -LiteralPath $codeSessionDir
            } catch {
                $unread += "$codeSessionDir could not be read: $(Get-ClaudeErrorText $_)"
                continue
            }
            if (-not $sessionDir -or -not $sessionDir.PSIsContainer) { continue }
            $sessionDirsSeen++
            $walkErrors = $null
            $recentFiles = Get-ChildItem -LiteralPath $codeSessionDir -Filter "*.json" -Recurse -File `
                    -ErrorAction SilentlyContinue -ErrorVariable walkErrors |
                Where-Object { ($nowUtc - $_.LastWriteTimeUtc).TotalMinutes -lt 10 } |
                Select-Object -First 1
            if ($recentFiles) { return $true }
            # A folder that went away during the walk (PathNotFound) is a
            # session that ended. @($null) holds one element, so a $null is
            # passed over.
            foreach ($we in @($walkErrors)) {
                if ($null -eq $we) { continue }
                if (-not (Test-ClaudeNotFoundError -ErrorRecord $we -ErrorId 'PathNotFound' `
                        -ExceptionType 'System.Management.Automation.ItemNotFoundException')) {
                    $unread += "$codeSessionDir could not be read in full: $(Get-ClaudeErrorText $we)"
                }
            }
        }
        if ($sessionDirsSeen -eq 0 -and $unread.Count -eq 0 -and -not $Silent -and -not $script:NoCodeSessionDirLogged) {
            # Said once: either Code has never run here, or the app has moved
            # its data somewhere Get-ClaudeDataCandidates does not look.
            $script:NoCodeSessionDirLogged = $true
            Write-WatchLog "Test-ClaudeCodeActive: no claude-code-sessions folder under any known data root; only the log and transcript checks can see a Code session"
        }

        # Check 2: app log recency. These are written as the app shows a
        # session, which can be minutes apart while the session works; check 3
        # covers the gaps. A folder whose newest log is older than the window
        # cannot hold a recent one, so it is skipped unread.
        foreach ($ld in @($cands.LogDirs)) {
            if (-not $ld.Newest -or ($nowUtc - $ld.Newest.ToUniversalTime()).TotalMinutes -ge 10) { continue }
            foreach ($pattern in @('unknown-window*.log', 'claude.ai-web*.log', 'remote-control*.log')) {
                # A folder that went away since it was listed holds nothing.
                try {
                    $rLogs = @(Get-ChildItem -LiteralPath $ld.Path -Filter $pattern -File -ErrorAction Stop)
                } catch {
                    if (-not (Test-ClaudeNotFoundError -ErrorRecord $_ -ErrorId 'PathNotFound' `
                            -ExceptionType 'System.Management.Automation.ItemNotFoundException')) {
                        $unread += "$($ld.Path) could not be read: $(Get-ClaudeErrorText $_)"
                    }
                    # The folder's other logs are not listed either: it is
                    # said once, and a folder gone holds none of them.
                    break
                }
                foreach ($rLog in $rLogs) {
                    if (($nowUtc - $rLog.LastWriteTimeUtc).TotalMinutes -lt 10) { return $true }
                }
            }
        }

        # Check 3: the CLI's transcripts, through Test-ClaudeCodeTranscriptRecent
        # in the shared region, so Fix -Quiet asks the same question. The
        # watchdog task runs as the user, so the config dir it finds is the
        # user's. This runs only when a repair is about to start, and, while
        # a held-back repair waits, at most every 2 minutes; not on every
        # poll. The window is longer than the CLI's 30-minute limit on a
        # foreground command, because a tool call writes nothing there until
        # it returns.
        if (Test-ClaudeCodeTranscriptRecent -Minutes 35) { return $true }

        if ($unread.Count -gt 0) {
            throw "no recent sign of a session, and what could not be read may hold one: $($unread -join '; ')"
        }
    } catch {
        # Fails CLOSED. Test-UserActivity calls this to decide whether a repair
        # may proceed, so an unknown answer must not read as "no session".
        # At most once an hour for the same error (Write-CheckError); until
        # 6.3.1 on each call.
        if (-not $Silent) {
            Write-CheckError -Check 'Code session' -ErrorRecord $_ `
                -Text "Test-ClaudeCodeActive failed ({0}); assuming a Code session IS active"
        }
        return $true
    }
    return $false
}

function Test-UserActivity {
    <#
    .SYNOPSIS
        Returns $true if the user appears to be actively using Claude Desktop.
        Design: false positives (blocking a fix) are cheap; false negatives
        (killing active work) are expensive. So we err on the side of caution.

        CHECKS (interactive session):
        1. Foreground window belongs to a Claude process (via GetWindowThreadProcessId)
        2. User input within 3 minutes + Claude running
        3. VM log active within 120s (covers Code thinking phases)
        4. CPU sampling -- any Claude process burning >100ms CPU in 500ms
        5. Active Claude Code session (session files, app logs or CLI transcripts)
        6. Active Cowork session (created but not yet cleaned up in VM log)

        SESSION 0 (SYSTEM scheduled tasks):
        - Win32 APIs return zero/stale data, so we skip checks 1-2
        - Falls back to VM log + CPU sampling + Code session + Cowork session detection

        -Silent writes no line to the watch log, here or in
        Test-ClaudeCodeActive. Invoke-HeldBackRetry asks this before a retry,
        so a retry shows the warning only when BLOCKED would not stop it, and
        its wait writes one line every 10 minutes at most.
    #>
    param([switch]$Silent)
    try {
        # Every process named claude, the Claude Code CLI included: a busy
        # project thread or terminal session is activity, and this only ever
        # holds a repair back. A list that could not be read throws to the
        # catch below, which answers yes. Until 6.3.1 it was empty, and the
        # answer was "not in use".
        $claudeProcs = @(Get-ClaudeProcessByName -Name $ProcessName)
        if ($claudeProcs.Count -eq 0) { return $false }

        if ($script:IsInteractiveSession -and $script:Win32ActivityLoaded) {
            # Check 1: Foreground window PID matches a Claude process (or its parent)
            # Uses GetWindowThreadProcessId -- works with Electron renderer processes
            $fgHwnd = [Win32Activity]::GetForegroundWindow()
            if ($fgHwnd -ne [IntPtr]::Zero) {
                $fgPid = 0
                [Win32Activity]::GetWindowThreadProcessId($fgHwnd, [ref]$fgPid) | Out-Null
                if ($fgPid -gt 0) {
                    foreach ($cp in $claudeProcs) {
                        if ($cp.Id -eq $fgPid) { return $true }
                    }
                    # Electron renderer -> main process: check parent PID.
                    # A window whose process exited since it was read has no
                    # entry, and is not Claude's: checks 2 to 6 follow. A read
                    # that fails throws to the catch below, which answers yes.
                    # Until 6.3.1 it was passed over, as a window not Claude's.
                    $fgProc = @(Get-CimInstance Win32_Process -Filter "ProcessId=$fgPid" -ErrorAction Stop)
                    if ($fgProc.Count -gt 0) {
                        $parentId = $fgProc[0].ParentProcessId
                        foreach ($cp in $claudeProcs) {
                            if ($cp.Id -eq $parentId) { return $true }
                        }
                    }
                }
            }

            # Check 2: User input within 3 minutes (extended from 2)
            $lastInput = New-Object Win32Activity+LASTINPUTINFO
            $lastInput.cbSize = [uint32][System.Runtime.InteropServices.Marshal]::SizeOf($lastInput)
            if ([Win32Activity]::GetLastInputInfo([ref]$lastInput)) {
                # TickCount is a SIGNED 32-bit value that goes negative once
                # uptime passes 24.9 days, while dwTime is unsigned. Subtracting
                # them directly then yields a hugely negative idle time, which
                # reads as "user is active" and pins this monitor into never
                # auto-fixing anything. Doing the arithmetic in unsigned 32-bit
                # space makes the wrap cancel out. TickCount64 would sidestep it
                # but does not exist on Windows PowerShell 5.1.
                # 0xFFFFFFFFL, with the L. A bare 0xFFFFFFFF is parsed by
                # PowerShell as Int32 -1, so the mask is a no-op on a negative
                # TickCount and the [uint32] cast then throws, in precisely the
                # past-24.9-days case this exists for.
                $nowTicks  = [uint32]([Environment]::TickCount -band 0xFFFFFFFFL)
                $lastTicks = [uint32]$lastInput.dwTime
                if ($nowTicks -ge $lastTicks) {
                    $idleMs = [long]$nowTicks - [long]$lastTicks
                } else {
                    $idleMs = [long]4294967296 - [long]$lastTicks + [long]$nowTicks
                }
                if ($idleMs -lt 180000) {
                    return $true  # User active within 3 min + Claude is running
                }
            }
        }

        # Check 3: VM log active within 120s (was 30s -- covers Code thinking phases)
        # Any VM log, read now (Get-ClaudeVmLogAge). This used to read the one
        # log chosen when the monitor started, which once meant a file last
        # written in March and on app 6.18355 can mean coworkd.log while the
        # host side writes to cowork_vm_node.log.
        $ageSec = Get-ClaudeVmLogAge -LogDir $script:ClaudeLogDir
        if ($null -ne $ageSec -and $ageSec -lt 120) {
            return $true  # VM was active recently -- Code may be thinking
        }

        # Check 4: CPU sampling -- any Claude process using >100ms CPU in 500ms.
        #
        # ONE 500ms sleep covering all processes, not one sleep per process.
        # There are sixteen Electron processes on this machine, so the old
        # per-process loop could sleep for eight seconds inside a thirty second
        # poll, and the cancel path ran the same shape again for another 48.
        $cpuBefore = @{}
        foreach ($cp in $claudeProcs) {
            # A process that exited since the list was read has no CPU time
            # to read, and is not busy. Any other failure throws to the catch
            # below, which answers yes: until 6.3.1 it read as not busy.
            try { $cpuBefore[$cp.Id] = $cp.TotalProcessorTime.TotalMilliseconds }
            catch { if (-not [bool]$cp.HasExited) { throw } }
        }
        if ($cpuBefore.Count -gt 0) {
            Start-Sleep -Milliseconds 500
            foreach ($cp in $claudeProcs) {
                try {
                    if (-not $cpuBefore.ContainsKey($cp.Id)) { continue }
                    $cp.Refresh()
                    if (($cp.TotalProcessorTime.TotalMilliseconds - $cpuBefore[$cp.Id]) -gt 100) {
                        return $true
                    }
                } catch {
                    # Exited during the sample: not busy. Anything else, as
                    # above.
                    if (-not [bool]$cp.HasExited) { throw }
                }
            }
        }

        # Check 5: Active Claude Code session
        # Code runs inside Claude Desktop but doesn't use the VM.
        # During API-wait periods, CPU can be near zero for minutes.
        # Its transcripts, session files and app logs are more reliable
        # indicators.
        if (Test-ClaudeCodeActive -Silent:$Silent) { return $true }

        # Check 6: Active Cowork session (API-wait safe)
        if (Test-CoworkSessionActive) { return $true }
    } catch {
        # Fails CLOSED, not open.
        #
        # This is the guard that stands between an automated repair and a user
        # who is mid-sentence. Returning $false on an unexpected error means
        # "nobody is using this, go ahead and kill it", which is the worst
        # possible answer to give when the truth is unknown.
        #
        # The error is logged rather than swallowed, because the previous empty
        # catch is why nobody knew this could happen. At most once an hour for
        # the same error (Write-CheckError); until 6.3.1 on each call.
        if (-not $Silent) {
            Write-CheckError -Check 'user activity' -ErrorRecord $_ `
                -Text "Test-UserActivity failed ({0}); assuming the user IS active"
        }
        return $true
    }
    return $false
}

function Test-CoworkSessionActive {
    <#
    .SYNOPSIS
        Returns $true when the Cowork VM has logged activity recently.

        Reads cowork_vm_node.log in the current log folder and looks for a
        recent [Keepalive], [startVM], [VM:start], [postConnect] or
        [vmOneShot] entry within the last 180 seconds.

        During API-wait periods, when the model is thinking server side, a
        Cowork session burns no local CPU but is very much alive. This check is
        what stops Watch killing Claude in those gaps.

        It previously paired [Process:UUID] Created against Cleaned up. Neither
        marker appears in the current log format at all.
    #>
    # cowork_vm_node.log by name, because that is the only log whose lines
    # this parses. This used to read whichever VM log was newest when the
    # monitor started. On app 6.18355 that is usually coworkd.log, which has
    # none of the tags below and writes its dates with slashes, so the answer
    # was always "no session" and the Cowork guard was off. The node log has a
    # [vmOneShot] line for every command coworkd runs, so nothing is lost.
    $vmLogFile = $null
    if ($script:ClaudeLogDir) { $vmLogFile = Join-Path $script:ClaudeLogDir 'cowork_vm_node.log' }
    if (-not $vmLogFile) { return $false }
    # Only read the last 200 lines to stay fast.
    #
    # A log that is not there gives PathNotFound with ObjectNotFound and
    # ItemNotFoundException (measured 2026-11-22): no session. Any other
    # error is said through the hourly throttle, and the answer is $true.
    # Every caller holds something back on an active session (a staleness
    # warning, an HCS trigger, a repair), so not knowing holds it back too.
    # Until 6.3.1 the log was read with SilentlyContinue after a Test-Path,
    # and a log that could not be read read as no session.
    try {
        $lines = @(Get-Content -LiteralPath $vmLogFile -Tail 200 -ErrorAction Stop)
    } catch {
        if (Test-ClaudeNotFoundError -ErrorRecord $_ -ErrorId 'PathNotFound' -ExceptionType 'System.Management.Automation.ItemNotFoundException') {
            return $false
        }
        Write-CheckError -Check 'Cowork session' -ErrorRecord $_ `
            -Text "The Cowork VM log $vmLogFile could not be read ({0}), so a Cowork session is taken to be active and nothing that would end one is done"
        return $true
    }
    try {
        if ($lines.Count -eq 0) { return $false }

        # Liveness comes from recent VM activity tags, not from pairing
        # [Process:UUID] Created against Cleaned up.
        #
        # Those two markers occur ZERO times in the live VM log. They belong to
        # the March-era coworkd.log format (6.18355 writes coworkd.log again, in
        # a newer format with neither). Reading that dead file, the last 200
        # lines held old Created entries whose matching Cleaned up lines had
        # scrolled away, so the function reported an active session forever and
        # every auto-fix took the BLOCKED branch.
        #
        # The tags below are what the current log actually emits. Measured
        # frequencies in the live file: [Keepalive] 631, [postConnect] 604,
        # [startVM] 580, [VM:start] 235, [vmOneShot] 171.
        #
        # Format is "2026-07-28 16:27:55 [info] [Keepalive] message": dashes,
        # second precision, no fractional part.
        $activityTag = '\[(Keepalive|startVM|VM:start|postConnect|vmOneShot)\]'
        $stampPat    = '^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})'
        # Local time, on purpose: the stamps are parsed from log text the app
        # writes in local time with no offset, so local now is the only time
        # they can be compared with. For the hour around a change of the
        # clock a line can read up to an hour newer or older than it is.
        $now         = Get-Date
        $windowSec   = 180

        for ($i = $lines.Count - 1; $i -ge 0; $i--) {
            $line = $lines[$i]
            if ($line -notmatch $activityTag) { continue }
            if ($line -notmatch $stampPat) { continue }
            try {
                $ts = [datetime]::ParseExact($Matches[1], 'yyyy-MM-dd HH:mm:ss',
                    [System.Globalization.CultureInfo]::InvariantCulture)
            } catch { continue }
            # Scanning newest first, so the first parseable hit decides.
            return (($now - $ts).TotalSeconds -le $windowSec)
        }
    } catch {
        # Lines that could not be read through. Said, and taken as an active
        # session, as a log that could not be read is. Until 6.3.1 this
        # answered no session without a word.
        Write-CheckError -Check 'Cowork session' -ErrorRecord $_ `
            -Text "The Cowork VM log $vmLogFile could not be read through ({0}), so a Cowork session is taken to be active and nothing that would end one is done"
        return $true
    }
    return $false
}

# -- Auto-fix function ---------------------------------------------------

function Clear-HeldBackRepair {
    <#
    .SYNOPSIS
        Ends a pending held-back repair, in memory. The caller saves the state.
    #>
    $script:PendingFixReason     = $null
    $script:PendingFixSince      = [datetime]::MinValue
    $script:PendingFixHeldAt     = [datetime]::MinValue
    $script:PendingFixTries      = 0
    $script:PendingFixRetriesHeld = 0
    $script:PendingFixNextCheck  = [datetime]::MinValue
    $script:PendingFixWaitLogged = [datetime]::MinValue
    $script:PendingFixLockLogged  = $false
}

function Close-HeldBackRepair {
    <#
    .SYNOPSIS
        Ends a pending held-back repair with one line that says why, and saves
        the state.
    .DESCRIPTION
        The line reads ">>> held-back repair for <reason> <Why>". The state is
        cleared at once, so nothing that asks again can repeat the line.
    #>
    param([string]$Why)
    Write-WatchLog ">>> held-back repair for $($script:PendingFixReason) $Why"
    Clear-HeldBackRepair
    Save-WatchState
}

function Register-HeldBackRepair {
    <#
    .SYNOPSIS
        Keeps a repair pending after a hold, in memory. The caller saves the
        state.
    .DESCRIPTION
        A hold is Fix's exit 3, BLOCKED or DEFERRED in Invoke-AutoFix, or a
        one-shot trigger in the wait one of them left. A repair not pending
        yet starts its 4 hours at -HeldAt. One already pending keeps when its
        4 hours started: a hold alone, a retry held back again or a re-checked
        trigger held back while it waits, does not stretch them, so a repair
        that keeps being held back cannot go on for ever. Only a one-shot
        trigger that fires on its own starts them again
        (Merge-HeldBackTrigger), because it is new evidence.

        -HeldAt is also the latest hold. The next retry waits from it, and a
        Fix log newer than it is a Fix run since. -FixHeld counts a run Fix
        held back after the 30-second warning, toward the 3 tries. With
        -Retry too, that run was a retry, and the wait before the next one
        doubles (Get-HeldBackRetryWait).
    #>
    param([string]$Reason, [datetime]$HeldAt, [switch]$FixHeld, [switch]$Retry)
    if (-not $script:PendingFixReason) {
        $script:PendingFixSince       = $HeldAt
        $script:PendingFixTries       = 0
        $script:PendingFixRetriesHeld = 0
        $script:PendingFixWaitLogged  = [datetime]::MinValue
        $script:PendingFixLockLogged  = $false
    }
    $script:PendingFixReason    = $Reason
    $script:PendingFixHeldAt    = $HeldAt
    $script:PendingFixNextCheck = [datetime]::MinValue
    if ($FixHeld) {
        $script:PendingFixTries++
        if ($Retry) { $script:PendingFixRetriesHeld++ }
    }
}

function Get-HeldBackRetryWait {
    <#
    .SYNOPSIS
        Minutes to wait after a hold before a repair is tried again.
    .DESCRIPTION
        2 minutes, the cooldown a hold leaves, doubled for each hold before
        it that counts, 60 at most: 2, 4, 8, 16, 32 and then 60 minutes, for
        0 to 5. Two callers count differently.

        Invoke-AutoFix counts the runs Fix held back in a row before this
        one, whatever the trigger, and sets the cooldown to it. A fault
        checked again on every poll, the service down or the vmcompute
        handle leak, then does not come back to the 30-second warning every
        few minutes for as long as Fix holds it back.

        Invoke-HeldBackRetry counts the retries of the pending repair that
        Fix held back, and waits that long after its latest hold. The third
        time Fix holds a pending repair back ends it (see Invoke-AutoFix),
        so there the longer waits are a bound, not a schedule.

        Whatever held Fix back, such as a user who moves the mouse when the
        warning comes up, is then not asked again every few minutes.
    .PARAMETER RetriesHeld
        The holds before this one that count.
    #>
    param([int]$RetriesHeld)
    return [int][math]::Min(60, 2 * [math]::Pow(2, [math]::Max(0, $RetriesHeld)))
}

function Get-ProcessStartUtc {
    <#
    .SYNOPSIS
        A process's start time as UTC; $null for one that has exited since
        it was listed; throws for one still running whose start time could
        not be read.
    .DESCRIPTION
        PowerShell reads a getter that throws as nothing (measured on 5.1,
        2026-11-22), so a start time that reads as nothing is either a
        process that has exited, routine with Electron's short-lived
        helpers, or one this monitor may not open, as one running with
        rights it lacks. HasExited tells them apart: its getter answers
        $true for a process that is gone, and for one this monitor may not
        open it throws Access is denied, which reads as nothing too, and
        [bool] makes that $false (measured on 5.1, 2026-11-22, for System,
        csrss and lsass). The getter method is then asked again, and throws
        the reason. Until 6.3.1 both were passed over, as if every one had
        exited.
    #>
    param($Process)
    $st = $Process.StartTime
    if ($null -ne $st) { return ([datetime]$st).ToUniversalTime() }
    if ([bool]$Process.HasExited) { return $null }
    return $Process.get_StartTime().ToUniversalTime()
}

function Get-ClaudeAppStartTime {
    <#
    .SYNOPSIS
        When Claude Desktop's main process started, as UTC, or $null when the
        app is not running or no start time could be read.
    .DESCRIPTION
        The oldest of Get-ClaudeDesktopProcess. Electron's main process
        starts every other one and they end with it, so a helper started
        again later, such as a GPU or renderer process, is younger and does
        not count. StartTime is read per process inside a try, as DEFERRED
        reads it: it throws for a process that has just exited.
    #>
    # A list that could not be read is said, at most once an hour, and the
    # answer is $null, which ends no held-back repair. Until 6.3.1 the list
    # could not fail here: one that could not be read was empty.
    try { $procs = @(Get-ClaudeDesktopProcess -Environment $script:ClaudeEnv) }
    catch {
        Write-CheckError -Check 'app start time' -ErrorRecord $_ `
            -Text "When Claude Desktop started could not be read ({0}), so a held-back repair is not ended on it"
        return $null
    }
    # A process that has exited since the list was read is not the main
    # process, and is passed over. One still running whose start time could
    # not be read is said, and the answer is $null: until 6.3.1 it was passed
    # over too, so if it was the main process the start of a younger helper
    # was taken for the app's, and could end a held-back repair.
    $oldest = $null
    foreach ($p in $procs) {
        try {
            $st = Get-ProcessStartUtc -Process $p
        } catch {
            Write-CheckError -Check 'app start time' -ErrorRecord $_ `
                -Text "When Claude Desktop started could not be read ({0}), so a held-back repair is not ended on it"
            return $null
        }
        if ($null -eq $st) { continue }
        if ($null -eq $oldest -or $st -lt $oldest) { $oldest = $st }
    }
    return $oldest
}

function Test-FixLogRepaired {
    <#
    .SYNOPSIS
        Whether a Fix log holds $FixRepairLine: that run got as far as the
        repair.
    .DESCRIPTION
        Opened to share reading, writing and deleting, as a transcript Fix
        is still writing, or a log it is rotating away, may be. A log that
        cannot be read answers $false: the repair stays pending, and at worst
        is tried again.
    #>
    param([string]$Path)
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read,
            ([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
        try {
            $reader = [System.IO.StreamReader]::new($stream, $true)
            return $reader.ReadToEnd().Contains($FixRepairLine)
        } finally {
            $stream.Dispose()
        }
    } catch {
        return $false
    }
}

function Get-FixRunAfter {
    <#
    .SYNOPSIS
        The latest time a Fix log that holds $FixRepairLine was created or
        written after -After, as UTC, or $null when there is none.
    .DESCRIPTION
        Fix writes fix_<time>_transcript.log in $FixLogDir from its start,
        and fix_<time>.log each time it saves its log, and both get
        $FixRepairLine once the run is past every check that holds it back.
        Such a log created or written after -After is a Fix run since then
        that started the repair: one run by hand, or the one that held Fix's
        lock when this monitor's run was held back and went on after it. A
        log without the line is a run that changed nothing, such as a
        Diagnostic run, a Cancel or a No at "Proceed?", and does not count:
        until the line was asked for, any Fix log did, so a Diagnostic run
        ended a pending repair as "the repair". This monitor's own run
        writes its logs before it returns, so before the hold that follows.
        A folder or a log that cannot be read answers as no run, and the
        repair stays pending. A folder that is there and cannot be read is
        logged, at most once an hour for the same error (Write-CheckError);
        until 6.3.0 it was passed over. No folder is no Fix run yet.
    #>
    param([datetime]$After)
    $latest = $null
    try {
        foreach ($f in @(Get-ChildItem -LiteralPath $FixLogDir -Filter 'fix_*.log' -File -ErrorAction Stop)) {
            $t = $f.LastWriteTimeUtc
            if ($f.CreationTimeUtc -gt $t) { $t = $f.CreationTimeUtc }
            if ($t -le $After -or ($null -ne $latest -and $t -le $latest)) { continue }
            if (Test-FixLogRepaired -Path $f.FullName) { $latest = $t }
        }
    } catch {
        if ([string]$_.FullyQualifiedErrorId -notlike 'PathNotFound*') {
            Write-CheckError -Check 'Fix logs' -ErrorRecord $_ `
                -Text "The Fix logs in $FixLogDir could not be read ({0}), so a Fix run is not seen there, and a held-back repair stays pending"
        }
    }
    return $latest
}

function Close-StaleHeldBackRepair {
    <#
    .SYNOPSIS
        Ends a pending held-back repair that is over, with one line that says
        why. Returns $true when it ended one.
    .DESCRIPTION
        Over is one of three things.

        4 hours since it was first held back, or since the latest one-shot
        trigger that took it over fired (Merge-HeldBackTrigger). By then what
        was wrong has had time to change, and a repair nobody has seen asked
        for in hours could close a session the user has just started. No hold
        starts the 4 hours again.

        Claude Desktop started again since then (Get-ClaudeAppStartTime):
        what was detected came from the app session that ended, and starting
        the app again is what a repair would have done.

        A Fix run since the latest hold that started the repair
        (Get-FixRunAfter): whoever ran Fix has run it. Not asked while a Fix
        run holds Fix's lock: the callers ask Test-FixLockHeld first, and
        that run's log says what it did once it is over.

        -AgeOnly asks the first only. It costs no process or folder listing,
        so it can be asked on every poll.
    #>
    param([switch]$AgeOnly)
    if (-not $script:PendingFixReason) { return $false }
    if (([datetime]::UtcNow - $script:PendingFixSince).TotalHours -ge 4) {
        Close-HeldBackRepair -Why 'dropped after 4 hours; run Fix-ClaudeDesktop.ps1 when you are ready'
        return $true
    }
    if ($AgeOnly) { return $false }
    $appStart = Get-ClaudeAppStartTime
    if ($null -ne $appStart -and $appStart -gt $script:PendingFixSince) {
        Close-HeldBackRepair -Why 'ended: Claude Desktop was started again since it was held back'
        return $true
    }
    $lastHold = $script:PendingFixHeldAt
    if ($lastHold -eq [datetime]::MinValue) { $lastHold = $script:PendingFixSince }
    if ($null -ne (Get-FixRunAfter -After $lastHold)) {
        Close-HeldBackRepair -Why 'ended: Fix-ClaudeDesktop ran after it was held back, so that run was the repair'
        return $true
    }
    return $false
}

function Test-FixLockHeld {
    <#
    .SYNOPSIS
        Whether a Fix-ClaudeDesktop run holds Fix's lock now.
    .DESCRIPTION
        Fix takes one of these two names, preferring Global and falling back
        to Local when it lacks SeCreateGlobalPrivilege. Both are checked here
        because missing the lock means starting a second repair on top of a
        running one.

        The name carries no version deliberately: two toolkit versions
        operating on the same service and the same VHDX files at once is
        exactly what the lock exists to prevent. This check used to look for
        "Global\ClaudeDesktopFix_v4.8", which stopped matching the moment
        Fix's own name moved on.

        Only a lock that does not exist is not held. One that exists but
        this monitor may not open, as when Fix runs elevated and the
        monitor does not, is held; and one that cannot be asked at all is
        taken as held, so no repair starts on top of a running one, and that
        is logged, at most once an hour for the same error
        (Write-CheckError). Until 6.3.0 any error read as not held.
    #>
    foreach ($fixLockName in @("Global\ClaudeDesktopFix", "Local\ClaudeDesktopFix")) {
        try {
            $existingLock = [System.Threading.Mutex]::OpenExisting($fixLockName)
            $existingLock.Dispose()
            return $true
        } catch [System.Threading.WaitHandleCannotBeOpenedException] {
            # Not held under this scope. Try the next one.
            $null = $_
        } catch [System.UnauthorizedAccessException] {
            # It exists, so a Fix run holds it.
            return $true
        } catch {
            Write-CheckError -Check "Fix's lock" -ErrorRecord $_ `
                -Text "Whether Fix-ClaudeDesktop is running could not be told from its lock $fixLockName ({0}); taken as running"
            return $true
        }
    }
    return $false
}

function Write-FixRunningSkip {
    <#
    .SYNOPSIS
        Says that a repair is skipped because a Fix run holds Fix's lock.
    .DESCRIPTION
        A Fix run by hand holds the lock while it waits in its menu, for as
        long as the user leaves it, and a re-checked trigger asks on every
        poll meanwhile, so the line is never written each time it is asked.

        With a repair pending, once while the lock stays held, with what
        becomes of the repair: the retry asks every 2 minutes and finds the
        lock gone once that run is over, which lets the next run be said
        again. With none, nothing asks between triggers, so the end of a run
        is not seen: at most once every 10 minutes instead (Test-LineDue).
        Until 6.3.0 it was written each time it was asked, up to 120 times
        an hour.
    #>
    if (-not $script:PendingFixReason) {
        if (Test-LineDue -Key 'SKIPPED|lock' -Now ([datetime]::UtcNow)) {
            Write-WatchLog "AUTO-FIX SKIPPED: Fix-ClaudeDesktop is already running"
        }
        return
    }
    if ($script:PendingFixLockLogged) { return }
    $script:PendingFixLockLogged = $true
    Write-WatchLog "AUTO-FIX SKIPPED: Fix-ClaudeDesktop is already running"
    Write-WatchLog "    The held-back repair for $($script:PendingFixReason) is kept while that run lasts: it ends if the run gets as far as the repair, and is tried again if the run stops before it, as a Diagnostic run does"
}

function Merge-HeldBackTrigger {
    <#
    .SYNOPSIS
        Merges a one-shot trigger that fired on its own into the pending
        held-back repair, and saves the state.
    .DESCRIPTION
        A pending repair that is over ends first (Close-StaleHeldBackRepair),
        so the trigger starts one of its own. Otherwise the trigger says what
        is wrong now: the pending repair is retried for it, and its 4 hours
        start again from -Now, because the trigger is new evidence. Its tries
        and the retries held back are kept, so the 30-second warning still
        comes at most 3 times, and a hold alone never starts the 4 hours
        again. Without the new start, a log error that fired 2 minutes
        before the mark would be dropped "after 4 hours" with the repair it
        merged into, and never tried. The start is persisted with the rest,
        so a restart of the monitor keeps it.

        -AgeOnly asks only the 4 hours, for the lock path: while a Fix run
        holds the lock, the Fix logs are not read until it is over.

        Returns $true when the trigger took the pending repair over, so the
        caller can say so, and $false when none was pending or it was over.
    #>
    param([string]$Reason, [datetime]$Now, [switch]$AgeOnly)
    if (-not $script:PendingFixReason) { return $false }
    if (Close-StaleHeldBackRepair -AgeOnly:$AgeOnly) { return $false }
    if (-not $script:PendingFixReason) { return $false }
    $script:PendingFixReason = $Reason
    $script:PendingFixSince  = $Now
    Save-WatchState
    return $true
}

function Get-ClaudeCodeSessionHold {
    <#
    .SYNOPSIS
        Why Fix's -Quiet gate would hold a repair back for a Claude Code
        session, or $null when it would not.
    .DESCRIPTION
        Fix's gate (Get-AutomatedFixBlocker) asks two things about Claude
        Code first, whether or not the app runs: whether the app runs a Code
        session (Get-ClaudeAppCliProcess), and whether a transcript was
        written in the last 35 minutes (Test-ClaudeCodeTranscriptRecent).
        Test-UserActivity asks both too, through Test-ClaudeCodeActive, but
        only when a process named claude is running. With none, such as with
        the app closed and a Code session that ended minutes ago, or one in a
        terminal whose process has another name, the monitor showed the
        30-second warning, Fix held the run back, and a fault checked on
        every poll came back to the warning again. Asked here as Fix asks
        them, BLOCKED stops that run before the warning.

        Fails closed, as Fix's gate does for the app's sessions: a list or a
        folder that cannot be read is a reason too.
    #>
    # Each reason that could not be read says why, as Fix's gate does. Until
    # 6.3.1 the monitor's said only that it could not.
    try {
        $appCli = @(Get-ClaudeAppCliProcess -Environment $script:ClaudeEnv)
        if ($appCli.Count -gt 0) { return "$($appCli.Count) Claude Code session(s) the app started are running" }
    } catch {
        return "the Claude Code sessions the app started could not be listed ($(Get-ClaudeErrorText $_))"
    }
    try {
        if (Test-ClaudeCodeTranscriptRecent -Minutes 35) { return "a Claude Code transcript was written in the last 35 minutes" }
    } catch {
        return "whether a Claude Code transcript was written in the last 35 minutes could not be read ($(Get-ClaudeErrorText $_))"
    }
    return $null
}

function Test-LineDue {
    <#
    .SYNOPSIS
        Whether a line that repeats may be written now: $true at most once
        every 10 minutes for each key.
    .DESCRIPTION
        For three lines. BLOCKED and COOLDOWN, which a trigger checked again
        on every poll reaches each time it fires while the hold or the wait
        lasts; the lines that follow COOLDOWN are written with it or not at
        all. And AUTO-FIX SKIPPED from Write-FixRunningSkip, which any
        trigger reaches while a Fix run holds Fix's lock and no repair is
        pending, under the one key SKIPPED|lock rather than a reason.
        Keys 10 minutes old are dropped first, so the table stays small; a
        time in the future, after the clock was set back, counts as old. In
        memory only: after a restart each line is written again once.
    #>
    param([string]$Key, [datetime]$Now)
    foreach ($k in @($script:RepeatLinesSaid.Keys)) {
        $age = ($Now - $script:RepeatLinesSaid[$k]).TotalMinutes
        if ($age -ge 10 -or $age -lt 0) { $script:RepeatLinesSaid.Remove($k) }
    }
    if ($script:RepeatLinesSaid.ContainsKey($Key)) { return $false }
    $script:RepeatLinesSaid[$Key] = $Now
    return $true
}

function Invoke-AutoFix {
    param(
        [string]$Reason,
        # A one-shot trigger: one that does not fire again for the same fault,
        # a log error or an Event Log error. A hold, Fix's exit 3, Fix's lock
        # held, or BLOCKED or DEFERRED here, keeps its repair pending for
        # Invoke-HeldBackRetry, or it would be lost. Without it the trigger is re-checked, the
        # service or the vmcompute check: asked again on every poll, it fires
        # again while the fault lasts, so its repair is not kept.
        [switch]$OneShot,
        # Invoke-HeldBackRetry's call, for a repair held back. $Reason is
        # then the reason it was held back for, and is logged with a prefix.
        [switch]$Retry
    )

    # What a hold keeps as the pending repair's reason: without the prefix,
    # so a retry held back again does not stack it.
    $pendingReason = $Reason
    # Whether a hold keeps this repair pending: a one-shot trigger's, or a
    # retry's.
    $keepOnHold = ($OneShot -or $Retry)
    if ($Retry) { $Reason = "retry of held-back repair: $Reason" }

    # Do not trigger auto-fix if Fix-ClaudeDesktop is already running: see
    # Test-FixLockHeld.
    #
    # A repair pending meanwhile is not ended while that run holds the lock.
    # The run may be the repair, or one that changes nothing: a Fix run by
    # hand takes the lock and waits in its menu, and the user may pick
    # Diagnostic or Cancel, or answer No at "Proceed?". Its log says which
    # once it is over, and Invoke-HeldBackRetry reads it then
    # (Close-StaleHeldBackRepair). Until then the pending repair is kept,
    # and said to be once (Write-FixRunningSkip). A one-shot trigger that
    # fires now is kept as Fix's own hold at its lock keeps it: merged into
    # the pending repair, or pending as a new one.
    if (Test-FixLockHeld) {
        if ($OneShot -and -not $Retry) {
            $null = Merge-HeldBackTrigger -Reason $Reason -Now ([datetime]::UtcNow) -AgeOnly
            if (-not $script:PendingFixReason) {
                Register-HeldBackRepair -Reason $pendingReason -HeldAt ([datetime]::UtcNow)
                Save-WatchState
            }
        }
        Write-FixRunningSkip
        return
    }
    $script:PendingFixLockLogged = $false

    # A one-shot trigger that fires on its own while a repair is pending
    # takes it over: see Merge-HeldBackTrigger. First, so the time it is
    # given is when the trigger fired, before any Fix log is read. What it
    # took over is kept for the COOLDOWN line below, $null when nothing. The
    # trigger the repair is pending for, fired again, takes nothing over,
    # and its 4 hours start again all the same: $firedAgain.
    $tookOverFrom = $null
    $firedAgain = $false
    if ($OneShot -and -not $Retry) {
        $pendingBefore = $script:PendingFixReason
        if (Merge-HeldBackTrigger -Reason $Reason -Now ([datetime]::UtcNow)) {
            if ($pendingBefore -eq $Reason) { $firedAgain = $true } else { $tookOverFrom = $pendingBefore }
        }
    }

    # A Fix run by hand that started the repair since the latest run Fix
    # held back ends those runs in a row, as a run Fix does not hold back
    # ends them: the count starts again, and the wait they left ends 5
    # minutes after that run at the latest, the cooldown after a repair. It
    # never makes the wait longer: a run by hand outside such runs does not
    # touch it. Asked here, before the cooldown, so that a fault found after
    # the user's repair is not held off for up to an hour by holds from
    # before it. Asked only while the count is above 0, and Get-FixRunAfter
    # reads the text only of the Fix logs written since that hold.
    if ($script:FixHeldInARow -gt 0) {
        $manualAt = Get-FixRunAfter -After $script:FixHeldLastTime
        if ($null -ne $manualAt) {
            $script:FixHeldInARow    = 0
            $script:FixHeldLastTime  = [datetime]::MinValue
            $script:CooldownFromHold = $false
            if ($script:LastFixTime -gt $manualAt) { $script:LastFixTime = $manualAt }
            Save-WatchState
        }
    }

    # UTC, as every time this monitor keeps.
    $now = [datetime]::UtcNow
    $elapsed = ($now - $script:LastFixTime).TotalMinutes

    if ($elapsed -lt $Cooldown) {
        # Says when a fix is allowed again, not how long ago one ran. BLOCKED,
        # DEFERRED, CANCELLED and HELD BACK shorten the wait by moving
        # $script:LastFixTime back, with no fix run. Until 6.3.0 this said
        # "last fix N min ago", counted from that time, so after a BLOCKED it
        # named a fix 3 minutes old that never ran.
        #
        # At most once every 10 minutes for the same reason (Test-LineDue):
        # after runs Fix held back in a row the wait can be an hour, and a
        # re-checked trigger comes here each time it fires meanwhile.
        #
        # A one-shot trigger with no repair pending is kept as one when a
        # hold left this wait ($script:CooldownFromHold), and
        # Invoke-HeldBackRetry tries it once the wait is over: the trigger
        # does not fire again, and after runs Fix held back in a row the wait
        # can be an hour, so without this it would be lost. In the cooldown
        # after a repair that ran it is skipped, as before: that repair has
        # just restarted the app and the service, which is what a log error
        # then most likely comes from, and the cooldown is there so that a
        # second repair does not start on top of the first. After an
        # escalation or a cancel too: the user has been told to run Fix by
        # hand. A repair pending already has taken the trigger over above,
        # and is tried again for it the same way, so a second line says so
        # too, when COOLDOWN is written: COOLDOWN alone read as if the
        # trigger were dropped, although a PRE-FIX WARNING for it follows.
        # The trigger it is pending for, fired again, is said to be tried
        # again, with its 4 hours from now: it takes over nothing, and
        # naming it as the repair it took over read as if it had.
        $sayCooldown = Test-LineDue -Key "COOLDOWN|$Reason" -Now $now
        if ($sayCooldown) {
            Write-WatchLog "COOLDOWN: Skipping fix ($Reason), a fix is allowed again in $([math]::Round($Cooldown - $elapsed, 1)) min"
            if ($tookOverFrom) {
                Write-WatchLog "    It takes over the held-back repair for $tookOverFrom, and is tried again once the wait is over, for up to 4 hours"
            } elseif ($firedAgain) {
                Write-WatchLog "    It is tried again once the wait is over, for up to 4 hours"
            }
        }
        if ($OneShot -and -not $Retry -and $script:CooldownFromHold -and -not $script:PendingFixReason) {
            Register-HeldBackRepair -Reason $pendingReason -HeldAt $now
            Save-WatchState
            if ($sayCooldown) { Write-WatchLog "    It is tried again once the wait is over, for up to 4 hours" }
        }
        return
    }

    # Persistent failure escalation
    if (Test-PersistentFailure) {
        Write-WatchLog "ESCALATION: Fix attempted 3+ times in 30 min, backing off"
        Write-WatchLog "  This likely requires a Hyper-V nuclear reset:"
        Write-WatchLog "  1. Open admin PowerShell"
        Write-WatchLog "  2. dism /online /disable-feature /featurename:Microsoft-Hyper-V-All"
        Write-WatchLog "  3. Reboot"
        Write-WatchLog "  4. dism /online /enable-feature /featurename:Microsoft-Hyper-V-All"
        Write-WatchLog "  5. Reboot"
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
            $notifyIcon = New-Object System.Windows.Forms.NotifyIcon
            $notifyIcon.Icon = [System.Drawing.SystemIcons]::Error
            $notifyIcon.Visible = $true
            $notifyIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Error
            $notifyIcon.BalloonTipTitle = "Claude Health Monitor, Manual Fix Required"
            $notifyIcon.BalloonTipText = "Auto-fix failed 3 times. Hyper-V nuclear reset needed.`nSee watch log for instructions."
            $notifyIcon.ShowBalloonTip(60000)
            [System.Media.SystemSounds]::Hand.Play()
        } catch {
            # No interactive desktop (Session 0) or the assemblies are
            # absent: the steps are in the log above all the same.
            $null = $_
        }
        # Back off for the 30 minutes that Test-PersistentFailure measures over.
        #
        # This used to add $Cooldown, which defaults to 5, so the "back off"
        # lasted 10 minutes, not 30. Test-PersistentFailure looks at a 30-minute
        # window, so the escalation stayed true and the balloon plus alert sound
        # repeated every 10 minutes indefinitely.
        $script:LastFixTime = [datetime]::UtcNow.AddMinutes(30 - $Cooldown)
        $script:CooldownFromHold = $false
        return
    }

    # ---- SAFETY: Never auto-fix while user is active ----
    # What Fix's -Quiet gate asks about Claude Code sessions is asked first
    # (Get-ClaudeCodeSessionHold), so the 30-second warning is not shown for
    # a run that gate would hold back for a Code session: Test-UserActivity
    # asks nothing when no process named claude is running, and Fix's gate
    # still asks about the transcripts then.
    $codeHold = Get-ClaudeCodeSessionHold
    if ($codeHold -or (Test-UserActivity)) {
        if ($codeHold) {
            $blockedWhy   = $codeHold
            $blockedUntil = 'once that is over'
        } else {
            $blockedWhy   = 'user is active (Claude in focus, recent input, or VM busy)'
            $blockedUntil = 'once the user is away'
        }
        # At most once every 10 minutes for the same reason (Test-LineDue):
        # a re-checked trigger comes back here each time it fires while the
        # hold lasts, which wrote these lines about 30 times an hour.
        $sayBlocked = Test-LineDue -Key "BLOCKED|$Reason" -Now $now
        if ($sayBlocked) { Write-WatchLog ">>> BLOCKED: $Reason, $blockedWhy" }
        # Set cooldown so we don't spam the log every 30s: a trigger that
        # fires again gets through again in about 2 minutes. A pending
        # held-back repair waits longer when Invoke-HeldBackRetry's own wait
        # is longer: Get-HeldBackRetryWait from its latest hold, which is
        # this one when the repair is kept below, 2 minutes doubled for each
        # retry Fix has held back.
        $script:LastFixTime = $now.AddMinutes(-($Cooldown - 2))
        $script:CooldownFromHold = $true
        # A one-shot trigger's repair, or a retry, is kept pending as HELD
        # BACK keeps it below, and Invoke-HeldBackRetry tries it again once
        # the hold is over. Until 6.3.0 nothing did, and a log error BLOCKED
        # here was lost. Not counted toward the 3 tries: no warning was shown.
        # A re-checked trigger fires again on its own, and a pending repair
        # stays as it is.
        if ($keepOnHold) {
            Register-HeldBackRepair -Reason $pendingReason -HeldAt $now
            Save-WatchState
            if ($sayBlocked) {
                Write-WatchLog "    It is tried again $blockedUntil, for up to 4 hours, or run Fix-ClaudeDesktop.ps1 manually when ready"
            }
        } elseif ($sayBlocked) {
            Write-WatchLog "    Run Fix-ClaudeDesktop.ps1 manually when ready"
        }
        return
    }

    # ---- SAFETY: Don't auto-fix during Claude startup ----
    # When Claude just launched, workspace initialization can take 30-60 seconds
    # and may show transient errors. Give it time to settle.
    #
    # The app's processes only. This measures how long ago the APP started, and
    # the app starts a fresh Claude Code CLI process for project threads as they
    # run, so counting those would read every thread wake as an app launch and
    # could defer a repair for as long as threads keep waking.
    #
    # A list that could not be read defers the repair, as a start too recent
    # does: whether the app has just started is not known. Until 6.3.1 such a
    # list was empty, and the repair went ahead.
    $claudeStartupProcs = @()
    $startupUnread = $null
    $retryWhen = 'once they can be listed'
    try {
        $claudeStartupProcs = @(Get-ClaudeDesktopProcess -Environment $script:ClaudeEnv)
    } catch {
        $startupUnread = "Claude's processes could not be listed ($(Get-CheckErrorText -ErrorRecord $_))"
    }
    # Each start time through Get-ProcessStartUtc. A process that exited
    # between the list and this line has none, and with sixteen short-lived
    # Electron helpers around, that is a routine occurrence rather than an
    # edge case: it is passed over. One still running whose start time could
    # not be read defers the repair, as a list that could not be read does:
    # it may be the app, just started. Until 6.3.1 it was passed over too.
    # UTC, for the same reason as Test-ClaudeCodeActive's ages: at the
    # change to summer time a process started a minute ago read as an hour
    # and a minute old.
    $youngest = $null
    if (-not $startupUnread) {
        foreach ($sp in $claudeStartupProcs) {
            try {
                $st = Get-ProcessStartUtc -Process $sp
            } catch {
                $startupUnread = "the start time of one of Claude's processes could not be read ($(Get-CheckErrorText -ErrorRecord $_))"
                $retryWhen = 'once it can be read'
                break
            }
            if ($null -ne $st -and ($null -eq $youngest -or $st -gt $youngest)) { $youngest = $st }
        }
    }
    if ($startupUnread) {
        Write-WatchLog ">>> DEFERRED: $Reason, $startupUnread, so whether it has just started is not known"
        $script:LastFixTime = $now.AddMinutes(-($Cooldown - 1))
        $script:CooldownFromHold = $true
        if ($keepOnHold) {
            Register-HeldBackRepair -Reason $pendingReason -HeldAt $now
            Save-WatchState
            Write-WatchLog "    It is tried again $retryWhen, for up to 4 hours"
        }
        return
    }
    if ($claudeStartupProcs.Count -gt 0) {
        $procAge = if ($youngest) { ([datetime]::UtcNow - $youngest.ToUniversalTime()).TotalSeconds } else { [double]::MaxValue }
        if ($procAge -lt 120) {
            Write-WatchLog ">>> DEFERRED: $Reason, Claude launched $([math]::Round($procAge))s ago (waiting for startup to complete)"
            $script:LastFixTime = $now.AddMinutes(-($Cooldown - 1))
            $script:CooldownFromHold = $true
            # Kept pending for a one-shot trigger or a retry, as BLOCKED keeps
            # it, and tried again once startup is over. Until 6.3.0 a log
            # error DEFERRED here was lost.
            if ($keepOnHold) {
                Register-HeldBackRepair -Reason $pendingReason -HeldAt $now
                Save-WatchState
                Write-WatchLog "    It is tried again once startup is over, for up to 4 hours"
            }
            return
        }
    }

    # ---- PRE-FIX WARNING: 30s grace period with notification ----
    # Show a balloon tip so the user knows what's about to happen.
    # The notification says "open Claude to cancel" because we only cancel
    # if Claude becomes actively used -- not just because the mouse moved.
    # This way, a genuinely hung VM still gets fixed even if the user is
    # browsing, gaming, or otherwise using the PC.
    Write-WatchLog ">>> PRE-FIX WARNING: $Reason, auto-fix in 30s (open Claude to cancel)"
    $notifyIcon = $null
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
        $notifyIcon = New-Object System.Windows.Forms.NotifyIcon
        $notifyIcon.Icon = [System.Drawing.SystemIcons]::Warning
        $notifyIcon.Visible = $true
        $notifyIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Warning
        $notifyIcon.BalloonTipTitle = "Claude Health Monitor"
        $notifyIcon.BalloonTipText = "VM appears hung. Auto-fix in 30s.`nSwitch to Claude or use Code to cancel."
        $notifyIcon.ShowBalloonTip(30000)
        # Play the Windows "Exclamation" sound so the user hears it even if not looking
        [System.Media.SystemSounds]::Exclamation.Play()
    } catch {
        # No interactive desktop (Session 0) or the assemblies are absent. The
        # repair does not depend on the notification, so carry on regardless.
        $null = $_
    }

    Start-Sleep -Seconds 30

    # Dismiss notification. One that was never shown, or is gone already,
    # has nothing to dismiss.
    try { if ($notifyIcon) { $notifyIcon.Visible = $false; $notifyIcon.Dispose() } } catch { $null = $_ }

    # Re-check: only cancel if Claude is ACTIVELY being used.
    # We check Claude-specific signals (foreground, CPU, VM log) but NOT
    # general user input -- the user may be at the keyboard in another app
    # while Claude's VM is genuinely dead.
    $cancelFix = $false
    $cancelReason = ""
    # Every process named claude, the Claude Code CLI included: this can only
    # cancel the repair, and erring that way is safe. A list that could not
    # be read cancels it, as does each read below that fails: until 6.3.1
    # each read as nothing, and the repair went ahead.
    $claudeProcs = @()
    try {
        $claudeProcs = @(Get-ClaudeProcessByName -Name $ProcessName)
    } catch {
        $cancelFix = $true
        $cancelReason = "Claude's processes could not be listed ($(Get-CheckErrorText -ErrorRecord $_))"
    }
    if ($claudeProcs.Count -gt 0) {
        # Cancel if Claude is the foreground window
        if ($script:IsInteractiveSession -and $script:Win32ActivityLoaded) {
            try {
                $fgHwnd = [Win32Activity]::GetForegroundWindow()
                if ($fgHwnd -ne [IntPtr]::Zero) {
                    $fgPid = 0
                    [Win32Activity]::GetWindowThreadProcessId($fgHwnd, [ref]$fgPid) | Out-Null
                    if ($fgPid -gt 0) {
                        foreach ($cp in $claudeProcs) {
                            if ($cp.Id -eq $fgPid) { $cancelFix = $true; $cancelReason = "Claude is now in focus"; break }
                        }
                        # A window whose process exited since it was read
                        # has no entry, and is not Claude's. A read that
                        # fails goes to the catch below, which says so.
                        # Until 6.3.1 it was passed over in silence.
                        if (-not $cancelFix) {
                            $fgProc = @(Get-CimInstance Win32_Process -Filter "ProcessId=$fgPid" -ErrorAction Stop)
                            if ($fgProc.Count -gt 0) {
                                $parentId = $fgProc[0].ParentProcessId
                                foreach ($cp in $claudeProcs) {
                                    if ($cp.Id -eq $parentId) { $cancelFix = $true; $cancelReason = "Claude is now in focus"; break }
                                }
                            }
                        }
                    }
                }
            } catch {
                Write-CheckError -Check 'focus before repair' -ErrorRecord $_ `
                    -Text "Whether Claude is in focus could not be read before the repair ({0}); the other checks still decide"
            }
        }

        # Cancel if any Claude process is burning CPU
        # Extended sampling: 3 checks x 1s to catch bursty Code patterns
        # (Code has near-zero CPU during API waits but spikes during processing)
        if (-not $cancelFix) {
            for ($sample = 1; $sample -le 3; $sample++) {
                if ($cancelFix) { break }
                foreach ($cp in $claudeProcs) {
                    try {
                        $cpu1 = $cp.TotalProcessorTime.TotalMilliseconds
                        Start-Sleep -Milliseconds 1000
                        $cp.Refresh()
                        $cpu2 = $cp.TotalProcessorTime.TotalMilliseconds
                        if (($cpu2 - $cpu1) -gt 50) { $cancelFix = $true; $cancelReason = "Claude CPU active (sample $sample/3)"; break }
                    } catch {
                        # Exited during the sample: not busy. Anything else,
                        # or an end that cannot be read, cancels, and the
                        # reason names both. get_HasExited(), not
                        # .HasExited: a getter that throws reads as nothing
                        # (measured 2026-11-23 on Windows PowerShell 5.1),
                        # so until 6.3.1 the catch here never ran and an end
                        # that could not be read was not said.
                        $sampleError = Get-CheckErrorText -ErrorRecord $_
                        $sampleGone = $false
                        $endError = $null
                        try { $sampleGone = [bool]$cp.get_HasExited() } catch { $endError = $_ }
                        if (-not $sampleGone) {
                            $cancelFix = $true
                            $cancelReason = "whether claude process $($cp.Id) is using CPU could not be read ($sampleError)"
                            if ($endError) { $cancelReason += ", nor whether it has ended ($(Get-CheckErrorText -ErrorRecord $endError))" }
                            break
                        }
                    }
                }
            }
        }
    }

    # Cancel if VM log became active during the 30s window
    if (-not $cancelFix) {
        # Any VM log, read now: the same test as Test-UserActivity's check 3.
        try {
            $ageSec = Get-ClaudeVmLogAge -LogDir $script:ClaudeLogDir
            if ($null -ne $ageSec -and $ageSec -lt 120) { $cancelFix = $true; $cancelReason = "VM log became active" }
        } catch {
            $cancelFix = $true
            $cancelReason = "the VM logs could not be read ($(Get-CheckErrorText -ErrorRecord $_))"
        }
    }

    # Cancel if a Claude Code session is active
    # This is the critical check: Code doesn't use the VM, so VM staleness
    # is expected during Code usage. Killing Claude kills the Code session.
    if (-not $cancelFix) {
        if (Test-ClaudeCodeActive) { $cancelFix = $true; $cancelReason = "Claude Code session active" }
    }

    # Cancel if a Cowork session is active (may have started during grace period)
    if (-not $cancelFix) {
        if (Test-CoworkSessionActive) { $cancelFix = $true; $cancelReason = "Cowork session active" }
    }

    if ($cancelFix) {
        Write-WatchLog ">>> CANCELLED: $cancelReason during 30s grace period"

        # Tell the user the VM may still need repair -- they came back to Claude
        # but the underlying issue hasn't gone away.
        try {
            $cancelNotify = New-Object System.Windows.Forms.NotifyIcon
            $cancelNotify.Icon = [System.Drawing.SystemIcons]::Information
            $cancelNotify.Visible = $true
            $cancelNotify.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
            $cancelNotify.BalloonTipTitle = "Claude Health Monitor"
            $cancelNotify.BalloonTipText = "Auto-fix cancelled (you're using Claude).`nIf Cowork is broken, run Fix-ClaudeDesktop.bat"
            $cancelNotify.ShowBalloonTip(15000)
            [System.Media.SystemSounds]::Asterisk.Play()
            # Clean up after a delay so the balloon stays visible
            Start-Sleep -Seconds 16
            $cancelNotify.Visible = $false
            $cancelNotify.Dispose()
        } catch {
            # No interactive desktop or no assemblies, as for the warning's
            # balloon: the CANCELLED line above is in the log.
            $null = $_
        }

        # A trigger that fires again gets through about 2 minutes from now.
        # Until 6.3.0 this was counted from $now, which was taken before the
        # 30-second warning, the samples and the 16 seconds the balloon is
        # shown, so the check came back a minute or so after the cancel.
        $script:LastFixTime = [datetime]::UtcNow.AddMinutes(-($Cooldown - 2))
        $script:CooldownFromHold = $false
        # The cancel ends a pending held-back repair: the user is at Claude,
        # and the balloon has told them to run Fix if Cowork is broken.
        if ($script:PendingFixReason) {
            Clear-HeldBackRepair
            Save-WatchState
        }
        return
    }

    $script:FixCount++
    Write-WatchLog ">>> AUTO-FIX #$($script:FixCount) TRIGGERED: $Reason"
    Write-WatchLog "    (Claude still inactive after 30s warning)"
    $script:LastFixTime = $now
    $script:CooldownFromHold = $false

    # Fix-ClaudeDesktop self-elevates with Start-Process -Verb RunAs -Wait when
    # it is not already running as admin. Invoking it in-process from an
    # unattended monitor therefore either raises a UAC consent dialog with
    # nobody there to answer it, blocking this monitor for as long as it sits
    # on screen, or fails outright. Neither is a repair.
    #
    # So only call it when this process already holds the token Fix needs. The
    # cooldown was set just above, so this logs at most once per cooldown
    # window rather than on every poll.
    #
    # A check that fails says so, naming why: until 6.3.0 it read as "not
    # elevated", which sent the user to reinstall a task that was fine.
    $watchIsAdmin = $false
    $adminCheckError = $null
    try {
        $watchIsAdmin = ([Security.Principal.WindowsPrincipal] `
            [Security.Principal.WindowsIdentity]::GetCurrent()
            ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { $adminCheckError = Get-CheckErrorText -ErrorRecord $_ }

    if (-not $watchIsAdmin) {
        if ($adminCheckError) {
            Write-WatchLog ">>> AUTO-FIX #$($script:FixCount) SKIPPED: whether this monitor is elevated could not be read ($adminCheckError), so Fix was not started: unelevated, it would raise a UAC prompt with nobody to answer it. Run Fix-ClaudeDesktop manually."
        } else {
            Write-WatchLog ">>> AUTO-FIX #$($script:FixCount) SKIPPED: this monitor is not elevated, so Fix would raise a UAC prompt with nobody to answer it. Reinstall the monitor task with highest privileges, or run Fix-ClaudeDesktop manually."
        }
        # Fix cannot be started from here, so a pending held-back repair ends:
        # retried, it would only show the warning again every cooldown.
        if ($script:PendingFixReason) {
            Clear-HeldBackRepair
            Save-WatchState
        }
        return
    }

    $fixExit = $null
    try {
        & $fixScript -Mode Smart -Quiet
        $fixExit = $LASTEXITCODE
    } catch {
        Write-WatchLog ">>> AUTO-FIX #$($script:FixCount) FAILED: $($_.Exception.Message)"
    }

    # Fix held the run back before it stopped the app or the service: at its
    # start, right before Step 1's kill, or another Fix run held the lock. It
    # is not recorded as an attempt at a repair: not toward the escalation,
    # and the counters and baselines keep what was detected.
    #
    # A one-shot trigger's repair, or a retry, is kept pending, and
    # Invoke-HeldBackRetry tries it again once the hold has cleared, so a
    # trigger that does not fire again, such as a log error, is not lost. Its
    # 4 hours run from when it was first held back, or from the latest
    # one-shot trigger that took it over (Merge-HeldBackTrigger), and a retry
    # held back again keeps that time: one that ran past the 4 hours is
    # dropped here. The third time Fix holds a repair back after the
    # 30-second warning, the first included, it is dropped too, and each
    # retry Fix holds back doubles the wait before the next one. Whatever
    # held Fix back each time, such as a user who moves the mouse when the
    # warning comes up, is disturbed by every warning more. A re-checked
    # trigger's repair is not kept: the check fires again while the fault
    # lasts.
    #
    # Each run Fix holds back in a row doubles the wait before the next
    # repair, whatever the trigger: 2, 4, 8, 16, 32, then 60 minutes
    # (Get-HeldBackRetryWait), one count, kept in watch-state.json. A fault
    # checked again on every poll, the service down or the vmcompute leak,
    # otherwise came back to the 30-second warning every 2 to 5 minutes for
    # as long as Fix held it back. A run Fix does not hold back starts the
    # count again; so does a Fix run by hand that started the repair since
    # the latest hold, seen above before the cooldown; and so does a hold 4
    # hours or more after the one before, which belongs to a fault of its
    # own and is seen here. The wait is kept in $script:LastFixTime, so a
    # one-shot trigger that fires during it is kept for after it (the
    # cooldown above).
    #
    # The wait runs from when Fix returned, not from $now, which was taken
    # before the 30-second warning. Fix can run for minutes before Step 1,
    # and 2 minutes from before all that could be over by the time it ends.
    if ($fixExit -eq $FixHeldBackExitCode) {
        $heldAt = [datetime]::UtcNow
        if (($heldAt - $script:FixHeldLastTime).TotalHours -ge 4) { $script:FixHeldInARow = 0 }
        $script:FixHeldInARow++
        $script:FixHeldLastTime = $heldAt
        $script:CooldownFromHold = $true
        $heldWait = Get-HeldBackRetryWait -RetriesHeld ($script:FixHeldInARow - 1)
        $script:LastFixTime = $heldAt.AddMinutes($heldWait - $Cooldown)
        $held = ">>> AUTO-FIX #$($script:FixCount) HELD BACK by Fix (exit $fixExit): not counted as an attempt"
        $dropWhy = $null
        if (-not $keepOnHold) {
            $held += '; the check fires again while the fault lasts'
            # A one-shot repair pending meanwhile now waits from this hold,
            # and this run's own Fix log is not read as a Fix run since.
            if ($script:PendingFixReason) { $script:PendingFixHeldAt = $heldAt }
        } else {
            Register-HeldBackRepair -Reason $pendingReason -HeldAt $heldAt -FixHeld -Retry:$Retry
            if (($heldAt - $script:PendingFixSince).TotalHours -ge 4) {
                $dropWhy = 'dropped after 4 hours; run Fix-ClaudeDesktop.ps1 when you are ready'
            } elseif ($script:PendingFixTries -ge 3) {
                $dropWhy = 'dropped after 3 tries; run Fix-ClaudeDesktop.ps1 when you are ready'
            } else {
                $held += '; it is retried when the hold clears, for up to 4 hours'
            }
        }
        Write-WatchLog $held
        if ($heldWait -gt 2) {
            Write-WatchLog "    Fix has held back $($script:FixHeldInARow) runs in a row, so the next repair waits $heldWait minutes"
        }
        if ($dropWhy) {
            Close-HeldBackRepair -Why $dropWhy
            return
        }
        Save-WatchState
        return
    }
    # Any other end means the repair ran, or could not start: a pending
    # held-back repair ends, and is saved with the history below, and the
    # count of runs Fix held back in a row starts again.
    Clear-HeldBackRepair
    $script:FixHeldInARow   = 0
    $script:FixHeldLastTime = [datetime]::MinValue
    # Fix ran the repair part of the way and stopped before going further,
    # because it found a Claude Code session, or could not list the sessions
    # to rule one out. It had run Steps 1 and 2, which stop the app and the
    # service when they are running, so it is an attempt like any other:
    # recorded below toward the escalation, what was detected is reset, and
    # the normal cooldown applies.
    if ($fixExit -eq $FixStoppedPartWayExitCode) {
        Write-WatchLog ">>> AUTO-FIX #$($script:FixCount) STOPPED PART-WAY by Fix (exit $fixExit): the repair ran part of the way, then Fix stopped before going further: a Claude Code session was found, or could not be ruled out (see the Fix log); counted as an attempt"
    } elseif ($null -ne $fixExit) {
        Write-WatchLog ">>> AUTO-FIX #$($script:FixCount) COMPLETE (exit $fixExit)"
    }

    # Record fix timestamp for persistent failure detection, and persist it so
    # the backoff survives the restart that a bad repair loop tends to cause.
    $null = $script:FixHistory.Add([datetime]::UtcNow)
    Save-WatchState

    # Reset state after fix.
    #
    # Clearing the table is enough: Test-LogsForErrors reseeds each stream at
    # its current size the first time it sees it again. The loop that used to
    # be here reseeded using the bare file path as the key, which no longer
    # matches the (path, creation time) key the detector uses, so every entry
    # it wrote would have been ignored and then re-seeded anyway.
    Start-Sleep -Seconds 10
    $script:LogBaselines = @{}
    $script:VmLogStaleCount    = 0
    $script:VmLogEverActive    = $false
    $script:ServiceDownCount   = 0
    $script:EventLogHitCount   = 0
    $script:VmcomputeLeakCount = 0
}

function Invoke-HeldBackRetry {
    <#
    .SYNOPSIS
        Tries again a repair held back for a one-shot trigger, once the hold
        has cleared.
    .DESCRIPTION
        Called on a poll where no check fired, with the app running. A hold is
        Fix's exit 3, Fix's lock held, BLOCKED or DEFERRED in Invoke-AutoFix,
        or the wait one of them leaves, for a one-shot trigger that fired in
        it. A one-shot trigger that fires on its own goes to Invoke-AutoFix as
        before, and takes the pending repair over there, which starts its 4
        hours again (Merge-HeldBackTrigger).

        Asked once the cooldown a hold leaves is over, doubled for each run
        Fix held back in a row, and the wait after the latest hold has passed
        (Get-HeldBackRetryWait: 2 minutes, doubled each time Fix holds a
        retry back), and then no more than every 2 minutes.

        While a Fix run holds Fix's lock it waits, and says so once
        (Write-FixRunningSkip): the run may be the repair, or one that
        changes nothing, such as a Diagnostic run, and its log says which
        once it is over. Then it ends if it is over
        (Close-StaleHeldBackRepair): 4 hours old, the app started again, or a
        Fix run since the latest hold that started the repair. Then it asks,
        silently, what Fix's -Quiet gate asks about Claude Code sessions
        (Get-ClaudeCodeSessionHold: one the app started, or a transcript in
        the last 35 minutes), whether Test-ClaudeCodeActive sees a session,
        and whether Test-UserActivity, which BLOCKED asks, sees the user.
        While one does, the repair waits: no warning, no Fix, and one line
        every 10 minutes at most. Otherwise Invoke-AutoFix runs it, through
        the same lock check, cooldown, escalation, BLOCKED, DEFERRED and
        30-second warning as any trigger, and a hold there is again not
        counted as an attempt.

        It also ends when a repair runs or cannot be started, when the
        warning is cancelled, and the third time Fix holds it back after the
        warning.
    #>
    if (-not $script:PendingFixReason) { return }
    # On every poll: it lists nothing, and the drop is said on time.
    if (Close-StaleHeldBackRepair -AgeOnly) { return }
    $pending = $script:PendingFixReason
    $now = [datetime]::UtcNow
    if (($now - $script:LastFixTime).TotalMinutes -lt $Cooldown) { return }
    $lastHold = $script:PendingFixHeldAt
    if ($lastHold -eq [datetime]::MinValue) { $lastHold = $script:PendingFixSince }
    if ($now -lt $lastHold.AddMinutes((Get-HeldBackRetryWait -RetriesHeld $script:PendingFixRetriesHeld))) { return }
    if ($now -lt $script:PendingFixNextCheck) { return }
    $script:PendingFixNextCheck = $now.AddMinutes(2)
    # While a Fix run holds the lock, the Fix logs are not read: that run is
    # not over, and only its log says whether it was the repair.
    if (Test-FixLockHeld) {
        Write-FixRunningSkip
        return
    }
    $script:PendingFixLockLogged = $false
    if (Close-StaleHeldBackRepair) { return }

    # Fails closed, as Fix's -Quiet gate does.
    $why = Get-ClaudeCodeSessionHold
    if (-not $why -and (Test-ClaudeCodeActive -Silent)) {
        $why = "a Claude Code session is active, or a transcript was written in the last 35 minutes"
    }
    # What BLOCKED asks, asked first and silently, so a retry shows the
    # warning only when BLOCKED would not stop it. Without it, a user at the
    # PC got a BLOCKED pair of lines every 2 minutes for as long as the
    # repair waited.
    if (-not $why -and (Test-UserActivity -Silent)) {
        $why = "the user is active (Claude in focus, recent input, or VM busy)"
    }
    if ($why) {
        if (($now - $script:PendingFixWaitLogged).TotalMinutes -ge 10) {
            Write-WatchLog "a held-back repair for $pending is still waiting: $why"
            $script:PendingFixWaitLogged = $now
        }
        return
    }
    Invoke-AutoFix -Reason $pending -Retry
}

# -- Prevent duplicate instances -----------------------------------------
# Deliberately carries no version, matching Fix's own lock. Two toolkit
# versions monitoring the same machine at once is exactly what this prevents,
# and a versioned name lets them straight past each other. The old name was
# pinned at v5.0.0.
$mutexName = "Global\ClaudeHealthMonitor"
$script:Mutex = $null
try {
    $script:Mutex = [System.Threading.Mutex]::new($false, $mutexName)
    if (-not $script:Mutex.WaitOne(0)) {
        if (-not $Quiet) {
            Write-Host "  [i] Another health monitor instance is already running." -ForegroundColor DarkGray
        }
        exit 0
    }
} catch [System.Threading.AbandonedMutexException] {
    # Taken: the monitor that held it ended without letting go of it.
    $null = $_
} catch {
    # Could not take the lock. Monitoring without one is better than not
    # monitoring, and Invoke-AutoFix checks Fix's own lock separately before it
    # does anything destructive. Said, naming why: until 6.3.0 it was not.
    Write-CheckError -Check 'single instance' -ErrorRecord $_ `
        -Text "The lock that keeps a second monitor from running ($mutexName) could not be taken ({0}); monitoring without it"
}

# Restore cooldown and fix history before the first poll, so a restart cannot
# be used to escape the backoff.
Import-WatchState

function Sync-ClaudeDataPaths {
    <#
    .SYNOPSIS
        Looks again for the app's data root and log folder, and follows them
        when they have moved.
    .DESCRIPTION
        Discovery used to run once, when this monitor started. The monitor runs
        for days and the app updates under it, so when app 6.18355 moved its
        logs from %APPDATA%\Claude\logs to %LOCALAPPDATA%\Claude\logs the log
        scan kept reading the abandoned folder and never said so.

        Runs every ten minutes, and on every poll where the app has just
        started, which is when an update that moved the folders shows up. A
        call costs Get-ClaudeDataCandidates: a few Test-Path calls and one
        directory listing per candidate.

        The VM log named in the log is picked again when the log folder
        changed, or when there was none and one has appeared since. Two VM
        logs in one folder take turns at being the newest, so following that
        would log a move every ten minutes for nothing. The checks themselves
        do not use that name: they read every VM log in $script:ClaudeLogDir
        on each poll (Get-ClaudeVmLogAge), or the node log by name
        (Test-CoworkSessionActive), so they follow a moved folder and a log
        that appears later on their own.

        A fallback answer is never followed. Discovery falls back to
        %APPDATA%\Claude when it finds nothing, and a candidate folder that
        could not be read for one poll is no reason to move back to a folder
        the app abandoned.
    .PARAMETER Force
        Look now, whatever the schedule says.
    #>
    param([switch]$Force)

    if (-not $Force -and [datetime]::UtcNow -lt $script:NextDataPathSync) { return }
    $script:NextDataPathSync = [datetime]::UtcNow.AddMinutes(10)
    try {
        $p = Resolve-ClaudeDataPaths -PackageFamilyName $script:ClaudeEnv.PackageFamilyName `
                                     -AppDataDir $ClaudeAppData
        # A candidate that could not be read may be the live folder, or may
        # make an old one look the freshest, so nothing is followed while one
        # stands. Until 6.3.1 it looked the same as a folder not there.
        if (@($p.Errors).Count -gt 0) {
            throw "$(@($p.Errors).Count) candidate folder(s) could not be read: $(@($p.Errors) -join '; ')"
        }
        if ($p.DataRoot -and $p.DataRootSource -ne 'fallback' -and
            $p.DataRoot -ne $script:ClaudeEnv.DataRoot) {
            Write-WatchLog "Claude data root moved: $($script:ClaudeEnv.DataRoot) -> $($p.DataRoot) ($($p.DataRootSource))"
            foreach ($k in @('DataRoot', 'DataRootSource', 'VmCachePath', 'BundlePath',
                             'SessionsDir', 'CodeSessionsDir', 'DataRootCandidates')) {
                $script:ClaudeEnv.$k = $p.$k
            }
            $script:StealthMarkerPath = Get-StealthMarkerPath
        }
        if ($p.LogDir -and $p.LogDirSource -ne 'fallback' -and
            $p.LogDir -ne $script:ClaudeLogDir) {
            Write-WatchLog "Claude log folder moved: $($script:ClaudeLogDir) -> $($p.LogDir) ($($p.LogDirSource))"
            $script:ClaudeLogDir = $p.LogDir
            foreach ($k in @('LogDir', 'LogDirSource', 'LogDirCandidates')) {
                $script:ClaudeEnv.$k = $p.$k
            }
            $vm = @(Get-ClaudeVmLogCandidates -LogDir $p.LogDir)
            $newVm = $null
            if ($vm.Count -gt 0) { $newVm = $vm[0].Path }
            $script:ClaudeEnv.VmLogCandidates = $vm
            if ($newVm -ne $script:ClaudeEnv.VmLogFile) {
                Write-WatchLog "VM log moved: $($script:ClaudeEnv.VmLogFile) -> $newVm"
                $script:ClaudeEnv.VmLogFile = $newVm
            }
        }
        $curVm = $script:ClaudeEnv.VmLogFile
        if (-not $curVm -or -not (Test-Path -LiteralPath $curVm)) {
            $vm = @(Get-ClaudeVmLogCandidates -LogDir $script:ClaudeLogDir)
            if ($vm.Count -gt 0) {
                $script:ClaudeEnv.VmLogCandidates = $vm
                Write-WatchLog "VM log found: $($vm[0].Path)"
                $script:ClaudeEnv.VmLogFile = $vm[0].Path
            }
        }
    } catch {
        # At most once an hour for the same error; until 6.3.1 on each look.
        Write-CheckError -Check 'data folders' -ErrorRecord $_ `
            -Text "Looking for the app's folders again failed ({0}); nothing was moved"
    }
}

# -- Main loop -----------------------------------------------------------
Write-WatchLog "================================================================"
Write-WatchLog "Health Monitor v$ToolkitVersion started"
Write-WatchLog "  Poll interval : ${PollInterval}s"
Write-WatchLog "  Cooldown      : ${Cooldown} min"
Write-WatchLog "  Fix script    : $fixScript"
Write-WatchLog "  Log directory : $ClaudeLogDir ($($script:ClaudeEnv.LogDirSource))"
Write-WatchLog "  Data root     : $($script:ClaudeEnv.DataRoot) ($($script:ClaudeEnv.DataRootSource))"
$vmLogNames = @($script:ClaudeEnv.VmLogCandidates | ForEach-Object { $_.Name }) -join ', '
if (-not $vmLogNames) { $vmLogNames = "none yet" }
Write-WatchLog "  VM logs       : $vmLogNames (all read on each poll)"
# Reports what this feature can actually do on THIS machine, in the same spirit
# as the monitor list below: a banner that advertises a capability the run does
# not have is how dead monitors went unnoticed for months.
$stealthNote = "off (-NoStealthRelaunch)"
if (-not $NoStealthRelaunch) {
    if (-not $script:StealthMarkerPath) {
        $stealthNote = "unavailable: no MSIX package family discovered, so there is no marker to watch"
    } elseif (-not $script:ClaudeEnv.Aumid) {
        $stealthNote = "inert: marker path known but no AUMID discovered, so there would be nothing to start"
    } else {
        $stealthNote = "on, marker max age ${MaxStealthMarkerAgeMinutes} min, $($script:StealthMarkerPath)"
    }
}
Write-WatchLog "  Relaunch      : $stealthNote"
Write-WatchLog "  Safety        : user-activity block, 180s grace period, consecutive-check gates, cowork-session awareness"
# Lists only monitors that can actually fire.
#
# The previous line advertised heartbeat(x3), guest-connect and staleness.
# Heartbeat could never fire, because Get-VM does not see an HCS compute
# system. Guest-connect could never fire, because its timestamp patterns did
# not match either real log format. Staleness could never fire, because it
# measured a file last written in March. Naming dead monitors in the startup
# banner is how they stayed unnoticed.
# Heartbeat, guest-connect and NAT were removed in 6.3.0. There is no
# heartbeat to read for an HCS compute system. The guest check read
# cowork-service.log for a message that file has never held, and the file is
# no longer written. The NAT check looked for a WinNAT rule, and the Cowork VM
# does not use one.
Write-WatchLog "  Auto-fix on   : logs, service(x2), events(x2)+HCS, vmcompute-health"
Write-WatchLog "  Warn only     : storage, VM log staleness(x5), clock drift, hcs-state"
Write-WatchLog "  Not monitored : Hyper-V heartbeat (no such thing for an HCS compute system)"
Write-WatchLog "================================================================"
if ($script:SessionReadError) {
    Write-CheckError -Check 'session' -ErrorRecord $script:SessionReadError `
        -Text "The session this monitor runs in could not be read ({0}), so the foreground-window and input checks are skipped, as in Session 0"
}
if ($script:Win32ActivityError) {
    Write-CheckError -Check 'activity checks' -ErrorRecord $script:Win32ActivityError `
        -Text "The foreground-window and input checks could not be loaded ({0}), so they are skipped, as in Session 0"
}
if ($script:LogCleanupError) {
    Write-CheckError -Check 'old logs' -ErrorRecord $script:LogCleanupError `
        -Text "The watch logs older than 30 days in $WatchLogDir could not all be deleted ({0}); they are tried again at the next rollover"
}
# Each read the discovery at start could not make, once (ReadErrors). Until
# 6.3.1 one that failed looked the same as one that found nothing.
foreach ($readError in @($script:ClaudeEnv.ReadErrors)) {
    Write-WatchLog "READ ERROR at start: $readError"
}
# Its own check: a poll that cannot read it either says so as well.
if ($script:SeedReadError) {
    Write-CheckError -Check 'Claude processes at start' -ErrorRecord $script:SeedReadError `
        -Text "Whether Claude is running could not be read at start ({0}); the first poll that reads it starts from there"
}

try {
    while ($true) {
        try {
            Update-LogFile

            # The app, not any "claude". With every process named claude
            # counted, a live Claude Code CLI process (a project thread, a
            # terminal session) kept this true after the app had gone, so the
            # exit below was never seen, a stealth update was never relaunched,
            # and the checks under it ran against an app that was not there.
            #
            # A list that could not be read ends the poll, said at most once
            # an hour: every check below turns on whether the app runs, and
            # the relaunch below starts it. Until 6.3.1 such a list was
            # empty, "Claude is not running".
            try {
                $claudeRunning = @(Get-ClaudeDesktopProcess -Environment $script:ClaudeEnv).Count -gt 0
            } catch {
                Write-CheckError -Check 'Claude processes' -ErrorRecord $_ `
                    -Text "Whether Claude is running could not be read ({0}), so this poll checked nothing and did nothing"
                continue
            }

            # An app that has just started may be a new version that keeps its
            # files somewhere else, so look at once rather than at the next
            # ten-minute mark.
            Sync-ClaudeDataPaths -Force:($claudeRunning -and -not $script:ClaudeWasRunning)

            # Record Claude disappearing. Acted on only for a stealth update.
            #
            # Every check below sits inside "if ($claudeRunning)", so when
            # Claude is gone this monitor previously did nothing at all: no
            # detection, no note, nothing. A user report describes exactly that
            # shape, an app that died and stayed dead for over an hour with
            # zero log activity to point at afterwards.
            #
            # A crash and a deliberate quit still look the same from here, and
            # relaunching an app the user closed on purpose would be worse than
            # useless, so neither is acted on. A stealth update is the one case
            # that does not look the same: the app writes a marker file on its
            # way out, and nothing else does. Where that marker is present the
            # sentence below would be false, so it is not printed.
            #
            # Whether the service is still up is recorded alongside it: Close
            # mode stops the service, so a live service with no Claude leans
            # towards an unexpected exit rather than a clean shutdown.
            #
            # Not read at start: seeded by the first poll that reads it, with
            # no transition to report.
            if ($null -eq $script:ClaudeWasRunning) {
                $script:ClaudeWasRunning = $claudeRunning
            } elseif ($claudeRunning -ne $script:ClaudeWasRunning) {
                if (-not $claudeRunning) {
                    # Not registered and could not be read are told apart
                    # (Get-ClaudeServiceOrNull). Until 6.3.1 both said
                    # "absent".
                    try {
                        $svcNow = Get-ClaudeServiceOrNull -Name $ServiceName
                        $svcState = if ($svcNow) { "$($svcNow.Status)" } else { "absent" }
                    } catch {
                        $svcState = "not known, it could not be read: $(Get-CheckErrorText -ErrorRecord $_)"
                    }
                    if ($NoStealthRelaunch) {
                        Write-WatchLog "Claude is no longer running (service: $svcState). Not acting: stealth relaunch is off (-NoStealthRelaunch)."
                    } elseif (Get-StealthRelaunchMarker) {
                        Write-WatchLog "Claude is no longer running (service: $svcState). A stealth-update marker is present, so this was an update quit, not a user quit."
                    } else {
                        Write-WatchLog "Claude is no longer running (service: $svcState). Not acting: a crash and a deliberate quit look the same from here."
                    }
                } else {
                    Write-WatchLog "Claude is running again"
                }
                $script:ClaudeWasRunning = $claudeRunning
            }

            # Evaluated every poll, not only on the transition above. A monitor
            # that restarts, or that starts at logon after the update already
            # happened, never sees a transition: ClaudeWasRunning is seeded from
            # the current state, so an app that is already down looks unchanged.
            # Invoke-StealthRelaunch is idempotent per marker, so re-asking
            # costs one Test-Path on the polls where there is nothing to do.
            if (-not $claudeRunning) {
                $null = Invoke-StealthRelaunch
            }

            if ($claudeRunning) {
                # ---- Critical checks (trigger auto-fix) ----

                # Check 1: a VM failure logged at [error] in the app's VM or
                # main log. No consecutive gate, which is why the match is
                # limited to [error] entries: see Find-LogErrorPattern.
                if (-not (Test-StartupGracePeriod)) {
                    $logError = Test-LogsForErrors
                    if ($logError) {
                        # One-shot: Test-LogsForErrors moves its baselines
                        # past what it read, so the same line never fires
                        # again, and a repair held back for it is kept.
                        Invoke-AutoFix -Reason "Log error: $logError" -OneShot
                        continue
                    }
                }

                # Check 2: Service died while Claude is running (2 consecutive checks)
                if (-not (Test-ServiceHealth)) {
                    # Re-checked: the service is asked about again on every
                    # poll, so this fires again while it stays down, and a
                    # repair held back for it is not kept.
                    Invoke-AutoFix -Reason "CoworkVMService stopped while Claude is running (confirmed)"
                    continue
                }

                # Check 3: Event Log errors (2 consecutive checks, grace period, tightened filters)
                $eventError = Test-EventLogErrors
                if ($eventError) {
                    # One-shot: an event counts only inside the 60-second
                    # look-back, and the count starts again from 0 once it
                    # fires, so the same event does not fire again. A repair
                    # held back for it is kept.
                    Invoke-AutoFix -Reason "Event Log: $eventError" -OneShot
                    continue
                }

                # Check 4b: workspace storage readable and not out of room.
                # Warning only: a restart does not create disk space. Each
                # different text is said at most once an hour, with no cap
                # on how many (Test-StorageWarningDue), so a new warning, or
                # one whose text changes to one not said in that hour, is
                # said at once. Until 6.3.1 it was written on every poll
                # while it lasted, about 120 identical lines an hour in logs
                # kept 30 days.
                $storageIssue = Test-WorkspaceStorage
                if ($storageIssue -and (Test-StorageWarningDue -Message $storageIssue -Now ([datetime]::UtcNow))) {
                    Write-WatchLog "STORAGE WARNING: $storageIssue; the same warning is logged at most once an hour"
                }

                # Check 5 was the Hyper-V heartbeat, removed in 6.3.0. It read
                # the heartbeat through Get-VM, which never lists an HCS
                # compute system such as the Cowork VM, so it could not fire,
                # and from 6.0.0 it was a function that returned nothing.
                # Check 6 and the hcs-state check watch the VM instead.

                # Check 6: VM log staleness. WARNING ONLY. It does not auto-fix.
                #
                # This was an auto-fix trigger, and until v6.0.0 it could never
                # fire, because it measured a log last written in March and was
                # gated behind a flag that never got set. Fixing both made it
                # live for the first time, and a foreground run on a healthy
                # machine tripped it three seconds after the grace period
                # expired: cowork_vm_node.log had simply been quiet for 474
                # seconds because nobody was using Cowork.
                #
                # The guard that was meant to prevent that, Test-CoworkSessionActive,
                # reads the freshness of the SAME log. So "log is stale" implies
                # "no active session" implies "check staleness": the guard can
                # never guard. That circularity is not fixable with the signals
                # available on the host, because an idle VM and a hung VM look
                # identical from outside.
                #
                # Warning it is, until there is a signal that can tell those two
                # apart. Authorising a destructive repair on evidence this weak
                # would produce a repair loop on any machine left idle with
                # Claude open.
                $staleIssue = Test-VmLogStaleness
                if ($staleIssue) {
                    Write-WatchLog "STALENESS WARNING: $staleIssue"
                }

                # Check 7: vmcompute handle leak (2 consecutive checks, 60s interval)
                $vmcomputeIssue = Test-VmcomputeHealth
                if ($vmcomputeIssue) {
                    # Re-checked: the handle count is read again every 60
                    # seconds, so this fires again while the leak lasts, and
                    # a repair held back for it is not kept.
                    Invoke-AutoFix -Reason $vmcomputeIssue
                    continue
                }

                # Check 8: HCS state health (stale VMs, shutdown failure frequency)
                $hcsStateIssue = Test-HcsStateHealth
                if ($hcsStateIssue) {
                    # Rate-limit HCS STATE logging to once per 10 minutes (was every 30s)
                    $hcsLogElapsed = ([datetime]::UtcNow - $script:LastHcsStateLogTime).TotalMinutes
                    if ($hcsLogElapsed -ge 10) {
                        Write-WatchLog "HCS STATE: $hcsStateIssue"
                        $script:LastHcsStateLogTime = [datetime]::UtcNow
                        # Informational plus a pre-emptive cleanup of the
                        # Cowork compute systems left behind; no auto-fix.
                        Stop-StaleCoworkVm
                    }
                } else {
                    # Reset timer when healthy so next issue logs immediately
                    if ($script:LastHcsStateLogTime -ne [datetime]::MinValue) {
                        $script:LastHcsStateLogTime = [datetime]::MinValue
                    }
                }

                # ---- Maintenance (non-fix actions) ----
                Set-VmWorkerPriority

                $timeIssue = Test-TimeSyncHealth
                if ($timeIssue) {
                    Write-WatchLog "TIME WARNING: $timeIssue"
                }

                # A repair Fix held back, tried again once the hold has
                # cleared. Only on a poll where no check fired: one that fires
                # has gone to Invoke-AutoFix above and ended the poll.
                Invoke-HeldBackRetry
            }
        } catch {
            # At most once an hour for the same error: one that breaks a poll
            # most often breaks every poll. Until 6.3.0 it was logged on each.
            Write-CheckError -Check 'poll' -ErrorRecord $_ -Text 'MONITOR ERROR: {0}'
        } finally {
            # The poll sleep lives in a finally, and that is the whole point.
            #
            # Five of the checks above end in "continue", which targets the
            # while loop and therefore used to jump straight past a Start-Sleep
            # that sat after the try. During a real fault, when those branches
            # are exactly the ones firing, the loop span without any delay at
            # all: a 1.8 second directory enumeration and a Start-Job spawn,
            # over and over, for as long as the fault lasted.
            #
            # PowerShell runs finally when control leaves the try by continue,
            # so the sleep now happens on every path.
            Start-Sleep -Seconds $PollInterval
        }
    }
} finally {
    if ($script:Mutex) {
        try {
            $script:Mutex.ReleaseMutex()
            $script:Mutex.Dispose()
        } catch {
            # Not held, as when it could not be taken: nothing to release,
            # and the monitor is ending.
            $null = $_
        }
    }
    Write-WatchLog "Health monitor stopped"
}
