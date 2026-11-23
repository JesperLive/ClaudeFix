<#
.SYNOPSIS
    Claude Desktop / Cowork, Preventive Configuration

.DESCRIPTION
    One-shot script that sets Windows up for Claude Desktop's Cowork VM: full
    performance power settings, nothing under the VM that sleeps, parks or
    saves power, and the tools that keep it running.

    Run once. Changes persist across reboots. Every setting it changes is
    recorded before it is changed, so -Undo can put back exactly what was
    there.

    What it does:
    - Makes an Ultimate Performance plan active (High performance where
      Windows has no Ultimate Performance), and records the plan you had
    - Sets that plan to full performance on AC power: no sleep, no hibernate
      timer, no disk spin-down, no SATA or NVMe link power saving, no USB
      selective suspend or USB 3 link power saving, no PCI Express link
      power saving, processors at 100% with no core parking, aggressive boost
      and no throttle states, and wireless adapters at maximum performance
    - Turns off hibernate and Fast Startup
    - Turns off Modern Standby (S0 low power idle) where the machine has it
    - Turns off power saving on every physical network adapter: Windows'
      "allow the computer to turn off this device", the driver framework's
      power-down while the PC is running, Energy Efficient Ethernet, Green
      Ethernet and the driver's other power saving options, hidden ones
      included, selective suspend and packet coalescing. It also clears
      "allow the computer to turn off this device" on the adapters that are
      not physical: hidden ones, such as Wi-Fi Direct, and virtual ones,
      such as a VPN's adapter. No adapter is restarted; the changes take
      effect at the next Windows restart
    - Turns off Windows power throttling
    - Raises the Cowork VM's worker process to AboveNormal priority
    - Makes the Hyper-V compute service (vmcompute) restart itself on failure
    - Removes stale Cowork VMs, only while Claude is closed
    - Checks the workspace storage location, the clock, antivirus exclusions
      (adding the Defender ones) and WSL2
    - Installs the health monitor, the Fix Claude Desktop shortcuts and the
      Claude (Admin) and Claude (no GPU) launchers
    - Removes what older versions set up that is no longer needed

    What it does NOT do:
    - Touch your Claude config or conversations
    - Change battery (DC) power settings or your screen timeout
    - Restart Windows or any network adapter. The summary at the end lists
      what needs a restart to take effect

.PARAMETER Undo
    Puts back every setting recorded by a run of 6.3.0 or later, removes what
    older versions changed, and removes the monitor, tasks, launchers and
    shortcuts. The record is %LOCALAPPDATA%\ClaudeFix\prevent-undo.json, and
    only the computer that wrote it uses it.

.NOTES
    Exit code: 0 when every step ran, 1 when a step failed or the run stopped.
    -Undo also exits 1 while a setting is still in the undo record.
    -WhatIf is refused: this script cannot preview its changes.

    Version : 6.3.3
    Author  : Jesper Driessen
    Licence : MIT
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$Undo
)

# -WhatIf is refused, before anything else runs, elevation included. Until
# 6.3.0 it was accepted and carried across the elevation boundary, but the
# only step that honoured it was the stale-VM removal: every other step
# changed the machine for real, so a "dry run" was a real run. A preview that
# covers one step in twenty-seven is worse than none. SupportsShouldProcess
# stays declared so -WhatIf reaches this refusal rather than failing as an
# unknown parameter.
if ($WhatIfPreference) {
    Write-Host "  [!] -WhatIf is not supported: this script cannot preview its changes." -ForegroundColor Red
    Write-Host "      Nothing was changed. For the list of what it does, run:" -ForegroundColor Yellow
    Write-Host "      Get-Help .\Prevent-ClaudeIssues.ps1 -Full" -ForegroundColor Yellow
    exit 1
}

# -- Admin elevation -------------------------------------------------
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $scriptFile = $PSCommandPath
    if (-not $scriptFile) { $scriptFile = $MyInvocation.MyCommand.Definition }

    $elevateArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptFile`""
    if ($Undo) { $elevateArgs += " -Undo" }

    try {
        # -Wait -PassThru so the child's exit code reaches the launcher. Without
        # it the outer shell returned 0 the moment UAC was accepted, so a failed
        # elevated run looked like success to Prevent-ClaudeIssues.bat.
        $elevated = Start-Process PowerShell -ArgumentList $elevateArgs -Verb RunAs -Wait -PassThru
        if ($elevated) { exit $elevated.ExitCode }
        exit 0
    } catch {
        Write-Host "  [!] UAC elevation was declined or failed." -ForegroundColor Red
        Write-Host "      This script requires Administrator privileges." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Press any key to close..." -ForegroundColor DarkGray
        [void][System.Console]::ReadKey($true)
    }
    # Explicit 1, not a bare exit. Declining UAC means nothing was configured,
    # and the launcher treats 1 as "the script explained itself and paused",
    # which is exactly what just happened above.
    exit 1
}

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

# -- Runtime discovery -----------------------------------------------
# -SkipCacheInventory: nothing here needs the VHDX inventory, and the walk is
# the expensive half of discovery.
$script:ClaudeEnv = Get-ClaudeEnvironment -SkipCacheInventory

# -- Constants -------------------------------------------------------
$ToolkitVersion   = "6.3.3"
$TaskName         = "ClaudeCoworkWatchdog"
# The logon-time fix task, retired in 6.3.0. Setup removes it where an older
# version installed it, and so does -Undo.
$BootTaskName     = "ClaudeCoworkBootFix"
$TaskPath         = "\Claude\"
$ClaudeAppData    = $script:ClaudeEnv.AppDataDir
# Where the app keeps its data and VM image, which is what the storage check
# and the Defender exclusion are about. Until app 6.18355 that was
# $ClaudeAppData; since, it is %LOCALAPPDATA%\Claude-Data\claude_<id>, and
# $ClaudeAppData holds only this toolkit's files. Discovery falls back to
# $ClaudeAppData when it finds no data root.
$ClaudeDataRoot   = $script:ClaudeEnv.DataRoot
if (-not $ClaudeDataRoot) { $ClaudeDataRoot = $ClaudeAppData }

# Both discovered. $ServiceName was removed in the first pass as an unused
# variable, which it was: everything referred to the literal "CoworkVMService"
# instead. It is back because those references now use it.
$ServiceName      = $script:ClaudeEnv.ServiceName
$ServiceExe       = $script:ClaudeEnv.ServiceProcessName

if (-not $ClaudeAppData) {
    Write-Host "  [!] Could not locate the per-user Claude folder under %APPDATA%." -ForegroundColor Red
    exit 1
}
# The undo record. Every setting this script changes is written here, with
# the value it had, BEFORE it is changed, so -Undo puts back what was there.
# Until 6.3.0 -Undo set fixed values instead (hibernate on, sleep after 30
# minutes, Fast Startup on, the adapters' power options deleted), whatever
# the machine had before.
#
# The settings it records belong to this computer, so the record is kept in
# %LOCALAPPDATA%, which does not roam, and it names the computer that wrote
# it. In the roaming %APPDATA% it followed a roaming profile to the next PC,
# where the first record of each setting won: that PC's own values were
# never recorded, and its -Undo wrote the first PC's. Not %ProgramData%: any
# user can create files there, and -Undo writes what this record says into
# HKLM and the power plans.
$LocalAppData     = [Environment]::GetFolderPath('LocalApplicationData')
if (-not $LocalAppData) {
    Write-Host "  [!] Could not locate the per-user %LOCALAPPDATA% folder." -ForegroundColor Red
    exit 1
}
$UndoDir          = Join-Path $LocalAppData "ClaudeFix"
$UndoFile         = Join-Path $UndoDir "prevent-undo.json"
# Written by versions before 6.3.0, and read now only to carry what they
# hold over to the undo record and to -Undo. Never written again.
$BackupFile       = Join-Path $ClaudeAppData "power-plan-backup.txt"
$CsBackupFile     = Join-Path $ClaudeAppData "connected-standby-backup.txt"
# Claude's icon, copied out of Claude.exe for the shortcuts. Until 6.3.0 the
# shortcuts pointed at Claude.exe itself, inside a folder whose name carries
# the app version, so the first Claude update after setup blanked them.
$ClaudeIconFile   = Join-Path $ClaudeAppData "Claude.ico"

$GuidPattern          = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'
# The hidden Ultimate Performance template, which "powercfg /duplicatescheme"
# copies into a plan of its own, and the two stock plans.
$UltimateTemplateGuid = "e9a42b02-d5df-448d-aa00-03f14749eb61"
$HighPerfGuid         = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
$BalancedGuid         = "381b4222-f694-41f0-9685-ff5bb260df2e"

# Every AC power setting step 2 puts on the active plan: the power plan's
# own word on what may sleep, spin down, park or save power. Sub and Setting
# are powercfg's GUIDs, Value is the AC index and Shown is how the log names
# it. A setting this machine does not have (no NVMe drive, no wireless
# adapter, a processor without a second core class) is listed as not present
# and left alone. Battery (DC) values are not touched.
$PowerOptions = @(
    @{ Label = "Sleep after";                      Sub = "238c9fa8-0aad-41ed-83f4-97be242c8f20"; Setting = "29f6c1db-86da-48c5-9fdb-f2b67b1f44da"; Value = 0;   Shown = "Never" }
    @{ Label = "Hibernate after";                  Sub = "238c9fa8-0aad-41ed-83f4-97be242c8f20"; Setting = "9d7815a6-7ee4-497e-8888-515a05f02364"; Value = 0;   Shown = "Never" }
    @{ Label = "Turn off hard disk after";         Sub = "0012ee47-9041-4b5d-9b77-535fba8b1442"; Setting = "6738e2c4-e8a5-4a42-b16a-e040e769756e"; Value = 0;   Shown = "Never" }
    @{ Label = "SATA link power management";       Sub = "0012ee47-9041-4b5d-9b77-535fba8b1442"; Setting = "0b2d69d7-a2a1-449c-9680-f91c70521c60"; Value = 0;   Shown = "Active" }
    @{ Label = "SATA adaptive link power";         Sub = "0012ee47-9041-4b5d-9b77-535fba8b1442"; Setting = "dab60367-53fe-4fbc-825e-521d069d2456"; Value = 0;   Shown = "Off" }
    @{ Label = "NVMe idle timeout (primary)";      Sub = "0012ee47-9041-4b5d-9b77-535fba8b1442"; Setting = "d639518a-e56d-4345-8af2-b9f32fb26109"; Value = 0;   Shown = "Off" }
    @{ Label = "NVMe idle timeout (secondary)";    Sub = "0012ee47-9041-4b5d-9b77-535fba8b1442"; Setting = "d3d55efd-c1ff-424e-9dc3-441be7833010"; Value = 0;   Shown = "Off" }
    @{ Label = "NVMe latency tolerance (primary)"; Sub = "0012ee47-9041-4b5d-9b77-535fba8b1442"; Setting = "fc95af4d-40e7-4b6d-835a-56d131dbc80e"; Value = 0;   Shown = "0 ms" }
    @{ Label = "NVMe latency tolerance (secondary)"; Sub = "0012ee47-9041-4b5d-9b77-535fba8b1442"; Setting = "dbc9e238-6de9-49e3-92cd-8c2b4946b472"; Value = 0; Shown = "0 ms" }
    @{ Label = "USB selective suspend";            Sub = "2a737441-1930-4402-8d77-b2bebba308a3"; Setting = "48e6b7a6-50f5-4782-a5d4-53bb8f07e226"; Value = 0;   Shown = "Off" }
    @{ Label = "USB 3 link power management";      Sub = "2a737441-1930-4402-8d77-b2bebba308a3"; Setting = "d4e98f31-5ffe-4ce1-be31-1b38b384c009"; Value = 0;   Shown = "Off" }
    @{ Label = "PCI Express link state";           Sub = "501a4d13-42af-4429-9fd1-a8218c268e20"; Setting = "ee12f906-d277-404b-b6da-e5fa1a576df5"; Value = 0;   Shown = "Off" }
    @{ Label = "Minimum processor state";          Sub = "54533251-82be-4824-96c1-47b60b740d00"; Setting = "893dee8e-2bef-41e0-89c6-b55d0929964c"; Value = 100; Shown = "100%" }
    @{ Label = "Maximum processor state";          Sub = "54533251-82be-4824-96c1-47b60b740d00"; Setting = "bc5038f7-23e0-4960-96da-33abaf5935ec"; Value = 100; Shown = "100%" }
    @{ Label = "Processor boost mode";             Sub = "54533251-82be-4824-96c1-47b60b740d00"; Setting = "be337238-0d82-4146-a960-4f3749d470c7"; Value = 2;   Shown = "Aggressive" }
    @{ Label = "Core parking minimum cores";       Sub = "54533251-82be-4824-96c1-47b60b740d00"; Setting = "0cc5b647-c1df-4637-891a-dec35c318583"; Value = 100; Shown = "100% (no parking)" }
    @{ Label = "Core parking minimum cores, class 1"; Sub = "54533251-82be-4824-96c1-47b60b740d00"; Setting = "0cc5b647-c1df-4637-891a-dec35c318584"; Value = 100; Shown = "100% (no parking)" }
    @{ Label = "Energy performance preference";    Sub = "54533251-82be-4824-96c1-47b60b740d00"; Setting = "36687f9e-e3a5-4dbf-b1dc-15eb381c6863"; Value = 0;   Shown = "0 (performance)" }
    @{ Label = "Energy performance preference, class 1"; Sub = "54533251-82be-4824-96c1-47b60b740d00"; Setting = "36687f9e-e3a5-4dbf-b1dc-15eb381c6864"; Value = 0; Shown = "0 (performance)" }
    @{ Label = "Processor throttle states";        Sub = "54533251-82be-4824-96c1-47b60b740d00"; Setting = "3b04d4fd-1cc7-4f23-ab1c-d1337819c4bb"; Value = 0;   Shown = "Off" }
    @{ Label = "Wireless adapter power saving";    Sub = "19cbb8fa-5279-450e-9fac-8a3d5fedd0c1"; Setting = "12bbebe6-58d6-4636-95bb-3217ef867c1a"; Value = 0;   Shown = "Maximum performance" }
)

# The network drivers' own power saving options, by registry keyword, each
# with the value that turns it off. The names on the adapter's Advanced tab
# vary by vendor and language, the keywords do not. Off is not 0 for every
# keyword: for Intel's MIMO power save 0 is Auto SMPS and 3 is No SMPS, for
# Broadcom's EEE control 0 is the most power saving and 2 is off, and for
# *IdleRestriction 1 is the setting that keeps the adapter from idling. A
# keyword is set only where the driver lists its off value as a valid one.
#   *EEE                    Energy-Efficient Ethernet (the standard keyword)
#   EEELinkAdvertisement    Energy Efficient Ethernet (Intel)
#   AdvancedEEE             Advanced EEE (Realtek)
#   EnableGreenEthernet     Green Ethernet (Realtek)
#   GigaLite                Gigabit Lite (Realtek)
#   PowerSavingMode         Power Saving Mode (Realtek)
#   AutoDisableGigabit      Auto Disable Gigabit (Realtek)
#   ULPMode                 Ultra Low Power Mode (Intel)
#   SipsEnabled             System Idle Power Saver (Intel)
#   MIMOPowerSaveMode       MIMO Power Save Mode (Intel Wi-Fi): 3, No SMPS
#   bLeisurePs              Wi-Fi power saving (Realtek Wi-Fi)
#   LPMEnable               Link power management
#   *IdleRestriction        Idle restriction (the standard keyword): 1
#   EeeCtrlMode             EEE control (Broadcom): 2
#   GreenEthernet           Green Ethernet
#   APSmode                 Auto power saving
#   AutoLinkDownPcieMacOff  Power down the PCIe link when the cable is out
#   EnableLDPS              Link Down Power Saving
#   SleepPhyOnLinkDown      Sleep the PHY when the link is down
#   EEEMode                 Energy Efficient Ethernet
#   LowPowerIdleMode        Low power idle
$NicPowerOff = [ordered]@{
    '*EEE'                   = '0'
    'EEELinkAdvertisement'   = '0'
    'AdvancedEEE'            = '0'
    'EnableGreenEthernet'    = '0'
    'GigaLite'               = '0'
    'PowerSavingMode'        = '0'
    'AutoDisableGigabit'     = '0'
    'ULPMode'                = '0'
    'SipsEnabled'            = '0'
    'MIMOPowerSaveMode'      = '3'
    'bLeisurePs'             = '0'
    'LPMEnable'              = '0'
    '*IdleRestriction'       = '1'
    'EeeCtrlMode'            = '2'
    'GreenEthernet'          = '0'
    'APSmode'                = '0'
    'AutoLinkDownPcieMacOff' = '0'
    'EnableLDPS'             = '0'
    'SleepPhyOnLinkDown'     = '0'
    'EEEMode'                = '0'
    'LowPowerIdleMode'       = '0'
}

# Two power saving options some drivers keep in the registry alone: the
# Advanced tab does not show them and the driver lists no valid values for
# them. Each with the value that turns it off. Both read 1 on the Realtek
# 2.5GbE and the TP-Link (a Realtek chip) Wi-Fi adapter they were measured
# on:
#   PowerDownPll   Realtek Ethernet: the driver may power down the PLL, the
#                  clock the link runs on
#   bFwCtrlLPS     Realtek Wi-Fi: the firmware may put the radio into Leisure
#                  Power Save, dozing between the access point's beacons, on
#                  its own
# Both only save power. Off keeps the clock or the radio awake, which costs
# a little power and changes nothing about the traffic the adapter carries.
# With no valid values listed, the driver does not say what any other value
# means, so an option is changed only when it reads exactly 1, the on value
# of the on and off pair these two are known by. Any other value is left as
# it is.
$NicHiddenPowerOff = [ordered]@{
    'PowerDownPll' = '0'
    'bFwCtrlLPS'   = '0'
}

# The driver option, by its standardized keyword, that each power saving
# setting of Set-NetAdapterPowerManagement is kept in. A driver that lists
# none for a setting gives the cmdlet nowhere to write it: on the PC this was
# measured on, the Realtek 2.5GbE driver, rt25cx21, lists no
# *SelectiveSuspend, and Set-NetAdapterPowerManagement -SelectiveSuspend
# Disabled returned without an error, wrote nothing to its driver key and
# read back Enabled.
$NicPmKeyword = [ordered]@{
    'SelectiveSuspend'        = '*SelectiveSuspend'
    'DeviceSleepOnDisconnect' = '*DeviceSleepOnDisconnect'
    'D0PacketCoalescing'      = '*PacketCoalescing'
}

# Reads Claude.exe's icon group as a data file and writes it out as an .ico.
# Compiled on first use only; see Save-ClaudeIcon.
$script:IconExtractorSource = @'
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace ClaudeFix {
    public static class IconExtractor {
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern IntPtr LoadLibraryEx(string path, IntPtr file, uint flags);
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool FreeLibrary(IntPtr module);
        delegate bool EnumResNameProc(IntPtr module, IntPtr type, IntPtr name, IntPtr param);
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern bool EnumResourceNames(IntPtr module, IntPtr type, EnumResNameProc proc, IntPtr param);
        [DllImport("kernel32.dll", SetLastError = true, EntryPoint = "FindResourceW")]
        static extern IntPtr FindResourceById(IntPtr module, IntPtr name, IntPtr type);
        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "FindResourceW")]
        static extern IntPtr FindResourceByName(IntPtr module, string name, IntPtr type);
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern uint SizeofResource(IntPtr module, IntPtr res);
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr LoadResource(IntPtr module, IntPtr res);
        [DllImport("kernel32.dll")]
        static extern IntPtr LockResource(IntPtr data);

        const uint LOAD_LIBRARY_AS_DATAFILE = 0x2;
        const uint LOAD_LIBRARY_AS_IMAGE_RESOURCE = 0x20;
        static readonly IntPtr RT_ICON = (IntPtr)3;
        static readonly IntPtr RT_GROUP_ICON = (IntPtr)14;

        static byte[] Read(IntPtr module, IntPtr res) {
            if (res == IntPtr.Zero) return null;
            uint size = SizeofResource(module, res);
            IntPtr handle = LoadResource(module, res);
            if (handle == IntPtr.Zero || size == 0) return null;
            IntPtr data = LockResource(handle);
            if (data == IntPtr.Zero) return null;
            byte[] bytes = new byte[size];
            Marshal.Copy(data, bytes, 0, (int)size);
            return bytes;
        }

        // GRPICONDIR: 6-byte header, then 14-byte entries whose last two
        // bytes are the RT_ICON id. An .ico file has the same header and
        // 16-byte entries that end in the image's size and file offset.
        public static bool ExtractGroup(string exePath, string icoPath) {
            IntPtr module = LoadLibraryEx(exePath, IntPtr.Zero, LOAD_LIBRARY_AS_DATAFILE | LOAD_LIBRARY_AS_IMAGE_RESOURCE);
            if (module == IntPtr.Zero) return false;
            try {
                object first = null;
                EnumResourceNames(module, RT_GROUP_ICON, delegate(IntPtr m, IntPtr t, IntPtr name, IntPtr param) {
                    long v = name.ToInt64();
                    if ((v >> 16) == 0) first = (int)v; else first = Marshal.PtrToStringUni(name);
                    return false;
                }, IntPtr.Zero);
                if (first == null) return false;
                IntPtr groupRes = (first is int)
                    ? FindResourceById(module, (IntPtr)(int)first, RT_GROUP_ICON)
                    : FindResourceByName(module, (string)first, RT_GROUP_ICON);
                byte[] group = Read(module, groupRes);
                if (group == null || group.Length < 6) return false;
                int count = BitConverter.ToUInt16(group, 4);
                if (count == 0 || group.Length < 6 + 14 * count) return false;
                byte[][] images = new byte[count][];
                for (int i = 0; i < count; i++) {
                    int id = BitConverter.ToUInt16(group, 6 + 14 * i + 12);
                    images[i] = Read(module, FindResourceById(module, (IntPtr)id, RT_ICON));
                    if (images[i] == null) return false;
                }
                using (MemoryStream ms = new MemoryStream())
                using (BinaryWriter w = new BinaryWriter(ms)) {
                    w.Write((ushort)0);
                    w.Write((ushort)1);
                    w.Write((ushort)count);
                    int offset = 6 + 16 * count;
                    for (int i = 0; i < count; i++) {
                        w.Write(group, 6 + 14 * i, 8);
                        w.Write((uint)images[i].Length);
                        w.Write((uint)offset);
                        offset += images[i].Length;
                    }
                    for (int i = 0; i < count; i++) w.Write(images[i]);
                    w.Flush();
                    File.WriteAllBytes(icoPath, ms.ToArray());
                }
                return true;
            } finally {
                FreeLibrary(module);
            }
        }
    }
}
'@

# -- Helpers ---------------------------------------------------------
function Log {
    param([string]$Message, [string]$Colour = "White", [switch]$Indent)
    $pfx = ""
    if ($Indent) { $pfx = "      " }
    Write-Host "$pfx$Message" -ForegroundColor $Colour
}

function Initialize-ClaudeAppData {
    # %APPDATA%\Claude was created by Claude Desktop on first launch until app
    # 6.18355, which keeps its data under %LOCALAPPDATA%\Claude-Data and need
    # not create it at all. This script is meant to be runnable BEFORE Claude
    # Desktop is installed, and it also
    # self-elevates -- which can land it in a different account's profile that
    # has never launched Claude. Either way the folder may not exist yet, and
    # every file this script writes lives inside it, except the undo record,
    # which is in %LOCALAPPDATA%\ClaudeFix.
    # Returns $true when the folder is there or was made, and throws, naming
    # the folder and why, when it could not be read or made. Until 6.3.1
    # Test-Path was asked, which answers $false for a folder that cannot be
    # read as for one that is not there (measured 2026-11-23), and a folder
    # that could not be read was said to be one that could not be created.
    $item = $null
    try {
        $item = Get-ClaudeItemOrNull -LiteralPath $script:ClaudeAppData
    } catch {
        throw "$($script:ClaudeAppData) could not be read ($(Get-ClaudeErrorText $_))"
    }
    if ($null -ne $item) { return $true }
    try {
        New-Item -Path $script:ClaudeAppData -ItemType Directory -Force -ErrorAction Stop | Out-Null
    } catch {
        throw "$($script:ClaudeAppData) could not be created ($(Get-ClaudeErrorText $_))"
    }
    return $true
}

function Test-DefenderExclusionCovers {
    <#
    .SYNOPSIS
        True when one of Defender's ExclusionPath entries covers Path: the
        same path, a folder above it, or a wildcard entry that matches it.
    .DESCRIPTION
        Until 6.2.0 this was a bare prefix test, which ignores folder
        boundaries. The data root %LOCALAPPDATA%\Claude-Data\claude_<id>
        sits beside the app's own %LOCALAPPDATA%\Claude, so an exclusion for
        that folder read as covering the data root and the data root was
        never excluded. Measured 2026-11-21: a machine in exactly that state.
        An empty entry used to cover every path.

        A Defender entry may use * and ? and environment variables, so those
        are honoured; [ and ] are taken literally, as they are in a path.
        An entry starting "N/A" is the placeholder Get-MpPreference returns
        to a caller that may not read the list, and covers nothing.
    #>
    param([string]$Path, [string[]]$Exclusions)
    if (-not $Path) { return $false }
    $p = $Path.TrimEnd('\')
    foreach ($ex in @($Exclusions)) {
        if (-not $ex -or $ex -like 'N/A*') { continue }
        $e = [Environment]::ExpandEnvironmentVariables($ex.Trim()).TrimEnd('\')
        if (-not $e) { continue }
        $pattern = $e -replace '[\[\]`]', '`$0'
        if (($p -like $pattern) -or ($p -like "$pattern\*")) { return $true }
    }
    return $false
}

function Test-DefenderListHidden {
    <#
    .SYNOPSIS
        True when an exclusion list from Get-MpPreference holds nothing but
        its "N/A" placeholder. Defender returns that in place of the list to
        a caller it hides the list from, which a policy can make every
        administrator. Such a list says nothing about which exclusions
        exist: read as empty, every exclusion looked missing, and step 13
        added and recorded ones that were already there, which -Undo then
        removed.
    #>
    param([string[]]$List)
    $seen = $false
    foreach ($e in @($List)) {
        if (-not $e) { continue }
        if ($e -notlike 'N/A*') { return $false }
        $seen = $true
    }
    return $seen
}

function ConvertTo-StringArray {
    <#
    .SYNOPSIS
        A value as a string array, empty for $null. Not a pipeline: piping
        $null into ForEach-Object runs the block once, so that gives one
        empty string where there is nothing.
    #>
    param($Value)
    $out = @()
    if ($null -ne $Value) {
        foreach ($v in @($Value)) {
            if ($null -ne $v) { $out += [string]$v }
        }
    }
    return ,([string[]]$out)
}

# -- Step results ----------------------------------------------------
# Each step reports what it actually did, and the closing summary and the
# exit code are built from those reports. Until 6.3.0 the summary was a
# fixed list that called every setting configured, printed after steps that
# had failed, said "not critical" or not run at all, and the script exited 0
# either way.
$script:StepResults  = New-Object System.Collections.ArrayList
$script:CurrentStep  = $null
$script:StatusRank   = @{ None = -1; Skipped = 0; Unchanged = 1; Done = 2; Warning = 3; Failed = 4 }
$script:Aborted      = $false
$script:ShortcutIcon = $null
# -Undo is done when every setting in the undo record is put back. One still
# recorded, an unplugged adapter's for one, makes it exit 1.
$script:UndoMode     = [bool]$Undo
# Set when -Undo could not read a list it checks, the NAT rules for one.
# What older versions left there was not looked at, so -Undo is not done
# either: it ends UNDO INCOMPLETE and exits 1, and says to run it again.
# Each run starts with it clear.
$script:UndoUnread   = $false

function Step {
    param([int]$Num, [int]$Total, [string]$Message, [string]$Name)
    if (-not $Name) { $Name = $Message }
    $script:CurrentStep = [pscustomobject]@{
        Num    = $Num
        Name   = $Name
        Status = 'None'
        Notes  = (New-Object System.Collections.ArrayList)
        Reboot = $false
    }
    $null = $script:StepResults.Add($script:CurrentStep)
    Log "[$Num/$Total] $Message" -Colour Yellow
}

function Add-StepResult {
    <#
    .SYNOPSIS
        Records an outcome for the step that is running. The worst outcome
        a step reports is the one the summary shows, and every Detail is
        listed under it. -Reboot puts the step on the list of what takes
        effect at the next Windows restart.
    #>
    param(
        [ValidateSet('Skipped', 'Unchanged', 'Done', 'Warning', 'Failed')]
        [string]$Status,
        [string]$Detail,
        [switch]$Reboot
    )
    $s = $script:CurrentStep
    if ($null -eq $s) { return }
    if ($script:StatusRank[$Status] -gt $script:StatusRank[$s.Status]) { $s.Status = $Status }
    if ($Detail) { $null = $s.Notes.Add($Detail) }
    if ($Reboot) { $s.Reboot = $true }
}

function Add-DiscoveryReadError {
    <#
    .SYNOPSIS
        Says each read the discovery at start could not make (the
        environment's ReadErrors), now and in the summary, where they are a
        step of their own, numbered 0, with the outcome Warning. Until 6.3.1
        they were not said, and a read that failed looked the same as one
        that found nothing. Nothing is recorded when every read was made.
    #>
    param([string[]]$ReadErrors)
    $lines = @($ReadErrors | Where-Object { $_ })
    if ($lines.Count -eq 0) { return }
    $s = [pscustomobject]@{
        Num    = 0
        Name   = 'Looking at this PC'
        Status = 'Warning'
        Notes  = (New-Object System.Collections.ArrayList)
        Reboot = $false
    }
    foreach ($line in $lines) {
        Log "[!] $line" -Colour Yellow -Indent
        $null = $s.Notes.Add($line)
    }
    $null = $script:StepResults.Add($s)
}

function Write-StepSummary {
    $words   = @{ None = 'ran'; Skipped = 'skipped'; Unchanged = 'no change needed'; Done = 'changed'; Warning = 'check the note'; Failed = 'FAILED' }
    $colours = @{ None = 'White'; Skipped = 'DarkGray'; Unchanged = 'Gray'; Done = 'Green'; Warning = 'Yellow'; Failed = 'Red' }
    Write-Host "  What each step did:" -ForegroundColor Cyan
    Write-Host ""
    foreach ($s in $script:StepResults) {
        $label = "{0,2}. {1} " -f $s.Num, $s.Name
        if ($label.Length -lt 40) { $label = $label + ('.' * (40 - $label.Length)) }
        Write-Host ("    {0} {1}" -f $label, $words[$s.Status]) -ForegroundColor $colours[$s.Status]
        foreach ($n in $s.Notes) {
            Write-Host "          $n" -ForegroundColor DarkGray
        }
    }
    # A failed step is listed too. The worst outcome is the one a step shows,
    # and a step that changed one adapter and failed on another still has
    # that change waiting for the restart.
    $rebootSteps = @($script:StepResults | Where-Object { $_.Reboot -and @('Done', 'Warning', 'Failed') -contains $_.Status })
    if ($rebootSteps.Count -gt 0) {
        Write-Host ""
        Write-Host "  Restart Windows for these to take effect:" -ForegroundColor DarkYellow
        foreach ($s in $rebootSteps) {
            Write-Host "    - $($s.Name)" -ForegroundColor DarkYellow
        }
    }
}

function Get-PreventExitCode {
    <#
    .SYNOPSIS
        1 when the run stopped early or any step failed, else 0. In -Undo
        also 1 while a setting is still in the undo record, and when a list
        it checks could not be read: -Undo did not put everything back, or
        does not know that it did.
    #>
    if ($script:Aborted) { return 1 }
    foreach ($s in $script:StepResults) {
        if ($s.Status -eq 'Failed') { return 1 }
    }
    if ($script:UndoMode -and $script:UndoRecord.Count -gt 0) { return 1 }
    if ($script:UndoMode -and $script:UndoUnread) { return 1 }
    return 0
}

function Write-Box {
    param([string]$Text, [string]$Colour)
    $border = "  +----------------------------------------------+"
    Write-Host ""
    Write-Host $border -ForegroundColor $Colour
    Write-Host ("  |" + (' ' * 11) + $Text.PadRight(35) + "|") -ForegroundColor $Colour
    Write-Host $border -ForegroundColor $Colour
    Write-Host ""
}

# -- Undo record -----------------------------------------------------
# One item per setting changed: Kind, Key and the Original value. Saved to
# $UndoFile after every addition, so a run that stops half way leaves a
# record of exactly what it changed.
$script:UndoRecord = New-Object System.Collections.ArrayList
# The computer a record written on another one names. That record is not
# used (see $UndoFile), and setup moves it aside, to the path kept here.
# ByGuid is set when the record's MachineGuid is what told it apart.
$script:UndoForeign      = $null
$script:UndoForeignAside = $null
$script:UndoForeignByGuid = $false

function Get-PreventMachineGuid {
    <#
    .SYNOPSIS
        This Windows installation's MachineGuid, or '' when it cannot be
        read. Renaming the computer does not change it, and a profile copied
        to another computer does not take it along.
    #>
    # Best effort: one that cannot be read is '', as one that is not there,
    # and Read-PreventUndo then tells the record's computer by its name
    # alone, as for a record that holds no MachineGuid. Nothing is said
    # about it, because nothing is claimed from it.
    try {
        $g = Get-RegistryValue -Path 'HKLM:\SOFTWARE\Microsoft\Cryptography' -Name 'MachineGuid'
    } catch {
        return ''
    }
    if ($null -eq $g) { return '' }
    return ([string]$g).Trim()
}

function Read-PreventUndo {
    <#
    .SYNOPSIS
        Loads the undo record into $script:UndoRecord. No file is an empty
        record. A file that cannot be read stops the run before anything is
        changed: carrying on would overwrite the original values of what an
        earlier run changed. A record that names another computer, or none,
        is not loaded: $script:UndoForeign is set to that name and the file
        is left as it is.
    .DESCRIPTION
        A record names the computer that wrote it by its MachineGuid and by
        its name. The MachineGuid decides when the record and this computer
        both have one: a renamed computer keeps it, and another computer
        with the same name has its own. The name decides otherwise, as it
        did for a record written before the MachineGuid was kept.
    #>
    $script:UndoRecord = New-Object System.Collections.ArrayList
    $script:UndoForeign = $null
    $script:UndoForeignByGuid = $false
    # Until 6.3.1 Test-Path was asked, which answers $false for a file that
    # cannot be read as for one that is not there (measured 2026-11-23): a
    # record that could not be read was an empty one, and -Undo said there
    # was nothing to undo.
    $there = $null
    try {
        $there = Get-ClaudeItemOrNull -LiteralPath $UndoFile
    } catch {
        throw "The undo record $UndoFile could not be read ($(Get-ClaudeErrorText $_)). Nothing was changed. Move the file aside and rerun."
    }
    if ($null -eq $there) { return }
    $doc = $null
    try {
        $doc = [System.IO.File]::ReadAllText($UndoFile) | ConvertFrom-Json -ErrorAction Stop
    } catch {
        throw "The undo record $UndoFile could not be read ($($_.Exception.Message)). Nothing was changed. Move the file aside and rerun."
    }
    if ($null -eq $doc -or -not $doc.PSObject.Properties['Version'] -or $doc.Version -ne 1) {
        throw "The undo record $UndoFile is not a version 1 record. Nothing was changed. Move the file aside and rerun."
    }
    # A record names the computer that wrote it. One that names another came
    # with a copied or roaming profile and holds that computer's values.
    $owner = ''
    if ($doc.PSObject.Properties['ComputerName']) { $owner = [string]$doc.ComputerName }
    $ownerGuid = ''
    if ($doc.PSObject.Properties['MachineGuid']) { $ownerGuid = ([string]$doc.MachineGuid).Trim() }
    $myGuid = Get-PreventMachineGuid
    if ($ownerGuid -and $myGuid) {
        $mine = [string]::Equals($ownerGuid, $myGuid, [StringComparison]::OrdinalIgnoreCase)
        $script:UndoForeignByGuid = -not $mine
    } else {
        $mine = [string]::Equals($owner, [Environment]::MachineName, [StringComparison]::OrdinalIgnoreCase)
    }
    if (-not $mine) {
        $script:UndoForeign = $owner
        if (-not $owner) { $script:UndoForeign = 'unknown' }
        return
    }
    if (-not $doc.PSObject.Properties['Items']) { return }
    foreach ($it in @($doc.Items)) {
        if ($null -eq $it -or -not $it.PSObject.Properties['Kind'] -or -not $it.PSObject.Properties['Key']) {
            throw "The undo record $UndoFile holds an item with no Kind or Key. Nothing was changed. Move the file aside and rerun."
        }
        $orig = $null
        if ($it.PSObject.Properties['Original']) { $orig = $it.Original }
        $null = $script:UndoRecord.Add([pscustomobject]@{ Kind = [string]$it.Kind; Key = [string]$it.Key; Original = $orig })
    }
}

function Save-PreventUndo {
    <#
    .SYNOPSIS
        Writes $script:UndoRecord to $UndoFile, with this computer's name and
        MachineGuid, through a temporary file so a failed write never leaves
        half a record. An empty record deletes the file. A record written on
        another computer is never overwritten or deleted.
    #>
    if ($script:UndoForeign) {
        throw "The undo record $UndoFile was written on another computer ($($script:UndoForeign)), so it is not overwritten."
    }
    if ($script:UndoRecord.Count -eq 0) {
        # A record that could not be read is not one that is gone. Until
        # 6.3.1 Test-Path was asked, which answers $false for both
        # (measured 2026-11-23), and the record was left as it was, unsaid.
        $there = $null
        try {
            $there = Get-ClaudeItemOrNull -LiteralPath $UndoFile
        } catch {
            throw "The undo record $UndoFile could not be read ($(Get-ClaudeErrorText $_)), so it was not emptied."
        }
        if ($null -ne $there) { Remove-PreventItem -Path $UndoFile }
        return
    }
    if (-not (Test-Path -LiteralPath $UndoDir)) {
        $null = New-Item -Path $UndoDir -ItemType Directory -Force -ErrorAction Stop
    }
    $doc = [pscustomobject]@{ Version = 1; ComputerName = [Environment]::MachineName; MachineGuid = (Get-PreventMachineGuid); Items = @($script:UndoRecord) }
    $json = ConvertTo-Json -InputObject $doc -Depth 6
    $tmp = "$UndoFile.tmp"
    [System.IO.File]::WriteAllText($tmp, $json, (New-Object System.Text.UTF8Encoding($false)))
    try {
        if (Test-Path -LiteralPath $UndoFile) {
            # [NullString]::Value, not $null. Windows PowerShell passes $null
            # to a .NET string parameter as "", Replace refuses "" as a backup
            # path ("The path is not of a legal form"), and every save after
            # the first would fail.
            [System.IO.File]::Replace($tmp, $UndoFile, [NullString]::Value)
        } else {
            [System.IO.File]::Move($tmp, $UndoFile)
        }
    } catch {
        # Best effort: the error that matters is the one thrown on the next
        # line, and a temporary file left behind is overwritten next time.
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        throw
    }
}

function Get-PreventUndoItem {
    param([string]$Kind, [string]$Key)
    foreach ($it in $script:UndoRecord) {
        if ($it.Kind -eq $Kind -and $it.Key -eq $Key) { return $it }
    }
    return $null
}

function Add-PreventUndoItem {
    <#
    .SYNOPSIS
        Records a setting's value before it is changed, and saves the record
        at once. The first record of a setting wins: a rerun finds the value
        this script set, which is not the original. Throws when the record
        cannot be saved, and the caller then does not make the change.
    #>
    param([string]$Kind, [string]$Key, $Original)
    if ($null -ne (Get-PreventUndoItem -Kind $Kind -Key $Key)) { return }
    $item = [pscustomobject]@{ Kind = $Kind; Key = $Key; Original = $Original }
    $null = $script:UndoRecord.Add($item)
    try {
        Save-PreventUndo
    } catch {
        $script:UndoRecord.Remove($item)
        throw "Could not save the undo record ($($_.Exception.Message)). The change was not made."
    }
}

function Clear-PreventUndoItem {
    <#
    .SYNOPSIS
        Drops a setting from the undo record once it is put back, and saves
        the record. When the save fails the item stays where it was, so the
        record in memory is what the file holds, and the error is thrown.
    #>
    param([string]$Kind, [string]$Key)
    $it = Get-PreventUndoItem -Kind $Kind -Key $Key
    if ($null -eq $it) { return }
    $at = $script:UndoRecord.IndexOf($it)
    $script:UndoRecord.RemoveAt($at)
    try {
        Save-PreventUndo
    } catch {
        $script:UndoRecord.Insert($at, $it)
        throw
    }
}

function Move-ForeignPreventUndo {
    <#
    .SYNOPSIS
        Moves an undo record written on another computer aside, so setup can
        start this computer's own, and returns where it went. The file is
        kept, never deleted: its values are that computer's.
    #>
    $from = [string]$script:UndoForeign
    $safe = $from -replace '[^A-Za-z0-9_-]', ''
    if ($safe.Length -gt 40) { $safe = $safe.Substring(0, 40) }
    if (-not $safe) { $safe = 'unknown' }
    $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
    $aside = Join-Path $UndoDir "prevent-undo.from-$safe.$stamp.json"
    $n = 1
    while (Test-Path -LiteralPath $aside) {
        $aside = Join-Path $UndoDir "prevent-undo.from-$safe.$stamp-$n.json"
        $n++
    }
    try {
        [System.IO.File]::Move($UndoFile, $aside)
    } catch {
        throw "The undo record $UndoFile was written on another computer ($from) and could not be moved aside ($($_.Exception.Message)). Nothing was changed."
    }
    Log "[!] The undo record was written on another computer ($from), so it was not used. It is kept as $aside, and this computer's own record starts here." -Colour Yellow -Indent
    $script:UndoForeignByGuid = $false
    $script:UndoForeign = $null
    return $aside
}

function Add-LegacyInstallRecord {
    <#
    .SYNOPSIS
        Records that a version before 6.3.0 ran here, before setup deletes
        the file or task that shows it. Throws, as Add-PreventUndoItem does,
        when the record cannot be saved, and the caller then deletes nothing.
    .DESCRIPTION
        -Undo step 2 removes three values older versions set only where one
        is known to have run (see Get-LegacyInstallTrace). Setup deletes what
        shows it, so without this record a -Undo after setup could not tell.
    #>
    param([string]$What, [string]$Where)
    Add-PreventUndoItem -Kind LegacyInstall -Key $What -Original $Where
}

function Get-LegacyInstallTrace {
    <#
    .SYNOPSIS
        What shows that a version before 6.3.0 ran on this computer, as
        @{ What; Where }, or $null when nothing does: setup's record of one,
        a file only those versions wrote, or their logon task.
    .DESCRIPTION
        ServicesPipeTimeout = 120000, LocalAccountTokenFilterPolicy = 1 and
        FilterAdministratorToken = 0 are what older versions set, and also
        what an administrator, another tool or a policy can set for reasons
        of their own. -Undo removes them only where this finds an older
        version, and setup's step 19 says which of the two it is.

        Throws, naming what could not be read, when it cannot tell. Until
        6.3.1 a file or a task that could not be read counted as not there,
        and so as no older version.
    #>
    foreach ($it in $script:UndoRecord) {
        if ($it.Kind -eq 'LegacyInstall') { return [pscustomobject]@{ What = [string]$it.Key; Where = [string]$it.Original } }
    }
    $files = [ordered]@{
        'power plan record'        = $BackupFile
        'Connected Standby record' = $CsBackupFile
        'Dynamic Memory flag'      = (Join-Path $ClaudeAppData 'disable-dynamic-memory.flag')
        'watchdog script'          = (Join-Path $ClaudeAppData 'cowork-watchdog.ps1')
        'admin launcher'           = (Join-Path $ClaudeAppData 'Launch-Claude-Admin.cmd')
    }
    foreach ($k in $files.Keys) {
        # Test-Path answers $false for a path it could not read as well.
        try {
            $item = Get-ClaudeItemOrNull -LiteralPath $files[$k]
        } catch {
            throw "$($files[$k]) could not be read ($(Get-ClaudeErrorText $_))"
        }
        if ($null -ne $item) { return [pscustomobject]@{ What = $k; Where = [string]$files[$k] } }
    }
    # Get-Command answers nothing for a command that is not installed, and
    # without the ScheduledTasks module there is no logon task to find.
    if (Get-Command -Name Get-ScheduledTask -ErrorAction SilentlyContinue) {
        try {
            $boot = Get-ClaudeScheduledTaskOrNull -TaskName $BootTaskName -TaskPath $TaskPath
        } catch {
            throw "the scheduled task $TaskPath$BootTaskName could not be read ($(Get-ClaudeErrorText $_))"
        }
        if ($null -ne $boot) {
            return [pscustomobject]@{ What = 'boot-fix task'; Where = "$TaskPath$BootTaskName" }
        }
    }
    return $null
}

function Get-ClaudeVmWithoutDynamicMemory {
    <#
    .SYNOPSIS
        The Hyper-V VMs named like Claude that have Dynamic Memory off, which
        older versions' monitor turned off where disable-dynamic-memory.flag
        was set. Nothing when Hyper-V's cmdlets are not installed. An object
        without DynamicMemoryEnabled, which another product's Get-VM gives,
        is not a Hyper-V VM and is left out.
    #>
    # Get-Command answers nothing for a command that is not installed:
    # without Hyper-V's cmdlets there are no Hyper-V VMs to read.
    if (-not (Get-Command -Name Get-VM -ErrorAction SilentlyContinue)) { return }
    Get-VM -ErrorAction Stop | Where-Object {
        $_.PSObject.Properties['DynamicMemoryEnabled'] -and [string]$_.Name -match 'claude' -and -not $_.DynamicMemoryEnabled
    }
}

function Test-WinNatMissingError {
    <#
    .SYNOPSIS
        Whether an error from Get-NetNat means WinNAT is not on this PC, so
        there is no NAT rule to read.
    .DESCRIPTION
        The NetNat cmdlets come with Windows, but WinNAT's WMI class and its
        provider can be missing, as on a PC without Hyper-V or containers.
        Get-NetNat then fails on every run. Read as a rule that could not
        be read, that would end every setup and every -Undo with a warning
        about a rule that cannot be there.

        The error is known by its codes, never by its message, which is in
        the language of Windows: the CIM status Invalid namespace (3),
        Invalid class (5) or Not supported (7), or the WMI code in its error
        data, 0x8004100E (invalid namespace), 0x80041010 (invalid class),
        0x8004100C (not supported), 0x80041011 (provider not found) or
        0x80041013 (provider load failure). Any other error, such as access
        denied or a timeout, is a rule that could not be read.
    #>
    param($ErrorRecord)
    $ex = $ErrorRecord.Exception
    for ($depth = 0; $null -ne $ex -and $depth -lt 5; $depth++) {
        if ($ex.GetType().FullName -eq 'Microsoft.Management.Infrastructure.CimException') {
            if (@(3, 5, 7) -contains [int]$ex.StatusCode) { return $true }
            # Read with -as, so that a value that is not a number cannot throw
            # in place of the error this is asked about.
            $code = $null
            if ($null -ne $ex.ErrorData) {
                $p = $ex.ErrorData.CimInstanceProperties['error_Code']
                if ($null -ne $p -and $null -ne $p.Value) { $code = $p.Value -as [long] }
            }
            return ($null -ne $code -and @(0x8004100EL, 0x80041010L, 0x8004100CL, 0x80041011L, 0x80041013L) -contains $code)
        }
        $ex = $ex.InnerException
    }
    return $false
}

function Test-CimClassMissingError {
    <#
    .SYNOPSIS
        Whether an error from a CIM read means the namespace or the class it
        asked for is not on this PC, as root\SecurityCenter2 is not on
        Windows Server.
    .DESCRIPTION
        Known by its codes, never by its message, as Test-WinNatMissingError
        knows WinNAT's: the CIM status Invalid namespace (3) or Invalid
        class (5), or the WMI code in its error data, 0x8004100E (invalid
        namespace) or 0x80041010 (invalid class). Measured on Windows 11
        with Get-CimInstance: a namespace that is not there is status 3
        with 0x8004100E, and a class that is not there status 5 with
        0x80041010. Any other error, Not supported and access denied among
        them, is a list that could not be read.
    #>
    param($ErrorRecord)
    $ex = $ErrorRecord.Exception
    $depth = 0
    while ($null -ne $ex -and $depth -lt 5) {
        if ($ex.GetType().FullName -ceq 'Microsoft.Management.Infrastructure.CimException') {
            if (@(3, 5) -contains [int]$ex.StatusCode) { return $true }
            # Read with -as, so that a value that is not a number cannot
            # throw in place of the error this is asked about.
            $wmiCode = $null
            if ($null -ne $ex.ErrorData) {
                $prop = $ex.ErrorData.CimInstanceProperties['error_Code']
                if ($null -ne $prop -and $null -ne $prop.Value) { $wmiCode = $prop.Value -as [long] }
            }
            return ($null -ne $wmiCode -and @(0x8004100EL, 0x80041010L) -contains $wmiCode)
        }
        $ex = $ex.InnerException
        $depth++
    }
    return $false
}

function Test-CoworkNat {
    <#
    .SYNOPSIS
        Whether the NAT rule CoworkNAT, which versions before 6.3.0 created
        and their monitor re-created, is on this PC: $true or $false, or
        $null when WinNAT is not on this PC: its cmdlets are not installed,
        or Get-NetNat fails because its class or provider is missing.
    .DESCRIPTION
        Only an ObjectNotFound error means there is no such rule, and only
        the errors Test-WinNatMissingError knows mean there is no WinNAT. Any
        other error is thrown, so that a rule that could not be read is
        reported rather than passed over as gone.
    #>
    # Get-Command answers nothing for a command that is not installed:
    # without the NetNat module there is no NAT rule to read.
    if (-not (Get-Command -Name Get-NetNat -ErrorAction SilentlyContinue)) { return $null }
    try {
        return (@(Get-NetNat -Name 'CoworkNAT' -ErrorAction Stop).Count -gt 0)
    } catch {
        if ($_.CategoryInfo.Category -eq [System.Management.Automation.ErrorCategory]::ObjectNotFound) { return $false }
        if (Test-WinNatMissingError -ErrorRecord $_) { return $null }
        throw
    }
}

function Get-PreventProcessByIdOrNull {
    <#
    .SYNOPSIS
        The process with that id, $null when no process has it, and throws,
        with the error Get-Process gave, when it could not be read.
    .DESCRIPTION
        Measured 2026-11-22 on Windows 11, Windows PowerShell 5.1, with a
        read-only probe (Get-Process -Id 999999): error id
        NoProcessFoundForGivenId, category ObjectNotFound, exception
        Microsoft.PowerShell.Commands.ProcessCommandException. Any other
        error means the process could not be read. Until 6.3.1 every error
        read as a process that had exited.
    #>
    param([Parameter(Mandatory)][int]$Id)
    try {
        return (Get-Process -Id $Id -ErrorAction Stop)
    } catch {
        if (Test-ClaudeNotFoundError -ErrorRecord $_ -ErrorId 'NoProcessFoundForGivenId' `
                -ExceptionType 'Microsoft.PowerShell.Commands.ProcessCommandException') { return $null }
        throw
    }
}

function Wait-PreventProcessExit {
    <#
    .SYNOPSIS
        Waits up to two seconds for the process with that id to go: $true
        once it is gone, $false when it still runs after that, and throws
        when it could not be read.
    .DESCRIPTION
        A process that was stopped can take a moment to go, and one that
        ended by itself before the stop is gone as well.
    #>
    param([Parameter(Mandatory)][int]$Id)
    for ($i = 0; $i -lt 10; $i++) {
        if ($null -eq (Get-PreventProcessByIdOrNull -Id $Id)) { return $true }
        Start-Sleep -Milliseconds 200
    }
    return $false
}

function Get-PreventTaskState {
    <#
    .SYNOPSIS
        The state of the scheduled task $TaskPath$Name, such as Ready or
        Running, '' when there is no such task, and throws when it could
        not be read (see Get-ClaudeScheduledTaskOrNull).
    #>
    param([Parameter(Mandatory)][string]$Name)
    $t = Get-ClaudeScheduledTaskOrNull -TaskName $Name -TaskPath $TaskPath
    if ($null -eq $t) { return '' }
    return [string]$t.State
}

function Remove-PreventTask {
    <#
    .SYNOPSIS
        Removes the scheduled task $TaskPath$Name. Returns 'Removed',
        'NotFound', or why it could not be removed.
    .DESCRIPTION
        Only an ObjectNotFound error is "not found". Until 6.3.0 every error,
        a scheduler that refused the delete among them, was reported as a
        task that was not there.
    #>
    # -WhatIf is refused before anything runs, so there is no preview for
    # ShouldProcess to serve. -Undo reports what each removal did instead.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = '-WhatIf is refused at the top of the script')]
    param([string]$Name)
    try {
        $null = Get-ScheduledTask -TaskName $Name -TaskPath $TaskPath -ErrorAction Stop
    } catch {
        if ($_.CategoryInfo.Category -eq [System.Management.Automation.ErrorCategory]::ObjectNotFound) { return 'NotFound' }
        return "it could not be read ($($_.Exception.Message))"
    }
    try {
        Unregister-ScheduledTask -TaskName $Name -TaskPath $TaskPath -Confirm:$false -ErrorAction Stop
    } catch {
        return "it could not be removed ($($_.Exception.Message))"
    }
    try {
        $null = Get-ScheduledTask -TaskName $Name -TaskPath $TaskPath -ErrorAction Stop
    } catch {
        if ($_.CategoryInfo.Category -eq [System.Management.Automation.ErrorCategory]::ObjectNotFound) { return 'Removed' }
        return "whether it is gone could not be read ($($_.Exception.Message))"
    }
    return 'it is still there after it was removed'
}

function Get-PreventTaskService {
    <# The Task Scheduler's COM service, connected. #>
    $svc = New-Object -ComObject Schedule.Service
    $svc.Connect()
    return $svc
}

function Remove-PreventTaskFolder {
    <#
    .SYNOPSIS
        Deletes the task folder $TaskPath once it holds no task and no other
        folder. Returns 'Removed', 'NotFound', 'NotEmpty', or why it could not
        be removed. Unregister-ScheduledTask leaves the folder behind, and
        the scheduled-task cmdlets have no way to delete one.
    #>
    # -WhatIf is refused before anything runs, so there is no preview for
    # ShouldProcess to serve. -Undo reports what each removal did instead.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = '-WhatIf is refused at the top of the script')]
    param()
    $name = $TaskPath.Trim('\')
    if (-not $name) { return 'NotFound' }
    try {
        $svc = Get-PreventTaskService
    } catch {
        return "the Task Scheduler could not be reached ($($_.Exception.Message))"
    }
    $folder = $null
    try {
        $folder = $svc.GetFolder("\$name")
    } catch {
        # 0x80070002 and 0x80070003: no such folder. The COM call throws the
        # error itself; a wrapper can hold it further down the chain.
        for ($e = $_.Exception; $null -ne $e; $e = $e.InnerException) {
            if ($e.HResult -eq -2147024894 -or $e.HResult -eq -2147024893) { return 'NotFound' }
        }
        return "it could not be read ($($_.Exception.Message))"
    }
    try {
        # 1 = TASK_ENUM_HIDDEN: a hidden task counts as one.
        if ($folder.GetTasks(1).Count -gt 0 -or $folder.GetFolders(0).Count -gt 0) { return 'NotEmpty' }
        $svc.GetFolder('\').DeleteFolder($name, 0)
    } catch {
        return "it could not be removed ($($_.Exception.Message))"
    }
    try {
        $null = $svc.GetFolder("\$name")
    } catch {
        # Only "no such folder" proves it is gone. Until 6.3.0 any error on
        # this read, a scheduler that stopped answering among them, was
        # reported as the folder removed.
        for ($e = $_.Exception; $null -ne $e; $e = $e.InnerException) {
            if ($e.HResult -eq -2147024894 -or $e.HResult -eq -2147024893) { return 'Removed' }
        }
        return "whether it is gone could not be read ($($_.Exception.Message))"
    }
    return 'it is still there after it was removed'
}

function Test-PreventPathThere {
    <#
    .SYNOPSIS
        $true when there is a file or folder at Path, $false when there is
        none, and $null when that could not be read. That one is said as a
        Warning on the current step, naming the path, why, and what was
        therefore not done ($NotDone); in -Undo it also marks the undo as
        not complete.
    .DESCRIPTION
        Test-Path cannot tell these apart: for a file or folder the user
        cannot read it answers $false, as for one that is not there
        (measured 2026-11-23, an explicit Deny on the folder that holds it:
        Test-Path writes ItemExistsUnauthorizedAccessError and returns
        $false). Until 6.3.1 every file an older version left was looked
        for with it, so one that could not be read was reported as gone.
    #>
    param([string]$Path, [string]$NotDone)
    try {
        return ($null -ne (Get-ClaudeItemOrNull -LiteralPath $Path))
    } catch {
        $said = "$Path could not be read ($(Get-ClaudeErrorText $_)), so $NotDone"
        Log "[!] $said" -Colour Yellow -Indent
        Add-StepResult -Status Warning -Detail $said
        if ($script:UndoMode) { $script:UndoUnread = $true }
        return $null
    }
}

function Find-PreventKitFile {
    <#
    .SYNOPSIS
        The first of Candidates that is a file, or $null when none is. A
        place that could not be read throws, naming it and why: the search
        does not go on to a later place, so nothing is installed from
        another copy in its stead. Until 6.3.1 Test-Path was asked, which
        answers $false for a file that cannot be read as for one that is
        not there (measured 2026-11-23).
    #>
    param([string[]]$Candidates)
    foreach ($c in @($Candidates)) {
        if (-not $c) { continue }
        $item = $null
        try {
            $item = Get-ClaudeItemOrNull -LiteralPath $c
        } catch {
            throw "$c could not be read ($(Get-ClaudeErrorText $_))"
        }
        if ($null -ne $item -and -not $item.PSIsContainer) { return $c }
    }
    return $null
}

function Remove-PreventItem {
    <#
    .SYNOPSIS
        Removes one file, and throws, naming it and why, when it cannot.
    .DESCRIPTION
        For a file the user may not delete, the error Remove-Item gives does
        not name it ("Access to the path is denied.", measured 2026-11-23),
        and the steps that remove a file this way report the error as it
        comes. Until 6.3.1 they said what failed, but not on which file.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = '-WhatIf is refused at the top of the script')]
    param([string]$Path)
    try {
        Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
    } catch {
        throw "$Path could not be removed ($(Get-ClaudeErrorText $_))"
    }
}

function Remove-PreventFile {
    <#
    .SYNOPSIS
        Deletes a file this script or an older version made, and reports it
        to the current step: Done once it is gone, Failed while it is still
        there or whether it is gone could not be read, and a Warning when
        whether it is there could not be read. Returns $true when there was
        a file, or one could not be ruled out. Until 6.3.0 -Undo logged
        every file it tried to delete as removed. Until 6.3.1 a file that
        could not be read was taken for none, so -Undo said "No shortcuts
        found" while one was left.
    #>
    # -WhatIf is refused before anything runs, so there is no preview for
    # ShouldProcess to serve. -Undo reports what each removal did instead.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = '-WhatIf is refused at the top of the script')]
    param([string]$Path, [string]$What = 'Removed')
    $there = Test-PreventPathThere -Path $Path -NotDone 'it was not removed'
    if ($null -eq $there) { return $true }
    if (-not $there) { return $false }
    $why = $null
    try {
        Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
    } catch {
        $why = $_.Exception.Message
    }
    $left = $null
    try {
        $left = Get-ClaudeItemOrNull -LiteralPath $Path
    } catch {
        $gone = "whether it is gone could not be read ($(Get-ClaudeErrorText $_))"
        if ($why) { $why = "$why; $gone" } else { $why = $gone }
        $left = $true
    }
    if ($null -ne $left) {
        if (-not $why) { $why = 'it is still there' }
        Log "[!] Could not remove ${Path}: $why" -Colour Red -Indent
        Add-StepResult -Status Failed -Detail "Could not remove ${Path}: $why"
    } else {
        Log "${What}: $Path" -Colour Green -Indent
        Add-StepResult -Status Done
    }
    return $true
}

function Install-HealthMonitorTask {
    <#
    .SYNOPSIS
        Step 15's install: registers the logon task that runs the health
        monitor at WatchScript, stops a monitor already running, and starts
        the task now when that cannot leave two monitors running. Each
        outcome goes to the step.
    .DESCRIPTION
        The monitor guards itself with the mutex Global\ClaudeHealthMonitor:
        a copy started while another holds it exits, 120 seconds after the
        task starts it, without watching anything. That guard does not
        cover every monitor. Versions up to 5.0.0 put their version in the
        mutex's name, so one of them and this one do not see each other,
        and a monitor that could not create its mutex runs without one. Two
        monitors running at once each repair Claude on their own.

        So the task is started now only when the running processes were
        listed and every monitor among them stopped. When the list cannot
        be read, or a monitor does not stop, that is a Warning that names
        it and why, and the task, registered all the same, starts at the
        next logon. A start that fails is a Warning too. Until 6.3.0 the
        list was read with SilentlyContinue and each stop's error was
        dropped, so a monitor that went on running was never mentioned and
        the task was started beside it; a start that failed was dropped as
        well, and the step said "Health monitor installed" either way.

        The task is registered first, with -Force, which replaces the one an
        earlier run made, and with -ErrorAction Stop: the ScheduledTasks
        cmdlets report a refusal as an error that does not stop the script.
        Until 6.3.0 the old task was removed and the running monitor stopped
        before the new task was registered, and a refusal was not caught, so
        a registration that failed left no monitor at all and was logged as
        installed. Now it is Failed, and the monitor already running is left
        as it is.
    #>
    param([string]$WatchScript)

    try {
        $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

        # In single quotes, a ' in the path is written twice.
        $delayedWatchCmd = "Start-Sleep -Seconds 120; & '$($WatchScript -replace "'", "''")' -Quiet"
        $action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
                      -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command `"$delayedWatchCmd`""

        $trigger = New-ScheduledTaskTrigger -AtLogOn

        $settings = New-ScheduledTaskSettingsSet `
                        -AllowStartIfOnBatteries `
                        -DontStopIfGoingOnBatteries `
                        -StartWhenAvailable `
                        -DontStopOnIdleEnd `
                        -RestartCount 3 `
                        -RestartInterval (New-TimeSpan -Minutes 1) `
                        -ExecutionTimeLimit (New-TimeSpan -Days 365)

        $principal = New-ScheduledTaskPrincipal `
                         -UserId $currentUser `
                         -RunLevel Highest `
                         -LogonType S4U

        Register-ScheduledTask `
            -TaskName $TaskName `
            -TaskPath $TaskPath `
            -Action $action `
            -Trigger $trigger `
            -Settings $settings `
            -Principal $principal `
            -Description "Monitors Claude logs for VirtioFS mount failures and auto-runs the fix script. Polls every 30s." `
            -Force -ErrorAction Stop | Out-Null
    } catch {
        Log "[!] Could not create health monitor task: $($_.Exception.Message)" -Colour Red -Indent
        Log "You can run Watch-ClaudeHealth.bat manually instead" -Colour DarkGray -Indent
        Add-StepResult -Status Failed -Detail "Could not create the health monitor task: $($_.Exception.Message)"
        return
    }
    Log "Health monitor installed (starts 120s after logon, polls every 30s)" -Colour Green -Indent

    # A monitor already running is stopped, so that the new one replaces it.
    $startNow = $true
    $running = @()
    try {
        $running = @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" -ErrorAction Stop |
                     Where-Object { [string]$_.CommandLine -match 'Watch-ClaudeHealth' })
    } catch {
        $startNow = $false
        $said = "Could not list the running processes, so a health monitor already running was not looked for or stopped, and the new one was not started now; it starts at the next logon: $($_.Exception.Message)"
        Log "[!] $said" -Colour Yellow -Indent
        Add-StepResult -Status Warning -Detail $said
    }
    foreach ($mon in $running) {
        $monId = [int]$mon.ProcessId
        $why = $null
        try {
            Stop-Process -Id $monId -Force -ErrorAction Stop
        } catch {
            $why = $_.Exception.Message
        }
        # Until 6.3.1 a process that could not be read counted as gone.
        try {
            $alive = -not (Wait-PreventProcessExit -Id $monId)
        } catch {
            $startNow = $false
            $said = "Could not read whether the health monitor already running (PID $monId) stopped, so the new one was not started now; it starts at the next logon: $(Get-ClaudeErrorText $_)"
            Log "[!] $said" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail $said
            continue
        }
        if ($alive) {
            $startNow = $false
            if (-not $why) { $why = 'it is still running' }
            $said = "Could not stop the health monitor already running (PID $monId), so it was not replaced, and the new one was not started now; it starts at the next logon: $why"
            Log "[!] $said" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail $said
        } elseif ($why) {
            # The stop failed, and the process is gone: it ended by itself.
            Log "The health monitor that was running (PID $monId) has ended" -Colour DarkGray -Indent
        } else {
            Log "Stopped the health monitor already running (PID $monId)" -Colour DarkGray -Indent
        }
    }

    if ($startNow) {
        # The task runs one instance at a time: MultipleInstances is
        # IgnoreNew, the default of New-ScheduledTaskSettingsSet (read
        # 2026-11-22 from the monitor's task on Windows 11), and the scheduler
        # ignores a start, with no error, while it reads an instance as
        # running. After the old monitor is stopped the task can read running
        # a moment longer. So the start waits up to 5 seconds for that to
        # end, and afterwards checks that the task runs: its action waits 120
        # seconds before the monitor starts, so a start that took reads
        # running for that long. Until 6.3.1 neither was checked, and
        # "Started it now" was said of a start the scheduler ignored.
        $notStarted = $null
        try {
            $state = Get-PreventTaskState -Name $TaskName
            for ($i = 0; $i -lt 20 -and $state -eq 'Running'; $i++) {
                Start-Sleep -Milliseconds 250
                $state = Get-PreventTaskState -Name $TaskName
            }
            if ($state -eq 'Running') {
                $notStarted = "it still reads as running 5 seconds after the old monitor was stopped, and the scheduler ignores a start while it does"
            }
        } catch {
            $notStarted = "its state could not be read ($(Get-ClaudeErrorText $_))"
        }
        $started = $false
        if ($notStarted) {
            $said = "The health monitor task was not started now, because $notStarted; the monitor starts at the next logon"
            Log "[!] $said" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail $said
        } else {
            try {
                Start-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction Stop
                $started = $true
            } catch {
                $said = "The health monitor task could not be started now, so the monitor starts at the next logon: $($_.Exception.Message)"
                Log "[!] $said" -Colour Yellow -Indent
                Add-StepResult -Status Warning -Detail $said
            }
        }
        if ($started) {
            $said = $null
            try {
                $state = Get-PreventTaskState -Name $TaskName
                for ($i = 0; $i -lt 12 -and $state -ne 'Running'; $i++) {
                    Start-Sleep -Milliseconds 250
                    $state = Get-PreventTaskState -Name $TaskName
                }
                if ($state -eq 'Running') {
                    Log "Started it now: the monitor begins watching in 120 seconds" -Colour Green -Indent
                } elseif ($state) {
                    $said = "The scheduler accepted the start of the health monitor task, but the task did not start: it reads $state 3 seconds later. The monitor starts at the next logon"
                } else {
                    $said = "The scheduler accepted the start of the health monitor task, but the task is no longer there 3 seconds later"
                }
            } catch {
                $said = "The health monitor task was started, but its state could not be read to check that it runs: $(Get-ClaudeErrorText $_)"
            }
            if ($said) {
                Log "[!] $said" -Colour Yellow -Indent
                Add-StepResult -Status Warning -Detail $said
            }
        }
    }

    Log "Task: Task Scheduler > $TaskPath$TaskName" -Colour DarkGray -Indent
    Log "Script: $WatchScript" -Colour DarkGray -Indent
    Log "Logs: $ClaudeAppData\watch-logs\" -Colour DarkGray -Indent
    Add-StepResult -Status Done
}

function Get-PreventUndoItemLabel {
    <#
    .SYNOPSIS
        One recorded setting in words, for the list -Undo prints of what it
        did not put back.
    #>
    param($Item)
    $key = [string]$Item.Key
    $parts = $key.Split('|')
    $last = $parts[$parts.Count - 1]
    switch ([string]$Item.Kind) {
        'PowerValue' {
            $label = $last
            foreach ($o in $PowerOptions) { if ($o.Setting -eq $last) { $label = $o.Label; break } }
            return "Power setting '$label' of the plan $($parts[0])"
        }
        'PowerPlan' {
            if ($key -eq 'Created') { return "The Ultimate Performance plan this script created ($($Item.Original))" }
            return "The power plan you had ($($Item.Original))"
        }
        'Hibernate' { return 'Hibernate' }
        'Registry' {
            $i = $key.LastIndexOf('|')
            if ($i -lt 1) { return "Registry value $key" }
            return "Registry value $last under $($key.Substring(0, $i))"
        }
        'RegistryKey'           { return "Registry key this script created: $key" }
        'NicDeviceEnable'       { return "Network adapter $($parts[0]): Windows may turn it off to save power" }
        'NicAdvanced'           { return "Network adapter $($parts[0]): its setting $last" }
        'NicPowerManagement'    { return "Network adapter $($parts[0]): its power option $last" }
        'DefenderPath'          { return "Defender exclusion of $key" }
        'DefenderProcess'       { return "Defender exclusion of the process $key" }
        'ServiceFailureActions' { return "Recovery settings of the service $key" }
        'LegacyInstall'         { return "What an older version set here (it left its $key)" }
        default                 { return "$($Item.Kind) $key" }
    }
}

function Write-UndoOutcome {
    <#
    .SYNOPSIS
        The end of -Undo: a box that says how it went, what each step did,
        and every setting still in the undo record, which -Undo did not put
        back. A run that left settings in the record, an unplugged adapter's
        for one, used to show UNDO COMPLETE and exit 0. So did a run that
        could not read a list it checks, such as the NAT rules: what older
        versions left there was not looked at, so that run is not complete
        either.
    #>
    $left = @($script:UndoRecord.ToArray())
    $failed = @($script:StepResults | Where-Object { $_.Status -eq 'Failed' }).Count -gt 0
    if ($failed) {
        Write-Box -Text "UNDO FINISHED WITH ERRORS" -Colour Red
    } elseif ($left.Count -gt 0 -or $script:UndoUnread) {
        Write-Box -Text "UNDO INCOMPLETE" -Colour Yellow
    } else {
        Write-Box -Text "UNDO COMPLETE" -Colour Green
    }
    Write-StepSummary
    if ($script:UndoForeign) {
        Write-Host ""
        Write-Host "  The undo record was written on another computer ($($script:UndoForeign))," -ForegroundColor Yellow
        Write-Host "  so nothing was put back from it. It is left as it is:" -ForegroundColor Yellow
        Write-Host "    $UndoFile" -ForegroundColor Yellow
    }
    if ($script:UndoUnread) {
        Write-Host ""
        Write-Host "  Some lists could not be read (the warnings above say which), so" -ForegroundColor Yellow
        Write-Host "  what is in them was not checked. Run -Undo again once that is" -ForegroundColor Yellow
        Write-Host "  resolved." -ForegroundColor Yellow
    }
    if ($left.Count -gt 0) {
        Write-Host ""
        Write-Host "  Not put back, and still in the undo record:" -ForegroundColor Yellow
        foreach ($it in $left) {
            Write-Host "    - $(Get-PreventUndoItemLabel -Item $it)" -ForegroundColor Yellow
        }
        Write-Host ""
        Write-Host "  The record is kept in:" -ForegroundColor Yellow
        Write-Host "    $UndoFile" -ForegroundColor Yellow
        Write-Host "  Run -Undo again once the cause shown above is resolved, and it" -ForegroundColor Yellow
        Write-Host "  puts these back." -ForegroundColor Yellow
        if (@($left | Where-Object { ([string]$_.Kind).StartsWith('Nic') }).Count -gt 0) {
            Write-Host "  For a network adapter, run it with that adapter plugged in." -ForegroundColor Yellow
        }
    }
}

# -- Power -----------------------------------------------------------
function Invoke-PowerCfg {
    <#
    .SYNOPSIS
        Runs powercfg and returns its exit code and its output as one
        string, so a caller can check both. Until 6.3.0 neither was read,
        and a powercfg that refused a change was logged as having made it.
    #>
    param([string[]]$Arguments)
    $out = & powercfg.exe @Arguments 2>&1
    $code = $LASTEXITCODE
    $lines = @()
    foreach ($l in @($out)) { if ($null -ne $l) { $lines += [string]$l } }
    return [pscustomobject]@{ ExitCode = $code; Output = ($lines -join "`n") }
}

function Test-PowerSettingGone {
    <#
    .SYNOPSIS
        True when Windows no longer defines this power setting: its key under
        Control\Power\PowerSettings is gone while that folder is there. A
        driver or an update can take a setting away, and powercfg then says
        it does not exist, in the machine's language, so its message is not
        what is read here.
    #>
    param([string]$Sub, [string]$Setting)
    $root = 'HKLM:\SYSTEM\CurrentControlSet\Control\Power\PowerSettings'
    if (-not (Test-Path -LiteralPath $root)) { return $false }
    return (-not (Test-Path -LiteralPath "$root\$Sub\$Setting"))
}

function Get-ActivePlanGuid {
    $r = Invoke-PowerCfg -Arguments @('/getactivescheme')
    if ($r.Output -match "($GuidPattern)") { return $Matches[1].ToLowerInvariant() }
    return $null
}

function Test-PowerPlanExists {
    <#
    .SYNOPSIS
        True when powercfg lists the plan, False when it does not, and $null
        when powercfg could not list the plans. Every caller keeps what it
        would drop for a missing plan when the answer is $null: a powercfg
        failure used to read as every plan gone, and -Undo cleared records
        it had not put back.
    #>
    param([string]$Plan)
    if (-not $Plan) { return $false }
    $r = Invoke-PowerCfg -Arguments @('/list')
    if ($r.ExitCode -ne 0) { return $null }
    return ($r.Output -match [regex]::Escape($Plan))
}

function Get-UltimatePlanGuids {
    <#
    .SYNOPSIS
        This machine's Ultimate Performance plans: the copies that
        "powercfg /duplicatescheme" made of the hidden template.
    .DESCRIPTION
        Found by the plan's FriendlyName, which for a copy of the template is
        the resource reference "@%SystemRoot%\system32\powrprof.dll,-19,
        Ultimate Performance". "powercfg /list" shows the translated name,
        so the test the earlier versions used, the words "Ultimate
        Performance" in that list, found nothing on a non-English Windows
        and every run made one more copy.

        Throws when the plans could not be read. A key that is not there
        holds no plan: Get-ChildItem says so with error id PathNotFound,
        category ObjectNotFound, exception
        System.Management.Automation.ItemNotFoundException (measured
        2026-11-22 on Windows 11, Windows PowerShell 5.1, with a read-only
        probe of a missing registry key). Until 6.3.1 every error read as no
        plan, and setup then made one more copy.
    #>
    $root = "HKLM:\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes"
    $found = @()
    $keys = @()
    try {
        $keys = @(Get-ChildItem -LiteralPath $root -ErrorAction Stop)
    } catch {
        if (-not (Test-ClaudeNotFoundError -ErrorRecord $_ -ErrorId 'PathNotFound' `
                -ExceptionType 'System.Management.Automation.ItemNotFoundException')) {
            throw "the power plans in the registry could not be read ($(Get-ClaudeErrorText $_))"
        }
    }
    foreach ($k in $keys) {
        $g = ([string]$k.PSChildName).ToLowerInvariant()
        if ($g -eq $UltimateTemplateGuid -or $g -notmatch "^$GuidPattern$") { continue }
        $fn = $k.GetValue('FriendlyName')
        if ($fn -and ([string]$fn) -match 'powrprof\.dll,-19,') { $found += $g }
    }
    return @($found)
}

function Get-AcPowerIndex {
    <#
    .SYNOPSIS
        The AC value of one power setting in one plan, or $null when this
        machine does not have that setting.
    .DESCRIPTION
        "powercfg /qh" ends a setting with its AC and its DC value:

            Current AC Power Setting Index: 0x00000064
            Current DC Power Setting Index: 0x00000064

        The labels are translated on a non-English Windows, so the two lines
        are found by being last and by their hex number, not by their words.
        /qh rather than /q, so hidden settings such as the NVMe timeouts are
        read too. A setting the machine does not have makes powercfg exit 1.
    #>
    param([string]$Plan, [string]$Sub, [string]$Setting)
    $r = Invoke-PowerCfg -Arguments @('/qh', $Plan, $Sub, $Setting)
    if ($r.ExitCode -ne 0) { return $null }
    $lines = @()
    foreach ($l in ($r.Output -split "`n")) { if ($l.Trim()) { $lines += $l.Trim() } }
    if ($lines.Count -lt 2) { return $null }
    $dc = $lines[$lines.Count - 1]
    $ac = $lines[$lines.Count - 2]
    if ($dc -notmatch ':\s*0x[0-9a-fA-F]+$') { return $null }
    if ($ac -match ':\s*0x([0-9a-fA-F]+)$') { return [Convert]::ToInt64($Matches[1], 16) }
    return $null
}

function Write-AcPowerValue {
    <#
    .SYNOPSIS
        Sets one AC power setting in a plan, recording the value it had
        first, and reads it back. Returns Missing, Unchanged, Done or Failed.
    #>
    param([string]$Plan, [string]$Sub, [string]$Setting, [long]$Value)
    $current = Get-AcPowerIndex -Plan $Plan -Sub $Sub -Setting $Setting
    if ($null -eq $current) { return 'Missing' }
    if ($current -eq $Value) { return 'Unchanged' }
    Add-PreventUndoItem -Kind PowerValue -Key "$Plan|$Sub|$Setting" -Original $current
    $r = Invoke-PowerCfg -Arguments @('/setacvalueindex', $Plan, $Sub, $Setting, [string]$Value)
    $after = Get-AcPowerIndex -Plan $Plan -Sub $Sub -Setting $Setting
    if ($r.ExitCode -eq 0 -and $null -ne $after -and $after -eq $Value) { return 'Done' }
    return 'Failed'
}

function Get-ModernStandbyState {
    <#
    .SYNOPSIS
        From "powercfg /a": Enabled when the machine uses S0 low power idle
        (Modern Standby), Disabled when it does not, $null when the text
        cannot be read.
    .DESCRIPTION
        The output lists the sleep states available and the ones that are
        not, each block an unindented header with its states indented below
        it. A state that is not available has the reasons why indented
        deeper under it:

            The following sleep states are available on this system:
                Standby (S3)

            The following sleep states are not available on this system:
                Standby (S0 Low Power Idle)
                    The system firmware does not support this standby state.

        The headers are translated and the "(S0" tag is not, so S0 is found
        by its tag on a line at the states' indent, and it is in use when no
        reason is indented under it.
        The first header used to be taken for the available block. With no
        state available powercfg prints only the other block, and S0 listed
        there read as in use, on every run, however often Windows restarted.
    #>
    param([string]$Text)
    if (-not $Text) { return $null }
    # Each line that is not blank, with the width of its indent. A tab
    # reaches the next multiple of 8, as on screen: the states are indented
    # with spaces and the reasons under them with tabs.
    $rows = New-Object System.Collections.ArrayList
    foreach ($line in ($Text -split "`r?`n")) {
        if (-not $line.Trim()) { continue }
        $width = 0
        foreach ($ch in $line.ToCharArray()) {
            if ($ch -eq [char]32) { $width++ }
            elseif ($ch -eq [char]9) { $width = $width - ($width % 8) + 8 }
            else { break }
        }
        $null = $rows.Add([pscustomobject]@{ Width = $width; Text = $line })
    }
    $sawHeader = $false
    $inUse = $false
    $stateWidth = -1
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $row = $rows[$i]
        if ($row.Width -eq 0) { $sawHeader = $true; $stateWidth = -1; continue }
        # The first line under a header is a state, and sets the states'
        # indent for its block.
        if ($stateWidth -lt 0) { $stateWidth = $row.Width }
        if ($row.Width -ne $stateWidth -or $row.Text -notmatch '\(S0\b') { continue }
        $reason = ($i + 1 -lt $rows.Count) -and ($rows[$i + 1].Width -gt $row.Width)
        if (-not $reason) { $inUse = $true }
    }
    if (-not $sawHeader) { return $null }
    if ($inUse) { return 'Enabled' }
    return 'Disabled'
}

# -- Registry --------------------------------------------------------
function Get-RegistryValue {
    <#
    .SYNOPSIS
        A registry value as stored, or $null when the key or the value is
        absent. Throws, naming the value, the key and why, when it could not
        be read.
    .DESCRIPTION
        A REG_BINARY or REG_MULTI_SZ value is returned whole, as the byte[] or
        string[] it is. A bare "return" of an array hands its elements down
        the pipeline one by one, and the caller gets an object[] of them, so
        a FailureActions value never read as a byte[] and step 9 could not
        recognise the recovery it had just set.

        Until 6.3.1 every error here was read as $null, the answer for a
        value that is not there. A key Windows will not open (measured
        2026-11-23, HKLM:\SAM\SAM: System.Security.SecurityException,
        PermissionDenied) was then one with no value: a step recorded that
        there was none before it wrote one, so -Undo would delete the value
        instead of putting it back, and -Undo said it had removed a value
        it never looked at. Get-ClaudeItemOrNull tells the key that is not
        there from the one that cannot be read; Test-Path cannot, as it
        answers $true for a key it cannot open, and even for a child of it
        that does not exist (measured the same day). Get-PreventMachineGuid
        catches the error; every other caller runs inside a try that reports
        it as the step's outcome, so nothing is changed from a value that
        was not read.
    #>
    param([string]$Path, [string]$Name)
    try {
        $k = Get-ClaudeItemOrNull -LiteralPath $Path
        if ($null -eq $k) { return $null }
        if (@($k.GetValueNames()) -notcontains $Name) { return $null }
        $v = $k.GetValue($Name)
    } catch {
        throw "The registry value $Name in $Path could not be read ($(Get-ClaudeErrorText $_))"
    }
    if ($v -is [array]) { return ,$v }
    return $v
}

function Write-RecordedDword {
    <#
    .SYNOPSIS
        Sets a DWORD registry value, recording what was there first: the old
        value, or that there was none. A missing key is created and recorded
        too, so -Undo removes it again. Returns Unchanged, Done or Failed.
    #>
    param([string]$Path, [string]$Name, [int]$Value)
    $current = Get-RegistryValue -Path $Path -Name $Name
    if ($current -is [int] -and $current -eq $Value) { return 'Unchanged' }
    if (-not (Test-Path -LiteralPath $Path)) {
        Add-PreventUndoItem -Kind RegistryKey -Key $Path -Original $null
        try {
            $null = New-Item -Path $Path -Force -ErrorAction Stop
        } catch {
            return 'Failed'
        }
    }
    Add-PreventUndoItem -Kind Registry -Key "$Path|$Name" -Original $current
    try {
        Set-ItemProperty -LiteralPath $Path -Name $Name -Value $Value -Type DWord -ErrorAction Stop
    } catch {
        return 'Failed'
    }
    $after = Get-RegistryValue -Path $Path -Name $Name
    if ($after -is [int] -and $after -eq $Value) { return 'Done' }
    return 'Failed'
}

# -- Services --------------------------------------------------------
function Get-ServiceFailureBytes {
    <# A service's FailureActions registry value, or $null when it has none. #>
    param([string]$Service)
    $v = Get-RegistryValue -Path "HKLM:\SYSTEM\CurrentControlSet\Services\$Service" -Name 'FailureActions'
    if ($v -is [byte[]]) { return ,$v }
    return $null
}

function ConvertFrom-FailureActionsBlob {
    <#
    .SYNOPSIS
        Decodes a FailureActions value, the SERVICE_FAILURE_ACTIONS that
        sc.exe stores: the reset period in seconds at offset 0, the number
        of actions at 12, and from offset 20 the actions, each a type (0
        none, 1 restart, 2 reboot, 3 run a command) and a delay in
        milliseconds. $null when there is nothing to decode.
    #>
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -lt 20) { return $null }
    $count = [int][BitConverter]::ToUInt32($Bytes, 12)
    if ($count -gt 64 -or $Bytes.Length -lt (20 + 8 * $count)) { return $null }
    $actions = @()
    for ($i = 0; $i -lt $count; $i++) {
        $o = 20 + 8 * $i
        $actions += [pscustomobject]@{ Type = [int][BitConverter]::ToUInt32($Bytes, $o); Delay = [long][BitConverter]::ToUInt32($Bytes, $o + 4) }
    }
    return [pscustomobject]@{ Reset = [long][BitConverter]::ToUInt32($Bytes, 0); Actions = @($actions) }
}

function Test-RestartRecovery {
    <#
    .SYNOPSIS
        True when a decoded FailureActions is the recovery step 9 sets:
        restart after 30, 60 and 120 seconds, the count reset after 300.
        Versions before 6.3.0 set the same on vmcompute and on the Cowork
        VM service, so -Undo uses this to recognise what they left.
    #>
    param($Decoded)
    if ($null -eq $Decoded -or $Decoded.Reset -ne 300) { return $false }
    $a = @($Decoded.Actions)
    $want = @(30000, 60000, 120000)
    if ($a.Count -ne $want.Count) { return $false }
    for ($i = 0; $i -lt $want.Count; $i++) {
        if ($a[$i].Type -ne 1 -or $a[$i].Delay -ne $want[$i]) { return $false }
    }
    return $true
}

# -- Storage ---------------------------------------------------------
function Get-CloudSyncFolderName {
    <#
    .SYNOPSIS
        The cloud service whose sync folder holds Path, or $null.
    .DESCRIPTION
        Until 6.3.0 this was a substring test for "OneDrive", "Google Drive",
        "Dropbox", "iCloud" and "Box" anywhere in the path, so a folder
        such as "Xbox" or "Boxes" read as a Box sync folder. Now: a path
        under a OneDrive folder Windows names in the environment, or a
        folder in the path that is a sync folder by name, or starts with
        that name and then " - " or " (", as "OneDrive - Contoso" and
        "Dropbox (Personal)" do. Below the user profile only the part after
        it is tested, so a user named Box is not a Box folder, and nor is
        the profile folder itself.
    #>
    param([string]$Path)
    if (-not $Path) { return $null }
    $p = $Path.TrimEnd('\')
    foreach ($ev in @('OneDrive', 'OneDriveConsumer', 'OneDriveCommercial')) {
        $root = [Environment]::GetEnvironmentVariable($ev)
        if (-not $root) { continue }
        $r = $root.TrimEnd('\')
        if ($p.Equals($r, [StringComparison]::OrdinalIgnoreCase) -or
            $p.StartsWith("$r\", [StringComparison]::OrdinalIgnoreCase)) { return 'OneDrive' }
    }
    $rest = $p
    $prof = [string]$env:USERPROFILE
    if ($prof) {
        $prof = $prof.TrimEnd('\')
        if ($p.Equals($prof, [StringComparison]::OrdinalIgnoreCase)) {
            $rest = ''
        } elseif ($p.StartsWith("$prof\", [StringComparison]::OrdinalIgnoreCase)) {
            $rest = $p.Substring($prof.Length + 1)
        }
    }
    $names = [ordered]@{
        'OneDrive'     = 'OneDrive'
        'Google Drive' = 'Google Drive'
        'My Drive'     = 'Google Drive'
        'Dropbox'      = 'Dropbox'
        'iCloudDrive'  = 'iCloud'
        'iCloud Drive' = 'iCloud'
        'Box'          = 'Box'
        'Box Sync'     = 'Box'
    }
    foreach ($seg in ($rest -split '\\')) {
        foreach ($n in $names.Keys) {
            if ($seg -eq $n -or $seg -like "$n - *" -or $seg -like "$n (*") { return $names[$n] }
        }
    }
    return $null
}

# -- Icon ------------------------------------------------------------
function Save-ClaudeIcon {
    <#
    .SYNOPSIS
        Copies Claude's icon out of Claude.exe into an .ico file. True when
        the file was written.
    .DESCRIPTION
        The whole icon group is copied, every size the exe carries, by
        reading it as a resource with LoadLibraryEx. Where that fails,
        System.Drawing's ExtractAssociatedIcon gives one size, which still
        serves a shortcut. Written to a temporary file first, so a failure
        never leaves half an icon in place of a good one.
    #>
    param([string]$ExePath, [string]$Destination)
    if (-not $ExePath -or -not (Test-Path -LiteralPath $ExePath)) { return $false }
    $tmp = "$Destination.tmp"
    $ok = $false
    try {
        $type = 'ClaudeFix.IconExtractor' -as [type]
        if (-not $type) {
            Add-Type -TypeDefinition $script:IconExtractorSource -ErrorAction Stop
            $type = 'ClaudeFix.IconExtractor' -as [type]
        }
        if ($type) { $ok = [bool]$type::ExtractGroup($ExePath, $tmp) }
    } catch {
        $ok = $false
    }
    if (-not $ok) {
        try {
            Add-Type -AssemblyName System.Drawing -ErrorAction Stop
            $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($ExePath)
            if ($icon) {
                $fs = [System.IO.File]::Create($tmp)
                try { $icon.Save($fs) } finally { $fs.Dispose(); $icon.Dispose() }
                $ok = $true
            }
        } catch {
            $ok = $false
        }
    }
    if ($ok) {
        try {
            [System.IO.File]::Copy($tmp, $Destination, $true)
        } catch {
            $ok = $false
        }
    }
    # Best effort: $ok already says whether the icon was saved, and a
    # temporary file left behind is overwritten next time.
    if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
    return $ok
}

function Initialize-ClaudeIcon {
    <#
    .SYNOPSIS
        The IconLocation for the shortcuts this script makes: Claude's icon,
        copied into the toolkit folder where an update cannot remove it, or
        a stock Windows icon when there is no Claude to copy it from.
    #>
    if ($script:ShortcutIcon) { return $script:ShortcutIcon }
    if ((Save-ClaudeIcon -ExePath $script:ClaudeEnv.ClaudeExe -Destination $ClaudeIconFile) -or
        (Test-Path -LiteralPath $ClaudeIconFile)) {
        $script:ShortcutIcon = "$ClaudeIconFile,0"
    } else {
        $script:ShortcutIcon = "%SystemRoot%\System32\shell32.dll,77"
    }
    return $script:ShortcutIcon
}

function Get-PowerPlanName {
    <# A plan's name as "powercfg /list" shows it, or the GUID when it has none. #>
    param([string]$Plan)
    if (-not $Plan) { return '(unknown)' }
    $r = Invoke-PowerCfg -Arguments @('/list')
    if ($r.Output -match ([regex]::Escape($Plan) + '\s+\(([^)]+)\)')) { return "$($Matches[1]) ($Plan)" }
    return $Plan
}

function Complete-Step {
    <# A step that ran without reporting anything changed nothing. #>
    if ($null -ne $script:CurrentStep -and $script:CurrentStep.Status -eq 'None') {
        Add-StepResult -Status Unchanged
    }
}

function Get-NicAdvancedValue {
    <#
    .SYNOPSIS
        What a driver option is set to, as a string array: the value in the
        registry, or the driver's default when the registry holds none.
    #>
    param($Property)
    $v = ConvertTo-StringArray $Property.RegistryValue
    if ($v.Count -eq 0) { $v = ConvertTo-StringArray $Property.DefaultRegistryValue }
    return ,$v
}

function Get-PreventNicAdvanced {
    <#
    .SYNOPSIS
        The driver options of the adapter named Name, matched exactly.
    .DESCRIPTION
        The NetAdapter cmdlets read -Name as a wildcard pattern, so until
        6.3.0 an adapter named "Ethernet [LAN]" matched nothing, without an
        error, and every option on it was passed over. Each wildcard
        character in the name is asked for as ?, which matches that
        character too, and the names that come back are compared as plain
        text.

        Only this adapter is read, not every adapter at once: then one
        adapter whose driver cannot be read would fail all the others with
        it. A plain name that matches nothing is an ObjectNotFound error from
        the cmdlet; it gives nothing here, as an adapter with no options
        does.
    #>
    param([string]$Name)
    $pattern = $Name -replace '[\[\]*?`]', '?'
    try {
        $rows = @(Get-NetAdapterAdvancedProperty -Name $pattern -AllProperties -ErrorAction Stop)
    } catch {
        if ($_.CategoryInfo.Category -eq [System.Management.Automation.ErrorCategory]::ObjectNotFound) { return }
        throw
    }
    $rows | Where-Object { [string]$_.Name -eq $Name }
}

function Get-PreventNicPower {
    <#
    .SYNOPSIS
        The power management settings of the adapter named Name, matched
        exactly and read on their own, for the reasons Get-PreventNicAdvanced
        gives.
    #>
    param([string]$Name)
    $pattern = $Name -replace '[\[\]*?`]', '?'
    try {
        $rows = @(Get-NetAdapterPowerManagement -Name $pattern -ErrorAction Stop)
    } catch {
        if ($_.CategoryInfo.Category -eq [System.Management.Automation.ErrorCategory]::ObjectNotFound) { return }
        throw
    }
    $rows | Where-Object { [string]$_.Name -eq $Name }
}

function Get-PreventDeviceEnable {
    <#
    .SYNOPSIS
        Whether one MSPower_DeviceEnable or MSPower_DeviceWakeEnable
        instance reads on, read by itself.
    .DESCRIPTION
        The instance is read on its own, with Get-CimInstance -InputObject,
        which is how Device Manager reads it, and never taken from a read of
        the whole class. On an adapter whose driver is built on NetAdapterCx,
        such as Realtek's 2.5GbE rt25cx21, the driver framework answers a
        read of one instance, and it holds the setting; NDIS answers a read
        of the whole class, and it does not. Measured on this PC after 6.3.2
        had cleared "Allow the computer to turn off this device to save
        power" on the Realtek and Windows had restarted: the class read True,
        the instance False, and Device Manager showed the box cleared. Until
        6.3.3 the class read decided and checked this switch, so every run
        warned that the box "still reads on until Windows restarts", which a
        restart never changed, recorded it as on, and cleared it again.

        A read that fails, or that returns nothing, throws.
    #>
    param([Parameter(Mandatory)]$Instance)
    $one = @(Get-CimInstance -InputObject $Instance -ErrorAction Stop)
    if ($one.Count -eq 0) { throw "Windows returned nothing for $([string]$Instance.InstanceName)." }
    return [bool]$one[0].Enable
}

function Disable-NicPowerSaving {
    <#
    .SYNOPSIS
        Turns off power saving on every physical network adapter, recording
        each setting before it is changed, and reports the outcome to step 6.
    .DESCRIPTION
        An adapter's power saving lives in five places, and each is a
        different switch:

        1. "Allow the computer to turn off this device to save power", the
           checkbox on the adapter's Power Management tab. In WMI that is
           MSPower_DeviceEnable. Each instance is read by itself, as Device
           Manager reads it, to decide whether to clear it and to check that
           it took: on some drivers a read of the whole class gives another
           answer (see Get-PreventDeviceEnable).
        2. IdleInWorkingState in the WDF key under the device's Enum key,
           Enum\<device id>\Device Parameters\WDF. A driver written on the
           Windows Driver Framework reads it to decide whether it may power
           the device down while the PC is running and the device is
           idle. It is set to 0 where it is there as a number. Where it
           is absent the driver offers no such setting: nothing is
           written, and the key is never created.
        3. PnPCapabilities in the adapter's driver key, under
           Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}\NNNN. Bit
           0x08 tells NDIS not to power the adapter down. Until 6.3.0 this
           value was written to the adapter's Enum key instead, where nothing
           reads it, so an adapter reported as done still had power saving
           on. Only the 0x08 bit is added; the wake bits beside it are left
           as they are. The key is never created.
        4. The driver's own options, by registry keyword: Energy Efficient
           Ethernet, Green Ethernet and the rest of $NicPowerOff, each set
           to the value that turns it off, and the hidden ones in
           $NicHiddenPowerOff, each only from its on value.
        5. Selective suspend, sleep on disconnect and D0 packet coalescing,
           through Set-NetAdapterPowerManagement. Each is kept in a driver
           option, $NicPmKeyword; a driver that has none for a setting
           leaves the cmdlet nowhere to write it, so a setting that does not
           take there is not a failure: nothing is recorded, and it is a
           warning. Selective suspend there is not even that where step 2
           turned the framework's idle power-down off: that is the switch a
           driver built on the framework is powered down by while idle.

        Adapters that are not physical have the first switch too: hidden
        ones, such as the Wi-Fi Direct virtual adapters that a Wi-Fi driver
        adds, and visible virtual ones, such as a VPN's TAP adapter.
        -Physical lists none of them, and Get-NetAdapter lists the hidden
        ones only with -IncludeHidden, so until 6.3.0 none was read and each
        kept it on. They are read and handled whether or not the physical
        adapters could be listed: a physical list that cannot be read fails
        the step for those adapters, naming why, and no more. Among the
        adapters -IncludeHidden lists, a physical one is known by its device
        id from -Physical, or, when that list could not be read, by
        ConnectorPresent, which Windows sets on physical adapters alone
        (measured on this PC: true on the Ethernet and Wi-Fi adapters, false
        on every hidden and virtual one), and it is left out. Every one
        that -IncludeHidden lists with a device id is taken, hidden or
        not, so that no adapter is left for Windows to
        power down. It is cleared on each the same way and recorded under
        its device id; -Undo finds it by that id, whether the adapter is
        physical or not. Each one changed gets one line, which calls it a
        hidden adapter where its Hidden property is true and a virtual
        adapter otherwise; one with nothing to change is not shown.

        An adapter's options are read by a pattern that matches its name,
        and only those whose name is the same text are kept (see
        Get-PreventNicAdvanced). Each change is made on the object read back
        for it.

        Every change is made with -NoRestart. Restarting an adapter drops
        the network link: until 6.3.0 this step did that, for about 14
        seconds, and the clock check after it then failed. The changes take
        effect at the next Windows restart, and the summary says so.
    #>
    # A physical list that cannot be read fails the step for those adapters
    # and goes on: the hidden and virtual adapters below are a list of their
    # own.
    $physicalError = $null
    $nics = @()
    try {
        $nics = @(Get-NetAdapter -Physical -ErrorAction Stop)
    } catch {
        $physicalError = $_.Exception.Message
        Log "[!] Could not list the physical network adapters, so their power saving was not changed: $physicalError" -Colour Red -Indent
        Add-StepResult -Status Failed -Detail "Could not list the physical network adapters, so their power saving was not changed: $physicalError"
    }
    if ($null -eq $physicalError -and $nics.Count -eq 0) {
        Log "No physical network adapters found" -Colour DarkGray -Indent
        Add-StepResult -Status Skipped
    }

    $devEnable = @()
    $devEnableError = $null
    try {
        $devEnable = @(Get-CimInstance -Namespace root\wmi -ClassName MSPower_DeviceEnable -ErrorAction Stop)
    } catch {
        $devEnableError = $_.Exception.Message
    }

    foreach ($nic in $nics) {
        $label   = "$($nic.Name) ($($nic.InterfaceDescription))"
        $pnp     = [string]$nic.PnPDeviceID
        $changed = New-Object System.Collections.ArrayList
        $warned  = New-Object System.Collections.ArrayList
        $failed  = New-Object System.Collections.ArrayList
        $wrote   = $false
        Log $label -Colour Cyan -Indent

        try {
            # 1. Allow the computer to turn off this device to save power.
            #    The class read above names the adapter's instances; each is
            #    then read by itself, before it is cleared and after, since
            #    on some drivers the class read gives another answer (see
            #    Get-PreventDeviceEnable). One that cannot be read is not
            #    written.
            if ($null -ne $devEnableError) {
                $null = $warned.Add("could not read whether Windows may turn it off: $devEnableError")
            } else {
                $inst = @($devEnable | Where-Object { ([string]$_.InstanceName).StartsWith($pnp + '_', [StringComparison]::OrdinalIgnoreCase) })
                foreach ($in in $inst) {
                    try {
                        $on = Get-PreventDeviceEnable -Instance $in
                    } catch {
                        $null = $warned.Add("could not read whether Windows may turn it off: $($_.Exception.Message)")
                        continue
                    }
                    if (-not $on) { continue }
                    Add-PreventUndoItem -Kind NicDeviceEnable -Key $pnp -Original $true
                    try {
                        $in | Set-CimInstance -Property @{ Enable = $false } -ErrorAction Stop
                        $wrote = $true
                        if (-not (Get-PreventDeviceEnable -Instance $in)) {
                            $null = $changed.Add("Windows may no longer turn it off to save power")
                        } else {
                            $null = $warned.Add("'allow the computer to turn off this device' still reads on after it was cleared")
                        }
                    } catch {
                        $null = $failed.Add("'allow the computer to turn off this device': $($_.Exception.Message)")
                    }
                }
            }

            # 2. The driver framework's idle power-down. An absent value, or
            #    one that is not a number, is left alone, so the key is
            #    never created.
            $wdfKey = "HKLM:\SYSTEM\CurrentControlSet\Enum\$pnp\Device Parameters\WDF"
            $idle = Get-RegistryValue -Path $wdfKey -Name 'IdleInWorkingState'
            $idleOff = $false
            if ($idle -is [int]) {
                $res = Write-RecordedDword -Path $wdfKey -Name 'IdleInWorkingState' -Value 0
                if ($res -eq 'Done' -or $res -eq 'Unchanged') { $idleOff = $true }
                if ($res -eq 'Done') {
                    $wrote = $true
                    $null = $changed.Add("the driver framework may no longer power it down while the PC is running (IdleInWorkingState 0)")
                } elseif ($res -eq 'Failed') {
                    $null = $failed.Add("could not write IdleInWorkingState to $wdfKey")
                }
            }

            # 3. PnPCapabilities in the driver key
            $drv = $null
            try {
                $drv = [string](Get-PnpDeviceProperty -InstanceId $pnp -KeyName 'DEVPKEY_Device_Driver' -ErrorAction Stop).Data
            } catch {
                $drv = $null
            }
            if (-not $drv) {
                $null = $warned.Add("could not find its driver key, so PnPCapabilities was not set")
            } else {
                $classKey = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\$drv"
                if (-not (Test-Path -LiteralPath $classKey)) {
                    $null = $warned.Add("its driver key $drv is missing, so PnPCapabilities was not set")
                } else {
                    $cap = Get-RegistryValue -Path $classKey -Name 'PnPCapabilities'
                    $want = 8
                    if ($cap -is [int]) { $want = $cap -bor 8 }
                    $res = Write-RecordedDword -Path $classKey -Name 'PnPCapabilities' -Value $want
                    if ($res -eq 'Done') {
                        $wrote = $true
                        $null = $changed.Add("NDIS may no longer power it down (PnPCapabilities $want)")
                    } elseif ($res -eq 'Failed') {
                        $null = $failed.Add("could not write PnPCapabilities to $classKey")
                    }
                }
            }

            # 4. The driver's own power saving options
            $props = @()
            $propsRead = $false
            try {
                $props = @(Get-PreventNicAdvanced -Name $nic.Name)
                $propsRead = $true
            } catch {
                $null = $failed.Add("could not read the driver's options: $($_.Exception.Message)")
            }
            foreach ($p in $props) {
                $kw = [string]$p.RegistryKeyword
                $valid = ConvertTo-StringArray $p.ValidRegistryValues
                $cur = ConvertTo-StringArray $p.RegistryValue
                if ($valid.Count -eq 0 -and $NicHiddenPowerOff.Contains($kw)) {
                    # Only from its on value, 1: see $NicHiddenPowerOff.
                    if ($cur.Count -ne 1 -or $cur[0] -ne '1') { continue }
                    $off = [string]$NicHiddenPowerOff[$kw]
                } else {
                    if (-not $NicPowerOff.Contains($kw)) { continue }
                    $off = [string]$NicPowerOff[$kw]
                    if ($valid -notcontains $off) { continue }
                    $eff = Get-NicAdvancedValue -Property $p
                    if ($eff.Count -eq 1 -and $eff[0] -eq $off) { continue }
                }
                $orig = $null
                if ($cur.Count -gt 0) { $orig = $cur }
                $shown = [string]$p.DisplayName
                if (-not $shown) { $shown = $kw }
                Add-PreventUndoItem -Kind NicAdvanced -Key "$pnp|$kw" -Original $orig
                try {
                    # The option is given as -InputObject, never piped: piped,
                    # it also binds its own DisplayValue and RegistryValue by
                    # name, and a hidden option's empty DisplayValue fails the
                    # cmdlet before it writes. Until 6.3.2 that failed every
                    # hidden option, PowerDownPll and bFwCtrlLPS among them,
                    # with "Cannot validate argument on parameter
                    # 'DisplayValue'".
                    Set-NetAdapterAdvancedProperty -InputObject $p -RegistryValue $off -NoRestart -ErrorAction Stop
                    $wrote = $true
                    $chk = @(Get-PreventNicAdvanced -Name $nic.Name | Where-Object { $_.RegistryKeyword -eq $kw })
                    $now = @()
                    if ($chk.Count -gt 0) { $now = Get-NicAdvancedValue -Property $chk[0] }
                    if ($now.Count -eq 1 -and $now[0] -eq $off) {
                        $null = $changed.Add("$shown off")
                    } else {
                        $null = $failed.Add("$shown still reads $($now -join ',')")
                    }
                } catch {
                    $null = $failed.Add("${shown}: $($_.Exception.Message)")
                }
            }

            # 5. Selective suspend, sleep on disconnect, packet coalescing
            $pm = $null
            try {
                $pmAll = @(Get-PreventNicPower -Name $nic.Name)
                if ($pmAll.Count -gt 0) {
                    $pm = $pmAll[0]
                } else {
                    $null = $warned.Add("could not read its power management settings: Windows lists none for it")
                }
            } catch {
                $null = $warned.Add("could not read its power management settings: $($_.Exception.Message)")
            }
            if ($null -ne $pm) {
                foreach ($prop in @('SelectiveSuspend', 'DeviceSleepOnDisconnect', 'D0PacketCoalescing')) {
                    if (-not $pm.PSObject.Properties[$prop]) { continue }
                    if ([string]$pm.$prop -ne 'Enabled') { continue }
                    Add-PreventUndoItem -Kind NicPowerManagement -Key "$pnp|$prop" -Original 'Enabled'
                    try {
                        $splat = @{ NoRestart = $true; ErrorAction = 'Stop' }
                        $splat[$prop] = 'Disabled'
                        Set-NetAdapterPowerManagement -InputObject $pm @splat
                        $pmNow = ''
                        $pm2 = @(Get-PreventNicPower -Name $nic.Name)
                        if ($pm2.Count -gt 0) { $pmNow = [string]$pm2[0].$prop }
                        $pmKw = [string]$NicPmKeyword[$prop]
                        if ($pmNow -eq 'Disabled') {
                            $wrote = $true
                            $null = $changed.Add("$prop off")
                        } elseif ($propsRead -and @($props | Where-Object { [string]$_.RegistryKeyword -eq $pmKw }).Count -eq 0) {
                            # The driver has no option for it, so nothing was
                            # written: there is nothing to put back, and a
                            # record an earlier run made of it goes too. For
                            # selective suspend that is not a failure where the
                            # driver framework's idle power-down is off (step 2
                            # above): a driver built on the framework is
                            # powered down while idle by that alone, and
                            # Windows' reading of selective suspend does not
                            # show it. Anywhere else it is said.
                            Clear-PreventUndoItem -Kind NicPowerManagement -Key "$pnp|$prop"
                            if ($prop -ne 'SelectiveSuspend' -or -not $idleOff) {
                                $null = $warned.Add("$prop reads $pmNow, and the driver has no $pmKw option for Windows to turn it off with")
                            }
                        } else {
                            $wrote = $true
                            $null = $failed.Add("$prop still reads $pmNow")
                        }
                    } catch {
                        $null = $failed.Add("${prop}: $($_.Exception.Message)")
                    }
                }
            }
        } catch {
            $null = $failed.Add($_.Exception.Message)
        }

        foreach ($c in $changed) { Log "+ $c" -Colour Green -Indent }
        foreach ($w in $warned)  { Log "! $w" -Colour Yellow -Indent }
        foreach ($f in $failed)  { Log "[!] $f" -Colour Red -Indent }
        if ($failed.Count -gt 0) {
            Add-StepResult -Status Failed -Detail "${label}: $($failed -join '; ')"
        }
        if ($warned.Count -gt 0) {
            Add-StepResult -Status Warning -Detail "${label}: $($warned -join '; ')" -Reboot:$wrote
        }
        if ($changed.Count -gt 0) {
            Add-StepResult -Status Done -Detail "${label}: $($changed.Count) power saving setting(s) turned off" -Reboot
        }
        if ($changed.Count -eq 0 -and $warned.Count -eq 0 -and $failed.Count -eq 0) {
            Log "Power saving already off" -Colour DarkGray -Indent
            Add-StepResult -Status Unchanged -Detail "${label}: power saving already off"
        }
    }

    # Adapters that are not physical: the checkbox alone, one line for each
    # one changed. -IncludeHidden lists the visible virtual adapters, such
    # as a VPN's TAP adapter, beside the hidden ones, and both are taken. A
    # physical adapter is not taken again, nor one without a device id,
    # such as the Teredo tunnel.
    $listed = @()
    try {
        $listed = @(Get-NetAdapter -IncludeHidden -ErrorAction Stop)
    } catch {
        Log "! Could not list the hidden and virtual network adapters: $($_.Exception.Message)" -Colour Yellow -Indent
        Add-StepResult -Status Warning -Detail "Could not list the hidden and virtual network adapters: $($_.Exception.Message)"
        return
    }
    $physicalIds = @($nics | ForEach-Object { [string]$_.PnPDeviceID })
    $hiddenIds = @{}
    $hidden = @()
    foreach ($h in $listed) {
        $hPnp = [string]$h.PnPDeviceID
        if (-not $hPnp -or $physicalIds -contains $hPnp -or $hiddenIds.ContainsKey($hPnp)) { continue }
        # A physical adapter -Physical did not list, because that list could
        # not be read, is left out too: it is neither hidden nor virtual. It
        # is read only where the adapter has it, as Hidden is below.
        $hConnector = $h.PSObject.Properties['ConnectorPresent']
        if ($null -ne $hConnector -and $hConnector.Value -eq $true) { continue }
        $hiddenIds[$hPnp] = $true
        $hidden += $h
    }
    if ($hidden.Count -gt 0 -and $null -ne $devEnableError) {
        Log "! Hidden and virtual network adapters: could not read whether Windows may turn them off: $devEnableError" -Colour Yellow -Indent
        Add-StepResult -Status Warning -Detail "Hidden and virtual network adapters: could not read whether Windows may turn them off: $devEnableError"
        return
    }
    foreach ($h in $hidden) {
        $hPnp = [string]$h.PnPDeviceID
        $hInst = @($devEnable | Where-Object { ([string]$_.InstanceName).StartsWith($hPnp + '_', [StringComparison]::OrdinalIgnoreCase) })
        if ($hInst.Count -eq 0) { continue }
        # Hidden is read only where the adapter has it: one without it counts
        # as not hidden, and under StrictMode reading a property an object
        # lacks throws.
        $hHidden = $h.PSObject.Properties['Hidden']
        $hKind = 'a virtual adapter'
        if ($null -ne $hHidden -and $hHidden.Value -eq $true) { $hKind = 'a hidden adapter' }
        $hLabel = "$($h.Name) ($($h.InterfaceDescription)), $hKind"
        # Each instance is read by itself, as on the physical adapters (see
        # Get-PreventDeviceEnable). One that cannot be read is said and not
        # written; the others are still taken.
        $hOn = @()
        $hReadError = $null
        foreach ($hi in $hInst) {
            try {
                if (Get-PreventDeviceEnable -Instance $hi) { $hOn += $hi }
            } catch {
                $hReadError = $_.Exception.Message
            }
        }
        if ($null -ne $hReadError) {
            Log "! ${hLabel}: could not read whether Windows may turn it off: $hReadError" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail "${hLabel}: could not read whether Windows may turn it off: $hReadError"
        }
        if ($hOn.Count -eq 0) { continue }
        try {
            Add-PreventUndoItem -Kind NicDeviceEnable -Key $hPnp -Original $true
            $hOn | Set-CimInstance -Property @{ Enable = $false } -ErrorAction Stop
            $hStill = @($hOn | Where-Object { Get-PreventDeviceEnable -Instance $_ })
            if ($hStill.Count -eq 0) {
                Log "+ ${hLabel}: Windows may no longer turn it off to save power" -Colour Green -Indent
                Add-StepResult -Status Done -Detail "${hLabel}: Windows may no longer turn it off to save power" -Reboot
            } else {
                Log "! ${hLabel}: 'allow the computer to turn off this device' still reads on after it was cleared" -Colour Yellow -Indent
                Add-StepResult -Status Warning -Detail "${hLabel}: 'allow the computer to turn off this device' still reads on after it was cleared" -Reboot
            }
        } catch {
            Log "[!] ${hLabel}: 'allow the computer to turn off this device': $($_.Exception.Message)" -Colour Red -Indent
            Add-StepResult -Status Failed -Detail "${hLabel}: 'allow the computer to turn off this device': $($_.Exception.Message)"
        }
    }
}

function Get-PreventNic {
    <# The physical adapter with this PnP device id, or $null. #>
    param([string]$PnpId)
    foreach ($n in @(Get-NetAdapter -Physical -ErrorAction Stop)) {
        if ([string]$n.PnPDeviceID -eq $PnpId) { return $n }
    }
    return $null
}

function Restore-PreventUndoItem {
    <#
    .SYNOPSIS
        Puts back one recorded setting. Returns $true when its record can go:
        the setting is back, or what it belonged to no longer exists. $false
        keeps the record, so the next -Undo tries again.
    #>
    param($Item, [string]$CreatedPlan)

    $key  = [string]$Item.Key
    $orig = $Item.Original
    switch ([string]$Item.Kind) {

        'PowerValue' {
            $parts = $key.Split('|')
            if ($parts.Count -ne 3) { throw "The record $key is not a power setting." }
            $plan = $parts[0]; $sub = $parts[1]; $setting = $parts[2]
            # A plan this script created is deleted below.
            if ($CreatedPlan -and $plan -eq $CreatedPlan) { return $true }
            $label = $setting
            foreach ($o in $PowerOptions) { if ($o.Setting -eq $setting) { $label = $o.Label; break } }
            # A plan that is gone has nothing to put back. When powercfg
            # cannot list the plans that is not known, and the record stays.
            $exists = Test-PowerPlanExists -Plan $plan
            if ($null -eq $exists) {
                Log "[!] ${label}: powercfg could not list the power plans, so it stays recorded" -Colour Red -Indent
                Add-StepResult -Status Failed -Detail "${label}: powercfg could not list the power plans"
                return $false
            }
            if (-not $exists) { return $true }
            $want = [long]$orig
            $r = Invoke-PowerCfg -Arguments @('/setacvalueindex', $plan, $sub, $setting, [string]$want)
            $now = Get-AcPowerIndex -Plan $plan -Sub $sub -Setting $setting
            if ($null -eq $now) {
                # The plan is there and the setting is not: it is not coming
                # back, and a record of it would keep every -Undo from
                # finishing.
                if (Test-PowerSettingGone -Sub $sub -Setting $setting) {
                    Log "[!] ${label}: this setting is no longer on this machine, so its record is dropped" -Colour Yellow -Indent
                    Add-StepResult -Status Warning -Detail "${label}: the setting is no longer on this machine, so its record is dropped"
                    return $true
                }
                # Not read back, so not known to be back: the record stays.
                Log "[!] ${label}: could not be read back, so it stays recorded ($($r.Output))" -Colour Red -Indent
                Add-StepResult -Status Failed -Detail "${label} could not be read back after putting it back"
                return $false
            }
            if ($now -eq $want) {
                Log "${label}: put back" -Colour Green -Indent
                Add-StepResult -Status Done
                return $true
            }
            Log "[!] ${label}: could not be put back ($($r.Output))" -Colour Red -Indent
            Add-StepResult -Status Failed -Detail "${label} could not be put back"
            return $false
        }

        'Hibernate' {
            $r = Invoke-PowerCfg -Arguments @('/hibernate', 'on')
            if ((Get-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Power' -Name 'HibernateEnabled') -eq 1) {
                Log "Hibernate: on" -Colour Green -Indent
                Add-StepResult -Status Done
                return $true
            }
            Log "[!] Could not turn hibernate back on: $($r.Output)" -Colour Red -Indent
            Add-StepResult -Status Failed -Detail "Could not turn hibernate back on: $($r.Output)"
            return $false
        }

        'Registry' {
            $i = $key.LastIndexOf('|')
            if ($i -lt 1) { throw "The record $key is not a registry value." }
            $path = $key.Substring(0, $i)
            $name = $key.Substring($i + 1)
            if (-not (Test-Path -LiteralPath $path)) { return $true }
            if ($null -eq $orig) {
                if ($null -ne (Get-RegistryValue -Path $path -Name $name)) {
                    Remove-ItemProperty -LiteralPath $path -Name $name -ErrorAction Stop
                }
                Log "${name}: removed (there was none) under $path" -Colour Green -Indent
            } else {
                Set-ItemProperty -LiteralPath $path -Name $name -Value ([int]$orig) -Type DWord -ErrorAction Stop
                Log "${name}: put back to $orig under $path" -Colour Green -Indent
            }
            Add-StepResult -Status Done -Reboot
            return $true
        }

        'RegistryKey' {
            if (-not (Test-Path -LiteralPath $key)) { return $true }
            # A value this script set in the key is still recorded, because
            # putting it back failed: the key keeps its record until that
            # value is gone. Dropped now, the key this script created stayed
            # behind once the next -Undo removed the value.
            foreach ($other in @($script:UndoRecord.ToArray())) {
                if ([string]$other.Kind -ne 'Registry') { continue }
                $vk = [string]$other.Key
                $at = $vk.LastIndexOf('|')
                if ($at -lt 1) { continue }
                if ([string]::Equals($vk.Substring(0, $at).TrimEnd('\'), $key.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) {
                    Log "Kept the record of ${key}: a value set in it is still recorded" -Colour DarkGray -Indent
                    return $false
                }
            }
            $k = Get-Item -LiteralPath $key -ErrorAction Stop
            if ($k.ValueCount -eq 0 -and $k.SubKeyCount -eq 0) {
                Remove-Item -LiteralPath $key -ErrorAction Stop
                Log "Removed the registry key this script created: $key" -Colour Green -Indent
                Add-StepResult -Status Done
            } else {
                Log "Left $key in place: it now holds settings this script did not make" -Colour DarkGray -Indent
                Add-StepResult -Status Unchanged -Detail "Left $key in place: it holds settings this script did not make"
            }
            return $true
        }

        'NicDeviceEnable' {
            $inst = @(Get-CimInstance -Namespace root\wmi -ClassName MSPower_DeviceEnable -ErrorAction Stop |
                      Where-Object { ([string]$_.InstanceName).StartsWith($key + '_', [StringComparison]::OrdinalIgnoreCase) })
            if ($inst.Count -eq 0) {
                Log "[!] The network adapter $key is not present; its setting stays recorded" -Colour Yellow -Indent
                Add-StepResult -Status Warning -Detail "Adapter $key not present: rerun -Undo with it connected"
                return $false
            }
            $inst | Set-CimInstance -Property @{ Enable = [bool]$orig } -ErrorAction Stop
            Log "Windows may turn the adapter off to save power again ($key)" -Colour Green -Indent
            Add-StepResult -Status Done -Reboot
            return $true
        }

        'NicAdvanced' {
            $i = $key.LastIndexOf('|')
            if ($i -lt 1) { throw "The record $key is not an adapter option." }
            $pnp = $key.Substring(0, $i)
            $kw  = $key.Substring($i + 1)
            $nic = Get-PreventNic -PnpId $pnp
            if ($null -eq $nic) {
                Log "[!] The network adapter $pnp is not present; its $kw stays recorded" -Colour Yellow -Indent
                Add-StepResult -Status Warning -Detail "Adapter $pnp not present: rerun -Undo with it connected"
                return $false
            }
            $p = @(Get-PreventNicAdvanced -Name $nic.Name | Where-Object { $_.RegistryKeyword -eq $kw })
            if ($p.Count -eq 0) {
                # The driver no longer has the option, as after a driver
                # update: there is nothing to put it back on, and a record of
                # it would keep every -Undo from finishing. It is said, as a
                # power setting that is gone is, so that the value is not
                # taken for put back.
                Log "[!] $($nic.Name): its driver no longer lists the option $kw, so it was not put back and its record is dropped" -Colour Yellow -Indent
                Add-StepResult -Status Warning -Detail "$($nic.Name): its driver no longer lists the option $kw, so it was not put back and its record is dropped"
                return $true
            }
            $want = ConvertTo-StringArray $orig
            if ($want.Count -eq 0) {
                Reset-NetAdapterAdvancedProperty -InputObject $p[0] -NoRestart -ErrorAction Stop
            } else {
                # -InputObject, not piped: see Disable-NicPowerSaving.
                Set-NetAdapterAdvancedProperty -InputObject $p[0] -RegistryValue $want -NoRestart -ErrorAction Stop
            }
            Log "$($nic.Name): $kw put back" -Colour Green -Indent
            Add-StepResult -Status Done -Reboot
            return $true
        }

        'NicPowerManagement' {
            $i = $key.LastIndexOf('|')
            if ($i -lt 1) { throw "The record $key is not an adapter setting." }
            $pnp  = $key.Substring(0, $i)
            $prop = $key.Substring($i + 1)
            if (@('SelectiveSuspend', 'DeviceSleepOnDisconnect', 'D0PacketCoalescing') -notcontains $prop) {
                throw "The record $key names an adapter setting this script does not change."
            }
            $nic = Get-PreventNic -PnpId $pnp
            if ($null -eq $nic) {
                Log "[!] The network adapter $pnp is not present; its $prop stays recorded" -Colour Yellow -Indent
                Add-StepResult -Status Warning -Detail "Adapter $pnp not present: rerun -Undo with it connected"
                return $false
            }
            $pmNow = @(Get-PreventNicPower -Name $nic.Name)
            if ($pmNow.Count -eq 0) { throw "Windows lists no power management settings for $($nic.Name)" }
            $splat = @{ NoRestart = $true; ErrorAction = 'Stop' }
            $splat[$prop] = [string]$orig
            Set-NetAdapterPowerManagement -InputObject $pmNow[0] @splat
            Log "$($nic.Name): $prop put back to $orig" -Colour Green -Indent
            Add-StepResult -Status Done -Reboot
            return $true
        }

        'DefenderPath' {
            Remove-MpPreference -ExclusionPath $key -ErrorAction Stop
            Log "Removed the Defender exclusion this script added: $key" -Colour Green -Indent
            Add-StepResult -Status Done
            return $true
        }

        'DefenderProcess' {
            Remove-MpPreference -ExclusionProcess $key -ErrorAction Stop
            Log "Removed the Defender process exclusion this script added: $key" -Colour Green -Indent
            Add-StepResult -Status Done
            return $true
        }

        'ServiceFailureActions' {
            # A service that is gone has no recovery to put back. Until 6.3.1
            # one that could not be read counted as gone, and its record
            # was dropped.
            try {
                $svc = Get-ClaudeServiceOrNull -Name $key
            } catch {
                throw "the service could not be read, so its recovery settings stay recorded ($(Get-ClaudeErrorText $_))"
            }
            if ($null -eq $svc) { return $true }
            $svcKey = "HKLM:\SYSTEM\CurrentControlSet\Services\$key"
            $reboot = $false
            if ($null -eq $orig) {
                # There was no recovery set: clear the actions, then remove the
                # value sc.exe leaves behind, so the service is as it was.
                $out = & sc.exe failure $key reset= 0 actions= '""' 2>&1
                $code = $LASTEXITCODE
                if ($code -ne 0) { throw "sc.exe failure returned ${code}: $(@($out) -join ' ')" }
                if ($null -ne (Get-ServiceFailureBytes -Service $key)) {
                    Remove-ItemProperty -LiteralPath $svcKey -Name 'FailureActions' -ErrorAction Stop
                }
                $ok = ($null -eq (Get-ServiceFailureBytes -Service $key))
            } else {
                $bytes = [Convert]::FromBase64String([string]$orig)
                $want = ConvertFrom-FailureActionsBlob -Bytes $bytes
                if ($null -eq $want) { throw "The recorded recovery settings for $key cannot be read." }
                $simple = $true
                foreach ($a in @($want.Actions)) { if ($a.Type -ne 1 -and $a.Type -ne 2) { $simple = $false } }
                if ($simple) {
                    $words = @()
                    foreach ($a in @($want.Actions)) {
                        $verb = 'restart'
                        if ($a.Type -eq 2) { $verb = 'reboot' }
                        $words += "$verb/$($a.Delay)"
                    }
                    $actions = '""'
                    if ($words.Count -gt 0) { $actions = $words -join '/' }
                    $out = & sc.exe failure $key reset= ([string]$want.Reset) actions= $actions 2>&1
                    $code = $LASTEXITCODE
                    if ($code -ne 0) { throw "sc.exe failure returned ${code}: $(@($out) -join ' ')" }
                } else {
                    # An action sc.exe cannot write back as it was (run a
                    # command, or none between others): the value itself is
                    # restored, and the service manager reads it at the next
                    # Windows restart.
                    Set-ItemProperty -LiteralPath $svcKey -Name 'FailureActions' -Value $bytes -Type Binary -ErrorAction Stop
                    $reboot = $true
                }
                $now = ConvertFrom-FailureActionsBlob -Bytes (Get-ServiceFailureBytes -Service $key)
                $ok = ($null -ne $now -and $now.Reset -eq $want.Reset -and @($now.Actions).Count -eq @($want.Actions).Count)
                if ($ok) {
                    $na = @($now.Actions); $wa = @($want.Actions)
                    for ($j = 0; $j -lt $wa.Count; $j++) {
                        if ($na[$j].Type -ne $wa[$j].Type -or $na[$j].Delay -ne $wa[$j].Delay) { $ok = $false }
                    }
                }
            }
            if ($ok) {
                Log "${key}: recovery settings put back" -Colour Green -Indent
                Add-StepResult -Status Done -Reboot:$reboot
                return $true
            }
            Log "[!] ${key}: recovery settings could not be put back" -Colour Red -Indent
            Add-StepResult -Status Failed -Detail "${key}: recovery settings could not be put back"
            return $false
        }

        default {
            throw "The undo record holds a setting of a kind this version does not know: $($Item.Kind)"
        }
    }
}

# -- Header ----------------------------------------------------------
Write-Host ""
Write-Host "  +----------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |  CLAUDE DESKTOP / COWORK, PREVENTION SETUP   |" -ForegroundColor Cyan
Write-Host "  |  v$ToolkitVersion                                      |" -ForegroundColor DarkGray
Write-Host "  +----------------------------------------------+" -ForegroundColor Cyan
Write-Host ""

if ($Undo) {
    Write-Host "  MODE: UNDO, putting back what this script changed" -ForegroundColor Yellow
    Write-Host ""
}
# Each read the discovery could not make, said before the first step.
Add-DiscoveryReadError -ReadErrors @($script:ClaudeEnv.ReadErrors)

try {
    # Before anything changes: a record that cannot be read stops the run
    # here, rather than being overwritten with values this script set.
    Read-PreventUndo

if ($Undo) {

    # ================================================================
    # UNDO MODE
    # ================================================================
    # Until 6.3.0 -Undo set fixed values, whatever the machine had before:
    # hibernate and Fast Startup on, sleep after 30 minutes, every Ultimate
    # Performance plan deleted (a plan you made yourself included), and the
    # power plan it had recorded, which a second run had overwritten with its
    # own. Now step 1 puts back what the undo record holds, and step 2 removes
    # only what older versions are known to have set.
    $steps = 5

    # ----------------------------------------------------------------
    # 1. Everything recorded by 6.3.0 or later
    # ----------------------------------------------------------------
    Step 1 $steps "Putting back the recorded settings..." -Name "Recorded settings"
    $planRestored = $false
    if ($script:UndoForeign) {
        Log "[!] The undo record $UndoFile was written on another computer ($($script:UndoForeign)), so nothing was put back from it" -Colour Red -Indent
        if ($script:UndoForeignByGuid) {
            # The MachineGuid decided, and a new name does not change it.
            Log "It holds another Windows installation's MachineGuid, which renaming this computer does not change" -Colour Yellow -Indent
        } else {
            Log "If this is that computer under a new name, set ComputerName in the file to $([Environment]::MachineName) and run -Undo again" -Colour Yellow -Indent
        }
        Add-StepResult -Status Failed -Detail "The undo record was written on another computer ($($script:UndoForeign)): nothing was put back from it"
    } elseif ($script:UndoRecord.Count -eq 0) {
        Log "Nothing recorded: no run of version 6.3.0 or later changed a setting here" -Colour DarkGray -Indent
        Add-StepResult -Status Unchanged
    } else {
        $createdPlan = $null
        $createdItem = Get-PreventUndoItem -Kind PowerPlan -Key Created
        if ($createdItem -and $createdItem.Original) { $createdPlan = ([string]$createdItem.Original).ToLowerInvariant() }

        # Newest first, and the registry keys this script created after all
        # the values: a key is removed only once no value set in it is still
        # recorded, and a value recorded by an earlier run than its key would
        # otherwise come after it.
        $items = @($script:UndoRecord.ToArray())
        [array]::Reverse($items)
        $items = @($items | Where-Object { $_.Kind -ne 'RegistryKey' }) + @($items | Where-Object { $_.Kind -eq 'RegistryKey' })
        foreach ($it in $items) {
            # The plans are put back below, and an older version's install
            # is step 2's to finish with.
            if ($it.Kind -eq 'PowerPlan' -or $it.Kind -eq 'LegacyInstall') { continue }
            try {
                if (Restore-PreventUndoItem -Item $it -CreatedPlan $createdPlan) {
                    Clear-PreventUndoItem -Kind $it.Kind -Key $it.Key
                }
            } catch {
                Log "[!] Could not put back $($it.Kind) $($it.Key): $($_.Exception.Message)" -Colour Red -Indent
                Add-StepResult -Status Failed -Detail "$($it.Kind) $($it.Key): $($_.Exception.Message)"
            }
        }

        # The plan you had. Activating it also applies any of its settings
        # put back above. This and the created plan below each have their
        # own try, as the items above do: a record that cannot be saved here
        # used to stop -Undo, and steps 2 to 5 never ran.
        $activeItem = Get-PreventUndoItem -Kind PowerPlan -Key Active
        try {
            $current = Get-ActivePlanGuid
            if ($activeItem) {
                $target = ([string]$activeItem.Original).ToLowerInvariant()
                $exists = $false
                if ($target -match "^$GuidPattern$" -and $target -ne $createdPlan) {
                    $exists = Test-PowerPlanExists -Plan $target
                    if ($null -eq $exists) { throw "powercfg could not list the power plans, so the plan you had stays recorded" }
                }
                if (-not $exists) {
                    Log "[!] The plan you had ($target) no longer exists, so Balanced is made active" -Colour Yellow -Indent
                    Add-StepResult -Status Warning -Detail "The plan you had ($target) is gone, so Balanced is active now"
                    $target = $BalancedGuid
                }
                if ($current -eq $target) {
                    $null = Invoke-PowerCfg -Arguments @('/setactive', $target)
                    Log "Power plan: $(Get-PowerPlanName -Plan $target) (already active)" -Colour DarkGray -Indent
                    Clear-PreventUndoItem -Kind PowerPlan -Key Active
                    $planRestored = $true
                } else {
                    $r = Invoke-PowerCfg -Arguments @('/setactive', $target)
                    if ((Get-ActivePlanGuid) -eq $target) {
                        Log "Power plan put back: $(Get-PowerPlanName -Plan $target)" -Colour Green -Indent
                        Add-StepResult -Status Done
                        Clear-PreventUndoItem -Kind PowerPlan -Key Active
                        $planRestored = $true
                    } else {
                        Log "[!] Could not make $target the active plan: $($r.Output)" -Colour Red -Indent
                        Add-StepResult -Status Failed -Detail "powercfg could not make $target the active plan"
                    }
                }
            } elseif ($current) {
                # Settings put back on the active plan take effect when it is
                # activated again.
                $null = Invoke-PowerCfg -Arguments @('/setactive', $current)
            }
        } catch {
            Log "[!] Power plan: $($_.Exception.Message)" -Colour Red -Indent
            Add-StepResult -Status Failed -Detail "Power plan: $($_.Exception.Message)"
        }

        # The Ultimate Performance plan this script created, and only that.
        if ($createdItem) {
            try {
                $createdExists = $false
                if ($createdPlan) {
                    $createdExists = Test-PowerPlanExists -Plan $createdPlan
                    if ($null -eq $createdExists) { throw "powercfg could not list the power plans, so the plan this script created ($createdPlan) stays recorded" }
                }
                if (-not $createdExists) {
                    Clear-PreventUndoItem -Kind PowerPlan -Key Created
                } elseif ((Get-ActivePlanGuid) -eq $createdPlan) {
                    Log "[!] The plan this script created ($createdPlan) is still active, so it was kept" -Colour Yellow -Indent
                    Add-StepResult -Status Warning -Detail "Kept the plan this script created ($createdPlan): it is still the active plan"
                } else {
                    $r = Invoke-PowerCfg -Arguments @('/delete', $createdPlan)
                    # Gone only when powercfg lists the plans and not this one.
                    if ($false -eq (Test-PowerPlanExists -Plan $createdPlan)) {
                        Log "Removed the Ultimate Performance plan this script created ($createdPlan)" -Colour Green -Indent
                        Add-StepResult -Status Done
                        Clear-PreventUndoItem -Kind PowerPlan -Key Created
                    } else {
                        Log "[!] Could not remove the plan $createdPlan : $($r.Output)" -Colour Red -Indent
                        Add-StepResult -Status Failed -Detail "Could not remove the plan this script created ($createdPlan)"
                    }
                }
            } catch {
                Log "[!] Power plan: $($_.Exception.Message)" -Colour Red -Indent
                Add-StepResult -Status Failed -Detail "Power plan: $($_.Exception.Message)"
            }
        }
        Complete-Step
    }

    # ----------------------------------------------------------------
    # 2. What versions before 6.3.0 changed without a record
    # ----------------------------------------------------------------
    Step 2 $steps "Removing what older versions changed..." -Name "Older versions' changes"
    $legacyCount = 0
    $unseenCount = 0
    $gatedLeft   = 0
    # What could not be read. Then it is not known that nothing from older
    # versions is here, and the step does not say so.
    $unreadCount = 0
    try {
        # Whether an older version ran here, decided before this step deletes
        # the files that show it, and recorded so that a -Undo which cannot
        # finish here still knows it next time.
        # One that cannot be told is said, and what is removed only where an
        # older version ran is left. Until 6.3.1 it read as no older version.
        $legacyTrace = $null
        $legacyUnread = $false
        try {
            $legacyTrace = Get-LegacyInstallTrace
        } catch {
            $legacyUnread = $true
            $unreadCount++
            $script:UndoUnread = $true
            $said = "Whether a version before 6.3.0 ran here could not be told, because $(Get-ClaudeErrorText $_), so what -Undo removes only where one ran is left as it is"
            Log "[!] $said" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail $said
        }
        $legacySeen = $null -ne $legacyTrace
        if ($legacySeen -and -not $script:UndoForeign -and $null -eq (Get-PreventUndoItem -Kind LegacyInstall -Key $legacyTrace.What)) {
            try {
                Add-LegacyInstallRecord -What $legacyTrace.What -Where $legacyTrace.Where
            } catch {
                Log "[!] $($_.Exception.Message)" -Colour Yellow -Indent
                Add-StepResult -Status Warning -Detail "Could not record that an older version ran here: $($_.Exception.Message)"
            }
        }

        # The power plan. Versions before 6.3.0 kept the plan they replaced in
        # $BackupFile. A 6.3.0 setup carries it over to the undo record and
        # deletes it, so it is here only when no 6.3.0 setup has run.
        # Until 6.3.1 each file below was looked for with Test-Path, which
        # answers $false for a file that cannot be read as for one that is
        # not there (measured 2026-11-23): one that could not be read was
        # not counted, and "Nothing from older versions found" was said.
        $there = Test-PreventPathThere -Path $BackupFile -NotDone 'the plan it holds was not put back, and it was not removed'
        if ($null -eq $there) { $unreadCount++ }
        if ($there) {
            $legacyCount++
            if ($planRestored) {
                Remove-PreventItem -Path $BackupFile
            } else {
                # A record or a plan list that could not be read keeps the
                # record for the next -Undo. Until 6.3.1 a record that could
                # not be read was taken for one that held no plan: Balanced
                # was made active and the record deleted.
                $legacyPlan = $null
                $legacyUnreadWhy = $null
                try {
                    $legacyPlan = ([string](Get-Content -LiteralPath $BackupFile -Raw -ErrorAction Stop)).Trim().ToLowerInvariant()
                } catch {
                    $legacyUnreadWhy = "$BackupFile could not be read ($(Get-ClaudeErrorText $_))"
                }
                $legacyExists = $false
                $ultimateNow = @()
                if ($null -eq $legacyUnreadWhy -and $legacyPlan -match "^$GuidPattern$") {
                    $legacyExists = Test-PowerPlanExists -Plan $legacyPlan
                    if ($legacyExists) {
                        try {
                            $ultimateNow = @(Get-UltimatePlanGuids)
                        } catch {
                            $legacyUnreadWhy = $_.Exception.Message
                        }
                    }
                }
                if ($null -ne $legacyUnreadWhy) {
                    $said = "The plan an older version recorded was not put back, because $legacyUnreadWhy; its record is kept for the next -Undo"
                    Log "[!] $said" -Colour Red -Indent
                    Add-StepResult -Status Failed -Detail $said
                } elseif ($null -eq $legacyExists) {
                    # Not known whether that plan is there: the older
                    # version's record stays for the next -Undo.
                    Log "[!] powercfg could not list the power plans, so the plan an older version recorded was not put back" -Colour Red -Indent
                    Add-StepResult -Status Failed -Detail "powercfg could not list the power plans: the older version's record of your plan is kept for the next -Undo"
                } else {
                    $target = $BalancedGuid
                    if ($legacyExists -and ($ultimateNow -notcontains $legacyPlan)) {
                        $target = $legacyPlan
                    } else {
                        Log "[!] The older version's record does not hold the plan you had, so Balanced is made active" -Colour Yellow -Indent
                        Add-StepResult -Status Warning -Detail "Balanced is active now: the older version's record of your plan was overwritten with its own"
                    }
                    $r = Invoke-PowerCfg -Arguments @('/setactive', $target)
                    if ((Get-ActivePlanGuid) -eq $target) {
                        Log "Power plan put back: $(Get-PowerPlanName -Plan $target)" -Colour Green -Indent
                        Add-StepResult -Status Done
                        Remove-PreventItem -Path $BackupFile
                    } else {
                        Log "[!] Could not make $target the active plan: $($r.Output)" -Colour Red -Indent
                        Add-StepResult -Status Failed -Detail "powercfg could not make $target the active plan"
                    }
                }
            }
        }

        # Network adapters. Versions before 6.3.0 wrote PnPCapabilities = 24
        # to the adapter's Enum key, where nothing reads it, and turned off its
        # wake settings: "Wake on Magic Packet", "Wake on Pattern Match" and
        # Windows' "allow this device to wake the computer". The 24 marks an
        # adapter they touched, and on those the wake settings go back on.
        # A list that cannot be read is said, as the NAT rules below are, and
        # -Undo ends incomplete, so that it is run again to look.
        $legacyNics = @()
        try {
            $legacyNics = @(Get-NetAdapter -Physical -ErrorAction Stop)
        } catch {
            $unreadCount++
            $script:UndoUnread = $true
            Log "[!] Could not list the network adapters, so the wake settings older versions turned off were not checked: $($_.Exception.Message)" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail "Could not list the network adapters, so the wake settings older versions turned off on them were not checked: $($_.Exception.Message)"
        }
        $wakeInst = $null
        foreach ($nic in $legacyNics) {
            $pnp = [string]$nic.PnPDeviceID
            $devPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\$pnp\Device Parameters"
            if ((Get-RegistryValue -Path $devPath -Name 'PnPCapabilities') -ne 24) { continue }
            $legacyCount++
            $label = "$($nic.Name) ($($nic.InterfaceDescription))"
            try {
                # Wake first and the marker last: the 24 is what shows an
                # older version touched this adapter, so it goes only once
                # wake is back on, and a wake setting that cannot be put back
                # leaves it for the next -Undo.
                $advAll = @(Get-PreventNicAdvanced -Name $nic.Name)
                foreach ($dn in @('Wake on Magic Packet', 'Wake on Pattern Match')) {
                    foreach ($wp in @($advAll | Where-Object { [string]$_.DisplayName -eq $dn })) {
                        if ([string]$wp.DisplayValue -eq 'Disabled') {
                            # -InputObject, not piped: piped, the option's own
                            # RegistryValue, the off one, is bound beside the
                            # DisplayValue asked for.
                            Set-NetAdapterAdvancedProperty -InputObject $wp -DisplayValue 'Enabled' -NoRestart -ErrorAction Stop
                        }
                    }
                }
                if ($null -eq $wakeInst) {
                    # root\wmi answers Not supported for a class no device
                    # has: then no device can wake the computer, and there
                    # is nothing to turn back on. Any other failure fails
                    # this adapter and keeps its 24 for the next -Undo: taken
                    # for no device, it let the 24 go with this setting off.
                    try {
                        $wakeInst = @(Get-CimInstance -Namespace root\wmi -ClassName MSPower_DeviceWakeEnable -ErrorAction Stop)
                    } catch {
                        $wakeEx = $_.Exception
                        if (-not ($wakeEx.GetType().FullName -eq 'Microsoft.Management.Infrastructure.CimException' -and [int]$wakeEx.StatusCode -eq 7)) { throw }
                        $wakeInst = @()
                    }
                }
                # Each instance is read by itself, as Device Manager reads
                # it: the class read names them, and on some drivers gives
                # another answer for the setting (see Get-PreventDeviceEnable).
                foreach ($wi in @($wakeInst | Where-Object { ([string]$_.InstanceName).StartsWith($pnp + '_', [StringComparison]::OrdinalIgnoreCase) })) {
                    if (-not (Get-PreventDeviceEnable -Instance $wi)) { $wi | Set-CimInstance -Property @{ Enable = $true } -ErrorAction Stop }
                }
                Remove-ItemProperty -LiteralPath $devPath -Name 'PnPCapabilities' -ErrorAction Stop
                Log "${label}: removed an older version's settings and turned wake back on" -Colour Green -Indent
                Add-StepResult -Status Done -Reboot
            } catch {
                Log "[!] ${label}: $($_.Exception.Message)" -Colour Red -Indent
                Add-StepResult -Status Failed -Detail "${label}: $($_.Exception.Message)"
            }
        }

        # Connected Standby. Versions before 6.3.0 set CsEnabled to 0 and kept
        # the value it had in $CsBackupFile.
        $there = Test-PreventPathThere -Path $CsBackupFile -NotDone 'Connected Standby was not put back from it, and it was not removed'
        if ($null -eq $there) { $unreadCount++ }
        if ($there) {
            $legacyCount++
            $csPut = $false
            try {
                $cs = [int]([string](Get-Content -LiteralPath $CsBackupFile -Raw -ErrorAction Stop)).Trim()
                Set-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Power' -Name 'CsEnabled' -Value $cs -Type DWord -ErrorAction Stop
                $csPut = $true
                Remove-PreventItem -Path $CsBackupFile
                Log "CsEnabled (Connected Standby): put back to $cs" -Colour Green -Indent
                Add-StepResult -Status Done -Reboot
            } catch {
                # A record that cannot be removed once CsEnabled is put back is
                # said as that. Until 6.3.1 it was "Could not put back
                # CsEnabled", with an error that did not name the record.
                if ($csPut) {
                    Log "[!] CsEnabled (Connected Standby) was put back to $cs, but $($_.Exception.Message)" -Colour Red -Indent
                    Add-StepResult -Status Failed -Detail "CsEnabled (Connected Standby) was put back to $cs, but $($_.Exception.Message)" -Reboot
                } else {
                    Log "[!] Could not put back CsEnabled: $($_.Exception.Message)" -Colour Red -Indent
                    Add-StepResult -Status Failed -Detail "Could not put back CsEnabled: $($_.Exception.Message)"
                }
            }
        }

        # Values older versions set, each removed only while it holds exactly
        # what they wrote, and only where an older version is known to have
        # run: the same values can be an administrator's or a policy's own.
        $policyPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
        $legacyValues = @(
            @{ Path = 'HKLM:\SYSTEM\CurrentControlSet\Control'; Name = 'ServicesPipeTimeout'; Value = 120000 }
            @{ Path = $policyPath; Name = 'LocalAccountTokenFilterPolicy'; Value = 1 }
            @{ Path = $policyPath; Name = 'FilterAdministratorToken'; Value = 0 }
        )
        foreach ($lv in $legacyValues) {
            $v = Get-RegistryValue -Path $lv.Path -Name $lv.Name
            if (-not ($v -is [int] -and $v -eq $lv.Value)) { continue }
            if ($legacyUnread) {
                $said = "$($lv.Name) = $($lv.Value) matches what older versions set, but whether one ran here could not be told, so it is left as it is"
                Log $said -Colour DarkGray -Indent
                Add-StepResult -Status Unchanged -Detail $said
                continue
            }
            if (-not $legacySeen) {
                $unseenCount++
                $said = "$($lv.Name) = $($lv.Value) matches what older versions set, but no older install was found here, so it is left as it is"
                Log $said -Colour DarkGray -Indent
                Add-StepResult -Status Unchanged -Detail $said
                continue
            }
            $legacyCount++
            try {
                Remove-ItemProperty -LiteralPath $lv.Path -Name $lv.Name -ErrorAction Stop
                Log "$($lv.Name) = $($lv.Value): removed" -Colour Green -Indent
                Add-StepResult -Status Done -Reboot
            } catch {
                Log "[!] Could not remove $($lv.Name): $($_.Exception.Message)" -Colour Red -Indent
                Add-StepResult -Status Failed -Detail "Could not remove $($lv.Name): $($_.Exception.Message)"
            }
        }

        # The NAT rule older versions created, and the monitor re-created.
        # It is removed whether or not an older install was found, unlike the
        # three values above. A read that fails is said, and -Undo ends
        # incomplete, so that it is run again to look.
        $natThere = $null
        try {
            $natThere = Test-CoworkNat
        } catch {
            $unreadCount++
            $script:UndoUnread = $true
            Log "[!] Could not read the NAT rules, so CoworkNAT was not checked: $($_.Exception.Message)" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail "Could not read the NAT rules, so the rule CoworkNAT older versions created was not checked: $($_.Exception.Message)"
        }
        if ($natThere -eq $true) {
            $legacyCount++
            try {
                Remove-NetNat -Name 'CoworkNAT' -Confirm:$false -ErrorAction Stop
                Log "Removed the NAT rule CoworkNAT" -Colour Green -Indent
                Add-StepResult -Status Done
            } catch {
                Log "[!] Could not remove the NAT rule CoworkNAT: $($_.Exception.Message)" -Colour Red -Indent
                Add-StepResult -Status Failed -Detail "Could not remove the NAT rule CoworkNAT: $($_.Exception.Message)"
            }
        }

        $flagFile = Join-Path $ClaudeAppData "disable-dynamic-memory.flag"
        $there = Test-PreventPathThere -Path $flagFile -NotDone 'it was not removed'
        if ($null -eq $there) { $unreadCount++ }
        if ($there) {
            $legacyCount++
            Remove-PreventItem -Path $flagFile
            Log "Removed $flagFile" -Colour Green -Indent
            Add-StepResult -Status Done
        }

        # Dynamic Memory, which older versions' monitor turned off on the
        # Hyper-V VMs named like Claude, with no record of it. It goes back on
        # where an older version is known to have run. Hyper-V changes it only
        # while a VM is off, so a VM that is not is named, and the record that
        # an older version ran here stays for the next -Undo.
        # Get-Command answers nothing for a command that is not installed:
        # without Hyper-V's cmdlets there is no Dynamic Memory to turn on.
        if ($legacySeen -and (Get-Command -Name Set-VMMemory -ErrorAction SilentlyContinue)) {
            $dmVms = @()
            try {
                $dmVms = @(Get-ClaudeVmWithoutDynamicMemory)
            } catch {
                $unreadCount++
                $script:UndoUnread = $true
                Log "[!] Could not read the Hyper-V VMs, so Dynamic Memory was not checked: $($_.Exception.Message)" -Colour Yellow -Indent
                Add-StepResult -Status Warning -Detail "Could not read the Hyper-V VMs, so Dynamic Memory an older version turned off was not checked: $($_.Exception.Message)"
            }
            foreach ($vm in $dmVms) {
                $legacyCount++
                $vmName = [string]$vm.Name
                if ([string]$vm.State -ne 'Off') {
                    $gatedLeft++
                    Log "[!] The VM $vmName is $($vm.State): turn it off and run -Undo again to turn its Dynamic Memory back on" -Colour Yellow -Indent
                    Add-StepResult -Status Warning -Detail "Dynamic Memory on the VM $vmName is still off: Hyper-V changes it only while the VM is off, so turn the VM off and run -Undo again"
                    continue
                }
                try {
                    Set-VMMemory -VM $vm -DynamicMemoryEnabled $true -ErrorAction Stop
                    if (@(Get-ClaudeVmWithoutDynamicMemory | Where-Object { [string]$_.Name -eq $vmName }).Count -eq 0) {
                        Log "Dynamic Memory on the VM ${vmName}: turned back on" -Colour Green -Indent
                        Add-StepResult -Status Done
                    } else {
                        $gatedLeft++
                        Log "[!] Dynamic Memory on the VM ${vmName} still reads off" -Colour Red -Indent
                        Add-StepResult -Status Failed -Detail "Dynamic Memory on the VM ${vmName} still reads off"
                    }
                } catch {
                    $gatedLeft++
                    Log "[!] Could not turn Dynamic Memory back on for the VM ${vmName}: $($_.Exception.Message)" -Colour Red -Indent
                    Add-StepResult -Status Failed -Detail "Could not turn Dynamic Memory back on for the VM ${vmName}: $($_.Exception.Message)"
                }
            }
        }

        # Service recovery. Versions before 6.3.0 set restart after 30, 60 and
        # 120 seconds on vmcompute and on the Cowork VM service without a
        # record. Where that is still exactly what is set, and step 1 had no
        # record to put back, the recovery is cleared as it was before.
        foreach ($svcName in @('vmcompute', $ServiceName)) {
            if (-not $svcName) { continue }
            # A service that is not installed has no recovery to clear. Until
            # 6.3.1 one that could not be read counted as not installed.
            try {
                $svc = Get-ClaudeServiceOrNull -Name $svcName
            } catch {
                $unreadCount++
                $script:UndoUnread = $true
                $said = "Could not read the service $svcName, so the restart recovery older versions set on it was not checked: $(Get-ClaudeErrorText $_)"
                Log "[!] $said" -Colour Yellow -Indent
                Add-StepResult -Status Warning -Detail $said
                continue
            }
            if ($null -eq $svc) { continue }
            if (-not (Test-RestartRecovery (ConvertFrom-FailureActionsBlob -Bytes (Get-ServiceFailureBytes -Service $svcName)))) { continue }
            # A record step 1 could not put back is what the next -Undo puts
            # back. Clearing the recovery here would leave it as neither.
            if ($null -ne (Get-PreventUndoItem -Kind ServiceFailureActions -Key $svcName)) {
                Log "${svcName}: its recovery settings are still recorded, so step 1 puts them back on the next -Undo" -Colour DarkGray -Indent
                continue
            }
            $legacyCount++
            $out = & sc.exe failure $svcName reset= 0 actions= '""' 2>&1
            $code = $LASTEXITCODE
            if ($code -eq 0 -and -not (Test-RestartRecovery (ConvertFrom-FailureActionsBlob -Bytes (Get-ServiceFailureBytes -Service $svcName)))) {
                Log "${svcName}: recovery set by an older version cleared" -Colour Green -Indent
                Add-StepResult -Status Done
            } else {
                Log "[!] ${svcName}: sc.exe failure returned ${code}: $(@($out) -join ' ')" -Colour Red -Indent
                Add-StepResult -Status Failed -Detail "${svcName}: could not clear the recovery an older version set"
            }
        }

        # The record that an older version ran here goes once nothing it set
        # is left for a later -Undo to finish, and every list was read: the
        # next -Undo looks at one that was not, and some of what it removes
        # it removes only where an older version is known to have run.
        if ($gatedLeft -eq 0 -and $unreadCount -eq 0 -and $script:CurrentStep.Status -ne 'Failed') {
            foreach ($li in @($script:UndoRecord.ToArray() | Where-Object { $_.Kind -eq 'LegacyInstall' })) {
                Clear-PreventUndoItem -Kind LegacyInstall -Key $li.Key
            }
        }

        if ($legacyCount -eq 0) {
            if ($unseenCount -eq 0 -and $unreadCount -eq 0) { Log "Nothing from older versions found" -Colour DarkGray -Indent }
        } else {
            Log "Hibernate and Fast Startup turned off before 6.3.0 were not recorded, so they are left as they are" -Colour DarkGray -Indent
            Add-StepResult -Status Unchanged -Detail "Hibernate and Fast Startup turned off by a version before 6.3.0 are left as they are: there is no record of what they were"
            # A note, not a check: plans that could not be read are not
            # named, and the note is left out.
            $ultimateLeft = @()
            try { $ultimateLeft = @(Get-UltimatePlanGuids) } catch { $ultimateLeft = @() }
            if ($ultimateLeft.Count -gt 0) {
                Log "Ultimate Performance plans are left in place: there is no record of which ones older versions made" -Colour DarkGray -Indent
            }
        }
    } catch {
        Log "[!] $($_.Exception.Message)" -Colour Red -Indent
        Add-StepResult -Status Failed -Detail $_.Exception.Message
    }
    Complete-Step

    # ----------------------------------------------------------------
    # 3. Health monitor and scheduled tasks
    # ----------------------------------------------------------------
    # Every removal in steps 3 to 5 is checked. Until 6.3.0 a task the
    # scheduler refused to delete was reported as not found, a process that
    # did not stop and a file that was not deleted as removed.
    Step 3 $steps "Removing scheduled tasks and health monitor..." -Name "Monitor and tasks"
    # Stop any running health monitor first. A process list that cannot be
    # read is said, and -Undo ends incomplete. Until 6.3.0 it read as no
    # monitor running, and one that was went on running.
    $monitors = @()
    try {
        $monitors = @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" -ErrorAction Stop |
            Where-Object { $_.CommandLine -match "Watch-ClaudeHealth" })
    } catch {
        $script:UndoUnread = $true
        Log "[!] Could not list the running processes, so a running health monitor was not looked for: $($_.Exception.Message)" -Colour Yellow -Indent
        Add-StepResult -Status Warning -Detail "Could not list the running processes, so a health monitor still running was not stopped: $($_.Exception.Message)"
    }
    foreach ($mon in $monitors) {
        $procId = [int]$mon.ProcessId
        $why = $null
        try {
            Stop-Process -Id $procId -Force -ErrorAction Stop
        } catch {
            $why = $_.Exception.Message
        }
        # Until 6.3.1 a process that could not be read counted as gone.
        try {
            $alive = -not (Wait-PreventProcessExit -Id $procId)
        } catch {
            $said = "Could not read whether the health monitor (PID $procId) stopped: $(Get-ClaudeErrorText $_)"
            Log "[!] $said" -Colour Red -Indent
            Add-StepResult -Status Failed -Detail $said
            continue
        }
        if ($alive) {
            if (-not $why) { $why = 'it is still running' }
            Log "[!] Could not stop the health monitor (PID $procId): $why" -Colour Red -Indent
            Add-StepResult -Status Failed -Detail "Could not stop the health monitor (PID $procId): $why"
        } elseif ($why) {
            # The stop failed, and the process is gone: it ended by itself.
            Log "The health monitor that was running (PID $procId) has ended" -Colour DarkGray -Indent
        } else {
            Log "Stopped running health monitor (PID $procId)" -Colour Green -Indent
            Add-StepResult -Status Done
        }
    }
    $foundAny = $false
    foreach ($tName in @($TaskName, $BootTaskName)) {
        $res = Remove-PreventTask -Name $tName
        if ($res -eq 'NotFound') {
            Log "$tName not found" -Colour DarkGray -Indent
            continue
        }
        $foundAny = $true
        if ($res -eq 'Removed') {
            Log "$tName removed" -Colour Green -Indent
            Add-StepResult -Status Done
        } else {
            Log "[!] The scheduled task ${tName}: $res" -Colour Red -Indent
            Add-StepResult -Status Failed -Detail "The scheduled task ${tName}: $res"
        }
    }
    if (-not $foundAny) {
        Log "No tasks to remove" -Colour DarkGray -Indent
    }
    # Clean up old watchdog script if present
    $null = Remove-PreventFile -Path (Join-Path $ClaudeAppData "cowork-watchdog.ps1") -What "Removed old watchdog script"
    Complete-Step

    # ----------------------------------------------------------------
    # 4. Fix Claude Desktop shortcuts
    # ----------------------------------------------------------------
    Step 4 $steps "Removing shortcuts..." -Name "Fix Claude Desktop shortcuts"
    $desktopLnk = Join-Path ([Environment]::GetFolderPath("Desktop")) "Fix Claude Desktop.lnk"
    $startMenuLnk = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Fix Claude Desktop.lnk"
    $removedLnk = $false
    foreach ($lnk in @($desktopLnk, $startMenuLnk)) {
        if (Remove-PreventFile -Path $lnk) { $removedLnk = $true }
    }
    if (-not $removedLnk) {
        Log "No shortcuts found" -Colour DarkGray -Indent
    }
    Complete-Step

    # ----------------------------------------------------------------
    # 5. Claude (Admin) and Claude (no GPU) launchers
    # ----------------------------------------------------------------
    # The RUNASADMIN compatibility flags on Claude.exe are left alone: no
    # version of this toolkit set them, so any there are someone else's.
    # Until 6.3.0 -Undo removed every one it found.
    Step 5 $steps "Removing Claude launchers..." -Name "Launchers"
    try {
        # Remove the LaunchClaudeAdmin scheduled task
        $res = Remove-PreventTask -Name 'LaunchClaudeAdmin'
        if ($res -eq 'Removed') {
            Log "Removed LaunchClaudeAdmin scheduled task" -Colour Green -Indent
            Add-StepResult -Status Done
        } elseif ($res -eq 'NotFound') {
            Log "LaunchClaudeAdmin task not found" -Colour DarkGray -Indent
        } else {
            Log "[!] The scheduled task LaunchClaudeAdmin: $res" -Colour Red -Indent
            Add-StepResult -Status Failed -Detail "The scheduled task LaunchClaudeAdmin: $res"
        }
        # The task folder, now that step 3 and the line above have removed
        # this script's tasks from it. It goes only when it is empty.
        $folderRes = Remove-PreventTaskFolder
        if ($folderRes -eq 'Removed') {
            Log "Removed the empty task folder $TaskPath" -Colour Green -Indent
            Add-StepResult -Status Done
        } elseif ($folderRes -eq 'NotEmpty') {
            Log "The task folder $TaskPath still holds other tasks, so it is kept" -Colour DarkGray -Indent
        } elseif ($folderRes -ne 'NotFound') {
            Log "[!] The task folder ${TaskPath}: $folderRes" -Colour Red -Indent
            Add-StepResult -Status Failed -Detail "The task folder ${TaskPath}: $folderRes"
        }
        # Remove admin shortcut
        $null = Remove-PreventFile -Path (Join-Path ([Environment]::GetFolderPath("Desktop")) "Claude (Admin).lnk")
        # Remove launcher scripts. The NoGPU pair is written by step 18 and has
        # to come out here too, otherwise Undo leaves a shortcut on the desktop
        # pointing at a script this tool no longer maintains.
        $launcherCmd     = Join-Path $ClaudeAppData "Launch-Claude-Admin.cmd"
        $launcherPs1     = Join-Path $ClaudeAppData "Launch-Claude-Admin.ps1"
        $launcherVbs     = Join-Path $ClaudeAppData "Launch-Claude-Admin.vbs"
        $launcherTrigVbs = Join-Path $ClaudeAppData "Launch-Claude-Admin-Trigger.vbs"
        $noGpuCmd        = Join-Path $ClaudeAppData "Launch-Claude-NoGPU.cmd"
        $noGpuPs1        = Join-Path $ClaudeAppData "Launch-Claude-NoGPU.ps1"
        foreach ($lf in @($launcherCmd, $launcherPs1, $launcherVbs, $launcherTrigVbs, $noGpuCmd, $noGpuPs1)) {
            $null = Remove-PreventFile -Path $lf -What "Removed launcher"
        }
        $null = Remove-PreventFile -Path (Join-Path ([Environment]::GetFolderPath("Desktop")) "Claude (no GPU).lnk")
        # The copy of Claude's icon the shortcuts above used, from 6.3.0.
        $null = Remove-PreventFile -Path $ClaudeIconFile
    } catch {
        Log "[!] Could not fully remove the launchers: $($_.Exception.Message)" -Colour Red -Indent
        Add-StepResult -Status Failed -Detail "Could not fully remove the launchers: $($_.Exception.Message)"
    }
    Complete-Step

    Write-UndoOutcome

} else {

    # ================================================================
    # SETUP MODE
    # ================================================================
    $steps = 19
    # Claude Desktop may not be installed yet, so its AppData folder may not
    # exist. Later steps write their files there. The undo record's folder is
    # made here too: no setting is changed without a record.
    try {
        $null = Initialize-ClaudeAppData
    } catch {
        throw "$($_.Exception.Message). Nothing was changed."
    }
    if (-not (Test-Path -LiteralPath $UndoDir)) {
        try {
            $null = New-Item -Path $UndoDir -ItemType Directory -Force -ErrorAction Stop
        } catch {
            throw "Could not create $UndoDir, which holds the undo record ($($_.Exception.Message)). Nothing was changed."
        }
    }
    # A record written on another computer is moved aside and kept, and this
    # computer's own record starts empty.
    if ($script:UndoForeign) {
        $script:UndoForeignAside = Move-ForeignPreventUndo
    }

    # ----------------------------------------------------------------
    # 1. Power plan
    # ----------------------------------------------------------------
    # Until 6.3.0 every run wrote the active plan to $BackupFile, so a second
    # run replaced the plan you had with this script's own Ultimate
    # Performance plan, and -Undo "restored" that. It found Ultimate
    # Performance plans by their English name, so on another language it made
    # a new copy every run. Now the plan you had is recorded once, the first
    # time, and the existing copy is found by what it is.
    Step 1 $steps "Configuring power plan..." -Name "Power plan"
    try {
        $current = Get-ActivePlanGuid
        $ultimates = @(Get-UltimatePlanGuids)
        if ($null -eq (Get-PreventUndoItem -Kind PowerPlan -Key Active)) {
            $original = $current
            # Until 6.3.1 Test-Path was asked, which answers $false for a
            # file that cannot be read as for one that is not there
            # (measured 2026-11-23): such a record was taken for none, and
            # the plan active now, likely the older version's own copy, was
            # recorded as yours.
            $legacyThere = $null
            try {
                $legacyThere = Get-ClaudeItemOrNull -LiteralPath $BackupFile
            } catch {
                throw "the older version's record of your plan, $BackupFile, could not be read ($(Get-ClaudeErrorText $_)), so the plan you had is not recorded yet, and the record is kept"
            }
            if ($null -ne $legacyThere) {
                # An older version ran, and the active plan is likely its own
                # Ultimate Performance copy. The plan you had is the one in its
                # record, or Balanced when that record is no use. Until this
                # was checked, a record naming a plan since deleted left its
                # copy recorded as yours, and -Undo kept it active.
                # A record that could not be read is not one that holds no
                # plan: the step stops, and the record is kept. Until 6.3.1 it
                # was taken for one, Balanced was recorded as the plan you
                # had, and the record was deleted below.
                $legacy = $null
                try {
                    $legacy = ([string](Get-Content -LiteralPath $BackupFile -Raw -ErrorAction Stop)).Trim().ToLowerInvariant()
                } catch {
                    throw "the older version's record of your plan, $BackupFile, could not be read ($(Get-ClaudeErrorText $_)), so the plan you had is not recorded yet, and the record is kept"
                }
                $why = $null
                if ($legacy -notmatch "^$GuidPattern$") {
                    $why = "an older version's record of your original plan cannot be read"
                } else {
                    $legacyExists = Test-PowerPlanExists -Plan $legacy
                    if ($null -eq $legacyExists) { throw "powercfg could not list the power plans, so the plan you had is not recorded yet, and the older version's record of it is kept" }
                    if (-not $legacyExists) {
                        $why = "the plan an older version recorded as yours ($legacy) no longer exists"
                    } elseif ($ultimates -contains $legacy) {
                        $why = "an older version overwrote its record of your original plan with its own"
                    }
                }
                if ($why) {
                    $original = $BalancedGuid
                    Log "[!] -Undo will return to Balanced: $why" -Colour Yellow -Indent
                    Add-StepResult -Status Warning -Detail "-Undo returns to Balanced: $why"
                } else {
                    $original = $legacy
                }
            }
            if (-not $original) {
                $original = $BalancedGuid
                Log "[!] Could not read the active plan, so -Undo will return to Balanced" -Colour Yellow -Indent
                Add-StepResult -Status Warning -Detail "-Undo returns to Balanced: the active plan could not be read"
            }
            Add-PreventUndoItem -Kind PowerPlan -Key Active -Original $original
            Log "Recorded the plan to return to on -Undo: $(Get-PowerPlanName -Plan $original)" -Colour DarkGray -Indent
        }
        if (Test-PreventPathThere -Path $BackupFile -NotDone 'it was not recorded or removed') {
            # It shows an older version ran here, which -Undo needs to know
            # once the file is gone.
            Add-LegacyInstallRecord -What 'power plan record' -Where $BackupFile
            Remove-PreventItem -Path $BackupFile
        }

        $target = $null
        $createdItem = Get-PreventUndoItem -Kind PowerPlan -Key Created
        if ($createdItem) {
            $cg = ([string]$createdItem.Original).ToLowerInvariant()
            $cgExists = Test-PowerPlanExists -Plan $cg
            if ($null -eq $cgExists) { throw "powercfg could not list the power plans, so the plan this script created ($cg) stays recorded" }
            if ($cgExists) { $target = $cg } else { Clear-PreventUndoItem -Kind PowerPlan -Key Created }
        }
        if (-not $target -and $current -and ($ultimates -contains $current)) { $target = $current }
        if (-not $target -and $ultimates.Count -gt 0) { $target = $ultimates[0] }
        if (-not $target) {
            # Recorded before it is made, so a run that stops half way still
            # leaves -Undo the plan to delete.
            $newGuid = [guid]::NewGuid().ToString()
            Add-PreventUndoItem -Kind PowerPlan -Key Created -Original $newGuid
            $r = Invoke-PowerCfg -Arguments @('/duplicatescheme', $UltimateTemplateGuid, $newGuid)
            $made = $false
            if ($r.ExitCode -eq 0) {
                $made = Test-PowerPlanExists -Plan $newGuid
                # Made, perhaps, but not listed: the record stays, so -Undo
                # deletes the plan if it is there.
                if ($null -eq $made) { throw "powercfg could not list the power plans after making $newGuid, which stays recorded for -Undo" }
            }
            if ($made) {
                $target = $newGuid
                Log "Created an Ultimate Performance plan ($newGuid)" -Colour Green -Indent
            } else {
                Clear-PreventUndoItem -Kind PowerPlan -Key Created
                Log "Windows has no Ultimate Performance plan to copy" -Colour DarkGray -Indent
            }
        }
        if (-not $target) {
            $highExists = Test-PowerPlanExists -Plan $HighPerfGuid
            if ($null -eq $highExists) { throw "powercfg could not list the power plans" }
            if ($highExists) { $target = $HighPerfGuid }
        }
        if (-not $target) {
            $target = $current
            Log "[!] Neither Ultimate Performance nor High performance is available; step 2 tunes your current plan" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail "No Ultimate Performance or High performance plan: step 2 tunes your current plan instead"
        }
        if ($target -eq $current) {
            Log "Active plan: $(Get-PowerPlanName -Plan $target) (no change needed)" -Colour DarkGray -Indent
            Add-StepResult -Status Unchanged
        } else {
            $r = Invoke-PowerCfg -Arguments @('/setactive', $target)
            if ((Get-ActivePlanGuid) -eq $target) {
                Log "Active plan: $(Get-PowerPlanName -Plan $target)" -Colour Green -Indent
                Add-StepResult -Status Done
            } else {
                Log "[!] Could not make $target the active plan: $($r.Output)" -Colour Red -Indent
                Add-StepResult -Status Failed -Detail "powercfg could not make $target the active plan"
            }
        }
    } catch {
        Log "[!] $($_.Exception.Message)" -Colour Red -Indent
        Add-StepResult -Status Failed -Detail $_.Exception.Message
    }

    # ----------------------------------------------------------------
    # 2. Full performance settings on the active plan
    # ----------------------------------------------------------------
    # Set on the plan that is active now, and each one read back. Until 6.3.0
    # five of these were set, through "powercfg /change" and SCHEME_CURRENT,
    # and none was checked.
    Step 2 $steps "Setting the active plan to full performance (AC power)..." -Name "Full performance plan settings"
    try {
        $plan = Get-ActivePlanGuid
        if (-not $plan) { throw "Could not read the active power plan." }
        $changedAny = $false
        $missingLabels = @()
        foreach ($opt in $PowerOptions) {
            $res = Write-AcPowerValue -Plan $plan -Sub $opt.Sub -Setting $opt.Setting -Value $opt.Value
            if ($res -eq 'Done') {
                Log "$($opt.Label): $($opt.Shown)" -Colour Green -Indent
                Add-StepResult -Status Done
                $changedAny = $true
            } elseif ($res -eq 'Missing') {
                $missingLabels += $opt.Label
            } elseif ($res -eq 'Failed') {
                Log "[!] $($opt.Label): could not be set" -Colour Red -Indent
                Add-StepResult -Status Failed -Detail "$($opt.Label) could not be set"
            }
        }
        if ($missingLabels.Count -gt 0) {
            Log "Not on this machine: $($missingLabels -join ', ')" -Colour DarkGray -Indent
        }
        if ($changedAny) {
            # A changed value of the active plan applies when the plan is
            # activated again.
            $null = Invoke-PowerCfg -Arguments @('/setactive', $plan)
        } elseif ($script:CurrentStep.Status -eq 'None') {
            Log "Every setting already at full performance" -Colour DarkGray -Indent
            Add-StepResult -Status Unchanged
        }
    } catch {
        Log "[!] $($_.Exception.Message)" -Colour Red -Indent
        Add-StepResult -Status Failed -Detail $_.Exception.Message
    }

    # ----------------------------------------------------------------
    # 3. Hibernate
    # ----------------------------------------------------------------
    Step 3 $steps "Turning off hibernate..." -Name "Hibernate"
    try {
        $pwrKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Power'
        $he = Get-RegistryValue -Path $pwrKey -Name 'HibernateEnabled'
        if ($null -eq $he) { $he = Get-RegistryValue -Path $pwrKey -Name 'HibernateEnabledDefault' }
        if ($he -is [int] -and $he -eq 0) {
            Log "Hibernate already off" -Colour DarkGray -Indent
            Add-StepResult -Status Unchanged
        } else {
            Add-PreventUndoItem -Kind Hibernate -Key 'Hibernate' -Original $true
            $r = Invoke-PowerCfg -Arguments @('/hibernate', 'off')
            if ((Get-RegistryValue -Path $pwrKey -Name 'HibernateEnabled') -eq 0) {
                Log "Hibernate: off" -Colour Green -Indent
                Add-StepResult -Status Done
            } else {
                Clear-PreventUndoItem -Kind Hibernate -Key 'Hibernate'
                Log "[!] Could not turn hibernate off: $($r.Output)" -Colour Red -Indent
                Add-StepResult -Status Failed -Detail "powercfg could not turn hibernate off: $($r.Output)"
            }
        }
    } catch {
        Log "[!] $($_.Exception.Message)" -Colour Red -Indent
        Add-StepResult -Status Failed -Detail $_.Exception.Message
    }

    # ----------------------------------------------------------------
    # 4. Fast Startup
    # ----------------------------------------------------------------
    Step 4 $steps "Turning off Fast Startup..." -Name "Fast Startup"
    try {
        $res = Write-RecordedDword -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power' -Name 'HiberbootEnabled' -Value 0
        if ($res -eq 'Done') {
            Log "Fast Startup: off" -Colour Green -Indent
            Add-StepResult -Status Done
        } elseif ($res -eq 'Unchanged') {
            Log "Fast Startup already off" -Colour DarkGray -Indent
            Add-StepResult -Status Unchanged
        } else {
            Log "[!] Could not turn Fast Startup off" -Colour Red -Indent
            Add-StepResult -Status Failed -Detail "Could not write HiberbootEnabled"
        }
    } catch {
        Log "[!] $($_.Exception.Message)" -Colour Red -Indent
        Add-StepResult -Status Failed -Detail $_.Exception.Message
    }

    # ----------------------------------------------------------------
    # 5. Modern Standby
    # ----------------------------------------------------------------
    # Until 6.3.0 this set CsEnabled to 0, which current Windows no longer
    # reads: on a machine with Modern Standby nothing changed, and the summary
    # said it was off. PlatformAoAcOverride = 0 is the switch Windows reads
    # now. "powercfg /a" says whether the machine uses Modern Standby at all.
    Step 5 $steps "Checking Modern Standby..." -Name "Modern Standby"
    try {
        $a = Invoke-PowerCfg -Arguments @('/a')
        $ms = Get-ModernStandbyState -Text $a.Output
        if ($null -eq $ms) {
            Log "[!] Could not read the sleep states from powercfg /a" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail "Could not read the sleep states (powercfg /a)"
        } elseif ($ms -eq 'Disabled') {
            Log "Modern Standby is not in use on this machine" -Colour DarkGray -Indent
            Add-StepResult -Status Unchanged
        } else {
            $res = Write-RecordedDword -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Power' -Name 'PlatformAoAcOverride' -Value 0
            if ($res -eq 'Done') {
                Log "Modern Standby: off from the next Windows restart" -Colour Green -Indent
                Add-StepResult -Status Done -Reboot
            } elseif ($res -eq 'Unchanged') {
                Log "[!] Modern Standby is turned off but still in use: restart Windows" -Colour Yellow -Indent
                Add-StepResult -Status Warning -Detail "Turned off but still in use: restart Windows" -Reboot
            } else {
                Log "[!] Could not turn Modern Standby off" -Colour Red -Indent
                Add-StepResult -Status Failed -Detail "Could not write PlatformAoAcOverride"
            }
        }
    } catch {
        Log "[!] $($_.Exception.Message)" -Colour Red -Indent
        Add-StepResult -Status Failed -Detail $_.Exception.Message
    }

    # ----------------------------------------------------------------
    # 6. Network adapter power saving
    # ----------------------------------------------------------------
    Step 6 $steps "Turning off network adapter power saving..." -Name "Network adapter power saving"
    try {
        Disable-NicPowerSaving
    } catch {
        Log "[!] $($_.Exception.Message)" -Colour Red -Indent
        Add-StepResult -Status Failed -Detail $_.Exception.Message
    }

    # ----------------------------------------------------------------
    # 7. Power throttling
    # ----------------------------------------------------------------
    # Windows' power throttling runs background processes on efficient
    # cores at low clocks, and the VM's worker can be one of them.
    Step 7 $steps "Turning off Windows power throttling..." -Name "Power throttling"
    try {
        $res = Write-RecordedDword -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Power\PowerThrottling' -Name 'PowerThrottlingOff' -Value 1
        if ($res -eq 'Done') {
            Log "Power throttling: off for every process, from the next Windows restart" -Colour Green -Indent
            Add-StepResult -Status Done -Reboot
        } elseif ($res -eq 'Unchanged') {
            Log "Power throttling already off" -Colour DarkGray -Indent
            Add-StepResult -Status Unchanged
        } else {
            Log "[!] Could not turn power throttling off" -Colour Red -Indent
            Add-StepResult -Status Failed -Detail "Could not write PowerThrottlingOff"
        }
    } catch {
        Log "[!] $($_.Exception.Message)" -Colour Red -Indent
        Add-StepResult -Status Failed -Detail $_.Exception.Message
    }

    # ----------------------------------------------------------------
    # 8. Cowork VM worker priority
    # ----------------------------------------------------------------
    # Until 6.3.0 this raised every vmwp.exe on the machine: every Hyper-V VM,
    # WSL2 and Windows Sandbox. Now only the worker running a Cowork compute
    # system. A priority does not survive the process, so the health monitor
    # raises it again whenever the VM starts.
    Step 8 $steps "Raising the Cowork VM's priority..." -Name "Cowork VM priority"
    try {
        $hcsExe = $script:ClaudeEnv.HcsDiagPath
        if (-not $hcsExe) {
            # The discovery leaves HcsDiagPath unset for an hcsdiag.exe that
            # is not there and for one that could not be read, and says the
            # second before step 1. The step says which: one that could not
            # be read is a Warning that names it and why. Until 6.3.1 both
            # were Skipped, "hcsdiag.exe is not available", as if Windows did
            # not have it.
            $hcsThere = Test-PreventPathThere -Path "$env:SystemRoot\System32\hcsdiag.exe" -NotDone "the Cowork VM's worker cannot be told from other VMs' and was not raised"
            if ($null -ne $hcsThere) {
                Log "hcsdiag.exe not available, so the Cowork VM's worker cannot be told from other VMs'" -Colour DarkGray -Indent
                Add-StepResult -Status Skipped -Detail "hcsdiag.exe is not available"
            }
        } else {
            $hcsList = & $hcsExe list 2>&1 | Out-String
            $listCode = $LASTEXITCODE
            if ($listCode -ne 0) { throw "hcsdiag list returned $listCode" }
            # A process list that cannot be read fails the step. Read as no
            # worker, it said the Cowork VM was not running.
            $workers = @()
            try {
                $workers = @(Get-CimInstance Win32_Process -Filter "Name = 'vmwp.exe'" -ErrorAction Stop)
            } catch {
                throw "Could not list the VM worker processes: $($_.Exception.Message)"
            }
            $ids = @(Get-CoworkWorkerProcessIds -ListOutput ([string]$hcsList) -Workers $workers)
            if ($ids.Count -eq 0) {
                Log "The Cowork VM is not running; the health monitor raises it when it starts" -Colour DarkGray -Indent
                Add-StepResult -Status Skipped -Detail "The Cowork VM is not running"
            }
            foreach ($id in $ids) {
                # A worker can exit between the query and here, when the VM
                # shuts down. Until 6.3.0 that failed the whole step, and
                # until 6.3.1 a worker that could not be read had "exited".
                try {
                    $proc = Get-PreventProcessByIdOrNull -Id $id
                } catch {
                    $said = "Could not read the Cowork VM worker (PID $id), so its priority was not checked or raised: $(Get-ClaudeErrorText $_)"
                    Log "[!] $said" -Colour Yellow -Indent
                    Add-StepResult -Status Warning -Detail $said
                    continue
                }
                if (-not $proc) {
                    Log "Cowork VM worker (PID $id) has exited" -Colour DarkGray -Indent
                    Add-StepResult -Status Skipped -Detail "The Cowork VM's worker exited"
                } elseif ($proc.PriorityClass -eq 'AboveNormal') {
                    Log "Cowork VM worker (PID $id) already AboveNormal" -Colour DarkGray -Indent
                    Add-StepResult -Status Unchanged
                } else {
                    $proc.PriorityClass = 'AboveNormal'
                    Log "Cowork VM worker (PID $id): AboveNormal" -Colour Green -Indent
                    Add-StepResult -Status Done
                }
            }
        }
    } catch {
        Log "[!] $($_.Exception.Message)" -Colour Red -Indent
        Add-StepResult -Status Failed -Detail $_.Exception.Message
    }

    # ----------------------------------------------------------------
    # 9. vmcompute recovery
    # ----------------------------------------------------------------
    Step 9 $steps "Configuring Hyper-V compute service recovery..." -Name "vmcompute recovery"
    try {
        # Until 6.3.1 a service that could not be read was "not installed".
        try {
            $vmc = Get-ClaudeServiceOrNull -Name vmcompute
        } catch {
            throw "Could not read the service vmcompute, so its restart recovery was not checked or set: $(Get-ClaudeErrorText $_)"
        }
        if ($null -eq $vmc) {
            Log "vmcompute is not installed (Hyper-V is off)" -Colour DarkGray -Indent
            Add-StepResult -Status Skipped
        } else {
            $bytes = Get-ServiceFailureBytes -Service vmcompute
            if (Test-RestartRecovery (ConvertFrom-FailureActionsBlob -Bytes $bytes)) {
                Log "vmcompute already restarts after 30s/60s/120s" -Colour DarkGray -Indent
                Add-StepResult -Status Unchanged
            } else {
                $origBlob = $null
                if ($null -ne $bytes) { $origBlob = [Convert]::ToBase64String($bytes) }
                Add-PreventUndoItem -Kind ServiceFailureActions -Key vmcompute -Original $origBlob
                $out = & sc.exe failure vmcompute reset= 300 actions= restart/30000/restart/60000/restart/120000 2>&1
                $code = $LASTEXITCODE
                if ($code -eq 0 -and (Test-RestartRecovery (ConvertFrom-FailureActionsBlob -Bytes (Get-ServiceFailureBytes -Service vmcompute)))) {
                    Log "vmcompute failure recovery: restart after 30s/60s/120s (reset after 300s)" -Colour Green -Indent
                    Add-StepResult -Status Done
                } else {
                    Log "[!] sc.exe failure returned ${code}: $(@($out) -join ' ')" -Colour Red -Indent
                    Add-StepResult -Status Failed -Detail "Could not set the vmcompute recovery (sc.exe returned $code)"
                }
            }
        }
    } catch {
        Log "[!] $($_.Exception.Message)" -Colour Red -Indent
        Add-StepResult -Status Failed -Detail $_.Exception.Message
    }

    # ----------------------------------------------------------------
    # 10. Stale Cowork VMs
    # ----------------------------------------------------------------
    Step 10 $steps "Removing stale Cowork VMs..." -Name "Stale VM removal"
    # Three corrections in v6.0.0:
    #   1. hcsdiag has no "close" verb. The verbs are list, exec, console, read,
    #      write, share and kill. Every call this block made was a silent no-op
    #      that then logged success.
    #   2. The parse was wrong, in its own way, and is now Get-CoworkHcsGuids in
    #      the shared ClaudeEnv region. Fix and Watch had their own copies with
    #      their own variations on the same two mistakes.
    #   3. It ran unconditionally. If Claude was open, "stale" meant the VM the
    #      user was actively working in. Now it only acts when Claude is closed.
    #   4. (6.3.0) With exactly one stale VM the parser returned a single
    #      string, and .Count on a string throws under StrictMode in
    #      Windows PowerShell 5.1, so the one case this step exists for
    #      logged "check failed" and removed nothing. Wrapped in @() now.
    $hcsExe = $script:ClaudeEnv.HcsDiagPath
    if ($hcsExe) {
        try {
            # Every process named claude counts, the Claude Code CLI included:
            # this only decides to leave the VM alone, and erring that way is safe.
            # A list that could not be read leaves them alone too, and says so:
            # until 6.3.1 it read as Claude not running, and the VMs were killed.
            $claudeLive = $false
            $claudeUnread = $null
            try {
                $claudeLive = @(Get-ClaudeProcessByName -Name claude).Count -gt 0
            } catch {
                $claudeUnread = Get-ClaudeErrorText $_
            }
            $hcsList = & $hcsExe list 2>&1 | Out-String
            # Until 6.3.0 a list that failed read as an empty one: "HCS state clean".
            $listCode = $LASTEXITCODE
            if ($listCode -ne 0) { throw "hcsdiag list returned $listCode" }
            $guids = @(Get-CoworkHcsGuids -ListOutput ([string]$hcsList))

            if ($guids.Count -eq 0) {
                Log "HCS state clean" -Colour Green -Indent
                Add-StepResult -Status Unchanged
            } elseif ($null -ne $claudeUnread) {
                $said = "Could not list the processes named claude, so it is not known whether Claude is running, and the $($guids.Count) cowork-vm instance(s) are left alone: $claudeUnread"
                Log "[!] $said" -Colour Yellow -Indent
                Add-StepResult -Status Warning -Detail $said
            } elseif ($claudeLive) {
                Log "$($guids.Count) cowork-vm instance(s) present but Claude is running, leaving them alone" -Colour DarkGray -Indent
                Add-StepResult -Status Skipped -Detail "Claude is running, so its VM is in use"
            } else {
                $sent = @()
                foreach ($g in $guids) {
                    if ($PSCmdlet.ShouldProcess($g, "hcsdiag kill")) {
                        & $hcsExe kill $g 2>&1 | Out-Null
                        $killCode = $LASTEXITCODE
                        if ($killCode -eq 0) {
                            $sent += $g
                        } else {
                            Log "[!] hcsdiag kill returned $killCode for ${g}" -Colour Red -Indent
                            Add-StepResult -Status Failed -Detail "hcsdiag kill returned $killCode for ${g}"
                        }
                    }
                }
                # An exit code of 0 is not proof either: the list is read
                # again, and only a compute system gone from it counts.
                if ($sent.Count -gt 0) {
                    # A list that prints nothing is an answer, '' (no compute
                    # system left); a list that fails gives nothing.
                    $left = Get-CoworkHcsSurvivors -Guids $sent -List {
                        $t = & $hcsExe list 2>&1 | Out-String
                        if ($LASTEXITCODE -eq 0) { [string]$t }
                    }
                    foreach ($g in $sent) {
                        if ($null -eq $left) {
                            Log "hcsdiag accepted the kill of ${g}, but the list could not be read again to confirm it" -Colour DarkYellow -Indent
                            Add-StepResult -Status Warning -Detail "Could not confirm ${g} is gone"
                        } elseif ($left -contains $g) {
                            Log "[!] ${g} is still listed after hcsdiag kill" -Colour Red -Indent
                            Add-StepResult -Status Failed -Detail "${g} is still listed after hcsdiag kill"
                        } else {
                            Log "Killed stale HCS compute system: $g" -Colour Green -Indent
                            Add-StepResult -Status Done
                        }
                    }
                }
            }
        } catch {
            Log "[!] HCS cleanup check failed: $($_.Exception.Message)" -Colour Red -Indent
            Add-StepResult -Status Failed -Detail "HCS cleanup check failed: $($_.Exception.Message)"
        }
    } else {
        # As in step 8: an hcsdiag.exe that could not be read is a Warning
        # that names it and why, not Skipped as one that is not there. Until
        # 6.3.1 both were "not available".
        $hcsThere = Test-PreventPathThere -Path "$env:SystemRoot\System32\hcsdiag.exe" -NotDone "stale Cowork VMs were not looked for"
        if ($null -ne $hcsThere) {
            Log "hcsdiag.exe not available" -Colour DarkGray -Indent
            Add-StepResult -Status Skipped -Detail "hcsdiag.exe is not available"
        }
    }
    Complete-Step

    # ----------------------------------------------------------------
    # 11. Storage location
    # ----------------------------------------------------------------
    Step 11 $steps "Checking workspace storage location..." -Name "Storage location"
    $vmCachePath = $script:ClaudeEnv.VmCachePath
    $storageWarnings = @()
    $storageUnread = @()

    # Both checks read the discovered data root, which is where the VM cache
    # actually lives. Until 6.3.0 the sync folder test was a substring match,
    # so a folder named "Xbox" or "Boxes" read as a Box folder; see
    # Get-CloudSyncFolderName.
    $cloud = Get-CloudSyncFolderName -Path $ClaudeDataRoot
    if ($cloud) {
        $storageWarnings += "The Claude data folder is inside a $cloud sync folder, which causes mount failures"
    }

    # Is the Claude data folder on an external/USB drive, or a network
    # drive? Each check that cannot be made goes in $storageUnread with why,
    # and is a Warning. Until 6.3.0 every read here used SilentlyContinue
    # inside a catch that kept nothing, so a check that failed read as "not
    # external, not a network drive" and the step said the storage was OK.
    $appDataDrive = $null
    try {
        $appDataDrive = (Split-Path $ClaudeDataRoot -Qualifier -ErrorAction Stop) + "\"
    } catch {
        $storageUnread += "Could not tell which drive the Claude data folder $ClaudeDataRoot is on, so whether it is an external or a network drive was not checked: $($_.Exception.Message)"
    }
    if ($appDataDrive) {
        # Get-Volume, Get-Partition and Get-Disk each give an ObjectNotFound
        # error for what is not there: a drive letter that is not a local
        # volume, such as a network drive, or a volume with no partition on
        # a disk behind it. That leaves no disk to judge, and is not a
        # failure. Any other error is.
        try {
            $volume = @(Get-Volume -DriveLetter ($appDataDrive[0]) -ErrorAction Stop)
            $partition = @()
            if ($volume.Count -gt 0) { $partition = @(Get-Partition -DriveLetter ($appDataDrive[0]) -ErrorAction Stop) }
            $disk = @()
            if ($partition.Count -gt 0) { $disk = @(Get-Disk -Number $partition[0].DiskNumber -ErrorAction Stop) }
            if ($disk.Count -gt 0 -and [string]$disk[0].BusType -match "USB|Thunderbolt|1394") {
                $storageWarnings += "The Claude data folder is on an external $($disk[0].BusType) drive, use a local SSD instead"
            }
        } catch {
            if ($_.CategoryInfo.Category -ne [System.Management.Automation.ErrorCategory]::ObjectNotFound) {
                $storageUnread += "Could not check whether the drive $appDataDrive, which holds the Claude data folder, is an external drive: $($_.Exception.Message)"
            }
        }

        # Win32_LogicalDisk lists a network drive, which Get-Volume does not:
        # until 6.3.0 this check ran only for a drive Get-Volume found, so it
        # could never see one. A letter it does not list is no network drive.
        try {
            $logDisk = @(Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$($appDataDrive.TrimEnd('\'))'" -ErrorAction Stop)
            if ($logDisk.Count -gt 0 -and $logDisk[0].DriveType -eq 4) {
                $storageWarnings += "The Claude data folder is on a network drive, VirtioFS requires local storage"
            }
        } catch {
            $storageUnread += "Could not check whether the drive $appDataDrive, which holds the Claude data folder, is a network drive: $($_.Exception.Message)"
        }
    }

    # Check the VM cache path specifically for problematic locations. A
    # cache folder that could not be read is said, and its drive, which is
    # in its path, is checked all the same. Until 6.3.1 Test-Path was asked,
    # which answers $false for a folder that cannot be read as for one that
    # is not there (measured 2026-11-23): neither was checked, and the step
    # said "Storage location OK".
    $vmCacheThere = $false
    if ($vmCachePath) {
        try {
            $vmCacheThere = ($null -ne (Get-ClaudeItemOrNull -LiteralPath $vmCachePath))
        } catch {
            $storageUnread += "Could not read the VM cache folder ${vmCachePath}: $(Get-ClaudeErrorText $_)"
            $vmCacheThere = $true
        }
    }
    if ($vmCacheThere) {
        try {
            $vmDrive = (Split-Path $vmCachePath -Qualifier -ErrorAction Stop) + "\"
            $vmDisk = @(Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$($vmDrive.TrimEnd('\'))'" -ErrorAction Stop)
            if ($vmDisk.Count -gt 0 -and $vmDisk[0].DriveType -eq 4) {
                $storageWarnings += "VM cache is on a network drive, this will cause failures"
            }
        } catch {
            $storageUnread += "Could not check whether the VM cache $vmCachePath is on a network drive: $($_.Exception.Message)"
        }
    }

    foreach ($u in $storageUnread) {
        Log "[!] $u" -Colour Yellow -Indent
        Add-StepResult -Status Warning -Detail $u
    }
    if ($storageWarnings.Count -gt 0) {
        foreach ($w in $storageWarnings) {
            Log "WARNING: $w" -Colour DarkYellow -Indent
            Add-StepResult -Status Warning -Detail $w
        }
        Log "Recommended: Move Claude's data to a local SSD (C:\Users\$env:USERNAME\)" -Colour DarkGray -Indent
    } elseif ($storageUnread.Count -eq 0) {
        Log "Storage location OK (local drive)" -Colour Green -Indent
        Add-StepResult -Status Unchanged
    }

    # ----------------------------------------------------------------
    # 12. Clock
    # ----------------------------------------------------------------
    # Until 6.3.0 the offset was read with a pattern that matched a decimal
    # point only, so where the locale writes a comma the check could never
    # measure; it ran seconds after the network adapter step (step 8 then,
    # step 6 now, which no longer restarts them) had restarted the adapters,
    # so it could not reach the time server either; and a resync that failed
    # was reported as done.
    Step 12 $steps "Checking the clock..." -Name "Clock"
    try {
        # Until 6.3.1 a service that could not be read was "not installed".
        $w32 = $null
        $w32Unread = $null
        try {
            $w32 = Get-ClaudeServiceOrNull -Name W32Time
        } catch {
            $w32Unread = Get-ClaudeErrorText $_
        }
        if ($null -ne $w32Unread) {
            $said = "Could not read the Windows Time service, so it was not started and the clock was not checked: $w32Unread"
            Log "[!] $said" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail $said
        } elseif (-not $w32) {
            Log "[!] The Windows Time service is not installed" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail "The Windows Time service is not installed"
        } else {
            if ($w32.Status -ne 'Running') {
                try {
                    Start-Service -Name W32Time -ErrorAction Stop
                    Log "Started the Windows Time service" -Colour Green -Indent
                    Add-StepResult -Status Done
                } catch {
                    Log "[!] Could not start the Windows Time service: $($_.Exception.Message)" -Colour Yellow -Indent
                    Add-StepResult -Status Warning -Detail "Could not start the Windows Time service"
                }
            }
            # Wait for the network, 30 seconds at most, before asking the time
            # server anything. The 30 seconds bound the whole wait, lookups
            # included: a lookup can itself take seconds, and one that never
            # returned held the step for as long as it took, 15 times over.
            $resolved = $false
            $deadlineMs = 30000
            $netWait = [System.Diagnostics.Stopwatch]::StartNew()
            while (-not $resolved) {
                $leftMs = $deadlineMs - [int]$netWait.ElapsedMilliseconds
                if ($leftMs -le 0) { break }
                try {
                    $lookup = [System.Net.Dns]::GetHostAddressesAsync('time.windows.com')
                    if ($lookup.Wait($leftMs) -and @($lookup.Result).Count -gt 0) { $resolved = $true }
                } catch {
                    # A lookup that fails is tried again until the 30 seconds
                    # are up, and then the step says the server could not be
                    # reached, which is what was seen.
                    $null = $_
                }
                if ($resolved) { break }
                $leftMs = $deadlineMs - [int]$netWait.ElapsedMilliseconds
                if ($leftMs -le 0) { break }
                Start-Sleep -Milliseconds ([math]::Min(2000, $leftMs))
            }
            if (-not $resolved) {
                Log "[!] time.windows.com could not be reached within 30 seconds, so the clock was not checked" -Colour Yellow -Indent
                Add-StepResult -Status Warning -Detail "No network: time.windows.com could not be reached within 30 seconds, so the clock was not checked"
            } else {
                $chart = & w32tm /stripchart /computer:time.windows.com /dataonly /samples:3 2>&1 | Out-String
                $offset = Get-W32tmOffsetSeconds -Text $chart
                if ($null -eq $offset) {
                    Log "[!] Could not measure the clock offset" -Colour Yellow -Indent
                    Add-StepResult -Status Warning -Detail "w32tm could not measure the clock offset"
                } else {
                    $drift = [math]::Round([math]::Abs($offset), 2)
                    if ($drift -gt 5) {
                        $null = & w32tm /resync /force 2>&1
                        $code = $LASTEXITCODE
                        if ($code -eq 0) {
                            Log "The clock was $drift s off and has been resynced" -Colour Green -Indent
                            Add-StepResult -Status Done
                        } else {
                            Log "[!] The clock is $drift s off and the resync failed (w32tm returned $code)" -Colour Yellow -Indent
                            Add-StepResult -Status Warning -Detail "The clock is $drift s off and w32tm /resync failed"
                        }
                    } else {
                        Log "Clock offset: $drift s" -Colour DarkGray -Indent
                        Add-StepResult -Status Unchanged
                    }
                }
            }
        }
    } catch {
        Log "[!] $($_.Exception.Message)" -Colour Red -Indent
        Add-StepResult -Status Failed -Detail $_.Exception.Message
    }

    # ----------------------------------------------------------------
    # 13. Antivirus exclusions
    # ----------------------------------------------------------------
    Step 13 $steps "Checking antivirus configuration..." -Name "Antivirus"
    # The antivirus products registered with Windows Security Center (WMI).
    # Windows Server has no Security Center, so there its namespace is
    # missing: that is no failure, and the running processes below are what
    # is checked. Any other error is a Warning, since an antivirus that is
    # not Defender may then go unreported. Until 6.3.0 every error here read
    # as no product registered.
    $avProducts = @()
    $avUnread = $false
    try {
        $avItems = @(Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName "AntiVirusProduct" -ErrorAction Stop)
        foreach ($av in $avItems) {
            $avProducts += $av.displayName
        }
    } catch {
        if (-not (Test-CimClassMissingError -ErrorRecord $_)) {
            $avUnread = $true
            $said = "Could not read the antivirus products registered with Windows, so one that is not Defender may go unreported: $($_.Exception.Message)"
            Log "[!] $said" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail $said
        }
    }

    # Also check for common AV processes
    $knownAvProcesses = @{
        "MsMpEng"      = "Windows Defender"
        "mbamservice"  = "Malwarebytes"
        "avp"          = "Kaspersky"
        "avgnt"        = "Avira"
        "ccSvcHst"     = "Norton/Symantec"
        "bdagent"      = "Bitdefender"
        "ekrn"         = "ESET"
        "SentinelAgent" = "SentinelOne"
        "CrowdStrike"  = "CrowdStrike Falcon"
        "CSFalconService" = "CrowdStrike Falcon"
        "TmCCSF"       = "Trend Micro"
        "SophosSafestore" = "Sophos"
    }
    # Until 6.3.1 a process list that could not be read read as no
    # antivirus running.
    $runningAv = @()
    $avProcUnread = $false
    foreach ($proc in $knownAvProcesses.Keys) {
        try {
            if (@(Get-ClaudeProcessByName -Name $proc).Count -gt 0) {
                $runningAv += $knownAvProcesses[$proc]
            }
        } catch {
            $avProcUnread = $true
            $said = "Could not list the running processes, so an antivirus that is not Defender may go unreported: $(Get-ClaudeErrorText $_)"
            Log "[!] $said" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail $said
            break
        }
    }
    # Deduplicate
    $allAv = @(($avProducts + $runningAv) | Sort-Object -Unique)

    if ($allAv.Count -gt 0) {
        Log "Detected: $($allAv -join ', ')" -Colour Cyan -Indent
        $isDefenderOnly = ($allAv.Count -eq 1 -and $allAv[0] -match "Windows Defender")

        if ($isDefenderOnly) {
            # Each exclusion is recorded before it is added, so -Undo removes
            # the ones this script added and no others. Until 6.3.0 -Undo left
            # them all in place.
            try {
                # Stop, not SilentlyContinue: a Defender that cannot be read
                # is a failed check, not "nothing excluded yet".
                $exclusions = @((Get-MpPreference -ErrorAction Stop).ExclusionPath)
                # A list Defender hides holds only "N/A": nothing can be told
                # from it, so nothing is added and nothing recorded.
                if (Test-DefenderListHidden -List $exclusions) {
                    $said = "Defender hides its exclusion list from administrators on this machine, so this step cannot tell which path exclusions exist; none were added"
                    Log "[!] $said" -Colour Yellow -Indent
                    Add-StepResult -Status Warning -Detail $said
                } else {
                    # The app's data root, which holds the VM image whose VirtioFS
                    # disk operations an AV filter driver gets in the way of. Until
                    # 6.2.0 this was %APPDATA%\Claude, which on app 6.18355 holds
                    # only this toolkit's files.
                    $neededPaths = @(
                        $ClaudeDataRoot,
                        (Join-Path $env:ProgramFiles "Hyper-V"),
                        "$env:SystemRoot\System32\vmwp.exe",
                        "$env:SystemRoot\System32\vmms.exe"
                    )
                    $missing = @()
                    foreach ($np in $neededPaths) {
                        if (-not (Test-DefenderExclusionCovers -Path $np -Exclusions $exclusions)) { $missing += $np }
                    }
                    if ($missing.Count -gt 0) {
                        Log "Adding Defender exclusions for Hyper-V/Claude paths:" -Colour Green -Indent
                        foreach ($mp in $missing) {
                            Add-PreventUndoItem -Kind DefenderPath -Key $mp -Original $null
                            try {
                                Add-MpPreference -ExclusionPath $mp -ErrorAction Stop
                                Log "  + $mp" -Colour Green -Indent
                                Add-StepResult -Status Done
                            } catch {
                                Clear-PreventUndoItem -Kind DefenderPath -Key $mp
                                Log "  ! Could not add: $mp ($($_.Exception.Message))" -Colour Red -Indent
                                Add-StepResult -Status Failed -Detail "Could not add the exclusion $mp"
                            }
                        }
                    } else {
                        Log "Defender path exclusions already configured" -Colour Green -Indent
                    }
                }
                # The process exclusions are checked on every run. Until 6.2.0
                # they were only looked at when a path was missing, so a machine
                # whose paths were all excluded never got them.
                $requiredProcs = @("vmwp.exe", "vmms.exe", "vmcompute.exe", "$ServiceExe.exe")
                $procExclusions = @((Get-MpPreference -ErrorAction Stop).ExclusionProcess)
                if (Test-DefenderListHidden -List $procExclusions) {
                    $said = "Defender hides its exclusion list from administrators on this machine, so this step cannot tell which process exclusions exist; none were added"
                    Log "[!] $said" -Colour Yellow -Indent
                    Add-StepResult -Status Warning -Detail $said
                } else {
                    $missingProcs = @($requiredProcs | Where-Object { -not ($procExclusions -contains $_) })
                    if ($missingProcs.Count -gt 0) {
                        foreach ($rp in $missingProcs) {
                            Add-PreventUndoItem -Kind DefenderProcess -Key $rp -Original $null
                            try {
                                Add-MpPreference -ExclusionProcess $rp -ErrorAction Stop
                            } catch {
                                Log "  ! Could not add process exclusion: $rp ($($_.Exception.Message))" -Colour Red -Indent
                            }
                        }
                        $procExclusions = @((Get-MpPreference -ErrorAction Stop).ExclusionProcess)
                        $stillMissing = @($requiredProcs | Where-Object { -not ($procExclusions -contains $_) })
                        $added = @($missingProcs | Where-Object { $stillMissing -notcontains $_ })
                        if ($added.Count -gt 0) {
                            Log "  + Process exclusions: $($added -join ', ')" -Colour Green -Indent
                        }
                        if ($stillMissing.Count -gt 0) {
                            foreach ($sm in $stillMissing) { Clear-PreventUndoItem -Kind DefenderProcess -Key $sm }
                            Log "  ! Could not add process exclusions: $($stillMissing -join ', ')" -Colour Red -Indent
                            Add-StepResult -Status Failed -Detail "Could not add the process exclusions $($stillMissing -join ', ')"
                        } else {
                            Log "  All process exclusions verified" -Colour Green -Indent
                            Add-StepResult -Status Done
                        }
                    } else {
                        Log "Defender process exclusions already configured" -Colour Green -Indent
                    }
                }
                Add-StepResult -Status Unchanged
            } catch {
                Log "[!] Could not check the Defender exclusions: $($_.Exception.Message)" -Colour Red -Indent
                Add-StepResult -Status Failed -Detail "Could not check the Defender exclusions: $($_.Exception.Message)"
            }
        } else {
            # Third-party AV, can only advise
            Log "Recommended exclusion paths for your AV:" -Colour DarkYellow -Indent
            Log "  - $ClaudeDataRoot\" -Colour White -Indent
            Log "  - $env:ProgramFiles\Hyper-V\" -Colour White -Indent
            Log "  - $env:SystemRoot\System32\vmwp.exe" -Colour White -Indent
            Log "  - $env:SystemRoot\System32\vmms.exe" -Colour White -Indent
            Log "  - Process: vmcompute.exe, $ServiceExe.exe" -Colour White -Indent
            Log "Adding these exclusions prevents AV filter drivers from" -Colour DarkGray -Indent
            Log "interfering with VirtioFS disk operations" -Colour DarkGray -Indent
            Add-StepResult -Status Warning -Detail "Add the exclusions listed above to $($allAv -join ', '): this script cannot"
        }
    } elseif ($avUnread -or $avProcUnread) {
        # Only what was read is said.
        if (-not $avProcUnread) {
            Log "No antivirus product found among the running processes" -Colour DarkGray -Indent
        } elseif (-not $avUnread) {
            Log "No antivirus product is registered with Windows" -Colour DarkGray -Indent
        }
    } else {
        Log "No antivirus product detected" -Colour DarkGray -Indent
        Add-StepResult -Status Skipped
    }

    # ----------------------------------------------------------------
    # 14. WSL2 / Hyper-V conflicts
    # ----------------------------------------------------------------
    # A WSL2 distro that is running shares Hyper-V with the Cowork VM. WSL or
    # Docker merely being installed is not a conflict, so those are only
    # listed. Until 6.3.0 they were warnings, and the test for a running
    # distro read wsl.exe's UTF-16 output as text, so it never matched.
    Step 14 $steps "Checking for WSL2 / Hyper-V conflicts..." -Name "WSL2 conflicts"
    try {
        # A feature that is not there comes back as nothing, with no error
        # (measured 2026-11-22: Get-WindowsOptionalFeature -Online
        # -FeatureName NoSuchFeature_cf). Any error means it could not be
        # read: until 6.3.1 that was said only as a grey line.
        $wslFeature = $null
        try {
            $wslFeature = Get-WindowsOptionalFeature -Online -FeatureName Microsoft-Windows-Subsystem-Linux -ErrorAction Stop
        } catch {
            $said = "Could not read whether the WSL feature is installed: $(Get-ClaudeErrorText $_)"
            Log "[!] $said" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail $said
        }
        if ($wslFeature -and $wslFeature.State -eq 'Enabled') {
            Log "WSL is installed" -Colour DarkGray -Indent
        }
        if (Test-Path -LiteralPath (Join-Path $env:ProgramFiles "Docker\Docker\Docker.exe")) {
            Log "Docker Desktop is installed (it may use WSL2)" -Colour DarkGray -Indent
        }
        # wsl.exe -l -v lists the distros. It ends with exit code 0 when it
        # did; otherwise its output is a message, not a list. Until 6.3.1 the
        # exit code and the errors were not looked at, and any failure read
        # as "No WSL2 distro running".
        $runningWsl = @()
        $wslUnread = $null
        $wslExe = Join-Path $env:SystemRoot "System32\wsl.exe"
        # A wsl.exe that could not be read is a list that could not be read.
        # Until 6.3.1 Test-Path was asked, which answers $false for a file
        # that cannot be read as for one that is not there (measured
        # 2026-11-23), and "No WSL2 distro running" was said.
        $wslExeThere = $false
        try {
            $wslExeThere = ($null -ne (Get-ClaudeItemOrNull -LiteralPath $wslExe))
        } catch {
            $wslUnread = "$wslExe could not be read ($(Get-ClaudeErrorText $_))"
        }
        if ($wslExeThere) {
            $oldWslUtf8 = $env:WSL_UTF8
            try {
                $env:WSL_UTF8 = '1'
                $wslOut = @(& $wslExe -l -v 2>&1)
                $wslCode = $LASTEXITCODE
                if ($wslCode -eq 0) {
                    foreach ($l in $wslOut) {
                        if ($l -is [System.Management.Automation.ErrorRecord]) { continue }
                        $t = ([string]$l) -replace "`0", ''
                        if ($t -match 'Running' -and $t -match '\s2\s*$') {
                            $runningWsl += (($t.Trim().TrimStart('*').Trim()) -split '\s+')[0]
                        }
                    }
                } else {
                    # The list fails where WSL is not installed (measured
                    # 2026-11-22: exit code 1). No distro runs while neither
                    # WSL service runs, WSLService (WSL from the Store) or
                    # LxssManager (the WSL in Windows): that is told by the
                    # services, never by wsl.exe's message, which is in the
                    # display language. Anything else is a list that could
                    # not be read.
                    $wslSvcUp = @()
                    foreach ($svcName in @('WSLService', 'LxssManager')) {
                        $svc = Get-ClaudeServiceOrNull -Name $svcName
                        if ($null -ne $svc -and [string]$svc.Status -ne 'Stopped') { $wslSvcUp += $svcName }
                    }
                    if ($wslSvcUp.Count -gt 0) {
                        $first = @($wslOut | ForEach-Object { ([string]$_) -replace "`0", '' } | Where-Object { $_.Trim() }) | Select-Object -First 1
                        $wslUnread = "wsl.exe -l -v ended with exit code $wslCode while $($wslSvcUp -join ' and ') runs: $first"
                    }
                }
            } catch {
                $wslUnread = Get-ClaudeErrorText $_
            } finally {
                $env:WSL_UTF8 = $oldWslUtf8
            }
        }
        if ($null -ne $wslUnread) {
            $said = "Could not read which WSL distros are running: $wslUnread"
            Log "[!] $said" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail $said
        } elseif ($runningWsl.Count -gt 0) {
            Log "[!] WSL2 distros running: $($runningWsl -join ', ')" -Colour Yellow -Indent
            Log "If Cowork has issues, try: wsl --shutdown" -Colour DarkGray -Indent
            Add-StepResult -Status Warning -Detail "WSL2 running ($($runningWsl -join ', ')): if Cowork has issues, run wsl --shutdown"
        } else {
            Log "No WSL2 distro running" -Colour DarkGray -Indent
            Add-StepResult -Status Unchanged
        }
    } catch {
        Log "[!] $($_.Exception.Message)" -Colour Red -Indent
        Add-StepResult -Status Failed -Detail $_.Exception.Message
    }

    # ----------------------------------------------------------------
    # 15. Health monitor (auto-detects and auto-fixes crashes)
    # ----------------------------------------------------------------
    Step 15 $steps "Installing health monitor..." -Name "Health monitor"

    # Find Watch-ClaudeHealth.ps1 beside this script, else in C:\ClaudeFix or
    # in the Desktop or Documents folder. A place that could not be read
    # stops the search: the monitor is not installed from a later place in
    # its stead, and the step says why. Until 6.3.1 Test-Path was asked,
    # which answers $false for a file that cannot be read as for one that is
    # not there (measured 2026-11-23): such a script was passed over for a
    # copy found later, and with none the step said it was not beside this
    # script.
    $myDir = $null
    $watchScript = $null
    $watchUnread = $null
    try {
        if ($PSCommandPath) { $myDir = Split-Path $PSCommandPath -Parent }
        $watchCandidates = @()
        if ($myDir) { $watchCandidates += (Join-Path $myDir "Watch-ClaudeHealth.ps1") }
        $watchCandidates += @(
            "C:\ClaudeFix\Watch-ClaudeHealth.ps1",
            (Join-Path $env:USERPROFILE "Desktop\Watch-ClaudeHealth.ps1"),
            (Join-Path $env:USERPROFILE "Documents\Watch-ClaudeHealth.ps1")
        )
        $watchScript = Find-PreventKitFile -Candidates $watchCandidates
    } catch {
        $watchUnread = Get-ClaudeErrorText $_
    }

    # The basic watchdog script older versions installed, replaced by the
    # health monitor. It shows an older version ran here, so that is recorded
    # before it is deleted.
    $oldWatchdog = Join-Path $ClaudeAppData "cowork-watchdog.ps1"
    if (Test-PreventPathThere -Path $oldWatchdog -NotDone 'it was not recorded or removed') {
        try {
            Add-LegacyInstallRecord -What 'watchdog script' -Where $oldWatchdog
            Remove-Item -LiteralPath $oldWatchdog -Force -ErrorAction Stop
            Log "Removed old basic watchdog script" -Colour DarkGray -Indent
        } catch {
            Log "[!] Could not remove the old watchdog script ${oldWatchdog}: $($_.Exception.Message)" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail "Could not remove the old watchdog script ${oldWatchdog}: $($_.Exception.Message)"
        }
    }

    if ($watchScript) {
        Install-HealthMonitorTask -WatchScript $watchScript
    } elseif ($watchUnread) {
        Log "[!] Could not look for Watch-ClaudeHealth.ps1: $watchUnread" -Colour Red -Indent
        Add-StepResult -Status Failed -Detail "Could not look for Watch-ClaudeHealth.ps1, because $watchUnread, so the health monitor was not installed"
    } else {
        Log "[!] Watch-ClaudeHealth.ps1 not found in same folder" -Colour Yellow -Indent
        Log "Put all scripts in the same folder and rerun to enable health monitor" -Colour DarkGray -Indent
        Add-StepResult -Status Failed -Detail "Watch-ClaudeHealth.ps1 was not found beside this script"
    }
    Complete-Step

    # ----------------------------------------------------------------
    # 16. Fix Claude Desktop shortcuts (Desktop + Start Menu)
    # ----------------------------------------------------------------
    Step 16 $steps "Creating Fix Claude Desktop shortcuts..." -Name "Fix Claude Desktop shortcuts"

    # Fix-ClaudeDesktop.bat beside this script, else beside the first
    # Fix-ClaudeDesktop.ps1 found beside this script, in C:\ClaudeFix, or in
    # the Desktop or Documents folder. A place that could not be read stops
    # the search: no shortcut is made to a later place in its stead, and the
    # step says why. Until 6.3.1 Test-Path was asked, which answers $false
    # for a file that cannot be read as for one that is not there (measured
    # 2026-11-23): such a file was passed over, and with none found the step
    # said the launcher was not beside this script.
    $fixBat = $null
    $fixUnread = $null
    try {
        if ($myDir) { $fixBat = Find-PreventKitFile -Candidates @(Join-Path $myDir "Fix-ClaudeDesktop.bat") }
        if (-not $fixBat) {
            $fixCandidates = @()
            if ($myDir) { $fixCandidates += (Join-Path $myDir "Fix-ClaudeDesktop.ps1") }
            $fixCandidates += @(
                "C:\ClaudeFix\Fix-ClaudeDesktop.ps1",
                (Join-Path $env:USERPROFILE "Desktop\Fix-ClaudeDesktop.ps1"),
                (Join-Path $env:USERPROFILE "Documents\Fix-ClaudeDesktop.ps1")
            )
            $fixScript = Find-PreventKitFile -Candidates $fixCandidates
            if ($fixScript) {
                $fixBat = Find-PreventKitFile -Candidates @(Join-Path (Split-Path $fixScript -Parent) "Fix-ClaudeDesktop.bat")
            }
        }
    } catch {
        $fixUnread = Get-ClaudeErrorText $_
    }

    $shell = $null
    if ($fixBat) {
        try {
            $shell = New-Object -ComObject WScript.Shell
        } catch {
            Log "[!] Windows Script Host is unavailable, cannot create shortcuts" -Colour Red -Indent
            Add-StepResult -Status Failed -Detail "Windows Script Host is unavailable"
        }
    }

    if ($fixBat -and $shell) {

        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $desktopLnk = Join-Path $desktopPath "Fix Claude Desktop.lnk"
        try {
            $sc = $shell.CreateShortcut($desktopLnk)
            $sc.TargetPath = $fixBat
            $sc.WorkingDirectory = Split-Path $fixBat -Parent
            $sc.Description = "Reset and fix Claude Desktop / Cowork VM"
            $sc.WindowStyle = 1
            $sc.IconLocation = "%SystemRoot%\System32\shell32.dll,77"
            $sc.Save()
            Log "Desktop shortcut created" -Colour Green -Indent
            Add-StepResult -Status Done
        } catch {
            Log "[!] Could not create desktop shortcut: $($_.Exception.Message)" -Colour Red -Indent
            Add-StepResult -Status Failed -Detail "Could not create the desktop shortcut: $($_.Exception.Message)"
        }

        $startMenuPath = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs"
        $startMenuLnk = Join-Path $startMenuPath "Fix Claude Desktop.lnk"
        try {
            $sc2 = $shell.CreateShortcut($startMenuLnk)
            $sc2.TargetPath = $fixBat
            $sc2.WorkingDirectory = Split-Path $fixBat -Parent
            $sc2.Description = "Reset and fix Claude Desktop / Cowork VM"
            $sc2.WindowStyle = 1
            $sc2.IconLocation = "%SystemRoot%\System32\shell32.dll,77"
            $sc2.Save()
            Log "Start Menu shortcut created" -Colour Green -Indent
            Add-StepResult -Status Done
            Log "You can pin this to your Taskbar: search 'Fix Claude Desktop' in Start" -Colour DarkGray -Indent
        } catch {
            Log "[!] Could not create Start Menu shortcut: $($_.Exception.Message)" -Colour Red -Indent
            Add-StepResult -Status Failed -Detail "Could not create the Start Menu shortcut: $($_.Exception.Message)"
        }
    } elseif ($fixUnread) {
        Log "[!] Could not look for Fix-ClaudeDesktop.bat: $fixUnread" -Colour Red -Indent
        Add-StepResult -Status Failed -Detail "Could not look for Fix-ClaudeDesktop.bat, because $fixUnread, so the shortcuts were not made"
    } elseif (-not $fixBat) {
        Log "[!] Fix-ClaudeDesktop.bat not found in same folder" -Colour Red -Indent
        Log "Put all scripts in the same folder and rerun to create shortcuts" -Colour DarkGray -Indent
        Add-StepResult -Status Failed -Detail "Fix-ClaudeDesktop.bat was not found beside this script"
    }
    Complete-Step

    # ----------------------------------------------------------------
    # 17. Claude (Admin) launcher (MSIX-aware)
    # ----------------------------------------------------------------
    Step 17 $steps "Configuring Claude Desktop to launch elevated..." -Name "Claude (Admin) launcher"
    try {
        # Scheduled task registration requires admin -- skip gracefully if not elevated
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if (-not $isAdmin) {
            Log "Skipped, requires admin privileges (run as Administrator to enable)" -Colour DarkYellow -Indent
            throw "SKIP"
        }
        # MSIX apps block all direct .exe access from WindowsApps (ACLs, -Verb RunAs,
        # dir enumeration all fail). The only reliable way to launch an MSIX app with
        # full admin privileges is a scheduled task with RunLevel=Highest + Interactive
        # logon. The task gets a full unfiltered admin token, no UAC prompt, and the
        # GUI is visible in the user's desktop session.
        #
        # The task action finds Claude at runtime via two methods:
        #   1) Get-AppxPackage (MSIX installs from Store/winget)
        #   2) Common install paths (traditional .exe installer)
        # This survives version updates. A third, reading the path off a
        # running "claude" process, was removed in 6.2.0: it only ran when no
        # app process was running, so all it could find was the Claude Code CLI.

        $elevTaskName = "LaunchClaudeAdmin"

        # An old task is removed first. One that could not be read, or could
        # not be removed, is said: registering with -Force below replaces it
        # all the same. Until 6.3.1 both were passed over without a word.
        $oldTask = $null
        try {
            $oldTask = Get-ClaudeScheduledTaskOrNull -TaskName $elevTaskName -TaskPath $TaskPath
        } catch {
            $said = "Could not read whether the scheduled task $TaskPath$elevTaskName is there, so it was not removed first; registering it replaces it: $(Get-ClaudeErrorText $_)"
            Log "[!] $said" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail $said
        }
        if ($null -ne $oldTask) {
            try {
                Unregister-ScheduledTask -TaskName $elevTaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction Stop
            } catch {
                $said = "Could not remove the old scheduled task $TaskPath$elevTaskName; registering it again replaces it: $(Get-ClaudeErrorText $_)"
                Log "[!] $said" -Colour Yellow -Indent
                Add-StepResult -Status Warning -Detail $said
            }
        }

        # PowerShell command that finds and launches Claude
        # Priority: 1) MSIX via Get-AppxPackage  2) Traditional .exe install paths
        # NOTE: We MUST use direct .exe launch (Start-Process $exe) to inherit the
        # task's elevated token. shell:AppsFolder routes through the non-elevated desktop
        # shell and the app gets medium integrity -- defeating the entire purpose.
        # Trade-off: MSIX installs will show a second taskbar icon. This is unavoidable
        # because Windows enforces medium integrity for all shell-activated MSIX apps.
        $launchCmd = @'
# If Claude Desktop is already running, don't launch again. The Claude Code CLI
# runs as "claude" too and does not count: with a project thread or a terminal
# session alive, this used to exit without launching anything. A process whose
# path or product name cannot be read does count, so a doubt ends in no
# second launch.
#
# A process list that cannot be read is a doubt as well: Claude is not
# started, and why is shown, since this runs with no window. Get-Process says
# that none runs with an error of its own: id NoProcessFoundForGivenName,
# category ObjectNotFound, exception ProcessCommandException (measured
# 2026-11-22). Until 6.3.1 every error read as none running.
$existing = @()
try {
    $existing = @(Get-Process -Name claude -ErrorAction Stop | Where-Object {
        $pexe = $null
        try { $pexe = $_.MainModule.FileName } catch { $null = $_ }
        if (-not $pexe) { return $true }
        if ($pexe -match '[\\/]claude-code[\\/]') { return $false }
        $product = ''
        try { $product = [string](Get-Item -LiteralPath $pexe).VersionInfo.ProductName } catch { $null = $_ }
        $product -ne 'Claude Code'
    })
} catch {
    if (-not ("$($_.FullyQualifiedErrorId)" -like 'NoProcessFoundForGivenName,*' -and
              [string]$_.CategoryInfo.Category -eq 'ObjectNotFound' -and
              $_.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.ProcessCommandException')) {
        $null = (New-Object -ComObject WScript.Shell).Popup("Claude was not started: the running processes could not be read, so it is not known whether Claude is already running.`n`n$($_.Exception.Message)", 0, 'Claude (Admin)', 48)
        exit 1
    }
}
if ($existing.Count -gt 0) { exit 0 }
$exe = $null
# A place that could not be read is a doubt too: Claude is not started from
# a later place in its stead, and why is shown. Get-Item says a file is not
# there with id PathNotFound, category ObjectNotFound, exception
# ItemNotFoundException (measured 2026-11-23). Until 6.3.1 Test-Path was
# asked, which answers $false for a file that cannot be read, and the next
# place was tried. Get-AppxPackage lists nothing, with no error, when no
# package matches (measured 2026-11-23), so any error is a list not read.
function Test-ClaudeExe {
    param([string]$Path)
    try {
        return (-not (Get-Item -LiteralPath $Path -Force -ErrorAction Stop).PSIsContainer)
    } catch {
        if ("$($_.FullyQualifiedErrorId)" -like 'PathNotFound,*' -and [string]$_.CategoryInfo.Category -eq 'ObjectNotFound' -and
            $_.Exception.GetType().FullName -eq 'System.Management.Automation.ItemNotFoundException') { return $false }
        throw "$Path could not be read: $($_.Exception.Message)"
    }
}
try {
    # 1. MSIX install (Windows Store / winget MSIX)
    $p = Get-AppxPackage -ErrorAction Stop | Where-Object { $_.Name -eq 'Claude' -or $_.PackageFamilyName -like 'Claude_*' } | Select-Object -First 1
    if ($p) { $e = Join-Path $p.InstallLocation 'app\Claude.exe'; if (Test-ClaudeExe $e) { $exe = $e } }
    # 2. Traditional installer paths
    if (-not $exe) {
        $candidates = @(
            (Join-Path $env:LOCALAPPDATA 'Programs\claude-desktop\Claude.exe'),
            (Join-Path $env:LOCALAPPDATA 'Claude Desktop\Claude.exe'),
            (Join-Path $env:LOCALAPPDATA 'Programs\Claude\Claude.exe'),
            (Join-Path ${env:ProgramFiles} 'Claude Desktop\Claude.exe'),
            (Join-Path ${env:ProgramFiles(x86)} 'Claude Desktop\Claude.exe')
        )
        foreach ($c in $candidates) { if (Test-ClaudeExe $c) { $exe = $c; break } }
    }
} catch {
    $null = (New-Object -ComObject WScript.Shell).Popup("Claude was not started: $($_.Exception.Message)", 0, 'Claude (Admin)', 48)
    exit 1
}
# There was a third method, reading the path off a running "claude" process.
# By this line no app process is running, so it could only ever find the Claude
# Code CLI, and then launched that instead of the app.
if ($exe) { Start-Process $exe } else { throw 'Claude Desktop not found. Is it installed?' }
'@

        # Encode as Base64 for -EncodedCommand (handles the multi-line launcher
        # safely inside the single-line VBScript string below).
        $encodedCmd = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($launchCmd))

        # The task must NOT execute PowerShell.exe directly. Under an Interactive
        # principal, Task Scheduler hands a console-subsystem program a brand-new
        # VISIBLE console window, and -WindowStyle Hidden cannot stop it: PowerShell
        # only hides its console AFTER conhost has drawn it, so the window still
        # flashes to the foreground on every launch -- enough to drop a fullscreen
        # game to the desktop. Measured 2026-10-09: the old PowerShell.exe action
        # took the foreground for ~45 ms each time the shortcut was clicked.
        #
        # wscript.exe is a GUI-subsystem host, so the task itself opens no console,
        # and WScript.Shell.Run(cmd, 0, ...) starts PowerShell with its console
        # hidden at creation time (vbHide 0) -- which is exactly what -WindowStyle
        # Hidden cannot do. Same elevated token (wscript runs at the task's
        # RunLevel and PowerShell inherits it), same EncodedCommand payload, and
        # zero windows. Verified clean with tools/probes/probe_console_flash.ps1.
        try {
            $null = Initialize-ClaudeAppData
        } catch {
            throw "the launcher could not be placed: $($_.Exception.Message)"
        }
        $wscriptExe = Join-Path $env:SystemRoot "System32\wscript.exe"
        $elevVbs = Join-Path $ClaudeAppData "Launch-Claude-Admin.vbs"
        $elevVbsBody = @"
' Launch-Claude-Admin.vbs
' Auto-generated by Prevent-ClaudeIssues.ps1 v$ToolkitVersion
' Starts Claude elevated with NO console window. The 0 (vbHide) hides the
' PowerShell console at creation; -WindowStyle Hidden alone cannot, and is
' deliberately not relied on here.
CreateObject("WScript.Shell").Run "PowerShell.exe -NoProfile -ExecutionPolicy Bypass -EncodedCommand $encodedCmd", 0, False
"@
        Set-Content -Path $elevVbs -Value $elevVbsBody -Encoding ASCII -Force -ErrorAction Stop

        $elevAction = New-ScheduledTaskAction -Execute $wscriptExe `
            -Argument "`"$elevVbs`""

        $elevSettings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -ExecutionTimeLimit (New-TimeSpan -Minutes 1)

        $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $elevPrincipal = New-ScheduledTaskPrincipal `
            -UserId $currentUser `
            -RunLevel Highest `
            -LogonType Interactive

        # A registration the scheduler refuses is a failed step, and no
        # shortcut is made. One an earlier run made is removed when its task
        # is not there now: it would start nothing. Until 6.3.1 a refused
        # registration went on to say "Scheduled task created" and to make
        # the shortcut.
        $regWhy = $null
        try {
            $null = Register-ScheduledTask `
                -TaskName $elevTaskName `
                -TaskPath $TaskPath `
                -Action $elevAction `
                -Settings $elevSettings `
                -Principal $elevPrincipal `
                -Description "Launches Claude Desktop with full admin privileges. Triggered by the 'Claude (Admin)' shortcut." `
                -Force -ErrorAction Stop
        } catch {
            $regWhy = Get-ClaudeErrorText $_
        }
        if ($null -ne $regWhy) {
            $said = "the scheduler refused the task $TaskPath$elevTaskName, so no Claude (Admin) shortcut was made ($regWhy)"
            $staleLnk = Join-Path ([Environment]::GetFolderPath("Desktop")) "Claude (Admin).lnk"
            $staleTold = $true
            try {
                $taskLeft = $null -ne (Get-ClaudeScheduledTaskOrNull -TaskName $elevTaskName -TaskPath $TaskPath)
                $lnkLeft = $null -ne (Get-ClaudeItemOrNull -LiteralPath $staleLnk)
            } catch {
                $staleTold = $false
                $said += ". Whether the shortcut $staleLnk an earlier run made still has its task could not be read: $(Get-ClaudeErrorText $_)"
            }
            if ($staleTold -and $lnkLeft -and -not $taskLeft) {
                try {
                    Remove-Item -LiteralPath $staleLnk -Force -ErrorAction Stop
                    $said += ". The shortcut $staleLnk an earlier run made was removed: the task it starts is not there"
                } catch {
                    $said += ". The shortcut $staleLnk an earlier run made starts a task that is not there, and could not be removed: $(Get-ClaudeErrorText $_)"
                }
            }
            throw $said
        }

        Log "Scheduled task created: $TaskPath$elevTaskName (Highest + Interactive, no console)" -Colour Green -Indent

        # The desktop shortcut must not open a console either. A .cmd target
        # opened a visible cmd window (measured 2026-10-09: ~1 s in the
        # foreground) before schtasks even ran. Trigger the task through
        # wscript + vbHide instead, so clicking the shortcut shows nothing.
        $triggerVbs = Join-Path $ClaudeAppData "Launch-Claude-Admin-Trigger.vbs"
        $triggerVbsBody = @"
' Launch-Claude-Admin-Trigger.vbs
' Auto-generated by Prevent-ClaudeIssues.ps1 v$ToolkitVersion
' Fires the elevated LaunchClaudeAdmin task with no console window. The 0
' (vbHide) runs schtasks hidden; True waits so a missing task is reported.
Dim sh, rc
Set sh = CreateObject("WScript.Shell")
rc = sh.Run("schtasks /run /tn ""\Claude\LaunchClaudeAdmin""", 0, True)
If rc <> 0 Then
    MsgBox "The LaunchClaudeAdmin scheduled task was not found." & vbCrLf & "Run Prevent-ClaudeIssues.bat to set it up.", vbExclamation, "Claude (Admin)"
End If
"@
        Set-Content -Path $triggerVbs -Value $triggerVbsBody -Encoding ASCII -Force -ErrorAction Stop
        Log "Launcher created: $triggerVbs" -Colour Green -Indent

        # Remove the old .cmd launcher left by versions <= 6.1.0, if present.
        # It shows an older version ran here, so that is recorded first.
        $oldLauncherCmd = Join-Path $ClaudeAppData "Launch-Claude-Admin.cmd"
        if (Test-PreventPathThere -Path $oldLauncherCmd -NotDone 'it was not recorded or removed') {
            try {
                Add-LegacyInstallRecord -What 'admin launcher' -Where $oldLauncherCmd
                Remove-Item -LiteralPath $oldLauncherCmd -Force -ErrorAction Stop
                Log "Removed superseded .cmd launcher: $oldLauncherCmd" -Colour DarkGray -Indent
            } catch {
                Log "[!] Could not remove the superseded .cmd launcher ${oldLauncherCmd}: $($_.Exception.Message)" -Colour Yellow -Indent
                Add-StepResult -Status Warning -Detail "Could not remove the superseded .cmd launcher ${oldLauncherCmd}: $($_.Exception.Message)"
            }
        }

        # Create desktop shortcut pointing to the launcher
        $desktopPath = [Environment]::GetFolderPath("Desktop")

        # Remove any old broken shortcuts
        $oldLnk = Join-Path $desktopPath "Claude (Admin).lnk"
        if (Test-Path $oldLnk) {
            # Best effort: the shortcut is saved over it below, and a save
            # that fails fails the step. A shortcut Test-Path could not read
            # answers $false here, and is the same case: the save says.
            Remove-Item $oldLnk -Force -ErrorAction SilentlyContinue
        }

        $adminLnkPath = Join-Path $desktopPath "Claude (Admin).lnk"
        $shell = New-Object -ComObject WScript.Shell
        $sc = $shell.CreateShortcut($adminLnkPath)
        $sc.TargetPath = $wscriptExe
        $sc.Arguments = "`"$triggerVbs`""
        $sc.WorkingDirectory = $ClaudeAppData
        $sc.Description = "Claude Desktop (Elevated via scheduled task)"
        # A copy of Claude's icon kept in the toolkit folder. Until 6.3.0 this
        # pointed at Claude.exe itself, which for an MSIX install sits in a
        # folder named after the version: the first update removed that
        # folder and the shortcut lost its icon.
        $sc.IconLocation = Initialize-ClaudeIcon
        $sc.Save()
        Log "Admin shortcut created: $adminLnkPath" -Colour Green -Indent
        Log "No UAC prompt, task runs with full admin token automatically" -Colour DarkGray -Indent
        Log "Survives Claude updates: the launcher finds Claude when it runs, and the icon is a copy kept in the toolkit folder" -Colour DarkGray -Indent
        Add-StepResult -Status Done

    } catch {
        if ($_.Exception.Message -eq 'SKIP') {
            Add-StepResult -Status Skipped -Detail "Needs Administrator"
        } else {
            Log "[!] Could not configure the elevated launcher: $($_.Exception.Message)" -Colour Red -Indent
            Add-StepResult -Status Failed -Detail "Could not configure the elevated launcher: $($_.Exception.Message)"
        }
    }
    Complete-Step

    # ----------------------------------------------------------------
    # 18. GPU-free launcher
    # ----------------------------------------------------------------
    # There is a fatal GPU-process crash on current builds, upstream issue
    # 80444, exit code 101457950 (0x060C201E). A page loaded in the in-app
    # Browser preview runs a WebGL capability probe, the GPU process dies, and
    # the whole app dies with it. Confirmed in a user's logs on 2026-07-29:
    # browser preview created at 10:02:22, 57 INVALID_ENUM probe warnings at
    # 10:02:27, GPU gone at 10:02:28.
    #
    # Running with --disable-gpu avoids it. The obstacle is that an MSIX app
    # launched normally gets no command line, so the flag has nowhere to go.
    # Invoke-CommandInDesktopPackage starts the executable INSIDE the package
    # identity, which is what makes this work without breaking the install.
    #
    # This writes the launcher rather than changing anything. It is opt-in: the
    # user runs it when they need it, and the normal shortcut is untouched.
    Step 18 $steps "Creating GPU-free launcher..." -Name "GPU-free launcher"
    try {
        $noGpuPs1 = Join-Path $ClaudeAppData "Launch-Claude-NoGPU.ps1"
        $noGpuCmd = Join-Path $ClaudeAppData "Launch-Claude-NoGPU.cmd"

        # Everything identifying the install comes from discovery. Hardcoding
        # Claude_pzs8sxrjxfjjc would work on the machine this was written on and
        # nowhere else, which is the mistake this whole layer exists to prevent.
        $pfn      = $script:ClaudeEnv.PackageFamilyName
        $appId    = $script:ClaudeEnv.ApplicationId
        $exePath  = $script:ClaudeEnv.ClaudeExe
        $isMsix   = $script:ClaudeEnv.IsMsix

        if (-not $isMsix -and -not $exePath) {
            Log "Claude not found, skipping launcher" -Colour DarkGray -Indent
            Add-StepResult -Status Skipped -Detail "Claude Desktop was not found"
        } else {
            # Two shapes. An MSIX install needs the package-identity launch. A
            # traditional install is an ordinary exe and takes the flag straight
            # on the command line, no ceremony required.
            if ($isMsix -and $pfn -and $appId) {
                # In single quotes, a ' is written twice.
                $launchBody = @"
Invoke-CommandInDesktopPackage ``
    -PackageFamilyName '$($pfn -replace "'", "''")' ``
    -AppId '$($appId -replace "'", "''")' ``
    -Command 'app\Claude.exe' ``
    -Args '--disable-gpu'
"@
                $howNote = "MSIX install, launching inside package identity ($pfn)"
            } else {
                $launchBody = "Start-Process -FilePath '$($exePath -replace "'", "''")' -ArgumentList '--disable-gpu'"
                $howNote = "Traditional install, flag passed directly"
            }

            $ps1Content = @"
# Launch-Claude-NoGPU.ps1
# Auto-generated by Prevent-ClaudeIssues.ps1 v$ToolkitVersion
#
# Starts Claude Desktop with hardware GPU acceleration turned off, to avoid the
# fatal GPU-process crash triggered by the in-app Browser preview.
# Upstream issue: https://github.com/anthropics/claude-code/issues/80444
#
# $howNote
#
# Run this INSTEAD of the normal Claude shortcut when you need the Browser pane.
# The regular shortcut is unchanged and still uses the GPU.

`$ErrorActionPreference = 'Stop'

# Claude is single-instance. Launching a second copy while one is running just
# focuses the existing window, and the flag is silently ignored. So the running
# instance has to go first, or this does nothing at all.
#
# The app only. The Claude Code CLI runs as "claude" too: the copies Claude
# starts for Code sessions and project threads, out of a claude-code folder,
# and any terminal session. Closing those would end that work. A process whose
# path cannot be read is left alone as well: it could be either, and an app
# that was started elevated cannot be closed from here anyway. One whose path
# is read and is not the CLI's, but whose product name is not, is the app.
# -Unread lists instead the ones whose path could not be read.
#
# Get-Process says that none runs with an error of its own: id
# NoProcessFoundForGivenName, category ObjectNotFound, exception
# ProcessCommandException (measured 2026-11-22). Any other error means the
# list could not be read, and is thrown. Until 6.3.1 it read as none running.
function Get-ClaudeAppProcess {
    param([switch]`$Unread)
    `$all = @()
    try {
        `$all = @(Get-Process claude -ErrorAction Stop)
    } catch {
        if (-not ("`$(`$_.FullyQualifiedErrorId)" -like 'NoProcessFoundForGivenName,*' -and
                  [string]`$_.CategoryInfo.Category -eq 'ObjectNotFound' -and
                  `$_.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.ProcessCommandException')) { throw }
    }
    @(`$all | Where-Object {
        `$exe = `$null
        try { `$exe = `$_.MainModule.FileName } catch { `$null = `$_ }
        if (`$Unread) { return (-not `$exe) }
        if (-not `$exe -or `$exe -match '[\\/]claude-code[\\/]') { return `$false }
        `$product = ''
        try { `$product = [string](Get-Item -LiteralPath `$exe).VersionInfo.ProductName } catch { `$null = `$_ }
        `$product -ne 'Claude Code'
    })
}

# A list that cannot be read stops here: Claude is neither closed nor started.
try {
    `$running = @(Get-ClaudeAppProcess)
    `$unread = @(Get-ClaudeAppProcess -Unread)
} catch {
    Write-Host "  [!] The running processes could not be read, so Claude was not closed or started: `$(`$_.Exception.Message)" -ForegroundColor Red
    Read-Host "  Press Enter to close"
    exit 1
}
# So does a Claude process whose path cannot be read. It is not closed, as it
# could be the CLI, and it could as well be Claude started as administrator,
# which would take the launch below and ignore the flag. Until 6.3.1 the
# launch went ahead all the same, and was then said to run without the GPU,
# or not to have started.
if (`$unread.Count -gt 0) {
    Write-Host "  [!] `$(`$unread.Count) running Claude process(es) could not be read (one started as administrator, or a Claude Code session), so Claude was not closed or started without the GPU: a copy already running takes the launch and ignores the flag. Close Claude and run this again." -ForegroundColor Red
    Read-Host "  Press Enter to close"
    exit 1
}
if (`$running.Count -gt 0) {
    Write-Host "  Closing `$(`$running.Count) running Claude process(es)..." -ForegroundColor DarkGray
    # A stop that fails shows below, where what still runs is read again.
    `$running | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3
    # A copy still running would take the launch below and ignore the flag.
    # Until 6.3.1 this was not checked, and a Claude that could not be closed
    # was then reported as running without GPU acceleration.
    `$left = -1
    try { `$left = @(Get-ClaudeAppProcess).Count } catch { `$left = -1 }
    if (`$left -ne 0) {
        Write-Host "  [!] Claude could not be closed, or whether it closed could not be read, so it was not started again without the GPU. Close Claude and run this again." -ForegroundColor Red
        Read-Host "  Press Enter to close"
        exit 1
    }
}

Write-Host "  Starting Claude with --disable-gpu..." -ForegroundColor Cyan
$launchBody

Start-Sleep -Seconds 3
`$after = -1
try { `$after = @(Get-ClaudeAppProcess).Count } catch { `$after = -1 }
if (`$after -gt 0) {
    Write-Host "  Claude is running without hardware GPU acceleration." -ForegroundColor Green
    Write-Host "  The UI may feel slightly slower. That is the trade." -ForegroundColor DarkGray
} elseif (`$after -eq 0) {
    Write-Host "  [!] Claude did not start. Launch it normally and report this." -ForegroundColor Red
    Read-Host "  Press Enter to close"
} else {
    Write-Host "  [!] Claude was started, but the running processes could not be read to check that it runs." -ForegroundColor Red
    Read-Host "  Press Enter to close"
}
"@
            Set-Content -Path $noGpuPs1 -Value $ps1Content -Encoding ASCII -Force

            $cmdContent = @"
@echo off
REM Claude Desktop (no GPU) launcher
REM Auto-generated by Prevent-ClaudeIssues.ps1 v$ToolkitVersion
REM Avoids the fatal GPU-process crash from the in-app Browser preview.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Launch-Claude-NoGPU.ps1"
"@
            Set-Content -Path $noGpuCmd -Value $cmdContent -Encoding ASCII -Force

            Log "Launcher created: $noGpuCmd" -Colour Green -Indent
            Add-StepResult -Status Done
            Log $howNote -Colour DarkGray -Indent

            # Desktop shortcut, same treatment as the admin launcher.
            try {
                $desktopDir = [Environment]::GetFolderPath("Desktop")
                $noGpuLnk   = Join-Path $desktopDir "Claude (no GPU).lnk"
                $wshell = New-Object -ComObject WScript.Shell
                $lnk = $wshell.CreateShortcut($noGpuLnk)
                $lnk.TargetPath       = $noGpuCmd
                $lnk.WorkingDirectory = $ClaudeAppData
                $lnk.Description      = "Claude Desktop with hardware GPU acceleration disabled"
                $lnk.IconLocation = Initialize-ClaudeIcon
                $lnk.Save()
                Log "Shortcut created: $noGpuLnk" -Colour Green -Indent
                Add-StepResult -Status Done
            } catch {
                Log "[!] Could not create the shortcut, the .cmd still works: $($_.Exception.Message)" -Colour Yellow -Indent
                Add-StepResult -Status Warning -Detail "No desktop shortcut: $($_.Exception.Message). The .cmd launcher still works"
            }

            Log "Use it only if you hit the crash; the normal shortcut is unchanged" -Colour DarkGray -Indent
        }
    } catch {
        Log "[!] Could not create the GPU-free launcher: $($_.Exception.Message)" -Colour Red -Indent
        Add-StepResult -Status Failed -Detail "Could not create the GPU-free launcher: $($_.Exception.Message)"
    }
    Complete-Step

    # ----------------------------------------------------------------
    # 19. What older versions set up that is no longer needed
    # ----------------------------------------------------------------
    # Removed here because it now does harm or nothing:
    #   - The logon task $BootTaskName restarted vmcompute 45 seconds after
    #     every logon while Claude was not yet running, which is a restart of
    #     the service every VM on the machine depends on.
    #   - disable-dynamic-memory.flag asked the monitor to turn off Dynamic
    #     Memory on a Hyper-V VM named like Claude. The Cowork VM is not a
    #     Hyper-V VM and has no such setting.
    # The settings they made are listed, and -Undo removes them. Three of
    # them, values an administrator or a policy can set too, -Undo removes
    # only where an older version is known to have run, and the list says
    # which it is. What shows that is recorded before it is deleted here.
    Step 19 $steps "Removing what older versions set up..." -Name "Older versions' leftovers"
    try {
        # Until 6.3.1 a file or a task that could not be read counted as not
        # there: no older version, and no logon task to remove.
        $legacyTrace = $null
        $legacyUnread = $false
        try {
            $legacyTrace = Get-LegacyInstallTrace
        } catch {
            $legacyUnread = $true
            $said = "Whether a version before 6.3.0 ran here could not be told, because $(Get-ClaudeErrorText $_)"
            Log "[!] $said" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail $said
        }
        $bootTask = $null
        try {
            $bootTask = Get-ClaudeScheduledTaskOrNull -TaskName $BootTaskName -TaskPath $TaskPath
        } catch {
            $said = "Could not read the scheduled task $TaskPath$BootTaskName, so the logon task older versions created was not checked or removed: $(Get-ClaudeErrorText $_)"
            Log "[!] $said" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail $said
        }
        if ($null -ne $bootTask) {
            Add-LegacyInstallRecord -What 'boot-fix task' -Where "$TaskPath$BootTaskName"
            Unregister-ScheduledTask -TaskName $BootTaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction Stop
            Log "Removed the logon task $BootTaskName" -Colour Green -Indent
            Add-StepResult -Status Done
        }
        # Until 6.3.1 the flag and the Connected Standby record below were
        # looked for with Test-Path, which answers $false for a file that
        # cannot be read as for one that is not there (measured 2026-11-23),
        # and "Nothing left from older versions" was said.
        $flagFile = Join-Path $ClaudeAppData "disable-dynamic-memory.flag"
        if (Test-PreventPathThere -Path $flagFile -NotDone 'it was not recorded or removed') {
            Add-LegacyInstallRecord -What 'Dynamic Memory flag' -Where $flagFile
            Remove-PreventItem -Path $flagFile
            Log "Removed $flagFile" -Colour Green -Indent
            Add-StepResult -Status Done
        }

        $leftovers = @()
        $gated = @()
        $policyPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
        $v = Get-RegistryValue -Path $policyPath -Name 'LocalAccountTokenFilterPolicy'
        if ($v -is [int] -and $v -eq 1) { $gated += "LocalAccountTokenFilterPolicy = 1" }
        $v = Get-RegistryValue -Path $policyPath -Name 'FilterAdministratorToken'
        if ($v -is [int] -and $v -eq 0) { $gated += "FilterAdministratorToken = 0" }
        $v = Get-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Control' -Name 'ServicesPipeTimeout'
        if ($v -is [int] -and $v -eq 120000) { $gated += "ServicesPipeTimeout = 120000" }
        if ($null -ne $legacyTrace) {
            $leftovers += $gated
        } elseif ($legacyUnread) {
            foreach ($g in $gated) {
                $said = "$g matches what older versions set, but whether one ran here could not be told, so it is not known whether -Undo removes it"
                Log "Left in place: $said" -Colour DarkGray -Indent
                Add-StepResult -Status Unchanged -Detail $said
            }
        } else {
            foreach ($g in $gated) {
                $said = "$g matches what older versions set, but no older install was found here, so -Undo leaves it"
                Log "Left in place: $said" -Colour DarkGray -Indent
                Add-StepResult -Status Unchanged -Detail $said
            }
        }
        if (Test-PreventPathThere -Path $CsBackupFile -NotDone 'whether Connected Standby is left as an older version set it is not known') {
            $leftovers += "CsEnabled (Connected Standby)"
        }
        # Restart recovery older versions set on the Cowork VM service. Setup
        # sets recovery on vmcompute only, with a record, so this is theirs.
        # Until 6.3.1 a service that could not be read counted as not there.
        $cvSvc = $null
        if ($ServiceName) {
            try {
                $cvSvc = Get-ClaudeServiceOrNull -Name $ServiceName
            } catch {
                $said = "Could not read the service $ServiceName, so the restart recovery older versions set on it was not checked: $(Get-ClaudeErrorText $_)"
                Log "[!] $said" -Colour Yellow -Indent
                Add-StepResult -Status Warning -Detail $said
            }
        }
        if ($null -ne $cvSvc -and
            $null -eq (Get-PreventUndoItem -Kind ServiceFailureActions -Key $ServiceName) -and
            (Test-RestartRecovery (ConvertFrom-FailureActionsBlob -Bytes (Get-ServiceFailureBytes -Service $ServiceName)))) {
            $leftovers += "the restart recovery on $ServiceName"
        }
        # The NAT rule older versions created: listed, and left for -Undo.
        try {
            if (Test-CoworkNat) { $leftovers += "the NAT rule CoworkNAT" }
        } catch {
            Log "[!] Could not read the NAT rules: $($_.Exception.Message)" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail "Could not read the NAT rules, so the rule CoworkNAT older versions created was not checked: $($_.Exception.Message)"
        }
        # A list that cannot be read is said, as the NAT rules are: passed
        # over, it read as no adapter an older version touched, and the step
        # said nothing was left.
        $legacyNics = @()
        try {
            $legacyNics = @(Get-NetAdapter -Physical -ErrorAction Stop)
        } catch {
            Log "[!] Could not list the network adapters: $($_.Exception.Message)" -Colour Yellow -Indent
            Add-StepResult -Status Warning -Detail "Could not list the network adapters, so the wake settings older versions turned off on them were not checked: $($_.Exception.Message)"
        }
        foreach ($nic in $legacyNics) {
            $devPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\$($nic.PnPDeviceID)\Device Parameters"
            if ((Get-RegistryValue -Path $devPath -Name 'PnPCapabilities') -eq 24) {
                $leftovers += "PnPCapabilities with the wake settings on $($nic.Name)"
            }
        }
        foreach ($lo in $leftovers) {
            Log "Left in place: $lo (-Undo removes it)" -Colour DarkGray -Indent
            Add-StepResult -Status Unchanged -Detail "$lo, set by an older version, is left in place; -Undo removes it"
        }
        # Dynamic Memory older versions turned off, where one is known to have
        # run. -Undo turns it back on, and Hyper-V lets it only while the VM
        # is off.
        # Get-Command answers nothing for a command that is not installed:
        # without Hyper-V's cmdlets there is no Dynamic Memory to read.
        if ($null -ne $legacyTrace -and (Get-Command -Name Set-VMMemory -ErrorAction SilentlyContinue)) {
            $dmVms = @()
            try {
                $dmVms = @(Get-ClaudeVmWithoutDynamicMemory)
            } catch {
                Log "[!] Could not read the Hyper-V VMs: $($_.Exception.Message)" -Colour Yellow -Indent
                Add-StepResult -Status Warning -Detail "Could not read the Hyper-V VMs, so Dynamic Memory an older version turned off was not checked: $($_.Exception.Message)"
            }
            foreach ($vm in $dmVms) {
                $said = "Dynamic Memory on the VM $($vm.Name), turned off by an older version, is left off; -Undo turns it back on"
                if ([string]$vm.State -ne 'Off') { $said += ", once the VM is turned off: it is $($vm.State) now" }
                Log "Left in place: $said" -Colour DarkGray -Indent
                Add-StepResult -Status Unchanged -Detail $said
            }
        }
        if ($script:CurrentStep.Status -eq 'None') {
            Log "Nothing left from older versions" -Colour DarkGray -Indent
            Add-StepResult -Status Unchanged
        }
    } catch {
        Log "[!] $($_.Exception.Message)" -Colour Red -Indent
        Add-StepResult -Status Failed -Detail $_.Exception.Message
    }

    # ================================================================
    # Summary
    # ================================================================
    if ((Get-PreventExitCode) -eq 0) {
        Write-Box -Text "SETUP COMPLETE" -Colour Green
    } else {
        Write-Box -Text "SETUP FINISHED WITH ERRORS" -Colour Red
    }
    Write-StepSummary
    Write-Host ""
    Write-Host "  TIP: Right-click 'Fix Claude Desktop' in Start Menu" -ForegroundColor DarkGray
    Write-Host "       and select 'Pin to taskbar' for quick access." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  To put back everything this script changed:" -ForegroundColor DarkGray
    Write-Host "    .\Prevent-ClaudeIssues.ps1 -Undo" -ForegroundColor DarkGray
    if ($script:UndoRecord.Count -gt 0) {
        Write-Host ""
        Write-Host "  What it changed, and the values it replaced, are recorded in:" -ForegroundColor DarkGray
        Write-Host "    $UndoFile" -ForegroundColor DarkGray
    }
    if ($script:UndoForeignAside) {
        Write-Host ""
        Write-Host "  An undo record written on another computer was found. It was not" -ForegroundColor Yellow
        Write-Host "  used, and it is kept as:" -ForegroundColor Yellow
        Write-Host "    $script:UndoForeignAside" -ForegroundColor Yellow
    }
}

} catch {
    $script:Aborted = $true
    $msg = $_.Exception.Message
    if ($null -ne $script:CurrentStep) { Add-StepResult -Status Failed -Detail "Stopped here: $msg" }
    Write-Box -Text "UNEXPECTED ERROR" -Colour Red
    Write-Host "  $msg" -ForegroundColor Red
    Write-Host "  Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor DarkGray
    if ($null -ne $script:CurrentStep) {
        Write-Host "  Step: $($script:CurrentStep.Num). $($script:CurrentStep.Name)" -ForegroundColor DarkGray
    }
    Write-Host ""
    if ($script:StepResults.Count -eq 0) {
        Write-Host "  It stopped before the first step, so nothing was changed." -ForegroundColor DarkYellow
    } elseif ($Undo) {
        Write-Host "  -Undo stopped early. What it put back stays put back; what it" -ForegroundColor DarkYellow
        Write-Host "  did not reach is still in the record, so rerun -Undo once the" -ForegroundColor DarkYellow
        Write-Host "  cause above is resolved." -ForegroundColor DarkYellow
    } else {
        Write-Host "  Setup stopped early, so the steps after this one did not run." -ForegroundColor DarkYellow
        Write-Host "  What the earlier steps changed is in place and recorded, so" -ForegroundColor DarkYellow
        Write-Host "  -Undo can put it back. Rerun the script once the cause above" -ForegroundColor DarkYellow
        Write-Host "  is resolved." -ForegroundColor DarkYellow
    }
    Write-Host "  If it keeps happening, open an issue at" -ForegroundColor DarkGray
    Write-Host "  https://github.com/JesperLive/ClaudeFix/issues" -ForegroundColor DarkGray
    if ($script:StepResults.Count -gt 0) {
        Write-Host ""
        Write-StepSummary
    }
}

# 1 when a step failed or the run stopped, so Prevent-ClaudeIssues.bat and
# anything else that runs this can tell. Until 6.3.0 it was 0 either way.
$exitCode = Get-PreventExitCode
Write-Host ""
Write-Host "  Press any key to close..." -ForegroundColor DarkGray
# Best effort: without a console to read a key from, the window simply closes.
try { [void][System.Console]::ReadKey($true) } catch { $null = $_ }
exit $exitCode
