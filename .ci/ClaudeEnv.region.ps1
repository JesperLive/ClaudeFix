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
