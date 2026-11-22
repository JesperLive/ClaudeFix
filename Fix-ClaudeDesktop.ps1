<#
.SYNOPSIS
    Claude Desktop / Cowork, Reset and Fix

.DESCRIPTION
    Kills the Claude Desktop app's processes, stops CoworkVMService, recovers
    from HCS (Host Compute Service) errors, performs orphan compute system
    cleanup, purges stale VM cache, restarts the service, and relaunches
    Claude Desktop with elevated privileges when available. Claude Code CLI
    processes, which are also named "claude", are left running: they carry
    project threads, Code sessions and terminal sessions.

    Use -Close for a clean shutdown without relaunching (kills Claude UI,
    waits for VM shutdown, restarts service for next launch). Pair with
    Stop-ClaudeDesktop.bat for double-click convenience.

    Does NOT touch config files or MCP servers. Session transcripts are left
    alone unless you explicitly pass -PurgeSessions.
    Fully automatic, no user interaction required.

    Works with or without admin privileges. If run without admin,
    service control falls back to process-level operations and Claude
    handles service restart automatically on launch.

.PARAMETER SkipLaunch
    Reset the VM service but don't relaunch Claude Desktop afterwards.

.PARAMETER BootPrep
    Retired in 6.3.0: says that it does nothing and exits 0, before anything
    else runs, so it writes, deletes, elevates and touches nothing: no log,
    no cleanup, no vmcompute, no VM. Until 6.3.0 it restarted vmcompute, from
    a task older versions of Prevent ran at every logon. The switch stays so
    that task, until Prevent or -Undo removes it, gets a no-op rather than an
    error.

.PARAMETER Quiet
    Suppress the "press any key" prompt at the end and skip the
    interactive menu. Defaults to Smart mode.

.PARAMETER Mode
    Skip the interactive menu and run in the specified mode:
      Quick      : Restart services + basic repair (Steps 1-5, skip cache purge)
      Deep       : Full nuclear reset (all steps including cache purge)
      Smart      : Try quick first, escalate to deep if needed (default)
      Diagnostic : Health check only, no changes

.PARAMETER KeepCache
    Skip the VM cache purge (Step 6). Use this to avoid re-downloading the VM
    bundle, which is roughly 13 GB on a current install, not the 2-3 GB this
    help used to claim. rootfs.vhdx alone is about 9 GB.

    sessiondata.vhdx and smol-bin.vhdx are backed up and restored across a
    purge, so your workspace state survives. rootfs.vhdx is not: rebuilding the
    VM image is the entire point of a Deep purge, and restoring it would put
    back whatever was wrong with it. That is the 9 GB you re-download.

    If the fix fails with -KeepCache, run again without it to force a clean
    rebuild.

.PARAMETER Close
    Perform a clean shutdown only: kill Claude UI, wait for VM to shut down
    gracefully, restart the service so it is ready for next launch. Does NOT
    relaunch Claude. Useful before a reboot or when you just want to fully
    stop Claude without running a repair.

.PARAMETER PurgeSessions
    Also delete session transcripts under local-agent-mode-sessions that are
    older than 7 days. Off by default. Earlier versions did this on every run
    in every mode while the description above claimed conversations were never
    touched, so it is now opt-in and stated plainly.

.PARAMETER WhatIf
    Show what would happen without actually doing anything.

.NOTES
    Version : 6.3.0
    Author  : Jesper Driessen
    Licence : MIT
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$SkipLaunch,
    [switch]$BootPrep,
    [Alias("Silent")]
    [switch]$Quiet,
    [switch]$KeepCache,
    [switch]$Close,
    [switch]$PurgeSessions,
    [ValidateSet('Quick','Deep','Smart','Diagnostic')]
    [string]$Mode
)

# -- -BootPrep: retired ----------------------------------------------
# Until 6.3.0 -BootPrep restarted vmcompute when no Cowork workspace was up,
# from a task older versions of Prevent ran 45 seconds after every logon. A
# vmcompute restart is a restart of the service every VM on the machine
# depends on. The task was retired in 6.3.0, and Prevent and -Undo remove it.
# Until they run, the task still passes -BootPrep, so the switch is kept and
# does nothing but say so.
#
# First, before anything else: the elevation below would raise a UAC prompt
# at logon, and what follows it saves the environment, prunes old logs and
# VHDX backups and writes a log. Until 6.3.0 all of that ran first, for a run
# that was to do nothing.
if ($BootPrep) {
    Write-Host "  [BootPrep] -BootPrep is retired since 6.3.0 and does nothing; vmcompute and the VMs were not touched" -ForegroundColor DarkGray
    exit 0
}

# -- Admin elevation (optional) --------------------------------------
$script:IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $script:IsAdmin) {
    $scriptFile = $PSCommandPath
    if (-not $scriptFile) { $scriptFile = $MyInvocation.MyCommand.Definition }

    $elevateArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptFile`""
    if ($SkipLaunch)       { $elevateArgs += " -SkipLaunch" }
    if ($Quiet)            { $elevateArgs += " -Quiet" }
    if ($WhatIfPreference) { $elevateArgs += " -WhatIf" }
    if ($KeepCache)        { $elevateArgs += " -KeepCache" }
    if ($Close)            { $elevateArgs += " -Close" }
    if ($PurgeSessions)    { $elevateArgs += " -PurgeSessions" }
    if ($Mode)             { $elevateArgs += " -Mode $Mode" }

    Write-Host ""
    Write-Host "  Requesting admin privileges for full service control..." -ForegroundColor DarkGray
    Write-Host "  (If you decline, the script will still work but some" -ForegroundColor DarkGray
    Write-Host "   operations may be slower or less thorough.)" -ForegroundColor DarkGray
    Write-Host ""

    try {
        # -PassThru so the child's exit code can be propagated. This used to be
        # a hardcoded "exit 0", so no caller could tell a clean repair from a
        # hard failure, or either from a run that was held back.
        $elevated = Start-Process PowerShell -ArgumentList $elevateArgs -Verb RunAs -Wait -PassThru
        if ($elevated) { exit $elevated.ExitCode }
        exit 0
    } catch {
        Write-Host "  [i] Running without admin, service control will be limited" -ForegroundColor Yellow
        Write-Host ""
        # Continue running as normal user
    }
}

# -- Running (elevated or not) ---------------------------------------
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
function Get-ClaudeEnvironment {
    <#
    .SYNOPSIS
        Discovers where this machine's Claude Desktop install actually lives.
    .DESCRIPTION
        Every location the toolkit acts on is queried at run time rather than
        guessed at authoring time. Nothing here throws. Anything that cannot be
        determined stays $null and is recorded in Warnings, so a partial result
        is still usable by a caller that only needs some of the fields.

        Locations are discoverable and are discovered. Thresholds and log
        signatures are NOT in scope here: those are not location problems and
        cannot be fixed by looking harder at the filesystem.
    .PARAMETER SkipCacheInventory
        Skip the recursive walk of the VM cache. That walk is the expensive half
        of discovery and only the backup and purge paths consume its results, so
        anything that polls should pass this. CacheSizeBytes is left at -1 when
        skipped, which is distinguishable from a genuine zero.
    #>
    param([switch]$SkipCacheInventory)

    # Version 2 (6.2.0): the app's folders come from Resolve-ClaudeDataPaths
    # instead of %APPDATA%\Claude. A version 1 record names the old folders,
    # so Import-ClaudeEnvironment turns it away.
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

        Warnings = @()
    }

    # -- Host capabilities ------------------------------------------------
    try {
        $info.IsAdmin = ([Security.Principal.WindowsPrincipal] `
            [Security.Principal.WindowsIdentity]::GetCurrent()
            ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { $info.Warnings += "Could not determine admin status: $($_.Exception.Message)" }

    # TickCount64 is .NET Core and up. Windows PowerShell 5.1 runs on .NET
    # Framework and does not have it, which is why the idle-time maths has to
    # survive the 24.9 day TickCount wrap instead of just using the 64-bit call.
    try { $null = [Environment]::TickCount64; $info.HasTickCount64 = $true }
    catch { $info.HasTickCount64 = $false }

    try {
        $hcs = Join-Path $env:SystemRoot "System32\hcsdiag.exe"
        if (Test-Path $hcs) { $info.HcsDiagPath = $hcs }
        else { $info.Warnings += "hcsdiag.exe not present; HCS cleanup unavailable" }
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
    try {
        $pkgs = @(Get-AppxPackage -Name "Claude*" -ErrorAction SilentlyContinue |
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
            # there. PackageFullName carries the same version as a second source.
            try { $info.PackageVersion = "$($p.Version)" } catch { $null = $_ }
            if (-not $info.PackageVersion -and $info.PackageFullName) {
                try {
                    $parts = @($info.PackageFullName -split '_')
                    if ($parts.Count -ge 2) { $info.PackageVersion = $parts[1] }
                } catch { $null = $_ }
            }
        }
    } catch { $info.Warnings += "Package query failed: $($_.Exception.Message)" }

    # Application Id comes from the manifest. The old hardcoded "App" fallback
    # produced an invalid AUMID, and it only ever ran when the manifest read had
    # already failed, which is precisely when a wrong guess does damage.
    if ($info.InstallLocation) {
        try {
            $manifestPath = Join-Path $info.InstallLocation "AppxManifest.xml"
            if (Test-Path $manifestPath) {
                [xml]$mx = Get-Content $manifestPath -ErrorAction Stop
                $ns = New-Object Xml.XmlNamespaceManager($mx.NameTable)
                $ns.AddNamespace("x", "http://schemas.microsoft.com/appx/manifest/foundation/windows10")
                $appNode = $mx.SelectSingleNode("//x:Application", $ns)
                if ($appNode) { $info.ApplicationId = $appNode.GetAttribute("Id") }
            }
        } catch { $info.Warnings += "Manifest read failed: $($_.Exception.Message)" }
    }
    if (-not $info.ApplicationId -and $info.PackageFound) {
        # Every shipped Claude manifest to date uses Id="Claude", matching the
        # package Name. Better than a literal guess, and it is recorded as one.
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
    } catch { $info.Warnings += "Data folder discovery failed: $($_.Exception.Message)" }

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
    if ($info.InstallLocation) {
        foreach ($rel in @('app\claude.exe', 'claude.exe')) {
            try {
                $cand = Join-Path $info.InstallLocation $rel
                if (Test-Path $cand) {
                    $info.ClaudeExe = $cand; $info.ClaudeExeSource = 'package'; break
                }
            } catch { $null = $_ }
        }
    }
    if (-not $info.ClaudeExe) {
        try {
            foreach ($rp in @(Get-Process -Name $info.ProcessName -ErrorAction SilentlyContinue)) {
                try {
                    $cand = $rp.MainModule.FileName
                    if ($cand -and -not (Test-ClaudeCodeCliPath $cand) -and (Test-Path $cand)) {
                        $info.ClaudeExe = $cand; $info.ClaudeExeSource = 'process'; break
                    }
                } catch { $null = $_ }
            }
        } catch { $null = $_ }
    }
    if (-not $info.ClaudeExe) {
        try {
            foreach ($wp in @(Get-CimInstance Win32_Process -Filter "Name LIKE '%claude%'" `
                              -ErrorAction SilentlyContinue)) {
                $cand = "$($wp.ExecutablePath)"
                if ($cand -and -not (Test-ClaudeCodeCliPath $cand) -and (Test-Path $cand)) {
                    $info.ClaudeExe = $cand; $info.ClaudeExeSource = 'wmi'; break
                }
            }
        } catch { $null = $_ }
    }
    if (-not $info.ClaudeExe) {
        foreach ($cand in @(
            (Join-Path $env:LOCALAPPDATA "Programs\claude\Claude.exe"),
            (Join-Path $env:LOCALAPPDATA "AnthropicClaude\Claude.exe"),
            (Join-Path $env:LOCALAPPDATA "Anthropic\Claude\Claude.exe"),
            (Join-Path $env:ProgramFiles "Claude\Claude.exe")
        )) {
            try { if ($cand -and (Test-Path $cand)) {
                $info.ClaudeExe = $cand; $info.ClaudeExeSource = 'commonpath'; break
            } } catch { $null = $_ }
        }
    }
    if (-not $info.ClaudeExe) { $info.Warnings += "Claude executable not located" }

    # -- Service -------------------------------------------------------------
    try {
        $svc = Get-Service -Name $info.ServiceName -ErrorAction SilentlyContinue
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
                        } catch { $null = $_ }
                    }
                }
            } catch {
                try { $info.ServiceStartMode = "$($svc.StartType)" } catch { $null = $_ }
            }
        } else {
            $info.Warnings += "$($info.ServiceName) is not registered"
        }
    } catch { $info.Warnings += "Service query failed: $($_.Exception.Message)" }

    # Recovery actions are read from the registry, not from parsed sc.exe output.
    # Matching the literal string RESTART in console text fails on any Windows
    # whose display language is not English.
    try {
        $svcKey = "HKLM:\SYSTEM\CurrentControlSet\Services\$($info.ServiceName)"
        if (Test-Path $svcKey) {
            $fa = (Get-ItemProperty -Path $svcKey -Name FailureActions `
                   -ErrorAction SilentlyContinue).FailureActions
            if ($fa) { $info.ServiceRecoveryConfigured = $true }
        }
    } catch { $null = $_ }

    # -- VM log selection ----------------------------------------------------
    # See Get-ClaudeVmLogCandidates for which logs count and why.
    try {
        if ($info.LogDir -and (Test-Path $info.LogDir)) {
            $cands = @(Get-ClaudeVmLogCandidates -LogDir $info.LogDir)
            $info.VmLogCandidates = $cands
            if ($cands.Count -gt 0) { $info.VmLogFile = $cands[0].Path }
            else { $info.Warnings += "No VM log present in $($info.LogDir)" }
        }
    } catch { $info.Warnings += "VM log probe failed: $($_.Exception.Message)" }

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
            foreach ($dir in @($info.VmCachePath, $info.BundlePath)) {
                if (-not $dir -or -not (Test-Path $dir)) { continue }
                foreach ($f in @(Get-ChildItem $dir -Recurse -File -ErrorAction SilentlyContinue)) {
                    $total += $f.Length
                    if ($f.Extension -eq '.vhdx') {
                        $vhdx += [pscustomobject]@{
                            Name   = $f.Name
                            Path   = $f.FullName
                            Length = $f.Length
                        }
                    }
                }
            }
            $info.VhdxFiles      = @($vhdx)
            $info.CacheSizeBytes = $total
        } catch { $info.Warnings += "Cache inventory failed: $($_.Exception.Message)" }
    }

    return $info
}

function Save-ClaudeEnvironment {
    <#
    .SYNOPSIS
        Persists a discovery result so sibling scripts and bug reports can use it.
    .DESCRIPTION
        Written as UTF-8 without a BOM. Out-File -Encoding utf8 emits a BOM on
        Windows PowerShell 5.1, which breaks anything reading the file as plain
        JSON. Never throws: persistence is a convenience, not a dependency.
    #>
    param(
        [Parameter(Mandatory)]$Environment,
        [Parameter(Mandatory)][string]$Path
    )
    try {
        $dir = Split-Path $Path -Parent
        if ($dir -and -not (Test-Path $dir)) {
            New-Item $dir -ItemType Directory -Force | Out-Null
        }
        $json = $Environment | ConvertTo-Json -Depth 4
        [System.IO.File]::WriteAllText($Path, $json, `
            (New-Object System.Text.UTF8Encoding($false)))
        return $true
    } catch { return $false }
}

function Import-ClaudeEnvironment {
    <#
    .SYNOPSIS
        Reads a persisted discovery result, but only if it still describes this
        machine's current install.
    .DESCRIPTION
        Returns $null when the file is missing, unreadable, from a different
        schema, or stamped with a package version other than the one asked for.
        That version stamp is the whole point: the existing .claude-exe-path
        cache has no stamp, which is why a stale entry there survives until
        something downstream trips over it.

        Pass -ExpectedPackageVersion from a live query when correctness matters.
        Omit it to accept whatever was recorded, which is fine for bug reports.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$ExpectedPackageVersion
    )
    try {
        if (-not (Test-Path $Path)) { return $null }
        $raw = Get-Content $Path -Raw -ErrorAction Stop
        if (-not $raw) { return $null }
        $obj = $raw | ConvertFrom-Json -ErrorAction Stop
        if (-not $obj) { return $null }
        if ($obj.SchemaVersion -ne 2) { return $null }
        if ($ExpectedPackageVersion -and
            "$($obj.PackageVersion)" -ne "$ExpectedPackageVersion") { return $null }
        return $obj
    } catch { return $null }
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
    #>
    param([string]$Path)
    if (-not $Path) { return $false }
    if ($Path -match '[\\/]claude-code[\\/]') { return $true }
    try {
        $vi = (Get-Item -LiteralPath $Path -ErrorAction Stop).VersionInfo
        if ($vi -and "$($vi.ProductName)" -eq 'Claude Code') { return $true }
    } catch { $null = $_ }
    return $false
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
    } catch { $null = $_ }

    $result = @()
    $cimPaths = $null
    $products = @{}
    foreach ($p in @(Get-Process -Name $Environment.ProcessName -ErrorAction SilentlyContinue)) {
        $path = $null
        try { $path = $p.MainModule.FileName } catch { $null = $_ }
        if (-not $path) {
            # MainModule needs more access than the image name does. One CIM
            # query answers for every process it could not read.
            if ($null -eq $cimPaths) {
                $cimPaths = @{}
                try {
                    $flt = "Name='$($Environment.ProcessName).exe'"
                    foreach ($w in @(Get-CimInstance Win32_Process -Filter $flt -ErrorAction Stop)) {
                        if ($w.ExecutablePath) { $cimPaths[[int]$w.ProcessId] = "$($w.ExecutablePath)" }
                    }
                } catch { $null = $_ }
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
                    if (-not $products.ContainsKey($path)) {
                        $products[$path] = ''
                        try {
                            $products[$path] = "$((Get-Item -LiteralPath $path -ErrorAction Stop).VersionInfo.ProductName)"
                        } catch { $null = $_ }
                    }
                    if ($products[$path] -eq 'Claude Code') { $kind = 'cli' }
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
        Get-ClaudeProcessInventory says how the two are told apart.
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
        counted: a repair does not end it.

        Throws when the parent ids cannot be read, so a caller deciding
        whether a repair may run can fail closed.
    .PARAMETER Environment
        A Get-ClaudeEnvironment result.
    #>
    param([Parameter(Mandatory)]$Environment)

    $inventory = @(Get-ClaudeProcessInventory -Environment $Environment)
    $cli = @($inventory | Where-Object { $_.Kind -eq 'cli' })
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
    #>
    param([int]$Minutes = 35)

    $cfgDir = $env:CLAUDE_CONFIG_DIR
    if (-not $cfgDir -and $env:USERPROFILE) { $cfgDir = Join-Path $env:USERPROFILE ".claude" }
    if (-not $cfgDir) { return $false }
    $projects = Join-Path $cfgDir "projects"
    if (-not (Test-Path -LiteralPath $projects -PathType Container)) { return $false }
    $now = [datetime]::UtcNow
    $recent = Get-ChildItem -LiteralPath $projects -Filter "*.jsonl" -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { ($now - $_.LastWriteTimeUtc).TotalMinutes -lt $Minutes } |
        Select-Object -First 1
    return [bool]$recent
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

        Cheap enough to call on every poll: Test-Path and one directory listing
        per candidate, nothing recursive. Never throws.
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
    foreach ($b in $bases) {
        foreach ($n in @('Claude-Data', 'Claude-3p')) {
            try {
                $parent = Join-Path $b.Path $n
                if (Test-Path -LiteralPath $parent -PathType Container) {
                    foreach ($d in @(Get-ChildItem -LiteralPath $parent -Directory -Force -ErrorAction SilentlyContinue)) {
                        $wanted += @{ Path = $d.FullName; Source = "$($b.Tag)\$n" }
                    }
                }
            } catch { $null = $_ }
        }
        $wanted += @{ Path = (Join-Path $b.Path 'Claude'); Source = "$($b.Tag)\Claude" }
    }

    $markers = @('Preferences', 'Local State', 'config.json', 'Network', 'Session Storage',
                 'window-state.json', 'lockfile')
    $roots = @()
    $logs  = @()
    $seen  = @{}
    foreach ($w in $wanted) {
        try {
            $full = "$($w.Path)".TrimEnd('\')
            if (-not $full -or $seen.ContainsKey($full.ToLowerInvariant())) { continue }
            $seen[$full.ToLowerInvariant()] = $true
            if (-not (Test-Path -LiteralPath $full -PathType Container)) { continue }

            $newest = $null
            $isApp  = $false
            foreach ($m in $markers) {
                $mi = Get-Item -LiteralPath (Join-Path $full $m) -Force -ErrorAction SilentlyContinue
                if (-not $mi) { continue }
                if ($m -eq 'Preferences' -or $m -eq 'Local State') { $isApp = $true }
                if ($null -eq $newest -or $mi.LastWriteTime -gt $newest) { $newest = $mi.LastWriteTime }
            }
            $roots += [pscustomobject]@{ Path = $full; Source = $w.Source; IsAppData = $isApp; Newest = $newest }

            $ld = Join-Path $full 'logs'
            if (Test-Path -LiteralPath $ld -PathType Container) {
                $lnewest = $null
                foreach ($f in @(Get-ChildItem -LiteralPath $ld -Filter '*.log' -File -Force -ErrorAction SilentlyContinue)) {
                    if ($null -eq $lnewest -or $f.LastWriteTime -gt $lnewest) { $lnewest = $f.LastWriteTime }
                }
                $logs += [pscustomobject]@{ Path = $ld; Source = $w.Source; Newest = $lnewest }
            }
        } catch { $null = $_ }
    }

    $byNewest = @{ Expression = { if ($_.Newest) { $_.Newest } else { [datetime]::MinValue } }; Descending = $true }
    return [pscustomobject]@{
        DataRoots = @($roots | Sort-Object -Property @{ Expression = { [int]$_.IsAppData }; Descending = $true }, $byNewest)
        LogDirs   = @($logs | Sort-Object -Property $byNewest)
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
        when the log folder moves under it. Never throws.
    #>
    param([string]$LogDir)

    $cands = @()
    if (-not $LogDir) { return $cands }
    foreach ($n in @('cowork_vm_node.log', 'coworkd.log', 'cowork-service.log')) {
        try {
            $p = Join-Path $LogDir $n
            if (Test-Path -LiteralPath $p) {
                $fi = Get-Item -LiteralPath $p -ErrorAction SilentlyContinue
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
        } catch { $null = $_ }
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
        the script started could be reading the quiet one. Never throws.
        Measured in UTC, so a daylight-saving change does not move it an hour.
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
# Everything below is read off this machine rather than assumed. See the
# ClaudeEnv region above for what is discovered, and for what deliberately
# is not: thresholds and log signatures are not location problems.
$script:ClaudeEnv = Get-ClaudeEnvironment

# -- Constants -------------------------------------------------------
$ToolkitVersion  = "6.3.0"
$ServiceName     = $script:ClaudeEnv.ServiceName
$ServiceExe      = $script:ClaudeEnv.ServiceProcessName
$ProcessName     = $script:ClaudeEnv.ProcessName
# This toolkit's folder, %APPDATA%\Claude: fix-logs, the exe cache, the VHDX
# backup. The app's own cache folders are under the discovered data root,
# which on app 6.18355 is %LOCALAPPDATA%\Claude-Data\claude_<id>.
$ClaudeAppData   = $script:ClaudeEnv.AppDataDir
$VmCachePath     = $script:ClaudeEnv.VmCachePath
$BundlePath      = $script:ClaudeEnv.BundlePath

# Without the toolkit's folder there is nowhere to write a transcript or to
# back the session data up to before a purge, so there is nothing useful left
# to do.
if (-not $ClaudeAppData) {
    Write-Host ""
    Write-Host "  [!] Could not locate the per-user Claude folder under %APPDATA%." -ForegroundColor Red
    Write-Host "      Nothing this script does is meaningful without it." -ForegroundColor DarkGray
    Write-Host ""
    exit 1
}

$ExePathCache    = Join-Path $ClaudeAppData ".claude-exe-path"
$LogDir          = Join-Path $ClaudeAppData "fix-logs"
$EnvArtifact     = Join-Path $LogDir "claude-env.json"
$ServiceTimeout  = 30   # VM shutdown takes 10-30s; too short = force-kill = HCS corruption
$StartPollMax    = 20   # seconds budget for the start poll, stepped 2s at a time
$MaxRetries      = 3    # how many times to retry the full fix cycle

# Seeded from discovery, which resolves the DESKTOP app rather than whichever
# process named "claude" answered first. The Claude Code CLI also runs under
# that name, and caching its path turns every later relaunch into a CLI start.
$script:CapturedClaudeExe = $script:ClaudeEnv.ClaudeExe
$script:ServicePrereq     = $null  # set by Test-CoworkServicePrereq

# Process exit code. 0 means the run finished without an unhandled error and
# without exhausting its retries. Anything scheduling this script can act on it.
$script:ExitCode = 0

# Exit code of a run that was held back before anything that stops the app or
# the service had run: another Fix run holds the lock; or, in a -Quiet run,
# the gate at the start (Get-AutomatedFixBlocker) found Claude or the PC in
# use (a Claude Code session or a recent transcript, claude using CPU, a recent
# VM log, or keyboard or mouse input in the last 3 minutes), or the check
# right before Step 1's kill found a Claude Code session or could not rule one
# out. Step 0 may already have removed a
# stale compute system by then. No other path uses it. Until 6.3.0 these
# exited 0, the same as a finished repair, so the monitor logged a run that
# did not repair anything as a completed fix, counted it toward "Auto-fix
# failed 3 times" and dropped what it had detected. The monitor reads this
# same number as FixHeldBackExitCode: the run is not an attempt, and a repair
# for a trigger that does not fire again, a log error or an Event Log error,
# is kept pending and tried again once the hold has cleared, and so is one
# for such a trigger that fires during the wait the hold leaves. Each run
# held back in a row doubles the monitor's wait before its next repair, from
# 2 minutes up to 60, until a run that is not held back, a run by hand that
# gets as far as the repair, which also ends that wait 5 minutes after it at
# the latest, or a hold 4 hours or more after the one before.
# A Fix run by hand ends that pending repair only when it gets as far as the
# repair: see $script:RepairUnderWayLine.
$script:BlockedExitCode = 3

# The line a run logs once it is past every check that holds it back with
# exit 3 and has begun the repair: after the menu, the Diagnostic branch, the
# "Proceed?" confirmation, the -Quiet gate and the check right before Step
# 1's kill. The monitor ends a held-back repair it keeps pending only for a
# Fix log written since its hold that holds this line, so a run that changed
# nothing, such as a Diagnostic run, a Cancel or a No at "Proceed?", does not
# end it. Such a log written since the latest of its runs held back with
# exit 3 also starts its count of runs held back in a row again, and ends
# the longer wait they left. A -WhatIf
# run changes nothing and does not log it. It must match
# $FixRepairLine in Watch-ClaudeHealth.ps1, which a test checks.
$script:RepairUnderWayLine = 'Repair under way: this run is past every check that holds it back'

# Exit code of a -Quiet run that stopped part-way: Steps 1 and 2 had run, so
# the app was stopped if it was running and so was the service, and then the
# check before Step 4's kill of the app's leftover processes, before the
# no-progress escalation or before the escalation to Deep found a Claude Code
# session, or could not rule one out. The check starts the service again
# first when this run stopped it. No other path uses this code. It is not 3,
# because 3 tells the monitor that the app and the service were left alone,
# and the monitor then keeps what it detected and, for a trigger that does not
# fire again, tries the repair again once the hold has cleared.
# This run did stop them, so the monitor counts it as an attempt at a repair,
# and, as after any run that was not held back, its wait before the next
# repair starts again from 2 minutes.
# The monitor reads this same number as FixStoppedPartWayExitCode.
$script:StoppedPartWayExitCode = 4

# Set when this run stops CoworkVMService: in Step 2, and in the restarts and
# recoveries after it. A -Quiet run that then stops part-way starts the
# service again before it exits, because Windows does not start a manually
# stopped service through its named-pipe trigger, and Cowork would otherwise
# stay down until the next repair, a full cooldown of the monitor later.
$script:ServiceStoppedByRun = $false

# Every Cowork compute system GUID any "hcsdiag list" showed during this run,
# added by Get-HcsList and never cleared. Step 5 matches vmwp.exe workers
# against it: a compute system leaves the list when it is killed, and a worker
# left hanging can then only be matched to it by a GUID seen earlier.
$script:CoworkGuidsSeen = @{}
# Set by Close-StaleHcsVms: whether hcsdiag answered its list, how many Cowork
# compute systems it listed, and how many kills it sent. A return of 0 alone
# cannot tell "nothing to kill" from "listed, no kill sent" from "kills sent,
# none confirmed" from "no answer".
$script:LastHcsListOk    = $false
$script:LastHcsListed    = 0
$script:LastHcsKillsSent = 0
# The last non-zero exit Invoke-HcsDiag logged, so a failure that repeats is
# not logged again on every pass of Step 9's wait.
$script:LastHcsDiagNote  = $null

# -- Logging ---------------------------------------------------------
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }
$script:SessionTimestamp = "{0:yyyyMMdd_HHmmss}" -f (Get-Date)
$LogFile = Join-Path $LogDir "fix_$($script:SessionTimestamp).log"

# Persist the discovery result next to the logs. Sibling scripts reuse it, the
# package version stamp lets them tell a current record from one written before
# an app update, and a bug report can attach this instead of asking someone to
# run ten commands by hand.
$null = Save-ClaudeEnvironment -Environment $script:ClaudeEnv -Path $EnvArtifact

# Clean up logs older than 30 days. Measured in UTC, like every age in this
# script: local time jumps an hour when summer time starts or ends.
try {
    Get-ChildItem $LogDir -Filter "fix_*.log" -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTimeUtc -lt [datetime]::UtcNow.AddDays(-30) } |
        Remove-Item -Force -ErrorAction SilentlyContinue
} catch { $null = $_ }

# The VHDX backup folder lives in the same directory but was never cleaned:
# the rotation above only matches fix_*.log, so up to 3 GB of stale
# sessiondata.vhdx accumulated here indefinitely. A backup older than 7 days
# has outlived any repair it could still be used to undo.
try {
    $staleBackupDir = Join-Path $LogDir "vhdx-backup"
    if (Test-Path $staleBackupDir) {
        Get-ChildItem $staleBackupDir -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "*.vhdx*" -and
                           $_.LastWriteTimeUtc -lt [datetime]::UtcNow.AddDays(-7) } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }
} catch { $null = $_ }
$script:LogLines = New-Object System.Collections.ArrayList

function Log {
    param([string]$Message, [string]$Colour = "White", [switch]$Indent)
    $pfx = ""
    if ($Indent) { $pfx = "      " }
    $ts = "[{0:HH:mm:ss}]" -f (Get-Date)
    $null = $script:LogLines.Add("$ts $pfx$Message")
    Write-Host "$pfx$Message" -ForegroundColor $Colour
}

function Save-Log {
    # WriteAllText rather than Out-File -Encoding utf8, which emits a UTF-8 BOM
    # on Windows PowerShell 5.1 and leaves a stray EF BB BF at the head of every
    # log file that gets pasted into a bug report.
    try {
        [System.IO.File]::WriteAllText($LogFile, (($script:LogLines) -join "`r`n"), `
            (New-Object System.Text.UTF8Encoding($false)))
    }
    catch { Write-Host "  [!] Could not write log file" -ForegroundColor DarkGray }
}

function Wait-ForAnyKey {
    <#
    .SYNOPSIS
        Waits for a keypress, and never throws when there is no console to read.
    .DESCRIPTION
        There were three divergent, unguarded copies of this. The one at the end
        of the script sat outside the try/finally, so on a host with no
        key-reading console an unhandled error was the script's final act.
    #>
    param([string]$Message = "  Press any key to close...")
    try { Write-Host $Message -ForegroundColor DarkGray } catch { $null = $_ }
    try {
        if ($Host.UI -and $Host.UI.RawUI) {
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            return
        }
    } catch { $null = $_ }
    try { [void][System.Console]::ReadKey($true) } catch { $null = $_ }
}

# Transcript backup (v4.8.0) -- catches output even if Save-Log fails
$script:TranscriptFile = Join-Path $LogDir "fix_$($script:SessionTimestamp)_transcript.log"
try { Start-Transcript -Path $script:TranscriptFile -Append -ErrorAction SilentlyContinue } catch { $null = $_ }

# -- Win32: bring window to foreground and flash taskbar icon --------
Add-Type -ErrorAction SilentlyContinue -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public struct FLASHWINFO {
    public uint cbSize;
    public IntPtr hwnd;
    public uint dwFlags;
    public uint uCount;
    public uint dwTimeout;
}

public static class Win32Window {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern bool FlashWindowEx(ref FLASHWINFO pwfi);
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();

    public static void BringToFront() {
        IntPtr h = GetConsoleWindow();
        if (h == IntPtr.Zero) return;
        ShowWindow(h, 9);           // SW_RESTORE
        SetForegroundWindow(h);
    }

    public static void Flash() {
        IntPtr h = GetConsoleWindow();
        if (h == IntPtr.Zero) return;
        FLASHWINFO fi = new FLASHWINFO();
        fi.cbSize  = (uint)Marshal.SizeOf(fi);
        fi.hwnd    = h;
        fi.dwFlags = 0x0003 | 0x000C;  // FLASHW_ALL | FLASHW_TIMERNOFG
        fi.uCount  = 0;                // flash until focused
        fi.dwTimeout = 0;
        FlashWindowEx(ref fi);
    }

    public static void StopFlash() {
        IntPtr h = GetConsoleWindow();
        if (h == IntPtr.Zero) return;
        FLASHWINFO fi = new FLASHWINFO();
        fi.cbSize  = (uint)Marshal.SizeOf(fi);
        fi.hwnd    = h;
        fi.dwFlags = 0;  // FLASHW_STOP
        fi.uCount  = 0;
        fi.dwTimeout = 0;
        FlashWindowEx(ref fi);
    }
}
'@

# -- Stop the desktop app (shared function) ---------------------------
function Stop-ClaudeDesktopApp {
    <#
    .SYNOPSIS
        Stops the Claude Desktop app's processes and no other process named
        claude, says how many others it left alone, and returns how many it
        stopped.
    .DESCRIPTION
        "claude" is not a unique process name: the Claude Code CLI runs under
        it too, for the app's project threads and for terminal sessions.
        Get-ClaudeProcessInventory tells the app apart. The two escalations
        after a relaunch kill through here; by then the relaunched app may
        have started new CLI processes, and the log says how many were spared.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param()

    $inventory = @(Get-ClaudeProcessInventory -Environment $script:ClaudeEnv)
    $appProcs  = @($inventory | Where-Object { $_.IsDesktop } | ForEach-Object { $_.Process })
    $spared    = @($inventory | Where-Object { -not $_.IsDesktop }).Count
    $stopped   = 0
    if ($appProcs.Count -gt 0 -and $PSCmdlet.ShouldProcess("$($appProcs.Count) Claude process(es)", "Stop")) {
        $appProcs | Stop-Process -Force -ErrorAction SilentlyContinue
        $stopped = $appProcs.Count
    }
    if ($spared -gt 0) {
        Log "Left $spared other process(es) named claude alone (Claude Code CLI, or not identifiable as the app)" `
            -Colour DarkGray -Indent
    }
    return $stopped
}

# -- Find Claude.exe (shared function) -------------------------------
function Find-ClaudeExe {
    <#
    .SYNOPSIS
        Resolves the Claude Desktop executable.
    .DESCRIPTION
        Discovery already did the hard part at startup, preferring the package,
        which is the only source that reliably identifies the desktop app.

        What used to live here was a six-tier search that could not reach a
        WindowsApps install at all: no search path covered WindowsApps, the
        registry branch missed the app\ subdirectory, and the brute-force scan
        could not read the TrustedInstaller ACL. It finished with a multi-second
        recursive scan of LocalAppData and Program Files whose result the direct
        launch path then explicitly rejected for being under WindowsApps.

        What remains are the branches that still earn their place: a cached
        path, the App Paths registry key, and a Start Menu shortcut, for
        traditional installs that discovery did not see.
    #>
    # 0. Whatever discovery resolved.
    if ($script:ClaudeEnv.ClaudeExe -and (Test-Path $script:ClaudeEnv.ClaudeExe)) {
        Log "Using discovered path ($($script:ClaudeEnv.ClaudeExeSource)): $($script:ClaudeEnv.ClaudeExe)" `
            -Colour DarkGray -Indent
        return $script:ClaudeEnv.ClaudeExe
    }
    if ($script:CapturedClaudeExe -and (Test-Path $script:CapturedClaudeExe)) {
        Log "Using path captured from running process: $($script:CapturedClaudeExe)" -Colour DarkGray -Indent
        return $script:CapturedClaudeExe
    }

    # 1. Cached path from a previous run. The CLI filter matters here too: a
    #    cache written before 6.2.0 may hold a CLI path, because the old
    #    Step 1 capture took the first process named claude and only knew the
    #    claude-code folder, not the terminal install in .local\bin.
    if (Test-Path $ExePathCache) {
        try {
            $cached = (Get-Content $ExePathCache -Raw -ErrorAction Stop).Trim()
            if ($cached -and -not (Test-ClaudeCodeCliPath $cached) -and (Test-Path $cached)) {
                Log "Found (cached): $cached" -Colour DarkGray -Indent
                return $cached
            }
            Log "Cached path invalid, searching..." -Colour DarkGray -Indent
            Remove-Item $ExePathCache -Force -ErrorAction SilentlyContinue
        } catch { $null = $_ }
    }

    # 2. App Paths registry, for a traditional install discovery did not see.
    try {
        $appPathsKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\claude.exe"
        if (Test-Path $appPathsKey) {
            $props   = Get-ItemProperty $appPathsKey -ErrorAction SilentlyContinue
            $appPath = $null
            if ($props) {
                if ($props.PSObject.Properties['(default)']) { $appPath = $props.'(default)' }
                if (-not $appPath -and $props.PSObject.Properties['Path']) {
                    $appPath = Join-Path $props.Path "Claude.exe"
                }
            }
            if ($appPath -and -not (Test-ClaudeCodeCliPath $appPath) -and (Test-Path $appPath)) {
                Log "Found (App Paths): $appPath" -Colour DarkGray -Indent
                return $appPath
            }
        }
    } catch { $null = $_ }

    # 3. Start Menu shortcut.
    #
    # The COM object and the CreateShortcut call are inside the try now. A
    # dangling .lnk made CreateShortcut throw, and with nothing catching it the
    # exception unwound all the way to the outer handler and abandoned the run.
    foreach ($mp in @($env:APPDATA, $env:ProgramData)) {
        if (-not $mp) { continue }
        try {
            $menuRoot = Join-Path $mp "Microsoft\Windows\Start Menu"
            if (-not (Test-Path $menuRoot)) { continue }
            $lnk = Get-ChildItem $menuRoot -Recurse -Filter "Claude*.lnk" -ErrorAction SilentlyContinue |
                   Select-Object -First 1
            if (-not $lnk) { continue }
            $shell  = New-Object -ComObject WScript.Shell
            $target = $shell.CreateShortcut($lnk.FullName).TargetPath
            if ($target -and -not (Test-ClaudeCodeCliPath $target) -and (Test-Path $target)) {
                Log "Found (Start Menu shortcut): $target" -Colour DarkGray -Indent
                return $target
            }
        } catch { $null = $_ }
    }

    # The where.exe last resort has been removed. It indexed $whereResult[0],
    # and on a single-line result that is a [string] whose element 0 is the
    # character "C", so the Test-Path guard was testing "C" rather than a path.

    return $null
}

# -- Service prerequisite check ---------------------------------------
function Test-CoworkServicePrereq {
    <#
    .SYNOPSIS
        Works out whether CoworkVMService is able to start at all.
    .DESCRIPTION
        "sc start" answers a blocked service with FAILED 87 (the parameter is
        incorrect) or FAILED 1058 (service disabled), and neither says why.
        Without this check the script retries, escalates to a cache purge, and
        reports failure against a service that was never going to start.

        Checks follow Anthropic's own remediation steps for Cowork on Windows:
        https://support.claude.com/en/articles/12622703-deploy-claude-desktop-for-windows

        Returns a hashtable: Found, StartType, Blocked, Reason, Fixes, Repaired
    #>
    $r = @{
        Found     = $false
        StartType = "Unknown"
        Blocked   = $false
        Reason    = $null
        Fixes     = @()
        Repaired  = $false
    }

    # 1. Virtual Machine Platform.
    #
    # Two lessons from the same user report, on the same day, in the same order
    # they were learned.
    #
    # First the wording. It used to say "Cowork runs in a Hyper-V VM and cannot
    # start without it", so the reporter ran Get-WindowsOptionalFeature with
    # -FeatureName *Hyper-V*, saw all seven Hyper-V features Enabled, and said
    # the script was wrong. That filter cannot match VirtualMachinePlatform,
    # because the name contains no "Hyper-V". Anything user-facing here has to
    # say that outright, which the text below now does.
    #
    # Then the check itself, which mattered far more.
    #
    # What Cowork actually needs is a working Host Compute Service stack. Virtual
    # Machine Platform is one way to get one. Full Hyper-V is another, and on a
    # machine that has it, VirtualMachinePlatform can read Disabled while
    # vmcompute and hns are present and HCS works perfectly.
    #
    # That is not hypothetical. The 2026-07-29 reporter's machine reads Disabled
    # here and runs Cowork: their cowork-service.log shows vmcompute.dll and
    # computecore.dll loading, HCN enumerating two networks, and a compute
    # system in State Running an hour before they ran this script.
    #
    # So check the thing that matters first. If the HCS services are there, this
    # feature flag is not the story and must not be told as one.
    $hcsPresent = $true
    foreach ($n in @('vmcompute', 'hns')) {
        if (-not (Get-Service -Name $n -ErrorAction SilentlyContinue)) { $hcsPresent = $false }
    }

    if ($script:IsAdmin -and -not $hcsPresent) {
        try {
            $vmp = Get-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -ErrorAction Stop
            if ($vmp -and $vmp.State -ne "Enabled") {
                $r.Blocked = $true
                $r.Reason  = "The Virtual Machine Platform feature is $($vmp.State), and the Host Compute Service stack is not registered either."
                $r.Fixes  += "This is NOT the same as Hyper-V. Enabling Hyper-V does not enable it, and"
                $r.Fixes  += "a *Hyper-V* search will not list it, because the name has no 'Hyper-V' in it."
                $r.Fixes  += "Check it with: Get-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform"
                $r.Fixes  += "Enable it with: Enable-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -All"
                $r.Fixes  += "Then RESTART the machine (use Restart, not shut down; Fast Startup can leave the services uninitialised)"
                return $r
            }
        } catch { $null = $_ }
    }

    # 2. Host Compute Service stack. Anthropic reports this as
    #    "Missing HCS services: HNS, vmcompute, vfpext".
    $missingHcs = @()
    foreach ($n in @('vmcompute', 'hns')) {
        if (-not (Get-Service -Name $n -ErrorAction SilentlyContinue)) { $missingHcs += $n }
    }
    if ($missingHcs.Count -gt 0) {
        $r.Blocked = $true
        $r.Reason  = "The Host Compute Service stack is not registered (missing: $($missingHcs -join ', ')), so the Cowork VM has nothing to run on."
        $r.Fixes  += "Enable-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -All"
        $r.Fixes  += "Then RESTART the machine"
        return $r
    }

    # 3. Hypervisor must be set to launch at boot. VMware and VirtualBox
    #    installs are the usual reason this gets switched off.
    if ($script:IsAdmin) {
        try {
            $bcd = & bcdedit /enum "{current}" 2>$null
            # Select-Object -First 1 matters: bcdedit can emit more than one
            # matching line, and .Trim() on the resulting array throws under
            # StrictMode. The empty catch around this then reported a genuinely
            # disabled hypervisor as fine.
            $hvLine = $bcd | Where-Object { $_ -match 'hypervisorlaunchtype' } |
                      Select-Object -First 1
            if ($hvLine -and $hvLine -notmatch '(?i)auto') {
                $r.Blocked = $true
                $r.Reason  = "The Windows hypervisor is not set to launch at boot ($($hvLine.Trim()))."
                $r.Fixes  += "bcdedit /set hypervisorlaunchtype auto"
                $r.Fixes  += "Then RESTART the machine"
                return $r
            }
        } catch { $null = $_ }
    }

    # 4. The service itself.
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $svc) {
        $r.Blocked = $true
        $r.Reason  = "CoworkVMService is not registered. A per-user MSIX install can finish without registering the Cowork virtualisation service, which leaves Claude installed but Cowork unable to start."
        $r.Fixes  += "Reinstall machine-wide: Add-AppxProvisionedPackage -Online -PackagePath 'Claude.msix' -SkipLicense -Regions 'all'"
        return $r
    }
    $r.Found = $true

    try {
        $r.StartType = (Get-CimInstance Win32_Service -Filter "Name='$ServiceName'" -ErrorAction Stop).StartMode
    } catch {
        try { $r.StartType = $svc.StartType.ToString() } catch { $null = $_ }
    }

    # 5. Start type Disabled is what answers "sc start" with 1058.
    if ($r.StartType -eq "Disabled") {
        if ($script:IsAdmin) {
            try {
                Set-Service -Name $ServiceName -StartupType Automatic -ErrorAction Stop
                $r.StartType = "Auto"
                $r.Repaired  = $true
            } catch {
                $r.Blocked = $true
                $r.Reason  = "CoworkVMService start type is Disabled and could not be re-enabled: $($_.Exception.Message)"
                $r.Fixes  += "Set-Service -Name CoworkVMService -StartupType Automatic"
                return $r
            }
        } else {
            $r.Blocked = $true
            $r.Reason  = "CoworkVMService start type is Disabled."
            $r.Fixes  += "Rerun this script as Administrator, or run: Set-Service -Name CoworkVMService -StartupType Automatic"
            return $r
        }
    }

    # 6. Duplicate package-family registrations produce error 87 on start.
    try {
        $pkgs = @(Get-AppxPackage -Name "*Claude*" -ErrorAction SilentlyContinue |
                  Where-Object { $_.Name -like "Claude*" })
        if ($pkgs.Count -gt 1) {
            $r.Blocked = $true
            $r.Reason  = "$($pkgs.Count) Claude packages are registered ($(($pkgs | ForEach-Object { $_.PackageFullName }) -join '; ')). Duplicate entries in the Claude package family make the service start fail with error 87."
            $r.Fixes  += "Remove the stale package with Remove-AppxPackage, keeping only the current version"
            return $r
        }
    } catch { $null = $_ }

    return $r
}

# -- Service restart function ----------------------------------------
function Restart-CoworkService {
    [CmdletBinding(SupportsShouldProcess)]
    param()

    # Under -WhatIf this reports success rather than failure, deliberately.
    # Returning $false would make Smart mode escalate to a Deep purge that it
    # also would not perform, and then report a cascade of work that never
    # happened. Reporting success keeps a dry run quiet, which is the point.
    if (-not $PSCmdlet.ShouldProcess($ServiceName, "Restart")) { return $true }

    # The prereq check runs for its ONE side effect, re-enabling a start type of
    # Disabled. Its findings are advisory and are not consulted until a start
    # has actually been attempted and failed.
    #
    # It used to hold a veto: Blocked meant return $false without ever calling
    # Start. A user transcript on 2026-07-29 showed what that costs. Their
    # machine reported Virtual Machine Platform as Disabled, so the check
    # blocked, and the script printed "Service failed to start" for a start it
    # never tried, then repeated it every five seconds.
    #
    # Their own service log settles it. Between 11:12:16, when this script
    # stopped the service, and 11:29:18, when Claude Desktop started it again,
    # cowork-service.log contains nothing at all. Not one failed start. And the
    # 11:29 start took 293ms and worked: vmcompute.dll and computecore.dll
    # loaded, HCN enumerated two networks, the named pipe came up. The VM had
    # been Running on that machine at 10:02 the same morning.
    #
    # A check that can be wrong must never be able to prevent the attempt that
    # would prove it wrong. Attempting a start on a machine that genuinely
    # cannot run one costs a few seconds and an error message. Refusing to
    # attempt one on a machine that can costs the user their working install.
    $script:ServicePrereq = Test-CoworkServicePrereq
    if ($script:ServicePrereq.Repaired) {
        Log "Service start type was Disabled, re-enabled (Automatic)" -Colour Green -Indent
    }
    if ($script:ServicePrereq.Blocked) {
        Log "Prerequisite check flagged: $($script:ServicePrereq.Reason)" -Colour DarkYellow -Indent
        Log "Trying to start anyway; the service is the authority on this, not the check." -Colour DarkGray -Indent
    }

    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $svc) {
        Log "[!] Service not found" -Colour Red -Indent
        return $false
    }

    # Stop if running.
    #
    # The wait loop used to exit on timeout without recording whether the
    # service had actually stopped. Execution then fell through to
    # Start-Service, which no-ops against a service that is still running, and
    # the poll at the end saw Running and returned $true. A hung service was
    # therefore reported as a successful restart. Step 2 had always handled this
    # correctly; the two implementations had drifted apart.
    if ($svc.Status -eq "Running") {
        $stopped = $false
        # See $script:ServiceStoppedByRun.
        $script:ServiceStoppedByRun = $true
        if ($script:IsAdmin) {
            try {
                $svc.Stop()
                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                while ($sw.Elapsed.TotalSeconds -lt $ServiceTimeout) {
                    Start-Sleep -Seconds 2
                    $curSvc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
                    if (-not $curSvc -or $curSvc.Status -eq "Stopped") { $stopped = $true; break }
                }
                $sw.Stop()
            } catch { $null = $_ }

            if (-not $stopped) {
                Log "Service did not stop within ${ServiceTimeout}s; force-killing $ServiceExe" -Colour DarkYellow -Indent
                try {
                    Stop-Process -Name $ServiceExe -Force -ErrorAction Stop
                    $stopped = $true
                } catch {
                    try {
                        Start-Process "taskkill" -ArgumentList "/F /IM $ServiceExe.exe" -NoNewWindow -Wait
                        Start-Sleep -Seconds 1
                        if (-not (Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue)) {
                            $stopped = $true
                        }
                    } catch { $null = $_ }
                }
            }

            if (-not $stopped) {
                Log "[!] Could not stop $ServiceName. Starting it now would report a false success." -Colour Red -Indent
                return $false
            }
        } else {
            # Without admin the service was never ours to control. Say so and
            # let the poll below report whatever state it is actually in.
            try { Stop-Process -Name $ServiceExe -Force -ErrorAction Stop }
            catch { Log "[i] Cannot stop the service without admin" -Colour DarkGray -Indent }
        }
        Start-Sleep -Seconds 1
    }

    # Start
    $startAttempted = $false
    $startError     = $null
    if ($script:IsAdmin) {
        $startAttempted = $true
        try {
            Start-Service -Name $ServiceName -ErrorAction Stop
        } catch {
            $startError = $_.Exception.Message
            try {
                Start-Process "sc.exe" -ArgumentList "start $ServiceName" -NoNewWindow -Wait
            } catch { $null = $_ }
        }
    }
    # Non-admin: we cannot start the service directly, but Claude will start it
    # when it launches. Poll to see if it comes up on its own.

    # Poll until Running
    $elapsed = 0
    while ($elapsed -lt $StartPollMax) {
        Start-Sleep -Seconds 2
        $elapsed += 2
        $svcNow = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($svcNow -and $svcNow.Status -eq "Running") {
            # It started. Whatever the prereq check thought, it was wrong or it
            # was irrelevant, so drop the finding rather than leaving a red
            # herring in $script:ServicePrereq for the summary to print.
            if ($script:ServicePrereq -and $script:ServicePrereq.Blocked) {
                Log "Started despite the prerequisite warning; the warning was wrong." -Colour Green -Indent
                $script:ServicePrereq.Blocked = $false
            }
            return $true
        }
    }

    # It did not come up. NOW the prereq findings are worth reading out, because
    # there is a real failure for them to explain. This ordering is the whole
    # point: diagnose a failure, never predict one.
    if ($startAttempted) {
        if ($startError) {
            Log "[!] Start failed: $startError" -Colour Red -Indent
        } else {
            Log "[!] The service did not reach Running within ${StartPollMax}s" -Colour Red -Indent
        }
        if ($script:ServicePrereq -and $script:ServicePrereq.Blocked) {
            Log "Most likely cause: $($script:ServicePrereq.Reason)" -Colour Yellow -Indent
            foreach ($f in $script:ServicePrereq.Fixes) {
                Log "  -> $f" -Colour Yellow -Indent
            }
        }
    } else {
        Log "[i] Not admin, so the service was never asked to start" -Colour DarkGray -Indent
    }
    return $false
}

function Stop-CoworkServiceWithTimeout {
    <#
    .SYNOPSIS
        Stops the service with a hard timeout and confirms that it stopped.
    .DESCRIPTION
        The job reports the status it observed AFTER its own stop attempt, and
        that reported status is what decides the outcome.

        Four near-identical copies of this logic treated a completed Wait-Job as
        proof of a stop. It is not: Stop-Service inside the job runs with
        -ErrorAction SilentlyContinue, so the job completes just as cheerfully
        when the stop failed as when it worked. Returns $true only when the
        service is genuinely Stopped or gone.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param([int]$TimeoutSeconds = 30)

    # As with Restart-CoworkService: a dry run reports success so the caller
    # does not go on to describe a cascade of recovery work it also would not
    # have performed.
    if (-not $PSCmdlet.ShouldProcess($ServiceName, "Stop")) { return $true }
    # See $script:ServiceStoppedByRun.
    $script:ServiceStoppedByRun = $true

    $stopped = $false
    $job     = $null
    try {
        $job = Start-Job -ScriptBlock {
            param($svc)
            Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
            $s = Get-Service -Name $svc -ErrorAction SilentlyContinue
            if (-not $s) { return "Absent" }
            return "$($s.Status)"
        } -ArgumentList $ServiceName

        if (Wait-Job $job -Timeout $TimeoutSeconds) {
            $observed = @(Receive-Job $job)
            if ($observed.Count -gt 0) {
                $last = "$($observed[-1])"
                if ($last -eq "Absent" -or $last -eq "Stopped") { $stopped = $true }
            }
        } else {
            Stop-Job $job -ErrorAction SilentlyContinue
        }
    } catch {
        # Job plumbing failure is non-fatal; the force-kill below is the answer.
        $null = $_
    } finally {
        if ($job) { Remove-Job $job -Force -ErrorAction SilentlyContinue }
    }

    if (-not $stopped) {
        Log "Service did not stop within ${TimeoutSeconds}s; force-killing $ServiceExe" -Colour DarkYellow -Indent
        try { Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue | Stop-Process -Force } catch { $null = $_ }
        Start-Sleep -Seconds 1
        $now = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if (-not $now -or $now.Status -eq "Stopped") { $stopped = $true }
        if (-not $stopped) {
            Log "[!] $ServiceName is still running after a force-kill attempt" -Colour Red -Indent
        }
    }
    return $stopped
}

function Start-ClaudeDesktop {
    <#
    .SYNOPSIS
        Relaunches Claude Desktop. Returns $true only if something started.
    .DESCRIPTION
        There were two copies of this. The deep-escalation one tried the
        scheduled task and nothing else, so when the task was absent it went on
        to wait 120 seconds for a Claude it had never launched.

        Order matters for an MSIX install: launching the exe directly produces
        a loose instance with a duplicate taskbar icon, so the AUMID comes
        first and the direct path is the fallback.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param()

    if (-not $PSCmdlet.ShouldProcess("Claude Desktop", "Launch")) { return $true }

    try {
        $task = Get-ScheduledTask -TaskName "LaunchClaudeAdmin" -TaskPath "\Claude\" `
                -ErrorAction SilentlyContinue
        if ($task) {
            Start-ScheduledTask -TaskName "LaunchClaudeAdmin" -TaskPath "\Claude\" -ErrorAction Stop
            Log "Relaunched Claude via scheduled task" -Colour Green -Indent
            return $true
        }
    } catch {
        Log "Scheduled-task relaunch failed: $($_.Exception.Message)" -Colour DarkYellow -Indent
    }

    if ($script:ClaudeEnv.IsMsix -and $script:ClaudeEnv.Aumid) {
        try {
            Start-Process $script:ClaudeEnv.Aumid -ErrorAction Stop
            Log "Relaunched Claude (MSIX)" -Colour Green -Indent
            return $true
        } catch {
            Log "MSIX relaunch failed: $($_.Exception.Message)" -Colour DarkYellow -Indent
        }
    }

    $exe = Find-ClaudeExe
    if ($exe) {
        try {
            Start-Process $exe -ErrorAction Stop
            Log "Relaunched Claude directly" -Colour Green -Indent
            return $true
        } catch {
            Log "Direct relaunch failed: $($_.Exception.Message)" -Colour DarkYellow -Indent
        }
    }

    Log "[!] Could not relaunch Claude by any method" -Colour Red -Indent
    return $false
}

# -- HCS error detection -----------------------------------------------
function Get-FixWinEvent {
    <#
    .SYNOPSIS
        Get-WinEvent -FilterHashtable with -ErrorAction Stop, where nothing to
        read is no error: returns the events, or nothing when none matches,
        the log does not exist or its provider is not installed. Any other
        error is thrown, for the caller to name.
    .DESCRIPTION
        Get-WinEvent says there is nothing to read with an error whose id
        starts NoMatchingEventsFound, NoMatchingLogsFound or
        NoMatchingProvidersFound (measured on Windows PowerShell 5.1), the
        usual answer on a machine without Hyper-V. Until 6.3.0 the reads used
        SilentlyContinue, which hid every other error with them, so a log
        that could not be read looked like one with nothing in it. The same
        as Get-MatchingWinEvent in Watch-ClaudeHealth.ps1.

        Each StartTime given here is the UTC time as round-trip text, ending
        in Z: given a [datetime], Get-WinEvent reads it as local time whatever
        its Kind, and a local time is wrong in the hour the clocks go back,
        which happens twice. Measured on Windows PowerShell 5.1 on this
        machine's log, in summer time and across that hour: the text was
        exact in both. Until 6.3.0 it was the local time.
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

function Test-RecentHcsErrors {
    <#
    .SYNOPSIS
        Checks for recent HCS errors. Now checks Information-level events too (v4.8.0).
        Returns: $null (clean), "shutdown_stale" (0xC037010D -- property query bug),
        "construct_failure" (0x800707DE), "guest_connect_failure" (isGuestConnected timeout),
        or "hcs_error" (other HCS issues). A log that cannot be read, an
        Event Log, a Claude log or the Claude log folder, is said, naming the
        error, and passed over: until 6.3.0 it was not said, and a log that
        could not be read looked like one with no HCS error in it.
    #>
    # Check 1: HCS Compute event log -- check ALL levels including Information
    try {
        # Check for 0xC037010D (shutdown failures) -- these are Information-level
        $hcsInfoFilter = @{
            LogName   = "Microsoft-Windows-Hyper-V-Compute-Operational"
            # The UTC time as round-trip text: see Get-FixWinEvent.
            StartTime = [datetime]::UtcNow.AddMinutes(-5).ToString('o')
        }
        $hcsInfoEvents = @(Get-FixWinEvent -Filter $hcsInfoFilter -MaxEvents 50)

        $shutdownFailures = 0
        $constructFailures = 0
        foreach ($evt in $hcsInfoEvents) {
            $msg = $evt.Message
            if ($msg -match "0xC037010D") { $shutdownFailures++ }
            if ($msg -match "0x800707DE") { $constructFailures++ }
        }

        # Construct failure (0x800707DE) is the most actionable
        if ($constructFailures -gt 0) {
            if ($shutdownFailures -gt 0) {
                Log "  HCS: $constructFailures construct failures + $shutdownFailures shutdown failures in 5 min" -Colour DarkYellow -Indent
            }
            return "construct_failure"
        }

        # Shutdown failures from Claude Desktop property query bug (literal '$')
        # These happen on EVERY VM shutdown -- only flag if excessive spike
        # 0xC037010D occurs on every normal VM shutdown (property query bug).
        # Only flag as stale if excessive spike (>15 in 5 min) indicating
        # actual HCS corruption rather than normal shutdown events.
        if ($shutdownFailures -gt 15) {
            return "shutdown_stale"
        }

        # Also check Critical/Error level (original check)
        $hcsCritFilter = @{
            LogName   = "Microsoft-Windows-Hyper-V-Compute-Admin"
            Level     = @(1, 2)
            # The UTC time as round-trip text: see Get-FixWinEvent.
            StartTime = [datetime]::UtcNow.AddMinutes(-5).ToString('o')
        }
        $hcsCritEvents = @(Get-FixWinEvent -Filter $hcsCritFilter -MaxEvents 10)
        if ($hcsCritEvents) {
            $hasRealError = $false
            foreach ($evt in $hcsCritEvents) {
                $xml = $evt.ToXml()
                # Skip 0xC037010D in Admin log -- already handled by threshold
                # check in Operational log above. These are normal shutdown events.
                if ($xml -match "0xC037010D" -or $xml -match "Invalid JSON document") {
                    continue
                }
                $hasRealError = $true
            }
            if ($hasRealError) { return "hcs_error" }
        }
    } catch {
        Log "  HCS: the Hyper-V Compute event logs could not be read ($($_.Exception.Message)), so recent HCS errors are not known" -Colour DarkYellow -Indent
    }

    # Check 2: Claude log files (keep existing logic)
    $hcsPatterns = @("HCS operation failed", "failed to create compute system",
                     "HcsWaitForOperationResult", "0x800707DE")
    # %ProgramData%\Claude does not exist on any current build, so the only
    # live log directory is the one discovery resolved: the freshest logs
    # folder, %LOCALAPPDATA%\Claude\logs on app 6.18355.
    $claudeLogDirs = @()
    if ($script:ClaudeEnv.LogDir) { $claudeLogDirs += $script:ClaudeEnv.LogDir }
    $recentLogs = @()
    foreach ($dir in $claudeLogDirs) {
        if (Test-Path -LiteralPath $dir) {
            try {
                $recentLogs += @(Get-ChildItem -LiteralPath $dir -Filter "*.log" -ErrorAction Stop |
                    Where-Object { ([datetime]::UtcNow - $_.LastWriteTimeUtc).TotalMinutes -lt 5 })
            } catch {
                Log "  HCS: the Claude log folder $dir could not be listed ($($_.Exception.Message)), so HCS errors in its logs are not known" -Colour DarkYellow -Indent
            }
        }
    }
    foreach ($logFile in $recentLogs) {
        try {
            $content = Get-Content -LiteralPath $logFile.FullName -Tail 50 -ErrorAction Stop
        } catch {
            Log "  HCS: the Claude log $($logFile.Name) could not be read ($($_.Exception.Message)), so HCS errors in it are not known" -Colour DarkYellow -Indent
            continue
        }
        $text = $content -join "`n"
        foreach ($pattern in $hcsPatterns) {
            if ($text -match [regex]::Escape($pattern)) { return "hcs_error" }
        }
    }

    # Check 3: cowork-service.log for guest connection failures
    $guestState = Test-CoworkServiceLog -WindowSeconds 60 -Brief
    if ($guestState -eq "guest-timeout" -or $guestState -eq "guest-error") {
        return "guest_connect_failure"
    }

    return $null
}

# -- cowork-service.log guest connection detection -------------------------
function Test-CoworkServiceLog {
    <#
    .SYNOPSIS
        Reads cowork-service.log to determine guest connection state.
        Returns: $null (log not found), "no-polling", "guest-error",
        "guest-timeout", "guest-connected", or "guest-polling".
    #>
    param(
        [int]$WindowSeconds = 30,
        [switch]$Brief
    )

    # a) Build candidate paths
    # %ProgramData% is gone, so discovery's log directory is the only candidate.
    # Note that cowork-service.log itself has not been written since March on
    # current builds. When that is the case the window filter below finds no
    # recent lines and the state comes back "no-polling", which is accurate.
    $candidatePaths = @()
    if ($script:ClaudeEnv.LogDir) {
        $candidatePaths += (Join-Path $script:ClaudeEnv.LogDir "cowork-service.log")
    }

    # b) Find the first path that exists
    $svcLogPath = $null
    foreach ($p in $candidatePaths) {
        if (Test-Path $p) { $svcLogPath = $p; break }
    }
    if (-not $svcLogPath) { return $null }

    # c) Read the last 200 lines (handle file locks gracefully)
    $lines = $null
    try {
        $lines = Get-Content $svcLogPath -Tail 200 -ErrorAction Stop
    } catch {
        # Tail failed (locked file) -- try stream-based read
        try {
            $stream = New-Object System.IO.FileStream(
                $svcLogPath, [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            $reader = New-Object System.IO.StreamReader($stream)
            $allText = $reader.ReadToEnd()
            $reader.Close()
            $stream.Close()
            $allLines = $allText -split "`r?`n"
            $lines = if ($allLines.Count -gt 200) { $allLines[-200..-1] } else { $allLines }
        } catch {
            return $null
        }
    }
    if (-not $lines -or $lines.Count -eq 0) { return $null }

    # d) Filter to lines within the last $WindowSeconds
    #
    # Local time, on purpose: the stamps below are parsed from log text the
    # service writes in local time with no offset, so local now is the only
    # time they can be compared with. For the hour around a change of the
    # clock a line can read up to an hour newer or older than it is.
    $now = Get-Date
    $recentLines = @()
    foreach ($line in $lines) {
        # Format C: "2026/03/23 01:46:54.946851" (ProgramData logs, v5.1.0)
        if ($line -match '^\s*(\d{4}/\d{2}/\d{2}\s+\d{2}:\d{2}:\d{2}\.\d+)') {
            try {
                $tsStr = $Matches[1].Trim()
                # Truncate to milliseconds if microseconds present
                if ($tsStr -match '^(.+\.\d{3})\d+$') { $tsStr = $Matches[1] }
                $tsStr = $tsStr -replace '/', '-'
                $ts = [datetime]::ParseExact($tsStr, "yyyy-MM-dd HH:mm:ss.fff",
                    [System.Globalization.CultureInfo]::InvariantCulture)
                if (($now - $ts).TotalSeconds -le $WindowSeconds) {
                    $recentLines += $line
                }
            } catch { $null = $_ }
        }
        # Format A: "yyyy-MM-dd HH:mm:ss.fff" (original format)
        elseif ($line -match '^\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\.\d{3})') {
            try {
                $ts = [datetime]::ParseExact($Matches[1].Trim(), "yyyy-MM-dd HH:mm:ss.fff",
                    [System.Globalization.CultureInfo]::InvariantCulture)
                if (($now - $ts).TotalSeconds -le $WindowSeconds) {
                    $recentLines += $line
                }
            } catch { $null = $_ }
        } elseif ($line -match '^\s*(\d{2}:\d{2}:\d{2}\.\d{3})') {
            try {
                $ts = [datetime]::ParseExact($Matches[1], "HH:mm:ss.fff",
                    [System.Globalization.CultureInfo]::InvariantCulture)
                # Assume today; if timestamp is in the future, assume yesterday
                $ts = $now.Date.Add($ts.TimeOfDay)
                if ($ts -gt $now) { $ts = $ts.AddDays(-1) }
                if (($now - $ts).TotalSeconds -le $WindowSeconds) {
                    $recentLines += $line
                }
            } catch { $null = $_ }
        }
    }

    # e) Count relevant patterns
    $guestConnectCalls = 0
    $guestConnectSuccess = 0
    $errors = 0

    for ($i = 0; $i -lt $recentLines.Count; $i++) {
        $l = $recentLines[$i]
        if ($l -match "method=isGuestConnected") {
            $guestConnectCalls++
            # Check next 2 lines for successful response
            for ($j = 1; $j -le 2 -and ($i + $j) -lt $recentLines.Count; $j++) {
                if ($recentLines[$i + $j] -match "Sent response|RPC to VM") {
                    $guestConnectSuccess++
                    break
                }
            }
        }
        if ($l -match "(?i)(error|timeout|failed|refused)") {
            $errors++
        }
    }

    # f) Determine state
    $state = $null
    if ($guestConnectCalls -eq 0) {
        $state = "no-polling"
    } elseif ($errors -gt 0) {
        $state = "guest-error"
    } elseif ($guestConnectCalls -gt 5 -and $guestConnectSuccess -eq 0) {
        $state = "guest-timeout"
    } elseif ($guestConnectSuccess -gt 0) {
        $state = "guest-connected"
    } else {
        $state = "guest-polling"
    }

    # g) Return
    if ($Brief) { return $state }
    Log "  cowork-svc    : $state (${guestConnectCalls} polls, ${guestConnectSuccess} connected, ${errors} errors in ${WindowSeconds}s)" -Colour DarkGray -Indent
    return $state
}

# Invoke-WithTimeout has been removed. Its only caller was the Get-VM
# heartbeat probe in Test-HyperVReady, which could never fire because the
# Cowork VM is an HCS compute system and does not appear in VMMS. With that
# probe gone the function had no callers, and its $Default parameter had never
# been supplied by anything.

function Wait-ForServiceProcessExit {
    <#
    .SYNOPSIS
        Waits for the service process to let go of the VHDX files.
    .DESCRIPTION
        Returns $true when the process is gone.

        Three identical copies of this loop reported the state observed BEFORE
        the final sleep, so a process that exited during that last second was
        still announced as holding the files locked. The status is re-read once
        the loop ends.
    #>
    param([int]$TimeoutSeconds = 6)

    $waited = 0
    while ($waited -lt $TimeoutSeconds) {
        if (-not (Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue)) { break }
        Start-Sleep -Seconds 1
        $waited++
    }

    $stillRunning = [bool](Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue)
    if ($waited -gt 0) {
        if ($stillRunning) {
            Log "Service process still running after ${waited}s; the VHDX files may be locked" -Colour DarkYellow -Indent
        } else {
            Log "Service process exited after ${waited}s" -Colour DarkGray -Indent
        }
    }
    return (-not $stillRunning)
}

function Invoke-HcsDiag {
    <#
    .SYNOPSIS
        Runs hcsdiag.exe with a timeout. Returns its output when it exited 0,
        '' when it exited 0 and printed nothing, and $null when it exited with
        any other code, timed out or is not installed.
    .DESCRIPTION
        Until 6.3.0 the exit code was thrown away. An error hcsdiag printed
        came back as if it were the list, so the check after a kill read the
        error, found no compute system in it and counted the kill as done.
        And "hcsdiag list" prints nothing at all when no compute system is
        left, which callers that tested the text for truth read as no answer.
        $null is now the only "did not answer". The exit code is read inside
        the job, where hcsdiag runs.

        A non-zero exit is logged once, and again only when what hcsdiag says
        changes or it has answered in between, as the monitor's copy does.
        Step 9's wait asks for the list on every pass, so a host where it
        keeps failing would otherwise get the same line every few seconds.

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
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        Log "hcsdiag timed out after ${TimeoutSeconds}s (args: $($Arguments -join ' '))" -Colour DarkYellow -Indent
        return $null
    }
    $result = @(Receive-Job $job -ErrorAction SilentlyContinue) |
        Where-Object { $null -ne $_ -and $_.PSObject.Properties['ExitCode'] } | Select-Object -Last 1
    Remove-Job $job -Force -ErrorAction SilentlyContinue
    if ($null -eq $result) { return $null }
    if ($result.ExitCode -ne 0) {
        $firstLine = @(([string]$result.Text) -split "`r?`n" | Where-Object { $_.Trim() }) | Select-Object -First 1
        $said = if ($firstLine) { ": $($firstLine.Trim())" } else { '' }
        # The verb only, so a kill of each GUID in turn is one note, not one each.
        $note = "hcsdiag $(@($Arguments)[0]) exited $($result.ExitCode)$said"
        if ($note -ne $script:LastHcsDiagNote) {
            Log $note -Colour DarkGray -Indent
            $script:LastHcsDiagNote = $note
        }
        return $null
    }
    $script:LastHcsDiagNote = $null
    return [string]$result.Text
}

function Get-HcsList {
    <#
    .SYNOPSIS
        "hcsdiag list" through Invoke-HcsDiag: the text, '' for an empty list,
        or $null when hcsdiag did not answer. Every Cowork GUID it shows is
        added to $script:CoworkGuidsSeen.
    .DESCRIPTION
        Every list this script reads goes through here, so step 5 can match a
        hung worker against any Cowork compute system seen during the run.
        Until 6.3.0 it matched against the one list the last cleanup had read,
        which step 5 cleared just before, and which each cleanup overwrote.
    #>
    $text = Invoke-HcsDiag -Arguments "list"
    if ($null -ne $text) {
        foreach ($g in @(Get-CoworkHcsGuids -ListOutput ([string]$text))) { $script:CoworkGuidsSeen[$g] = $true }
    }
    return $text
}

function Close-StaleHcsVms {
    <#
    .SYNOPSIS
        Finds and kills stale cowork-vm compute systems via hcsdiag.
        Returns the number of compute systems acted on.
    .DESCRIPTION
        hcsdiag exposes list, exec, console, read, write, share and kill.
        There is no "close" verb, so kill is the only valid action. Passing
        an unsupported verb makes hcsdiag print usage and exit without doing
        anything, which reads as success to the caller.

        The parse of "hcsdiag list" lives in Get-CoworkHcsGuids in the shared
        ClaudeEnv region, because Prevent and Watch need exactly the same thing
        and used to each carry their own subtly different copy.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [ValidateSet('kill')]
        [string]$Action = "kill"
    )
    $killed = 0
    # Records whether hcsdiag actually answered. A return of 0 means "nothing to
    # kill", but that is true both when hcsdiag reported an empty list and when
    # hcsdiag could not be reached at all, and callers were printing "none found
    # via hcsdiag" for both. Step 0 already told the truth here; step 5 did not.
    $script:LastHcsListOk = $false
    # And what it listed and whether kills went out, for the callers that say
    # what was found.
    $script:LastHcsListed    = 0
    $script:LastHcsKillsSent = 0
    try {
        # Get-HcsList also records what the list shows, for step 5's worker
        # check: a compute system leaves the list when it is killed.
        $hcsList = Get-HcsList
        # $null is "did not answer". An empty list is an answer, with nothing
        # in it to kill. Until 6.3.0 the test was "-not $hcsList", and '' is
        # false, so a host with no compute system left read as no answer.
        if ($null -eq $hcsList) { return 0 }
        $script:LastHcsListOk = $true
        # @(): with exactly one stale VM the parser returns a single string,
        # and .Count on a string throws under StrictMode in Windows PowerShell
        # 5.1. Until 6.3.0 the catch below swallowed that, so the one case
        # this function exists for returned 0, which callers print as "none
        # found".
        $entries = @(Get-CoworkHcsGuids -ListOutput ([string]$hcsList))
        $script:LastHcsListed = $entries.Count
        if ($entries.Count -eq 0) { return 0 }

        $sent = @()
        foreach ($guid in $entries) {
            try {
                if ($PSCmdlet.ShouldProcess($guid, "hcsdiag $Action")) {
                    # What the kill returns is not read: the list afterwards
                    # says whether it worked.
                    $null = Invoke-HcsDiag -Arguments $Action,$guid
                    $sent += $guid
                }
            } catch { $null = $_ }
        }
        $script:LastHcsKillsSent = $sent.Count
        if ($sent.Count -eq 0) { return 0 }

        # Counted from the list afterwards, not from the calls made. Until
        # 6.3.0 every kill sent was counted, and the callers printed "Killed N"
        # whether hcsdiag had removed anything or not.
        $left = Get-CoworkHcsSurvivors -Guids $sent -List { Get-HcsList }
        if ($null -eq $left) {
            Log "hcsdiag did not answer after the kill, so whether $($sent.Count) compute system(s) are gone is not known" -Colour DarkYellow -Indent
            return 0
        }
        foreach ($g in $left) {
            # Listed when last read, which is not proof the kill failed for
            # good: a compute system can take longer to leave the list than
            # the check waits. Until 6.3.0 this said the kill did not remove it.
            Log "[!] compute system $g was still listed after the kill" -Colour DarkYellow -Indent
        }
        $killed = $sent.Count - @($left).Count
    } catch { $null = $_ }
    return $killed
}

function Test-WorkerExited {
    <#
    .SYNOPSIS
        Whether a held worker process has exited: its handle says so, or no
        process with its id and start time is running any more.
    .DESCRIPTION
        HasExited answers through the handle the Process object holds. Where
        no handle could be held it opens the process by id, which after the
        worker exits can name a new process, so the start time is compared
        too. When the start time cannot be read the answer is "not exited",
        so a worker is never reported gone on a guess.
    #>
    param($Held, [int]$ProcessId, [datetime]$StartTime)
    try { if ($Held.HasExited) { return $true } } catch { $null = $_ }
    $now = $null
    try { $now = Get-Process -Id $ProcessId -ErrorAction Stop } catch { return $true }
    try { return ($now.StartTime -ne $StartTime) } catch { return $false }
}

function Stop-HungCoworkWorker {
    <#
    .SYNOPSIS
        Stops the worker process (vmwp.exe) of any Cowork compute system seen
        during this run that is still running. Returns how many it stopped
        (Stopped) and how many it found but could not confirm gone (Left).
    .DESCRIPTION
        Step 5's Method 2. Only the worker of a Cowork compute system: until
        6.3.0 this killed every vmwp.exe on the machine, so a Hyper-V VM, WSL2
        or Windows Sandbox running beside Claude went down with it. A worker
        is the Cowork VM's when the GUID its command line starts with is in
        $script:CoworkGuidsSeen, every Cowork GUID any hcsdiag list showed in
        this run. Until 6.3.0 it was matched against the one list the last
        cleanup read, which step 5 cleared first, so a compute system killed
        earlier in the run, whose worker hung on, was never matched.

        Each worker is held as a Process object before anything is sent to
        it, and only when its start time is the one the process query read,
        so a PID Windows has since given to another process is left alone.
        The object keeps a handle open where Windows allows one, and while it
        does the PID cannot be reused. After hcsdiag's kill it waits on that
        object, up to five seconds, for the worker to go. The force-kill goes
        through that object too, and the worker counts as stopped only once
        it has exited. Until 6.3.0 the
        fallback killed by bare PID two seconds later and counted the worker
        stopped as soon as Stop-Process returned, without looking.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param()

    $stopped = 0
    $left    = 0
    $vmwpProcs = @(Get-CimInstance Win32_Process -Filter "Name='vmwp.exe'" -ErrorAction SilentlyContinue)
    $coworkWorkers = @($vmwpProcs | Where-Object {
        $wg = Get-VmWorkerGuid -CommandLine ([string]$_.CommandLine)
        $wg -and $script:CoworkGuidsSeen.ContainsKey($wg)
    })
    if ($vmwpProcs.Count -eq 0) {
        Log "No VM worker processes running" -Colour DarkGray -Indent
    } elseif ($script:CoworkGuidsSeen.Count -eq 0) {
        Log "$($vmwpProcs.Count) VM worker process(es) running, none known to be the Cowork VM's, so none were touched" -Colour DarkGray -Indent
    } elseif ($coworkWorkers.Count -eq 0) {
        Log "No hung Cowork VM worker" -Colour DarkGray -Indent
    }
    foreach ($vmwp in $coworkWorkers) {
        $vmwpPid  = [int]$vmwp.ProcessId
        $vmwpGuid = Get-VmWorkerGuid -CommandLine ([string]$vmwp.CommandLine)
        if (-not $PSCmdlet.ShouldProcess("vmwp.exe PID $vmwpPid (GUID $vmwpGuid)", "Kill")) { continue }
        $held = $null
        try {
            try { $held = Get-Process -Id $vmwpPid -ErrorAction Stop } catch { $held = $null }
            if (-not $held) {
                Log "The Cowork VM's worker (vmwp.exe PID $vmwpPid) had exited before it was stopped" -Colour DarkGray -Indent
                continue
            }
            $heldStart = $null
            try { $heldStart = $held.StartTime } catch { $heldStart = $null }
            if ($null -eq $heldStart -or $null -eq $vmwp.CreationDate) {
                Log "vmwp.exe PID $vmwpPid could not be confirmed as the Cowork VM's worker, so it was left alone" -Colour DarkYellow -Indent
                $left++
                continue
            }
            if ([math]::Abs(($heldStart - [datetime]$vmwp.CreationDate).TotalSeconds) -ge 1) {
                Log "The Cowork VM's worker (PID $vmwpPid) had exited, and the PID now belongs to another process, which was left alone" -Colour DarkGray -Indent
                continue
            }
            # Hold a handle, so the PID cannot be reused while this runs.
            # Best effort: Test-WorkerExited also compares the start time.
            try { $null = $held.Handle } catch { $null = $_ }
            if ($script:IsAdmin) {
                # Cleaner than killing the worker. What hcsdiag returns is not
                # read: the worker gone is the proof.
                $null = Invoke-HcsDiag -Arguments "kill",$vmwpGuid
                # Wait on the held object for the worker to go, up to five
                # seconds, and only then fall back to the force-kill below.
                # Until 6.3.0 this slept a fixed two seconds, so a worker that
                # took longer to leave after a kill that worked was force-killed
                # anyway, with the VHDX risk that brings. Where the object has
                # no handle to wait on, the two seconds are kept.
                $waited = $false
                try { $null = $held.WaitForExit(5000); $waited = $true } catch { $null = $_ }
                if (-not $waited) { Start-Sleep -Seconds 2 }
                if (Test-WorkerExited -Held $held -ProcessId $vmwpPid -StartTime $heldStart) {
                    Log "Stopped the Cowork VM's worker (vmwp.exe PID $vmwpPid) through hcsdiag" -Colour Green -Indent
                    $stopped++
                    continue
                }
            }
            # Fallback: kill it through the object held above, never by a bare
            # PID, and give it five seconds to go.
            try { $held.Kill() } catch { $null = $_ }
            try { $null = $held.WaitForExit(5000) } catch { $null = $_ }
            if (Test-WorkerExited -Held $held -ProcessId $vmwpPid -StartTime $heldStart) {
                Log "WARNING: Force-killed vmwp.exe (PID $vmwpPid, GUID $vmwpGuid); VHDX corruption risk" -Colour DarkYellow -Indent
                $stopped++
            } else {
                Log "[!] vmwp.exe (PID $vmwpPid) is unkillable, host restart may be needed" -Colour Red -Indent
                $left++
            }
        } finally {
            if ($held) { try { $held.Dispose() } catch { $null = $_ } }
        }
    }
    return [pscustomobject]@{ Stopped = $stopped; Left = $left }
}

function Get-AppSessionBlocker {
    <#
    .SYNOPSIS
        Why the app's own Claude Code sessions hold an automated repair back,
        or $null when the app runs none.
    .DESCRIPTION
        A repair stops the app, and every Claude Code session the app started
        stops with it. Fails closed: when the sessions cannot be listed, that
        is the reason given. Get-AutomatedFixBlocker asks this first, and
        Assert-NoNewAppSession asks it again before each kill of the app.
    #>
    try {
        $appCli = @(Get-ClaudeAppCliProcess -Environment $script:ClaudeEnv)
        if ($appCli.Count -gt 0) {
            return "$($appCli.Count) Claude Code session(s) the app started are running (PID $(@($appCli | ForEach-Object { $_.Id }) -join ', '))"
        }
    } catch {
        return "the Claude Code sessions the app started could not be listed ($($_.Exception.Message))"
    }
    return $null
}

function Start-ServiceStoppedByRun {
    <#
    .SYNOPSIS
        Starts CoworkVMService again when this run stopped it and it is not
        running. Best effort: it logs what happened and does not throw.
    .DESCRIPTION
        Assert-NoNewAppSession calls this before a -Quiet run ends part-way,
        after Step 2 stopped the service and before Step 7 would start it
        again, or after a recovery in Step 9 stopped it and could not start
        it. Windows does not start a manually stopped service through its
        named-pipe trigger, so without this Cowork stayed down for as long as
        the Claude Code session the run stopped for went on, and the monitor
        waits a full cooldown before it tries again. Nothing is started when
        this run did not stop the service, or when it is running.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param()
    if (-not $script:ServiceStoppedByRun) { return }
    try {
        $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if (-not $svc -or $svc.Status -eq "Running") { return }
        if (-not $PSCmdlet.ShouldProcess($ServiceName, "Start again")) { return }
        Start-Service -Name $ServiceName -ErrorAction Stop
        Log "Started $ServiceName again, which this run had stopped, so Cowork is not left down" -Colour Green -Indent
    } catch {
        Log "[!] Could not start $ServiceName again ($($_.Exception.Message)); run Fix when Cowork is needed" -Colour DarkYellow -Indent
    }
}

function Assert-NoNewAppSession {
    <#
    .SYNOPSIS
        In a -Quiet run, ends the run when the app runs a Claude Code session,
        or when that cannot be ruled out, right before a step that stops the
        app. -ExitCode is the code it ends with: $script:BlockedExitCode
        before Step 1, when nothing that stops the app or the service has run
        yet, and $script:StoppedPartWayExitCode after it, when Steps 1 and 2
        have run. On that second path it first starts the service again when
        this run stopped it.
    .DESCRIPTION
        The -Quiet gate asks once, at the start. A Smart run takes minutes,
        and the app can start a project thread or a Code session in that
        time, which the next kill of the app would end. Until 6.3.0 the start
        was the only time it was asked. This asks again before Step 1's kill,
        before Step 4's kill of the app's leftover processes and before each
        escalation that stops the app, and fails closed the same way. Manual
        runs are not asked: the user wants the fix.

        exit here ends the script, not only this function, and the finally at
        the end of the script still releases the lock.
    #>
    param(
        [string]$Before,
        [Parameter(Mandatory)][int]$ExitCode
    )
    if (-not $Quiet) { return }
    $why = Get-AppSessionBlocker
    if ($why) {
        if ($ExitCode -eq $script:BlockedExitCode) {
            Log "BLOCKED: $why; $Before would stop the app, so the automated fix ends here" -Colour Yellow
        } else {
            Log "STOPPED PART-WAY: $why; $Before would stop the app, so the automated fix ends here, after the steps it has run" -Colour Yellow
            # Nothing after this exit would start the service again.
            Start-ServiceStoppedByRun
        }
        Save-Log
        exit $ExitCode
    }
}

function Get-AutomatedFixBlocker {
    <#
    .SYNOPSIS
        Why an automated (-Quiet) repair must not run now, or $null when
        nothing holds it back.
    .DESCRIPTION
        A repair stops the app, and every Claude Code session the app started
        stops with it. Manual runs (no -Quiet) are not asked: the user wants
        the fix. The checks:

        1. A Claude Code CLI the app started is running: a project thread or
           Code session. Added in 6.3.0, and it fails closed: when the
           sessions cannot be listed, the repair waits.
        2. A Claude Code transcript written in the last 35 minutes, the same
           question the monitor asks before its own repair. Added in 6.3.0.
        3. A claude process using CPU: over 100 ms in 500 ms.
        4. A VM log written in the last 120 seconds.
        5. Keyboard or mouse input in the last 3 minutes.

        Until 6.3.0 only checks 3 to 5 ran. A project thread waiting on a
        25-minute command uses almost no CPU and writes nothing, so an
        unattended -Quiet run passed all three and ended it.
    #>
    $appSession = Get-AppSessionBlocker
    if ($appSession) { return $appSession }

    if (Test-ClaudeCodeTranscriptRecent -Minutes 35) {
        return "a Claude Code transcript was written in the last 35 minutes"
    }

    # Check 3: any claude process burning CPU. The Claude Code CLI counts as
    # well: a busy terminal session is a reason to wait too.
    foreach ($cp in @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)) {
        try {
            $cpu1 = $cp.TotalProcessorTime.TotalMilliseconds
            Start-Sleep -Milliseconds 500
            $cp.Refresh()
            $cpu2 = $cp.TotalProcessorTime.TotalMilliseconds
            if (($cpu2 - $cpu1) -gt 100) { return "claude process $($cp.Id) is using CPU" }
        } catch { $null = $_ }
    }

    # Check 4: VM log active within 120s (Code may be thinking)
    #
    # Every VM log in the app's log folder, read now (Get-ClaudeVmLogAge). The
    # list before discovery led with C:\ProgramData\Claude\Logs\coworkd.log, on
    # a hardcoded drive letter, in a tree that does not exist.
    $ageSec = Get-ClaudeVmLogAge -LogDir $script:ClaudeEnv.LogDir
    if ($null -ne $ageSec -and $ageSec -lt 120) { return "a VM log was written $([int]$ageSec)s ago" }

    # Check 5: user input within 3 minutes (interactive sessions only)
    try {
        $sessionId = (Get-Process -Id $PID -ErrorAction Stop).SessionId
        if ($sessionId -gt 0) {
            Add-Type -ErrorAction SilentlyContinue -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class FixActivityCheck {
    [StructLayout(LayoutKind.Sequential)]
    public struct LASTINPUTINFO { public uint cbSize; public uint dwTime; }
    [DllImport("user32.dll")] public static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);
}
'@
            $lastInput = New-Object FixActivityCheck+LASTINPUTINFO
            $lastInput.cbSize = [uint32][System.Runtime.InteropServices.Marshal]::SizeOf($lastInput)
            if ([FixActivityCheck]::GetLastInputInfo([ref]$lastInput)) {
                # TickCount is a SIGNED 32-bit value that goes negative once
                # uptime passes 24.9 days, while dwTime is unsigned. Doing
                # the subtraction directly then produces a hugely negative
                # idle time, which reads as "user is active", so every
                # -Quiet run on a long-uptime machine exited having done
                # nothing. TickCount64 would avoid this but does not exist
                # on Windows PowerShell 5.1, which is the engine the
                # elevation path always lands in. Doing the arithmetic in
                # unsigned 32-bit space makes the wrap cancel out instead.
                # 0xFFFFFFFFL, with the L. A bare 0xFFFFFFFF is parsed by
                # PowerShell as Int32 -1, so the mask is a no-op on a
                # negative TickCount and the [uint32] cast then throws
                # "Value was either too large or too small for a UInt32",
                # in precisely the past-24.9-days case this exists for.
                $nowTicks  = [uint32]([Environment]::TickCount -band 0xFFFFFFFFL)
                $lastTicks = [uint32]$lastInput.dwTime
                if ($nowTicks -ge $lastTicks) {
                    $idleMs = [long]$nowTicks - [long]$lastTicks
                } else {
                    $idleMs = [long]4294967296 - [long]$lastTicks + [long]$nowTicks
                }
                if ($idleMs -lt 180000) { return "keyboard or mouse input $([int]($idleMs / 1000))s ago" }
            }
        }
    } catch { $null = $_ }
    return $null
}

# ====================================================================
# MAIN
# ====================================================================
try {

# -- Header ----------------------------------------------------------
Write-Host ""
Write-Host "  +-------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |  CLAUDE DESKTOP / COWORK, RESET & FIX     |" -ForegroundColor Cyan
Write-Host "  |  v$ToolkitVersion                                   |" -ForegroundColor DarkGray
Write-Host "  +-------------------------------------------+" -ForegroundColor Cyan
Write-Host ""

# -- Prevent concurrent Fix runs ----------------------------------------
# Deliberately carries no version. Two different toolkit versions running at
# once against the same service and the same VHDX files is exactly what this
# mutex exists to prevent, and a versioned name would let them straight past
# each other. The old name was still pinned at v4.8.
# A Global\ mutex needs SeCreateGlobalPrivilege, which a non-elevated run does
# not have. That failure used to be swallowed whole, leaving no lock at all, so
# two runs could overlap across the same service and the same VHDX files, which
# is the one thing this is here to prevent. Fall back to a session-local name,
# which still catches the common case of someone double-clicking twice.
$fixMutex     = $null
$fixMutexHeld = $false
foreach ($fixMutexName in @("Global\ClaudeDesktopFix", "Local\ClaudeDesktopFix")) {
    try {
        $candidate = [System.Threading.Mutex]::new($false, $fixMutexName)
        if ($candidate.WaitOne(0)) {
            $fixMutex     = $candidate
            $fixMutexHeld = $true
        } else {
            try { $candidate.Dispose() } catch { $null = $_ }
            Log "Another Fix instance is already running, exiting" -Colour DarkGray
            Save-Log
            # A manual run waits for a key before it ends, as every other
            # manual run does, so the window does not close before this can
            # be read. The launchers do not pause again after it. Until 6.3.0
            # it closed at once, and Fix-ClaudeDesktop.bat then reported a
            # PowerShell failure.
            if (-not $Quiet) { Wait-ForAnyKey }
            # Held back, not done: see $script:BlockedExitCode.
            exit $script:BlockedExitCode
        }
        break
    } catch {
        $fixMutex = $null
    }
}
if (-not $fixMutexHeld) {
    Log "Could not take a concurrency lock; a second Fix run could overlap this one" -Colour DarkYellow
}

if (-not $script:IsAdmin) {
    Log "Running without admin (limited service control)" -Colour DarkGray
}
if ($WhatIfPreference) {
    Log "DRY RUN, no changes will be made" -Colour Yellow
}
# Where this run reads the app's logs and finds its VM cache. Both were fixed
# paths under %APPDATA%\Claude until 6.2.0, and on app 6.18355 neither is there,
# so a run that found nothing to read or purge had no way to say why.
# claude-env.json in fix-logs lists every candidate that was passed over.
Log "App data : $($script:ClaudeEnv.DataRoot) ($($script:ClaudeEnv.DataRootSource))" -Colour DarkGray
Log "App logs : $($script:ClaudeEnv.LogDir) ($($script:ClaudeEnv.LogDirSource))" -Colour DarkGray
Write-Host ""

# -- Close mode: clean shutdown only, no relaunch ----------------
if ($Close) {
    Log "CLOSE MODE, performing clean shutdown" -Colour Yellow
    Log ""
    # 1) Stop the service -- this triggers graceful VM shutdown
    Log "[1/5] Stopping CoworkVMService (graceful)..." -Colour Yellow
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($svc -and $svc.Status -eq "Running") {
        if ($script:IsAdmin) {
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $svcCtl = New-Object System.ServiceProcess.ServiceController($ServiceName)
            try { $svcCtl.Stop() } catch { $null = $_ }
            $maxWait = 45
            $stopped = $false
            while ($sw.Elapsed.TotalSeconds -lt $maxWait) {
                Start-Sleep -Seconds 3
                $elapsed = [math]::Round($sw.Elapsed.TotalSeconds)
                $curSvc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
                if (-not $curSvc -or $curSvc.Status -eq "Stopped") {
                    $stopped = $true
                    Log "Service stopped gracefully (${elapsed}s)" -Colour Green -Indent
                    break
                }
                # $null is no answer, which says nothing about the VM. Until
                # 6.3.0 it read as "VM gone".
                $hcsOut = Get-HcsList
                if ($null -eq $hcsOut) {
                    Log "Waiting for service... (${elapsed}s, hcsdiag did not answer)" -Colour DarkGray -Indent
                } elseif ($hcsOut -match "cowork-vm") {
                    Log "Waiting for VM shutdown... (${elapsed}s)" -Colour DarkGray -Indent
                } else {
                    Log "VM gone, waiting for service... (${elapsed}s)" -Colour DarkGray -Indent
                }
            }
            $sw.Stop()
            if (-not $stopped) {
                Log "Service still running after ${maxWait}s, force-killing" -Colour DarkYellow -Indent
                Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue | Stop-Process -Force
                Start-Sleep -Seconds 2
            }
        } else {
            Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue | Stop-Process -Force
            Log "Killed service process (no admin)" -Colour DarkGray -Indent
        }
    } else {
        Log "Service not running" -Colour DarkGray -Indent
    }
    # 2) Kill the Claude UI, THEN clean HCS.
    #
    # These two were the other way round, which meant the compute system was
    # torn down while Claude was still running and Claude simply recreated it a
    # moment later. The main repair path has always had this order right.
    Log "[2/5] Terminating Claude processes..." -Colour Yellow
    # The desktop app only. The Claude Code CLI runs as "claude" too, and this
    # used to stop the copies running project threads and any terminal session
    # along with the app. See Get-ClaudeProcessInventory.
    $procInventory = @(Get-ClaudeProcessInventory -Environment $script:ClaudeEnv)
    $procs = @($procInventory | Where-Object { $_.IsDesktop } | ForEach-Object { $_.Process })
    $procsSpared = @($procInventory | Where-Object { -not $_.IsDesktop }).Count
    if ($procs.Count -gt 0) {
        $procs | Stop-Process -Force -ErrorAction SilentlyContinue
        Log "Killed $($procs.Count) Claude process(es)" -Colour Green -Indent
    } else {
        Log "No Claude processes found" -Colour DarkGray -Indent
    }
    if ($procsSpared -gt 0) {
        Log "Left $procsSpared other process(es) named claude alone (Claude Code CLI, or not identifiable as the app)" `
            -Colour DarkGray -Indent
    }

    # 3) Clean up any remaining HCS compute systems
    Log "[3/5] Cleaning HCS compute systems..." -Colour Yellow
    if ($script:IsAdmin) {
        Start-Sleep -Seconds 2
        try {
            $cleaned = Close-StaleHcsVms -Action kill
            if ($cleaned -gt 0) {
                Log "Killed $cleaned remaining compute system(s)" -Colour Green -Indent
            } elseif ($script:LastHcsKillsSent -gt 0) {
                # The lines above say which were still listed, or that the
                # check after the kill did not answer. Until 6.3.0 this said
                # "No remaining compute systems".
                Log "No compute system confirmed gone" -Colour DarkYellow -Indent
            } elseif ($script:LastHcsListed -gt 0) {
                # Listed, and no kill went out: -WhatIf, or every kill threw.
                # Until 6.3.0 this said "No remaining compute systems".
                Log "$($script:LastHcsListed) compute system(s) listed, none killed" -Colour DarkYellow -Indent
            } elseif ($script:LastHcsListOk) {
                Log "No remaining compute systems" -Colour DarkGray -Indent
            } else {
                Log "hcsdiag did not answer, so this check was not run" -Colour DarkGray -Indent
            }
            # Kills went out for some of those listed and not for the rest,
            # which are still there. Until 6.3.0 nothing said so.
            $notSent = $script:LastHcsListed - $script:LastHcsKillsSent
            if ($script:LastHcsKillsSent -gt 0 -and $notSent -gt 0) {
                Log "$notSent compute system(s) listed, not killed" -Colour DarkYellow -Indent
            }
        } catch {
            Log "HCS cleanup error: $($_.Exception.Message)" -Colour DarkYellow -Indent
        }
    } else {
        Log "Skipping (no admin)" -Colour DarkGray -Indent
    }
    # 4) Restart the service so it is ready for next launch
    #    Without this, Windows will not auto-start the service via
    #    the named pipe trigger (manually-stopped services are ignored).
    Log "[4/5] Restarting service (idle, ready for next launch)..." -Colour Yellow
    # No longer conditional on this run having been the one to stop it. A
    # service that was already stopped on entry is exactly the case this step
    # exists for: Windows will not honour the named-pipe start trigger for a
    # manually stopped service, so leaving it down is what breaks the next
    # launch. The $serviceWasStopped flag that used to gate this is gone, since
    # nothing else read it.
    if ($script:IsAdmin) {
        try {
            Start-Service -Name $ServiceName -ErrorAction Stop
            # Seconds, stepped 2 at a time. The bound used to read 15, which
            # with a 2-second step actually allowed 16.
            $svcPoll = 0
            $svcPollMax = 16
            $svcOk = $false
            while ($svcPoll -lt $svcPollMax) {
                Start-Sleep -Seconds 2
                $svcPoll += 2
                $curSvc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
                if ($curSvc -and $curSvc.Status -eq "Running") {
                    $svcOk = $true
                    break
                }
            }
            if ($svcOk) {
                Log "Service running (idle)" -Colour Green -Indent
            } else {
                Log "Service did not reach Running state, may need Fix on relaunch" -Colour DarkYellow -Indent
            }
        } catch {
            Log "Service restart failed: $($_.Exception.Message)" -Colour DarkYellow -Indent
            Log "You may need to run Fix on next launch" -Colour DarkGray -Indent
        }
    } else {
        Log "Skipping (no admin)" -Colour DarkGray -Indent
        Log "Claude will start the service itself on next launch" -Colour DarkGray -Indent
    }
    # 5) Verify clean state
    Log "[5/5] Verifying clean state..." -Colour Yellow
    Start-Sleep -Seconds 1
    $remainingProcs = @(Get-ClaudeDesktopProcess -Environment $script:ClaudeEnv)
    $remainingSvc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    $remainingHcs = $false
    $hcsUnknown   = $false
    try {
        $hcsCheck = Get-HcsList
        if ($null -eq $hcsCheck) { $hcsUnknown = $true }
        else { $remainingHcs = ($hcsCheck -match "cowork-vm") }
    } catch { $hcsUnknown = $true }
    if ($hcsUnknown) {
        Log "Whether a compute system is left is not known: hcsdiag did not answer" -Colour DarkGray -Indent
    }
    $svcRunning = $remainingSvc -and $remainingSvc.Status -eq "Running"
    # Whether the branch that calls the shutdown clean ran. The summary after
    # it says the app is shut down only then.
    $closeClean = $false
    if ($remainingProcs.Count -eq 0 -and -not $remainingHcs) {
        if ($hcsUnknown) {
            # Not "complete": hcsdiag did not say whether a compute system is
            # left. Until 6.3.0 this said "Clean shutdown complete" here too.
            $svcSaid = if ($svcRunning) { 'the service is running' } else { 'the service is not running' }
            Log "No app processes are left and $svcSaid; whether a compute system is left is not known" -Colour DarkYellow -Indent
        } elseif ($svcRunning) {
            Log "Clean shutdown complete (service idle, ready for relaunch)" -Colour Green -Indent
            $closeClean = $true
        } else {
            Log "Clean shutdown complete (service not running, may need Fix on relaunch)" -Colour DarkYellow -Indent
            $closeClean = $true
        }
    } else {
        if ($remainingProcs.Count -gt 0) { Log "Warning: $($remainingProcs.Count) Claude processes still running" -Colour DarkYellow -Indent }
        if ($remainingHcs) { Log "Warning: HCS compute system still present" -Colour DarkYellow -Indent }
    }
    # Until 6.3.0 this said in green that the app was shut down, and with the
    # service running that a relaunch would work at once, after either
    # warning above and when hcsdiag had not answered.
    Write-Host ""
    if ($closeClean) {
        Write-Host "  Claude Desktop is shut down." -ForegroundColor Green
        if ($svcRunning) {
            Write-Host "  Service is idle and ready, relaunch should work immediately." -ForegroundColor Green
        } else {
            Write-Host "  Service is not running, you may need to run Fix after relaunch." -ForegroundColor DarkYellow
        }
    } elseif ($hcsUnknown -and $remainingProcs.Count -eq 0) {
        Write-Host "  Claude Desktop's processes are stopped, but whether a compute system is left is not known." -ForegroundColor DarkYellow
        Write-Host "  If the relaunch fails, run Fix." -ForegroundColor DarkYellow
    } else {
        Write-Host "  Claude Desktop is not fully shut down: see the warnings above." -ForegroundColor DarkYellow
        Write-Host "  If the relaunch fails, run Fix." -ForegroundColor DarkYellow
    }
    Write-Host ""
    Save-Log
    # Nulled after disposal. The finally block at the end of the script releases
    # and disposes this same handle, and on an early exit like this one that
    # second pass raised ObjectDisposedException into a bare catch.
    if ($fixMutex) {
        try { $fixMutex.ReleaseMutex() } catch { $null = $_ }
        try { $fixMutex.Dispose() } catch { $null = $_ }
        $fixMutex = $null
    }
    if (-not $Quiet) { Wait-ForAnyKey }
    exit 0
}

# ====================================================================
# INTERACTIVE MENU -- shown when run manually without -Mode or -Quiet
# ====================================================================
$script:SelectedMode = $Mode   # may be empty
if (-not $Quiet -and -not $Mode -and [Environment]::UserInteractive) {
    # Try PromptForChoice first (works in full console hosts)
    $menuSuccess = $false
    try {
        $modeTitle   = "  Select repair mode"
        $modeMessage = "  What kind of fix do you want to run?"
        $modeChoices = [System.Management.Automation.Host.ChoiceDescription[]]@(
            (New-Object System.Management.Automation.Host.ChoiceDescription "&Quick Fix",   "Restart services + basic repair (Steps 1-5, skip cache purge)"),
            (New-Object System.Management.Automation.Host.ChoiceDescription "&Deep Fix",    "Full nuclear reset (all steps including cache purge)"),
            (New-Object System.Management.Automation.Host.ChoiceDescription "&Smart Fix",   "Try quick first, escalate to deep if needed (recommended)"),
            (New-Object System.Management.Automation.Host.ChoiceDescription "D&iagnostic",  "Health check only, no changes"),
            (New-Object System.Management.Automation.Host.ChoiceDescription "&Cancel",      "Exit without doing anything")
        )
        $modeDefault = 2  # Smart Fix
        $modeResult  = $host.UI.PromptForChoice($modeTitle, $modeMessage, $modeChoices, $modeDefault)
        $menuSuccess = $true

        switch ($modeResult) {
            0 { $script:SelectedMode = "Quick" }
            1 { $script:SelectedMode = "Deep" }
            2 { $script:SelectedMode = "Smart" }
            3 { $script:SelectedMode = "Diagnostic" }
            4 {
                Log "Cancelled by user" -Colour DarkGray
                Save-Log
                exit 0
            }
        }

        # Second menu: options
        Write-Host ""
        $optTitle   = "  Options"
        $optMessage = "  Toggle any options, then Continue:"
        $optDone = $false
        while (-not $optDone) {
            $kcLabel   = if ($KeepCache) { "&Keep cache [ON]" } else { "&Keep cache [off]" }
            $slLabel   = if ($SkipLaunch) { "&Skip relaunch [ON]" } else { "&Skip relaunch [off]" }
            $wiLabel   = if ($WhatIfPreference) { "&WhatIf mode [ON]" } else { "&WhatIf mode [off]" }
            $optChoices = [System.Management.Automation.Host.ChoiceDescription[]]@(
                (New-Object System.Management.Automation.Host.ChoiceDescription $kcLabel,  "Toggle cache preservation"),
                (New-Object System.Management.Automation.Host.ChoiceDescription $slLabel,  "Toggle Claude relaunch"),
                (New-Object System.Management.Automation.Host.ChoiceDescription $wiLabel,  "Toggle dry-run mode"),
                (New-Object System.Management.Automation.Host.ChoiceDescription "&Continue", "Accept current options")
            )
            $optResult = $host.UI.PromptForChoice($optTitle, $optMessage, $optChoices, 3)
            switch ($optResult) {
                0 { $KeepCache = -not $KeepCache }
                1 { $SkipLaunch = -not $SkipLaunch }
                2 { $WhatIfPreference = -not $WhatIfPreference }
                3 { $optDone = $true }
            }
        }
    } catch {
        # Fallback: simple Read-Host menu for hosts that don't support PromptForChoice
        if (-not $menuSuccess) {
            Write-Host ""
            Write-Host "  Select repair mode:" -ForegroundColor Cyan
            Write-Host "  1) Quick Fix    : Restart services + basic repair"
            Write-Host "  2) Deep Fix     : Full nuclear reset (cache purge)"
            Write-Host "  3) Smart Fix    : Quick first, escalate if needed (recommended)"
            Write-Host "  4) Diagnostic   : Health check only, no changes"
            Write-Host "  C) Cancel"
            Write-Host ""
            $choice = Read-Host "  Selection [3]"
            if (-not $choice) { $choice = "3" }
            switch ($choice.Trim()) {
                "1" { $script:SelectedMode = "Quick" }
                "2" { $script:SelectedMode = "Deep" }
                "3" { $script:SelectedMode = "Smart" }
                "4" { $script:SelectedMode = "Diagnostic" }
                { $_ -eq "C" -or $_ -eq "c" } {
                    Log "Cancelled by user" -Colour DarkGray
                    Save-Log
                    exit 0
                }
                default { $script:SelectedMode = "Smart" }
            }
        }
    }

    # Summary and final confirm
    Write-Host ""
    Write-Host "  +-------------------------------------------+" -ForegroundColor Cyan
    Write-Host "  |  SELECTED OPTIONS                         |" -ForegroundColor Cyan
    Write-Host "  +-------------------------------------------+" -ForegroundColor Cyan
    Write-Host "    Mode:          $($script:SelectedMode)" -ForegroundColor White
    Write-Host "    Keep cache:    $(if ($KeepCache) { 'Yes' } else { 'No' })" -ForegroundColor White
    Write-Host "    Skip relaunch: $(if ($SkipLaunch) { 'Yes' } else { 'No' })" -ForegroundColor White
    Write-Host "    WhatIf:        $(if ($WhatIfPreference) { 'Yes' } else { 'No' })" -ForegroundColor White
    Write-Host ""
    # Guarded, like the PromptForChoice block above always was. In a host with
    # redirected stdin Read-Host throws, and unguarded it aborted the whole run
    # rather than the prompt. An unreadable prompt means proceed, which is the
    # documented default answer.
    $confirm = ""
    try {
        $confirm = Read-Host "  Proceed? (Y/n)"
    } catch {
        Write-Host "  (no console input available, proceeding)" -ForegroundColor DarkGray
    }
    if ($confirm -and $confirm.Trim() -match "^[Nn]") {
        Log "Cancelled by user" -Colour DarkGray
        Save-Log
        exit 0
    }
    Write-Host ""
}

# Default mode when none selected
if (-not $script:SelectedMode) { $script:SelectedMode = "Smart" }

# ====================================================================
# DIAGNOSTIC MODE -- report health and exit
# ====================================================================
if ($script:SelectedMode -eq "Diagnostic") {
    Log "Running diagnostic check (no changes)..." -Colour Cyan
    Write-Host ""

    # Service status
    $diagSvc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($diagSvc) {
        $diagCol = if ($diagSvc.Status -eq "Running") { "Green" } else { "Yellow" }
        Log "  $ServiceName : $($diagSvc.Status)" -Colour $diagCol
    } else {
        Log "  $ServiceName : Not installed" -Colour Red
    }

    # vmcompute status
    $diagVmc = Get-Service -Name "vmcompute" -ErrorAction SilentlyContinue
    if ($diagVmc) {
        $diagCol = if ($diagVmc.Status -eq "Running") { "Green" } else { "Yellow" }
        Log "  vmcompute     : $($diagVmc.Status)" -Colour $diagCol
    }

    # HCS errors
    $diagHcs = Test-RecentHcsErrors
    if ($diagHcs -eq "shutdown_stale") {
        Log "  HCS health    : Shutdown stale (0xC037010D property query bug)" -Colour Yellow
    } elseif ($diagHcs -eq "construct_failure") {
        Log "  HCS health    : Construct failure detected (0x800707DE)" -Colour Red
    } elseif ($diagHcs -eq "guest_connect_failure") {
        Log "  HCS health    : Guest connection failure (isGuestConnected timeout)" -Colour Red
    } elseif ($diagHcs -eq "hcs_error") {
        Log "  HCS health    : Errors detected" -Colour Yellow
    } else {
        Log "  HCS health    : Clean" -Colour Green
    }

    # Claude processes
    $diagProcs = @(Get-ClaudeProcessInventory -Environment $script:ClaudeEnv)
    $diagDesktop = @($diagProcs | Where-Object { $_.IsDesktop }).Count
    Log "  Claude procs  : $diagDesktop app, $($diagProcs.Count - $diagDesktop) other (Claude Code CLI or unidentified)" -Colour Cyan

    # VM cache
    foreach ($cp in @(
        @{ Path = $VmCachePath; Label = "claude-code-vm" },
        @{ Path = $BundlePath;  Label = "vm_bundles" }
    )) {
        if (Test-Path $cp.Path) {
            $sz = (Get-ChildItem $cp.Path -Recurse -ErrorAction SilentlyContinue |
                   Measure-Object -Property Length -Sum).Sum
            Log "  $($cp.Label) : $([math]::Round($sz/1MB,1)) MB" -Colour DarkGray
        } else {
            Log "  $($cp.Label) : Not present" -Colour DarkGray
        }
    }

    # Recent event log errors
    try {
        $diagEvFilter = @{
            LogName      = "Application"
            ProviderName = "CoworkVMService"
            Level        = 2
            # The UTC time as round-trip text: see Get-FixWinEvent.
            StartTime    = [datetime]::UtcNow.AddHours(-1).ToString('o')
        }
        $diagEvents = @(Get-FixWinEvent -Filter $diagEvFilter -MaxEvents 5)
        if ($diagEvents) {
            Write-Host ""
            Log "  Recent service errors (last hour):" -Colour DarkYellow
            foreach ($de in $diagEvents) {
                $deTime = "{0:HH:mm}" -f $de.TimeCreated
                $deMsg  = ($de.Message -split "`n")[0]
                Log "    [$deTime] $deMsg" -Colour DarkGray
            }
        }
    } catch {
        Log "  Service errors: the Application log could not be read ($($_.Exception.Message))" -Colour DarkGray
    }

    # HCS state via hcsdiag (v4.8.0)
    if ($script:IsAdmin) {
        try {
            $diagHcsList = Get-HcsList
            # '' is an answer: nothing listed. Until 6.3.0 it read as
            # "Unable to query".
            if ($null -ne $diagHcsList) {
                # Counts DISTINCT GUIDs, not occurrences of the name.
                #
                # "hcsdiag list" prints the compute system name twice for a
                # single VM: once on its own line and again at the end of the
                # detail line. Counting matches of "cowork-vm" therefore
                # reported 2 instances when exactly one was running.
                $diagGuidPat = '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})'
                $diagSeen = @{}
                foreach ($diagLine in ($diagHcsList -split "`r?`n")) {
                    if ($diagLine -match "cowork-vm") {
                        if ($diagLine -match $diagGuidPat) { $diagSeen[$Matches[1]] = $true }
                    }
                }
                $vmEntries = $diagSeen.Count
                if ($vmEntries -gt 0) {
                    Log "  HCS VMs       : $vmEntries cowork-vm instance(s)" -Colour $(if ($vmEntries -gt 1) { "Yellow" } else { "Green" })
                } else {
                    Log "  HCS VMs       : None" -Colour DarkGray
                }
            } else {
                Log "  HCS VMs       : Unable to query (hcsdiag timed out, failed or is not installed)" -Colour DarkGray
            }
        } catch { $null = $_ }
    }

    # 0xC037010D and 0x800707DE frequency.
    #
    # The filter is built outside both try blocks. It used to be created inside
    # the first one, and the second block then reused it: if the first block
    # threw before the assignment, the second referenced an unset variable and
    # failed under StrictMode.
    $diagEventFilter = @{
        LogName   = "Microsoft-Windows-Hyper-V-Compute-Operational"
        # The UTC time as round-trip text: see Get-FixWinEvent.
        StartTime = [datetime]::UtcNow.AddHours(-24).ToString('o')
    }
    try {
        $diagShutdownEvents = @(Get-FixWinEvent -Filter $diagEventFilter |
            Where-Object { $_.Message -match "0xC037010D" })
        # The age in UTC: local time jumps an hour when summer time starts or
        # ends. Until 6.3.0 it was local.
        $hourAgoUtc = [datetime]::UtcNow.AddHours(-1)
        $last1h = @($diagShutdownEvents | Where-Object { $_.TimeCreated -and $_.TimeCreated.ToUniversalTime() -gt $hourAgoUtc }).Count
        $last24h = $diagShutdownEvents.Count
        $statusColour = if ($last1h -gt 10) { "Red" } elseif ($last1h -gt 3) { "Yellow" } else { "Green" }
        Log "  Shutdown fails: $last1h (1h) / $last24h (24h)" -Colour $statusColour
    } catch {
        Log "  Shutdown fails: the Hyper-V Compute log could not be read ($($_.Exception.Message))" -Colour DarkGray
    }

    # 0x800707DE frequency (v4.8.0)
    try {
        $diagConstructEvents = @(Get-FixWinEvent -Filter $diagEventFilter |
            Where-Object { $_.Message -match "0x800707DE" })
        if ($diagConstructEvents.Count -gt 0) {
            Log "  Construct fails: $($diagConstructEvents.Count) (24h)" -Colour Red
        }
    } catch {
        Log "  Construct fails: the Hyper-V Compute log could not be read ($($_.Exception.Message))" -Colour DarkGray
    }

    # Session file count.
    #
    # Reported, not judged. This used to turn red past 1000 files, which reads
    # as a fault, but a session file count says nothing about HCS or VM health.
    # A working install accumulates thousands of them in normal use.
    $sessionDir = $script:ClaudeEnv.SessionsDir
    if ($sessionDir -and (Test-Path $sessionDir)) {
        $sessionFiles = @(Get-ChildItem $sessionDir -Recurse -File -ErrorAction SilentlyContinue)
        $sessionSum = ($sessionFiles | Measure-Object -Property Length -Sum).Sum
        if (-not $sessionSum) { $sessionSum = 0 }
        Log "  Session files : $($sessionFiles.Count) ($([math]::Round($sessionSum/1MB,1)) MB, -PurgeSessions to trim)" -Colour DarkGray
    }

    # vmcompute handle count (v4.8.0)
    try {
        $diagVmcompute = Get-Process -Name "vmcompute" -ErrorAction SilentlyContinue
        if ($diagVmcompute) {
            $hc = $diagVmcompute.HandleCount
            $hcCol = if ($hc -gt 10000) { "Red" } elseif ($hc -gt 5000) { "Yellow" } else { "Green" }
            Log "  vmcompute     : $hc handles" -Colour $hcCol
        }
    } catch { $null = $_ }

    # CoworkVMService recovery config.
    #
    # Read from the registry by discovery, not by matching the literal string
    # "RESTART" in sc.exe console output. That match fails on any Windows whose
    # display language is not English, reporting recovery as unconfigured on a
    # machine where it is configured correctly. The service name also came from
    # a hardcoded literal rather than $ServiceName.
    try {
        if ($script:ClaudeEnv.ServiceRecoveryConfigured) {
            Log "  Svc recovery  : Configured" -Colour Green
        } else {
            Log "  Svc recovery  : NOT CONFIGURED (run Prevent to fix)" -Colour Yellow
        }
    } catch { $null = $_ }

    # Defender exclusion completeness (v4.8.0)
    try {
        $procExcl = (Get-MpPreference -ErrorAction SilentlyContinue).ExclusionProcess
        $requiredProcs = @("vmwp.exe", "vmms.exe", "vmcompute.exe", "cowork-svc.exe")
        $missingExcl = @($requiredProcs | Where-Object { $procExcl -notcontains $_ })
        if ($missingExcl.Count -gt 0) {
            Log "  Defender procs: MISSING $($missingExcl -join ', ') (run Prevent to fix)" -Colour Yellow
        } else {
            Log "  Defender procs: All exclusions present" -Colour Green
        }
    } catch { $null = $_ }

    # -- Support bundle ------------------------------------------------
    #
    # A user reporting a Cowork failure gets asked for logs, and the honest
    # answer is that Claude's own logs are close to empty when the service
    # never started: there was nothing running to write them. What actually
    # records the failure is the Windows Service Control Manager channel, and
    # nobody thinks to look there.
    #
    # So collect it. Anthropic's own troubleshooting note points at
    # Help > Troubleshooting > Show Logs, which regenerates three snapshot
    # files at the moment it is clicked; this gathers the same folder plus the
    # event log, without the user having to know any of that.
    #
    # Copies, never parses. supported-features-info.json is documented as
    # showing which check failed, but on the builds seen here it holds
    # codenamed feature flags and nothing about virtualization, so branching on
    # it would be inventing a contract that does not exist.
    Write-Host ""
    Log "Collecting a support bundle..." -Colour Cyan
    $bundleDir = $null
    try {
        $stamp     = Get-Date -Format "yyyyMMdd_HHmmss"
        $bundleDir = Join-Path $LogDir "support_$stamp"
        New-Item -ItemType Directory -Path $bundleDir -Force -ErrorAction Stop | Out-Null

        # Enumerated, not listed. The first version of this hardcoded eight
        # filenames read off the machine it was written on, and the very first
        # real export it was compared against had coworkd-user.log and
        # vm-info.json instead, neither of them on the list. vm-info.json was
        # the single most useful file in that export: it carries the VM bundle
        # size, whether it finished downloading, and isGuestConnected.
        #
        # Log filenames vary by app build. Take everything and cap it, rather
        # than guessing which names this build uses. Same authoring-time-guess
        # bug that Get-ClaudeEnvironment exists to end, reintroduced by hand.
        $wanted = @()
        if ($srcLogsProbe = $script:ClaudeEnv.LogDir) {
            try {
                $wanted = @(Get-ChildItem $srcLogsProbe -File -ErrorAction Stop |
                            Where-Object { $_.Extension -in @('.log', '.json', '.txt') } |
                            Sort-Object Length |
                            Select-Object -ExpandProperty Name)
            } catch { $null = $_ }
        }
        if ($wanted.Count -eq 0) {
            # Fall back to the known names so an unreadable folder still
            # produces something rather than an empty bundle.
            $wanted = @(
                'system-info.txt', 'supported-features-info.json', 'gpu-info.json',
                'vm-info.json', 'mcp-info.json',
                'main.log', 'cowork-service.log', 'coworkd.log', 'coworkd-user.log',
                'cowork_vm_node.log', 'cowork_host_loop_debug.log'
            )
        }
        # Whole-file copies came to 19.6 MB on the machine this was built on,
        # nearly all of it main.log and cowork_vm_node.log. Discord caps
        # attachments at 5 MB, so the bundle is built to a 4 MB budget and the
        # folder is zipped afterwards. The zip is what a person sends.
        #
        # A per-file cap does not achieve this: five logs at 2 MB each is 10 MB
        # no matter how modest each one looks. The budget is shared instead.
        # Small metadata files are copied whole, since together they are a few
        # KB, and whatever is left is split evenly between the large logs that
        # actually exist. A startup failure is at the END of a log, so each one
        # keeps its most recent bytes.
        $bundleBudget = 4MB
        $reserve      = 256KB   # service-events.txt and windows-features.txt
        # Classified by size, not by name. The name list was wrong on the first
        # real machine it met, and there is no reason to guess when the size is
        # right there. Anything under 128 KB is cheap enough to take whole.
        $smallCutoff  = 128KB

        $copied   = @()
        $missing  = @()
        $trimmed  = @()
        $srcLogs  = $script:ClaudeEnv.LogDir

        # Pass 1: find what is actually there, and how big.
        $present = @{}
        foreach ($w in $wanted) {
            $p = if ($srcLogs) { Join-Path $srcLogs $w } else { $null }
            if ($p -and (Test-Path $p)) {
                try { $present[$w] = (Get-Item $p -ErrorAction Stop).Length }
                catch { $missing += $w }
            } else {
                $missing += $w
            }
        }

        # Pass 2: work out the share for the large logs.
        #
        # Enumerating the folder turned up 40 files on the machine this was
        # written on, 32 of them per-MCP-server logs. Splitting the budget
        # evenly gave each one 114 KB, which meant cowork-service.log, the file
        # that actually explains a Cowork failure, got the same slice as the log
        # for some unrelated MCP server. So rank them.
        #
        # The pattern is deliberately broad rather than a list of names. It
        # catches log names this build uses that no one here has seen, which is
        # the entire reason the folder is enumerated instead of listed.
        # unknown-window.log earns its place here. The GPU-process crash that
        # kills the whole app (exit code 101457950 = 0x060C201E, upstream issue
        # 80444) leaves its entire signature in that file and nothing but a
        # single "GPU process gone" line in main.log. A bundle that trimmed it
        # to the long-tail share would drop the only evidence of the cause.
        $priorityPattern = '(?i)cowork|^vm-|^main\.|^unknown-window|system-info|supported-features|gpu-info'
        $smallNames = @($present.Keys | Where-Object { $present[$_] -le $smallCutoff })
        $bigNames   = @($present.Keys | Where-Object { $present[$_] -gt $smallCutoff })
        $bigTop     = @($bigNames | Where-Object { $_ -match $priorityPattern })
        $bigRest    = @($bigNames | Where-Object { $_ -notmatch $priorityPattern })

        $smallBytes = 0
        foreach ($s in $smallNames) { $smallBytes += $present[$s] }
        $logBudget = $bundleBudget - $reserve - $smallBytes
        if ($logBudget -lt 0) { $logBudget = [long]0 }

        # Four fifths to the files that answer the question, one fifth to
        # everything else, so the long tail is represented without drowning it.
        $shareTop  = [long]0
        $shareRest = [long]0
        if ($bigTop.Count -gt 0)  { $shareTop  = [long](($logBudget * 0.8) / $bigTop.Count) }
        if ($bigRest.Count -gt 0) { $shareRest = [long](($logBudget * 0.2) / $bigRest.Count) }
        if ($bigTop.Count -eq 0 -and $bigRest.Count -gt 0) {
            $shareRest = [long]($logBudget / $bigRest.Count)
        }
        if ($bigTop.Count -gt 0 -and $shareTop -lt 64KB)   { $shareTop  = [long]64KB }
        if ($bigRest.Count -gt 0 -and $shareRest -lt 16KB) { $shareRest = [long]16KB }

        # Pass 3: copy, tailing anything over its share.
        foreach ($w in $present.Keys) {
            $p    = Join-Path $srcLogs $w
            $dest = Join-Path $bundleDir $w
            $len  = $present[$w]
            $cap  = if ($len -le $smallCutoff) { [long]$len }
                    elseif ($w -match $priorityPattern) { $shareTop }
                    else { $shareRest }
            try {
                if ($len -le $cap) {
                    Copy-Item $p $dest -Force -ErrorAction Stop
                } else {
                    # Byte-level tail. Opened with ReadWrite sharing because
                    # Claude holds these open while it runs, and a plain read
                    # lock fails against a live log.
                    $fs = [System.IO.File]::Open($p, [System.IO.FileMode]::Open,
                                                 [System.IO.FileAccess]::Read,
                                                 [System.IO.FileShare]::ReadWrite)
                    try {
                        $null = $fs.Seek(-$cap, [System.IO.SeekOrigin]::End)
                        $buf  = New-Object byte[] $cap
                        $read = $fs.Read($buf, 0, $cap)
                        $out  = [System.IO.File]::Create($dest)
                        try {
                            $hdr = [System.Text.Encoding]::ASCII.GetBytes(
                                "=== TRUNCATED: last $([math]::Round($cap/1KB)) KB of a $([math]::Round($len/1MB,1)) MB file ===`r`n")
                            $out.Write($hdr, 0, $hdr.Length)
                            $out.Write($buf, 0, $read)
                        } finally { $out.Dispose() }
                    } finally { $fs.Dispose() }
                    $trimmed += $w
                }
                $copied += $w
            } catch { $missing += $w }
        }
        Log "  Copied  : $($copied.Count) of $($wanted.Count) log files" -Colour Green
        if ($trimmed.Count -gt 0) {
            Log "  Trimmed : $($trimmed.Count) large log(s), $([math]::Round($shareTop/1KB)) KB for Cowork logs, $([math]::Round($shareRest/1KB)) KB for the rest" -Colour DarkGray
        }
        if ($missing.Count -gt 0) {
            Log "  Skipped : $($missing -join ', ')" -Colour DarkGray
        }
        # Absence is evidence, and now that the folder is enumerated rather than
        # guessed at, an absent VM log means it really is not there rather than
        # meaning this build spells it differently.
        if (-not ($present.Keys -contains 'cowork_vm_node.log')) {
            Log "  Note    : no VM log in the folder, so the Cowork VM may never have started here" -Colour Yellow
        }

        # The Service Control Manager entries for the three services that have
        # to be alive. This is where a service that refuses to start is
        # actually recorded.
        try {
            $svcNames = @($ServiceName, 'vmcompute', 'hns') | Where-Object { $_ }
            $pattern  = ($svcNames | ForEach-Object { [regex]::Escape($_) }) -join '|'
            # None in the window is no error: Get-FixWinEvent. Until 6.3.0
            # it was logged as a System log that could not be read.
            $evts = @(Get-FixWinEvent -Filter @{
                          LogName      = 'System'
                          ProviderName = 'Service Control Manager'
                          # The UTC time as round-trip text: see Get-FixWinEvent.
                          StartTime    = [datetime]::UtcNow.AddDays(-7).ToString('o')
                      } -MaxEvents 500 |
                      Where-Object { $_.Message -match $pattern })
            if ($evts.Count -gt 0) {
                $evts | Select-Object TimeCreated, Id, LevelDisplayName, Message |
                    Format-List | Out-File (Join-Path $bundleDir 'service-events.txt') -Encoding ascii
            } else {
                "No Service Control Manager events for $($svcNames -join ', ') in the last 7 days." |
                    Out-File (Join-Path $bundleDir 'service-events.txt') -Encoding ascii
            }
            Log "  Events  : $($evts.Count) service event(s) in the last 7 days" -Colour Green
        } catch {
            Log "  Events  : could not read the System log ($($_.Exception.Message))" -Colour DarkGray
        }

        # Feature states, by exact name. A *Hyper-V* search does not match
        # VirtualMachinePlatform, which is how a machine with every Hyper-V
        # feature enabled can still be missing the one Cowork needs.
        try {
            $featLines = @()
            foreach ($fn in @('VirtualMachinePlatform', 'Microsoft-Hyper-V-All',
                              'Microsoft-Hyper-V', 'HypervisorPlatform')) {
                $st = 'unknown (needs admin)'
                if ($script:IsAdmin) {
                    try {
                        $f = Get-WindowsOptionalFeature -Online -FeatureName $fn -ErrorAction Stop
                        $st = if ($f) { "$($f.State)" } else { 'not present on this edition' }
                    } catch { $st = "query failed: $($_.Exception.Message)" }
                }
                $featLines += ("{0,-26} {1}" -f $fn, $st)
            }
            $featLines | Out-File (Join-Path $bundleDir 'windows-features.txt') -Encoding ascii
            Log "  Features: recorded by exact name" -Colour Green
        } catch { $null = $_ }

        # Zip it. One attachment beats ten, and logs compress about ten to one,
        # so the folder was already inside the 4 MB budget and the archive lands
        # far below it. The folder stays behind for anyone who would rather read
        # it in place.
        try {
            $zipPath = "$bundleDir.zip"
            Compress-Archive -Path (Join-Path $bundleDir '*') -DestinationPath $zipPath `
                             -CompressionLevel Optimal -Force -ErrorAction Stop
            $zipLen = (Get-Item $zipPath -ErrorAction Stop).Length
            Log "  Archive : $([math]::Round($zipLen/1KB)) KB" -Colour Green
            $script:SupportZip = $zipPath
        } catch {
            Log "  Archive : could not zip ($($_.Exception.Message)); send the folder instead" -Colour DarkGray
            $script:SupportZip = $null
        }
    } catch {
        Log "  Could not build the bundle: $($_.Exception.Message)" -Colour Yellow
        $bundleDir = $null
    }

    Write-Host ""
    Log "Diagnostic complete, no changes were made" -Colour Green
    if ($bundleDir) {
        if ($script:SupportZip) {
            Log "Support bundle: $script:SupportZip" -Colour Cyan
            Log "Attach that one file to a bug report. It is under the 4 MB limit." -Colour DarkGray
        } else {
            Log "Support bundle: $bundleDir" -Colour Cyan
            Log "Attach that folder to a bug report; it has the logs and the event entries." -Colour DarkGray
        }
    }
    Save-Log

    if (-not $Quiet) {
        Write-Host ""
        try { [Win32Window]::Flash() } catch { $null = $_ }
        Wait-ForAnyKey
        try { [Win32Window]::StopFlash() } catch { $null = $_ }
    }
    if ($fixMutex) {
        try { $fixMutex.ReleaseMutex() } catch { $null = $_ }
        try { $fixMutex.Dispose() } catch { $null = $_ }
        $fixMutex = $null
    }
    exit 0
}

# ====================================================================
# SAFETY GATE -- Block when called by automation while user is active
# ====================================================================
# When -Quiet is set (called by the health monitor), Get-AutomatedFixBlocker
# says whether Claude is in use. Manual runs (no -Quiet) always proceed: the
# user explicitly wants a fix.
if ($Quiet) {
    $blocker = Get-AutomatedFixBlocker
    if ($blocker) {
        Log "BLOCKED: $blocker, skipping automated fix" -Colour Yellow
        Save-Log
        # Held back, not done: see $script:BlockedExitCode.
        exit $script:BlockedExitCode
    }
}

$vmReady = $false
# One flag shared by both Deep escalation blocks. Previously only the second
# block was gated, so a single Smart run could purge the cache twice, the
# second time part way through the re-download, which left the fix unable to
# converge. Both blocks now also require admin: without it the purge destroys
# the cache and then cannot restart the service.
$script:DeepEscalated = $false

# ====================================================================
# STEP 0 -- Pre-emptive HCS state cleanup (v4.8.0)
# ====================================================================
Log "[0/10] Checking for stale HCS compute systems..." -Colour Yellow
if ($script:IsAdmin) {
    try {
        # $null is no answer; '' is an empty list, which is clean. Until 6.3.0
        # an empty list read as "hcsdiag unavailable".
        $hcsList = Get-HcsList
        if ($null -eq $hcsList) {
            Log "hcsdiag unavailable, timed out or failed, skipping HCS cleanup" -Colour DarkGray -Indent
        } elseif ($hcsList -match "cowork-vm") {
            # Only STALE compute systems get killed here, and a compute system
            # is only stale if Claude is not running.
            #
            # This step runs before Step 1 kills Claude. Until v6.0.0 the kill
            # was a no-op, because the call used hcsdiag's non-existent "close"
            # verb, so the ordering never mattered. Now that it actually kills,
            # tearing down a live VM underneath a running Claude just prompts
            # Claude to build another one, which Step 1 then orphans.
            #
            # A VM belonging to a running Claude is dealt with by Step 5, after
            # the process is gone. Every process named claude counts, the CLI
            # included: this only
            # decides to leave the VM alone until Step 5.
            $claudeUp = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue).Count -gt 0
            if ($claudeUp) {
                Log "cowork-vm present but Claude is running; leaving it for Step 5" -Colour DarkGray -Indent
            } else {
                Log "Found stale cowork-vm in HCS, cleaning up" -Colour DarkYellow -Indent
                $cleaned = Close-StaleHcsVms -Action kill
                if ($cleaned -gt 0) {
                    Log "Killed $cleaned stale HCS compute system(s)" -Colour Green -Indent
                }
            }
        } else {
            Log "HCS state clean, no stale cowork-vm found" -Colour Green -Indent
        }
    } catch {
        Log "HCS cleanup failed (non-critical): $($_.Exception.Message)" -Colour DarkGray -Indent
    }
} else {
    Log "Skipped (requires admin)" -Colour DarkGray -Indent
}
Start-Sleep -Seconds 1

# Session transcript cleanup. Opt-in only, via -PurgeSessions.
#
# This used to run on every invocation in every mode, deleting anything under
# local-agent-mode-sessions older than 7 days, while the .DESCRIPTION promised
# that conversations were never touched. There was also a second, identical
# pass inside the cache purge, so the work was done twice.
#
# The Watch threshold this was written to keep under has itself been removed:
# a session file count has no bearing on HCS or VM health.
if ($PurgeSessions) {
    $sessionDir = $script:ClaudeEnv.SessionsDir
    if ($sessionDir -and (Test-Path $sessionDir)) {
        try {
            # UTC: local time jumps an hour when summer time starts or ends.
            $cutoff = [datetime]::UtcNow.AddDays(-7)
            $oldFiles = @(Get-ChildItem $sessionDir -Recurse -File -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTimeUtc -lt $cutoff })
            if ($oldFiles.Count -gt 0) {
                $sizeMB = [math]::Round(($oldFiles | Measure-Object -Property Length -Sum).Sum / 1MB, 1)
                if ($PSCmdlet.ShouldProcess("$($oldFiles.Count) session file(s)", "Delete ($sizeMB MB)")) {
                    $oldFiles | Remove-Item -Force -ErrorAction SilentlyContinue
                    Log "Deleted $($oldFiles.Count) session files older than 7 days ($sizeMB MB)" -Colour Green -Indent
                    # Remove the directories those files left empty
                    Get-ChildItem $sessionDir -Directory -ErrorAction SilentlyContinue |
                        Where-Object { @(Get-ChildItem $_.FullName -Recurse -File -ErrorAction SilentlyContinue).Count -eq 0 } |
                        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                }
            } else {
                Log "No session files older than 7 days" -Colour DarkGray -Indent
            }
        } catch {
            Log "Session cleanup failed (non-critical): $($_.Exception.Message)" -Colour DarkGray -Indent
        }
    }
}

# ====================================================================
# STEP 1 -- Kill the Claude Desktop app's processes
# ====================================================================
Log "[1/10] Terminating Claude processes..." -Colour Yellow

# The desktop app only. The Claude Code CLI runs as "claude" too: the copies the
# app starts for project threads and Code sessions, and any terminal session.
# This step used to kill all of them along with the app.
$claudeInventory = @(Get-ClaudeProcessInventory -Environment $script:ClaudeEnv)
$claudeProcs     = @($claudeInventory | Where-Object { $_.IsDesktop } | ForEach-Object { $_.Process })
$claudeSpared    = @($claudeInventory | Where-Object { -not $_.IsDesktop }).Count
if ($claudeSpared -gt 0) {
    Log "Leaving $claudeSpared other process(es) named claude alone (Claude Code CLI, or not identifiable as the app)" `
        -Colour DarkGray -Indent
}
if ($claudeProcs.Count -gt 0) {
    # Discovery already resolved this, preferring the package, which is the only
    # authoritative source for the DESKTOP app. Reading it off a live process is
    # a fallback, not the primary, because the Claude Code CLI also runs as a
    # process named "claude", out of claude-code\<version>\ under the app's data
    # folder or out of a terminal install. Latching onto one of those caches the
    # CLI path and turns every later relaunch into a CLI start.
    if ($script:CapturedClaudeExe) {
        Log "Exe path from discovery ($($script:ClaudeEnv.ClaudeExeSource)): $($script:CapturedClaudeExe)" `
            -Colour DarkGray -Indent
    } else {
        foreach ($cp in $claudeProcs) {
            try {
                $exeFromProc = $cp.MainModule.FileName
                if ($exeFromProc -and -not (Test-ClaudeCodeCliPath $exeFromProc) -and
                    (Test-Path $exeFromProc)) {
                    $script:CapturedClaudeExe = $exeFromProc
                    Log "Captured exe path: $exeFromProc" -Colour DarkGray -Indent
                    break
                }
            } catch { $null = $_ }
        }
    }
    if ($script:CapturedClaudeExe) {
        # Cache for future runs. WriteAllText rather than Out-File -Encoding utf8,
        # which emits a BOM on Windows PowerShell 5.1.
        try {
            [System.IO.File]::WriteAllText($ExePathCache, $script:CapturedClaudeExe, `
                (New-Object System.Text.UTF8Encoding($false)))
        } catch { $null = $_ }
    }
    # The gate at the start asked whether the app runs a Claude Code session;
    # step 0 can take a while, so it is asked again right before the kill.
    # Nothing that stops the app or the service has run yet, so a session
    # found here holds the run back. Step 0 may have removed a stale compute
    # system by now.
    Assert-NoNewAppSession -Before "Step 1" -ExitCode $script:BlockedExitCode
    # The confirming line sits INSIDE the gate. It used to sit outside, so a
    # -WhatIf run printed "Killed 16 Claude process(es)" having killed none.
    if ($PSCmdlet.ShouldProcess("$($claudeProcs.Count) Claude process(es)", "Stop")) {
        $claudeProcs | Stop-Process -Force -ErrorAction SilentlyContinue
        Log "Killed $($claudeProcs.Count) Claude process(es)" -Colour Green -Indent
    }
} else {
    Log "No Claude processes running" -Colour DarkGray -Indent
}
Start-Sleep -Seconds 1

# Past every check that holds a run back with exit 3: the menu, Diagnostic,
# Cancel, a No at "Proceed?", the -Quiet gate and the check right before
# Step 1's kill all come before this, and Step 1 has stopped the app if it
# was running. This run is the repair, and says so once, in the log and in
# the transcript, so that the monitor can tell it from a run that stopped
# before it (see $script:RepairUnderWayLine). Not in a -WhatIf run, which
# changes nothing.
if (-not $WhatIfPreference) { Log $script:RepairUnderWayLine -Colour DarkGray }

# ====================================================================
# STEP 2 -- Stop CoworkVMService
# ====================================================================
Log "[2/10] Stopping $ServiceName..." -Colour Yellow

$svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if (-not $svc) {
    Log "Service not found, is Cowork installed?" -Colour DarkGray -Indent
} elseif ($svc.Status -ne "Running") {
    Log "Service already stopped ($($svc.Status))" -Colour DarkGray -Indent
} else {
    if ($PSCmdlet.ShouldProcess($ServiceName, "Stop")) {
        # See $script:ServiceStoppedByRun: a -Quiet run that stops part-way
        # after this starts the service again.
        $script:ServiceStoppedByRun = $true
        $stopped = $false
        if ($script:IsAdmin) {
            try {
                $svc.Stop()
                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                while ($sw.Elapsed.TotalSeconds -lt $ServiceTimeout) {
                    Start-Sleep -Seconds 3
                    $curSvc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
                    if (-not $curSvc -or $curSvc.Status -eq "Stopped") {
                        $stopped = $true
                        break
                    }
                }
                $sw.Stop()
                if ($stopped) {
                    Log "Service stopped gracefully ($([math]::Round($sw.Elapsed.TotalSeconds))s)" -Colour Green -Indent
                }
            } catch { $null = $_ }
        }
        if (-not $stopped) {
            Log "Force-killing $ServiceExe (last resort after ${ServiceTimeout}s)..." -Colour DarkYellow -Indent
            try { Stop-Process -Name $ServiceExe -Force -ErrorAction Stop }
            catch {
                try { Start-Process "taskkill" -ArgumentList "/F /IM $ServiceExe.exe" -NoNewWindow -Wait }
                catch { $null = $_ }
            }
            # Confirm rather than assume. Without admin both attempts above fail,
            # and this used to report "Force-killed" in green either way.
            Start-Sleep -Seconds 1
            if (Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue) {
                Log "[!] $ServiceExe is still running; the force-kill did not take" -Colour Red -Indent
            } else {
                Log "Force-killed" -Colour Green -Indent
            }
        }
    }
}
Start-Sleep -Seconds 1

# ====================================================================
# STEP 3 -- HCS service recovery
# ====================================================================
Log "[3/10] Checking HCS service health..." -Colour Yellow

# Initialised before the try, not inside it. Test-RecentHcsErrors calls
# Test-CoworkServiceLog from outside that function's own try, so a throw there
# left $hcsDetected never assigned, and [bool]$hcsDetected in Step 9 then failed
# under StrictMode with "variable has not been set".
$hcsDetected = $null
try {
    $hcsDetected = Test-RecentHcsErrors
    if ($hcsDetected -eq "shutdown_stale") {
        Log "HCS shutdown failures (0xC037010D), stale state from property query bug" -Colour DarkYellow -Indent
        Log "vmcompute restart will clear this (same recovery as construct failure)" -Colour DarkGray -Indent
    }
    if ($hcsDetected -eq "construct_failure") {
        Log "HCS construct failure (0x800707DE), stale state from failed shutdowns" -Colour DarkYellow -Indent
    }
    if ($hcsDetected -eq "guest_connect_failure") {
        Log "Guest connection failure, isGuestConnected RPC timing out" -Colour DarkYellow -Indent
        Log "Service restart will clear stale guest state" -Colour DarkGray -Indent
    }
    if ($hcsDetected) {
        if ($script:IsAdmin) {
            Log "HCS errors detected, restarting vmcompute service" -Colour DarkYellow -Indent
            try {
                Stop-Service vmcompute -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 3
                Start-Service vmcompute -ErrorAction SilentlyContinue
                # Wait up to 15 seconds for Running status
                $vmcElapsed = 0
                $vmcRunning = $false
                while ($vmcElapsed -lt 15) {
                    Start-Sleep -Seconds 3
                    $vmcElapsed += 3
                    $vmcSvc = Get-Service -Name "vmcompute" -ErrorAction SilentlyContinue
                    if ($vmcSvc -and $vmcSvc.Status -eq "Running") { $vmcRunning = $true; break }
                }
                if ($vmcRunning) {
                    Log "vmcompute service restarted successfully" -Colour Green -Indent
                } else {
                    Log "vmcompute not running after 15s, escalating" -Colour DarkYellow -Indent
                    # Escalation 1: restart vmms (Virtual Machine Management)
                    $vmmsOk = $false
                    $vmmsSvc = Get-Service -Name "vmms" -ErrorAction SilentlyContinue
                    if ($vmmsSvc) {
                        Log "Restarting vmms (Virtual Machine Management)..." -Colour DarkYellow -Indent
                        Stop-Service vmms -Force -ErrorAction SilentlyContinue
                        Start-Sleep -Seconds 3
                        Start-Service vmms -ErrorAction SilentlyContinue
                        Log "vmms service restarted" -Colour Green -Indent
                        # Re-check vmcompute (vmms restart often brings it back)
                        Start-Sleep -Seconds 3
                        $vmcSvc2 = Get-Service -Name "vmcompute" -ErrorAction SilentlyContinue
                        if ($vmcSvc2 -and $vmcSvc2.Status -eq "Running") { $vmmsOk = $true }
                    } else {
                        Log "vmms service not found, skipping" -Colour DarkGray -Indent
                    }
                    # Escalation 2: restart HvHost -- ONLY in Deep mode (very disruptive)
                    if (-not $vmmsOk -and $script:SelectedMode -eq "Deep") {
                        $hvHostSvc = Get-Service -Name "HvHost" -ErrorAction SilentlyContinue
                        if ($hvHostSvc) {
                            Log "WARNING: Restarting HvHost affects ALL Hyper-V VMs" -Colour Red -Indent
                            Log "Restarting HvHost (Host Compute Service Host)..." -Colour DarkYellow -Indent
                            Restart-Service HvHost -Force -ErrorAction SilentlyContinue
                            Log "HvHost service restarted" -Colour Green -Indent
                        } else {
                            Log "HvHost service not found, skipping" -Colour DarkGray -Indent
                        }
                    }
                }
            } catch {
                Log "[!] vmcompute restart failed: $($_.Exception.Message)" -Colour Red -Indent
            }
        } else {
            Log "HCS errors detected but no admin, vmcompute restart requires elevation" -Colour DarkYellow -Indent
        }
    } else {
        Log "No HCS issues detected" -Colour DarkGray -Indent
    }
} catch {
    Log "[!] HCS check failed: $($_.Exception.Message), continuing" -Colour DarkGray -Indent
}

# ====================================================================
# STEP 4 -- Verify no orphan processes remain
# ====================================================================
Log "[4/10] Checking for orphan processes..." -Colour Yellow

# The app's own processes and the service only. A Claude Code CLI process is not
# an orphan of the app: it belongs to a project thread or a terminal session.
$orphanInventory = @(Get-ClaudeProcessInventory -Environment $script:ClaudeEnv)
$orphanSpared    = @($orphanInventory | Where-Object { -not $_.IsDesktop }).Count
$orphanAppCount  = @($orphanInventory | Where-Object { $_.IsDesktop }).Count
$remaining = @(@($orphanInventory | Where-Object { $_.IsDesktop } | ForEach-Object { $_.Process }) +
               @(Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue))
if ($orphanSpared -gt 0) {
    Log "Leaving $orphanSpared other process(es) named claude alone (Claude Code CLI, or not identifiable as the app)" `
        -Colour DarkGray -Indent
}
if ($remaining.Count -gt 0) {
    # A process of the app found here can be one the app started after Step
    # 1, running a Claude Code session that the kill below would end, so a
    # -Quiet run asks again first, as before Step 1. Steps 1 and 2 have run,
    # so a session found here ends the run as stopped part-way, not as held
    # back, after starting again the service Step 2 stopped. The service
    # alone is stopped without asking. Until 6.3.0 nothing was asked here.
    if ($orphanAppCount -gt 0) { Assert-NoNewAppSession -Before "Step 4" -ExitCode $script:StoppedPartWayExitCode }
    if ($PSCmdlet.ShouldProcess("$($remaining.Count) orphan process(es)", "Force-kill")) {
        foreach ($proc in $remaining) {
            try { $proc | Stop-Process -Force -ErrorAction Stop } catch { $null = $_ }
        }
        Start-Sleep -Seconds 1
        $stubborn = @(@(Get-ClaudeDesktopProcess -Environment $script:ClaudeEnv) +
                      @(Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue))
        if ($stubborn.Count -gt 0) {
            Log "[!] $($stubborn.Count) process(es) refuse to die, a reboot may be needed" -Colour Red -Indent
        } else {
            Log "Cleaned up $($remaining.Count) orphan(s)" -Colour Green -Indent
        }
    }
} else {
    Log "All clear" -Colour Green -Indent
}

# ====================================================================
# STEP 5 -- Kill orphan HCS compute systems
# ====================================================================
Log "[5/10] Checking for orphan compute systems..." -Colour Yellow
try {
    $orphanKilled = $false
    # Something found and not confirmed gone: a compute system listed that no
    # kill went out for, whether none or only some of the kills were sent; a
    # kill hcsdiag did not confirm; or a worker still running. "No orphan
    # compute systems found" is not printed then. Until 6.3.0 it was.
    $orphanLeft = $false
    # Why the compute systems were not checked, or $null once hcsdiag's list
    # answered; and whether the worker check below finished. "No orphan
    # compute systems found" is printed only when both checks ran. Until
    # 6.3.0 it was printed in green when this run had no admin, when hcsdiag
    # did not answer, and when either check failed.
    $hcsNotChecked = 'needs admin'
    $workerChecked = $false
    # Method 1: hcsdiag (most reliable for HCS compute systems)
    if ($script:IsAdmin) {
        $hcsNotChecked = 'hcsdiag did not answer'
        try {
            # Close-StaleHcsVms owns the hcsdiag list parsing and the kill,
            # including the ShouldProcess gate. This used to carry a second,
            # independent parser that expected the GUID on a line of its own,
            # which hcsdiag never emits.
            $killedCount = Close-StaleHcsVms -Action kill
            if ($script:LastHcsListOk) { $hcsNotChecked = $null }
            if ($killedCount -gt 0) {
                Log "Killed $killedCount orphan compute system(s)" -Colour Green -Indent
                $orphanKilled = $true
            } elseif ($script:LastHcsKillsSent -gt 0) {
                # The lines above say which were still listed, or that the
                # check after the kill did not answer.
                Log "No orphan compute system confirmed gone" -Colour DarkYellow -Indent
            } elseif ($script:LastHcsListed -gt 0) {
                # Listed, and no kill went out: -WhatIf, or every kill threw.
                # Until 6.3.0 this ended in "No orphan compute systems found".
                Log "$($script:LastHcsListed) orphan compute system(s) listed, none killed" -Colour DarkYellow -Indent
                $orphanLeft = $true
            } elseif ($script:LastHcsListOk) {
                Log "No orphan compute systems via hcsdiag" -Colour DarkGray -Indent
            } else {
                # Step 0 already says this properly when hcsdiag is missing or
                # hangs. This branch used to print "No orphan compute systems
                # via hcsdiag" regardless, so a run that opened with "hcsdiag
                # unavailable" then claimed a clean hcsdiag result five steps
                # later. Seen in a user transcript on 2026-07-29.
                Log "hcsdiag did not answer, so this check was not run" -Colour DarkGray -Indent
            }
            # Kills went out for some of those listed and not for the rest:
            # those are still there, and are left. Until 6.3.0 nothing said so,
            # and they were not counted as left.
            $notSent = $script:LastHcsListed - $script:LastHcsKillsSent
            if ($script:LastHcsKillsSent -gt 0 -and $notSent -gt 0) {
                Log "$notSent orphan compute system(s) listed, not killed" -Colour DarkYellow -Indent
                $orphanLeft = $true
            }
            if ($script:LastHcsKillsSent -gt $killedCount) { $orphanLeft = $true }
        } catch {
            $hcsNotChecked = 'the hcsdiag query failed'
            Log "hcsdiag query failed: $($_.Exception.Message)" -Colour DarkGray -Indent
        }
    }
    # Note: Get-VM does not see HCS compute systems (like cowork-vm).
    # All cleanup is handled by hcsdiag above. (v4.8.0)

    # Method 2: a hung Cowork VM worker (vmwp.exe), and only the Cowork VM's.
    # Stop-HungCoworkWorker says how it tells which worker that is.
    try {
        $workerResult = Stop-HungCoworkWorker
        if ($workerResult.Stopped -gt 0) { $orphanKilled = $true }
        if ($workerResult.Left -gt 0) { $orphanLeft = $true }
        $workerChecked = $true
    } catch {
        Log "vmwp.exe check failed (non-critical): $($_.Exception.Message)" -Colour DarkGray -Indent
    }

    if (-not $orphanKilled -and -not $orphanLeft) {
        if ($null -eq $hcsNotChecked -and $workerChecked) {
            Log "No orphan compute systems found" -Colour Green -Indent
        } else {
            $notChecked = @()
            if ($null -ne $hcsNotChecked) { $notChecked += "compute systems were not checked ($hcsNotChecked)" }
            if (-not $workerChecked) { $notChecked += 'the VM worker check did not finish' }
            Log "Nothing was stopped, but $($notChecked -join ', and ')" -Colour DarkYellow -Indent
        }
    }
} catch {
    Log "Orphan VM check failed (non-critical): $($_.Exception.Message)" -Colour DarkGray -Indent
}
Start-Sleep -Seconds 1

# ====================================================================
# STEP 6 -- Purge VM cache (skipped with -KeepCache or Quick mode)
# ====================================================================
# Quick mode and Smart mode (before escalation) skip cache purge.
# Deep mode and Smart-escalated do full purge with VHDX backup/restore.
$skipCachePurge = $KeepCache -or ($script:SelectedMode -in "Quick","Smart")
$vhdxBackedUp = @{}

# -- VHDX integrity check helper --
function Test-VhdxHeader {
    param([string]$Path)
    $fs = $null
    try {
        $fs = [System.IO.File]::OpenRead($Path)
        $buf = New-Object byte[] 4
        $fs.Seek(65536, 'Begin') | Out-Null
        $read = $fs.Read($buf, 0, 4)
        return ($read -eq 4 -and $buf[0] -eq 0x68 -and $buf[1] -eq 0x65 -and
                $buf[2] -eq 0x61 -and $buf[3] -eq 0x64)
    } catch {
        return $false
    } finally {
        # Closed in a finally, not on the success path only. A failed read used
        # to leave a handle open on the .tmp file, and the Remove-Item that was
        # meant to clean it up then failed silently.
        if ($fs) { try { $fs.Close() } catch { $null = $_ } }
    }
}

# The VHDX files worth carrying across a purge. rootfs.vhdx is deliberately
# absent: a Deep purge exists to rebuild the VM image, so restoring the image
# would put back the very corruption the purge was run to clear.
$script:PreservedVhdx = @('sessiondata.vhdx', 'smol-bin.vhdx')

function Get-CoworkVhdxSource {
    param([Parameter(Mandatory)][string]$Name)
    # Prefer the inventory discovery already built at startup. Fall back to a
    # live scan when the tree has changed since then, which is exactly what an
    # escalation purge followed by a re-download does.
    try {
        foreach ($v in @($script:ClaudeEnv.VhdxFiles)) {
            if ($v.Name -eq $Name -and (Test-Path $v.Path)) {
                return (Get-Item $v.Path -ErrorAction Stop)
            }
        }
    } catch { $null = $_ }
    foreach ($cd in @($VmCachePath, $BundlePath)) {
        if (-not $cd -or -not (Test-Path $cd)) { continue }
        try {
            $hit = Get-ChildItem $cd -Recurse -Filter $Name -ErrorAction SilentlyContinue |
                   Select-Object -First 1
            if ($hit) { return $hit }
        } catch { $null = $_ }
    }
    return $null
}

function Backup-CoworkVhdx {
    <#
    .SYNOPSIS
        Backs up the preserved VHDX files ahead of a cache purge.
    .DESCRIPTION
        Returns a hashtable with three keys:
          BackedUp    name -> the ORIGINAL full path it came from, so the
                      restore can put it back exactly where it was instead of
                      reconstructing a directory depth by hand
          SourceFound $true when at least one preserved VHDX existed
          SafeToPurge $false means the caller must NOT delete the cache

        SafeToPurge is the point of this function. The old code checked free
        space against a 720MB literal while sessiondata.vhdx alone is over
        3 GB, so the check passed, the copy ran out of room part way through,
        the catch deleted the partial, and the purge then destroyed the
        original anyway.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$BackupDir
    )
    $r = @{ BackedUp = @{}; SourceFound = $false; SafeToPurge = $false }

    try {
        if (-not (Test-Path $BackupDir)) {
            New-Item $BackupDir -ItemType Directory -Force | Out-Null
        }
    } catch {
        Log "[!] Could not create the backup folder: $($_.Exception.Message)" -Colour Red -Indent
        return $r
    }

    $sources = @{}
    foreach ($name in $script:PreservedVhdx) {
        $src = Get-CoworkVhdxSource -Name $name
        if ($src) { $sources[$name] = $src }
    }
    if ($sources.Count -eq 0) {
        # Nothing to preserve means nothing to lose.
        $r.SafeToPurge = $true
        return $r
    }
    $r.SourceFound = $true

    # Requirement computed from the real file sizes, plus headroom.
    $needed = [long]0
    foreach ($k in @($sources.Keys)) { $needed += [long]$sources[$k].Length }
    $needed = [long]($needed * 1.1) + 64MB

    # Drive taken from the PATH, not from Get-Item on the directory.
    #
    # The backup directory may not exist yet: New-Item above is
    # ShouldProcess-gated, so a dry run does not create it, and Get-Item on a
    # missing path throws. That left free space "unknown" and refused the purge
    # on every -WhatIf run, so a dry run reported the opposite of what a real
    # run would do. Split-Path -Qualifier works on a path that is not there.
    $free = [long](-1)
    try {
        $qual = Split-Path $BackupDir -Qualifier -ErrorAction Stop
        if ($qual) {
            $psd = Get-PSDrive $qual.TrimEnd(':') -ErrorAction Stop
            if ($psd -and $null -ne $psd.Free) { $free = [long]$psd.Free }
        }
    } catch { $null = $_ }

    if ($free -lt 0) {
        # A redirected or UNC %APPDATA% can report no free space at all.
        # Unknown counts as do-not-purge: destroying a 3 GB VHDX on the strength
        # of a number that could not be read is not a trade worth making.
        Log "Free space on the backup volume is unknown; not purging" -Colour DarkYellow -Indent
        return $r
    }
    if ($free -lt $needed) {
        Log ("Not enough space for the VHDX backup ({0} MB free, {1} MB needed); not purging" -f `
            [math]::Round($free/1MB,0), [math]::Round($needed/1MB,0)) -Colour DarkYellow -Indent
        return $r
    }

    foreach ($name in @($sources.Keys)) {
        $src   = $sources[$name]
        $tmp   = Join-Path $BackupDir "$name.tmp"
        $final = Join-Path $BackupDir $name
        try {
            if (-not $PSCmdlet.ShouldProcess($name, "Backup")) {
                # Dry run. Record what WOULD have been backed up, so the caller
                # goes on to show the purge it would then have performed.
                #
                # Without this the backup copies nothing, SafeToPurge comes back
                # false, and -WhatIf reports "purge skipped" for a Deep run that
                # in reality would purge. A dry run that describes the opposite
                # of the real behaviour is worse than no dry run.
                $r.BackedUp[$name] = $src.FullName
                continue
            }

            Copy-Item $src.FullName $tmp -Force -ErrorAction Stop
            if (Test-VhdxHeader $tmp) {
                if (Test-Path $final) { Remove-Item $final -Force -ErrorAction SilentlyContinue }
                Rename-Item $tmp $name -ErrorAction Stop
                $r.BackedUp[$name] = $src.FullName
                Log ("Backed up {0} ({1} MB)" -f $name, `
                    [math]::Round($src.Length/1MB,1)) -Colour Green -Indent
            } else {
                Log "WARNING: $name failed its VHDX header check; discarding the copy" -Colour DarkYellow -Indent
                Remove-Item $tmp -Force -ErrorAction SilentlyContinue
            }
        } catch {
            Log "[!] Failed to back up $name : $($_.Exception.Message)" -Colour Red -Indent
            if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
        }
    }

    # Session data is the only genuinely unrecoverable file in the set. If it
    # was there and did not make it into the backup, the purge does not run.
    if ($sources.ContainsKey('sessiondata.vhdx') -and
        -not $r.BackedUp.ContainsKey('sessiondata.vhdx')) {
        Log "sessiondata.vhdx was not backed up; not purging, so it is not lost" -Colour DarkYellow -Indent
        return $r
    }
    $r.SafeToPurge = $true
    return $r
}

function Restore-CoworkVhdx {
    <#
    .SYNOPSIS
        Puts backed-up VHDX files back where they came from.
    .DESCRIPTION
        Restores to the recorded original path. The escalation restore this
        replaces rebuilt the path by hand, one directory level too shallow, and
        walked the cache directories in the opposite order to the backup, so it
        wrote to the wrong place on the occasions it wrote at all.

        Must run BEFORE the service is started. Once the service is up it
        recreates the tree itself, the "target does not exist" guard goes false,
        and the backup is abandoned without a word.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$BackupDir,
        [Parameter(Mandatory)]$BackedUp
    )
    $restored = 0
    foreach ($name in @($BackedUp.Keys)) {
        $backup = Join-Path $BackupDir $name
        $target = $BackedUp[$name]
        if (-not $target) { continue }
        if (-not (Test-Path $backup)) { continue }
        if (Test-Path $target) { continue }
        try {
            $parent = Split-Path $target -Parent
            if ($parent -and -not (Test-Path $parent)) {
                New-Item $parent -ItemType Directory -Force | Out-Null
            }
            if ($PSCmdlet.ShouldProcess($name, "Restore")) {
                Copy-Item $backup $target -Force -ErrorAction Stop
                Log "Restored $name" -Colour Green -Indent
                $restored++
            }
        } catch {
            Log "[!] Could not restore $name : $($_.Exception.Message); the service will recreate it" `
                -Colour DarkYellow -Indent
        }
    }
    return $restored
}

if ($skipCachePurge) {
    $skipReason = if ($KeepCache) { "-KeepCache" } else { "$($script:SelectedMode) mode" }
    Log "[6/10] Keeping VM cache ($skipReason)" -Colour DarkGray
    foreach ($item in @(
        @{ Path = $VmCachePath; Label = "claude-code-vm" },
        @{ Path = $BundlePath;  Label = "vm_bundles" }
    )) {
        if (Test-Path $item.Path) {
            $size = (Get-ChildItem $item.Path -Recurse -ErrorAction SilentlyContinue |
                     Measure-Object -Property Length -Sum).Sum
            $sizeMB = [math]::Round($size / 1MB, 1)
            Log "$($item.Label) preserved ($sizeMB MB)" -Colour DarkGray -Indent
        }
    }
} else {
    Log "[6/10] Smart cache purge..." -Colour Yellow

    # Phase 0: HCS state cleanup before cache purge (v4.8.0)
    if ($script:IsAdmin) {
        try {
            $hcsList = Get-HcsList
            if ($hcsList -and $hcsList -match "cowork-vm") {
                $cleaned = Close-StaleHcsVms -Action kill
                if ($cleaned -gt 0) {
                    Log "HCS: closed $cleaned stale cowork-vm(s)" -Colour Green -Indent
                }
            }
        } catch {
            Log "HCS deep cleanup failed (non-critical): $($_.Exception.Message)" -Colour DarkGray -Indent
        }
    }

    # Phase 0b used to be a second, identical session-log purge. It is gone:
    # session cleanup now happens once, near the top, and only when the caller
    # asked for it with -PurgeSessions.

    # Phase 1: let the service release its handles, then back up.
    $backupDir = Join-Path $LogDir "vhdx-backup"

    $null = Wait-ForServiceProcessExit

    $backupResult = Backup-CoworkVhdx -BackupDir $backupDir
    $vhdxBackedUp = $backupResult.BackedUp

    # Phase 2: purge, but only once the backup has confirmed it is safe to.
    # This used to run unconditionally, so a backup that failed for any reason
    # was followed immediately by deletion of the file it had failed to copy.
    if (-not $backupResult.SafeToPurge) {
        Log "Cache purge skipped: the session data is not safely backed up" -Colour DarkYellow -Indent
    } else {
        foreach ($item in @(
            @{ Path = $VmCachePath; Label = "claude-code-vm" },
            @{ Path = $BundlePath;  Label = "vm_bundles" }
        )) {
            if (Test-Path $item.Path) {
                $size = (Get-ChildItem $item.Path -Recurse -File -ErrorAction SilentlyContinue |
                         Measure-Object -Property Length -Sum).Sum
                if (-not $size) { $size = 0 }
                $sizeMB = [math]::Round($size / 1MB, 1)
                # The confirming line sits inside the gate now, so a -WhatIf run
                # no longer reports a deletion that did not happen.
                if ($PSCmdlet.ShouldProcess($item.Label, "Delete ($sizeMB MB)")) {
                    Remove-Item $item.Path -Recurse -Force -ErrorAction SilentlyContinue
                    Log "$($item.Label) removed ($sizeMB MB freed)" -Colour Green -Indent
                }
            } else {
                Log "$($item.Label) not present" -Colour DarkGray -Indent
            }
        }
    }

    # Phase 3 (restore) runs in Step 7 below, BEFORE the service is started.
    #
    # The MSIX smol-bin fallback that used to sit here has been removed. It
    # looked for resources\app\claudevm.bundle\smol-bin.vhdx, but the real
    # package layout is app\resources\, and no smol-bin.vhdx ships inside the
    # package on any current build, so that path never resolved.
}

# -- Temp file cleanup (Change 5) --
try {
    $tempPatterns = @("$env:TEMP\anthropic-*", "$env:TEMP\claude-*")
    $tempCleaned = 0
    $tempBytes   = [long]0
    foreach ($pattern in $tempPatterns) {
        foreach ($item in @(Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue)) {
            # $_.Length throws on a DirectoryInfo under StrictMode, and the
            # empty catch wrapped around this whole block hid it, so the temp
            # cleanup never actually ran. Directories are now sized by summing
            # the files inside them.
            $itemBytes = [long]0
            try {
                if ($item.PSIsContainer) {
                    $sum = (Get-ChildItem $item.FullName -Recurse -File -ErrorAction SilentlyContinue |
                            Measure-Object -Property Length -Sum).Sum
                    if ($sum) { $itemBytes = [long]$sum }
                } else {
                    $itemBytes = [long]$item.Length
                }
            } catch { $null = $_ }

            # Counter and byte total moved inside the gate, so -WhatIf no longer
            # reports having cleaned something it left alone.
            if ($PSCmdlet.ShouldProcess($item.FullName, "Remove temp item")) {
                Remove-Item $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
                $tempBytes += $itemBytes
                $tempCleaned++
            }
        }
    }
    if ($tempCleaned -gt 0) {
        $freed = if ($tempBytes -gt 1MB) { "{0:N1} MB" -f ($tempBytes/1MB) } else { "{0:N0} KB" -f ($tempBytes/1KB) }
        Log "Cleaned $tempCleaned temp items ($freed freed)" -Colour DarkGray -Indent
    }
} catch { $null = $_ }

# -- AnthropicClaude traditional-install path cleanup (Change 6) --
try {
    $tradPaths = @(
        (Join-Path $env:LOCALAPPDATA 'AnthropicClaude\sessions'),
        (Join-Path $env:LOCALAPPDATA 'AnthropicClaude\vm-state')
    )
    foreach ($tp in $tradPaths) {
        if (Test-Path $tp) {
            $count = (Get-ChildItem $tp -Recurse -ErrorAction SilentlyContinue).Count
            if ($count -gt 0) {
                if ($PSCmdlet.ShouldProcess($tp, "Clean traditional install path")) {
                    Remove-Item "$tp\*" -Recurse -Force -ErrorAction SilentlyContinue
                    Log "Cleaned traditional install path: $tp ($count items)" -Colour DarkGray -Indent
                }
            }
        }
    }
} catch { $null = $_ }

# ====================================================================
# STEP 7 -- Restart CoworkVMService (with extended polling)
# ====================================================================
Log "[7/10] Starting $ServiceName..." -Colour Yellow

# Phase 3 restore, and it has to happen BEFORE the service starts. Once the
# service is up it recreates the cache tree itself, which makes the "target does
# not exist" guard false and leaves the backup sitting unused. That is exactly
# what used to happen: this block ran after Step 7 had already started the
# service, so a 3 GB backup was abandoned on every Deep run, silently.
# $vhdxBackedUp is always a hashtable here, so the old "-and $vhdxBackedUp"
# term was checking truthiness of an object that is truthy even when empty.
# The count is what actually matters.
if (-not $skipCachePurge -and $vhdxBackedUp.Count -gt 0) {
    $restoredCount = Restore-CoworkVhdx -BackupDir $backupDir -BackedUp $vhdxBackedUp
    if ($restoredCount -gt 0) {
        Log "Restored $restoredCount VHDX file(s) before starting the service" -Colour Green -Indent
    }
}

if ($PSCmdlet.ShouldProcess($ServiceName, "Start")) {
    if ($script:IsAdmin) {
        $svcOk = Restart-CoworkService
        if ($svcOk) {
            Log "Service running" -Colour Green -Indent
        } else {
            Log "[!] Service failed to start; will retry after launch" -Colour Yellow -Indent
        }
    } else {
        Log "Skipping manual service start (no admin)" -Colour DarkGray -Indent
        Log "Claude will restart the service automatically when it launches" -Colour DarkGray -Indent
    }
}

# -- Smart mode escalation: if service didn't start, escalate to Deep --
if ($script:SelectedMode -eq "Smart") {
    $smartSvc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    $prereqBlocked = ($null -ne $script:ServicePrereq -and $script:ServicePrereq.Blocked)
    if ($prereqBlocked) {
        Log "Smart mode: not escalating to Deep, $($script:ServicePrereq.Reason)" -Colour Yellow
        Log "Purging the VM cache cannot fix that. Resolve it and rerun." -Colour DarkGray -Indent
    }
    $smartSvcDown = (-not $smartSvc -or $smartSvc.Status -ne "Running")
    if (-not $prereqBlocked -and $smartSvcDown -and -not $script:IsAdmin) {
        Log "Smart mode: service not running, but a Deep purge needs admin. Rerun elevated." -Colour DarkYellow
    }
    if (-not $prereqBlocked -and $smartSvcDown -and $script:IsAdmin -and -not $script:DeepEscalated) {
        $script:DeepEscalated = $true
        Log "Smart mode: service not running after quick fix, escalating to Deep" -Colour Yellow
        # Run the cache purge that was skipped -- with VHDX backup
        Log "[7/10] Escalated cache purge..." -Colour Yellow
        # Phase 0: HCS cleanup
        if ($script:IsAdmin) {
            try { Close-StaleHcsVms -Action kill | Out-Null } catch { $null = $_ }
        }
        # Phase 1: back up, through the same implementation the main path uses.
        # This used to be its own unvalidated copy that skipped the header check
        # and the tmp-plus-rename, writing over the good backup from Phase 1
        # with an unverified one.
        $escalateBackupDir = Join-Path $LogDir "vhdx-backup"

        $null = Wait-ForServiceProcessExit

        $escalateBackup = Backup-CoworkVhdx -BackupDir $escalateBackupDir

        # Phase 2: purge, gated on the backup exactly as the main path is.
        if (-not $escalateBackup.SafeToPurge) {
            Log "Escalated purge skipped: the session data is not safely backed up" -Colour DarkYellow -Indent
        } else {
            foreach ($item in @(
                @{ Path = $VmCachePath; Label = "claude-code-vm" },
                @{ Path = $BundlePath;  Label = "vm_bundles" }
            )) {
                if (Test-Path $item.Path) {
                    $size = (Get-ChildItem $item.Path -Recurse -File -ErrorAction SilentlyContinue |
                             Measure-Object -Property Length -Sum).Sum
                    if (-not $size) { $size = 0 }
                    $sizeMB = [math]::Round($size / 1MB, 1)
                    if ($PSCmdlet.ShouldProcess($item.Label, "Delete ($sizeMB MB)")) {
                        Remove-Item $item.Path -Recurse -Force -ErrorAction SilentlyContinue
                        Log "$($item.Label) removed ($sizeMB MB freed)" -Colour Green -Indent
                    }
                }
            }
        }

        # Phase 3: restore BEFORE the service comes back, then start it.
        $null = Restore-CoworkVhdx -BackupDir $escalateBackupDir -BackedUp $escalateBackup.BackedUp

        Log "Retrying service start after the deep purge..." -Colour Yellow -Indent
        if ($script:IsAdmin) {
            $svcOk2 = Restart-CoworkService
            if ($svcOk2) {
                Log "Service running after escalation" -Colour Green -Indent
            } else {
                Log "[!] Service still failed after the deep purge" -Colour Red -Indent
            }
        }
    }
}

# ====================================================================
# STEP 8 -- Relaunch Claude Desktop
# ====================================================================
if ($SkipLaunch) {
    Log "[8/10] Skipping Claude launch (-SkipLaunch)" -Colour DarkGray
    Log "[9/10] Skipping health check (-SkipLaunch)" -Colour DarkGray
} else {
    Log "[8/10] Launching Claude Desktop..." -Colour Yellow

    # Pre-launch guard: ensure HCS is clean before launching (v5.0.0)
    if ($script:IsAdmin) {
        try {
            $preLaunchCleaned = Close-StaleHcsVms -Action kill
            if ($preLaunchCleaned -gt 0) {
                Log "Pre-launch: cleaned $preLaunchCleaned stale HCS compute system(s)" -Colour Green -Indent
                Start-Sleep -Seconds 2
            }
        } catch { $null = $_ }
    }

    $claudeExe = Find-ClaudeExe

    # Detect MSIX install and use shell:AppsFolder protocol if so
    $launched = $false

    # Method 0: Elevated launch via scheduled task (created by Prevent-ClaudeIssues)
    # This gives Claude a full admin token without UAC prompt.
    # The task uses direct .exe launch (not shell:AppsFolder) so it inherits
    # the Highest RunLevel -- shell:AppsFolder would route through the
    # non-elevated desktop shell and defeat the purpose.
    $elevTaskPath = "\Claude\LaunchClaudeAdmin"
    try {
        $taskExists = Get-ScheduledTask -TaskName "LaunchClaudeAdmin" -TaskPath "\Claude\" -ErrorAction SilentlyContinue
        if ($taskExists) {
            if ($PSCmdlet.ShouldProcess($elevTaskPath, "Launch (elevated task)")) {
                Start-ScheduledTask -TaskName "LaunchClaudeAdmin" -TaskPath "\Claude\" -ErrorAction Stop
                # Verify the task actually started something (give it 3 seconds).
                # The app, not any "claude": a Claude Code CLI process that
                # survived Step 1 would otherwise read as a successful launch.
                Start-Sleep -Seconds 3
                $claudeProc = @(Get-ClaudeDesktopProcess -Environment $script:ClaudeEnv)
                if ($claudeProc.Count -gt 0) {
                    Log "Launched elevated via scheduled task: $elevTaskPath" -Colour Green -Indent
                    $launched = $true
                } else {
                    Log "Scheduled task ran but Claude process not detected, falling back" -Colour DarkYellow -Indent
                }
            }
        } else {
            Log "LaunchClaudeAdmin task not found, falling back to standard launch" -Colour DarkGray -Indent
            Log "(Run Prevent-ClaudeIssues.bat to enable elevated launch)" -Colour DarkGray -Indent
        }
    } catch {
        Log "Elevated launch failed: $($_.Exception.Message), falling back" -Colour DarkYellow -Indent
    }

    # Method A: MSIX.
    #
    # The package query, the manifest read and the Application Id lookup all
    # used to happen here, and the Get-AppxPackage call was the one optional
    # module call in the file with no guard around it. The Id also fell back to
    # a hardcoded "App", producing an invalid AUMID exactly when the manifest
    # read had failed, which is when a wrong guess actually costs something.
    # Discovery does all of it once at startup, inside try/catch, and records
    # anything it could not determine.
    if (-not $launched -and $script:ClaudeEnv.IsMsix -and $script:ClaudeEnv.Aumid) {
        Log "Detected MSIX install: $($script:ClaudeEnv.PackageFamilyName)" -Colour DarkGray -Indent
        $shellUri = $script:ClaudeEnv.Aumid
        if ($PSCmdlet.ShouldProcess($shellUri, "Launch (MSIX)")) {
            try {
                Start-Process $shellUri -ErrorAction Stop
                Log "Launched (MSIX): $shellUri" -Colour Green -Indent
                $launched = $true
            } catch {
                Log "[!] MSIX launch failed: $($_.Exception.Message)" -Colour Red -Indent
            }
        }
    }

    # Method B: Direct exe (non-MSIX installs only)
    if (-not $launched -and $claudeExe) {
        # Skip direct launch for WindowsApps paths -- it creates a loose instance
        # with a duplicate taskbar icon. Fall through to Method C (shortcut).
        if ($claudeExe -match "WindowsApps") {
            Log "Exe is in WindowsApps, skipping direct launch (would create loose instance)" -Colour DarkYellow -Indent
        } else {
            if ($PSCmdlet.ShouldProcess($claudeExe, "Launch")) {
                try {
                    Start-Process $claudeExe -ErrorAction Stop
                    Log "Launched: $claudeExe" -Colour Green -Indent
                    $launched = $true
                } catch {
                    Log "[!] Failed to launch: $($_.Exception.Message)" -Colour Red -Indent
                }
            }
        }
    }

    # Method C: Start Menu shortcut as last resort
    if (-not $launched) {
        Log "Trying Start Menu shortcut as last resort..." -Colour DarkYellow -Indent
        $lnkFile = Get-ChildItem "$env:APPDATA\Microsoft\Windows\Start Menu" -Recurse -Filter "Claude*.lnk" -ErrorAction SilentlyContinue |
                   Select-Object -First 1
        if (-not $lnkFile) {
            $lnkFile = Get-ChildItem "$env:ProgramData\Microsoft\Windows\Start Menu" -Recurse -Filter "Claude*.lnk" -ErrorAction SilentlyContinue |
                       Select-Object -First 1
        }
        if ($lnkFile) {
            try {
                Start-Process $lnkFile.FullName -ErrorAction Stop
                Log "Launched via shortcut: $($lnkFile.FullName)" -Colour Green -Indent
                $launched = $true
            } catch {
                Log "[!] Shortcut launch failed: $($_.Exception.Message)" -Colour Red -Indent
            }
        } else {
            Log "[!] No Claude shortcut found" -Colour Red -Indent
        }
    }

    if (-not $launched) {
        Log "[!] All launch methods exhausted" -Colour Red -Indent
        Log "Please launch Claude manually from the Start Menu" -Colour Yellow -Indent
    }

    # ====================================================================
    # STEP 9 -- Wait for Cowork workspace readiness (now Step 9/10)
    # ====================================================================
    # What each pass of the wait reads, in the order it reads them:
    #   D. The size of the VM cache. While it keeps growing the deadline
    #      moves out, so a slow download is not taken for a stuck VM.
    #   E. cowork-service.log: the guest connection state, from the
    #      isGuestConnected polling it logs. Connected, with a VM log marker,
    #      ends the wait. A timeout, 90 s or more in, starts a recovery; an
    #      error is only logged.
    #   A. Every VM log, from where each stood when the wait began:
    #      "Startup complete" or "[Keepalive]" in cowork_vm_node.log,
    #      "connected, waiting for commands" in coworkd.log.
    #   A2. The tail of any VM log written in the last 3 minutes, for a
    #      workspace that was already up when the wait began.
    #   B. Whether "hcsdiag list" shows a cowork-vm compute system, noted
    #      once. A compute system existing does not mean the workspace is
    #      ready, and it is not read as if it did.
    #   C. Step markers in the VM logs, such as "guest_vsock_connect
    #      completed", shown as progress.
    # Until 6.3.0 this list named Get-VM "claudevm" with its heartbeat, which
    # never lists an HCS compute system, and a file size fallback the wait
    # does not have.
    # The named pipe RPC is NOT usable (requires signed client executable).
    # ====================================================================
    Log "[9/10] Waiting for Cowork workspace..." -Colour Yellow

    $vmReady = $false

    # Every VM log in the app's log folder is read, each from where it stood
    # when the wait began. App 6.18355 writes a boot to two of them in the
    # same second: "[VM:start] Startup complete" and "[Keepalive] Starting"
    # to cowork_vm_node.log, "[coworkd] connected, waiting for commands" to
    # coworkd.log. This used to read only the log that was newest when the
    # script started, which on 6.18355 is usually coworkd.log, and none of the
    # markers below matched a boot there. A Smart run could then wait out its
    # deadline on a VM that had come up, and escalate to Deep, which deletes
    # the VM cache.
    #
    # The list before discovery led with C:\ProgramData\Claude\Logs\coworkd.log.
    # That tree does not exist on any current build, the drive letter was
    # hardcoded, and the coworkd.log in %APPDATA%\Claude\logs was last
    # written in March. 6.18355 writes a new one in %LOCALAPPDATA%\Claude\logs.
    $vmLogDir   = $script:ClaudeEnv.LogDir
    $vmCandList = @(Get-ClaudeVmLogCandidates -LogDir $vmLogDir)
    if ($vmCandList.Count -gt 0) {
        foreach ($cand in $vmCandList) {
            Log "Reading VM log: $($cand.Path) (last written $($cand.AgeHours)h ago)" -Colour DarkGray -Indent
        }
    } else {
        Log "No VM log file found; readiness will rely on guest and HCS signals" -Colour DarkYellow -Indent
    }

    # Where each VM log stands now, by path. A log that is missing here and
    # appears later is read from its start. Never throws: the candidate list
    # is built that way.
    function Get-VmLogSizes {
        $sizes = @{}
        foreach ($cand in @(Get-ClaudeVmLogCandidates -LogDir $vmLogDir)) {
            $sizes[$cand.Path] = [long]$cand.Length
        }
        return $sizes
    }

    # Record where each log stands at the START so we only check new entries
    $logBaseline = Get-VmLogSizes

    # Helper: check the entries written to any VM log since the baseline for
    # boot completion markers
    function Test-VmLogReady {
        param([hashtable]$Baseline)
        $newContent = ''
        foreach ($cand in @(Get-ClaudeVmLogCandidates -LogDir $vmLogDir)) {
            try {
                $start = [long]0
                if ($Baseline -and $Baseline.ContainsKey($cand.Path)) {
                    $start = [long]$Baseline[$cand.Path]
                }
                # Shorter than at the baseline means the log was rotated, and
                # everything in the new file is new.
                if ([long]$cand.Length -lt $start) { $start = [long]0 }
                if ([long]$cand.Length -le $start) { continue }
                # Read only the new portion of the log
                $stream = $null
                $reader = $null
                try {
                    $stream = New-Object System.IO.FileStream(
                        $cand.Path, [System.IO.FileMode]::Open,
                        [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                    $stream.Position = $start
                    $reader = New-Object System.IO.StreamReader($stream)
                    $newContent += $reader.ReadToEnd() + "`n"
                } finally {
                    if ($reader) { try { $reader.Close() } catch { $null = $_ } }
                    if ($stream) { try { $stream.Close() } catch { $null = $_ } }
                }
            } catch { $null = $_ }
        }
        if (-not $newContent) { return $null }
        # Check for completion markers (most definitive first)
        # cowork_vm_node.log
        if ($newContent -match "Startup complete") { return "startup-complete" }
        # coworkd.log's line for the same boot, in the same second
        if ($newContent -match "\[coworkd\] connected, waiting for commands") { return "startup-complete" }
        if ($newContent -match "\[Keepalive\]") { return "keepalive" }
        if ($newContent -match "guest_vsock_connect completed") { return "vsock-connected" }
        if ($newContent -match "sdk_install completed") { return "sdk-installed" }
        # coworkd.log. The id is "oneshot-<uuid>" on 6.18355, which the
        # hex-only pattern this used to be never matched.
        if ($newContent -match "\[process:[^\]]+\] started PID") { return "process-started" }
        if ($newContent -match "\[coworkd\] mounted .+ at /sessions/") { return "mounts-ready" }
        if ($newContent -match "full egress mode enabled") { return "egress-ready" }
        return $null
    }

    # Helper: check HCS compute system state via hcsdiag.
    # Invoke-HcsDiag already runs the call in a timed-out background job, which
    # is what keeps this from hanging when vmcompute is unstable.
    function Test-HyperVReady {
        try {
            $hcsList = Get-HcsList
            if (-not $hcsList) { return $null }
            # The Integration Services heartbeat probe that used to sit here is
            # gone. Get-VM enumerates VMMS virtual machines, and the Cowork VM
            # is an HCS compute system that never appears there, a fact this
            # script already notes in Step 5. So the probe could not fire, and
            # neither could the "heartbeat" readiness path it fed.
            if ($hcsList -match "cowork-vm") { return "running" }
            return $null
        } catch { return $null }
    }

    # Wall-clock deadline that extends for as long as the VM cache is visibly
    # growing.
    #
    # The old $vmElapsed was not wall-clock at all. It advanced only by the 5s
    # sleep, while the service restart, the guest-timeout recovery and the
    # no-progress escalation added roughly 50s, 38s and 42s each without
    # touching it, so a "240 second" wait could run for six or seven minutes.
    #
    # A fixed ceiling also cannot tell slow apart from stuck. After a purge the
    # VM has 13.5 GB to pull before it can report ready, which on a slow link is
    # not a failure. So: start with a base allowance and push the deadline out
    # again every time the cache grows. No growth for the length of one progress
    # window is what stuck actually looks like.
    $vmSw            = [System.Diagnostics.Stopwatch]::StartNew()
    $vmProgressGrant = 180
    $vmHardCap       = 3600
    $vmDeadline      = 240
    if (-not $skipCachePurge) {
        # The cache was just emptied, so there is a full bundle to fetch before
        # anything can be reported. Start with more patience.
        $vmDeadline = 300
    }
    $lastCacheBytes = [long](-1)
    $vmElapsed      = 0
    $lastStatus = ""
    $hvChecked  = $false
    $svcRestartAttempts = 0
    $svcRestartLimit    = 3
    $script:NoProgressEscalated = $false
    # Whether the note that the VM may be stuck has been logged for the
    # current stretch with no log or guest activity.
    $hcsStuckNoted = $false

    # Skip Hyper-V cmdlet checks if vmcompute was just restarted (may hang on unstable service)
    $skipHvChecks = [bool]$hcsDetected
    if ($skipHvChecks) {
        Log "Skipping Hyper-V VM checks (vmcompute was just restarted)" -Colour DarkGray -Indent
    }

    # Skip the wait only when the service is BOTH flagged AND demonstrably not
    # running. v6.0.2 skipped on the flag alone, which was the same mistake as
    # the veto it was written to fix: it trusted a prediction over the machine.
    # On the reporter's box the service came up in 293ms the moment anything
    # actually asked it to, so a flag-only skip would have abandoned a workspace
    # that was about to work.
    #
    # Restart-CoworkService clears the flag when a start succeeds, so by the time
    # this runs, Blocked surviving alongside a stopped service means both the
    # prediction and the attempt agreed. That is worth acting on. Either one
    # alone is not.
    $svcAtWaitStart = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    $prereqStillBlocked = ($null -ne $script:ServicePrereq -and
                           $script:ServicePrereq.Blocked -and
                           (-not $svcAtWaitStart -or $svcAtWaitStart.Status -ne "Running"))
    if ($prereqStillBlocked) {
        Log "Not waiting for the workspace: the service is down and the cause needs fixing first." -Colour Red -Indent
        Log "Claude Desktop itself will run; Cowork will not until that is fixed." -Colour Yellow -Indent
    }

    while (-not $prereqStillBlocked -and
           $vmSw.Elapsed.TotalSeconds -lt $vmDeadline -and
           $vmSw.Elapsed.TotalSeconds -lt $vmHardCap) {
        Start-Sleep -Seconds 5
        $vmElapsed = [int]$vmSw.Elapsed.TotalSeconds

        # Measure the cache once per iteration and extend the deadline while it
        # is still filling. The status block below reuses this figure.
        $curCacheBytes = [long]0
        if ($VmCachePath -and (Test-Path $VmCachePath)) {
            try {
                $sum = (Get-ChildItem $VmCachePath -Recurse -File -ErrorAction SilentlyContinue |
                        Measure-Object -Property Length -Sum).Sum
                if ($sum) { $curCacheBytes = [long]$sum }
            } catch { $null = $_ }
        }
        if ($lastCacheBytes -ge 0 -and ($curCacheBytes - $lastCacheBytes) -gt 1MB) {
            $grantTo = $vmSw.Elapsed.TotalSeconds + $vmProgressGrant
            if ($grantTo -gt $vmDeadline) { $vmDeadline = $grantTo }
        }
        $lastCacheBytes = $curCacheBytes

        # Service health check.
        #
        # The restart result used to be piped to Out-Null. Restart-CoworkService
        # returns $false when Test-CoworkServicePrereq finds a blocker the user
        # has to clear, such as Virtual Machine Platform being disabled, and
        # throwing that away meant the loop retried something that could not
        # work. A user on 2026-07-29 watched the same four-line remediation
        # print every five seconds until they pressed Ctrl+C.
        #
        # Two bounds now. A blocked prereq stops immediately, because nothing
        # about it will change while the script runs. Anything else gets three
        # restarts, because a service dying repeatedly for an unknown reason is
        # not going to be fixed by a fourth.
        $svcNow = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if (-not $svcNow -or $svcNow.Status -ne "Running") {
            $svcRestartAttempts++
            if ($svcRestartAttempts -gt $svcRestartLimit) {
                Log "Service has died $svcRestartAttempts times; giving up on the wait" -Colour Red -Indent
                break
            }
            Log "Service died, restarting ($svcRestartAttempts of $svcRestartLimit)..." -Colour DarkYellow -Indent
            $restartOk = Restart-CoworkService
            if (-not $restartOk -and
                $null -ne $script:ServicePrereq -and $script:ServicePrereq.Blocked) {
                Log "This needs to be fixed on the machine first; not retrying." -Colour Red -Indent
                break
            }
            continue
        }

        # Re-enable Hyper-V checks after grace period (vmcompute stable by now)
        if ($skipHvChecks -and $vmElapsed -ge 20) {
            $skipHvChecks = $false
            $hvChecked = $false          # force re-probe on next iteration
            Log "Re-enabling Hyper-V VM checks (grace period elapsed)" -Colour DarkGray -Indent
        }

        # E. Check cowork-service.log for guest connection state
        $guestState = $null
        if ($vmElapsed -ge 30) {
            $guestState = Test-CoworkServiceLog -WindowSeconds 15 -Brief
            if ($guestState -eq "guest-connected") {
                # Guest reports connected but VM log hasn't caught up yet --
                # give it a few more seconds then accept
                Start-Sleep -Seconds 5
                $logStatus = Test-VmLogReady -Baseline $logBaseline
                if ($logStatus) {
                    $vmReady = $true
                    Log "Workspace ready (guest connected + log: $logStatus)" -Colour Green -Indent
                    break
                }
                # Guest says connected but no log markers yet -- keep waiting
            } elseif ($guestState -eq "guest-timeout" -and $vmElapsed -ge 90) {
                Log "[!] Guest connection timeout detected (isGuestConnected failing)" -Colour DarkYellow -Indent
                Log "Attempting targeted recovery..." -Colour Yellow -Indent
                # Targeted recovery: close HCS, restart cowork-svc
                try {
                    $cleaned = Close-StaleHcsVms -Action kill
                    if ($cleaned -gt 0) {
                        Log "Closed $cleaned stale HCS compute system(s)" -Colour Green -Indent
                    }
                } catch { $null = $_ }
                $null = Stop-CoworkServiceWithTimeout
                Start-Sleep -Seconds 3
                Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 5
                # Reset baseline so we check fresh log entries
                $logBaseline = Get-VmLogSizes
                Log "Service restarted, continuing to wait..." -Colour DarkGray -Indent
            } elseif ($guestState -eq "guest-error" -and $vmElapsed -ge 60) {
                Log "[!] Guest connection errors detected in cowork-service.log" -Colour DarkYellow -Indent
            }
        }

        # No-progress detector: if 60s in and nothing happening, escalate (v5.0.0)
        if ($vmElapsed -ge 60 -and -not $vmReady) {
            $hasLogActivity = ($null -ne (Test-VmLogReady -Baseline $logBaseline))
            $hasGuestActivity = ($guestState -and $guestState -ne "no-polling")
            $hasHcsVm = $false
            try {
                $hcsCheck = Get-HcsList
                $hasHcsVm = $hcsCheck -and ($hcsCheck -match "cowork-vm")
            } catch { $null = $_ }
            if (-not $hasLogActivity -and -not $hasGuestActivity) {
                # Once for each stretch with no activity, not on every pass.
                # Until 6.3.0 it was logged on every pass after 60 s whenever
                # hcsdiag listed the VM, so a slow boot that was making
                # progress filled the log with it.
                if ($hasHcsVm -and -not $hcsStuckNoted) {
                    Log "HCS VM present but no log/guest activity, VM may be stuck" -Colour DarkGray -Indent
                    $hcsStuckNoted = $true
                }
                if (-not $script:NoProgressEscalated) {
                    $script:NoProgressEscalated = $true
                    Log "[!] No progress after ${vmElapsed}s, no log activity, no guest polling" -Colour Red -Indent
                    Log "Escalating: killing Claude, cleaning HCS, relaunching..." -Colour Yellow -Indent
                    # Kill the app, not the Claude Code CLI processes beside it,
                    # and not while the relaunched app runs a session.
                    Assert-NoNewAppSession -Before "the no-progress escalation" -ExitCode $script:StoppedPartWayExitCode
                    $null = Stop-ClaudeDesktopApp
                    Start-Sleep -Seconds 2
                    # Clean HCS
                    try { Close-StaleHcsVms -Action kill | Out-Null } catch { $null = $_ }
                    Start-Sleep -Seconds 2
                    $null = Stop-CoworkServiceWithTimeout
                    Start-Sleep -Seconds 3
                    Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 5
                    $null = Start-ClaudeDesktop
                    # Reset log baseline and timers
                    $logBaseline = Get-VmLogSizes
                    # Don't reset vmElapsed -- let the outer timeout still apply
                    Log "Waiting for workspace after relaunch..." -Colour DarkGray -Indent
                    continue
                }
            } else {
                $hcsStuckNoted = $false
            }
        }

        # A. Check log file for boot completion (new entries since baseline)
        $logStatus = Test-VmLogReady -Baseline $logBaseline
        if ($logStatus -in @("startup-complete", "keepalive", "process-started", "egress-ready")) {
            $vmReady = $true
            Log "Workspace ready (log: $logStatus)" -Colour Green -Indent
            break
        }

        # A2. Fallback: workspace may already be running (baseline captured after boot).
        #     Check the tail of each VM log written in the last 3 minutes for
        #     boot markers, regardless of baseline.
        if (-not $logStatus -and $vmElapsed -ge 30) {
            foreach ($cand in @(Get-ClaudeVmLogCandidates -LogDir $vmLogDir)) {
                try {
                    $recentlyWritten = ($cand.LastWriteTimeUtc -gt [datetime]::UtcNow.AddMinutes(-3))
                    if (-not $recentlyWritten -or $cand.Length -le 0) { continue }
                    $tailSize = [math]::Min([long]$cand.Length, 8192)
                    $tailContent = ''
                    $stream = $null
                    $reader = $null
                    try {
                        $stream = New-Object System.IO.FileStream(
                            $cand.Path, [System.IO.FileMode]::Open,
                            [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                        $stream.Position = [long]$cand.Length - $tailSize
                        $reader = New-Object System.IO.StreamReader($stream)
                        $tailContent = $reader.ReadToEnd()
                    } finally {
                        if ($reader) { try { $reader.Close() } catch { $null = $_ } }
                        if ($stream) { try { $stream.Close() } catch { $null = $_ } }
                    }
                    if ($tailContent -match "Startup complete|\[Keepalive\]|connected, waiting for commands|started PID|full egress mode") {
                        $vmReady = $true
                        break
                    }
                } catch { $null = $_ }
            }
            if ($vmReady) {
                Log "Workspace ready (log tail: recently active with boot markers)" -Colour Green -Indent
                break
            }
        }

        # B. Note once whether an HCS compute system exists at all.
        #
        # This used to carry a second branch that accepted "running+heartbeat"
        # as evidence of readiness. Test-HyperVReady can never return that
        # value, so the branch never ran, and the "heartbeat(x3)" monitor the
        # startup banner advertises has never once fired.
        if (-not $hvChecked) {
            $hvChecked = $true
            $hvState = if ($skipHvChecks) { $null } else { Test-HyperVReady }
            if ($null -ne $hvState) {
                Log "HCS compute system present: $hvState" -Colour DarkGray -Indent
            }
        }

        # C. Progress from log step markers
        $curStatus = ""
        if ($logStatus -eq "sdk-installed") {
            $curStatus = "Finishing setup... (${vmElapsed}s)"
        } elseif ($logStatus -eq "vsock-connected") {
            $curStatus = "Installing SDK... (${vmElapsed}s)"
        } else {
            # Reuses the figure measured at the top of this iteration rather
            # than walking the cache tree a second time.
            $curSizeMB = [math]::Round($curCacheBytes / 1MB, 0)
            if ($curSizeMB -lt 50) {
                $curStatus = "Setting up workspace... (${vmElapsed}s, ${curSizeMB} MB)"
            } else {
                if ($guestState -eq "guest-polling") {
                    $curStatus = "Starting workspace... (${vmElapsed}s, guest polling)"
                } else {
                    $curStatus = "Starting workspace... (${vmElapsed}s)"
                }
            }
        }

        if ($curStatus -ne $lastStatus) {
            Log $curStatus -Colour DarkGray -Indent
            $lastStatus = $curStatus
        }
    }

    if (-not $vmReady) {
        # Final check: if log has vsock-connected or sdk-installed, partially ready
        $finalLog = Test-VmLogReady -Baseline $logBaseline
        if ($finalLog) {
            Log "Workspace partially ready (log: $finalLog), may still be loading" -Colour Yellow -Indent
        } else {
            Log "[!] Workspace not confirmed ready after ${vmElapsed}s" -Colour Yellow -Indent
            Log "Open a Cowork session in Claude to trigger setup" -Colour DarkGray -Indent
        }
    }
}

# -- Smart mode: escalate to Deep if workspace didn't come up (v5.0.0) --
if (-not $vmReady -and -not $SkipLaunch -and $script:IsAdmin -and
    $script:SelectedMode -eq "Smart" -and -not $script:DeepEscalated) {
    # -SkipLaunch never enters the Step 9 branch, so $vmLogDir and
    # $logBaseline below only exist when that branch ran. Without this
    # guard, -SkipLaunch purged 13.5 GB and then died under StrictMode with
    # the cache already gone.
    $script:DeepEscalated = $true
    Log "" -Colour White
    Log "Smart mode: workspace not ready after quick fix, escalating to Deep" -Colour Yellow
    # Phase 0: Kill the app (not the Claude Code CLI) + HCS cleanup, and not
    # while the relaunched app runs a session.
    Assert-NoNewAppSession -Before "the escalation to Deep" -ExitCode $script:StoppedPartWayExitCode
    $null = Stop-ClaudeDesktopApp
    Start-Sleep -Seconds 2
    try { Close-StaleHcsVms -Action kill | Out-Null } catch { $null = $_ }
    # Phase 1: stop the service
    $null = Stop-CoworkServiceWithTimeout
    Start-Sleep -Seconds 3
    # Phase 2: Cache purge with VHDX backup
    $escalateBackupDir = Join-Path $LogDir "vhdx-backup"

    $null = Wait-ForServiceProcessExit

    $deepBackup = Backup-CoworkVhdx -BackupDir $escalateBackupDir

    # This block previously deleted both cache trees with no ShouldProcess gate
    # at all, unlike its two siblings, so a -WhatIf run destroyed 13.5 GB.
    if (-not $deepBackup.SafeToPurge) {
        Log "Deep escalation purge skipped: the session data is not safely backed up" -Colour DarkYellow -Indent
    } else {
        foreach ($item in @(
            @{ Path = $VmCachePath; Label = "claude-code-vm" },
            @{ Path = $BundlePath;  Label = "vm_bundles" }
        )) {
            if (Test-Path $item.Path) {
                $size = (Get-ChildItem $item.Path -Recurse -File -ErrorAction SilentlyContinue |
                         Measure-Object -Property Length -Sum).Sum
                if (-not $size) { $size = 0 }
                $sizeMB = [math]::Round($size / 1MB, 1)
                if ($PSCmdlet.ShouldProcess($item.Label, "Delete ($sizeMB MB)")) {
                    Remove-Item $item.Path -Recurse -Force -ErrorAction SilentlyContinue
                    Log "$($item.Label) removed ($sizeMB MB freed)" -Colour Green -Indent
                }
            }
        }
    }

    # Restore first, then start. The old order started the service and only then
    # tried to restore, by which point the service had rebuilt the tree and the
    # restore quietly declined to overwrite it.
    $null = Restore-CoworkVhdx -BackupDir $escalateBackupDir -BackedUp $deepBackup.BackedUp
    Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 5
    # Relaunch Claude. This used to try the scheduled task and nothing else, so
    # when the task was absent it fell through to the 120s wait below having
    # launched nothing at all, and then reported a timeout.
    $relaunched = Start-ClaudeDesktop
    $escalateTimeout = 120
    $escalateElapsed = 0
    if ($relaunched) {
        Log "Waiting for workspace after deep escalation..." -Colour Yellow -Indent
        $logBaseline = Get-VmLogSizes
    } else {
        Log "Nothing was relaunched, so there is nothing to wait for" -Colour DarkYellow -Indent
    }
    while ($relaunched -and $escalateElapsed -lt $escalateTimeout) {
        Start-Sleep -Seconds 5
        $escalateElapsed += 5
        $logStatus = Test-VmLogReady -Baseline $logBaseline
        if ($logStatus -in @("startup-complete", "keepalive", "process-started", "egress-ready")) {
            $vmReady = $true
            Log "Workspace ready after deep escalation ($logStatus)" -Colour Green -Indent
            break
        }
        if ($escalateElapsed % 30 -eq 0) {
            Log "Still waiting... (${escalateElapsed}s)" -Colour DarkGray -Indent
        }
    }
}

# -- Post-fix retry loop (v4.8.0) --
# If workspace didn't come up after the full fix, retry service restart
if (-not $vmReady -and -not $SkipLaunch) {
    Log "" -Colour White
    Log "Workspace not ready, attempting retry cycle ($MaxRetries max)..." -Colour Yellow
    for ($retryNum = 1; $retryNum -le $MaxRetries; $retryNum++) {
        Log "Retry $retryNum/$MaxRetries, quick service restart..." -Colour Yellow -Indent
        if ($script:IsAdmin) {
            # Quick cycle: stop the service, clean HCS, start it again
            $null = Stop-CoworkServiceWithTimeout
            Start-Sleep -Seconds 2
            try {
                    $cleaned = Close-StaleHcsVms -Action kill
                    if ($cleaned -gt 0) {
                        Log "Cleaned $cleaned stale HCS state(s)" -Colour Green -Indent
                    }
            } catch { $null = $_ }
            Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 5

            # Reset the baseline after the restart. Without this the loop below
            # re-reads content written before this retry began and treats it as
            # evidence that the retry worked.
            $logBaseline = Get-VmLogSizes

            # Wait up to 60s for workspace
            $retryElapsed = 0
            while ($retryElapsed -lt 60) {
                Start-Sleep -Seconds 5
                $retryElapsed += 5
                $retryStatus = Test-VmLogReady -Baseline $logBaseline
                # The same four definitive markers the main loop requires. This
                # used to accept ANY non-null marker, so a partial signal such
                # as sdk-installed reported a failed fix as a success and
                # suppressed the manual-intervention warning at the end.
                if ($retryStatus -in @("startup-complete", "keepalive",
                                       "process-started", "egress-ready")) {
                    Log "Workspace ready on retry $retryNum ($retryStatus)" -Colour Green -Indent
                    $vmReady = $true
                    break
                }
            }
            if ($vmReady) { break }
            Log "Retry $retryNum failed" -Colour DarkYellow -Indent
        } else {
            Log "Retry requires admin, skipping" -Colour DarkGray -Indent
            break
        }
    }
    if (-not $vmReady) {
        Log "All $MaxRetries retries exhausted, manual intervention may be needed" -Colour Red
        $script:ExitCode = 1
    }
}

# -- Bring this window to the front (Claude may have taken focus) ----
try { [Win32Window]::BringToFront() } catch { $null = $_ }

# ====================================================================
# Summary
# ====================================================================
Write-Host ""
Write-Host "  +-------------------------------------------+" -ForegroundColor Green
Write-Host "  |           OPERATION COMPLETE              |" -ForegroundColor Green
Write-Host "  +-------------------------------------------+" -ForegroundColor Green
Write-Host ""

$fSvc   = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
$fProcs = @(Get-ClaudeDesktopProcess -Environment $script:ClaudeEnv)
$svcOk  = $fSvc -and $fSvc.Status -eq "Running"

$svcStatusText  = "Not found"
$svcStatusColor = "Yellow"
if ($fSvc) {
    $svcStatusText = "$($fSvc.Status)"
    if ($svcOk) { $svcStatusColor = "Green" }
}
Write-Host "  VM Service:       $svcStatusText" -ForegroundColor $svcStatusColor
Write-Host "  Claude processes: $($fProcs.Count) active (app only)" -ForegroundColor Cyan

# If the service can never start, say so here rather than leaving the reader
# to guess from an "sc start" error code.
if (-not $svcOk -and $null -ne $script:ServicePrereq -and $script:ServicePrereq.Blocked) {
    Write-Host ""
    Write-Host "  Why the service is not running:" -ForegroundColor Yellow
    Write-Host "    $($script:ServicePrereq.Reason)" -ForegroundColor Red
    if ($script:ServicePrereq.Fixes.Count -gt 0) {
        Write-Host ""
        Write-Host "  To fix it (elevated PowerShell):" -ForegroundColor Yellow
        foreach ($f in $script:ServicePrereq.Fixes) {
            Write-Host "    $f" -ForegroundColor White
        }
    }
    Write-Host ""
    Write-Host "  Reference: https://support.claude.com/en/articles/12622703-deploy-claude-desktop-for-windows" -ForegroundColor DarkGray
}

# Quick peek at recent event log errors
try {
    $evFilter = @{
        LogName      = "Application"
        ProviderName = "CoworkVMService"
        Level        = 2
        # The UTC time as round-trip text: see Get-FixWinEvent.
        StartTime    = [datetime]::UtcNow.AddHours(-1).ToString('o')
    }
    $events = @(Get-FixWinEvent -Filter $evFilter -MaxEvents 3)

    if ($events) {
        Write-Host ""
        Log "Recent service errors (last hour):" -Colour DarkYellow -Indent
        foreach ($ev in $events) {
            $evTime = "{0:HH:mm}" -f $ev.TimeCreated
            $evMsg  = ($ev.Message -split "`n")[0]
            Log "  [$evTime] $evMsg" -Colour DarkGray -Indent
        }
    }
} catch {
    Log "Recent service errors: the Application log could not be read ($($_.Exception.Message))" -Colour DarkGray -Indent
}

Write-Host ""
Log "Log saved to: $LogFile" -Colour DarkGray -Indent

} catch {
    Write-Host ""
    Write-Host "  +-------------------------------------------+" -ForegroundColor Red
    Write-Host "  |           UNEXPECTED ERROR                |" -ForegroundColor Red
    Write-Host "  +-------------------------------------------+" -ForegroundColor Red
    Write-Host ""
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor DarkGray
    Write-Host ""
    $script:ExitCode = 1
} finally {
    Save-Log
    try { Stop-Transcript -ErrorAction SilentlyContinue } catch { $null = $_ }
    if ($fixMutex) {
        try { $fixMutex.ReleaseMutex() } catch { $null = $_ }
        try { $fixMutex.Dispose() } catch { $null = $_ }
        $fixMutex = $null
    }
}

# -- Always pause unless -Quiet --------------------------------------
if (-not $Quiet) {
    Write-Host ""
    try { [Win32Window]::Flash() } catch { $null = $_ }
    Wait-ForAnyKey
    try { [Win32Window]::StopFlash() } catch { $null = $_ }
}

exit $script:ExitCode
