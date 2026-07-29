<#
.SYNOPSIS
    Claude Desktop / Cowork, Reset and Fix

.DESCRIPTION
    Kills all Claude processes, stops CoworkVMService, recovers from HCS
    (Host Compute Service) errors, performs orphan compute system cleanup,
    purges stale VM cache, restarts the service, and relaunches Claude
    Desktop with elevated privileges when available.

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
    Version : 6.0.2
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
    if ($BootPrep)         { $elevateArgs += " -BootPrep" }
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
        # hard failure. BootPrep exits 1, which made the inconsistency visible.
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

    $info = [ordered]@{
        SchemaVersion = 1
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

        AppDataDir  = $null
        LogDir      = $null
        VmCachePath = $null
        BundlePath  = $null
        SessionsDir = $null

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

    # -- Per-user data locations -------------------------------------------
    # %APPDATA% is the only location that exists on current builds. The old
    # %ProgramData%\Claude tree is gone, so it is not probed here at all.
    try {
        if ($env:APPDATA) {
            $info.AppDataDir  = Join-Path $env:APPDATA "Claude"
            $info.LogDir      = Join-Path $info.AppDataDir "logs"
            $info.VmCachePath = Join-Path $info.AppDataDir "claude-code-vm"
            $info.BundlePath  = Join-Path $info.AppDataDir "vm_bundles"
            $info.SessionsDir = Join-Path $info.AppDataDir "local-agent-mode-sessions"
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

    # -- Claude executable ---------------------------------------------------
    # Ordered by authority, not by convenience. The package branch runs first
    # because it is the only source that reliably identifies the DESKTOP app,
    # and it is also what makes a WindowsApps install reachable at all: no
    # filesystem search path covers WindowsApps and a brute scan cannot read its
    # TrustedInstaller ACL.
    #
    # Reading the path off a running process looks more reliable but is not.
    # The Claude Code CLI also runs as a process named "claude", out of
    # %APPDATA%\Claude\claude-code\<version>\claude.exe. Measured on this
    # machine: 3 of 16 "claude" processes were the CLI. Taking the first
    # readable one can therefore cache the CLI path and turn every later
    # relaunch into a CLI start instead of an app start.
    $cliPattern = '[\\/]claude-code[\\/]'
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
                    if ($cand -and $cand -notmatch $cliPattern -and (Test-Path $cand)) {
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
                if ($cand -and $cand -notmatch $cliPattern -and (Test-Path $cand)) {
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
    # Deliberately restricted to VM-relevant logs. main.log is excluded: it is
    # the Electron main process log and it can sit frozen for a day at a time on
    # a perfectly healthy machine, so keying staleness on it is a guaranteed
    # false positive. Newest wins, and the age is recorded so a caller can
    # reject a candidate that predates its own lookback window.
    try {
        if ($info.LogDir -and (Test-Path $info.LogDir)) {
            $cands = @()
            foreach ($n in @('cowork_vm_node.log', 'coworkd.log', 'cowork-service.log')) {
                try {
                    $p = Join-Path $info.LogDir $n
                    if (Test-Path $p) {
                        $fi = Get-Item $p -ErrorAction SilentlyContinue
                        if ($fi) {
                            $cands += [pscustomobject]@{
                                Name          = $n
                                Path          = $fi.FullName
                                LastWriteTime = $fi.LastWriteTime
                                AgeHours      = [math]::Round( `
                                    ((Get-Date) - $fi.LastWriteTime).TotalHours, 2)
                                Length        = $fi.Length
                            }
                        }
                    }
                } catch { $null = $_ }
            }
            $cands = @($cands | Sort-Object LastWriteTime -Descending)
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
        if ($obj.SchemaVersion -ne 1) { return $null }
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
#endregion ClaudeEnv

# -- Runtime discovery -----------------------------------------------
# Everything below is read off this machine rather than assumed. See the
# ClaudeEnv region above for what is discovered, and for what deliberately
# is not: thresholds and log signatures are not location problems.
$script:ClaudeEnv = Get-ClaudeEnvironment

# -- Constants -------------------------------------------------------
$ToolkitVersion  = "6.0.2"
$ServiceName     = $script:ClaudeEnv.ServiceName
$ServiceExe      = $script:ClaudeEnv.ServiceProcessName
$ProcessName     = $script:ClaudeEnv.ProcessName
$ClaudeAppData   = $script:ClaudeEnv.AppDataDir
$VmCachePath     = $script:ClaudeEnv.VmCachePath
$BundlePath      = $script:ClaudeEnv.BundlePath

# Without the per-user data folder there is no cache to purge, no log to read
# and nowhere to write a transcript, so there is nothing useful left to do.
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

# -- Logging ---------------------------------------------------------
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }
$script:SessionTimestamp = "{0:yyyyMMdd_HHmmss}" -f (Get-Date)
$LogFile = Join-Path $LogDir "fix_$($script:SessionTimestamp).log"

# Persist the discovery result next to the logs. Sibling scripts reuse it, the
# package version stamp lets them tell a current record from one written before
# an app update, and a bug report can attach this instead of asking someone to
# run ten commands by hand.
$null = Save-ClaudeEnvironment -Environment $script:ClaudeEnv -Path $EnvArtifact

# Clean up logs older than 30 days
try {
    Get-ChildItem $LogDir -Filter "fix_*.log" -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
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
                           $_.LastWriteTime -lt (Get-Date).AddDays(-7) } |
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

    # 1. Cached path from a previous run. The claude-code filter matters here
    #    too: a cache written by an older build may hold the CLI path.
    if (Test-Path $ExePathCache) {
        try {
            $cached = (Get-Content $ExePathCache -Raw -ErrorAction Stop).Trim()
            if ($cached -and $cached -notmatch '[\\/]claude-code[\\/]' -and (Test-Path $cached)) {
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
            if ($appPath -and (Test-Path $appPath)) {
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
            if ($target -and $target -notmatch '[\\/]claude-code[\\/]' -and (Test-Path $target)) {
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

    # 1. Virtual Machine Platform -- the documented Cowork requirement.
    #
    # The wording below is deliberate. It used to say "Cowork runs in a Hyper-V
    # VM and cannot start without it", and a user on 2026-07-29 read that, ran
    # Get-WindowsOptionalFeature -FeatureName *Hyper-V*, saw all seven Hyper-V
    # features Enabled, and reported the script as wrong. That filter cannot
    # match VirtualMachinePlatform: the name does not contain "Hyper-V". They
    # are separate features and enabling Hyper-V does not enable this one.
    # Anything user-facing here has to say that outright.
    if ($script:IsAdmin) {
        try {
            $vmp = Get-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -ErrorAction Stop
            if ($vmp -and $vmp.State -ne "Enabled") {
                $r.Blocked = $true
                $r.Reason  = "The Virtual Machine Platform feature is $($vmp.State). Cowork cannot start without it."
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

    $script:ServicePrereq = Test-CoworkServicePrereq
    if ($script:ServicePrereq.Repaired) {
        Log "Service start type was Disabled, re-enabled (Automatic)" -Colour Green -Indent
    }
    if ($script:ServicePrereq.Blocked) {
        Log "[!] $($script:ServicePrereq.Reason)" -Colour Red -Indent
        foreach ($f in $script:ServicePrereq.Fixes) {
            Log "  -> $f" -Colour Yellow -Indent
        }
        return $false
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
    if ($script:IsAdmin) {
        try {
            Start-Service -Name $ServiceName -ErrorAction Stop
        } catch {
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
        if ($svcNow -and $svcNow.Status -eq "Running") { return $true }
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
function Test-RecentHcsErrors {
    <#
    .SYNOPSIS
        Checks for recent HCS errors. Now checks Information-level events too (v4.8.0).
        Returns: $null (clean), "shutdown_stale" (0xC037010D -- property query bug),
        "construct_failure" (0x800707DE), "guest_connect_failure" (isGuestConnected timeout),
        or "hcs_error" (other HCS issues).
    #>
    # Check 1: HCS Compute event log -- check ALL levels including Information
    try {
        # Check for 0xC037010D (shutdown failures) -- these are Information-level
        $hcsInfoFilter = @{
            LogName   = "Microsoft-Windows-Hyper-V-Compute-Operational"
            StartTime = (Get-Date).AddMinutes(-5)
        }
        $hcsInfoEvents = @(Get-WinEvent -FilterHashtable $hcsInfoFilter -MaxEvents 50 -ErrorAction SilentlyContinue)

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
            StartTime = (Get-Date).AddMinutes(-5)
        }
        $hcsCritEvents = @(Get-WinEvent -FilterHashtable $hcsCritFilter -MaxEvents 10 -ErrorAction SilentlyContinue)
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
    } catch { $null = $_ }

    # Check 2: Claude log files (keep existing logic)
    $hcsPatterns = @("HCS operation failed", "failed to create compute system",
                     "HcsWaitForOperationResult", "0x800707DE")
    # %ProgramData%\Claude does not exist on any current build, so the only
    # live log directory is the per-user one that discovery resolved.
    $claudeLogDirs = @()
    if ($script:ClaudeEnv.LogDir) { $claudeLogDirs += $script:ClaudeEnv.LogDir }
    $recentLogs = @()
    foreach ($dir in $claudeLogDirs) {
        if (Test-Path $dir) {
            $recentLogs += @(Get-ChildItem $dir -Filter "*.log" -ErrorAction SilentlyContinue |
                Where-Object { ((Get-Date) - $_.LastWriteTime).TotalMinutes -lt 5 })
        }
    }
    if ($recentLogs.Count -gt 0) {
        try {
            foreach ($logFile in $recentLogs) {
                try {
                    $content = Get-Content $logFile.FullName -Tail 50 -ErrorAction SilentlyContinue
                    $text = $content -join "`n"
                    foreach ($pattern in $hcsPatterns) {
                        if ($text -match [regex]::Escape($pattern)) { return "hcs_error" }
                    }
                } catch { $null = $_ }
            }
        } catch { $null = $_ }
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
        Runs hcsdiag.exe with a timeout. Returns output string or $null on timeout/error.
    #>
    param(
        [string[]]$Arguments,
        [int]$TimeoutSeconds = 15
    )
    $hcsdiagPath = "$env:SystemRoot\System32\hcsdiag.exe"
    if (-not (Test-Path $hcsdiagPath)) { return $null }
    $job = Start-Job -ScriptBlock {
        param($p, $a)
        & $p @a 2>&1 | Out-String
    } -ArgumentList $hcsdiagPath, $Arguments
    $completed = Wait-Job $job -Timeout $TimeoutSeconds
    if ($completed) {
        $result = Receive-Job $job
        Remove-Job $job -Force
        return $result
    } else {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        Log "hcsdiag timed out after ${TimeoutSeconds}s (args: $($Arguments -join ' '))" -Colour DarkYellow -Indent
        return $null
    }
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
    try {
        $hcsList = Invoke-HcsDiag -Arguments "list"
        if (-not $hcsList) { return 0 }
        $script:LastHcsListOk = $true
        $entries = Get-CoworkHcsGuids -ListOutput ([string]$hcsList)
        if ($entries.Count -eq 0) { return 0 }

        foreach ($guid in $entries) {
            try {
                if ($PSCmdlet.ShouldProcess($guid, "hcsdiag $Action")) {
                    Invoke-HcsDiag -Arguments $Action,$guid | Out-Null
                    $killed++
                }
            } catch { $null = $_ }
        }
    } catch { $null = $_ }
    return $killed
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
            exit 0
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
                $hcsOut = Invoke-HcsDiag -Arguments "list"
                $vmStillExists = $hcsOut -and ($hcsOut -match "cowork-vm")
                if ($vmStillExists) {
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
    $procs = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
    if ($procs.Count -gt 0) {
        $procs | Stop-Process -Force -ErrorAction SilentlyContinue
        Log "Killed $($procs.Count) Claude process(es)" -Colour Green -Indent
    } else {
        Log "No Claude processes found" -Colour DarkGray -Indent
    }

    # 3) Clean up any remaining HCS compute systems
    Log "[3/5] Cleaning HCS compute systems..." -Colour Yellow
    if ($script:IsAdmin) {
        Start-Sleep -Seconds 2
        try {
            $cleaned = Close-StaleHcsVms -Action kill
            if ($cleaned -gt 0) {
                Log "Killed $cleaned remaining compute system(s)" -Colour Green -Indent
            } else {
                Log "No remaining compute systems" -Colour DarkGray -Indent
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
    $remainingProcs = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
    $remainingSvc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    $remainingHcs = $false
    try {
        $hcsCheck = Invoke-HcsDiag -Arguments "list"
        $remainingHcs = $hcsCheck -and ($hcsCheck -match "cowork-vm")
    } catch { $null = $_ }
    $svcRunning = $remainingSvc -and $remainingSvc.Status -eq "Running"
    if ($remainingProcs.Count -eq 0 -and -not $remainingHcs) {
        if ($svcRunning) {
            Log "Clean shutdown complete (service idle, ready for relaunch)" -Colour Green -Indent
        } else {
            Log "Clean shutdown complete (service not running, may need Fix on relaunch)" -Colour DarkYellow -Indent
        }
    } else {
        if ($remainingProcs.Count -gt 0) { Log "Warning: $($remainingProcs.Count) Claude processes still running" -Colour DarkYellow -Indent }
        if ($remainingHcs) { Log "Warning: HCS compute system still present" -Colour DarkYellow -Indent }
    }
    Write-Host ""
    if ($svcRunning) {
        Write-Host "  Claude Desktop is shut down." -ForegroundColor Green
        Write-Host "  Service is idle and ready, relaunch should work immediately." -ForegroundColor Green
    } else {
        Write-Host "  Claude Desktop is shut down." -ForegroundColor Green
        Write-Host "  Service is not running, you may need to run Fix after relaunch." -ForegroundColor DarkYellow
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
    $diagProcs = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
    Log "  Claude procs  : $($diagProcs.Count) running" -Colour Cyan

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
            StartTime    = (Get-Date).AddHours(-1)
        }
        $diagEvents = @(Get-WinEvent -FilterHashtable $diagEvFilter -MaxEvents 5 -ErrorAction SilentlyContinue)
        if ($diagEvents) {
            Write-Host ""
            Log "  Recent service errors (last hour):" -Colour DarkYellow
            foreach ($de in $diagEvents) {
                $deTime = "{0:HH:mm}" -f $de.TimeCreated
                $deMsg  = ($de.Message -split "`n")[0]
                Log "    [$deTime] $deMsg" -Colour DarkGray
            }
        }
    } catch { $null = $_ }

    # HCS state via hcsdiag (v4.8.0)
    if ($script:IsAdmin) {
        try {
            $diagHcsList = Invoke-HcsDiag -Arguments "list"
            if ($diagHcsList) {
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
                Log "  HCS VMs       : Unable to query (timeout or not available)" -Colour DarkGray
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
        StartTime = (Get-Date).AddHours(-24)
    }
    try {
        $diagShutdownEvents = @(Get-WinEvent -FilterHashtable $diagEventFilter -ErrorAction SilentlyContinue |
            Where-Object { $_.Message -match "0xC037010D" })
        $last1h = @($diagShutdownEvents | Where-Object { $_.TimeCreated -gt (Get-Date).AddHours(-1) }).Count
        $last24h = $diagShutdownEvents.Count
        $statusColour = if ($last1h -gt 10) { "Red" } elseif ($last1h -gt 3) { "Yellow" } else { "Green" }
        Log "  Shutdown fails: $last1h (1h) / $last24h (24h)" -Colour $statusColour
    } catch { $null = $_ }

    # 0x800707DE frequency (v4.8.0)
    try {
        $diagConstructEvents = @(Get-WinEvent -FilterHashtable $diagEventFilter -ErrorAction SilentlyContinue |
            Where-Object { $_.Message -match "0x800707DE" })
        if ($diagConstructEvents.Count -gt 0) {
            Log "  Construct fails: $($diagConstructEvents.Count) (24h)" -Colour Red
        }
    } catch { $null = $_ }

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

        $wanted = @(
            'system-info.txt', 'supported-features-info.json', 'gpu-info.json',
            'main.log', 'cowork-service.log', 'coworkd.log',
            'cowork_vm_node.log', 'cowork_host_loop_debug.log'
        )
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
        $smallFiles   = @('system-info.txt', 'supported-features-info.json', 'gpu-info.json')

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
        $smallBytes = 0
        foreach ($s in $smallFiles) { if ($present.ContainsKey($s)) { $smallBytes += $present[$s] } }
        $bigNames = @($present.Keys | Where-Object { $smallFiles -notcontains $_ })
        $share    = [long]0
        if ($bigNames.Count -gt 0) {
            $share = [long](($bundleBudget - $reserve - $smallBytes) / $bigNames.Count)
            if ($share -lt 64KB) { $share = [long]64KB }
        }

        # Pass 3: copy, tailing anything over its share.
        foreach ($w in $present.Keys) {
            $p    = Join-Path $srcLogs $w
            $dest = Join-Path $bundleDir $w
            $len  = $present[$w]
            $cap  = if ($smallFiles -contains $w) { [long]$len } else { $share }
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
            Log "  Trimmed : $($trimmed.Count) large log(s) to $([math]::Round($share/1KB)) KB each" -Colour DarkGray
        }
        if ($missing.Count -gt 0) {
            # Absence is evidence. No cowork_vm_node.log means the VM has never
            # come up on this machine, which is worth saying out loud rather
            # than leaving as a quiet gap in the bundle.
            Log "  Absent  : $($missing -join ', ')" -Colour DarkGray
            if ($missing -contains 'cowork_vm_node.log') {
                Log "  Note    : no VM log at all means the Cowork VM has never started here" -Colour Yellow
            }
        }

        # The Service Control Manager entries for the three services that have
        # to be alive. This is where a service that refuses to start is
        # actually recorded.
        try {
            $svcNames = @($ServiceName, 'vmcompute', 'hns') | Where-Object { $_ }
            $pattern  = ($svcNames | ForEach-Object { [regex]::Escape($_) }) -join '|'
            $evts = @(Get-WinEvent -FilterHashtable @{
                          LogName      = 'System'
                          ProviderName = 'Service Control Manager'
                          StartTime    = (Get-Date).AddDays(-7)
                      } -MaxEvents 500 -ErrorAction Stop |
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
# BOOT PREP MODE -- Non-destructive vmcompute preparation (v4.8.4)
# ====================================================================
# When -BootPrep is set (called by boot task 30s after logon), do a
# lightweight vmcompute restart to clear stale state from previous session.
# This runs BEFORE the user opens Claude, preventing construct failures.
# Does NOT kill Claude, stop services, or touch any files.
if ($BootPrep) {
    Log "=== ClaudeFix Boot Prep v$ToolkitVersion ===" -Colour Cyan
    Log "[BootPrep] Non-destructive boot preparation (30s post-logon)" -Colour DarkGray
    if (-not $script:IsAdmin) {
        Log "[BootPrep] Not running as admin, cannot restart vmcompute" -Colour Yellow
        Save-Log
        exit 1
    }
    # Wait for vmcompute to be running (may still be starting after boot)
    $vmcWait = 0
    $vmcReady = $false
    while ($vmcWait -lt 30) {
        $vmcSvc = Get-Service -Name "vmcompute" -ErrorAction SilentlyContinue
        if ($vmcSvc -and $vmcSvc.Status -eq "Running") { $vmcReady = $true; break }
        if ($vmcWait -eq 0) {
            Log "[BootPrep] Waiting for vmcompute service to start..." -Colour DarkGray
        }
        Start-Sleep -Seconds 3
        $vmcWait += 3
    }
    if (-not $vmcReady) {
        Log "[BootPrep] vmcompute not running after 30s, cannot prepare" -Colour Yellow
        Save-Log
        exit 1
    }
    Log "[BootPrep] vmcompute is running" -Colour DarkGray
    # Check if there is an active workspace (do not disrupt healthy sessions)
    $activeWorkspace = $false
    try {
        $hcsList = Invoke-HcsDiag -Arguments "list"
        if ($hcsList) {
            if ($hcsList -match "cowork-vm") {
                # Distinguish stale VMs (from before reboot) from active ones
                # If Claude is running and cowork-vm exists, it is likely active
                $claudeRunning = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue).Count -gt 0
                if ($claudeRunning) {
                    $activeWorkspace = $true
                    Log "[BootPrep] Active workspace detected (Claude running + cowork-vm in HCS)" -Colour DarkGray
                } else {
                    # Stale VM from previous session -- clean it up
                    Log "[BootPrep] Stale cowork-vm found in HCS, cleaning" -Colour DarkYellow
                    try {
                        $cleaned = Close-StaleHcsVms -Action kill
                        if ($cleaned -gt 0) {
                            Log "[BootPrep] Closed $cleaned stale HCS system(s)" -Colour Green -Indent
                        }
                    } catch {
                        Log "[BootPrep] HCS cleanup failed: $($_.Exception.Message)" -Colour DarkGray
                    }
                }
            }
        }
    } catch {
        Log "[BootPrep] HCS check failed: $($_.Exception.Message)" -Colour DarkGray
    }
    if ($activeWorkspace) {
        Log "[BootPrep] Skipping vmcompute restart (active workspace)" -Colour Green
        Log "[BootPrep] Boot prep complete, no action needed" -Colour Green
        Save-Log
        exit 0
    }
    # Proactive vmcompute restart to clear stale boot state
    Log "[BootPrep] Restarting vmcompute to clear stale boot state..." -Colour DarkYellow
    try {
        Stop-Service vmcompute -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
        Start-Service vmcompute -ErrorAction SilentlyContinue
        $vmcElapsed = 0
        $vmcRunning = $false
        while ($vmcElapsed -lt 15) {
            Start-Sleep -Seconds 3
            $vmcElapsed += 3
            $vmcSvc2 = Get-Service -Name "vmcompute" -ErrorAction SilentlyContinue
            if ($vmcSvc2 -and $vmcSvc2.Status -eq "Running") { $vmcRunning = $true; break }
        }
        if ($vmcRunning) {
            Log "[BootPrep] vmcompute restarted successfully, ready for Claude" -Colour Green
        } else {
            Log "[BootPrep] vmcompute not running after restart" -Colour Yellow
        }
    } catch {
        Log "[BootPrep] vmcompute restart failed: $($_.Exception.Message)" -Colour Red
    }
    Log "[BootPrep] Boot prep complete" -Colour Green
    Save-Log
    exit 0
}

# ====================================================================
# SAFETY GATE -- Block when called by automation while user is active
# ====================================================================
# When -Quiet is set (called by health monitor or boot task), check if
# Claude Desktop is actively in use. This prevents killing Claude while
# the user is typing, Cowork is running, or Code is doing a task.
# Manual runs (no -Quiet) always proceed -- user explicitly wants a fix.
if ($Quiet) {
    $isActive = $false

    # Check 1: Any Claude process burning CPU (active request processing)
    $claudeCheck = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
    foreach ($cp in $claudeCheck) {
        try {
            $cpu1 = $cp.TotalProcessorTime.TotalMilliseconds
            Start-Sleep -Milliseconds 500
            $cp.Refresh()
            $cpu2 = $cp.TotalProcessorTime.TotalMilliseconds
            if (($cpu2 - $cpu1) -gt 100) { $isActive = $true; break }
        } catch { $null = $_ }
    }

    # Check 2: VM log active within 120s (Code may be thinking)
    #
    # Discovery already ranked the VM logs by recency. The list this replaces
    # led with C:\ProgramData\Claude\Logs\coworkd.log, on a hardcoded drive
    # letter, in a tree that does not exist.
    if (-not $isActive) {
        $safetyVmLog = $script:ClaudeEnv.VmLogFile
        if ($safetyVmLog -and (Test-Path $safetyVmLog)) {
            try {
                $ageSec = ((Get-Date) - (Get-Item $safetyVmLog).LastWriteTime).TotalSeconds
                if ($ageSec -lt 120) { $isActive = $true }
            } catch { $null = $_ }
        }
    }

    # Check 3: User input within 3 minutes (interactive sessions only)
    if (-not $isActive) {
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
                    if ($idleMs -lt 180000) { $isActive = $true }
                }
            }
        } catch { $null = $_ }
    }

    if ($isActive) {
        $msg = "BLOCKED: User/Code appears active, skipping automated fix"
        Log $msg -Colour Yellow
        Save-Log
        exit 0
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
        $hcsList = Invoke-HcsDiag -Arguments "list"
        if (-not $hcsList) {
            Log "hcsdiag unavailable or timed out, skipping HCS cleanup" -Colour DarkGray -Indent
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
            # the process is gone. BootPrep already draws this distinction.
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
            $cutoff = (Get-Date).AddDays(-7)
            $oldFiles = @(Get-ChildItem $sessionDir -Recurse -File -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -lt $cutoff })
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
# STEP 1 -- Kill all Claude processes
# ====================================================================
Log "[1/10] Terminating Claude processes..." -Colour Yellow

$claudeProcs = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
if ($claudeProcs.Count -gt 0) {
    # Discovery already resolved this, preferring the package, which is the only
    # authoritative source for the DESKTOP app. Reading it off a live process is
    # a fallback, not the primary, because the Claude Code CLI also runs as a
    # process named "claude" out of %APPDATA%\Claude\claude-code\<version>\.
    # Latching onto one of those caches the CLI path and turns every later
    # relaunch into a CLI start.
    if ($script:CapturedClaudeExe) {
        Log "Exe path from discovery ($($script:ClaudeEnv.ClaudeExeSource)): $($script:CapturedClaudeExe)" `
            -Colour DarkGray -Indent
    } else {
        foreach ($cp in $claudeProcs) {
            try {
                $exeFromProc = $cp.MainModule.FileName
                if ($exeFromProc -and $exeFromProc -notmatch '[\\/]claude-code[\\/]' -and
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

$remaining = @(Get-Process -ErrorAction SilentlyContinue | Where-Object {
    ($_.Name -eq $ProcessName) -or ($_.Name -eq $ServiceExe)
})
if ($remaining.Count -gt 0) {
    if ($PSCmdlet.ShouldProcess("$($remaining.Count) orphan process(es)", "Force-kill")) {
        foreach ($proc in $remaining) {
            try { $proc | Stop-Process -Force -ErrorAction Stop } catch { $null = $_ }
        }
        Start-Sleep -Seconds 1
        $stubborn = @(Get-Process -ErrorAction SilentlyContinue | Where-Object {
            ($_.Name -eq $ProcessName) -or ($_.Name -eq $ServiceExe)
        })
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
    # Method 1: hcsdiag (most reliable for HCS compute systems)
    if ($script:IsAdmin) {
        try {
            # Close-StaleHcsVms owns the hcsdiag list parsing and the kill,
            # including the ShouldProcess gate. This used to carry a second,
            # independent parser that expected the GUID on a line of its own,
            # which hcsdiag never emits.
            $killedCount = Close-StaleHcsVms -Action kill
            if ($killedCount -gt 0) {
                Log "Killed $killedCount orphan compute system(s)" -Colour Green -Indent
                $orphanKilled = $true
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
        } catch {
            Log "hcsdiag query failed: $($_.Exception.Message)" -Colour DarkGray -Indent
        }
    }
    # Note: Get-VM does not see HCS compute systems (like cowork-vm).
    # All cleanup is handled by hcsdiag above. (v4.8.0)

    # Method 2: Kill hung vmwp.exe (VM Worker Process)
    try {
        $vmwpProcs = @(Get-CimInstance Win32_Process -Filter "Name='vmwp.exe'" -ErrorAction SilentlyContinue)
        if ($vmwpProcs.Count -gt 0) {
            foreach ($vmwp in $vmwpProcs) {
                $vmwpPid = $vmwp.ProcessId
                $vmwpCmd = $vmwp.CommandLine
                if (-not $vmwpCmd) {
                    Log "vmwp.exe (PID $vmwpPid) has no command line, skipping" -Colour DarkGray -Indent
                    continue
                }
                # Extract GUID from command line
                if ($vmwpCmd -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})') {
                    $vmwpGuid = $Matches[1]
                    # Only kill if related to Claude's VM (match claude/cowork in hcsdiag or if this is the only VM)
                    if ($PSCmdlet.ShouldProcess("vmwp.exe PID $vmwpPid (GUID $vmwpGuid)", "Kill")) {
                        $vmwpKilled = $false
                        # Try hcsdiag kill first (cleaner)
                        if ($script:IsAdmin) {
                            try {
                                $killResult = Invoke-HcsDiag -Arguments "kill",$vmwpGuid
                                if ($null -ne $killResult) { $vmwpKilled = $true }
                            } catch { $null = $_ }
                        }
                        # Fallback: force-kill the process
                        if (-not $vmwpKilled) {
                            try {
                                Stop-Process -Id $vmwpPid -Force -ErrorAction Stop
                                $vmwpKilled = $true
                            } catch {
                                Log "[!] vmwp.exe (PID $vmwpPid) is unkillable, host restart may be needed" -Colour Red -Indent
                            }
                        }
                        if ($vmwpKilled) {
                            Log "WARNING: Force-killed vmwp.exe (PID $vmwpPid, GUID $vmwpGuid); VHDX corruption risk" -Colour DarkYellow -Indent
                            $orphanKilled = $true
                        }
                    }
                }
            }
        } else {
            Log "No hung VM worker processes found" -Colour DarkGray -Indent
        }
    } catch {
        Log "vmwp.exe check failed (non-critical): $($_.Exception.Message)" -Colour DarkGray -Indent
    }

    if (-not $orphanKilled) {
        Log "No orphan compute systems found" -Colour Green -Indent
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
            $hcsList = Invoke-HcsDiag -Arguments "list"
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
                # Verify the task actually started something (give it 3 seconds)
                Start-Sleep -Seconds 3
                $claudeProc = Get-Process -Name $ProcessName -ErrorAction SilentlyContinue
                if ($claudeProc) {
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
    # Detection strategy (in order of reliability):
    #   A. Log file: cowork_vm_node.log -- "Startup complete" or "Keepalive"
    #   B. Hyper-V VM state: Get-VM "claudevm" shows Running + heartbeat
    #   C. Log file: step markers like "guest_vsock_connect completed"
    #   D. File size stability fallback (last resort)
    #   E. cowork-service.log -- guest connection state via isGuestConnected RPC
    # The named pipe RPC is NOT usable (requires signed client executable).
    # ====================================================================
    Log "[9/10] Waiting for Cowork workspace..." -Colour Yellow

    $vmReady = $false

    # Discovery already selected the newest VM-relevant log at startup, and it
    # records each candidate's age, so a stale file is reported rather than
    # silently preferred.
    #
    # The list this replaces led with C:\ProgramData\Claude\Logs\coworkd.log.
    # That tree does not exist on any current build, the drive letter was
    # hardcoded, and coworkd.log itself was last written in March.
    $vmLogFile  = $script:ClaudeEnv.VmLogFile
    $vmCandList = @($script:ClaudeEnv.VmLogCandidates)
    if ($vmLogFile -and $vmCandList.Count -gt 0) {
        Log "Using log file: $vmLogFile (last written $($vmCandList[0].AgeHours)h ago)" -Colour DarkGray -Indent
        foreach ($cand in $vmCandList) {
            if ($cand.Name -ne $vmCandList[0].Name -and $cand.AgeHours -gt 24) {
                Log "Ignoring stale candidate $($cand.Name) ($($cand.AgeHours)h old)" -Colour DarkGray -Indent
            }
        }
    } else {
        Log "No VM log file found; readiness will rely on guest and HCS signals" -Colour DarkYellow -Indent
    }

    # Safe size read for the chosen VM log. Returns 0 when no log was found,
    # or when the file vanished between the existence test and the read, which
    # would otherwise be a property access on $null under StrictMode.
    function Get-VmLogSize {
        if (-not $vmLogFile) { return [long]0 }
        $item = Get-Item $vmLogFile -ErrorAction SilentlyContinue
        if (-not $item) { return [long]0 }
        return [long]$item.Length
    }

    # Record the log file size at the START so we only check new entries
    $logBaselineSize = Get-VmLogSize

    # Helper: check recent log entries for boot completion markers
    function Test-VmLogReady {
        param([long]$Baseline)
        # $vmLogFile is $null whenever no candidate log existed. Test-Path $null
        # throws, and this call site sits outside the try below, so the throw
        # unwound to the outer catch and abandoned every remaining step.
        if (-not $vmLogFile -or -not (Test-Path $vmLogFile)) { return $null }
        try {
            $fi = Get-Item $vmLogFile -ErrorAction Stop
            if ($fi.Length -le $Baseline) { return $null }
            # Read only the new portion of the log
            $stream = $null
            $reader = $null
            try {
                $stream = New-Object System.IO.FileStream(
                    $vmLogFile, [System.IO.FileMode]::Open,
                    [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                $stream.Position = $Baseline
                $reader = New-Object System.IO.StreamReader($stream)
                $newContent = $reader.ReadToEnd()
            } finally {
                if ($reader) { try { $reader.Close() } catch { $null = $_ } }
                if ($stream) { try { $stream.Close() } catch { $null = $_ } }
            }
            # Check for completion markers (most definitive first)
            # Old markers (cowork_vm_node.log)
            if ($newContent -match "Startup complete") { return "startup-complete" }
            if ($newContent -match "\[Keepalive\]") { return "keepalive" }
            if ($newContent -match "guest_vsock_connect completed") { return "vsock-connected" }
            if ($newContent -match "sdk_install completed") { return "sdk-installed" }
            # New markers (coworkd.log in ProgramData, v5.1.0)
            if ($newContent -match "\[process:[0-9a-f-]+\] started PID") { return "process-started" }
            if ($newContent -match "\[coworkd\] mounted .+ at /sessions/") { return "mounts-ready" }
            if ($newContent -match "full egress mode enabled") { return "egress-ready" }
            return $null
        } catch { return $null }
    }

    # Helper: check HCS compute system state via hcsdiag.
    # Invoke-HcsDiag already runs the call in a timed-out background job, which
    # is what keeps this from hanging when vmcompute is unstable.
    function Test-HyperVReady {
        try {
            $hcsList = Invoke-HcsDiag -Arguments "list"
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

    # Skip Hyper-V cmdlet checks if vmcompute was just restarted (may hang on unstable service)
    $skipHvChecks = [bool]$hcsDetected
    if ($skipHvChecks) {
        Log "Skipping Hyper-V VM checks (vmcompute was just restarted)" -Colour DarkGray -Indent
    }

    # Do not wait at all for a workspace that cannot come up. Step 7 has already
    # printed what is wrong and what to run, and none of it changes while this
    # script is running, so the only thing waiting produces is the same message
    # on repeat.
    $prereqStillBlocked = ($null -ne $script:ServicePrereq -and $script:ServicePrereq.Blocked)
    if ($prereqStillBlocked) {
        Log "Not waiting for the workspace: the blocker above has to be cleared first." -Colour Red -Indent
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
                $logStatus = Test-VmLogReady -Baseline $logBaselineSize
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
                if ($vmLogFile) {
                    $logBaselineSize = Get-VmLogSize
                }
                Log "Service restarted, continuing to wait..." -Colour DarkGray -Indent
            } elseif ($guestState -eq "guest-error" -and $vmElapsed -ge 60) {
                Log "[!] Guest connection errors detected in cowork-service.log" -Colour DarkYellow -Indent
            }
        }

        # No-progress detector: if 60s in and nothing happening, escalate (v5.0.0)
        if ($vmElapsed -ge 60 -and -not $vmReady) {
            $hasLogActivity = ($null -ne (Test-VmLogReady -Baseline $logBaselineSize))
            $hasGuestActivity = ($guestState -and $guestState -ne "no-polling")
            $hasHcsVm = $false
            try {
                $hcsCheck = Invoke-HcsDiag -Arguments "list"
                $hasHcsVm = $hcsCheck -and ($hcsCheck -match "cowork-vm")
            } catch { $null = $_ }
            if ($hasHcsVm) {
                Log "HCS VM present but no log/guest activity, VM may be stuck" -Colour DarkGray -Indent
            }
            if (-not $hasLogActivity -and -not $hasGuestActivity) {
                if (-not $script:NoProgressEscalated) {
                    $script:NoProgressEscalated = $true
                    Log "[!] No progress after ${vmElapsed}s, no log activity, no guest polling" -Colour Red -Indent
                    Log "Escalating: killing Claude, cleaning HCS, relaunching..." -Colour Yellow -Indent
                    # Kill Claude
                    Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Stop-Process -Force
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
                    if ($vmLogFile) {
                        $logBaselineSize = Get-VmLogSize
                    }
                    # Don't reset vmElapsed -- let the outer timeout still apply
                    Log "Waiting for workspace after relaunch..." -Colour DarkGray -Indent
                    continue
                }
            }
        }

        # A. Check log file for boot completion (new entries since baseline)
        $logStatus = Test-VmLogReady -Baseline $logBaselineSize
        if ($logStatus -in @("startup-complete", "keepalive", "process-started", "egress-ready")) {
            $vmReady = $true
            Log "Workspace ready (log: $logStatus)" -Colour Green -Indent
            break
        }

        # A2. Fallback: workspace may already be running (baseline captured after boot).
        #     Check the tail of the log for recent boot markers regardless of baseline.
        if (-not $logStatus -and $vmElapsed -ge 30 -and $vmLogFile -and (Test-Path $vmLogFile)) {
            try {
                $fi = Get-Item $vmLogFile -ErrorAction Stop
                $recentlyWritten = ($fi.LastWriteTime -gt (Get-Date).AddMinutes(-3))
                if ($recentlyWritten -and $fi.Length -gt 0) {
                    $tailSize = [math]::Min($fi.Length, 8192)
                    $stream = $null
                    $reader = $null
                    try {
                        $stream = New-Object System.IO.FileStream(
                            $vmLogFile, [System.IO.FileMode]::Open,
                            [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                        $stream.Position = $fi.Length - $tailSize
                        $reader = New-Object System.IO.StreamReader($stream)
                        $tailContent = $reader.ReadToEnd()
                    } finally {
                        if ($reader) { try { $reader.Close() } catch { $null = $_ } }
                        if ($stream) { try { $stream.Close() } catch { $null = $_ } }
                    }
                    if ($tailContent -match "Startup complete|\[Keepalive\]|started PID|full egress mode") {
                        $vmReady = $true
                        Log "Workspace ready (log tail: recently active with boot markers)" -Colour Green -Indent
                        break
                    }
                }
            } catch { $null = $_ }
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
        $finalLog = Test-VmLogReady -Baseline $logBaselineSize
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
    # -SkipLaunch never enters the Step 9 branch, so $vmLogFile and
    # $logBaselineSize below only exist when that branch ran. Without this
    # guard, -SkipLaunch purged 13.5 GB and then died under StrictMode with
    # the cache already gone.
    $script:DeepEscalated = $true
    Log "" -Colour White
    Log "Smart mode: workspace not ready after quick fix, escalating to Deep" -Colour Yellow
    # Phase 0: Kill Claude + HCS cleanup
    Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Stop-Process -Force
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
        if ($vmLogFile) {
            $logBaselineSize = Get-VmLogSize
        }
    } else {
        Log "Nothing was relaunched, so there is nothing to wait for" -Colour DarkYellow -Indent
    }
    while ($relaunched -and $escalateElapsed -lt $escalateTimeout) {
        Start-Sleep -Seconds 5
        $escalateElapsed += 5
        $logStatus = Test-VmLogReady -Baseline $logBaselineSize
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
            $logBaselineSize = Get-VmLogSize

            # Wait up to 60s for workspace
            $retryElapsed = 0
            while ($retryElapsed -lt 60) {
                Start-Sleep -Seconds 5
                $retryElapsed += 5
                $retryStatus = Test-VmLogReady -Baseline $logBaselineSize
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
$fProcs = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
$svcOk  = $fSvc -and $fSvc.Status -eq "Running"

$svcStatusText  = "Not found"
$svcStatusColor = "Yellow"
if ($fSvc) {
    $svcStatusText = "$($fSvc.Status)"
    if ($svcOk) { $svcStatusColor = "Green" }
}
Write-Host "  VM Service:       $svcStatusText" -ForegroundColor $svcStatusColor
Write-Host "  Claude processes: $($fProcs.Count) active" -ForegroundColor Cyan

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
        StartTime    = (Get-Date).AddHours(-1)
    }
    $events = Get-WinEvent -FilterHashtable $evFilter -MaxEvents 3 -ErrorAction SilentlyContinue

    if ($events) {
        Write-Host ""
        Log "Recent service errors (last hour):" -Colour DarkYellow -Indent
        foreach ($ev in $events) {
            $evTime = "{0:HH:mm}" -f $ev.TimeCreated
            $evMsg  = ($ev.Message -split "`n")[0]
            Log "  [$evTime] $evMsg" -Colour DarkGray -Indent
        }
    }
} catch { $null = $_ }

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
