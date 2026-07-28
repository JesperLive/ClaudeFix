<#
.SYNOPSIS
    Claude Desktop / Cowork, Health Monitor

.DESCRIPTION
    Persistent background monitor that detects VirtioFS/Plan9 mount failures
    and HCS/vmcompute errors in Claude Desktop and automatically runs the fix
    script.

    Monitors:
    - Claude log files for "bad address", mount failure, and HCS error messages
    - CoworkVMService status (stopped while Claude is running)
    - Windows Event Log for service, Hyper-V, and HCS Compute errors/warnings
    - WinNAT rules (VM network connectivity)
    - Hyper-V Integration Services heartbeat
    - VM log staleness (hung VM detection)
    - vmcompute handle leak detection (HCS service health)
    - cowork-service.log for isGuestConnected RPC timeout patterns
    - Host clock drift (NTP/time sync)

    SAFETY FEATURES (v3.3+):
    - Auto-fix is BLOCKED when user is active (Claude in focus, recent input, VM busy, CPU active)
    - Claude Code session awareness (detects active Code sessions via session files + renderer logs)
    - Cowork session awareness (detects active Cowork VM processes via log parsing)
    - Claude startup grace period (defers auto-fix for 120s after Claude launch)
    - Extended CPU sampling (3 samples over 3s to catch bursty API-wait patterns)
    - Electron-aware window detection (GetWindowThreadProcessId, not MainWindowHandle)
    - Session 0 safe (falls back to process heuristics when Win32 APIs are unavailable)
    - Startup grace period (skips first 180s to avoid matching pre-existing events)
    - Consecutive-check requirements on all heuristic triggers
    - Event log matching tightened to Claude-specific messages only
    - Cooldown between fixes (default 5 min)
    - VM log window extended to 120s (covers Code thinking phases)

    When a failure is detected AND the user is idle, automatically runs
    Fix-ClaudeDesktop.ps1 to reset the VM and relaunch Claude.

    Designed to run as a hidden scheduled task (installed by Prevent-ClaudeIssues.ps1)
    but can also be started manually for foreground monitoring.

.PARAMETER PollInterval
    Seconds between health checks (default: 30).

.PARAMETER Cooldown
    Minutes to wait between auto-fix runs (default: 5).

.PARAMETER Quiet
    Suppress console output (for scheduled task use).

.NOTES
    Version : 6.0.1
    Author  : Jesper Driessen
    Licence : MIT
#>

[CmdletBinding()]
param(
    [int]$PollInterval = 30,
    [int]$Cooldown = 5,
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

# -- Runtime discovery ---------------------------------------------------
# -SkipCacheInventory because this polls: the cache walk is the expensive half
# of discovery (measured 590ms of a 775ms total) and nothing here needs the
# VHDX inventory.
$script:ClaudeEnv = Get-ClaudeEnvironment -SkipCacheInventory

# -- Constants -----------------------------------------------------------
$ToolkitVersion = "6.0.1"
$ServiceName    = $script:ClaudeEnv.ServiceName
$ProcessName    = $script:ClaudeEnv.ProcessName
$ClaudeAppData  = $script:ClaudeEnv.AppDataDir

# The %ProgramData%\Claude branch is gone. That tree does not exist on any
# current build, so the "prefer ProgramData" line above resolved to the AppData
# path every time, which made three separate fallbacks downstream dead code.
$ClaudeLogDir   = $script:ClaudeEnv.LogDir

if (-not $ClaudeAppData -or -not $ClaudeLogDir) {
    Write-Host "  [!] Could not locate the per-user Claude folder. Nothing to monitor." -ForegroundColor Red
    exit 1
}
$WatchLogDir    = Join-Path $ClaudeAppData "watch-logs"

# -- Legacy error signatures ---------------------------------------------
# SECONDARY set, kept for continuity. None of these eleven strings occurs
# anywhere in ~35 MB of current logs on build 1.24012.9.0: 0 matches for all
# eleven. The one previously described as "the MOST RELIABLE" trigger has no
# corroboration on this build at all.
#
# Primary detection now runs on signals that verifiably exist. See
# Test-PrimarySignals below.
#
# These are matched literally. The old set mixed a regex,
# "isGuestConnected.*timeout", into a list consumed through [regex]::Escape(),
# so that entry could never match anything; the literal
# "Request timed out: isGuestConnected" above it covers the real message.
$LegacyErrorPatterns = @(
    "Plan9 mount failed",
    "bad address",
    "failed to ensure virtiofs mount",
    "RPC error -1",
    "HCS operation failed",
    "failed to create compute system",
    "HcsWaitForOperationResult",
    "VM is already running",
    "Request timed out: isGuestConnected",
    "guest connection failed"
)

# -- Win32 APIs for user activity detection ------------------------------
# NOTE: These APIs only work in interactive sessions (Session 1+).
# In Session 0 (SYSTEM scheduled tasks), they return zero/stale data.
# We detect this and fall back to process-only heuristics.
$script:IsInteractiveSession = $false
try {
    $sessionId = (Get-Process -Id $PID -ErrorAction Stop).SessionId
    $script:IsInteractiveSession = ($sessionId -gt 0)
} catch { $null = $_ }

# Set once the Add-Type below is confirmed to have produced a usable type.
$script:Win32ActivityLoaded = $false

Add-Type -ErrorAction SilentlyContinue -TypeDefinition @'
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

# Verify the type actually loaded, rather than assuming it did.
#
# The Add-Type above runs with -ErrorAction SilentlyContinue and its result was
# never checked. If it failed, the first [Win32Activity] reference threw inside
# the single try wrapping the user-activity checks, so checks 2 through 6 were
# all skipped, the empty catch swallowed the error, and the function returned
# $false, which means "user is not active". The auto-fix then proceeded while
# someone was working.
try {
    $null = [Win32Activity]
    $script:Win32ActivityLoaded = $true
} catch {
    $null = $_
    $script:Win32ActivityLoaded = $false
}

# -- Find Fix script -----------------------------------------------------
$myDir = if ($PSCommandPath) { Split-Path $PSCommandPath -Parent } else { $PWD.Path }
$fixScript = Join-Path $myDir "Fix-ClaudeDesktop.ps1"

if (-not (Test-Path $fixScript)) {
    $fallbacks = @(
        "C:\ClaudeFix\Fix-ClaudeDesktop.ps1",
        (Join-Path $env:USERPROFILE "Desktop\Fix-ClaudeDesktop.ps1"),
        (Join-Path $env:USERPROFILE "Documents\Fix-ClaudeDesktop.ps1")
    )
    foreach ($fb in $fallbacks) {
        if (Test-Path $fb) { $fixScript = $fb; break }
    }
}

# The fatal check for a missing Fix script is deliberately NOT here. It used
# to be, which meant it ran before $script:WatchLogFile existed, so its only
# output was a Write-Host with no console to write to under
# -WindowStyle Hidden. Task Scheduler retried three times and gave up in
# silence. See the check after Write-WatchLog is defined, below.

# -- Logging -------------------------------------------------------------
if (-not (Test-Path $WatchLogDir)) { New-Item $WatchLogDir -ItemType Directory -Force | Out-Null }

# Clean old watch logs (>30 days)
try {
    Get-ChildItem $WatchLogDir -Filter "watch_*.log" -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
        Remove-Item -Force -ErrorAction SilentlyContinue
} catch { $null = $_ }

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
    } catch { $null = $_ }
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
        try {
            Get-ChildItem $WatchLogDir -Filter "watch_*.log" -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
                Remove-Item -Force -ErrorAction SilentlyContinue
        } catch { $null = $_ }
    }
}

# Fatal check for the Fix script, placed AFTER Write-WatchLog exists so the
# reason lands in the watch log rather than on a console that is not there.
if (-not (Test-Path $fixScript)) {
    Write-WatchLog "FATAL: Fix-ClaudeDesktop.ps1 not found. Looked next to this script at '$myDir' and in the standard fallback locations. The health monitor cannot start."
    if (-not $Quiet) {
        Write-Host "  [!] Fix-ClaudeDesktop.ps1 not found. Health monitor cannot start." -ForegroundColor Red
        Write-Host "      Reason written to: $script:WatchLogFile" -ForegroundColor DarkGray
    }
    exit 1
}

# -- hcsdiag helper (timeout-protected) ----------------------------------
function Invoke-HcsDiag {
    <#
    .SYNOPSIS
        Runs hcsdiag.exe with a timeout. Returns output string or $null on timeout.
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
        Write-WatchLog "hcsdiag timed out after ${TimeoutSeconds}s (args: $($Arguments -join ' '))"
        return $null
    }
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

    $now = Get-Date
    if (($now - $script:LastHcsListTime).TotalSeconds -lt $MaxAgeSeconds) {
        return $script:LastHcsListText
    }
    $script:LastHcsListTime = $now
    $script:LastHcsListText = Invoke-HcsDiag -Arguments "list"
    return $script:LastHcsListText
}

# -- State ---------------------------------------------------------------
$script:StartTime            = Get-Date
$script:LastHcsListTime      = [datetime]::MinValue
$script:LastHcsListText      = $null
$script:LastFixTime          = [datetime]::MinValue
$script:LogBaselines         = @{}   # Track file sizes per log file
$script:FixCount             = 0
$script:LastNatCheckTime     = [datetime]::MinValue
$script:LastTimeSyncCheck    = [datetime]::MinValue
$script:LastHeartbeatStatus  = $null
$script:HeartbeatFailCount   = 0     # Consecutive heartbeat failures
$script:VmLogStaleCount      = 0     # Consecutive stale checks
$script:VmLogEverActive      = $false  # Has VM log ever been active this session?
$script:ServiceDownCount     = 0     # Consecutive service-down checks
$script:EventLogHitCount     = 0     # Consecutive event log hits
$script:LastVmcomputeCheck   = [datetime]::MinValue
$script:VmcomputeLeakCount   = 0     # Consecutive vmcompute handle leak checks
$script:GuestFailCount       = 0     # Consecutive guest connection failures
$script:LastHcsStateLogTime  = [datetime]::MinValue
$script:LastShutdownFailLogTime = [datetime]::MinValue
$script:FixHistory           = New-Object System.Collections.ArrayList
$script:StatePath            = Join-Path $WatchLogDir "watch-state.json"
# Seeded from the current state so the first poll does not report a transition
# that did not happen.
$script:ClaudeWasRunning     = (@(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue).Count -gt 0)

function Import-WatchState {
    <#
    .SYNOPSIS
        Restores cooldown and fix history from disk.
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
    #>
    try {
        if (Test-Path $script:StatePath) {
            $raw = Get-Content $script:StatePath -Raw -ErrorAction Stop
            if ($raw) {
                $st = $raw | ConvertFrom-Json -ErrorAction Stop
                if ($st.LastFixTime) {
                    $script:LastFixTime = [datetime]::Parse($st.LastFixTime,
                        [System.Globalization.CultureInfo]::InvariantCulture)
                }
                if ($st.FixCount) { $script:FixCount = [int]$st.FixCount }
                foreach ($h in @($st.FixHistory)) {
                    if (-not $h) { continue }
                    $null = $script:FixHistory.Add([datetime]::Parse($h,
                        [System.Globalization.CultureInfo]::InvariantCulture))
                }
            }
        }
    } catch { $null = $_ }

    # Seed from the log itself, not from the saved file.
    try {
        $seedLog = $script:ClaudeEnv.VmLogFile
        if ($seedLog -and (Test-Path $seedLog)) {
            $seedAge = ((Get-Date) - (Get-Item $seedLog).LastWriteTime).TotalMinutes
            if ($seedAge -lt 60) { $script:VmLogEverActive = $true }
        }
    } catch { $null = $_ }
}

function Save-WatchState {
    <#
    .SYNOPSIS
        Persists cooldown and fix history. Never throws; this is bookkeeping.
    #>
    try {
        $cutoff = (Get-Date).AddHours(-6)
        # Trimmed on write. $script:FixHistory was appended to and never
        # trimmed, so it grew for the life of the daemon and every
        # Test-PersistentFailure call filtered a list that only got longer.
        $recent = @($script:FixHistory | Where-Object { $_ -gt $cutoff })
        $script:FixHistory.Clear()
        foreach ($r in $recent) { $null = $script:FixHistory.Add($r) }

        $payload = [ordered]@{
            LastFixTime = $script:LastFixTime.ToString('o')
            FixCount    = $script:FixCount
            FixHistory  = @($recent | ForEach-Object { $_.ToString('o') })
        }
        [System.IO.File]::WriteAllText($script:StatePath,
            ($payload | ConvertTo-Json -Depth 3),
            (New-Object System.Text.UTF8Encoding($false)))
    } catch { $null = $_ }
}

# Baselines are seeded lazily by Test-LogsForErrors: any stream it has not seen
# before starts at that file's current size. The eager loop that used to live
# here keyed on path alone, which no longer matches the (path, creation time)
# key the detector uses, and it would have had to be kept in sync by hand.

# -- Detection functions -------------------------------------------------

function Test-StartupGracePeriod {
    <#
    .SYNOPSIS
        Returns $true if the monitor just started and should skip heuristic checks.
        Grace period: 180 seconds. This prevents matching pre-existing event log
        entries or stale logs from before the monitor was running.
    #>
    return ((Get-Date) - $script:StartTime).TotalSeconds -lt 180
}

function Test-PersistentFailure {
    <#
    .SYNOPSIS
        Returns $true if the fix has been attempted 3+ times in 30 minutes,
        indicating an unrecoverable error (e.g., HCS JSON corruption).
    #>
    $cutoff = (Get-Date).AddMinutes(-30)
    $recentFixes = @($script:FixHistory | Where-Object { $_ -gt $cutoff })
    return $recentFixes.Count -ge 3
}

function Test-LogsForErrors {
    <#
    .SYNOPSIS
        Scans all Claude log files for new VirtioFS error messages since last check.
        Returns the error pattern found, or $null if clean.
        This is the MOST RELIABLE trigger -- actual error strings in Claude's own logs.
    #>
    if (-not (Test-Path $ClaudeLogDir)) { return $null }

    $logFiles = Get-ChildItem $ClaudeLogDir -Filter "*.log" -ErrorAction SilentlyContinue
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

            foreach ($pattern in $LegacyErrorPatterns) {
                if ($newContent -match [regex]::Escape($pattern)) {
                    return $pattern
                }
            }
        } catch {
            # Read failed. The baseline stays put, so this window is retried on
            # the next poll rather than silently skipped.
            $null = $_
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
    $claude = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
    if ($claude.Count -eq 0) {
        $script:ServiceDownCount = 0
        return $true
    }

    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
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

function Test-EventLogErrors {
    <#
    .SYNOPSIS
        Checks Windows Event Log for Claude-specific errors.
        TIGHTENED in v3.0:
        - VMMS check now requires "claude" in the message (no more "failed"/"unexpected" wildcards)
        - Requires 2 consecutive hits before triggering
        - Skipped during startup grace period
    #>
    # Skip during startup grace period (avoids pre-existing events)
    if (Test-StartupGracePeriod) { return $null }

    $lookback = [math]::Max($PollInterval, 60)
    $since = (Get-Date).AddSeconds(-$lookback)
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
        $events = Get-WinEvent -FilterHashtable $evFilter -MaxEvents 1 -ErrorAction SilentlyContinue
        if ($events) {
            $msg = "$($events[0].Message)"
            foreach ($pattern in $LegacyErrorPatterns) {
                if ($msg -match [regex]::Escape($pattern)) {
                    $foundIssue = "CoworkVMService: $(& $firstLine $msg)"
                    break
                }
            }
        }
    } catch { $null = $_ }

    # Check 2: Hyper-V Worker -- must mention claude/Plan9/virtio/shared memory
    if (-not $foundIssue) {
        try {
            $hvWorkerFilter = @{
                LogName   = "Microsoft-Windows-Hyper-V-Worker-Admin"
                Level     = @(1, 2)  # Critical, Error only (dropped warnings)
                StartTime = $since
            }
            $hvEvents = Get-WinEvent -FilterHashtable $hvWorkerFilter -MaxEvents 1 -ErrorAction SilentlyContinue
            if ($hvEvents) {
                $msg = "$($hvEvents[0].Message)"
                if ($msg -match "claude" -or $msg -match "Plan9" -or $msg -match "virtio" -or $msg -match "shared memory") {
                    $foundIssue = "Hyper-V Worker: $(& $firstLine $msg)"
                }
            }
        } catch { $null = $_ }
    }

    # Check 3: Hyper-V VMMS -- MUST mention "claude" (no more generic "failed" matching)
    if (-not $foundIssue) {
        try {
            $vmmsFilter = @{
                LogName   = "Microsoft-Windows-Hyper-V-VMMS-Admin"
                Level     = @(1, 2)  # Critical, Error
                StartTime = $since
            }
            $vmmsEvents = Get-WinEvent -FilterHashtable $vmmsFilter -MaxEvents 1 -ErrorAction SilentlyContinue
            if ($vmmsEvents) {
                $msg = "$($vmmsEvents[0].Message)"
                # TIGHTENED: must specifically mention claude or cowork
                if ($msg -match "claude" -or $msg -match "cowork") {
                    $foundIssue = "Hyper-V VMMS: $(& $firstLine $msg)"
                }
            }
        } catch { $null = $_ }
    }

    # Check 4: HCS Compute log -- dedicated channel for Host Compute Service
    if (-not $foundIssue) {
        try {
            $hcsFilter = @{
                LogName   = "Microsoft-Windows-Hyper-V-Compute-Admin"
                Level     = @(1, 2)  # Critical, Error
                StartTime = $since
            }
            $hcsEvents = Get-WinEvent -FilterHashtable $hcsFilter -MaxEvents 1 -ErrorAction SilentlyContinue
            if ($hcsEvents) {
                $msg = "$($hcsEvents[0].Message)"
                if ($msg -match "claude" -or $msg -match "cowork" -or $msg -match "failed to create" -or $msg -match "HcsWaitForOperationResult") {
                    $foundIssue = "HCS Compute: $(& $firstLine $msg)"
                }
            }
        } catch { $null = $_ }
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
        disk nor a missing file is something a VM restart repairs.
    #>
    if (Test-StartupGracePeriod) { return $null }

    $cachePath = $script:ClaudeEnv.VmCachePath
    if (-not $cachePath -or -not (Test-Path $cachePath)) { return $null }

    # 1. Readability.
    try {
        $null = @(Get-ChildItem $cachePath -Force -ErrorAction Stop | Select-Object -First 1)
    } catch {
        return "VM cache directory is not readable: $($_.Exception.Message)"
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
    } catch { $null = $_ }

    return $null
}

function Test-WinNatHealth {
    <#
    .SYNOPSIS
        Checks that a WinNAT rule exists for the Cowork VM's network.
        Without NAT, the VM has no outbound connectivity, causing API calls
        and mount operations to fail silently.
        Returns $null if healthy, or a description string if unhealthy.
        Does NOT trigger auto-fix -- logs warning only.
    #>
    # Only check every 60 seconds (NAT doesn't change that often)
    $now = Get-Date
    if (($now - $script:LastNatCheckTime).TotalSeconds -lt 60) { return $null }
    $script:LastNatCheckTime = $now

    try {
        # The Hyper-V "Default Switch" provides NAT natively through HNS
        # without requiring a WinNAT (Get-NetNat) rule. If it exists, NAT is fine.
        $defaultSwitch = Get-VMSwitch -Name "Default Switch" -ErrorAction SilentlyContinue
        if ($defaultSwitch) { return $null }

        $natRules = @(Get-NetNat -ErrorAction SilentlyContinue)
        if ($natRules.Count -eq 0) {
            # Advisory only. This block used to create a NAT rule by itself.
            #
            # It selected the first internal Hyper-V switch matching
            # "WSL|claude|nat", which on a developer machine is very often WSL's
            # own switch, then claimed a NAT over it with a hardcoded /24 and no
            # confirmation. Windows permits only a small number of NAT
            # instances, so taking one over another product's switch can break
            # that product's networking AND leave no room to create the correct
            # rule afterwards.
            #
            # The caller documents this check as warning-only and does not
            # auto-fix on it. Silently reconfiguring host networking from an
            # unattended monitor is not what that contract promises, so this
            # reports and leaves the decision to a person.
            $internalSwitches = @(Get-VMSwitch -SwitchType Internal -ErrorAction SilentlyContinue)
            if ($internalSwitches.Count -eq 0) {
                # Nothing to NAT. Without an internal switch there is no subnet
                # for a NAT rule to serve, so a missing rule is not a fault.
                # This used to warn every 60 seconds, forever, on any machine
                # with no Hyper-V switch at all.
                return $null
            }

            $coworkSwitch = @($internalSwitches | Where-Object { $_.Name -match "claude|cowork" } |
                              Select-Object -First 1)
            if ($coworkSwitch.Count -gt 0) {
                return "No WinNAT rule found while internal switch '$($coworkSwitch[0].Name)' exists. If the VM has no network, create a NAT for its subnet with New-NetNat."
            }
            return $null
        }
    } catch {
        # Get-NetNat is not present on every SKU. Nothing to report.
        $null = $_
    }
    return $null
}

function Test-VmHeartbeat {
    <#
    .SYNOPSIS
        Always returns $null. Retained as a no-op so the check numbering in the
        main loop does not shift.
    .DESCRIPTION
        This measured the Hyper-V Integration Services heartbeat via Get-VM.

        Get-VM enumerates VMMS virtual machines. The Cowork VM is an HCS
        compute system and never appears there: Get-VM returns nothing at all
        on this machine. So $claudeVm was always $null, the function returned
        at the first guard every time, and the "heartbeat(x3)" monitor that the
        startup banner advertises has never fired once in the product's life.

        There is no Integration Services heartbeat for an HCS compute system to
        replace it with. Liveness is covered by Test-VmLogStaleness, which
        watches the VM log's write time, and by the hcsdiag instance state in
        Test-HcsStateHealth.
    #>
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
    # Discovery selects the newest VM-relevant log. Preferring coworkd.log, as
    # this used to, meant measuring the staleness of a file last written in
    # March: permanently stale, and the only thing standing between that and a
    # constant repair loop was the VmLogEverActive gate below.
    #
    # main.log is deliberately NOT a candidate. It is the Electron main process
    # log and it can sit frozen for a day at a time on a completely healthy
    # machine, which makes staleness there a guaranteed false positive.
    $vmLogFile = $script:ClaudeEnv.VmLogFile
    if (-not $vmLogFile -or -not (Test-Path $vmLogFile)) {
        $script:VmLogStaleCount = 0
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
        $lastWrite = (Get-Item $vmLogFile -ErrorAction Stop).LastWriteTime
        $staleSec = ((Get-Date) - $lastWrite).TotalSeconds

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
    } catch { $null = $_ }
    return $null
}

function Test-TimeSyncHealth {
    <#
    .SYNOPSIS
        Checks for significant host clock drift. Self-repairs via NTP resync.
        Only checks every 5 minutes. Does NOT trigger auto-fix.
    #>
    $now = Get-Date
    if (($now - $script:LastTimeSyncCheck).TotalMinutes -lt 5) { return $null }
    $script:LastTimeSyncCheck = $now

    try {
        $w32svc = Get-Service -Name "W32Time" -ErrorAction SilentlyContinue
        if ($w32svc -and $w32svc.Status -ne "Running") {
            try {
                Start-Service -Name "W32Time" -ErrorAction Stop
                Write-WatchLog "REPAIRED: Started W32Time service"
            } catch {
                return "W32Time service stopped, clock may drift"
            }
        }

        # /stripchart makes a network call, and this runs inside the monitor's
        # single-threaded poll loop. Unbounded, a blocked or slow NTP server
        # stalls the entire monitor for as long as the network takes to give up.
        $w32tmText = ""
        $w32tmJob = Start-Job -ScriptBlock {
            & w32tm /stripchart /computer:time.windows.com /dataonly /samples:1 2>&1 | Out-String
        }
        if (Wait-Job $w32tmJob -Timeout 15) {
            $w32tmText = (@(Receive-Job $w32tmJob) -join "`n")
        } else {
            Stop-Job $w32tmJob -ErrorAction SilentlyContinue
            Write-WatchLog "w32tm stripchart timed out after 15s; skipping the clock-drift check"
        }
        Remove-Job $w32tmJob -Force -ErrorAction SilentlyContinue

        # Matched against a single joined string, not the raw result.
        #
        # w32tm returns an ARRAY of output lines, and -match against an array
        # uses FILTER semantics: it returns the matching elements and never
        # populates $Matches. Reading $Matches[1] on the next line therefore
        # threw under StrictMode, the empty catch below swallowed it, and
        # neither the drift check nor the NTP resync has ever run.
        if ($w32tmText -match "(-?\d+\.\d+)s") {
            $drift = [math]::Abs([double]$Matches[1])
            if ($drift -gt 5.0) {
                try {
                    & w32tm /resync /force 2>&1 | Out-Null
                    Write-WatchLog "REPAIRED: Forced NTP resync (drift was ${drift}s)"
                } catch { $null = $_ }
                if ($drift -gt 30.0) {
                    return "Clock drift ${drift}s, may cause VM connectivity issues"
                }
            }
        }
    } catch { $null = $_ }
    return $null
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
    $now = Get-Date
    if (($now - $script:LastVmcomputeCheck).TotalSeconds -lt 60) { return $null }
    $script:LastVmcomputeCheck = $now

    try {
        # @()-wrapped. Get-Process returns a scalar for a single match, and
        # .HandleCount on a scalar is fine, but the collection form is what the
        # rest of this file assumes and it also survives zero matches.
        $vmcomputeProcs = @(Get-Process -Name "vmcompute" -ErrorAction SilentlyContinue)
        if ($vmcomputeProcs.Count -eq 0) { return $null }

        $handleCount = 0
        foreach ($vp in $vmcomputeProcs) {
            try { if ($vp.HandleCount -gt $handleCount) { $handleCount = $vp.HandleCount } }
            catch { $null = $_ }
        }

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
    } catch { $null = $_ }
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
    } catch { $null = $_ }

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
                StartTime = (Get-Date).AddHours(-1)
            }
            $allShutdownEvents = @(Get-WinEvent -FilterHashtable $shutdownFilter `
                    -MaxEvents 500 -ErrorAction SilentlyContinue |
                Where-Object { $_.Message -match "0xC037010D" })
            $shutdownCount = $allShutdownEvents.Count
            if ($shutdownCount -gt 0) {
                $shutdownLogElapsed = ((Get-Date) - $script:LastShutdownFailLogTime).TotalMinutes
                if ($shutdownLogElapsed -ge 10) {
                    Write-WatchLog "HCS: $shutdownCount shutdown failures (0xC037010D) in last hour (property query bug, informational)"
                    $script:LastShutdownFailLogTime = Get-Date
                }
            } else {
                # Reset rate-limit timer when count drops to zero
                $script:LastShutdownFailLogTime = [datetime]::MinValue
            }
            if ($shutdownCount -gt 10) {
                $issues += "HCS shutdown failure spike: $shutdownCount events in last hour"
            }
            # Spike detection: excessive shutdown failures may indicate real corruption
            $fiveMinAgo = (Get-Date).AddMinutes(-5)
            $recentSpike = @($allShutdownEvents | Where-Object { $_.TimeCreated -gt $fiveMinAgo })
            if ($recentSpike.Count -gt 15) {
                # Before restarting vmcompute, check if there is an active session
                $coworkSvc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
                $claudeRunning = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue).Count -gt 0
                # Same two lock names Fix actually takes. This was still
                # pinned to "Global\ClaudeDesktopFix_v4.8", so it never found
                # the lock and would have restarted vmcompute out from under a
                # running repair.
                $fixRunning = $false
                foreach ($fixLockName in @("Global\ClaudeDesktopFix", "Local\ClaudeDesktopFix")) {
                    try {
                        $probe = [System.Threading.Mutex]::OpenExisting($fixLockName)
                        $probe.Dispose()
                        $fixRunning = $true
                        break
                    } catch { $null = $_ }
                }
                if ($claudeRunning -and $coworkSvc -and $coworkSvc.Status -eq "Running") {
                    Write-WatchLog "HCS SPIKE: $($recentSpike.Count) failures in 5 min but active session detected, skipping vmcompute restart"
                } elseif ($fixRunning) {
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
    } catch { $null = $_ }

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

function Test-GuestConnectionHealth {
    <#
    .SYNOPSIS
        Monitors cowork-service.log for isGuestConnected RPC timeout patterns.
        Returns $null if healthy, or a trigger string if guest connection is failing.
    #>
    # Skip during startup grace period
    if (Test-StartupGracePeriod) { return $null }

    # %ProgramData% is gone. Discovery's log directory is the only candidate.
    $svcLogPath = $null
    if ($ClaudeLogDir) {
        $candidate = Join-Path $ClaudeLogDir "cowork-service.log"
        if (Test-Path $candidate) { $svcLogPath = $candidate }
    }
    if (-not $svcLogPath) { return $null }

    # Reject a source that predates the window we are about to ask about.
    #
    # cowork-service.log has not been written since 2026-03-18 on current
    # builds. Parsing a file months out of date to answer "what happened in the
    # last 60 seconds" cannot produce a useful answer, and the only reason it
    # produced a harmless one before was that the timestamp regexes below never
    # matched anything either.
    try {
        $svcLogAge = ((Get-Date) - (Get-Item $svcLogPath -ErrorAction Stop).LastWriteTime).TotalSeconds
        if ($svcLogAge -gt 300) { return $null }
    } catch {
        $null = $_
        return $null
    }

    try {
        $lines = Get-Content $svcLogPath -Tail 100 -ErrorAction Stop
        $now = Get-Date
        $recentLines = @()
        foreach ($line in $lines) {
            # Two real formats, both anchored to a full date.
            #
            #   2026/03/18 12:39:04.457195   slashes, SIX fractional digits
            #   2026-07-28 16:27:55          dashes, no fractional part
            #
            # The patterns this replaces expected "yyyy-MM-dd HH:mm:ss.fff" and
            # a bare "HH:mm:ss.fff". Neither matches either real format, so this
            # whole check was dead.
            #
            # The bare time-only branch is also gone on purpose. It rebased a
            # time onto today's date, so a line from months ago read as minutes
            # old. That was a false-positive generator waiting for the regex
            # above it to start matching.
            $stamp = $null
            if ($line -match '^\s*(\d{4})[/-](\d{2})[/-](\d{2})\s+(\d{2}):(\d{2}):(\d{2})(?:\.(\d+))?') {
                try {
                    $ms = 0
                    if ($Matches[7]) {
                        # Truncate to milliseconds; the source emits microseconds.
                        $ms = [int]($Matches[7].PadRight(3, '0').Substring(0, 3))
                    }
                    $stamp = New-Object System.DateTime(
                        [int]$Matches[1], [int]$Matches[2], [int]$Matches[3],
                        [int]$Matches[4], [int]$Matches[5], [int]$Matches[6], $ms)
                } catch { $null = $_ }
            }
            if ($stamp -and ($now - $stamp).TotalSeconds -le 60) {
                $recentLines += $line
            }
        }

        $polls = @($recentLines | Where-Object { $_ -match "method=isGuestConnected" })
        $responses = @($recentLines | Where-Object { $_ -match "Sent response" })

        if ($polls.Count -gt 10 -and $responses.Count -eq 0) {
            return "Guest connection timeout: $($polls.Count) polls, 0 responses in 60s"
        }
    } catch { $null = $_ }
    return $null
}

# -- VM maintenance functions ---------------------------------------------

function Set-VmWorkerPriority {
    # SupportsShouldProcess: this genuinely changes system state, raising the
    # scheduling priority of a live process.
    [CmdletBinding(SupportsShouldProcess)]
    param()
    try {
        $vmwpProcs = @(Get-Process -Name "vmwp" -ErrorAction SilentlyContinue)
        foreach ($p in $vmwpProcs) {
            try {
                if ($p.PriorityClass -ne 'AboveNormal') {
                    if ($PSCmdlet.ShouldProcess("vmwp.exe PID $($p.Id)", "Set priority to AboveNormal")) {
                        $p.PriorityClass = 'AboveNormal'
                        Write-WatchLog "Boosted vmwp.exe (PID $($p.Id)) to AboveNormal"
                    }
                }
            } catch { $null = $_ }
        }
    } catch { $null = $_ }
}

function Set-DynamicMemoryFlag {
    <#
    .SYNOPSIS
        Applies a pending "disable Dynamic Memory" request for a VMMS virtual
        machine, if one was left behind.
    .DESCRIPTION
        Renamed from Apply-DynamicMemoryFlag: Apply is not an approved
        PowerShell verb, which is the one thing the analyzer flagged here.

        Note this is inert for the Cowork VM itself. Get-VM and Set-VMMemory
        operate on VMMS virtual machines, and the Cowork VM is an HCS compute
        system that does not appear to either. The flag file is only ever
        actioned for a VM that VMMS can actually see, and nothing in this
        toolkit writes one.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param()
    $flagFile = Join-Path $ClaudeAppData "disable-dynamic-memory.flag"
    if (-not (Test-Path $flagFile)) { return }

    try {
        $vmName = (Get-Content $flagFile -Raw -ErrorAction Stop).Trim()
        if (-not $vmName) { return }

        $vm = Get-VM -VMName $vmName -ErrorAction SilentlyContinue
        if (-not $vm) { return }

        if ($vm.State -eq "Off") {
            Set-VMMemory -VMName $vmName -DynamicMemoryEnabled $false -ErrorAction Stop
            Remove-Item $flagFile -Force -ErrorAction SilentlyContinue
            Write-WatchLog "Dynamic Memory disabled for VM '$vmName' (flag applied)"
        }
    } catch { $null = $_ }
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
        1. Session files: claude-code-sessions folder for recently-modified files (10 min)
        2. Renderer logs: unknown-window*.log or claude.ai-web*.log recently written (10 min)
           These logs show LOCAL_SESSION activity when Code is streaming/processing.
    #>
    try {
        # Check 1: Claude Code session persistence files
        # Code writes session state to: %AppData%\Claude\claude-code-sessions\{accountId}\{orgId}\
        $codeSessionDir = Join-Path $ClaudeAppData "claude-code-sessions"
        if (Test-Path $codeSessionDir) {
            # Find any .json file modified in the last 10 minutes (Code writes as it works)
            $recentFiles = Get-ChildItem -Path $codeSessionDir -Filter "*.json" -Recurse -ErrorAction SilentlyContinue |
                Where-Object { ((Get-Date) - $_.LastWriteTime).TotalMinutes -lt 10 } |
                Select-Object -First 1
            if ($recentFiles) { return $true }
        }

        # Check 2: Renderer log recency
        # When Code is active, the renderer writes LOCAL_SESSION messages to these logs.
        # Even during quiet API-wait periods, messages appear every few minutes.
        $rendererLogs = @(
            Get-ChildItem -Path $ClaudeLogDir -Filter "unknown-window*.log" -ErrorAction SilentlyContinue
            Get-ChildItem -Path $ClaudeLogDir -Filter "claude.ai-web*.log" -ErrorAction SilentlyContinue
        )
        foreach ($rLog in $rendererLogs) {
            if ($rLog -and ((Get-Date) - $rLog.LastWriteTime).TotalMinutes -lt 10) {
                return $true
            }
        }
    } catch {
        # Fails CLOSED. Test-UserActivity calls this to decide whether a repair
        # may proceed, so an unknown answer must not read as "no session".
        Write-WatchLog "Test-ClaudeCodeActive failed ($($_.Exception.Message)); assuming a Code session IS active"
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
        5. Active Claude Code session (session files or renderer logs)
        6. Active Cowork session (created but not yet cleaned up in VM log)

        SESSION 0 (SYSTEM scheduled tasks):
        - Win32 APIs return zero/stale data, so we skip checks 1-2
        - Falls back to VM log + CPU sampling + Code session + Cowork session detection
    #>
    try {
        $claudeProcs = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
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
                    # Electron renderer -> main process: check parent PID
                    try {
                        $parentId = (Get-CimInstance Win32_Process -Filter "ProcessId=$fgPid" -ErrorAction SilentlyContinue).ParentProcessId
                        foreach ($cp in $claudeProcs) {
                            if ($cp.Id -eq $parentId) { return $true }
                        }
                    } catch { $null = $_ }
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
        # Discovery selects the newest VM-relevant log. This chain used to
        # prefer coworkd.log, which on current builds was last written in
        # March, so the freshness test below was reading a dead file.
        $vmLog = $script:ClaudeEnv.VmLogFile
        if ($vmLog -and (Test-Path $vmLog)) {
            $ageSec = ((Get-Date) - (Get-Item $vmLog).LastWriteTime).TotalSeconds
            if ($ageSec -lt 120) {
                return $true  # VM was active recently -- Code may be thinking
            }
        }

        # Check 4: CPU sampling -- any Claude process using >100ms CPU in 500ms.
        #
        # ONE 500ms sleep covering all processes, not one sleep per process.
        # There are sixteen Electron processes on this machine, so the old
        # per-process loop could sleep for eight seconds inside a thirty second
        # poll, and the cancel path ran the same shape again for another 48.
        $cpuBefore = @{}
        foreach ($cp in $claudeProcs) {
            try { $cpuBefore[$cp.Id] = $cp.TotalProcessorTime.TotalMilliseconds }
            catch { $null = $_ }
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
                } catch { $null = $_ }
            }
        }

        # Check 5: Active Claude Code session
        # Code runs inside Claude Desktop but doesn't use the VM.
        # During API-wait periods, CPU can be near zero for minutes.
        # Session files and renderer logs are more reliable indicators.
        if (Test-ClaudeCodeActive) { return $true }

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
        # catch is why nobody knew this could happen.
        Write-WatchLog "Test-UserActivity failed ($($_.Exception.Message)); assuming the user IS active"
        return $true
    }
    return $false
}

function Test-CoworkSessionActive {
    <#
    .SYNOPSIS
        Returns $true when the Cowork VM has logged activity recently.

        Reads the newest VM log and looks for a recent [Keepalive], [startVM],
        [VM:start], [postConnect] or [vmOneShot] entry within the last 180
        seconds.

        During API-wait periods, when the model is thinking server side, a
        Cowork session burns no local CPU but is very much alive. This check is
        what stops Watch killing Claude in those gaps.

        It previously paired [Process:UUID] Created against Cleaned up. Neither
        marker appears in the current log format at all.
    #>
    # Discovery selects the newest VM-relevant log rather than preferring
    # coworkd.log, which is months stale on current builds.
    $vmLogFile = $script:ClaudeEnv.VmLogFile
    if (-not $vmLogFile -or -not (Test-Path $vmLogFile)) { return $false }
    try {
        # Only read the last 200 lines to stay fast
        $lines = @(Get-Content $vmLogFile -Tail 200 -ErrorAction SilentlyContinue)
        if ($lines.Count -eq 0) { return $false }

        # Liveness comes from recent VM activity tags, not from pairing
        # [Process:UUID] Created against Cleaned up.
        #
        # Those two markers occur ZERO times in the live VM log. They belong to
        # the March-era coworkd.log format. Reading that dead file, the last 200
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
        # Unparseable log. Report no active session, which leaves the other
        # guards in Test-UserActive to decide.
        $null = $_
        return $false
    }
    return $false
}

# -- Auto-fix function ---------------------------------------------------

function Invoke-AutoFix {
    param([string]$Reason)

    # Do not trigger auto-fix if Fix-ClaudeDesktop is already running.
    #
    # Fix takes one of these two names, preferring Global and falling back to
    # Local when it lacks SeCreateGlobalPrivilege. Both are checked here because
    # missing the lock means starting a second repair on top of a running one.
    #
    # The name carries no version deliberately: two toolkit versions operating
    # on the same service and the same VHDX files at once is exactly what the
    # lock exists to prevent. This check used to look for
    # "Global\ClaudeDesktopFix_v4.8", which stopped matching the moment Fix's
    # own name moved on.
    foreach ($fixLockName in @("Global\ClaudeDesktopFix", "Local\ClaudeDesktopFix")) {
        try {
            $existingLock = [System.Threading.Mutex]::OpenExisting($fixLockName)
            $existingLock.Dispose()
            Write-WatchLog "AUTO-FIX SKIPPED: Fix-ClaudeDesktop is already running"
            return
        } catch {
            # Not held under this scope. Try the next one.
            $null = $_
        }
    }

    $now = Get-Date
    $elapsed = ($now - $script:LastFixTime).TotalMinutes

    if ($elapsed -lt $Cooldown) {
        Write-WatchLog "COOLDOWN: Skipping fix ($Reason), last fix $([math]::Round($elapsed, 1)) min ago"
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
        } catch { $null = $_ }
        # Back off for the 30 minutes that Test-PersistentFailure measures over.
        #
        # This used to add $Cooldown, which defaults to 5, so the "back off"
        # lasted 10 minutes, not 30. Test-PersistentFailure looks at a 30-minute
        # window, so the escalation stayed true and the balloon plus alert sound
        # repeated every 10 minutes indefinitely.
        $script:LastFixTime = (Get-Date).AddMinutes(30 - $Cooldown)
        return
    }

    # ---- SAFETY: Never auto-fix while user is active ----
    if (Test-UserActivity) {
        Write-WatchLog ">>> BLOCKED: $Reason, user is active (Claude in focus, recent input, or VM busy)"
        Write-WatchLog "    Run Fix-ClaudeDesktop.ps1 manually when ready"
        # Set cooldown so we don't spam the log every 30s -- retry in ~2 min
        $script:LastFixTime = $now.AddMinutes(-($Cooldown - 2))
        return
    }

    # ---- SAFETY: Don't auto-fix during Claude startup ----
    # When Claude just launched, workspace initialization can take 30-60 seconds
    # and may show transient errors. Give it time to settle.
    $claudeStartupProcs = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
    if ($claudeStartupProcs.Count -gt 0) {
        # .StartTime is read per process inside a try.
        #
        # It throws for a process that exited between the Get-Process above and
        # this line, and with sixteen short-lived Electron helpers around, that
        # is a routine occurrence rather than an edge case. Unprotected, one
        # exited helper abandoned the whole repair for that cycle.
        $youngest = $null
        foreach ($sp in $claudeStartupProcs) {
            try {
                $st = $sp.StartTime
                if ($null -eq $youngest -or $st -gt $youngest) { $youngest = $st }
            } catch { $null = $_ }
        }
        $procAge = if ($youngest) { ((Get-Date) - $youngest).TotalSeconds } else { [double]::MaxValue }
        if ($procAge -lt 120) {
            Write-WatchLog ">>> DEFERRED: $Reason, Claude launched $([math]::Round($procAge))s ago (waiting for startup to complete)"
            $script:LastFixTime = $now.AddMinutes(-($Cooldown - 1))
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

    # Dismiss notification
    try { if ($notifyIcon) { $notifyIcon.Visible = $false; $notifyIcon.Dispose() } } catch { $null = $_ }

    # Re-check: only cancel if Claude is ACTIVELY being used.
    # We check Claude-specific signals (foreground, CPU, VM log) but NOT
    # general user input -- the user may be at the keyboard in another app
    # while Claude's VM is genuinely dead.
    $cancelFix = $false
    $cancelReason = ""
    $claudeProcs = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
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
                        if (-not $cancelFix) {
                            try {
                                $parentId = (Get-CimInstance Win32_Process -Filter "ProcessId=$fgPid" -ErrorAction SilentlyContinue).ParentProcessId
                                foreach ($cp in $claudeProcs) {
                                    if ($cp.Id -eq $parentId) { $cancelFix = $true; $cancelReason = "Claude is now in focus"; break }
                                }
                            } catch { $null = $_ }
                        }
                    }
                }
            } catch { $null = $_ }
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
                    } catch { $null = $_ }
                }
            }
        }
    }

    # Cancel if VM log became active during the 30s window
    if (-not $cancelFix) {
        # Discovery selects the newest VM-relevant log. This chain used to
        # prefer coworkd.log, which on current builds was last written in
        # March, so the freshness test below was reading a dead file.
        $vmLog = $script:ClaudeEnv.VmLogFile
        if ($vmLog -and (Test-Path $vmLog)) {
            $ageSec = ((Get-Date) - (Get-Item $vmLog).LastWriteTime).TotalSeconds
            if ($ageSec -lt 120) { $cancelFix = $true; $cancelReason = "VM log became active" }
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
        } catch { $null = $_ }

        $script:LastFixTime = $now.AddMinutes(-($Cooldown - 2))
        return
    }

    $script:FixCount++
    Write-WatchLog ">>> AUTO-FIX #$($script:FixCount) TRIGGERED: $Reason"
    Write-WatchLog "    (Claude still inactive after 30s warning)"
    $script:LastFixTime = $now

    # Fix-ClaudeDesktop self-elevates with Start-Process -Verb RunAs -Wait when
    # it is not already running as admin. Invoking it in-process from an
    # unattended monitor therefore either raises a UAC consent dialog with
    # nobody there to answer it, blocking this monitor for as long as it sits
    # on screen, or fails outright. Neither is a repair.
    #
    # So only call it when this process already holds the token Fix needs. The
    # cooldown was set just above, so this logs at most once per cooldown
    # window rather than on every poll.
    $watchIsAdmin = $false
    try {
        $watchIsAdmin = ([Security.Principal.WindowsPrincipal] `
            [Security.Principal.WindowsIdentity]::GetCurrent()
            ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { $null = $_ }

    if (-not $watchIsAdmin) {
        Write-WatchLog ">>> AUTO-FIX #$($script:FixCount) SKIPPED: this monitor is not elevated, so Fix would raise a UAC prompt with nobody to answer it. Reinstall the monitor task with highest privileges, or run Fix-ClaudeDesktop manually."
        return
    }

    try {
        & $fixScript -Mode Smart -Quiet
        Write-WatchLog ">>> AUTO-FIX #$($script:FixCount) COMPLETE (exit $LASTEXITCODE)"
    } catch {
        Write-WatchLog ">>> AUTO-FIX #$($script:FixCount) FAILED: $($_.Exception.Message)"
    }

    # Record fix timestamp for persistent failure detection, and persist it so
    # the backoff survives the restart that a bad repair loop tends to cause.
    $null = $script:FixHistory.Add((Get-Date))
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
    $script:HeartbeatFailCount = 0
    $script:VmcomputeLeakCount = 0
    $script:GuestFailCount     = 0
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
} catch {
    # Could not take the lock. Monitoring without one is better than not
    # monitoring, and Invoke-AutoFix checks Fix's own lock separately before it
    # does anything destructive.
    $null = $_
}

# Restore cooldown and fix history before the first poll, so a restart cannot
# be used to escape the backoff.
Import-WatchState

# -- Main loop -----------------------------------------------------------
Write-WatchLog "================================================================"
Write-WatchLog "Health Monitor v$ToolkitVersion started"
Write-WatchLog "  Poll interval : ${PollInterval}s"
Write-WatchLog "  Cooldown      : ${Cooldown} min"
Write-WatchLog "  Fix script    : $fixScript"
Write-WatchLog "  Log directory : $ClaudeLogDir"
Write-WatchLog "  VM log        : $($script:ClaudeEnv.VmLogFile)"
Write-WatchLog "  Safety        : user-activity block, 180s grace period, consecutive-check gates, cowork-session awareness"
# Lists only monitors that can actually fire.
#
# The previous line advertised heartbeat(x3), guest-connect and staleness.
# Heartbeat could never fire, because Get-VM does not see an HCS compute
# system. Guest-connect could never fire, because its timestamp patterns did
# not match either real log format. Staleness could never fire, because it
# measured a file last written in March. Naming dead monitors in the startup
# banner is how they stayed unnoticed.
# Guest-connect is reported as live or dormant based on the actual source file,
# not asserted. Test-GuestConnectionHealth rejects cowork-service.log when it is
# more than 300s stale, and on current builds that file has not been written
# since March, so the check returns immediately. Claiming it as an active
# monitor would repeat exactly the mistake this banner used to make.
$guestSrc = $null
if ($ClaudeLogDir) { $guestSrc = Join-Path $ClaudeLogDir "cowork-service.log" }
$guestLive = $false
if ($guestSrc -and (Test-Path $guestSrc)) {
    try {
        $guestLive = (((Get-Date) - (Get-Item $guestSrc).LastWriteTime).TotalSeconds -le 300)
    } catch { $null = $_ }
}

$autoFixList = "logs, service(x2), events(x2)+HCS, vmcompute-health"
if ($guestLive) { $autoFixList += ", guest-connect(x3)" }

Write-WatchLog "  Auto-fix on   : $autoFixList"
Write-WatchLog "  Warn only     : NAT, storage, VM log staleness(x5), clock drift, hcs-state"
Write-WatchLog "  Not monitored : Hyper-V heartbeat (no such thing for an HCS compute system)"
if (-not $guestLive) {
    Write-WatchLog "  Dormant       : guest-connect (cowork-service.log is stale or absent on this build)"
}
Write-WatchLog "================================================================"

try {
    while ($true) {
        try {
            Update-LogFile

            $claudeRunning = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue).Count -gt 0

            # Record Claude disappearing. Reported, never acted on.
            #
            # Every check below sits inside "if ($claudeRunning)", so when
            # Claude is gone this monitor previously did nothing at all: no
            # detection, no note, nothing. A user report describes exactly that
            # shape, an app that died and stayed dead for over an hour with
            # zero log activity to point at afterwards.
            #
            # This monitor cannot tell a crash from someone quitting, and
            # relaunching an app the user closed on purpose would be worse than
            # useless, so it does not try. What it can do is leave a timestamp,
            # which is the thing that was missing.
            #
            # Whether the service is still up is recorded alongside it: Close
            # mode stops the service, so a live service with no Claude leans
            # towards an unexpected exit rather than a clean shutdown.
            if ($claudeRunning -ne $script:ClaudeWasRunning) {
                if (-not $claudeRunning) {
                    $svcNow = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
                    $svcState = if ($svcNow) { "$($svcNow.Status)" } else { "absent" }
                    Write-WatchLog "Claude is no longer running (service: $svcState). Not acting: a crash and a deliberate quit look the same from here."
                } else {
                    Write-WatchLog "Claude is running again"
                }
                $script:ClaudeWasRunning = $claudeRunning
            }

            if ($claudeRunning) {
                # ---- Critical checks (trigger auto-fix) ----

                # Check 1: VirtioFS errors in log files (most reliable -- no consecutive gate needed)
                if (-not (Test-StartupGracePeriod)) {
                    $logError = Test-LogsForErrors
                    if ($logError) {
                        Invoke-AutoFix -Reason "Log error: $logError"
                        continue
                    }
                }

                # Check 2: Service died while Claude is running (2 consecutive checks)
                if (-not (Test-ServiceHealth)) {
                    Invoke-AutoFix -Reason "CoworkVMService stopped while Claude is running (confirmed)"
                    continue
                }

                # Check 3: Event Log errors (2 consecutive checks, grace period, tightened filters)
                $eventError = Test-EventLogErrors
                if ($eventError) {
                    Invoke-AutoFix -Reason "Event Log: $eventError"
                    continue
                }

                # Check 4: WinNAT connectivity (warning only -- does NOT auto-fix)
                $natIssue = Test-WinNatHealth
                if ($natIssue) {
                    Write-WatchLog "NAT WARNING: $natIssue"
                }

                # Check 4b: workspace storage readable and not out of room.
                # Warning only: a restart does not create disk space.
                $storageIssue = Test-WorkspaceStorage
                if ($storageIssue) {
                    Write-WatchLog "STORAGE WARNING: $storageIssue"
                }

                # Check 5: Hyper-V heartbeat (3 consecutive checks, grace period)
                $hbIssue = Test-VmHeartbeat
                if ($hbIssue) {
                    Invoke-AutoFix -Reason $hbIssue
                    continue
                }

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
                    Invoke-AutoFix -Reason $vmcomputeIssue
                    continue
                }

                # Check 8: HCS state health (stale VMs, shutdown failure frequency)
                $hcsStateIssue = Test-HcsStateHealth
                if ($hcsStateIssue) {
                    # Rate-limit HCS STATE logging to once per 10 minutes (was every 30s)
                    $hcsLogElapsed = ((Get-Date) - $script:LastHcsStateLogTime).TotalMinutes
                    if ($hcsLogElapsed -ge 10) {
                        Write-WatchLog "HCS STATE: $hcsStateIssue"
                        $script:LastHcsStateLogTime = Get-Date
                        # Informational plus a pre-emptive cleanup; no auto-fix.
                        #
                        # Rewritten. The block this replaces carried the same two
                        # bugs as Fix-ClaudeDesktop findings 1 and 2, and the
                        # audit did not catch this second copy of them:
                        #
                        #   1. It parsed GUIDs by looking for a line matching
                        #      ^\s*GUID\s*$. "hcsdiag list" never puts the GUID
                        #      on a line of its own, so $vmGuids was always
                        #      empty and this entire block was dead.
                        #   2. It then called "hcsdiag close", and hcsdiag has
                        #      no close verb. Even had the parse worked, the
                        #      call would have done nothing while logging
                        #      "proactively closed stale cowork-vm".
                        #
                        # The parse now lives in Get-CoworkHcsGuids in the shared
                        # ClaudeEnv region. This was the third of four copies.
                        try {
                            $hcsList = Get-HcsListCached
                            if ($hcsList) {
                                $vmGuids = @(Get-CoworkHcsGuids -ListOutput ([string]$hcsList))

                                # Keep the newest entry, which is the live VM,
                                # and kill anything else left behind.
                                if ($vmGuids.Count -gt 1) {
                                    foreach ($guid in $vmGuids[0..($vmGuids.Count - 2)]) {
                                        Invoke-HcsDiag -Arguments "kill",$guid | Out-Null
                                        Write-WatchLog "HCS: killed stale cowork-vm ($guid)"
                                    }
                                }
                            }
                        } catch { $null = $_ }
                    }
                } else {
                    # Reset timer when healthy so next issue logs immediately
                    if ($script:LastHcsStateLogTime -ne [datetime]::MinValue) {
                        $script:LastHcsStateLogTime = [datetime]::MinValue
                    }
                }

                # Check 10: Guest connection health (cowork-service.log)
                $guestHealth = Test-GuestConnectionHealth
                if ($guestHealth) {
                    $script:GuestFailCount++
                    Write-WatchLog "CHECK 10 FAIL ($($script:GuestFailCount)/3): $guestHealth"
                    if ($script:GuestFailCount -ge 3) {
                        Invoke-AutoFix -Reason "guest-connect-timeout"
                        $script:GuestFailCount = 0
                        continue
                    }
                } else {
                    $script:GuestFailCount = 0
                }

                # ---- Maintenance (non-fix actions) ----
                Set-VmWorkerPriority

                $timeIssue = Test-TimeSyncHealth
                if ($timeIssue) {
                    Write-WatchLog "TIME WARNING: $timeIssue"
                }
            }

            Set-DynamicMemoryFlag
        } catch {
            Write-WatchLog "MONITOR ERROR: $($_.Exception.Message)"
        } finally {
            # The poll sleep lives in a finally, and that is the whole point.
            #
            # Seven of the checks above end in "continue", which targets the
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
        } catch { $null = $_ }
    }
    Write-WatchLog "Health monitor stopped"
}
