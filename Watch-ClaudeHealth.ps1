<#
.SYNOPSIS
    Claude Desktop / Cowork -- Health Monitor

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
    - Host clock drift (NTP/time sync)

    SAFETY FEATURES (v3.3):
    - Auto-fix is BLOCKED when user is active (Claude in focus, recent input, VM busy, CPU active)
    - Claude Code session awareness (detects active Code sessions via session files + renderer logs)
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
    Version : 4.8.0
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

# -- Constants -----------------------------------------------------------
$Version        = "4.8.0"
$ServiceName    = "CoworkVMService"
$ClaudeAppData  = Join-Path $env:APPDATA "Claude"
$ClaudeLogDir   = Join-Path $ClaudeAppData "logs"
$WatchLogDir    = Join-Path $ClaudeAppData "watch-logs"

# Error patterns that indicate VirtioFS mount failure
$ErrorPatterns = @(
    "Plan9 mount failed",
    "bad address",
    "failed to ensure virtiofs mount",
    "RPC error -1",
    "HCS operation failed",
    "failed to create compute system",
    "HcsWaitForOperationResult",
    "VM is already running"
)

# -- Win32 APIs for user activity detection ------------------------------
# NOTE: These APIs only work in interactive sessions (Session 1+).
# In Session 0 (SYSTEM scheduled tasks), they return zero/stale data.
# We detect this and fall back to process-only heuristics.
$script:IsInteractiveSession = $false
try {
    $sessionId = (Get-Process -Id $PID -ErrorAction Stop).SessionId
    $script:IsInteractiveSession = ($sessionId -gt 0)
} catch {}

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

if (-not (Test-Path $fixScript)) {
    Write-Host "  [!] Fix-ClaudeDesktop.ps1 not found. Health monitor cannot start." -ForegroundColor Red
    exit 1
}

# -- Logging -------------------------------------------------------------
if (-not (Test-Path $WatchLogDir)) { New-Item $WatchLogDir -ItemType Directory -Force | Out-Null }

# Clean old watch logs (>30 days)
try {
    Get-ChildItem $WatchLogDir -Filter "watch_*.log" -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
        Remove-Item -Force -ErrorAction SilentlyContinue
} catch {}

$script:WatchLogFile = Join-Path $WatchLogDir ("watch_{0:yyyyMMdd}.log" -f (Get-Date))

function Write-WatchLog {
    param([string]$Message)
    $line = "[{0:yyyy-MM-dd HH:mm:ss}] $Message" -f (Get-Date)
    try { $line | Out-File -FilePath $script:WatchLogFile -Append -Encoding utf8 -ErrorAction SilentlyContinue } catch {}
    if (-not $Quiet) { Write-Host $line -ForegroundColor DarkGray }
}

# Rotate to a new daily log file if the date rolls over
function Update-LogFile {
    $newPath = Join-Path $WatchLogDir ("watch_{0:yyyyMMdd}.log" -f (Get-Date))
    if ($newPath -ne $script:WatchLogFile) {
        $script:WatchLogFile = $newPath
    }
}

# -- State ---------------------------------------------------------------
$script:StartTime            = Get-Date
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
$script:FixHistory           = New-Object System.Collections.ArrayList

# Initialize baselines for existing log files
if (Test-Path $ClaudeLogDir) {
    Get-ChildItem $ClaudeLogDir -Filter "*.log" -ErrorAction SilentlyContinue | ForEach-Object {
        $script:LogBaselines[$_.FullName] = $_.Length
    }
}

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
    foreach ($logFile in $logFiles) {
        $path = $logFile.FullName
        $currentSize = $logFile.Length

        $baseline = 0
        if ($script:LogBaselines.ContainsKey($path)) {
            $baseline = $script:LogBaselines[$path]
        }

        # Update baseline
        $script:LogBaselines[$path] = $currentSize

        # Handle log rotation (file got smaller = was recreated)
        if ($currentSize -lt $baseline) {
            $baseline = 0
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
                if ($reader) { try { $reader.Close() } catch {} }
                if ($stream) { try { $stream.Close() } catch {} }
            }

            foreach ($pattern in $ErrorPatterns) {
                if ($newContent -match [regex]::Escape($pattern)) {
                    return $pattern
                }
            }
        } catch {}
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
    $claude = @(Get-Process -Name "claude" -ErrorAction SilentlyContinue)
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
        Write-WatchLog "Service not running (check $($script:ServiceDownCount)/2 -- waiting to confirm)"
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
            $msg = ($events[0].Message -split "`n")[0]
            foreach ($pattern in $ErrorPatterns) {
                if ($msg -match [regex]::Escape($pattern)) {
                    $foundIssue = "CoworkVMService: $msg"
                    break
                }
            }
        }
    } catch {}

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
                $msg = ($hvEvents[0].Message -split "`n")[0]
                if ($msg -match "claude" -or $msg -match "Plan9" -or $msg -match "virtio" -or $msg -match "shared memory") {
                    $foundIssue = "Hyper-V Worker: $msg"
                }
            }
        } catch {}
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
                $msg = ($vmmsEvents[0].Message -split "`n")[0]
                # TIGHTENED: must specifically mention claude or cowork
                if ($msg -match "claude" -or $msg -match "cowork") {
                    $foundIssue = "Hyper-V VMMS: $msg"
                }
            }
        } catch {}
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
                $msg = ($hcsEvents[0].Message -split "`n")[0]
                if ($msg -match "claude" -or $msg -match "cowork" -or $msg -match "failed to create" -or $msg -match "HcsWaitForOperationResult") {
                    $foundIssue = "HCS Compute: $msg"
                }
            }
        } catch {}
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
            $hvSwitch = Get-VMSwitch -SwitchType Internal -ErrorAction SilentlyContinue |
                        Where-Object { $_.Name -match "WSL|claude|nat" } |
                        Select-Object -First 1
            if (-not $hvSwitch) {
                $hvSwitch = Get-VMSwitch -SwitchType Internal -ErrorAction SilentlyContinue |
                            Select-Object -First 1
            }

            if ($hvSwitch) {
                $adapter = Get-NetAdapter -ErrorAction SilentlyContinue |
                           Where-Object { $_.InterfaceDescription -match "Hyper-V Virtual Ethernet Adapter" -and $_.Name -match $hvSwitch.Name } |
                           Select-Object -First 1
                if (-not $adapter) {
                    $adapter = Get-NetAdapter -Name "vEthernet ($($hvSwitch.Name))" -ErrorAction SilentlyContinue
                }

                if ($adapter) {
                    $ipAddr = Get-NetIPAddress -InterfaceIndex $adapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                              Select-Object -First 1
                    if ($ipAddr) {
                        $prefix = ($ipAddr.IPAddress -split '\.')[0..2] -join '.'
                        $subnet = "$prefix.0/24"
                        try {
                            New-NetNat -Name "CoworkNAT" -InternalIPInterfaceAddressPrefix $subnet -ErrorAction Stop | Out-Null
                            Write-WatchLog "REPAIRED: Created WinNAT rule 'CoworkNAT' for $subnet"
                            return $null
                        } catch {
                            return "WinNAT missing and auto-repair failed: $($_.Exception.Message)"
                        }
                    }
                }
            }
            return "No WinNAT rules found -- VM may have no network connectivity"
        }
    } catch {
        # Get-NetNat not available -- skip this check silently
    }
    return $null
}

function Test-VmHeartbeat {
    <#
    .SYNOPSIS
        Checks the Hyper-V Integration Services heartbeat for the Claude VM.
        Now requires 3 CONSECUTIVE failures before triggering (was instant).
        Returns $null if healthy or unavailable, or a description if unhealthy.
    #>
    # Skip during startup grace period
    if (Test-StartupGracePeriod) { return $null }

    try {
        $claudeVm = Get-VM -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match "claude" } |
                    Select-Object -First 1
        if (-not $claudeVm) { return $null }
        if ($claudeVm.State -ne "Running") { return $null }

        $hb = Get-VMIntegrationService -VMName $claudeVm.Name -Name "Heartbeat" -ErrorAction SilentlyContinue
        if (-not $hb) { return $null }
        if (-not $hb.Enabled) { return $null }

        $status = $hb.PrimaryStatusDescription
        if ($status -eq "OK" -or $status -eq "No Contact") {
            if ($script:LastHeartbeatStatus -eq "Lost" -and $status -eq "OK") {
                Write-WatchLog "Heartbeat recovered for VM '$($claudeVm.Name)'"
            }
            $script:LastHeartbeatStatus = $status
            $script:HeartbeatFailCount = 0
            return $null
        }

        # Heartbeat is not OK
        $script:HeartbeatFailCount++
        $script:LastHeartbeatStatus = "Lost"

        if ($script:HeartbeatFailCount -ge 3) {
            $script:HeartbeatFailCount = 0
            return "VM heartbeat: $status for 3 consecutive checks (VM may be hung)"
        }

        Write-WatchLog "Heartbeat issue (check $($script:HeartbeatFailCount)/3): $status"
    } catch {}
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
    $vmLogFile = Join-Path $ClaudeLogDir "cowork_vm_node.log"
    if (-not (Test-Path $vmLogFile)) {
        $script:VmLogStaleCount = 0
        return $null
    }

    # Skip during startup grace period
    if (Test-StartupGracePeriod) { return $null }

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
                return "VM log stale for $([math]::Round($staleSec))s -- VM may be hung"
            }
            if ($script:VmLogStaleCount -eq 1) {
                Write-WatchLog "VM log stale ($([math]::Round($staleSec))s) -- monitoring (check 1/5)"
            }
        } else {
            $script:VmLogStaleCount = 0
        }
    } catch {}
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
                return "W32Time service stopped -- clock may drift"
            }
        }

        $w32tmResult = & w32tm /stripchart /computer:time.windows.com /dataonly /samples:1 2>&1
        if ($w32tmResult -match "(-?\d+\.\d+)s") {
            $drift = [math]::Abs([double]$Matches[1])
            if ($drift -gt 5.0) {
                try {
                    & w32tm /resync /force 2>&1 | Out-Null
                    Write-WatchLog "REPAIRED: Forced NTP resync (drift was ${drift}s)"
                } catch {}
                if ($drift -gt 30.0) {
                    return "Clock drift ${drift}s -- may cause VM connectivity issues"
                }
            }
        }
    } catch {}
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
        $vmcompute = Get-Process -Name "vmcompute" -ErrorAction SilentlyContinue
        if (-not $vmcompute) { return $null }

        $handleCount = $vmcompute.HandleCount
        if ($handleCount -gt 10000) {
            $script:VmcomputeLeakCount++
            if ($script:VmcomputeLeakCount -ge 2) {
                $script:VmcomputeLeakCount = 0
                return "vmcompute handle leak critical ($handleCount handles) -- service needs restart"
            }
            Write-WatchLog "vmcompute handle leak detected ($handleCount handles) -- check $($script:VmcomputeLeakCount)/2"
        } elseif ($handleCount -gt 5000) {
            Write-WatchLog "vmcompute handle leak detected ($handleCount handles) -- service may need restart"
            $script:VmcomputeLeakCount = 0
        } else {
            $script:VmcomputeLeakCount = 0
        }
    } catch {}
    return $null
}

function Test-HcsStateHealth {
    <#
    .SYNOPSIS
        Monitors HCS state for stale cowork-vm and 0xC037010D shutdown failure frequency.
        Returns $null if healthy, or a trigger string if intervention needed.
        Also tracks session file count as secondary metric.
    #>
    # Skip during startup grace period
    if (Test-StartupGracePeriod) { return $null }

    $issues = @()

    # Check 1: Stale cowork-vm in hcsdiag (should not exist while service is healthy)
    try {
        $hcsdiagPath = "$env:SystemRoot\System32\hcsdiag.exe"
        if (Test-Path $hcsdiagPath) {
            $hcsList = & $hcsdiagPath list 2>&1 | Out-String
            # Count cowork-vm entries
            $vmCount = ([regex]::Matches($hcsList, "cowork-vm")).Count
            if ($vmCount -gt 1) {
                $issues += "Multiple cowork-vm instances in HCS ($vmCount found)"
            }
        }
    } catch {}

    # Check 2: 0xC037010D shutdown failure frequency
    try {
        $shutdownFilter = @{
            LogName   = "Microsoft-Windows-Hyper-V-Compute-Operational"
            StartTime = (Get-Date).AddHours(-1)
        }
        $allShutdownEvents = @(Get-WinEvent -FilterHashtable $shutdownFilter -ErrorAction SilentlyContinue |
            Where-Object { $_.Message -match "0xC037010D" })
        $shutdownCount = $allShutdownEvents.Count
        if ($shutdownCount -gt 0) {
            Write-WatchLog "HCS: $shutdownCount shutdown failures (0xC037010D) in last hour"
        }
        if ($shutdownCount -gt 10) {
            $issues += "HCS shutdown failure spike: $shutdownCount events in last hour"
        }
        # Spike detection: >5 in last 5 minutes triggers vmcompute restart
        $fiveMinAgo = (Get-Date).AddMinutes(-5)
        $recentSpike = @($allShutdownEvents | Where-Object { $_.TimeCreated -gt $fiveMinAgo })
        if ($recentSpike.Count -gt 5) {
            Write-WatchLog "HCS SPIKE: $($recentSpike.Count) shutdown failures in 5 min — restarting vmcompute"
            try {
                Restart-Service -Name "vmcompute" -Force -ErrorAction Stop
                Start-Sleep -Seconds 3
                Write-WatchLog "vmcompute restarted (pre-emptive, cleared stale HCS state)"
            } catch {
                Write-WatchLog "vmcompute restart failed: $($_.Exception.Message)"
            }
        }
    } catch {}

    # Check 3: Session file count (secondary metric — warn at 500, critical at 1000)
    try {
        $sessionDir = Join-Path $env:APPDATA "Claude\local-agent-mode-sessions"
        if (Test-Path $sessionDir) {
            $fileCount = (Get-ChildItem $sessionDir -Recurse -File -ErrorAction SilentlyContinue).Count
            if ($fileCount -gt 1000) {
                $issues += "Session file accumulation critical: $fileCount files"
            } elseif ($fileCount -gt 500) {
                Write-WatchLog "Session file count elevated: $fileCount files"
            }
        }
    } catch {}

    if ($issues.Count -gt 0) {
        return ($issues -join "; ")
    }
    return $null
}

# -- VM maintenance functions ---------------------------------------------

function Set-VmWorkerPriority {
    try {
        $vmwpProcs = @(Get-Process -Name "vmwp" -ErrorAction SilentlyContinue)
        foreach ($p in $vmwpProcs) {
            try {
                if ($p.PriorityClass -ne 'AboveNormal') {
                    $p.PriorityClass = 'AboveNormal'
                    Write-WatchLog "Boosted vmwp.exe (PID $($p.Id)) to AboveNormal"
                }
            } catch {}
        }
    } catch {}
}

function Apply-DynamicMemoryFlag {
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
    } catch {}
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
    } catch {}
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

        SESSION 0 (SYSTEM scheduled tasks):
        - Win32 APIs return zero/stale data, so we skip checks 1-2
        - Falls back to VM log + CPU sampling + Code session detection
    #>
    try {
        $claudeProcs = @(Get-Process -Name "claude" -ErrorAction SilentlyContinue)
        if ($claudeProcs.Count -eq 0) { return $false }

        if ($script:IsInteractiveSession) {
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
                    # Electron renderer → main process: check parent PID
                    try {
                        $parentId = (Get-CimInstance Win32_Process -Filter "ProcessId=$fgPid" -ErrorAction SilentlyContinue).ParentProcessId
                        foreach ($cp in $claudeProcs) {
                            if ($cp.Id -eq $parentId) { return $true }
                        }
                    } catch {}
                }
            }

            # Check 2: User input within 3 minutes (extended from 2)
            $lastInput = New-Object Win32Activity+LASTINPUTINFO
            $lastInput.cbSize = [uint32][System.Runtime.InteropServices.Marshal]::SizeOf($lastInput)
            if ([Win32Activity]::GetLastInputInfo([ref]$lastInput)) {
                $idleMs = [Environment]::TickCount - $lastInput.dwTime
                if ($idleMs -lt 180000) {
                    return $true  # User active within 3 min + Claude is running
                }
            }
        }

        # Check 3: VM log active within 120s (was 30s -- covers Code thinking phases)
        $vmLog = Join-Path $ClaudeLogDir "cowork_vm_node.log"
        if (Test-Path $vmLog) {
            $ageSec = ((Get-Date) - (Get-Item $vmLog).LastWriteTime).TotalSeconds
            if ($ageSec -lt 120) {
                return $true  # VM was active recently -- Code may be thinking
            }
        }

        # Check 4: CPU sampling -- any Claude process using >100ms CPU in 500ms
        # Catches active request processing even when UI is idle
        foreach ($cp in $claudeProcs) {
            try {
                $cpu1 = $cp.TotalProcessorTime.TotalMilliseconds
                Start-Sleep -Milliseconds 500
                $cp.Refresh()
                $cpu2 = $cp.TotalProcessorTime.TotalMilliseconds
                if (($cpu2 - $cpu1) -gt 100) { return $true }
            } catch {}
        }

        # Check 5: Active Claude Code session
        # Code runs inside Claude Desktop but doesn't use the VM.
        # During API-wait periods, CPU can be near zero for minutes.
        # Session files and renderer logs are more reliable indicators.
        if (Test-ClaudeCodeActive) { return $true }
    } catch {}
    return $false
}

# -- Auto-fix function ---------------------------------------------------

function Invoke-AutoFix {
    param([string]$Reason)

    $now = Get-Date
    $elapsed = ($now - $script:LastFixTime).TotalMinutes

    if ($elapsed -lt $Cooldown) {
        Write-WatchLog "COOLDOWN: Skipping fix ($Reason) -- last fix $([math]::Round($elapsed, 1)) min ago"
        return
    }

    # Persistent failure escalation
    if (Test-PersistentFailure) {
        Write-WatchLog "ESCALATION: Fix attempted 3+ times in 30 min -- backing off"
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
            $notifyIcon.BalloonTipTitle = "Claude Health Monitor -- Manual Fix Required"
            $notifyIcon.BalloonTipText = "Auto-fix failed 3 times. Hyper-V nuclear reset needed.`nSee watch log for instructions."
            $notifyIcon.ShowBalloonTip(60000)
            [System.Media.SystemSounds]::Hand.Play()
        } catch {}
        # Back off for 30 minutes
        $script:LastFixTime = (Get-Date).AddMinutes($Cooldown)
        return
    }

    # ---- SAFETY: Never auto-fix while user is active ----
    if (Test-UserActivity) {
        Write-WatchLog ">>> BLOCKED: $Reason -- user is active (Claude in focus, recent input, or VM busy)"
        Write-WatchLog "    Run Fix-ClaudeDesktop.ps1 manually when ready"
        # Set cooldown so we don't spam the log every 30s -- retry in ~2 min
        $script:LastFixTime = $now.AddMinutes(-($Cooldown - 2))
        return
    }

    # ---- PRE-FIX WARNING: 30s grace period with notification ----
    # Show a balloon tip so the user knows what's about to happen.
    # The notification says "open Claude to cancel" because we only cancel
    # if Claude becomes actively used -- not just because the mouse moved.
    # This way, a genuinely hung VM still gets fixed even if the user is
    # browsing, gaming, or otherwise using the PC.
    Write-WatchLog ">>> PRE-FIX WARNING: $Reason -- auto-fix in 30s (open Claude to cancel)"
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
        # Notification failed (Session 0, missing assemblies) -- proceed without
    }

    Start-Sleep -Seconds 30

    # Dismiss notification
    try { if ($notifyIcon) { $notifyIcon.Visible = $false; $notifyIcon.Dispose() } } catch {}

    # Re-check: only cancel if Claude is ACTIVELY being used.
    # We check Claude-specific signals (foreground, CPU, VM log) but NOT
    # general user input -- the user may be at the keyboard in another app
    # while Claude's VM is genuinely dead.
    $cancelFix = $false
    $cancelReason = ""
    $claudeProcs = @(Get-Process -Name "claude" -ErrorAction SilentlyContinue)
    if ($claudeProcs.Count -gt 0) {
        # Cancel if Claude is the foreground window
        if ($script:IsInteractiveSession) {
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
                            } catch {}
                        }
                    }
                }
            } catch {}
        }

        # Cancel if any Claude process is burning CPU
        # Extended sampling: 3 checks × 1s to catch bursty Code patterns
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
                    } catch {}
                }
            }
        }
    }

    # Cancel if VM log became active during the 30s window
    if (-not $cancelFix) {
        $vmLog = Join-Path $ClaudeLogDir "cowork_vm_node.log"
        if (Test-Path $vmLog) {
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
        } catch {}

        $script:LastFixTime = $now.AddMinutes(-($Cooldown - 2))
        return
    }

    $script:FixCount++
    Write-WatchLog ">>> AUTO-FIX #$($script:FixCount) TRIGGERED: $Reason"
    Write-WatchLog "    (Claude still inactive after 30s warning)"
    $script:LastFixTime = $now

    try {
        & $fixScript -Mode Smart -Quiet
        Write-WatchLog ">>> AUTO-FIX #$($script:FixCount) COMPLETE"
    } catch {
        Write-WatchLog ">>> AUTO-FIX #$($script:FixCount) FAILED: $($_.Exception.Message)"
    }

    # Record fix timestamp for persistent failure detection
    $null = $script:FixHistory.Add((Get-Date))

    # Reset state after fix
    Start-Sleep -Seconds 10
    $script:LogBaselines = @{}
    if (Test-Path $ClaudeLogDir) {
        Get-ChildItem $ClaudeLogDir -Filter "*.log" -ErrorAction SilentlyContinue | ForEach-Object {
            $script:LogBaselines[$_.FullName] = $_.Length
        }
    }
    $script:VmLogStaleCount    = 0
    $script:VmLogEverActive    = $false
    $script:ServiceDownCount   = 0
    $script:EventLogHitCount   = 0
    $script:HeartbeatFailCount = 0
    $script:VmcomputeLeakCount = 0
}

# -- Prevent duplicate instances -----------------------------------------
$mutexName = "Global\ClaudeHealthMonitor_v4.8"
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
    # Mutex creation failed -- continue anyway (better to monitor than not)
}

# -- Main loop -----------------------------------------------------------
Write-WatchLog "================================================================"
Write-WatchLog "Health Monitor v$Version started"
Write-WatchLog "  Poll interval : ${PollInterval}s"
Write-WatchLog "  Cooldown      : ${Cooldown} min"
Write-WatchLog "  Fix script    : $fixScript"
Write-WatchLog "  Log directory : $ClaudeLogDir"
Write-WatchLog "  Safety        : user-activity block, 180s grace period, consecutive-check gates"
Write-WatchLog "  Monitors      : logs, service(x2), events(x2)+HCS, NAT, heartbeat(x3), staleness(x5), vmcompute-health, hcs-state, time sync"
Write-WatchLog "================================================================"

try {
    while ($true) {
        try {
            Update-LogFile

            $claudeRunning = @(Get-Process -Name "claude" -ErrorAction SilentlyContinue).Count -gt 0

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

                # Check 5: Hyper-V heartbeat (3 consecutive checks, grace period)
                $hbIssue = Test-VmHeartbeat
                if ($hbIssue) {
                    Invoke-AutoFix -Reason $hbIssue
                    continue
                }

                # Check 6: VM log staleness (5 consecutive checks, 300s threshold, must have been active)
                $staleIssue = Test-VmLogStaleness
                if ($staleIssue) {
                    Invoke-AutoFix -Reason $staleIssue
                    continue
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
                    Write-WatchLog "HCS STATE: $hcsStateIssue"
                    # Don't auto-fix — this is informational + pre-emptive cleanup
                    # Try hcsdiag close for stale VMs proactively
                    try {
                        $hcsdiagPath = "$env:SystemRoot\System32\hcsdiag.exe"
                        if (Test-Path $hcsdiagPath) {
                            $hcsList = & $hcsdiagPath list 2>&1 | Out-String
                            $vmEntries = ([regex]::Matches($hcsList, "cowork-vm")).Count
                            if ($vmEntries -gt 1) {
                                # Parse all cowork-vm GUIDs
                                $guidPattern = '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})'
                                $vmGuids = @()
                                $currentGuid = $null
                                foreach ($line in ($hcsList -split "`r?`n")) {
                                    if ($line -match "^\s*$guidPattern\s*$") {
                                        $currentGuid = $Matches[1]
                                    } elseif ($currentGuid -and $line -match "cowork-vm") {
                                        $vmGuids += $currentGuid
                                        $currentGuid = $null
                                    }
                                }
                                # Close all EXCEPT the last one (most likely the active VM)
                                if ($vmGuids.Count -gt 1) {
                                    $staleGuids = $vmGuids[0..($vmGuids.Count - 2)]
                                    foreach ($guid in $staleGuids) {
                                        & $hcsdiagPath close $guid 2>&1 | Out-Null
                                        Write-WatchLog "HCS: proactively closed stale cowork-vm ($guid)"
                                    }
                                }
                            }
                        }
                    } catch {}
                }

                # ---- Maintenance (non-fix actions) ----
                Set-VmWorkerPriority

                $timeIssue = Test-TimeSyncHealth
                if ($timeIssue) {
                    Write-WatchLog "TIME WARNING: $timeIssue"
                }
            }

            Apply-DynamicMemoryFlag
        } catch {
            Write-WatchLog "MONITOR ERROR: $($_.Exception.Message)"
        }

        Start-Sleep -Seconds $PollInterval
    }
} finally {
    if ($script:Mutex) {
        try {
            $script:Mutex.ReleaseMutex()
            $script:Mutex.Dispose()
        } catch {}
    }
    Write-WatchLog "Health monitor stopped"
}
