<#
.SYNOPSIS
    Claude Desktop / Cowork -- Health Monitor

.DESCRIPTION
    Persistent background monitor that detects VirtioFS/Plan9 mount failures
    in Claude Desktop and automatically runs the fix script.

    Monitors:
    - Claude log files for "bad address" and mount failure messages
    - CoworkVMService status (stopped while Claude is running)
    - Windows Event Log for service and Hyper-V errors/warnings
    - WinNAT rules (VM network connectivity)
    - Hyper-V Integration Services heartbeat
    - VM log staleness (hung VM detection)
    - Host clock drift (NTP/time sync)

    When a failure is detected, automatically runs Fix-ClaudeDesktop.ps1
    to reset the VM and relaunch Claude. A cooldown prevents rapid cycling.

    Designed to run as a hidden scheduled task (installed by Prevent-ClaudeIssues.ps1)
    but can also be started manually for foreground monitoring.

.PARAMETER PollInterval
    Seconds between health checks (default: 30).

.PARAMETER Cooldown
    Minutes to wait between auto-fix runs (default: 3).

.PARAMETER Quiet
    Suppress console output (for scheduled task use).

.NOTES
    Version : 2.0.0
    Author  : Jesper Driessen
    Licence : MIT
#>

[CmdletBinding()]
param(
    [int]$PollInterval = 30,
    [int]$Cooldown = 3,
    [switch]$Quiet
)

Set-StrictMode -Version Latest

# -- Constants -----------------------------------------------------------
$Version        = "2.0.0"
$ServiceName    = "CoworkVMService"
$ClaudeAppData  = Join-Path $env:APPDATA "Claude"
$ClaudeLogDir   = Join-Path $ClaudeAppData "logs"
$WatchLogDir    = Join-Path $ClaudeAppData "watch-logs"

# Error patterns that indicate VirtioFS mount failure
$ErrorPatterns = @(
    "Plan9 mount failed",
    "bad address",
    "failed to ensure virtiofs mount",
    "RPC error -1"
)

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
$script:LastFixTime          = [datetime]::MinValue
$script:LogBaselines         = @{}   # Track file sizes per log file
$script:FixCount             = 0
$script:LastNatCheckTime     = [datetime]::MinValue
$script:LastTimeSyncCheck    = [datetime]::MinValue
$script:LastHeartbeatStatus  = $null
$script:VmLogStaleCount      = 0     # Consecutive stale checks

# Initialize baselines for existing log files
if (Test-Path $ClaudeLogDir) {
    Get-ChildItem $ClaudeLogDir -Filter "*.log" -ErrorAction SilentlyContinue | ForEach-Object {
        $script:LogBaselines[$_.FullName] = $_.Length
    }
}

# -- Detection functions -------------------------------------------------

function Test-LogsForErrors {
    <#
    .SYNOPSIS
        Scans all Claude log files for new VirtioFS error messages since last check.
        Returns the error pattern found, or $null if clean.
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
            $stream = [System.IO.FileStream]::new(
                $path,
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read,
                [System.IO.FileShare]::ReadWrite)
            $stream.Position = $baseline
            $reader = [System.IO.StreamReader]::new($stream)
            $newContent = $reader.ReadToEnd()
            $reader.Close()
            $stream.Close()

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
        Returns $true if healthy (Claude not running, or service running).
        Returns $false if Claude is running but the service has stopped.
    #>
    $claude = @(Get-Process -Name "claude" -ErrorAction SilentlyContinue)
    if ($claude.Count -eq 0) { return $true }

    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $svc) { return $true }

    if ($svc.Status -ne "Running") { return $false }
    return $true
}

function Test-EventLogErrors {
    <#
    .SYNOPSIS
        Checks Windows Event Log for CoworkVMService errors AND Hyper-V
        Worker/VMMS warnings in the last PollInterval window.
        Returns the first relevant message, or $null if clean.
    #>
    $lookback = [math]::Max($PollInterval, 60)  # at least 60s lookback
    $since = (Get-Date).AddSeconds(-$lookback)

    # Check 1: CoworkVMService errors (original)
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
                    return "CoworkVMService: $msg"
                }
            }
        }
    } catch {}

    # Check 2: Hyper-V Worker errors (vmwp.exe crashes / VM failures)
    try {
        $hvWorkerFilter = @{
            LogName   = "Microsoft-Windows-Hyper-V-Worker-Admin"
            Level     = @(1, 2, 3)  # Critical, Error, Warning
            StartTime = $since
        }
        $hvEvents = Get-WinEvent -FilterHashtable $hvWorkerFilter -MaxEvents 1 -ErrorAction SilentlyContinue
        if ($hvEvents) {
            $msg = ($hvEvents[0].Message -split "`n")[0]
            # Only trigger on messages related to our VM
            if ($msg -match "claude" -or $msg -match "Plan9" -or $msg -match "virtio" -or $msg -match "shared memory") {
                return "Hyper-V Worker: $msg"
            }
        }
    } catch {}

    # Check 3: Hyper-V VMMS errors (VM management service issues)
    try {
        $vmmsFilter = @{
            LogName   = "Microsoft-Windows-Hyper-V-VMMS-Admin"
            Level     = @(1, 2)  # Critical, Error
            StartTime = $since
        }
        $vmmsEvents = Get-WinEvent -FilterHashtable $vmmsFilter -MaxEvents 1 -ErrorAction SilentlyContinue
        if ($vmmsEvents) {
            $msg = ($vmmsEvents[0].Message -split "`n")[0]
            if ($msg -match "claude" -or $msg -match "failed" -or $msg -match "unexpected") {
                return "Hyper-V VMMS: $msg"
            }
        }
    } catch {}

    return $null
}

function Test-WinNatHealth {
    <#
    .SYNOPSIS
        Checks that a WinNAT rule exists for the Cowork VM's network.
        Without NAT, the VM has no outbound connectivity, causing API calls
        and mount operations to fail silently.
        Returns $null if healthy, or a description string if unhealthy.
    #>
    # Only check every 60 seconds (NAT doesn't change that often)
    $now = Get-Date
    if (($now - $script:LastNatCheckTime).TotalSeconds -lt 60) { return $null }
    $script:LastNatCheckTime = $now

    try {
        $natRules = @(Get-NetNat -ErrorAction SilentlyContinue)
        if ($natRules.Count -eq 0) {
            # No NAT rules at all -- this is bad if Claude is running
            # Try to find the Hyper-V internal switch subnet and recreate
            $hvSwitch = Get-VMSwitch -SwitchType Internal -ErrorAction SilentlyContinue |
                        Where-Object { $_.Name -match "WSL|claude|Default|nat" } |
                        Select-Object -First 1
            if (-not $hvSwitch) {
                # Also check for any internal switch
                $hvSwitch = Get-VMSwitch -SwitchType Internal -ErrorAction SilentlyContinue |
                            Select-Object -First 1
            }

            if ($hvSwitch) {
                # Find the adapter connected to this switch
                $adapter = Get-NetAdapter -ErrorAction SilentlyContinue |
                           Where-Object { $_.InterfaceDescription -match "Hyper-V Virtual Ethernet Adapter" -and $_.Name -match $hvSwitch.Name } |
                           Select-Object -First 1
                if (-not $adapter) {
                    # Try matching by vEthernet name pattern
                    $adapter = Get-NetAdapter -Name "vEthernet ($($hvSwitch.Name))" -ErrorAction SilentlyContinue
                }

                if ($adapter) {
                    $ipAddr = Get-NetIPAddress -InterfaceIndex $adapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                              Select-Object -First 1
                    if ($ipAddr) {
                        # Derive the subnet (assume /24 for typical Hyper-V NAT)
                        $prefix = ($ipAddr.IPAddress -split '\.')[0..2] -join '.'
                        $subnet = "$prefix.0/24"
                        try {
                            New-NetNat -Name "CoworkNAT" -InternalIPInterfaceAddressPrefix $subnet -ErrorAction Stop | Out-Null
                            Write-WatchLog "REPAIRED: Created WinNAT rule 'CoworkNAT' for $subnet"
                            return $null  # Fixed it
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
        Returns $null if healthy or unavailable, or a description if unhealthy.
    #>
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
            # "No Contact" is normal during VM boot
            if ($script:LastHeartbeatStatus -eq "Lost" -and $status -eq "OK") {
                Write-WatchLog "Heartbeat recovered for VM '$($claudeVm.Name)'"
            }
            $script:LastHeartbeatStatus = $status
            return $null
        }

        # Heartbeat is not OK (could be "Error", "Lost Communication", etc.)
        if ($script:LastHeartbeatStatus -ne "Lost") {
            $script:LastHeartbeatStatus = "Lost"
            return "VM heartbeat: $status (VM may be hung or unresponsive)"
        }
    } catch {}
    return $null
}

function Test-VmLogStaleness {
    <#
    .SYNOPSIS
        Checks if the VM's node log has gone stale (no updates for >2 minutes
        while Claude is running). This catches silent VM hangs where no error
        is logged but the VM has stopped responding.
        Returns $null if OK, or a description if stale.
    #>
    $vmLogFile = Join-Path $ClaudeLogDir "cowork_vm_node.log"
    if (-not (Test-Path $vmLogFile)) {
        $script:VmLogStaleCount = 0
        return $null
    }

    try {
        $lastWrite = (Get-Item $vmLogFile -ErrorAction Stop).LastWriteTime
        $staleSec = ((Get-Date) - $lastWrite).TotalSeconds

        if ($staleSec -gt 120) {
            # Log hasn't been written to in 2+ minutes
            $script:VmLogStaleCount++
            # Require 3 consecutive stale checks (90 seconds) to avoid false positives
            if ($script:VmLogStaleCount -ge 3) {
                $script:VmLogStaleCount = 0
                return "VM log stale for $([math]::Round($staleSec))s -- VM may be hung"
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
        Checks for significant host clock drift. If the host clock drifts >5 seconds,
        Hyper-V time synchronization can't correct the guest clock, causing
        TLS certificate validation failures and API timeouts inside the VM.
        Only checks every 5 minutes to avoid overhead.
        Returns $null if OK, or a description if drifted.
    #>
    $now = Get-Date
    if (($now - $script:LastTimeSyncCheck).TotalMinutes -lt 5) { return $null }
    $script:LastTimeSyncCheck = $now

    try {
        # Check if W32Time service is running
        $w32svc = Get-Service -Name "W32Time" -ErrorAction SilentlyContinue
        if ($w32svc -and $w32svc.Status -ne "Running") {
            # Try to start it
            try {
                Start-Service -Name "W32Time" -ErrorAction Stop
                Write-WatchLog "REPAIRED: Started W32Time service"
            } catch {
                return "W32Time service stopped -- clock may drift"
            }
        }

        # Check actual drift via w32tm (non-blocking, quick)
        $w32tmResult = & w32tm /stripchart /computer:time.windows.com /dataonly /samples:1 2>&1
        if ($w32tmResult -match "(-?\d+\.\d+)s") {
            $drift = [math]::Abs([double]$Matches[1])
            if ($drift -gt 5.0) {
                # Force resync
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

# -- VM maintenance functions ---------------------------------------------

function Set-VmWorkerPriority {
    <#
    .SYNOPSIS
        Ensures all vmwp.exe (Hyper-V VM Worker) processes run at AboveNormal
        priority. This is not persistent across reboots, so the health monitor
        re-applies it on every poll cycle.
    #>
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
    <#
    .SYNOPSIS
        If the Prevent script wrote a flag file because the VM was running,
        check if the VM is now off and disable dynamic memory.
    #>
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

# -- Auto-fix function ---------------------------------------------------

function Invoke-AutoFix {
    param([string]$Reason)

    $now = Get-Date
    $elapsed = ($now - $script:LastFixTime).TotalMinutes

    if ($elapsed -lt $Cooldown) {
        Write-WatchLog "COOLDOWN: Skipping fix ($Reason) -- last fix $([math]::Round($elapsed, 1)) min ago"
        return
    }

    $script:FixCount++
    Write-WatchLog ">>> AUTO-FIX #$($script:FixCount) TRIGGERED: $Reason"
    $script:LastFixTime = $now

    try {
        # Run the fix script directly (inherits elevation from scheduled task)
        # -Quiet suppresses "Press any key" prompt
        & $fixScript -Quiet
        Write-WatchLog ">>> AUTO-FIX #$($script:FixCount) COMPLETE"
    } catch {
        Write-WatchLog ">>> AUTO-FIX #$($script:FixCount) FAILED: $($_.Exception.Message)"
    }

    # Reset log baselines after fix (log files may have been recreated)
    Start-Sleep -Seconds 10
    $script:LogBaselines = @{}
    if (Test-Path $ClaudeLogDir) {
        Get-ChildItem $ClaudeLogDir -Filter "*.log" -ErrorAction SilentlyContinue | ForEach-Object {
            $script:LogBaselines[$_.FullName] = $_.Length
        }
    }
    # Reset stale counter
    $script:VmLogStaleCount = 0
}

# -- Prevent duplicate instances -----------------------------------------
$mutexName = "Global\ClaudeHealthMonitor_v2"
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
Write-WatchLog "  Monitors      : logs, service, events, NAT, heartbeat, staleness, time sync"
Write-WatchLog "================================================================"

try {
    while ($true) {
        try {
            # Rotate log file at midnight
            Update-LogFile

            # Only monitor when Claude is actually running
            $claudeRunning = @(Get-Process -Name "claude" -ErrorAction SilentlyContinue).Count -gt 0

            if ($claudeRunning) {
                # ---- Critical checks (trigger auto-fix) ----

                # Check 1: VirtioFS errors in log files (most reliable)
                $logError = Test-LogsForErrors
                if ($logError) {
                    Invoke-AutoFix -Reason "Log error: $logError"
                    continue
                }

                # Check 2: Service died while Claude is running
                if (-not (Test-ServiceHealth)) {
                    Invoke-AutoFix -Reason "CoworkVMService stopped while Claude is running"
                    continue
                }

                # Check 3: Event Log errors (catches errors not in log files)
                $eventError = Test-EventLogErrors
                if ($eventError) {
                    Invoke-AutoFix -Reason "Event Log: $eventError"
                    continue
                }

                # Check 4: WinNAT connectivity (VM has no network without it)
                $natIssue = Test-WinNatHealth
                if ($natIssue) {
                    # NAT issues are self-repaired when possible; only trigger fix
                    # if repair failed (the function returns $null on successful repair)
                    Write-WatchLog "NAT WARNING: $natIssue"
                    # Don't auto-fix for NAT -- it may self-recover after repair
                }

                # Check 5: Hyper-V heartbeat (detects hung VMs)
                $hbIssue = Test-VmHeartbeat
                if ($hbIssue) {
                    Invoke-AutoFix -Reason $hbIssue
                    continue
                }

                # Check 6: VM log staleness (catches silent hangs)
                $staleIssue = Test-VmLogStaleness
                if ($staleIssue) {
                    Invoke-AutoFix -Reason $staleIssue
                    continue
                }

                # ---- Maintenance (non-fix actions) ----

                # Keep vmwp.exe at elevated priority
                Set-VmWorkerPriority

                # Check time sync health (auto-repairs drift)
                $timeIssue = Test-TimeSyncHealth
                if ($timeIssue) {
                    Write-WatchLog "TIME WARNING: $timeIssue"
                }
            }

            # Maintenance: apply deferred dynamic memory flag (VM must be off)
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
