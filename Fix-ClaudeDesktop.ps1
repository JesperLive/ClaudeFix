<#
.SYNOPSIS
    Claude Desktop / Cowork -- Reset & Fix

.DESCRIPTION
    Kills all Claude processes, stops CoworkVMService, recovers from HCS
    (Host Compute Service) errors, performs orphan compute system cleanup,
    purges stale VM cache, restarts the service, and relaunches Claude
    Desktop with elevated privileges when available.

    Does NOT touch: config files, MCP servers, conversations.
    Fully automatic -- no user interaction required.

    Works with or without admin privileges. If run without admin,
    service control falls back to process-level operations and Claude
    handles service restart automatically on launch.

.PARAMETER SkipLaunch
    Reset the VM service but don't relaunch Claude Desktop afterwards.

.PARAMETER Quiet
    Suppress the "press any key" prompt at the end.

.PARAMETER KeepCache
    Skip the VM cache purge (Step 6). Use this to avoid re-downloading
    the ~2-3 GB VM bundle. If the fix fails with -KeepCache, run again
    without it to force a clean rebuild.

.PARAMETER WhatIf
    Show what would happen without actually doing anything.

.NOTES
    Version : 4.6.0
    Author  : Jesper Driessen
    Licence : MIT
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$SkipLaunch,
    [switch]$Quiet,
    [switch]$KeepCache
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

    Write-Host ""
    Write-Host "  Requesting admin privileges for full service control..." -ForegroundColor DarkGray
    Write-Host "  (If you decline, the script will still work but some" -ForegroundColor DarkGray
    Write-Host "   operations may be slower or less thorough.)" -ForegroundColor DarkGray
    Write-Host ""

    try {
        Start-Process PowerShell -ArgumentList $elevateArgs -Verb RunAs -Wait
        exit  # Elevated copy ran successfully
    } catch {
        Write-Host "  [i] Running without admin -- service control will be limited" -ForegroundColor Yellow
        Write-Host ""
        # Continue running as normal user
    }
}

# -- Running (elevated or not) ---------------------------------------
Set-StrictMode -Version Latest

# -- Constants -------------------------------------------------------
$Version         = "4.6.0"
$ServiceName     = "CoworkVMService"
$ServiceExe      = "cowork-svc"
$ProcessName     = "claude"
$ClaudeAppData   = Join-Path $env:APPDATA "Claude"
$VmCachePath     = Join-Path $ClaudeAppData "claude-code-vm"
$BundlePath      = Join-Path $ClaudeAppData "vm_bundles"
$ExePathCache    = Join-Path $ClaudeAppData ".claude-exe-path"
$LogDir          = Join-Path $ClaudeAppData "fix-logs"
$ServiceTimeout  = 8
$StartPollMax    = 20   # increased from 12 -- give the service more time after boot
$PostLaunchWait  = 10   # seconds to wait after launching Claude before health check
$MaxRetries      = 3    # how many times to retry the full fix cycle
$script:CapturedClaudeExe = $null  # set in Step 1 from running process

# -- Logging ---------------------------------------------------------
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }
$LogFile = Join-Path $LogDir ("fix_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))

# Clean up logs older than 30 days
try {
    Get-ChildItem $LogDir -Filter "fix_*.log" -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-30) } |
        Remove-Item -Force -ErrorAction SilentlyContinue
} catch {}
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
    try { $script:LogLines | Out-File -FilePath $LogFile -Encoding utf8 -Force }
    catch { Write-Host "  [!] Could not write log file" -ForegroundColor DarkGray }
}

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
    # 0. Path captured from running process (Step 1)
    if ($script:CapturedClaudeExe -and (Test-Path $script:CapturedClaudeExe)) {
        Log "Using path captured from running process: $($script:CapturedClaudeExe)" -Colour DarkGray -Indent
        return $script:CapturedClaudeExe
    }

    # 1. Cached path
    if (Test-Path $ExePathCache) {
        $cached = (Get-Content $ExePathCache -Raw).Trim()
        if ($cached -and (Test-Path $cached)) {
            Log "Found (cached): $cached" -Colour DarkGray -Indent
            return $cached
        } else {
            Log "Cached path invalid, searching..." -Colour DarkGray -Indent
            Remove-Item $ExePathCache -Force -ErrorAction SilentlyContinue
        }
    }

    # 2. App Paths registry (where modern apps register themselves)
    $appPathsKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\claude.exe"
    if (Test-Path $appPathsKey) {
        $appPath = (Get-ItemProperty $appPathsKey -ErrorAction SilentlyContinue).'(default)'
        if (-not $appPath) {
            $appPath = (Get-ItemProperty $appPathsKey -ErrorAction SilentlyContinue).'Path'
            if ($appPath) { $appPath = Join-Path $appPath "Claude.exe" }
        }
        if ($appPath -and (Test-Path $appPath)) {
            Log "Found (App Paths): $appPath" -Colour DarkGray -Indent
            $appPath | Out-File -FilePath $ExePathCache -Encoding utf8 -Force -ErrorAction SilentlyContinue
            return $appPath
        }
    }
    Log "Not in App Paths, checking common locations..." -Colour DarkGray -Indent

    # 3. Common install paths (including Squirrel/Electron locations)
    $searchPaths = @(
        (Join-Path $env:LOCALAPPDATA "Programs\claude\Claude.exe"),
        (Join-Path $env:LOCALAPPDATA "claude\Claude.exe"),
        (Join-Path $env:LOCALAPPDATA "Claude\Claude.exe"),
        (Join-Path $env:LOCALAPPDATA "AnthropicClaude\Claude.exe"),
        (Join-Path $env:LOCALAPPDATA "Anthropic\Claude\Claude.exe"),
        (Join-Path $env:ProgramFiles "Claude\Claude.exe"),
        (Join-Path $env:ProgramFiles "Anthropic\Claude\Claude.exe"),
        (Join-Path ${env:ProgramFiles(x86)} "Claude\Claude.exe")
    )
    foreach ($p in $searchPaths) {
        if (Test-Path $p) {
            Log "Found (common path): $p" -Colour DarkGray -Indent
            $p | Out-File -FilePath $ExePathCache -Encoding utf8 -Force -ErrorAction SilentlyContinue
            return $p
        }
    }
    Log "Not in common paths, checking registry..." -Colour DarkGray -Indent

    # 4. Registry uninstall keys
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    foreach ($rp in $regPaths) {
        $entries = Get-ItemProperty $rp -ErrorAction SilentlyContinue |
                   Where-Object { $_.PSObject.Properties['DisplayName'] -and $_.DisplayName -match "Claude" }
        foreach ($entry in $entries) {
            # Try InstallLocation
            if ($entry.InstallLocation) {
                $candidate = Join-Path $entry.InstallLocation "Claude.exe"
                if (Test-Path $candidate) {
                    Log "Found (registry InstallLocation): $candidate" -Colour DarkGray -Indent
                    $candidate | Out-File -FilePath $ExePathCache -Encoding utf8 -Force -ErrorAction SilentlyContinue
                    return $candidate
                }
            }
            # Try DisplayIcon (often points to the exe)
            if ($entry.DisplayIcon) {
                $iconPath = $entry.DisplayIcon -replace ',.*$', ''
                if ($iconPath -match "Claude\.exe$" -and (Test-Path $iconPath)) {
                    Log "Found (registry DisplayIcon): $iconPath" -Colour DarkGray -Indent
                    $iconPath | Out-File -FilePath $ExePathCache -Encoding utf8 -Force -ErrorAction SilentlyContinue
                    return $iconPath
                }
            }
        }
    }
    Log "Not in registry, checking Start Menu..." -Colour DarkGray -Indent

    # 5. Start Menu shortcuts
    $menuPaths = @(
        (Join-Path $env:APPDATA "Microsoft\Windows\Start Menu"),
        (Join-Path $env:ProgramData "Microsoft\Windows\Start Menu")
    )
    foreach ($mp in $menuPaths) {
        $lnk = Get-ChildItem $mp -Recurse -Filter "Claude*.lnk" -ErrorAction SilentlyContinue |
                Select-Object -First 1
        if ($lnk) {
            $shell  = New-Object -ComObject WScript.Shell
            $target = $shell.CreateShortcut($lnk.FullName).TargetPath
            if ($target -and (Test-Path $target)) {
                Log "Found (Start Menu shortcut): $target" -Colour DarkGray -Indent
                $target | Out-File -FilePath $ExePathCache -Encoding utf8 -Force -ErrorAction SilentlyContinue
                return $target
            }
        }
    }
    Log "Not in Start Menu, scanning LocalAppData..." -Colour DarkGray -Indent

    # 6. Brute-force scan (broad pattern -- catches claude.exe, Claude.exe, claude-desktop.exe, etc.)
    $scanDirs = @($env:LOCALAPPDATA, $env:ProgramFiles, ${env:ProgramFiles(x86)})
    foreach ($scanDir in $scanDirs) {
        if (-not $scanDir -or -not (Test-Path $scanDir)) { continue }
        $found = Get-ChildItem $scanDir -Recurse -Filter "*claude*.exe" `
                     -Depth 5 -ErrorAction SilentlyContinue |
                 Where-Object { $_.Name -notmatch "unins|setup|update" } |
                 Select-Object -First 1
        if ($found) {
            Log "Found (scan): $($found.FullName)" -Colour DarkGray -Indent
            $found.FullName | Out-File -FilePath $ExePathCache -Encoding utf8 -Force -ErrorAction SilentlyContinue
            return $found.FullName
        }
    }

    # 7. where.exe as last resort
    try {
        $whereResult = where.exe Claude.exe 2>&1
        if ($LASTEXITCODE -eq 0 -and $whereResult -and (Test-Path $whereResult[0])) {
            $exePath = "$($whereResult[0])"
            Log "Found (where.exe): $exePath" -Colour DarkGray -Indent
            $exePath | Out-File -FilePath $ExePathCache -Encoding utf8 -Force -ErrorAction SilentlyContinue
            return $exePath
        }
    } catch {}

    return $null
}

# -- Service restart function ----------------------------------------
function Restart-CoworkService {
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $svc) {
        Log "[!] Service not found" -Colour Red -Indent
        return $false
    }

    # Stop if running
    if ($svc.Status -eq "Running") {
        if ($script:IsAdmin) {
            try {
                $svc.Stop()
                $svc.WaitForStatus("Stopped", (New-TimeSpan -Seconds $ServiceTimeout))
            } catch {
                try { Stop-Process -Name $ServiceExe -Force -ErrorAction Stop } catch {}
            }
        } else {
            # Non-admin: try to kill the service process directly
            try { Stop-Process -Name $ServiceExe -Force -ErrorAction Stop } catch {
                Log "[i] Cannot stop service without admin -- Claude relaunch will restart it" -Colour DarkGray -Indent
            }
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
            } catch {}
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

# -- HCS error detection -----------------------------------------------
function Test-RecentHcsErrors {
    <#
    .SYNOPSIS
        Checks for recent HCS (Host Compute Service) errors.
        Returns $null if clean, "json_corruption" if 0xC037010D detected,
        or "hcs_error" for other HCS errors.
    #>
    # Check 1: HCS Compute event log
    try {
        $hcsFilter = @{
            LogName   = "Microsoft-Windows-Hyper-V-Compute-Admin"
            Level     = @(1, 2)  # Critical, Error
            StartTime = (Get-Date).AddMinutes(-5)
        }
        $hcsEvents = @(Get-WinEvent -FilterHashtable $hcsFilter -MaxEvents 10 -ErrorAction SilentlyContinue)
        if ($hcsEvents) {
            foreach ($evt in $hcsEvents) {
                $xml = $evt.ToXml()
                if ($xml -match "0xC037010D" -or $xml -match "Invalid JSON document") {
                    return "json_corruption"
                }
            }
            return "hcs_error"
        }
    } catch {}

    # Check 2: Claude log files for HCS error patterns
    $hcsPatterns = @("HCS operation failed", "failed to create compute system", "HcsWaitForOperationResult")
    $claudeLogDir = Join-Path $env:APPDATA "Claude\logs"
    if (Test-Path $claudeLogDir) {
        try {
            $recentLogs = Get-ChildItem $claudeLogDir -Filter "*.log" -ErrorAction SilentlyContinue |
                          Where-Object { ((Get-Date) - $_.LastWriteTime).TotalMinutes -lt 5 }
            foreach ($logFile in $recentLogs) {
                try {
                    $content = Get-Content $logFile.FullName -Tail 50 -ErrorAction SilentlyContinue
                    $text = $content -join "`n"
                    foreach ($pattern in $hcsPatterns) {
                        if ($text -match [regex]::Escape($pattern)) { return "hcs_error" }
                    }
                } catch {}
            }
        } catch {}
    }

    return $null
}

# ====================================================================
# MAIN
# ====================================================================
try {

# -- Header ----------------------------------------------------------
Write-Host ""
Write-Host "  +-------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |  CLAUDE DESKTOP / COWORK -- RESET & FIX   |" -ForegroundColor Cyan
Write-Host "  |  v$Version                                  |" -ForegroundColor DarkGray
Write-Host "  +-------------------------------------------+" -ForegroundColor Cyan
Write-Host ""

# -- Prevent concurrent Fix runs ----------------------------------------
$fixMutexName = "Global\ClaudeDesktopFix_v4.6"
$fixMutex = $null
try {
    $fixMutex = [System.Threading.Mutex]::new($false, $fixMutexName)
    if (-not $fixMutex.WaitOne(0)) {
        Log "Another Fix instance is already running -- exiting" -Colour DarkGray
        Save-Log
        exit 0
    }
} catch {
    # Mutex creation failed -- continue anyway
}

if (-not $script:IsAdmin) {
    Log "Running without admin (limited service control)" -Colour DarkGray
}
if ($WhatIfPreference) {
    Log "DRY RUN -- no changes will be made" -Colour Yellow
}
Write-Host ""

# ====================================================================
# SAFETY GATE -- Block when called by automation while user is active
# ====================================================================
# When -Quiet is set (called by health monitor or boot task), check if
# Claude Desktop is actively in use. This prevents killing Claude while
# the user is typing, Cowork is running, or Code is doing a task.
# Manual runs (no -Quiet) always proceed -- user explicitly wants a fix.
if ($Quiet) {
    $ClaudeLogDir = Join-Path $ClaudeAppData "logs"
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
        } catch {}
    }

    # Check 2: VM log active within 120s (Code may be thinking)
    if (-not $isActive) {
        $vmLog = Join-Path $ClaudeLogDir "cowork_vm_node.log"
        if (Test-Path $vmLog) {
            $ageSec = ((Get-Date) - (Get-Item $vmLog).LastWriteTime).TotalSeconds
            if ($ageSec -lt 120) { $isActive = $true }
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
                    $idleMs = [Environment]::TickCount - $lastInput.dwTime
                    if ($idleMs -lt 180000) { $isActive = $true }
                }
            }
        } catch {}
    }

    if ($isActive) {
        $msg = "BLOCKED: User/Code appears active -- skipping automated fix"
        Log $msg -Colour Yellow
        Save-Log
        exit 0
    }
}

# ====================================================================
# STEP 1 -- Kill all Claude processes
# ====================================================================
Log "[1/9] Terminating Claude processes..." -Colour Yellow

$claudeProcs = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
if ($claudeProcs.Count -gt 0) {
    # Capture the exe path BEFORE killing -- this is the most reliable detection method
    foreach ($cp in $claudeProcs) {
        try {
            $exeFromProc = $cp.MainModule.FileName
            if ($exeFromProc -and (Test-Path $exeFromProc)) {
                $script:CapturedClaudeExe = $exeFromProc
                Log "Captured exe path: $exeFromProc" -Colour DarkGray -Indent
                # Cache it for future runs
                $exeFromProc | Out-File -FilePath $ExePathCache -Encoding utf8 -Force -ErrorAction SilentlyContinue
                break
            }
        } catch {}
    }
    if (-not $script:CapturedClaudeExe) {
        # Try via WMI/CIM as fallback (works even when MainModule is access-denied)
        try {
            $wmiProc = Get-CimInstance Win32_Process -Filter "Name LIKE '%claude%'" -ErrorAction SilentlyContinue |
                       Select-Object -First 1
            if ($wmiProc -and $wmiProc.ExecutablePath -and (Test-Path $wmiProc.ExecutablePath)) {
                $script:CapturedClaudeExe = $wmiProc.ExecutablePath
                Log "Captured exe path (WMI): $($wmiProc.ExecutablePath)" -Colour DarkGray -Indent
                $wmiProc.ExecutablePath | Out-File -FilePath $ExePathCache -Encoding utf8 -Force -ErrorAction SilentlyContinue
            }
        } catch {}
    }
    if ($PSCmdlet.ShouldProcess("$($claudeProcs.Count) Claude process(es)", "Stop")) {
        $claudeProcs | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    Log "Killed $($claudeProcs.Count) Claude process(es)" -Colour Green -Indent
} else {
    Log "No Claude processes running" -Colour DarkGray -Indent
}
Start-Sleep -Seconds 1

# ====================================================================
# STEP 2 -- Stop CoworkVMService
# ====================================================================
Log "[2/9] Stopping $ServiceName..." -Colour Yellow

$svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if (-not $svc) {
    Log "Service not found -- is Cowork installed?" -Colour DarkGray -Indent
} elseif ($svc.Status -ne "Running") {
    Log "Service already stopped ($($svc.Status))" -Colour DarkGray -Indent
} else {
    if ($PSCmdlet.ShouldProcess($ServiceName, "Stop")) {
        $stopped = $false
        if ($script:IsAdmin) {
            try {
                $svc.Stop()
                $svc.WaitForStatus("Stopped", (New-TimeSpan -Seconds $ServiceTimeout))
                Log "Service stopped gracefully" -Colour Green -Indent
                $stopped = $true
            } catch {}
        }
        if (-not $stopped) {
            Log "Force-killing $ServiceExe..." -Colour DarkYellow -Indent
            try { Stop-Process -Name $ServiceExe -Force -ErrorAction Stop }
            catch {
                try { Start-Process "taskkill" -ArgumentList "/F /IM $ServiceExe.exe" -NoNewWindow -Wait }
                catch {}
            }
            Log "Force-killed" -Colour Green -Indent
        }
    }
}
Start-Sleep -Seconds 1

# ====================================================================
# STEP 3 -- HCS service recovery
# ====================================================================
Log "[3/9] Checking HCS service health..." -Colour Yellow

try {
    $hcsDetected = Test-RecentHcsErrors
    if ($hcsDetected -eq "json_corruption") {
        Log "CRITICAL: HCS JSON corruption detected (0xC037010D)" -Colour Red -Indent
        Log "This error is NOT recoverable by restarting vmcompute." -Colour Red -Indent
        Log "Required fix: Hyper-V nuclear reset" -Colour Yellow -Indent
        Log "  1. Open admin PowerShell" -Colour White -Indent
        Log "  2. dism /online /disable-feature /featurename:Microsoft-Hyper-V-All" -Colour White -Indent
        Log "  3. Reboot" -Colour White -Indent
        Log "  4. dism /online /enable-feature /featurename:Microsoft-Hyper-V-All" -Colour White -Indent
        Log "  5. Reboot" -Colour White -Indent
        if (-not $Quiet) {
            Log "" -Colour White
            Log "The script will still attempt a vmcompute restart, but it is" -Colour DarkGray -Indent
            Log "unlikely to resolve JSON corruption. If Cowork fails to start" -Colour DarkGray -Indent
            Log "after this fix, perform the nuclear reset above." -Colour DarkGray -Indent
        }
    }
    if ($hcsDetected) {
        if ($script:IsAdmin) {
            Log "HCS errors detected -- restarting vmcompute service" -Colour DarkYellow -Indent
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
                    Log "vmcompute not running after 15s -- also restarting vmms" -Colour DarkYellow -Indent
                    Stop-Service vmms -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 3
                    Start-Service vmms -ErrorAction SilentlyContinue
                    Log "vmms service restarted" -Colour Green -Indent
                }
            } catch {
                Log "[!] vmcompute restart failed: $($_.Exception.Message)" -Colour Red -Indent
            }
        } else {
            Log "HCS errors detected but no admin -- vmcompute restart requires elevation" -Colour DarkYellow -Indent
        }
    } else {
        Log "No HCS issues detected" -Colour DarkGray -Indent
    }
} catch {
    Log "[!] HCS check failed: $($_.Exception.Message) -- continuing" -Colour DarkGray -Indent
}

# ====================================================================
# STEP 4 -- Verify no orphan processes remain
# ====================================================================
Log "[4/9] Checking for orphan processes..." -Colour Yellow

$remaining = @(Get-Process -ErrorAction SilentlyContinue | Where-Object {
    ($_.Name -eq $ProcessName) -or ($_.Name -eq $ServiceExe)
})
if ($remaining.Count -gt 0) {
    if ($PSCmdlet.ShouldProcess("$($remaining.Count) orphan process(es)", "Force-kill")) {
        foreach ($proc in $remaining) {
            try { $proc | Stop-Process -Force -ErrorAction Stop } catch {}
        }
        Start-Sleep -Seconds 1
        $stubborn = @(Get-Process -ErrorAction SilentlyContinue | Where-Object {
            ($_.Name -eq $ProcessName) -or ($_.Name -eq $ServiceExe)
        })
        if ($stubborn.Count -gt 0) {
            Log "[!] $($stubborn.Count) process(es) refuse to die -- a reboot may be needed" -Colour Red -Indent
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
Log "[5/9] Checking for orphan compute systems..." -Colour Yellow
try {
    $orphanKilled = $false
    # Method 1: hcsdiag (most reliable for HCS compute systems)
    if ($script:IsAdmin) {
        $hcsdiagPath = "$env:SystemRoot\System32\hcsdiag.exe"
        if (Test-Path $hcsdiagPath) {
            try {
                $hcsList = & $hcsdiagPath list 2>&1 | Out-String
                if ($hcsList -match "(?i)claude|cowork") {
                    Log "Found orphan compute system(s) via hcsdiag" -Colour DarkYellow -Indent
                    $lines = $hcsList -split "`r?`n"
                    $currentGuid = $null
                    $isClaudeVm = $false
                    foreach ($line in $lines) {
                        if ($line -match '^\s*([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\s*$') {
                            if ($isClaudeVm -and $currentGuid) {
                                if ($PSCmdlet.ShouldProcess($currentGuid, "hcsdiag kill")) {
                                    & $hcsdiagPath kill $currentGuid 2>&1 | Out-Null
                                    Log "Killed orphan compute system: $currentGuid" -Colour Green -Indent
                                    $orphanKilled = $true
                                }
                            }
                            $currentGuid = $Matches[1]
                            $isClaudeVm = $false
                        } elseif ($currentGuid -and $line -match '(?i)claude|cowork') {
                            $isClaudeVm = $true
                        }
                    }
                    if ($isClaudeVm -and $currentGuid) {
                        if ($PSCmdlet.ShouldProcess($currentGuid, "hcsdiag kill")) {
                            & $hcsdiagPath kill $currentGuid 2>&1 | Out-Null
                            Log "Killed orphan compute system: $currentGuid" -Colour Green -Indent
                            $orphanKilled = $true
                        }
                    }
                }
            } catch {
                Log "hcsdiag query failed: $($_.Exception.Message)" -Colour DarkGray -Indent
            }
        }
    }
    # Method 2: Hyper-V cmdlets fallback (Stop-VM -TurnOff)
    try {
        $claudeVm = Get-VM -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match "claude" }
        if ($claudeVm) {
            foreach ($vm in $claudeVm) {
                if ($vm.State -ne "Off") {
                    if ($PSCmdlet.ShouldProcess($vm.Name, "Stop-VM -TurnOff")) {
                        Stop-VM -Name $vm.Name -TurnOff -Force -ErrorAction SilentlyContinue
                        Log "Force-stopped VM '$($vm.Name)' via Hyper-V" -Colour Green -Indent
                        $orphanKilled = $true
                    }
                }
            }
        }
    } catch {
        Log "Hyper-V VM check skipped: $($_.Exception.Message)" -Colour DarkGray -Indent
    }
    if (-not $orphanKilled) {
        Log "No orphan compute systems found" -Colour Green -Indent
    }
} catch {
    Log "Orphan VM check failed (non-critical): $($_.Exception.Message)" -Colour DarkGray -Indent
}
Start-Sleep -Seconds 1

# ====================================================================
# STEP 6 -- Purge VM cache (skipped with -KeepCache)
# ====================================================================
if ($KeepCache) {
    Log "[6/9] Keeping VM cache (-KeepCache)" -Colour DarkGray
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
    Log "[6/9] Purging VM cache..." -Colour Yellow
    $cacheDirs = @(
        @{ Path = $VmCachePath; Label = "claude-code-vm" },
        @{ Path = $BundlePath;  Label = "vm_bundles" }
    )
    foreach ($item in $cacheDirs) {
        if (Test-Path $item.Path) {
            $size = (Get-ChildItem $item.Path -Recurse -ErrorAction SilentlyContinue |
                     Measure-Object -Property Length -Sum).Sum
            $sizeMB = [math]::Round($size / 1MB, 1)
            if ($PSCmdlet.ShouldProcess($item.Label, "Delete ($sizeMB MB)")) {
                Remove-Item $item.Path -Recurse -Force -ErrorAction SilentlyContinue
            }
            Log "$($item.Label) removed ($sizeMB MB freed)" -Colour Green -Indent
        } else {
            Log "$($item.Label) not present" -Colour DarkGray -Indent
        }
    }
}

# ====================================================================
# STEP 7 -- Restart CoworkVMService (with extended polling)
# ====================================================================
Log "[7/9] Starting $ServiceName..." -Colour Yellow

if ($PSCmdlet.ShouldProcess($ServiceName, "Start")) {
    if ($script:IsAdmin) {
        $svcOk = Restart-CoworkService
        if ($svcOk) {
            Log "Service running" -Colour Green -Indent
        } else {
            Log "[!] Service failed to start -- will retry after launch" -Colour Yellow -Indent
        }
    } else {
        Log "Skipping manual service start (no admin)" -Colour DarkGray -Indent
        Log "Claude will restart the service automatically when it launches" -Colour DarkGray -Indent
    }
}

# ====================================================================
# STEP 8 -- Relaunch Claude Desktop
# ====================================================================
if ($SkipLaunch) {
    Log "[8/9] Skipping Claude launch (-SkipLaunch)" -Colour DarkGray
    Log "[9/9] Skipping health check (-SkipLaunch)" -Colour DarkGray
} else {
    Log "[8/9] Launching Claude Desktop..." -Colour Yellow

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
                $claudeProc = Get-Process -Name "claude" -ErrorAction SilentlyContinue
                if ($claudeProc) {
                    Log "Launched elevated via scheduled task: $elevTaskPath" -Colour Green -Indent
                    $launched = $true
                } else {
                    Log "Scheduled task ran but Claude process not detected -- falling back" -Colour DarkYellow -Indent
                }
            }
        } else {
            Log "LaunchClaudeAdmin task not found -- falling back to standard launch" -Colour DarkGray -Indent
            Log "(Run Prevent-ClaudeIssues.bat to enable elevated launch)" -Colour DarkGray -Indent
        }
    } catch {
        Log "Elevated launch failed: $($_.Exception.Message) -- falling back" -Colour DarkYellow -Indent
    }

    # Method A: MSIX / AppX -- query the package and launch properly
    $appxPkg = Get-AppxPackage -Name "*Claude*" -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $launched -and $appxPkg) {
        $pfn = $appxPkg.PackageFamilyName
        Log "Detected MSIX install: $pfn" -Colour DarkGray -Indent
        # Get the Application ID from the manifest
        $appId = $null
        try {
            $manifestPath = Join-Path $appxPkg.InstallLocation "AppxManifest.xml"
            if (Test-Path $manifestPath) {
                [xml]$manifest = Get-Content $manifestPath -ErrorAction Stop
                $ns = New-Object Xml.XmlNamespaceManager($manifest.NameTable)
                $ns.AddNamespace("x", "http://schemas.microsoft.com/appx/manifest/foundation/windows10")
                $appNode = $manifest.SelectSingleNode("//x:Application", $ns)
                if ($appNode) { $appId = $appNode.GetAttribute("Id") }
            }
        } catch {}
        if (-not $appId) { $appId = "App" }  # common default

        $shellUri = "shell:AppsFolder\$pfn!$appId"
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
            Log "Exe is in WindowsApps -- skipping direct launch (would create loose instance)" -Colour DarkYellow -Indent
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
    # STEP 9 -- Wait for Cowork workspace readiness
    # ====================================================================
    # Detection strategy (in order of reliability):
    #   A. Log file: cowork_vm_node.log -- "Startup complete" or "Keepalive"
    #   B. Hyper-V VM state: Get-VM "claudevm" shows Running + heartbeat
    #   C. Log file: step markers like "guest_vsock_connect completed"
    #   D. File size stability fallback (last resort)
    # The named pipe RPC is NOT usable (requires signed client executable).
    # ====================================================================
    Log "[9/9] Waiting for Cowork workspace..." -Colour Yellow

    $vmReady = $false
    $vmLogDir = Join-Path $ClaudeAppData "logs"
    $vmLogFile = Join-Path $vmLogDir "cowork_vm_node.log"

    # Record the log file size at the START so we only check new entries
    $logBaselineSize = 0
    if (Test-Path $vmLogFile) {
        $logBaselineSize = (Get-Item $vmLogFile -ErrorAction SilentlyContinue).Length
    }

    # Helper: check recent log entries for boot completion markers
    function Test-VmLogReady {
        param([long]$Baseline)
        if (-not (Test-Path $vmLogFile)) { return $null }
        try {
            $fi = Get-Item $vmLogFile -ErrorAction Stop
            if ($fi.Length -le $Baseline) { return $null }
            # Read only the new portion of the log
            $stream = New-Object System.IO.FileStream(
                $vmLogFile, [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            $stream.Position = $Baseline
            $reader = New-Object System.IO.StreamReader($stream)
            $newContent = $reader.ReadToEnd()
            $reader.Close()
            $stream.Close()
            # Check for completion markers (most definitive first)
            if ($newContent -match "Startup complete") { return "startup-complete" }
            if ($newContent -match "\[Keepalive\]") { return "keepalive" }
            if ($newContent -match "guest_vsock_connect completed") { return "vsock-connected" }
            if ($newContent -match "sdk_install completed") { return "sdk-installed" }
            return $null
        } catch { return $null }
    }

    # Helper: check Hyper-V VM state (may not work without Hyper-V module)
    function Test-HyperVReady {
        try {
            $vm = Get-VM -VMName "claudevm" -ErrorAction Stop
            if ($vm.State -eq "Running") {
                # Check heartbeat if available
                $hb = Get-VMIntegrationService -VMName "claudevm" -Name "Heartbeat" -ErrorAction SilentlyContinue
                if ($hb -and $hb.PrimaryStatusDescription -eq "OK") {
                    return "running+heartbeat"
                }
                return "running"
            }
            return $null
        } catch { return $null }
    }

    $vmTimeout = 240   # 4 min -- full boot can take a while after fresh purge
    $vmElapsed = 0
    $lastStatus = ""
    $hvChecked = $false
    $hvAvailable = $false

    while ($vmElapsed -lt $vmTimeout) {
        Start-Sleep -Seconds 5
        $vmElapsed += 5

        # Service health check
        $svcNow = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if (-not $svcNow -or $svcNow.Status -ne "Running") {
            Log "Service died -- restarting..." -Colour DarkYellow -Indent
            Restart-CoworkService | Out-Null
            continue
        }

        # A. Check log file for boot completion
        $logStatus = Test-VmLogReady -Baseline $logBaselineSize
        if ($logStatus -eq "startup-complete" -or $logStatus -eq "keepalive") {
            $vmReady = $true
            Log "Workspace ready (log: $logStatus)" -Colour Green -Indent
            break
        }

        # B. Check Hyper-V (once, to see if cmdlets work)
        if (-not $hvChecked) {
            $hvChecked = $true
            $hvState = Test-HyperVReady
            if ($null -ne $hvState) {
                $hvAvailable = $true
                Log "Hyper-V VM detected: $hvState" -Colour DarkGray -Indent
            }
        } elseif ($hvAvailable -and $vmElapsed -ge 20) {
            $hvState = Test-HyperVReady
            if ($hvState -eq "running+heartbeat") {
                # Heartbeat OK -- give logs a few more seconds then accept
                Start-Sleep -Seconds 5
                $logStatus = Test-VmLogReady -Baseline $logBaselineSize
                if ($logStatus) {
                    $vmReady = $true
                    Log "Workspace ready (heartbeat + log: $logStatus)" -Colour Green -Indent
                    break
                }
            }
        }

        # C. Progress from log step markers
        $curStatus = ""
        if ($logStatus -eq "sdk-installed") {
            $curStatus = "Finishing setup... (${vmElapsed}s)"
        } elseif ($logStatus -eq "vsock-connected") {
            $curStatus = "Installing SDK... (${vmElapsed}s)"
        } else {
            # Check if VM files are being downloaded
            $curSizeMB = 0
            if (Test-Path $VmCachePath) {
                $curSize = (Get-ChildItem $VmCachePath -Recurse -ErrorAction SilentlyContinue |
                            Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                $curSizeMB = [math]::Round($curSize / 1MB, 0)
            }
            if ($curSizeMB -lt 50) {
                $curStatus = "Setting up workspace... (${vmElapsed}s, ${curSizeMB} MB)"
            } else {
                $curStatus = "Starting workspace... (${vmElapsed}s)"
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
            Log "Workspace partially ready (log: $finalLog) -- may still be loading" -Colour Yellow -Indent
        } else {
            Log "[!] Workspace not confirmed ready after ${vmTimeout}s" -Colour Yellow -Indent
            Log "Open a Cowork session in Claude to trigger setup" -Colour DarkGray -Indent
        }
    }
}

# -- Bring this window to the front (Claude may have taken focus) ----
try { [Win32Window]::BringToFront() } catch {}

# ====================================================================
# Summary
# ====================================================================
Write-Host ""
Write-Host "  +-------------------------------------------+" -ForegroundColor Green
Write-Host "  |           OPERATION COMPLETE               |" -ForegroundColor Green
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
} catch {}

Write-Host ""
Log "Log saved to: $LogFile" -Colour DarkGray -Indent

Save-Log

} catch {
    Write-Host ""
    Write-Host "  +-------------------------------------------+" -ForegroundColor Red
    Write-Host "  |           UNEXPECTED ERROR                 |" -ForegroundColor Red
    Write-Host "  +-------------------------------------------+" -ForegroundColor Red
    Write-Host ""
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor DarkGray
    Write-Host ""
    Save-Log
} finally {
    if ($fixMutex) {
        try { $fixMutex.ReleaseMutex(); $fixMutex.Dispose() } catch {}
    }
}

# -- Always pause unless -Quiet --------------------------------------
if (-not $Quiet) {
    Write-Host ""
    Write-Host "  Press any key to close..." -ForegroundColor DarkGray
    try { [Win32Window]::Flash() } catch {}
    [void][System.Console]::ReadKey($true)
    try { [Win32Window]::StopFlash() } catch {}
}
