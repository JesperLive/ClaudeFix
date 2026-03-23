<#
.SYNOPSIS
    Claude Desktop / Cowork -- Reset & Fix

.DESCRIPTION
    Kills all Claude processes, stops CoworkVMService, recovers from HCS
    (Host Compute Service) errors, performs orphan compute system cleanup,
    purges stale VM cache, restarts the service, and relaunches Claude
    Desktop with elevated privileges when available.

    Use -Close for a clean shutdown without relaunching (kills Claude UI,
    waits for VM shutdown, restarts service for next launch). Pair with
    Stop-ClaudeDesktop.bat for double-click convenience.

    Does NOT touch: config files, MCP servers, conversations.
    Fully automatic -- no user interaction required.

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
      Quick      -- Restart services + basic repair (Steps 1-5, skip cache purge)
      Deep       -- Full nuclear reset (all steps including cache purge)
      Smart      -- Try quick first, escalate to deep if needed (default)
      Diagnostic -- Health check only, no changes

.PARAMETER KeepCache
    Skip the VM cache purge (Step 6). Use this to avoid re-downloading
    the ~2-3 GB VM bundle. If the fix fails with -KeepCache, run again
    without it to force a clean rebuild.

.PARAMETER Close
    Perform a clean shutdown only: kill Claude UI, wait for VM to shut down
    gracefully, restart the service so it is ready for next launch. Does NOT
    relaunch Claude. Useful before a reboot or when you just want to fully
    stop Claude without running a repair.

.PARAMETER WhatIf
    Show what would happen without actually doing anything.

.NOTES
    Version : 5.3.1
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
    if ($Mode)             { $elevateArgs += " -Mode $Mode" }

    Write-Host ""
    Write-Host "  Requesting admin privileges for full service control..." -ForegroundColor DarkGray
    Write-Host "  (If you decline, the script will still work but some" -ForegroundColor DarkGray
    Write-Host "   operations may be slower or less thorough.)" -ForegroundColor DarkGray
    Write-Host ""

    try {
        Start-Process PowerShell -ArgumentList $elevateArgs -Verb RunAs -Wait
        exit 0  # Elevated copy ran successfully
    } catch {
        Write-Host "  [i] Running without admin -- service control will be limited" -ForegroundColor Yellow
        Write-Host ""
        # Continue running as normal user
    }
}

# -- Running (elevated or not) ---------------------------------------
Set-StrictMode -Version Latest

# -- Constants -------------------------------------------------------
$Version         = "5.3.1"
$ServiceName     = "CoworkVMService"
$ServiceExe      = "cowork-svc"
$ProcessName     = "claude"
$ClaudeAppData   = Join-Path $env:APPDATA "Claude"
$VmCachePath     = Join-Path $ClaudeAppData "claude-code-vm"
$BundlePath      = Join-Path $ClaudeAppData "vm_bundles"
$ExePathCache    = Join-Path $ClaudeAppData ".claude-exe-path"
$LogDir          = Join-Path $ClaudeAppData "fix-logs"
$ServiceTimeout  = 30   # VM shutdown takes 10-30s; too short = force-kill = HCS corruption
$StartPollMax    = 20   # increased from 12 -- give the service more time after boot
$PostLaunchWait  = 10   # seconds to wait after launching Claude before health check
$MaxRetries      = 3    # how many times to retry the full fix cycle
$script:CapturedClaudeExe = $null  # set in Step 1 from running process

# -- Logging ---------------------------------------------------------
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }
$script:SessionTimestamp = "{0:yyyyMMdd_HHmmss}" -f (Get-Date)
$LogFile = Join-Path $LogDir "fix_$($script:SessionTimestamp).log"

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

# Transcript backup (v4.8.0) -- catches output even if Save-Log fails
$script:TranscriptFile = Join-Path $LogDir "fix_$($script:SessionTimestamp)_transcript.log"
try { Start-Transcript -Path $script:TranscriptFile -Append -ErrorAction SilentlyContinue } catch {}

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
                $sw = [System.Diagnostics.Stopwatch]::StartNew()
                while ($sw.Elapsed.TotalSeconds -lt $ServiceTimeout) {
                    Start-Sleep -Seconds 2
                    $curSvc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
                    if (-not $curSvc -or $curSvc.Status -eq "Stopped") { break }
                }
                $sw.Stop()
            } catch {
                try { Stop-Process -Name $ServiceExe -Force -ErrorAction Stop } catch {}
            }
        } else {
            try { Stop-Process -Name $ServiceExe -Force -ErrorAction Stop } catch {
                Log "[i] Cannot stop service without admin" -Colour DarkGray -Indent
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
    } catch {}

    # Check 2: Claude log files (keep existing logic)
    $hcsPatterns = @("HCS operation failed", "failed to create compute system",
                     "HcsWaitForOperationResult", "0x800707DE")
    $claudeLogDirs = @(
        (Join-Path $env:ProgramData "Claude\Logs"),
        (Join-Path $env:APPDATA "Claude\logs")
    )
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
                } catch {}
            }
        } catch {}
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
    $candidatePaths = @(
        (Join-Path $env:ProgramData "Claude\Logs\cowork-service.log"),
        (Join-Path $ClaudeAppData "logs\cowork-service.log")
    )

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
            } catch {}
        }
        # Format A: "yyyy-MM-dd HH:mm:ss.fff" (original format)
        elseif ($line -match '^\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\.\d{3})') {
            try {
                $ts = [datetime]::ParseExact($Matches[1].Trim(), "yyyy-MM-dd HH:mm:ss.fff",
                    [System.Globalization.CultureInfo]::InvariantCulture)
                if (($now - $ts).TotalSeconds -le $WindowSeconds) {
                    $recentLines += $line
                }
            } catch {}
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
            } catch {}
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

function Invoke-WithTimeout {
    <#
    .SYNOPSIS
        Runs a scriptblock in a background job with a timeout.
        Returns $Default if the job does not complete in time.
    #>
    param(
        [scriptblock]$ScriptBlock,
        [int]$TimeoutSeconds = 10,
        $Default = $null
    )
    $job = Start-Job -ScriptBlock $ScriptBlock
    $completed = Wait-Job $job -Timeout $TimeoutSeconds
    if ($completed) {
        $result = Receive-Job $job
        Remove-Job $job -Force
        return $result
    } else {
        Stop-Job $job -ErrorAction SilentlyContinue
        Remove-Job $job -Force -ErrorAction SilentlyContinue
        return $Default
    }
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
        Finds and closes stale cowork-vm entries in HCS via hcsdiag.
        Returns the number of VMs closed.
    #>
    param(
        [string]$Action = "close",   # "close" or "kill"
        [switch]$KeepOne             # Keep one instance (the active one) -- only close extras
    )
    $closed = 0
    try {
        $hcsList = Invoke-HcsDiag -Arguments "list"
        if (-not $hcsList) { return 0 }
        if ($hcsList -notmatch "cowork-vm") { return 0 }
        # Parse GUID + name pairs
        $guidPattern = '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})'
        $entries = @()
        $currentGuid = $null
        foreach ($line in ($hcsList -split "`r?`n")) {
            if ($line -match "^\s*$guidPattern\s*$") {
                $currentGuid = $Matches[1]
            } elseif ($currentGuid -and $line -match "cowork-vm") {
                $entries += $currentGuid
                $currentGuid = $null
            } elseif ($line -match "^\s*$guidPattern") {
                # New GUID line without matching cowork-vm for previous
                $currentGuid = $Matches[1]
            }
        }
        # If KeepOne, skip the last entry (most likely the active one)
        if ($KeepOne -and $entries.Count -le 1) { return 0 }
        $toClose = if ($KeepOne) { $entries[0..($entries.Count - 2)] } else { $entries }
        foreach ($guid in $toClose) {
            try {
                Invoke-HcsDiag -Arguments $Action,$guid | Out-Null
                $closed++
            } catch {}
        }
    } catch {}
    return $closed
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
$fixMutexName = "Global\ClaudeDesktopFix_v4.8"
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

# -- Close mode: clean shutdown only, no relaunch ----------------
if ($Close) {
    Log "CLOSE MODE -- performing clean shutdown" -Colour Yellow
    Log ""
    # 1) Stop the service -- this triggers graceful VM shutdown
    Log "[1/5] Stopping CoworkVMService (graceful)..." -Colour Yellow
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    $serviceWasStopped = $false
    if ($svc -and $svc.Status -eq "Running") {
        if ($script:IsAdmin) {
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $svcCtl = New-Object System.ServiceProcess.ServiceController($ServiceName)
            try { $svcCtl.Stop() } catch {}
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
                Log "Service still running after ${maxWait}s -- force-killing" -Colour DarkYellow -Indent
                Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue | Stop-Process -Force
                Start-Sleep -Seconds 2
            }
            $serviceWasStopped = $true
        } else {
            Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue | Stop-Process -Force
            Log "Killed service process (no admin)" -Colour DarkGray -Indent
            $serviceWasStopped = $true
        }
    } else {
        Log "Service not running" -Colour DarkGray -Indent
    }
    # 2) Clean up any remaining HCS compute systems
    Log "[2/5] Cleaning HCS compute systems..." -Colour Yellow
    if ($script:IsAdmin) {
        Start-Sleep -Seconds 2
        try {
            $cleaned = Close-StaleHcsVms -Action "close"
            if ($cleaned -gt 0) {
                Log "Closed $cleaned remaining compute system(s)" -Colour Green -Indent
            } else {
                Log "No remaining compute systems" -Colour DarkGray -Indent
            }
        } catch {
            Log "HCS cleanup error: $($_.Exception.Message)" -Colour DarkYellow -Indent
        }
    } else {
        Log "Skipping (no admin)" -Colour DarkGray -Indent
    }
    # 3) Kill remaining Claude Desktop processes (UI)
    Log "[3/5] Terminating Claude processes..." -Colour Yellow
    $procs = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
    if ($procs.Count -gt 0) {
        $procs | Stop-Process -Force
        Log "Killed $($procs.Count) Claude process(es)" -Colour Green -Indent
    } else {
        Log "No Claude processes found" -Colour DarkGray -Indent
    }
    # 4) Restart the service so it is ready for next launch
    #    Without this, Windows will not auto-start the service via
    #    the named pipe trigger (manually-stopped services are ignored).
    Log "[4/5] Restarting service (idle, ready for next launch)..." -Colour Yellow
    if ($serviceWasStopped -and $script:IsAdmin) {
        try {
            Start-Service -Name $ServiceName -ErrorAction Stop
            $svcPoll = 0
            $svcOk = $false
            while ($svcPoll -lt 15) {
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
                Log "Service did not reach Running state -- may need Fix on relaunch" -Colour DarkYellow -Indent
            }
        } catch {
            Log "Service restart failed: $($_.Exception.Message)" -Colour DarkYellow -Indent
            Log "You may need to run Fix on next launch" -Colour DarkGray -Indent
        }
    } elseif (-not $script:IsAdmin) {
        Log "Skipping (no admin)" -Colour DarkGray -Indent
    } else {
        Log "Skipping (service was not stopped by us)" -Colour DarkGray -Indent
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
    } catch {}
    $svcRunning = $remainingSvc -and $remainingSvc.Status -eq "Running"
    if ($remainingProcs.Count -eq 0 -and -not $remainingHcs) {
        if ($svcRunning) {
            Log "Clean shutdown complete (service idle, ready for relaunch)" -Colour Green -Indent
        } else {
            Log "Clean shutdown complete (service not running -- may need Fix on relaunch)" -Colour DarkYellow -Indent
        }
    } else {
        if ($remainingProcs.Count -gt 0) { Log "Warning: $($remainingProcs.Count) Claude processes still running" -Colour DarkYellow -Indent }
        if ($remainingHcs) { Log "Warning: HCS compute system still present" -Colour DarkYellow -Indent }
    }
    Write-Host ""
    if ($svcRunning) {
        Write-Host "  Claude Desktop is shut down." -ForegroundColor Green
        Write-Host "  Service is idle and ready -- relaunch should work immediately." -ForegroundColor Green
    } else {
        Write-Host "  Claude Desktop is shut down." -ForegroundColor Green
        Write-Host "  Service is not running -- you may need to run Fix after relaunch." -ForegroundColor DarkYellow
    }
    Write-Host ""
    Save-Log
    if ($fixMutex) { try { $fixMutex.ReleaseMutex(); $fixMutex.Dispose() } catch {} }
    if (-not $Quiet) {
        Write-Host "  Press any key to close..." -ForegroundColor DarkGray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
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
            Write-Host "  1) Quick Fix     -- Restart services + basic repair"
            Write-Host "  2) Deep Fix      -- Full nuclear reset (cache purge)"
            Write-Host "  3) Smart Fix     -- Quick first, escalate if needed (recommended)"
            Write-Host "  4) Diagnostic    -- Health check only, no changes"
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
    $confirm = Read-Host "  Proceed? (Y/n)"
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
    } catch {}

    # HCS state via hcsdiag (v4.8.0)
    if ($script:IsAdmin) {
        try {
            $diagHcsList = Invoke-HcsDiag -Arguments "list"
            if ($diagHcsList) {
                $vmEntries = ([regex]::Matches($diagHcsList, "cowork-vm")).Count
                if ($vmEntries -gt 0) {
                    Log "  HCS VMs       : $vmEntries cowork-vm instance(s)" -Colour $(if ($vmEntries -gt 1) { "Yellow" } else { "Green" })
                } else {
                    Log "  HCS VMs       : None" -Colour DarkGray
                }
            } else {
                Log "  HCS VMs       : Unable to query (timeout or not available)" -Colour DarkGray
            }
        } catch {}
    }

    # 0xC037010D frequency (v4.8.0)
    try {
        $diagShutdownFilter = @{
            LogName   = "Microsoft-Windows-Hyper-V-Compute-Operational"
            StartTime = (Get-Date).AddHours(-24)
        }
        $diagShutdownEvents = @(Get-WinEvent -FilterHashtable $diagShutdownFilter -ErrorAction SilentlyContinue |
            Where-Object { $_.Message -match "0xC037010D" })
        $last1h = @($diagShutdownEvents | Where-Object { $_.TimeCreated -gt (Get-Date).AddHours(-1) }).Count
        $last24h = $diagShutdownEvents.Count
        $statusColour = if ($last1h -gt 10) { "Red" } elseif ($last1h -gt 3) { "Yellow" } else { "Green" }
        Log "  Shutdown fails: $last1h (1h) / $last24h (24h)" -Colour $statusColour
    } catch {}

    # 0x800707DE frequency (v4.8.0)
    try {
        $diagConstructEvents = @(Get-WinEvent -FilterHashtable $diagShutdownFilter -ErrorAction SilentlyContinue |
            Where-Object { $_.Message -match "0x800707DE" })
        if ($diagConstructEvents.Count -gt 0) {
            Log "  Construct fails: $($diagConstructEvents.Count) (24h)" -Colour Red
        }
    } catch {}

    # Session file count (v4.8.0)
    $sessionDir = Join-Path $env:APPDATA "Claude\local-agent-mode-sessions"
    if (Test-Path $sessionDir) {
        $sessionFiles = @(Get-ChildItem $sessionDir -Recurse -File -ErrorAction SilentlyContinue)
        $sessionSize = ($sessionFiles | Measure-Object -Property Length -Sum).Sum
        $sessionCol = if ($sessionFiles.Count -gt 1000) { "Red" } elseif ($sessionFiles.Count -gt 500) { "Yellow" } else { "Green" }
        Log "  Session files : $($sessionFiles.Count) ($([math]::Round($sessionSize/1MB,1)) MB)" -Colour $sessionCol
    }

    # vmcompute handle count (v4.8.0)
    try {
        $diagVmcompute = Get-Process -Name "vmcompute" -ErrorAction SilentlyContinue
        if ($diagVmcompute) {
            $hc = $diagVmcompute.HandleCount
            $hcCol = if ($hc -gt 10000) { "Red" } elseif ($hc -gt 5000) { "Yellow" } else { "Green" }
            Log "  vmcompute     : $hc handles" -Colour $hcCol
        }
    } catch {}

    # CoworkVMService recovery config (v4.8.0)
    try {
        $svcRecovery = & sc.exe qfailure CoworkVMService 2>&1
        if ($svcRecovery -match "RESTART") {
            Log "  Svc recovery  : Configured" -Colour Green
        } else {
            Log "  Svc recovery  : NOT CONFIGURED (run Prevent to fix)" -Colour Yellow
        }
    } catch {}

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
    } catch {}

    Write-Host ""
    Log "Diagnostic complete -- no changes were made" -Colour Green
    Save-Log

    if (-not $Quiet) {
        Write-Host ""
        Write-Host "  Press any key to close..." -ForegroundColor DarkGray
        try { [Win32Window]::Flash() } catch {}
        [void][System.Console]::ReadKey($true)
        try { [Win32Window]::StopFlash() } catch {}
    }
    if ($fixMutex) { try { $fixMutex.ReleaseMutex(); $fixMutex.Dispose() } catch {} }
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
    Log "=== ClaudeFix Boot Prep v$Version ===" -Colour Cyan
    Log "[BootPrep] Non-destructive boot preparation (30s post-logon)" -Colour DarkGray
    if (-not $script:IsAdmin) {
        Log "[BootPrep] Not running as admin -- cannot restart vmcompute" -Colour Yellow
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
        Log "[BootPrep] vmcompute not running after 30s -- cannot prepare" -Colour Yellow
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
                $claudeRunning = @(Get-Process -Name "claude" -ErrorAction SilentlyContinue).Count -gt 0
                if ($claudeRunning) {
                    $activeWorkspace = $true
                    Log "[BootPrep] Active workspace detected (Claude running + cowork-vm in HCS)" -Colour DarkGray
                } else {
                    # Stale VM from previous session -- clean it up
                    Log "[BootPrep] Stale cowork-vm found in HCS -- cleaning" -Colour DarkYellow
                    try {
                        $cleaned = Close-StaleHcsVms -Action "close"
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
        Log "[BootPrep] Boot prep complete -- no action needed" -Colour Green
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
            Log "[BootPrep] vmcompute restarted successfully -- ready for Claude" -Colour Green
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
    # Check ProgramData first (v5.1.0), fall back to AppData
    if (-not $isActive) {
        $safetyLogCandidates = @(
            (Join-Path "C:\ProgramData\Claude\Logs" "coworkd.log"),
            (Join-Path $ClaudeLogDir "cowork_vm_node.log"),
            (Join-Path $ClaudeLogDir "coworkd.log")
        )
        foreach ($vmLog in $safetyLogCandidates) {
            if (Test-Path $vmLog) {
                $ageSec = ((Get-Date) - (Get-Item $vmLog).LastWriteTime).TotalSeconds
                if ($ageSec -lt 120) { $isActive = $true; break }
            }
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

$vmReady = $false
$script:SmartWorkspaceEscalated = $false

# ====================================================================
# STEP 0 -- Pre-emptive HCS state cleanup (v4.8.0)
# ====================================================================
Log "[0/10] Checking for stale HCS compute systems..." -Colour Yellow
if ($script:IsAdmin) {
    try {
        $hcsList = Invoke-HcsDiag -Arguments "list"
        if (-not $hcsList) {
            Log "hcsdiag unavailable or timed out -- skipping HCS cleanup" -Colour DarkGray -Indent
        } elseif ($hcsList -match "cowork-vm") {
            Log "Found stale cowork-vm in HCS -- cleaning up" -Colour DarkYellow -Indent
            $cleaned = Close-StaleHcsVms -Action "close"
            if ($cleaned -gt 0) {
                Log "Closed $cleaned stale HCS compute system(s)" -Colour Green -Indent
            }
        } else {
            Log "HCS state clean -- no stale cowork-vm found" -Colour Green -Indent
        }
    } catch {
        Log "HCS cleanup failed (non-critical): $($_.Exception.Message)" -Colour DarkGray -Indent
    }
} else {
    Log "Skipped (requires admin)" -Colour DarkGray -Indent
}
Start-Sleep -Seconds 1

# Clean up old session files to prevent accumulation (>7 days)
# These accumulate over days and Watch flags them as "critical" at 1000+
$sessionDir = Join-Path $ClaudeAppData "local-agent-mode-sessions"
if (Test-Path $sessionDir) {
    try {
        $cutoff = (Get-Date).AddDays(-7)
        $oldFiles = @(Get-ChildItem $sessionDir -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt $cutoff })
        if ($oldFiles.Count -gt 0) {
            $sizeMB = [math]::Round(($oldFiles | Measure-Object -Property Length -Sum).Sum / 1MB, 1)
            $oldFiles | Remove-Item -Force -ErrorAction SilentlyContinue
            Log "Cleaned $($oldFiles.Count) session files older than 7 days ($sizeMB MB)" -Colour Green -Indent
            # Remove empty directories
            Get-ChildItem $sessionDir -Directory -ErrorAction SilentlyContinue |
                Where-Object { @(Get-ChildItem $_.FullName -Recurse -File -ErrorAction SilentlyContinue).Count -eq 0 } |
                Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch {
        Log "Session cleanup failed (non-critical): $($_.Exception.Message)" -Colour DarkGray -Indent
    }
}

# ====================================================================
# STEP 1 -- Kill all Claude processes
# ====================================================================
Log "[1/10] Terminating Claude processes..." -Colour Yellow

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
Log "[2/10] Stopping $ServiceName..." -Colour Yellow

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
            } catch {}
        }
        if (-not $stopped) {
            Log "Force-killing $ServiceExe (last resort after ${ServiceTimeout}s)..." -Colour DarkYellow -Indent
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
Log "[3/10] Checking HCS service health..." -Colour Yellow

try {
    $hcsDetected = Test-RecentHcsErrors
    if ($hcsDetected -eq "shutdown_stale") {
        Log "HCS shutdown failures (0xC037010D) -- stale state from property query bug" -Colour DarkYellow -Indent
        Log "vmcompute restart will clear this (same recovery as construct failure)" -Colour DarkGray -Indent
    }
    if ($hcsDetected -eq "construct_failure") {
        Log "HCS construct failure (0x800707DE) -- stale state from failed shutdowns" -Colour DarkYellow -Indent
    }
    if ($hcsDetected -eq "guest_connect_failure") {
        Log "Guest connection failure -- isGuestConnected RPC timing out" -Colour DarkYellow -Indent
        Log "Service restart will clear stale guest state" -Colour DarkGray -Indent
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
                    Log "vmcompute not running after 15s -- escalating" -Colour DarkYellow -Indent
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
                        Log "vmms service not found -- skipping" -Colour DarkGray -Indent
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
                            Log "HvHost service not found -- skipping" -Colour DarkGray -Indent
                        }
                    }
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
Log "[4/10] Checking for orphan processes..." -Colour Yellow

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
Log "[5/10] Checking for orphan compute systems..." -Colour Yellow
try {
    $orphanKilled = $false
    # Method 1: hcsdiag (most reliable for HCS compute systems)
    if ($script:IsAdmin) {
        try {
            $hcsList = Invoke-HcsDiag -Arguments "list"
            if ($hcsList -and $hcsList -match "(?i)claude|cowork") {
                Log "Found orphan compute system(s) via hcsdiag" -Colour DarkYellow -Indent
                $lines = $hcsList -split "`r?`n"
                $currentGuid = $null
                $isClaudeVm = $false
                foreach ($line in $lines) {
                    if ($line -match '^\s*([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\s*$') {
                        if ($isClaudeVm -and $currentGuid) {
                            if ($PSCmdlet.ShouldProcess($currentGuid, "hcsdiag kill")) {
                                Invoke-HcsDiag -Arguments "kill",$currentGuid | Out-Null
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
                        Invoke-HcsDiag -Arguments "kill",$currentGuid | Out-Null
                        Log "Killed orphan compute system: $currentGuid" -Colour Green -Indent
                        $orphanKilled = $true
                    }
                }
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
                    Log "vmwp.exe (PID $vmwpPid) has no command line -- skipping" -Colour DarkGray -Indent
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
                            } catch {}
                        }
                        # Fallback: force-kill the process
                        if (-not $vmwpKilled) {
                            try {
                                Stop-Process -Id $vmwpPid -Force -ErrorAction Stop
                                $vmwpKilled = $true
                            } catch {
                                Log "[!] vmwp.exe (PID $vmwpPid) is unkillable -- host restart may be needed" -Colour Red -Indent
                            }
                        }
                        if ($vmwpKilled) {
                            Log "WARNING: Force-killed vmwp.exe (PID $vmwpPid, GUID $vmwpGuid) -- VHDX corruption risk" -Colour DarkYellow -Indent
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
    try {
        $fs = [System.IO.File]::OpenRead($Path)
        $buf = New-Object byte[] 4
        $fs.Seek(65536, 'Begin') | Out-Null
        $read = $fs.Read($buf, 0, 4)
        $fs.Close()
        return ($read -eq 4 -and $buf[0] -eq 0x68 -and $buf[1] -eq 0x65 -and $buf[2] -eq 0x61 -and $buf[3] -eq 0x64)
    } catch { return $false }
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
                $cleaned = Close-StaleHcsVms -Action "close"
                if ($cleaned -gt 0) {
                    Log "HCS: closed $cleaned stale cowork-vm(s)" -Colour Green -Indent
                }
            }
        } catch {
            Log "HCS deep cleanup failed (non-critical): $($_.Exception.Message)" -Colour DarkGray -Indent
        }
    }

    # Phase 0b: Clean old session conversation logs (>7 days)
    $sessionLogDir = Join-Path $env:APPDATA "Claude\local-agent-mode-sessions"
    if (Test-Path $sessionLogDir) {
        try {
            $oldLogs = Get-ChildItem $sessionLogDir -Recurse -File -ErrorAction SilentlyContinue |
                       Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-7) }
            if ($oldLogs.Count -gt 0) {
                $oldSize = ($oldLogs | Measure-Object -Property Length -Sum).Sum
                $oldLogs | Remove-Item -Force -ErrorAction SilentlyContinue
                Log "Cleaned $($oldLogs.Count) old session logs ($([math]::Round($oldSize/1MB,1)) MB)" -Colour Green -Indent
            }
        } catch {
            Log "Session log cleanup failed (non-critical): $($_.Exception.Message)" -Colour DarkGray -Indent
        }
    }

    # Phase 1: Backup sessiondata.vhdx and smol-bin.vhdx
    $backupDir = Join-Path $LogDir "vhdx-backup"
    if (-not (Test-Path $backupDir)) { New-Item $backupDir -ItemType Directory -Force | Out-Null }

    $drive = (Get-Item $backupDir).PSDrive
    $needed = 720MB  # 580 + 36 + margin
    $free = (Get-PSDrive $drive.Name).Free
    $spaceOk = $free -ge $needed

    if (-not $spaceOk) {
        Log "Insufficient disk space for VHDX backup ($([math]::Round($free/1MB,0)) MB free, need $([math]::Round($needed/1MB,0)) MB) -- doing full nuke" -Colour DarkYellow -Indent
    }

    $vhdxFiles = @('sessiondata.vhdx', 'smol-bin.vhdx')
    $vhdxBackedUp = @{}
    $cacheDirs = @($VmCachePath, $BundlePath)

    # Verify service process is fully gone before touching VHDX files
    $handleWait = 0
    while ($handleWait -lt 6) {
        $svcProc = Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue
        if (-not $svcProc) { break }
        Start-Sleep -Seconds 1
        $handleWait++
    }
    if ($handleWait -gt 0) {
        if ($svcProc) {
            Log "Service process still running after ${handleWait}s -- VHDX files may be locked" -Colour DarkYellow -Indent
        } else {
            Log "Service process exited after ${handleWait}s" -Colour DarkGray -Indent
        }
    }

    if ($spaceOk) {
        foreach ($vhdx in $vhdxFiles) {
            # Find the file in BundlePath or VmCachePath
            $src = $null
            foreach ($cd in $cacheDirs) {
                if (-not $cd -or -not (Test-Path $cd)) { continue }
                $candidate = Get-ChildItem $cd -Recurse -Filter $vhdx -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($candidate) { $src = $candidate.FullName; break }
            }
            if ($src -and (Test-Path $src)) {
                $tmp   = Join-Path $backupDir "$vhdx.tmp"
                $final = Join-Path $backupDir $vhdx
                try {
                    if ($PSCmdlet.ShouldProcess($vhdx, "Backup")) {
                        Copy-Item $src $tmp -Force -ErrorAction Stop
                        # Validate VHDX header
                        if (Test-VhdxHeader $tmp) {
                            if (Test-Path $final) { Remove-Item $final -Force -ErrorAction SilentlyContinue }
                            Rename-Item $tmp $vhdx -ErrorAction Stop
                            $vhdxBackedUp[$vhdx] = $src  # remember original location
                            Log "Backed up $vhdx ($([math]::Round((Get-Item $final).Length/1MB,1)) MB)" -Colour Green -Indent
                        } else {
                            Log "WARNING: $vhdx backup has invalid VHDX header -- skipping" -Colour DarkYellow -Indent
                            Remove-Item $tmp -Force -ErrorAction SilentlyContinue
                        }
                    }
                } catch {
                    Log "[!] Failed to backup $vhdx : $($_.Exception.Message)" -Colour Red -Indent
                    if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
                }
            }
        }
    }

    # Phase 2: Nuke -- delete bulk VM files, keep session data if backed up
    foreach ($item in @(
        @{ Path = $VmCachePath; Label = "claude-code-vm" },
        @{ Path = $BundlePath;  Label = "vm_bundles" }
    )) {
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

    # Phase 3: Restore backed-up VHDXs (after service restart in Step 7)
    # Deferred -- see $vhdxBackedUp usage after Step 7 below.

    # Phase 4: MSIX smol-bin fallback (if backup doesn't exist or is corrupt)
    if (-not $vhdxBackedUp.ContainsKey('smol-bin.vhdx')) {
        try {
            $pkg = Get-AppxPackage | Where-Object { $_.Name -eq 'Claude' -or $_.PackageFamilyName -like 'Claude_*' } | Select-Object -First 1
            if ($pkg) {
                $msixSmolBin = Join-Path $pkg.InstallLocation 'resources\app\claudevm.bundle\smol-bin.vhdx'
                if (Test-Path $msixSmolBin) {
                    $final = Join-Path $backupDir 'smol-bin.vhdx'
                    Copy-Item $msixSmolBin $final -Force -ErrorAction Stop
                    $vhdxBackedUp['smol-bin.vhdx'] = $null  # no original location -- will use bundle path
                    Log "Recovered smol-bin.vhdx from MSIX package" -Colour Green -Indent
                }
            }
        } catch {
            Log "MSIX smol-bin recovery skipped: $($_.Exception.Message)" -Colour DarkGray -Indent
        }
    }
}

# -- Temp file cleanup (Change 5) --
try {
    $tempPatterns = @("$env:TEMP\anthropic-*", "$env:TEMP\claude-*")
    $tempCleaned = 0
    $tempBytes = 0
    foreach ($pattern in $tempPatterns) {
        Get-ChildItem -Path $pattern -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            $tempBytes += $_.Length
            if ($PSCmdlet.ShouldProcess($_.FullName, "Remove temp file")) {
                Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
            }
            $tempCleaned++
        }
    }
    if ($tempCleaned -gt 0) {
        $freed = if ($tempBytes -gt 1MB) { "{0:N1} MB" -f ($tempBytes/1MB) } else { "{0:N0} KB" -f ($tempBytes/1KB) }
        Log "Cleaned $tempCleaned temp items ($freed freed)" -Colour DarkGray -Indent
    }
} catch {}

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
                }
                Log "Cleaned traditional install path: $tp ($count items)" -Colour DarkGray -Indent
            }
        }
    }
} catch {}

# ====================================================================
# STEP 7 -- Restart CoworkVMService (with extended polling)
# ====================================================================
Log "[7/10] Starting $ServiceName..." -Colour Yellow

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

# -- Phase 3: Restore backed-up VHDXs after service restart --
if (-not $skipCachePurge -and $vhdxBackedUp -and $vhdxBackedUp.Count -gt 0) {
    foreach ($vhdx in $vhdxBackedUp.Keys) {
        $backup = Join-Path $backupDir $vhdx
        # Determine restore target: original location or first existing cache dir
        $target = $vhdxBackedUp[$vhdx]
        if (-not $target) {
            # smol-bin from MSIX -- pick the bundle path if it was recreated
            foreach ($cd in @($BundlePath, $VmCachePath)) {
                if (Test-Path $cd) { $target = Join-Path $cd $vhdx; break }
            }
        }
        if ($target -and (Test-Path $backup) -and -not (Test-Path $target)) {
            try {
                # Ensure parent directory exists
                $parentDir = Split-Path $target -Parent
                if (-not (Test-Path $parentDir)) { New-Item $parentDir -ItemType Directory -Force | Out-Null }
                Copy-Item $backup $target -Force -ErrorAction Stop
                Log "Restored $vhdx from backup" -Colour Green -Indent
            } catch {
                Log "[!] Failed to restore $vhdx : $($_.Exception.Message) -- service will recreate" -Colour DarkYellow -Indent
            }
        }
    }
}

# -- Smart mode escalation: if service didn't start, escalate to Deep --
if ($script:SelectedMode -eq "Smart") {
    $smartSvc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if (-not $smartSvc -or $smartSvc.Status -ne "Running") {
        Log "Smart mode: service not running after quick fix -- escalating to Deep" -Colour Yellow
        # Run the cache purge that was skipped -- with VHDX backup
        Log "[7/10] Escalated cache purge..." -Colour Yellow
        # Phase 0: HCS cleanup
        if ($script:IsAdmin) {
            try { Close-StaleHcsVms -Action "close" | Out-Null } catch {}
        }
        # Phase 1: Backup VHDX before nuke
        $escalateBackupDir = Join-Path $LogDir "vhdx-backup"
        if (-not (Test-Path $escalateBackupDir)) { New-Item $escalateBackupDir -ItemType Directory -Force | Out-Null }
        # Verify service process is fully gone before touching VHDX files
        $handleWait = 0
        while ($handleWait -lt 6) {
            $svcProc = Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue
            if (-not $svcProc) { break }
            Start-Sleep -Seconds 1
            $handleWait++
        }
        if ($handleWait -gt 0) {
            if ($svcProc) {
                Log "Service process still running after ${handleWait}s -- VHDX files may be locked" -Colour DarkYellow -Indent
            } else {
                Log "Service process exited after ${handleWait}s" -Colour DarkGray -Indent
            }
        }
        foreach ($vhdx in @('sessiondata.vhdx', 'smol-bin.vhdx')) {
            foreach ($cd in @($VmCachePath, $BundlePath)) {
                if (-not $cd -or -not (Test-Path $cd)) { continue }
                $src = Get-ChildItem $cd -Recurse -Filter $vhdx -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($src) {
                    $dest = Join-Path $escalateBackupDir $vhdx
                    try {
                        Copy-Item $src.FullName $dest -Force -ErrorAction Stop
                        Log "Backed up $vhdx before escalation" -Colour Green -Indent
                    } catch {}
                    break
                }
            }
        }
        # Phase 2: Nuke
        foreach ($item in @(
            @{ Path = $VmCachePath; Label = "claude-code-vm" },
            @{ Path = $BundlePath;  Label = "vm_bundles" }
        )) {
            if (Test-Path $item.Path) {
                $size = (Get-ChildItem $item.Path -Recurse -ErrorAction SilentlyContinue |
                         Measure-Object -Property Length -Sum).Sum
                $sizeMB = [math]::Round($size / 1MB, 1)
                if ($PSCmdlet.ShouldProcess($item.Label, "Delete ($sizeMB MB)")) {
                    Remove-Item $item.Path -Recurse -Force -ErrorAction SilentlyContinue
                }
                Log "$($item.Label) removed ($sizeMB MB freed)" -Colour Green -Indent
            }
        }
        # Retry service start
        Log "Retrying service start after deep purge..." -Colour Yellow -Indent
        if ($script:IsAdmin) {
            $svcOk2 = Restart-CoworkService
            if ($svcOk2) {
                Log "Service running after escalation" -Colour Green -Indent
            } else {
                Log "[!] Service still failed after deep purge" -Colour Red -Indent
            }
        }
        # Phase 3: Restore VHDX after service restart
        foreach ($vhdx in @('sessiondata.vhdx', 'smol-bin.vhdx')) {
            $backup = Join-Path $escalateBackupDir $vhdx
            if (Test-Path $backup) {
                foreach ($cd in @($BundlePath, $VmCachePath)) {
                    if (Test-Path $cd) {
                        $target = Join-Path $cd $vhdx
                        if (-not (Test-Path $target)) {
                            try {
                                Copy-Item $backup $target -Force -ErrorAction Stop
                                Log "Restored $vhdx from escalation backup" -Colour Green -Indent
                            } catch {}
                        }
                        break
                    }
                }
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
            $preLaunchCleaned = Close-StaleHcsVms -Action "close"
            if ($preLaunchCleaned -gt 0) {
                Log "Pre-launch: cleaned $preLaunchCleaned stale HCS compute system(s)" -Colour Green -Indent
                Start-Sleep -Seconds 2
            }
        } catch {}
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

    # Pick the most recently written log as primary (v5.1.0)
    # ProgramData is where active logs live since ~March 2026;
    # AppData paths are kept as fallbacks for older installs.
    $vmLogDirPD  = "C:\ProgramData\Claude\Logs"
    $vmLogDirAD  = Join-Path $ClaudeAppData "logs"
    $coworkdLogPD = Join-Path $vmLogDirPD "coworkd.log"
    $coworkdLogAD = Join-Path $vmLogDirAD "coworkd.log"
    $vmNodeLogAD  = Join-Path $vmLogDirAD "cowork_vm_node.log"

    $vmLogFile = $null
    $vmLogCandidates = @($coworkdLogPD, $vmNodeLogAD, $coworkdLogAD) |
        Where-Object { Test-Path $_ } |
        Sort-Object { (Get-Item $_).LastWriteTime } -Descending
    if ($vmLogCandidates) {
        $vmLogFile = $vmLogCandidates[0]
        Log "Using log file: $vmLogFile" -Colour DarkGray -Indent
    }

    # Record the log file size at the START so we only check new entries
    $logBaselineSize = 0
    if ($vmLogFile -and (Test-Path $vmLogFile)) {
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
                if ($reader) { try { $reader.Close() } catch {} }
                if ($stream) { try { $stream.Close() } catch {} }
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

    # Helper: check HCS compute system state via hcsdiag (v4.8.0)
    # Uses Invoke-WithTimeout to avoid hangs when vmcompute is unstable (v5.0.0)
    function Test-HyperVReady {
        try {
            $hcsList = Invoke-HcsDiag -Arguments "list"
            if (-not $hcsList) { return $null }
            if ($hcsList -match "cowork-vm") {
                # Also check Integration Services heartbeat if available
                try {
                    $claudeVm = Invoke-WithTimeout -TimeoutSeconds 8 -ScriptBlock {
                        Get-VM -ErrorAction SilentlyContinue |
                        Where-Object { $_.Name -match "claude" } |
                        Select-Object -First 1
                    }
                    if ($claudeVm -and $claudeVm.State -eq "Running") {
                        $hb = Get-VMIntegrationService -VMName $claudeVm.Name -Name "Heartbeat" -ErrorAction SilentlyContinue
                        if ($hb -and $hb.PrimaryStatusDescription -eq "OK") {
                            return "running+heartbeat"
                        }
                    }
                } catch {}
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
    $script:NoProgressEscalated = $false

    # Skip Hyper-V cmdlet checks if vmcompute was just restarted (may hang on unstable service)
    $skipHvChecks = [bool]$hcsDetected
    if ($skipHvChecks) {
        Log "Skipping Hyper-V VM checks (vmcompute was just restarted)" -Colour DarkGray -Indent
    }

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
                    $cleaned = Close-StaleHcsVms -Action "close"
                    if ($cleaned -gt 0) {
                        Log "Closed $cleaned stale HCS compute system(s)" -Colour Green -Indent
                    }
                } catch {}
                # Stop service with timeout (v5.1.0)
                $stopJob = Start-Job -ScriptBlock {
                    param($svc)
                    Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
                } -ArgumentList $ServiceName
                $stopDone = Wait-Job $stopJob -Timeout 30
                if (-not $stopDone) {
                    Stop-Job $stopJob -ErrorAction SilentlyContinue
                    Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue | Stop-Process -Force
                    Log "Service stop timed out -- force-killed" -Colour DarkYellow -Indent
                }
                Remove-Job $stopJob -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 3
                Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 5
                # Reset baseline so we check fresh log entries
                if (Test-Path $vmLogFile) {
                    $logBaselineSize = (Get-Item $vmLogFile -ErrorAction SilentlyContinue).Length
                }
                Log "Service restarted -- continuing to wait..." -Colour DarkGray -Indent
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
            } catch {}
            if ($hasHcsVm) {
                Log "HCS VM present but no log/guest activity -- VM may be stuck" -Colour DarkGray -Indent
            }
            if (-not $hasLogActivity -and -not $hasGuestActivity) {
                if (-not $script:NoProgressEscalated) {
                    $script:NoProgressEscalated = $true
                    Log "[!] No progress after ${vmElapsed}s -- no log activity, no guest polling" -Colour Red -Indent
                    Log "Escalating: killing Claude, cleaning HCS, relaunching..." -Colour Yellow -Indent
                    # Kill Claude
                    Get-Process -Name "claude" -ErrorAction SilentlyContinue | Stop-Process -Force
                    Start-Sleep -Seconds 2
                    # Clean HCS
                    try { Close-StaleHcsVms -Action "close" | Out-Null } catch {}
                    Start-Sleep -Seconds 2
                    # Restart cowork service (with timeout, v5.1.0)
                    $stopJob = Start-Job -ScriptBlock {
                        param($svc)
                        Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
                    } -ArgumentList $ServiceName
                    $stopDone = Wait-Job $stopJob -Timeout 30
                    if (-not $stopDone) {
                        Stop-Job $stopJob -ErrorAction SilentlyContinue
                        Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue | Stop-Process -Force
                        Log "Service stop timed out -- force-killed" -Colour DarkYellow -Indent
                    }
                    Remove-Job $stopJob -Force -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 3
                    Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 5
                    # Relaunch Claude (reuse the scheduled task if available)
                    try {
                        $taskExists = Get-ScheduledTask -TaskName "LaunchClaudeAdmin" -TaskPath "\Claude\" -ErrorAction SilentlyContinue
                        if ($taskExists) {
                            Start-ScheduledTask -TaskName "LaunchClaudeAdmin" -TaskPath "\Claude\"
                            Log "Relaunched Claude via scheduled task" -Colour Green -Indent
                        } elseif ($script:CapturedClaudeExe) {
                            Start-Process $script:CapturedClaudeExe -ErrorAction SilentlyContinue
                            Log "Relaunched Claude directly" -Colour Green -Indent
                        }
                    } catch {}
                    # Reset log baseline and timers
                    if (Test-Path $vmLogFile) {
                        $logBaselineSize = (Get-Item $vmLogFile -ErrorAction SilentlyContinue).Length
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
        if (-not $logStatus -and $vmElapsed -ge 30 -and (Test-Path $vmLogFile)) {
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
                        if ($reader) { try { $reader.Close() } catch {} }
                        if ($stream) { try { $stream.Close() } catch {} }
                    }
                    if ($tailContent -match "Startup complete|\[Keepalive\]|started PID|full egress mode") {
                        $vmReady = $true
                        Log "Workspace ready (log tail: recently active with boot markers)" -Colour Green -Indent
                        break
                    }
                }
            } catch {}
        }

        # B. Check Hyper-V (once, to see if cmdlets work)
        if (-not $hvChecked) {
            $hvChecked = $true
            $hvState = if ($skipHvChecks) { $null } else { Test-HyperVReady }
            if ($null -ne $hvState) {
                $hvAvailable = $true
                Log "Hyper-V VM detected: $hvState" -Colour DarkGray -Indent
            }
        } elseif ($hvAvailable -and -not $skipHvChecks -and $vmElapsed -ge 20) {
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
            Log "Workspace partially ready (log: $finalLog) -- may still be loading" -Colour Yellow -Indent
        } else {
            Log "[!] Workspace not confirmed ready after ${vmTimeout}s" -Colour Yellow -Indent
            Log "Open a Cowork session in Claude to trigger setup" -Colour DarkGray -Indent
        }
    }
}

# -- Smart mode: escalate to Deep if workspace didn't come up (v5.0.0) --
if (-not $vmReady -and $script:SelectedMode -eq "Smart" -and -not $script:SmartWorkspaceEscalated) {
    $script:SmartWorkspaceEscalated = $true
    Log "" -Colour White
    Log "Smart mode: workspace not ready after quick fix -- escalating to Deep" -Colour Yellow
    # Phase 0: Kill Claude + HCS cleanup
    Get-Process -Name "claude" -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2
    try { Close-StaleHcsVms -Action "close" | Out-Null } catch {}
    # Phase 1: Stop service (with timeout, v5.1.0)
    $stopJob = Start-Job -ScriptBlock {
        param($svc)
        Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
    } -ArgumentList $ServiceName
    $stopDone = Wait-Job $stopJob -Timeout 30
    if (-not $stopDone) {
        Stop-Job $stopJob -ErrorAction SilentlyContinue
        Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue | Stop-Process -Force
        Log "Service stop timed out -- force-killed" -Colour DarkYellow -Indent
    }
    Remove-Job $stopJob -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3
    # Phase 2: Cache purge with VHDX backup
    $escalateBackupDir = Join-Path $LogDir "vhdx-backup"
    if (-not (Test-Path $escalateBackupDir)) {
        New-Item $escalateBackupDir -ItemType Directory -Force | Out-Null
    }
    # Verify service process is fully gone before touching VHDX files
    $handleWait = 0
    while ($handleWait -lt 6) {
        $svcProc = Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue
        if (-not $svcProc) { break }
        Start-Sleep -Seconds 1
        $handleWait++
    }
    if ($handleWait -gt 0) {
        if ($svcProc) {
            Log "Service process still running after ${handleWait}s -- VHDX files may be locked" -Colour DarkYellow -Indent
        } else {
            Log "Service process exited after ${handleWait}s" -Colour DarkGray -Indent
        }
    }
    foreach ($vhdx in @('sessiondata.vhdx', 'smol-bin.vhdx')) {
        foreach ($cd in @($VmCachePath, $BundlePath)) {
            if (-not $cd -or -not (Test-Path $cd)) { continue }
            $src = Get-ChildItem $cd -Recurse -Filter $vhdx -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($src) {
                $dest = Join-Path $escalateBackupDir $vhdx
                try { Copy-Item $src.FullName $dest -Force -ErrorAction Stop } catch {}
                break
            }
        }
    }
    # Nuke VM cache
    foreach ($item in @(
        @{ Path = $VmCachePath; Label = "claude-code-vm" },
        @{ Path = $BundlePath;  Label = "vm_bundles" }
    )) {
        if (Test-Path $item.Path) {
            $size = (Get-ChildItem $item.Path -Recurse -ErrorAction SilentlyContinue |
                     Measure-Object -Property Length -Sum).Sum
            $sizeMB = [math]::Round($size / 1MB, 1)
            Remove-Item $item.Path -Recurse -Force -ErrorAction SilentlyContinue
            Log "$($item.Label) removed ($sizeMB MB freed)" -Colour Green -Indent
        }
    }
    # Phase 3: Restart service + relaunch + re-enter step 9 wait
    Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 5
    # Restore VHDX
    foreach ($vhdx in @('sessiondata.vhdx', 'smol-bin.vhdx')) {
        $backup = Join-Path $escalateBackupDir $vhdx
        if (Test-Path $backup) {
            foreach ($cd in @($BundlePath, $VmCachePath)) {
                if (Test-Path $cd) {
                    $target = Join-Path $cd $vhdx
                    if (-not (Test-Path $target)) {
                        try { Copy-Item $backup $target -Force -ErrorAction Stop } catch {}
                    }
                    break
                }
            }
        }
    }
    # Relaunch Claude
    try {
        $taskExists = Get-ScheduledTask -TaskName "LaunchClaudeAdmin" -TaskPath "\Claude\" -ErrorAction SilentlyContinue
        if ($taskExists) {
            Start-ScheduledTask -TaskName "LaunchClaudeAdmin" -TaskPath "\Claude\"
            Log "Relaunched Claude via scheduled task" -Colour Green -Indent
        }
    } catch {}
    # Re-enter a shorter step 9 wait (120s this time)
    Log "Waiting for workspace after deep escalation..." -Colour Yellow -Indent
    $escalateTimeout = 120
    $escalateElapsed = 0
    if (Test-Path $vmLogFile) {
        $logBaselineSize = (Get-Item $vmLogFile -ErrorAction SilentlyContinue).Length
    }
    while ($escalateElapsed -lt $escalateTimeout) {
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
    Log "Workspace not ready -- attempting retry cycle ($MaxRetries max)..." -Colour Yellow
    for ($retryNum = 1; $retryNum -le $MaxRetries; $retryNum++) {
        Log "Retry $retryNum/$MaxRetries -- quick service restart..." -Colour Yellow -Indent
        if ($script:IsAdmin) {
            # Quick cycle: stop service, hcsdiag cleanup, restart service (with timeout, v5.1.0)
            $stopJob = Start-Job -ScriptBlock {
                param($svc)
                Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
            } -ArgumentList $ServiceName
            $stopDone = Wait-Job $stopJob -Timeout 30
            if (-not $stopDone) {
                Stop-Job $stopJob -ErrorAction SilentlyContinue
                Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue | Stop-Process -Force
                Log "Service stop timed out -- force-killed" -Colour DarkYellow -Indent
            }
            Remove-Job $stopJob -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            try {
                    $cleaned = Close-StaleHcsVms -Action "close"
                    if ($cleaned -gt 0) {
                        Log "Cleaned $cleaned stale HCS state(s)" -Colour Green -Indent
                    }
            } catch {}
            Start-Service -Name $ServiceName -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 5
            # Wait up to 60s for workspace
            $retryElapsed = 0
            while ($retryElapsed -lt 60) {
                Start-Sleep -Seconds 5
                $retryElapsed += 5
                $retryStatus = Test-VmLogReady -Baseline $logBaselineSize
                if ($retryStatus) {
                    Log "Workspace ready on retry $retryNum ($retryStatus)" -Colour Green -Indent
                    $vmReady = $true
                    break
                }
            }
            if ($vmReady) { break }
            Log "Retry $retryNum failed" -Colour DarkYellow -Indent
        } else {
            Log "Retry requires admin -- skipping" -Colour DarkGray -Indent
            break
        }
    }
    if (-not $vmReady) {
        Log "All $MaxRetries retries exhausted -- manual intervention may be needed" -Colour Red
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

} catch {
    Write-Host ""
    Write-Host "  +-------------------------------------------+" -ForegroundColor Red
    Write-Host "  |           UNEXPECTED ERROR                 |" -ForegroundColor Red
    Write-Host "  +-------------------------------------------+" -ForegroundColor Red
    Write-Host ""
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor DarkGray
    Write-Host ""
} finally {
    Save-Log
    try { Stop-Transcript -ErrorAction SilentlyContinue } catch {}
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
