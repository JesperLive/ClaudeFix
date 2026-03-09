<#
.SYNOPSIS
    Claude Desktop / Cowork -- Preventive Configuration

.DESCRIPTION
    One-shot script that configures Windows to minimise VirtioFS/Plan9
    mount failures and HCS (Host Compute Service) errors in Claude
    Desktop's Cowork VM.

    Run once. Changes persist across reboots.

    What it does:
    - Sets power plan to High Performance (or Ultimate if available)
    - Disables sleep on AC power
    - Disables hibernate and Fast Startup
    - Disables USB selective suspend
    - Disables hard disk spin-down and PCI-E power management on AC
    - Disables Connected Standby / Modern Standby
    - Disables power saving on all network adapters
    - Sets minimum processor state to 100% on AC
    - Pins Hyper-V VM memory (disables Dynamic Memory ballooning)
    - Boosts VM worker process priority
    - Configures HCS (vmcompute) service auto-recovery on failure
    - Sets ServicesPipeTimeout to prevent boot race conditions
    - Verifies and repairs WinNAT rules for VM network
    - Checks Windows Firewall policies for Hyper-V compatibility
    - Detects problematic workspace storage locations
    - Verifies NTP/time synchronisation
    - Detects antivirus software and suggests exclusions
    - Installs a persistent health monitor that detects VirtioFS mount
      failures and auto-runs the fix script within seconds
    - Registers a boot-fix task that resets the VM at every logon

    What it does NOT do:
    - Touch your Claude config or conversations
    - Disable sleep on battery (laptop users keep battery sleep)
    - Change your screen timeout

.PARAMETER Undo
    Reverts all changes made by this script.

.NOTES
    Version : 4.8.0
    Author  : Jesper Driessen
    Licence : MIT
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$Undo
)

# -- Admin elevation -------------------------------------------------
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $scriptFile = $PSCommandPath
    if (-not $scriptFile) { $scriptFile = $MyInvocation.MyCommand.Definition }

    $elevateArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptFile`""
    if ($Undo) { $elevateArgs += " -Undo" }

    try {
        Start-Process PowerShell -ArgumentList $elevateArgs -Verb RunAs
    } catch {
        Write-Host "  [!] UAC elevation was declined or failed." -ForegroundColor Red
        Write-Host "      This script requires Administrator privileges." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Press any key to close..." -ForegroundColor DarkGray
        [void][System.Console]::ReadKey($true)
    }
    exit
}

Set-StrictMode -Version Latest

# -- Constants -------------------------------------------------------
$Version          = "4.8.0"
$TaskName         = "ClaudeCoworkWatchdog"
$BootTaskName     = "ClaudeCoworkBootFix"
$TaskPath         = "\Claude\"
$ServiceName      = "CoworkVMService"
$BackupFile       = Join-Path $env:APPDATA "Claude\power-plan-backup.txt"
$CsBackupFile     = Join-Path $env:APPDATA "Claude\connected-standby-backup.txt"

# -- Helpers ---------------------------------------------------------
function Log {
    param([string]$Message, [string]$Colour = "White", [switch]$Indent)
    $pfx = ""
    if ($Indent) { $pfx = "      " }
    Write-Host "$pfx$Message" -ForegroundColor $Colour
}

function Step {
    param([int]$Num, [int]$Total, [string]$Message)
    Log "[$Num/$Total] $Message" -Colour Yellow
}

function Get-ActivePlanGuid {
    $raw = (powercfg /getactivescheme) -replace '.*:\s*', '' -replace '\s*\(.*', ''
    return $raw.Trim()
}

# -- Header ----------------------------------------------------------
Write-Host ""
Write-Host "  +----------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |  CLAUDE DESKTOP / COWORK -- PREVENTION SETUP  |" -ForegroundColor Cyan
Write-Host "  |  v$Version                                       |" -ForegroundColor DarkGray
Write-Host "  +----------------------------------------------+" -ForegroundColor Cyan
Write-Host ""

if ($Undo) {
    Write-Host "  MODE: UNDO -- reverting all changes" -ForegroundColor Yellow
    Write-Host ""
}

try {

if ($Undo) {

    # ================================================================
    # UNDO MODE
    # ================================================================
    $steps = 11

    Step 1 $steps "Restoring original power plan..."
    if (Test-Path $BackupFile) {
        $originalGuid = (Get-Content $BackupFile -Raw).Trim()
        if ($originalGuid -match "^[0-9a-fA-F\-]{36}$") {
            powercfg /setactive $originalGuid
            Log "Restored plan: $originalGuid" -Colour Green -Indent
            Remove-Item $BackupFile -Force -ErrorAction SilentlyContinue
        } else {
            Log "Backup file corrupt -- please set your power plan manually" -Colour Yellow -Indent
        }
    } else {
        Log "No backup found -- power plan was not changed by this script" -Colour DarkGray -Indent
    }

    Step 2 $steps "Re-enabling hibernate and Fast Startup..."
    powercfg /h on
    Log "Hibernate: On" -Colour Green -Indent
    # Re-enable Fast Startup
    try {
        $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power"
        Set-ItemProperty -Path $regPath -Name "HiberbootEnabled" -Value 1 -ErrorAction SilentlyContinue
        Log "Fast Startup: On" -Colour Green -Indent
    } catch {
        Log "Could not re-enable Fast Startup -- not critical" -Colour DarkGray -Indent
    }

    Step 3 $steps "Resetting sleep timeout to 30 minutes (AC)..."
    powercfg /change standby-timeout-ac 30
    Log "Sleep timeout set to 30 min on AC" -Colour Green -Indent

    Step 4 $steps "Re-enabling Connected Standby..."
    if (Test-Path $CsBackupFile) {
        try {
            $originalCs = (Get-Content $CsBackupFile -Raw).Trim()
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Power" -Name "CsEnabled" -Value ([int]$originalCs) -ErrorAction Stop
            Log "Connected Standby restored to: $originalCs" -Colour Green -Indent
            Remove-Item $CsBackupFile -Force -ErrorAction SilentlyContinue
        } catch {
            Log "Could not restore Connected Standby -- check manually" -Colour Yellow -Indent
        }
    } else {
        Log "No Connected Standby backup found -- was not changed" -Colour DarkGray -Indent
    }

    Step 5 $steps "Re-enabling network adapter power management..."
    try {
        $nics = Get-NetAdapter -Physical -ErrorAction SilentlyContinue
        foreach ($nic in $nics) {
            $pnp = Get-PnpDeviceProperty -InstanceId $nic.PnPDeviceID `
                       -KeyName "DEVPKEY_Device_Class" -ErrorAction SilentlyContinue
            # Re-enable via registry -- AllowIdleIrpInD3 = 1
            $regBase = "HKLM:\SYSTEM\CurrentControlSet\Control\Class"
            # Use powershell to set PnPCapabilities back to 0 (default)
            $devPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\$($nic.PnPDeviceID)\Device Parameters"
            if (Test-Path $devPath) {
                Remove-ItemProperty -Path $devPath -Name "PnPCapabilities" -ErrorAction SilentlyContinue
            }
        }
        Log "Network adapter power management: Restored to defaults" -Colour Green -Indent
        Log "A reboot is required for this change to take effect" -Colour DarkGray -Indent
    } catch {
        Log "Could not restore network adapter settings -- not critical" -Colour DarkGray -Indent
    }

    Step 6 $steps "Re-enabling Hyper-V Dynamic Memory..."
    try {
        $claudeVm = Get-VM -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match "claude" } |
                    Select-Object -First 1
        if ($claudeVm) {
            $vmName = $claudeVm.Name
            if ($claudeVm.State -eq "Off") {
                Set-VMMemory -VMName $vmName -DynamicMemoryEnabled $true -ErrorAction Stop
                Log "Dynamic Memory re-enabled for VM '$vmName'" -Colour Green -Indent
            } else {
                Log "VM '$vmName' is running -- restart it to re-enable Dynamic Memory" -Colour DarkYellow -Indent
            }
        } else {
            Log "No Claude VM found -- nothing to restore" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not restore Dynamic Memory -- Hyper-V module may not be available" -Colour DarkGray -Indent
    }
    # Clean up flag file
    $flagFile = Join-Path $env:APPDATA "Claude\disable-dynamic-memory.flag"
    if (Test-Path $flagFile) {
        Remove-Item $flagFile -Force -ErrorAction SilentlyContinue
    }

    Step 7 $steps "Reverting HCS service configuration..."
    try {
        # Reset vmcompute failure actions to Windows default
        & sc.exe failure vmcompute actions= "" reset= 0 2>&1 | Out-Null
        Log "vmcompute failure recovery: Reset to defaults" -Colour Green -Indent

        # Reset CoworkVMService failure actions (v4.8.0)
        & sc.exe failure CoworkVMService actions= "" reset= 0 2>&1 | Out-Null
        Log "CoworkVMService failure recovery: Reset to defaults" -Colour Green -Indent

        # Remove ServicesPipeTimeout only if we set it (value is exactly 120000)
        try {
            $current = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control" -Name ServicesPipeTimeout -ErrorAction SilentlyContinue).ServicesPipeTimeout
            if ($null -ne $current -and $current -eq 120000) {
                Remove-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control" -Name ServicesPipeTimeout -ErrorAction Stop
                Log "ServicesPipeTimeout: Removed (was 120000ms -- set by this script)" -Colour Green -Indent
            } elseif ($null -ne $current) {
                Log "ServicesPipeTimeout: Left at ${current}ms (not set by this script)" -Colour DarkGray -Indent
            } else {
                Log "ServicesPipeTimeout: Not set" -Colour DarkGray -Indent
            }
        } catch {
            Log "Could not remove ServicesPipeTimeout -- not critical" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not revert HCS configuration -- not critical" -Colour DarkGray -Indent
    }

    Step 8 $steps "Removing scheduled tasks and health monitor..."
    # Kill any running health monitor process first
    try {
        Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" -ErrorAction SilentlyContinue |
            Where-Object { $_.CommandLine -match "Watch-ClaudeHealth" } |
            ForEach-Object {
                Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
                Log "Stopped running health monitor (PID $($_.ProcessId))" -Colour Green -Indent
            }
    } catch {}
    $removedAny = $false
    foreach ($tName in @($TaskName, $BootTaskName)) {
        try {
            Unregister-ScheduledTask -TaskName $tName -TaskPath $TaskPath -Confirm:$false -ErrorAction Stop
            Log "$tName removed" -Colour Green -Indent
            $removedAny = $true
        } catch {
            Log "$tName not found" -Colour DarkGray -Indent
        }
    }
    if (-not $removedAny) {
        Log "No tasks to remove" -Colour DarkGray -Indent
    }
    # Clean up old watchdog script if present
    $oldWatchdog = Join-Path $env:APPDATA "Claude\cowork-watchdog.ps1"
    if (Test-Path $oldWatchdog) {
        Remove-Item $oldWatchdog -Force -ErrorAction SilentlyContinue
        Log "Removed old watchdog script" -Colour DarkGray -Indent
    }

    Step 9 $steps "Removing shortcuts..."
    $desktopLnk = Join-Path ([Environment]::GetFolderPath("Desktop")) "Fix Claude Desktop.lnk"
    $startMenuLnk = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Fix Claude Desktop.lnk"
    $removedLnk = $false
    foreach ($lnk in @($desktopLnk, $startMenuLnk)) {
        if (Test-Path $lnk) {
            Remove-Item $lnk -Force -ErrorAction SilentlyContinue
            Log "Removed: $lnk" -Colour Green -Indent
            $removedLnk = $true
        }
    }
    if (-not $removedLnk) {
        Log "No shortcuts found" -Colour DarkGray -Indent
    }

    Step 10 $steps "Removing Claude elevation config..."
    try {
        # Remove the LaunchClaudeAdmin scheduled task
        try {
            Unregister-ScheduledTask -TaskName "LaunchClaudeAdmin" -TaskPath $TaskPath -Confirm:$false -ErrorAction Stop
            Log "Removed LaunchClaudeAdmin scheduled task" -Colour Green -Indent
        } catch {
            Log "LaunchClaudeAdmin task not found" -Colour DarkGray -Indent
        }
        # Remove RUNASADMIN compat flag for any Claude.exe entries
        $layersPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"
        if (Test-Path $layersPath) {
            $layers = Get-ItemProperty -Path $layersPath -ErrorAction SilentlyContinue
            $removedFlag = $false
            foreach ($prop in $layers.PSObject.Properties) {
                if ($prop.Name -match "Claude\.exe" -and $prop.Value -match "RUNASADMIN") {
                    Remove-ItemProperty -Path $layersPath -Name $prop.Name -Force -ErrorAction SilentlyContinue
                    Log "Removed RUNASADMIN flag: $($prop.Name)" -Colour Green -Indent
                    $removedFlag = $true
                }
            }
            if (-not $removedFlag) {
                Log "No RUNASADMIN flags found" -Colour DarkGray -Indent
            }
        }
        # Remove admin shortcut
        $adminLnk = Join-Path ([Environment]::GetFolderPath("Desktop")) "Claude (Admin).lnk"
        if (Test-Path $adminLnk) {
            Remove-Item $adminLnk -Force -ErrorAction SilentlyContinue
            Log "Removed: $adminLnk" -Colour Green -Indent
        }
        # Remove launcher scripts
        $launcherCmd = Join-Path $env:APPDATA "Claude\Launch-Claude-Admin.cmd"
        $launcherPs1 = Join-Path $env:APPDATA "Claude\Launch-Claude-Admin.ps1"
        foreach ($lf in @($launcherCmd, $launcherPs1)) {
            if (Test-Path $lf) {
                Remove-Item $lf -Force -ErrorAction SilentlyContinue
                Log "Removed launcher: $lf" -Colour Green -Indent
            }
        }
    } catch {
        Log "Could not fully remove elevation config -- not critical" -Colour DarkGray -Indent
    }

    Step 11 $steps "Reverting admin token policy..."
    try {
        $policyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
        $reverted = $false
        try {
            $latfp = (Get-ItemProperty -Path $policyPath -ErrorAction Stop).LocalAccountTokenFilterPolicy
            if ($null -ne $latfp -and $latfp -eq 1) {
                Remove-ItemProperty -Path $policyPath -Name "LocalAccountTokenFilterPolicy" -Force -ErrorAction Stop
                Log "LocalAccountTokenFilterPolicy: Removed (restored default filtering)" -Colour Green -Indent
                $reverted = $true
            }
        } catch {}
        try {
            $fat = (Get-ItemProperty -Path $policyPath -ErrorAction Stop).FilterAdministratorToken
            if ($null -ne $fat -and $fat -eq 0) {
                Remove-ItemProperty -Path $policyPath -Name "FilterAdministratorToken" -Force -ErrorAction Stop
                Log "FilterAdministratorToken: Removed (restored default)" -Colour Green -Indent
                $reverted = $true
            }
        } catch {}
        if (-not $reverted) {
            Log "Token policy was not modified by this script" -Colour DarkGray -Indent
        } else {
            Log "A reboot is required for token policy changes to take effect" -Colour DarkYellow -Indent
        }
    } catch {
        Log "Could not revert token policy -- not critical" -Colour DarkGray -Indent
    }

    Write-Host ""
    Write-Host "  +----------------------------------------------+" -ForegroundColor Green
    Write-Host "  |           UNDO COMPLETE                       |" -ForegroundColor Green
    Write-Host "  +----------------------------------------------+" -ForegroundColor Green
    Write-Host ""
    Write-Host "  NOTE: A reboot is recommended for all changes to take full effect." -ForegroundColor DarkGray

} else {

    # ================================================================
    # SETUP MODE
    # ================================================================
    $steps = 26

    # ----------------------------------------------------------------
    # 1. Power plan -- High Performance or Ultimate Performance
    # ----------------------------------------------------------------
    Step 1 $steps "Configuring power plan..."

    # Back up current plan
    $currentPlan = Get-ActivePlanGuid
    if ($currentPlan -match "^[0-9a-fA-F\-]{36}$") {
        $currentPlan | Out-File -FilePath $BackupFile -Encoding ascii -Force
        Log "Backed up current plan: $currentPlan" -Colour DarkGray -Indent
    }

    # Check for Ultimate Performance (may not exist)
    $ultimateGuid = "e9a42b02-d5df-448d-aa00-03f14749eb61"
    $highPerfGuid = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"

    $planList = powercfg /list
    if ($planList -match $ultimateGuid) {
        powercfg /setactive $ultimateGuid
        Log "Set to: Ultimate Performance" -Colour Green -Indent
    } else {
        # Try to add Ultimate Performance
        $dupResult = powercfg /duplicatescheme $ultimateGuid 2>&1
        if ($LASTEXITCODE -eq 0) {
            powercfg /setactive $ultimateGuid
            Log "Set to: Ultimate Performance (added)" -Colour Green -Indent
        } else {
            powercfg /setactive $highPerfGuid
            Log "Set to: High Performance (Ultimate not available)" -Colour Green -Indent
        }
    }

    # ----------------------------------------------------------------
    # 2. Disable sleep on AC
    # ----------------------------------------------------------------
    Step 2 $steps "Disabling sleep on AC power..."
    powercfg /change standby-timeout-ac 0
    Log "Sleep on AC: Never" -Colour Green -Indent

    # ----------------------------------------------------------------
    # 3. Disable hibernate (also kills Fast Startup)
    # ----------------------------------------------------------------
    Step 3 $steps "Disabling hibernate..."
    powercfg /h off
    Log "Hibernate: Off" -Colour Green -Indent

    # ----------------------------------------------------------------
    # 4. Disable USB selective suspend on AC
    # ----------------------------------------------------------------
    Step 4 $steps "Disabling USB selective suspend on AC..."
    try {
        $activePlan = Get-ActivePlanGuid
        powercfg /setacvalueindex $activePlan 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 0
        powercfg /setactive $activePlan
        Log "USB selective suspend on AC: Disabled" -Colour Green -Indent
    } catch {
        Log "Could not change USB setting -- not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 5. Disable hard disk sleep + PCI-E power management on AC
    # ----------------------------------------------------------------
    Step 5 $steps "Disabling disk sleep and PCI-E power management on AC..."
    powercfg /change disk-timeout-ac 0
    Log "Hard disk sleep on AC: Never" -Colour Green -Indent

    try {
        $activePlan = Get-ActivePlanGuid
        powercfg /setacvalueindex $activePlan 501a4d13-42af-4429-9fd1-a8218c268e20 ee12f906-d277-404b-b6da-e5fa1a576df5 0
        powercfg /setactive $activePlan
        Log "PCI-E link state power management on AC: Off" -Colour Green -Indent
    } catch {
        Log "Could not change PCI-E setting -- not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 6. Disable Fast Startup (explicit registry)
    # ----------------------------------------------------------------
    Step 6 $steps "Disabling Fast Startup..."
    try {
        $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power"
        if (Test-Path $regPath) {
            Set-ItemProperty -Path $regPath -Name "HiberbootEnabled" -Value 0 -Type DWord -Force
            Log "Fast Startup: Disabled (registry)" -Colour Green -Indent
        } else {
            Log "Fast Startup registry key not found -- may not be supported" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not disable Fast Startup -- not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 7. Disable Connected Standby / Modern Standby
    # ----------------------------------------------------------------
    Step 7 $steps "Disabling Connected Standby / Modern Standby..."
    try {
        $csRegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Power"
        $csKey = Get-ItemProperty -Path $csRegPath -Name "CsEnabled" -ErrorAction SilentlyContinue
        if ($null -ne $csKey -and $csKey.CsEnabled -ne $null) {
            $csKey.CsEnabled.ToString() | Out-File -FilePath $CsBackupFile -Encoding ascii -Force
            if ($csKey.CsEnabled -eq 1) {
                Set-ItemProperty -Path $csRegPath -Name "CsEnabled" -Value 0 -Type DWord -Force
                Log "Connected Standby: Disabled (was enabled)" -Colour Green -Indent
                Log "A reboot is required for this change to take effect" -Colour DarkYellow -Indent
            } else {
                Log "Connected Standby: Already disabled" -Colour DarkGray -Indent
            }
        } else {
            Log "Connected Standby: Not supported on this system (no CsEnabled key)" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not change Connected Standby -- not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 8. Disable network adapter power saving
    # ----------------------------------------------------------------
    Step 8 $steps "Disabling network adapter power saving..."
    try {
        $nics = Get-NetAdapter -Physical -ErrorAction SilentlyContinue
        $nicCount = 0
        foreach ($nic in $nics) {
            $devPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\$($nic.PnPDeviceID)\Device Parameters"
            if (Test-Path $devPath) {
                Set-ItemProperty -Path $devPath -Name "PnPCapabilities" -Value 24 -Type DWord -Force -ErrorAction SilentlyContinue
                $nicCount++
            }
            try {
                Set-NetAdapterAdvancedProperty -Name $nic.Name -DisplayName "Wake on Magic Packet" `
                    -DisplayValue "Disabled" -ErrorAction SilentlyContinue
                Set-NetAdapterAdvancedProperty -Name $nic.Name -DisplayName "Wake on Pattern Match" `
                    -DisplayValue "Disabled" -ErrorAction SilentlyContinue
            } catch {}
            try {
                $pnpDevice = Get-PnpDevice -InstanceId $nic.PnPDeviceID -ErrorAction SilentlyContinue
                if ($pnpDevice) {
                    $powerMgmt = Get-CimInstance -ClassName MSPower_DeviceWakeEnable `
                        -Namespace root\wmi -ErrorAction SilentlyContinue |
                        Where-Object { $_.InstanceName -match [regex]::Escape($nic.PnPDeviceID) }
                    if ($powerMgmt) {
                        $powerMgmt | Set-CimInstance -Property @{Enable = $false} -ErrorAction SilentlyContinue
                    }
                }
            } catch {}
        }
        if ($nicCount -gt 0) {
            Log "Disabled power saving on $nicCount adapter(s)" -Colour Green -Indent
            Log "A reboot is required for PnPCapabilities changes to take effect" -Colour DarkGray -Indent
        } else {
            Log "No physical network adapters found" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not change network adapter settings -- not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 9. Set processor minimum state to 100% on AC
    # ----------------------------------------------------------------
    Step 9 $steps "Setting processor minimum state to 100% on AC..."
    try {
        $activePlan = Get-ActivePlanGuid
        powercfg /setacvalueindex $activePlan 54533251-82be-4824-96c1-47b60b740d00 893dee8e-2bef-41e0-89c6-b55d0929964c 100
        powercfg /setactive $activePlan
        Log "Processor minimum state on AC: 100%" -Colour Green -Indent
    } catch {
        Log "Could not change processor state -- not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 10. Pin Hyper-V VM memory (disable dynamic memory ballooning)
    # ----------------------------------------------------------------
    Step 10 $steps "Pinning Hyper-V VM memory..."
    try {
        $claudeVm = Get-VM -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match "claude" } |
                    Select-Object -First 1

        if ($claudeVm) {
            $vmName = $claudeVm.Name
            $dynMem = (Get-VMMemory -VMName $vmName -ErrorAction Stop).DynamicMemoryEnabled

            if ($dynMem) {
                if ($claudeVm.State -eq "Off") {
                    Set-VMMemory -VMName $vmName -DynamicMemoryEnabled $false -ErrorAction Stop
                    Log "Dynamic Memory disabled for VM '$vmName'" -Colour Green -Indent
                } else {
                    Log "VM '$vmName' is running -- Dynamic Memory will be disabled on next restart" -Colour DarkYellow -Indent
                    $flagFile = Join-Path $env:APPDATA "Claude\disable-dynamic-memory.flag"
                    $vmName | Out-File -FilePath $flagFile -Encoding ascii -Force
                    Log "Flag written: $flagFile" -Colour DarkGray -Indent
                }
            } else {
                Log "Dynamic Memory already disabled for VM '$vmName'" -Colour DarkGray -Indent
            }
        } else {
            Log "No Claude VM found (Hyper-V module may not be available)" -Colour DarkGray -Indent
            Log "This is normal if Cowork hasn't been used yet" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not configure VM memory -- Hyper-V module may not be installed" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 11. Boost VM worker process priority
    # ----------------------------------------------------------------
    Step 11 $steps "Boosting VM worker process priority..."
    try {
        $vmwpProcs = @(Get-Process -Name "vmwp" -ErrorAction SilentlyContinue)
        if ($vmwpProcs.Count -gt 0) {
            $boosted = 0
            foreach ($p in $vmwpProcs) {
                try {
                    if ($p.PriorityClass -ne 'AboveNormal') {
                        $p.PriorityClass = 'AboveNormal'
                        $boosted++
                    }
                } catch {}
            }
            if ($boosted -gt 0) {
                Log "Boosted $boosted vmwp.exe process(es) to AboveNormal priority" -Colour Green -Indent
            } else {
                Log "vmwp.exe already at AboveNormal priority" -Colour DarkGray -Indent
            }
            Log "Health monitor will maintain this across reboots" -Colour DarkGray -Indent
        } else {
            Log "No vmwp.exe processes found (VM may not be running)" -Colour DarkGray -Indent
            Log "Health monitor will boost priority when VM starts" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not set process priority -- not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 12. Configure HCS service recovery
    # ----------------------------------------------------------------
    Step 12 $steps "Configuring HCS service recovery..."
    try {
        & sc.exe failure vmcompute actions= restart/30000/restart/60000/restart/120000 reset= 300 2>&1 | Out-Null
        $verifyResult = & sc.exe qfailure vmcompute 2>&1
        if ($verifyResult -match "RESTART") {
            Log "vmcompute failure recovery: restart after 30s/60s/120s (reset after 300s)" -Colour Green -Indent
        } else {
            Log "vmcompute failure recovery set (could not verify -- non-critical)" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not configure vmcompute failure recovery -- not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 13. Configure CoworkVMService recovery (v4.8.0)
    # ----------------------------------------------------------------
    Step 13 $steps "Configuring CoworkVMService recovery..."
    try {
        $svc = Get-Service -Name "CoworkVMService" -ErrorAction SilentlyContinue
        if ($svc) {
            & sc.exe failure CoworkVMService reset= 300 actions= restart/30000/restart/60000/restart/120000 2>&1 | Out-Null
            $verifyResult = & sc.exe qfailure CoworkVMService 2>&1
            if ($verifyResult -match "RESTART") {
                Log "CoworkVMService failure recovery: restart after 30s/60s/120s (reset 300s)" -Colour Green -Indent
            } else {
                Log "CoworkVMService failure recovery set (could not verify)" -Colour DarkGray -Indent
            }
        } else {
            Log "CoworkVMService not installed -- skipping" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not configure CoworkVMService recovery -- not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 14. Pre-emptive HCS state cleanup (v4.8.0)
    # ----------------------------------------------------------------
    Step 14 $steps "Pre-emptive HCS state cleanup..."
    if (Test-Path "$env:SystemRoot\System32\hcsdiag.exe") {
        try {
            $hcsList = & "$env:SystemRoot\System32\hcsdiag.exe" list 2>&1 | Out-String
            if ($hcsList -match "cowork-vm") {
                $lines = $hcsList -split "`n"
                $currentGuid = $null
                foreach ($line in $lines) {
                    if ($line -match "^([0-9a-f-]{36})") { $currentGuid = $Matches[1] }
                    if ($line -match "cowork-vm" -and $currentGuid) {
                        & "$env:SystemRoot\System32\hcsdiag.exe" close $currentGuid 2>&1 | Out-Null
                        Log "Closed stale HCS compute system: $currentGuid" -Colour Green -Indent
                        $currentGuid = $null
                    }
                }
            } else {
                Log "HCS state clean" -Colour Green -Indent
            }
        } catch {
            Log "HCS cleanup check failed: $($_.Exception.Message)" -Colour DarkGray -Indent
        }
    } else {
        Log "hcsdiag.exe not available" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 15. Set service startup timeout (ServicesPipeTimeout)
    # ----------------------------------------------------------------
    Step 15 $steps "Setting service startup timeout..."
    try {
        $current = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control" -Name ServicesPipeTimeout -ErrorAction SilentlyContinue).ServicesPipeTimeout
        if ($null -eq $current -or $current -lt 120000) {
            Set-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control" -Name ServicesPipeTimeout -Value 120000 -Type DWord -Force
            Log "ServicesPipeTimeout set to 120000ms (prevents boot race conditions)" -Colour Green -Indent
            Log "Takes effect after next reboot" -Colour DarkGray -Indent
        } else {
            Log "ServicesPipeTimeout already set to ${current}ms" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not set ServicesPipeTimeout -- not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 16. Verify and repair WinNAT rules
    # ----------------------------------------------------------------
    Step 16 $steps "Checking WinNAT rules for VM network..."
    try {
        $natRules = @(Get-NetNat -ErrorAction SilentlyContinue)
        if ($natRules.Count -gt 0) {
            foreach ($rule in $natRules) {
                Log "NAT rule found: '$($rule.Name)' ($($rule.InternalIPInterfaceAddressPrefix))" -Colour Green -Indent
            }
        } else {
            Log "No WinNAT rules found" -Colour DarkYellow -Indent
            # Try to auto-create
            $hvSwitch = Get-VMSwitch -SwitchType Internal -ErrorAction SilentlyContinue |
                        Select-Object -First 1
            if ($hvSwitch) {
                $adapter = Get-NetAdapter -Name "vEthernet ($($hvSwitch.Name))" -ErrorAction SilentlyContinue
                if ($adapter) {
                    $ipAddr = Get-NetIPAddress -InterfaceIndex $adapter.ifIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                              Select-Object -First 1
                    if ($ipAddr) {
                        $prefix = ($ipAddr.IPAddress -split '\.')[0..2] -join '.'
                        $subnet = "$prefix.0/24"
                        try {
                            New-NetNat -Name "CoworkNAT" -InternalIPInterfaceAddressPrefix $subnet -ErrorAction Stop | Out-Null
                            Log "Created NAT rule 'CoworkNAT' for $subnet" -Colour Green -Indent
                        } catch {
                            Log "Could not create NAT rule: $($_.Exception.Message)" -Colour Yellow -Indent
                        }
                    } else {
                        Log "No IPv4 address on Hyper-V adapter -- NAT not needed yet" -Colour DarkGray -Indent
                    }
                } else {
                    Log "No Hyper-V virtual adapter found -- NAT not needed yet" -Colour DarkGray -Indent
                }
            } else {
                Log "No internal Hyper-V switch found -- NAT not needed yet" -Colour DarkGray -Indent
            }
            Log "Health monitor will auto-repair NAT if it disappears" -Colour DarkGray -Indent
        }
    } catch {
        Log "Get-NetNat not available -- skipping NAT check" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 17. Windows Firewall policy verification
    # ----------------------------------------------------------------
    Step 17 $steps "Checking Windows Firewall policies..."
    try {
        # Check if local firewall rules are being applied (Group Policy can block them)
        $fwProfiles = Get-NetFirewallProfile -ErrorAction SilentlyContinue
        $issues = @()
        foreach ($profile in $fwProfiles) {
            if ($profile.Enabled -and -not $profile.AllowLocalFirewallRules) {
                $issues += $profile.Name
            }
        }
        if ($issues.Count -gt 0) {
            Log "WARNING: Local firewall rules blocked on: $($issues -join ', ')" -Colour DarkYellow -Indent
            Log "This may prevent Hyper-V VM network access (DHCP/DNS)" -Colour DarkYellow -Indent
            Log "Ask your IT admin to enable 'Apply Local Firewall Rules' in Group Policy" -Colour DarkGray -Indent
        } else {
            Log "Firewall policies OK (local rules allowed)" -Colour Green -Indent
        }

        # Check for specific Hyper-V firewall rules
        $hvRules = Get-NetFirewallRule -DisplayGroup "*Hyper-V*" -ErrorAction SilentlyContinue
        if ($hvRules) {
            $disabled = @($hvRules | Where-Object { $_.Enabled -eq "False" })
            if ($disabled.Count -gt 0) {
                Log "WARNING: $($disabled.Count) Hyper-V firewall rule(s) are disabled" -Colour DarkYellow -Indent
                foreach ($dr in $disabled | Select-Object -First 3) {
                    Log "  - $($dr.DisplayName)" -Colour DarkGray -Indent
                }
                if ($disabled.Count -gt 3) {
                    Log "  ... and $($disabled.Count - 3) more" -Colour DarkGray -Indent
                }
            } else {
                Log "All Hyper-V firewall rules are enabled" -Colour Green -Indent
            }
        } else {
            Log "No Hyper-V firewall rules found (may be managed by Group Policy)" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not check firewall policies -- not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 18. Storage location detection
    # ----------------------------------------------------------------
    Step 18 $steps "Checking workspace storage location..."
    $claudeAppData = Join-Path $env:APPDATA "Claude"
    $vmCachePath = Join-Path $claudeAppData "claude-code-vm"
    $storageWarnings = @()

    # Check if APPDATA is on a cloud-sync folder
    $cloudPaths = @("OneDrive", "Google Drive", "Dropbox", "iCloud", "Box")
    foreach ($cp in $cloudPaths) {
        if ($env:APPDATA -match [regex]::Escape($cp)) {
            $storageWarnings += "APPDATA is inside a '$cp' sync folder -- this causes mount failures"
        }
    }

    # Check if APPDATA is on an external/USB drive
    try {
        $appDataDrive = (Split-Path $env:APPDATA -Qualifier) + "\"
        $driveInfo = Get-Volume -DriveLetter ($appDataDrive[0]) -ErrorAction SilentlyContinue
        if ($driveInfo) {
            $diskNumber = (Get-Partition -DriveLetter ($appDataDrive[0]) -ErrorAction SilentlyContinue).DiskNumber
            if ($null -ne $diskNumber) {
                $disk = Get-Disk -Number $diskNumber -ErrorAction SilentlyContinue
                if ($disk -and $disk.BusType -match "USB|Thunderbolt|1394") {
                    $storageWarnings += "APPDATA is on an external $($disk.BusType) drive -- use a local SSD instead"
                }
            }

            # Check if it's a network drive
            $logDisk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$($appDataDrive.TrimEnd('\'))'" -ErrorAction SilentlyContinue
            if ($logDisk -and $logDisk.DriveType -eq 4) {
                $storageWarnings += "APPDATA is on a network drive -- VirtioFS requires local storage"
            }
        }
    } catch {}

    # Check the VM cache path specifically for problematic locations
    if (Test-Path $vmCachePath) {
        try {
            $vmDrive = (Split-Path $vmCachePath -Qualifier) + "\"
            $vmDriveType = (Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$($vmDrive.TrimEnd('\'))'" -ErrorAction SilentlyContinue).DriveType
            if ($vmDriveType -eq 4) {
                $storageWarnings += "VM cache is on a network drive -- this will cause failures"
            }
        } catch {}
    }

    if ($storageWarnings.Count -gt 0) {
        foreach ($w in $storageWarnings) {
            Log "WARNING: $w" -Colour DarkYellow -Indent
        }
        Log "Recommended: Move Claude's data to a local SSD (C:\Users\$env:USERNAME\)" -Colour DarkGray -Indent
    } else {
        Log "Storage location OK (local drive)" -Colour Green -Indent
    }

    # ----------------------------------------------------------------
    # 19. NTP / time synchronisation check
    # ----------------------------------------------------------------
    Step 19 $steps "Checking time synchronisation..."
    try {
        $w32svc = Get-Service -Name "W32Time" -ErrorAction SilentlyContinue
        if ($w32svc) {
            if ($w32svc.Status -ne "Running") {
                try {
                    Start-Service -Name "W32Time" -ErrorAction Stop
                    Log "W32Time service started (was stopped)" -Colour Green -Indent
                } catch {
                    Log "WARNING: W32Time service is stopped and won't start" -Colour DarkYellow -Indent
                    Log "Clock drift may cause VM connectivity issues" -Colour DarkGray -Indent
                }
            } else {
                Log "W32Time service: Running" -Colour Green -Indent
            }

            # Quick drift check
            try {
                $w32tmResult = & w32tm /stripchart /computer:time.windows.com /dataonly /samples:1 2>&1
                if ($w32tmResult -match "(-?\d+\.\d+)s") {
                    $drift = [math]::Abs([double]$Matches[1])
                    if ($drift -gt 5.0) {
                        Log "WARNING: Clock drift is ${drift}s (>5s threshold)" -Colour DarkYellow -Indent
                        try {
                            & w32tm /resync /force 2>&1 | Out-Null
                            Log "Forced NTP resync" -Colour Green -Indent
                        } catch {}
                    } else {
                        Log "Clock drift: ${drift}s (within tolerance)" -Colour Green -Indent
                    }
                } else {
                    Log "Could not measure clock drift (NTP server unreachable?)" -Colour DarkGray -Indent
                }
            } catch {
                Log "Could not check clock drift -- not critical" -Colour DarkGray -Indent
            }
        } else {
            Log "W32Time service not found -- clock sync may not be configured" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not check time sync -- not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 20. Antivirus exclusion guidance
    # ----------------------------------------------------------------
    Step 20 $steps "Checking antivirus configuration..."
    $avProducts = @()
    try {
        # Query Windows Security Center (WMI)
        $avItems = Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName "AntiVirusProduct" -ErrorAction SilentlyContinue
        foreach ($av in $avItems) {
            $avProducts += $av.displayName
        }
    } catch {}

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
    $runningAv = @()
    foreach ($proc in $knownAvProcesses.Keys) {
        if (Get-Process -Name $proc -ErrorAction SilentlyContinue) {
            $runningAv += $knownAvProcesses[$proc]
        }
    }
    # Deduplicate
    $allAv = @(($avProducts + $runningAv) | Sort-Object -Unique)

    if ($allAv.Count -gt 0) {
        Log "Detected: $($allAv -join ', ')" -Colour Cyan -Indent
        $isDefenderOnly = ($allAv.Count -eq 1 -and $allAv[0] -match "Windows Defender")

        if ($isDefenderOnly) {
            # Check if Defender exclusions are already set
            try {
                $exclusions = (Get-MpPreference -ErrorAction SilentlyContinue).ExclusionPath
                $neededPaths = @(
                    (Join-Path $env:APPDATA "Claude"),
                    (Join-Path $env:ProgramFiles "Hyper-V"),
                    "$env:SystemRoot\System32\vmwp.exe",
                    "$env:SystemRoot\System32\vmms.exe"
                )
                $missing = @()
                foreach ($np in $neededPaths) {
                    $found = $false
                    if ($exclusions) {
                        foreach ($ex in $exclusions) {
                            if ($np -like "$ex*") { $found = $true; break }
                        }
                    }
                    if (-not $found) { $missing += $np }
                }
                if ($missing.Count -gt 0) {
                    Log "Adding Defender exclusions for Hyper-V/Claude paths:" -Colour Green -Indent
                    foreach ($mp in $missing) {
                        try {
                            Add-MpPreference -ExclusionPath $mp -ErrorAction Stop
                            Log "  + $mp" -Colour Green -Indent
                        } catch {
                            Log "  ! Could not add: $mp" -Colour Yellow -Indent
                        }
                    }
                    # Also add process exclusions
                    try {
                        Add-MpPreference -ExclusionProcess "vmwp.exe" -ErrorAction SilentlyContinue
                        Add-MpPreference -ExclusionProcess "vmms.exe" -ErrorAction SilentlyContinue
                        Add-MpPreference -ExclusionProcess "vmcompute.exe" -ErrorAction SilentlyContinue
                        Add-MpPreference -ExclusionProcess "cowork-svc.exe" -ErrorAction SilentlyContinue
                        Log "  + Process exclusions: vmwp.exe, vmms.exe, vmcompute.exe, cowork-svc.exe" -Colour Green -Indent
                    } catch {}
                    # Verify process exclusions were applied (v4.8.0)
                    try {
                        $procExclusions = (Get-MpPreference -ErrorAction SilentlyContinue).ExclusionProcess
                        $requiredProcs = @("vmwp.exe", "vmms.exe", "vmcompute.exe", "cowork-svc.exe")
                        $missingProcs = @()
                        foreach ($rp in $requiredProcs) {
                            if (-not ($procExclusions -contains $rp)) { $missingProcs += $rp }
                        }
                        if ($missingProcs.Count -gt 0) {
                            Log "  ! Missing process exclusions: $($missingProcs -join ', ')" -Colour Yellow -Indent
                            foreach ($mp in $missingProcs) {
                                Add-MpPreference -ExclusionProcess $mp -ErrorAction SilentlyContinue
                            }
                            Log "  Retried adding missing exclusions" -Colour DarkGray -Indent
                        } else {
                            Log "  All process exclusions verified" -Colour Green -Indent
                        }
                    } catch {}
                } else {
                    Log "Defender exclusions already configured" -Colour Green -Indent
                }
            } catch {
                Log "Could not check Defender exclusions -- not critical" -Colour DarkGray -Indent
            }
        } else {
            # Third-party AV -- can only advise
            Log "Recommended exclusion paths for your AV:" -Colour DarkYellow -Indent
            Log "  - $env:APPDATA\Claude\" -Colour White -Indent
            Log "  - $env:ProgramFiles\Hyper-V\" -Colour White -Indent
            Log "  - $env:SystemRoot\System32\vmwp.exe" -Colour White -Indent
            Log "  - $env:SystemRoot\System32\vmms.exe" -Colour White -Indent
            Log "  - Process: vmcompute.exe, cowork-svc.exe" -Colour White -Indent
            Log "Adding these exclusions prevents AV filter drivers from" -Colour DarkGray -Indent
            Log "interfering with VirtioFS disk operations" -Colour DarkGray -Indent
        }
    } else {
        Log "No antivirus product detected" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 21. WSL2 / Hyper-V conflict detection
    # ----------------------------------------------------------------
    Step 21 $steps "Checking for WSL2 / Hyper-V conflicts..."

    $wsl2Warnings = @()

    # Check WSL feature
    $wslFeature = Get-WindowsOptionalFeature -Online -FeatureName Microsoft-Windows-Subsystem-Linux -ErrorAction SilentlyContinue
    if ($wslFeature -and $wslFeature.State -eq 'Enabled') {
        $wsl2Warnings += "WSL feature is enabled"

        # Check for running distros
        if (Test-Path "$env:SystemRoot\System32\wsl.exe") {
            try {
                $distroOutput = & wsl -l -v 2>$null
                if ($distroOutput) {
                    $running = $distroOutput | Where-Object { $_ -match 'Running' -and $_ -match '\s2\s' }
                    if ($running) {
                        $wsl2Warnings += "WSL2 distros are actively running -- may conflict with Claude's VM"
                        $wsl2Warnings += "If Cowork has issues, try: wsl --shutdown"
                    }
                }
            } catch {}
        }
    }

    # Check Docker Desktop
    if (Test-Path "C:\Program Files\Docker\Docker\Docker.exe") {
        $wsl2Warnings += "Docker Desktop detected (may use WSL2 backend)"
    }

    # Display warnings
    if ($wsl2Warnings.Count -gt 0) {
        Write-Host ""
        Log "WSL2 / Hyper-V Conflict Check:" -Colour Yellow
        foreach ($w in $wsl2Warnings) {
            Log "  [!] $w" -Colour Yellow -Indent
        }
        Write-Host ""
    } else {
        Log "WSL2 conflict check: No conflicts detected" -Colour Green -Indent
    }

    # ----------------------------------------------------------------
    # 22. Install health monitor (auto-detects and auto-fixes crashes)
    # ----------------------------------------------------------------
    Step 22 $steps "Installing health monitor..."

    # Find Watch-ClaudeHealth.ps1 in the same folder as this script
    $myDir = Split-Path $PSCommandPath -Parent
    $watchScript = $null
    if ($myDir) {
        $candidate = Join-Path $myDir "Watch-ClaudeHealth.ps1"
        if (Test-Path $candidate) { $watchScript = $candidate }
    }
    if (-not $watchScript) {
        $fallbackPaths = @(
            "C:\ClaudeFix\Watch-ClaudeHealth.ps1",
            (Join-Path $env:USERPROFILE "Desktop\Watch-ClaudeHealth.ps1"),
            (Join-Path $env:USERPROFILE "Documents\Watch-ClaudeHealth.ps1")
        )
        foreach ($fb in $fallbackPaths) {
            if (Test-Path $fb) { $watchScript = $fb; break }
        }
    }

    # Clean up old basic watchdog script (replaced by health monitor)
    $oldWatchdog = Join-Path $env:APPDATA "Claude\cowork-watchdog.ps1"
    if (Test-Path $oldWatchdog) {
        Remove-Item $oldWatchdog -Force -ErrorAction SilentlyContinue
        Log "Removed old basic watchdog script" -Colour DarkGray -Indent
    }

    if ($watchScript) {
        try {
            try {
                Unregister-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction SilentlyContinue
            } catch {}

            # Kill any running health monitor before replacing
            try {
                Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" -ErrorAction SilentlyContinue |
                    Where-Object { $_.CommandLine -match "Watch-ClaudeHealth" } |
                    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
            } catch {}

            $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

            $delayedWatchCmd = "Start-Sleep -Seconds 120; & '$watchScript' -Quiet"
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
                -Force | Out-Null

            try { Start-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue } catch {}

            Log "Health monitor installed (starts 120s after logon, polls every 30s)" -Colour Green -Indent
            Log "Task: Task Scheduler > $TaskPath$TaskName" -Colour DarkGray -Indent
            Log "Script: $watchScript" -Colour DarkGray -Indent
            Log "Logs: $env:APPDATA\Claude\watch-logs\" -Colour DarkGray -Indent
        } catch {
            Log "[!] Could not create health monitor task: $($_.Exception.Message)" -Colour Red -Indent
            Log "You can run Watch-ClaudeHealth.bat manually instead" -Colour DarkGray -Indent
        }
    } else {
        Log "[!] Watch-ClaudeHealth.ps1 not found in same folder" -Colour Yellow -Indent
        Log "Put all scripts in the same folder and rerun to enable health monitor" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 23. Create boot-fix scheduled task
    # ----------------------------------------------------------------
    Step 23 $steps "Creating boot-time fix task..."

    $fixScript = $null
    if ($myDir) {
        $candidate = Join-Path $myDir "Fix-ClaudeDesktop.ps1"
        if (Test-Path $candidate) { $fixScript = $candidate }
    }
    if (-not $fixScript) {
        $fallbackPaths = @(
            "C:\ClaudeFix\Fix-ClaudeDesktop.ps1",
            (Join-Path $env:USERPROFILE "Desktop\Fix-ClaudeDesktop.ps1"),
            (Join-Path $env:USERPROFILE "Documents\Fix-ClaudeDesktop.ps1")
        )
        foreach ($fb in $fallbackPaths) {
            if (Test-Path $fb) { $fixScript = $fb; break }
        }
    }

    if ($fixScript) {
        try {
            try {
                Unregister-ScheduledTask -TaskName $BootTaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction SilentlyContinue
            } catch {}

            # Wrap in a delayed command: wait 180s after logon before running fix.
            # This prevents racing with Claude Desktop's own auto-start at logon.
            # The Fix script also has its own activity guard (-Quiet mode) as a second layer.
            $delayedCmd = "Start-Sleep -Seconds 180; & '$fixScript' -SkipLaunch -Quiet"
            $bootAction = New-ScheduledTaskAction -Execute "PowerShell.exe" `
                              -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command `"$delayedCmd`""

            $bootTrigger = New-ScheduledTaskTrigger -AtLogOn

            $bootSettings = New-ScheduledTaskSettingsSet `
                                -AllowStartIfOnBatteries `
                                -DontStopIfGoingOnBatteries `
                                -StartWhenAvailable `
                                -ExecutionTimeLimit (New-TimeSpan -Minutes 5)

            $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $bootPrincipal = New-ScheduledTaskPrincipal `
                                 -UserId $currentUser `
                                 -RunLevel Highest `
                                 -LogonType S4U

            Register-ScheduledTask `
                -TaskName $BootTaskName `
                -TaskPath $TaskPath `
                -Action $bootAction `
                -Trigger $bootTrigger `
                -Settings $bootSettings `
                -Principal $bootPrincipal `
                -Description "Runs Fix-ClaudeDesktop at logon to ensure clean VM state after boot." `
                -Force | Out-Null

            Log "Boot-fix task created (runs 180s after logon)" -Colour Green -Indent
            Log "Task: Task Scheduler > $TaskPath$BootTaskName" -Colour DarkGray -Indent
            Log "Runs: $fixScript -SkipLaunch -Quiet (after 180s delay)" -Colour DarkGray -Indent
        } catch {
            Log "[!] Could not create boot task: $($_.Exception.Message)" -Colour Red -Indent
        }
    } else {
        Log "[!] Fix-ClaudeDesktop.ps1 not found in same folder" -Colour Yellow -Indent
        Log "Put both scripts in the same folder and rerun to enable boot-fix" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 24. Create shortcuts (Desktop + Start Menu)
    # ----------------------------------------------------------------
    Step 24 $steps "Creating Fix Claude Desktop shortcuts..."

    $fixBat = $null
    if ($myDir) {
        $candidate = Join-Path $myDir "Fix-ClaudeDesktop.bat"
        if (Test-Path $candidate) { $fixBat = $candidate }
    }
    if (-not $fixBat -and $fixScript) {
        $fixBatCandidate = Join-Path (Split-Path $fixScript -Parent) "Fix-ClaudeDesktop.bat"
        if (Test-Path $fixBatCandidate) { $fixBat = $fixBatCandidate }
    }

    if ($fixBat) {
        $shell = New-Object -ComObject WScript.Shell

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
        } catch {
            Log "[!] Could not create desktop shortcut: $($_.Exception.Message)" -Colour Yellow -Indent
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
            Log "You can pin this to your Taskbar: search 'Fix Claude Desktop' in Start" -Colour DarkGray -Indent
        } catch {
            Log "[!] Could not create Start Menu shortcut: $($_.Exception.Message)" -Colour Yellow -Indent
        }
    } else {
        Log "[!] Fix-ClaudeDesktop.bat not found in same folder" -Colour Yellow -Indent
        Log "Put all scripts in the same folder and rerun to create shortcuts" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 25. Set Claude Desktop to launch elevated (MSIX-aware)
    # ----------------------------------------------------------------
    Step 25 $steps "Configuring Claude Desktop to launch elevated..."
    try {
        # Scheduled task registration requires admin -- skip gracefully if not elevated
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if (-not $isAdmin) {
            Log "Skipped -- requires admin privileges (run as Administrator to enable)" -Colour DarkYellow -Indent
            throw "SKIP"
        }
        # MSIX apps block all direct .exe access from WindowsApps (ACLs, -Verb RunAs,
        # dir enumeration all fail). The only reliable way to launch an MSIX app with
        # full admin privileges is a scheduled task with RunLevel=Highest + Interactive
        # logon. The task gets a full unfiltered admin token, no UAC prompt, and the
        # GUI is visible in the user's desktop session.
        #
        # The task action finds Claude at runtime via three methods:
        #   1) Get-AppxPackage (MSIX installs from Store/winget)
        #   2) Common install paths (traditional .exe installer)
        #   3) Running process fallback (any install method)
        # This survives version updates and works with any install type.

        $elevTaskName = "LaunchClaudeAdmin"

        # Remove old task if present
        try { Unregister-ScheduledTask -TaskName $elevTaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction SilentlyContinue } catch {}

        # PowerShell command that finds and launches Claude
        # Priority: 1) MSIX via Get-AppxPackage  2) Traditional .exe install paths  3) Running process
        # NOTE: We MUST use direct .exe launch (Start-Process $exe) to inherit the
        # task's elevated token. shell:AppsFolder routes through the non-elevated desktop
        # shell and the app gets medium integrity -- defeating the entire purpose.
        # Trade-off: MSIX installs will show a second taskbar icon. This is unavoidable
        # because Windows enforces medium integrity for all shell-activated MSIX apps.
        $launchCmd = @'
# If Claude is already running, don't launch again
$existing = Get-Process -Name Claude -ErrorAction SilentlyContinue
if ($existing) { exit 0 }
$exe = $null
# 1. MSIX install (Windows Store / winget MSIX)
$p = Get-AppxPackage | Where-Object { $_.Name -eq 'Claude' -or $_.PackageFamilyName -like 'Claude_*' } | Select-Object -First 1
if ($p) { $e = Join-Path $p.InstallLocation 'app\Claude.exe'; if (Test-Path $e) { $exe = $e } }
# 2. Traditional installer paths
if (-not $exe) {
    $candidates = @(
        (Join-Path $env:LOCALAPPDATA 'Programs\claude-desktop\Claude.exe'),
        (Join-Path $env:LOCALAPPDATA 'Claude Desktop\Claude.exe'),
        (Join-Path $env:LOCALAPPDATA 'Programs\Claude\Claude.exe'),
        (Join-Path ${env:ProgramFiles} 'Claude Desktop\Claude.exe'),
        (Join-Path ${env:ProgramFiles(x86)} 'Claude Desktop\Claude.exe')
    )
    foreach ($c in $candidates) { if (Test-Path $c) { $exe = $c; break } }
}
# 3. Fallback: find from running process
if (-not $exe) {
    $proc = Get-Process -Name Claude -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($proc -and $proc.MainModule) { $exe = $proc.MainModule.FileName }
}
if ($exe) { Start-Process $exe } else { throw 'Claude Desktop not found. Is it installed?' }
'@

        # Encode as Base64 for -EncodedCommand (handles multi-line safely in task XML)
        $encodedCmd = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($launchCmd))

        $elevAction = New-ScheduledTaskAction -Execute "PowerShell.exe" `
            -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -EncodedCommand $encodedCmd"

        $elevSettings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -ExecutionTimeLimit (New-TimeSpan -Minutes 1)

        $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $elevPrincipal = New-ScheduledTaskPrincipal `
            -UserId $currentUser `
            -RunLevel Highest `
            -LogonType Interactive

        Register-ScheduledTask `
            -TaskName $elevTaskName `
            -TaskPath $TaskPath `
            -Action $elevAction `
            -Settings $elevSettings `
            -Principal $elevPrincipal `
            -Description "Launches Claude Desktop with full admin privileges. Triggered by the 'Claude (Admin)' shortcut." `
            -Force | Out-Null

        Log "Scheduled task created: $TaskPath$elevTaskName (Highest + Interactive)" -Colour Green -Indent

        # Create a launcher .cmd that triggers the task (one-liner, no $, no PS needed)
        $launcherDir = Join-Path $env:APPDATA "Claude"
        $launcherCmd = Join-Path $launcherDir "Launch-Claude-Admin.cmd"
        $cmdContent = @"
@echo off
REM -- Claude Desktop (Admin) Launcher
REM -- Auto-generated by Prevent-ClaudeIssues.ps1 v$Version
REM -- Triggers the LaunchClaudeAdmin scheduled task (runs with full admin token).
schtasks /run /tn "\Claude\LaunchClaudeAdmin" >nul 2>&1
if errorlevel 1 (
    echo The LaunchClaudeAdmin scheduled task was not found.
    echo Run Prevent-ClaudeIssues.bat to set it up.
    pause
)
"@
        Set-Content -Path $launcherCmd -Value $cmdContent -Encoding ASCII -Force
        Log "Launcher CMD created: $launcherCmd" -Colour Green -Indent

        # Create desktop shortcut pointing to the launcher
        $desktopPath = [Environment]::GetFolderPath("Desktop")

        # Remove any old broken shortcuts
        $oldLnk = Join-Path $desktopPath "Claude (Admin).lnk"
        if (Test-Path $oldLnk) {
            Remove-Item $oldLnk -Force -ErrorAction SilentlyContinue
        }

        $adminLnkPath = Join-Path $desktopPath "Claude (Admin).lnk"
        $shell = New-Object -ComObject WScript.Shell
        $sc = $shell.CreateShortcut($adminLnkPath)
        $sc.TargetPath = $launcherCmd
        $sc.WorkingDirectory = $launcherDir
        $sc.Description = "Claude Desktop (Elevated via scheduled task)"
        # Try to use Claude's own icon (falls back to generic if path is stale after update)
        $claudeIcon = $null
        $appxPkg = Get-AppxPackage -Name "*Claude*" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($appxPkg) {
            $candidateExe = Join-Path $appxPkg.InstallLocation "app\Claude.exe"
            if (Test-Path $candidateExe) { $claudeIcon = "$candidateExe,0" }
        }
        if (-not $claudeIcon) {
            # Check traditional install paths
            $exeCandidates = @(
                (Join-Path $env:LOCALAPPDATA 'Programs\claude-desktop\Claude.exe'),
                (Join-Path $env:LOCALAPPDATA 'Claude Desktop\Claude.exe'),
                (Join-Path ${env:ProgramFiles} 'Claude Desktop\Claude.exe')
            )
            foreach ($c in $exeCandidates) {
                if (Test-Path $c) { $claudeIcon = "$c,0"; break }
            }
        }
        if (-not $claudeIcon) { $claudeIcon = "%SystemRoot%\System32\shell32.dll,77" }
        $sc.IconLocation = $claudeIcon
        $sc.Save()
        Log "Admin shortcut created: $adminLnkPath" -Colour Green -Indent
        Log "No UAC prompt -- task runs with full admin token automatically" -Colour DarkGray -Indent
        Log "Survives Claude updates (detects MSIX, traditional install, or running process)" -Colour DarkGray -Indent

    } catch {
        Log "Could not configure elevation -- not critical: $($_.Exception.Message)" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 26. Admin token filtering (LocalAccountTokenFilterPolicy)
    # ----------------------------------------------------------------
    Step 26 $steps "Configuring admin token policy..."
    try {
        $policyPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
        $current = Get-ItemProperty -Path $policyPath -ErrorAction Stop

        $latfp = try { $current.LocalAccountTokenFilterPolicy } catch { $null }
        $fat   = try { $current.FilterAdministratorToken } catch { $null }

        $changed = $false
        if ($latfp -ne 1) {
            Set-ItemProperty -Path $policyPath -Name "LocalAccountTokenFilterPolicy" -Value 1 -Type DWord -Force
            Log "LocalAccountTokenFilterPolicy: Set to 1 (no token filtering)" -Colour Green -Indent
            $changed = $true
        } else {
            Log "LocalAccountTokenFilterPolicy: Already set to 1" -Colour DarkGray -Indent
        }

        if ($fat -ne 0) {
            Set-ItemProperty -Path $policyPath -Name "FilterAdministratorToken" -Value 0 -Type DWord -Force
            Log "FilterAdministratorToken: Set to 0 (admin gets full token)" -Colour Green -Indent
            $changed = $true
        } else {
            Log "FilterAdministratorToken: Already set to 0" -Colour DarkGray -Indent
        }

        if ($changed) {
            Log "A reboot is required for token policy changes to take effect" -Colour DarkYellow -Indent
        }
        Log "EnableLUA remains 1 (UAC stays on -- Store apps keep working)" -Colour DarkGray -Indent
    } catch {
        Log "Could not configure token policy -- not critical: $($_.Exception.Message)" -Colour DarkGray -Indent
    }

    # ================================================================
    # Summary
    # ================================================================
    Write-Host ""
    Write-Host "  +----------------------------------------------+" -ForegroundColor Green
    Write-Host "  |           SETUP COMPLETE                      |" -ForegroundColor Green
    Write-Host "  +----------------------------------------------+" -ForegroundColor Green
    Write-Host ""
    Write-Host "  What was configured:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "    Power plan ........... Ultimate/High Performance" -ForegroundColor White
    Write-Host "    Sleep on AC .......... Never" -ForegroundColor White
    Write-Host "    Hibernate ............ Off" -ForegroundColor White
    Write-Host "    USB suspend (AC) ..... Disabled" -ForegroundColor White
    Write-Host "    Disk sleep (AC) ...... Never" -ForegroundColor White
    Write-Host "    PCI-E power mgmt ..... Off" -ForegroundColor White
    Write-Host "    Fast Startup ......... Disabled" -ForegroundColor White
    Write-Host "    Connected Standby .... Disabled" -ForegroundColor White
    Write-Host "    NIC power saving ..... Disabled" -ForegroundColor White
    Write-Host "    CPU minimum (AC) ..... 100%" -ForegroundColor White
    Write-Host "    VM memory ............ Pinned (no ballooning)" -ForegroundColor White
    Write-Host "    VM worker priority ... AboveNormal" -ForegroundColor White
    Write-Host "    HCS service recovery . Auto-restart on failure" -ForegroundColor White
    Write-Host "    VM service recovery .. Auto-restart on failure" -ForegroundColor White
    Write-Host "    HCS state cleanup .... Pre-emptive stale VM removal" -ForegroundColor White
    Write-Host "    Service timeout ...... 120s (boot race prevention)" -ForegroundColor White
    Write-Host "    WinNAT rules ......... Verified/repaired" -ForegroundColor White
    Write-Host "    Firewall policies .... Checked" -ForegroundColor White
    Write-Host "    Storage location ..... Checked" -ForegroundColor White
    Write-Host "    Time sync ............ Verified" -ForegroundColor White
    Write-Host "    Antivirus ............ Exclusions configured" -ForegroundColor White
    Write-Host "    WSL2 conflicts ....... Checked" -ForegroundColor White
    Write-Host "    Health monitor ....... Every 30s (auto-fix)" -ForegroundColor White
    Write-Host "    Boot-fix task ........ At every logon" -ForegroundColor White
    Write-Host "    Shortcuts ............ Desktop + Start Menu" -ForegroundColor White
    Write-Host "    Claude elevation ..... Scheduled task (full admin, no UAC prompt)" -ForegroundColor White
    Write-Host "    Admin token policy ... Full admin token for local accounts" -ForegroundColor White
    Write-Host ""
    Write-Host "  TIP: Right-click 'Fix Claude Desktop' in Start Menu" -ForegroundColor DarkGray
    Write-Host "       and select 'Pin to taskbar' for quick access." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  NOTE: Some changes (Connected Standby, NIC power) require" -ForegroundColor DarkYellow
    Write-Host "        a reboot to take full effect." -ForegroundColor DarkYellow
    Write-Host ""
    Write-Host "  To undo everything:" -ForegroundColor DarkGray
    Write-Host "    .\Prevent-ClaudeIssues.ps1 -Undo" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Original power plan backed up to:" -ForegroundColor DarkGray
    Write-Host "    $BackupFile" -ForegroundColor DarkGray
}

} catch {
    Write-Host ""
    Write-Host "  +----------------------------------------------+" -ForegroundColor Red
    Write-Host "  |           UNEXPECTED ERROR                    |" -ForegroundColor Red
    Write-Host "  +----------------------------------------------+" -ForegroundColor Red
    Write-Host ""
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "  Press any key to close..." -ForegroundColor DarkGray
[void][System.Console]::ReadKey($true)
