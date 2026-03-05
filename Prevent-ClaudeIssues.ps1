<#
.SYNOPSIS
    Claude Desktop / Cowork -- Preventive Configuration

.DESCRIPTION
    One-shot script that configures Windows to minimise VirtioFS/Plan9
    mount failures in Claude Desktop's Cowork VM.

    Run once. Changes persist across reboots.

    What it does:
    - Sets power plan to High Performance (or Ultimate if available)
    - Disables sleep on AC power
    - Disables hibernate
    - Disables USB selective suspend
    - Disables hard disk spin-down on AC
    - Disables PCI Express link state power management on AC
    - Registers a scheduled task that monitors CoworkVMService health
      and auto-restarts it if it enters a failed state

    What it does NOT do:
    - Touch your Claude config or conversations
    - Disable sleep on battery (laptop users keep battery sleep)
    - Change your screen timeout

.PARAMETER Undo
    Reverts all changes made by this script.

.NOTES
    Version : 1.3.0
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
$Version          = "1.3.0"
$TaskName         = "ClaudeCoworkWatchdog"
$BootTaskName     = "ClaudeCoworkBootFix"
$TaskPath         = "\Claude\"
$ServiceName      = "CoworkVMService"
$BackupFile       = Join-Path $env:APPDATA "Claude\power-plan-backup.txt"

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
    $steps = 5

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

    Step 2 $steps "Re-enabling hibernate..."
    powercfg /h on
    Log "Hibernate enabled" -Colour Green -Indent

    Step 3 $steps "Resetting sleep timeout to 30 minutes (AC)..."
    powercfg /change standby-timeout-ac 30
    Log "Sleep timeout set to 30 min on AC" -Colour Green -Indent

    Step 4 $steps "Removing scheduled tasks..."
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

    Step 5 $steps "Removing shortcuts..."
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

    Write-Host ""
    Write-Host "  +----------------------------------------------+" -ForegroundColor Green
    Write-Host "  |           UNDO COMPLETE                       |" -ForegroundColor Green
    Write-Host "  +----------------------------------------------+" -ForegroundColor Green

} else {

    # ================================================================
    # SETUP MODE
    # ================================================================
    $steps = 8

    # ----------------------------------------------------------------
    # 1. Power plan -- High Performance or Ultimate Performance
    # ----------------------------------------------------------------
    Step 1 $steps "Configuring power plan..."

    # Back up current plan
    $currentPlan = (powercfg /getactivescheme) -replace '.*:\s*', '' -replace '\s*\(.*', ''
    $currentPlan = $currentPlan.Trim()
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
    # 3. Disable hibernate
    # ----------------------------------------------------------------
    Step 3 $steps "Disabling hibernate..."
    powercfg /h off
    Log "Hibernate: Off" -Colour Green -Indent

    # ----------------------------------------------------------------
    # 4. Disable USB selective suspend on AC
    # ----------------------------------------------------------------
    Step 4 $steps "Disabling USB selective suspend on AC..."
    # USB selective suspend: GUID 2a737441-1930-4402-8d77-b2bebba308a3
    # Setting GUID:          48e6b7a6-50f5-4782-a5d4-53bb8f07e226
    # 0 = Disabled, 1 = Enabled
    try {
        $activePlan = (powercfg /getactivescheme) -replace '.*:\s*', '' -replace '\s*\(.*', ''
        $activePlan = $activePlan.Trim()
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
    # Hard disk timeout: 0 = never
    powercfg /change disk-timeout-ac 0
    Log "Hard disk sleep on AC: Never" -Colour Green -Indent

    # PCI Express Link State Power Management
    # Sub-group: 501a4d13-42af-4429-9fd1-a8218c268e20
    # Setting:   ee12f906-d277-404b-b6da-e5fa1a576df5
    # 0 = Off
    try {
        $activePlan = (powercfg /getactivescheme) -replace '.*:\s*', '' -replace '\s*\(.*', ''
        $activePlan = $activePlan.Trim()
        powercfg /setacvalueindex $activePlan 501a4d13-42af-4429-9fd1-a8218c268e20 ee12f906-d277-404b-b6da-e5fa1a576df5 0
        powercfg /setactive $activePlan
        Log "PCI-E link state power management on AC: Off" -Colour Green -Indent
    } catch {
        Log "Could not change PCI-E setting -- not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 6. Create watchdog scheduled task
    # ----------------------------------------------------------------
    Step 6 $steps "Creating CoworkVMService watchdog task..."

    $watchdogScript = @'
$svc = Get-Service -Name "CoworkVMService" -ErrorAction SilentlyContinue
if (-not $svc) { exit }
if ($svc.Status -eq "Running") { exit }

# Service is not running -- check if Claude Desktop is
$claude = Get-Process -Name "claude" -ErrorAction SilentlyContinue
if (-not $claude) { exit }

# Claude is running but service is not -- restart it
try {
    Start-Service -Name "CoworkVMService" -ErrorAction Stop
} catch {
    try {
        Restart-Service -Name "CoworkVMService" -Force -ErrorAction Stop
    } catch {}
}
'@

    $watchdogPath = Join-Path $env:APPDATA "Claude\cowork-watchdog.ps1"
    $watchdogScript | Out-File -FilePath $watchdogPath -Encoding ascii -Force

    try {
        # Remove existing task if present
        try {
            Unregister-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction SilentlyContinue
        } catch {}

        $action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
                      -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$watchdogPath`""

        # Create a Once trigger, then patch in repetition via COM
        $trigger = New-ScheduledTaskTrigger -Once -At "00:00" `
                       -RepetitionInterval (New-TimeSpan -Minutes 5)

        # Set repetition duration to indefinite (PS 5.1 workaround)
        $trigger.Repetition.StopAtDurationEnd = $false
        $trigger.Repetition.Duration = "P9999D"

        $settings = New-ScheduledTaskSettingsSet `
                        -AllowStartIfOnBatteries `
                        -DontStopIfGoingOnBatteries `
                        -StartWhenAvailable `
                        -ExecutionTimeLimit (New-TimeSpan -Minutes 1)

        $principal = New-ScheduledTaskPrincipal `
                         -UserId "SYSTEM" `
                         -RunLevel Highest `
                         -LogonType ServiceAccount

        Register-ScheduledTask `
            -TaskName $TaskName `
            -TaskPath $TaskPath `
            -Action $action `
            -Trigger $trigger `
            -Settings $settings `
            -Principal $principal `
            -Description "Monitors CoworkVMService and restarts it if Claude Desktop is running but the service has stopped." `
            -Force | Out-Null

        Log "Watchdog task created (runs every 5 minutes)" -Colour Green -Indent
        Log "Task: Task Scheduler > $TaskPath$TaskName" -Colour DarkGray -Indent
        Log "Script: $watchdogPath" -Colour DarkGray -Indent
    } catch {
        Log "[!] Could not create scheduled task: $($_.Exception.Message)" -Colour Red -Indent
        Log "Watchdog script saved to: $watchdogPath" -Colour DarkGray -Indent
        Log "You can manually create a task to run it every 5 minutes" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 7. Create boot-fix scheduled task
    # ----------------------------------------------------------------
    Step 7 $steps "Creating boot-time fix task..."

    # Find Fix-ClaudeDesktop.ps1 -- look in the same folder as this script
    $fixScript = $null
    $myDir = Split-Path $PSCommandPath -Parent
    if ($myDir) {
        $candidate = Join-Path $myDir "Fix-ClaudeDesktop.ps1"
        if (Test-Path $candidate) { $fixScript = $candidate }
    }
    # Fallback: check common locations
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
            # Remove existing task if present
            try {
                Unregister-ScheduledTask -TaskName $BootTaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction SilentlyContinue
            } catch {}

            $bootAction = New-ScheduledTaskAction -Execute "PowerShell.exe" `
                              -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$fixScript`" -SkipLaunch -Quiet"

            $bootTrigger = New-ScheduledTaskTrigger -AtLogOn

            $bootSettings = New-ScheduledTaskSettingsSet `
                                -AllowStartIfOnBatteries `
                                -DontStopIfGoingOnBatteries `
                                -StartWhenAvailable `
                                -ExecutionTimeLimit (New-TimeSpan -Minutes 5)

            $bootPrincipal = New-ScheduledTaskPrincipal `
                                 -UserId "SYSTEM" `
                                 -RunLevel Highest `
                                 -LogonType ServiceAccount

            Register-ScheduledTask `
                -TaskName $BootTaskName `
                -TaskPath $TaskPath `
                -Action $bootAction `
                -Trigger $bootTrigger `
                -Settings $bootSettings `
                -Principal $bootPrincipal `
                -Description "Runs Fix-ClaudeDesktop at logon to ensure clean VM state after boot." `
                -Force | Out-Null

            Log "Boot-fix task created (runs at every logon)" -Colour Green -Indent
            Log "Task: Task Scheduler > $TaskPath$BootTaskName" -Colour DarkGray -Indent
            Log "Runs: $fixScript -Quiet" -Colour DarkGray -Indent
        } catch {
            Log "[!] Could not create boot task: $($_.Exception.Message)" -Colour Red -Indent
        }
    } else {
        Log "[!] Fix-ClaudeDesktop.ps1 not found in same folder" -Colour Yellow -Indent
        Log "Put both scripts in the same folder and rerun to enable boot-fix" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 8. Create shortcuts (Desktop + Start Menu) for Fix-ClaudeDesktop
    # ----------------------------------------------------------------
    Step 8 $steps "Creating Fix Claude Desktop shortcuts..."

    $fixBat = $null
    if ($myDir) {
        $candidate = Join-Path $myDir "Fix-ClaudeDesktop.bat"
        if (Test-Path $candidate) { $fixBat = $candidate }
    }
    if (-not $fixBat -and $fixScript) {
        # If we found the .ps1 earlier, check for .bat next to it
        $fixBatCandidate = Join-Path (Split-Path $fixScript -Parent) "Fix-ClaudeDesktop.bat"
        if (Test-Path $fixBatCandidate) { $fixBat = $fixBatCandidate }
    }

    if ($fixBat) {
        $shell = New-Object -ComObject WScript.Shell

        # Desktop shortcut
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $desktopLnk = Join-Path $desktopPath "Fix Claude Desktop.lnk"
        try {
            $sc = $shell.CreateShortcut($desktopLnk)
            $sc.TargetPath = $fixBat
            $sc.WorkingDirectory = Split-Path $fixBat -Parent
            $sc.Description = "Reset and fix Claude Desktop / Cowork VM"
            $sc.WindowStyle = 1  # Normal window
            # Use shield icon from shell32.dll (index 77)
            $sc.IconLocation = "%SystemRoot%\System32\shell32.dll,77"
            $sc.Save()
            Log "Desktop shortcut created" -Colour Green -Indent
        } catch {
            Log "[!] Could not create desktop shortcut: $($_.Exception.Message)" -Colour Yellow -Indent
        }

        # Start Menu shortcut
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
    Write-Host "    Watchdog task ........ Every 5 min" -ForegroundColor White
    Write-Host "    Boot-fix task ........ At every logon" -ForegroundColor White
    Write-Host "    Shortcuts ............ Desktop + Start Menu" -ForegroundColor White
    Write-Host ""
    Write-Host "  TIP: Right-click 'Fix Claude Desktop' in Start Menu" -ForegroundColor DarkGray
    Write-Host "       and select 'Pin to taskbar' for quick access." -ForegroundColor DarkGray
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
