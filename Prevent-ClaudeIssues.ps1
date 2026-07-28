<#
.SYNOPSIS
    Claude Desktop / Cowork, Preventive Configuration

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
    Version : 6.0.1
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

    # -WhatIf has to survive the elevation boundary. Without this line, running
    # the script unelevated with -WhatIf relaunched it elevated WITHOUT -WhatIf
    # and applied every change for real, which is the exact opposite of what was
    # asked for. Fix-ClaudeDesktop already forwarded it; this one did not.
    if ($WhatIfPreference) { $elevateArgs += " -WhatIf" }

    try {
        # -Wait -PassThru so the child's exit code reaches the launcher. Without
        # it the outer shell returned 0 the moment UAC was accepted, so a failed
        # elevated run looked like success to Prevent-ClaudeIssues.bat.
        $elevated = Start-Process PowerShell -ArgumentList $elevateArgs -Verb RunAs -Wait -PassThru
        if ($elevated) { exit $elevated.ExitCode }
        exit 0
    } catch {
        Write-Host "  [!] UAC elevation was declined or failed." -ForegroundColor Red
        Write-Host "      This script requires Administrator privileges." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Press any key to close..." -ForegroundColor DarkGray
        [void][System.Console]::ReadKey($true)
    }
    # Explicit 1, not a bare exit. Declining UAC means nothing was configured,
    # and the launcher treats 1 as "the script explained itself and paused",
    # which is exactly what just happened above.
    exit 1
}

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
# -SkipCacheInventory: nothing here needs the VHDX inventory, and the walk is
# the expensive half of discovery.
$script:ClaudeEnv = Get-ClaudeEnvironment -SkipCacheInventory

# -- Constants -------------------------------------------------------
$ToolkitVersion   = "6.0.1"
$TaskName         = "ClaudeCoworkWatchdog"
$BootTaskName     = "ClaudeCoworkBootFix"
$TaskPath         = "\Claude\"
$ClaudeAppData    = $script:ClaudeEnv.AppDataDir

# Both discovered. $ServiceName was removed in the first pass as an unused
# variable, which it was: everything referred to the literal "CoworkVMService"
# instead. It is back because those references now use it.
$ServiceName      = $script:ClaudeEnv.ServiceName
$ServiceExe       = $script:ClaudeEnv.ServiceProcessName

if (-not $ClaudeAppData) {
    Write-Host "  [!] Could not locate the per-user Claude folder under %APPDATA%." -ForegroundColor Red
    exit 1
}
$BackupFile       = Join-Path $ClaudeAppData "power-plan-backup.txt"
$CsBackupFile     = Join-Path $ClaudeAppData "connected-standby-backup.txt"

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

function Initialize-ClaudeAppData {
    # %APPDATA%\Claude is created by Claude Desktop on first launch. This script
    # is meant to be runnable BEFORE Claude Desktop is installed, and it also
    # self-elevates -- which can land it in a different account's profile that
    # has never launched Claude. Either way the folder may not exist yet, and
    # every state file this script writes lives inside it.
    if (Test-Path -LiteralPath $script:ClaudeAppData) { return $true }
    try {
        New-Item -Path $script:ClaudeAppData -ItemType Directory -Force -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Save-StateFile {
    param([string]$Path, [string]$Value, [string]$Label)
    if (-not (Initialize-ClaudeAppData)) {
        Log "Could not create $script:ClaudeAppData, $Label not saved" -Colour DarkYellow -Indent
        return $false
    }
    try {
        $Value | Out-File -FilePath $Path -Encoding ascii -Force -ErrorAction Stop
        return $true
    } catch {
        Log "Could not write $Label, continuing" -Colour DarkYellow -Indent
        return $false
    }
}

function Get-ActivePlanGuid {
    $raw = (powercfg /getactivescheme) -replace '.*:\s*', '' -replace '\s*\(.*', ''
    return $raw.Trim()
}

# -- Header ----------------------------------------------------------
Write-Host ""
Write-Host "  +----------------------------------------------+" -ForegroundColor Cyan
Write-Host "  |  CLAUDE DESKTOP / COWORK, PREVENTION SETUP   |" -ForegroundColor Cyan
Write-Host "  |  v$ToolkitVersion                                      |" -ForegroundColor DarkGray
Write-Host "  +----------------------------------------------+" -ForegroundColor Cyan
Write-Host ""

if ($Undo) {
    Write-Host "  MODE: UNDO, reverting all changes" -ForegroundColor Yellow
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
            Log "Backup file corrupt, please set your power plan manually" -Colour Yellow -Indent
        }
        # Clean up any ClaudeFix-created Ultimate Performance plans
        $planLines = powercfg /list | Where-Object { $_ -match 'Ultimate Performance' }
        $deleteCount = 0
        foreach ($line in $planLines) {
            if ($line -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})') {
                powercfg /delete $Matches[1] 2>$null
                $deleteCount++
            }
        }
        if ($deleteCount -gt 0) {
            Log "Deleted $deleteCount Ultimate Performance plan(s)" -Colour DarkGray -Indent
        }
    } else {
        Log "No backup found, power plan was not changed by this script" -Colour DarkGray -Indent
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
        Log "Could not re-enable Fast Startup, not critical" -Colour DarkGray -Indent
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
            Log "Could not restore Connected Standby, check manually" -Colour Yellow -Indent
        }
    } else {
        Log "No Connected Standby backup found, was not changed" -Colour DarkGray -Indent
    }

    Step 5 $steps "Re-enabling network adapter power management..."
    try {
        $nics = Get-NetAdapter -Physical -ErrorAction SilentlyContinue
        foreach ($nic in $nics) {
            # The $pnp query and the $regBase path that used to be here are
            # gone: both were assigned and never read. Removing the PnP
            # capabilities value below is what actually restores the default,
            # and it needs neither of them.
            $devPath = "HKLM:\SYSTEM\CurrentControlSet\Enum\$($nic.PnPDeviceID)\Device Parameters"
            if (Test-Path $devPath) {
                Remove-ItemProperty -Path $devPath -Name "PnPCapabilities" -ErrorAction SilentlyContinue
            }
        }
        Log "Network adapter power management: Restored to defaults" -Colour Green -Indent
        Log "A reboot is required for this change to take effect" -Colour DarkGray -Indent
    } catch {
        Log "Could not restore network adapter settings, not critical" -Colour DarkGray -Indent
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
                Log "VM '$vmName' is running, restart it to re-enable Dynamic Memory" -Colour DarkYellow -Indent
            }
        } else {
            Log "No Claude VM found, nothing to restore" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not restore Dynamic Memory, Hyper-V module may not be available" -Colour DarkGray -Indent
    }
    # Clean up flag file
    $flagFile = Join-Path $ClaudeAppData "disable-dynamic-memory.flag"
    if (Test-Path $flagFile) {
        Remove-Item $flagFile -Force -ErrorAction SilentlyContinue
    }

    Step 7 $steps "Reverting HCS service configuration..."
    try {
        # Reset vmcompute failure actions to Windows default
        & sc.exe failure vmcompute actions= "" reset= 0 2>&1 | Out-Null
        Log "vmcompute failure recovery: Reset to defaults" -Colour Green -Indent

        # Reset CoworkVMService failure actions (v4.8.0)
        & sc.exe failure $ServiceName actions= "" reset= 0 2>&1 | Out-Null
        Log "$ServiceName failure recovery: Reset to defaults" -Colour Green -Indent

        # Remove ServicesPipeTimeout only if we set it (value is exactly 120000)
        try {
            $current = (Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control" -Name ServicesPipeTimeout -ErrorAction SilentlyContinue).ServicesPipeTimeout
            if ($null -ne $current -and $current -eq 120000) {
                Remove-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control" -Name ServicesPipeTimeout -ErrorAction Stop
                Log "ServicesPipeTimeout: Removed (was 120000ms, set by this script)" -Colour Green -Indent
            } elseif ($null -ne $current) {
                Log "ServicesPipeTimeout: Left at ${current}ms (not set by this script)" -Colour DarkGray -Indent
            } else {
                Log "ServicesPipeTimeout: Not set" -Colour DarkGray -Indent
            }
        } catch {
            Log "Could not remove ServicesPipeTimeout, not critical" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not revert HCS configuration, not critical" -Colour DarkGray -Indent
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
    } catch { $null = $_ }
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
    $oldWatchdog = Join-Path $ClaudeAppData "cowork-watchdog.ps1"
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
        $launcherCmd = Join-Path $ClaudeAppData "Launch-Claude-Admin.cmd"
        $launcherPs1 = Join-Path $ClaudeAppData "Launch-Claude-Admin.ps1"
        foreach ($lf in @($launcherCmd, $launcherPs1)) {
            if (Test-Path $lf) {
                Remove-Item $lf -Force -ErrorAction SilentlyContinue
                Log "Removed launcher: $lf" -Colour Green -Indent
            }
        }
    } catch {
        Log "Could not fully remove elevation config, not critical" -Colour DarkGray -Indent
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
        } catch { $null = $_ }
        try {
            $fat = (Get-ItemProperty -Path $policyPath -ErrorAction Stop).FilterAdministratorToken
            if ($null -ne $fat -and $fat -eq 0) {
                Remove-ItemProperty -Path $policyPath -Name "FilterAdministratorToken" -Force -ErrorAction Stop
                Log "FilterAdministratorToken: Removed (restored default)" -Colour Green -Indent
                $reverted = $true
            }
        } catch { $null = $_ }
        if (-not $reverted) {
            Log "Token policy was not modified by this script" -Colour DarkGray -Indent
        } else {
            Log "A reboot is required for token policy changes to take effect" -Colour DarkYellow -Indent
        }
    } catch {
        Log "Could not revert token policy, not critical" -Colour DarkGray -Indent
    }

    Write-Host ""
    Write-Host "  +----------------------------------------------+" -ForegroundColor Green
    Write-Host "  |           UNDO COMPLETE                      |" -ForegroundColor Green
    Write-Host "  +----------------------------------------------+" -ForegroundColor Green
    Write-Host ""
    Write-Host "  NOTE: A reboot is recommended for all changes to take full effect." -ForegroundColor DarkGray

} else {

    # ================================================================
    # SETUP MODE
    # ================================================================
    $steps = 26

    # Claude Desktop may not be installed yet, so its AppData folder may not
    # exist. Create it up front -- several steps below write state files there.
    if (-not (Initialize-ClaudeAppData)) {
        Log "[!] Could not create $ClaudeAppData" -Colour Yellow
        Log "Backups and state files will be skipped; setup continues" -Colour DarkGray -Indent
        Write-Host ""
    }

    # ----------------------------------------------------------------
    # 1. Power plan -- Ultimate Performance (with deduplication)
    # ----------------------------------------------------------------
    Step 1 $steps "Configuring power plan..."
    try {
        # Back up current plan
        $currentPlan = Get-ActivePlanGuid
        if ($currentPlan -match "^[0-9a-fA-F\-]{36}$") {
            if (Save-StateFile -Path $BackupFile -Value $currentPlan -Label "power plan backup") {
                Log "Backed up current plan: $currentPlan" -Colour DarkGray -Indent
            }
        }

        $ultimateGuid = "e9a42b02-d5df-448d-aa00-03f14749eb61"
        $highPerfGuid = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"

        # Find all existing "Ultimate Performance" plans
        $planLines = powercfg /list | Where-Object { $_ -match 'Ultimate Performance' }
        $ultimatePlans = @()
        foreach ($line in $planLines) {
            if ($line -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})') {
                $ultimatePlans += $Matches[1]
            }
        }

        if ($ultimatePlans.Count -gt 0) {
            # Keep the first one, delete the rest
            $keepGuid = $ultimatePlans[0]
            if ($ultimatePlans.Count -gt 1) {
                $deleteCount = 0
                for ($i = 1; $i -lt $ultimatePlans.Count; $i++) {
                    powercfg /delete $ultimatePlans[$i] 2>$null
                    $deleteCount++
                }
                Log "Cleaned up $deleteCount duplicate Ultimate Performance plan(s)" -Colour DarkGray -Indent
            }
            powercfg /setactive $keepGuid
            Log "Set to: Ultimate Performance" -Colour Green -Indent
        } else {
            # No Ultimate Performance plan exists -- create one
            $dupResult = powercfg /duplicatescheme $ultimateGuid 2>&1
            if ($LASTEXITCODE -eq 0 -and $dupResult -match '([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})') {
                $newGuid = $Matches[1]
                powercfg /setactive $newGuid
                Log "Set to: Ultimate Performance (created)" -Colour Green -Indent
            } else {
                # Ultimate Performance not available on this system -- fall back to High Performance
                powercfg /setactive $highPerfGuid
                Log "Set to: High Performance (Ultimate not available)" -Colour Green -Indent
            }
        }

        # Verify activation
        $verifyPlan = Get-ActivePlanGuid
        $verifyName = (powercfg /getactivescheme) -replace '.*\(', '' -replace '\).*', ''
        Log "Active plan verified: $verifyName ($verifyPlan)" -Colour DarkGray -Indent
    } catch {
        Log "Could not configure power plan, not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 2. Disable sleep on AC
    # ----------------------------------------------------------------
    Step 2 $steps "Disabling sleep on AC power..."
    try {
        powercfg /change standby-timeout-ac 0
        Log "Sleep on AC: Never" -Colour Green -Indent
    } catch {
        Log "Could not change sleep timeout, not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 3. Disable hibernate (also kills Fast Startup)
    # ----------------------------------------------------------------
    Step 3 $steps "Disabling hibernate..."
    try {
        powercfg /h off
        Log "Hibernate: Off" -Colour Green -Indent
    } catch {
        Log "Could not disable hibernate, not critical" -Colour DarkGray -Indent
    }

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
        Log "Could not change USB setting, not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 5. Disable hard disk sleep + PCI-E power management on AC
    # ----------------------------------------------------------------
    Step 5 $steps "Disabling disk sleep and PCI-E power management on AC..."
    try {
        powercfg /change disk-timeout-ac 0
        Log "Hard disk sleep on AC: Never" -Colour Green -Indent
    } catch {
        Log "Could not change disk timeout, not critical" -Colour DarkGray -Indent
    }

    try {
        $activePlan = Get-ActivePlanGuid
        powercfg /setacvalueindex $activePlan 501a4d13-42af-4429-9fd1-a8218c268e20 ee12f906-d277-404b-b6da-e5fa1a576df5 0
        powercfg /setactive $activePlan
        Log "PCI-E link state power management on AC: Off" -Colour Green -Indent
    } catch {
        Log "Could not change PCI-E setting, not critical" -Colour DarkGray -Indent
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
            Log "Fast Startup registry key not found, may not be supported" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not disable Fast Startup, not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 7. Disable Connected Standby / Modern Standby
    # ----------------------------------------------------------------
    Step 7 $steps "Disabling Connected Standby / Modern Standby..."
    try {
        $csRegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Power"
        $csKey = Get-ItemProperty -Path $csRegPath -Name "CsEnabled" -ErrorAction SilentlyContinue
        # $null on the LEFT of both comparisons. With $null on the right,
        # PowerShell's comparison operators treat an array-valued left side as
        # a filter and return the non-matching elements rather than a boolean,
        # so the test can quietly produce the wrong answer.
        if ($null -ne $csKey -and $null -ne $csKey.CsEnabled) {
            Save-StateFile -Path $CsBackupFile -Value $csKey.CsEnabled.ToString() -Label "Connected Standby backup" | Out-Null
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
        Log "Could not change Connected Standby, not critical" -Colour DarkGray -Indent
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
            } catch { $null = $_ }
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
            } catch { $null = $_ }
        }
        if ($nicCount -gt 0) {
            Log "Disabled power saving on $nicCount adapter(s)" -Colour Green -Indent
            Log "A reboot is required for PnPCapabilities changes to take effect" -Colour DarkGray -Indent
        } else {
            Log "No physical network adapters found" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not change network adapter settings, not critical" -Colour DarkGray -Indent
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
        Log "Could not change processor state, not critical" -Colour DarkGray -Indent
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
                    Log "VM '$vmName' is running, Dynamic Memory will be disabled on next restart" -Colour DarkYellow -Indent
                    $flagFile = Join-Path $ClaudeAppData "disable-dynamic-memory.flag"
                    if (Save-StateFile -Path $flagFile -Value $vmName -Label "dynamic memory flag") {
                        Log "Flag written: $flagFile" -Colour DarkGray -Indent
                    }
                }
            } else {
                Log "Dynamic Memory already disabled for VM '$vmName'" -Colour DarkGray -Indent
            }
        } else {
            Log "No Claude VM found (Hyper-V module may not be available)" -Colour DarkGray -Indent
            Log "This is normal if Cowork hasn't been used yet" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not configure VM memory, Hyper-V module may not be installed" -Colour DarkGray -Indent
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
                } catch { $null = $_ }
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
        Log "Could not set process priority, not critical" -Colour DarkGray -Indent
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
            Log "vmcompute failure recovery set (could not verify, non-critical)" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not configure vmcompute failure recovery, not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 13. Configure CoworkVMService recovery (v4.8.0)
    # ----------------------------------------------------------------
    Step 13 $steps "Configuring CoworkVMService recovery..."
    try {
        $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($svc) {
            & sc.exe failure $ServiceName reset= 300 actions= restart/30000/restart/60000/restart/120000 2>&1 | Out-Null
            $verifyResult = & sc.exe qfailure $ServiceName 2>&1
            if ($verifyResult -match "RESTART") {
                Log "$ServiceName failure recovery: restart after 30s/60s/120s (reset 300s)" -Colour Green -Indent
            } else {
                Log "$ServiceName failure recovery set (could not verify)" -Colour DarkGray -Indent
            }
        } else {
            Log "$ServiceName not installed, skipping" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not configure CoworkVMService recovery, not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 14. Pre-emptive HCS state cleanup (v4.8.0)
    # ----------------------------------------------------------------
    Step 14 $steps "Pre-emptive HCS state cleanup..."
    # Three corrections in v6.0.0:
    #   1. hcsdiag has no "close" verb. The verbs are list, exec, console, read,
    #      write, share and kill. Every call this block made was a silent no-op
    #      that then logged success.
    #   2. The parse was wrong, in its own way, and is now Get-CoworkHcsGuids in
    #      the shared ClaudeEnv region. Fix and Watch had their own copies with
    #      their own variations on the same two mistakes.
    #   3. It ran unconditionally. If Claude was open, "stale" meant the VM the
    #      user was actively working in. Now it only acts when Claude is closed.
    $hcsExe = $script:ClaudeEnv.HcsDiagPath
    if ($hcsExe) {
        try {
            $claudeLive = @(Get-Process -Name claude -ErrorAction SilentlyContinue).Count -gt 0
            $hcsList = & $hcsExe list 2>&1 | Out-String
            $guids = Get-CoworkHcsGuids -ListOutput ([string]$hcsList)

            if ($guids.Count -eq 0) {
                Log "HCS state clean" -Colour Green -Indent
            } elseif ($claudeLive) {
                Log "$($guids.Count) cowork-vm instance(s) present but Claude is running, leaving them alone" -Colour DarkGray -Indent
            } else {
                foreach ($g in $guids) {
                    if ($PSCmdlet.ShouldProcess($g, "hcsdiag kill")) {
                        & $hcsExe kill $g 2>&1 | Out-Null
                        if ($LASTEXITCODE -eq 0) {
                            Log "Killed stale HCS compute system: $g" -Colour Green -Indent
                        } else {
                            Log "hcsdiag kill returned $LASTEXITCODE for ${g}" -Colour DarkYellow -Indent
                        }
                    }
                }
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
        Log "Could not set ServicesPipeTimeout, not critical" -Colour DarkGray -Indent
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
                        Log "No IPv4 address on Hyper-V adapter, NAT not needed yet" -Colour DarkGray -Indent
                    }
                } else {
                    Log "No Hyper-V virtual adapter found, NAT not needed yet" -Colour DarkGray -Indent
                }
            } else {
                Log "No internal Hyper-V switch found, NAT not needed yet" -Colour DarkGray -Indent
            }
            Log "Health monitor will auto-repair NAT if it disappears" -Colour DarkGray -Indent
        }
    } catch {
        Log "Get-NetNat not available, skipping NAT check" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 17. Windows Firewall policy verification
    # ----------------------------------------------------------------
    Step 17 $steps "Checking Windows Firewall policies..."
    try {
        # Check if local firewall rules are being applied (Group Policy can block them)
        $fwProfiles = Get-NetFirewallProfile -ErrorAction SilentlyContinue
        $issues = @()
        # $fwProfile, not $profile. $profile is a PowerShell automatic variable
        # holding the path to the current profile script, and assigning to it
        # inside a loop clobbers it for the rest of the session.
        foreach ($fwProfile in $fwProfiles) {
            if ($fwProfile.Enabled -and -not $fwProfile.AllowLocalFirewallRules) {
                $issues += $fwProfile.Name
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
        Log "Could not check firewall policies, not critical" -Colour DarkGray -Indent
    }

    # ----------------------------------------------------------------
    # 18. Storage location detection
    # ----------------------------------------------------------------
    Step 18 $steps "Checking workspace storage location..."
    $vmCachePath = $script:ClaudeEnv.VmCachePath
    $storageWarnings = @()

    # Both checks used to read $env:APPDATA. They now read the discovered Claude
    # folder, which is where the VM cache actually lives. On a default install
    # the two are the same; if a user has redirected the folder they are not,
    # and the redirected one is the answer that matters.
    $cloudPaths = @("OneDrive", "Google Drive", "Dropbox", "iCloud", "Box")
    foreach ($cp in $cloudPaths) {
        if ($ClaudeAppData -match [regex]::Escape($cp)) {
            $storageWarnings += "The Claude folder is inside a '$cp' sync folder, this causes mount failures"
        }
    }

    # Check if the Claude folder is on an external/USB drive
    try {
        $appDataDrive = (Split-Path $ClaudeAppData -Qualifier) + "\"
        $driveInfo = Get-Volume -DriveLetter ($appDataDrive[0]) -ErrorAction SilentlyContinue
        if ($driveInfo) {
            $diskNumber = (Get-Partition -DriveLetter ($appDataDrive[0]) -ErrorAction SilentlyContinue).DiskNumber
            if ($null -ne $diskNumber) {
                $disk = Get-Disk -Number $diskNumber -ErrorAction SilentlyContinue
                if ($disk -and $disk.BusType -match "USB|Thunderbolt|1394") {
                    $storageWarnings += "The Claude folder is on an external $($disk.BusType) drive, use a local SSD instead"
                }
            }

            # Check if it's a network drive
            $logDisk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$($appDataDrive.TrimEnd('\'))'" -ErrorAction SilentlyContinue
            if ($logDisk -and $logDisk.DriveType -eq 4) {
                $storageWarnings += "The Claude folder is on a network drive, VirtioFS requires local storage"
            }
        }
    } catch { $null = $_ }

    # Check the VM cache path specifically for problematic locations
    if ($vmCachePath -and (Test-Path $vmCachePath)) {
        try {
            $vmDrive = (Split-Path $vmCachePath -Qualifier) + "\"
            $vmDriveType = (Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$($vmDrive.TrimEnd('\'))'" -ErrorAction SilentlyContinue).DriveType
            if ($vmDriveType -eq 4) {
                $storageWarnings += "VM cache is on a network drive, this will cause failures"
            }
        } catch { $null = $_ }
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
                        } catch { $null = $_ }
                    } else {
                        Log "Clock drift: ${drift}s (within tolerance)" -Colour Green -Indent
                    }
                } else {
                    Log "Could not measure clock drift (NTP server unreachable?)" -Colour DarkGray -Indent
                }
            } catch {
                Log "Could not check clock drift, not critical" -Colour DarkGray -Indent
            }
        } else {
            Log "W32Time service not found, clock sync may not be configured" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not check time sync, not critical" -Colour DarkGray -Indent
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
    } catch { $null = $_ }

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
                    $ClaudeAppData,
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
                    $requiredProcs = @("vmwp.exe", "vmms.exe", "vmcompute.exe", "$ServiceExe.exe")
                    try {
                        foreach ($rp in $requiredProcs) {
                            Add-MpPreference -ExclusionProcess $rp -ErrorAction SilentlyContinue
                        }
                        Log "  + Process exclusions: $($requiredProcs -join ', ')" -Colour Green -Indent
                    } catch { $null = $_ }
                    # Verify process exclusions were applied (v4.8.0)
                    try {
                        $procExclusions = (Get-MpPreference -ErrorAction SilentlyContinue).ExclusionProcess
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
                    } catch { $null = $_ }
                } else {
                    Log "Defender exclusions already configured" -Colour Green -Indent
                }
            } catch {
                Log "Could not check Defender exclusions, not critical" -Colour DarkGray -Indent
            }
        } else {
            # Third-party AV, can only advise
            Log "Recommended exclusion paths for your AV:" -Colour DarkYellow -Indent
            Log "  - $ClaudeAppData\" -Colour White -Indent
            Log "  - $env:ProgramFiles\Hyper-V\" -Colour White -Indent
            Log "  - $env:SystemRoot\System32\vmwp.exe" -Colour White -Indent
            Log "  - $env:SystemRoot\System32\vmms.exe" -Colour White -Indent
            Log "  - Process: vmcompute.exe, $ServiceExe.exe" -Colour White -Indent
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

    # Check WSL feature (the DISM module is absent on some Windows editions,
    # which is a command-resolution error that -ErrorAction cannot suppress)
    $wslFeature = $null
    try {
        $wslFeature = Get-WindowsOptionalFeature -Online -FeatureName Microsoft-Windows-Subsystem-Linux -ErrorAction SilentlyContinue
    } catch {
        Log "Could not query Windows optional features, skipping WSL check" -Colour DarkGray -Indent
    }
    if ($wslFeature -and $wslFeature.State -eq 'Enabled') {
        $wsl2Warnings += "WSL feature is enabled"

        # Check for running distros
        if (Test-Path "$env:SystemRoot\System32\wsl.exe") {
            try {
                $distroOutput = & wsl -l -v 2>$null
                if ($distroOutput) {
                    $running = $distroOutput | Where-Object { $_ -match 'Running' -and $_ -match '\s2\s' }
                    if ($running) {
                        $wsl2Warnings += "WSL2 distros are actively running, may conflict with Claude's VM"
                        $wsl2Warnings += "If Cowork has issues, try: wsl --shutdown"
                    }
                }
            } catch { $null = $_ }
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
    $myDir = $null
    $watchScript = $null
    try {
        if ($PSCommandPath) { $myDir = Split-Path $PSCommandPath -Parent }
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
        $oldWatchdog = Join-Path $ClaudeAppData "cowork-watchdog.ps1"
        if (Test-Path $oldWatchdog) {
            Remove-Item $oldWatchdog -Force -ErrorAction SilentlyContinue
            Log "Removed old basic watchdog script" -Colour DarkGray -Indent
        }
    } catch {
        Log "Could not locate health monitor script, skipping" -Colour DarkGray -Indent
    }

    if ($watchScript) {
        try {
            try {
                Unregister-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction SilentlyContinue
            } catch { $null = $_ }

            # Kill any running health monitor before replacing
            try {
                Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" -ErrorAction SilentlyContinue |
                    Where-Object { $_.CommandLine -match "Watch-ClaudeHealth" } |
                    ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
            } catch { $null = $_ }

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

            try { Start-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue } catch { $null = $_ }

            Log "Health monitor installed (starts 120s after logon, polls every 30s)" -Colour Green -Indent
            Log "Task: Task Scheduler > $TaskPath$TaskName" -Colour DarkGray -Indent
            Log "Script: $watchScript" -Colour DarkGray -Indent
            Log "Logs: $ClaudeAppData\watch-logs\" -Colour DarkGray -Indent
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
    try {
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
    } catch {
        Log "Could not locate fix script, skipping boot task" -Colour DarkGray -Indent
    }

    if ($fixScript) {
        try {
            try {
                Unregister-ScheduledTask -TaskName $BootTaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction SilentlyContinue
            } catch { $null = $_ }

            # Wrap in a delayed command: wait 45s after logon before running BootPrep.
            # Increased from 30s to avoid race with Claude auto-launch VM construction.
            # Pre-check: skip if the VM service is already running (Claude got here
            # first) or if Claude itself is running (do not disrupt active session).
            # BootPrep is non-destructive: only restarts vmcompute, does not kill Claude.
            #
            # The service name is interpolated from discovery at install time rather
            # than written as a literal. The generated task is a standalone script
            # with no access to the discovery region, so the value has to be baked in
            # here, on the machine the task is being installed on.
            $delayedCmd = @"
Start-Sleep -Seconds 45
`$svc = Get-Service -Name '$ServiceName' -ErrorAction SilentlyContinue
if (`$svc -and `$svc.Status -eq 'Running') {
    # Service already running, Claude got here first, do not restart vmcompute
    exit 0
}
`$claude = Get-Process -Name 'claude' -ErrorAction SilentlyContinue
if (`$claude) {
    # Claude is running, do not disrupt
    exit 0
}
& '$fixScript' -BootPrep -Quiet
"@
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
                -Description "Non-destructive vmcompute restart at logon (45s delay, skips if Claude already running)." `
                -Force | Out-Null

            Log "Boot-fix task created (runs 45s after logon, skips if Claude running)" -Colour Green -Indent
            Log "Task: Task Scheduler > $TaskPath$BootTaskName" -Colour DarkGray -Indent
            Log "Runs: $fixScript -BootPrep -Quiet (after 45s delay, with service pre-check)" -Colour DarkGray -Indent
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
    try {
        if ($myDir) {
            $candidate = Join-Path $myDir "Fix-ClaudeDesktop.bat"
            if (Test-Path $candidate) { $fixBat = $candidate }
        }
        if (-not $fixBat -and $fixScript) {
            $fixBatCandidate = Join-Path (Split-Path $fixScript -Parent) "Fix-ClaudeDesktop.bat"
            if (Test-Path $fixBatCandidate) { $fixBat = $fixBatCandidate }
        }
    } catch {
        Log "Could not locate fix launcher, skipping shortcuts" -Colour DarkGray -Indent
    }

    $shell = $null
    if ($fixBat) {
        try {
            $shell = New-Object -ComObject WScript.Shell
        } catch {
            Log "Windows Script Host is unavailable, cannot create shortcuts" -Colour DarkYellow -Indent
        }
    }

    if ($fixBat -and $shell) {

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
            Log "Skipped, requires admin privileges (run as Administrator to enable)" -Colour DarkYellow -Indent
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
        try { Unregister-ScheduledTask -TaskName $elevTaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction SilentlyContinue } catch { $null = $_ }

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
        Initialize-ClaudeAppData | Out-Null
        $launcherCmd = Join-Path $ClaudeAppData "Launch-Claude-Admin.cmd"
        $cmdContent = @"
@echo off
REM Claude Desktop (Admin) Launcher
REM Auto-generated by Prevent-ClaudeIssues.ps1 v$ToolkitVersion
REM Triggers the LaunchClaudeAdmin scheduled task (runs with full admin token).
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
        $sc.WorkingDirectory = $ClaudeAppData
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
        Log "No UAC prompt, task runs with full admin token automatically" -Colour DarkGray -Indent
        Log "Survives Claude updates (detects MSIX, traditional install, or running process)" -Colour DarkGray -Indent

    } catch {
        Log "Could not configure elevation, not critical: $($_.Exception.Message)" -Colour DarkGray -Indent
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
        Log "EnableLUA remains 1 (UAC stays on, Store apps keep working)" -Colour DarkGray -Indent
    } catch {
        Log "Could not configure token policy, not critical: $($_.Exception.Message)" -Colour DarkGray -Indent
    }

    # ================================================================
    # Summary
    # ================================================================
    Write-Host ""
    Write-Host "  +----------------------------------------------+" -ForegroundColor Green
    Write-Host "  |           SETUP COMPLETE                     |" -ForegroundColor Green
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
    Write-Host "  |           UNEXPECTED ERROR                   |" -ForegroundColor Red
    Write-Host "  +----------------------------------------------+" -ForegroundColor Red
    Write-Host ""
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Setup stopped early, so some steps did not run." -ForegroundColor DarkYellow
    Write-Host "  Nothing was left half-applied, rerun the script once the" -ForegroundColor DarkGray
    Write-Host "  cause above is resolved, or open an issue at" -ForegroundColor DarkGray
    Write-Host "  https://github.com/JesperLive/ClaudeFix/issues" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "  Press any key to close..." -ForegroundColor DarkGray
[void][System.Console]::ReadKey($true)
