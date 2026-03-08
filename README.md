# Claude Desktop / Cowork VM Fix for Windows

**Fix "VirtioFS mount failed" and "HCS operation failed" crashes in Claude Desktop on Windows -- without rebooting.**

If you're seeing any of these errors in Claude Desktop's Cowork mode, this toolkit will fix and prevent them:

```
RPC error -1: failed to ensure virtiofs mount: Plan9 mount failed: bad address
```
```
Workspace failed to start
```
```
Setting up workspace... (stuck forever)
```
```
HCS operation failed: failed to create compute system
```
```
VM is already running
```

Two scripts: one **prevents** the crash, the other **fixes** it when it happens anyway.

| Script | Purpose | Run when |
|--------|---------|----------|
| `Prevent-ClaudeIssues.bat` | Configure Windows to minimise crashes | Once |
| `Fix-ClaudeDesktop.bat` | Reset and relaunch when it breaks | Every time it breaks |
| `Watch-ClaudeHealth.bat` | Foreground health monitor (auto-installed by Prevent) | Optional / manual |

---

## Quick Start

1. Download all files into one folder (e.g. `C:\ClaudeFix\`)
2. Run **`Prevent-ClaudeIssues.bat`** once (configures Windows, creates watchdog, boot-fix task, and shortcuts)
3. Use the Desktop shortcut or search "Fix Claude Desktop" in Start when Cowork breaks
4. Right-click the Start Menu entry and select **Pin to taskbar** for one-click access

After running the prevention script, Claude will be automatically repaired at every logon and a background health monitor will detect and auto-fix VirtioFS crashes within seconds.

---

## Compatibility

| | Supported |
|---|---|
| Windows 10 Pro/Enterprise/Education (build 19041+) | Yes |
| Windows 11 Pro/Enterprise/Education | Yes |
| Windows 10/11 Home | No (Cowork requires Hyper-V, which is not available on Home) |
| Windows 7 / 8 / 8.1 | No (Cowork requires Hyper-V, Claude requires Win 10 19041+) |
| MSIX / Microsoft Store install | Yes |
| Traditional (.exe) install | Yes |
| PowerShell 5.1 | Yes (ships with Windows) |
| PowerShell 7+ | Yes |

| Script | Admin required? |
|--------|----------------|
| `Fix-ClaudeDesktop` | No (recommended, but works without -- uses force-kill instead of service control) |
| `Prevent-ClaudeIssues` | **Yes** (power settings, registry, scheduled tasks all require admin) |
| `Watch-ClaudeHealth` | No (auto-installed as elevated task by Prevent) |

---

## The Problem

Claude Desktop runs a lightweight Hyper-V VM to power Cowork mode. The VM uses VirtioFS (Plan9 protocol) to share the filesystem between host and guest. This mount frequently breaks:

```
RPC error -1: failed to ensure virtiofs mount: Plan9 mount failed: bad address
```

Closing and reopening Claude doesn't help -- the `CoworkVMService` stays in a broken state. The typical "fix" before these scripts was rebooting the entire PC.

### Why It Happens

The VirtioFS connection degrades when:

- Windows enters sleep or hibernate (the VM doesn't survive the transition)
- The VM sits idle for extended periods
- Power management reduces PCI-E or USB link states mid-operation
- `CoworkVMService` crashes and doesn't auto-recover
- Stale VM cache files from a previous session conflict with the new boot
- Hyper-V Dynamic Memory reclaims shared memory regions from the VM
- WinNAT rules disappear (VPN reconnect, network reset), killing VM connectivity
- Antivirus filter drivers interfere with VirtioFS disk operations
- Host clock drift causes TLS and API failures inside the VM

### HCS Failures

A separate class of failure occurs before the VM even boots. The Host Compute Service (HCS, `vmcompute.exe`) manages Hyper-V compute systems. When it fails, Claude shows:

```
Failed to start Claude's workspace
```

The underlying HCS error (`failed to create compute system: HcsWaitForOperationResult failed`) can be caused by:

- Handle leaks in `vmcompute.exe` after many VM start/stop cycles
- The `vmcompute` service crashing or entering a stuck state
- Boot race conditions where services start before the kernel is ready
- Antivirus filter drivers interfering with compute system creation

### Tracked Issues

- [#26554](https://github.com/anthropics/claude-code/issues/26554) -- VirtioFS mount fails with "bad address"
- [#27576](https://github.com/anthropics/claude-code/issues/27576) -- Mount failure after ~1 hour of use
- [#28890](https://github.com/anthropics/claude-code/issues/28890) -- Mount goes stale after idle
- [#29587](https://github.com/anthropics/claude-code/issues/29587) -- Cowork fails after brief use
- [#29848](https://github.com/anthropics/claude-code/issues/29848) -- Recurring VM crashes
- [#31520](https://github.com/anthropics/claude-code/issues/31520) -- Community recovery script for VirtioFS failures (ClaudeFix covers all steps and more)
- [#31703](https://github.com/anthropics/claude-code/issues/31703) -- HCS/VM service failures on v1.1.5368 (still open as of March 2026)

---

## Fix-ClaudeDesktop

One-click fix when Claude Desktop / Cowork is broken. No reboot needed.

### What It Does

| Step | Action |
|------|--------|
| 1 | Captures Claude.exe path, then force-kills all claude.exe processes |
| 2 | Stops CoworkVMService (graceful with admin, force-kill without) |
| 3 | Checks for HCS errors and restarts vmcompute service if needed (admin only) |
| 4 | Verifies no orphan processes remain |
| 5 | Kills orphan HCS compute systems via hcsdiag and Hyper-V cmdlets |
| 6 | Deletes stale VM cache (claude-code-vm and vm_bundles) to force a clean rebuild |
| 7 | Restarts CoworkVMService (admin) or defers to Claude auto-restart (non-admin) |
| 8 | Relaunches Claude Desktop with elevated privileges via scheduled task (Method 0), falling back to MSIX shell protocol or direct exe launch |
| 9 | Monitors cowork_vm_node.log for boot completion, confirms workspace is ready |

**Step 8** first attempts to relaunch Claude with elevated privileges via the `LaunchClaudeAdmin` scheduled task created by Prevent (Method 0). This gives Claude a full admin token without a UAC prompt. If the task doesn't exist or fails, it falls through to three standard methods: Method A launches MSIX installs via `shell:AppsFolder` protocol (no duplicate taskbar icons), Method B launches traditional `.exe` installs directly, and Method C uses Start Menu shortcuts as a last resort. All methods respect `-WhatIf` and each has a `$launched` guard to prevent double-launch.

**Step 5** terminates orphan HCS compute systems that survive service shutdown. When CoworkVMService stops, the underlying Hyper-V VM may remain registered in HCS, causing "VM is already running" errors on restart. The script first uses `hcsdiag list` to find any claude/cowork compute systems and `hcsdiag kill` to terminate them (admin only). As a fallback, it uses `Stop-VM -TurnOff -Force` via Hyper-V cmdlets. Both methods are tried because hcsdiag operates at the HCS layer (catching lightweight containers) while Stop-VM operates at the Hyper-V management layer (catching full VMs). This step is non-fatal -- failures don't block the rest of the fix.

**Step 3** checks recent Windows Event Log entries and Claude logs for HCS error patterns (`HCS operation failed`, `failed to create compute system`, `HcsWaitForOperationResult`). If detected and running as admin, it stops and restarts the `vmcompute` service. If `vmcompute` fails to restart within 15 seconds, it also restarts the `vmms` (Virtual Machine Management) service as a fallback. This step is wrapped in a try/catch so failures don't block the rest of the fix process. Without admin, HCS errors are logged but require manual elevation.

**Step 9** monitors the VM boot log for definitive completion markers (`"Startup complete"`, `"[Keepalive]"`), showing real-time progress through the boot stages. Falls back to Hyper-V heartbeat checks and directory monitoring if logs are unavailable. After completion, the PowerShell window is brought to the foreground and the taskbar icon flashes until you dismiss it.

**Step 6 note:** Purging the VM cache (`vm_bundles`) forces Claude to re-download approximately 2-3 GB of VM resources on the next launch. This is necessary for a clean recovery but means the first Cowork session after a fix will take longer to start. If you're running the fix script frequently, be aware of this download cost.

### What It Does NOT Touch

- `claude_desktop_config.json` -- your MCP servers and settings are safe
- `config.json` -- app configuration is safe
- Conversations -- stored server-side, not in the local VM cache

### Parameters

| Parameter | Description |
|-----------|-------------|
| `-SkipLaunch` | Reset the VM but don't relaunch Claude |
| `-Quiet` | Suppress the "press any key" prompt (for scheduled tasks) |
| `-KeepCache` | Skip the VM cache purge (avoids ~2-3 GB re-download). Use when running Fix frequently. If the fix fails with `-KeepCache`, run again without it. |
| `-WhatIf` | Dry run -- show what would happen without changing anything |

### Diagnostics

Each run writes a timestamped log to `%APPDATA%\Claude\fix-logs\`. Recent `CoworkVMService` errors from the Windows Event Log are shown in the summary.

---

## Prevent-ClaudeIssues

Run once. Configures Windows to keep the Cowork VM alive as long as possible.

### What It Does

| Step | Setting | Value | Why |
|------|---------|-------|-----|
| 1 | Power plan | Ultimate / High Performance | Prevents aggressive power saving that kills VM mounts |
| 2 | Sleep on AC | Never | Sleep kills VirtioFS mounts -- the #1 cause of crashes |
| 3 | Hibernate | Off | Incompatible with running Hyper-V VMs |
| 4 | USB selective suspend | Disabled (AC) | Can interrupt VM communication channels |
| 5 | Hard disk sleep + PCI-E | Never / Off (AC) | Prevents disk spin-down and link state changes during VM I/O |
| 6 | Fast Startup | Disabled | Prevents kernel hibernate on shutdown -- services don't reinitialise cleanly |
| 7 | Connected Standby | Disabled | Modern Standby can enter low-power states even with sleep "disabled" |
| 8 | Network adapter power saving | Disabled | Prevents Windows from sleeping virtual network adapters used by Hyper-V |
| 9 | Processor minimum state | 100% (AC) | Prevents CPU throttling that can starve the VM |
| 10 | Hyper-V VM memory | Pinned (no ballooning) | Prevents memory reclaim from invalidating VirtioFS shared regions |
| 11 | VM worker priority | AboveNormal | Prevents host from deprioritizing vmwp.exe under load |
| 12 | HCS service recovery | Auto-restart 30s/60s/120s | Configures `vmcompute` to auto-restart on failure with escalating delays |
| 13 | Service startup timeout | 120000ms | Prevents boot race conditions where services start before dependencies are ready |
| 14 | WinNAT rules | Verified / repaired | Ensures VM has outbound network connectivity |
| 15 | Firewall policies | Checked | Detects Group Policy blocking Hyper-V network rules |
| 16 | Storage location | Checked | Warns if workspace is on cloud-sync, USB, or network drive |
| 17 | Time synchronisation | Verified | Ensures NTP is running and clock drift is within tolerance |
| 18 | Antivirus exclusions | Configured / advised | Prevents AV filter drivers from blocking VirtioFS disk ops |
| 19 | Health monitor | Every 30 seconds | Detects VirtioFS errors and auto-runs the full fix script |
| 20 | Boot-fix task | At logon | Runs the full fix script at every logon for a clean start |
| 21 | Shortcuts | Desktop + Start Menu | Quick access to Fix-ClaudeDesktop |
| 22 | Claude elevation | Scheduled task + Desktop shortcut | Ensures Claude Desktop launches with full admin privileges. Desktop shortcut uses Claude's own icon and includes a process guard to prevent double-launch |
| 23 | Admin token policy | LocalAccountTokenFilterPolicy=1 | Disables remote/network admin token filtering -- complementary to Step 22 for tools that use COM or WMI elevation |

Battery settings are not changed -- laptop users keep normal battery behaviour.

### Hyper-V VM Memory and Worker Priority

**Dynamic Memory pinning** (step 10) -- Hyper-V's Dynamic Memory feature allows Windows to balloon memory in and out of VMs based on demand. When memory is reclaimed from the Cowork VM, VirtioFS shared memory regions can become invalid -- this is the direct cause of the `EFAULT` ("bad address") error. Disabling Dynamic Memory pins the VM's allocation so it can't be reclaimed. This requires the VM to be stopped; if it's running when you run Prevent, a flag file is written and the health monitor applies the change the next time the VM restarts.

**VM worker process priority** (step 11) -- `vmwp.exe` is the Hyper-V Virtual Machine Worker Process that hosts each VM on the host side. At Normal priority, it can be starved under heavy CPU or I/O load, causing the VirtioFS connection to stall or time out. Setting it to AboveNormal gives it scheduling preference. This is not persistent across reboots, so the health monitor re-applies it on every poll cycle.

### HCS Service Recovery

**HCS service recovery** (step 12) -- The `vmcompute` service (Host Compute Service) manages all Hyper-V compute system operations. If it crashes, every VM creation call fails with `HCS operation failed`. The script configures Windows Service Control Manager to auto-restart `vmcompute` with escalating delays: 30 seconds after the first failure, 60 seconds after the second, 120 seconds after the third. The failure counter resets after 300 seconds of healthy operation. This is a permanent OS-level setting that survives reboots.

**Service startup timeout** (step 13) -- The default `ServicesPipeTimeout` of 30 seconds can be too short on heavily loaded systems or during Windows Update reboots. If services like `vmcompute` don't start within this window, dependent services fail silently. Setting it to 120 seconds (120000ms) gives boot-time services more room. This is idempotent -- if the timeout is already >=120000ms (set by another tool), it's left untouched. Requires a reboot to take effect.

### Network and NAT

**WinNAT rules** (step 14) -- The Cowork VM needs a WinNAT rule to route traffic from its internal Hyper-V switch to the host's network. If this rule disappears (VPN reconnect, network adapter change, Windows Update), the VM silently loses all outbound connectivity. API calls fail, package downloads stall, and the workspace becomes unresponsive. The prevention script checks for existing NAT rules and auto-creates one if missing. The health monitor continuously monitors NAT health and repairs it automatically.

**Firewall policies** (step 15) -- Group Policy can set "Apply Local Firewall Rules" to disabled, which blocks the DHCP and DNS rules that Hyper-V's Host Network Service (HNS) creates for VMs. The script detects this and warns you to contact your IT admin. It also checks that Hyper-V-specific firewall rules are enabled.

### Storage, Time Sync, and Antivirus

**Storage location** (step 16) -- VirtioFS mounts fail when Claude's data directory is on a cloud-sync folder (OneDrive, Google Drive, Dropbox), an external USB drive, or a network share. The script detects these conditions and warns you to move Claude's data to a local SSD.

**Time synchronisation** (step 17) -- If the host clock drifts more than 5 seconds from NTP, Hyper-V's time synchronisation integration service can't correct the guest clock. This causes TLS certificate validation failures and API timeouts inside the VM. The script checks the W32Time service, measures actual drift, and forces a resync if needed. The health monitor continues to check every 5 minutes.

**Antivirus exclusions** (step 18) -- Antivirus filter drivers sit in the I/O path between VirtioFS and the host filesystem. They can delay or block disk operations that VirtioFS depends on, causing timeouts and mount failures. For Windows Defender, the script automatically adds exclusions for Claude's data directory, Hyper-V binaries, and the CoworkVMService process. For third-party AV products, it lists the recommended exclusion paths for you to add manually.

### The Health Monitor

A persistent background process that starts at logon and polls every 30 seconds. It monitors eight sources for VirtioFS failures:

1. **Claude log files** -- scans all `*.log` files in `%APPDATA%\Claude\logs\` for error patterns like "Plan9 mount failed" and "bad address"
2. **Service status** -- detects when `CoworkVMService` stops while `claude.exe` is still running (2 consecutive checks)
3. **Windows Event Log** -- checks for Claude-specific `CoworkVMService` errors and Hyper-V Worker/VMMS errors (2 consecutive checks, Claude-only matching)
4. **WinNAT health** -- detects missing NAT rules and auto-repairs them (every 60 seconds, warning only)
5. **Hyper-V heartbeat** -- monitors the VM's Integration Services heartbeat to detect hung VMs (3 consecutive checks)
6. **VM log staleness** -- catches silent hangs where the VM stops writing logs (5 consecutive stale checks, 5-minute threshold, only if VM was previously active)
7. **Clock drift** -- checks NTP drift every 5 minutes and auto-resyncs if >5 seconds (warning only)
8. **vmcompute health** -- monitors `vmcompute.exe` handle count every 60 seconds. Warning at 5000 handles, critical trigger at 10000 handles (2 consecutive checks required). Catches handle leaks that precede HCS failures

### Safety Features (v4.3)

The entire toolkit is designed to **never interrupt active work** — whether you're in Chat, Cowork, or Code:

- **Electron-aware activity detection** -- uses `GetWindowThreadProcessId` to correctly detect Claude's Electron renderer windows (the old `MainWindowHandle` comparison failed because Electron child processes own the visible window, not the main process)
- **Session 0 safe** -- when running as a SYSTEM scheduled task (Session 0), Win32 window/input APIs return garbage. The monitor detects this and falls back to process-only heuristics (VM log + CPU sampling)
- **Claude Code session awareness** -- detects active Code sessions via session persistence files and renderer logs (`unknown-window*.log`, `claude.ai-web*.log`). Code runs inside Claude Desktop but doesn't use the Cowork VM, so VM staleness is expected and normal during Code usage. Without this check, the monitor could kill an active Code session while trying to fix a "stale" VM that was never needed
- **Extended CPU sampling** -- the 30s grace period re-check now takes 3 samples × 1s (was 1 × 500ms) with a lower threshold (50ms, was 100ms). Code has bursty CPU with long near-zero periods during API waits; the wider sampling window catches these patterns
- **CPU sampling** -- measures actual processor time on Claude processes over 500ms. Catches active request processing even when the UI is idle (e.g., Code thinking)
- **Extended VM log window** -- activity check uses 120s window (was 30s). Code's "thinking" phases can leave the VM log quiet for minutes; the old 30s window caused false negatives
- **User input window** -- 3 minutes (was 2). More buffer for reading/reviewing before auto-fix considers you idle
- **Fix script activity guard** -- `Fix-ClaudeDesktop.ps1` itself now checks for active use when called with `-Quiet` (by the monitor or boot task). Three checks: CPU sampling, VM log, user input. Blocks and exits if anything is active. Manual runs (no `-Quiet`) always proceed
- **Boot-fix 90s delay** -- the logon task now waits 90 seconds before running, preventing a race condition where it would kill Claude Desktop as it was auto-starting at logon
- **Startup grace period** -- heuristic checks (event log, heartbeat, staleness) are skipped for the first 90 seconds after the monitor starts, preventing false triggers from pre-existing events
- **Consecutive-check gates** -- every heuristic trigger requires multiple consecutive failures before firing (service: 2, event log: 2, heartbeat: 3, staleness: 5). Only the log-file pattern check (actual VirtioFS error strings) triggers immediately
- **Tightened event log matching** -- Hyper-V VMMS events must mention "claude" or "cowork" (no generic "failed"/"unexpected" matching). Worker events: Critical/Error only
- **VM log staleness requires prior activity** -- only triggers if the VM log was previously active this session, preventing false positives in Chat mode
- **5-minute cooldown** -- between auto-fixes
- **30s pre-fix warning with notification** -- before any auto-fix, a Windows balloon notification with an audible chime warns you. You have 30 seconds to switch to Claude to cancel. If you don't, the fix proceeds. If you do switch to Claude, the fix cancels and a second notification tells you to run Fix-ClaudeDesktop.bat manually if Cowork is broken
- **Smart cancellation** -- the 30s grace period only cancels if Claude is *actively being used* (foreground window, CPU activity, VM log alive, or active Code session). General mouse/keyboard activity in other apps does **not** cancel the fix — so a genuinely hung VM still gets repaired while you're browsing or gaming
- **Default Switch NAT awareness** -- the NAT health check now recognises Hyper-V's "Default Switch" as providing NAT natively (via HNS), eliminating false "WinNAT missing" warnings on standard configurations

When a failure is detected **and the user is idle**, it shows a warning notification with a 30-second countdown, then runs `Fix-ClaudeDesktop.ps1 -Quiet`. If you switch to Claude during the countdown, the fix cancels and you're notified to run it manually. If the user was already detected as active before the countdown, it logs a `BLOCKED` message and waits.

The monitor also performs continuous maintenance: re-applying vmwp.exe AboveNormal priority on every cycle, and applying deferred Dynamic Memory changes when the VM is stopped.

The monitor uses a global mutex to ensure only one instance runs at a time. It logs its activity to `%APPDATA%\Claude\watch-logs\` (auto-cleaned after 30 days).

Visible in Task Scheduler under `\Claude\ClaudeCoworkWatchdog`. Can also be started manually with `Watch-ClaudeHealth.bat` for foreground monitoring.

### Claude Elevation and Admin Token Policy

**Claude elevation** (step 22) -- Claude Desktop is installed as an MSIX (Microsoft Store) package. By default, it launches with a standard (non-elevated) user token, even if you're an administrator. This means its child processes (including MCP servers like Desktop Commander) also run without admin privileges and cannot perform system-level operations. MSIX apps block all direct `.exe` access from `WindowsApps` (ACLs, `Start-Process -Verb RunAs`, `dir` enumeration all fail), so the only reliable approach is a **scheduled task**. The script creates a `\Claude\LaunchClaudeAdmin` task with `RunLevel=Highest` + `LogonType=Interactive`, which gives the process a full unfiltered admin token with no UAC prompt, while keeping the GUI visible in the user's desktop session. The task's action finds Claude at runtime via three methods: (1) `Get-AppxPackage` for MSIX installs, (2) common install paths for traditional `.exe` installs, (3) running-process detection as a final fallback. This survives version updates and works with any install type. A "Claude (Admin)" desktop shortcut triggers this task via `schtasks /run`. **Note:** MSIX installs will show a second Claude icon on the taskbar when launched elevated. This is unavoidable -- Windows enforces medium integrity for all shell-activated MSIX apps, so the only way to get a full admin token is to launch the `.exe` directly, which bypasses the MSIX app model's icon grouping. The scheduled task includes a process guard: if Claude is already running when the task is triggered (e.g., clicking the Desktop shortcut while Claude is open), it exits cleanly without launching a second instance. The Desktop shortcut uses Claude's actual icon (resolved at Prevent runtime via `Get-AppxPackage`); if the icon path becomes stale after a Claude update, it falls back to a generic Windows icon until Prevent is re-run.

**Admin token policy** (step 23) -- Windows filters admin tokens for local accounts during remote/network logins via `LocalAccountTokenFilterPolicy`. Setting it to `1` (along with `FilterAdministratorToken=0`) allows tools that use COM elevation, WMI, or remote PowerShell to receive full admin tokens. This is complementary to Step 22 -- the scheduled task handles the main elevation for Claude Desktop itself, while the token policy helps any tools that use COM-based or network-based elevation. UAC stays enabled and Store apps continue to work. Requires a reboot.

### Fast Startup, Connected Standby, NIC Power Saving

These three settings are the most commonly overlooked causes of VirtioFS failures:

**Fast Startup** (step 6) -- With Fast Startup on, a Windows "shutdown" is actually a kernel hibernate. Services like `CoworkVMService` don't fully reinitialise on the next boot, which can leave stale VM state behind. Disabling hibernate (`powercfg /h off`) usually kills Fast Startup too, but the script also explicitly sets the `HiberbootEnabled` registry key to 0 for safety.

**Connected Standby / Modern Standby** (step 7) -- On newer hardware (most laptops since ~2018, some desktops), Windows uses Modern Standby instead of traditional S3 sleep. Modern Standby can enter low-power states *even when sleep is "disabled"* in the power plan. Setting `CsEnabled = 0` in the registry forces traditional power behaviour. Requires a reboot.

**Network adapter power saving** (step 8) -- Windows can put physical network adapters to sleep to save power. This also affects the virtual network switch that Hyper-V uses for the VM's network bridge. The script disables power saving on all physical adapters via `PnPCapabilities`, wake-on-LAN settings, and WMI power management properties. Requires a reboot.

### The Boot-Fix Task

A scheduled task runs at every user logon as the current user (elevated). It waits **90 seconds** (to let Claude Desktop auto-start first), then executes `Fix-ClaudeDesktop.ps1 -SkipLaunch -Quiet` to ensure the VM service starts cleanly. The Fix script's own activity guard provides a second layer of protection — if Claude is already running and active after the 90s delay, the fix is blocked and exits silently.

The prevention script auto-detects `Fix-ClaudeDesktop.ps1` in the same folder. If it can't find it, this step is skipped with a warning.

Visible in Task Scheduler under `\Claude\ClaudeCoworkBootFix`.

### Shortcuts

Creates a "Fix Claude Desktop" shortcut on the Desktop and in the Start Menu. You can pin the Start Menu entry to your taskbar for one-click access.

### Undo Everything

```
.\Prevent-ClaudeIssues.ps1 -Undo
```

Restores your original power plan, re-enables hibernate, resets sleep to 30 minutes, reverts HCS service recovery configuration, removes all scheduled tasks (Watchdog, BootFix, LaunchClaudeAdmin), deletes the shortcuts and launcher scripts, removes the RUNASADMIN registry flags, and reverts admin token policy changes.

---

## Requirements

- Windows 10 (build 19041+) or Windows 11
- Claude Desktop with Cowork mode
- Hyper-V capable edition (Pro, Enterprise, Education -- not Home)
- PowerShell 5.1+ (included with Windows)
- Admin privileges: required for Prevent, optional for Fix

---

## Troubleshooting

**The fix script says "Workspace ready" but Cowork still shows an error**
Try running the fix script a second time. Some stale states need two cycles to fully clear. If it persists after 2-3 runs, reboot and let the boot-fix task handle it.

**"Failed to start Claude's workspace" with HCS error**
This is an HCS (Host Compute Service) failure, not a VirtioFS mount failure. Run `Fix-ClaudeDesktop.bat` as admin -- Step 3 will detect the HCS error and restart the `vmcompute` service. If it keeps happening, run `Prevent-ClaudeIssues.bat` to configure automatic HCS service recovery. The health monitor (installed by Prevent) also watches for `vmcompute` handle leaks that precede these failures.

**"VM is already running" after running the fix script**
The old VM wasn't fully released before the service restarted. Run the fix script again -- Step 5 now explicitly kills orphan compute systems via hcsdiag. If it persists, open an admin PowerShell and run `hcsdiag list` to see all compute systems, then `hcsdiag kill <id>` for any claude/cowork entries.

**"Service not found -- is Cowork installed?"**
The `CoworkVMService` Windows service is only installed when you've used Cowork mode at least once. Open Claude Desktop, start a Cowork session, let it install the VM components, then run the scripts.

**The prevention script asks for admin but I don't have it**
The prevention script genuinely needs admin for power settings, scheduled tasks, and registry changes. The fix script can run without admin, but the prevention script cannot. Ask your IT department for temporary elevation, or have them run `Prevent-ClaudeIssues.bat` for you.

**Connected Standby changes didn't take effect**
This setting requires a full reboot (not just sleep/wake). Shut down, wait 10 seconds, power on. Verify with `powercfg /a` -- if it no longer lists "Standby (S0 Low Power Idle)" then Modern Standby is disabled.

**The prevention script warns about WinNAT / Firewall**
If you see a WinNAT warning, the health monitor will attempt to repair it automatically. If you see a firewall warning, ask your IT admin to enable "Apply Local Firewall Rules" in Group Policy for the affected profiles.

**Antivirus warnings**
If you have third-party antivirus, add the recommended exclusion paths shown by the prevention script. For Windows Defender, exclusions are added automatically.

**Claude launches with a duplicate taskbar icon**
When using the "Claude (Admin)" shortcut, MSIX installs will show a second Claude icon on the taskbar. This is expected -- Windows enforces medium integrity for MSIX apps launched through the shell protocol (`shell:AppsFolder`), so the only way to get a full admin token is to launch the `.exe` directly, which bypasses icon grouping. Close the non-admin Claude first if you want a single icon. The regular fix script (`Fix-ClaudeDesktop`) uses the shell protocol and does not have this issue.

**Logs are piling up in `%APPDATA%\Claude\fix-logs\`**
The fix script auto-cleans logs older than 30 days. You can safely delete everything in that folder manually if needed.

---

## File List

```
Fix-ClaudeDesktop.bat       -- Fix launcher (double-click when broken)
Fix-ClaudeDesktop.ps1       -- Fix script
Prevent-ClaudeIssues.bat    -- Prevention launcher (run once)
Prevent-ClaudeIssues.ps1    -- Prevention script
Watch-ClaudeHealth.bat      -- Health monitor launcher (manual foreground mode)
Watch-ClaudeHealth.ps1      -- Health monitor (auto-detects and auto-fixes crashes)
README.md                   -- This file
LICENSE                     -- MIT licence
```

Current versions: Fix 4.5.0, Watch 4.4.0, Prevent 4.5.0

---

## Credits & Acknowledgements

This toolkit was built by combining community knowledge from multiple independent contributors who shared their findings and workarounds:

- **Jonas Kamsker** ([blog.kamsker.at](https://blog.kamsker.at/blog/cowork-windows-broken/)) -- Comprehensive diagnostics, DNS/NAT fix scripts, and VM state recovery techniques
- **Elliot Segler** ([elliotsegler.com](https://www.elliotsegler.com/fixing-claude-coworks-network-conflict-on-windows.html)) -- Network conflict resolution and HNS-based recovery for subnet collisions
- **@garabedjunior-dotcom** ([GitHub #29848](https://github.com/anthropics/claude-code/issues/29848)) -- Community troubleshooting scripts and MCP crash diagnosis
- Everyone who reported and documented VirtioFS failures in the [claude-code issue tracker](https://github.com/anthropics/claude-code/issues)

Special thanks to the community contributors on GitHub issues #26554, #27576, #28890, #29587, and #29848 whose collective debugging narrowed down the root causes.

---

## Licence

MIT

---

*Built out of frustration with rebooting the PC every time Claude decides to break.*
