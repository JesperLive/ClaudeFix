# Claude Desktop / Cowork VM Fix for Windows

**Fix "VirtioFS mount failed", "HCS operation failed", and "Failed to start Claude's workspace" errors in Claude Desktop Cowork mode on Windows — without rebooting.**

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
```
HCS operation failed: The operation failed because of a virtual disk system limitation (0x800707DE)
```
```
Failed to start Claude's workspace
```
```
VM service not running. The service failed to start.
```
```
VM boot failed: VM is already running
```
```
Request timed out: isGuestConnected
```

Four tools: one prevents the crash, one fixes it, one monitors health, and one stops Claude cleanly.

| Script | Purpose | Run when |
|--------|---------|----------|
| Prevent-ClaudeIssues.bat | Configure Windows to minimise crashes | Once |
| Fix-ClaudeDesktop.bat | Reset and relaunch when it breaks | Every time it breaks |
| Stop-ClaudeDesktop.bat | Clean shutdown without repair | When you want to fully close Claude |
| Watch-ClaudeHealth.bat | Foreground health monitor (auto-installed by Prevent) | Optional / manual |

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

### Guest Connection Timeout

A third class of failure occurs after the VM boots but the guest operating system never establishes its connection back to the host. Claude Desktop polls `cowork-svc.exe` via a named pipe (`\\.\pipe\cowork-vm-service`) every ~1 second with `isGuestConnected` RPC calls. If the guest never reports as connected, the UI shows `"Request timed out: isGuestConnected"`. This can be caused by: the VM booting but vsock/networking failing to initialise, `cowork-svc.exe` being CPU-starved by its own Authenticode signature verification loop (each poll re-hashes the 204 MB `claude.exe` binary, consuming ~960ms per call), or HCS state corruption preventing guest-host communication. The service log at `C:\ProgramData\Claude\Logs\cowork-service.log` records each RPC call. See issue #31848 for the Authenticode CPU burn.

### Tracked Issues

- [#26554](https://github.com/anthropics/claude-code/issues/26554) -- VirtioFS mount fails with "bad address" (closed March 18 2026 as "completed" -- no linked PR or fix details; underlying architecture unchanged)
- [#27576](https://github.com/anthropics/claude-code/issues/27576) -- Mount failure after ~1 hour of use
- [#28890](https://github.com/anthropics/claude-code/issues/28890) -- Mount goes stale after idle
- [#29587](https://github.com/anthropics/claude-code/issues/29587) -- Cowork fails after brief use
- [#29848](https://github.com/anthropics/claude-code/issues/29848) -- Recurring VM crashes
- [#31520](https://github.com/anthropics/claude-code/issues/31520) -- Community recovery script for VirtioFS failures (ClaudeFix covers all steps and more)
- [#31703](https://github.com/anthropics/claude-code/issues/31703) -- HCS/VM service failures on v1.1.5368 (still open as of March 2026, no Anthropic response)
- [#32172](https://github.com/anthropics/claude-code/issues/32172) -- HCS 0x800707DE construct failure after VirtioFS mount error (still open as of March 2026, no Anthropic response)
- [#29045](https://github.com/anthropics/claude-code/issues/29045) — Claude Desktop spawns Hyper-V VM on every launch, even for chat-only use
- [#27801](https://github.com/anthropics/claude-code/issues/27801) — "Failed to start Claude's workspace" — VM service not running, persists after reboot
- [#31848](https://github.com/anthropics/claude-code/issues/31848) -- cowork-svc.exe CPU burn from Authenticode re-verification on every isGuestConnected poll (still open as of March 2026)
- [#31314](https://github.com/anthropics/claude-code/issues/31314) -- cowork-svc.exe 195 MB/s sustained I/O from signature polling loop (still open as of March 2026)

---

## Fix-ClaudeDesktop

One-click fix when Claude Desktop / Cowork is broken. No reboot needed.

### Interactive Menu

When run manually (without `-Quiet` or `-Mode`), the script shows an interactive menu:

1. **Quick Fix** -- Restart services + basic repair (Steps 1-5, skip cache purge)
2. **Deep Fix** -- Full nuclear reset (all steps including cache purge)
3. **Smart Fix** -- Try quick first, escalate to deep if needed (default)
4. **Diagnostic** -- Health check only, no changes

After selecting a mode, a second menu lets you toggle options (Keep cache, Skip relaunch, WhatIf). A summary is shown before the final Y/N confirmation.

The menu is skipped when:
- `-Mode` is passed as a parameter (uses that mode directly)
- `-Quiet` is set (defaults to Smart mode, for automated callers)
- The host is non-interactive (defaults to Smart mode)

### What It Does

| Step | Action |
|------|--------|
| 0 | Pre-emptive HCS state cleanup (stale cowork-vm entries) and session file housekeeping (>7 days) |
| 1 | Captures Claude.exe path, then force-kills all claude.exe processes |
| 2 | Stops CoworkVMService (graceful with admin, force-kill without) |
| 3 | Checks for HCS errors and restarts vmcompute service if needed; escalates to vmms and HvHost (Deep mode only) |
| 4 | Verifies no orphan processes remain |
| 5 | Kills orphan HCS compute systems via hcsdiag, Hyper-V cmdlets, and hung vmwp.exe processes |
| 6 | Smart cache purge: backs up session VHDXs, nukes VM cache, restores session data; cleans temp files and legacy paths |
| 7 | Restarts CoworkVMService (admin) or defers to Claude auto-restart (non-admin); restores VHDX backups |
| 8 | Relaunches Claude Desktop with elevated privileges via scheduled task (Method 0), falling back to MSIX shell protocol or direct exe launch |
| 9 | Monitors coworkd.log and cowork-service.log for boot completion and guest connection state, confirms workspace is ready |

**Step 8** first attempts to relaunch Claude with elevated privileges via the `LaunchClaudeAdmin` scheduled task created by Prevent (Method 0). This gives Claude a full admin token without a UAC prompt. If the task doesn't exist or fails, it falls through to three standard methods: Method A launches MSIX installs via `shell:AppsFolder` protocol (no duplicate taskbar icons), Method B launches traditional `.exe` installs directly, and Method C uses Start Menu shortcuts as a last resort. All methods respect `-WhatIf` and each has a `$launched` guard to prevent double-launch.

**Step 5** terminates orphan HCS compute systems that survive service shutdown. When CoworkVMService stops, the underlying Hyper-V VM may remain registered in HCS, causing "VM is already running" errors on restart. The script first uses `hcsdiag list` to find any claude/cowork compute systems and `hcsdiag kill` to terminate them (admin only). As a fallback, it uses `Stop-VM -TurnOff -Force` via Hyper-V cmdlets, then checks for hung `vmwp.exe` processes (VM worker) and kills them via hcsdiag or force-kill. Both methods are tried because hcsdiag operates at the HCS layer (catching lightweight containers) while Stop-VM operates at the Hyper-V management layer (catching full VMs). This step is non-fatal -- failures don't block the rest of the fix.

**Step 3** checks recent Windows Event Log entries and Claude logs for HCS error patterns (`HCS operation failed`, `failed to create compute system`, `HcsWaitForOperationResult`). If detected and running as admin, it stops and restarts the `vmcompute` service. If `vmcompute` fails to restart within 15 seconds, it escalates: first restarting `vmms` (Virtual Machine Management), then in Deep mode only, restarting `HvHost` (which affects all Hyper-V VMs). This step is wrapped in a try/catch so failures don't block the rest of the fix process. Without admin, HCS errors are logged but require manual elevation.

**Step 9** monitors the VM boot log for definitive completion markers (`"Startup complete"`, `"[Keepalive]"`), showing real-time progress through the boot stages. Additionally monitors `cowork-service.log` for `isGuestConnected` RPC state -- if the guest reports connected, this accelerates readiness detection; if guest-timeout is detected after 90 seconds, performs targeted recovery (HCS cleanup + service restart) instead of waiting for the full 240-second timeout. Falls back to Hyper-V heartbeat checks and directory monitoring if logs are unavailable. After completion, the PowerShell window is brought to the foreground and the taskbar icon flashes until you dismiss it.

**Step 6 -- Smart Cache Purge:** Instead of blindly deleting everything, the script now backs up `sessiondata.vhdx` and `smol-bin.vhdx` before purging (with VHDX header integrity validation), then restores them after the service restart. If `smol-bin.vhdx` can't be backed up, it attempts recovery from the MSIX package. The step also cleans Claude temp files (`%TEMP%\anthropic-*`, `%TEMP%\claude-*`) and legacy `AnthropicClaude\sessions` and `vm-state` directories. Quick mode skips this step entirely.

**Smart mode escalation:** If the service doesn't start after the quick fix (Steps 1-5), Smart mode automatically escalates to a full deep purge with VHDX preservation: Phase 0 cleans stale HCS state, Phase 1 backs up session VHDXs, Phase 2 nukes the VM cache, the service is restarted, and Phase 3 restores the backed-up VHDXs. This ensures session data survives even when escalation is needed.

### What It Does NOT Touch

- `claude_desktop_config.json` -- your MCP servers and settings are safe
- `config.json` -- app configuration is safe
- Conversations -- stored server-side, not in the local VM cache

### Parameters

| Parameter | Description |
|-----------|-------------|
| `-Mode` | `Quick`, `Deep`, `Smart`, or `Diagnostic`. Skips the interactive menu. |
| `-BootPrep` | Non-destructive boot preparation mode. Unconditionally restarts `vmcompute` if no active Cowork workspace exists. Used by the logon boot-fix task. Does not kill Claude, stop services, or purge cache. |
| `-SkipLaunch` | Reset the VM but don't relaunch Claude |
| `-Quiet` / `-Silent` | Suppress the interactive menu and "press any key" prompt (for scheduled tasks). Defaults to Smart mode. |
| `-KeepCache` | Skip the VM cache purge (avoids ~2-3 GB re-download). Use when running Fix frequently. If the fix fails with `-KeepCache`, run again without it. |
| `-WhatIf` | Dry run -- show what would happen without changing anything |

### Diagnostics

Each run writes a timestamped log to `%APPDATA%\Claude\fix-logs\`. The health monitor logs to `%APPDATA%\Claude\watch-logs\`. Recent CoworkVMService errors from the Windows Event Log are shown in the summary.

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
| 12 | HCS service recovery | Auto-restart 30s/60s/120s | Configures vmcompute to auto-restart on failure with escalating delays |
| 13 | CoworkVMService recovery | Auto-restart 10s/30s/60s | Configures CoworkVMService to auto-restart on failure (independent of vmcompute) |
| 14 | HCS state cleanup | Pre-emptive close | Closes stale cowork-vm entries in HCS before they accumulate and block new VM creation |
| 15 | Service startup timeout | 120000ms | Prevents boot race conditions where services start before dependencies are ready |
| 16 | WinNAT rules | Verified / repaired | Ensures VM has outbound network connectivity |
| 17 | Firewall policies | Checked | Detects Group Policy blocking Hyper-V network rules |
| 18 | Storage location | Checked | Warns if workspace is on cloud-sync, USB, or network drive |
| 19 | Time synchronisation | Verified | Ensures NTP is running and clock drift is within tolerance |
| 20 | Antivirus exclusions | Configured / advised | Prevents AV filter drivers from blocking VirtioFS disk ops |
| 21 | WSL2 conflict detection | Checked | Warns about WSL2 distros and Docker Desktop that may conflict with Claude's VM |
| 22 | Health monitor | Every 30 seconds | Detects VirtioFS errors and auto-runs the full fix script |
| 23 | Boot-fix task | At logon (45s delay) | Runs the full fix script at every logon for a clean start |
| 24 | Shortcuts | Desktop + Start Menu | Quick access to Fix-ClaudeDesktop |
| 25 | Claude elevation | Scheduled task + Desktop shortcut | Ensures Claude Desktop launches with full admin privileges |
| 26 | Admin token policy | LocalAccountTokenFilterPolicy=1 | Disables remote/network admin token filtering |

Battery settings are not changed -- laptop users keep normal battery behaviour.

### Hyper-V VM Memory and Worker Priority

**Dynamic Memory pinning** (step 10) -- Hyper-V's Dynamic Memory feature allows Windows to balloon memory in and out of VMs based on demand. When memory is reclaimed from the Cowork VM, VirtioFS shared memory regions can become invalid -- this is the direct cause of the `EFAULT` ("bad address") error. Disabling Dynamic Memory pins the VM's allocation so it can't be reclaimed. This requires the VM to be stopped; if it's running when you run Prevent, a flag file is written and the health monitor applies the change the next time the VM restarts.

**VM worker process priority** (step 11) -- `vmwp.exe` is the Hyper-V Virtual Machine Worker Process that hosts each VM on the host side. At Normal priority, it can be starved under heavy CPU or I/O load, causing the VirtioFS connection to stall or time out. Setting it to AboveNormal gives it scheduling preference. This is not persistent across reboots, so the health monitor re-applies it on every poll cycle.

### HCS Service Recovery

**HCS service recovery** (step 12) -- The `vmcompute` service (Host Compute Service) manages all Hyper-V compute system operations. If it crashes, every VM creation call fails with `HCS operation failed`. The script configures Windows Service Control Manager to auto-restart `vmcompute` with escalating delays: 30 seconds after the first failure, 60 seconds after the second, 120 seconds after the third. The failure counter resets after 300 seconds of healthy operation. This is a permanent OS-level setting that survives reboots.

### CoworkVMService Recovery

**CoworkVMService recovery** (step 13) -- Independent of the `vmcompute` recovery in Step 12, the `CoworkVMService` itself can crash. The script configures Windows Service Control Manager to auto-restart it with escalating delays: 10 seconds, 30 seconds, 60 seconds. This ensures the Cowork VM restarts automatically without user intervention, even if `vmcompute` is healthy.

### HCS State Cleanup

**HCS state cleanup** (step 14) -- Over time, stale `cowork-vm` entries can accumulate in the Host Compute Service after unclean shutdowns. These orphan entries can block new VM creation when the service tries to create a new compute system. The prevention script proactively enumerates HCS entries via `hcsdiag list` and closes any stale cowork-vm instances. This is also performed by the health monitor on every cycle (Check 8) and by the fix script during cache purge.

**Service startup timeout** (step 15) -- The default `ServicesPipeTimeout` of 30 seconds can be too short on heavily loaded systems or during Windows Update reboots. If services like `vmcompute` don't start within this window, dependent services fail silently. Setting it to 120 seconds (120000ms) gives boot-time services more room. This is idempotent -- if the timeout is already >=120000ms (set by another tool), it's left untouched. Requires a reboot to take effect.

### Network and NAT

**WinNAT rules** (step 16) -- The Cowork VM needs a WinNAT rule to route traffic from its internal Hyper-V switch to the host's network. If this rule disappears (VPN reconnect, network adapter change, Windows Update), the VM silently loses all outbound connectivity. API calls fail, package downloads stall, and the workspace becomes unresponsive. The prevention script checks for existing NAT rules and auto-creates one if missing. The health monitor continuously monitors NAT health and repairs it automatically.

**Firewall policies** (step 17) -- Group Policy can set "Apply Local Firewall Rules" to disabled, which blocks the DHCP and DNS rules that Hyper-V's Host Network Service (HNS) creates for VMs. The script detects this and warns you to contact your IT admin. It also checks that Hyper-V-specific firewall rules are enabled.

### Storage, Time Sync, and Antivirus

**Storage location** (step 18) -- VirtioFS mounts fail when Claude's data directory is on a cloud-sync folder (OneDrive, Google Drive, Dropbox), an external USB drive, or a network share. The script detects these conditions and warns you to move Claude's data to a local SSD.

**Time synchronisation** (step 19) -- If the host clock drifts more than 5 seconds from NTP, Hyper-V's time synchronisation integration service can't correct the guest clock. This causes TLS certificate validation failures and API timeouts inside the VM. The script checks the W32Time service, measures actual drift, and forces a resync if needed. The health monitor continues to check every 5 minutes.

**Antivirus exclusions** (step 20) -- Antivirus filter drivers sit in the I/O path between VirtioFS and the host filesystem. They can delay or block disk operations that VirtioFS depends on, causing timeouts and mount failures. For Windows Defender, the script automatically adds exclusions for Claude's data directory, Hyper-V binaries, and the CoworkVMService process. For third-party AV products, it lists the recommended exclusion paths for you to add manually.

### The Health Monitor

A persistent background process that starts at logon and polls every 30 seconds. It monitors ten sources for VirtioFS failures:

1. **Claude log files** -- scans coworkd.log in `C:\ProgramData\Claude\Logs\` (with `%APPDATA%\Claude\logs\` fallback) for error patterns like "Plan9 mount failed" and "bad address"
2. **Service status** -- detects when `CoworkVMService` stops while `claude.exe` is still running (2 consecutive checks)
3. **Windows Event Log** -- checks for Claude-specific `CoworkVMService` errors and Hyper-V Worker/VMMS errors (2 consecutive checks, Claude-only matching)
4. **WinNAT health** -- detects missing NAT rules and auto-repairs them (every 60 seconds, warning only)
5. **Hyper-V heartbeat** -- monitors the VM's Integration Services heartbeat to detect hung VMs (3 consecutive checks)
6. **VM log staleness** -- catches silent hangs where the VM stops writing logs (5 consecutive stale checks, 5-minute threshold, only if VM was previously active)
7. **Clock drift** -- checks NTP drift every 5 minutes and auto-resyncs if >5 seconds (warning only)
8. **vmcompute health** -- monitors `vmcompute.exe` handle count every 60 seconds. Warning at 5000 handles, critical trigger at 10000 handles (2 consecutive checks required). Catches handle leaks that precede HCS failures
9. **HCS state health** -- monitors for stale cowork-vm entries in HCS via hcsdiag and tracks 0xC037010D shutdown failure frequency in the Hyper-V Compute event log. Proactively closes stale VMs (keeping the active one) and auto-restarts vmcompute if >15 shutdown failures occur within 5 minutes. Guarded by active session detection and Fix mutex check to prevent interference with running sessions
10. **Guest connection health** -- monitors `cowork-service.log` for `isGuestConnected` RPC polling patterns. Detects sustained polling with zero guest responses, indicating a guest connection timeout (3 consecutive checks required before trigger)

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
- **BootPrep mode** (v4.8.4) -- the logon task now uses non-destructive `-BootPrep` mode with a 45-second delay instead of the previous 180-second full fix. It unconditionally restarts `vmcompute` for a clean HCS state without killing Claude, stopping services, or purging cache. If Claude is already running with an active workspace, it exits silently
- **Startup grace period** -- all heuristic checks (including log scanning, event log, heartbeat, staleness) are skipped for the first 180 seconds after the monitor starts, preventing false triggers from pre-existing events
- **Consecutive-check gates** -- every heuristic trigger requires multiple consecutive failures before firing (service: 2, event log: 2, heartbeat: 3, staleness: 5). Only the log-file pattern check (actual VirtioFS error strings) triggers immediately
- **Tightened event log matching** -- Hyper-V VMMS events must mention "claude" or "cowork" (no generic "failed"/"unexpected" matching). Worker events: Critical/Error only
- **VM log staleness requires prior activity** -- only triggers if the VM log was previously active this session, preventing false positives in Chat mode
- **5-minute cooldown** -- between auto-fixes
- **30s pre-fix warning with notification** -- before any auto-fix, a Windows balloon notification with an audible chime warns you. You have 30 seconds to switch to Claude to cancel. If you don't, the fix proceeds. If you do switch to Claude, the fix cancels and a second notification tells you to run Fix-ClaudeDesktop.bat manually if Cowork is broken
- **Smart cancellation** -- the 30s grace period only cancels if Claude is *actively being used* (foreground window, CPU activity, VM log alive, or active Code session). General mouse/keyboard activity in other apps does **not** cancel the fix — so a genuinely hung VM still gets repaired while you're browsing or gaming
- **Default Switch NAT awareness** -- the NAT health check now recognises Hyper-V's "Default Switch" as providing NAT natively (via HNS), eliminating false "WinNAT missing" warnings on standard configurations
- **Invoke-HcsDiag timeout wrapper** (v5.0.0) -- all hcsdiag calls across Fix and Watch are wrapped with a 15-second Start-Job timeout to prevent indefinite hangs when HCS is corrupted
- **vmcompute restart session guard** (v5.0.0) -- the health monitor checks for active Claude sessions and Fix mutex before restarting vmcompute, preventing disruption of running Cowork sessions
- **ServiceTimeout raised** (v5.0.0) -- Fix script service stop timeout increased from 8s to 30s, preventing premature force-kills that cause HCS corruption

When a failure is detected **and the user is idle**, it shows a warning notification with a 30-second countdown, then runs `Fix-ClaudeDesktop.ps1 -Quiet`. If you switch to Claude during the countdown, the fix cancels and you're notified to run it manually. If the user was already detected as active before the countdown, it logs a `BLOCKED` message and waits.

The monitor also performs continuous maintenance: re-applying vmwp.exe AboveNormal priority on every cycle, and applying deferred Dynamic Memory changes when the VM is stopped.

The monitor uses a global mutex to ensure only one instance runs at a time. It logs its activity to `%APPDATA%\Claude\watch-logs\` (auto-cleaned after 30 days).

Visible in Task Scheduler under `\Claude\ClaudeCoworkWatchdog`. Can also be started manually with `Watch-ClaudeHealth.bat` for foreground monitoring.

### Claude Elevation and Admin Token Policy

**Claude elevation** (step 25) -- Claude Desktop is installed as an MSIX (Microsoft Store) package. By default, it launches with a standard (non-elevated) user token, even if you're an administrator. This means its child processes (including MCP servers like Desktop Commander) also run without admin privileges and cannot perform system-level operations. MSIX apps block all direct `.exe` access from `WindowsApps` (ACLs, `Start-Process -Verb RunAs`, `dir` enumeration all fail), so the only reliable approach is a **scheduled task**. The script creates a `\Claude\LaunchClaudeAdmin` task with `RunLevel=Highest` + `LogonType=Interactive`, which gives the process a full unfiltered admin token with no UAC prompt, while keeping the GUI visible in the user's desktop session. The task's action finds Claude at runtime via three methods: (1) `Get-AppxPackage` for MSIX installs, (2) common install paths for traditional `.exe` installs, (3) running-process detection as a final fallback. This survives version updates and works with any install type. A "Claude (Admin)" desktop shortcut triggers this task via `schtasks /run`. **Note:** MSIX installs will show a second Claude icon on the taskbar when launched elevated. This is unavoidable -- Windows enforces medium integrity for all shell-activated MSIX apps, so the only way to get a full admin token is to launch the `.exe` directly, which bypasses the MSIX app model's icon grouping. The scheduled task includes a process guard: if Claude is already running when the task is triggered (e.g., clicking the Desktop shortcut while Claude is open), it exits cleanly without launching a second instance. The Desktop shortcut uses Claude's actual icon (resolved at Prevent runtime via `Get-AppxPackage`); if the icon path becomes stale after a Claude update, it falls back to a generic Windows icon until Prevent is re-run.

**Admin token policy** (step 26) -- Windows filters admin tokens for local accounts during remote/network logins via `LocalAccountTokenFilterPolicy`. Setting it to `1` (along with `FilterAdministratorToken=0`) allows tools that use COM elevation, WMI, or remote PowerShell to receive full admin tokens. This is complementary to Step 25 -- the scheduled task handles the main elevation for Claude Desktop itself, while the token policy helps any tools that use COM-based or network-based elevation. UAC stays enabled and Store apps continue to work. Requires a reboot.

### Fast Startup, Connected Standby, NIC Power Saving

These three settings are the most commonly overlooked causes of VirtioFS failures:

**Fast Startup** (step 6) -- With Fast Startup on, a Windows "shutdown" is actually a kernel hibernate. Services like `CoworkVMService` don't fully reinitialise on the next boot, which can leave stale VM state behind. Disabling hibernate (`powercfg /h off`) usually kills Fast Startup too, but the script also explicitly sets the `HiberbootEnabled` registry key to 0 for safety.

**Connected Standby / Modern Standby** (step 7) -- On newer hardware (most laptops since ~2018, some desktops), Windows uses Modern Standby instead of traditional S3 sleep. Modern Standby can enter low-power states *even when sleep is "disabled"* in the power plan. Setting `CsEnabled = 0` in the registry forces traditional power behaviour. Requires a reboot.

**Network adapter power saving** (step 8) -- Windows can put physical network adapters to sleep to save power. This also affects the virtual network switch that Hyper-V uses for the VM's network bridge. The script disables power saving on all physical adapters via `PnPCapabilities`, wake-on-LAN settings, and WMI power management properties. Requires a reboot.

### The Boot-Fix Task

A scheduled task runs at every user logon as the current user (elevated). It waits 45 seconds (to let Windows services initialise), then executes `Fix-ClaudeDesktop.ps1 -BootPrep -Quiet` for a non-destructive vmcompute preparation.

**BootPrep mode** (added in v4.8.4) is a lightweight, non-destructive boot preparation mode that:

1. Waits up to 45 seconds for `vmcompute` to be running
2. Checks for an active Cowork workspace via `hcsdiag` — if Claude is already running with an active VM, it exits cleanly
3. Cleans any stale VMs left from a previous session (only if Claude is not running)
4. **Unconditionally restarts `vmcompute`** to ensure a fresh HCS state before Claude launches

Unlike the previous approach (v4.8.0–v4.8.3 used `-SkipLaunch -Quiet` with a 180-second delay), BootPrep:
- Runs in 45 seconds instead of 180, so the system is ready faster
- Never kills Claude processes or stops CoworkVMService
- Never purges cache or touches VM bundle files
- Is safe to run even if Claude is already open (it detects and exits)

This prevents the `0x800707DE` "failed to create compute system" HCS error that occurs when `vmcompute` has stale state after a reboot.

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

**"Request timed out: isGuestConnected"**
The VM booted but the guest OS did not connect back to the host. Run `Fix-ClaudeDesktop.bat` -- Step 9 now detects this specific failure and performs targeted recovery (HCS cleanup + service restart). If it recurs frequently, the Authenticode verification loop in `cowork-svc.exe` may be starving the service of CPU. See issue #31848. The health monitor (installed by Prevent) also detects this pattern automatically.

**Logs are piling up in `%APPDATA%\Claude\fix-logs\`**
The fix script auto-cleans logs older than 30 days. You can safely delete everything in that folder manually if needed.

---

## Stop-ClaudeDesktop

Clean shutdown without repair. Double-click `Stop-ClaudeDesktop.bat` when you want to fully close Claude Desktop without triggering a fix cycle.

### What It Does

1. Stops `CoworkVMService` gracefully (waits up to 45s for the VM to shut down)
2. Cleans any remaining HCS compute systems
3. Kills all Claude Desktop UI processes
4. Restarts the service in idle state (ready for next launch)
5. Verifies clean state

The service is restarted after stopping because Windows ignores ETW named pipe triggers for manually-stopped services. Without the restart, opening Claude Desktop would fail to auto-start `CoworkVMService`.

### When to Use

- Before a planned reboot
- When you want to fully close Claude without leaving VM processes behind
- When Claude is misbehaving and you want a clean stop before reopening

Unlike `Fix-ClaudeDesktop`, this does not purge cache, restart HCS services, or relaunch Claude.

---

## Known Limitations

### Claude Desktop closes and reopens on first launch after reboot

After a reboot, Claude Desktop may briefly close and reopen itself on the first launch. This is a **cosmetic behaviour caused by a race condition inside Claude Desktop** (not a ClaudeFix issue):

1. Claude Desktop restores multiple windows on startup
2. Each window independently tries to start the Cowork VM simultaneously
3. The first call succeeds; the second gets `"VM is already running"`
4. Claude's internal error handler triggers an auto-reinstall, causing the visible close-reopen
5. The VM starts successfully on the second attempt (~8–10 seconds total)

**This does not affect functionality.** No data is lost, no settings are changed, and the workspace starts correctly. ClaudeFix's BootPrep mode prevents the more serious `0x800707DE` construct failure — the close-reopen is a separate Claude Desktop bug that cannot be fixed externally.

Relevant log pattern in `C:\ProgramData\Claude\Logs\coworkd.log` (or legacy `%APPDATA%\Claude\logs\cowork_vm_node.log`):
```
[VM:start] Beginning startup, VM instance ID: ...
[VM:start] Beginning startup, VM instance ID: ...  (duplicate)
[error] VM boot failed: VM is already running
Auto-reinstalling workspace after startup failure
```

---

## File List

```
Fix-ClaudeDesktop.bat       -- Fix launcher (double-click when broken)
Fix-ClaudeDesktop.ps1       -- Fix script
Stop-ClaudeDesktop.bat      -- Clean shutdown launcher (double-click to stop Claude)
Prevent-ClaudeIssues.bat    -- Prevention launcher (run once)
Prevent-ClaudeIssues.ps1    -- Prevention script
Watch-ClaudeHealth.bat      -- Health monitor launcher (manual foreground mode)
Watch-ClaudeHealth.ps1      -- Health monitor (auto-detects and auto-fixes crashes)
README.md                   -- This file
LICENSE                     -- MIT licence
```

Current versions: Fix 5.3.1, Watch 5.0.0, Prevent 2.0.0

---

## Changelog

### v5.3.1 -- Close Mode Fix (2026-03-23)

- **Close mode rewrite** (Fix) -- stop-then-restart pattern. Stops CoworkVMService (triggers graceful VM shutdown ~31s), cleans HCS orphans, kills Claude UI, restarts service so it is ready for the next launch. Fixes the Windows behaviour where manually-stopped services ignore ETW named pipe triggers
- **Stop-ClaudeDesktop.bat** -- new launcher for clean shutdown via Close mode. Double-click to stop Claude without running a repair
- **Test-RecentHcsErrors Admin log fix** (Fix) -- Admin log code path now skips 0xC037010D events with continue instead of returning shutdown_stale, which was bypassing the >15 threshold on any single event
- **Version bump** -- Fix updated to v5.3.1

### v5.0.0 -- HCS Safety & Close Mode (2026-03-23)

- **Invoke-HcsDiag helper** (Fix + Watch) -- all hcsdiag calls wrapped with Start-Job + 15s timeout to prevent indefinite hangs when HCS is corrupted. 14 call sites in Fix, 3 in Watch
- **ServiceTimeout raised** (Fix) -- service stop timeout increased from 8s to 30s. Prevents premature force-kills that cause HCS state corruption
- **Graceful service polling** (Fix) -- WaitForStatus replaced with polling loop, preventing PowerShell timeout exceptions during slow VM shutdowns
- **VHDX handle verification** (Fix) -- 3 sites now verify VHDX files are not locked before backup/restore operations
- **0xC037010D threshold raised** (Fix) -- from 8 to 15 in Test-RecentHcsErrors. This event occurs on every normal VM shutdown (property query bug) and is not a blocker
- **vmcompute restart guards** (Watch) -- active session detection + Fix mutex check before vmcompute restart. Threshold raised from >5 to >15. Prevents disrupting running Cowork sessions
- **Log paths updated** (Fix + Watch) -- primary log path changed to C:\ProgramData\Claude\Logs\coworkd.log with fallback to %APPDATA%\Claude\logs\cowork_vm_node.log
- **BootPrep delay increased** (Prevent) -- boot-fix task delay raised from 30s to 45s with smart pre-check: skips if CoworkVMService or Claude is already running
- **Version bump** -- Fix 5.0.0, Watch 5.0.0, Prevent 2.0.0

### v4.9.0 -- Guest Connection Timeout Detection (2026-03-22)

- **New detection source** (Fix) -- reads `C:\ProgramData\Claude\Logs\cowork-service.log` for `isGuestConnected` RPC state via new `Test-CoworkServiceLog` helper
- **Guest timeout in HCS detection** (Fix) -- `Test-RecentHcsErrors` now returns `"guest_connect_failure"` when repeated `isGuestConnected` polls get no guest response
- **Targeted Step 9 recovery** (Fix) -- when guest-timeout is detected after 90s, performs HCS cleanup + service restart instead of waiting for the full 240s timeout
- **Guest connection health monitor** (Watch) -- new Check 10 monitors `cowork-service.log` for sustained `isGuestConnected` polling with zero guest responses (3 consecutive checks required before trigger)
- **Error patterns expanded** (Watch) -- added `isGuestConnected` timeout patterns to log-scan error list
- **README updated** with `isGuestConnected` error documentation, new tracked issues (#31848, #31314), troubleshooting entry, and updated step descriptions
- **Version bump** -- Fix and Watch updated to v4.9.0

### v4.8.6 -- Audit Cleanup (2026-03-22)
- **Dead code removal** (Fix) -- removed legacy $SkipLaunch vmcompute restart path from Step 3 (obsoleted by BootPrep in v4.8.4)
- **Stale .NOTES headers** (Watch, Prevent) -- corrected comment-block version numbers that were behind the runtime $Version
- **Issue tracker updates** (README) -- updated #26554 closure status (closed as "completed" with no linked fix), added #32172 (HCS 0x800707DE, still open)
- **Step 0 documented** (README) -- added Step 0 row to Fix "What It Does" table (pre-emptive HCS state cleanup + session file housekeeping)
- **Removed 24 obsolete files** -- deleted one-off diagnostic scripts and outdated CodePrompt files from Claude\ subfolder
- **Version bump** -- all three scripts updated to v4.8.6; Watch-ClaudeHealth mutex updated to v4.8.6

### v4.8.5 — Power Plan Fix (2026-03-09)

- **Fixed power plan activation** (Prevent) — `powercfg /duplicatescheme` returns a new GUID, but the code was trying to activate the template GUID which doesn't exist as an actual plan. Now parses the new GUID from the command output and activates that
- **Power plan deduplication** (Prevent) — Step 1 now detects existing "Ultimate Performance" plans before creating a new one. If duplicates exist from previous runs, all but one are deleted
- **Power plan cleanup on Undo** (Prevent) — `-Undo` now deletes all ClaudeFix-created Ultimate Performance plans in addition to restoring the original plan
- **Activation verification** (Prevent) — Step 1 now logs the actual active plan name and GUID after activation to confirm success

### v4.8.4 — Non-Destructive Boot Prep (2026-03-09)

- **BootPrep mode** (Fix) — new `-BootPrep` parameter for lightweight, non-destructive boot preparation. Unconditionally restarts `vmcompute` when no active workspace exists, preventing `0x800707DE` construct failures without killing Claude or stopping services
- **30-second boot task** (Prevent) — boot-fix scheduled task now uses `-BootPrep -Quiet` with a 30-second delay (was 180s with `-SkipLaunch -Quiet`), making the system ready for Claude 150 seconds faster
- **Stale VM cleanup at boot** (Fix) — BootPrep cleans orphan HCS compute systems from previous sessions before restarting `vmcompute`
- **Active workspace detection** (Fix) — BootPrep safely exits if Claude is already running with an active Cowork workspace, preventing interference with in-progress sessions
- **Version bump** — all three scripts updated to v4.8.4; Watch-ClaudeHealth mutex updated to `v4.8.4`

### v4.8.0 -- HCS Hardening & Robustness (2026-03-09)

- **CoworkVMService auto-recovery** (Prevent Step 13) -- configures `CoworkVMService` to auto-restart on failure with escalating delays (10s/30s/60s), independent of vmcompute recovery
- **Pre-emptive HCS state cleanup** (Prevent Step 14) -- closes stale cowork-vm entries in HCS during prevention setup and on every health monitor cycle
- **HCS state health monitor** (Watch Check 8) -- new check that detects stale cowork-vm entries, tracks 0xC037010D shutdown failure frequency, auto-restarts vmcompute on spikes (>5 in 5 min), and monitors session file accumulation
- **Close-StaleHcsVms helper** (Fix) -- unified function for consistent GUID parsing and HCS cleanup across all code paths (Step 0, Phase 0, Smart escalation, retry loop)
- **Test-HyperVReady heartbeat support** (Fix) -- VM readiness check now queries Hyper-V Integration Services heartbeat, enabling the wait loop to detect full VM boot (not just HCS registration)
- **Smart escalation VHDX preservation** (Fix) -- escalation from Quick to Deep now includes Phase 0 (HCS cleanup), Phase 1 (VHDX backup), Phase 2 (nuke), and Phase 3 (VHDX restore), preserving session data through the escalation
- **FileStream leak protection** (Fix + Watch) -- all FileStream usage in Test-VmLogReady and Test-LogsForErrors wrapped in try/finally to prevent handle leaks on exceptions
- **Event log deduplication** (Watch) -- Test-HcsStateHealth now fetches shutdown events once and filters in-memory for spike detection, eliminating redundant Get-WinEvent calls
- **Consistent log timestamps** (Fix) -- shared `$script:SessionTimestamp` ensures log file and transcript file use identical timestamps
- **Explicit auto-fix mode** (Watch) -- Invoke-AutoFix now passes `-Mode Smart -Quiet` explicitly instead of relying on defaults
- **Prevent step count fix** -- corrected `$steps = 27` to `$steps = 26` (was off-by-one)
- **Watch Check 8 safety** -- stale VM cleanup now skips the last (most likely active) cowork-vm entry, preventing accidental termination of the running VM

### v4.7.0 — Interactive Menu + 7 New Features (2026-03-08)

- **Interactive menu** -- Fix script now shows a mode selection menu when run manually (Quick, Deep, Smart, Diagnostic). Bypassed with `-Mode`, `-Quiet`, or non-interactive hosts. Falls back to simple numbered menu if `PromptForChoice` is unavailable
- **`-Mode` parameter** -- `Quick`, `Deep`, `Smart`, or `Diagnostic`. Skips the menu and runs in the specified mode
- **`-Silent` alias** -- `-Silent` is now an alias for `-Quiet` for backward compatibility
- **Diagnostic mode** -- Health-check-only mode that reports service status, HCS health, process counts, and VM cache state without making any changes
- **HvHost service restart fallback** -- Step 3 now escalates through vmcompute → vmms → HvHost when services fail to restart. HvHost restart is only attempted in Deep mode due to its impact on all Hyper-V VMs
- **vmwp.exe kill** -- Step 5 now detects and kills hung VM worker processes (`vmwp.exe`) via hcsdiag or force-kill, with VHDX corruption warnings
- **Smart cache purge with VHDX backup/restore** -- Step 6 backs up `sessiondata.vhdx` and `smol-bin.vhdx` (with VHDX header integrity validation) before purging, then restores them after service restart. Falls back to MSIX package recovery for `smol-bin.vhdx`
- **Temp file cleanup** -- Removes `%TEMP%\anthropic-*` and `%TEMP%\claude-*` files during cache purge
- **AnthropicClaude path cleanup** -- Cleans `AnthropicClaude\sessions` and `vm-state` directories from traditional installs
- **Smart mode escalation** -- If the service doesn't start after Quick-mode steps, Smart mode automatically escalates to a deep cache purge and retries
- **WSL2 conflict detection** (Prevent) -- New Step 19 checks for WSL feature, running WSL2 distros, and Docker Desktop. Warns about potential conflicts with Claude's Hyper-V VM

### v4.6.0 — Race Condition & Escalation Fixes (2026-03-08)

- **Startup grace period covers ALL watchdog checks** -- `Test-LogsForErrors` is now gated by `Test-StartupGracePeriod`, preventing stale log entries from triggering a fix before the monitor has settled
- **Grace period increased from 90s to 180s** -- gives Claude's VM more time to initialise before heuristic checks start
- **Persistent failure backoff** -- if the fix script runs 3+ times within 30 minutes and the problem persists, the watchdog stops retrying and shows a notification with Hyper-V nuclear reset instructions (DISM disable/enable)
- **Fix script mutual exclusion** -- a global mutex (`Global\ClaudeDesktopFix_v4.6`) prevents concurrent Fix runs from the watchdog, boot-fix task, and manual invocation
- **HCS JSON corruption detection** -- `Test-RecentHcsErrors` now detects Invalid JSON document patterns in HCS error logs (distinct from normal 0xC037010D shutdown events), and warns the user that a Hyper-V nuclear reset is needed
- **Watchdog delayed 120s at logon** -- was immediate; prevents the watchdog from racing Claude's own VM initialisation
- **Boot-fix delayed 180s at logon** -- was 90s; same race condition mitigation

---

## Credits & Acknowledgements

This toolkit was built by combining community knowledge from multiple independent contributors who shared their findings and workarounds:

- **Jonas Kamsker** ([blog.kamsker.at](https://blog.kamsker.at/blog/cowork-windows-broken/)) -- Comprehensive diagnostics, DNS/NAT fix scripts, and VM state recovery techniques
- **Elliot Segler** ([elliotsegler.com](https://www.elliotsegler.com/fixing-claude-coworks-network-conflict-on-windows.html)) -- Network conflict resolution and HNS-based recovery for subnet collisions
- **@garabedjunior-dotcom** ([GitHub #29848](https://github.com/anthropics/claude-code/issues/29848)) -- Community troubleshooting scripts and MCP crash diagnosis
- **@Onimir89** ([GitHub: Restart_claude](https://github.com/Onimir89/Restart_claude)) — Independent VM restart script demonstrating the service-reset recovery pattern
- Everyone who reported and documented VirtioFS failures in the [claude-code issue tracker](https://github.com/anthropics/claude-code/issues)

Special thanks to the community contributors on GitHub issues #25206, #26554, #27576, #27801, #28890, #29045, #29587, #29848, and #31520 whose collective debugging narrowed down the root causes.

---

## Licence

MIT

---

*Out of frustration, built with [Claude Desktop](https://claude.ai), for [Anthropic](https://github.com/anthropics).*
