# Claude Desktop / Cowork VM Fix for Windows

**Fix the "VirtioFS mount failed: bad address" crash in Claude Desktop on Windows -- without rebooting.**

If you're seeing any of these errors in Claude Desktop's Cowork mode, this toolkit will fix them:

```
RPC error -1: failed to ensure virtiofs mount: Plan9 mount failed: bad address
```
```
Workspace failed to start
```
```
Setting up workspace... (stuck forever)
```

Two scripts: one **prevents** the crash, the other **fixes** it when it happens anyway.

| Script | Purpose | Run when |
|--------|---------|----------|
| `Prevent-ClaudeIssues.bat` | Configure Windows to minimise crashes | Once |
| `Fix-ClaudeDesktop.bat` | Reset and relaunch when it breaks | Every time it breaks |

---

## Quick Start

1. Download all files into one folder (e.g. `C:\ClaudeFix\`)
2. Run **`Prevent-ClaudeIssues.bat`** once (configures Windows, creates watchdog, boot-fix task, and shortcuts)
3. Use the Desktop shortcut or search "Fix Claude Desktop" in Start when Cowork breaks
4. Right-click the Start Menu entry and select **Pin to taskbar** for one-click access

After running the prevention script, Claude will be automatically repaired at every logon and the VM service will be monitored continuously.

---

## Compatibility

| | Supported |
|---|---|
| Windows 10 (1607+) | Yes |
| Windows 11 | Yes |
| Admin privileges | Optional (recommended) |
| MSIX / Microsoft Store install | Yes |
| Traditional (.exe) install | Yes |
| PowerShell 5.1 | Yes (ships with Windows) |
| PowerShell 7+ | Yes |

**Admin privileges are requested but not required.** If you decline UAC, the fix script still works -- it just can't control the Windows service directly. Claude handles that automatically when it launches.

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

### Tracked Issues

- [#26554](https://github.com/anthropics/claude-code/issues/26554) -- VirtioFS mount fails with "bad address"
- [#27576](https://github.com/anthropics/claude-code/issues/27576) -- Mount failure after ~1 hour of use
- [#28890](https://github.com/anthropics/claude-code/issues/28890) -- Mount goes stale after idle
- [#29587](https://github.com/anthropics/claude-code/issues/29587) -- Cowork fails after brief use
- [#29848](https://github.com/anthropics/claude-code/issues/29848) -- Recurring VM crashes

---

## Fix-ClaudeDesktop

One-click fix when Claude Desktop / Cowork is broken. No reboot needed.

### What It Does

| Step | Action |
|------|--------|
| 1 | Captures Claude.exe path, then force-kills all `claude.exe` processes |
| 2 | Stops `CoworkVMService` (graceful with admin, force-kill without) |
| 3 | Verifies no orphan processes remain |
| 4 | Deletes stale VM cache (`claude-code-vm` and `vm_bundles`) to force a clean rebuild |
| 5 | Restarts `CoworkVMService` (admin) or defers to Claude auto-restart (non-admin) |
| 6 | Auto-detects and relaunches Claude Desktop (MSIX-aware -- no duplicate taskbar icons) |
| 7 | Monitors `cowork_vm_node.log` for boot completion, confirms workspace is ready |

**Step 6** detects both MSIX (Microsoft Store) and traditional installs. MSIX apps are launched via `shell:AppsFolder` protocol to avoid creating loose instances with duplicate taskbar icons. Eight different detection methods ensure the executable is found regardless of install location.

**Step 7** monitors the VM boot log for definitive completion markers (`"Startup complete"`, `"[Keepalive]"`), showing real-time progress through the boot stages. Falls back to Hyper-V heartbeat checks and directory monitoring if logs are unavailable. After completion, the PowerShell window is brought to the foreground and the taskbar icon flashes until you dismiss it.

### What It Does NOT Touch

- `claude_desktop_config.json` -- your MCP servers and settings are safe
- `config.json` -- app configuration is safe
- Conversations -- stored server-side, not in the local VM cache

### Parameters

| Parameter | Description |
|-----------|-------------|
| `-SkipLaunch` | Reset the VM but don't relaunch Claude |
| `-Quiet` | Suppress the "press any key" prompt (for scheduled tasks) |
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
| 10 | Watchdog task | Every 5 minutes | Auto-restarts `CoworkVMService` if it dies while Claude is running |
| 11 | Boot-fix task | At logon | Runs the full fix script at every logon for a clean start |
| 12 | Shortcuts | Desktop + Start Menu | Quick access to Fix-ClaudeDesktop |

Battery settings are not changed -- laptop users keep normal battery behaviour.

### The Watchdog

A scheduled task runs every 5 minutes as SYSTEM. It checks whether `CoworkVMService` has stopped while `claude.exe` is still running, and restarts the service if so.

This catches cases where the service simply stops. Stale VirtioFS mounts still need the full fix script.

Visible in Task Scheduler under `\Claude\ClaudeCoworkWatchdog`.

### Fast Startup, Connected Standby, NIC Power Saving

These three settings are the most commonly overlooked causes of VirtioFS failures:

**Fast Startup** (step 6) -- With Fast Startup on, a Windows "shutdown" is actually a kernel hibernate. Services like `CoworkVMService` don't fully reinitialise on the next boot, which can leave stale VM state behind. Disabling hibernate (`powercfg /h off`) usually kills Fast Startup too, but the script also explicitly sets the `HiberbootEnabled` registry key to 0 for safety.

**Connected Standby / Modern Standby** (step 7) -- On newer hardware (most laptops since ~2018, some desktops), Windows uses Modern Standby instead of traditional S3 sleep. Modern Standby can enter low-power states *even when sleep is "disabled"* in the power plan. Setting `CsEnabled = 0` in the registry forces traditional power behaviour. Requires a reboot.

**Network adapter power saving** (step 8) -- Windows can put physical network adapters to sleep to save power. This also affects the virtual network switch that Hyper-V uses for the VM's network bridge. The script disables power saving on all physical adapters via `PnPCapabilities`, wake-on-LAN settings, and WMI power management properties. Requires a reboot.

### The Boot-Fix Task

A scheduled task runs at every user logon as the current user (elevated). It executes `Fix-ClaudeDesktop.ps1 -SkipLaunch -Quiet` to ensure the VM service starts cleanly before Claude tries to use it. It runs as your user account (not SYSTEM) so that it can find Claude's data in the correct `%APPDATA%` location.

The prevention script auto-detects `Fix-ClaudeDesktop.ps1` in the same folder. If it can't find it, this step is skipped with a warning.

Visible in Task Scheduler under `\Claude\ClaudeCoworkBootFix`.

### Shortcuts

Creates a "Fix Claude Desktop" shortcut on the Desktop and in the Start Menu. You can pin the Start Menu entry to your taskbar for one-click access.

### Undo Everything

```
.\Prevent-ClaudeIssues.ps1 -Undo
```

Restores your original power plan, re-enables hibernate, resets sleep to 30 minutes, removes both scheduled tasks, and deletes the shortcuts.

---

## Requirements

- Windows 10 (1607+) or Windows 11
- Claude Desktop with Cowork mode
- PowerShell 5.1+ (included with Windows)
- Admin privileges: optional but recommended

---

## Troubleshooting

**The fix script says "Workspace ready" but Cowork still shows an error**
Try running the fix script a second time. Some stale states need two cycles to fully clear. If it persists after 2-3 runs, reboot and let the boot-fix task handle it.

**"Service not found -- is Cowork installed?"**
The `CoworkVMService` Windows service is only installed when you've used Cowork mode at least once. Open Claude Desktop, start a Cowork session, let it install the VM components, then run the scripts.

**The prevention script asks for admin but I don't have it**
The prevention script genuinely needs admin for power settings, scheduled tasks, and registry changes. The fix script can run without admin, but the prevention script cannot. Ask your IT department for temporary elevation, or have them run `Prevent-ClaudeIssues.bat` for you.

**Connected Standby changes didn't take effect**
This setting requires a full reboot (not just sleep/wake). Shut down, wait 10 seconds, power on. Verify with `powercfg /a` -- if it no longer lists "Standby (S0 Low Power Idle)" then Modern Standby is disabled.

**Claude launches with a duplicate taskbar icon**
This happens when an MSIX (Microsoft Store) app is launched via its `.exe` path directly instead of through the shell protocol. The fix script handles this automatically, but if you see it, make sure you're using the latest version of the fix script.

**Logs are piling up in `%APPDATA%\Claude\fix-logs\`**
The fix script auto-cleans logs older than 30 days. You can safely delete everything in that folder manually if needed.

---

## File List

```
Fix-ClaudeDesktop.bat       -- Fix launcher (double-click when broken)
Fix-ClaudeDesktop.ps1       -- Fix script
Prevent-ClaudeIssues.bat    -- Prevention launcher (run once)
Prevent-ClaudeIssues.ps1    -- Prevention script
README.md                   -- This file
LICENSE                     -- MIT licence
```

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
