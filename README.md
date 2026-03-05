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
| 6 | Watchdog task | Every 5 minutes | Auto-restarts `CoworkVMService` if it dies while Claude is running |
| 7 | Boot-fix task | At logon | Runs the full fix script at every logon for a clean start |
| 8 | Shortcuts | Desktop + Start Menu | Quick access to Fix-ClaudeDesktop |

Battery settings are not changed -- laptop users keep normal battery behaviour.

### The Watchdog

A scheduled task runs every 5 minutes as SYSTEM. It checks whether `CoworkVMService` has stopped while `claude.exe` is still running, and restarts the service if so.

This catches cases where the service simply stops. Stale VirtioFS mounts still need the full fix script.

Visible in Task Scheduler under `\Claude\ClaudeCoworkWatchdog`.

### The Boot-Fix Task

A scheduled task runs at every user logon as SYSTEM. It executes `Fix-ClaudeDesktop.ps1 -SkipLaunch -Quiet` to ensure the VM service starts cleanly before Claude tries to use it.

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

## File List

```
Fix-ClaudeDesktop.bat       -- Fix launcher (double-click when broken)
Fix-ClaudeDesktop.ps1       -- Fix script
Prevent-ClaudeIssues.bat    -- Prevention launcher (run once)
Prevent-ClaudeIssues.ps1    -- Prevention script
README.md                   -- This file
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
