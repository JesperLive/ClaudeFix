# Claude Desktop / Cowork VM Fix for Windows

> **Maintenance status (2026-07-29):** Actively maintained again. The 6.0.x line
> landed after a full audit of all three scripts, and 6.0.1 through 6.0.4 came
> from real user bug reports over two days. Report breakage and it gets looked
> at. If nothing lands within 7 days of a report, the Troubleshooting section
> below has manual recovery steps that do not depend on this toolkit.

**Fix "VirtioFS mount failed", "HCS operation failed", and "Failed to start Claude's workspace" errors in Claude Desktop Cowork mode on Windows, without rebooting.**

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
2. Run **`Prevent-ClaudeIssues.bat`** once (configures Windows, installs the health monitor, and creates the shortcuts and launchers)
3. Use the Desktop shortcut or search "Fix Claude Desktop" in Start when Cowork breaks
4. Right-click the Start Menu entry and select **Pin to taskbar** for one-click access

After running the prevention script, a background health monitor checks for VirtioFS and HCS failures every 30 seconds and runs the fix when you are idle.

---

## Compatibility

| | Supported |
|---|---|
| Windows 10 Pro/Enterprise/Education (build 19041+) | Yes |
| Windows 11 Pro/Enterprise/Education | Yes |
| Windows 10/11 Home | Mostly, see the note below |
| Windows 7 / 8 / 8.1 | No (Claude Desktop requires Windows 10 19041+) |
| MSIX / Microsoft Store install | Yes |
| Traditional (.exe) install | Yes |
| PowerShell 5.1 | Yes (ships with Windows) |
| PowerShell 7+ | Yes |

**About Home editions.** Earlier versions of this table said Home was
unsupported, on the reasoning that Cowork needs Hyper-V and Home does not have
it. That was wrong on both halves.

What Cowork needs is
[Virtual Machine Platform](https://support.claude.com/en/articles/12622703-deploy-claude-desktop-for-windows),
which is a separate Windows feature from the full Hyper-V role and *is*
available on Home. Anthropic's own deployment page names that feature and no
other. Reports from Home users bear it out: the machine in
[#80444](https://github.com/anthropics/claude-code/issues/80444) runs Windows 11
Home, and so does the one in
[#61372](https://github.com/anthropics/claude-code/issues/61372).

Two things are genuinely narrower on Home. `New-NetNat` and the NetNat WMI
components are missing or unreliable there, so the manual NAT recovery
documented by
[Jonas Kamsker](https://blog.kamsker.at/blog/cowork-windows-broken/) will not
run. Prevent's own WinNAT check, which also depended on them, was retired in
6.3.0. The Hyper-V PowerShell module is also absent, though nothing in this
toolkit depends on it any more, because the Cowork VM is an HCS compute system
that `Get-VM` never returns on any edition.

Everything else in this toolkit works on Home.

| Script | Admin required? |
|--------|----------------|
| `Fix-ClaudeDesktop` | No (recommended, but works without it, using force-kill instead of service control) |
| `Prevent-ClaudeIssues` | **Yes** (power settings, registry, scheduled tasks all require admin) |
| `Watch-ClaudeHealth` | No (auto-installed as elevated task by Prevent) |

---

## The Problem

Claude Desktop runs a lightweight VM, an HCS compute system rather than a Hyper-V VM, to power Cowork mode. The VM uses VirtioFS (Plan9 protocol) to share the filesystem between host and guest. This mount frequently breaks:

```
RPC error -1: failed to ensure virtiofs mount: Plan9 mount failed: bad address
```

Closing and reopening Claude doesn't help; the `CoworkVMService` stays in a broken state. The typical "fix" before these scripts was rebooting the entire PC.

### Why It Happens

The VirtioFS connection degrades when:

- Windows enters sleep or hibernate (the VM doesn't survive the transition)
- The VM sits idle for extended periods
- Power management reduces PCI-E or USB link states mid-operation
- `CoworkVMService` crashes and doesn't auto-recover
- Stale VM cache files from a previous session conflict with the new boot
- Antivirus filter drivers interfere with VirtioFS disk operations
- Host clock drift causes TLS and API failures inside the VM

Earlier versions of this list also named Hyper-V Dynamic Memory and
disappearing WinNAT rules. Both were dropped in 6.3.0: the Cowork VM is not a
Hyper-V VM and has no Dynamic Memory setting, and it does not use a WinNAT
rule.

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

A third class of failure occurs after the VM boots but the guest operating system never establishes its connection back to the host. Claude Desktop polls `cowork-svc.exe` via a named pipe (`\\.\pipe\cowork-vm-service`) every ~1 second with `isGuestConnected` RPC calls. If the guest never reports as connected, the UI shows `"Request timed out: isGuestConnected"`. This can be caused by: the VM booting but vsock/networking failing to initialise, `cowork-svc.exe` being CPU-starved by its own Authenticode signature verification loop (each poll re-hashes the 204 MB `claude.exe` binary, consuming ~960ms per call), or HCS state corruption preventing guest-host communication. The service log `cowork-service.log` in the app's log folder (see [Where the app keeps its files](#where-the-app-keeps-its-files)) records each RPC call. See issue #31848 for the Authenticode CPU burn.

Note that on current builds this file has not been written since March. The toolkit checks its age before parsing it, so a stale copy is ignored rather than mined for answers about the last 60 seconds. The health monitor's guest-connection check, which looked in this file for a message it has never held, was retired in 6.3.0; a `Request timed out: isGuestConnected` the app logs at `[error]` in its VM or main log still starts a repair through the monitor's log check.

### Tracked Issues

State checked against the tracker on 2026-07-29. Several of these were described
here as "still open, no Anthropic response" long after they had been closed, so
the state column is now something to re-verify rather than assume.

A closed issue does not mean the failure cannot recur. Most of these closed
without a linked PR or any published fix detail, and the underlying
architecture is unchanged, which is why the toolkit still handles them.

**Open**

- [#27801](https://github.com/anthropics/claude-code/issues/27801): "Failed to start Claude's workspace", VM service not running, persists after reboot
- [#29045](https://github.com/anthropics/claude-code/issues/29045): Claude Desktop spawns a Hyper-V VM on every launch, even for chat-only use
- [#80444](https://github.com/anthropics/claude-code/issues/80444): fatal GPU-process crash via the in-app Browser tab, and the MSIX package left needing Repair afterwards. Prevent's step 18 writes a launcher that works around the crash half of this

**Closed**

- [#25206](https://github.com/anthropics/claude-code/issues/25206): VM starts then crashes within 5 minutes, service unrecoverable
- [#26554](https://github.com/anthropics/claude-code/issues/26554): VirtioFS mount fails with "bad address"
- [#27576](https://github.com/anthropics/claude-code/issues/27576): mount failure after roughly an hour of use
- [#28890](https://github.com/anthropics/claude-code/issues/28890): mount goes stale after idle
- [#29587](https://github.com/anthropics/claude-code/issues/29587): Cowork fails after brief use
- [#29848](https://github.com/anthropics/claude-code/issues/29848): recurring VM crashes
- [#31314](https://github.com/anthropics/claude-code/issues/31314): cowork-svc.exe sustaining 195 MB/s of I/O from the signature polling loop
- [#31520](https://github.com/anthropics/claude-code/issues/31520): community recovery script for VirtioFS failures, which this toolkit covers and extends
- [#31703](https://github.com/anthropics/claude-code/issues/31703): HCS and VM service failures on v1.1.5368
- [#31848](https://github.com/anthropics/claude-code/issues/31848): cowork-svc.exe CPU burn from Authenticode re-verification on every isGuestConnected poll
- [#32172](https://github.com/anthropics/claude-code/issues/32172): HCS 0x800707DE construct failure after a VirtioFS mount error

---

## Fix-ClaudeDesktop

One-click fix when Claude Desktop / Cowork is broken. No reboot needed.

### Interactive Menu

When run manually (without `-Quiet` or `-Mode`), the script shows an interactive menu:

1. **Quick Fix**: Restart services + basic repair (every step but the cache purge and the HvHost restart)
2. **Deep Fix**: Full nuclear reset (all steps including cache purge)
3. **Smart Fix**: Try quick first, escalate to deep if needed (default)
4. **Diagnostic**: Health check only, no changes

After selecting a mode, a second menu lets you toggle options (Keep cache, Skip relaunch, WhatIf). A summary is shown before the final Y/N confirmation.

The menu is skipped when:
- `-Mode` is passed as a parameter (uses that mode directly)
- `-Quiet` is set (defaults to Smart mode, for automated callers)
- The host is non-interactive (defaults to Smart mode)

### What It Does

| Step | Action |
|------|--------|
| 0 | Pre-emptive HCS state cleanup (stale cowork-vm entries); with `-PurgeSessions`, also deletes session files older than 7 days |
| 1 | Captures Claude.exe path, then force-kills the Claude Desktop app's processes. Claude Code CLI processes, also named `claude.exe`, are left running (since 6.2.0) |
| 2 | Stops CoworkVMService (graceful with admin, force-kill without) |
| 3 | Checks for HCS errors and restarts vmcompute service if needed; escalates to vmms, and in Deep mode only to HvHost |
| 4 | Verifies no orphan app or service processes remain |
| 5 | Kills orphan HCS compute systems via hcsdiag, and a hung Cowork VM worker (`vmwp.exe`) |
| 6 | Smart cache purge: backs up session VHDXs, nukes VM cache, restores session data; cleans temp files and legacy paths |
| 7 | Restarts CoworkVMService (admin) or defers to Claude auto-restart (non-admin); restores VHDX backups |
| 8 | Relaunches Claude Desktop with elevated privileges via scheduled task (Method 0), falling back to MSIX shell protocol or direct exe launch |
| 9 | Watches every VM log for boot completion and hcsdiag for instance state, confirms workspace is ready |

**Step 8** first attempts to relaunch Claude with elevated privileges via the `LaunchClaudeAdmin` scheduled task created by Prevent (Method 0). This gives Claude a full admin token without a UAC prompt. If the task doesn't exist or fails, it falls through to three standard methods: Method A launches MSIX installs via `shell:AppsFolder` protocol (no duplicate taskbar icons), Method B launches traditional `.exe` installs directly, and Method C uses Start Menu shortcuts as a last resort. When whether the task exists cannot be read, the step says so and goes on to the standard methods; until 6.3.1 that read as a missing task, with Prevent to run for it. When no Claude shortcut is found and part of the Start Menu could not be searched, that is said too, since the shortcut may be there. All methods respect `-WhatIf` and each has a `$launched` guard to prevent double-launch.

**Step 5** terminates orphan HCS compute systems that survive service shutdown. When CoworkVMService stops, the underlying VM may remain registered in HCS, causing "VM is already running" errors on restart. The script uses `hcsdiag list` to find cowork compute systems and `hcsdiag kill` to terminate them (admin only), and counts a kill only once the compute system is gone from a second list. It then looks for a hung Cowork VM worker: the `vmwp.exe` whose first argument is the GUID of a Cowork compute system that any `hcsdiag list` showed during the run, so the worker of a compute system killed earlier in the run is still found. It holds that process before it acts, so a process that has since taken over its PID is left alone. It stops the worker through hcsdiag, waits up to 5 seconds for it to exit, and force-kills it only if it is still running then. A worker counts as stopped only once it has exited, and one it could not stop is reported. When the VM worker processes cannot be listed, no worker is stopped, and the step says the worker check did not finish. A clean-up of compute systems, here and in the other steps that run one, says when a kill could not be sent and when it stopped on an error, so that what it removed is not known. When nothing was stopped, the step ends with "No orphan compute systems found" only when hcsdiag answered and the worker check finished; otherwise it says "Nothing was stopped, but" and what was not checked. Compute systems that were listed and not killed, as under `-WhatIf`, are reported as listed and not killed. When no list showed a Cowork compute system, as when the script runs without admin, no worker is touched. Until 6.3.0 this step killed every `vmwp.exe` on the machine, so a Hyper-V VM, WSL2 or Windows Sandbox running beside Claude went down with it, and it counted every kill it sent whether hcsdiag had removed anything or not. This step is non-fatal, failures don't block the rest of the fix.

There is no Hyper-V cmdlet fallback, and earlier versions of this section were wrong to describe one. `Stop-VM` operates at the Hyper-V management layer, and the Cowork VM is an HCS compute system that never appears to `Get-VM` at all, so that path could never have done anything. hcsdiag is the only tool that sees this VM.

If hcsdiag is missing or does not answer, the step now says so rather than reporting a clean result. From 6.3.0 an hcsdiag call that exits with an error code counts as not answering, and an empty list counts as nothing listed. Both outcomes previously returned "no orphan compute systems found", which meant a run whose first step reported hcsdiag unavailable would still claim a successful hcsdiag check five steps later.

**Step 3** checks recent Windows Event Log entries and Claude logs for HCS error patterns (`HCS operation failed`, `failed to create compute system`, `HcsWaitForOperationResult`). If detected and running as admin, it stops and restarts the `vmcompute` service. If `vmcompute` fails to restart within 15 seconds, it escalates: first restarting `vmms` (Virtual Machine Management), then in Deep mode only, restarting `HvHost` (which affects all Hyper-V VMs). It escalates only on a `vmcompute` known to be down: when whether `vmcompute` came back cannot be read, it escalates no further, and a `vmms` or `HvHost` it cannot read is not restarted. A `vmms` or `HvHost` restart that reported an error is said, not called a restart. Until 6.3.1 a `vmcompute` that could not be read escalated to both, and each restart was said whatever happened. This step is wrapped in a try/catch so failures don't block the rest of the fix process. When the Hyper-V Compute event logs, a Claude log or the Claude log folder cannot be read, it says so, naming the error, and goes on to the next Claude log; the two Hyper-V Compute logs are read together, so a failure on one leaves both unchecked, and the line says so. Until 6.3.0 each read as one with no HCS error, and until 6.3.1 so did a Claude log folder that could not be read at all. Without admin, HCS errors are logged but require manual elevation.

**Step 9** monitors the VM boot log for definitive completion markers (`"Startup complete"`, `"[Keepalive]"`), showing real-time progress through the boot stages. The log is chosen by write time at startup rather than from a fixed preference list, so a build that names its logs differently still gets read and a stale file gets reported rather than believed. It also watches `cowork-service.log` for `isGuestConnected` state where that file is fresh enough to mean anything.

The wait is not a fixed countdown. It starts at 240 seconds, 300 after a cache purge, and pushes the deadline out again every time the VM cache grows, with a hard ceiling of one hour. A fresh bundle download is slow rather than stuck, and the difference between the two is whether bytes are still arriving. No growth for one progress window is what stuck actually looks like.

There is no Hyper-V heartbeat fallback. Earlier versions of this section listed one; there is no heartbeat to read, because `Get-VM` does not return this VM at all.

If the prerequisite check flagged something and the service also failed to start, the wait is skipped rather than spent repeating an error, and a restart during the wait that fails on such a cause stops the wait. Either way Smart mode's escalation to Deep and the retry cycle after it are skipped too, since neither a purge nor a restart can fix that: the run says so, without telling you to open a Cowork session, and exits 1. Until 6.3.1 both ran, and the escalation purged the VM cache of a service that was disabled. If the service is running, the wait proceeds regardless of what the check thought. A service whose state cannot be read does not skip the wait: the wait says so once and does not restart the service on it. Until 6.3.1 the wait was skipped on it, and during the wait it read as a service that had died and was restarted.

A VM log that cannot be read is said once, naming the log (or its folder, when the error gives no path) and the error, and from then on the logs cannot say whether the workspace is ready: a guest that says it is connected is not taken as ready on a VM log that could not be read. A `cowork-service.log` that cannot be read is said once per run. When the workspace was not confirmed ready and a VM log could not be read, the run neither escalates nor retries on it: it ends with `Whether the workspace came up is not known: a VM log could not be read`, tells you to open a Cowork session and run Fix again if it does not start, and exits 1, unless the service cannot start (above), when the run ends with that cause instead. Until 6.3.1 a VM log that could not be read was passed over as one with nothing new in it, and the no-progress escalation and the retries ran on it. After a guest-connection timeout, the targeted recovery says `Service restarted, continuing to wait...` only when its start of the service went out, and `Continuing to wait...` otherwise.

After completion, the PowerShell window is brought to the foreground and the taskbar icon flashes until you dismiss it.

**Step 6, Smart Cache Purge:** Instead of blindly deleting everything, the script now backs up `sessiondata.vhdx` and `smol-bin.vhdx` before purging (with VHDX header integrity validation), then restores them in Step 7, before the service is started. If `sessiondata.vhdx` can't be backed up, the cache is not purged. The step also cleans Claude temp files (`%TEMP%\anthropic-*`, `%TEMP%\claude-*`) and legacy `AnthropicClaude\sessions` and `vm-state` directories. Quick and Smart runs skip the cache purge and keep the cache, as `-KeepCache` does; the temp-file and legacy-folder cleanup still runs. Smart mode purges only when it escalates (below), and never with `-KeepCache`. A cache folder that cannot be read is reported as not purged, here and in the escalation purges. Until 6.3.1 Step 6 reported it as not present, and the escalation purges passed over it without a word. A delete that failed in part says the folder could not be removed in full, where it said removed. The VHDX backup does not purge when it cannot read whether a VHDX is in the VM cache, and when the free space on the backup volume is unknown it says why and does not purge. When Step 6 keeps the cache, a cache folder that cannot be read is said, and kept. When the app's data folder is not known, a Step 6 that purges, as in Deep mode, says the cache folders are not known, and one that keeps the cache passes over them, as the escalation purges do; until 6.3.1 Step 6, the escalation purges and the Diagnostic run stopped there with UNEXPECTED ERROR. Restoring the VHDX backups does not write a backup over a file it could not read, and a backup that could not be read is said, and not restored.

**Smart mode escalation:** Smart mode first runs as Quick mode does: every step, with the cache purge in Step 6 skipped and no HvHost restart in Step 3, which only Deep mode runs. If the service doesn't start in Step 7, or the workspace is not ready after Step 9, Smart mode escalates once, as administrator, to a full deep purge with VHDX preservation: it cleans stale HCS state, backs up the session VHDXs, nukes the VM cache, restores the backed-up VHDXs, and then starts the service. The escalation after Step 9 also stops the Claude app first and relaunches it, and is not run with `-SkipLaunch`. Neither escalation runs with `-KeepCache` (Keep cache [ON] in the menu), since each is a purge of the cache it keeps: the run says it did not escalate, and why; until 6.3.1 both purged it. Neither runs when the prerequisite check found that the service cannot start at all, as when it is disabled, and it did not start, which a purge cannot fix; until 6.3.1 the one after Step 9 still ran then. This ensures session data survives even when escalation is needed. When the service's state cannot be read, Smart mode does not escalate to Deep, and says so.

**A read that fails is said** (v6.3.1). Every read that fails and that Fix would act on is now said, naming the reason, mostly in a dark yellow `[!] ... could not be read` or `could not be listed` line. Fix does not do what an empty answer would have led it to: it stops, kills, restarts, purges and deletes nothing on it. A few reads that decide nothing stay silent, such as the cache size that only extends the Step 9 wait, and the taskbar flash at the end, which no longer prints a red error when it cannot be loaded. The reads discovery could not make are logged at the top of the run, and old logs, old VHDX backups and a discovery record that could not be cleaned up or saved are said once logging starts. Finding `Claude.exe` passes over a cached path, the App Paths key or a Start Menu folder it cannot read, and says so; the cached path is kept. A folder at one of those paths is not taken for `Claude.exe`: until 6.3.1 Step 8 started it, which opened the folder in Explorer. Step 0 leaves a stale cowork-vm for Step 5, and says so, when whether Claude is running cannot be read. Steps 1 and 4 stop none of Claude's processes, or cowork-svc's, when they cannot be listed, and Step 1 counts only the processes that stopped, followed by `[!] N could not be stopped (reason)`. Step 2, Step 9 and the service restarts do not stop, force-kill or restart a service whose state cannot be read, and a stop the service refused says `Asking ... to stop failed (reason)`, where it said nothing. When whether the service process has exited cannot be read, Fix says so, with the warning that the VHDX files may be locked. `-PurgeSessions` purges nothing when the sessions folder cannot be read, and says so; it says when part of the folder cannot be read and does not consider the files there, reports `Deleted N of M` with `[!] K could not be deleted`, and keeps the session folders it could not read in full. Under `-WhatIf` the service process is neither force-killed nor started, and neither is reported as done. Until 6.3.1 most of these reads passed over their errors, so a read that failed looked like a service that is not installed, no processes, no files or nothing new.

### What It Does NOT Touch

- `claude_desktop_config.json`: your MCP servers and settings are safe
- `config.json`: app configuration is safe
- Conversations: stored server-side, not in the local VM cache
- Session transcripts under `local-agent-mode-sessions`: left alone unless
  you pass `-PurgeSessions`

Versions before 6.0.0 deleted session transcripts older than 7 days on every
run, in every mode, while this section said conversations were never touched.
That deletion is now opt-in and named.

### Parameters

| Parameter | Description |
|-----------|-------------|
| `-Mode` | `Quick`, `Deep`, `Smart`, or `Diagnostic`. Skips the interactive menu. |
| `-BootPrep` | Retired in 6.3.0: it exits 0 before anything else runs, saying only that it is retired, so it asks for no elevation, writes no log or file, deletes nothing, and touches neither `vmcompute` nor any VM. It restarted `vmcompute`, and removed a stale Cowork compute system, when no Cowork workspace was up. It is kept because the logon boot-fix task older versions of Prevent installed still passes it until Prevent or `-Undo` removes that task. |
| `-SkipLaunch` | Reset the VM but don't relaunch Claude |
| `-Quiet` / `-Silent` | Suppress the interactive menu and "press any key" prompt (for scheduled tasks). Defaults to Smart mode. Holds the run back while Claude or the PC is in use: see [Exit codes](#exit-codes). |
| `-KeepCache` | Skip the VM cache purge, and in Smart mode the escalation to Deep, which is a purge too. Avoids re-downloading the VM bundle, which is roughly 13 GB, not the 2-3 GB earlier versions of this table claimed. Use when running Fix frequently. If the fix fails with `-KeepCache`, run again without it. |
| `-PurgeSessions` | Also delete session transcripts older than 7 days. Off by default. |
| `-WhatIf` | Dry run: show what would happen without changing anything |

**About that 13 GB.** `sessiondata.vhdx` and `smol-bin.vhdx` are backed up and
restored across a purge, so your workspace state survives. `rootfs.vhdx`, which
is about 9 GB of it, is not: rebuilding the VM image is the whole point of a
deep purge, and restoring it would put back whatever was wrong with it. That
9 GB is what gets re-downloaded.

### Exit codes

| Code | Meaning |
|------|---------|
| 0 | The run finished. A `-Close`, `-Mode Diagnostic` or `-SkipLaunch` run, `-BootPrep`, and a menu you cancelled end with 0 too |
| 1 | The workspace was not ready after the retries (a run without admin does not retry, and no run retries a service the prerequisite check found cannot start), whether it came up is not known because a VM log could not be read, an unexpected error stopped the run, or the `%APPDATA%\Claude` folder was not found |
| 3 | Held back before the app or the service was stopped: another Fix run holds the lock, or a `-Quiet` run found Claude or the PC in use, could not rule out a Claude Code session, or could not read whether they are in use. Step 0 may already have removed a stale compute system |
| 4 | A `-Quiet` run stopped part-way. Steps 1 and 2 had run, stopping the app and the service if they were running, and then, before a later step that stops the app, a Claude Code session the app started was found or could not be ruled out |

A `-Quiet` run, which is how the health monitor runs Fix, is held back when any of these is true: the app runs a Claude Code session, a project thread or Code session (when the parent of a Claude Code process cannot be read, it counts as one); a Claude Code transcript was written in the last 35 minutes; a `claude` process is using CPU; a VM log was written in the last 2 minutes; or there was keyboard or mouse input in the last 3 minutes. A check that cannot be made holds the run back too, and the `BLOCKED` line names what could not be read: the Claude Code sessions the app started, the transcripts, the `claude` processes or their CPU time, the VM logs, or the time of the last input, and why, naming the transcript folder when that is what could not be read, and the error that kept the input check from loading. Until 6.3.1 a transcript folder, a process list or a VM log that could not be read answered that nothing was going on, and so did an input check that failed, and the unattended repair ran. A Smart run takes minutes, and the app can start a session in that time, so a `-Quiet` run asks about the app's Claude Code sessions again right before Step 1, before Step 4 stops the app's leftover processes, and before the no-progress and Deep escalations stop the app. Before Step 1 the run is held back (3). After it the run stops part-way (4), and first starts CoworkVMService again when this run stopped it, so Cowork is not left down; when that start fails it says so, and you run Fix when Cowork is needed. Manual runs are not asked: you want the fix. Once a run is past every check that holds it back, it logs `Repair under way: this run is past every check that holds it back` (not under `-WhatIf`); the health monitor reads that line to tell a Fix run that started the repair from one that stopped before it.

When the script relaunches itself elevated, the elevated run's exit code is passed back. `Fix-ClaudeDesktop.bat` and `Stop-ClaudeDesktop.bat` pass 0, 1, 3 and 4 through. For 3 and 4 they print a line saying what happened and that the lines above say why, and they do not pause. A manual run that finds another Fix running waits for a key before it ends. Any other code means PowerShell could not run the script, and the launcher says so and pauses. The launchers find their script beside them by its name, so they start it from a folder that may be read but not listed, and from a folder whose path holds a parenthesis, such as `ClaudeFix (1)`. When the folder cannot be read at all, a launcher says the script `could not be looked for: this folder could not be read` and exits 1 (v6.3.1).

The health monitor does not count a run Fix held back (3) as an attempt at a repair: it logs `HELD BACK` and tries the repair again once the hold has cleared, waiting longer each time Fix holds it back in a row (see [Safety Features](#safety-features)). A run that stopped part-way (4) is logged as `STOPPED PART-WAY` and counts as an attempt, toward the escalation after 3 attempts in 30 minutes, and the normal cooldown follows. Until 6.3.0 a run held back exited 0, the same as a finished repair, so the monitor counted it toward "Auto-fix failed 3 times".

### Diagnostics

Each run writes a timestamped log to `%APPDATA%\Claude\fix-logs\`. The health monitor logs to `%APPDATA%\Claude\watch-logs\`. Recent CoworkVMService errors from the Windows Event Log are shown in the summary; when the Application log cannot be read, the summary says so, naming the error.

### Where the app keeps its files

The toolkit's own files stay in `%APPDATA%\Claude`: `fix-logs`, `watch-logs`,
`watch-state.json` and the support bundles. Prevent's undo record is the one
exception: it is `%LOCALAPPDATA%\ClaudeFix\prevent-undo.json`, in the local
profile, so that a roaming profile does not take it to another computer (see
[Undo Everything](#undo-everything)). The app's own files have moved
between versions, so from 6.2.0 the toolkit looks for them instead of
assuming a folder:

| Version | App logs | App data (VM cache, bundles, sessions) |
|---|---|---|
| Before 6.18355 | `%APPDATA%\Claude\logs` | `%APPDATA%\Claude` |
| 6.18355 (measured 2026-11-21) | `%LOCALAPPDATA%\Claude\logs` | `%LOCALAPPDATA%\Claude-Data\claude_<id>` |

Every folder the app has used is a candidate: `Claude-Data\*`, `Claude-3p\*`
and `Claude` under `%LOCALAPPDATA%` and `%APPDATA%`, the MSIX package's
`LocalCache` copies of both, and `CLAUDE_USER_DATA_DIR` when it is set. The
data root is the freshest candidate holding `Preferences` or `Local State`,
and the log folder is the candidate whose newest log was written last. When
nothing is found, both fall back to `%APPDATA%\Claude`, the pre-6.2.0
behaviour. The choice and every candidate passed over are written to
`Warnings` in `fix-logs\claude-env.json`, Fix prints both folders at the top
of its log, and the health monitor prints them in its startup banner, looks
again every ten minutes and whenever the app starts, and logs a line when
either folder moves. While a candidate folder cannot be read, the monitor
moves nothing and says so, at most once an hour; a candidate that
disappears while it is listed is passed over, not reported as unreadable.

What this lookup could not read is said, not taken as an answer (v6.3.1). An `hcsdiag.exe` that could not be read is named as that, where it was called not present; when `Claude.exe` is not located, each place that could not be read is named (the app package, a running Claude process, the process list, the usual install folders), so not located is not taken for not installed; and a VM log folder that could not be read is a read error, where it read as one with no VM log. These are shown with the other reads the lookup could not make: at the top of Fix's log, as `READ ERROR at start` in the health monitor, and as step 0, "Looking at this PC", in Prevent. A VM cache folder that could not be read makes the cache size in `claude-env.json` a lower bound, and its `Warnings` name the folder. A folder named `claude.exe` is not taken for the app.

**Support bundles.** `-Mode Diagnostic` changes nothing and, at the end, writes a zip you can attach to a bug report:

```
.\Fix-ClaudeDesktop.ps1 -Mode Diagnostic
```

It collects Claude's log folder, the Service Control Manager entries for `CoworkVMService`, `vmcompute` and `hns` over the last seven days, and the state of the Windows features by exact name.

The event log is the part people miss. When a service refuses to start there is very little in Claude's own logs, because nothing ran to write them, and Windows is the only thing that recorded what happened.

The log folder is enumerated rather than pulled from a list of expected filenames, because those names differ between builds. Files are ranked so that Cowork, VM and renderer logs get most of the space, and the whole bundle is built to a 4 MB budget and then zipped, so it fits an attachment limit. Missing files are reported rather than skipped quietly: no VM log at all is itself the answer. A log that is there but could not be read or copied is named, with why, as not in the bundle, and the note that the Cowork VM may never have started is not given for a VM log that could not be read; until 6.3.1 such a log was listed as skipped.

A read the diagnostic check cannot make says `could not be read`, or `could not be listed`, with the reason: the services, the `claude` processes, the vmcompute handle count, the HCS VMs, the VM cache folders, the session files, the service recovery and the Defender exclusions. Without admin, the Defender line says `not checked (needs admin to read the exclusions)`. When the log folder cannot be listed, the bundle says so and tries the known log names, and Windows features that could not be recorded are said, with the reason. The summary at the end of every run says `VM Service: could not be read (reason)` and `Claude processes: could not be listed (reason)` in the same case. Until 6.3.1 the diagnostic check called a service that could not be read `Not installed` and a VM cache folder it could not read `Not present`, and the summary called the service `Not found` and a process list it could not read `0 active`. When the app's data folder is not known, the VM cache folders say `Not known`, where the run stopped with UNEXPECTED ERROR.

---

## Prevent-ClaudeIssues

Run once. Configures Windows to keep the Cowork VM alive as long as possible. If you ran a version before 6.3.0, run it again: several settings those versions reported as made were never made, and a new run makes them, removes the retired boot-fix task, and lists what else they left.

### What It Does

| Step | Setting | Value | Why |
|------|---------|-------|-----|
| 1 | Power plan | Ultimate Performance, else High performance | Prevents aggressive power saving that kills VM mounts. The plan you had is recorded once, for `-Undo`. If an older version left `power-plan-backup.txt` holding a plan that still exists, that plan is recorded instead, or Balanced, with a warning, when it is an Ultimate Performance plan. When the power plans in the registry or that file cannot be read, the step fails, keeps the file, and makes no Ultimate Performance copy. Once the plan you had is recorded, a file that cannot be read is a warning, and is kept |
| 2 | Power plan settings (AC) | Full performance, see below | Nothing under the VM sleeps, parks or saves power while on AC |
| 3 | Hibernate | Off | Incompatible with running Hyper-V VMs |
| 4 | Fast Startup | Off (`HiberbootEnabled` = 0) | Prevents kernel hibernate on shutdown, so services don't reinitialise cleanly |
| 5 | Modern Standby | Off where the machine has it (`PlatformAoAcOverride` = 0) | Modern Standby can enter low-power states even with sleep "disabled" |
| 6 | Network adapter power saving | Off, no adapter restarted | Stops Windows and the driver powering the adapter down |
| 7 | Windows power throttling | Off (`PowerThrottlingOff` = 1) | It runs background processes on efficient cores at low clocks, and the VM's worker can be one of them |
| 8 | Cowork VM worker priority | AboveNormal, the Cowork VM's `vmwp.exe` only | Prevents host from deprioritizing the VM's worker under load |
| 9 | vmcompute recovery | Restart after 30s/60s/120s | Configures vmcompute to auto-restart on failure with escalating delays |
| 10 | Stale VM removal | Only when no process named `claude` runs, the Claude Code CLI included, and the processes can be listed | Removes stale cowork-vm entries before they accumulate and block new VM creation |
| 11 | Storage location | Checked | Warns if the app's data is on cloud-sync, USB, or network drive |
| 12 | Clock | Checked, resynced past 5 seconds | Ensures NTP is running and clock drift is within tolerance |
| 13 | Antivirus exclusions | Added for Defender / advised | Prevents AV filter drivers from blocking VirtioFS disk ops |
| 14 | WSL2 conflicts | Checked | Warns about a running WSL2 distro; lists WSL and Docker Desktop when installed. A `wsl.exe -l -v` that fails counts as no distro only while no WSL service runs, and otherwise is a warning with its exit code; a WSL feature or a `wsl.exe` that cannot be read is a warning too |
| 15 | Health monitor | Logon task, started now where safe; polls every 30 seconds from 120 seconds after it starts | Detects VirtioFS and HCS failures and runs the fix when you are idle |
| 16 | Fix Claude Desktop shortcuts | Desktop + Start Menu | Quick access to Fix-ClaudeDesktop |
| 17 | Claude (Admin) launcher | Scheduled task + Desktop shortcut | Launches Claude Desktop with full admin privileges, with no UAC prompt |
| 18 | GPU-free launcher | Written, not applied | Opt-in way to start Claude with `--disable-gpu`, for the in-app Browser crash |
| 19 | Older versions' leftovers | Removed, or listed | Removes what older versions set up that now does harm or nothing, and lists what `-Undo` removes |

**Step 2** sets these on the active plan, for AC power only, and reads each one back: sleep after, hibernate after and turn off hard disk after at Never; SATA link power management Active and SATA adaptive link power off; NVMe idle timeout (primary and secondary) off and NVMe latency tolerance (primary and secondary) 0 ms; USB selective suspend and USB 3 link power management off; PCI Express link state power management off; minimum and maximum processor state 100%; processor boost mode Aggressive; core parking minimum cores 100%, so no core is parked, and energy performance preference 0 (performance), each for both processor classes; processor throttle states off; and wireless adapter power saving at Maximum performance. A setting this machine does not have is reported as not on this machine and left alone. Until 6.3.0 five of these were set, through `powercfg /change` and `powercfg /setacvalueindex` on the active plan, and none was read back.

Battery (DC) settings and the screen timeout are not changed, so laptop users keep normal battery behaviour.

**Each step reports what it did**: changed, no change needed, skipped, check the note, or FAILED, with the reason. The summary at the end lists every step that way, and what needs a Windows restart to take effect. The exit code is 0 when every step ran and 1 when a step failed or the run stopped. Until 6.3.0 the summary was a fixed list that called every setting configured, printed after steps that had failed, and the script exited 0 either way.

**A read that fails is said** (v6.3.1). Before step 1, Prevent and `-Undo` look up where the app and its data live; each read they could not make there is printed before step 1, and the summary lists them first, as step 0, "Looking at this PC", with check the note. They do not change the exit code. A service, a process, a scheduled task or a file a step cannot read is reported as unreadable, with the reason, and the step changes nothing on it. Until 6.3.1 each was taken as one that is not there. Setup stops before anything changes, naming the folder and why, when the Claude folder in `%APPDATA%` cannot be read, and Prevent and `-Undo` stop when the undo record cannot be read (see [Undo Everything](#undo-everything)). A registry value whose key cannot be opened fails the step that reads it, naming the value, the key and why, where it read as a value that is not there. A file an older version left that cannot be read is a warning naming it, and is kept, and a file that cannot be removed is named in the step's outcome.

**`-WhatIf` is refused**, before the UAC prompt: nothing changes and the script exits 1. It used to be accepted, but the only step that honoured it was the stale VM removal, so a dry run changed the machine for real. For the list of what Prevent does, run `Get-Help .\Prevent-ClaudeIssues.ps1 -Full`.

### Cowork VM Worker Priority

**VM worker process priority** (step 8): `vmwp.exe` is the Virtual Machine Worker Process that hosts each VM on the host side. At Normal priority, it can be starved under heavy CPU or I/O load, causing the VirtioFS connection to stall or time out. Setting it to AboveNormal gives it scheduling preference. Only the Cowork VM's worker is raised: the `vmwp.exe` whose first argument is the GUID of a Cowork compute system in `hcsdiag list`. Until 6.3.0 every `vmwp.exe` on the machine was raised, which meant every Hyper-V VM, WSL2 and Windows Sandbox. Without hcsdiag the step is skipped, and an `hcsdiag.exe` that cannot be read is a warning that names it (v6.3.1), because the Cowork VM's worker cannot be told from the others, and when the Cowork VM is not running there is nothing to raise. A worker that cannot be read by its PID is a warning, and its priority is not checked or raised. A priority does not outlive its process, so the health monitor raises it again on every poll.

**Dynamic Memory pinning** was retired in 6.3.0: it asked the health monitor, through the flag file `disable-dynamic-memory.flag`, to turn off Dynamic Memory on a Hyper-V VM named like Claude, and the Cowork VM is not a Hyper-V VM and has no such setting; Prevent's step 19 and `-Undo` delete the flag file older versions left. Where an older install is found, `-Undo` also turns Dynamic Memory back on for a Hyper-V VM named like Claude that an older version left without it, once that VM is turned off.

### HCS Service Recovery

**HCS service recovery** (step 9): The `vmcompute` service (Host Compute Service) manages all Hyper-V compute system operations. If it crashes, every VM creation call fails with `HCS operation failed`. The script configures Windows Service Control Manager to auto-restart `vmcompute` with escalating delays: 30 seconds after the first failure, 60 seconds after the second, 120 seconds after the third. The failure counter resets after 300 seconds of healthy operation. This is a permanent OS-level setting that survives reboots. The recovery that was set before is recorded first, so `-Undo` puts it back. When `vmcompute` is not installed (Hyper-V is off), the step is skipped. When it cannot be read, the step fails, and its recovery is not checked or set.

### CoworkVMService Recovery

**CoworkVMService recovery** was retired in 6.3.0: Prevent no longer sets a recovery on `CoworkVMService`, Prevent's step 19 lists the restart after 30, 60 and 120 seconds that older versions set there as left in place, and `-Undo` clears it while that is still exactly what is set.

### HCS State Cleanup

**HCS state cleanup** (step 10): Over time, stale `cowork-vm` entries can accumulate in the Host Compute Service after unclean shutdowns. These orphan entries can block new VM creation when the service tries to create a new compute system. The prevention script enumerates HCS entries via `hcsdiag list` and kills any stale cowork-vm instances, but only when no process named `claude` is running, the Claude Code CLI included: with Claude open, the "stale" compute system is the one you are working in. An open Claude Code CLI, a terminal one included, also makes the step skip as "Claude is running". When the processes named `claude` cannot be listed, whether Claude is running is not known, so the step warns and leaves the cowork-vm instances alone (v6.3.1). Without `hcsdiag.exe` the step is skipped; one that cannot be read is a warning that names it, and no stale VM is looked for (v6.3.1).

Two corrections here, both from the 6.0.x audit. Every version before 6.0.0 called `hcsdiag close`, and hcsdiag has no `close` verb, so the call did nothing and then logged success. And the parse of `hcsdiag list` existed in four places across the three scripts, each with its own variation of the same two mistakes, so fixing one copy did nothing for the others. There is one shared parser now.

Three more in 6.3.0. With exactly one stale VM, the case the step exists for, the parse returned a single string, and `.Count` on a string throws under StrictMode in Windows PowerShell 5.1, so the step logged "check failed" and removed nothing. An `hcsdiag list` that failed read as an empty one, so the step said "HCS state clean"; its exit code is checked now. And a kill counts only once the compute system is gone from a second list: one still listed fails the step, and when the list cannot be read again the kill is reported as unconfirmed, with a warning.

**Service startup timeout** was retired in 6.3.0: Prevent no longer sets `ServicesPipeTimeout`; step 19 lists the 120000 an older version set as left in place, and `-Undo` removes it while it holds exactly that value, where an older install is found (see [Undo Everything](#undo-everything)).

### Network and NAT

**WinNAT check** (formerly step 16) was retired in 6.3.0: it looked for a WinNAT rule, and the Cowork VM does not use one. Older versions created a rule named `CoworkNAT`, which the health monitor re-created; Prevent's step 19 lists it when it is still there, and `-Undo` removes it, whether or not an older install is found. A machine without WinNAT has no such rule, and gets no warning. When the NAT rules cannot be read, as when access is denied, step 19 and `-Undo` warn that `CoworkNAT` was not checked, and `-Undo` ends incomplete, so that it is run again to look. For networking trouble by hand, [Jonas Kamsker's write-up](https://blog.kamsker.at/blog/cowork-windows-broken/) and [Elliot Segler's](https://www.elliotsegler.com/fixing-claude-coworks-network-conflict-on-windows.html) cover the VM's networking, the latter the subnet-collision case where `172.16.0.0/24` clashes with your own network.

**Firewall policy check** (formerly step 17) was retired in 6.3.0: Prevent no longer checks the Group Policy firewall settings.

### Storage, Time Sync, and Antivirus

**Storage location** (step 11): VirtioFS mounts fail when Claude's data directory is on a cloud-sync folder (OneDrive, Google Drive, Dropbox, iCloud, Box), an external USB, Thunderbolt or 1394 drive, or a network share. The script checks the app's data root (see [Where the app keeps its files](#where-the-app-keeps-its-files)) for all three and the VM cache for a network drive, and warns you to move Claude's data to a local SSD. A sync folder is a OneDrive folder Windows names in the environment, or a folder in the path that is named after the service or starts with that name and then " - " or " (", as "OneDrive - Contoso" and "Dropbox (Personal)" do. Below the user profile only the part after it is tested. Until 6.3.0 this was a substring test, so a folder named "Xbox" or "Boxes" read as a Box sync folder. A check that cannot be made, of which drive holds the data root, whether that drive is external or a network drive, or whether the VM cache is on a network drive, is a warning that names it and why, and `Storage location OK` is said only when every check was made; a data root on a network share with no drive letter gets a warning that neither check was made. A VM cache folder that cannot be read is a warning that names it, and whether its drive is a network drive is still checked; until 6.3.1 it read as no VM cache, and the step said `Storage location OK` (v6.3.1). Until 6.3.0 a check that failed read as a local drive, and the network drive check ran only for a drive `Get-Volume` lists, which a mapped network drive is not, so it never found one.

**Time synchronisation** (step 12): checks the host clock against NTP. The script starts the W32Time service if it is stopped, waits up to 30 seconds for `time.windows.com` to resolve, measures actual drift, and forces a resync past 5 seconds; a resync that fails is a warning. When the W32Time service cannot be read, the step warns, does not start it and does not check the clock; when it is not installed, the step warns that it is not installed. When `time.windows.com` cannot be reached within those 30 seconds, the step warns that the clock was not checked. The offset is read with either decimal separator. Until 6.3.0 it was read with a decimal point only, so where the locale writes a comma nothing was measured, and a resync that failed was reported as done. The health monitor continues to check every 5 minutes.

**Antivirus exclusions** (step 13): Antivirus filter drivers sit in the I/O path between VirtioFS and the host filesystem. They can delay or block disk operations that VirtioFS depends on, causing timeouts and mount failures. When Windows Defender is the only antivirus it finds, the script adds path exclusions for the app's data root (see [Where the app keeps its files](#where-the-app-keeps-its-files)), the Hyper-V folder, `vmwp.exe` and `vmms.exe`, and process exclusions for `vmwp.exe`, `vmms.exe`, `vmcompute.exe` and `cowork-svc.exe`, the CoworkVMService's executable. A path is added only if no existing exclusion covers it, and a process only if there is no process exclusion with that name. An existing path exclusion covers a path when it is the same path, a folder above it, or a wildcard that matches it, so an exclusion for the app's own `%LOCALAPPDATA%\Claude` does not count for `%LOCALAPPDATA%\Claude-Data` beside it. Each exclusion is recorded before it is added, so `-Undo` removes the ones Prevent added and no others; until 6.3.0 `-Undo` left them all in place. An exclusion Defender refuses fails the step, and so does a Defender whose settings cannot be read. Where Defender hides its exclusion lists from administrators, which a policy can make it do, the step warns that it cannot tell which exclusions exist and adds none: read as empty, the list made every exclusion look missing, so ones that were already there were recorded as Prevent's, and `-Undo` would remove them. When Windows' list of registered antivirus products cannot be read, the step warns that a product other than Defender may go unreported, and goes by the running processes alone; where Windows has no such list, as on Windows Server, nothing is said. Until 6.3.0 a list that could not be read meant no product registered. When the running processes cannot be listed, the step warns that an antivirus that is not Defender may go unreported (v6.3.1). For third-party AV products, it lists the recommended exclusion paths for you to add manually.

### The Health Monitor

A persistent background process that starts 120 seconds after logon and polls every 30 seconds.

Prevent's step 15 registers it as a logon task, replacing the one an earlier run made, then stops a monitor already running and starts the task at once. It starts it now only when every running monitor was found and stopped, because two monitors running at once would each repair Claude on their own. When the running processes cannot be listed, a monitor is still running 2 seconds after it was stopped, or whether it stopped cannot be read, the step warns, naming why, and the new monitor starts at the next logon. The scheduler ignores a start while the task still reads as running, which it can do for a moment after the old monitor has ended, so the step waits up to 5 seconds for it to stop reading as running, starts it, and checks for up to 3 seconds that it runs; only then does it say "Started it now: the monitor begins watching in 120 seconds". When the task still reads as running after those 5 seconds, its state cannot be read before the start, or it reads another state 3 seconds after the start, the step warns that the monitor starts at the next logon. When the task is no longer there 3 seconds after the start, the step warns that it is gone, and when its state cannot be read then, that whether the monitor runs could not be checked. Until 6.3.1 neither was checked, and "Started it now" was said of a start the scheduler ignored. A start that fails is a warning too, and a registration that fails fails the step and leaves the running monitor as it is. Until 6.3.0 the old task was removed and the running monitor stopped before the new task was registered, and every error was dropped, so a registration that failed left no monitor at all and was reported as installed, and the new monitor was started beside one that did not stop. The step looks for `Watch-ClaudeHealth.ps1` beside Prevent, then in `C:\ClaudeFix`, the Desktop and Documents. When one of those places cannot be read, the step fails, naming it and why, and installs no copy found later; until 6.3.1 it was passed over. An old `cowork-watchdog.ps1` that cannot be read is a warning, and is kept (v6.3.1).

1. **Claude log files**: scans the app's VM log (`cowork_vm_node*.log`) and main log (`main*.log`) for a VirtioFS mount or HCS failure signature inside an entry the app logged at `[error]`, reading only what each file has gained since the last poll. A file seen for the first time is read from its current end, so old errors are not replayed. Lines at `[info]` or `[warn]`, lines from `[remote-bash]` (remote sessions' commands), and the browser, web and MCP logs never trigger a repair: app 6.18355 logs a harmless `RPC error -1` at `[info] [remote-bash]` when a remote session's probe finds its worker already running. A log folder that cannot be listed is said, at most once an hour, and what was read of each log so far is kept until it can be listed; a folder that is not there is silent (v6.3.1).
2. **Service status**: detects when `CoworkVMService` stops while Claude is still running (2 consecutive checks). When the service or Claude's processes cannot be read, it says so, at most once an hour, and does nothing, with its count left as it was (v6.3.1)
3. **Windows Event Log**: checks for Claude-specific `CoworkVMService` errors and Hyper-V Worker/VMMS errors (2 consecutive checks, whole message examined, not just the first line)
4. **Workspace storage**: checks that the VM cache is readable and its volume has room (warning only). Each different `STORAGE WARNING` for a VM cache folder that cannot be read or a volume low on space is logged at most once an hour, with no cap on how many, so a new warning, or one whose text changes to one not logged in that hour, is logged at once, where until 6.3.1 it was logged on every poll; a VM cache folder that cannot be read is named (v6.3.1). A free space that cannot be read is logged at most once an hour for the same error
5. **VM log staleness**: warns when the VM stops writing logs (5 consecutive stale checks, 5-minute threshold, only if the VM was previously active). Warning only: it starts no repair, because an idle VM and a hung VM look the same from the host. Every VM log is read on each poll, and `main.log` is not one: it can sit frozen for a day on a perfectly healthy machine. When the VM logs cannot be read, the check says so, at most once an hour, and does nothing, with its count left as it was (v6.3.1).
6. **Clock drift**: checks the clock against `time.windows.com` every 5 minutes and forces an NTP resync past 5 seconds (warning only). The resync is logged as repaired, with the drift before and after, only when the clock, measured again, is within 5 seconds. Otherwise the warning gives the drift, why the resync did not help and the time source Windows syncs from, and no resync is forced for 60 minutes while the drift lasts; the clock is still measured every 5 minutes, the warning is repeated while the drift is over 30 seconds, and the monitor logs when it is back within 5 seconds; a wait that reaches more than 60 minutes ahead, because the clock was set back during it, is ended. Each w32tm call is given 15 seconds. A resync, a measurement after it or a time source query that fails is named in the warning, and the drift is still reported. When the check cannot be made, because w32tm cannot be started or answers with an error and no offset, as when it cannot reach the time server, a `TIME WARNING` says that the clock check failed and names why, and a measurement that does not answer within 15 seconds is logged as a skipped check; each is logged at most once an hour for the same error. The offset is read with either decimal separator. Until 6.3.0 a resync was logged as repaired without measuring again, a resync that failed included, one that changed nothing was forced again every 5 minutes, and a check that could not be made was passed over without a word. From 6.3.1 a Windows Time service that cannot be read is said, at most once an hour, and is not started then, while the clock is still measured; one that is not installed is passed over without a word. The 60 minutes with no forced resync run from when the resync finished, not from just before it
7. **vmcompute health**: monitors `vmcompute.exe` handle count every 60 seconds. Warning at 12000, trigger at 20000, 2 consecutive checks required. A process list or handle count that cannot be read is said, naming the error, and leaves the run of high counts as it was (v6.3.1)
8. **HCS state health**: watches for stale cowork-vm entries via hcsdiag and tracks 0xC037010D shutdown failure frequency. Kills stale instances, keeping the one whose worker process started last, compared in UTC so that the hour the clocks go back cannot reverse the order, and confirms each kill by listing again: one still in the list is logged as still listed after the kill. With no worker process to tell the live one by, it kills none, and when the worker processes cannot be listed it kills none and says so. From 6.3.1 it also kills none while a worker's command line, or a Cowork VM worker's start time, cannot be read, and says so at most once an hour. Until 6.3.0 it kept the last entry hcsdiag printed and logged every kill it sent as done. Restarts vmcompute if more than 15 shutdown failures occur within 5 minutes. That restart is held, and this is said, when the Cowork service or Claude's processes cannot be read (v6.3.1). Guarded by active-session detection and a check for a running Fix

**A check that fails is said** (v6.3.0). When one of the checks named in this paragraph cannot be made, it logs one line that names it and the error. The same error is logged at most once an hour, and no more than 3 different errors of one check in an hour. That covers the Claude logs, the Event Logs, storage, VM log staleness, the clock, vmcompute, the HCS checks, the VM worker priority, the state in `watch-state.json`, the Fix logs and Fix's lock, the lock that keeps a second monitor from running, the session the monitor runs in, the focus check before a repair, and the poll itself, as `MONITOR ERROR`. A Fix lock the monitor may not open, as when Fix runs elevated and the monitor does not, means Fix is running, and so does one it cannot ask about at all, so no repair starts on top of a running Fix. When the check that the monitor is elevated fails, the line that skips the repair names the error. Until 6.3.0 `MONITOR ERROR` was written on every poll and the rest were passed over without a word, so a check that could not run looked like one that found nothing wrong: a vmcompute handle count that could not be read was taken as 0, any error from Fix's lock read as Fix not running, and a failed elevation check read as a monitor that is not elevated. From 6.3.1 the same goes for the service check, the Windows Time service, the Claude log folder, the Cowork VM log, the list of Claude's processes and their start times, the VM workers' command lines and start times, the hcsdiag job, the activity checks loaded at start, the look for the app's folders and the deletion of old watch logs, and none of them acts on a read it could not make. A poll that cannot list Claude's processes checks nothing and does nothing; the service check and the staleness check do nothing and keep their counts; and the guards before a repair take a session, or the user, to be active (see [Safety Features](#safety-features)). The reads discovery could not make are logged at start, each as `READ ERROR at start`. Until 6.3.1 a service whose status could not be read counted as one that is not installed, and a process list that could not be read as one with no process in it. Its storage check reports a VM cache folder that cannot be read, naming it, where it read as no VM cache (v6.3.1).

**A saved state that cannot be read holds repairs back** (v6.3.1). `watch-state.json` keeps the cooldown, the repairs of the last 30 minutes, which stop a fourth after three, and a held-back repair. When the file is there but cannot be read or parsed, is empty, or holds no saved state, a cooldown starts at once, as after a repair, and no automatic repair runs until it is over: `-Cooldown` minutes, 5 by default, or a longer wait already restored from the file. This is logged at most once an hour, as `STATE: The state saved in ... could not be restored in full`, with the reason and how long no repair runs. What could not be restored starts empty. Until 6.3.1 such a file read as no state at all, as on a first start, so a repair loop bad enough to leave the file broken could run again at once; an empty file, or one that could not be looked for, did so without a word. A file that is not there is a first start, and is silent.

**What it does not monitor, and why.** Earlier versions listed a Hyper-V
Integration Services heartbeat check. There is no such heartbeat to read: the
Cowork VM is an HCS compute system and does not appear to `Get-VM` at all, so
that monitor could never fire, and from 6.0.0 it did nothing. 6.3.0 removed it
from the monitor and from its help; the startup banner names the heartbeat
under "Not monitored". Liveness comes from log write times and hcsdiag
instance state instead.

**Retired in 6.3.0: the guest-connection and NAT checks.** The
guest-connection check read `cowork-service.log` for a message that file has
never held, and the file is no longer written. The NAT check looked for a
WinNAT rule, and the Cowork VM does not use one. A `Request timed out:
isGuestConnected` the app logs at `[error]` in its VM or main log is still
caught by the log check (item 1).

### Safety Features

The entire toolkit is designed to **never interrupt active work**, whether you're in Chat, Cowork, or Code. Version tags below mark when each safeguard landed:

- **Electron-aware activity detection**: uses `GetWindowThreadProcessId` to correctly detect Claude's Electron renderer windows (the old `MainWindowHandle` comparison failed because Electron child processes own the visible window, not the main process)
- **Session 0 safe**: when running as a SYSTEM scheduled task (Session 0), Win32 window/input APIs return garbage. The monitor detects this and falls back to process-only heuristics (VM log + CPU sampling). When the foreground-window and input checks cannot be loaded, the monitor says so after its startup banner, naming the error, and skips them, as in Session 0 (v6.3.1)
- **Claude Code session awareness**: detects active Code sessions via a running Claude Code CLI the app started, for a project thread or Code session (v6.3.0; when the parent of such a process cannot be read, the monitor assumes a session is active), Claude Code's own transcripts (`%USERPROFILE%\.claude\projects\**\*.jsonl`, or under `CLAUDE_CONFIG_DIR`), the app's session persistence files, and its logs (`unknown-window*.log`, `claude.ai-web*.log`, `remote-control*.log`). A session file or log counts when written in the last 10 minutes, a transcript in the last 35: a tool call writes nothing to the transcript while it runs, and a foreground command may run for 30 minutes. The transcripts are the signal that keeps up: measured 2026-11-21, two working project threads had written theirs that minute while the app's session file and logs were 13 to 38 minutes old. A terminal Claude Code session counts too, which only delays a repair by two minutes at a time. Code runs inside Claude Desktop but doesn't use the Cowork VM, so VM staleness is expected and normal during Code usage. When a folder, the process list or a CPU time the check needs cannot be read and nothing else shows a session, the monitor assumes a session is active, and says why at most once an hour (v6.3.1); until then a folder that could not be read was passed over as one with nothing recent in it. Without this check, the monitor could kill an active Code session while trying to fix a "stale" VM that was never needed
- **Extended CPU sampling**: the 30s grace period re-check now takes 3 samples x 1s (was 1 x 500ms) with a lower threshold (50ms, was 100ms). Code has bursty CPU with long near-zero periods during API waits; the wider sampling window catches these patterns
- **CPU sampling**: measures actual processor time on Claude processes over 500ms. Catches active request processing even when the UI is idle (e.g., Code thinking)
- **Extended VM log window**: activity check uses 120s window (was 30s). Code's "thinking" phases can leave the VM log quiet for minutes; the old 30s window caused false negatives
- **User input window**: 3 minutes (was 2). More buffer for reading/reviewing before auto-fix considers you idle
- **Fix script activity guard**: `Fix-ClaudeDesktop.ps1` itself checks for active use when called with `-Quiet`, as the health monitor calls it. Five checks: a Claude Code CLI the app started (when the parent of such a process cannot be read, the repair waits), a Claude Code transcript written in the last 35 minutes, CPU sampling, VM log, user input. Holds the run back and exits 3 if anything is active, and, from 6.3.1, when one of the checks cannot be read, naming it; from 6.3.0 it also asks about the app's Claude Code sessions again before each later step that stops the app (see [Exit codes](#exit-codes)). Until 6.3.0 only the last three ran, and a project thread waiting on a 25-minute command uses almost no CPU and writes nothing, so an unattended run passed all three and ended it. Manual runs (no `-Quiet`) always proceed
- **No `vmcompute` restart at logon** (v6.3.0): the boot-fix task, which ran Fix in `-BootPrep` mode 45 seconds after every logon, was retired. It restarted `vmcompute` while Claude was not yet running, which is a restart of the service every VM on the machine depends on. Prevent and `-Undo` remove it. `-BootPrep` itself is retired too: until they do, the task runs Fix, which exits at once, saying only that `-BootPrep` is retired, and writes no log or file, deletes nothing, and touches neither `vmcompute` nor any VM
- **Startup grace period**: all heuristic checks (log scanning, event log, staleness) are skipped for the first 180 seconds after the monitor starts, preventing false triggers from pre-existing events
- **Consecutive-check gates**: every heuristic trigger requires multiple consecutive failures before firing (service: 2, event log: 2, vmcompute handles: 2). The staleness warning also waits for 5 stale checks. Only the log-file check (a VM failure signature in an `[error]` entry) triggers immediately
- **Tightened event log matching**: Hyper-V VMMS events must mention "claude" or "cowork" (no generic "failed"/"unexpected" matching). Worker events: Critical/Error only
- **VM log staleness requires prior activity**: the staleness warning is only logged if the VM log was previously active this session, so Chat mode does not raise it; it never starts a repair
- **5-minute cooldown**: between auto-fixes. The `COOLDOWN` line says when a fix is allowed again (v6.3.0). A run Fix held back is not counted as one: the wait it leaves is 2 minutes, doubled for each run Fix holds back in a row, up to 60 minutes (v6.3.0)
- **Times in UTC** (v6.3.0): the file and event ages Fix and the monitor measure, the monitor's own times and the times it keeps in `watch-state.json` are in UTC, so the change to or from summer time no longer moves them by an hour. Until then a file written a minute before the change to summer time read as an hour and a minute old. A `watch-state.json` an older version wrote is converted when it is read. Each Event Log query in Fix and the monitor is given the start of its window as UTC text, so the window is right in the hour the clocks go back, which a local time names twice. One thing stays in local time on purpose: the times in the app's own log lines, which the app writes in local time with no offset, so the Cowork session check compares them with local time, and for the hour around a change of the clock a line can read up to an hour newer or older than it is
- **30s pre-fix warning with notification**: before any auto-fix, a Windows balloon notification with an audible chime warns you. You have 30 seconds to switch to Claude to cancel. If you don't, the fix proceeds. If you do switch to Claude, the fix cancels and a second notification tells you to run Fix-ClaudeDesktop.bat manually if Cowork is broken. A trigger that fires again is let through about 2 minutes after the cancel, and a held-back repair waiting to be tried again ends with the cancel; until 6.3.0 those 2 minutes were timed from before the warning, so the check came back a minute or so after the cancel. From 6.3.1 the fix also cancels when Claude's processes, their CPU time or the VM logs cannot be read during the 30 seconds, and the `CANCELLED` line names the read. When whether Claude is in focus cannot be read, that is said, and the other checks still decide
- **Smart cancellation**: the 30s grace period only cancels if Claude is *actively being used* (foreground window, CPU activity, VM log alive, or active Code session). General mouse/keyboard activity in other apps does **not** cancel the fix, so a genuinely hung VM still gets repaired while you're browsing or gaming
- **Invoke-HcsDiag timeout wrapper** (v5.0.0): all hcsdiag calls across Fix and Watch are wrapped with a 15-second Start-Job timeout to prevent indefinite hangs when HCS is corrupted. From 6.3.0 a call that exits with an error code counts as no answer, and its exit code is logged once, and again only when what hcsdiag says changes or after it has answered in between, not on every check. An empty `hcsdiag list` counts as nothing listed, where it read as no answer. From 6.3.1 a background hcsdiag call that cannot run hcsdiag says why, at most once an hour, instead of reading as hcsdiag not answering
- **vmcompute restart session guard** (v5.0.0): the health monitor checks for active Claude sessions and Fix mutex before restarting vmcompute, preventing disruption of running Cowork sessions
- **ServiceTimeout raised** (v5.0.0): Fix script service stop timeout increased from 8s to 30s, preventing premature force-kills that cause HCS corruption
- **Prerequisites cannot veto a start** (v6.0.3): the service is always asked to start, and the prerequisite findings are only read out if it refuses. Before this, a wrong prerequisite verdict returned failure without ever calling `Start-Service`, so the script reported "Service failed to start" for a start it had never attempted, on a machine where the service came up in 293ms the moment anything asked it to
- **Bounded restart attempts** (v6.0.2): the workspace wait restarts a dead service at most three times, and stops immediately when a known blocker survives a real failed start. It used to retry every five seconds until the deadline or the keyboard
- **Session transcripts are opt-in** (v6.0.0): deleting transcripts older than 7 days now requires `-PurgeSessions`. Every earlier version did it on every run, in every mode, while this README said conversations were never touched

When a failure is detected **and the user is idle**, it shows a warning notification with a 30-second countdown, then runs `Fix-ClaudeDesktop.ps1 -Quiet`. If you switch to Claude during the countdown, the fix cancels and you're notified to run it manually. If the user was already detected as active before the countdown, or a Claude Code session the app started is running, or a Claude Code transcript was written in the last 35 minutes, it logs a `BLOCKED` message and waits. It does the same when those sessions cannot be listed or the transcripts cannot be read, and from 6.3.1 the line says why, naming the transcript folder. In the first 120 seconds after Claude starts it logs `DEFERRED`. It logs `DEFERRED` too when Claude's processes cannot be listed or the start time of one of them cannot be read, since whether Claude has just started is then not known (v6.3.1). When Fix itself holds the run back, the monitor logs `HELD BACK`, which is not counted as an attempt; when Fix stops part-way, it logs `STOPPED PART-WAY` and counts the run as an attempt (see [Exit codes](#exit-codes)).

**A repair that is held back** (v6.3.0). A check that is made again on every poll, such as the service check, fires again while the fault lasts. A log error or an Event Log error fires once, so when Fix holds its repair back (exit 3), or the monitor logs it as `BLOCKED` or `DEFERRED`, the repair is kept pending in `watch-state.json`, and so is one for such an error that fires during the wait a hold leaves. It is tried again once the hold has cleared, with the 30-second warning at most 3 times. It ends after 4 hours, when Claude Desktop is started again, or when a Fix run logs `Repair under way`, since that run was the repair. When the time Claude Desktop started cannot be read, the repair is not ended on it, and this is said. A Fix run that stops before that, such as a Diagnostic run, a Cancel or a No at "Proceed?", does not end it, and while a Fix run holds Fix's lock the repair waits for it to end. A log error or an Event Log error that fires while a repair is pending takes it over: the repair is tried for the new reason, its 4 hours start again from then, and its tries are kept. In the cooldown, the `COOLDOWN` line for it is followed by one that says it takes over the held-back repair, naming the reason that was pending, and is tried again once the wait is over, for up to 4 hours. The error the repair is already pending for, firing again, takes nothing over: its 4 hours start again, and that line says only that it is tried again. Each run Fix holds back in a row doubles the wait before the next repair: 2, 4, 8, 16, 32, then 60 minutes. The count starts again after a run Fix does not hold back, after a hold 4 hours or more after the one before, and after a Fix run by hand that started the repair, which also ends the wait 5 minutes after that run at the latest. A trigger in the cooldown after a repair that ran, an escalation or a cancel is skipped.

`BLOCKED` and `COOLDOWN` lines for the same reason, and a held-back repair's "still waiting" line, are written at most once every 10 minutes, and so is `AUTO-FIX SKIPPED: Fix-ClaudeDesktop is already running` when no repair is pending; with one pending, it is written once for each Fix run that holds Fix's lock. Until 6.3.0 a trigger checked again on every poll wrote the same lines about 30 times an hour while a hold lasted.

**Stealth-update relaunch** (v6.1.0): Claude applies background updates by quitting itself once idle and handing off to the Electron updater. On the MSIX path `quitAndInstall()` calls the AppX deployment API with `ForceApplicationShutdownOption`, and registering a package does not start the app afterwards, so the app stays down until someone starts it by hand. Before it quits, the app writes a `stealth-relaunch` marker file recording window position and navigation history, and restores from it if the next launch happens within five minutes. The monitor reads that marker and performs the launch the update path never does, via the same `shell:AppsFolder` activation the fix script uses, so on the default 30-second poll the app is usually back inside the restore window and you see nothing at all. The marker is the discriminator, and it is never deleted here: it exists **only** when the app went down through the update path, and a deliberate quit writes nothing, so an app you closed on purpose still stays closed. A marker older than `-MaxStealthMarkerAgeMinutes` (default 10) means the machine was off or you chose to stay closed, and is left alone. The marker's timestamp is recorded before the launch is attempted, so a launch that fails cannot become a relaunch attempt every thirty seconds. `-NoStealthRelaunch` turns the whole thing off. A marker that cannot be looked for is logged as `STEALTH WARNING: marker at ... could not be read`, once while it lasts, where it read as no marker. When `watch-state.json` could not be restored, a marker written before this monitor started is not acted on, since whether the monitor before it already relaunched Claude for it is not known: it is recorded as decided, and logged (v6.3.1).

The monitor also performs continuous maintenance: re-applying AboveNormal priority to the Cowork VM's worker (`vmwp.exe`) on every cycle. When the worker processes cannot be listed, or the worker cannot be raised, it says so, naming why; a worker that has exited since the list was read is passed over. Until 6.3.0 it raised every `vmwp.exe` on the machine, and it also applied deferred Dynamic Memory changes, which were retired because the Cowork VM is not a Hyper-V VM.

The monitor uses a global mutex to ensure only one instance runs at a time; when it cannot take it, it says so and monitors without it. It logs its activity to `%APPDATA%\Claude\watch-logs\`, and deletes the watch logs older than 30 days at start and at the daily rollover. Logs it cannot delete are said, naming the folder and the error, and tried again at the next rollover; a log folder that is not there is not an error (v6.3.1). At start it looks for `Fix-ClaudeDesktop.ps1` beside itself, then in `C:\ClaudeFix`, the Desktop and Documents. When one of those places cannot be read, it stops with `FATAL: Fix-ClaudeDesktop.ps1 at '...' could not be read`, naming the error, and does not start with a copy found later (v6.3.1).

Visible in Task Scheduler under `\Claude\ClaudeCoworkWatchdog`. Can also be started manually with `Watch-ClaudeHealth.bat` for foreground monitoring.

### Claude (Admin) and Claude (no GPU) Launchers

**Claude elevation** (step 17): Claude Desktop is installed as an MSIX (Microsoft Store) package. By default, it launches with a standard (non-elevated) user token, even if you're an administrator. This means its child processes (including MCP servers like Desktop Commander) also run without admin privileges and cannot perform system-level operations. MSIX apps block all direct `.exe` access from `WindowsApps` (ACLs, `Start-Process -Verb RunAs`, `dir` enumeration all fail), so the only reliable approach is a **scheduled task**. The script creates a `\Claude\LaunchClaudeAdmin` task with `RunLevel=Highest` + `LogonType=Interactive`, which gives the process a full unfiltered admin token with no UAC prompt, while keeping the GUI visible in the user's desktop session. The task's action finds Claude at runtime via two methods: (1) `Get-AppxPackage` for MSIX installs, (2) common install paths for traditional `.exe` installs. A third, reading the path off a running `claude` process, was removed in 6.2.0: it only ran when no app process was running, so all it could find was the Claude Code CLI. This survives version updates and works with any install type. A "Claude (Admin)" desktop shortcut triggers this task via `schtasks /run`, through a hidden Windows Script Host script that shows a message when the task is missing. **Note:** MSIX installs will show a second Claude icon on the taskbar when launched elevated. This is unavoidable, Windows enforces medium integrity for all shell-activated MSIX apps, so the only way to get a full admin token is to launch the `.exe` directly, which bypasses the MSIX app model's icon grouping. The scheduled task includes a process guard: if Claude is already running when the task is triggered (e.g., clicking the Desktop shortcut while Claude is open), it exits cleanly without launching a second instance. A Claude Code CLI, which also runs as `claude`, does not count; a process whose path cannot be read does, so a doubt ends in no second launch. When the running processes cannot be read at all, the task does not start Claude: it shows a message that Claude was not started because the running processes could not be read, with the error, and exits 1 (v6.3.1). When `Claude.exe` at one of the places it looks, or the list of app packages, cannot be read, it shows `Claude was not started` with the reason and exits 1, and starts no Claude from a later place; a folder named `Claude.exe` is not taken for it (v6.3.1). When the scheduler refuses the task, the step fails with the scheduler's reason and no shortcut is made, and a "Claude (Admin)" shortcut an earlier run left on the Desktop whose task is gone is removed. An old task that cannot be read or removed is a warning, and registering the new one replaces it. The Desktop shortcut uses a copy of Claude's icon, `Claude.ico` in `%APPDATA%\Claude`, where a Claude update cannot remove it, or a stock Windows icon when there is no Claude to copy it from. Until 6.3.0 it pointed at `Claude.exe` itself, inside a folder whose name carries the app version, so the first Claude update after setup blanked it.

**Admin token policy** was retired in 6.3.0: Prevent no longer sets `LocalAccountTokenFilterPolicy` or `FilterAdministratorToken`; step 19 lists them when an older version's values are still there, and `-Undo` removes them while they hold exactly 1 and 0, where an older install is found (see [Undo Everything](#undo-everything)).

**GPU-free launcher** (step 18): Current Claude builds have a fatal GPU-process crash, [upstream issue #80444](https://github.com/anthropics/claude-code/issues/80444), exit code `101457950` (`0x060C201E`). A page opened in the in-app Browser preview runs a WebGL capability probe, the GPU process dies, and the whole app dies with it. Running Claude with `--disable-gpu` avoids it, but an MSIX app launched normally never receives a command line, so the flag has nowhere to go. The script writes `Launch-Claude-NoGPU.ps1` and `Launch-Claude-NoGPU.cmd` into `%APPDATA%\Claude` and a "Claude (no GPU)" desktop shortcut with the same copy of Claude's icon, which use `Invoke-CommandInDesktopPackage` to start the executable *inside* the package identity, so the flag arrives and the install stays intact. On a traditional install the flag goes straight on the command line. The launcher closes the running app first, and never a Claude Code CLI, which would end that work: Claude is single-instance, so launching a second copy over a running one just focuses the existing window and the flag is silently ignored. It starts Claude again only once no copy it tried to close is left: when one is still running, or whether it closed cannot be read, it starts nothing, and when the running processes cannot be read at all it closes nothing and starts nothing. While a running Claude process cannot be read, as one started as administrator or a Claude Code session, it closes nothing and starts nothing too, since a copy already running would take the launch and ignore the flag (v6.3.1). Either way it says why and waits for Enter. After the start, it says when it cannot read the running processes to check that Claude runs (v6.3.1). The launchers are written by Prevent, so run it again to get these changes. This is opt-in, the normal shortcut is untouched and still uses the GPU, and `-Undo` removes the launcher, the shortcut and the icon copy. Expect the UI to feel slightly slower without hardware acceleration.

### Hibernate, Fast Startup, Modern Standby, NIC Power Saving, Power Throttling

These settings are the most commonly overlooked causes of VirtioFS failures. Each is recorded before it is changed, so `-Undo` puts back what was there:

**Hibernate and Fast Startup** (steps 3 and 4): With Fast Startup on, a Windows "shutdown" is actually a kernel hibernate. Services like `CoworkVMService` don't fully reinitialise on the next boot, which can leave stale VM state behind. Turning hibernate off (`powercfg /hibernate off`) usually kills Fast Startup too, but the script also explicitly sets the `HiberbootEnabled` registry value to 0. Hibernate is read back after it is turned off.

**Modern Standby** (step 5): On newer hardware (most laptops since ~2018, some desktops), Windows uses Modern Standby instead of traditional S3 sleep. Modern Standby can enter low-power states *even when sleep is "disabled"* in the power plan. The script asks `powercfg /a` whether the machine uses it and, where it does, sets `PlatformAoAcOverride = 0` in the registry. Requires a restart. Until 6.3.0 it set `CsEnabled = 0`, which current Windows no longer reads, so on a machine with Modern Standby nothing changed and the summary said it was off.

**Network adapter power saving** (step 6): Windows, the driver framework and the adapter's driver can each power a network adapter down to save power. The script turns that off on every physical adapter in the five places it lives: Windows' "Allow the computer to turn off this device to save power" (`MSPower_DeviceEnable`); the driver framework's power-down while the PC is running, `IdleInWorkingState` in the `WDF` key under the adapter's Enum key, set to 0 where the value is there, with the key never created; bit 0x08 of `PnPCapabilities` in the adapter's driver key, which tells NDIS not to power the adapter down, with the wake bits beside it left as they are and the key never created; the driver's own options by registry keyword (`*EEE`, `EEELinkAdvertisement`, `AdvancedEEE`, `EnableGreenEthernet`, `GigaLite`, `PowerSavingMode`, `AutoDisableGigabit`, `ULPMode`, `SipsEnabled`, `MIMOPowerSaveMode`, `bLeisurePs`, `LPMEnable`, `*IdleRestriction`, `EeeCtrlMode`, `GreenEthernet`, `APSmode`, `AutoLinkDownPcieMacOff`, `EnableLDPS`, `SleepPhyOnLinkDown`, `EEEMode` and `LowPowerIdleMode`), each set to the value that turns it off, which is 0 for most, 3 for `MIMOPowerSaveMode`, 1 for `*IdleRestriction` and 2 for `EeeCtrlMode`, and only where the driver lists that value as a valid one, and two options some Realtek drivers keep in the registry alone, `PowerDownPll` (Ethernet) and `bFwCtrlLPS` (Wi-Fi), set to 0 only when they read exactly 1; and selective suspend, sleep on disconnect and D0 packet coalescing. Each of those last three lives in a driver option (`*SelectiveSuspend`, `*DeviceSleepOnDisconnect`, `*PacketCoalescing`), and a driver that has no such option gives Windows nowhere to write the setting: it is then left as it reads, nothing is recorded for `-Undo`, and the step warns that the driver has no option for it. Selective suspend is not warned about where the driver framework's power-down while idle (`IdleInWorkingState`) is off, as on Realtek's 2.5GbE `rt25cx21` driver: that switch is the one such a driver is powered down by, and Windows' reading of selective suspend stays `Enabled` whatever it is (v6.3.2). Until 6.3.2 the hidden options failed with "Cannot validate argument on parameter 'DisplayValue'", because the option was piped to the cmdlet and its empty display value went with it, and selective suspend on such a driver failed the step with "SelectiveSuspend still reads Enabled". Hidden adapters, such as Wi-Fi Direct, and virtual ones, such as a VPN's, have Windows' "Allow the computer to turn off this device" box cleared, even when no physical adapter is found, and when the physical adapters cannot be listed: that fails the step, with "Could not list the physical network adapters, so their power saving was not changed" and why, and no physical adapter is changed or taken for a hidden or virtual one. When the hidden and virtual adapters cannot be listed, or whether Windows may turn them off cannot be read, the step warns. No adapter is restarted: the changes take effect at the next Windows restart, and the summary says so. Until 6.3.0 `PnPCapabilities` was written to the adapter's Enum key, where nothing reads it, and the adapters were restarted, which dropped the network link for about 14 seconds and made the clock check after it fail. Those versions also turned off the adapters' wake settings; `-Undo` turns them back on for the adapters they touched.

**Power throttling** (step 7): Windows' power throttling runs background processes on efficient cores at low clocks, and the VM's worker can be one of them. The script sets `PowerThrottlingOff = 1`, which turns it off for every process from the next Windows restart.

### The Boot-Fix Task

Retired in 6.3.0: the `\Claude\ClaudeCoworkBootFix` task, which ran `Fix-ClaudeDesktop.ps1 -BootPrep -Quiet` 45 seconds after every logon, restarted `vmcompute` while Claude was not yet running, which is a restart of the service every VM on the machine depends on, so Prevent's step 19 and `-Undo` now remove it. When step 19 cannot read the task, it warns that the task was not checked or removed (v6.3.1). `-Undo` fails its step 3 on a task it cannot read, as it has since 6.3.0.

### Shortcuts

Step 16 creates a "Fix Claude Desktop" shortcut on the Desktop and in the Start Menu, pointing at `Fix-ClaudeDesktop.bat`, with a stock Windows icon. You can pin the Start Menu entry to your taskbar for one-click access. When `Fix-ClaudeDesktop.bat` cannot be found, the step fails and says so: keep all the scripts in one folder. It looks for `Fix-ClaudeDesktop.bat` beside Prevent, then for `Fix-ClaudeDesktop.ps1` in `C:\ClaudeFix`, the Desktop and Documents, with its `.bat` beside it. When one of those places cannot be read, the step fails, naming it and why, and makes no shortcut to a copy found later (v6.3.1).

### Undo Everything

```
.\Prevent-ClaudeIssues.ps1 -Undo
```

Every setting Prevent changes is recorded first, with the value it had, in `%LOCALAPPDATA%\ClaudeFix\prevent-undo.json`. The first record of a setting wins, because a rerun finds the value Prevent set, which is not the original. The record is in the local profile, not the roaming `%APPDATA%`, and names the computer that wrote it by its MachineGuid, which renaming the computer does not change, and by its name, so a roaming profile cannot take it to another computer and have that one put back this one's values. A record written on another computer is never used: Prevent moves it aside, as `prevent-undo.from-<computer>.<time>.json` in the same folder, and starts this computer's own, and `-Undo` puts nothing back from it, fails its first step and says why. `-Undo` runs five steps:

1. **Recorded settings**: puts back everything a run of 6.3.0 or later recorded, newest first, then makes the plan you had active again, or Balanced, with a warning, when that plan is gone or is the one Prevent created. It deletes only the Ultimate Performance plan Prevent created, and keeps that one while it is still the active plan. A power setting Windows no longer has, or a driver option the adapter's driver no longer lists, is dropped from the record with a warning. A setting for a network adapter that is not plugged in stays in the record: run `-Undo` again with it connected.
2. **Older versions' changes**: removes what versions before 6.3.0 set without a record. Their power plan backup is put back, or Balanced, with a warning, when it holds an Ultimate Performance plan or a plan that is gone. A backup that cannot be read fails the step and is kept for the next `-Undo`. When even whether `power-plan-backup.txt`, `connected-standby-backup.txt` or `disable-dynamic-memory.flag` is there cannot be read, that is a warning, the file is kept, and `-Undo` ends UNDO INCOMPLETE (v6.3.1). On each network adapter those versions marked with `PnPCapabilities = 24` in its Enum key, the wake settings are turned back on before the value is removed. `CsEnabled` is put back from its backup; when it was put back but the backup could not be removed, the step fails and says that it was put back. The `CoworkNAT` rule and the Dynamic Memory flag file are removed. `ServicesPipeTimeout = 120000`, `LocalAccountTokenFilterPolicy = 1` and `FilterAdministratorToken = 0` are removed while they hold exactly those values, and only where an older install is found: a file only those versions wrote in `%APPDATA%\Claude` (`power-plan-backup.txt`, `connected-standby-backup.txt`, `disable-dynamic-memory.flag`, `cowork-watchdog.ps1` or `Launch-Claude-Admin.cmd`), their boot-fix task, or the note Prevent makes of one of these before it removes it. When one of those files or the boot-fix task cannot be read, whether an older version ran is not known: the step warns, leaves the three values as they are, and `-Undo` ends UNDO INCOMPLETE, so the next `-Undo` looks again (v6.3.1). An administrator, another tool or a policy can set the same values for reasons of their own, so elsewhere each is logged as matching what older versions set and left as it is. Where an older install is found, Dynamic Memory is turned back on for a Hyper-V VM named like Claude that is turned off; for one that is running, the step warns you to turn it off and run `-Undo` again. The restart after 30, 60 and 120 seconds on `vmcompute` and `CoworkVMService` is cleared while it is still exactly that, unless step 1 could not put back a recorded recovery: then that record is kept, and step 1 puts it back on the next `-Undo`. A service that cannot be read is a warning, and `-Undo` ends incomplete. Hibernate and Fast Startup turned off by those versions are left as they are, and so are Ultimate Performance plans: there is no record of what they were, or of which ones those versions made.
3. **Monitor and tasks**: stops the health monitor and removes the `ClaudeCoworkWatchdog` and `ClaudeCoworkBootFix` tasks and the old `cowork-watchdog.ps1`.
4. **Fix Claude Desktop shortcuts**: removes them from the Desktop and the Start Menu.
5. **Launchers**: removes the `LaunchClaudeAdmin` task, the `\Claude\` task folder once it holds no other task, the "Claude (Admin)" shortcut and its launcher scripts, the "Claude (no GPU)" launcher and shortcut, and the copy of Claude's icon. The RUNASADMIN compatibility flags on `Claude.exe` are left alone: no version of this toolkit set them, so any there are someone else's. Until 6.3.0 `-Undo` removed every one it found.

Every removal in steps 3 to 5 is checked, and one that did not happen fails the step. Until 6.3.0 a task the scheduler refused to delete was reported as not found, and a process that did not stop and a file that was not deleted as removed. A health monitor whose stop cannot be read fails step 3 too; until 6.3.1 a process that could not be read counted as gone. A file in steps 3 to 5 that cannot be read is a warning that names it, and is kept, and a removal whose result cannot be read fails the step; until 6.3.1 such a file was taken for none, so step 4 could say `No shortcuts found` while one was left.

`-Undo` ends with UNDO COMPLETE only when it is done. While a setting is still in the record it ends with UNDO INCOMPLETE, lists what is left and exits 1. When it cannot read what it needs to check, the NAT rules, the network adapters, the Hyper-V VMs, the running processes, a service, the files and the task that show an older version ran, or a file it would remove, it warns, ends UNDO INCOMPLETE and exits 1 too, and the next `-Undo` looks again. A recorded recovery whose service cannot be read stays in the record. A run in which a step failed ends with UNDO FINISHED WITH ERRORS and exits 1.

Until 6.3.0 `-Undo` set fixed values whatever the machine had before: hibernate and Fast Startup on, sleep after 30 minutes, every Ultimate Performance plan deleted, a plan you made yourself included, and the power plan it had recorded, which a second run had overwritten with its own. A setting that cannot be put back stays in the record: rerun `-Undo` once the cause it shows is resolved. A record that cannot be read, is not a version 1 record, or holds an item with no Kind or Key stops both Prevent and `-Undo` before anything changes. Until 6.3.1 a record whose file could not be read at all was taken for none, and `-Undo` said there was nothing to undo. Move `%LOCALAPPDATA%\ClaudeFix\prevent-undo.json` aside and rerun, as the message says; `-Undo` then cannot put back what that record held.

---

## Requirements

- Windows 10 (build 19041+) or Windows 11, any edition including Home. See the Compatibility note above for what is narrower on Home
- Claude Desktop with Cowork mode
- The Virtual Machine Platform Windows feature, which is what Cowork actually requires. It is not the same as the Hyper-V role, and a `*Hyper-V*` search will not find it, because the feature name contains no "Hyper-V". Check it with `Get-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform`
- PowerShell 5.1+ (included with Windows)
- Admin privileges: required for Prevent, optional for Fix

---

## Troubleshooting

**The fix script says "Workspace ready" but Cowork still shows an error**
Try running the fix script a second time. Some stale states need two cycles to fully clear. If it persists after 2-3 runs, restart Windows.

**"Failed to start Claude's workspace" with HCS error**
This is an HCS (Host Compute Service) failure, not a VirtioFS mount failure. Run `Fix-ClaudeDesktop.bat` as admin. Step 3 will detect the HCS error and restart the `vmcompute` service. If it keeps happening, run `Prevent-ClaudeIssues.bat` to configure automatic HCS service recovery. The health monitor (installed by Prevent) also watches for `vmcompute` handle leaks that precede these failures.

**"VM is already running" after running the fix script**
The old VM wasn't fully released before the service restarted. Run the fix script again. Step 5 now explicitly kills orphan compute systems via hcsdiag. If it persists, open an admin PowerShell and run `hcsdiag list` to see all compute systems, then `hcsdiag kill <id>` for any claude/cowork entries.

**"Service not found, is Cowork installed?"**
The `CoworkVMService` Windows service is only installed when you've used Cowork mode at least once. Open Claude Desktop, start a Cowork session, let it install the VM components, then run the scripts.

**The prevention script asks for admin but I don't have it**
The prevention script genuinely needs admin for power settings, scheduled tasks, and registry changes. The fix script can run without admin, but the prevention script cannot. Ask your IT department for temporary elevation, or have them run `Prevent-ClaudeIssues.bat` for you.

**Modern Standby changes didn't take effect**
This setting requires a full reboot (not just sleep/wake). Shut down, wait 10 seconds, power on. Verify with `powercfg /a`: if "Standby (S0 Low Power Idle)" is no longer among the sleep states it lists as available (it may still be listed among the ones that are not), Modern Standby is disabled. Versions before 6.3.0 set `CsEnabled`, which current Windows no longer reads: run `Prevent-ClaudeIssues.bat` again, and it sets `PlatformAoAcOverride` instead.

**The prevention script lists settings as "Left in place"**
Those are settings a version before 6.3.0 made and this version no longer makes: `LocalAccountTokenFilterPolicy`, `FilterAdministratorToken`, `ServicesPipeTimeout`, `CsEnabled`, the restart recovery on `CoworkVMService`, the `CoworkNAT` NAT rule, `PnPCapabilities` with the wake settings on an adapter, and Dynamic Memory left off on a Hyper-V VM named like Claude. Prevent leaves them as they are; `-Undo` removes them, and turns Dynamic Memory back on once that VM is turned off. The first three are listed this way only where an older install is found. Elsewhere the line says that the value matches what older versions set but no older install was found here, so `-Undo` leaves it: an administrator, another tool or a policy may have set it. When whether an older version ran here could not be told, because a file or the boot-fix task it would have left could not be read, the step warns, and the line says that it is not known whether `-Undo` removes the value.

**Antivirus warnings**
If you have third-party antivirus, add the recommended exclusion paths shown by the prevention script. For Windows Defender, exclusions are added automatically.

**Claude launches with a duplicate taskbar icon**
When using the "Claude (Admin)" shortcut, MSIX installs will show a second Claude icon on the taskbar. This is expected: Windows enforces medium integrity for MSIX apps launched through the shell protocol (`shell:AppsFolder`), so the only way to get a full admin token is to launch the `.exe` directly, which bypasses icon grouping. Close the non-admin Claude first if you want a single icon. The regular fix script (`Fix-ClaudeDesktop`) uses the shell protocol and does not have this issue.

**"Request timed out: isGuestConnected"**
The VM booted but the guest OS did not connect back to the host. Run `Fix-ClaudeDesktop.bat`, which stops Claude and `CoworkVMService` and relaunches Claude. Its Step 9 has a targeted recovery for this failure (HCS cleanup and a service restart), but it finds the failure in `cowork-service.log`, which current builds no longer write, so on current builds that recovery does not run. If it recurs frequently, the Authenticode verification loop in `cowork-svc.exe` may be starving the service of CPU. See issue #31848. The health monitor (installed by Prevent) starts a repair when the app logs this at `[error]` in its VM or main log.

**Logs are piling up in `%APPDATA%\Claude\fix-logs\`**
The fix script auto-cleans logs older than 30 days. When some cannot be deleted, it says so once its log starts. You can safely delete everything in that folder manually if needed.

**Claude vanishes the moment the agent uses the in-app Browser**
That is the GPU-process crash in [#80444](https://github.com/anthropics/claude-code/issues/80444). Check `main.log` in the app's log folder (`%LOCALAPPDATA%\Claude\logs` on app 6.18355; see [Where the app keeps its files](#where-the-app-keeps-its-files)) for `GPU process gone` with `exitCode: 101457950`, and `unknown-window.log` for a burst of `WebGL: INVALID_ENUM: getInternalformatParameter` warnings a second or two earlier. Run `Prevent-ClaudeIssues.bat` and launch from the "Claude (no GPU)" shortcut it creates. If that does not hold, avoid the in-app Browser pane on this build and use a real browser window for testing.

**A launcher says its script "could not be looked for"**
The `.bat` launcher could not read the folder it is in, so whether its script is there cannot be told. Move the toolkit to a folder your account can read, such as `C:\ClaudeFix`, and run it from there. A folder that may be read but not listed is fine, and so is a path with a parenthesis in it, such as `ClaudeFix (1)`; until 6.3.1 `Fix-ClaudeDesktop.bat`, `Stop-ClaudeDesktop.bat` and `Watch-ClaudeHealth.bat` stopped there with "... was unexpected at this time."

**The script says Virtual Machine Platform is disabled, but Hyper-V is on**
Those are different Windows features and both statements can be true at once. Run `Get-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform` to check the one that matters; a `*Hyper-V*` search cannot return it. That said, from v6.0.3 the script only names this feature as a cause when `vmcompute` and `hns` are genuinely missing, because full Hyper-V supplies the same underlying stack and Cowork runs perfectly well on it. If an earlier version told you to enable a feature and restart, and Cowork was working before that, you did not need to.

---

## Stop-ClaudeDesktop

Clean shutdown without repair. Double-click `Stop-ClaudeDesktop.bat` when you want to fully close Claude Desktop without triggering a fix cycle.

### What It Does

1. Stops `CoworkVMService` gracefully (waits up to 45s for the VM to shut down)
2. Kills the Claude Desktop app's processes, leaving Claude Code CLI processes running
3. Cleans any remaining HCS compute systems, counting one as removed only once it is gone from a second list, and saying when some were listed and not killed
4. Restarts the service in idle state (ready for next launch)
5. Verifies clean state

The service is restarted after stopping because Windows ignores ETW named pipe triggers for manually-stopped services. Without the restart, opening Claude Desktop would fail to auto-start `CoworkVMService`.

It says in green that Claude Desktop is shut down only when the check in step 5 found no app process and hcsdiag listed no Cowork compute system. Otherwise it says in yellow that the app is not fully shut down, after the warnings that say why, or, when hcsdiag did not answer, that the app's processes are stopped but whether a compute system is left is not known; if the relaunch then fails, run Fix. When the app's processes could not be listed, it says that whether they are stopped is not known, and when the service's state could not be read at the end, it says `Clean shutdown complete (whether the service is running is not known)` and that whether the service is running is not known; in both cases, if the relaunch fails, run Fix (v6.3.1). Until 6.3.0 it said in green that the app was shut down in every case. When another Fix run is already going, it does nothing and exits 3 (see [Exit codes](#exit-codes)).

A read that fails is said, and nothing is done on it (v6.3.1): a `CoworkVMService` that cannot be read is not stopped, a stop it refused says `Asking ... to stop failed (reason)` and the run waits to see whether it stops, a process list that cannot be read stops nothing, and the count of processes stopped is followed by `[!] N could not be stopped` when some were not. The line about the service process says `No service process was running` when there was none, and `Killed service process` only when a kill went out.

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
5. The VM starts successfully on the second attempt (~8-10 seconds total)

**This does not affect functionality.** No data is lost, no settings are changed, and the workspace starts correctly. Earlier versions of this section credited the logon task that ran ClaudeFix's BootPrep mode with preventing the more serious `0x800707DE` construct failure. That task was retired in 6.3.0; Fix's Step 3 still restarts `vmcompute` when it finds an HCS error. The close-reopen is a separate Claude Desktop bug that cannot be fixed externally.

Relevant log pattern in `cowork_vm_node.log`, in the app's log folder (see [Where the app keeps its files](#where-the-app-keeps-its-files)):
```
[VM:start] Beginning startup, VM instance ID: ...
[VM:start] Beginning startup, VM instance ID: ...  (duplicate)
[error] VM boot failed: VM is already running
Auto-reinstalling workspace after startup failure
```

---

## File List

```
Fix-ClaudeDesktop.bat       Fix launcher (double-click when broken)
Fix-ClaudeDesktop.ps1       Fix script
Stop-ClaudeDesktop.bat      Clean shutdown launcher (double-click to stop Claude)
Prevent-ClaudeIssues.bat    Prevention launcher (run once)
Prevent-ClaudeIssues.ps1    Prevention script
Watch-ClaudeHealth.bat      Health monitor launcher (manual foreground mode)
Watch-ClaudeHealth.ps1      Health monitor (auto-detects and auto-fixes crashes)
README.md                   This file
LICENSE                     MIT licence
```

Current toolkit version: 6.3.1 (one version across all three scripts).

---

## Changelog

See [CHANGELOG.md](CHANGELOG.md).

Releases before 6.0.0 shipped under `UpdateN` and `HotfixN` tags while each
script carried its own version number. From 6.0.0 the tags are semver and the
three scripts share one version. The old tags stay where they are.

---

## Credits & Acknowledgements

This toolkit was built by combining community knowledge from multiple independent contributors who shared their findings and workarounds:

- **Jonas Kamsker** ([blog.kamsker.at](https://blog.kamsker.at/blog/cowork-windows-broken/)): Comprehensive diagnostics, DNS/NAT fix scripts, and VM state recovery techniques
- **Elliot Segler** ([elliotsegler.com](https://www.elliotsegler.com/fixing-claude-coworks-network-conflict-on-windows.html)): Network conflict resolution and HNS-based recovery for subnet collisions
- **@garabedjunior-dotcom** ([GitHub #29848](https://github.com/anthropics/claude-code/issues/29848)): Community troubleshooting scripts and MCP crash diagnosis
- **@Onimir89** ([GitHub: Restart_claude](https://github.com/Onimir89/Restart_claude)): independent VM restart script demonstrating the service-reset recovery pattern
- Everyone who reported and documented VirtioFS failures in the [claude-code issue tracker](https://github.com/anthropics/claude-code/issues)

Special thanks to the community contributors on GitHub issues #25206, #26554, #27576, #27801, #28890, #29045, #29587, #29848, and #31520 whose collective debugging narrowed down the root causes.

---

## Licence

MIT

---

*Out of frustration, built with [Claude Desktop](https://claude.ai), for [Anthropic](https://github.com/anthropics).*
