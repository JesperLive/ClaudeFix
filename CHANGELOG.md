# Changelog

Reconstructed from the 14 GitHub releases and their tags. Everything before
6.0.0 shipped under `UpdateN` and `HotfixN` tags while the three scripts each
carried their own independent version number, so the version column below is
the one the release title claimed, not something the tag encodes.

From 6.0.0 onward there is one version for the whole toolkit, and tags are
semver. The old tags stay where they are.

## [6.0.4] - 2026-07-29

### Added

- `Prevent-ClaudeIssues` writes a GPU-free launcher: `Launch-Claude-NoGPU.cmd`
  in the Claude folder, plus a "Claude (no GPU)" desktop shortcut.

  Current builds have a fatal GPU-process crash, upstream
  [#80444](https://github.com/anthropics/claude-code/issues/80444), exit code
  101457950 (0x060C201E). A page loaded in the in-app Browser preview runs a
  WebGL capability probe, the GPU process dies, and the whole app goes with it.
  Seen in a user's logs: browser preview created at 10:02:22, 57 probe warnings
  at 10:02:27, GPU gone at 10:02:28.

  Running with `--disable-gpu` avoids it, but an MSIX app launched normally
  never receives a command line, so the flag has nowhere to go. The launcher
  uses `Invoke-CommandInDesktopPackage`, which starts the executable inside the
  package identity, so the flag arrives and the install stays intact. On a
  traditional install it just passes the flag on the command line.

  It also closes any running instance first. Claude is single-instance, so
  launching a second copy over a running one only focuses the existing window
  and the flag is silently ignored. That is the step most hand-written versions
  of this miss.

  Opt-in. The normal shortcut is untouched and still uses the GPU. `-Undo`
  removes the launcher and the shortcut.

- A CI gate that builds the generated launcher for both install shapes and
  parses it. Script-writing-script code is never executed at build time, so a
  dropped escape would have shipped and only surfaced when a user
  double-clicked. It also fails on a hardcoded package family, and on step
  numbering that disagrees with the declared total.

### Changed

- The admin shortcut's icon lookup reads `Get-ClaudeEnvironment` instead of
  re-running `Get-AppxPackage` and then walking three hardcoded install paths.
  It predated the discovery layer and was missed when the rest of the script
  adopted it.

## [6.0.3] - 2026-07-29

The same user sent their logs, and they showed 6.0.2 had fixed the wrong thing.

6.0.2 stopped the script repeating an unfixable message. It did not question
whether the message was true. It was not.

### Fixed

- The prerequisite check could veto a start attempt, and did. `Restart-CoworkService`
  called `Test-CoworkServicePrereq`, and on a Blocked result returned false
  without ever calling `Start-Service`. "Service failed to start" was printed
  for a start that never happened.

  The reporter's own `cowork-service.log` settles it. Between 11:12:16, when the
  script stopped the service, and 11:29:18, when Claude Desktop started it
  again, the file contains nothing at all. Not one failed start. The 11:29 start
  took 293 ms and worked.

  Prerequisites are now advisory. The service is always asked to start, and the
  findings are read out only if it refuses. A check that can be wrong must never
  prevent the attempt that would prove it wrong.

- Virtual Machine Platform is no longer treated as fatal on its own. What Cowork
  needs is a working Host Compute Service stack. That feature is one way to get
  one; full Hyper-V is another. On the reporter's machine `VirtualMachinePlatform`
  reads Disabled while `vmcompute.dll` and `computecore.dll` load, HCN
  enumerates two networks, and a compute system was in state Running an hour
  before they ran the script. The feature is only cited now when the HCS
  services are genuinely absent.

- 6.0.2 skipped the workspace wait whenever the prerequisite flag was set. That
  was the same error in a new place, and on a machine like the reporter's it
  would have abandoned a workspace that was about to come up. The wait is now
  skipped only when the flag survives an actual failed start.

### Changed

- The support bundle enumerates the log folder instead of copying eight
  hardcoded filenames. The list was written from one machine; the first real
  export it met had `coworkd-user.log` and `vm-info.json`, neither of them on
  it, and `vm-info.json` was the most useful file in the export. It carries the
  VM bundle size, whether the download finished, and `isGuestConnected`.

  Files are ranked rather than treated equally, because enumeration turned up 40
  files here and 32 were per-server MCP logs. Cowork and VM logs take four
  fifths of the budget, everything else shares the rest.

### Added

- A CI gate that stubs the prerequisite check to report a blocker and asserts a
  start is attempted anyway, in all three shapes: check wrong and service fine,
  check clean, and service genuinely dead. Verified by restoring the veto and
  watching five assertions fail.

## [6.0.2] - 2026-07-29

> Corrected by 6.0.3. This entry says the script "correctly" found Virtual
> Machine Platform disabled. It did read the feature state correctly, but
> treating that as fatal was wrong, and 6.0.2 did not question it. Read this
> entry with 6.0.3 above.

From a user report. The repair script found that Virtual Machine Platform was
disabled, said so, and then spent the rest of the run acting as though it had
not.

### Fixed

- After diagnosing a blocker the user has to clear, the script kept going. It
  launched Claude, entered the workspace wait, found the service dead, restarted
  it, hit the same wall, and repeated the identical four-line remediation every
  five seconds until the user pressed Ctrl+C. The wait loop was discarding the
  return value of the restart, which is exactly the flag that says this cannot
  work. It now skips the wait entirely when the blocker is known, and caps
  restarts at three otherwise.
- The Virtual Machine Platform message said "Cowork runs in a Hyper-V VM". The
  reporter read that, searched their features for `*Hyper-V*`, saw all seven
  Hyper-V features enabled, and concluded the script was wrong. That search
  cannot match `VirtualMachinePlatform`, because the name contains no "Hyper-V".
  The message now says so outright and gives the exact command to check.
- Step 5 printed "No orphan compute systems via hcsdiag" on a run whose step 0
  had already reported hcsdiag unavailable. A count of zero meant both "hcsdiag
  said none" and "hcsdiag never answered", and the caller could not tell them
  apart. It can now.

### Added

- Diagnostic mode writes a support bundle: the Claude log folder, the Service
  Control Manager entries for the three services that have to be alive, and the
  Windows feature states listed by exact name. Built to a 4 MB budget shared
  across the files, then zipped, so it fits an attachment limit. Absences are
  reported rather than left as gaps, because no `cowork_vm_node.log` at all
  means the VM has never started on that machine.

  The event log matters here. When a service refuses to start there is almost
  nothing in Claude's own logs, because nothing ran to write them, and Windows
  is the only thing that recorded what happened.

## [6.0.1] - 2026-07-28

Follow-up to 6.0.0. Four defects, three of them in code 6.0.0 changed, plus the
launchers, which no release had ever looked at.

### Fixed

- `Prevent-ClaudeIssues.ps1 -WhatIf` applied every change for real. The script
  relaunches itself elevated, and that relaunch forwarded `-Undo` but not
  `-WhatIf`. A dry run therefore became a real run the moment UAC was accepted.
  The same relaunch also returned success immediately instead of waiting for
  the elevated process, so a failed run reported as a clean one.
- All four `.bat` launchers treated any non-zero exit as a launcher failure and
  asked the user to screenshot an error. 6.0.0 made the repair script exit 1
  deliberately when retries are exhausted, which turned a normal, already
  explained outcome into a second alarming prompt.
- The HCS cleanup step in `Prevent-ClaudeIssues.ps1` still used `hcsdiag close`,
  which is not a verb hcsdiag has, and still ran even while Claude was open, so
  "stale" could mean the VM the user was working in. 6.0.0 fixed this in the
  repair script and the health monitor and missed the third copy.
- The `hcsdiag list` parse now exists once, in the shared discovery region,
  instead of four times across three scripts. Each copy had its own version of
  the same two mistakes, and fixing one did nothing for the others.

### Changed

- The health monitor reports guest-connect as dormant, rather than active, when
  `cowork-service.log` is stale or missing. On builds that do not write that log
  the check cannot fire, and the startup banner used to claim it was armed.
- `Prevent-ClaudeIssues.ps1` uses discovered paths and names throughout rather
  than only in its constants block. The service name, service executable and
  Claude data folder are read from discovery at every site, including the boot
  task it generates, where the value is baked in at install time because that
  task runs standalone.

### Added

- Two CI gates. One runs each launcher with a stubbed exit code and asserts what
  the user sees. The other reads the source and asserts that any script which
  self-elevates forwards `-WhatIf` across the boundary, waits for the child, and
  propagates its exit code. The parser gate now also runs against the shared
  hcsdiag parser directly, including the double-count case.

## [6.0.0] - 2026-07-28

One version across all three scripts. The largest change is that the toolkit
now discovers where Claude is installed instead of assuming, and that several
repairs which reported success were not doing anything at all.

### Fixed, and these were not working before

- HCS cleanup did nothing. Every call used `hcsdiag close`, and hcsdiag has no
  `close` verb. Ten call sites in the repair script and one more in the health
  monitor logged success for an operation that never ran.
- The `hcsdiag list` parser could not match real output. It expected a GUID
  alone on a line; hcsdiag prints it mid-line alongside the name. Both copies
  of the parser, in two different scripts, always returned an empty list.
- The health monitor could never repair anything. It preferred a log file last
  written in March, so its session check reported an active session forever and
  every repair took the "blocked" branch.
- Session liveness was derived from log markers that do not exist in the
  current log format.
- The VHDX backup checked for 720 MB of free space before copying a file that
  is over 3 GB. The check passed, the copy ran out of room, the partial was
  deleted, and the purge then destroyed the original.
- The VHDX restore ran after the service had been restarted, by which point the
  service had recreated the directory tree, so the restore declined to
  overwrite it. A 3 GB backup was abandoned on every deep run, silently.
- The escalation restore wrote one directory level too shallow and walked the
  cache directories in the opposite order to the backup.
- A clock-drift check that has never executed: matching against an array never
  populates `$Matches`, so reading it threw and an empty catch hid it.
- A Hyper-V heartbeat monitor that has never fired, because `Get-VM` does not
  enumerate HCS compute systems. It is now documented as such rather than
  advertised in the startup banner.
- Idle-time detection broke after 24.9 days of uptime, because `TickCount` is
  signed and wraps negative. Every quiet run then decided the user was active
  and exited without doing anything.
- The temp-file cleanup never ran: reading `.Length` on a directory throws
  under StrictMode, and the surrounding empty catch swallowed it.

### Fixed, and these were actively harmful

- `-SkipLaunch` could purge 13 GB of cache and then fail, because the
  escalation path it fell into read variables that only exist in the branch it
  had skipped.
- A single smart run could purge the cache twice, the second time part way
  through the re-download, so the repair could never converge.
- The deep escalation had no administrator check, so a non-elevated run could
  destroy the cache and then be unable to restart the service.
- One of the three purge blocks had no `-WhatIf` gate at all, so a dry run
  deleted 13 GB.
- The health monitor's user-activity guard failed open. Any unexpected error
  meant "nobody is using this", and the repair proceeded while someone worked.
- The monitor called the repair script in-process, and the repair script
  self-elevates, so from an unattended task it either raised a consent dialog
  nobody could answer while blocking the monitor, or failed.
- A `New-NetNat` auto-repair could claim a NAT over WSL's or Docker's switch,
  with a hardcoded /24 and no confirmation.
- Session transcripts older than 7 days were deleted on every run, in every
  mode, while the help text stated that conversations were never touched. This
  is now `-PurgeSessions`, off by default.

### Changed

- Locations are discovered at run time: package, executable, service binary,
  log directory, VM cache and VHDX inventory. The dead `%ProgramData%\Claude`
  paths and two hardcoded `C:` drive letters are gone.
- The health monitor's cooldown and repair history persist across restarts, so
  a repair loop can no longer escape its own backoff by crashing.
- Log baselines are keyed on file identity rather than path, so Electron's log
  rotation can no longer make historical errors look new.
- Detection runs on signals verifiable on the current build. The eleven legacy
  error strings, none of which appears anywhere in 35 MB of current logs, are
  kept as a secondary set.
- The workspace wait is wall-clock and extends while the cache is still
  downloading, so a slow connection is no longer mistaken for a hung VM.
- Exit codes mean something: 0 for a clean run, 1 for an unhandled error or
  exhausted retries.

### Added

- `-PurgeSessions` for session transcript cleanup, off by default.
- A storage check for whether the VM cache is readable and its volume has room.
- CI on every push: syntax, ASCII-only, undefined variables, version
  agreement, shared-region drift, backup purge logic, and PSScriptAnalyzer.
- A bug report template that collects the diagnostics these issues always need.

### Removed

- A brute-force filesystem scan for the executable that could not reach a
  WindowsApps install and whose result was rejected by the launch path anyway.
- An MSIX recovery path that looked for `resources\app\...` when the real
  layout is `app\resources\`, for a file that does not ship in the package.
- A session-file counter that called 10,433 files critical on every poll, held
  a repair trigger permanently open, and cost 1.8 seconds of disk enumeration
  every 30 seconds.

## Earlier releases

These shipped before the toolkit had a single version. Titles are as published.

| Tag | Date | Release |
|---|---|---|
| Update11 | 2026-07-27 | ClaudeFix v5.4.0: missing AppData folder fix + service prerequisite diagnostics |
| Update10 | 2026-03-23 | ClaudeFix Hotfix |
| Update9 | 2026-03-22 | v4.8.6 Audit cleanup, dead code removal, version sync |
| Hotfix3 | 2026-03-09 | ClaudeFix v4.8.5 Power Plan Fix + Boot Prep + README Overhaul |
| Update8 | 2026-03-09 | ClaudeFix v4.8.4 Non-Destructive Boot Prep + README Overhaul |
| Hotfix2 | 2026-03-09 | ClaudeFix v4.8.0: HCS hardening, robustness fixes, encoding fix |
| Update7 | 2026-03-09 | ClaudeFix v4.8.0: HCS hardening, robustness fixes, README update |
| Update6 | 2026-03-08 | v4.7.0: Interactive menu, smart VHDX backup/restore, HvHost fallback, vmwp kill, WSL2 detection |
| Update5 | 2026-03-08 | v4.6.0: Fix race conditions and add persistent failure escalation |
| Update4 | 2026-03-08 | Update 4 |
| Update3 | 2026-03-07 | Quick Hotfix and additional information in README |
| Update2 | 2026-03-07 | ClaudeFix 07/03/2026 |
| Hotfix | 2026-03-06 | ClaudeFix Hotfix |
| Release | 2026-03-05 | ClaudeFix Release |
