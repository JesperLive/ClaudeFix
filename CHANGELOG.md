# Changelog

Reconstructed from the 14 GitHub releases and their tags. Everything before
6.0.0 shipped under `UpdateN` and `HotfixN` tags while the three scripts each
carried their own independent version number, so the version column below is
the one the release title claimed, not something the tag encodes.

From 6.0.0 onward there is one version for the whole toolkit, and tags are
semver. The old tags stay where they are.

## [6.3.3] - 2026-11-23

Prevent's step 6, network adapter power saving, no longer warns on every run
that "Allow the computer to turn off this device to save power" still reads
on, on an adapter where it is off. Run `Prevent-ClaudeIssues.bat` again; this
fix needs no restart.

### Fixed

- "Allow the computer to turn off this device to save power" is read the way
  Device Manager reads it.

  Step 6 read `MSPower_DeviceEnable` for the whole class and took each
  adapter's setting from that list. On a driver built on NetAdapterCx, such
  as Realtek's 2.5GbE `rt25cx21`, NDIS answers that list, while the driver
  framework, which holds the setting, answers a read of one instance.
  Measured on the PC this was written for, after 6.3.2 had cleared the box
  and Windows had restarted: the list read `True`, the instance `False`, and
  Device Manager showed the box cleared. So every run recorded the box as on,
  cleared it again, and warned that it "still reads on until Windows
  restarts", which no restart changed. Each instance is now read by itself
  (`Get-CimInstance -InputObject`), both to decide and to check, on physical,
  hidden and virtual adapters alike. An instance that cannot be read that
  way is named in a warning and not written. A box that really does not take
  now warns that it "still reads on after it was cleared", and promises no
  restart. `-Undo`, turning "Allow this device to wake the computer" back on
  for adapters a version before 6.3.0 touched, reads each instance the same
  way.

## [6.3.2] - 2026-11-23

Prevent's step 6, network adapter power saving, no longer fails on the
Realtek drivers it was written for. Run `Prevent-ClaudeIssues.bat` again,
then restart Windows.

### Fixed

- The hidden Realtek options are turned off.

  `PowerDownPll` (Ethernet) and `bFwCtrlLPS` (Wi-Fi) failed with "Cannot
  validate argument on parameter 'DisplayValue'. The argument is null or
  empty." Each option was piped to `Set-NetAdapterAdvancedProperty`, which
  binds the piped option's own DisplayValue and RegistryValue by name, and a
  hidden option has no display value, so the cmdlet refused it before it
  wrote anything. Every adapter write now names its option with
  `-InputObject`. The same binding made `-Undo`'s turning wake back on, for
  adapters a version before 6.3.0 touched, send the off value it read beside
  the `Enabled` it asked for; that write names its option the same way.

- Selective suspend on a driver that has no option for it.

  On the Realtek 2.5GbE driver, `rt25cx21`, `Set-NetAdapterPowerManagement
  -SelectiveSuspend Disabled` returns without an error and writes nothing:
  the driver has no `*SelectiveSuspend` option to write it to, and Windows
  goes on reading it as `Enabled`. Step 6 failed with "SelectiveSuspend still
  reads Enabled". Such a driver is powered down while idle by the driver
  framework's own switch, `IdleInWorkingState`, which step 6 already turns
  off, and its selective suspend reading does not follow that switch. A
  setting that does not take on a driver with no option for it is now left
  as it is, with nothing recorded for `-Undo` and a record an earlier run
  made of it dropped, and the step warns that the driver has no option for
  it; for selective suspend it says nothing where `IdleInWorkingState` is
  off. A driver that has the option and does not take the setting still
  fails the step.

### Tests

- The test stand-ins for `Set-NetAdapterAdvancedProperty`,
  `Reset-NetAdapterAdvancedProperty` and `Set-NetAdapterPowerManagement` now
  bind their parameters as Windows' own do, measured on PowerShell 5.1 from
  the cmdlets' own parameter blocks: a hidden option piped in fails on its
  empty DisplayValue, and a visible one brings its RegistryValue. Visible
  options in the test adapters carry a display value, as Windows lists them.
  A check fails if any adapter write in Prevent is piped or does not name
  its option with `-InputObject`. Seven cases cover selective suspend and
  sleep on disconnect on a driver with no option for them. 850 cases in
  Prevent's steps, up from 839; run against 6.3.1, 28 fail.

## [6.3.1] - 2026-11-23

A read that fails is no longer taken as an answer. Many of the toolkit's reads
of a service, the running processes, a scheduled task, a folder or a log
passed over an error, so a read that failed looked like "not installed", "not
running", "not found" or "nothing recent", and the scripts acted on that. Each
of those reads now tells "not there" from "could not be read" by the error's
identity, and says what it could not read and why. Where the answer decides
whether something is stopped, killed, restarted or deleted, or whether a
repair may run, a read that failed changes nothing and holds the repair back.

The Claude (Admin) and Claude (no GPU) launchers are written by Prevent, so
run `Prevent-ClaudeIssues.bat` again to get the launcher fixes below.

### Fixed

- Finding the app says what it could not read.

  Fix, Prevent and the health monitor first look up where the app and its
  data live. An `hcsdiag.exe` that could not be read was called not present,
  a VM log folder that could not be read was passed over as one with no VM
  log, and when `Claude.exe` was not located, a place that could not be read
  went unsaid, so not located read as not installed. Each is now a read
  error, naming the path and why, shown with the other reads this lookup
  could not make: at the top of Fix's log, as `READ ERROR at start` in the
  health monitor, and as step 0 in Prevent. A VM cache folder that could not
  be read makes the cache size in `fix-logs\claude-env.json` a lower bound,
  and says so, naming the folder. A folder named `claude.exe` is no longer
  taken for the app, and the lookup no longer fails when `LOCALAPPDATA` is
  not set.

- Prevent and `-Undo` say what they could not read about this PC.

  Before step 1, Prevent and `-Undo` look up where the app and its data live.
  A read they could not make there went unsaid. Each one is now printed
  before step 1, and the summary lists them first, as step 0,
  "Looking at this PC", with the outcome Warning. They do not change the exit
  code.

- Prevent and `-Undo` no longer take a service, process, scheduled task or
  file they could not read as one that is not there.

  Each is now reported as unreadable, with the reason, and the step changes
  nothing on it. `-Undo` keeps what it could not check, such as a service's
  recorded recovery settings, ends with UNDO INCOMPLETE and exits 1, so the
  next `-Undo` looks again. A health monitor whose stop cannot be checked
  fails `-Undo` step 3. In setup, step 8 does not raise a Cowork VM worker it
  cannot read, step 9 fails when it cannot read vmcompute, step 12 does not
  start a Windows Time service it cannot read and does not check the clock
  then, and step 13 warns that an antivirus that is not Defender may go
  unreported when it cannot list the running processes.

- Prevent stops, naming it, when it cannot read the Claude folder in
  `%APPDATA%` or the undo record.

  Setup took a Claude folder it could not read for one that is not there, and
  said it could not be created. Prevent and `-Undo` took an undo record they
  could not read for none, so `-Undo` said there was nothing to undo. Each
  now stops the run before anything changes, naming the folder or file and
  why.

- `-Undo` keeps a file it cannot read, and says so.

  The files older versions left, the Fix Claude Desktop shortcuts and the
  launcher files were taken for none when they could not be read, so `-Undo`
  could say "No shortcuts found" while one was left. Each is now a Warning
  naming it, the file is kept, and `-Undo` ends UNDO INCOMPLETE, so the next
  `-Undo` looks again. A removal whose result cannot be read fails the step.

  A file that cannot be removed is named in the step's outcome, where the
  error did not name it, and `-Undo` says `CsEnabled` was put back when only
  its backup could not be removed, where it said it could not put it back.

- A registry value whose key cannot be opened fails the step that reads it.

  It read as a value that is not there, so a step could record that there was
  none before it wrote one, and `-Undo` would then delete the value instead
  of putting it back. The step now fails, naming the value, the key and why.

- Step 10 leaves the cowork-vm instances alone when it cannot list the claude
  processes.

  Whether Claude is running is not known then, so nothing is removed, and the
  step warns.

- Steps 8 and 10 warn when they cannot read `hcsdiag.exe`.

  They said Skipped, "hcsdiag.exe is not available". The Warning now names it
  and why: step 8 raises no VM worker, since the Cowork VM's cannot be told
  from the others, and step 10 looks for no stale Cowork VM.

- Step 11 warns when it cannot read the VM cache folder, and still checks its
  drive.

  It read as no VM cache: whether the cache is on a network drive was not
  checked, and the step said "Storage location OK".

- Step 1 fails on power plans, or an older version's power-plan record, that
  it cannot read.

  Step 1 and `-Undo` step 2 took an older version's `power-plan-backup.txt`
  that could not be read for an empty one. They now fail and keep the record
  for the next run. Power plans that cannot be read from the registry were
  taken for none, so another Ultimate Performance copy was made. They now fail
  step 1, and no copy is made.

  When whether the record is there cannot be read, step 1 fails the same way
  while the plan you had is not yet recorded. Once it is, and in `-Undo`,
  that is a Warning, the record is kept, and `-Undo` ends UNDO INCOMPLETE.

- Step 14 reads wsl.exe's exit code.

  A `wsl.exe -l -v` that failed read as "No WSL2 distro running". A non-zero
  exit now counts as no distro only while neither WSLService nor LxssManager
  runs. While one of them does, the step warns and names the exit code and
  the first line wsl.exe printed. A WSL feature that cannot be read is a
  Warning too.

  A `wsl.exe` that cannot be read is a Warning naming it, where it read as
  "No WSL2 distro running".

- Step 15 says "Started it now" only when the monitor's task runs.

  The scheduler ignores a start while the task still reads as running, and
  the old monitor can take a moment to end. The step now waits up to 5
  seconds for the task to stop reading as running, starts it, and checks for
  up to 3 seconds that it runs. When it still reads as running, its state
  cannot be read before the start, the start fails, or it reads another state
  3 seconds after the start, the step warns that the monitor starts at the
  next logon. When the task is no longer there 3 seconds after the start, the
  step warns that it is gone, and when its state cannot be read then, that
  whether the monitor runs could not be checked. When it cannot read whether
  the old monitor stopped, it warns and starts nothing now.

- Steps 15 and 16 fail when a place they look for the toolkit's scripts
  cannot be read.

  They look beside Prevent, then in `C:\ClaudeFix`, the Desktop and
  Documents. A place that could not be read was passed over, so the health
  monitor could be installed, or the shortcuts made, from an older copy found
  later, and with none found the step said the script was not beside Prevent.
  The step now fails, naming the place and why, and installs nothing from
  elsewhere.

- Steps 15, 17 and 19 warn about an older version's file they cannot read.

  `cowork-watchdog.ps1`, `Launch-Claude-Admin.cmd`,
  `disable-dynamic-memory.flag` and `connected-standby-backup.txt` were
  passed over as gone when they could not be read. Each is now a Warning
  naming it and why, and the file is kept.

- Step 17 fails when the scheduler refuses the Claude (Admin) task.

  The step now fails with the scheduler's reason, and no shortcut is made. A
  Claude (Admin) shortcut an earlier run left on the Desktop, whose task is
  gone, is removed. An old task that cannot be read or removed is a Warning.

- Neither launcher starts Claude when the running processes cannot be read.

  The Claude (Admin) launcher shows "Claude was not started" with the reason,
  and exits 1. The Claude (no GPU) launcher says so and waits for Enter. It
  also no longer starts Claude again without the GPU while a copy it tried to
  close is still running, or when whether it closed cannot be read, and it
  says when it cannot check that the Claude it started runs.

  The Claude (Admin) launcher also shows "Claude was not started", with the
  reason, when `Claude.exe` at a place it looks, or the list of app packages,
  cannot be read, and starts no Claude from a later place. A folder named
  `Claude.exe` is not taken for it.

  The Claude (no GPU) launcher closes nothing and starts nothing while a
  running Claude process cannot be read, as one started as administrator or a
  Claude Code session, since a copy already running takes the launch and
  ignores the flag. It went ahead, and then said Claude ran without the GPU
  or had not started.

- Step 19 and `-Undo` step 2 say when they cannot tell whether an older
  version ran here.

  They tell by the files and the boot-fix task older versions left. When
  those cannot be read, the values `-Undo` removes only where an older version
  ran (ServicesPipeTimeout, LocalAccountTokenFilterPolicy and
  FilterAdministratorToken at the values older versions set) are left as they
  are, with a warning, and `-Undo` ends UNDO INCOMPLETE. Step 19 warns when it
  cannot read the boot-fix task `\Claude\ClaudeCoworkBootFix`, which is then
  neither checked nor removed, and when it cannot read CoworkVMService, whose
  restart recovery from older versions is then not checked.

- Fix says each read it would act on that fails, and no longer acts on it as
  on an empty one.

  Each one is now said, with the reason, mostly in a DarkYellow `[!] ...
  could not be read (reason)` or `could not be listed` line, and Fix does not
  take the action an empty read would have justified: it does not stop, kill,
  restart, purge or delete on it. A few reads that decide nothing stay
  silent, such as the cache size that only extends the Step 9 wait. The reads
  discovery could not make are logged at the top of the run, and old logs,
  old VHDX backups and a discovery record that could not be cleaned up or
  saved are said once logging starts.

  Step 3 says when it cannot read the Claude log folder, which was passed
  over as one with no HCS error. A taskbar flash that cannot be loaded no
  longer prints a red Add-Type error.

- Fix passes over a Claude.exe path it cannot read, and says so, and no
  longer takes a folder for it.

  A cached path, the App Paths key or a Start Menu folder that cannot be read
  is named and passed over, and the cached path is kept. When the cached
  path, the App Paths key or a Start Menu shortcut named a folder, Step 8
  started it, which opened the folder in Explorer instead of starting Claude.
  A folder is now passed over.

- `-Quiet` holds a repair back on a read it cannot make.

  The Claude Code sessions the app started, the Claude Code transcripts, the
  processes named claude, their CPU time, the VM logs and the time of the last
  keyboard or mouse input could each be passed over when they could not be
  read, so an unattended repair could close the app under a working session.
  Each one now holds the run back: Fix logs `BLOCKED: <what could not be
  read>, skipping automated fix` and exits 3.

  The reason says why each read failed: an input check that could not be
  loaded gives its error, and a transcript folder that could not be read is
  named.

- Fix stops nothing on a process list it cannot read.

  Step 0 leaves a stale cowork-vm for Step 5, and says so, when whether
  Claude is running cannot be read. Steps 1 and 4 and `-Close` say when
  Claude's processes, or cowork-svc's, could not be listed, and stop none of
  them. Step 1 and `-Close` count only the processes that stopped, with
  `[!] N could not be stopped (reason)` after it, and `-Close` no longer calls
  the shutdown complete when its process list could not be read. Step 5 stops
  no VM worker when it cannot list them, and says the check did not finish.

- Fix does not stop, force-kill or restart a service it cannot read.

  Steps 2 and 9, `-Close` and the service restart say why instead, and a stop
  that was refused says "Asking ... to stop failed (reason)" instead of
  nothing. Step 3 still restarts vmcompute on HCS errors, but when whether it
  came back cannot be read it escalates no further, and it does not restart a
  vmms or HvHost it cannot read. A vmms or HvHost restart that reported an
  error is said instead of being called restarted. Smart mode does not
  escalate to Deep when the service's state cannot be read. The wait for the
  service process to exit says when the process list could not be read, with
  the warning that the VHDX files may be locked.

- Step 9 does not escalate or retry on a VM log it cannot read, and the run
  then exits 1.

  A guest that says it is connected, with a VM log that could not be read,
  no longer counts as ready. When a VM log could not be read, the run ends
  with "Whether the workspace came up is not known: a VM log could not be
  read", escalates and retries nothing on it, and exits 1; for a service
  that cannot start (below), it ends with that cause instead. A
  cowork-service.log that cannot be read is said once per run. After the
  guest-connection timeout, "Service restarted" is said only when the start
  went out, and "Continuing to wait..." otherwise.

  The line that says a VM log could not be read names the log, or its folder
  when the error gives no path.

- Step 6, both escalation purges and the VHDX backup do not purge what they
  cannot read.

  A cache folder that cannot be read is reported as not purged; Step 6 called
  it not present, and the escalation purges said nothing. A delete that
  failed in part says "could not be removed in full" instead of "removed".
  The VHDX backup does not purge when it cannot read whether a VHDX is in the
  VM cache, and the free-space line says why the free space is unknown. With
  the cache kept, by `-KeepCache`, Quick mode, or Smart mode before it
  escalates, a cache folder that cannot be read is said.

  When the app's data folder is not known, a Step 6 that purges says the
  cache folders are not known; a Step 6 that keeps the cache, and the
  escalation purges, pass over them. Step 6, the escalation purges and the
  Diagnostic run stopped there with UNEXPECTED ERROR.

  A VHDX restore no longer writes a backup over a file it could not read, and
  says when a backup could not be read, and so was not restored.

- Step 8 tries the other launch methods when it cannot read the
  LaunchClaudeAdmin task.

  It said "task not found, run Prevent". It now says the task could not be
  read and falls back to the standard launch. When it finds no Start Menu
  shortcut and part of the Start Menu could not be searched, it says so.

- The HCS clean-up says when a kill could not be sent, and when it stopped on
  an error.

  Its non-critical lines now name the error.

- `-PurgeSessions` says what it could not read or delete.

  When the sessions folder cannot be read, it says so and purges nothing.
  Part of the folder that cannot be read is said, and the files there are not
  considered. It reports "Deleted N of M" with `[!] K could not be deleted`,
  and keeps session folders it could not read in full.

- `-Close` says what it did to the service process.

  It says "No service process was running" when there was none, and "Killed
  service process" only when a kill went out. When the service could not be
  read, it ends with "Clean shutdown complete (whether the service is running
  is not known)" and "Whether the service is running is not known: see the
  warning above. If the relaunch fails, run Fix."

- `-Diagnostics` and the final summary say "could not be read" instead of
  absent.

  The services, the vmcompute handle count, the HCS VMs, the session files,
  the log listing and the Windows features say "could not be read" or "could
  not be recorded", with the reason, where they reported a failed read as
  absent or not configured. Without admin, the Defender line says "not
  checked (needs admin to read the exclusions)". The final summary shows "VM
  Service: could not be read (reason)" and "Claude processes: could not be
  listed (reason)".

  The VM cache folders say "could not be read" where they said "Not present",
  and "Not known" when the app's data folder is not known. The support bundle
  names each log it could not read or copy, and says it is not in the bundle,
  where it was listed as skipped, and gives no "the Cowork VM may never have
  started" note for a VM log it could not read.

- `-WhatIf` on Fix no longer force-kills or starts the service process, or
  reports either as done.

- The health monitor says what it could not read at start, and a poll that
  cannot list Claude's processes does nothing.

  Each read discovery could not make is logged at start as `READ ERROR at
  start`, and a start that could not list Claude's processes is said on its
  own. A poll that cannot list Claude's processes checks nothing and does
  nothing. The monitor's new "could not be read" lines that end with "the
  same error is logged at most once an hour" are written at most once an
  hour for the same check and error.

- The health monitor's service check and staleness check do nothing on a read
  they cannot make.

  A CoworkVMService whose status could not be read counted as one that is not
  installed. When the service or Claude's processes cannot be read, the
  service check now says so and does nothing, and the count of failed checks
  stays as it was. When the VM logs cannot be read, the staleness check says
  so and does nothing.

- The health monitor says when it cannot read the Windows Time service, and
  does not start it then.

  The clock is still measured every 5 minutes. A Windows Time service that is
  not installed is still passed over without a word.

- After a forced clock resync, the hour with no further forced resync runs
  from when the resync finished, not from just before it.

- The health monitor does not kill a Cowork VM as stale while it cannot tell
  the live one.

  When the VM workers cannot be listed, or a worker's command line or a
  Cowork VM worker's start time cannot be read, no VM is killed, and this is
  said.

- A vmcompute process list or handle count that cannot be read is said,
  naming the error, and no longer resets the run of high handle counts.

- The restart of vmcompute after a spike of HCS failures is held when the
  Cowork service or Claude's processes cannot be read, and this is said.

- A Claude log folder that cannot be listed is said, and the log triggers
  keep their place until it can be listed.

  A folder that is not there is still silent.

- The health monitor's guards answer "active" on a read they cannot make.

  The Cowork guard takes a session to be active when the Cowork VM log cannot
  be read or read through, so nothing that would end a session is done. The
  Claude Code and user-activity guards answer "active" when a folder, the
  process list or a CPU time could not be read, and say why at most once an
  hour, where they were logged on every check.

- A repair is deferred or cancelled when a read just before it fails.

  A repair waits, saying why, when Claude's processes cannot be listed or
  the start time of one of them cannot be read: a running Claude process
  whose start time cannot be read is no longer taken as exited. The 30-second
  grace period before a repair cancels it when Claude's processes, their CPU
  time or the VM logs cannot be read.

  Where the CPU time of a `claude` process cannot be read, the reason now
  says too when whether it has ended could not be read, here and in Fix's
  `-Quiet` check.

- The health monitor's data-folder sync moves nothing while a candidate
  folder cannot be read.

  A candidate folder that disappears while it is listed is no longer reported
  as unreadable.

- The health monitor's stop line says the service's state is "not known"
  when it could not be read, rather than "absent".

- Watch logs over 30 days old that cannot be deleted are said, naming the
  folder and the error.

  This is checked at startup and at the daily rollover. A log folder that is
  not there is not an error.

- The health monitor says, after its startup banner, when the foreground
  window and input checks could not be loaded, naming the error.

- A background hcsdiag call that cannot run hcsdiag says why, at most once an
  hour, instead of reading as hcsdiag not answering.

- The health monitor holds automatic repairs for a cooldown when it cannot
  read its saved state.

  `watch-state.json` keeps the cooldown, the repairs of the last 30 minutes,
  which stop a fourth after three, and a held-back repair. A file that was
  there but could not be read or parsed read as no state at all, as on a
  first start, so a repair loop bad enough to leave the file broken could run
  again at once; an empty file, or one that could not be looked for, did so
  without a word. Now a cooldown starts, as after a repair, and no automatic
  repair runs until it is over. This is logged at most once an hour, as
  `STATE: The state saved in ... could not be restored in full`, naming the
  reason and how long no repair runs. A file that is not there is still a
  first start.

  A stealth-update marker written before the monitor started is then not
  acted on, since whether the monitor before it relaunched Claude for it is
  not known, and this is logged.

- The health monitor stops, naming the error, when it cannot read its Fix
  script.

  It looks for `Fix-ClaudeDesktop.ps1` beside itself, then in `C:\ClaudeFix`,
  the Desktop and Documents. A place that could not be read was passed over,
  and the monitor started with a copy found later. It now stops with `FATAL:
  Fix-ClaudeDesktop.ps1 at '...' could not be read`, naming the error.

- A stealth-update marker the health monitor cannot look for is logged as
  `STEALTH WARNING: marker at ... could not be read`, once while it lasts,
  where it read as no marker.

- The health monitor's storage check reports a VM cache folder it cannot
  read, naming it, where it read as no VM cache. Each different storage
  warning text, this one or a volume low on space, is logged at most once an
  hour, with no cap on how many, so a new warning, or one whose text changes
  to one not logged in that hour, is logged at once. The low-space warning
  was written on every poll, about 120 times an hour.

- The health monitor's `BLOCKED` and "still waiting" lines say why the Claude
  Code sessions the app started could not be listed, or why the transcript
  folder could not be read, naming it.

- The `.bat` launchers start their script from a folder whose path holds a
  parenthesis, or that may be read but not listed.

  `Fix-ClaudeDesktop.bat`, `Stop-ClaudeDesktop.bat` and
  `Watch-ClaudeHealth.bat` stopped with "... was unexpected at this time."
  before running anything when their folder's path held a parenthesis, as in
  "ClaudeFix (1)" or "Program Files (x86)". All four launchers said their
  script was "not found" in a folder that may be read but not listed. When a
  launcher cannot look for its script because its folder cannot be read, it
  now says "could not be looked for: this folder could not be read", and
  exits 1.

  `Watch-ClaudeHealth.bat`'s message after the monitor stops on startup, and
  its message when PowerShell could not run the script, now show when the
  profile path holds a parenthesis.

- The Quick Fix line in Fix's menu, in its `-Mode` help and in the README
  says what Quick mode runs.

  It said "Steps 1-5, skip cache purge". Quick mode runs every step but two:
  the cache purge in Step 6, and the HvHost restart in Step 3, which only
  Deep mode runs. The README's Smart mode escalation said the same of
  Smart's first pass, and left out its second escalation, when the
  workspace is not ready after Step 9.

- Fix in Smart mode no longer purges the VM cache with `-KeepCache` set
  (Keep cache [ON] in the menu).

  Both of Smart's escalations to Deep purge the VM cache, and neither looked
  at `-KeepCache`, so a Smart run set to keep the cache deleted the roughly
  13 GB bundle when it escalated. Neither escalation runs with `-KeepCache`
  now, and the run says `Smart mode: ... not escalating to Deep: -KeepCache
  keeps the VM cache`; the retry cycle after Step 9 still runs.

- Fix no longer escalates to Deep, or retries, after Step 9 on a service
  that cannot start.

  When the prerequisite check found that CoworkVMService cannot start, as
  when it is disabled, and it was down when Step 9's wait began, the wait
  was skipped, and a restart in the wait that failed on such a cause
  stopped the wait; either way Smart mode's escalation after it still
  purged the VM cache, and the retry cycle still restarted the service
  three times with a wait of up to a minute each. Neither can fix that,
  and neither runs then now: the run says
  `Workspace not ready, but not retrying:` and the cause (in Smart mode as
  administrator, before it has escalated to Deep,
  `not escalating to Deep or retrying:`), and exits 1, as it did after the
  retries. Step 9 no longer says then that the workspace was not confirmed
  ready, after 0s when the wait was skipped, nor to open a Cowork session
  to trigger setup.

## [6.3.0] - 2026-11-22

If you ran Prevent from an older version, run `Prevent-ClaudeIssues.bat`
again. Several settings older versions reported as made were never made. The
new run makes them, records what each one was before it changes it, and
removes the boot-fix task. Until it runs, a boot-fix task an older version
installed keeps running `Fix-ClaudeDesktop.ps1 -BootPrep` at every logon. In
this version `-BootPrep` only says that it is retired and exits, so the task
no longer restarts `vmcompute`, and only Prevent or `-Undo` removes the task.

`-Undo` from 6.3.0 still cleans up after older versions. Its first step puts
back what a 6.3.0 run recorded, and its second removes what versions before
6.3.0 are known to have set, so it is the undo to run whichever version made
the changes. What it cannot know, it leaves: hibernate and Fast Startup turned
off before 6.3.0, and Ultimate Performance plans, because no record says what
they were or which ones older versions made.

### Fixed

- `-Undo` puts back what the machine had, not fixed values.

  It set the same values everywhere: hibernate and Fast Startup on, sleep
  after 30 minutes, every Ultimate Performance plan deleted, a plan you made
  yourself included, and the power plan it had recorded. Every setting Prevent
  changes is now recorded before it is changed, with the value it had, in
  `%LOCALAPPDATA%\ClaudeFix\prevent-undo.json`, and `-Undo` puts those values
  back, newest first. The first record of a setting wins, so a rerun does not
  record the value Prevent set as yours. `-Undo` deletes only the Ultimate
  Performance plan Prevent created, and keeps that one while it is still the
  active plan.

  The record is in the local profile, not the roaming one, and it names the
  computer that wrote it, so a roaming profile cannot carry it to another PC
  and have that PC put back this one's values. A record written on another
  computer is never used: Prevent moves it aside, as
  `prevent-undo.from-<computer>.<time>.json` in the same folder, and starts
  this computer's own record, and `-Undo` puts nothing back from it and says
  why. A record that cannot be read, is not a version 1 record, or holds an
  item with no Kind or Key stops Prevent and `-Undo` before anything changes,
  rather than being overwritten with values Prevent set.

- `-Undo` says when it did not put everything back, and exits 1.

  It ended with UNDO COMPLETE whatever had happened. A setting it cannot put
  back now stays in the record for the next `-Undo`, as one for a network
  adapter that is not plugged in does, and the run ends with UNDO
  INCOMPLETE, lists what is still in the record and exits 1. When `-Undo`
  cannot read a list it needs to check (the NAT rules, the network adapters,
  the Hyper-V VMs or the running processes), it warns, ends UNDO INCOMPLETE
  and exits 1, and the next `-Undo` looks again. A run in which a step failed
  ends with UNDO FINISHED WITH ERRORS. A power setting Windows no longer has,
  and a driver option the adapter's driver no longer lists, are dropped from
  the record with a warning: there is nothing left to put back.

- `-Undo` checks what it removes, and leaves the RUNASADMIN flags alone.

  A scheduled task the scheduler refused to delete was reported as not
  found, a health monitor that did not stop as stopped, and a file that was
  not deleted as removed. Each removal in its last three steps is now
  checked, and one that did not happen fails the step. The `\Claude\` task
  folder is removed too, once it holds no other task. `-Undo` also removed
  every RUNASADMIN compatibility flag on `Claude.exe` it found, and now
  leaves them: no version of this toolkit set them, so any there are someone
  else's.

- A second Prevent run no longer replaces the record of your power plan.

  Every run wrote the active plan to its backup file, so a second run recorded
  Prevent's own Ultimate Performance plan as yours, and `-Undo` "restored"
  that. The plan you had is now recorded once, the first time. An older
  version's backup file is carried over and deleted, and when it holds an
  Ultimate Performance plan the run says that `-Undo` will return to Balanced.

- Prevent no longer makes another Ultimate Performance plan on every run on a
  non-English Windows.

  It found the plans by the words "Ultimate Performance" in `powercfg /list`,
  which shows the translated name, so on another language it found none and
  made one more copy. It now finds them by their FriendlyName, the resource
  reference every copy of the template carries.

- A change powercfg refused is no longer reported as made.

  powercfg's exit code was read only when an Ultimate Performance plan was
  created, and no change was checked: the active plan was read back and
  printed as verified, whichever plan it was. Each power setting is now read
  back after it is written, and the plan switch, hibernate and the plan
  deletion in `-Undo` are each checked the same way.

- Modern Standby is turned off where the machine has it.

  The step set `CsEnabled` to 0, which current Windows no longer reads, so on
  a machine with Modern Standby nothing changed and the summary said it was
  off. It now asks `powercfg /a` whether the machine uses Modern Standby and,
  where it does, sets `PlatformAoAcOverride` to 0, the switch Windows reads
  now. It takes effect at the next Windows restart.

- Network adapter power saving is turned off where Windows reads it, and the
  clock check after it can reach the time server.

  `PnPCapabilities` was written to the adapter's Enum key, where nothing reads
  it, so an adapter reported as done still had power saving on. It now goes in
  the adapter's driver key, and only bit 0x08 is added: the wake bits beside
  it are left as they are, and the key is never created. The adapters were
  also restarted, which dropped the network link for about 14 seconds, and the
  clock check later in the same run then failed. No adapter is restarted any
  more: every change is made with `-NoRestart` or is a WMI setting or a
  registry value, and takes effect at the next Windows restart, which the
  summary says.

- Only the Cowork VM's worker process is raised, and only its worker is
  killed.

  Prevent and the health monitor set every `vmwp.exe` on the machine to
  AboveNormal: every Hyper-V VM, WSL2 and Windows Sandbox. Fix's Step 5 killed
  every `vmwp.exe` on the machine, so one of those running beside Claude went
  down with it. The Cowork VM's worker is now the `vmwp.exe` whose first
  argument is the GUID of a Cowork compute system `hcsdiag list` showed, in
  Fix at any point in the run, and when no list showed one, no worker is
  touched. Fix stops it through hcsdiag, waits up to five seconds, and only
  then force-kills it, and never a process that has since taken over its
  PID. A worker counts as stopped only once it has exited, where any output
  from hcsdiag, an error included, counted as a kill.

- A single stale VM is removed, and a removal is reported only once it
  happened.

  With one stale VM, the case the removal exists for, the parse returned a
  single string, and `.Count` on a string throws under StrictMode in Windows
  PowerShell 5.1. Prevent logged "check failed" and removed nothing, and Fix
  reported none found. An `hcsdiag list` that failed read as an empty one, so
  Prevent said "HCS state clean". And every kill sent was counted: Fix printed
  "Killed N" and the monitor logged "killed" whether hcsdiag had removed
  anything or not.

  Prevent now checks the list's exit code, and in Prevent, Fix and the monitor
  a kill counts only when the compute system is gone from a second list. When
  the monitor finds more than one Cowork VM, it keeps the one whose worker
  process started last, compared in UTC, since in the hour the clocks go
  back a later start can show an earlier local time. It kept the last entry
  hcsdiag printed, on the belief that the list runs oldest to newest, which
  nothing documents. With no worker to tell the live one by, it kills none.

- An unattended repair no longer ends a Claude Code session the app started.

  Fix's `-Quiet` guard, which the health monitor's repairs go through, checked
  CPU, the VM log and input. A project thread waiting on a 25-minute command
  uses almost no CPU and writes nothing, so an unattended run passed all three
  and ended it. The guard now first holds the repair back while a Claude Code
  CLI the app started is running, and while a Claude Code transcript was
  written in the last 35 minutes, the same question the monitor asks before
  its 30-second warning. Both fail closed: when the parent of a Claude Code
  CLI process cannot be read, the repair waits. A CLI in a terminal is not
  one the app started.

  The guard asked once, at the start, and a Smart run takes minutes, long
  enough for the app to start a project thread. A `-Quiet` run now asks
  about the app's sessions again right before each step that stops the app:
  Step 1, Step 4 when it finds the app's leftover processes, the
  no-progress escalation in Step 9 and the escalation to Deep. Before Step 1
  the run is held back (exit 3). After it the run stops part-way (exit 4),
  and first starts CoworkVMService again when this run stopped it, since
  Windows does not start a manually stopped service on its trigger; when
  that start fails, it says so and asks you to run Fix when Cowork is
  needed. When the app's sessions cannot be listed, the run stops. Manual
  runs are not asked.

- Fix's exit code says when a run was held back or stopped part-way.

  A run held back before it stopped the app or the service exited 0, the
  same as a finished repair: another Fix run holding the lock, and the
  `-Quiet` guard finding Claude or the PC in use. The health monitor then
  logged it as a completed fix and counted it toward "Auto-fix failed 3
  times". Those runs now exit 3, and a `-Quiet` run that stops part-way,
  above, exits 4. 0 is still a finished run, and 1 a workspace that was not
  ready after the retries or an unexpected error. The README's new "Exit
  codes" section lists them. Once a run is past every check that holds it
  back, Fix logs "Repair under way", which the monitor reads to tell a Fix
  run that started the repair from one that stopped before it.

- The health monitor no longer counts a repair Fix held back, and tries a
  held-back repair again once the hold clears.

  A run Fix held back (exit 3) is logged as `HELD BACK` and is not counted
  toward "Auto-fix failed 3 times". A run that stopped part-way (exit 4) is
  logged as `STOPPED PART-WAY`, saying that a Claude Code session was found
  or could not be ruled out, and is counted like any other repair: toward
  the escalation, with the normal cooldown after it.

  A check that is made again on every poll, such as the service check, fires
  again while the fault lasts. A log error or an Event Log error fires once,
  so a repair for one that Fix held back, or that the monitor BLOCKED or
  DEFERRED, was lost. That repair is now kept pending in `watch-state.json`,
  as is one for such an error that fires during the wait a hold leaves, and
  it is tried again once the hold has cleared, with the 30-second warning at
  most 3 times. It ends after 4 hours, when Claude Desktop is started again,
  or when a Fix run logs that it started the repair. A Fix run that stopped
  before the repair, such as a Diagnostic run, a Cancel or a No at
  "Proceed?", does not end it, and while a Fix run holds Fix's lock the
  repair waits for it to end. A log error or an Event Log error that fires
  while a repair is pending takes it over: the repair is tried for the new
  reason, its 4 hours start again from then, and its tries are kept. In the
  cooldown, the COOLDOWN line for it is followed by one that says it takes
  over the held-back repair, naming the reason that was pending, and is tried
  again once the wait is over, for up to 4 hours. The error the repair is
  already pending for, firing again, takes nothing over: its 4 hours start
  again, and that line says only that it is tried again.

  Each run Fix holds back in a row doubles the wait before the next repair:
  2, 4, 8, 16, 32, then 60 minutes. The count starts again after a run Fix
  does not hold back, after a hold 4 hours or more after the one before, and
  after a Fix run by hand that started the repair, which also ends the wait
  5 minutes after that run at the latest. A trigger in the cooldown after a
  repair that ran, an escalation or a cancel is skipped, as before.

  A repair you cancel by switching to Claude during the 30-second warning is
  let through again about 2 minutes after the cancel. That was timed from
  before the warning, so it came back a minute or so after the cancel.

- The health monitor no longer repeats its BLOCKED, COOLDOWN and SKIPPED
  lines on every poll, and COOLDOWN says when a fix is allowed again.

  A trigger that is checked again came back to them each time it fired
  while a hold lasted, which wrote the same lines about 30 times an hour,
  and "AUTO-FIX SKIPPED: Fix-ClaudeDesktop is already running" up to 120
  times an hour while a Fix run by hand waited in its menu. BLOCKED and
  COOLDOWN lines for the same reason are now written at most once every 10
  minutes, and so is SKIPPED when no repair is pending; with one pending,
  SKIPPED is written once for each Fix run that holds Fix's lock. A pending
  repair that is still waiting says why, at most once every 10 minutes.
  COOLDOWN said "last fix N min ago", counted from a time that BLOCKED,
  DEFERRED, CANCELLED and HELD BACK move back with no fix run, so after a
  BLOCKED it named a fix that never ran. It now says "a fix is allowed again
  in N min".

- `Fix-ClaudeDesktop.bat` and `Stop-ClaudeDesktop.bat` say what exit codes 3
  and 4 mean.

  For 3 they say that the run was held back before the app or the service
  was stopped, and for 4 that the repair ran part of the way and stopped
  because a Claude Code session was found or could not be ruled out, and
  that the lines above say which. Neither pauses for them, as neither does
  for 0 and 1.
  Any other code is still reported as PowerShell failing to run the script.
  A manual run that finds another Fix run already going now waits for a key
  before its window closes, where it closed at once.

- An hcsdiag call that fails no longer reads as an answer, and an empty list
  no longer reads as a failure.

  Fix and the health monitor threw hcsdiag's exit code away, so an error it
  printed came back as if it were the list. A call that exits with an error
  code now counts as no answer, and its exit code and first line are logged
  once, and again only when what hcsdiag says changes or after it has
  answered in between, so a host where it keeps failing no longer gets the
  same line on every check. `hcsdiag list` prints nothing when no compute
  system is left, and that read as no answer: Fix's Step 0 said hcsdiag was
  unavailable. It now counts as nothing listed. A compute system still in
  the list after a kill is reported as still listed after the kill, since it
  can take longer to leave the list than the check waits. While `-Close`
  waits for the service to stop, an hcsdiag that does not answer is reported
  as such, where it read as the VM gone.

- Fix's Step 5 says "No orphan compute systems found" only when it looked
  and found none.

  It said so whenever nothing was killed: after a worker it could not stop,
  when hcsdiag had not answered or the worker check failed, and when compute
  systems were listed but not killed. It now says "Nothing was stopped, but"
  and what was not checked, the compute systems or the VM worker, and how
  many compute systems were listed and not killed. `-Close` says the same of
  the ones it listed.

- `-Close` and `Stop-ClaudeDesktop.bat` say Claude Desktop is shut down only
  when it is.

  The closing lines said in green that Claude Desktop was shut down after a
  warning that Claude processes or a compute system were still there, and
  when hcsdiag had not answered. That is now said only when the check found
  no app process and hcsdiag listed no Cowork compute system. Otherwise it
  says, in yellow, that the app is not fully shut down, or that its
  processes are stopped but whether a compute system is left is not known,
  and to run Fix if the relaunch fails. The log no longer says "Clean
  shutdown complete" when hcsdiag did not answer.

- Step 9 notes a VM that may be stuck once, and only while nothing happens.

  "HCS VM present but no log/guest activity, VM may be stuck" was logged on
  every pass after 60 seconds whenever hcsdiag listed the VM, so a slow boot
  that was making progress filled the log with it. It is now logged once
  for each stretch with no log or guest activity.

- Ages are measured in UTC, so the change to or from summer time no longer
  moves them by an hour.

  Fix, the health monitor and the discovery code they share measured file
  and event ages against local time, so at the change to summer time a file
  written a minute ago read as an hour and a minute old. That moved the
  activity checks, the log ages Steps 3 and 9 read, the shutdown failures
  `-Mode Diagnostic` counts for the last hour, and when old logs, VHDX
  backups and, with `-PurgeSessions`, session transcripts are deleted. The
  monitor's cooldown, the escalation's 30-minute window, its startup grace
  period and how long ago Claude started are in UTC too, and so are the
  times in `watch-state.json`; one an older version wrote is converted when
  it is read. Each Event Log read in Fix and the monitor now gives the start
  of its window in UTC, so the window is right in the hour the clocks go
  back, which a local time names twice. The times in the app's own log
  lines, which it writes in local time with no offset, are still compared
  with local time.

- The storage check no longer reads a folder named Xbox as a Box sync folder.

  It was a substring test for "OneDrive", "Google Drive", "Dropbox", "iCloud"
  and "Box" anywhere in the path. A sync folder is now a OneDrive folder
  Windows names in the environment, or a folder in the path that is named
  after the service or starts with that name and then " - " or " (", as
  "OneDrive - Contoso" and "Dropbox (Personal)" do. Below the user profile
  only the part after it is tested, so a user named Box is not a Box folder.

- Prevent's storage check says when a check could not be made, and finds a
  mapped network drive.

  Every read in step 11 dropped its errors, so a check that failed read as
  a local drive and the step said the storage was OK. A check that cannot be
  made, of which drive holds the Claude data folder, whether that drive is
  external or a network drive, or whether the VM cache is on a network
  drive, is now a warning that names it and why, and "Storage location OK"
  is said only when every check was made. A data folder on a network share
  with no drive letter gets a warning that neither check was made. The
  network drive check ran only for a drive `Get-Volume` lists, which a
  mapped network drive is not, so it never found one; it now runs for any
  drive letter.

- The clock check measures on every locale and reports a resync that failed.

  The offset was read with a pattern that took a decimal point only, so where
  the locale writes a comma neither Prevent nor the monitor measured anything,
  and a resync that failed was reported as done. Both now read either
  separator and the resync's exit code, and a failed resync is a warning.
  Prevent also starts the Windows Time service if it is stopped, and waits up
  to 30 seconds for `time.windows.com` to resolve before it asks. When it
  cannot be reached in that time, the step warns that the clock was not
  checked.

- The health monitor logs a clock resync as repaired only when the clock is
  right after it, and stops forcing one that does not help.

  A resync that exited 0 was logged REPAIRED without measuring again, and
  one that changed nothing was forced again every 5 minutes for as long as
  the drift lasted. The clock is now measured again after the resync, and
  REPAIRED, with the drift before and after, means it is within 5 seconds.
  Otherwise the warning gives the drift, why the resync did not help and the
  time source Windows syncs from, and no resync is forced for 60 minutes
  while the drift lasts. The clock is still measured every 5 minutes, the
  warning is repeated while the drift is over 30 seconds, and the monitor
  logs when it is back within 5 seconds. A wait that reaches more than 60
  minutes ahead, because the clock was set back during it, is ended. The
  resync and the time source query now have the 15 seconds the measurement
  had, where the resync could stall the monitor for as long as the network
  took. A resync, a measurement after it or a time source query that fails
  is named in the warning, and the drift is still reported.

  A clock check that failed was passed over without a word, so it looked
  like a clock that was right. When w32tm cannot be started, or answers with
  an error and no offset, as when it cannot reach the time server, a TIME
  WARNING now says that the clock check failed and names why. A measurement
  that does not answer within 15 seconds is logged as a skipped check, a
  line that was written on every check. Each is now logged at most once an
  hour for the same error.

- The health monitor says when one of its checks fails, and Fix's HCS error
  check says when it cannot read a log.

  Most of the monitor's checks passed a failure over without a word, so a
  check that could not run looked like one that found nothing wrong. Each
  now logs one line that names the check and the error: the Claude logs,
  the four Event Logs, storage, VM log staleness, vmcompute, the HCS checks,
  the VM worker processes, saving and restoring `watch-state.json`, the Fix
  logs and Fix's lock, the lock that keeps a second monitor from running,
  the session it runs in, the focus check before a repair, and the poll
  itself, as `MONITOR ERROR`, which was written on every poll. The same
  error is logged at most once an hour, and no more than 3 different errors
  of one check in an hour. A vmcompute handle count that cannot be read is
  said, where it was taken as 0. When the VM worker processes cannot be
  listed, the monitor raises no worker and kills no VM. A Fix lock the
  monitor may not open, as when Fix runs elevated and the monitor does not,
  means Fix is running, where any error read as not running, and so does
  one it cannot ask about at all. When the check that the monitor is
  elevated fails, the line that skips the repair names the error, where it
  said the monitor was not elevated and to reinstall its task.

  Fix's Step 3 and `-Mode Diagnostic` say when the Hyper-V Compute event
  logs, a Claude log or the Claude log folder cannot be read, naming the
  error, and go on to the next Claude log, where each read as a log with
  no HCS error. Fix's summary and `-Mode Diagnostic` also say when the
  Application log, where they look for CoworkVMService errors, cannot be
  read, where that read as no errors. A support bundle with no Service
  Control Manager event for these services in the last 7 days reports 0
  events and says so in `service-events.txt`, where it logged that the
  System log could not be read.

- `-Undo` removes the Defender exclusions Prevent added.

  It left them all in place. Each exclusion is now recorded before it is
  added, so `-Undo` removes the ones Prevent added and no others. One that
  Defender refuses is taken back out of the record and fails the step. A
  Defender whose settings cannot be read fails the step too, where it read as
  nothing excluded yet. Where Defender hides its exclusion lists from
  administrators, the step says that it cannot tell which exclusions exist
  and adds none.

- Prevent's antivirus check says when the antivirus products Windows knows
  of cannot be read.

  The list Windows Security Center keeps was read with its errors dropped,
  so a list that could not be read meant no antivirus product registered,
  and one that is not Defender could go unreported. That is now a warning,
  and when no known antivirus process is running either, the step says "No
  antivirus product found among the running processes" instead of "No
  antivirus product detected". Where Windows has no such list, as Windows
  Server has no Security Center, nothing is said, as before.

- The Claude (Admin) and Claude (no GPU) shortcuts keep their icon across a
  Claude update.

  Their icon pointed at `Claude.exe` itself, which for an MSIX install sits in
  a folder named after the version, and the first update removed that folder.
  Prevent now copies Claude's icon to `Claude.ico` in `%APPDATA%\Claude`, uses
  a stock Windows icon when there is no Claude to copy it from, and `-Undo`
  removes the copy.

- The health monitor task and the Claude (no GPU) launcher work from a path
  with an apostrophe in it.

  The path was put between single quotes as it was, so a `'` in it, as in a
  user folder named O'Brien, ended the string early and the command failed.
  Each `'` is now doubled, as PowerShell expects inside single quotes.

- Prevent's health monitor step starts the new monitor only when no other
  is left running, and says when it could not replace one.

  Step 15 removed the old task and stopped the running monitor before it
  registered the new task, and dropped every error: a registration that
  failed left no monitor and was reported as installed, a monitor that did
  not stop was not mentioned and the new one was started beside it, and a
  start that failed was not said. The task is now registered first,
  replacing the old one, and a registration that fails fails the step and
  leaves the running monitor as it is. The new monitor is started at once
  only when every running monitor was found and stopped: when the running
  processes cannot be listed, or a monitor is still running 2 seconds after
  it was stopped, the step warns, naming why, and the new monitor starts at
  the next logon. A start that fails is a warning too. A monitor that ended
  by itself before it could be stopped is said to have ended, here and in
  `-Undo`, not to have been stopped.

- The WSL2 check sees a running distro.

  It read `wsl.exe`'s UTF-16 output as text, so the test for a running WSL2
  distro never matched. It now sets `WSL_UTF8`. WSL or Docker Desktop merely
  being installed is listed, no longer a warning.

- Prevent says what each step did, and exits 1 when one failed.

  The summary was a fixed list that called every setting configured, printed
  after steps that had failed, said "not critical" or not run at all, and the
  script exited 0 either way. Each step now reports changed, no change needed,
  skipped, check the note or FAILED, with the reason. The summary lists what
  needs a Windows restart to take effect, and the exit code is 1 when a step
  failed or the run stopped.

- `-WhatIf` on Prevent is refused instead of changing the machine.

  It was accepted and carried across the elevation boundary, but the only step
  that honoured it was the stale VM removal, so a "dry run" was a real run. It
  is now refused before anything runs, elevation included, and the script
  exits 1.

### Changed

- Prevent sets the active power plan to full performance on AC power.

  It set five values, through `powercfg /change` and, on the active plan,
  `powercfg /setacvalueindex`, and checked none. It now sets these on the
  active plan and reads each one back: sleep, hibernate and hard disk
  turn-off at Never, SATA link power management Active and adaptive link
  power off, the NVMe idle timeouts off and NVMe latency tolerance 0 ms, USB
  selective suspend and USB 3 link power management off, PCI Express link
  state off, minimum and maximum processor state 100%, boost mode
  Aggressive, no core parking, energy performance preference 0, processor
  throttle states off, and wireless adapters at maximum performance. A
  setting the machine does not have is left alone. Battery (DC) settings and
  the screen timeout are not changed.

- Network adapter power saving is turned off in the five places it lives,
  and on hidden and virtual adapters too.

  On each physical adapter: Windows' "allow the computer to turn off this
  device to save power"; the driver framework's power-down while the PC is
  running (`IdleInWorkingState` in the adapter's WDF key, set to 0 where it
  is there); bit 0x08 of the driver key's `PnPCapabilities`; the driver's
  own options by registry keyword (Energy Efficient Ethernet, Green
  Ethernet, Gigabit Lite, Ultra Low Power Mode, System Idle Power Saver,
  Wi-Fi power save and the rest, each set to its own off value and only
  where the driver lists that value), and two hidden Realtek options,
  `PowerDownPll` and `bFwCtrlLPS`, changed only when they read 1; and
  selective suspend, sleep on disconnect and D0 packet coalescing. Hidden
  adapters, such as Wi-Fi Direct, and virtual ones, such as a VPN's, have
  the "allow the computer to turn off this device" box cleared, even when
  the physical adapters cannot be listed, which fails the step for those,
  naming why. Wake settings are no longer turned off.

- Prevent turns off Windows power throttling.

  Power throttling runs background processes on efficient cores at low clocks,
  and the VM's worker can be one of them. Prevent sets `PowerThrottlingOff` to
  1, which turns it off for every process from the next Windows restart.

- The vmcompute recovery Prevent replaces is recorded first, so `-Undo` puts
  it back. Without `vmcompute`, the step is skipped.

- Prevent runs 19 steps and `-Undo` 5. Prevent's last step removes what older
  versions set up that now does harm or nothing, and lists as "Left in place"
  these settings they made, which `-Undo` removes:
  `LocalAccountTokenFilterPolicy`, `FilterAdministratorToken`,
  `ServicesPipeTimeout`, `CsEnabled`, the restart recovery on
  CoworkVMService, the `CoworkNAT` rule, an adapter's `PnPCapabilities` with
  its wake settings, and Dynamic Memory left off on a Hyper-V VM named like
  Claude. The first three are listed, and removed, only where an older
  install is found; elsewhere the step says that they match what older
  versions set and that `-Undo` leaves them.

### Removed

- The boot-fix task.

  `\Claude\ClaudeCoworkBootFix` ran `Fix-ClaudeDesktop.ps1 -BootPrep -Quiet`
  45 seconds after every logon. It restarted `vmcompute` while Claude was not
  yet running, which is a restart of the service every VM on the machine
  depends on. Prevent and `-Undo` remove it.

- What Fix's `-BootPrep` did.

  `-BootPrep` restarted `vmcompute`, and removed a stale Cowork compute
  system, when no Cowork workspace was up. The switch is kept, because a
  boot-fix task an older version installed still passes it until Prevent or
  `-Undo` removes the task, but it now exits 0 before anything else runs,
  saying only that it is retired: it writes no log or file, deletes nothing,
  and touches neither `vmcompute` nor any VM.

- Dynamic Memory pinning and its flag file.

  `disable-dynamic-memory.flag` asked the monitor to turn off Dynamic Memory
  on a Hyper-V VM named like Claude. The Cowork VM is not a Hyper-V VM and has
  no such setting. Prevent and `-Undo` delete the flag. Where an older
  install is found, `-Undo` turns Dynamic Memory back on for such a VM that
  is turned off, and for one that is running says to turn it off and run
  `-Undo` again.

- The health monitor's guest-connection and NAT checks.

  The guest-connection check read `cowork-service.log` for a message that file
  has never held, and the file is no longer written. The NAT check looked for
  a WinNAT rule, and the Cowork VM does not use one. A `Request timed out:
  isGuestConnected` the app logs at `[error]` in its VM or main log still
  starts a repair through the log check.

- The health monitor's Hyper-V heartbeat check.

  It read the heartbeat through `Get-VM`, which never lists an HCS compute
  system such as the Cowork VM, so it could not fire, and from 6.0.0 it did
  nothing. The monitor's help no longer lists it, nor the guest-connection
  and NAT checks, and the startup banner still says the heartbeat is not
  monitored.

- Prevent's WinNAT and firewall policy checks.

  The Cowork VM does not use a WinNAT rule. The `CoworkNAT` rule older
  versions created, and the monitor re-created, is listed by a new run when it
  is still there, and `-Undo` removes it. A machine without WinNAT has no such
  rule, and is not warned about. When the NAT rules cannot be read, Prevent
  and `-Undo` warn that `CoworkNAT` was not checked.

- Four settings Prevent no longer makes: Connected Standby (`CsEnabled`),
  `ServicesPipeTimeout`, the admin token policy
  (`LocalAccountTokenFilterPolicy` and `FilterAdministratorToken`), and the
  CoworkVMService recovery.

  Current Windows no longer reads `CsEnabled`; the Modern Standby step sets
  `PlatformAoAcOverride` instead. A new run lists what older versions left
  of these as left in place. `-Undo` puts `CsEnabled` back from its backup.
  It removes `ServicesPipeTimeout`, `LocalAccountTokenFilterPolicy` and
  `FilterAdministratorToken` while they hold exactly what those versions
  wrote, 120000, 1 and 0, and only where an older install is found: a
  backup file, flag file, script or launcher an older version left, or the
  boot-fix task. It clears the restart after 30, 60 and 120 seconds on
  CoworkVMService and `vmcompute` where it is still exactly that, except
  while Prevent's record of it is kept, since its first step puts that back.

## [6.2.1] - 2026-11-21

### Fixed

- The health monitor no longer restarts Claude over a harmless log line.

  The log check starts a repair on a single hit, and it matched its error
  strings against every new byte of every `*.log` in the app's logs folder, at
  any log level. App 6.18355 writes `RPC error -1` in an `[info]
  [remote-bash]` line when a remote session's probe finds its worker already
  running and carries on. Measured 2026-11-21: 14 such lines in `main.log` and
  its rotations, on 2026-10-31, 11-07, 11-16 and 11-18, and each one reached
  the repair path, which restarts the app.

  A signature now counts only inside an entry the app logged at `[error]`,
  never from a `remote*` component, and only in `cowork_vm_node*.log` and
  `main*.log`. The browser, web and MCP logs hold what web pages and MCP
  servers print, not the VM. The real VM failures in the same logs, the
  `[error] [VM:start] Startup failed` entries of 2026-09-11 with `Plan9 mount
  failed`, still start a repair. `bad address` left the pattern list: it is
  the bare text of EFAULT, and the failure it described also carries
  `Plan9 mount failed` or `failed to ensure virtiofs mount`.

  `.ci/Test-LogTriggers.ps1` feeds the captured lines, worker names and
  process ids replaced, to the shipped `Test-LogsForErrors` against a
  throwaway logs folder.

## [6.2.0] - 2026-11-21

Read "Repairs that could not run before can run now" below before you pull
this onto a machine where the health monitor or the boot-fix task runs. Both
run the scripts straight from this folder, so pulling is deploying.

### Fixed

- A repair no longer ends every Claude Code session on the machine.

  "claude" is not a unique process name. Measured 2026-11-21 on app 6.18355:
  18 processes answered to it, 14 of them the desktop app, 3 Claude Code CLI
  processes the app had started for project threads, and 1 a CLI in a
  terminal. Every kill in the toolkit stopped all 18. Fix's Step 1, `-Close`,
  the orphan sweep and both escalations now stop only processes whose image
  is under the app's package install folder (or the discovered `claude.exe`'s
  install root), and say how many other `claude` processes they left alone.
  The launch check, the remaining-process check and the closing summary count
  the app only, so a surviving CLI no longer reads as a successful launch.

- A relaunch starts the app, not the terminal CLI.

  The old Step 1 cached the exe of the first process named `claude`, and the
  cache check only knew the CLI by its `claude-code` folder, so a cache
  written while a terminal CLI ran first could hold `.local\bin\claude.exe`
  and every relaunch would start the CLI. The cached path, the App Paths key
  and the Start Menu shortcut are now each refused when they name the CLI by
  folder or by product name, and a refused cache is deleted.

- The health monitor sees the app quit while a CLI is running.

  With every `claude` counted, a live CLI kept "Claude is running" true after
  the app had gone, so the exit was never logged, a stealth update was never
  relaunched, the checks ran against an absent app, and every thread wake read
  as an app launch that deferred repairs. The running test, its startup seed,
  the service check and the startup deferral now count the app only. The
  checks that only hold a repair back (user activity, the pre-fix recheck,
  the vmcompute spike guard) still count every `claude` process.

- The Claude Code guard finds a running session again.

  `Test-ClaudeCodeActive` keeps the monitor from repairing the app under a
  working Code session. It looked in `%APPDATA%\Claude\claude-code-sessions`
  and `%APPDATA%\Claude\logs`, and on app 6.18355 neither is written, so it
  answered "no session" with one running. It now searches every candidate
  data root, recursively, and every candidate log folder written in the last
  ten minutes, and counts `remote-control.log` as a session too.

  Those are files the app writes about a session, and it does not write them
  while the session works: measured 2026-11-21, two working project threads
  had written their transcripts that minute while the newest session file was
  16 minutes old and the logs 13 to 38. The guard now also reads Claude
  Code's own transcripts, any `*.jsonl` under `%USERPROFILE%\.claude\projects`
  (or `CLAUDE_CONFIG_DIR`) written in the last 35 minutes. Run on that
  machine, the old guard answered no and the new one yes. A terminal Claude
  Code session counts as well, which delays a repair by two minutes at a
  time.

  The window is 35 minutes because a tool call writes nothing to the
  transcript while it runs, and a foreground command may run for 30.
  Measured 2026-11-21 in one project thread: three Bash calls of 25 minutes
  each, nothing written during the first two. On the transcripts alone, a
  10-minute window answered "no session" from minute 10 of each. The cost is
  that a repair waits up to 35 minutes after the last Claude Code activity.

- The app's logs, VM cache, bundles and session folders are found where the
  app keeps them.

  They were fixed paths under `%APPDATA%\Claude`. App 6.18355 writes its logs
  to `%LOCALAPPDATA%\Claude\logs` and its data to
  `%LOCALAPPDATA%\Claude-Data\claude_<id>`, and `%APPDATA%\Claude` holds only
  this toolkit's own files. The log scan read a folder nobody wrote, the VM
  log check had no file, and Fix's cache backup and purge pointed at a folder
  that did not exist. Discovery now picks the freshest candidate for each; see
  "Where the app keeps its files" in the README. The toolkit's own files
  (`fix-logs`, `watch-logs`, `watch-state.json`) stay in `%APPDATA%\Claude`.

- Prevent's generated launchers leave the CLI alone. The "Claude (no GPU)"
  launcher closes only the app, and the admin launcher's already-running test
  skips the CLI. Its third way of finding the exe, reading it off a running
  `claude` process, is gone: it only ran when the app was not running, so it
  could only find the CLI. Run Prevent again to regenerate both.

- Every VM log is read, each in its own format.

  App 6.18355 writes a VM boot to two logs in the same second:
  `cowork_vm_node.log` gets "Startup complete" and "[Keepalive] Starting",
  `coworkd.log` gets "connected, waiting for commands", with slashes in its
  dates. The checks read only the VM log that was newest when the script
  started, which is usually `coworkd.log`, and looked in it for the node
  log's markers. Fix's wait for the VM after a restart therefore saw none of
  its readiness markers, and a Smart repair could wait out its deadline on a
  VM that had booted and escalate to Deep, which deletes the VM cache. The
  monitor's Cowork guard, which holds a repair back while a Cowork session
  waits on the model, always answered "no session". The monitor's VM log
  freshness checks kept reading the file named at startup after the other
  log had become the newer one.

  Fix now reads every VM log from where each stood when the wait began,
  reads a log that was rotated or that appeared during the wait from its
  start, and knows coworkd's boot line and its process starts, whose id is
  now `oneshot-<uuid>` and never matched the old hex-only pattern. The
  Cowork guard reads `cowork_vm_node.log` by name, and the freshness checks
  read every VM log on each poll. Fed the same log lines, the code before
  this fix answered "no session" and saw no boot at all.

### Changed

- The health monitor looks for the app's folders again every ten minutes and
  whenever the app starts, and logs a line when the data root, the log folder
  or the VM log moves. A fallback answer is never followed. Its startup banner
  shows the log folder and the data root, each with how it was found.

- Fix prints the app's data root and log folder at the top of its log.

- The stealth-update marker is looked for in the data root, where the app's
  userData now is. That location is inferred from the 6.18355 layout, not
  seen during a real stealth update.

- Prevent's storage check and its Defender exclusion use the app's data root.
  Running Prevent adds an exclusion for `%LOCALAPPDATA%\Claude-Data\claude_<id>`
  on a 6.18355 machine; the old `%APPDATA%\Claude` exclusion is not removed.

  An existing exclusion now counts only when it is the same path, a folder
  above it, or a wildcard that matches it. The old test compared the start of
  the string, so an exclusion for the app's own `%LOCALAPPDATA%\Claude`
  counted for `%LOCALAPPDATA%\Claude-Data` beside it, and an empty entry
  counted for everything. Measured 2026-11-21: a machine with that exclusion,
  where the old test would never have added the data root. The process
  exclusions are now checked on every run; they were only looked at when a
  path exclusion was missing.

- `claude-env.json` is schema version 2: it gains `DataRoot`, `LogDir`, how
  each was found, the VM-cache, bundle and session folders under the data
  root, and the candidates passed over. `Import-ClaudeEnvironment` refuses a
  version 1 file.

- CI: the kill check parses the three scripts and the scripts they generate,
  instead of matching one line at a time, and follows a process lookup through
  a variable or a foreach. Run on 0aa8d07 it finds all six broad kills, where
  the line check found the two escalations only, and it proves itself against
  copies of those kills on every run. The launcher check builds the "Claude
  (no GPU)" launcher from Prevent's own source and runs its process filter
  against fake processes, instead of testing a hand copy that had drifted.

### Repairs that could not run before can run now

The path fix re-arms checks that pointed at empty folders:

- The log scan reads live logs again, so a VirtioFS or HCS error there can
  start an automatic repair. New log files are read from their current end,
  so the backlog is not replayed.
- The VM log checks have files to read, and Fix's wait for the VM after a
  restart reads every one of them (see above).
- A deep or escalated Fix backs up and purges the real VM cache. Measured
  2026-11-21 that is about 18 GB under
  `%LOCALAPPDATA%\Claude-Data\claude_<id>\claude-code-vm`: `rootfs.vhdx`
  (8.6 GB) is deleted and re-downloaded, `sessiondata.vhdx` (9.7 GB) is
  backed up to `fix-logs\vhdx-backup` and restored. Check the free space on
  the drive holding `%APPDATA%` before a deep purge.
- A stealth-update relaunch can fire again.

## [6.1.1] - 2026-10-09

### Fixed

- The "Claude (Admin)" launcher no longer flashes a console window.

  The `LaunchClaudeAdmin` scheduled task executed `PowerShell.exe` directly.
  Under its Interactive principal, Task Scheduler hands a console-subsystem
  program a brand-new visible console window, and `-WindowStyle Hidden` cannot
  prevent it: PowerShell only hides its console after conhost has already drawn
  it. Measured 2026-10-09, the task's window took the foreground for ~45 ms on
  every launch, and the desktop shortcut's `.cmd` added a second cmd window of
  its own (~1 s in the foreground) before `schtasks` even ran. Either is enough
  to drop a fullscreen game to the desktop.

  The task now runs `wscript.exe` -- a GUI-subsystem host, so it opens no
  console of its own -- pointing at a generated `Launch-Claude-Admin.vbs`, which
  starts PowerShell with its console hidden at creation time via
  `WScript.Shell.Run(cmd, 0, ...)` (vbHide). The desktop shortcut now points at
  `wscript.exe` running `Launch-Claude-Admin-Trigger.vbs`, which fires the task
  the same hidden way. Claude still launches with the full, unfiltered admin
  token: `wscript` runs at the task's `RunLevel Highest` and PowerShell inherits
  it, so the elevation path, the Claude-discovery payload and the no-UAC-prompt
  behaviour are all unchanged -- only the window is gone. Verified with a window
  monitor across both launch steps: zero windows.

  `Undo` removes the two new `.vbs` files, and setup deletes the superseded
  `Launch-Claude-Admin.cmd` left by an earlier install.

## [6.1.0] - 2026-08-28

### Added

- `Watch-ClaudeHealth` relaunches Claude Desktop after a stealth update, and
  only then.

  Claude applies background updates by quitting itself once idle and handing
  off to the Electron updater. On the MSIX path `quitAndInstall()` calls the
  AppX deployment API with `ForceApplicationShutdownOption`, and registering a
  package does not start the app afterwards. Measured on 2026-08-28: quit at
  19:44:07, package registered successfully at 19:44:39, nothing came back, and
  the app stayed down until it was started by hand at 19:47:43. Not a crash --
  no 1000, no 1002, no WER `APPCRASH`, no dump.

  The app writes a marker file, `stealth-relaunch` in its userData folder, on
  the way out, recording window position and navigation history, and restores
  from it if the next launch happens within five minutes. That marker is the
  discriminator this monitor was missing: it exists if and only if the app went
  down through the stealth path. A deliberate quit writes nothing, so an app the
  user closed on purpose still stays closed, which is what the previous
  behaviour was protecting and what the log line "a crash and a deliberate quit
  look the same from here" was about.

  The marker is never deleted here. The app consumes it on next launch, and
  deleting it would throw away the window and tab restore that makes the whole
  thing invisible to the user. On the default 30-second poll the relaunch lands
  well inside the five-minute honour window.

  Bounded on both sides. `-MaxStealthMarkerAgeMinutes` (default 10) refuses to
  resurrect the app from a marker old enough to mean the machine was off or the
  user chose to stay closed. The marker's `ts` is recorded, and persisted, before
  the launch is attempted rather than after, so a launch that throws cannot turn
  into a relaunch attempt every thirty seconds forever.

  `-NoStealthRelaunch` turns the whole thing off. No new scheduled task:
  `\Claude\ClaudeCoworkWatchdog` already runs this script on a logon trigger.

- `LastStealthMarkerTs` in `watch-state.json`. State files written by earlier
  versions do not have the field and load without it.

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
