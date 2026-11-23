<#
.SYNOPSIS
    Drives the desktop-or-CLI process classifier in the shared ClaudeEnv region
    against fake processes, and checks no script kills by name any more.
.DESCRIPTION
    "claude" is not a unique process name. The Claude Code CLI runs under it
    too: the copies Claude Desktop starts for Code sessions and project
    threads, and any terminal install. Until 6.2.0 every kill in this toolkit
    stopped all of them. Measured 2026-11-21 on app 6.18355: 18 processes
    named claude, 14 the app, 3 app-started CLI, 1 terminal CLI.

    Get-Process, Get-Item and Get-CimInstance are replaced by functions here,
    which PowerShell resolves before the cmdlets, so the region's own code runs
    against processes this script describes. Nothing real is touched.

    The second half parses the three scripts, and the scripts they generate,
    and fails on a kill that takes processes chosen by the name claude without
    narrowing them to the app. Before trusting a clean result it proves the
    scan flags every broad kill the toolkit shipped until 6.2.0, and passes the
    narrowed kills that replaced them. The section comment says what the scan
    follows and what it does not.

    Exits 0 when every case passes, 1 otherwise.
#>
[CmdletBinding()]
param(
    [string]$Root
)

if (-not $Root) {
    $here = $PSScriptRoot
    if (-not $here) { $here = Split-Path $MyInvocation.MyCommand.Path -Parent }
    $Root = Split-Path $here -Parent
}

$pass = 0
$fail = 0
function Assert-Case {
    param([string]$Name, [bool]$Condition, [string]$Detail = '')
    if ($Condition) { Write-Host ("  PASS  {0}" -f $Name); $script:pass++ }
    else { Write-Host ("  FAIL  {0}  {1}" -f $Name, $Detail); $script:fail++ }
}

. (Join-Path $PSScriptRoot 'ClaudeEnv.region.ps1')

# ---------------------------------------------------------------------
# Fakes. Each fake process is Id, Path (what MainModule.FileName returns, or
# $null for a process whose module list cannot be read), CimPath (what
# Win32_Process.ExecutablePath returns for it) and Parent (its
# Win32_Process.ParentProcessId). $script:CimThrows makes the process query
# fail, as it does when WMI is broken or the caller lacks the right.
#
# Each fake writes the error its cmdlet writes, with the identity measured
# 2026-11-22 on Windows 11, Windows PowerShell 5.1 (the region's helpers
# carry the measurements): Get-Process for a name nothing runs under writes
# NoProcessFoundForGivenName, ObjectNotFound, ProcessCommandException;
# Get-Item for a path with nothing there writes PathNotFound, ObjectNotFound,
# ItemNotFoundException. A read that fails for another reason (access denied
# could not be produced read-only on this PC) gets another identity here,
# which is what the helpers must not take for "not there". The error is
# written, not thrown, as the cmdlets do; -ErrorAction Stop makes it
# terminating in the caller.
# ---------------------------------------------------------------------
Add-Type -AssemblyName 'Microsoft.PowerShell.Commands.Management'
$script:FakeProcs     = @()
$script:FakeProducts  = @{}
$script:FakeMissing   = @{}
$script:FakeDenied    = @{}
$script:FakeNoVersion = @{}
$script:FakeReal      = @{}
$script:CimThrows     = $false
$script:ProcListFails = $false

function New-FakeErrorRecord {
    param([string]$TypeName, [string]$Message, [string]$Id, [string]$Category, $Target)
    $ex = New-Object $TypeName $Message
    return (New-Object System.Management.Automation.ErrorRecord($ex, $Id,
        [System.Management.Automation.ErrorCategory]$Category, $Target))
}

function New-FakeProc {
    param([int]$Id, [string]$Path, [string]$CimPath, [int]$Parent = 0)
    $o = [pscustomobject]@{ Id = $Id; Name = 'claude'; CimPath = $CimPath; FakePath = $Path; Parent = $Parent }
    if ($Path) {
        $o | Add-Member -MemberType NoteProperty -Name MainModule -Value ([pscustomobject]@{ FileName = $Path })
    } else {
        $o | Add-Member -MemberType ScriptProperty -Name MainModule -Value { throw 'Access is denied' }
    }
    return $o
}

function Get-Process {
    [CmdletBinding()]
    param([string]$Name)
    if ($script:ProcListFails) {
        $PSCmdlet.WriteError((New-FakeErrorRecord 'System.ComponentModel.Win32Exception' 'Access is denied' `
            'CouldNotEnumerateProcesses' 'NotSpecified' $Name))
        return
    }
    $hits = @($script:FakeProcs | Where-Object { $_.Name -eq $Name })
    if ($hits.Count -eq 0) {
        $PSCmdlet.WriteError((New-FakeErrorRecord 'Microsoft.PowerShell.Commands.ProcessCommandException' `
            "Cannot find a process with the name `"$Name`". Verify the process name and call the cmdlet again." `
            'NoProcessFoundForGivenName' 'ObjectNotFound' $Name))
        return
    }
    return $hits
}

# A path not listed in FakeMissing or FakeDenied is a file that is there,
# with the product name FakeProducts gives it, or none. FakeNoVersion makes
# its VersionInfo read as $null, as a getter that throws does. A path in
# FakeReal is read by the real Get-Item, for a real folder this script makes.
function Get-Item {
    [CmdletBinding()]
    param([string]$LiteralPath, [switch]$Force)
    if ($script:FakeReal.ContainsKey($LiteralPath)) {
        return (Microsoft.PowerShell.Management\Get-Item -LiteralPath $LiteralPath -Force:$Force)
    }
    if ($script:FakeDenied.ContainsKey($LiteralPath)) {
        $PSCmdlet.WriteError((New-FakeErrorRecord 'System.UnauthorizedAccessException' `
            'Access is denied' 'ItemExistsUnauthorizedAccessError' 'PermissionDenied' $LiteralPath))
        return
    }
    if ($script:FakeMissing.ContainsKey($LiteralPath)) {
        $PSCmdlet.WriteError((New-FakeErrorRecord 'System.Management.Automation.ItemNotFoundException' `
            "Cannot find path '$LiteralPath' because it does not exist." 'PathNotFound' 'ObjectNotFound' $LiteralPath))
        return
    }
    $product = ''
    if ($script:FakeProducts.ContainsKey($LiteralPath)) { $product = $script:FakeProducts[$LiteralPath] }
    $vi = [pscustomobject]@{ ProductName = $product }
    if ($script:FakeNoVersion.ContainsKey($LiteralPath)) { $vi = $null }
    return [pscustomobject]@{ FullName = $LiteralPath; PSIsContainer = $false; VersionInfo = $vi }
}

# The real cmdlet writes a CimException (measured 2026-11-22 for an invalid
# class: HRESULT 0x80041010, MetadataError) and does not throw.
function Get-CimInstance {
    [CmdletBinding()]
    param([string]$ClassName, [string]$Filter)
    $null = $ClassName, $Filter
    if ($script:CimThrows) {
        $PSCmdlet.WriteError((New-FakeErrorRecord 'Microsoft.Management.Infrastructure.CimException' 'Invalid class' `
            'HRESULT 0x80041010' 'MetadataError' $ClassName))
        return
    }
    return @($script:FakeProcs | ForEach-Object {
        [pscustomobject]@{ ProcessId = $_.Id; ExecutablePath = $_.CimPath; ParentProcessId = $_.Parent }
    })
}

function Get-KindById {
    param($Environment)
    $map = @{}
    foreach ($r in @(Get-ClaudeProcessInventory -Environment $Environment)) { $map[[int]$r.Id] = $r.Kind }
    return $map
}

$wa      = 'C:\Program Files\WindowsApps'
$install = "$wa\Claude_1.0.4.0_x64__pzs8sxrjxfjjc"
$msixEnv = [pscustomobject]@{
    ProcessName       = 'claude'
    InstallLocation   = $install
    PackageFamilyName = 'Claude_pzs8sxrjxfjjc'
    ClaudeExe         = "$install\app\Claude.exe"
}

$appCli   = 'C:\Users\u\AppData\Local\Claude-Data\claude_pzs8sxrjxfjjc\claude-code\2.1.327\claude.exe'
$termCli  = 'C:\Users\u\.local\bin\claude.exe'
$stranger = 'C:\tools\claude.exe'
$script:FakeProducts[$termCli]  = 'Claude Code'
$script:FakeProducts[$stranger] = 'Some Other Tool'

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- MSIX install: the 6.18355 population"

$script:FakeProcs = @(
    (New-FakeProc 101 "$install\app\Claude.exe"),
    (New-FakeProc 102 "$install\app\Claude.exe"),
    (New-FakeProc 103 "$wa\Claude_1.0.3.0_x64__pzs8sxrjxfjjc\app\Claude.exe"),
    (New-FakeProc 201 $appCli),
    (New-FakeProc 202 $termCli),
    (New-FakeProc 301 $stranger),
    (New-FakeProc 302 "$wa\Claude_1.0.4.0_x64__someoneelse\app\Claude.exe"),
    (New-FakeProc 303 "${install}X\app\Claude.exe"),
    (New-FakeProc 401 $null "$install\app\Claude.exe"),
    (New-FakeProc 402 $null $null)
)
$k = Get-KindById $msixEnv
Assert-Case "app under InstallLocation is desktop"              ($k[101] -eq 'desktop') "got $($k[101])"
Assert-Case "every app process is desktop, not just the first"  ($k[102] -eq 'desktop') "got $($k[102])"
Assert-Case "app from the previous package version is desktop"  ($k[103] -eq 'desktop') "got $($k[103])"
Assert-Case "app-started CLI under claude-code is cli"          ($k[201] -eq 'cli')     "got $($k[201])"
Assert-Case "terminal CLI is cli by its product name"           ($k[202] -eq 'cli')     "got $($k[202])"
Assert-Case "an unrelated claude.exe is other"                  ($k[301] -eq 'other')   "got $($k[301])"
Assert-Case "same package name, other publisher is other"       ($k[302] -eq 'other')   "got $($k[302])"
Assert-Case "a folder that only starts like InstallLocation is other" ($k[303] -eq 'other') "got $($k[303])"
Assert-Case "unreadable module, path from CIM, is desktop"      ($k[401] -eq 'desktop') "got $($k[401])"
Assert-Case "no path anywhere is unknown"                       ($k[402] -eq 'unknown') "got $($k[402])"

$desk = @(Get-ClaudeDesktopProcess -Environment $msixEnv | ForEach-Object { [int]$_.Id } | Sort-Object)
Assert-Case "Get-ClaudeDesktopProcess returns the app and nothing else" `
    (($desk -join ',') -eq '101,102,103,401') "got $($desk -join ',')"

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Squirrel install: no package, ClaudeExe stands in"

$sqRoot = 'C:\Users\u\AppData\Local\AnthropicClaude'
$sqEnv = [pscustomobject]@{
    ProcessName       = 'claude'
    InstallLocation   = $null
    PackageFamilyName = $null
    ClaudeExe         = "$sqRoot\app-1.2.3\Claude.exe"
}
$script:FakeProcs = @(
    (New-FakeProc 501 "$sqRoot\app-1.2.3\Claude.exe"),
    (New-FakeProc 502 "$sqRoot\app-1.2.4\Claude.exe"),
    (New-FakeProc 503 $appCli),
    (New-FakeProc 504 $termCli)
)
$k = Get-KindById $sqEnv
Assert-Case "app in the discovered app-<version> is desktop"   ($k[501] -eq 'desktop') "got $($k[501])"
Assert-Case "app in a newer app-<version> is desktop"          ($k[502] -eq 'desktop') "got $($k[502])"
Assert-Case "app-started CLI is cli"                           ($k[503] -eq 'cli')     "got $($k[503])"
Assert-Case "terminal CLI is cli"                              ($k[504] -eq 'cli')     "got $($k[504])"

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- nothing discovered: the product name decides"

$bareEnv = [pscustomobject]@{
    ProcessName       = 'claude'
    InstallLocation   = $null
    PackageFamilyName = $null
    ClaudeExe         = $termCli
}
$someApp = 'D:\Apps\Claude\Claude.exe'
$script:FakeProducts[$someApp] = 'Claude'
$script:FakeProcs = @(
    (New-FakeProc 601 $someApp),
    (New-FakeProc 602 $termCli),
    (New-FakeProc 603 $stranger)
)
$k = Get-KindById $bareEnv
Assert-Case "a ClaudeExe that is the CLI is not used as the app's root" ($k[602] -eq 'cli') "got $($k[602])"
Assert-Case "ProductName Claude is desktop"                    ($k[601] -eq 'desktop') "got $($k[601])"
Assert-Case "anything else is other"                           ($k[603] -eq 'other')   "got $($k[603])"

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Test-ClaudeCodeCliPath"

Assert-Case "claude-code folder is the CLI"    (Test-ClaudeCodeCliPath $appCli)
Assert-Case "product name Claude Code is the CLI" (Test-ClaudeCodeCliPath $termCli)
Assert-Case "the app is not the CLI"           (-not (Test-ClaudeCodeCliPath "$install\app\Claude.exe"))
Assert-Case "an empty path is not the CLI"     (-not (Test-ClaudeCodeCliPath ''))

# Until 6.3.1 a file whose product name could not be read answered "not the
# CLI", so the callers that pick the app's exe could take it.
$deniedExe  = 'D:\Locked\claude.exe'
$noVerExe   = 'D:\NoVersion\claude.exe'
$missingExe = 'D:\Gone\claude.exe'
$script:FakeDenied[$deniedExe]     = $true
$script:FakeNoVersion[$noVerExe]   = $true
$script:FakeMissing[$missingExe]   = $true
Assert-Case "a file that cannot be read is not taken for the app" (Test-ClaudeCodeCliPath $deniedExe)
Assert-Case "a file whose version info reads as null is not taken for the app" (Test-ClaudeCodeCliPath $noVerExe)
Assert-Case "a path with nothing there is not the CLI" (-not (Test-ClaudeCodeCliPath $missingExe))

# A folder is not the CLI. A Claude*.lnk or an App Paths value can name one,
# and a folder has no VersionInfo, so reading it threw under StrictMode
# Latest, as the scripts run, and Fix's search for Claude.exe then said the
# Start Menu or the App Paths key could not be read. A real folder, read by
# the real Get-Item, made here and removed after.
$realDir = Join-Path ([System.IO.Path]::GetTempPath()) ('claudefix-ci-tpi-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$dirAnswer = $null
$dirError  = ''
try {
    $null = New-Item -ItemType Directory -Path $realDir
    $script:FakeReal[$realDir] = $true
    $dirAnswer = & { Set-StrictMode -Version Latest; Test-ClaudeCodeCliPath $realDir }
} catch {
    $dirError = $_.Exception.Message
} finally {
    $script:FakeReal.Remove($realDir)
    # Only the folder made above, which is empty.
    Remove-Item -LiteralPath $realDir -Force -ErrorAction SilentlyContinue
}
Assert-Case "a folder is not the CLI, and asking does not throw" (
    $dirAnswer -is [bool] -and -not $dirAnswer -and -not $dirError) "got '$dirAnswer', error '$dirError'"

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Get-ClaudeProcessInventory: a list that could not be read"

# Until 6.3.1 Get-Process ran with SilentlyContinue, so a list that could
# not be read came back empty and every caller read "Claude is not running".
$script:FakeProcs = @()
$got = '<threw>'
try { $got = "$(@(Get-ClaudeProcessInventory -Environment $msixEnv).Count)" } catch { $got = "<threw: $($_.FullyQualifiedErrorId)>" }
Assert-Case "none running: an empty inventory, no error" ($got -eq '0') "got $got"
$got = '<threw>'
try { $got = "$(@(Get-ClaudeDesktopProcess -Environment $msixEnv).Count)" } catch { $got = "<threw: $($_.FullyQualifiedErrorId)>" }
Assert-Case "none running: no desktop process, no error" ($got -eq '0') "got $got"

$script:FakeProcs = @((New-FakeProc 101 "$install\app\Claude.exe"))
$script:ProcListFails = $true
$threwWith = ''
try { $null = Get-ClaudeProcessInventory -Environment $msixEnv } catch { $threwWith = "$($_.FullyQualifiedErrorId)" }
Assert-Case "a list that could not be read throws, with the error it came with" ($threwWith -match '^CouldNotEnumerateProcesses') "threw: $threwWith"
$threwWith = ''
try { $null = Get-ClaudeDesktopProcess -Environment $msixEnv } catch { $threwWith = "$($_.FullyQualifiedErrorId)" }
Assert-Case "and Get-ClaudeDesktopProcess passes it on" ($threwWith -match '^CouldNotEnumerateProcesses') "threw: $threwWith"
$threwWith = ''
try { $null = Get-ClaudeAppCliProcess -Environment $msixEnv } catch { $threwWith = "$($_.FullyQualifiedErrorId)" }
Assert-Case "and Get-ClaudeAppCliProcess passes it on" ($threwWith -match '^CouldNotEnumerateProcesses') "threw: $threwWith"
$script:ProcListFails = $false

# A process whose module cannot be read is placed by one CIM query. When
# that fails too, the process cannot be placed, and until 6.3.1 the failure
# was passed over: an app process read as unknown and was not counted.
$script:FakeProcs = @(
    (New-FakeProc 101 "$install\app\Claude.exe"),
    (New-FakeProc 401 $null "$install\app\Claude.exe")
)
$script:CimThrows = $true
$threwWith = ''
try { $null = Get-ClaudeProcessInventory -Environment $msixEnv } catch { $threwWith = "$($_.FullyQualifiedErrorId)|$($_.Exception.GetType().Name)" }
Assert-Case "an unreadable module and a failed CIM query throw the CIM error" ($threwWith -match '^HRESULT 0x80041010.*\|CimException$') "threw: $threwWith"
$script:CimThrows = $false

# A product name that could not be read leaves the process unknown, not
# other: it may be the CLI. Until 6.3.1 it was other.
$script:FakeProcs = @(
    (New-FakeProc 101 "$install\app\Claude.exe" '' 4000),
    (New-FakeProc 701 $deniedExe '' 101),
    (New-FakeProc 702 $noVerExe '' 101),
    (New-FakeProc 703 $deniedExe '' 900)
)
$k = Get-KindById $msixEnv
Assert-Case "a product name that cannot be read is unknown"           ($k[701] -eq 'unknown') "got $($k[701])"
Assert-Case "a version info that reads as null is unknown"            ($k[702] -eq 'unknown') "got $($k[702])"
$got = '<threw>'
try { $got = (@(Get-ClaudeAppCliProcess -Environment $msixEnv | ForEach-Object { [int]$_.Id } | Sort-Object) -join ',') } catch { $got = "<threw: $($_.Exception.Message)>" }
Assert-Case "an unknown child of the app counts as one of its sessions" ($got -eq '701,702') "got $got"

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- the helpers that tell not there from could not read"

# Each helper is driven with a fake of its cmdlet in a scope of its own: the
# error that says "not there", the same id from a script function (whose
# id PowerShell suffixes with its own name), a failure of another kind, an
# error that has the not-there id but another category or type, and a read
# that answers.
$helperCase = & {
    Set-StrictMode -Version Latest
    $mode = 'ok'
    function Get-Service {
        [CmdletBinding()] param([string]$Name)
        switch ($mode) {
            'missing' { $PSCmdlet.WriteError((New-FakeErrorRecord 'Microsoft.PowerShell.Commands.ServiceCommandException' `
                            "Cannot find any service with service name '$Name'." 'NoServiceFoundForGivenName' 'ObjectNotFound' $Name)); return }
            'denied'  { $PSCmdlet.WriteError((New-FakeErrorRecord 'System.InvalidOperationException' `
                            'Cannot open Service Control Manager on computer. This operation might require other privileges.' 'CouldNotGetServiceInfo' 'PermissionDenied' $Name)); return }
            'wrongcat' { $PSCmdlet.WriteError((New-FakeErrorRecord 'Microsoft.PowerShell.Commands.ServiceCommandException' `
                            'x' 'NoServiceFoundForGivenName' 'NotSpecified' $Name)); return }
            'wrongtype' { $PSCmdlet.WriteError((New-FakeErrorRecord 'System.InvalidOperationException' `
                            "Cannot find any service with service name '$Name'." 'NoServiceFoundForGivenName' 'ObjectNotFound' $Name)); return }
        }
        return [pscustomobject]@{ Name = $Name; Status = 'Running' }
    }
    function Get-ScheduledTask {
        [CmdletBinding()] param([string]$TaskName, [string]$TaskPath)
        switch ($mode) {
            'missing' { $PSCmdlet.WriteError((New-FakeErrorRecord 'Microsoft.PowerShell.Cmdletization.Cim.CimJobException' `
                            "No MSFT_ScheduledTask objects found with property 'TaskName' equal to '$TaskName'." 'CmdletizationQuery_NotFound' 'ObjectNotFound' $TaskName)); return }
            'denied'  { $PSCmdlet.WriteError((New-FakeErrorRecord 'Microsoft.Management.Infrastructure.CimException' `
                            'Access denied' 'HRESULT 0x80070005' 'PermissionDenied' $TaskName)); return }
            'wrongcat' { $PSCmdlet.WriteError((New-FakeErrorRecord 'Microsoft.PowerShell.Cmdletization.Cim.CimJobException' `
                            'x' 'CmdletizationQuery_NotFound' 'InvalidOperation' $TaskName)); return }
            'wrongtype' { $PSCmdlet.WriteError((New-FakeErrorRecord 'System.InvalidOperationException' `
                            'x' 'CmdletizationQuery_NotFound' 'ObjectNotFound' $TaskName)); return }
        }
        return [pscustomobject]@{ TaskName = $TaskName; TaskPath = $TaskPath; State = 'Ready' }
    }
    function Get-Process {
        [CmdletBinding()] param([string]$Name)
        switch ($mode) {
            'missing' { $PSCmdlet.WriteError((New-FakeErrorRecord 'Microsoft.PowerShell.Commands.ProcessCommandException' `
                            "Cannot find a process with the name `"$Name`"." 'NoProcessFoundForGivenName' 'ObjectNotFound' $Name)); return }
            'denied'  { $PSCmdlet.WriteError((New-FakeErrorRecord 'System.ComponentModel.Win32Exception' `
                            "Cannot find a process with the name `"$Name`"." 'CouldNotEnumerateProcesses' 'NotSpecified' $Name)); return }
            'wrongcat' { $PSCmdlet.WriteError((New-FakeErrorRecord 'Microsoft.PowerShell.Commands.ProcessCommandException' `
                            'x' 'NoProcessFoundForGivenName' 'NotSpecified' $Name)); return }
            'wrongtype' { $PSCmdlet.WriteError((New-FakeErrorRecord 'System.InvalidOperationException' `
                            'x' 'NoProcessFoundForGivenName' 'ObjectNotFound' $Name)); return }
        }
        return @([pscustomobject]@{ Id = 11; Name = $Name }, [pscustomobject]@{ Id = 12; Name = $Name })
    }
    function Get-Item {
        [CmdletBinding()] param([string]$LiteralPath, [switch]$Force)
        $null = $Force
        switch ($mode) {
            'missing' { $PSCmdlet.WriteError((New-FakeErrorRecord 'System.Management.Automation.ItemNotFoundException' `
                            "Cannot find path '$LiteralPath' because it does not exist." 'PathNotFound' 'ObjectNotFound' $LiteralPath)); return }
            'denied'  { $PSCmdlet.WriteError((New-FakeErrorRecord 'System.UnauthorizedAccessException' `
                            'Access is denied' 'ItemExistsUnauthorizedAccessError' 'PermissionDenied' $LiteralPath)); return }
            'wrongcat' { $PSCmdlet.WriteError((New-FakeErrorRecord 'System.Management.Automation.ItemNotFoundException' `
                            'x' 'PathNotFound' 'ReadError' $LiteralPath)); return }
            'wrongtype' { $PSCmdlet.WriteError((New-FakeErrorRecord 'System.IO.IOException' `
                            "Cannot find path '$LiteralPath' because it does not exist." 'PathNotFound' 'ObjectNotFound' $LiteralPath)); return }
        }
        return [pscustomobject]@{ FullName = $LiteralPath; PSIsContainer = $false }
    }
    $calls = [ordered]@{
        service = { Get-ClaudeServiceOrNull -Name 'CoworkVMService' }
        task    = { Get-ClaudeScheduledTaskOrNull -TaskName 'ClaudeWatch' -TaskPath '\Claude\' }
        # The leading comma keeps an empty list from unrolling to nothing.
        process = { ,@(Get-ClaudeProcessByName -Name 'claude') }
        item    = { Get-ClaudeItemOrNull -LiteralPath 'C:\x\y.log' }
    }
    $r = [ordered]@{}
    foreach ($h in $calls.Keys) {
        foreach ($m in @('ok', 'missing', 'denied', 'wrongcat', 'wrongtype')) {
            $mode = $m
            try {
                $v = & $calls[$h]
                if ($null -eq $v) { $r["$h.$m"] = 'null' }
                elseif ($v -is [array]) { $r["$h.$m"] = "array:$($v.Count)" }
                else { $r["$h.$m"] = 'object' }
            } catch {
                $r["$h.$m"] = "threw:$($_.FullyQualifiedErrorId)|$($_.Exception.GetType().Name)"
            }
        }
    }
    $r
}
foreach ($h in @('service', 'task', 'item')) {
    Assert-Case "$h helper: a read that answers returns it"   ($helperCase["$h.ok"] -eq 'object') "got $($helperCase["$h.ok"])"
    Assert-Case "$h helper: not there is `$null"              ($helperCase["$h.missing"] -eq 'null') "got $($helperCase["$h.missing"])"
}
Assert-Case "process helper: every process it lists comes back" ($helperCase['process.ok'] -eq 'array:2') "got $($helperCase['process.ok'])"
Assert-Case "process helper: none running is empty"             ($helperCase['process.missing'] -eq 'array:0') "got $($helperCase['process.missing'])"
# The error thrown is the one the cmdlet wrote, told by its id. A bare "it
# threw" also passes where the helper does not exist and the call itself
# throws CommandNotFoundException.
$deniedId   = @{ service = 'CouldNotGetServiceInfo'; task = 'HRESULT 0x80070005'; process = 'CouldNotEnumerateProcesses'
                 item = 'ItemExistsUnauthorizedAccessError' }
$notThereId = @{ service = 'NoServiceFoundForGivenName'; task = 'CmdletizationQuery_NotFound'; process = 'NoProcessFoundForGivenName'
                 item = 'PathNotFound' }
foreach ($h in @('service', 'task', 'process', 'item')) {
    Assert-Case "$h helper: a failed read throws the original error" ($helperCase["$h.denied"] -like "threw:$($deniedId[$h]),*") `
        "got $($helperCase["$h.denied"])"
    Assert-Case "$h helper: the not-there id with another category throws" ($helperCase["$h.wrongcat"] -like "threw:$($notThereId[$h]),*") `
        "got $($helperCase["$h.wrongcat"])"
    Assert-Case "$h helper: the not-there id with another exception type throws" ($helperCase["$h.wrongtype"] -like "threw:$($notThereId[$h]),*") `
        "got $($helperCase["$h.wrongtype"])"
}
Assert-Case "service helper: the error thrown is the one Get-Service wrote" `
    ($helperCase['service.denied'] -match '^threw:CouldNotGetServiceInfo,Get-Service\|InvalidOperationException$') "got $($helperCase['service.denied'])"
Assert-Case "process helper: a message that reads like not there does not decide" `
    ($helperCase['process.denied'] -match '^threw:CouldNotEnumerateProcesses,Get-Process\|Win32Exception$') "got $($helperCase['process.denied'])"
Assert-Case "task helper: the error thrown is the one Get-ScheduledTask wrote" `
    ($helperCase['task.denied'] -match '^threw:HRESULT 0x80070005,Get-ScheduledTask\|CimException$') "got $($helperCase['task.denied'])"
Assert-Case "item helper: the error thrown is the one Get-Item wrote" `
    ($helperCase['item.denied'] -match '^threw:ItemExistsUnauthorizedAccessError,Get-Item\|UnauthorizedAccessException$') "got $($helperCase['item.denied'])"

# Test-ClaudeNotFoundError directly: the real cmdlet's id carries the
# cmdlet's class after the comma, a fake's its own name.
$realRec = New-FakeErrorRecord 'Microsoft.PowerShell.Commands.ServiceCommandException' 'x' `
    'NoServiceFoundForGivenName,Microsoft.PowerShell.Commands.GetServiceCommand' 'ObjectNotFound' 'n'
Assert-Case "Test-ClaudeNotFoundError reads the real cmdlet's id" `
    (Test-ClaudeNotFoundError -ErrorRecord $realRec -ErrorId 'NoServiceFoundForGivenName' -ExceptionType 'Microsoft.PowerShell.Commands.ServiceCommandException')
Assert-Case "Test-ClaudeNotFoundError: no record is not not-there" `
    (-not (Test-ClaudeNotFoundError -ErrorRecord $null -ErrorId 'NoServiceFoundForGivenName' -ExceptionType 'Microsoft.PowerShell.Commands.ServiceCommandException'))
Assert-Case "Test-ClaudeNotFoundError: another id is not not-there" `
    (-not (Test-ClaudeNotFoundError -ErrorRecord $realRec -ErrorId 'PathNotFound' -ExceptionType 'Microsoft.PowerShell.Commands.ServiceCommandException'))

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Get-ClaudeAppCliProcess: the sessions the app started"

# The CLI copies the app starts are its children; a terminal CLI's parent is a
# shell or python. Measured 2026-11-21: four app-started CLIs, parent the app's
# main process, and one terminal CLI under python.exe. A repair ends the
# first kind and not the second, so only the first holds an automated repair
# back.
function Get-AppCliIds {
    param($Environment)
    return (@(Get-ClaudeAppCliProcess -Environment $Environment | ForEach-Object { [int]$_.Id } | Sort-Object) -join ',')
}
$script:CimThrows = $false
$script:FakeProcs = @(
    (New-FakeProc 101 "$install\app\Claude.exe" '' 4000),
    (New-FakeProc 102 "$install\app\Claude.exe" '' 101),
    (New-FakeProc 201 $appCli '' 101),
    (New-FakeProc 203 $appCli '' 102),
    (New-FakeProc 202 $termCli '' 900),
    (New-FakeProc 204 $appCli '' 999),
    (New-FakeProc 205 $termCli '' 101)
)
$got = '<threw>'
try { $got = Get-AppCliIds $msixEnv } catch { $got = "<threw: $($_.Exception.Message)>" }
Assert-Case "a CLI whose parent is the app's main process counts"       ($got -match '(^|,)201(,|$)') "got $got"
Assert-Case "a CLI whose parent is another of the app's processes counts" ($got -match '(^|,)203(,|$)') "got $got"
Assert-Case "a terminal CLI under python does not count"                ($got -notmatch '(^|,)202(,|$)') "got $got"
Assert-Case "a CLI whose parent has exited does not count"              ($got -notmatch '(^|,)204(,|$)') "got $got"
Assert-Case "a CLI by product name, started by the app, counts"         ($got -match '(^|,)205(,|$)') "got $got"
Assert-Case "exactly the app's sessions, nothing else"                  ($got -eq '201,203,205') "got $got"

$script:FakeProcs = @(
    (New-FakeProc 201 $appCli '' 101),
    (New-FakeProc 202 $termCli '' 900)
)
$got = '<threw>'
try { $got = Get-AppCliIds $msixEnv } catch { $got = "<threw: $($_.Exception.Message)>" }
Assert-Case "no app running: no session counts" ($got -eq '') "got $got"

$script:FakeProcs = @((New-FakeProc 101 "$install\app\Claude.exe" '' 4000))
$script:CimThrows = $true
$got = '<threw>'
try { $got = Get-AppCliIds $msixEnv } catch { $got = "<threw: $($_.Exception.Message)>" }
Assert-Case "no CLI at all: empty, without asking WMI" ($got -eq '') "got $got"

$script:FakeProcs = @(
    (New-FakeProc 101 "$install\app\Claude.exe" '' 4000),
    (New-FakeProc 201 $appCli '' 101)
)
$threwWith = ''
try { $null = Get-ClaudeAppCliProcess -Environment $msixEnv } catch { $threwWith = $_.Exception.Message }
Assert-Case "a failed parent query throws, so a caller can fail closed" ($threwWith -eq 'Invalid class') "threw: $threwWith"
$script:CimThrows = $false

# ---------------------------------------------------------------------
# Fix's own functions, lifted out by the parser and run here against the same
# fakes.
# ---------------------------------------------------------------------
$fixPath = Join-Path $Root 'Fix-ClaudeDesktop.ps1'
$fixAst  = [System.Management.Automation.Language.Parser]::ParseFile($fixPath, [ref]$null, [ref]$null)
function Get-FixFunction {
    param([string]$Name)
    return $fixAst.Find({ param($n)
        $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Stop-ClaudeDesktopApp stops the app and counts the rest"

# The two escalations after a relaunch kill through this. It has to stop the
# app's processes, leave the CLI, and say how many it left.
$stopDef = Get-FixFunction 'Stop-ClaudeDesktopApp'
Assert-Case "Stop-ClaudeDesktopApp was found in Fix" ($null -ne $stopDef)
if ($stopDef) {
    $script:ClaudeEnv = $msixEnv
    $script:FakeProcs = @(
        (New-FakeProc 101 "$install\app\Claude.exe"),
        (New-FakeProc 102 "$install\app\Claude.exe"),
        (New-FakeProc 201 $appCli),
        (New-FakeProc 202 $termCli),
        (New-FakeProc 301 $stranger),
        (New-FakeProc 401 $null "$install\app\Claude.exe"),
        (New-FakeProc 402 $null $null)
    )
    $outcome = & {
        $stoppedIds = New-Object System.Collections.ArrayList
        $logged     = New-Object System.Collections.ArrayList
        function Stop-Process {
            param([Parameter(ValueFromPipeline = $true)]$InputObject, [switch]$Force)
            process { $null = $Force; [void]$stoppedIds.Add([int]$InputObject.Id) }
        }
        function Log {
            param([string]$Message, [string]$Colour, [switch]$Indent)
            $null = $Colour, $Indent
            [void]$logged.Add($Message)
        }
        . ([scriptblock]::Create($stopDef.Extent.Text))
        $dryReturned = Stop-ClaudeDesktopApp -WhatIf
        $dryStopped  = $stoppedIds.Count
        $returned    = Stop-ClaudeDesktopApp
        [pscustomobject]@{
            DryReturned = $dryReturned
            DryStopped  = $dryStopped
            Returned    = $returned
            Stopped     = (@($stoppedIds | Sort-Object) -join ',')
            Logged      = ($logged -join ' | ')
        }
    }
    Assert-Case "under -WhatIf it stops nothing and says so"   ($outcome.DryStopped -eq 0 -and $outcome.DryReturned -eq 0) `
        "stopped $($outcome.DryStopped), returned $($outcome.DryReturned)"
    Assert-Case "it stops the app's processes and nothing else" ($outcome.Stopped -eq '101,102,401') "stopped $($outcome.Stopped)"
    Assert-Case "it returns how many it stopped"                ($outcome.Returned -eq 3) "returned $($outcome.Returned)"
    Assert-Case "it says how many it left alone" `
        ($outcome.Logged -match 'Left 4 other process\(es\) named claude alone') "logged: $($outcome.Logged)"

    # A process list that could not be read stops nothing, says so and
    # returns $null. Until 6.3.1 it read as empty: nothing stopped, nothing
    # said, and 0 returned, as for an app that was not running.
    $script:ProcListFails = $true
    try {
        $unreadStop = & {
            $stoppedIds = New-Object System.Collections.ArrayList
            $logged     = New-Object System.Collections.ArrayList
            function Stop-Process {
                param([Parameter(ValueFromPipeline = $true)]$InputObject, [switch]$Force)
                process { $null = $Force; [void]$stoppedIds.Add([int]$InputObject.Id) }
            }
            function Log {
                param([string]$Message, [string]$Colour, [switch]$Indent)
                $null = $Colour, $Indent
                [void]$logged.Add($Message)
            }
            . ([scriptblock]::Create($stopDef.Extent.Text))
            $returned = Stop-ClaudeDesktopApp
            [pscustomobject]@{ IsNull = ($null -eq $returned); Returned = "$returned"; Stopped = $stoppedIds.Count; Logged = ($logged -join ' | ') }
        }
    } catch {
        # Escaped: a failed case, not the end of the gate.
        $unreadStop = [pscustomobject]@{ IsNull = $false; Returned = "escaped: $($_.Exception.Message)"; Stopped = -1; Logged = '' }
    } finally {
        $script:ProcListFails = $false
    }
    Assert-Case "a process list that could not be read: nothing stopped, `$null returned" `
        ($unreadStop.IsNull -and $unreadStop.Stopped -eq 0) "returned '$($unreadStop.Returned)', stopped $($unreadStop.Stopped)"
    Assert-Case "and it says so" `
        ($unreadStop.Logged -match "^\[!\] Claude's processes could not be listed \(Access is denied\), so none were stopped$") "logged: $($unreadStop.Logged)"

    # A kill that Stop-Process refuses is not counted, and is said. Until
    # 6.3.1 every process listed was counted as stopped. The refusal is the
    # error the real one writes for a process it may not stop.
    $firstErrDef = Get-FixFunction 'Get-FixFirstError'
    $refusedStop = & {
        $stoppedIds = New-Object System.Collections.ArrayList
        $logged     = New-Object System.Collections.ArrayList
        function Stop-Process {
            param([Parameter(ValueFromPipeline = $true)]$InputObject, [switch]$Force)
            process {
                $null = $Force
                if ([int]$InputObject.Id -eq 102) {
                    $PSCmdlet.WriteError((New-Object System.Management.Automation.ErrorRecord(
                        (New-Object System.ComponentModel.Win32Exception 5), 'CouldNotStopProcess', 'CloseError', $InputObject)))
                    return
                }
                [void]$stoppedIds.Add([int]$InputObject.Id)
            }
        }
        function Log {
            param([string]$Message, [string]$Colour, [switch]$Indent)
            $null = $Colour, $Indent
            [void]$logged.Add($Message)
        }
        . ([scriptblock]::Create($firstErrDef.Extent.Text))
        . ([scriptblock]::Create($stopDef.Extent.Text))
        $returned = Stop-ClaudeDesktopApp
        [pscustomobject]@{ Returned = $returned; Stopped = (@($stoppedIds | Sort-Object) -join ','); Logged = ($logged -join ' | ') }
    }
    Assert-Case "a kill that is refused is not counted" `
        ($refusedStop.Returned -eq 2 -and $refusedStop.Stopped -eq '101,401') "returned $($refusedStop.Returned), stopped $($refusedStop.Stopped)"
    Assert-Case "and it says so" `
        ($refusedStop.Logged -match '\[!\] 1 Claude process\(es\) could not be stopped \(Access is denied\)') "logged: $($refusedStop.Logged)"
}

# Every kill site in Fix says how many processes named claude it spared, so a
# log reader can tell a spared CLI from one the kill missed.
$fixSrc = [System.IO.File]::ReadAllText($fixPath)
foreach ($site in @(
    @{ Name = '-Close';                 Rx = 'Log "Left \$procsSpared other process\(es\) named claude alone' }
    @{ Name = 'Step 1';                 Rx = 'Log "Leaving \$claudeSpared other process\(es\) named claude alone' }
    @{ Name = 'Step 4 orphan cleanup';  Rx = 'Log "Leaving \$orphanSpared other process\(es\) named claude alone' }
    @{ Name = 'Stop-ClaudeDesktopApp';  Rx = 'Log "Left \$spared other process\(es\) named claude alone' }
)) {
    Assert-Case "Fix $($site.Name) logs the processes it spared" ($fixSrc -match $site.Rx)
}
$helperCalls = [regex]::Matches($fixSrc, '(?m)^\s*\$null = Stop-ClaudeDesktopApp\s*$').Count
Assert-Case "Fix's two escalation kills go through Stop-ClaudeDesktopApp" ($helperCalls -eq 2) "found $helperCalls"
Assert-Case "Fix -Close stops only the desktop set" `
    ($fixSrc -match '\$procs = @\(\$procInventory \| Where-Object \{ \$_\.IsDesktop \}')
Assert-Case "Fix Step 1 stops only the desktop set" `
    ($fixSrc -match '\$claudeProcs\s+= @\(\$claudeInventory \| Where-Object \{ \$_\.IsDesktop \}')
Assert-Case "Fix Step 4 stops only the desktop set and the service" `
    ($fixSrc -match '\$remaining = @\(@\(\$orphanInventory \| Where-Object \{ \$_\.IsDesktop \}')

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Fix -Quiet holds back for the app's sessions"

# Until 6.3.0 the -Quiet gate asked only about CPU, a VM log under 120 s old
# and recent input. A project thread waiting on a 25-minute command uses no
# CPU and writes nothing, so an unattended run passed all three and ended the
# thread. Get-AutomatedFixBlocker asks about the app's sessions first, fails
# closed when they cannot be listed, then asks the monitor's transcript
# question. Everything it calls is faked inside the scope that runs it.
$blockDef = Get-FixFunction 'Get-AutomatedFixBlocker'
Assert-Case "Get-AutomatedFixBlocker was found in Fix" ($null -ne $blockDef)
# Since 6.3.0 the question about the app's sessions is its own function, so
# the check before each kill of the app asks the same one.
$sessionDef = Get-FixFunction 'Get-AppSessionBlocker'
if ($blockDef) {
    $blocked = & {
        Set-StrictMode -Version Latest
        $ProcessName      = 'claude'
        $script:ClaudeEnv = [pscustomobject]@{ ProcessName = 'claude'; LogDir = 'C:\nowhere' }
        $appCliNow        = @()
        $appCliThrows     = $false
        $transcriptRecent = $false
        $busyProc         = $null
        $vmLogAge         = $null
        $transcriptThrows = $false
        $vmLogThrows      = $false
        $procListFails    = $false
        $idFails          = $false
        $sessionOne       = $false
        $addTypeFails     = ''
        $transcriptAsked  = New-Object System.Collections.ArrayList
        function Get-ClaudeAppCliProcess {
            param($Environment)
            $null = $Environment
            if ($appCliThrows) { throw 'Access denied' }
            return @($appCliNow)
        }
        # Both throw when they cannot tell, since 6.3.1, as the region's do.
        function Test-ClaudeCodeTranscriptRecent {
            param([int]$Minutes)
            [void]$transcriptAsked.Add($Minutes)
            if ($transcriptThrows) { throw 'Access to the path is denied' }
            return $transcriptRecent
        }
        function Get-ClaudeVmLogAge {
            param($LogDir)
            $null = $LogDir
            if ($vmLogThrows) { throw 'Access to the path is denied' }
            return $vmLogAge
        }
        function Start-Sleep { param([int]$Milliseconds, [int]$Seconds) $null = $Milliseconds, $Seconds }
        # As the cmdlet: no match writes NoProcessFoundForGivenName,
        # ProcessCommandException, ObjectNotFound (measured 2026-11-22), and a
        # list that cannot be read another error; neither throws, and
        # -ErrorAction Stop makes it terminating in the caller.
        function Get-Process {
            [CmdletBinding()]
            param([string]$Name, [int]$Id)
            if ($Id) {
                if ($idFails) {
                    $PSCmdlet.WriteError((New-FakeErrorRecord 'System.ComponentModel.Win32Exception' 'Access is denied' `
                        'CouldNotEnumerateProcesses' 'NotSpecified' $Id))
                    return
                }
                return [pscustomobject]@{ Id = $Id; SessionId = $(if ($sessionOne) { 1 } else { 0 }) }
            }
            if ($procListFails) {
                $PSCmdlet.WriteError((New-FakeErrorRecord 'System.ComponentModel.Win32Exception' 'Access is denied' `
                    'CouldNotEnumerateProcesses' 'NotSpecified' $Name))
                return
            }
            if ($Name -eq 'claude' -and $busyProc) { return @($busyProc) }
            $PSCmdlet.WriteError((New-FakeErrorRecord 'Microsoft.PowerShell.Commands.ProcessCommandException' `
                "Cannot find a process with the name `"$Name`". Verify the process name and call the cmdlet again." `
                'NoProcessFoundForGivenName' 'ObjectNotFound' $Name))
        }
        # As the cmdlet (measured 2026-11-23 on Windows PowerShell 5.1): a
        # definition that does not compile writes one SOURCE_CODE_ERROR per
        # compiler error, which -ErrorAction Stop makes terminating, and then
        # throws COMPILER_ERRORS whatever the ErrorAction; a name another
        # definition holds throws TYPE_ALREADY_EXISTS whatever the
        # ErrorAction. Otherwise it adds nothing: the type the input check
        # calls is a stand-in, loaded below.
        function Add-Type {
            [CmdletBinding()]
            param([string]$TypeDefinition)
            $null = $TypeDefinition
            if ($addTypeFails -eq 'compile') {
                $PSCmdlet.WriteError((New-FakeErrorRecord 'System.Exception' `
                    ("c:\Temp\cfx1\cfx1.0.cs(6) : Invalid expression term '}'`r`n`r`nc:\Temp\cfx1\cfx1.0.cs(6) : >>> " +
                        '[DllImport("user32.dll")] public static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);') `
                    'SOURCE_CODE_ERROR' 'InvalidData' $null))
                $PSCmdlet.ThrowTerminatingError((New-FakeErrorRecord 'System.InvalidOperationException' `
                    'Cannot add type. Compilation errors occurred.' 'COMPILER_ERRORS' 'InvalidData' $null))
            }
            if ($addTypeFails -eq 'exists') {
                $PSCmdlet.ThrowTerminatingError((New-FakeErrorRecord 'System.Exception' `
                    "Cannot add type. The type name 'FixActivityCheck' already exists." 'TYPE_ALREADY_EXISTS' 'InvalidOperation' $null))
            }
        }
        if ($sessionDef) { . ([scriptblock]::Create($sessionDef.Extent.Text)) }
        . ([scriptblock]::Create($blockDef.Extent.Text))

        $r = [ordered]@{}
        $appCliNow = @([pscustomobject]@{ Id = 201 }, [pscustomobject]@{ Id = 203 })
        $transcriptRecent = $true
        $r['cli'] = Get-AutomatedFixBlocker
        $appCliNow = @()
        $appCliThrows = $true
        $r['throws'] = Get-AutomatedFixBlocker
        $appCliThrows = $false
        $r['transcript'] = Get-AutomatedFixBlocker
        $r['minutes'] = (@($transcriptAsked) -join ',')
        $transcriptRecent = $false
        $busyProc = [pscustomobject]@{ Id = 77; Ms = 0 }
        $busyProc | Add-Member -MemberType ScriptProperty -Name TotalProcessorTime -Value { [timespan]::FromMilliseconds($this.Ms) }
        $busyProc | Add-Member -MemberType ScriptMethod -Name Refresh -Value { $this.Ms += 400 }
        $r['cpu'] = Get-AutomatedFixBlocker
        $busyProc = $null
        $vmLogAge = 30.4
        $r['vmlog'] = Get-AutomatedFixBlocker
        $vmLogAge = 600
        $r['clear'] = Get-AutomatedFixBlocker
        # Each read that could not be made, since 6.3.1. A throw out of the
        # function is a case result too, so an old copy fails here and not
        # the whole run.
        $ask = { try { Get-AutomatedFixBlocker } catch { "THREW: $($_.Exception.Message)" } }
        $transcriptThrows = $true
        $r['transcriptUnread'] = & $ask
        $transcriptThrows = $false
        $procListFails = $true
        $r['procsUnread'] = & $ask
        $procListFails = $false
        $vmLogThrows = $true
        $r['vmlogUnread'] = & $ask
        $vmLogThrows = $false
        $idFails = $true
        $r['inputUnread'] = & $ask
        $idFails = $false
        # A process whose second sample fails: running, then whose end
        # cannot be read either, then ended. As a Process: a getter that
        # throws reads as nothing, and its get_ method throws (measured
        # 2026-11-23).
        $busyProc = [pscustomobject]@{ Id = 78; Gone = $false; EndUnread = $false }
        $busyProc | Add-Member -MemberType ScriptProperty -Name TotalProcessorTime -Value { [timespan]::FromMilliseconds(0) }
        $busyProc | Add-Member -MemberType ScriptMethod -Name Refresh -Value { throw 'Access is denied' }
        $busyProc | Add-Member -MemberType ScriptProperty -Name HasExited -Value {
            if ($this.EndUnread) { throw 'Access is denied' }
            return $this.Gone
        }
        $busyProc | Add-Member -MemberType ScriptMethod -Name get_HasExited -Value {
            if ($this.EndUnread) { throw 'Access is denied' }
            return $this.Gone
        }
        $r['sampleUnread'] = & $ask
        $busyProc.EndUnread = $true
        $r['sampleEndUnread'] = & $ask
        $busyProc.EndUnread = $false
        $busyProc.Gone = $true
        $r['sampleEnded'] = & $ask
        $busyProc = $null
        # Check 5, in a session with a desktop. Its type comes from the
        # compiler, so the case where it is not there runs first, and a
        # stand-in of that name is loaded after it: GetLastInputInfo answers
        # an input IdleMs ago, or fails.
        $r['typeBefore'] = [bool]('FixActivityCheck' -as [type])
        $sessionOne = $true
        $addTypeFails = 'compile'
        $r['inputTypeUnloaded'] = & $ask
        $addTypeFails = ''
        Microsoft.PowerShell.Utility\Add-Type -ErrorAction Stop -TypeDefinition @'
using System;
public static class FixActivityCheck {
    public struct LASTINPUTINFO { public uint cbSize; public uint dwTime; }
    public static uint IdleMs = 600000;
    public static bool Fails = false;
    public static bool GetLastInputInfo(ref LASTINPUTINFO plii) {
        if (Fails) { return false; }
        plii.dwTime = unchecked((uint)Environment.TickCount - IdleMs);
        return true;
    }
}
'@
        $standIn = 'FixActivityCheck' -as [type]
        $standIn::IdleMs = 10000
        $r['inputRecent'] = & $ask
        $standIn::IdleMs = 600000
        $r['inputIdle'] = & $ask
        $addTypeFails = 'exists'
        $r['inputTypeThere'] = & $ask
        $addTypeFails = ''
        $standIn::Fails = $true
        $r['inputApiFails'] = & $ask
        $standIn::Fails = $false
        $sessionOne = $false
        $r
    }
    Assert-Case "a running app session blocks, naming its PIDs" `
        ("$($blocked['cli'])" -match '^2 Claude Code session\(s\) the app started are running \(PID 201, 203\)$') "got: $($blocked['cli'])"
    Assert-Case "sessions that cannot be listed block (fail closed)" `
        ("$($blocked['throws'])" -match '^the Claude Code sessions the app started could not be listed \(Access denied\)$') "got: $($blocked['throws'])"
    Assert-Case "a transcript written in the last 35 minutes blocks" `
        ("$($blocked['transcript'])" -match 'transcript was written in the last 35 minutes') "got: $($blocked['transcript'])"
    Assert-Case "the transcript question is asked with 35 minutes" `
        ($blocked['minutes'] -match '^35(,35)*$') "asked with: $($blocked['minutes'])"
    Assert-Case "a busy claude process still blocks" `
        ("$($blocked['cpu'])" -eq 'claude process 77 is using CPU') "got: $($blocked['cpu'])"
    Assert-Case "a fresh VM log still blocks" `
        ("$($blocked['vmlog'])" -eq 'a VM log was written 30s ago') "got: $($blocked['vmlog'])"
    Assert-Case "nothing going on: no blocker" ($null -eq $blocked['clear']) "got: $($blocked['clear'])"
    Assert-Case "a transcript folder that could not be read blocks (fail closed)" `
        ("$($blocked['transcriptUnread'])" -eq 'whether a Claude Code transcript was written in the last 35 minutes could not be read (Access to the path is denied)') `
        "got: $($blocked['transcriptUnread'])"
    Assert-Case "a process list that could not be read blocks (fail closed)" `
        ("$($blocked['procsUnread'])" -eq 'the processes named claude could not be listed (Access is denied)') "got: $($blocked['procsUnread'])"
    Assert-Case "VM logs that could not be read block (fail closed)" `
        ("$($blocked['vmlogUnread'])" -eq 'the VM logs could not be read (Access to the path is denied)') "got: $($blocked['vmlogUnread'])"
    Assert-Case "an input check that failed blocks (fail closed)" `
        ("$($blocked['inputUnread'])" -eq 'the time of the last keyboard or mouse input could not be read (Access is denied)') "got: $($blocked['inputUnread'])"
    Assert-Case "a running process that could not be sampled blocks" `
        ("$($blocked['sampleUnread'])" -match '^whether claude process 78 is using CPU could not be read \(.*Access is denied.*\)$') "got: $($blocked['sampleUnread'])"
    Assert-Case "one that ended before its second sample does not" ($null -eq $blocked['sampleEnded']) "got: $($blocked['sampleEnded'])"
    # Until 6.3.1 the end was read through the getter, which reads as
    # nothing when it throws, so an end that could not be read was not said.
    Assert-Case "one whose end could not be read either blocks, saying both" `
        ("$($blocked['sampleEndUnread'])" -match '^whether claude process 78 is using CPU could not be read \(.*Access is denied.*\), nor whether it has ended \(.*Access is denied.*\)$') `
        "got: $($blocked['sampleEndUnread'])"
    Assert-Case "no stand-in input type was loaded before the input cases" ($blocked['typeBefore'] -eq $false) "got: $($blocked['typeBefore'])"
    # Until 6.3.1 the Add-Type ran with SilentlyContinue, which hid the
    # compiler's error, and only "Cannot add type. Compilation errors
    # occurred." was said.
    Assert-Case "an input check whose type did not compile blocks, naming the compiler's error" `
        ("$($blocked['inputTypeUnloaded'])" -match "^the time of the last keyboard or mouse input could not be read \(the check for it could not be loaded: .*Invalid expression term '\}'.*\)$") `
        "got: $($blocked['inputTypeUnloaded'])"
    Assert-Case "input 10 seconds ago blocks" ("$($blocked['inputRecent'])" -match '^keyboard or mouse input 1[01]s ago$') "got: $($blocked['inputRecent'])"
    Assert-Case "no input for 10 minutes: no blocker" ($null -eq $blocked['inputIdle']) "got: $($blocked['inputIdle'])"
    Assert-Case "an Add-Type that fails with the type already there: the type answers, no blocker" `
        ($null -eq $blocked['inputTypeThere']) "got: $($blocked['inputTypeThere'])"
    Assert-Case "a GetLastInputInfo that fails blocks" `
        ("$($blocked['inputApiFails'])" -eq 'the time of the last keyboard or mouse input could not be read (GetLastInputInfo failed)') `
        "got: $($blocked['inputApiFails'])"
}

# Fix's window type, loaded at the top of the script, is best effort: with
# it the window is brought forward and flashed, and nothing else. Add-Type
# throws COMPILER_ERRORS for a type that does not compile whatever its
# ErrorAction (measured 2026-11-23), so until 6.3.1 SilentlyContinue alone
# left that error in red on the screen of a run it does not touch. The
# statement that loads it runs here against a fake Add-Type that fails as
# the cmdlet does.
$winAdd = $fixAst.Find({ param($n)
    $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Add-Type' -and
    $n.Extent.Text -match 'public static class Win32Window' }, $true)
Assert-Case "Fix's window type is loaded by an Add-Type" ($null -ne $winAdd)
if ($winAdd) {
    $winStmt = $winAdd
    while ($winStmt.Parent -and $winStmt.Parent -isnot [System.Management.Automation.Language.NamedBlockAst]) { $winStmt = $winStmt.Parent }
    $winOut = try {
        & {
            Set-StrictMode -Version Latest
            function Add-Type {
                [CmdletBinding()]
                param([string]$TypeDefinition)
                $null = $TypeDefinition
                $PSCmdlet.WriteError((New-FakeErrorRecord 'System.Exception' 'c:\Temp\cfx2\cfx2.0.cs(12) : ; expected' 'SOURCE_CODE_ERROR' 'InvalidData' $null))
                $PSCmdlet.ThrowTerminatingError((New-FakeErrorRecord 'System.InvalidOperationException' `
                    'Cannot add type. Compilation errors occurred.' 'COMPILER_ERRORS' 'InvalidData' $null))
            }
            . ([scriptblock]::Create($winStmt.Extent.Text))
            'ran on'
        } 2>&1
    } catch { "THREW: $($_.Exception.Message)" }
    $winText = @($winOut | ForEach-Object { "$_" }) -join ' / '
    Assert-Case "a window type that does not compile is passed over in silence, and the run goes on" ($winText -eq 'ran on') "got: $winText"
}
# Held back, not done: exit $script:BlockedExitCode. Until 6.3.0 it was exit 0,
# the same as a finished repair. Test-RunOutcomes runs the gate.
$gateRx = '(?ms)^if \(\$Quiet\) \{\s*\r?\n\s*\$blocker = Get-AutomatedFixBlocker\s*\r?\n\s*if \(\$blocker\) \{\s*\r?\n\s*Log "BLOCKED: \$blocker, skipping automated fix"[^\r\n]*\r?\n\s*Save-Log\s*\r?\n(\s*#[^\r\n]*\r?\n)*\s*exit \$script:BlockedExitCode\s*\r?\n'
Assert-Case "Fix's -Quiet gate exits through Get-AutomatedFixBlocker, as held back" ($fixSrc -match $gateRx)
$blockerCalls = [regex]::Matches($fixSrc, '(?m)^\s*\$blocker = Get-AutomatedFixBlocker\s*$').Count
Assert-Case "and the gate at the start asks it once" ($blockerCalls -eq 1) "found $blockerCalls call(s)"

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Fix -Quiet asks again before each kill of the app"

# The gate asks at the start, and a Smart run takes minutes: the app can
# start a Code session or a project thread in that time. Until 6.3.0 nothing
# asked again, and the next kill of the app ended it. Assert-NoNewAppSession
# asks right before Step 1's kill, Step 4's kill of the app's leftover
# processes and each escalation that stops the app. Before Step 1 nothing
# that stops the app or the service has run, so a session found there ends
# the run as held back, exit 3. After it Steps 1 and 2 have run, so a session
# found ends the run as stopped part-way, exit 4, and the service is started
# again first when this run stopped it. Until 6.3.0 Step 4 was not asked, and
# every guard ended the run as held back.
$fixLines = $fixSrc -split "\r?\n"
$guardLineRx = '^(if \(\$orphanAppCount -gt 0\) \{ )?Assert-NoNewAppSession -Before "([^"]+)" -ExitCode \$script:(BlockedExitCode|StoppedPartWayExitCode)( \})?$'
function Get-GuardBefore {
    # The guard call on the nearest code line above line $At, or $null when
    # that line is anything else; comment and blank lines in between are
    # skipped.
    param([int]$At)
    for ($k = $At - 1; $k -ge 0 -and $k -ge $At - 8; $k--) {
        $t = $fixLines[$k].Trim()
        if ($t -eq '' -or $t.StartsWith('#')) { continue }
        if ($t -match $guardLineRx) { return [pscustomobject]@{ Before = $Matches[2]; Code = $Matches[3]; Line = $t } }
        return $null
    }
    return $null
}
$killSites = @()
for ($k = 0; $k -lt $fixLines.Count; $k++) {
    if ($fixLines[$k] -match '^\s*\$null = Stop-ClaudeDesktopApp\s*$') { $killSites += $k }
    if ($fixLines[$k] -match '^\s*if \(\$PSCmdlet\.ShouldProcess\("\$\(\$claudeProcs\.Count\) Claude process\(es\)", "Stop"\)\) \{\s*$') { $killSites += $k }
    if ($fixLines[$k] -match '^\s*if \(\$PSCmdlet\.ShouldProcess\("\$\(\$remaining\.Count\) orphan process\(es\)", "Force-kill"\)\) \{\s*$') { $killSites += $k }
}
Assert-Case "four kills of the app found (Step 1, Step 4 and the two escalations)" ($killSites.Count -eq 4) "found $($killSites.Count)"
$guardedBy = @()
foreach ($k in $killSites) {
    $g = Get-GuardBefore -At $k
    Assert-Case "line $($k + 1) asks Assert-NoNewAppSession right before it" ($null -ne $g) $fixLines[$k].Trim()
    if ($g) {
        $want = if ($g.Before -eq 'Step 1') { 'BlockedExitCode' } else { 'StoppedPartWayExitCode' }
        Assert-Case "  and a session found before $($g.Before) ends the run with `$script:$want" ($g.Code -eq $want) $g.Line
        $guardedBy += $g.Before
    }
}
$guardedSeen = (@($guardedBy | Sort-Object) -join ' | ')
$guardedWant = (@('Step 1', 'Step 4', 'the escalation to Deep', 'the no-progress escalation') | Sort-Object) -join ' | '
Assert-Case "the guards stand before Step 1, Step 4 and both escalations" ($guardedSeen -eq $guardedWant) "before: $guardedSeen"
$guardCalls = [regex]::Matches($fixSrc, '(?m)^\s*(if \(\$orphanAppCount -gt 0\) \{ )?Assert-NoNewAppSession -Before "').Count
# Once more before Step 1, in its branch for a list of processes that could
# not be read: whether the app runs a session is not known then either, and
# nothing has been stopped yet (6.3.1).
Assert-Case "and nowhere else, but once more in Step 1 for a list that could not be read" ($guardCalls -eq 5) "found $guardCalls call(s)"
Assert-Case "that one is Step 1's branch for a list that could not be read" `
    ($fixSrc -match '(?m)^\} elseif \(\$claudeUnread\) \{\s*\r?\n(\s*#[^\r\n]*\r?\n)*\s*Assert-NoNewAppSession -Before "Step 1" -ExitCode \$script:BlockedExitCode\s*\r?\n\} else \{')
# Step 4 asks only when a process of the app is left over, not for the
# service alone, which no Claude Code session runs in.
Assert-Case "Step 4 counts the app's own processes it is about to stop" `
    ($fixSrc -match '(?m)^\$orphanAppCount\s+= @\(\$orphanInventory \| Where-Object \{ \$_\.IsDesktop \}\)\.Count\s*$')

# Run in a child process, because it ends the run with exit. Everything it
# calls is faked there; the exit code is the outcome.
$guardDef   = Get-FixFunction 'Assert-NoNewAppSession'
$codeLine   = [regex]::Match($fixSrc, '(?m)^\$script:BlockedExitCode = \d+\s*$').Value
$partLine   = [regex]::Match($fixSrc, '(?m)^\$script:StoppedPartWayExitCode = \d+\s*$').Value
Assert-Case "Fix sets both codes the guard ends a run with" ($codeLine -ne '' -and $partLine -ne '') "[$($codeLine.Trim())] [$($partLine.Trim())]"
Assert-Case "Assert-NoNewAppSession and Get-AppSessionBlocker were found in Fix" ($null -ne $guardDef -and $null -ne $sessionDef)
# The service as Step 2 left it: stopped by this run, unless the case says
# 'notstopped' (it was not running then), or running again ('svcup'), as at
# an escalation after Step 7. Start-Service only says what it was handed,
# and throws for 'startfails'.
$restartDef = Get-FixFunction 'Start-ServiceStoppedByRun'
Assert-Case "Start-ServiceStoppedByRun was found in Fix" ($null -ne $restartDef)
$restartText = if ($restartDef) { $restartDef.Extent.Text } else { '' }
# The service is read through Fix's Get-FixServiceState and the region's
# Get-ClaudeServiceOrNull, lifted as they are; only Get-Service is faked.
# 'svcdenied' makes it fail with an identity that is not "no such service"
# (access denied), as a read that could not be made.
$svcReadText = @('Get-FixServiceState', 'Get-ClaudeServiceOrNull', 'Test-ClaudeNotFoundError' | ForEach-Object {
    $d = Get-FixFunction $_
    if ($d) { $d.Extent.Text } }) -join "`r`n"
$svcFakes = @(
    '$ServiceName = "CoworkVMService"',
    '$script:ServiceStoppedByRun = ($Case -notmatch ''notstopped'')',
    'function Get-ClaudeErrorText { param($ErrorRecord) return [string]$ErrorRecord.Exception.Message }',
    'function Get-Service {',
    '    [CmdletBinding()]',
    '    param([string]$Name)',
    '    if ($Case -match ''svcdenied'') {',
    '        $PSCmdlet.WriteError((New-Object System.Management.Automation.ErrorRecord((New-Object System.ComponentModel.Win32Exception 5), ''CouldNotGetServiceInfo'', ''PermissionDenied'', $Name)))',
    '        return',
    '    }',
    '    $st = if ($Case -match ''svcup'') { ''Running'' } else { ''Stopped'' }',
    '    return [pscustomobject]@{ Name = $Name; Status = $st }',
    '}',
    $svcReadText,
    'function Start-Service {',
    '    param([string]$Name, [string]$ErrorAction)',
    '    $null = $ErrorAction',
    '    if ($Case -match ''startfails'') { throw "Cannot start service $Name" }',
    '    [Console]::Out.WriteLine("STARTED $Name")',
    '}',
    $restartText
) -join "`r`n"
# Every place Fix stops the service notes it, so the guard knows the stop
# was this run's: Step 2, inside its ShouldProcess, Restart-CoworkService
# and Stop-CoworkServiceWithTimeout.
Assert-Case "Fix starts each run with `$script:ServiceStoppedByRun false" `
    ($fixSrc -match '(?m)^\$script:ServiceStoppedByRun = \$false\s*$')
$s2Stop = @($fixAst.FindAll({ param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and
    $n.Clauses[0].Item1.Extent.Text -eq '$PSCmdlet.ShouldProcess($ServiceName, "Stop")' }, $true))
Assert-Case "Step 2 notes that it stopped the service" `
    ($s2Stop.Count -eq 1 -and $s2Stop[0].Clauses[0].Item2.Extent.Text -match '(?m)^\s*\$script:ServiceStoppedByRun = \$true\s*$') `
    "found $($s2Stop.Count) stop block(s)"
foreach ($fn in 'Restart-CoworkService', 'Stop-CoworkServiceWithTimeout') {
    $fnDef = Get-FixFunction $fn
    Assert-Case "and so does $fn" ($null -ne $fnDef -and $fnDef.Extent.Text -match '(?m)^\s*\$script:ServiceStoppedByRun = \$true\s*$')
}
$guardFakes = @(
    $svcFakes,
    '$script:ClaudeEnv = [pscustomobject]@{ ProcessName = ''claude'' }',
    'function Log { param([string]$Message, [string]$Colour, [switch]$Indent) $null = $Colour, $Indent; Write-Output "LOG: $Message" }',
    'function Save-Log { Write-Output ''SAVED'' }',
    'function Get-ClaudeAppCliProcess {',
    '    param($Environment)',
    '    $null = $Environment',
    '    if ($Case -match ''throws'') { throw ''Access denied'' }',
    '    if ($Case -match ''session'') { return @([pscustomobject]@{ Id = 301 }) }',
    '    return @()',
    '}'
) -join "`r`n"
if ($guardDef -and $sessionDef) {
    $guardScript = Join-Path ([System.IO.Path]::GetTempPath()) ('cf-guard-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.ps1')
    $guardBody = @(
        'param([string]$Case, [string]$Before, [string]$Code, [switch]$Quiet)',
        'Set-StrictMode -Version Latest',
        $codeLine,
        $partLine,
        $guardFakes,
        $sessionDef.Extent.Text,
        $guardDef.Extent.Text,
        '$exitWith = if ($Code -eq ''Part'') { $script:StoppedPartWayExitCode } else { $script:BlockedExitCode }',
        'Assert-NoNewAppSession -Before $Before -ExitCode $exitWith',
        "Write-Output 'RETURNED'",
        'exit 0'
    ) -join "`r`n"
    [System.IO.File]::WriteAllText($guardScript, $guardBody, [System.Text.Encoding]::ASCII)
    try {
        foreach ($gc in @(
            @{ Case = 'session'; Before = 'Step 1'; Code = 'Blocked'; Quiet = $true; Exit = 3; Not = 'STOPPED PART-WAY|STARTED CoworkVMService'
               Rx = 'LOG: BLOCKED: 1 Claude Code session\(s\) the app started are running \(PID 301\); Step 1 would stop the app, so the automated fix ends here' },
            @{ Case = 'throws';  Before = 'Step 1'; Code = 'Blocked'; Quiet = $true; Exit = 3; Not = 'STOPPED PART-WAY|STARTED CoworkVMService'
               Rx = 'LOG: BLOCKED: the Claude Code sessions the app started could not be listed \(Access denied\)' },
            @{ Case = 'session'; Before = 'Step 4'; Code = 'Part'; Quiet = $true; Exit = 4; Not = 'BLOCKED'
               Rx = 'LOG: STOPPED PART-WAY: 1 Claude Code session\(s\) the app started are running \(PID 301\); Step 4 would stop the app, so the automated fix ends here, after the steps it has run' },
            @{ Case = 'throws';  Before = 'the escalation to Deep'; Code = 'Part'; Quiet = $true; Exit = 4; Not = 'BLOCKED'
               Rx = 'LOG: STOPPED PART-WAY: the Claude Code sessions the app started could not be listed \(Access denied\); the escalation to Deep would stop the app' },
            @{ Case = 'none';    Before = 'Step 4'; Code = 'Part'; Quiet = $true;  Exit = 0; Not = 'BLOCKED|STOPPED|STARTED CoworkVMService'; Rx = 'RETURNED' },
            @{ Case = 'session'; Before = 'Step 4'; Code = 'Part'; Quiet = $false; Exit = 0; Not = 'BLOCKED|STOPPED|STARTED CoworkVMService'; Rx = 'RETURNED' }
        )) {
            $gArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $guardScript, '-Case', $gc.Case, '-Before', $gc.Before, '-Code', $gc.Code)
            if ($gc.Quiet) { $gArgs += '-Quiet' }
            $gOut = (& powershell.exe @gArgs 2>&1 | Out-String)
            $gExit = $LASTEXITCODE
            $label = "$($gc.Case) before $($gc.Before)$(if ($gc.Quiet) { ', -Quiet' } else { ', manual' })"
            Assert-Case "$label exits $($gc.Exit)" ($gExit -eq $gc.Exit -and $gOut -match $gc.Rx -and $gOut -notmatch $gc.Not) "exit $gExit; output: $($gOut.Trim())"
        }
    } finally {
        # Clean-up of this section's own temp script: its cases are counted,
        # and a file left in %TEMP% changes nothing.
        Remove-Item -LiteralPath $guardScript -Force -ErrorAction SilentlyContinue
    }
}

# Step 4 itself, lifted from its inventory of the app's processes to the end
# of its kill and run in a child process. Every command it calls is faked
# there, and the kill is a fake Stop-Process that only says which Id it was
# handed; ShouldProcess is a stand-in that says it was asked. Step 2 ran
# before it, so the service is stopped and this run stopped it, unless the
# case says otherwise. Until 6.3.0 an exit 4 here left it stopped, and
# Windows does not start a manually stopped service through its trigger.
$s4First = @($fixAst.FindAll({ param($n)
    $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$orphanUnread' -and
    $n.Right.Extent.Text -eq '$null' }, $true))
$s4Last  = @($fixAst.FindAll({ param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -eq '$remaining.Count -gt 0' }, $true))
Assert-Case "Step 4's inventory and its kill were found" ($s4First.Count -eq 1 -and $s4Last.Count -eq 1) "found $($s4First.Count), $($s4Last.Count)"
if ($s4First.Count -eq 1 -and $s4Last.Count -eq 1 -and $guardDef -and $sessionDef) {
    $s4Text = $fixAst.Extent.Text.Substring($s4First[0].Extent.StartOffset, $s4Last[0].Extent.EndOffset - $s4First[0].Extent.StartOffset)
    $s4Swaps = ([regex]::Matches($s4Text, [regex]::Escape('$PSCmdlet.ShouldProcess('))).Count
    $s4Text = $s4Text.Replace('$PSCmdlet.ShouldProcess(', '$fakeCmdlet.ShouldProcess(')
    $s4Faked = @('Get-ClaudeProcessInventory', 'Get-Process', 'Log', 'Save-Log', 'Stop-Process', 'Start-Sleep',
        'Get-ClaudeDesktopProcess', 'Get-ClaudeAppCliProcess', 'Assert-NoNewAppSession', 'Get-AppSessionBlocker',
        'Where-Object', 'ForEach-Object', 'Start-ServiceStoppedByRun', 'Get-Service', 'Start-Service', 'Get-ClaudeErrorText',
        'Get-ClaudeProcessByName', 'Get-FixServiceState')
    $s4Parsed = [System.Management.Automation.Language.Parser]::ParseInput($s4Text, [ref]$null, [ref]$null)
    $s4Unfaked = @(@($s4Parsed, $guardDef, $sessionDef, $restartDef) | Where-Object { $_ } | ForEach-Object {
        $_.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) } | ForEach-Object {
        $name = $_.GetCommandName()
        if (-not $name) { "<called through $($_.CommandElements[0].Extent.Text)>" } elseif ($s4Faked -notcontains $name) { $name }
    } | Sort-Object -Unique)
    Assert-Case "Step 4's ShouldProcess was swapped for the stand-in, and every command it calls is faked" `
        ($s4Swaps -eq 1 -and $s4Unfaked.Count -eq 0) "swapped $s4Swaps; not faked: $($s4Unfaked -join ', ')"
    if ($s4Swaps -eq 1 -and $s4Unfaked.Count -eq 0) {
        $s4Script = Join-Path ([System.IO.Path]::GetTempPath()) ('cf-step4-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.ps1')
        $s4Body = @(
            'param([string]$Case, [switch]$Quiet)',
            'Set-StrictMode -Version Latest',
            $codeLine,
            $partLine,
            $guardFakes,
            '$ServiceExe = "cowork-svc"',
            'function Start-Sleep { param([int]$Seconds, [int]$Milliseconds) $null = $Seconds, $Milliseconds }',
            'function Get-ClaudeErrorText { param($ErrorRecord) return [string]$ErrorRecord.Exception.Message }',
            'function Get-ClaudeProcessInventory {',
            '    param($Environment)',
            '    $null = $Environment',
            '    if ($Case -match ''unread'') { throw ''Access is denied'' }',
            '    if ($Case -match ''desktop'') { return @([pscustomobject]@{ IsDesktop = $true; Process = [pscustomobject]@{ Id = 101 } }) }',
            '    return @()',
            '}',
            'function Get-Process {',
            '    [CmdletBinding()]',
            '    param([string]$Name)',
            '    if ($Case -match ''svcexefails'') {',
            '        $PSCmdlet.WriteError((New-Object System.Management.Automation.ErrorRecord((New-Object System.ComponentModel.Win32Exception 5), ''CouldNotEnumerateProcesses'', ''PermissionDenied'', $Name)))',
            '        return',
            '    }',
            '    if ($Case -match ''service'' -and $Name -eq $ServiceExe) { return @([pscustomobject]@{ Id = 900 }) }',
            '    return @()',
            '}',
            'function Get-ClaudeProcessByName {',
            '    param([string]$Name)',
            '    if ($Case -match ''svcexefails'') { throw (New-Object System.ComponentModel.Win32Exception 5) }',
            '    if ($Case -match ''service'' -and $Name -eq $ServiceExe) { return @([pscustomobject]@{ Id = 900 }) }',
            '    return @()',
            '}',
            'function Stop-Process { param([switch]$Force, [string]$ErrorAction) $null = $Force, $ErrorAction; foreach ($p in $input) { [Console]::Out.WriteLine("KILLED $($p.Id)") } }',
            'function Get-ClaudeDesktopProcess {',
            '    param($Environment)',
            '    $null = $Environment',
            '    if ($Case -match ''recheck'') { throw ''Access is denied'' }',
            '    return @()',
            '}',
            '$fakeCmdlet = New-Object psobject',
            '$fakeCmdlet | Add-Member -MemberType ScriptMethod -Name ShouldProcess -Value { param($Target, $Action) [Console]::Out.WriteLine("ASKED $Action"); return $true }',
            $sessionDef.Extent.Text,
            $guardDef.Extent.Text,
            $s4Text,
            "Write-Output 'STEP 4 DONE'",
            'exit 0'
        ) -join "`r`n"
        [System.IO.File]::WriteAllText($s4Script, $s4Body, [System.Text.Encoding]::ASCII)
        try {
            foreach ($sc in @(
                # Until 6.3.0 this one went straight on to the kill. It
                # starts the service Step 2 stopped again before it exits.
                @{ Case = 'desktop-session'; Quiet = $true;  Exit = 4
                   Has = @('STOPPED PART-WAY: 1 Claude Code session\(s\) the app started are running \(PID 301\); Step 4 would stop the app',
                           'STARTED CoworkVMService', 'LOG: Started CoworkVMService again, which this run had stopped')
                   Not = 'ASKED|KILLED|STEP 4 DONE|Could not start' },
                # The other paths start nothing.
                @{ Case = 'desktop';         Quiet = $true;  Exit = 0; Has = @('ASKED Force-kill', 'KILLED 101', 'STEP 4 DONE'); Not = 'STOPPED|BLOCKED|STARTED CoworkVMService' },
                @{ Case = 'service-session'; Quiet = $true;  Exit = 0; Has = @('ASKED Force-kill', 'KILLED 900', 'STEP 4 DONE'); Not = 'STOPPED|BLOCKED|KILLED 101|STARTED CoworkVMService' },
                @{ Case = 'desktop-session'; Quiet = $false; Exit = 0; Has = @('ASKED Force-kill', 'KILLED 101', 'STEP 4 DONE'); Not = 'STOPPED|BLOCKED|STARTED CoworkVMService' },
                # Not this run's stop, or running already: nothing to start.
                @{ Case = 'desktop-session-notstopped'; Quiet = $true; Exit = 4; Has = @('STOPPED PART-WAY: 1 Claude Code session')
                   Not = 'ASKED|KILLED|STEP 4 DONE|STARTED CoworkVMService|Could not start' },
                @{ Case = 'desktop-session-svcup'; Quiet = $true; Exit = 4; Has = @('STOPPED PART-WAY: 1 Claude Code session')
                   Not = 'ASKED|KILLED|STEP 4 DONE|STARTED CoworkVMService|Could not start' },
                # A start that fails is logged, and the run still ends with 4.
                @{ Case = 'desktop-session-startfails'; Quiet = $true; Exit = 4
                   Has = @('STOPPED PART-WAY: 1 Claude Code session', 'LOG: \[!\] Could not start CoworkVMService again \(Cannot start service CoworkVMService\); run Fix when Cowork is needed')
                   Not = 'ASKED|KILLED|STEP 4 DONE|STARTED CoworkVMService' },
                # A list that could not be read: none of the app's processes
                # is stopped, that is said, and "All clear" is not. The
                # service's are read apart and still stopped. Until 6.3.1 the
                # list read as empty.
                @{ Case = 'unread'; Quiet = $false; Exit = 0
                   Has = @('LOG: \[!\] Claude''s processes could not be listed \(Access is denied\), so none of the app''s were stopped', 'STEP 4 DONE')
                   Not = 'ASKED|KILLED|All clear' },
                @{ Case = 'unread-service'; Quiet = $true; Exit = 0
                   Has = @('could not be listed \(Access is denied\)', 'ASKED Force-kill', 'KILLED 900', 'STEP 4 DONE')
                   Not = 'KILLED 101|All clear|STOPPED PART-WAY|BLOCKED' },
                # A re-check that could not be read is not "Cleaned up".
                @{ Case = 'desktop-recheck'; Quiet = $false; Exit = 0
                   Has = @('KILLED 101', 'LOG: \[!\] Whether the orphans stopped is not known: Claude''s processes could not be listed \(Access is denied\)', 'STEP 4 DONE')
                   Not = 'Cleaned up|refuse to die' },
                # The service's processes could not be listed: none of them
                # is stopped, that is said, and neither "All clear" nor
                # "Cleaned up" is. Until 6.3.1 that list read as empty.
                @{ Case = 'svcexefails'; Quiet = $false; Exit = 0
                   Has = @('LOG: \[!\] cowork-svc''s processes could not be listed \(Access is denied\), so none were stopped', 'STEP 4 DONE')
                   Not = 'ASKED|KILLED|All clear' },
                @{ Case = 'desktop-svcexefails'; Quiet = $false; Exit = 0
                   Has = @('LOG: \[!\] cowork-svc''s processes could not be listed \(Access is denied\), so none were stopped', 'KILLED 101',
                           'LOG: \[!\] Whether the orphans stopped is not known: cowork-svc''s processes could not be listed \(Access is denied\)', 'STEP 4 DONE')
                   Not = 'KILLED 900|All clear|Cleaned up' },
                # Whether the service runs could not be read: it is not
                # started again, and that is said. Until 6.3.1 the run ended
                # here without a word.
                @{ Case = 'desktop-session-svcdenied'; Quiet = $true; Exit = 4
                   Has = @('STOPPED PART-WAY: 1 Claude Code session',
                           'LOG: \[!\] Whether CoworkVMService is running could not be read \(Access is denied\), so it was not started again; run Fix when Cowork is needed')
                   Not = 'ASKED|KILLED|STEP 4 DONE|STARTED CoworkVMService|Could not start' }
            )) {
                $sArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $s4Script, '-Case', $sc.Case)
                if ($sc.Quiet) { $sArgs += '-Quiet' }
                $sOut = (& powershell.exe @sArgs 2>&1 | Out-String)
                $sExit = $LASTEXITCODE
                $missing = @($sc.Has | Where-Object { $sOut -notmatch $_ })
                $label = "Step 4, $($sc.Case)$(if ($sc.Quiet) { ', -Quiet' } else { ', manual' })"
                Assert-Case "$label exits $($sc.Exit)" ($sExit -eq $sc.Exit -and $missing.Count -eq 0 -and $sOut -notmatch $sc.Not) `
                    "exit $sExit; missing [$($missing -join '; ')]; output: $($sOut.Trim())"
            }
        } finally {
            # Clean-up of this section's own temp script: its cases are
            # counted, and a file left in %TEMP% changes nothing.
            Remove-Item -LiteralPath $s4Script -Force -ErrorAction SilentlyContinue
        }
    }
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Find-ClaudeExe never hands back the CLI"

# A cache written before 6.2.0 can hold the terminal CLI's path: the old Step 1
# capture took the first process named claude, and the cache check only knew
# the claude-code folder. Every relaunch would then start the CLI. The App
# Paths and Start Menu tiers get the same check. Test-Path, Get-ItemProperty,
# Get-ChildItem and New-Object are faked inside the scope that runs it, so
# neither the registry nor the real Start Menu is read; the cache is a real
# file in a sandbox, because the function deletes a cache it rejects.
$findDef = Get-FixFunction 'Find-ClaudeExe'
Assert-Case "Find-ClaudeExe was found in Fix" ($null -ne $findDef)
if ($findDef) {
    $sandbox = Join-Path ([System.IO.Path]::GetTempPath()) ('cf-findexe-' + [guid]::NewGuid().ToString('N'))
    $realTermCli = Join-Path $sandbox '.local\bin\claude.exe'
    $realAppCli  = Join-Path $sandbox 'Claude-Data\claude_x\claude-code\2.1.327\claude.exe'
    $realApp     = Join-Path $sandbox 'Claude\Claude.exe'
    foreach ($f in $realTermCli, $realAppCli, $realApp) {
        $null = New-Item -ItemType Directory -Path (Split-Path $f -Parent) -Force
        [System.IO.File]::WriteAllText($f, '')
    }
    $script:FakeProducts[$realTermCli] = 'Claude Code'
    $script:FakeProducts[$realAppCli]  = 'Claude Code'
    $script:FakeProducts[$realApp]     = 'Claude'
    $script:ClaudeEnv         = [pscustomobject]@{ ClaudeExe = $null; ClaudeExeSource = $null }
    $script:CapturedClaudeExe = $null
    $presentDef  = Get-FixFunction 'Test-FixPathPresent'
    $firstErrDef = Get-FixFunction 'Get-FixFirstError'
    try {
        $found = & {
            $ExePathCache   = Join-Path $sandbox '.claude-exe-path'
            $appPathsTarget = $null
            $lnkTarget      = $null
            # What could not be read: 'cached' (the path the cache names),
            # 'reg' (the App Paths key), 'menu' (the Start Menu folders), or
            # 'partial' (a folder under the Start Menu, so the search writes
            # an error and finds nothing). Test-Path answers $false for each,
            # as the real one does for a path it cannot read; the reads
            # since 6.3.1 get the error. Access denied is the identity used.
            $readFails      = ''
            $findLog        = New-Object System.Collections.ArrayList
            function Log { param([string]$Message, [string]$Colour, [switch]$Indent) $null = $Colour, $Indent; [void]$findLog.Add($Message) }
            # The errors Get-Item really gives (measured 2026-11-23): for a
            # registry key it may not open, SecurityException, "Requested
            # registry access is not allowed."; for a file or folder,
            # ItemExistsUnauthorizedAccessError, "Access is denied", which
            # does not name the path. Get-ChildItem's error for a folder in
            # its walk, DirUnauthorizedAccessError, does name it.
            function New-DeniedRecord {
                param($Target)
                # Module-qualified: New-Object is faked in this scope.
                if ("$Target" -like 'HK*:*') {
                    return (Microsoft.PowerShell.Utility\New-Object System.Management.Automation.ErrorRecord(
                        (Microsoft.PowerShell.Utility\New-Object System.Security.SecurityException 'Requested registry access is not allowed.'),
                        'System.Security.SecurityException', 'PermissionDenied', $Target))
                }
                return (Microsoft.PowerShell.Utility\New-Object System.Management.Automation.ErrorRecord(
                    (Microsoft.PowerShell.Utility\New-Object System.UnauthorizedAccessException 'Access is denied'),
                    'ItemExistsUnauthorizedAccessError', 'PermissionDenied', $Target))
            }
            function New-WalkDeniedRecord {
                param($Target)
                return (Microsoft.PowerShell.Utility\New-Object System.Management.Automation.ErrorRecord(
                    (Microsoft.PowerShell.Utility\New-Object System.UnauthorizedAccessException "Access to the path '$Target' is denied."),
                    'DirUnauthorizedAccessError', 'PermissionDenied', $Target))
            }
            function Test-Path {
                param($Path)
                if ("$Path" -like 'HKLM:*')      { return ($readFails -ne 'reg' -and [bool]$appPathsTarget) }
                if ("$Path" -like '*Start Menu') { return ($readFails -ne 'menu' -and ([bool]$lnkTarget -or $readFails -eq 'partial')) }
                if ($readFails -eq 'cached' -and "$Path" -eq $realApp) { return $false }
                return (Microsoft.PowerShell.Management\Test-Path -LiteralPath $Path)
            }
            function Get-ClaudeItemOrNull {
                param([string]$LiteralPath)
                if ($LiteralPath -like 'HKLM:*') {
                    if ($readFails -eq 'reg') { throw (New-DeniedRecord $LiteralPath) }
                    if ($appPathsTarget) { return [pscustomobject]@{ PSPath = $LiteralPath } }
                    return $null
                }
                if ($LiteralPath -like '*Start Menu') {
                    if ($readFails -eq 'menu') { throw (New-DeniedRecord $LiteralPath) }
                    if ($lnkTarget -or $readFails -eq 'partial') { return [pscustomobject]@{ PSPath = $LiteralPath } }
                    return $null
                }
                if ($readFails -eq 'cached' -and $LiteralPath -eq $realApp) { throw (New-DeniedRecord $LiteralPath) }
                # A sandbox file: the script's Get-Item fake gives its product
                # name, which Test-ClaudeCodeCliPath reads.
                if (Microsoft.PowerShell.Management\Test-Path -LiteralPath $LiteralPath) { return (Get-Item -LiteralPath $LiteralPath -ErrorAction Stop) }
                return $null
            }
            function Get-ItemProperty {
                param($Path, $LiteralPath, $ErrorAction)
                $null = $Path, $LiteralPath, $ErrorAction
                return [pscustomobject]@{ '(default)' = $appPathsTarget }
            }
            function Get-ChildItem {
                [CmdletBinding()]
                param([Parameter(Position = 0)]$Path, $LiteralPath, [switch]$Recurse, $Filter)
                $null = $Path, $LiteralPath, $Recurse, $Filter
                if ($readFails -eq 'partial') {
                    $PSCmdlet.WriteError((New-WalkDeniedRecord 'Start Menu\Programs\Locked'))
                    return
                }
                if ($lnkTarget) { return [pscustomobject]@{ FullName = 'Claude.lnk' } }
            }
            function New-Object {
                param($ComObject)
                $null = $ComObject
                $shell = [pscustomobject]@{ Target = $lnkTarget }
                $shell | Add-Member -MemberType ScriptMethod -Name CreateShortcut -Value {
                    param($LinkPath)
                    $null = $LinkPath
                    [pscustomobject]@{ TargetPath = $this.Target }
                }
                return $shell
            }
            if ($presentDef)  { . ([scriptblock]::Create($presentDef.Extent.Text)) }
            if ($firstErrDef) { . ([scriptblock]::Create($firstErrDef.Extent.Text)) }
            . ([scriptblock]::Create($findDef.Extent.Text))

            $r = [ordered]@{}
            foreach ($c in @(
                @{ Key = 'cacheTerm'; Cache = $realTermCli }
                @{ Key = 'cacheCli';  Cache = $realAppCli }
                @{ Key = 'cacheApp';  Cache = $realApp }
                @{ Key = 'regTerm';   Reg = $realTermCli }
                @{ Key = 'regApp';    Reg = $realApp }
                @{ Key = 'lnkTerm';   Lnk = $realTermCli }
                @{ Key = 'lnkApp';    Lnk = $realApp }
                @{ Key = 'cacheUnread'; Cache = $realApp; Fails = 'cached' }
                @{ Key = 'regUnread';   Reg = $realApp; Fails = 'reg' }
                @{ Key = 'regTargetUnread'; Reg = $realApp; Fails = 'cached' }
                @{ Key = 'menuUnread';  Lnk = $realApp; Fails = 'menu' }
                @{ Key = 'menuPartial'; Fails = 'partial' }
            )) {
                # The cache the case before wrote. Until 6.3.1 one that could
                # not be deleted was passed over, and this case then read the
                # last one's. File.Delete is no error for a file that is not
                # there, and throws for one it cannot delete, which stops the
                # gate.
                [System.IO.File]::Delete($ExePathCache)
                if ($c['Cache']) { [System.IO.File]::WriteAllText($ExePathCache, $c['Cache']) }
                $appPathsTarget = $c['Reg']
                $lnkTarget      = $c['Lnk']
                $readFails      = "$($c['Fails'])"
                $findLog.Clear()
                $got = Find-ClaudeExe
                # GetAttributes tells a cache that is not there
                # (FileNotFoundException) from one that could not be read,
                # which throws and fails the case. File.Exists answered $false
                # for both, which would pass "the cache holding it is removed".
                $kept = $(try { $null = [System.IO.File]::GetAttributes($ExePathCache); $true }
                          catch [System.IO.FileNotFoundException] { $false })
                $r[$c.Key] = [pscustomobject]@{
                    Got       = "$got"
                    CacheKept = $kept
                    Logged    = ($findLog -join ' | ')
                }
            }
            $r
        }
    } finally {
        # Clean-up of this section's own sandbox: its results are in hand,
        # and a folder left in %TEMP% changes nothing.
        Remove-Item -LiteralPath $sandbox -Recurse -Force -ErrorAction SilentlyContinue
        $script:ClaudeEnv = $null
    }
    Assert-Case "a cached terminal CLI path is refused"      ($found['cacheTerm'].Got -eq '') "got $($found['cacheTerm'].Got)"
    Assert-Case "and the cache holding it is removed"        (-not $found['cacheTerm'].CacheKept)
    Assert-Case "a cached app-started CLI path is refused"   ($found['cacheCli'].Got -eq '')  "got $($found['cacheCli'].Got)"
    Assert-Case "and the cache holding it is removed"        (-not $found['cacheCli'].CacheKept)
    Assert-Case "a cached app path is used"                  ($found['cacheApp'].Got -eq $realApp) "got $($found['cacheApp'].Got); logged: $($found['cacheApp'].Logged)"
    Assert-Case "and its cache is kept"                      ($found['cacheApp'].CacheKept)
    Assert-Case "App Paths naming the terminal CLI is refused" ($found['regTerm'].Got -eq '') "got $($found['regTerm'].Got)"
    Assert-Case "App Paths naming the app is used"           ($found['regApp'].Got -eq $realApp) "got $($found['regApp'].Got); logged: $($found['regApp'].Logged)"
    Assert-Case "a Start Menu shortcut to the terminal CLI is refused" ($found['lnkTerm'].Got -eq '') "got $($found['lnkTerm'].Got)"
    Assert-Case "a Start Menu shortcut to the app is used"   ($found['lnkApp'].Got -eq $realApp) "got $($found['lnkApp'].Got); logged: $($found['lnkApp'].Logged)"
    # What could not be read is said, and a cache whose path could not be
    # read is kept. Until 6.3.1 each read as nothing there, without a word,
    # and the cache was deleted on it.
    Assert-Case "a cached path that could not be read: the cache is kept" `
        ($found['cacheUnread'].Got -eq '' -and $found['cacheUnread'].CacheKept) "got '$($found['cacheUnread'].Got)', kept $($found['cacheUnread'].CacheKept)"
    Assert-Case "and it says so" `
        ($found['cacheUnread'].Logged -match '\[!\] The cached Claude\.exe could not be read \(Access is denied\), so it was passed over' -and
         $found['cacheUnread'].Logged -notmatch 'Cached path invalid') "logged: $($found['cacheUnread'].Logged)"
    Assert-Case "an App Paths key that could not be read is said" `
        ($found['regUnread'].Got -eq '' -and $found['regUnread'].Logged -match '\[!\] The App Paths key for claude\.exe could not be read \(Requested registry access is not allowed\.\)') `
        "got '$($found['regUnread'].Got)'; logged: $($found['regUnread'].Logged)"
    Assert-Case "an App Paths target that could not be read is said" `
        ($found['regTargetUnread'].Got -eq '' -and $found['regTargetUnread'].Logged -match '\[!\] The Claude\.exe App Paths names could not be read \(Access is denied\)') `
        "got '$($found['regTargetUnread'].Got)'; logged: $($found['regTargetUnread'].Logged)"
    Assert-Case "a Start Menu that could not be searched is said" `
        ($found['menuUnread'].Got -eq '' -and $found['menuUnread'].Logged -match '\[!\] The Start Menu at .*Start Menu could not be searched for a Claude shortcut \(Access is denied\)') `
        "got '$($found['menuUnread'].Got)'; logged: $($found['menuUnread'].Logged)"
    Assert-Case "part of the Start Menu that could not be searched is said" `
        ($found['menuPartial'].Got -eq '' -and $found['menuPartial'].Logged -match '\[!\] Part of .*Start Menu could not be searched for a Claude shortcut \(Access to the path ''Start Menu\\Programs\\Locked'' is denied\.\)') `
        "got '$($found['menuPartial'].Got)'; logged: $($found['menuPartial'].Logged)"
}

# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- no script kills every process named claude"

# The scan reads the parse tree, not lines, so a pipe split over lines with a
# backtick, a here-string whose variables are escaped, or a lookup assigned on
# one line and stopped thirty lines later cannot hide a kill from it. It flags:
#   - processes chosen by the name claude (Get-Process claude, Get-Process
#     -Name $ProcessName, Get-Process or Get-CimInstance piped into a
#     Where-Object on the name, Get-CimInstance filtered on claude.exe) that
#     reach Stop-Process, .Kill() or a Terminate in the same pipeline;
#   - the same lookup held in a variable, when the variable's latest
#     assignment above the kill is that lookup (or a pipeline from a variable
#     that was), and the variable is piped to a kill, passed to Stop-Process,
#     has .Kill() called on it, or is walked by a foreach that kills;
#   - Stop-Process -Name claude, taskkill /IM claude (direct or through
#     Start-Process), and wmic ending claude.exe;
#   - all of that inside a script the toolkit generates from a string, which
#     is parsed as code in its turn.
# A lookup stops being broad once a later stage of its pipeline filters on
# IsDesktop, the claude-code folder, the product name or
# Test-ClaudeCodeCliPath. What it does not follow: a lookup returned from a
# function, a variable across function scopes, or a variable reassigned
# further down inside a loop. Assignments are read in source order.

$script:AstT = @{
    Cmd     = ([System.Management.Automation.Language.CommandAst])
    CmdExpr = ([System.Management.Automation.Language.CommandExpressionAst])
    Param   = ([System.Management.Automation.Language.CommandParameterAst])
    Pipe    = ([System.Management.Automation.Language.PipelineAst])
    Var     = ([System.Management.Automation.Language.VariableExpressionAst])
    ArrExpr = ([System.Management.Automation.Language.ArrayExpressionAst])
    SubExpr = ([System.Management.Automation.Language.SubExpressionAst])
    Paren   = ([System.Management.Automation.Language.ParenExpressionAst])
    Binary  = ([System.Management.Automation.Language.BinaryExpressionAst])
    ArrLit  = ([System.Management.Automation.Language.ArrayLiteralAst])
    Convert = ([System.Management.Automation.Language.ConvertExpressionAst])
    ForEach = ([System.Management.Automation.Language.ForEachStatementAst])
    Invoke  = ([System.Management.Automation.Language.InvokeMemberExpressionAst])
    Member  = ([System.Management.Automation.Language.MemberExpressionAst])
    Str     = ([System.Management.Automation.Language.StringConstantExpressionAst])
    XStr    = ([System.Management.Automation.Language.ExpandableStringExpressionAst])
}
$script:ClaudeNameRx = '(?i)^["'']?(claude(\.exe)?\*?|\$(script:)?ProcessName)["'']?$'
$script:NameFilterRx = '(?i)(\$(script:)?ProcessName\b|["'']claude(\.exe)?["'']|-(eq|like|match)\s+claude\b)'
$script:NarrowRx     = '(?i)IsDesktop|claude-code|ProductName|Test-ClaudeCodeCliPath'
$script:KillRx       = '(?i)\bStop-Process\b|\.Kill\(|\bTerminate\b'
$script:ScanAssign   = @{}

function Get-ScanVarKey {
    param($VarAst)
    return ($VarAst.VariablePath.UserPath -replace '^(script|global|local|private):', '').ToLowerInvariant()
}

function Get-ScanCmdName {
    param($Cmd)
    $n = $Cmd.GetCommandName()
    if ($n) { return [string]$n }
    return ''
}

function Test-ScanNamesClaude {
    param($Node)
    if ($null -eq $Node) { return $false }
    if ($Node -is $script:AstT.ArrLit) {
        foreach ($e in $Node.Elements) { if (Test-ScanNamesClaude $e) { return $true } }
        return $false
    }
    return ($Node.Extent.Text -match $script:ClaudeNameRx)
}

function Test-ScanNameLookup {
    # Get-Process given the name claude, by -Name or by position, or the
    # region's Get-ClaudeProcessByName (6.3.1), which is Get-Process -Name.
    param($Cmd)
    if ((Get-ScanCmdName $Cmd) -notmatch '(?i)^(Get-Process|gps|ps|Get-ClaudeProcessByName)$') { return $false }
    $els = $Cmd.CommandElements
    for ($i = 1; $i -lt $els.Count; $i++) {
        $e = $els[$i]
        if ($e -is $script:AstT.Param) {
            if ($e.Argument -and (Test-ScanNamesClaude $e.Argument)) { return $true }
        } elseif (Test-ScanNamesClaude $e) {
            return $true
        }
    }
    return $false
}

function Test-ScanBroadPipeline {
    # True when the pipeline starts from processes chosen by the name claude
    # and no later stage narrows them to the app.
    param($Pipe)
    $els = $Pipe.PipelineElements
    for ($i = 1; $i -lt $els.Count; $i++) {
        if ($els[$i].Extent.Text -match $script:NarrowRx) { return $false }
    }
    $head = $els[0]
    if ($head -is $script:AstT.CmdExpr) { return (Test-ScanBroadExpr $head.Expression) }
    if ($head -isnot $script:AstT.Cmd) { return $false }
    if (Test-ScanNameLookup $head) { return $true }
    $name = Get-ScanCmdName $head
    if ($name -match '(?i)^(Get-CimInstance|gcim|Get-WmiObject|gwmi)$' -and $head.Extent.Text -match '(?i)claude\.exe') {
        return $true
    }
    if ($name -match '(?i)^(Get-Process|gps|ps|Get-CimInstance|gcim|Get-WmiObject|gwmi)$') {
        for ($i = 1; $i -lt $els.Count; $i++) {
            $e = $els[$i]
            if ($e -is $script:AstT.Cmd -and (Get-ScanCmdName $e) -match '^(Where-Object|where|\?)$' -and
                $e.Extent.Text -match '(?i)\bName\b' -and $e.Extent.Text -match $script:NameFilterRx) {
                return $true
            }
        }
    }
    return $false
}

function Test-ScanBroadStatement {
    param($Statement)
    if ($Statement -is $script:AstT.Pipe)    { return (Test-ScanBroadPipeline $Statement) }
    if ($Statement -is $script:AstT.CmdExpr) { return (Test-ScanBroadExpr $Statement.Expression) }
    return $false
}

function Test-ScanBroadExpr {
    param($Expr)
    if ($null -eq $Expr) { return $false }
    if ($Expr -is $script:AstT.Var) {
        return (Test-ScanTainted (Get-ScanVarKey $Expr) $Expr.Extent.StartOffset)
    }
    if ($Expr -is $script:AstT.Invoke)  { return $false }
    if ($Expr -is $script:AstT.Member)  { return (Test-ScanBroadExpr $Expr.Expression) }
    if ($Expr -is $script:AstT.Convert) { return (Test-ScanBroadExpr $Expr.Child) }
    if ($Expr -is $script:AstT.Paren)   { return (Test-ScanBroadStatement $Expr.Pipeline) }
    if ($Expr -is $script:AstT.ArrExpr -or $Expr -is $script:AstT.SubExpr) {
        foreach ($s in $Expr.SubExpression.Statements) { if (Test-ScanBroadStatement $s) { return $true } }
        return $false
    }
    if ($Expr -is $script:AstT.Binary) {
        return ((Test-ScanBroadExpr $Expr.Left) -or (Test-ScanBroadExpr $Expr.Right))
    }
    if ($Expr -is $script:AstT.ArrLit) {
        foreach ($e in $Expr.Elements) { if (Test-ScanBroadExpr $e) { return $true } }
    }
    return $false
}

function Test-ScanTainted {
    # True when the variable's latest assignment above this point is a broad
    # lookup. A += keeps looking further up; a plain = ends the search.
    param([string]$Key, [int]$Offset)
    $list = $script:ScanAssign[$Key]
    if (-not $list) { return $false }
    for ($i = $list.Count - 1; $i -ge 0; $i--) {
        $a = $list[$i]
        if ($a.Extent.EndOffset -gt $Offset) { continue }
        if (Test-ScanBroadStatement $a.Right) { return $true }
        if ($a.Operator -eq [System.Management.Automation.Language.TokenKind]::Equals) { return $false }
    }
    return $false
}

function Get-ScanLabel {
    param([string]$Where, $Node)
    return "$Where line $($Node.Extent.StartLineNumber): " + (($Node.Extent.Text -split "`r?`n")[0].Trim())
}

function Find-BroadKill {
    # Parses $Code and returns the broad kills in it, how many kill sites it
    # examined, and how many parse errors it met. $Where labels each hit.
    param([string]$Code, [string]$Where, [int]$Depth = 0)
    $errs = $null
    $ast  = [System.Management.Automation.Language.Parser]::ParseInput($Code, [ref]$null, [ref]$errs)
    $saved = $script:ScanAssign
    $assign = @{}
    $assignNodes = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true) |
        Sort-Object { $_.Extent.StartOffset })
    foreach ($a in $assignNodes) {
        $left = $a.Left
        if ($left -is $script:AstT.Convert) { $left = $left.Child }
        if ($left -isnot $script:AstT.Var) { continue }
        $k = Get-ScanVarKey $left
        if (-not $assign.ContainsKey($k)) { $assign[$k] = New-Object System.Collections.ArrayList }
        [void]$assign[$k].Add($a)
    }
    $script:ScanAssign = $assign
    $hits  = New-Object System.Collections.ArrayList
    $kills = 0
    try {
        foreach ($n in $ast.FindAll({ param($x) $null = $x; $true }, $true)) {
            if ($n -is $script:AstT.Pipe) {
                $els = $n.PipelineElements
                $killing = $false
                for ($i = 1; $i -lt $els.Count; $i++) { if ($els[$i].Extent.Text -match $script:KillRx) { $killing = $true } }
                if ($killing -and (Test-ScanBroadPipeline $n)) { [void]$hits.Add((Get-ScanLabel $Where $n)) }
            } elseif ($n -is $script:AstT.Cmd) {
                $name = Get-ScanCmdName $n
                $text = $n.Extent.Text
                if ($name -match '(?i)^(Stop-Process|spps|kill)$') {
                    $kills++
                    $els = $n.CommandElements
                    for ($i = 1; $i -lt $els.Count; $i++) {
                        $e = $els[$i]
                        if ($e -is $script:AstT.Param) {
                            $pn = $e.ParameterName.ToLowerInvariant()
                            if ($pn -and ('name'.StartsWith($pn) -or 'processname'.StartsWith($pn))) {
                                $val = $e.Argument
                                if (-not $val -and $i + 1 -lt $els.Count) { $val = $els[$i + 1] }
                                if (Test-ScanNamesClaude $val) { [void]$hits.Add((Get-ScanLabel $Where $n)) }
                            }
                            continue
                        }
                        foreach ($v in $e.FindAll({ param($x) $x -is [System.Management.Automation.Language.VariableExpressionAst] }, $true)) {
                            if (Test-ScanTainted (Get-ScanVarKey $v) $v.Extent.StartOffset) { [void]$hits.Add((Get-ScanLabel $Where $n)) }
                        }
                    }
                } elseif ($name -match '(?i)^taskkill(\.exe)?$' -or
                          ($name -match '(?i)^(Start-Process|saps|start)$' -and $text -match '(?i)taskkill')) {
                    $kills++
                    if ($text -match '(?i)/IM\s+["'']?(claude|\$(script:)?ProcessName)\b') { [void]$hits.Add((Get-ScanLabel $Where $n)) }
                } elseif ($name -match '(?i)^wmic(\.exe)?$') {
                    $kills++
                    if ($text -match '(?i)claude\.exe' -and $text -match '(?i)\b(delete|terminate)\b') { [void]$hits.Add((Get-ScanLabel $Where $n)) }
                }
            } elseif ($n -is $script:AstT.ForEach) {
                if ($n.Body.Extent.Text -match $script:KillRx -and (Test-ScanBroadStatement $n.Condition)) {
                    [void]$hits.Add((Get-ScanLabel $Where $n))
                }
            } elseif ($n -is $script:AstT.Invoke) {
                if ($n.Member.Extent.Text -eq 'Kill') {
                    $kills++
                    if (Test-ScanBroadExpr $n.Expression) { [void]$hits.Add((Get-ScanLabel $Where $n)) }
                }
            } elseif ($Depth -lt 2 -and (($n -is $script:AstT.XStr) -or
                      ($n -is $script:AstT.Str -and "$($n.StringConstantType)" -ne 'BareWord'))) {
                $v = [string]$n.Value
                if ($v -match '(?i)Stop-Process|taskkill|wmic|\.Kill\(') {
                    $inner = Find-BroadKill -Code $v -Where "$Where line $($n.Extent.StartLineNumber), generated script," -Depth ($Depth + 1)
                    foreach ($h in $inner.Hits) { [void]$hits.Add($h) }
                    $kills += $inner.Kills
                }
            }
        }
    } finally {
        $script:ScanAssign = $saved
    }
    return [pscustomobject]@{
        Hits   = @($hits | Select-Object -Unique)
        Kills  = $kills
        Errors = @($errs).Count
    }
}

# The scan has to flag every broad kill this toolkit shipped until 6.2.0, or a
# clean result below means nothing. The first five are copied from 0aa8d07,
# the last commit before the fix, with only enough around them to parse; the
# rest are shapes a regression could take.
$mustFlag = [ordered]@{
    'the region helper that lists by name, into a kill' = 'Get-ClaudeProcessByName -Name $ProcessName | Stop-Process -Force'
    'the region helper that lists by name, held and killed' = @'
$all = @(Get-ClaudeProcessByName -Name $ProcessName)
$all | Stop-Process -Force
'@
    'Fix -Close (0aa8d07)' = @'
$procs = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
if ($procs.Count -gt 0) {
    $procs | Stop-Process -Force -ErrorAction SilentlyContinue
}
'@
    'Fix Step 1 (0aa8d07)' = @'
$claudeProcs = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
if ($claudeProcs.Count -gt 0) {
    if ($PSCmdlet.ShouldProcess("$($claudeProcs.Count) Claude process(es)", "Stop")) {
        $claudeProcs | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}
'@
    'Fix Step 4 orphans (0aa8d07)' = @'
$remaining = @(Get-Process -ErrorAction SilentlyContinue | Where-Object {
    ($_.Name -eq $ProcessName) -or ($_.Name -eq $ServiceExe)
})
if ($remaining.Count -gt 0) {
    foreach ($proc in $remaining) {
        try { $proc | Stop-Process -Force -ErrorAction Stop } catch { $null = $_ }
    }
}
'@
    'Fix escalations (0aa8d07)' = @'
Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Stop-Process -Force
'@
    'Prevent NoGPU launcher (0aa8d07)' = @'
$ps1Content = @"
`$ErrorActionPreference = 'Stop'
`$running = @(Get-Process claude -ErrorAction SilentlyContinue)
if (`$running.Count -gt 0) {
    `$running | Stop-Process -Force -ErrorAction SilentlyContinue
}
"@
'@
    'Stop-Process -Name claude'          = 'Stop-Process -Name claude -Force'
    'Stop-Process -Name, quoted variable' = 'Stop-Process -Name "$ProcessName" -Force'
    'taskkill'                           = 'taskkill /F /IM claude.exe'
    'taskkill through Start-Process'     = 'Start-Process "taskkill" -ArgumentList "/F /IM claude.exe" -NoNewWindow -Wait'
    'a double-quoted name'               = 'Get-Process "claude" | Stop-Process -Force'
    'a pipe split by a backtick' = @'
Get-Process -Name $ProcessName -ErrorAction SilentlyContinue `
    | Stop-Process -Force
'@
    'a reassignment, then .Kill() in a foreach' = @'
$all = @(Get-Process -Name claude)
$all = @($all | Sort-Object Id)
foreach ($p in $all) { $p.Kill() }
'@
    'a variable passed to Stop-Process' = @'
$c = Get-Process claude
Stop-Process -InputObject $c -Force
'@
    'a CIM lookup by name' = @'
Get-CimInstance Win32_Process -Filter "Name = 'claude.exe'" | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
'@
}
foreach ($k in $mustFlag.Keys) {
    $r = Find-BroadKill -Code $mustFlag[$k] -Where 'control'
    Assert-Case "control flagged: $k" ($r.Hits.Count -gt 0 -and $r.Errors -eq 0) "hits $($r.Hits.Count), parse errors $($r.Errors)"
}

# And it has to pass the kills that replaced them, and the lookups that never
# kill, or it would be red for good and get switched off.
$mustPass = [ordered]@{
    'CPU sampling across every claude (Fix until 6.3.1)' = @'
$claudeCheck = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)
foreach ($cp in $claudeCheck) { $cpu1 = $cp.TotalProcessorTime.TotalMilliseconds }
'@
    'CPU sampling across every claude (Fix today)' = @'
$cpuProcs = @(Get-ClaudeProcessByName -Name $ProcessName)
foreach ($cp in $cpuProcs) { $cpu1 = $cp.TotalProcessorTime.TotalMilliseconds }
'@
    'the service, by name' = 'Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue | Stop-Process -Force'
    'taskkill on the service' = 'Start-Process "taskkill" -ArgumentList "/F /IM $ServiceExe.exe" -NoNewWindow -Wait'
    'Step 4 orphans (Fix today)' = @'
$orphanUnread    = $null
$orphanInventory = @()
try {
    $orphanInventory = @(Get-ClaudeProcessInventory -Environment $script:ClaudeEnv)
} catch {
    $orphanUnread = Get-ClaudeErrorText $_
}
$remaining = @(@($orphanInventory | Where-Object { $_.IsDesktop } | ForEach-Object { $_.Process }) +
               @(Get-Process -Name $ServiceExe -ErrorAction SilentlyContinue))
foreach ($proc in $remaining) {
    try { $proc | Stop-Process -Force -ErrorAction Stop } catch { $null = $_ }
}
'@
    'the NoGPU launcher filter (Prevent today)' = @'
$ps1Content = @"
function Get-ClaudeAppProcess {
    @(Get-Process claude -ErrorAction SilentlyContinue | Where-Object {
        `$exe = `$null
        try { `$exe = `$_.MainModule.FileName } catch { `$null = `$_ }
        if (-not `$exe -or `$exe -match '[\\/]claude-code[\\/]') { return `$false }
        `$product = ''
        try { `$product = [string](Get-Item -LiteralPath `$exe).VersionInfo.ProductName } catch { `$null = `$_ }
        `$product -ne 'Claude Code'
    })
}
`$running = @(Get-ClaudeAppProcess)
`$running | Stop-Process -Force -ErrorAction SilentlyContinue
"@
'@
    'a broad lookup narrowed before the kill' = @'
$p = @(Get-Process -Name claude)
$p = @($p | Where-Object { -not (Test-ClaudeCodeCliPath $_.Path) })
$p | Stop-Process -Force
'@
    'a comment' = '# Get-Process claude | Stop-Process'
}
foreach ($k in $mustPass.Keys) {
    $r = Find-BroadKill -Code $mustPass[$k] -Where 'control'
    Assert-Case "control passed: $k" ($r.Hits.Count -eq 0 -and $r.Errors -eq 0) "hits: $($r.Hits -join ' | ')"
}

foreach ($n in 'Fix-ClaudeDesktop.ps1', 'Watch-ClaudeHealth.ps1', 'Prevent-ClaudeIssues.ps1') {
    $r = Find-BroadKill -Code ([System.IO.File]::ReadAllText((Join-Path $Root $n))) -Where $n
    Write-Host ("  ....  {0}: {1} kill site(s) examined" -f $n, $r.Kills)
    Assert-Case "$n parses"                                      ($r.Errors -eq 0) "$($r.Errors) parse error(s)"
    Assert-Case "$n kills no process chosen by the name claude"  ($r.Hits.Count -eq 0) ($r.Hits -join ' | ')
}

Write-Host ""
Write-Host ("Process inventory: {0} passed, {1} failed" -f $pass, $fail)
if ($fail -gt 0) {
    Write-Host "Process inventory gate FAILED" -ForegroundColor Red
    exit 1
}
Write-Host "Process inventory gate PASSED" -ForegroundColor Green
exit 0
