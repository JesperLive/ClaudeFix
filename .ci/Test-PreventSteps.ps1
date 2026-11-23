<#
.SYNOPSIS
    Drives Prevent-ClaudeIssues.ps1's own functions and steps against fakes.
.DESCRIPTION
    The functions are lifted out of Prevent with the language parser, and
    the steps that are not functions (1, 8 to 15, 17 and 19 of setup, the
    parts of 15 and 17 that remove an older version's files, and 1 to 5 of
    -Undo) are cut out of the source between their Step calls and
    wrapped in a function, so what runs here is the shipped text. So are
    the two scripts setup writes for later, the Claude (Admin) task's
    payload and the GPU-free launcher, which run here in this process. The
    monitor's clock functions are lifted from Watch-ClaudeHealth.ps1 too.

    Nothing on this machine is changed. Files are read only in the folder
    this run makes under %TEMP%: Test-Path, Get-Item and Get-ChildItem
    pass a path that is not a registry path to the real cmdlet, and
    Get-Content has no fake, but none is given a path outside that folder
    (measured 2026-11-22 with a copy of this file that logged each one:
    1833 calls, none outside it). Beyond that folder, what is read for real
    is the name of this computer ([Environment]::MachineName), which the
    undo record carries; the name of the account the run is under
    ([Security.Principal.WindowsIdentity]::GetCurrent().Name), which step
    15 gives the monitor's task; this process's environment variables,
    some of which cases set, for this process only; the clock; and the
    list of commands, to check that each name below is its fake. Every other
    command those steps reach that reads or changes the machine is a fake
    defined below: powercfg, sc.exe,
    w32tm and hcsdiag, the
    registry (an in-memory copy), the network adapter, storage and CIM
    cmdlets, the service, process, Defender, scheduled task and Hyper-V
    cmdlets, and the Task Scheduler's COM service. The static .NET calls
    that cannot be shadowed (the clock step's DNS lookup and Stopwatch, the
    Desktop and %APPDATA% folders), and the folders step 15 looks in for the
    monitor's script, are swapped in the cut-out text for fakes. Anything
    else that could change the machine is a tripwire that throws and is
    listed at the end, and the run refuses to start unless every one of
    those names resolves to its fake. The monitor runs w32tm.exe by its full
    path in a job: the lifted copy is pointed at the fake instead, and the
    fake Start-Job refuses any path. Files go to a folder under %TEMP% that
    is removed afterwards.

    Exits 0 when every case behaves, 1 otherwise.
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
Set-StrictMode -Version Latest

$script:pass = 0
$script:fail = 0
function Assert-Case {
    param([string]$Name, $Ok, [string]$Detail)
    if ($Ok -is [bool] -and $Ok) {
        $script:pass++
        Write-Host "  PASS  $Name" -ForegroundColor Green
    } else {
        $script:fail++
        $d = ''
        if ($Detail) { $d = " ($Detail)" }
        Write-Host "  FAIL  $Name$d" -ForegroundColor Red
    }
}
function Invoke-Section {
    param([string]$Title, [scriptblock]$Body)
    Write-Host ""
    Write-Host "--- $Title"
    try {
        & $Body
    } catch {
        Assert-Case "$Title ran to the end" $false "$($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)"
    }
}

# ---------------------------------------------------------------------
# Parse the two scripts
# ---------------------------------------------------------------------
$preventPath = Join-Path $Root 'Prevent-ClaudeIssues.ps1'
$watchPath   = Join-Path $Root 'Watch-ClaudeHealth.ps1'
function Read-ScriptAst {
    param([string]$Path)
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    if (@($errors).Count -gt 0) { throw "parse errors in $Path" }
    return $ast
}
$pAst = Read-ScriptAst $preventPath
$wAst = Read-ScriptAst $watchPath
$pSrc = [System.IO.File]::ReadAllText($preventPath)

function Find-FunctionDef {
    param($Ast, [string]$Name)
    return $Ast.Find({
        param($n)
        $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name
    }, $true)
}
function Find-TopAssignment {
    param($Ast, [string]$Left)
    foreach ($s in $Ast.EndBlock.Statements) {
        if ($s -is [System.Management.Automation.Language.AssignmentStatementAst] -and $s.Left.Extent.Text -eq $Left) { return $s }
    }
    return $null
}

# ---------------------------------------------------------------------
# Sandbox and the script-level values the lifted code reads
# ---------------------------------------------------------------------
$sandbox = Join-Path ([System.IO.Path]::GetTempPath()) ('cf-prevent-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $sandbox -Force
$ClaudeAppData = Join-Path $sandbox 'Claude'
# The undo record's folder stands in for %LOCALAPPDATA%\ClaudeFix.
$UndoDir       = Join-Path $sandbox 'ClaudeFix'
$UndoFile      = Join-Path $UndoDir 'prevent-undo.json'
$BackupFile    = Join-Path $ClaudeAppData 'power-plan-backup.txt'
$CsBackupFile  = Join-Path $ClaudeAppData 'connected-standby-backup.txt'
$ServiceName   = 'CoworkVMService'
$script:ClaudeEnv = [pscustomobject]@{
    HcsDiagPath = 'Invoke-FakeHcs'
    ClaudeExe   = $null
    AppDataDir  = $ClaudeAppData
    ProcessName = 'claude'
    VmCachePath = $null
}
$ServiceExe     = 'cowork-svc'
$ClaudeDataRoot = Join-Path $sandbox 'Claude-Data'
$TaskName       = 'ClaudeCoworkWatchdog'
$BootTaskName   = 'ClaudeCoworkBootFix'
$TaskPath       = '\Claude\'
$ClaudeIconFile = Join-Path $ClaudeAppData 'Claude.ico'
# The Desktop and %APPDATA% folders -Undo deletes shortcuts from stand in
# the sandbox: the cut-out steps ask Get-FakeFolder for them.
$fakeDesktop    = Join-Path $sandbox 'Desktop'
$fakeAppData    = Join-Path $sandbox 'AppData'
# Step 15 looks for the monitor's script beside Prevent, then in
# C:\ClaudeFix and in the user's Desktop and Documents folders: each is a
# folder in the sandbox, and Prevent's own path is $script:FakeCommandPath,
# which Reset-State clears.
$fakeProfile    = Join-Path $sandbox 'Profile'
$fakeKitRoot    = Join-Path $sandbox 'ClaudeFixRoot'
function Get-FakeFolder {
    param([string]$Which)
    if ($Which -eq 'Desktop') { return $fakeDesktop }
    if ($Which -eq 'Profile') { return $fakeProfile }
    if ($Which -eq 'ClaudeFix') { return $fakeKitRoot }
    return $fakeAppData
}
function Get-FakeCommandPath { return $script:FakeCommandPath }

# A fake reports what a cmdlet reports: the same error id, category and
# exception type, which is what Prevent tells "not there" from "could not
# be read" by. The exception types of the service and process cmdlets live
# in this assembly.
Add-Type -AssemblyName 'Microsoft.PowerShell.Commands.Management'
function New-FakeErrorRecord {
    param([string]$TypeName, [string]$Message, [string]$Id, [string]$Category, $Target)
    $ex = New-Object $TypeName $Message
    return (New-Object System.Management.Automation.ErrorRecord($ex, $Id,
        [System.Management.Automation.ErrorCategory]$Category, $Target))
}
function Test-InSandbox {
    param([string]$Path)
    return ([string]$Path).StartsWith($sandbox, [StringComparison]::OrdinalIgnoreCase)
}
function Select-FakePath {
    param($Path, $LiteralPath)
    foreach ($x in @($LiteralPath) + @($Path)) { if ($x) { return [string]$x } }
    return ''
}

# ---------------------------------------------------------------------
# Tripwires: anything that could change the machine and that no case
# here expects to be called.
# ---------------------------------------------------------------------
$script:Tripped = New-Object System.Collections.ArrayList
$tripwires = @(
    'Restart-NetAdapter', 'Disable-NetAdapter', 'Enable-NetAdapter', 'Rename-NetAdapter',
    'Set-MpPreference', 'Stop-Service', 'Restart-Service', 'Set-Service', 'New-Service',
    'Set-ScheduledTask', 'Disable-ScheduledTask', 'Enable-ScheduledTask', 'Stop-ScheduledTask', 'New-ScheduledTask',
    'New-ItemProperty', 'Clear-ItemProperty', 'Rename-ItemProperty', 'Copy-Item', 'Move-Item',
    'Rename-Item', 'Add-Content', 'Restart-Computer', 'Stop-Computer',
    'New-NetNat', 'New-NetFirewallRule', 'Remove-NetFirewallRule',
    'Set-VM', 'Start-VM', 'Stop-VM', 'Start-Process', 'bcdedit.exe', 'reg.exe', 'netsh.exe'
)
foreach ($tw in $tripwires) {
    . ([scriptblock]::Create("function $tw { [void]`$script:Tripped.Add('$tw ' + (@(`$args) -join ' ')); throw 'TRIPWIRE: $tw' }"))
}

# ---------------------------------------------------------------------
# Fake registry: HKLM:\ and HKCU:\ paths live in $script:Reg, a table of
# key path to an ordered table of value name to value. Every other path is
# passed to the real cmdlet when it is in the sandbox, and trips otherwise.
# ---------------------------------------------------------------------
function Test-RegPath { param([string]$Path) return ([string]$Path) -match '^(HKLM|HKCU):\\' }
function Get-RegNorm  { param([string]$Path) return $Path.TrimEnd('\').ToLowerInvariant() }
function Add-FakeRegKey {
    param([string]$Path)
    $n = Get-RegNorm $Path
    if (-not $script:Reg.ContainsKey($n)) {
        $script:Reg[$n] = [pscustomobject]@{ Path = $Path.TrimEnd('\'); Vals = [ordered]@{} }
    }
    return $script:Reg[$n]
}
function Set-FakeRegValue {
    param([string]$Path, [string]$Name, $Value)
    $k = Add-FakeRegKey $Path
    $k.Vals[$Name] = $Value
}
function Get-FakeRegValue {
    param([string]$Path, [string]$Name)
    $n = Get-RegNorm $Path
    if (-not $script:Reg.ContainsKey($n)) { return $null }
    $v = $script:Reg[$n].Vals
    if (-not $v.Contains($Name)) { return $null }
    $x = $v[$Name]
    if ($x -is [array]) { return ,$x }
    return $x
}
function Test-FakeRegValue {
    param([string]$Path, [string]$Name)
    $n = Get-RegNorm $Path
    return ($script:Reg.ContainsKey($n) -and $script:Reg[$n].Vals.Contains($Name))
}
function New-FakeRegKeyObject {
    param($Entry)
    $leaf = $Entry.Path.Substring($Entry.Path.LastIndexOf('\') + 1)
    $o = [pscustomobject]@{ PSChildName = $leaf; Name = $Entry.Path; FakeEntry = $Entry }
    $o | Add-Member -MemberType ScriptMethod -Name GetValueNames -Value { return @($this.FakeEntry.Vals.Keys) }
    $o | Add-Member -MemberType ScriptMethod -Name GetValue -Value {
        param($n)
        $v = $this.FakeEntry.Vals
        if (-not $v.Contains($n)) { return $null }
        $x = $v[$n]
        if ($x -is [array]) { return ,$x }
        return $x
    }
    $o | Add-Member -MemberType ScriptProperty -Name ValueCount -Value { return $this.FakeEntry.Vals.Count }
    $o | Add-Member -MemberType ScriptProperty -Name SubKeyCount -Value {
        $pre = (Get-RegNorm $this.FakeEntry.Path) + '\'
        return @($script:Reg.Keys | Where-Object { $_.StartsWith($pre) -and $_.Substring($pre.Length) -notmatch '\\' }).Count
    }
    return $o
}

function Test-Path {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string[]]$Path, [string[]]$LiteralPath, [string]$PathType, [switch]$IsValid)
    $p = Select-FakePath $Path $LiteralPath
    # A file or folder $script:ItemReadFault matches cannot be read: Test-Path
    # answers $false for it, as for one that is not there, and writes this
    # error (measured 2026-11-23, an explicit Deny on the folder that holds
    # it). A registry key is not faulted here: Test-Path answers $true for a
    # key that cannot be read (measured the same day).
    if ($script:ItemReadFault -and -not (Test-RegPath $p) -and $p -like $script:ItemReadFault) {
        $PSCmdlet.WriteError((New-FakeErrorRecord 'System.UnauthorizedAccessException' 'Access is denied' `
            'ItemExistsUnauthorizedAccessError' 'PermissionDenied' $p))
        return $false
    }
    if (Test-RegPath $p) { return $script:Reg.ContainsKey((Get-RegNorm $p)) }
    return Microsoft.PowerShell.Management\Test-Path @PSBoundParameters
}
function Get-Item {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string[]]$Path, [string[]]$LiteralPath, [switch]$Force)
    $p = Select-FakePath $Path $LiteralPath
    # $script:ItemReadFault is a -like pattern: a path it matches cannot be
    # read. The error is the one Get-Item -Force gives for a file under an
    # explicit Deny (measured 2026-11-23): its message does not name the
    # path. Until 6.3.1 this fake's message did, as no real one measured
    # does, so a test could not see a line that failed to name it.
    if ($script:ItemReadFault -and $p -like $script:ItemReadFault) {
        $PSCmdlet.WriteError((New-FakeErrorRecord 'System.UnauthorizedAccessException' 'Access is denied' `
            'ItemExistsUnauthorizedAccessError' 'PermissionDenied' $p))
        return
    }
    if (Test-RegPath $p) {
        $n = Get-RegNorm $p
        # $script:RegItemReadFault is a -like pattern: a key it matches cannot
        # be opened, with the error Get-Item gives for a key Windows will not
        # open (measured 2026-11-23, HKLM:\SAM\SAM). The fake Test-Path still
        # answers whether the key is there; the real one answers $true.
        if ($script:RegItemReadFault -and $p -like $script:RegItemReadFault) {
            $PSCmdlet.WriteError((New-FakeErrorRecord 'System.Security.SecurityException' 'Requested registry access is not allowed.' `
                'System.Security.SecurityException' 'PermissionDenied' $p))
            return
        }
        # A key that is not there gives the error Get-Item gives (measured
        # 2026-11-23); until 6.3.1 this fake threw a bare message instead.
        if (-not $script:Reg.ContainsKey($n)) {
            $PSCmdlet.WriteError((New-FakeErrorRecord 'System.Management.Automation.ItemNotFoundException' `
                "Cannot find path '$p' because it does not exist." 'PathNotFound' 'ObjectNotFound' $p))
            return
        }
        return (New-FakeRegKeyObject $script:Reg[$n])
    }
    return Microsoft.PowerShell.Management\Get-Item @PSBoundParameters
}
function Get-ChildItem {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string[]]$Path, [string[]]$LiteralPath, [switch]$Force, [switch]$Recurse, [string]$Filter)
    $p = Select-FakePath $Path $LiteralPath
    if (Test-RegPath $p) {
        $pre = (Get-RegNorm $p) + '\'
        # $script:RegReadFault is a -like pattern: a key it matches cannot be
        # read. A key that is not there gives the error Get-ChildItem gives.
        if ($script:RegReadFault -and $p -like $script:RegReadFault) {
            $PSCmdlet.WriteError((New-FakeErrorRecord 'System.Security.SecurityException' 'Requested registry access is not allowed.' `
                'System.Security.SecurityException' 'PermissionDenied' $p))
            return
        }
        if (-not $script:Reg.ContainsKey((Get-RegNorm $p)) -and @($script:Reg.Keys | Where-Object { $_.StartsWith($pre) }).Count -eq 0) {
            $PSCmdlet.WriteError((New-FakeErrorRecord 'System.Management.Automation.ItemNotFoundException' `
                "Cannot find path '$p' because it does not exist." 'PathNotFound' 'ObjectNotFound' $p))
            return
        }
        $out = @()
        foreach ($k in @($script:Reg.Keys | Sort-Object)) {
            if ($k.StartsWith($pre) -and $k.Substring($pre.Length) -notmatch '\\') { $out += (New-FakeRegKeyObject $script:Reg[$k]) }
        }
        return $out
    }
    return Microsoft.PowerShell.Management\Get-ChildItem @PSBoundParameters
}
function New-Item {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string[]]$Path, [string]$ItemType, [switch]$Force, $Value, [string]$Name)
    $p = Select-FakePath $Path $null
    if (Test-RegPath $p) {
        [void]$script:Calls.Add("New-Item $p")
        if ($script:RegNewKeyFault) { throw "Requested registry access is not allowed." }
        $null = Add-FakeRegKey $p
        return (New-FakeRegKeyObject $script:Reg[(Get-RegNorm $p)])
    }
    if (Test-InSandbox $p) { return Microsoft.PowerShell.Management\New-Item @PSBoundParameters }
    [void]$script:Tripped.Add("New-Item $p")
    throw "TRIPWIRE: New-Item $p"
}
# Set-Content writes only in the sandbox, where step 17 writes its
# launchers; anywhere else it is a tripwire. $script:SetContentFault is a
# -like pattern: a file it matches is refused, as a denied one is.
function Set-Content {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string[]]$Path, [string[]]$LiteralPath, [Parameter(Position = 1)]$Value,
          [string]$Encoding, [switch]$Force, [switch]$NoNewline)
    $p = Select-FakePath $Path $LiteralPath
    if (-not (Test-InSandbox $p)) { [void]$script:Tripped.Add("Set-Content $p"); throw "TRIPWIRE: Set-Content $p"; return }
    # The error is the one Set-Content gives for a file the user may not
    # write (measured 2026-11-23); until 6.3.1 its id and category were not.
    if ($script:SetContentFault -and $p -like $script:SetContentFault) {
        $PSCmdlet.WriteError((New-FakeErrorRecord 'System.UnauthorizedAccessException' "Access to the path '$p' is denied." `
            'System.UnauthorizedAccessException' 'NotSpecified' $p))
        return
    }
    Microsoft.PowerShell.Management\Set-Content @PSBoundParameters
}
function Set-ItemProperty {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string[]]$Path, [string[]]$LiteralPath, [Parameter(Position = 1)][string]$Name,
          [Parameter(Position = 2)]$Value, [string]$Type, [switch]$Force)
    $p = Select-FakePath $Path $LiteralPath
    if (-not (Test-RegPath $p)) {
        [void]$script:Tripped.Add("Set-ItemProperty $p")
        throw "TRIPWIRE: Set-ItemProperty $p"
    }
    $id = (Get-RegNorm $p) + '|' + $Name.ToLowerInvariant()
    [void]$script:Calls.Add("Set-ItemProperty $p|$Name ($Type)")
    if ($script:RegFault.ContainsKey($id)) {
        if ($script:RegFault[$id] -eq 'throw') { throw "Requested registry access is not allowed." }
        return
    }
    if (-not $script:Reg.ContainsKey((Get-RegNorm $p))) { throw "Cannot find path '$p' because it does not exist." }
    $v = $Value
    if ($Type -eq 'DWord') { $v = [int]$Value }
    elseif ($Type -eq 'Binary') { $v = [byte[]]$Value }
    elseif ($Type -eq 'String') { $v = [string]$Value }
    Set-FakeRegValue $p $Name $v
}
function Remove-ItemProperty {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string[]]$Path, [string[]]$LiteralPath, [Parameter(Position = 1)][string[]]$Name, [switch]$Force)
    $p = Select-FakePath $Path $LiteralPath
    if (-not (Test-RegPath $p)) {
        [void]$script:Tripped.Add("Remove-ItemProperty $p")
        throw "TRIPWIRE: Remove-ItemProperty $p"
    }
    [void]$script:Calls.Add("Remove-ItemProperty $p|$($Name -join ',')")
    $n = Get-RegNorm $p
    foreach ($x in $Name) {
        if ($script:RegRemoveFault.ContainsKey($n + '|' + $x.ToLowerInvariant())) { throw "Requested registry access is not allowed." }
    }
    if (-not $script:Reg.ContainsKey($n)) { throw "Cannot find path '$p' because it does not exist." }
    foreach ($x in $Name) {
        if (-not $script:Reg[$n].Vals.Contains($x)) { throw "Property $x does not exist at path $p." }
        $script:Reg[$n].Vals.Remove($x)
    }
}
function Remove-Item {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string[]]$Path, [string[]]$LiteralPath, [switch]$Force, [switch]$Recurse)
    $p = Select-FakePath $Path $LiteralPath
    if (Test-RegPath $p) {
        [void]$script:Calls.Add("Remove-Item $p")
        $n = Get-RegNorm $p
        if (-not $script:Reg.ContainsKey($n)) { throw "Cannot find path '$p' because it does not exist." }
        foreach ($k in @($script:Reg.Keys)) {
            if ($k -eq $n -or $k.StartsWith("$n\")) { $script:Reg.Remove($k) }
        }
        return
    }
    # $script:FileRemoveFault and $script:FileRemoveNoop are -like patterns:
    # a file the first matches is refused, as a locked or denied one is, and
    # one the second matches stays without an error. The error is the one
    # Remove-Item gives for a file the user may not delete (measured
    # 2026-11-23): its message does not name the file. Until 6.3.1 this
    # fake's message did, so a test could not see a line that failed to.
    if ($script:FileRemoveFault -and $p -like $script:FileRemoveFault) {
        $PSCmdlet.WriteError((New-FakeErrorRecord 'System.ArgumentException' 'Access to the path is denied.' `
            'RemoveFileSystemItemArgumentError' 'InvalidArgument' $p))
        return
    }
    if ($script:FileRemoveNoop -and $p -like $script:FileRemoveNoop) { return }
    if (Test-InSandbox $p) {
        Microsoft.PowerShell.Management\Remove-Item @PSBoundParameters
        # $script:ItemReadFaultAfterRemove becomes $script:ItemReadFault once
        # a file is removed: whether it is gone then cannot be read.
        if ($script:ItemReadFaultAfterRemove) { $script:ItemReadFault = $script:ItemReadFaultAfterRemove }
        return
    }
    [void]$script:Tripped.Add("Remove-Item $p")
    throw "TRIPWIRE: Remove-Item $p"
}
# Get-ItemProperty reads the same fake registry. A key that is not there,
# or a -Name it does not hold, is an error, as the real cmdlet gives it.
function Get-ItemProperty {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string[]]$Path, [string[]]$LiteralPath, [Parameter(Position = 1)][string[]]$Name)
    $p = Select-FakePath $Path $LiteralPath
    if (Test-RegPath $p) {
        $n = Get-RegNorm $p
        if (-not $script:Reg.ContainsKey($n)) {
            Write-Error -Message "Cannot find path '$p' because it does not exist." -Category ObjectNotFound
            return
        }
        $entry = $script:Reg[$n]
        foreach ($x in @($Name | Where-Object { $_ })) {
            if (-not $entry.Vals.Contains($x)) {
                Write-Error -Message "Property $x does not exist at path $p." -Category InvalidArgument
                return
            }
        }
        $o = [ordered]@{}
        foreach ($k in @($entry.Vals.Keys)) {
            if ($Name -and $Name -notcontains $k) { continue }
            $o[$k] = $entry.Vals[$k]
        }
        $o['PSPath'] = $entry.Path
        $o['PSChildName'] = $entry.Path.Substring($entry.Path.LastIndexOf('\') + 1)
        return [pscustomobject]$o
    }
    if (Test-InSandbox $p) { return Microsoft.PowerShell.Management\Get-ItemProperty @PSBoundParameters }
    [void]$script:Tripped.Add("Get-ItemProperty $p")
    throw "TRIPWIRE: Get-ItemProperty $p"
}

# ---------------------------------------------------------------------
# Fake powercfg. Plans are $script:Plans (GUID to name), the active one is
# $script:ActivePlan, and every AC index is $script:PcIdx keyed
# "plan|sub|setting". Each plan also has a key under PowerSchemes in the fake
# registry, which is where Get-UltimatePlanGuids finds the copies of the
# Ultimate Performance template. $script:ListOk is how many more /list calls
# succeed before it fails (-1: it never fails), and $script:QhFails makes
# every /qh fail, as powercfg does when it cannot reach the power service.
# ---------------------------------------------------------------------
$plBal       = '381b4222-f694-41f0-9685-ff5bb260df2e'
$plHigh      = '8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c'
$plSaver     = 'a1841308-3541-4fab-bc81-f71556f20b4a'
$plTemplate  = 'e9a42b02-d5df-448d-aa00-03f14749eb61'
$schemesRoot = 'HKLM:\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes'
$ultFriendly = '@%SystemRoot%\system32\powrprof.dll,-19,Ultimate Performance'
$subSleep    = '238c9fa8-0aad-41ed-83f4-97be242c8f20'
$setSleep    = '29f6c1db-86da-48c5-9fdb-f2b67b1f44da'
$powerKey    = 'HKLM:\SYSTEM\CurrentControlSet\Control\Power'

function Add-FakePlan {
    param([string]$Guid, [string]$Name, [string]$Friendly)
    $g = $Guid.ToLowerInvariant()
    $script:Plans[$g] = $Name
    Set-FakeRegValue "$schemesRoot\$g" 'FriendlyName' $Friendly
}
function Set-PowerFixture {
    Add-FakePlan $plBal   'Balanced'         '@%SystemRoot%\system32\powrprof.dll,-2,Balanced'
    Add-FakePlan $plHigh  'High performance' '@%SystemRoot%\system32\powrprof.dll,-14,High performance'
    Add-FakePlan $plSaver 'Power saver'      '@%SystemRoot%\system32\powrprof.dll,-16,Power saver'
    # The hidden template has a key of its own beside the plans, with the
    # same resource name its copies carry. It is not one of this machine's
    # plans and must never be taken for one.
    Set-FakeRegValue "$schemesRoot\$plTemplate" 'FriendlyName' $ultFriendly
    $script:ActivePlan = $plBal
}
function Format-FakeQh {
    param([string]$Plan, [string]$Sub, [string]$Setting, [long]$Ac)
    $lines = @(
        "Power Scheme GUID: $Plan  ($($script:Plans[$Plan]))",
        "  Subgroup GUID: $Sub  (Sleep)",
        "    Power Setting GUID: $Setting  (Sleep after)",
        '      Minimum Possible Setting: 0x00000000',
        '      Maximum Possible Setting: 0xffffffff',
        '      Possible Settings increment: 0x00000001',
        '      Possible Settings units: Seconds'
    )
    $hex = '0x{0:x8}' -f $Ac
    # The DC value differs from every AC value used here, so a reader that
    # took the wrong one of the two lines is caught.
    $dcHex = '0x0000abcd'
    if ($script:PcLang -eq 'nl') {
        $lines += "    Index van huidige wisselstroominstelling: $hex"
        if (-not $script:PcTruncate) { $lines += "    Index van huidige gelijkstroominstelling: $dcHex" }
    } else {
        $lines += "    Current AC Power Setting Index: $hex"
        if (-not $script:PcTruncate) { $lines += "    Current DC Power Setting Index: $dcHex" }
    }
    $lines += ''
    return $lines
}
function powercfg.exe {
    $a = @($args | ForEach-Object { [string]$_ })
    [void]$script:PcCalls.Add($a -join ' ')
    $global:LASTEXITCODE = 0
    if ($a.Count -eq 0) { $global:LASTEXITCODE = 1; return }
    $verb = $a[0].ToLowerInvariant()
    if ($verb -eq '/getactivescheme') {
        if ($script:ActivePlan) { return "Power Scheme GUID: $($script:ActivePlan)  ($($script:Plans[$script:ActivePlan]))" }
        $global:LASTEXITCODE = 1
        return
    }
    if ($verb -eq '/list') {
        if ($script:ListOk -eq 0) {
            $global:LASTEXITCODE = 1
            return 'Unable to perform operation. An unexpected error (0x65b) has occurred.'
        }
        if ($script:ListOk -gt 0) { $script:ListOk-- }
        $out = @('', 'Existing Power Schemes (* Active)', '-----------------------------------')
        foreach ($g in @($script:Plans.Keys)) {
            $star = ''
            if ($g -eq $script:ActivePlan) { $star = ' *' }
            $out += "Power Scheme GUID: $g  ($($script:Plans[$g]))$star"
        }
        return $out
    }
    if ($verb -eq '/duplicatescheme') {
        if ($script:NoTemplate -or $a.Count -lt 3 -or $a[1].ToLowerInvariant() -ne $plTemplate) {
            $global:LASTEXITCODE = 1
            return 'Unable to create a new power scheme. The specified power scheme does not exist.'
        }
        Add-FakePlan $a[2] 'Ultimate Performance' $ultFriendly
        return "Power Scheme GUID: $($a[2].ToLowerInvariant())  (Ultimate Performance)"
    }
    if ($verb -eq '/setactive') {
        $g = $a[1].ToLowerInvariant()
        if ($script:SetActiveRefuse -or -not $script:Plans.Contains($g)) {
            $global:LASTEXITCODE = 1
            return 'Unable to perform operation. An unexpected error (0x65b) has occurred.'
        }
        $script:ActivePlan = $g
        return
    }
    if ($verb -eq '/delete') {
        $g = $a[1].ToLowerInvariant()
        if ($script:DeleteRefuse -or $g -eq $script:ActivePlan -or -not $script:Plans.Contains($g)) {
            $global:LASTEXITCODE = 1
            return 'Unable to perform operation. An unexpected error (0x65b) has occurred.'
        }
        $script:Plans.Remove($g)
        $n = Get-RegNorm "$schemesRoot\$g"
        if ($script:Reg.ContainsKey($n)) { $script:Reg.Remove($n) }
        return
    }
    if ($verb -eq '/qh') {
        $k = ("{0}|{1}|{2}" -f $a[1], $a[2], $a[3]).ToLowerInvariant()
        if ($script:QhFails -or -not $script:PcIdx.ContainsKey($k) -or -not $script:Plans.Contains($a[1].ToLowerInvariant())) {
            $global:LASTEXITCODE = 1
            return 'The power scheme, subgroup or setting specified does not exist.'
        }
        return (Format-FakeQh $a[1] $a[2] $a[3] $script:PcIdx[$k])
    }
    if ($verb -eq '/setacvalueindex') {
        $k = ("{0}|{1}|{2}" -f $a[1], $a[2], $a[3]).ToLowerInvariant()
        if ($script:PcRefuse.ContainsKey($k) -or -not $script:PcIdx.ContainsKey($k)) {
            $global:LASTEXITCODE = 1
            return 'Invalid Parameters -- try "/?" for help'
        }
        $script:PcIdx[$k] = [long]$a[4]
        return
    }
    if ($verb -eq '/hibernate') {
        if ($script:HibernateRefuse) {
            $global:LASTEXITCODE = 1
            return 'Hibernation failed with the following error: The request is not supported.'
        }
        $on = 0
        if ($a[1] -eq 'on') { $on = 1 }
        Set-FakeRegValue $powerKey 'HibernateEnabled' ([int]$on)
        return
    }
    if ($verb -eq '/a') { return @($script:PowerAText -split "`n") }
    $global:LASTEXITCODE = 1
    [void]$script:Tripped.Add("powercfg.exe $($a -join ' ')")
}

# ---------------------------------------------------------------------
# Fake sc.exe: "failure <svc> reset= N actions= a/d/a/d" writes the
# FailureActions value into the fake registry, as the service manager does.
# ---------------------------------------------------------------------
function New-FailureBlob {
    param([long]$Reset, [long[]]$Pairs = @(), [int]$CountOverride = -1)
    $n = [int]([math]::Floor(@($Pairs).Count / 2))
    $count = $n
    if ($CountOverride -ge 0) { $count = $CountOverride }
    $bytes = New-Object byte[] (20 + 8 * $n)
    [BitConverter]::GetBytes([uint32]$Reset).CopyTo($bytes, 0)
    [BitConverter]::GetBytes([uint32]$count).CopyTo($bytes, 12)
    if ($n -gt 0) { [BitConverter]::GetBytes([uint32]20).CopyTo($bytes, 16) }
    for ($i = 0; $i -lt $n; $i++) {
        [BitConverter]::GetBytes([uint32]$Pairs[2 * $i]).CopyTo($bytes, 20 + 8 * $i)
        [BitConverter]::GetBytes([uint32]$Pairs[2 * $i + 1]).CopyTo($bytes, 24 + 8 * $i)
    }
    return ,$bytes
}
function sc.exe {
    $a = @($args | ForEach-Object { [string]$_ })
    [void]$script:ScCalls.Add($a -join ' ')
    $global:LASTEXITCODE = 0
    if ($a.Count -lt 2 -or $a[0] -ne 'failure') {
        $global:LASTEXITCODE = 1
        [void]$script:Tripped.Add("sc.exe $($a -join ' ')")
        return '[SC] unexpected call'
    }
    if ($script:ScFails) {
        $global:LASTEXITCODE = 5
        return '[SC] ChangeServiceConfig2 FAILED 5:', 'Access is denied.'
    }
    $svc = $a[1]
    $reset = 0
    $acts = ''
    for ($i = 2; $i -lt $a.Count - 1; $i++) {
        if ($a[$i] -eq 'reset=') { $reset = [long]$a[$i + 1] }
        if ($a[$i] -eq 'actions=') { $acts = $a[$i + 1] }
    }
    $pairs = @()
    if ($acts -and $acts -ne '""') {
        $parts = $acts.Split('/')
        for ($j = 0; $j + 1 -lt $parts.Count; $j += 2) {
            $t = 0
            if ($parts[$j] -eq 'restart') { $t = 1 } elseif ($parts[$j] -eq 'reboot') { $t = 2 } elseif ($parts[$j] -eq 'run') { $t = 3 }
            $pairs += [long]$t
            $pairs += [long]$parts[$j + 1]
        }
    }
    Set-FakeRegValue "HKLM:\SYSTEM\CurrentControlSet\Services\$svc" 'FailureActions' (New-FailureBlob -Reset $reset -Pairs $pairs)
    return '[SC] ChangeServiceConfig2 SUCCESS'
}

# ---------------------------------------------------------------------
# Fake w32tm, Start-Job and friends (the monitor runs w32tm in a job), DNS
# for Prevent's clock step, Start-Sleep, services, Defender, and the NAT rule
# -Undo step 2 looks for (there is none).
# ---------------------------------------------------------------------
function w32tm {
    $a = @($args | ForEach-Object { [string]$_ })
    [void]$script:W32Calls.Add($a -join ' ')
    $global:LASTEXITCODE = 0
    if ($a.Count -gt 0 -and $a[0] -eq '/stripchart') {
        # Each measurement takes the next chart in $script:ChartSeq while
        # there is one, then $script:ChartLines.
        if ($script:ChartSeq.Count -gt 0) {
            $c = $script:ChartSeq[0]
            $script:ChartSeq.RemoveAt(0)
            return @($c)
        }
        return @($script:ChartLines)
    }
    if ($a.Count -gt 0 -and $a[0] -eq '/resync') {
        # $script:ResyncTakesMs: how long the resync takes, for real, as a
        # resync that sets the clock forward moves the time read after it.
        # Start-Sleep is faked here, so the thread sleeps.
        if ($script:ResyncTakesMs -gt 0) { [System.Threading.Thread]::Sleep($script:ResyncTakesMs) }
        $script:ResyncDoneAt = [datetime]::UtcNow
        $global:LASTEXITCODE = $script:ResyncCode
        if ($script:ResyncCode -ne 0) { return 'The computer did not resync because no time data was available.' }
        return 'The command completed successfully.'
    }
    if ($a.Count -gt 1 -and $a[0] -eq '/query' -and $a[1] -eq '/source') { return @($script:TimeSource) }
    $global:LASTEXITCODE = 1
    [void]$script:Tripped.Add("w32tm $($a -join ' ')")
}
# The monitor runs w32tm in a job. The lifted Invoke-W32tm is given the
# fake's name; a path in the arguments would be the real w32tm.exe, so it
# trips before anything runs. $script:JobHangs makes every job hang, and
# $script:JobHangMatch only the ones whose arguments match it.
# $script:JobThrows, when set, is thrown instead, as a job that cannot start
# throws: a message, or an exception.
#
# $script:JobFails is the message of a job that ends Failed, as a real one
# whose process cannot be created does. Measured on Windows PowerShell 5.1:
# Start-Job does not throw, Wait-Job hands the job back, State is Failed,
# and Receive-Job writes one error and no output, whose message has no
# "[localhost] " and whose ErrorDetails do. $script:JobFailMatch limits it
# to the jobs whose arguments match, after the first $script:JobFailSkip of
# those. $script:JobNoExit is the message of a job that completes with no
# exit code and that error, as one whose program is not there does.
function New-FakeJob {
    param([string]$State, [object[]]$Out = $null, [bool]$Hang = $false, [string[]]$Errors = @(), [string]$ErrorDetails = '')
    return [pscustomobject]@{ FakeJob = $true; State = $State; Out = $Out; Hang = $Hang; Errors = @($Errors); ErrorDetails = $ErrorDetails }
}
function Start-Job {
    [CmdletBinding()]
    param([Parameter(Position = 0)][scriptblock]$ScriptBlock, [object[]]$ArgumentList)
    $flat = @($ArgumentList | ForEach-Object { $_ } | ForEach-Object { [string]$_ })
    if (@($flat | Where-Object { $_ -match '\\|\.exe$' }).Count -gt 0) {
        [void]$script:Tripped.Add("Start-Job $($flat -join ' ')")
        throw "TRIPWIRE: Start-Job $($flat -join ' ')"
    }
    if ($null -ne $script:JobThrows) {
        [void]$script:Calls.Add('Start-Job failed')
        throw $script:JobThrows
        return
    }
    $hang = [bool]$script:JobHangs
    if ($script:JobHangMatch -and (($flat -join ' ') -match $script:JobHangMatch)) { $hang = $true }
    if ($null -ne $script:JobFails -and (-not $script:JobFailMatch -or (($flat -join ' ') -match $script:JobFailMatch))) {
        if ($script:JobFailSkip -gt 0) {
            $script:JobFailSkip--
        } else {
            [void]$script:Calls.Add('Start-Job Failed')
            return (New-FakeJob -State 'Failed' -Errors @($script:JobFails) -ErrorDetails "[localhost] $($script:JobFails)")
        }
    }
    if ($null -ne $script:JobNoExit) {
        return (New-FakeJob -State 'Completed' -Out @([pscustomobject]@{ Text = ''; ExitCode = $null }) -Errors @($script:JobNoExit))
    }
    $out = $null
    if (-not $hang) { $out = @(& $ScriptBlock @ArgumentList) }
    $state = 'Completed'
    if ($hang) { $state = 'Running' }
    return (New-FakeJob -State $state -Out $out -Hang $hang)
}
function Wait-Job {
    [CmdletBinding()]
    param([Parameter(Position = 0, ValueFromPipeline)]$Job, [int]$Timeout)
    process {
        if ($Job.Hang) { return }
        return $Job
    }
}
function Receive-Job {
    [CmdletBinding()]
    param([Parameter(Position = 0, ValueFromPipeline)]$Job)
    process {
        foreach ($m in @($Job.Errors)) {
            $rec = [System.Management.Automation.ErrorRecord]::new(
                [System.Management.Automation.Remoting.PSRemotingTransportException]::new($m), 'JobFailure', 'OpenError', $null)
            if ($Job.ErrorDetails) { $rec.ErrorDetails = [System.Management.Automation.ErrorDetails]::new($Job.ErrorDetails) }
            $PSCmdlet.WriteError($rec)
        }
        return $Job.Out
    }
}
function Stop-Job {
    [CmdletBinding()]
    param([Parameter(Position = 0, ValueFromPipeline)]$Job)
    process { [void]$script:Calls.Add('Stop-Job') }
}
function Remove-Job {
    [CmdletBinding()]
    param([Parameter(Position = 0, ValueFromPipeline)]$Job, [switch]$Force)
    process { $null = $Job }
}

# The clock step's wait for the network runs on $script:FakeClockMs, which
# its Stopwatch reads and Start-Sleep and the lookups move. A lookup takes
# $script:DnsLatencyMs, or never answers when $script:DnsHang is set, and
# answers from try $script:DnsOkAfter on (0: never). Resolve-FakeTimeHost
# stands in for the blocking GetHostAddresses, which older versions called;
# Resolve-FakeTimeHostAsync for GetHostAddressesAsync, whose Wait(ms) gives
# up at ms, and without one waits as long as the lookup takes.
function Resolve-FakeTimeHost {
    param([string]$HostName)
    $null = $HostName
    $script:DnsTries++
    if ($script:DnsHang) {
        $script:FakeClockMs += 120000
        throw 'No such host is known'
    }
    $script:FakeClockMs += $script:DnsLatencyMs
    if ($script:DnsOkAfter -gt 0 -and $script:DnsTries -ge $script:DnsOkAfter) { return @('20.101.57.9') }
    throw 'No such host is known'
}
function Resolve-FakeTimeHostAsync {
    param([string]$HostName)
    $null = $HostName
    $task = [pscustomobject]@{ Result = $null }
    $task | Add-Member -MemberType ScriptMethod -Name Wait -Value {
        $limit = [long]-1
        if ($args.Count -gt 0) { $limit = [long]$args[0] }
        $script:DnsTries++
        $took = [long]$script:DnsLatencyMs
        if ($script:DnsHang) { $took = [long]::MaxValue }
        if ($limit -ge 0 -and $took -gt $limit) {
            $script:FakeClockMs += $limit
            return $false
        }
        if ($took -eq [long]::MaxValue) {
            # A wait with no bound on a lookup that never answers does not
            # return: ten minutes stand for that.
            $script:FakeClockMs += 600000
            throw 'The lookup never answered, and the wait had no bound.'
        }
        $script:FakeClockMs += $took
        if ($script:DnsOkAfter -gt 0 -and $script:DnsTries -ge $script:DnsOkAfter) {
            $this.Result = @('20.101.57.9')
            return $true
        }
        throw 'One or more errors occurred. (No such host is known)'
    }
    return $task
}
function New-FakeStopwatch {
    $sw = [pscustomobject]@{ StartMs = [long]$script:FakeClockMs }
    $sw | Add-Member -MemberType ScriptProperty -Name ElapsedMilliseconds -Value { return [long]($script:FakeClockMs - $this.StartMs) }
    return $sw
}
function Start-Sleep {
    param([int]$Seconds, [int]$Milliseconds)
    $script:Slept += $Seconds + ($Milliseconds / 1000.0)
    $script:FakeClockMs += 1000 * [long]$Seconds + $Milliseconds
}
# A service that is not there gives the error Get-Service gives for one; a
# name in $script:ServiceReadFault cannot be read.
function Get-Service {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string[]]$Name)
    $out = @()
    foreach ($n in $Name) {
        if (@($script:ServiceReadFault) -contains $n) {
            $PSCmdlet.WriteError((New-FakeErrorRecord 'System.ComponentModel.Win32Exception' 'Access is denied' `
                'CouldNotGetServiceInfo' 'PermissionDenied' $n))
            continue
        }
        if ($script:Services.ContainsKey($n)) {
            $out += [pscustomobject]@{ Name = $n; Status = $script:Services[$n] }
        } else {
            $PSCmdlet.WriteError((New-FakeErrorRecord 'Microsoft.PowerShell.Commands.ServiceCommandException' `
                "Cannot find any service with service name '$n'." 'NoServiceFoundForGivenName' 'ObjectNotFound' $n))
        }
    }
    return $out
}
function Start-Service {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string[]]$Name)
    [void]$script:SvcCalls.Add("Start-Service $($Name -join ',')")
    if ($script:StartServiceThrows) { throw "Service 'Windows Time (W32Time)' cannot be started due to the following error: Access is denied." }
    foreach ($n in $Name) { $script:Services[$n] = 'Running' }
}
function Remove-MpPreference {
    [CmdletBinding()]
    param([string[]]$ExclusionPath, [string[]]$ExclusionProcess)
    if ($script:MpFails) { throw 'Operation failed with the following error: 0x800106ba'; return }
    if ($ExclusionPath) { [void]$script:MpCalls.Add("ExclusionPath $($ExclusionPath -join ',')") }
    if ($ExclusionProcess) { [void]$script:MpCalls.Add("ExclusionProcess $($ExclusionProcess -join ',')") }
}
# Defender's exclusion lists are $script:MpPref.Path and .Process. Where
# Defender hides them from administrators it answers with one "N/A" entry,
# which $script:MpHidePath and $script:MpHideProcess stand for, and an empty
# list is $null, as Get-MpPreference gives it.
function Get-MpPreference {
    [CmdletBinding()]
    param()
    if ($script:MpReadFails) { throw 'Operation failed with the following error: 0x800106ba'; return }
    $hidden = [string[]]@('N/A: Must be an administrator to view exclusions')
    $paths = $null
    $procs = $null
    if ($script:MpHidePath) { $paths = $hidden } elseif ($script:MpPref.Path.Count -gt 0) { $paths = [string[]]@($script:MpPref.Path) }
    if ($script:MpHideProcess) { $procs = $hidden } elseif ($script:MpPref.Process.Count -gt 0) { $procs = [string[]]@($script:MpPref.Process) }
    return [pscustomobject]@{ ExclusionPath = $paths; ExclusionProcess = $procs }
}
function Add-MpPreference {
    [CmdletBinding()]
    param([string[]]$ExclusionPath, [string[]]$ExclusionProcess)
    if ($script:MpFails) { throw 'Operation failed with the following error: 0x800106ba'; return }
    foreach ($x in @($ExclusionPath | Where-Object { $_ })) {
        [void]$script:MpAdds.Add("Path $x")
        [void]$script:MpPref.Path.Add($x)
    }
    foreach ($x in @($ExclusionProcess | Where-Object { $_ })) {
        [void]$script:MpAdds.Add("Process $x")
        [void]$script:MpPref.Process.Add($x)
    }
}
# Stop-Process ends a process in $script:Procs. $script:StopProcFault makes
# it refuse, and $script:StopProcLinger leaves the process running without
# an error. A process that is not there, one that ended by itself, is an
# ObjectNotFound error. Each error is one that is not terminating, as
# Windows' are.
function Stop-Process {
    [CmdletBinding()]
    param([int[]]$Id, [switch]$Force)
    $null = $Force
    [void]$script:ProcCalls.Add("Stop-Process $($Id -join ',')")
    if ($script:StopProcFault) {
        Write-Error -Message "Cannot stop process ""powershell ($($Id -join ','))"" because of the following error: Access is denied" -Category CloseError
        return
    }
    $live = @($script:Procs | ForEach-Object { $_.Id })
    $gone = @($Id | Where-Object { $live -notcontains $_ })
    if (-not $script:StopProcLinger) { $script:Procs = @($script:Procs | Where-Object { $Id -notcontains $_.Id }) }
    if ($gone.Count -gt 0) { Write-Error -Message "Cannot find a process with the process identifier $($gone -join ',')." -Category ObjectNotFound }
}
# Fake WinNAT: $script:Nats are the NAT rules Get-NetNat lists. A name that
# matches none is an ObjectNotFound error, as in Windows. $script:NatReadFault
# is a list that cannot be read at all, an error that is not terminating, as
# the CIM cmdlets give; $script:NatRemoveFault is a removal that is refused.
# Every call is in $script:NatCalls. A removal is also a tripwire, recorded
# in $script:Tripped, unless the case set $script:NatRemoveExpected, which
# Reset-State clears: only -Undo step 2 removes a rule, so a removal
# anywhere else, in setup or in a function it calls, must fail the suite.
function New-FakeNat {
    param([string]$Name)
    return [pscustomobject]@{ Name = $Name; InternalIPInterfaceAddressPrefix = '172.16.0.0/24' }
}
# A CimException as Get-NetNat writes one, built in memory, nothing asked of
# the machine. Measured on this PC with Get-CimInstance and with a CDXML
# cmdlet on a class and a namespace that are not there: the exception is
# Microsoft.Management.Infrastructure.CimException itself, StatusCode is
# the CIM status (5 Invalid class, 3 Invalid namespace), and its error data
# holds the WMI code as error_Code (2147749904 = 0x80041010), with the
# message ending in a space. Built this way NativeErrorCode reads Ok, not
# the status; Prevent does not read it. -Code -1 leaves the WMI code out.
# $script:NatReadCim is the error Get-NetNat then writes.
function New-FakeCimError {
    param([uint32]$Status, [long]$Code, [string]$Message, [string]$Category = 'MetadataError')
    $none = [Microsoft.Management.Infrastructure.CimFlags]::None
    $inst = [Microsoft.Management.Infrastructure.CimInstance]::new('MSFT_ExtendedStatus')
    $inst.CimInstanceProperties.Add([Microsoft.Management.Infrastructure.CimProperty]::Create('CIMStatusCode', $Status, [Microsoft.Management.Infrastructure.CimType]::UInt32, $none))
    $inst.CimInstanceProperties.Add([Microsoft.Management.Infrastructure.CimProperty]::Create('Message', $Message, [Microsoft.Management.Infrastructure.CimType]::String, $none))
    $id = 'Get-NetNat'
    if ($Code -ge 0) {
        $hex = 'HRESULT 0x{0:x8}' -f $Code
        $inst.CimInstanceProperties.Add([Microsoft.Management.Infrastructure.CimProperty]::Create('error_Code', [uint32]$Code, [Microsoft.Management.Infrastructure.CimType]::UInt32, $none))
        $inst.CimInstanceProperties.Add([Microsoft.Management.Infrastructure.CimProperty]::Create('MessageID', $hex, [Microsoft.Management.Infrastructure.CimType]::String, $none))
        $id = "$hex,Get-NetNat"
    }
    return [pscustomobject]@{
        Exception = [Microsoft.Management.Infrastructure.CimException]::new($inst)
        Id        = $id
        Category  = [System.Management.Automation.ErrorCategory]$Category
    }
}
function Get-NetNat {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string[]]$Name)
    [void]$script:NatCalls.Add("Get-NetNat $($Name -join ',')".TrimEnd())
    if ($script:NatReadFault) { Write-Error -Message 'Access denied' -Category PermissionDenied; return }
    if ($null -ne $script:NatReadCim) {
        $f = $script:NatReadCim
        $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new($f.Exception, $f.Id, $f.Category, $null))
        return
    }
    if (-not $Name) { return @($script:Nats) }
    $out = @()
    foreach ($n in $Name) {
        $hit = @($script:Nats | Where-Object { $_.Name -like $n })
        if ($hit.Count -eq 0) {
            Write-Error -Message "No MSFT_NetNat objects found with property 'Name' equal to '$n'.  Verify the value of the property and retry." -Category ObjectNotFound
            continue
        }
        $out += $hit
    }
    return $out
}
function Remove-NetNat {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Position = 0)][string[]]$Name)
    $confirm = 'unset'
    if ($PSBoundParameters.ContainsKey('Confirm')) { $confirm = [bool]$PSBoundParameters['Confirm'] }
    [void]$script:NatCalls.Add("Remove-NetNat $($Name -join ',') Confirm=$confirm")
    if (-not $script:NatRemoveExpected) {
        [void]$script:Tripped.Add("Remove-NetNat $($Name -join ',') Confirm=$confirm")
        throw 'TRIPWIRE: Remove-NetNat'
        return
    }
    if ($script:NatRemoveFault) { throw 'Access is denied.'; return }
    $script:Nats = @($script:Nats | Where-Object { $Name -notcontains $_.Name })
}

# ---------------------------------------------------------------------
# Fake processes, CIM and hcsdiag.
# ---------------------------------------------------------------------
function Get-Process {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string[]]$Name, [int[]]$Id)
    # $script:ProcReadFault holds names and ids that cannot be read, or '*'
    # for every one: an error that is not the one for "not there".
    $fault = @($script:ProcReadFault)
    if ($PSBoundParameters.ContainsKey('Id')) {
        if ($fault -contains '*' -or $fault -contains [string]$Id[0]) {
            $PSCmdlet.WriteError((New-FakeErrorRecord 'System.ComponentModel.Win32Exception' 'Access is denied' `
                'CouldNotEnumerateProcesses' 'NotSpecified' $Id))
            return
        }
        $found = @($script:Procs | Where-Object { $Id -contains $_.Id })
        if ($found.Count -eq 0) {
            # The error Get-Process writes, with its id, category and
            # exception type, which tell a process that exited from a read
            # that failed.
            $PSCmdlet.WriteError((New-FakeErrorRecord 'Microsoft.PowerShell.Commands.ProcessCommandException' `
                "Cannot find a process with the process identifier $($Id -join ',')." 'NoProcessFoundForGivenId' 'ObjectNotFound' $Id))
            return
        }
        return $found
    }
    if ($Name) {
        $out = @()
        foreach ($n in $Name) {
            if ($fault -contains '*' -or $fault -contains $n) {
                $PSCmdlet.WriteError((New-FakeErrorRecord 'System.ComponentModel.Win32Exception' 'Access is denied' `
                    'CouldNotEnumerateProcesses' 'NotSpecified' $n))
                continue
            }
            $m = @($script:Procs | Where-Object { $_.ProcessName -eq $n })
            if ($m.Count -eq 0) {
                $PSCmdlet.WriteError((New-FakeErrorRecord 'Microsoft.PowerShell.Commands.ProcessCommandException' `
                    "Cannot find a process with the name `"$n`". Verify the process name and call the cmdlet again." `
                    'NoProcessFoundForGivenName' 'ObjectNotFound' $n))
            } else {
                $out += $m
            }
        }
        return $out
    }
    return @($script:Procs)
}
function Get-CimInstance {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string]$ClassName, [string]$Namespace, [string]$Filter)
    [void]$script:CimCalls.Add($ClassName)
    if ($ClassName -eq 'MSPower_DeviceEnable') {
        if ($script:DevEnableThrows) { throw 'Invalid namespace'; return }
        return @($script:DevEnable)
    }
    # $script:AvReadCim is the CIM error the antivirus read writes, and
    # $script:AvReadFault one that is not a CIM error.
    if ($ClassName -eq 'AntiVirusProduct') {
        if ($null -ne $script:AvReadCim) {
            $f = $script:AvReadCim
            $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new($f.Exception, $f.Id, $f.Category, $null))
            return
        }
        if ($script:AvReadFault) { Write-Error -Message 'The RPC server is unavailable.' -Category ResourceUnavailable; return }
        return @($script:AvProducts)
    }
    # $script:LogicalDisks maps a drive, "C:", to its DriveType (4 is a
    # network drive); a drive not in it is not listed, with no error, as in
    # Windows. A drive in $script:LogicalDiskFault cannot be read.
    if ($ClassName -eq 'Win32_LogicalDisk') {
        $dev = ''
        if ($Filter -match "DeviceID='([^']*)'") { $dev = $Matches[1] }
        [void]$script:StorageCalls.Add("Win32_LogicalDisk $dev")
        if (@($script:LogicalDiskFault) -contains $dev) { Write-Error -Message 'The RPC server is unavailable.' -Category ResourceUnavailable; return }
        if ($script:LogicalDisks.ContainsKey($dev)) { return [pscustomobject]@{ DeviceID = $dev; DriveType = [uint32]$script:LogicalDisks[$dev] } }
        return
    }
    if ($ClassName -eq 'MSPower_DeviceWakeEnable') {
        if ($null -ne $script:WakeReadCim) {
            $f = $script:WakeReadCim
            $PSCmdlet.WriteError([System.Management.Automation.ErrorRecord]::new($f.Exception, $f.Id, $f.Category, $null))
            return
        }
        return @($script:WakeEnable)
    }
    if ($ClassName -eq 'Win32_Process' -and $Filter -match 'powershell\.exe') {
        if ($script:PsProcsFault) { Write-Error -Message 'The RPC server is unavailable.' -Category ResourceUnavailable; return }
        return @($script:PsProcs)
    }
    if ($ClassName -eq 'Win32_Process' -and $Filter -match 'vmwp\.exe') {
        if ($script:VmWorkersFault) { Write-Error -Message 'The RPC server is unavailable.' -Category ResourceUnavailable; return }
        return @($script:VmWorkers)
    }
    [void]$script:Tripped.Add("Get-CimInstance $ClassName $Filter")
    return @()
}
function Set-CimInstance {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline)]$InputObject, [hashtable]$Property)
    process {
        [void]$script:NicCalls.Add("Set-CimInstance $($InputObject.InstanceName) Enable=$($Property['Enable'])")
        if ($script:SetCimThrows) { throw 'Access denied'; return }
        if (-not $script:SetCimNoop) { $InputObject.Enable = [bool]$Property['Enable'] }
    }
}

# Fake storage, for step 11. $script:Drives maps a drive letter, "C", to
# the local volume on it, @{ Disk = <number>; Bus = <BusType> }, or to
# $null for a volume with no partition behind it. A letter not in it is
# not a local volume, and Get-Volume, Get-Partition and Get-Disk then write
# the ObjectNotFound error they give in Windows, in the words measured on
# this PC. $script:StorageFault names the one of the three that cannot be
# read at all, 'Volume', 'Partition' or 'Disk': an error that is not
# terminating, as theirs are. Every call is in $script:StorageCalls, with
# the Win32_LogicalDisk reads of the fake Get-CimInstance.
function Get-Volume {
    [CmdletBinding()]
    param([char[]]$DriveLetter)
    $l = [string]$DriveLetter[0]
    [void]$script:StorageCalls.Add("Get-Volume $l")
    if ($script:StorageFault -eq 'Volume') { Write-Error -Message 'The RPC server is unavailable.' -Category ResourceUnavailable; return }
    if (-not $script:Drives.ContainsKey($l)) {
        Write-Error -Message "No MSFT_Volume objects found with property 'DriveLetter' equal to '$l'.  Verify the value of the property and retry." -Category ObjectNotFound
        return
    }
    return [pscustomobject]@{ DriveLetter = $l }
}
function Get-Partition {
    [CmdletBinding()]
    param([char[]]$DriveLetter)
    $l = [string]$DriveLetter[0]
    [void]$script:StorageCalls.Add("Get-Partition $l")
    if ($script:StorageFault -eq 'Partition') { Write-Error -Message 'Access denied' -Category PermissionDenied; return }
    if (-not $script:Drives.ContainsKey($l) -or $null -eq $script:Drives[$l]) {
        Write-Error -Message "No MSFT_Partition objects found with property 'DriveLetter' equal to '$l'.  Verify the value of the property and retry." -Category ObjectNotFound
        return
    }
    return [pscustomobject]@{ DriveLetter = $l; DiskNumber = [uint32]$script:Drives[$l].Disk }
}
function Get-Disk {
    [CmdletBinding()]
    param([uint32[]]$Number)
    $n = $Number[0]
    [void]$script:StorageCalls.Add("Get-Disk $n")
    if ($script:StorageFault -eq 'Disk') { Write-Error -Message 'The device is not ready.' -Category DeviceError; return }
    foreach ($v in @($script:Drives.Values)) {
        if ($null -ne $v -and [uint32]$v.Disk -eq $n) { return [pscustomobject]@{ Number = $n; BusType = [string]$v.Bus } }
    }
    Write-Error -Message "No MSFT_Disk objects found with property 'Number' equal to '$n'.  Verify the value of the property and retry." -Category ObjectNotFound
}

$guidA   = 'DE1517EC-1234-4567-89AB-77A48CB1AD97'
$guidB   = 'AB1517EC-1234-4567-89AB-77A48CB1AD98'
$guidWsl = '11111111-2222-3333-4444-555555555555'
$guidVm  = 'CAFEF00D-0000-1111-2222-333344445555'
# $script:HcsOnlyCowork lists the Cowork VMs alone, so that once they are
# killed hcsdiag lists nothing, as it does on a machine with no other VM.
function Format-FakeHcsList {
    $out = @()
    if (-not $script:HcsOnlyCowork) {
        $out += "wsl-utility`n    VM,   Running, $guidWsl, wsl-utility"
        $out += "ubuntu-dev`n    VM,   Running, $guidVm, ubuntu-dev"
    }
    foreach ($g in $script:HcsLive) {
        $n = 'cowork-vm-' + $g.Substring(0, 8).ToLowerInvariant()
        $out += "$n`n    VM,                       `tRunning, $g, $n"
    }
    return ($out -join "`n")
}
function Invoke-FakeHcs {
    $a = @($args | ForEach-Object { [string]$_ })
    [void]$script:HcsCalls.Add($a -join ' ')
    $global:LASTEXITCODE = 0
    if ($a.Count -gt 0 -and $a[0] -eq 'list') {
        if ($script:HcsListCode -ne 0) {
            $global:LASTEXITCODE = $script:HcsListCode
            return 'The operation failed.'
        }
        if ($script:HcsListDownAfterKill -and $script:Killed.Count -gt 0) {
            $global:LASTEXITCODE = 1
            return 'The RPC server is unavailable.'
        }
        $text = Format-FakeHcsList
        if ($text) { return $text }
        return
    }
    if ($a.Count -gt 1 -and $a[0] -eq 'kill') {
        if ($script:HcsKillCode -ne 0) {
            $global:LASTEXITCODE = $script:HcsKillCode
            return 'Access is denied.'
        }
        [void]$script:Killed.Add($a[1])
        if (@($script:HcsStuck) -notcontains $a[1]) { $script:HcsLive.Remove($a[1]) }
        return
    }
    $global:LASTEXITCODE = 1
    [void]$script:Tripped.Add("hcsdiag $($a -join ' ')")
}
function New-FakeWorker {
    param([int]$ProcessId, [string]$Guid, $CreationDate = $null)
    return [pscustomobject]@{ ProcessId = $ProcessId; CommandLine = "`"C:\Windows\System32\vmwp.exe`" $Guid 0x330"; CreationDate = $CreationDate }
}
function New-FakeProcess {
    param([int]$Id, [string]$ProcessName, [string]$Priority = 'Normal')
    return [pscustomobject]@{ Id = $Id; ProcessName = $ProcessName; PriorityClass = $Priority }
}
function Set-VmFixture {
    [void]$script:HcsLive.Add($guidA)
    $script:VmWorkers = @(
        (New-FakeWorker 10 $guidA),
        (New-FakeWorker 11 $guidWsl),
        (New-FakeWorker 12 $guidVm),
        [pscustomobject]@{ ProcessId = 13; CommandLine = $null }
    )
    $script:Procs = @(
        (New-FakeProcess 10 'vmwp'),
        (New-FakeProcess 11 'vmwp'),
        (New-FakeProcess 12 'vmwp')
    )
}
function Get-Priority {
    param([int]$Id)
    foreach ($p in $script:Procs) { if ($p.Id -eq $Id) { return [string]$p.PriorityClass } }
    return '<gone>'
}

# ---------------------------------------------------------------------
# Fake Task Scheduler. $script:Tasks holds every task as "<path>|<name>",
# and $script:TaskFolders every folder below the root, as "\Claude". A task
# that is not there is an ObjectNotFound error, as Get-ScheduledTask gives
# it; $script:TaskQueryFault is a scheduler that cannot be reached. A key
# also in $script:HiddenTasks is a hidden task, which a folder's GetTasks
# lists only with flag 1 (TASK_ENUM_HIDDEN), as the real one does. The COM
# service Remove-PreventTaskFolder uses is faked by its function's name, so
# the real one is never connected, and the lift refuses to go on when that
# name holds Prevent's own.
# ---------------------------------------------------------------------
function Get-ScheduledTask {
    [CmdletBinding()]
    param([string]$TaskName, [string]$TaskPath)
    [void]$script:TaskCalls.Add("Get $TaskPath$TaskName")
    if ($script:TaskQueryFault) {
        Write-Error -Message 'The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)' -Category ConnectionError
        return
    }
    $key = "$TaskPath|$TaskName"
    if (-not $script:Tasks.ContainsKey($key)) {
        $PSCmdlet.WriteError((New-FakeTaskNotFound $TaskName))
        return
    }
    # $script:TaskStopReads: how many more reads a task that reads Running
    # goes on reading so, as an instance the scheduler is still winding
    # down does. 0 ends it at the next read, -1 never.
    if ($script:Tasks[$key] -eq 'Running' -and $script:TaskStopReads.ContainsKey($key)) {
        $left = [int]$script:TaskStopReads[$key]
        if ($left -eq 0) {
            $script:Tasks[$key] = 'Ready'
            $script:TaskStopReads.Remove($key)
        } elseif ($left -gt 0) {
            $script:TaskStopReads[$key] = $left - 1
        }
    }
    return [pscustomobject]@{ TaskName = $TaskName; TaskPath = $TaskPath; State = $script:Tasks[$key] }
}
# The error Get-ScheduledTask gives for a task that is not there.
function New-FakeTaskNotFound {
    param([string]$TaskName)
    return (New-FakeErrorRecord 'Microsoft.PowerShell.Cmdletization.Cim.CimJobException' `
        "No MSFT_ScheduledTask objects found with property 'TaskName' equal to '$TaskName'.  Verify the value of the property and retry." `
        'CmdletizationQuery_NotFound' 'ObjectNotFound' $TaskName)
}
function Unregister-ScheduledTask {
    [CmdletBinding(SupportsShouldProcess)]
    param([string]$TaskName, [string]$TaskPath)
    [void]$script:TaskCalls.Add("Unregister $TaskPath$TaskName")
    if ($script:UnregisterFault) { throw 'Access is denied.'; return }
    if (-not $script:Tasks.ContainsKey("$TaskPath|$TaskName")) {
        $PSCmdlet.WriteError((New-FakeTaskNotFound $TaskName))
        return
    }
    if ($script:UnregisterNoop) { return }
    $script:Tasks.Remove("$TaskPath|$TaskName")
}
# Steps 15 and 17 make a task: the monitor's, and Claude (Admin)'s.
# Register-ScheduledTask,
# Start-ScheduledTask and the New-ScheduledTask* cmdlets that build the
# task answer only when the case set $script:TaskInstallExpected, which
# Reset-State clears; anywhere else each is a tripwire, recorded in
# $script:Tripped, that throws, since only steps 15 and 17 make one. A
# registration is in $script:TaskCalls and $script:Tasks, and what it was
# given in $script:RegisteredTask. $script:RegisterFault and
# $script:StartTaskFault are a scheduler that refuses: an error that is
# not terminating, as the ScheduledTasks cmdlets give.
function Test-TaskInstallExpected {
    param([string]$Name, [string]$What)
    if ($script:TaskInstallExpected) { return $true }
    [void]$script:Tripped.Add("$Name $What".TrimEnd())
    return $false
}
function Register-ScheduledTask {
    [CmdletBinding()]
    param([string]$TaskName, [string]$TaskPath, $Action, $Trigger, $Settings, $Principal, [string]$Description, [switch]$Force)
    if (-not (Test-TaskInstallExpected 'Register-ScheduledTask' "$TaskPath$TaskName")) { throw 'TRIPWIRE: Register-ScheduledTask'; return }
    [void]$script:TaskCalls.Add("Register $TaskPath$TaskName")
    if ($script:RegisterFault) { Write-Error -Message 'Access is denied.' -Category PermissionDenied; return }
    # Registering does not end an instance that runs.
    if ($script:Tasks["$TaskPath|$TaskName"] -ne 'Running') { $script:Tasks["$TaskPath|$TaskName"] = 'Ready' }
    $script:RegisteredTask = [pscustomobject]@{
        Action = $Action; Trigger = $Trigger; Settings = $Settings; Principal = $Principal; Description = $Description; Force = [bool]$Force
    }
    return [pscustomobject]@{ TaskName = $TaskName; TaskPath = $TaskPath; State = 'Ready' }
}
function Start-ScheduledTask {
    [CmdletBinding()]
    param([string]$TaskName, [string]$TaskPath)
    if (-not (Test-TaskInstallExpected 'Start-ScheduledTask' "$TaskPath$TaskName")) { throw 'TRIPWIRE: Start-ScheduledTask'; return }
    [void]$script:TaskCalls.Add("Start $TaskPath$TaskName")
    if ($script:StartTaskFault) {
        Write-Error -Message 'The operator or administrator has refused the request. (Exception from HRESULT: 0x800710E0)' -Category InvalidOperation
        return
    }
    if (-not $script:Tasks.ContainsKey("$TaskPath|$TaskName")) {
        $PSCmdlet.WriteError((New-FakeTaskNotFound $TaskName))
        return
    }
    # A start the scheduler takes without an error and then does nothing
    # with ($script:StartTaskNoop), and the task's MultipleInstances
    # IgnoreNew: a start while an instance runs is ignored, with no error.
    if ($script:StartTaskNoop -or $script:Tasks["$TaskPath|$TaskName"] -eq 'Running') { return }
    $script:Tasks["$TaskPath|$TaskName"] = 'Running'
    # From here on the task cannot be read: a scheduler that stopped answering.
    if ($script:TaskQueryFaultAfterStart) { $script:TaskQueryFault = $true }
    # The task is deleted right after the start, by another program or a policy.
    if ($script:TaskGoneAfterStart) { $script:Tasks.Remove("$TaskPath|$TaskName") }
}
# Setup step 14 reads the WSL feature and runs wsl.exe -l -v; the step
# runs Invoke-FakeWsl in its place, from the path Get-FakeWslPath gives.
function Get-WindowsOptionalFeature {
    [CmdletBinding()]
    param([switch]$Online, [string]$FeatureName)
    if ($script:WslFeatureFault) {
        $PSCmdlet.WriteError((New-FakeErrorRecord 'System.Runtime.InteropServices.COMException' 'Class not registered' `
            'Microsoft.Dism.Commands.GetWindowsOptionalFeatureCommand' 'NotSpecified' $FeatureName))
        return
    }
    if ($script:WslFeatureState) { return [pscustomobject]@{ FeatureName = $FeatureName; State = $script:WslFeatureState } }
}
function Invoke-FakeWsl {
    [void]$script:WslCalls.Add(@($args) -join ' ')
    if ($script:WslThrows) { throw 'The system cannot execute the specified program.' }
    foreach ($l in @($script:WslLines)) { $l }
    $global:LASTEXITCODE = $script:WslExit
}
function Get-FakeWslPath { return $script:FakeWslPath }
# Setup steps 8 and 10 look for hcsdiag.exe when the discovery found none, to
# tell one that is not there from one that could not be read; they look at
# the path Get-FakeHcsDiagPath gives, in the sandbox.
function Get-FakeHcsDiagPath { return $script:FakeHcsDiagPath }
function Get-FakeDockerPath { return (Join-Path $sandbox 'no-docker\Docker.exe') }
# Setup step 17 asks whether it runs as an administrator and as whom, and
# makes its shortcut through WScript.Shell. Each is swapped for these in the
# cut-out step; a shortcut that is saved is in $script:SavedShortcuts.
function Test-FakeAdmin { return [bool]$script:FakeIsAdmin }
function Get-FakeUserName { return 'TESTPC\tester' }
function New-FakeShell {
    $sh = [pscustomobject]@{ Kind = 'FakeShell' }
    $sh | Add-Member -MemberType ScriptMethod -Name CreateShortcut -Value {
        param($LnkPath)
        $lnk = [pscustomobject]@{ Path = $LnkPath; TargetPath = ''; Arguments = ''; WorkingDirectory = ''; Description = ''; IconLocation = ''; WindowStyle = 1 }
        $lnk | Add-Member -MemberType ScriptMethod -Name Save -Value { [void]$script:SavedShortcuts.Add($this) }
        return $lnk
    }
    return $sh
}
function Initialize-ClaudeIcon { return 'fake-claude.ico,0' }
function New-ScheduledTaskAction {
    [CmdletBinding()]
    param([string]$Execute, [string]$Argument)
    if (-not (Test-TaskInstallExpected 'New-ScheduledTaskAction' $Execute)) { throw 'TRIPWIRE: New-ScheduledTaskAction'; return }
    return [pscustomobject]@{ Execute = $Execute; Arguments = $Argument }
}
function New-ScheduledTaskTrigger {
    [CmdletBinding()]
    param([switch]$AtLogOn)
    if (-not (Test-TaskInstallExpected 'New-ScheduledTaskTrigger' '')) { throw 'TRIPWIRE: New-ScheduledTaskTrigger'; return }
    return [pscustomobject]@{ AtLogOn = [bool]$AtLogOn }
}
function New-ScheduledTaskSettingsSet {
    [CmdletBinding()]
    param([switch]$AllowStartIfOnBatteries, [switch]$DontStopIfGoingOnBatteries, [switch]$StartWhenAvailable, [switch]$DontStopOnIdleEnd,
          [int]$RestartCount, [TimeSpan]$RestartInterval, [TimeSpan]$ExecutionTimeLimit)
    if (-not (Test-TaskInstallExpected 'New-ScheduledTaskSettingsSet' '')) { throw 'TRIPWIRE: New-ScheduledTaskSettingsSet'; return }
    return [pscustomobject]@{
        AllowStartIfOnBatteries = [bool]$AllowStartIfOnBatteries; DontStopIfGoingOnBatteries = [bool]$DontStopIfGoingOnBatteries
        StartWhenAvailable = [bool]$StartWhenAvailable; DontStopOnIdleEnd = [bool]$DontStopOnIdleEnd
        RestartCount = $RestartCount; RestartInterval = $RestartInterval; ExecutionTimeLimit = $ExecutionTimeLimit
    }
}
function New-ScheduledTaskPrincipal {
    [CmdletBinding()]
    param([string]$UserId, [string]$RunLevel, [string]$LogonType)
    if (-not (Test-TaskInstallExpected 'New-ScheduledTaskPrincipal' $UserId)) { throw 'TRIPWIRE: New-ScheduledTaskPrincipal'; return }
    return [pscustomobject]@{ UserId = $UserId; RunLevel = $RunLevel; LogonType = $LogonType }
}
function New-FakeTaskFolder {
    param([string]$Path)
    $f = [pscustomobject]@{ Path = $Path }
    $f | Add-Member -MemberType ScriptMethod -Name GetTasks -Value {
        param($Flags)
        $pre = $this.Path.TrimEnd('\') + '\|'
        $hidden = ([int]$Flags -band 1) -ne 0
        return ,@($script:Tasks.Keys | Where-Object {
            $_.StartsWith($pre, [StringComparison]::OrdinalIgnoreCase) -and ($hidden -or @($script:HiddenTasks) -notcontains $_)
        })
    }
    $f | Add-Member -MemberType ScriptMethod -Name GetFolders -Value {
        $pre = $this.Path.TrimEnd('\') + '\'
        return ,@($script:TaskFolders | Where-Object {
            $_.StartsWith($pre, [StringComparison]::OrdinalIgnoreCase) -and $_.Length -gt $pre.Length -and $_.Substring($pre.Length) -notmatch '\\'
        })
    }
    $f | Add-Member -MemberType ScriptMethod -Name DeleteFolder -Value {
        param($Name, $Flags)
        $null = $Flags
        $full = $this.Path.TrimEnd('\') + '\' + $Name
        [void]$script:TaskCalls.Add("DeleteFolder $full")
        if ($script:TaskFolderDeleteFault) {
            throw (New-Object System.Runtime.InteropServices.COMException('Access is denied. (Exception from HRESULT: 0x80070005 (E_ACCESSDENIED))', -2147024891))
        }
        $inside = @($script:Tasks.Keys | Where-Object { $_.StartsWith("$full\|", [StringComparison]::OrdinalIgnoreCase) }).Count +
            @($script:TaskFolders | Where-Object { $_.StartsWith("$full\", [StringComparison]::OrdinalIgnoreCase) }).Count
        if ($inside -gt 0) {
            throw (New-Object System.Runtime.InteropServices.COMException('The directory is not empty. (Exception from HRESULT: 0x80070091)', -2147024751))
        }
        if (-not $script:TaskFolderDeleteNoop) { $script:TaskFolders.Remove($full) }
    }
    return $f
}
function Get-PreventTaskService {
    [void]$script:TaskCalls.Add('Connect')
    if ($script:TaskSvcFault) { throw (New-Object System.UnauthorizedAccessException('Access is denied. (Exception from HRESULT: 0x80070005 (E_ACCESSDENIED))')) }
    $svc = [pscustomobject]@{ FakeTaskService = $true }
    $svc | Add-Member -MemberType ScriptMethod -Name GetFolder -Value {
        param($Path)
        if ($Path -eq '\') { return (New-FakeTaskFolder '\') }
        if ($script:TaskFolderReadFault) { throw (New-Object System.UnauthorizedAccessException('Access is denied. (Exception from HRESULT: 0x80070005 (E_ACCESSDENIED))')) }
        if (@($script:TaskFolders) -notcontains $Path) {
            # $script:TaskFolderGoneReadFault: a folder that is gone cannot be
            # read either, but for a reason other than its absence.
            if ($script:TaskFolderGoneReadFault) { throw (New-Object System.Runtime.InteropServices.COMException('The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)', -2147023174)) }
            throw (New-Object System.IO.FileNotFoundException('The system cannot find the file specified. (Exception from HRESULT: 0x80070002)'))
        }
        return (New-FakeTaskFolder $Path)
    }
    return $svc
}

# ---------------------------------------------------------------------
# Fake Hyper-V: $script:Vms are the VMs Get-VM lists. Set-VMMemory refuses a
# VM that is not off, as Hyper-V does; $script:VmNoop is a change that does
# not take.
# ---------------------------------------------------------------------
# Get-Command answers as Windows' own does, except that a name in
# $script:AbsentCommands is not there, as on a PC without Hyper-V's or
# WinNAT's cmdlets. The fakes stay defined: code that calls one without
# asking first still reaches the fake, and the call is recorded.
$script:AbsentCommands = @()
function Get-Command {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string[]]$Name, [System.Management.Automation.CommandTypes]$CommandType)
    if ($null -ne $Name) {
        $keep = @($Name | Where-Object { @($script:AbsentCommands) -notcontains $_ })
        if ($keep.Count -eq 0) { return }
        $PSBoundParameters['Name'] = [string[]]$keep
    }
    Microsoft.PowerShell.Core\Get-Command @PSBoundParameters
}
function New-FakeVm {
    param([string]$Name, [string]$State = 'Off', [bool]$Dynamic = $false)
    return [pscustomobject]@{ Name = $Name; State = $State; DynamicMemoryEnabled = $Dynamic }
}
function Get-VM {
    [CmdletBinding()]
    param()
    [void]$script:VmCalls.Add('Get-VM')
    if ($script:VmReadFault) { throw 'The Hyper-V Virtual Machine Management service is not running.'; return }
    return @($script:Vms)
}
function Set-VMMemory {
    [CmdletBinding()]
    param($VM, [bool]$DynamicMemoryEnabled)
    [void]$script:VmCalls.Add("Set-VMMemory $($VM.Name) DynamicMemoryEnabled=$DynamicMemoryEnabled")
    if ([string]$VM.State -ne 'Off') {
        throw "Failed to modify device 'Memory'. Dynamic Memory cannot be changed while the virtual machine '$($VM.Name)' is running."
        return
    }
    if (-not $script:VmNoop) { $VM.DynamicMemoryEnabled = $DynamicMemoryEnabled }
}

# ---------------------------------------------------------------------
# Fake network adapter cmdlets. The fixture is the two adapters measured on
# sataana on 2026-11-21: a Realtek 2.5GbE with Energy-Efficient Ethernet,
# Green Ethernet, Gigabit Lite and Power Saving Mode on, selective suspend
# on and PnPCapabilities 24, and a TP-Link USB WiFi with nothing to turn off.
# ---------------------------------------------------------------------
$ethPnp    = 'PCI\VEN_10EC&DEV_8125&SUBSYS_87D71043&REV_05\01000000684CE00000'
$wifiPnp   = 'USB\VID_2357&PID_011F\00e04c000001'
$classRoot = 'HKLM:\SYSTEM\CurrentControlSet\Control\Class'
$netClass  = '{4d36e972-e325-11ce-bfc1-08002be10318}'
$ethKey    = "$classRoot\$netClass\0002"
$wifiKey   = "$classRoot\$netClass\0003"

function New-AdvProp {
    param([string]$Adapter, [string]$Keyword, [string]$DisplayName, $Value, $Valid, $Default, [string]$DisplayValue)
    $v = $null
    if ($null -ne $Value) { $v = [string[]]@($Value) }
    $d = $null
    if ($null -ne $Default) { $d = [string[]]@($Default) }
    $vv = $null
    if ($null -ne $Valid) { $vv = [string[]]@($Valid) }
    return [pscustomobject]@{
        Name                 = $Adapter
        DisplayName          = $DisplayName
        RegistryKeyword      = $Keyword
        RegistryValue        = $v
        DefaultRegistryValue = $d
        ValidRegistryValues  = $vv
        DisplayValue         = $DisplayValue
    }
}
function Set-NicFixture {
    # ConnectorPresent is true on a physical adapter alone, as measured on
    # this PC.
    $script:Nics = @(
        [pscustomobject]@{ Name = 'Ethernet'; InterfaceDescription = 'Realtek Gaming 2.5GbE Family Controller'; PnPDeviceID = $ethPnp; ConnectorPresent = $true },
        [pscustomobject]@{ Name = 'WiFi'; InterfaceDescription = 'TP-Link Wireless USB Adapter'; PnPDeviceID = $wifiPnp; ConnectorPresent = $true }
    )
    # The second instance is a decoy whose name starts with the Realtek's
    # device id without being it: only "<id>_" is the Realtek's own.
    $script:DevEnable = @(
        [pscustomobject]@{ InstanceName = "${ethPnp}_0"; Enable = $true },
        [pscustomobject]@{ InstanceName = "${ethPnp}X_0"; Enable = $true },
        [pscustomobject]@{ InstanceName = "${wifiPnp}_0"; Enable = $false }
    )
    $script:DrvKeys = @{ $ethPnp = "$netClass\0002"; $wifiPnp = "$netClass\0003" }
    Set-FakeRegValue $ethKey 'PnPCapabilities' ([int]24)
    Set-FakeRegValue $ethKey 'DriverDesc' 'Realtek Gaming 2.5GbE Family Controller'
    Set-FakeRegValue $wifiKey 'PnPCapabilities' ([int]24)
    Set-FakeRegValue $wifiKey 'DriverDesc' 'TP-Link Wireless USB Adapter'
    $script:Adv = @{
        'Ethernet' = @(
            (New-AdvProp 'Ethernet' '*EEE'                 'Energy-Efficient Ethernet' '1'    @('0', '1')),
            (New-AdvProp 'Ethernet' '*FlowControl'         'Flow Control'              '3'    @('0', '3')),
            (New-AdvProp 'Ethernet' '*InterruptModeration' 'Interrupt Moderation'      '1'    @('0', '1')),
            (New-AdvProp 'Ethernet' 'AdvancedEEE'          'Advanced EEE'              '0'    @('0', '1')),
            (New-AdvProp 'Ethernet' 'EEEMaxSupportSpeed'   'EEE Max Support Speed'     '1000' @('10', '100', '1000', '2500')),
            (New-AdvProp 'Ethernet' 'EnableGreenEthernet'  'Green Ethernet'            '1'    @('0', '1')),
            (New-AdvProp 'Ethernet' 'GigaLite'             'Gigabit Lite'              '1'    @('0', '1')),
            (New-AdvProp 'Ethernet' 'PnPCapabilities'      ''                          '24'   $null),
            (New-AdvProp 'Ethernet' 'PowerSavingMode'      'Power Saving Mode'         '1'    @('0', '1')),
            (New-AdvProp 'Ethernet' '*WakeOnMagicPacket'   'Wake on Magic Packet'      '0'    @('0', '1')),
            (New-AdvProp 'Ethernet' 'S5WakeOnLan'          'Shutdown Wake-On-Lan'      '1'    @('0', '1'))
        )
        'WiFi'     = @(
            (New-AdvProp 'WiFi' 'DynamicBatchEnable' '' '1'   $null),
            (New-AdvProp 'WiFi' 'HWRTSEnable'        '' '1'   $null),
            (New-AdvProp 'WiFi' 'PnPCapabilities'    '' '24'  $null),
            (New-AdvProp 'WiFi' 'TxDutyEnable'       '' '1'   $null),
            (New-AdvProp 'WiFi' 'Vht24g' 'VHT 2.4G IOT' '255' @('0', '255', '254'))
        )
    }
    $script:Pm = @{
        'Ethernet' = [pscustomobject]@{
            Name = 'Ethernet'; AllowComputerToTurnOffDevice = 'Enabled'; D0PacketCoalescing = 'Unsupported'
            DeviceSleepOnDisconnect = 'Unsupported'; SelectiveSuspend = 'Enabled'; WakeOnMagicPacket = 'Disabled'
        }
        'WiFi'     = [pscustomobject]@{
            Name = 'WiFi'; AllowComputerToTurnOffDevice = 'Disabled'; D0PacketCoalescing = 'Unsupported'
            DeviceSleepOnDisconnect = 'Inactive'; SelectiveSuspend = 'Disabled'; WakeOnMagicPacket = 'Unsupported'
        }
    }
}
# -IncludeHidden lists the adapters that are not physical as well, which
# -Physical never does: the hidden ones, and visible virtual ones such as a
# VPN's adapter, all in $script:HiddenNics, where Hidden says which is
# which. The physical ones are in both lists, as they are in Windows'.
# A list that cannot be read is an error that is not terminating, as the
# CIM cmdlets give, so that a read with SilentlyContinue passes it over and
# one with Stop does not.
function Get-NetAdapter {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string[]]$Name, [switch]$Physical, [switch]$IncludeHidden)
    if ($script:NicThrows) { Write-Error -Message 'The WMI provider is not loaded'; return }
    # $script:PhysicalNicThrows: only the -Physical list cannot be read.
    if ($Physical -and $script:PhysicalNicThrows) { Write-Error -Message 'The WMI provider is not loaded'; return }
    if ($IncludeHidden) {
        if ($script:HiddenNicThrows) { Write-Error -Message 'Generic failure'; return }
        return @($script:Nics) + @($script:HiddenNics)
    }
    return @($script:Nics)
}
function Get-PnpDeviceProperty {
    [CmdletBinding()]
    param([string[]]$InstanceId, [string[]]$KeyName)
    $id = $InstanceId[0]
    if ($KeyName -contains 'DEVPKEY_Device_Driver' -and $script:DrvKeys.ContainsKey($id)) {
        return [pscustomobject]@{ InstanceId = $id; KeyName = 'DEVPKEY_Device_Driver'; Data = $script:DrvKeys[$id] }
    }
    throw 'Element not found.'
}
# The NetAdapter cmdlets read -Name as a wildcard, as Windows' own do: an
# adapter named "Ethernet [LAN]" is a character class there, and the name
# matches nothing, silently. A plain name that matches nothing is an
# ObjectNotFound error.
function Select-FakeNicKey {
    [CmdletBinding()]
    param([string[]]$Keys, [string[]]$Name)
    if (-not $Name) { return ,@($Keys) }
    $out = @()
    foreach ($n in $Name) {
        $hit = @($Keys | Where-Object { $_ -like $n })
        if ($hit.Count -eq 0 -and -not [System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($n)) {
            Write-Error -Message "No MSFT_NetAdapter objects found with property 'Name' equal to '$n'." -Category ObjectNotFound
        }
        $out += $hit
    }
    return ,@($out | Select-Object -Unique)
}
# Windows lists an option with no display name, such as PowerDownPll, only
# with -AllProperties: on the PC the fixture was measured on, the Realtek
# lists 28 options without it and 106 with it.
function Get-NetAdapterAdvancedProperty {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string[]]$Name, [switch]$AllProperties, [string[]]$RegistryKeyword, [string[]]$DisplayName)
    $out = @()
    foreach ($n in (Select-FakeNicKey -Keys @($script:Adv.Keys) -Name $Name)) {
        if (@($script:AdvReadThrows) -contains $n) { throw 'No MSFT_NetAdapterAdvancedPropertySettingData objects found'; return }
        foreach ($p in @($script:Adv[$n])) {
            if (-not $AllProperties -and -not [string]$p.DisplayName) { continue }
            if ($RegistryKeyword -and $RegistryKeyword -notcontains $p.RegistryKeyword) { continue }
            if ($DisplayName -and $DisplayName -notcontains $p.DisplayName) { continue }
            $out += $p
        }
    }
    return $out
}
function Set-NetAdapterAdvancedProperty {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline)]$InputObject, [string[]]$RegistryValue, [string]$DisplayValue, [switch]$NoRestart,
          [Parameter(Position = 0)][string[]]$Name, [string[]]$RegistryKeyword, [string[]]$DisplayName)
    process {
        $targets = @()
        if ($null -ne $InputObject) {
            $targets = @($InputObject)
        } else {
            foreach ($n in (Select-FakeNicKey -Keys @($script:Adv.Keys) -Name $Name)) {
                foreach ($p in @($script:Adv[$n])) {
                    if ($RegistryKeyword -and $RegistryKeyword -notcontains $p.RegistryKeyword) { continue }
                    if ($DisplayName -and $DisplayName -notcontains $p.DisplayName) { continue }
                    $targets += $p
                }
            }
        }
        foreach ($t in $targets) {
            if ($PSBoundParameters.ContainsKey('DisplayValue')) {
                [void]$script:NicCalls.Add("Set-AdvDisplay $($t.Name)|$($t.DisplayName)=$DisplayValue NoRestart=$([bool]$NoRestart)")
                if ($script:AdvDisplayFault) { throw 'The parameter is incorrect.'; return }
                $t.DisplayValue = $DisplayValue
                continue
            }
            $kw = [string]$t.RegistryKeyword
            [void]$script:NicCalls.Add("Set-Adv $($t.Name)|$kw=$($RegistryValue -join ',') NoRestart=$([bool]$NoRestart)")
            $f = $null
            if ($script:AdvFault.ContainsKey($kw)) { $f = $script:AdvFault[$kw] }
            if ($f -eq 'throw') { throw 'The parameter is incorrect.'; return }
            if ($f -ne 'stuck') { $t.RegistryValue = [string[]]@($RegistryValue) }
        }
    }
}
function Reset-NetAdapterAdvancedProperty {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline)]$InputObject, [switch]$NoRestart, [string[]]$Name, [string[]]$RegistryKeyword)
    process {
        [void]$script:NicCalls.Add("Reset-Adv $($InputObject.Name)|$($InputObject.RegistryKeyword) NoRestart=$([bool]$NoRestart)")
        $InputObject.RegistryValue = $null
    }
}
function Get-NetAdapterPowerManagement {
    [CmdletBinding()]
    param([Parameter(Position = 0)][string[]]$Name)
    $out = @()
    foreach ($n in (Select-FakeNicKey -Keys @($script:Pm.Keys) -Name $Name)) {
        if (@($script:PmReadThrows) -contains $n) { throw 'No MSFT_NetAdapterPowerManagementSettingData objects found'; return }
        $out += $script:Pm[$n]
    }
    return $out
}
function Set-NetAdapterPowerManagement {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline)]$InputObject, [Parameter(Position = 0)][string[]]$Name, [switch]$NoRestart,
          [string]$SelectiveSuspend, [string]$DeviceSleepOnDisconnect, [string]$D0PacketCoalescing)
    process {
        $targets = @()
        if ($null -ne $InputObject) {
            $targets = @($InputObject)
        } else {
            foreach ($n in (Select-FakeNicKey -Keys @($script:Pm.Keys) -Name $Name)) { $targets += $script:Pm[$n] }
        }
        foreach ($t in $targets) {
            foreach ($prop in @('SelectiveSuspend', 'DeviceSleepOnDisconnect', 'D0PacketCoalescing')) {
                if (-not $PSBoundParameters.ContainsKey($prop)) { continue }
                $val = $PSBoundParameters[$prop]
                [void]$script:NicCalls.Add("Set-Pm $($t.Name)|$prop=$val NoRestart=$([bool]$NoRestart)")
                if ($script:PmFault.ContainsKey($prop)) { throw 'The request is not supported.'; return }
                $t.$prop = $val
            }
        }
    }
}
# Renames an adapter in the fixture, everywhere the cmdlets above read it.
function Rename-FakeNic {
    param([string]$Old, [string]$New)
    foreach ($n in $script:Nics) { if ($n.Name -eq $Old) { $n.Name = $New } }
    if ($script:Adv.ContainsKey($Old)) {
        $a = $script:Adv[$Old]
        $script:Adv.Remove($Old)
        foreach ($p in @($a)) { $p.Name = $New }
        $script:Adv[$New] = $a
    }
    if ($script:Pm.ContainsKey($Old)) {
        $p = $script:Pm[$Old]
        $script:Pm.Remove($Old)
        $p.Name = $New
        $script:Pm[$New] = $p
    }
}

# ---------------------------------------------------------------------
# Log and Write-WatchLog keep what they are given; the monitor's cached
# hcsdiag list answers from the same fake.
# ---------------------------------------------------------------------
function Log {
    param([string]$Message, [string]$Colour = "White", [switch]$Indent)
    $null = $Colour, $Indent
    [void]$script:LogLines.Add($Message)
}
function Write-WatchLog {
    param([string]$Message, [string]$Level)
    $null = $Level
    [void]$script:WatchLog.Add($Message)
}
function Get-HcsListCached {
    param([int]$MaxAgeSeconds = 25)
    $null = $MaxAgeSeconds
    if ($script:HcsListDown) { return $null }
    return (Format-FakeHcsList)
}
# The monitor's Invoke-HcsDiag, on the fake hcsdiag, to its contract: the
# output when it exited 0, '' when it printed nothing, $null otherwise.
function Invoke-HcsDiag {
    param([string[]]$Arguments, [int]$TimeoutSeconds = 15)
    $null = $TimeoutSeconds
    $out = @(Invoke-FakeHcs @Arguments)
    if ($LASTEXITCODE -ne 0) { return $null }
    return (@($out | ForEach-Object { [string]$_ }) -join "`n")
}

function Reset-State {
    $script:StepResults    = New-Object System.Collections.ArrayList
    $script:CurrentStep    = $null
    $script:Aborted        = $false
    $script:UndoRecord     = New-Object System.Collections.ArrayList
    $script:LogLines       = New-Object System.Collections.ArrayList
    $script:WatchLog       = New-Object System.Collections.ArrayList
    $script:Calls          = New-Object System.Collections.ArrayList
    $script:NicCalls       = New-Object System.Collections.ArrayList
    $script:PcCalls        = New-Object System.Collections.ArrayList
    $script:ScCalls        = New-Object System.Collections.ArrayList
    $script:W32Calls       = New-Object System.Collections.ArrayList
    $script:HcsCalls       = New-Object System.Collections.ArrayList
    $script:MpCalls        = New-Object System.Collections.ArrayList
    $script:SvcCalls       = New-Object System.Collections.ArrayList
    $script:CimCalls       = New-Object System.Collections.ArrayList
    $script:Reg            = @{}
    $script:RegFault       = @{}
    $script:RegNewKeyFault = $false
    $script:Slept          = 0
    $script:Plans          = [ordered]@{}
    $script:ActivePlan     = $null
    $script:PcIdx          = @{}
    $script:PcRefuse       = @{}
    $script:PcLang         = 'en'
    $script:PcTruncate     = $false
    $script:NoTemplate     = $false
    $script:SetActiveRefuse = $false
    $script:DeleteRefuse   = $false
    $script:HibernateRefuse = $false
    $script:PowerAText     = ''
    $script:Nics           = @()
    $script:NicThrows      = $false
    $script:HiddenNics     = @()
    $script:HiddenNicThrows = $false
    $script:DevEnable      = @()
    $script:DevEnableThrows = $false
    $script:SetCimThrows   = $false
    $script:SetCimNoop     = $false
    $script:DrvKeys        = @{}
    $script:Adv            = @{}
    $script:AdvReadThrows  = @()
    $script:AdvFault       = @{}
    $script:Pm             = @{}
    $script:PmReadThrows   = @()
    $script:PmFault        = @{}
    $script:Services       = @{}
    $script:StartServiceThrows = $false
    $script:ScFails        = $false
    $script:MpFails        = $false
    $script:Procs          = @()
    $script:VmWorkers      = @()
    $script:VmWorkersFault = $false
    $script:HcsLive        = New-Object System.Collections.ArrayList
    $script:HcsStuck       = @()
    $script:HcsListCode    = 0
    $script:HcsKillCode    = 0
    $script:HcsListDown    = $false
    $script:HcsListDownAfterKill = $false
    $script:Killed         = New-Object System.Collections.ArrayList
    $script:ChartLines     = @()
    $script:ResyncCode     = 0
    $script:ResyncTakesMs  = 0
    $script:ResyncDoneAt   = [datetime]::MinValue
    $script:JobHangs       = $false
    $script:DnsOkAfter     = 1
    $script:DnsTries       = 0
    $script:LastTimeSyncCheck = [datetime]::MinValue
    $script:UndoMode       = $false
    $script:UndoUnread     = $false
    $script:UndoForeign    = $null
    $script:UndoForeignAside = $null
    $script:ListOk         = -1
    $script:QhFails        = $false
    $script:UndoForeignByGuid = $false
    $script:ResyncBackoffUntil = [datetime]::MinValue
    $script:ChartSeq       = New-Object System.Collections.ArrayList
    $script:TimeSource     = 'time.windows.com,0x9'
    $script:FakeClockMs    = [long]0
    $script:DnsHang        = $false
    $script:DnsLatencyMs   = 0
    $script:JobHangMatch   = $null
    $script:JobThrows      = $null
    $script:JobFails       = $null
    $script:JobFailMatch   = $null
    $script:JobFailSkip    = 0
    $script:JobNoExit      = $null
    $script:CheckErrorsSaid = @{}
    $script:LastHcsListTime = [datetime]::MinValue
    $script:HcsOnlyCowork  = $false
    $script:MpPref         = [pscustomobject]@{ Path = (New-Object System.Collections.ArrayList); Process = (New-Object System.Collections.ArrayList) }
    $script:MpHidePath     = $false
    $script:MpHideProcess  = $false
    $script:MpAdds         = New-Object System.Collections.ArrayList
    $script:MpReadFails    = $false
    $script:AvProducts     = @()
    $script:PsProcs        = @()
    $script:PsProcsFault   = $false
    $script:ProcCalls      = New-Object System.Collections.ArrayList
    $script:StopProcFault  = $false
    $script:StopProcLinger = $false
    $script:Tasks          = @{}
    $script:TaskFolders    = New-Object System.Collections.ArrayList
    $script:TaskCalls      = New-Object System.Collections.ArrayList
    $script:TaskQueryFault = $false
    $script:UnregisterFault = $false
    $script:UnregisterNoop = $false
    $script:TaskSvcFault   = $false
    $script:TaskFolderReadFault = $false
    $script:TaskFolderDeleteFault = $false
    $script:TaskFolderDeleteNoop = $false
    $script:TaskFolderGoneReadFault = $false
    $script:HiddenTasks    = New-Object System.Collections.ArrayList
    $script:Vms            = @()
    $script:VmCalls        = New-Object System.Collections.ArrayList
    $script:VmReadFault    = $false
    $script:VmNoop         = $false
    $script:AbsentCommands = @()
    $script:Nats           = @()
    $script:NatCalls       = New-Object System.Collections.ArrayList
    $script:NatReadFault   = $false
    $script:NatReadCim     = $null
    $script:NatRemoveFault = $false
    $script:NatRemoveExpected = $false
    $script:FileRemoveFault = $null
    $script:FileRemoveNoop = $null
    $script:WakeEnable     = @()
    $script:WakeReadCim    = $null
    $script:AdvDisplayFault = $false
    $script:RegRemoveFault = @{}
    $script:PhysicalNicThrows = $false
    $script:AvReadCim      = $null
    $script:AvReadFault    = $false
    $script:Drives         = @{}
    $script:LogicalDisks   = @{}
    $script:LogicalDiskFault = @()
    $script:StorageFault   = $null
    $script:StorageCalls   = New-Object System.Collections.ArrayList
    $script:TaskInstallExpected = $false
    $script:RegisterFault  = $false
    $script:StartTaskFault = $false
    $script:RegisteredTask = $null
    $script:FakeCommandPath = $null
    $script:ClaudeEnv.VmCachePath = $null
    $script:ServiceReadFault = @()
    $script:ProcReadFault = @()
    $script:ItemReadFault = $null
    $script:ItemReadFaultAfterRemove = $null
    $script:RegReadFault = $null
    $script:RegItemReadFault = $null
    $script:SetContentFault = $null
    $script:TaskStopReads = @{}
    $script:StartTaskNoop = $false
    $script:TaskQueryFaultAfterStart = $false
    $script:TaskGoneAfterStart = $false
    $script:WslFeatureFault = $false
    $script:WslFeatureState = $null
    $script:WslLines = @()
    $script:WslExit = 0
    $script:WslThrows = $false
    $script:WslCalls = New-Object System.Collections.ArrayList
    $script:FakeWslPath = Join-Path $sandbox 'no-wsl\wsl.exe'
    $script:FakeHcsDiagPath = Join-Path $sandbox 'no-hcs\hcsdiag.exe'
    # -Undo step 2 reads $planRestored, which step 1 sets once it has put
    # your plan back.
    $script:FakePlanRestored = $false
    $script:FakeIsAdmin = $true
    $script:SavedShortcuts = New-Object System.Collections.ArrayList
    foreach ($f in @($UndoDir, $BackupFile, $CsBackupFile, $ClaudeAppData, $fakeDesktop, $fakeAppData)) {
        if (Microsoft.PowerShell.Management\Test-Path -LiteralPath $f) {
            Microsoft.PowerShell.Management\Remove-Item -LiteralPath $f -Recurse -Force
        }
    }
    foreach ($d in @($ClaudeAppData, $UndoDir, $fakeDesktop, $fakeAppData)) {
        if (-not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $d)) {
            $null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $d -Force
        }
    }
}
Reset-State

# ---------------------------------------------------------------------
# Refuse to run unless every command that could reach the machine is a fake.
# ---------------------------------------------------------------------
$mustBeFakes = @(
    'powercfg.exe', 'sc.exe', 'w32tm', 'Test-Path', 'Get-Item', 'Get-ChildItem', 'New-Item', 'Set-ItemProperty',
    'Remove-ItemProperty', 'Remove-Item', 'Get-ItemProperty', 'Start-Job', 'Wait-Job', 'Receive-Job', 'Stop-Job', 'Remove-Job',
    'Start-Sleep', 'Get-Service', 'Start-Service', 'Get-MpPreference', 'Add-MpPreference', 'Remove-MpPreference',
    'Get-Process', 'Stop-Process', 'Get-CimInstance', 'Set-CimInstance', 'Get-NetAdapter', 'Get-PnpDeviceProperty',
    'Get-NetAdapterAdvancedProperty', 'Set-NetAdapterAdvancedProperty', 'Reset-NetAdapterAdvancedProperty',
    'Get-NetAdapterPowerManagement', 'Set-NetAdapterPowerManagement', 'Get-NetNat', 'Remove-NetNat', 'Get-ScheduledTask',
    'Unregister-ScheduledTask', 'Get-PreventTaskService', 'Get-VM', 'Set-VMMemory', 'Get-Volume', 'Get-Partition', 'Get-Disk',
    'Register-ScheduledTask', 'Start-ScheduledTask', 'New-ScheduledTaskAction', 'New-ScheduledTaskTrigger', 'New-ScheduledTaskSettingsSet',
    'New-ScheduledTaskPrincipal', 'Invoke-HcsDiag', 'Get-HcsListCached', 'Set-Content', 'Get-WindowsOptionalFeature',
    'Initialize-ClaudeIcon'
) + $tripwires
$notFaked = @()
foreach ($c in $mustBeFakes) {
    $cmd = Get-Command $c -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -eq $cmd -or $cmd.CommandType -ne 'Function') { $notFaked += $c }
}
if ($notFaked.Count -gt 0) {
    Write-Host "Refusing to run: these do not resolve to this test's fakes: $($notFaked -join ', ')" -ForegroundColor Red
    exit 1
}

# ---------------------------------------------------------------------
# Lift Prevent's functions and constants, the monitor's two step functions,
# and the steps that are not functions.
# ---------------------------------------------------------------------
Write-Host ""
Write-Host "--- Lifting the code under test"
$preventFns = @(
    'Initialize-ClaudeAppData', 'ConvertTo-StringArray', 'Step', 'Add-StepResult', 'Add-DiscoveryReadError', 'Write-StepSummary',
    'Get-PreventExitCode', 'Read-PreventUndo', 'Save-PreventUndo', 'Get-PreventUndoItem', 'Add-PreventUndoItem',
    'Clear-PreventUndoItem', 'Invoke-PowerCfg', 'Get-ActivePlanGuid', 'Test-PowerPlanExists',
    'Get-UltimatePlanGuids', 'Get-AcPowerIndex', 'Write-AcPowerValue', 'Get-ModernStandbyState',
    'Get-RegistryValue', 'Write-RecordedDword', 'Get-ServiceFailureBytes', 'ConvertFrom-FailureActionsBlob',
    'Test-RestartRecovery', 'Get-CloudSyncFolderName', 'Get-PowerPlanName', 'Complete-Step',
    'Get-NicAdvancedValue', 'Disable-NicPowerSaving', 'Get-PreventNic', 'Restore-PreventUndoItem',
    'Get-CoworkHcsGuids', 'Get-CoworkWorkerProcessIds', 'Get-VmWorkerGuid', 'Get-CoworkHcsSurvivors',
    'Get-W32tmOffsetSeconds', 'Write-Box', 'Write-UndoOutcome', 'Get-PreventUndoItemLabel', 'Move-ForeignPreventUndo',
    'Get-PreventMachineGuid', 'Test-DefenderListHidden', 'Test-DefenderExclusionCovers', 'Get-PreventNicAdvanced',
    'Get-PreventNicPower', 'Test-PowerSettingGone', 'Add-LegacyInstallRecord', 'Get-LegacyInstallTrace',
    'Get-ClaudeVmWithoutDynamicMemory', 'Remove-PreventTask', 'Remove-PreventTaskFolder', 'Remove-PreventFile',
    'Test-CoworkNat', 'Test-WinNatMissingError', 'Install-HealthMonitorTask', 'Test-CimClassMissingError',
    'Test-ClaudeNotFoundError', 'Get-ClaudeErrorText', 'Get-ClaudeServiceOrNull', 'Get-ClaudeProcessByName',
    'Get-ClaudeScheduledTaskOrNull', 'Get-ClaudeItemOrNull', 'Get-PreventProcessByIdOrNull', 'Wait-PreventProcessExit',
    'Get-PreventTaskState', 'Test-PreventPathThere', 'Find-PreventKitFile', 'Remove-PreventItem'
)
# Get-PreventTaskService is never lifted: it connects to the real Task
# Scheduler, and the fake of that name above stands in for it.
$notLifted = @()
foreach ($fn in $preventFns) {
    $def = Find-FunctionDef $pAst $fn
    if ($null -eq $def) { $notLifted += "Prevent:$fn"; continue }
    . ([scriptblock]::Create($def.Extent.Text))
}
foreach ($fn in @('Set-VmWorkerPriority', 'Stop-StaleCoworkVm', 'Test-TimeSyncHealth', 'Measure-ClockOffset', 'Get-TimeSourceName',
                  'Get-CheckErrorText', 'Test-CheckErrorDue', 'Write-CheckError')) {
    $def = Find-FunctionDef $wAst $fn
    if ($null -eq $def) { $notLifted += "Watch:$fn"; continue }
    . ([scriptblock]::Create($def.Extent.Text))
}
# The monitor runs w32tm.exe by its full path, in a job. The lifted copy
# runs the fake by name instead: the path in its parameter's default is
# replaced, and it must be there exactly once. Otherwise a stub that throws
# stands in, and the lift is reported as failed.
$w32tmPathText = '"$env:SystemRoot\System32\w32tm.exe"'
$def = Find-FunctionDef $wAst 'Invoke-W32tm'
$hits = 0
if ($null -ne $def) { $hits = ([regex]::Matches($def.Extent.Text, [regex]::Escape($w32tmPathText))).Count }
if ($hits -eq 1) {
    . ([scriptblock]::Create($def.Extent.Text.Replace($w32tmPathText, "'w32tm'")))
} else {
    $notLifted += "Watch:Invoke-W32tm (its w32tm.exe path appears $hits time(s), not once)"
    . ([scriptblock]::Create("function Invoke-W32tm { throw 'Invoke-W32tm was not lifted' }"))
}
foreach ($left in @('$GuidPattern', '$UltimateTemplateGuid', '$HighPerfGuid', '$BalancedGuid', '$PowerOptions',
                    '$NicPowerOff', '$NicHiddenPowerOff', '$script:StatusRank')) {
    $as = Find-TopAssignment $pAst $left
    if ($null -eq $as) { $notLifted += "Prevent:$left"; continue }
    . ([scriptblock]::Create($as.Extent.Text))
}
Assert-Case "every function and constant was found" ($notLifted.Count -eq 0) "missing: $($notLifted -join ', ')"

# Two lifted names would reach the machine if the fakes did not hold: the
# Task Scheduler's COM service, which Prevent's own Get-PreventTaskService
# connects to, and the real w32tm.exe, which the monitor's Invoke-W32tm
# runs by its path. Refuse to go on unless both are the harness's.
$escapes = @()
$tsCmd = Get-Command Get-PreventTaskService -CommandType Function -ErrorAction SilentlyContinue
if ($null -eq $tsCmd -or $tsCmd.ScriptBlock.ToString() -match 'ComObject') {
    $escapes += 'Get-PreventTaskService is not the fake'
}
$iwCmd = Get-Command Invoke-W32tm -CommandType Function -ErrorAction SilentlyContinue
if ($null -eq $iwCmd) {
    $escapes += 'Invoke-W32tm is not defined'
} else {
    $iwAst = $iwCmd.ScriptBlock.Ast
    foreach ($pa in @($iwAst.FindAll({
        param($n)
        $n -is [System.Management.Automation.Language.ParameterAst] -and $n.Name.VariablePath.UserPath -eq 'W32tmPath'
    }, $true))) {
        $dv = '<none>'
        if ($null -ne $pa.DefaultValue) { $dv = $pa.DefaultValue.Extent.Text }
        if ($dv -ne "'w32tm'") { $escapes += "Invoke-W32tm's W32tmPath defaults to $dv, not 'w32tm'" }
    }
    $exeNamed = @($iwAst.FindAll({
        param($n)
        ($n -is [System.Management.Automation.Language.StringConstantExpressionAst] -or
         $n -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) -and $n.Value -match '(?i)w32tm\.exe'
    }, $true))
    if ($exeNamed.Count -gt 0) { $escapes += "Invoke-W32tm still names w32tm.exe: $($exeNamed[0].Extent.Text)" }
}
if ($escapes.Count -gt 0) {
    Write-Host "Refusing to run: $($escapes -join '; ')" -ForegroundColor Red
    # Best effort: the run exits 1 below either way, and a folder left in
    # %TEMP% changes nothing.
    Microsoft.PowerShell.Management\Remove-Item -LiteralPath $sandbox -Recurse -Force -ErrorAction SilentlyContinue
    exit 1
}

function Get-SourceSlice {
    param([string]$Source, [string]$From, [string]$To)
    $i = $Source.IndexOf($From, [StringComparison]::Ordinal)
    if ($i -lt 0) { return $null }
    if ($Source.IndexOf($From, $i + 1, [StringComparison]::Ordinal) -ge 0) { return $null }
    $j = $Source.IndexOf($To, $i + $From.Length, [StringComparison]::Ordinal)
    if ($j -lt 0) { return $null }
    return $Source.Substring($i, $j - $i)
}
# A static .NET call cannot be shadowed, and a few in these steps reach the
# machine: the clock step's DNS lookup and its Stopwatch, and the Desktop
# and %APPDATA% folders -Undo deletes shortcuts from. Each is swapped, in
# the cut-out text only, for a fake, and a step that still holds one after
# the swaps is refused. %APPDATA% becomes a subexpression, which also
# holds inside a double-quoted string. The clock step must ask DNS exactly once; it is
# swapped in either form, the lookup with a timeout or the blocking one
# older versions made, so the same tests run against either. Setup step
# 15 looks for the monitor's script beside Prevent, by $PSCommandPath, then
# in C:\ClaudeFix and in the user's profile folder: each is swapped for a
# folder in the sandbox, so that no case reads this machine's.
# -Undo step 2 reads $planRestored, which step 1 sets, so its wrapper sets
# it first. The parts of setup steps 15 and 17 cut here are the ones that
# remove an older version's file; a test opens the step before it runs one.
# Step 15 is cut out whole as well.
$stepSwaps = @(
    @{ Old = "[System.Net.Dns]::GetHostAddressesAsync('time.windows.com')"; New = "(Resolve-FakeTimeHostAsync 'time.windows.com')"; Dns = $true }
    @{ Old = "[System.Net.Dns]::GetHostAddresses('time.windows.com')"; New = "(Resolve-FakeTimeHost 'time.windows.com')"; Dns = $true }
    @{ Old = '[System.Diagnostics.Stopwatch]::StartNew()'; New = '(New-FakeStopwatch)'; Dns = $false }
    @{ Old = '[Environment]::GetFolderPath("Desktop")'; New = '(Get-FakeFolder Desktop)'; Dns = $false }
    @{ Old = '$env:APPDATA'; New = '$(Get-FakeFolder AppData)'; Dns = $false }
    @{ Old = '$env:USERPROFILE'; New = '$(Get-FakeFolder Profile)'; Dns = $false }
    @{ Old = '$PSCommandPath'; New = '(Get-FakeCommandPath)'; Dns = $false }
    @{ Old = '"C:\ClaudeFix\Watch-ClaudeHealth.ps1"'; New = '(Join-Path (Get-FakeFolder ClaudeFix) "Watch-ClaudeHealth.ps1")'; Dns = $false }
)
$unswapped = '\[System\.Net\.Dns\]::|\[System\.Diagnostics\.Stopwatch\]::|\[Environment\]::GetFolderPath|\$env:(APPDATA|USERPROFILE|LOCALAPPDATA)|-ComObject|\$PSCommandPath|C:\\ClaudeFix\\'
$stepWrappers = @(
    @{ Fn = 'Invoke-SetupStep1'; Total = 19; From = 'Step 1 $steps "Configuring power plan...'; To = 'Step 2 $steps "Setting the active plan' }
    @{ Fn = 'Invoke-UndoStep1';  Total = 5;  From = 'Step 1 $steps "Putting back the recorded settings...'; To = 'Step 2 $steps "Removing what older' }
    @{ Fn = 'Invoke-UndoStep2';  Total = 5;  From = 'Step 2 $steps "Removing what older'; To = 'Step 3 $steps "Removing scheduled'
       Pre = '$planRestored = [bool]$script:FakePlanRestored' }
    @{ Fn = 'Invoke-UndoStep3';  Total = 5;  From = 'Step 3 $steps "Removing scheduled'; To = 'Step 4 $steps "Removing shortcuts' }
    @{ Fn = 'Invoke-UndoStep4';  Total = 5;  From = 'Step 4 $steps "Removing shortcuts'; To = 'Step 5 $steps "Removing Claude launchers' }
    @{ Fn = 'Invoke-UndoStep5';  Total = 5;  From = 'Step 5 $steps "Removing Claude launchers'; To = '    Write-UndoOutcome' }
    @{ Fn = 'Invoke-Step8';      Total = 19; From = 'Step 8 $steps "Raising the Cowork VM'; To = 'Step 9 $steps'
       Swaps = @( @{ Old = '"$env:SystemRoot\System32\hcsdiag.exe"'; New = '(Get-FakeHcsDiagPath)' } ) }
    @{ Fn = 'Invoke-Step9';      Total = 19; From = 'Step 9 $steps "Configuring Hyper-V compute service recovery'; To = 'Step 10 $steps' }
    @{ Fn = 'Invoke-Step10';     Total = 19; From = 'Step 10 $steps "Removing stale Cowork VMs'; To = 'Step 11 $steps'
       Swaps = @( @{ Old = '"$env:SystemRoot\System32\hcsdiag.exe"'; New = '(Get-FakeHcsDiagPath)' } ) }
    @{ Fn = 'Invoke-Step11';     Total = 19; From = 'Step 11 $steps "Checking workspace storage'; To = 'Step 12 $steps' }
    @{ Fn = 'Invoke-Step12';     Total = 19; From = 'Step 12 $steps "Checking the clock'; To = 'Step 13 $steps'; DnsOnce = $true }
    @{ Fn = 'Invoke-Step13';     Total = 19; From = 'Step 13 $steps "Checking antivirus'; To = 'Step 14 $steps' }
    @{ Fn = 'Invoke-Step14';     Total = 19; From = 'Step 14 $steps "Checking for WSL2'; To = 'Step 15 $steps'
       Swaps = @(
           @{ Old = '(Join-Path $env:ProgramFiles "Docker\Docker\Docker.exe")'; New = '(Get-FakeDockerPath)' }
           @{ Old = 'Join-Path $env:SystemRoot "System32\wsl.exe"'; New = 'Get-FakeWslPath' }
           @{ Old = '& $wslExe -l -v 2>&1'; New = 'Invoke-FakeWsl -l -v 2>&1' }
       ) }
    @{ Fn = 'Invoke-Step15Legacy'; Total = 19; From = '    # The basic watchdog script older versions installed'; To = '    if ($watchScript) {' }
    @{ Fn = 'Invoke-Step15';     Total = 19; From = 'Step 15 $steps "Installing health monitor'; To = 'Step 16 $steps' }
    # Step 16 reads $myDir, which step 15 sets from Prevent's own path.
    @{ Fn = 'Invoke-Step16';     Total = 19; From = 'Step 16 $steps "Creating Fix Claude Desktop shortcuts'; To = 'Step 17 $steps'
       Pre = '$myDir = $null; if ($script:FakeCommandPath) { $myDir = Split-Path $script:FakeCommandPath -Parent }'
       Swaps = @(
           @{ Old = '"C:\ClaudeFix\Fix-ClaudeDesktop.ps1"'; New = '(Join-Path (Get-FakeFolder ClaudeFix) "Fix-ClaudeDesktop.ps1")' }
           @{ Old = 'New-Object -ComObject WScript.Shell'; New = 'New-FakeShell' }
       ) }
    @{ Fn = 'Invoke-Step17Legacy'; Total = 19; From = '# Remove the old .cmd launcher left by versions <= 6.1.0, if present.'
       To = '        # Create desktop shortcut pointing to the launcher' }
    # The launcher payload is a single-quoted here-string that step 17 only
    # encodes, so the %LOCALAPPDATA% and WScript.Shell in it are inert text;
    # they are swapped all the same, so that the check below stays simple.
    @{ Fn = 'Invoke-Step17';     Total = 19; From = 'Step 17 $steps "Configuring Claude Desktop to launch elevated'; To = 'Step 18 $steps'
       Pre = "`$ToolkitVersion = '0.0.0-test'"
       Swaps = @(
           @{ Old = '([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)'
              New = '(Test-FakeAdmin)' }
           @{ Old = '[System.Security.Principal.WindowsIdentity]::GetCurrent().Name'; New = '(Get-FakeUserName)' }
           @{ Old = 'New-Object -ComObject WScript.Shell'; New = 'New-FakeShell' }
           @{ Old = '$env:LOCALAPPDATA'; New = '$env:FAKE_LOCALAPPDATA' }
       ) }
    @{ Fn = 'Invoke-Step19';     Total = 19; From = 'Step 19 $steps "Removing what older'; To = '# Summary' }
)
foreach ($w in $stepWrappers) {
    $slice = Get-SourceSlice $pSrc $w.From $w.To
    $why = ''
    if ($null -eq $slice) {
        $why = "not found between '$($w.From)' and '$($w.To)'"
    } else {
        $hits = 0
        $ownSwaps = @()
        if ($w.ContainsKey('Swaps')) { $ownSwaps = @($w.Swaps) }
        foreach ($sw in $ownSwaps) {
            if ($slice.IndexOf($sw.Old, [StringComparison]::Ordinal) -lt 0) { $why = "it no longer holds $($sw.Old)"; break }
            $slice = $slice.Replace($sw.Old, $sw.New)
        }
        foreach ($sw in $stepSwaps) {
            $n = ([regex]::Matches($slice, [regex]::Escape($sw.Old))).Count
            if ($sw.Dns) { $hits += $n }
            $slice = $slice.Replace($sw.Old, $sw.New)
        }
        if ($why) { }
        elseif ($w.ContainsKey('DnsOnce') -and $hits -ne 1) { $why = "the DNS call appears $hits time(s), not once" }
        elseif ($slice -match $unswapped) { $why = "it still holds $($Matches[0]), which reaches the machine" }
    }
    if ($why) {
        Assert-Case "$($w.Fn): the step was cut out of Prevent" $false $why
        . ([scriptblock]::Create("function $($w.Fn) { throw 'the step was not found in Prevent' }"))
        continue
    }
    $pre = ''
    if ($w.ContainsKey('Pre')) { $pre = "    $($w.Pre)`n" }
    $text = "function $($w.Fn) {`n    [CmdletBinding(SupportsShouldProcess)]`n    param()`n    `$steps = $($w.Total)`n$pre$slice`n}"
    . ([scriptblock]::Create($text))
    Assert-Case "$($w.Fn): the step was cut out of Prevent" $true
}

function Get-LastStep {
    if ($script:StepResults.Count -eq 0) { return $null }
    return $script:StepResults[$script:StepResults.Count - 1]
}
function Get-StepNotes {
    $s = Get-LastStep
    if ($null -eq $s) { return '' }
    return (@($s.Notes) -join ' | ')
}
function Get-StepStatus {
    $s = Get-LastStep
    if ($null -eq $s) { return '<no step>' }
    return [string]$s.Status
}
function Test-LogHas {
    param([string]$Pattern)
    return (@($script:LogLines | Where-Object { $_ -match $Pattern }).Count -gt 0)
}
function Get-UndoKeys {
    return (@($script:UndoRecord | ForEach-Object { "$($_.Kind):$($_.Key)" }) -join ' ; ')
}
function Find-Undo {
    param([string]$Kind, [string]$Key)
    foreach ($it in $script:UndoRecord) { if ($it.Kind -eq $Kind -and $it.Key -eq $Key) { return $it } }
    return $null
}
function Test-UndoFileClean {
    if (-not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $UndoFile)) { return $false }
    $b = [System.IO.File]::ReadAllBytes($UndoFile)
    $bom = $b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF
    $tmp = Microsoft.PowerShell.Management\Test-Path -LiteralPath "$UndoFile.tmp"
    return (-not $bom -and -not $tmp)
}
# The end of -Undo after the steps a case ran: what Write-UndoOutcome
# prints, its box, whether it says that some lists could not be read and
# to run -Undo again, and the exit code -Undo then has.
function Get-UndoEnd {
    $script:UndoMode = $true
    $text = ''
    try { $text = @(Write-UndoOutcome 6>&1 | ForEach-Object { [string]$_ }) -join "`n" } catch { $text = "threw: $($_.Exception.Message)" }
    $box = @($text -split "`n" | Where-Object { $_ -match '^\s*\|\s+UNDO ' } | ForEach-Object { $_.Trim(' ', '|') })
    return [pscustomobject]@{
        Text   = $text
        Box    = ($box -join ' / ')
        Unread = ($text -match 'Some lists could not be read' -and $text -match 'Run -Undo again once that is\s+resolved')
        Exit   = (Get-PreventExitCode)
    }
}
function Get-LegacyInstallCount { return @($script:UndoRecord | Where-Object { $_.Kind -eq 'LegacyInstall' }).Count }
function Add-LegacyTrace {
    # The record setup's Add-LegacyInstallRecord writes.
    Add-PreventUndoItem -Kind LegacyInstall -Key 'Dynamic Memory flag' -Original (Join-Path $ClaudeAppData 'disable-dynamic-memory.flag')
}

# =====================================================================
Invoke-Section 'Step results, the summary and the exit code' {
    Reset-State
    Step 1 19 'one' -Name 'First'
    Add-StepResult -Status Done
    Add-StepResult -Status Unchanged
    Add-StepResult -Status Warning -Detail 'note one'
    Add-StepResult -Status Done -Reboot
    $s1 = Get-LastStep
    Assert-Case "the worst outcome a step reports is the one it keeps" ($s1.Status -eq 'Warning') "got $($s1.Status)"
    Assert-Case "a detail is kept" ((@($s1.Notes) -join '|') -eq 'note one') "got $(@($s1.Notes) -join '|')"
    Assert-Case "-Reboot is kept" ($s1.Reboot -eq $true)
    Step 2 19 'two' -Name 'Second'
    Complete-Step
    Assert-Case "a step that reported nothing changed nothing" ((Get-StepStatus) -eq 'Unchanged') "got $(Get-StepStatus)"
    Step 3 19 'three' -Name 'Third'
    Add-StepResult -Status Skipped
    Add-StepResult -Status Done -Reboot
    Complete-Step
    Assert-Case "Done outranks Skipped, and Complete-Step leaves it" ((Get-StepStatus) -eq 'Done') "got $(Get-StepStatus)"
    Step 4 19 'four' -Name 'Fourth'
    Add-StepResult -Status Skipped -Reboot
    Assert-Case "no step failed: exit code 0" ((Get-PreventExitCode) -eq 0)
    $out = @(Write-StepSummary 6>&1 | ForEach-Object { [string]$_ })
    $text = $out -join "`n"
    Assert-Case "the summary lists what each step did" ($text -match 'What each step did:' -and $text -match ' 1\. First \.+ check the note' -and
        $text -match ' 2\. Second \.+ no change needed' -and $text -match ' 3\. Third \.+ changed' -and $text -match ' 4\. Fourth \.+ skipped') $text
    Assert-Case "and the notes under it" ($text -match '(?m)^\s+note one$')
    $rebootAt = [array]::IndexOf($out, ($out | Where-Object { $_ -match 'Restart Windows for these to take effect:' } | Select-Object -First 1))
    $rebootList = @()
    if ($rebootAt -ge 0) { $rebootList = @($out[($rebootAt + 1)..($out.Count - 1)] | ForEach-Object { $_.Trim() }) }
    Assert-Case "the reboot list holds the steps that changed something that waits for a restart" (($rebootList -join ',') -eq '- First,- Third') "got $($rebootList -join ',')"
    Step 5 19 'five' -Name 'Fifth'
    Add-StepResult -Status Failed -Detail 'it broke'
    Add-StepResult -Status Done
    Assert-Case "Failed is never overwritten" ((Get-StepStatus) -eq 'Failed')
    Assert-Case "a failed step: exit code 1" ((Get-PreventExitCode) -eq 1)
    Reset-State
    $script:Aborted = $true
    Assert-Case "a run that stopped: exit code 1" ((Get-PreventExitCode) -eq 1)
    Reset-State
    Add-StepResult -Status Done
    Assert-Case "a result with no step running is dropped" ($script:StepResults.Count -eq 0)
    $threw = $false
    Step 1 19 'x' -Name 'X'
    try { Add-StepResult -Status Bogus } catch { $threw = $true }
    Assert-Case "an outcome that is not one of the five is refused" $threw

    # The reads the discovery could not make: said at once, and in the
    # summary as a step of their own with a warning. Until 6.3.1 they were
    # not said at all.
    Reset-State
    Add-DiscoveryReadError -ReadErrors @()
    Assert-Case "every read made: nothing said, and no step recorded" ($script:LogLines.Count -eq 0 -and $script:StepResults.Count -eq 0) `
        "said $($script:LogLines.Count), steps $($script:StepResults.Count)"
    Add-DiscoveryReadError -ReadErrors @('Service CoworkVMService could not be read: Access is denied',
        'Data folder candidate C:\x could not be read: Access to the path is denied.')
    Assert-Case "each read that failed is said at once, as a warning" ((@($script:LogLines) -join '|') -eq
        ('[!] Service CoworkVMService could not be read: Access is denied|' +
            '[!] Data folder candidate C:\x could not be read: Access to the path is denied.')) "got $(@($script:LogLines) -join '|')"
    $d = $script:StepResults | Select-Object -First 1
    Assert-Case "and recorded as step 0, a warning, with each read under it" ($script:StepResults.Count -eq 1 -and $d.Num -eq 0 -and
        $d.Status -eq 'Warning' -and @($d.Notes).Count -eq 2) "steps $($script:StepResults.Count)"
    Step 1 19 'one' -Name 'First'
    Add-StepResult -Status Done
    $text = @(Write-StepSummary 6>&1 | ForEach-Object { [string]$_ }) -join "`n"
    Assert-Case "the summary shows it first, with the reads" ($text -match ' 0\. Looking at this PC \.+ check the note' -and
        $text -match '(?m)^\s+Service CoworkVMService could not be read: Access is denied$' -and
        $text.IndexOf('Looking at this PC') -lt $text.IndexOf('First')) $text
    Assert-Case "a warning is not a failure: exit code 0" ((Get-PreventExitCode) -eq 0)
    Reset-State
}

# =====================================================================
Invoke-Section 'The restart list' {
    Reset-State
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Add-StepResult -Status Done -Reboot
    Add-StepResult -Status Failed -Detail 'WiFi: The request is not supported.'
    Step 7 19 'throttle' -Name 'Power throttling'
    Add-StepResult -Status Unchanged -Reboot
    $out = @(Write-StepSummary 6>&1 | ForEach-Object { [string]$_ })
    $at = -1
    for ($i = 0; $i -lt $out.Count; $i++) { if ($out[$i] -match 'Restart Windows for these to take effect:') { $at = $i } }
    $list = @()
    for ($i = $at + 1; $at -ge 0 -and $i -lt $out.Count; $i++) { $list += $out[$i].Trim() }
    Assert-Case "a step that changed one adapter and failed on another is listed, and a step that changed nothing is not" (($list -join ',') -eq '- Network adapter power saving') "got '$($list -join ',')'"
}

# =====================================================================
Invoke-Section 'The undo record' {
    Reset-State
    Assert-Case "no file: an empty record" ($null -eq (Read-PreventUndo) -and $script:UndoRecord.Count -eq 0)
    Add-PreventUndoItem -Kind Registry -Key 'HKLM:\T|X' -Original 100
    Add-PreventUndoItem -Kind Registry -Key 'HKLM:\T|Y' -Original $null
    Add-PreventUndoItem -Kind NicAdvanced -Key 'P|K' -Original ([string[]]@('1'))
    Add-PreventUndoItem -Kind PowerValue -Key 'a|b|c' -Original ([long]-5)
    Add-PreventUndoItem -Kind NicDeviceEnable -Key 'P' -Original $true
    Add-PreventUndoItem -Kind NicPowerManagement -Key 'P|SelectiveSuspend' -Original 'Enabled'
    Add-PreventUndoItem -Kind Registry -Key 'HKLM:\T|X' -Original 999
    Assert-Case "six saves, each replacing the file, all landed" ($script:UndoRecord.Count -eq 6) "got $($script:UndoRecord.Count)"
    Assert-Case "the first record of a setting wins" ((Find-Undo 'Registry' 'HKLM:\T|X').Original -eq 100)
    Assert-Case "the file has no byte order mark and no temporary file is left" (Test-UndoFileClean)
    $script:UndoRecord = New-Object System.Collections.ArrayList
    Read-PreventUndo
    Assert-Case "it reads back with the same six items" ($script:UndoRecord.Count -eq 6) (Get-UndoKeys)
    $x = Find-Undo 'Registry' 'HKLM:\T|X'
    $y = Find-Undo 'Registry' 'HKLM:\T|Y'
    $k = Find-Undo 'NicAdvanced' 'P|K'
    $pv = Find-Undo 'PowerValue' 'a|b|c'
    $de = Find-Undo 'NicDeviceEnable' 'P'
    $pm = Find-Undo 'NicPowerManagement' 'P|SelectiveSuspend'
    Assert-Case "a number comes back" ($null -ne $x -and $x.Original -eq 100)
    Assert-Case "'there was none' comes back as null" ($null -ne $y -and $null -eq $y.Original)
    Assert-Case "a one-value list comes back as a list" ($null -ne $k -and @($k.Original).Count -eq 1 -and @($k.Original)[0] -eq '1')
    Assert-Case "a negative number comes back" ($null -ne $pv -and $pv.Original -eq -5)
    Assert-Case "true comes back" ($null -ne $de -and $de.Original -eq $true)
    Assert-Case "a word comes back" ($null -ne $pm -and $pm.Original -eq 'Enabled')
    foreach ($it in @($script:UndoRecord.ToArray())) { Clear-PreventUndoItem -Kind $it.Kind -Key $it.Key }
    Assert-Case "clearing the last item deletes the file" (-not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $UndoFile))

    [System.IO.File]::WriteAllText($UndoFile, '{ this is not json')
    $msg = ''
    try { Read-PreventUndo } catch { $msg = $_.Exception.Message }
    Assert-Case "a file that is not JSON stops the run" ($msg -match 'could not be read' -and $msg -match 'Nothing was changed') $msg
    [System.IO.File]::WriteAllText($UndoFile, '{ "Version": 2, "Items": [] }')
    $msg = ''
    try { Read-PreventUndo } catch { $msg = $_.Exception.Message }
    Assert-Case "a record of another version stops the run" ($msg -match 'is not a version 1 record') $msg
    [System.IO.File]::WriteAllText($UndoFile, ('{ "Version": 1, "ComputerName": "' + [Environment]::MachineName + '", "Items": [ { "Kind": "Registry", "Original": 1 } ] }'))
    $msg = ''
    try { Read-PreventUndo } catch { $msg = $_.Exception.Message }
    Assert-Case "an item with no key stops the run" ($msg -match 'holds an item with no Kind or Key') $msg

    Reset-State
    $null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $UndoFile -Force
    $msg = ''
    try { Add-PreventUndoItem -Kind Registry -Key 'HKLM:\T|Z' -Original 1 } catch { $msg = $_.Exception.Message }
    Assert-Case "a record that cannot be saved throws, so the change is not made" ($msg -match 'Could not save the undo record' -and $msg -match 'The change was not made') $msg
    Assert-Case "and the item is not kept" ($script:UndoRecord.Count -eq 0)
    Assert-Case "and no temporary file is left" (-not (Microsoft.PowerShell.Management\Test-Path -LiteralPath "$UndoFile.tmp"))
    Microsoft.PowerShell.Management\Remove-Item -LiteralPath $UndoFile -Recurse -Force
}

# =====================================================================
Invoke-Section 'The undo record: where it is kept, and which computer it belongs to' {
    $where = @()
    foreach ($left in @('$LocalAppData', '$UndoDir', '$UndoFile')) {
        $as = Find-TopAssignment $pAst $left
        if ($null -eq $as) { $where += '<none>' } else { $where += $as.Right.Extent.Text }
    }
    $where = $where -join ' ; '
    Assert-Case "it is %LOCALAPPDATA%\ClaudeFix\prevent-undo.json, which does not roam to another computer" ($where -eq
        ("[Environment]::GetFolderPath('LocalApplicationData') ; " + 'Join-Path $LocalAppData "ClaudeFix" ; Join-Path $UndoDir "prevent-undo.json"')) $where

    Reset-State
    Microsoft.PowerShell.Management\Remove-Item -LiteralPath $UndoDir -Recurse -Force
    $owner = '<no file>'
    try { Add-PreventUndoItem -Kind Registry -Key 'HKLM:\T|X' -Original 1 } catch { $owner = $_.Exception.Message }
    if (Microsoft.PowerShell.Management\Test-Path -LiteralPath $UndoFile) {
        $doc = [System.IO.File]::ReadAllText($UndoFile) | ConvertFrom-Json
        $owner = '<no name>'
        if ($doc.PSObject.Properties['ComputerName']) { $owner = [string]$doc.ComputerName }
    }
    Assert-Case "the first save makes its folder, and the record names this computer" ($owner -eq [Environment]::MachineName) "got $owner"

    Reset-State
    $foreign = '{ "Version": 1, "ComputerName": "OTHER-PC", "Items": [ { "Kind": "PowerPlan", "Key": "Active", "Original": "' + $plBal + '" } ] }'
    [System.IO.File]::WriteAllText($UndoFile, $foreign)
    Read-PreventUndo
    Assert-Case "a record that names another computer is not loaded" ($script:UndoRecord.Count -eq 0 -and $script:UndoForeign -eq 'OTHER-PC') "$($script:UndoRecord.Count) item(s), from '$($script:UndoForeign)'"
    $msg = ''
    try { Add-PreventUndoItem -Kind Registry -Key 'HKLM:\T|Y' -Original 1 } catch { $msg = $_.Exception.Message }
    Assert-Case "nor written over: a change that cannot be recorded is not made" ($msg -match 'Could not save the undo record' -and $msg -match 'another computer \(OTHER-PC\)' -and
        [System.IO.File]::ReadAllText($UndoFile) -eq $foreign) $msg
    $script:UndoRecord = New-Object System.Collections.ArrayList
    $msg = ''
    try { Save-PreventUndo } catch { $msg = $_.Exception.Message }
    Assert-Case "nor deleted by a save of an empty record" ((Microsoft.PowerShell.Management\Test-Path -LiteralPath $UndoFile) -and
        [System.IO.File]::ReadAllText($UndoFile) -eq $foreign) $msg

    Reset-State
    $own = '{ "Version": 1, "ComputerName": "' + [Environment]::MachineName.ToLowerInvariant() + '", "Items": [ { "Kind": "Registry", "Key": "HKLM:\\T|X", "Original": 5 } ] }'
    [System.IO.File]::WriteAllText($UndoFile, $own)
    Read-PreventUndo
    Assert-Case "this computer's own record loads, whatever the case of its name" ($script:UndoRecord.Count -eq 1 -and $null -eq $script:UndoForeign)
    Reset-State
    [System.IO.File]::WriteAllText($UndoFile, '{ "Version": 1, "Items": [ { "Kind": "Registry", "Key": "HKLM:\\T|X", "Original": 5 } ] }')
    Read-PreventUndo
    Assert-Case "a record that names no computer is not loaded either" ($script:UndoRecord.Count -eq 0 -and $script:UndoForeign -eq 'unknown') "$($script:UndoRecord.Count) item(s), from '$($script:UndoForeign)'"

    # Setup moves a foreign record aside. The clock is fixed, so two moves in
    # one second are what is tested.
    function Get-Date { return [datetime]'2026-11-21 10:00:00' }
    Reset-State
    [System.IO.File]::WriteAllText($UndoFile, $foreign)
    Read-PreventUndo
    $aside = ''
    $err = ''
    try { $aside = [string](Move-ForeignPreventUndo) } catch { $err = $_.Exception.Message }
    $kept = $false
    if ($aside -and (Microsoft.PowerShell.Management\Test-Path -LiteralPath $aside)) { $kept = [System.IO.File]::ReadAllText($aside) -eq $foreign }
    Assert-Case "setup moves it aside and keeps it whole" ($kept -and -not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $UndoFile)) "$aside $err"
    Assert-Case "beside the record, named for the computer that wrote it" ($aside -and (Split-Path -Path $aside -Parent) -eq $UndoDir -and
        (Split-Path -Path $aside -Leaf) -eq 'prevent-undo.from-OTHER-PC.20261121-100000.json') $aside
    Assert-Case "it says so, and this computer's own record starts empty" ((Test-LogHas 'another computer \(OTHER-PC\)') -and $null -eq $script:UndoForeign -and
        $script:UndoRecord.Count -eq 0) ($script:LogLines -join ' / ')
    $msg = ''
    try { Add-PreventUndoItem -Kind Registry -Key 'HKLM:\T|X' -Original 1 } catch { $msg = $_.Exception.Message }
    Assert-Case "and is saved" ($msg -eq '' -and (Test-UndoFileClean)) $msg
    [System.IO.File]::WriteAllText($UndoFile, $foreign)
    Read-PreventUndo
    $aside2 = ''
    try { $aside2 = [string](Move-ForeignPreventUndo) } catch { $err = $_.Exception.Message }
    Assert-Case "a second one moved aside in the same second does not replace the first" ($aside2 -and (Split-Path -Path $aside2 -Leaf) -eq 'prevent-undo.from-OTHER-PC.20261121-100000-1.json' -and
        (Microsoft.PowerShell.Management\Test-Path -LiteralPath $aside) -and (Microsoft.PowerShell.Management\Test-Path -LiteralPath $aside2)) "$aside2 $err"
    [System.IO.File]::WriteAllText($UndoFile, '{ "Version": 1, "ComputerName": "..\\..\\evil/x:y", "Items": [] }')
    Read-PreventUndo
    $aside3 = ''
    try { $aside3 = [string](Move-ForeignPreventUndo) } catch { $err = $_.Exception.Message }
    Assert-Case "a computer name that is a path stays a file name in the same folder" ($aside3 -and (Split-Path -Path $aside3 -Parent) -eq $UndoDir -and
        (Split-Path -Path $aside3 -Leaf) -eq 'prevent-undo.from-evilxy.20261121-100000.json') "$aside3 $err"

    $modeIf = $pAst.Find({
        param($n)
        $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -eq '$Undo' -and
        $null -ne $n.ElseClause -and $n.ElseClause.Extent.Text.Contains('$steps = 19')
    }, $true)
    $moves = @()
    $step1 = $null
    if ($null -ne $modeIf) {
        $moves = @($modeIf.ElseClause.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and [string]$n.GetCommandName() -eq 'Move-ForeignPreventUndo' }, $true))
        $step1 = $modeIf.ElseClause.Find({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.Extent.Text.StartsWith('Step 1 $steps') }, $true)
    }
    Assert-Case "setup moves a record written on another computer aside before step 1" ($moves.Count -eq 1 -and $null -ne $step1 -and
        $moves[0].Extent.StartOffset -lt $step1.Extent.StartOffset) "$($moves.Count) call(s)"
}

# =====================================================================
Invoke-Section '-Undo that leaves settings recorded' {
    Reset-State
    $script:UndoMode = $true
    Step 1 5 'undo' -Name 'Recorded settings'
    Add-StepResult -Status Warning -Detail "The adapter $ethPnp is not plugged in"
    $null = $script:UndoRecord.Add([pscustomobject]@{ Kind = 'NicAdvanced'; Key = "$ethPnp|*EEE"; Original = [string[]]@('1') })
    $null = $script:UndoRecord.Add([pscustomobject]@{ Kind = 'PowerValue'; Key = "$plBal|$subSleep|$setSleep"; Original = 1800 })
    Assert-Case "settings still recorded: exit code 1" ((Get-PreventExitCode) -eq 1)
    $text = ''
    try { $text = @(Write-UndoOutcome 6>&1 | ForEach-Object { [string]$_ }) -join "`n" } catch { $text = "threw: $($_.Exception.Message)" }
    Assert-Case "the box says UNDO INCOMPLETE, never UNDO COMPLETE" ($text -match 'UNDO INCOMPLETE' -and $text -notmatch 'UNDO COMPLETE') $text
    Assert-Case "each setting still recorded is listed" ($text.Contains("Network adapter ${ethPnp}: its setting *EEE") -and
        $text.Contains("Power setting 'Sleep after' of the plan $plBal")) $text
    Assert-Case "with where the record is, and that -Undo again puts them back, with the adapter plugged in" ($text.Contains($UndoFile) -and
        $text -match 'Run -Undo again' -and $text -match 'puts these back' -and $text -match 'plugged in') $text
    $script:UndoMode = $false
    Assert-Case "setup keeps its record on purpose: exit code 0" ((Get-PreventExitCode) -eq 0)

    Reset-State
    $script:UndoMode = $true
    Step 1 5 'undo' -Name 'Recorded settings'
    Add-StepResult -Status Done
    $text = ''
    try { $text = @(Write-UndoOutcome 6>&1 | ForEach-Object { [string]$_ }) -join "`n" } catch { $text = "threw: $($_.Exception.Message)" }
    Assert-Case "everything put back: UNDO COMPLETE, nothing listed, exit code 0" ($text -match 'UNDO COMPLETE' -and $text -notmatch 'still in the undo record' -and
        (Get-PreventExitCode) -eq 0) $text
    Add-StepResult -Status Failed -Detail 'it broke'
    $text = ''
    try { $text = @(Write-UndoOutcome 6>&1 | ForEach-Object { [string]$_ }) -join "`n" } catch { $text = "threw: $($_.Exception.Message)" }
    Assert-Case "a failed step: UNDO FINISHED WITH ERRORS" ($text -match 'UNDO FINISHED WITH ERRORS' -and $text -notmatch 'UNDO COMPLETE') $text

    # A list -Undo could not read, the NAT rules for one: what older
    # versions left there was not looked at, so -Undo is not complete. It
    # used to end UNDO COMPLETE with exit code 0, which told neither the
    # user nor a calling script to run it again.
    Reset-State
    Step 2 5 'undo' -Name "Older versions' changes"
    Add-StepResult -Status Warning -Detail 'Could not read the NAT rules'
    $script:UndoUnread = $true
    $end = Get-UndoEnd
    Assert-Case "a list -Undo could not read: UNDO INCOMPLETE, never UNDO COMPLETE, and exit code 1" ($end.Box -eq 'UNDO INCOMPLETE' -and $end.Exit -eq 1) $end.Text
    Assert-Case "and it says some lists could not be read, and to run -Undo again once that is resolved" ($end.Unread -and
        $end.Text -notmatch 'still in the undo record') $end.Text
    $script:UndoMode = $false
    Assert-Case "setup does not read it: exit code 0" ((Get-PreventExitCode) -eq 0)

    Reset-State
    Step 1 5 'undo' -Name 'Recorded settings'
    Add-StepResult -Status Done
    $end = Get-UndoEnd
    Assert-Case "every list read and everything put back: UNDO COMPLETE, exit code 0, and no word of lists not read" ($end.Box -eq 'UNDO COMPLETE' -and
        -not $end.Unread -and $end.Text -notmatch 'could not be read' -and $end.Exit -eq 0) $end.Text

    $uu = Find-TopAssignment $pAst '$script:UndoUnread'
    $runTry = @($pAst.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.TryStatementAst] -and $_.Body.Extent.Text.Contains('Read-PreventUndo') })
    $uuText = '<none>'
    if ($null -ne $uu) { $uuText = $uu.Extent.Text }
    Assert-Case "each run starts with it clear, before -Undo or setup begins" ($null -ne $uu -and $uu.Right.Extent.Text -eq '$false' -and $runTry.Count -eq 1 -and
        $uu.Extent.StartOffset -lt $runTry[0].Extent.StartOffset) "$uuText; $($runTry.Count) run block(s)"

    $um = Find-TopAssignment $pAst '$script:UndoMode'
    Assert-Case "the undo mode is the -Undo switch" ($null -ne $um -and $um.Right.Extent.Text -eq '[bool]$Undo')
    $modeIf = $pAst.Find({
        param($n)
        $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -eq '$Undo' -and
        $null -ne $n.ElseClause -and $n.ElseClause.Extent.Text.Contains('$steps = 19')
    }, $true)
    $lastUndo = '<none>'
    if ($null -ne $modeIf) {
        $st = @($modeIf.Clauses[0].Item2.Statements)
        $lastUndo = $st[$st.Count - 1].Extent.Text
    }
    Assert-Case "and -Undo ends with that summary" ($lastUndo -eq 'Write-UndoOutcome') $lastUndo
}

# =====================================================================
Invoke-Section 'A record that cannot be saved after a setting is put back' {
    Reset-State
    Add-PreventUndoItem -Kind Registry -Key 'HKLM:\T|A' -Original 1
    Add-PreventUndoItem -Kind Registry -Key 'HKLM:\T|B' -Original 2
    Add-PreventUndoItem -Kind Registry -Key 'HKLM:\T|C' -Original 3
    Microsoft.PowerShell.Management\Remove-Item -LiteralPath $UndoFile -Force
    $null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $UndoFile -Force
    $threw = $false
    try { Clear-PreventUndoItem -Kind Registry -Key 'HKLM:\T|B' } catch { $threw = $true }
    Assert-Case "clearing it throws, and the setting stays where it was in the record" ($threw -and
        (Get-UndoKeys) -eq 'Registry:HKLM:\T|A ; Registry:HKLM:\T|B ; Registry:HKLM:\T|C') (Get-UndoKeys)
}

# =====================================================================
Invoke-Section 'Write-RecordedDword' {
    Reset-State
    Set-FakeRegValue 'HKLM:\T\K' 'Same' ([int]1)
    Set-FakeRegValue 'HKLM:\T\K' 'Diff' ([int]0)
    Set-FakeRegValue 'HKLM:\T\K' 'Sz' '1'
    $r = Write-RecordedDword -Path 'HKLM:\T\K' -Name 'Same' -Value 1
    Assert-Case "the value it already has: Unchanged, nothing recorded or written" ($r -eq 'Unchanged' -and $script:UndoRecord.Count -eq 0 -and $script:Calls.Count -eq 0) "got $r"
    $r = Write-RecordedDword -Path 'HKLM:\T\K' -Name 'New' -Value 1
    Assert-Case "an absent value: Done, recorded as none" ($r -eq 'Done' -and (Get-FakeRegValue 'HKLM:\T\K' 'New') -eq 1 -and
        $null -ne (Find-Undo 'Registry' 'HKLM:\T\K|New') -and $null -eq (Find-Undo 'Registry' 'HKLM:\T\K|New').Original) "got $r; $(Get-UndoKeys)"
    $r = Write-RecordedDword -Path 'HKLM:\T\K' -Name 'Diff' -Value 1
    Assert-Case "another value: Done, the old one recorded" ($r -eq 'Done' -and (Find-Undo 'Registry' 'HKLM:\T\K|Diff').Original -eq 0)
    $r = Write-RecordedDword -Path 'HKLM:\T\K' -Name 'Sz' -Value 1
    $now = Get-FakeRegValue 'HKLM:\T\K' 'Sz'
    Assert-Case "a string '1' is not the DWORD 1: rewritten as a DWORD" ($r -eq 'Done' -and $now -is [int] -and $now -eq 1 -and
        (Find-Undo 'Registry' 'HKLM:\T\K|Sz').Original -eq '1') "got $r"
    $r = Write-RecordedDword -Path 'HKLM:\T\Missing\Sub' -Name 'V' -Value 7
    Assert-Case "a missing key is created and recorded, before the value" ($r -eq 'Done' -and (Test-Path -LiteralPath 'HKLM:\T\Missing\Sub') -and
        (Get-UndoKeys) -match 'RegistryKey:HKLM:\\T\\Missing\\Sub ; Registry:HKLM:\\T\\Missing\\Sub\|V') (Get-UndoKeys)
    $script:RegFault['hklm:\t\k|boom'] = 'throw'
    $r = Write-RecordedDword -Path 'HKLM:\T\K' -Name 'Boom' -Value 1
    Assert-Case "a write that throws: Failed" ($r -eq 'Failed') "got $r"
    $script:RegFault['hklm:\t\k|quiet'] = 'ignore'
    $r = Write-RecordedDword -Path 'HKLM:\T\K' -Name 'Quiet' -Value 1
    Assert-Case "a write that does not take: Failed (read back)" ($r -eq 'Failed') "got $r"
    $script:RegNewKeyFault = $true
    $r = Write-RecordedDword -Path 'HKLM:\T\Other' -Name 'V' -Value 1
    Assert-Case "a key that cannot be created: Failed" ($r -eq 'Failed') "got $r"
}

# =====================================================================
Invoke-Section 'Power settings: reading and writing one AC value' {
    Reset-State
    Set-PowerFixture
    $script:PcIdx["$plBal|$subSleep|$setSleep"] = [long]1800
    $v = Get-AcPowerIndex -Plan $plBal -Sub $subSleep -Setting $setSleep
    Assert-Case "the AC value is read, not the DC value" ($v -eq 1800) "got $v"
    $script:PcLang = 'nl'
    $v = Get-AcPowerIndex -Plan $plBal -Sub $subSleep -Setting $setSleep
    Assert-Case "on a Dutch Windows too" ($v -eq 1800) "got $v"
    $script:PcLang = 'en'
    $script:PcTruncate = $true
    $v = Get-AcPowerIndex -Plan $plBal -Sub $subSleep -Setting $setSleep
    Assert-Case "output without the DC line is not read" ($null -eq $v) "got $v"
    $script:PcTruncate = $false
    $v = Get-AcPowerIndex -Plan $plBal -Sub $subSleep -Setting 'deadbeef-0000-0000-0000-000000000000'
    Assert-Case "a setting this machine does not have: null" ($null -eq $v) "got $v"
    $r = Write-AcPowerValue -Plan $plBal -Sub $subSleep -Setting $setSleep -Value 0
    Assert-Case "a value set and read back: Done" ($r -eq 'Done' -and $script:PcIdx["$plBal|$subSleep|$setSleep"] -eq 0) "got $r"
    Assert-Case "the value it had is recorded" ((Find-Undo 'PowerValue' "$plBal|$subSleep|$setSleep").Original -eq 1800)
    $r = Write-AcPowerValue -Plan $plBal -Sub $subSleep -Setting $setSleep -Value 0
    Assert-Case "set again: Unchanged" ($r -eq 'Unchanged') "got $r"
    $r = Write-AcPowerValue -Plan $plBal -Sub $subSleep -Setting 'deadbeef-0000-0000-0000-000000000000' -Value 0
    Assert-Case "a setting this machine does not have: Missing" ($r -eq 'Missing') "got $r"
    $script:PcIdx["$plBal|$subSleep|setting-two"] = [long]5
    $script:PcRefuse["$plBal|$subSleep|setting-two"] = $true
    $r = Write-AcPowerValue -Plan $plBal -Sub $subSleep -Setting 'setting-two' -Value 0
    Assert-Case "a value powercfg refuses: Failed" ($r -eq 'Failed') "got $r"
}

# =====================================================================
Invoke-Section 'Get-ModernStandbyState' {
    $sataana = @(
        'The following sleep states are available on this system:',
        '    Standby (S3)',
        '',
        'The following sleep states are not available on this system:',
        '    Standby (S1)',
        "`tThe system firmware does not support this standby state.",
        '',
        '    Hibernate',
        "`tAn internal system component has disabled hibernation.",
        "`t`tGuarded Host",
        '',
        '    Standby (S0 Low Power Idle)',
        "`tThe system firmware does not support this standby state."
    ) -join "`r`n"
    Assert-Case "the machine measured on 2026-11-21: Disabled" ((Get-ModernStandbyState -Text $sataana) -eq 'Disabled')
    $s0 = @(
        'The following sleep states are available on this system:',
        '    Standby (S0 Low Power Idle) Network Connected',
        '    Hibernate',
        '',
        'The following sleep states are not available on this system:',
        '    Standby (S3)'
    ) -join "`n"
    Assert-Case "S0 in the first block: Enabled" ((Get-ModernStandbyState -Text $s0) -eq 'Enabled')
    $nl = @(
        'De volgende slaapstanden zijn beschikbaar op dit systeem:',
        '    Stand-by (S0 Laag energieverbruik inactief) Netwerk verbonden',
        '',
        'De volgende slaapstanden zijn niet beschikbaar op dit systeem:',
        '    Stand-by (S3)'
    ) -join "`n"
    Assert-Case "translated headers: Enabled" ((Get-ModernStandbyState -Text $nl) -eq 'Enabled')
    Assert-Case "no text: null" ($null -eq (Get-ModernStandbyState -Text ''))
    Assert-Case "no header: null" ($null -eq (Get-ModernStandbyState -Text "    Standby (S3)`n    Hibernate"))
}

# =====================================================================
Invoke-Section 'Get-ModernStandbyState: S0 is decided by its own line' {
    $onlyNot = @(
        'The following sleep states are not available on this system:',
        '    Standby (S1)',
        "`tThe system firmware does not support this standby state.",
        '    Standby (S0 Low Power Idle)',
        "`tThe system firmware does not support this standby state.",
        '    Hibernate',
        "`tHibernation has not been enabled."
    ) -join "`r`n"
    Assert-Case "only the not-available block, S0 with its reason: Disabled" ((Get-ModernStandbyState -Text $onlyNot) -eq 'Disabled') "got $(Get-ModernStandbyState -Text $onlyNot)"
    $s0Last = @(
        'The following sleep states are not available on this system:',
        '    Standby (S3)',
        "`tThe system firmware does not support this standby state.",
        '    Standby (S0 Low Power Idle)',
        "`tThe system firmware does not support this standby state."
    ) -join "`n"
    Assert-Case "the same with S0 last: Disabled" ((Get-ModernStandbyState -Text $s0Last) -eq 'Disabled') "got $(Get-ModernStandbyState -Text $s0Last)"
    $nlNot = @(
        'De volgende slaapstanden zijn niet beschikbaar op dit systeem:',
        '    Stand-by (S0 Laag energieverbruik inactief)',
        "`tDe systeemfirmware ondersteunt deze stand-bymodus niet.",
        '    Stand-by (S3)',
        "`tDe systeemfirmware ondersteunt deze stand-bymodus niet."
    ) -join "`r`n"
    Assert-Case "the same in Dutch: Disabled" ((Get-ModernStandbyState -Text $nlNot) -eq 'Disabled') "got $(Get-ModernStandbyState -Text $nlNot)"
    $spaces = @(
        'The following sleep states are not available on this system:',
        '    Standby (S0 Low Power Idle)',
        '        The system firmware does not support this standby state.'
    ) -join "`n"
    Assert-Case "a reason indented with spaces, not a tab: Disabled" ((Get-ModernStandbyState -Text $spaces) -eq 'Disabled') "got $(Get-ModernStandbyState -Text $spaces)"
    $both = @(
        'The following sleep states are available on this system:',
        '    Standby (S0 Low Power Idle) Network Connected',
        '    Hibernate',
        '    Fast Startup',
        '',
        'The following sleep states are not available on this system:',
        '    Standby (S1)',
        "`tThis standby state is disabled when S0 low power idle is supported.",
        '    Standby (S3)',
        "`tThis standby state is disabled when S0 low power idle is supported."
    ) -join "`r`n"
    Assert-Case "S0 available, with a not-available block below: Enabled" ((Get-ModernStandbyState -Text $both) -eq 'Enabled')
    $bothNot = @(
        'The following sleep states are available on this system:',
        '    Standby (S3)',
        '    Hibernate',
        '',
        'The following sleep states are not available on this system:',
        '    Standby (S1)',
        "`tThe system firmware does not support this standby state.",
        '    Standby (S0 Low Power Idle)',
        "`tThe system firmware does not support this standby state."
    ) -join "`r`n"
    Assert-Case "S0 not available, with an available block above: Disabled" ((Get-ModernStandbyState -Text $bothNot) -eq 'Disabled')
    $availOnly = @(
        'The following sleep states are available on this system:',
        '    Standby (S0 Low Power Idle) Network Connected',
        '    Hibernate'
    ) -join "`n"
    Assert-Case "only the available block, with S0: Enabled" ((Get-ModernStandbyState -Text $availOnly) -eq 'Enabled')
    $s0EndOfFirst = @(
        'The following sleep states are available on this system:',
        '    Hibernate',
        '    Standby (S0 Low Power Idle) Network Disconnected',
        'The following sleep states are not available on this system:',
        '    Standby (S3)',
        "`tThis standby state is disabled when S0 low power idle is supported."
    ) -join "`n"
    Assert-Case "S0 last in its block, the next header right below it: Enabled" ((Get-ModernStandbyState -Text $s0EndOfFirst) -eq 'Enabled')
    $de = @(
        'Die folgenden Standbyzustaende sind auf diesem System verfuegbar:',
        '    Standby (S0 Niedriger Energiezustand im Leerlauf) Netzwerk verbunden',
        '    Ruhezustand',
        '',
        'Die folgenden Standbyzustaende sind auf diesem System nicht verfuegbar:',
        '    Standby (S3)',
        "`tDieser Standbyzustand ist deaktiviert, wenn S0 im Leerlauf mit niedrigem Energieverbrauch unterstuetzt wird."
    ) -join "`r`n"
    Assert-Case "German headers, S0 available: Enabled" ((Get-ModernStandbyState -Text $de) -eq 'Enabled')
    $reasonNamesS0 = @(
        'The following sleep states are available on this system:',
        '    Standby (S3)',
        '',
        'The following sleep states are not available on this system:',
        '    Standby (S1)',
        "`tThis standby state is disabled when Standby (S0 Low Power Idle) is supported."
    ) -join "`n"
    Assert-Case "a reason line that names (S0 is not the state: Disabled" ((Get-ModernStandbyState -Text $reasonNamesS0) -eq 'Disabled') "got $(Get-ModernStandbyState -Text $reasonNamesS0)"
}

# =====================================================================
Invoke-Section 'Service recovery values' {
    $set = New-FailureBlob -Reset 300 -Pairs @(1, 30000, 1, 60000, 1, 120000)
    $d = ConvertFrom-FailureActionsBlob -Bytes $set
    Assert-Case "a recovery value decodes" ($null -ne $d -and $d.Reset -eq 300 -and @($d.Actions).Count -eq 3 -and
        @($d.Actions)[2].Type -eq 1 -and @($d.Actions)[2].Delay -eq 120000)
    Assert-Case "and is the recovery step 9 sets" (Test-RestartRecovery $d)
    Assert-Case "no value: null" ($null -eq (ConvertFrom-FailureActionsBlob -Bytes $null))
    Assert-Case "under 20 bytes: null" ($null -eq (ConvertFrom-FailureActionsBlob -Bytes (New-Object byte[] 19)))
    Assert-Case "a count over 64: null" ($null -eq (ConvertFrom-FailureActionsBlob -Bytes (New-FailureBlob -Reset 1 -Pairs @(1, 1) -CountOverride 65)))
    Assert-Case "fewer actions than the count: null" ($null -eq (ConvertFrom-FailureActionsBlob -Bytes (New-FailureBlob -Reset 300 -Pairs @(1, 30000, 1, 60000) -CountOverride 3)))
    $zero = ConvertFrom-FailureActionsBlob -Bytes (New-FailureBlob -Reset 0)
    Assert-Case "no actions decodes as none" ($null -ne $zero -and @($zero.Actions).Count -eq 0 -and -not (Test-RestartRecovery $zero))
    Assert-Case "another delay is not it" (-not (Test-RestartRecovery (ConvertFrom-FailureActionsBlob -Bytes (New-FailureBlob -Reset 300 -Pairs @(1, 30000, 1, 60000, 1, 60000)))))
    Assert-Case "a reboot action is not it" (-not (Test-RestartRecovery (ConvertFrom-FailureActionsBlob -Bytes (New-FailureBlob -Reset 300 -Pairs @(1, 30000, 1, 60000, 2, 120000)))))
    Assert-Case "another reset period is not it" (-not (Test-RestartRecovery (ConvertFrom-FailureActionsBlob -Bytes (New-FailureBlob -Reset 86400 -Pairs @(1, 30000, 1, 60000, 1, 120000)))))
    Assert-Case "two actions are not it" (-not (Test-RestartRecovery (ConvertFrom-FailureActionsBlob -Bytes (New-FailureBlob -Reset 300 -Pairs @(1, 30000, 1, 60000)))))
    Assert-Case "nothing is not it" (-not (Test-RestartRecovery $null))
}

# =====================================================================
Invoke-Section 'Get-CloudSyncFolderName' {
    $saved = @{}
    foreach ($ev in @('OneDrive', 'OneDriveConsumer', 'OneDriveCommercial', 'USERPROFILE')) {
        $saved[$ev] = [Environment]::GetEnvironmentVariable($ev, 'Process')
        [Environment]::SetEnvironmentVariable($ev, $null, 'Process')
    }
    try {
        [Environment]::SetEnvironmentVariable('USERPROFILE', 'C:\Users\Box', 'Process')
        $cases = @(
            @('C:\Users\Box\AppData\Local\Claude-Data', $null, 'a user named Box is not a Box folder'),
            @('C:\Users\Box', $null, 'the profile itself is not a Box folder'),
            @('C:\Users\Box\Xbox\Claude', $null, 'Xbox is not Box'),
            @('C:\Users\Box\Boxes\Claude', $null, 'Boxes is not Box'),
            @('D:\Dropboxes\Claude', $null, 'Dropboxes is not Dropbox'),
            @('C:\Users\Box\Box\Claude', 'Box', 'a folder named Box is'),
            @('C:\Users\Box\Box Sync\Claude', 'Box', 'Box Sync is Box'),
            @('C:\Users\Box\OneDrive - Contoso\Claude', 'OneDrive', 'OneDrive - Contoso is OneDrive'),
            @('C:\Users\Box\Dropbox (Personal)\Claude', 'Dropbox', 'Dropbox (Personal) is Dropbox'),
            @('G:\My Drive\Claude', 'Google Drive', 'My Drive is Google Drive'),
            @('C:\Users\Box\Google Drive\Claude', 'Google Drive', 'Google Drive is'),
            @('C:\Users\Box\iCloudDrive\Claude', 'iCloud', 'iCloudDrive is iCloud'),
            @('D:\Box', 'Box', 'outside the profile, a Box segment is Box'),
            @('', $null, 'no path: none')
        )
        foreach ($c in $cases) {
            $got = Get-CloudSyncFolderName -Path $c[0]
            Assert-Case $c[2] ($got -eq $c[1] -or ($null -eq $c[1] -and $null -eq $got)) "'$($c[0])' gave '$got'"
        }
        [Environment]::SetEnvironmentVariable('OneDrive', 'E:\Sync\Personal\', 'Process')
        Assert-Case "under the OneDrive folder Windows names" ((Get-CloudSyncFolderName -Path 'E:\Sync\Personal\Claude') -eq 'OneDrive')
        Assert-Case "beside it, with a name that starts the same, is not" ($null -eq (Get-CloudSyncFolderName -Path 'E:\Sync\PersonalX\Claude'))
        [Environment]::SetEnvironmentVariable('OneDriveCommercial', 'C:\Corp\Work', 'Process')
        Assert-Case "the work OneDrive too, in any case" ((Get-CloudSyncFolderName -Path 'c:\corp\work\Claude') -eq 'OneDrive')
    } finally {
        foreach ($ev in $saved.Keys) { [Environment]::SetEnvironmentVariable($ev, $saved[$ev], 'Process') }
    }
}

# =====================================================================
Invoke-Section 'Network adapter power saving: the measured Realtek and WiFi' {
    Reset-State
    Set-NicFixture
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Complete-Step
    $s = Get-LastStep
    $adv = @($script:NicCalls | Where-Object { $_ -like 'Set-Adv *' })
    $kws = @($adv | ForEach-Object { ($_ -split '\|')[1] -replace '=.*$', '' } | Sort-Object) -join ','
    Assert-Case "exactly the four power saving options are turned off" ($kws -eq '*EEE,EnableGreenEthernet,GigaLite,PowerSavingMode') "got $kws"
    Assert-Case "each to 0, without restarting the adapter" (@($adv | Where-Object { $_ -notmatch '=0 NoRestart=True$' }).Count -eq 0) ($adv -join ' / ')
    $pmc = @($script:NicCalls | Where-Object { $_ -like 'Set-Pm *' })
    Assert-Case "selective suspend is turned off, without a restart, and nothing else there" (($pmc -join ',') -eq 'Set-Pm Ethernet|SelectiveSuspend=Disabled NoRestart=True') ($pmc -join ',')
    $cim = @($script:NicCalls | Where-Object { $_ -like 'Set-CimInstance *' })
    Assert-Case "'allow the computer to turn off this device' is cleared on the Realtek only" (($cim -join ',') -eq "Set-CimInstance ${ethPnp}_0 Enable=False") ($cim -join ',')
    Assert-Case "an instance whose name only starts with the Realtek's id is left alone" ($script:DevEnable[1].Enable -eq $true)
    Assert-Case "PnPCapabilities 24 already has the 0x08 bit: left as it is" ((Get-FakeRegValue $ethKey 'PnPCapabilities') -eq 24 -and
        @($script:Calls | Where-Object { $_ -like 'Set-ItemProperty*' }).Count -eq 0)
    Assert-Case "six undo records: the checkbox, four options, selective suspend" ($script:UndoRecord.Count -eq 6) (Get-UndoKeys)
    Assert-Case "the checkbox is recorded as it was" ((Find-Undo 'NicDeviceEnable' $ethPnp).Original -eq $true)
    $eee = Find-Undo 'NicAdvanced' "$ethPnp|*EEE"
    Assert-Case "an option is recorded with its registry value" ($null -ne $eee -and (@($eee.Original) -join ',') -eq '1')
    Assert-Case "selective suspend is recorded as it was" ((Find-Undo 'NicPowerManagement' "$ethPnp|SelectiveSuspend").Original -eq 'Enabled')
    Assert-Case "the step reports Done, after a Windows restart" ($s.Status -eq 'Done' -and $s.Reboot) "got $($s.Status), reboot $($s.Reboot)"
    $notes = Get-StepNotes
    Assert-Case "the Realtek: six settings turned off" ($notes -match [regex]::Escape('Ethernet (Realtek Gaming 2.5GbE Family Controller): 6 power saving setting(s) turned off')) $notes
    Assert-Case "the WiFi: already off" ($notes -match [regex]::Escape('WiFi (TP-Link Wireless USB Adapter): power saving already off')) $notes
    Assert-Case "the undo record on disk is clean" (Test-UndoFileClean)

    $script:NicCalls = New-Object System.Collections.ArrayList
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Complete-Step
    Assert-Case "a second run changes nothing" ((Get-StepStatus) -eq 'Unchanged' -and $script:NicCalls.Count -eq 0 -and $script:UndoRecord.Count -eq 6) "got $(Get-StepStatus), $($script:NicCalls.Count) call(s)"

    # -Undo, from the record as it was saved.
    $script:UndoRecord = New-Object System.Collections.ArrayList
    Read-PreventUndo
    $items = @($script:UndoRecord.ToArray())
    [array]::Reverse($items)
    Step 1 5 'undo' -Name 'Recorded settings'
    $kept = @()
    foreach ($it in $items) {
        if (Restore-PreventUndoItem -Item $it -CreatedPlan $null) { Clear-PreventUndoItem -Kind $it.Kind -Key $it.Key } else { $kept += $it.Key }
    }
    $vals = @($script:Adv['Ethernet'] | Where-Object { @('*EEE', 'EnableGreenEthernet', 'GigaLite', 'PowerSavingMode') -contains $_.RegistryKeyword } |
        ForEach-Object { "$($_.RegistryKeyword)=$(@($_.RegistryValue) -join ',')" }) -join ' '
    Assert-Case "-Undo puts every option back" ($vals -eq '*EEE=1 EnableGreenEthernet=1 GigaLite=1 PowerSavingMode=1') $vals
    Assert-Case "and selective suspend, and the checkbox" ($script:Pm['Ethernet'].SelectiveSuspend -eq 'Enabled' -and $script:DevEnable[0].Enable -eq $true)
    Assert-Case "without restarting the adapter" (@($script:NicCalls | Where-Object { $_ -match 'NoRestart=False' }).Count -eq 0)
    Assert-Case "and the record is gone" ($kept.Count -eq 0 -and $script:UndoRecord.Count -eq 0 -and -not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $UndoFile)) "kept: $($kept -join ', ')"
}

# =====================================================================
Invoke-Section 'Network adapter power saving: other drivers' {
    Reset-State
    $aPnp = 'PCI\VEN_8086&DEV_15F3\SYNTHA'
    $bPnp = 'PCI\VEN_8086&DEV_15F3\SYNTHB'
    $aKey = "$classRoot\$netClass\0010"
    $bKey = "$classRoot\$netClass\0011"
    $script:Nics = @(
        [pscustomobject]@{ Name = 'A'; InterfaceDescription = 'Intel A'; PnPDeviceID = $aPnp },
        [pscustomobject]@{ Name = 'B'; InterfaceDescription = 'Intel B'; PnPDeviceID = $bPnp }
    )
    $script:DrvKeys = @{ $aPnp = "$netClass\0010"; $bPnp = "$netClass\0011" }
    Set-FakeRegValue $aKey 'DriverDesc' 'Intel A'
    Set-FakeRegValue $bKey 'PnPCapabilities' ([int]16)
    $script:Adv = @{
        'A' = @(
            (New-AdvProp 'A' 'AutoDisableGigabit'   'Auto Disable Gigabit' $null @('0', '1') '1'),
            (New-AdvProp 'A' 'ULPMode'              'Ultra Low Power Mode' '0'   @('0', '1')),
            (New-AdvProp 'A' 'EEELinkAdvertisement' 'Energy Efficient Ethernet' '1' @('0', '1')),
            (New-AdvProp 'A' 'SipsEnabled'          'System Idle Power Saver' '1' @('1', '2'))
        )
        'B' = @()
    }
    $script:Pm = @{
        'A' = [pscustomobject]@{ Name = 'A'; SelectiveSuspend = 'Disabled'; DeviceSleepOnDisconnect = 'Enabled'; D0PacketCoalescing = 'Enabled' }
        'B' = [pscustomobject]@{ Name = 'B'; WakeOnMagicPacket = 'Enabled' }
    }
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    $cap = Get-FakeRegValue $aKey 'PnPCapabilities'
    Assert-Case "no PnPCapabilities: 8 is written" ($cap -is [int] -and $cap -eq 8) "got $cap"
    Assert-Case "recorded as there was none" ($null -ne (Find-Undo 'Registry' "$aKey|PnPCapabilities") -and $null -eq (Find-Undo 'Registry' "$aKey|PnPCapabilities").Original)
    Assert-Case "PnPCapabilities 16: the 0x08 bit is added, the wake bit kept" ((Get-FakeRegValue $bKey 'PnPCapabilities') -eq 24 -and (Find-Undo 'Registry' "$bKey|PnPCapabilities").Original -eq 16)
    $adg = Find-Undo 'NicAdvanced' "$aPnp|AutoDisableGigabit"
    Assert-Case "an option at its driver default is turned off and recorded as unset" ($null -ne $adg -and $null -eq $adg.Original -and
        (@($script:Adv['A'][0].RegistryValue) -join ',') -eq '0')
    Assert-Case "an option already 0 is left alone" ($null -eq (Find-Undo 'NicAdvanced' "$aPnp|ULPMode"))
    Assert-Case "an option where 0 is not a valid value is left alone" ($null -eq (Find-Undo 'NicAdvanced' "$aPnp|SipsEnabled") -and
        (@($script:Adv['A'][3].RegistryValue) -join ',') -eq '1')
    Assert-Case "Intel's Energy Efficient Ethernet keyword is turned off" ((@($script:Adv['A'][2].RegistryValue) -join ',') -eq '0')
    Assert-Case "sleep on disconnect and packet coalescing are turned off" ($script:Pm['A'].DeviceSleepOnDisconnect -eq 'Disabled' -and $script:Pm['A'].D0PacketCoalescing -eq 'Disabled')
    Assert-Case "an adapter without those settings is fine" ((Get-StepStatus) -eq 'Done') "got $(Get-StepStatus): $(Get-StepNotes)"

    $script:UndoRecord = New-Object System.Collections.ArrayList
    Read-PreventUndo
    $adg = Find-Undo 'NicAdvanced' "$aPnp|AutoDisableGigabit"
    Assert-Case "unset survives the record on disk" ($null -ne $adg -and $null -eq $adg.Original)
    $items = @($script:UndoRecord.ToArray())
    [array]::Reverse($items)
    Step 1 5 'undo' -Name 'Recorded settings'
    foreach ($it in $items) { if (Restore-PreventUndoItem -Item $it -CreatedPlan $null) { Clear-PreventUndoItem -Kind $it.Kind -Key $it.Key } }
    Assert-Case "-Undo removes the PnPCapabilities it added" (-not (Test-FakeRegValue $aKey 'PnPCapabilities'))
    Assert-Case "and puts back the one it changed" ((Get-FakeRegValue $bKey 'PnPCapabilities') -eq 16)
    Assert-Case "and resets the option that was at its default" ($null -eq $script:Adv['A'][0].RegistryValue -and
        @($script:NicCalls | Where-Object { $_ -eq 'Reset-Adv A|AutoDisableGigabit NoRestart=True' }).Count -eq 1)
    Assert-Case "and the record is gone" ($script:UndoRecord.Count -eq 0)
}

# =====================================================================
Invoke-Section 'Network adapter power saving: what goes wrong' {
    Reset-State
    $script:NicThrows = $true
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Assert-Case "the adapters cannot be listed: Failed, saying so for each list" ((Get-StepStatus) -eq 'Failed' -and (Get-StepNotes) -eq (
        'Could not list the physical network adapters, so their power saving was not changed: The WMI provider is not loaded | ' +
        'Could not list the hidden and virtual network adapters: The WMI provider is not loaded')) (Get-StepNotes)

    Reset-State
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Assert-Case "no adapters: Skipped" ((Get-StepStatus) -eq 'Skipped')

    Reset-State
    Set-NicFixture
    $script:DevEnableThrows = $true
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    $s = Get-LastStep
    Assert-Case "the checkbox cannot be read: a warning, the rest still done" ($s.Status -eq 'Warning' -and $s.Reboot -and
        (Get-StepNotes) -match 'could not read whether Windows may turn it off: Invalid namespace' -and
        @($script:NicCalls | Where-Object { $_ -like 'Set-Adv *' }).Count -eq 4) (Get-StepNotes)

    Reset-State
    Set-NicFixture
    $script:AdvFault['GigaLite'] = 'throw'
    $script:AdvFault['PowerSavingMode'] = 'stuck'
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    $notes = Get-StepNotes
    Assert-Case "an option that throws and one that does not take: Failed, naming both" ((Get-StepStatus) -eq 'Failed' -and
        $notes -match 'Gigabit Lite: The parameter is incorrect\.' -and $notes -match 'Power Saving Mode still reads 1') $notes
    Assert-Case "and the other options are still turned off" ((@($script:Adv['Ethernet'][0].RegistryValue) -join ',') -eq '0' -and
        (@($script:Adv['Ethernet'][5].RegistryValue) -join ',') -eq '0')

    Reset-State
    Set-NicFixture
    $script:DrvKeys = @{}
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Assert-Case "no driver key known: a warning" ((Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -match 'could not find its driver key, so PnPCapabilities was not set') (Get-StepNotes)

    Reset-State
    Set-NicFixture
    $script:Reg.Remove((Get-RegNorm $ethKey))
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Assert-Case "its driver key is missing: a warning" ((Get-StepNotes) -match [regex]::Escape("its driver key $netClass\0002 is missing, so PnPCapabilities was not set")) (Get-StepNotes)
    Assert-Case "and the key is never created" (-not (Test-Path -LiteralPath $ethKey) -and $null -eq (Find-Undo 'RegistryKey' $ethKey))

    Reset-State
    Set-NicFixture
    $script:RegFault[(Get-RegNorm $wifiKey) + '|pnpcapabilities'] = 'throw'
    Set-FakeRegValue $wifiKey 'PnPCapabilities' ([int]16)
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Assert-Case "PnPCapabilities cannot be written: Failed" ((Get-StepStatus) -eq 'Failed' -and (Get-StepNotes) -match 'could not write PnPCapabilities to') (Get-StepNotes)

    Reset-State
    Set-NicFixture
    $script:PmReadThrows = @('Ethernet')
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Assert-Case "power management cannot be read: a warning" ((Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -match 'could not read its power management settings') (Get-StepNotes)

    # Each adapter is read on its own: one whose driver cannot be read does
    # not keep the others from being turned off, or put back by -Undo.
    Reset-State
    Set-NicFixture
    $script:AdvReadThrows = @('WiFi')
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    $notes = Get-StepNotes
    $adv = @($script:NicCalls | Where-Object { $_ -like 'Set-Adv Ethernet|*' })
    Assert-Case "one adapter's driver options cannot be read: it fails, and the other is still turned off" ((Get-StepStatus) -eq 'Failed' -and
        $notes -match [regex]::Escape("WiFi (TP-Link Wireless USB Adapter): could not read the driver's options: No MSFT_NetAdapterAdvancedPropertySettingData objects found") -and
        $notes -match [regex]::Escape('Ethernet (Realtek Gaming 2.5GbE Family Controller): 6 power saving setting(s) turned off') -and $adv.Count -eq 4) "$notes | $($adv -join ' / ')"

    Reset-State
    Set-NicFixture
    $script:PmReadThrows = @('WiFi')
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    $notes = Get-StepNotes
    $pmc = @($script:NicCalls | Where-Object { $_ -like 'Set-Pm *' })
    Assert-Case "one adapter's power management cannot be read: a warning for it, and the other's selective suspend is still turned off" (
        $notes -match [regex]::Escape('WiFi (TP-Link Wireless USB Adapter): could not read its power management settings: No MSFT_NetAdapterPowerManagementSettingData objects found') -and
        $notes -match [regex]::Escape('Ethernet (Realtek Gaming 2.5GbE Family Controller): 6 power saving setting(s) turned off') -and
        ($pmc -join ',') -eq 'Set-Pm Ethernet|SelectiveSuspend=Disabled NoRestart=True') "$notes | $($pmc -join ' / ')"

    Reset-State
    Set-NicFixture
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Complete-Step
    $script:AdvReadThrows = @('WiFi')
    $script:PmReadThrows = @('WiFi')
    $items = @($script:UndoRecord.ToArray())
    [array]::Reverse($items)
    Step 1 5 'undo' -Name 'Recorded settings'
    $kept = @()
    foreach ($it in $items) {
        # As -Undo step 1 does: an item whose restore throws is kept.
        try {
            if (Restore-PreventUndoItem -Item $it -CreatedPlan $null) { Clear-PreventUndoItem -Kind $it.Kind -Key $it.Key } else { $kept += $it.Key }
        } catch {
            $kept += "$($it.Key) ($($_.Exception.Message))"
        }
    }
    $vals = @($script:Adv['Ethernet'] | Where-Object { @('*EEE', 'EnableGreenEthernet', 'GigaLite', 'PowerSavingMode') -contains $_.RegistryKeyword } |
        ForEach-Object { "$($_.RegistryKeyword)=$(@($_.RegistryValue) -join ',')" }) -join ' '
    Assert-Case "-Undo: an adapter that cannot be read does not keep the other's settings from being put back" (
        $vals -eq '*EEE=1 EnableGreenEthernet=1 GigaLite=1 PowerSavingMode=1' -and $script:Pm['Ethernet'].SelectiveSuspend -eq 'Enabled' -and
        $kept.Count -eq 0) "$vals; selective suspend $($script:Pm['Ethernet'].SelectiveSuspend); kept: $($kept -join ', ')"

    # Each wildcard character is asked for as ?, so "Ethernet [LAN]" reads
    # "Ethernet (LAN)" as well: only the exact name is kept. The other one
    # here is not a physical adapter, so nothing of it may change.
    Reset-State
    Set-NicFixture
    Rename-FakeNic 'Ethernet' 'Ethernet [LAN]'
    $script:Adv['Ethernet (LAN)'] = @(New-AdvProp 'Ethernet (LAN)' '*EEE' 'Energy-Efficient Ethernet' '1' @('0', '1'))
    $script:Pm['Ethernet (LAN)'] = [pscustomobject]@{
        Name = 'Ethernet (LAN)'; AllowComputerToTurnOffDevice = 'Enabled'; D0PacketCoalescing = 'Enabled'
        DeviceSleepOnDisconnect = 'Enabled'; SelectiveSuspend = 'Enabled'; WakeOnMagicPacket = 'Disabled'
    }
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    $notes = Get-StepNotes
    $other = @($script:NicCalls | Where-Object { $_ -like '*Ethernet (LAN)*' })
    Assert-Case "an adapter whose name only matches the pattern is left alone" ($other.Count -eq 0 -and
        (@($script:Adv['Ethernet (LAN)'][0].RegistryValue) -join ',') -eq '1' -and $script:Pm['Ethernet (LAN)'].SelectiveSuspend -eq 'Enabled' -and
        $notes -match [regex]::Escape('Ethernet [LAN] (Realtek Gaming 2.5GbE Family Controller): 6 power saving setting(s) turned off')) "$notes | $($other -join ' / ')"

    # A physical adapter that Windows lists no driver options and no power
    # settings for: the read by its plain name is an ObjectNotFound error
    # from the cmdlet, which means there is nothing to turn off, not that the
    # read failed.
    Reset-State
    Set-NicFixture
    $null = $script:Adv.Remove('WiFi')
    $null = $script:Pm.Remove('WiFi')
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    $notes = Get-StepNotes
    Assert-Case "an adapter Windows lists no options or power settings for: a warning that says so, not a failure" ((Get-StepStatus) -eq 'Warning' -and
        $notes -match [regex]::Escape('WiFi (TP-Link Wireless USB Adapter): could not read its power management settings: Windows lists none for it') -and
        $notes -notmatch 'No MSFT_' -and $notes -notmatch "could not read the driver's options" -and
        $notes -match [regex]::Escape('Ethernet (Realtek Gaming 2.5GbE Family Controller): 6 power saving setting(s) turned off')) $notes

    Reset-State
    Set-NicFixture
    $null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $UndoFile -Force
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Assert-Case "the undo record cannot be saved: Failed" ((Get-StepStatus) -eq 'Failed' -and (Get-StepNotes) -match 'Could not save the undo record') (Get-StepNotes)
    Assert-Case "and nothing on the adapter was changed" ($script:NicCalls.Count -eq 0 -and $script:DevEnable[0].Enable -eq $true -and
        (@($script:Adv['Ethernet'][0].RegistryValue) -join ',') -eq '1') ($script:NicCalls -join ' / ')
    Microsoft.PowerShell.Management\Remove-Item -LiteralPath $UndoFile -Recurse -Force
}

# =====================================================================
# What the first fixture leaves out, as measured on this PC: the
# Realtek's WDF key, where IdleInWorkingState lets the driver framework
# power it down while the PC is running; a hidden option on each adapter,
# PowerDownPll and bFwCtrlLPS, both 1, and bLeisurePs, hidden and 0; and
# the hidden adapters Get-NetAdapter lists only with -IncludeHidden. Of
# those, the two Wi-Fi Direct adapters have the checkbox on, the kernel
# debugger adapter has no instance of it, and Teredo has no device id. The
# last instance is a decoy: its name only starts with the second Wi-Fi
# Direct adapter's id. Add-TapFixture adds a visible virtual adapter, a
# VPN's TAP adapter, which -IncludeHidden lists beside the hidden ones.
$wdfEth  = "HKLM:\SYSTEM\CurrentControlSet\Enum\$ethPnp\Device Parameters\WDF"
$wdfWifi = "HKLM:\SYSTEM\CurrentControlSet\Enum\$wifiPnp\Device Parameters\WDF"
$wfd1    = '{5d624f94-8850-40c3-a3fa-a4fd2080baf3}\vwifimp_wfd\b&6a0164d&0&11'
$wfd2    = '{5d624f94-8850-40c3-a3fa-a4fd2080baf3}\vwifimp_wfd\b&6a0164d&0&12'
$tapPnp  = 'ROOT\NET\0000'
function Set-NicMoreFixture {
    Set-NicFixture
    Set-FakeRegValue $wdfEth 'WdfDefaultIdleInWorkingState' ([int]1)
    Set-FakeRegValue $wdfEth 'IdleInWorkingState' ([int]1)
    Set-FakeRegValue $wdfEth 'WakeFromSleepState' ([int]0)
    $script:Adv['Ethernet'] += (New-AdvProp 'Ethernet' 'PowerDownPll' '' '1' $null)
    $script:Adv['WiFi'] += (New-AdvProp 'WiFi' 'bFwCtrlLPS' '' '1' $null)
    $script:Adv['WiFi'] += (New-AdvProp 'WiFi' 'bLeisurePs' '' '0' $null)
    $script:HiddenNics = @(
        [pscustomobject]@{ Name = 'Ethernet (Kernel Debugger)'; InterfaceDescription = 'Microsoft Kernel Debug Network Adapter'; PnPDeviceID = 'ROOT\KDNIC\0000'; Hidden = $true; ConnectorPresent = $false },
        [pscustomobject]@{ Name = 'Local Area Connection* 2'; InterfaceDescription = 'Microsoft Wi-Fi Direct Virtual Adapter #2'; PnPDeviceID = $wfd2; Hidden = $true; ConnectorPresent = $false },
        [pscustomobject]@{ Name = 'Teredo Tunneling Pseudo-Interface'; InterfaceDescription = ''; PnPDeviceID = ''; Hidden = $true; ConnectorPresent = $false },
        [pscustomobject]@{ Name = 'Local Area Connection* 1'; InterfaceDescription = 'Microsoft Wi-Fi Direct Virtual Adapter'; PnPDeviceID = $wfd1; Hidden = $true; ConnectorPresent = $false }
    )
    $script:DevEnable += [pscustomobject]@{ InstanceName = "${wfd1}_0"; Enable = $true }
    $script:DevEnable += [pscustomobject]@{ InstanceName = "${wfd2}_0"; Enable = $true }
    $script:DevEnable += [pscustomobject]@{ InstanceName = "${wfd2}9_0"; Enable = $true }
}
function Add-TapFixture {
    $script:HiddenNics += [pscustomobject]@{ Name = 'OpenVPN TAP'; InterfaceDescription = 'TAP-Windows Adapter V9'; PnPDeviceID = $tapPnp; Hidden = $false; ConnectorPresent = $false }
    $script:DevEnable += [pscustomobject]@{ InstanceName = "${tapPnp}_0"; Enable = $true }
}
function Get-FakeAdv {
    param([string]$Adapter, [string]$Keyword)
    foreach ($p in @($script:Adv[$Adapter])) { if ($p.RegistryKeyword -eq $Keyword) { return (@($p.RegistryValue) -join ',') } }
    return '(none)'
}
function Get-FakeDevEnable {
    param([string]$InstanceName)
    foreach ($d in @($script:DevEnable)) { if ($d.InstanceName -eq $InstanceName) { return $d.Enable } }
    return $null
}

Invoke-Section 'Network adapter power saving: the framework idle power-down, hidden options, hidden and virtual adapters' {
    Reset-State
    Set-NicMoreFixture
    Add-TapFixture
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Complete-Step
    $notes = Get-StepNotes

    # The framework's idle power-down
    $wdfRec = Find-Undo 'Registry' "$wdfEth|IdleInWorkingState"
    Assert-Case "the Realtek's IdleInWorkingState is set to 0" ((Get-FakeRegValue $wdfEth 'IdleInWorkingState') -eq 0) "got $(Get-FakeRegValue $wdfEth 'IdleInWorkingState')"
    Assert-Case "recorded with the 1 it held" ($null -ne $wdfRec -and $wdfRec.Original -eq 1) (Get-UndoKeys)
    $wdfSets = @($script:Calls | Where-Object { $_ -like 'Set-ItemProperty*WDF*' })
    Assert-Case "the default and the wake value beside it are left alone" ((Get-FakeRegValue $wdfEth 'WdfDefaultIdleInWorkingState') -eq 1 -and
        (Get-FakeRegValue $wdfEth 'WakeFromSleepState') -eq 0 -and $wdfSets.Count -eq 1) ($wdfSets -join ' / ')
    Assert-Case "and it is in the Realtek's list" ($script:LogLines -contains '+ the driver framework may no longer power it down while the PC is running (IdleInWorkingState 0)') ($script:LogLines -join ' / ')
    Assert-Case "the WiFi has no WDF key: none is created" (-not (Test-Path -LiteralPath $wdfWifi) -and
        @($script:Calls | Where-Object { $_ -like 'New-Item*' }).Count -eq 0 -and $null -eq (Find-Undo 'RegistryKey' $wdfWifi))

    # The hidden options
    $adv = @($script:NicCalls | Where-Object { $_ -like 'Set-Adv *' })
    Assert-Case "PowerDownPll and bFwCtrlLPS, hidden, are set to 0 without restarting the adapter" (
        $adv -contains 'Set-Adv Ethernet|PowerDownPll=0 NoRestart=True' -and $adv -contains 'Set-Adv WiFi|bFwCtrlLPS=0 NoRestart=True') ($adv -join ' / ')
    Assert-Case "and read back as 0" ((Get-FakeAdv 'Ethernet' 'PowerDownPll') -eq '0' -and (Get-FakeAdv 'WiFi' 'bFwCtrlLPS') -eq '0')
    $pll = Find-Undo 'NicAdvanced' "$ethPnp|PowerDownPll"
    $lps = Find-Undo 'NicAdvanced' "$wifiPnp|bFwCtrlLPS"
    Assert-Case "each is recorded with the 1 it held" ($null -ne $pll -and (@($pll.Original) -join ',') -eq '1' -and
        $null -ne $lps -and (@($lps.Original) -join ',') -eq '1') (Get-UndoKeys)
    Assert-Case "the Realtek: eight settings turned off, the WiFi: one" (
        $notes -match [regex]::Escape('Ethernet (Realtek Gaming 2.5GbE Family Controller): 8 power saving setting(s) turned off') -and
        $notes -match [regex]::Escape('WiFi (TP-Link Wireless USB Adapter): 1 power saving setting(s) turned off')) $notes

    # The hidden adapters, and the TAP adapter, virtual but not hidden
    $cim = @($script:NicCalls | Where-Object { $_ -like 'Set-CimInstance *' })
    Assert-Case "the checkbox is cleared on the Realtek, the two Wi-Fi Direct adapters and the TAP adapter, nowhere else" (
        ($cim -join ',') -eq ("Set-CimInstance ${ethPnp}_0 Enable=False,Set-CimInstance ${wfd2}_0 Enable=False,Set-CimInstance ${wfd1}_0 Enable=False," +
        "Set-CimInstance ${tapPnp}_0 Enable=False")) ($cim -join ',')
    Assert-Case "an instance whose name only starts with a hidden adapter's id is left alone" ((Get-FakeDevEnable "${wfd2}9_0") -eq $true)
    $r1 = Find-Undo 'NicDeviceEnable' $wfd1
    $r2 = Find-Undo 'NicDeviceEnable' $wfd2
    $r3 = Find-Undo 'NicDeviceEnable' $tapPnp
    Assert-Case "each adapter that is not physical is recorded under its device id, the TAP adapter too" ($null -ne $r1 -and $r1.Original -eq $true -and
        $null -ne $r2 -and $r2.Original -eq $true -and $null -ne $r3 -and $r3.Original -eq $true) (Get-UndoKeys)
    $hl = @($script:LogLines | Where-Object { $_ -match 'hidden adapter|virtual adapter' })
    $want = @('+ Local Area Connection* 2 (Microsoft Wi-Fi Direct Virtual Adapter #2), a hidden adapter: Windows may no longer turn it off to save power',
              '+ Local Area Connection* 1 (Microsoft Wi-Fi Direct Virtual Adapter), a hidden adapter: Windows may no longer turn it off to save power',
              '+ OpenVPN TAP (TAP-Windows Adapter V9), a virtual adapter: Windows may no longer turn it off to save power')
    Assert-Case "one line for each, a hidden one called hidden and the TAP adapter virtual, none for one with nothing to change" (($hl -join ' / ') -eq ($want -join ' / ') -and
        -not ($script:LogLines -match 'Kernel Debug|Teredo')) ($script:LogLines -join ' / ')
    Assert-Case "and the same in the step's summary" ($notes -match [regex]::Escape('Local Area Connection* 1 (Microsoft Wi-Fi Direct Virtual Adapter), a hidden adapter: Windows may no longer') -and
        $notes -match [regex]::Escape('OpenVPN TAP (TAP-Windows Adapter V9), a virtual adapter: Windows may no longer turn it off to save power')) $notes
    Assert-Case "twelve undo records: the six before, the WDF value, two options and three checkboxes" ($script:UndoRecord.Count -eq 12) (Get-UndoKeys)

    $script:NicCalls = New-Object System.Collections.ArrayList
    $setsBefore = @($script:Calls | Where-Object { $_ -like 'Set-ItemProperty*' }).Count
    $linesBefore = $script:LogLines.Count
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Complete-Step
    $again = @($script:LogLines | Select-Object -Skip $linesBefore)
    Assert-Case "a second run changes nothing, and says nothing of the hidden or virtual adapters" ((Get-StepStatus) -eq 'Unchanged' -and $script:NicCalls.Count -eq 0 -and
        @($script:Calls | Where-Object { $_ -like 'Set-ItemProperty*' }).Count -eq $setsBefore -and $script:UndoRecord.Count -eq 12 -and
        -not ($again -match 'hidden|virtual')) "got $(Get-StepStatus), $($script:NicCalls.Count) call(s); $($again -join ' / ')"

    # -Undo, from the record as it was saved. The adapters that are not
    # physical are not listed any more: -Undo finds their checkboxes by
    # instance name.
    $script:HiddenNics = @()
    $script:UndoRecord = New-Object System.Collections.ArrayList
    Read-PreventUndo
    $items = @($script:UndoRecord.ToArray())
    [array]::Reverse($items)
    Step 1 5 'undo' -Name 'Recorded settings'
    $kept = @()
    foreach ($it in $items) {
        # As -Undo step 1 does: an item whose restore throws is kept.
        try {
            if (Restore-PreventUndoItem -Item $it -CreatedPlan $null) { Clear-PreventUndoItem -Kind $it.Kind -Key $it.Key } else { $kept += $it.Key }
        } catch {
            $kept += "$($it.Key) ($($_.Exception.Message))"
        }
    }
    Assert-Case "-Undo puts IdleInWorkingState back to 1" ((Get-FakeRegValue $wdfEth 'IdleInWorkingState') -eq 1) "got $(Get-FakeRegValue $wdfEth 'IdleInWorkingState')"
    Assert-Case "and PowerDownPll and bFwCtrlLPS back to 1" ((Get-FakeAdv 'Ethernet' 'PowerDownPll') -eq '1' -and (Get-FakeAdv 'WiFi' 'bFwCtrlLPS') -eq '1')
    Assert-Case "and the hidden adapters' checkboxes, with neither one listed as an adapter" ((Get-FakeDevEnable "${wfd1}_0") -eq $true -and
        (Get-FakeDevEnable "${wfd2}_0") -eq $true) "kept: $($kept -join ', ')"
    Assert-Case "and the TAP adapter's, which is virtual but not hidden" ((Get-FakeDevEnable "${tapPnp}_0") -eq $true) "kept: $($kept -join ', ')"
    Assert-Case "without restarting an adapter" (@($script:NicCalls | Where-Object { $_ -match 'NoRestart=False' }).Count -eq 0)
    Assert-Case "and the record is gone" ($kept.Count -eq 0 -and $script:UndoRecord.Count -eq 0 -and -not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $UndoFile)) "kept: $($kept -join ', ')"

    # The hidden options are there only with -AllProperties, as in Windows.
    Reset-State
    Set-NicMoreFixture
    $advPlain = @(Get-NetAdapterAdvancedProperty -Name 'Ethernet' | Where-Object { $_.RegistryKeyword -eq 'PowerDownPll' })
    $advAll = @(Get-NetAdapterAdvancedProperty -Name 'Ethernet' -AllProperties | Where-Object { $_.RegistryKeyword -eq 'PowerDownPll' })
    Assert-Case "the fake lists PowerDownPll only with -AllProperties" ($advPlain.Count -eq 0 -and $advAll.Count -eq 1) "without: $($advPlain.Count), with: $($advAll.Count)"

    # -Undo puts a hidden option back on its own, with no setup run before
    # it to have read it.
    foreach ($o in @($script:Adv['Ethernet'])) { if ($o.RegistryKeyword -eq 'PowerDownPll') { $o.RegistryValue = [string[]]@('0') } }
    Step 1 5 'undo' -Name 'Recorded settings'
    $pllItem = [pscustomobject]@{ Kind = 'NicAdvanced'; Key = "$ethPnp|PowerDownPll"; Original = @('1') }
    $ok = Restore-PreventUndoItem -Item $pllItem -CreatedPlan $null
    Assert-Case "-Undo puts a hidden option back on its own, without a restart" ($ok -eq $true -and (Get-FakeAdv 'Ethernet' 'PowerDownPll') -eq '1' -and
        ($script:NicCalls -join ',') -eq 'Set-Adv Ethernet|PowerDownPll=1 NoRestart=True') "$ok; $($script:NicCalls -join ',')"
}

# =====================================================================
Invoke-Section 'Network adapter power saving: the framework idle power-down, hidden options, hidden and virtual adapters, what goes wrong' {
    # The framework's idle power-down
    Reset-State
    Set-NicMoreFixture
    $script:RegFault[(Get-RegNorm $wdfEth) + '|idleinworkingstate'] = 'throw'
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Assert-Case "IdleInWorkingState cannot be written: Failed, naming the key" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -match [regex]::Escape("could not write IdleInWorkingState to $wdfEth")) (Get-StepNotes)
    Assert-Case "and the rest is still turned off" ((Get-FakeAdv 'Ethernet' 'PowerDownPll') -eq '0' -and (Get-FakeAdv 'Ethernet' '*EEE') -eq '0')

    Reset-State
    Set-NicMoreFixture
    Set-FakeRegValue $wdfEth 'IdleInWorkingState' ([int]0)
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Assert-Case "IdleInWorkingState already 0: nothing written, nothing recorded" ((Get-FakeRegValue $wdfEth 'IdleInWorkingState') -eq 0 -and
        @($script:Calls | Where-Object { $_ -like 'Set-ItemProperty*WDF*' }).Count -eq 0 -and $null -eq (Find-Undo 'Registry' "$wdfEth|IdleInWorkingState") -and
        -not ($script:LogLines -match 'IdleInWorkingState')) ($script:Calls -join ' / ')

    Reset-State
    Set-NicMoreFixture
    $script:Reg[(Get-RegNorm $wdfEth)].Vals.Remove('IdleInWorkingState')
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Assert-Case "a WDF key without IdleInWorkingState: none is written, nothing recorded" (-not (Test-FakeRegValue $wdfEth 'IdleInWorkingState') -and
        @($script:Calls | Where-Object { $_ -like 'Set-ItemProperty*WDF*' }).Count -eq 0 -and $null -eq (Find-Undo 'Registry' "$wdfEth|IdleInWorkingState")) ($script:Calls -join ' / ')

    Reset-State
    Set-NicMoreFixture
    $script:Reg.Remove((Get-RegNorm $wdfEth))
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Assert-Case "no WDF key: none is created, nothing recorded" (-not (Test-Path -LiteralPath $wdfEth) -and
        @($script:Calls | Where-Object { $_ -like 'New-Item*' -or $_ -like 'Set-ItemProperty*WDF*' }).Count -eq 0 -and
        $null -eq (Find-Undo 'RegistryKey' $wdfEth) -and $null -eq (Find-Undo 'Registry' "$wdfEth|IdleInWorkingState")) ($script:Calls -join ' / ')

    # The hidden options
    Reset-State
    Set-NicMoreFixture
    foreach ($p in @($script:Adv['Ethernet'])) { if ($p.RegistryKeyword -eq 'PowerDownPll') { $p.RegistryValue = [string[]]@('2') } }
    foreach ($p in @($script:Adv['WiFi'])) { if ($p.RegistryKeyword -eq 'bFwCtrlLPS') { $p.RegistryValue = [string[]]@('0') } }
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    $adv = @($script:NicCalls | Where-Object { $_ -like 'Set-Adv *' })
    Assert-Case "a hidden option that reads 2, not its on value 1, is left alone, and one already 0 too" (
        @($adv | Where-Object { $_ -match 'PowerDownPll|bFwCtrlLPS' }).Count -eq 0 -and (Get-FakeAdv 'Ethernet' 'PowerDownPll') -eq '2' -and
        $null -eq (Find-Undo 'NicAdvanced' "$ethPnp|PowerDownPll") -and $null -eq (Find-Undo 'NicAdvanced' "$wifiPnp|bFwCtrlLPS")) ($adv -join ' / ')

    # A hidden option with no value, and one with two: neither is its on
    # value alone.
    Reset-State
    Set-NicMoreFixture
    foreach ($p in @($script:Adv['Ethernet'])) { if ($p.RegistryKeyword -eq 'PowerDownPll') { $p.RegistryValue = $null } }
    foreach ($p in @($script:Adv['WiFi'])) { if ($p.RegistryKeyword -eq 'bFwCtrlLPS') { $p.RegistryValue = [string[]]@('1', '2') } }
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Assert-Case "a hidden option with no value, and one with two, are left alone" (
        @($script:NicCalls | Where-Object { $_ -match 'PowerDownPll|bFwCtrlLPS' }).Count -eq 0 -and
        (Get-FakeAdv 'WiFi' 'bFwCtrlLPS') -eq '1,2' -and (Get-StepStatus) -eq 'Done') "$(Get-StepStatus): $($script:NicCalls -join ' / ')"

    Reset-State
    Set-NicMoreFixture
    $script:AdvFault['bFwCtrlLPS'] = 'throw'
    $script:AdvFault['PowerDownPll'] = 'stuck'
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    $notes = Get-StepNotes
    Assert-Case "a hidden option that throws and one that does not take: Failed, naming both by keyword" ((Get-StepStatus) -eq 'Failed' -and
        $notes -match 'bFwCtrlLPS: The parameter is incorrect\.' -and $notes -match 'PowerDownPll still reads 1') $notes

    # A driver that lists values for PowerDownPll says what they mean, and
    # 0 is not among them here: the hidden table does not apply.
    Reset-State
    Set-NicMoreFixture
    $script:Adv['Ethernet'] = @(@($script:Adv['Ethernet'] | Where-Object { $_.RegistryKeyword -ne 'PowerDownPll' }) +
        (New-AdvProp 'Ethernet' 'PowerDownPll' 'Power Down PLL' '1' @('1', '2')))
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Assert-Case "PowerDownPll with values 1 and 2 listed, and no 0, is left alone" ((Get-FakeAdv 'Ethernet' 'PowerDownPll') -eq '1' -and
        @($script:NicCalls | Where-Object { $_ -like 'Set-Adv *PowerDownPll*' }).Count -eq 0) ($script:NicCalls -join ' / ')

    # The hidden adapters
    Reset-State
    Set-NicMoreFixture
    $script:DevEnable[0].Enable = $false
    $script:SetCimThrows = $true
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Assert-Case "a hidden adapter's checkbox cannot be cleared: Failed, naming the adapter" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -match [regex]::Escape("Local Area Connection* 1 (Microsoft Wi-Fi Direct Virtual Adapter), a hidden adapter: 'allow the computer to turn off this device': Access denied")) (Get-StepNotes)

    Reset-State
    Set-NicMoreFixture
    $script:DevEnable[0].Enable = $false
    foreach ($d in $script:DevEnable) { if ($d.InstanceName -like '*vwifimp*') { $d.Enable = $false } }
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Assert-Case "hidden adapters already off: no call, no record, no line" (@($script:NicCalls | Where-Object { $_ -like 'Set-CimInstance *' }).Count -eq 0 -and
        $null -eq (Find-Undo 'NicDeviceEnable' $wfd1) -and $null -eq (Find-Undo 'NicDeviceEnable' $wfd2) -and
        -not ($script:LogLines -match 'hidden') -and -not ((Get-StepNotes) -match 'hidden')) ($script:LogLines -join ' / ')

    # Windows' own read-back can lag: the physical adapter's checkbox still
    # reads on, and it is not taken again as a hidden one.
    Reset-State
    Set-NicMoreFixture
    $script:SetCimNoop = $true
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    $ethSets = @($script:NicCalls | Where-Object { $_ -eq "Set-CimInstance ${ethPnp}_0 Enable=False" })
    Assert-Case "a physical adapter whose checkbox still reads on is not taken again as a hidden one" ($ethSets.Count -eq 1 -and
        -not ($script:LogLines -match 'Realtek.*hidden adapter')) ($script:NicCalls -join ' / ')
    Assert-Case "a hidden adapter whose checkbox still reads on: a warning that says so" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -match [regex]::Escape("Local Area Connection* 2 (Microsoft Wi-Fi Direct Virtual Adapter #2), a hidden adapter: 'allow the computer to turn off this device' still reads on until Windows restarts")) (Get-StepNotes)

    # The same, with no ConnectorPresent on the physical adapters: the
    # -Physical list's device ids alone leave them out of the hidden ones.
    Reset-State
    Set-NicMoreFixture
    $script:Nics = @($script:Nics | ForEach-Object { [pscustomobject]@{ Name = $_.Name; InterfaceDescription = $_.InterfaceDescription; PnPDeviceID = $_.PnPDeviceID } })
    $script:SetCimNoop = $true
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    $ethSets = @($script:NicCalls | Where-Object { $_ -eq "Set-CimInstance ${ethPnp}_0 Enable=False" })
    Assert-Case "a physical adapter with no ConnectorPresent is known by its device id, and not taken again as a hidden one" ($ethSets.Count -eq 1 -and
        -not ($script:LogLines -match 'Realtek.*hidden adapter')) ($script:NicCalls -join ' / ')

    # The same adapter listed twice, with the read-back lagging so that its
    # checkbox still reads on the second time; and Teredo, with no device
    # id, beside an instance whose name an empty id would match.
    Reset-State
    Set-NicMoreFixture
    $script:SetCimNoop = $true
    $script:HiddenNics += [pscustomobject]@{ Name = 'Local Area Connection* 1'; InterfaceDescription = 'Microsoft Wi-Fi Direct Virtual Adapter'; PnPDeviceID = $wfd1; Hidden = $true }
    $script:DevEnable += [pscustomobject]@{ InstanceName = '_0'; Enable = $true }
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    $w1 = @($script:NicCalls | Where-Object { $_ -eq "Set-CimInstance ${wfd1}_0 Enable=False" })
    $l1 = @($script:LogLines | Where-Object { ([string]$_).Contains('Local Area Connection* 1 (') })
    Assert-Case "a hidden adapter listed twice is taken once, with one line" ($w1.Count -eq 1 -and $l1.Count -eq 1) ($script:NicCalls -join ' / ')
    Assert-Case "an adapter without a device id takes no instance and leaves no record" (
        @($script:NicCalls | Where-Object { $_ -like 'Set-CimInstance _0 *' }).Count -eq 0 -and $null -eq (Find-Undo 'NicDeviceEnable' '')) (Get-UndoKeys)

    Reset-State
    Set-NicMoreFixture
    $script:HiddenNicThrows = $true
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Assert-Case "the hidden adapters cannot be listed: a warning, and the physical ones are still done" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -match 'Could not list the hidden and virtual network adapters: Generic failure' -and
        (Get-FakeDevEnable "${ethPnp}_0") -eq $false -and (Get-FakeAdv 'Ethernet' '*EEE') -eq '0') (Get-StepNotes)

    Reset-State
    Set-NicMoreFixture
    $script:DevEnableThrows = $true
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    $hl = @($script:LogLines | Where-Object { $_ -match 'Hidden and virtual network adapters' })
    Assert-Case "the checkboxes cannot be read: one line for the hidden adapters, saying so" ($hl.Count -eq 1 -and
        $hl[0] -eq '! Hidden and virtual network adapters: could not read whether Windows may turn them off: Invalid namespace' -and (Get-StepStatus) -eq 'Warning') ($script:LogLines -join ' / ')

    Reset-State
    Set-NicFixture
    $script:DevEnableThrows = $true
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Assert-Case "the checkboxes cannot be read with no hidden adapter: nothing said of hidden ones" (
        -not ($script:LogLines -match 'Hidden and virtual network adapters') -and -not ((Get-StepNotes) -match 'Hidden and virtual network adapters')) ($script:LogLines -join ' / ')

    Reset-State
    $script:HiddenNics = @([pscustomobject]@{ Name = 'Local Area Connection* 1'; InterfaceDescription = 'Microsoft Wi-Fi Direct Virtual Adapter'; PnPDeviceID = $wfd1; Hidden = $true })
    $script:DevEnable = @([pscustomobject]@{ InstanceName = "${wfd1}_0"; Enable = $true })
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Assert-Case "no physical adapter: a hidden one is still turned off" ((Get-StepStatus) -eq 'Done' -and (Get-FakeDevEnable "${wfd1}_0") -eq $false) (Get-StepNotes)

    # An adapter object without a Hidden property counts as not hidden, and
    # reading it does not throw under StrictMode.
    Reset-State
    $script:HiddenNics = @([pscustomobject]@{ Name = 'OpenVPN TAP'; InterfaceDescription = 'TAP-Windows Adapter V9'; PnPDeviceID = $tapPnp })
    $script:DevEnable = @([pscustomobject]@{ InstanceName = "${tapPnp}_0"; Enable = $true })
    Step 6 19 'nic' -Name 'Network adapter power saving'
    $threw = $null
    try { Disable-NicPowerSaving } catch { $threw = $_.Exception.Message }
    Assert-Case "an adapter with no Hidden property is a virtual adapter, read without an error" ($null -eq $threw -and (Get-StepStatus) -eq 'Done' -and
        (Get-FakeDevEnable "${tapPnp}_0") -eq $false -and
        (Get-StepNotes) -eq 'OpenVPN TAP (TAP-Windows Adapter V9), a virtual adapter: Windows may no longer turn it off to save power') "threw: $threw; $(Get-StepNotes)"

    # -Undo with a hidden adapter's instance gone: it stays recorded.
    Reset-State
    Set-NicMoreFixture
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Complete-Step
    $script:DevEnable = @($script:DevEnable | Where-Object { $_.InstanceName -ne "${wfd1}_0" })
    Step 1 5 'undo' -Name 'Recorded settings'
    $gone = Find-Undo 'NicDeviceEnable' $wfd1
    $ok = $null
    if ($null -ne $gone) { $ok = Restore-PreventUndoItem -Item $gone -CreatedPlan $null }
    Assert-Case "-Undo: a hidden adapter that is gone keeps its record, with a warning" ($ok -eq $false -and (Get-StepStatus) -eq 'Warning') "recorded: $($null -ne $gone); $(Get-StepNotes)"

    # The undo record cannot be saved: a hidden adapter is not changed
    # either, and the step goes on to say so.
    Reset-State
    Set-NicMoreFixture
    $null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $UndoFile -Force
    Step 6 19 'nic' -Name 'Network adapter power saving'
    try {
        Disable-NicPowerSaving
    } finally {
        Microsoft.PowerShell.Management\Remove-Item -LiteralPath $UndoFile -Recurse -Force
    }
    $notes = Get-StepNotes
    Assert-Case "the undo record cannot be saved: a hidden adapter is not changed, and is named" ((Get-StepStatus) -eq 'Failed' -and
        (Get-FakeDevEnable "${wfd1}_0") -eq $true -and (Get-FakeDevEnable "${wfd2}_0") -eq $true -and
        @($script:NicCalls | Where-Object { $_ -like 'Set-CimInstance *vwifimp*' }).Count -eq 0 -and
        $notes -match [regex]::Escape("Local Area Connection* 1 (Microsoft Wi-Fi Direct Virtual Adapter), a hidden adapter: 'allow the computer to turn off this device': Could not save the undo record")) $notes
}

# =====================================================================
# The physical adapters cannot be listed. Until 6.3.0 the step ended
# there, so the hidden and virtual adapters, a list of their own, kept
# Windows' power saving on as well. That list is read all the same now,
# and a physical adapter it holds, known by ConnectorPresent, is left out:
# it is neither hidden nor virtual.
Invoke-Section 'Network adapter power saving: the physical adapters cannot be listed' {
    Reset-State
    Set-NicMoreFixture
    Add-TapFixture
    $script:PhysicalNicThrows = $true
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    $said = 'Could not list the physical network adapters, so their power saving was not changed: The WMI provider is not loaded'
    $notes = Get-StepNotes
    Assert-Case "the physical list fails: Failed, naming why" ((Get-StepStatus) -eq 'Failed' -and @((Get-LastStep).Notes)[0] -eq $said -and
        (Test-LogHas ('^' + [regex]::Escape("[!] $said") + '$'))) "$(Get-StepStatus): $notes"
    Assert-Case "and the hidden and virtual adapters are still turned off, each recorded for -Undo" ((Get-FakeDevEnable "${wfd1}_0") -eq $false -and
        (Get-FakeDevEnable "${wfd2}_0") -eq $false -and (Get-FakeDevEnable "${tapPnp}_0") -eq $false -and
        $null -ne (Find-Undo 'NicDeviceEnable' $wfd1) -and $null -ne (Find-Undo 'NicDeviceEnable' $wfd2) -and
        $null -ne (Find-Undo 'NicDeviceEnable' $tapPnp)) "$notes; $(Get-UndoKeys)"
    Assert-Case "and the physical ones, which that list holds too, are neither changed nor called hidden or virtual" (
        (Get-FakeDevEnable "${ethPnp}_0") -eq $true -and (Get-FakeAdv 'Ethernet' '*EEE') -eq '1' -and
        (Get-FakeRegValue $ethKey 'PnPCapabilities') -eq 24 -and $script:UndoRecord.Count -eq 3 -and
        -not ($notes -match 'Ethernet \(|WiFi \(') -and
        @($script:NicCalls | Where-Object { $_.StartsWith("Set-CimInstance $ethPnp") -or $_.StartsWith("Set-CimInstance $wifiPnp") }).Count -eq 0) (
        "$notes; $(Get-UndoKeys); $($script:NicCalls -join ' / ')")
    Assert-Case "and the step does not say no physical adapter was found" (-not (Test-LogHas '^No physical network adapters found$') -and
        (Get-StepStatus) -ne 'Skipped') ($script:LogLines -join ' / ')

    # A list read with no physical adapter in it is still Skipped for them,
    # and is not a failure.
    Reset-State
    $script:HiddenNics = @([pscustomobject]@{ Name = 'OpenVPN TAP'; InterfaceDescription = 'TAP-Windows Adapter V9'; PnPDeviceID = $tapPnp; Hidden = $false; ConnectorPresent = $false })
    $script:DevEnable = @([pscustomobject]@{ InstanceName = "${tapPnp}_0"; Enable = $true })
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Assert-Case "no physical adapter listed: said, and the virtual one is turned off" ((Test-LogHas '^No physical network adapters found$') -and
        (Get-StepStatus) -eq 'Done' -and (Get-FakeDevEnable "${tapPnp}_0") -eq $false) "$(Get-StepStatus): $(Get-StepNotes)"
}

# =====================================================================
Invoke-Section 'Restore-PreventUndoItem, kind by kind' {
    function New-UndoItem { param([string]$Kind, [string]$Key, $Original) return [pscustomobject]@{ Kind = $Kind; Key = $Key; Original = $Original } }

    Reset-State
    Set-PowerFixture
    $script:PcIdx["$plBal|$subSleep|$setSleep"] = [long]0
    Step 1 5 'undo' -Name 'Recorded settings'
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'PowerValue' "$plBal|$subSleep|$setSleep" 1800) -CreatedPlan $null
    Assert-Case "a power value is put back" ($ok -eq $true -and $script:PcIdx["$plBal|$subSleep|$setSleep"] -eq 1800 -and (Get-StepStatus) -eq 'Done')
    $before = $script:PcCalls.Count
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'PowerValue' "$plHigh|$subSleep|$setSleep" 1800) -CreatedPlan $plHigh
    Assert-Case "a value on the plan this script created is left for the plan's removal" ($ok -eq $true -and
        @($script:PcCalls | Select-Object -Skip $before | Where-Object { $_ -like '/setacvalueindex*' }).Count -eq 0)
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'PowerValue' "deadbeef-0000-0000-0000-000000000000|$subSleep|$setSleep" 1800) -CreatedPlan $null
    Assert-Case "a value on a plan that is gone has nothing to put back" ($ok -eq $true)
    $script:PcRefuse["$plBal|$subSleep|$setSleep"] = $true
    $script:PcIdx["$plBal|$subSleep|$setSleep"] = [long]0
    Step 1 5 'undo' -Name 'Recorded settings'
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'PowerValue' "$plBal|$subSleep|$setSleep" 1800) -CreatedPlan $null
    Assert-Case "a value powercfg refuses stays recorded, and fails the step by name" ($ok -eq $false -and (Get-StepStatus) -eq 'Failed' -and (Get-StepNotes) -match '^Sleep after could not be put back$') (Get-StepNotes)
    $threw = $false
    try { $null = Restore-PreventUndoItem -Item (New-UndoItem 'PowerValue' 'a|b' 1) -CreatedPlan $null } catch { $threw = $true }
    Assert-Case "a power record that is not plan|sub|setting throws" $threw

    Reset-State
    Set-FakeRegValue $powerKey 'HibernateEnabled' ([int]0)
    Step 1 5 'undo' -Name 'Recorded settings'
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'Hibernate' 'Hibernate' 1) -CreatedPlan $null
    Assert-Case "hibernate is turned back on" ($ok -eq $true -and (Get-FakeRegValue $powerKey 'HibernateEnabled') -eq 1)
    Set-FakeRegValue $powerKey 'HibernateEnabled' ([int]0)
    $script:HibernateRefuse = $true
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'Hibernate' 'Hibernate' 1) -CreatedPlan $null
    Assert-Case "hibernate that will not come back stays recorded" ($ok -eq $false -and (Get-StepStatus) -eq 'Failed')

    Reset-State
    Set-FakeRegValue 'HKLM:\T\K' 'Mine' ([int]1)
    Set-FakeRegValue 'HKLM:\T\K' 'Was5' ([int]1)
    Step 1 5 'undo' -Name 'Recorded settings'
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'Registry' 'HKLM:\T\K|Mine' $null) -CreatedPlan $null
    Assert-Case "a value there was none of is removed" ($ok -eq $true -and -not (Test-FakeRegValue 'HKLM:\T\K' 'Mine'))
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'Registry' 'HKLM:\T\K|Was5' 5) -CreatedPlan $null
    $now = Get-FakeRegValue 'HKLM:\T\K' 'Was5'
    Assert-Case "a value is put back as a DWORD" ($ok -eq $true -and $now -is [int] -and $now -eq 5)
    Assert-Case "and the step waits for a restart" ((Get-LastStep).Reboot -eq $true)
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'Registry' 'HKLM:\T\Gone|V' 5) -CreatedPlan $null
    Assert-Case "a value in a key that is gone has nothing to put back" ($ok -eq $true)
    $threw = $false
    try { $null = Restore-PreventUndoItem -Item (New-UndoItem 'Registry' 'nobar' 5) -CreatedPlan $null } catch { $threw = $true }
    Assert-Case "a registry record with no value name throws" $threw

    Reset-State
    $null = Add-FakeRegKey 'HKLM:\T\Empty'
    Set-FakeRegValue 'HKLM:\T\Used' 'Theirs' ([int]1)
    Step 1 5 'undo' -Name 'Recorded settings'
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'RegistryKey' 'HKLM:\T\Empty' $null) -CreatedPlan $null
    Assert-Case "a key this script created is removed when empty" ($ok -eq $true -and -not (Test-Path -LiteralPath 'HKLM:\T\Empty'))
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'RegistryKey' 'HKLM:\T\Used' $null) -CreatedPlan $null
    Assert-Case "and left when something else put a value in it" ($ok -eq $true -and (Test-Path -LiteralPath 'HKLM:\T\Used'))

    Reset-State
    Set-NicFixture
    $script:DevEnable[0].Enable = $false
    $script:DevEnable[1].Enable = $false
    Step 1 5 'undo' -Name 'Recorded settings'
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'NicDeviceEnable' $ethPnp $true) -CreatedPlan $null
    Assert-Case "the checkbox is put back on its own adapter only" ($ok -eq $true -and $script:DevEnable[0].Enable -eq $true -and $script:DevEnable[1].Enable -eq $false)
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'NicDeviceEnable' 'PCI\NOPE' $true) -CreatedPlan $null
    Assert-Case "an adapter that is not plugged in stays recorded" ($ok -eq $false -and (Get-StepStatus) -eq 'Warning')

    Reset-State
    Set-NicFixture
    $script:Adv['Ethernet'][0].RegistryValue = [string[]]@('0')
    Step 1 5 'undo' -Name 'Recorded settings'
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'NicAdvanced' "$ethPnp|*EEE" @('1')) -CreatedPlan $null
    Assert-Case "an option is put back, without a restart" ($ok -eq $true -and (@($script:Adv['Ethernet'][0].RegistryValue) -join ',') -eq '1' -and
        ($script:NicCalls -join ',') -eq 'Set-Adv Ethernet|*EEE=1 NoRestart=True') ($script:NicCalls -join ',')
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'NicAdvanced' "$ethPnp|GigaLite" $null) -CreatedPlan $null
    Assert-Case "an option that was unset is reset" ($ok -eq $true -and $null -eq $script:Adv['Ethernet'][6].RegistryValue)
    $before = $script:NicCalls.Count
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'NicAdvanced' "$ethPnp|NoSuchKeyword" @('1')) -CreatedPlan $null
    Assert-Case "an option the driver no longer has has nothing to put back" ($ok -eq $true -and $script:NicCalls.Count -eq $before)
    $gone = 'Ethernet: its driver no longer lists the option NoSuchKeyword, so it was not put back and its record is dropped'
    Assert-Case "and its record is dropped with a warning that names the adapter and the option" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq $gone -and (Test-LogHas ('^' + [regex]::Escape("[!] $gone") + '$'))) "$(Get-StepStatus): $(Get-StepNotes)"
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'NicAdvanced' 'PCI\NOPE|*EEE' @('1')) -CreatedPlan $null
    Assert-Case "an option on an adapter not plugged in stays recorded" ($ok -eq $false)

    Reset-State
    Set-NicFixture
    $script:Pm['Ethernet'].SelectiveSuspend = 'Disabled'
    Step 1 5 'undo' -Name 'Recorded settings'
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'NicPowerManagement' "$ethPnp|SelectiveSuspend" 'Enabled') -CreatedPlan $null
    Assert-Case "selective suspend is put back, without a restart" ($ok -eq $true -and $script:Pm['Ethernet'].SelectiveSuspend -eq 'Enabled' -and
        ($script:NicCalls -join ',') -eq 'Set-Pm Ethernet|SelectiveSuspend=Enabled NoRestart=True') ($script:NicCalls -join ',')
    $threw = $false
    try { $null = Restore-PreventUndoItem -Item (New-UndoItem 'NicPowerManagement' "$ethPnp|WakeOnMagicPacket" 'Enabled') -CreatedPlan $null } catch { $threw = $true }
    Assert-Case "a setting this script never changes is refused" $threw
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'NicPowerManagement' 'PCI\NOPE|SelectiveSuspend' 'Enabled') -CreatedPlan $null
    Assert-Case "a setting on an adapter not plugged in stays recorded" ($ok -eq $false)

    Reset-State
    Step 1 5 'undo' -Name 'Recorded settings'
    $ok1 = Restore-PreventUndoItem -Item (New-UndoItem 'DefenderPath' 'C:\Users\X\AppData\Local\Claude-Data' $null) -CreatedPlan $null
    $ok2 = Restore-PreventUndoItem -Item (New-UndoItem 'DefenderProcess' 'vmwp.exe' $null) -CreatedPlan $null
    Assert-Case "the Defender exclusions this script added are removed, and only those" ($ok1 -and $ok2 -and
        ($script:MpCalls -join ',') -eq 'ExclusionPath C:\Users\X\AppData\Local\Claude-Data,ExclusionProcess vmwp.exe') ($script:MpCalls -join ',')
    $script:MpFails = $true
    $threw = $false
    try { $null = Restore-PreventUndoItem -Item (New-UndoItem 'DefenderPath' 'C:\X' $null) -CreatedPlan $null } catch { $threw = $true }
    Assert-Case "an exclusion that cannot be removed throws, so it stays recorded" $threw

    Reset-State
    $svcKey = 'HKLM:\SYSTEM\CurrentControlSet\Services\vmcompute'
    $script:Services['vmcompute'] = 'Running'
    Set-FakeRegValue $svcKey 'FailureActions' (New-FailureBlob -Reset 300 -Pairs @(1, 30000, 1, 60000, 1, 120000))
    Step 1 5 'undo' -Name 'Recorded settings'
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'ServiceFailureActions' 'vmcompute' $null) -CreatedPlan $null
    Assert-Case "a service that had no recovery has none again, and no value left" ($ok -eq $true -and -not (Test-FakeRegValue $svcKey 'FailureActions') -and
        ($script:ScCalls -join ',') -eq 'failure vmcompute reset= 0 actions= ""') ($script:ScCalls -join ',')
    $orig = New-FailureBlob -Reset 86400 -Pairs @(1, 60000, 2, 120000)
    $script:ScCalls = New-Object System.Collections.ArrayList
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'ServiceFailureActions' 'vmcompute' ([Convert]::ToBase64String($orig))) -CreatedPlan $null
    $now = ConvertFrom-FailureActionsBlob -Bytes (Get-FakeRegValue $svcKey 'FailureActions')
    Assert-Case "restart and reboot actions are put back through sc.exe" ($ok -eq $true -and ($script:ScCalls -join ',') -eq 'failure vmcompute reset= 86400 actions= restart/60000/reboot/120000' -and
        $now.Reset -eq 86400 -and @($now.Actions).Count -eq 2) ($script:ScCalls -join ',')
    $run = New-FailureBlob -Reset 60 -Pairs @(3, 1000, 1, 5000)
    $script:ScCalls = New-Object System.Collections.ArrayList
    Step 1 5 'undo' -Name 'Recorded settings'
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'ServiceFailureActions' 'vmcompute' ([Convert]::ToBase64String($run))) -CreatedPlan $null
    $now = Get-FakeRegValue $svcKey 'FailureActions'
    Assert-Case "an action sc.exe cannot write is put back as the value, after a restart" ($ok -eq $true -and $script:ScCalls.Count -eq 0 -and
        [Convert]::ToBase64String($now) -eq [Convert]::ToBase64String($run) -and (Get-LastStep).Reboot -eq $true)
    $script:ScFails = $true
    $threw = $false
    try { $null = Restore-PreventUndoItem -Item (New-UndoItem 'ServiceFailureActions' 'vmcompute' $null) -CreatedPlan $null } catch { $threw = $true }
    Assert-Case "sc.exe refusing throws, so it stays recorded" $threw
    $script:Services.Remove('vmcompute')
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'ServiceFailureActions' 'vmcompute' $null) -CreatedPlan $null
    Assert-Case "a service that is gone has nothing to put back" ($ok -eq $true)

    $threw = $false
    try { $null = Restore-PreventUndoItem -Item (New-UndoItem 'Teleporter' 'x' $null) -CreatedPlan $null } catch { $threw = $_.Exception.Message -match 'does not know' }
    Assert-Case "a kind this version does not know throws" $threw
}

# =====================================================================
Invoke-Section '-Undo: a key this script created, with a value still recorded in it' {
    function New-UndoItem { param([string]$Kind, [string]$Key, $Original) return [pscustomobject]@{ Kind = $Kind; Key = $Key; Original = $Original } }

    Reset-State
    Set-FakeRegValue 'HKLM:\T\Throttle' 'PowerThrottlingOff' ([int]1)
    $null = $script:UndoRecord.Add((New-UndoItem 'Registry' 'hklm:\t\throttle|PowerThrottlingOff' $null))
    Step 1 5 'undo' -Name 'Recorded settings'
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'RegistryKey' 'HKLM:\T\Throttle' $null) -CreatedPlan $null
    Assert-Case "the key keeps its record, and stays" ($ok -eq $false -and (Test-Path -LiteralPath 'HKLM:\T\Throttle')) "got $ok"
    Assert-Case "and is not said to hold settings this script did not make" (-not (Test-LogHas 'did not make') -and (Get-StepNotes) -notmatch 'did not make') (Get-StepNotes)
    $script:UndoRecord = New-Object System.Collections.ArrayList
    $null = $script:UndoRecord.Add((New-UndoItem 'Registry' 'HKLM:\T\EmptyX|V' $null))
    $null = $script:UndoRecord.Add((New-UndoItem 'Registry' 'HKLM:\T\Empty\Sub|V' $null))
    $null = Add-FakeRegKey 'HKLM:\T\Empty'
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'RegistryKey' 'HKLM:\T\Empty' $null) -CreatedPlan $null
    Assert-Case "a value recorded in another key, or one whose name starts the same, does not hold it" ($ok -eq $true -and -not (Test-Path -LiteralPath 'HKLM:\T\Empty')) "got $ok"

    Reset-State
    Set-PowerFixture
    $null = Write-RecordedDword -Path 'HKLM:\T\Throttle' -Name 'PowerThrottlingOff' -Value 1
    $script:RegRemoveFault['hklm:\t\throttle|powerthrottlingoff'] = $true
    Invoke-UndoStep1
    Assert-Case "-Undo: a value that cannot be removed stays recorded, and so does the key it is in" ($null -ne (Find-Undo 'Registry' 'HKLM:\T\Throttle|PowerThrottlingOff') -and
        $null -ne (Find-Undo 'RegistryKey' 'HKLM:\T\Throttle') -and (Test-Path -LiteralPath 'HKLM:\T\Throttle') -and (Get-StepStatus) -eq 'Failed') "$(Get-UndoKeys); $(Get-StepStatus)"
    Assert-Case "and the key is not said to hold settings this script did not make" (-not (Test-LogHas 'did not make')) ($script:LogLines -join ' / ')
    $script:RegRemoveFault = @{}
    $script:StepResults = New-Object System.Collections.ArrayList
    Invoke-UndoStep1
    Assert-Case "the next -Undo removes the value, then the key" (-not (Test-Path -LiteralPath 'HKLM:\T\Throttle') -and $script:UndoRecord.Count -eq 0) (Get-UndoKeys)

    Reset-State
    Set-PowerFixture
    Set-FakeRegValue 'HKLM:\T\K2' 'V' ([int]1)
    Add-PreventUndoItem -Kind Registry -Key 'HKLM:\T\K2|V' -Original $null
    Add-PreventUndoItem -Kind RegistryKey -Key 'HKLM:\T\K2' -Original $null
    Invoke-UndoStep1
    Assert-Case "a value recorded before the key it is in: both go in one -Undo" (-not (Test-Path -LiteralPath 'HKLM:\T\K2') -and $script:UndoRecord.Count -eq 0) (Get-UndoKeys)
}

# =====================================================================
Invoke-Section 'powercfg that cannot list the plans' {
    function New-UndoItem { param([string]$Kind, [string]$Key, $Original) return [pscustomobject]@{ Kind = $Kind; Key = $Key; Original = $Original } }

    Reset-State
    Set-PowerFixture
    Assert-Case "a plan powercfg lists: true" ((Test-PowerPlanExists -Plan $plHigh) -eq $true)
    Assert-Case "a plan it does not: false" ((Test-PowerPlanExists -Plan 'deadbeef-0000-0000-0000-000000000000') -eq $false)
    $script:ListOk = 0
    $e = Test-PowerPlanExists -Plan $plHigh
    Assert-Case "powercfg /list failing: null, not false" ($null -eq $e) "got '$e'"

    $script:PcIdx["$plBal|$subSleep|$setSleep"] = [long]0
    $script:PcCalls = New-Object System.Collections.ArrayList
    Step 1 5 'undo' -Name 'Recorded settings'
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'PowerValue' "$plBal|$subSleep|$setSleep" 1800) -CreatedPlan $null
    Assert-Case "-Undo: a power value whose plan cannot be listed stays recorded, is left alone, and fails the step" ($ok -eq $false -and (Get-StepStatus) -eq 'Failed' -and
        $script:PcIdx["$plBal|$subSleep|$setSleep"] -eq 0 -and @($script:PcCalls | Where-Object { $_ -like '/setacvalueindex*' }).Count -eq 0) "got $ok; $(Get-StepNotes)"
    $script:ListOk = -1
    $script:QhFails = $true
    Step 1 5 'undo' -Name 'Recorded settings'
    $ok = Restore-PreventUndoItem -Item (New-UndoItem 'PowerValue' "$plBal|$subSleep|$setSleep" 1800) -CreatedPlan $null
    Assert-Case "a power value that cannot be read back stays recorded, and fails the step" ($ok -eq $false -and (Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -match 'could not be read back') "got $ok; $(Get-StepNotes)"
}

# =====================================================================
Invoke-Section 'Setup step 1: the power plan' {
    Reset-State
    Set-PowerFixture
    Invoke-SetupStep1
    $created = (Find-Undo 'PowerPlan' 'Created')
    $active = (Find-Undo 'PowerPlan' 'Active')
    Assert-Case "first run: the plan you had is recorded" ($null -ne $active -and $active.Original -eq $plBal) (Get-UndoKeys)
    Assert-Case "an Ultimate Performance plan is created, recorded before it is made" ($null -ne $created -and $script:Plans.Contains([string]$created.Original) -and
        (Get-UndoKeys) -match 'PowerPlan:Active ; PowerPlan:Created') (Get-UndoKeys)
    Assert-Case "and made active" ($null -ne $created -and $script:ActivePlan -eq $created.Original -and (Get-StepStatus) -eq 'Done') "active $($script:ActivePlan), $(Get-StepStatus)"
    $planCount = $script:Plans.Count
    $script:StepResults = New-Object System.Collections.ArrayList
    $script:PcCalls = New-Object System.Collections.ArrayList
    Invoke-SetupStep1
    Assert-Case "second run: the plan you had is still the one recorded" ((Find-Undo 'PowerPlan' 'Active').Original -eq $plBal)
    Assert-Case "and no second copy is made" ($script:Plans.Count -eq $planCount -and @($script:PcCalls | Where-Object { $_ -like '/duplicatescheme*' }).Count -eq 0 -and
        (Get-StepStatus) -eq 'Unchanged') "$($script:Plans.Count) plans, $(Get-StepStatus)"

    Reset-State
    Set-PowerFixture
    $ultOld = '17bdb7be-3cb2-4559-9868-fd8383619e03'
    Add-FakePlan $ultOld 'Ultimate Performance' $ultFriendly
    $script:ActivePlan = $ultOld
    [System.IO.File]::WriteAllText($BackupFile, "$ultOld`r`n")
    Invoke-SetupStep1
    Assert-Case "an older version's record of its own plan: -Undo returns to Balanced, and says so" ((Find-Undo 'PowerPlan' 'Active').Original -eq $plBal -and
        (Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -match 'overwrote its record') (Get-StepNotes)
    Assert-Case "the old record is removed and no copy is made" (-not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $BackupFile) -and
        $null -eq (Find-Undo 'PowerPlan' 'Created') -and $script:ActivePlan -eq $ultOld)

    Reset-State
    Set-PowerFixture
    Add-FakePlan $ultOld 'Ultimate Performance' $ultFriendly
    $script:ActivePlan = $plHigh
    [System.IO.File]::WriteAllText($BackupFile, $plHigh.ToUpperInvariant())
    Invoke-SetupStep1
    Assert-Case "an older version's record of your plan is carried over" ((Find-Undo 'PowerPlan' 'Active').Original -eq $plHigh -and
        -not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $BackupFile))
    Assert-Case "and the Ultimate Performance plan already there is used" ($script:ActivePlan -eq $ultOld -and $null -eq (Find-Undo 'PowerPlan' 'Created') -and (Get-StepStatus) -eq 'Done')

    Reset-State
    $script:Plans = [ordered]@{}
    Add-FakePlan $plBal 'Gebalanceerd' '@%SystemRoot%\system32\powrprof.dll,-2,Balanced'
    $ultNl = '0aa0aa0a-1111-2222-3333-444455556666'
    Add-FakePlan $ultNl 'Ultieme prestaties' $ultFriendly
    Set-FakeRegValue "$schemesRoot\$plTemplate" 'FriendlyName' $ultFriendly
    $script:ActivePlan = $plBal
    Invoke-SetupStep1
    Assert-Case "a Dutch Windows' Ultimate Performance copy is found by what it is" ($script:ActivePlan -eq $ultNl -and
        @($script:PcCalls | Where-Object { $_ -like '/duplicatescheme*' }).Count -eq 0) "active $($script:ActivePlan)"

    Reset-State
    Set-PowerFixture
    $script:NoTemplate = $true
    Invoke-SetupStep1
    Assert-Case "no template to copy: High performance, and no plan to delete on -Undo" ($script:ActivePlan -eq $plHigh -and
        $null -eq (Find-Undo 'PowerPlan' 'Created') -and (Get-StepStatus) -eq 'Done') "$($script:ActivePlan); $(Get-UndoKeys)"

    Reset-State
    $script:Plans = [ordered]@{}
    Add-FakePlan $plBal 'Balanced' '@%SystemRoot%\system32\powrprof.dll,-2,Balanced'
    $script:ActivePlan = $plBal
    $script:NoTemplate = $true
    Invoke-SetupStep1
    Assert-Case "neither plan: your plan stays, with a warning" ($script:ActivePlan -eq $plBal -and (Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -match 'No Ultimate Performance or High performance plan') (Get-StepNotes)

    Reset-State
    Set-PowerFixture
    $script:SetActiveRefuse = $true
    Invoke-SetupStep1
    Assert-Case "a plan powercfg will not activate: Failed" ((Get-StepStatus) -eq 'Failed' -and (Get-StepNotes) -match 'powercfg could not make .* the active plan') (Get-StepNotes)

    Reset-State
    Set-PowerFixture
    Add-PreventUndoItem -Kind PowerPlan -Key Active -Original $plSaver
    Add-PreventUndoItem -Kind PowerPlan -Key Created -Original '99999999-9999-9999-9999-999999999999'
    Invoke-SetupStep1
    $c = Find-Undo 'PowerPlan' 'Created'
    Assert-Case "a created plan someone deleted is forgotten, and a new one made" ($null -ne $c -and $c.Original -ne '99999999-9999-9999-9999-999999999999' -and
        $script:ActivePlan -eq $c.Original -and (Find-Undo 'PowerPlan' 'Active').Original -eq $plSaver)
}

# =====================================================================
Invoke-Section "Setup step 1: an older version's record that is no use, and powercfg that cannot list" {
    $ultOld = '17bdb7be-3cb2-4559-9868-fd8383619e03'
    foreach ($c in @(
            @('deadbeef-0000-0000-0000-000000000000', 'no longer exists', 'a plan since deleted'),
            @('High performance', 'cannot be read', 'no plan id')
        )) {
        Reset-State
        Set-PowerFixture
        Add-FakePlan $ultOld 'Ultimate Performance' $ultFriendly
        $script:ActivePlan = $ultOld
        [System.IO.File]::WriteAllText($BackupFile, "$($c[0])`r`n")
        Invoke-SetupStep1
        $a = Find-Undo 'PowerPlan' 'Active'
        Assert-Case "an older version's record naming $($c[2]): -Undo returns to Balanced, with a warning" ($null -ne $a -and $a.Original -eq $plBal -and
            (Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -match $c[1]) "$(Get-UndoKeys); $(Get-StepNotes)"
        Assert-Case "and that record is removed ($($c[2]))" (-not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $BackupFile))
    }

    Reset-State
    Set-PowerFixture
    [System.IO.File]::WriteAllText($BackupFile, $plHigh)
    $script:ListOk = 0
    Invoke-SetupStep1
    Assert-Case "an older version's record is kept while powercfg cannot say whether its plan is there" ((Microsoft.PowerShell.Management\Test-Path -LiteralPath $BackupFile) -and
        $null -eq (Find-Undo 'PowerPlan' 'Active') -and (Get-StepStatus) -eq 'Failed') "$(Get-UndoKeys); $(Get-StepNotes)"

    $made = '0fff0fff-1111-2222-3333-444455556666'
    Reset-State
    Set-PowerFixture
    Add-FakePlan $made 'Ultimate Performance' $ultFriendly
    $script:ActivePlan = $made
    Add-PreventUndoItem -Kind PowerPlan -Key Active -Original $plBal
    Add-PreventUndoItem -Kind PowerPlan -Key Created -Original $made
    $script:ListOk = 0
    Invoke-SetupStep1
    Assert-Case "powercfg cannot list the plans: the plan this script created stays recorded, and the step fails" ($null -ne (Find-Undo 'PowerPlan' 'Created') -and
        (Get-StepStatus) -eq 'Failed') "$(Get-UndoKeys); $(Get-StepNotes)"

    Reset-State
    Set-PowerFixture
    $script:ListOk = 0
    Invoke-SetupStep1
    $c = Find-Undo 'PowerPlan' 'Created'
    Assert-Case "a plan made while powercfg cannot list the plans stays recorded for -Undo, and the step fails" ($null -ne $c -and $script:Plans.Contains([string]$c.Original) -and
        (Get-StepStatus) -eq 'Failed') "$(Get-UndoKeys); $(Get-StepNotes)"

    Reset-State
    Set-PowerFixture
    $script:NoTemplate = $true
    $script:ListOk = 0
    $script:PcCalls = New-Object System.Collections.ArrayList
    Invoke-SetupStep1
    Assert-Case "no plan to copy, and powercfg cannot say whether High performance is there: Failed, and the plan left as it is" ((Get-StepStatus) -eq 'Failed' -and
        $script:ActivePlan -eq $plBal -and @($script:PcCalls | Where-Object { $_ -like '/setactive*' }).Count -eq 0) "$(Get-StepStatus): $(Get-StepNotes)"
}

# =====================================================================
Invoke-Section '-Undo step 1: the recorded settings and the plan' {
    Reset-State
    Set-PowerFixture
    $script:PcIdx["$plBal|$subSleep|$setSleep"] = [long]1800
    Invoke-SetupStep1
    $created = [string](Find-Undo 'PowerPlan' 'Created').Original
    $script:PcIdx["$created|$subSleep|$setSleep"] = [long]1800
    $null = Write-AcPowerValue -Plan $created -Sub $subSleep -Setting $setSleep -Value 0
    $null = Write-AcPowerValue -Plan $plBal -Sub $subSleep -Setting $setSleep -Value 0
    $null = Write-RecordedDword -Path 'HKLM:\T\Throttle' -Name 'PowerThrottlingOff' -Value 1
    $ultUser = '0bb0bb0b-1111-2222-3333-444455556666'
    Add-FakePlan $ultUser 'My own Ultimate' $ultFriendly
    $script:UndoRecord = New-Object System.Collections.ArrayList
    Read-PreventUndo
    $script:StepResults = New-Object System.Collections.ArrayList
    $script:PcCalls = New-Object System.Collections.ArrayList
    Invoke-UndoStep1
    Assert-Case "the plan you had is active again" ($script:ActivePlan -eq $plBal)
    Assert-Case "its setting is put back" ($script:PcIdx["$plBal|$subSleep|$setSleep"] -eq 1800)
    Assert-Case "the plan this script created is deleted" (-not $script:Plans.Contains($created))
    Assert-Case "and a value on it was not set before it went" (@($script:PcCalls | Where-Object { $_ -like "/setacvalueindex $created*" }).Count -eq 0)
    Assert-Case "an Ultimate Performance plan you made yourself is kept" ($script:Plans.Contains($ultUser) -and $script:Plans.Contains($plHigh))
    Assert-Case "the registry value is removed again" (-not (Test-FakeRegValue 'HKLM:\T\Throttle' 'PowerThrottlingOff'))
    Assert-Case "and the key it made" (-not (Test-Path -LiteralPath 'HKLM:\T\Throttle'))
    Assert-Case "everything put back: no record left" ($script:UndoRecord.Count -eq 0 -and -not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $UndoFile)) (Get-UndoKeys)
    Assert-Case "the step reports Done" ((Get-StepStatus) -eq 'Done') "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    Set-PowerFixture
    Add-PreventUndoItem -Kind PowerPlan -Key Active -Original 'deadbeef-0000-0000-0000-000000000000'
    $script:ActivePlan = $plHigh
    Invoke-UndoStep1
    Assert-Case "the plan you had is gone: Balanced, with a warning" ($script:ActivePlan -eq $plBal -and (Get-StepStatus) -eq 'Warning') (Get-StepNotes)

    Reset-State
    Set-PowerFixture
    Add-FakePlan '0ccc0ccc-1111-2222-3333-444455556666' 'Ultimate Performance' $ultFriendly
    $script:ActivePlan = '0ccc0ccc-1111-2222-3333-444455556666'
    Add-PreventUndoItem -Kind PowerPlan -Key Active -Original $plBal
    Add-PreventUndoItem -Kind PowerPlan -Key Created -Original '0ccc0ccc-1111-2222-3333-444455556666'
    $script:SetActiveRefuse = $true
    Invoke-UndoStep1
    Assert-Case "a plan that cannot be switched back: Failed, and the created plan kept" ((Get-StepStatus) -eq 'Failed' -and
        $script:Plans.Contains('0ccc0ccc-1111-2222-3333-444455556666') -and (Get-StepNotes) -match 'Kept the plan this script created') (Get-StepNotes)
    Assert-Case "both stay recorded for the next -Undo" ($null -ne (Find-Undo 'PowerPlan' 'Active') -and $null -ne (Find-Undo 'PowerPlan' 'Created'))

    Reset-State
    Set-PowerFixture
    Invoke-UndoStep1
    Assert-Case "nothing recorded: Unchanged" ((Get-StepStatus) -eq 'Unchanged' -and (Test-LogHas 'Nothing recorded'))
}

# =====================================================================
Invoke-Section '-Undo step 1: what does not stop it, and what it does not touch' {
    $made = '0ddd0ddd-1111-2222-3333-444455556666'
    Reset-State
    Set-PowerFixture
    Add-FakePlan $made 'Ultimate Performance' $ultFriendly
    $script:ActivePlan = $made
    Add-PreventUndoItem -Kind PowerPlan -Key Active -Original $plBal
    Add-PreventUndoItem -Kind PowerPlan -Key Created -Original $made
    Microsoft.PowerShell.Management\Remove-Item -LiteralPath $UndoFile -Force
    $null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $UndoFile -Force
    $threw = ''
    try { Invoke-UndoStep1 } catch { $threw = $_.Exception.Message }
    Assert-Case "a record that cannot be saved in the plan part does not stop -Undo, so steps 2 to 5 run" ($threw -eq '') $threw
    Assert-Case "the step fails, naming the power plan" ((Get-StepStatus) -eq 'Failed' -and (Get-StepNotes) -match 'Power plan: ') (Get-StepNotes)
    Assert-Case "the plan you had is active, and both plans stay recorded" ($script:ActivePlan -eq $plBal -and $null -ne (Find-Undo 'PowerPlan' 'Active') -and
        $null -ne (Find-Undo 'PowerPlan' 'Created')) (Get-UndoKeys)

    Reset-State
    Set-PowerFixture
    Add-FakePlan $made 'Ultimate Performance' $ultFriendly
    $script:ActivePlan = $made
    Add-PreventUndoItem -Kind PowerPlan -Key Active -Original $plHigh
    Add-PreventUndoItem -Kind PowerPlan -Key Created -Original $made
    $script:ListOk = 0
    $script:PcCalls = New-Object System.Collections.ArrayList
    Invoke-UndoStep1
    $acts = @($script:PcCalls | Where-Object { $_ -like '/setactive*' -or $_ -like '/delete*' })
    Assert-Case "powercfg cannot list the plans: no plan is switched or deleted" ($acts.Count -eq 0 -and $script:ActivePlan -eq $made -and $script:Plans.Contains($made)) ($script:PcCalls -join ' / ')
    Assert-Case "both plans stay recorded, and the step fails" ($null -ne (Find-Undo 'PowerPlan' 'Active') -and $null -ne (Find-Undo 'PowerPlan' 'Created') -and
        (Get-StepStatus) -eq 'Failed') "$(Get-UndoKeys); $(Get-StepNotes)"

    Reset-State
    Set-PowerFixture
    Add-FakePlan $made 'Ultimate Performance' $ultFriendly
    Add-PreventUndoItem -Kind PowerPlan -Key Created -Original $made
    $script:DeleteRefuse = $true
    $script:ListOk = 1
    Invoke-UndoStep1
    Assert-Case "a plan whose removal powercfg cannot confirm stays recorded, and the step fails" ($script:Plans.Contains($made) -and
        $null -ne (Find-Undo 'PowerPlan' 'Created') -and (Get-StepStatus) -eq 'Failed') "$(Get-UndoKeys); $(Get-StepNotes)"

    Reset-State
    Set-PowerFixture
    $script:ActivePlan = $plHigh
    $foreign = '{ "Version": 1, "ComputerName": "OTHER-PC", "Items": [ { "Kind": "PowerPlan", "Key": "Active", "Original": "' + $plBal + '" } ] }'
    [System.IO.File]::WriteAllText($UndoFile, $foreign)
    Read-PreventUndo
    $script:UndoMode = $true
    Invoke-UndoStep1
    Assert-Case "a record written on another computer: Failed, and nothing put back from it" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -match 'another computer \(OTHER-PC\)' -and $script:ActivePlan -eq $plHigh -and [System.IO.File]::ReadAllText($UndoFile) -eq $foreign) "$(Get-StepNotes); active $($script:ActivePlan)"
    $text = ''
    try { $text = @(Write-UndoOutcome 6>&1 | ForEach-Object { [string]$_ }) -join "`n" } catch { $text = "threw: $($_.Exception.Message)" }
    Assert-Case "and the summary names that computer and the file, which is left as it is" ($text -match 'another computer \(OTHER-PC\)' -and $text.Contains($UndoFile) -and
        (Get-PreventExitCode) -eq 1) $text
}

# =====================================================================
Invoke-Section "-Undo step 2: an older version's record of your plan" {
    Reset-State
    Set-PowerFixture
    [System.IO.File]::WriteAllText($BackupFile, "$plHigh`r`n")
    $script:ListOk = 0
    $script:PcCalls = New-Object System.Collections.ArrayList
    Invoke-UndoStep2
    Assert-Case "powercfg cannot list the plans: the record is kept and no plan is switched" ((Microsoft.PowerShell.Management\Test-Path -LiteralPath $BackupFile) -and
        $script:ActivePlan -eq $plBal -and @($script:PcCalls | Where-Object { $_ -like '/setactive*' }).Count -eq 0) ($script:PcCalls -join ' / ')
    Assert-Case "and the step fails" ((Get-StepStatus) -eq 'Failed') "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    Set-PowerFixture
    [System.IO.File]::WriteAllText($BackupFile, "$plHigh`r`n")
    Invoke-UndoStep2
    Assert-Case "powercfg working: the plan it recorded is made active, and the record removed" ($script:ActivePlan -eq $plHigh -and
        -not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $BackupFile) -and (Get-StepStatus) -eq 'Done') "$(Get-StepStatus): $(Get-StepNotes)"
}

# =====================================================================
Invoke-Section 'Setup step 8: only the Cowork VM is raised' {
    Reset-State
    Set-VmFixture
    Invoke-Step8
    Assert-Case "the Cowork VM's worker is raised" ((Get-Priority 10) -eq 'AboveNormal')
    Assert-Case "WSL's and another VM's are not" ((Get-Priority 11) -eq 'Normal' -and (Get-Priority 12) -eq 'Normal')
    Assert-Case "Done" ((Get-StepStatus) -eq 'Done') "$(Get-StepStatus): $(Get-StepNotes)"
    $script:StepResults = New-Object System.Collections.ArrayList
    Invoke-Step8
    Assert-Case "already raised: Unchanged" ((Get-StepStatus) -eq 'Unchanged')

    Reset-State
    Set-VmFixture
    $script:Procs = @((New-FakeProcess 11 'vmwp'), (New-FakeProcess 12 'vmwp'))
    Invoke-Step8
    Assert-Case "the worker exited after the query: Skipped, not Failed" ((Get-StepStatus) -eq 'Skipped' -and (Get-StepNotes) -match "worker exited") (Get-StepNotes)

    Reset-State
    Set-VmFixture
    $script:HcsLive = New-Object System.Collections.ArrayList
    Invoke-Step8
    Assert-Case "no Cowork VM: Skipped, nothing raised" ((Get-StepStatus) -eq 'Skipped' -and (Get-Priority 11) -eq 'Normal' -and (Get-Priority 12) -eq 'Normal') (Get-StepNotes)

    Reset-State
    Set-VmFixture
    $script:HcsListCode = 1
    Invoke-Step8
    Assert-Case "hcsdiag fails: Failed, nothing raised" ((Get-StepStatus) -eq 'Failed' -and (Get-StepNotes) -match 'hcsdiag list returned 1' -and (Get-Priority 10) -eq 'Normal') (Get-StepNotes)

    # A process list that cannot be read: read as no worker, it said the
    # Cowork VM was not running.
    Reset-State
    Set-VmFixture
    $script:VmWorkersFault = $true
    Invoke-Step8
    Assert-Case "the VM workers cannot be listed: Failed, saying so, and nothing raised" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq 'Could not list the VM worker processes: The RPC server is unavailable.' -and (Get-Priority 10) -eq 'Normal' -and
        -not (Test-LogHas 'is not running')) "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    Set-VmFixture
    $script:ClaudeEnv.HcsDiagPath = $null
    try { Invoke-Step8 } finally { $script:ClaudeEnv.HcsDiagPath = 'Invoke-FakeHcs' }
    Assert-Case "no hcsdiag: Skipped" ((Get-StepStatus) -eq 'Skipped' -and (Get-StepNotes) -match 'hcsdiag.exe is not available')

    # 6.3.1: the discovery leaves HcsDiagPath unset for an hcsdiag.exe that
    # could not be read too. Until 6.3.1 that was Skipped, "not available",
    # as for one that is not there.
    Reset-State
    Set-VmFixture
    $script:ClaudeEnv.HcsDiagPath = $null
    $script:FakeHcsDiagPath = Join-Path $sandbox 'hcs\hcsdiag.exe'
    $script:ItemReadFault = $script:FakeHcsDiagPath
    try { Invoke-Step8 } finally { $script:ClaudeEnv.HcsDiagPath = 'Invoke-FakeHcs' }
    Assert-Case "an hcsdiag.exe that cannot be read: a Warning that names it and why, not 'not available', and nothing raised" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq "$($script:FakeHcsDiagPath) could not be read (Access is denied), so the Cowork VM's worker cannot be told from other VMs' and was not raised" -and
        (Get-Priority 10) -eq 'Normal' -and -not (Test-LogHas 'not available')) "$(Get-StepStatus): $(Get-StepNotes)"
}

# =====================================================================
Invoke-Section 'Setup step 9: vmcompute recovery' {
    $svcKey = 'HKLM:\SYSTEM\CurrentControlSet\Services\vmcompute'
    Reset-State
    Invoke-Step9
    Assert-Case "no vmcompute: Skipped" ((Get-StepStatus) -eq 'Skipped')

    Reset-State
    $script:Services['vmcompute'] = 'Running'
    $null = Add-FakeRegKey $svcKey
    Invoke-Step9
    Assert-Case "no recovery set: set, and recorded as none" ((Get-StepStatus) -eq 'Done' -and $null -ne (Find-Undo 'ServiceFailureActions' 'vmcompute') -and
        $null -eq (Find-Undo 'ServiceFailureActions' 'vmcompute').Original -and
        ($script:ScCalls -join ',') -eq 'failure vmcompute reset= 300 actions= restart/30000/restart/60000/restart/120000') ($script:ScCalls -join ',')
    $script:StepResults = New-Object System.Collections.ArrayList
    $script:ScCalls = New-Object System.Collections.ArrayList
    Invoke-Step9
    Assert-Case "already set: Unchanged, sc.exe not run" ((Get-StepStatus) -eq 'Unchanged' -and $script:ScCalls.Count -eq 0)

    Reset-State
    $script:Services['vmcompute'] = 'Running'
    $other = New-FailureBlob -Reset 86400 -Pairs @(1, 60000)
    Set-FakeRegValue $svcKey 'FailureActions' $other
    Invoke-Step9
    Assert-Case "another recovery: recorded as it was" ((Find-Undo 'ServiceFailureActions' 'vmcompute').Original -eq [Convert]::ToBase64String($other) -and (Get-StepStatus) -eq 'Done')

    Reset-State
    $script:Services['vmcompute'] = 'Running'
    $null = Add-FakeRegKey $svcKey
    $script:ScFails = $true
    Invoke-Step9
    Assert-Case "sc.exe refuses: Failed, saying so" ((Get-StepStatus) -eq 'Failed' -and (Get-StepNotes) -match 'sc.exe returned 5') (Get-StepNotes)
}

# =====================================================================
Invoke-Section 'Setup step 10: stale Cowork VMs' {
    Reset-State
    [void]$script:HcsLive.Add($guidA)
    Invoke-Step10
    Assert-Case "one stale VM is killed and confirmed gone" (($script:Killed -join ',') -eq $guidA -and (Get-StepStatus) -eq 'Done') "$($script:Killed -join ','); $(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    [void]$script:HcsLive.Add($guidA)
    [void]$script:HcsLive.Add($guidB)
    Invoke-Step10
    Assert-Case "two: both killed, WSL and the other VM never" ((@($script:Killed) | Sort-Object) -join ',' -eq (@($guidA, $guidB) | Sort-Object) -join ',' -and
        $script:Killed -notcontains $guidWsl -and $script:Killed -notcontains $guidVm -and (Get-StepStatus) -eq 'Done') ($script:Killed -join ',')

    Reset-State
    [void]$script:HcsLive.Add($guidA)
    $script:HcsStuck = @($guidA)
    Invoke-Step10
    Assert-Case "a VM still listed after the kill: Failed" ((Get-StepStatus) -eq 'Failed' -and (Get-StepNotes) -match "$guidA is still listed after hcsdiag kill") (Get-StepNotes)

    Reset-State
    [void]$script:HcsLive.Add($guidA)
    $script:HcsKillCode = 5
    Invoke-Step10
    Assert-Case "a kill that fails: Failed, naming it" ((Get-StepStatus) -eq 'Failed' -and (Get-StepNotes) -match "hcsdiag kill returned 5 for $guidA") (Get-StepNotes)

    Reset-State
    [void]$script:HcsLive.Add($guidA)
    $script:HcsListDownAfterKill = $true
    Invoke-Step10
    Assert-Case "the list cannot be read again: a warning, not Done" ((Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -match "Could not confirm $guidA is gone") (Get-StepNotes)

    Reset-State
    [void]$script:HcsLive.Add($guidA)
    $script:Procs = @((New-FakeProcess 500 'claude'))
    Invoke-Step10
    Assert-Case "Claude running: its VM is left alone" ($script:Killed.Count -eq 0 -and (Get-StepStatus) -eq 'Skipped')

    Reset-State
    Invoke-Step10
    Assert-Case "no Cowork VM: Unchanged" ((Get-StepStatus) -eq 'Unchanged' -and (Test-LogHas 'HCS state clean'))

    Reset-State
    [void]$script:HcsLive.Add($guidA)
    $script:HcsListCode = 1
    Invoke-Step10
    Assert-Case "a list that fails is not an empty one: Failed, nothing killed" ((Get-StepStatus) -eq 'Failed' -and $script:Killed.Count -eq 0 -and
        (Get-StepNotes) -match 'hcsdiag list returned 1') (Get-StepNotes)

    Reset-State
    $script:ClaudeEnv.HcsDiagPath = $null
    try { Invoke-Step10 } finally { $script:ClaudeEnv.HcsDiagPath = 'Invoke-FakeHcs' }
    Assert-Case "no hcsdiag: Skipped" ((Get-StepStatus) -eq 'Skipped')

    # 6.3.1: as in step 8, an hcsdiag.exe that could not be read is not one
    # that is not there.
    Reset-State
    [void]$script:HcsLive.Add($guidA)
    $script:ClaudeEnv.HcsDiagPath = $null
    $script:FakeHcsDiagPath = Join-Path $sandbox 'hcs\hcsdiag.exe'
    $script:ItemReadFault = $script:FakeHcsDiagPath
    try { Invoke-Step10 } finally { $script:ClaudeEnv.HcsDiagPath = 'Invoke-FakeHcs' }
    Assert-Case "an hcsdiag.exe that cannot be read: a Warning that names it and why, not 'not available', and nothing killed" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq "$($script:FakeHcsDiagPath) could not be read (Access is denied), so stale Cowork VMs were not looked for" -and
        $script:Killed.Count -eq 0 -and -not (Test-LogHas 'not available')) "$(Get-StepStatus): $(Get-StepNotes)"
}

# =====================================================================
Invoke-Section 'Setup step 12: the clock' {
    $chart = @('Tracking time.windows.com [20.101.57.9:123].', 'Collecting 3 samples.', 'The current time is 21-11-2026 20:22:22.')
    Reset-State
    Invoke-Step12
    Assert-Case "no Windows Time service: a warning" ((Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -match 'not installed')

    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:DnsOkAfter = 0
    $script:ChartLines = $chart + @('20:22:22, -12,5000000s')
    Invoke-Step12
    Assert-Case "no network: waits 30 seconds, then a warning, and w32tm is not asked" ((Get-StepStatus) -eq 'Warning' -and $script:FakeClockMs -eq 30000 -and
        $script:DnsTries -eq 15 -and $script:W32Calls.Count -eq 0 -and
        (Get-StepNotes) -match 'No network: time\.windows\.com could not be reached within 30 seconds, so the clock was not checked') "clock $($script:FakeClockMs) ms, tried $($script:DnsTries); $(Get-StepNotes)"

    # The 30 seconds bound the lookups too: one that never answers, or
    # lookups that each take seconds, still end the wait at 30 seconds.
    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:DnsHang = $true
    $script:ChartLines = $chart + @('20:22:22, -12,5000000s')
    Invoke-Step12
    Assert-Case "a lookup that never answers: the wait still ends at 30 seconds" ((Get-StepStatus) -eq 'Warning' -and $script:FakeClockMs -eq 30000 -and
        $script:DnsTries -eq 1 -and $script:W32Calls.Count -eq 0) "clock $($script:FakeClockMs) ms, tried $($script:DnsTries)"

    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:DnsOkAfter = 0
    $script:DnsLatencyMs = 10000
    $script:ChartLines = $chart + @('20:22:22, -12,5000000s')
    Invoke-Step12
    Assert-Case "lookups that take 10 seconds each: the wait still ends at 30 seconds" ((Get-StepStatus) -eq 'Warning' -and $script:FakeClockMs -eq 30000 -and
        $script:DnsTries -eq 3 -and $script:W32Calls.Count -eq 0) "clock $($script:FakeClockMs) ms, tried $($script:DnsTries)"

    # The sleep between tries is cut to the time left. Lookups of 1.5 seconds
    # and sleeps of 2 leave half a second after the ninth try; a whole sleep
    # there would end the wait at 31.5 seconds.
    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:DnsOkAfter = 0
    $script:DnsLatencyMs = 1500
    $script:ChartLines = $chart + @('20:22:22, -12,5000000s')
    Invoke-Step12
    Assert-Case "lookups of 1.5 seconds: the last sleep is cut to the time left, so the wait ends at 30 seconds" ((Get-StepStatus) -eq 'Warning' -and
        $script:FakeClockMs -eq 30000 -and $script:DnsTries -eq 9 -and $script:W32Calls.Count -eq 0) "clock $($script:FakeClockMs) ms, tried $($script:DnsTries)"

    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:DnsOkAfter = 3
    $script:ChartLines = $chart + @('20:22:22, -12,5000000s')
    Invoke-Step12
    Assert-Case "the network comes up: measured, in a comma locale, and resynced" ((Get-StepStatus) -eq 'Done' -and $script:Slept -eq 4 -and
        @($script:W32Calls | Where-Object { $_ -eq '/resync /force' }).Count -eq 1 -and (Test-LogHas '12\.5 s off and has been resynced')) "$(Get-StepStatus), slept $($script:Slept)"

    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:ChartLines = $chart + @('20:22:22, -12,5000000s')
    $script:ResyncCode = 1
    Invoke-Step12
    Assert-Case "a resync that fails is reported" ((Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -match 'w32tm /resync failed') (Get-StepNotes)

    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:ChartLines = $chart + @('20:22:22, +00.0684055s', '20:22:24, +00.0701000s')
    Invoke-Step12
    Assert-Case "a small offset: Unchanged, no resync" ((Get-StepStatus) -eq 'Unchanged' -and @($script:W32Calls | Where-Object { $_ -like '/resync*' }).Count -eq 0)

    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:ChartLines = $chart + @('20:22:22, error: 0x800705B4')
    Invoke-Step12
    Assert-Case "no offset measured: a warning" ((Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -match 'could not measure') (Get-StepNotes)

    Reset-State
    $script:Services['W32Time'] = 'Stopped'
    $script:ChartLines = $chart + @('20:22:22, +00.0684055s')
    Invoke-Step12
    Assert-Case "a stopped service is started" ((Get-StepStatus) -eq 'Done' -and ($script:SvcCalls -join ',') -eq 'Start-Service W32Time')

    Reset-State
    $script:Services['W32Time'] = 'Stopped'
    $script:StartServiceThrows = $true
    $script:ChartLines = $chart + @('20:22:22, +00.0684055s')
    Invoke-Step12
    Assert-Case "a service that will not start: a warning" ((Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -match 'Could not start the Windows Time service') (Get-StepNotes)
}

# =====================================================================
Invoke-Section 'The monitor: worker priority and the clock' {
    Reset-State
    Set-VmFixture
    Set-VmWorkerPriority
    Assert-Case "the monitor raises the Cowork VM's worker" ((Get-Priority 10) -eq 'AboveNormal' -and
        @($script:WatchLog | Where-Object { $_ -match 'vmwp\.exe PID 10\) to AboveNormal' }).Count -eq 1)
    Assert-Case "and no other VM's" ((Get-Priority 11) -eq 'Normal' -and (Get-Priority 12) -eq 'Normal')
    $script:WatchLog = New-Object System.Collections.ArrayList
    Set-VmWorkerPriority
    Assert-Case "already raised: nothing logged" ($script:WatchLog.Count -eq 0)

    Reset-State
    Set-VmFixture
    $script:HcsLive = New-Object System.Collections.ArrayList
    Set-VmWorkerPriority
    Assert-Case "no Cowork VM: the process list is not even read" (@($script:CimCalls | Where-Object { $_ -eq 'Win32_Process' }).Count -eq 0 -and (Get-Priority 11) -eq 'Normal')

    Reset-State
    Set-VmFixture
    $script:HcsListDown = $true
    Set-VmWorkerPriority
    Assert-Case "hcsdiag not answering: nothing raised" ((Get-Priority 10) -eq 'Normal')

    # The clock: Test-TimeSyncHealth runs w32tm through the lifted
    # Invoke-W32tm, whose job is the fake Start-Job. $script:ChartSeq holds
    # the charts the measurements take in turn, before $script:ChartLines.
    $chart = @('Tracking time.windows.com [20.101.57.9:123].', 'Collecting 1 samples.', 'The current time is 21-11-2026 20:22:22.')
    $held = 'time source: time\.windows\.com,0x9\. No resync is forced for 60 min while the drift lasts; it is still measured every 5 min$'
    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:ChartLines = $chart + @('20:22:22, +00.0684055s')
    $r = Test-TimeSyncHealth
    Assert-Case "a small offset: no problem, no resync" ($null -eq $r -and @($script:W32Calls | Where-Object { $_ -like '/resync*' }).Count -eq 0) "got $r"

    Reset-State
    $script:Services['W32Time'] = 'Running'
    [void]$script:ChartSeq.Add([object[]]($chart + @('20:22:22, -12,5000000s')))
    $script:ChartLines = $chart + @('20:22:22, +00,0700000s')
    $r = Test-TimeSyncHealth
    Assert-Case "a comma locale is measured, resynced, measured again, and only then logged as repaired" ($null -eq $r -and
        @($script:W32Calls | Where-Object { $_ -eq '/resync /force' }).Count -eq 1 -and @($script:W32Calls | Where-Object { $_ -like '/stripchart*' }).Count -eq 2 -and
        ($script:WatchLog -join ' | ') -eq 'REPAIRED: Forced NTP resync (drift was 12.5s, now 0.07s)' -and
        $script:ResyncBackoffUntil -eq [datetime]::MinValue) "got $r; $($script:WatchLog -join ' | ')"

    Reset-State
    $script:Services['W32Time'] = 'Running'
    # The clock would read in range if it were measured again: a resync that
    # failed is reported as failed all the same, and not measured again.
    [void]$script:ChartSeq.Add([object[]]($chart + @('20:22:22, -12,5000000s')))
    $script:ChartLines = $chart + @('20:22:22, +00,0700000s')
    $script:ResyncCode = 1
    $r = Test-TimeSyncHealth
    Assert-Case "a resync that fails: reported with the time source, not logged as repaired" ($r -eq ('Clock drift 12.5s and w32tm /resync failed (exit 1); ' +
        'time source: time.windows.com,0x9. No resync is forced for 60 min while the drift lasts; it is still measured every 5 min') -and
        @($script:WatchLog | Where-Object { $_ -match 'REPAIRED' }).Count -eq 0 -and
        @($script:W32Calls | Where-Object { $_ -like '/stripchart*' }).Count -eq 1) "got $r; $($script:WatchLog -join ' | ')"
    Assert-Case "and no resync is forced for an hour" ($script:ResyncBackoffUntil -gt [datetime]::UtcNow.AddMinutes(59) -and
        $script:ResyncBackoffUntil -lt [datetime]::UtcNow.AddMinutes(61)) "until $($script:ResyncBackoffUntil.ToString('o'))"

    # The hour runs from the time after the resync. Until 6.3.1 it ran from
    # the time read before it, so a resync that set the clock forward left
    # a hold that had already ended, or nearly. Here the resync takes 400 ms
    # of real time, which the hold must not lose.
    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:ChartLines = $chart + @('20:22:22, -12,5000000s')
    $script:ResyncCode = 1
    $script:ResyncTakesMs = 400
    $null = Test-TimeSyncHealth
    Assert-Case "the hour of no forced resync runs from the time after the resync, not before it" ($script:ResyncDoneAt -ne [datetime]::MinValue -and
        $script:ResyncBackoffUntil -ge $script:ResyncDoneAt.AddMinutes(60) -and $script:ResyncBackoffUntil -lt [datetime]::UtcNow.AddMinutes(61)) `
        "until $($script:ResyncBackoffUntil.ToString('o')), resync done $($script:ResyncDoneAt.ToString('o'))"

    # The Windows Time service that cannot be read: said once an hour,
    # not started, and the clock still measured. Until 6.3.1 it was read
    # with SilentlyContinue, as a service that is not installed, unsaid.
    $w32Line = 'TIME WARNING: The Windows Time service could not be read (Access is denied), so it is not started if it is stopped; ' +
        'the clock is still measured every 5 min; the same error is logged at most once an hour'
    Reset-State
    $script:Services['W32Time'] = 'Stopped'
    $script:ServiceReadFault = @('W32Time')
    $script:ChartLines = $chart + @('20:22:22, +00.0684055s')
    $r = Test-TimeSyncHealth
    $script:LastTimeSyncCheck = [datetime]::UtcNow.AddMinutes(-6)
    $null = Test-TimeSyncHealth
    Assert-Case "the Windows Time service cannot be read: said once, not started, and the clock still measured" ($null -eq $r -and
        ($script:WatchLog -join ' | ') -eq $w32Line -and $script:SvcCalls.Count -eq 0 -and
        @($script:W32Calls | Where-Object { $_ -like '/stripchart*' }).Count -eq 2) "got $r; $($script:WatchLog -join ' | '); $($script:SvcCalls -join ',')"
    Reset-State
    $script:ServiceReadFault = @()
    $script:ChartLines = $chart + @('20:22:22, +00.0684055s')
    $r = Test-TimeSyncHealth
    Assert-Case "a Windows Time service that is not installed: not started, and nothing said" ($null -eq $r -and $script:WatchLog.Count -eq 0 -and
        $script:SvcCalls.Count -eq 0 -and @($script:W32Calls | Where-Object { $_ -like '/stripchart*' }).Count -eq 1) "got $r; $($script:WatchLog -join ' | ')"

    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:ChartLines = $chart + @('20:22:22, -12,5000000s')
    $r = Test-TimeSyncHealth
    Assert-Case "a resync that returns 0 and changes nothing: reported, not logged as repaired" ($r -match ('^Clock drift 12\.5s and it is still 12\.5s after w32tm /resync; ' + $held) -and
        @($script:WatchLog | Where-Object { $_ -match 'REPAIRED' }).Count -eq 0 -and @($script:W32Calls | Where-Object { $_ -like '/stripchart*' }).Count -eq 2) "got $r; $($script:WatchLog -join ' | ')"

    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:ResyncBackoffUntil = [datetime]::UtcNow.AddMinutes(30)
    $script:ChartLines = $chart + @('20:22:22, -12,5000000s')
    $r = Test-TimeSyncHealth
    Assert-Case "while the hold lasts: measured, not resynced, not reported again" ($null -eq $r -and
        @($script:W32Calls | Where-Object { $_ -like '/resync*' }).Count -eq 0 -and @($script:W32Calls | Where-Object { $_ -like '/stripchart*' }).Count -eq 1) "got $r"

    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:ResyncBackoffUntil = [datetime]::UtcNow.AddMinutes(30)
    $script:ChartLines = $chart + @('20:22:22, +45.0000000s')
    $r = Test-TimeSyncHealth
    Assert-Case "while the hold lasts, a drift past 30 seconds is reported" ($r -eq 'Clock drift 45s, may cause VM connectivity issues' -and
        @($script:W32Calls | Where-Object { $_ -like '/resync*' }).Count -eq 0) "got $r"

    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:ResyncBackoffUntil = [datetime]::UtcNow.AddMinutes(30)
    $script:ChartLines = $chart + @('20:22:22, +00.0700000s')
    $r = Test-TimeSyncHealth
    Assert-Case "the clock back within 5 seconds ends the hold" ($null -eq $r -and $script:ResyncBackoffUntil -eq [datetime]::MinValue -and
        ($script:WatchLog -join ' | ') -eq 'Clock drift back to 0.07s; forced resyncs allowed again') "got $r; $($script:WatchLog -join ' | ')"

    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:JobHangs = $true
    $r = Test-TimeSyncHealth
    Assert-Case "w32tm hanging: the job is stopped and the check skipped" ($null -eq $r -and @($script:Calls | Where-Object { $_ -eq 'Stop-Job' }).Count -eq 1 -and
        ($script:WatchLog -join ' | ') -eq 'w32tm stripchart did not answer within 15s; skipping the clock-drift check') ($script:WatchLog -join ' | ')

    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:ChartLines = $chart + @('20:22:22, -12,5000000s')
    $script:JobHangMatch = '/resync'
    $r = Test-TimeSyncHealth
    Assert-Case "a resync that does not answer is stopped and reported" ($r -match ('^Clock drift 12\.5s and w32tm /resync did not answer within 15s; ' + $held) -and
        @($script:Calls | Where-Object { $_ -eq 'Stop-Job' }).Count -eq 1) "got $r"

    Reset-State
    $script:Services['W32Time'] = 'Running'
    [void]$script:ChartSeq.Add([object[]]($chart + @('20:22:22, -12,5000000s')))
    $script:ChartLines = $chart + @('20:22:22, error: 0x800705B4')
    $r = Test-TimeSyncHealth
    Assert-Case "a clock that cannot be measured after the resync is reported, naming why, not logged as repaired" ($r -match ('^Clock drift 12\.5s and the clock could not be measured after w32tm /resync \(w32tm /stripchart gave no offset: error: 0x800705B4\); ' + $held) -and
        @($script:WatchLog | Where-Object { $_ -match 'REPAIRED' }).Count -eq 0) "got $r"

    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:LastTimeSyncCheck = [datetime]::UtcNow
    $script:ChartLines = $chart + @('20:22:22, -12,5000000s')
    $r = Test-TimeSyncHealth
    Assert-Case "within five minutes of the last check: not asked again" ($null -eq $r -and $script:W32Calls.Count -eq 0)

    Reset-State
    $script:Services['W32Time'] = 'Stopped'
    $script:ChartLines = $chart + @('20:22:22, +00.0684055s')
    $null = Test-TimeSyncHealth
    Assert-Case "a stopped Windows Time service is started" (($script:SvcCalls -join ',') -eq 'Start-Service W32Time' -and
        @($script:WatchLog | Where-Object { $_ -eq 'REPAIRED: Started W32Time service' }).Count -eq 1)

    # A clock check that fails part-way, here because Start-Job throws:
    # said in one line naming the error, and still $null, so the poll goes
    # on. The same error at most once an hour, another at once. Until 6.3.0
    # nothing was said, and a check that could not run looked like a clock
    # that was right. Each next check is let through the 5-minute gate by
    # moving the last one back.
    $jobErr = 'The background job could not be started: Access is denied.'
    $tail = ', so whether the clock is right is not known. It runs again every 5 min; the same error is logged at most once an hour'
    $failLine = 'TIME WARNING: The clock check failed (w32tm /stripchart could not be run: ' + $jobErr + ')' + $tail
    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:JobThrows = $jobErr
    $r = Test-TimeSyncHealth
    Assert-Case "a clock check that fails: one line naming the error, and nothing returned" ($null -eq $r -and
        ($script:WatchLog -join ' | ') -eq $failLine) "got $r; $($script:WatchLog -join ' | ')"
    $script:LastTimeSyncCheck = [datetime]::UtcNow.AddMinutes(-6)
    $r = Test-TimeSyncHealth
    Assert-Case "and the same error at the next check, within the hour: not logged again" ($null -eq $r -and $script:WatchLog.Count -eq 1 -and
        @($script:Calls | Where-Object { $_ -eq 'Start-Job failed' }).Count -eq 2) "got $r; $($script:WatchLog -join ' | ')"
    # Another error is said at once, on one line.
    $script:LastTimeSyncCheck = [datetime]::UtcNow.AddMinutes(-6)
    $script:JobThrows = "The background job could not be started.`r`n  The paging file is too small for this operation to complete."
    $r = Test-TimeSyncHealth
    $otherLine = 'TIME WARNING: The clock check failed (w32tm /stripchart could not be run: The background job could not be started. The paging file is too small for this operation to complete.)' + $tail
    Assert-Case "another error within the hour is logged at once, on one line" ($null -eq $r -and $script:WatchLog.Count -eq 2 -and
        $script:WatchLog[1] -eq $otherLine) "got $r; $($script:WatchLog -join ' | ')"
    # An hour on, the first error is logged again.
    foreach ($k in @($script:CheckErrorsSaid.Keys)) { $script:CheckErrorsSaid[$k] = $script:CheckErrorsSaid[$k].AddMinutes(-60) }
    $script:LastTimeSyncCheck = [datetime]::UtcNow.AddMinutes(-6)
    $script:JobThrows = $jobErr
    $r = Test-TimeSyncHealth
    Assert-Case "an hour on, the same error is logged again" ($null -eq $r -and $script:WatchLog.Count -eq 3 -and $script:WatchLog[2] -eq $failLine) `
        "got $r; $($script:WatchLog -join ' | ')"
    # A time logged that is in the future, as after the clock was set back,
    # counts as old: logged again rather than held off.
    foreach ($k in @($script:CheckErrorsSaid.Keys)) { $script:CheckErrorsSaid[$k] = [datetime]::UtcNow.AddMinutes(30) }
    $script:LastTimeSyncCheck = [datetime]::UtcNow.AddMinutes(-6)
    $r = Test-TimeSyncHealth
    Assert-Case "after the clock was set back, the same error is logged again" ($null -eq $r -and $script:WatchLog.Count -eq 4 -and
        $script:WatchLog[3] -eq $failLine) "got $r; $($script:WatchLog -join ' | ')"
    # An error with no message is named by its type.
    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:JobThrows = [System.InvalidOperationException]::new('')
    $r = Test-TimeSyncHealth
    Assert-Case "an error with no message: named by its type" ($null -eq $r -and
        ($script:WatchLog -join ' | ') -eq ('TIME WARNING: The clock check failed (w32tm /stripchart could not be run: System.InvalidOperationException)' + $tail)) "got $r; $($script:WatchLog -join ' | ')"

    # The real w32tm.exe is never reached: the lifted copy runs the fake by
    # name, and a path given to it trips the fake Start-Job before any job.
    $iw = (Get-Command Invoke-W32tm -CommandType Function).ScriptBlock.Ast
    $pathParam = @($iw.FindAll({ param($n) $n -is [System.Management.Automation.Language.ParameterAst] -and $n.Name.VariablePath.UserPath -eq 'W32tmPath' }, $true))
    $dflt = '<none>'
    if ($pathParam.Count -eq 1 -and $null -ne $pathParam[0].DefaultValue) { $dflt = $pathParam[0].DefaultValue.Extent.Text }
    Assert-Case "the monitor's w32tm runs the fake, by name" ($dflt -eq "'w32tm'") "default $dflt"
    $before = $script:Tripped.Count
    $threw = ''
    try { $null = Invoke-W32tm -W32tmPath 'C:\Windows\System32\w32tm.exe' -Arguments '/query', '/source' } catch { $threw = $_.Exception.Message }
    $trip = 'Start-Job C:\Windows\System32\w32tm.exe /query /source'
    Assert-Case "and w32tm.exe by its path trips before any job runs" ($threw -match '^(w32tm /query could not be run: )?TRIPWIRE: Start-Job ' -and $script:Tripped.Count -eq $before + 1 -and
        $script:Tripped[$script:Tripped.Count - 1] -eq $trip) "threw '$threw'"
    if ($script:Tripped.Count -gt $before -and $script:Tripped[$script:Tripped.Count - 1] -eq $trip) { $script:Tripped.RemoveAt($script:Tripped.Count - 1) }
}

# =====================================================================
# The clock check when w32tm runs and fails, rather than not answering: a
# job that ends Failed, one that ends with no exit code, and w32tm that
# answers with an error and no offset. Each is a TIME WARNING naming the
# error, at most once an hour for the same error. Until 6.3.0 the first two
# were logged as a check that did not answer, on every check, and the
# third was not logged at all.
Invoke-Section 'The monitor: a clock check that fails part-way' {
    $chart = @('Tracking time.windows.com [20.101.57.9:123].', 'Collecting 1 samples.', 'The current time is 21-11-2026 20:22:22.')
    $tail = ', so whether the clock is right is not known. It runs again every 5 min; the same error is logged at most once an hour'
    $held = 'time source: time\.windows\.com,0x9\. No resync is forced for 60 min while the drift lasts; it is still measured every 5 min$'
    $jobFail = 'An error occurred while starting the background process. Error reported: The directory name is invalid.'

    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:JobFails = $jobFail
    $r = Test-TimeSyncHealth
    $line = 'TIME WARNING: The clock check failed (w32tm /stripchart could not be run: ' + $jobFail + ')' + $tail
    Assert-Case "a job that ends Failed: a TIME WARNING naming its error, not a check that did not answer" ($null -eq $r -and
        ($script:WatchLog -join ' | ') -eq $line) "got $r; $($script:WatchLog -join ' | ')"
    $script:LastTimeSyncCheck = [datetime]::UtcNow.AddMinutes(-6)
    $r = Test-TimeSyncHealth
    Assert-Case "and the same at the next check, within the hour: not logged again" ($null -eq $r -and $script:WatchLog.Count -eq 1 -and
        @($script:Calls | Where-Object { $_ -eq 'Start-Job Failed' }).Count -eq 2) "got $r; $($script:WatchLog -join ' | ')"

    Reset-State
    $script:Services['W32Time'] = 'Running'
    $noTerm = "The term 'w32tm' is not recognized as the name of a cmdlet, function, script file, or operable program."
    $script:JobNoExit = $noTerm
    $r = Test-TimeSyncHealth
    Assert-Case "a job that ends with no exit code: a TIME WARNING naming its error" ($null -eq $r -and
        ($script:WatchLog -join ' | ') -eq ('TIME WARNING: The clock check failed (w32tm /stripchart could not be run: ' + $noTerm + ')' + $tail)) "got $r; $($script:WatchLog -join ' | ')"

    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:ChartLines = $chart + @('20:22:22, error: 0x800705B4')
    $r = Test-TimeSyncHealth
    $line = 'TIME WARNING: The clock check failed (w32tm /stripchart gave no offset: error: 0x800705B4)' + $tail
    Assert-Case "w32tm that answers with an error and no offset: a TIME WARNING naming the error, and no resync" ($null -eq $r -and
        ($script:WatchLog -join ' | ') -eq $line -and @($script:W32Calls | Where-Object { $_ -like '/resync*' }).Count -eq 0) "got $r; $($script:WatchLog -join ' | ')"
    $script:LastTimeSyncCheck = [datetime]::UtcNow.AddMinutes(-6)
    $r = Test-TimeSyncHealth
    Assert-Case "and the same at the next check, within the hour: not logged again" ($null -eq $r -and $script:WatchLog.Count -eq 1) "got $r; $($script:WatchLog -join ' | ')"

    # A measurement that does not answer is still the line it always was,
    # now at most once an hour: the check runs every 5 minutes.
    $toLine = 'w32tm stripchart did not answer within 15s; skipping the clock-drift check'
    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:JobHangs = $true
    $null = Test-TimeSyncHealth
    $script:LastTimeSyncCheck = [datetime]::UtcNow.AddMinutes(-6)
    $null = Test-TimeSyncHealth
    Assert-Case "a measurement that does not answer at two checks in the hour: said once" (($script:WatchLog -join ' | ') -eq $toLine -and
        @($script:Calls | Where-Object { $_ -eq 'Stop-Job' }).Count -eq 2) ($script:WatchLog -join ' | ')
    foreach ($k in @($script:CheckErrorsSaid.Keys)) { $script:CheckErrorsSaid[$k] = $script:CheckErrorsSaid[$k].AddMinutes(-60) }
    $script:LastTimeSyncCheck = [datetime]::UtcNow.AddMinutes(-6)
    $null = Test-TimeSyncHealth
    Assert-Case "and said again an hour on" ($script:WatchLog.Count -eq 2 -and $script:WatchLog[1] -eq $toLine) ($script:WatchLog -join ' | ')

    # A drift found is reported, whatever fails after it.
    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:ChartLines = $chart + @('20:22:22, -12,5000000s')
    $script:JobFails = $jobFail
    $script:JobFailMatch = '/resync'
    $r = Test-TimeSyncHealth
    Assert-Case "a resync whose job ends Failed: the drift is still reported, naming the error" ($r -eq ('Clock drift 12.5s and w32tm /resync could not be run: ' + $jobFail +
        '; time source: time.windows.com,0x9. No resync is forced for 60 min while the drift lasts; it is still measured every 5 min') -and
        $script:WatchLog.Count -eq 0) "got $r; $($script:WatchLog -join ' | ')"

    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:ChartLines = $chart + @('20:22:22, -12,5000000s')
    $script:JobFails = $jobFail
    $script:JobFailMatch = '/stripchart'
    $script:JobFailSkip = 1
    $r = Test-TimeSyncHealth
    Assert-Case "a measurement after the resync whose job ends Failed: the drift is still reported, naming the error" ($r -match ('^Clock drift 12\.5s and the clock could not be measured after w32tm /resync \(w32tm /stripchart could not be run: ' +
        [regex]::Escape($jobFail) + '\); ' + $held) -and @($script:WatchLog | Where-Object { $_ -match 'REPAIRED|did not answer' }).Count -eq 0) "got $r; $($script:WatchLog -join ' | ')"

    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:ChartLines = $chart + @('20:22:22, -12,5000000s')
    $script:JobFails = $jobFail
    $script:JobFailMatch = '/query'
    $r = Test-TimeSyncHealth
    Assert-Case "a time source that cannot be asked: the drift is still reported, and why the source is unknown" ($r -match ('^Clock drift 12\.5s and it is still 12\.5s after w32tm /resync; time source: unknown \(w32tm /query could not be run: ' +
        [regex]::Escape($jobFail) + '\)\. No resync is forced for 60 min')) "got $r"

    # A hold more than 60 minutes ahead cannot have been set by this check,
    # which sets 60: the clock was set back since. It is dropped. Until
    # 6.3.0 it kept resyncs off for as long again.
    Reset-State
    $script:Services['W32Time'] = 'Running'
    [void]$script:ChartSeq.Add([object[]]($chart + @('20:22:22, -12,5000000s')))
    $script:ChartLines = $chart + @('20:22:22, +00,0700000s')
    $script:ResyncBackoffUntil = [datetime]::UtcNow.AddMinutes(150)
    $r = Test-TimeSyncHealth
    Assert-Case "a hold 150 min ahead, after the clock was set back: dropped, and the drift resynced" ($null -eq $r -and
        @($script:W32Calls | Where-Object { $_ -eq '/resync /force' }).Count -eq 1 -and
        ($script:WatchLog -join ' | ') -eq 'REPAIRED: Forced NTP resync (drift was 12.5s, now 0.07s)' -and
        $script:ResyncBackoffUntil -eq [datetime]::MinValue) "got $r; $($script:WatchLog -join ' | ')"
    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:ChartLines = $chart + @('20:22:22, -12,5000000s')
    $script:ResyncBackoffUntil = [datetime]::UtcNow.AddMinutes(59)
    $r = Test-TimeSyncHealth
    Assert-Case "a hold 59 min ahead holds: no resync" ($null -eq $r -and @($script:W32Calls | Where-Object { $_ -like '/resync*' }).Count -eq 0 -and
        $script:ResyncBackoffUntil -gt [datetime]::UtcNow.AddMinutes(58)) "got $r"
}

# =====================================================================
# The throttle every check that fails logs through: Test-CheckErrorDue,
# Write-CheckError and Get-CheckErrorText.
Invoke-Section 'The monitor: a check that fails is logged at most once an hour' {
    $t0 = [datetime]::UtcNow
    Reset-State
    $a = Test-CheckErrorDue -Check 'storage' -Message 'Access is denied.' -Now $t0
    $b = Test-CheckErrorDue -Check 'storage' -Message 'Access is denied.' -Now $t0.AddMinutes(59)
    $c = Test-CheckErrorDue -Check 'storage' -Message 'The device is not ready.' -Now $t0.AddMinutes(1)
    $d = Test-CheckErrorDue -Check 'clock' -Message 'Access is denied.' -Now $t0.AddMinutes(1)
    Assert-Case "the first error of a check is said, and the same again within the hour is not" ($a -and -not $b) "first $a, again $b"
    Assert-Case "another error of the check, or the same error of another check, is said at once" ($c -and $d) "another $c, other check $d"
    $e = Test-CheckErrorDue -Check 'storage' -Message 'Access is denied.' -Now $t0.AddMinutes(60)
    Assert-Case "an hour on, the same error is said again" $e

    Reset-State
    $null = Test-CheckErrorDue -Check 'storage' -Message 'Access is denied.' -Now $t0
    $f = Test-CheckErrorDue -Check 'storage' -Message 'Access is denied.' -Now $t0.AddMinutes(-5)
    Assert-Case "a time said that is ahead of now, after the clock was set back, counts as old" $f

    Reset-State
    $said = @(1..4 | ForEach-Object { Test-CheckErrorDue -Check 'storage' -Message "error $_" -Now $t0.AddMinutes($_) })
    Assert-Case "no more than 3 errors of one check in an hour" (($said -join ',') -eq 'True,True,True,False') ($said -join ',')
    $g = Test-CheckErrorDue -Check 'clock' -Message 'error 4' -Now $t0.AddMinutes(5)
    Assert-Case "which leaves the other checks alone" $g

    Reset-State
    $null = Test-CheckErrorDue -Check 'storage' -Message 'old' -Now $t0
    $null = Test-CheckErrorDue -Check 'clock' -Message 'new' -Now $t0.AddMinutes(61)
    Assert-Case "an entry an hour old is dropped from the table" ((@($script:CheckErrorsSaid.Keys) -join ',') -eq 'clock|new') (@($script:CheckErrorsSaid.Keys) -join ',')

    Reset-State
    $rec = [System.Management.Automation.ErrorRecord]::new([System.IO.IOException]::new("The device is not ready.`r`n  Try again."), 'IOError', 'ReadError', $null)
    Write-CheckError -Check 'storage' -ErrorRecord $rec -Text 'STORAGE WARNING: the read failed ({0})'
    Write-CheckError -Check 'storage' -ErrorRecord $rec -Text 'STORAGE WARNING: the read failed ({0})'
    Assert-Case "Write-CheckError: the line, with the error on one line in it, once" (($script:WatchLog -join ' | ') -eq
        'STORAGE WARNING: the read failed (The device is not ready. Try again.); the same error is logged at most once an hour') ($script:WatchLog -join ' | ')
    $empty = [System.Management.Automation.ErrorRecord]::new([System.InvalidOperationException]::new(''), 'Empty', 'NotSpecified', $null)
    Assert-Case "Get-CheckErrorText: an error with no message is named by its type" ((Get-CheckErrorText -ErrorRecord $empty) -eq 'System.InvalidOperationException')
}

# =====================================================================
# The VM worker processes, listed with Stop. Until 6.3.0 a list that could
# not be read, or a raise that failed, was passed over without a word.
Invoke-Section "The monitor: the VM's worker process, when it cannot be read or raised" {
    Reset-State
    Set-VmFixture
    $script:VmWorkersFault = $true
    Set-VmWorkerPriority
    Set-VmWorkerPriority
    $line = "The VM worker processes (vmwp.exe) could not be listed (The RPC server is unavailable.), so the Cowork VM's worker was not raised to AboveNormal; the same error is logged at most once an hour"
    Assert-Case "a process list that cannot be read: said, once, and nothing raised" (($script:WatchLog -join ' | ') -eq $line -and
        (Get-Priority 10) -eq 'Normal') ($script:WatchLog -join ' | ')

    Reset-State
    Set-VmFixture
    $script:Procs = @($script:Procs | Where-Object { $_.Id -ne 10 })
    Set-VmWorkerPriority
    Assert-Case "a worker that exited since the list was read: passed over without a word" ($script:WatchLog.Count -eq 0) ($script:WatchLog -join ' | ')

    Reset-State
    Set-VmFixture
    $locked = [pscustomobject]@{ Id = 10; ProcessName = 'vmwp' }
    $locked | Add-Member -MemberType ScriptProperty -Name PriorityClass -Value { 'Normal' } -SecondValue { throw 'Access is denied.' }
    $script:Procs = @($locked) + @($script:Procs | Where-Object { $_.Id -ne 10 })
    Set-VmWorkerPriority
    Set-VmWorkerPriority
    $raise = @($script:WatchLog | Where-Object { $_ -match "^The Cowork VM's worker \(vmwp\.exe PID 10\) could not be raised to AboveNormal \(.*Access is denied\..*\); the same error is logged at most once an hour$" })
    Assert-Case "a raise that fails: said once, naming why" ($raise.Count -eq 1 -and $script:WatchLog.Count -eq 1) ($script:WatchLog -join ' | ')
}

# =====================================================================
# The Cowork compute systems left behind beside the live one, which is
# the one whose worker started last. Until 6.3.0 this was part of the main
# loop: it kept the one hcsdiag listed last, logged "killed" for each kill
# it sent, and took a worker list that could not be read for one with no
# Cowork worker in it.
Invoke-Section 'The monitor: Cowork compute systems left behind' {
    $older = [datetime]::Now.AddMinutes(-10)
    $newer = [datetime]::Now.AddMinutes(-1)
    function Set-TwoCoworkVms {
        [void]$script:HcsLive.Add($guidB)
        [void]$script:HcsLive.Add($guidA)
        $script:VmWorkers = @((New-FakeWorker 10 $guidA $older), (New-FakeWorker 14 $guidB $newer), (New-FakeWorker 11 $guidWsl $older))
    }
    Reset-State
    Set-TwoCoworkVms
    $script:LastHcsListTime = [datetime]::UtcNow
    Stop-StaleCoworkVm
    Assert-Case "two Cowork VMs: the one whose worker started first is killed, though hcsdiag lists it last, and confirmed gone" (($script:Killed -join ',') -eq $guidA -and
        ($script:WatchLog -join ' | ') -eq "HCS: killed stale cowork-vm ($guidA)") "killed $($script:Killed -join ','); $($script:WatchLog -join ' | ')"
    Assert-Case "and the cached list is read afresh after it" ($script:LastHcsListTime -eq [datetime]::MinValue)

    Reset-State
    Set-TwoCoworkVms
    $script:VmWorkersFault = $true
    Stop-StaleCoworkVm
    Stop-StaleCoworkVm
    $line = 'HCS: 2 cowork-vm compute systems, and the VM worker processes (vmwp.exe) could not be listed to tell the live one by (The RPC server is unavailable.); none killed; the same error is logged at most once an hour'
    Assert-Case "a worker list that cannot be read: nothing killed, said once, and not as a list with no Cowork worker in it" ($script:Killed.Count -eq 0 -and
        ($script:WatchLog -join ' | ') -eq $line) ($script:WatchLog -join ' | ')

    Reset-State
    Set-TwoCoworkVms
    $script:VmWorkers = @((New-FakeWorker 11 $guidWsl $older))
    Stop-StaleCoworkVm
    Assert-Case "a worker list read with no Cowork worker in it: nothing killed, and said so" ($script:Killed.Count -eq 0 -and
        ($script:WatchLog -join ' | ') -eq 'HCS: 2 cowork-vm compute systems and none has a worker process to tell the live one by; none killed') ($script:WatchLog -join ' | ')

    # A worker that cannot be read is not one that is not there. Until
    # 6.3.1 a worker whose command line or start time could not be read was
    # passed over: when it was the live VM's, the live VM was killed.
    $unreadLine = 'HCS: 2 cowork-vm compute systems, and the command line or start time of 1 VM worker process(es) (vmwp.exe) ' +
        'could not be read to tell the live one by; none killed; the same error is logged at most once an hour'
    Reset-State
    Set-TwoCoworkVms
    $script:VmWorkers = @((New-FakeWorker 10 $guidA $older), [pscustomobject]@{ ProcessId = 14; CommandLine = $null; CreationDate = $newer },
        (New-FakeWorker 11 $guidWsl $older))
    Stop-StaleCoworkVm
    Stop-StaleCoworkVm
    Assert-Case "the live VM's worker whose command line cannot be read: nothing killed, said once" ($script:Killed.Count -eq 0 -and
        ($script:WatchLog -join ' | ') -eq $unreadLine) "killed $($script:Killed -join ','); $($script:WatchLog -join ' | ')"

    Reset-State
    Set-TwoCoworkVms
    $script:VmWorkers = @((New-FakeWorker 10 $guidA $older), (New-FakeWorker 14 $guidB $null), (New-FakeWorker 11 $guidWsl $older))
    Stop-StaleCoworkVm
    Assert-Case "the live VM's worker whose start time cannot be read: nothing killed, and said" ($script:Killed.Count -eq 0 -and
        ($script:WatchLog -join ' | ') -eq $unreadLine) "killed $($script:Killed -join ','); $($script:WatchLog -join ' | ')"

    # A worker of another VM with no start time does not hold the kill back.
    Reset-State
    Set-TwoCoworkVms
    $script:VmWorkers = @((New-FakeWorker 10 $guidA $older), (New-FakeWorker 14 $guidB $newer), (New-FakeWorker 11 $guidWsl $null))
    Stop-StaleCoworkVm
    Assert-Case "another VM's worker with no start time: the stale Cowork VM is still killed" (($script:Killed -join ',') -eq $guidA -and
        ($script:WatchLog -join ' | ') -eq "HCS: killed stale cowork-vm ($guidA)") "killed $($script:Killed -join ','); $($script:WatchLog -join ' | ')"

    Reset-State
    Set-TwoCoworkVms
    $script:HcsStuck = @($guidA)
    Stop-StaleCoworkVm
    Assert-Case "a kill that leaves it listed: said so, not that it was killed" (($script:Killed -join ',') -eq $guidA -and
        ($script:WatchLog -join ' | ') -eq "HCS: stale cowork-vm ($guidA) was still listed after the kill") ($script:WatchLog -join ' | ')

    Reset-State
    [void]$script:HcsLive.Add($guidA)
    $script:VmWorkers = @((New-FakeWorker 10 $guidA $older))
    Stop-StaleCoworkVm
    Assert-Case "one Cowork VM: nothing killed, nothing said, no process list read" ($script:Killed.Count -eq 0 -and $script:WatchLog.Count -eq 0 -and
        @($script:CimCalls | Where-Object { $_ -eq 'Win32_Process' }).Count -eq 0) ($script:WatchLog -join ' | ')
}

# =====================================================================
# hcsdiag list prints nothing at all once no compute system is left, which
# is what killing a lone Cowork VM leaves. Until 6.3.0 that empty answer
# read as "hcsdiag did not answer", and the kill was reported as unconfirmed
# on every machine without WSL or another VM running.
Invoke-Section 'Setup step 10: the only compute system, killed' {
    Reset-State
    $script:HcsOnlyCowork = $true
    [void]$script:HcsLive.Add($guidA)
    Invoke-Step10
    Assert-Case "the lone Cowork VM is killed and confirmed gone: Done" (($script:Killed -join ',') -eq $guidA -and (Get-StepStatus) -eq 'Done' -and
        (Test-LogHas ('^Killed stale HCS compute system: ' + [regex]::Escape($guidA) + '$'))) "$(Get-StepStatus): $(Get-StepNotes)"
    Assert-Case "after the list was read again" (@($script:HcsCalls | Where-Object { $_ -eq 'list' }).Count -ge 2 -and
        -not (Test-LogHas 'could not be read again')) ($script:HcsCalls -join ' / ')

    # Step 10's own -List block, cut out of Prevent: what it hands
    # Get-CoworkHcsSurvivors for an empty list, and for a list that failed.
    $calls = @($pAst.FindAll({
        param($n)
        $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Get-CoworkHcsSurvivors'
    }, $true))
    Assert-Case "Prevent asks Get-CoworkHcsSurvivors in one place" ($calls.Count -eq 1) "found $($calls.Count)"
    $listAst = $null
    if ($calls.Count -eq 1) {
        $els = @($calls[0].CommandElements)
        for ($i = 0; $i -lt $els.Count; $i++) {
            $e = $els[$i]
            if ($e -is [System.Management.Automation.Language.CommandParameterAst] -and $e.ParameterName -eq 'List') {
                $a = $e.Argument
                if ($null -eq $a -and $i + 1 -lt $els.Count) { $a = $els[$i + 1] }
                if ($a -is [System.Management.Automation.Language.ScriptBlockExpressionAst]) { $listAst = $a }
            }
        }
    }
    Assert-Case "its -List is a script block" ($null -ne $listAst)
    if ($null -ne $listAst) {
        $listBlock = $listAst.ScriptBlock.GetScriptBlock()
        $hcsExe = 'Invoke-FakeHcs'
        Reset-State
        $script:HcsOnlyCowork = $true
        $r = & $listBlock
        $got = 'nothing'
        if ($null -ne $r) { $got = "$($r.GetType().Name) '$r'" }
        Assert-Case "a list that exits 0 and prints nothing gives '', an answer" ($r -is [string] -and $r.Trim() -eq '') "got $got"
        Reset-State
        $script:HcsListCode = 1
        $r = & $listBlock
        $got = 'nothing'
        if ($null -ne $r) { $got = "$($r.GetType().Name) '$r'" }
        Assert-Case "a list that fails gives nothing" ($null -eq $r) "got $got"
    }

    # The shared function's side of it: an empty answer means every one is
    # gone, and no answer means not known.
    Reset-State
    $r = Get-CoworkHcsSurvivors -Guids @($guidA) -List { '' }
    Assert-Case "Get-CoworkHcsSurvivors: an empty answer is an empty list, not null" ($null -ne $r -and @($r).Count -eq 0)
    $r = Get-CoworkHcsSurvivors -Guids @($guidA) -List { }
    Assert-Case "and no answer is null" ($null -eq $r)
}

# =====================================================================
# Each driver option is set to the value that turns its power saving off,
# and only when the driver offers that value. Until 6.3.0 every option was
# set to 0, which is Auto SMPS for Intel's MIMO power save and Max Power
# Save for Broadcom's EEE, and the Wi-Fi options were not in the list.
Invoke-Section 'Network adapter power saving: each option to its own off value' {
    Reset-State
    $wPnp = 'PCI\VEN_8086&DEV_2725&SUBSYS_00248086&REV_1A\SYNTHWIFI'
    $bPnp = 'PCI\VEN_14E4&DEV_1682&SUBSYS_168214E4&REV_01\SYNTHBCM'
    $wKey = "$classRoot\$netClass\0020"
    $bKey = "$classRoot\$netClass\0021"
    $script:Nics = @(
        [pscustomobject]@{ Name = 'Wi-Fi'; InterfaceDescription = 'Intel Wi-Fi 6E AX210 160MHz'; PnPDeviceID = $wPnp },
        [pscustomobject]@{ Name = 'Ethernet 2'; InterfaceDescription = 'Broadcom NetXtreme Gigabit Ethernet'; PnPDeviceID = $bPnp }
    )
    $script:DrvKeys = @{ $wPnp = "$netClass\0020"; $bPnp = "$netClass\0021" }
    Set-FakeRegValue $wKey 'PnPCapabilities' ([int]24)
    Set-FakeRegValue $wKey 'DriverDesc' 'Intel Wi-Fi 6E AX210 160MHz'
    Set-FakeRegValue $bKey 'PnPCapabilities' ([int]24)
    Set-FakeRegValue $bKey 'DriverDesc' 'Broadcom NetXtreme Gigabit Ethernet'
    $script:DevEnable = @(
        [pscustomobject]@{ InstanceName = "${wPnp}_0"; Enable = $false },
        [pscustomobject]@{ InstanceName = "${bPnp}_0"; Enable = $false }
    )
    # Intel: 0 is Auto SMPS, the driver default, and 3 is No SMPS. Broadcom:
    # 0 is Max Power Save, 1 Balance, 2 Max Performance. Idle restriction:
    # 0 lets the adapter idle down at any time. This Green Ethernet offers
    # no 0, and the throughput booster is not a power saving option.
    $script:Adv = @{
        'Wi-Fi'      = @(
            (New-AdvProp 'Wi-Fi' 'MIMOPowerSaveMode'        'MIMO Power Save Mode' '0' @('0', '3')),
            (New-AdvProp 'Wi-Fi' 'ThroughputBoosterEnabled' 'Throughput Booster'   '1' @('0', '1'))
        )
        'Ethernet 2' = @(
            (New-AdvProp 'Ethernet 2' 'EeeCtrlMode'      'EEE Control Policies'        '1' @('0', '1', '2')),
            (New-AdvProp 'Ethernet 2' '*IdleRestriction' 'Idle Power Down Restriction' '0' @('0', '1')),
            (New-AdvProp 'Ethernet 2' 'GreenEthernet'    'Green Ethernet'              '1' @('1', '2'))
        )
    }
    $script:Pm = @{}
    foreach ($n in @('Wi-Fi', 'Ethernet 2')) {
        $script:Pm[$n] = [pscustomobject]@{
            Name = $n; AllowComputerToTurnOffDevice = 'Unsupported'; D0PacketCoalescing = 'Unsupported'
            DeviceSleepOnDisconnect = 'Unsupported'; SelectiveSuspend = 'Unsupported'; WakeOnMagicPacket = 'Unsupported'
        }
    }
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Complete-Step
    $s = Get-LastStep
    $adv = @($script:NicCalls | Where-Object { $_ -like 'Set-Adv *' }) -join ' / '
    Assert-Case "exactly three options are changed, each to its own off value, without a restart" ($adv -eq ('Set-Adv Wi-Fi|MIMOPowerSaveMode=3 NoRestart=True / ' +
        'Set-Adv Ethernet 2|EeeCtrlMode=2 NoRestart=True / Set-Adv Ethernet 2|*IdleRestriction=1 NoRestart=True')) $adv
    Assert-Case "Intel MIMO power save is No SMPS (3), not Auto SMPS (0)" ((@($script:Adv['Wi-Fi'][0].RegistryValue) -join ',') -eq '3')
    Assert-Case "Broadcom EEE is Max Performance (2), never Max Power Save (0)" ((@($script:Adv['Ethernet 2'][0].RegistryValue) -join ',') -eq '2')
    Assert-Case "idle restriction is 1, idle only when the user is away" ((@($script:Adv['Ethernet 2'][1].RegistryValue) -join ',') -eq '1')
    Assert-Case "an option whose off value the driver does not offer is left alone" ((@($script:Adv['Ethernet 2'][2].RegistryValue) -join ',') -eq '1' -and
        $null -eq (Find-Undo 'NicAdvanced' "$bPnp|GreenEthernet"))
    Assert-Case "an option that is not power saving is left alone" ((@($script:Adv['Wi-Fi'][1].RegistryValue) -join ',') -eq '1' -and
        $null -eq (Find-Undo 'NicAdvanced' "$wPnp|ThroughputBoosterEnabled"))
    $orig = @(@("$wPnp|MIMOPowerSaveMode", "$bPnp|EeeCtrlMode", "$bPnp|*IdleRestriction") | ForEach-Object {
        $it = Find-Undo 'NicAdvanced' $_
        if ($null -eq $it) { '<none>' } else { @($it.Original) -join ',' }
    }) -join ' '
    Assert-Case "three undo records, each with the value it had" ($script:UndoRecord.Count -eq 3 -and $orig -eq '0 1 0') "$orig; $(Get-UndoKeys)"
    $notes = Get-StepNotes
    Assert-Case "each adapter says how many it turned off" ($notes -match [regex]::Escape('Wi-Fi (Intel Wi-Fi 6E AX210 160MHz): 1 power saving setting(s) turned off') -and
        $notes -match [regex]::Escape('Ethernet 2 (Broadcom NetXtreme Gigabit Ethernet): 2 power saving setting(s) turned off')) $notes
    Assert-Case "Done, after a Windows restart" ($s.Status -eq 'Done' -and $s.Reboot) "got $($s.Status), reboot $($s.Reboot)"

    $script:NicCalls = New-Object System.Collections.ArrayList
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Complete-Step
    Assert-Case "a second run sees each one off and changes nothing" ((Get-StepStatus) -eq 'Unchanged' -and $script:NicCalls.Count -eq 0) "got $(Get-StepStatus), $($script:NicCalls -join ' / ')"

    $script:UndoRecord = New-Object System.Collections.ArrayList
    Read-PreventUndo
    $items = @($script:UndoRecord.ToArray())
    [array]::Reverse($items)
    Step 1 5 'undo' -Name 'Recorded settings'
    $kept = @()
    foreach ($it in $items) {
        if (Restore-PreventUndoItem -Item $it -CreatedPlan $null) { Clear-PreventUndoItem -Kind $it.Kind -Key $it.Key } else { $kept += $it.Key }
    }
    $vals = "$(@($script:Adv['Wi-Fi'][0].RegistryValue) -join ',') $(@($script:Adv['Ethernet 2'][0].RegistryValue) -join ',') $(@($script:Adv['Ethernet 2'][1].RegistryValue) -join ',')"
    Assert-Case "-Undo puts each one back as it was" ($vals -eq '0 1 0') $vals
    Assert-Case "and the record is gone" ($kept.Count -eq 0 -and $script:UndoRecord.Count -eq 0) "kept: $($kept -join ', ')"
}

# =====================================================================
# The NetAdapter cmdlets read -Name as a wildcard pattern, so until 6.3.0
# an adapter named "Ethernet [LAN]" matched nothing, without an error:
# setup passed over every option on it, and -Undo cleared the records of
# one renamed so after setup without putting a value back.
Invoke-Section 'Network adapter power saving: an adapter whose name holds [ or ]' {
    $offKws = @('*EEE', 'EnableGreenEthernet', 'GigaLite', 'PowerSavingMode')
    Reset-State
    Set-NicFixture
    Rename-FakeNic 'Ethernet' 'Ethernet [LAN]'
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Complete-Step
    $adv = @($script:NicCalls | Where-Object { $_.StartsWith('Set-Adv ') })
    $kws = @($adv | ForEach-Object { ($_ -split '\|')[1] -replace '=.*$', '' } | Sort-Object) -join ','
    Assert-Case "the four power saving options on Ethernet [LAN] are found and turned off" ($kws -eq ($offKws -join ',')) "got $kws"
    Assert-Case "each on that adapter, by its own name" ($adv.Count -eq 4 -and @($adv | Where-Object { -not $_.StartsWith('Set-Adv Ethernet [LAN]|') }).Count -eq 0) ($adv -join ' / ')
    $pmc = @($script:NicCalls | Where-Object { $_.StartsWith('Set-Pm ') }) -join ' / '
    Assert-Case "and its selective suspend" ($pmc -eq 'Set-Pm Ethernet [LAN]|SelectiveSuspend=Disabled NoRestart=True') $pmc
    $notes = Get-StepNotes
    Assert-Case "it says six settings turned off on Ethernet [LAN]: Done" ((Get-StepStatus) -eq 'Done' -and
        $notes -match [regex]::Escape('Ethernet [LAN] (Realtek Gaming 2.5GbE Family Controller): 6 power saving setting(s) turned off')) "$(Get-StepStatus): $notes"

    $script:NicCalls = New-Object System.Collections.ArrayList
    $script:UndoRecord = New-Object System.Collections.ArrayList
    Read-PreventUndo
    $items = @($script:UndoRecord.ToArray())
    [array]::Reverse($items)
    Step 1 5 'undo' -Name 'Recorded settings'
    $kept = @()
    foreach ($it in $items) {
        if (Restore-PreventUndoItem -Item $it -CreatedPlan $null) { Clear-PreventUndoItem -Kind $it.Kind -Key $it.Key } else { $kept += $it.Key }
    }
    $vals = @($script:Adv['Ethernet [LAN]'] | Where-Object { $offKws -contains $_.RegistryKeyword } |
        ForEach-Object { "$($_.RegistryKeyword)=$(@($_.RegistryValue) -join ',')" }) -join ' '
    Assert-Case "-Undo puts every option on Ethernet [LAN] back" ($vals -eq '*EEE=1 EnableGreenEthernet=1 GigaLite=1 PowerSavingMode=1') $vals
    Assert-Case "and its selective suspend goes back on" ($script:Pm['Ethernet [LAN]'].SelectiveSuspend -eq 'Enabled')
    Assert-Case "and the record is gone" ($kept.Count -eq 0 -and $script:UndoRecord.Count -eq 0) "kept: $($kept -join ', ')"

    # Renamed between setup and -Undo: -Undo finds the adapter by its device
    # id, and reads its options by the name it has now.
    Reset-State
    Set-NicFixture
    Step 6 19 'nic' -Name 'Network adapter power saving'
    Disable-NicPowerSaving
    Complete-Step
    Rename-FakeNic 'Ethernet' 'Ethernet [LAN]'
    $script:NicCalls = New-Object System.Collections.ArrayList
    $script:UndoRecord = New-Object System.Collections.ArrayList
    Read-PreventUndo
    $items = @($script:UndoRecord.ToArray())
    [array]::Reverse($items)
    Step 1 5 'undo' -Name 'Recorded settings'
    $kept = @()
    foreach ($it in $items) {
        if (Restore-PreventUndoItem -Item $it -CreatedPlan $null) { Clear-PreventUndoItem -Kind $it.Kind -Key $it.Key } else { $kept += $it.Key }
    }
    $vals = @($script:Adv['Ethernet [LAN]'] | Where-Object { $offKws -contains $_.RegistryKeyword } |
        ForEach-Object { "$($_.RegistryKeyword)=$(@($_.RegistryValue) -join ',')" }) -join ' '
    Assert-Case "renamed to Ethernet [LAN] after setup: -Undo still puts every option back" ($vals -eq '*EEE=1 EnableGreenEthernet=1 GigaLite=1 PowerSavingMode=1') $vals
    Assert-Case "and its selective suspend, renamed" ($script:Pm['Ethernet [LAN]'].SelectiveSuspend -eq 'Enabled')
    $named = @($script:NicCalls | Where-Object { $_.StartsWith('Set-Adv ') -or $_.StartsWith('Set-Pm ') })
    Assert-Case "each by the new name" ($named.Count -eq 5 -and @($named | Where-Object { $_ -notmatch '^Set-(Adv|Pm) Ethernet \[LAN\]\|' }).Count -eq 0) ($named -join ' / ')
    Assert-Case "and the renamed adapter's record is gone" ($kept.Count -eq 0 -and $script:UndoRecord.Count -eq 0) "kept: $($kept -join ', ')"
}

# =====================================================================
# Versions before 6.3.0 wrote PnPCapabilities = 24 to the adapter's Enum
# key and turned its wake settings off. Until 6.3.0 -Undo removed that 24
# first, and the 24 is the only sign that an older version touched the
# adapter: a wake setting that could not be put back was never tried again.
Invoke-Section "-Undo step 2: an older version's wake settings on a network adapter" {
    $enumKey = "HKLM:\SYSTEM\CurrentControlSet\Enum\$ethPnp\Device Parameters"
    $label = 'Ethernet (Realtek Gaming 2.5GbE Family Controller)'
    function Set-LegacyWakeFixture {
        Reset-State
        Set-NicFixture
        Set-FakeRegValue $enumKey 'PnPCapabilities' ([int]24)
        foreach ($p in $script:Adv['Ethernet']) {
            if ($p.RegistryKeyword -eq '*WakeOnMagicPacket') { $p.DisplayValue = 'Disabled' }
        }
        $script:Adv['Ethernet'] = @($script:Adv['Ethernet']) + @(New-AdvProp 'Ethernet' '*WakeOnPattern' 'Wake on Pattern Match' '0' @('0', '1') -DisplayValue 'Disabled')
        # The second instance only starts with the Realtek's device id.
        $script:WakeEnable = @(
            [pscustomobject]@{ InstanceName = "${ethPnp}_0"; Enable = $false },
            [pscustomobject]@{ InstanceName = "${ethPnp}X_0"; Enable = $false }
        )
    }
    $both = 'Set-AdvDisplay Ethernet|Wake on Magic Packet=Enabled NoRestart=True / Set-AdvDisplay Ethernet|Wake on Pattern Match=Enabled NoRestart=True'

    Set-LegacyWakeFixture
    Invoke-UndoStep2
    $disp = @($script:NicCalls | Where-Object { $_.StartsWith('Set-AdvDisplay ') }) -join ' / '
    Assert-Case "Wake on Magic Packet and Wake on Pattern Match go back on, without a restart" ($disp -eq $both) $disp
    $cim = @($script:NicCalls | Where-Object { $_.StartsWith('Set-CimInstance ') }) -join ' / '
    Assert-Case "and 'allow this device to wake the computer', on the Realtek's own instance only" ($cim -eq "Set-CimInstance ${ethPnp}_0 Enable=True" -and
        $script:WakeEnable[0].Enable -eq $true -and $script:WakeEnable[1].Enable -eq $false) $cim
    Assert-Case "the 24 is removed" (-not (Test-FakeRegValue $enumKey 'PnPCapabilities'))
    Assert-Case "Done, after a Windows restart, and it says so" ((Get-StepStatus) -eq 'Done' -and (Get-LastStep).Reboot -and
        (Test-LogHas ('^' + [regex]::Escape("${label}: removed an older version's settings and turned wake back on") + '$'))) "$(Get-StepStatus): $(Get-StepNotes)"

    Set-LegacyWakeFixture
    $script:AdvDisplayFault = $true
    Invoke-UndoStep2
    $notes = Get-StepNotes
    Assert-Case "a wake setting that cannot be put back: Failed, with the adapter and why" ((Get-StepStatus) -eq 'Failed' -and
        $notes -match [regex]::Escape("${label}: The parameter is incorrect.")) "$(Get-StepStatus): $notes"
    Assert-Case "and the 24 stays, so the next -Undo tries again" ((Get-FakeRegValue $enumKey 'PnPCapabilities') -eq 24 -and
        @($script:Calls | Where-Object { $_.StartsWith('Remove-ItemProperty ') }).Count -eq 0) ($script:Calls -join ' / ')
    $script:AdvDisplayFault = $false
    $script:StepResults = New-Object System.Collections.ArrayList
    $script:NicCalls = New-Object System.Collections.ArrayList
    Invoke-UndoStep2
    $disp = @($script:NicCalls | Where-Object { $_.StartsWith('Set-AdvDisplay ') }) -join ' / '
    Assert-Case "which it does: wake goes back on and the 24 is removed" ((Get-StepStatus) -eq 'Done' -and $disp -eq $both -and
        $script:WakeEnable[0].Enable -eq $true -and -not (Test-FakeRegValue $enumKey 'PnPCapabilities')) "$(Get-StepStatus): $disp; $(Get-StepNotes)"

    Set-LegacyWakeFixture
    $script:SetCimThrows = $true
    Invoke-UndoStep2
    Assert-Case "'allow this device to wake' that cannot be set: Failed, and the 24 stays" ((Get-StepStatus) -eq 'Failed' -and
        (Get-FakeRegValue $enumKey 'PnPCapabilities') -eq 24 -and (Get-StepNotes) -match [regex]::Escape("${label}: Access denied")) "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    Set-NicFixture
    Set-FakeRegValue $enumKey 'PnPCapabilities' ([int]24)
    $script:WakeEnable = @([pscustomobject]@{ InstanceName = "${ethPnp}_0"; Enable = $true })
    Invoke-UndoStep2
    Assert-Case "wake already on: nothing is set, and the 24 is removed" ($script:NicCalls.Count -eq 0 -and (Get-StepStatus) -eq 'Done' -and
        -not (Test-FakeRegValue $enumKey 'PnPCapabilities')) "$(Get-StepStatus): $($script:NicCalls -join ' / ')"

    # An adapter list that cannot be read is said, and the 24 stays for the
    # next -Undo. Passed over, it read as nothing from older versions found.
    Set-LegacyWakeFixture
    $script:NicThrows = $true
    Invoke-UndoStep2
    Assert-Case "the adapters cannot be listed: a warning that says so, and the 24 stays" ((Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -eq
        'Could not list the network adapters, so the wake settings older versions turned off on them were not checked: The WMI provider is not loaded' -and
        (Get-FakeRegValue $enumKey 'PnPCapabilities') -eq 24 -and $script:NicCalls.Count -eq 0) "$(Get-StepStatus): $(Get-StepNotes); $($script:NicCalls -join ' / ')"
    Assert-Case "and it does not say nothing from older versions was found" ((Test-LogHas ('^' + [regex]::Escape(
        '[!] Could not list the network adapters, so the wake settings older versions turned off were not checked: The WMI provider is not loaded') + '$')) -and
        -not (Test-LogHas '^Nothing from older versions found$')) ($script:LogLines -join ' / ')

    $end = Get-UndoEnd
    Assert-Case "and -Undo ends UNDO INCOMPLETE, saying to run it again, with exit code 1" ($end.Box -eq 'UNDO INCOMPLETE' -and $end.Unread -and $end.Exit -eq 1) $end.Text

    # With an older install recorded, that record stays, so that the next
    # -Undo still knows one ran here. That -Undo, with the adapters read,
    # puts wake back on, drops the record and is complete.
    Set-LegacyWakeFixture
    Add-LegacyTrace
    $script:NicThrows = $true
    Invoke-UndoStep2
    Assert-Case "with an older install recorded, an adapter list that cannot be read keeps that record" ((Get-StepStatus) -eq 'Warning' -and
        (Get-LegacyInstallCount) -eq 1 -and (Get-FakeRegValue $enumKey 'PnPCapabilities') -eq 24) "$(Get-StepStatus): $(Get-StepNotes); $(Get-UndoKeys)"
    $script:NicThrows = $false
    # The next run starts with the flag clear, as Prevent sets it.
    $script:UndoUnread = $false
    $script:StepResults = New-Object System.Collections.ArrayList
    $script:NicCalls = New-Object System.Collections.ArrayList
    Invoke-UndoStep2
    $end = Get-UndoEnd
    Assert-Case "and the next -Undo, with the list read, puts wake back on, drops the record, and ends UNDO COMPLETE with exit code 0" ((Get-StepStatus) -eq 'Done' -and
        -not (Test-FakeRegValue $enumKey 'PnPCapabilities') -and (Get-LegacyInstallCount) -eq 0 -and $end.Box -eq 'UNDO COMPLETE' -and -not $end.Unread -and
        $end.Exit -eq 0) "$(Get-StepStatus): $(Get-StepNotes); $(Get-UndoKeys); $($end.Text)"

    Set-LegacyWakeFixture
    Rename-FakeNic 'Ethernet' 'Ethernet [LAN]'
    Invoke-UndoStep2
    $disp = @($script:NicCalls | Where-Object { $_.StartsWith('Set-AdvDisplay ') }) -join ' / '
    Assert-Case "an adapter named Ethernet [LAN]: its wake settings are found and put back" ($disp -eq $both.Replace('Ethernet|', 'Ethernet [LAN]|') -and
        -not (Test-FakeRegValue $enumKey 'PnPCapabilities')) $disp

    # The devices that can wake the computer, read for the 'allow this
    # device to wake the computer' setting. root\wmi answers Not supported
    # for a class no device has: then there is nothing to turn back on. Any
    # other failure fails the adapter, and the 24 stays for the next -Undo.
    # Taken for no device, it let the 24 go with that setting still off.
    Set-LegacyWakeFixture
    $script:WakeReadCim = New-FakeCimError -Status 2 -Code 0x80041003L -Message 'Access denied ' -Category PermissionDenied
    Invoke-UndoStep2
    Assert-Case "the devices that can wake cannot be listed: Failed, naming the adapter and why, and the 24 stays" ((Get-StepStatus) -eq 'Failed' -and
        @((Get-LastStep).Notes)[0] -eq "${label}: Access denied " -and (Get-FakeRegValue $enumKey 'PnPCapabilities') -eq 24 -and
        @($script:NicCalls | Where-Object { $_.StartsWith('Set-CimInstance ') }).Count -eq 0) "$(Get-StepStatus): $(Get-StepNotes); $($script:NicCalls -join ' / ')"
    $script:WakeReadCim = $null
    $script:StepResults = New-Object System.Collections.ArrayList
    $script:NicCalls = New-Object System.Collections.ArrayList
    Invoke-UndoStep2
    Assert-Case "which the next -Undo finishes: wake goes back on and the 24 is removed" ((Get-StepStatus) -eq 'Done' -and $script:WakeEnable[0].Enable -eq $true -and
        -not (Test-FakeRegValue $enumKey 'PnPCapabilities')) "$(Get-StepStatus): $(Get-StepNotes); $($script:NicCalls -join ' / ')"

    Set-LegacyWakeFixture
    $script:WakeReadCim = New-FakeCimError -Status 7 -Code 0x8004100CL -Message 'Not supported ' -Category NotImplemented
    Invoke-UndoStep2
    $disp = @($script:NicCalls | Where-Object { $_.StartsWith('Set-AdvDisplay ') }) -join ' / '
    Assert-Case "no device can wake the computer (Not supported): the wake options go back on, the 24 is removed, and nothing fails" ((Get-StepStatus) -eq 'Done' -and
        $disp -eq $both -and @($script:NicCalls | Where-Object { $_.StartsWith('Set-CimInstance ') }).Count -eq 0 -and
        -not (Test-FakeRegValue $enumKey 'PnPCapabilities')) "$(Get-StepStatus): $(Get-StepNotes); $($script:NicCalls -join ' / ')"
}

# =====================================================================
# ServicesPipeTimeout = 120000, LocalAccountTokenFilterPolicy = 1 and
# FilterAdministratorToken = 0 are what versions before 6.3.0 set, and also
# what an administrator or a policy can set. Until 6.3.0 -Undo removed them
# wherever they held those values; now only where an older version is known
# to have run, and the record of that is kept until they are gone.
Invoke-Section '-Undo step 2: values older versions set, removed only where one ran' {
    $policyPath  = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
    $controlPath = 'HKLM:\SYSTEM\CurrentControlSet\Control'
    $flagPath    = Join-Path $ClaudeAppData 'disable-dynamic-memory.flag'
    function Set-LegacyValues {
        Set-FakeRegValue $controlPath 'ServicesPipeTimeout' ([int]120000)
        Set-FakeRegValue $policyPath 'LocalAccountTokenFilterPolicy' ([int]1)
        Set-FakeRegValue $policyPath 'FilterAdministratorToken' ([int]0)
    }
    function Get-LegacyValuesLeft {
        $left = @()
        foreach ($x in @(@($controlPath, 'ServicesPipeTimeout'), @($policyPath, 'LocalAccountTokenFilterPolicy'), @($policyPath, 'FilterAdministratorToken'))) {
            if (Test-FakeRegValue $x[0] $x[1]) { $left += $x[1] }
        }
        return ($left -join ',')
    }
    function Get-LegacyRecordCount { return @($script:UndoRecord | Where-Object { $_.Kind -eq 'LegacyInstall' }).Count }
    $tail = ' matches what older versions set, but no older install was found here, so it is left as it is'
    # "record from setup" is the record Add-LegacyInstallRecord writes (step
    # 19's section drives that one), written here as the record itself so
    # that the same cases also run against versions without it.

    Reset-State
    Set-LegacyValues
    Invoke-UndoStep2
    $notes = Get-StepNotes
    Assert-Case "no older install found: each of the three is named and left as it is" ($notes -eq
        ("ServicesPipeTimeout = 120000$tail | LocalAccountTokenFilterPolicy = 1$tail | FilterAdministratorToken = 0$tail")) $notes
    Assert-Case "none is removed, and the step is Unchanged" ((Get-LegacyValuesLeft) -eq 'ServicesPipeTimeout,LocalAccountTokenFilterPolicy,FilterAdministratorToken' -and
        @($script:Calls | Where-Object { $_.StartsWith('Remove-ItemProperty ') }).Count -eq 0 -and (Get-StepStatus) -eq 'Unchanged') "$(Get-StepStatus): left $(Get-LegacyValuesLeft)"
    Assert-Case "and nothing is recorded" ($script:UndoRecord.Count -eq 0) (Get-UndoKeys)

    foreach ($how in @('flag file', 'logon task', 'record from setup')) {
        Reset-State
        Set-LegacyValues
        switch ($how) {
            'flag file'         { [System.IO.File]::WriteAllText($flagPath, '1') }
            'logon task'        { $script:Tasks["$TaskPath|$BootTaskName"] = 'Ready' }
            'record from setup' { Add-PreventUndoItem -Kind LegacyInstall -Key 'boot-fix task' -Original "$TaskPath$BootTaskName" }
        }
        Invoke-UndoStep2
        Assert-Case "an older install shown by its $how, all three are removed: Done" ((Get-LegacyValuesLeft) -eq '' -and (Get-StepStatus) -eq 'Done' -and
            (Test-LogHas '^ServicesPipeTimeout = 120000: removed$') -and (Test-LogHas '^LocalAccountTokenFilterPolicy = 1: removed$') -and
            (Test-LogHas '^FilterAdministratorToken = 0: removed$')) "$(Get-StepStatus): left $(Get-LegacyValuesLeft); $(Get-StepNotes)"
        Assert-Case "and with them the record that one ran here ($how)" ((Get-LegacyRecordCount) -eq 0) (Get-UndoKeys)
    }

    Reset-State
    Set-LegacyValues
    Set-FakeRegValue $controlPath 'ServicesPipeTimeout' ([int]60000)
    [System.IO.File]::WriteAllText($flagPath, '1')
    Invoke-UndoStep2
    Assert-Case "a value older versions did not write is left, even where one ran" ((Get-FakeRegValue $controlPath 'ServicesPipeTimeout') -eq 60000 -and
        (Get-LegacyValuesLeft) -eq 'ServicesPipeTimeout') "left $(Get-LegacyValuesLeft)"

    # One that cannot be removed fails the step and keeps the record that
    # an older version ran here, after the file that showed it is gone.
    Reset-State
    Set-LegacyValues
    [System.IO.File]::WriteAllText($flagPath, '1')
    $script:RegRemoveFault[(Get-RegNorm $policyPath) + '|filteradministratortoken'] = $true
    Invoke-UndoStep2
    $notes = Get-StepNotes
    Assert-Case "a value that cannot be removed: Failed, and it says which and why" ((Get-StepStatus) -eq 'Failed' -and
        $notes -match [regex]::Escape('Could not remove FilterAdministratorToken: Requested registry access is not allowed.')) "$(Get-StepStatus): $notes"
    Assert-Case "the other two are removed" ((Get-LegacyValuesLeft) -eq 'FilterAdministratorToken') "left $(Get-LegacyValuesLeft)"
    Assert-Case "the flag is gone, and the record that an older version ran here is kept" (-not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $flagPath) -and
        (Get-LegacyRecordCount) -eq 1 -and $null -ne (Find-Undo 'LegacyInstall' 'Dynamic Memory flag')) (Get-UndoKeys)
    $script:RegRemoveFault = @{}
    $script:StepResults = New-Object System.Collections.ArrayList
    $script:LogLines = New-Object System.Collections.ArrayList
    Invoke-UndoStep2
    Assert-Case "so the next -Undo removes it, and then the record" ((Get-LegacyValuesLeft) -eq '' -and (Get-StepStatus) -eq 'Done' -and
        (Get-LegacyRecordCount) -eq 0 -and (Test-LogHas '^FilterAdministratorToken = 0: removed$')) "$(Get-StepStatus): left $(Get-LegacyValuesLeft); $(Get-UndoKeys)"
}

# =====================================================================
# The monitor of versions before 6.3.0 turned Dynamic Memory off on the
# Hyper-V VMs named like Claude, with no record of it, and until 6.3.0
# nothing turned it back on. -Undo now does, where an older version is known
# to have run, and only while the VM is off, which is when Hyper-V allows it.
Invoke-Section '-Undo step 2: Dynamic Memory an older version turned off' {
    function Set-DmFixture {
        param([switch]$NoTrace)
        Reset-State
        # The record setup's Add-LegacyInstallRecord writes.
        if (-not $NoTrace) { Add-PreventUndoItem -Kind LegacyInstall -Key 'Dynamic Memory flag' -Original (Join-Path $ClaudeAppData 'disable-dynamic-memory.flag') }
        $script:Vms = @(
            (New-FakeVm 'Claude Cowork' 'Off' $false),
            (New-FakeVm 'claude-dm-on' 'Off' $true),
            (New-FakeVm 'ubuntu-dev' 'Off' $false)
        )
    }
    function Get-LegacyRecordCount { return @($script:UndoRecord | Where-Object { $_.Kind -eq 'LegacyInstall' }).Count }
    function Get-VmSets { return (@($script:VmCalls | Where-Object { $_.StartsWith('Set-VMMemory ') }) -join ' / ') }
    $running = 'Dynamic Memory on the VM Claude Cowork is still off: Hyper-V changes it only while the VM is off, so turn the VM off and run -Undo again'

    Set-DmFixture
    Invoke-UndoStep2
    Assert-Case "on the Claude VM that has it off, and on no other VM, Dynamic Memory goes back on" ((Get-VmSets) -eq 'Set-VMMemory Claude Cowork DynamicMemoryEnabled=True' -and
        $script:Vms[0].DynamicMemoryEnabled -and -not $script:Vms[2].DynamicMemoryEnabled) (Get-VmSets)
    Assert-Case "Done, and it says so" ((Get-StepStatus) -eq 'Done' -and (Test-LogHas '^Dynamic Memory on the VM Claude Cowork: turned back on$')) "$(Get-StepStatus): $(Get-StepNotes)"
    Assert-Case "and the record that an older version ran here goes" ((Get-LegacyRecordCount) -eq 0) (Get-UndoKeys)
    $end = Get-UndoEnd
    Assert-Case "and -Undo ends UNDO COMPLETE with exit code 0" ($end.Box -eq 'UNDO COMPLETE' -and -not $end.Unread -and $end.Exit -eq 0) $end.Text

    Set-DmFixture
    $script:Vms[0].State = 'Running'
    Invoke-UndoStep2
    Assert-Case "a running VM is named, and not changed" ((Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -match [regex]::Escape($running) -and (Get-VmSets) -eq '') "$(Get-StepStatus): $(Get-StepNotes); $(Get-VmSets)"
    Assert-Case "and the record that an older version ran here is kept for the next -Undo" ((Get-LegacyRecordCount) -eq 1) (Get-UndoKeys)
    $script:Vms[0].State = 'Off'
    $script:StepResults = New-Object System.Collections.ArrayList
    $script:LogLines = New-Object System.Collections.ArrayList
    Invoke-UndoStep2
    Assert-Case "which turns it back on once the VM is off, and then drops the record" ((Get-StepStatus) -eq 'Done' -and $script:Vms[0].DynamicMemoryEnabled -and
        (Get-LegacyRecordCount) -eq 0) "$(Get-StepStatus): $(Get-StepNotes); $(Get-UndoKeys)"

    Set-DmFixture
    $script:VmReadFault = $true
    Invoke-UndoStep2
    Assert-Case "Hyper-V cannot be read: a warning that says so, and the record is kept" ((Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -eq
        'Could not read the Hyper-V VMs, so Dynamic Memory an older version turned off was not checked: The Hyper-V Virtual Machine Management service is not running.' -and
        (Get-LegacyRecordCount) -eq 1) "$(Get-StepStatus): $(Get-StepNotes)"
    Assert-Case "and it does not then say nothing from older versions was found" (-not (Test-LogHas '^Nothing from older versions found$')) (
        $script:LogLines -join ' / ')

    $end = Get-UndoEnd
    Assert-Case "and -Undo ends UNDO INCOMPLETE, saying some lists could not be read, with exit code 1" ($end.Box -eq 'UNDO INCOMPLETE' -and $end.Unread -and
        $end.Exit -eq 1) $end.Text

    Set-DmFixture
    $script:VmNoop = $true
    Invoke-UndoStep2
    Assert-Case "a change that does not take: Failed, and the record is kept" ((Get-StepStatus) -eq 'Failed' -and (Get-StepNotes) -match
        [regex]::Escape('Dynamic Memory on the VM Claude Cowork still reads off') -and (Get-LegacyRecordCount) -eq 1) "$(Get-StepStatus): $(Get-StepNotes)"

    # Another product's Get-VM lists objects with no Dynamic Memory at all.
    Set-DmFixture
    $script:Vms = @([pscustomobject]@{ Name = 'Claude (VMware)'; State = 'Off' })
    Invoke-UndoStep2
    Assert-Case "a VM with no Dynamic Memory property is not a Hyper-V VM: left alone, no warning" ((Get-VmSets) -eq '' -and
        @('Warning', 'Failed') -notcontains (Get-StepStatus) -and (Get-LegacyRecordCount) -eq 0) "$(Get-StepStatus): $(Get-StepNotes)"

    Set-DmFixture -NoTrace
    Invoke-UndoStep2
    Assert-Case "no older install found: Hyper-V is not even read" ($script:VmCalls.Count -eq 0 -and -not $script:Vms[0].DynamicMemoryEnabled) ($script:VmCalls -join ' / ')

    # Hyper-V's cmdlets are not installed: Dynamic Memory is not looked at,
    # no VM cmdlet is called, and with nothing else from older versions
    # the step says so. Nothing is left for a later -Undo, so the record
    # that an older version ran here goes.
    foreach ($absent in @(@('Set-VMMemory'), @('Get-VM'), @('Get-VM', 'Set-VMMemory'))) {
        Set-DmFixture
        $script:AbsentCommands = $absent
        Invoke-UndoStep2
        Assert-Case "no $($absent -join ' or '): no VM cmdlet is called, nothing fails, and nothing from older versions is found" (
            $script:VmCalls.Count -eq 0 -and -not $script:Vms[0].DynamicMemoryEnabled -and (Get-StepStatus) -eq 'Unchanged' -and
            (Get-StepNotes) -eq '' -and (Test-LogHas '^Nothing from older versions found$') -and -not (Test-LogHas 'Dynamic Memory|Hyper-V') -and
            (Get-LegacyRecordCount) -eq 0) "$(Get-StepStatus): $(Get-StepNotes); $($script:VmCalls -join ' / '); $($script:LogLines -join ' / ')"
    }
}

# =====================================================================
# Where WinNAT's class or provider is missing, as on a PC without Hyper-V
# or containers, Get-NetNat fails on every run although there is no rule.
# Those errors are known by their CIM status or WMI code, never by their
# message; any other error is a rule that could not be read.
Invoke-Section 'Test-WinNatMissingError: the Get-NetNat errors that mean there is no WinNAT' {
    $rows = @(
        @{ Why = 'Invalid namespace'; Status = 3; Code = 0x8004100EL; Want = $true }
        @{ Why = 'Invalid class'; Status = 5; Code = 0x80041010L; Want = $true }
        @{ Why = 'Not supported'; Status = 7; Code = 0x8004100CL; Want = $true }
        @{ Why = 'Provider not found'; Status = 1; Code = 0x80041011L; Want = $true }
        @{ Why = 'Provider load failure'; Status = 1; Code = 0x80041013L; Want = $true }
        @{ Why = 'Invalid namespace, with no WMI code'; Status = 3; Code = -1; Want = $true }
        @{ Why = 'Invalid class, with no WMI code'; Status = 5; Code = -1; Want = $true }
        @{ Why = 'Not supported, with no WMI code'; Status = 7; Code = -1; Want = $true }
        @{ Why = 'Invalid namespace, under a Failed status'; Status = 1; Code = 0x8004100EL; Want = $true }
        @{ Why = 'Invalid class, under a Failed status'; Status = 1; Code = 0x80041010L; Want = $true }
        @{ Why = 'Not supported, under a Failed status'; Status = 1; Code = 0x8004100CL; Want = $true }
        @{ Why = 'A failure with no WMI code'; Status = 1; Code = -1; Want = $false }
        @{ Why = 'Access denied'; Status = 2; Code = 0x80041003L; Want = $false }
        @{ Why = 'Access denied, with no WMI code'; Status = 2; Code = -1; Want = $false }
        @{ Why = 'The wait operation timed out'; Status = 1; Code = 0x80070102L; Want = $false }
        @{ Why = 'Generic failure'; Status = 1; Code = 0x80041001L; Want = $false }
        @{ Why = 'Invalid class, said by a message with an access denied code'; Status = 2; Code = 0x80041003L; Want = $false }
    )
    foreach ($r in $rows) {
        $f = New-FakeCimError -Status $r.Status -Code $r.Code -Message "$($r.Why) "
        $got = Test-WinNatMissingError -ErrorRecord ([System.Management.Automation.ErrorRecord]::new($f.Exception, $f.Id, $f.Category, $null))
        $codeText = 'no WMI code'
        if ($r.Code -ge 0) { $codeText = '0x{0:X8}' -f $r.Code }
        Assert-Case "$($r.Why) (CIM status $($r.Status), $codeText): $(if ($r.Want) { 'no WinNAT' } else { 'a rule that could not be read' })" (
            $got -is [bool] -and $got -eq $r.Want) "got $got"
    }
    $f = New-FakeCimError -Status 5 -Code 0x80041010L -Message 'Invalid class '
    $wrapped = [System.Management.Automation.RuntimeException]::new('The CIM job failed', $f.Exception)
    Assert-Case "an Invalid class inside another exception is found there" (
        (Test-WinNatMissingError -ErrorRecord ([System.Management.Automation.ErrorRecord]::new($wrapped, 'CimJob', 'NotSpecified', $null))) -eq $true)
    Assert-Case "an error that is not a CIM one, worded as Invalid class, is not taken for no WinNAT" (
        (Test-WinNatMissingError -ErrorRecord ([System.Management.Automation.ErrorRecord]::new(
            [System.InvalidOperationException]::new('Invalid class '), 'x', 'MetadataError', $null))) -eq $false)
}

# =====================================================================
# The NAT rule CoworkNAT, which versions before 6.3.0 created and their
# monitor re-created. -Undo step 2 removes it whether or not an older
# install is found, removes no other rule, and says so when it cannot read
# the rules or remove this one.
Invoke-Section '-Undo step 2: the NAT rule CoworkNAT older versions created' {
    function Get-NatNames { return (@($script:Nats | ForEach-Object { $_.Name }) -join ',') }
    $natUnread = 'Could not read the NAT rules, so the rule CoworkNAT older versions created was not checked: Access denied'

    # A removal no case expects trips, so that one in setup, or anywhere
    # else but here, fails the suite.
    Reset-State
    $script:Nats = @((New-FakeNat 'CoworkNAT'))
    $before = $script:Tripped.Count
    $threw = ''
    try { Remove-NetNat -Name 'CoworkNAT' -Confirm:$false } catch { $threw = $_.Exception.Message }
    $trip = 'Remove-NetNat CoworkNAT Confirm=False'
    Assert-Case "a NAT removal no case expects trips, and the rule stays" ($threw -eq 'TRIPWIRE: Remove-NetNat' -and
        $script:Tripped.Count -eq $before + 1 -and $script:Tripped[$script:Tripped.Count - 1] -eq $trip -and (Get-NatNames) -eq 'CoworkNAT') (
        "threw '$threw'; $($script:Tripped -join ' / ')")
    if ($script:Tripped.Count -gt $before -and $script:Tripped[$script:Tripped.Count - 1] -eq $trip) { $script:Tripped.RemoveAt($script:Tripped.Count - 1) }

    Reset-State
    $script:Nats = @((New-FakeNat 'WSLNat'), (New-FakeNat 'CoworkNAT'))
    $script:NatRemoveExpected = $true
    Invoke-UndoStep2
    Assert-Case "CoworkNAT is removed, without a prompt, and no other rule, with no older install found" (
        ($script:NatCalls -join ' / ') -eq 'Get-NetNat CoworkNAT / Remove-NetNat CoworkNAT Confirm=False' -and (Get-NatNames) -eq 'WSLNat') (
        "$($script:NatCalls -join ' / '); left: $(Get-NatNames)")
    Assert-Case "Done, and it says so" ((Get-StepStatus) -eq 'Done' -and (Test-LogHas '^Removed the NAT rule CoworkNAT$')) "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    $script:Nats = @((New-FakeNat 'WSLNat'))
    Invoke-UndoStep2
    Assert-Case "no CoworkNAT: nothing is removed, and nothing from older versions is found" (($script:NatCalls -join ' / ') -eq 'Get-NetNat CoworkNAT' -and
        (Get-NatNames) -eq 'WSLNat' -and (Get-StepStatus) -eq 'Unchanged' -and (Test-LogHas '^Nothing from older versions found$')) (
        "$(Get-StepStatus): $(Get-StepNotes); $($script:NatCalls -join ' / ')")
    $end = Get-UndoEnd
    Assert-Case "and -Undo ends UNDO COMPLETE with exit code 0" ($end.Box -eq 'UNDO COMPLETE' -and -not $end.Unread -and $end.Exit -eq 0) $end.Text

    Reset-State
    $script:Nats = @((New-FakeNat 'CoworkNAT'))
    $script:NatRemoveFault = $true
    $script:NatRemoveExpected = $true
    Invoke-UndoStep2
    Assert-Case "a removal that is refused: Failed, saying why, and the rule is still there" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -match [regex]::Escape('Could not remove the NAT rule CoworkNAT: Access is denied.') -and (Get-NatNames) -eq 'CoworkNAT' -and
        (Test-LogHas ('^' + [regex]::Escape('[!] Could not remove the NAT rule CoworkNAT: Access is denied.') + '$'))) "$(Get-StepStatus): $(Get-StepNotes)"
    $script:NatRemoveFault = $false
    $script:StepResults = New-Object System.Collections.ArrayList
    $script:LogLines = New-Object System.Collections.ArrayList
    Invoke-UndoStep2
    Assert-Case "and the next -Undo removes it" ((Get-StepStatus) -eq 'Done' -and (Get-NatNames) -eq '') "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    $script:Nats = @((New-FakeNat 'CoworkNAT'))
    $script:AbsentCommands = @('Get-NetNat')
    Invoke-UndoStep2
    Assert-Case "no WinNAT cmdlets: no rule is read or removed" ($script:NatCalls.Count -eq 0 -and (Get-NatNames) -eq 'CoworkNAT' -and
        (Get-StepStatus) -eq 'Unchanged' -and (Test-LogHas '^Nothing from older versions found$')) "$(Get-StepStatus): $($script:NatCalls -join ' / ')"

    Reset-State
    $script:Nats = @((New-FakeNat 'CoworkNAT'))
    $script:NatReadFault = $true
    Invoke-UndoStep2
    Assert-Case "the NAT rules cannot be read: a warning that says so, and nothing is removed" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq $natUnread -and (Get-NatNames) -eq 'CoworkNAT' -and
        @($script:NatCalls | Where-Object { $_.StartsWith('Remove-NetNat') }).Count -eq 0) "$(Get-StepStatus): $(Get-StepNotes); $($script:NatCalls -join ' / ')"
    Assert-Case "and it does not then say nothing from older versions was found" (-not (Test-LogHas '^Nothing from older versions found$')) (
        $script:LogLines -join ' / ')

    $end = Get-UndoEnd
    Assert-Case "and -Undo ends UNDO INCOMPLETE, saying to run it again, with exit code 1" ($end.Box -eq 'UNDO INCOMPLETE' -and $end.Unread -and $end.Exit -eq 1) $end.Text

    Reset-State
    Add-LegacyTrace
    $script:Nats = @((New-FakeNat 'CoworkNAT'))
    $script:NatReadFault = $true
    Invoke-UndoStep2
    Assert-Case "with an older install recorded, NAT rules that cannot be read keep that record for the next -Undo" ((Get-StepStatus) -eq 'Warning' -and
        (Get-LegacyInstallCount) -eq 1 -and (Get-NatNames) -eq 'CoworkNAT') "$(Get-StepStatus): $(Get-StepNotes); $(Get-UndoKeys)"

    # WinNAT is not on this PC: no rule, so nothing is removed or said.
    foreach ($nw in @(@{ Why = 'Invalid class'; Status = 5; Code = 0x80041010L }, @{ Why = 'Provider load failure'; Status = 1; Code = 0x80041013L })) {
        Reset-State
        $script:Nats = @((New-FakeNat 'CoworkNAT'))
        $script:NatReadCim = New-FakeCimError -Status $nw.Status -Code $nw.Code -Message "$($nw.Why) "
        Invoke-UndoStep2
        Assert-Case "Get-NetNat fails with $($nw.Why): no WinNAT, so no warning, nothing is removed, and nothing from older versions is found" (
            (Get-StepStatus) -eq 'Unchanged' -and (Get-StepNotes) -eq '' -and ($script:NatCalls -join ' / ') -eq 'Get-NetNat CoworkNAT' -and
            (Test-LogHas '^Nothing from older versions found$') -and -not (Test-LogHas 'NAT rule')) (
            "$(Get-StepStatus): $(Get-StepNotes); $($script:NatCalls -join ' / '); $($script:LogLines -join ' / ')")
        $end = Get-UndoEnd
        Assert-Case "and -Undo ends UNDO COMPLETE with exit code 0" ($end.Box -eq 'UNDO COMPLETE' -and -not $end.Unread -and $end.Exit -eq 0) $end.Text
    }
    Reset-State
    $script:Nats = @((New-FakeNat 'CoworkNAT'))
    $script:NatReadCim = New-FakeCimError -Status 2 -Code 0x80041003L -Message 'Access denied ' -Category PermissionDenied
    Invoke-UndoStep2
    Assert-Case "Get-NetNat fails with a CIM Access denied: still a warning, and nothing is removed" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq "$natUnread " -and ($script:NatCalls -join ' / ') -eq 'Get-NetNat CoworkNAT' -and
        -not (Test-LogHas '^Nothing from older versions found$')) "$(Get-StepStatus): $(Get-StepNotes); $($script:NatCalls -join ' / ')"
    $end = Get-UndoEnd
    Assert-Case "and -Undo ends UNDO INCOMPLETE, saying to run it again, with exit code 1" ($end.Box -eq 'UNDO INCOMPLETE' -and $end.Unread -and $end.Exit -eq 1) $end.Text
}

# =====================================================================
# Versions before 6.3.0 set restart recovery on vmcompute with no record,
# and -Undo step 2 clears it where it is still exactly theirs. A 6.3.0
# setup records vmcompute's recovery; where step 1 could not put that back,
# clearing it here left it as neither, and until 6.3.0 it was cleared anyway.
Invoke-Section '-Undo step 2: service recovery still recorded for step 1' {
    $svcKey = 'HKLM:\SYSTEM\CurrentControlSet\Services\vmcompute'
    function Set-RecoveryFixture {
        Reset-State
        $script:Services['vmcompute'] = 'Running'
        Set-FakeRegValue $svcKey 'FailureActions' (New-FailureBlob -Reset 300 -Pairs @(1, 30000, 1, 60000, 1, 120000))
    }
    function Test-RecoveryOn { return [bool](Test-RestartRecovery (ConvertFrom-FailureActionsBlob -Bytes (Get-ServiceFailureBytes -Service 'vmcompute'))) }

    Set-RecoveryFixture
    Add-PreventUndoItem -Kind ServiceFailureActions -Key 'vmcompute' -Original $null
    Invoke-UndoStep2
    Assert-Case "vmcompute's recovery still recorded: not cleared here, and it says why" ($script:ScCalls.Count -eq 0 -and
        (Test-LogHas '^vmcompute: its recovery settings are still recorded, so step 1 puts them back on the next -Undo$') -and
        -not (Test-LogHas 'recovery set by an older version cleared')) ($script:ScCalls -join ' / ')
    Assert-Case "the recovery is as it was, and its record stays for step 1" ((Test-RecoveryOn) -and $null -ne (Find-Undo 'ServiceFailureActions' 'vmcompute')) (Get-UndoKeys)

    Set-RecoveryFixture
    Invoke-UndoStep2
    Assert-Case "no record: the recovery an older version set is cleared: Done" (($script:ScCalls -join ' / ') -eq 'failure vmcompute reset= 0 actions= ""' -and
        (Get-StepStatus) -eq 'Done' -and (Test-LogHas '^vmcompute: recovery set by an older version cleared$')) "$(Get-StepStatus): $($script:ScCalls -join ' / ')"
    Assert-Case "and it is off" (-not (Test-RecoveryOn))
}

# =====================================================================
# Step 19 removes the logon task and the flag file older versions left,
# and names what -Undo removes. The task and the flag show that an older
# version ran here, so each is recorded before it goes: until 6.3.0 neither
# was, and once setup had run, -Undo could not tell that one had. The three
# values older versions set are called theirs only where one is known to
# have run, as -Undo removes them only there.
Invoke-Section 'Setup step 19: what older versions left' {
    $flagPath    = Join-Path $ClaudeAppData 'disable-dynamic-memory.flag'
    $policyPath  = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
    $controlPath = 'HKLM:\SYSTEM\CurrentControlSet\Control'
    $taskKey     = "$TaskPath|$BootTaskName"
    $left        = ', set by an older version, is left in place; -Undo removes it'
    function Set-LegacyValues {
        Set-FakeRegValue $controlPath 'ServicesPipeTimeout' ([int]120000)
        Set-FakeRegValue $policyPath 'LocalAccountTokenFilterPolicy' ([int]1)
        Set-FakeRegValue $policyPath 'FilterAdministratorToken' ([int]0)
    }
    function Reset-Step {
        $script:StepResults = New-Object System.Collections.ArrayList
        $script:LogLines = New-Object System.Collections.ArrayList
        $script:VmCalls = New-Object System.Collections.ArrayList
    }

    Reset-State
    Invoke-Step19
    Assert-Case "nothing from older versions: Unchanged, and it says so" ((Get-StepStatus) -eq 'Unchanged' -and (Get-StepNotes) -eq '' -and
        (Test-LogHas '^Nothing left from older versions$')) "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    $script:Tasks[$taskKey] = 'Ready'
    Invoke-Step19
    Assert-Case "the logon task older versions installed is removed: Done" (-not $script:Tasks.ContainsKey($taskKey) -and (Get-StepStatus) -eq 'Done' -and
        (Test-LogHas ('^Removed the logon task ' + $BootTaskName + '$'))) "$(Get-StepStatus): $(Get-StepNotes)"
    $rec = Find-Undo 'LegacyInstall' 'boot-fix task'
    Assert-Case "after it is recorded as what shows an older version ran here" ($null -ne $rec -and [string]$rec.Original -eq '\Claude\ClaudeCoworkBootFix') (Get-UndoKeys)
    $script:UndoRecord = New-Object System.Collections.ArrayList
    Read-PreventUndo
    # Guarded, so that the cases after this one also run against a version
    # without it.
    $trace = $null
    if (Get-Command -Name Get-LegacyInstallTrace -CommandType Function -ErrorAction SilentlyContinue) { $trace = Get-LegacyInstallTrace }
    Assert-Case "so a later -Undo, reading the record, still finds that one did" ($null -ne $trace -and $trace.What -eq 'boot-fix task') (Get-UndoKeys)

    Reset-State
    [System.IO.File]::WriteAllText($flagPath, '1')
    Invoke-Step19
    $rec = Find-Undo 'LegacyInstall' 'Dynamic Memory flag'
    Assert-Case "the Dynamic Memory flag is removed, after it is recorded: Done" (-not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $flagPath) -and
        (Get-StepStatus) -eq 'Done' -and (Test-LogHas ('^Removed ' + [regex]::Escape($flagPath) + '$')) -and
        $null -ne $rec -and [string]$rec.Original -eq $flagPath) "$(Get-StepStatus): $(Get-StepNotes); $(Get-UndoKeys)"

    # A record that cannot be saved: nothing is deleted.
    Reset-State
    $script:Tasks[$taskKey] = 'Ready'
    $null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $UndoFile -Force
    Invoke-Step19
    Assert-Case "the task cannot be recorded: it stays, and the step fails" ($script:Tasks.ContainsKey($taskKey) -and (Get-StepStatus) -eq 'Failed' -and
        @($script:TaskCalls | Where-Object { $_.StartsWith('Unregister ') }).Count -eq 0) "$(Get-StepStatus): $(Get-StepNotes); $($script:TaskCalls -join ' / ')"
    Reset-State
    [System.IO.File]::WriteAllText($flagPath, '1')
    $null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $UndoFile -Force
    Invoke-Step19
    Assert-Case "the flag cannot be recorded: it stays, and the step fails" ((Microsoft.PowerShell.Management\Test-Path -LiteralPath $flagPath) -and
        (Get-StepStatus) -eq 'Failed') "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    Set-LegacyValues
    Invoke-Step19
    $tail = ' matches what older versions set, but no older install was found here, so -Undo leaves it'
    Assert-Case "the three values, no older install found: not called an older version's" ((Get-StepNotes) -eq
        "LocalAccountTokenFilterPolicy = 1$tail | FilterAdministratorToken = 0$tail | ServicesPipeTimeout = 120000$tail" -and
        (Get-StepStatus) -eq 'Unchanged') "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    Set-LegacyValues
    [System.IO.File]::WriteAllText($CsBackupFile, '1')
    Invoke-Step19
    Assert-Case "with an older install found (its CsEnabled record): each is named as an older version's, for -Undo" ((Get-StepNotes) -eq
        "LocalAccountTokenFilterPolicy = 1$left | FilterAdministratorToken = 0$left | ServicesPipeTimeout = 120000$left | CsEnabled (Connected Standby)$left" -and
        (Get-StepStatus) -eq 'Unchanged') "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    $script:Services[$ServiceName] = 'Running'
    Set-FakeRegValue "HKLM:\SYSTEM\CurrentControlSet\Services\$ServiceName" 'FailureActions' (New-FailureBlob -Reset 300 -Pairs @(1, 30000, 1, 60000, 1, 120000))
    Invoke-Step19
    Assert-Case "the restart recovery older versions set on CoworkVMService is named" ((Get-StepNotes) -eq "the restart recovery on CoworkVMService$left" -and
        (Get-StepStatus) -eq 'Unchanged') "$(Get-StepStatus): $(Get-StepNotes)"
    Add-PreventUndoItem -Kind ServiceFailureActions -Key $ServiceName -Original $null
    Reset-Step
    Invoke-Step19
    Assert-Case "but not while it is recorded, for step 1 of -Undo to put back" ((Get-StepNotes) -eq '' -and (Test-LogHas '^Nothing left from older versions$')) (Get-StepNotes)

    Reset-State
    Set-NicFixture
    Set-FakeRegValue "HKLM:\SYSTEM\CurrentControlSet\Enum\$ethPnp\Device Parameters" 'PnPCapabilities' ([int]24)
    Invoke-Step19
    Assert-Case "an adapter an older version touched is named, and only that one" ((Get-StepNotes) -eq "PnPCapabilities with the wake settings on Ethernet$left" -and
        (Get-StepStatus) -eq 'Unchanged') "$(Get-StepStatus): $(Get-StepNotes)"
    Reset-State
    Set-NicFixture
    Set-FakeRegValue "HKLM:\SYSTEM\CurrentControlSet\Enum\$ethPnp\Device Parameters" 'PnPCapabilities' ([int]24)
    $script:NicThrows = $true
    Invoke-Step19
    Assert-Case "the adapters cannot be listed: a warning that says so, and nothing is said to be left" ((Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -eq
        'Could not list the network adapters, so the wake settings older versions turned off on them were not checked: The WMI provider is not loaded' -and
        (Test-LogHas '^\[!\] Could not list the network adapters: The WMI provider is not loaded$') -and -not (Test-LogHas '^Nothing left from older versions$')) (
        "$(Get-StepStatus): $(Get-StepNotes); $($script:LogLines -join ' / ')")

    # Dynamic Memory, where an older version is known to have run: named,
    # and left for -Undo, which Hyper-V lets change it only while the VM is
    # off.
    $dm = 'Dynamic Memory on the VM Claude Cowork, turned off by an older version, is left off; -Undo turns it back on'
    Reset-State
    Add-PreventUndoItem -Kind LegacyInstall -Key 'boot-fix task' -Original "$TaskPath$BootTaskName"
    $script:Vms = @((New-FakeVm 'Claude Cowork' 'Running' $false), (New-FakeVm 'ubuntu-dev' 'Off' $false))
    Invoke-Step19
    Assert-Case "a running Claude VM with Dynamic Memory off is named, with what -Undo needs" ((Get-StepNotes) -eq
        "$dm, once the VM is turned off: it is Running now" -and (Get-StepStatus) -eq 'Unchanged') "$(Get-StepStatus): $(Get-StepNotes)"
    Assert-Case "and setup does not change it" (@($script:VmCalls | Where-Object { $_.StartsWith('Set-VMMemory ') }).Count -eq 0 -and -not $script:Vms[0].DynamicMemoryEnabled) ($script:VmCalls -join ' / ')
    $script:Vms[0].State = 'Off'
    Reset-Step
    Invoke-Step19
    Assert-Case "one that is off is named without that" ((Get-StepNotes) -eq $dm) (Get-StepNotes)
    $script:VmReadFault = $true
    Reset-Step
    Invoke-Step19
    Assert-Case "Hyper-V cannot be read: a warning that says so" ((Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -eq
        'Could not read the Hyper-V VMs, so Dynamic Memory an older version turned off was not checked: The Hyper-V Virtual Machine Management service is not running.') "$(Get-StepStatus): $(Get-StepNotes)"
    Reset-State
    $script:Vms = @((New-FakeVm 'Claude Cowork' 'Off' $false))
    Invoke-Step19
    Assert-Case "no older install found: Hyper-V is not even read" ($script:VmCalls.Count -eq 0 -and (Test-LogHas '^Nothing left from older versions$')) ($script:VmCalls -join ' / ')

    # Hyper-V's cmdlets are not installed: an older install is known, but
    # Dynamic Memory is not looked at and nothing is listed for it.
    foreach ($absent in @(@('Set-VMMemory'), @('Get-VM'), @('Get-VM', 'Set-VMMemory'))) {
        Reset-State
        Add-PreventUndoItem -Kind LegacyInstall -Key 'boot-fix task' -Original "$TaskPath$BootTaskName"
        $script:Vms = @((New-FakeVm 'Claude Cowork' 'Off' $false))
        $script:AbsentCommands = $absent
        Invoke-Step19
        Assert-Case "no $($absent -join ' or '): no VM cmdlet is called, nothing is listed, and nothing fails" ($script:VmCalls.Count -eq 0 -and
            (Get-StepStatus) -eq 'Unchanged' -and (Get-StepNotes) -eq '' -and (Test-LogHas '^Nothing left from older versions$')) (
            "$(Get-StepStatus): $(Get-StepNotes); $($script:VmCalls -join ' / ')")
    }

    # The NAT rule CoworkNAT older versions created, and their monitor
    # re-created: listed for -Undo, and left in place, whether or not an
    # older install is found. Another rule is not theirs, a PC without
    # WinNAT's cmdlets is not asked, and a list that cannot be read is a
    # warning that says so.
    $natLeft = "the NAT rule CoworkNAT$left"
    $natUnread = 'Could not read the NAT rules, so the rule CoworkNAT older versions created was not checked: Access denied'
    function Get-NatRemoves { return @($script:NatCalls | Where-Object { $_.StartsWith('Remove-NetNat') }).Count }
    Reset-State
    $script:Nats = @((New-FakeNat 'WSLNat'), (New-FakeNat 'CoworkNAT'))
    Invoke-Step19
    Assert-Case "a CoworkNAT rule is listed for -Undo, and setup leaves it" ((Get-StepNotes) -eq $natLeft -and (Get-StepStatus) -eq 'Unchanged' -and
        (Test-LogHas '^Left in place: the NAT rule CoworkNAT \(-Undo removes it\)$') -and $script:Nats.Count -eq 2 -and (Get-NatRemoves) -eq 0) (
        "$(Get-StepStatus): $(Get-StepNotes); $($script:NatCalls -join ' / ')")
    Reset-State
    $script:Nats = @((New-FakeNat 'WSLNat'))
    Invoke-Step19
    Assert-Case "with no CoworkNAT rule, another one is not named" ((Get-StepNotes) -eq '' -and (Test-LogHas '^Nothing left from older versions$') -and
        ($script:NatCalls -join ' / ') -eq 'Get-NetNat CoworkNAT') "$(Get-StepNotes); $($script:NatCalls -join ' / ')"
    Reset-State
    $script:Nats = @((New-FakeNat 'CoworkNAT'))
    $script:AbsentCommands = @('Get-NetNat')
    Invoke-Step19
    Assert-Case "no WinNAT cmdlets: no NAT rule is read, and none is named" ($script:NatCalls.Count -eq 0 -and (Get-StepNotes) -eq '' -and
        (Test-LogHas '^Nothing left from older versions$')) "$(Get-StepNotes); $($script:NatCalls -join ' / ')"
    Reset-State
    $script:Nats = @((New-FakeNat 'CoworkNAT'))
    $script:NatReadFault = $true
    Invoke-Step19
    Assert-Case "the NAT rules cannot be read: a warning that says so" ((Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -eq $natUnread -and
        (Test-LogHas '^\[!\] Could not read the NAT rules: Access denied$')) "$(Get-StepStatus): $(Get-StepNotes)"
    foreach ($nw in @(@{ Why = 'Invalid class'; Status = 5; Code = 0x80041010L }, @{ Why = 'Provider load failure'; Status = 1; Code = 0x80041013L })) {
        Reset-State
        $script:Nats = @((New-FakeNat 'CoworkNAT'))
        $script:NatReadCim = New-FakeCimError -Status $nw.Status -Code $nw.Code -Message "$($nw.Why) "
        Invoke-Step19
        Assert-Case "Get-NetNat fails with $($nw.Why): no WinNAT, so no warning and no rule named" ((Get-StepStatus) -eq 'Unchanged' -and
            (Get-StepNotes) -eq '' -and (Test-LogHas '^Nothing left from older versions$') -and -not (Test-LogHas 'NAT rule')) (
            "$(Get-StepStatus): $(Get-StepNotes); $($script:LogLines -join ' / ')")
    }
    Reset-State
    $script:Nats = @((New-FakeNat 'CoworkNAT'))
    $script:NatReadCim = New-FakeCimError -Status 2 -Code 0x80041003L -Message 'Access denied ' -Category PermissionDenied
    Invoke-Step19
    Assert-Case "Get-NetNat fails with a CIM Access denied: still a warning that says so" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq "$natUnread " -and (Test-LogHas '^\[!\] Could not read the NAT rules: Access denied $')) "$(Get-StepStatus): $(Get-StepNotes)"
    Reset-State
    $script:Nats = @((New-FakeNat 'CoworkNAT'))
    $script:NatRemoveFault = $true
    Invoke-Step19
    Assert-Case "a rule that cannot be removed is no matter to setup, which never removes one" ((Get-StepNotes) -eq $natLeft -and
        (Get-StepStatus) -eq 'Unchanged' -and (Get-NatRemoves) -eq 0) "$(Get-StepStatus): $(Get-StepNotes); $($script:NatCalls -join ' / ')"
}

# =====================================================================
# Step 13 adds the Defender exclusions that are missing and records each
# one first, so that -Undo removes those it added and no others. Where a
# policy hides Defender's lists from administrators, each list comes back
# as one "N/A" entry. Until 6.3.0 that read as nothing excluded: every
# exclusion looked missing, so step 13 added and recorded ones that were
# already there, which -Undo then removed, and the process check that read
# the list again reported the ones it had added as failed.
Invoke-Section 'Setup step 13: Defender exclusions, and lists Defender hides' {
    $paths = @($ClaudeDataRoot, (Join-Path $env:ProgramFiles 'Hyper-V'), "$env:SystemRoot\System32\vmwp.exe", "$env:SystemRoot\System32\vmms.exe")
    $procs = @('vmwp.exe', 'vmms.exe', 'vmcompute.exe', "$ServiceExe.exe")
    $pathSaid = 'Defender hides its exclusion list from administrators on this machine, so this step cannot tell which path exclusions exist; none were added'
    $procSaid = 'Defender hides its exclusion list from administrators on this machine, so this step cannot tell which process exclusions exist; none were added'
    $wantPaths = (@($paths | ForEach-Object { "Path $_" }) -join ' / ')
    $wantProcs = (@($procs | ForEach-Object { "Process $_" }) -join ' / ')
    function Get-RecordKeys {
        param([string]$Kind)
        return (@($script:UndoRecord | Where-Object { $_.Kind -eq $Kind } | ForEach-Object { [string]$_.Key }) -join ' / ')
    }
    function Reset-Defender {
        Reset-State
        $script:AvProducts = @([pscustomobject]@{ displayName = 'Windows Defender' })
    }

    Reset-Defender
    $script:MpHidePath = $true
    $script:MpHideProcess = $true
    Invoke-Step13
    Assert-Case "both lists hidden: nothing is added and nothing recorded" ($script:MpAdds.Count -eq 0 -and $script:UndoRecord.Count -eq 0 -and
        -not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $UndoFile)) "$($script:MpAdds -join ' / '); $(Get-UndoKeys)"
    Assert-Case "and the step says so for each list, as a Warning" ((Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -eq "$pathSaid | $procSaid" -and
        (Test-LogHas ('^\[!\] ' + [regex]::Escape($pathSaid) + '$')) -and (Test-LogHas ('^\[!\] ' + [regex]::Escape($procSaid) + '$'))) "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-Defender
    Invoke-Step13
    Assert-Case "nothing excluded yet: the four paths, then the four processes, are added" (($script:MpAdds -join ' / ') -eq "$wantPaths / $wantProcs") ($script:MpAdds -join ' / ')
    Assert-Case "each one recorded for -Undo, and the step is Done" ((Get-RecordKeys 'DefenderPath') -eq ($paths -join ' / ') -and
        (Get-RecordKeys 'DefenderProcess') -eq ($procs -join ' / ') -and $script:UndoRecord.Count -eq 8 -and (Get-StepStatus) -eq 'Done' -and
        (Get-StepNotes) -eq '' -and (Test-LogHas '^  All process exclusions verified$')) "$(Get-StepStatus): $(Get-StepNotes); $(Get-UndoKeys)"
    $script:UndoRecord = New-Object System.Collections.ArrayList
    Read-PreventUndo
    Assert-Case "and the record is on disk" ($script:UndoRecord.Count -eq 8) (Get-UndoKeys)

    Reset-Defender
    foreach ($x in $paths) { [void]$script:MpPref.Path.Add($x) }
    foreach ($x in $procs) { [void]$script:MpPref.Process.Add($x) }
    Invoke-Step13
    Assert-Case "every exclusion already there: nothing added or recorded, Unchanged" ($script:MpAdds.Count -eq 0 -and $script:UndoRecord.Count -eq 0 -and
        (Get-StepStatus) -eq 'Unchanged' -and (Test-LogHas '^Defender path exclusions already configured$') -and
        (Test-LogHas '^Defender process exclusions already configured$')) "$(Get-StepStatus): $($script:MpAdds -join ' / '); $(Get-UndoKeys)"

    Reset-Defender
    $script:MpHidePath = $true
    Invoke-Step13
    Assert-Case "the path list hidden: only the processes are added, and only they are recorded" (($script:MpAdds -join ' / ') -eq $wantProcs -and
        (Get-RecordKeys 'DefenderPath') -eq '' -and (Get-RecordKeys 'DefenderProcess') -eq ($procs -join ' / ')) "$($script:MpAdds -join ' / '); $(Get-UndoKeys)"
    Assert-Case "a Warning for the hidden list, and the processes counted as added" ((Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -eq $pathSaid -and
        (Test-LogHas '^  All process exclusions verified$')) "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-Defender
    $script:MpHideProcess = $true
    Invoke-Step13
    Assert-Case "the process list hidden: only the paths are added, and only they are recorded" (($script:MpAdds -join ' / ') -eq $wantPaths -and
        (Get-RecordKeys 'DefenderPath') -eq ($paths -join ' / ') -and (Get-RecordKeys 'DefenderProcess') -eq '') "$($script:MpAdds -join ' / '); $(Get-UndoKeys)"
    Assert-Case "and a Warning that names the process list" ((Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -eq $procSaid) "$(Get-StepStatus): $(Get-StepNotes)"

    # The test itself, last: a version without it stops here, after the
    # cases above have run.
    $na = 'N/A: Must be an administrator to view exclusions'
    Assert-Case "Test-DefenderListHidden: a list that is only the N/A entry is hidden" (Test-DefenderListHidden -List @($na))
    Assert-Case "an empty list is empty, not hidden" (-not (Test-DefenderListHidden -List @()) -and -not (Test-DefenderListHidden -List $null))
    Assert-Case "a list with a real entry beside the N/A one is not hidden" (-not (Test-DefenderListHidden -List @($na, 'C:\Data')) -and
        -not (Test-DefenderListHidden -List @('C:\Data', $na)))
}

# =====================================================================
# Setup step 11 checks whether the Claude data folder is on an external or
# a network drive, and the VM cache on a network drive. Get-Volume,
# Get-Partition and Get-Disk say ObjectNotFound for what is not there,
# which leaves no disk to judge; any other error, and any error from
# Win32_LogicalDisk, is a check that could not be made, a Warning that
# names it and why. Until 6.3.0 every read used SilentlyContinue inside a
# catch that kept nothing, so a check that failed read as a local drive,
# and the network check ran only for a drive Get-Volume found, which a
# network drive never is.
Invoke-Section 'Setup step 11: where the Claude data folder is, and the checks that cannot be made' {
    $sbDrive = Split-Path $sandbox -Qualifier
    $dl = 'Q'
    if ($sbDrive -eq 'Q:') { $dl = 'R' }
    $ClaudeDataRoot = "${dl}:\Claude-Data"
    $vmCache = Join-Path $sandbox 'vm_bundles'
    $null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $vmCache -Force
    $ok = '^Storage location OK \(local drive\)$'
    $ext = "Could not check whether the drive ${dl}:\, which holds the Claude data folder, is an external drive: "
    $net = "Could not check whether the drive ${dl}:\, which holds the Claude data folder, is a network drive: "
    function Reset-Storage {
        param([string]$Bus = 'NVMe')
        Reset-State
        $script:Drives = @{ $dl = @{ Disk = 1; Bus = $Bus } }
        $script:LogicalDisks = @{ "${dl}:" = 3 }
    }
    function Get-Reads { return ($script:StorageCalls -join ' / ') }
    function Test-Unread {
        param([string]$Said)
        return ((Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -eq $Said -and (Test-LogHas ('^' + [regex]::Escape("[!] $Said") + '$')) -and
            -not (Test-LogHas $ok))
    }

    Reset-Storage
    Invoke-Step11
    Assert-Case "a local drive: OK, Unchanged, after each check was made" ((Get-StepStatus) -eq 'Unchanged' -and (Get-StepNotes) -eq '' -and (Test-LogHas $ok) -and
        (Get-Reads) -eq "Get-Volume $dl / Get-Partition $dl / Get-Disk 1 / Win32_LogicalDisk ${dl}:") "$(Get-StepStatus): $(Get-StepNotes); $(Get-Reads)"

    Reset-Storage -Bus 'USB'
    Invoke-Step11
    Assert-Case "a USB drive: a Warning that says so" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq 'The Claude data folder is on an external USB drive, use a local SSD instead' -and -not (Test-LogHas $ok)) "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    $script:LogicalDisks = @{ "${dl}:" = 4 }
    Invoke-Step11
    Assert-Case "a network drive, which Get-Volume does not list: a Warning that says so" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq 'The Claude data folder is on a network drive, VirtioFS requires local storage' -and
        (Get-Reads) -eq "Get-Volume $dl / Win32_LogicalDisk ${dl}:") "$(Get-StepStatus): $(Get-StepNotes); $(Get-Reads)"

    Reset-State
    Invoke-Step11
    Assert-Case "a drive letter nothing lists: not found is no failure, and OK" ((Get-StepStatus) -eq 'Unchanged' -and (Get-StepNotes) -eq '' -and
        (Test-LogHas $ok)) "$(Get-StepStatus): $(Get-StepNotes); $(Get-Reads)"

    Reset-Storage
    $script:Drives[$dl] = $null
    Invoke-Step11
    Assert-Case "a volume with no partition behind it: no disk to judge, and the network check still made" ((Get-StepStatus) -eq 'Unchanged' -and
        (Get-StepNotes) -eq '' -and (Get-Reads) -eq "Get-Volume $dl / Get-Partition $dl / Win32_LogicalDisk ${dl}:") "$(Get-StepStatus): $(Get-StepNotes); $(Get-Reads)"

    foreach ($r in @(
        @{ Fault = 'Volume'; Said = 'The RPC server is unavailable.'; Reads = "Get-Volume $dl / Win32_LogicalDisk ${dl}:" },
        @{ Fault = 'Partition'; Said = 'Access denied'; Reads = "Get-Volume $dl / Get-Partition $dl / Win32_LogicalDisk ${dl}:" },
        @{ Fault = 'Disk'; Said = 'The device is not ready.'; Reads = "Get-Volume $dl / Get-Partition $dl / Get-Disk 1 / Win32_LogicalDisk ${dl}:" }
    )) {
        Reset-Storage -Bus 'USB'
        $script:StorageFault = $r.Fault
        Invoke-Step11
        Assert-Case "Get-$($r.Fault) cannot be read: a Warning that names the external drive check and why, and the network check still made" (
            (Test-Unread "$ext$($r.Said)") -and (Get-Reads) -eq $r.Reads) "$(Get-StepStatus): $(Get-StepNotes); $(Get-Reads)"
    }

    Reset-Storage
    $script:LogicalDiskFault = @("${dl}:")
    Invoke-Step11
    Assert-Case "Win32_LogicalDisk cannot be read: a Warning that names the network drive check and why" (
        Test-Unread "${net}The RPC server is unavailable.") "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-Storage
    $script:StorageFault = 'Volume'
    $script:LogicalDiskFault = @("${dl}:")
    Invoke-Step11
    Assert-Case "both checks cannot be made: a Warning for each, in order" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq "${ext}The RPC server is unavailable. | ${net}The RPC server is unavailable.") "$(Get-StepStatus): $(Get-StepNotes)"

    $ClaudeDataRoot = '\\server\share\Claude'
    Reset-Storage
    Invoke-Step11
    $said = "Could not tell which drive the Claude data folder \\server\share\Claude is on, so whether it is an external or a network drive was not checked: "
    Assert-Case "a data folder on a network share, with no drive letter: a Warning that neither check was made" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes).StartsWith($said) -and (Get-StepNotes).Length -gt $said.Length -and (Get-Reads) -eq '' -and -not (Test-LogHas $ok)) (
        "$(Get-StepStatus): $(Get-StepNotes); $(Get-Reads)")
    $ClaudeDataRoot = "${dl}:\Claude-Data"

    Reset-Storage
    $script:ClaudeEnv.VmCachePath = $vmCache
    $script:LogicalDisks[$sbDrive] = 3
    Invoke-Step11
    Assert-Case "the VM cache on a local drive: OK, after its drive was checked" ((Get-StepStatus) -eq 'Unchanged' -and (Get-StepNotes) -eq '' -and
        (Get-Reads).EndsWith("Win32_LogicalDisk $sbDrive")) "$(Get-StepStatus): $(Get-StepNotes); $(Get-Reads)"

    Reset-Storage
    $script:ClaudeEnv.VmCachePath = $vmCache
    $script:LogicalDisks[$sbDrive] = 4
    Invoke-Step11
    Assert-Case "the VM cache on a network drive: a Warning that says so" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq 'VM cache is on a network drive, this will cause failures') "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-Storage
    $script:ClaudeEnv.VmCachePath = $vmCache
    $script:LogicalDiskFault = @($sbDrive)
    Invoke-Step11
    Assert-Case "the VM cache's drive cannot be read: a Warning that names it and why" (
        Test-Unread "Could not check whether the VM cache $vmCache is on a network drive: The RPC server is unavailable.") "$(Get-StepStatus): $(Get-StepNotes)"

    # 6.3.1: until then the cache folder was looked for with Test-Path, which
    # answers $false for a folder that cannot be read as for one that is not
    # there: its drive was not checked, and "Storage location OK" was said.
    Reset-Storage
    $script:ClaudeEnv.VmCachePath = $vmCache
    $script:LogicalDisks[$sbDrive] = 3
    $script:ItemReadFault = $vmCache
    Invoke-Step11
    Assert-Case "the VM cache folder cannot be read: a Warning that names it and why, and its drive checked all the same" (
        (Test-Unread "Could not read the VM cache folder ${vmCache}: Access is denied") -and (Get-Reads).EndsWith("Win32_LogicalDisk $sbDrive")) (
        "$(Get-StepStatus): $(Get-StepNotes); $(Get-Reads)")
    Reset-State
}

# =====================================================================
# Setup step 13 reads the antivirus products registered with Windows
# Security Center. Windows Server has none: there its namespace is
# missing, which is no failure, known by the CIM codes and never by the
# message. Any other error is a Warning, since an antivirus that is not
# Defender may then go unreported. Until 6.3.0 every error read as no
# product registered.
Invoke-Section 'Setup step 13: the registered antivirus products cannot be read' {
    $said = 'Could not read the antivirus products registered with Windows, so one that is not Defender may go unreported: '
    foreach ($c in @(
        @{ Why = 'a namespace that is not there'; Status = 3; Code = 0x8004100EL; Message = 'Invalid namespace ' },
        @{ Why = 'a class that is not there'; Status = 5; Code = 0x80041010L; Message = 'Invalid class ' },
        @{ Why = 'the namespace status alone'; Status = 3; Code = -1L; Message = 'Invalid namespace ' },
        @{ Why = 'the WMI code for a class that is not there alone'; Status = 1; Code = 0x80041010L; Message = 'Failed ' }
    )) {
        Reset-State
        $script:AvReadCim = New-FakeCimError -Status $c.Status -Code $c.Code -Message $c.Message
        Invoke-Step13
        Assert-Case "$($c.Why): nothing said of it, and no antivirus detected" ((Get-StepStatus) -eq 'Skipped' -and (Get-StepNotes) -eq '' -and
            (Test-LogHas '^No antivirus product detected$')) "$(Get-StepStatus): $(Get-StepNotes)"
    }

    foreach ($c in @(
        @{ Why = 'access denied'; Err = (New-FakeCimError -Status 2 -Code 0x80041003L -Message 'Access denied ' -Category PermissionDenied); Text = 'Access denied ' },
        @{ Why = 'not supported'; Err = (New-FakeCimError -Status 7 -Code 0x8004100CL -Message 'Not supported ' -Category NotImplemented); Text = 'Not supported ' },
        @{ Why = 'a message naming the namespace, with codes that do not'; Err = (New-FakeCimError -Status 1 -Code 0x80041001L -Message 'Invalid namespace '); Text = 'Invalid namespace ' }
    )) {
        Reset-State
        $script:AvReadCim = $c.Err
        Invoke-Step13
        Assert-Case "the list cannot be read ($($c.Why)): a Warning that says so and why" ((Get-StepStatus) -eq 'Warning' -and
            (Get-StepNotes) -eq "$said$($c.Text)" -and (Test-LogHas ('^' + [regex]::Escape("[!] $said$($c.Text)") + '$'))) "$(Get-StepStatus): $(Get-StepNotes)"
        Assert-Case "and it is not said that no antivirus was detected" (-not (Test-LogHas '^No antivirus product detected$') -and
            (Test-LogHas '^No antivirus product found among the running processes$')) ($script:LogLines -join ' / ')
    }

    Reset-State
    $script:AvReadFault = $true
    Invoke-Step13
    Assert-Case "an error that is not a CIM one: a Warning that says so and why" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq "${said}The RPC server is unavailable.") "$(Get-StepStatus): $(Get-StepNotes)"

    # Defender found among the running processes: its exclusions are still
    # made, and the Warning stays.
    Reset-State
    $script:AvReadFault = $true
    $script:Procs = @(New-FakeProcess 900 'MsMpEng')
    Invoke-Step13
    Assert-Case "the list cannot be read, Defender running: the exclusions are made, and the Warning stays" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq "${said}The RPC server is unavailable." -and $script:MpAdds.Count -eq 8 -and
        (Test-LogHas '^Detected: Windows Defender$')) "$(Get-StepStatus): $(Get-StepNotes); $($script:MpAdds.Count) added"
}

# =====================================================================
# Test-CimClassMissingError: a namespace or a class that is not there, by
# the CIM status or the WMI code, never by the message.
Invoke-Section 'Test-CimClassMissingError' {
    function New-CimRecord {
        param([uint32]$Status, [long]$Code, [string]$Message = 'Failed ')
        $f = New-FakeCimError -Status $Status -Code $Code -Message $Message
        return [System.Management.Automation.ErrorRecord]::new($f.Exception, $f.Id, $f.Category, $null)
    }
    foreach ($c in @(
        @{ S = 3; C = 0x8004100EL; Want = $true }, @{ S = 5; C = 0x80041010L; Want = $true },
        @{ S = 3; C = -1L; Want = $true }, @{ S = 5; C = -1L; Want = $true },
        @{ S = 1; C = 0x8004100EL; Want = $true }, @{ S = 1; C = 0x80041010L; Want = $true },
        @{ S = 2; C = 0x80041003L; Want = $false }, @{ S = 7; C = 0x8004100CL; Want = $false },
        @{ S = 6; C = 0x80041002L; Want = $false }, @{ S = 1; C = -1L; Want = $false }
    )) {
        $got = Test-CimClassMissingError -ErrorRecord (New-CimRecord -Status $c.S -Code $c.C)
        $codeText = 'no WMI code'
        if ($c.C -ge 0) { $codeText = '0x{0:X8}' -f $c.C }
        Assert-Case "status $($c.S), $codeText`: $($c.Want)" ($got -eq $c.Want) "got $got"
    }
    Assert-Case "the message is never read: status 1 with no code saying Invalid namespace is not missing" (
        -not (Test-CimClassMissingError -ErrorRecord (New-CimRecord -Status 1 -Code (-1L) -Message 'Invalid namespace ')))
    $plain = [System.Management.Automation.ErrorRecord]::new([System.Management.Automation.RuntimeException]::new('Invalid namespace'), 'x', 'MetadataError', $null)
    Assert-Case "an error that is not a CIM one is not missing" (-not (Test-CimClassMissingError -ErrorRecord $plain))
    $inner = (New-FakeCimError -Status 3 -Code 0x8004100EL -Message 'Invalid namespace ').Exception
    $wrapped = [System.Management.Automation.ErrorRecord]::new([System.Management.Automation.RuntimeException]::new('wrapped', $inner), 'x', 'MetadataError', $null)
    Assert-Case "a CIM error inside another exception is found" (Test-CimClassMissingError -ErrorRecord $wrapped)
}

# =====================================================================
# -Undo step 3 stops the health monitor and removes the two scheduled
# tasks and the old watchdog script, and checks that each one went. Until
# 6.3.0 a monitor that did not stop was logged as stopped, every error from
# the scheduler as a task that was not there, and a file that stayed as
# removed, so -Undo could finish with the monitor still installed.
Invoke-Section '-Undo step 3: the health monitor, its tasks and the old watchdog script' {
    $wd = Join-Path $ClaudeAppData 'cowork-watchdog.ps1'
    $wdTask = "$TaskPath|$TaskName"
    $bootTask = "$TaskPath|$BootTaskName"
    function Set-Monitors {
        $script:PsProcs = @(
            [pscustomobject]@{ ProcessId = 4242; CommandLine = 'powershell.exe -NoProfile -File "C:\ClaudeFix\Watch-ClaudeHealth.ps1" -Quiet' },
            [pscustomobject]@{ ProcessId = 4343; CommandLine = 'powershell.exe -NoProfile -File "C:\Tools\Backup.ps1"' }
        )
        $script:Procs = @((New-FakeProcess 4242 'powershell'), (New-FakeProcess 4343 'powershell'))
    }
    function Get-ProcIds { return (@($script:Procs | ForEach-Object { [string]$_.Id }) -join ',') }

    Reset-State
    Set-Monitors
    Invoke-UndoStep3
    Assert-Case "the health monitor is stopped, and only it" (($script:ProcCalls -join ' / ') -eq 'Stop-Process 4242' -and (Get-ProcIds) -eq '4343' -and
        (Test-LogHas '^Stopped running health monitor \(PID 4242\)$') -and (Get-StepStatus) -eq 'Done') "$($script:ProcCalls -join ' / '); left $(Get-ProcIds); $(Get-StepStatus)"
    Assert-Case "and with nothing else there, the tasks and the script are reported absent" ((Test-LogHas '^No tasks to remove$') -and
        (Test-LogHas ('^' + $TaskName + ' not found$')) -and (Test-LogHas ('^' + $BootTaskName + ' not found$')) -and (Get-StepNotes) -eq '') ($script:LogLines -join ' / ')

    $end = Get-UndoEnd
    Assert-Case "and -Undo ends UNDO COMPLETE with exit code 0" ($end.Box -eq 'UNDO COMPLETE' -and -not $end.Unread -and $end.Exit -eq 0) $end.Text

    # A process list that cannot be read is said, and -Undo ends incomplete.
    # Read as no monitor running, a monitor that was went on running.
    Reset-State
    Set-Monitors
    $script:PsProcsFault = $true
    Invoke-UndoStep3
    Assert-Case "the running processes cannot be listed: a warning that says so, and no process is stopped" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq 'Could not list the running processes, so a health monitor still running was not stopped: The RPC server is unavailable.' -and
        @($script:ProcCalls).Count -eq 0 -and (Test-LogHas ('^' + [regex]::Escape(
        '[!] Could not list the running processes, so a running health monitor was not looked for: The RPC server is unavailable.') + '$'))) (
        "$(Get-StepStatus): $(Get-StepNotes); $($script:ProcCalls -join ' / ')")
    $end = Get-UndoEnd
    Assert-Case "and -Undo ends UNDO INCOMPLETE, saying to run it again, with exit code 1" ($end.Box -eq 'UNDO INCOMPLETE' -and $end.Unread -and $end.Exit -eq 1) $end.Text

    Reset-State
    Set-Monitors
    $script:StopProcFault = $true
    Invoke-UndoStep3
    Assert-Case "a monitor that cannot be stopped: Failed, with the reason" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -match '^Could not stop the health monitor \(PID 4242\): Cannot stop process' -and -not (Test-LogHas '^Stopped running')) "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    Set-Monitors
    $script:StopProcLinger = $true
    Invoke-UndoStep3
    Assert-Case "a monitor still running after it was stopped: Failed, after waiting 2 seconds" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq 'Could not stop the health monitor (PID 4242): it is still running' -and $script:FakeClockMs -eq 2000 -and
        -not (Test-LogHas '^Stopped running')) "$(Get-StepStatus): $(Get-StepNotes); clock $($script:FakeClockMs)"

    # Listed, but gone by the time it is stopped: the stop's error says so.
    # It ended by itself, so it is not said to have been stopped, and
    # nothing failed.
    Reset-State
    Set-Monitors
    $script:Procs = @(New-FakeProcess 4343 'powershell')
    Invoke-UndoStep3
    Assert-Case "a monitor that ended before it was stopped: said to have ended, not stopped, and nothing failed" ((Get-StepStatus) -eq 'Unchanged' -and
        (Get-StepNotes) -eq '' -and (Test-LogHas '^The health monitor that was running \(PID 4242\) has ended$') -and
        -not (Test-LogHas '^Stopped running')) "$(Get-StepStatus): $(Get-StepNotes); $($script:LogLines -join ' / ')"

    Reset-State
    $script:Tasks[$wdTask] = 'Running'
    $script:Tasks[$bootTask] = 'Ready'
    Invoke-UndoStep3
    Assert-Case "both tasks are removed: Done" ($script:Tasks.Count -eq 0 -and (Get-StepStatus) -eq 'Done' -and (Test-LogHas ('^' + $TaskName + ' removed$')) -and
        (Test-LogHas ('^' + $BootTaskName + ' removed$')) -and -not (Test-LogHas '^No tasks to remove$')) "$(Get-StepStatus): $($script:LogLines -join ' / ')"

    Reset-State
    $script:Tasks[$wdTask] = 'Ready'
    Invoke-UndoStep3
    Assert-Case "one task there: it is removed, and the other is not found" (-not $script:Tasks.ContainsKey($wdTask) -and (Get-StepStatus) -eq 'Done' -and
        (Test-LogHas ('^' + $BootTaskName + ' not found$')) -and -not (Test-LogHas '^No tasks to remove$')) "$(Get-StepStatus): $($script:LogLines -join ' / ')"

    Reset-State
    $script:Tasks[$wdTask] = 'Ready'
    $script:TaskQueryFault = $true
    Invoke-UndoStep3
    Assert-Case "a scheduler that cannot be read: Failed for each task, none called not found" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -match ('^The scheduled task ' + $TaskName + ': it could not be read \(The RPC server is unavailable') -and
        (Get-StepNotes) -match ('\| The scheduled task ' + $BootTaskName + ': it could not be read \(') -and
        -not (Test-LogHas 'not found') -and -not (Test-LogHas '^No tasks to remove$')) "$(Get-StepStatus): $(Get-StepNotes)"
    Assert-Case "and nothing is unregistered" (@($script:TaskCalls | Where-Object { $_.StartsWith('Unregister ') }).Count -eq 0 -and
        $script:Tasks.ContainsKey($wdTask)) ($script:TaskCalls -join ' / ')

    Reset-State
    $script:Tasks[$wdTask] = 'Ready'
    $script:UnregisterFault = $true
    Invoke-UndoStep3
    Assert-Case "a task the scheduler refuses to remove: Failed, with its reason" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq "The scheduled task ${TaskName}: it could not be removed (Access is denied.)" -and $script:Tasks.ContainsKey($wdTask) -and
        -not (Test-LogHas ('^' + $TaskName + ' (removed|not found)$'))) "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    $script:Tasks[$wdTask] = 'Ready'
    $script:UnregisterNoop = $true
    Invoke-UndoStep3
    Assert-Case "a task still there after it was removed: Failed" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq "The scheduled task ${TaskName}: it is still there after it was removed" -and
        -not (Test-LogHas ('^' + $TaskName + ' removed$'))) "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    [System.IO.File]::WriteAllText($wd, '# old')
    Invoke-UndoStep3
    Assert-Case "the old watchdog script is removed: Done" (-not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $wd) -and (Get-StepStatus) -eq 'Done' -and
        (Test-LogHas ('^Removed old watchdog script: ' + [regex]::Escape($wd) + '$'))) "$(Get-StepStatus): $($script:LogLines -join ' / ')"

    Reset-State
    [System.IO.File]::WriteAllText($wd, '# old')
    $script:FileRemoveFault = '*cowork-watchdog.ps1'
    Invoke-UndoStep3
    Assert-Case "a script that cannot be deleted: Failed, with the reason, and it is kept" ((Microsoft.PowerShell.Management\Test-Path -LiteralPath $wd) -and
        (Get-StepStatus) -eq 'Failed' -and (Get-StepNotes) -eq "Could not remove ${wd}: Access to the path is denied." -and
        -not (Test-LogHas '^Removed old watchdog script')) "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    [System.IO.File]::WriteAllText($wd, '# old')
    $script:FileRemoveNoop = '*cowork-watchdog.ps1'
    Invoke-UndoStep3
    Assert-Case "a script still there after it was deleted: Failed" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq "Could not remove ${wd}: it is still there" -and -not (Test-LogHas '^Removed old watchdog script')) "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    Invoke-UndoStep3
    Assert-Case "nothing there at all: Unchanged" ((Get-StepStatus) -eq 'Unchanged' -and (Get-StepNotes) -eq '' -and @($script:ProcCalls).Count -eq 0) "$(Get-StepStatus): $(Get-StepNotes)"
}

# =====================================================================
# -Undo step 4 deletes the two Fix Claude Desktop shortcuts. Until 6.3.0
# each one it tried was logged as removed, whether or not it went.
Invoke-Section '-Undo step 4: the Fix Claude Desktop shortcuts' {
    $menuDir = Join-Path $fakeAppData 'Microsoft\Windows\Start Menu\Programs'
    $desk = Join-Path $fakeDesktop 'Fix Claude Desktop.lnk'
    $menu = Join-Path $menuDir 'Fix Claude Desktop.lnk'
    function Set-Shortcuts {
        $null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $menuDir -Force
        [System.IO.File]::WriteAllText($desk, 'lnk')
        [System.IO.File]::WriteAllText($menu, 'lnk')
    }

    Reset-State
    Set-Shortcuts
    Invoke-UndoStep4
    Assert-Case "both shortcuts are removed: Done" (-not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $desk) -and
        -not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $menu) -and (Get-StepStatus) -eq 'Done' -and
        (Test-LogHas ('^Removed: ' + [regex]::Escape($desk) + '$')) -and (Test-LogHas ('^Removed: ' + [regex]::Escape($menu) + '$')) -and
        -not (Test-LogHas '^No shortcuts found$')) "$(Get-StepStatus): $($script:LogLines -join ' / ')"

    Reset-State
    Set-Shortcuts
    $script:FileRemoveFault = '*Start Menu*'
    Invoke-UndoStep4
    Assert-Case "one that cannot be deleted: Failed, it is kept, and the other is still removed" ((Microsoft.PowerShell.Management\Test-Path -LiteralPath $menu) -and
        -not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $desk) -and (Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq "Could not remove ${menu}: Access to the path is denied." -and
        -not (Test-LogHas ('^Removed: ' + [regex]::Escape($menu) + '$'))) "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    Set-Shortcuts
    $script:FileRemoveNoop = '*Start Menu*'
    Invoke-UndoStep4
    Assert-Case "one still there after it was deleted: Failed" ((Get-StepStatus) -eq 'Failed' -and (Get-StepNotes) -eq "Could not remove ${menu}: it is still there") "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    Invoke-UndoStep4
    Assert-Case "none there: it says so, Unchanged" ((Get-StepStatus) -eq 'Unchanged' -and (Test-LogHas '^No shortcuts found$')) "$(Get-StepStatus): $($script:LogLines -join ' / ')"
}

# =====================================================================
# -Undo step 5 removes the LaunchClaudeAdmin task, the launchers and their
# shortcuts, and then the \Claude\ task folder, once it is empty. Every
# removal is checked. Until 6.3.0 the folder was left behind, a task the
# scheduler refused to delete was reported as not found, and every file
# was logged as removed. The RUNASADMIN marks other software sets on
# Claude.exe are not this script's, and stay.
Invoke-Section '-Undo step 5: the launchers and the task folder' {
    $adminTask = "$TaskPath|LaunchClaudeAdmin"
    $layers = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers'
    $claudeExe = 'C:\Program Files\Claude\Claude.exe'
    $files = @(
        (Join-Path $fakeDesktop 'Claude (Admin).lnk'),
        (Join-Path $ClaudeAppData 'Launch-Claude-Admin.cmd'),
        (Join-Path $ClaudeAppData 'Launch-Claude-Admin.ps1'),
        (Join-Path $ClaudeAppData 'Launch-Claude-Admin.vbs'),
        (Join-Path $ClaudeAppData 'Launch-Claude-Admin-Trigger.vbs'),
        (Join-Path $ClaudeAppData 'Launch-Claude-NoGPU.cmd'),
        (Join-Path $ClaudeAppData 'Launch-Claude-NoGPU.ps1'),
        (Join-Path $fakeDesktop 'Claude (no GPU).lnk'),
        $ClaudeIconFile
    )
    function Set-Launchers {
        param([switch]$NoFolder)
        $script:Tasks[$adminTask] = 'Ready'
        if (-not $NoFolder) { [void]$script:TaskFolders.Add('\Claude') }
        foreach ($f in $files) { [System.IO.File]::WriteAllText($f, 'x') }
        Set-FakeRegValue $layers $claudeExe '~ RUNASADMIN'
    }
    function Get-FilesLeft { return @($files | Where-Object { Microsoft.PowerShell.Management\Test-Path -LiteralPath $_ }).Count }
    function Get-FolderCalls { return (@($script:TaskCalls | Where-Object { $_.StartsWith('DeleteFolder ') }) -join ' / ') }

    Reset-State
    Set-Launchers
    Invoke-UndoStep5
    Assert-Case "the task, the nine files and the empty task folder are removed: Done" (-not $script:Tasks.ContainsKey($adminTask) -and (Get-FilesLeft) -eq 0 -and
        $script:TaskFolders.Count -eq 0 -and (Get-FolderCalls) -eq 'DeleteFolder \Claude' -and (Get-StepStatus) -eq 'Done' -and (Get-StepNotes) -eq '') "$(Get-StepStatus): $(Get-StepNotes); $(Get-FilesLeft) file(s) left; $($script:TaskCalls -join ' / ')"
    Assert-Case "each one logged as it went" ((Test-LogHas '^Removed LaunchClaudeAdmin scheduled task$') -and (Test-LogHas '^Removed the empty task folder \\Claude\\$') -and
        @($script:LogLines | Where-Object { $_ -match '^Removed launcher: ' }).Count -eq 6 -and
        @($script:LogLines | Where-Object { $_ -match '^Removed: ' }).Count -eq 3) ($script:LogLines -join ' / ')
    Assert-Case "the RUNASADMIN mark on Claude.exe is left alone" ((Get-FakeRegValue $layers $claudeExe) -eq '~ RUNASADMIN' -and
        @($script:Calls | Where-Object { $_ -like 'Remove-ItemProperty*' }).Count -eq 0) ($script:Calls -join ' / ')

    Reset-State
    Set-Launchers
    $script:Tasks["$TaskPath|Other"] = 'Ready'
    Invoke-UndoStep5
    Assert-Case "a folder that holds another task is kept, and that is not a failure" ($script:TaskFolders -contains '\Claude' -and (Get-FolderCalls) -eq '' -and
        $script:Tasks.ContainsKey("$TaskPath|Other") -and (Get-StepStatus) -eq 'Done' -and
        (Test-LogHas '^The task folder \\Claude\\ still holds other tasks, so it is kept$')) "$(Get-StepStatus): $(Get-StepNotes); $($script:TaskCalls -join ' / ')"

    Reset-State
    Set-Launchers
    [void]$script:TaskFolders.Add('\Claude\Sub')
    Invoke-UndoStep5
    Assert-Case "so is a folder that holds another folder" ($script:TaskFolders -contains '\Claude' -and (Get-FolderCalls) -eq '' -and (Get-StepStatus) -eq 'Done') "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    Set-Launchers
    $script:Tasks["$TaskPath|Hidden"] = 'Ready'
    [void]$script:HiddenTasks.Add("$TaskPath|Hidden")
    Invoke-UndoStep5
    Assert-Case "and a folder that holds only a hidden task" ($script:TaskFolders -contains '\Claude' -and (Get-FolderCalls) -eq '' -and (Get-StepStatus) -eq 'Done' -and
        (Test-LogHas '^The task folder \\Claude\\ still holds other tasks, so it is kept$')) "$(Get-StepStatus): $(Get-StepNotes); $($script:TaskCalls -join ' / ')"

    Reset-State
    Set-Launchers -NoFolder
    Invoke-UndoStep5
    Assert-Case "no task folder: nothing to remove, and not a failure" ((Get-StepStatus) -eq 'Done' -and (Get-StepNotes) -eq '' -and (Get-FolderCalls) -eq '' -and
        -not (Test-LogHas 'task folder')) "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    Set-Launchers
    $script:UnregisterFault = $true
    Invoke-UndoStep5
    Assert-Case "a LaunchClaudeAdmin task that cannot be removed: Failed, and its folder is kept" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq 'The scheduled task LaunchClaudeAdmin: it could not be removed (Access is denied.)' -and
        $script:TaskFolders -contains '\Claude' -and (Get-FolderCalls) -eq '' -and -not (Test-LogHas '^LaunchClaudeAdmin task not found$')) "$(Get-StepStatus): $(Get-StepNotes)"
    Assert-Case "and the files are still removed" ((Get-FilesLeft) -eq 0) "$(Get-FilesLeft) file(s) left"

    Reset-State
    Set-Launchers
    $script:TaskFolderDeleteFault = $true
    Invoke-UndoStep5
    Assert-Case "a task folder the scheduler refuses to delete: Failed" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -match '^The task folder \\Claude\\: it could not be removed \(.*Access is denied' -and $script:TaskFolders -contains '\Claude') "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    Set-Launchers
    $script:TaskFolderDeleteNoop = $true
    Invoke-UndoStep5
    Assert-Case "a task folder still there after it was deleted: Failed" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq 'The task folder \Claude\: it is still there after it was removed' -and -not (Test-LogHas '^Removed the empty task folder')) "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    Set-Launchers
    $script:TaskFolderGoneReadFault = $true
    Invoke-UndoStep5
    Assert-Case "a deleted folder whose absence cannot be read: Failed, not called removed" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -match '^The task folder \\Claude\\: whether it is gone could not be read \(.*The RPC server is unavailable' -and
        -not (Test-LogHas '^Removed the empty task folder')) "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    Set-Launchers
    $script:TaskSvcFault = $true
    Invoke-UndoStep5
    Assert-Case "a Task Scheduler that cannot be reached: Failed" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -match '^The task folder \\Claude\\: the Task Scheduler could not be reached \(Access is denied') "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    Set-Launchers
    $script:TaskFolderReadFault = $true
    Invoke-UndoStep5
    Assert-Case "a task folder that cannot be read: Failed" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -match '^The task folder \\Claude\\: it could not be read \(.*Access is denied' -and (Get-FolderCalls) -eq '') "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    Set-Launchers
    $script:FileRemoveFault = '*Launch-Claude-Admin.vbs'
    Invoke-UndoStep5
    $vbs = $files[3]
    Assert-Case "a launcher that cannot be deleted: Failed, it is kept, and the others are removed" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq "Could not remove ${vbs}: Access to the path is denied." -and (Get-FilesLeft) -eq 1 -and
        (Microsoft.PowerShell.Management\Test-Path -LiteralPath $vbs)) "$(Get-StepStatus): $(Get-StepNotes); $(Get-FilesLeft) file(s) left"

    Reset-State
    Invoke-UndoStep5
    Assert-Case "nothing there at all: Unchanged" ((Get-StepStatus) -eq 'Unchanged' -and (Get-StepNotes) -eq '' -and (Test-LogHas '^LaunchClaudeAdmin task not found$')) "$(Get-StepStatus): $(Get-StepNotes)"
}

# =====================================================================
# Setup step 15 installs the health monitor's logon task. The monitor
# guards itself with the mutex Global\ClaudeHealthMonitor, which versions
# up to 5.0.0 named with their version and which a monitor that could not
# create it runs without, so the task is started now only when every
# monitor already running was found and stopped. Until 6.3.0 the process
# list was read with SilentlyContinue, each stop's error and a failed
# start were dropped, the old task was removed before the new one was
# registered, and a refused registration was not caught: the step said
# "Health monitor installed" with two monitors running, or none.
Invoke-Section 'Setup step 15: the health monitor task, and a monitor already running' {
    $kit = Join-Path $sandbox 'kit'
    $null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $kit -Force
    $watch = Join-Path $kit 'Watch-ClaudeHealth.ps1'
    [System.IO.File]::WriteAllText($watch, '# the monitor')
    $taskKey = "$TaskPath|$TaskName"
    $reg = "Register $TaskPath$TaskName"
    # A start reads the task before, for an old instance that still reads
    # running, and after, to check that it took.
    $get = "Get $TaskPath$TaskName"
    $start = "$get / Start $TaskPath$TaskName / $get"
    $installed = '^Health monitor installed \(starts 120s after logon, polls every 30s\)$'
    $started = '^Started it now: the monitor begins watching in 120 seconds$'
    $tripBefore = $script:Tripped.Count
    function Reset-Install {
        Reset-State
        $script:TaskInstallExpected = $true
        $script:FakeCommandPath = Join-Path $kit 'Prevent-ClaudeIssues.ps1'
    }
    function Set-RunningMonitor {
        $script:PsProcs = @(
            [pscustomobject]@{ ProcessId = 4242; CommandLine = 'powershell.exe -NoProfile -File "C:\ClaudeFix\Watch-ClaudeHealth.ps1" -Quiet' },
            [pscustomobject]@{ ProcessId = 4343; CommandLine = 'powershell.exe -NoProfile -File "C:\Tools\Backup.ps1"' }
        )
        $script:Procs = @((New-FakeProcess 4242 'powershell'), (New-FakeProcess 4343 'powershell'))
    }
    function Get-Calls { return ($script:TaskCalls -join ' / ') }
    function Get-ProcIds { return (@($script:Procs | ForEach-Object { [string]$_.Id }) -join ',') }
    function Get-Seen { return "$(Get-StepStatus): $(Get-StepNotes); tasks: $(Get-Calls); stops: $($script:ProcCalls -join ' / '); $($script:LogLines -join ' / ')" }

    Reset-Install
    Invoke-Step15
    Assert-Case "no monitor running: the task is registered and started now, and the step is Done" ((Get-Calls) -eq "$reg / $start" -and
        $script:ProcCalls.Count -eq 0 -and (Get-StepStatus) -eq 'Done' -and (Get-StepNotes) -eq '' -and $script:Tasks[$taskKey] -eq 'Running' -and
        (Test-LogHas $installed) -and (Test-LogHas $started)) (Get-Seen)
    $arg = ''
    if ($null -ne $script:RegisteredTask) { $arg = [string]$script:RegisteredTask.Action.Arguments }
    Assert-Case "and the task runs the monitor beside Prevent, replacing the one an earlier run made" ($null -ne $script:RegisteredTask -and
        $script:RegisteredTask.Force -and $arg.Contains("& '$watch' -Quiet") -and $script:RegisteredTask.Principal.LogonType -eq 'S4U' -and
        (Test-LogHas ('^Script: ' + [regex]::Escape($watch) + '$'))) $arg

    Reset-Install
    Set-RunningMonitor
    Invoke-Step15
    Assert-Case "a monitor already running: it is stopped, and only it, and the new one is started" (($script:ProcCalls -join ' / ') -eq 'Stop-Process 4242' -and
        (Get-ProcIds) -eq '4343' -and (Get-Calls) -eq "$reg / $start" -and (Get-StepStatus) -eq 'Done' -and (Get-StepNotes) -eq '' -and
        (Test-LogHas '^Stopped the health monitor already running \(PID 4242\)$') -and (Test-LogHas $started)) (Get-Seen)

    Reset-Install
    Set-RunningMonitor
    $script:PsProcsFault = $true
    Invoke-Step15
    $said = 'Could not list the running processes, so a health monitor already running was not looked for or stopped, and the new one was not started now; it starts at the next logon: The RPC server is unavailable.'
    Assert-Case "the running processes cannot be listed: a Warning that says so, nothing stopped, and the task not started now" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq $said -and $script:ProcCalls.Count -eq 0 -and (Get-Calls) -eq $reg -and
        (Test-LogHas ('^' + [regex]::Escape("[!] $said") + '$')) -and -not (Test-LogHas $started)) (Get-Seen)
    Assert-Case "and the task is installed all the same, for the next logon" ($script:Tasks[$taskKey] -eq 'Ready' -and (Test-LogHas $installed)) (Get-Seen)

    Reset-Install
    Set-RunningMonitor
    $script:StopProcFault = $true
    Invoke-Step15
    $said = 'Could not stop the health monitor already running (PID 4242), so it was not replaced, and the new one was not started now; it starts at the next logon: Cannot stop process'
    Assert-Case "a monitor that cannot be stopped: a Warning with the reason, and the new one not started beside it" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes).StartsWith($said) -and (Get-Calls) -eq $reg -and (Get-ProcIds) -eq '4242,4343' -and
        -not (Test-LogHas '^Stopped the health monitor') -and -not (Test-LogHas $started)) (Get-Seen)

    Reset-Install
    Set-RunningMonitor
    $script:StopProcLinger = $true
    Invoke-Step15
    Assert-Case "a monitor still running after it was stopped: a Warning, after waiting 2 seconds, and the new one not started" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq 'Could not stop the health monitor already running (PID 4242), so it was not replaced, and the new one was not started now; it starts at the next logon: it is still running' -and
        $script:FakeClockMs -eq 2000 -and (Get-Calls) -eq $reg -and -not (Test-LogHas $started)) "$(Get-Seen); clock $($script:FakeClockMs)"

    # Listed, but gone by the time it is stopped: the stop's error says so,
    # and the new one is started.
    Reset-Install
    Set-RunningMonitor
    $script:Procs = @(New-FakeProcess 4343 'powershell')
    Invoke-Step15
    Assert-Case "a monitor that ended before it was stopped: said to have ended, and the new one is started" ((Get-StepStatus) -eq 'Done' -and
        (Get-StepNotes) -eq '' -and (Get-Calls) -eq "$reg / $start" -and (Test-LogHas '^The health monitor that was running \(PID 4242\) has ended$') -and
        -not (Test-LogHas '^Stopped the health monitor')) (Get-Seen)

    Reset-Install
    $script:StartTaskFault = $true
    Invoke-Step15
    $said = 'The health monitor task could not be started now, so the monitor starts at the next logon: The operator or administrator has refused the request. (Exception from HRESULT: 0x800710E0)'
    Assert-Case "the task cannot be started now: a Warning that says so and why" ((Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -eq $said -and
        (Get-Calls) -eq "$reg / $get / Start $TaskPath$TaskName" -and (Test-LogHas ('^' + [regex]::Escape("[!] $said") + '$')) -and -not (Test-LogHas $started) -and
        (Test-LogHas $installed)) (Get-Seen)

    Reset-Install
    Set-RunningMonitor
    $script:RegisterFault = $true
    Invoke-Step15
    Assert-Case "the task cannot be registered: Failed with the reason, and the monitor already running is left as it is" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq 'Could not create the health monitor task: Access is denied.' -and (Get-Calls) -eq $reg -and
        $script:ProcCalls.Count -eq 0 -and (Get-ProcIds) -eq '4242,4343' -and -not $script:Tasks.ContainsKey($taskKey) -and
        (Test-LogHas '^\[!\] Could not create health monitor task: Access is denied\.$') -and -not (Test-LogHas $installed)) (Get-Seen)

    Reset-Install
    $script:FakeCommandPath = Join-Path $sandbox 'elsewhere\Prevent-ClaudeIssues.ps1'
    Invoke-Step15
    Assert-Case "the monitor's script is not found: Failed, and no task is touched" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq 'Watch-ClaudeHealth.ps1 was not found beside this script' -and (Get-Calls) -eq '' -and $script:ProcCalls.Count -eq 0) (Get-Seen)

    Assert-Case "and no case here reached a tripwire" ($script:Tripped.Count -eq $tripBefore) (@($script:Tripped) -join ' / ')

    # The gate itself: with no step 15 expected, a task made is a tripwire.
    Reset-State
    $threw = $null
    try { Register-ScheduledTask -TaskName 'Elsewhere' -TaskPath '\' } catch { $threw = $_.Exception.Message }
    $last = ''
    if ($script:Tripped.Count -gt $tripBefore) { $last = [string]$script:Tripped[$script:Tripped.Count - 1] }
    Assert-Case "a task registered outside step 15 trips, and throws" ($threw -eq 'TRIPWIRE: Register-ScheduledTask' -and $last -eq 'Register-ScheduledTask \Elsewhere') "$threw; $last"
    while ($script:Tripped.Count -gt $tripBefore) { $script:Tripped.RemoveAt($script:Tripped.Count - 1) }
}

# =====================================================================
# Setup steps 15 and 17 delete the watchdog script and the .cmd launcher
# older versions left. Each shows that an older version ran here, which
# -Undo step 2 needs to know, so each is recorded before it goes, and a
# file whose record cannot be saved stays. Until 6.3.0 neither was
# recorded, and once setup had run, -Undo could not tell.
Invoke-Section "Setup steps 15 and 17: an older version's watchdog script and .cmd launcher" {
    $oldWd = Join-Path $ClaudeAppData 'cowork-watchdog.ps1'
    $oldCmd = Join-Path $ClaudeAppData 'Launch-Claude-Admin.cmd'

    Reset-State
    [System.IO.File]::WriteAllText($oldWd, '# old')
    Step 15 19 'monitor' -Name 'Health monitor'
    Invoke-Step15Legacy
    $rec = Find-Undo 'LegacyInstall' 'watchdog script'
    Assert-Case "step 15: the old watchdog script is removed" (-not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $oldWd) -and
        (Test-LogHas '^Removed old basic watchdog script$') -and (Get-StepNotes) -eq '') "$(Get-StepStatus): $(Get-StepNotes); $($script:LogLines -join ' / ')"
    $script:UndoRecord = New-Object System.Collections.ArrayList
    Read-PreventUndo
    $saved = Find-Undo 'LegacyInstall' 'watchdog script'
    Assert-Case "after it is recorded, on disk, as what shows an older version ran here" ($null -ne $rec -and $null -ne $saved -and
        [string]$saved.Original -eq $oldWd) (Get-UndoKeys)

    Reset-State
    [System.IO.File]::WriteAllText($oldWd, '# old')
    $null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $UndoFile -Force
    Step 15 19 'monitor' -Name 'Health monitor'
    Invoke-Step15Legacy
    Assert-Case "a record that cannot be saved: the script stays, with a Warning that says why" ((Microsoft.PowerShell.Management\Test-Path -LiteralPath $oldWd) -and
        (Get-StepStatus) -eq 'Warning' -and (Get-StepNotes).StartsWith("Could not remove the old watchdog script ${oldWd}: Could not save the undo record (") -and
        -not (Test-LogHas '^Removed old basic watchdog script$')) "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    [System.IO.File]::WriteAllText($oldWd, '# old')
    $script:FileRemoveFault = '*cowork-watchdog.ps1'
    Step 15 19 'monitor' -Name 'Health monitor'
    Invoke-Step15Legacy
    Assert-Case "a script that cannot be deleted: kept, with a Warning that says why" ((Microsoft.PowerShell.Management\Test-Path -LiteralPath $oldWd) -and
        (Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -eq "Could not remove the old watchdog script ${oldWd}: Access to the path is denied." -and
        -not (Test-LogHas '^Removed old basic watchdog script$')) "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    Step 15 19 'monitor' -Name 'Health monitor'
    Invoke-Step15Legacy
    Assert-Case "no old script: nothing recorded and nothing said" ($script:UndoRecord.Count -eq 0 -and (Get-StepStatus) -eq 'None' -and
        @($script:LogLines).Count -eq 1) "$(Get-StepStatus): $($script:LogLines -join ' / '); $(Get-UndoKeys)"

    Reset-State
    [System.IO.File]::WriteAllText($oldCmd, '@echo off')
    Step 17 19 'launchers' -Name 'Claude (Admin) launcher'
    Invoke-Step17Legacy
    Assert-Case "step 17: the old .cmd launcher is removed" (-not (Microsoft.PowerShell.Management\Test-Path -LiteralPath $oldCmd) -and
        (Test-LogHas ('^Removed superseded \.cmd launcher: ' + [regex]::Escape($oldCmd) + '$')) -and (Get-StepNotes) -eq '') "$(Get-StepStatus): $(Get-StepNotes); $($script:LogLines -join ' / ')"
    $script:UndoRecord = New-Object System.Collections.ArrayList
    Read-PreventUndo
    $saved = Find-Undo 'LegacyInstall' 'admin launcher'
    Assert-Case "after it is recorded, on disk, as what shows an older version ran here" ($null -ne $saved -and [string]$saved.Original -eq $oldCmd) (Get-UndoKeys)

    Reset-State
    [System.IO.File]::WriteAllText($oldCmd, '@echo off')
    $null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $UndoFile -Force
    Step 17 19 'launchers' -Name 'Claude (Admin) launcher'
    Invoke-Step17Legacy
    Assert-Case "a record that cannot be saved: the launcher stays, with a Warning that says why" ((Microsoft.PowerShell.Management\Test-Path -LiteralPath $oldCmd) -and
        (Get-StepStatus) -eq 'Warning' -and (Get-StepNotes).StartsWith("Could not remove the superseded .cmd launcher ${oldCmd}: Could not save the undo record (")) "$(Get-StepStatus): $(Get-StepNotes)"

    Reset-State
    [System.IO.File]::WriteAllText($oldCmd, '@echo off')
    $script:FileRemoveFault = '*Launch-Claude-Admin.cmd'
    Step 17 19 'launchers' -Name 'Claude (Admin) launcher'
    Invoke-Step17Legacy
    Assert-Case "a launcher that cannot be deleted: kept, with a Warning that says why" ((Microsoft.PowerShell.Management\Test-Path -LiteralPath $oldCmd) -and
        (Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -eq "Could not remove the superseded .cmd launcher ${oldCmd}: Access to the path is denied.") "$(Get-StepStatus): $(Get-StepNotes)"
}

# =====================================================================
# The undo record names the Windows installation that wrote it by its
# MachineGuid, beside the computer's name. The MachineGuid decides where
# the record and this computer both have one: renaming a computer does not
# change it, and another computer given the same name has its own. The
# name decides otherwise, as it did for a record from before 6.3.0. Until
# 6.3.0 only the name was kept, so a renamed computer could not use its
# own record, and a profile copied to a computer of the same name could
# use another's.
Invoke-Section 'The undo record: which computer wrote it, by its MachineGuid' {
    $crypto = 'HKLM:\SOFTWARE\Microsoft\Cryptography'
    $mine = 'aaaaaaaa-1111-2222-3333-444444444444'
    $other = 'bbbbbbbb-5555-6666-7777-888888888888'
    $me = [Environment]::MachineName
    function Set-Record {
        param([string]$Name, [string]$Guid)
        $head = '{ "Version": 1, "ComputerName": "' + $Name + '", '
        if ($Guid) { $head += '"MachineGuid": "' + $Guid + '", ' }
        [System.IO.File]::WriteAllText($UndoFile, $head + '"Items": [ { "Kind": "Registry", "Key": "HKLM:\\T|X", "Original": 5 } ] }')
    }
    function Get-Loaded { return "$($script:UndoRecord.Count) item(s), from '$($script:UndoForeign)', by GUID $($script:UndoForeignByGuid)" }

    Reset-State
    Set-FakeRegValue $crypto 'MachineGuid' $mine
    Add-PreventUndoItem -Kind Registry -Key 'HKLM:\T|X' -Original 1
    $doc = [System.IO.File]::ReadAllText($UndoFile) | ConvertFrom-Json
    $got = '<none>'
    if ($doc.PSObject.Properties['MachineGuid']) { $got = [string]$doc.MachineGuid }
    Assert-Case "a save names this installation's MachineGuid beside the computer's name" ($got -eq $mine -and [string]$doc.ComputerName -eq $me) "got $got, $($doc.ComputerName)"

    Reset-State
    Set-FakeRegValue $crypto 'MachineGuid' $mine
    Set-Record 'OLD-NAME' $mine.ToUpperInvariant()
    Read-PreventUndo
    Assert-Case "a record under another name with this MachineGuid, in any case, is this computer's: it loads" ($script:UndoRecord.Count -eq 1 -and
        $null -eq $script:UndoForeign -and -not $script:UndoForeignByGuid) (Get-Loaded)
    Reset-State
    Set-FakeRegValue $crypto 'MachineGuid' $mine
    Set-Record 'OLD-NAME' "  $mine  "
    Read-PreventUndo
    Assert-Case "and so is one whose MachineGuid has spaces around it" ($script:UndoRecord.Count -eq 1 -and $null -eq $script:UndoForeign) (Get-Loaded)

    Reset-State
    Set-FakeRegValue $crypto 'MachineGuid' $mine
    Set-Record $me $other
    Read-PreventUndo
    Assert-Case "a record under this name with another MachineGuid is another computer's: it is not loaded" ($script:UndoRecord.Count -eq 0 -and
        $script:UndoForeign -eq $me -and $script:UndoForeignByGuid) (Get-Loaded)
    Set-PowerFixture
    $script:UndoMode = $true
    Invoke-UndoStep1
    Assert-Case "-Undo step 1 fails, and says the MachineGuid told them apart, not that a rename would help" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq "The undo record was written on another computer ($me): nothing was put back from it" -and
        (Test-LogHas "^It holds another Windows installation's MachineGuid, which renaming this computer does not change$") -and
        -not (Test-LogHas 'set ComputerName')) "$(Get-StepStatus): $(Get-StepNotes); $($script:LogLines -join ' / ')"

    Reset-State
    Set-Record $me $other
    Read-PreventUndo
    Assert-Case "a record with a MachineGuid where this computer has none: the name decides, and it loads" ($script:UndoRecord.Count -eq 1 -and
        $null -eq $script:UndoForeign) (Get-Loaded)
    Reset-State
    Set-Record 'OTHER-PC' $mine
    Read-PreventUndo
    Assert-Case "and under another name it is not loaded" ($script:UndoRecord.Count -eq 0 -and $script:UndoForeign -eq 'OTHER-PC' -and
        -not $script:UndoForeignByGuid) (Get-Loaded)

    Reset-State
    Set-FakeRegValue $crypto 'MachineGuid' $mine
    Set-Record 'OTHER-PC' ''
    Read-PreventUndo
    Assert-Case "a record from before 6.3.0, under another name: the name decides, and it is not loaded" ($script:UndoRecord.Count -eq 0 -and
        $script:UndoForeign -eq 'OTHER-PC' -and -not $script:UndoForeignByGuid) (Get-Loaded)
    Set-PowerFixture
    $script:UndoMode = $true
    Invoke-UndoStep1
    Assert-Case "-Undo step 1 fails, and says how to claim it if this is that computer renamed" ((Get-StepStatus) -eq 'Failed' -and
        (Test-LogHas ('^If this is that computer under a new name, set ComputerName in the file to ' + [regex]::Escape($me) + ' and run -Undo again$')) -and
        -not (Test-LogHas 'MachineGuid')) "$(Get-StepStatus): $($script:LogLines -join ' / ')"
    Reset-State
    Set-FakeRegValue $crypto 'MachineGuid' $mine
    Set-Record $me ''
    Read-PreventUndo
    Assert-Case "and one under this name loads" ($script:UndoRecord.Count -eq 1 -and $null -eq $script:UndoForeign) (Get-Loaded)

    Reset-State
    Set-FakeRegValue $crypto 'MachineGuid' "  $mine  "
    $padded = Get-PreventMachineGuid
    Reset-State
    $none = Get-PreventMachineGuid
    Assert-Case "Get-PreventMachineGuid: the value trimmed, and '' when there is none" ($padded -eq $mine -and $none -is [string] -and $none -eq '') "'$padded' / '$none'"
}

# =====================================================================
# -Undo sets a recorded power setting back and reads it again. A driver or
# an update can take a setting away, and powercfg then can neither set nor
# read it, in the machine's language. Where Windows no longer defines the
# setting (its key under Control\Power\PowerSettings is gone while that
# folder is there) the record is dropped, with a Warning. Until 6.3.0 it
# stayed, and every -Undo after that failed on it.
Invoke-Section '-Undo: a power setting that is no longer on this machine' {
    $root = 'HKLM:\SYSTEM\CurrentControlSet\Control\Power\PowerSettings'
    $key = "$plBal|$subSleep|$setSleep"
    $dropped = 'Sleep after: the setting is no longer on this machine, so its record is dropped'
    function Set-Settings {
        param([switch]$Defined, [switch]$NoRoot)
        Reset-State
        Set-PowerFixture
        if (-not $NoRoot) {
            $null = Add-FakeRegKey $root
            $null = Add-FakeRegKey "$root\$subSleep"
        }
        if ($Defined) { $null = Add-FakeRegKey "$root\$subSleep\$setSleep" }
    }
    function Restore-Sleep {
        Step 1 5 'undo' -Name 'Recorded settings'
        return (Restore-PreventUndoItem -Item ([pscustomobject]@{ Kind = 'PowerValue'; Key = $key; Original = 1800 }) -CreatedPlan $null)
    }

    Set-Settings
    $ok = Restore-Sleep
    Assert-Case "a setting Windows no longer defines: its record is dropped, with a Warning" ($ok -eq $true -and (Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq $dropped -and
        (Test-LogHas '^\[!\] Sleep after: this setting is no longer on this machine, so its record is dropped$')) "$ok; $(Get-StepStatus): $(Get-StepNotes)"

    Set-Settings -Defined
    $ok = Restore-Sleep
    Assert-Case "one still defined that cannot be read back stays recorded, and fails the step" ($ok -eq $false -and (Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq 'Sleep after could not be read back after putting it back') "$ok; $(Get-StepStatus): $(Get-StepNotes)"

    Set-Settings -NoRoot
    $ok = Restore-Sleep
    Assert-Case "where the PowerSettings folder cannot be seen, nothing is known: it stays recorded" ($ok -eq $false -and (Get-StepStatus) -eq 'Failed') "$ok; $(Get-StepStatus): $(Get-StepNotes)"

    Set-Settings
    Add-PreventUndoItem -Kind PowerValue -Key $key -Original 1800
    $script:StepResults = New-Object System.Collections.ArrayList
    $script:UndoMode = $true
    Invoke-UndoStep1
    $script:UndoRecord = New-Object System.Collections.ArrayList
    Read-PreventUndo
    Assert-Case "-Undo step 1 drops it from the record on disk, and the step is a Warning" ($null -eq (Find-Undo 'PowerValue' $key) -and
        (Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -eq $dropped) "$(Get-StepStatus): $(Get-StepNotes); $(Get-UndoKeys)"

    # The test itself, last: a version without it stops here.
    Set-Settings
    $gone = Test-PowerSettingGone -Sub $subSleep -Setting $setSleep
    Set-Settings -Defined
    $there = Test-PowerSettingGone -Sub $subSleep -Setting $setSleep
    Set-Settings -NoRoot
    $unseen = Test-PowerSettingGone -Sub $subSleep -Setting $setSleep
    Assert-Case "Test-PowerSettingGone: gone only when the folder is there and the setting's key is not" ($gone -eq $true -and $there -eq $false -and
        $unseen -eq $false) "gone $gone, defined $there, no folder $unseen"
}

# =====================================================================
# Setup writes paths and names into PowerShell commands it generates: the
# monitor's logon task runs the monitor script by its path, and the NoGPU
# launcher starts Claude by its package or by its exe. Each goes in single
# quotes, where a ' is written twice. Until 6.3.0 it was not, so a ' in a
# user or folder name (C:\Users\O'Brien) made a command PowerShell could
# not parse, while setup reported the step as done. Each command is built
# here from Prevent's own text, with names that hold a ', and parsed.
Invoke-Section 'Generated commands: a path or name that holds an apostrophe' {
    $watchScript = "C:\Users\O'Brien\ClaudeFix\Watch-ClaudeHealth.ps1"
    $pfn = 'Claude_pzs8sxrjxfjjc'
    $appId = "App's"
    $exePath = "C:\Users\O'Brien\AppData\Local\AnthropicClaude\claude.exe"
    $null = $watchScript, $pfn, $appId, $exePath
    function Test-Parses {
        param([string]$Text)
        $tokens = $null
        $errs = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($Text, [ref]$tokens, [ref]$errs)
        $vals = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] }, $true) | ForEach-Object { [string]$_.Value })
        return [pscustomobject]@{ Errors = @($errs).Count; Values = $vals }
    }

    $bad = Test-Parses "Start-Sleep -Seconds 120; & 'C:\Users\O'Brien\ClaudeFix\Watch-ClaudeHealth.ps1' -Quiet"
    Assert-Case "a ' written once does not parse, so the check below can fail" ($bad.Errors -gt 0) "$($bad.Errors) error(s)"

    $assigns = @($pAst.FindAll({
        param($n)
        $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
        $n.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
        @('delayedWatchCmd', 'launchBody') -contains $n.Left.VariablePath.UserPath
    }, $true))
    Assert-Case "the three are found: the monitor's task and the launcher's two shapes" ($assigns.Count -eq 3) "$($assigns.Count) found"
    foreach ($as in $assigns) {
        $name = $as.Left.VariablePath.UserPath
        $line = $as.Extent.StartLineNumber
        $want = @($exePath)
        $what = "the NoGPU launcher for an exe"
        if ($name -eq 'delayedWatchCmd') {
            $want = @($watchScript)
            $what = "the monitor's logon task"
        } elseif ($as.Right.Extent.Text -match 'PackageFamilyName') {
            $want = @($pfn, $appId)
            $what = "the NoGPU launcher for the MSIX package"
        }
        # Only a string is built here: anything else in Prevent's text is
        # not run.
        $expr = $null
        if ($as.Right -is [System.Management.Automation.Language.CommandExpressionAst]) { $expr = $as.Right.Expression }
        if ($expr -isnot [System.Management.Automation.Language.ExpandableStringExpressionAst] -and
            $expr -isnot [System.Management.Automation.Language.StringConstantExpressionAst]) {
            Assert-Case "${what}: built from a string" $false "Prevent:${line}: $($as.Right.Extent.Text)"
            continue
        }
        $text = [string](& ([scriptblock]::Create($expr.Extent.Text)))
        $r = Test-Parses $text
        $missing = @($want | Where-Object { $r.Values -notcontains $_ })
        Assert-Case "${what}: it parses, and holds each name whole" ($r.Errors -eq 0 -and $missing.Count -eq 0) "Prevent:${line}: $($r.Errors) error(s), missing $($missing -join ', '): $text"
    }
}

# =====================================================================
# Setup step 15 starts the monitor's task once the old monitor is
# stopped. The task runs one instance at a time (MultipleInstances
# IgnoreNew), and the scheduler ignores a start, with no error, while it
# reads an instance as running, which it can for a moment after the old
# monitor is stopped. So the start waits for that to end, up to 5
# seconds, and is checked afterwards. Until 6.3.1 neither was done, and
# "Started it now" was said of a start the scheduler ignored.
Invoke-Section 'Setup step 15: the start waits for the old instance, and is checked' {
    $kit = Join-Path $sandbox 'kit'
    $null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $kit -Force
    [System.IO.File]::WriteAllText((Join-Path $kit 'Watch-ClaudeHealth.ps1'), '# the monitor')
    $taskKey = "$TaskPath|$TaskName"
    $reg = "Register $TaskPath$TaskName"
    $get = "Get $TaskPath$TaskName"
    $start = "Start $TaskPath$TaskName"
    $started = '^Started it now: the monitor begins watching in 120 seconds$'
    $rpc = 'The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)'
    $tripBefore = $script:Tripped.Count
    function Reset-Install {
        Reset-State
        $script:TaskInstallExpected = $true
        $script:FakeCommandPath = Join-Path $kit 'Prevent-ClaudeIssues.ps1'
    }
    function Get-Calls { return ($script:TaskCalls -join ' / ') }
    function Get-Count { param([string]$Call) return @($script:TaskCalls | Where-Object { $_ -eq $Call }).Count }
    function Get-Seen { return "$(Get-StepStatus): $(Get-StepNotes); tasks: $(Get-Calls); clock $($script:FakeClockMs) ms" }

    Reset-Install
    $script:Tasks[$taskKey] = 'Running'
    $script:TaskStopReads[$taskKey] = 0
    Invoke-Step15
    Assert-Case "an instance that has ended by the first read: started at once, and Done" ((Get-Calls) -eq "$reg / $get / $start / $get" -and
        $script:FakeClockMs -eq 0 -and (Get-StepStatus) -eq 'Done' -and (Get-StepNotes) -eq '' -and (Test-LogHas $started)) (Get-Seen)

    Reset-Install
    $script:Tasks[$taskKey] = 'Running'
    $script:TaskStopReads[$taskKey] = 3
    Invoke-Step15
    Assert-Case "an instance that ends after 3 more reads: started after 750 ms, and Done" ((Get-Calls) -eq "$reg / $get / $get / $get / $get / $start / $get" -and
        $script:FakeClockMs -eq 750 -and (Get-StepStatus) -eq 'Done' -and (Get-StepNotes) -eq '' -and (Test-LogHas $started)) (Get-Seen)

    Reset-Install
    $script:Tasks[$taskKey] = 'Running'
    $script:TaskStopReads[$taskKey] = -1
    Invoke-Step15
    $said = 'The health monitor task was not started now, because it still reads as running 5 seconds after the old monitor was stopped, and the scheduler ignores a start while it does; the monitor starts at the next logon'
    Assert-Case "an instance that never ends: not started, after 5 seconds, and a Warning that says so" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq $said -and (Get-Count $start) -eq 0 -and (Get-Count $get) -eq 21 -and $script:FakeClockMs -eq 5000 -and
        (Test-LogHas ('^' + [regex]::Escape("[!] $said") + '$')) -and -not (Test-LogHas $started)) (Get-Seen)

    Reset-Install
    $script:StartTaskNoop = $true
    Invoke-Step15
    $said = 'The scheduler accepted the start of the health monitor task, but the task did not start: it reads Ready 3 seconds later. The monitor starts at the next logon'
    Assert-Case "a start the scheduler takes and does nothing with: a Warning after 3 seconds, not 'Started it now'" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq $said -and (Get-Count $start) -eq 1 -and (Get-Count $get) -eq 14 -and $script:FakeClockMs -eq 3000 -and
        -not (Test-LogHas $started)) (Get-Seen)

    Reset-Install
    $script:TaskGoneAfterStart = $true
    Invoke-Step15
    $said = 'The scheduler accepted the start of the health monitor task, but the task is no longer there 3 seconds later'
    Assert-Case "a task that is gone after the start: a Warning that says so, not 'Started it now'" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq $said -and (Get-Count $start) -eq 1 -and (Get-Count $get) -eq 14 -and $script:FakeClockMs -eq 3000 -and
        -not (Test-LogHas $started)) (Get-Seen)

    Reset-Install
    $script:TaskQueryFault = $true
    Invoke-Step15
    $said = "The health monitor task was not started now, because its state could not be read ($rpc); the monitor starts at the next logon"
    Assert-Case "a task whose state cannot be read before the start: not started, and a Warning that says why" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq $said -and (Get-Calls) -eq "$reg / $get" -and -not (Test-LogHas $started)) (Get-Seen)

    Reset-Install
    $script:TaskQueryFaultAfterStart = $true
    Invoke-Step15
    $said = "The health monitor task was started, but its state could not be read to check that it runs: $rpc"
    Assert-Case "a task whose state cannot be read after the start: a Warning, not 'Started it now'" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq $said -and (Get-Calls) -eq "$reg / $get / $start / $get" -and -not (Test-LogHas $started)) (Get-Seen)

    # The old monitor was stopped, and whether it ended cannot be read.
    Reset-Install
    $script:PsProcs = @([pscustomobject]@{ ProcessId = 4242; CommandLine = 'powershell.exe -NoProfile -File "C:\ClaudeFix\Watch-ClaudeHealth.ps1" -Quiet' })
    $script:Procs = @(New-FakeProcess 4242 'powershell')
    $script:ProcReadFault = @('4242')
    Invoke-Step15
    $said = 'Could not read whether the health monitor already running (PID 4242) stopped, so the new one was not started now; it starts at the next logon: Access is denied'
    Assert-Case "an old monitor that cannot be read after it was stopped: the new one not started, and a Warning that says so" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq $said -and (Get-Calls) -eq $reg -and -not (Test-LogHas '^Stopped the health monitor') -and -not (Test-LogHas $started)) (Get-Seen)

    Assert-Case "and no case here reached a tripwire" ($script:Tripped.Count -eq $tripBefore) (@($script:Tripped) -join ' / ')
}

# =====================================================================
# Setup step 17 makes the Claude (Admin) task and the shortcut that starts
# it. Until 6.3.1 the old task was removed with its errors dropped, and a
# registration the scheduler refused went on to say "Scheduled task
# created" and to make a shortcut that started nothing.
Invoke-Section 'Setup step 17: the Claude (Admin) task' {
    $adminKey = "$TaskPath|LaunchClaudeAdmin"
    $get = "Get ${TaskPath}LaunchClaudeAdmin"
    $unreg = "Unregister ${TaskPath}LaunchClaudeAdmin"
    $reg = "Register ${TaskPath}LaunchClaudeAdmin"
    $lnk = Join-Path $fakeDesktop 'Claude (Admin).lnk'
    $elevVbs = Join-Path $ClaudeAppData 'Launch-Claude-Admin.vbs'
    $trigVbs = Join-Path $ClaudeAppData 'Launch-Claude-Admin-Trigger.vbs'
    $created = '^Scheduled task created: \\Claude\\LaunchClaudeAdmin \(Highest \+ Interactive, no console\)$'
    $rpc = 'The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)'
    $tripBefore = $script:Tripped.Count
    function Reset-Admin {
        Reset-State
        $script:TaskInstallExpected = $true
    }
    function Get-Calls { return ($script:TaskCalls -join ' / ') }
    function Get-Saved { return (@($script:SavedShortcuts | ForEach-Object { [string]$_.Path }) -join ' / ') }
    function Get-Seen { return "$(Get-StepStatus): $(Get-StepNotes); tasks: $(Get-Calls); shortcuts: $(Get-Saved)" }
    function Test-File { param([string]$Path) return (Microsoft.PowerShell.Management\Test-Path -LiteralPath $Path) }

    Reset-Admin
    Invoke-Step17
    Assert-Case "no task yet: it is registered, both launchers written and the shortcut saved: Done" ((Get-Calls) -eq "$get / $reg" -and
        $script:Tasks[$adminKey] -eq 'Ready' -and (Get-Saved) -eq $lnk -and (Get-StepStatus) -eq 'Done' -and (Get-StepNotes) -eq '' -and
        (Test-LogHas $created) -and (Test-File $elevVbs) -and (Test-File $trigVbs)) (Get-Seen)
    $p = $script:RegisteredTask
    Assert-Case "and the task runs as this user, elevated and on the desktop, replacing any other" ($null -ne $p -and $p.Force -and
        $p.Principal.UserId -eq 'TESTPC\tester' -and $p.Principal.RunLevel -eq 'Highest' -and $p.Principal.LogonType -eq 'Interactive') (Get-Seen)

    Reset-Admin
    $script:Tasks[$adminKey] = 'Ready'
    Invoke-Step17
    Assert-Case "an old task: it is removed, then registered again: Done" ((Get-Calls) -eq "$get / $unreg / $reg" -and
        $script:Tasks[$adminKey] -eq 'Ready' -and (Get-Saved) -eq $lnk -and (Get-StepStatus) -eq 'Done' -and (Get-StepNotes) -eq '') (Get-Seen)

    Reset-Admin
    $script:RegisterFault = $true
    Invoke-Step17
    $said = 'Could not configure the elevated launcher: the scheduler refused the task \Claude\LaunchClaudeAdmin, so no Claude (Admin) shortcut was made (Access is denied.)'
    Assert-Case "a registration the scheduler refuses: Failed with the reason, no shortcut made, and no task said to be created" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq $said -and (Get-Saved) -eq '' -and -not (Test-LogHas $created) -and (Get-Calls) -eq "$get / $reg / $get" -and
        -not (Test-File $trigVbs)) (Get-Seen)

    Reset-Admin
    $script:RegisterFault = $true
    [System.IO.File]::WriteAllText($lnk, 'an earlier shortcut')
    Invoke-Step17
    Assert-Case "and a shortcut an earlier run made, whose task is not there: removed, and said" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq "$said. The shortcut $lnk an earlier run made was removed: the task it starts is not there" -and -not (Test-File $lnk)) (Get-Seen)

    Reset-Admin
    $script:Tasks[$adminKey] = 'Ready'
    $script:UnregisterFault = $true
    $script:RegisterFault = $true
    [System.IO.File]::WriteAllText($lnk, 'an earlier shortcut')
    Invoke-Step17
    Assert-Case "a shortcut an earlier run made, whose task is still there: kept" ((Get-StepStatus) -eq 'Failed' -and (Test-File $lnk) -and
        (Get-StepNotes).EndsWith($said) -and -not ((Get-StepNotes) -match 'The shortcut ')) (Get-Seen)

    Reset-Admin
    $script:RegisterFault = $true
    [System.IO.File]::WriteAllText($lnk, 'an earlier shortcut')
    $script:ItemReadFault = $lnk
    Invoke-Step17
    Assert-Case "a shortcut an earlier run made that cannot be read: kept, and said" ((Get-StepStatus) -eq 'Failed' -and (Test-File $lnk) -and
        (Get-StepNotes) -eq "$said. Whether the shortcut $lnk an earlier run made still has its task could not be read: Access is denied") (Get-Seen)

    Reset-Admin
    $script:Tasks[$adminKey] = 'Ready'
    $script:UnregisterFault = $true
    Invoke-Step17
    $said = 'Could not remove the old scheduled task \Claude\LaunchClaudeAdmin; registering it again replaces it: Access is denied.'
    Assert-Case "an old task that cannot be removed: a Warning that says so, and it is registered over all the same" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq $said -and (Get-Calls) -eq "$get / $unreg / $reg" -and (Get-Saved) -eq $lnk -and (Test-LogHas $created)) (Get-Seen)

    Reset-Admin
    $script:Tasks[$adminKey] = 'Ready'
    $script:TaskQueryFault = $true
    Invoke-Step17
    $said = "Could not read whether the scheduled task \Claude\LaunchClaudeAdmin is there, so it was not removed first; registering it replaces it: $rpc"
    Assert-Case "an old task that cannot be read: a Warning, nothing removed, and it is registered over" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq $said -and (Get-Calls) -eq "$get / $reg" -and (Get-Saved) -eq $lnk) (Get-Seen)

    Reset-Admin
    $script:SetContentFault = '*Trigger.vbs'
    Invoke-Step17
    Assert-Case "a launcher that cannot be written: Failed with the reason, and no shortcut" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq "Could not configure the elevated launcher: Access to the path '$trigVbs' is denied." -and (Get-Saved) -eq '' -and
        -not (Test-File $trigVbs)) (Get-Seen)

    Reset-Admin
    $script:FakeIsAdmin = $false
    Invoke-Step17
    Assert-Case "not run as an administrator: Skipped, and nothing is made" ((Get-StepStatus) -eq 'Skipped' -and (Get-Calls) -eq '' -and
        (Get-Saved) -eq '' -and -not (Test-File $elevVbs)) (Get-Seen)

    Assert-Case "and no case here reached a tripwire" ($script:Tripped.Count -eq $tripBefore) (@($script:Tripped) -join ' / ')
}

# =====================================================================
# Whether an older version ran here decides what -Undo removes, and what
# setup step 19 says. A file or a task that cannot be read is not one
# that is not there: until 6.3.1 it counted as not there, and so as no
# older version, and setup step 19 left the logon task unchecked.
Invoke-Section "-Undo step 2 and setup step 19: an older version's traces that cannot be read" {
    $policyPath  = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
    $controlPath = 'HKLM:\SYSTEM\CurrentControlSet\Control'
    $flagPath    = Join-Path $ClaudeAppData 'disable-dynamic-memory.flag'
    $rpc = 'The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)'
    function Set-LegacyValues {
        Set-FakeRegValue $controlPath 'ServicesPipeTimeout' ([int]120000)
        Set-FakeRegValue $policyPath 'LocalAccountTokenFilterPolicy' ([int]1)
        Set-FakeRegValue $policyPath 'FilterAdministratorToken' ([int]0)
    }
    function Get-LegacyValuesLeft {
        $left = @()
        foreach ($x in @(@($controlPath, 'ServicesPipeTimeout'), @($policyPath, 'LocalAccountTokenFilterPolicy'), @($policyPath, 'FilterAdministratorToken'))) {
            if (Test-FakeRegValue $x[0] $x[1]) { $left += $x[1] }
        }
        return ($left -join ',')
    }
    function Get-Seen { return "$(Get-StepStatus): $(Get-StepNotes); unread $($script:UndoUnread); left $(Get-LegacyValuesLeft)" }
    $all = 'ServicesPipeTimeout,LocalAccountTokenFilterPolicy,FilterAdministratorToken'
    $taskWhy = "the scheduled task $TaskPath$BootTaskName could not be read ($rpc)"
    $told = ' matches what older versions set, but whether one ran here could not be told, so it is left as it is'

    Reset-State
    Set-LegacyValues
    $script:TaskQueryFault = $true
    Invoke-UndoStep2
    $said = "Whether a version before 6.3.0 ran here could not be told, because $taskWhy, so what -Undo removes only where one ran is left as it is"
    Assert-Case "-Undo, the logon task cannot be read: a Warning that says so, and the three values left, each said" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq "$said | ServicesPipeTimeout = 120000$told | LocalAccountTokenFilterPolicy = 1$told | FilterAdministratorToken = 0$told" -and
        (Get-LegacyValuesLeft) -eq $all -and @($script:Calls | Where-Object { $_.StartsWith('Remove-ItemProperty ') }).Count -eq 0) (Get-Seen)
    $end = Get-UndoEnd
    Assert-Case "and -Undo ends UNDO INCOMPLETE, saying to run it again, with exit code 1" ($end.Box -eq 'UNDO INCOMPLETE' -and $end.Unread -and $end.Exit -eq 1) $end.Text

    Reset-State
    Set-LegacyValues
    $script:ItemReadFault = $flagPath
    Invoke-UndoStep2
    Assert-Case "-Undo, a file that shows an older version cannot be read: the same, naming the file" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes).StartsWith("Whether a version before 6.3.0 ran here could not be told, because $flagPath could not be read (Access is denied), so") -and
        (Get-LegacyValuesLeft) -eq $all -and $script:UndoUnread) (Get-Seen)

    Reset-State
    Set-LegacyValues
    $script:TaskQueryFault = $true
    Invoke-Step19
    $s1 = "Whether a version before 6.3.0 ran here could not be told, because $taskWhy"
    $s2 = "Could not read the scheduled task $TaskPath$BootTaskName, so the logon task older versions created was not checked or removed: $rpc"
    $g = ' matches what older versions set, but whether one ran here could not be told, so it is not known whether -Undo removes it'
    Assert-Case "setup step 19, the logon task cannot be read: a Warning for each, and the values said to be undecided" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq "$s1 | $s2 | LocalAccountTokenFilterPolicy = 1$g | FilterAdministratorToken = 0$g | ServicesPipeTimeout = 120000$g" -and
        @($script:TaskCalls | Where-Object { $_ -like 'Unregister *' }).Count -eq 0) (Get-Seen)

    Reset-State
    $script:Services[$ServiceName] = 'Running'
    $script:ServiceReadFault = @($ServiceName)
    Invoke-Step19
    Assert-Case "setup step 19, the Cowork VM service cannot be read: a Warning that says so" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq "Could not read the service $ServiceName, so the restart recovery older versions set on it was not checked: Access is denied") (Get-Seen)
}

# =====================================================================
# A service that cannot be read is not one that is not installed. Until
# 6.3.1 each of these read it with SilentlyContinue: -Undo dropped the
# record of a recovery it had not put back, setup step 9 said vmcompute
# was not there, and setup step 12 said the Windows Time service was not
# installed.
Invoke-Section 'Services that cannot be read' {
    function New-UndoItem { param([string]$Kind, [string]$Key, $Original) return [pscustomobject]@{ Kind = $Kind; Key = $Key; Original = $Original } }
    $svcKey = 'HKLM:\SYSTEM\CurrentControlSet\Services\vmcompute'
    $denied = 'Access is denied'
    function Set-Recovery {
        Reset-State
        $script:Services['vmcompute'] = 'Running'
        Set-FakeRegValue $svcKey 'FailureActions' (New-FailureBlob -Reset 300 -Pairs @(1, 30000, 1, 60000, 1, 120000))
    }
    function Test-RecoveryOn { return [bool](Test-RestartRecovery (ConvertFrom-FailureActionsBlob -Bytes (Get-ServiceFailureBytes -Service 'vmcompute'))) }
    function Get-Seen { return "$(Get-StepStatus): $(Get-StepNotes); sc: $($script:ScCalls -join ' / ')" }

    Set-Recovery
    $script:ServiceReadFault = @('vmcompute')
    Step 1 5 'undo' -Name 'Recorded settings'
    $threw = ''
    try { $null = Restore-PreventUndoItem -Item (New-UndoItem 'ServiceFailureActions' 'vmcompute' $null) -CreatedPlan $null } catch { $threw = $_.Exception.Message }
    Assert-Case "-Undo step 1, a service that cannot be read: it throws, so its recovery stays recorded, and sc.exe is not run" (
        $threw -eq "the service could not be read, so its recovery settings stay recorded ($denied)" -and $script:ScCalls.Count -eq 0 -and (Test-RecoveryOn)) "threw: $threw; $(Get-Seen)"

    Set-Recovery
    $script:ServiceReadFault = @('vmcompute')
    Invoke-UndoStep2
    Assert-Case "-Undo step 2, vmcompute cannot be read: a Warning, its recovery left, and -Undo incomplete" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq "Could not read the service vmcompute, so the restart recovery older versions set on it was not checked: $denied" -and
        $script:ScCalls.Count -eq 0 -and (Test-RecoveryOn) -and $script:UndoUnread) (Get-Seen)

    Reset-State
    $script:Services['vmcompute'] = 'Running'
    $null = Add-FakeRegKey $svcKey
    $script:ServiceReadFault = @('vmcompute')
    Invoke-Step9
    Assert-Case "setup step 9, vmcompute cannot be read: Failed, saying so, and nothing set or recorded" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq "Could not read the service vmcompute, so its restart recovery was not checked or set: $denied" -and $script:ScCalls.Count -eq 0 -and
        $null -eq (Find-Undo 'ServiceFailureActions' 'vmcompute')) (Get-Seen)

    Reset-State
    $script:Services['W32Time'] = 'Running'
    $script:ServiceReadFault = @('W32Time')
    Invoke-Step12
    Assert-Case "setup step 12, the Windows Time service cannot be read: a Warning that says so, nothing started, and w32tm not asked" (
        (Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -eq "Could not read the Windows Time service, so it was not started and the clock was not checked: $denied" -and
        $script:SvcCalls.Count -eq 0 -and $script:W32Calls.Count -eq 0) "$(Get-Seen); $($script:SvcCalls -join ' / ')"
}

# =====================================================================
# A process that cannot be read is not one that has exited. Until 6.3.1
# each of these read it with SilentlyContinue: setup step 10 killed the
# Cowork VM with Claude perhaps running, step 8 said the worker had
# exited, step 13 said no antivirus was found, and -Undo step 3 said a
# monitor was stopped.
Invoke-Section 'Processes that cannot be read' {
    $denied = 'Access is denied'
    function Get-Seen { return "$(Get-StepStatus): $(Get-StepNotes)" }

    Reset-State
    [void]$script:HcsLive.Add($guidA)
    $script:ProcReadFault = @('claude')
    Invoke-Step10
    Assert-Case "setup step 10, the processes named claude cannot be listed: no VM is killed, and a Warning says why" ($script:Killed.Count -eq 0 -and
        (Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq "Could not list the processes named claude, so it is not known whether Claude is running, and the 1 cowork-vm instance(s) are left alone: $denied") (
        "$(Get-Seen); killed $($script:Killed -join ',')")

    Reset-State
    Set-VmFixture
    $script:ProcReadFault = @('10')
    Invoke-Step8
    Assert-Case "setup step 8, the Cowork VM's worker cannot be read: a Warning, and it is not raised" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq "Could not read the Cowork VM worker (PID 10), so its priority was not checked or raised: $denied" -and
        (Get-Priority 10) -eq 'Normal' -and -not (Test-LogHas 'has exited')) (Get-Seen)

    Reset-State
    $script:ProcReadFault = @('*')
    Invoke-Step13
    $said = "Could not list the running processes, so an antivirus that is not Defender may go unreported: $denied"
    Assert-Case "setup step 13, the running processes cannot be listed: a Warning, and only what was read is said" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq $said -and -not (Test-LogHas '^No antivirus product detected$') -and (Test-LogHas '^No antivirus product is registered with Windows$')) (Get-Seen)

    Reset-State
    $script:AvProducts = @([pscustomobject]@{ displayName = 'Windows Defender' })
    $script:ProcReadFault = @('*')
    Invoke-Step13
    Assert-Case "and with Defender registered, its exclusions are added all the same" ($script:MpAdds.Count -gt 0 -and
        @((Get-StepNotes) -split ' \| ') -contains $said) "$(Get-Seen); adds $($script:MpAdds.Count)"

    Reset-State
    $script:PsProcs = @([pscustomobject]@{ ProcessId = 4242; CommandLine = 'powershell.exe -NoProfile -File "C:\ClaudeFix\Watch-ClaudeHealth.ps1" -Quiet' })
    $script:Procs = @(New-FakeProcess 4242 'powershell')
    $script:ProcReadFault = @('4242')
    Invoke-UndoStep3
    Assert-Case "-Undo step 3, whether a stopped monitor ended cannot be read: Failed, and not said to be stopped" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq "Could not read whether the health monitor (PID 4242) stopped: $denied" -and -not (Test-LogHas '^Stopped running health monitor')) (Get-Seen)
}

# =====================================================================
# Setup step 14 reads the WSL feature and the running distros. Until
# 6.3.1 a feature that could not be read was said only as a grey line, and
# wsl.exe's exit code and errors were not looked at: any failure read as
# "No WSL2 distro running".
Invoke-Section 'Setup step 14: WSL, and what cannot be read' {
    function Set-Wsl {
        param([string[]]$Lines, [int]$Exit)
        # A wsl.exe of its own, so that the one Reset-State names stays absent.
        $script:FakeWslPath = Join-Path $sandbox 'wsl\wsl.exe'
        $null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path (Split-Path $script:FakeWslPath -Parent) -Force
        [System.IO.File]::WriteAllText($script:FakeWslPath, '')
        $script:WslLines = $Lines
        $script:WslExit = $Exit
    }
    function Get-Seen { return "$(Get-StepStatus): $(Get-StepNotes); wsl: $($script:WslCalls -join ' / '); $($script:LogLines -join ' / ')" }
    $noDistro = '^No WSL2 distro running$'
    $none = 'Windows Subsystem for Linux has no installed distributions.'
    $wslUtf8 = $env:WSL_UTF8

    Reset-State
    Invoke-Step14
    Assert-Case "no wsl.exe: Unchanged, with no distro running, and wsl.exe not run" ((Get-StepStatus) -eq 'Unchanged' -and (Get-StepNotes) -eq '' -and
        (Test-LogHas $noDistro) -and $script:WslCalls.Count -eq 0) (Get-Seen)

    # 6.3.1: until then wsl.exe was looked for with Test-Path, which answers
    # $false for a file that cannot be read as for one that is not there.
    Reset-State
    Set-Wsl @('  NAME      STATE           VERSION', '* Ubuntu    Running         2') 0
    $script:ItemReadFault = $script:FakeWslPath
    Invoke-Step14
    Assert-Case "a wsl.exe that cannot be read: a Warning that names it and why, and no distro said to run or not" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq "Could not read which WSL distros are running: $($script:FakeWslPath) could not be read (Access is denied)" -and
        $script:WslCalls.Count -eq 0 -and -not (Test-LogHas $noDistro)) (Get-Seen)

    Reset-State
    $script:WslFeatureState = 'Enabled'
    Set-Wsl @('  NAME      STATE           VERSION', '* Ubuntu    Running         2', '  Debian    Stopped         2', '  Legacy    Running         1') 0
    Invoke-Step14
    Assert-Case "a WSL2 distro running: a Warning naming it, and only it" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq 'WSL2 running (Ubuntu): if Cowork has issues, run wsl --shutdown' -and ($script:WslCalls -join ' / ') -eq '-l -v' -and
        (Test-LogHas '^WSL is installed$') -and $env:WSL_UTF8 -eq $wslUtf8) (Get-Seen)

    Reset-State
    Set-Wsl @($none) 1
    Invoke-Step14
    Assert-Case "wsl.exe fails while neither WSL service is there: no distro runs, Unchanged" ((Get-StepStatus) -eq 'Unchanged' -and
        (Get-StepNotes) -eq '' -and (Test-LogHas $noDistro)) (Get-Seen)

    Reset-State
    Set-Wsl @($none) 1
    $script:Services['WSLService'] = 'Stopped'
    $script:Services['LxssManager'] = 'Stopped'
    Invoke-Step14
    Assert-Case "and while both are there and stopped: the same" ((Get-StepStatus) -eq 'Unchanged' -and (Test-LogHas $noDistro)) (Get-Seen)

    Reset-State
    Set-Wsl @('', 'Access is denied.') 1
    $script:Services['WSLService'] = 'Running'
    Invoke-Step14
    $said = 'Could not read which WSL distros are running: wsl.exe -l -v ended with exit code 1 while WSLService runs: Access is denied.'
    Assert-Case "wsl.exe fails while WSLService runs: a Warning that says so, not 'No WSL2 distro running'" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq $said -and -not (Test-LogHas $noDistro)) (Get-Seen)

    Reset-State
    Set-Wsl @($none) 1
    $script:Services['LxssManager'] = 'Running'
    Invoke-Step14
    Assert-Case "and while LxssManager runs: the same, naming it" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq "Could not read which WSL distros are running: wsl.exe -l -v ended with exit code 1 while LxssManager runs: $none") (Get-Seen)

    Reset-State
    Set-Wsl @($none) 1
    $script:ServiceReadFault = @('WSLService')
    Invoke-Step14
    Assert-Case "wsl.exe fails and a WSL service cannot be read: a Warning" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq 'Could not read which WSL distros are running: Access is denied' -and -not (Test-LogHas $noDistro)) (Get-Seen)

    Reset-State
    Set-Wsl @() 0
    $script:WslThrows = $true
    Invoke-Step14
    Assert-Case "wsl.exe cannot be run: a Warning" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq 'Could not read which WSL distros are running: The system cannot execute the specified program.' -and
        $env:WSL_UTF8 -eq $wslUtf8) (Get-Seen)

    Reset-State
    $script:WslFeatureFault = $true
    Invoke-Step14
    Assert-Case "the WSL feature cannot be read: a Warning that says so" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq 'Could not read whether the WSL feature is installed: Class not registered' -and -not (Test-LogHas '^WSL is installed$')) (Get-Seen)
}

# =====================================================================
# Power plans that cannot be read. Until 6.3.1 every error reading the
# plans in the registry read as no Ultimate Performance plan: setup made
# one more copy, and -Undo took the plan an older version recorded for
# yours. A record that could not be read was taken for one that held no
# plan: Balanced was made active, and the record deleted.
Invoke-Section 'Power plans that cannot be read' {
    $why = 'the power plans in the registry could not be read (Requested registry access is not allowed.)'
    function Get-SetActive { return @($script:PcCalls | Where-Object { $_ -like '/setactive*' }).Count }
    function Get-Seen { return "$(Get-StepStatus): $(Get-StepNotes); active $($script:ActivePlan); $($script:PcCalls -join ' / ')" }

    # A PowerSchemes key that is not there holds no plan, and is not an
    # error: Get-ChildItem says so with PathNotFound, and that alone.
    Reset-State
    $guids = $null
    $err = $null
    try { $guids = @(Get-UltimatePlanGuids) } catch { $err = $_.Exception.Message }
    $schemesKeys = @($script:Reg.Keys | Where-Object { $_.StartsWith((Get-RegNorm $schemesRoot)) }).Count
    Assert-Case "no PowerSchemes key: no Ultimate plan, and no error" ($null -eq $err -and $null -ne $guids -and $guids.Count -eq 0 -and
        $schemesKeys -eq 0) "error: $err; keys: $schemesKeys"

    Reset-State
    Set-PowerFixture
    $script:RegReadFault = $schemesRoot
    Invoke-SetupStep1
    Assert-Case "setup step 1, the plans cannot be read from the registry: Failed, and no copy is made" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq $why -and @($script:PcCalls | Where-Object { $_ -like '/duplicatescheme*' }).Count -eq 0 -and
        $null -eq (Find-Undo 'PowerPlan' 'Created') -and $null -eq (Find-Undo 'PowerPlan' 'Active')) (Get-Seen)

    # An older version's record of the plan you had that cannot be read was
    # taken, until 6.3.1, for one that held no plan: Balanced was recorded as
    # yours, and the record deleted.
    Reset-State
    Set-PowerFixture
    $script:ActivePlan = $plSaver
    [System.IO.File]::WriteAllText($BackupFile, "$plHigh`r`n")
    $lock = [System.IO.File]::Open($BackupFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
    try { Invoke-SetupStep1 } finally { $lock.Dispose() }
    $head = "the older version's record of your plan, $BackupFile, could not be read ("
    Assert-Case "setup step 1, an older version's record cannot be read: Failed, nothing recorded, and the record kept" ((Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes).StartsWith($head) -and (Get-StepNotes).EndsWith('), so the plan you had is not recorded yet, and the record is kept') -and
        $null -eq (Find-Undo 'PowerPlan' 'Active') -and (Get-LegacyInstallCount) -eq 0 -and
        (Microsoft.PowerShell.Management\Test-Path -LiteralPath $BackupFile) -and (Get-SetActive) -eq 0) (Get-Seen)

    Reset-State
    Set-PowerFixture
    $script:ActivePlan = $plSaver
    [System.IO.File]::WriteAllText($BackupFile, "$plHigh`r`n")
    $script:RegReadFault = $schemesRoot
    Invoke-UndoStep2
    Assert-Case "-Undo step 2, the plans cannot be read: the plan an older version recorded is not put back, and its record is kept: Failed" (
        (Get-StepStatus) -eq 'Failed' -and
        @((Get-StepNotes) -split ' \| ') -contains "The plan an older version recorded was not put back, because $why; its record is kept for the next -Undo" -and
        (Microsoft.PowerShell.Management\Test-Path -LiteralPath $BackupFile) -and $script:ActivePlan -eq $plSaver -and (Get-SetActive) -eq 0) (Get-Seen)

    Reset-State
    Set-PowerFixture
    $script:ActivePlan = $plSaver
    [System.IO.File]::WriteAllText($BackupFile, "$plHigh`r`n")
    $lock = [System.IO.File]::Open($BackupFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
    try { Invoke-UndoStep2 } finally { $lock.Dispose() }
    $head = "The plan an older version recorded was not put back, because $BackupFile could not be read ("
    Assert-Case "-Undo step 2, the older version's record cannot be read: not taken for an empty one, and kept: Failed" ((Get-StepStatus) -eq 'Failed' -and
        @((Get-StepNotes) -split ' \| ' | Where-Object { $_.StartsWith($head) -and $_.EndsWith('); its record is kept for the next -Undo') }).Count -eq 1 -and
        (Microsoft.PowerShell.Management\Test-Path -LiteralPath $BackupFile) -and $script:ActivePlan -eq $plSaver -and (Get-SetActive) -eq 0) (Get-Seen)
}

# =====================================================================
# The scripts Prevent writes for later: the Claude (Admin) task's payload
# and the GPU-free launcher. Each is run here, in this process, against
# fakes: Invoke-GeneratedScript turns each exit into a return of
# 'EXIT n', and refuses to run a script that reaches a command that is
# neither its own, nor in the list it is given, nor one of four that only
# shape what they are given. Until 6.3.1 both read the process list with
# SilentlyContinue, so a list that could not be read was taken for no
# Claude running: the payload started a second Claude, and the launcher
# started one beside a Claude it had not closed, then said it ran without
# the GPU.
function New-GenClaude {
    param([int]$Id, [string]$Exe = 'C:\FakeApps\Claude\app\Claude.exe')
    return [pscustomobject]@{ Id = $Id; ProcessName = 'claude'; PriorityClass = 'Normal'; MainModule = [pscustomobject]@{ FileName = $Exe } }
}
function Reset-Generated {
    Reset-State
    $script:GenHost = New-Object System.Collections.ArrayList
    $script:GenPopups = New-Object System.Collections.ArrayList
    $script:GenStarts = New-Object System.Collections.ArrayList
    $script:GenAppx = @([pscustomobject]@{ Name = 'Claude'; PackageFamilyName = 'Claude_test'; InstallLocation = 'C:\FakeApps\Claude' })
    $script:GenFiles = @('C:\FakeApps\Claude\app\Claude.exe')
    $script:GenProduct = @{ 'C:\FakeApps\Claude\app\Claude.exe' = 'Claude' }
    $script:GenStartRuns = $true
    $script:GenFaultAfterStop = $false
    $script:GenFaultAfterStart = $false
    $script:GenItemFault = $null
    $script:GenAppxFault = $false
}
function Invoke-GeneratedScript {
    param([string]$Text, [string[]]$Allowed)
    $shapers = @('Where-Object', 'ForEach-Object', 'Select-Object', 'Join-Path')
    $errs = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($Text, [ref]$null, [ref]$errs)
    if (@($errs).Count -gt 0) { return "REFUSED: it does not parse: $($errs[0].Message)" }
    $exits = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.ExitStatementAst] }, $true) |
        Sort-Object { $_.Extent.StartOffset } -Descending)
    foreach ($x in $exits) {
        $code = '0'
        if ($null -ne $x.Pipeline) { $code = $x.Pipeline.Extent.Text }
        $Text = $Text.Substring(0, $x.Extent.StartOffset) + "return 'EXIT $code'" + $Text.Substring($x.Extent.EndOffset)
    }
    $ast = [System.Management.Automation.Language.Parser]::ParseInput($Text, [ref]$null, [ref]$errs)
    if (@($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.ExitStatementAst] }, $true)).Count -gt 0) { return 'REFUSED: an exit is left' }
    $own = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object { $_.Name })
    $bad = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { [string]$_.GetCommandName() } |
        Where-Object { $own -notcontains $_ -and $Allowed -notcontains $_ -and $shapers -notcontains $_ } | Select-Object -Unique)
    if ($bad.Count -gt 0) { return "REFUSED: it runs $($bad -join ', ')" }

    function Write-Host { param([Parameter(Position = 0)]$Object, $ForegroundColor) [void]$script:GenHost.Add([string]$Object) }
    function Read-Host { param([Parameter(Position = 0)]$Prompt) [void]$script:GenHost.Add("READ $Prompt") }
    # Each of these is seen by every function the script calls as well, the
    # fakes among them, since PowerShell scopes are dynamic: none may be one
    # that those call.
    function New-GenShell {
        $sh = [pscustomobject]@{ Kind = 'FakeShell' }
        $sh | Add-Member -MemberType ScriptMethod -Name Popup -Value { [void]$script:GenPopups.Add([string]$args[0]); return 1 }
        return $sh
    }
    # $script:GenAppxFault: the package list cannot be read. The payload
    # takes any error from it for that, so the error's identity is not
    # what is tested.
    function Get-AppxPackage {
        [CmdletBinding()]
        param()
        if ($script:GenAppxFault) {
            $PSCmdlet.WriteError((New-FakeErrorRecord 'System.UnauthorizedAccessException' 'Access is denied.' 'GetAppxPackageError' 'PermissionDenied' $null))
            return
        }
        return @($script:GenAppx)
    }
    # $script:GenItemFault is a -like pattern: a file it matches cannot be
    # read. Test-Path then answers $false, and Get-Item gives the error
    # Get-Item -Force gives under an explicit Deny; a file not in
    # $script:GenFiles or $script:GenProduct is not there (both measured
    # 2026-11-23).
    function Test-Path {
        [CmdletBinding()]
        param([Parameter(Position = 0)][string]$Path, [string]$LiteralPath)
        $p = Select-FakePath $Path $LiteralPath
        if ($script:GenItemFault -and $p -like $script:GenItemFault) { return $false }
        return (@($script:GenFiles) -contains $p)
    }
    function Get-Item {
        [CmdletBinding()]
        param([string]$LiteralPath, [switch]$Force)
        $null = $Force
        if ($script:GenItemFault -and $LiteralPath -like $script:GenItemFault) {
            $PSCmdlet.WriteError((New-FakeErrorRecord 'System.UnauthorizedAccessException' 'Access is denied' `
                'ItemExistsUnauthorizedAccessError' 'PermissionDenied' $LiteralPath))
            return
        }
        if (@($script:GenFiles) -notcontains $LiteralPath -and -not $script:GenProduct.ContainsKey($LiteralPath)) {
            $PSCmdlet.WriteError((New-FakeErrorRecord 'System.Management.Automation.ItemNotFoundException' `
                "Cannot find path '$LiteralPath' because it does not exist." 'PathNotFound' 'ObjectNotFound' $LiteralPath))
            return
        }
        return [pscustomobject]@{ PSIsContainer = $false; VersionInfo = [pscustomobject]@{ ProductName = [string]$script:GenProduct[$LiteralPath] } }
    }
    function Start-Process {
        param([Parameter(Position = 0)][string]$FilePath, [string[]]$ArgumentList)
        [void]$script:GenStarts.Add(("$FilePath " + ($ArgumentList -join ' ')).Trim())
        if ($script:GenStartRuns) { $script:Procs = @($script:Procs) + @(New-GenClaude 99) }
        if ($script:GenFaultAfterStart) { $script:ProcReadFault = @('claude') }
    }
    function Stop-Process {
        [CmdletBinding()]
        param([Parameter(ValueFromPipeline = $true)]$InputObject, [switch]$Force)
        process {
            $null = $Force
            [void]$script:ProcCalls.Add("Stop-Process $($InputObject.Id)")
            if (-not $script:StopProcLinger) { $script:Procs = @($script:Procs | Where-Object { $_.Id -ne $InputObject.Id }) }
            if ($script:GenFaultAfterStop) { $script:ProcReadFault = @('claude') }
        }
    }
    foreach ($n in @($Allowed)) {
        $c = $null
        try { $c = Get-Command -Name $n -ErrorAction Stop } catch { return "REFUSED: $n is not found: $($_.Exception.Message)" }
        if ($c.CommandType -ne 'Function') { return "REFUSED: $n is not a fake" }
    }
    Set-StrictMode -Off
    try {
        return (@(& ([scriptblock]::Create($Text))) -join ' / ')
    } catch {
        return "THREW: $($_.Exception.Message)"
    }
}

Invoke-Section 'The Claude (Admin) task payload: whether Claude already runs' {
    $m = [regex]::Match($pSrc, '(?s)\$launchCmd = @''\r?\n(.*?)\r?\n''@')
    Assert-Case "the payload is found in Prevent" $m.Success
    $shellCall = 'New-Object -ComObject WScript.Shell'
    $payload = $m.Groups[1].Value
    Assert-Case "and it makes its message boxes through WScript.Shell, in its two places" (([regex]::Matches($payload, [regex]::Escape($shellCall))).Count -eq 2)
    $payload = $payload.Replace($shellCall, 'New-GenShell')
    $allowed = @('Get-Process', 'Get-Item', 'New-GenShell', 'Get-AppxPackage', 'Test-Path', 'Start-Process')
    $app = 'C:\FakeApps\Claude\app\Claude.exe'
    $out = ''
    function Get-Seen { return "out: $out; popups: $($script:GenPopups -join ' / '); starts: $($script:GenStarts -join ' / ')" }

    Reset-Generated
    $out = Invoke-GeneratedScript $payload $allowed
    Assert-Case "no Claude running: Claude is started" ($out -eq '' -and ($script:GenStarts -join ' / ') -eq $app -and $script:GenPopups.Count -eq 0) (Get-Seen)

    Reset-Generated
    $script:Procs = @(New-GenClaude 77)
    $out = Invoke-GeneratedScript $payload $allowed
    Assert-Case "Claude running: it ends, and starts nothing" ($out -eq 'EXIT 0' -and $script:GenStarts.Count -eq 0 -and $script:GenPopups.Count -eq 0) (Get-Seen)

    Reset-Generated
    $script:Procs = @(New-GenClaude 78 'C:\Users\tester\AppData\Roaming\Claude\claude-code\2.1.0\claude.exe')
    $out = Invoke-GeneratedScript $payload $allowed
    Assert-Case "only the Claude Code CLI running: Claude is started" ($out -eq '' -and ($script:GenStarts -join ' / ') -eq $app) (Get-Seen)

    Reset-Generated
    $script:ProcReadFault = @('claude')
    $out = Invoke-GeneratedScript $payload $allowed
    $said = "Claude was not started: the running processes could not be read, so it is not known whether Claude is already running.`n`nAccess is denied"
    Assert-Case "the processes cannot be listed: Claude is not started, and a message says why" ($out -eq 'EXIT 1' -and $script:GenStarts.Count -eq 0 -and
        ($script:GenPopups -join ' / ') -eq $said) (Get-Seen)

    # 6.3.1: a place Claude may be installed that cannot be read is a doubt,
    # not a place where it is not. Until 6.3.1 it was looked for with
    # Test-Path, which answers $false for both, and Claude was started from
    # a later place: an older install the user had left.
    $first = Join-Path $env:LOCALAPPDATA 'Programs\claude-desktop\Claude.exe'
    $third = Join-Path $env:LOCALAPPDATA 'Programs\Claude\Claude.exe'

    Reset-Generated
    $script:GenFiles = @($app, $third)
    $script:GenItemFault = $app
    $out = Invoke-GeneratedScript $payload $allowed
    Assert-Case "the app's own Claude.exe cannot be read: nothing is started from another place, and a message says why" ($out -eq 'EXIT 1' -and
        $script:GenStarts.Count -eq 0 -and ($script:GenPopups -join ' / ') -eq "Claude was not started: $app could not be read: Access is denied") (Get-Seen)

    Reset-Generated
    $script:GenAppx = @()
    $script:GenFiles = @($first, $third)
    $script:GenItemFault = $first
    $out = Invoke-GeneratedScript $payload $allowed
    Assert-Case "an installer's place cannot be read: nothing is started from a later one, and a message says why" ($out -eq 'EXIT 1' -and
        $script:GenStarts.Count -eq 0 -and ($script:GenPopups -join ' / ') -eq "Claude was not started: $first could not be read: Access is denied") (Get-Seen)

    Reset-Generated
    $script:GenAppx = @()
    $script:GenFiles = @($third)
    $out = Invoke-GeneratedScript $payload $allowed
    Assert-Case "no app package, and the installer's third place holds it: Claude is started from there" ($out -eq '' -and
        ($script:GenStarts -join ' / ') -eq $third -and $script:GenPopups.Count -eq 0) (Get-Seen)

    Reset-Generated
    $script:GenAppx = @()
    $script:GenFiles = @()
    $out = Invoke-GeneratedScript $payload $allowed
    Assert-Case "Claude installed nowhere: nothing is started, and that is what is said" ($out -eq 'THREW: Claude Desktop not found. Is it installed?' -and
        $script:GenStarts.Count -eq 0 -and $script:GenPopups.Count -eq 0) (Get-Seen)

    Reset-Generated
    $script:GenFiles = @($app, $third)
    $script:GenAppxFault = $true
    $out = Invoke-GeneratedScript $payload $allowed
    Assert-Case "the app packages cannot be listed: nothing is started, and a message says why" ($out -eq 'EXIT 1' -and $script:GenStarts.Count -eq 0 -and
        ($script:GenPopups -join ' / ') -eq 'Claude was not started: Access is denied.') (Get-Seen)
}

Invoke-Section 'The GPU-free launcher: a Claude that is not closed, and a list that cannot be read' {
    $as = $pAst.Find({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$ps1Content' }, $true)
    $tpl = $null
    if ($null -ne $as -and $as.Right -is [System.Management.Automation.Language.CommandExpressionAst] -and
        $as.Right.Expression -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) { $tpl = $as.Right.Expression }
    $vars = @()
    if ($null -ne $tpl) { $vars = @($tpl.NestedExpressions | ForEach-Object { $_.Extent.Text } | Sort-Object -Unique) }
    Assert-Case "the launcher's template is found, and expands only `$howNote, `$launchBody and `$ToolkitVersion" ($null -ne $tpl -and
        ($vars -join ',') -eq '$howNote,$launchBody,$ToolkitVersion') ($vars -join ',')
    if ($null -eq $tpl) { return }
    $app = 'C:\FakeApps\Claude\app\Claude.exe'
    $howNote = 'Traditional install, flag passed directly'
    $launchBody = "Start-Process -FilePath '$app' -ArgumentList '--disable-gpu'"
    $ToolkitVersion = '0.0.0-test'
    $launcher = [string](& ([scriptblock]::Create($tpl.Extent.Text)))
    $allowed = @('Get-Process', 'Get-Item', 'Write-Host', 'Read-Host', 'Stop-Process', 'Start-Sleep', 'Start-Process')
    $out = ''
    function Get-Seen { return "out: $out; stops: $($script:ProcCalls -join ' / '); starts: $($script:GenStarts -join ' / '); $($script:GenHost -join ' / ')" }
    function Test-HostSaid { param([string]$Line) return (@($script:GenHost | Where-Object { ([string]$_).Trim() -eq $Line }).Count -eq 1) }
    $runs = 'Claude is running without hardware GPU acceleration.'
    $notClosed = '[!] Claude could not be closed, or whether it closed could not be read, so it was not started again without the GPU. Close Claude and run this again.'

    Reset-Generated
    $script:Procs = @(New-GenClaude 77)
    $out = Invoke-GeneratedScript $launcher $allowed
    Assert-Case "Claude running: it is closed, started again without the GPU, and said to run so" ($out -eq '' -and ($script:ProcCalls -join ' / ') -eq 'Stop-Process 77' -and
        ($script:GenStarts -join ' / ') -eq "$app --disable-gpu" -and (Test-HostSaid $runs)) (Get-Seen)

    Reset-Generated
    $out = Invoke-GeneratedScript $launcher $allowed
    Assert-Case "no Claude running: started without the GPU, and nothing stopped" ($out -eq '' -and $script:ProcCalls.Count -eq 0 -and
        ($script:GenStarts -join ' / ') -eq "$app --disable-gpu" -and (Test-HostSaid $runs)) (Get-Seen)

    Reset-Generated
    $script:Procs = @(New-GenClaude 77)
    $script:StopProcLinger = $true
    $out = Invoke-GeneratedScript $launcher $allowed
    Assert-Case "a Claude that is not closed: it is not started again, and a message says so" ($out -eq 'EXIT 1' -and $script:GenStarts.Count -eq 0 -and
        (Test-HostSaid $notClosed) -and -not (Test-HostSaid $runs)) (Get-Seen)

    Reset-Generated
    $script:Procs = @(New-GenClaude 77)
    $script:GenFaultAfterStop = $true
    $out = Invoke-GeneratedScript $launcher $allowed
    Assert-Case "a Claude whose closing cannot be read: the same" ($out -eq 'EXIT 1' -and $script:GenStarts.Count -eq 0 -and (Test-HostSaid $notClosed)) (Get-Seen)

    Reset-Generated
    $script:Procs = @(New-GenClaude 77)
    $script:ProcReadFault = @('claude')
    $out = Invoke-GeneratedScript $launcher $allowed
    Assert-Case "the processes cannot be listed: nothing closed or started, and a message says why" ($out -eq 'EXIT 1' -and $script:ProcCalls.Count -eq 0 -and
        $script:GenStarts.Count -eq 0 -and (Test-HostSaid '[!] The running processes could not be read, so Claude was not closed or started: Access is denied')) (Get-Seen)

    Reset-Generated
    $script:GenStartRuns = $false
    $out = Invoke-GeneratedScript $launcher $allowed
    Assert-Case "a start that does not take: said not to have started" ($out -eq '' -and
        (Test-HostSaid '[!] Claude did not start. Launch it normally and report this.') -and -not (Test-HostSaid $runs)) (Get-Seen)

    Reset-Generated
    $script:GenFaultAfterStart = $true
    $out = Invoke-GeneratedScript $launcher $allowed
    Assert-Case "a start whose result cannot be read: said so, not said to run" ($out -eq '' -and
        (Test-HostSaid '[!] Claude was started, but the running processes could not be read to check that it runs.') -and -not (Test-HostSaid $runs)) (Get-Seen)

    # 6.3.1: a Claude process whose path cannot be read could be Claude
    # started as administrator, which takes the launch and ignores the flag.
    # Until 6.3.1 the launch went ahead, and Claude was said to run without
    # the GPU. MainModule throws for it, as for a process this one may not
    # open.
    $unreadSaid = '[!] 1 running Claude process(es) could not be read (one started as administrator, or a Claude Code session), so Claude was not closed or started without the GPU: a copy already running takes the launch and ignores the flag. Close Claude and run this again.'
    $shut = [pscustomobject]@{ Id = 79; ProcessName = 'claude'; PriorityClass = 'Normal' }
    $shut | Add-Member -MemberType ScriptProperty -Name MainModule -Value { throw 'Access is denied' }
    Reset-Generated
    $script:Procs = @($shut)
    $out = Invoke-GeneratedScript $launcher $allowed
    Assert-Case "a Claude process whose path cannot be read: nothing closed or started, and a message says why" ($out -eq 'EXIT 1' -and
        $script:ProcCalls.Count -eq 0 -and $script:GenStarts.Count -eq 0 -and (Test-HostSaid $unreadSaid) -and -not (Test-HostSaid $runs)) (Get-Seen)

    Reset-Generated
    $script:Procs = @((New-GenClaude 77), $shut)
    $out = Invoke-GeneratedScript $launcher $allowed
    Assert-Case "and beside a Claude that can be read: that one is not closed either" ($out -eq 'EXIT 1' -and
        $script:ProcCalls.Count -eq 0 -and $script:GenStarts.Count -eq 0 -and (Test-HostSaid $unreadSaid)) (Get-Seen)
}

# =====================================================================
Invoke-Section 'What the script says it does' {
    $first = $pAst.EndBlock.Statements[0]
    $isRefusal = $first -is [System.Management.Automation.Language.IfStatementAst] -and $first.Clauses.Count -eq 1 -and
        $first.Clauses[0].Item1.Extent.Text -eq '$WhatIfPreference' -and $null -eq $first.ElseClause
    $ends = $false
    if ($isRefusal) {
        $body = @($first.Clauses[0].Item2.Statements)
        $last = $body[$body.Count - 1]
        $ends = $last -is [System.Management.Automation.Language.ExitStatementAst] -and $last.Pipeline.Extent.Text -eq '1'
    }
    Assert-Case "-WhatIf is refused before anything else runs" ($isRefusal -and $ends) $first.Extent.Text.Split("`n")[0]
    $attrs = (@($pAst.ParamBlock.Attributes) | ForEach-Object { $_.Extent.Text }) -join ' '
    Assert-Case "and still declared, so it reaches the refusal" ($attrs -match 'SupportsShouldProcess') $attrs
    Assert-Case "and never carried into the elevated run" ($pSrc -notmatch '(?i)elevateArgs\s*\+=\s*"\s*-WhatIf')

    $setupAt = $pSrc.IndexOf('$steps = 19')
    $undoAt = $pSrc.IndexOf('$steps = 5')
    $setupNames = @()
    $undoNames = @()
    if ($setupAt -gt 0 -and $undoAt -gt 0 -and $undoAt -lt $setupAt) {
        $rx = '(?m)^\s*Step (\d+) \$steps "[^"]*" -Name "([^"]+)"'
        $setupNames = @([regex]::Matches($pSrc.Substring($setupAt), $rx) | ForEach-Object { "$($_.Groups[1].Value). $($_.Groups[2].Value)" })
        $undoNames = @([regex]::Matches($pSrc.Substring($undoAt, $setupAt - $undoAt), $rx) | ForEach-Object { "$($_.Groups[1].Value). $($_.Groups[2].Value)" })
    }
    $wantSetup = @('Power plan', 'Full performance plan settings', 'Hibernate', 'Fast Startup', 'Modern Standby',
                   'Network adapter power saving', 'Power throttling', 'Cowork VM priority', 'vmcompute recovery', 'Stale VM removal',
                   'Storage location', 'Clock', 'Antivirus', 'WSL2 conflicts', 'Health monitor',
                   'Fix Claude Desktop shortcuts', 'Claude (Admin) launcher', 'GPU-free launcher', "Older versions' leftovers")
    $want = @(for ($i = 0; $i -lt $wantSetup.Count; $i++) { "$($i + 1). $($wantSetup[$i])" })
    Assert-Case "setup runs these 19 steps, in order" (($setupNames -join ' / ') -eq ($want -join ' / ')) ($setupNames -join ' / ')
    $wantUndo = @('1. Recorded settings', "2. Older versions' changes", '3. Monitor and tasks', '4. Fix Claude Desktop shortcuts', '5. Launchers')
    Assert-Case "-Undo runs these 5" (($undoNames -join ' / ') -eq ($wantUndo -join ' / ')) ($undoNames -join ' / ')

    $tail = $pAst.EndBlock.Statements[$pAst.EndBlock.Statements.Count - 1]
    Assert-Case "the script ends with the exit code its steps earned" ($tail -is [System.Management.Automation.Language.ExitStatementAst] -and
        $tail.Pipeline.Extent.Text -eq '$exitCode' -and $null -ne (Find-TopAssignment $pAst '$exitCode') -and
        (Find-TopAssignment $pAst '$exitCode').Right.Extent.Text -eq 'Get-PreventExitCode') $tail.Extent.Text

    $soft = @($pAst.FindAll({
        param($n)
        ($n -is [System.Management.Automation.Language.StringConstantExpressionAst] -or
         $n -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) -and $n.Value -match '(?i)not critical'
    }, $true))
    Assert-Case "no message calls a failure 'not critical'" ($soft.Count -eq 0) "$($soft.Count) found"

    $setupIf = $pAst.Find({
        param($n)
        $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -eq '$Undo' -and
        $null -ne $n.ElseClause -and $n.ElseClause.Extent.Text.Contains('$steps = 19')
    }, $true)
    # The setup branch, and every Prevent function it reaches through any
    # number of calls: a retired change moved into a function such as
    # Disable-NicPowerSaving is still made by setup.
    $fnDefs = @{}
    foreach ($fd in $pAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) { $fnDefs[$fd.Name] = $fd }
    $setupCmds = New-Object System.Collections.ArrayList
    $reached = @{}
    if ($null -ne $setupIf) {
        $bodies = New-Object System.Collections.Queue
        $bodies.Enqueue($setupIf.ElseClause)
        while ($bodies.Count -gt 0) {
            $body = $bodies.Dequeue()
            foreach ($c in $body.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
                [void]$setupCmds.Add($c)
                $cn = [string]$c.GetCommandName()
                if ($cn -and $fnDefs.ContainsKey($cn) -and -not $reached.ContainsKey($cn)) {
                    $reached[$cn] = $true
                    $bodies.Enqueue($fnDefs[$cn].Body)
                }
            }
        }
    }
    $bad = @()
    if ($null -ne $setupIf) {
        foreach ($c in $setupCmds) {
            $name = [string]$c.GetCommandName()
            $t = $c.Extent.Text
            if (@('New-NetNat', 'Remove-NetNat', 'New-NetFirewallRule', 'Set-NetFirewallRule', 'Set-VMMemory', 'Set-VM', 'Restart-NetAdapter', 'Disable-NetAdapter') -contains $name) { $bad += $t }
            if (@('Set-ItemProperty', 'New-ItemProperty', 'Write-RecordedDword') -contains $name -and
                $t -match 'ServicesPipeTimeout|LocalAccountTokenFilterPolicy|CsEnabled') { $bad += $t }
            if ($name -eq 'Register-ScheduledTask' -and $t -match 'BootTaskName') { $bad += $t }
            if ($name -eq 'sc.exe' -and $t -match '\$ServiceName') { $bad += $t }
        }
    }
    Assert-Case "setup no longer makes the retired changes" ($null -ne $setupIf -and $bad.Count -eq 0) (($bad | ForEach-Object { $_.Split("`n")[0] }) -join ' / ')
    # Disable-NicPowerSaving is called by step 6, and Get-PreventNicAdvanced
    # only by functions: both reached shows the check follows the calls.
    Assert-Case "and that check reads the functions setup calls, and the ones they call" ($reached.ContainsKey('Disable-NicPowerSaving') -and
        $reached.ContainsKey('Get-PreventNicAdvanced') -and -not $setupIf.ElseClause.Extent.Text.Contains('Get-PreventNicAdvanced')) (
        "$($reached.Count) reached: $((@($reached.Keys) | Sort-Object) -join ', ')")
    $restarts = @($pAst.FindAll({
        param($n)
        $n -is [System.Management.Automation.Language.CommandAst] -and @('Restart-NetAdapter', 'Disable-NetAdapter', 'Enable-NetAdapter') -contains [string]$n.GetCommandName()
    }, $true))
    Assert-Case "no network adapter is restarted anywhere" ($restarts.Count -eq 0)
    $retired = @($wAst.FindAll({
        param($n)
        ($n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
         @('Set-DynamicMemoryFlag', 'Test-WinNatHealth', 'Test-GuestConnectionHealth') -contains $n.Name) -or
        ($n -is [System.Management.Automation.Language.CommandAst] -and
         @('Set-DynamicMemoryFlag', 'Test-WinNatHealth', 'Test-GuestConnectionHealth', 'Set-VMMemory', 'New-NetNat') -contains [string]$n.GetCommandName())
    }, $true))
    Assert-Case "the monitor's retired checks are gone" ($retired.Count -eq 0) "$($retired.Count) found"

    # No ClaudeFix version set Claude.exe to run as administrator through the
    # AppCompatFlags Layers key, so the RUNASADMIN mark there is the user's
    # own. Until 6.3.0 -Undo deleted every such mark on a name containing
    # Claude.exe; nothing in the script may name that key or the mark now.
    $runAs = @($pAst.FindAll({
        param($n)
        ($n -is [System.Management.Automation.Language.StringConstantExpressionAst] -or
         $n -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) -and $n.Value -match '(?i)RUNASADMIN|AppCompatFlags'
    }, $true))
    Assert-Case "-Undo leaves the user's run-as-administrator marks alone" ($runAs.Count -eq 0) (($runAs | ForEach-Object { $_.Extent.Text }) -join ' / ')
}

# =====================================================================
# 6.3.1: a file or folder that cannot be read is not one that is not
# there. Test-Path answers $false for both (measured 2026-11-23, an explicit
# Deny on the folder that holds it), and until 6.3.1 each place below used
# it: -Undo said there was nothing to undo or nothing left, setup recorded
# the wrong plan as yours, installed the monitor or made the shortcuts from
# another copy, and said it had checked what it had not.
# $script:ItemReadFault makes a path unreadable to the fake Get-Item and
# Test-Path alike.
Invoke-Section 'Files and folders that cannot be read are not ones that are not there' {
    $denied = 'Access is denied'
    function Get-Seen { return "$(Get-StepStatus): $(Get-StepNotes); unread $($script:UndoUnread); undo: $(Get-UndoKeys); calls: $($script:TaskCalls -join ' / '); $($script:LogLines -join ' / ')" }
    function Test-File { param([string]$Path) return (Microsoft.PowerShell.Management\Test-Path -LiteralPath $Path) }
    $flag = Join-Path $ClaudeAppData 'disable-dynamic-memory.flag'

    # The Claude AppData folder, which setup makes first.
    Reset-State
    $script:ItemReadFault = $ClaudeAppData
    $r = $null
    $threw = ''
    try { $r = Initialize-ClaudeAppData } catch { $threw = $_.Exception.Message }
    Assert-Case "the Claude AppData folder cannot be read: setup stops, naming it and why" ($threw -eq "$ClaudeAppData could not be read ($denied)") "returned $r; threw $threw"
    Reset-State
    Microsoft.PowerShell.Management\Remove-Item -LiteralPath $ClaudeAppData -Recurse -Force
    $r = Initialize-ClaudeAppData
    Assert-Case "and a folder that is not there is made" ($r -eq $true -and (Test-File $ClaudeAppData)) "returned $r"

    # The undo record.
    Reset-State
    [System.IO.File]::WriteAllText($UndoFile, '{"Version":1,"Items":[]}')
    $script:ItemReadFault = $UndoFile
    $threw = ''
    try { Read-PreventUndo } catch { $threw = $_.Exception.Message }
    Assert-Case "an undo record that cannot be read: the run stops before anything is changed, naming it and why" (
        $threw -eq "The undo record $UndoFile could not be read ($denied). Nothing was changed. Move the file aside and rerun.") $threw
    Reset-State
    [System.IO.File]::WriteAllText($UndoFile, '{"Version":1,"Items":[]}')
    $script:UndoRecord = New-Object System.Collections.ArrayList
    $script:ItemReadFault = $UndoFile
    $threw = ''
    try { Save-PreventUndo } catch { $threw = $_.Exception.Message }
    Assert-Case "an emptied record over one that cannot be read: it says so, and the file is kept" (
        $threw -eq "The undo record $UndoFile could not be read ($denied), so it was not emptied." -and (Test-File $UndoFile)) $threw

    # Remove-PreventFile, which -Undo steps 3 to 5 remove files with.
    $f = Join-Path $ClaudeAppData 'Launch-Claude-NoGPU.cmd'
    Reset-State
    $script:UndoMode = $true
    [System.IO.File]::WriteAllText($f, 'x')
    $script:ItemReadFault = $f
    Step 5 5 'launchers' -Name 'Launchers'
    $r = Remove-PreventFile -Path $f
    Assert-Case "a file to remove that cannot be read: kept, counted as there, a Warning that names it and why, and -Undo not complete" ($r -eq $true -and
        (Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -eq "$f could not be read ($denied), so it was not removed" -and $script:UndoUnread -and
        (Test-File $f)) (Get-Seen)
    Reset-State
    [System.IO.File]::WriteAllText($f, 'x')
    $script:ItemReadFaultAfterRemove = $f
    Step 5 5 'launchers' -Name 'Launchers'
    $r = Remove-PreventFile -Path $f
    Assert-Case "a file removed, whether it is gone then cannot be read: Failed, and not said removed" ($r -eq $true -and (Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes) -eq "Could not remove ${f}: whether it is gone could not be read ($denied)" -and -not (Test-LogHas '^Removed: ')) (Get-Seen)

    # -Undo step 4.
    $desk = Join-Path $fakeDesktop 'Fix Claude Desktop.lnk'
    Reset-State
    $script:UndoMode = $true
    [System.IO.File]::WriteAllText($desk, 'lnk')
    $script:ItemReadFault = $desk
    Invoke-UndoStep4
    Assert-Case "-Undo step 4, a shortcut that cannot be read: a Warning that names it, no 'No shortcuts found', and -Undo not complete" (
        (Get-StepStatus) -eq 'Warning' -and (Get-StepNotes) -eq "$desk could not be read ($denied), so it was not removed" -and
        -not (Test-LogHas '^No shortcuts found$') -and $script:UndoUnread -and (Test-File $desk)) (Get-Seen)

    # -Undo step 2, where an older version ran: the record of that keeps the
    # files from being read for it, so each is read only here.
    foreach ($c in @(
            @($BackupFile, 'the plan it holds was not put back, and it was not removed'),
            @($CsBackupFile, 'Connected Standby was not put back from it, and it was not removed'),
            @($flag, 'it was not removed'))) {
        Reset-State
        Set-PowerFixture
        $script:UndoMode = $true
        Add-LegacyTrace
        [System.IO.File]::WriteAllText($c[0], "$plHigh`r`n")
        $script:ItemReadFault = $c[0]
        Invoke-UndoStep2
        Assert-Case "-Undo step 2, $(Split-Path $c[0] -Leaf) cannot be read: a Warning that names it, the mark of an older version kept, and -Undo not complete" (
            (Get-StepStatus) -eq 'Warning' -and (Get-StepNotes).Contains("$($c[0]) could not be read ($denied), so $($c[1])") -and
            $null -ne (Find-Undo 'LegacyInstall' 'Dynamic Memory flag') -and $script:UndoUnread -and (Test-File $c[0]) -and
            -not (Test-LogHas '^Nothing from older versions found$')) (Get-Seen)
    }

    # Setup step 1: an older version's record of your plan.
    Reset-State
    Set-PowerFixture
    [System.IO.File]::WriteAllText($BackupFile, "$plHigh`r`n")
    $script:ItemReadFault = $BackupFile
    Invoke-SetupStep1
    Assert-Case "setup step 1, an older version's plan record that cannot be read: the step stops, no plan is recorded as yours, and the record is kept" (
        (Get-StepStatus) -eq 'Failed' -and $null -eq (Find-Undo 'PowerPlan' 'Active') -and (Test-File $BackupFile) -and
        (Get-StepNotes).Contains("the older version's record of your plan, $BackupFile, could not be read ($denied), so the plan you had is not recorded yet, and the record is kept")) (Get-Seen)
    Reset-State
    Set-PowerFixture
    Add-PreventUndoItem -Kind PowerPlan -Key Active -Original $plBal
    [System.IO.File]::WriteAllText($BackupFile, "$plHigh`r`n")
    $script:ItemReadFault = $BackupFile
    Invoke-SetupStep1
    Assert-Case "and with your plan recorded already: a Warning that it was not recorded or removed, and it is kept" (
        (Get-StepNotes).Contains("$BackupFile could not be read ($denied), so it was not recorded or removed") -and (Get-StepStatus) -eq 'Warning' -and
        $null -eq (Find-Undo 'LegacyInstall' 'power plan record') -and (Test-File $BackupFile)) (Get-Seen)

    # Setup step 15: the monitor's script, and the old watchdog script.
    $kit = Join-Path $sandbox 'kit-unread'
    $null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $kit -Force
    $beside = Join-Path $kit 'Watch-ClaudeHealth.ps1'
    [System.IO.File]::WriteAllText($beside, '# the monitor')
    $null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $fakeKitRoot -Force
    $other = Join-Path $fakeKitRoot 'Watch-ClaudeHealth.ps1'
    [System.IO.File]::WriteAllText($other, '# another copy')
    Reset-State
    $script:TaskInstallExpected = $true
    $script:FakeCommandPath = Join-Path $kit 'Prevent-ClaudeIssues.ps1'
    $script:ItemReadFault = $beside
    Invoke-Step15
    Assert-Case "setup step 15, the monitor's script beside Prevent cannot be read: Failed, saying why, and the copy in C:\ClaudeFix not installed" (
        (Get-StepStatus) -eq 'Failed' -and $script:TaskCalls.Count -eq 0 -and $null -eq $script:RegisteredTask -and
        (Get-StepNotes) -eq "Could not look for Watch-ClaudeHealth.ps1, because $beside could not be read ($denied), so the health monitor was not installed") (Get-Seen)
    Microsoft.PowerShell.Management\Remove-Item -LiteralPath $other -Force
    $oldWd = Join-Path $ClaudeAppData 'cowork-watchdog.ps1'
    Reset-State
    [System.IO.File]::WriteAllText($oldWd, '# old')
    $script:ItemReadFault = $oldWd
    Step 15 19 'monitor' -Name 'Health monitor'
    Invoke-Step15Legacy
    Assert-Case "an old watchdog script that cannot be read: kept, not recorded, and a Warning that names it" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq "$oldWd could not be read ($denied), so it was not recorded or removed" -and
        $null -eq (Find-Undo 'LegacyInstall' 'watchdog script') -and (Test-File $oldWd)) (Get-Seen)

    # Setup step 16: the Fix Claude Desktop shortcuts.
    $kit16 = Join-Path $sandbox 'kit16'
    $null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $kit16 -Force
    $bat = Join-Path $kit16 'Fix-ClaudeDesktop.bat'
    [System.IO.File]::WriteAllText($bat, '@echo off')
    function Get-Targets { return (@($script:SavedShortcuts | ForEach-Object { [string]$_.TargetPath }) -join ',') }
    Reset-State
    $script:FakeCommandPath = Join-Path $kit16 'Prevent-ClaudeIssues.ps1'
    Invoke-Step16
    Assert-Case "setup step 16: both shortcuts start the launcher beside Prevent" ((Get-StepStatus) -eq 'Done' -and (Get-Targets) -eq "$bat,$bat") "$(Get-Seen); $(Get-Targets)"
    $null = Microsoft.PowerShell.Management\New-Item -ItemType Directory -Path $fakeKitRoot -Force
    $otherPs1 = Join-Path $fakeKitRoot 'Fix-ClaudeDesktop.ps1'
    $otherBat = Join-Path $fakeKitRoot 'Fix-ClaudeDesktop.bat'
    [System.IO.File]::WriteAllText($otherPs1, '# another copy')
    [System.IO.File]::WriteAllText($otherBat, '@echo off')
    Reset-State
    $script:FakeCommandPath = Join-Path $kit16 'Prevent-ClaudeIssues.ps1'
    $script:ItemReadFault = $bat
    Invoke-Step16
    Assert-Case "the launcher beside Prevent cannot be read: Failed, saying why, and no shortcut made to the copy in C:\ClaudeFix" (
        (Get-StepStatus) -eq 'Failed' -and $script:SavedShortcuts.Count -eq 0 -and
        (Get-StepNotes) -eq "Could not look for Fix-ClaudeDesktop.bat, because $bat could not be read ($denied), so the shortcuts were not made") "$(Get-Seen); $(Get-Targets)"
    Reset-State
    $script:FakeCommandPath = Join-Path $sandbox 'elsewhere\Prevent-ClaudeIssues.ps1'
    Invoke-Step16
    Assert-Case "no launcher beside Prevent: the shortcuts start the copy beside the Fix script in C:\ClaudeFix" ((Get-StepStatus) -eq 'Done' -and
        (Get-Targets) -eq "$otherBat,$otherBat") "$(Get-Seen); $(Get-Targets)"
    Microsoft.PowerShell.Management\Remove-Item -LiteralPath $otherPs1, $otherBat -Force

    # Setup step 17: the .cmd launcher versions up to 6.1.0 made.
    $oldCmd = Join-Path $ClaudeAppData 'Launch-Claude-Admin.cmd'
    Reset-State
    [System.IO.File]::WriteAllText($oldCmd, '@echo off')
    $script:ItemReadFault = $oldCmd
    Step 17 19 'admin' -Name 'Claude (Admin) launcher'
    Invoke-Step17Legacy
    Assert-Case "setup step 17, an old .cmd launcher that cannot be read: kept, not recorded, and a Warning that names it" ((Get-StepStatus) -eq 'Warning' -and
        (Get-StepNotes) -eq "$oldCmd could not be read ($denied), so it was not recorded or removed" -and
        $null -eq (Find-Undo 'LegacyInstall' 'admin launcher') -and (Test-File $oldCmd)) (Get-Seen)

    # Setup step 19, where an older version ran.
    foreach ($c in @(
            @($flag, 'it was not recorded or removed'),
            @($CsBackupFile, 'whether Connected Standby is left as an older version set it is not known'))) {
        Reset-State
        Add-LegacyTrace
        [System.IO.File]::WriteAllText($c[0], '0')
        $script:ItemReadFault = $c[0]
        Invoke-Step19
        Assert-Case "setup step 19, $(Split-Path $c[0] -Leaf) cannot be read: a Warning that names it, and no 'Nothing left from older versions'" (
            (Get-StepStatus) -eq 'Warning' -and (Get-StepNotes).Contains("$($c[0]) could not be read ($denied), so $($c[1])") -and
            -not (Test-LogHas '^Nothing left from older versions$') -and (Test-File $c[0])) (Get-Seen)
    }
    Reset-State
}

Invoke-Section 'Registry keys that cannot be opened are not ones with no value' {
    # Until 6.3.1 Get-RegistryValue read any error as a value that is not
    # there: a step recorded "there was none" before it wrote one, and -Undo
    # said it removed a value it never looked at.
    $refused = 'Requested registry access is not allowed.'
    $policyPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
    function Get-Seen { return "$(Get-StepStatus): $(Get-StepNotes); unread $($script:UndoUnread); undo: $(Get-UndoKeys); calls: $($script:Calls -join ' / '); $($script:LogLines -join ' / ')" }
    function New-UndoItem { param([string]$Kind, [string]$Key, $Original) return [pscustomobject]@{ Kind = $Kind; Key = $Key; Original = $Original } }

    Reset-State
    Set-FakeRegValue 'HKLM:\T\Shut' 'V' ([int]5)
    Set-FakeRegValue 'HKLM:\T\Open' 'V' ([int]5)
    $script:RegItemReadFault = 'HKLM:\T\Shut'
    $v = 'unset'
    $err = $null
    try { $v = Get-RegistryValue -Path 'HKLM:\T\Shut' -Name 'V' } catch { $err = $_.Exception.Message }
    Assert-Case "Get-RegistryValue, a key that cannot be opened: throws, naming the value, the key and why" (
        $err -eq "The registry value V in HKLM:\T\Shut could not be read ($refused)") "got '$v'; error: $err"
    $err = $null
    $a = 'unset'; $b2 = 'unset'; $c = 'unset'
    try {
        $a = Get-RegistryValue -Path 'HKLM:\T\Open' -Name 'V'
        $b2 = Get-RegistryValue -Path 'HKLM:\T\Open' -Name 'Absent'
        $c = Get-RegistryValue -Path 'HKLM:\T\NoSuchKey' -Name 'V'
    } catch { $err = $_.Exception.Message }
    Assert-Case "and a value that is there, one that is not and a key that is not are read as before" (
        $null -eq $err -and $a -eq 5 -and $null -eq $b2 -and $null -eq $c) "error: $err; got '$a' '$b2' '$c'"

    $r = 'unset'
    $err = $null
    try { $r = Write-RecordedDword -Path 'HKLM:\T\Shut' -Name 'V' -Value 1 } catch { $err = $_.Exception.Message }
    Assert-Case "Write-RecordedDword, a key that cannot be opened: throws, and nothing is recorded or written" (
        $err -eq "The registry value V in HKLM:\T\Shut could not be read ($refused)" -and $script:UndoRecord.Count -eq 0 -and
        $script:Calls.Count -eq 0 -and (Get-FakeRegValue 'HKLM:\T\Shut' 'V') -eq 5) "got $r; error: $err; $(Get-UndoKeys); calls: $($script:Calls -join ' / ')"

    # The MachineGuid is best effort: the record is then told apart by the
    # computer's name, as for a record that holds none.
    Reset-State
    Set-FakeRegValue 'HKLM:\SOFTWARE\Microsoft\Cryptography' 'MachineGuid' 'abc'
    $script:RegItemReadFault = 'HKLM:\SOFTWARE\Microsoft\Cryptography'
    $g = 'unset'
    $err = $null
    try { $g = Get-PreventMachineGuid } catch { $err = $_.Exception.Message }
    Assert-Case "Get-PreventMachineGuid, its key cannot be opened: '', and the run goes on" ($null -eq $err -and $g -eq '') "got '$g'; error: $err"

    # -Undo step 1, a value recorded as none.
    Reset-State
    Set-PowerFixture
    Set-FakeRegValue 'HKLM:\T\Shut' 'V' ([int]1)
    $null = $script:UndoRecord.Add((New-UndoItem 'Registry' 'HKLM:\T\Shut|V' $null))
    $script:RegItemReadFault = 'HKLM:\T\Shut'
    Invoke-UndoStep1
    Assert-Case "-Undo step 1, a value recorded as none in a key that cannot be opened: Failed, naming it, not said removed, and its record kept" (
        (Get-StepStatus) -eq 'Failed' -and
        @((Get-StepNotes) -split ' \| ') -contains "Registry HKLM:\T\Shut|V: The registry value V in HKLM:\T\Shut could not be read ($refused)" -and
        -not (Test-LogHas 'there was none') -and $null -ne (Find-Undo 'Registry' 'HKLM:\T\Shut|V') -and
        (Get-FakeRegValue 'HKLM:\T\Shut' 'V') -eq 1) (Get-Seen)

    # -Undo step 2, where an older version ran.
    Reset-State
    Add-LegacyTrace
    Set-FakeRegValue $policyPath 'LocalAccountTokenFilterPolicy' ([int]1)
    $script:RegItemReadFault = $policyPath
    Invoke-UndoStep2
    Assert-Case "-Undo step 2, the policy key cannot be opened: Failed, naming it, the value left, and the mark of an older version kept" (
        (Get-StepStatus) -eq 'Failed' -and
        @((Get-StepNotes) -split ' \| ') -contains "The registry value LocalAccountTokenFilterPolicy in $policyPath could not be read ($refused)" -and
        (Get-FakeRegValue $policyPath 'LocalAccountTokenFilterPolicy') -eq 1 -and $null -ne (Find-Undo 'LegacyInstall' 'Dynamic Memory flag') -and
        -not (Test-LogHas '^Nothing from older versions found$')) (Get-Seen)

    # Setup step 19.
    Reset-State
    Set-FakeRegValue $policyPath 'LocalAccountTokenFilterPolicy' ([int]1)
    $script:RegItemReadFault = $policyPath
    Invoke-Step19
    Assert-Case "setup step 19, the policy key cannot be opened: Failed, naming it, and no 'Nothing left from older versions'" (
        (Get-StepStatus) -eq 'Failed' -and
        (Get-StepNotes).Contains("The registry value LocalAccountTokenFilterPolicy in $policyPath could not be read ($refused)") -and
        -not (Test-LogHas '^Nothing left from older versions$')) (Get-Seen)
    Reset-State
}

Invoke-Section 'A file that cannot be removed is named' {
    # Until 6.3.1 these removals let Remove-Item's error through as it came,
    # and for a file the user may not delete it does not name the file
    # (measured 2026-11-23): the step said what failed, but not on which file.
    function Get-Seen { return "$(Get-StepStatus): $(Get-StepNotes); undo: $(Get-UndoKeys); $($script:LogLines -join ' / ')" }
    function Test-File { param([string]$Path) return (Microsoft.PowerShell.Management\Test-Path -LiteralPath $Path) }
    $refusedRm = 'Access to the path is denied.'
    $flag = Join-Path $ClaudeAppData 'disable-dynamic-memory.flag'

    # Emptying the undo record.
    Reset-State
    [System.IO.File]::WriteAllText($UndoFile, '{"Version":1,"Items":[]}')
    $script:UndoRecord = New-Object System.Collections.ArrayList
    $script:FileRemoveFault = $UndoFile
    $threw = ''
    try { Save-PreventUndo } catch { $threw = $_.Exception.Message }
    Assert-Case "an emptied record that cannot be removed: it throws, naming it and why, and the file is kept" (
        $threw -eq "$UndoFile could not be removed ($refusedRm)" -and (Test-File $UndoFile)) $threw

    # -Undo step 2, an older version's record of your plan: once your plan
    # is put back from it, and once step 1 has put it back from this
    # script's own record.
    foreach ($restored in @($false, $true)) {
        Reset-State
        Set-PowerFixture
        $script:UndoMode = $true
        Add-LegacyTrace
        $script:FakePlanRestored = $restored
        [System.IO.File]::WriteAllText($BackupFile, "$plHigh`r`n")
        $script:FileRemoveFault = $BackupFile
        Invoke-UndoStep2
        Assert-Case "-Undo step 2, the older version's plan record cannot be removed (plan put back by step 1: $restored): Failed, naming it and why, and the mark of an older version kept" (
            (Get-StepStatus) -eq 'Failed' -and (Get-StepNotes).Contains("$BackupFile could not be removed ($refusedRm)") -and
            $null -ne (Find-Undo 'LegacyInstall' 'Dynamic Memory flag') -and (Test-File $BackupFile)) (Get-Seen)
    }

    # -Undo step 2, Connected Standby: put back, and its record not removed.
    Reset-State
    Set-PowerFixture
    $script:UndoMode = $true
    Add-LegacyTrace
    Set-FakeRegValue $powerKey 'CsEnabled' ([int]0)
    [System.IO.File]::WriteAllText($CsBackupFile, '1')
    $script:FileRemoveFault = $CsBackupFile
    Invoke-UndoStep2
    Assert-Case "-Undo step 2, CsEnabled put back and its record cannot be removed: Failed, saying it was put back, naming the record and why" (
        (Get-StepStatus) -eq 'Failed' -and
        @((Get-StepNotes) -split ' \| ') -contains "CsEnabled (Connected Standby) was put back to 1, but $CsBackupFile could not be removed ($refusedRm)" -and
        (Get-FakeRegValue $powerKey 'CsEnabled') -eq 1 -and (Get-LastStep).Reboot -and -not (Test-LogHas 'Could not put back CsEnabled') -and
        (Test-File $CsBackupFile)) (Get-Seen)

    # -Undo step 2, the Dynamic Memory flag.
    Reset-State
    Set-PowerFixture
    $script:UndoMode = $true
    Add-LegacyTrace
    [System.IO.File]::WriteAllText($flag, '')
    $script:FileRemoveFault = $flag
    Invoke-UndoStep2
    Assert-Case "-Undo step 2, the Dynamic Memory flag cannot be removed: Failed, naming it and why, and the mark of an older version kept" (
        (Get-StepStatus) -eq 'Failed' -and (Get-StepNotes).Contains("$flag could not be removed ($refusedRm)") -and
        $null -ne (Find-Undo 'LegacyInstall' 'Dynamic Memory flag') -and (Test-File $flag)) (Get-Seen)

    # Setup step 1, an older version's record of your plan.
    Reset-State
    Set-PowerFixture
    [System.IO.File]::WriteAllText($BackupFile, "$plHigh`r`n")
    $script:FileRemoveFault = $BackupFile
    Invoke-SetupStep1
    Assert-Case "setup step 1, an older version's plan record that cannot be removed: Failed, naming it and why, and it is kept" (
        (Get-StepStatus) -eq 'Failed' -and (Get-StepNotes).Contains("$BackupFile could not be removed ($refusedRm)") -and
        (Test-File $BackupFile)) (Get-Seen)

    # Setup step 19, the Dynamic Memory flag.
    Reset-State
    Add-LegacyTrace
    [System.IO.File]::WriteAllText($flag, '0')
    $script:FileRemoveFault = $flag
    Invoke-Step19
    Assert-Case "setup step 19, the Dynamic Memory flag cannot be removed: Failed, naming it and why, and it is kept" (
        (Get-StepStatus) -eq 'Failed' -and (Get-StepNotes).Contains("$flag could not be removed ($refusedRm)") -and
        (Test-File $flag)) (Get-Seen)
    Reset-State
}

# ---------------------------------------------------------------------
Write-Host ""
Assert-Case "no tripwire was touched" ($script:Tripped.Count -eq 0) ($script:Tripped -join ' / ')
# Best effort: every case is counted by now, and a folder left in %TEMP%
# changes nothing.
Microsoft.PowerShell.Management\Remove-Item -LiteralPath $sandbox -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host ("Prevent steps: {0} passed, {1} failed" -f $script:pass, $script:fail)
if ($script:fail -eq 0) { Write-Host "Prevent steps gate PASSED"; exit 0 }
Write-Host "Prevent steps gate FAILED"
exit 1
