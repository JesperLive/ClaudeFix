<#
.SYNOPSIS
    Exercises Backup-CoworkVhdx's purge decision against synthetic inputs.
.DESCRIPTION
    Pulls the function body out of Fix-ClaudeDesktop.ps1 with the language
    parser, defines it in this session, and drives it through the cases that
    decide whether 13.5 GB of VM cache gets deleted.

    Nothing destructive runs. Sources and the backup folder are small temporary
    files, and the free-space query is stubbed so the low-space and
    unknown-space branches can actually be reached.

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

function Get-FunctionSource {
    param([string]$Path, [string]$Name)
    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    if (@($errors).Count -gt 0) { throw "parse errors in $Path" }
    $fn = $ast.Find({
        param($n)
        $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name
    }, $true)
    if (-not $fn) { throw "function '$Name' not found in $Path" }
    return $fn.Extent.Text
}

$fixPath = Join-Path $Root 'Fix-ClaudeDesktop.ps1'
Invoke-Expression (Get-FunctionSource -Path $fixPath -Name 'Backup-CoworkVhdx')
# Get-ClaudeErrorText and Get-ClaudeItemOrNull, which Fix's functions call,
# live in the shared region.
. (Join-Path $Root '.ci\ClaudeEnv.region.ps1')
Write-Host "Extracted Backup-CoworkVhdx from Fix-ClaudeDesktop.ps1"

# -- Stubs the extracted function depends on ---------------------------
$script:PreservedVhdx = @('sessiondata.vhdx', 'smol-bin.vhdx')
$script:LogLines = @()
function Log {
    param([string]$Message, [string]$Colour = "White", [switch]$Indent)
    $script:LogLines += $Message
}

$script:FakeSources    = @{}
$script:FakeFreeBytes  = [long]0
$script:FakeHeaderOk   = $true

# A source of THROWS is one that could not be looked for: the real one
# throws for it since 6.3.1.
function Get-CoworkVhdxSource {
    param([Parameter(Mandatory)][string]$Name)
    if ($script:FakeSources.ContainsKey($Name)) {
        $v = $script:FakeSources[$Name]
        if ($v -is [string] -and $v -eq 'THROWS') { throw 'C:\nx\cache could not be searched in full (Access is denied)' }
        return $v
    }
    return $null
}
function Test-VhdxHeader {
    param([string]$Path)
    return $script:FakeHeaderOk
}
function Get-PSDrive {
    # CmdletBinding supplies -ErrorAction as a common parameter. Without it the
    # real call site's "-ErrorAction Stop" fails to bind, the stub throws a
    # binding error, and every case silently lands in the unknown-space branch.
    [CmdletBinding()]
    param([Parameter(Position = 0)]$Name)
    if ($script:FakeFreeBytes -lt 0) { throw "free space unavailable" }
    return [pscustomobject]@{ Name = "$Name"; Free = $script:FakeFreeBytes }
}

# -- Harness -----------------------------------------------------------
$sandbox = Join-Path $env:TEMP ("cfix-backuptest-" + [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item $sandbox -ItemType Directory -Force | Out-Null

function New-FakeVhdx {
    param([string]$Name, [int]$SizeKb)
    $dir = Join-Path $sandbox "src"
    if (-not (Test-Path $dir)) { New-Item $dir -ItemType Directory -Force | Out-Null }
    $p = Join-Path $dir $Name
    $bytes = New-Object byte[] ($SizeKb * 1024)
    [System.IO.File]::WriteAllBytes($p, $bytes)
    return (Get-Item $p)
}

$pass = 0
$fail = 0
function Assert-Case {
    param([string]$Name, [bool]$Condition, [string]$Detail = "")
    if ($Condition) {
        Write-Host ("  PASS  {0}" -f $Name)
        $script:pass++
    } else {
        Write-Host ("  FAIL  {0}  {1}" -f $Name, $Detail)
        $script:fail++
    }
}

function Reset-Case {
    param([string]$Label)
    Write-Host ""
    Write-Host ("--- {0}" -f $Label)
    $script:LogLines = @()
    $bd = Join-Path $sandbox ("bk-" + [guid]::NewGuid().ToString("N").Substring(0, 6))
    return $bd
}

# Case 1: nothing to preserve. Purging is safe because there is nothing to lose.
$bd = Reset-Case "no preserved VHDX present"
$script:FakeSources = @{}
$script:FakeFreeBytes = [long]500GB
$r = Backup-CoworkVhdx -BackupDir $bd
Assert-Case "SourceFound is false" (-not $r.SourceFound)
Assert-Case "SafeToPurge is true"  ($r.SafeToPurge) "nothing to lose, purge should proceed"
Assert-Case "BackedUp is empty"    (@($r.BackedUp.Keys).Count -eq 0)

# Case 2: both files present, ample space, copies succeed.
$bd = Reset-Case "both files present, ample space"
$script:FakeSources = @{
    'sessiondata.vhdx' = (New-FakeVhdx -Name 'sessiondata.vhdx' -SizeKb 64)
    'smol-bin.vhdx'    = (New-FakeVhdx -Name 'smol-bin.vhdx'    -SizeKb 16)
}
$script:FakeFreeBytes = [long]500GB
$script:FakeHeaderOk  = $true
$r = Backup-CoworkVhdx -BackupDir $bd
Assert-Case "SourceFound is true" ($r.SourceFound)
Assert-Case "SafeToPurge is true" ($r.SafeToPurge)
Assert-Case "both files recorded" (@($r.BackedUp.Keys).Count -eq 2) ("got " + @($r.BackedUp.Keys).Count)
Assert-Case "records the ORIGINAL source path, not the backup path" `
    ($r.BackedUp['sessiondata.vhdx'] -eq $script:FakeSources['sessiondata.vhdx'].FullName) `
    $r.BackedUp['sessiondata.vhdx']
Assert-Case "backup file actually landed" (Test-Path (Join-Path $bd 'sessiondata.vhdx'))
Assert-Case "no .tmp left behind" (-not (Test-Path (Join-Path $bd 'sessiondata.vhdx.tmp')))

# Case 3: not enough free space. This is the 720MB literal bug.
$bd = Reset-Case "insufficient free space"
$script:FakeFreeBytes = [long]1MB
$r = Backup-CoworkVhdx -BackupDir $bd
Assert-Case "SafeToPurge is false" (-not $r.SafeToPurge) "purge must not run when the backup cannot fit"
Assert-Case "nothing was backed up" (@($r.BackedUp.Keys).Count -eq 0)
Assert-Case "reason was logged" (@($script:LogLines | Where-Object { $_ -match 'Not enough space' }).Count -ge 1)

# Case 4: free space cannot be determined. Unknown must mean do-not-purge.
$bd = Reset-Case "free space unknown"
$script:FakeFreeBytes = [long](-1)
$r = Backup-CoworkVhdx -BackupDir $bd
Assert-Case "SafeToPurge is false" (-not $r.SafeToPurge) "unknown space must not authorise a purge"
Assert-Case "reason was logged" (@($script:LogLines | Where-Object { $_ -match 'unknown' }).Count -ge 1)

# Case 5: the copy lands but fails its header check. Session data is not safe.
$bd = Reset-Case "VHDX header check fails"
$script:FakeFreeBytes = [long]500GB
$script:FakeHeaderOk  = $false
$r = Backup-CoworkVhdx -BackupDir $bd
Assert-Case "SafeToPurge is false" (-not $r.SafeToPurge) "a corrupt backup must not authorise a purge"
Assert-Case "sessiondata not recorded" (-not $r.BackedUp.ContainsKey('sessiondata.vhdx'))
Assert-Case "discarded .tmp is cleaned up" (-not (Test-Path (Join-Path $bd 'sessiondata.vhdx.tmp')))

# Case 6: the backup directory does not exist yet.
#
# New-Item is ShouldProcess-gated, so a -WhatIf run never creates it. Deriving
# the drive with Get-Item on the missing path threw, free space came back
# "unknown", and the purge was refused, which made every dry run report the
# opposite of what a real run does.
$bd = Reset-Case "backup directory does not exist yet"
$script:FakeHeaderOk = $true
$script:FakeFreeBytes = [long]500GB
$script:FakeSources = @{
    'sessiondata.vhdx' = (New-FakeVhdx -Name 'sessiondata.vhdx' -SizeKb 64)
}
$missingDir = Join-Path $sandbox ("never-created-" + [guid]::NewGuid().ToString("N").Substring(0, 6))
Assert-Case "the directory really is absent to start with" (-not (Test-Path $missingDir))
$r = Backup-CoworkVhdx -BackupDir $missingDir
Assert-Case "free space is still determined" `
    (@($script:LogLines | Where-Object { $_ -match 'unknown' }).Count -eq 0) `
    ($script:LogLines -join ' | ')
Assert-Case "SafeToPurge is true" ($r.SafeToPurge)

# Case 7: only smol-bin present. Session data absent means nothing to protect.
$bd = Reset-Case "only smol-bin present"
$script:FakeHeaderOk = $true
$script:FakeSources = @{ 'smol-bin.vhdx' = (New-FakeVhdx -Name 'smol-bin.vhdx' -SizeKb 16) }
$r = Backup-CoworkVhdx -BackupDir $bd
Assert-Case "SourceFound is true" ($r.SourceFound)
Assert-Case "SafeToPurge is true" ($r.SafeToPurge) "no sessiondata means nothing irreplaceable"

# Case 8, since 6.3.1: whether a preserved VHDX is in the cache could not be
# read. Not purging, and that is said. Until 6.3.1 it read as nothing to
# preserve, and the cache was purged with sessiondata.vhdx in it.
$bd = Reset-Case "a preserved VHDX that could not be looked for"
$script:FakeHeaderOk = $true
$script:FakeFreeBytes = [long]500GB
$script:FakeSources = @{ 'sessiondata.vhdx' = 'THROWS' }
$r = $null
$escaped = $null
try { $r = Backup-CoworkVhdx -BackupDir $bd } catch { $escaped = $_.Exception.Message }
Assert-Case "nothing escapes" ($null -eq $escaped) "$escaped"
Assert-Case "SafeToPurge is false" ($null -ne $r -and -not $r.SafeToPurge) "the VHDX may be there, so the cache must stay"
Assert-Case "and why is said" `
    (@($script:LogLines | Where-Object {
        $_ -eq '[!] Whether sessiondata.vhdx is in the VM cache could not be read (C:\nx\cache could not be searched in full (Access is denied)); not purging' }).Count -eq 1) `
    ($script:LogLines -join ' | ')

# Get-CoworkVhdxSource itself, lifted and run against fakes of what it reads.
# Since 6.3.1 a place a VHDX could be that could not be read throws, so the
# purge above is refused; a place that is not there is passed over. Until
# 6.3.1 each read as no VHDX there. The fakes write the errors the cmdlets
# write: "not there" with the identity measured 2026-11-22 on Windows 11,
# Windows PowerShell 5.1 (error id PathNotFound, category ObjectNotFound,
# System.Management.Automation.ItemNotFoundException), and "could not be
# read" with access denied. Every path is under C:\nx, which is never
# created, and every read of it is faked.
Write-Host ""
Write-Host "--- Get-CoworkVhdxSource: a place that could not be read"
$found = & {
    Set-StrictMode -Version Latest
    Invoke-Expression (Get-FunctionSource -Path $fixPath -Name 'Get-CoworkVhdxSource')
    # Get-CoworkVhdxSource names a search's first error with
    # Get-FixFirstError. Until 6.3.1 a lift of it that failed was passed
    # over, and the cases below then failed on a command not found rather
    # than on the reason; now the lift stops the gate and says why.
    Invoke-Expression (Get-FunctionSource -Path $fixPath -Name 'Get-FixFirstError')
    function New-BkRecord {
        param([string]$TypeName, [string]$Message, [string]$Id, [string]$Category, $Target)
        $ex = Microsoft.PowerShell.Utility\New-Object $TypeName $Message
        return (Microsoft.PowerShell.Utility\New-Object System.Management.Automation.ErrorRecord($ex, $Id,
                [System.Management.Automation.ErrorCategory]$Category, $Target))
    }
    # BkItems maps a path to its item, or to Denied; a path not listed is not
    # there. BkScan maps a folder to what a search of it finds and whether
    # part of it could not be read.
    function Get-Item {
        [CmdletBinding()]
        param([Parameter(Position = 0)][string]$Path, [string]$LiteralPath, [switch]$Force)
        $null = $Force
        $p = if ($LiteralPath) { $LiteralPath } else { $Path }
        if (-not $script:BkItems.ContainsKey($p)) {
            $PSCmdlet.WriteError((New-BkRecord 'System.Management.Automation.ItemNotFoundException' `
                "Cannot find path '$p' because it does not exist." 'PathNotFound' 'ObjectNotFound' $p))
            return
        }
        $v = $script:BkItems[$p]
        if ($v -is [string] -and $v -eq 'Denied') {
            $PSCmdlet.WriteError((New-BkRecord 'System.UnauthorizedAccessException' 'Access is denied' `
                'ItemExistsUnauthorizedAccessError' 'PermissionDenied' $p))
            return
        }
        return $v
    }
    function Test-Path {
        param([Parameter(Position = 0)][string]$Path, [string]$LiteralPath)
        $p = if ($LiteralPath) { $LiteralPath } else { $Path }
        $v = $script:BkItems[$p]
        return ($null -ne $v -and -not ($v -is [string] -and $v -eq 'Denied'))
    }
    function Get-ChildItem {
        [CmdletBinding()]
        param([Parameter(Position = 0)][string]$Path, [string]$LiteralPath, [string]$Filter, [switch]$Recurse)
        $null = $Filter, $Recurse
        $p = if ($LiteralPath) { $LiteralPath } else { $Path }
        $s = $script:BkScan[$p]
        if ($null -eq $s) { return }
        if ($s['Denied']) {
            $PSCmdlet.WriteError((New-BkRecord 'System.UnauthorizedAccessException' "Access to the path '$p\sub' is denied." `
                'DirUnauthorizedAccessError' 'PermissionDenied' "$p\sub"))
        }
        if ($s['Hit']) { $s['Hit'] }
    }
    $cacheItem = [pscustomobject]@{ FullName = 'C:\nx\cache' }
    $hitItem   = [pscustomobject]@{ Name = 'sessiondata.vhdx'; FullName = 'C:\nx\cache\sub2\sessiondata.vhdx'; Length = 1MB }
    $knownItem = [pscustomobject]@{ Name = 'sessiondata.vhdx'; FullName = 'C:\nx\known\sessiondata.vhdx'; Length = 1MB }
    $known     = @([pscustomobject]@{ Name = 'sessiondata.vhdx'; Path = 'C:\nx\known\sessiondata.vhdx' })
    $o = [ordered]@{}
    foreach ($c in @(
            @{ N = 'knownUnread';  Known = $known; Items = @{ 'C:\nx\known\sessiondata.vhdx' = 'Denied' } },
            @{ N = 'folderUnread'; Known = @();    Items = @{ 'C:\nx\cache' = 'Denied' } },
            @{ N = 'scanUnread';   Known = @();    Items = @{ 'C:\nx\cache' = $cacheItem }; Scan = @{ 'C:\nx\cache' = @{ Denied = $true } } },
            @{ N = 'hit';          Known = @();    Items = @{ 'C:\nx\cache' = $cacheItem }; Scan = @{ 'C:\nx\cache' = @{ Hit = $hitItem } } },
            @{ N = 'known';        Known = $known; Items = @{ 'C:\nx\known\sessiondata.vhdx' = $knownItem } },
            @{ N = 'nothing';      Known = $known; Items = @{ 'C:\nx\bundle' = $cacheItem }; Scan = @{ 'C:\nx\bundle' = @{} } })) {
        $script:ClaudeEnv = [pscustomobject]@{ VhdxFiles = $c.Known }
        $script:BkItems = $c.Items
        $script:BkScan = if ($c['Scan']) { $c.Scan } else { @{} }
        $VmCachePath = 'C:\nx\cache'
        $BundlePath  = 'C:\nx\bundle'
        try {
            $got = Get-CoworkVhdxSource -Name 'sessiondata.vhdx'
            $o[$c.N] = if ($null -eq $got) { 'null' } else { "item $($got.FullName)" }
        } catch {
            $o[$c.N] = "threw $($_.Exception.Message)"
        }
    }
    $script:ClaudeEnv = $null
    $o
}
Assert-Case "a known path that could not be read: throws" `
    ($found['knownUnread'] -eq "threw Access is denied") $found['knownUnread']
Assert-Case "a cache folder that could not be read: throws" `
    ($found['folderUnread'] -eq "threw Access is denied") $found['folderUnread']
Assert-Case "a search that could not read all of the cache and found nothing: throws" `
    ($found['scanUnread'] -eq "threw C:\nx\cache could not be searched in full (Access to the path 'C:\nx\cache\sub' is denied.)") $found['scanUnread']
Assert-Case "a search that found it: the file" ($found['hit'] -eq 'item C:\nx\cache\sub2\sessiondata.vhdx') $found['hit']
Assert-Case "a known path that is there: the file" ($found['known'] -eq 'item C:\nx\known\sessiondata.vhdx') $found['known']
Assert-Case "nowhere, and every place read: null" ($found['nothing'] -eq 'null') $found['nothing']

# Restore-CoworkVhdx, lifted and run against fakes of what it reads and
# writes. Since 6.3.1 a backup that could not be read is said and not
# restored, and a target that could not be read is said and not written
# over: it may be the disk itself, in a folder that cannot be listed and can
# still be written. Until 6.3.1 Test-Path answered $false for both: the
# backup was passed over without a word, and the target was overwritten.
# Every path is under C:\nx, which is never created; the copy and the
# folder it makes are fakes, and nothing runs unless every command the
# function calls is one of them or a pure one.
Write-Host ""
Write-Host "--- Restore-CoworkVhdx: a backup or a target that could not be read"
$restoreSrc = Get-FunctionSource -Path $fixPath -Name 'Restore-CoworkVhdx'
$restoreAst = [System.Management.Automation.Language.Parser]::ParseInput($restoreSrc, [ref]$null, [ref]$null)
$restoreSafe = @('Join-Path', 'Split-Path', 'Test-Path', 'Get-Item', 'New-Item', 'Copy-Item', 'Out-Null', 'Log',
    'Get-ClaudeItemOrNull', 'Get-ClaudeErrorText')
$restoreCalls = @($restoreAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) |
    ForEach-Object { $_.GetCommandName() } | Where-Object { $restoreSafe -notcontains $_ } | Sort-Object -Unique)
Assert-Case "every command Restore-CoworkVhdx calls is faked or pure" ($restoreCalls.Count -eq 0) ($restoreCalls -join ', ')
$rs = $null
if ($restoreCalls.Count -eq 0) {
    $rs = & {
        Set-StrictMode -Version Latest
        Invoke-Expression $restoreSrc
        function New-BkRecord {
            param([string]$TypeName, [string]$Message, [string]$Id, [string]$Category, $Target)
            $ex = Microsoft.PowerShell.Utility\New-Object $TypeName $Message
            return (Microsoft.PowerShell.Utility\New-Object System.Management.Automation.ErrorRecord($ex, $Id,
                    [System.Management.Automation.ErrorCategory]$Category, $Target))
        }
        # BkItems maps a path to its item, or to Denied; a path not listed
        # is not there. Get-Item writes the errors the cmdlet writes
        # (measured 2026-11-22 and 2026-11-23); Test-Path answers $false for
        # both, as the cmdlet does.
        function Get-Item {
            [CmdletBinding()]
            param([Parameter(Position = 0)][string]$Path, [string]$LiteralPath, [switch]$Force)
            $null = $Force
            $p = if ($LiteralPath) { $LiteralPath } else { $Path }
            if (-not $script:BkItems.ContainsKey($p)) {
                $PSCmdlet.WriteError((New-BkRecord 'System.Management.Automation.ItemNotFoundException' `
                    "Cannot find path '$p' because it does not exist." 'PathNotFound' 'ObjectNotFound' $p))
                return
            }
            $v = $script:BkItems[$p]
            if ($v -is [string] -and $v -eq 'Denied') {
                $PSCmdlet.WriteError((New-BkRecord 'System.UnauthorizedAccessException' 'Access is denied' `
                    'ItemExistsUnauthorizedAccessError' 'PermissionDenied' $p))
                return
            }
            return $v
        }
        function Test-Path {
            param([Parameter(Position = 0)][string]$Path, [string]$LiteralPath)
            $p = if ($LiteralPath) { $LiteralPath } else { $Path }
            $v = $script:BkItems[$p]
            return ($null -ne $v -and -not ($v -is [string] -and $v -eq 'Denied'))
        }
        function New-Item {
            [CmdletBinding()]
            param([Parameter(Position = 0)][string]$Path, [string]$ItemType, [switch]$Force)
            $null = $ItemType, $Force
            [void]$script:BkWrites.Add("MKDIR $Path")
        }
        function Copy-Item {
            [CmdletBinding()]
            param([Parameter(Position = 0)][string]$Path, [Parameter(Position = 1)][string]$Destination, [switch]$Force)
            $null = $Force
            [void]$script:BkWrites.Add("COPY $Path -> $Destination")
        }
        $script:BkWrites = [System.Collections.ArrayList]::new()
        $script:LogLines = @()
        $folder = [pscustomobject]@{ FullName = 'C:\nx\cache' }
        $file = [pscustomobject]@{ FullName = 'a file' }
        # a: restored. b: its target could not be read. c: its backup could
        # not be read. d: its target is there. e: no backup.
        $script:BkItems = @{
            'C:\nx\cache' = $folder
            'C:\nx\bk\a.vhdx' = $file
            'C:\nx\bk\b.vhdx' = $file; 'C:\nx\cache\b.vhdx' = 'Denied'
            'C:\nx\bk\c.vhdx' = 'Denied'
            'C:\nx\bk\d.vhdx' = $file; 'C:\nx\cache\d.vhdx' = $file
        }
        $backedUp = [ordered]@{}
        foreach ($n in 'a', 'b', 'c', 'd', 'e') { $backedUp["$n.vhdx"] = "C:\nx\cache\$n.vhdx" }
        $o = [ordered]@{}
        try {
            $o['count'] = Restore-CoworkVhdx -BackupDir 'C:\nx\bk' -BackedUp $backedUp
        } catch { $o['count'] = "threw $($_.Exception.Message)" }
        $o['writes'] = @($script:BkWrites) -join ' | '
        $o['said'] = @($script:LogLines) -join ' | '
        $script:BkItems = $null
        $o
    }
}
if ($rs) {
    Assert-Case "one backup restored, and only that one" ($rs['count'] -eq 1 -and $rs['writes'] -eq 'COPY C:\nx\bk\a.vhdx -> C:\nx\cache\a.vhdx') `
        "count $($rs['count']); writes: $($rs['writes'])"
    Assert-Case "a target that could not be read is not written over, and that is said" `
        ($rs['writes'] -notmatch 'b\.vhdx' -and
         @($rs['said'] -split ' \| ') -contains ("[!] Whether C:\nx\cache\b.vhdx is there could not be read (Access is denied), " +
            "so the backup of b.vhdx was not restored over it")) $rs['said']
    Assert-Case "a backup that could not be read is not restored, and that is said" `
        ($rs['writes'] -notmatch 'c\.vhdx' -and
         @($rs['said'] -split ' \| ') -contains ("[!] The backup of c.vhdx at C:\nx\bk\c.vhdx could not be read (Access is denied), " +
            "so it was not restored")) $rs['said']
    Assert-Case "a target that is there and a backup that is not: passed over, nothing said" `
        ($rs['said'] -notmatch '[de]\.vhdx' -and $rs['writes'] -notmatch '[de]\.vhdx') $rs['said']
} else {
    Assert-Case "Restore-CoworkVhdx ran" $false
}

# Clean-up of this script's own sandbox: every case is counted by now,
# and a folder left in %TEMP% changes nothing.
Remove-Item $sandbox -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host ("Backup logic: {0} passed, {1} failed" -f $pass, $fail)
if ($fail -eq 0) { Write-Host "Backup logic gate PASSED"; exit 0 }
Write-Host "Backup logic gate FAILED"
exit 1
