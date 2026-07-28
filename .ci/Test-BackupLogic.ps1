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

function Get-CoworkVhdxSource {
    param([Parameter(Mandatory)][string]$Name)
    if ($script:FakeSources.ContainsKey($Name)) { return $script:FakeSources[$Name] }
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

Remove-Item $sandbox -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host ("Backup logic: {0} passed, {1} failed" -f $pass, $fail)
if ($fail -eq 0) { Write-Host "Backup logic gate PASSED"; exit 0 }
Write-Host "Backup logic gate FAILED"
exit 1
