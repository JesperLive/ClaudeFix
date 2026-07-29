# Test-AlwaysTriesStart.ps1
#
# Restart-CoworkService used to consult Test-CoworkServicePrereq first and
# return $false on a Blocked result without ever calling Start-Service. On the
# 2026-07-29 reporter's machine that check wrongly flagged Virtual Machine
# Platform, so the script printed "Service failed to start" for a start it had
# never attempted, then repeated it every five seconds.
#
# Their own cowork-service.log proved it: between 11:12:16, when the script
# stopped the service, and 11:29:18, when Claude Desktop started it again,
# there is not a single line. No failed start, because no start.
#
# The rule this gate enforces: a check that can be wrong must never prevent the
# attempt that would prove it wrong. Prerequisites are advisory. The service is
# the authority on whether the service can run.
#
# This drives the real function out of the shipped file with every dependency
# stubbed, with the prereq check hardwired to report a blocker, and asserts the
# start was attempted regardless.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root    = Split-Path -Parent $PSScriptRoot
$fixPath = Join-Path $root 'Fix-ClaudeDesktop.ps1'

$script:pass = 0
$script:fail = 0
function Assert {
    param([string]$Label, [bool]$Condition, [string]$Detail = '')
    if ($Condition) { Write-Host "  PASS  $Label" -ForegroundColor Green; $script:pass++ }
    else { Write-Host "  FAIL  $Label  $Detail" -ForegroundColor Red; $script:fail++ }
}

function Get-FunctionSource {
    param([string]$Path, [string]$Name)
    $tokens = $null; $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    if (@($errors).Count -gt 0) { throw "parse errors in $Path" }
    $fn = $ast.Find({
        param($n)
        $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name
    }, $true)
    if (-not $fn) { throw "function '$Name' not found in $Path" }
    return $fn.Extent.Text
}

# -- Harness -------------------------------------------------------
$ServiceName   = 'CoworkVMService'
$ServiceExe    = 'cowork-svc'
$StartPollMax  = 4
$script:IsAdmin = $true

$script:StartCalls   = 0
$script:ServiceState = 'Stopped'
$script:PrereqBlocked = $true

function Log { param([string]$Message, [string]$Colour = 'White', [switch]$Indent) }

function Test-CoworkServicePrereq {
    return @{
        Found = $true; StartType = 'Auto'; Repaired = $false
        Blocked = $script:PrereqBlocked
        Reason  = 'The Virtual Machine Platform feature is Disabled.'
        Fixes   = @('Enable-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform -All')
    }
}

# [CmdletBinding()] already supplies -ErrorAction. Declaring it again as a
# parameter throws "A parameter with the name 'ErrorAction' was defined multiple
# times", which is a harness fault and says nothing about the code under test.
function Get-Service {
    [CmdletBinding()]
    param([string]$Name)
    if ($script:ServiceState -eq 'Absent') { return $null }
    return [pscustomobject]@{ Name = $Name; Status = $script:ServiceState }
}

function Start-Service {
    [CmdletBinding()]
    param([string]$Name)
    $script:StartCalls++
    # Model the reporter's machine: the service starts fine when asked.
    $script:ServiceState = 'Running'
}

function Stop-Service { [CmdletBinding()] param([string]$Name, [switch]$Force) }
function Stop-Process { [CmdletBinding()] param([string]$Name, [switch]$Force) }
function Start-Process { [CmdletBinding()] param($FilePath, $ArgumentList, [switch]$NoNewWindow, [switch]$Wait) }
function Start-Sleep { [CmdletBinding()] param([int]$Seconds, [int]$Milliseconds) }

Invoke-Expression (Get-FunctionSource -Path $fixPath -Name 'Restart-CoworkService')

# ------------------------------------------------------------------
Write-Host ""
Write-Host "--- prereq reports Blocked, service actually works" -ForegroundColor Cyan
# This is the reporter's machine exactly: a wrong prerequisite verdict over a
# service that starts in under a second.
$script:PrereqBlocked = $true
$script:ServiceState  = 'Stopped'
$script:StartCalls    = 0

$result = Restart-CoworkService

Assert "a start was attempted despite the blocker" ($script:StartCalls -ge 1) "StartCalls=$($script:StartCalls)"
Assert "the restart reports success"                ($result -eq $true)      "got $result"
Assert "the service ended up Running"               ($script:ServiceState -eq 'Running')

# ------------------------------------------------------------------
Write-Host ""
Write-Host "--- prereq clean, service works" -ForegroundColor Cyan
$script:PrereqBlocked = $false
$script:ServiceState  = 'Stopped'
$script:StartCalls    = 0

$result = Restart-CoworkService
Assert "a start was attempted"      ($script:StartCalls -ge 1) "StartCalls=$($script:StartCalls)"
Assert "the restart reports success" ($result -eq $true)       "got $result"

# ------------------------------------------------------------------
Write-Host ""
Write-Host "--- service genuinely will not start" -ForegroundColor Cyan
# Redefine Start-Service to do nothing, so the poll times out. The point is that
# failure is still discovered by trying, not predicted.
function Start-Service {
    [CmdletBinding()]
    param([string]$Name)
    $script:StartCalls++
}
$script:PrereqBlocked = $true
$script:ServiceState  = 'Stopped'
$script:StartCalls    = 0

$result = Restart-CoworkService
Assert "a start was still attempted"   ($script:StartCalls -ge 1) "StartCalls=$($script:StartCalls)"
Assert "the restart reports failure"   ($result -eq $false)       "got $result"

# ------------------------------------------------------------------
Write-Host ""
Write-Host "--- the source carries no early return on Blocked" -ForegroundColor Cyan
# Belt and braces against the shape coming back somewhere the harness above
# happens not to reach.
#
# A plain text scan between ".Blocked" and "Start-Service" is too crude: it
# fires on the legitimate early returns for "service not registered" and "could
# not stop it", which are conclusions drawn from the machine rather than
# predictions about it. The thing to forbid is narrower, so find the actual if
# statement that tests .Blocked and assert THAT branch does not bail.
$src = Get-FunctionSource -Path $fixPath -Name 'Restart-CoworkService'
$tok = $null; $err = $null
$fnAst = [System.Management.Automation.Language.Parser]::ParseInput($src, [ref]$tok, [ref]$err)

$blockedIfs = @($fnAst.FindAll({
    param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and
    $n.Clauses[0].Item1.Extent.Text -match '\.Blocked'
}, $true))

Assert "the function still consults .Blocked" ($blockedIfs.Count -ge 1) "found $($blockedIfs.Count)"

$vetoFound = $false
foreach ($ifAst in $blockedIfs) {
    # Only the branch guarding the START matters. The failure-reporting branch
    # near the end legitimately sits after a start has already been attempted,
    # so restrict this to if statements that appear before the first
    # Start-Service in the function body.
    $startIdx = $src.IndexOf('Start-Service')
    $ifIdx    = $ifAst.Extent.StartOffset - $fnAst.Extent.StartOffset
    if ($startIdx -ge 0 -and $ifIdx -lt $startIdx) {
        if ($ifAst.Clauses[0].Item2.Extent.Text -match 'return\s+\$false') { $vetoFound = $true }
    }
}
Assert "no .Blocked branch before the start returns `$false" (-not $vetoFound)

Write-Host ""
Write-Host "Always-tries-start: $script:pass passed, $script:fail failed"
if ($script:fail -gt 0) {
    Write-Host "Always-tries-start gate FAILED" -ForegroundColor Red
    exit 1
}
Write-Host "Always-tries-start gate PASSED" -ForegroundColor Green
exit 0
