# Test-GeneratedLaunchers.ps1
#
# Prevent-ClaudeIssues writes PowerShell out of here-strings: the boot task, the
# admin launcher, and now Launch-Claude-NoGPU.ps1. None of that is ever parsed
# at build time, so a stray brace or a backtick that was meant to escape and
# did not would ship silently and only surface when a user double-clicks it.
#
# The NoGPU launcher is a here-string inside a here-string, near enough: it
# interpolates $ToolkitVersion and the discovered package identity while
# escaping the runtime variables with backticks. That is exactly the shape that
# gets it wrong. So build it here the way the script builds it, and parse it.
#
# Also asserts the identity is interpolated rather than hardcoded. Writing
# Claude_pzs8sxrjxfjjc into the source would work on the machine this was
# written on and nowhere else, which is the bug the discovery layer exists to
# prevent and which has been reintroduced by hand once already.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root    = Split-Path -Parent $PSScriptRoot
$srcPath = Join-Path $root 'Prevent-ClaudeIssues.ps1'

$script:pass = 0
$script:fail = 0
function Assert {
    param([string]$Label, [bool]$Condition, [string]$Detail = '')
    if ($Condition) { Write-Host "  PASS  $Label" -ForegroundColor Green; $script:pass++ }
    else { Write-Host "  FAIL  $Label  $Detail" -ForegroundColor Red; $script:fail++ }
}

function Test-Parses {
    param([string]$Label, [string]$Code)
    $errors = $null
    $null = [System.Management.Automation.Language.Parser]::ParseInput($Code, [ref]$null, [ref]$errors)
    $n = @($errors).Count
    Assert $Label ($n -eq 0) $(if ($n) { $errors[0].Message } else { '' })
}

# ------------------------------------------------------------------
# Rebuild the generated launcher from Prevent's own source, for both install
# shapes, and parse the result. The parser lifts the launch-body choice and the
# launcher here-string out of the source and they are evaluated here, so what
# is tested is what Prevent writes. Until 6.2.0 this file held a hand copy of
# the here-string, and the copy had already drifted from the source.
# ------------------------------------------------------------------
$src    = [System.IO.File]::ReadAllText($srcPath)
$srcAst = [System.Management.Automation.Language.Parser]::ParseInput($src, [ref]$null, [ref]$null)

$bodyIf = $srcAst.Find({ param($n)
    $n -is [System.Management.Automation.Language.IfStatementAst] -and
    $n.Clauses[0].Item1.Extent.Text -eq '$isMsix -and $pfn -and $appId' }, $true)
$ps1Assign = $srcAst.Find({ param($n)
    $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and
    $n.Left.Extent.Text -eq '$ps1Content' }, $true)
$template = $null
if ($ps1Assign -and $ps1Assign.Right -is [System.Management.Automation.Language.CommandExpressionAst] -and
    $ps1Assign.Right.Expression -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) {
    $template = $ps1Assign.Right.Expression
}
Assert "the launch-body choice was found in Prevent"  ($null -ne $bodyIf)
Assert "the launcher here-string was found in Prevent" ($null -ne $template)

if ($bodyIf -and $template) {
    # Evaluating source text is only safe while it runs nothing. These two hold
    # it to that, and the second is also the escaping check: a runtime variable
    # that lost its backtick would expand when the launcher is written, and
    # show up here as a fourth name.
    $cmdsInIf = @($bodyIf.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true))
    Assert "the launch-body choice runs no command" ($cmdsInIf.Count -eq 0) `
        (($cmdsInIf | ForEach-Object { $_.Extent.Text }) -join ' | ')
    $expanded = @($template.NestedExpressions | ForEach-Object { $_.Extent.Text } | Sort-Object -Unique)
    $expandsOnlyThree = (($expanded -join ',') -eq '$howNote,$launchBody,$ToolkitVersion')
    Assert "writing the launcher expands only `$howNote, `$launchBody and `$ToolkitVersion" `
        $expandsOnlyThree "expands $($expanded -join ', ')"

    $ToolkitVersion = '0.0.0-test'
    $cases = @()
    if ($cmdsInIf.Count -eq 0 -and $expandsOnlyThree) { $cases = @(
        @{ Name = 'MSIX install'; IsMsix = $true; Pfn = 'Claude_testfamily1'; AppId = 'ClaudeTest'; Exe = $null
           Launch = "-PackageFamilyName 'Claude_testfamily1'" }
        @{ Name = 'Traditional install'; IsMsix = $false; Pfn = $null; AppId = $null
           Exe = 'C:\Program Files\Claude\Claude.exe'
           Launch = "Start-Process -FilePath 'C:\Program Files\Claude\Claude.exe' -ArgumentList '--disable-gpu'" }
    ) }
    foreach ($case in $cases) {
        Write-Host ""
        Write-Host "--- generated launcher: $($case.Name)" -ForegroundColor Cyan

        $isMsix  = $case.IsMsix
        $pfn     = $case.Pfn
        $appId   = $case.AppId
        $exePath = $case.Exe
        $launchBody = $null
        $howNote    = $null
        . ([scriptblock]::Create($bodyIf.Extent.Text))
        $ps1Content = & ([scriptblock]::Create($template.Extent.Text))

        Test-Parses "the generated script parses" $ps1Content

        # The runtime variables must survive as literal $ in the output, not be
        # expanded away at generation time. If the backticks were dropped,
        # $running would interpolate to nothing and the guard would be dead code.
        Assert "runtime `$running survived as a variable" ($ps1Content -match '\$running\.Count')
        Assert "runtime `$ErrorActionPreference survived"  ($ps1Content -match '\$ErrorActionPreference')
        Assert "the version WAS interpolated"             ($ps1Content -match [regex]::Escape($ToolkitVersion))
        Assert "the launch for this install was inlined"  ($ps1Content -match [regex]::Escape($case.Launch))

        # Only the app is closed. Every process named claude used to be, which
        # took the Claude Code CLI down with it: project threads, Code sessions,
        # terminal sessions.
        Assert "the close goes through the app filter"    ($ps1Content -match '\$running = @\(Get-ClaudeAppProcess\)')

        # Run the generated filter against fake processes. Get-Process and
        # Get-Item are functions inside this scope, which PowerShell resolves
        # before the cmdlets, so nothing real is listed or touched.
        $genAst = [System.Management.Automation.Language.Parser]::ParseInput($ps1Content, [ref]$null, [ref]$null)
        $filter = $genAst.Find({ param($n)
            $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $n.Name -eq 'Get-ClaudeAppProcess' }, $true)
        Assert "the generated launcher defines Get-ClaudeAppProcess" ($null -ne $filter)
        if ($filter) {
            $picked = & {
                function Get-Process {
                    param([string]$Name, [string]$ErrorAction)
                    $null = $Name, $ErrorAction
                    $made = @()
                    foreach ($p in @(@(1, 'C:\Program Files\WindowsApps\Claude_1.0.0.0_x64__x\app\Claude.exe'),
                                     @(2, 'C:\Users\u\AppData\Local\Claude-Data\claude_x\claude-code\2.1.327\claude.exe'),
                                     @(3, 'C:\Users\u\.local\bin\claude.exe'),
                                     @(4, $null))) {
                        $o = [pscustomobject]@{ Id = $p[0] }
                        if ($p[1]) {
                            $o | Add-Member -MemberType NoteProperty -Name MainModule -Value ([pscustomobject]@{ FileName = $p[1] })
                        } else {
                            $o | Add-Member -MemberType ScriptProperty -Name MainModule -Value { throw 'Access is denied' }
                        }
                        $made += $o
                    }
                    return $made
                }
                function Get-Item {
                    param([string]$LiteralPath)
                    $product = 'Claude'
                    if ($LiteralPath -like '*\.local\bin\*') { $product = 'Claude Code' }
                    return [pscustomobject]@{ VersionInfo = [pscustomobject]@{ ProductName = $product } }
                }
                . ([scriptblock]::Create($filter.Extent.Text))
                @(Get-ClaudeAppProcess | ForEach-Object { $_.Id })
            }
            Assert "the generated filter picks the app and only the app" `
                ((@($picked) -join ',') -eq '1') "picked $(@($picked) -join ',')"
        }
    }
}

# ------------------------------------------------------------------
Write-Host ""
Write-Host "--- source does not hardcode this machine's install" -ForegroundColor Cyan

# The literal package family of the machine this was developed on. Allowed in
# comments, never in code.
$codeLines = @($src -split "`r?`n" | Where-Object { $_ -notmatch '^\s*#' })
$hardcoded = @($codeLines | Where-Object { $_ -match 'Claude_pzs8sxrjxfjjc' })
Assert "no hardcoded package family in code" ($hardcoded.Count -eq 0) ($hardcoded -join ' | ')

# The admin launcher payload is a single-quoted here-string, so the source text
# IS the generated script. Parse it as it stands.
$adminPayload = [regex]::Match($src, "(?s)\`$launchCmd = @'\r?\n(.*?)\r?\n'@")
Assert "the admin launcher payload was found" $adminPayload.Success
if ($adminPayload.Success) {
    Test-Parses "the admin launcher payload parses" $adminPayload.Groups[1].Value
}

# The admin launcher is not evaluated above, so its process checks are read
# off the source. The NoGPU pair stays as a second, plainer witness.
Assert "the NoGPU launcher closes through the app filter" `
    ($src -match '`\$running = @\(Get-ClaudeAppProcess\)')
Assert "no generated launcher closes every process named claude" `
    ($src -notmatch '`\$running = @\(Get-Process claude')
Assert "the admin launcher's already-running check skips the CLI" `
    ($src -match "(?s)\`$existing = @\(Get-Process -Name claude.{0,400}claude-code")
Assert "the admin launcher no longer reads its exe off a running process" `
    ($src -notmatch '\$proc = Get-Process -Name Claude')

Assert "the launcher reads PackageFamilyName from discovery" `
    ($src -match '\$script:ClaudeEnv\.PackageFamilyName')
Assert "the launcher reads ApplicationId from discovery" `
    ($src -match '\$script:ClaudeEnv\.ApplicationId')

# Undo has to take back everything setup puts down, or a user who reverts is
# left with a desktop shortcut pointing at a script nothing maintains.
Assert "Undo removes Launch-Claude-NoGPU.ps1"  ($src -match 'Launch-Claude-NoGPU\.ps1')
Assert "Undo removes Launch-Claude-NoGPU.cmd"  ($src -match 'Launch-Claude-NoGPU\.cmd')
Assert "Undo removes the desktop shortcut"     ($src -match 'Claude \(no GPU\)\.lnk')

# Step numbering has to agree with the declared total, or the progress display
# reads "Step 27 of 26".
$declared = [regex]::Match($src, '(?m)^\s*\$steps\s*=\s*(\d+)\s*$\s*[\s\S]{0,4000}?Step 1 \$steps "Configuring power plan')
$maxStep  = 0
foreach ($m in [regex]::Matches($src, '(?m)^\s*Step (\d+) \$steps')) {
    $v = [int]$m.Groups[1].Value
    if ($v -gt $maxStep) { $maxStep = $v }
}
$totals = @([regex]::Matches($src, '(?m)^\s*\$steps = (\d+)') | ForEach-Object { [int]$_.Groups[1].Value })
Assert "highest step number matches a declared total" `
    ($totals -contains $maxStep) "max step $maxStep, declared totals $($totals -join ', ')"

Write-Host ""
Write-Host "Generated launchers: $script:pass passed, $script:fail failed"
if ($script:fail -gt 0) {
    Write-Host "Generated launcher gate FAILED" -ForegroundColor Red
    exit 1
}
Write-Host "Generated launcher gate PASSED" -ForegroundColor Green
exit 0
