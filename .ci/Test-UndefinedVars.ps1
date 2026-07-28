<#
.SYNOPSIS
    Flags variables that are read but never assigned anywhere in a script.
.DESCRIPTION
    The parser gate cannot catch this. PowerShell resolves variables at run
    time, so a script that reads a variable nobody ever assigns parses
    perfectly and then dies under Set-StrictMode the moment that line executes.

    That is not hypothetical here: rewiring the constants block left four
    references to a removed $ClaudeLogDirAD and two to a renamed
    $ErrorPatterns. Both files parsed clean. Only a grep found them.

    This walks the AST, collects every name that is assigned, declared as a
    parameter, or bound by a foreach, and reports every read that matches none
    of them. Scope prefixes are normalised, so $script:Foo and $Foo count as
    the same name.

    Exits 0 when clean, 1 otherwise.
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

# PowerShell automatic and preference variables, plus provider-qualified names.
$known = @(
    '_', 'psitem', 'args', 'input', 'this', 'true', 'false', 'null',
    'error', 'host', 'pscommandpath', 'myinvocation', 'pid', 'psversiontable',
    'lastexitcode', 'pscmdlet', 'psboundparameters', 'pwd', 'home',
    'psscriptroot', 'matches', 'stacktrace', 'executioncontext', 'shellid',
    'pshome', 'psculture', 'psuiculture', 'psedition', 'profile',
    'iswindows', 'islinux', 'ismacos', 'nestedpromptlevel', 'outputencoding',
    'whatifpreference', 'confirmpreference', 'debugpreference',
    'erroractionpreference', 'verbosepreference', 'warningpreference',
    'informationpreference', 'progresspreference', 'psdefaultparametervalues',
    'psnativecommandargumentpassing', 'erroractionpreference', 'foreach',
    'switch', 'psculture', 'consolefilename', 'psemailserver'
)

function Get-BaseName {
    param([string]$Name)
    $n = $Name
    foreach ($prefix in @('script:', 'global:', 'local:', 'private:', 'using:')) {
        if ($n.ToLower().StartsWith($prefix)) { $n = $n.Substring($prefix.Length) }
    }
    return $n.ToLower()
}

function Test-ScriptVariables {
    param([string]$Path)

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    if (@($errors).Count -gt 0) {
        Write-Host ("PARSE-FAIL : {0}" -f (Split-Path $Path -Leaf))
        return 1
    }

    $assigned = New-Object 'System.Collections.Generic.HashSet[string]'

    # Assignment targets, including the left side of "if ($x = ...)".
    foreach ($a in @($ast.FindAll({
        param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true))) {
        foreach ($v in @($a.Left.FindAll({
            param($n) $n -is [System.Management.Automation.Language.VariableExpressionAst] }, $true))) {
            $null = $assigned.Add((Get-BaseName $v.VariablePath.UserPath))
        }
    }

    # Parameters, both param() blocks and function parameters.
    foreach ($p in @($ast.FindAll({
        param($n) $n -is [System.Management.Automation.Language.ParameterAst] }, $true))) {
        $null = $assigned.Add((Get-BaseName $p.Name.VariablePath.UserPath))
    }

    # foreach ($x in ...) bindings.
    foreach ($f in @($ast.FindAll({
        param($n) $n -is [System.Management.Automation.Language.ForEachStatementAst] }, $true))) {
        $null = $assigned.Add((Get-BaseName $f.Variable.VariablePath.UserPath))
    }

    # trap / catch bindings and data statements.
    foreach ($d in @($ast.FindAll({
        param($n) $n -is [System.Management.Automation.Language.DataStatementAst] }, $true))) {
        if ($d.Variable) { $null = $assigned.Add((Get-BaseName $d.Variable)) }
    }

    $reported = New-Object 'System.Collections.Generic.HashSet[string]'
    $findings = @()

    foreach ($v in @($ast.FindAll({
        param($n) $n -is [System.Management.Automation.Language.VariableExpressionAst] }, $true))) {

        $userPath = $v.VariablePath.UserPath
        # Provider-qualified names such as $env:APPDATA are not script variables.
        if ($v.VariablePath.IsDriveQualified) { continue }

        $base = Get-BaseName $userPath
        if ($known -contains $base) { continue }
        if ($assigned.Contains($base)) { continue }
        if ($reported.Contains($base)) { continue }

        $null = $reported.Add($base)
        $findings += [pscustomobject]@{
            Name = $userPath
            Line = $v.Extent.StartLineNumber
        }
    }

    $name = Split-Path $Path -Leaf
    if ($findings.Count -eq 0) {
        Write-Host ("OK      : {0}" -f $name)
        return 0
    }
    Write-Host ("FAIL    : {0} ({1} name(s) read but never assigned)" -f $name, $findings.Count)
    foreach ($f in $findings) {
        Write-Host ("          line {0,-5} `${1}" -f $f.Line, $f.Name)
    }
    return 1
}

$failed = 0
foreach ($n in @('Fix-ClaudeDesktop.ps1', 'Watch-ClaudeHealth.ps1', 'Prevent-ClaudeIssues.ps1')) {
    $p = Join-Path $Root $n
    if (-not (Test-Path $p)) {
        Write-Host ("MISSING : {0}" -f $n)
        $failed++
        continue
    }
    $failed += (Test-ScriptVariables -Path $p)
}

Write-Host ""
if ($failed -eq 0) { Write-Host "Undefined-variable gate PASSED"; exit 0 }
Write-Host ("Undefined-variable gate FAILED ({0} file(s))" -f $failed)
exit 1
