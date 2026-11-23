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

    Only "=" to the variable itself assigns it: "$x = ", "[int]$x = ",
    "$a, $b = " and "($x) = ". "$x[$k] = " and "$x.Name = " need $x to be
    set already, and under Set-StrictMode die on a $x never set. "$x += "
    dies the same way at script scope and on $script:x; in a function it
    sets a local $x, but a read of $x before it, or after it when it did
    not run, still dies, so it does not count either. Until 6.3.1 every
    variable on the left of any assignment counted, $k too, so deleting
    the line that sets up a table passed this gate whenever an element of
    the table was assigned somewhere, as Watch's $script:StorageWarningsSaid
    is.

    A read of $script:x also needs x set at script scope: by "=" to
    $script:x or $global:x anywhere, or by "=", a foreach or the script's
    own param() in the script's own scope. "=" to $private:x there
    counts only for a read of $script:x there too: no other scope sees a
    private name, so a function's $script:x dies on it, and a read in the
    script's own scope, or in a block that runs in place, does not
    (measured). A name set only in a scope of its own is that scope's
    local, and $script:x dies on it. A scope of its own is a function, a
    trap, or a script block that does not run in the scope around it.
    Measured on Windows PowerShell 5.1, these run in the scope around them:
    a script block given to ForEach-Object (as -Begin, -Process or -End
    too), Where-Object, their aliases %, foreach, ? and where,
    Measure-Command, Invoke-Command or icm with -NoNewScope (any prefix of
    it, as a plain switch or ":$true"), and one dot-sourced where it is
    written, ". { }". These do not: "& { }", Invoke-Command without
    -NoNewScope, Sort-Object, a calculated property of Select-Object, a
    trap, and the .ForEach() method of a List[T]. The .ForEach() and
    .Where() methods of an array do run in place, but which method a call
    reaches is not known from the script, so a script block given to any
    method counts as a scope of its own. Any other script block counts as a
    scope of its own, so one stored and dot-sourced later must not be the
    only place a name read as $script:x is set. Until 6.3.1 an assignment
    anywhere counted for $script:x.

    Not counted, so a variable they set must be set with "=" first:
    -ErrorVariable, -OutVariable, -PipelineVariable, Set-Variable and
    New-Variable. A name read without a scope is not checked per scope, nor
    per path through the code: PowerShell's dynamic scoping lets a function
    read its caller's variables, so a read in one function may be set by
    another.

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

function Get-AssignedVariable {
    # The variables the left side of an "=" sets: the variable itself, also
    # through a type constraint or an attribute, in parentheses, or each of
    # a list. An element or a member sets none.
    param($Target)
    if ($Target -is [System.Management.Automation.Language.VariableExpressionAst]) { $Target; return }
    if ($Target -is [System.Management.Automation.Language.AttributedExpressionAst]) {
        Get-AssignedVariable -Target $Target.Child
        return
    }
    if ($Target -is [System.Management.Automation.Language.ParenExpressionAst]) {
        $pipe = $Target.Pipeline
        if ($pipe -is [System.Management.Automation.Language.PipelineAst] -and @($pipe.PipelineElements).Count -eq 1 -and
            $pipe.PipelineElements[0] -is [System.Management.Automation.Language.CommandExpressionAst]) {
            Get-AssignedVariable -Target $pipe.PipelineElements[0].Expression
        }
        return
    }
    if ($Target -is [System.Management.Automation.Language.ArrayLiteralAst]) {
        foreach ($e in $Target.Elements) { Get-AssignedVariable -Target $e }
    }
}

function Test-RunsInPlace {
    # Whether a script block written here runs in the scope around it (see
    # the description).
    param($Block)
    $parent = $Block.Parent
    if ($parent -is [System.Management.Automation.Language.CommandParameterAst]) { $parent = $parent.Parent }
    if ($parent -is [System.Management.Automation.Language.CommandAst]) {
        if ($parent.InvocationOperator -eq [System.Management.Automation.Language.TokenKind]::Dot -and
            $parent.CommandElements[0] -eq $Block) { return $true }
        $name = $parent.GetCommandName()
        if ($name -in @('ForEach-Object', '%', 'foreach', 'Where-Object', '?', 'where', 'Measure-Command')) { return $true }
        if ($name -in @('Invoke-Command', 'icm')) { return (Test-NoNewScope -Command $parent) }
    }
    return $false
}

function Test-NoNewScope {
    # Whether an Invoke-Command is given -NoNewScope: by any prefix of the
    # name ("-N" is enough, measured), as a plain switch or ":$true".
    param($Command)
    foreach ($e in $Command.CommandElements) {
        if ($e -is [System.Management.Automation.Language.CommandParameterAst] -and
            'NoNewScope'.StartsWith($e.ParameterName, [System.StringComparison]::OrdinalIgnoreCase)) {
            if ($null -eq $e.Argument) { return $true }
            return ($e.Argument -is [System.Management.Automation.Language.VariableExpressionAst] -and
                    $e.Argument.VariablePath.UserPath -eq 'true')
        }
    }
    return $false
}

function Test-InChildScope {
    # Whether a node is in a scope of its own, whose plain names are its
    # locals: a function, a trap, or a script block that does not run in
    # the scope around it.
    param($Node)
    for ($p = $Node.Parent; $p; $p = $p.Parent) {
        if ($p -is [System.Management.Automation.Language.FunctionDefinitionAst]) { return $true }
        if ($p -is [System.Management.Automation.Language.TrapStatementAst]) { return $true }
        if ($p -is [System.Management.Automation.Language.ScriptBlockExpressionAst] -and
            -not (Test-RunsInPlace -Block $p)) { return $true }
    }
    return $false
}

function Test-ScriptScopeName {
    # Whether an assignment to this variable, at this node, sets it at
    # script scope: a $script: or $global: name anywhere, or a plain or
    # $local: name in the script's own scope. Not a $private: name, which
    # no function sees.
    param($Variable)
    $path = $Variable.VariablePath
    if ($path.IsScript -or $path.IsGlobal) { return $true }
    if ($path.IsLocal -or $path.IsUnqualified) { return -not (Test-InChildScope -Node $Variable) }
    return $false
}

function Test-OwnPrivateName {
    # Whether this is a $private: name in the script's own scope, which a
    # read of $script:x sees in that scope only.
    param($Variable)
    return ($Variable.VariablePath.IsPrivate -and -not (Test-InChildScope -Node $Variable))
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
    # The names of those set at script scope, for reads of $script:x.
    $scriptSet = New-Object 'System.Collections.Generic.HashSet[string]'
    # The names set as $private:x in the script's own scope, for reads of
    # $script:x there.
    $privateSet = New-Object 'System.Collections.Generic.HashSet[string]'

    # Assignment targets, including the left side of "if ($x = ...)": "="
    # to the variable itself only (see the description).
    foreach ($a in @($ast.FindAll({
        param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true))) {
        if ($a.Operator -ne [System.Management.Automation.Language.TokenKind]::Equals) { continue }
        foreach ($v in @(Get-AssignedVariable -Target $a.Left)) {
            $null = $assigned.Add((Get-BaseName $v.VariablePath.UserPath))
            if (Test-ScriptScopeName -Variable $v) { $null = $scriptSet.Add((Get-BaseName $v.VariablePath.UserPath)) }
            if (Test-OwnPrivateName -Variable $v) { $null = $privateSet.Add((Get-BaseName $v.VariablePath.UserPath)) }
        }
    }

    # Parameters, both param() blocks and function parameters. Only the
    # script's own param() sets a name at script scope.
    foreach ($p in @($ast.FindAll({
        param($n) $n -is [System.Management.Automation.Language.ParameterAst] }, $true))) {
        $null = $assigned.Add((Get-BaseName $p.Name.VariablePath.UserPath))
        if ($p.Parent -eq $ast.ParamBlock) { $null = $scriptSet.Add((Get-BaseName $p.Name.VariablePath.UserPath)) }
    }

    # foreach ($x in ...) bindings.
    foreach ($f in @($ast.FindAll({
        param($n) $n -is [System.Management.Automation.Language.ForEachStatementAst] }, $true))) {
        $null = $assigned.Add((Get-BaseName $f.Variable.VariablePath.UserPath))
        if (Test-ScriptScopeName -Variable $f.Variable) { $null = $scriptSet.Add((Get-BaseName $f.Variable.VariablePath.UserPath)) }
        if (Test-OwnPrivateName -Variable $f.Variable) { $null = $privateSet.Add((Get-BaseName $f.Variable.VariablePath.UserPath)) }
    }

    # Data statements.
    foreach ($d in @($ast.FindAll({
        param($n) $n -is [System.Management.Automation.Language.DataStatementAst] }, $true))) {
        if ($d.Variable) {
            $null = $assigned.Add((Get-BaseName $d.Variable))
            if (-not (Test-InChildScope -Node $d)) { $null = $scriptSet.Add((Get-BaseName $d.Variable)) }
        }
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
        if ($assigned.Contains($base)) {
            # Set somewhere; a read of $script:x needs it set at script scope.
            if (-not $v.VariablePath.IsScript -or $scriptSet.Contains($base)) { continue }
            if ($privateSet.Contains($base) -and -not (Test-InChildScope -Node $v)) { continue }
            $kind = 'local'
        } else {
            $kind = 'never'
        }
        if ($reported.Contains($base)) { continue }

        $null = $reported.Add($base)
        $findings += [pscustomobject]@{
            Name = $userPath
            Line = $v.Extent.StartLineNumber
            Kind = $kind
        }
    }

    $name = Split-Path $Path -Leaf
    if ($findings.Count -eq 0) {
        Write-Host ("OK      : {0}" -f $name)
        return 0
    }
    foreach ($k in @(
            @{ Kind = 'never'; Text = 'read but never assigned' },
            @{ Kind = 'local'; Text = 'read as $script: but never set at script scope' })) {
        $these = @($findings | Where-Object { $_.Kind -eq $k.Kind })
        if ($these.Count -eq 0) { continue }
        Write-Host ("FAIL    : {0} ({1} name(s) {2}: {3})" -f $name, $these.Count, $k.Text,
            (@($these | ForEach-Object { '$' + $_.Name }) -join ', '))
        foreach ($f in $these) {
            Write-Host ("          line {0,-5} `${1}" -f $f.Line, $f.Name)
        }
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
