<#
.SYNOPSIS
    PCCM test-only audit: variables a PowerShell runner can READ BEFORE ASSIGNMENT.

.DESCRIPTION
    WHY THIS EXISTS. Two Windows turns were spent on startup defects, and the
    second was `$script:P8ZPath` - read by a helper copied out of an accepted
    runner whose four script-scope initialisers had not been copied with it.
    Under `Set-StrictMode -Version 2.0` that is a terminated session on the
    first line the runner writes.

    A TEXT SEARCH CANNOT FIND THIS. The name appears in the file; what it does
    not have is an assignment reachable before the read. So this walks the AST
    and models PowerShell's actual scoping:

      $script:X       must have an initialiser at FILE SCOPE - not inside a
                      function - and, for a read that is itself at file scope,
                      one that appears before it.
      a local         must be assigned, or be a parameter, in its own function.
      a bare name     may resolve outward: to a file-scope variable, or to a
                      local of a lexically enclosing function, which is how the
                      accepted Phase-6 fixture helper reads $controls.

    IT IS NOT A GENERAL POWERSHELL LINTER and does not try to be. It answers one
    question, over the files this project actually ships.

.NOTES
    Exit 0 and print CLEAN, or exit 1 and print one line per finding.
#>
param([string]$Path)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$errors = $null; $tokens = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
if ($errors) { $errors | ForEach-Object { 'PARSE ' + $_.Message }; exit 2 }

$AUTOMATIC = @('_','null','true','false','args','psscriptroot','psitem','pscmdlet',
               'error','lastexitcode','myinvocation','pwd','host','env','matches',
               'psboundparameters','input','this','ofs','psversiontable','foreach',
               'switch','stacktrace','executioncontext','profile','shellid',
               'iscoreclr','pshome','psculture','psuiculture','nestedpromptlevel')

function Get-VarName { param($V) return $V.VariablePath.UserPath.ToLowerInvariant() }

# The function that lexically encloses a node, or $null for file scope.
function Get-Enclosing { param($Node)
    $p = $Node.Parent
    while ($null -ne $p) {
        if ($p -is [System.Management.Automation.Language.FunctionDefinitionAst]) { return $p }
        $p = $p.Parent
    }
    return $null
}

function Get-Targets { param($Node)
    $out = @()
    foreach ($a in $Node.FindAll({param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst]}, $true)) {
        $left = $a.Left
        # Only a bare variable (or one being indexed / member-accessed) is a
        # definition of that variable's name.
        if ($left -is [System.Management.Automation.Language.VariableExpressionAst]) {
            $out += [pscustomobject]@{ Var = $left; Ast = $a }
        }
    }
    foreach ($f in $Node.FindAll({param($n) $n -is [System.Management.Automation.Language.ForEachStatementAst]}, $true)) {
        $out += [pscustomobject]@{ Var = $f.Variable; Ast = $f }
    }
    return $out
}

# ---- FILE-SCOPE ASSIGNMENTS: the only place a $script: initialiser may live ----
$fileScope = @{}
foreach ($t in Get-Targets -Node $ast) {
    if ($null -ne (Get-Enclosing -Node $t.Var)) { continue }
    $k = Get-VarName -V $t.Var
    if (-not $fileScope.ContainsKey($k)) { $fileScope[$k] = $t.Var.Extent.StartOffset }
}
# The bare name a $script: variable is also reachable under.
$fileBare = @{}
foreach ($k in $fileScope.Keys) {
    $b = ($k -replace '^script:', '')
    if (-not $fileBare.ContainsKey($b)) { $fileBare[$b] = $fileScope[$k] }
}

$findings = @()

# ---- EVERY FUNCTION, AND FILE SCOPE, CHECKED SEPARATELY ----
$scopes = @([pscustomobject]@{ Name = '<file scope>'; Node = $ast; IsFunction = $false })
foreach ($f in $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]}, $true)) {
    $scopes += [pscustomobject]@{ Name = $f.Name; Node = $f; IsFunction = $true }
}

foreach ($scope in $scopes) {
    $local = @{}
    foreach ($p in $scope.Node.FindAll({param($n) $n -is [System.Management.Automation.Language.ParameterAst]}, $true)) {
        $local[(Get-VarName -V $p.Name)] = 0
    }
    foreach ($t in Get-Targets -Node $scope.Node) {
        $k = Get-VarName -V $t.Var
        if (-not $local.ContainsKey($k)) { $local[$k] = $t.Var.Extent.StartOffset }
        elseif ($t.Var.Extent.StartOffset -lt $local[$k]) { $local[$k] = $t.Var.Extent.StartOffset }
    }
    foreach ($c in $scope.Node.FindAll({param($n) $n -is [System.Management.Automation.Language.CatchClauseAst]}, $true)) {
        $local['_'] = 0
    }

    foreach ($r in $scope.Node.FindAll({param($n) $n -is [System.Management.Automation.Language.VariableExpressionAst]}, $true)) {
        # Only the reads that BELONG to this scope.
        $owner = Get-Enclosing -Node $r
        if ($scope.IsFunction) { if ($owner -ne $scope.Node) { continue } }
        else { if ($null -ne $owner) { continue } }

        $name = $r.VariablePath.UserPath
        $k = $name.ToLowerInvariant()
        if ($AUTOMATIC -contains $k) { continue }
        if ($r.VariablePath.IsGlobal -or $r.VariablePath.DriveName -eq 'env') { continue }
        # The assignment target itself is not a read.
        $p = $r.Parent
        if ($p -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $p.Left.Extent.StartOffset -eq $r.Extent.StartOffset) { continue }
        if ($p -is [System.Management.Automation.Language.ForEachStatementAst] -and
            $p.Variable.Extent.StartOffset -eq $r.Extent.StartOffset) { continue }

        if ($r.VariablePath.IsScript) {
            $bare = ($k -replace '^script:', '')
            $at = $null
            if ($fileScope.ContainsKey($k))    { $at = $fileScope[$k] }
            elseif ($fileBare.ContainsKey($bare)) { $at = $fileBare[$bare] }
            if ($null -eq $at) {
                $findings += ('UNINITIALISED SCRIPT SCOPE  $' + $name + '  read in ' +
                              $scope.Name + ' at line ' + $r.Extent.StartLineNumber)
            } elseif ((-not $scope.IsFunction) -and ($at -gt $r.Extent.StartOffset)) {
                $findings += ('SCRIPT SCOPE READ BEFORE ITS INITIALISER  $' + $name +
                              '  at line ' + $r.Extent.StartLineNumber)
            }
            continue
        }
        if ($local.ContainsKey($k)) { continue }
        # A function may legitimately read a file-scope variable by its bare name.
        if ($scope.IsFunction -and ($fileScope.ContainsKey($k) -or $fileBare.ContainsKey($k))) { continue }
        # AND POWERSHELL SCOPING REACHES OUTWARD. A function defined inside
        # another sees its parent's locals, which is how the accepted Phase-6
        # fixture helper reads $controls. Walk the enclosing chain before
        # calling anything unassigned.
        if ($scope.IsFunction) {
            $outer = Get-Enclosing -Node $scope.Node
            $reached = $false
            while ($null -ne $outer) {
                foreach ($t in Get-Targets -Node $outer) {
                    if ((Get-VarName -V $t.Var) -eq $k) { $reached = $true }
                }
                foreach ($pp in $outer.FindAll({param($n) $n -is [System.Management.Automation.Language.ParameterAst]}, $true)) {
                    if ((Get-VarName -V $pp.Name) -eq $k) { $reached = $true }
                }
                $outer = Get-Enclosing -Node $outer
            }
            if ($reached) { continue }
        }
        $findings += ('UNASSIGNED  $' + $name + '  read in ' + $scope.Name +
                      ' at line ' + $r.Extent.StartLineNumber)
    }
}

# ---- RULE 2: A $script: INITIALISER MUST RUN BEFORE ANYTHING CALLS A HELPER ----
#
# Rule 1 accepts any file-scope assignment for a read inside a function, because
# a function body runs when it is called and not where it is written. That is
# too generous on its own: `$script:P8ZPath` WAS assigned at file scope - far
# down, after the banner had already been written through it. So the initialiser
# must come before the first top-level call to a function this file defines.
$ourFunctions = @{}
foreach ($f in $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst]}, $true)) {
    $ourFunctions[$f.Name.ToLowerInvariant()] = $true
}
$firstCall = [int]::MaxValue
foreach ($c in $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.CommandAst]}, $true)) {
    if ($null -ne (Get-Enclosing -Node $c)) { continue }
    $cn = $c.GetCommandName()
    if ($null -eq $cn) { continue }
    if (-not $ourFunctions.ContainsKey($cn.ToLowerInvariant())) { continue }
    if ($c.Extent.StartOffset -lt $firstCall) { $firstCall = $c.Extent.StartOffset }
}
if ($firstCall -lt [int]::MaxValue) {
    $scriptReads = @{}
    foreach ($r in $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.VariableExpressionAst]}, $true)) {
        if (-not $r.VariablePath.IsScript) { continue }
        $p2 = $r.Parent
        if ($p2 -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $p2.Left.Extent.StartOffset -eq $r.Extent.StartOffset) { continue }
        $scriptReads[$r.VariablePath.UserPath.ToLowerInvariant()] = $r.Extent.StartLineNumber
    }
    foreach ($k in $scriptReads.Keys) {
        $at = $null
        if ($fileScope.ContainsKey($k)) { $at = $fileScope[$k] }
        if ($null -eq $at) { continue }   # Rule 1 already reported it.
        if ($at -gt $firstCall) {
            $findings += ('SCRIPT SCOPE INITIALISED TOO LATE  $' + $k +
                          '  is first assigned after the first helper call')
        }
    }
}

# ---- RULE 3: WHAT A finally READS MUST BE ASSIGNED BEFORE ITS try ----
#
# A cleanup block runs however early the try failed - including before its first
# statement completed. A variable assigned only inside the try is not there yet.
foreach ($t in $ast.FindAll({param($n) $n -is [System.Management.Automation.Language.TryStatementAst]}, $true)) {
    if ($null -eq $t.Finally) { continue }
    $before = @{}
    $owner = Get-Enclosing -Node $t
    $ownerNode = $ast
    if ($null -ne $owner) { $ownerNode = $owner }
    foreach ($tt in Get-Targets -Node $ownerNode) {
        if ($tt.Var.Extent.StartOffset -lt $t.Extent.StartOffset) {
            $before[(Get-VarName -V $tt.Var)] = $true
        }
    }
    foreach ($pp in $ownerNode.FindAll({param($n) $n -is [System.Management.Automation.Language.ParameterAst]}, $true)) {
        $before[(Get-VarName -V $pp.Name)] = $true
    }
    # AND POWERSHELL SCOPING REACHES OUTWARD, so an enclosing function's
    # parameters and locals are in scope in a nested one's cleanup - which is
    # how the accepted Phase-6 restoration reads $SimInspection.
    $outerNode = $owner
    while ($null -ne $outerNode) {
        $outerNode = Get-Enclosing -Node $outerNode
        if ($null -eq $outerNode) {
            foreach ($tt in Get-Targets -Node $ast) {
                if ($null -eq (Get-Enclosing -Node $tt.Var)) { $before[(Get-VarName -V $tt.Var)] = $true }
            }
            break
        }
        foreach ($tt in Get-Targets -Node $outerNode) { $before[(Get-VarName -V $tt.Var)] = $true }
        foreach ($pp in $outerNode.FindAll({param($n) $n -is [System.Management.Automation.Language.ParameterAst]}, $true)) {
            $before[(Get-VarName -V $pp.Name)] = $true
        }
    }
    foreach ($r in $t.Finally.FindAll({param($n) $n -is [System.Management.Automation.Language.VariableExpressionAst]}, $true)) {
        $k = $r.VariablePath.UserPath.ToLowerInvariant()
        if ($AUTOMATIC -contains $k) { continue }
        if ($r.VariablePath.IsScript -or $r.VariablePath.IsGlobal) { continue }
        if ($r.VariablePath.DriveName -eq 'env') { continue }
        $p3 = $r.Parent
        if ($p3 -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $p3.Left.Extent.StartOffset -eq $r.Extent.StartOffset) { continue }
        if ($p3 -is [System.Management.Automation.Language.ForEachStatementAst] -and
            $p3.Variable.Extent.StartOffset -eq $r.Extent.StartOffset) { continue }
        # Assigned inside the finally itself, before this read.
        $inFinally = $false
        foreach ($tt in Get-Targets -Node $t.Finally) {
            if ((Get-VarName -V $tt.Var) -eq $k -and
                $tt.Var.Extent.StartOffset -lt $r.Extent.StartOffset) { $inFinally = $true }
        }
        if ($inFinally) { continue }
        if (-not $before.ContainsKey($k)) {
            $findings += ('CLEANUP READS A VARIABLE ASSIGNED ONLY INSIDE ITS try  $' +
                          $r.VariablePath.UserPath + '  at line ' + $r.Extent.StartLineNumber)
        }
    }
}

# ---- RULE 4: A VARIABLE WHOSE EVERY ASSIGNMENT READS ITSELF ----
#
# `$comAcquired = $comAcquired + 1` is a read as much as a write. If that is the
# ONLY kind of assignment a variable has, its first execution reads something
# that was never set - which is precisely what deleting a pre-`try` initialiser
# leaves behind, and what Rule 1 cannot see because an assignment does exist.
foreach ($scope in $scopes) {
    $selfOnly = @{}
    foreach ($t in Get-Targets -Node $scope.Node) {
        $owner = Get-Enclosing -Node $t.Var
        if ($scope.IsFunction) { if ($owner -ne $scope.Node) { continue } }
        else { if ($null -ne $owner) { continue } }
        $k = Get-VarName -V $t.Var
        if ($t.Ast -isnot [System.Management.Automation.Language.AssignmentStatementAst]) {
            $selfOnly[$k] = $false
            continue
        }
        $reads = $false
        foreach ($v in $t.Ast.Right.FindAll({param($n) $n -is [System.Management.Automation.Language.VariableExpressionAst]}, $true)) {
            if ((Get-VarName -V $v) -eq $k) { $reads = $true }
        }
        if (-not $selfOnly.ContainsKey($k)) { $selfOnly[$k] = $reads }
        elseif (-not $reads) { $selfOnly[$k] = $false }
    }
    foreach ($pp in $scope.Node.FindAll({param($n) $n -is [System.Management.Automation.Language.ParameterAst]}, $true)) {
        $selfOnly[(Get-VarName -V $pp.Name)] = $false
    }
    foreach ($k in $selfOnly.Keys) {
        if (-not $selfOnly[$k]) { continue }
        if ($k -like 'script:*') { continue }
        if ($scope.IsFunction -and ($fileScope.ContainsKey($k) -or $fileBare.ContainsKey($k))) { continue }
        $findings += ('EVERY ASSIGNMENT TO $' + $k + ' READS IT FIRST  in ' + $scope.Name +
                      '; it is never given a starting value')
    }
}

$findings | Sort-Object -Unique
if ($findings.Count -gt 0) { exit 1 }
Write-Output 'CLEAN'
