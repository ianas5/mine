<#
.SYNOPSIS
    PCCM test-only audit: commands a PowerShell runner can CALL BUT NOT RESOLVE.

.DESCRIPTION
    WHY THIS EXISTS. Phase-9 Windows run 1 died on
    `The term 'Write-RowObject' is not recognized` after five checks - past the
    Stage-B bootstrap, past the VBAProject compile, past a worksheet-called
    adapter answering INVALID. Nothing in the runner called that function. A
    function in a file the runner DOT-SOURCES called it, and that file does not
    define it: its definition lives in a Phase-4 driver the runner deliberately
    does not dot-source. The same gap killed W1 and the timing harness before.

    A TEXT SEARCH CANNOT FIND THIS EITHER. The name does not appear in the
    runner at all. So this walks the AST of the runner AND of every file it
    dot-sources, builds the call graph, and asks one question of every command
    the runner can actually reach: does it resolve?

      a function        defined in the runner or in a dot-sourced file, and
                        defined EXACTLY ONCE across them - two definitions mean
                        the last one dot-sourced silently wins
      a cmdlet, alias   resolvable by Get-Command on this host
      anything else     a finding

    REACHABILITY IS COMPUTED, NOT ASSUMED. The dot-sourced Gate-B files define
    hundreds of functions the Phase-9 runner never calls, and many of those do
    depend on the Phase-4 driver. Requiring all of them to resolve would be a
    finding about a file this runner does not use. So the closure starts at the
    runner's own top-level statements and follows only what is actually called.

    IT IS NOT A GENERAL POWERSHELL LINTER. It answers one question, over the
    files this project ships.

.NOTES
    Exit 0 and print CLEAN, or exit 1 and print one line per finding.
#>
param(
    [string]$Path,
    # Commands that exist on Windows PowerShell 5.1 and not on the Linux pwsh
    # this audit runs under. Each is named, never pattern-matched, so a typo
    # cannot be absorbed by a wildcard.
    [string[]]$WindowsOnly = @('Get-WmiObject', 'Get-CimInstance')
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Get-Ast {
    param([string]$File)
    $errors = $null; $tokens = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $File, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        foreach ($e in $errors) { Write-Output ('PARSE ' + $File + ': ' + $e.Message) }
        exit 2
    }
    return $ast
}

# WHICH FILES ARE IN SCOPE: the runner, and every file it dot-sources at top
# level. A dot-source inside a function would be a different question and this
# project has none.
function Get-DotSourcedFiles {
    param($Ast, [string]$Directory)
    $found = New-Object System.Collections.ArrayList
    foreach ($command in $Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
        if ($command.InvocationOperator -ne [System.Management.Automation.Language.TokenKind]::Dot) { continue }
        $text = $command.Extent.Text
        if ($text -match "'([^']+\.ps1)'") {
            $null = $found.Add((Join-Path $Directory $Matches[1]))
        }
    }
    return $found
}

$runner = (Resolve-Path -LiteralPath $Path).Path
$directory = Split-Path -Parent $runner
$runnerAst = Get-Ast -File $runner
$files = New-Object System.Collections.ArrayList
$null = $files.Add($runner)
foreach ($f in @(Get-DotSourcedFiles -Ast $runnerAst -Directory $directory)) {
    if (Test-Path -LiteralPath $f) { $null = $files.Add((Resolve-Path -LiteralPath $f).Path) }
    else { Write-Output ('MISSING dot-sourced file: ' + $f) }
}

# EVERY FUNCTION DEFINITION, AND WHERE. A name defined twice is reported even if
# both definitions happen to be identical: the reader cannot tell which one runs.
$definitions = @{}
$asts = @{}
foreach ($file in $files) {
    $ast = if ($file -eq $runner) { $runnerAst } else { Get-Ast -File $file }
    $asts[$file] = $ast
    foreach ($fn in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        $key = $fn.Name.ToLowerInvariant()
        if (-not $definitions.ContainsKey($key)) {
            $definitions[$key] = New-Object System.Collections.ArrayList
        }
        $null = $definitions[$key].Add(
            (Split-Path -Leaf $file) + ':' + [string]$fn.Extent.StartLineNumber)
    }
}

# THE COMMANDS EACH FUNCTION NAMES, and the commands the runner's own top level
# names. A command whose name is not a bare constant - `& $bootstrap` - is not a
# name this audit can check and is skipped rather than guessed at.
function Get-CalledNames {
    param($Node)
    $names = New-Object System.Collections.ArrayList
    foreach ($command in $Node.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
        if ($command.InvocationOperator -eq [System.Management.Automation.Language.TokenKind]::Dot) { continue }
        $element = $command.CommandElements[0]
        if ($element -isnot [System.Management.Automation.Language.StringConstantExpressionAst]) { continue }
        $null = $names.Add([string]$element.Value)
    }
    return $names
}

$bodyOf = @{}
foreach ($file in $files) {
    foreach ($fn in $asts[$file].FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        $key = $fn.Name.ToLowerInvariant()
        if (-not $bodyOf.ContainsKey($key)) { $bodyOf[$key] = New-Object System.Collections.ArrayList }
        foreach ($n in @(Get-CalledNames -Node $fn.Body)) { $null = $bodyOf[$key].Add($n) }
    }
}

# THE CLOSURE. Start at the runner's top level - everything outside a function
# definition in the runner file - and follow calls through defined functions.
$topLevel = New-Object System.Collections.ArrayList
foreach ($n in @(Get-CalledNames -Node $runnerAst)) { $null = $topLevel.Add($n) }
$insideFunctions = New-Object System.Collections.Generic.HashSet[string]
foreach ($fn in $runnerAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    foreach ($n in @(Get-CalledNames -Node $fn.Body)) { $null = $insideFunctions.Add($n) }
}

$queue = New-Object System.Collections.Queue
$seen = New-Object System.Collections.Generic.HashSet[string]
foreach ($n in $topLevel) { if ($seen.Add($n.ToLowerInvariant())) { $queue.Enqueue($n) } }
$reached = New-Object System.Collections.ArrayList
while ($queue.Count -gt 0) {
    $name = [string]$queue.Dequeue()
    $null = $reached.Add($name)
    $key = $name.ToLowerInvariant()
    if ($bodyOf.ContainsKey($key)) {
        foreach ($callee in $bodyOf[$key]) {
            if ($seen.Add($callee.ToLowerInvariant())) { $queue.Enqueue($callee) }
        }
    }
}

$findings = New-Object System.Collections.ArrayList
foreach ($key in ($definitions.Keys | Sort-Object)) {
    if ($definitions[$key].Count -gt 1) {
        $null = $findings.Add('DUPLICATE ' + $key + ' defined at ' +
                              ($definitions[$key] -join ', '))
    }
}
foreach ($name in ($reached | Sort-Object -Unique)) {
    $key = $name.ToLowerInvariant()
    if ($definitions.ContainsKey($key)) { continue }
    if ($WindowsOnly -contains $name) { continue }
    $resolved = $null
    try { $resolved = Get-Command -Name $name -ErrorAction SilentlyContinue } catch { $resolved = $null }
    if ($null -eq $resolved) {
        $null = $findings.Add('UNRESOLVED ' + $name)
    }
}

if ($findings.Count -gt 0) {
    foreach ($finding in $findings) { Write-Output $finding }
    exit 1
}
Write-Output ('CLEAN ' + [string]$reached.Count + ' reachable commands over ' +
              [string]$files.Count + ' files')
exit 0
