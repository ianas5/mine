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
    [string[]]$WindowsOnly = @('Get-WmiObject', 'Get-CimInstance'),
    # FUNCTIONS THE RUNNER DELIBERATELY REDEFINES OVER A DOT-SOURCED FILE.
    #
    # The blanket rule below - a name defined twice is a finding - was true of
    # every runner in this tree until one needed to change a helper's behaviour
    # WITHOUT editing the accepted file that defines it. Forbidding that outright
    # would have forced the edit; permitting it silently would restore exactly the
    # ambiguity the rule exists to prevent.
    #
    # So it is DECLARED, and the declaration is CHECKED rather than believed. A
    # name listed here must be defined exactly twice - once in the runner, once in
    # a file the runner dot-sources - and the runner's definition must come AFTER
    # the dot-source statement that loads the other one, which is what makes it
    # the definition that actually wins. A declaration that is not true of the
    # files is a finding, so this is strictly stronger than the rule it relaxes:
    # nothing before checked the ORDER at all.
    #
    # Every caller that does not pass this keeps the original behaviour exactly.
    [string[]]$DeclaredOverride = @()
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
    # THE LINE COMES BACK WITH THE PATH. An override only wins if it is defined
    # after the dot-source that would otherwise overwrite it, so the order cannot
    # be checked without knowing where the dot-source is.
    $found = New-Object System.Collections.ArrayList
    foreach ($command in $Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
        if ($command.InvocationOperator -ne [System.Management.Automation.Language.TokenKind]::Dot) { continue }
        $text = $command.Extent.Text
        if ($text -match "'([^']+\.ps1)'") {
            $null = $found.Add([pscustomobject]@{
                Path = (Join-Path $Directory $Matches[1])
                Line = [int]$command.Extent.StartLineNumber
            })
        }
    }
    return $found
}

$runner = (Resolve-Path -LiteralPath $Path).Path
$directory = Split-Path -Parent $runner
$runnerAst = Get-Ast -File $runner
$files = New-Object System.Collections.ArrayList
$null = $files.Add($runner)
$dotSourceLine = @{}
foreach ($entry in @(Get-DotSourcedFiles -Ast $runnerAst -Directory $directory)) {
    if (Test-Path -LiteralPath $entry.Path) {
        $resolvedPath = (Resolve-Path -LiteralPath $entry.Path).Path
        $null = $files.Add($resolvedPath)
        $dotSourceLine[(Split-Path -Leaf $resolvedPath)] = [int]$entry.Line
    }
    else { Write-Output ('MISSING dot-sourced file: ' + $entry.Path) }
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
        $null = $definitions[$key].Add([pscustomobject]@{
            File = (Split-Path -Leaf $file)
            Line = [int]$fn.Extent.StartLineNumber
            IsRunner = ($file -eq $runner)
            Where = (Split-Path -Leaf $file) + ':' + [string]$fn.Extent.StartLineNumber
        })
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

# A DECLARATION THAT NAMES NOTHING IS A STALE EXEMPTION, and a stale exemption is
# where the next real duplicate hides. Every declared name must actually be
# overridden by this runner.
# COMMA-SEPARATED TOO, because `pwsh -File` hands every argument to the script as
# ONE string and cannot build an array - and -File is how every caller in this
# tree invokes this audit. A function name cannot contain a comma, so splitting
# on one is unambiguous. Blank fragments are dropped rather than becoming a name
# nothing can match.
$declaredNames = New-Object System.Collections.ArrayList
foreach ($declared in @($DeclaredOverride)) {
    foreach ($fragment in ([string]$declared).Split(',')) {
        $trimmed = $fragment.Trim()
        if ($trimmed.Length -gt 0) { $null = $declaredNames.Add($trimmed) }
    }
}

$declaredKeys = New-Object System.Collections.Generic.HashSet[string]
foreach ($declared in $declaredNames) {
    $null = $declaredKeys.Add(([string]$declared).ToLowerInvariant())
}
foreach ($declared in $declaredNames) {
    $key = ([string]$declared).ToLowerInvariant()
    if (-not $definitions.ContainsKey($key)) {
        $null = $findings.Add('DECLARED-OVERRIDE ' + $key +
                              ' is not defined anywhere in these files')
        continue
    }
    if ($definitions[$key].Count -lt 2) {
        $null = $findings.Add('DECLARED-OVERRIDE ' + $key +
                              ' overrides nothing: it is defined once, at ' +
                              $definitions[$key][0].Where)
    }
}

foreach ($key in ($definitions.Keys | Sort-Object)) {
    if ($definitions[$key].Count -le 1) { continue }
    $places = @($definitions[$key])
    if (-not $declaredKeys.Contains($key)) {
        $null = $findings.Add('DUPLICATE ' + $key + ' defined at ' +
                              (($places | ForEach-Object { $_.Where }) -join ', '))
        continue
    }
    # DECLARED: prove the runner's definition is the one that wins. Exactly two
    # definitions, one of them the runner's, and the runner's below the
    # dot-source that loads the other - otherwise the accepted file's definition
    # overwrites it and the runner silently runs the behaviour it meant to
    # replace, with its own corrected source still sitting in the file.
    if ($places.Count -ne 2) {
        $null = $findings.Add('DECLARED-OVERRIDE ' + $key + ' has ' +
                              [string]$places.Count + ' definitions, not two: ' +
                              (($places | ForEach-Object { $_.Where }) -join ', '))
        continue
    }
    $mine = @($places | Where-Object { $_.IsRunner })
    $theirs = @($places | Where-Object { -not $_.IsRunner })
    if (($mine.Count -ne 1) -or ($theirs.Count -ne 1)) {
        $null = $findings.Add('DECLARED-OVERRIDE ' + $key +
                              ' is not one runner definition over one dot-sourced one: ' +
                              (($places | ForEach-Object { $_.Where }) -join ', '))
        continue
    }
    if (-not $dotSourceLine.ContainsKey($theirs[0].File)) {
        $null = $findings.Add('DECLARED-OVERRIDE ' + $key + ': ' + $theirs[0].File +
                              ' is not dot-sourced by this runner')
        continue
    }
    $loadedAt = [int]$dotSourceLine[$theirs[0].File]
    if ([int]$mine[0].Line -lt $loadedAt) {
        $null = $findings.Add('DECLARED-OVERRIDE ' + $key + ' does not win: defined at ' +
                              $mine[0].Where + ', but ' + $theirs[0].File +
                              ' is dot-sourced at line ' + [string]$loadedAt +
                              ' and redefines it')
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
