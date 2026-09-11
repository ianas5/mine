<#
.SYNOPSIS
    PCCM test-only harness: the benchmark's SAMPLE and ITERATION shapes, executed.

.DESCRIPTION
    WHY THIS EXISTS. The PERF-SMALL run at ce5951f reached the timed section and
    lost it to two shape defects that every source-reading control passed:

      * `return ,@($problems)` in `Test-BenchmarkSample` paired with
        `@(Test-BenchmarkSample ...)` at the call site. `@()` collects pipeline
        items and does NOT flatten a nested array, so the count was 1 whatever
        the sample found - every execution INVALID, and `-join` rendered the
        inner array as `System.Object[]`.

      * the run loop's `$iterations` IS the script parameter `[int[]]$Iterations`,
        because PowerShell variable names are case-insensitive and a typed
        parameter keeps its constraint for the variable's whole life. Every
        assignment was coerced back to `[int[]]`, so `[double]$iterations` raised.

    Neither is a property of text. One is a COUNT and the other is a CLR TYPE, and
    a control that asserted them from source would be asserting a belief about
    the language. So this RUNS the real code.

    SECTION A lifts the runner's OWN param block and the run loop's OWN iteration
    selection lines, composes them into a script, and runs it once per planned run
    in the real plan - with and without `-Iterations` supplied - reporting the CLR
    type the selection produced.

    SECTION B lifts `Test-BenchmarkSample` and `Assert-BenchmarkProblemList` by
    AST and drives them over scripted evidence, reporting the count, the rendered
    message, and whether the guard refuses a deliberately re-wrapped list.

    Excel is never started and no workbook is opened.

.NOTES
    Prints tagged lines: ITER|... and SAMPLE|... . Exit 0 always; the pytest
    control reads the lines and decides.
#>
param(
    [string]$Runner,
    [string]$Plan
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here
if ([string]::IsNullOrWhiteSpace($Runner)) {
    $Runner = Join-Path $root 'bootstrap/windows/phase10_benchmark.ps1'
}
if ([string]::IsNullOrWhiteSpace($Plan)) {
    $Plan = Join-Path $root 'build/phase10_benchmark_plan.json'
}
$runnerPath = (Resolve-Path -LiteralPath $Runner).Path

$errors = $null; $tokens = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile(
    $runnerPath, [ref]$tokens, [ref]$errors)
if ($errors -and $errors.Count -gt 0) {
    Write-Output ('PARSE|' + [string]$errors.Count + ' parse error(s) in the runner')
    exit 0
}

# ===========================================================================
# SECTION A - THE ITERATION SELECTION, UNDER THE RUNNER'S OWN PARAM BLOCK
# ===========================================================================
# THE PARAM BLOCK IS THE DEFECT'S OTHER HALF, so it is lifted rather than
# retyped: a test that declared its own `[int[]]$Iterations` would be testing its
# own guess about what the runner declares.
if ($null -eq $ast.ParamBlock) {
    Write-Output 'MISSING|the runner has no param block'
    exit 0
}
$paramBlock = $ast.ParamBlock.Extent.Text

$source = Get-Content -LiteralPath $runnerPath -Raw
$selectMark = '        $runIterations = $null'
if ($source -notmatch [regex]::Escape($selectMark)) {
    Write-Output 'MISSING|the iteration selection anchor is not in the runner'
    exit 0
}
$from = $source.IndexOf($selectMark)
$rest = $source.Substring($from)
$endMark = ' iterations' + [char]39
$cut = $rest.IndexOf($endMark)
if ($cut -lt 0) {
    Write-Output 'MISSING|the iteration selection tail anchor is not in the runner'
    exit 0
}
# The two selection statements and the label line that follows them.
$selection = $rest.Substring(0, $cut + $endMark.Length) + ' }'

$probe = Join-Path ([System.IO.Path]::GetTempPath()) ('pccm-iter-probe-' +
    [System.Guid]::NewGuid().ToString('N') + '.ps1')
$body = @()
$body += $paramBlock
$body += 'Set-StrictMode -Version 2.0'
$body += '$ErrorActionPreference = ' + [char]39 + 'Stop' + [char]39
$body += '$run = [pscustomobject]@{ iterations = $env:PCCM_PROBE_ITERATIONS }'
$body += 'if ([string]::IsNullOrWhiteSpace([string]$run.iterations)) {'
$body += '    $run = [pscustomobject]@{ iterations = $null }'
$body += '} else {'
$body += '    $run = [pscustomobject]@{ iterations = [int]$env:PCCM_PROBE_ITERATIONS }'
$body += '}'
$body += '$operation = [pscustomobject]@{ label = ' + [char]39 + 'Probe' + [char]39 + ' }'
$body += $selection
$body += '$kind = $(if ($null -eq $runIterations) { ' + [char]39 + '<null>' + [char]39 +
         ' } else { $runIterations.GetType().FullName })'
$body += '$shown = $(if ($null -eq $runIterations) { ' + [char]39 + 'n/a' + [char]39 +
         ' } else { [string]$runIterations })'
$body += '$asDouble = ' + [char]39 + 'n/a' + [char]39
$body += 'if ($null -ne $runIterations) {'
$body += '    try { $asDouble = [string]([double]$runIterations) }'
$body += '    catch { $asDouble = ' + [char]39 + 'RAISED: ' + [char]39 +
         ' + $_.Exception.Message }'
$body += '}'
$body += 'Write-Output (' + [char]39 + 'PROBE|' + [char]39 +
         ' + $kind + ' + [char]39 + '|' + [char]39 + ' + $shown + ' + [char]39 + '|' +
         [char]39 + ' + $asDouble + ' + [char]39 + '|' + [char]39 + ' + $label)'
Set-Content -LiteralPath $probe -Value ($body -join "`r`n") -Encoding ASCII

function Invoke-IterationProbe {
    param([string]$Label, $Iterations, [switch]$SupplyParameter)
    if ($null -eq $Iterations) { $env:PCCM_PROBE_ITERATIONS = '' }
    else { $env:PCCM_PROBE_ITERATIONS = [string]$Iterations }
    # THE REAL PARAM BLOCK IS LIFTED, SO ITS REAL REQUIREMENTS APPLY: -Scenario
    # is mandatory with a ValidateSet. Supplying it is not a workaround - a probe
    # that declared a looser block would be testing its own guess about the
    # constraint that caused the defect.
    $arguments = @('-NoProfile', '-File', $probe, '-Scenario', 'PERF-SMALL')
    if ($SupplyParameter -and ($null -ne $Iterations)) {
        # A LEGITIMATE SCOPED INVOCATION. The parameter really is an [int[]]; the
        # defect was never its declaration, it was a local reusing its name.
        $arguments += @('-Iterations', [string]$Iterations)
    }
    $out = & (Get-Process -Id $PID).Path @arguments 2>&1
    $line = ''
    foreach ($item in @($out)) {
        if ([string]$item -like 'PROBE|*') { $line = [string]$item }
    }
    if ([string]::IsNullOrWhiteSpace($line)) { $line = 'PROBE|<none>|<none>|<none>|' + ($out -join ' ') }
    Write-Output ('ITER|' + $Label + '|' + $line.Substring('PROBE|'.Length))
}

$planned = @()
if (Test-Path -LiteralPath $Plan) {
    $planJson = Get-Content -LiteralPath $Plan -Raw | ConvertFrom-Json
    foreach ($run in @($planJson.runs)) {
        if ([string]$run.scenario -ne 'PERF-SMALL') { continue }
        $planned += [pscustomobject]@{
            Operation = [string]$run.operation
            Iterations = $(if ($null -eq $run.iterations) { $null } else { [int]$run.iterations })
        }
    }
}
if ($planned.Count -eq 0) {
    Write-Output 'MISSING|the plan declares no PERF-SMALL runs'
} else {
    foreach ($entry in $planned) {
        Invoke-IterationProbe -Label ('plan-' + $entry.Operation + '-' +
            $(if ($null -eq $entry.Iterations) { 'none' } else { [string]$entry.Iterations })) `
            -Iterations $entry.Iterations
    }
    # AND WITH THE PARAMETER ACTUALLY SUPPLIED, which is the state the constraint
    # was written for and the one a scoped run really uses.
    Invoke-IterationProbe -Label 'parameter-supplied-10000' -Iterations 10000 -SupplyParameter
}
Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
$env:PCCM_PROBE_ITERATIONS = ''

# ===========================================================================
# SECTION B - THE SAMPLE VALIDATOR AND ITS SHAPE GUARD
# ===========================================================================
$wanted = @('Test-BenchmarkSample', 'Assert-BenchmarkProblemList')
$found = @{}
foreach ($fn in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    if ($wanted -contains $fn.Name) { $found[$fn.Name] = $fn.Extent.Text }
}
foreach ($name in $wanted) {
    if (-not $found.ContainsKey($name)) {
        Write-Output ('MISSING|' + $name + ' is not defined in the runner')
        exit 0
    }
    Invoke-Expression $found[$name]
}

function New-Evidence {
    param([string]$Result = 'OK|done', $Published = $null, [string]$Annual = 'CURRENT')
    $evidence = @{}
    $evidence['automation_result'] = $Result
    $evidence['published_iterations'] = $Published
    $evidence['annual_distribution_state'] = $Annual
    return $evidence
}

function New-Operation {
    param([string]$Key = 'calculate', [string]$Kind = 'command',
          [bool]$IterationDependent = $false)
    return [pscustomobject]@{ key = $Key; kind = $Kind
                              iteration_dependent = $IterationDependent }
}

$cases = @(
    @{ Name = 'clean-command'; Operation = (New-Operation)
       Evidence = (New-Evidence); Requested = $null; Wrap = $false },
    @{ Name = 'clean-recalculation'
       Operation = (New-Operation -Key 'recalculation' -Kind 'recalculation')
       Evidence = (New-Evidence -Result 'OK|recalculation'); Requested = $null; Wrap = $false },
    @{ Name = 'endpoint-refusal'; Operation = (New-Operation)
       Evidence = (New-Evidence -Result 'FAIL|Calculate|Discount Rate: the value is blank')
       Requested = $null; Wrap = $false },
    @{ Name = 'malformed-announcement'; Operation = (New-Operation)
       Evidence = (New-Evidence -Result 'something else entirely'); Requested = $null; Wrap = $false },
    @{ Name = 'missing-published-iterations'
       Operation = (New-Operation -Key 'simulation' -IterationDependent $true)
       Evidence = (New-Evidence); Requested = 10000; Wrap = $false },
    @{ Name = 'wrong-published-iterations'
       Operation = (New-Operation -Key 'simulation' -IterationDependent $true)
       Evidence = (New-Evidence -Published 10000); Requested = 100000; Wrap = $false },
    @{ Name = 'right-published-iterations'
       Operation = (New-Operation -Key 'simulation' -IterationDependent $true)
       Evidence = (New-Evidence -Published 50000); Requested = 50000; Wrap = $false },
    @{ Name = 'annual-state-wrong'
       Operation = (New-Operation -Key 'annual' -IterationDependent $true)
       Evidence = (New-Evidence -Published 10000 -Annual 'PENDING'); Requested = 10000; Wrap = $false },
    @{ Name = 'two-problems-at-once'
       Operation = (New-Operation -Key 'annual' -IterationDependent $true)
       Evidence = (New-Evidence -Result 'FAIL|Annual|refused' -Published 10 -Annual 'PENDING')
       Requested = 10000; Wrap = $false },
    # THE DEFECT ITSELF, DELIBERATELY RE-CREATED. The guard must refuse it.
    @{ Name = 'double-wrapped'; Operation = (New-Operation)
       Evidence = (New-Evidence); Requested = $null; Wrap = $true }
)

foreach ($case in $cases) {
    if ([bool]$case.Wrap) {
        # THE DEFECT, RECREATED EXACTLY. `@()` must be applied TO THE CALL: it
        # collects the ONE object that `return ,@(...)` emits into a new
        # one-element array whose element is the real list. Wrapping an
        # already-assigned array is a no-op and reproduces nothing.
        $problems = @(Test-BenchmarkSample -Operation $case.Operation -Evidence $case.Evidence `
            -RequestedIterations $case.Requested)
    } else {
        $problems = Test-BenchmarkSample -Operation $case.Operation -Evidence $case.Evidence `
            -RequestedIterations $case.Requested
    }
    $guard = 'ACCEPTED'
    try {
        Assert-BenchmarkProblemList -Problems $problems -Where ([string]$case.Operation.key)
    } catch {
        $guard = 'REFUSED: ' + ((([string]$_.Exception.Message) -replace '\s+', ' '))
    }
    $count = @($problems).Count
    $valid = ([bool]($count -eq 0))
    $rendered = (@($problems) -join '; ')
    Write-Output ('SAMPLE|' + $case.Name + '|count=' + [string]$count + '|valid=' + [string]$valid +
                  '|guard=' + $guard + '|rendered=' + $rendered)
}
exit 0
