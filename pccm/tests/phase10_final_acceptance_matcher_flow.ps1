<#
.SYNOPSIS
    PCCM test-only harness: THE MODEL CHECK MATCHER, EXECUTED UNDER STRICT MODE.
.DESCRIPTION
    WHY THIS EXISTS. Final acceptance run 4 at 95e322f passed every check through
    Calculate and then died inside Assert-FaModelCheck with "The property
    'Subject' cannot be found on this object": the advisory expectation names no
    Subject, and under Set-StrictMode reading an absent hashtable key is an
    error. The static controls proved the expected SETS and never executed the
    matcher. This lifts the REAL Test-FaExpectedEntry, Assert-FaModelCheck,
    Format-FaActionable and Format-FaCell out of the runner BY AST, stands in a
    surface reader and a check recorder, reads the vocabulary from the REAL
    generated Phase-9 projection, and drives the matcher over the three real
    expectation shapes, two refusals and four malformed definitions, under
    Set-StrictMode 2.0. Excel is never started. IT ASSERTS NOTHING. It prints
    tagged lines and tests/test_phase10_final_acceptance_source.py decides.
.NOTES
    Prints:
      MATCH|<case>|ok=<True/False>|<detail>
      REFUSED|<case>|<message>              a throw the matcher raised on purpose
      ERROR|<case>|<exception type>|<message>   any other throw - the run-4 class
    Exit 0 always.
#>
param([string]$Runner, [string]$Projection)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$source = Get-Content -LiteralPath $Runner -Raw
$ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$null, [ref]$null)
foreach ($name in @('Format-FaCell', 'Format-FaActionable', 'Test-FaExpectedEntry', 'Assert-FaModelCheck')) {
    $definition = $ast.FindAll({ param($node)
        ($node -is [System.Management.Automation.Language.FunctionDefinitionAst]) -and ($node.Name -eq $name) }, $true)
    if (@($definition).Count -ne 1) { throw ('the runner defines ' + $name + ' ' + [string]@($definition).Count + ' times') }
    . ([scriptblock]::Create(@($definition)[0].Extent.Text))
}
$script:FaErrorCodes = @{ -2146826246 = '#N/A' }
# NOT $projection: the [string] parameter of that name would coerce the object back to text.
$inspection = Get-Content -LiteralPath $Projection -Raw | ConvertFrom-Json
$script:Surface = $null
# RECORDED, NOT EMITTED: the runner assigns the recorder's output to $null, so a
# stand-in that wrote to the output stream would be discarded exactly as run 12
# of the equivalence gate taught. Every line is printed at the end.
$script:Lines = New-Object System.Collections.ArrayList

# STAND-INS. The reader returns the case's surface; the recorder prints instead
# of throwing, so every case reports its own verdict.
function Get-FaModelCheckSurface { param($Workbook, $Projection) return $script:Surface }
function Add-FaCheck {
    param([string]$Scenario, [bool]$Ok, [string]$Detail = '', [switch]$Continue)
    $null = $script:Lines.Add('MATCH|' + $Scenario + '|ok=' + [string]$Ok + '|' + $Detail)
    return $Ok
}

function New-Row {
    param([string]$Id, [string]$Group, [string]$Severity, $Subject, [string]$Message)
    return [pscustomobject]@{ check_id = $Id; group = $Group; severity = $Severity; subject = $Subject; message = $Message; guidance = '' }
}
function New-Surface {
    param([string]$Overall, [int]$Errors, [int]$Warnings, $Rows)
    return [pscustomobject]@{ Summary = @{ overall_status = $Overall; error_count = $Errors; warning_count = $Warnings }; Shown = @($Rows) }
}
function Invoke-Case {
    param([string]$Case, $Surface, $Expected)
    $script:Surface = $Surface
    try {
        $null = Assert-FaModelCheck -Workbook $null -Projection $inspection -Scenario $Case -Expected $Expected
    } catch {
        $message = [string]$_.Exception.Message
        if ($message -like 'RUNNER DEFINITION ERROR*') { $null = $script:Lines.Add('REFUSED|' + $Case + '|' + $message) }
        else { $null = $script:Lines.Add('ERROR|' + $Case + '|' + $_.Exception.GetType().FullName + '|' + $message) }
    }
}

$advisoryMessage = [string]$inspection.advisory.message
$advisory = @{ Id = [string]$inspection.advisory.check_id; Severity = [string]$inspection.advisory.severity; Message = $advisoryMessage }
$advisoryRow = New-Row -Id ([string]$inspection.advisory.check_id) -Group 'Inputs' -Severity 'WARNING' -Subject 'Monte Carlo Iterations' -Message $advisoryMessage
$infoRow = New-Row -Id 'CAL-040' -Group 'Calculation' -Severity 'INFO' -Subject $null -Message 'current'

# A. THE ADVISORY: Id + Severity + Message, NO Subject - the run-4 shape.
Invoke-Case -Case 'A.advisory' -Expected @($advisory) `
    -Surface (New-Surface -Overall 'WARNING' -Errors 0 -Warnings 1 -Rows @($advisoryRow, $infoRow))
# B. THE INVALID ERROR: Id + Severity + Subject, NO Message, beside the advisory.
$calcError = @{ Id = 'CAL-010'; Severity = 'ERROR'; Subject = 'CL-001' }
Invoke-Case -Case 'B.invalid' -Expected @($calcError, $advisory) `
    -Surface (New-Surface -Overall 'ERROR' -Errors 1 -Warnings 1 -Rows @(
        (New-Row -Id 'CAL-010' -Group 'Calculation' -Severity 'ERROR' -Subject 'CL-001' -Message 'refused'),
        $advisoryRow,
        (New-Row -Id 'SIM-020' -Group 'Simulation' -Severity 'INFO' -Subject $null -Message 'context')))
# C. NOT CALCULATED: AnyOf + Severity + Subject, NO Id, NO Message, beside the advisory.
$notCalculated = @{ AnyOf = @('CAL-020', 'CAL-030'); Severity = 'WARNING'; Subject = '<blank>' }
Invoke-Case -Case 'C.after-reset' -Expected @($advisory, $notCalculated) `
    -Surface (New-Surface -Overall 'WARNING' -Errors 0 -Warnings 2 -Rows @(
        $advisoryRow,
        (New-Row -Id 'CAL-020' -Group 'Calculation' -Severity 'WARNING' -Subject $null -Message 'not calculated')))
# D. AN UNRELATED ACTIONABLE ROW must not satisfy the checkpoint.
Invoke-Case -Case 'D.unrelated' -Expected @($advisory) `
    -Surface (New-Surface -Overall 'WARNING' -Errors 0 -Warnings 2 -Rows @(
        $advisoryRow, (New-Row -Id 'ANN-050' -Group 'Annual' -Severity 'WARNING' -Subject $null -Message 'historical')))
# E. THE ADVISORY MISSING behind a matching overall word.
Invoke-Case -Case 'E.missing' -Expected @($advisory) `
    -Surface (New-Surface -Overall 'WARNING' -Errors 0 -Warnings 1 -Rows @(
        (New-Row -Id 'ANN-050' -Group 'Annual' -Severity 'WARNING' -Subject $null -Message 'historical')))
# F. A SUBJECT THAT DIFFERS is not a match.
Invoke-Case -Case 'F.wrong-subject' -Expected @($calcError, $advisory) `
    -Surface (New-Surface -Overall 'ERROR' -Errors 1 -Warnings 1 -Rows @(
        (New-Row -Id 'CAL-010' -Group 'Calculation' -Severity 'ERROR' -Subject 'CL-002' -Message 'refused'), $advisoryRow))
# G. THE ACTUAL INVALID SEQUENCE (run 5): the error, the advisory and the two
# historical Annual WARNINGs, expected as two AnyOf entries over the projected
# Annual WARNING population.
$annualWarningIds = @($inspection.evaluation.declared_checks | Where-Object {
    ([string]$_.group -ceq [string]$inspection.vocabulary.group_order[4]) -and ([string]$_.severity -ceq 'WARNING') } |
    ForEach-Object { [string]$_.check_id })
$annualHistorical = @{ AnyOf = $annualWarningIds; Severity = 'WARNING'; Subject = '<blank>' }
$run5Rows = @(
    (New-Row -Id 'CAL-010' -Group 'Calculation' -Severity 'ERROR' -Subject 'CL-001' -Message 'refused'),
    $advisoryRow,
    (New-Row -Id 'ANN-010' -Group 'Annual' -Severity 'WARNING' -Subject $null -Message 'historical profile'),
    (New-Row -Id 'ANN-050' -Group 'Annual' -Severity 'WARNING' -Subject $null -Message 'historical distributions'),
    (New-Row -Id 'SIM-020' -Group 'Simulation' -Severity 'INFO' -Subject $null -Message 'context'))
Invoke-Case -Case 'G.invalid-after-annual' -Expected @($calcError, $advisory, $annualHistorical, $annualHistorical) `
    -Surface (New-Surface -Overall 'ERROR' -Errors 1 -Warnings 3 -Rows $run5Rows)
# H. ONE HISTORICAL ANNUAL WARNING OMITTED from the expectation is refused.
Invoke-Case -Case 'H.one-annual-omitted' -Expected @($calcError, $advisory, $annualHistorical) `
    -Surface (New-Surface -Overall 'ERROR' -Errors 1 -Warnings 3 -Rows $run5Rows)
# MALFORMED DEFINITIONS are refused before any row is read.
$plain = New-Surface -Overall 'WARNING' -Errors 0 -Warnings 1 -Rows @($advisoryRow)
Invoke-Case -Case 'M1.both-selectors' -Surface $plain -Expected @(@{ Id = 'INP-010'; AnyOf = @('INP-010'); Severity = 'WARNING' })
Invoke-Case -Case 'M2.no-selector' -Surface $plain -Expected @(@{ Severity = 'WARNING' })
Invoke-Case -Case 'M3.no-severity' -Surface $plain -Expected @(@{ Id = 'INP-010' })
Invoke-Case -Case 'M4.empty-anyof' -Surface $plain -Expected @(@{ AnyOf = @(); Severity = 'WARNING' })
foreach ($line in $script:Lines) { Write-Output ([string]$line) }
exit 0
