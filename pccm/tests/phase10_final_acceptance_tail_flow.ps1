<#
.SYNOPSIS
    Executes the final acceptance runner's remaining-tail PREDICATES - the
    parenthesised expression each Add-FaCheck decides on - lifted out of the
    runner by AST and evaluated over stand-in states under Set-StrictMode 2.0.
    Excel is never started. The projected vocabulary is bound exactly as the
    runner binds it: the runner's own `$status... = [string]$p7...` assignment
    lines are executed over the accepted Phase-7 inspection, and the protection
    projection is the accepted Phase-10 one. IT ASSERTS NOTHING: it prints
    TAIL|<case>|<True or False> lines and the Python control decides.

    FINAL ACCEPTANCE RUN 18 (ad330bc): reset.states failed on the production-
    derived state calc=NOT CALCULATED sim=INVALID annual=NOT PRODUCED
    profile=NOT PRODUCED because the runner expected a blank simulation; the
    same expectation sat in refused.outcome.annual; reset.rollback never named
    the profile; and rowo.open-suppressed expected an unprotected file where
    Stage B saves a protected one. Each predicate is exercised here on the
    state production derives, and on the states it must refuse.
#>
param([string]$Runner, [string]$BuildDir)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$source = Get-Content -LiteralPath $Runner -Raw
$ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$null, [ref]$null)

# THE PROJECTIONS, accepted, and the runner's own vocabulary bindings over them.
$p7 = Get-Content -LiteralPath (Join-Path $BuildDir 'phase7_acceptance_inspection.json') -Raw | ConvertFrom-Json
$protection = Get-Content -LiteralPath (Join-Path $BuildDir 'phase10_protection_inspection.json') -Raw | ConvertFrom-Json
$bindings = @($ast.FindAll({ param($node)
    ($node -is [System.Management.Automation.Language.AssignmentStatementAst]) -and
    ((($node.Left.Extent.Text -match '^\$(status[A-Za-z]+|attempt[A-Za-z]+|annualNotProduced|profileNotProduced|annualHistorical|profileHistorical)$') -and
      ($node.Right.Extent.Text -match '^\[string\]\$p7\.')) -or
     (($node.Left.Extent.Text -match '^\$script:(ResetFailpoint|OpenFailpoint)$') -and
      ($node.Right -is [System.Management.Automation.Language.CommandExpressionAst]) -and
      ($node.Right.Expression -is [System.Management.Automation.Language.StringConstantExpressionAst]))) }, $false))
foreach ($binding in $bindings) { . ([scriptblock]::Create($binding.Extent.Text)) }
foreach ($required in @('statusNotCalculated', 'statusCurrent', 'statusInvalid', 'annualNotProduced', 'profileNotProduced', 'profileHistorical')) {
    if (-not (Test-Path -LiteralPath ('variable:' + $required))) { throw ('the runner does not bind $' + $required + ' from the projection') }
}

# THE PREDICATE of one recorded check: the second argument of its Add-FaCheck.
function Get-Predicate {
    param([string]$Scenario)
    $calls = @($ast.FindAll({ param($node)
        ($node -is [System.Management.Automation.Language.CommandAst]) -and
        ($node.GetCommandName() -eq 'Add-FaCheck') -and
        (@($node.CommandElements).Count -ge 3) -and
        ($node.CommandElements[1].Extent.Text -eq ("'" + $Scenario + "'")) }, $true))
    if ($calls.Count -ne 1) { throw ('the runner records ' + $Scenario + ' ' + [string]$calls.Count + ' time(s)') }
    return [scriptblock]::Create($calls[0].CommandElements[2].Extent.Text)
}
function Emit { param([string]$Case, $Value) Write-Output ('TAIL|' + $Case + '|' + [string]$Value) }
function New-States {
    param([string]$Calculation, [string]$Simulation, [string]$Annual, [string]$Profile)
    return [pscustomobject]@{ Calculation = $Calculation; Simulation = $Simulation; Annual = $Annual; Profile = $Profile }
}
$declared = @($protection.sheets).Count
function New-ProtectionState {
    param([bool]$Applied, [int]$Protected, [bool]$Structure, [int]$Depth)
    return [pscustomobject]@{ Applied = $Applied; Depth = $Depth; Structure = $Structure; Sheets = $declared; Protected = $Protected
                              Raw = ('OK|applied=' + [string]$Applied + '|depth=' + [string]$Depth + '|structure=' + [string]$Structure +
                                     '|sheets=' + [string]$declared + '|protected=' + [string]$Protected) }
}
Emit 'vocabulary' ('calc=' + $statusNotCalculated + '|sim=' + $statusInvalid + '|annual=' + $annualNotProduced + '|profile=' + $profileNotProduced + '|sheets=' + [string]$declared)

# 1. reset.states: the production-derived state passes; a blank, CURRENT or
#    wrongly-calculated variant, and an annual product left behind, do not.
$predicate = Get-Predicate 'reset.states'
$afterReset = New-States $statusNotCalculated $statusInvalid $annualNotProduced $profileNotProduced
$statesReset = $afterReset
Emit 'reset.states.production-derived' (& $predicate)
$statesReset = New-States $statusNotCalculated '' $annualNotProduced $profileNotProduced
Emit 'reset.states.blank-simulation' (& $predicate)
$statesReset = New-States $statusNotCalculated $statusCurrent $annualNotProduced $profileNotProduced
Emit 'reset.states.current-simulation' (& $predicate)
$statesReset = New-States $statusInvalid $statusInvalid $annualNotProduced $profileNotProduced
Emit 'reset.states.invalid-calculation' (& $predicate)
$statesReset = New-States $statusNotCalculated $statusInvalid $statusCurrent $profileNotProduced
Emit 'reset.states.annual-left' (& $predicate)
$statesReset = New-States $statusNotCalculated $statusInvalid $annualNotProduced $statusCurrent
Emit 'reset.states.profile-left' (& $predicate)

# 2. refused.outcome.annual: a FAIL and the unchanged derived state pass; the
#    production refusal text is composed as modAppState.Announce composes it.
$predicate = Get-Predicate 'refused.outcome.annual'
$refusedAnnual = 'FAIL|The annual stochastic answer was not produced.|the simulation is INVALID; the annual step needs a CURRENT simulation'
$statesAfterRefusal = $afterReset
Emit 'refused.annual.unchanged' (& $predicate)
$statesAfterRefusal = New-States $statusNotCalculated '' $annualNotProduced $profileNotProduced
Emit 'refused.annual.blank-simulation' (& $predicate)
$statesAfterRefusal = New-States $statusNotCalculated $statusInvalid $annualNotProduced $statusCurrent
Emit 'refused.annual.profile-left' (& $predicate)
$statesAfterRefusal = $afterReset
$refusedAnnual = 'OK|The annual stochastic answer was produced.'
Emit 'refused.annual.not-refused' (& $predicate)

# 3. reset.rollback: production's exact composed failure - FAIL|message|detail,
#    the failpoint's own sentence, a CRLF, the put-back sentence - with every
#    digest equal and every derived state CURRENT passes; a historical profile,
#    a moved publication or a success do not.
$predicate = Get-Predicate 'reset.rollback'
$injected = ("FAIL|Reset Results|Injected structural failure after stage '" + $script:ResetFailpoint + "'." + "`r`n" +
             'Every publication this command had cleared was put back, and no input, counter or identity was touched at any point.')
$publicationFull = 'publication'; $publicationRolledBack = 'publication'
$simStateFull = 'full:' + "`r`n" + '    shared.x = 1'; $simStateRolledBack = $simStateFull
$preservedFull = 'kept'; $preservedRolledBack = 'kept'
$statesRolledBack = New-States $statusCurrent $statusCurrent $statusCurrent $statusCurrent
Emit 'rollback.restored' (& $predicate)
$statesRolledBack = New-States $statusCurrent $statusCurrent $statusCurrent $profileHistorical
Emit 'rollback.profile-historical' (& $predicate)
$statesRolledBack = New-States $statusCurrent $statusCurrent $statusCurrent $statusCurrent
$publicationRolledBack = 'moved'
Emit 'rollback.publication-moved' (& $predicate)
$publicationRolledBack = 'publication'
$injected = 'OK|Results reset. Model inputs and identity counters were preserved.'
Emit 'rollback.not-failed' (& $predicate)

# 4. rowo.open-suppressed: the file as Stage B saved it - every declared sheet
#    and the structure protected, the window closed, nothing recorded - passes;
#    an unprotected file, an open window, a record, or no state do not.
$predicate = Get-Predicate 'rowo.open-suppressed'
$resultBeforeHandler = ''
$suppressed = New-ProtectionState $true $declared $true 0
Emit 'open-suppressed.saved-protection' (& $predicate)
$suppressed = New-ProtectionState $false 0 $false 0
Emit 'open-suppressed.unprotected-file' (& $predicate)
$suppressed = New-ProtectionState $true ($declared - 1) $true 0
Emit 'open-suppressed.one-sheet-short' (& $predicate)
$suppressed = New-ProtectionState $true $declared $true 1
Emit 'open-suppressed.window-open' (& $predicate)
$suppressed = New-ProtectionState $true $declared $true 0
$resultBeforeHandler = 'Workbook_Open: something'
Emit 'open-suppressed.recorded' (& $predicate)
$resultBeforeHandler = ''
$suppressed = $null
Emit 'open-suppressed.no-state' (& $predicate)

# 5. rowo.released-not-half-protected, the other half of the pair: only the
#    fully released state passes; the saved-protected state does not.
$predicate = Get-Predicate 'rowo.released-not-half-protected'
$released = New-ProtectionState $false 0 $false 0
Emit 'released.unprotected' (& $predicate)
$released = New-ProtectionState $true $declared $true 0
Emit 'released.still-protected' (& $predicate)
$released = New-ProtectionState $false 1 $false 0
Emit 'released.half-protected' (& $predicate)
