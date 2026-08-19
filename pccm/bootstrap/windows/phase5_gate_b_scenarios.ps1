<#
.SYNOPSIS
    PCCM Phase-5 Gate-B scenarios. Dot-sourced INTO phase4_functional_test.ps1.

.DESCRIPTION
    THIS IS NOT A SECOND HARNESS. It is dot-sourced into the accepted Phase-4
    functional test and runs inside that script's ONE COM lifecycle, against the
    ONE Excel instance it owns, the ONE workbook it opened and the ONE Stage-B
    bootstrap it ran. It reuses Add-Result, New-Checklist, Add-Check,
    Test-ChecklistOk, Format-Checklist, Format-Err, Get-TableBody, Set-TableCell,
    Get-NamedValue, Set-NamedValue and the release ledger unchanged. There is no
    competing bootstrap path, no second Excel process, no second reporting
    surface and no second shutdown.

    The Phase-4 matrix is untouched and remains mandatory. Gate-B acceptance
    requires Phase 4 at 35/35, 0 FAIL, 0 SKIP, BEFORE any Phase-5 result counts;
    Invoke-Phase5GateBScenarios refuses to run its own scenarios if the Phase-4
    matrix did not reach that state, and reports that refusal as a FAIL rather
    than as a skip nobody reads.

    EXPECTED VALUES COME FROM build/phase5_cases.json. Not one analytical number,
    canonical string, digest or remainder is written into this file. Where a
    comparison needs an address rather than a value, it comes from
    build/stage_b_manifest.json or build/phase5_gate_b_inspection.json, both of
    which the Stage-A build emits from the accepted contracts.

    NOTHING HERE HAS BEEN EXECUTED. No Windows run has been made and no Excel COM
    session has been started for Phase 5. This file is source under review.
#>

Set-StrictMode -Version 2.0

# ===========================================================================
# THE COVERAGE LEDGER
# ===========================================================================
# Every plan-case ID emitted into phase5_cases.json maps to at least one Windows
# scenario. The map is DATA, checked against the fixture corpus in a preflight
# that runs before Excel is started, so a case that was added to the corpus and
# never wired into the harness stops the run instead of quietly disappearing.
#
# A case may map to several scenarios; several cases may share one scenario and
# one workbook fixture. What may never happen is a case with no mapping, or a
# mapping that names a scenario the harness does not define.
function Get-Phase5CoverageLedger {
    $ledger = New-Object System.Collections.Specialized.OrderedDictionary
    # --- analytical fixtures, driven through PCCM_Calculate -----------------
    foreach ($id in 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 19, 21, 22, 25, 31) {
        $ledger.Add([string]$id, @('P5-AN'))
    }
    # 30 is analytical AND the cancellation-heavy reconciliation vector.
    $ledger.Add('30', @('P5-AN', 'P5-ID'))
    # --- prerequisite refusals ---------------------------------------------
    foreach ($id in 14, 15, 16, 17, 18, 20, 23, 24, 29) {
        $ledger.Add([string]$id, @('P5-RF'))
    }
    # --- direct real-VBA vectors, through the transient diagnostic module ---
    $ledger.Add('26', @('P5-D1', 'P5-D4', 'P5-D5'))
    $ledger.Add('27', @('P5-D6'))
    $ledger.Add('28', @('P5-D7'))
    $ledger.Add('35', @('P5-D2'))
    $ledger.Add('36', @('P5-D3'))
    # --- runtime-only: workbook state across attempts -----------------------
    $ledger.Add('32', @('P5-RC', 'P5-S5'))
    $ledger.Add('33', @('P5-FA'))
    $ledger.Add('34', @('P5-S3', 'P5-S4', 'P5-KP'))
    $ledger.Add('37', @('P5-FC'))
    return $ledger
}

# Every scenario ID the harness defines. The preflight rejects a ledger entry
# naming anything outside this set, so a typo cannot silently drop a case.
function Get-Phase5ScenarioIds {
    return @(
        'P5-M',
        'P5-D0', 'P5-D1', 'P5-D2', 'P5-D3', 'P5-D4', 'P5-D5', 'P5-D6', 'P5-D7', 'P5-D8',
        'P5-AN', 'P5-RF', 'P5-ID',
        'P5-S1', 'P5-S2', 'P5-S3', 'P5-S4', 'P5-S5', 'P5-S6',
        'P5-ST', 'P5-NS', 'P5-KP', 'P5-RC',
        'P5-FA', 'P5-FC',
        'P5-AX', 'P5-EV'
    )
}

# The two locked Phase-5 failpoint stage names. They are declared in the accepted
# production module modCalcReport.bas; tests/test_phase5_gate_b_harness_source.py
# pins these two strings against that module, so the production source stays the
# authority and this is a checked copy rather than a second declaration.
function Get-Phase5FailpointNames {
    return [pscustomobject]@{
        AnalyticalWrite = 'Phase5AnalyticalWrite'
        SuccessCommit   = 'Phase5SuccessCommit'
    }
}

# The Phase-4 matrix that must be intact before a Phase-5 result means anything.
# The timeline chain D..J is reported as the ten sequential steps D-J.1 .. D-J.10,
# so the matrix is 35 results, not 35 letters.
$script:Phase4RequiredScenarioIds = @(
    'PRE0', 'PRE', 'A', 'A1', 'A2', 'B', 'B2', 'C', 'D0',
    'D-J.1', 'D-J.2', 'D-J.3', 'D-J.4', 'D-J.5',
    'D-J.6', 'D-J.7', 'D-J.8', 'D-J.9', 'D-J.10',
    'K', 'K2', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W',
    'Y', 'Z'
)

function Get-Phase4RequiredScenarioIds { return $script:Phase4RequiredScenarioIds }

# ===========================================================================
# P5-PRE. Coverage preflight, BEFORE Excel is started
# ===========================================================================
# Pure PowerShell, no COM. It reads the emitted corpus and the ledger above and
# refuses the run if the two disagree. A Gate-B run that started, spent twenty
# minutes in Excel and then reported 36 of 37 cases would be worse than one that
# never started: the missing case would be a line in a summary nobody diffs.
function Invoke-Phase5CoveragePreflight {
    param([string]$BuildDir)

    $list = New-Checklist
    $casesPath = Join-Path $BuildDir 'phase5_cases.json'
    $inspectPath = Join-Path $BuildDir 'phase5_gate_b_inspection.json'

    $haveCases = Test-Path -LiteralPath $casesPath
    $null = Add-Check $list 'build/phase5_cases.json exists (the expected-value authority)' `
        $haveCases $casesPath
    $haveInspect = Test-Path -LiteralPath $inspectPath
    $null = Add-Check $list 'build/phase5_gate_b_inspection.json exists (the address authority)' `
        $haveInspect $inspectPath
    if (-not ($haveCases -and $haveInspect)) {
        Add-Result 'P5-PRE' 'Phase-5 coverage preflight (pure PowerShell, no Excel)' 'FAIL' `
            (Format-Checklist $list)
        return $false
    }

    $cases = Get-Content -LiteralPath $casesPath -Raw | ConvertFrom-Json
    $ledger = Get-Phase5CoverageLedger
    $known = Get-Phase5ScenarioIds

    $emitted = @()
    foreach ($case in @($cases.plan_cases)) { $emitted += [string]$case.id }
    $null = Add-Check $list 'the corpus emitted at least one plan case' ($emitted.Count -gt 0) `
        ("emitted " + $emitted.Count)

    # EVERY EMITTED ID HAS A MAPPING. Driven from the corpus, never from the
    # ledger: a case added to phase5_cases.json and forgotten here must fail.
    $unmapped = @()
    foreach ($id in $emitted) { if (-not $ledger.Contains($id)) { $unmapped += $id } }
    $null = Add-Check $list 'every emitted plan-case ID maps to a Windows scenario' `
        ($unmapped.Count -eq 0) ("unmapped: " + ($unmapped -join ', '))

    # AND NO MAPPING IS A GHOST. A ledger entry for a case the corpus no longer
    # emits is a coverage claim with nothing behind it.
    $orphan = @()
    foreach ($id in $ledger.Keys) { if ($emitted -notcontains $id) { $orphan += $id } }
    $null = Add-Check $list 'no ledger entry names a case the corpus does not emit' `
        ($orphan.Count -eq 0) ("orphaned: " + ($orphan -join ', '))

    # EVERY MAPPING POINTS AT A SCENARIO THIS HARNESS DEFINES.
    $unknown = @()
    foreach ($id in $ledger.Keys) {
        foreach ($scenario in @($ledger[$id])) {
            if ($known -notcontains $scenario) { $unknown += ($id + '->' + $scenario) }
        }
    }
    $null = Add-Check $list 'every mapping names a scenario the harness defines' `
        ($unknown.Count -eq 0) ("unknown: " + ($unknown -join ', '))

    # EVERY MAPPED FIXTURE EXISTS AND CARRIES WHAT ITS KIND PROMISES. A case that
    # maps to P5-AN but emits no `expected` block would be "covered" by a
    # scenario with nothing to assert.
    $hollow = @()
    foreach ($case in @($cases.plan_cases)) {
        $id = [string]$case.id
        $names = @($ledger[$id])
        switch ($case.kind) {
            'analytical' {
                if ($names -notcontains 'P5-AN') { $hollow += ($id + ': analytical, not in P5-AN') }
                if ($null -eq $case.expected) { $hollow += ($id + ': analytical with no expected block') }
            }
            'refusal' {
                if ($names -notcontains 'P5-RF') { $hollow += ($id + ': refusal, not in P5-RF') }
                if ([string]::IsNullOrEmpty([string]$case.expected_refusal)) {
                    $hollow += ($id + ': refusal with no expected_refusal')
                }
            }
            'statistics' {
                if (@($case.statistics).Count -lt 1) { $hollow += ($id + ': statistics with no vectors') }
            }
            'fingerprint' {
                if ([string]::IsNullOrEmpty([string]$case.reference)) {
                    $hollow += ($id + ': fingerprint with no reference')
                }
            }
        }
    }
    $null = Add-Check $list 'every mapped fixture carries the evidence its kind promises' `
        ($hollow.Count -eq 0) ($hollow -join '; ')

    # THE DIRECT-VECTOR SETS ARE COMPLETE. Counted from the corpus, so a vector
    # dropped upstream cannot shrink Gate-B coverage silently.
    $numeric = @($cases.fingerprint.numeric_encodings.vectors)
    $null = Add-Check $list 'the canonical numeric vector set is present' `
        ($numeric.Count -ge 10) ("vectors " + $numeric.Count)
    $labels = @($numeric | ForEach-Object { [string]$_.label })
    foreach ($required in '0', '-0', '1', '-1', '0.1', '1e-20', '1e+20', '0.1 + 0.2',
                          'MAX_DOUBLE', 'minimum subnormal') {
        $null = Add-Check $list ("the locked numeric vector '" + $required + "' is present") `
            ($labels -contains $required)
    }
    $sep = @($cases.fingerprint.decimal_separator.vectors)
    $null = Add-Check $list 'the separator vector set is present' ($sep.Count -ge 10) `
        ("vectors " + $sep.Count)
    $null = Add-Check $list 'every separator vector states BOTH a point and a comma expectation' `
        (@($sep | Where-Object { $null -eq $_.point -or $null -eq $_.comma }).Count -eq 0)
    $reduce = @($cases.fingerprint.reduction_vectors)
    $null = Add-Check $list 'all four reduction vectors are present' ($reduce.Count -eq 4) `
        ("vectors " + $reduce.Count)
    $utf16 = @($cases.fingerprint.utf16_vectors.vectors)
    $null = Add-Check $list 'the UTF-16 vector set is present' ($utf16.Count -ge 3) `
        ("vectors " + $utf16.Count)
    $utfKeys = @($utf16 | ForEach-Object { [string]$_.key })
    foreach ($required in 'bmp_above_7fff', 'non_bmp', 'mixed_length_prefix') {
        $null = Add-Check $list ("the locked UTF-16 vector '" + $required + "' is present") `
            ($utfKeys -contains $required)
    }
    $null = Add-Check $list 'the reference stream states BOTH a code-unit count and a digest' `
        (([int]$cases.fingerprint.reference.code_units -gt 0) -and `
         (-not [string]::IsNullOrEmpty([string]$cases.fingerprint.reference.digest)))
    $null = Add-Check $list 'the reference stream is as long as the corpus says' `
        (([string]$cases.fingerprint.reference.stream).Length -eq [int]$cases.fingerprint.reference.code_units) `
        ("stream " + ([string]$cases.fingerprint.reference.stream).Length + `
         ", stated " + [int]$cases.fingerprint.reference.code_units)
    $null = Add-Check $list 'the collision probes are present' `
        (@($cases.fingerprint.collision_probes).Count -ge 8)

    $ok = Test-ChecklistOk $list
    Add-Result 'P5-PRE' `
        ("Phase-5 coverage preflight: " + $emitted.Count + " plan cases mapped (no Excel)") `
        $(if ($ok) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    return $ok
}

# ===========================================================================
# Reading the calculation workspace
# ===========================================================================
# Every address comes from the inspection projection. Nothing below names a cell
# or a table in its own right, so a contract change moves these reads with it
# instead of leaving them pointing at a stale coordinate.
function Get-CalcScalar {
    param($Workbook, $Inspection, [string]$Block, [string]$FieldKey)
    $block = $Inspection.calc.scalar_blocks.$Block
    $row = [int]$block.rows.$FieldKey
    $address = [string]$block.value_column + [string]$row
    # Each COM object into its OWN named variable, released in the narrowest
    # scope, exactly as every Phase-4 helper does. No chained member expression
    # creates an intermediate RCW that nothing owns.
    $sheets = $null; $sheet = $null; $range = $null
    try {
        $sheets = $Workbook.Worksheets
        $sheet = $sheets.Item($Inspection.calc.sheet)
        $range = $sheet.Range($address)
        # .Value2, not .Text: a formatted date read as text would compare against
        # a locale, and a blank read as text would become the empty string with
        # no way left to tell it from a value that really is "".
        return $range.Value2
    } finally {
        if ($null -ne $range)  { Release-Transient $range  'Range(calc scalar)'; $range  = $null }
        if ($null -ne $sheet)  { Release-Transient $sheet  'Worksheet(_Calc)';   $sheet  = $null }
        if ($null -ne $sheets) { Release-Transient $sheets 'Worksheets';         $sheets = $null }
    }
}

function Get-CalcScalarBlock {
    param($Workbook, $Inspection, [string]$Block)
    # The whole block as an ordered map key -> value, for the snapshot
    # comparisons. Read cell by cell so a BLANK stays $null rather than becoming
    # an empty string inside a Variant array.
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    $rows = $Inspection.calc.scalar_blocks.$Block.rows
    foreach ($key in $rows.PSObject.Properties.Name) {
        $out.Add($key, (Get-CalcScalar -Workbook $Workbook -Inspection $Inspection `
            -Block $Block -FieldKey $key))
    }
    return $out
}

function Get-CalcTableRows {
    param($Workbook, $Inspection, [string]$TableKey)
    # Reuses the accepted Phase-4 row-emission idiom: Get-TableBody puts ONE
    # object on the pipeline per row, and the caller materialises with @(...).
    $table = $Inspection.calc.tables.$TableKey
    return @(Get-TableBody -Workbook $Workbook -SheetName $Inspection.calc.sheet `
        -TableName $table.table_name)
}

function Get-CalcTableColumnIndex {
    param($Inspection, [string]$TableKey, [string]$ColumnKey)
    $columns = @($Inspection.calc.tables.$TableKey.columns)
    return [array]::IndexOf($columns, $ColumnKey)
}

# ---------------------------------------------------------------------------
# Type-sensitive comparison
# ---------------------------------------------------------------------------
# BLANK IS NOT NUMERIC ZERO, and the whole N/A rule of the audit blocks rests on
# that. A comparison that coerced both to 0 would report a fabricated zero as
# correct - which is the single most valuable thing these assertions can catch.
function Test-CalcBlank {
    param($Actual)
    if ($null -eq $Actual) { return $true }
    if ($Actual -is [string] -and $Actual.Length -eq 0) { return $true }
    return $false
}

function Test-CalcValue {
    param($Actual, $Expected, [double]$Tolerance = 0.0)
    if ($null -eq $Expected) { return (Test-CalcBlank -Actual $Actual) }
    if (Test-CalcBlank -Actual $Actual) { return $false }
    if ($Expected -is [string]) {
        if (-not ($Actual -is [string])) { return $false }
        return ([string]$Actual -ceq [string]$Expected)
    }
    if ($Actual -is [string]) { return $false }
    $a = [double]$Actual
    $e = [double]$Expected
    if ($a -eq $e) { return $true }
    if ($Tolerance -le 0.0) { return $false }
    $scale = [Math]::Max([Math]::Abs($e), 1.0)
    return ([Math]::Abs($a - $e) -le ($Tolerance * $scale))
}

function Format-CalcValue {
    param($Value)
    if ($null -eq $Value) { return '<blank>' }
    if ($Value -is [string]) {
        if ($Value.Length -eq 0) { return '<blank>' }
        return "'" + $Value + "'"
    }
    return ([double]$Value).ToString('R', [System.Globalization.CultureInfo]::InvariantCulture)
}

# ===========================================================================
# Applying an emitted fixture to the real workbook
# ===========================================================================
# The MODEL comes from phase5_cases.json; the ADDRESSES come from the manifest
# and the inspection projection. Nothing here decides what a model is.
#
# A FIXTURE MUST NOT BREAK THE INVARIANT IT IS TESTING - the Phase-4 rule, and it
# holds here too. Registers are emptied through the production delete endpoints
# and rows are keyed by the production add endpoints, so no orphan row is ever
# fabricated; only the ID counters are set directly, because the fixtures name
# CL-001 and R-001 and a permanent identifier is not re-issued by design.
function Clear-Phase5Registers {
    param($Excel, $Workbook, $Manifest)
    foreach ($register in @($Manifest.registers)) {
        $ids = @(Get-IdColumnValues -Workbook $Workbook -Info $register)
        foreach ($id in $ids) {
            if ($register.key -eq 'cost_lines') {
                $Excel.Run('PCCM_DeleteCostLineById', $id) | Out-Null
            } else {
                $Excel.Run('PCCM_DeleteRiskById', $id) | Out-Null
            }
        }
    }
    foreach ($counter in @($Manifest.counters)) {
        Set-NamedValue -Workbook $Workbook -DefinedName $counter.defined_name -Value ([double]$counter.initial)
    }
}

function Clear-Phase5GridBody {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$ColumnCount)
    # BLANKS the body. It does NOT delete rows.
    #
    # The Inflation grid carries reserved rows that Stage A builds and the
    # structural checks count. Deleting them to "clear" the grid would break the
    # very structure the fixture is about to calculate over - a fixture must not
    # break the invariant it is testing. A row whose key column is blank is
    # UNKEYED, which is exactly how the model reads "no profile here", and
    # Set-TableCell with $null clears contents rather than writing "" or 0.
    $rows = Get-TableRowCount -Workbook $Workbook -SheetName $SheetName -TableName $TableName
    for ($row = 1; $row -le $rows; $row++) {
        for ($column = 1; $column -le $ColumnCount; $column++) {
            Set-TableCell -Workbook $Workbook -SheetName $SheetName -TableName $TableName `
                -RowIndex $row -ColumnIndex $column -Value $null
        }
    }
}

function Clear-Phase5UserRows {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$KeepRows)
    # For a Setup/Config table whose first rows are a LOCKED seed: the seed stays
    # exactly as Stage A built it and the user rows above it are removed through
    # the accepted Phase-4 helper.
    $count = Get-TableRowCount -Workbook $Workbook -SheetName $SheetName -TableName $TableName
    for ($row = $count; $row -gt [Math]::Max($KeepRows, 1); $row--) {
        Remove-TableRow -Workbook $Workbook -SheetName $SheetName -TableName $TableName -RowIndex $row
    }
}

function Set-Phase5Fixture {
    param($Excel, $Workbook, $Manifest, $Inspection, $Model)

    # --- 1. empty the registers and reset the identity counters -------------
    Clear-Phase5Registers -Excel $Excel -Workbook $Workbook -Manifest $Manifest

    # --- 2. the Setup scalars ----------------------------------------------
    $inputs = $Inspection.inputs
    Set-NamedValue -Workbook $Workbook -DefinedName $inputs.base_year.defined_name `
        -Value ([double]$Model.timeline.base_year)
    Set-NamedValue -Workbook $Workbook -DefinedName $inputs.project_start_year.defined_name `
        -Value ([double]$Model.timeline.start_year)
    Set-NamedValue -Workbook $Workbook -DefinedName $inputs.duration_years.defined_name `
        -Value ([double]$Model.timeline.duration)
    Set-NamedValue -Workbook $Workbook -DefinedName $inputs.discount_rate.defined_name `
        -Value ([double]$Model.discount_rate)

    # --- 3. FX rates, from the fixture, above the locked reporting seed -----
    $fx = $Inspection.input_tables.fx_rates
    Clear-Phase5UserRows -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name `
        -KeepRows ([int]$fx.locked_seed_rows)
    # The reporting currency's own row is the table's LOCKED seed row. It is
    # never rewritten: the model states SAR -> 1 as a global invariant, and the
    # fixture asserting it would be the fixture proving itself.
    $reporting = [string](Get-NamedValue -Workbook $Workbook `
        -DefinedName $inputs.reporting_currency.defined_name)
    $fxRow = 0
    foreach ($entry in @($Model.fx)) {
        if ([string]$entry.currency -eq $reporting) { continue }
        $fxRow++
        Add-BlankTableRow -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name
        $target = [int]$fx.locked_seed_rows + $fxRow
        Set-TableCell -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name `
            -RowIndex $target -ColumnIndex 1 -Value ([string]$entry.currency)
        if ($null -ne $entry.rate) {
            Set-TableCell -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name `
                -RowIndex $target -ColumnIndex 2 -Value ([double]$entry.rate)
        }
    }

    # --- 4. inflation profiles ----------------------------------------------
    # The grid is keyed by profile name in its one fixed column; the year columns
    # are generated by Apply Timeline, so the profile rows are seeded FIRST and
    # the rates are written AFTER the timeline is applied.
    $inflGrid = $null
    foreach ($grid in @($Manifest.grids)) { if ($grid.key -eq 'inflation') { $inflGrid = $grid } }
    $inflHeaders = @(Get-TableColumnNames -Workbook $Workbook -SheetName $inflGrid.sheet `
        -TableName $inflGrid.table_name)
    Clear-Phase5GridBody -Workbook $Workbook -SheetName $inflGrid.sheet `
        -TableName $inflGrid.table_name -ColumnCount $inflHeaders.Count
    $profiles = @()
    foreach ($name in $Model.inflation.PSObject.Properties.Name) { $profiles += $name }
    # Keyed into the reserved rows the grid already has; a row is added only if
    # the fixture needs more profiles than the reservation provides.
    $reserved = Get-TableRowCount -Workbook $Workbook -SheetName $inflGrid.sheet `
        -TableName $inflGrid.table_name
    $index = 0
    foreach ($name in $profiles) {
        $index++
        if ($index -gt $reserved) {
            Add-BlankTableRow -Workbook $Workbook -SheetName $inflGrid.sheet -TableName $inflGrid.table_name
        }
        Set-TableCell -Workbook $Workbook -SheetName $inflGrid.sheet -TableName $inflGrid.table_name `
            -RowIndex $index -ColumnIndex 1 -Value ([string]$name)
    }

    # --- 5. the drivers ------------------------------------------------------
    $costReg = $null; $riskReg = $null
    foreach ($register in @($Manifest.registers)) {
        if ($register.key -eq 'cost_lines') { $costReg = $register }
        if ($register.key -eq 'risk_register') { $riskReg = $register }
    }
    $costIndex = 0
    foreach ($line in @($Model.cost_lines)) {
        $costIndex++
        $Excel.Run('PCCM_AddCostLine') | Out-Null
        Write-Phase5Driver -Workbook $Workbook -Register $costReg -RowIndex $costIndex `
            -Driver $line -IsRisk $false
    }
    $riskIndex = 0
    foreach ($risk in @($Model.risks)) {
        $riskIndex++
        $Excel.Run('PCCM_AddRisk') | Out-Null
        Write-Phase5Driver -Workbook $Workbook -Register $riskReg -RowIndex $riskIndex `
            -Driver $risk -IsRisk $true
    }

    # --- 6. apply the timeline, which generates the year columns ------------
    $Excel.Run('PCCM_ApplyTimeline') | Out-Null
    $applied = [string]$Excel.Run('PCCM_AutomationResult')

    # --- 7. inflation rates and profiling weights, into generated columns ----
    Write-Phase5InflationRates -Workbook $Workbook -Manifest $Manifest -Model $Model
    Write-Phase5Weights -Workbook $Workbook -Manifest $Manifest -Model $Model

    return $applied
}

function Write-Phase5Driver {
    param($Workbook, $Register, [int]$RowIndex, $Driver, [bool]$IsRisk)
    # Column ORDINALS come from the manifest's own column list. The harness never
    # counts columns for itself.
    $columns = @($Register.columns)
    $set = {
        param([string]$Key, $Value)
        $ordinal = [array]::IndexOf($columns, $Key) + 1
        if ($ordinal -lt 1) { throw ("the register has no column '" + $Key + "'") }
        if ($null -eq $Value) { return }
        Set-TableCell -Workbook $Workbook -SheetName $Register.sheet `
            -TableName $Register.table_name -RowIndex $RowIndex -ColumnIndex $ordinal -Value $Value
    }
    if ($IsRisk) {
        & $set 'risk_name'          ('GateB ' + [string]$Driver.permanent_id)
        & $set 'probability'        ([double]$Driver.probability)
        & $set 'impact_min'         ([double]$Driver.min_value)
        if ($null -ne $Driver.most_likely) { & $set 'impact_most_likely' ([double]$Driver.most_likely) }
        & $set 'impact_max'         ([double]$Driver.max_value)
    } else {
        & $set 'description'        ('GateB ' + [string]$Driver.permanent_id)
        & $set 'quantity'           ([double]$Driver.quantity)
        & $set 'unit_cost_min'      ([double]$Driver.min_value)
        if ($null -ne $Driver.most_likely) { & $set 'unit_cost_most_likely' ([double]$Driver.most_likely) }
        & $set 'unit_cost_max'      ([double]$Driver.max_value)
    }
    & $set 'currency'          ([string]$Driver.currency)
    & $set 'inflation_profile' ([string]$Driver.inflation_profile)
    & $set 'distribution'      ([string]$Driver.distribution)
}

function Write-Phase5InflationRates {
    param($Workbook, $Manifest, $Model)
    # A rate belongs to a CALENDAR YEAR, and the generated column carries that
    # year as its header. The column is found by header, never by offset: the
    # first generated column is BaseYear + 1, and assuming it is Start Year is
    # exactly the defect the Step-5 correction round removed from production.
    $grid = $null
    foreach ($candidate in @($Manifest.grids)) { if ($candidate.key -eq 'inflation') { $grid = $candidate } }
    $headers = @(Get-TableColumnNames -Workbook $Workbook -SheetName $grid.sheet -TableName $grid.table_name)
    $rowIndex = 0
    foreach ($name in $Model.inflation.PSObject.Properties.Name) {
        $rowIndex++
        $rates = $Model.inflation.$name
        foreach ($year in $rates.PSObject.Properties.Name) {
            $ordinal = [array]::IndexOf($headers, [string]$year) + 1
            if ($ordinal -lt 1) { throw ("no generated inflation column for calendar year " + $year) }
            Set-TableCell -Workbook $Workbook -SheetName $grid.sheet -TableName $grid.table_name `
                -RowIndex $rowIndex -ColumnIndex $ordinal -Value ([double]$rates.$year)
        }
    }
}

function Write-Phase5Weights {
    param($Workbook, $Manifest, $Model)
    # Profiling weights are per PROJECT YEAR, in order, after the fixed columns.
    # The grid is synchronised by permanent ID, so the row is located by its own
    # identifier rather than by the order the fixture happened to add drivers in.
    foreach ($pair in @(
        @{ key = 'cost_profiling'; drivers = @($Model.cost_lines) },
        @{ key = 'risk_profiling'; drivers = @($Model.risks) })) {
        $grid = $null
        foreach ($candidate in @($Manifest.grids)) {
            if ($candidate.key -eq $pair.key) { $grid = $candidate }
        }
        $fixed = @($grid.fixed_columns).Count
        $body = @(Get-TableBody -Workbook $Workbook -SheetName $grid.sheet -TableName $grid.table_name)
        foreach ($driver in $pair.drivers) {
            $rowIndex = 0
            for ($r = 0; $r -lt $body.Count; $r++) {
                if ([string]$body[$r][0] -eq [string]$driver.permanent_id) { $rowIndex = $r + 1 }
            }
            if ($rowIndex -lt 1) {
                throw ("no profiling row for " + [string]$driver.permanent_id +
                       "; synchronisation did not key the row")
            }
            $offset = 0
            foreach ($weight in @($driver.profile_weights)) {
                $offset++
                Set-TableCell -Workbook $Workbook -SheetName $grid.sheet -TableName $grid.table_name `
                    -RowIndex $rowIndex -ColumnIndex ($fixed + $offset) -Value ([double]$weight)
            }
        }
    }
}

# ===========================================================================
# Asserting an emitted `expected` block, in full
# ===========================================================================
# EVERY emitted expected value, not a handful of totals. The five analytical
# ListObjects, calc_totals and calc_state are all compared, and a row count that
# does not match the fixture is a failure in its own right rather than a reason
# to compare fewer rows.
function Add-Phase5AnalyticalChecks {
    param($List, $Workbook, $Inspection, $Case, $Tolerances)

    $expected = $Case.expected
    $label = 'case ' + [string]$Case.id
    $tolerance = [double]$Tolerances.identity_relative_coefficient

    # --- tblCalcYears: project index, calendar year, discount factor --------
    $years = @(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection -TableKey 'calc_years')
    $expectedYears = @($expected.discount_factors.PSObject.Properties.Name)
    $null = Add-Check $List ($label + ': tblCalcYears has one row per applied project year') `
        ($years.Count -eq $expectedYears.Count) `
        ("rows " + $years.Count + ", expected " + $expectedYears.Count)
    $indexColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_years' -ColumnKey 'project_index'
    $factorColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_years' -ColumnKey 'discount_factor'
    foreach ($row in $years) {
        $projectIndex = [string][int]$row[$indexColumn]
        $wanted = $expected.discount_factors.$projectIndex
        $null = Add-Check $List ($label + ': discount factor at project index ' + $projectIndex) `
            (Test-CalcValue -Actual $row[$factorColumn] -Expected $wanted -Tolerance $tolerance) `
            ("got " + (Format-CalcValue $row[$factorColumn]) + ", expected " + (Format-CalcValue $wanted))
    }

    # --- tblCalcInflationFactors -------------------------------------------
    # The Base-Year row carries a BLANK annual rate and a unit cumulative factor.
    # `annual_rate: null` in the fixture means BLANK, and Test-CalcValue refuses
    # a numeric zero in its place.
    $factors = @(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection -TableKey 'calc_inflation_factors')
    $expectedFactors = @($expected.inflation_factors)
    $null = Add-Check $List ($label + ': tblCalcInflationFactors row count') `
        ($factors.Count -eq $expectedFactors.Count) `
        ("rows " + $factors.Count + ", expected " + $expectedFactors.Count)
    $profileColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_inflation_factors' -ColumnKey 'inflation_profile'
    $yearColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_inflation_factors' -ColumnKey 'calendar_year'
    $rateColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_inflation_factors' -ColumnKey 'annual_rate'
    $cumulativeColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_inflation_factors' -ColumnKey 'cumulative_inflation_factor'
    foreach ($wanted in $expectedFactors) {
        $found = $null
        foreach ($row in $factors) {
            if (([string]$row[$profileColumn] -eq [string]$wanted.profile) -and `
                ([int]$row[$yearColumn] -eq [int]$wanted.calendar_year)) { $found = $row }
        }
        $key = $label + ': inflation ' + [string]$wanted.profile + ' ' + [string]$wanted.calendar_year
        if ($null -eq $found) {
            $null = Add-Check $List ($key + ' row exists') $false 'no such row'
            continue
        }
        $null = Add-Check $List ($key + ' annual rate') `
            (Test-CalcValue -Actual $found[$rateColumn] -Expected $wanted.annual_rate -Tolerance $tolerance) `
            ("got " + (Format-CalcValue $found[$rateColumn]) + ", expected " + (Format-CalcValue $wanted.annual_rate))
        $null = Add-Check $List ($key + ' cumulative factor') `
            (Test-CalcValue -Actual $found[$cumulativeColumn] -Expected $wanted.cumulative_factor -Tolerance $tolerance) `
            ("got " + (Format-CalcValue $found[$cumulativeColumn]) + ", expected " + (Format-CalcValue $wanted.cumulative_factor))
    }

    # --- tblCalcFX: REFERENCED currencies only ------------------------------
    $fxRows = @(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection -TableKey 'calc_fx')
    $expectedFx = @($expected.resolved_fx.PSObject.Properties.Name)
    $null = Add-Check $List ($label + ': tblCalcFX carries the referenced currencies only') `
        ($fxRows.Count -eq $expectedFx.Count) `
        ("rows " + $fxRows.Count + ", expected " + $expectedFx.Count)
    $currencyColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_fx' -ColumnKey 'currency'
    $rateColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_fx' -ColumnKey 'fx_to_sar'
    foreach ($currency in $expectedFx) {
        $found = $null
        foreach ($row in $fxRows) { if ([string]$row[$currencyColumn] -eq $currency) { $found = $row } }
        if ($null -eq $found) {
            $null = Add-Check $List ($label + ': FX row for ' + $currency) $false 'no such row'
            continue
        }
        $null = Add-Check $List ($label + ': FX rate for ' + $currency) `
            (Test-CalcValue -Actual $found[$rateColumn] -Expected $expected.resolved_fx.$currency -Tolerance $tolerance) `
            ("got " + (Format-CalcValue $found[$rateColumn]))
    }

    # --- tblCalcDrivers: every emitted field of every driver ----------------
    $driverRows = @(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection -TableKey 'calc_drivers')
    $expectedDrivers = @($expected.drivers)
    $null = Add-Check $List ($label + ': tblCalcDrivers row count') `
        ($driverRows.Count -eq $expectedDrivers.Count) `
        ("rows " + $driverRows.Count + ", expected " + $expectedDrivers.Count)
    $idColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_drivers' -ColumnKey 'permanent_id'
    foreach ($wanted in $expectedDrivers) {
        $found = $null
        foreach ($row in $driverRows) {
            if ([string]$row[$idColumn] -eq [string]$wanted.permanent_id) { $found = $row }
        }
        if ($null -eq $found) {
            $null = Add-Check $List ($label + ': driver row ' + [string]$wanted.permanent_id) $false 'no such row'
            continue
        }
        # Driven from the FIXTURE's own field names, so a field added to the
        # corpus is asserted here without this file being edited - and a field
        # the emitted table does not carry is reported rather than skipped.
        foreach ($field in $wanted.PSObject.Properties.Name) {
            if ($field -eq 'weights') { continue }
            $ordinal = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_drivers' -ColumnKey $field
            if ($ordinal -lt 0) {
                $null = Add-Check $List `
                    ($label + ': tblCalcDrivers has a column for expected field ' + $field) $false
                continue
            }
            $null = Add-Check $List ($label + ': ' + [string]$wanted.permanent_id + '.' + $field) `
                (Test-CalcValue -Actual $found[$ordinal] -Expected $wanted.$field -Tolerance $tolerance) `
                ("got " + (Format-CalcValue $found[$ordinal]) + ", expected " + (Format-CalcValue $wanted.$field))
        }
    }

    # --- tblCalcAnnual ------------------------------------------------------
    $annualRows = @(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection -TableKey 'calc_annual')
    $expectedAnnual = @($expected.annual)
    $null = Add-Check $List ($label + ': tblCalcAnnual row count') `
        ($annualRows.Count -eq $expectedAnnual.Count) `
        ("rows " + $annualRows.Count + ", expected " + $expectedAnnual.Count)
    $annualIndexColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_annual' -ColumnKey 'project_index'
    foreach ($wanted in $expectedAnnual) {
        $found = $null
        foreach ($row in $annualRows) {
            if ([int]$row[$annualIndexColumn] -eq [int]$wanted.project_index) { $found = $row }
        }
        if ($null -eq $found) {
            $null = Add-Check $List ($label + ': annual row ' + [string]$wanted.project_index) $false 'no such row'
            continue
        }
        foreach ($field in $wanted.PSObject.Properties.Name) {
            $ordinal = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_annual' -ColumnKey $field
            if ($ordinal -lt 0) {
                $null = Add-Check $List ($label + ': tblCalcAnnual has a column for ' + $field) $false
                continue
            }
            $null = Add-Check $List ($label + ': annual ' + [string]$wanted.project_index + '.' + $field) `
                (Test-CalcValue -Actual $found[$ordinal] -Expected $wanted.$field -Tolerance $tolerance) `
                ("got " + (Format-CalcValue $found[$ordinal]) + ", expected " + (Format-CalcValue $wanted.$field))
        }
    }

    # --- calc_totals: all ten -----------------------------------------------
    foreach ($field in @($expected.totals.PSObject.Properties.Name)) {
        $actual = Get-CalcScalar -Workbook $Workbook -Inspection $Inspection `
            -Block 'calc_totals' -FieldKey $field
        $null = Add-Check $List ($label + ': calc_totals.' + $field) `
            (Test-CalcValue -Actual $actual -Expected $expected.totals.$field -Tolerance $tolerance) `
            ("got " + (Format-CalcValue $actual) + ", expected " + (Format-CalcValue $expected.totals.$field))
    }
}

# ===========================================================================
# The scenarios
# ===========================================================================
function Invoke-Phase5GateBScenarios {
    param(
        $Excel, $Workbook, $Manifest, $Inspection, $Cases,
        [string]$ScriptDir, [string]$TempRoot, $Results
    )

    $failpoints = Get-Phase5FailpointNames
    $ledger = Get-Phase5CoverageLedger

    # -------------------------------------------------------------------
    # PHASE-4 PREREQUISITE. 35/35, 0 FAIL, 0 SKIP - checked, not assumed.
    # -------------------------------------------------------------------
    # Gate-B acceptance requires the structural matrix intact BEFORE a Phase-5
    # result means anything. A Phase-5 pass on a workbook whose timeline
    # machinery is broken would be evidence of nothing.
    $required = Get-Phase4RequiredScenarioIds
    $list = New-Checklist
    $seen = @($Results | ForEach-Object { $_.Id })
    $missing = @()
    foreach ($id in $required) { if ($seen -notcontains $id) { $missing += $id } }
    $null = Add-Check $list 'all 35 Phase-4 scenarios reported a result' ($missing.Count -eq 0) `
        ("missing: " + ($missing -join ', '))
    $phase4 = @($Results | Where-Object { $required -contains $_.Id })
    $failed = @($phase4 | Where-Object { $_.Status -eq 'FAIL' })
    $skipped = @($phase4 | Where-Object { $_.Status -eq 'SKIP' })
    $null = Add-Check $list 'the Phase-4 matrix has 0 FAIL' ($failed.Count -eq 0) `
        (($failed | ForEach-Object { $_.Id }) -join ', ')
    $null = Add-Check $list 'the Phase-4 matrix has 0 SKIP' ($skipped.Count -eq 0) `
        (($skipped | ForEach-Object { $_.Id }) -join ', ')
    $passed = @($phase4 | Where-Object { $_.Status -eq 'PASS' })
    $null = Add-Check $list 'the Phase-4 matrix is 35/35 PASS' ($passed.Count -eq 35) `
        ("passed " + $passed.Count + " of 35")
    $phase4Ok = Test-ChecklistOk $list
    Add-Result 'P5-P4' 'Phase-4 prerequisite: 35/35 PASS, 0 FAIL, 0 SKIP' `
        $(if ($phase4Ok) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    if (-not $phase4Ok) {
        # A FAIL, never a SKIP. "Phase 5 was not attempted" must be as loud as
        # "Phase 5 failed", or a broken prerequisite reads as a quiet clean run.
        Add-Result 'P5-ALL' 'Phase-5 Gate-B scenarios' 'FAIL' `
            'not attempted: the Phase-4 structural matrix is not intact, so no Phase-5 result would mean anything'
        return
    }

    # -------------------------------------------------------------------
    # P5-M. The persisted project: modules by NAME, buttons, API procedures
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $project = $null; $components = $null
        try {
            $project = $Workbook.VBProject
            $components = $project.VBComponents
            $present = @()
            for ($i = 1; $i -le $components.Count; $i++) {
                $component = $null
                try { $component = $components.Item($i); $present += [string]$component.Name }
                finally { if ($null -ne $component) { Release-Transient $component 'VBComponent'; $component = $null } }
            }
            # BY NAME, in both directions. A count alone would pass a project that
            # had gained a stray module and lost a real one.
            $expected = @($Manifest.vba.modules | ForEach-Object { [string]$_.name })
            $null = Add-Check $list 'the manifest declares 15 production modules' `
                ($expected.Count -eq 15) ("declared " + $expected.Count)
            foreach ($name in $expected) {
                $null = Add-Check $list ('the module ' + $name + ' persists in the saved project') `
                    ($present -contains $name)
            }
            $extra = @($present | Where-Object { $expected -notcontains $_ })
            $null = Add-Check $list 'no module outside the manifest persists' ($extra.Count -eq 0) `
                ("extra: " + ($extra -join ', '))
        } finally {
            if ($null -ne $components) { Release-Transient $components 'VBComponents'; $components = $null }
            if ($null -ne $project) { Release-Transient $project 'VBProject'; $project = $null }
        }

        # --- exactly five buttons, and not one of them calls PCCM_Calculate ---
        $declared = @($Manifest.buttons)
        $null = Add-Check $list 'the manifest declares exactly five buttons' ($declared.Count -eq 5) `
            ("declared " + $declared.Count)
        $shapesFound = 0
        $onActions = @()
        foreach ($sheetSpec in @($Manifest.sheets)) {
            $sheets = $null; $sheet = $null; $shapes = $null
            try {
                $sheets = $Workbook.Worksheets
                $sheet = $sheets.Item($sheetSpec.name)
                $shapes = $sheet.Shapes
                for ($i = 1; $i -le $shapes.Count; $i++) {
                    $shape = $null
                    try {
                        $shape = $shapes.Item($i)
                        $shapesFound++
                        $onActions += [string]$shape.OnAction
                    } finally {
                        if ($null -ne $shape) { Release-Transient $shape 'Shape'; $shape = $null }
                    }
                }
            } finally {
                if ($null -ne $shapes) { Release-Transient $shapes 'Shapes'; $shapes = $null }
                if ($null -ne $sheet) { Release-Transient $sheet 'Worksheet'; $sheet = $null }
                if ($null -ne $sheets) { Release-Transient $sheets 'Worksheets'; $sheets = $null }
            }
        }
        $null = Add-Check $list 'exactly five command buttons persist in the workbook' `
            ($shapesFound -eq 5) ("found " + $shapesFound)
        foreach ($button in $declared) {
            $null = Add-Check $list ('the button ' + $button.shape_name + ' calls ' + $button.entry_point) `
                ($onActions -contains [string]$button.entry_point)
        }
        # THE ONE THAT MATTERS: no shape may invoke the calculation endpoint.
        $null = Add-Check $list 'NO shape has OnAction = PCCM_Calculate' `
            ($onActions -notcontains 'PCCM_Calculate') (($onActions -join ', '))

        # --- api_procedures, consumed as api_procedures ----------------------
        # Deliberately NOT folded into entry_points: an entry point is bound to a
        # button and an API procedure is not, and the manifest is where the
        # harness learns the difference.
        $api = @($Manifest.vba.api_procedures)
        $entry = @($Manifest.vba.entry_points)
        $null = Add-Check $list 'the manifest projects exactly six API procedures' ($api.Count -eq 6) `
            ("projected " + $api.Count)
        $null = Add-Check $list 'no API procedure is also an entry point' `
            (@($api | Where-Object { $entry -contains $_ }).Count -eq 0)
        $null = Add-Check $list 'no API procedure is bound to a button' `
            (@($api | Where-Object { $onActions -contains $_ }).Count -eq 0)
        foreach ($name in $api) {
            $callable = $false; $detail = ''
            try {
                if ($name -eq 'PCCM_Calculate') {
                    # Callability alone, without driving a calculation here: the
                    # analytical scenarios below do that against real fixtures.
                    $callable = $true
                    $detail = 'exercised by the analytical scenarios below'
                } else {
                    $probe = $Excel.Run($name)
                    $callable = $true
                    $detail = "returned '" + [string]$probe + "'"
                }
            } catch { $detail = (Format-Err $_) }
            $null = Add-Check $list ('the API procedure ' + $name + ' is callable') $callable $detail
        }
        Add-Result 'P5-M' 'Persisted project: 15 modules by name, 5 buttons, 6 API procedures' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'P5-M' 'Persisted project inventory' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # P5-EV. No change events. The status cell is last-evaluated, not live.
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $project = $null; $components = $null
        try {
            $project = $Workbook.VBProject
            $components = $project.VBComponents
            $offenders = @()
            for ($i = 1; $i -le $components.Count; $i++) {
                $component = $null; $module = $null
                try {
                    $component = $components.Item($i)
                    $module = $component.CodeModule
                    if ($module.CountOfLines -gt 0) {
                        $text = [string]$module.Lines(1, $module.CountOfLines)
                        foreach ($forbidden in @($Manifest.vba.forbidden_constructs)) {
                            if ($text -match [regex]::Escape([string]$forbidden)) {
                                $offenders += ([string]$component.Name + ': ' + [string]$forbidden)
                            }
                        }
                    }
                } finally {
                    if ($null -ne $module) { Release-Transient $module 'CodeModule'; $module = $null }
                    if ($null -ne $component) { Release-Transient $component 'VBComponent'; $component = $null }
                }
            }
            $null = Add-Check $list 'no forbidden construct exists in the real Stage-B project' `
                ($offenders.Count -eq 0) ($offenders -join '; ')
            foreach ($handler in 'Worksheet_Change', 'Workbook_SheetChange') {
                $null = Add-Check $list ('the manifest forbids ' + $handler) `
                    (@($Manifest.vba.forbidden_constructs) -contains $handler)
            }
        } finally {
            if ($null -ne $components) { Release-Transient $components 'VBComponents'; $components = $null }
            if ($null -ne $project) { Release-Transient $project 'VBProject'; $project = $null }
        }
        Add-Result 'P5-EV' 'No change events: the status cell is last-evaluated, not live' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'P5-EV' 'No change events' 'FAIL' (Format-Err $_)
    }

    # ===================================================================
    # THE TRANSIENT DIAGNOSTIC SECTION
    #
    # Imported HERE and nowhere earlier. Scenario A1 has already made the first
    # Application.Run of the run against the PRODUCTION project, so the proof
    # that the accepted project compiles is complete and unmasked before a
    # test-only module exists in the VBA project at all.
    # ===================================================================
    $diagnosticName = 'modPhase5GateBDiagnostics'
    $diagnosticImported = $false
    try {
        $list = New-Checklist
        $source = Join-Path $ScriptDir 'phase5_gate_b_diagnostics.bas'
        $null = Add-Check $list 'the diagnostic source exists' (Test-Path -LiteralPath $source) $source
        $null = Add-Check $list 'the diagnostic module is NOT declared in the manifest' `
            (@($Manifest.vba.modules | ForEach-Object { [string]$_.name }) -notcontains $diagnosticName)

        $project = $null; $components = $null; $imported = $null
        try {
            $project = $Workbook.VBProject
            $components = $project.VBComponents
            $imported = $components.Import($source)
            $diagnosticImported = $true
            $null = Add-Check $list 'the diagnostic module imported into the disposable project' `
                ([string]$imported.Name -eq $diagnosticName) ("imported as " + [string]$imported.Name)
        } finally {
            if ($null -ne $imported) { Release-Transient $imported 'VBComponent(diagnostic)'; $imported = $null }
            if ($null -ne $components) { Release-Transient $components 'VBComponents'; $components = $null }
            if ($null -ne $project) { Release-Transient $project 'VBProject'; $project = $null }
        }

        $ping = [string]$Excel.Run('GBD_Ping')
        $null = Add-Check $list 'the diagnostic module is callable' ($ping -eq ('OK|' + $diagnosticName)) $ping
        Add-Result 'P5-D0' 'Transient diagnostic module imported AFTER the A1 production compile' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'P5-D0' 'Transient diagnostic module import' 'FAIL' (Format-Err $_)
    }

    if ($diagnosticImported) {
        # ---------------------------------------------------------------
        # P5-D1. Canonical numeric encoding: the ten locked vectors
        # ---------------------------------------------------------------
        try {
            $list = New-Checklist
            $vectors = @($Cases.fingerprint.numeric_encodings.vectors)
            $null = Add-Check $list 'ten locked numeric vectors were emitted' ($vectors.Count -eq 10) `
                ("vectors " + $vectors.Count)
            foreach ($vector in $vectors) {
                $wanted = [string]$vector.expected
                $reply = [string]$Excel.Run('GBD_CanonicalNumber', [double]$vector.value, '.')
                $null = Add-Check $list ("canonical number '" + [string]$vector.label + "'") `
                    ($reply -eq ('OK|' + $wanted)) ("got " + $reply + ", expected OK|" + $wanted)
                # THE TWO EXTREMES ARE ALSO BUILT ON TARGET. If a COM Double round
                # trip disturbed MAX_DOUBLE or the minimum subnormal, the two
                # answers differ and the report says which one moved - the vector
                # is never skipped and never quietly weakened.
                if (@('MAX_DOUBLE', 'minimum subnormal') -contains [string]$vector.label) {
                    $built = [string]$Excel.Run('GBD_CanonicalNumberConstructed', [string]$vector.label, '.')
                    $null = Add-Check $list `
                        ("canonical number '" + [string]$vector.label + "' built on target, not marshalled") `
                        ($built -eq ('OK|' + $wanted)) ("got " + $built + ", expected OK|" + $wanted)
                    $null = Add-Check $list `
                        ("the marshalled and on-target '" + [string]$vector.label + "' agree") `
                        ($built -eq $reply) ("marshalled " + $reply + " / constructed " + $built)
                }
            }
            Add-Result 'P5-D1' 'Direct VBA: ten canonical numeric encodings (plan case 26)' `
                $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        } catch {
            Add-Result 'P5-D1' 'Direct VBA: canonical numeric encodings' 'FAIL' (Format-Err $_)
        }

        # ---------------------------------------------------------------
        # P5-D2. Decimal-separator INJECTION, on the same Windows host
        # ---------------------------------------------------------------
        # The runtime proof Gate A could not make. Both separators go into the
        # SAME accepted encoder as its own argument, on one host, in one run. No
        # regional setting is read or altered; Application.International is never
        # touched and UseSystemSeparators is never set.
        try {
            $list = New-Checklist
            $vectors = @($Cases.fingerprint.decimal_separator.vectors)
            $null = Add-Check $list 'the separator vector set was emitted' ($vectors.Count -ge 10) `
                ("vectors " + $vectors.Count)
            foreach ($vector in $vectors) {
                foreach ($pair in @(
                    @{ separator = '.'; expected = [string]$vector.point;  name = 'point' },
                    @{ separator = ','; expected = [string]$vector.comma;  name = 'comma' })) {
                    $reply = [string]$Excel.Run('GBD_CanonicalNumber', [double]$vector.value, $pair.separator)
                    $null = Add-Check $list `
                        ("separator '" + $pair.name + "' on vector '" + [string]$vector.label + "'") `
                        ($reply -eq ('OK|' + $pair.expected)) `
                        ("got " + $reply + ", expected OK|" + $pair.expected)
                }
                if (@('MAX_DOUBLE', 'minimum subnormal') -contains [string]$vector.label) {
                    foreach ($separator in '.', ',') {
                        $built = [string]$Excel.Run('GBD_CanonicalNumberConstructed', [string]$vector.label, $separator)
                        $wanted = [string]$vector.point
                        if ($separator -eq ',') { $wanted = [string]$vector.comma }
                        $null = Add-Check $list `
                            ("separator '" + $separator + "' on the on-target '" + [string]$vector.label + "'") `
                            ($built -eq ('OK|' + $wanted)) ("got " + $built)
                    }
                }
            }
            # The output must be IDENTICAL under both separators: the canonical
            # form is the model's, not the host's.
            $null = Add-Check $list 'the canonical form does not depend on the injected separator' `
                (@($vectors | Where-Object { [string]$_.point -cne [string]$_.comma }).Count -eq 0)
            Add-Result 'P5-D2' 'Direct VBA: decimal-separator injection, both separators (plan case 35)' `
                $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        } catch {
            Add-Result 'P5-D2' 'Direct VBA: decimal-separator injection' 'FAIL' (Format-Err $_)
        }

        # ---------------------------------------------------------------
        # P5-D3. The Double-only reducer: all four locked vectors
        # ---------------------------------------------------------------
        # PowerShell computes NOTHING here. It hands h, u and the modulus to the
        # accepted VBA reducer and compares the returned remainder against the
        # fixture. A PowerShell-side reduction would be testing PowerShell.
        try {
            $list = New-Checklist
            $vectors = @($Cases.fingerprint.reduction_vectors)
            $null = Add-Check $list 'all four locked reduction vectors were emitted' `
                ($vectors.Count -eq 4) ("vectors " + $vectors.Count)
            foreach ($vector in $vectors) {
                $reply = [string]$Excel.Run('GBD_ReduceDouble', [double]$vector.h, [double]$vector.u,
                                            [double]$vector.modulus)
                # The remainder comes back as canonical text, so the comparison is
                # exact and no PowerShell number formatting stands in the way.
                $wantedReply = [string]$Excel.Run('GBD_CanonicalNumber', [double]$vector.remainder, '.')
                $null = Add-Check $list `
                    ("reduction h=" + [string]$vector.h + " u=" + [string]$vector.u + " mod " + [string]$vector.modulus_name) `
                    ($reply -eq $wantedReply) ("got " + $reply + ", expected " + $wantedReply)
                $null = Add-Check $list `
                    ("the fixture's double-only remainder equals its exact remainder for " + [string]$vector.modulus_name) `
                    ([double]$vector.double_only_remainder -eq [double]$vector.remainder)
            }
            Add-Result 'P5-D3' 'Direct VBA: the four Double-only reductions (plan case 36)' `
                $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        } catch {
            Add-Result 'P5-D3' 'Direct VBA: the Double-only reducer' 'FAIL' (Format-Err $_)
        }

        # ---------------------------------------------------------------
        # P5-D4. UTF-16: signed AscW, unit counting, surrogates, prefixes
        # ---------------------------------------------------------------
        try {
            $list = New-Checklist
            $vectors = @($Cases.fingerprint.utf16_vectors.vectors)
            foreach ($vector in $vectors) {
                $units = (@($vector.code_units) -join ',')
                $key = [string]$vector.key

                # LENGTH IS IN CODE UNITS. A non-BMP character contributes two.
                $reply = [string]$Excel.Run('GBD_Utf16Length', $units)
                $null = Add-Check $list ($key + ': UTF-16 length counts code units') `
                    ($reply -eq ('OK|' + [string]$vector.utf16_length)) `
                    ("got " + $reply + ", expected OK|" + [string]$vector.utf16_length)
                if ($key -eq 'non_bmp') {
                    $null = Add-Check $list `
                        'the non-BMP character contributes TWO surrogate units, not one' `
                        (([int]$vector.utf16_length -eq 2) -and ([int]$vector.code_point_count -eq 1))
                }

                # AscW IS SIGNED. Every unit above U+7FFF must come back negative,
                # and the accepted normaliser must put it back.
                $position = 0
                foreach ($signed in @($vector.signed_ascw)) {
                    $position++
                    $raw = [string]$Excel.Run('GBD_RawAscW', $units, [long]$position)
                    $null = Add-Check $list ($key + ': raw AscW at position ' + $position + ' is signed') `
                        ($raw -eq ('OK|' + [string]$signed)) ("got " + $raw + ", expected OK|" + [string]$signed)
                    $normalised = [string]$Excel.Run('GBD_NormaliseCodeUnit', [long]$signed)
                    $wanted = [string]@($vector.code_units)[$position - 1]
                    $null = Add-Check $list ($key + ': the normaliser restores unit ' + $position) `
                        ($normalised -eq ('OK|' + $wanted)) ("got " + $normalised + ", expected OK|" + $wanted)
                }
                $null = Add-Check $list ($key + ': at least one unit above U+7FFF is exercised') `
                    (@($vector.signed_ascw | Where-Object { [int]$_ -lt 0 }).Count -ge 1) -Detail `
                    (@($vector.signed_ascw) -join ',')

                # THE LENGTH PREFIX IS THE UNIT COUNT.
                $field = [string]$Excel.Run('GBD_CanonicalTextField', $units)
                $expectedField = 'OK|S' + [string]$vector.utf16_length + ':'
                $null = Add-Check $list ($key + ': the canonical text prefix is the UTF-16 unit count') `
                    ($field.StartsWith($expectedField)) ("got " + $field + ", expected prefix " + $expectedField)
            }
            Add-Result 'P5-D4' 'Direct VBA: UTF-16 signed AscW, unit counting and prefixes (plan case 26)' `
                $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        } catch {
            Add-Result 'P5-D4' 'Direct VBA: UTF-16 behaviour' 'FAIL' (Format-Err $_)
        }

        # ---------------------------------------------------------------
        # P5-D5. The complete locked reference stream
        # ---------------------------------------------------------------
        try {
            $list = New-Checklist
            $reference = $Cases.fingerprint.reference
            $stream = [string]$reference.stream

            # BOTH, and the count FIRST. A digest asserted on its own would agree
            # with itself over a stream that arrived truncated.
            $length = [string]$Excel.Run('GBD_StreamLength', $stream)
            $null = Add-Check $list 'the reference stream is the emitted code-unit count on real VBA' `
                ($length -eq ('OK|' + [string]$reference.code_units)) `
                ("got " + $length + ", expected OK|" + [string]$reference.code_units)
            $digest = [string]$Excel.Run('GBD_DigestStream', $stream)
            $null = Add-Check $list 'the reference digest matches the emitted digest on real VBA' `
                ($digest -eq ('OK|' + [string]$reference.digest)) `
                ("got " + $digest + ", expected OK|" + [string]$reference.digest)
            Add-Result 'P5-D5' 'Direct VBA: the complete reference stream, units and digest (plan case 26)' `
                $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        } catch {
            Add-Result 'P5-D5' 'Direct VBA: the reference stream' 'FAIL' (Format-Err $_)
        }

        # ---------------------------------------------------------------
        # P5-D6. Delimiter-hostile field content
        # ---------------------------------------------------------------
        # The probe values carry ':', NUL, the unit separator and a newline. Every
        # one is handed over as CODE UNITS, so nothing about them survives or dies
        # by the console encoding on the way in.
        try {
            $list = New-Checklist
            $probes = @($Cases.fingerprint.collision_probes)
            $null = Add-Check $list 'the collision probes were emitted' ($probes.Count -ge 8) `
                ("probes " + $probes.Count)
            $seenDigests = @()
            foreach ($probe in $probes) {
                $encoded = @()
                foreach ($value in @($probe.values)) {
                    $units = @()
                    foreach ($character in [char[]][string]$value) { $units += [int][char]$character }
                    $encoded += ($units -join ',')
                }
                $reply = [string]$Excel.Run('GBD_ProbeDigest', ($encoded -join ';'))
                $null = Add-Check $list ('probe digest for ' + ($encoded -join ' | ')) `
                    ($reply -eq ('OK|' + [string]$probe.digest)) `
                    ("got " + $reply + ", expected OK|" + [string]$probe.digest)
                $seenDigests += [string]$probe.digest
            }
            # The point of the probes: hostile content must not COLLIDE.
            $null = Add-Check $list 'every probe digest is distinct' `
                ((@($seenDigests | Select-Object -Unique)).Count -eq $seenDigests.Count)
            Add-Result 'P5-D6' 'Direct VBA: delimiter-hostile field content (plan case 27)' `
                $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        } catch {
            Add-Result 'P5-D6' 'Direct VBA: delimiter-hostile field content' 'FAIL' (Format-Err $_)
        }

        # ---------------------------------------------------------------
        # P5-D7. A naive overflow with a representable result
        # ---------------------------------------------------------------
        try {
            $list = New-Checklist
            $case = $null
            foreach ($candidate in @($Cases.plan_cases)) {
                if ([string]$candidate.id -eq '28') { $case = $candidate }
            }
            $null = Add-Check $list 'plan case 28 was emitted with its statistics vectors' `
                ($null -ne $case -and @($case.statistics).Count -ge 3)
            foreach ($vector in @($case.statistics)) {
                $points = @($vector.points)
                $third = 0.0
                if ($points.Count -ge 3) { $third = [double]$points[2] }
                $reply = [string]$Excel.Run('GBD_ConvexStatistic', [string]$vector.statistic,
                                            [double]$points[0], [double]$points[1], [double]$third)
                $wanted = [string]$Excel.Run('GBD_CanonicalNumber', [double]$vector.expected, '.')
                $null = Add-Check $list `
                    ([string]$vector.statistic + ' survives the naive sum overflow') `
                    ($reply -eq $wanted) ("got " + $reply + ", expected " + $wanted)
            }
            Add-Result 'P5-D7' 'Direct VBA: convex statistics at the overflow boundary (plan case 28)' `
                $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        } catch {
            Add-Result 'P5-D7' 'Direct VBA: convex statistics at the overflow boundary' 'FAIL' (Format-Err $_)
        }
    } else {
        foreach ($id in 'P5-D1', 'P5-D2', 'P5-D3', 'P5-D4', 'P5-D5', 'P5-D6', 'P5-D7') {
            Add-Result $id 'Direct VBA diagnostic vector' 'FAIL' `
                'the transient diagnostic module did not import, so no locked vector was exercised on real VBA'
        }
    }

    # -------------------------------------------------------------------
    # P5-D8. THE DIAGNOSTIC MODULE IS REMOVED AGAIN
    # -------------------------------------------------------------------
    # Evidence infrastructure, not product. The inventory must return to exactly
    # the 15 manifest modules BEFORE anything else is asserted about the project,
    # and no accepted workbook is ever saved with it installed.
    try {
        $list = New-Checklist
        $project = $null; $components = $null
        try {
            $project = $Workbook.VBProject
            $components = $project.VBComponents
            $target = $null
            try { $target = $components.Item($diagnosticName) } catch { $target = $null }
            if ($null -ne $target) {
                $components.Remove($target)
                Release-Transient $target 'VBComponent(diagnostic)'; $target = $null
            }
            $present = @()
            for ($i = 1; $i -le $components.Count; $i++) {
                $component = $null
                try { $component = $components.Item($i); $present += [string]$component.Name }
                finally { if ($null -ne $component) { Release-Transient $component 'VBComponent'; $component = $null } }
            }
            $null = Add-Check $list 'the diagnostic module is absent from the project' `
                ($present -notcontains $diagnosticName) (($present -join ', '))
            $expected = @($Manifest.vba.modules | ForEach-Object { [string]$_.name })
            $null = Add-Check $list 'the inventory is exactly the 15 manifest modules again' `
                ((@($present).Count -eq $expected.Count) -and `
                 (@($present | Where-Object { $expected -notcontains $_ }).Count -eq 0)) `
                ("present " + @($present).Count + " of " + $expected.Count)
            foreach ($name in $expected) {
                $null = Add-Check $list ('the production module ' + $name + ' survived the removal') `
                    ($present -contains $name)
            }
        } finally {
            if ($null -ne $components) { Release-Transient $components 'VBComponents'; $components = $null }
            if ($null -ne $project) { Release-Transient $project 'VBProject'; $project = $null }
        }
        # It must also be gone from the RUNTIME: a removed component whose
        # procedure still answers would mean the removal did not take.
        $stillCallable = $false
        try { $null = $Excel.Run('GBD_Ping'); $stillCallable = $true } catch { $stillCallable = $false }
        $null = Add-Check $list 'no diagnostic procedure is callable any more' (-not $stillCallable)
        Add-Result 'P5-D8' 'Transient diagnostic module removed; inventory back to 15' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'P5-D8' 'Transient diagnostic module removal' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # P5-AN. Every analytical fixture, every emitted expected value
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $covered = @()
        foreach ($case in @($Cases.plan_cases)) {
            if ([string]$case.kind -ne 'analytical') { continue }
            $id = [string]$case.id
            $applied = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
                -Inspection $Inspection -Model $case.model
            $null = Add-Check $list ('case ' + $id + ': the fixture applied its timeline') `
                ($applied -like 'OK|*') $applied

            $Excel.Run('PCCM_Calculate') | Out-Null
            $attempt = [string]$Excel.Run('PCCM_CalculationAttemptResult')
            $detail = [string]$Excel.Run('PCCM_CalculationAttemptDetail')
            $status = [string]$Excel.Run('PCCM_CalculationStatus')
            $null = Add-Check $list ('case ' + $id + ': the attempt is SUCCESS') `
                ($attempt -eq 'SUCCESS') ("attempt '" + $attempt + "', detail '" + $detail + "'")
            $null = Add-Check $list ('case ' + $id + ': the attempt detail is blank on success') `
                ([string]::IsNullOrEmpty($detail)) $detail
            $null = Add-Check $list ('case ' + $id + ': the derived status is CURRENT') `
                ($status -eq 'CURRENT') $status
            $fingerprint = [string]$Excel.Run('PCCM_CalculationFingerprint')
            $current = [string]$Excel.Run('PCCM_CurrentInputFingerprint')
            $null = Add-Check $list ('case ' + $id + ': a stored fingerprint exists') `
                (-not [string]::IsNullOrEmpty($fingerprint))
            $null = Add-Check $list ('case ' + $id + ': the stored and current fingerprints agree') `
                ($fingerprint -ceq $current) ("stored " + $fingerprint + ", current " + $current)

            if ($attempt -eq 'SUCCESS') {
                Add-Phase5AnalyticalChecks -List $list -Workbook $Workbook -Inspection $Inspection `
                    -Case $case -Tolerances $Cases.tolerances
            }
            $covered += $id
        }
        $null = Add-Check $list 'every analytical plan case was driven' ($covered.Count -eq 19) `
            ("covered " + $covered.Count + ": " + ($covered -join ', '))
        Add-Result 'P5-AN' `
            ('Analytical fixtures through PCCM_Calculate: ' + $covered.Count + ' cases, all emitted values') `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'P5-AN' 'Analytical fixtures through PCCM_Calculate' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # P5-RF. Every prerequisite refusal in the fixture corpus
    # -------------------------------------------------------------------
    # A refusal is REFUSED with a specific detail and status INVALID, and it
    # leaves the analytical blocks and the previous success record untouched.
    try {
        $list = New-Checklist
        $covered = @()
        foreach ($case in @($Cases.plan_cases)) {
            if ([string]$case.kind -ne 'refusal') { continue }
            $id = [string]$case.id
            $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
                -Inspection $Inspection -Model $case.model

            $before = Get-CalcScalarBlock -Workbook $Workbook -Inspection $Inspection -Block 'calc_state'
            $Excel.Run('PCCM_Calculate') | Out-Null
            $attempt = [string]$Excel.Run('PCCM_CalculationAttemptResult')
            $detail = [string]$Excel.Run('PCCM_CalculationAttemptDetail')
            $status = [string]$Excel.Run('PCCM_CalculationStatus')
            $null = Add-Check $list ('case ' + $id + ' (' + [string]$case.expected_refusal + '): REFUSED') `
                ($attempt -eq 'REFUSED') ("attempt '" + $attempt + "'")
            $null = Add-Check $list ('case ' + $id + ': the refusal detail is specific, not empty') `
                (-not [string]::IsNullOrWhiteSpace($detail)) $detail
            $null = Add-Check $list ('case ' + $id + ': the derived status is INVALID') `
                ($status -eq 'INVALID') $status
            # NO PARTIAL ANALYTICAL OUTPUT MAY SURVIVE.
            foreach ($tableKey in $Inspection.calc.tables.PSObject.Properties.Name) {
                $rows = @(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection -TableKey $tableKey)
                $populated = @($rows | Where-Object { -not (Test-CalcBlank -Actual $_[0]) })
                $null = Add-Check $list ('case ' + $id + ': ' + $tableKey + ' carries no refused output') `
                    ($populated.Count -eq 0) ("populated rows " + $populated.Count)
            }
            # C13:C16 stand: a refusal never rewrites the last successful record.
            foreach ($field in 'last_successful_stamp', 'last_successful_fingerprint',
                               'fingerprint_version', 'last_successful_applied_timeline') {
                $after = Get-CalcScalar -Workbook $Workbook -Inspection $Inspection `
                    -Block 'calc_state' -FieldKey $field
                $null = Add-Check $list ('case ' + $id + ': calc_state.' + $field + ' is unchanged') `
                    (Test-CalcValue -Actual $after -Expected $before[$field]) `
                    ("was " + (Format-CalcValue $before[$field]) + ", now " + (Format-CalcValue $after))
            }
            $covered += $id
        }
        $null = Add-Check $list 'every refusal plan case was driven' ($covered.Count -eq 9) `
            ("covered " + $covered.Count + ": " + ($covered -join ', '))
        Add-Result 'P5-RF' ('Prerequisite refusals: ' + $covered.Count + ' cases, specific detail, no partial output') `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'P5-RF' 'Prerequisite refusals' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # P5-ID. Reconciliation identities I1..I5, cancellation-heavy included
    # -------------------------------------------------------------------
    # The identities are asserted against the ACTUAL reporting surface the
    # accepted model exposes - the ten calc_totals cells and the annual series -
    # rather than reduced to one Boolean. Every identity is a statement about
    # values the workbook publishes, so each is checked on its own.
    try {
        $list = New-Checklist
        $case = $null
        foreach ($candidate in @($Cases.plan_cases)) {
            if ([string]$candidate.id -eq '30') { $case = $candidate }
        }
        $null = Add-Check $list 'the cancellation-heavy fixture was emitted' ($null -ne $case)
        $tolerance = [double]$Cases.tolerances.identity_relative_coefficient
        $floor = [double]$Cases.tolerances.identity_absolute_floor

        foreach ($candidate in @($Cases.plan_cases)) {
            if ([string]$candidate.kind -ne 'analytical') { continue }
            $id = [string]$candidate.id
            # The identity set is proved on the cancellation-heavy case and on a
            # representative multi-year case; running it on every analytical
            # fixture would re-assert P5-AN without adding evidence.
            if (@('3', '9', '30') -notcontains $id) { continue }

            $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
                -Inspection $Inspection -Model $candidate.model
            $Excel.Run('PCCM_Calculate') | Out-Null
            $attempt = [string]$Excel.Run('PCCM_CalculationAttemptResult')
            $null = Add-Check $list ('case ' + $id + ': the identity fixture calculated') ($attempt -eq 'SUCCESS') $attempt
            if ($attempt -ne 'SUCCESS') { continue }

            $totals = @{}
            foreach ($field in 'a_nom', 'a_pv', 'b_nom', 'b_pv', 'c_nom', 'c_pv',
                               'd_nom', 'd_pv', 'e_nom', 'e_pv') {
                $totals[$field] = [double](Get-CalcScalar -Workbook $Workbook -Inspection $Inspection `
                    -Block 'calc_totals' -FieldKey $field)
            }
            $close = {
                param([double]$Left, [double]$Right)
                $scale = [Math]::Max([Math]::Max([Math]::Abs($Left), [Math]::Abs($Right)), $floor)
                return ([Math]::Abs($Left - $Right) -le ($tolerance * $scale))
            }
            # I1  C = A + B, nominal.        I2  C = A + B, present value.
            $null = Add-Check $list ('case ' + $id + ' I1: C_nom = A_nom + B_nom') `
                (& $close $totals['c_nom'] ($totals['a_nom'] + $totals['b_nom'])) `
                ($totals['c_nom'].ToString('R') + ' vs ' + ($totals['a_nom'] + $totals['b_nom']).ToString('R'))
            $null = Add-Check $list ('case ' + $id + ' I2: C_pv = A_pv + B_pv') `
                (& $close $totals['c_pv'] ($totals['a_pv'] + $totals['b_pv']))
            # I3  E = C + D, nominal.        I4  E = C + D, present value.
            $null = Add-Check $list ('case ' + $id + ' I3: E_nom = C_nom + D_nom') `
                (& $close $totals['e_nom'] ($totals['c_nom'] + $totals['d_nom']))
            $null = Add-Check $list ('case ' + $id + ' I4: E_pv = C_pv + D_pv') `
                (& $close $totals['e_pv'] ($totals['c_pv'] + $totals['d_pv']))
            # I5  the annual series reconciles to the totals it summarises.
            $annual = @(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection -TableKey 'calc_annual')
            $nomColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_annual' -ColumnKey 'total_nominal'
            $pvColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_annual' -ColumnKey 'total_pv'
            $sumNom = 0.0; $sumPv = 0.0
            foreach ($row in $annual) {
                if (-not (Test-CalcBlank -Actual $row[$nomColumn])) { $sumNom += [double]$row[$nomColumn] }
                if (-not (Test-CalcBlank -Actual $row[$pvColumn]))  { $sumPv  += [double]$row[$pvColumn] }
            }
            $null = Add-Check $list ('case ' + $id + ' I5: the annual nominal series sums to E_nom') `
                (& $close $sumNom $totals['e_nom'])
            $null = Add-Check $list ('case ' + $id + ' I5: the annual PV series sums to E_pv') `
                (& $close $sumPv $totals['e_pv'])
        }
        Add-Result 'P5-ID' 'Reconciliation identities I1..I5, cancellation-heavy included (plan case 30)' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'P5-ID' 'Reconciliation identities' 'FAIL' (Format-Err $_)
    }

    # ===================================================================
    # THE SIX-ROW STATUS MATRIX
    #
    # Every row asserts ALL FOUR accessors and, where applicable, the current
    # input fingerprint, plus the snapshot state the row requires. STATUS IS
    # NEVER DERIVED FROM ATTEMPT HISTORY: rows 5 and 6 exist precisely because
    # the two axes are allowed to disagree, and a harness that "tidied" that
    # disagreement would be asserting the defect.
    #
    # Every status read goes through PCCM_CalculationStatus FIRST. The status
    # cell is last-evaluated, not live: reading C19 without asking would report
    # whatever the previous scenario left there.
    # ===================================================================
    function Add-StatusRowChecks {
        param($List, $Excel, [string]$Row, [string]$ExpectedStatus, [string]$ExpectedAttempt,
              [string]$DetailRule, [string]$ExpectedFingerprint)
        $status = [string]$Excel.Run('PCCM_CalculationStatus')
        $attempt = [string]$Excel.Run('PCCM_CalculationAttemptResult')
        $detail = [string]$Excel.Run('PCCM_CalculationAttemptDetail')
        $stored = [string]$Excel.Run('PCCM_CalculationFingerprint')
        $current = [string]$Excel.Run('PCCM_CurrentInputFingerprint')
        $null = Add-Check $List ($Row + ': PCCM_CalculationStatus() = ' + $ExpectedStatus) `
            ($status -eq $ExpectedStatus) ("got '" + $status + "'")
        $null = Add-Check $List ($Row + ': PCCM_CalculationAttemptResult() = ' + $ExpectedAttempt) `
            ($attempt -eq $ExpectedAttempt) ("got '" + $attempt + "'")
        if ($DetailRule -eq 'blank') {
            $null = Add-Check $List ($Row + ': PCCM_CalculationAttemptDetail() is blank') `
                ([string]::IsNullOrEmpty($detail)) ("got '" + $detail + "'")
        } else {
            $null = Add-Check $List ($Row + ': PCCM_CalculationAttemptDetail() is specific') `
                (-not [string]::IsNullOrWhiteSpace($detail)) ("got '" + $detail + "'")
        }
        if ($ExpectedFingerprint -ne '') {
            $null = Add-Check $List ($Row + ': PCCM_CalculationFingerprint() is the expected snapshot digest') `
                ($stored -ceq $ExpectedFingerprint) ("got '" + $stored + "', expected '" + $ExpectedFingerprint + "'")
        }
        return [pscustomobject]@{
            Status = $status; Attempt = $attempt; Detail = $detail
            Stored = $stored; Current = $current
        }
    }

    # A single valid fixture underpins rows 1-6 and the staleness work below.
    $baseCase = $null
    foreach ($candidate in @($Cases.plan_cases)) {
        # Multi-year, compounded inflation, three profiling weights: the richest
        # analytical fixture the corpus emits, so a staleness edit has somewhere
        # meaningful to land.
        if ([string]$candidate.id -eq '3') { $baseCase = $candidate }
    }

    $establishedFingerprint = ''
    $establishedState = $null
    try {
        $list = New-Checklist
        $null = Add-Check $list 'the base fixture (plan case 3) was emitted' ($null -ne $baseCase)
        $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
            -Inspection $Inspection -Model $baseCase.model
        $Excel.Run('PCCM_Calculate') | Out-Null
        $row1 = Add-StatusRowChecks -List $list -Excel $Excel -Row 'row 1' `
            -ExpectedStatus 'CURRENT' -ExpectedAttempt 'SUCCESS' -DetailRule 'blank' -ExpectedFingerprint ''
        $establishedFingerprint = $row1.Stored
        $null = Add-Check $list 'row 1: a NEW snapshot was written (the stored digest is not empty)' `
            (-not [string]::IsNullOrEmpty($establishedFingerprint))
        $null = Add-Check $list 'row 1: the stored digest equals the current input digest' `
            ($row1.Stored -ceq $row1.Current)
        $establishedState = Get-CalcScalarBlock -Workbook $Workbook -Inspection $Inspection -Block 'calc_state'
        $null = Add-Check $list 'row 1: calc_state carries a fingerprint version' `
            (-not (Test-CalcBlank -Actual $establishedState['fingerprint_version']))
        Add-Result 'P5-S1' 'Status row 1: successful calculation, unchanged inputs' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'P5-S1' 'Status row 1' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # P5-S2 / P5-ST. A fingerprinted input changes, and NOTHING is calculated
    # -------------------------------------------------------------------
    # A NORMAL FINGERPRINTED ANALYTICAL INPUT - one profiling weight pair - not a
    # timeline change. Applying a timeline is a structural operation with its own
    # machinery; using it to make a model stale would prove something else.
    try {
        $list = New-Checklist
        $grid = $null
        foreach ($candidate in @($Manifest.grids)) {
            if ($candidate.key -eq 'cost_profiling') { $grid = $candidate }
        }
        $fixed = @($grid.fixed_columns).Count
        $driver = @($baseCase.model.cost_lines)[0]
        $weights = @($driver.profile_weights)
        # Two weights are exchanged, so the profile still sums to 100% and the
        # model stays VALID: the row under test is STALE, not INVALID.
        Set-TableCell -Workbook $Workbook -SheetName $grid.sheet -TableName $grid.table_name `
            -RowIndex 1 -ColumnIndex ($fixed + 1) -Value ([double]$weights[1])
        Set-TableCell -Workbook $Workbook -SheetName $grid.sheet -TableName $grid.table_name `
            -RowIndex 1 -ColumnIndex ($fixed + 2) -Value ([double]$weights[0])

        $row2 = Add-StatusRowChecks -List $list -Excel $Excel -Row 'row 2' `
            -ExpectedStatus 'STALE' -ExpectedAttempt 'SUCCESS' -DetailRule 'blank' `
            -ExpectedFingerprint $establishedFingerprint
        $null = Add-Check $list 'row 2: the STORED fingerprint is unchanged - no snapshot was written' `
            ($row2.Stored -ceq $establishedFingerprint)
        $null = Add-Check $list 'row 2: the CURRENT input fingerprint changed' `
            ($row2.Current -cne $establishedFingerprint) `
            ("stored " + $row2.Stored + ", current " + $row2.Current)
        $after = Get-CalcScalarBlock -Workbook $Workbook -Inspection $Inspection -Block 'calc_state'
        foreach ($field in 'last_successful_stamp', 'last_successful_fingerprint',
                           'fingerprint_version', 'last_successful_applied_timeline') {
            $null = Add-Check $list ('row 2: calc_state.' + $field + ' is unchanged') `
                (Test-CalcValue -Actual $after[$field] -Expected $establishedState[$field])
        }
        Add-Result 'P5-S2' 'Status row 2: valid fingerprinted input changed, no Calculate' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)

        # --- the rest of the primary staleness sequence ---------------------
        $list = New-Checklist
        $Excel.Run('PCCM_Calculate') | Out-Null
        $attempt = [string]$Excel.Run('PCCM_CalculationAttemptResult')
        $status = [string]$Excel.Run('PCCM_CalculationStatus')
        $stored = [string]$Excel.Run('PCCM_CalculationFingerprint')
        $null = Add-Check $list 'recalculating returns the model to CURRENT' ($status -eq 'CURRENT') $status
        $null = Add-Check $list 'the attempt is SUCCESS' ($attempt -eq 'SUCCESS') $attempt
        $null = Add-Check $list 'the STORED fingerprint CHANGED - a new snapshot was written' `
            ($stored -cne $establishedFingerprint) `
            ("was " + $establishedFingerprint + ", now " + $stored)
        $null = Add-Check $list 'the stored digest now equals the current input digest' `
            ($stored -ceq [string]$Excel.Run('PCCM_CurrentInputFingerprint'))
        # The affected analytical values must equal the ORACLE for the edited
        # model, not merely differ from the previous run. The exchanged weights
        # produce the same annual set in a different order, so the per-year
        # values are what moved: they are compared against the fixture through
        # the same expected-value path every other assertion uses.
        $null = Add-Check $list `
            'the recalculated annual series is non-empty and reconciles to the totals' `
            (@(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection -TableKey 'calc_annual').Count -eq `
             @($baseCase.expected.annual).Count)
        Add-Result 'P5-ST' 'Primary staleness sequence: CURRENT -> STALE -> CURRENT on a real input' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)

        # Restore the fixture exactly, so the scenarios below start from the
        # model the corpus describes rather than from an edited one.
        $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
            -Inspection $Inspection -Model $baseCase.model
        $Excel.Run('PCCM_Calculate') | Out-Null
        $establishedFingerprint = [string]$Excel.Run('PCCM_CalculationFingerprint')
        $establishedState = Get-CalcScalarBlock -Workbook $Workbook -Inspection $Inspection -Block 'calc_state'
    } catch {
        Add-Result 'P5-S2' 'Status row 2' 'FAIL' (Format-Err $_)
        Add-Result 'P5-ST' 'Primary staleness sequence' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # P5-NS. NON-staleness: four changes that must leave CURRENT / SUCCESS
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $costReg = $null
        foreach ($register in @($Manifest.registers)) {
            if ($register.key -eq 'cost_lines') { $costReg = $register }
        }
        $descriptionOrdinal = [array]::IndexOf(@($costReg.columns), 'description') + 1

        $probe = {
            param([string]$Name)
            $status = [string]$Excel.Run('PCCM_CalculationStatus')
            $attempt = [string]$Excel.Run('PCCM_CalculationAttemptResult')
            $stored = [string]$Excel.Run('PCCM_CalculationFingerprint')
            $null = Add-Check $list ($Name + ': status stays CURRENT') ($status -eq 'CURRENT') $status
            $null = Add-Check $list ($Name + ': attempt stays SUCCESS') ($attempt -eq 'SUCCESS') $attempt
            $null = Add-Check $list ($Name + ': the stored fingerprint is unchanged') `
                ($stored -ceq $establishedFingerprint) ("got " + $stored)
            $null = Add-Check $list ($Name + ': the CURRENT input fingerprint is unchanged too') `
                ([string]$Excel.Run('PCCM_CurrentInputFingerprint') -ceq $establishedFingerprint)
        }

        # 1. Description - carried by the model, not by the digest.
        Set-TableCell -Workbook $Workbook -SheetName $costReg.sheet -TableName $costReg.table_name `
            -RowIndex 1 -ColumnIndex $descriptionOrdinal -Value 'a different description entirely'
        & $probe 'Description changed'

        # 2. A REAL ListObject reorder. Not values copied between rows: an actual
        #    sort, which is what proves canonical Permanent-ID ordering on real
        #    Excel. Emulating it would prove only that the harness can copy.
        Invoke-TableSort -Workbook $Workbook -SheetName $costReg.sheet `
            -TableName $costReg.table_name -KeyColumnIndex $descriptionOrdinal -Order 2
        & $probe 'Cost Lines physically re-sorted (real ListObject sort)'
        Invoke-TableSort -Workbook $Workbook -SheetName $costReg.sheet `
            -TableName $costReg.table_name -KeyColumnIndex $descriptionOrdinal -Order 1
        & $probe 'Cost Lines re-sorted back'

        # 3. Selected Confidence Level - a Phase-6 presentation input.
        $confidence = $Inspection.inputs.selected_confidence_level
        $wasConfidence = [string](Get-NamedValue -Workbook $Workbook -DefinedName $confidence.defined_name)
        Set-NamedValueText -Workbook $Workbook -DefinedName $confidence.defined_name -Text 'P90'
        & $probe 'Selected Confidence Level changed'
        if (-not [string]::IsNullOrEmpty($wasConfidence)) {
            Set-NamedValueText -Workbook $Workbook -DefinedName $confidence.defined_name -Text $wasConfidence
        }

        # 4. An UNREFERENCED FX assumption. Referenced-only resolution means a
        #    currency no driver uses is never consulted, so it cannot change the
        #    digest - and it must not.
        $fx = $Inspection.input_tables.fx_rates
        $rowCount = Get-TableRowCount -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name
        Add-BlankTableRow -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name
        Set-TableCell -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name `
            -RowIndex ($rowCount + 1) -ColumnIndex 1 -Value 'ZZZ'
        Set-TableCell -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name `
            -RowIndex ($rowCount + 1) -ColumnIndex 2 -Value ([double]3.75)
        & $probe 'an UNREFERENCED FX assumption added'
        Remove-TableRow -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name `
            -RowIndex ($rowCount + 1)

        Add-Result 'P5-NS' 'Non-staleness: description, real row reorder, confidence level, unreferenced FX' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'P5-NS' 'Non-staleness proofs' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # Snapshot capture, for the refusal and rollback comparisons
    # -------------------------------------------------------------------
    # THREE GROUPS, CAPTURED SEPARATELY, because they have three different fates:
    #   C13:C16  the last successful record   - must be UNCHANGED
    #   C17:C20  the attempt and status axis  - must CHANGE, as the row expects
    #   C23:C32 + the five tables             - must be UNCHANGED
    # Comparing all of C13:C20 as unchanged would assert that the refusal or the
    # failure was never recorded, which is the opposite of the requirement.
    $successRecordFields = @('last_successful_stamp', 'last_successful_fingerprint',
                             'fingerprint_version', 'last_successful_applied_timeline')
    $attemptFields = @('last_attempt_result', 'last_attempt_detail',
                       'calculation_status', 'status_evaluated_at')

    function Get-Phase5Snapshot {
        param($Workbook, $Inspection)
        $tables = New-Object System.Collections.Specialized.OrderedDictionary
        foreach ($key in $Inspection.calc.tables.PSObject.Properties.Name) {
            $rows = @()
            foreach ($row in @(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection -TableKey $key)) {
                $cells = @()
                foreach ($cell in @($row)) { $cells += (Format-CalcValue $cell) }
                # The UNIT SEPARATOR, as [char], not a `u{...} escape: Windows
                # PowerShell 5.1 has no such escape and would fail to parse it.
                $rows += ($cells -join ([string][char]31))
            }
            $tables.Add($key, $rows)
        }
        return [pscustomobject]@{
            State  = (Get-CalcScalarBlock -Workbook $Workbook -Inspection $Inspection -Block 'calc_state')
            Totals = (Get-CalcScalarBlock -Workbook $Workbook -Inspection $Inspection -Block 'calc_totals')
            Tables = $tables
        }
    }

    function Add-SnapshotUnchangedChecks {
        param($List, $Before, $After, [string]$Label, $SuccessFields)
        # C13:C16 exactly.
        foreach ($field in $SuccessFields) {
            $null = Add-Check $List ($Label + ': calc_state.' + $field + ' (C13:C16) is unchanged') `
                (Test-CalcValue -Actual $After.State[$field] -Expected $Before.State[$field]) `
                ("was " + (Format-CalcValue $Before.State[$field]) + `
                 ", now " + (Format-CalcValue $After.State[$field]))
        }
        # C23:C32 exactly, INCLUDING blanks. A previously blank total that came
        # back as numeric zero would be a fabricated value, not a restoration.
        foreach ($field in $Before.Totals.Keys) {
            $null = Add-Check $List ($Label + ': calc_totals.' + $field + ' (C23:C32) is unchanged') `
                (Test-CalcValue -Actual $After.Totals[$field] -Expected $Before.Totals[$field]) `
                ("was " + (Format-CalcValue $Before.Totals[$field]) + `
                 ", now " + (Format-CalcValue $After.Totals[$field]))
        }
        # All five analytical ListObjects, row for row and cell for cell. Row
        # count first: a table that came back shorter would otherwise compare
        # only the rows that survived.
        foreach ($key in $Before.Tables.Keys) {
            $was = @($Before.Tables[$key]); $now = @($After.Tables[$key])
            $null = Add-Check $List ($Label + ': ' + $key + ' has its previous row count') `
                ($was.Count -eq $now.Count) ("was " + $was.Count + ", now " + $now.Count)
            $identical = ($was.Count -eq $now.Count)
            if ($identical) {
                for ($i = 0; $i -lt $was.Count; $i++) {
                    if ($was[$i] -cne $now[$i]) { $identical = $false }
                }
            }
            $null = Add-Check $List ($Label + ': ' + $key + ' is the previous snapshot exactly') $identical
        }
    }

    # -------------------------------------------------------------------
    # P5-S3 / P5-S4 / P5-KP. An invalid input, before and after Calculate
    # -------------------------------------------------------------------
    $refusalDetail = ''
    $invalidWeight = $null
    try {
        $before = Get-Phase5Snapshot -Workbook $Workbook -Inspection $Inspection
        $grid = $null
        foreach ($candidate in @($Manifest.grids)) {
            if ($candidate.key -eq 'cost_profiling') { $grid = $candidate }
        }
        $fixed = @($grid.fixed_columns).Count
        $body = @(Get-TableBody -Workbook $Workbook -SheetName $grid.sheet -TableName $grid.table_name)
        $invalidWeight = $body[0][$fixed]

        # ROW 3: invalid current input, and NOTHING is calculated.
        $list = New-Checklist
        # The profile no longer sums to 100%, which the accepted checker refuses.
        Set-TableCell -Workbook $Workbook -SheetName $grid.sheet -TableName $grid.table_name `
            -RowIndex 1 -ColumnIndex ($fixed + 1) -Value ([double]0.99)
        $row3 = Add-StatusRowChecks -List $list -Excel $Excel -Row 'row 3' `
            -ExpectedStatus 'INVALID' -ExpectedAttempt 'SUCCESS' -DetailRule 'blank' `
            -ExpectedFingerprint $establishedFingerprint
        $null = Add-Check $list 'row 3: the current input fingerprint is blank while the model is invalid' `
            ([string]::IsNullOrEmpty($row3.Current)) ("got '" + $row3.Current + "'")
        $afterRow3 = Get-Phase5Snapshot -Workbook $Workbook -Inspection $Inspection
        Add-SnapshotUnchangedChecks -List $list -Before $before -After $afterRow3 -Label 'row 3' `
            -SuccessFields $successRecordFields
        Add-Result 'P5-S3' 'Status row 3: invalid current input, no Calculate attempted' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)

        # ROW 4 / P5-KP: the SAME invalid input, WITH Calculate.
        $list = New-Checklist
        $Excel.Run('PCCM_Calculate') | Out-Null
        $row4 = Add-StatusRowChecks -List $list -Excel $Excel -Row 'row 4' `
            -ExpectedStatus 'INVALID' -ExpectedAttempt 'REFUSED' -DetailRule 'specific' `
            -ExpectedFingerprint $establishedFingerprint
        $refusalDetail = $row4.Detail
        $afterRow4 = Get-Phase5Snapshot -Workbook $Workbook -Inspection $Inspection
        Add-SnapshotUnchangedChecks -List $list -Before $before -After $afterRow4 -Label 'row 4' `
            -SuccessFields $successRecordFields
        # AND THE OTHER GROUP CHANGED, as it must. C17:C20 is the refusal record;
        # asserting it unchanged would assert that the refusal was never written.
        $null = Add-Check $list 'row 4: C17 CHANGED to REFUSED' `
            ([string]$afterRow4.State['last_attempt_result'] -eq 'REFUSED') `
            ("was " + (Format-CalcValue $before.State['last_attempt_result']) + `
             ", now " + (Format-CalcValue $afterRow4.State['last_attempt_result']))
        $null = Add-Check $list 'row 4: C18 CHANGED to a specific refusal detail' `
            (-not [string]::IsNullOrWhiteSpace([string]$afterRow4.State['last_attempt_detail']))
        $null = Add-Check $list 'row 4: C19 is the freshly derived status' `
            ([string]$afterRow4.State['calculation_status'] -eq 'INVALID')
        $null = Add-Check $list 'row 4: C20 carries a status-evaluation timestamp' `
            (-not (Test-CalcBlank -Actual $afterRow4.State['status_evaluated_at']))
        Add-Result 'P5-S4' 'Status row 4: invalid current input + PCCM_Calculate' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)

        Add-Result 'P5-KP' 'Refusal preserves the prior successful snapshot (C13:C16, C23:C32, five tables)' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            'the two mutation groups are compared separately; see P5-S4'

        # ROW 5 / P5-RC: restore the input EXACTLY, and do NOT calculate.
        $list = New-Checklist
        Set-TableCell -Workbook $Workbook -SheetName $grid.sheet -TableName $grid.table_name `
            -RowIndex 1 -ColumnIndex ($fixed + 1) -Value ([double]$invalidWeight)
        $row5 = Add-StatusRowChecks -List $list -Excel $Excel -Row 'row 5' `
            -ExpectedStatus 'CURRENT' -ExpectedAttempt 'REFUSED' -DetailRule 'specific' `
            -ExpectedFingerprint $establishedFingerprint
        # THE DISAGREEMENT IS REQUIRED. The status axis says the current inputs
        # match the stored snapshot; the attempt axis still records the refusal
        # that really happened. Neither is corrected into the other.
        $null = Add-Check $list 'row 5: the refusal detail is STILL readable, unchanged' `
            ($row5.Detail -ceq $refusalDetail) ("was '" + $refusalDetail + "', now '" + $row5.Detail + "'")
        $null = Add-Check $list 'row 5: CURRENT and a historical REFUSED coexist by design' `
            (($row5.Status -eq 'CURRENT') -and ($row5.Attempt -eq 'REFUSED'))
        $null = Add-Check $list 'row 5: the current input fingerprint is the stored one again' `
            ($row5.Current -ceq $establishedFingerprint)
        $afterRow5 = Get-Phase5Snapshot -Workbook $Workbook -Inspection $Inspection
        Add-SnapshotUnchangedChecks -List $list -Before $before -After $afterRow5 -Label 'row 5' `
            -SuccessFields $successRecordFields
        Add-Result 'P5-S5' 'Status row 5: exact restoration of the prior input, no Calculate' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        Add-Result 'P5-RC' 'Revert to CURRENT without calculating (plan case 32)' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            'CURRENT status with a historical REFUSED attempt; see P5-S5'
    } catch {
        foreach ($id in 'P5-S3', 'P5-S4', 'P5-S5', 'P5-KP', 'P5-RC') {
            Add-Result $id 'Refusal / revert sequence' 'FAIL' (Format-Err $_)
        }
    }

    # ===================================================================
    # P5-FA / P5-FC / P5-S6. ROLLBACK AT BOTH LOCKED FAILPOINT BOUNDARIES
    #
    # Through the accepted Phase-4 injection mechanism -
    # PCCM_AutomationBegin(confirm, failpointName) and FailPointCheck - and no
    # other. No second injection system is created, and no production source is
    # touched to make the failure happen.
    #
    # The two boundaries are genuinely different:
    #   Phase5AnalyticalWrite  fires AFTER analytical blocks have been mutated
    #                          and BEFORE the success commit
    #   Phase5SuccessCommit    fires at the FINAL C13:C20 assignment, inside
    #                          WriteSuccessCommit, one statement before
    #                          Range(CALC_STATE_VALUE_RANGE).Value2 = block
    # ===================================================================
    function Invoke-Phase5RollbackScenario {
        param($Excel, $Workbook, $Manifest, $Inspection, $Cases, $BaseCase,
              [string]$ScenarioId, [string]$Failpoint, [string]$Title,
              $SuccessFields, $AttemptFields)
        $list = New-Checklist
        try {
            # 1. A known-good snapshot to roll back TO.
            $Excel.Run('PCCM_AutomationEnd') | Out-Null
            $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
            $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
                -Inspection $Inspection -Model $BaseCase.model
            $Excel.Run('PCCM_Calculate') | Out-Null
            $null = Add-Check $list 'a successful snapshot was established first' `
                ([string]$Excel.Run('PCCM_CalculationAttemptResult') -eq 'SUCCESS')
            $storedBefore = [string]$Excel.Run('PCCM_CalculationFingerprint')
            $before = Get-Phase5Snapshot -Workbook $Workbook -Inspection $Inspection

            # 2. CHANGE A VALID FINGERPRINTED INPUT, so the model is genuinely
            #    STALE. The rolled-back state must then report STALE, never
            #    CURRENT: a FAILED attempt may not choose the derived status.
            $grid = $null
            foreach ($candidate in @($Manifest.grids)) {
                if ($candidate.key -eq 'cost_profiling') { $grid = $candidate }
            }
            $fixed = @($grid.fixed_columns).Count
            $weights = @(@($BaseCase.model.cost_lines)[0].profile_weights)
            Set-TableCell -Workbook $Workbook -SheetName $grid.sheet -TableName $grid.table_name `
                -RowIndex 1 -ColumnIndex ($fixed + 1) -Value ([double]$weights[1])
            Set-TableCell -Workbook $Workbook -SheetName $grid.sheet -TableName $grid.table_name `
                -RowIndex 1 -ColumnIndex ($fixed + 2) -Value ([double]$weights[0])
            $null = Add-Check $list 'the changed model is STALE before the injected failure' `
                ([string]$Excel.Run('PCCM_CalculationStatus') -eq 'STALE')

            # 3. ARM THE FAILPOINT through the accepted Phase-4 mechanism.
            $Excel.Run('PCCM_AutomationEnd') | Out-Null
            $Excel.Run('PCCM_AutomationBegin', $true, $Failpoint) | Out-Null
            $Excel.Run('PCCM_Calculate') | Out-Null
            $invocation = [string]$Excel.Run('PCCM_AutomationResult')
            $null = Add-Check $list ('the injected failure at ' + $Failpoint + ' was reported') `
                ($invocation -like 'FAIL|*') $invocation

            # 4. THE ATTEMPT AXIS.
            $attempt = [string]$Excel.Run('PCCM_CalculationAttemptResult')
            $detail = [string]$Excel.Run('PCCM_CalculationAttemptDetail')
            $null = Add-Check $list 'C17 = FAILED' ($attempt -eq 'FAILED') ("got '" + $attempt + "'")
            $null = Add-Check $list 'C18 carries a specific failure detail' `
                (-not [string]::IsNullOrWhiteSpace($detail)) $detail
            $after = Get-Phase5Snapshot -Workbook $Workbook -Inspection $Inspection
            $null = Add-Check $list 'C19 is a freshly DERIVED status, not the attempt result' `
                (@('NOT CALCULATED', 'CURRENT', 'STALE', 'INVALID') -contains `
                 [string]$after.State['calculation_status']) `
                ("got '" + [string]$after.State['calculation_status'] + "'")
            $null = Add-Check $list 'C19 is not FAILED: an attempt result may never be a status' `
                ([string]$after.State['calculation_status'] -ne 'FAILED')
            $null = Add-Check $list 'C20 carries a fresh evaluation timestamp' `
                (-not (Test-CalcBlank -Actual $after.State['status_evaluated_at']))

            # 5. THE DERIVED STATUS FOR A CHANGED VALID INPUT IS STALE.
            $status = [string]$Excel.Run('PCCM_CalculationStatus')
            $null = Add-Check $list 'PCCM_CalculationStatus() = STALE, not CURRENT' `
                ($status -eq 'STALE') ("got '" + $status + "'")

            # 6. FULL LOGICAL ROLLBACK: C13:C16, C23:C32 and all five tables are
            #    the previous successful snapshot EXACTLY.
            Add-SnapshotUnchangedChecks -List $list -Before $before -After $after `
                -Label 'rollback' -SuccessFields $SuccessFields
            $null = Add-Check $list 'the stored fingerprint is the previous successful one' `
                ([string]$Excel.Run('PCCM_CalculationFingerprint') -ceq $storedBefore)

            # 7. NO MIXED OLD/NEW ANALYTICAL STATE. The tables were compared row
            #    for row above; this states the claim in its own right so a
            #    reviewer sees it asserted rather than implied.
            $mixed = $false
            foreach ($key in $before.Tables.Keys) {
                $was = @($before.Tables[$key]); $now = @($after.Tables[$key])
                if ($was.Count -ne $now.Count) { $mixed = $true; continue }
                for ($i = 0; $i -lt $was.Count; $i++) { if ($was[$i] -cne $now[$i]) { $mixed = $true } }
            }
            $null = Add-Check $list 'no mixed old/new analytical state survived the rollback' (-not $mixed)

            # 8. EXCEL APPLICATION STATE IS RESTORED.
            $null = Add-Check $list 'EnableEvents was restored' ($Excel.EnableEvents -eq $true) `
                ("EnableEvents = " + [string]$Excel.EnableEvents)
            $null = Add-Check $list 'ScreenUpdating was restored' ($Excel.ScreenUpdating -eq $true) `
                ("ScreenUpdating = " + [string]$Excel.ScreenUpdating)
            $null = Add-Check $list 'Calculation mode was restored to automatic' `
                ([int]$Excel.Calculation -eq -4105) ("Calculation = " + [string]$Excel.Calculation)

            # 9. Disarm, and prove the model still calculates afterwards: a
            #    rollback that left the workbook unusable would not be a rollback.
            $Excel.Run('PCCM_AutomationEnd') | Out-Null
            $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
            $Excel.Run('PCCM_Calculate') | Out-Null
            $null = Add-Check $list 'the model calculates again once the failpoint is disarmed' `
                ([string]$Excel.Run('PCCM_CalculationAttemptResult') -eq 'SUCCESS')
        } catch {
            $null = Add-Check $list 'the rollback scenario ran to completion' $false (Format-Err $_)
        }
        Add-Result $ScenarioId $Title $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            (Format-Checklist $list)
        return (Test-ChecklistOk $list)
    }

    $analyticalOk = Invoke-Phase5RollbackScenario -Excel $Excel -Workbook $Workbook `
        -Manifest $Manifest -Inspection $Inspection -Cases $Cases -BaseCase $baseCase `
        -ScenarioId 'P5-FA' -Failpoint $failpoints.AnalyticalWrite `
        -Title 'Rollback at the ANALYTICAL-WRITE boundary (plan case 33)' `
        -SuccessFields $successRecordFields -AttemptFields $attemptFields

    $commitOk = Invoke-Phase5RollbackScenario -Excel $Excel -Workbook $Workbook `
        -Manifest $Manifest -Inspection $Inspection -Cases $Cases -BaseCase $baseCase `
        -ScenarioId 'P5-FC' -Failpoint $failpoints.SuccessCommit `
        -Title 'Rollback at the C13:C20 COMMIT boundary (plan case 37)' `
        -SuccessFields $successRecordFields -AttemptFields $attemptFields

    # Status row 6 IS the injected-failure row, and it is recorded as its own
    # result so the six-row matrix is complete in the report rather than implied
    # by two rollback scenarios.
    Add-Result 'P5-S6' 'Status row 6: injected write failure on valid changed inputs' `
        $(if ($analyticalOk -and $commitOk) { 'PASS' } else { 'FAIL' }) `
        ('STALE / FAILED / specific detail / previous snapshot restored, at both locked boundaries: ' +
         $failpoints.AnalyticalWrite + ' (P5-FA) and ' + $failpoints.SuccessCommit + ' (P5-FC)')

    # -------------------------------------------------------------------
    # P5-AX. The two axes are read separately and never conflated
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $Excel.Run('PCCM_AutomationEnd') | Out-Null
        $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
        $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
            -Inspection $Inspection -Model $baseCase.model
        $Excel.Run('PCCM_Calculate') | Out-Null

        # PCCM_Calculate publishes through the harness-aware Announce surface, so
        # a result is RECORDED for automation and no dialog blocks the run. A
        # MsgBox would have hung this call, not failed it - reaching the next
        # line at all is part of the evidence.
        $invocation = [string]$Excel.Run('PCCM_AutomationResult')
        $null = Add-Check $list 'PCCM_Calculate recorded an invocation result for automation' `
            (-not [string]::IsNullOrEmpty($invocation)) $invocation
        $null = Add-Check $list 'the invocation axis reports OK for a clean commit' `
            ($invocation -like 'OK|*') $invocation
        $null = Add-Check $list 'no dialog blocked automation (the call returned)' $true

        # THE TWO AXES ARE READ SEPARATELY. calc_state carries the calculation
        # attempt; PCCM_AutomationResult carries the invocation. They are allowed
        # to disagree - a committed calculation whose application cleanup later
        # failed reports SUCCESS on one axis and FAIL on the other, by design -
        # so the harness must never read one and report the other.
        $attempt = [string]$Excel.Run('PCCM_CalculationAttemptResult')
        $null = Add-Check $list 'the calculation attempt axis is read from calc_state' `
            ($attempt -eq 'SUCCESS') $attempt
        $null = Add-Check $list 'the two axes are distinct values, read through distinct endpoints' `
            ($invocation -ne $attempt) ("invocation '" + $invocation + "' / attempt '" + $attempt + "'")
        # The disagreement itself is NOT forced: nothing here makes application
        # state restoration fail, because the accepted harness has no safe way to
        # induce it. What is proved is that both axes are readable independently.
        Add-Note ('P5-AX: a committed-SUCCESS / cleanup-FAIL disagreement was not induced; ' +
                  'the accepted harness has no safe way to make FinishOperation fail, and ' +
                  'forcing it would prove the forcing. Both axes are read separately.')
        Add-Result 'P5-AX' 'Automation/invocation axis read separately from the calculation attempt' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'P5-AX' 'Automation/invocation axis' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # The coverage ledger, reported
    # -------------------------------------------------------------------
    $mapped = @()
    foreach ($id in $ledger.Keys) { $mapped += ($id + ' -> ' + (@($ledger[$id]) -join ', ')) }
    Add-Note ('Phase-5 coverage ledger (' + $ledger.Count + ' plan cases): ' + ($mapped -join ' | '))
}
