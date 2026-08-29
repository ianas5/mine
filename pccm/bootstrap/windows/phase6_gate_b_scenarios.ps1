<#
    PCCM - PHASE 6 STEP 13 - GATE-B RUNTIME SCENARIOS (Windows / Excel COM)

    Dot-sourced into phase4_functional_test.ps1. It runs inside THAT script's one
    COM lifecycle, against the one Excel instance it owns, the one workbook it
    opened and the one Stage-B bootstrap it ran, and reports through the same
    Add-Result. It creates no Excel process, no release ledger, no bootstrap
    invocation and no shutdown of its own.

    THE LINE THIS FILE MUST NOT CROSS
    ---------------------------------
    Static/source evidence is not runtime evidence. Everything here is a
    STATEMENT ABOUT SOURCE until it has actually executed in Excel on Windows.
    No claim in this file's comments is a claim that anything has run.

    NOTHING HERE HAS BEEN EXECUTED. As submitted, no Windows run has been made.

    WHAT THIS FILE MAY NOT DO
    -------------------------
    * it restates no `_SimData` address. Every sheet, column, row, range, cell
      and defined name comes from build/phase6_gate_b_inspection.json;
    * it restates no expected simulation number. Every digest, seed, ladder
      value and bound comes from build/phase6_gate_b_cases.json, which the
      accepted Python oracle generates;
    * it recomputes no simulation, no digest and no statistic of its own. There
      is deliberately no fallback path: if the expectation artefact is missing
      or malformed the run FAILS, it does not fall back to self-consistency;
    * it modifies no production VBA, no contract and no workbook layout.

    WHAT IT DOES WRITE
    ------------------
    Two Setup controls, a Phase-5 fixture through the accepted Phase-5 helpers,
    and - in the recovery and run-ID scenarios only - specific `_SimData`
    machine cells whose ORIGINAL VALUE IS CAPTURED FIRST AND RESTORED AFTER, on
    the disposable %TEMP% copy the Phase-4 harness already drives. A scenario
    whose restoration cannot be verified is a FAIL, never a note.
#>

Set-StrictMode -Version 2.0

# ===========================================================================
# THE SCENARIO SET
# ===========================================================================
# Every ID this file may record. The preflight refuses a coverage entry naming
# anything outside this set, and P6-FIN proves each required ID has exactly one
# result, so neither a typo nor a silent omission can shrink the matrix.
function Get-Phase6ScenarioIds {
    return @(
        'P6-PRE', 'P6-ART', 'P6-CMP', 'P6-M', 'P6-API', 'P6-BTN', 'P6-INIT',
        'P6-FX1', 'P6-DET', 'P6-ORA', 'P6-FIXED-INERT',
        'P6-AU1', 'P6-AU2',
        'P6-BANK', 'P6-ACC',
        'P6-RF1', 'P6-PRESERVE',
        'P6-FP1', 'P6-FP2', 'P6-FP3',
        'P6-REC1', 'P6-REC2', 'P6-REC3', 'P6-REC4', 'P6-REC5',
        'P6-RIDMAX', 'P6-AXIS',
        'P6-SU', 'P6-XX', 'P6-LDG', 'P6-FIN'
    )
}

# The FUNCTIONAL scenarios a complete Step-13 run must produce a result for.
#
# THREE THINGS ARE DELIBERATELY ABSENT, and the absences are what keep the
# finalisation acyclic:
#
#   P6-FIN  is the completeness verdict OVER this set. Requiring itself would
#           make the verdict its own precondition.
#   P6-LDG  is the ledger's verdict over every guarded result INCLUDING P6-FIN,
#           so it is emitted last, through the UNGUARDED reporter, and must not
#           be something P6-FIN needs in order to exist. That circularity is the
#           Round-4A defect the Phase-5 block already removed: a ledger verdict
#           emitted before the last guarded result cannot see a duplicate of it.
#   P6-SU / P6-XX  are failure channels a clean run never reaches.
function Get-Phase6RequiredScenarioIds {
    return @(
        'P6-PRE', 'P6-ART', 'P6-CMP', 'P6-M', 'P6-API', 'P6-BTN', 'P6-INIT',
        'P6-FX1', 'P6-DET', 'P6-ORA', 'P6-FIXED-INERT',
        'P6-AU1', 'P6-AU2',
        'P6-BANK', 'P6-ACC',
        'P6-RF1', 'P6-PRESERVE',
        'P6-FP1', 'P6-FP2', 'P6-FP3',
        'P6-REC1', 'P6-REC2', 'P6-REC3', 'P6-REC4', 'P6-REC5',
        'P6-RIDMAX', 'P6-AXIS'
    )
}

# The scenarios that WRITE machine state directly, and therefore the ones whose
# restoration failure can contaminate everything after them.
function Get-Phase6FixtureScenarioIds {
    return @('P6-REC1', 'P6-REC2', 'P6-REC3', 'P6-REC4', 'P6-REC5', 'P6-RIDMAX')
}

# The accepted production baseline this harness is evidence infrastructure FOR.
#
# A CHECKED COPY, not a second authority: `tests/test_phase6_gate_b_harness_source.py`
# pins this string against its own `PRODUCTION_BASELINE`, and that test proves
# `git diff <baseline> -- pccm/src pccm/spec` is empty for the tree being
# reviewed. It is NOT `git rev-parse HEAD`: HEAD on the runtime tree is the
# RUNTIME-HARNESS commit, and reporting it as the production baseline would
# conflate the two identities the Step-13 authorisation requires to be distinct.
function Get-Phase6ProductionBaseline { return 'bc7949b' }

# The accepted Phase-6 production modules whose bytes must be the baseline's.
# Named, never counted, and never inferred from the compiled project: that a
# project CONTAINS a module named modSimReport is a different fact from the
# modSimReport source being exercised being the accepted one.
function Get-Phase6ProductionModules {
    return @('modSimContract', 'modSimRng', 'modSimSample', 'modSimEngine',
             'modSimStats', 'modSimFingerprint', 'modSimNonce', 'modSimReport')
}

# The three accepted Phase-6 failpoint stage names.
#
# DECLARED IN PRODUCTION VBA, NOT IN A CONTRACT. `FAILPOINT_SIM_AFTER_NONCE`
# lives in modSimNonce.bas; `FAILPOINT_SIM_CANDIDATE_BANK` and
# `FAILPOINT_SIM_FINAL_COMMIT` live in modSimReport.bas. That is why they are
# not projected into the inspection artefact, which carries only what a
# contract owns. tests/test_phase6_gate_b_harness_source.py pins these three
# strings against those two modules, so production stays the authority and this
# is a CHECKED COPY rather than a second declaration.
function Get-Phase6FailpointNames {
    return [pscustomobject]@{
        AfterNoncePersisted = 'Phase6AfterNoncePersisted'
        CandidateBank       = 'Phase6CandidateBank'
        FinalCommit         = 'Phase6FinalCommit'
    }
}

# ===========================================================================
# ONE RESULT PER SCENARIO ID
# ===========================================================================
# The same fail-closed ledger the Phase-5 block uses, and for the same reason: a
# second result for an ID must not be reducible to a Note, because the driver
# declares success from the FAIL count and Notes do not contribute to it.
$script:Phase6RecordedIds = New-Object System.Collections.ArrayList
$script:Phase6LedgerViolations = New-Object System.Collections.ArrayList
$script:Phase6LedgerReported = $false

function Reset-Phase6ResultLedger {
    $script:Phase6RecordedIds = New-Object System.Collections.ArrayList
    $script:Phase6LedgerViolations = New-Object System.Collections.ArrayList
    $script:Phase6LedgerReported = $false
}

# ===========================================================================
# FIXTURE INTEGRITY - the fail-closed flag
# ===========================================================================
# Several scenarios write `_SimData` machine cells directly, and every one of
# them shares ONE workbook with everything that follows. If a fixture cannot be
# restored - because the write itself raised, because the read-back disagrees,
# or because an exception carried the scenario past its restore - then the
# workbook is no longer the workbook the later scenarios assume.
#
# CONTINUING WOULD PRODUCE BEHAVIOURAL EVIDENCE FROM A STATE THE HARNESS PUT
# THERE. That is worse than no evidence, because it looks like evidence. So the
# flag latches false, and every remaining STATEFUL scenario records
# FAIL / not attempted instead of running. Lifecycle and finalisation still run:
# the point is to stop making claims about behaviour, not to stop reporting.
$script:Phase6FixtureIntegrity = $true
$script:Phase6ContaminationReason = ''

function Reset-Phase6FixtureIntegrity {
    $script:Phase6FixtureIntegrity = $true
    $script:Phase6ContaminationReason = ''
}

function Test-Phase6FixtureIntegrity { return $script:Phase6FixtureIntegrity }

function Get-Phase6ContaminationReason { return $script:Phase6ContaminationReason }

function Set-Phase6Contaminated {
    param([string]$Reason)
    # LATCHING. The first contamination is the one that matters; a later
    # scenario cannot clear it by succeeding at something else.
    if ($script:Phase6FixtureIntegrity) {
        $script:Phase6FixtureIntegrity = $false
        $script:Phase6ContaminationReason = $Reason
        Add-Note ('P6 HARNESS STATE CONTAMINATED: ' + $Reason +
                  '. Every remaining stateful Step-13 scenario will record ' +
                  'FAIL / not attempted rather than produce behavioural evidence ' +
                  'from a workbook this harness could not restore.')
    }
}

function Test-Phase6ResultRecorded {
    param([string]$Id)
    return (@($script:Phase6RecordedIds | Where-Object { $_ -eq $Id }).Count -gt 0)
}

function Get-Phase6LedgerViolations { return @($script:Phase6LedgerViolations) }

function Add-Phase6Result {
    param([string]$Id, [string]$Name, [string]$Status, [string]$Detail = '')
    if (Test-Phase6ResultRecorded -Id $Id) {
        $null = $script:Phase6LedgerViolations.Add(
            $Id + ' (attempted as ' + $Status + '): ' + $Detail)
        Add-Note ('P6 ledger VIOLATION: a second result for ' + $Id +
                  ' was attempted (' + $Status + ') and refused. The first result ' +
                  'stands and P6-LDG will FAIL the run. Detail: ' + $Detail)
        return
    }
    $null = $script:Phase6RecordedIds.Add($Id)
    Add-Result $Id $Name $Status $Detail
}

function Add-Phase6LedgerIntegrityResult {
    if ($script:Phase6LedgerReported) { return }
    $script:Phase6LedgerReported = $true
    $violations = Get-Phase6LedgerViolations
    Add-Result 'P6-LDG' 'Phase-6 result ledger integrity' `
        $(if ($violations.Count -eq 0) { 'PASS' } else { 'FAIL' }) `
        $(if ($violations.Count -eq 0) { 'one result per scenario ID' }
          else { 'duplicate result attempts: ' + ($violations -join '; ') })
}

function Format-Phase6Err {
    param($ErrorRecord)
    return (Format-Err $ErrorRecord)
}

# ===========================================================================
# THE PREFLIGHT - pure PowerShell, BEFORE Excel is started
# ===========================================================================
# Everything checkable without a workbook is checked before a COM object exists,
# so an artefact that never arrived stops the run at a point where the diagnosis
# is one line rather than a cascade of COM failures forty minutes later.
function Invoke-Phase6CoveragePreflight {
    param([string]$BuildDir)

    $ok = $true
    $inspectionPath = Join-Path $BuildDir 'phase6_gate_b_inspection.json'
    $casesPath = Join-Path $BuildDir 'phase6_gate_b_cases.json'

    foreach ($pair in @(@{ Path = $inspectionPath; What = 'inspection projection' },
                        @{ Path = $casesPath;      What = 'parity expectation corpus' })) {
        if (-not (Test-Path -LiteralPath $pair.Path)) {
            Write-Host ("  [FAIL] Phase-6 preflight: the " + $pair.What + ' is missing: ' +
                        $pair.Path) -ForegroundColor Red
            $ok = $false
        }
    }
    if (-not $ok) { return $false }

    $inspection = Get-Content -LiteralPath $inspectionPath -Raw | ConvertFrom-Json
    $cases = Get-Content -LiteralPath $casesPath -Raw | ConvertFrom-Json

    $list = New-Checklist

    # THE PROJECTION CARRIES WHAT THE SCENARIOS ASK FOR. Named, not counted: a
    # count would pass while the one key a scenario needs was the missing one.
    foreach ($key in @('schema_version', 'purpose', 'provenance', 'sim_data',
                       'publication', 'controls', 'command_surface')) {
        $null = Add-Check $list ('the inspection projection carries ' + $key) `
            ($null -ne $inspection.PSObject.Properties[$key])
    }
    foreach ($key in @('sheet', 'required_visibility', 'run_identity',
                       'pending_auto_nonce', 'iteration_records',
                       'summary_statistics', 'contingency_ladder')) {
        $null = Add-Check $list ('sim_data carries ' + $key) `
            ($null -ne $inspection.sim_data.PSObject.Properties[$key])
    }
    foreach ($key in @('next_auto_nonce', 'last_run_id', 'last_attempt_result',
                       'last_attempt_detail', 'last_attempt_seed_mode',
                       'last_attempt_effective_seed', 'last_attempt_auto_nonce',
                       'simulation_status', 'status_evaluated_at', 'active_bank',
                       'run_id', 'request_fingerprint', 'result_digest',
                       'seed_mode', 'supplied_seed', 'effective_seed',
                       'consumed_auto_nonce', 'iterations_run',
                       'last_successful_stamp', 'applied_timeline')) {
        $null = Add-Check $list ('the run-identity rows name ' + $key) `
            ($null -ne $inspection.sim_data.run_identity.rows.PSObject.Properties[$key])
    }
    foreach ($key in @('monte_carlo_iterations', 'random_seed')) {
        $null = Add-Check $list ('the simulation control ' + $key + ' is projected') `
            ($null -ne $inspection.controls.PSObject.Properties[$key])
    }

    # THE EXPECTATION CORPUS IS PRESENT AND NON-EMPTY, and every parity case
    # carries the fields a comparison consumes. A corpus that lost its digests
    # must stop the run here, not silently compare nothing on Windows.
    $parity = @($cases.parity_cases)
    $null = Add-Check $list 'the parity corpus carries at least one case' `
        ($parity.Count -ge 1) ('cases: ' + $parity.Count)
    $null = Add-Check $list 'the parity corpus states its case count truthfully' `
        ([int]$cases.case_count -eq $parity.Count) `
        ('stated ' + [string]$cases.case_count + ', present ' + $parity.Count)
    foreach ($case in $parity) {
        foreach ($key in @('result_digest', 'effective_seed', 'iterations_run',
                           'summary', 'deterministic_base')) {
            $null = Add-Check $list ($case.id + ' carries ' + $key) `
                ($null -ne $case.expected_exact.PSObject.Properties[$key])
        }
        foreach ($measure in @('nominal', 'pv')) {
            $null = Add-Check $list ($case.id + ' carries the ' + $measure + ' ladder') `
                ($null -ne $case.expected_exact.summary.PSObject.Properties[$measure])
        }
        $null = Add-Check $list ($case.id + ' names an existing plan case') `
            ([int]$case.plan_case_id -gt 0)
    }
    $null = Add-Check $list 'the corpus states an exact comparison policy' `
        ([string]$cases.comparison_policy -like 'EXACT*') `
        ([string]$cases.comparison_policy)
    foreach ($key in @('business_minimum_iterations', 'max_iterations_representable',
                       'seed_minimum', 'seed_maximum', 'run_id_maximum',
                       'nonce_initial')) {
        $null = Add-Check $list ('the corpus carries the bound ' + $key) `
            ($null -ne $cases.bounds.PSObject.Properties[$key])
    }
    foreach ($key in @('attempt_results', 'sim_states', 'seed_modes', 'quantile_labels')) {
        $null = Add-Check $list ('the corpus carries the vocabulary ' + $key) `
            ($null -ne $cases.vocabulary.PSObject.Properties[$key])
    }

    # THE COVERAGE MAP. Every parity case maps to P6-ORA, and no scenario ID
    # outside the declared set may be referenced.
    $declared = Get-Phase6ScenarioIds
    foreach ($id in (Get-Phase6RequiredScenarioIds)) {
        $null = Add-Check $list ('the required scenario ' + $id + ' is declared') `
            ($declared -contains $id)
    }

    if (Test-ChecklistOk $list) {
        Write-Host '  [PASS] PRE6  Phase-6 Gate-B artefact preflight' -ForegroundColor Green
        return $true
    }
    Write-Host '  [FAIL] PRE6  Phase-6 Gate-B artefact preflight' -ForegroundColor Red
    foreach ($line in $list) { Write-Host ('        ' + $line) -ForegroundColor DarkGray }
    return $false
}

# ===========================================================================
# `_SimData` ACCESS - every address comes from the projection
# ===========================================================================
# There is no ListObject here: the run-identity block is plain cells, so the
# Phase-5 typed TABLE reader does not apply and this is the `Get-CalcScalar`
# discipline instead - one named COM variable per object, released in the
# narrowest scope, `.Value2` and never `.Text`.
#
# THE COLUMN IS CHOSEN BY THE ROW'S OWN GROUP. `snapshot` rows exist once per
# bank; `counter`, `attempt`, `derived` and `control` rows are shared and live
# in the shared value column. A caller that had to remember which was which
# would eventually read a shared row out of a bank column, so it cannot: it
# names the field, and optionally the bank, and the projection decides.
function Get-SimColumnFor {
    param($Inspection, [string]$FieldKey, [string]$Bank = '')
    $identity = $Inspection.sim_data.run_identity
    $group = [string]$identity.groups.$FieldKey
    if ($group -eq 'snapshot') {
        if ([string]::IsNullOrEmpty($Bank)) {
            throw ("the _SimData field '" + $FieldKey + "' is a per-bank snapshot row; " +
                   'reading it without naming a bank would read whichever column ' +
                   'happened to be the shared one')
        }
        return [string]$identity.bank_value_columns.$Bank
    }
    if (-not [string]::IsNullOrEmpty($Bank)) {
        throw ("the _SimData field '" + $FieldKey + "' is shared (" + $group +
               '), so it has no per-bank column and naming one would read the ' +
               'wrong cell')
    }
    return [string]$identity.value_column
}

function Get-SimAddressFor {
    param($Inspection, [string]$FieldKey, [string]$Bank = '')
    $identity = $Inspection.sim_data.run_identity
    if ($null -eq $identity.rows.PSObject.Properties[$FieldKey]) {
        throw ("the inspection projection carries no _SimData row named '" +
               $FieldKey + "'")
    }
    $row = [int]$identity.rows.$FieldKey
    return (Get-SimColumnFor -Inspection $Inspection -FieldKey $FieldKey -Bank $Bank) +
           [string]$row
}

function Get-SimRawCell {
    param($Workbook, $Inspection, [string]$Address)
    $sheets = $null; $sheet = $null; $range = $null
    try {
        $sheets = $Workbook.Worksheets
        $sheet = $sheets.Item([string]$Inspection.sim_data.sheet)
        $range = $sheet.Range($Address)
        return $range.Value2
    } finally {
        if ($null -ne $range)  { Release-Transient $range  'Range(_SimData)';      $range  = $null }
        if ($null -ne $sheet)  { Release-Transient $sheet  'Worksheet(_SimData)';  $sheet  = $null }
        if ($null -ne $sheets) { Release-Transient $sheets 'Worksheets';           $sheets = $null }
    }
}

function Set-SimRawCell {
    param($Workbook, $Inspection, [string]$Address, $Value)
    # THE SAME TYPE DISCIPLINE Set-Phase5TypedCell established, and for the same
    # reason: PowerShell binds a COM property call site per ARGUMENT TYPE, so
    # one polymorphic assignment line silently fails the second time it is
    # reached with a different type. One assignment site per type, and an
    # unsupported captured type FAILS CLOSED rather than being normalised.
    $sheets = $null; $sheet = $null; $range = $null
    try {
        $sheets = $Workbook.Worksheets
        $sheet = $sheets.Item([string]$Inspection.sim_data.sheet)
        $range = $sheet.Range($Address)
        if ($null -eq $Value) {
            $null = $range.ClearContents()
        } elseif ($Value -is [string]) {
            $range.Value2 = [string]$Value
        } elseif ($Value -is [bool]) {
            $range.Value2 = [bool]$Value
        } elseif ($Value.GetType().FullName -ceq 'System.Double') {
            $range.Value2 = [double]$Value
        } else {
            throw ('Set-SimRawCell cannot write a value of type ' +
                   $Value.GetType().FullName + '; this helper writes exactly what ' +
                   'Excel Value2 publishes - an empty cell, System.String, ' +
                   'System.Double or System.Boolean - and converting any other ' +
                   'type would normalise the value instead of writing it')
        }
    } finally {
        if ($null -ne $range)  { Release-Transient $range  'Range(_SimData)';      $range  = $null }
        if ($null -ne $sheet)  { Release-Transient $sheet  'Worksheet(_SimData)';  $sheet  = $null }
        if ($null -ne $sheets) { Release-Transient $sheets 'Worksheets';           $sheets = $null }
    }
}

function Get-SimField {
    param($Workbook, $Inspection, [string]$FieldKey, [string]$Bank = '')
    return (Get-SimRawCell -Workbook $Workbook -Inspection $Inspection `
        -Address (Get-SimAddressFor -Inspection $Inspection -FieldKey $FieldKey -Bank $Bank))
}

function Set-SimField {
    param($Workbook, $Inspection, [string]$FieldKey, $Value, [string]$Bank = '')
    Set-SimRawCell -Workbook $Workbook -Inspection $Inspection `
        -Address (Get-SimAddressFor -Inspection $Inspection -FieldKey $FieldKey -Bank $Bank) `
        -Value $Value
}

function Get-SimPendingCell {
    param($Inspection)
    return [string]$Inspection.sim_data.pending_auto_nonce.cell
}

function Get-SimPending {
    param($Workbook, $Inspection)
    return (Get-SimRawCell -Workbook $Workbook -Inspection $Inspection `
        -Address (Get-SimPendingCell -Inspection $Inspection))
}

function Set-SimPending {
    param($Workbook, $Inspection, $Value)
    Set-SimRawCell -Workbook $Workbook -Inspection $Inspection `
        -Address (Get-SimPendingCell -Inspection $Inspection) -Value $Value
}

# THE SHARED FINAL-COMMIT BLOCK, cell by cell. Read individually so a BLANK
# stays $null instead of becoming an empty string inside a Variant array - the
# distinction between "no publication" and "a published empty string" is exactly
# what several of these scenarios turn on.
function Get-SimSharedBlock {
    param($Workbook, $Inspection)
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    $identity = $Inspection.sim_data.run_identity
    foreach ($key in $identity.rows.PSObject.Properties.Name) {
        if ([string]$identity.groups.$key -eq 'snapshot') { continue }
        $out.Add($key, (Get-SimField -Workbook $Workbook -Inspection $Inspection -FieldKey $key))
    }
    return $out
}

function Get-SimBankBlock {
    param($Workbook, $Inspection, [string]$Bank)
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    $identity = $Inspection.sim_data.run_identity
    foreach ($key in $identity.rows.PSObject.Properties.Name) {
        if ([string]$identity.groups.$key -ne 'snapshot') { continue }
        $out.Add($key, (Get-SimField -Workbook $Workbook -Inspection $Inspection `
            -FieldKey $key -Bank $Bank))
    }
    return $out
}

function Get-SimSummaryValue {
    param($Workbook, $Inspection, [string]$Bank, [string]$Measure, [string]$RowKey)
    $block = $Inspection.sim_data.summary_statistics
    $column = [string]$block.bank_value_columns.$Bank.$Measure
    $row = [int]$block.rows.$RowKey
    return (Get-SimRawCell -Workbook $Workbook -Inspection $Inspection `
        -Address ($column + [string]$row))
}

function Get-SimIterationValue {
    param($Workbook, $Inspection, [string]$Bank, [string]$ColumnKey, [int]$Offset)
    $block = $Inspection.sim_data.iteration_records
    $column = [string]$block.banks.$Bank.$ColumnKey
    $row = [int]$block.first_iteration_row + $Offset
    return (Get-SimRawCell -Workbook $Workbook -Inspection $Inspection `
        -Address ($column + [string]$row))
}

# ===========================================================================
# COMPARISON AND FORMATTING
# ===========================================================================
function Format-SimValue {
    param($Value)
    if ($null -eq $Value) { return '<blank>' }
    if ($Value -is [string]) { return ("'" + $Value + "'") }
    if ($Value -is [bool]) { return ('<bool ' + [string]$Value + '>') }
    if ($Value.GetType().FullName -ceq 'System.Double') {
        # ROUND-TRIP, NOT DISPLAY. 'R' is the shortest text that reads back to
        # the same binary64, and InvariantCulture keeps a comma out of it on a
        # locale that would otherwise supply one.
        return ([double]$Value).ToString('R', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    return ('<' + $Value.GetType().Name + ' ' + [string]$Value + '>')
}

function Test-SimBlank {
    param($Value)
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    return $false
}

function Test-SimExactDouble {
    param($Actual, [double]$Expected)
    # EXACT, AND ONLY FOR A REAL DOUBLE. A number Excel published as text is a
    # different fact from a number it published as a number, and coercing the
    # first into the second here would hide it. There is deliberately no
    # tolerance parameter: the corpus states EXACT and the canonical encoder
    # already normalises the host decimal separator before hashing.
    if ($null -eq $Actual) { return $false }
    if ($Actual.GetType().FullName -cne 'System.Double') { return $false }
    return ([double]$Actual -eq $Expected)
}

function Test-SimExactText {
    param($Actual, [string]$Expected)
    if ($null -eq $Actual) { return $false }
    if ($Actual -isnot [string]) { return $false }
    return ([string]$Actual -ceq $Expected)
}

function Test-SimSameValue {
    param($A, $B)
    # "Unchanged" means the SAME Value2, type included. A cell that held Double
    # 3 and now holds String '3' has changed, and an equality that cast both to
    # text would call that unchanged.
    if (($null -eq $A) -and ($null -eq $B)) { return $true }
    if (($null -eq $A) -or ($null -eq $B)) { return $false }
    if ($A.GetType().FullName -cne $B.GetType().FullName) { return $false }
    if ($A -is [string]) { return ([string]$A -ceq [string]$B) }
    if ($A -is [bool]) { return ([bool]$A -eq [bool]$B) }
    return ([double]$A -eq [double]$B)
}

function Add-SimUnchangedChecks {
    param($List, $Before, $After, [string]$Label)
    foreach ($key in $Before.Keys) {
        $null = Add-Check $List ($Label + ': ' + $key + ' is unchanged') `
            (Test-SimSameValue -A $Before[$key] -B $After[$key]) `
            ('was ' + (Format-SimValue $Before[$key]) + ', now ' + (Format-SimValue $After[$key]))
    }
}

# ===========================================================================
# EVIDENCE CAPTURE
# ===========================================================================
# The record every scenario returns, in the order the Step-13 authorisation
# asks for it. Read from the PERSISTED CELLS, never inferred from VBA state:
# nothing here asks a Public accessor what a cell contains when the cell can be
# read directly, and a Private Boolean inside modSimReport is not observable
# from PowerShell at all and is never claimed to be.
function Get-Phase6State {
    param($Workbook, $Inspection)
    $state = New-Object System.Collections.Specialized.OrderedDictionary
    $state.Add('shared', (Get-SimSharedBlock -Workbook $Workbook -Inspection $Inspection))
    foreach ($bank in @($Inspection.publication.bank_labels)) {
        $state.Add(('bank_' + $bank),
                   (Get-SimBankBlock -Workbook $Workbook -Inspection $Inspection -Bank $bank))
    }
    $state.Add('pending_auto_nonce', (Get-SimPending -Workbook $Workbook -Inspection $Inspection))
    return $state
}

function Format-Phase6State {
    param($State, [string]$Label)
    $lines = @($Label + ':')
    foreach ($key in $State.Keys) {
        if ($key -eq 'pending_auto_nonce') {
            $lines += ('    pending_auto_nonce = ' + (Format-SimValue $State[$key]))
            continue
        }
        $block = $State[$key]
        foreach ($field in $block.Keys) {
            $lines += ('    ' + $key + '.' + $field + ' = ' + (Format-SimValue $block[$field]))
        }
    }
    return ($lines -join "`r`n")
}

function Get-Phase6ActiveBank {
    param($State)
    $value = $State['shared']['active_bank']
    if (Test-SimBlank -Value $value) { return '' }
    return [string]$value
}

function Get-Phase6CandidateTarget {
    param($Inspection, [string]$ActiveBank)
    # THE SELECTOR MAP IS THE CONTRACT'S, projected. The harness does not know
    # that a blank selector targets A, or that A targets B; it asks.
    #
    # ENTRIES, NOT PROPERTIES, and that is not a style choice. The contract keys
    # this map by the ACTIVE BANK, and the key for "nothing has ever been
    # published" is the empty string. Projected as a JSON object it became a
    # property whose NAME was empty, and Windows PowerShell 5.1's
    # ConvertFrom-Json cannot materialise such an object at all - it threw
    # PSArgumentException on the whole artefact, and Step-13 Run 1 died in the
    # preflight before Excel was started. The projection now carries the absence
    # as a null VALUE, which is the same fact in a shape 5.1 can read.
    #
    # THE BLANK IS NORMALISED ONLY FOR COMPARISON. Nothing here renames it, and
    # no A/B mapping is restated: the answer is whatever the matching entry
    # says.
    $entries = @($Inspection.publication.candidate_target)
    if ($entries.Count -eq 0) {
        throw 'the publication candidate_target projection carries no entries'
    }
    $matched = @($entries | Where-Object {
        if ($null -eq $_.active_bank) {
            [string]::IsNullOrEmpty($ActiveBank)
        } else {
            ([string]$_.active_bank) -ceq $ActiveBank
        }
    })
    # EXACTLY ONE, FAIL-CLOSED BOTH WAYS. Zero means the projection cannot
    # answer for this state; two means it answers twice and the harness would be
    # choosing which answer to believe.
    if ($matched.Count -eq 0) {
        throw ('the publication candidate_target projection has no entry for active bank ' +
               [char]39 + $ActiveBank + [char]39)
    }
    if ($matched.Count -gt 1) {
        throw ('the publication candidate_target projection has ' + $matched.Count +
               ' entries for active bank ' + [char]39 + $ActiveBank + [char]39 +
               '; exactly one is required')
    }
    return [string]$matched[0].candidate_bank
}

# ONE INVOCATION OF THE PRODUCTION ENDPOINT, and its announcement.
#
# Deliberately NOT Invoke-Phase5ProductionOperation: that helper throws when the
# operation did not succeed, and most Phase-6 scenarios EXPECT a refusal or a
# failure. Throwing on the expected outcome would turn every negative scenario
# into a harness error.
#
# Application.Run is SYNCHRONOUS. Nothing observes workbook state while the call
# is in progress, and no scenario in this file claims to: an armed failpoint
# raises inside VBA, VBA's own handler runs, and PowerShell regains control only
# after the endpoint has returned. Every state comparison below is therefore
# between a capture taken BEFORE the call and one taken AFTER it.
function Invoke-Phase6Simulation {
    param($Excel, [string]$FailAfterStage = '')
    $Excel.Run('PCCM_AutomationBegin', $true, $FailAfterStage) | Out-Null
    try {
        $Excel.Run('PCCM_RunSimulation') | Out-Null
        return [string]$Excel.Run('PCCM_AutomationResult')
    } finally {
        # THE ARMED STAGE IS DISARMED WHATEVER HAPPENED. A failpoint left armed
        # would fire inside the next scenario and be reported against it.
        $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
    }
}

function Test-Phase6Announced {
    param([string]$Result, [string]$Kind)
    if ($Kind -eq 'OK') { return ($Result -like 'OK|*') }
    return ($Result -like 'FAIL|*')
}

# ===========================================================================
# DIRECT MACHINE-STATE FIXTURES, AND THEIR RESTORATION
# ===========================================================================
# Several scenarios must put the workbook into a state only a crashed or
# corrupted run could otherwise produce. Because every scenario shares ONE
# workbook and ONE session, a fixture that is not restored exactly is
# contamination of everything after it.
#
# THE POLICY, FOR EVERY FIXTURE-WRITING SCENARIO:
#
#   capture the typed original -> write the fixture -> VERIFY the fixture took
#   -> run the scenario -> capture POST evidence BEFORE cleanup -> restore ->
#   VERIFY the restoration is exact.
#
# A scenario whose restoration cannot be verified is a FAIL. It is never a note,
# and never a silent carry-on: the evidence value of everything after it depends
# on the workbook being back where it started.
function New-Phase6CellFixture {
    param($Workbook, $Inspection, [string]$Address)
    return [pscustomobject]@{
        Address  = $Address
        Original = (Get-SimRawCell -Workbook $Workbook -Inspection $Inspection -Address $Address)
        Written  = $false
    }
}

function Set-Phase6CellFixture {
    param($Workbook, $Inspection, $Fixture, $Value, $List, [string]$Label)
    # MARKED BEFORE THE WRITE, not after. A COM assignment that raises may still
    # have changed the cell, and a flag set only on success would let exactly
    # that case escape restoration.
    $Fixture.Written = $true
    Set-SimRawCell -Workbook $Workbook -Inspection $Inspection `
        -Address $Fixture.Address -Value $Value
    $readBack = Get-SimRawCell -Workbook $Workbook -Inspection $Inspection `
        -Address $Fixture.Address
    return (Add-Check $List ($Label + ': the fixture at ' + $Fixture.Address + ' took') `
        (Test-SimSameValue -A $readBack -B $Value) `
        ('wrote ' + (Format-SimValue $Value) + ', read back ' + (Format-SimValue $readBack)))
}

function Restore-Phase6CellFixture {
    param($Workbook, $Inspection, $Fixture, $List, [string]$Label)
    # BEST EFFORT, AND IT NEVER THROWS. This runs from a `finally`, on the
    # exception path as well as the success path, so a raise here would replace
    # the original scenario failure with a cleanup failure and lose the thing
    # that actually went wrong. A write or read-back that raises becomes a FAILED
    # CHECK, which is what the caller needs in order to latch contamination.
    if (-not $Fixture.Written) { return $true }
    $readBack = $null
    $failure = ''
    try {
        Set-SimRawCell -Workbook $Workbook -Inspection $Inspection `
            -Address $Fixture.Address -Value $Fixture.Original
        $readBack = Get-SimRawCell -Workbook $Workbook -Inspection $Inspection `
            -Address $Fixture.Address
    } catch {
        $failure = Format-Phase6Err $_
    }
    if (-not [string]::IsNullOrEmpty($failure)) {
        return (Add-Check $List ($Label + ': ' + $Fixture.Address + ' is restored exactly') `
            $false ('restoration raised: ' + $failure))
    }
    return (Add-Check $List ($Label + ': ' + $Fixture.Address + ' is restored exactly') `
        (Test-SimSameValue -A $readBack -B $Fixture.Original) `
        ('original ' + (Format-SimValue $Fixture.Original) + ', restored ' +
         (Format-SimValue $readBack)))
}

# THE ONE PLACE A FIXTURE IS UNWOUND, and it is always reached.
#
# Every direct machine-state write goes through this on the way out - success
# path, assertion-failure path and exception path alike - because a fixture that
# is not restored is not a failed scenario, it is a corrupted workbook for
# everything after it. If the restoration cannot be VERIFIED, the harness latches
# contaminated and the remaining stateful scenarios stop producing behavioural
# evidence rather than producing untrustworthy evidence.
function Complete-Phase6Fixture {
    param($Workbook, $Inspection, $Fixture, $List, [string]$Label)
    if (-not $Fixture.Written) { return $true }
    $restored = $false
    try {
        $restored = Restore-Phase6CellFixture -Workbook $Workbook -Inspection $Inspection `
            -Fixture $Fixture -List $List -Label $Label
    } catch {
        # Restore-Phase6CellFixture is written not to throw; this is the belt
        # for the braces, and it still records rather than swallowing.
        $null = Add-Check $List ($Label + ': ' + $Fixture.Address + ' is restored exactly') `
            $false ('restoration raised outside the helper: ' + (Format-Phase6Err $_))
        $restored = $false
    }
    if (-not $restored) {
        Set-Phase6Contaminated -Reason ($Label + ' could not restore ' + $Fixture.Address +
            ' (original ' + (Format-SimValue $Fixture.Original) + ')')
    }
    return $restored
}

# ===========================================================================
# THE PARITY COMPARISON
# ===========================================================================
# EXACT, and there is no other mode. If Excel and the oracle disagree, the
# scenario FAILS and the disagreement is reported as runtime evidence. Nothing
# here recomputes the expectation, falls back to comparing Excel against itself,
# or admits a tolerance.
function Add-Phase6ParityChecks {
    param($Workbook, $Inspection, $Cases, $Case, $List, [string]$Bank, [string]$Label)

    $expected = $Case.expected_exact

    foreach ($pair in @(
        @{ Field = 'result_digest';  Expected = [string]$expected.result_digest },
        @{ Field = 'request_fingerprint'; Expected = $(
            if ($null -ne $expected.PSObject.Properties['request_fingerprint'])
            { [string]$expected.request_fingerprint } else { $null }) })) {
        if ($null -eq $pair.Expected) { continue }
        $actual = Get-SimField -Workbook $Workbook -Inspection $Inspection `
            -FieldKey $pair.Field -Bank $Bank
        $null = Add-Check $List ($Label + ': published ' + $pair.Field + ' equals the oracle') `
            (Test-SimExactText -Actual $actual -Expected $pair.Expected) `
            ('oracle ' + [char]39 + $pair.Expected + [char]39 + ', workbook ' +
             (Format-SimValue $actual))
    }

    foreach ($pair in @(
        @{ Field = 'effective_seed'; Expected = [double]$expected.effective_seed },
        @{ Field = 'iterations_run'; Expected = [double]$expected.iterations_run },
        @{ Field = 'rng_version';    Expected = [double]$expected.rng_version },
        @{ Field = 'sim_method_version'; Expected = [double]$expected.sim_method_version })) {
        $actual = Get-SimField -Workbook $Workbook -Inspection $Inspection `
            -FieldKey $pair.Field -Bank $Bank
        $null = Add-Check $List ($Label + ': published ' + $pair.Field + ' equals the oracle') `
            (Test-SimExactDouble -Actual $actual -Expected $pair.Expected) `
            ('oracle ' + [string]$pair.Expected + ', workbook ' + (Format-SimValue $actual))
    }

    # THE SUMMARY LADDER, BOTH MEASURES, EVERY PUBLISHED ROW. Mandatory: a
    # digest match with a wrong ladder would mean the retained totals were right
    # and the statistics layer was not.
    foreach ($measure in @('nominal', 'pv')) {
        $ladder = $expected.summary.$measure
        foreach ($rowKey in @('mean', 'sample_standard_deviation', 'minimum', 'maximum')) {
            $actual = Get-SimSummaryValue -Workbook $Workbook -Inspection $Inspection `
                -Bank $Bank -Measure $measure -RowKey $rowKey
            $null = Add-Check $List `
                ($Label + ': ' + $measure + ' ' + $rowKey + ' equals the oracle') `
                (Test-SimExactDouble -Actual $actual -Expected ([double]$ladder.$rowKey)) `
                ('oracle ' + (Format-SimValue ([double]$ladder.$rowKey)) +
                 ', workbook ' + (Format-SimValue $actual))
        }
        $labels = @($Cases.vocabulary.quantile_labels)
        for ($index = 0; $index -lt $labels.Count; $index++) {
            $label = [string]$labels[$index]
            $rowKey = 'quantile_' + [string]($index + 1)
            $actual = Get-SimSummaryValue -Workbook $Workbook -Inspection $Inspection `
                -Bank $Bank -Measure $measure -RowKey $rowKey
            $null = Add-Check $List `
                ($Label + ': ' + $measure + ' ' + $rowKey + ' (' + $label + ') equals the oracle') `
                (Test-SimExactDouble -Actual $actual `
                    -Expected ([double]$ladder.quantiles.$label)) `
                ('oracle ' + (Format-SimValue ([double]$ladder.quantiles.$label)) +
                 ', workbook ' + (Format-SimValue $actual))
        }
        $actual = Get-SimSummaryValue -Workbook $Workbook -Inspection $Inspection `
            -Bank $Bank -Measure $measure -RowKey 'deterministic_base_a'
        $null = Add-Check $List `
            ($Label + ': ' + $measure + ' deterministic base A equals the oracle') `
            (Test-SimExactDouble -Actual $actual `
                -Expected ([double]$expected.deterministic_base.$measure)) `
            ('oracle ' + (Format-SimValue ([double]$expected.deterministic_base.$measure)) +
             ', workbook ' + (Format-SimValue $actual))
    }
}

# ===========================================================================
# THE DRIVER
# ===========================================================================
function Invoke-Phase6GateBScenarios {
    param(
        $Excel, $Workbook, $Manifest, $Inspection, $Cases,
        $SimInspection, $GateBCases,
        [string]$ScriptDir, [string]$TempRoot, $Results,
        [string]$HarnessCommit, [string]$RepoRoot
    )

    Reset-Phase6ResultLedger
    Reset-Phase6FixtureIntegrity
    $failpoints = Get-Phase6FailpointNames
    $banks = @($SimInspection.publication.bank_labels)
    $controls = $SimInspection.controls
    $bounds = $GateBCases.bounds
    $vocabulary = $GateBCases.vocabulary
    $parityCases = @($GateBCases.parity_cases)
    $goldenCase = @($parityCases | Where-Object { [int]$_.plan_case_id -eq 1 })[0]
    $planCases = @($Cases.plan_cases)

    function Get-PlanCase {
        param([int]$Id)
        $found = @($planCases | Where-Object { [int]$_.id -eq $Id })
        if ($found.Count -ne 1) {
            throw ('phase5_cases.json does not carry exactly one plan case ' + [string]$Id)
        }
        return $found[0]
    }

    # Establish one plan-case fixture and the two simulation controls. The
    # fixture itself goes through the ACCEPTED Phase-5 helper - this file writes
    # no register row, no grid cell and no timeline of its own.
    function Set-Phase6Fixture {
        param([int]$PlanCaseId, $SuppliedSeed, [int]$Iterations)
        $case = Get-PlanCase -Id $PlanCaseId
        $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
            -Inspection $Inspection -Model $case.model
        Set-NamedValue -Workbook $Workbook `
            -DefinedName ([string]$controls.monte_carlo_iterations.defined_name) `
            -Value ([double]$Iterations)
        # A BLANK Random Seed IS the AUTO request. There is no seed-mode cell.
        Set-NamedValue -Workbook $Workbook `
            -DefinedName ([string]$controls.random_seed.defined_name) `
            -Value $(if ($null -eq $SuppliedSeed) { $null } else { [double]$SuppliedSeed })
        $null = Invoke-Phase5ProductionOperation -Excel $Excel -Operation 'PCCM_Calculate' `
            -Stage ('establishing the Phase-6 fixture for plan case ' + [string]$PlanCaseId)
        return $case
    }

    # -------------------------------------------------------------------
    # P6-PRE. The prerequisite evidence, and the fail-closed path.
    # -------------------------------------------------------------------
    # A Phase-6 result on a workbook whose Phase-4 structural matrix or Phase-5
    # calculation evidence is not intact would be evidence of nothing.
    #
    # THE LIFECYCLE TOPOLOGY IS THE WHOLE DIFFICULTY, and getting it wrong once
    # already cost Gate-B Run 1. This block executes INSIDE the live automation
    # session:
    #
    #   Invoke-Phase5GateBScenarios -> Invoke-Phase6GateBScenarios ->
    #   PCCM_AutomationEnd -> workbook close -> Excel quit -> COM release ->
    #   Z -> Y -> Add-Phase4FinalCompletenessResult (P5-FIN) -> P5-LDG -> P6-LDG
    #
    # So Y and Z CANNOT have been recorded yet. Demanding the full 35-case
    # Phase-4 set here would be unsatisfiable by construction, and the accepted
    # Phase-5 source already derived the partition for exactly this reason:
    # `Get-Phase4PrerequisiteScenarioIds` is the required set MINUS the deferred
    # finalisation cases. This block uses that same derived partition, proves the
    # partition is real, and proves the deferral is real.
    #
    # THE 35/35 DEMAND IS NOT WEAKENED. It is still made, later and by the
    # already-accepted `Add-Phase4FinalCompletenessResult` / P5-FIN, after Y and
    # Z exist. Nothing here replaces it.
    $prerequisiteOk = $false
    try {
        $list = New-Checklist
        $seen = @($Results | ForEach-Object { $_.Id })

        $phase4Required = @(Get-Phase4RequiredScenarioIds)
        $phase4Deferred = @(Get-Phase4FinalizationScenarioIds)
        $phase4Prerequisite = @(Get-Phase4PrerequisiteScenarioIds)

        # THE PARTITION IS PROVED, NOT ASSERTED IN PROSE. Nothing may leave the
        # matrix by being called a lifecycle case.
        $overlap = @($phase4Prerequisite | Where-Object { $phase4Deferred -contains $_ })
        $stray = @($phase4Deferred | Where-Object { $phase4Required -notcontains $_ })
        $null = Add-Check $list `
            'the prerequisite and deferred Phase-4 sets partition the whole matrix' `
            ((($phase4Prerequisite.Count + $phase4Deferred.Count) -eq $phase4Required.Count) -and
             ($overlap.Count -eq 0) -and ($stray.Count -eq 0)) `
            ('prerequisite ' + $phase4Prerequisite.Count + ' + deferred ' +
             $phase4Deferred.Count + ' vs matrix ' + $phase4Required.Count +
             '; overlap: ' + ($overlap -join ', ') + '; not in matrix: ' + ($stray -join ', '))

        # AND THE DEFERRAL IS REAL. If a deferred case had already run it was
        # never a post-session case, and excluding it here would be a hole.
        $earlyDeferred = @($phase4Deferred | Where-Object { $seen -contains $_ })
        $null = Add-Check $list `
            ('the post-session Phase-4 cases have not run yet, so Phase 6 runs before them: ' +
             ($phase4Deferred -join ', ')) `
            ($earlyDeferred.Count -eq 0) ('already recorded: ' + ($earlyDeferred -join ', '))

        $missing = @($phase4Prerequisite | Where-Object { $seen -notcontains $_ })
        $null = Add-Check $list `
            ('all ' + $phase4Prerequisite.Count + ' pre-Phase-6 Phase-4 scenarios reported a result') `
            ($missing.Count -eq 0) ('missing: ' + ($missing -join ', '))
        $phase4 = @($Results | Where-Object { $phase4Prerequisite -contains $_.Id })
        $phase4Failed = @($phase4 | Where-Object { $_.Status -eq 'FAIL' })
        $phase4Skipped = @($phase4 | Where-Object { $_.Status -eq 'SKIP' })
        $phase4Passed = @($phase4 | Where-Object { $_.Status -eq 'PASS' })
        $null = Add-Check $list 'the Phase-4 prerequisite matrix has 0 FAIL' `
            ($phase4Failed.Count -eq 0) `
            ((@($phase4Failed | ForEach-Object { $_.Id })) -join ', ')
        $null = Add-Check $list 'the Phase-4 prerequisite matrix has 0 SKIP' `
            ($phase4Skipped.Count -eq 0) `
            ((@($phase4Skipped | ForEach-Object { $_.Id })) -join ', ')
        $null = Add-Check $list `
            ('the Phase-4 prerequisite matrix is ' + $phase4Prerequisite.Count + '/' +
             $phase4Prerequisite.Count + ' PASS') `
            ($phase4Passed.Count -eq $phase4Prerequisite.Count) `
            ('passed ' + $phase4Passed.Count + ' of ' + $phase4Prerequisite.Count)

        # --- and the Phase-5 block, on exactly the same terms ---------------
        # `Get-Phase5ScenarioIds` is the DERIVED in-session Phase-5 set: it does
        # not contain P5-FIN or P5-LDG, which are themselves post-session, so
        # demanding it here is satisfiable at this point in the lifecycle.
        $phase5Required = @(Get-Phase5ScenarioIds)
        $phase5Missing = @($phase5Required | Where-Object { $seen -notcontains $_ })
        $null = Add-Check $list `
            ('all ' + $phase5Required.Count + ' in-session Phase-5 scenarios reported a result') `
            ($phase5Missing.Count -eq 0) ('missing: ' + ($phase5Missing -join ', '))
        $phase5 = @($Results | Where-Object { $_.Id -like 'P5-*' })
        $phase5NotPassed = @($phase5 | Where-Object { $_.Status -ne 'PASS' })
        $null = Add-Check $list 'no recorded Phase-5 result failed or was skipped' `
            ($phase5NotPassed.Count -eq 0) `
            ('not PASS: ' + (@($phase5NotPassed | ForEach-Object { $_.Id + '=' + $_.Status }) -join ', '))
        foreach ($id in @('P5-P4', 'P5-CMP', 'P5-M')) {
            $found = @($Results | Where-Object { $_.Id -eq $id })
            $null = Add-Check $list ('the Phase-5 scenario ' + $id + ' ran exactly once and passed') `
                (($found.Count -eq 1) -and ($found[0].Status -eq 'PASS')) `
                ('results: ' + $found.Count)
        }
        # P5-ALL exists ONLY when Phase 5 refused to run, and P5-XX only when
        # driving it threw. Either is a Phase-5 failure by another name.
        foreach ($id in @('P5-ALL', 'P5-XX')) {
            $null = Add-Check $list ('the Phase-5 failure channel ' + $id + ' was not used') `
                ($seen -notcontains $id)
        }
        # AND THE PHASE-5 DEFERRAL IS REAL TOO, for the same reason as Y and Z.
        foreach ($id in @('P5-FIN', 'P5-LDG')) {
            $null = Add-Check $list `
                ('the post-session Phase-5 result ' + $id + ' has not run yet') `
                ($seen -notcontains $id)
        }

        # THE PENDING LEDGER STATE, READ DIRECTLY. This is the check that
        # "P5-LDG has not run yet" makes necessary rather than sufficient.
        #
        # The accepted Phase-5 guard deliberately leaves an INTERMEDIATE state
        # visible to nobody but the ledger: a duplicate result attempt does not
        # overwrite the first result and does not create a FAIL. It is recorded
        # as a violation and converted into a FAIL later, by P5-LDG - which is
        # deferred until after Phase 6. So a scan of recorded results can show
        # every Phase-5 scenario PASS while Phase 5 is ALREADY KNOWN to have a
        # harness-integrity violation.
        #
        # Phase 6 must not produce behavioural evidence on top of that. The
        # authority is Phase 5's own, consumed here and not reimplemented, and
        # P5-LDG still emits its verdict at the accepted point in the lifecycle.
        $phase5LedgerViolations = @(Get-Phase5LedgerViolations)
        $null = Add-Check $list `
            'the Phase-5 result ledger holds no pending duplicate attempts' `
            ($phase5LedgerViolations.Count -eq 0) `
            ('pending Phase-5 duplicate attempts: ' + ($phase5LedgerViolations -join '; '))

        $prerequisiteOk = Test-ChecklistOk $list
        Add-Phase6Result 'P6-PRE' `
            ('Phase-4 (' + $phase4Prerequisite.Count + '/' + $phase4Prerequisite.Count +
             ', ' + ($phase4Deferred -join ', ') + ' deferred to P5-FIN) and Phase-5 ' +
             'evidence is intact') `
            $(if ($prerequisiteOk) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        $prerequisiteOk = $false
        Add-Phase6Result 'P6-PRE' 'Phase-4 and Phase-5 evidence is intact' 'FAIL' `
            (Format-Phase6Err $_)
    }

    if (-not $prerequisiteOk) {
        foreach ($id in (Get-Phase6RequiredScenarioIds)) {
            if (Test-Phase6ResultRecorded -Id $id) { continue }
            Add-Phase6Result $id 'not attempted' 'FAIL' `
                ('the Phase-4/Phase-5 prerequisite evidence is not intact, so this ' +
                 'scenario was not executed. A Phase-6 result on this workbook would ' +
                 'be evidence of nothing, and "not attempted" must be as loud as "failed".')
        }
        # P6-FIN through the GUARDED reporter, like every other Phase-6 result.
        # P6-LDG is emitted by the driver, last of all.
        Add-Phase6Result 'P6-FIN' 'Phase-6 completeness' 'FAIL' `
            'the Step-13 matrix was not executed: its prerequisite evidence is not intact'
        return
    }

    # -------------------------------------------------------------------
    # P6-ART. Runtime artefact identity.
    # -------------------------------------------------------------------
    # TWO COMMITS, AND THEY ARE NOT THE SAME COMMIT.
    #
    #   production baseline    the accepted source this harness is evidence
    #                          infrastructure FOR. A pinned review authority.
    #   runtime harness commit whatever HEAD is on the machine that ran this.
    #
    # Reporting HEAD as the production baseline would conflate them, and HEAD on
    # an authorised runtime tree is always the harness commit, never the
    # baseline. Both are printed, separately and by name.
    #
    # THE BINDING IS BY BLOB IDENTITY, NOT BY MODULE NAME. That the compiled
    # project contains a module called modSimReport says nothing about whose
    # modSimReport it is. Each accepted Phase-6 production module's blob id in
    # the runtime checkout is compared against the SAME PATH at the baseline, so
    # a source that moved is caught even if the name did not. git computes both
    # sides under identical attribute rules, so no line-ending or encoding
    # difference can masquerade as a match or a mismatch.
    #
    # NO GIT MEANS NO PASS. A runtime result with no attributable source
    # revision is weaker evidence, and recording "unknown" while passing would
    # hand that weakness on as though it were strength.
    #
    # The executed .xlsm hash is CAPTURED, never asserted: Excel's own save is
    # not claimed to be byte-reproducible and no control in this project proves
    # that it is.
    try {
        $list = New-Checklist
        $lines = @()
        $baseline = Get-Phase6ProductionBaseline
        $lines += ('production baseline commit    : ' + $baseline)
        $lines += ('runtime harness commit (HEAD) : ' + $HarnessCommit)

        $gitOk = (-not [string]::IsNullOrWhiteSpace($HarnessCommit)) -and
                 ($HarnessCommit -match '^[0-9a-f]{7,40}$')
        $null = Add-Check $list 'the runtime harness commit was read from git, not guessed' `
            $gitOk ([string]$HarnessCommit)
        $null = Add-Check $list 'the harness commit is not being reported as the production baseline' `
            ($HarnessCommit -notlike ($baseline + '*')) `
            ('baseline ' + $baseline + ', HEAD ' + [string]$HarnessCommit)

        if ($gitOk) {
            # THE WHOLE-TREE STATEMENT, from git itself, AND FROM THE REPOSITORY
            # ROOT.
            #
            # `git -C <path>` runs as though git had been started in <path>, and
            # a `git diff` pathspec is resolved relative to that directory. The
            # first version ran with -C <repo>/pccm and pathspecs `pccm/src` and
            # `pccm/spec`, which resolve to <repo>/pccm/pccm/src - a path that
            # matches nothing. A pathspec matching nothing produces no diff, so
            # `--quiet` exited 0 and the check passed whatever the tree held.
            # That is fail-open, and on the one statement the whole freeze claim
            # rests on. The pathspecs are repository-root relative, so the
            # working directory must be the repository root.
            $null = & git -C $RepoRoot diff --quiet $baseline -- 'pccm/src' 'pccm/spec' 2>$null
            $treeClean = ($LASTEXITCODE -eq 0)
            $null = Add-Check $list `
                ('the complete pccm/src and pccm/spec trees are unchanged from ' + $baseline) `
                $treeClean ('git diff --quiet exit code ' + [string]$LASTEXITCODE)
            # AND THE PATHSPEC REALLY MATCHED SOMETHING. A freeze proved by a
            # pathspec that names nothing is not a freeze; this asks git how many
            # files the pathspec covers at the baseline.
            $tracked = @(& git -C $RepoRoot ls-tree -r --name-only $baseline -- 'pccm/src' 'pccm/spec' 2>$null)
            $null = Add-Check $list `
                'the freeze pathspec matches the production trees it names' `
                ($tracked.Count -gt 0) ('files under the pathspec at the baseline: ' + $tracked.Count)

            # AND THE PER-MODULE STATEMENT, blob for blob.
            foreach ($module in (Get-Phase6ProductionModules)) {
                $relative = 'pccm/src/vba/' + $module + '.bas'
                $accepted = ''
                $current = ''
                try {
                    $accepted = [string](& git -C $RepoRoot rev-parse ($baseline + ':' + $relative) 2>$null)
                    $current = [string](& git -C $RepoRoot rev-parse ('HEAD:' + $relative) 2>$null)
                } catch {
                    $accepted = ''
                    $current = ''
                }
                $null = Add-Check $list `
                    ('the ' + $module + ' source being exercised is the accepted baseline source') `
                    ((-not [string]::IsNullOrWhiteSpace($accepted)) -and ($current -ceq $accepted)) `
                    ('baseline blob ' + $accepted + ', runtime blob ' + $current)
                $lines += ('  ' + $module.PadRight(32) + ' blob ' + $current)
            }
        } else {
            $null = Add-Check $list `
                'the accepted production source can be bound to the baseline' $false `
                ('git is not available on this machine, so no runtime result here is ' +
                 'attributable to a source revision')
        }

        # EVERY ARTEFACT IS HASHED WHERE IT WAS ACTUALLY CONSUMED - the
        # disposable copy the Phase-4 harness made and the bootstrap read - not
        # in build/, so the hashes describe this run rather than the directory
        # it was seeded from.
        foreach ($item in @(
            @{ Label = 'Stage-A workbook (build input)';
               Path = (Join-Path $TempRoot ([string]$Manifest.stage_a_filename)) },
            @{ Label = 'executed .xlsm (this session)';
               Path = (Join-Path $TempRoot ([string]$Manifest.stage_b_filename)) })) {
            if (Test-Path -LiteralPath $item.Path) {
                $hash = (Get-FileHash -LiteralPath $item.Path -Algorithm SHA256).Hash
                $lines += ($item.Label + ' : ' + $hash)
            } else {
                $lines += ($item.Label + ' : NOT FOUND at ' + $item.Path)
            }
        }
        foreach ($item in @(
            @{ Label = 'stage_b_manifest.json         '; Name = 'stage_b_manifest.json' },
            @{ Label = 'phase6_gate_b_inspection.json '; Name = 'phase6_gate_b_inspection.json' },
            @{ Label = 'phase6_gate_b_cases.json      '; Name = 'phase6_gate_b_cases.json' })) {
            $path = Join-Path $TempRoot $item.Name
            if (Test-Path -LiteralPath $path) {
                $lines += ($item.Label + ': ' + (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash)
            } else {
                $lines += ($item.Label + ': NOT FOUND')
            }
        }
        $null = Add-Check $list 'the inspection projection states its provenance' `
            (-not [string]::IsNullOrWhiteSpace([string]$SimInspection.provenance.sim_contract_version)) `
            ('sim_contract ' + [string]$SimInspection.provenance.sim_contract_version)
        $null = Add-Check $list 'the parity corpus and the inspection projection agree on the contract' `
            ([string]$GateBCases.sim_contract_version -ceq
             [string]$SimInspection.provenance.sim_contract_version) `
            ('corpus ' + [string]$GateBCases.sim_contract_version + ', projection ' +
             [string]$SimInspection.provenance.sim_contract_version)
        $null = Add-Check $list 'the parity corpus and the manifest agree on the model version' `
            ([string]$GateBCases.model_version -ceq [string]$Manifest.model_version) `
            ('corpus ' + [string]$GateBCases.model_version + ', manifest ' +
             [string]$Manifest.model_version)
        Add-Phase6Result 'P6-ART' 'Runtime artefact identity' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            ((Format-Checklist $list) + "`r`n" + ($lines -join "`r`n"))
    } catch {
        Add-Phase6Result 'P6-ART' 'Runtime artefact identity' 'FAIL' (Format-Phase6Err $_)
    }

    # -------------------------------------------------------------------
    # P6-CMP and P6-M. DERIVED, not re-executed.
    # -------------------------------------------------------------------
    # P5-CMP is the one actual VBE compile execution of this run and P5-M is the
    # one actual complete module inventory. Running either again to obtain a
    # Phase-6 label would report a second compile that did not happen. These two
    # scenarios therefore assert what those executions establish FOR PHASE 6 -
    # that the project which compiled, and the inventory which was proved, are
    # the ones containing the eight accepted modSim* modules - and say plainly
    # that the execution belongs to the Phase-5 scenario.
    try {
        $list = New-Checklist
        $cmp = @($Results | Where-Object { $_.Id -eq 'P5-CMP' })
        $null = Add-Check $list 'P5-CMP executed a VBE compile of the persisted project' `
            ($cmp.Count -eq 1) ('P5-CMP results: ' + $cmp.Count)
        $null = Add-Check $list 'that compile PASSED' `
            (($cmp.Count -eq 1) -and ($cmp[0].Status -eq 'PASS')) `
            $(if ($cmp.Count -eq 1) { 'status ' + $cmp[0].Status } else { 'no P5-CMP result' })
        $modules = @($Manifest.vba.modules | ForEach-Object { [string]$_.name })
        foreach ($name in @('modSimContract', 'modSimRng', 'modSimSample', 'modSimEngine',
                            'modSimStats', 'modSimFingerprint', 'modSimNonce', 'modSimReport')) {
            $null = Add-Check $list ('the compiled project includes ' + $name) `
                ($modules -contains $name) ('manifest modules: ' + $modules.Count)
        }
        Add-Phase6Result 'P6-CMP' `
            'The compiled project includes the Phase-6 modules (derived from P5-CMP)' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            ((Format-Checklist $list) + "`r`n" +
             '    DERIVED. The compile itself is P5-CMP; no second compile was performed.')
    } catch {
        Add-Phase6Result 'P6-CMP' 'The compiled project includes the Phase-6 modules' 'FAIL' `
            (Format-Phase6Err $_)
    }

    try {
        $list = New-Checklist
        $inventory = @($Results | Where-Object { $_.Id -eq 'P5-M' })
        $null = Add-Check $list 'P5-M executed a complete VBComponent inventory' `
            ($inventory.Count -eq 1) ('P5-M results: ' + $inventory.Count)
        $null = Add-Check $list 'that inventory PASSED' `
            (($inventory.Count -eq 1) -and ($inventory[0].Status -eq 'PASS')) `
            $(if ($inventory.Count -eq 1) { 'status ' + $inventory[0].Status } else { 'no P5-M result' })
        $modules = @($Manifest.vba.modules | ForEach-Object { [string]$_.name })
        $null = Add-Check $list 'the inventory set is the manifest module set, Phase 6 included' `
            ((@($modules | Where-Object { $_ -like 'modSim*' }).Count -eq 8) -and
             ($modules.Count -eq (@($modules | Sort-Object -Unique)).Count)) `
            ('modSim* modules: ' + (@($modules | Where-Object { $_ -like 'modSim*' }) -join ', '))
        Add-Phase6Result 'P6-M' `
            'The persisted project inventory covers the Phase-6 modules (derived from P5-M)' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            ((Format-Checklist $list) + "`r`n" +
             '    DERIVED. The inventory itself is P5-M; no second inventory was performed.')
    } catch {
        Add-Phase6Result 'P6-M' 'The persisted project inventory covers the Phase-6 modules' `
            'FAIL' (Format-Phase6Err $_)
    }

    # -------------------------------------------------------------------
    # P6-API. Declared for seven, callable for six.
    # -------------------------------------------------------------------
    # THREE KINDS OF EVIDENCE, AND THEY ARE NOT INTERCHANGEABLE - the same
    # distinction P5-M was corrected to make after Runtime Run 7:
    #
    #   A  DECLARATION  the name exists in the persisted VBA project
    #   B  CALLABILITY  Application.Run reached it and it answered
    #   C  EXECUTION    it ran against a valid fixture and did its work
    #
    # PCCM_RunSimulation gets A here and nothing else. It is stateful - it
    # allocates a run id, spends a random sequence and publishes a bank - so
    # invoking it at inventory time would establish a publication no scenario
    # asked for. Its first execution is P6-FX1, and no callability pass is
    # recorded on its behalf here.
    try {
        $list = New-Checklist
        $surface = $SimInspection.command_surface
        $endpoint = [string]$surface.automation_endpoint
        $accessors = @($surface.read_accessors)
        $declared = @(Get-Phase5ProjectProcedureNames -Workbook $Workbook)
        $null = Add-Check $list 'the persisted project could be read for declared procedures' `
            ($declared.Count -gt 0) ('declared procedures found: ' + $declared.Count)
        $null = Add-Check $list 'the contract settles six read accessors' `
            ($accessors.Count -eq 6) ('projected ' + $accessors.Count)
        foreach ($name in (@($endpoint) + $accessors)) {
            $null = Add-Check $list ('the Phase-6 procedure ' + $name + ' is declared') `
                ($declared -contains $name) `
                ('declared PCCM_*: ' + ((@($declared | Where-Object { $_ -like 'PCCM_*' })) -join ', '))
        }
        foreach ($name in $accessors) {
            $answered = $false
            $note = ''
            try {
                $value = $Excel.Run($name)
                $answered = $true
                $note = 'returned ' + (Format-SimValue $value)
            } catch {
                $note = Format-Phase6Err $_
            }
            $null = Add-Check $list ('the read accessor ' + $name + ' is callable') `
                $answered $note
        }
        Add-Note ('P6-API: ' + $endpoint + ' is declared; it is stateful, so its runtime ' +
                  'execution is deferred to P6-FX1, which is the first PCCM_RunSimulation ' +
                  'of the run. P6-API records no callability evidence for it.')
        Add-Phase6Result 'P6-API' 'The Phase-6 public surface' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Phase6Result 'P6-API' 'The Phase-6 public surface' 'FAIL' (Format-Phase6Err $_)
    }

    # -------------------------------------------------------------------
    # P6-BTN. Phase 6 introduces no user-facing run control.
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $endpoint = [string]$SimInspection.command_surface.automation_endpoint
        $onActions = @()
        foreach ($button in @($Manifest.buttons)) {
            $sheets = $null; $ws = $null; $shapes = $null; $shp = $null
            try {
                $sheets = $Workbook.Worksheets
                $ws = $sheets.Item([string]$button.sheet)
                $shapes = $ws.Shapes
                $shp = $shapes.Item([string]$button.shape_name)
                $onActions += [string]$shp.OnAction
            } finally {
                if ($null -ne $shp)    { Release-Transient $shp    'Shape';      $shp    = $null }
                if ($null -ne $shapes) { Release-Transient $shapes 'Shapes';     $shapes = $null }
                if ($null -ne $ws)     { Release-Transient $ws     'Worksheet';  $ws     = $null }
                if ($null -ne $sheets) { Release-Transient $sheets 'Worksheets'; $sheets = $null }
            }
        }
        $null = Add-Check $list ('NO shape has OnAction = ' + $endpoint) `
            ($onActions -notcontains $endpoint) ($onActions -join ', ')
        foreach ($name in @($SimInspection.command_surface.read_accessors)) {
            $null = Add-Check $list ('NO shape has OnAction = ' + $name) `
                ($onActions -notcontains $name) ($onActions -join ', ')
        }
        $null = Add-Check $list 'the button count is unchanged from the manifest' `
            ($onActions.Count -eq @($Manifest.buttons).Count) `
            ('found ' + $onActions.Count + ' of ' + @($Manifest.buttons).Count)
        Add-Phase6Result 'P6-BTN' 'Phase 6 adds no user-facing run control' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Phase6Result 'P6-BTN' 'Phase 6 adds no user-facing run control' 'FAIL' `
            (Format-Phase6Err $_)
    }

    # -------------------------------------------------------------------
    # P6-INIT. A workbook that has never simulated.
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $state = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
        $null = Add-Check $list 'the active bank is blank' `
            (Test-SimBlank -Value $state['shared']['active_bank']) `
            (Format-SimValue $state['shared']['active_bank'])
        $null = Add-Check $list 'the pending AUTO nonce sidecar is blank' `
            (Test-SimBlank -Value $state['pending_auto_nonce']) `
            (Format-SimValue $state['pending_auto_nonce'])
        $null = Add-Check $list 'the AUTO nonce counter is at its contract initial' `
            (Test-SimExactDouble -Actual $state['shared']['next_auto_nonce'] `
                -Expected ([double]$bounds.nonce_initial)) `
            ('expected ' + [string]$bounds.nonce_initial + ', found ' +
             (Format-SimValue $state['shared']['next_auto_nonce']))
        $null = Add-Check $list 'the run-id counter is at its contract initial' `
            (Test-SimExactDouble -Actual $state['shared']['last_run_id'] `
                -Expected ([double]$bounds.run_id_initial)) `
            ('expected ' + [string]$bounds.run_id_initial + ', found ' +
             (Format-SimValue $state['shared']['last_run_id']))
        $null = Add-Check $list 'the last attempt result is the never-attempted token' `
            (Test-SimExactText -Actual $state['shared']['last_attempt_result'] `
                -Expected ([string]$vocabulary.attempt_results[0])) `
            ('expected ' + [string]$vocabulary.attempt_results[0] + ', found ' +
             (Format-SimValue $state['shared']['last_attempt_result']))
        foreach ($bank in $banks) {
            $null = Add-Check $list ('bank ' + $bank + ' carries no request fingerprint') `
                (Test-SimBlank -Value $state[('bank_' + $bank)]['request_fingerprint']) `
                (Format-SimValue $state[('bank_' + $bank)]['request_fingerprint'])
        }
        Add-Phase6Result 'P6-INIT' 'A workbook that has never simulated' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            ((Format-Checklist $list) + "`r`n" + (Format-Phase6State -State $state -Label 'state'))
    } catch {
        Add-Phase6Result 'P6-INIT' 'A workbook that has never simulated' 'FAIL' `
            (Format-Phase6Err $_)
    }

    # -------------------------------------------------------------------
    # P6-FX1. The first real PCCM_RunSimulation.
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $null = Set-Phase6Fixture -PlanCaseId ([int]$goldenCase.plan_case_id) `
            -SuppliedSeed ([double]$GateBCases.supplied_seed) `
            -Iterations ([int]$GateBCases.iterations)
        $before = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
        $activeBefore = Get-Phase6ActiveBank -State $before
        $target = Get-Phase6CandidateTarget -Inspection $SimInspection -ActiveBank $activeBefore
        $announced = Invoke-Phase6Simulation -Excel $Excel
        $after = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection

        $null = Add-Check $list 'the endpoint announced success' `
            (Test-Phase6Announced -Result $announced -Kind 'OK') $announced
        $null = Add-Check $list ('the active bank moved to the candidate target ' + $target) `
            ((Get-Phase6ActiveBank -State $after) -ceq $target) `
            ('was ' + [char]39 + $activeBefore + [char]39 + ', now ' + [char]39 +
             (Get-Phase6ActiveBank -State $after) + [char]39)
        $null = Add-Check $list 'the run id advanced to the contract first successful value' `
            (Test-SimExactDouble -Actual $after['shared']['last_run_id'] `
                -Expected ([double]$bounds.run_id_first_successful_value)) `
            (Format-SimValue $after['shared']['last_run_id'])
        $null = Add-Check $list 'the attempt result is SUCCESS' `
            (Test-SimExactText -Actual $after['shared']['last_attempt_result'] -Expected 'SUCCESS') `
            (Format-SimValue $after['shared']['last_attempt_result'])
        $null = Add-Check $list 'the published seed mode is FIXED' `
            (Test-SimExactText -Actual $after[('bank_' + $target)]['seed_mode'] -Expected 'FIXED') `
            (Format-SimValue $after[('bank_' + $target)]['seed_mode'])
        $null = Add-Check $list 'the effective seed is the supplied seed' `
            (Test-SimExactDouble -Actual $after[('bank_' + $target)]['effective_seed'] `
                -Expected ([double]$GateBCases.supplied_seed)) `
            (Format-SimValue $after[('bank_' + $target)]['effective_seed'])
        $null = Add-Check $list 'the iterations run is the requested count' `
            (Test-SimExactDouble -Actual $after[('bank_' + $target)]['iterations_run'] `
                -Expected ([double]$GateBCases.iterations)) `
            (Format-SimValue $after[('bank_' + $target)]['iterations_run'])
        $null = Add-Check $list 'the first published iteration row carries a number' `
            (((Get-SimIterationValue -Workbook $Workbook -Inspection $SimInspection `
                -Bank $target -ColumnKey 'total_nominal' -Offset 0)).GetType().FullName -ceq
             'System.Double') `
            (Format-SimValue (Get-SimIterationValue -Workbook $Workbook -Inspection $SimInspection `
                -Bank $target -ColumnKey 'total_nominal' -Offset 0))

        Add-Phase6Result 'P6-FX1' 'FIXED-seed execution publishes a simulation result' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            ((Format-Checklist $list) + "`r`n" +
             (Format-Phase6State -State $before -Label 'before') + "`r`n" +
             '    action: PCCM_RunSimulation, FIXED seed ' + [string]$GateBCases.supplied_seed +
             ', ' + [string]$GateBCases.iterations + ' iterations' + "`r`n" +
             '    announcement: ' + $announced + "`r`n" +
             (Format-Phase6State -State $after -Label 'after'))
    } catch {
        Add-Phase6Result 'P6-FX1' 'FIXED-seed execution publishes a simulation result' 'FAIL' `
            (Format-Phase6Err $_)
    }

    # -------------------------------------------------------------------
    # P6-ORA. Cross-implementation parity, every authorised case.
    # -------------------------------------------------------------------
    # THE QUESTION THIS ANSWERS, AND NO OTHER SCENARIO DOES: does real Excel
    # produce the number the accepted Python oracle produces for the SAME model?
    # Each case names an EXISTING Phase-5 plan case, driven through the accepted
    # fixture machinery, so the two implementations are not describing "similar"
    # fixtures. If Excel and the oracle disagree, this FAILS and the exact
    # disagreement is the evidence.
    try {
        $list = New-Checklist
        $evidence = @()
        foreach ($case in $parityCases) {
            $label = 'plan case ' + [string]$case.plan_case_id +
                     ' (' + [string]$case.sampling_mechanism + ')'
            $null = Set-Phase6Fixture -PlanCaseId ([int]$case.plan_case_id) `
                -SuppliedSeed ([double]$GateBCases.supplied_seed) `
                -Iterations ([int]$GateBCases.iterations)

            # THE CURRENT ANALYTICAL IDENTITY FIRST, FOR EVERY CASE.
            #
            # Comparing a simulation before proving WHICH model produced it
            # compares two different questions. An earlier scenario having driven
            # the same plan case is not that proof: what has to be established is
            # that the workbook, as it stands right now and immediately before
            # this simulation, holds the model whose oracle result is about to be
            # compared.
            #
            # THE ACCEPTED PHASE-5 MACHINERY DOES THIS ALREADY, so it is reused
            # rather than reimplemented. `Add-Phase5AnalyticalChecks` compares the
            # CURRENT `_Calc` snapshot against that plan case's own emitted
            # expectations, and `Add-Phase5SuccessStateChecks` compares the
            # committed calc_state record. Using them is not a second fingerprint
            # implementation - it is exactly the current-fixture identity check
            # P5-AN is already trusted for.
            $planCase = Get-PlanCase -Id ([int]$case.plan_case_id)
            $calcAttempt = [string]$Excel.Run('PCCM_CalculationAttemptResult')
            $calcStatus = [string]$Excel.Run('PCCM_CalculationStatus')
            $calcStored = [string]$Excel.Run('PCCM_CalculationFingerprint')
            $calcCurrent = [string]$Excel.Run('PCCM_CurrentInputFingerprint')
            $null = Add-Check $list ($label + ': the fixture calculation succeeded') `
                ($calcAttempt -ceq 'SUCCESS') ([char]39 + $calcAttempt + [char]39)
            $null = Add-Check $list ($label + ': the calculation status is CURRENT') `
                ($calcStatus -ceq 'CURRENT') ([char]39 + $calcStatus + [char]39)
            $null = Add-Check $list `
                ($label + ': the stored analytical fingerprint IS the current one') `
                ((-not [string]::IsNullOrEmpty($calcCurrent)) -and ($calcStored -ceq $calcCurrent)) `
                ('stored ' + [char]39 + $calcStored + [char]39 + ', current ' +
                 [char]39 + $calcCurrent + [char]39)
            Add-Phase5AnalyticalChecks -List $list -Workbook $Workbook `
                -Inspection $Inspection -Case $planCase -Tolerances $Cases.tolerances `
                -Label ($label + ' analytical')
            Add-Phase5SuccessStateChecks -List $list -Excel $Excel -Workbook $Workbook `
                -Inspection $Inspection -Case $planCase -Cases $Cases `
                -Label ($label + ' calc_state')
            if ([bool]$case.analytical_identity.fingerprint_independently_derivable) {
                # AND, FOR THE ONE CASE WHERE IT IS DERIVABLE, the independent
                # reference digest as well. This is additional to the checks
                # above, never a substitute for them.
                $null = Add-Check $list `
                    ($label + ': the calculation fingerprint equals the accepted reference') `
                    ($calcStored -ceq [string]$case.expected_exact.calculation_fingerprint) `
                    ('oracle ' + [string]$case.expected_exact.calculation_fingerprint +
                     ', workbook ' + [char]39 + $calcStored + [char]39)
            }

            $activeBefore = Get-Phase6ActiveBank `
                -State (Get-Phase6State -Workbook $Workbook -Inspection $SimInspection)
            $target = Get-Phase6CandidateTarget -Inspection $SimInspection -ActiveBank $activeBefore
            $announced = Invoke-Phase6Simulation -Excel $Excel
            $null = Add-Check $list ($label + ': the endpoint announced success') `
                (Test-Phase6Announced -Result $announced -Kind 'OK') $announced

            Add-Phase6ParityChecks -Workbook $Workbook -Inspection $SimInspection `
                -Cases $GateBCases -Case $case -List $list -Bank $target -Label $label
            $evidence += ('    ' + $label + ' -> bank ' + $target + ', oracle digest ' +
                          [string]$case.expected_exact.result_digest + ', workbook digest ' +
                          (Format-SimValue (Get-SimField -Workbook $Workbook `
                              -Inspection $SimInspection -FieldKey 'result_digest' -Bank $target)))
        }
        Add-Phase6Result 'P6-ORA' 'Cross-implementation parity with the accepted oracle' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            ((Format-Checklist $list) + "`r`n" + ($evidence -join "`r`n"))
    } catch {
        Add-Phase6Result 'P6-ORA' 'Cross-implementation parity with the accepted oracle' 'FAIL' `
            (Format-Phase6Err $_)
    }

    # -------------------------------------------------------------------
    # P6-DET. The same inputs and the same FIXED seed, twice.
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $null = Set-Phase6Fixture -PlanCaseId ([int]$goldenCase.plan_case_id) `
            -SuppliedSeed ([double]$GateBCases.supplied_seed) `
            -Iterations ([int]$GateBCases.iterations)
        $digests = @()
        $announcements = @()
        for ($pass = 1; $pass -le 2; $pass++) {
            $activeBefore = Get-Phase6ActiveBank `
                -State (Get-Phase6State -Workbook $Workbook -Inspection $SimInspection)
            $target = Get-Phase6CandidateTarget -Inspection $SimInspection -ActiveBank $activeBefore
            $announcements += (Invoke-Phase6Simulation -Excel $Excel)
            $digests += (Get-SimField -Workbook $Workbook -Inspection $SimInspection `
                -FieldKey 'result_digest' -Bank $target)
        }
        foreach ($announced in $announcements) {
            $null = Add-Check $list 'the repeated run announced success' `
                (Test-Phase6Announced -Result $announced -Kind 'OK') $announced
        }
        $null = Add-Check $list 'both runs published the same result digest' `
            (Test-SimSameValue -A $digests[0] -B $digests[1]) `
            ((Format-SimValue $digests[0]) + ' vs ' + (Format-SimValue $digests[1]))
        $null = Add-Check $list 'and it is the digest the oracle predicted' `
            (Test-SimExactText -Actual $digests[1] `
                -Expected ([string]$goldenCase.expected_exact.result_digest)) `
            ('oracle ' + [string]$goldenCase.expected_exact.result_digest)
        Add-Phase6Result 'P6-DET' 'Repeatability: same inputs, same FIXED seed, same digest' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Phase6Result 'P6-DET' 'Repeatability: same inputs, same FIXED seed, same digest' `
            'FAIL' (Format-Phase6Err $_)
    }

    # -------------------------------------------------------------------
    # P6-FIXED-INERT. FIXED touches nothing on the AUTO axis.
    # -------------------------------------------------------------------
    # The armed failpoint is the load-bearing part. Phase6AfterNoncePersisted
    # names a PERSISTED AUTO ADVANCE; in FIXED mode that boundary does not
    # exist, so a FIXED run with it armed must still SUCCEED. A run that failed
    # here would mean the failpoint had been placed on a path FIXED reaches.
    try {
        $list = New-Checklist
        $null = Set-Phase6Fixture -PlanCaseId ([int]$goldenCase.plan_case_id) `
            -SuppliedSeed ([double]$GateBCases.supplied_seed) `
            -Iterations ([int]$GateBCases.iterations)
        $before = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
        $announced = Invoke-Phase6Simulation -Excel $Excel `
            -FailAfterStage ([string]$failpoints.AfterNoncePersisted)
        $after = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection

        $null = Add-Check $list ('a FIXED run succeeds with ' +
            [string]$failpoints.AfterNoncePersisted + ' armed') `
            (Test-Phase6Announced -Result $announced -Kind 'OK') $announced
        $null = Add-Check $list 'the AUTO nonce counter is unchanged' `
            (Test-SimSameValue -A $before['shared']['next_auto_nonce'] `
                -B $after['shared']['next_auto_nonce']) `
            ('was ' + (Format-SimValue $before['shared']['next_auto_nonce']) + ', now ' +
             (Format-SimValue $after['shared']['next_auto_nonce']))
        $null = Add-Check $list 'the pending AUTO nonce sidecar is unchanged' `
            (Test-SimSameValue -A $before['pending_auto_nonce'] -B $after['pending_auto_nonce']) `
            ('was ' + (Format-SimValue $before['pending_auto_nonce']) + ', now ' +
             (Format-SimValue $after['pending_auto_nonce']))
        $null = Add-Check $list 'and the sidecar is blank' `
            (Test-SimBlank -Value $after['pending_auto_nonce']) `
            (Format-SimValue $after['pending_auto_nonce'])
        $null = Add-Check $list 'no AUTO nonce was recorded as consumed' `
            (Test-SimBlank -Value $after[(
                'bank_' + (Get-Phase6ActiveBank -State $after))]['consumed_auto_nonce']) `
            (Format-SimValue $after[(
                'bank_' + (Get-Phase6ActiveBank -State $after))]['consumed_auto_nonce'])
        Add-Phase6Result 'P6-FIXED-INERT' 'FIXED mode never touches the AUTO axis' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            ((Format-Checklist $list) + "`r`n" +
             (Format-Phase6State -State $before -Label 'before') + "`r`n" +
             '    announcement: ' + $announced + "`r`n" +
             (Format-Phase6State -State $after -Label 'after'))
    } catch {
        Add-Phase6Result 'P6-FIXED-INERT' 'FIXED mode never touches the AUTO axis' 'FAIL' `
            (Format-Phase6Err $_)
    }

    # -------------------------------------------------------------------
    # P6-AU1 / P6-AU2. AUTO allocation, and no replay.
    # -------------------------------------------------------------------
    $autoFirst = $null
    try {
        $list = New-Checklist
        $null = Set-Phase6Fixture -PlanCaseId ([int]$goldenCase.plan_case_id) `
            -SuppliedSeed $null -Iterations ([int]$GateBCases.iterations)
        $before = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
        $counterBefore = [double]$before['shared']['next_auto_nonce']
        $activeBefore = Get-Phase6ActiveBank -State $before
        $target = Get-Phase6CandidateTarget -Inspection $SimInspection -ActiveBank $activeBefore
        $announced = Invoke-Phase6Simulation -Excel $Excel
        $after = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection

        $null = Add-Check $list 'the AUTO run announced success' `
            (Test-Phase6Announced -Result $announced -Kind 'OK') $announced
        $null = Add-Check $list 'the published seed mode is AUTO' `
            (Test-SimExactText -Actual $after[('bank_' + $target)]['seed_mode'] -Expected 'AUTO') `
            (Format-SimValue $after[('bank_' + $target)]['seed_mode'])
        $null = Add-Check $list 'the consumed AUTO nonce is the counter value the run started from' `
            (Test-SimExactDouble -Actual $after[('bank_' + $target)]['consumed_auto_nonce'] `
                -Expected $counterBefore) `
            ('counter before ' + [string]$counterBefore + ', consumed ' +
             (Format-SimValue $after[('bank_' + $target)]['consumed_auto_nonce']))
        $null = Add-Check $list 'the counter advanced by exactly one' `
            (Test-SimExactDouble -Actual $after['shared']['next_auto_nonce'] `
                -Expected ($counterBefore + 1)) `
            ('was ' + [string]$counterBefore + ', now ' +
             (Format-SimValue $after['shared']['next_auto_nonce']))
        $null = Add-Check $list 'the pending AUTO nonce sidecar is clear after a clean success' `
            (Test-SimBlank -Value $after['pending_auto_nonce']) `
            (Format-SimValue $after['pending_auto_nonce'])
        $null = Add-Check $list 'no supplied seed was published' `
            (Test-SimBlank -Value $after[('bank_' + $target)]['supplied_seed']) `
            (Format-SimValue $after[('bank_' + $target)]['supplied_seed'])
        $null = Add-Check $list 'an effective seed was published' `
            (($after[('bank_' + $target)]['effective_seed']).GetType().FullName -ceq 'System.Double') `
            (Format-SimValue $after[('bank_' + $target)]['effective_seed'])

        $autoFirst = [pscustomobject]@{
            Nonce         = $after[('bank_' + $target)]['consumed_auto_nonce']
            EffectiveSeed = $after[('bank_' + $target)]['effective_seed']
            Digest        = $after[('bank_' + $target)]['result_digest']
            CounterAfter  = $after['shared']['next_auto_nonce']
        }
        Add-Phase6Result 'P6-AU1' 'AUTO-seed execution consumes and persists one nonce' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            ((Format-Checklist $list) + "`r`n" +
             (Format-Phase6State -State $before -Label 'before') + "`r`n" +
             '    announcement: ' + $announced + "`r`n" +
             (Format-Phase6State -State $after -Label 'after'))
    } catch {
        Add-Phase6Result 'P6-AU1' 'AUTO-seed execution consumes and persists one nonce' 'FAIL' `
            (Format-Phase6Err $_)
    }

    try {
        $list = New-Checklist
        $before = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
        $counterBefore = [double]$before['shared']['next_auto_nonce']
        $target = Get-Phase6CandidateTarget -Inspection $SimInspection `
            -ActiveBank (Get-Phase6ActiveBank -State $before)
        $announced = Invoke-Phase6Simulation -Excel $Excel
        $after = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection

        # THE NO-REPLAY INVARIANT, STATED AS THE CONTRACT STATES IT. The claim
        # is that the second run takes the NEXT nonce and that the two runs have
        # distinct AUTO identities - NOT that two different seeds must produce
        # two different digests. A degenerate model can legitimately produce
        # identical retained totals under two seeds, so digest inequality is
        # recorded as evidence and is never a pass criterion here.
        $null = Add-Check $list 'the second AUTO run announced success' `
            (Test-Phase6Announced -Result $announced -Kind 'OK') $announced
        $null = Add-Check $list 'it consumed the next nonce, not the first one again' `
            (Test-SimExactDouble -Actual $after[('bank_' + $target)]['consumed_auto_nonce'] `
                -Expected $counterBefore) `
            ('counter before ' + [string]$counterBefore + ', consumed ' +
             (Format-SimValue $after[('bank_' + $target)]['consumed_auto_nonce']))
        if ($null -ne $autoFirst) {
            $null = Add-Check $list 'the two runs consumed different AUTO nonces' `
                (-not (Test-SimSameValue -A $autoFirst.Nonce `
                    -B $after[('bank_' + $target)]['consumed_auto_nonce'])) `
                ((Format-SimValue $autoFirst.Nonce) + ' then ' +
                 (Format-SimValue $after[('bank_' + $target)]['consumed_auto_nonce']))
            $null = Add-Check $list 'and derived different effective seeds from them' `
                (-not (Test-SimSameValue -A $autoFirst.EffectiveSeed `
                    -B $after[('bank_' + $target)]['effective_seed'])) `
                ((Format-SimValue $autoFirst.EffectiveSeed) + ' then ' +
                 (Format-SimValue $after[('bank_' + $target)]['effective_seed']))
            Add-Note ('P6-AU2: digests for the two AUTO identities were ' +
                      (Format-SimValue $autoFirst.Digest) + ' and ' +
                      (Format-SimValue $after[('bank_' + $target)]['result_digest']) +
                      '. Recorded as evidence; digest inequality is not a contract rule.')
        }
        $null = Add-Check $list 'the counter advanced by exactly one again' `
            (Test-SimExactDouble -Actual $after['shared']['next_auto_nonce'] `
                -Expected ($counterBefore + 1)) `
            (Format-SimValue $after['shared']['next_auto_nonce'])
        $null = Add-Check $list 'the sidecar is clear' `
            (Test-SimBlank -Value $after['pending_auto_nonce']) `
            (Format-SimValue $after['pending_auto_nonce'])
        Add-Phase6Result 'P6-AU2' 'An AUTO retry takes the next nonce, never a consumed one' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            ((Format-Checklist $list) + "`r`n" +
             (Format-Phase6State -State $before -Label 'before') + "`r`n" +
             '    announcement: ' + $announced + "`r`n" +
             (Format-Phase6State -State $after -Label 'after'))
    } catch {
        Add-Phase6Result 'P6-AU2' 'An AUTO retry takes the next nonce, never a consumed one' `
            'FAIL' (Format-Phase6Err $_)
    }

    # -------------------------------------------------------------------
    # P6-BANK. Dual-bank publication, as an AGGREGATE of observable states.
    # -------------------------------------------------------------------
    # WHAT POWERSHELL CAN AND CANNOT SEE. Application.Run is synchronous. No
    # scenario here inspects the workbook while PCCM_RunSimulation is executing,
    # and neither failpoint suspends the transaction for inspection: Phase6FinalCommit
    # fires AFTER the D22:D30 assignment and BEFORE production's own restore, so
    # by the time PowerShell regains control the prior block has already been
    # put back. Any claim of a mid-call observation would be a claim about a
    # moment this harness never occupied.
    #
    # So the conclusion is assembled from three INDEPENDENTLY OBSERVABLE states:
    #
    #   this scenario  a successful run moves the selector to the candidate
    #                  target, the shared commit block agrees with the target
    #                  bank, and the other bank is untouched;
    #   P6-FP2         a failure after the candidate writes leaves the selector
    #                  and the prior publication where they were;
    #   P6-FP3         a failure at the commit boundary returns with the prior
    #                  shared block restored and the prior bank still active.
    #
    # THE ORDERING STATEMENT - that the selector assignment happens only inside
    # FinalCommit, after candidate publication - is SOURCE evidence, proved by
    # the Linux transaction-order controls. It is not observed here.
    try {
        $list = New-Checklist
        $null = Set-Phase6Fixture -PlanCaseId ([int]$goldenCase.plan_case_id) `
            -SuppliedSeed ([double]$GateBCases.supplied_seed) `
            -Iterations ([int]$GateBCases.iterations)
        $before = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
        $activeBefore = Get-Phase6ActiveBank -State $before
        $target = Get-Phase6CandidateTarget -Inspection $SimInspection -ActiveBank $activeBefore
        $other = @($banks | Where-Object { $_ -cne $target })[0]
        $announced = Invoke-Phase6Simulation -Excel $Excel
        $after = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection

        $null = Add-Check $list 'the run announced success' `
            (Test-Phase6Announced -Result $announced -Kind 'OK') $announced
        $null = Add-Check $list ('the selector moved ' + [char]39 + $activeBefore + [char]39 +
            ' -> ' + [char]39 + $target + [char]39 + ', the contract''s candidate target') `
            ((Get-Phase6ActiveBank -State $after) -ceq $target) `
            (Format-SimValue $after['shared']['active_bank'])
        # THE PUBLISHED BANK AND THE SHARED COMMIT BLOCK AGREE. The commit block
        # is one write; if it and the bank it published disagreed, the workbook
        # would be describing two different runs.
        $null = Add-Check $list 'the committed seed mode matches the published bank' `
            (Test-SimSameValue -A $after['shared']['last_attempt_seed_mode'] `
                -B $after[('bank_' + $target)]['seed_mode']) `
            ((Format-SimValue $after['shared']['last_attempt_seed_mode']) + ' vs ' +
             (Format-SimValue $after[('bank_' + $target)]['seed_mode']))
        $null = Add-Check $list 'the committed effective seed matches the published bank' `
            (Test-SimSameValue -A $after['shared']['last_attempt_effective_seed'] `
                -B $after[('bank_' + $target)]['effective_seed']) `
            ((Format-SimValue $after['shared']['last_attempt_effective_seed']) + ' vs ' +
             (Format-SimValue $after[('bank_' + $target)]['effective_seed']))
        $null = Add-Check $list 'the committed run id matches the published bank' `
            (Test-SimSameValue -A $after['shared']['last_run_id'] `
                -B $after[('bank_' + $target)]['run_id']) `
            ((Format-SimValue $after['shared']['last_run_id']) + ' vs ' +
             (Format-SimValue $after[('bank_' + $target)]['run_id']))
        # AND THE BANK THAT WAS NOT THE TARGET DID NOT MOVE.
        Add-SimUnchangedChecks -List $list -Before $before[('bank_' + $other)] `
            -After $after[('bank_' + $other)] -Label ('bank ' + $other)

        Add-Phase6Result 'P6-BANK' 'Dual-bank publication in real Excel' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            ((Format-Checklist $list) + "`r`n" +
             '    NOTE: this scenario observes only states BEFORE and AFTER a completed ' +
             "`r`n" + '    Application.Run. It makes no mid-call observation, and the ' +
             "`r`n" + '    ordering of the selector write inside FinalCommit is SOURCE ' +
             "`r`n" + '    evidence, corroborated at runtime by P6-FP2 and P6-FP3.' + "`r`n" +
             (Format-Phase6State -State $before -Label 'before') + "`r`n" +
             '    announcement: ' + $announced + "`r`n" +
             (Format-Phase6State -State $after -Label 'after'))
    } catch {
        Add-Phase6Result 'P6-BANK' 'Dual-bank publication in real Excel' 'FAIL' `
            (Format-Phase6Err $_)
    }

    # -------------------------------------------------------------------
    # P6-ACC. All six read accessors against the published authority.
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $before = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
        $active = Get-Phase6ActiveBank -State $before

        $status = [string]$Excel.Run('PCCM_SimulationStatus')
        $after = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection

        $null = Add-Check $list 'PCCM_SimulationStatus reports CURRENT for the run just published' `
            ($status -ceq [string]$vocabulary.sim_states[0]) `
            ('expected ' + [string]$vocabulary.sim_states[0] + ', got ' + [char]39 + $status + [char]39)
        $null = Add-Check $list 'and the persisted status cell agrees with what it returned' `
            (Test-SimExactText -Actual $after['shared']['simulation_status'] -Expected $status) `
            (Format-SimValue $after['shared']['simulation_status'])

        # ASKING FOR THE STATUS IS NOT A SIMULATION. Only the two derived rows
        # may move; every other shared row, both banks and the sidecar must be
        # exactly as they were. Proved by capturing the WHOLE block at both ends
        # rather than by reading the source.
        foreach ($key in $before['shared'].Keys) {
            if (($key -eq 'simulation_status') -or ($key -eq 'status_evaluated_at')) { continue }
            $null = Add-Check $list ('the status accessor did not touch ' + $key) `
                (Test-SimSameValue -A $before['shared'][$key] -B $after['shared'][$key]) `
                ('was ' + (Format-SimValue $before['shared'][$key]) + ', now ' +
                 (Format-SimValue $after['shared'][$key]))
        }
        foreach ($bank in $banks) {
            Add-SimUnchangedChecks -List $list -Before $before[('bank_' + $bank)] `
                -After $after[('bank_' + $bank)] -Label ('the status accessor left bank ' + $bank)
        }
        $null = Add-Check $list 'the status accessor did not touch the sidecar' `
            (Test-SimSameValue -A $before['pending_auto_nonce'] -B $after['pending_auto_nonce']) `
            (Format-SimValue $after['pending_auto_nonce'])

        $attemptResult = [string]$Excel.Run('PCCM_SimulationAttemptResult')
        $attemptDetail = [string]$Excel.Run('PCCM_SimulationAttemptDetail')
        $storedFingerprint = [string]$Excel.Run('PCCM_SimulationRequestFingerprint')
        $currentFingerprint = [string]$Excel.Run('PCCM_CurrentSimulationRequestFingerprint')
        $storedDigest = [string]$Excel.Run('PCCM_SimulationResultDigest')

        $null = Add-Check $list 'PCCM_SimulationAttemptResult is SUCCESS' `
            ($attemptResult -ceq 'SUCCESS') ([char]39 + $attemptResult + [char]39)
        $null = Add-Check $list 'and it equals the persisted attempt row' `
            (Test-SimExactText -Actual $after['shared']['last_attempt_result'] -Expected $attemptResult) `
            (Format-SimValue $after['shared']['last_attempt_result'])
        $null = Add-Check $list 'PCCM_SimulationAttemptDetail equals the persisted detail row' `
            (($attemptDetail -ceq [string]$after['shared']['last_attempt_detail']) -or
             ((Test-SimBlank -Value $after['shared']['last_attempt_detail']) -and
              [string]::IsNullOrEmpty($attemptDetail))) `
            ('accessor ' + [char]39 + $attemptDetail + [char]39 + ', cell ' +
             (Format-SimValue $after['shared']['last_attempt_detail']))
        $null = Add-Check $list 'PCCM_SimulationRequestFingerprint is the ACTIVE bank''s stored value' `
            (Test-SimExactText -Actual $after[('bank_' + $active)]['request_fingerprint'] `
                -Expected $storedFingerprint) `
            ('accessor ' + [char]39 + $storedFingerprint + [char]39 + ', bank ' + $active + ' ' +
             (Format-SimValue $after[('bank_' + $active)]['request_fingerprint']))
        $null = Add-Check $list 'PCCM_SimulationResultDigest is the ACTIVE bank''s stored value' `
            (Test-SimExactText -Actual $after[('bank_' + $active)]['result_digest'] `
                -Expected $storedDigest) `
            ('accessor ' + [char]39 + $storedDigest + [char]39 + ', bank ' + $active + ' ' +
             (Format-SimValue $after[('bank_' + $active)]['result_digest']))
        $null = Add-Check $list 'PCCM_CurrentSimulationRequestFingerprint recomputed a value' `
            (-not [string]::IsNullOrEmpty($currentFingerprint)) `
            ([char]39 + $currentFingerprint + [char]39)
        $null = Add-Check $list 'and while the run is CURRENT it equals the stored one' `
            ($currentFingerprint -ceq $storedFingerprint) `
            ('current ' + [char]39 + $currentFingerprint + [char]39 + ', stored ' +
             [char]39 + $storedFingerprint + [char]39)
        $null = Add-Check $list 'the published request fingerprint is the oracle''s for this case' `
            ($storedFingerprint -ceq [string]$goldenCase.expected_exact.request_fingerprint) `
            ('oracle ' + [string]$goldenCase.expected_exact.request_fingerprint)

        Add-Phase6Result 'P6-ACC' 'The six read accessors agree with the published authority' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            ((Format-Checklist $list) + "`r`n" +
             (Format-Phase6State -State $before -Label 'before') + "`r`n" +
             (Format-Phase6State -State $after -Label 'after'))
    } catch {
        Add-Phase6Result 'P6-ACC' 'The six read accessors agree with the published authority' `
            'FAIL' (Format-Phase6Err $_)
    }

    # THE PERSISTED FACTS EACH UNSUCCESSFUL ATTEMPT LEFT BEHIND. Collected as
    # the scenarios run and read back by P6-AXIS, which draws its conclusion
    # from persisted evidence only.
    $axisObservations = New-Object System.Collections.ArrayList

    # -------------------------------------------------------------------
    # P6-RF1 and P6-PRESERVE. A refusal before allocation, and what survives it.
    # -------------------------------------------------------------------
    $preservedBefore = $null
    try {
        $list = New-Checklist
        $preservedBefore = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
        # BELOW THE CONTRACT MINIMUM, read from the corpus rather than spelled.
        Set-NamedValue -Workbook $Workbook `
            -DefinedName ([string]$controls.monte_carlo_iterations.defined_name) `
            -Value ([double]$bounds.business_minimum_iterations - 1)
        $announced = Invoke-Phase6Simulation -Excel $Excel
        $after = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection

        $null = Add-Check $list 'the endpoint announced a failure' `
            (Test-Phase6Announced -Result $announced -Kind 'FAIL') $announced
        $null = Add-Check $list 'the attempt result is REFUSED' `
            (Test-SimExactText -Actual $after['shared']['last_attempt_result'] -Expected 'REFUSED') `
            (Format-SimValue $after['shared']['last_attempt_result'])
        $null = Add-Check $list 'no AUTO nonce was consumed: the counter did not move' `
            (Test-SimSameValue -A $preservedBefore['shared']['next_auto_nonce'] `
                -B $after['shared']['next_auto_nonce']) `
            ('was ' + (Format-SimValue $preservedBefore['shared']['next_auto_nonce']) + ', now ' +
             (Format-SimValue $after['shared']['next_auto_nonce']))
        $null = Add-Check $list 'the sidecar was not touched' `
            (Test-SimSameValue -A $preservedBefore['pending_auto_nonce'] `
                -B $after['pending_auto_nonce']) `
            (Format-SimValue $after['pending_auto_nonce'])
        $null = Add-Check $list 'the run id was not allocated' `
            (Test-SimSameValue -A $preservedBefore['shared']['last_run_id'] `
                -B $after['shared']['last_run_id']) `
            (Format-SimValue $after['shared']['last_run_id'])
        $null = Add-Check $list 'the active bank did not move' `
            (Test-SimSameValue -A $preservedBefore['shared']['active_bank'] `
                -B $after['shared']['active_bank']) `
            (Format-SimValue $after['shared']['active_bank'])

        $null = $axisObservations.Add([pscustomobject]@{
            Scenario     = 'P6-RF1'
            What         = 'refusal before allocation'
            AttemptResult = $after['shared']['last_attempt_result']
            CounterMoved = (-not (Test-SimSameValue `
                -A $preservedBefore['shared']['next_auto_nonce'] `
                -B $after['shared']['next_auto_nonce']))
            AttemptNonce = $after['shared']['last_attempt_auto_nonce']
        })

        Add-Phase6Result 'P6-RF1' 'A prerequisite refusal spends nothing' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            ((Format-Checklist $list) + "`r`n" +
             (Format-Phase6State -State $preservedBefore -Label 'before') + "`r`n" +
             '    action: PCCM_RunSimulation with iterations = ' +
             [string]([double]$bounds.business_minimum_iterations - 1) + "`r`n" +
             '    announcement: ' + $announced + "`r`n" +
             (Format-Phase6State -State $after -Label 'after'))
    } catch {
        Add-Phase6Result 'P6-RF1' 'A prerequisite refusal spends nothing' 'FAIL' `
            (Format-Phase6Err $_)
    }

    try {
        $list = New-Checklist
        if ($null -eq $preservedBefore) {
            throw 'the pre-refusal publication was never captured, so nothing can be compared'
        }
        $after = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
        foreach ($bank in $banks) {
            Add-SimUnchangedChecks -List $list -Before $preservedBefore[('bank_' + $bank)] `
                -After $after[('bank_' + $bank)] -Label ('bank ' + $bank + ' survived the refusal')
        }
        foreach ($key in @('active_bank', 'last_run_id', 'next_auto_nonce')) {
            $null = Add-Check $list ('the shared ' + $key + ' survived the refusal') `
                (Test-SimSameValue -A $preservedBefore['shared'][$key] -B $after['shared'][$key]) `
                ('was ' + (Format-SimValue $preservedBefore['shared'][$key]) + ', now ' +
                 (Format-SimValue $after['shared'][$key]))
        }
        Add-Phase6Result 'P6-PRESERVE' 'A refused attempt destroys no prior publication' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Phase6Result 'P6-PRESERVE' 'A refused attempt destroys no prior publication' 'FAIL' `
            (Format-Phase6Err $_)
    }

    # -------------------------------------------------------------------
    # P6-FP1. Phase6AfterNoncePersisted, AUTO. The two-axis model in Excel.
    # -------------------------------------------------------------------
    # THIS IS THE RUNTIME PROOF OF THE STEP-12 CORRECTION, stated in persisted
    # terms only. The failpoint fires once the advance is KNOWN consumed and the
    # marker is cleared, so the workbook must show, together:
    #
    #   the counter advanced m -> m+1        the nonce really was spent
    #   the sidecar is clear                 the transaction resolved
    #   the attempt records seed and nonce   the identity was selected
    #   the attempt result is REFUSED        not AUTO_NONCE_INDETERMINATE
    #
    # The fifth token belongs to PERSISTENCE_INDETERMINATE and to nothing else,
    # and this run is a definite CONSUMED. `NonceConsumed` is a Private Boolean
    # inside modSimReport; PowerShell cannot observe it and nothing here claims
    # to. What is proved at runtime is the PERSISTED consequence.
    try {
        $list = New-Checklist
        $null = Set-Phase6Fixture -PlanCaseId ([int]$goldenCase.plan_case_id) `
            -SuppliedSeed $null -Iterations ([int]$GateBCases.iterations)
        $before = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
        $counterBefore = [double]$before['shared']['next_auto_nonce']
        $announced = Invoke-Phase6Simulation -Excel $Excel `
            -FailAfterStage ([string]$failpoints.AfterNoncePersisted)
        $after = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection

        $null = Add-Check $list 'the endpoint announced a failure' `
            (Test-Phase6Announced -Result $announced -Kind 'FAIL') $announced
        $null = Add-Check $list 'the counter advanced by one: the nonce WAS consumed' `
            (Test-SimExactDouble -Actual $after['shared']['next_auto_nonce'] `
                -Expected ($counterBefore + 1)) `
            ('was ' + [string]$counterBefore + ', now ' +
             (Format-SimValue $after['shared']['next_auto_nonce']))
        $null = Add-Check $list 'the pending sidecar is clear: the transaction resolved' `
            (Test-SimBlank -Value $after['pending_auto_nonce']) `
            (Format-SimValue $after['pending_auto_nonce'])
        $null = Add-Check $list 'the attempt row records the attempted AUTO nonce' `
            (Test-SimExactDouble -Actual $after['shared']['last_attempt_auto_nonce'] `
                -Expected $counterBefore) `
            (Format-SimValue $after['shared']['last_attempt_auto_nonce'])
        $null = Add-Check $list 'the attempt row records an effective seed' `
            (($after['shared']['last_attempt_effective_seed']).GetType().FullName -ceq
             'System.Double') `
            (Format-SimValue $after['shared']['last_attempt_effective_seed'])
        $null = Add-Check $list 'the attempt result is REFUSED, not the indeterminate token' `
            (Test-SimExactText -Actual $after['shared']['last_attempt_result'] -Expected 'REFUSED') `
            (Format-SimValue $after['shared']['last_attempt_result'])
        $null = Add-Check $list 'the active bank did not move' `
            (Test-SimSameValue -A $before['shared']['active_bank'] -B $after['shared']['active_bank']) `
            (Format-SimValue $after['shared']['active_bank'])

        $null = $axisObservations.Add([pscustomobject]@{
            Scenario      = 'P6-FP1'
            What          = 'failure after a persisted AUTO advance'
            AttemptResult = $after['shared']['last_attempt_result']
            CounterMoved  = (Test-SimExactDouble -Actual $after['shared']['next_auto_nonce'] `
                -Expected ($counterBefore + 1))
            AttemptNonce  = $after['shared']['last_attempt_auto_nonce']
        })

        Add-Phase6Result 'P6-FP1' `
            'A failure after a persisted AUTO advance records REFUSED with the advance intact' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            ((Format-Checklist $list) + "`r`n" +
             (Format-Phase6State -State $before -Label 'before') + "`r`n" +
             '    action: PCCM_RunSimulation, AUTO, failpoint ' +
             [string]$failpoints.AfterNoncePersisted + "`r`n" +
             '    announcement: ' + $announced + "`r`n" +
             (Format-Phase6State -State $after -Label 'after'))
    } catch {
        Add-Phase6Result 'P6-FP1' `
            'A failure after a persisted AUTO advance records REFUSED with the advance intact' `
            'FAIL' (Format-Phase6Err $_)
    }

    # -------------------------------------------------------------------
    # P6-FP2. Phase6CandidateBank: written, not verified, not published.
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $null = Set-Phase6Fixture -PlanCaseId ([int]$goldenCase.plan_case_id) `
            -SuppliedSeed ([double]$GateBCases.supplied_seed) `
            -Iterations ([int]$GateBCases.iterations)
        $before = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
        $activeBefore = Get-Phase6ActiveBank -State $before
        $candidate = Get-Phase6CandidateTarget -Inspection $SimInspection -ActiveBank $activeBefore
        $announced = Invoke-Phase6Simulation -Excel $Excel `
            -FailAfterStage ([string]$failpoints.CandidateBank)
        $after = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection

        $null = Add-Check $list 'the endpoint announced a failure' `
            (Test-Phase6Announced -Result $announced -Kind 'FAIL') $announced
        $null = Add-Check $list 'the attempt result is FAILED' `
            (Test-SimExactText -Actual $after['shared']['last_attempt_result'] -Expected 'FAILED') `
            (Format-SimValue $after['shared']['last_attempt_result'])
        $null = Add-Check $list ('the selector did not move: bank ' + [char]39 + $activeBefore +
            [char]39 + ' is still active') `
            ((Get-Phase6ActiveBank -State $after) -ceq $activeBefore) `
            (Format-SimValue $after['shared']['active_bank'])
        $null = Add-Check $list 'the run id was not allocated' `
            (Test-SimSameValue -A $before['shared']['last_run_id'] -B $after['shared']['last_run_id']) `
            (Format-SimValue $after['shared']['last_run_id'])
        if (-not [string]::IsNullOrEmpty($activeBefore)) {
            Add-SimUnchangedChecks -List $list -Before $before[('bank_' + $activeBefore)] `
                -After $after[('bank_' + $activeBefore)] `
                -Label ('the prior publication in bank ' + $activeBefore + ' is authoritative and unchanged')
        }
        # THE CANDIDATE BANK MAY HOLD WRITTEN DATA, AND THAT IS NOT A DEFECT.
        # Production wrote it before the failpoint and never verified it, so it
        # has no semantic standing. Its contents are RECORDED as evidence and
        # deliberately not asserted either way.
        Add-Note ('P6-FP2: the unverified candidate bank ' + $candidate +
                  ' holds request fingerprint ' +
                  (Format-SimValue $after[('bank_' + $candidate)]['request_fingerprint']) +
                  ' and result digest ' +
                  (Format-SimValue $after[('bank_' + $candidate)]['result_digest']) +
                  '. It has no semantic standing: the selector never moved to it.')

        Add-Phase6Result 'P6-FP2' 'A candidate-bank failure publishes nothing' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            ((Format-Checklist $list) + "`r`n" +
             (Format-Phase6State -State $before -Label 'before') + "`r`n" +
             '    action: PCCM_RunSimulation, FIXED, failpoint ' +
             [string]$failpoints.CandidateBank + "`r`n" +
             '    announcement: ' + $announced + "`r`n" +
             (Format-Phase6State -State $after -Label 'after'))
    } catch {
        Add-Phase6Result 'P6-FP2' 'A candidate-bank failure publishes nothing' 'FAIL' `
            (Format-Phase6Err $_)
    }

    # -------------------------------------------------------------------
    # P6-FP3. Phase6FinalCommit: the prior shared block comes back.
    # -------------------------------------------------------------------
    # THE RESTORE HAS ALREADY HAPPENED BY THE TIME POWERSHELL SEES ANYTHING.
    # The failpoint fires after the D22:D30 assignment and before production's
    # own restore, both inside one Application.Run. What is observable here is
    # the state AFTER the call returned - which is exactly the claim that
    # matters: the prior block is back and the prior bank is still active.
    try {
        $list = New-Checklist
        $null = Set-Phase6Fixture -PlanCaseId ([int]$goldenCase.plan_case_id) `
            -SuppliedSeed ([double]$GateBCases.supplied_seed) `
            -Iterations ([int]$GateBCases.iterations)
        $before = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
        $activeBefore = Get-Phase6ActiveBank -State $before
        $candidate = Get-Phase6CandidateTarget -Inspection $SimInspection -ActiveBank $activeBefore
        $announced = Invoke-Phase6Simulation -Excel $Excel `
            -FailAfterStage ([string]$failpoints.FinalCommit)
        $after = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection

        $null = Add-Check $list 'the endpoint announced a failure' `
            (Test-Phase6Announced -Result $announced -Kind 'FAIL') $announced
        $null = Add-Check $list 'the attempt result is FAILED' `
            (Test-SimExactText -Actual $after['shared']['last_attempt_result'] -Expected 'FAILED') `
            (Format-SimValue $after['shared']['last_attempt_result'])
        $null = Add-Check $list ('the active bank after return is the prior bank ' +
            [char]39 + $activeBefore + [char]39) `
            ((Get-Phase6ActiveBank -State $after) -ceq $activeBefore) `
            (Format-SimValue $after['shared']['active_bank'])
        # THE PRIOR SHARED BLOCK IS BACK, field by field, except the two attempt
        # rows the failure path legitimately rewrites afterwards.
        foreach ($key in $before['shared'].Keys) {
            if (($key -eq 'last_attempt_result') -or ($key -eq 'last_attempt_detail')) { continue }
            $null = Add-Check $list ('the restored shared block preserves ' + $key) `
                (Test-SimSameValue -A $before['shared'][$key] -B $after['shared'][$key]) `
                ('was ' + (Format-SimValue $before['shared'][$key]) + ', now ' +
                 (Format-SimValue $after['shared'][$key]))
        }
        if (-not [string]::IsNullOrEmpty($activeBefore)) {
            Add-SimUnchangedChecks -List $list -Before $before[('bank_' + $activeBefore)] `
                -After $after[('bank_' + $activeBefore)] `
                -Label ('the prior publication in bank ' + $activeBefore + ' is unchanged')
        }
        Add-Note ('P6-FP3: the candidate bank ' + $candidate + ' remains unpublished; ' +
                  'the selector never named it.')

        Add-Phase6Result 'P6-FP3' 'A final-commit failure restores the prior shared block' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            ((Format-Checklist $list) + "`r`n" +
             (Format-Phase6State -State $before -Label 'before') + "`r`n" +
             '    action: PCCM_RunSimulation, FIXED, failpoint ' +
             [string]$failpoints.FinalCommit + "`r`n" +
             '    announcement: ' + $announced + "`r`n" +
             (Format-Phase6State -State $after -Label 'after'))
    } catch {
        Add-Phase6Result 'P6-FP3' 'A final-commit failure restores the prior shared block' 'FAIL' `
            (Format-Phase6Err $_)
    }

    # -------------------------------------------------------------------
    # P6-REC1 .. P6-REC5. The durable F21 recovery protocol.
    # -------------------------------------------------------------------
    # THESE FIVE STATES ARE NOT ALL THE SAME KIND OF THING, and calling them all
    # "what a crashed run leaves behind" would be wrong:
    #
    #   REC1, REC2  CRASH-EQUIVALENT write-ahead states. A marker with the
    #               counter either advanced past it or still at it is exactly
    #               what an interrupted transaction leaves, and production
    #               resolves both and clears the marker.
    #   REC3        an INCONSISTENT durable-recovery state: a marker the counter
    #               agrees with in neither direction. Production must retain it
    #               and stay blocked.
    #   REC4        a CORRUPTED / unreadable pending marker.
    #   REC5        a CORRUPTED / unreadable counter.
    #
    # All five are honestly constructible by writing cells on the disposable
    # copy. None of them fakes a COM failure.
    function Invoke-Phase6RecoveryScenario {
        param([string]$Id, [string]$Name, [string]$Address, $FixtureValue,
              [scriptblock]$Assert, [switch]$SecondAttempt)

        # FAIL-CLOSED ON A CONTAMINATED WORKBOOK. If an earlier fixture could not
        # be restored, this scenario would be describing a state the harness put
        # there, so it is not run and says so.
        if (-not (Test-Phase6FixtureIntegrity)) {
            Add-Phase6Result $Id $Name 'FAIL' `
                ('not attempted: the harness could not restore an earlier fixture, so ' +
                 'this workbook is no longer trustworthy for behavioural evidence. ' +
                 (Get-Phase6ContaminationReason))
            return
        }

        $list = New-Checklist
        $fixture = $null
        $evidence = ''
        $scenarioFailure = ''
        try {
            $null = Set-Phase6Fixture -PlanCaseId ([int]$goldenCase.plan_case_id) `
                -SuppliedSeed $null -Iterations ([int]$GateBCases.iterations)

            $fixture = New-Phase6CellFixture -Workbook $Workbook -Inspection $SimInspection `
                -Address $Address
            # THE FLAG IS SET BY Set-Phase6CellFixture ITSELF, before the
            # verification, so a write that took but failed its read-back is
            # still unwound.
            $fixtureOk = Set-Phase6CellFixture -Workbook $Workbook -Inspection $SimInspection `
                -Fixture $fixture -Value $FixtureValue -List $list -Label $Id

            # THE ESTABLISHMENT VERDICT IS LOAD-BEARING. A COM write can return
            # normally and still leave a cell holding something other than what
            # was asked for - a coerced type, a rejected value, a protected
            # sheet. Recording the failed check and then running the endpoint
            # anyway would produce behavioural observations against a fixture
            # the harness had ITSELF just proved was not established, which is
            # exactly the sequence this scenario exists to respect:
            #
            #   write -> VERIFY -> only then invoke production
            #
            # The throw goes to the catch, the finally still runs, and the
            # fixture is still marked written, so restoration is still attempted
            # and still verified.
            if (-not $fixtureOk) {
                throw ('the direct machine-state fixture at ' + $Address +
                       ' could not be established exactly; production was not invoked')
            }

            $before = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
            $announced = Invoke-Phase6Simulation -Excel $Excel
            $after = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
            $secondAnnounced = ''
            $afterSecond = $null
            if ($SecondAttempt) {
                # THE LOCK OUTLIVES THE ATTEMPT ROW. A second AUTO run must stay
                # blocked, which is the whole point of a DURABLE marker.
                $secondAnnounced = Invoke-Phase6Simulation -Excel $Excel
                $afterSecond = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
            }

            # POST EVIDENCE IS CAPTURED BEFORE CLEANUP, ALWAYS.
            $evidence = (Format-Phase6State -State $before -Label 'before') + "`r`n" +
                        '    fixture: ' + $Address + ' = ' + (Format-SimValue $FixtureValue) + "`r`n" +
                        '    announcement: ' + $announced + "`r`n" +
                        (Format-Phase6State -State $after -Label 'after')
            if ($SecondAttempt) {
                $evidence += ("`r`n" + '    second announcement: ' + $secondAnnounced + "`r`n" +
                              (Format-Phase6State -State $afterSecond -Label 'after second attempt'))
            }

            & $Assert $list $before $after $announced $afterSecond $secondAnnounced
        } catch {
            # THE ORIGINAL FAILURE IS PRESERVED. Cleanup runs next and must not
            # replace this with its own story.
            $scenarioFailure = Format-Phase6Err $_
        } finally {
            # RESTORATION IS REACHED ON EVERY PATH. An exception between the
            # fixture write and the restore is exactly the case that would
            # otherwise leave a modified F21 or counter in the workbook and let
            # every later scenario run against it.
            if ($null -ne $fixture) {
                $null = Complete-Phase6Fixture -Workbook $Workbook -Inspection $SimInspection `
                    -Fixture $fixture -List $list -Label $Id
            }
        }

        if (-not [string]::IsNullOrEmpty($scenarioFailure)) {
            $null = Add-Check $list ($Id + ': the scenario ran to completion') $false `
                $scenarioFailure
        }
        Add-Phase6Result $Id $Name `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            ((Format-Checklist $list) + $(if ($evidence) { "`r`n" + $evidence } else { '' }))
    }

    $pendingCell = Get-SimPendingCell -Inspection $SimInspection
    $counterCell = Get-SimAddressFor -Inspection $SimInspection -FieldKey 'next_auto_nonce'
    $runIdCell = Get-SimAddressFor -Inspection $SimInspection -FieldKey 'last_run_id'
    $counterNow = [double](Get-SimRawCell -Workbook $Workbook -Inspection $SimInspection `
        -Address $counterCell)

    # REC1: marker m with the counter already at m+1. ONE CELL WRITTEN - the
    # counter is left exactly where production put it, so the fixture creates
    # the state by choosing m rather than by corrupting a second cell.
    if ($counterNow -ge ([double]$bounds.nonce_first_valid_allocation + 1)) {
        Invoke-Phase6RecoveryScenario -Id 'P6-REC1' `
            -Name 'A crash-equivalent marker whose advance DID persist' `
            -Address $pendingCell -FixtureValue ($counterNow - 1) -Assert {
                param($list, $before, $after, $announced, $afterSecond, $secondAnnounced)
                $null = Add-Check $list 'the run succeeded after reconciling the marker' `
                    (Test-Phase6Announced -Result $announced -Kind 'OK') $announced
                $null = Add-Check $list 'the resolved marker was cleared' `
                    (Test-SimBlank -Value $after['pending_auto_nonce']) `
                    (Format-SimValue $after['pending_auto_nonce'])
                $null = Add-Check $list 'the pending nonce was NOT reissued' `
                    (-not (Test-SimSameValue -A $before['pending_auto_nonce'] `
                        -B $after['shared']['last_attempt_auto_nonce'])) `
                    ('pending was ' + (Format-SimValue $before['pending_auto_nonce']) +
                     ', attempted ' + (Format-SimValue $after['shared']['last_attempt_auto_nonce']))
            }
    } else {
        Add-Phase6Result 'P6-REC1' 'A crash-equivalent marker whose advance DID persist' 'FAIL' `
            ('the AUTO counter is at ' + [string]$counterNow + ', so no marker m exists for ' +
             'which the counter already reads m+1 without corrupting a second cell. The ' +
             'earlier AUTO scenarios should have advanced it; that they did not is the defect.')
    }

    # REC2: marker m with the counter still at m. ONE CELL WRITTEN.
    Invoke-Phase6RecoveryScenario -Id 'P6-REC2' `
        -Name 'A crash-equivalent marker whose advance never persisted' `
        -Address $pendingCell -FixtureValue ([double](Get-SimRawCell -Workbook $Workbook `
            -Inspection $SimInspection -Address $counterCell)) -Assert {
            param($list, $before, $after, $announced, $afterSecond, $secondAnnounced)
            $null = Add-Check $list 'the run succeeded after reconciling the marker' `
                (Test-Phase6Announced -Result $announced -Kind 'OK') $announced
            $null = Add-Check $list 'the resolved marker was cleared' `
                (Test-SimBlank -Value $after['pending_auto_nonce']) `
                (Format-SimValue $after['pending_auto_nonce'])
            $null = Add-Check $list 'the never-consumed nonce was legitimately taken by this run' `
                (Test-SimSameValue -A $before['pending_auto_nonce'] `
                    -B $after['shared']['last_attempt_auto_nonce']) `
                ('pending was ' + (Format-SimValue $before['pending_auto_nonce']) +
                 ', attempted ' + (Format-SimValue $after['shared']['last_attempt_auto_nonce']))
        }

    # REC3: an inconsistent state - the counter is neither m nor m+1. ONE CELL
    # WRITTEN: m is chosen so the EXISTING counter already disagrees with it.
    Invoke-Phase6RecoveryScenario -Id 'P6-REC3' `
        -Name 'An inconsistent durable-recovery state blocks every AUTO run' `
        -Address $pendingCell -FixtureValue ([double](Get-SimRawCell -Workbook $Workbook `
            -Inspection $SimInspection -Address $counterCell) + 5) -SecondAttempt -Assert {
            param($list, $before, $after, $announced, $afterSecond, $secondAnnounced)
            $null = Add-Check $list 'the run was refused' `
                (Test-Phase6Announced -Result $announced -Kind 'FAIL') $announced
            $null = Add-Check $list 'the attempt result is REFUSED' `
                (Test-SimExactText -Actual $after['shared']['last_attempt_result'] -Expected 'REFUSED') `
                (Format-SimValue $after['shared']['last_attempt_result'])
            $null = Add-Check $list 'the marker was RETAINED, not normalised away' `
                (Test-SimSameValue -A $before['pending_auto_nonce'] -B $after['pending_auto_nonce']) `
                ('was ' + (Format-SimValue $before['pending_auto_nonce']) + ', now ' +
                 (Format-SimValue $after['pending_auto_nonce']))
            $null = Add-Check $list 'no allocation was made: the counter did not move' `
                (Test-SimSameValue -A $before['shared']['next_auto_nonce'] `
                    -B $after['shared']['next_auto_nonce']) `
                (Format-SimValue $after['shared']['next_auto_nonce'])
            $null = Add-Check $list 'a SECOND AUTO run is still blocked: the lock is durable' `
                (Test-Phase6Announced -Result $secondAnnounced -Kind 'FAIL') $secondAnnounced
            $null = Add-Check $list 'and the marker survived that attempt too' `
                (Test-SimSameValue -A $before['pending_auto_nonce'] `
                    -B $afterSecond['pending_auto_nonce']) `
                (Format-SimValue $afterSecond['pending_auto_nonce'])
            $null = $axisObservations.Add([pscustomobject]@{
                Scenario      = 'P6-REC3'
                What          = 'refusal while reconciling a PRIOR marker'
                AttemptResult = $after['shared']['last_attempt_result']
                CounterMoved  = (-not (Test-SimSameValue `
                    -A $before['shared']['next_auto_nonce'] `
                    -B $after['shared']['next_auto_nonce']))
                AttemptNonce  = $after['shared']['last_attempt_auto_nonce']
            })
        }

    # REC4: a corrupted, unreadable pending marker.
    Invoke-Phase6RecoveryScenario -Id 'P6-REC4' `
        -Name 'A corrupted pending marker refuses and is retained' `
        -Address $pendingCell -FixtureValue 'not-a-nonce' -Assert {
            param($list, $before, $after, $announced, $afterSecond, $secondAnnounced)
            $null = Add-Check $list 'the run was refused' `
                (Test-Phase6Announced -Result $announced -Kind 'FAIL') $announced
            $null = Add-Check $list 'the attempt result is REFUSED' `
                (Test-SimExactText -Actual $after['shared']['last_attempt_result'] -Expected 'REFUSED') `
                (Format-SimValue $after['shared']['last_attempt_result'])
            $null = Add-Check $list 'the corrupted marker was retained, not overwritten' `
                (Test-SimSameValue -A $before['pending_auto_nonce'] -B $after['pending_auto_nonce']) `
                (Format-SimValue $after['pending_auto_nonce'])
            $null = Add-Check $list 'the counter did not move' `
                (Test-SimSameValue -A $before['shared']['next_auto_nonce'] `
                    -B $after['shared']['next_auto_nonce']) `
                (Format-SimValue $after['shared']['next_auto_nonce'])
        }

    # REC5: a corrupted, unreadable counter. Honest: ReadShared refuses a value
    # that is not a whole number in range, which is a state a corrupted workbook
    # really can be in - not a simulated COM raise.
    Invoke-Phase6RecoveryScenario -Id 'P6-REC5' `
        -Name 'A corrupted AUTO nonce counter refuses before any allocation' `
        -Address $counterCell -FixtureValue 'not-a-counter' -Assert {
            param($list, $before, $after, $announced, $afterSecond, $secondAnnounced)
            $null = Add-Check $list 'the run was refused' `
                (Test-Phase6Announced -Result $announced -Kind 'FAIL') $announced
            $null = Add-Check $list 'the attempt result is REFUSED' `
                (Test-SimExactText -Actual $after['shared']['last_attempt_result'] -Expected 'REFUSED') `
                (Format-SimValue $after['shared']['last_attempt_result'])
            $null = Add-Check $list 'the corrupted counter was not repaired by production' `
                (Test-SimSameValue -A $before['shared']['next_auto_nonce'] `
                    -B $after['shared']['next_auto_nonce']) `
                (Format-SimValue $after['shared']['next_auto_nonce'])
            $null = Add-Check $list 'no nonce identity was selected for this attempt' `
                (Test-SimBlank -Value $after['shared']['last_attempt_auto_nonce']) `
                (Format-SimValue $after['shared']['last_attempt_auto_nonce'])
            $null = Add-Check $list 'the sidecar was not written' `
                (Test-SimBlank -Value $after['pending_auto_nonce']) `
                (Format-SimValue $after['pending_auto_nonce'])
            $null = $axisObservations.Add([pscustomobject]@{
                Scenario      = 'P6-REC5'
                What          = 'refusal on an unreadable counter, before any allocation'
                AttemptResult = $after['shared']['last_attempt_result']
                CounterMoved  = (-not (Test-SimSameValue `
                    -A $before['shared']['next_auto_nonce'] `
                    -B $after['shared']['next_auto_nonce']))
                AttemptNonce  = $after['shared']['last_attempt_auto_nonce']
            })
        }

    # -------------------------------------------------------------------
    # P6-RIDMAX. Run-ID exhaustion refuses before allocation.
    # -------------------------------------------------------------------
    # QUALITATIVELY DIFFERENT FROM THE NON-INDUCIBLE COM-FAILURE CASES. This one
    # is honestly constructible, cheap, deterministic, pre-allocation and
    # non-computational: set the persisted Last Run ID to the accepted maximum
    # and ask for a run.
    #
    # WRITING A LARGE NUMBER IS NOT THE PROOF. The proof is the production
    # endpoint's REFUSAL and the untouched allocation and publication state that
    # comes with it.
    #
    # It writes a machine cell, so it carries the same fail-closed contract as
    # the recovery scenarios: restoration on every path, and contamination
    # latched if the restoration cannot be verified.
    if (-not (Test-Phase6FixtureIntegrity)) {
        Add-Phase6Result 'P6-RIDMAX' 'Run-ID exhaustion refuses before allocation' 'FAIL' `
            ('not attempted: the harness could not restore an earlier fixture, so this ' +
             'workbook is no longer trustworthy for behavioural evidence. ' +
             (Get-Phase6ContaminationReason))
    } else {
        $list = New-Checklist
        $fixture = $null
        $evidence = ''
        $scenarioFailure = ''
        try {
            $null = Set-Phase6Fixture -PlanCaseId ([int]$goldenCase.plan_case_id) `
                -SuppliedSeed $null -Iterations ([int]$GateBCases.iterations)

            $fixture = New-Phase6CellFixture -Workbook $Workbook -Inspection $SimInspection `
                -Address $runIdCell
            $null = Add-Check $list `
                'the original Last Run ID was captured before anything was written' `
                ($null -ne $fixture) (Format-SimValue $fixture.Original)
            $fixtureOk = Set-Phase6CellFixture -Workbook $Workbook -Inspection $SimInspection `
                -Fixture $fixture -Value ([double]$bounds.run_id_maximum) -List $list `
                -Label 'P6-RIDMAX'
            # LOAD-BEARING, exactly as in the recovery scenarios: a fixture that
            # did not take is not a fixture, and production is not invoked
            # against one.
            if (-not $fixtureOk) {
                throw ('the direct machine-state fixture at ' + $runIdCell +
                       ' could not be established exactly; production was not invoked')
            }

            $before = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
            $announced = Invoke-Phase6Simulation -Excel $Excel
            $after = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection

            # POST EVIDENCE IS CAPTURED BEFORE CLEANUP, ALWAYS.
            $evidence = (Format-Phase6State -State $before -Label 'before') + "`r`n" +
                        '    fixture: ' + $runIdCell + ' = ' + [string]$bounds.run_id_maximum + "`r`n" +
                        '    announcement: ' + $announced + "`r`n" +
                        (Format-Phase6State -State $after -Label 'after')

            $null = Add-Check $list 'the endpoint refused' `
                (Test-Phase6Announced -Result $announced -Kind 'FAIL') $announced
            $null = Add-Check $list 'the attempt result is REFUSED' `
                (Test-SimExactText -Actual $after['shared']['last_attempt_result'] -Expected 'REFUSED') `
                (Format-SimValue $after['shared']['last_attempt_result'])
            $null = Add-Check $list `
                'the refusal happened BEFORE AUTO allocation: the counter did not move' `
                (Test-SimSameValue -A $before['shared']['next_auto_nonce'] `
                    -B $after['shared']['next_auto_nonce']) `
                ('was ' + (Format-SimValue $before['shared']['next_auto_nonce']) + ', now ' +
                 (Format-SimValue $after['shared']['next_auto_nonce']))
            $null = Add-Check $list 'no AUTO nonce identity was selected' `
                (Test-SimBlank -Value $after['shared']['last_attempt_auto_nonce']) `
                (Format-SimValue $after['shared']['last_attempt_auto_nonce'])
            $null = Add-Check $list 'the sidecar was not written' `
                (Test-SimSameValue -A $before['pending_auto_nonce'] -B $after['pending_auto_nonce']) `
                (Format-SimValue $after['pending_auto_nonce'])
            $null = Add-Check $list 'the active bank did not move' `
                (Test-SimSameValue -A $before['shared']['active_bank'] -B $after['shared']['active_bank']) `
                (Format-SimValue $after['shared']['active_bank'])
            foreach ($bank in $banks) {
                Add-SimUnchangedChecks -List $list -Before $before[('bank_' + $bank)] `
                    -After $after[('bank_' + $bank)] `
                    -Label ('run-ID exhaustion left bank ' + $bank)
            }
            $null = Add-Check $list 'the run id was not incremented past its maximum' `
                (Test-SimExactDouble -Actual $after['shared']['last_run_id'] `
                    -Expected ([double]$bounds.run_id_maximum)) `
                (Format-SimValue $after['shared']['last_run_id'])

            $null = $axisObservations.Add([pscustomobject]@{
                Scenario      = 'P6-RIDMAX'
                What          = 'run-ID exhaustion, refused before allocation'
                AttemptResult = $after['shared']['last_attempt_result']
                CounterMoved  = (-not (Test-SimSameValue -A $before['shared']['next_auto_nonce'] `
                    -B $after['shared']['next_auto_nonce']))
                AttemptNonce  = $after['shared']['last_attempt_auto_nonce']
            })
        } catch {
            $scenarioFailure = Format-Phase6Err $_
        } finally {
            if ($null -ne $fixture) {
                $null = Complete-Phase6Fixture -Workbook $Workbook -Inspection $SimInspection `
                    -Fixture $fixture -List $list -Label 'P6-RIDMAX'
            }
        }

        if (-not [string]::IsNullOrEmpty($scenarioFailure)) {
            $null = Add-Check $list 'P6-RIDMAX: the scenario ran to completion' $false `
                $scenarioFailure
        }
        Add-Phase6Result 'P6-RIDMAX' 'Run-ID exhaustion refuses before allocation' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            ((Format-Checklist $list) + $(if ($evidence) { "`r`n" + $evidence } else { '' }))
    }

    # -------------------------------------------------------------------
    # P6-AXIS. One result string does not encode the allocation classification.
    # -------------------------------------------------------------------
    # THE CONCLUSION, IN PERSISTED TERMS ONLY. Several unsuccessful attempts
    # above recorded the SAME Last Attempt Result - REFUSED - while the
    # persisted AUTO axis around them differed:
    #
    #   P6-FP1     REFUSED, and the counter DID advance m -> m+1
    #   P6-RF1     REFUSED, and the counter did not move
    #   P6-REC3    REFUSED, reconciling a PRIOR marker, no transition of its own
    #   P6-REC5    REFUSED, and no identity was even selected
    #
    # That is direct runtime evidence that Last Attempt Result alone does not
    # encode the physical allocation classification, which is exactly what the
    # Step-12 two-axis correction asserted.
    #
    # WHAT IS NOT CLAIMED HERE. `NonceConsumed` is a Private Boolean inside
    # modSimReport. PowerShell cannot observe it, so its projection from the
    # allocation state remains SOURCE evidence; this scenario proves the
    # persisted consequence, not the field.
    try {
        $list = New-Checklist
        $observations = @($axisObservations)
        $null = Add-Check $list 'at least three unsuccessful attempts were observed' `
            ($observations.Count -ge 3) ('observations: ' + $observations.Count)
        $refused = @($observations | Where-Object {
            (Test-SimExactText -Actual $_.AttemptResult -Expected 'REFUSED') })
        $null = Add-Check $list 'several of them recorded the SAME attempt result' `
            ($refused.Count -ge 2) ('REFUSED observations: ' + $refused.Count)
        $advanced = @($refused | Where-Object { $_.CounterMoved })
        $unmoved = @($refused | Where-Object { -not $_.CounterMoved })
        $null = Add-Check $list 'at least one REFUSED attempt DID advance the AUTO counter' `
            ($advanced.Count -ge 1) `
            ('advanced: ' + (@($advanced | ForEach-Object { $_.Scenario }) -join ', '))
        $null = Add-Check $list 'at least one REFUSED attempt did NOT advance it' `
            ($unmoved.Count -ge 1) `
            ('did not advance: ' + (@($unmoved | ForEach-Object { $_.Scenario }) -join ', '))
        $null = Add-Check $list `
            'therefore Last Attempt Result alone does not encode the allocation classification' `
            (($advanced.Count -ge 1) -and ($unmoved.Count -ge 1))
        $lines = @()
        foreach ($item in $observations) {
            $lines += ('    ' + $item.Scenario + '  ' + (Format-SimValue $item.AttemptResult) +
                       '  counter advanced: ' + [string]$item.CounterMoved +
                       '  attempted nonce: ' + (Format-SimValue $item.AttemptNonce) +
                       '  (' + $item.What + ')')
        }
        Add-Phase6Result 'P6-AXIS' `
            'The attempt axis and the allocation axis are separate, in persisted evidence' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            ((Format-Checklist $list) + "`r`n" + ($lines -join "`r`n") + "`r`n" +
             '    NOT CLAIMED: the private NonceConsumed projection.' + "`r`n" +
             '    PowerShell cannot observe it; that remains source evidence.')
    } catch {
        Add-Phase6Result 'P6-AXIS' `
            'The attempt axis and the allocation axis are separate, in persisted evidence' `
            'FAIL' (Format-Phase6Err $_)
    }

    # -------------------------------------------------------------------
    # P6-FIN. Completeness, through the GUARDED reporter.
    # -------------------------------------------------------------------
    # THE ORDER IS THE ROUND-4A CORRECTION, reused rather than reinvented.
    #
    #   every functional P6 scenario
    #     -> P6-FIN, through Add-Phase6Result, so it is itself a guarded result
    #        and a duplicate attempt at it is a recorded violation
    #     -> P6-LDG, emitted by the DRIVER through the unguarded Add-Result,
    #        last of all, so the ledger can see a duplicate P6-FIN and can never
    #        suppress the result that reports on the ledger.
    #
    # P6-LDG is deliberately NOT in the set P6-FIN requires: making the ledger's
    # verdict a precondition of the completeness verdict that precedes it is the
    # circular ordering this correction removes.
    try {
        $list = New-Checklist
        $recorded = @($Results | Where-Object { $_.Id -like 'P6-*' })
        foreach ($id in (Get-Phase6RequiredScenarioIds)) {
            $matching = @($recorded | Where-Object { $_.Id -eq $id })
            $null = Add-Check $list ('the required scenario ' + $id + ' has exactly one result') `
                ($matching.Count -eq 1) ('results for ' + $id + ': ' + $matching.Count)
            $null = Add-Check $list ('the required scenario ' + $id + ' passed') `
                (($matching.Count -eq 1) -and ($matching[0].Status -eq 'PASS')) `
                $(if ($matching.Count -eq 1) { 'status ' + $matching[0].Status } else { 'no result' })
        }
        $skipped = @($recorded | Where-Object { $_.Status -eq 'SKIP' })
        $null = Add-Check $list 'no Phase-6 scenario was skipped' ($skipped.Count -eq 0) `
            ((@($skipped | ForEach-Object { $_.Id })) -join ', ')
        # AND THE WORKBOOK WAS STILL TRUSTWORTHY AT THE END. A run that
        # contaminated its own fixture state does not finish green even if every
        # scenario before the contamination passed.
        $null = Add-Check $list 'the harness restored every fixture it wrote' `
            (Test-Phase6FixtureIntegrity) (Get-Phase6ContaminationReason)
        Add-Phase6Result 'P6-FIN' 'Phase-6 completeness' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Phase6Result 'P6-FIN' 'Phase-6 completeness' 'FAIL' (Format-Phase6Err $_)
    }
}
