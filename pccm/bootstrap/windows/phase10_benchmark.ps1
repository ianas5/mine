<#
.SYNOPSIS
    PCCM Phase 10 Step 4 - the delivery PERFORMANCE BENCHMARK runner.

.DESCRIPTION
    THIS RECORDS. IT DOES NOT JUDGE.

    It measures how long the four real user commands take, on the real
    production .xlsm, on the real target machine, at three declared model sizes
    and the declared iteration counts. It is the first delivery-oriented
    baseline, so it carries no pass mark: a slow number here is a FACT about
    this machine, not a failure, and the ratios by which a LATER run is judged
    are recorded in the plan rather than applied to this one.

    IT DECLARES NOTHING OF ITS OWN
    ------------------------------
    The matrix - sizes, driver split, project years, iteration counts,
    forbidden combinations, how many cold and warm runs, which statistic
    compares them, and what the environment record must contain - is READ from
    `build/phase10_benchmark_plan.json`, which the Stage-A build emits from
    `builder/pccm_builder/benchmark.py`. This file executes that plan literally.
    If the plan and this runner ever disagreed, the plan would be right.

    WHAT IT IS NOT
    --------------
      * It is NOT an acceptance harness. It records no Gate-B result, no
        Phase-7/8/9 scenario outcome, and its output is never a pass or a fail.
      * It is NOT a surrogate. Nothing here re-implements calculation,
        sampling, ranking or the annual replay. It calls the production
        endpoints by name and times them.
      * It adds NO timing code to production VBA. Every measurement is taken
        from PowerShell, around one synchronous call.
      * It writes NOTHING into the workbook that anything reads back. No result
        is stored in a cell; the workbook is never saved.

    COLD AND WARM
    -------------
    Defined in the plan and obeyed here:

      COLD  the FIRST execution of that operation in this Excel session against
            this scenario and iteration count. Reported separately, never
            averaged into anything.
      WARM  each of the three executions immediately after it, with no other
            operation between them and no reopen.

    Excel startup, workbook open, the Stage-B bootstrap and the scenario
    fixture are SETUP. They are timed separately, reported separately, and are
    never inside an operation's elapsed time.

    A FAST FAILURE IS NOT A FAST OPERATION
    --------------------------------------
    Every timed execution carries its own evidence: the automation result, the
    resulting state, and for a stochastic operation the iteration count the
    workbook actually published. A sample that fails a gate is marked INVALID,
    excluded from the median, and reported with its refusal text.

.PARAMETER Scenario
    PERF-SMALL, PERF-MEDIUM or PERF-LARGE. Exactly one per invocation - these
    runs are long, and one command at a time is the point.

.PARAMETER Iterations
    Optionally narrow the iteration counts to a subset of the scenario's
    declared list. A count the plan does not declare for this scenario is
    REFUSED, which is what makes the Large x 100,000 cap unreachable from the
    command line.

.PARAMETER Operations
    Optionally narrow to a subset of the declared operation keys.

.PARAMETER BuildDir
    The Stage-A build directory. Defaults to <repo>/pccm/build.

.PARAMETER WorkDir
    Where the disposable copy of the build is made and the .xlsm is opened
    from. Defaults to the system temp directory. Set it deliberately to measure
    from a particular kind of location - the report records which kind it was.

.PARAMETER OutDir
    Where the JSON and Markdown artifacts are written. Defaults to the work
    directory.

.NOTES
    SAFETY. No security setting is altered, no registry key is touched, no
    Trusted Location is added, and no Excel process this script did not create
    is ever terminated. Shutdown is the accepted `com_lifecycle.ps1` path.
    The workbook is never saved.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('PERF-SMALL', 'PERF-MEDIUM', 'PERF-LARGE')]
    [string]$Scenario,
    [int[]]$Iterations,
    [string[]]$Operations,
    [string]$BuildDir,
    [string]$WorkDir,
    [string]$OutDir,
    [ValidateSet('Endpoints', 'Bulk')]
    [string]$FixtureMode = 'Endpoints',
    [switch]$KeepArtifacts
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# THE ACCEPTED FILES ARE REUSED, NOT REIMPLEMENTED. All three are
# definition-only at top level: dot-sourcing them defines functions and runs no
# scenario. Nothing here calls Invoke-Phase5GateBScenarios or
# Invoke-Phase6GateBScenarios, so no Gate-B result is produced or implied.
. (Join-Path $scriptDir 'com_lifecycle.ps1')
. (Join-Path $scriptDir 'phase5_gate_b_scenarios.ps1')
. (Join-Path $scriptDir 'phase6_gate_b_scenarios.ps1')

$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $scriptDir))
$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
if ([string]::IsNullOrWhiteSpace($BuildDir)) { $BuildDir = Join-Path $pccmRoot 'build' }

# The schema this runner understands. A plan from a different schema is refused
# rather than half-executed.
$script:BenchmarkSupportedSchema = 1

# ===========================================================================
# THE COM PRIMITIVES
# ===========================================================================
# COPIED, NOT REINVENTED, AND NOT EXTRACTED. `phase5_gate_b_scenarios.ps1` USES
# these helpers and does not DEFINE them - every standalone runner in this tree
# carries its own copy, because the alternative is editing the accepted Gate-B
# harness to dot-source a new shared file, and changing the bytes of an accepted
# harness for the sake of a measurement is not a trade this task may make.
#
# The block below is `bootstrap/windows/phase7_timing_scenarios.ps1`'s, verbatim.
# A control asserts it stayed verbatim, so a fix to one is a fix to both.

function Write-RowObject {
    param([object[]]$Row)
    Write-Output -NoEnumerate $Row
}

function Get-NamedValue {
    param($Workbook, [string]$DefinedName)
    $names = $null; $nm = $null; $rng = $null
    try {
        $names = $Workbook.Names
        $nm = $names.Item($DefinedName)
        $rng = $nm.RefersToRange
        $v = $rng.Value2
        if ($null -eq $v) { return '' }
        return [string]$v
    } finally {
        if ($null -ne $rng)   { Release-Transient $rng   'Range(name)'; $rng   = $null }
        if ($null -ne $nm)    { Release-Transient $nm    'Name';        $nm    = $null }
        if ($null -ne $names) { Release-Transient $names 'Names';       $names = $null }
    }
}

function Set-NamedValue {
    param($Workbook, [string]$DefinedName, $Value)
    $names = $null; $nm = $null; $rng = $null
    try {
        $names = $Workbook.Names
        $nm = $names.Item($DefinedName)
        $rng = $nm.RefersToRange
        if ($null -eq $Value) { $null = $rng.ClearContents() } else { $rng.Value2 = [double]$Value }
    } finally {
        if ($null -ne $rng)   { Release-Transient $rng   'Range(name)'; $rng   = $null }
        if ($null -ne $nm)    { Release-Transient $nm    'Name';        $nm    = $null }
        if ($null -ne $names) { Release-Transient $names 'Names';       $names = $null }
    }
}

function Get-TableColumnNames {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $cols = $null
    $out = @()
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $cols = $lo.ListColumns
        $colCount = [int]$cols.Count
        for ($i = 1; $i -le $colCount; $i++) {
            $c = $null
            try { $c = $cols.Item($i); $out += [string]$c.Name }
            finally { if ($null -ne $c) { Release-Transient $c 'ListColumn'; $c = $null } }
        }
    } finally {
        if ($null -ne $cols)            { Release-Transient $cols            'ListColumns'; $cols            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
    return $out
}

function Get-TableBody {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $body = $null
    $rowsObj = $null; $colsObj = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $body = $lo.DataBodyRange
        # An empty body is a valid outcome: emit NOTHING. The caller's @(...) turns
        # zero pipeline objects into an empty collection, which is exactly right.
        if ($null -eq $body) { return }

        # Row and column counts are read through named, released objects rather than
        # through $body.Rows.Count, which would mint an unowned Range on every
        # iteration of the loop.
        $rowsObj = $body.Rows
        $colsObj = $body.Columns
        $rowCount = [int]$rowsObj.Count
        $colCount = [int]$colsObj.Count
        Release-Transient $rowsObj 'Range(rows)'; $rowsObj = $null
        Release-Transient $colsObj 'Range(columns)'; $colsObj = $null

        for ($r = 1; $r -le $rowCount; $r++) {
            $line = @()
            for ($c = 1; $c -le $colCount; $c++) {
                $cell = $null
                try {
                    $cell = $body.Cells($r, $c)
                    $v = $cell.Value2
                    if ($null -eq $v) { $line += '' } else { $line += [string]$v }
                } finally {
                    if ($null -ne $cell) { Release-Transient $cell 'Range(cell)'; $cell = $null }
                }
            }
            Write-RowObject $line
        }
    } finally {
        if ($null -ne $rowsObj)         { Release-Transient $rowsObj         'Range(rows)';    $rowsObj         = $null }
        if ($null -ne $colsObj)         { Release-Transient $colsObj         'Range(columns)'; $colsObj         = $null }
        if ($null -ne $body)            { Release-Transient $body            'Range(body)';    $body            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';     $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects';    $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';      $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';     $localWorksheets = $null }
    }
    # No trailing return: every row has already been emitted, one object each.
}

function Set-TableCell {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$RowIndex, [int]$ColumnIndex, $Value)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $body = $null; $cell = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $body = $lo.DataBodyRange
        $cell = $body.Cells($RowIndex, $ColumnIndex)
        if ($null -eq $Value) {
            # A genuine blank, not zero. The two are different assumptions and the
            # harness has to be able to create each of them deliberately.
            $null = $cell.ClearContents()
        } elseif ($Value -is [string]) {
            $cell.Value2 = [string]$Value
        } else {
            $cell.Value2 = [double]$Value
        }
    } finally {
        if ($null -ne $cell)            { Release-Transient $cell            'Range(cell)'; $cell            = $null }
        if ($null -ne $body)            { Release-Transient $body            'Range(body)'; $body            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
}

function Get-TableRowCount {
    param($Workbook, [string]$SheetName, [string]$TableName)
    return @(Get-TableBody -Workbook $Workbook -SheetName $SheetName -TableName $TableName).Count
}

# ===========================================================================
# THE FIXTURE ON A PROTECTED WORKBOOK
# ===========================================================================
# WHAT CHANGED HERE, AND WHY IT IS HERE RATHER THAN IN THE ACCEPTED HARNESS.
#
# Benchmark Run 3 aborted on a COM `ListRow.Delete` against `tblFXRates` with
# "Table features aren't available because the sheet is protected." Protection
# is a Phase-10 addition - `build_stage_b.ps1` applies it and `Workbook_Open`
# re-applies it - and every harness in this tree was written against the
# workbook as it stood BEFORE that.
#
# The runtime protection reconciliation settled the capability split from
# Windows evidence rather than from assumption:
#
#   code writes a VALUE to a locked cell             PERMITTED (UserInterfaceOnly)
#   code performs a LISTOBJECT STRUCTURAL operation  REFUSED
#
# Production reaches the second capability through
# `modAppState.BeginStructuralOperation`, which opens `modProtection`'s
# structural window from inside VBA. A PowerShell COM caller never enters that
# window and is not given one: the protection boundary is closed, and a
# measurement is not a reason to reopen it.
#
# So the benchmark fixture resets WITHOUT structural mutation. It can, because
# the accepted tables are built carrying RESERVED BLANK BODY ROWS -
# `tblFXRates` is `data_rows: 12` with a single seeded row, and
# `tblInflationProfiles` is `data_rows: 10` with none - and because production
# reads those tables by CONTENT and never by shape. The proof is beside
# `Reset-Phase5FxTable` below.
#
# THE ACCEPTED HARNESS IS NOT EDITED. `phase5_gate_b_scenarios.ps1` stays
# byte-identical to what Gate B was accepted on. It USES these helpers and does
# not DEFINE them, and PowerShell resolves a function name at CALL time, so the
# definitions below - which follow the dot-source at the top of this file - are
# the ones its fixture tree reaches IN THIS PROCESS ONLY. No other runner is
# affected and no accepted evidence is disturbed.

# A BLANK ROW IS FOUND, NOT MADE. The accepted tables ship with reserved blank
# body rows and the fixture fills them from the top, so "make a blank row
# available" is a SEARCH over what is already there. The returned index is the
# first wholly blank body row, which is the row the fixture tree then writes.
#
# IT REFUSES RATHER THAN GROWING. A table with no blank row left genuinely needs
# a structural add; that add is genuinely refused under protection; and the
# caller is told exactly that - naming the table and its capacity - instead of
# meeting a bare 1004 with no context. Nothing here relaxes a requirement: a
# fixture that cannot be built is a fixture that does not get measured.
function Add-BlankTableRow {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $body = @(Get-TableBody -Workbook $Workbook -SheetName $SheetName -TableName $TableName)
    if ($body.Count -lt 1) {
        throw ('the benchmark fixture needs a blank row in ' + $TableName + ' on ' +
               $SheetName + ', and the table has no body row at all. Creating one is a ' +
               'ListObject structural operation, which worksheet protection refuses to a ' +
               'COM caller; only a production endpoint may perform one.')
    }
    for ($row = 1; $row -le $body.Count; $row++) {
        $blank = $true
        foreach ($value in @($body[$row - 1])) {
            if ([string]$value -ne '') { $blank = $false; break }
        }
        if ($blank) { return [int]$row }
    }
    throw ('the benchmark fixture needs a blank row in ' + $TableName + ' on ' + $SheetName +
           ', and all ' + [string]$body.Count + ' reserved body rows are populated. Growing ' +
           'the table is a ListObject structural operation, which worksheet protection ' +
           'refuses to a COM caller; only a production endpoint may perform one.')
}

# `Remove-TableRow` REFUSES. IT NO LONGER DELETES.
#
# Its `$victim.Delete()` is the exact line Benchmark Run 3 died on, and nothing
# in the benchmark's fixture path needs it: `Reset-Phase5FxTable` below resets by
# CONTENT. The two callers that remain in the dot-sourced Gate-B file are
# unreachable from here - `Clear-Phase5UserRows`, which has no caller anywhere in
# the tree, and the `fx_remove` arm of `Invoke-Phase5Mutation`, which belongs to
# the Gate-B scenarios this runner never executes.
#
# IT IS NOT DELETED FROM THE FILE, and the reason is a defect this project has
# already paid for. Phase-9 Windows run 1 died on
# `The term 'Write-RowObject' is not recognized` - a function in a DOT-SOURCED
# file calling a name the runner did not define. Removing this definition
# recreates that exact shape: `phase5_gate_b_scenarios.ps1` names
# `Remove-TableRow` in two places, and a name a reachable file can call must
# resolve. `tests/powershell_command_resolution_audit.ps1` reports it, which is
# how this was found rather than discovered on Windows.
#
# So the name resolves and the CAPABILITY is gone. A caller that appears in
# future is told what it tried to do and why it cannot, at the call site, which
# is a far better diagnosis than a 1004 from inside Excel - and there is no body
# here for a later edit to quietly fill back in.
function Remove-TableRow {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$RowIndex)
    throw ('the benchmark tried to delete row ' + [string]$RowIndex + ' of ' + $TableName +
           ' on ' + $SheetName + '. Deleting a ListRow is a ListObject structural ' +
           'operation, which worksheet protection refuses to a COM caller; only a ' +
           'production endpoint may perform one. The benchmark fixture resets by ' +
           'content instead - see Reset-Phase5FxTable below - so reaching this is a ' +
           'path that was never meant to run, not a protection problem to work around.')
}

# THE FX RESET, BY CONTENT.
#
# The accepted reset deleted every body row below the seed and then rewrote row
# 1 from the capture. Deleting was never what the fixture NEEDED - it was how
# the harness spelled "nothing below the seed survives" - and on the protected
# workbook it is refused. What the fixture requires afterwards is exactly two
# things:
#
#   row 1                 IS the captured locked seed, value for value and type
#                         for type
#   every row below it    carries no currency and no rate
#
# AND PRODUCTION CANNOT TELL A BLANKED ROW FROM AN ABSENT ONE. That is the whole
# claim, and it is proved in production's own source rather than assumed:
#
#   modCalcResolve.MatchingFxRows walks `1 To modWorkbook.BodyRowCount(table)`
#   and counts the rows whose Currency cell EQUALS the key it was handed.
#   modCalcResolve.RawCellText - the reader it counts through - exits False on
#   `IsEmpty`, so a blank row is never a candidate and can never be a match.
#   The row count is a loop bound and nothing else, and the row index is used
#   only to fetch the matched row's own rate. No production path reads
#   `tblFXRates`'s ListRows.Count, its DataBodyRange dimensions or a row
#   position AS A VALUE, and the reporting-currency invariant is "appears
#   exactly once", which rows carrying nothing cannot affect.
#
# It is also the shape production already runs against: Stage A builds
# `tblFXRates` with twelve body rows of which eleven are blank, and every
# accepted Phase-4 through Phase-9 run resolved FX over exactly that.
#
# DELETING WAS THE OPERATION THAT DID DAMAGE. Those eleven reserved rows carry
# the input contract's data validation - `lstCurrencies` on Currency, decimal
# greater than zero on the rate - so deleting them removed validated rows from a
# table the contract declares as `data_rows: 12`, and the accepted reset was
# quietly shrinking the delivered shape. Blanking leaves the validation where
# the contract put it.
#
# THE SEED RESTORATION IS UNCHANGED from the accepted reset: the same typed
# writer, the same strict comparator, the same refusal to decide what the value
# ought to be.
function Reset-Phase5FxTable {
    param($Workbook, $Inspection, $Seed)
    $fx = $Inspection.input_tables.fx_rates
    $seedRows = [int]$fx.locked_seed_rows
    # The row-1 arithmetic below is only correct for a single seed row.
    # `Save-Phase5LockedFxSeed` already refuses anything else; this refuses it
    # again here rather than letting a contract change drift silently past.
    if ($seedRows -ne 1) {
        throw ('the FX table declares ' + [string]$seedRows + ' locked seed rows; the ' +
               'benchmark fixture reset assumes exactly one')
    }
    $columns = @(Get-TableColumnNames -Workbook $Workbook -SheetName $fx.sheet `
        -TableName $fx.table_name)
    $rows = Get-TableRowCount -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name
    if ($rows -lt ($seedRows + 1)) {
        throw ('the FX table carries ' + [string]$rows + ' body row(s), so there is no row ' +
               'below the locked seed for the fixture to write a rate into. Adding one is a ' +
               'ListObject structural operation, which worksheet protection refuses to a ' +
               'COM caller; only a production endpoint may perform one.')
    }

    # Everything below the seed is BLANKED, whatever it is - the same reach as
    # the accepted delete loop, by a different mechanism. `Set-TableCell` with
    # $null is ClearContents, which makes a GENUINE blank: an empty string would
    # leave `IsEmpty` False and production would see a populated row.
    for ($row = $seedRows + 1; $row -le $rows; $row++) {
        for ($column = 1; $column -le $columns.Count; $column++) {
            Set-TableCell -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name `
                -RowIndex $row -ColumnIndex $column -Value $null
        }
    }

    # Row 1 is REWRITTEN from the capture. It is not trusted to still be the seed.
    #
    # THE CAPTURED VALUE IS WRITTEN BACK AS ITSELF. No [double], no [string], no
    # decision about what the value ought to be: Set-Phase5TypedCell assigns
    # Value2 directly, so a numeric seed stays numeric and a defective text seed
    # stays text and is exposed by the production calculation rather than being
    # quietly corrected here.
    Set-Phase5TypedCell -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name `
        -RowIndex 1 -ColumnIndex 1 -Value $Seed.Currency
    Set-Phase5TypedCell -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name `
        -RowIndex 1 -ColumnIndex 2 -Value $Seed.Rate

    # ONE TYPED READ-BACK PROVES BOTH HALVES. A restoration nobody checked is an
    # assumption, and a blanking nobody checked is a fixture that may be carrying
    # the previous scenario's rates into a measurement.
    $body = @(Get-Phase5TypedTableBody -Workbook $Workbook -SheetName $fx.sheet `
        -TableName $fx.table_name)
    if ((-not (Test-Phase5ExactValue -Actual $body[0][0] -Expected $Seed.Currency)) -or `
        (-not (Test-Phase5ExactValue -Actual $body[0][1] -Expected $Seed.Rate))) {
        throw ("the locked FX seed did not restore: row 1 is " +
               (Format-Phase5Typed $body[0][0]) + " / " + (Format-Phase5Typed $body[0][1]) +
               ", captured " + (Format-Phase5Typed $Seed.Currency) + " / " +
               (Format-Phase5Typed $Seed.Rate))
    }
    # STRICTLY $null, not "empty-looking". ClearContents leaves Value2 $null; a
    # cell holding the empty string is NOT a genuine blank and production would
    # read it as populated.
    for ($row = $seedRows + 1; $row -le $body.Count; $row++) {
        for ($column = 1; $column -le $columns.Count; $column++) {
            if ($null -ne $body[$row - 1][$column - 1]) {
                throw ("the FX reset did not clear row " + [string]$row + " column " +
                       [string]$column + ": it still holds " +
                       (Format-Phase5Typed $body[$row - 1][$column - 1]) +
                       ". Stale fixture data below the locked seed would be resolved by " +
                       "production as a real rate, so the measurement would not be of the " +
                       "model this run claims to build. ClearContents on a locked cell is " +
                       "the same capability as the value write UserInterfaceOnly permits, " +
                       "so a cell that refused to clear is new protection behaviour and " +
                       "must be reported, not retried.")
            }
        }
    }
}

# ===========================================================================
# THE BULK FIXTURE
# ===========================================================================
# WHY THIS EXISTS, MEASURED RATHER THAN ASSERTED. PERF-LARGE was operator-aborted
# after more than four hours in `BUILDING PERF-LARGE`, before the first timed
# operation. The cost is not the harness's COM chatter - it is that the fixture
# invoked `PCCM_AddCostLine` / `PCCM_AddRisk` three hundred times, and EACH of
# those is a full production structural operation:
#
#   modAppState.BeginStructuralOperation   unprotect 14 sheets, then a second
#                                          pass over all 14 to PROVE the release
#   modWorkbook.SnapshotTable x2           every cell of the register AND of the
#                                          profiling grid, for the rollback
#   modProfiling.SyncRows                  every existing weight into a
#                                          Dictionary, then the grid rewritten
#   modStructuralCheck.ValidateStructure   the register and the grid again
#   modAppState.FinishOperation            re-apply protection to 14 sheets and
#                                          verify, then .Calculate four sheets
#
# That is O(register rows x project years) PER ADD, so N adds is O(N^2 x years):
#
#   PERF-SMALL     20 adds      46,580 in-VBA cell visits,    80 recalcs
#   PERF-MEDIUM   100 adds     543,760 in-VBA cell visits,   400 recalcs
#   PERF-LARGE    300 adds   5,816,730 in-VBA cell visits, 1,200 recalcs
#
# ALL OF THAT IS PRODUCTION BEHAVING CORRECTLY. One user adding one cost line
# SHOULD snapshot for rollback, re-sync the grid, validate and re-protect. It is
# right per click and catastrophic as a bulk loader, and production is not
# changed to make a benchmark convenient.
#
# WHAT THIS BUILDER MAY AND MAY NOT WRITE
# ---------------------------------------
# PRODUCTION STILL PRODUCES, and this builder never writes:
#
#   the year columns on all three grids ...... modProfiling.SetYearColumns and
#                                              modInflation.SetYearColumns
#   the profiling ROWS and their ID keying ... modProfiling.SyncRows, which keys
#                                              a grid row to a register row by
#                                              permanent ID
#   the applied timeline defined names ....... PCCM_ApplyTimeline
#   structural validation .................... modStructuralCheck.ValidateStructure
#   EVERY calculation, simulation, sensitivity
#   and annual result ........................ the timed endpoints, untouched
#
# All of the first four arrive from ONE real `PCCM_ApplyTimeline`, which does
# exactly them and nothing else this fixture needs.
#
# THIS BUILDER WRITES ONLY WHAT A USER TYPES, plus the two identity artifacts
# production would have issued:
#
#   the register business columns, the FX rates, the Config profile names, the
#   profiling weights and the Setup scalars ... all user input
#   the permanent IDs and the two counters .... deterministic: modDrivers.AllocateId
#                                              increments the counter and formats
#                                              prefix + zero-padded sequence, so
#                                              N adds always yield CL-001..CL-00N
#                                              with the counter left at N. The
#                                              prefix and pad width are read from
#                                              the manifest's counter projection,
#                                              never from a literal here.
#
# The identity artifacts are the ONE thing here that is not a plain user input,
# and they are what the semantic-equivalence gate exists to prove - see
# `tests/phase10_fixture_equivalence.ps1`, which builds PERF-SMALL BOTH ways in
# two disposable workbooks and compares them field for field and by production's
# own calculation fingerprint.
#
# IT PUBLISHES NOTHING. There is no write to _Calc, no write to _SimData, no
# fingerprint, no state label, and no result of any kind. Controls ban each by
# name.

# ONE RECTANGULAR ASSIGNMENT. This is the whole performance argument: a
# 300 x 11 block is ONE cross-process call instead of 3,300 of them.
#
# PowerShell hands a 2-D object[,] to Value2 as a VARIANT array, which is the
# only shape Excel accepts for a block write. A jagged array of arrays is NOT
# that shape and Excel rejects it, so the array is built with New-Object
# 'object[,]' rather than by nesting @() literals.
function Set-BenchmarkRangeBlock {
    param($Workbook, [string]$SheetName, [string]$TableName,
          [int]$FirstRow, [int]$FirstColumn, $Block, [string]$Description)
    if ($Block -isnot [System.Array]) {
        throw ('the bulk write for ' + $Description + ' was handed a ' +
               $Block.GetType().FullName + ', not an array')
    }
    if ($Block.Rank -ne 2) {
        throw ('the bulk write for ' + $Description + ' was handed a rank-' +
               [string]$Block.Rank + ' array; Excel accepts only a rectangular ' +
               'two-dimensional block')
    }
    $rows = [int]$Block.GetLength(0)
    $columns = [int]$Block.GetLength(1)
    if (($rows -lt 1) -or ($columns -lt 1)) {
        throw ('the bulk write for ' + $Description + ' is empty (' + [string]$rows +
               'x' + [string]$columns + ')')
    }
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null
    $body = $null; $anchor = $null; $target = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $body = $lo.DataBodyRange
        if ($null -eq $body) {
            throw ('the bulk write for ' + $Description + ' found no body in ' + $TableName)
        }
        $anchor = $body.Cells($FirstRow, $FirstColumn)
        $target = $anchor.Resize($rows, $columns)
        # RANGE.RESIZE, NOT A LISTOBJECT OPERATION. Resizing a Range selects a
        # different rectangle of the same sheet; it does not add or remove table
        # rows or columns, so it is not the structural class protection refuses.
        $target.Value2 = $Block
    } finally {
        if ($null -ne $target)          { Release-Transient $target          'Range(block)'; $target          = $null }
        if ($null -ne $anchor)          { Release-Transient $anchor          'Range(anchor)'; $anchor         = $null }
        if ($null -ne $body)            { Release-Transient $body            'Range(body)';  $body            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';   $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects';  $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';    $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';   $localWorksheets = $null }
    }
}

# ENSURE CAPACITY. NEVER SHRINK TO THE DRIVER COUNT.
#
# WHAT EQUIVALENCE RUN 2 GOT WRONG. This took its argument as the number of body
# rows the table should END UP with, and refused when the table already held more:
#
#   FAIL|Bulk|RAISED|tblCostLines already holds 25 body rows where the fixture
#   needs 12. This builder never deletes rows...
#
# PHYSICAL CAPACITY IS NOT SEMANTIC COUNT. Stage A builds `tblCostLines` with
# `reserved_rows: 25` - twenty-five blank body rows - and twelve Cost Lines
# occupying twelve of them is not an error, it is the state production reaches.
#
# PRODUCTION'S OWN RULE, from modDrivers.AddDriver:
#
#   targetRow = FirstFreeRow(Kind, orphanRow)      a blank RESERVED row
#   If targetRow = 0 Then                          only when none is left
#       register.ListRows.Add                      ...is the table grown
#
# and the comment beside it: "Reserved rows were only ever initial capacity,
# never a business maximum." So twelve Adds into a twenty-five row table leave
# twenty-five physical rows with a thirteen-row blank suffix, and a hundred and
# eighty Adds grow it to a hundred and eighty with no suffix at all.
#
# ALL FIVE BENCHMARK-POPULATED TABLES FOLLOW THAT RULE, checked in source rather
# than assumed: modProfiling.SyncRows and modInflation.SyncProfileRows both grow
# only when `writeRow > BodyRowCount(target)` and then CLEAR the tail rather than
# delete it. No production path shrinks a body.
#
# So the target is `max(existing capacity, semantic count)`, the suffix is left
# exactly as Stage A built it, and nothing is ever deleted.
function Set-BenchmarkRegisterRowCount {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$MinimumRows)
    $current = Get-TableRowCount -Workbook $Workbook -SheetName $SheetName -TableName $TableName
    # RESERVED CAPACITY IS ENOUGH, AND IS LEFT ALONE. Shrinking to the driver count
    # would delete rows production would have kept blank, and the two fixtures
    # would then differ physically - which the equivalence snapshot compares, and
    # rightly reports as a difference.
    if ($current -ge $MinimumRows) { return $current }
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $rows = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $rows = $lo.ListRows
        for ($i = $current; $i -lt $MinimumRows; $i++) {
            $added = $null
            try { $added = $rows.Add() }
            finally { if ($null -ne $added) { Release-Transient $added 'ListRow'; $added = $null } }
        }
    } finally {
        if ($null -ne $rows)            { Release-Transient $rows            'ListRows';    $rows            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
    # PROVED, NOT ASSUMED. A grow that silently did nothing would leave the bulk
    # write landing outside the table. Exactly the requested minimum, not more: a
    # table that overshot would carry rows production never created.
    $after = Get-TableRowCount -Workbook $Workbook -SheetName $SheetName -TableName $TableName
    if ($after -ne $MinimumRows) {
        throw ($TableName + ' holds ' + [string]$after + ' body rows after growing it to at ' +
               'least ' + [string]$MinimumRows + '. ListRows.Add is a structural operation ' +
               'and is only permitted inside the fixture maintenance window.')
    }
    return $after
}

# THE IDENTIFIER PRODUCTION WOULD HAVE ISSUED, from the manifest's own counter
# projection. `modDrivers.AllocateId` reads the counter, increments it, persists
# it and formats prefix + the sequence zero-padded to the declared width - so
# the Nth add issues prefix + N and leaves the counter at N.
function Get-BenchmarkPermanentId {
    param($Counter, [int]$Sequence)
    $prefix = [string]$Counter.prefix
    $pad = [int]$Counter.pad_width
    if ([string]::IsNullOrEmpty($prefix)) {
        throw ('the manifest counter ' + [string]$Counter.defined_name + ' declares no prefix')
    }
    if ($pad -lt 1) {
        throw ('the manifest counter ' + [string]$Counter.defined_name +
               ' declares a pad width of ' + [string]$pad)
    }
    return ($prefix + ([string]$Sequence).PadLeft($pad, [char]48))
}

# THE REGISTER BODY AS ONE BLOCK, in the manifest's declared column order.
#
# COLUMN ORDINALS COME FROM THE MANIFEST, never from a count here - the same
# rule `Write-Phase5Driver` follows, and the reason a register column can be
# added to the contract without this builder inventing a position for it. A
# column no driver fills stays $null, which is a genuine blank: `category`,
# `uom` and `risk_owner` are blank in the accepted endpoint-built fixture too,
# and writing '' instead would make them populated.
function New-BenchmarkRegisterBlock {
    param($Register, $Counter, $Drivers, [bool]$IsRisk)
    $columns = @($Register.columns | ForEach-Object { [string]$_ })
    $drivers = @($Drivers)
    if ($drivers.Count -lt 1) {
        throw ('the fixture model carries no drivers for ' + [string]$Register.table_name)
    }
    $block = New-Object 'object[,]' $drivers.Count, $columns.Count
    $ids = @()
    for ($index = 0; $index -lt $drivers.Count; $index++) {
        $driver = $drivers[$index]
        $sequence = $index + 1
        $issued = Get-BenchmarkPermanentId -Counter $Counter -Sequence $sequence
        # THE MODEL'S OWN IDENTIFIER IS COMPARED, NOT TRUSTED. The emitted model
        # names CL-001.. in order and production issues in sequence; if the two
        # ever disagreed the fixture would be describing a different workbook.
        $declared = [string]$driver.permanent_id
        if ($issued -cne $declared) {
            throw ('the fixture model declares ' + $declared + ' as driver ' +
                   [string]$sequence + ' where production would issue ' + $issued)
        }
        $ids += $issued
        # The values are the accepted endpoint fixture's, field for field.
        $values = @{}
        $values[$columns[0]] = $issued
        if ($IsRisk) {
            $values['risk_name'] = ('GateB ' + $issued)
            $values['probability'] = [double]$driver.probability
            $values['impact_min'] = [double]$driver.min_value
            if ($null -ne $driver.most_likely) { $values['impact_most_likely'] = [double]$driver.most_likely }
            $values['impact_max'] = [double]$driver.max_value
        } else {
            $values['description'] = ('GateB ' + $issued)
            $values['quantity'] = [double]$driver.quantity
            $values['unit_cost_min'] = [double]$driver.min_value
            if ($null -ne $driver.most_likely) { $values['unit_cost_most_likely'] = [double]$driver.most_likely }
            $values['unit_cost_max'] = [double]$driver.max_value
        }
        $values['currency'] = [string]$driver.currency
        $values['inflation_profile'] = [string]$driver.inflation_profile
        $values['distribution'] = [string]$driver.distribution
        for ($c = 0; $c -lt $columns.Count; $c++) {
            $key = $columns[$c]
            if ($values.ContainsKey($key)) { $block[$index, $c] = $values[$key] }
            else { $block[$index, $c] = $null }
        }
    }
    return [pscustomobject]@{
        Block = $block
        Ids = $ids
        CounterValue = [double]$drivers.Count
        Columns = $columns
    }
}

# THE PROFILING WEIGHTS AS ONE BLOCK PER GRID, in the order production keyed the
# rows. The row order is READ BACK from the grid rather than assumed, because
# `modProfiling.SyncRows` rebuilds the grid from the register and nothing binds
# its physical order to the order this builder wrote the register in.
function New-BenchmarkWeightBlock {
    param($Workbook, $Grid, $Drivers, [int]$Years)
    $drivers = @($Drivers)
    $body = @(Get-TableBody -Workbook $Workbook -SheetName $Grid.sheet -TableName $Grid.table_name)
    $byId = @{}
    foreach ($driver in $drivers) { $byId[[string]$driver.permanent_id] = $driver }
    $rows = @()
    foreach ($row in $body) {
        $key = [string]$row[0]
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        if (-not $byId.ContainsKey($key)) {
            throw ('the profiling grid ' + [string]$Grid.table_name + ' carries a row for ' +
                   $key + ', which the fixture model does not declare')
        }
        $rows += $key
    }
    if ($rows.Count -ne $drivers.Count) {
        throw ('the profiling grid ' + [string]$Grid.table_name + ' carries ' +
               [string]$rows.Count + ' keyed row(s) where the fixture model declares ' +
               [string]$drivers.Count + '. PCCM_ApplyTimeline synchronises the grid from ' +
               'the register, so this means the register was not written as intended.')
    }
    # CONTIGUOUS FROM ROW 1, or the single rectangular write would straddle a gap.
    for ($r = 0; $r -lt $rows.Count; $r++) {
        if ([string]::IsNullOrWhiteSpace([string]$body[$r][0])) {
            throw ('the profiling grid ' + [string]$Grid.table_name + ' has a blank key at ' +
                   'row ' + [string]($r + 1) + ', so its keyed rows are not contiguous and a ' +
                   'block write would land on the wrong rows')
        }
    }
    $block = New-Object 'object[,]' $rows.Count, $Years
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $driver = $byId[$rows[$r]]
        $weights = @($driver.profile_weights)
        if ($weights.Count -ne $Years) {
            throw ('driver ' + $rows[$r] + ' declares ' + [string]$weights.Count +
                   ' weights for ' + [string]$Years + ' project years')
        }
        for ($c = 0; $c -lt $Years; $c++) {
            if ($null -eq $weights[$c]) { $block[$r, $c] = $null }
            else { $block[$r, $c] = [double]$weights[$c] }
        }
    }
    # A MATRIX IS NEVER THE THING A FUNCTION EMITS. `return $block` wrote the
    # rank-2 object[,] to the output stream, and PowerShell ENUMERATES a
    # multidimensional array into its elements in row-major order - so the caller's
    # assignment received an Object[] of rows x years scalars and Excel was handed a
    # rank-1 array. That is what Windows reported for tblCostProfiling, and
    # Set-BenchmarkRangeBlock's rank guard is what caught it before the write.
    #
    # THE FIX IS THE PATTERN THE REGISTER BUILDER ALREADY USES. A matrix carried as a
    # PROPERTY of a record cannot be enumerated on its way back: the pipeline emits
    # one pscustomobject, and `.Block` is the same array object the loop filled.
    # `return ,$block` would also work, but a comma is one keystroke from being
    # tidied away by someone who does not know what it is holding up.
    return [pscustomobject]@{
        Block   = $block
        Rows    = $rows.Count
        Columns = $Years
        Keys    = $rows
    }
}

# THE BULK FIXTURE, STEP FOR STEP AGAINST THE ACCEPTED ONE.
#
# The accepted `Invoke-Phase5FixtureSteps` runs A-H. This runs the same steps and
# reaches the same end state; only step F changes, and only in HOW the register
# rows arrive - not in what they contain.
#
#   A  registers empty, counters set          ASSERTED empty, then the counters
#                                             are set to N with the register block
#   B  the four Setup scalars                 identical, the accepted setter
#   C  FX reset and the foreign row           identical, the accepted helpers
#   D  the Config profile master              identical, the accepted helper
#   E/F REGISTERS, THEN ONE ApplyTimeline     the one change: the register bodies
#                                             are written as two blocks and ONE
#                                             real PCCM_ApplyTimeline synchronises
#                                             both grids, instead of N production
#                                             Add operations each re-syncing
#   G  rates and weights                      the same values, as blocks
#   H  the closing coherence check            identical, production's own report
#
# THE ORDER DIFFERENCE IS THE POINT AND IS NOT A SEMANTIC ONE. In the accepted
# path ApplyTimeline runs over EMPTY registers and each Add then syncs one grid
# row; here it runs over FULL registers and syncs all of them at once.
# `modProfiling.SyncRows` rebuilds the grid from the register by permanent ID in
# register order either way, so both paths end with the same keyed rows in the
# same order - which the equivalence gate checks rather than assumes.
function Set-BenchmarkBulkFixture {
    param($Excel, $Workbook, $Manifest, $Inspection, $Model, $ScenarioSpec)
    $years = [int]$ScenarioSpec.years

    # --- A. the registers must START empty ---------------------------------
    # Not cleared - ASSERTED. This builder writes permanent IDs from sequence 1,
    # so a register that already carried rows would be given duplicates, and the
    # counter would disagree with the highest identifier present.
    $registerByKey = @{}
    foreach ($register in @($Manifest.registers)) {
        $registerByKey[[string]$register.key] = $register
        $existing = @(Get-IdColumnValues -Workbook $Workbook -Info $register)
        if ($existing.Count -ne 0) {
            throw ('the bulk fixture requires an empty ' + [string]$register.table_name +
                   ' and found ' + [string]$existing.Count + ' keyed row(s). It writes ' +
                   'identifiers from sequence 1 and runs against a fresh disposable ' +
                   'workbook.')
        }
    }
    $counterByRegister = @{}
    foreach ($counter in @($Manifest.counters)) {
        $counterByRegister[[string]$counter.driver_register] = $counter
    }
    foreach ($key in @('cost_lines', 'risk_register')) {
        if (-not $registerByKey.ContainsKey($key)) {
            throw ('the manifest declares no register ' + $key)
        }
        if (-not $counterByRegister.ContainsKey($key)) {
            throw ('the manifest declares no identity counter for ' + $key)
        }
    }

    # --- B. the Setup scalars, through the accepted setter ------------------
    $inputs = $Inspection.inputs
    Set-NamedValue -Workbook $Workbook -DefinedName $inputs.base_year.defined_name `
        -Value ([double]$Model.timeline.base_year)
    Set-NamedValue -Workbook $Workbook -DefinedName $inputs.project_start_year.defined_name `
        -Value ([double]$Model.timeline.start_year)
    Set-NamedValue -Workbook $Workbook -DefinedName $inputs.duration_years.defined_name `
        -Value ([double]$Model.timeline.duration)
    Set-NamedValue -Workbook $Workbook -DefinedName $inputs.discount_rate.defined_name `
        -Value ([double]$Model.discount_rate)

    # --- C. FX, through the accepted reset and append ----------------------
    # TWENTY-FOUR CELLS. Left cell-by-cell on purpose: it is already bounded by
    # the contract's twelve reserved rows, the reset is a settled control over a
    # defective build, and replacing it would put that control at risk for no
    # measurable gain.
    $fx = $Inspection.input_tables.fx_rates
    Reset-Phase5FxTable -Workbook $Workbook -Inspection $Inspection `
        -Seed (Get-Phase5LockedFxSeed)
    $reporting = [string](Get-NamedValue -Workbook $Workbook `
        -DefinedName $inputs.reporting_currency.defined_name)
    $fxRow = 0
    foreach ($entry in @($Model.fx)) {
        if ([string]$entry.currency -eq $reporting) { continue }
        $fxRow++
        $null = Add-BlankTableRow -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name
        $target = [int]$fx.locked_seed_rows + $fxRow
        Set-TableCell -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name `
            -RowIndex $target -ColumnIndex 1 -Value ([string]$entry.currency)
        if ($null -ne $entry.rate) {
            Set-TableCell -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name `
                -RowIndex $target -ColumnIndex 2 -Value ([double]$entry.rate)
        }
    }

    # --- D. the Config profile master, through the accepted helper ----------
    Set-Phase5InflationProfileMaster -Workbook $Workbook -Inspection $Inspection `
        -Profiles @($Model.inflation.PSObject.Properties.Name)

    # --- E. THE REGISTERS, AS TWO BLOCKS ------------------------------------
    $blocks = @{}
    foreach ($pair in @(
            @{ key = 'cost_lines'; drivers = @($Model.cost_lines); risk = $false },
            @{ key = 'risk_register'; drivers = @($Model.risks); risk = $true })) {
        $register = $registerByKey[$pair.key]
        $counter = $counterByRegister[$pair.key]
        $prepared = New-BenchmarkRegisterBlock -Register $register -Counter $counter `
            -Drivers $pair.drivers -IsRisk ([bool]$pair.risk)
        # THE DRIVER COUNT IS A FLOOR, NOT A TARGET. Stage A's reserved capacity
        # stands when it is already enough, exactly as it does after N production
        # Adds, and the blank suffix is left where the contract put it.
        $physical = Set-BenchmarkRegisterRowCount -Workbook $Workbook -SheetName $register.sheet `
            -TableName $register.table_name -MinimumRows @($pair.drivers).Count
        if ($physical -lt @($pair.drivers).Count) {
            throw ([string]$register.table_name + ' holds ' + [string]$physical +
                   ' body rows, fewer than the ' + [string]@($pair.drivers).Count +
                   ' the fixture must populate')
        }
        # THE SAME GEOMETRY PROOF. This block never collapsed - it travels as a
        # record property and always has - but the caller asserting what it asked for
        # is what keeps that true rather than lucky.
        if ([int]$prepared.Block.GetLength(0) -ne @($pair.drivers).Count) {
            throw ([string]$register.table_name + ' block has ' +
                   [string]$prepared.Block.GetLength(0) + ' rows where the fixture ' +
                   'declares ' + [string]@($pair.drivers).Count)
        }
        if ([int]$prepared.Block.GetLength(1) -ne @($prepared.Columns).Count) {
            throw ([string]$register.table_name + ' block has ' +
                   [string]$prepared.Block.GetLength(1) + ' columns where the contract ' +
                   'declares ' + [string]@($prepared.Columns).Count)
        }
        Set-BenchmarkRangeBlock -Workbook $Workbook -SheetName $register.sheet `
            -TableName $register.table_name -FirstRow 1 -FirstColumn 1 -Block $prepared.Block `
            -Description ([string]$register.table_name)
        # THE COUNTER IS THE MODEL'S RECORD OF EVERY IDENTIFIER EVER ISSUED, and
        # after N adds it holds N. modDrivers.TryReadCounter refuses anything that
        # is not a whole number, so a counter that landed as text would make the
        # very next production mutation refuse.
        Set-NamedValue -Workbook $Workbook -DefinedName ([string]$counter.defined_name) `
            -Value ([double]$prepared.CounterValue)
        $blocks[$pair.key] = $prepared
    }
    # AND THE REGISTERS REALLY CARRY WHAT WAS WRITTEN, before a production command
    # is asked to synchronise anything from them.
    foreach ($pair in @(
            @{ key = 'cost_lines'; drivers = @($Model.cost_lines) },
            @{ key = 'risk_register'; drivers = @($Model.risks) })) {
        $register = $registerByKey[$pair.key]
        $ids = @(Get-IdColumnValues -Workbook $Workbook -Info $register)
        $expected = @($blocks[$pair.key].Ids)
        if ($ids.Count -ne $expected.Count) {
            throw ([string]$register.table_name + ' carries ' + [string]$ids.Count +
                   ' keyed row(s) after the block write, not ' + [string]$expected.Count)
        }
        for ($i = 0; $i -lt $expected.Count; $i++) {
            if ([string]$ids[$i] -cne [string]$expected[$i]) {
                throw ([string]$register.table_name + ' row ' + [string]($i + 1) + ' carries ' +
                       [string]$ids[$i] + ' where the fixture wrote ' + [string]$expected[$i])
            }
        }
        # AND THE RESERVED SUFFIX IS STILL BLANK. The block write covers exactly the
        # semantic rows; anything below them must be as Stage A left it, because a
        # value there is a row production never keyed - which modDrivers.AddDriver
        # reports as an orphan and refuses to mutate over.
        $body = @(Get-TableBody -Workbook $Workbook -SheetName $register.sheet `
            -TableName $register.table_name)
        for ($row = $expected.Count; $row -lt $body.Count; $row++) {
            foreach ($value in @($body[$row])) {
                if ([string]$value -ne '') {
                    throw ([string]$register.table_name + ' row ' + [string]($row + 1) +
                           ' is a reserved row below the ' + [string]$expected.Count +
                           ' semantic drivers and carries ' + [char]39 + [string]$value +
                           [char]39 + '. Production leaves reserved rows blank, and a ' +
                           'populated unkeyed row is the orphan it refuses to mutate over.')
                }
            }
        }
    }

    # --- F. ONE REAL PCCM_ApplyTimeline -------------------------------------
    # This is where every structural thing the fixture needs comes from
    # production: the year columns on all three grids, one profiling row per
    # register row keyed by permanent ID, the applied-timeline defined names, and
    # modStructuralCheck.ValidateStructure. Nothing here reproduces any of it.
    $applied = Invoke-Phase5ProductionOperation -Excel $Excel `
        -Operation 'PCCM_ApplyTimeline' -Stage 'the bulk fixture structural baseline'
    $null = Assert-Phase5StructurallyCoherent -Excel $Excel `
        -Stage 'after the bulk fixture applied the timeline'

    # --- G. the rates and the weights, as blocks ----------------------------
    $inflationGrid = $null
    $gridByKey = @{}
    foreach ($grid in @($Manifest.grids)) {
        $gridByKey[[string]$grid.key] = $grid
        if ([string]$grid.key -eq 'inflation') { $inflationGrid = $grid }
    }
    if ($null -eq $inflationGrid) { throw 'the manifest declares no inflation grid' }
    Write-Phase5InflationRates -Workbook $Workbook -Manifest $Manifest -Model $Model

    foreach ($pair in @(
            @{ key = 'cost_profiling'; drivers = @($Model.cost_lines) },
            @{ key = 'risk_profiling'; drivers = @($Model.risks) })) {
        if (-not $gridByKey.ContainsKey($pair.key)) {
            throw ('the manifest declares no grid ' + $pair.key)
        }
        $grid = $gridByKey[$pair.key]
        $fixed = @($grid.fixed_columns).Count
        $prepared = New-BenchmarkWeightBlock -Workbook $Workbook -Grid $grid `
            -Drivers $pair.drivers -Years $years
        # THE GEOMETRY IS PROVED AGAINST WHAT THIS CALLER ASKED FOR, not merely
        # against whatever arrived. Set-BenchmarkRangeBlock resizes the anchor to the
        # block's own dimensions, so a block of the wrong shape would write a
        # perfectly consistent rectangle in the wrong place.
        if ([int]$prepared.Rows -ne @($pair.drivers).Count) {
            throw ('the weight block for ' + [string]$grid.table_name + ' has ' +
                   [string]$prepared.Rows + ' rows where the fixture declares ' +
                   [string]@($pair.drivers).Count)
        }
        if ([int]$prepared.Columns -ne $years) {
            throw ('the weight block for ' + [string]$grid.table_name + ' has ' +
                   [string]$prepared.Columns + ' columns where the model runs for ' +
                   [string]$years + ' year(s)')
        }
        Set-BenchmarkRangeBlock -Workbook $Workbook -SheetName $grid.sheet `
            -TableName $grid.table_name -FirstRow 1 -FirstColumn ($fixed + 1) `
            -Block $prepared.Block -Description ([string]$grid.table_name)
    }

    # --- H. THE FIXTURE ENDS COHERENT ---------------------------------------
    $null = Assert-Phase5StructurallyCoherent -Excel $Excel `
        -Stage 'at the end of bulk fixture establishment'
    return $applied
}

# ===========================================================================
# THE FIXTURE MAINTENANCE WINDOW
# ===========================================================================
# WHAT THE WINDOWS RUN AT 7077608 SETTLED. The content-based reset is correct
# and it is still refused, because the refusal is not about deletion:
#
#   statement : $null = $cell.ClearContents()
#   message   : The cell or chart you're trying to change is on a protected
#               sheet. To make a change, unprotect the sheet.
#
# An external COM caller cannot clear a cell on a protected sheet. A VALUE write
# from the same caller CAN - Benchmark Run 3 wrote four Setup scalars that way
# and they succeeded - so `Value2 =` and `ClearContents` are different
# capabilities to an out-of-process client, and only the second one needs help.
#
# THE WINDOW IS PRODUCTION'S OWN, NOT A SECOND ONE. `modProtection` already owns
# a depth-counted structural window that releases WORKSHEET protection only and
# never touches workbook structure. Nothing here re-implements it: there is no
# `.Unprotect` in this runner, and a control refuses one.
#
# WHY A SHIM. `ProtectionBeginStructural` and `ProtectionEndStructural` both take
# `ByRef detail As String`. Every procedure any accepted harness in this tree has
# ever reached through `Application.Run` - production and Gate-B diagnostic alike
# - takes ByVal parameters or none, so there is no precedent for marshalling a
# ByRef out-parameter across that boundary and no reason to discover its failure
# mode on Windows, on the one call whose failure has to abort the run.
# `phase10_fixture_window.bas` owns the String inside VBA and returns a String,
# which is the shape that is proven. It is imported into the DISPOSABLE copy
# exactly as `phase5_gate_b_diagnostics.bas` has been since Phase 5, and it is
# never declared in the manifest.
#
# SETUP ONLY, AND PROVABLY SO. The window opens immediately before
# `Set-Phase5Fixture` and closes immediately after it, inside the fixture
# stopwatch, whose figure the runner already reports as setup and uses in no
# measurement. It is closed - and the closure verified - before the first timed
# operation exists. A close that fails throws, and a throw here reaches the
# abandon path, so no sample can be recorded from a workbook whose protection
# was not restored.
$script:FixtureWindowModule = 'modPhase10FixtureWindow'
$script:FixtureWindowSource = 'phase10_fixture_window.bas'

function Import-BenchmarkFixtureWindow {
    param($Excel, $Workbook, $Manifest, [string]$ScriptDir)
    $source = Join-Path $ScriptDir $script:FixtureWindowSource
    if (-not (Test-Path -LiteralPath $source)) {
        throw ('the fixture-window shim ' + $source + ' is missing, so the benchmark ' +
               'cannot open the accepted protection window')
    }
    # IT IS NOT PRODUCTION, AND THAT IS CHECKED RATHER THAN SAID. A module the
    # manifest declares would be a production module, and importing one over the
    # built project would be replacing production at runtime.
    $declared = @($Manifest.vba.modules | ForEach-Object { [string]$_.name })
    if ($declared -contains $script:FixtureWindowModule) {
        throw ('the manifest declares ' + $script:FixtureWindowModule + ' as a production ' +
               'module; the benchmark will not import a test module over production')
    }
    $project = $null; $components = $null; $imported = $null
    try {
        $project = $Workbook.VBProject
        $components = $project.VBComponents
        $imported = $components.Import($source)
        $name = [string]$imported.Name
        if ($name -ne $script:FixtureWindowModule) {
            throw ('the fixture-window shim imported as ' + $name + ', not ' +
                   $script:FixtureWindowModule)
        }
    } finally {
        if ($null -ne $imported)   { Release-Transient $imported   'VBComponent(shim)'; $imported   = $null }
        if ($null -ne $components) { Release-Transient $components 'VBComponents';      $components = $null }
        if ($null -ne $project)    { Release-Transient $project    'VBProject';         $project    = $null }
    }
    # AND IT ANSWERS. An import that reported success and a project that will not
    # call it are different things.
    $ping = [string]$Excel.Run('P10FW_Ping')
    if ($ping -ne ('OK|' + $script:FixtureWindowModule)) {
        throw ('the fixture-window shim imported but does not answer: ' + $ping)
    }
}

# THE STATE, PARSED FROM ONE STRING. Read in VBA and returned as text, because
# probe Run 8 proved `Worksheet.ProtectContents` across COM can be a terminating
# PropertyNotFoundException on an object PowerShell has no type information for.
function Get-BenchmarkProtectionState {
    param($Excel)
    $raw = [string]$Excel.Run('P10FW_State')
    if ($raw -notlike 'OK|*') {
        throw ('the workbook could not report its protection state: ' + $raw)
    }
    $state = @{}
    foreach ($field in $raw.Substring(3).Split([char]124)) {
        $pair = $field.Split([char]61)
        if ($pair.Count -eq 2) { $state[$pair[0]] = $pair[1] }
    }
    foreach ($required in @('applied', 'depth', 'structure', 'sheets', 'protected')) {
        if (-not $state.ContainsKey($required)) {
            throw ('the protection state is missing ' + $required + ': ' + $raw)
        }
    }
    return [pscustomobject]@{
        Applied   = ([string]$state['applied'] -eq 'True')
        Depth     = [int]$state['depth']
        Structure = ([string]$state['structure'] -eq 'True')
        Sheets    = [int]$state['sheets']
        Protected = [int]$state['protected']
        Raw       = $raw
    }
}

# THE FULL ACCEPTED STATE, AGAINST THE DECLARED ONE. The sheet count comes from
# `stage_b_manifest.json`'s protection projection, never from a literal here, so
# a contract that gained a worksheet is a contract change rather than a silent
# pass at the old number.
function Assert-BenchmarkProtectionApplied {
    param($Excel, $Manifest, [string]$Stage)
    $expected = @($Manifest.protection.sheets).Count
    $state = Get-BenchmarkProtectionState -Excel $Excel
    $problems = @()
    if (-not $state.Applied)                { $problems += 'the workbook does not report itself protected' }
    if (-not $state.Structure)              { $problems += 'workbook structure protection is not applied' }
    if ($state.Sheets -ne $expected)        { $problems += ('the workbook holds ' + [string]$state.Sheets +
                                                            ' worksheets where the manifest declares ' +
                                                            [string]$expected) }
    if ($state.Protected -ne $expected)     { $problems += ([string]$state.Protected + ' of ' +
                                                            [string]$expected + ' worksheets are protected') }
    if ($state.Depth -ne 0)                 { $problems += ('the structural window is still open at depth ' +
                                                            [string]$state.Depth) }
    if ($problems.Count -gt 0) {
        throw ('protection is not in the accepted state ' + $Stage + ': ' +
               ($problems -join '; ') + '. Reported: ' + $state.Raw)
    }
    return $state
}

# OPENING IS A TRANSACTION, AND IT OWNS ITS OWN ROLLBACK.
#
# THE GAP THIS CLOSES. `P10FW_Begin` can succeed - the window is open, every
# worksheet released - and one of this function's OWN post-open checks can then
# throw. The caller has not reached its try/finally yet, so nothing would close
# what was opened, and the error would escape to the abandon path leaving a
# half-open window behind. `modProtection` names that outcome the worst one there
# is, and it is not made acceptable by the workbook being disposable: the
# contract is that once Begin succeeds, exactly one compensating End is attempted
# before the failure escapes.
#
# A `catch`, NOT A `finally`. A finally that threw would REPLACE the original
# exception, and the reason the window is being rolled back is the thing worth
# reporting. The catch composes both into one sentence instead: what failed, and
# whether the rollback took.
#
# AND ROLLBACK RUNS ONLY ON FAILURE. `return` inside a `try` does not enter its
# `catch`, so a successful open closes nothing and hands the open window to the
# caller, whose try/finally performs the one normal close. There is no path on
# which both this function and the caller close the same window, so the depth
# cannot be decremented twice.
function Open-BenchmarkFixtureWindow {
    param($Excel, $Manifest)
    # THE STATE BEFORE IS PROVED FIRST. Opening a window over a workbook that was
    # already unprotected would close onto a state nobody established.
    #
    # NOTHING IS OWED YET. Every failure above the Begin below leaves the workbook
    # exactly as it was found, so there is nothing to compensate and no End is
    # called - a compensating close here would decrement a depth nobody raised.
    $null = Assert-BenchmarkProtectionApplied -Excel $Excel -Manifest $Manifest `
        -Stage 'before the fixture maintenance window was opened'
    $reply = [string]$Excel.Run('P10FW_Begin')
    if ($reply -notlike 'OK|*') {
        throw ('the fixture maintenance window could not be opened: ' + $reply)
    }

    # FROM HERE THE WINDOW IS OPEN AND EVERY FAILURE MUST COMPENSATE.
    try {
        $state = Get-BenchmarkProtectionState -Excel $Excel
        if ($state.Depth -ne 1) {
            throw ('the fixture maintenance window opened to depth ' + [string]$state.Depth +
                   ', not 1. The benchmark opens exactly one outer window and closes it.')
        }
        # WORKBOOK STRUCTURE PROTECTION IS NOT PART OF THE WINDOW AND MUST NOT
        # MOVE. `Structure` is `ThisWorkbook.ProtectStructure`, not worksheet
        # protection - the window releases worksheets and nothing else, and the
        # only `ThisWorkbook.Unprotect` in production lives in the maintenance
        # path this harness never calls.
        if (-not $state.Structure) {
            throw ('opening the fixture maintenance window released WORKBOOK STRUCTURE ' +
                   'protection, which it must never do. Reported: ' + $state.Raw)
        }
        return $state
    } catch {
        $original = [string]$_.Exception.Message
        $recovery = Invoke-BenchmarkWindowRollback -Excel $Excel
        throw ($original + ' The window was open when this failed, so a compensating ' +
               'close was attempted: ' + $recovery)
    }
}

# THE ROLLBACK ITSELF, WHICH MUST NOT RAISE.
#
# It is called from a catch block that is about to rethrow, and a rollback that
# threw would destroy the diagnosis it exists to accompany. So it reports in a
# sentence instead - and it reports a REFUSAL and a RAISE differently, because
# "the workbook says it could not restore protection" and "the call never
# arrived" are different facts about the machine.
#
# NOT A CATCH-AND-IGNORE. Nothing here is swallowed: every outcome becomes text
# that the caller concatenates into the exception it throws.
function Invoke-BenchmarkWindowRollback {
    param($Excel)
    try {
        $reply = [string]$Excel.Run('P10FW_End')
        if ($reply -notlike 'OK|*') {
            return ('IT REFUSED, so worksheet protection is NOT restored and this ' +
                    'workbook must not be measured - ' + $reply)
        }
        return ('it succeeded and protection is restored - ' + $reply)
    } catch {
        return ('IT RAISED, so worksheet protection is NOT restored and this workbook ' +
                'must not be measured - ' + [string]$_.Exception.Message)
    }
}

function Close-BenchmarkFixtureWindow {
    param($Excel, $Manifest)
    $reply = [string]$Excel.Run('P10FW_End')
    if ($reply -notlike 'OK|*') {
        throw ('the fixture maintenance window could not be closed and protection was ' +
               'not restored: ' + $reply + '. No measurement may be taken from this ' +
               'workbook.')
    }
    # AND THE OWNER'S WORD IS NOT THE END OF IT. `ProtectionEndStructural` already
    # requires `ProtectionIsApplied` before reporting success; this asks the
    # workbook again, independently, and counts the sheets.
    return (Assert-BenchmarkProtectionApplied -Excel $Excel -Manifest $Manifest `
        -Stage 'after the fixture maintenance window was closed')
}

function Get-IdColumnValues {
    param($Workbook, $Info)
    $out = @()
    foreach ($row in @(Get-TableBody -Workbook $Workbook -SheetName $Info.sheet -TableName $Info.table_name)) {
        if ($row[0] -ne '') { $out += $row[0] }
    }
    return $out
}

# ===========================================================================
# THE PHASE-7 ADDRESSES
# ===========================================================================
# `phase6_gate_b_inspection.json` is a PINNED artefact whose SHA-256 the
# Phase-6 Gate-B fixture-integrity checks depend on. Extending it to carry the


# ===========================================================================
# WHERE THE RUN HAD GOT TO
# ===========================================================================
# W1 DIAGNOSTIC ROOT. Windows Run 1 reached the outer handler and said only
# "The property 'Value' cannot be found on this object." It did not say which
# stage it was in, which is the one thing that would have located it in a
# thousand-line script from a console transcript.
#
# So the run carries a CURSOR. It is set at stage boundaries and at each timed
# execution, and it is READ ONLY BY THE FAILURE PATH - never inside a measured
# interval, and never by a successful run, so it can change nothing about a
# timing. The whole cursor costs four string assignments per operation.
$script:BenchmarkCursor = [pscustomobject]@{
    Stage      = 'startup'
    Action     = 'loading the harness'
    Scenario   = ''
    Operation  = ''
    Iterations = ''
    Phase      = ''
}

function Set-BenchmarkStage {
    param([string]$Stage, [string]$Action = '')
    $script:BenchmarkCursor.Stage = $Stage
    $script:BenchmarkCursor.Action = $Action
}

function Set-BenchmarkOperationContext {
    param([string]$Scenario = '', [string]$Operation = '', $Iterations = $null,
          [string]$Phase = '')
    $script:BenchmarkCursor.Scenario = $Scenario
    $script:BenchmarkCursor.Operation = $Operation
    $script:BenchmarkCursor.Iterations = $(if ($null -eq $Iterations) { 'n/a' }
                                           else { [string]$Iterations })
    $script:BenchmarkCursor.Phase = $Phase
}

# WHAT FAILED, WHERE, AND WHAT THE RUN WAS DOING AT THE TIME. Structured, so the
# same facts reach the console, the log and the JSON artifact rather than only
# the first of the three.
function New-BenchmarkFailureRecord {
    param($ErrorRecord, [string]$Fallback = '')
    $cursor = $script:BenchmarkCursor
    $record = New-Object System.Collections.Specialized.OrderedDictionary
    $record.Add('stage', [string]$cursor.Stage)
    $record.Add('action', [string]$cursor.Action)
    $record.Add('scenario', [string]$cursor.Scenario)
    $record.Add('operation', [string]$cursor.Operation)
    $record.Add('iterations', [string]$cursor.Iterations)
    $record.Add('execution_phase', [string]$cursor.Phase)
    if ($null -eq $ErrorRecord) {
        $record.Add('exception_type', '')
        $record.Add('message', $Fallback)
        $record.Add('script_line_number', 0)
        $record.Add('script_line', '')
        $record.Add('command', '')
        return $record
    }
    $exception = Get-BenchmarkProperty -InputObject $ErrorRecord -Name 'Exception'
    $invocation = Get-BenchmarkProperty -InputObject $ErrorRecord -Name 'InvocationInfo'
    $record.Add('exception_type',
                $(if ($null -eq $exception) { Get-BenchmarkUnavailable }
                  else { $exception.GetType().FullName }))
    $record.Add('message', [string](Format-BenchmarkFact (
        Get-BenchmarkProperty -InputObject $exception -Name 'Message')))
    $record.Add('script_line_number', [int](Format-BenchmarkNumber (
        Get-BenchmarkProperty -InputObject $invocation -Name 'ScriptLineNumber')))
    $line = [string](Get-BenchmarkProperty -InputObject $invocation -Name 'Line')
    $record.Add('script_line', $line.Trim())
    $record.Add('command', [string](Format-BenchmarkFact (
        Get-BenchmarkProperty -InputObject $invocation -Name 'MyCommand')))
    return $record
}

function Format-BenchmarkNumber {
    param($Value)
    if ($null -eq $Value) { return 0 }
    return [int]$Value
}

function Format-BenchmarkFailure {
    param($Record)
    $lines = @('  stage                : ' + [string]$Record['stage'],
               '  doing                : ' + [string]$Record['action'],
               '  scenario             : ' + [string]$Record['scenario'],
               '  operation            : ' + [string]$Record['operation'],
               '  iterations           : ' + [string]$Record['iterations'],
               '  execution phase      : ' + [string]$Record['execution_phase'],
               '  exception            : ' + [string]$Record['exception_type'],
               '  message              : ' + [string]$Record['message'],
               '  at line              : ' + [string]$Record['script_line_number'],
               '  statement            : ' + [string]$Record['script_line'],
               '  command              : ' + [string]$Record['command'])
    return ($lines -join "`r`n")
}


# ===========================================================================
# THE REPORT SINK
# ===========================================================================
$script:BenchmarkLines = New-Object System.Collections.ArrayList
$script:BenchmarkTextPath = ''

function Write-BenchmarkLine {
    param([string]$Text = '')
    $null = $script:BenchmarkLines.Add($Text)
    Write-Host $Text
    # WRITTEN THROUGH ON EVERY LINE. A run that is stopped or killed must still
    # leave every measurement it had already taken on disk.
    if (-not [string]::IsNullOrWhiteSpace($script:BenchmarkTextPath)) {
        try {
            Set-Content -LiteralPath $script:BenchmarkTextPath `
                -Value ($script:BenchmarkLines -join "`r`n") -Encoding UTF8
        } catch {
            # THE CONSOLE LINE ALREADY WENT OUT. Losing the write-through means
            # the log file is behind, which is worth saying once - and it must
            # not recurse back into this function.
            $script:BenchmarkTextPath = ''
            Write-Host ('  (the run log could not be written through: ' +
                        $_.Exception.Message + ')') -ForegroundColor DarkYellow
        }
    }
}

function Format-BenchmarkSeconds {
    param([double]$Milliseconds)
    return ('{0:N3} s ({1:N0} ms)' -f ($Milliseconds / 1000.0), $Milliseconds)
}

# THE COMPARISON STATISTIC. Three warm samples, sorted, middle one. The mean is
# deliberately not offered: it would average away the outlier a scheduler or a
# background process introduces, and that outlier is information.
function Get-BenchmarkMedian {
    param([double[]]$Values)
    $sorted = @($Values | Sort-Object)
    if ($sorted.Count -eq 0) { return $null }
    $middle = [int][Math]::Floor($sorted.Count / 2)
    if (($sorted.Count % 2) -eq 1) { return [double]$sorted[$middle] }
    return [double](([double]$sorted[$middle - 1] + [double]$sorted[$middle]) / 2.0)
}


# ===========================================================================
# THE SOURCE REVISION
# ===========================================================================
# FAIL CLOSED. If `pccm/src`, `pccm/spec` or `pccm/builder` is dirty, the
# workbook this run would build cannot be attributed to a revision, and a
# measurement nobody can attribute is not evidence. Same policy and same
# pathspecs as the accepted Phase-7 measurement.
function Get-BenchmarkSourceRevision {
    param([string]$RepoRoot)
    $head = ''
    try { $head = [string](& git -C $RepoRoot rev-parse HEAD 2>$null) } catch { $head = '' }
    $head = $head.Trim()
    if ([string]::IsNullOrWhiteSpace($head)) {
        throw ('git could not report HEAD for ' + $RepoRoot +
               '; the benchmark workbook cannot be attributed to a source revision')
    }
    $dirty = @()
    foreach ($pathspec in @('pccm/src', 'pccm/spec', 'pccm/builder')) {
        $lines = @()
        try { $lines = @(& git -C $RepoRoot status --porcelain -- $pathspec 2>$null) } catch { $lines = @() }
        foreach ($line in $lines) {
            if (-not [string]::IsNullOrWhiteSpace([string]$line)) { $dirty += [string]$line }
        }
    }
    return [pscustomobject]@{ Head = $head; Dirty = $dirty }
}

# ===========================================================================
# THE ENVIRONMENT RECORD
# ===========================================================================
# A TIMING WITHOUT THIS IS A NUMBER WITH NO INTERPRETATION. Every field the plan
# names is present in the output; one the machine cannot answer is recorded as
# 'unavailable', never omitted - so a missing field always means the harness did
# not ask, rather than the machine did not say.
function Get-BenchmarkUnavailable { return 'unavailable' }

# W1 ROOT CAUSE, AND THE ONE PLACE A PROPERTY IS NOW READ OFF SOMETHING THAT
# MIGHT NOT BE THERE.
#
# Windows Run 1 aborted here, before a single operation was timed:
#
#     $value = [string](Get-Item -LiteralPath ('Env:' + $name) `
#                       -ErrorAction SilentlyContinue).Value
#
# `Get-Item Env:OneDriveCommercial` on a machine with a CONSUMER OneDrive and no
# work account writes a suppressed ItemNotFoundException and EMITS NOTHING. A
# pipeline that emitted nothing is $null in an expression, and under
# `Set-StrictMode -Version 2.0` `$null.Value` is not a quiet $null - it is a
# terminating PropertyNotFoundException with FullyQualifiedErrorId
# PropertyNotFoundStrict. The harness asked an optional environment variable for
# a property before establishing that the variable existed.
#
# NORMALISE, VALIDATE, THEN USE. `@()` turns "emitted nothing" into an empty
# array rather than $null, and the property is read only from an object that is
# proved to carry it. $null here is the ABSENCE OF A FACT and never a substitute
# for one: every caller either records it as unavailable or refuses.
function Get-BenchmarkProperty {
    param($InputObject, [string]$Name)
    if ($null -eq $InputObject) { return $null }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

# A property that MUST be there. Absent is a refusal with the shape named, never
# a default - a benchmark that proceeded on a substituted value would measure a
# scenario nobody asked for.
function Get-BenchmarkRequiredProperty {
    param($InputObject, [string]$Name, [string]$Where)
    if ($null -eq $InputObject) {
        throw ($Where + ": expected an object carrying '" + $Name + "' and got nothing")
    }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) {
        throw ($Where + ": the object is a " + $InputObject.GetType().FullName +
               " and carries no '" + $Name + "' property. It is not defaulted.")
    }
    return $property.Value
}

# A FACT THE MACHINE ANSWERED, OR THE WORD FOR "IT DID NOT". Never an empty
# string and never a silently omitted key: a missing field must always mean the
# harness did not ask.
function Format-BenchmarkFact {
    param($Value)
    if ($null -eq $Value) { return (Get-BenchmarkUnavailable) }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return (Get-BenchmarkUnavailable) }
    return $text
}

function Get-BenchmarkOneDriveRoots {
    $roots = @()
    foreach ($name in @('OneDrive', 'OneDriveCommercial', 'OneDriveConsumer')) {
        # AN ABSENT VARIABLE CONTRIBUTES NOTHING, and that is a fact about the
        # machine rather than a hole in the record.
        foreach ($item in @(Get-Item -LiteralPath ('Env:' + $name) -ErrorAction SilentlyContinue)) {
            $value = [string](Get-BenchmarkProperty -InputObject $item -Name 'Value')
            if (-not [string]::IsNullOrWhiteSpace($value)) { $roots += $value }
        }
    }
    return ,@($roots | Select-Object -Unique)
}

# WHAT KIND OF PLACE THIS PATH IS. An observation about the path and nothing
# more: no claim is made here about what any of them does to a measurement.
function Get-BenchmarkLocationType {
    param([string]$Path, [string[]]$OneDriveRoots)
    if ([string]::IsNullOrWhiteSpace($Path)) { return 'unknown' }
    $full = ''
    try { $full = [System.IO.Path]::GetFullPath($Path) } catch { $full = $Path }
    foreach ($root in @($OneDriveRoots)) {
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        if ($full.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
            return 'onedrive'
        }
    }
    if ($full -match '(?i)[\\/](OneDrive|Dropbox|Google Drive|Box|iCloudDrive)[\\/-]') {
        return 'synced'
    }
    try {
        $qualifier = [System.IO.Path]::GetPathRoot($full)
        if ($full.StartsWith('\\')) { return 'synced' }
        foreach ($drive in @(Get-CimInstance -ClassName Win32_LogicalDisk `
                -Filter ("DeviceID='" + $qualifier.TrimEnd('\') + "'") -ErrorAction SilentlyContinue)) {
            $type = Get-BenchmarkProperty -InputObject $drive -Name 'DriveType'
            if (($null -ne $type) -and ([int]$type -eq 4)) { return 'synced' }
        }
    } catch {
        # UNCLASSIFIABLE IS NOT LOCAL. Returning 'local' after a failed lookup
        # would state the very thing the contract asked never to be assumed - a
        # synced repository silently recorded as an unsynchronised path.
        return 'unknown'
    }
    return 'local'
}

# EXCEL'S BITNESS IS EXCEL'S, NOT POWERSHELL'S.
#
# The first draft recorded `[System.Environment]::Is64BitProcess`, which is the
# bitness of the POWERSHELL HOST. A 64-bit host automating a 32-bit Excel would
# have been recorded as 64-bit Excel - wrong evidence in the one field a reader
# would use to say which build was exercised. Windows Run 1 never reached this
# line, so no evidence is invalidated by correcting it.
#
# The answer is derived from the EXCEL PROCESS'S OWN IMAGE PATH, and the path is
# recorded beside the word so the derivation is auditable rather than asserted.
function Get-BenchmarkExcelImage {
    param($Identity)
    $path = ''
    $processId = [int](Get-BenchmarkProperty -InputObject $Identity -Name 'ProcessId')
    if ($processId -gt 0) {
        foreach ($process in @(Get-Process -Id $processId -ErrorAction SilentlyContinue)) {
            $candidate = [string](Get-BenchmarkProperty -InputObject $process -Name 'Path')
            if (-not [string]::IsNullOrWhiteSpace($candidate)) { $path = $candidate }
        }
    }
    $bitness = Get-BenchmarkUnavailable
    if (-not [string]::IsNullOrWhiteSpace($path)) {
        $bitness = $(if ($path -match '(?i)\\Program Files \(x86\)\\') { '32-bit' } else { '64-bit' })
        $bitness = $bitness + ' (derived from the Excel image path)'
    }
    return [pscustomobject]@{ Path = $path; Bitness = $bitness }
}

function Get-BenchmarkEnvironment {
    param($Excel, $Identity, [string]$WorkbookPath, [string]$RepositoryPath,
          $ReleaseIdentity, [string]$HarnessVersion, [int]$SchemaVersion, $Revision)
    $unknown = Get-BenchmarkUnavailable
    # ASSIGNED, NOT RE-WRAPPED - the same two-place contract as the problem
    # list. Double-wrapped, every OneDrive root became one nested array, so a
    # workbook under OneDrive was recorded as a LOCAL location and the roots
    # field held a list containing a list.
    $roots = Get-BenchmarkOneDriveRoots

    # NORMALISED THE SAME WAY THE ONEDRIVE ROOTS ARE. A CIM query that matches
    # nothing emits nothing, and `@()` turns that into an empty collection
    # instead of a $null whose properties StrictMode refuses to read.
    $os = $null; $cpu = $null; $computer = $null
    foreach ($item in @(Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue)) { $os = $item }
    foreach ($item in @(Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue)) { if ($null -eq $cpu) { $cpu = $item } }
    foreach ($item in @(Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue)) { $computer = $item }

    $ram = $unknown
    $memory = Get-BenchmarkProperty -InputObject $computer -Name 'TotalPhysicalMemory'
    if ($null -ne $memory) { $ram = [string]('{0:N1}' -f ([double]$memory / 1GB)) }

    $image = Get-BenchmarkExcelImage -Identity $Identity
    $excelVersion = $unknown; $excelBuild = $unknown
    $excelBitness = [string]$image.Bitness
    $excelPath = $(if ([string]::IsNullOrWhiteSpace($image.Path)) { $unknown } else { [string]$image.Path })
    $calcMode = $unknown; $otherBooks = $unknown
    if ($null -ne $Excel) {
        # THESE FOUR ARE COM CALLS, NOT PROPERTY LOOKUPS ON A POWERSHELL OBJECT.
        # A COM member that an Excel build does not implement raises rather than
        # returning $null, and there is nothing to normalise first - so each is
        # asked on its own and a failure records 'unavailable' for that one field
        # and no other. Nothing here can swallow a failure of the benchmark
        # itself: the whole block runs before any operation is timed and cannot
        # affect a sample.
        try { $excelVersion = [string]$Excel.Version } catch { $excelVersion = $unknown }
        try { $excelBuild = [string]$Excel.Build } catch { $excelBuild = $unknown }
        try { $calcMode = [string]$Excel.Calculation } catch { $calcMode = $unknown }
        try { $otherBooks = [string]($Excel.Workbooks.Count) } catch { $otherBooks = $unknown }
    }

    $branch = $unknown
    try {
        $named = [string](& git -C $repoRoot rev-parse --abbrev-ref HEAD 2>$null)
        if (-not [string]::IsNullOrWhiteSpace($named)) { $branch = $named.Trim() }
    } catch {
        $branch = $unknown
    }

    $record = New-Object System.Collections.Specialized.OrderedDictionary
    $record.Add('captured_at_utc', ((Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')))
    $record.Add('host_name', [string]$env:COMPUTERNAME)
    $record.Add('powershell_version', [string]$PSVersionTable.PSVersion)
    $record.Add('windows_edition', (Format-BenchmarkFact (Get-BenchmarkProperty -InputObject $os -Name 'Caption')))
    $record.Add('windows_version', (Format-BenchmarkFact (Get-BenchmarkProperty -InputObject $os -Name 'Version')))
    $record.Add('windows_build', (Format-BenchmarkFact (Get-BenchmarkProperty -InputObject $os -Name 'BuildNumber')))
    $record.Add('cpu_model', (Format-BenchmarkFact (Get-BenchmarkProperty -InputObject $cpu -Name 'Name')))
    $record.Add('logical_processors',
                (Format-BenchmarkFact (Get-BenchmarkProperty -InputObject $cpu -Name 'NumberOfLogicalProcessors')))
    $record.Add('installed_ram_gb', $ram)
    $record.Add('excel_version', $excelVersion)
    $record.Add('excel_build', $excelBuild)
    $record.Add('excel_bitness', $excelBitness)
    $record.Add('excel_executable_path', $excelPath)
    $record.Add('excel_calculation_mode', $calcMode)
    $record.Add('other_workbooks_open', $otherBooks)
    $record.Add('workbook_path', $WorkbookPath)
    $record.Add('workbook_location_type',
                (Get-BenchmarkLocationType -Path $WorkbookPath -OneDriveRoots $roots))
    $record.Add('workbook_location_note',
                ('the timed workbook is a disposable Stage-B build made under -WorkDir; ' +
                 'the repository itself is recorded separately below and is NOT ' +
                 'assumed to be an unsynchronised local path'))
    $record.Add('repository_path', $RepositoryPath)
    $record.Add('repository_location_type',
                (Get-BenchmarkLocationType -Path $RepositoryPath -OneDriveRoots $roots))
    $record.Add('onedrive_roots', @($roots))
    $record.Add('git_branch', $branch)
    $record.Add('git_commit', [string]$Revision.Head)
    $record.Add('git_worktree_clean', ([bool](@($Revision.Dirty).Count -eq 0)))
    # W2 ROOT CAUSE, CORRECTED AT THE AUTHORITY RATHER THAN AT THE LINE.
    #
    # These three used to be read off `stage_b_manifest.json`, which carries
    # `model_version` and has NEVER carried the other two. It cannot: P10-3
    # settled that the model version and the builder version are INDEPENDENT
    # authorities, and that manifest is a projection of the MODEL side alone.
    # `$Manifest.builder_version` was therefore a property that never existed,
    # and under StrictMode 2.0 reading it is a terminating
    # PropertyNotFoundException - which is exactly where Windows Run 2 died.
    #
    # All three now come from `release_identity` in the generated benchmark
    # plan, where each is projected from its OWN owner and carries the name of
    # that owner beside it. The runner parses no Python and infers nothing.
    #
    # AND THEY ARE REQUIRED. A release identity is not an optional environment
    # fact: a timing nobody can attribute to a release cannot be compared with
    # anything later, so an absent value throws with the authority named instead
    # of being recorded as 'unavailable'.
    $record.Add('model_version', [string](Get-BenchmarkRequiredProperty `
        -InputObject $ReleaseIdentity -Name 'model_version' `
        -Where 'the benchmark plan release identity'))
    $record.Add('builder_version', [string](Get-BenchmarkRequiredProperty `
        -InputObject $ReleaseIdentity -Name 'builder_version' `
        -Where 'the benchmark plan release identity'))
    $record.Add('build_phase', [string](Get-BenchmarkRequiredProperty `
        -InputObject $ReleaseIdentity -Name 'build_phase' `
        -Where 'the benchmark plan release identity'))
    $record.Add('harness_version', $HarnessVersion)
    $record.Add('schema_version', $SchemaVersion)
    return $record
}

# ===========================================================================
# THE SCENARIO MODEL
# ===========================================================================
# THE SHAPE IS THE ACCEPTED PHASE-7 TIMING SHAPE, generalised over the number of
# project years the plan declares. Every field varies with the index and every
# driver has Min < Most Likely < Max, so no driver is degenerate: the
# zero-variance arm of the analysis is not what is being timed, and the ranked
# table is not mostly refusals.
#
# DETERMINISTIC, WITH NO RANDOMNESS ANYWHERE. The same scenario id and the same
# plan produce byte-identical inputs on every machine and every run.
function New-BenchmarkDriver {
    param([int]$Index, [bool]$IsRisk, [int]$Years)
    $families = @('Triangular', 'Beta-PERT', 'Uniform')
    $profiles = @('Standard', 'Escalated')
    $currencies = @('SAR', 'USD')

    $base = 100.0 + (7.0 * [double]$Index)
    $driver = [pscustomobject]@{
        permanent_id      = $(if ($IsRisk) { 'R-{0:D3}' -f $Index } else { 'CL-{0:D3}' -f $Index })
        distribution      = [string]$families[($Index - 1) % $families.Count]
        currency          = [string]$currencies[($Index - 1) % $currencies.Count]
        inflation_profile = [string]$profiles[($Index - 1) % $profiles.Count]
        min_value         = $base
        most_likely       = [double]($base * 1.35)
        max_value         = [double]($base * 2.10)
        profile_weights   = (New-BenchmarkWeights -Years $Years -Offset $Index)
    }
    if ($IsRisk) {
        # 0.1 .. 0.7, never 0 and never 1: a Risk that always or never occurs
        # carries no occurrence variance, and this measures the populated path.
        Add-Member -InputObject $driver -MemberType NoteProperty -Name 'probability' `
            -Value ([double](((($Index - 1) % 7) + 1) / 10.0))
    } else {
        Add-Member -InputObject $driver -MemberType NoteProperty -Name 'quantity' `
            -Value ([double](1 + (($Index - 1) % 5)))
    }
    return $driver
}

# THE WEIGHTS SUM TO EXACTLY ONE, and they are not flat.
#
# A flat profile would be cheap in a way real spend is not, so the shape is a
# deterministic ramp that varies per driver. The LAST year absorbs whatever the
# division left behind, so the total is exactly 1 in binary floating point
# rather than 0.9999999999999999 - which the accepted profiling rule would
# refuse, correctly, and this harness must never provoke.
function New-BenchmarkWeights {
    param([int]$Years, [int]$Offset)
    if ($Years -lt 1) { throw ('a scenario needs at least one project year; asked for ' + [string]$Years) }
    if ($Years -eq 1) { return ,@([double]1) }
    $raw = @()
    $total = 0.0
    for ($year = 1; $year -le $Years; $year++) {
        $shape = [double](1 + ((($year + $Offset) % 5)))
        $raw += $shape
        $total += $shape
    }
    $weights = @()
    $running = 0.0
    for ($index = 0; $index -lt ($Years - 1); $index++) {
        $value = [double]([Math]::Round([double]$raw[$index] / $total, 6))
        $weights += $value
        $running += $value
    }
    $weights += [double](1.0 - $running)
    return ,@($weights)
}

function New-BenchmarkModel {
    param($ScenarioSpec)
    $costCount = [int]$ScenarioSpec.cost_lines
    $riskCount = [int]$ScenarioSpec.risks
    $years = [int]$ScenarioSpec.years
    if (($costCount -lt 1) -or ($riskCount -lt 1)) {
        throw ('the plan gives ' + [string]$ScenarioSpec.id + ' a degenerate split: ' +
               [string]$costCount + ' Cost Lines and ' + [string]$riskCount + ' Risks')
    }

    $costLines = @()
    for ($index = 1; $index -le $costCount; $index++) {
        $costLines += (New-BenchmarkDriver -Index $index -IsRisk $false -Years $years)
    }
    $risks = @()
    for ($index = 1; $index -le $riskCount; $index++) {
        $risks += (New-BenchmarkDriver -Index $index -IsRisk $true -Years $years)
    }

    # base_year = start_year - 1 is the accepted shape: the generated inflation
    # columns begin at BaseYear + 1, so they are exactly the project years.
    $startYear = 2027
    $standard = New-Object System.Collections.Specialized.OrderedDictionary
    $escalated = New-Object System.Collections.Specialized.OrderedDictionary
    for ($offset = 0; $offset -lt $years; $offset++) {
        $calendar = [string]($startYear + $offset)
        $standard.Add($calendar, [double]0.03)
        $escalated.Add($calendar, [double](0.06 - (0.0005 * [double]$offset)))
    }

    return [pscustomobject]@{
        timeline      = [pscustomobject]@{ base_year = ($startYear - 1); start_year = $startYear
                                           duration = $years }
        discount_rate = 0.05
        fx            = @(
            [pscustomobject]@{ currency = 'SAR'; rate = 1.0 },
            [pscustomobject]@{ currency = 'USD'; rate = 3.75 }
        )
        inflation     = [pscustomobject]@{
            'Standard'  = [pscustomobject]$standard
            'Escalated' = [pscustomobject]$escalated
        }
        cost_lines    = $costLines
        risks         = $risks
    }
}

# THE SEED IS FIXED, and that is what makes two runs comparable. An AUTO run
# samples a different sequence every time; the same fixed seed does the same
# work, so a difference between two benchmark runs is a difference in the
# machine rather than in the sample.
function Get-BenchmarkSeed { return 20260101 }

# ===========================================================================
# THE EVIDENCE
# ===========================================================================
# WHAT PROVES AN EXECUTION ACTUALLY DID THE WORK. Read AFTER the clock stops,
# so no read is ever inside a measured interval.
# ONE INTEGER OUT OF PRODUCTION'S OWN ANNOUNCEMENT. Returns $null when the
# sentence does not carry one, which is itself evidence: a gate that needs the
# number then fails rather than silently passing on a default.
function Get-BenchmarkAnnouncedNumber {
    param([string]$Text, [string]$Pattern)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $match = [regex]::Match($Text, $Pattern)
    if (-not $match.Success) { return $null }
    return [int]$match.Groups[1].Value
}


function Get-BenchmarkEvidence {
    param($Excel, $Workbook, $SimInspection, [string]$OperationKey, [string]$AutomationResult)
    $evidence = New-Object System.Collections.Specialized.OrderedDictionary
    $evidence.Add('automation_result', $AutomationResult)

    if ($OperationKey -eq 'calculate' -or $OperationKey -eq 'recalculation') {
        try { $evidence.Add('calculation_status', [string]$Excel.Run('PCCM_CalculationStatus')) }
        catch { $evidence.Add('calculation_status', (Get-BenchmarkUnavailable)) }
        if ($OperationKey -eq 'calculate') {
            try { $evidence.Add('calculation_fingerprint', [string]$Excel.Run('PCCM_CalculationFingerprint')) }
            catch { $evidence.Add('calculation_fingerprint', (Get-BenchmarkUnavailable)) }
        }
        return $evidence
    }

    if ($OperationKey -eq 'annual') {
        foreach ($pair in @(
            @{ key = 'annual_distribution_state'; call = 'PCCM_AnnualDistributionState' },
            @{ key = 'annual_profile_state';      call = 'PCCM_AnnualProfileState' },
            @{ key = 'annual_profile_px';         call = 'PCCM_AnnualProfilePx' },
            @{ key = 'annual_year_count';         call = 'PCCM_AnnualYearCount' })) {
            try { $evidence.Add($pair.key, [string]$Excel.Run($pair.call)) }
            catch { $evidence.Add($pair.key, (Get-BenchmarkUnavailable)) }
        }
        # "Annual stochastic complete: Y project year(s) over K iterations."
        $evidence.Add('published_iterations', (Get-BenchmarkAnnouncedNumber `
            -Text $AutomationResult -Pattern 'over\s+(\d+)\s+iterations'))
        return $evidence
    }

    # simulation and sensitivity both answer through the published run identity.
    try { $evidence.Add('simulation_status', [string]$Excel.Run('PCCM_SimulationStatus')) }
    catch { $evidence.Add('simulation_status', (Get-BenchmarkUnavailable)) }

    $state = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
    $bank = Get-Phase6ActiveBank -State $state
    $evidence.Add('active_bank', $bank)
    if ([string]::IsNullOrEmpty($bank)) {
        $evidence.Add('published_iterations', (Get-BenchmarkUnavailable))
        $evidence.Add('run_id', (Get-BenchmarkUnavailable))
    } else {
        # THE NUMBER, NOT ITS RENDERING. The gate compares an integer with an
        # integer; a formatted string would compare "10000" with "10,000" on
        # exactly the locale nobody tested on.
        $block = $state[('bank_' + $bank)]
        $published = $block['iterations_run']
        if ($published -is [double]) {
            $evidence.Add('published_iterations', [int][double]$published)
        } else {
            $evidence.Add('published_iterations', $null)
        }
        $evidence.Add('run_id', (Format-SimValue $block['run_id']))
    }
    if ($OperationKey -eq 'sensitivity') {
        # The sensitivity announcement carries the iteration count it replayed
        # over, which is the number the gate must check - the bank block above
        # records what the SIMULATION published, and the two agreeing is the
        # point rather than the assumption.
        $evidence['published_iterations'] = (Get-BenchmarkAnnouncedNumber `
            -Text $AutomationResult -Pattern 'over\s+(\d+)\s+iterations')
        # PRODUCTION'S OWN SENTENCE, PARSED - not a cell read.
        #
        # PCCM_RunSensitivity announces "Sensitivity complete: N ranked of M
        # drivers, over K iterations." Everything this harness needs to know that
        # the ranking really happened is in that sentence, so nothing here needs
        # a _SimData address. A harness that read the sensitivity block directly
        # would be a second declaration of geometry that no build artifact
        # currently projects, and the accepted Phase-7 harness already had to
        # correct exactly that kind of typed copy once.
        $evidence.Add('sensitivity_ranked', (Get-BenchmarkAnnouncedNumber -Text $AutomationResult `
            -Pattern '(\d+)\s+ranked\s+of'))
        $evidence.Add('sensitivity_record_count', (Get-BenchmarkAnnouncedNumber -Text $AutomationResult `
            -Pattern 'ranked\s+of\s+(\d+)\s+drivers'))
    }
    return $evidence
}

# THE GATES. A sample that fails one is not a slow sample or a fast sample - it
# is not a sample. It is excluded from every statistic and reported with its
# reason.
# THE SHAPE IS CHECKED, NOT TRUSTED.
#
# The pairing above is a two-place contract - a comma in the helper and no `@()`
# at the call site - and a contract spread over two places is one an edit can
# half-keep. This refuses the half-kept version AT THE CALL, by name, instead of
# letting it reappear as every sample being invalid for an unreadable reason.
#
# IT THROWS. A malformed problem list is a defect in this harness, not a fact
# about the workbook, so it must not be recorded as an invalid sample - that
# would be the harness marking its own bug as the model's. It also does not
# flatten: flattening would hide the defect and keep the run going.
function Assert-BenchmarkProblemList {
    param($Problems, [string]$Where)
    if ($null -eq $Problems) {
        throw ('the sample validator returned nothing for ' + $Where +
               '; it must return a list, empty when the sample is good')
    }
    foreach ($problem in $Problems) {
        if ($problem -is [System.Array]) {
            throw ('the sample validator returned a NESTED list for ' + $Where +
                   ': an element is a ' + $problem.GetType().FullName + '. A caller ' +
                   'wrapped `return ,@(...)` in `@()` again, which makes every sample ' +
                   'invalid for an unreadable reason. Assign the result; do not re-wrap it.')
        }
        if ($problem -isnot [string]) {
            throw ('the sample validator returned a ' + $problem.GetType().FullName +
                   ' for ' + $Where + ' where every problem must be a string')
        }
    }
}

function Test-BenchmarkSample {
    param($Operation, $Evidence, $RequestedIterations)
    $problems = @()
    $result = [string]$Evidence['automation_result']
    if ([string]$Operation.kind -eq 'command') {
        if ($result -notlike 'OK|*') {
            $problems += ('the command did not announce success: ' + $result)
        }
    }
    if ([bool]$Operation.iteration_dependent -and $null -ne $RequestedIterations) {
        # THE ITERATION COUNT IS VERIFIED, NEVER ASSUMED. A run that quietly
        # executed ten thousand iterations when a hundred thousand were asked
        # for would be the fastest and most useless measurement in the report.
        $published = $Evidence['published_iterations']
        if ($null -eq $published) {
            $problems += 'the operation did not report how many iterations it executed'
        } elseif (([int]$published) -ne ([int]$RequestedIterations)) {
            $problems += ('the workbook executed ' + [string]$published + ' iterations where ' +
                          [string][int]$RequestedIterations + ' were requested')
        }
    }
    if ([string]$Operation.key -eq 'annual') {
        $distribution = [string]$Evidence['annual_distribution_state']
        if ($distribution -ne 'CURRENT') {
            $problems += ('the annual distribution state is ' + $distribution + ', not CURRENT')
        }
    }
    return ,@($problems)
}

# ===========================================================================
# ONE TIMED EXECUTION
# ===========================================================================
# NOTHING BETWEEN THE TWO STOPWATCH STATEMENTS except the one call. The
# automation envelope is opened before the clock starts and the announcement and
# every evidence read happen after it stops, so neither is inside the interval.
function Invoke-BenchmarkExecution {
    param($Excel, $Workbook, $SimInspection, $Operation, $RequestedIterations, [string]$Phase)

    $result = ''
    if ([string]$Operation.kind -eq 'recalculation') {
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        $Excel.CalculateFull()
        $watch.Stop()
        $result = 'OK|recalculation'
    } else {
        $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        $Excel.Run([string]$Operation.endpoint) | Out-Null
        $watch.Stop()
        $result = [string]$Excel.Run('PCCM_AutomationResult')
    }

    $evidence = Get-BenchmarkEvidence -Excel $Excel -Workbook $Workbook `
        -SimInspection $SimInspection -OperationKey ([string]$Operation.key) `
        -AutomationResult $result
    # ASSIGNED, NOT RE-WRAPPED.
    #
    # `Test-BenchmarkSample` ends `return ,@($problems)`. The leading comma hands
    # back ONE object that IS the array, which is what keeps an EMPTY result from
    # being enumerated into nothing by the pipeline. Wrapping that one object in
    # `@()` again collects it into a NEW one-element array whose single element
    # is the real list.
    #
    # `@()` COLLECTS PIPELINE ITEMS; IT DOES NOT FLATTEN A NESTED ARRAY. So the
    # count was 1 whatever the sample actually found: every execution was marked
    # INVALID, and `-join` rendered the inner array as its type name. That is the
    # whole of `INVALID: System.Object[]` - Calculate and the recalculation had
    # executed correctly and had no problems at all.
    $problems = Test-BenchmarkSample -Operation $Operation -Evidence $evidence `
        -RequestedIterations $RequestedIterations
    Assert-BenchmarkProblemList -Problems $problems -Where ([string]$Operation.key)

    return [pscustomobject]@{
        Phase      = $Phase
        ElapsedMs  = [double]$watch.Elapsed.TotalMilliseconds
        Valid      = ([bool]($problems.Count -eq 0))
        Problems   = $problems
        Evidence   = $evidence
    }
}

# ===========================================================================
# PREFLIGHT: THE PLAN, THE AUTHORITIES, AND THE SOURCE REVISION
# ===========================================================================
Write-Host ''
Write-Host 'PCCM - Phase 10 delivery performance benchmark' -ForegroundColor Cyan
Write-Host '=============================================' -ForegroundColor Cyan
Write-Host ''
Write-Host 'This run RECORDS a baseline. It does not judge one: there is no' -ForegroundColor Yellow
Write-Host 'absolute pass mark before the target machine has been observed.' -ForegroundColor Yellow
Write-Host ''

Set-BenchmarkStage -Stage 'preflight' -Action 'reading the benchmark plan and the contract projections'
$planPath       = Join-Path $BuildDir 'phase10_benchmark_plan.json'
$manifestPath   = Join-Path $BuildDir 'stage_b_manifest.json'
$inspectPath    = Join-Path $BuildDir 'phase5_gate_b_inspection.json'
$simInspectPath = Join-Path $BuildDir 'phase6_gate_b_inspection.json'
foreach ($required in @($planPath, $manifestPath, $inspectPath, $simInspectPath)) {
    if (-not (Test-Path -LiteralPath $required)) {
        Write-Host ("$required not found. Run the Stage-A build first: " +
                    'python3 pccm/builder/build_stage_a.py') -ForegroundColor Red
        exit 1
    }
}
$plan          = Get-Content -LiteralPath $planPath       -Raw | ConvertFrom-Json
$manifest      = Get-Content -LiteralPath $manifestPath   -Raw | ConvertFrom-Json
$inspection    = Get-Content -LiteralPath $inspectPath    -Raw | ConvertFrom-Json
$simInspection = Get-Content -LiteralPath $simInspectPath -Raw | ConvertFrom-Json

if ([int]$plan.schema_version -ne $script:BenchmarkSupportedSchema) {
    Write-Host ('the benchmark plan is schema ' + [string]$plan.schema_version +
                ' and this runner understands schema ' +
                [string]$script:BenchmarkSupportedSchema + '. Nothing was measured.') -ForegroundColor Red
    exit 1
}

# THE RELEASE IDENTITY, CHECKED BEFORE ANYTHING IS SPENT.
#
# W2 DIED SIXTY-EIGHT SECONDS INTO A RUN because a release value was missing and
# nothing had looked for it until the environment capture. It is checked here,
# before the bootstrap and before Excel is started, and a plan that cannot
# identify the release it measures is REFUSED rather than measured.
#
# NO SUBSTITUTE IS ACCEPTED. Not an empty string, not 'unknown', not a default
# 1.0.0: a timing nobody can attribute to a release is not evidence, and the
# refusal names the authority the projection was supposed to come from.
$releaseIdentity = $null
foreach ($candidate in @($plan.release_identity)) { $releaseIdentity = $candidate }
if ($null -eq $releaseIdentity) {
    Write-Host ('the benchmark plan carries no release_identity block. Rebuild it: ' +
                'python3 pccm/builder/build_stage_a.py') -ForegroundColor Red
    exit 1
}
$releaseProblems = @()
foreach ($field in @('model_version', 'builder_version', 'build_phase')) {
    $value = [string](Get-BenchmarkProperty -InputObject $releaseIdentity -Name $field)
    if ([string]::IsNullOrWhiteSpace($value)) {
        $authority = [string](Get-BenchmarkProperty `
            -InputObject (Get-BenchmarkProperty -InputObject $releaseIdentity -Name 'authorities') `
            -Name $field)
        if ([string]::IsNullOrWhiteSpace($authority)) { $authority = 'its declared owner' }
        $releaseProblems += ('  ' + $field + ' is missing from the plan projection; its ' +
                             'authority is ' + $authority)
    }
}
if ($releaseProblems.Count -gt 0) {
    Write-Host 'REFUSED, BEFORE ANYTHING WAS BUILT OR MEASURED.' -ForegroundColor Red
    Write-Host ''
    Write-Host ('The benchmark plan does not identify the release it would measure, and a ' +
                'timing nobody can attribute to a release is not evidence:') -ForegroundColor Red
    foreach ($line in $releaseProblems) { Write-Host $line -ForegroundColor Red }
    Write-Host ''
    Write-Host 'Rebuild the plan: python3 pccm/builder/build_stage_a.py' -ForegroundColor Red
    exit 1
}

# THE SCENARIO IS THE PLAN'S, NOT THIS FILE'S. Sizes, split, years and iteration
# counts are read; none of them is declared here.
$scenarioSpec = $null
foreach ($candidate in @($plan.scenarios)) {
    if ([string]$candidate.id -eq $Scenario) { $scenarioSpec = $candidate }
}
if ($null -eq $scenarioSpec) {
    Write-Host ('the plan declares no scenario ' + $Scenario) -ForegroundColor Red
    exit 1
}

$declaredIterations = @($scenarioSpec.iterations | ForEach-Object { [int]$_ })
$requestedIterations = $declaredIterations
if ($null -ne $Iterations -and @($Iterations).Count -gt 0) {
    # A COUNT THE PLAN DOES NOT DECLARE FOR THIS SCENARIO IS REFUSED. This is
    # what puts the Large x 100,000 cap out of reach from the command line: it
    # is not merely absent from the matrix, it cannot be asked for.
    $unknown = @($Iterations | Where-Object { $declaredIterations -notcontains [int]$_ })
    if ($unknown.Count -gt 0) {
        Write-Host ('REFUSED: ' + $Scenario + ' declares iteration counts ' +
                    ($declaredIterations -join ', ') + '. Asked for ' +
                    ($unknown -join ', ') + '.') -ForegroundColor Red
        foreach ($rule in @($plan.forbidden)) {
            if ([string]$rule.scenario -eq $Scenario) {
                Write-Host ('  ' + [string]$rule.iterations + ': ' + [string]$rule.reason) -ForegroundColor Red
            }
        }
        exit 1
    }
    $requestedIterations = @($Iterations | ForEach-Object { [int]$_ })
}

$operationByKey = @{}
foreach ($operation in @($plan.operations)) { $operationByKey[[string]$operation.key] = $operation }

$plannedRuns = @()
foreach ($run in @($plan.runs)) {
    if ([string]$run.scenario -ne $Scenario) { continue }
    if ($null -ne $Operations -and @($Operations).Count -gt 0) {
        if ($Operations -notcontains [string]$run.operation) { continue }
    }
    if ($null -ne $run.iterations) {
        if ($requestedIterations -notcontains [int]$run.iterations) { continue }
    }
    $plannedRuns += $run
}
if ($plannedRuns.Count -eq 0) {
    Write-Host 'nothing to measure with those filters' -ForegroundColor Red
    exit 1
}

# FAIL CLOSED ON AN UNATTRIBUTABLE WORKBOOK. A measurement nobody can attribute
# to a revision is not evidence. The accepted Phase-7 check, unchanged.
Set-BenchmarkStage -Stage 'preflight' -Action 'attributing the workbook to a source revision'
$revision = $null
try {
    $revision = Get-BenchmarkSourceRevision -RepoRoot $repoRoot
} catch {
    Write-Host (Format-Err $_) -ForegroundColor Red
    exit 1
}
if (@($revision.Dirty).Count -gt 0) {
    Write-Host 'REFUSED, BEFORE EXCEL WAS STARTED.' -ForegroundColor Red
    Write-Host ''
    Write-Host ('pccm/src, pccm/spec or pccm/builder is modified, so the workbook this ' +
                'run would build cannot be attributed to a source revision:') -ForegroundColor Red
    foreach ($line in @($revision.Dirty)) { Write-Host ('    ' + $line) -ForegroundColor Red }
    Write-Host ''
    Write-Host 'Commit or stash those changes and run again.' -ForegroundColor Red
    exit 1
}

# ===========================================================================
# A DISPOSABLE COPY OF THE BUILD, AND THE STAGE-B BOOTSTRAP
# ===========================================================================
# The real build output is never opened, never mutated and never saved over.
# THIS IS SETUP. Its cost is timed and reported, and it is part of no
# measurement.
if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = [System.IO.Path]::GetTempPath() }
$stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
$tempRoot = Join-Path $WorkDir ('pccm-phase10-benchmark-' + $Scenario.ToLower() + '-' + $stamp)
$null = New-Item -ItemType Directory -Path $tempRoot -Force
if ([string]::IsNullOrWhiteSpace($OutDir)) { $OutDir = $tempRoot }
$null = New-Item -ItemType Directory -Path $OutDir -Force -ErrorAction SilentlyContinue

Set-BenchmarkStage -Stage 'setup' -Action 'copying the build and running the Stage-B bootstrap'
$bootstrapWatch = [System.Diagnostics.Stopwatch]::StartNew()
Copy-Item -LiteralPath (Join-Path $BuildDir ([string]$manifest.stage_a_filename)) -Destination $tempRoot
Copy-Item -LiteralPath $manifestPath   -Destination $tempRoot
Copy-Item -LiteralPath $inspectPath    -Destination $tempRoot
Copy-Item -LiteralPath $simInspectPath -Destination $tempRoot
Copy-Item -LiteralPath (Join-Path $BuildDir 'vba') -Destination $tempRoot -Recurse

$script:BenchmarkTextPath = Join-Path $OutDir ('phase10_benchmark_' + $Scenario + '_' + $stamp + '.log')
$stageBPath = Join-Path $tempRoot ([string]$manifest.stage_b_filename)

$bootstrap = Join-Path $scriptDir 'build_stage_b.ps1'
& $bootstrap -BuildDir $tempRoot -Force
$bootstrapExit = $LASTEXITCODE
$bootstrapWatch.Stop()
if (($bootstrapExit -ne 0) -or (-not (Test-Path -LiteralPath $stageBPath))) {
    Write-Host ''
    Write-Host ('The Stage-B bootstrap did not complete (exit ' + [string]$bootstrapExit +
                '). Nothing was measured.') -ForegroundColor Red
    exit 1
}

# The last moment the executed .xlsm is both built and unlocked.
$artefacts = Get-Phase6RuntimeArtefactIdentity -TempRoot $tempRoot -Manifest $manifest

Write-BenchmarkLine 'PCCM - PHASE 10 DELIVERY PERFORMANCE BENCHMARK'
Write-BenchmarkLine '============================================='
Write-BenchmarkLine ''
Write-BenchmarkLine 'THIS IS A BASELINE RECORDING, NOT AN ACCEPTANCE RUN.'
Write-BenchmarkLine 'No Gate-B scenario ran. No pass or fail is decided here.'
Write-BenchmarkLine ''
Write-BenchmarkLine ('baseline id            : ' + [string]$plan.baseline_id)
Write-BenchmarkLine ('plan schema            : ' + [string]$plan.schema_version)
Write-BenchmarkLine ('harness version        : ' + [string]$plan.harness_version)
Write-BenchmarkLine 'release identity, each value from its own authority:'
foreach ($field in @('model_version', 'builder_version', 'build_phase')) {
    Write-BenchmarkLine ('  ' + $field.PadRight(20) + ' : ' +
                         [string](Get-BenchmarkProperty -InputObject $releaseIdentity -Name $field) +
                         '   [' + [string](Get-BenchmarkProperty `
                             -InputObject (Get-BenchmarkProperty -InputObject $releaseIdentity `
                                 -Name 'authorities') -Name $field) + ']')
}
Write-BenchmarkLine ('scenario               : ' + $Scenario + ' - ' + [string]$scenarioSpec.title)
Write-BenchmarkLine ('  drivers              : ' + [string]$scenarioSpec.drivers +
                     ' (' + [string]$scenarioSpec.cost_lines + ' Cost Lines + ' +
                     [string]$scenarioSpec.risks + ' Risks)')
Write-BenchmarkLine ('  project years        : ' + [string]$scenarioSpec.years)
Write-BenchmarkLine ('  iterations declared  : ' + ($declaredIterations -join ', '))
Write-BenchmarkLine ('  iterations this run  : ' + ($requestedIterations -join ', '))
Write-BenchmarkLine ('  timed runs planned   : ' + [string]$plannedRuns.Count)
Write-BenchmarkLine ''
Write-BenchmarkLine 'COMBINATIONS THE PLAN FORBIDS'
foreach ($rule in @($plan.forbidden)) {
    Write-BenchmarkLine ('  ' + [string]$rule.scenario + ' x ' + [string]$rule.iterations +
                         ' for ' + ((@($rule.operations)) -join ', ') + ': ' + [string]$rule.reason)
}
Write-BenchmarkLine ''
Write-BenchmarkLine 'TIMING METHOD'
Write-BenchmarkLine ('  cold runs            : ' + [string]$plan.timing.cold_runs +
                     '  (' + [string]$plan.timing.cold_definition + ')')
Write-BenchmarkLine ('  warm runs            : ' + [string]$plan.timing.warm_runs +
                     '  (' + [string]$plan.timing.warm_definition + ')')
Write-BenchmarkLine ('  compared on          : ' + [string]$plan.timing.comparison_statistic)
Write-BenchmarkLine ('  clock                : ' + [string]$plan.timing.timing_source)
Write-BenchmarkLine ('  NOT in elapsed time  : ' + ((@($plan.timing.excluded_from_elapsed)) -join ', '))
Write-BenchmarkLine ''
Write-BenchmarkLine 'BASELINE AND REGRESSION POLICY'
Write-BenchmarkLine ('  absolute threshold   : ' + [string]$plan.regression_policy.absolute_threshold_note)
foreach ($rule in @($plan.regression_policy.rules)) { Write-BenchmarkLine ('  ' + [string]$rule) }
Write-BenchmarkLine ''
Write-BenchmarkLine 'HISTORICAL PHASE-7 TIMINGS'
Write-BenchmarkLine ('  ' + [string]$plan.historical_context.status)
foreach ($observation in @($plan.historical_context.observations)) {
    Write-BenchmarkLine ('    ' + [string]$observation.operation + '  ' +
                         [string]$observation.drivers + ' drivers x ' +
                         [string]$observation.iterations + ' iterations  ~' +
                         [string]$observation.seconds + ' s')
}
Write-BenchmarkLine ''
Write-BenchmarkLine 'SOURCE REVISION AND BUILD IDENTITY'
Write-BenchmarkLine ('  git HEAD             : ' + [string]$revision.Head)
Write-BenchmarkLine  '  pccm/src, pccm/spec, pccm/builder : clean (proved before Excel was started)'
foreach ($item in $artefacts) {
    $shown = $item.Hash
    if ([string]::IsNullOrWhiteSpace($shown)) { $shown = '(' + $item.Problem + ')' }
    Write-BenchmarkLine ('  ' + $item.Label.PadRight(30) + ' SHA-256 ' + $shown)
}
Write-BenchmarkLine ''
Write-BenchmarkLine ('  Stage-B bootstrap    : ' +
                     (Format-BenchmarkSeconds $bootstrapWatch.Elapsed.TotalMilliseconds) +
                     '  [SETUP, part of no measurement]')
Write-BenchmarkLine ''

# ===========================================================================
# THE MEASUREMENT SESSION
# ===========================================================================
$preExisting = @(Get-PreExistingExcelPids)
$excel = $null; $workbooks = $null; $wb = $null
$excelIdentity = $null
$rel = $null
$environment = $null
$results = New-Object System.Collections.ArrayList
# DECLARED BEFORE THE TRY, because the artifact is written whatever happened and
# StrictMode will not read an unassigned variable. An abandoned run still emits a
# report; it emits one that says what it managed to establish.
$actual = New-Object System.Collections.Specialized.OrderedDictionary
$setupTimings = New-Object System.Collections.Specialized.OrderedDictionary
$setupTimings.Add('stage_b_bootstrap_ms', [double]$bootstrapWatch.Elapsed.TotalMilliseconds)
$sessionWatch = [System.Diagnostics.Stopwatch]::StartNew()
$abandoned = ''
$failure = $null

try {
    Set-BenchmarkStage -Stage 'setup' -Action 'starting an owned Excel instance'
    $startupWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $excel = New-Object -ComObject Excel.Application
    $excelIdentity = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $startupWatch.Stop()

    Set-BenchmarkStage -Stage 'setup' -Action 'opening the benchmark workbook'
    $openWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $workbooks = $excel.Workbooks
    $wb = $workbooks.Open($stageBPath)
    $openWatch.Stop()

    # EXCEL STARTUP AND WORKBOOK OPEN ARE RECORDED AND EXCLUDED. They are here so
    # a reader can see they were paid, and nowhere near an operation's elapsed
    # time. "Cold" never means either of them.
    $setupTimings.Add('excel_startup_ms', [double]$startupWatch.Elapsed.TotalMilliseconds)
    $setupTimings.Add('workbook_open_ms', [double]$openWatch.Elapsed.TotalMilliseconds)

    Set-BenchmarkStage -Stage 'setup' -Action 'capturing the environment inventory'
    $environment = Get-BenchmarkEnvironment -Excel $excel -Identity $excelIdentity `
        -WorkbookPath $stageBPath -RepositoryPath $repoRoot -ReleaseIdentity $releaseIdentity `
        -HarnessVersion ([string]$plan.harness_version) `
        -SchemaVersion ([int]$plan.schema_version) -Revision $revision

    Write-BenchmarkLine 'ENVIRONMENT'
    Write-BenchmarkLine '-----------'
    foreach ($field in @($plan.environment_fields)) {
        $value = $environment[[string]$field]
        if ($value -is [System.Array]) { $value = (@($value) -join '; ') }
        if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) {
            $value = '(none)'
        }
        Write-BenchmarkLine ('  ' + ([string]$field).PadRight(26) + ' : ' + [string]$value)
    }
    Write-BenchmarkLine ''

    $costRegister = $null; $riskRegister = $null
    foreach ($register in @($manifest.registers)) {
        if ([string]$register.key -eq 'cost_lines')    { $costRegister = $register }
        if ([string]$register.key -eq 'risk_register') { $riskRegister = $register }
    }

    Set-BenchmarkStage -Stage 'setup' -Action 'opening the automation envelope and capturing the FX seed'
    $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
    $null = Save-Phase5LockedFxSeed -Workbook $wb -Inspection $inspection

    Set-BenchmarkStage -Stage 'setup' -Action 'importing the fixture-window shim into the disposable project'
    Import-BenchmarkFixtureWindow -Excel $excel -Workbook $wb -Manifest $manifest -ScriptDir $scriptDir
    $protectionBefore = Assert-BenchmarkProtectionApplied -Excel $excel -Manifest $manifest `
        -Stage 'as the workbook was opened'
    Write-BenchmarkLine 'PROTECTION'
    Write-BenchmarkLine '----------'
    Write-BenchmarkLine ('  as opened            : ' + $protectionBefore.Raw)
    Write-BenchmarkLine ''

    # --- THE SCENARIO, THROUGH THE ACCEPTED FIXTURE ------------------------
    # SETUP. Timed, reported, and part of no measurement. At three hundred
    # drivers over forty project years this is the expensive part of the run,
    # and it is expensive exactly once.
    Write-BenchmarkLine ('BUILDING ' + $Scenario)
    Write-BenchmarkLine ('-' * (9 + $Scenario.Length))
    Set-BenchmarkStage -Stage 'scenario' -Action ('building the ' + $Scenario + ' fixture through the accepted production endpoints')
    Set-BenchmarkOperationContext -Scenario $Scenario
    $model = New-BenchmarkModel -ScenarioSpec $scenarioSpec
    $fixtureWatch = [System.Diagnostics.Stopwatch]::StartNew()
    # THE WINDOW SPANS THE FIXTURE AND NOTHING ELSE.
    #
    # `finally` rather than a straight line, because the one outcome that must be
    # impossible is a fixture that raised and left the workbook unprotected. If
    # the fixture throws, the window still closes; if the close then fails too,
    # its failure is the one that reaches the abandon path, because an
    # unprotected workbook is the worse fact.
    $null = Open-BenchmarkFixtureWindow -Excel $excel -Manifest $manifest
    try {
        # WHICH BUILDER, RECORDED IN THE ARTIFACT. Two fixture methods that reach
        # the same state are still two methods, and a baseline must never be
        # compared across them without someone seeing that it was.
        if ($FixtureMode -eq 'Bulk') {
            $null = Set-BenchmarkBulkFixture -Excel $excel -Workbook $wb -Manifest $manifest `
                -Inspection $inspection -Model $model -ScenarioSpec $scenarioSpec
        } else {
            $null = Set-Phase5Fixture -Excel $excel -Workbook $wb -Manifest $manifest `
                -Inspection $inspection -Model $model
        }
    } finally {
        $protectionAfter = Close-BenchmarkFixtureWindow -Excel $excel -Manifest $manifest
    }
    # OUTSIDE THE WINDOW ON PURPOSE. This is a VALUE write to a named cell, which
    # Benchmark Run 3 proved an external COM caller can make on a protected
    # sheet, so it does not need the window and does not get it.
    Set-NamedValue -Workbook $wb `
        -DefinedName ([string]$simInspection.controls.random_seed.defined_name) `
        -Value ([double](Get-BenchmarkSeed))
    $fixtureWatch.Stop()
    $setupTimings.Add('scenario_fixture_ms', [double]$fixtureWatch.Elapsed.TotalMilliseconds)
    Write-BenchmarkLine ('  after the fixture    : ' + $protectionAfter.Raw +
                         '   [window closed and verified before any timed run]')

    # THE DIMENSIONS ARE READ BACK OUT OF THE WORKBOOK, not taken from the model.
    # What is timed is what the workbook holds.
    Set-BenchmarkStage -Stage 'scenario' -Action 'reading the built dimensions back out of the workbook'
    $costIds = @(Get-IdColumnValues -Workbook $wb -Info $costRegister)
    $riskIds = @(Get-IdColumnValues -Workbook $wb -Info $riskRegister)
    $actual.Add('cost_lines', [int]$costIds.Count)
    $actual.Add('risks', [int]$riskIds.Count)
    $actual.Add('drivers', [int]($costIds.Count + $riskIds.Count))
    $actual.Add('years', [int]$scenarioSpec.years)
    $actual.Add('seed_mode', 'FIXED')
    $actual.Add('seed', [int](Get-BenchmarkSeed))
    Write-BenchmarkLine ('  Cost Lines in book   : ' + [string]$actual['cost_lines'] +
                         ' (plan: ' + [string]$scenarioSpec.cost_lines + ')')
    Write-BenchmarkLine ('  Risks in book        : ' + [string]$actual['risks'] +
                         ' (plan: ' + [string]$scenarioSpec.risks + ')')
    Write-BenchmarkLine ('  project years        : ' + [string]$actual['years'])
    Write-BenchmarkLine ('  seed                 : FIXED ' + [string](Get-BenchmarkSeed) +
                         '  (the same work every run, so a difference is the machine)')
    Write-BenchmarkLine ('  fixture method       : ' + [string]$FixtureMode +
                         $(if ($FixtureMode -eq 'Bulk') {
                             '  [bulk inputs + ONE real PCCM_ApplyTimeline; every result still from production]'
                           } else { '  [one production Add per driver]' }))
    Write-BenchmarkLine ('  fixture build time   : ' +
                         (Format-BenchmarkSeconds $fixtureWatch.Elapsed.TotalMilliseconds) +
                         '  [SETUP, part of no measurement]')
    Write-BenchmarkLine ''

    if (([int]$actual['drivers']) -ne ([int]$scenarioSpec.drivers)) {
        $abandoned = ('the workbook holds ' + [string]$actual['drivers'] +
                      ' drivers where the plan declares ' + [string]$scenarioSpec.drivers)
        throw $abandoned
    }

    # --- THE RUNS -----------------------------------------------------------
    $currentIterations = $null
    foreach ($run in $plannedRuns) {
        $operation = $operationByKey[[string]$run.operation]
        # NAMED SO IT CANNOT BE THE PARAMETER.
        #
        # This used to be `$iterations`, and PowerShell variable names are
        # CASE-INSENSITIVE: that IS the script parameter `[int[]]$Iterations`. A
        # typed parameter keeps its type constraint for the whole life of the
        # variable, so every assignment to it is coerced back to `[int[]]` -
        # `$iterations = [int]$run.iterations` stored the one-element array
        # `@(10000)`, not the integer.
        #
        # `[string]` OF A ONE-ELEMENT ARRAY IS THE ELEMENT, which is why the
        # banner read "Run Simulation @ 10000 iterations" and looked right, and
        # `[double]` of the same array raised
        # "Cannot convert the System.Int32[] value ... to type System.Double" at
        # the iteration-control write.
        #
        # The plan carries a SCALAR here - `runs[].iterations` is 10000, not
        # [10000] - so nothing needed selecting out of an array. The shape defect
        # was the name.
        $runIterations = $null
        if ($null -ne $run.iterations) { $runIterations = [int]$run.iterations }

        $label = [string]$operation.label
        if ($null -ne $runIterations) { $label = $label + '  @ ' + [string]$runIterations + ' iterations' }
        Write-BenchmarkLine ($label)
        Write-BenchmarkLine ('-' * $label.Length)

        if (($null -ne $runIterations) -and ($runIterations -ne $currentIterations)) {
            # SETTING THE CONTROL IS SETUP. It is outside every clock, and the
            # value is read back so the run records what the workbook was asked
            # for rather than what this script intended.
            Set-BenchmarkStage -Stage 'measurement' -Action ('setting the iteration control to ' + [string]$runIterations + ' and re-establishing the deterministic basis')
            Set-NamedValue -Workbook $wb `
                -DefinedName ([string]$simInspection.controls.monte_carlo_iterations.defined_name) `
                -Value ([double]$runIterations)
            $null = Invoke-Phase5ProductionOperation -Excel $excel -Operation 'PCCM_Calculate' `
                -Stage ('re-establishing the deterministic basis for ' + [string]$runIterations +
                        ' iterations')
            $currentIterations = $runIterations
        }
        $iterationsSet = [string](Get-NamedValue -Workbook $wb `
            -DefinedName ([string]$simInspection.controls.monte_carlo_iterations.defined_name))

        $samples = New-Object System.Collections.ArrayList
        $total = [int]$run.cold_runs + [int]$run.warm_runs
        for ($index = 0; $index -lt $total; $index++) {
            $phase = $(if ($index -lt [int]$run.cold_runs) { 'cold' } else { 'warm' })
            # THE CURSOR IS SET HERE, OUTSIDE THE CLOCK, and read only if
            # something throws. `Invoke-BenchmarkExecution` starts its stopwatch
            # after this returns, so no diagnostic work is ever inside a
            # measured interval.
            Set-BenchmarkStage -Stage 'measurement' -Action ('executing ' + [string]$operation.label)
            Set-BenchmarkOperationContext -Scenario $Scenario -Operation ([string]$operation.key) `
                -Iterations $runIterations -Phase $phase
            $execution = Invoke-BenchmarkExecution -Excel $excel -Workbook $wb `
                -SimInspection $simInspection -Operation $operation `
                -RequestedIterations $runIterations -Phase $phase
            $null = $samples.Add($execution)
            $marker = $(if ($execution.Valid) { '' } else { '   INVALID: ' + (($execution.Problems) -join '; ') })
            Write-BenchmarkLine ('  ' + $phase.PadRight(5) +
                                 ' ' + (Format-BenchmarkSeconds $execution.ElapsedMs) + $marker)
        }

        $warm = @($samples | Where-Object { $_.Phase -eq 'warm' })
        $validWarm = @($warm | Where-Object { $_.Valid })
        $cold = @($samples | Where-Object { $_.Phase -eq 'cold' })[0]

        # THE MEDIAN EXISTS ONLY IF EVERY WARM SAMPLE WAS VALID. A median of two
        # good runs and one refusal is not a measurement of anything.
        $median = $null
        if ($validWarm.Count -eq [int]$run.warm_runs) {
            $median = Get-BenchmarkMedian -Values @($validWarm | ForEach-Object { [double]$_.ElapsedMs })
        }
        $valid = ([bool](($cold.Valid) -and ($validWarm.Count -eq [int]$run.warm_runs)))

        if ($null -ne $median) {
            Write-BenchmarkLine ('  WARM MEDIAN : ' + (Format-BenchmarkSeconds $median))
        } else {
            Write-BenchmarkLine '  WARM MEDIAN : NOT COMPUTED - this scenario did not produce three valid warm samples'
        }
        Write-BenchmarkLine ('  iterations control   : ' + $iterationsSet)
        Write-BenchmarkLine ''

        $row = New-Object System.Collections.Specialized.OrderedDictionary
        $row.Add('scenario', $Scenario)
        $row.Add('operation', [string]$operation.key)
        $row.Add('operation_label', [string]$operation.label)
        $row.Add('kind', [string]$operation.kind)
        $row.Add('endpoint', [string]$operation.endpoint)
        $row.Add('iterations_requested', $runIterations)
        $row.Add('iterations_control', $iterationsSet)
        $row.Add('valid', $valid)
        $row.Add('cold_ms', [double]$cold.ElapsedMs)
        $row.Add('warm_ms', @($warm | ForEach-Object { [double]$_.ElapsedMs }))
        $row.Add('warm_median_ms', $median)
        $row.Add('samples', @($samples | ForEach-Object {
            $sample = New-Object System.Collections.Specialized.OrderedDictionary
            $sample.Add('phase', [string]$_.Phase)
            $sample.Add('elapsed_ms', [double]$_.ElapsedMs)
            $sample.Add('valid', [bool]$_.Valid)
            $sample.Add('problems', @($_.Problems))
            $sample.Add('evidence', $_.Evidence)
            $sample
        }))
        $null = $results.Add($row)
    }

    $excel.Run('PCCM_AutomationEnd') | Out-Null
} catch {
    # WHAT FAILED, WHERE, AND WHAT THE RUN WAS DOING. Windows Run 1 reached this
    # handler and reported one sentence; it now reports the cursor, the
    # exception type, the failing statement and its line.
    $failure = New-BenchmarkFailureRecord -ErrorRecord $_ -Fallback $abandoned
    $abandoned = [string]$failure['message']
    Write-BenchmarkLine ''
    Write-BenchmarkLine 'THE BENCHMARK SESSION RAISED'
    Write-BenchmarkLine '----------------------------'
    Write-BenchmarkLine (Format-BenchmarkFailure $failure)
    Write-BenchmarkLine ''
    Write-BenchmarkLine 'Any operation that produced three valid warm samples before this'
    Write-BenchmarkLine 'point is above and is still a valid measurement. THIS RUN IS NOT A'
    Write-BenchmarkLine 'BASELINE: a baseline is the whole declared matrix for a scenario.'
} finally {
    # --- shutdown, the accepted path, leaf before parent --------------------
    $rel = New-ReleaseLedger 'phase-10 benchmark instance'
    try {
        if ($null -ne $wb) {
            # NEVER SAVED. The disposable copy is discarded, no measurement
            # carries a save cost, and no benchmark number ever reaches a cell.
            try { $wb.Close($false); $rel.WorkbookClosed = $true }
            catch { $null = $rel.Failed.Add('Workbook.Close') }
        }
        Invoke-NamedRelease $rel $wb        'Workbook';  $wb        = $null
        Invoke-NamedRelease $rel $workbooks 'Workbooks'; $workbooks = $null
        if ($null -ne $excel) {
            try { $excel.Quit(); $rel.QuitCalled = $true }
            catch { $null = $rel.Failed.Add('Application.Quit') }
        }
        Invoke-NamedRelease $rel $excel 'Application'; $excel = $null

        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

        $rel.NaturalExit = Wait-ExcelExit -Identity $excelIdentity
        if (-not $rel.NaturalExit) {
            $rel.EmergencyRequired = $true
            Write-BenchmarkLine (Invoke-EmergencyExcelCleanup -Identity $excelIdentity `
                -Label 'phase-10 benchmark instance')
        }
    } catch {
        Write-BenchmarkLine ('Shutdown raised: ' + (Format-Err $_))
    }
    Write-BenchmarkLine 'SHUTDOWN'
    Write-BenchmarkLine '--------'
    Write-BenchmarkLine (Format-ReleaseLedger $rel)
    $transient = @(Get-TransientFailures)
    if ($transient.Count -gt 0) {
        Write-BenchmarkLine ('transient COM release failures: ' + ($transient -join '; '))
    } else {
        Write-BenchmarkLine 'every transient COM object released cleanly'
    }
}
$sessionWatch.Stop()

# ===========================================================================
# THE ARTIFACTS
# ===========================================================================
# ONE MACHINE-READABLE FILE AND ONE HUMAN-READABLE ONE, from the same object.
# Neither is written into the workbook: a benchmark number is evidence about a
# machine, and a workbook cell that held one would be a state authority nobody
# asked for.
$shutdownRecord = New-Object System.Collections.Specialized.OrderedDictionary
$shutdownRecord.Add('workbook_closed', [bool]$rel.WorkbookClosed)
$shutdownRecord.Add('quit_called', [bool]$rel.QuitCalled)
$shutdownRecord.Add('natural_exit', [bool]$rel.NaturalExit)
$shutdownRecord.Add('emergency_required', [bool]$rel.EmergencyRequired)
$shutdownRecord.Add('failed_releases', @($rel.Failed))
$shutdownRecord.Add('transient_release_failures', @(Get-TransientFailures))
$shutdownRecord.Add('workbook_saved', $false)

if ($null -eq $environment) {
    $environment = Get-BenchmarkEnvironment -Excel $null -Identity $excelIdentity `
        -WorkbookPath $stageBPath -RepositoryPath $repoRoot -ReleaseIdentity $releaseIdentity `
        -HarnessVersion ([string]$plan.harness_version) `
        -SchemaVersion ([int]$plan.schema_version) -Revision $revision
}

# ===========================================================================
# WAS THIS A BASELINE?
# ===========================================================================
# W1 SETTLED THIS BY FAILING. Windows Run 1 aborted in setup, produced no timed
# operation at all, and the only thing it timed was a 68.4 s Stage-B bootstrap -
# which is SETUP and is not a user-operation figure under any reading. A run
# like that must not be able to look like evidence, so the artifact says what it
# was in a field, the summary says it in a sentence, and the process says it in
# an exit code.
#
# THREE OUTCOMES, AND THEY ARE DIFFERENT THINGS:
#
#   BASELINE RECORDED  every run the plan declares for this scenario completed
#                      with a cold sample and three valid warm samples. This is
#                      the only outcome that is a baseline.
#
#   SCOPED RUN         the operator narrowed the run with -Iterations or
#                      -Operations. Everything asked for succeeded; it is
#                      complete evidence for what was asked and is NOT the
#                      scenario's baseline.
#
#   ABORTED            something failed, refused, or produced fewer than three
#                      valid warm samples. Not a baseline, not a partial median,
#                      and not a successful benchmark.
$plannedCount = @($plannedRuns).Count
$completed = @($results | Where-Object { [bool]$_['valid'] -and ($null -ne $_['warm_median_ms']) })
$runComplete = ([bool](([string]::IsNullOrWhiteSpace($abandoned)) -and
                       (@($completed).Count -eq $plannedCount) -and ($plannedCount -gt 0)))
# `@($null).Count` IS 1, NOT 0. An unsupplied [int[]] parameter is $null, and
# wrapping $null in @() produces a one-element array holding $null - so a plain
# count would have marked every full run as scoped.
$scoped = ([bool]((($null -ne $Iterations) -and (@($Iterations).Count -gt 0)) -or
                  (($null -ne $Operations) -and (@($Operations).Count -gt 0))))
$baselineEstablished = ([bool]($runComplete -and (-not $scoped)))
$baselineStatus = 'ABORTED BEFORE A COMPLETE BASELINE - this run is NOT a baseline, NOT a partial warm median, and NOT a successful benchmark'
if ($runComplete) {
    $baselineStatus = $(if ($scoped) {
        'SCOPED RUN COMPLETE - complete evidence for what was asked, but NOT the scenario baseline'
    } else { 'BASELINE RECORDED' })
}

$report = New-Object System.Collections.Specialized.OrderedDictionary
$report.Add('schema_version', [int]$plan.schema_version)
$report.Add('baseline_id', [string]$plan.baseline_id)
$report.Add('harness_version', [string]$plan.harness_version)
$report.Add('kind', 'pccm-phase10-performance-baseline')
$report.Add('baseline_status', $baselineStatus)
# WHICH FIXTURE BUILT THE WORKBOOK THAT WAS MEASURED. Two methods that reach the
# same state are still two methods; a reader comparing baselines must be able to
# see which one produced each.
$report.Add('fixture_mode', [string]$FixtureMode)
$report.Add('baseline_established', $baselineEstablished)
$report.Add('run_complete', $runComplete)
$report.Add('scoped', $scoped)
$report.Add('runs_planned', $plannedCount)
$report.Add('runs_with_a_valid_warm_median', @($completed).Count)
$report.Add('judgement', ('NONE. This artifact records measurements. No absolute pass ' +
                          'or fail is contracted before a baseline exists.'))
$report.Add('generated_at_utc', ((Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')))
$report.Add('release_identity', $releaseIdentity)
$report.Add('environment', $environment)
$report.Add('scenario', $scenarioSpec)
$report.Add('scenario_actual', $actual)
$report.Add('setup_ms', $setupTimings)
$report.Add('setup_note', ('setup is reported so it can be seen to have been paid. ' +
                           'None of it is inside any operation elapsed time, and none of ' +
                           'it is a user-operation performance figure: a Stage-B bootstrap ' +
                           'time is evidence about a build, never a baseline.'))
$report.Add('timing', $plan.timing)
$report.Add('correctness_gates', $plan.correctness_gates)
$report.Add('regression_policy', $plan.regression_policy)
$report.Add('historical_context', $plan.historical_context)
$report.Add('forbidden', $plan.forbidden)
$report.Add('results', @($results))
$report.Add('abandoned', $abandoned)
$report.Add('failure', $failure)
$report.Add('shutdown', $shutdownRecord)
$report.Add('session_wall_clock_ms', [double]$sessionWatch.Elapsed.TotalMilliseconds)

$jsonPath = Join-Path $OutDir ('phase10_benchmark_' + $Scenario + '_' + $stamp + '.json')
$report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $jsonPath -Encoding UTF8

# --- the human-readable summary -------------------------------------------
$md = New-Object System.Collections.ArrayList
$null = $md.Add('# PCCM performance baseline - ' + $Scenario)
$null = $md.Add('')
$null = $md.Add('**Status: ' + $baselineStatus + '**')
$null = $md.Add('')
$null = $md.Add('**This records. It does not judge.** No absolute pass or fail is')
$null = $md.Add('contracted before a baseline exists on the target machine.')
$null = $md.Add('')
$null = $md.Add('| | |')
$null = $md.Add('|---|---|')
$null = $md.Add('| Baseline id | `' + [string]$plan.baseline_id + '` |')
$null = $md.Add('| Baseline established | ' + $(if ($baselineEstablished) { 'yes' } else { '**no**' }) + ' |')
$null = $md.Add('| Runs with a valid warm median | ' + [string]@($completed).Count +
                ' of ' + [string]$plannedCount + ' planned |')
$null = $md.Add('| Scenario | ' + $Scenario + ' - ' + [string]$scenarioSpec.title + ' |')
$null = $md.Add('| Drivers | ' + [string]$actual['drivers'] + ' (' +
                [string]$actual['cost_lines'] + ' Cost Lines + ' + [string]$actual['risks'] + ' Risks) |')
$null = $md.Add('| Project years | ' + [string]$actual['years'] + ' |')
$null = $md.Add('| Seed | FIXED ' + [string]$actual['seed'] + ' |')
$null = $md.Add('| Commit | `' + [string]$environment['git_commit'] + '` |')
$null = $md.Add('| Model / builder version | ' + [string]$environment['model_version'] +
                ' / ' + [string]$environment['builder_version'] + ' |')
$null = $md.Add('| Excel | ' + [string]$environment['excel_version'] + ' build ' +
                [string]$environment['excel_build'] + ', ' + [string]$environment['excel_bitness'] + ' |')
$null = $md.Add('| Windows | ' + [string]$environment['windows_edition'] + ' ' +
                [string]$environment['windows_version'] + ' build ' + [string]$environment['windows_build'] + ' |')
$null = $md.Add('| CPU / RAM | ' + [string]$environment['cpu_model'] + ' / ' +
                [string]$environment['installed_ram_gb'] + ' GB |')
$null = $md.Add('| Workbook location | ' + [string]$environment['workbook_location_type'] +
                ' - `' + [string]$environment['workbook_path'] + '` |')
$null = $md.Add('| Repository location | ' + [string]$environment['repository_location_type'] +
                ' - `' + [string]$environment['repository_path'] + '` |')
$null = $md.Add('| Other workbooks open | ' + [string]$environment['other_workbooks_open'] + ' |')
$null = $md.Add('')
$null = $md.Add('## Measurements')
$null = $md.Add('')
$null = $md.Add('| Operation | Iterations | Cold | Warm 1 | Warm 2 | Warm 3 | **Warm median** | Valid |')
$null = $md.Add('|---|---|---|---|---|---|---|---|')
foreach ($row in $results) {
    $warmCells = @()
    foreach ($value in @($row['warm_ms'])) { $warmCells += ('{0:N3} s' -f ([double]$value / 1000.0)) }
    while ($warmCells.Count -lt 3) { $warmCells += '-' }
    $medianCell = '-'
    if ($null -ne $row['warm_median_ms']) {
        $medianCell = '**' + ('{0:N3} s' -f ([double]$row['warm_median_ms'] / 1000.0)) + '**'
    }
    $iterationCell = '-'
    if ($null -ne $row['iterations_requested']) { $iterationCell = [string]$row['iterations_requested'] }
    $null = $md.Add('| ' + [string]$row['operation_label'] + ' | ' + $iterationCell + ' | ' +
                    ('{0:N3} s' -f ([double]$row['cold_ms'] / 1000.0)) + ' | ' +
                    ($warmCells -join ' | ') + ' | ' + $medianCell + ' | ' +
                    $(if ([bool]$row['valid']) { 'yes' } else { 'NO' }) + ' |')
}
$null = $md.Add('')
$null = $md.Add('Setup, excluded from every figure above: Stage-B bootstrap ' +
                ('{0:N1} s' -f ([double]$setupTimings['stage_b_bootstrap_ms'] / 1000.0)) +
                $(if ($setupTimings.Contains('scenario_fixture_ms')) {
                    ', scenario fixture ' +
                    ('{0:N1} s' -f ([double]$setupTimings['scenario_fixture_ms'] / 1000.0)) } else { '' }) + '.')
$null = $md.Add('')
if (-not $baselineEstablished) {
    $null = $md.Add('> This run did **not** establish the ' + $Scenario + ' baseline, so')
    $null = $md.Add('> nothing in it may be quoted as one and nothing may be compared')
    $null = $md.Add('> against it under the policy below.')
    if ($null -ne $failure) {
        $null = $md.Add('>')
        $null = $md.Add('> It stopped in stage **' + [string]$failure['stage'] + '** while ' +
                        [string]$failure['action'] + ': `' + [string]$failure['message'] + '`')
    }
    $null = $md.Add('')
}
$null = $md.Add('## How a later run is compared with this one')
$null = $md.Add('')
foreach ($rule in @($plan.regression_policy.rules)) { $null = $md.Add('- ' + [string]$rule) }
$null = $md.Add('')
$null = $md.Add('Compared warm median against warm median, for the same plan, scenario,')
$null = $md.Add('operation and iteration count.')
$null = $md.Add('')
$null = $md.Add('## The Phase-7 timings')
$null = $md.Add('')
$null = $md.Add([string]$plan.historical_context.status + '. Source: ' +
                [string]$plan.historical_context.source + '.')
$null = $md.Add('')
foreach ($observation in @($plan.historical_context.observations)) {
    $null = $md.Add('- ' + [string]$observation.operation + ', ' + [string]$observation.drivers +
                    ' drivers x ' + [string]$observation.iterations + ' iterations: ~' +
                    [string]$observation.seconds + ' s')
}
$mdPath = Join-Path $OutDir ('phase10_benchmark_' + $Scenario + '_' + $stamp + '.md')
Set-Content -LiteralPath $mdPath -Value ($md -join "`r`n") -Encoding UTF8

Write-BenchmarkLine ''
Write-BenchmarkLine 'SUMMARY'
Write-BenchmarkLine '-------'
foreach ($row in $results) {
    $shown = 'NOT COMPUTED'
    if ($null -ne $row['warm_median_ms']) { $shown = (Format-BenchmarkSeconds ([double]$row['warm_median_ms'])) }
    $iterationCell = ''
    if ($null -ne $row['iterations_requested']) { $iterationCell = ' @ ' + [string]$row['iterations_requested'] }
    Write-BenchmarkLine ('  ' + ([string]$row['operation_label'] + $iterationCell).PadRight(42) +
                         ' warm median ' + $shown +
                         $(if ([bool]$row['valid']) { '' } else { '   (INVALID SAMPLE)' }))
}
Write-BenchmarkLine ''
Write-BenchmarkLine ('  whole session        : ' + (Format-BenchmarkSeconds $sessionWatch.Elapsed.TotalMilliseconds))
Write-BenchmarkLine ''
Write-BenchmarkLine ('  BASELINE STATUS      : ' + $baselineStatus)
Write-BenchmarkLine ('  valid warm medians   : ' + [string]@($completed).Count +
                     ' of ' + [string]$plannedCount + ' planned run(s)')
Write-BenchmarkLine ''
Write-BenchmarkLine 'NO CONCLUSION IS DRAWN HERE. These numbers are the first delivery'
Write-BenchmarkLine 'baseline; whether any of them is acceptable is a decision taken'
Write-BenchmarkLine 'against this evidence, not inside this file.'
Write-BenchmarkLine ''
Write-BenchmarkLine ('  json                 : ' + $jsonPath)
Write-BenchmarkLine ('  markdown             : ' + $mdPath)
Write-BenchmarkLine ('  log                  : ' + $script:BenchmarkTextPath)

if ($KeepArtifacts) {
    Write-Host ''
    Write-Host ('Working copy kept in ' + $tempRoot) -ForegroundColor Yellow
} elseif ($OutDir -ne $tempRoot) {
    try { Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue }
    catch { Write-Host ('  (the working copy could not be removed: ' + $tempRoot + ')') -ForegroundColor DarkYellow }
}
Write-Host ''
Write-Host ('Benchmark artifacts in ' + $OutDir) -ForegroundColor Yellow

# AND THE PROCESS SAYS IT TOO. A run that produced no complete measurement must
# not exit 0: a caller, a scheduled task or a transcript reader would otherwise
# record an abort as a success.
if (-not $runComplete) {
    Write-Host ''
    Write-Host $baselineStatus -ForegroundColor Red
    exit 1
}
if (-not $baselineEstablished) {
    Write-Host ''
    Write-Host $baselineStatus -ForegroundColor Yellow
}
exit 0
