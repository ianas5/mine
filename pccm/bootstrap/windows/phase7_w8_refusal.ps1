<#
.SYNOPSIS
    PCCM Phase 7 - the MINIMAL W8 runner: the annual endpoint REFUSES, twice,
    and the answer it refuses to replace is still there afterwards.

.DESCRIPTION
    W8 IS THE ONE SCENARIO WHERE NOTHING IS PRODUCED, and that is the whole
    point. A refusal is only worth anything if it costs nothing: the run that
    already succeeded must survive it intact, and the sheet must still say - in
    the contract's own words - which answer is published and which run owns it.

    ONE BASELINE AND TWO STATES, REACHED THE WAY A USER REACHES THEM:

      BASELINE  five drivers, four project years, FIXED seed, the projected
                business-minimum iteration count, selected Px as the corpus
                names it. Calculate, simulate, publish the annual answer, and
                reconcile that answer before anything is allowed to disturb it.
      PART A    the SIMULATION REQUEST moves and Phase 5 does not. The Monte
                Carlo iteration control is a Phase-6 request input; it enters
                the request fingerprint and no Phase-5 calculation, so moving it
                to another valid projected value makes the published run stale
                and nothing else.
      PART B    the MODEL itself moves, through the accepted cost-line register:
                one driver's maximum is put below its own minimum, which is the
                ordering the accepted numerical checker refuses. Phase 5 can no
                longer be prepared at all, so no current request fingerprint can
                be formed and the simulation is invalid.

    NO HIDDEN STATE IS TOUCHED. Not one cell of `_Calc`, not one cell of
    `_SimData`, no stamp, no fingerprint and no digest is written by this
    runner. Both states are produced by writing to a control and to a register
    the accepted contracts project, and every status is production's own answer
    to production's own question.

    TWO AXES, AND W8 EXISTS TO KEEP THEM APART. A PERSISTENT STATE says what the
    model, the simulation and the annual answer ARE; an ATTEMPT RESULT says what
    happened the last time somebody pressed something. REFUSED belongs to the
    second and may never appear in the first - "an invalid model is INVALID
    whether or not anyone pressed Calculate". So the runner projects four
    vocabularies, proves they are disjoint before it uses them, and then proves
    at runtime that no accessor and no persisted state cell ever answers with a
    word from the attempt axis.

    PHYSICAL PERSISTENCE AND SEMANTIC AUTHORITY ARE PROVED SEPARATELY. The
    stored payload is compared directly - the whole simulation block, the whole
    annual surface, every published iteration value - and the accessors are
    asked separately what that payload now MEANS. A refusal must move neither,
    and the second fact is not evidence for the first.

    WHAT W8 IS NOT. It moves no reporting selector - W6 owns that - cycles no
    bank and shrinks no duration - W7 owns those - runs no sensitivity, and
    produces no second successful run of anything. The baseline is reconciled;
    the two refusals produce nothing to reconcile.

.PARAMETER BuildDir
    The Stage-A build directory to copy from. Defaults to <repo>/pccm/build.

.NOTES
    WINDOWS POWERSHELL 5.1 is the target shell. `Join-Path a b c` - a child per
    positional argument - is PowerShell 6+ only, and the roots below are
    resolved the way the harnesses that have actually run on 5.1 resolve them.
#>

[CmdletBinding()]
param(
    [string]$BuildDir
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# ALL THREE ARE DEFINITION-ONLY AT TOP LEVEL. Dot-sourcing them defines
# functions and script variables and runs no scenario.
. (Join-Path $scriptDir 'com_lifecycle.ps1')
. (Join-Path $scriptDir 'phase5_gate_b_scenarios.ps1')
. (Join-Path $scriptDir 'phase6_gate_b_scenarios.ps1')

$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
$repoRoot = Split-Path -Parent $pccmRoot
if ([string]::IsNullOrWhiteSpace($BuildDir)) { $BuildDir = Join-Path $pccmRoot 'build' }

# ===========================================================================
# THE TEN HELPERS COPIED FROM THE ACCEPTED PHASE-7 TIMING HARNESS
# ===========================================================================
# THE ACCEPTED PHASE-5 FIXTURE CHOREOGRAPHY CALLS THEM and its own file does not
# define them, because in a Gate-B run the Phase-4 driver dot-sources the
# scenarios file and the helpers are already in scope. Reaching the scenarios
# file directly leaves that dependency unmet, and W1 died on exactly that, one
# helper at a time.
#
# FOUR OF THE TEN ARE ALSO CALLED HERE. W8 reads and writes the projected
# simulation request through `Get-NamedValue` and `Set-NamedValue`, and reaches
# the cost-line register through `Get-TableBody` and `Set-TableCell` - the same
# two primitives the accepted fixture writer uses, at an ordinal taken from the
# manifest's own column list rather than from an Excel header caption.
#
# They are copied BYTE FOR BYTE from `phase7_timing_scenarios.ps1` so there is
# one behaviour rather than two, and a source control pins them to it.
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

function Get-TableRowCount {
    param($Workbook, [string]$SheetName, [string]$TableName)
    return @(Get-TableBody -Workbook $Workbook -SheetName $SheetName -TableName $TableName).Count
}

function Add-BlankTableRow {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $rows = $null; $added = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $rows = $lo.ListRows
        $added = $rows.Add()
        return [int]$added.Index
    } finally {
        if ($null -ne $added)           { Release-Transient $added           'ListRow';     $added           = $null }
        if ($null -ne $rows)            { Release-Transient $rows            'ListRows';    $rows            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
}

function Remove-TableRow {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$RowIndex)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $rows = $null; $victim = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $rows = $lo.ListRows
        $victim = $rows.Item($RowIndex)
        $victim.Delete()
    } finally {
        if ($null -ne $victim)          { Release-Transient $victim          'ListRow';     $victim          = $null }
        if ($null -ne $rows)            { Release-Transient $rows            'ListRows';    $rows            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
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
# THE REPORT AND THE CHECK LEDGER
# ===========================================================================
$script:W8Lines = New-Object System.Collections.ArrayList
$script:W8Path = ''
$script:W8Checks = New-Object System.Collections.ArrayList

function Write-W8Line {
    param([string]$Text = '')
    $null = $script:W8Lines.Add($Text)
    Write-Host $Text
    # Written through on every line, so a run that is stopped still leaves every
    # observation it had already taken on disk.
    if (-not [string]::IsNullOrWhiteSpace($script:W8Path)) {
        try {
            Set-Content -LiteralPath $script:W8Path `
                -Value ($script:W8Lines -join "`r`n") -Encoding UTF8
        } catch { }
    }
}

# ONE CHECK. `Kind` separates a PREREQUISITE - a step W8 had to perform to reach
# its subject - from a RESULT, which is W8's own claim about the refusal. A
# failed prerequisite still fails the invocation; it is just not evidence about
# the property under test, and the report says which it was.
function Add-W8Check {
    param([string]$Label, [bool]$Ok, [string]$Detail = '', [string]$Kind = 'RESULT')
    $null = $script:W8Checks.Add([pscustomobject]@{
        Kind = $Kind; Label = $Label; Ok = $Ok; Detail = $Detail
    })
    $verdict = 'FAIL'
    if ($Ok) { $verdict = 'PASS' }
    $line = '  [' + $verdict + '] ' + $Label
    if (-not [string]::IsNullOrWhiteSpace($Detail)) { $line = $line + ' -- ' + $Detail }
    Write-W8Line $line
    return $Ok
}

# ===========================================================================
# EVERY RELEASE, WITH ITS COUNT
# ===========================================================================
# THE DISCIPLINE W1 CLOSED ON. ReleaseComObject returns the count REMAINING on
# that RCW: zero means the runtime let go, and anything above zero is a retained
# reference, named. `Invoke-NamedRelease` writes the ledger line but keeps that
# number to itself, and `Release-Transient` says nothing at all on success -
# which is how W1's first failure went unexplained. This wraps the SAME accepted
# primitive and records the number. com_lifecycle.ps1 is not modified.
$script:W8Residual = New-Object System.Collections.ArrayList

function Invoke-W8Release {
    param($Ledger, $Obj, [string]$Label)
    $rec = Release-ComObjectSafe -Obj $Obj -Label $Label
    if ($rec.Status -eq 'PASS') {
        $Ledger.Attempted = $Ledger.Attempted + 1
        $Ledger.Succeeded = $Ledger.Succeeded + 1
        $null = $Ledger.Lines.Add(
            ("      {0,-24} | PASS    | ReleaseComObject returned {1}" -f $rec.Label, $rec.Count))
        if ([int]$rec.Count -ne 0) {
            $null = $script:W8Residual.Add(
                ([string]$rec.Label + ' left ' + [string]$rec.Count + ' outstanding'))
        }
    } elseif ($rec.Status -eq 'FAIL') {
        $Ledger.Attempted = $Ledger.Attempted + 1
        $null = $Ledger.Failed.Add($rec.Label)
        $null = $Ledger.Lines.Add(("      {0,-24} | FAIL    | {1}" -f $rec.Label, $rec.Error))
    } else {
        $null = $Ledger.Lines.Add(("      {0,-24} | SKIPPED | {1}" -f $rec.Label, $rec.Error))
    }
}

# ===========================================================================
# THE CANDIDATE IDENTITY
# ===========================================================================
function Get-W8SourceRevision {
    param([string]$RepoRoot)
    $head = ''
    try { $head = [string](& git -C $RepoRoot rev-parse HEAD 2>$null) } catch { $head = '' }
    $head = $head.Trim()
    if ([string]::IsNullOrWhiteSpace($head)) {
        throw ('git could not report HEAD for ' + $RepoRoot +
               '; a W8 result could not be attributed to a source revision')
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
# THE STATE WORDS, PROJECTED - NEVER TYPED
# ===========================================================================
# SIX VOCABULARIES, TWO ARTEFACTS, AND NOT ONE MEMBER SPELLED IN THIS FILE:
#
#   sim_states            phase6_gate_b_cases.json         the SIMULATION's
#   attempt_results       phase6_gate_b_cases.json         persistent axis and
#                                                          its attempt axis
#   derived_status        phase7_acceptance_inspection.json the CALCULATION's
#   attempt_result        phase7_acceptance_inspection.json two axes
#   distribution_states   phase7_acceptance_inspection.json the ANNUAL handoff's
#   profile_states        phase7_acceptance_inspection.json two products
#
# ORDINALS, NOT LITERALS. A contract vocabulary is an ORDERED declaration and
# the position of a member in it is as much the contract's statement as the
# spelling is. W6 and W7 already read the current annual state as
# `distribution_states[1]`; W8 reads the rest the same way, and proves the
# SHAPE it is indexing into before it indexes - so a contract that reordered or
# resized a list fails a prerequisite here rather than silently renaming a state
# W8 then goes on to assert.
function Get-W8Vocabulary {
    param($GateBCases, $P7)
    $sim = @($GateBCases.vocabulary.sim_states | ForEach-Object { [string]$_ })
    $simAttempts = @($GateBCases.vocabulary.attempt_results | ForEach-Object { [string]$_ })
    $calc = @($P7.model_states.derived_status | ForEach-Object { [string]$_ })
    $calcAttempts = @($P7.model_states.attempt_result | ForEach-Object { [string]$_ })
    $distribution = @($P7.handoff.distribution_states | ForEach-Object { [string]$_ })
    $profile = @($P7.handoff.profile_states | ForEach-Object { [string]$_ })
    return [pscustomobject]@{
        SimStates            = $sim
        SimAttempts          = $simAttempts
        CalcStates           = $calc
        CalcAttempts         = $calcAttempts
        DistributionStates   = $distribution
        ProfileStates        = $profile
        SimCurrent           = [string]$sim[0]
        SimStale             = [string]$sim[1]
        SimInvalid           = [string]$sim[2]
        CalcCurrent          = [string]$calc[1]
        CalcInvalid          = [string]$calc[3]
        CalcRefused          = [string]$calcAttempts[2]
        AnnualNotProduced    = [string]$distribution[0]
        AnnualCurrent        = [string]$distribution[1]
        # THE PROFILE'S OWN LIST, not the distribution's. The two products do
        # not share a currentness rule, so they do not share a vocabulary
        # either - and reading one product's state out of the other's list is
        # the collapse the contract declares two lists to prevent.
        AnnualProfileCurrent = [string]$profile[1]
        # THE STATE A PUBLISHED ANNUAL ANSWER CARRIES ONCE ITS OWNING RUN IS NO
        # LONGER THE CURRENT ONE. The contract declares the word as the LAST
        # member of BOTH handoff lists and gives it no key of its own - unlike
        # the inconsistent-stamp state, which has one - so it is taken as the
        # last member of each and required below to be the same word in both.
        DistributionSuperseded = [string]$distribution[$distribution.Count - 1]
        ProfileSuperseded      = [string]$profile[$profile.Count - 1]
        InconsistentStamp      = [string]$P7.handoff.inconsistent_stamp_state
    }
}

# The members two vocabularies share. Empty is the answer every call site wants.
function Get-W8SharedMembers {
    param($First, $Second)
    $shared = @()
    foreach ($item in @($First)) {
        foreach ($other in @($Second)) {
            if (([string]$item) -ceq ([string]$other)) { $shared += [string]$item }
        }
    }
    return @($shared)
}

# EXACT, BINARY, AND A NON-STRING IS NEVER A MEMBER. A state word published as a
# number is a different fact from a state word, and coercing it here would hide
# the one thing membership is being asked about.
function Test-W8Member {
    param($Value, $List)
    if ($null -eq $Value) { return $false }
    if ($Value -isnot [string]) { return $false }
    foreach ($item in @($List)) {
        if (([string]$Value) -ceq ([string]$item)) { return $true }
    }
    return $false
}

# ===========================================================================
# WHICH PERSISTED ROWS A REFUSAL MAY MOVE, AND WHICH IT MAY NOT
# ===========================================================================
# THE PARTITION IS THE PROJECTION'S, not a list typed here. Every row of the
# run-identity block carries a `group` in phase6_gate_b_inspection.json, and the
# groups already say what each row is for:
#
#   snapshot  the published run's own identity, per bank - run id, fingerprint,
#             digest, seeds, iterations run, applied timeline
#   counter   the AUTO nonce and the last allocated run id
#   attempt   what happened the last time a SIMULATION was attempted
#   derived   the simulation status and the moment it was evaluated
#   control   the active publication bank
#
# ONLY `derived` IS EXCLUDED FROM THE FROZEN SET, and not quietly.
# PCCM_SimulationStatus recomputes and rewrites both of its rows every time it
# is asked, and the annual endpoint's own precondition asks it. Requiring the
# evaluation TIMESTAMP not to move would be requiring the annual step not to
# check whether it may run.
#
# THE `attempt` ROWS ARE FROZEN, DELIBERATELY. The annual endpoint is not a
# simulation and states that it "allocates no run id, advances no AUTO nonce,
# touches no pending-nonce marker, writes no attempt row". That is a CLAIM, so
# W8 treats it as one: the rows are held in the frozen set and a refusal that
# wrote a simulation attempt row fails a RESULT check rather than being excused
# by a rule this runner wrote for it.
function Get-W8RowsInGroup {
    param($Inspection, [string]$Group)
    $out = @()
    $identity = $Inspection.sim_data.run_identity
    foreach ($key in $identity.rows.PSObject.Properties.Name) {
        if (([string]$identity.groups.$key) -ceq $Group) { $out += [string]$key }
    }
    return @($out)
}

function Get-W8RunInvariants {
    param($Workbook, $Inspection)
    $excluded = @(Get-W8RowsInGroup -Inspection $Inspection -Group 'derived')
    $state = Get-Phase6State -Workbook $Workbook -Inspection $Inspection
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($key in $state['shared'].Keys) {
        if ($excluded -contains [string]$key) { continue }
        $out.Add(('shared.' + $key), $state['shared'][$key])
    }
    foreach ($bank in @($Inspection.publication.bank_labels)) {
        $block = $state[('bank_' + $bank)]
        foreach ($key in $block.Keys) { $out.Add(('bank_' + $bank + '.' + $key), $block[$key]) }
    }
    $out.Add('pending_auto_nonce', $state['pending_auto_nonce'])
    return $out
}

function Add-W8InvariantChecks {
    param([string]$Label, $Before, $After)
    $moved = @()
    foreach ($key in $Before.Keys) {
        if (-not (Test-SimSameValue -A $Before[$key] -B $After[$key])) {
            $moved += ([string]$key + ': ' + (Format-SimValue $Before[$key]) + ' -> ' +
                       (Format-SimValue $After[$key]))
        }
    }
    return (Add-W8Check ($Label + ': no run identity, nonce, pending marker, digest or ' +
                         'simulation attempt row moved') `
        ($moved.Count -eq 0) ($moved -join '; '))
}

# THE ROWS THE FROZEN SET LEAVES OUT, PLUS THE ATTEMPT ROWS, CAPTURED ON THEIR
# OWN. They are reported rather than silently tolerated, and the two facts that
# matter about them are asserted separately: the STATUS WORD is what the
# contract says it should be, and the timestamp beside it is free to move.
function Get-W8StatusMetadata {
    param($Workbook, $Inspection)
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($group in @('attempt', 'derived')) {
        foreach ($key in @(Get-W8RowsInGroup -Inspection $Inspection -Group $group)) {
            $out.Add(($group + '.' + [string]$key),
                     (Get-SimField -Workbook $Workbook -Inspection $Inspection -FieldKey $key))
        }
    }
    return $out
}

# WHICH ROW CARRIES A STATE WORD - asked of the values, not of the key names. A
# row is the status row when what it holds is a member of the projected
# persistent vocabulary, and the caller requires there to be exactly one. That
# is the same question W8 asks everywhere else, and it needs no row name typed
# into this file to ask it.
function Get-W8RowsHolding {
    param($Block, $List, [string]$Prefix = '')
    $out = @()
    foreach ($key in $Block.Keys) {
        $name = [string]$key
        if ((-not [string]::IsNullOrEmpty($Prefix)) -and (-not $name.StartsWith($Prefix))) { continue }
        if (Test-W8Member -Value $Block[$name] -List $List) { $out += $name }
    }
    return @($out)
}

# ===========================================================================
# COLUMN ARITHMETIC, AND THE ANNUAL LADDER IT REACHES
# ===========================================================================
# The projection gives the FIRST ladder column per bank per measure and the
# ladder's LENGTH; the eleven columns follow it. That is the contract's shape,
# so the offset is computed rather than eleven letters being typed in.
#
# THE DEFECT THAT STOPPED THE FIRST W5 RUN, and the reason it survived every
# static control: this loop used to read
#
#     foreach ($character in [string]$Letters.ToUpperInvariant().ToCharArray())
#
# and a cast binds tighter than the enumeration, so `[string]` was applied to
# the whole char[] rather than to each character. PowerShell renders an array as
# a string by JOINING its elements with $OFS - a single space by default - so
# ['A','D'] became the ONE-element sequence "A D", and the body then asked for
# [char]"A D": "String must be exactly one character long."
#
# It worked for every single-letter column, because ['A'] renders as "A" and
# [char]"A" is fine. The first multi-letter column it ever met was the projected
# quantile first column, and that is exactly where the run died. There is no
# cast here now: .ToCharArray() already yields the chars, one at a time, which
# is the form the accepted harness has always used.
function ConvertTo-W8ColumnNumber {
    param([string]$Letters)
    $number = 0
    foreach ($character in $Letters.ToUpperInvariant().ToCharArray()) {
        $number = ($number * 26) + ([int]$character - 64)
    }
    return $number
}

# The inverse, in the same bijective base 26: there is no zero digit, so the
# remainder is taken on (n - 1) and the quotient steps down by the digit that
# was just emitted. Z -> AA is the boundary that proves it.
function ConvertFrom-W8ColumnNumber {
    param([int]$Number)
    $letters = ''
    $remaining = $Number
    while ($remaining -gt 0) {
        $remainder = ($remaining - 1) % 26
        $letters = ([string][char](65 + $remainder)) + $letters
        $remaining = [int](($remaining - $remainder - 1) / 26)
    }
    return $letters
}

# ===========================================================================
# THE PHASE-7 ANNUAL SURFACE, READ WITHOUT RUNNING IT
# ===========================================================================
# EVERY ADDRESS COMES FROM THE PROJECTION. Not one column letter or row number
# is written here: they are read from phase7_acceptance_inspection.json, which
# is projected from sim_contract.yaml by the build that projects them into
# modSimContract. A contract move therefore moves this runner too.
function Get-W8AnnualStamp {
    param($Workbook, $Inspection, $P7, [string]$Bank)
    $stamp = $P7.annual_records.stamp
    $column = [string]$stamp.bank_value_columns.$Bank
    if ([string]::IsNullOrEmpty($column)) {
        throw ('the annual stamp projection carries no column for bank ' + [char]39 + $Bank + [char]39)
    }
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($key in $stamp.rows.PSObject.Properties.Name) {
        $address = $column + [string]([int]$stamp.rows.$key)
        $out.Add($key, (Get-SimRawCell -Workbook $Workbook -Inspection $Inspection -Address $address))
    }
    return $out
}

# ONE ANNUAL RECORD, WHOLE: its index pair, both ladders, both profile values.
# Every address comes from the projection.
function Get-W8AnnualRecord {
    param($Workbook, $Inspection, $P7, [string]$Bank, [int]$Offset)
    $records = $P7.annual_records
    $row = [int]$records.first_record_row + $Offset
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($key in $records.index_columns.$Bank.PSObject.Properties.Name) {
        $column = [string]$records.index_columns.$Bank.$key
        $out.Add($key, (Get-SimRawCell -Workbook $Workbook -Inspection $Inspection `
            -Address ($column + [string]$row)))
    }
    $count = [int]$records.quantile_count
    foreach ($measure in $records.quantile_first_column.$Bank.PSObject.Properties.Name) {
        $first = ConvertTo-W8ColumnNumber -Letters ([string]$records.quantile_first_column.$Bank.$measure)
        $ladder = New-Object System.Collections.ArrayList
        for ($index = 0; $index -lt $count; $index++) {
            $column = ConvertFrom-W8ColumnNumber -Number ($first + $index)
            $null = $ladder.Add((Get-SimRawCell -Workbook $Workbook -Inspection $Inspection `
                -Address ($column + [string]$row)))
        }
        $out.Add(('ladder_' + $measure), @($ladder))
    }
    foreach ($measure in $records.selected_px_profile_columns.$Bank.PSObject.Properties.Name) {
        $column = [string]$records.selected_px_profile_columns.$Bank.$measure
        $out.Add(('profile_' + $measure), (Get-SimRawCell -Workbook $Workbook `
            -Inspection $Inspection -Address ($column + [string]$row)))
    }
    return $out
}

# ===========================================================================
# THE WHOLE ANNUAL ANSWER, AS PLAIN DATA
# ===========================================================================
# Captured before a state change and compared after it. What comes back holds
# strings, doubles and nulls only - never a COM object - so a captured surface
# can be held across an endpoint call without keeping anything alive.
function Get-W8AnnualSurface {
    param($Workbook, $Inspection, $P7, [string]$Bank, [int]$YearCount)
    $surface = New-Object System.Collections.Specialized.OrderedDictionary
    $stamp = Get-W8AnnualStamp -Workbook $Workbook -Inspection $Inspection -P7 $P7 -Bank $Bank
    foreach ($key in $stamp.Keys) { $surface.Add(('stamp.' + [string]$key), $stamp[$key]) }
    $ladderCount = [int]$P7.annual_records.quantile_count
    for ($offset = 0; $offset -lt $YearCount; $offset++) {
        $record = Get-W8AnnualRecord -Workbook $Workbook -Inspection $Inspection -P7 $P7 `
            -Bank $Bank -Offset $offset
        $year = 'year' + [string]($offset + 1)
        foreach ($key in $record.Keys) {
            $name = [string]$key
            if ($name.StartsWith('ladder_')) {
                $ladder = @($record[$name])
                for ($index = 0; $index -lt $ladderCount; $index++) {
                    $rung = $null
                    if ($index -lt $ladder.Count) { $rung = $ladder[$index] }
                    $surface.Add(($year + '.' + $name + '.' + [string]($index + 1)), $rung)
                }
            } else {
                $surface.Add(($year + '.' + $name), $record[$name])
            }
        }
    }
    return $surface
}

# ONE BANK, CAPTURED WHOLE: the simulation block AND the annual answer, as plain
# data, so the STORED PAYLOAD can be compared directly rather than through an
# accessor's opinion of it.
function Get-W8BankCapture {
    param($Workbook, $Inspection, $P7, [string]$Bank, [int]$YearCount)
    $capture = New-Object System.Collections.Specialized.OrderedDictionary
    $block = Get-SimBankBlock -Workbook $Workbook -Inspection $Inspection -Bank $Bank
    foreach ($key in $block.Keys) { $capture.Add(('sim.' + [string]$key), $block[$key]) }
    $annual = Get-W8AnnualSurface -Workbook $Workbook -Inspection $Inspection -P7 $P7 `
        -Bank $Bank -YearCount $YearCount
    foreach ($key in $annual.Keys) { $capture.Add(('annual.' + [string]$key), $annual[$key]) }
    return $capture
}

# Two captured surfaces, compared key by key on the accepted "same value" rule -
# type included, so a number that became text is a change.
function Compare-W8Surface {
    param($Before, $After)
    $moved = New-Object System.Collections.ArrayList
    foreach ($key in $Before.Keys) {
        $name = [string]$key
        if (-not (Test-SimSameValue -A $Before[$name] -B $After[$name])) {
            $null = $moved.Add($name + ': ' + (Format-SimValue $Before[$name]) + ' -> ' +
                               (Format-SimValue $After[$name]))
        }
        if ($moved.Count -gt 12) { break }
    }
    $detail = ($moved -join '; ')
    if ($moved.Count -gt 12) { $detail = $detail + ' ... (truncated)' }
    return [pscustomobject]@{ Moved = $moved.Count; Detail = $detail }
}

# THE ITERATION COLUMN, READ AS ONE RANGE rather than as N round trips. The
# three COM objects it opens are released into the run's own ledger and counted
# where they are acquired, so the acquire/release balance stays true.
function Get-W8IterationBlock {
    param($Workbook, $Inspection, [string]$Bank, [int]$Count, $Ledger)
    $block = $Inspection.sim_data.iteration_records
    $first = [string]$block.banks.$Bank.iteration_index
    $last = [string]$block.banks.$Bank.total_pv
    $firstRow = [int]$block.first_iteration_row
    $lastRow = $firstRow + $Count - 1
    $sheets = $null; $sheet = $null; $range = $null
    $acquired = 0
    try {
        $sheets = $Workbook.Worksheets
        $acquired = $acquired + 1
        $sheet = $sheets.Item([string]$Inspection.sim_data.sheet)
        $acquired = $acquired + 1
        $range = $sheet.Range($first + [string]$firstRow + ':' + $last + [string]$lastRow)
        $acquired = $acquired + 1
        return [pscustomobject]@{ Values = $range.Value2; Acquired = $acquired }
    } finally {
        if ($null -ne $range)  { Invoke-W8Release $Ledger $range  'Range(iterations)';   $range  = $null }
        if ($null -ne $sheet)  { Invoke-W8Release $Ledger $sheet  'Worksheet(_SimData)'; $sheet  = $null }
        if ($null -ne $sheets) { Invoke-W8Release $Ledger $sheets 'Worksheets';          $sheets = $null }
    }
}

# The published iteration block, before and after, compared value for value.
function Compare-W8IterationGrid {
    param($Before, $After)
    if (($null -eq $Before) -or ($null -eq $After)) { return 'an iteration block was not read' }
    if (($Before.GetLength(0) -ne $After.GetLength(0)) -or
        ($Before.GetLength(1) -ne $After.GetLength(1))) {
        return 'the iteration block changed shape'
    }
    for ($row = 1; $row -le [int]$Before.GetLength(0); $row++) {
        for ($column = 1; $column -le [int]$Before.GetLength(1); $column++) {
            if (-not (Test-SimSameValue -A $Before[$row, $column] -B $After[$row, $column])) {
                return ('iteration row ' + [string]$row + ' column ' + [string]$column + ': ' +
                        (Format-SimValue $Before[$row, $column]) + ' -> ' +
                        (Format-SimValue $After[$row, $column]))
            }
        }
    }
    return ''
}

function Get-W8Handoff {
    param($Excel, $P7)
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($accessor in @($P7.command_surface.handoff_accessors)) {
        $out.Add([string]$accessor, $Excel.Run([string]$accessor))
    }
    return $out
}

# AN ATTEMPT, THROUGH THE PUBLISHED AUTOMATION SURFACE AND ITS PUBLISHED ANSWER.
# The runner never judges the outcome itself: it asks production what happened.
# The accepted Phase-5 wrapper THROWS on anything but success, which is right for
# a fixture step and wrong for W8 - here a refusal is the subject, not a fault.
function Invoke-W8Endpoint {
    param($Excel, [string]$Endpoint)
    $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
    $Excel.Run($Endpoint) | Out-Null
    return [string]$Excel.Run('PCCM_AutomationResult')
}

# A defined name that carries TEXT. Set-NamedValue writes a Double - PowerShell
# binds a COM property call site per argument type, so one polymorphic
# assignment silently fails the second time it is reached with another type -
# and the selected confidence level is a label, not a number.
function Set-W8NamedText {
    param($Workbook, [string]$DefinedName, [string]$Value)
    $names = $null; $nm = $null; $rng = $null
    try {
        $names = $Workbook.Names
        $nm = $names.Item($DefinedName)
        $rng = $nm.RefersToRange
        $rng.Value2 = [string]$Value
    } finally {
        if ($null -ne $rng)   { Release-Transient $rng   'Range(name)'; $rng   = $null }
        if ($null -ne $nm)    { Release-Transient $nm    'Name';        $nm    = $null }
        if ($null -ne $names) { Release-Transient $names 'Names';       $names = $null }
    }
}

# ===========================================================================
# THE CONTRACT'S OWN TYPE-7 PERCENTILE
# ===========================================================================
# sim_contract.yaml statistics.percentile, transcribed and nothing else:
#
#     h  = (n - 1) * p
#     lo = floor(h)
#     hi = min(lo + 1, n - 1)
#     f  = h - lo
#     value = (1 - f) * x[lo] + f * x[hi]
#
# Convex interpolation, on a COPY that is sorted - the contract's
# `sorting: on_copies_only`. This derives the TOTAL percentile from the
# published iteration column. It is deliberately NOT the selected-Px annual
# BLEND: that is production's, and reimplementing it here would compare an
# implementation against a copy of itself.
function Get-W8Type7Value {
    param([double[]]$Values, [double]$Probability)
    $sorted = @($Values | Sort-Object)
    $n = $sorted.Count
    if ($n -lt 1) { throw 'a type-7 percentile needs at least one value' }
    if ($n -eq 1) { return [double]$sorted[0] }
    $h = ($n - 1) * $Probability
    $lo = [int][Math]::Floor($h)
    $hi = [Math]::Min($lo + 1, $n - 1)
    $f = $h - $lo
    return ((1 - $f) * [double]$sorted[$lo]) + ($f * [double]$sorted[$hi])
}

# THE IDENTITY RULE, READ FROM THE CORPUS. |delta| <= max(floor, coefficient *
# max(scale floor, conditioning scale)). The conditioning scale names the
# MAGNITUDE OF THE ARITHMETIC PERFORMED - here the annual terms summed and the
# aggregate they are compared against - never the magnitude of the net result,
# which is the ERRATUM C1 correction this project already accepted.
function Get-W8IdentityAllowance {
    param($Provenance, [double]$ConditioningScale)
    $floor = [double]$Provenance.identity_absolute_floor
    $coefficient = [double]$Provenance.identity_relative_coefficient
    $scaleFloor = [double]$Provenance.conditioning_scale_floor
    $scale = [Math]::Max($scaleFloor, [Math]::Abs($ConditioningScale))
    return [Math]::Max($floor, $coefficient * $scale)
}

# The per-year profile sums and their conditioning scale, read back from the
# persisted records. The blend itself is never recomputed here.
function Get-W8ProfileSums {
    param($Workbook, $Inspection, $P7, [string]$Bank, [int]$YearCount, $Measures)
    $sums = @{}
    $scale = @{}
    $problems = New-Object System.Collections.ArrayList
    foreach ($measure in @($Measures)) { $sums[$measure] = 0.0; $scale[$measure] = 0.0 }
    for ($offset = 0; $offset -lt $YearCount; $offset++) {
        $record = Get-W8AnnualRecord -Workbook $Workbook -Inspection $Inspection -P7 $P7 `
            -Bank $Bank -Offset $offset
        foreach ($measure in @($Measures)) {
            $value = $record[('profile_' + $measure)]
            if ($value -isnot [double]) {
                $null = $problems.Add('year ' + [string]($offset + 1) + ' ' + $measure +
                                      ' profile published ' + (Format-SimValue $value))
                continue
            }
            $sums[$measure] = $sums[$measure] + [double]$value
            $scale[$measure] = $scale[$measure] + [Math]::Abs([double]$value)
        }
    }
    return [pscustomobject]@{ Sums = $sums; Scale = $scale; Problems = @($problems) }
}

# ===========================================================================
# THE BASELINE IS A REAL ANSWER, NOT ONLY A STORAGE SHAPE
# ===========================================================================
# W5's three checks, at the rung the corpus selects. THE BASELINE ONLY: the two
# refusals produce nothing, so there is nothing after them to reconcile, and a
# second reconciliation of an unchanged payload would be the runner agreeing
# with itself.
function Invoke-W8Reconciliation {
    param($Workbook, $Inspection, $P7, $Provenance, $Grid, [string]$Bank,
          [string]$Label, [double]$Probability, [int]$LadderIndex,
          $ProfileSums, $ProfileScale, [string]$Stage)
    $semantics = $P7.summary_semantics
    $rowKey = 'quantile_' + [string]($LadderIndex + 1)
    $contingencyBlock = $Inspection.sim_data.contingency_ladder
    foreach ($measure in @($semantics.contingency_measures | ForEach-Object { [string]$_ })) {
        $column = 2
        if ($measure -eq 'pv') { $column = 3 }
        $totals = New-Object System.Collections.ArrayList
        for ($row = 1; $row -le [int]$Grid.GetLength(0); $row++) {
            $value = $Grid[$row, $column]
            if ($value -is [double]) { $null = $totals.Add([double]$value) }
        }
        $derived = Get-W8Type7Value -Values ([double[]]@($totals)) -Probability $Probability

        # (A) THE TOTAL, from the block the projection names as the total
        # percentile block, cross-checked before anything is reconciled to it.
        $total = Get-SimSummaryValue -Workbook $Workbook -Inspection $Inspection `
            -Bank $Bank -Measure $measure -RowKey $rowKey
        $totalScale = [Math]::Abs($derived)
        if ($total -is [double]) { $totalScale = $totalScale + [Math]::Abs([double]$total) }
        $totalAllowance = Get-W8IdentityAllowance -Provenance $Provenance `
            -ConditioningScale $totalScale
        $totalDelta = [double]::PositiveInfinity
        if ($total -is [double]) { $totalDelta = [Math]::Abs([double]$total - $derived) }
        $null = Add-W8Check ($Stage + ': the published ' + $measure + ' ' + $Label +
                             ' TOTAL equals the contract' + [char]39 +
                             's Type-7 value over the iteration column') `
            (($total -is [double]) -and ($totalDelta -le $totalAllowance)) `
            ('published ' + (Format-SimValue $total) + ', derived ' + [string]$derived +
             ', delta ' + [string]$totalDelta + ', allowance ' + [string]$totalAllowance)

        # (B) sum_y Profile_Px(y) = reported Px TOTAL.
        $sum = [double]$ProfileSums[$measure]
        $scale = [double]$ProfileScale[$measure]
        if ($total -is [double]) { $scale = $scale + [Math]::Abs([double]$total) }
        $identityAllowance = Get-W8IdentityAllowance -Provenance $Provenance `
            -ConditioningScale $scale
        $delta = [double]::PositiveInfinity
        if ($total -is [double]) { $delta = [Math]::Abs($sum - [double]$total) }
        $null = Add-W8Check ($Stage + ': the ' + $measure +
                             ' selected-Px profile sums to the reported ' + $Label + ' TOTAL') `
            (($total -is [double]) -and ($delta -le $identityAllowance)) `
            ('sum ' + [string]$sum + ', total ' + (Format-SimValue $total) +
             ', delta ' + [string]$delta + ', allowance ' + [string]$identityAllowance +
             ' (conditioning scale ' + [string]$scale + ')')

        # (C) THE CONTINGENCY, on the contract's own formula and the SAME
        # measure's deterministic base.
        $baseline = Get-SimSummaryValue -Workbook $Workbook -Inspection $Inspection `
            -Bank $Bank -Measure $measure -RowKey ([string]$semantics.baseline_metric_key)
        $contingency = Get-SimRawCell -Workbook $Workbook -Inspection $Inspection `
            -Address ([string]$contingencyBlock.bank_value_columns.$Bank.$measure +
                      [string]([int]$contingencyBlock.rows.$rowKey))
        $expectedContingency = [double]::NaN
        $contingencyDelta = [double]::PositiveInfinity
        $contingencyScale = 0.0
        if (($total -is [double]) -and ($baseline -is [double])) {
            $expectedContingency = [double]$total - [double]$baseline
            $contingencyScale = [Math]::Abs([double]$total) + [Math]::Abs([double]$baseline)
            if ($contingency -is [double]) {
                $contingencyDelta = [Math]::Abs([double]$contingency - $expectedContingency)
            }
        }
        $contingencyAllowance = Get-W8IdentityAllowance -Provenance $Provenance `
            -ConditioningScale $contingencyScale
        $null = Add-W8Check ($Stage + ': the published ' + $measure + ' ' + $Label +
                             ' contingency is ' + [string]$semantics.contingency_formula) `
            (($contingency -is [double]) -and ($baseline -is [double]) -and
             ($total -is [double]) -and ($contingencyDelta -le $contingencyAllowance)) `
            ('published ' + (Format-SimValue $contingency) + ', total ' +
             (Format-SimValue $total) + ' - base ' + (Format-SimValue $baseline) + ' = ' +
             [string]$expectedContingency + ', delta ' + [string]$contingencyDelta +
             ', allowance ' + [string]$contingencyAllowance)

        Write-W8Line ('    ' + $Stage.PadRight(10) + $measure.PadRight(8) +
                      ' profile sum ' + [string]$sum + '  total ' + (Format-SimValue $total) +
                      '  contingency ' + (Format-SimValue $contingency))
    }
}

# ===========================================================================
# THE PUBLICATION, AND THE BANK THAT MUST STAY EMPTY
# ===========================================================================
# THE AUTHORITY RULE, APPLIED AS THE CONTRACT STATES IT: the marker plus the
# stamped year count, never the last non-blank row.
function Test-W8Authoritative {
    param($Stamp, $P7, [int]$ExpectedYears)
    $marker = [string]$P7.annual_records.stamp.published_marker
    $markerOk = (Test-SimExactText -Actual $Stamp['published'] -Expected $marker)
    $countOk = (Test-SimExactDouble -Actual $Stamp['year_count'] -Expected ([double]$ExpectedYears))
    return [pscustomobject]@{
        Ok = ($markerOk -and $countOk)
        Detail = ('marker ' + (Format-SimValue $Stamp['published']) + ', year_count ' +
                  (Format-SimValue $Stamp['year_count']) + ', expected ' + [string]$ExpectedYears)
    }
}

# THE BANK A SUCCESSFUL RUN IS ABOUT TO PUBLISH TO, from the projected cycle.
# The runner never assumes a blank selector means A.
function Get-W8CandidateBank {
    param($P7, [string]$ActiveBank)
    foreach ($entry in @($P7.publication_semantics.candidate_target)) {
        $active = ''
        if ($null -ne $entry.active_bank) { $active = [string]$entry.active_bank }
        if ($active -ceq $ActiveBank) { return [string]$entry.candidate_bank }
    }
    return ''
}

# THE OTHER LABEL THE PROJECTION DECLARES. Set difference over the projected
# bank labels - not "the one that is not A".
function Get-W8OtherBank {
    param($P7, [string]$Bank)
    foreach ($label in @($P7.publication_semantics.bank_labels)) {
        if (([string]$label) -cne $Bank) { return [string]$label }
    }
    return ''
}

# ===========================================================================
# THE REGISTER EDIT THAT MAKES THE MODEL INVALID
# ===========================================================================
# A NORMAL USER-FACING INPUT AND NOTHING ELSE. One driver's maximum is put below
# its own minimum, in the register the Stage-B manifest projects, at the ordinal
# the manifest's own column list gives - which is exactly how the accepted
# fixture writer reaches the same cell. No Excel header caption is matched, no
# column is counted here, and nothing hidden is touched.
function Get-W8Register {
    param($Manifest, [string]$Key)
    foreach ($register in @($Manifest.registers)) {
        if (([string]$register.key) -ceq $Key) { return $register }
    }
    throw ('the Stage-B manifest projects no ' + [char]39 + $Key + [char]39 + ' register')
}

function Get-W8RegisterColumnIndex {
    param($Register, [string]$ColumnKey)
    $ordinal = [array]::IndexOf(@($Register.columns), $ColumnKey) + 1
    if ($ordinal -lt 1) {
        throw ('the register carries no ' + [char]39 + $ColumnKey + [char]39 + ' column')
    }
    return $ordinal
}

# THE ROW IS FOUND BY THE PERMANENT ID THE CORPUS NAMES, at its PHYSICAL
# position in the table body. A filtered id list would renumber the rows the
# moment a blank one appeared above the driver being edited.
function Get-W8RegisterRowIndex {
    param($Workbook, $Register, [string]$PermanentId)
    $body = @(Get-TableBody -Workbook $Workbook -SheetName ([string]$Register.sheet) `
        -TableName ([string]$Register.table_name))
    for ($index = 0; $index -lt $body.Count; $index++) {
        $cells = @($body[$index])
        if ($cells.Count -lt 1) { continue }
        if (([string]$cells[0]) -ceq $PermanentId) { return ($index + 1) }
    }
    return 0
}

# ===========================================================================
# WHAT A REFUSAL MUST HAVE LEFT ALONE
# ===========================================================================
# PHYSICAL PERSISTENCE, PROVED AS PHYSICAL PERSISTENCE. The stored payload of
# the bank that published is compared against the baseline capture value for
# value and type for type: the whole simulation identity block, the whole annual
# stamp, every index cell, every rung of both ladders for every published year,
# both profile columns, and every published iteration value. No accessor is
# consulted here - that is a separate question with a separate answer, and
# proving one from the other is what W8 must not do.
function Invoke-W8PreservationChecks {
    param($Workbook, $Inspection, $P7, $Baseline, [string]$Stage, [int]$Iterations, $Ledger)
    $capture = Get-W8BankCapture -Workbook $Workbook -Inspection $Inspection -P7 $P7 `
        -Bank ([string]$Baseline.Bank) -YearCount ([int]$Baseline.YearCount)
    $moved = Compare-W8Surface -Before $Baseline.Capture -After $capture
    $null = Add-W8Check ($Stage + ': the published bank ' + [string]$Baseline.Bank +
                         ' payload is value-identical to the successful baseline') `
        ($moved.Moved -eq 0) $moved.Detail

    $grid = Get-W8IterationBlock -Workbook $Workbook -Inspection $Inspection `
        -Bank ([string]$Baseline.Bank) -Count $Iterations -Ledger $Ledger
    $gridMoved = Compare-W8IterationGrid -Before $Baseline.Iterations -After $grid.Values
    $null = Add-W8Check ($Stage + ': no published iteration value in bank ' +
                         [string]$Baseline.Bank + ' changed') `
        ([string]::IsNullOrWhiteSpace($gridMoved)) $gridMoved

    # NOTHING APPEARED IN THE OTHER BANK EITHER. A refusal that quietly published
    # somewhere else would leave the baseline intact and still be a new run.
    $otherCapture = Get-W8BankCapture -Workbook $Workbook -Inspection $Inspection -P7 $P7 `
        -Bank ([string]$Baseline.OtherBank) -YearCount ([int]$Baseline.YearCount)
    $otherMoved = Compare-W8Surface -Before $Baseline.OtherCapture -After $otherCapture
    $null = Add-W8Check ($Stage + ': bank ' + [string]$Baseline.OtherBank +
                         ' is value-identical to the successful baseline') `
        ($otherMoved.Moved -eq 0) $otherMoved.Detail
    $otherStamp = Get-W8AnnualStamp -Workbook $Workbook -Inspection $Inspection -P7 $P7 `
        -Bank ([string]$Baseline.OtherBank)
    $null = Add-W8Check ($Stage + ': no annual publication marker appeared in bank ' +
                         [string]$Baseline.OtherBank) `
        (-not (Test-SimExactText -Actual $otherStamp['published'] `
            -Expected ([string]$P7.annual_records.stamp.published_marker))) `
        ('published = ' + (Format-SimValue $otherStamp['published']))
    return [int]$grid.Acquired
}

# SEMANTIC AUTHORITY, ASKED SEPARATELY. What the accessors now SAY about the
# payload that is still there - and every answer is required to be a member of
# the vocabulary that owns it and of no other.
function Invoke-W8HandoffChecks {
    param($Excel, $P7, $Vocabulary, $Baseline, [string]$Stage,
          [string]$ExpectedDistribution, [string]$ExpectedProfile)
    $handoff = Get-W8Handoff -Excel $Excel -P7 $P7
    $accessors = @($P7.command_surface.handoff_accessors | ForEach-Object { [string]$_ })
    $distribution = $handoff[$accessors[0]]
    $profile = $handoff[$accessors[1]]

    $null = Add-W8Check ($Stage + ': both annual states are members of their own projected ' +
                         'vocabulary and of no attempt vocabulary') `
        ((Test-W8Member -Value $distribution -List $Vocabulary.DistributionStates) -and
         (Test-W8Member -Value $profile -List $Vocabulary.ProfileStates) -and
         (-not (Test-W8Member -Value $distribution -List $Vocabulary.SimAttempts)) -and
         (-not (Test-W8Member -Value $profile -List $Vocabulary.SimAttempts)) -and
         (-not (Test-W8Member -Value $distribution -List $Vocabulary.CalcAttempts)) -and
         (-not (Test-W8Member -Value $profile -List $Vocabulary.CalcAttempts))) `
        ((Format-SimValue $distribution) + ' / ' + (Format-SimValue $profile))
    $null = Add-W8Check ($Stage + ': the annual distribution reports ' + $ExpectedDistribution) `
        (Test-SimExactText -Actual $distribution -Expected $ExpectedDistribution) `
        (Format-SimValue $distribution)
    $null = Add-W8Check ($Stage + ': the annual profile reports ' + $ExpectedProfile) `
        (Test-SimExactText -Actual $profile -Expected $ExpectedProfile) `
        (Format-SimValue $profile)
    # THE ANSWER IS NOT RETRACTED. The stamp still says which Px the persisted
    # profile is the profile FOR, and how many years it covers; a state that
    # stopped being current does not blank either, and nothing is relabelled.
    $null = Add-W8Check ($Stage + ': the stamped Px and year count are still the baseline' +
                         [char]39 + 's') `
        ((Test-SimSameValue -A $Baseline.Handoff[$accessors[2]] -B $handoff[$accessors[2]]) -and
         (Test-SimSameValue -A $Baseline.Handoff[$accessors[3]] -B $handoff[$accessors[3]])) `
        ((Format-SimValue $handoff[$accessors[2]]) + ' / ' +
         (Format-SimValue $handoff[$accessors[3]]))
    return $handoff
}

# ===========================================================================
# ONE REFUSAL, PROVED WHOLE
# ===========================================================================
# The same probe for both parts, because the two states must cost exactly the
# same: nothing. What differs is only the state word the refusal is expected to
# name, and that word is projected, not typed.
function Invoke-W8RefusalProbe {
    param($Excel, $Workbook, $Inspection, $P7, $Vocabulary, $Baseline, [string]$Stage,
          [string]$ExpectedSimState, [int]$Iterations, $Ledger)

    $invariantsBefore = Get-W8RunInvariants -Workbook $Workbook -Inspection $Inspection
    $metadataBefore = Get-W8StatusMetadata -Workbook $Workbook -Inspection $Inspection

    $announcement = Invoke-W8Endpoint -Excel $Excel `
        -Endpoint ([string]$P7.command_surface.annual_endpoint)
    $null = Add-W8Check ($Stage + ': ' + [string]$P7.command_surface.annual_endpoint +
                         ' REFUSED') ($announcement -like 'FAIL|*') $announcement
    # A REFUSAL FOR THE RIGHT REASON. Production names the state it refused for
    # in its own detail, so the runner requires the projected word to be in it
    # rather than accepting any failure at all as evidence.
    $null = Add-W8Check ($Stage + ': the refusal names the ' + $ExpectedSimState +
                         ' simulation it refused for') `
        ($announcement -clike ('*' + $ExpectedSimState + '*')) $announcement

    # NOTHING THE SUCCESSFUL PUBLICATION OWNS MOVED - identity, counters, the
    # pending marker, the active bank, and the simulation attempt rows.
    $null = Add-W8InvariantChecks ($Stage + ': the refusal') $invariantsBefore `
        (Get-W8RunInvariants -Workbook $Workbook -Inspection $Inspection)
    $null = Add-W8InvariantChecks ($Stage + ': the refusal, measured from the baseline') `
        $Baseline.Invariants (Get-W8RunInvariants -Workbook $Workbook -Inspection $Inspection)

    # THE ATTEMPT AND STATUS ROWS, INSPECTED RATHER THAN FROZEN WHOLESALE.
    $metadataAfter = Get-W8StatusMetadata -Workbook $Workbook -Inspection $Inspection
    $attemptMoved = @()
    foreach ($key in $metadataBefore.Keys) {
        if (-not ([string]$key).StartsWith('attempt.')) { continue }
        if (-not (Test-SimSameValue -A $metadataBefore[$key] -B $metadataAfter[$key])) {
            $attemptMoved += ([string]$key + ': ' + (Format-SimValue $metadataBefore[$key]) +
                              ' -> ' + (Format-SimValue $metadataAfter[$key]))
        }
    }
    $null = Add-W8Check ($Stage + ': the refusal wrote no SIMULATION attempt row') `
        ($attemptMoved.Count -eq 0) ($attemptMoved -join '; ')

    # THE STATUS ROW IS FOUND BY WHAT IT HOLDS, and the timestamp beside it is
    # allowed to move because evaluating the status is what the annual endpoint's
    # own precondition does.
    $statusRows = @(Get-W8RowsHolding -Block $metadataAfter -List $Vocabulary.SimStates `
        -Prefix 'derived.')
    $null = Add-W8Check ($Stage + ': exactly one derived row carries a simulation state word') `
        ($statusRows.Count -eq 1) (($statusRows -join ', ') + ' of ' +
                                   [string]@($metadataAfter.Keys).Count + ' inspected rows')
    $statusValue = $null
    if ($statusRows.Count -eq 1) { $statusValue = $metadataAfter[$statusRows[0]] }
    $null = Add-W8Check ($Stage + ': the persisted simulation status is still ' +
                         $ExpectedSimState) `
        (Test-SimExactText -Actual $statusValue -Expected $ExpectedSimState) `
        (Format-SimValue $statusValue)
    $null = Add-W8Check ($Stage + ': no persisted status or attempt row answers with an ' +
                         'attempt word where a state word belongs') `
        ((@(Get-W8RowsHolding -Block $metadataAfter -List $Vocabulary.SimAttempts `
            -Prefix 'derived.')).Count -eq 0) `
        ((@(Get-W8RowsHolding -Block $metadataAfter -List $Vocabulary.SimAttempts `
            -Prefix 'derived.')) -join ', ')
    foreach ($key in $metadataAfter.Keys) {
        Write-W8Line ('    ' + $Stage + ' ' + ([string]$key).PadRight(34) +
                      (Format-SimValue $metadataBefore[$key]) + ' -> ' +
                      (Format-SimValue $metadataAfter[$key]))
    }

    # AND THE PAYLOAD IS STILL THERE.
    return (Invoke-W8PreservationChecks -Workbook $Workbook -Inspection $Inspection -P7 $P7 `
        -Baseline $Baseline -Stage ($Stage + ' after the refusal') -Iterations $Iterations `
        -Ledger $Ledger)
}

# ===========================================================================
# PREFLIGHT
# ===========================================================================
Write-Host ''
Write-Host 'PCCM - Phase 7 W8 (the two live refusals and what survives them)' -ForegroundColor Cyan
Write-Host '================================================================' -ForegroundColor Cyan
Write-Host ''

$manifestPath   = Join-Path $BuildDir 'stage_b_manifest.json'
$inspectPath    = Join-Path $BuildDir 'phase5_gate_b_inspection.json'
$simInspectPath = Join-Path $BuildDir 'phase6_gate_b_inspection.json'
$gateBCasePath  = Join-Path $BuildDir 'phase6_gate_b_cases.json'
$p7InspectPath  = Join-Path $BuildDir 'phase7_acceptance_inspection.json'
$casesPath      = Join-Path $BuildDir 'phase7_acceptance_cases.json'
foreach ($required in @($manifestPath, $inspectPath, $simInspectPath, $gateBCasePath,
                        $p7InspectPath, $casesPath)) {
    if (-not (Test-Path -LiteralPath $required)) {
        Write-Host ($required + ' not found. Run the Stage-A build first: ' +
                    'python3 pccm/builder/build_stage_a.py') -ForegroundColor Red
        exit 1
    }
}
$manifest      = Get-Content -LiteralPath $manifestPath   -Raw | ConvertFrom-Json
$inspection    = Get-Content -LiteralPath $inspectPath    -Raw | ConvertFrom-Json
$simInspection = Get-Content -LiteralPath $simInspectPath -Raw | ConvertFrom-Json
$gateBCases    = Get-Content -LiteralPath $gateBCasePath  -Raw | ConvertFrom-Json
$p7            = Get-Content -LiteralPath $p7InspectPath  -Raw | ConvertFrom-Json
$cases         = Get-Content -LiteralPath $casesPath      -Raw | ConvertFrom-Json

# THE BEHAVIOURAL FIXTURE W4 ESTABLISHED AND W5 AND W6 BOUND TO. W8 introduces
# no new model, no new seed and no new selector: the baseline it refuses to
# damage has to be the one the accepted rounds already reconciled.
$case = $cases.scenarios | Where-Object { [string]$_.id -ceq 'W4' }
if ($null -eq $case) {
    Write-Host 'the acceptance corpus carries no behavioural scenario for W8 to reuse.' -ForegroundColor Red
    exit 1
}

$revision = $null
try { $revision = Get-W8SourceRevision -RepoRoot $repoRoot }
catch { Write-Host (Format-Err $_) -ForegroundColor Red; exit 1 }
if ($revision.Dirty.Count -gt 0) {
    Write-Host 'REFUSED, BEFORE EXCEL WAS STARTED.' -ForegroundColor Red
    Write-Host ''
    Write-Host ('pccm/src, pccm/spec or pccm/builder is modified, so a W8 result could ' +
                'not be attributed to a source revision:') -ForegroundColor Red
    foreach ($line in $revision.Dirty) { Write-Host ('    ' + $line) -ForegroundColor Red }
    exit 1
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) `
    ('pccm-phase7-w8-' + (Get-Date).ToString('yyyyMMdd-HHmmss'))
$null = New-Item -ItemType Directory -Path $tempRoot -Force
Copy-Item -LiteralPath (Join-Path $BuildDir ([string]$manifest.stage_a_filename)) -Destination $tempRoot
foreach ($artefact in @($manifestPath, $inspectPath, $simInspectPath, $gateBCasePath,
                        $p7InspectPath, $casesPath)) {
    Copy-Item -LiteralPath $artefact -Destination $tempRoot
}
Copy-Item -LiteralPath (Join-Path $BuildDir 'vba') -Destination $tempRoot -Recurse

$script:W8Path = Join-Path $tempRoot 'phase7_w8_refusal.txt'
$stageBPath = Join-Path $tempRoot ([string]$manifest.stage_b_filename)

$bootstrap = Join-Path $scriptDir 'build_stage_b.ps1'
& $bootstrap -BuildDir $tempRoot -Force
$bootstrapExit = $LASTEXITCODE
$bootstrapOk = (($bootstrapExit -eq 0) -and (Test-Path -LiteralPath $stageBPath))

$model = $case.model
$yearCount = [int]$model.timeline.duration
$driverCount = @($model.cost_lines).Count + @($model.risks).Count
$baselineIterations = [int]$case.iterations
$suppliedSeed = [double]$case.supplied_seed
$selectedLabel = [string]$case.selected_confidence_level
$quantileLabels = @($gateBCases.vocabulary.quantile_labels | ForEach-Object { [string]$_ })
$selectedIndex = [array]::IndexOf($quantileLabels, $selectedLabel)
$selectedProbability = 0.0
if ($selectedLabel -match '^P(\d+)$') { $selectedProbability = [double]$Matches[1] / 100.0 }
$measures = @($p7.summary_semantics.contingency_measures | ForEach-Object { [string]$_ })
$vocabulary = Get-W8Vocabulary -GateBCases $gateBCases -P7 $p7

# THE SECOND VALID ITERATION COUNT, DERIVED FROM THE PROJECTED BOUNDS. It is a
# request the workbook would happily run; W8 never runs it, and that is the
# point - a stale simulation is a published run that no longer answers the
# question being asked, not a broken one.
$minimumIterations = [int]$gateBCases.bounds.business_minimum_iterations
$ceilingIterations = [int]$gateBCases.bounds.max_iterations_representable
$staleIterations = $minimumIterations + 1

# THE INVALIDATING EDIT, DERIVED FROM THE CORPUS DRIVER IT IS APPLIED TO. Its
# maximum is put one below its own minimum, which is the ordering every
# distribution family the accepted numerical checker knows refuses.
$costRegister = Get-W8Register -Manifest $manifest -Key 'cost_lines'
$victim = @($model.cost_lines)[0]
$victimId = [string]$victim.permanent_id
$victimMinimum = [double]$victim.min_value
$victimMaximum = [double]$victim.max_value
$invalidMaximum = $victimMinimum - 1.0

Write-W8Line 'PCCM - PHASE 7 W8: THE TWO LIVE REFUSALS AND WHAT SURVIVES THEM'
Write-W8Line '=============================================================='
Write-W8Line ''
Write-W8Line 'This is the MINIMAL W8 runner. One successful FIXED-seed annual baseline,'
Write-W8Line 'then two refusals reached through ordinary inputs: the simulation request'
Write-W8Line 'moves while Phase 5 stays current, and then the model itself stops being'
Write-W8Line 'valid. No selector move (W6), no bank cycling or shrink (W7), no'
Write-W8Line 'sensitivity, and not one hidden cell written by this runner.'
Write-W8Line ''
Write-W8Line ('run started            : ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
Write-W8Line ('host                   : ' + [string]$env:COMPUTERNAME)
Write-W8Line ('PowerShell             : ' + [string]$PSVersionTable.PSVersion)
Write-W8Line ('git HEAD               : ' + [string]$revision.Head)
Write-W8Line  'pccm/src, pccm/spec, pccm/builder : clean (proved before Excel was started)'
Write-W8Line ('model version          : ' + [string]$manifest.model_version)
Write-W8Line ('sim contract version   : ' + [string]$p7.provenance.sim_contract_version)
Write-W8Line ('build directory        : ' + $BuildDir)
Write-W8Line ('working copy           : ' + $tempRoot)
$artefacts = @(
    [pscustomobject]@{ Label = 'Stage-A workbook'
                       Path = (Join-Path $tempRoot ([string]$manifest.stage_a_filename)) },
    [pscustomobject]@{ Label = 'Stage-B workbook'; Path = $stageBPath }
)
foreach ($artefact in $artefacts) {
    $shown = '(not present)'
    if (Test-Path -LiteralPath ([string]$artefact.Path)) {
        try { $shown = [string](Get-FileHash -LiteralPath ([string]$artefact.Path) -Algorithm SHA256).Hash }
        catch { $shown = '(unreadable)' }
    }
    Write-W8Line ('  ' + ([string]$artefact.Label).PadRight(20) + ' SHA-256 ' + $shown)
}
Write-W8Line ''
Write-W8Line 'THE BASELINE AND THE TWO MOVES'
Write-W8Line '------------------------------'
Write-W8Line ('  baseline             : ' + [string]$driverCount + ' drivers, ' +
              [string]$yearCount + ' project years, FIXED seed ' + [string]$suppliedSeed +
              ', ' + [string]$baselineIterations + ' iterations, selected ' + $selectedLabel)
Write-W8Line ('  part A               : iterations ' + [string]$baselineIterations + ' -> ' +
              [string]$staleIterations + ', a valid request the workbook would run')
Write-W8Line ('  part B               : cost line ' + $victimId + ' maximum ' +
              [string]$victimMaximum + ' -> ' + [string]$invalidMaximum + ', below its own ' +
              'minimum ' + [string]$victimMinimum)
Write-W8Line ('  identity rule        : ' + [string]$cases.provenance.identity_rule)
Write-W8Line ''
Write-W8Line 'THE VOCABULARIES, PROJECTED. Not one member is spelled in the runner.'
Write-W8Line ('  simulation states    : ' + (@($vocabulary.SimStates) -join ', '))
Write-W8Line ('  simulation attempts  : ' + (@($vocabulary.SimAttempts) -join ', '))
Write-W8Line ('  model states         : ' + (@($vocabulary.CalcStates) -join ', '))
Write-W8Line ('  model attempts       : ' + (@($vocabulary.CalcAttempts) -join ', '))
Write-W8Line ('  annual distribution  : ' + (@($vocabulary.DistributionStates) -join ', '))
Write-W8Line ('  annual profile       : ' + (@($vocabulary.ProfileStates) -join ', '))
Write-W8Line ''
Write-W8Line 'THERE IS NO INDEPENDENT ANNUAL WINDOWS ORACLE, and none is claimed. The'
Write-W8Line 'baseline is reconciled the way W5 and W6 reconcile it, so what the refusals'
Write-W8Line 'preserve is a real successful answer and not only a storage shape. The two'
Write-W8Line 'refusals produce nothing, so there is nothing after them to reconcile.'
Write-W8Line ''

$null = Add-W8Check 'the Stage-B workbook was generated from the current Stage-A build' `
    $bootstrapOk ('bootstrap exit ' + [string]$bootstrapExit) 'PREREQUISITE'
if (-not $bootstrapOk) {
    Write-W8Line ''
    Write-W8Line 'STOP. Excel was never started for the W8 session, and nothing was accepted.'
    Write-W8Line ('report                 : ' + $script:W8Path)
    Write-W8Line 'W8: FAIL'
    exit 1
}

$null = Add-W8Check 'the W8 baseline is the accepted behavioural fixture' `
    (($driverCount -eq 5) -and ($yearCount -eq 4) -and (([string]$case.seed_mode) -ceq 'FIXED') -and
     ($baselineIterations -eq $minimumIterations) -and ($selectedIndex -ge 0)) `
    ([string]$driverCount + ' drivers, ' + [string]$yearCount + ' years, ' +
     [string]$case.seed_mode + ' seed, ' + [string]$baselineIterations +
     ' iterations against a projected business minimum of ' + [string]$minimumIterations +
     ', ' + $selectedLabel + ' at ladder index ' + [string]$selectedIndex) 'PREREQUISITE'
$null = Add-W8Check 'the stale request is another VALID iteration count, not a broken one' `
    (($staleIterations -ne $baselineIterations) -and ($staleIterations -ge $minimumIterations) -and
     ($staleIterations -le $ceilingIterations)) `
    ([string]$staleIterations + ' within the projected [' + [string]$minimumIterations + ', ' +
     [string]$ceilingIterations + ']') 'PREREQUISITE'
# The ordinal is resolved HERE rather than inside the check expression, because
# a missing column throws and a throw inside a check would end the run without
# the report ever saying which column the register stopped carrying.
$maximumOrdinal = 0
$ordinalFailure = ''
try { $maximumOrdinal = Get-W8RegisterColumnIndex -Register $costRegister -ColumnKey 'unit_cost_max' }
catch { $ordinalFailure = (Format-Err $_) }
$null = Add-W8Check 'the invalidating edit puts a cost line maximum below its own minimum' `
    (($invalidMaximum -lt $victimMinimum) -and ($invalidMaximum -ne $victimMaximum) -and
     ($maximumOrdinal -ge 1)) `
    ($victimId + ': min ' + [string]$victimMinimum + ', max ' + [string]$victimMaximum +
     ' -> ' + [string]$invalidMaximum + ' at register ordinal ' + [string]$maximumOrdinal +
     ' ' + $ordinalFailure) 'PREREQUISITE'

# ---- THE VOCABULARIES ARE THE SHAPE THE ORDINALS ASSUME ------------------
$null = Add-W8Check 'every projected vocabulary has the shape its ordinals are read at' `
    ((@($vocabulary.SimStates).Count -eq 3) -and (@($vocabulary.CalcStates).Count -eq 4) -and
     (@($vocabulary.CalcAttempts).Count -ge 3) -and
     (@($vocabulary.DistributionStates).Count -eq 3) -and
     (@($vocabulary.ProfileStates).Count -eq 4)) `
    ('simulation ' + [string]@($vocabulary.SimStates).Count + ', model ' +
     [string]@($vocabulary.CalcStates).Count + ', model attempts ' +
     [string]@($vocabulary.CalcAttempts).Count + ', distribution ' +
     [string]@($vocabulary.DistributionStates).Count + ', profile ' +
     [string]@($vocabulary.ProfileStates).Count) 'PREREQUISITE'
$null = Add-W8Check ('the superseded annual state is one word, declared by both handoff ' +
                     'lists, and is none of the other three') `
    ((([string]$vocabulary.DistributionSuperseded) -ceq ([string]$vocabulary.ProfileSuperseded)) -and
     (([string]$vocabulary.DistributionSuperseded) -cne ([string]$vocabulary.AnnualCurrent)) -and
     (([string]$vocabulary.ProfileSuperseded) -cne ([string]$vocabulary.AnnualProfileCurrent)) -and
     (([string]$vocabulary.DistributionSuperseded) -cne ([string]$vocabulary.AnnualNotProduced)) -and
     (([string]$vocabulary.DistributionSuperseded) -cne ([string]$vocabulary.InconsistentStamp))) `
    ([string]$vocabulary.DistributionSuperseded + ' / ' + [string]$vocabulary.ProfileSuperseded +
     ', against ' + [string]$vocabulary.AnnualCurrent + ', ' +
     [string]$vocabulary.AnnualNotProduced + ', ' + [string]$vocabulary.InconsistentStamp) 'PREREQUISITE'

# ---- A REFUSED ATTEMPT IS NOT A PERSISTENT STATE, IN THE PROJECTION -------
# The static half of the distinction. Neither engine's persistent vocabulary may
# contain a word from either engine's attempt vocabulary, and the annual handoff
# may not contain one either. If any of these overlapped, every runtime check
# below would be unable to tell an attempt outcome from a state.
$overlaps = New-Object System.Collections.ArrayList
foreach ($pair in @(
    [pscustomobject]@{ Name = 'simulation states vs simulation attempts'
                       A = $vocabulary.SimStates;          B = $vocabulary.SimAttempts },
    [pscustomobject]@{ Name = 'model states vs model attempts'
                       A = $vocabulary.CalcStates;         B = $vocabulary.CalcAttempts },
    [pscustomobject]@{ Name = 'annual distribution states vs simulation attempts'
                       A = $vocabulary.DistributionStates; B = $vocabulary.SimAttempts },
    [pscustomobject]@{ Name = 'annual profile states vs simulation attempts'
                       A = $vocabulary.ProfileStates;      B = $vocabulary.SimAttempts },
    [pscustomobject]@{ Name = 'annual distribution states vs model attempts'
                       A = $vocabulary.DistributionStates; B = $vocabulary.CalcAttempts },
    [pscustomobject]@{ Name = 'annual profile states vs model attempts'
                       A = $vocabulary.ProfileStates;      B = $vocabulary.CalcAttempts })) {
    $shared = @(Get-W8SharedMembers -First $pair.A -Second $pair.B)
    if ($shared.Count -gt 0) {
        $null = $overlaps.Add([string]$pair.Name + ': ' + ($shared -join ', '))
    }
}
$null = Add-W8Check ('no persistent state vocabulary shares a member with an attempt ' +
                     'vocabulary') (@($overlaps).Count -eq 0) ((@($overlaps)) -join '; ')
$null = Add-W8Check ('the refusal word W8 tests for is an attempt result and no engine' +
                     [char]39 + 's persistent state') `
    ((Test-W8Member -Value ([string]$vocabulary.CalcRefused) -List $vocabulary.CalcAttempts) -and
     (-not (Test-W8Member -Value ([string]$vocabulary.CalcRefused) -List $vocabulary.CalcStates)) -and
     (-not (Test-W8Member -Value ([string]$vocabulary.CalcRefused) -List $vocabulary.SimStates)) -and
     (-not (Test-W8Member -Value ([string]$vocabulary.CalcRefused) `
        -List $vocabulary.DistributionStates)) -and
     (-not (Test-W8Member -Value ([string]$vocabulary.CalcRefused) -List $vocabulary.ProfileStates))) `
    ([string]$vocabulary.CalcRefused)

# ---- THE FROZEN / MOVABLE PARTITION IS THE PROJECTION'S -------------------
$derivedRows = @(Get-W8RowsInGroup -Inspection $simInspection -Group 'derived')
$attemptRows = @(Get-W8RowsInGroup -Inspection $simInspection -Group 'attempt')
$snapshotRows = @(Get-W8RowsInGroup -Inspection $simInspection -Group 'snapshot')
$null = Add-W8Check ('the run-identity projection separates derived rows from attempt rows ' +
                     'and from the published identity') `
    (($derivedRows.Count -ge 1) -and ($attemptRows.Count -ge 1) -and ($snapshotRows.Count -ge 1) -and
     (@(Get-W8SharedMembers -First $derivedRows -Second $attemptRows).Count -eq 0) -and
     (@(Get-W8SharedMembers -First $derivedRows -Second $snapshotRows).Count -eq 0)) `
    ('derived: ' + ($derivedRows -join ', ') + ' | attempt: ' + ($attemptRows -join ', ') +
     ' | snapshot rows: ' + [string]$snapshotRows.Count) 'PREREQUISITE'

# ===========================================================================
# THE W8 SESSION
# ===========================================================================
$preExisting = @(Get-PreExistingExcelPids)
$excel = $null; $workbooks = $null; $wb = $null
$excelIdentity = $null
$naturalExit = $false
$emergencyRequired = $false
$fatal = ''
$comAcquired = 0

$rel = New-ReleaseLedger 'phase 7 W8 session'

try {
    $excel = New-Object -ComObject Excel.Application
    $comAcquired = $comAcquired + 1
    $excelIdentity = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    Write-W8Line ('EXCEL PROCESS OWNERSHIP: this run created PID ' +
                  [string]$excelIdentity.ProcessId + '. No process it did not create is')
    Write-W8Line 'ever terminated, and the workbook is never saved.'
    Write-W8Line ''
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $workbooks = $excel.Workbooks
    $comAcquired = $comAcquired + 1
    $wb = $workbooks.Open($stageBPath)
    $comAcquired = $comAcquired + 1

    $compileFailure = ''
    try { $null = $excel.Run('PCCM_CalculationStatus') } catch { $compileFailure = (Format-Err $_) }
    $null = Add-W8Check 'the current VBAProject compiles in real Excel' `
        ([string]::IsNullOrWhiteSpace($compileFailure)) $compileFailure 'PREREQUISITE'
    if (-not [string]::IsNullOrWhiteSpace($compileFailure)) { throw $compileFailure }
    $null = Save-Phase5LockedFxSeed -Workbook $wb -Inspection $inspection

    # ===================================================================
    # THE BASELINE - ONE SUCCESSFUL ANNUAL RUN
    # ===================================================================
    Write-W8Line ''
    Write-W8Line 'THE BASELINE - ONE SUCCESSFUL ANNUAL RUN'
    Write-W8Line '----------------------------------------'
    $applied = ''
    $fixtureFailure = ''
    try {
        $applied = [string](Set-Phase5Fixture -Excel $excel -Workbook $wb -Manifest $manifest `
            -Inspection $inspection -Model $model)
        Set-NamedValue -Workbook $wb `
            -DefinedName ([string]$simInspection.controls.monte_carlo_iterations.defined_name) `
            -Value ([double]$baselineIterations)
        Set-NamedValue -Workbook $wb `
            -DefinedName ([string]$simInspection.controls.random_seed.defined_name) `
            -Value $suppliedSeed
        Set-W8NamedText -Workbook $wb `
            -DefinedName ([string]$inspection.inputs.($p7.selector_semantics.selector_input_key).defined_name) `
            -Value $selectedLabel
    } catch { $fixtureFailure = (Format-Err $_) }
    $null = Add-W8Check 'the behavioural model and the simulation request were applied' `
        (($applied -like 'OK|*') -and [string]::IsNullOrWhiteSpace($fixtureFailure)) `
        ($applied + $fixtureFailure) 'PREREQUISITE'
    if (-not ($applied -like 'OK|*')) { throw ('the W8 model was not applied: ' + $applied) }

    $calcFailure = ''
    try {
        $null = [string](Invoke-Phase5ProductionOperation -Excel $excel `
            -Operation 'PCCM_Calculate' -Stage 'W8 baseline calculate')
    } catch { $calcFailure = (Format-Err $_) }
    $null = Add-W8Check 'PCCM_Calculate succeeded' `
        ([string]::IsNullOrWhiteSpace($calcFailure)) $calcFailure 'PREREQUISITE'
    if (-not [string]::IsNullOrWhiteSpace($calcFailure)) { throw $calcFailure }
    $null = Add-W8Check 'the baseline calculation reports the projected current model state' `
        ((([string]$excel.Run('PCCM_CalculationStatus')) -ceq ([string]$vocabulary.CalcCurrent))) `
        ([string]$excel.Run('PCCM_CalculationStatus')) 'PREREQUISITE'

    $stateStart = Get-Phase6State -Workbook $wb -Inspection $simInspection
    $expectedBank = Get-W8CandidateBank -P7 $p7 `
        -ActiveBank (Get-Phase6ActiveBank -State $stateStart)
    $nonceStart = $stateStart['shared']['next_auto_nonce']
    $pendingStart = $stateStart['pending_auto_nonce']

    $simResult = Invoke-Phase6Simulation -Excel $excel
    $null = Add-W8Check 'PCCM_RunSimulation succeeded' `
        (Test-Phase6Announced -Result $simResult -Kind 'OK') $simResult 'PREREQUISITE'
    if (-not (Test-Phase6Announced -Result $simResult -Kind 'OK')) {
        throw 'the W8 baseline simulation did not succeed'
    }
    $null = Add-W8Check 'the baseline simulation reports the projected current state' `
        ((([string]$excel.Run('PCCM_SimulationStatus')) -ceq ([string]$vocabulary.SimCurrent))) `
        ([string]$excel.Run('PCCM_SimulationStatus')) 'PREREQUISITE'

    $annual = Invoke-W8Endpoint -Excel $excel `
        -Endpoint ([string]$p7.command_surface.annual_endpoint)
    $null = Add-W8Check ([string]$p7.command_surface.annual_endpoint + ' succeeded') `
        ($annual -like 'OK|*') $annual 'PREREQUISITE'
    if (-not ($annual -like 'OK|*')) { throw ('the W8 baseline annual run did not succeed: ' + $annual) }

    # ---- THE BASELINE, CAPTURED WHOLE AND AS PLAIN DATA ----------------
    $stateBaseline = Get-Phase6State -Workbook $wb -Inspection $simInspection
    $bank = Get-Phase6ActiveBank -State $stateBaseline
    $otherBank = Get-W8OtherBank -P7 $p7 -Bank $bank
    $null = Add-W8Check ('the baseline published to the bank the projected cycle names (' +
                         $expectedBank + ')') `
        (($bank -ceq $expectedBank) -and (-not [string]::IsNullOrEmpty($otherBank)) -and
         ($otherBank -cne $bank)) ($bank + ' vs ' + $expectedBank + ', other ' + $otherBank) `
        'PREREQUISITE'
    $stampBaseline = Get-W8AnnualStamp -Workbook $wb -Inspection $simInspection -P7 $p7 -Bank $bank
    $authority = Test-W8Authoritative -Stamp $stampBaseline -P7 $p7 -ExpectedYears $yearCount
    $null = Add-W8Check ('the baseline is authoritative for exactly ' + [string]$yearCount +
                         ' years') ([bool]$authority.Ok) ([string]$authority.Detail)
    $captureBaseline = Get-W8BankCapture -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bank -YearCount $yearCount
    $otherCaptureBaseline = Get-W8BankCapture -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $otherBank -YearCount $yearCount
    $iterationsBaseline = Get-W8IterationBlock -Workbook $wb -Inspection $simInspection `
        -Bank $bank -Count $baselineIterations -Ledger $rel
    $comAcquired = $comAcquired + [int]$iterationsBaseline.Acquired
    $handoffBaseline = Get-W8Handoff -Excel $excel -P7 $p7
    $accessors = @($p7.command_surface.handoff_accessors | ForEach-Object { [string]$_ })
    $null = Add-W8Check 'the baseline reports both annual products current and the right year count' `
        (((([string]$handoffBaseline[$accessors[0]]) -ceq ([string]$vocabulary.AnnualCurrent))) -and
         ((([string]$handoffBaseline[$accessors[1]]) -ceq ([string]$vocabulary.AnnualProfileCurrent))) -and
         ((Test-SimExactText -Actual $handoffBaseline[$accessors[2]] -Expected $selectedLabel)) -and
         (([int]$handoffBaseline[$accessors[3]]) -eq $yearCount)) `
        ((Format-SimValue $handoffBaseline[$accessors[0]]) + ' / ' +
         (Format-SimValue $handoffBaseline[$accessors[1]]) + ' / ' +
         (Format-SimValue $handoffBaseline[$accessors[2]]) + ' / ' +
         (Format-SimValue $handoffBaseline[$accessors[3]]))
    $profileBaseline = Get-W8ProfileSums -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bank -YearCount $yearCount -Measures $measures
    $null = Add-W8Check 'every baseline profile value is a number' `
        (@($profileBaseline.Problems).Count -eq 0) ((@($profileBaseline.Problems)) -join '; ')
    Invoke-W8Reconciliation -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Provenance $cases.provenance -Grid $iterationsBaseline.Values -Bank $bank `
        -Label $selectedLabel -Probability $selectedProbability -LadderIndex $selectedIndex `
        -ProfileSums $profileBaseline.Sums -ProfileScale $profileBaseline.Scale -Stage 'baseline'

    $baseline = [pscustomobject]@{
        Bank         = $bank
        OtherBank    = $otherBank
        YearCount    = $yearCount
        Capture      = $captureBaseline
        OtherCapture = $otherCaptureBaseline
        Iterations   = $iterationsBaseline.Values
        Handoff      = $handoffBaseline
        Invariants   = (Get-W8RunInvariants -Workbook $wb -Inspection $simInspection)
    }

    Write-W8Line ''
    Write-W8Line 'THE BASELINE, AS PLAIN DATA'
    Write-W8Line '---------------------------'
    Write-W8Line (Format-Phase6State -State $stateBaseline -Label '  the run identity')
    Write-W8Line ('  active bank = ' + $bank + ', other bank = ' + $otherBank)
    foreach ($key in $stampBaseline.Keys) {
        Write-W8Line ('    annual stamp.' + ([string]$key).PadRight(26) +
                      (Format-SimValue $stampBaseline[$key]))
    }
    for ($offset = 0; $offset -lt $yearCount; $offset++) {
        $record = Get-W8AnnualRecord -Workbook $wb -Inspection $simInspection -P7 $p7 `
            -Bank $bank -Offset $offset
        foreach ($key in $record.Keys) {
            $name = [string]$key
            $shown = ''
            if ($name.StartsWith('ladder_')) {
                $parts = @()
                foreach ($rung in @($record[$name])) { $parts += (Format-SimValue $rung) }
                $shown = ($parts -join ' ')
            } else {
                $shown = (Format-SimValue $record[$name])
            }
            Write-W8Line ('    year ' + [string]($offset + 1) + ' ' + $name.PadRight(20) + $shown)
        }
    }
    foreach ($accessor in $accessors) {
        Write-W8Line ('    ' + $accessor.PadRight(34) + (Format-SimValue $handoffBaseline[$accessor]))
    }
    $grid = $iterationsBaseline.Values
    Write-W8Line ('    iteration block ' + [string]$grid.GetLength(0) + ' x ' +
                  [string]$grid.GetLength(1) + ', first ' +
                  (Format-SimValue $grid[1, 2]) + ' / ' + (Format-SimValue $grid[1, 3]) +
                  ', last ' + (Format-SimValue $grid[$grid.GetLength(0), 2]) + ' / ' +
                  (Format-SimValue $grid[$grid.GetLength(0), 3]))

    # ===================================================================
    # PART A - THE SIMULATION REQUEST MOVES AND PHASE 5 DOES NOT
    # ===================================================================
    Write-W8Line ''
    Write-W8Line 'PART A - A VALID REQUEST CHANGE MAKES THE PUBLISHED RUN STALE'
    Write-W8Line '------------------------------------------------------------'
    $iterationsControl = [string]$simInspection.controls.monte_carlo_iterations.defined_name
    $iterationsBefore = Get-NamedValue -Workbook $wb -DefinedName $iterationsControl
    Set-NamedValue -Workbook $wb -DefinedName $iterationsControl -Value ([double]$staleIterations)
    $iterationsAfter = Get-NamedValue -Workbook $wb -DefinedName $iterationsControl
    $null = Add-W8Check ('the projected iteration control now carries the second valid value') `
        ((([string]$iterationsAfter) -ceq ([string]$staleIterations)) -and
         (([string]$iterationsBefore) -cne ([string]$iterationsAfter))) `
        ($iterationsControl + ': ' + [string]$iterationsBefore + ' -> ' +
         [string]$iterationsAfter) 'PREREQUISITE'

    # PHASE 5 IS UNTOUCHED. The iteration count enters the SIMULATION request
    # fingerprint and no calculation, so the model must still be current - and
    # if it were not, part A would be testing part B's state under part A's name.
    $null = Add-W8Check 'part A: the model is still the projected current state' `
        ((([string]$excel.Run('PCCM_CalculationStatus')) -ceq ([string]$vocabulary.CalcCurrent))) `
        ([string]$excel.Run('PCCM_CalculationStatus'))
    $null = Add-W8Check 'part A: the simulation reports the projected stale state' `
        ((([string]$excel.Run('PCCM_SimulationStatus')) -ceq ([string]$vocabulary.SimStale))) `
        ([string]$excel.Run('PCCM_SimulationStatus'))
    $null = Add-W8InvariantChecks 'part A: the request change' $baseline.Invariants `
        (Get-W8RunInvariants -Workbook $wb -Inspection $simInspection)
    $comAcquired = $comAcquired + (Invoke-W8PreservationChecks -Workbook $wb `
        -Inspection $simInspection -P7 $p7 -Baseline $baseline `
        -Stage 'part A after the request change' -Iterations $baselineIterations -Ledger $rel)
    $null = Invoke-W8HandoffChecks -Excel $excel -P7 $p7 -Vocabulary $vocabulary `
        -Baseline $baseline -Stage 'part A after the request change' `
        -ExpectedDistribution ([string]$vocabulary.DistributionSuperseded) `
        -ExpectedProfile ([string]$vocabulary.ProfileSuperseded)

    $comAcquired = $comAcquired + (Invoke-W8RefusalProbe -Excel $excel -Workbook $wb `
        -Inspection $simInspection -P7 $p7 -Vocabulary $vocabulary -Baseline $baseline `
        -Stage 'part A' -ExpectedSimState ([string]$vocabulary.SimStale) `
        -Iterations $baselineIterations -Ledger $rel)
    $null = Invoke-W8HandoffChecks -Excel $excel -P7 $p7 -Vocabulary $vocabulary `
        -Baseline $baseline -Stage 'part A after the refusal' `
        -ExpectedDistribution ([string]$vocabulary.DistributionSuperseded) `
        -ExpectedProfile ([string]$vocabulary.ProfileSuperseded)

    # ===================================================================
    # PART B - THE MODEL ITSELF STOPS BEING VALID
    # ===================================================================
    Write-W8Line ''
    Write-W8Line 'PART B - AN ORDINARY REGISTER EDIT MAKES THE MODEL INVALID'
    Write-W8Line '---------------------------------------------------------'
    $victimRow = Get-W8RegisterRowIndex -Workbook $wb -Register $costRegister -PermanentId $victimId
    $null = Add-W8Check ('the cost line the corpus names is in the register') ($victimRow -ge 1) `
        ($victimId + ' at body row ' + [string]$victimRow) 'PREREQUISITE'
    if ($victimRow -lt 1) { throw ('the W8 register edit could not find ' + $victimId) }
    Set-TableCell -Workbook $wb -SheetName ([string]$costRegister.sheet) `
        -TableName ([string]$costRegister.table_name) -RowIndex $victimRow `
        -ColumnIndex $maximumOrdinal -Value $invalidMaximum

    # RECALCULATED THROUGH THE ACCEPTED WORKFLOW, and the workflow refuses. The
    # announcement is the ATTEMPT's answer; the state cell beside it is the
    # MODEL's, and W8 reads both because they are the two axes it exists to keep
    # apart.
    $calcAnnouncement = Invoke-W8Endpoint -Excel $excel -Endpoint 'PCCM_Calculate'
    $null = Add-W8Check 'part B: PCCM_Calculate refuses the edited model' `
        ($calcAnnouncement -like 'FAIL|*') $calcAnnouncement
    $calcState = Get-CalcScalarBlock -Workbook $wb -Inspection $inspection -Block 'calc_state'
    $calcStatusRows = @(Get-W8RowsHolding -Block $calcState -List $vocabulary.CalcStates)
    $calcAttemptRows = @(Get-W8RowsHolding -Block $calcState -List $vocabulary.CalcAttempts)
    $null = Add-W8Check ('part B: the persisted model state carries one state word and one ' +
                         'attempt word, in different rows') `
        (($calcStatusRows.Count -eq 1) -and ($calcAttemptRows.Count -eq 1) -and
         (@(Get-W8SharedMembers -First $calcStatusRows -Second $calcAttemptRows).Count -eq 0)) `
        ('state row ' + ($calcStatusRows -join ', ') + ', attempt row ' +
         ($calcAttemptRows -join ', '))
    $calcStatusValue = $null
    $calcAttemptValue = $null
    if ($calcStatusRows.Count -eq 1) { $calcStatusValue = $calcState[$calcStatusRows[0]] }
    if ($calcAttemptRows.Count -eq 1) { $calcAttemptValue = $calcState[$calcAttemptRows[0]] }
    $null = Add-W8Check ('part B: the refused attempt is recorded as an attempt result and the ' +
                         'model state is a state word, and they are different words') `
        ((Test-SimExactText -Actual $calcAttemptValue -Expected ([string]$vocabulary.CalcRefused)) -and
         (Test-SimExactText -Actual $calcStatusValue -Expected ([string]$vocabulary.CalcInvalid)) -and
         (([string]$vocabulary.CalcRefused) -cne ([string]$vocabulary.CalcInvalid))) `
        ('attempt ' + (Format-SimValue $calcAttemptValue) + ', state ' +
         (Format-SimValue $calcStatusValue))
    $null = Add-W8Check 'part B: the model no longer reports the projected current state' `
        (((([string]$excel.Run('PCCM_CalculationStatus')) -cne ([string]$vocabulary.CalcCurrent))) -and
         ((([string]$excel.Run('PCCM_CalculationStatus')) -ceq ([string]$vocabulary.CalcInvalid)))) `
        ([string]$excel.Run('PCCM_CalculationStatus'))
    $null = Add-W8Check 'part B: the simulation reports the projected invalid state' `
        (((([string]$excel.Run('PCCM_SimulationStatus')) -cne ([string]$vocabulary.SimCurrent))) -and
         ((([string]$excel.Run('PCCM_SimulationStatus')) -ceq ([string]$vocabulary.SimInvalid)))) `
        ([string]$excel.Run('PCCM_SimulationStatus'))
    $null = Add-W8InvariantChecks 'part B: the model edit and the refused recalculation' `
        $baseline.Invariants (Get-W8RunInvariants -Workbook $wb -Inspection $simInspection)
    $comAcquired = $comAcquired + (Invoke-W8PreservationChecks -Workbook $wb `
        -Inspection $simInspection -P7 $p7 -Baseline $baseline `
        -Stage 'part B after the model edit' -Iterations $baselineIterations -Ledger $rel)
    $null = Invoke-W8HandoffChecks -Excel $excel -P7 $p7 -Vocabulary $vocabulary `
        -Baseline $baseline -Stage 'part B after the model edit' `
        -ExpectedDistribution ([string]$vocabulary.DistributionSuperseded) `
        -ExpectedProfile ([string]$vocabulary.ProfileSuperseded)

    $comAcquired = $comAcquired + (Invoke-W8RefusalProbe -Excel $excel -Workbook $wb `
        -Inspection $simInspection -P7 $p7 -Vocabulary $vocabulary -Baseline $baseline `
        -Stage 'part B' -ExpectedSimState ([string]$vocabulary.SimInvalid) `
        -Iterations $baselineIterations -Ledger $rel)
    $null = Invoke-W8HandoffChecks -Excel $excel -P7 $p7 -Vocabulary $vocabulary `
        -Baseline $baseline -Stage 'part B after the refusal' `
        -ExpectedDistribution ([string]$vocabulary.DistributionSuperseded) `
        -ExpectedProfile ([string]$vocabulary.ProfileSuperseded)

    # ---- NOTHING WAS ALLOCATED ACROSS THE WHOLE SESSION ----------------
    $stateEnd = Get-Phase6State -Workbook $wb -Inspection $simInspection
    $null = Add-W8Check 'no refusal consumed a run id, an AUTO nonce or the pending marker' `
        ((Test-SimSameValue -A $nonceStart -B $stateEnd['shared']['next_auto_nonce']) -and
         (Test-SimSameValue -A $pendingStart -B $stateEnd['pending_auto_nonce']) -and
         (Test-SimSameValue -A $stateBaseline['shared']['last_run_id'] `
             -B $stateEnd['shared']['last_run_id']) -and
         (Test-SimSameValue -A $stateBaseline['shared']['active_bank'] `
             -B $stateEnd['shared']['active_bank'])) `
        ('nonce ' + (Format-SimValue $nonceStart) + ' -> ' +
         (Format-SimValue $stateEnd['shared']['next_auto_nonce']) + ', pending ' +
         (Format-SimValue $pendingStart) + ' -> ' +
         (Format-SimValue $stateEnd['pending_auto_nonce']) + ', last run id ' +
         (Format-SimValue $stateEnd['shared']['last_run_id']) + ', active bank ' +
         (Format-SimValue $stateEnd['shared']['active_bank']))

    Write-W8Line ''
    Write-W8Line (Format-Phase6State -State $stateEnd -Label '  the run identity after both refusals')

    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    try { $excel.Run('PCCM_AutomationEnd') | Out-Null } catch { }
} catch {
    $fatal = (Format-Err $_)
    Write-W8Line ''
    Write-W8Line ('THE W8 SESSION DID NOT COMPLETE: ' + $fatal)
} finally {
    try {
        if ($null -ne $wb) {
            try { $wb.Close($false); $rel.WorkbookClosed = $true }
            catch { $null = $rel.Failed.Add('Workbook.Close') }
        }
        Invoke-W8Release $rel $wb        'Workbook';  $wb        = $null
        Invoke-W8Release $rel $workbooks 'Workbooks'; $workbooks = $null
        if ($null -ne $excel) {
            try { $excel.Quit(); $rel.QuitCalled = $true }
            catch { $null = $rel.Failed.Add('Application.Quit') }
        }
        Invoke-W8Release $rel $excel 'Excel.Application'; $excel = $null
    } finally {
        $Error.Clear()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

        if ($null -ne $excelIdentity) {
            $naturalExit = Wait-ExcelExit -Identity $excelIdentity -TimeoutSeconds 90
        }
        $rel.NaturalExit = $naturalExit
        Write-W8Line ''
        Write-W8Line 'EXCEL SHUTDOWN'
        Write-W8Line '--------------'
        if ($naturalExit) {
            Write-W8Line ('EXCEL SHUTDOWN: the owned process (PID ' +
                          [string]$excelIdentity.ProcessId + ') exited naturally.')
        } else {
            $emergencyRequired = $true
            $rel.EmergencyRequired = $true
            $cleaned = Invoke-EmergencyExcelCleanup -Identity $excelIdentity -Label 'W8'
            Write-W8Line ('EXCEL SHUTDOWN: emergency cleanup was required (' + [string]$cleaned + ')')
        }
        Write-W8Line (Format-ReleaseLedger $rel)
        foreach ($residual in @($script:W8Residual)) {
            Write-W8Line ('      OUTSTANDING: ' + [string]$residual)
        }
        foreach ($transient in @(Get-TransientFailures)) {
            Write-W8Line ('      transient release FAILED: ' + [string]$transient)
        }
    }
}

# ===========================================================================
# THE VERDICT
# ===========================================================================
$null = Add-W8Check 'the owned Excel process exited naturally' $naturalExit
$null = Add-W8Check 'no emergency cleanup was required' (-not $emergencyRequired)
$null = Add-W8Check 'every COM object this runner acquired was released' `
    ([int]$rel.Attempted -eq [int]$comAcquired) `
    ([string]$comAcquired + ' acquired, ' + [string]$rel.Attempted + ' released')
$null = Add-W8Check 'every COM release succeeded' ($rel.Failed.Count -eq 0) ($rel.Failed -join ', ')
$null = Add-W8Check 'every COM release left 0 outstanding references' `
    (@($script:W8Residual).Count -eq 0) ((@($script:W8Residual)) -join '; ')
$null = Add-W8Check 'every transient release inside the accepted readers succeeded' `
    (@(Get-TransientFailures).Count -eq 0) ((@(Get-TransientFailures)) -join '; ')

$results = @($script:W8Checks | Where-Object { [string]$_.Kind -eq 'RESULT' })
$prereqs = @($script:W8Checks | Where-Object { [string]$_.Kind -eq 'PREREQUISITE' })
$failedResults = @($results | Where-Object { -not $_.Ok })
$failedPrereqs = @($prereqs | Where-Object { -not $_.Ok })

Write-W8Line ''
Write-W8Line 'VERDICT'
Write-W8Line '-------'
Write-W8Line ('prerequisites          : ' + [string]$prereqs.Count + ' checked, ' +
              [string]$failedPrereqs.Count + ' failed')
Write-W8Line ('scenario results       : ' + [string]$results.Count + ' checked, ' +
              [string]$failedResults.Count + ' failed')
foreach ($failure in @($failedPrereqs + $failedResults)) {
    $suffix = ''
    if (-not [string]::IsNullOrWhiteSpace([string]$failure.Detail)) {
        $suffix = ' -- ' + [string]$failure.Detail
    }
    Write-W8Line ('  FAILED [' + [string]$failure.Kind + '] ' + [string]$failure.Label + $suffix)
}
$ok = (($failedResults.Count -eq 0) -and ($failedPrereqs.Count -eq 0) -and
       [string]::IsNullOrWhiteSpace($fatal) -and ($results.Count -gt 0))
Write-W8Line ''
if ($ok) {
    Write-W8Line 'W8: PASS'
} else {
    Write-W8Line 'W8: FAIL'
    Write-W8Line ''
    Write-W8Line 'STOP AND REVIEW. Do not treat Phase 7 as closed until this is understood.'
}
Write-W8Line ''
Write-W8Line ('report                 : ' + $script:W8Path)
Write-Host ''
Write-Host ('The report is at ' + $script:W8Path) -ForegroundColor Cyan
Write-Host 'The working copy is left in place so the report survives; delete it when done.'
if ($ok) { exit 0 } else { exit 1 }
