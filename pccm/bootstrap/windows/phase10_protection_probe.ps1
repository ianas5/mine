<#
.SYNOPSIS
    PCCM Phase 10 - the PROTECTION / TABLE-STRUCTURE PROBE.

.DESCRIPTION
    ONE QUESTION, ASKED OF REAL EXCEL, IN ABOUT A MINUTE.

    Benchmark Windows Run 3 died on `ListRow.Delete()` with "Table features
    aren't available because the sheet is protected." That call came from
    PowerShell. Static reading of `pccm/src/vba` says the same class of call -
    `ListColumns.Add`, `ListRows.Add`, `ListRows(...).Delete` - sits inside
    SEVEN of the user commands, two of them unconditionally. But "should" is not
    evidence, and the correction it implies edits accepted production modules
    and an accepted contract.

    So this probe asks Excel directly, and separates the two capabilities the
    benchmark run conflated:

      1. CONTROL - can code write a VALUE to a locked cell on a protected sheet?
         Run 3 already suggests yes. This confirms it deliberately.

      2. THE QUESTION - can the real production commands perform their
         CONTRACTED STRUCTURAL operation on a protected sheet?

    AN ANNOUNCEMENT IS NOT ENOUGH. A command that says OK but changed no shape
    has not done its structural work, so every endpoint is measured by what
    happened to the TABLES as well as by what it said.

    THREE OUTCOMES, AND ONLY THREE:

      BLOCKED       a real production endpoint did not succeed while protection
                    was in force.
      FINE          every required endpoint succeeded AND its structural effect
                    actually occurred, with protection in force throughout.
      INCONCLUSIVE  anything else - including any failure of this script. A
                    probe-internal error is NEVER a statement about production.

    IT MEASURES NOTHING. No timing, no baseline, no Gate-B result. It opens a
    disposable Stage-B build, presses real buttons through their real entry
    points, and never saves.

.PARAMETER BuildDir
    The Stage-A build directory. Defaults to <repo>/pccm/build.

.PARAMETER WorkDir
    Where the disposable copy is made. Defaults to the system temp directory.

.NOTES
    SAFETY. No security setting is altered, no registry key is touched, and no
    Excel process this script did not create is ever terminated. Shutdown is the
    accepted `com_lifecycle.ps1` path. The workbook is never saved.
#>

[CmdletBinding()]
param(
    [string]$BuildDir,
    [string]$WorkDir,
    [switch]$KeepArtifacts
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptDir 'com_lifecycle.ps1')

$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $scriptDir))
$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
if ([string]::IsNullOrWhiteSpace($BuildDir)) { $BuildDir = Join-Path $pccmRoot 'build' }
if ([string]::IsNullOrWhiteSpace($WorkDir))  { $WorkDir  = [System.IO.Path]::GetTempPath() }

# ===========================================================================
# WHERE THE PROBE HAD GOT TO
# ===========================================================================
# PROBE RUN 1 REPORTED ONLY THE EXCEPTION. It said "The property 'Count' cannot
# be found on this object" and nothing about which of the six `.Count` reads in
# this file it meant, which stage it was in, or which command it was testing.
# The cursor below is set at every stage boundary and before every endpoint, and
# is read ONLY by the failure path - never inside a decision, never by a
# successful run, so it cannot influence the verdict or touch the workbook.
$script:ProbeCursor = [pscustomobject]@{
    Stage    = 'startup'
    Action   = 'loading the probe'
    Endpoint = ''
    Detail   = ''
}

function Set-ProbeStage {
    param([string]$Stage, [string]$Action = '', [string]$Endpoint = '', [string]$Detail = '')
    $script:ProbeCursor.Stage = $Stage
    $script:ProbeCursor.Action = $Action
    $script:ProbeCursor.Endpoint = $Endpoint
    $script:ProbeCursor.Detail = $Detail
}

$script:ProbeLines = New-Object System.Collections.ArrayList
$script:ProbePath = ''

function Write-ProbeLine {
    param([string]$Text = '')
    $null = $script:ProbeLines.Add($Text)
    Write-Host $Text
    if (-not [string]::IsNullOrWhiteSpace($script:ProbePath)) {
        try {
            Set-Content -LiteralPath $script:ProbePath -Value ($script:ProbeLines -join "`r`n") -Encoding UTF8
        } catch {
            $script:ProbePath = ''
            Write-Host '  (the probe log could not be written through)' -ForegroundColor DarkYellow
        }
    }
}

# ===========================================================================
# SHAPES THAT CANNOT SURPRISE THE PROBE
# ===========================================================================
# PROBE RUN 1 ROOT CAUSE, AS A CLASS RATHER THAN A LINE.
#
# `$sheets.Count` read a member off whatever `$Workbook.Worksheets` handed back.
# Five of the six `.Count` reads in the first draft were wrapped in `@()`, which
# guarantees an array before anything is read from it; that one was not, and it
# is the only one whose operand shape was not guaranteed by construction. Under
# `Set-StrictMode -Version 2.0` a member that is not there is a terminating
# PropertyNotFoundException, so the probe died before it asked its question.
#
# THE SHAPE IS NOW GONE, NOT GUARDED. Nothing in this file indexes a COM
# collection by position or reads `.Count` off one: collections are enumerated,
# and the two accessors below are the only way a property is read off anything
# whose shape is not certain.
function Get-ProbeProperty {
    param($InputObject, [string]$Name)
    if ($null -eq $InputObject) { return $null }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

# A FACT THE PROBE CANNOT DO WITHOUT. Absent is a refusal naming the shape, never
# a zero: a probe that counted nothing and reported zero would answer its own
# question with a fabrication.
function Get-ProbeRequiredProperty {
    param($InputObject, [string]$Name, [string]$Where)
    if ($null -eq $InputObject) {
        throw ($Where + ": expected an object carrying '" + $Name + "' and got nothing")
    }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) {
        throw ($Where + ': the object is a ' + $InputObject.GetType().FullName +
               " and carries no '" + $Name + "' property. It is not defaulted.")
    }
    return $property.Value
}

function Get-ProbeNamedValue {
    param($Workbook, [string]$DefinedName)
    $names = $null; $nm = $null; $range = $null
    try {
        $names = $Workbook.Names
        $nm = $names.Item($DefinedName)
        $range = $nm.RefersToRange
        return $range.Value2
    } finally {
        if ($null -ne $range) { Release-Transient $range 'Range'; $range = $null }
        if ($null -ne $nm)    { Release-Transient $nm    'Name';  $nm    = $null }
        if ($null -ne $names) { Release-Transient $names 'Names'; $names = $null }
    }
}

# WHAT ENDED PROTECTION PROBE RUN 4, and it was this function's own doing.
#
#     System.InvalidCastException: Unable to cast object of type 'System.Double'
#     to type 'System.String'.
#       at phase10_protection_probe.ps1:175   source: $range.Value2 = $Value
#
# The probe had reimplemented the accepted Set-NamedValue and, in reimplementing
# it, dropped BOTH of the things that make it work: the [double] cast at the
# assignment and the ClearContents branch for a null. (The dropped ClearContents
# branch is already on the record above - it is what made Run 2 contradict
# itself.) A polymorphic `$range.Value2 = $Value` with no cast is exactly the
# shape Phase-5 Runtime Run 4 already proved defective, at
# phase5_gate_b_scenarios.ps1:922, with the same exception and the same type
# pair: PowerShell binds a COM property setter per call site, so one line that
# is asked to carry more than one CLR type cannot.
#
# THE FIX IS NOT A NEW SETTER. It is the accepted one, copied VERBATIM - the
# same thing phase10_benchmark.ps1 does, and proved a verbatim copy by control
# rather than asserted to be one. These harnesses are scripts, not modules;
# copying the primitive and pinning the copy is how this repository shares it.
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

# ONE COM ASSIGNMENT SITE PER CLR TYPE, for the cell writes that are not named
# ranges. The cast on each branch is a no-op whose only job is to give that
# branch its OWN bound call site; a single line serving two types is the Run-4
# defect. An unsupported type is refused BY NAME rather than coerced, because
# coercing it is how a Double becomes the string '2026' and the workbook then
# refuses a timeline for a reason that has nothing to do with protection.
function Set-ProbeCellExact {
    param($Cell, $Value)
    if ($null -eq $Value)        { $null = $Cell.ClearContents() }
    elseif ($Value -is [string]) { $Cell.Value2 = [string]$Value }
    elseif ($Value -is [double]) { $Cell.Value2 = [double]$Value }
    elseif ($Value -is [bool])   { $Cell.Value2 = [bool]$Value }
    else {
        throw ('the captured cell value is a ' + $Value.GetType().FullName +
               ', which this probe will not write back by coercion')
    }
}

# EXACT MEANS THE TYPE TOO. [string]2026 -ceq [string]'2026' is true, so a
# stringified numeric would pass a text comparison while being the wrong thing
# in the cell. This is the accepted Test-Phase5ExactValue rule: CLR type
# identity first, then value.
function Test-ProbeExactValue {
    param($Actual, $Expected)
    if ($null -eq $Expected) { return ($null -eq $Actual) }
    if ($null -eq $Actual)   { return $false }
    if ($Actual.GetType().FullName -cne $Expected.GetType().FullName) { return $false }
    return ([string]$Actual -ceq [string]$Expected)
}

# ===========================================================================
# THE _Calc PUBLICATION TABLES: WATCHING WHAT CALCULATE ACTUALLY RESHAPES
# ===========================================================================
# WINDOWS RUN 6 EXPOSED A WEAKNESS IN THIS PROBE'S OWN PRINCIPLE. The five
# watched tables are the registers and the grids, so Calculate was judged "by
# its announcement" - and this probe's whole argument is that an announcement of
# success is not proof that a structural operation happened. It is the argument
# that made Benchmark Run 3 a harness defect rather than a production one.
#
# SO THE _Calc TABLES ARE WATCHED TOO, for Calculate only. modCalcReport's
# ResizeBody adds and deletes ListRows on all five of them, which is the
# structural work protection would have blocked.
#
# THE AUTHORITY IS THE INSPECTION, NOT A LITERAL. calc.sheet and
# calc.tables[*].table_name come from the same Gate-B inspection everything else
# here reads, and calc.tables[*].row_rule states what each table's row count
# means - "one row per applied project year" for calc_years and calc_annual.
# That rule is what makes the observation predictive rather than merely
# different: with the applied duration at 3 they must end at 3 rows.
function Get-ProbeCalcTables {
    param($Inspection)
    $calc = Get-ProbeRequiredProperty -InputObject $Inspection -Name 'calc' `
        -Where 'the Gate-B inspection'
    $sheet = Get-ProbeScalarString -InputObject $calc -Name 'sheet' `
        -Where 'the Gate-B inspection calc block'
    $tables = Get-ProbeRequiredProperty -InputObject $calc -Name 'tables' `
        -Where 'the Gate-B inspection calc block'
    $out = @()
    foreach ($property in @($tables.PSObject.Properties)) {
        $spec = $property.Value
        $out += [pscustomobject]@{
            Key      = [string]$property.Name
            Sheet    = $sheet
            Table    = (Get-ProbeScalarString -InputObject $spec -Name 'table_name' `
                            -Where ('the ' + [string]$property.Name + ' calc table'))
            RowRule  = (Get-ProbeScalarString -InputObject $spec -Name 'row_rule' `
                            -Where ('the ' + [string]$property.Name + ' calc table'))
        }
    }
    return $out
}

# One row-count reading per _Calc table. Read through the ALREADY-RESOLVED
# worksheet, exactly as the watched targets are, so no second Worksheets
# acquisition appears in this script.
function Get-ProbeCalcShapes {
    param($Worksheet, $CalcTables)
    $los = $null
    $shapes = @()
    try {
        $los = $Worksheet.ListObjects
        foreach ($entry in @($CalcTables)) {
            $lo = $null
            try {
                $lo = $los.Item([string]$entry.Table)
                $rows = $null; $cols = $null
                try {
                    $rows = $lo.ListRows
                    $cols = $lo.ListColumns
                    $shapes += [pscustomobject]@{
                        Key  = [string]$entry.Key
                        Table = [string]$entry.Table
                        Rows = (Measure-ProbeCollection -Collection $rows -Label 'ListRow' `
                                    -Where ([string]$entry.Table))
                        Cols = (Measure-ProbeCollection -Collection $cols -Label 'ListColumn' `
                                    -Where ([string]$entry.Table))
                    }
                } finally {
                    if ($null -ne $cols) { Release-Transient $cols 'ListColumns'; $cols = $null }
                    if ($null -ne $rows) { Release-Transient $rows 'ListRows';    $rows = $null }
                }
            } finally {
                if ($null -ne $lo) { Release-Transient $lo 'ListObject'; $lo = $null }
            }
        }
    } finally {
        if ($null -ne $los) { Release-Transient $los 'ListObjects'; $los = $null }
    }
    return $shapes
}

# The tables whose row rule is "one row per applied project year". After a
# successful Calculate they must hold exactly the applied duration, and that is
# a stronger claim than "something changed": only ResizeBody can produce it.
function Get-ProbePerYearCalcTables {
    param($CalcTables)
    $out = @()
    foreach ($entry in @($CalcTables)) {
        if ([string]$entry.RowRule -eq 'one row per applied project year') {
            $out += [string]$entry.Table
        }
    }
    return $out
}

# ===========================================================================
# WHICH WAY DID THE TABLE MOVE?
# ===========================================================================
# WHY DIRECTION AND NOT MERELY CHANGE. Windows Run 7 proved ListColumns.Add and
# ListRows.Add under the production structural window. It did NOT prove
# ListColumns.Delete or ListRows.Delete - and ListRow.Delete() is the exact call
# Benchmark Run 3 died on:
#
#   "Table features aren't available because the sheet is protected"
#     at a ListRow.Delete()
#
# So "the shape changed" is not enough for the delete path. A shrink round has
# to observe columns and rows going DOWN, and a growth reading must never be
# accepted as delete evidence.
function Get-ProbeColumnDirection {
    param($Before, $After, $Resolution)
    $grew = @(); $shrank = @()
    foreach ($entry in @($Resolution.Targets)) {
        $key = [string]$entry.Key
        $b = $Before[$key]
        $a = $After[$key]
        if ($null -eq $b -or $null -eq $a) { continue }
        if ([int]$a.Columns -gt [int]$b.Columns) { $grew += $key }
        elseif ([int]$a.Columns -lt [int]$b.Columns) { $shrank += $key }
    }
    return [pscustomobject]@{ Grew = @($grew); Shrank = @($shrank) }
}

# ===========================================================================
# ONE CALCULATE ROUND, RUN TWICE
# ===========================================================================
# THE GROWTH ROUND AND THE SHRINK ROUND SHARE ONE IMPLEMENTATION. Two copies of
# eighty lines of shape reading would drift, and the second copy is exactly
# where a weaker check would appear.
#
# ExpectedRows is the applied duration for THIS round. calc_years and calc_annual
# carry the row rule "one row per applied project year", and Stage A builds every
# _Calc table with a single body row - so the growth round expects 3 and the
# shrink round expects 1, and only ResizeBody can produce either.
function Invoke-ProbeCalculateRound {
    param($Excel, $Workbook, $Resolution, $CalcTables, $PerYearTables,
          [int]$ExpectedRows, [string]$Label)
    $lines = New-Object System.Collections.ArrayList
    $calcSheet = $null
    $before = @(); $after = @()
    $outcome = $null
    try {
        Set-ProbeStage -Stage 'endpoint' `
            -Action ('PREPARING TO TEST: reading the _Calc table shapes before Calculate (' +
                     $Label + ')') -Endpoint 'PCCM_Calculate'
        $calcSheet = $Resolution.Sheets.Item([string]@($CalcTables)[0].Sheet)
        $before = @(Get-ProbeCalcShapes -Worksheet $calcSheet -CalcTables $CalcTables)

        $outcome = Invoke-ProbeEndpoint -Excel $Excel -Workbook $Workbook `
            -Endpoint 'PCCM_Calculate' -Resolution $Resolution

        Set-ProbeStage -Stage 'endpoint' `
            -Action ('reading the _Calc table shapes after Calculate (' + $Label + ')') `
            -Endpoint 'PCCM_Calculate'
        $after = @(Get-ProbeCalcShapes -Worksheet $calcSheet -CalcTables $CalcTables)
    } finally {
        if ($null -ne $calcSheet) { Release-Transient $calcSheet 'Worksheet(_Calc)'; $calcSheet = $null }
    }

    $changed = @(); $shrank = @()
    foreach ($b in @($before)) {
        $match = @(@($after) | Where-Object { [string]$_.Key -eq [string]$b.Key })
        if (@($match).Count -ne 1) {
            $null = $lines.Add('    ' + [string]$b.Table + ': NOT READ BACK')
            continue
        }
        $now = @($match)[0]
        $line = ('    ' + [string]$b.Table + ': ' + [string]$b.Rows + 'x' + [string]$b.Cols +
                 ' -> ' + [string]$now.Rows + 'x' + [string]$now.Cols)
        if (([int]$b.Rows -ne [int]$now.Rows) -or ([int]$b.Cols -ne [int]$now.Cols)) {
            $changed += [string]$b.Table
            $line = $line + '   CHANGED'
        }
        if ([int]$now.Rows -lt [int]$b.Rows) {
            $shrank += [string]$b.Table
            $line = $line + '   (rows DELETED)'
        }
        $null = $lines.Add($line)
    }

    # THE PREDICTIVE CHECK. Only demanded of a Calculate that SUCCEEDED: a
    # refusal is entitled to leave the tables alone, and failing it for that
    # would manufacture a failure - the same mistake as calling any refusal a
    # protection block.
    $proof = 'NOT ESTABLISHED'
    if ([string]$outcome.Outcome -ne 'SUCCEEDED') {
        $null = $lines.Add('    Calculate did not succeed, so no shape is required of it. ' +
                           'A refusal is entitled to leave the tables alone.')
    }
    if ([string]$outcome.Outcome -eq 'SUCCEEDED') {
        $wrong = @()
        foreach ($table in @($PerYearTables)) {
            $row = @(@($after) | Where-Object { [string]$_.Table -eq [string]$table })
            if (@($row).Count -ne 1) { $wrong += ($table + ': not read back'); continue }
            if ([int]@($row)[0].Rows -ne $ExpectedRows) {
                $wrong += ($table + ': ' + [string]@($row)[0].Rows + ' rows, expected ' +
                           [string]$ExpectedRows + ' (one per applied project year)')
            }
        }
        if ((@($wrong).Count -eq 0) -and (@($PerYearTables).Count -gt 0) -and
            (@($changed).Count -gt 0)) {
            $proof = 'OBSERVED'
        } else {
            $proof = 'CONTRADICTED'
            $null = $lines.Add('    BUT Calculate announced success WITHOUT the contracted shape:')
            foreach ($problem in @($wrong)) { $null = $lines.Add('      ' + $problem) }
            if (@($changed).Count -eq 0) {
                $null = $lines.Add('      no _Calc table changed shape at all')
            }
        }
    }

    return [pscustomobject]@{
        Label        = $Label
        Outcome      = $outcome
        Expected     = $ExpectedRows
        Changed      = @($changed)
        RowsDeleted  = @($shrank)
        Proof        = $proof
        Lines        = @($lines)
    }
}

# ===========================================================================
# IS THIS REFUSAL ACTUALLY ABOUT PROTECTION?
# ===========================================================================
# RUN 5 IS WHY THIS EXISTS. It produced the real answer - PCCM_ApplyTimeline was
# invoked and refused with
#
#   Error 1004: Table features aren't available because the sheet is protected.
#
# and that IS protection blocking a structural operation. But the same run also
# summarised "2 of 4 production endpoints did not succeed", and the second one
# was PCCM_Calculate refusing because the applied timeline was still pending.
# That is a business prerequisite, not protection, and counting it made the
# verdict look better-evidenced than it was.
#
# SO BLOCKED NEEDS EVIDENCE ATTRIBUTABLE TO THE PROTECTED OPERATION, and a
# generic refusal is not it. A validation refusal, a missing input, stale state
# or an unmet structural prerequisite must never produce BLOCKED - each of those
# would still be refused on a completely unprotected workbook.
#
# THE SIGNATURE IS EXCEL'S OWN, matched on BOTH halves. '1004' alone is a very
# common Excel error number and appears in refusals that have nothing to do with
# protection; the protection wording alone could appear in a sentence the probe
# or the workbook wrote ABOUT protection. Both, together, in an announcement from
# an endpoint that was actually invoked, is the evidence.
function Test-ProbeProtectionBlocked {
    param($Outcome)
    # NOT INVOKED IS NOT A RESULT. An endpoint the probe never entered says
    # nothing about production at all, protection included.
    if (-not [bool]$Outcome.Invoked) { return $false }
    if ([string]$Outcome.Outcome -eq 'SUCCEEDED') { return $false }

    $text = ([string]$Outcome.Result + ' ' + [string]$Outcome.Raised)
    if (-not ($text -match '1004')) { return $false }
    if (-not ($text -match '(?i)protect')) { return $false }
    # AND EXCEL'S OWN SENTENCE, not merely the two tokens in one string. This is
    # the wording Run 5 recorded, and it names the capability at issue.
    return [bool]($text -match "(?i)table features aren't available because the sheet is protected")
}

# ===========================================================================
# THE ENDPOINT PRECONDITIONS: the timeline inputs
# ===========================================================================
# NOT A SECOND FIXTURE CONTRACT. The values are the accepted Gate-B/Phase-7
# shape, and the reason they are valid is already on the record at
# phase7_timing_scenarios.ps1:465 - "base_year 2026 with start_year 2027 is the
# accepted shape: the generated inflation columns begin at BaseYear + 1, so
# 2027..2029 are exactly the three project years". The defined names come from
# the same Gate-B inspection the accepted fixture reads, never from a literal.
#
# THE MINIMUM, AND ONLY THE MINIMUM. PCCM_ApplyTimeline reads a timeline triple;
# the accepted fixture also sets discount_rate, which this command does not
# need. Setting more than the command requires would widen the probe into a
# fixture builder.
#
# ALL THREE ARE INTEGER INPUTS and are written as System.Double. modTimeline's
# ReadTriple gates every one of them through TryReadDouble and then d = Int(d),
# so a value written as text is not "nearly right": the command REFUSES, and a
# refusal has no way to distinguish itself from protection blocking the work.
# Stringifying here would manufacture the very verdict the probe exists to test.
# WHAT WINDOWS RUN 6 ADDED TO THIS LIST, and why it is not a workaround.
#
#   FAIL|Calculate|Discount Rate: the value is blank. A blank is not zero.
#
# That is production refusing correctly. modCalcResolve.ResolveAppliedTimeline
# requires NM_INPUT_DISCOUNT_RATE through NumericNamedCell, which refuses a
# blank BY DESIGN - "a blank is an unmade assumption, not zero" - and refuses a
# numeric-looking STRING too, because IsRealNumber tests the VarType rather than
# parsing. So the fix is to supply the input the way a user supplies it: a real
# number in the accepted Setup cell, through the accepted setter that casts
# [double]. Nothing is hard-coded around the validation and nothing is relaxed.
#
# THE VALUE IS THE ACCEPTED FIXTURE'S. phase7_timing_scenarios.ps1 declares
# discount_rate = 0.05 beside the same timeline triple; this reuses that
# declaration rather than inventing a second one.
function Get-ProbeDeclaredInputs {
    param($Inspection)
    $inputs = Get-ProbeRequiredProperty -InputObject $Inspection -Name 'inputs' `
        -Where 'the Gate-B inspection'
    # Endpoint says WHICH command needs the value, so a transcript can show that
    # the discount rate is a Calculate prerequisite and not a timeline one.
    $wanted = @(
        @{ Key = 'base_year';          Value = [double]2026; Endpoint = 'PCCM_ApplyTimeline' },
        @{ Key = 'project_start_year'; Value = [double]2027; Endpoint = 'PCCM_ApplyTimeline' },
        @{ Key = 'duration_years';     Value = [double](Get-ProbeGrowthYears); Endpoint = 'PCCM_ApplyTimeline' },
        @{ Key = 'discount_rate';      Value = [double]0.05; Endpoint = 'PCCM_Calculate' }
    )
    return New-ProbeInputRecords -Inputs $inputs -Wanted $wanted
}

# THE TWO APPLIED DURATIONS THIS PROBE USES, DECLARED IN ONE PLACE. The growth
# round applies 3 project years and the shrink round applies 1; every expectation
# downstream is derived from these rather than restating a number.
function Get-ProbeGrowthYears { return 3 }
function Get-ProbeShrinkYears { return 1 }

# THE SHRINK PRECONDITION. One input changes and it changes the same way every
# other one did: a genuine Double through the accepted Set-NamedValue, read back
# and type-checked before the endpoint is touched. 1 is inside the accepted bound
# - modTimeline.ReadTriple requires 1 <= d <= LIMIT_MAX_YEAR_COLUMNS - so this is
# a valid timeline, not a value chosen to force a failure.
function Get-ProbeShrinkInputs {
    param($Inspection)
    $inputs = Get-ProbeRequiredProperty -InputObject $Inspection -Name 'inputs' `
        -Where 'the Gate-B inspection'
    $wanted = @(
        @{ Key = 'duration_years'; Value = [double](Get-ProbeShrinkYears); Endpoint = 'PCCM_ApplyTimeline' }
    )
    return New-ProbeInputRecords -Inputs $inputs -Wanted $wanted
}

# ONE RECORD BUILDER FOR BOTH SETS. A second copy is where a missing scalar-shape
# check or a missing [double] would appear.
function New-ProbeInputRecords {
    param($Inputs, $Wanted)
    $out = @()
    foreach ($entry in @($Wanted)) {
        $key = [string]$entry.Key
        $spec = Get-ProbeRequiredProperty -InputObject $Inputs -Name $key `
            -Where 'the Gate-B inspection inputs'
        $out += [pscustomobject]@{
            Key         = $key
            Endpoint    = [string]$entry.Endpoint
            DefinedName = (Get-ProbeScalarString -InputObject $spec -Name 'defined_name' `
                               -Where ('the ' + $key + ' input'))
            Value       = [double]$entry.Value
        }
    }
    # NO UNARY COMMA - the caller's @() keeps a short result an array without
    # double-wrapping it. That is the Run-3 rule and it still applies here.
    return $out
}

function Set-ProbeDeclaredInputs {
    param($Workbook, $Inputs)
    foreach ($entry in @($Inputs)) {
        Set-NamedValue -Workbook $Workbook -DefinedName ([string]$entry.DefinedName) `
            -Value ([double]$entry.Value)
    }
}

# READ BACK BEFORE THE ENDPOINT IS TOUCHED. Returns the problems as plain
# strings; an empty result means every input landed with the right value AND the
# right type. A setter or readback failure is PROBE INSTRUMENTATION failure - it
# is never a statement about the production endpoint, which has not run.
function Test-ProbeDeclaredInputs {
    param($Workbook, $Inputs)
    $problems = @()
    foreach ($entry in @($Inputs)) {
        $name = [string]$entry.DefinedName
        $expected = [double]$entry.Value
        $actual = Get-ProbeNamedValue -Workbook $Workbook -DefinedName $name
        if ($null -eq $actual) {
            $problems += ($name + ': reads back blank, so the write did not land')
            continue
        }
        # TYPE FIRST, for the reason above: a numeric input that came back as
        # text would pass any value-only comparison and then be refused by the
        # command for a reason the probe would misread as protection.
        if ($actual -isnot [double]) {
            $problems += ($name + ': reads back as ' + $actual.GetType().FullName +
                          ', not System.Double - the numeric input was stringified')
            continue
        }
        if ([double]$actual -ne $expected) {
            $problems += ($name + ': reads back ' + [string]$actual +
                          ', expected ' + [string]$expected)
        }
    }
    return $problems
}

# IS THE WORKBOOK ACTUALLY PROTECTED RIGHT NOW? Asked of the sheets themselves,
# at the moment it matters, by ENUMERATION - there is no index and no `.Count`.
function Get-ProbeProtectionState {
    param($Workbook, [string]$Where)
    $sheets = $null
    $protectedNames = @(); $unprotectedNames = @()
    try {
        $sheets = $Workbook.Worksheets
        foreach ($sheet in @($sheets)) {
            try {
                if ($sheet.ProtectContents) { $protectedNames += [string]$sheet.Name }
                else                        { $unprotectedNames += [string]$sheet.Name }
            } finally {
                if ($null -ne $sheet) { Release-Transient $sheet 'Worksheet' }
            }
        }
    } finally {
        if ($null -ne $sheets) { Release-Transient $sheets 'Worksheets'; $sheets = $null }
    }
    # A WORKBOOK WITH NO SHEETS IS NOT AN ANSWER. Refusing here is what stops a
    # collection the probe failed to enumerate from being reported as "0 of 0
    # protected", which would read as evidence.
    if ((@($protectedNames).Count + @($unprotectedNames).Count) -lt 1) {
        throw ($Where + ': the workbook enumerated no worksheets, so its protection ' +
               'state could not be established')
    }
    return [pscustomobject]@{
        Protected   = @($protectedNames).Count
        Total       = (@($protectedNames).Count + @($unprotectedNames).Count)
        Unprotected = @($unprotectedNames)
        Structure   = [bool]$Workbook.ProtectStructure
    }
}

function Format-ProbeProtection {
    param($State)
    $unprotected = @(Get-ProbeProperty -InputObject $State -Name 'Unprotected')
    $text = ([string](Get-ProbeProperty -InputObject $State -Name 'Protected') + ' of ' +
             [string](Get-ProbeProperty -InputObject $State -Name 'Total') +
             ' sheets protected, structure ' +
             [string](Get-ProbeProperty -InputObject $State -Name 'Structure'))
    if (@($unprotected).Count -gt 0) { $text = $text + '; NOT protected: ' + ($unprotected -join ', ') }
    return $text
}

# COUNT BY ENUMERATION, NEVER BY `.Count` ON A COM COLLECTION.
#
# `.Count` on a COM collection is the exact shape that ended Probe Run 1, and it
# is not made safe by the collection usually having one. `@()` normalises
# whatever the adapter returns before anything is read from it, and each
# enumerated RCW is released as the accepted transient policy requires.
#
# AN ABSENT COLLECTION IS A REFUSAL, NOT A ZERO. A table whose columns could not
# be enumerated must not be reported as having none: that would be a shape
# change the probe invented.
function Measure-ProbeCollection {
    param($Collection, [string]$Label, [string]$Where)
    if ($null -eq $Collection) {
        throw ($Where + ': expected a ' + $Label + ' collection and got nothing')
    }
    $count = 0
    foreach ($item in @($Collection)) {
        if ($null -eq $item) { continue }
        $count++
        Release-Transient $item $Label
    }
    return $count
}

# EVERY TABLE THE STRUCTURAL COMMANDS RESHAPE, from the manifest rather than
# from a list typed here. The tab name and the table name are two independent
# fields of the same entry; neither is derived from the other.
function Get-ProbeWatchedTables {
    param($Manifest)
    $watched = @()
    foreach ($register in @($Manifest.registers)) {
        $watched += [pscustomobject]@{
            Key = [string]$register.key
            Sheet = [string]$register.sheet
            Table = [string]$register.table_name
        }
    }
    foreach ($grid in @($Manifest.grids)) {
        $watched += [pscustomobject]@{
            Key = [string]$grid.key
            Sheet = [string]$grid.sheet
            Table = [string]$grid.table_name
        }
    }
    # NO UNARY COMMA. `return ,@($x)` emits ONE pipeline item that IS the array,
    # and a caller writing `@(f)` then wraps it AGAIN - `@()` collects pipeline
    # items, it does not flatten a nested array. That double wrap is what turned
    # five records into one, and `[string]` on its array-valued properties is
    # what produced 'Cost Lines Risk Register Cost Profiling ...'.
    #
    # Emitting the records normally gives the caller five pipeline items, and the
    # caller's `@()` keeps a zero- or one-record result an array.
    return $watched
}

# ===========================================================================
# THE SHAPE OF A WATCHED RECORD, PROVED BEFORE IT IS USED
# ===========================================================================
# PROBE RUN 3 ROOT CAUSE, AS A RULE RATHER THAN A LINE. Five records collapsed
# into one whose Key, Sheet and Table were each five-element arrays, and
# `[string]` on an array joins it with spaces - so the probe asked Excel for a
# worksheet named 'Cost Lines Risk Register Cost Profiling Risk Profiling
# Inflation' and got DISP_E_BADINDEX. The coercion HID the shape defect: it
# turned a structural error into a plausible-looking name.
#
# SO THE SHAPE IS CHECKED BEFORE ANY COERCION. `[string]` is never allowed to be
# the thing that discovers a property is a collection.
function Get-ProbeScalarString {
    param($InputObject, [string]$Name, [string]$Where)
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) {
        throw ($Where + ": the record carries no '" + $Name + "' property")
    }
    $value = $property.Value
    if ($null -eq $value) {
        throw ($Where + ": '" + $Name + "' is null")
    }
    # THE ARRAY TEST COMES FIRST, BEFORE ANY CAST. A property holding five values
    # must refuse as probe-invalid, not quietly become five words.
    if ($value -is [System.Array] -or $value -is [System.Collections.IEnumerable] -and
        $value -isnot [string]) {
        throw ($Where + ": '" + $Name + "' holds a collection of " +
               [string](@($value).Count) + ' values where one was expected. The ' +
               'watched records have been collapsed into an aggregate.')
    }
    if ($value -isnot [string]) {
        throw ($Where + ": '" + $Name + "' is a " + $value.GetType().FullName +
               ' where a string was expected')
    }
    if ([string]::IsNullOrWhiteSpace($value)) {
        throw ($Where + ": '" + $Name + "' is blank")
    }
    return [string]$value
}

# FIVE DISTINCT RECORDS, OR THE PROBE STOPS. Checked against the manifest's own
# count so the number is never a literal in this file.
function Assert-ProbeWatchedShape {
    param($Watched, $Manifest)
    $records = @($Watched)
    $expected = (@($Manifest.registers).Count + @($Manifest.grids).Count)
    if (@($records).Count -ne $expected) {
        throw ('the probe holds ' + [string]@($records).Count + ' watched record(s) where the ' +
               'manifest declares ' + [string]$expected + '. The records have been ' +
               'collapsed, dropped or double-wrapped.')
    }
    $keys = @(); $pairs = @()
    $index = 0
    foreach ($record in $records) {
        $index++
        $where = 'watched record ' + [string]$index
        if ($record -is [System.Array]) {
            throw ($where + ' is an array of ' + [string](@($record).Count) +
                   ' items, not a record. The collection boundary collapsed.')
        }
        $key = Get-ProbeScalarString -InputObject $record -Name 'Key' -Where $where
        $sheet = Get-ProbeScalarString -InputObject $record -Name 'Sheet' -Where $where
        $table = Get-ProbeScalarString -InputObject $record -Name 'Table' -Where $where
        if ($keys -contains $key) { throw ($where + ": the key '" + $key + "' is not unique") }
        $keys += $key
        $pair = $sheet + '!' + $table
        if ($pairs -contains $pair) { throw ($where + ': ' + $pair + ' is not unique') }
        $pairs += $pair
    }
    return @($records).Count
}

# ===========================================================================
# RESOLVE ONCE, HOLD FOR THE SESSION
# ===========================================================================
# PROBE RUN 2 ROOT CAUSE. Every shape read did this:
#
#     $sheets = $Workbook.Worksheets
#     $ws = $sheets.Item($SheetName)
#     ... finally { Release-Transient $sheets 'Worksheets' }
#
# `Workbook.Worksheets` hands back the SAME underlying collection every time, so
# a run that read five tables around each of four endpoints acquired and
# released that one object dozens of times. The third cycle came back unusable
# and `Item('Cost Lines')` answered DISP_E_BADINDEX - not because the sheet was
# missing (all five tab names and all five table names exist in the built
# workbook and a control now proves it) but because the collection had been
# released out from under the lookup.
#
# THE CHURN IS GONE. Every worksheet and every ListObject the probe watches is
# resolved ONCE, held for the session, and released once at the end. Resolution
# is also EVIDENCE: the tab name, the CodeName and the table name are recorded
# for each target, so a later reader can see exactly what was measured.
#
# NEITHER IDENTIFIER IS INFERRED FROM THE OTHER. The tab name and the table name
# both come from the manifest, independently. The CodeName is read from Excel
# and is recorded only - it is never used to find anything.
function Resolve-ProbeTargets {
    param($Workbook, $Watched)
    $sheets = $Workbook.Worksheets
    $targets = New-Object System.Collections.ArrayList
    $index = 0
    foreach ($entry in @($Watched)) {
        $index++
        $where = 'watched record ' + [string]$index
        # SCALAR PROVED, THEN CONVERTED. Never the other way round.
        $sheetName = Get-ProbeScalarString -InputObject $entry -Name 'Sheet' -Where $where
        $tableName = Get-ProbeScalarString -InputObject $entry -Name 'Table' -Where $where
        $ws = $null
        try {
            $ws = $sheets.Item($sheetName)
        } catch {
            throw ('the workbook has no worksheet named ' + [char]39 + $sheetName + [char]39 +
                   ' (asked for by the manifest entry ' +
                   (Get-ProbeScalarString -InputObject $entry -Name 'Key' -Where $where) + '): ' +
                   (Format-Err $_))
        }
        $los = $null; $lo = $null
        try {
            $los = $ws.ListObjects
            $lo = $los.Item($tableName)
        } catch {
            throw ('worksheet ' + [char]39 + $sheetName + [char]39 + ' carries no table named ' +
                   [char]39 + $tableName + [char]39 + ': ' + (Format-Err $_))
        }
        $null = $targets.Add([pscustomobject]@{
            Key       = (Get-ProbeScalarString -InputObject $entry -Name 'Key' -Where $where)
            Sheet     = $sheetName
            CodeName  = [string]$ws.CodeName
            Table     = $tableName
            Worksheet = $ws
            ListObject = $lo
            ListObjects = $los
        })
    }
    return [pscustomobject]@{ Sheets = $sheets; Targets = @($targets) }
}

function Release-ProbeTargets {
    param($Resolution)
    if ($null -eq $Resolution) { return }
    foreach ($target in @($Resolution.Targets)) {
        if ($null -ne $target.ListObject)  { Release-Transient $target.ListObject  'ListObject' }
        if ($null -ne $target.ListObjects) { Release-Transient $target.ListObjects 'ListObjects' }
        if ($null -ne $target.Worksheet)   { Release-Transient $target.Worksheet   'Worksheet' }
    }
    if ($null -ne $Resolution.Sheets) { Release-Transient $Resolution.Sheets 'Worksheets' }
}

# THE SHAPE OF ONE ALREADY-RESOLVED TABLE, READ AND NEVER CHANGED. Rows and
# columns only, by enumeration - this is a read, not a structural operation.
function Get-ProbeTableShape {
    param($Target)
    $where = [string]$Target.Sheet + '!' + [string]$Target.Table
    $lo = $Target.ListObject
    $columnCollection = $null; $rowCollection = $null
    try {
        $columnCollection = $lo.ListColumns
        $columns = Measure-ProbeCollection -Collection $columnCollection `
            -Label 'ListColumn' -Where $where
        $rowCollection = $lo.ListRows
        $rows = Measure-ProbeCollection -Collection $rowCollection `
            -Label 'ListRow' -Where $where
    } finally {
        if ($null -ne $rowCollection)    { Release-Transient $rowCollection    'ListRows';    $rowCollection = $null }
        if ($null -ne $columnCollection) { Release-Transient $columnCollection 'ListColumns'; $columnCollection = $null }
    }
    return [pscustomobject]@{ Table = [string]$Target.Table; Columns = $columns; Rows = $rows }
}

function Get-ProbeAllShapes {
    param($Resolution)
    $shapes = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($target in @($Resolution.Targets)) {
        $shapes.Add([string]$target.Key, (Get-ProbeTableShape -Target $target))
    }
    return $shapes
}

function Get-ProbeShapeDelta {
    param($Before, $After, $Resolution)
    $changes = @()
    foreach ($entry in @($Resolution.Targets)) {
        $key = [string]$entry.Key
        $b = $Before[$key]
        $a = $After[$key]
        if ($null -eq $b -or $null -eq $a) { continue }
        if (([int]$b.Columns -ne [int]$a.Columns) -or ([int]$b.Rows -ne [int]$a.Rows)) {
            $changes += ($key + ' ' + [string]$b.Rows + 'x' + [string]$b.Columns +
                         ' -> ' + [string]$a.Rows + 'x' + [string]$a.Columns)
        }
    }
    # SAME REASON, AND THIS ONE WAS WORSE THAN AN ABORT. Double-wrapped, an
    # EMPTY change list arrived at the caller as a one-element array, so
    # `StructuralEffect` would have been true for every endpoint - and the probe
    # could have reached PRODUCTION IS FINE having observed no shape change at
    # all. A wrong conclusion is more dangerous than a failed run.
    return $changes
}

# ===========================================================================
# THE LOCKED-CELL CONTROL
# ===========================================================================
# WHAT PROBE RUN 2 GOT WRONG, TWICE.
#
#   IT DID NOT USE A LOCKED CELL. It wrote to `inpDiscountRate`, which is an
#   EDITABLE INPUT and is Locked=False by design. Writing there proves nothing
#   about UserInterfaceOnly: a user could type in that cell.
#
#   IT REPORTED SUCCESS AND FAILURE AT ONCE. The write and the RESTORE sat in
#   the same try. The write succeeded and printed SUCCEEDED; the restore threw -
#   the original value was BLANK, and the accepted Set-NamedValue has a
#   ClearContents branch for exactly that which this copy had dropped - and the
#   catch then printed REFUSED. `$controlWorked` had already been set true, so
#   the closing summary said True. Three contradictory lines from one try block.
#
# THE CAPABILITY AND THE CLEANUP ARE NOW SEPARATE FACTS, in separate try blocks.
# A failed restore is NEVER rewritten as "protection blocks value writes": it
# makes the control INCONCLUSIVE and says which step failed.
#
# THE TARGET IS PROVED LOCKED BEFORE ANYTHING IS WRITTEN. It is the header cell
# of the first watched table - model-controlled, non-blank, and reached through
# the ListObject the manifest named rather than through a typed address.
function Invoke-ProbeLockedCellControl {
    param($Workbook, $Resolution)
    $lines = New-Object System.Collections.ArrayList
    $target = @($Resolution.Targets)[0]
    $where = [string]$target.Sheet + '!' + [string]$target.Table
    $probeText = 'PCCM-PROBE'

    $header = $null; $cell = $null
    try {
        $header = $target.ListObject.HeaderRowRange
        $cell = $header.Cells(1, 1)

        # 1. THE PRECONDITIONS, PROVED NOW rather than assumed from a projection.
        if (-not $target.Worksheet.ProtectContents) {
            $null = $lines.Add('INCONCLUSIVE: ' + [string]$target.Sheet + ' is not protected, so a write to it proves nothing')
            return [pscustomobject]@{ Result = 'INCONCLUSIVE'; Detail = 'the control sheet was not protected'; Lines = @($lines) }
        }
        if (-not $cell.Locked) {
            $null = $lines.Add('INCONCLUSIVE: the ' + $where + ' header cell is NOT locked, so writing to it says nothing about UserInterfaceOnly')
            return [pscustomobject]@{ Result = 'INCONCLUSIVE'; Detail = 'the control cell was not locked'; Lines = @($lines) }
        }
        $original = $cell.Value2
        if ([string]::IsNullOrWhiteSpace([string]$original)) {
            $null = $lines.Add('INCONCLUSIVE: the ' + $where + ' header cell is blank, so an exact restoration could not be verified')
            return [pscustomobject]@{ Result = 'INCONCLUSIVE'; Detail = 'the control cell was blank'; Lines = @($lines) }
        }
        $null = $lines.Add('target             : ' + $where + ' header cell, Locked=True, sheet protected')
        $null = $lines.Add('original value     : ' + [char]39 + [string]$original + [char]39)

        # 2. THE CAPABILITY, in its own try so nothing after it can rewrite the answer.
        $writeRaised = ''
        try { Set-ProbeCellExact -Cell $cell -Value $probeText } catch { $writeRaised = (Format-Err $_) }
        if ($writeRaised -ne '') {
            $null = $lines.Add('write capability   : REFUSED - ' + $writeRaised)
            $null = $lines.Add('so protection is blocking code VALUE writes as well, which is a')
            $null = $lines.Add('DIFFERENT and larger finding than the table-structure one.')
            return [pscustomobject]@{ Result = 'REFUSED'; Detail = ('the write was refused: ' + $writeRaised); Lines = @($lines) }
        }
        $readBack = $cell.Value2
        if (-not (Test-ProbeExactValue -Actual $readBack -Expected $probeText)) {
            $null = $lines.Add('write capability   : INCONCLUSIVE - the write raised nothing but the cell reads back ' + [char]39 + [string]$readBack + [char]39)
            return [pscustomobject]@{ Result = 'INCONCLUSIVE'; Detail = 'the written value did not read back'; Lines = @($lines) }
        }
        $null = $lines.Add('write capability   : SUCCEEDED - the value was written and read back')

        # 3. THE CLEANUP, AS ITS OWN FACT. A failure here does not change the
        #    capability answer; it makes the control untrustworthy, which is a
        #    different thing and is reported as one.
        $restoreRaised = ''
        # SAME CLASS AS THE RUN-4 DEFECT. $original is whatever the header cell
        # actually held, so this site is polymorphic by nature: a String today, a
        # Double the moment a watched table's first header cell is numeric. It
        # gets the per-type dispatch for the same reason the named-value setter
        # does, before another Windows run finds it.
        try { Set-ProbeCellExact -Cell $cell -Value $original } catch { $restoreRaised = (Format-Err $_) }
        if ($restoreRaised -ne '') {
            $null = $lines.Add('control cleanup    : FAILED - ' + $restoreRaised)
            return [pscustomobject]@{ Result = 'INCONCLUSIVE'; Detail = 'the original value could not be restored'; Lines = @($lines) }
        }
        $restored = $cell.Value2
        if (-not (Test-ProbeExactValue -Actual $restored -Expected $original)) {
            $null = $lines.Add('control cleanup    : FAILED - the cell reads back ' + [char]39 + [string]$restored + [char]39 + ' after restoration')
            return [pscustomobject]@{ Result = 'INCONCLUSIVE'; Detail = 'the restored value did not match the original'; Lines = @($lines) }
        }
        $null = $lines.Add('control cleanup    : the original value was restored and verified')

        # 4. AND PROTECTION SURVIVED THE WHOLE THING.
        if (-not $target.Worksheet.ProtectContents) {
            $null = $lines.Add('control cleanup    : FAILED - ' + [string]$target.Sheet + ' is no longer protected')
            return [pscustomobject]@{ Result = 'INCONCLUSIVE'; Detail = 'protection was lost during the control'; Lines = @($lines) }
        }
        $null = $lines.Add('so UserInterfaceOnly IS honoured for code that writes VALUES to a')
        $null = $lines.Add('locked cell. That is a SEPARATE capability from permission to perform')
        $null = $lines.Add('a ListObject structural operation, and it settles nothing about one.')
        return [pscustomobject]@{ Result = 'SUCCEEDED'; Detail = 'a value was written to a proved-locked cell and restored exactly'; Lines = @($lines) }
    } finally {
        if ($null -ne $cell)   { Release-Transient $cell   'Range'; $cell   = $null }
        if ($null -ne $header) { Release-Transient $header 'Range'; $header = $null }
    }
}

# ===========================================================================
# ONE ENDPOINT, ONCE
# ===========================================================================
# THE THREE WAYS AN ENDPOINT CAN FAIL ARE DIFFERENT FACTS and are recorded as
# different facts:
#
#   REFUSED   the workbook announced FAIL|... - production declined, in its own
#             words. This is the answer the question is about.
#   RAISED    Application.Run itself threw - an Excel/VBA runtime failure that
#             never reached an announcement.
#   PROBE     this script failed around the call. Never a statement about
#             production.
function Invoke-ProbeEndpoint {
    param($Excel, $Workbook, [string]$Endpoint, $Resolution)

    # THE EVIDENCE ORDER, AND THE FLAG THAT KEEPS IT HONEST.
    #
    # Probe Run 2 said "doing: invoking the production entry point" and named
    # PCCM_ApplyTimeline while it was still collecting PRE-command evidence, and
    # the endpoint was never reached. A probe-side worksheet lookup was one line
    # away from being read as a production failure.
    #
    # $invoked is set in the instant before Application.Run and NEVER anywhere
    # else. Until it is true, nothing about this endpoint is a statement about
    # production - and the outcome carries the flag so a reader can see which it
    # was without trusting a sentence.
    $invoked = $false

    Set-ProbeStage -Stage 'endpoint' -Action 'PREPARING TO TEST: reading protection before the command' -Endpoint $Endpoint
    $protectionBefore = Get-ProbeProtectionState -Workbook $Workbook -Where ('before ' + $Endpoint)
    Set-ProbeStage -Stage 'endpoint' -Action 'PREPARING TO TEST: reading table shapes before the command' -Endpoint $Endpoint
    $shapesBefore = Get-ProbeAllShapes -Resolution $Resolution

    $raised = ''
    $result = ''
    try {
        $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
        Set-ProbeStage -Stage 'endpoint' -Action 'ENDPOINT INVOKED: Application.Run has been entered' -Endpoint $Endpoint
        $invoked = $true
        $Excel.Run($Endpoint) | Out-Null
        $result = [string]$Excel.Run('PCCM_AutomationResult')
    } catch {
        $raised = (Format-Err $_)
    }

    # THE DECLARED ORDER: the structural effect first, then protection last -
    # protection-after is the closest reading to "did the command leave the
    # workbook as it found it".
    Set-ProbeStage -Stage 'endpoint' -Action 'reading the structural effect back' -Endpoint $Endpoint
    $shapesAfter = Get-ProbeAllShapes -Resolution $Resolution
    Set-ProbeStage -Stage 'endpoint' -Action 'reading protection after the command' -Endpoint $Endpoint
    $protectionAfter = Get-ProbeProtectionState -Workbook $Workbook -Where ('after ' + $Endpoint)
    $changes = @(Get-ProbeShapeDelta -Before $shapesBefore -After $shapesAfter -Resolution $Resolution)

    # AN ENDPOINT THAT WAS NEVER ENTERED IS NOT A RESULT. It is a probe failure,
    # and it is named as one.
    $announced = ([bool](($raised -eq '') -and $invoked))
    $succeeded = ([bool]($announced -and ($result -like 'OK|*')))
    $outcome = 'REFUSED'
    if (-not $invoked)    { $outcome = 'NOT INVOKED' }
    elseif (-not $announced) { $outcome = 'RAISED' }
    elseif ($succeeded)   { $outcome = 'SUCCEEDED' }

    return [pscustomobject]@{
        Endpoint          = $Endpoint
        Invoked           = $invoked
        Outcome           = $outcome
        Result            = $result
        Raised            = $raised
        Changes           = $changes
        ProtectedBefore   = [int](Get-ProbeProperty -InputObject $protectionBefore -Name 'Protected')
        ProtectedAfter    = [int](Get-ProbeProperty -InputObject $protectionAfter  -Name 'Protected')
        TotalSheets       = [int](Get-ProbeProperty -InputObject $protectionAfter  -Name 'Total')
        ProtectionBefore  = (Format-ProbeProtection -State $protectionBefore)
        ProtectionAfter   = (Format-ProbeProtection -State $protectionAfter)
        StructuralEffect  = ([bool](@($changes).Count -gt 0))
        # THE STRUCTURE FLAG AS A BOOLEAN, not only inside the formatted text.
        # Windows Run 7 proved ListColumns.Add does not need workbook-structure
        # protection released, so the verdict has to be able to CHECK that it
        # stayed applied rather than read it in a sentence.
        StructureBefore   = ([bool](Get-ProbeProperty -InputObject $protectionBefore -Name 'Structure'))
        StructureAfter    = ([bool](Get-ProbeProperty -InputObject $protectionAfter  -Name 'Structure'))
        # THE SHAPES THEMSELVES, so a caller can ask which DIRECTION a table
        # moved. `Changes` is formatted text and parsing it back would be a
        # second, worse reading of evidence the probe already holds.
        ShapesBefore      = $shapesBefore
        ShapesAfter       = $shapesAfter
    }
}

function Write-ProbeOutcome {
    param($Outcome, [string]$Expectation)
    Write-ProbeLine ('  ' + [string]$Outcome.Endpoint)
    Write-ProbeLine ('    endpoint invoked   : ' + [string]$Outcome.Invoked)
    Write-ProbeLine ('    outcome            : ' + [string]$Outcome.Outcome)
    if ([string]$Outcome.Raised -ne '') {
        Write-ProbeLine ('    RAISED             : ' + [string]$Outcome.Raised)
    } else {
        Write-ProbeLine ('    announced          : ' + [string]$Outcome.Result)
    }
    Write-ProbeLine ('    protection before  : ' + [string]$Outcome.ProtectionBefore)
    Write-ProbeLine ('    protection after   : ' + [string]$Outcome.ProtectionAfter)
    if ([bool]$Outcome.StructuralEffect) {
        Write-ProbeLine  '    structural effect  : YES'
        foreach ($change in @($Outcome.Changes)) { Write-ProbeLine ('      ' + $change) }
    } else {
        Write-ProbeLine  '    structural effect  : NONE - no watched table changed shape'
    }
    Write-ProbeLine ('    what this settles  : ' + $Expectation)
    Write-ProbeLine ''
}

# WHAT FAILED, WHERE, AND WHAT THE PROBE WAS DOING AT THE TIME.
function Format-ProbeFailure {
    param($ErrorRecord)
    $cursor = $script:ProbeCursor
    $exception = Get-ProbeProperty -InputObject $ErrorRecord -Name 'Exception'
    $invocation = Get-ProbeProperty -InputObject $ErrorRecord -Name 'InvocationInfo'
    $line = [string](Get-ProbeProperty -InputObject $invocation -Name 'Line')
    $lines = @(
        '  stage              : ' + [string]$cursor.Stage,
        '  doing              : ' + [string]$cursor.Action,
        '  endpoint           : ' + $(if ([string]::IsNullOrWhiteSpace([string]$cursor.Endpoint)) { '(none)' } else { [string]$cursor.Endpoint }),
        '  detail             : ' + [string]$cursor.Detail,
        '  exception          : ' + $(if ($null -eq $exception) { 'unavailable' } else { $exception.GetType().FullName }),
        '  message            : ' + [string](Get-ProbeProperty -InputObject $exception -Name 'Message'),
        '  at line            : ' + [string](Get-ProbeProperty -InputObject $invocation -Name 'ScriptLineNumber'),
        '  statement          : ' + $line.Trim(),
        '  command            : ' + [string](Get-ProbeProperty -InputObject $invocation -Name 'MyCommand'))
    return ($lines -join "`r`n")
}

# ===========================================================================
# PREFLIGHT AND THE DISPOSABLE BUILD
# ===========================================================================
Set-ProbeStage -Stage 'preflight' -Action 'reading the build projections'
Write-Host ''
Write-Host 'PCCM - Phase 10 protection / table-structure probe' -ForegroundColor Cyan
Write-Host '=================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host 'This asks Excel one question. It measures nothing and decides' -ForegroundColor Yellow
Write-Host 'nothing: no timing, no baseline, no Gate-B result.' -ForegroundColor Yellow
Write-Host ''

$manifestPath = Join-Path $BuildDir 'stage_b_manifest.json'
$inspectPath  = Join-Path $BuildDir 'phase5_gate_b_inspection.json'
foreach ($required in @($manifestPath, $inspectPath)) {
    if (-not (Test-Path -LiteralPath $required)) {
        Write-Host ("$required not found. Run the Stage-A build first: " +
                    'python3 pccm/builder/build_stage_a.py') -ForegroundColor Red
        exit 1
    }
}
$manifest   = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
$inspection = Get-Content -LiteralPath $inspectPath  -Raw | ConvertFrom-Json
$watched = @(Get-ProbeWatchedTables -Manifest $manifest)
# THE COLLECTION BOUNDARY IS PROVED HERE, before Excel is started and before any
# identifier is coerced to a string. Probe Run 3 asked Excel for a worksheet
# whose name was five names joined by spaces; nothing between the manifest and
# the lookup had checked the shape.
try {
    $null = Assert-ProbeWatchedShape -Watched $watched -Manifest $manifest
} catch {
    Write-Host ''
    Write-Host 'REFUSED, BEFORE EXCEL WAS STARTED.' -ForegroundColor Red
    Write-Host ('the watched target records are malformed: ' + (Format-Err $_)) -ForegroundColor Red
    exit 2
}

Set-ProbeStage -Stage 'setup' -Action 'copying the build and running the Stage-B bootstrap'
$stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
$tempRoot = Join-Path $WorkDir ('pccm-phase10-protection-probe-' + $stamp)
$null = New-Item -ItemType Directory -Path $tempRoot -Force
Copy-Item -LiteralPath (Join-Path $BuildDir ([string]$manifest.stage_a_filename)) -Destination $tempRoot
Copy-Item -LiteralPath $manifestPath -Destination $tempRoot
Copy-Item -LiteralPath $inspectPath  -Destination $tempRoot
Copy-Item -LiteralPath (Join-Path $BuildDir 'vba') -Destination $tempRoot -Recurse

$script:ProbePath = Join-Path $tempRoot ('phase10_protection_probe_' + $stamp + '.log')
$stageBPath = Join-Path $tempRoot ([string]$manifest.stage_b_filename)

& (Join-Path $scriptDir 'build_stage_b.ps1') -BuildDir $tempRoot -Force
if (($LASTEXITCODE -ne 0) -or (-not (Test-Path -LiteralPath $stageBPath))) {
    Write-Host ''
    Write-Host 'The Stage-B bootstrap did not complete. Nothing was probed.' -ForegroundColor Red
    exit 1
}

Write-ProbeLine 'PCCM - PHASE 10 PROTECTION / TABLE-STRUCTURE PROBE'
Write-ProbeLine '================================================='
Write-ProbeLine ''
Write-ProbeLine ('run started : ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
Write-ProbeLine ('workbook    : ' + $stageBPath)
Write-ProbeLine ''
Write-ProbeLine 'THE QUESTION'
Write-ProbeLine '------------'
Write-ProbeLine 'Benchmark Run 3 died on a ListRow.Delete() with "Table features'
Write-ProbeLine 'aren''t available because the sheet is protected." That call came from'
Write-ProbeLine 'PowerShell. Seven user commands make the SAME CLASS of call from VBA.'
Write-ProbeLine 'This asks whether the real commands can do their contracted work on'
Write-ProbeLine 'the protected workbook - and checks that the shape actually changed,'
Write-ProbeLine 'because an announcement of success is not a structural operation.'
Write-ProbeLine ''
Write-ProbeLine ('TABLES WATCHED FOR A STRUCTURAL EFFECT (' + [string]@($watched).Count + ' records)')
foreach ($entry in @($watched)) {
    # ONE BLOCK PER RECORD. A single line per collection is how five records
    # printed as one in Run 3 and nobody noticed until Excel refused the name.
    Write-ProbeLine ('  ' + [string]$entry.Key + ':')
    Write-ProbeLine ('    tab   = ' + [string]$entry.Sheet)
    Write-ProbeLine ('    table = ' + [string]$entry.Table)
}
Write-ProbeLine ''
# ===========================================================================
# THE SESSION
# ===========================================================================
$preExisting = @(Get-PreExistingExcelPids)
$excel = $null; $workbooks = $null; $wb = $null; $excelIdentity = $null; $rel = $null
$resolution = $null

# THE VERDICT STARTS INCONCLUSIVE AND IS ONLY EVER MOVED BY THE ANSWER ITSELF.
# Every failure path leaves it where it is, so a probe that broke cannot say
# anything about production.
$verdict = 'INCONCLUSIVE'
$verdictReason = 'the probe did not reach its conclusion'
$outcomes = New-Object System.Collections.ArrayList
# FOUR STATES, AND 'NOT ATTEMPTED' IS ONE OF THEM.
#
# Probe Run 3 printed "UserInterfaceOnly permits code VALUE writes: False" after
# failing before the control ever ran. False reads as "blocking was observed",
# which is a claim about Excel that nothing had tested. A capability that was
# never exercised is NOT KNOWN, and the boolean is emitted only when it is.
$controlResult = 'NOT ATTEMPTED'
$controlDetail = 'the probe did not reach the locked-cell control'

try {
    Set-ProbeStage -Stage 'setup' -Action 'starting an owned Excel instance'
    $excel = New-Object -ComObject Excel.Application
    $excelIdentity = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    Set-ProbeStage -Stage 'setup' -Action 'opening the protected workbook'
    $workbooks = $excel.Workbooks
    $wb = $workbooks.Open($stageBPath)

    # --- is protection actually in force? ---------------------------------
    Set-ProbeStage -Stage 'protection' -Action 'reading the protection state as the workbook opened'
    $opened = Get-ProbeProtectionState -Workbook $wb -Where 'as the workbook opened'
    Write-ProbeLine 'PROTECTION STATE AS THE WORKBOOK OPENED'
    Write-ProbeLine '---------------------------------------'
    Write-ProbeLine ('  ' + (Format-ProbeProtection -State $opened))
    Set-ProbeStage -Stage 'protection' -Action 'asking the protection owner for its own answer'
    Write-ProbeLine ('  the owner reports : ' + [string]$excel.Run('ProtectionIsApplied'))
    Write-ProbeLine ''

    $protectionInForce = ([bool](([int]$opened.Protected -eq [int]$opened.Total) -and
                                 ([int]$opened.Total -gt 0)))

    # --- resolve every watched sheet and table, ONCE ----------------------
    # A REQUIRED SHEET OR TABLE THAT CANNOT BE RESOLVED IS A PROBE FAILURE, and
    # it throws here rather than being substituted with index 1 or skipped. The
    # resolved identifiers are evidence, printed below.
    Set-ProbeStage -Stage 'resolve' -Action 'resolving every watched worksheet and table'
    $resolution = Resolve-ProbeTargets -Workbook $wb -Watched $watched
    # AND THE RESOLVED SET IS THE SAME SIZE AS THE DECLARED ONE. A resolution
    # that lost or merged a target would otherwise be measured happily.
    $null = Assert-ProbeWatchedShape -Watched $resolution.Targets -Manifest $manifest
    Write-ProbeLine 'RESOLVED TARGETS'
    Write-ProbeLine '----------------'
    Write-ProbeLine ('  ' + [string]@($resolution.Targets).Count + ' targets resolved')
    foreach ($target in @($resolution.Targets)) {
        Write-ProbeLine ('  ' + [string]$target.Key + ':')
        Write-ProbeLine ('    tab      = ' + [string]$target.Sheet)
        Write-ProbeLine ('    codename = ' + [string]$target.CodeName)
        Write-ProbeLine ('    table    = ' + [string]$target.Table)
    }
    Write-ProbeLine ''

    if (-not $protectionInForce) {
        Write-ProbeLine 'THE PROBE STOPS HERE, INCONCLUSIVE: the workbook did not come back'
        Write-ProbeLine 'protected, so nothing below would be a statement about protected'
        Write-ProbeLine 'behaviour. That is itself a Workbook_Open protection-lifecycle'
        Write-ProbeLine 'finding, and it is reported as one rather than measured around.'
        Write-ProbeLine ''
        $verdictReason = 'the workbook was not protected when it opened'
    } else {
        # --- CONTROL: a value written to a locked cell ---------------------
        Set-ProbeStage -Stage 'control' -Action 'writing a VALUE to a locked cell on a protected sheet'
        Write-ProbeLine 'CONTROL - CAN CODE WRITE A VALUE TO A LOCKED CELL?'
        Write-ProbeLine '--------------------------------------------------'
        $control = Invoke-ProbeLockedCellControl -Workbook $wb -Resolution $resolution
        $controlResult = [string]$control.Result
        $controlDetail = [string]$control.Detail
        foreach ($line in @($control.Lines)) { Write-ProbeLine ('  ' + $line) }
        Write-ProbeLine ''
        if ($controlResult -eq 'INCONCLUSIVE') {
            # A PROBE THAT CANNOT RESTORE WHAT IT CHANGED HAS ESTABLISHED NOTHING.
            # The production question is not asked over an untrustworthy instrument.
            $verdictReason = ('the locked-cell control could not be trusted: ' + $controlDetail +
                              '. The production question was not asked.')
            Write-ProbeLine 'THE PROBE STOPS HERE, INCONCLUSIVE: the control did not settle.'
            Write-ProbeLine ''
        } else {

        # --- THE QUESTION: the real commands -------------------------------
        Write-ProbeLine 'THE QUESTION - CAN THE REAL COMMANDS DO THEIR STRUCTURAL WORK?'
        Write-ProbeLine '--------------------------------------------------------------'
        Write-ProbeLine 'Each is the accepted public entry point, invoked once, exactly as its'
        Write-ProbeLine 'button invokes it. What is reported is what the workbook announced AND'
        Write-ProbeLine 'what happened to the tables.'
        Write-ProbeLine ''

        # Apply Timeline is FIRST because it is unconditionally structural: a
        # fresh workbook has no year columns, so any timeline adds ListColumns to
        # three grids. It is also the first button a real user presses.
        # SETTING PRECONDITIONS IS NOT INVOKING THE ENDPOINT, and the stage wording
        # now says so in its own words. Run 4 died here, in fixture preparation,
        # and a transcript that called this "setting the timeline inputs" under
        # stage 'endpoint' was one careless reading away from implying that
        # PCCM_ApplyTimeline had run. It had not.
        Set-ProbeStage -Stage 'endpoint' `
            -Action 'SETTING ENDPOINT PRECONDITIONS: writing the declared inputs (production NOT invoked)' `
            -Endpoint 'PCCM_ApplyTimeline'
        $declaredInputs = @(Get-ProbeDeclaredInputs -Inspection $inspection)
        $growYears = [int](Get-ProbeGrowthYears)
        $shrinkYears = [int](Get-ProbeShrinkYears)
        Set-ProbeDeclaredInputs -Workbook $wb -Inputs $declaredInputs

        Set-ProbeStage -Stage 'endpoint' `
            -Action 'VERIFYING ENDPOINT PRECONDITIONS: reading the declared inputs back (production NOT invoked)' `
            -Endpoint 'PCCM_ApplyTimeline'
        $inputProblems = @(Test-ProbeDeclaredInputs -Workbook $wb -Inputs $declaredInputs)

        Write-ProbeLine 'Endpoint preconditions - the declared inputs, written and read back:'
        foreach ($entry in $declaredInputs) {
            Write-ProbeLine ('  ' + [string]$entry.Key + '  ' + [string]$entry.DefinedName +
                             ' = ' + [string]$entry.Value + '  (System.Double, for ' +
                             [string]$entry.Endpoint + ')')
        }
        if (@($inputProblems).Count -gt 0) {
            foreach ($problem in $inputProblems) { Write-ProbeLine ('  PROBLEM: ' + $problem) }
            # PROBE INSTRUMENTATION FAILURE. The outer catch leaves the verdict
            # INCONCLUSIVE and names the stage, and no endpoint is INVOKED -
            # $invoked is only ever set beside Application.Run.
            throw ('the declared inputs could not be established, so no production endpoint ' +
                   'was invoked: ' + ($inputProblems -join '; '))
        }
        Write-ProbeLine '  all read back as System.Double with the expected values'
        Write-ProbeLine ''

        $timeline = Invoke-ProbeEndpoint -Excel $excel -Workbook $wb `
            -Endpoint 'PCCM_ApplyTimeline' -Resolution $resolution
        $null = $outcomes.Add($timeline)
        Write-ProbeOutcome -Outcome $timeline -Expectation `
            ('ListColumns.Add on three grids. A fresh workbook has no year columns, so ' +
             'this fires for ANY timeline and is not capacity-dependent.')

        # --- CALCULATE, AND IT RUNS HERE FOR A REASON ----------------------
        # BEFORE THE ADD COMMANDS, because that is the only ordering in which
        # Calculate's own prerequisites are legitimately satisfiable without
        # inventing business data.
        #
        # AddDriver WRITES A PERMANENT ID into the row it adds. modCalcResolve's
        # ReadRegister reads every row whose id column is non-blank, so after an
        # Add the register holds one identified driver with every other field
        # empty - and Calculate then refuses on a blank required field. Filling
        # those fields would mean the probe manufacturing a cost line, which is
        # a fixture this batch is not authorised to invent.
        #
        # AN EMPTY DRIVER SET IS VALID, AND THAT IS THE CONTRACT'S OWN WORDING:
        # "A workbook with no Cost Lines and no Risks resolves to zero drivers
        # and an empty reference set; no minimum-driver rule is invented here."
        # With zero drivers no currency and no inflation profile is REFERENCED,
        # so ResolveFxRates and ResolveInflationRates have nothing to resolve and
        # the deliberately-blank inflation grid is never read. The applied
        # timeline and the discount rate are therefore the whole prerequisite.
        #
        # THIS IS ALSO THE USER'S OWN FIRST CALCULATE: apply a timeline, press
        # Calculate. It is not a contrived state.
        $calcTables = @(Get-ProbeCalcTables -Inspection $inspection)
        $perYearTables = @(Get-ProbePerYearCalcTables -CalcTables $calcTables)

        $growRound = Invoke-ProbeCalculateRound -Excel $excel -Workbook $wb `
            -Resolution $resolution -CalcTables $calcTables -PerYearTables $perYearTables `
            -ExpectedRows $growYears -Label 'growth'
        $calculate = $growRound.Outcome
        $null = $outcomes.Add($calculate)
        Write-ProbeOutcome -Outcome $calculate -Expectation `
            ('ResizeBody adds ListRows on the _Calc tables. Judged by the TABLE SHAPES ' +
             'below, not by the announcement.')
        Write-ProbeLine ('  _Calc table shapes, before -> after (' + [string]$growRound.Label + '):')
        foreach ($line in @($growRound.Lines)) { Write-ProbeLine $line }
        Write-ProbeLine ('  growth evidence: ' + [string]$growRound.Proof)
        Write-ProbeLine ''

        # --- THE SHRINK ROUND: THE DELETE PATH -----------------------------
        # WHY THIS ROUND EXISTS. Benchmark Run 3 died on a ListRow.Delete() with
        # "Table features aren't available because the sheet is protected". Every
        # successful round above proves the ADD direction: ListColumns.Add on the
        # grids, ListRows.Add in ResizeBody. Neither proves the DELETE direction,
        # and the delete is the call that actually failed.
        #
        # ONE INPUT CHANGES, AND IT CHANGES THE SAME WAY EVERY OTHER INPUT DID:
        # a genuine Double through the accepted Set-NamedValue, read back and
        # type-checked before the endpoint is touched. Duration 1 is inside the
        # accepted bound (modTimeline requires 1 <= d <= LIMIT_MAX_YEAR_COLUMNS).
        #
        # WHAT IT MUST PRODUCE. Shrinking from 3 project years to 1 drives
        # modProfiling.SetYearColumns and modInflation.SetYearColumns down their
        # `ListColumns(...).Delete` loops on all three grids, and then
        # ResizeBody down its `ListRows(...).Delete` loop on the per-year _Calc
        # tables. Both are the production structural window's delete path.
        Set-ProbeStage -Stage 'endpoint' `
            -Action ('SETTING ENDPOINT PRECONDITIONS: shrinking the applied duration to ' +
                     [string]$shrinkYears + ' (production NOT invoked)') `
            -Endpoint 'PCCM_ApplyTimeline'
        $shrinkInputs = @(Get-ProbeShrinkInputs -Inspection $inspection)
        Set-ProbeDeclaredInputs -Workbook $wb -Inputs $shrinkInputs

        Set-ProbeStage -Stage 'endpoint' `
            -Action 'VERIFYING ENDPOINT PRECONDITIONS: reading the shrink input back (production NOT invoked)' `
            -Endpoint 'PCCM_ApplyTimeline'
        $shrinkProblems = @(Test-ProbeDeclaredInputs -Workbook $wb -Inputs $shrinkInputs)
        Write-ProbeLine 'SHRINK ROUND - the delete path Benchmark Run 3 died on'
        Write-ProbeLine '------------------------------------------------------'
        foreach ($entry in $shrinkInputs) {
            Write-ProbeLine ('  ' + [string]$entry.Key + '  ' + [string]$entry.DefinedName +
                             ' = ' + [string]$entry.Value + '  (System.Double, for ' +
                             [string]$entry.Endpoint + ')')
        }
        if (@($shrinkProblems).Count -gt 0) {
            foreach ($problem in $shrinkProblems) { Write-ProbeLine ('  PROBLEM: ' + $problem) }
            throw ('the shrink input could not be established, so the delete path was NOT ' +
                   'exercised: ' + ($shrinkProblems -join '; '))
        }
        Write-ProbeLine '  read back as System.Double with the expected value'

        $shrinkTimeline = Invoke-ProbeEndpoint -Excel $excel -Workbook $wb `
            -Endpoint 'PCCM_ApplyTimeline' -Resolution $resolution
        $null = $outcomes.Add($shrinkTimeline)
        Write-ProbeOutcome -Outcome $shrinkTimeline -Expectation `
            ('ListColumns.Delete on all three grids. The year columns must come DOWN from ' +
             'the ' + [string]$growYears + '-year state to the ' + [string]$shrinkYears +
             '-year state; a shape that merely CHANGED is not delete evidence.')

        $growDirection = Get-ProbeColumnDirection -Before $timeline.ShapesBefore `
            -After $timeline.ShapesAfter -Resolution $resolution
        $shrinkDirection = Get-ProbeColumnDirection -Before $shrinkTimeline.ShapesBefore `
            -After $shrinkTimeline.ShapesAfter -Resolution $resolution
        Write-ProbeLine ('  grids that GAINED columns on the ' + [string]$growYears +
                         '-year apply : ' + ((@($growDirection.Grew) -join ', ')))
        Write-ProbeLine ('  grids that LOST columns on the ' + [string]$shrinkYears +
                         '-year apply  : ' + ((@($shrinkDirection.Shrank) -join ', ')))

        # EVERY GRID THAT GREW MUST SHRINK. Deriving the expectation from what was
        # actually observed growing keeps this predictive without a literal, and
        # refuses a partial shrink that happened to touch one table.
        $missedShrink = @()
        foreach ($key in @($growDirection.Grew)) {
            if (@($shrinkDirection.Shrank) -notcontains $key) { $missedShrink += $key }
        }
        $columnDeleteProof = 'NOT ESTABLISHED'
        if ([string]$shrinkTimeline.Outcome -eq 'SUCCEEDED') {
            if ((@($growDirection.Grew).Count -gt 0) -and (@($missedShrink).Count -eq 0)) {
                $columnDeleteProof = 'OBSERVED'
            } else {
                $columnDeleteProof = 'CONTRADICTED'
                if (@($missedShrink).Count -gt 0) {
                    Write-ProbeLine ('  BUT these grids grew and did NOT shrink: ' +
                                     ((@($missedShrink) -join ', ')))
                }
                if (@($growDirection.Grew).Count -eq 0) {
                    Write-ProbeLine '  BUT no grid gained columns in the growth round, so there is nothing to have deleted'
                }
            }
        }
        Write-ProbeLine ('  ListColumns.Delete evidence: ' + $columnDeleteProof)
        Write-ProbeLine ''

        $shrinkRound = Invoke-ProbeCalculateRound -Excel $excel -Workbook $wb `
            -Resolution $resolution -CalcTables $calcTables -PerYearTables $perYearTables `
            -ExpectedRows $shrinkYears -Label 'shrink'
        $shrinkCalculate = $shrinkRound.Outcome
        $null = $outcomes.Add($shrinkCalculate)
        Write-ProbeOutcome -Outcome $shrinkCalculate -Expectation `
            ('ResizeBody DELETES ListRows down to the ' + [string]$shrinkYears +
             '-year state. This is the ListRow.Delete() Benchmark Run 3 died on.')
        Write-ProbeLine ('  _Calc table shapes, before -> after (' + [string]$shrinkRound.Label + '):')
        foreach ($line in @($shrinkRound.Lines)) { Write-ProbeLine $line }

        # THE PER-YEAR TABLES MUST HAVE LOST ROWS, not merely landed on the right
        # number. Accepting 3 -> 3 would call an unchanged table delete evidence.
        $rowDeleteProof = 'NOT ESTABLISHED'
        if ([string]$shrinkCalculate.Outcome -eq 'SUCCEEDED') {
            $missedRowDelete = @()
            foreach ($table in @($perYearTables)) {
                if (@($shrinkRound.RowsDeleted) -notcontains [string]$table) {
                    $missedRowDelete += [string]$table
                }
            }
            if (([string]$shrinkRound.Proof -eq 'OBSERVED') -and
                (@($perYearTables).Count -gt 0) -and (@($missedRowDelete).Count -eq 0)) {
                $rowDeleteProof = 'OBSERVED'
            } else {
                $rowDeleteProof = 'CONTRADICTED'
                if (@($missedRowDelete).Count -gt 0) {
                    Write-ProbeLine ('  BUT these _Calc tables did not lose rows: ' +
                                     ((@($missedRowDelete) -join ', ')))
                }
            }
        }
        Write-ProbeLine ('  ListRows.Delete evidence: ' + $rowDeleteProof)

        # THE ADD DIRECTION, STATED AS ITS OWN CLASS rather than left implicit in
        # "the shape changed". Run 7 established it; naming it keeps the four
        # classes symmetrical so a future run cannot report three of them.
        $columnAddProof = 'NOT ESTABLISHED'
        if ([string]$timeline.Outcome -eq 'SUCCEEDED') {
            if (@($growDirection.Grew).Count -gt 0) { $columnAddProof = 'OBSERVED' }
            else { $columnAddProof = 'CONTRADICTED' }
        }

        # The two Calculate rounds summarised into the one word the criteria
        # block reports. OBSERVED only when BOTH rounds took their contracted
        # shape; a single CONTRADICTED round contradicts the whole claim.
        $calcStructuralProof = 'NOT ESTABLISHED'
        if (([string]$growRound.Proof -eq 'OBSERVED') -and
            ([string]$shrinkRound.Proof -eq 'OBSERVED')) {
            $calcStructuralProof = 'OBSERVED'
        } elseif (([string]$growRound.Proof -eq 'CONTRADICTED') -or
                  ([string]$shrinkRound.Proof -eq 'CONTRADICTED')) {
            $calcStructuralProof = 'CONTRADICTED'
        }
        Write-ProbeLine ''

        $addCost = Invoke-ProbeEndpoint -Excel $excel -Workbook $wb `
            -Endpoint 'PCCM_AddCostLine' -Resolution $resolution
        $null = $outcomes.Add($addCost)
        Write-ProbeOutcome -Outcome $addCost -Expectation `
            ('SyncRows runs on every add and can reshape the Cost Profiling grid. The ' +
             'register itself grows only PAST its 25 reserved rows, so on a fresh workbook ' +
             'ListRows.Add does not fire and NO shape change is the correct observation - ' +
             'it settles endpoint functionality under protection, NOT capacity expansion.')

        $addRisk = Invoke-ProbeEndpoint -Excel $excel -Workbook $wb `
            -Endpoint 'PCCM_AddRisk' -Resolution $resolution
        $null = $outcomes.Add($addRisk)
        Write-ProbeOutcome -Outcome $addRisk -Expectation `
            ('the same path on the Risk Register and the Risk Profiling grid, and the same ' +
             'reading: no shape change on a fresh workbook is expected, and does not settle ' +
             'ListRows.Add capacity expansion either.')

        # --- THE VERDICT ---------------------------------------------------
        # BLOCKED and FINE are the only two conclusions, and each needs its own
        # evidence. Anything else stays INCONCLUSIVE, which is where it started.
        Set-ProbeStage -Stage 'verdict' -Action 'weighing the outcomes'
        $notSucceeded = @(@($outcomes) | Where-Object { [string]$_.Outcome -ne 'SUCCEEDED' })
        $lostProtection = @(@($outcomes) | Where-Object {
            ([int]$_.ProtectedAfter -ne [int]$_.TotalSheets) -or
            ([int]$_.ProtectedBefore -ne [int]$_.TotalSheets) -or
            (-not [bool]$_.StructureBefore) -or (-not [bool]$_.StructureAfter) })
        $structural = @(@($outcomes) | Where-Object { [bool]$_.StructuralEffect })
        # BLOCKED IS DECIDED BY THIS LIST AND NOT BY $notSucceeded. A refusal
        # only counts when Excel itself said the sheet's protection is what
        # stopped a table operation.
        $protectionBlocked = @(@($outcomes) | Where-Object { Test-ProbeProtectionBlocked -Outcome $_ })
        $refusedForOtherReasons = @(@($outcomes) | Where-Object {
            ([string]$_.Outcome -ne 'SUCCEEDED') -and (-not (Test-ProbeProtectionBlocked -Outcome $_)) })

        if (@($protectionBlocked).Count -gt 0) {
            $verdict = 'PRODUCTION IS BLOCKED BY PROTECTION'
            $verdictReason = ([string]@($protectionBlocked).Count + ' of ' + [string]@($outcomes).Count +
                              ' production endpoints were refused by Excel with ' +
                              [char]34 + 'table features aren' + [char]39 +
                              't available because the sheet is protected' + [char]34 + ': ' +
                              ((@($protectionBlocked) | ForEach-Object { [string]$_.Endpoint }) -join ', '))
            if (@($refusedForOtherReasons).Count -gt 0) {
                # NAMED, AND NAMED AS NOT COUNTING. Run 5's PCCM_Calculate refused
                # because the applied timeline was pending; recording it silently
                # among the blocked ones is what made that verdict too broad.
                $verdictReason = ($verdictReason + '. Also refused, but NOT for protection reasons ' +
                                  'and NOT counted here: ' +
                                  ((@($refusedForOtherReasons) | ForEach-Object { [string]$_.Endpoint }) -join ', '))
            }
        } elseif (@($notSucceeded).Count -gt 0) {
            # REFUSED, BUT NOT BY PROTECTION. That is a real result and it is
            # reported - it is simply not an answer to THIS question.
            $verdictReason = ([string]@($notSucceeded).Count + ' of ' + [string]@($outcomes).Count +
                              ' production endpoints did not succeed, but none of them was ' +
                              'refused for a protection reason, so protection is not what ' +
                              'stopped them and the question is not settled: ' +
                              ((@($notSucceeded) | ForEach-Object {
                                  [string]$_.Endpoint + ' (' + [string]$_.Outcome + ')' }) -join ', '))
        } elseif (@($structural).Count -lt 1) {
            $verdictReason = ('every endpoint announced success but no watched table ever ' +
                              'changed shape, so the structural effect was not observed and ' +
                              'the question is not settled')
        } elseif (@($lostProtection).Count -gt 0) {
            $verdictReason = ('the commands succeeded but protection was not in force after ' +
                              [string]@($lostProtection).Count + ' of them, so they were not ' +
                              'a test of protected behaviour')
        } elseif (($columnDeleteProof -ne 'OBSERVED') -or ($rowDeleteProof -ne 'OBSERVED')) {
            # THE DELETE PATH IS THE ONE BENCHMARK RUN 3 DIED ON. Every ADD
            # direction succeeding says nothing about ListColumns.Delete or
            # ListRow.Delete, and calling that run a harness defect on ADD
            # evidence alone is precisely the overreach this branch refuses.
            $verdictReason = ('the ADD direction was observed but the DELETE direction was ' +
                              'not: ListColumns.Delete is ' + $columnDeleteProof +
                              ' and ListRows.Delete is ' + $rowDeleteProof +
                              '. Benchmark Run 3 died on a ListRow.Delete(), so the question ' +
                              'is not settled without it')
        } elseif ($growRound.Proof -ne 'OBSERVED') {
            $verdictReason = ('the growth round did not observe its contracted _Calc shape (' +
                              [string]$growRound.Proof + '), so the ADD direction is not settled')
        } elseif ($shrinkRound.Proof -ne 'OBSERVED') {
            # THE PROBE'S OWN PRINCIPLE, APPLIED TO ITSELF. An announcement of
            # success is not proof of a structural operation, so a Calculate that
            # announced success while the _Calc tables did not take the shape
            # their row rule contracts cannot carry the run to FINE.
            $verdictReason = ('Calculate announced success but the _Calc tables did not take ' +
                              'the shape their row rule contracts, so the structural work was ' +
                              'not observed and the question is not settled')
        } else {
            $verdict = 'PRODUCTION IS FINE UNDER PROTECTION'
            $verdictReason = ('LISTCOLUMN ADD: OBSERVED; LISTCOLUMN DELETE: OBSERVED; ' +
                              'LISTROW GROWTH: OBSERVED; LISTROW DELETE: OBSERVED. ' +
                              'changed shape, and every sheet was protected before and ' +
                              'after each one - so Benchmark Run 3 was a HARNESS defect only')
        }

        # --- WHAT THE POST-FIX RUN HAS TO ESTABLISH ------------------------
        # SEPARATED FROM THE VERDICT ON PURPOSE. The reconciliation's own
        # acceptance question is narrower than "is production fine": it is
        # whether protection stopped blocking STRUCTURAL INITIALISATION. Making
        # that a fourth verdict word would invite it to be read as acceptance,
        # and PCCM_Calculate refusing for want of business inputs is not a
        # protection result either way. So the five criteria are reported as
        # themselves, each from evidence this run actually holds.
        Set-ProbeStage -Stage 'verdict' -Action 'reporting the structural-initialisation criteria'
        $timelineOutcome = @(@($outcomes) | Where-Object { [string]$_.Endpoint -eq 'PCCM_ApplyTimeline' })
        $driverOutcomes = @(@($outcomes) | Where-Object {
            ([string]$_.Endpoint -eq 'PCCM_AddCostLine') -or ([string]$_.Endpoint -eq 'PCCM_AddRisk') })
        $criteria = New-Object System.Collections.ArrayList
        $null = $criteria.Add([pscustomobject]@{ Key = 'A'; Text = 'the workbook opened protected'
            Met = ([bool]$protectionInForce) })
        $null = $criteria.Add([pscustomobject]@{ Key = 'B'; Text = 'a locked-cell code VALUE write is still permitted'
            Met = ([bool]($controlResult -eq 'SUCCEEDED')) })
        $null = $criteria.Add([pscustomobject]@{ Key = 'C'; Text = 'PCCM_ApplyTimeline SUCCEEDED, created year columns, and left protection applied'
            Met = ([bool]((@($timelineOutcome).Count -eq 1) -and
                          ([string]@($timelineOutcome)[0].Outcome -eq 'SUCCEEDED') -and
                          [bool]@($timelineOutcome)[0].StructuralEffect -and
                          ([int]@($timelineOutcome)[0].ProtectedAfter -eq [int]@($timelineOutcome)[0].TotalSheets))) })
        $null = $criteria.Add([pscustomobject]@{ Key = 'D'; Text = 'Add Cost Line and Add Risk remain functional'
            Met = ([bool]((@($driverOutcomes).Count -eq 2) -and
                          (@(@($driverOutcomes) | Where-Object { [string]$_.Outcome -ne 'SUCCEEDED' }).Count -eq 0))) })
        $null = $criteria.Add([pscustomobject]@{ Key = 'E'; Text = 'no endpoint left a worksheet unprotected'
            Met = ([bool](@($lostProtection).Count -eq 0)) })

        Write-ProbeLine ''
        Write-ProbeLine 'STRUCTURAL-INITIALISATION CRITERIA'
        Write-ProbeLine '----------------------------------'
        foreach ($criterion in @($criteria)) {
            $mark = 'NOT MET'
            if ([bool]$criterion.Met) { $mark = 'MET    ' }
            Write-ProbeLine ('  ' + [string]$criterion.Key + '. ' + $mark + '  ' + [string]$criterion.Text)
        }
        $unmet = @(@($criteria) | Where-Object { -not [bool]$_.Met })
        if (@($unmet).Count -eq 0) {
            Write-ProbeLine '  STRUCTURAL INITIALISATION: PROVEN on this run.'
        } else {
            Write-ProbeLine ('  STRUCTURAL INITIALISATION: NOT PROVEN - ' +
                             ((@($unmet) | ForEach-Object { [string]$_.Key }) -join ', ') + ' not met.')
        }
        Write-ProbeLine ''
        Write-ProbeLine 'STRUCTURAL EVIDENCE BY CLASS'
        Write-ProbeLine '----------------------------'
        Write-ProbeLine ('  LISTCOLUMN ADD    : ' + $columnAddProof)
        Write-ProbeLine ('  LISTCOLUMN DELETE : ' + $columnDeleteProof)
        Write-ProbeLine ('  LISTROW GROWTH    : ' + [string]$growRound.Proof)
        Write-ProbeLine ('  LISTROW DELETE    : ' + $rowDeleteProof)
        Write-ProbeLine ('  CALCULATE STRUCTURAL EVIDENCE: ' + $calcStructuralProof +
                         '  (OBSERVED = the _Calc per-project-year tables hold the applied ' +
                         'duration; CONTRADICTED = it announced success without them; ' +
                         'NOT ESTABLISHED = Calculate did not succeed, which requires no shape)')
        Write-ProbeLine 'This is NOT final acceptance and NOT a benchmark baseline.'
        Write-ProbeLine ''
        }
    }
    $excel.Run('PCCM_AutomationEnd') | Out-Null
} catch {
    # A PROBE FAILURE IS NEVER A STATEMENT ABOUT PRODUCTION. The verdict is left
    # exactly where it was.
    Write-ProbeLine ''
    Write-ProbeLine 'THE PROBE RAISED'
    Write-ProbeLine '----------------'
    Write-ProbeLine (Format-ProbeFailure $_)
    Write-ProbeLine ''
    $verdictReason = ('the probe itself failed in stage ' + [string]$script:ProbeCursor.Stage +
                      ' while ' + [string]$script:ProbeCursor.Action)
} finally {
    # THE HELD TARGETS GO FIRST, leaf before parent, before the workbook closes.
    try { Release-ProbeTargets -Resolution $resolution }
    catch { Write-ProbeLine ('the resolved targets could not be released: ' + (Format-Err $_)) }
    $resolution = $null
    $rel = New-ReleaseLedger 'phase-10 protection probe'
    try {
        if ($null -ne $wb) {
            # NEVER SAVED. The disposable copy is discarded.
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
            Write-ProbeLine (Invoke-EmergencyExcelCleanup -Identity $excelIdentity `
                -Label 'phase-10 protection probe')
        }
    } catch {
        Write-ProbeLine ('Shutdown raised: ' + (Format-Err $_))
    }
    Write-ProbeLine 'SHUTDOWN'
    Write-ProbeLine (Format-ReleaseLedger $rel)
    $transient = @(Get-TransientFailures)
    if (@($transient).Count -gt 0) {
        Write-ProbeLine ('transient COM release failures: ' + ($transient -join '; '))
    } else {
        Write-ProbeLine 'every transient COM object released cleanly'
    }
}

Write-ProbeLine ''
Write-ProbeLine 'CONTROL'
Write-ProbeLine '-------'
Write-ProbeLine ('  status : ' + $controlResult)
Write-ProbeLine ('  detail : ' + $controlDetail)
if ($controlResult -eq 'SUCCEEDED') {
    Write-ProbeLine '  UserInterfaceOnly code-value-write capability: CONFIRMED'
} elseif ($controlResult -eq 'REFUSED') {
    Write-ProbeLine '  UserInterfaceOnly code-value-write capability: REFUSED BY EXCEL'
} else {
    # NOT ATTEMPTED and INCONCLUSIVE are both "nothing was learned", and neither
    # is written as a boolean. There is no False to be misread.
    Write-ProbeLine '  UserInterfaceOnly code-value-write capability: NOT TESTED'
}
Write-ProbeLine '  It is a SEPARATE capability from permission to perform a ListObject'
Write-ProbeLine '  structural operation, and it settles nothing about one.'
Write-ProbeLine ''
Write-ProbeLine 'VERDICT'
Write-ProbeLine '-------'
Write-ProbeLine ('  ' + $verdict)
Write-ProbeLine ('  ' + $verdictReason)
Write-ProbeLine ''
Write-ProbeLine 'THIS IS NOT A BASELINE, NOT A GATE-B RESULT AND NOT AN ACCEPTANCE RUN.'
Write-ProbeLine 'It is one question, asked once, so a correction is made against evidence'
Write-ProbeLine 'rather than against a reading of the documentation.'
Write-ProbeLine ''
Write-ProbeLine ('  log : ' + $script:ProbePath)

$kept = $script:ProbePath
if (-not $KeepArtifacts) {
    try {
        $kept = Join-Path $WorkDir ('pccm-phase10-protection-probe-' + $stamp + '.log')
        Copy-Item -LiteralPath $script:ProbePath -Destination $kept -Force
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    } catch { $kept = $script:ProbePath }
}
Write-Host ''
Write-Host ('Probe log kept at ' + $kept) -ForegroundColor Yellow

# THE PROCESS SAYS IT TOO. Only a settled question exits 0.
if ($verdict -eq 'INCONCLUSIVE') { exit 2 }
exit 0
