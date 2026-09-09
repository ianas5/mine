<#
.SYNOPSIS
    PCCM Phase 9 - P9-1: the Model Check surface, in real Excel.

.DESCRIPTION
    THIS IS NOT A PHASE-8 SUITE AND MUST NOT BE READ AS ONE. The P8-1, P8-2,
    P8-3 and P8-Z acceptances stand at their own commits and nothing here
    re-establishes or disturbs any of them. No chart is inspected, no annual
    table is reconciled, no sensitivity ranking is read. This runner asks the
    questions that belong to the Model Check sheet and to nothing else.

    WHAT ONLY EXCEL CAN ANSWER. Everything about the ORDERING, the counts, the
    precedence, the duplicate rule, the overflow arithmetic and the advisory
    boundary was settled on Linux against the sheet's own formulas. Four things
    were not, and they are why this file exists:

      1. DO THE ACCESSORS ANSWER FROM A CELL? Six worksheet functions are called
         from the readings block. Two of the accepted adapters reach a path that
         P8-1 proved could show #VALUE! when it persisted. Every reading cell
         must hold a value, not an error.

      2. DOES THE ANCHOR MAKE AN UN-VOLATILE ACCESSOR LIVE? PCCM_StructuralReport
         is not Application.Volatile and is not ours to make so: it is shared
         with command paths that have no reason to re-run on every calculation
         cycle. The sheet declares the dependency instead, by making that cell a
         dependent of the volatile `evaluated` heartbeat. If that does not work,
         a structural fault created after the workbook opened will not appear,
         and the sheet would be quietly reporting an old workbook.

      3. IS THE SURFACE ACTUALLY READ-ONLY? _Calc C19:C20 is the persisted
         calculation status pair. A recalculation of Model Check must not
         rewrite it. That is the whole worksheet-safety claim, and it is
         observable: read the pair, recalculate, read it again.

      4. DOES THE REGISTER FILL AND EMPTY THE WAY IT WAS BUILT TO? An unused
         slot must be #N/A - never blank, never 0 - and the counts above it must
         equal the rows below it, in every state the model passes through.

      5. IS THE ACTIONABLE ERROR'S SUBJECT THE RIGHT DRIVER? P9-2B threads the
         offending permanent id out of the owner that refuses, so scenario F can
         assert EQUALITY against the id the runner itself invalidated rather
         than looking for it inside a sentence.

      6. IS THE ACTIONABLE ERROR'S REASON THE LIVE ONE? P9-2A points the
         calculation ERROR at the sentence the CURRENT preparation wrote. In
         scenario F the persisted last attempt describes a SUCCESSFUL
         calculation of a model that no longer exists, so a row that reached for
         history instead of the live model reads plausibly and is wrong - which
         is exactly the shape of defect a static control cannot see and a run
         can.

    HOW THE STATES ARE REACHED. Through production, and only through production.
    The accepted W4 fixture is applied, then PCCM_Calculate, then
    PCCM_RunSimulation, then ordinary edits: an iteration count moved across the
    recommendation, a driver value changed, a three-point ordering broken, an
    unkeyed register row added. Nothing here writes a state word, a status cell
    or a Model Check cell.

    NOT ONE ADDRESS IS TYPED. Every row, column, vocabulary word and threshold
    comes from build/phase9_model_check_inspection.json. A hand-written address
    is a second declaration of something a projection already owns, and P7-4 is
    what that costs.

.NOTES
    THIS FILE HAS NOT RUN. There is no Windows and no Excel where it was
    written, and nothing in it claims otherwise.
#>
[CmdletBinding()]
param(
    [string]$BuildDir
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# DEFINITION-ONLY AT TOP LEVEL. Dot-sourcing defines functions and script
# variables and runs no scenario.
. (Join-Path $scriptDir 'com_lifecycle.ps1')
. (Join-Path $scriptDir 'phase5_gate_b_scenarios.ps1')
. (Join-Path $scriptDir 'phase6_gate_b_scenarios.ps1')

# EVERY PATH DERIVED FROM WHERE THE SCRIPT IS, never from the working directory:
#   bootstrap/windows  ->  pccm  ->  the repository root
$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
$repoRoot = Split-Path -Parent $pccmRoot
if ([string]::IsNullOrWhiteSpace($BuildDir)) { $BuildDir = Join-Path $pccmRoot 'build' }

# EVERY SCRIPT-SCOPE VARIABLE THE HELPERS READ, INITIALISED HERE. P8-Z run 2
# died because a copied function carried a dependency on script state that
# nothing had created; the lesson is that copying a function copies what it
# reads.
$script:P9Lines = New-Object System.Collections.ArrayList
# THE REPORT PATH, EMPTY UNTIL THERE IS ONE. Write-P9Line tests it before using
# it, so lines written before the temp directory exists are held and flushed.
$script:P9Path = ''
$script:P9Checks = New-Object System.Collections.ArrayList
$script:P9Residual = New-Object System.Collections.ArrayList

$script:P9ErrorCodes = @{
    -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!'
    -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!'
    -2146826246 = '#N/A'
}


# ===========================================================================
# THE WORKBOOK HELPERS THIS RUNNER NEEDS, DEFINED HERE
# ===========================================================================
# NOT DOT-SOURCED FROM phase8_pz_zero_variance.ps1. That file is a SCENARIO,
# not a definition library: dot-sourcing it would run a Phase-8 acceptance
# inside a Phase-9 runner. So the four table helpers and the two named-value
# helpers are defined here, with every script variable they read already
# created above - which is the P8-Z run-2 lesson stated as code rather than as
# a comment.

# `Write-RowObject` IS HERE BECAUSE OF A TRANSITIVE DEPENDENCY, and Windows run 1
# is what it cost to leave it out. NOTHING IN THIS FILE CALLS IT.
# `Get-Phase5TypedTableBody` does, and that function lives in
# `phase5_gate_b_scenarios.ps1`, which this runner DOES dot-source - while its
# own definition lives in `phase4_functional_test.ps1`, which this runner
# deliberately does not, because that file is a Phase-4 driver and not a
# definition library. In the accepted Gate-B runs the Phase-4 driver dot-sources
# the scenarios file, so the helper is in scope; reaching the scenarios file
# directly leaves the dependency unmet.
#
# THIS RUNNER REACHES IT THROUGH `Save-Phase5LockedFxSeed`, one line after the
# third prerequisite. W1 died the same way before a single check was recorded,
# the timing harness met the same gap the same way, and P9 run 1 died on it
# after five - which is why test_29 below now proves every command this runner
# names resolves, rather than trusting an assembly to have copied everything.
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

# THE SEVEN TABLE HELPERS THE DOT-SOURCED FIXTURE NEEDS, UNDER THEIR OWN NAMES.
# THIS RUNNER CALLS NONE OF THEM. `Set-Phase5Fixture` does - through
# Invoke-Phase5FixtureSteps, Reset-Phase5FxTable, Clear-Phase5Registers,
# Write-Phase5InflationRates and Invoke-Phase5AddDriverAndRequireSuccess - and
# that fixture lives in `phase5_gate_b_scenarios.ps1`, which this runner
# dot-sources while deliberately not dot-sourcing the Phase-4 driver the helpers
# live in.
#
# THE ASSEMBLY THAT BUILT THIS RUNNER COPIED THE THREE IT USED ITSELF AND
# RENAMED THEM INTO A P9 NAMESPACE - Set-P9TableCell, Add-P9BlankTableRow,
# Remove-P9TableRow below - and dropped the four it had no use for. The fixture
# needs all seven under the ORIGINAL names, so run 1 died on the first one it
# reached and seven more were waiting behind it. They are copied BYTE FOR BYTE
# from phase8_pz_zero_variance.ps1 so there is one behaviour rather than two,
# and test_29 below pins them to it.
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

function Set-P9TableCell {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$RowIndex,
          [int]$ColumnIndex, $Value)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $body = $null; $cell = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $body = $lo.DataBodyRange
        $cell = $body.Cells($RowIndex, $ColumnIndex)
        if ($null -eq $Value) {
            # A genuine blank, not zero. The two are different assumptions.
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

function Add-P9BlankTableRow {
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

function Remove-P9TableRow {
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

function Get-P9Register {
    param($Manifest, [string]$Key)
    foreach ($register in @($Manifest.registers)) {
        if (([string]$register.key) -ceq $Key) { return $register }
    }
    throw ('the Stage-B manifest projects no ' + [char]39 + $Key + [char]39 + ' register')
}

function Get-P9RegisterColumnIndex {
    param($Register, [string]$ColumnKey)
    # THE MANIFEST'S OWN COLUMN ORDER, by key. A header-text match would be a
    # second declaration of a name the manifest already owns.
    $ordinal = [array]::IndexOf(@($Register.columns), $ColumnKey) + 1
    if ($ordinal -lt 1) {
        throw ('the register carries no ' + [char]39 + $ColumnKey + [char]39 + ' column')
    }
    return $ordinal
}


# ===========================================================================
# REPORTING
# ===========================================================================
function Write-P9Line {
    param([string]$Text = '')
    $null = $script:P9Lines.Add($Text)
    Write-Host $Text
    # Written through on every line, so a run that is stopped still leaves every
    # observation it had already taken on disk.
    if (-not [string]::IsNullOrWhiteSpace($script:P9Path)) {
        try {
            Set-Content -LiteralPath $script:P9Path `
                -Value ($script:P9Lines -join "`r`n") -Encoding UTF8
        } catch { }
    }
}

function Add-P9Check {
    param([string]$Label, [bool]$Ok, [string]$Detail = '', [string]$Kind = 'RESULT')
    $null = $script:P9Checks.Add([pscustomobject]@{
        Kind = $Kind; Label = $Label; Ok = $Ok; Detail = $Detail
    })
    $verdict = 'FAIL'
    if ($Ok) { $verdict = 'PASS' }
    $line = '  [' + $verdict + '] ' + $Label
    if (-not [string]::IsNullOrWhiteSpace($Detail)) { $line = $line + ' -- ' + $Detail }
    Write-P9Line $line
    return $Ok
}

function Invoke-P9Release {
    param($Ledger, $Obj, [string]$Label)
    $rec = Release-ComObjectSafe -Obj $Obj -Label $Label
    if ($rec.Status -eq 'PASS') {
        $Ledger.Attempted = $Ledger.Attempted + 1
        $Ledger.Succeeded = $Ledger.Succeeded + 1
        $null = $Ledger.Lines.Add(
            ("      {0,-24} | PASS    | ReleaseComObject returned {1}" -f $rec.Label, $rec.Count))
        if ([int]$rec.Count -ne 0) {
            $null = $script:P9Residual.Add(
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

function Get-P9SourceRevision {
    param([string]$RepoRoot)
    # THE SAME REFUSAL THE ACCEPTED RUNNERS MAKE, and the same SHAPE. A result
    # from a tree that is not a commit cannot be attributed to one, so a
    # modified pccm/src, pccm/spec or pccm/builder stops the run before Excel is
    # started.
    #
    # `git -C` RATHER THAN Push-Location, AND NO $LASTEXITCODE. The first draft
    # of this used both, and the dry run refused it on Linux: under StrictMode
    # $LASTEXITCODE has not been set in a fresh session, so the very first
    # refusal this runner makes would have terminated it instead. The accepted
    # runners never read it, and neither does this now.
    $head = ''
    try { $head = [string](& git -C $RepoRoot rev-parse --short HEAD 2>$null) } catch { $head = '' }
    $head = $head.Trim()
    if ([string]::IsNullOrWhiteSpace($head)) {
        throw ('git could not report HEAD for ' + $RepoRoot +
               '; a P9-1 result could not be attributed to a source revision')
    }
    $dirty = @()
    foreach ($pathspec in @('pccm/src', 'pccm/spec', 'pccm/builder')) {
        $lines = @()
        try { $lines = @(& git -C $RepoRoot status --porcelain -- $pathspec 2>$null) }
        catch { $lines = @() }
        foreach ($line in $lines) {
            if (-not [string]::IsNullOrWhiteSpace([string]$line)) { $dirty += [string]$line }
        }
    }
    return [pscustomobject]@{ Head = $head; Dirty = $dirty }
}


# ===========================================================================
# READING THE SHEET
# ===========================================================================
function Get-P9Block {
    param($Workbook, [string]$SheetName, [string]$Address)
    # ONE COM CALL PER BLOCK, not one per cell. A hundred register rows read a
    # cell at a time is ten thousand round trips per scenario; Value2 over a
    # range returns the whole rectangle as one array.
    #
    # THE ARRAY IS RETURNED INSIDE AN OBJECT, AND THAT IS THE WHOLE POINT.
    # Windows run 2 died because this function returned the array directly:
    # `Value2` over a multi-cell range marshals to a RANK-2 object[,], and
    # PowerShell ENUMERATES a multidimensional array on its way out of a
    # function, so the caller received a FLAT object[]. `$block[$row, 1]` on a
    # flat array is not a two-dimensional read - it is PowerShell's multi-index
    # selection, and it quietly returns a TWO-ELEMENT object[]. Every reading
    # looked plausible until the first `[double]`, which is why Scenario A got
    # as far as it did. A property is never unrolled, so the rank survives.
    #
    # NOTHING INDEXES `.Rect` DIRECTLY. Get-P9BlockCell below is the only
    # reader, because the failure above was a SILENT wrong answer before it was
    # a crash.
    $worksheets = $null; $ws = $null; $rng = $null
    try {
        $worksheets = $Workbook.Worksheets
        $ws = $worksheets.Item($SheetName)
        $rng = $ws.Range($Address)
        return [pscustomobject]@{ Sheet = $SheetName; Address = $Address; Rect = $rng.Value2 }
    } finally {
        if ($null -ne $rng)        { Release-Transient $rng        'Range(block)'; $rng        = $null }
        if ($null -ne $ws)         { Release-Transient $ws         'Worksheet';    $ws         = $null }
        if ($null -ne $worksheets) { Release-Transient $worksheets 'Worksheets';   $worksheets = $null }
    }
}

function Get-P9BlockCell {
    param($Block, [int]$Row, [int]$Column)
    # EXACTLY ONE CELL, OR A FAILURE THAT SAYS WHERE. The contract:
    #
    #   a rank-2 block      the cell at ($Row, $Column) in the SHEET's own
    #                       1-based coordinates, read through GetValue so the
    #                       array's real lower bounds are honoured
    #   a single-cell read  Value2 hands back a scalar rather than an array;
    #                       only (1, 1) may ask for it
    #   a rank-1 block      REFUSED. That is the run-2 shape - a rectangle that
    #                       has been flattened - and there is no honest way to
    #                       recover ($Row, $Column) from it
    #   more than one value REFUSED, with the sheet and address named. Taking
    #                       element 0 would be a plausible wrong answer, which
    #                       is exactly what run 2 produced
    #
    # WHAT IT DOES NOT DO. It does not convert, coerce or stringify: a blank
    # comes back as $null, an Excel error comes back as the negative Int32 that
    # Test-P9Error recognises, and text comes back as text. Deciding what a cell
    # MEANS belongs to the caller that knows which cell it asked for.
    $where = 'block ' + [string]$Block.Sheet + '!' + [string]$Block.Address +
             ' at (' + [string]$Row + ',' + [string]$Column + ')'
    $values = $Block.Rect
    if ($values -is [array]) {
        if ($values.Rank -ne 2) {
            throw ($where + ': the block arrived with rank ' + [string]$values.Rank +
                   ' instead of 2, so the rectangle has been flattened and no ' +
                   'row/column read of it can be trusted')
        }
        $cell = $values.GetValue($Row, $Column)
    } else {
        if (($Row -ne 1) -or ($Column -ne 1)) {
            throw ($where + ': the block is a single cell and only (1,1) exists in it')
        }
        $cell = $values
    }
    if ($cell -is [array]) {
        throw ($where + ': the read produced ' + [string]@($cell).Count +
               ' values where exactly one was expected')
    }
    return $cell
}

function Get-P9ColumnLetter {
    param([int]$Index)
    $letters = ''
    $remaining = $Index
    while ($remaining -gt 0) {
        $remainder = ($remaining - 1) % 26
        $letters = [string][char](65 + $remainder) + $letters
        $remaining = [int](($remaining - $remainder - 1) / 26)
    }
    return $letters
}

function ConvertTo-P9ColumnNumber {
    param([string]$Letters)
    $value = 0
    foreach ($character in $Letters.ToUpperInvariant().ToCharArray()) {
        $value = ($value * 26) + ([int][char]$character - 64)
    }
    return $value
}

function Test-P9Error {
    param($Value)
    if ($null -eq $Value) { return $false }
    if ($Value -is [int]) { return $script:P9ErrorCodes.ContainsKey([int]$Value) }
    return $false
}

function Format-P9Cell {
    param($Value)
    if ($null -eq $Value) { return '<blank>' }
    if (Test-P9Error $Value) { return [string]$script:P9ErrorCodes[[int]$Value] }
    if ($Value -is [string]) { return $Value }
    return [string]$Value
}

function Test-P9Na {
    param($Value)
    # #N/A ONLY. Blank is not #N/A and neither is 0: an unused slot that came
    # back as either would be an issue with a fabricated identity, which is the
    # P8-3 hazard this whole representation exists to avoid.
    if (-not (Test-P9Error $Value)) { return $false }
    return ([string]$script:P9ErrorCodes[[int]$Value] -eq '#N/A')
}

function Read-P9Surface {
    param($Workbook, $Projection)
    # EVERY ADDRESS FROM THE PROJECTION. Nothing below spells a row, a column or
    # a vocabulary word of its own.
    $sheet = [string]$Projection.sheet
    $valueColumn = [string]$Projection.summary.value_column

    $summaryRows = @()
    foreach ($property in $Projection.summary.rows.PSObject.Properties) {
        $summaryRows += [pscustomobject]@{ Key = [string]$property.Name
                                           Row = [int]$property.Value.row }
    }
    $firstSummary = ($summaryRows | Measure-Object -Property Row -Minimum).Minimum
    $lastSummary = ($summaryRows | Measure-Object -Property Row -Maximum).Maximum
    $summaryBlock = Get-P9Block -Workbook $Workbook -SheetName $sheet `
        -Address ($valueColumn + [string]$firstSummary + ':' + $valueColumn + [string]$lastSummary)
    $summary = @{}
    foreach ($entry in $summaryRows) {
        $summary[$entry.Key] = Get-P9BlockCell -Block $summaryBlock `
            -Row ($entry.Row - $firstSummary + 1) -Column 1
    }

    $readingRows = @()
    foreach ($property in $Projection.evaluation.readings.rows.PSObject.Properties) {
        $readingRows += [pscustomobject]@{ Key = [string]$property.Name
                                           Row = [int]$property.Value.row }
    }
    $firstReading = ($readingRows | Measure-Object -Property Row -Minimum).Minimum
    $lastReading = ($readingRows | Measure-Object -Property Row -Maximum).Maximum
    $readingBlock = Get-P9Block -Workbook $Workbook -SheetName $sheet `
        -Address ($valueColumn + [string]$firstReading + ':' + $valueColumn + [string]$lastReading)
    $readings = @{}
    foreach ($entry in $readingRows) {
        $readings[$entry.Key] = Get-P9BlockCell -Block $readingBlock `
            -Row ($entry.Row - $firstReading + 1) -Column 1
    }

    $columns = @($Projection.register.columns)
    $numbers = @()
    foreach ($column in $columns) { $numbers += (ConvertTo-P9ColumnNumber ([string]$column.column)) }
    $firstColumn = ($numbers | Measure-Object -Minimum).Minimum
    $lastColumn = ($numbers | Measure-Object -Maximum).Maximum
    $firstRow = [int]$Projection.register.first_row
    $lastRow = [int]$Projection.register.last_row
    $registerBlock = Get-P9Block -Workbook $Workbook -SheetName $sheet `
        -Address ((Get-P9ColumnLetter $firstColumn) + [string]$firstRow + ':' +
                  (Get-P9ColumnLetter $lastColumn) + [string]$lastRow)

    $rows = @()
    for ($offset = 0; $offset -lt ($lastRow - $firstRow + 1); $offset++) {
        $record = @{}
        for ($index = 0; $index -lt $columns.Count; $index++) {
            $key = [string]$columns[$index].column
            $record[[string]$columns[$index].key] = Get-P9BlockCell -Block $registerBlock `
                -Row ($offset + 1) -Column ((ConvertTo-P9ColumnNumber $key) - $firstColumn + 1)
        }
        $rows += [pscustomobject]$record
    }
    return [pscustomobject]@{ Summary = $summary; Readings = $readings; Rows = $rows }
}

function Get-P9Shown {
    param($Surface, $Projection)
    # A ROW IS SHOWN WHILE ITS CHECK ID IS NOT #N/A. The window empties from the
    # bottom, so the first #N/A ends the population - but the check below walks
    # the whole window anyway rather than stopping, because a gap in the middle
    # would be a defect this runner must see rather than skip past.
    $shown = @()
    foreach ($row in @($Surface.Rows)) {
        if (-not (Test-P9Na $row.check_id)) { $shown += $row }
    }
    return $shown
}

function Test-P9WindowShape {
    param($Surface, $Projection, [string]$Stage)
    # THE FILLED ROWS ARE A PREFIX, AND EVERY UNUSED SLOT IS #N/A IN EVERY
    # COLUMN. A blank or a zero anywhere in the window would be an issue with no
    # identity, and a shown row after an unused one would mean the register was
    # not filled in order.
    $columns = @()
    foreach ($column in @($Projection.register.columns)) { $columns += [string]$column.key }
    $seenGap = $false
    $problems = @()
    $index = 0
    foreach ($row in @($Surface.Rows)) {
        $index = $index + 1
        $isUnused = $true
        foreach ($key in $columns) { if (-not (Test-P9Na $row.$key)) { $isUnused = $false } }
        $anyUnused = $false
        foreach ($key in $columns) { if (Test-P9Na $row.$key) { $anyUnused = $true } }
        if ($isUnused) {
            $seenGap = $true
            continue
        }
        if ($anyUnused) {
            $problems += ('row ' + [string]$index + ' is partly #N/A')
            continue
        }
        if ($seenGap) { $problems += ('row ' + [string]$index + ' is filled after an unused slot') }
        foreach ($key in $columns) {
            $value = $row.$key
            if ($null -eq $value) { $problems += ('row ' + [string]$index + ' ' + $key + ' is blank') }
            elseif (($value -isnot [string]) -and (-not (Test-P9Error $value)) -and
                    ([double]$value -eq 0)) {
                $problems += ('row ' + [string]$index + ' ' + $key + ' is a numeric zero')
            }
        }
    }
    return (Add-P9Check ($Stage + ': the window is a filled prefix and every unused slot is #N/A') `
        ($problems.Count -eq 0) ($problems -join '; '))
}

function Test-P9Reconciles {
    param($Surface, $Projection, [string]$Stage)
    # THE COUNTS ARE OF THE ROWS. A reader can add up the register and get the
    # summary, or this fails naming the difference.
    $shown = @(Get-P9Shown -Surface $Surface -Projection $Projection)
    $vocabulary = @($Projection.vocabulary.severity_order)
    $tally = @{}
    foreach ($word in $vocabulary) { $tally[[string]$word] = 0 }
    $unknown = @()
    foreach ($row in $shown) {
        $severity = [string]$row.severity
        if ($tally.ContainsKey($severity)) { $tally[$severity] = $tally[$severity] + 1 }
        else { $unknown += $severity }
    }
    $null = Add-P9Check ($Stage + ': every displayed severity is a declared one') `
        ($unknown.Count -eq 0) (($unknown | Select-Object -Unique) -join ', ')

    $errors = [double]$Surface.Summary['error_count']
    $warnings = [double]$Surface.Summary['warning_count']
    $info = [double]$Surface.Summary['info_count']
    $total = [double]$Surface.Summary['total_checks']
    $windowed = ($total -gt [double]@($Surface.Rows).Count)

    # WHEN THE POPULATION FITS THE WINDOW the counts equal the rows exactly.
    # When it does not, the counts are of the WHOLE population and are therefore
    # larger - which is the claim, not a discrepancy.
    if (-not $windowed) {
        $null = Add-P9Check ($Stage + ': the error count equals the ERROR rows shown') `
            ($errors -eq [double]$tally[[string]$vocabulary[0]]) `
            ([string]$errors + ' counted, ' + [string]$tally[[string]$vocabulary[0]] + ' shown')
        $null = Add-P9Check ($Stage + ': the warning count equals the WARNING rows shown') `
            ($warnings -eq [double]$tally[[string]$vocabulary[1]]) `
            ([string]$warnings + ' counted, ' + [string]$tally[[string]$vocabulary[1]] + ' shown')
        $null = Add-P9Check ($Stage + ': the information count equals the INFO rows shown') `
            ($info -eq [double]$tally[[string]$vocabulary[2]]) `
            ([string]$info + ' counted, ' + [string]$tally[[string]$vocabulary[2]] + ' shown')
        $null = Add-P9Check ($Stage + ': the total equals the rows shown') `
            ($total -eq [double]$shown.Count) `
            ([string]$total + ' counted, ' + [string]$shown.Count + ' shown')
    }
    $null = Add-P9Check ($Stage + ': the total is the sum of the three counts') `
        ($total -eq ($errors + $warnings + $info)) `
        ([string]$total + ' vs ' + [string]($errors + $warnings + $info))

    # PRECEDENCE, READ OFF THE COUNTS. Information never appears in it.
    $states = @($Projection.vocabulary.overall_states)
    $expected = [string]$states[0]
    if ($errors -gt 0) { $expected = [string]$states[2] }
    elseif ($warnings -gt 0) { $expected = [string]$states[1] }
    $null = Add-P9Check ($Stage + ': the Overall Status follows the actionable counts') `
        ([string]$Surface.Summary['overall_status'] -ceq $expected) `
        ('reads ' + (Format-P9Cell $Surface.Summary['overall_status']) + ', expected ' + $expected)

    # AND THE DISCLOSURE IS EXACTLY THE CONTRACTED SENTENCE, OR NOTHING.
    $window = [double]$Projection.register.row_window
    $disclosure = [string](Format-P9Cell $Surface.Summary['disclosure'])
    if ($total -gt $window) {
        $wanted = ([string]$Projection.register.disclosure_template).
            Replace('{window}', [string][int]$window).Replace('{total}', [string][int]$total)
        $null = Add-P9Check ($Stage + ': the overflow is disclosed in the contracted wording') `
            ($disclosure -ceq $wanted) ('reads ' + $disclosure + ', expected ' + $wanted)
    } else {
        $null = Add-P9Check ($Stage + ': nothing is disclosed while the population fits') `
            ([string]::IsNullOrEmpty($disclosure) -or ($disclosure -eq '<blank>')) $disclosure
    }
    return $shown
}

function Test-P9ReadingsAnswered {
    param($Surface, $Projection, [string]$Stage)
    # THE QUESTION LINUX CANNOT ANSWER. A worksheet function that reaches a
    # writer shows #VALUE! from a cell and returns normally from Application.Run,
    # which is exactly how the P8-1 defect hid for a whole phase.
    $broken = @()
    foreach ($property in $Projection.evaluation.readings.rows.PSObject.Properties) {
        $key = [string]$property.Name
        $value = $Surface.Readings[$key]
        if (Test-P9Error $value) {
            $broken += ($key + ' = ' + (Format-P9Cell $value) + ' (' +
                        [string]$property.Value.source + ')')
        }
    }
    return (Add-P9Check ($Stage + ': every reading answered from its cell without an error') `
        ($broken.Count -eq 0) ($broken -join '; '))
}

function Find-P9Row {
    param($Shown, [string]$CheckId)
    foreach ($row in @($Shown)) { if ([string]$row.check_id -ceq $CheckId) { return $row } }
    return $null
}

function Test-P9RowPresent {
    param($Shown, [string]$CheckId, [bool]$Expected, [string]$Stage, [string]$Why = '')
    $found = (Find-P9Row -Shown $Shown -CheckId $CheckId)
    $present = ($null -ne $found)
    $label = $Stage + ': ' + $CheckId + $(if ($Expected) { ' is shown' } else { ' is not shown' })
    if (-not [string]::IsNullOrWhiteSpace($Why)) { $label = $label + ' - ' + $Why }
    return (Add-P9Check $label ($present -eq $Expected) `
        ('present=' + [string]$present))
}

function Get-P9AdvisoryId {
    param($Projection)
    return [string]$Projection.advisory.check_id
}

function Invoke-P9Recalculate {
    param($Excel)
    # AN ORDINARY CALCULATION CYCLE, not CalculateFull. The claim under test is
    # that the sheet is LIVE on a normal recalculation - a full rebuild would
    # re-run everything and prove nothing about the anchor.
    $Excel.Calculate()
}


# ===========================================================================
# THE PROJECTIONS, READ THE ONE WAY THEY ARE MEANT TO BE READ
# ===========================================================================
$manifestPath   = Join-Path $BuildDir 'stage_b_manifest.json'
$inspectPath    = Join-Path $BuildDir 'phase5_gate_b_inspection.json'
$simInspectPath = Join-Path $BuildDir 'phase6_gate_b_inspection.json'
$casesPath      = Join-Path $BuildDir 'phase7_acceptance_cases.json'
$modelCheckPath = Join-Path $BuildDir 'phase9_model_check_inspection.json'

foreach ($required in @($manifestPath, $inspectPath, $simInspectPath, $casesPath,
                        $modelCheckPath)) {
    if (-not (Test-Path -LiteralPath $required)) {
        Write-Host ('REFUSED, BEFORE EXCEL WAS STARTED: ' + $required +
                    ' does not exist. Build Stage A first.') -ForegroundColor Red
        exit 1
    }
}

$manifest   = Get-Content -LiteralPath $manifestPath   -Raw | ConvertFrom-Json
$inspection = Get-Content -LiteralPath $inspectPath    -Raw | ConvertFrom-Json
$simInspect = Get-Content -LiteralPath $simInspectPath -Raw | ConvertFrom-Json
$cases      = Get-Content -LiteralPath $casesPath      -Raw | ConvertFrom-Json
$projection = Get-Content -LiteralPath $modelCheckPath -Raw | ConvertFrom-Json

$modelCheckSheet = [string]$projection.sheet
$recommended = [int]$projection.advisory.threshold
$advisoryId = Get-P9AdvisoryId -Projection $projection

$case = $null
foreach ($scenario in @($cases.scenarios)) { if ([string]$scenario.id -ceq 'W4') { $case = $scenario } }
if ($null -eq $case) { throw 'the acceptance corpus carries no W4 scenario' }
$model = $case.model
$suppliedSeed = [double]$case.supplied_seed

# THE PERSISTED CALCULATION STATUS PAIR, from the Phase-5 projection that owns
# it. This runner never says C19 out loud.
$calcSheet = [string]$inspection.calc.sheet
$calcStateBlock = $inspection.calc.scalar_blocks.calc_state
$calcStatusRow = [int]$calcStateBlock.rows.calculation_status
$calcEvaluatedRow = [int]$calcStateBlock.rows.status_evaluated_at
$calcValueColumn = [string]$calcStateBlock.value_column
$calcStatusRange = ($calcValueColumn + [string]$calcStatusRow + ':' +
                    $calcValueColumn + [string]$calcEvaluatedRow)

# THE TWO REQUEST CONTROLS, from the Phase-6 projection that owns them. This
# runner never spells a defined name.
$iterationsName = [string]$simInspect.controls.monte_carlo_iterations.defined_name
$seedName = [string]$simInspect.controls.random_seed.defined_name
# AND THE PHASE-9 PROJECTION HAS TO AGREE ABOUT WHICH INPUT THE ADVISORY READS.
# Two projections naming two different inputs would mean the advisory advised
# about something the runner never moved.
$advisoryReading = [string]$projection.advisory.reading
$advisorySource = [string]$projection.evaluation.readings.rows.$advisoryReading.source

# THE SOURCE REVISION, BEFORE ANYTHING ELSE HAPPENS.
$revision = $null
try { $revision = Get-P9SourceRevision -RepoRoot $repoRoot }
catch { Write-Host (Format-Err $_) -ForegroundColor Red; exit 1 }
if ($revision.Dirty.Count -gt 0) {
    Write-Host 'REFUSED, BEFORE EXCEL WAS STARTED.' -ForegroundColor Red
    Write-Host ''
    Write-Host ('pccm/src, pccm/spec or pccm/builder is modified, so a P9-1 result ' +
                'could not be attributed to a source revision:') -ForegroundColor Red
    foreach ($line in $revision.Dirty) { Write-Host ('    ' + $line) -ForegroundColor Red }
    exit 1
}


# ===========================================================================
# THE DISPOSABLE WORKING COPY, AND THE STAGE-B BOOTSTRAP
# ===========================================================================
$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) `
    ('pccm-phase9-p1-' + (Get-Date).ToString('yyyyMMdd-HHmmss'))
$null = New-Item -ItemType Directory -Path $tempRoot -Force
Copy-Item -LiteralPath (Join-Path $BuildDir ([string]$manifest.stage_a_filename)) -Destination $tempRoot
foreach ($artefact in @($manifestPath, $inspectPath, $simInspectPath, $casesPath,
                        $modelCheckPath)) {
    Copy-Item -LiteralPath $artefact -Destination $tempRoot
}
Copy-Item -LiteralPath (Join-Path $BuildDir 'vba') -Destination $tempRoot -Recurse

$script:P9Path = Join-Path $tempRoot 'phase9_p1_model_check.txt'
$stageBPath = Join-Path $tempRoot ([string]$manifest.stage_b_filename)

$bootstrap = Join-Path $scriptDir 'build_stage_b.ps1'
& $bootstrap -BuildDir $tempRoot -Force
$bootstrapExit = $LASTEXITCODE
$bootstrapOk = (($bootstrapExit -eq 0) -and (Test-Path -LiteralPath $stageBPath))

Write-P9Line 'PCCM - PHASE 9 P9-1: THE MODEL CHECK SURFACE IN REAL EXCEL'
Write-P9Line '=========================================================='
Write-P9Line ''
Write-P9Line ('source revision   : ' + $revision.Head)
Write-P9Line ('sheet             : ' + $modelCheckSheet)
Write-P9Line ('fixture           : ' + [string]$case.id + ' model, seed ' + [string]$suppliedSeed)
Write-P9Line ('row window        : ' + [string]$projection.register.row_window)
Write-P9Line ('advisory          : ' + $advisoryId + ', ' +
              [string]$projection.advisory.comparison + ' ' + [string]$recommended)
Write-P9Line ('threshold owner   : ' + [string]$projection.provenance.threshold_owner)
Write-P9Line ''
Write-P9Line 'THIS RUNNER ANSWERS PHASE-9 QUESTIONS ONLY. It re-runs no Phase-8'
Write-P9Line 'acceptance and disturbs none of it.'
Write-P9Line ''

$null = Add-P9Check 'the Stage-B workbook was bootstrapped' $bootstrapOk `
    ('build_stage_b.ps1 exit ' + [string]$bootstrapExit) 'PREREQUISITE'
if (-not $bootstrapOk) {
    Write-P9Line ''
    Write-P9Line 'STOP. The Stage-B workbook was not produced; nothing below could run.'
    Write-Host ('The report is at ' + $script:P9Path) -ForegroundColor Cyan
    exit 1
}


# ===========================================================================
# THE SESSION
# ===========================================================================
# EVERY VARIABLE THE CLEANUP READS IS ASSIGNED BEFORE THE TRY IS ENTERED, so the
# finally block is safe however early the try fails - including before Excel
# exists at all.
$preExisting = @(Get-PreExistingExcelPids)
$rel = New-ReleaseLedger
$excel = $null; $workbooks = $null; $wb = $null
$excelIdentity = $null
$naturalExit = $false
$emergencyRequired = $false
$fatal = ''
$comAcquired = 0
try {
    $excel = New-Object -ComObject Excel.Application
    $comAcquired = $comAcquired + 1
    $excelIdentity = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $workbooks = $excel.Workbooks
    $comAcquired = $comAcquired + 1
    $wb = $workbooks.Open($stageBPath)
    $comAcquired = $comAcquired + 1

    $compileFailure = ''
    try { $null = $excel.Run('PCCM_CalculationStatus') } catch { $compileFailure = (Format-Err $_) }
    $null = Add-P9Check 'the current VBAProject compiles in real Excel' `
        ([string]::IsNullOrWhiteSpace($compileFailure)) $compileFailure 'PREREQUISITE'
    if (-not [string]::IsNullOrWhiteSpace($compileFailure)) { throw $compileFailure }

    # THE TWO NEW PROCEDURES EXIST AND ANSWER OUT OF CELL, before anything asks
    # them from one. If they do not, everything below would fail for a reason
    # that had nothing to do with the sheet.
    $adapterFailure = ''
    $adapterAnswer = ''
    try { $adapterAnswer = [string]$excel.Run('PCCM_ModelCheckCalculationState') }
    catch { $adapterFailure = (Format-Err $_) }
    $null = Add-P9Check 'the authorised calculation adapter answers out of cell' `
        ([string]::IsNullOrWhiteSpace($adapterFailure)) `
        ($adapterFailure + $adapterAnswer) 'PREREQUISITE'

    $null = Save-Phase5LockedFxSeed -Workbook $wb -Inspection $inspection

    # -------------------------------------------------------------------
    # THE GEOMETRY IS WHAT THE PROJECTION SAYS IT IS
    # -------------------------------------------------------------------
    Write-P9Line 'GEOMETRY'
    Write-P9Line '--------'
    $sheets = $null; $modelCheckWs = $null; $activeWindow = $null
    try {
        $sheets = $wb.Worksheets
        $modelCheckWs = $sheets.Item($modelCheckSheet)
        $modelCheckWs.Activate()
        $activeWindow = $excel.ActiveWindow
        $freeze = [string]$projection.freeze_panes
        $wantedRow = 0; $wantedColumn = 0
        if ($freeze -match '^([A-Z]+)(\d+)$') {
            $wantedColumn = (ConvertTo-P9ColumnNumber $Matches[1]) - 1
            $wantedRow = [int]$Matches[2] - 1
        }
        $null = Add-P9Check 'the sheet freezes where the projection says' `
            (([int]$activeWindow.SplitRow -eq $wantedRow) -and
             ([int]$activeWindow.SplitColumn -eq $wantedColumn)) `
            ('split ' + [string]$activeWindow.SplitRow + '/' +
             [string]$activeWindow.SplitColumn + ', expected ' + [string]$wantedRow + '/' +
             [string]$wantedColumn)
        # ASKED ONCE AND RELEASED. ChartObjects() hands back a COM collection, and
        # calling it twice to build a message would leak the second one.
        $chartObjects = $modelCheckWs.ChartObjects()
        $chartCount = [int]$chartObjects.Count
        Release-Transient $chartObjects 'ChartObjects'
        $chartObjects = $null
        $null = Add-P9Check 'the sheet draws no chart' ($chartCount -eq 0) `
            ([string]$chartCount + ' chart object(s)')
    } finally {
        if ($null -ne $activeWindow)  { Release-Transient $activeWindow  'Window';     $activeWindow  = $null }
        if ($null -ne $modelCheckWs)  { Release-Transient $modelCheckWs  'Worksheet';  $modelCheckWs  = $null }
        if ($null -ne $sheets)        { Release-Transient $sheets        'Worksheets'; $sheets        = $null }
    }

    $headerRow = [int]$projection.register.header_row
    $headerProblems = @()
    foreach ($column in @($projection.register.columns)) {
        $cell = Get-P9BlockCell -Block (Get-P9Block -Workbook $wb -SheetName $modelCheckSheet `
            -Address ([string]$column.column + [string]$headerRow)) -Row 1 -Column 1
        if ([string]$cell -cne [string]$column.header) {
            $headerProblems += ([string]$column.column + $headerRow + ' reads ' +
                                (Format-P9Cell $cell) + ', expected ' + [string]$column.header)
        }
    }
    $null = Add-P9Check 'every register column carries the header the projection names' `
        ($headerProblems.Count -eq 0) ($headerProblems -join '; ')

    # -------------------------------------------------------------------
    # SCENARIO A - THE UNTOUCHED WORKBOOK
    # -------------------------------------------------------------------
    # NOTHING HAS BEEN CALCULATED AND NOTHING PUBLISHED. What is asserted is the
    # RULE rather than a particular word: whatever the calculation axis says, a
    # simulation INVALID that follows from it is context and is not counted a
    # second time.
    Write-P9Line ''
    Write-P9Line 'SCENARIO A - THE UNTOUCHED WORKBOOK'
    Write-P9Line '-----------------------------------'
    Invoke-P9Recalculate -Excel $excel
    $surfaceA = Read-P9Surface -Workbook $wb -Projection $projection
    $null = Test-P9ReadingsAnswered -Surface $surfaceA -Projection $projection -Stage 'A'
    $null = Test-P9WindowShape -Surface $surfaceA -Projection $projection -Stage 'A'
    $shownA = @(Test-P9Reconciles -Surface $surfaceA -Projection $projection -Stage 'A')
    Write-P9Line ('    calculation state : ' + (Format-P9Cell $surfaceA.Readings['calculation_state']))
    Write-P9Line ('    simulation state  : ' + (Format-P9Cell $surfaceA.Readings['simulation_state']))
    $calcStateA = [string](Format-P9Cell $surfaceA.Readings['calculation_state'])
    $simStateA = [string](Format-P9Cell $surfaceA.Readings['simulation_state'])
    if ($simStateA -ceq 'INVALID') {
        # THE STEP-1 PRECEDENCE, OBSERVED. A simulation INVALID while the
        # calculation is not CURRENT is the calculation's fault and is INFO.
        $simRows = @()
        foreach ($row in $shownA) {
            if ([string]$row.group -ceq 'Simulation') { $simRows += $row }
        }
        $counted = @()
        foreach ($row in $simRows) {
            if ([string]$row.severity -cne [string]$projection.vocabulary.informational_severity) {
                $counted += ([string]$row.check_id + '=' + [string]$row.severity)
            }
        }
        if ($calcStateA -cne 'CURRENT') {
            $null = Add-P9Check 'A: simulation INVALID under a non-current calculation is context' `
                ($counted.Count -eq 0) ($counted -join ', ')
        }
    }
    $null = Add-P9Check 'A: nothing on the untouched workbook is counted as an error twice' `
        ([double]$surfaceA.Summary['error_count'] -le 1) `
        ('errors=' + (Format-P9Cell $surfaceA.Summary['error_count']))

    # -------------------------------------------------------------------
    # SCENARIO B - THE FIXTURE, CALCULATED, NEVER SIMULATED
    # -------------------------------------------------------------------
    Write-P9Line ''
    Write-P9Line 'SCENARIO B - CALCULATED, NEVER SIMULATED'
    Write-P9Line '---------------------------------------'
    $applied = [string](Set-Phase5Fixture -Excel $excel -Workbook $wb -Manifest $manifest `
        -Inspection $inspection -Model $model)
    $null = Add-P9Check 'the accepted W4 fixture applied' ($applied -like 'OK|*') $applied 'PREREQUISITE'
    Set-NamedValue -Workbook $wb -DefinedName $iterationsName -Value ([double]$recommended)
    Set-NamedValue -Workbook $wb -DefinedName $seedName -Value $suppliedSeed
    $null = Invoke-Phase5ProductionOperation -Excel $excel -Operation 'PCCM_Calculate' -Stage 'B'
    Invoke-P9Recalculate -Excel $excel
    $surfaceB = Read-P9Surface -Workbook $wb -Projection $projection
    $null = Test-P9ReadingsAnswered -Surface $surfaceB -Projection $projection -Stage 'B'
    $null = Test-P9WindowShape -Surface $surfaceB -Projection $projection -Stage 'B'
    $shownB = @(Test-P9Reconciles -Surface $surfaceB -Projection $projection -Stage 'B')
    $null = Add-P9Check 'B: the live calculation state reads CURRENT' `
        ([string](Format-P9Cell $surfaceB.Readings['calculation_state']) -ceq 'CURRENT') `
        (Format-P9Cell $surfaceB.Readings['calculation_state'])
    # A VALID MODEL THAT HAS SIMPLY NOT BEEN SIMULATED IS NOT DEFECTIVE.
    $null = Add-P9Check 'B: an unsimulated valid model is PASS' `
        ([string]$surfaceB.Summary['overall_status'] -ceq [string]$projection.vocabulary.overall_states[0]) `
        (Format-P9Cell $surfaceB.Summary['overall_status'])
    $null = Test-P9RowPresent -Shown $shownB -CheckId $advisoryId -Expected $false -Stage 'B' `
        -Why 'the request is at the recommendation'

    # -------------------------------------------------------------------
    # WORKSHEET SAFETY - THE SURFACE CHANGES NO RECORD
    # -------------------------------------------------------------------
    # THE CLAIM THE WHOLE SHEET RESTS ON, and it is observable. _Calc holds the
    # persisted calculation status and the moment it was evaluated. If any
    # cell-called path persisted, an ordinary recalculation would move them.
    Write-P9Line ''
    Write-P9Line 'WORKSHEET SAFETY - A RECALCULATION CHANGES NO RECORD'
    Write-P9Line '---------------------------------------------------'
    $before = Get-P9Block -Workbook $wb -SheetName $calcSheet -Address $calcStatusRange
    Invoke-P9Recalculate -Excel $excel
    Invoke-P9Recalculate -Excel $excel
    $after = Get-P9Block -Workbook $wb -SheetName $calcSheet -Address $calcStatusRange
    $changed = @()
    for ($index = 1; $index -le 2; $index++) {
        $wasValue = (Format-P9Cell (Get-P9BlockCell -Block $before -Row $index -Column 1))
        $nowValue = (Format-P9Cell (Get-P9BlockCell -Block $after -Row $index -Column 1))
        if ($wasValue -cne $nowValue) {
            $changed += ($calcValueColumn + [string]($calcStatusRow + $index - 1) + ': ' +
                         $wasValue + ' -> ' + $nowValue)
        }
    }
    $null = Add-P9Check 'recalculating the workbook rewrote no persisted status cell' `
        ($changed.Count -eq 0) ($changed -join '; ')

    # -------------------------------------------------------------------
    # SCENARIO C AND D - THE PUBLISHED RUN, AND THE ADVISORY BOUNDARY
    # -------------------------------------------------------------------
    Write-P9Line ''
    Write-P9Line 'SCENARIO C - A CURRENT SIMULATION AT THE RECOMMENDATION'
    Write-P9Line '------------------------------------------------------'
    $simResult = [string](Invoke-Phase6Simulation -Excel $excel)
    $null = Add-P9Check 'the simulation published' ($simResult -like 'OK|*') $simResult 'PREREQUISITE'
    Invoke-P9Recalculate -Excel $excel
    $surfaceC = Read-P9Surface -Workbook $wb -Projection $projection
    $null = Test-P9ReadingsAnswered -Surface $surfaceC -Projection $projection -Stage 'C'
    $null = Test-P9WindowShape -Surface $surfaceC -Projection $projection -Stage 'C'
    $shownC = @(Test-P9Reconciles -Surface $surfaceC -Projection $projection -Stage 'C')
    $null = Add-P9Check 'C: the live simulation state reads CURRENT' `
        ([string](Format-P9Cell $surfaceC.Readings['simulation_state']) -ceq 'CURRENT') `
        (Format-P9Cell $surfaceC.Readings['simulation_state'])
    $null = Add-P9Check 'C: a current model with a current run is PASS' `
        ([string]$surfaceC.Summary['overall_status'] -ceq [string]$projection.vocabulary.overall_states[0]) `
        (Format-P9Cell $surfaceC.Summary['overall_status'])
    $null = Test-P9RowPresent -Shown $shownC -CheckId $advisoryId -Expected $false -Stage 'C'

    Write-P9Line ''
    Write-P9Line 'SCENARIO D AND G - THE ADVISORY BOUNDARY, ONE STEP EITHER SIDE'
    Write-P9Line '--------------------------------------------------------------'
    foreach ($offset in @(-1, 0, 1)) {
        $requested = $recommended + $offset
        Set-NamedValue -Workbook $wb -DefinedName $iterationsName -Value ([double]$requested)
        Invoke-P9Recalculate -Excel $excel
        $surface = Read-P9Surface -Workbook $wb -Projection $projection
        $shown = @(Get-P9Shown -Surface $surface -Projection $projection)
        $stage = 'D/G at ' + [string]$requested
        $null = Test-P9WindowShape -Surface $surface -Projection $projection -Stage $stage
        $expected = ($offset -lt 0)
        $null = Test-P9RowPresent -Shown $shown -CheckId $advisoryId -Expected $expected -Stage $stage `
            -Why 'the boundary is strictly less than'
        if ($expected) {
            $row = Find-P9Row -Shown $shown -CheckId $advisoryId
            $null = Add-P9Check ($stage + ': the advisory carries the contracted wording') `
                ((([string]$row.message) -ceq [string]$projection.advisory.message) -and
                 (([string]$row.guidance) -ceq [string]$projection.advisory.guidance)) `
                ((Format-P9Cell $row.message) + ' | ' + (Format-P9Cell $row.guidance))
            $null = Add-P9Check ($stage + ': the advisory is a warning, not an error') `
                ([string]$row.severity -ceq [string]$projection.advisory.severity) `
                (Format-P9Cell $row.severity)
        }
        # AND IT NEVER REFUSES. The hard minimum is untouched, so the request is
        # still a request the model will accept.
        $null = Add-P9Check ($stage + ': the advisory refuses nothing') `
            ([string](Format-P9Cell $surface.Readings['calculation_state']) -ceq 'CURRENT') `
            (Format-P9Cell $surface.Readings['calculation_state'])
    }
    Set-NamedValue -Workbook $wb -DefinedName $iterationsName -Value ([double]$recommended)

    # -------------------------------------------------------------------
    # SCENARIO E AND M - DRIFT, AND THE PERSISTED ROW THAT MAY NOT LIE
    # -------------------------------------------------------------------
    Write-P9Line ''
    Write-P9Line 'SCENARIO E AND M - REQUEST DRIFT, AND LIVE VERSUS PERSISTED'
    Write-P9Line '----------------------------------------------------------'
    $driftName = $seedName
    Set-NamedValue -Workbook $wb -DefinedName $driftName -Value ($suppliedSeed + 1)
    Invoke-P9Recalculate -Excel $excel
    $surfaceE = Read-P9Surface -Workbook $wb -Projection $projection
    $null = Test-P9ReadingsAnswered -Surface $surfaceE -Projection $projection -Stage 'E'
    $null = Test-P9WindowShape -Surface $surfaceE -Projection $projection -Stage 'E'
    $shownE = @(Test-P9Reconciles -Surface $surfaceE -Projection $projection -Stage 'E')
    $liveSim = [string](Format-P9Cell $surfaceE.Readings['simulation_state'])
    $persistedSim = [string](Format-P9Cell $surfaceE.Readings['simulation_status_last_evaluated'])
    Write-P9Line ('    live simulation state    : ' + $liveSim)
    Write-P9Line ('    persisted (last evaluated): ' + $persistedSim)
    $null = Add-P9Check 'E: the live simulation state moved off CURRENT when the request did' `
        ($liveSim -cne 'CURRENT') $liveSim
    # THE PERSISTED ROW IS ALLOWED TO STILL SAY CURRENT. What it may never do is
    # be the row the summary counts, or be presented as the live answer.
    $persistedRows = @()
    foreach ($row in $shownE) {
        if (([string]$row.message) -match '\(last evaluated\)|\(last attempt\)') {
            $persistedRows += $row
        }
    }
    $countedPersisted = @()
    foreach ($row in $persistedRows) {
        if ([string]$row.severity -cne [string]$projection.vocabulary.informational_severity) {
            $countedPersisted += ([string]$row.check_id + '=' + [string]$row.severity)
        }
    }
    $null = Add-P9Check 'M: every persisted row is labelled and none of them is counted' `
        (($persistedRows.Count -gt 0) -and ($countedPersisted.Count -eq 0)) `
        ([string]$persistedRows.Count + ' labelled, ' + ($countedPersisted -join ', '))
    Set-NamedValue -Workbook $wb -DefinedName $driftName -Value $suppliedSeed

    # -------------------------------------------------------------------
    # SCENARIO F - AN INVALID THREE-POINT ORDERING
    # -------------------------------------------------------------------
    # THE ROOT CAUSE IS COUNTED ONCE. The calculation cannot form a package, so
    # the simulation cannot form a request; that is one fault with one remedy,
    # and the simulation row is context.
    Write-P9Line ''
    Write-P9Line 'SCENARIO F - AN INVALID THREE-POINT ORDERING'
    Write-P9Line '-------------------------------------------'
    $costRegister = Get-P9Register -Manifest $manifest -Key 'cost_lines'
    $maxColumnIndex = Get-P9RegisterColumnIndex -Register $costRegister -ColumnKey 'unit_cost_max'
    $descriptionIndex = Get-P9RegisterColumnIndex -Register $costRegister -ColumnKey 'description'
    $breakable = ($maxColumnIndex -ge 1)
    $null = Add-P9Check 'the cost register projects its maximum and description columns' `
        (($maxColumnIndex -ge 1) -and ($descriptionIndex -ge 1)) `
        ('max at ' + [string]$maxColumnIndex + ', description at ' + [string]$descriptionIndex) `
        'PREREQUISITE'
    if ($breakable) {
        $firstLine = @($model.cost_lines)[0]
        $constantId = [string]$firstLine.permanent_id
        $restoreMax = [double]$firstLine.max_value
        # A MAXIMUM BELOW THE MINIMUM. Production refuses the package; nothing
        # here writes a state word.
        Set-P9TableCell -Workbook $wb -SheetName ([string]$costRegister.sheet) `
            -TableName ([string]$costRegister.table_name) -RowIndex 1 `
            -ColumnIndex $maxColumnIndex -Value ([double]$firstLine.min_value - 1)
        Invoke-P9Recalculate -Excel $excel
        $surfaceF = Read-P9Surface -Workbook $wb -Projection $projection
        $null = Test-P9ReadingsAnswered -Surface $surfaceF -Projection $projection -Stage 'F'
        $null = Test-P9WindowShape -Surface $surfaceF -Projection $projection -Stage 'F'
        $shownF = @(Test-P9Reconciles -Surface $surfaceF -Projection $projection -Stage 'F')
        $null = Add-P9Check 'F: the live calculation state reads INVALID' `
            ([string](Format-P9Cell $surfaceF.Readings['calculation_state']) -ceq 'INVALID') `
            (Format-P9Cell $surfaceF.Readings['calculation_state'])
        $null = Add-P9Check 'F: the root cause is counted exactly once' `
            ([double]$surfaceF.Summary['error_count'] -eq 1) `
            ('errors=' + (Format-P9Cell $surfaceF.Summary['error_count']))

        # -----------------------------------------------------------------
        # P9-2A. THE ACTIONABLE ERROR CARRIES THE LIVE REASON
        # -----------------------------------------------------------------
        # THE OWNER'S OWN SENTENCE, AS OF THE MODEL AS IT STANDS. Two things
        # are asked of it, and the second is the one that matters: that it is
        # the LIVE reason and not the persisted last attempt, which is history
        # and - at this point in the run - describes a SUCCESSFUL calculation
        # of a model that no longer exists.
        $liveReason = [string](Format-P9Cell $surfaceF.Readings['calculation_refusal_detail'])
        $persistedReason = [string](Format-P9Cell $surfaceF.Readings['calculation_attempt_detail'])
        Write-P9Line ('    live reason      : ' + $liveReason)
        Write-P9Line ('    persisted detail : ' + $persistedReason)
        $errorRows = @()
        foreach ($row in $shownF) {
            if ([string]$row.severity -ceq [string]@($projection.vocabulary.severity_order)[0]) {
                $errorRows += $row
            }
        }
        $null = Add-P9Check 'F: exactly one actionable ERROR row is displayed' `
            ($errorRows.Count -eq 1) ([string]$errorRows.Count + ' row(s)')
        if ($errorRows.Count -eq 1) {
            $null = Add-P9Check 'F: the actionable ERROR carries the live refusal reason' `
                (([string]$errorRows[0].message -ceq $liveReason) -and
                 (-not [string]::IsNullOrWhiteSpace($liveReason))) `
                (Format-P9Cell $errorRows[0].message)
            # AND THE SUBJECT IS THE ID ITSELF, as a VALUE. P9-2B threads it out
            # of the owner that refused; this asserts EQUALITY with the id the
            # runner itself invalidated, not that it appears somewhere in a
            # sentence. A row naming any other driver fails here.
            $null = Add-P9Check 'F: the actionable ERROR Subject is the invalidated permanent id' `
                ([string]$errorRows[0].subject -ceq $constantId) `
                ('subject=' + (Format-P9Cell $errorRows[0].subject) + ' expected ' + $constantId)
            $null = Add-P9Check 'F: the live subject reading agrees with the displayed Subject' `
                ([string](Format-P9Cell $surfaceF.Readings['calculation_refusal_subject']) -ceq `
                 [string]$errorRows[0].subject) `
                (Format-P9Cell $surfaceF.Readings['calculation_refusal_subject'])
            # AND THE SENTENCE STILL NAMES IT TOO, which is the owner's own
            # message and not something the sheet composed.
            $null = Add-P9Check 'F: the live reason names the same driver' `
                ([string]$errorRows[0].message -clike ('*' + $constantId + '*')) `
                ('looking for ' + $constantId + ' in: ' + (Format-P9Cell $errorRows[0].message))
            $null = Add-P9Check 'F: the actionable ERROR is not the persisted detail' `
                (([string]$errorRows[0].message -cne $persistedReason) -or
                 [string]::IsNullOrWhiteSpace($persistedReason)) `
                ('live=' + $liveReason + ' persisted=' + $persistedReason)
        }
        # AND THE SIMULATION THAT FOLLOWS FROM IT IS CONTEXT, NOT A SECOND ERROR.
        $simCountedF = @()
        foreach ($row in $shownF) {
            if (([string]$row.group -ceq 'Simulation') -and
                ([string]$row.severity -cne [string]$projection.vocabulary.informational_severity)) {
                $simCountedF += ([string]$row.check_id + '=' + [string]$row.severity)
            }
        }
        $null = Add-P9Check 'F: simulation INVALID creates no duplicate actionable ERROR' `
            ($simCountedF.Count -eq 0) ($simCountedF -join ', ')

        # AND THE PERSISTED LAST-ATTEMPT ROWS ARE SEPARATE, LABELLED AND NOT
        # COUNTED. They may say anything; what they may not do is be the live
        # answer or move the summary.
        $attemptRows = @()
        foreach ($row in $shownF) {
            if (([string]$row.message) -match '\(last attempt\)') { $attemptRows += $row }
        }
        $countedAttempts = @()
        foreach ($row in $attemptRows) {
            if ([string]$row.severity -cne [string]$projection.vocabulary.informational_severity) {
                $countedAttempts += ([string]$row.check_id + '=' + [string]$row.severity)
            }
        }
        $null = Add-P9Check 'F: every (last attempt) row is labelled and none is counted' `
            (($attemptRows.Count -gt 0) -and ($countedAttempts.Count -eq 0)) `
            ([string]$attemptRows.Count + ' labelled, ' + ($countedAttempts -join ', '))
        $simCounted = @()
        foreach ($row in $shownF) {
            if (([string]$row.group -ceq 'Simulation') -and
                ([string]$row.severity -cne [string]$projection.vocabulary.informational_severity)) {
                $simCounted += ([string]$row.check_id + '=' + [string]$row.severity)
            }
        }
        $null = Add-P9Check 'F: the simulation that follows from it is context, not a second error' `
            ($simCounted.Count -eq 0) ($simCounted -join ', ')
        Set-P9TableCell -Workbook $wb -SheetName ([string]$costRegister.sheet) `
            -TableName ([string]$costRegister.table_name) -RowIndex 1 `
            -ColumnIndex $maxColumnIndex -Value $restoreMax
        Invoke-P9Recalculate -Excel $excel
    }

    # -------------------------------------------------------------------
    # SCENARIO F2 - A REFUSAL FROM AN OWNER THE ORDERING CHECK NEVER REACHES
    # -------------------------------------------------------------------
    # WHY A SECOND INVALID SCENARIO. Scenario F breaks the three-point ordering,
    # which modCalcCheck refuses BEFORE any factor, audit or total is built, so
    # it proves the Subject for exactly one owner. P9-3 threaded the owners
    # DOWNSTREAM of that check, and a rule proved on one owner is not proved on
    # the others.
    #
    # SO THIS ONE PASSES EVERY VALIDATION AND FAILS IN THE ARITHMETIC. Min, Most
    # Likely and Max are set equal, so the ordering holds; the quantity stays
    # strictly positive; the profile weights still sum to one. What cannot be
    # represented is the PRODUCT, and the driver audit is where that is found -
    # a P9-3 family in modCalcReport. The Subject must still be the driver this
    # runner broke, and the sentence must not be the ordering one again.
    Write-P9Line ''
    Write-P9Line 'SCENARIO F2 - A REFUSAL RAISED AFTER VALIDATION PASSES'
    Write-P9Line '-----------------------------------------------------'
    $minColumnIndex = Get-P9RegisterColumnIndex -Register $costRegister -ColumnKey 'unit_cost_min'
    $likelyColumnIndex = Get-P9RegisterColumnIndex -Register $costRegister -ColumnKey 'unit_cost_most_likely'
    $quantityColumnIndex = Get-P9RegisterColumnIndex -Register $costRegister -ColumnKey 'quantity'
    $reachable = (($minColumnIndex -ge 1) -and ($likelyColumnIndex -ge 1) -and
                  ($maxColumnIndex -ge 1) -and ($quantityColumnIndex -ge 1))
    $null = Add-P9Check 'the cost register projects the columns scenario F2 needs' `
        $reachable `
        ('min=' + [string]$minColumnIndex + ' likely=' + [string]$likelyColumnIndex +
         ' max=' + [string]$maxColumnIndex + ' quantity=' + [string]$quantityColumnIndex) `
        'PREREQUISITE'
    if ($reachable) {
        $brokenLine = @($model.cost_lines)[0]
        $constantId2 = [string]$brokenLine.permanent_id
        $restoreMin = [double]$brokenLine.min_value
        $restoreLikely = [double]$brokenLine.most_likely
        $restoreMax2 = [double]$brokenLine.max_value
        $restoreQuantity = [double]$brokenLine.quantity
        $huge = [double]1e+300
        foreach ($columnIndex in @($minColumnIndex, $likelyColumnIndex, $maxColumnIndex,
                                   $quantityColumnIndex)) {
            Set-P9TableCell -Workbook $wb -SheetName ([string]$costRegister.sheet) `
                -TableName ([string]$costRegister.table_name) -RowIndex 1 `
                -ColumnIndex $columnIndex -Value $huge
        }
        Invoke-P9Recalculate -Excel $excel
        $surfaceF2 = Read-P9Surface -Workbook $wb -Projection $projection
        $null = Test-P9ReadingsAnswered -Surface $surfaceF2 -Projection $projection -Stage 'F-ARITHMETIC'
        $shownF2 = @(Test-P9Reconciles -Surface $surfaceF2 -Projection $projection -Stage 'F-ARITHMETIC')
        $liveReason2 = [string](Format-P9Cell $surfaceF2.Readings['calculation_refusal_detail'])
        $liveSubject2 = [string](Format-P9Cell $surfaceF2.Readings['calculation_refusal_subject'])
        Write-P9Line ('    live reason      : ' + $liveReason2)
        Write-P9Line ('    live subject     : ' + $liveSubject2)
        $null = Add-P9Check 'F2: the live calculation state reads INVALID' `
            ([string](Format-P9Cell $surfaceF2.Readings['calculation_state']) -ceq 'INVALID') `
            (Format-P9Cell $surfaceF2.Readings['calculation_state'])
        $null = Add-P9Check 'F2: the root cause is counted exactly once' `
            ([double]$surfaceF2.Summary['error_count'] -eq 1) `
            ('errors=' + (Format-P9Cell $surfaceF2.Summary['error_count']))
        $errorRows2 = @()
        foreach ($row in $shownF2) {
            if ([string]$row.severity -ceq [string]@($projection.vocabulary.severity_order)[0]) {
                $errorRows2 += $row
            }
        }
        $null = Add-P9Check 'F2: exactly one actionable ERROR row is displayed' `
            ($errorRows2.Count -eq 1) ([string]$errorRows2.Count + ' row(s)')
        if ($errorRows2.Count -eq 1) {
            $null = Add-P9Check 'F2: the actionable ERROR Subject is the invalidated permanent id' `
                ([string]$errorRows2[0].subject -ceq $constantId2) `
                ('subject=' + (Format-P9Cell $errorRows2[0].subject) + ' expected ' + $constantId2)
            $null = Add-P9Check 'F2: the live subject reading agrees with the displayed Subject' `
                ($liveSubject2 -ceq [string]$errorRows2[0].subject) $liveSubject2
            $null = Add-P9Check 'F2: the refusal comes from an owner past the ordering check' `
                (-not ([string]$errorRows2[0].message -clike '*Min <= Most Likely <= Max*')) `
                (Format-P9Cell $errorRows2[0].message)
            $null = Add-P9Check 'F2: the live reason names the same driver' `
                ([string]$errorRows2[0].message -clike ('*' + $constantId2 + '*')) `
                ('looking for ' + $constantId2 + ' in: ' + (Format-P9Cell $errorRows2[0].message))
        }
        $restoreBy = @{}
        $restoreBy[$minColumnIndex] = $restoreMin
        $restoreBy[$likelyColumnIndex] = $restoreLikely
        $restoreBy[$maxColumnIndex] = $restoreMax2
        $restoreBy[$quantityColumnIndex] = $restoreQuantity
        foreach ($columnIndex in @($restoreBy.Keys)) {
            Set-P9TableCell -Workbook $wb -SheetName ([string]$costRegister.sheet) `
                -TableName ([string]$costRegister.table_name) -RowIndex 1 `
                -ColumnIndex $columnIndex -Value ([double]$restoreBy[$columnIndex])
        }
        Invoke-P9Recalculate -Excel $excel
    }

    # -------------------------------------------------------------------
    # THE ANCHOR - A STRUCTURAL FAULT RAISED AFTER THE WORKBOOK OPENED
    # -------------------------------------------------------------------
    # THE QUESTION LINUX CANNOT ANSWER. PCCM_StructuralReport is not volatile,
    # so its cell is made a dependent of the volatile evaluation heartbeat
    # instead. If that does not work the report below will not change, and the
    # sheet would be reporting a workbook that no longer exists.
    #
    # THE FAULT IS PRODUCTION'S OWN: an unkeyed register row holding data is
    # CHK_NO_ORPHAN_STRUCTURAL_DATA, and nothing here writes a fault, a severity
    # or a Model Check cell.
    Write-P9Line ''
    Write-P9Line 'THE ANCHOR - A FAULT RAISED AFTER THE WORKBOOK OPENED'
    Write-P9Line '----------------------------------------------------'
    $surfaceBefore = Read-P9Surface -Workbook $wb -Projection $projection
    $beforeFaults = [double]$surfaceBefore.Readings['structural_fault_count']
    $canOrphan = ($descriptionIndex -ge 1)
    if ($canOrphan) {
        $addedRow = (Add-P9BlankTableRow -Workbook $wb -SheetName ([string]$costRegister.sheet) `
            -TableName ([string]$costRegister.table_name))
        Set-P9TableCell -Workbook $wb -SheetName ([string]$costRegister.sheet) `
            -TableName ([string]$costRegister.table_name) -RowIndex $addedRow `
            -ColumnIndex $descriptionIndex -Value 'an unkeyed row holding data'
        Invoke-P9Recalculate -Excel $excel
        $surfaceS = Read-P9Surface -Workbook $wb -Projection $projection
        $afterFaults = [double]$surfaceS.Readings['structural_fault_count']
        $null = Add-P9Check 'the anchored structural report re-evaluated on an ordinary recalculation' `
            ($afterFaults -gt $beforeFaults) `
            ([string]$beforeFaults + ' -> ' + [string]$afterFaults + ' fault(s)')
        $null = Test-P9WindowShape -Surface $surfaceS -Projection $projection -Stage 'STRUCT'
        $shownS = @(Test-P9Reconciles -Surface $surfaceS -Projection $projection -Stage 'STRUCT')
        $structuralRows = @()
        foreach ($row in $shownS) {
            if ([string]$row.group -ceq [string]@($projection.vocabulary.group_order)[0]) {
                $structuralRows += $row
            }
        }
        $null = Add-P9Check 'the fault is shown one row per fault, with the owner as its identity' `
            (($structuralRows.Count -ge 1) -and
             (-not [string]::IsNullOrWhiteSpace([string]$structuralRows[0].check_id))) `
            ([string]$structuralRows.Count + ' row(s), first id ' +
             (Format-P9Cell $structuralRows[0].check_id))
        # AND STRUCTURAL FAULTS SORT AHEAD OF EVERYTHING, which is what makes the
        # candidate order the displayed order.
        $null = Add-P9Check 'the structural rows are the first rows in the register' `
            (($shownS.Count -ge $structuralRows.Count) -and
             ([string]$shownS[0].group -ceq [string]@($projection.vocabulary.group_order)[0])) `
            ('first group ' + (Format-P9Cell $shownS[0].group))
        Remove-P9TableRow -Workbook $wb -SheetName ([string]$costRegister.sheet) `
            -TableName ([string]$costRegister.table_name) -RowIndex $addedRow
        Invoke-P9Recalculate -Excel $excel
        $surfaceR = Read-P9Surface -Workbook $wb -Projection $projection
        $null = Add-P9Check 'the fault cleared when the row was removed' `
            ([double]$surfaceR.Readings['structural_fault_count'] -eq $beforeFaults) `
            ([string]$surfaceR.Readings['structural_fault_count'])
    }

    # -------------------------------------------------------------------
    # AND NOTHING WAS PERSISTED BY ANY OF IT
    # -------------------------------------------------------------------
    $finalStatus = Get-P9Block -Workbook $wb -SheetName $calcSheet -Address $calcStatusRange
    Invoke-P9Recalculate -Excel $excel
    $stillStatus = Get-P9Block -Workbook $wb -SheetName $calcSheet -Address $calcStatusRange
    $stillChanged = @()
    for ($index = 1; $index -le 2; $index++) {
        if ((Format-P9Cell (Get-P9BlockCell -Block $finalStatus -Row $index -Column 1)) -cne
            (Format-P9Cell (Get-P9BlockCell -Block $stillStatus -Row $index -Column 1))) {
            $stillChanged += ($calcValueColumn + [string]($calcStatusRow + $index - 1))
        }
    }
    $null = Add-P9Check 'after every scenario, a recalculation still rewrites no persisted cell' `
        ($stillChanged.Count -eq 0) ($stillChanged -join ', ')

    # THE WORKBOOK IS NEVER SAVED. A run that wrote its fixture back would make
    # the next run start somewhere else.
    # EVERY LIFECYCLE FACT IS RECORDED WHERE IT HAPPENS. Run 1 reported
    # Workbook.Close, Application.Quit and natural PID exit as False while the
    # verdict below passed the natural-exit check, because this runner performed
    # the actions and - alone among the accepted runners - never wrote them to
    # the ledger it then printed. The fields were not pre-cleanup state and not
    # final state; they were their initialised False.
    try { $wb.Close($false); $rel.WorkbookClosed = $true }
    catch { $null = $rel.Failed.Add('Workbook.Close'); throw }
    Invoke-P9Release -Ledger $rel -Obj $wb -Label 'Workbook'
    $wb = $null
    Invoke-P9Release -Ledger $rel -Obj $workbooks -Label 'Workbooks'
    $workbooks = $null
    try { $excel.Quit(); $rel.QuitCalled = $true }
    catch { $null = $rel.Failed.Add('Application.Quit'); throw }
    Invoke-P9Release -Ledger $rel -Obj $excel -Label 'Application'
    $excel = $null
    $Error.Clear()
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    $naturalExit = (Wait-ExcelExit -Identity $excelIdentity -TimeoutSeconds 90)
} catch {
    $fatal = (Format-Err $_)
    Write-P9Line ''
    Write-P9Line ('FATAL: ' + $fatal)
} finally {
    # THE FATAL PATH RECORDS WHAT IT DID, on the same terms as the path above.
    # A Close or a Quit that throws here is not swallowed: it joins the failed
    # labels, so `every COM release succeeded` reports it instead of the run
    # looking clean because nobody wrote the failure down.
    if ($null -ne $wb) {
        try { $wb.Close($false); $rel.WorkbookClosed = $true }
        catch { $null = $rel.Failed.Add('Workbook.Close') }
        Invoke-P9Release -Ledger $rel -Obj $wb -Label 'Workbook'
        $wb = $null
    }
    if ($null -ne $workbooks) {
        Invoke-P9Release -Ledger $rel -Obj $workbooks -Label 'Workbooks'
        $workbooks = $null
    }
    if ($null -ne $excel) {
        try { $excel.Quit(); $rel.QuitCalled = $true }
        catch { $null = $rel.Failed.Add('Application.Quit') }
        Invoke-P9Release -Ledger $rel -Obj $excel -Label 'Application'
        $excel = $null
        $Error.Clear()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        $naturalExit = (Wait-ExcelExit -Identity $excelIdentity -TimeoutSeconds 90)
    }
    if (($null -ne $excelIdentity) -and (-not $naturalExit)) {
        $emergencyRequired = $true
        $null = Invoke-EmergencyExcelCleanup -Identity $excelIdentity
    }
    # ONE FACT, WRITTEN ONCE, READ BY BOTH THE LEDGER AND THE VERDICT. The two
    # cannot disagree about this run because there is no second copy to drift.
    $rel.NaturalExit = $naturalExit
    $rel.EmergencyRequired = $emergencyRequired
    Write-P9Line ''
    Write-P9Line 'COM LIFECYCLE'
    Write-P9Line '-------------'
    foreach ($line in @(Format-ReleaseLedger -Ledger $rel)) { Write-P9Line ([string]$line) }
    foreach ($transient in @(Get-TransientFailures)) {
        Write-P9Line ('      transient release FAILED: ' + [string]$transient)
    }
}


# ===========================================================================
# THE LIFECYCLE VERDICT, on the same terms the accepted runners use
# ===========================================================================
# THIS RUN'S OWN LIFECYCLE FACTS. A result from a session that leaked Excel is
# not evidence about a worksheet.
$null = Add-P9Check 'the owned Excel process exited naturally' $rel.NaturalExit
$null = Add-P9Check 'no emergency cleanup was required' (-not $rel.EmergencyRequired)
$null = Add-P9Check 'every COM object this runner acquired was released' `
    ([int]$rel.Attempted -eq [int]$comAcquired) `
    ([string]$comAcquired + ' acquired, ' + [string]$rel.Attempted + ' released')
$null = Add-P9Check 'every COM release succeeded' ($rel.Failed.Count -eq 0) ($rel.Failed -join ', ')
$null = Add-P9Check 'every COM release left 0 outstanding references' `
    (@($script:P9Residual).Count -eq 0) ((@($script:P9Residual)) -join '; ')

$results = @($script:P9Checks | Where-Object { [string]$_.Kind -eq 'RESULT' })
$prereqs = @($script:P9Checks | Where-Object { [string]$_.Kind -eq 'PREREQUISITE' })
$failedResults = @($results | Where-Object { -not $_.Ok })
$failedPrereqs = @($prereqs | Where-Object { -not $_.Ok })
Write-P9Line ''
Write-P9Line 'VERDICT'
Write-P9Line '-------'
Write-P9Line ('prerequisites     : ' + [string]$prereqs.Count + ' checked, ' +
              [string]$failedPrereqs.Count + ' failed')
Write-P9Line ('scenario results  : ' + [string]$results.Count + ' checked, ' +
              [string]$failedResults.Count + ' failed')
foreach ($failure in @($failedPrereqs + $failedResults)) {
    Write-P9Line ('    FAILED: ' + [string]$failure.Label + ' -- ' + [string]$failure.Detail)
}
$ok = (($failedResults.Count -eq 0) -and ($failedPrereqs.Count -eq 0) -and
       [string]::IsNullOrWhiteSpace($fatal))
Write-P9Line ('P9-1 ' + $(if ($ok) { 'PASS' } else { 'FAIL' }))
Write-P9Line ''
Write-P9Line 'THIS RUNNER ANSWERS PHASE-9 QUESTIONS ONLY. No result here re-establishes'
Write-P9Line 'or disturbs any accepted Phase-8 evidence.'
Write-P9Line ''
Write-P9Line ('report            : ' + $script:P9Path)
Write-Host ('The report is at ' + $script:P9Path) -ForegroundColor Cyan
Write-Host 'The working copy is left in place so the report survives; delete it when done.'
if ($ok) { exit 0 } else { exit 1 }
