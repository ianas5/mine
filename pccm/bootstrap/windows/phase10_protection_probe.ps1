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

function Set-ProbeNamedValue {
    param($Workbook, [string]$DefinedName, $Value)
    $names = $null; $nm = $null; $range = $null
    try {
        $names = $Workbook.Names
        $nm = $names.Item($DefinedName)
        $range = $nm.RefersToRange
        $range.Value2 = $Value
    } finally {
        if ($null -ne $range) { Release-Transient $range 'Range'; $range = $null }
        if ($null -ne $nm)    { Release-Transient $nm    'Name';  $nm    = $null }
        if ($null -ne $names) { Release-Transient $names 'Names'; $names = $null }
    }
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

# THE SHAPE OF ONE TABLE, READ AND NEVER CHANGED. Rows and columns only, through
# the ListObject's own counts - this is a read, not a structural operation.
function Get-ProbeTableShape {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $sheets = $null; $ws = $null; $los = $null; $lo = $null
    try {
        $sheets = $Workbook.Worksheets
        $ws = $sheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $columns = 0; $rows = 0
        $columnCollection = $null; $rowCollection = $null
        $where = $SheetName + '!' + $TableName
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
        return [pscustomobject]@{ Table = $TableName; Columns = $columns; Rows = $rows }
    } finally {
        if ($null -ne $lo)     { Release-Transient $lo     'ListObject';  $lo     = $null }
        if ($null -ne $los)    { Release-Transient $los    'ListObjects'; $los    = $null }
        if ($null -ne $ws)     { Release-Transient $ws     'Worksheet';   $ws     = $null }
        if ($null -ne $sheets) { Release-Transient $sheets 'Worksheets';  $sheets = $null }
    }
}

# EVERY TABLE THE STRUCTURAL COMMANDS RESHAPE, from the manifest rather than
# from a list typed here.
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
    return ,@($watched)
}

function Get-ProbeAllShapes {
    param($Workbook, $Watched)
    $shapes = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($entry in @($Watched)) {
        $shape = Get-ProbeTableShape -Workbook $Workbook -SheetName $entry.Sheet -TableName $entry.Table
        $shapes.Add([string]$entry.Key, $shape)
    }
    return $shapes
}

# WHAT ACTUALLY CHANGED. An announcement of success with no shape change is not
# a structural operation having happened.
function Get-ProbeShapeDelta {
    param($Before, $After, $Watched)
    $changes = @()
    foreach ($entry in @($Watched)) {
        $key = [string]$entry.Key
        $b = $Before[$key]
        $a = $After[$key]
        if ($null -eq $b -or $null -eq $a) { continue }
        if (([int]$b.Columns -ne [int]$a.Columns) -or ([int]$b.Rows -ne [int]$a.Rows)) {
            $changes += ($key + ' ' + [string]$b.Rows + 'x' + [string]$b.Columns +
                         ' -> ' + [string]$a.Rows + 'x' + [string]$a.Columns)
        }
    }
    return ,@($changes)
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
    param($Excel, $Workbook, [string]$Endpoint, $Watched)

    Set-ProbeStage -Stage 'endpoint' -Action 'invoking the production entry point' -Endpoint $Endpoint

    $protectionBefore = Get-ProbeProtectionState -Workbook $Workbook -Where ('before ' + $Endpoint)
    $shapesBefore = Get-ProbeAllShapes -Workbook $Workbook -Watched $Watched

    $raised = ''
    $result = ''
    try {
        $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
        $Excel.Run($Endpoint) | Out-Null
        $result = [string]$Excel.Run('PCCM_AutomationResult')
    } catch {
        $raised = (Format-Err $_)
    }

    Set-ProbeStage -Stage 'endpoint' -Action 'reading the structural effect back' -Endpoint $Endpoint
    $protectionAfter = Get-ProbeProtectionState -Workbook $Workbook -Where ('after ' + $Endpoint)
    $shapesAfter = Get-ProbeAllShapes -Workbook $Workbook -Watched $Watched
    $changes = @(Get-ProbeShapeDelta -Before $shapesBefore -After $shapesAfter -Watched $Watched)

    $announced = ([bool]($raised -eq ''))
    $succeeded = ([bool]($announced -and ($result -like 'OK|*')))
    $outcome = 'REFUSED'
    if (-not $announced)  { $outcome = 'RAISED' }
    elseif ($succeeded)   { $outcome = 'SUCCEEDED' }

    return [pscustomobject]@{
        Endpoint          = $Endpoint
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
    }
}

function Write-ProbeOutcome {
    param($Outcome, [string]$Expectation)
    Write-ProbeLine ('  ' + [string]$Outcome.Endpoint)
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
Write-ProbeLine 'TABLES WATCHED FOR A STRUCTURAL EFFECT'
foreach ($entry in @($watched)) {
    Write-ProbeLine ('  ' + ([string]$entry.Key).PadRight(16) + [string]$entry.Sheet + '!' + [string]$entry.Table)
}
Write-ProbeLine ''
# ===========================================================================
# THE SESSION
# ===========================================================================
$preExisting = @(Get-PreExistingExcelPids)
$excel = $null; $workbooks = $null; $wb = $null; $excelIdentity = $null; $rel = $null

# THE VERDICT STARTS INCONCLUSIVE AND IS ONLY EVER MOVED BY THE ANSWER ITSELF.
# Every failure path leaves it where it is, so a probe that broke cannot say
# anything about production.
$verdict = 'INCONCLUSIVE'
$verdictReason = 'the probe did not reach its conclusion'
$outcomes = New-Object System.Collections.ArrayList
$controlWorked = $false
$controlDetail = 'not attempted'

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
        $discountName = [string]$inspection.inputs.discount_rate.defined_name
        $original = Get-ProbeNamedValue -Workbook $wb -DefinedName $discountName
        try {
            Set-ProbeNamedValue -Workbook $wb -DefinedName $discountName -Value ([double]0.05)
            $controlWorked = $true
            $controlDetail = 'a cell VALUE write on a protected sheet SUCCEEDED'
            Write-ProbeLine ('  ' + $controlDetail)
            Write-ProbeLine '  so UserInterfaceOnly is honoured for code that writes VALUES.'
            Set-ProbeNamedValue -Workbook $wb -DefinedName $discountName -Value $original
        } catch {
            $controlDetail = ('a cell VALUE write on a protected sheet was REFUSED: ' + (Format-Err $_))
            Write-ProbeLine ('  ' + $controlDetail)
            Write-ProbeLine '  so protection is blocking code writes as well, which is a DIFFERENT'
            Write-ProbeLine '  and larger finding than the table-structure one.'
        }
        Write-ProbeLine ''

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
        Set-ProbeStage -Stage 'endpoint' -Action 'setting the timeline inputs' -Endpoint 'PCCM_ApplyTimeline'
        Set-ProbeNamedValue -Workbook $wb -DefinedName ([string]$inspection.inputs.base_year.defined_name) -Value ([double]2026)
        Set-ProbeNamedValue -Workbook $wb -DefinedName ([string]$inspection.inputs.project_start_year.defined_name) -Value ([double]2027)
        Set-ProbeNamedValue -Workbook $wb -DefinedName ([string]$inspection.inputs.duration_years.defined_name) -Value ([double]3)

        $timeline = Invoke-ProbeEndpoint -Excel $excel -Workbook $wb `
            -Endpoint 'PCCM_ApplyTimeline' -Watched $watched
        $null = $outcomes.Add($timeline)
        Write-ProbeOutcome -Outcome $timeline -Expectation `
            ('ListColumns.Add on three grids. A fresh workbook has no year columns, so ' +
             'this fires for ANY timeline and is not capacity-dependent.')

        $addCost = Invoke-ProbeEndpoint -Excel $excel -Workbook $wb `
            -Endpoint 'PCCM_AddCostLine' -Watched $watched
        $null = $outcomes.Add($addCost)
        Write-ProbeOutcome -Outcome $addCost -Expectation `
            ('the register grows only past its reserved rows, but SyncRows runs on every ' +
             'add and reshapes the Cost Profiling grid.')

        $addRisk = Invoke-ProbeEndpoint -Excel $excel -Workbook $wb `
            -Endpoint 'PCCM_AddRisk' -Watched $watched
        $null = $outcomes.Add($addRisk)
        Write-ProbeOutcome -Outcome $addRisk -Expectation `
            'the same path on the Risk Register and the Risk Profiling grid.'

        $calculate = Invoke-ProbeEndpoint -Excel $excel -Workbook $wb `
            -Endpoint 'PCCM_Calculate' -Watched $watched
        $null = $outcomes.Add($calculate)
        Write-ProbeOutcome -Outcome $calculate -Expectation `
            ('the _Calc tables are resized to the model on every Calculate. They are not ' +
             'watched above, so judge this one by its announcement.')

        # --- THE VERDICT ---------------------------------------------------
        # BLOCKED and FINE are the only two conclusions, and each needs its own
        # evidence. Anything else stays INCONCLUSIVE, which is where it started.
        Set-ProbeStage -Stage 'verdict' -Action 'weighing the outcomes'
        $notSucceeded = @(@($outcomes) | Where-Object { [string]$_.Outcome -ne 'SUCCEEDED' })
        $lostProtection = @(@($outcomes) | Where-Object {
            [int]$_.ProtectedAfter -ne [int]$_.TotalSheets })
        $structural = @(@($outcomes) | Where-Object { [bool]$_.StructuralEffect })

        if (@($notSucceeded).Count -gt 0) {
            $verdict = 'PRODUCTION IS BLOCKED BY PROTECTION'
            $verdictReason = ([string]@($notSucceeded).Count + ' of ' + [string]@($outcomes).Count +
                              ' production endpoints did not succeed while every sheet was protected')
        } elseif (@($structural).Count -lt 1) {
            $verdictReason = ('every endpoint announced success but no watched table ever ' +
                              'changed shape, so the structural effect was not observed and ' +
                              'the question is not settled')
        } elseif (@($lostProtection).Count -gt 0) {
            $verdictReason = ('the commands succeeded but protection was not in force after ' +
                              [string]@($lostProtection).Count + ' of them, so they were not ' +
                              'a test of protected behaviour')
        } else {
            $verdict = 'PRODUCTION IS FINE UNDER PROTECTION'
            $verdictReason = ('every production endpoint succeeded, the tables actually ' +
                              'changed shape, and every sheet was protected before and ' +
                              'after each one - so Benchmark Run 3 was a HARNESS defect only')
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
Write-ProbeLine ('  ' + $controlDetail)
Write-ProbeLine ('  UserInterfaceOnly permits code VALUE writes: ' + [string]$controlWorked)
Write-ProbeLine '  That is a SEPARATE capability from permission to perform a ListObject'
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
