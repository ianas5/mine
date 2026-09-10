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
# THE PHASE-7 ADDRESSES
# ===========================================================================
# `phase6_gate_b_inspection.json` is a PINNED artefact whose SHA-256 the
# Phase-6 Gate-B fixture-integrity checks depend on. Extending it to carry the


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
        } catch { }
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

function Get-BenchmarkOneDriveRoots {
    $roots = @()
    foreach ($name in @('OneDrive', 'OneDriveCommercial', 'OneDriveConsumer')) {
        $value = [string](Get-Item -LiteralPath ('Env:' + $name) -ErrorAction SilentlyContinue).Value
        if (-not [string]::IsNullOrWhiteSpace($value)) { $roots += $value }
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
        $drive = Get-CimInstance -ClassName Win32_LogicalDisk `
            -Filter ("DeviceID='" + $qualifier.TrimEnd('\') + "'") -ErrorAction SilentlyContinue
        if ($null -ne $drive -and [int]$drive.DriveType -eq 4) { return 'synced' }
    } catch { }
    return 'local'
}

function Get-BenchmarkEnvironment {
    param($Excel, [string]$WorkbookPath, [string]$RepositoryPath,
          $Manifest, [string]$HarnessVersion, [int]$SchemaVersion, $Revision)
    $unknown = Get-BenchmarkUnavailable
    $roots = @(Get-BenchmarkOneDriveRoots)

    $os = $null; $cpu = $null; $computer = $null
    try { $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue } catch { }
    try { $cpu = @(Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue)[0] } catch { }
    try { $computer = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue } catch { }

    $ram = $unknown
    if ($null -ne $computer -and $null -ne $computer.TotalPhysicalMemory) {
        $ram = [string]('{0:N1}' -f ([double]$computer.TotalPhysicalMemory / 1GB))
    }

    $excelVersion = $unknown; $excelBuild = $unknown; $excelBitness = $unknown
    $calcMode = $unknown; $otherBooks = $unknown
    if ($null -ne $Excel) {
        try { $excelVersion = [string]$Excel.Version } catch { }
        try { $excelBuild = [string]$Excel.Build } catch { }
        try { $excelBitness = [string]$Excel.OperatingSystem } catch { }
        try {
            # Excel reports its own bitness through the process, not through the
            # Application object; both are recorded because neither alone is
            # enough to say which build is installed.
            $excelBitness = [string]([System.Environment]::Is64BitProcess)
            $excelBitness = $(if ([System.Environment]::Is64BitProcess) { '64-bit host process' }
                              else { '32-bit host process' })
        } catch { }
        try { $calcMode = [string]$Excel.Calculation } catch { }
        try { $otherBooks = [string]($Excel.Workbooks.Count) } catch { }
    }

    $branch = $unknown
    try {
        $named = [string](& git -C $repoRoot rev-parse --abbrev-ref HEAD 2>$null)
        if (-not [string]::IsNullOrWhiteSpace($named)) { $branch = $named.Trim() }
    } catch { }

    $record = New-Object System.Collections.Specialized.OrderedDictionary
    $record.Add('captured_at_utc', ((Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')))
    $record.Add('host_name', [string]$env:COMPUTERNAME)
    $record.Add('powershell_version', [string]$PSVersionTable.PSVersion)
    $record.Add('windows_edition', $(if ($null -ne $os) { [string]$os.Caption } else { $unknown }))
    $record.Add('windows_version', $(if ($null -ne $os) { [string]$os.Version } else { $unknown }))
    $record.Add('windows_build', $(if ($null -ne $os) { [string]$os.BuildNumber } else { $unknown }))
    $record.Add('cpu_model', $(if ($null -ne $cpu) { [string]$cpu.Name } else { $unknown }))
    $record.Add('logical_processors',
                $(if ($null -ne $cpu -and $null -ne $cpu.NumberOfLogicalProcessors) {
                    [string]$cpu.NumberOfLogicalProcessors } else { $unknown }))
    $record.Add('installed_ram_gb', $ram)
    $record.Add('excel_version', $excelVersion)
    $record.Add('excel_build', $excelBuild)
    $record.Add('excel_bitness', $excelBitness)
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
    $record.Add('model_version', [string]$Manifest.model_version)
    $record.Add('builder_version', [string]$Manifest.builder_version)
    $record.Add('build_phase', [string]$Manifest.build_phase)
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
    $problems = @(Test-BenchmarkSample -Operation $Operation -Evidence $evidence `
        -RequestedIterations $RequestedIterations)

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

try {
    $startupWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $excel = New-Object -ComObject Excel.Application
    $excelIdentity = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $startupWatch.Stop()

    $openWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $workbooks = $excel.Workbooks
    $wb = $workbooks.Open($stageBPath)
    $openWatch.Stop()

    # EXCEL STARTUP AND WORKBOOK OPEN ARE RECORDED AND EXCLUDED. They are here so
    # a reader can see they were paid, and nowhere near an operation's elapsed
    # time. "Cold" never means either of them.
    $setupTimings.Add('excel_startup_ms', [double]$startupWatch.Elapsed.TotalMilliseconds)
    $setupTimings.Add('workbook_open_ms', [double]$openWatch.Elapsed.TotalMilliseconds)

    $environment = Get-BenchmarkEnvironment -Excel $excel -WorkbookPath $stageBPath `
        -RepositoryPath $repoRoot -Manifest $manifest `
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

    $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
    $null = Save-Phase5LockedFxSeed -Workbook $wb -Inspection $inspection

    # --- THE SCENARIO, THROUGH THE ACCEPTED FIXTURE ------------------------
    # SETUP. Timed, reported, and part of no measurement. At three hundred
    # drivers over forty project years this is the expensive part of the run,
    # and it is expensive exactly once.
    Write-BenchmarkLine ('BUILDING ' + $Scenario)
    Write-BenchmarkLine ('-' * (9 + $Scenario.Length))
    $model = New-BenchmarkModel -ScenarioSpec $scenarioSpec
    $fixtureWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $null = Set-Phase5Fixture -Excel $excel -Workbook $wb -Manifest $manifest `
        -Inspection $inspection -Model $model
    Set-NamedValue -Workbook $wb `
        -DefinedName ([string]$simInspection.controls.random_seed.defined_name) `
        -Value ([double](Get-BenchmarkSeed))
    $fixtureWatch.Stop()
    $setupTimings.Add('scenario_fixture_ms', [double]$fixtureWatch.Elapsed.TotalMilliseconds)

    # THE DIMENSIONS ARE READ BACK OUT OF THE WORKBOOK, not taken from the model.
    # What is timed is what the workbook holds.
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
        $iterations = $null
        if ($null -ne $run.iterations) { $iterations = [int]$run.iterations }

        $label = [string]$operation.label
        if ($null -ne $iterations) { $label = $label + '  @ ' + [string]$iterations + ' iterations' }
        Write-BenchmarkLine ($label)
        Write-BenchmarkLine ('-' * $label.Length)

        if (($null -ne $iterations) -and ($iterations -ne $currentIterations)) {
            # SETTING THE CONTROL IS SETUP. It is outside every clock, and the
            # value is read back so the run records what the workbook was asked
            # for rather than what this script intended.
            Set-NamedValue -Workbook $wb `
                -DefinedName ([string]$simInspection.controls.monte_carlo_iterations.defined_name) `
                -Value ([double]$iterations)
            $null = Invoke-Phase5ProductionOperation -Excel $excel -Operation 'PCCM_Calculate' `
                -Stage ('re-establishing the deterministic basis for ' + [string]$iterations +
                        ' iterations')
            $currentIterations = $iterations
        }
        $iterationsSet = [string](Get-NamedValue -Workbook $wb `
            -DefinedName ([string]$simInspection.controls.monte_carlo_iterations.defined_name))

        $samples = New-Object System.Collections.ArrayList
        $total = [int]$run.cold_runs + [int]$run.warm_runs
        for ($index = 0; $index -lt $total; $index++) {
            $phase = $(if ($index -lt [int]$run.cold_runs) { 'cold' } else { 'warm' })
            $execution = Invoke-BenchmarkExecution -Excel $excel -Workbook $wb `
                -SimInspection $simInspection -Operation $operation `
                -RequestedIterations $iterations -Phase $phase
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
        $row.Add('iterations_requested', $iterations)
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
    if ([string]::IsNullOrWhiteSpace($abandoned)) { $abandoned = (Format-Err $_) }
    Write-BenchmarkLine ''
    Write-BenchmarkLine ('THE BENCHMARK SESSION RAISED: ' + $abandoned)
    Write-BenchmarkLine 'Whatever was measured before this point is above and is still valid.'
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
    $environment = Get-BenchmarkEnvironment -Excel $null -WorkbookPath $stageBPath `
        -RepositoryPath $repoRoot -Manifest $manifest `
        -HarnessVersion ([string]$plan.harness_version) `
        -SchemaVersion ([int]$plan.schema_version) -Revision $revision
}

$report = New-Object System.Collections.Specialized.OrderedDictionary
$report.Add('schema_version', [int]$plan.schema_version)
$report.Add('baseline_id', [string]$plan.baseline_id)
$report.Add('harness_version', [string]$plan.harness_version)
$report.Add('kind', 'pccm-phase10-performance-baseline')
$report.Add('judgement', ('NONE. This artifact records measurements. No absolute pass ' +
                          'or fail is contracted before a baseline exists.'))
$report.Add('generated_at_utc', ((Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')))
$report.Add('environment', $environment)
$report.Add('scenario', $scenarioSpec)
$report.Add('scenario_actual', $actual)
$report.Add('setup_ms', $setupTimings)
$report.Add('setup_note', ('setup is reported so it can be seen to have been paid. ' +
                           'None of it is inside any operation elapsed time.'))
$report.Add('timing', $plan.timing)
$report.Add('correctness_gates', $plan.correctness_gates)
$report.Add('regression_policy', $plan.regression_policy)
$report.Add('historical_context', $plan.historical_context)
$report.Add('forbidden', $plan.forbidden)
$report.Add('results', @($results))
$report.Add('abandoned', $abandoned)
$report.Add('shutdown', $shutdownRecord)
$report.Add('session_wall_clock_ms', [double]$sessionWatch.Elapsed.TotalMilliseconds)

$jsonPath = Join-Path $OutDir ('phase10_benchmark_' + $Scenario + '_' + $stamp + '.json')
$report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $jsonPath -Encoding UTF8

# --- the human-readable summary -------------------------------------------
$md = New-Object System.Collections.ArrayList
$null = $md.Add('# PCCM performance baseline - ' + $Scenario)
$null = $md.Add('')
$null = $md.Add('**This records. It does not judge.** No absolute pass or fail is')
$null = $md.Add('contracted before a baseline exists on the target machine.')
$null = $md.Add('')
$null = $md.Add('| | |')
$null = $md.Add('|---|---|')
$null = $md.Add('| Baseline id | `' + [string]$plan.baseline_id + '` |')
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
    try { Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
}
Write-Host ''
Write-Host ('Benchmark artifacts in ' + $OutDir) -ForegroundColor Yellow
