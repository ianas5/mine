<#
.SYNOPSIS
    PCCM Stage-B bootstrap: turn the generated Stage-A .xlsx into the macro-enabled
    .xlsm, with the locked worksheet CodeNames, the Phase-4 VBA modules and the
    Phase-4 command buttons.

.DESCRIPTION
    Stage A (Linux, Python, openpyxl) produces PCCM_stageA.xlsx plus two generated
    inputs this script consumes:

        build/stage_b_manifest.json   sheet CodeNames, module list, button
                                      definitions, entry points, file format
        build/vba/modConstants.bas    the generated VBA constants module

    Nothing about the model is restated here. Every sheet name, CodeName, macro
    name, button caption and module name comes from the manifest, so this script
    cannot drift away from the contracts that produced it.

    What it does, in order:
      1. read and validate the manifest and the VBA sources
      2. open a NEW, owned Excel instance and capture its process identity
      3. open the Stage-A workbook and save it as .xlsm (FileFormat 52)
      4. apply the 14 locked worksheet CodeNames
      5. import every declared VBA module
      6. create or refresh the Phase-4 buttons and assign OnAction
      7. save, close, and release COM in explicit named order
      8. reopen in a FRESH Excel instance and verify what actually persisted
      9. close naturally

    It changes no security setting. It never force-stops an Excel process it did
    not create, and a forced stop is never reported as success.

.PARAMETER BuildDir
    The Stage-A build directory. Defaults to <repo>/pccm/build.

.PARAMETER Force
    Overwrite an existing Stage-B workbook. Without it, an existing file is left
    untouched and the run stops.
#>

[CmdletBinding()]
param(
    [string]$BuildDir,
    [switch]$Force
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptDir 'com_lifecycle.ps1')

$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
if ([string]::IsNullOrWhiteSpace($BuildDir)) { $BuildDir = Join-Path $pccmRoot 'build' }

$steps       = New-Object System.Collections.ArrayList
$failures    = New-Object System.Collections.ArrayList
$notes       = New-Object System.Collections.ArrayList

function Add-Step {
    param([string]$Name, [string]$Status, [string]$Detail = '')
    $null = $steps.Add([pscustomobject]@{ Name = $Name; Status = $Status; Detail = $Detail })
    if ($Status -eq 'FAIL') { $null = $failures.Add($Name) }
    $colour = 'Green'
    if ($Status -eq 'FAIL') { $colour = 'Red' } elseif ($Status -eq 'SKIP') { $colour = 'Yellow' }
    Write-Host ("  [{0}] {1}" -f $Status, $Name) -ForegroundColor $colour
    if ($Detail) { Write-Host ("        {0}" -f $Detail) -ForegroundColor DarkGray }
}

function Add-Note { param([string]$Text) $null = $notes.Add($Text) }

Write-Host ''
Write-Host 'PCCM - Stage-B bootstrap (.xlsx -> .xlsm)' -ForegroundColor Cyan
Write-Host '=========================================' -ForegroundColor Cyan
Write-Host ''

# ===========================================================================
# 1. Inputs
# ===========================================================================
$manifestPath = Join-Path $BuildDir 'stage_b_manifest.json'
$manifest     = $null
$stageAPath   = $null
$stageBPath   = $null
$moduleFiles  = @()

try {
    if (-not (Test-Path -LiteralPath $manifestPath)) {
        throw "stage_b_manifest.json not found at $manifestPath. Run the Stage-A build first: python3 pccm/builder/build_stage_a.py"
    }
    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json

    $stageAPath = Join-Path $BuildDir $manifest.stage_a_filename
    $stageBPath = Join-Path $BuildDir $manifest.stage_b_filename
    if (-not (Test-Path -LiteralPath $stageAPath)) {
        throw "Stage-A workbook not found at $stageAPath."
    }
    if ((Test-Path -LiteralPath $stageBPath) -and (-not $Force)) {
        throw "$stageBPath already exists. Re-run with -Force to replace it."
    }

    # Resolve every declared module to a file. A module the contract declares but
    # that is missing on disk stops the build here rather than producing a
    # half-populated VBA project.
    $srcDir  = Join-Path $pccmRoot $manifest.vba.source_dir
    $genDir  = Join-Path $pccmRoot $manifest.vba.generated_dir
    $missing = @()
    foreach ($m in $manifest.vba.modules) {
        $dir = $srcDir
        if ($m.generated) { $dir = $genDir }
        $file = Join-Path $dir ($m.name + '.bas')
        if (Test-Path -LiteralPath $file) { $moduleFiles += $file } else { $missing += $file }
    }
    if ($missing.Count -gt 0) { throw ("VBA module file(s) not found: " + ($missing -join ', ')) }

    Add-Step 'Read Stage-A build outputs' 'PASS' ("model {0}; structure contract {1}; {2} module(s)" -f $manifest.model_version, $manifest.structure_contract_version, $moduleFiles.Count)
} catch {
    Add-Step 'Read Stage-A build outputs' 'FAIL' (Format-Err $_)
    Write-Host ''
    Write-Host 'STAGE-B BOOTSTRAP FAILED before Excel was started.' -ForegroundColor Red
    exit 1
}

# ===========================================================================
# 2-7. Build instance
# ===========================================================================
$preExisting = Get-PreExistingExcelPids

# Explicit named ownership. Every long-lived COM object has its own variable and
# its own release point; there is no stack and no release plan.
$excel = $null; $workbooks = $null; $wb = $null; $worksheets = $null
$vbproj = $null; $vbcomps = $null
$id1 = $null
$rel1 = $null
$buildOk = $false

try {
    $excel = New-Object -ComObject Excel.Application
    $id1 = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    Add-Step 'Open an owned Excel instance' 'PASS' ("pid {0} (identity source {1})" -f $id1.ProcessId, $id1.Source)

    $workbooks = $excel.Workbooks
    $wb = $workbooks.Open($stageAPath)
    Add-Step 'Open the Stage-A workbook' 'PASS' $stageAPath

    # --- 3. save as .xlsm --------------------------------------------------
    if (Test-Path -LiteralPath $stageBPath) { Remove-Item -LiteralPath $stageBPath -Force }
    $wb.SaveAs($stageBPath, [int]$manifest.xlsm_file_format)
    $actualFormat = [int]$wb.FileFormat
    if ($actualFormat -ne [int]$manifest.xlsm_file_format) {
        throw ("SaveAs produced FileFormat {0}, expected {1}." -f $actualFormat, $manifest.xlsm_file_format)
    }
    Add-Step 'Save as macro-enabled .xlsm' 'PASS' ("FileFormat={0}; {1}" -f $actualFormat, $stageBPath)

    # --- 4. CodeNames -------------------------------------------------------
    $worksheets = $wb.Worksheets
    try {
        $vbproj = $wb.VBProject
    } catch {
        if (Test-TrustAccessError $_) {
            Add-Note (Get-TrustAccessGuidance)
            throw 'Excel refused programmatic access to the VBA project. See the guidance below.'
        }
        throw
    }
    $vbcomps = $vbproj.VBComponents

    $codeNameFails = @()
    foreach ($sheet in $manifest.sheets) {
        $ws = $null; $comp = $null; $props = $null; $prop = $null
        try {
            $ws = $worksheets.Item($sheet.name)
            $current = [string]$ws.CodeName
            if ($current -ne $sheet.codename) {
                $comp = $vbcomps.Item($current)
                $props = $comp.Properties
                $prop = $props.Item('_CodeName')
                $prop.Value = $sheet.codename
            }
            if ([string]$ws.CodeName -ne $sheet.codename) {
                $codeNameFails += ("{0}: CodeName is '{1}', expected '{2}'" -f $sheet.name, [string]$ws.CodeName, $sheet.codename)
            }
        } catch {
            $codeNameFails += ("{0}: {1}" -f $sheet.name, (Format-Err $_))
        } finally {
            if ($null -ne $prop)  { Release-Transient $prop  'Property(_CodeName)'; $prop  = $null }
            if ($null -ne $props) { Release-Transient $props 'Properties';          $props = $null }
            if ($null -ne $comp)  { Release-Transient $comp  'VBComponent';         $comp  = $null }
            if ($null -ne $ws)    { Release-Transient $ws    'Worksheet';           $ws    = $null }
        }
    }
    if ($codeNameFails.Count -gt 0) { throw ("CodeName assignment failed: " + ($codeNameFails -join '; ')) }
    Add-Step 'Apply the locked worksheet CodeNames' 'PASS' ("{0} sheets" -f $manifest.sheets.Count)

    # --- 5. import VBA ------------------------------------------------------
    # Remove any same-named component first so a re-run replaces rather than
    # duplicating (Excel would otherwise create modConstants1).
    foreach ($file in $moduleFiles) {
        $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($file)
        $existing = $null
        try {
            $existing = $vbcomps.Item($moduleName)
        } catch { $existing = $null }
        if ($null -ne $existing) {
            $vbcomps.Remove($existing)
            Release-Transient $existing 'VBComponent(existing)'; $existing = $null
        }
        $vbcomps.Import($file)
    }
    $importedNames = @()
    for ($i = 1; $i -le $vbcomps.Count; $i++) {
        $c = $null
        try {
            $c = $vbcomps.Item($i)
            $importedNames += [string]$c.Name
        } finally {
            if ($null -ne $c) { Release-Transient $c 'VBComponent(enum)'; $c = $null }
        }
    }
    $missingModules = @()
    foreach ($m in $manifest.vba.modules) {
        if ($importedNames -notcontains $m.name) { $missingModules += $m.name }
    }
    if ($missingModules.Count -gt 0) { throw ("VBA module(s) missing after import: " + ($missingModules -join ', ')) }
    Add-Step 'Import the Phase-4 VBA modules' 'PASS' (($manifest.vba.modules | ForEach-Object { $_.name }) -join ', ')

    # --- 6. buttons ---------------------------------------------------------
    foreach ($button in $manifest.buttons) {
        $ws = $null; $shapes = $null; $shp = $null; $anchor = $null; $tf = $null; $tr = $null; $existing = $null
        try {
            $ws = $worksheets.Item($button.sheet)
            $shapes = $ws.Shapes
            # Refresh rather than accumulate: a re-run must leave exactly one button.
            try {
                $existing = $shapes.Item($button.shape_name)
            } catch { $existing = $null }
            if ($null -ne $existing) {
                $existing.Delete()
                Release-Transient $existing 'Shape(existing)'; $existing = $null
            }
            $anchor = $ws.Range($button.anchor_cell)
            $shp = $shapes.AddShape(5, [double]$anchor.Left, [double]$anchor.Top, [double]$button.width, [double]$button.height)
            $shp.Name = $button.shape_name
            $tf = $shp.TextFrame2
            $tr = $tf.TextRange
            $tr.Text = $button.caption
            $shp.OnAction = $button.entry_point
            $readBack = [string]$shp.OnAction
            if ($readBack -notlike ('*' + $button.entry_point + '*')) {
                throw ("OnAction read back as '{0}', expected '{1}'." -f $readBack, $button.entry_point)
            }
        } finally {
            if ($null -ne $tr)     { Release-Transient $tr     'TextRange';  $tr     = $null }
            if ($null -ne $tf)     { Release-Transient $tf     'TextFrame2'; $tf     = $null }
            if ($null -ne $shp)    { Release-Transient $shp    'Shape';      $shp    = $null }
            if ($null -ne $anchor) { Release-Transient $anchor 'Range';      $anchor = $null }
            if ($null -ne $shapes) { Release-Transient $shapes 'Shapes';     $shapes = $null }
            if ($null -ne $ws)     { Release-Transient $ws     'Worksheet';  $ws     = $null }
        }
    }
    Add-Step 'Create the Phase-4 command buttons' 'PASS' (($manifest.buttons | ForEach-Object { $_.shape_name }) -join ', ')

    # --- 7. save ------------------------------------------------------------
    $wb.Save()
    Add-Step 'Save the Stage-B workbook' 'PASS' $stageBPath
    $buildOk = $true
} catch {
    Add-Step 'Stage-B build' 'FAIL' (Format-Err $_)
    $buildOk = $false
}

# --- shutdown of the build instance, leaf before parent --------------------
$rel1 = New-ReleaseLedger 'build instance'
try {
    Invoke-NamedRelease $rel1 $vbcomps    'VBComponents'; $vbcomps    = $null
    Invoke-NamedRelease $rel1 $vbproj     'VBProject';    $vbproj     = $null
    Invoke-NamedRelease $rel1 $worksheets 'Worksheets';   $worksheets = $null

    if ($null -ne $wb) {
        try { $wb.Close($false); $rel1.WorkbookClosed = $true } catch { $null = $rel1.Failed.Add('Workbook.Close') }
    }
    Invoke-NamedRelease $rel1 $wb        'Workbook';  $wb        = $null
    Invoke-NamedRelease $rel1 $workbooks 'Workbooks'; $workbooks = $null

    if ($null -ne $excel) {
        try { $excel.Quit(); $rel1.QuitCalled = $true } catch { $null = $rel1.Failed.Add('Application.Quit') }
    }
    Invoke-NamedRelease $rel1 $excel 'Application'; $excel = $null

    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

    $rel1.NaturalExit = Wait-ExcelExit -Identity $id1
    if (-not $rel1.NaturalExit) {
        $rel1.EmergencyRequired = $true
        Add-Note (Invoke-EmergencyExcelCleanup -Identity $id1 -Label 'build instance')
    }
} catch {
    Add-Note ('Shutdown of the build instance raised: ' + (Format-Err $_))
}

if ($rel1.NaturalExit -and $rel1.Failed.Count -eq 0) {
    Add-Step 'Build instance closed naturally' 'PASS' ("pid {0} exited without a forced stop" -f $id1.ProcessId)
} else {
    Add-Step 'Build instance closed naturally' 'FAIL' ("natural exit={0}; failed releases={1}" -f $rel1.NaturalExit, (($rel1.Failed | Select-Object -Unique) -join ', '))
}

# ===========================================================================
# 8. Verify in a FRESH instance
# ===========================================================================
$excel2 = $null; $workbooks2 = $null; $wb2 = $null; $worksheets2 = $null
$vbproj2 = $null; $vbcomps2 = $null
$id2 = $null
$rel2 = $null

if ($buildOk) {
    $preExisting2 = Get-PreExistingExcelPids
    try {
        $excel2 = New-Object -ComObject Excel.Application
        $id2 = Get-ExcelIdentity -ExcelApp $excel2 -PreExistingPids $preExisting2
        $excel2.Visible = $false
        $excel2.DisplayAlerts = $false
        $excel2.AskToUpdateLinks = $false

        $workbooks2 = $excel2.Workbooks
        $wb2 = $workbooks2.Open($stageBPath)

        $problems = @()

        if ([int]$wb2.FileFormat -ne [int]$manifest.xlsm_file_format) {
            $problems += ("FileFormat is {0}, expected {1}" -f [int]$wb2.FileFormat, $manifest.xlsm_file_format)
        }

        $worksheets2 = $wb2.Worksheets
        foreach ($sheet in $manifest.sheets) {
            $ws = $null
            try {
                $ws = $worksheets2.Item($sheet.name)
                if ([string]$ws.CodeName -ne $sheet.codename) {
                    $problems += ("{0}: CodeName persisted as '{1}', expected '{2}'" -f $sheet.name, [string]$ws.CodeName, $sheet.codename)
                }
            } catch {
                $problems += ("{0}: {1}" -f $sheet.name, (Format-Err $_))
            } finally {
                if ($null -ne $ws) { Release-Transient $ws 'Worksheet2'; $ws = $null }
            }
        }

        $vbproj2 = $wb2.VBProject
        $vbcomps2 = $vbproj2.VBComponents
        $persisted = @()
        for ($i = 1; $i -le $vbcomps2.Count; $i++) {
            $c = $null
            try { $c = $vbcomps2.Item($i); $persisted += [string]$c.Name }
            finally { if ($null -ne $c) { Release-Transient $c 'VBComponent2(enum)'; $c = $null } }
        }
        foreach ($m in $manifest.vba.modules) {
            if ($persisted -notcontains $m.name) { $problems += ("VBA module '{0}' did not persist" -f $m.name) }
        }

        foreach ($button in $manifest.buttons) {
            $ws = $null; $shapes = $null; $shp = $null
            try {
                $ws = $worksheets2.Item($button.sheet)
                $shapes = $ws.Shapes
                $shp = $shapes.Item($button.shape_name)
                $onAction = [string]$shp.OnAction
                if ($onAction -notlike ('*' + $button.entry_point + '*')) {
                    $problems += ("{0}: OnAction persisted as '{1}', expected '{2}'" -f $button.shape_name, $onAction, $button.entry_point)
                }
            } catch {
                $problems += ("{0}: {1}" -f $button.shape_name, (Format-Err $_))
            } finally {
                if ($null -ne $shp)    { Release-Transient $shp    'Shape2';     $shp    = $null }
                if ($null -ne $shapes) { Release-Transient $shapes 'Shapes2';    $shapes = $null }
                if ($null -ne $ws)     { Release-Transient $ws     'Worksheet2'; $ws     = $null }
            }
        }

        if ($problems.Count -gt 0) {
            Add-Step 'Verify the reopened .xlsm' 'FAIL' ($problems -join '; ')
        } else {
            Add-Step 'Verify the reopened .xlsm' 'PASS' ("{0} CodeNames, {1} modules, {2} buttons persisted" -f $manifest.sheets.Count, $manifest.vba.modules.Count, $manifest.buttons.Count)
        }
    } catch {
        Add-Step 'Verify the reopened .xlsm' 'FAIL' (Format-Err $_)
    }

    $rel2 = New-ReleaseLedger 'verification instance'
    try {
        Invoke-NamedRelease $rel2 $vbcomps2    'VBComponents2'; $vbcomps2    = $null
        Invoke-NamedRelease $rel2 $vbproj2     'VBProject2';    $vbproj2     = $null
        Invoke-NamedRelease $rel2 $worksheets2 'Worksheets2';   $worksheets2 = $null

        if ($null -ne $wb2) {
            try { $wb2.Close($false); $rel2.WorkbookClosed = $true } catch { $null = $rel2.Failed.Add('Workbook2.Close') }
        }
        Invoke-NamedRelease $rel2 $wb2        'Workbook2';  $wb2        = $null
        Invoke-NamedRelease $rel2 $workbooks2 'Workbooks2'; $workbooks2 = $null

        if ($null -ne $excel2) {
            try { $excel2.Quit(); $rel2.QuitCalled = $true } catch { $null = $rel2.Failed.Add('Application2.Quit') }
        }
        Invoke-NamedRelease $rel2 $excel2 'Application2'; $excel2 = $null

        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

        $rel2.NaturalExit = Wait-ExcelExit -Identity $id2
        if (-not $rel2.NaturalExit) {
            $rel2.EmergencyRequired = $true
            Add-Note (Invoke-EmergencyExcelCleanup -Identity $id2 -Label 'verification instance')
        }
    } catch {
        Add-Note ('Shutdown of the verification instance raised: ' + (Format-Err $_))
    }

    if ($rel2.NaturalExit -and $rel2.Failed.Count -eq 0) {
        Add-Step 'Verification instance closed naturally' 'PASS' ("pid {0} exited without a forced stop" -f $id2.ProcessId)
    } else {
        Add-Step 'Verification instance closed naturally' 'FAIL' ("natural exit={0}; failed releases={1}" -f $rel2.NaturalExit, (($rel2.Failed | Select-Object -Unique) -join ', '))
    }
} else {
    Add-Step 'Verify the reopened .xlsm' 'SKIP' 'the build did not complete'
}

# ===========================================================================
# Report
# ===========================================================================
$transient = Get-TransientFailures
if ($transient.Count -gt 0) {
    Add-Step 'Transient COM releases' 'FAIL' ($transient -join '; ')
} else {
    Add-Step 'Transient COM releases' 'PASS' 'every transient object released cleanly'
}

Write-Host ''
Write-Host 'Shutdown ledgers' -ForegroundColor Cyan
Write-Host '----------------'
Write-Host '  build instance:'
Write-Host (Format-ReleaseLedger $rel1)
Write-Host '  verification instance:'
Write-Host (Format-ReleaseLedger $rel2)

if ($notes.Count -gt 0) {
    Write-Host ''
    Write-Host 'Notes' -ForegroundColor Yellow
    Write-Host '-----'
    foreach ($n in $notes) { Write-Host ("  " + $n) }
}

Write-Host ''
if ($failures.Count -eq 0) {
    Write-Host ("STAGE-B BOOTSTRAP COMPLETE: {0}" -f $stageBPath) -ForegroundColor Green
    Write-Host ''
    exit 0
}
Write-Host ("STAGE-B BOOTSTRAP FAILED: {0}" -f (($failures | Select-Object -Unique) -join ', ')) -ForegroundColor Red
Write-Host ''
exit 1
