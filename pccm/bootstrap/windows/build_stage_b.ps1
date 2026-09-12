<#
.SYNOPSIS
    PCCM Stage-B bootstrap: turn the generated Stage-A .xlsx into the macro-enabled
    .xlsm, with the locked worksheet CodeNames, every manifest-declared VBA module
    and the command buttons.

.DESCRIPTION
    Stage A (Linux, Python, openpyxl) produces PCCM_stageA.xlsx, the manifest, and
    the generated VBA projections this script consumes:

        build/stage_b_manifest.json   sheet CodeNames, module list, button
                                      definitions, entry points, file format
        build/vba/*.bas               the generated VBA projection modules

    A module is generated when the MANIFEST says so - `generated: true` on its
    vba.modules entry - and this script reads that flag rather than a list of its
    own. At the time of writing the generated set is modConstants, modCalcContract
    and modSimContract; that is the CURRENT INVENTORY, not a dependency. Adding or
    removing a generated projection is a contract change and needs no edit here.

    Nothing about the model is restated here. Every sheet name, CodeName, macro
    name, button caption and module name comes from the manifest, so this script
    cannot drift away from the contracts that produced it.

    What it does, in order:
      1. read and validate the manifest and the VBA sources
      2. open a NEW, owned Excel instance and capture its process identity
      3. open the Stage-A workbook and save it as .xlsm (FileFormat 52)
      4. apply the 14 locked worksheet CodeNames
      5. import every manifest-declared VBA module, source and generated alike
      6. create or refresh the declared command buttons and assign OnAction
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

# ===========================================================================
# WHICH COM CALL WAS REFUSED
# ===========================================================================
# TWO CONSECUTIVE EQUIVALENCE RUNS DIED IN THE BUILD BLOCK AND THE TRANSCRIPT
# COULD NOT SAY WHERE. The block is one try/catch around eleven distinct COM
# operations, and its catch reported the whole region by one name:
#
#     [FAIL] Stage-B build
#            System.Runtime.InteropServices.COMException: Call was rejected by
#            callee. (0x80010001 RPC_E_CALL_REJECTED)
#
# Twice, identically, and there was no way to tell SaveAs from VBProject from a
# button. The region was identified; the CALL was not.
#
# SO THE BLOCK NAMES WHAT IT IS DOING. One variable, set immediately before each
# operation, out of a CLOSED vocabulary - a label that is not on the list throws
# HERE, rather than turning up in a diagnostic nobody can map back to a call
# site. It retries nothing, waits for nothing and decides nothing. Its whole job
# is that the NEXT Windows run names the refused call instead of narrowing it to
# a region, because the correction that follows depends entirely on which call
# it was: a refused property get may simply be reissued, and a refused SaveAs,
# Import or AddShape may not.
$script:StageBBuildOps = @(
    'open.workbook'
    'saveas.xlsm'
    'worksheets.acquire'
    'codename.write'
    'vbproject.acquire'
    'vbcomponents.acquire'
    'vbcomponents.import'
    'thisworkbook.write'
    'button.add'
    'protection.apply'
    'workbook.save'
)
# NOT one of the labels, and deliberately not blank: a failure before the first
# labelled operation must read as "no operation had begun", which is a different
# finding from "the label was lost".
$script:StageBBuildOp         = '<before the first labelled operation>'
$script:StageBBuildRejections = New-Object System.Collections.ArrayList

function Set-StageBBuildOp {
    param([string]$Operation)
    if ($script:StageBBuildOps -notcontains $Operation) {
        throw ('Set-StageBBuildOp: ' + $Operation + ' is not in the closed Stage-B build ' +
               'operation vocabulary (' + ($script:StageBBuildOps -join ', ') + ').')
    }
    $script:StageBBuildOp = $Operation
}

function Get-StageBBuildOp { return [string]$script:StageBBuildOp }

function Get-StageBBuildRejections { return @($script:StageBBuildRejections) }

# The HRESULT as HEX, from the first COMException in the chain, or '' when the
# error is not a COM failure at all. PRINTED EITHER WAY, because "the call was
# refused" and "a VBA error came back" are different findings and reporting them
# in the same words is how four runs produced one sentence.
function Get-StageBComHResult {
    param($ErrorRecord)
    if ($null -eq $ErrorRecord) { return '' }
    $ex = $null
    try { $ex = $ErrorRecord.Exception } catch { return '' }
    # BOUNDED, like the classifier it sits beside: a cyclic InnerException chain
    # would otherwise hang a diagnostic.
    for ($depth = 0; $depth -lt 5; $depth++) {
        if ($null -eq $ex) { return '' }
        if ($ex -is [System.Runtime.InteropServices.COMException]) {
            try { return ('0x' + ([int]$ex.ErrorCode).ToString('x8')) } catch { return '' }
        }
        $next = $null
        try { $next = $ex.InnerException } catch { $next = $null }
        $ex = $next
    }
    return ''
}

# ONE line, and it says WHICH OF THE TWO IT IS. A refused call never ran, so it
# may be reissued; a call Excel accepted and failed describes something that
# actually happened. The classifier that draws that line is the accepted one in
# com_lifecycle.ps1 - this only formats its answer.
function New-StageBRejectionLine {
    param([string]$Operation, $ErrorRecord)
    $shown = Get-StageBComHResult $ErrorRecord
    if ($shown -eq '') { $shown = 'none' }
    $name = Get-ComRejectionName $ErrorRecord
    if ([string]::IsNullOrWhiteSpace($name)) {
        return ('COMFAIL|build|' + $Operation + '|hresult=' + $shown +
                '|the call was accepted and failed; it was NOT refused')
    }
    return ('COMREJECT|build|' + $Operation + '|hresult=' + $shown + '|' + $name +
            '|the call was refused before it ran')
}

# BUILD-ONLY, READ-ONLY, AND NARROWER THAN THE HELPER IT FORWARDS TO. It takes an
# object and a MEMBER NAME and hands both to the accepted Invoke-ComRetryRead
# unchanged: same two retryable HRESULTs, same three bounds, same original error
# rethrown on exhaustion. There is no scriptblock parameter and no -Key either,
# so the only thing expressible through this is a plain PROPERTY GET - not a
# write, not an Open, not a SaveAs, not even an Item lookup. All it adds is the
# operation label and one concise line when a read actually had to be reissued.
function Invoke-StageBBuildRead {
    param($Target, [string]$Member, [string]$Operation, [string]$Description)
    $null = Set-StageBBuildOp $Operation
    $record = Invoke-ComRetryRead -Target $Target -Member $Member -Description $Description
    if ([int]$record.Attempts -gt 1) {
        $null = $script:StageBBuildRejections.Add(
            ('COMREJECT|build|' + $Operation + '|attempts=' + [string]$record.Attempts +
             '|waited=' + [string]$record.WaitedMs + '|' + [string]$record.Rejections +
             '|answered'))
    }
    # NOTHING IS NOT AN ANSWER. A PowerShell host can DISCARD an exception thrown
    # by a property getter and hand back $null instead of raising - proved on this
    # repository's own harness host. If Excel's adapter ever behaves that way, the
    # operation that was refused must still be the one that gets named here,
    # rather than surfacing three statements later as a null-reference on an
    # object nobody can trace back to a call.
    if ($null -eq $record.Value) {
        throw ('Invoke-StageBBuildRead: ' + $Description + ' answered with nothing at ' +
               $Operation + '. The read was not refused and it was not answered.')
    }
    return $record
}

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
    #
    # The two directories resolve DIFFERENTLY, and the difference matters:
    #
    #   source modules     repository-relative. They are version-controlled input
    #                      and are the same files whatever build is being assembled.
    #   generated modules  BUILD-DIRECTORY-relative. A generated projection is
    #                      emitted beside the workbook, the manifest and the
    #                      scenario fixture it belongs to, so it must come from the
    #                      SUPPLIED BuildDir.
    #
    # Resolving a generated module against the repository root instead meant the
    # functional harness copied a disposable build to %TEMP% and then imported
    # modConstants.bas from the real repository build -- testing a different
    # generated source than the manifest and scenarios sitting beside the workbook.
    $srcDir  = Join-Path $pccmRoot $manifest.vba.source_dir
    $genDir  = Join-Path $BuildDir (Split-Path -Leaf $manifest.vba.generated_dir)
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
$preExisting = @(Get-PreExistingExcelPids)

# Explicit named ownership. Every long-lived COM object has its own variable and
# its own release point; there is no stack and no release plan.
$excel = $null; $workbooks = $null; $wb = $null; $worksheets = $null
$vbproj = $null; $vbcomps = $null
$buildExcelIdentity = $null
$rel1 = $null
$buildOk = $false
# THE LEDGER IS PROCESS-WIDE AND NOW HAS TWO READERS. Without a baseline the
# verification step would report the BUILD's reissued reads as its own, and
# "no verification read was refused" would stop being true.
$buildRetryBase     = @(Get-ComRetryLedger).Count
$buildRetryWaitBase = Get-ComRetryWaitTotal

try {
    $excel = New-Object -ComObject Excel.Application
    $buildExcelIdentity = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    Add-Step 'Open an owned Excel instance' 'PASS' ("pid {0} (identity source {1})" -f $buildExcelIdentity.ProcessId, $buildExcelIdentity.Source)

    Set-StageBBuildOp 'open.workbook'
    $workbooks = $excel.Workbooks
    $wb = $workbooks.Open($stageAPath)
    Add-Step 'Open the Stage-A workbook' 'PASS' $stageAPath

    # --- 3. save as .xlsm --------------------------------------------------
    if (Test-Path -LiteralPath $stageBPath) { Remove-Item -LiteralPath $stageBPath -Force }
    # THE SaveAs ITSELF IS NOT RETRIED and must not be. A SaveAs Excel may or may
    # not have accepted is not something to reissue on a guess; if this is the
    # refused call, the label above is what says so and the correction is a
    # separate decision.
    Set-StageBBuildOp 'saveas.xlsm'
    $wb.SaveAs($stageBPath, [int]$manifest.xlsm_file_format)
    $actualFormat = [int](Invoke-StageBBuildRead -Target $wb -Member 'FileFormat' `
                              -Operation 'saveas.xlsm' `
                              -Description 'the Stage-B workbook FileFormat').Value
    if ($actualFormat -ne [int]$manifest.xlsm_file_format) {
        throw ("SaveAs produced FileFormat {0}, expected {1}." -f $actualFormat, $manifest.xlsm_file_format)
    }
    Add-Step 'Save as macro-enabled .xlsm' 'PASS' ("FileFormat={0}; {1}" -f $actualFormat, $stageBPath)

    # --- 4. CodeNames -------------------------------------------------------
    # THREE PLAIN PROPERTY GETS, AND THE REOPEN PATH ALREADY RETRIES ALL THREE.
    # Worksheets, VBProject and VBComponents are acquisitions: they read a member
    # and move nothing. The verification block below reissues exactly these
    # members on exactly these classes of object, so opting the BUILD's copies in
    # adds no new judgement - it applies one already accepted, at the reads that
    # sit closest to the workbook opening.
    $worksheets = (Invoke-StageBBuildRead -Target $wb -Member 'Worksheets' `
                       -Operation 'worksheets.acquire' `
                       -Description 'the Stage-B workbook Worksheets collection').Value
    try {
        $vbproj = (Invoke-StageBBuildRead -Target $wb -Member 'VBProject' `
                       -Operation 'vbproject.acquire' `
                       -Description 'the Stage-B workbook VBProject').Value
    } catch {
        # STILL REACHED. A Trust Center refusal is not a message-filter rejection,
        # so the helper rethrows it on the first attempt and this guidance path is
        # exactly as available as it was before the read was wrapped.
        if (Test-TrustAccessError $_) {
            Add-Note (Get-TrustAccessGuidance)
            throw 'Excel refused programmatic access to the VBA project. See the guidance below.'
        }
        throw
    }
    $vbcomps = (Invoke-StageBBuildRead -Target $vbproj -Member 'VBComponents' `
                    -Operation 'vbcomponents.acquire' `
                    -Description 'the Stage-B VBComponents collection').Value

    Set-StageBBuildOp 'codename.write'
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
    # duplicating (Excel would otherwise create a modConstants1 beside it).
    Set-StageBBuildOp 'vbcomponents.import'
    foreach ($file in $moduleFiles) {
        $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($file)
        # The acquire/use path is wrapped. With the release written inline after
        # $vbcomps.Remove($existing), a throw from Remove skipped it and left an owned
        # VBComponent RCW alive on the exception path -- the same unowned-intermediate
        # failure the readiness gate ruled out, just reached through an error path.
        # Once acquired, $existing is released and nulled whatever happens next.
        $existing = $null
        try {
            try {
                $existing = $vbcomps.Item($moduleName)
            } catch { $existing = $null }
            if ($null -ne $existing) {
                $vbcomps.Remove($existing)
            }
        } finally {
            if ($null -ne $existing) { Release-Transient $existing 'VBComponent(existing)'; $existing = $null }
        }
        # Import returns the VBComponent it created. Discarding that return left an
        # unowned RCW alive, which is exactly the invisible-intermediate pattern the
        # readiness gate ruled out. It is named, released once, and nulled here.
        $imported = $null
        try {
            $imported = $vbcomps.Import($file)
        } finally {
            if ($null -ne $imported) { Release-Transient $imported 'VBComponent(imported)'; $imported = $null }
        }
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
    Add-Step 'Import every manifest-declared VBA module' 'PASS' (($manifest.vba.modules | ForEach-Object { $_.name }) -join ', ')

    # --- 5b. the ThisWorkbook document module -------------------------------
    # IT CANNOT BE IMPORTED. Excel creates ThisWorkbook with the workbook, so
    # there is no component to replace - only one to write into. The existing
    # code is cleared first so a re-run leaves exactly one copy of the handler
    # rather than appending a second Workbook_Open beside it, which VBA would
    # accept at import time and refuse at compile time.
    $docModule = $manifest.vba.document_module
    if ($null -ne $docModule) {
        $docFile = Join-Path $srcDir ([string]$docModule.file)
        if (-not (Test-Path -LiteralPath $docFile)) {
            throw ("The document module source is missing: " + $docFile)
        }
        $docText = Get-Content -LiteralPath $docFile -Raw
        $docComp = $null; $codeModule = $null
        Set-StageBBuildOp 'thisworkbook.write'
        try {
            $docComp = $vbcomps.Item([string]$docModule.component)
            $codeModule = $docComp.CodeModule
            if ([int]$codeModule.CountOfLines -gt 0) {
                $codeModule.DeleteLines(1, [int]$codeModule.CountOfLines)
            }
            $codeModule.AddFromString($docText)
            # READ BACK, exactly as the buttons read back their OnAction. A
            # component that silently kept its old text would leave a workbook
            # that opens unprotected and says nothing.
            $written = [string]$codeModule.Lines(1, [int]$codeModule.CountOfLines)
            foreach ($event in @($docModule.events)) {
                if ($written -notlike ('*' + [string]$event + '*')) {
                    throw ("The document module was written but does not contain " + [string]$event + '.')
                }
            }
            if ($written -notlike ('*' + [string]$docModule.delegates_to + '*')) {
                throw ("The document module does not delegate to " + [string]$docModule.delegates_to + '.')
            }
        } finally {
            if ($null -ne $codeModule) { Release-Transient $codeModule 'CodeModule';   $codeModule = $null }
            if ($null -ne $docComp)    { Release-Transient $docComp    'VBComponent(doc)'; $docComp = $null }
        }
        Add-Step 'Write the ThisWorkbook document module' 'PASS' `
            ([string]$docModule.component + ': ' + (@($docModule.events) -join ', '))
    }

    # --- 6. buttons ---------------------------------------------------------
    Set-StageBBuildOp 'button.add'
    foreach ($button in $manifest.buttons) {
        $ws = $null; $shapes = $null; $shp = $null; $anchor = $null; $tf = $null; $tr = $null; $existing = $null
        try {
            $ws = $worksheets.Item($button.sheet)
            $shapes = $ws.Shapes
            # Refresh rather than accumulate: a re-run must leave exactly one button.
            try {
                $existing = $shapes.Item($button.shape_name)
            } catch { $existing = $null }
            # No inline release here: $existing.Delete() can throw, and an inline
            # release would then be skipped, leaking the Shape RCW. It is released in
            # the enclosing finally instead -- before $shapes, its parent.
            if ($null -ne $existing) {
                $existing.Delete()
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
            # Leaf before parent: the pre-existing Shape releases before $shapes.
            if ($null -ne $existing) { Release-Transient $existing 'Shape(existing)'; $existing = $null }
            if ($null -ne $shapes) { Release-Transient $shapes 'Shapes';     $shapes = $null }
            if ($null -ne $ws)     { Release-Transient $ws     'Worksheet';  $ws     = $null }
        }
    }
    Add-Step 'Create the Phase-4 command buttons' 'PASS' (($manifest.buttons | ForEach-Object { $_.shape_name }) -join ', ')

    # --- 7. protection ------------------------------------------------------
    # LAST, AND THAT ORDER IS THE POINT. Every step above writes cells, shapes,
    # code names and VBA components; protecting before them would have to
    # unprotect again for each one, and the workbook would spend the build in a
    # state no control could describe. The locked/unlocked state of each CELL is
    # already in the file - Stage A resolved it from the contracts that declare
    # which inputs are editable - so this applies the ACTION and decides nothing.
    #
    # NO PASSWORD ARGUMENT ANYWHERE. Not an empty string, not a variable.
    $protection = $manifest.protection
    if ($null -ne $protection) {
        if (-not $protection.passwordless) {
            throw 'The manifest asks for protection with a password. This build has none and will not invent one.'
        }
        Set-StageBBuildOp 'protection.apply'
        $protectFails = @()
        foreach ($sheetName in @($protection.sheets)) {
            $pws = $null
            try {
                $pws = $worksheets.Item([string]$sheetName)
                if ($pws.ProtectContents) { $pws.Unprotect() }
                # Worksheet.Protect(Password, DrawingObjects, Contents,
                # Scenarios, UserInterfaceOnly, ...). [Type]::Missing is how a
                # COM call says "no password" - $null would be marshalled as an
                # empty one, which is a password.
                $pws.Protect([Type]::Missing, $true, $true, $false, $true) | Out-Null
                if (-not $pws.ProtectContents) {
                    $protectFails += ([string]$sheetName + ': protection did not take')
                }
            } catch {
                $protectFails += ([string]$sheetName + ': ' + (Format-Err $_))
            } finally {
                if ($null -ne $pws) { Release-Transient $pws 'Worksheet(protect)'; $pws = $null }
            }
        }
        if ($protectFails.Count -gt 0) {
            throw ('Sheet protection failed: ' + ($protectFails -join '; '))
        }
        if ($protection.protect_structure -and (-not $wb.ProtectStructure)) {
            $wb.Protect([Type]::Missing, $true, $false)
        }
        if ($protection.protect_structure -and (-not $wb.ProtectStructure)) {
            throw 'Workbook structure protection did not take.'
        }
        Add-Step 'Apply passwordless protection' 'PASS' `
            ([string]@($protection.sheets).Count + ' sheet(s); structure=' +
             [string]$wb.ProtectStructure + '; UserInterfaceOnly=True')
    }

    # --- 8. save ------------------------------------------------------------
    Set-StageBBuildOp 'workbook.save'
    $wb.Save()
    Add-Step 'Save the Stage-B workbook' 'PASS' $stageBPath
    $buildOk = $true
} catch {
    # THE OPERATION, NOT THE REGION. This is the line that four runs could not
    # answer, and the reason the exact rejected call is still unknown.
    $failedOp = Get-StageBBuildOp
    Add-Step 'Stage-B build' 'FAIL' ('operation=' + $failedOp + '; ' + (Format-Err $_))
    # A DIAGNOSTIC MAY NOT REPLACE THE FAILURE IT DESCRIBES. If classifying the
    # error throws, the build failure above still stands and is still reported.
    try   { Add-Note (New-StageBRejectionLine -Operation $failedOp -ErrorRecord $_) }
    catch { Add-Note ('COMFAIL|build|' + $failedOp + '|hresult=unclassified|the error could ' +
                      'not be classified; the build failure above stands') }
    $buildOk = $false
}

# REPORTED ON EVERY RUN, REFUSED OR NOT. A silent clean build and a build that
# spent nine seconds reissuing three refused reads looked identical in the last
# four transcripts, so "nothing was refused" is now a printed line too. It
# describes what happened and decides nothing: a build that failed stays failed.
$buildRejections = @(Get-StageBBuildRejections)
$buildLedger     = @(@(Get-ComRetryLedger) | Select-Object -Skip $buildRetryBase)
$buildRetryWait  = (Get-ComRetryWaitTotal) - $buildRetryWaitBase
if ($buildLedger.Count -eq 0 -and $buildRejections.Count -eq 0) {
    Add-Step 'Transient COM rejections (build)' 'PASS' 'COMREJECT|build|none|attempts=0|waited=0'
} else {
    Add-Step 'Transient COM rejections (build)' 'PASS' `
        ("{0} build read(s) were refused and reissued; {1} ms waited in total" -f $buildLedger.Count, $buildRetryWait)
    foreach ($line in $buildRejections) { Add-Note $line }
    foreach ($line in $buildLedger)     { Add-Note ('COM read (build): ' + $line) }
}

# The verification step reports ITS OWN slice, from here on.
$verifyRetryBase     = @(Get-ComRetryLedger).Count
$verifyRetryWaitBase = Get-ComRetryWaitTotal

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

    $rel1.NaturalExit = Wait-ExcelExit -Identity $buildExcelIdentity
    if (-not $rel1.NaturalExit) {
        $rel1.EmergencyRequired = $true
        Add-Note (Invoke-EmergencyExcelCleanup -Identity $buildExcelIdentity -Label 'build instance')
    }
} catch {
    Add-Note ('Shutdown of the build instance raised: ' + (Format-Err $_))
}

if ($rel1.NaturalExit -and $rel1.Failed.Count -eq 0) {
    Add-Step 'Build instance closed naturally' 'PASS' ("pid {0} exited without a forced stop" -f $buildExcelIdentity.ProcessId)
} else {
    Add-Step 'Build instance closed naturally' 'FAIL' ("natural exit={0}; failed releases={1}" -f $rel1.NaturalExit, (($rel1.Failed | Select-Object -Unique) -join ', '))
}

# ===========================================================================
# 8. Verify in a FRESH instance
# ===========================================================================
$excel2 = $null; $workbooks2 = $null; $wb2 = $null; $worksheets2 = $null
$vbproj2 = $null; $vbcomps2 = $null
$verifyExcelIdentity = $null
$rel2 = $null

if ($buildOk) {
    $preExisting2 = @(Get-PreExistingExcelPids)
    try {
        $excel2 = New-Object -ComObject Excel.Application
        $verifyExcelIdentity = Get-ExcelIdentity -ExcelApp $excel2 -PreExistingPids $preExisting2
        $excel2.Visible = $false
        $excel2.DisplayAlerts = $false
        $excel2.AskToUpdateLinks = $false

        $workbooks2 = $excel2.Workbooks
        $wb2 = $workbooks2.Open($stageBPath)

        $problems = @()

        # EVERY READ IN THIS BLOCK GOES THROUGH Invoke-ComRetryRead, and that is
        # the whole change. The reopened workbook has just run Workbook_Open -
        # the first VBA this bootstrap ever executes - and twice in a row Excel
        # refused an incoming call here with RPC_E_CALL_REJECTED. A refused call
        # never ran, so reissuing it is safe; the helper reissues nothing else.
        #
        # NOTHING IS VERIFIED LESS. The same 14 CodeNames, the same 32 modules
        # and the same 11 buttons are read from the same reopened instance, which
        # still opens exactly as it will for a person: macros enabled, events on,
        # protection applied by Workbook_Open, nothing suppressed to make the
        # verification easier. Only the delivery of each read is retried.
        #
        # The BUILD block above is deliberately NOT wrapped. Its calls write -
        # SaveAs, Import, AddShape, Protect, Save - and a write that Excel may or
        # may not have accepted is not something to reissue on a guess.
        $ff = [int](Invoke-ComRetryRead -Target $wb2 -Member 'FileFormat' `
                        -Description 'the reopened workbook FileFormat').Value
        if ($ff -ne [int]$manifest.xlsm_file_format) {
            $problems += ("FileFormat is {0}, expected {1}" -f $ff, $manifest.xlsm_file_format)
        }

        $worksheets2 = (Invoke-ComRetryRead -Target $wb2 -Member 'Worksheets' `
                            -Description 'the reopened workbook Worksheets collection').Value
        foreach ($sheet in $manifest.sheets) {
            $ws = $null
            try {
                $ws = (Invoke-ComRetryRead -Target $worksheets2 -Member 'Item' -Key ([string]$sheet.name) `
                           -Description ("Worksheets.Item('" + [string]$sheet.name + "')")).Value
                $codeName = [string](Invoke-ComRetryRead -Target $ws -Member 'CodeName' `
                                         -Description ([string]$sheet.name + '.CodeName')).Value
                if ($codeName -ne $sheet.codename) {
                    $problems += ("{0}: CodeName persisted as '{1}', expected '{2}'" -f $sheet.name, $codeName, $sheet.codename)
                }
            } catch {
                $problems += ("{0}: {1}" -f $sheet.name, (Format-Err $_))
            } finally {
                if ($null -ne $ws) { Release-Transient $ws 'Worksheet2'; $ws = $null }
            }
        }

        $vbproj2 = (Invoke-ComRetryRead -Target $wb2 -Member 'VBProject' `
                        -Description 'the reopened workbook VBProject').Value
        $vbcomps2 = (Invoke-ComRetryRead -Target $vbproj2 -Member 'VBComponents' `
                         -Description 'the reopened VBComponents collection').Value
        # Read ONCE, not once per iteration. The old loop condition re-read .Count
        # on every pass - fourteen-odd unguarded COM reads where one suffices.
        $compCount = [int](Invoke-ComRetryRead -Target $vbcomps2 -Member 'Count' `
                               -Description 'VBComponents.Count').Value
        $persisted = @()
        for ($i = 1; $i -le $compCount; $i++) {
            $c = $null
            try {
                $c = (Invoke-ComRetryRead -Target $vbcomps2 -Member 'Item' -Key $i `
                          -Description ('VBComponents.Item(' + [string]$i + ')')).Value
                $persisted += [string](Invoke-ComRetryRead -Target $c -Member 'Name' `
                                           -Description ('VBComponents.Item(' + [string]$i + ').Name')).Value
            }
            finally { if ($null -ne $c) { Release-Transient $c 'VBComponent2(enum)'; $c = $null } }
        }
        foreach ($m in $manifest.vba.modules) {
            if ($persisted -notcontains $m.name) { $problems += ("VBA module '{0}' did not persist" -f $m.name) }
        }

        foreach ($button in $manifest.buttons) {
            $ws = $null; $shapes = $null; $shp = $null
            try {
                $ws = (Invoke-ComRetryRead -Target $worksheets2 -Member 'Item' -Key ([string]$button.sheet) `
                           -Description ("Worksheets.Item('" + [string]$button.sheet + "') for " + [string]$button.shape_name)).Value
                $shapes = (Invoke-ComRetryRead -Target $ws -Member 'Shapes' `
                               -Description ([string]$button.sheet + '.Shapes')).Value
                $shp = (Invoke-ComRetryRead -Target $shapes -Member 'Item' -Key ([string]$button.shape_name) `
                            -Description ("Shapes.Item('" + [string]$button.shape_name + "')")).Value
                $onAction = [string](Invoke-ComRetryRead -Target $shp -Member 'OnAction' `
                                         -Description ([string]$button.shape_name + '.OnAction')).Value
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

    # REPORTED WHETHER OR NOT ANY READ WAS REFUSED, and reported as its own step
    # so a retried run can never look like an untroubled one. This step describes
    # what happened; it decides nothing. A verification that failed above stays
    # failed, and an exhausted retry appears here AND as that failure.
    $retryLines = @(@(Get-ComRetryLedger) | Select-Object -Skip $verifyRetryBase)
    if ($retryLines.Count -eq 0) {
        Add-Step 'Transient COM rejections' 'PASS' 'no verification read was refused; 0 ms waited'
    } else {
        Add-Step 'Transient COM rejections' 'PASS' `
            ("{0} read(s) were refused and reissued; {1} ms waited in total" -f $retryLines.Count, ((Get-ComRetryWaitTotal) - $verifyRetryWaitBase))
        foreach ($line in $retryLines) { Add-Note ('COM read: ' + $line) }
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

        $rel2.NaturalExit = Wait-ExcelExit -Identity $verifyExcelIdentity
        if (-not $rel2.NaturalExit) {
            $rel2.EmergencyRequired = $true
            Add-Note (Invoke-EmergencyExcelCleanup -Identity $verifyExcelIdentity -Label 'verification instance')
        }
    } catch {
        Add-Note ('Shutdown of the verification instance raised: ' + (Format-Err $_))
    }

    if ($rel2.NaturalExit -and $rel2.Failed.Count -eq 0) {
        Add-Step 'Verification instance closed naturally' 'PASS' ("pid {0} exited without a forced stop" -f $verifyExcelIdentity.ProcessId)
    } else {
        Add-Step 'Verification instance closed naturally' 'FAIL' ("natural exit={0}; failed releases={1}" -f $rel2.NaturalExit, (($rel2.Failed | Select-Object -Unique) -join ', '))
    }
} else {
    Add-Step 'Verify the reopened .xlsm' 'SKIP' 'the build did not complete'
}

# ===========================================================================
# Report
# ===========================================================================
# @(...) AT THE CALLER, not inside the helper. A PowerShell function that returns an
# empty collection emits ZERO pipeline objects, so the assignment lands $null however
# the helper wrote the return -- and under Set-StrictMode reading .Count on $null
# raises PropertyNotFoundException. That fired on the first real Gate-B run, on the
# SUCCESS path specifically, because zero transient failures is what produces it.
$transient = @(Get-TransientFailures)
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
