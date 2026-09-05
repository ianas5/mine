<#
.SYNOPSIS
    PCCM Phase 7 - the MINIMAL W1 runner: the project compiles, and the public
    command surface is really there.

.DESCRIPTION
    THIS REPLACES `phase7_acceptance_scenarios.ps1` AS THE WINDOWS EXECUTION
    AUTHORITY FOR W1. That harness is left in history as evidence of the
    attempted acceptance matrix; it is not patched again and it is not run.
    Three separate execution-layer defects came out of its size and its
    inherited dependency graph, so this runner inherits neither.

    WHAT IT PROVES, AND NOTHING ELSE
    --------------------------------
      1. the candidate identity: git HEAD, with pccm/src, pccm/spec and
         pccm/builder clean, proved BEFORE Excel is started
      2. a Stage-B workbook generated from the CURRENT Stage-A build, in a
         disposable copy - the real build output is never opened or saved over
      3. every manifest-declared module is loaded in the project
      4. the COMPLETE VBAProject compiles in real Excel
      5. the required Phase-7 public procedures exist in their expected modules
      6. the four safe read-only accessors are callable and answer the unrun
         state
      7. the owned Excel process exits NATURALLY
      8. no emergency cleanup was required

    There is no W2-W8 logic here, no fixture choreography, no simulation, no
    sensitivity run and no annual computation. A procedure's presence in its
    module plus a compiled project is the whole of the W1 claim.

    THE ONE DOT-SOURCE IS `com_lifecycle.ps1`, which is definition-only at top
    level: dot-sourcing it defines functions and runs nothing. The Stage-B
    bootstrap is invoked as a CHILD PROCESS, exactly as the accepted harnesses
    invoke it, so none of its top-level code runs in this scope. That is the
    entire external dependency surface of this file.

    THE TWO DEFECTS OF THE FAILED W1, AND WHAT WAS DONE ABOUT THEM
    -------------------------------------------------------------
    FALSE PROCEDURE-EXISTENCE FAILURES. The old detector walked a module line
    by line from line 1 and asked `CodeModule.ProcOfLine`. Line 1 of every
    module is in the DECLARATIONS section - Option Explicit, the Public Consts,
    the Public Types - and ProcOfLine REFUSES a line that is not inside a
    procedure, raising "Sub or Function not defined". The refusal happened on
    the first line of all four modules, the caller caught it, and every one of
    the eight checks failed carrying that same message. The name it was given
    was never the problem: the module-qualified form appeared only in the
    CHECK LABEL, and the probe itself was always passed a bare name.

    So this runner never asks a CodeModule to resolve a line, a name, or
    anything else. It reads the module's text ONCE - `CodeModule.Lines` - and
    matches a VBA procedure DECLARATION for the bare name in that specific
    module's text. That call cannot refuse, which is what makes the detector
    total: it answers YES or NO, and never a third thing that a catch block has
    to turn into a failure.

    THE RETAINED COM REFERENCE, AND WHAT THE FIRST W1 RUN SETTLED. PID 27384
    was first explained by `$Error`-rooted RCWs: a raising COM call leaves an
    ErrorRecord that still references the object the call was made on, and
    `$Error` keeps 256 of them for the life of the session. That explanation is
    now REFUTED for the second run. At 408cb5d no COM call raised - 57 checks,
    55 of them green - `$Error` was cleared before the collect, and PID 42716
    still had to be force-stopped. Three more explanations are refuted too, and
    each by something checkable rather than by argument:

      * the automation guard holds no object. PCCM_AutomationBegin sets five
        module-level Booleans and Strings and nothing else.
      * no COM object crosses Application.Run. Every procedure this runner
        calls is declared `As String`, `As Long`, or is a Sub, so the VARIANT
        that comes back is a BSTR, an I4 or Empty.
      * no acquisition is missing a release. There are 37 acquisitions in the
        session and 37 releases, each in a `finally`.

    WHAT WAS ACTUALLY WRONG WAS THE EVIDENCE. The shutdown ledger represented
    THREE objects - Workbook, Workbooks, Application. The other 34 - VBProject,
    VBComponents, one VBComponent per module and one CodeModule per inspected
    module - were released through `Release-Transient`, which records NOTHING
    when a release succeeds: not the label, not the count. So "every ledgered
    release succeeded" was a true statement about 3 objects out of 37, and the
    one number that names a retained reference - what `ReleaseComObject`
    returned - was never observed for any of the 34.

    `Marshal.ReleaseComObject` RETURNS THE REFERENCE COUNT REMAINING ON THAT
    RCW. Zero means the runtime let go of it. Anything above zero means the
    same underlying object was marshalled more times than it was released, and
    that is the signature of the retained reference, per object, by name. So
    every release in this runner now goes into the ledger with its count, and a
    non-zero count is a FAILING CHECK that names the class. A run that still
    will not exit with every count at zero is not holding a managed reference
    at all, and the report will say so instead of leaving it open.

    THE OTHER TWO HIDDEN-REFERENCE PATTERNS ARE REFUSED STRUCTURALLY, and
    controls hold them refused: this runner never chains two COM property
    accesses in one expression - `$a.B.C` acquires an intermediate that no
    variable holds and no `finally` can release - and it never enumerates a COM
    collection with `foreach`, which acquires an IEnumVARIANT nobody releases
    either. Every component is fetched by index into a named variable.

.PARAMETER BuildDir
    The Stage-A build directory to copy from. Defaults to <repo>/pccm/build.

.PARAMETER KeepArtifacts
    Reserved for symmetry with the accepted harnesses; the working copy is left
    in place either way so the report survives the run.

.NOTES
    WINDOWS POWERSHELL 5.1 is the target shell. `Join-Path a b c` - a child per
    positional argument - is PowerShell 6+ only, and the roots below are
    resolved the way the harnesses that have actually run on 5.1 resolve them.

    SHUTDOWN is the accepted `com_lifecycle.ps1` path: Workbook.Close($false),
    Application.Quit, named releases leaf before parent, Wait-ExcelExit, and
    Invoke-EmergencyExcelCleanup ONLY for a process whose identity is still
    positively verified - and never as a pass.
#>

[CmdletBinding()]
param(
    [string]$BuildDir,
    [switch]$KeepArtifacts
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# THE ONLY DOT-SOURCE. Definition-only at top level, and the owner of every
# ownership, release, identity and shutdown decision this runner makes.
. (Join-Path $scriptDir 'com_lifecycle.ps1')

$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
$repoRoot = Split-Path -Parent $pccmRoot
if ([string]::IsNullOrWhiteSpace($BuildDir)) { $BuildDir = Join-Path $pccmRoot 'build' }

# ===========================================================================
# THE DECLARED W1 SURFACE
# ===========================================================================
# The eight required procedures and the module each must be in. Module
# PLACEMENT is not projected by any contract - only the command names are - so
# the table is declared here and a Linux control pins it to both the projection
# and the .bas sources it describes.
function Get-W1RequiredSurface {
    return @(
        [pscustomobject]@{ Module = 'modSimAnnualRun'
                           Procedures = @('PCCM_RunAnnualStochastic') },
        [pscustomobject]@{ Module = 'modSimAnnualStore'
                           Procedures = @('PCCM_AnnualDistributionState',
                                          'PCCM_AnnualProfileState',
                                          'PCCM_AnnualProfilePx',
                                          'PCCM_AnnualYearCount') },
        [pscustomobject]@{ Module = 'modSimPostReport'
                           Procedures = @('PCCM_RunSensitivity') },
        [pscustomobject]@{ Module = 'modSimReport'
                           Procedures = @('PCCM_RunSimulation',
                                          'PCCM_SimulationStatus') }
    )
}

# THE NEGATIVE CONTROL'S SUBJECT. A name no module can contain, probed in every
# expected module in the same session as the real ones. A detector that answers
# the same for this and for a real procedure has proved nothing, and this run
# fails rather than reporting a surface it cannot actually see.
$script:W1AbsentProcedure = 'PCCM_W1NegativeControl_NoSuchProcedure'

# ===========================================================================
# THE REPORT AND THE CHECK LEDGER
# ===========================================================================
$script:W1Lines = New-Object System.Collections.ArrayList
$script:W1Path = ''
$script:W1Checks = New-Object System.Collections.ArrayList

function Write-W1Line {
    param([string]$Text = '')
    $null = $script:W1Lines.Add($Text)
    Write-Host $Text
    # Written through on every line, so a run that is stopped still leaves every
    # observation it had already taken on disk.
    if (-not [string]::IsNullOrWhiteSpace($script:W1Path)) {
        try {
            Set-Content -LiteralPath $script:W1Path `
                -Value ($script:W1Lines -join "`r`n") -Encoding UTF8
        } catch { }
    }
}

function Add-W1Check {
    param([string]$Label, [bool]$Ok, [string]$Detail = '')
    $null = $script:W1Checks.Add([pscustomobject]@{ Label = $Label; Ok = $Ok; Detail = $Detail })
    $verdict = 'FAIL'
    if ($Ok) { $verdict = 'PASS' }
    $line = '  [' + $verdict + '] ' + $Label
    if (-not [string]::IsNullOrWhiteSpace($Detail)) { $line = $line + ' -- ' + $Detail }
    Write-W1Line $line
    return $Ok
}

function Format-W1Value {
    param($Value)
    if ($null -eq $Value) { return '<null>' }
    if ($Value -is [System.DBNull]) { return '<empty>' }
    $text = ''
    try { $text = [string]$Value } catch { return '<unprintable>' }
    if ([string]::IsNullOrEmpty($text)) { return '<blank>' }
    return $text
}

# ===========================================================================
# THE CANDIDATE IDENTITY
# ===========================================================================
# An acceptance result taken from a modified tree names nothing, so this is
# proved before Excel is started and refuses the run outright.
function Get-W1SourceRevision {
    param([string]$RepoRoot)
    $head = ''
    try { $head = [string](& git -C $RepoRoot rev-parse HEAD 2>$null) } catch { $head = '' }
    $head = $head.Trim()
    if ([string]::IsNullOrWhiteSpace($head)) {
        throw ('git could not report HEAD for ' + $RepoRoot +
               '; a W1 result could not be attributed to a source revision')
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
# THE PROCEDURE DETECTOR
# ===========================================================================
# A VBA procedure DECLARATION for the BARE name, at the start of a line of the
# module's own text. It is a pure string function: no COM, no Excel, no state,
# and it cannot raise - so it returns an answer for a name that is absent
# instead of an error a caller has to interpret, which is exactly what the old
# detector could not do.
#
# A comment cannot match: a VBA comment line starts with an apostrophe or Rem,
# and the anchor requires the declaration keywords first. `End Sub` cannot
# match either. The trailing guard refuses a longer identifier, so
# PCCM_AnnualYearCount is not found by a probe for PCCM_AnnualYear.
function Test-W1ProcedureDeclared {
    param([string]$ModuleText, [string]$ProcedureName)
    if ([string]::IsNullOrEmpty($ModuleText)) { return $false }
    if ([string]::IsNullOrWhiteSpace($ProcedureName)) { return $false }
    $escaped = [regex]::Escape($ProcedureName)
    $pattern = '(?im)^[ \t]*(?:Public[ \t]+|Private[ \t]+|Friend[ \t]+)?(?:Static[ \t]+)?' +
               '(?:Sub|Function|Property[ \t]+(?:Get|Let|Set))[ \t]+' + $escaped + '(?![A-Za-z0-9_])'
    return [regex]::IsMatch($ModuleText, $pattern)
}

# ===========================================================================
# EVERY RELEASE, WITH ITS COUNT
# ===========================================================================
# `Invoke-NamedRelease` writes the accepted ledger line but keeps the count to
# itself, and `Release-Transient` says nothing at all on success - which is how
# 34 of this run's 37 releases produced no evidence. This wraps the SAME
# accepted primitive, `Release-ComObjectSafe`, writes the SAME ledger line, and
# additionally records the one number that identifies a retained reference:
# what ReleaseComObject had left. com_lifecycle.ps1 is not modified.
$script:W1Residual = New-Object System.Collections.ArrayList

function Invoke-W1Release {
    param($Ledger, $Obj, [string]$Label)
    $rec = Release-ComObjectSafe -Obj $Obj -Label $Label
    if ($rec.Status -eq 'PASS') {
        $Ledger.Attempted = $Ledger.Attempted + 1
        $Ledger.Succeeded = $Ledger.Succeeded + 1
        $null = $Ledger.Lines.Add(
            ("      {0,-24} | PASS    | ReleaseComObject returned {1}" -f $rec.Label, $rec.Count))
        if ([int]$rec.Count -ne 0) {
            $null = $script:W1Residual.Add(
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
# THE ONE COM PASS OVER THE PROJECT
# ===========================================================================
# VBProject and VBComponents are acquired ONCE for the whole run, each
# VBComponent and CodeModule is acquired and released inside the narrowest
# possible scope, and what comes back out is PLAIN DATA: a list of names and a
# map of name to text. No COM object crosses this boundary, so nothing later in
# the run can hold one alive by accident.
function Read-W1VbaProject {
    param($Workbook, [string[]]$WantedModules, $Ledger)
    $result = [pscustomobject]@{
        Loaded       = @()
        Texts        = @{}
        Problem      = ''
        TrustRefused = $false
        Acquired     = 0
    }
    $vbproj = $null; $comps = $null
    try {
        # ONE HOP PER STATEMENT, ALWAYS INTO A NAMED VARIABLE. Writing this as
        # $Workbook.VBProject.VBComponents would acquire a VBProject that no
        # variable holds, so no finally could release it and no ledger line
        # could record it - a retained reference by construction.
        $vbproj = $Workbook.VBProject
        $result.Acquired = $result.Acquired + 1
        $comps = $vbproj.VBComponents
        $result.Acquired = $result.Acquired + 1
        $names = New-Object System.Collections.ArrayList
        $total = [int]$comps.Count
        # BY INDEX, NEVER BY foreach. Enumerating a COM collection acquires an
        # IEnumVARIANT that PowerShell never hands back and nothing releases.
        for ($index = 1; $index -le $total; $index++) {
            $comp = $null; $code = $null
            try {
                $comp = $comps.Item($index)
                $result.Acquired = $result.Acquired + 1
                $name = [string]$comp.Name
                $null = $names.Add($name)
                if ($WantedModules -contains $name) {
                    $code = $comp.CodeModule
                    $result.Acquired = $result.Acquired + 1
                    $lineCount = [int]$code.CountOfLines
                    $text = ''
                    if ($lineCount -ge 1) { $text = [string]$code.Lines(1, $lineCount) }
                    $result.Texts[$name] = $text
                }
            } finally {
                if ($null -ne $code) {
                    Invoke-W1Release $Ledger $code ('CodeModule[' + [string]$index + ']')
                    $code = $null
                }
                if ($null -ne $comp) {
                    Invoke-W1Release $Ledger $comp ('VBComponent[' + [string]$index + ']')
                    $comp = $null
                }
            }
        }
        $result.Loaded = @($names)
    } catch {
        # THE TRUST-CENTRE QUESTION IS ANSWERED HERE, where the real ErrorRecord
        # is, and only a boolean leaves this scope: an ErrorRecord that escaped
        # would keep the object the failed call was made on alive.
        $result.TrustRefused = (Test-TrustAccessError $_)
        $result.Problem = (Format-Err $_)
    } finally {
        if ($null -ne $comps)  { Invoke-W1Release $Ledger $comps  'VBComponents'; $comps  = $null }
        if ($null -ne $vbproj) { Invoke-W1Release $Ledger $vbproj 'VBProject';    $vbproj = $null }
    }
    return $result
}

# ===========================================================================
# PREFLIGHT
# ===========================================================================
Write-Host ''
Write-Host 'PCCM - Phase 7 W1 (minimal runner)' -ForegroundColor Cyan
Write-Host '==================================' -ForegroundColor Cyan
Write-Host ''

$manifestPath  = Join-Path $BuildDir 'stage_b_manifest.json'
$p7InspectPath = Join-Path $BuildDir 'phase7_acceptance_inspection.json'
foreach ($required in @($manifestPath, $p7InspectPath)) {
    if (-not (Test-Path -LiteralPath $required)) {
        Write-Host ($required + ' not found. Run the Stage-A build first: ' +
                    'python3 pccm/builder/build_stage_a.py') -ForegroundColor Red
        exit 1
    }
}
$manifest = Get-Content -LiteralPath $manifestPath  -Raw | ConvertFrom-Json
$p7       = Get-Content -LiteralPath $p7InspectPath -Raw | ConvertFrom-Json

$revision = $null
try { $revision = Get-W1SourceRevision -RepoRoot $repoRoot }
catch { Write-Host (Format-Err $_) -ForegroundColor Red; exit 1 }
if ($revision.Dirty.Count -gt 0) {
    Write-Host 'REFUSED, BEFORE EXCEL WAS STARTED.' -ForegroundColor Red
    Write-Host ''
    Write-Host ('pccm/src, pccm/spec or pccm/builder is modified, so a W1 result could ' +
                'not be attributed to a source revision:') -ForegroundColor Red
    foreach ($line in $revision.Dirty) { Write-Host ('    ' + $line) -ForegroundColor Red }
    exit 1
}

# ===========================================================================
# A DISPOSABLE COPY OF THE BUILD, AND THE STAGE-B BOOTSTRAP
# ===========================================================================
$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) `
    ('pccm-phase7-w1-' + (Get-Date).ToString('yyyyMMdd-HHmmss'))
$null = New-Item -ItemType Directory -Path $tempRoot -Force
Copy-Item -LiteralPath (Join-Path $BuildDir ([string]$manifest.stage_a_filename)) -Destination $tempRoot
Copy-Item -LiteralPath $manifestPath  -Destination $tempRoot
Copy-Item -LiteralPath $p7InspectPath -Destination $tempRoot
Copy-Item -LiteralPath (Join-Path $BuildDir 'vba') -Destination $tempRoot -Recurse

$script:W1Path = Join-Path $tempRoot 'phase7_w1_smoke.txt'
$stageAPath = Join-Path $tempRoot ([string]$manifest.stage_a_filename)
$stageBPath = Join-Path $tempRoot ([string]$manifest.stage_b_filename)

# INVOKED AS A CHILD PROCESS, never dot-sourced: build_stage_b.ps1 has
# executable top-level code, and this runner takes on none of it.
$bootstrap = Join-Path $scriptDir 'build_stage_b.ps1'
& $bootstrap -BuildDir $tempRoot -Force
$bootstrapExit = $LASTEXITCODE
$bootstrapOk = (($bootstrapExit -eq 0) -and (Test-Path -LiteralPath $stageBPath))

Write-W1Line 'PCCM - PHASE 7 W1: COMPILE AND PUBLIC SURFACE'
Write-W1Line '============================================'
Write-W1Line ''
Write-W1Line 'This is the MINIMAL W1 runner. It is not Gate B, it is not the Phase-7'
Write-W1Line 'acceptance matrix, and it records no result about W2-W8. The historical'
Write-W1Line 'Phase-6 runtime authority remains Run 6 on its own closure commit.'
Write-W1Line ''
Write-W1Line ('run started            : ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
Write-W1Line ('host                   : ' + [string]$env:COMPUTERNAME)
Write-W1Line ('PowerShell             : ' + [string]$PSVersionTable.PSVersion)
Write-W1Line ('git HEAD               : ' + [string]$revision.Head)
Write-W1Line  'pccm/src, pccm/spec, pccm/builder : clean (proved before Excel was started)'
Write-W1Line ('model version          : ' + [string]$manifest.model_version)
Write-W1Line ('sim contract version   : ' + [string]$p7.provenance.sim_contract_version)
Write-W1Line ('build directory        : ' + $BuildDir)
Write-W1Line ('working copy           : ' + $tempRoot)
$artefacts = @(
    [pscustomobject]@{ Label = 'Stage-A workbook'; Path = $stageAPath },
    [pscustomobject]@{ Label = 'Stage-B workbook'; Path = $stageBPath }
)
foreach ($artefact in $artefacts) {
    $shown = '(not present)'
    if (Test-Path -LiteralPath ([string]$artefact.Path)) {
        try { $shown = [string](Get-FileHash -LiteralPath ([string]$artefact.Path) -Algorithm SHA256).Hash }
        catch { $shown = '(unreadable)' }
    }
    Write-W1Line ('  ' + ([string]$artefact.Label).PadRight(20) + ' SHA-256 ' + $shown)
}
Write-W1Line ''

$null = Add-W1Check 'the Stage-B workbook was generated from the current Stage-A build' `
    $bootstrapOk ('bootstrap exit ' + [string]$bootstrapExit)
if (-not $bootstrapOk) {
    Write-W1Line ''
    Write-W1Line 'STOP. Excel was never started for the W1 session, and nothing was accepted.'
    Write-W1Line ('report                 : ' + $script:W1Path)
    Write-W1Line 'W1: FAIL'
    exit 1
}

# ===========================================================================
# THE W1 SESSION
# ===========================================================================
$surface = @(Get-W1RequiredSurface)
$wantedModules = @($surface | ForEach-Object { [string]$_.Module })

$preExisting = @(Get-PreExistingExcelPids)
$excel = $null; $workbooks = $null; $wb = $null
$excelIdentity = $null
$naturalExit = $false
$emergencyRequired = $false
$fatal = ''
$comAcquired = 0

# THE LEDGER IS CREATED BEFORE EXCEL EXISTS, so that every release of the run -
# the VBE objects included - lands in ONE record with its count. Building it in
# the shutdown, as the first version did, is precisely why 34 of 37 releases
# were never represented anywhere.
$rel = New-ReleaseLedger 'phase 7 W1 session'

try {
    $excel = New-Object -ComObject Excel.Application
    $comAcquired = $comAcquired + 1
    $excelIdentity = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    Write-W1Line ('EXCEL PROCESS OWNERSHIP: this run created PID ' +
                  [string]$excelIdentity.ProcessId + '. No process it did not create is')
    Write-W1Line 'ever terminated, and the workbook is never saved.'
    Write-W1Line ''
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $workbooks = $excel.Workbooks
    $comAcquired = $comAcquired + 1
    $wb = $workbooks.Open($stageBPath)
    $comAcquired = $comAcquired + 1

    # AUTOMATION ON FIRST, no failure stage armed. A modal confirmation on a
    # headless run would hang until a person closed it, which is the one way
    # this runner could still orphan a process.
    #
    # IT IS ALSO THE FIRST Application.Run OF THE SESSION, and VBA compiles the
    # whole project before it executes any statement - so this call is where a
    # compile failure would actually surface. It is guarded and fed into the
    # compile check below rather than allowed to escape, because a compile
    # failure reported as an unexplained fatal is a compile failure with no
    # evidence attached.
    $automationProblem = ''
    try { $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null }
    catch { $automationProblem = (Format-Err $_) }

    # -------------------------------------------------------------------
    # 3. EVERY MANIFEST-DECLARED MODULE IS LOADED
    # -------------------------------------------------------------------
    Write-W1Line 'MODULES, COMPILE, SURFACE'
    Write-W1Line '-------------------------'
    $project = Read-W1VbaProject -Workbook $wb -WantedModules $wantedModules -Ledger $rel
    $comAcquired = $comAcquired + [int]$project.Acquired
    # COLLECTED HERE, WHILE EXCEL IS STILL HEALTHY, rather than only at
    # shutdown. Anything the runtime created and dropped during the VBE pass is
    # finalised now, so the shutdown measures the objects this runner actually
    # named instead of racing a finaliser queue.
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    if ($project.TrustRefused) { Write-W1Line (Get-TrustAccessGuidance) }
    $declaredModules = @($manifest.vba.modules | ForEach-Object { [string]$_.name })
    $missingModules = @($declaredModules | Where-Object { $project.Loaded -notcontains $_ })
    $null = Add-W1Check ('all ' + [string]$declaredModules.Count +
                         ' manifest-declared modules are loaded in the project') `
        (($missingModules.Count -eq 0) -and [string]::IsNullOrWhiteSpace($project.Problem)) `
        (('missing: ' + ($missingModules -join ', ') + ' ' + $project.Problem).Trim())

    # -------------------------------------------------------------------
    # 4. THE COMPLETE VBAProject COMPILES
    # -------------------------------------------------------------------
    # THE MECHANISM ALREADY OBSERVED TO WORK, unchanged: VBA compiles the whole
    # project before it executes any statement, so the first Application.Run IS
    # a full-project compile. The subject is a published READ accessor - it runs
    # no simulation, writes nothing and consumes no run identity. The historical
    # VBE modal-dialog route is not reopened: it raises a dialog on failure, and
    # a headless run cannot answer one.
    $compileFailure = $automationProblem
    if ([string]::IsNullOrWhiteSpace($compileFailure)) {
        try { $null = $excel.Run([string]$p7.command_surface.handoff_accessors[0]) }
        catch { $compileFailure = (Format-Err $_) }
    }
    $compiled = [string]::IsNullOrWhiteSpace($compileFailure)
    $null = Add-W1Check 'the complete VBAProject compiles in real Excel' $compiled $compileFailure
    $null = Add-W1Check 'the automation guard is on for the whole session' `
        ([string]::IsNullOrWhiteSpace($automationProblem)) $automationProblem
    if (-not $compiled) {
        Write-W1Line ''
        Write-W1Line 'STOP. Nothing after a compile failure is evidence about anything.'
        throw ('the VBAProject does not compile: ' + $compileFailure)
    }

    # -------------------------------------------------------------------
    # 5. THE REQUIRED PROCEDURES, IN THEIR EXPECTED MODULES
    # -------------------------------------------------------------------
    # The declared table is checked against the contract projection first, so a
    # command the contract has moved or renamed cannot be silently missed by a
    # runner still looking for the old one.
    $projectedCommands = @([string]$p7.command_surface.annual_endpoint)
    foreach ($accessor in @($p7.command_surface.handoff_accessors)) { $projectedCommands += [string]$accessor }
    $declaredProcedures = @()
    foreach ($entry in $surface) { $declaredProcedures += @($entry.Procedures) }
    $unclaimed = @($projectedCommands | Where-Object { $declaredProcedures -notcontains $_ })
    $null = Add-W1Check 'the declared W1 surface covers every contract-projected command' `
        ($unclaimed.Count -eq 0) ('not claimed by any module: ' + ($unclaimed -join ', '))

    $positives = 0
    $negatives = 0
    foreach ($entry in $surface) {
        $moduleName = [string]$entry.Module
        $text = ''
        if ($project.Texts.ContainsKey($moduleName)) { $text = [string]$project.Texts[$moduleName] }
        $null = Add-W1Check ('the code of ' + $moduleName + ' was read out of the compiled project') `
            (-not [string]::IsNullOrWhiteSpace($text))

        foreach ($wanted in @($entry.Procedures)) {
            $found = Test-W1ProcedureDeclared -ModuleText $text -ProcedureName $wanted
            if ($found) { $positives = $positives + 1 }
            $null = Add-W1Check ($wanted + ' is declared in ' + $moduleName) $found

            # THE SHARPER NEGATIVE CONTROL: the same detector, the same session,
            # the same real name - asked of the modules it is NOT in. The failed
            # W1 could not tell one module from another, so a detector that says
            # yes everywhere is refused here.
            foreach ($other in $surface) {
                if ([string]$other.Module -eq $moduleName) { continue }
                $otherText = ''
                if ($project.Texts.ContainsKey([string]$other.Module)) {
                    $otherText = [string]$project.Texts[[string]$other.Module]
                }
                $elsewhere = Test-W1ProcedureDeclared -ModuleText $otherText -ProcedureName $wanted
                if (-not $elsewhere) { $negatives = $negatives + 1 }
                $null = Add-W1Check ($wanted + ' is NOT declared in ' + [string]$other.Module) `
                    (-not $elsewhere)
            }
        }

        # AND THE ABSENT NAME, in every expected module.
        $sentinel = Test-W1ProcedureDeclared -ModuleText $text -ProcedureName $script:W1AbsentProcedure
        if (-not $sentinel) { $negatives = $negatives + 1 }
        $null = Add-W1Check ('the detector does NOT find ' + $script:W1AbsentProcedure +
                             ' in ' + $moduleName) (-not $sentinel)
    }
    $null = Add-W1Check 'the procedure detector discriminates: it answered YES and NO in this session' `
        (($positives -gt 0) -and ($negatives -gt 0)) `
        ([string]$positives + ' present, ' + [string]$negatives + ' absent')

    # -------------------------------------------------------------------
    # 6. THE SAFE READ-ONLY ACCESSORS ARE CALLABLE
    # -------------------------------------------------------------------
    # Only these four, and only on the fresh workbook. Nothing here runs a
    # simulation, a sensitivity pass or an annual computation: presence in the
    # module plus a compiled project is the whole existence claim for the rest.
    Write-W1Line ''
    Write-W1Line 'THE SAFE ACCESSORS ON AN UNRUN WORKBOOK'
    Write-W1Line '---------------------------------------'
    $accessors = @($p7.command_surface.handoff_accessors | ForEach-Object { [string]$_ })
    $expectations = @(
        [pscustomobject]@{ Name = $accessors[0]; Kind = 'TEXT'
                           Expect = [string]$p7.handoff.distribution_states[0] },
        [pscustomobject]@{ Name = $accessors[1]; Kind = 'TEXT'
                           Expect = [string]$p7.handoff.profile_states[0] },
        [pscustomobject]@{ Name = $accessors[2]; Kind = 'BLANK'; Expect = '' },
        [pscustomobject]@{ Name = $accessors[3]; Kind = 'ZERO';  Expect = '0' }
    )
    foreach ($expectation in $expectations) {
        $value = $null
        $problem = ''
        try { $value = $excel.Run([string]$expectation.Name) } catch { $problem = (Format-Err $_) }
        $callable = [string]::IsNullOrWhiteSpace($problem)
        $null = Add-W1Check ([string]$expectation.Name + ' is callable through Application.Run') `
            $callable $problem
        if (-not $callable) { continue }
        $answered = $false
        if ([string]$expectation.Kind -eq 'TEXT') {
            $answered = (([string]$value) -ceq [string]$expectation.Expect)
        } elseif ([string]$expectation.Kind -eq 'BLANK') {
            $answered = (($null -eq $value) -or ($value -is [System.DBNull]) -or
                         [string]::IsNullOrEmpty([string]$value))
        } else {
            try { $answered = ([double]$value -eq 0) } catch { $answered = $false }
        }
        $wanted = [string]$expectation.Expect
        if ([string]$expectation.Kind -eq 'BLANK') { $wanted = 'blank' }
        $null = Add-W1Check ([string]$expectation.Name + ' answers the unrun state (' + $wanted + ')') `
            $answered ('returned ' + (Format-W1Value $value))
    }

    # THE AUTOMATION GUARD IS TURNED OFF THE WAY THE PROVEN DRIVER TURNS IT OFF.
    # The Phase-4/5/6 driver - the one harness whose owned Excel has exited
    # naturally on this machine - ends its session with this call, and neither
    # the frozen Phase-7 harness nor the first version of this runner did. It
    # holds no COM reference either way (ClearAutomation resets five scalars),
    # so this is parity with the proven path, not a theory about the leak.
    try { $excel.Run('PCCM_AutomationEnd') | Out-Null } catch { }
} catch {
    $fatal = (Format-Err $_)
    Write-W1Line ''
    Write-W1Line ('THE W1 SESSION DID NOT COMPLETE: ' + $fatal)
} finally {
    # THE ACCEPTED SHUTDOWN PATH, against the ledger this run has been writing
    # to since before Excel existed.
    try {
        if ($null -ne $wb) {
            try { $wb.Close($false); $rel.WorkbookClosed = $true }
            catch { $null = $rel.Failed.Add('Workbook.Close') }
        }
        Invoke-W1Release $rel $wb        'Workbook';  $wb        = $null
        Invoke-W1Release $rel $workbooks 'Workbooks'; $workbooks = $null
        if ($null -ne $excel) {
            try { $excel.Quit(); $rel.QuitCalled = $true }
            catch { $null = $rel.Failed.Add('Application.Quit') }
        }
        Invoke-W1Release $rel $excel 'Excel.Application'; $excel = $null
    } finally {
        # $Error IS CLEARED BEFORE THE COLLECT, and this is the line that keeps
        # PID 27384 from happening again. An ErrorRecord from a failed COM call
        # still references the object the call was made on, $Error holds 256 of
        # them for the life of the session, and a rooted RCW is one the two
        # collects below cannot reclaim - so Excel stays alive with its
        # reference count above zero and Quit leaves the process running.
        $Error.Clear()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

        # 90 SECONDS, EXPLICITLY, AND IT IS NOT A FIX. The accepted default is
        # 25. Tearing down a VBE that has compiled 28 modules is the slowest
        # shutdown in this repo, and a 25-second bound cannot tell "still
        # referenced" from "still closing". Widening it removes the timing
        # explanation; the release counts above decide the other one.
        if ($null -ne $excelIdentity) {
            $naturalExit = Wait-ExcelExit -Identity $excelIdentity -TimeoutSeconds 90
        }
        $rel.NaturalExit = $naturalExit
        Write-W1Line ''
        Write-W1Line 'EXCEL SHUTDOWN'
        Write-W1Line '--------------'
        if ($naturalExit) {
            Write-W1Line ('EXCEL SHUTDOWN: the owned process (PID ' +
                          [string]$excelIdentity.ProcessId + ') exited naturally.')
        } else {
            # THE SAFETY NET, WHICH IS NEVER A PASS. It runs so a failed run does
            # not leave a process behind, and the verdict below refuses the run
            # because it was needed at all.
            $emergencyRequired = $true
            $rel.EmergencyRequired = $true
            $cleaned = Invoke-EmergencyExcelCleanup -Identity $excelIdentity -Label 'W1'
            Write-W1Line ('EXCEL SHUTDOWN: emergency cleanup was required (' + [string]$cleaned + ')')
        }
        # EVERY RELEASE OF THE RUN, WITH ITS COUNT. This is the evidence the
        # first run could not produce: if a reference was retained, the object
        # that retained it is named here with a number beside it.
        Write-W1Line (Format-ReleaseLedger $rel)
        foreach ($residual in @($script:W1Residual)) {
            Write-W1Line ('      OUTSTANDING: ' + [string]$residual)
        }
    }
}

# ===========================================================================
# THE VERDICT
# ===========================================================================
$null = Add-W1Check 'the owned Excel process exited naturally' $naturalExit
$null = Add-W1Check 'no emergency cleanup was required' (-not $emergencyRequired)
$null = Add-W1Check 'every COM object acquired in this run was released' `
    ([int]$rel.Attempted -eq [int]$comAcquired) `
    ([string]$comAcquired + ' acquired, ' + [string]$rel.Attempted + ' released')
$null = Add-W1Check 'every COM release succeeded' ($rel.Failed.Count -eq 0) `
    ($rel.Failed -join ', ')
# THE CHECK THE FAILED RUN COULD NOT MAKE. ReleaseComObject returns what is
# left on that RCW; anything above zero IS the retained reference, named.
$null = Add-W1Check 'every COM release left 0 outstanding references' `
    (@($script:W1Residual).Count -eq 0) ((@($script:W1Residual)) -join '; ')

$failed = @($script:W1Checks | Where-Object { -not $_.Ok })
Write-W1Line ''
Write-W1Line 'VERDICT'
Write-W1Line '-------'
Write-W1Line ('checks                 : ' + [string]$script:W1Checks.Count + ' checked, ' +
              [string]$failed.Count + ' failed')
foreach ($failure in $failed) {
    $suffix = ''
    if (-not [string]::IsNullOrWhiteSpace([string]$failure.Detail)) {
        $suffix = ' -- ' + [string]$failure.Detail
    }
    Write-W1Line ('  FAILED ' + [string]$failure.Label + $suffix)
}
$ok = (($failed.Count -eq 0) -and [string]::IsNullOrWhiteSpace($fatal) -and
       ($script:W1Checks.Count -gt 0))
Write-W1Line ''
if ($ok) {
    Write-W1Line 'W1: PASS'
} else {
    Write-W1Line 'W1: FAIL'
    Write-W1Line ''
    Write-W1Line 'STOP AND REVIEW. Do not run any later scenario until this is understood.'
}
Write-W1Line ''
Write-W1Line ('report                 : ' + $script:W1Path)
Write-Host ''
Write-Host ('The report is at ' + $script:W1Path) -ForegroundColor Cyan
Write-Host 'The working copy is left in place so the report survives; delete it when done.'
if ($ok) { exit 0 } else { exit 1 }
