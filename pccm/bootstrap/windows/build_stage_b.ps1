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

# SUB-OPERATIONS, BECAUSE 'saveas.xlsm' WAS TOO COARSE AND A WINDOWS RUN PROVED IT.
# Run 5 failed in the PRE-SAVE FileFormat observation and reported itself as
# operation=saveas.xlsm, which reads as though the save had been attempted. It had
# not. A failed pre- or post-condition READ and a failed SaveAs are different
# findings, and the label has to say which.
$script:StageBBuildSubOps = @(
    'open.ready.fullname'
    'open.ready.fileformat'
    'saveas.presave.fullname'
    'saveas.presave.fileformat'
    'saveas.presave.target'
    'saveas.call'
    'saveas.post.fullname'
    'saveas.post.fileformat'
    'saveas.post.target'
)
$script:StageBBuildSubOp = ''

function Set-StageBBuildOp {
    param([string]$Operation)
    if ($script:StageBBuildOps -notcontains $Operation) {
        throw ('Set-StageBBuildOp: ' + $Operation + ' is not in the closed Stage-B build ' +
               'operation vocabulary (' + ($script:StageBBuildOps -join ', ') + ').')
    }
    $script:StageBBuildOp = $Operation
    # A NEW TOP-LEVEL OPERATION CLEARS THE SUB-OPERATION. A stale one would name a
    # step that finished for a failure somewhere else entirely.
    $script:StageBBuildSubOp = ''
}

function Get-StageBBuildOp { return [string]$script:StageBBuildOp }

function Set-StageBBuildStep {
    param([string]$Step)
    if ($script:StageBBuildSubOps -notcontains $Step) {
        throw ('Set-StageBBuildStep: ' + $Step + ' is not in the closed Stage-B build ' +
               'sub-operation vocabulary (' + ($script:StageBBuildSubOps -join ', ') + ').')
    }
    $script:StageBBuildSubOp = $Step
}

function Get-StageBBuildStep { return [string]$script:StageBBuildSubOp }

# THE LABEL A DIAGNOSTIC ACTUALLY PRINTS: the sub-operation when one is in flight,
# the top-level operation otherwise. 'saveas.xlsm' stays as the region's name; it
# is no longer what a failed read reports itself as.
function Get-StageBBuildLabel {
    if ($script:StageBBuildSubOp -ne '') { return [string]$script:StageBBuildSubOp }
    return [string]$script:StageBBuildOp
}

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

# ===========================================================================
# THE READ FORM IS THE ONE WINDOWS HAS ALREADY EXECUTED
# ===========================================================================
# RUN 5 KILLED A WRAPPER THAT HAD NEVER RUN. `Invoke-StageBBuildRead` took the COM
# object and the member name and forwarded both to the accepted helper. On Windows
# it produced:
#
#     operation=saveas.xlsm; Invoke-StageBBuildRead: the Stage-A workbook
#     FileFormat before SaveAs answered with nothing at saveas.xlsm.
#
# with COMREJECT|build|none|attempts=0|waited=0 beside it. That telemetry line is
# what makes this diagnosable: attempts=0 means the retry loop answered on its
# FIRST try and no read was ever reissued, so the reissue path, the rejection
# ledger and the backoff are all excluded - the single `$value = $Target.$Member`
# in the accepted helper returned $null without raising.
#
# WHAT IS RULED OUT, FROM SOURCE. The wrapper could not have polluted the return:
# its only emission was `return $record`, and both side effects were suppressed
# with `$null =`. It could not have consumed the value: the telemetry branch was
# gated on Attempts -gt 1, which attempts=0 proves did not run. Type coercion is
# excluded because the [int] cast sat OUTSIDE the wrapper and was never reached,
# and the null classification is the REPORT rather than the cause - $record.Value
# was genuinely $null, because $value has exactly one assignment in the helper.
#
# WHAT IS NOT CLAIMED. Why dynamic member access answered with nothing through
# that extra hop is NOT established, and nothing here theorises about it. What IS
# established is which form has worked: the reopen verification reads FileFormat,
# Worksheets, VBProject, VBComponents, CodeName, Count, Item, Shapes, OnAction and
# Name through `Invoke-ComRetryRead` called DIRECTLY, and has done so on Windows
# across every accepted run. The wrapper had never once returned a value there.
#
# SO THE UNPROVEN LAYER IS REMOVED RATHER THAN REPAIRED. Every build read now uses
# the proven form verbatim. The operation label is set before the call and the
# reissue telemetry is recorded AFTER it, both out of band - neither is in the
# expression that produces the value, so neither can pollute or consume it.

# OUT OF BAND, AND A STATEMENT RATHER THAN AN EXPRESSION. It records and returns
# nothing, so it cannot appear in a value's own expression even by accident.
function Add-StageBReadRejection {
    param([string]$Operation, $Record)
    if ($null -eq $Record) { return }
    if ([int]$Record.Attempts -le 1) { return }
    $null = $script:StageBBuildRejections.Add(
        ('COMREJECT|build|' + $Operation + '|attempts=' + [string]$Record.Attempts +
         '|waited=' + [string]$Record.WaitedMs + '|' + [string]$Record.Rejections +
         '|answered'))
}

# TYPE VALIDATION AT THE CONSUMER, NOT TRUTHINESS. `if (-not $value)` would reject
# a legitimate 0, and for a workbook property 0 is a real value elsewhere in the
# object model. These take the ALREADY-EXTRACTED primitive - never a COM object -
# so nothing about a COM member lookup is repeated here.
function Get-StageBScalarInt {
    param($Value, [string]$What)
    if ($null -eq $Value) {
        throw ($What + ' answered with nothing. The read was not refused and it was ' +
               'not answered.')
    }
    if ($Value -is [System.Array]) {
        throw ($What + ' answered with ' + [string]@($Value).Count + ' values where one ' +
               'scalar was required.')
    }
    $text = ([string]$Value).Trim()
    if ($text -notmatch '^-?[0-9]+$') {
        throw ($What + ' answered ' + $text + ', which is not an integer.')
    }
    return [int]$text
}

function Get-StageBNonEmptyString {
    param($Value, [string]$What)
    if ($null -eq $Value) {
        throw ($What + ' answered with nothing. The read was not refused and it was ' +
               'not answered.')
    }
    if ($Value -is [System.Array]) {
        throw ($What + ' answered with ' + [string]@($Value).Count + ' values where one ' +
               'string was required.')
    }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        throw ($What + ' answered an empty string.')
    }
    return $text
}

# ===========================================================================
# THE OPENED WORKBOOK DOES NOT ALWAYS ANSWER FOR ITSELF YET
# ===========================================================================
# TWO WINDOWS RUNS, AND THE FAILURE MOVED WITH THE POSITION RATHER THAN THE MEMBER.
# At cc9cf8d the first read after Workbooks.Open was FileFormat and it answered with
# nothing. The read wrapper was removed, so at c66e752 the first read after
# Workbooks.Open became FullName - and IT answered with nothing, in the same place,
# with COMREJECT|build|none|attempts=0|waited=0 beside it both times. In that same
# c66e752 run the FIRST Excel session read FullName and FileFormat successfully and
# completed its SaveAs.
#
# So it is not the member and it is not the wrapper. What Windows shows is that
# Workbooks.Open can return a Workbook RCW in the second isolated session BEFORE its
# own read-only properties reliably answer. That is WHERE the gap is observed.
#
# WHY EXCEL BEHAVES THIS WAY IS NOT ESTABLISHED, and nothing here asserts one: not a
# message-filter race, not modal state, not OneDrive, not a file lock, not lifecycle
# overlap, not a marshaling cause, not a bug. The gate is a bounded OBSERVATION, and
# it waits only because an observation did not answer.
#
# AND IT IS NOT A DELAY. Nothing sleeps on the way past a workbook that answers: the
# first attempt returns and the run continues. A sleep happens only after a
# readiness observation came back with no answer, or was refused by the message
# filter and the accepted helper could not resolve it within its own bounds.

# READ-ONLY, AND STRICTER THAN THE HELPER IT CALLS. Invoke-ComRetryRead returning a
# record whose Value is $null is a SUCCESS to the helper - it read the member and
# that is what came back. To this gate it means NOT READY YET, which is the whole
# difference and the reason the gate exists rather than a wider retry.
function Wait-StageBWorkbookReady {
    param($Workbook, [string]$ExpectedPath,
          [int]$MaxAttempts   = 12,
          [int]$FirstDelayMs  = 250,
          [int]$MaxDelayMs    = 2000,
          [int]$TotalBudgetMs = 15000)
    if ($null -eq $Workbook) { throw 'Wait-StageBWorkbookReady: no workbook.' }
    if ([string]::IsNullOrWhiteSpace($ExpectedPath)) {
        throw 'Wait-StageBWorkbookReady: no expected path to recognise the workbook by.'
    }
    if ($MaxAttempts -lt 1) { throw 'Wait-StageBWorkbookReady: MaxAttempts must be at least 1.' }
    if ($TotalBudgetMs -lt 0) { throw 'Wait-StageBWorkbookReady: TotalBudgetMs may not be negative.' }

    $attempt     = 0
    $waitedMs    = 0
    $delay       = $FirstDelayMs
    $nameState   = 'unresolved'
    $formatState = 'unresolved'
    $fullName    = ''
    $format      = 0
    $ready       = $false

    # THE BOUND IS IN THE HEAD, as it is in the read helper and the SaveAs loop.
    while ($attempt -lt $MaxAttempts) {
        $attempt     = $attempt + 1
        $nameState   = 'unresolved'
        $formatState = 'unresolved'

        # --- A. which file is this workbook? --------------------------------
        Set-StageBBuildStep 'open.ready.fullname'
        $nameValue = $null
        try {
            $nameRead = Invoke-ComRetryRead -Target $Workbook -Member 'FullName' `
                            -Description 'the opened workbook FullName'
            Add-StageBReadRejection -Operation (Get-StageBBuildLabel) -Record $nameRead
            $nameValue = $nameRead.Value
        } catch {
            # A REFUSAL THE HELPER COULD NOT RESOLVE WITHIN ITS OWN BOUNDS IS 'NOT
            # READY YET'. Anything else Excel accepted and failed, and that aborts.
            if ([string]::IsNullOrWhiteSpace((Get-ComRejectionName $_))) { throw }
            $nameValue = $null
        }
        if ($null -eq $nameValue) {
            $nameState = 'no-answer'
        } elseif ($nameValue -is [System.Array]) {
            throw ('READY: the opened workbook FullName answered with ' +
                   [string]@($nameValue).Count + ' values where one string was required.')
        } elseif ([string]::IsNullOrWhiteSpace([string]$nameValue)) {
            $nameState = 'no-answer'
        } else {
            # A REAL PATH THAT IS THE WRONG PATH IS AN IDENTITY FAILURE, NOT A DELAY.
            # Waiting cannot change which workbook this is, so polling would only
            # spend the budget on a question already answered.
            if ((Get-StageBComparablePath ([string]$nameValue)) -ne
                (Get-StageBComparablePath $ExpectedPath)) {
                throw ('READY: the opened workbook is bound to ' + [string]$nameValue +
                       ' and not to ' + $ExpectedPath + '. Waiting cannot change which ' +
                       'workbook this is.')
            }
            $fullName  = [string]$nameValue
            $nameState = 'ok'
        }

        # --- B. what format is it in? ---------------------------------------
        # OBSERVED, NEVER ASSUMED. This value becomes the original source format the
        # SaveAs NOT-EXECUTED verdict is measured against, so a literal here would
        # be this script restating a contract it is supposed to read.
        Set-StageBBuildStep 'open.ready.fileformat'
        $formatValue = $null
        try {
            $formatRead = Invoke-ComRetryRead -Target $Workbook -Member 'FileFormat' `
                              -Description 'the opened workbook FileFormat'
            Add-StageBReadRejection -Operation (Get-StageBBuildLabel) -Record $formatRead
            $formatValue = $formatRead.Value
        } catch {
            if ([string]::IsNullOrWhiteSpace((Get-ComRejectionName $_))) { throw }
            $formatValue = $null
        }
        if ($null -eq $formatValue) {
            $formatState = 'no-answer'
        } elseif ($formatValue -is [System.Array]) {
            throw ('READY: the opened workbook FileFormat answered with ' +
                   [string]@($formatValue).Count + ' values where one scalar was required.')
        } elseif ([string]::IsNullOrWhiteSpace([string]$formatValue)) {
            $formatState = 'no-answer'
        } elseif ((([string]$formatValue).Trim()) -notmatch '^-?[0-9]+$') {
            # A REAL BUT UNUSABLE ANSWER IS AN ERROR, not a readiness delay.
            throw ('READY: the opened workbook FileFormat answered ' +
                   ([string]$formatValue).Trim() + ', which is not an integer.')
        } else {
            $format      = [int](([string]$formatValue).Trim())
            $formatState = 'ok'
        }

        if (($nameState -eq 'ok') -and ($formatState -eq 'ok')) { $ready = $true; break }

        # NO SLEEP ON THE WAY OUT. The break above happens first, so a workbook that
        # answered is never waited on.
        if ($attempt -ge $MaxAttempts) { break }
        if (($waitedMs + $delay) -gt $TotalBudgetMs) { break }
        Start-Sleep -Milliseconds $delay
        $waitedMs = $waitedMs + $delay
        $delay = [Math]::Min(($delay + $FirstDelayMs), $MaxDelayMs)
    }

    if (-not $ready) {
        Add-Note ('READY|open|exhausted|attempts=' + [string]$attempt + '|waited=' +
                  [string]$waitedMs + '|fullname=' + $nameState + '|fileformat=' + $formatState)
        throw ('READY: the opened workbook did not answer its own properties after ' +
               [string]$attempt + ' attempt(s) and ' + [string]$waitedMs + ' ms ' +
               '(fullname=' + $nameState + ', fileformat=' + $formatState + '). ' +
               'Stage-B stops here; SaveAs was not attempted.')
    }
    # ONE LINE, whatever happened. A run that waited nine seconds and one that did
    # not must not look alike.
    Add-Note ('READY|open|attempt=' + [string]$attempt + '|fullname=True|fileformat=' +
              [string]$format + '|waited=' + [string]$waitedMs)
    return [pscustomobject]@{
        FullName   = $fullName
        FileFormat = $format
        Attempts   = $attempt
        WaitedMs   = $waitedMs
    }
}

# ===========================================================================
# SaveAs: A NON-IDEMPOTENT CALL SETTLED BY ITS POSTCONDITION, NOT BY A CONTRACT
# ===========================================================================
# WINDOWS NAMED IT. The labelled build reported, repeatably:
#
#     [FAIL] Stage-B build
#            operation=saveas.xlsm; System.Runtime.InteropServices.COMException:
#            Call was rejected by callee. 0x80010001 RPC_E_CALL_REJECTED
#     COMREJECT|build|none|attempts=0|waited=0
#
# The last line matters as much as the first: NO READ was reissued, so the read
# retry had nothing to do with this. The rejection is in Workbook.SaveAs, before
# the worksheets, the VBProject, the modules, the buttons, the protection and the
# final save.
#
# AND THE OLD WORDING WILL NOT DO. 'the call was refused before it ran' is what
# the message-filter contract says, and for a property get that is enough. SaveAs
# WRITES A FILE AND REBINDS THE WORKBOOK. Reissuing one on the strength of a
# contract is precisely the guess this project refuses, so the contract is not
# used as the licence here at all: the licence is OBSERVED STATE.
#
# WHY THE OBSERVATION IS CONCLUSIVE HERE, AND WOULD NOT BE ANYWHERE. The build
# DELETES the target immediately before the call, so at the moment SaveAs is
# attempted the target provably does not exist and the workbook is provably bound
# to the Stage-A path in the Stage-A format. Three independent facts therefore
# separate the two outcomes, and they cannot be half-true together:
#
#   the workbook's FullName   - bound to the target, or still to the source
#   the workbook's FileFormat - the target's 52, or the source's original
#   the target file on disk   - present, or absent
#
# Anything other than all-three-target or all-three-source is AMBIGUOUS, and an
# ambiguous save is not retried and not cleaned up: it aborts and says so, with
# the evidence intact.
#
# WHY THIS IS NOT A GENERAL RETRY. It is one function, for one call, gated on one
# observation. Nothing else in this bootstrap gains a mutation retry, and
# Invoke-ComRetryRead is untouched - it still cannot express a write.

# Excel answers with an absolute path in the host's own separator. The build's own
# paths come through Join-Path from a caller-supplied BuildDir, so both sides are
# normalised before they are compared; a case or separator difference is not a
# rebind, and reading one as a rebind would manufacture an ambiguous state.
function Get-StageBComparablePath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $text = $Path
    try {
        $text = $text.Replace('/', [System.IO.Path]::DirectorySeparatorChar)
        $text = [System.IO.Path]::GetFullPath($text)
    } catch { $text = $Path }
    return $text.TrimEnd([System.IO.Path]::DirectorySeparatorChar).ToLowerInvariant()
}

# READ-ONLY, AND IT DECIDES NOTHING BY ITSELF. It reports what is observable and
# which of the three states that adds up to. Both COM reads use the form the reopen
# verification has executed on Windows - the accepted helper, called directly - so
# an inspection that is itself refused is reissued rather than mistaken for a
# finding, and each carries its own sub-operation label.
function Get-StageBSaveAsPostcondition {
    param($Workbook, [string]$SourcePath, [string]$TargetPath,
          [int]$TargetFormat, [int]$SourceFormat)
    $fullName  = ''
    $format    = 0
    $readError = ''
    try {
        Set-StageBBuildStep 'saveas.post.fullname'
        $nameRead = Invoke-ComRetryRead -Target $Workbook -Member 'FullName' `
                        -Description 'the workbook FullName after SaveAs'
        Add-StageBReadRejection -Operation (Get-StageBBuildLabel) -Record $nameRead
        $fullName = Get-StageBNonEmptyString -Value $nameRead.Value `
                        -What 'the workbook FullName after SaveAs'

        Set-StageBBuildStep 'saveas.post.fileformat'
        $formatRead = Invoke-ComRetryRead -Target $Workbook -Member 'FileFormat' `
                          -Description 'the workbook FileFormat after SaveAs'
        Add-StageBReadRejection -Operation (Get-StageBBuildLabel) -Record $formatRead
        $format = Get-StageBScalarInt -Value $formatRead.Value `
                      -What 'the workbook FileFormat after SaveAs'
    } catch {
        # A READ THAT COULD NOT BE ANSWERED IS NOT EVIDENCE OF ANYTHING. It makes
        # the state ambiguous, which is the state that never retries.
        $readError = Format-Err $_
    }
    Set-StageBBuildStep 'saveas.post.target'
    $seen          = Get-StageBComparablePath $fullName
    $boundToTarget = ($seen -ne '') -and ($seen -eq (Get-StageBComparablePath $TargetPath))
    $boundToSource = ($seen -ne '') -and ($seen -eq (Get-StageBComparablePath $SourcePath))
    $targetExists  = [bool](Test-Path -LiteralPath $TargetPath)
    $sourceExists  = [bool](Test-Path -LiteralPath $SourcePath)

    # THREE FACTS, ALL OF THEM, FOR EITHER ANSWER. Any single one of them on its
    # own is exactly the partial evidence that must NOT settle this: a target that
    # exists while the workbook is still bound to the source is a half-written
    # file, not a completed save.
    $state = 'ambiguous'
    if ($readError -ne '') {
        $state = 'ambiguous'
    } elseif ($boundToTarget -and ($format -eq $TargetFormat) -and $targetExists) {
        $state = 'completed'
    } elseif ($boundToSource -and ($format -eq $SourceFormat) -and (-not $targetExists)) {
        $state = 'not-executed'
    }
    return [pscustomobject]@{
        State         = $state
        FullName      = $fullName
        FileFormat    = $format
        BoundToTarget = $boundToTarget
        BoundToSource = $boundToSource
        TargetExists  = $targetExists
        SourceExists  = $sourceExists
        ReadError     = $readError
        Detail        = ('boundToTarget=' + [string]$boundToTarget +
                         '|boundToSource=' + [string]$boundToSource +
                         '|format=' + [string]$format +
                         '|targetExists=' + [string]$targetExists +
                         '|sourceExists=' + [string]$sourceExists +
                         '|readError=' + $(if ($readError -eq '') { 'none' } else { $readError }))
    }
}

# THE ONLY MUTATION IN THIS BOOTSTRAP WITH A RECOVERY PATH, and the bounds are the
# accepted envelope's rather than new numbers. Every reissue is gated on an
# observation, so the sequence is: attempt, and on a REFUSED call inspect before
# deciding - never inspect after deciding, and never decide without inspecting.
function Invoke-StageBSaveAs {
    param($Workbook, [string]$SourcePath, [string]$TargetPath,
          [int]$TargetFormat, [int]$SourceFormat,
          [string]$SourceFullName,
          [int]$MaxAttempts   = 12,
          [int]$FirstDelayMs  = 250,
          [int]$MaxDelayMs    = 2000,
          [int]$TotalBudgetMs = 15000)
    if ($null -eq $Workbook) { throw 'Invoke-StageBSaveAs: no workbook.' }
    if ($MaxAttempts -lt 1) { throw 'Invoke-StageBSaveAs: MaxAttempts must be at least 1.' }
    if ($TotalBudgetMs -lt 0) { throw 'Invoke-StageBSaveAs: TotalBudgetMs may not be negative.' }

    $attempt  = 0
    $waitedMs = 0
    $delay    = $FirstDelayMs

    # THE BOUND IS IN THE HEAD, as it is in the read helper: a loop whose only
    # exits are a break and a throw is bounded in fact and unbounded to a reader.
    while ($attempt -lt $MaxAttempts) {
        $attempt = $attempt + 1
        $refused = ''
        try {
            # THE CALL'S OWN LABEL. A failure from here IS the save; a failure from
            # an observation is not, and run 5 could not tell the two apart.
            Set-StageBBuildStep 'saveas.call'
            $Workbook.SaveAs($TargetPath, $TargetFormat)
        } catch {
            $refused = Get-ComRejectionName $_
            $hres = Get-StageBComHResult $_
            if ($hres -eq '') { $hres = 'none' }
            if ([string]::IsNullOrWhiteSpace($refused)) {
                # EXCEL ACCEPTED THIS ONE AND IT FAILED. That describes something
                # which actually happened, and nothing here may reissue it.
                Add-Note ('SAVEAS|attempt=' + [string]$attempt + '|error|' + $hres)
                throw
            }
            Add-Note ('SAVEAS|attempt=' + [string]$attempt + '|rejected|' + $hres)
            # INSPECT BEFORE ANY SECOND CALL. This is the whole settlement.
            $state = Get-StageBSaveAsPostcondition -Workbook $Workbook `
                -SourcePath $SourcePath -TargetPath $TargetPath `
                -TargetFormat $TargetFormat -SourceFormat $SourceFormat
            Add-Note ('SAVEAS|postcondition|' + $state.State + '|' + $state.Detail)
            if ($state.State -eq 'completed') {
                # IT HAPPENED. Reissuing it now would overwrite the workbook this
                # build already owns, so the exception is the news and the state
                # is the answer.
                Add-Note ('SAVEAS|attempt=' + [string]$attempt +
                          '|completed-despite-rejection')
                return (New-StageBSaveAsResult -State $state -Attempts $attempt -WaitedMs $waitedMs)
            }
            if ($state.State -ne 'not-executed') {
                throw ('SAVEAS AMBIGUOUS after ' + [string]$attempt + ' attempt(s): the ' +
                       'save can be proved neither to have happened nor to have been ' +
                       'skipped, so it is not reissued and nothing is cleaned up. ' +
                       $state.Detail)
            }
            if ($attempt -ge $MaxAttempts) { throw }
            if (($waitedMs + $delay) -gt $TotalBudgetMs) { throw }
            Start-Sleep -Milliseconds $delay
            $waitedMs = $waitedMs + $delay
            $delay = [Math]::Min(($delay + $FirstDelayMs), $MaxDelayMs)
            continue
        }
        # THE CALL RETURNED, WHICH IS NOT THE SAME AS THE SAVE HAVING HAPPENED. The
        # same three facts are required of a normal return as of a refused one.
        $state = Get-StageBSaveAsPostcondition -Workbook $Workbook `
            -SourcePath $SourcePath -TargetPath $TargetPath `
            -TargetFormat $TargetFormat -SourceFormat $SourceFormat
        Add-Note ('SAVEAS|attempt=' + [string]$attempt + '|success')
        Add-Note ('SAVEAS|verified|path=' + [string]$state.BoundToTarget +
                  '|format=' + [string]$state.FileFormat +
                  '|exists=' + [string]$state.TargetExists)
        if ($state.State -ne 'completed') {
            throw ('SaveAs returned without error but its postconditions do not prove ' +
                   'the save: ' + $state.Detail)
        }
        return (New-StageBSaveAsResult -State $state -Attempts $attempt -WaitedMs $waitedMs)
    }
    # UNREACHABLE WHILE EVERY PATH ABOVE RETURNS OR THROWS, and it still throws
    # rather than letting a changed head bound turn 'never saved' into a save.
    throw ('Invoke-StageBSaveAs: the save was never completed after ' +
           [string]$attempt + ' attempt(s).')
}

function New-StageBSaveAsResult {
    param($State, [int]$Attempts, [int]$WaitedMs)
    return [pscustomobject]@{
        State      = [string]$State.State
        FullName   = [string]$State.FullName
        FileFormat = [int]$State.FileFormat
        Attempts   = $Attempts
        WaitedMs   = $WaitedMs
        Detail     = [string]$State.Detail
    }
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
    # IMMEDIATELY, ON THE OBJECT Open JUST RETURNED, and before anything at all is
    # read or written through it. Two Windows runs failed on the FIRST property read
    # after this line - different members, same position.
    $ready = Wait-StageBWorkbookReady -Workbook $wb -ExpectedPath $stageAPath
    Add-Step 'Open the Stage-A workbook' 'PASS' `
        ($stageAPath + '; ready on attempt ' + [string]$ready.Attempts + ', waited ' +
         [string]$ready.WaitedMs + ' ms, FileFormat=' + [string]$ready.FileFormat)

    # --- 3. save as .xlsm --------------------------------------------------
    if (Test-Path -LiteralPath $stageBPath) { Remove-Item -LiteralPath $stageBPath -Force }
    # THE BASELINE IS WHAT THE READINESS GATE ALREADY OBSERVED, and it is NOT read
    # again. A second read for symmetry's sake would be a second chance for the
    # workbook to answer differently, and the value the NOT-EXECUTED verdict is
    # measured against has to be the one readiness actually established.
    Set-StageBBuildOp 'saveas.xlsm'

    Set-StageBBuildStep 'saveas.presave.fullname'
    $sourceFullName = Get-StageBNonEmptyString -Value $ready.FullName `
                          -What 'the Stage-A workbook FullName observed at open'

    Set-StageBBuildStep 'saveas.presave.fileformat'
    $sourceFormat = Get-StageBScalarInt -Value $ready.FileFormat `
                        -What 'the Stage-A workbook FileFormat observed at open'

    # AND THE BASELINE MUST BE INTERNALLY CONSISTENT BEFORE ANYTHING IS SAVED. If
    # the workbook is not the Stage-A file, or the target is already there, then
    # 'bound to source and target absent' cannot mean what the classification needs
    # it to mean - so the save is not attempted at all.
    Set-StageBBuildStep 'saveas.presave.target'
    $preSeen = Get-StageBComparablePath $sourceFullName
    if ($preSeen -ne (Get-StageBComparablePath $stageAPath)) {
        throw ('SAVEAS BASELINE: the workbook is bound to ' + $sourceFullName +
               ' and not to the Stage-A path ' + $stageAPath +
               ', so NOT EXECUTED could not be recognised. The save was not attempted.')
    }
    if (Test-Path -LiteralPath $stageBPath) {
        throw ('SAVEAS BASELINE: ' + $stageBPath + ' is present before the save, so ' +
               'the target existing afterwards would prove nothing. The save was ' +
               'not attempted.')
    }
    Add-Note ('SAVEAS|baseline|fullname=ok|format=' + [string]$sourceFormat +
              '|target-absent=True')

    $saveAs = Invoke-StageBSaveAs -Workbook $wb -SourcePath $stageAPath -TargetPath $stageBPath `
                  -TargetFormat ([int]$manifest.xlsm_file_format) -SourceFormat $sourceFormat `
                  -SourceFullName $sourceFullName
    $actualFormat = [int]$saveAs.FileFormat
    if ($actualFormat -ne [int]$manifest.xlsm_file_format) {
        throw ("SaveAs produced FileFormat {0}, expected {1}." -f $actualFormat, $manifest.xlsm_file_format)
    }
    Add-Step 'Save as macro-enabled .xlsm' 'PASS' `
        ("FileFormat={0}; attempt(s)={1}; waited={2} ms; {3}" -f $actualFormat, $saveAs.Attempts, $saveAs.WaitedMs, $stageBPath)

    # --- 4. CodeNames -------------------------------------------------------
    # THREE PLAIN PROPERTY GETS, AND THE REOPEN PATH ALREADY RETRIES ALL THREE.
    # Worksheets, VBProject and VBComponents are acquisitions: they read a member
    # and move nothing. The verification block below reissues exactly these
    # members on exactly these classes of object, so opting the BUILD's copies in
    # adds no new judgement - it applies one already accepted, at the reads that
    # sit closest to the workbook opening.
    # A COM OBJECT IS TAKEN STRAIGHT OFF THE RECORD, NOT THROUGH ANOTHER FUNCTION.
    # The acquired collection is assigned from $read.Value with no further parameter
    # binding, so the value path is the proven one end to end.
    Set-StageBBuildOp 'worksheets.acquire'
    $wsRead = Invoke-ComRetryRead -Target $wb -Member 'Worksheets' `
                  -Description 'the Stage-B workbook Worksheets collection'
    Add-StageBReadRejection -Operation (Get-StageBBuildLabel) -Record $wsRead
    if ($null -eq $wsRead.Value) {
        throw 'the Stage-B workbook Worksheets collection answered with nothing.'
    }
    $worksheets = $wsRead.Value
    try {
        Set-StageBBuildOp 'vbproject.acquire'
        $vbpRead = Invoke-ComRetryRead -Target $wb -Member 'VBProject' `
                       -Description 'the Stage-B workbook VBProject'
        Add-StageBReadRejection -Operation (Get-StageBBuildLabel) -Record $vbpRead
        if ($null -eq $vbpRead.Value) {
            throw 'the Stage-B workbook VBProject answered with nothing.'
        }
        $vbproj = $vbpRead.Value
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
    Set-StageBBuildOp 'vbcomponents.acquire'
    $vbcRead = Invoke-ComRetryRead -Target $vbproj -Member 'VBComponents' `
                   -Description 'the Stage-B VBComponents collection'
    Add-StageBReadRejection -Operation (Get-StageBBuildLabel) -Record $vbcRead
    if ($null -eq $vbcRead.Value) {
        throw 'the Stage-B VBComponents collection answered with nothing.'
    }
    $vbcomps = $vbcRead.Value

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
    # THE PRECISE LABEL. Run 5 reported operation=saveas.xlsm for a failure in the
    # pre-save observation, which reads as though the save had been attempted.
    $failedOp = Get-StageBBuildLabel
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
