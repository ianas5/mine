<#
.SYNOPSIS
    PCCM Phase-5 Gate-B scenarios. Dot-sourced INTO phase4_functional_test.ps1.

.DESCRIPTION
    THIS IS NOT A SECOND HARNESS. It is dot-sourced into the accepted Phase-4
    functional test and runs inside that script's ONE COM lifecycle, against the
    ONE Excel instance it owns, the ONE workbook it opened and the ONE Stage-B
    bootstrap it ran. It reuses Add-Result, New-Checklist, Add-Check,
    Test-ChecklistOk, Format-Checklist, Format-Err, Get-TableBody, Set-TableCell,
    Get-NamedValue, Set-NamedValue and the release ledger unchanged. There is no
    competing bootstrap path, no second Excel process, no second reporting
    surface and no second shutdown.

    The Phase-4 matrix is untouched and remains mandatory. Gate-B acceptance
    requires Phase 4 at 35/35, 0 FAIL, 0 SKIP, BEFORE any Phase-5 result counts;
    Invoke-Phase5GateBScenarios refuses to run its own scenarios if the Phase-4
    matrix did not reach that state, and reports that refusal as a FAIL rather
    than as a skip nobody reads.

    EXPECTED VALUES COME FROM build/phase5_cases.json. Not one analytical number,
    canonical string, digest or remainder is written into this file. Where a
    comparison needs an address rather than a value, it comes from
    build/stage_b_manifest.json or build/phase5_gate_b_inspection.json, both of
    which the Stage-A build emits from the accepted contracts.

    NOTHING HERE HAS BEEN EXECUTED. No Windows run has been made and no Excel COM
    session has been started for Phase 5. This file is source under review.
#>

Set-StrictMode -Version 2.0

# ===========================================================================
# THE COVERAGE LEDGER
# ===========================================================================
# Every plan-case ID emitted into phase5_cases.json maps to at least one Windows
# scenario. The map is DATA, checked against the fixture corpus in a preflight
# that runs before Excel is started, so a case that was added to the corpus and
# never wired into the harness stops the run instead of quietly disappearing.
#
# A case may map to several scenarios; several cases may share one scenario and
# one workbook fixture. What may never happen is a case with no mapping, or a
# mapping that names a scenario the harness does not define.
function Get-Phase5CoverageLedger {
    $ledger = New-Object System.Collections.Specialized.OrderedDictionary
    # --- analytical fixtures, driven through PCCM_Calculate -----------------
    foreach ($id in 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 19, 21, 22, 25, 31) {
        $ledger.Add([string]$id, @('P5-AN'))
    }
    # 30 is analytical AND the cancellation-heavy reconciliation vector.
    $ledger.Add('30', @('P5-AN', 'P5-ID'))
    # --- prerequisite refusals ---------------------------------------------
    foreach ($id in 14, 15, 16, 17, 18, 20, 23, 24, 29) {
        $ledger.Add([string]$id, @('P5-RF'))
    }
    # --- direct real-VBA vectors, through the transient diagnostic module ---
    $ledger.Add('26', @('P5-D1', 'P5-D4', 'P5-D5'))
    $ledger.Add('27', @('P5-D6'))
    $ledger.Add('28', @('P5-D7'))
    $ledger.Add('35', @('P5-D2'))
    $ledger.Add('36', @('P5-D3'))
    # --- runtime-only: workbook state across attempts -----------------------
    $ledger.Add('32', @('P5-RC', 'P5-S5'))
    $ledger.Add('33', @('P5-FA'))
    $ledger.Add('34', @('P5-S3', 'P5-S4', 'P5-KP'))
    $ledger.Add('37', @('P5-FC'))
    return $ledger
}

# Every scenario ID the harness defines. The preflight rejects a ledger entry
# naming anything outside this set, so a typo cannot silently drop a case.
function Get-Phase5ScenarioIds {
    return @(
        'P5-CMP', 'P5-FX', 'P5-FIX', 'P5-M',
        'P5-D0', 'P5-D1', 'P5-D2', 'P5-D3', 'P5-D4', 'P5-D5', 'P5-D6', 'P5-D7',
        'P5-DC', 'P5-D8',
        'P5-AN', 'P5-RF', 'P5-PQ', 'P5-PN', 'P5-AR', 'P5-ID',
        'P5-S1', 'P5-S2', 'P5-S3', 'P5-S4', 'P5-S5', 'P5-S6',
        'P5-ST', 'P5-NS', 'P5-KP', 'P5-RC',
        'P5-FA', 'P5-FC',
        'P5-AX', 'P5-EV'
    )
}

# The two locked Phase-5 failpoint stage names. They are declared in the accepted
# production module modCalcReport.bas; tests/test_phase5_gate_b_harness_source.py
# pins these two strings against that module, so the production source stays the
# authority and this is a checked copy rather than a second declaration.
function Get-Phase5FailpointNames {
    return [pscustomobject]@{
        AnalyticalWrite = 'Phase5AnalyticalWrite'
        SuccessCommit   = 'Phase5SuccessCommit'
    }
}

# The Phase-4 matrix that must be intact before a Phase-5 result means anything.
# The timeline chain D..J is reported as the ten sequential steps D-J.1 .. D-J.10,
# so the matrix is 35 results, not 35 letters.
$script:Phase4RequiredScenarioIds = @(
    'PRE0', 'PRE', 'A', 'A1', 'A2', 'B', 'B2', 'C', 'D0',
    'D-J.1', 'D-J.2', 'D-J.3', 'D-J.4', 'D-J.5',
    'D-J.6', 'D-J.7', 'D-J.8', 'D-J.9', 'D-J.10',
    'K', 'K2', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W',
    'Y', 'Z'
)

# Two of those 35 are not Phase-4 BEHAVIOUR cases, and they are the whole of the
# Run-1 sequencing defect:
#
#   Z  asserts the owned Excel process exited naturally. It is recorded after
#      Workbook.Close, Application.Quit and the COM release ledger have run.
#   Y  asserts every transient COM object released cleanly across the WHOLE run,
#      Phase-5's own transients included. It is recorded last of all, from
#      Get-TransientFailures.
#
# Both are POST-SESSION lifecycle assertions. Phase 5 runs inside the live
# automation session, so neither can exist while Invoke-Phase5GateBScenarios is
# executing - and evaluating Y early would attest to Phase-4's transients only,
# which is weaker evidence, not stronger.
#
# They stay in the 35-case matrix and stay mandatory. What moved is WHERE the
# 35/35 demand is made: P5-FIN, after Y and Z, gates final acceptance.
$script:Phase4FinalizationScenarioIds = @('Y', 'Z')

function Get-Phase4RequiredScenarioIds { return $script:Phase4RequiredScenarioIds }

function Get-Phase4FinalizationScenarioIds { return $script:Phase4FinalizationScenarioIds }

# DERIVED, never a second hand-maintained list. A case added to the matrix is a
# prerequisite case automatically; the only way to defer one is to name it above,
# and P5-P4 proves the two sets still partition the matrix exactly.
function Get-Phase4PrerequisiteScenarioIds {
    $deferred = $script:Phase4FinalizationScenarioIds
    return @($script:Phase4RequiredScenarioIds | Where-Object { $deferred -notcontains $_ })
}

# ===========================================================================
# P5-FIN. Final Phase-4 completeness, AFTER the post-session lifecycle cases
# ===========================================================================
# The other half of the Run-1 correction, and the reason the 35/35 requirement
# did not weaken when P5-P4 narrowed to what can exist at its point in the
# lifecycle. This runs last, after Y and Z have been recorded, and it demands:
#
#   all 35 matrix cases reported a result
#   0 FAIL across all 35
#   0 SKIP across all 35
#   35/35 PASS
#   and each deferred case, BY NAME, ran exactly once and PASSED
#
# The named check is what makes final acceptance impossible without Y and Z: a
# bare 35/35 count would be satisfiable by a matrix that lost Z and counted
# something else twice. Because the driver's exit code is driven by the FAIL
# count, a FAIL here fails the whole run - including a Phase-4 SKIP, which the
# summary alone would have printed and then exited 0 on.
function Add-Phase4FinalCompletenessResult {
    param($Results)

    $required = Get-Phase4RequiredScenarioIds
    $deferred = Get-Phase4FinalizationScenarioIds
    $list = New-Checklist
    $seen = @($Results | ForEach-Object { $_.Id })

    $missing = @()
    foreach ($id in $required) { if ($seen -notcontains $id) { $missing += $id } }
    $null = Add-Check $list 'all 35 Phase-4 scenarios reported a result' ($missing.Count -eq 0) `
        ("missing: " + ($missing -join ', '))

    $phase4 = @($Results | Where-Object { $required -contains $_.Id })
    $failed = @($phase4 | Where-Object { $_.Status -eq 'FAIL' })
    $skipped = @($phase4 | Where-Object { $_.Status -eq 'SKIP' })
    $null = Add-Check $list 'the final Phase-4 matrix has 0 FAIL' ($failed.Count -eq 0) `
        (($failed | ForEach-Object { $_.Id }) -join ', ')
    $null = Add-Check $list 'the final Phase-4 matrix has 0 SKIP' ($skipped.Count -eq 0) `
        (($skipped | ForEach-Object { $_.Id }) -join ', ')
    $passed = @($phase4 | Where-Object { $_.Status -eq 'PASS' })
    $null = Add-Check $list 'the final Phase-4 matrix is 35/35 PASS' ($passed.Count -eq 35) `
        ("passed " + $passed.Count + " of 35")

    # BY NAME, one at a time, and exactly one record each.
    foreach ($id in $deferred) {
        $record = @($Results | Where-Object { $_.Id -eq $id })
        $null = Add-Check $list `
            ('the deferred lifecycle case ' + $id + ' ran exactly once and PASSED') `
            (($record.Count -eq 1) -and ([string]$record[0].Status -eq 'PASS')) `
            ("recorded " + $record.Count + " result(s): " +
             (($record | ForEach-Object { [string]$_.Status }) -join ', '))
    }

    Add-Phase5Result 'P5-FIN' `
        'Final Phase-4 completeness: 35/35 PASS, 0 FAIL, 0 SKIP, deferred lifecycle cases included' `
        $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
}

# ===========================================================================
# P5-PRE. Coverage preflight, BEFORE Excel is started
# ===========================================================================
# Pure PowerShell, no COM. It reads the emitted corpus and the ledger above and
# refuses the run if the two disagree. A Gate-B run that started, spent twenty
# minutes in Excel and then reported 36 of 37 cases would be worse than one that
# never started: the missing case would be a line in a summary nobody diffs.
function Invoke-Phase5CoveragePreflight {
    param([string]$BuildDir)

    # THE FIRST PHASE-5 ENTRY POINT, so the one-result ledger starts here. It is
    # deliberately NOT reset again in Invoke-Phase5GateBScenarios: that would
    # discard P5-PRE's record and let a later catch emit it a second time.
    Reset-Phase5ResultLedger

    $list = New-Checklist
    $casesPath = Join-Path $BuildDir 'phase5_cases.json'
    $inspectPath = Join-Path $BuildDir 'phase5_gate_b_inspection.json'

    $haveCases = Test-Path -LiteralPath $casesPath
    $null = Add-Check $list 'build/phase5_cases.json exists (the expected-value authority)' `
        $haveCases $casesPath
    $haveInspect = Test-Path -LiteralPath $inspectPath
    $null = Add-Check $list 'build/phase5_gate_b_inspection.json exists (the address authority)' `
        $haveInspect $inspectPath
    if (-not ($haveCases -and $haveInspect)) {
        Add-Phase5Result 'P5-PRE' 'Phase-5 coverage preflight (pure PowerShell, no Excel)' 'FAIL' `
            (Format-Checklist $list)
        return $false
    }

    $cases = Get-Content -LiteralPath $casesPath -Raw | ConvertFrom-Json
    $ledger = Get-Phase5CoverageLedger
    $known = Get-Phase5ScenarioIds

    $emitted = @()
    foreach ($case in @($cases.plan_cases)) { $emitted += [string]$case.id }
    $null = Add-Check $list 'the corpus emitted at least one plan case' ($emitted.Count -gt 0) `
        ("emitted " + $emitted.Count)

    # EVERY EMITTED ID HAS A MAPPING. Driven from the corpus, never from the
    # ledger: a case added to phase5_cases.json and forgotten here must fail.
    $unmapped = @()
    foreach ($id in $emitted) { if (-not $ledger.Contains($id)) { $unmapped += $id } }
    $null = Add-Check $list 'every emitted plan-case ID maps to a Windows scenario' `
        ($unmapped.Count -eq 0) ("unmapped: " + ($unmapped -join ', '))

    # AND NO MAPPING IS A GHOST. A ledger entry for a case the corpus no longer
    # emits is a coverage claim with nothing behind it.
    $orphan = @()
    foreach ($id in $ledger.Keys) { if ($emitted -notcontains $id) { $orphan += $id } }
    $null = Add-Check $list 'no ledger entry names a case the corpus does not emit' `
        ($orphan.Count -eq 0) ("orphaned: " + ($orphan -join ', '))

    # EVERY MAPPING POINTS AT A SCENARIO THIS HARNESS DEFINES.
    $unknown = @()
    foreach ($id in $ledger.Keys) {
        foreach ($scenario in @($ledger[$id])) {
            if ($known -notcontains $scenario) { $unknown += ($id + '->' + $scenario) }
        }
    }
    $null = Add-Check $list 'every mapping names a scenario the harness defines' `
        ($unknown.Count -eq 0) ("unknown: " + ($unknown -join ', '))

    # EVERY MAPPED FIXTURE EXISTS AND CARRIES WHAT ITS KIND PROMISES. A case that
    # maps to P5-AN but emits no `expected` block would be "covered" by a
    # scenario with nothing to assert.
    $hollow = @()
    foreach ($case in @($cases.plan_cases)) {
        $id = [string]$case.id
        $names = @($ledger[$id])
        switch ($case.kind) {
            'analytical' {
                if ($names -notcontains 'P5-AN') { $hollow += ($id + ': analytical, not in P5-AN') }
                if ($null -eq $case.expected) { $hollow += ($id + ': analytical with no expected block') }
            }
            'refusal' {
                if ($names -notcontains 'P5-RF') { $hollow += ($id + ': refusal, not in P5-RF') }
                if ([string]::IsNullOrEmpty([string]$case.expected_refusal)) {
                    $hollow += ($id + ': refusal with no expected_refusal')
                }
            }
            'statistics' {
                if (@($case.statistics).Count -lt 1) { $hollow += ($id + ': statistics with no vectors') }
            }
            'fingerprint' {
                if ([string]::IsNullOrEmpty([string]$case.reference)) {
                    $hollow += ($id + ': fingerprint with no reference')
                }
            }
        }
    }
    $null = Add-Check $list 'every mapped fixture carries the evidence its kind promises' `
        ($hollow.Count -eq 0) ($hollow -join '; ')

    # THE DIRECT-VECTOR SETS ARE COMPLETE. Counted from the corpus, so a vector
    # dropped upstream cannot shrink Gate-B coverage silently.
    $numeric = @($cases.fingerprint.numeric_encodings.vectors)
    $null = Add-Check $list 'the canonical numeric vector set is present' `
        ($numeric.Count -ge 10) ("vectors " + $numeric.Count)
    $labels = @($numeric | ForEach-Object { [string]$_.label })
    foreach ($required in '0', '-0', '1', '-1', '0.1', '1e-20', '1e+20', '0.1 + 0.2',
                          'MAX_DOUBLE', 'minimum subnormal') {
        $null = Add-Check $list ("the locked numeric vector '" + $required + "' is present") `
            ($labels -contains $required)
    }
    $sep = @($cases.fingerprint.decimal_separator.vectors)
    $null = Add-Check $list 'the separator vector set is present' ($sep.Count -ge 10) `
        ("vectors " + $sep.Count)
    $null = Add-Check $list 'every separator vector states BOTH a point and a comma expectation' `
        (@($sep | Where-Object { $null -eq $_.point -or $null -eq $_.comma }).Count -eq 0)
    $reduce = @($cases.fingerprint.reduction_vectors)
    $null = Add-Check $list 'all four reduction vectors are present' ($reduce.Count -eq 4) `
        ("vectors " + $reduce.Count)
    $utf16 = @($cases.fingerprint.utf16_vectors.vectors)
    $null = Add-Check $list 'the UTF-16 vector set is present' ($utf16.Count -ge 3) `
        ("vectors " + $utf16.Count)
    $utfKeys = @($utf16 | ForEach-Object { [string]$_.key })
    foreach ($required in 'bmp_above_7fff', 'non_bmp', 'mixed_length_prefix') {
        $null = Add-Check $list ("the locked UTF-16 vector '" + $required + "' is present") `
            ($utfKeys -contains $required)
    }
    $null = Add-Check $list 'the reference stream states BOTH a code-unit count and a digest' `
        (([int]$cases.fingerprint.reference.code_units -gt 0) -and `
         (-not [string]::IsNullOrEmpty([string]$cases.fingerprint.reference.digest)))
    $null = Add-Check $list 'the reference stream is as long as the corpus says' `
        (([string]$cases.fingerprint.reference.stream).Length -eq [int]$cases.fingerprint.reference.code_units) `
        ("stream " + ([string]$cases.fingerprint.reference.stream).Length + `
         ", stated " + [int]$cases.fingerprint.reference.code_units)
    $null = Add-Check $list 'the collision probes are present' `
        (@($cases.fingerprint.collision_probes).Count -ge 8)

    $ok = Test-ChecklistOk $list
    Add-Phase5Result 'P5-PRE' `
        ("Phase-5 coverage preflight: " + $emitted.Count + " plan cases mapped (no Excel)") `
        $(if ($ok) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    return $ok
}

# ===========================================================================
# VBA CODE VERSUS COMMENTARY
# ===========================================================================
# RUN-2 ROOT. P5-EV read each CodeModule's whole text and searched it literally
# for every manifest forbidden_construct:
#
#   FAIL no forbidden construct exists in the real Stage-B project
#        -- modAppState: Worksheet_Change; modAppState: NPV
#
# Both hits are PROSE in accepted production source:
#
#   modAppState.bas:7   ' ... No cost, risk, escalation, FX, NPV, EMV,
#   modAppState.bas:78  ' ... no input Worksheet_Change handler, and this
#                       '     guarantees that stays true even if
#
# A comment explaining that there is no Worksheet_Change handler was read as a
# Worksheet_Change handler. That is a harness false positive with no product
# meaning, and the comment must stay: it is the reason the guarantee exists.
#
# The Python side already draws this line - builder/pccm_builder/vba_source.py
# strips comments and string literals before any structural scan, for exactly
# this reason. This is the same rule, applied at runtime to the code Excel
# actually holds.
#
# THE SCAN IS NOT WEAKENED. It still runs over every component, still uses the
# manifest's own list, and a real declaration still fails. What it no longer
# does is read prose as code. Two things are deliberately NOT done:
#   * no construct is removed from the manifest list;
#   * no blanket text substitution is applied - a stripped line keeps its
#     length-zero content, so a real `Private Sub Worksheet_Change(` on the very
#     next line is untouched and still fails.
#
# A VBA comment starts at an apostrophe that is NOT inside a double-quoted
# string, and runs to end of line. Doubled quotes ("") are the VBA escape and
# stay inside the literal. Rem-form comments are handled by the same rule: a
# line whose first token is Rem is commentary.
function Remove-VbaCommentary {
    param([string]$Code)
    $out = New-Object System.Text.StringBuilder
    foreach ($line in ($Code -split "`r?`n")) {
        $inString = $false
        # Rem is a STATEMENT, so it is commentary wherever a statement may
        # begin: at the start of the line, or after a colon separator. This
        # tracks that boundary through the same single pass that already
        # understands string literals, rather than post-matching a regex over
        # the whole line - a regex cannot see whether a colon was inside a
        # literal, and a line-anchored one misses the inline form entirely.
        $atStatementStart = $true
        $kept = New-Object System.Text.StringBuilder
        $i = 0
        while ($i -lt $line.Length) {
            $ch = $line[$i]

            # REM-FORM COMMENTARY. Three conditions, all required:
            #   * outside a string literal, so x = "Rem NPV" is data;
            #   * at a statement boundary, so nothing mid-expression matches;
            #   * the COMPLETE keyword - followed by whitespace or end of line -
            #     so Remember and RemoteValue are identifiers, not comments.
            # Once it begins, the rest of the physical line is commentary.
            if ((-not $inString) -and $atStatementStart -and
                ((($i + 3) -le $line.Length)) -and
                ($line.Substring($i, 3) -eq 'Rem') -and
                (((($i + 3) -eq $line.Length)) -or [char]::IsWhiteSpace($line[$i + 3]))) {
                break
            }

            if ($ch -eq '"') {
                # A DOUBLED QUOTE INSIDE A STRING IS AN ESCAPED QUOTE, not a
                # close. Without this, "he said ""don't""" would be read as
                # closing after `said `, and the apostrophe in don't would
                # truncate the rest of a real statement.
                if ($inString -and (($i + 1) -lt $line.Length) -and ($line[$i + 1] -eq '"')) {
                    $null = $kept.Append('""')
                    $atStatementStart = $false
                    $i += 2
                    continue
                }
                $inString = -not $inString
                $null = $kept.Append($ch)
                $atStatementStart = $false
                $i++
                continue
            }

            # [char]39, not a quoted apostrophe. A double-quoted string whose
            # only content is an apostrophe desynchronises every naive
            # quote-stripping sweep that reads this file - including the
            # accepted invocation sweep in tests/test_phase4_stage_b_source.py,
            # which strips single-quoted strings before double-quoted ones. The
            # codebase already uses [char]31 for the unit separator for the same
            # class of reason.
            if (($ch -eq ([char]39)) -and (-not $inString)) { break }

            $null = $kept.Append($ch)
            # Whitespace neither opens nor closes a statement, so it leaves the
            # boundary alone: `x = 1 :   Rem ...` is still a Rem statement. A
            # colon opens the next statement, but ONLY outside a literal, so the
            # colon in "text : Rem NPV" does not.
            if (-not [char]::IsWhiteSpace($ch)) {
                $atStatementStart = (($ch -eq ':') -and (-not $inString))
            }
            $i++
        }
        $null = $out.AppendLine($kept.ToString())
    }
    return $out.ToString()
}

# THE SECOND HALF OF THE SAME RULE. A construct named inside a string literal is
# DATA, not an executable occurrence of it:
#
#     MsgBox "NPV is not available"          <- prose, in a message
#     Err.Raise 5, , "Worksheet_Change"      <- prose, in an error string
#
# The Python authority has always done both - VbaModule.code is
# strip_strings(strip_comments(raw)) - and contains_construct() scans that. The
# runtime scanner had only the first half, so the two had different semantics
# for the same question. This is the same regex the Python side uses:
# `"(?:[^"]|"")*"`, replaced by an EMPTY literal.
#
# The literal is replaced rather than deleted so the statement around it keeps
# its shape: `x = Rnd()` after `MsgBox "..."` on the same line survives intact,
# and a forbidden token that follows a string literal is still found.
function Remove-VbaStringLiterals {
    param([string]$Code)
    return [regex]::Replace($Code, '"(?:[^"]|"")*"', '""')
}

# The executable code of a module: commentary gone, string payloads emptied,
# every executable token preserved. Comments are stripped FIRST - the comment
# scanner is the one that understands string literals, so it must run while the
# literals are still intact.
function Get-VbaExecutableCode {
    param([string]$Code)
    return (Remove-VbaStringLiterals -Code (Remove-VbaCommentary -Code $Code))
}

# Does the code declare a procedure with this name? Sub/Function, any accessor,
# optionally Static, and the name must be followed by '(' - so a mention of the
# identifier inside another statement is still caught by the general scan while
# a DECLARATION is reported as the declaration it is.
function Test-VbaProcedureDeclared {
    param([string]$Code, [string]$ProcedureName)
    $pattern = '(?im)^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function)\s+' +
               [regex]::Escape($ProcedureName) + '\s*\('
    return ($Code -match $pattern)
}

# ===========================================================================
# THE VBPROJECT COMPONENT INVENTORY
# ===========================================================================
# RUN-2 ROOT. P5-M and P5-D8 both enumerated VBComponents and compared every
# component NAME against the manifest's 15-entry vba.modules collection:
#
#   FAIL the inventory is exactly the 15 manifest modules again -- present 30 of 15
#   FAIL no module outside the manifest persists -- extra: ThisWorkbook,
#        shDashboard, shSetup, ... shSimData
#
# All fifteen production modules were individually confirmed present and the
# diagnostic procedure was no longer callable. The 30 is arithmetic, not a
# defect: a VBProject holds one DOCUMENT component per worksheet plus one for
# ThisWorkbook, and this workbook has 14 sheets.
#
#   15 standard modules + 14 sheet documents + 1 ThisWorkbook = 30 components
#
# The manifest's vba.modules describes the production STANDARD MODULES. It has
# never described document components, and it never could - they are created by
# Excel when a sheet exists, not imported by the bootstrap.
#
# So the inventory is partitioned BY COMPONENT TYPE and each partition is judged
# against what actually governs it. Nothing is weakened to "at least 15": the
# standard-module set must still equal the manifest set exactly, in both
# directions, by name.
#
# The VBIDE type constants, named rather than spelled as bare integers at the
# comparison sites:
$script:VbextComponentTypes = @{
    StdModule        = 1     # vbext_ct_StdModule    - the production namespace
    ClassModule      = 2     # vbext_ct_ClassModule
    MSForm           = 3     # vbext_ct_MSForm
    ActiveXDesigner  = 11    # vbext_ct_ActiveXDesigner
    Document         = 100   # vbext_ct_Document     - sheets and ThisWorkbook
}

# ===========================================================================
# WHAT THE PERSISTED PROJECT ACTUALLY DECLARES
# ===========================================================================
# REVIEW OF ae52bdd, BLOCKER 2. P5-M reported "the API procedure PCCM_Calculate
# is callable" as a PASS while never invoking it: the branch that skipped the
# call set the flag to true and carried a note about a future exercise. An
# expected future exercise is not present evidence.
#
# The honest replacement for the six is DECLARATION evidence, read from the
# persisted project's own code through the SAME machinery P5-EV uses -
# CodeModule text, comment- and literal-stripped by Get-VbaExecutableCode, and
# matched as a real declaration by Test-VbaProcedureDeclared. A procedure named
# only in a comment or a string is not declared, and a manifest that names it is
# not the project that holds it.
#
# This is deliberately NOT callability. It says the name exists in code. What
# crossing Application.Run proves is a different claim, made only where it is
# actually observed.
function Get-Phase5ProjectProcedureNames {
    param($Workbook)
    $names = @()
    $project = $null; $components = $null
    try {
        $project = $Workbook.VBProject
        $components = $project.VBComponents
        $count = [int]$components.Count
        for ($i = 1; $i -le $count; $i++) {
            $component = $null; $module = $null
            try {
                $component = $components.Item($i)
                $module = $component.CodeModule
                if ([int]$module.CountOfLines -gt 0) {
                    $raw = [string]$module.Lines(1, [int]$module.CountOfLines)
                    $code = Get-VbaExecutableCode -Code $raw
                    foreach ($line in ($code -split "`n")) {
                        $match = [regex]::Match(
                            $line,
                            '^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function)\s+(\w+)\s*\(')
                        if ($match.Success) { $names += [string]$match.Groups[1].Value }
                    }
                }
            } finally {
                if ($null -ne $module)    { Release-Transient $module    'CodeModule';  $module    = $null }
                if ($null -ne $component) { Release-Transient $component 'VBComponent'; $component = $null }
            }
        }
    } finally {
        if ($null -ne $components) { Release-Transient $components 'VBComponents'; $components = $null }
        if ($null -ne $project)    { Release-Transient $project    'VBProject';    $project    = $null }
    }
    return $names
}

function Get-VbComponentTypeName {
    param([int]$TypeValue)
    foreach ($key in $script:VbextComponentTypes.Keys) {
        if ([int]$script:VbextComponentTypes[$key] -eq $TypeValue) { return $key }
    }
    return ('type' + [string]$TypeValue)
}

# Every component as plain data: name and numeric type, nothing else. The COM
# objects are released inside the loop and none escapes.
function Get-Phase5VbComponentInventory {
    param($Workbook)
    # ONE RECORD PER COMPONENT, EMITTED. Not one array returned.
    #
    # This used to accumulate into $out and end with `return ,$out`. The unary
    # comma exists to stop PowerShell unrolling a collection - it is the right
    # tool for a function that must hand back ONE row whose own elements must not
    # become separate pipeline objects, which is why Write-RowObject uses
    # -NoEnumerate for table rows. It is the WRONG tool here: this function
    # produces a SEQUENCE of records, and wrapping the sequence made the caller's
    # @(...) see a single nested array. Every downstream
    #
    #     $Components | Where-Object { [int]$_.Type -eq ... }
    #
    # would then filter one array-shaped object with no .Type at all, and no
    # partition would ever match. A textual source test cannot see that; the
    # cardinality has to be stated in the emission itself.
    #
    # The contract, deliberately the plain pipeline one:
    #     zero components -> nothing is emitted
    #     one component   -> one PSCustomObject
    #     N components    -> N PSCustomObjects
    # and the caller's @(...) is what turns 0/1/N into an Object[].
    #
    # A PSCustomObject is not a collection, so emitting one cannot unroll into
    # its properties. There is nothing here for the comma to protect.
    #
    # NOTE: this says nothing about Get-Phase5TypedTableBody. That function emits
    # one object[] PER ROW and must keep -NoEnumerate, or a row's cells would
    # each become a pipeline object and the row boundaries would be lost.
    $project = $null; $components = $null
    try {
        $project = $Workbook.VBProject
        $components = $project.VBComponents
        $count = [int]$components.Count
        for ($i = 1; $i -le $count; $i++) {
            $component = $null
            try {
                $component = $components.Item($i)
                # Plain data only, read before the COM object is released.
                Write-Output ([pscustomobject]@{
                    Name = [string]$component.Name
                    Type = [int]$component.Type
                })
            } finally {
                if ($null -ne $component) { Release-Transient $component 'VBComponent'; $component = $null }
            }
        }
    } finally {
        if ($null -ne $components) { Release-Transient $components 'VBComponents'; $components = $null }
        if ($null -ne $project)    { Release-Transient $project    'VBProject';     $project    = $null }
    }
}

function Format-VbComponentList {
    param($Components)
    return (@($Components | ForEach-Object {
        [string]$_.Name + ' (' + (Get-VbComponentTypeName -TypeValue ([int]$_.Type)) + ')'
    }) -join ', ')
}

# The whole inventory judgement, shared by P5-M and P5-D8 so the two cannot
# drift apart. $ExpectedSheetCount is the manifest's own sheet count; the +1 is
# ThisWorkbook, which every workbook has exactly one of.
function Add-Phase5ModuleInventoryChecks {
    param($List, $Components, $ExpectedModules, [int]$ExpectedSheetCount, [string]$Label)

    $standard = @($Components | Where-Object { [int]$_.Type -eq $script:VbextComponentTypes.StdModule })
    $documents = @($Components | Where-Object { [int]$_.Type -eq $script:VbextComponentTypes.Document })
    $other = @($Components | Where-Object {
        ([int]$_.Type -ne $script:VbextComponentTypes.StdModule) -and
        ([int]$_.Type -ne $script:VbextComponentTypes.Document) })
    $standardNames = @($standard | ForEach-Object { [string]$_.Name })

    # THE PRODUCTION NAMESPACE, EXACTLY. Both directions, by name, no tolerance.
    foreach ($name in $ExpectedModules) {
        $null = Add-Check $List ($Label + ': the production module ' + $name + ' is a standard module') `
            ($standardNames -contains $name) `
            ('standard modules: ' + ($standardNames -join ', '))
    }
    $strayStandard = @($standardNames | Where-Object { $ExpectedModules -notcontains $_ })
    $null = Add-Check $List ($Label + ': no standard module outside the manifest persists') `
        ($strayStandard.Count -eq 0) ('stray standard modules: ' + ($strayStandard -join ', '))
    $null = Add-Check $List `
        ($Label + ': the standard-module set is exactly the ' + [string]@($ExpectedModules).Count +
         ' manifest modules') `
        ($standardNames.Count -eq @($ExpectedModules).Count) `
        ('present ' + $standardNames.Count + ' of ' + @($ExpectedModules).Count + ': ' +
         ($standardNames -join ', '))

    # THE DOCUMENT COMPONENTS ARE COUNTED, NOT WAVED THROUGH. One per sheet plus
    # ThisWorkbook - so a stray document component cannot hide here either.
    $expectedDocuments = $ExpectedSheetCount + 1
    $null = Add-Check $List `
        ($Label + ': the document components are the ' + [string]$ExpectedSheetCount +
         ' sheets plus ThisWorkbook') `
        ($documents.Count -eq $expectedDocuments) `
        ('found ' + $documents.Count + ', expected ' + $expectedDocuments + ': ' +
         (Format-VbComponentList $documents))
    $null = Add-Check $List ($Label + ': exactly one ThisWorkbook document component') `
        (@($documents | Where-Object { [string]$_.Name -eq 'ThisWorkbook' }).Count -eq 1) `
        (Format-VbComponentList $documents)

    # AND NOTHING ELSE AT ALL. A class module, a UserForm or an ActiveX designer
    # is code that the manifest does not describe and the bootstrap did not put
    # there, whatever it is called.
    $null = Add-Check $List `
        ($Label + ': no class module, UserForm or designer component exists') `
        ($other.Count -eq 0) (Format-VbComponentList $other)

    # The whole-project arithmetic, stated so the evidence is self-checking.
    $null = Add-Check $List `
        ($Label + ': the project is exactly ' + [string]@($ExpectedModules).Count +
         ' standard modules + ' + [string]$expectedDocuments + ' document components') `
        (@($Components).Count -eq (@($ExpectedModules).Count + $expectedDocuments)) `
        ('total components ' + @($Components).Count + ': ' + (Format-VbComponentList $Components))
}

# ===========================================================================
# ONE RESULT PER SCENARIO ID
# ===========================================================================
# RUNTIME RUN 4 LEDGER DEFECT. The final report said 19 failed over 17 unique
# Phase-5 IDs, because P5-S2 and P5-ST were each recorded TWICE: the try block
# recorded both scenarios on their real evidence, then a LATER setup step in the
# same try - restoring the base fixture - threw, and the enclosing catch
# recorded both IDs again as failures. The second record overwrote nothing; it
# simply appended, so "38 Phase-5 scenarios reported" was a count of records,
# not of scenarios.
#
# Two things are wrong there and both are fixed:
#
#   OWNERSHIP  a setup step that runs AFTER a scenario has been recorded does
#              not belong inside that scenario's try. Those steps now sit in
#              their own try/catch and report as P5-SU, a setup failure, which
#              is what they actually are.
#
#   STRUCTURE  a catch that emits several IDs must not re-emit one that already
#              has a result. Add-Phase5Result records every ID it emits and
#              refuses a second, turning the attempt into a NOTE so the
#              suppression is visible rather than silent.
#
# The guard is not a print-time filter: nothing downstream de-duplicates, and
# the ledger genuinely contains one record per ID.
$script:Phase5RecordedIds = New-Object System.Collections.ArrayList
$script:Phase5LedgerViolations = New-Object System.Collections.ArrayList
$script:Phase5LedgerReported = $false

function Reset-Phase5ResultLedger {
    $script:Phase5RecordedIds = New-Object System.Collections.ArrayList
    $script:Phase5LedgerViolations = New-Object System.Collections.ArrayList
    $script:Phase5LedgerReported = $false
}

function Test-Phase5ResultRecorded {
    param([string]$Id)
    return (@($script:Phase5RecordedIds) -contains $Id)
}

function Get-Phase5LedgerViolations { return @($script:Phase5LedgerViolations) }

function Add-Phase5Result {
    param([string]$Id, [string]$Name, [string]$Status, [string]$Detail = '')
    if (Test-Phase5ResultRecorded -Id $Id) {
        # A DUPLICATE ATTEMPT IS ITSELF A HARNESS-INTEGRITY FAILURE, and it must
        # not be reducible to a Note. Review round 4A: the driver declares
        # success from the FAIL count, and Notes do not contribute to it - so a
        # future ownership defect could record P5-X as PASS, attempt P5-X as
        # FAIL, have the attempt suppressed, and still finish green. That is
        # fail-open behaviour in an evidence harness.
        #
        # The first result still stands and no second result with this ID is
        # appended, so the one-result-per-ID property is unchanged. What changes
        # is that the attempt is recorded as a violation, and P5-LDG turns any
        # violation into a real FAIL. Throwing from here instead would take the
        # shutdown ledger, Y, Z and P5-FIN down with it.
        $null = $script:Phase5LedgerViolations.Add(
            $Id + ' (attempted as ' + $Status + '): ' + $Detail)
        Add-Note ('P5 ledger VIOLATION: a second result for ' + $Id +
                  ' was attempted (' + $Status + ') and refused. The first result ' +
                  'stands and P5-LDG will FAIL the run. Detail: ' + $Detail)
        return
    }
    $null = $script:Phase5RecordedIds.Add($Id)
    Add-Result $Id $Name $Status $Detail
}

function Add-Phase5LedgerIntegrityResult {
    # ONE result, always, whatever happened above. It is emitted through
    # Add-Result rather than Add-Phase5Result so that the ledger's own report can
    # never be suppressed by the ledger, and it carries its own emitted-once
    # flag so that many duplicate attempts still produce exactly one P5-LDG.
    if ($script:Phase5LedgerReported) { return }
    $script:Phase5LedgerReported = $true
    $violations = @($script:Phase5LedgerViolations)
    if ($violations.Count -eq 0) {
        Add-Result 'P5-LDG' 'Phase-5 result ledger: one result per scenario ID' 'PASS' `
            ('scenario results recorded: ' + @($script:Phase5RecordedIds).Count +
             '; duplicate attempts: 0')
        return
    }
    Add-Result 'P5-LDG' 'Phase-5 result ledger: one result per scenario ID' 'FAIL' `
        ('a scenario result was attempted more than once. The first result for ' +
         'each ID stands, but a duplicate attempt means a scenario boundary owns ' +
         'a failure that is not its own, so the run cannot be trusted: ' +
         ($violations -join ' | '))
}

# ===========================================================================
# UNEXPECTED-ERROR EVIDENCE
# ===========================================================================
# RUN-2 DIAGNOSTIC GAP. Every Phase-5 catch site reported through the accepted
# Phase-4 Format-Err helper, which returns exception TYPE and MESSAGE and nothing else.
# Run 2 therefore recorded eleven scenarios as
#
#   System.InvalidCastException: Unable to cast object of type 'System.Double'
#   to type 'System.String'.
#
# with no file, no line and no call chain - the same sentence eleven times, and
# no way to tell which statement produced it. The message named a type pair that
# PowerShell's own [string] conversion cannot fail on, so the text alone did not
# even identify the KIND of boundary involved.
#
# This adds the location. It does NOT change Format-Err: that helper is accepted
# Phase-4 source and its callers are Phase-4 scenarios.
#
# PLAIN DATA ONLY. Strings, integers and type names are read out of the
# ErrorRecord. No COM object, no Range, no Workbook and no Application is
# captured, held or rendered - the diagnostic ledger must never become a reason
# for an RCW to outlive its scope.
function Format-Phase5Err {
    param($ErrorRecord)
    if ($null -eq $ErrorRecord) { return 'unknown error' }
    $parts = @()

    # The exception chain, outermost first. A MethodInvocationException wrapping
    # the real cause is exactly the shape a .NET or COM binding fault takes, and
    # reporting only the outer type hides it.
    $exception = $null
    try { $exception = $ErrorRecord.Exception } catch { }
    $depth = 0
    while (($null -ne $exception) -and ($depth -lt 5)) {
        $type = ''
        try { $type = [string]$exception.GetType().FullName } catch { $type = 'unknown type' }
        $message = ''
        try { $message = [string]$exception.Message } catch { }
        if ($depth -eq 0) {
            $parts += ($type + ': ' + $message)
        } else {
            $parts += ('  inner[' + [string]$depth + '] ' + $type + ': ' + $message)
        }
        $next = $null
        try { $next = $exception.InnerException } catch { }
        $exception = $next
        $depth++
    }
    if ($parts.Count -eq 0) {
        try { $parts += [string]$ErrorRecord } catch { $parts += 'unknown error' }
    }

    # WHERE. Script, line, column and the offending source line itself.
    $invocation = $null
    try { $invocation = $ErrorRecord.InvocationInfo } catch { }
    if ($null -ne $invocation) {
        $script = ''
        try { $script = [string]$invocation.ScriptName } catch { }
        if ([string]::IsNullOrWhiteSpace($script)) { $script = '<no script>' }
        else { $script = Split-Path -Leaf $script }
        $line = 0
        try { $line = [int]$invocation.ScriptLineNumber } catch { }
        $column = 0
        try { $column = [int]$invocation.OffsetInLine } catch { }
        $parts += ('  at ' + $script + ':' + [string]$line + ' col ' + [string]$column)
        $text = ''
        try { $text = [string]$invocation.Line } catch { }
        if (-not [string]::IsNullOrWhiteSpace($text)) {
            $parts += ('  source: ' + $text.Trim())
        }
    }

    # THE CALL CHAIN. Which helper the throwing statement was reached through is
    # the whole question when one shared path serves eleven scenarios.
    $stack = ''
    try { $stack = [string]$ErrorRecord.ScriptStackTrace } catch { }
    if (-not [string]::IsNullOrWhiteSpace($stack)) {
        foreach ($frame in ($stack -split "`r?`n")) {
            if (-not [string]::IsNullOrWhiteSpace($frame)) { $parts += ('  ' + $frame.Trim()) }
        }
    }

    $category = ''
    try { $category = [string]$ErrorRecord.CategoryInfo.Category } catch { }
    $target = ''
    try { $target = [string]$ErrorRecord.CategoryInfo.TargetType } catch { }
    if (-not [string]::IsNullOrWhiteSpace($category)) {
        $suffix = ''
        if (-not [string]::IsNullOrWhiteSpace($target)) { $suffix = ', target type ' + $target }
        $parts += ('  category: ' + $category + $suffix)
    }

    return ($parts -join [string][char]10)
}

# ===========================================================================
# Reading the calculation workspace
# ===========================================================================
# Every address comes from the inspection projection. Nothing below names a cell
# or a table in its own right, so a contract change moves these reads with it
# instead of leaving them pointing at a stale coordinate.
function Get-CalcScalar {
    param($Workbook, $Inspection, [string]$Block, [string]$FieldKey)
    # RUN-2 ROOT. This used to read:
    #
    #     $block = $Inspection.calc.scalar_blocks.$Block
    #
    # PowerShell variable names are CASE-INSENSITIVE, so $block IS $Block - the
    # [string]-typed parameter. A typed parameter keeps its type constraint for
    # the life of the variable, so assigning the block PSCustomObject to it
    # CONVERTED it to a string ("@{value_column=C; rows=...}"). The next line
    # then asked a String for .rows and Set-StrictMode turned that into
    #
    #     PropertyNotFoundException: The property 'rows' cannot be found on this object
    #
    # every single time. Run 2 reported it from P5-S2, P5-ST, P5-S3, P5-S4,
    # P5-S5, P5-KP and P5-RC - seven scenarios, one defect, and nothing to do
    # with the inspection projection, which carries 'rows' for both blocks.
    #
    # The local is named for what it holds and can no longer collide with a
    # parameter. A source test scans EVERY typed parameter in this file for the
    # same shadowing, so the class of defect is closed, not just this instance.
    $blockSpec = $Inspection.calc.scalar_blocks.$Block
    $row = [int]$blockSpec.rows.$FieldKey
    $address = [string]$blockSpec.value_column + [string]$row
    # Each COM object into its OWN named variable, released in the narrowest
    # scope, exactly as every Phase-4 helper does. No chained member expression
    # creates an intermediate RCW that nothing owns.
    $sheets = $null; $sheet = $null; $range = $null
    try {
        $sheets = $Workbook.Worksheets
        $sheet = $sheets.Item($Inspection.calc.sheet)
        $range = $sheet.Range($address)
        # .Value2, not .Text: a formatted date read as text would compare against
        # a locale, and a blank read as text would become the empty string with
        # no way left to tell it from a value that really is "".
        return $range.Value2
    } finally {
        if ($null -ne $range)  { Release-Transient $range  'Range(calc scalar)'; $range  = $null }
        if ($null -ne $sheet)  { Release-Transient $sheet  'Worksheet(_Calc)';   $sheet  = $null }
        if ($null -ne $sheets) { Release-Transient $sheets 'Worksheets';         $sheets = $null }
    }
}

function Get-CalcScalarBlock {
    param($Workbook, $Inspection, [string]$Block)
    # The whole block as an ordered map key -> value, for the snapshot
    # comparisons. Read cell by cell so a BLANK stays $null rather than becoming
    # an empty string inside a Variant array.
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    $rows = $Inspection.calc.scalar_blocks.$Block.rows
    foreach ($key in $rows.PSObject.Properties.Name) {
        $out.Add($key, (Get-CalcScalar -Workbook $Workbook -Inspection $Inspection `
            -Block $Block -FieldKey $key))
    }
    return $out
}

# ===========================================================================
# THE PHASE-5 TYPED TABLE READER
# ===========================================================================
# The accepted Phase-4 Get-TableBody converts every cell to a display-neutral
# STRING - `if ($null -eq $v) { '' } else { [string]$v }`. That is right for the
# structural comparisons it was written for, and WRONG for Phase-5 analytical
# evidence, because Test-CalcValue is deliberately type-sensitive: a numeric
# expectation requires a numeric actual, and a blank must never equal a numeric
# zero. A correct Excel cell holding Value2 = 1.05 arrived as the String "1.05"
# and every analytical comparison returned False before it ever compared a
# number. The first successful Gate-B analytical scenario would have failed with
# production behaving perfectly.
#
# So Phase 5 reads the body itself, preserving what Excel published:
#
#   blank Value2   -> $null          (an absence, distinguishable from "")
#   text Value2    -> String
#   numeric Value2 -> the numeric scalar, NEVER stringified
#   Boolean        -> Boolean
#
# Nothing here formats. Formatting belongs to failure diagnostics, and a reader
# that formats is a reader that has already decided the comparison.
#
# THE ROW SHAPE IS THE ACCEPTED ONE: exactly one non-enumerated object[] per
# physical row, through Write-RowObject, so the caller's @(...) still gives
# 0/1/N with row boundaries intact. The row is allocated at the known column
# count and assigned BY INDEX, because `$line += $null` appends nothing and a
# blank cell would silently vanish from the row.
#
# The Phase-4 helper is not modified.
function Get-Phase5TypedTableBody {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $body = $null
    $rowsObj = $null; $colsObj = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $body = $lo.DataBodyRange
        # An empty body is a valid outcome: emit NOTHING, exactly as the accepted
        # reader does.
        if ($null -eq $body) { return }

        $rowsObj = $body.Rows
        $colsObj = $body.Columns
        $rowCount = [int]$rowsObj.Count
        $colCount = [int]$colsObj.Count
        Release-Transient $rowsObj 'Range(rows)'; $rowsObj = $null
        Release-Transient $colsObj 'Range(columns)'; $colsObj = $null

        for ($r = 1; $r -le $rowCount; $r++) {
            $line = New-Object 'object[]' $colCount
            for ($c = 1; $c -le $colCount; $c++) {
                $cell = $null
                try {
                    $cell = $body.Cells($r, $c)
                    # Value2, and NOTHING else. No [string], no Format, no
                    # coalescing of $null into ''.
                    $line[$c - 1] = $cell.Value2
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
}

function Get-Phase5TypedCell {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$RowIndex, [int]$ColumnIndex)
    # One cell, same discipline, for the places that need a single Value2 rather
    # than a whole body.
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null
    $body = $null; $cell = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $body = $lo.DataBodyRange
        if ($null -eq $body) { return $null }
        $cell = $body.Cells($RowIndex, $ColumnIndex)
        return $cell.Value2
    } finally {
        if ($null -ne $cell)            { Release-Transient $cell            'Range(cell)';  $cell            = $null }
        if ($null -ne $body)            { Release-Transient $body            'Range(body)';  $body            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';   $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects';  $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';    $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';   $localWorksheets = $null }
    }
}

function Set-Phase5TypedCell {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$RowIndex,
          [int]$ColumnIndex, $Value)
    # WRITES BACK THE TYPE EXCEL PUBLISHED, never a type inferred from the
    # contract. See the dispatch below for why there is one COM assignment site
    # per type and why that is not the same thing as coercion.
    # The accepted Phase-4 Set-TableCell is not modified.
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null
    $body = $null; $cell = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $body = $lo.DataBodyRange
        $cell = $body.Cells($RowIndex, $ColumnIndex)
        # ONE COM ASSIGNMENT SITE PER TYPE. Runtime Run 4 located R5 here:
        #
        #   System.InvalidCastException: Unable to cast object of type
        #   'System.Double' to type 'System.String'.
        #     at phase5_gate_b_scenarios.ps1:922   source: $cell.Value2 = $Value
        #     at Set-Phase5TypedCell -> Reset-Phase5FxTable -> Set-Phase5Fixture
        #
        # The locked seed is String 'SAR' then Double 1, restored through this
        # one helper. The String assignment succeeded; the Double assignment
        # through the SAME source line failed. PowerShell caches the COM
        # property binding PER CALL SITE, so the site had already been bound for
        # a String argument and the Double could not be marshalled through it.
        #
        # The accepted Phase-4 Set-TableCell never hit this because it has two
        # distinct assignment lines, one per branch, and each gets its own
        # binding. The polymorphic single line was written to avoid coercion and
        # instead produced a defect that no Linux test could see.
        #
        # THIS IS NOT A RETREAT TO INFERENCE. Set-TableCell asks what the value
        # OUGHT to be; this asks what Excel ACTUALLY PUBLISHED and writes that
        # same type back. A captured String '1' is written as String '1' and a
        # captured Double 1 as Double 1 - a defective text seed is still
        # restored as text, which is the whole point of the helper. The dispatch
        # is on the CAPTURED CLR type, and the cast on each branch is a no-op
        # that exists only to give the branch its own bound call site.
        #
        # An unsupported type FAILS LOUDLY rather than being coerced by
        # whichever branch happens to accept it.
        # THE SUPPORTED CAPTURED TYPES, AND ONLY THOSE. Review round 4A: the
        # earlier form accepted Single, Int16, Int32, Int64, Byte and Decimal and
        # wrote all of them as Double. That is NORMALISATION, not restoration -
        # a captured Int32 1 came back as Double 1 - and the read-back could not
        # see it, because the comparator ended in [double]$Actual -eq
        # [double]$Expected. A DateTime branch was there too, with no runtime
        # evidence behind it and no defined exact-type contract, so it is gone.
        #
        # What Run 4's restoration path actually needs is $null, String and
        # Double. Boolean is kept because Excel Value2 really does publish it and
        # it round-trips as itself. Everything else FAILS CLOSED, before any
        # assignment, naming the real CLR type - the harness does not get to
        # decide that some other numeric type "is really" a Double.
        if ($null -eq $Value) {
            $null = $cell.ClearContents()
        } elseif ($Value -is [string]) {
            $cell.Value2 = [string]$Value
        } elseif ($Value -is [bool]) {
            $cell.Value2 = [bool]$Value
        } elseif ($Value.GetType().FullName -ceq 'System.Double') {
            $cell.Value2 = [double]$Value
        } else {
            throw ("Set-Phase5TypedCell cannot restore a value of type " +
                   $Value.GetType().FullName + "; this helper restores exactly " +
                   "what Excel Value2 published - an empty cell, System.String, " +
                   "System.Double or System.Boolean - and converting any other " +
                   "type would normalise the capture instead of restoring it")
        }
    } finally {
        if ($null -ne $cell)            { Release-Transient $cell            'Range(cell)';  $cell            = $null }
        if ($null -ne $body)            { Release-Transient $body            'Range(body)';  $body            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';   $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects';  $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';    $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';   $localWorksheets = $null }
    }
}

function Get-CalcTableRows {
    param($Workbook, $Inspection, [string]$TableKey)
    # THE TYPED READER. Every `_Calc` oracle comparison therefore consumes the
    # Value2 types Excel actually published, and a workbook that wrote a number
    # as text fails rather than passing through a stringifying reader.
    $table = $Inspection.calc.tables.$TableKey
    return @(Get-Phase5TypedTableBody -Workbook $Workbook -SheetName $Inspection.calc.sheet `
        -TableName $table.table_name)
}

function Get-CalcTableColumnIndex {
    param($Inspection, [string]$TableKey, [string]$ColumnKey)
    $columns = @($Inspection.calc.tables.$TableKey.columns)
    return [array]::IndexOf($columns, $ColumnKey)
}

# ---------------------------------------------------------------------------
# Type-sensitive comparison
# ---------------------------------------------------------------------------
# BLANK IS NOT NUMERIC ZERO, and the whole N/A rule of the audit blocks rests on
# that. A comparison that coerced both to 0 would report a fabricated zero as
# correct - which is the single most valuable thing these assertions can catch.
function Test-CalcBlank {
    param($Actual)
    if ($null -eq $Actual) { return $true }
    if ($Actual -is [string] -and $Actual.Length -eq 0) { return $true }
    return $false
}

function Test-CalcValue {
    param($Actual, $Expected, [double]$Tolerance = 0.0)
    if ($null -eq $Expected) { return (Test-CalcBlank -Actual $Actual) }
    if (Test-CalcBlank -Actual $Actual) { return $false }
    if ($Expected -is [string]) {
        if (-not ($Actual -is [string])) { return $false }
        return ([string]$Actual -ceq [string]$Expected)
    }
    if ($Actual -is [string]) { return $false }
    $a = [double]$Actual
    $e = [double]$Expected
    if ($a -eq $e) { return $true }
    if ($Tolerance -le 0.0) { return $false }
    $scale = [Math]::Max([Math]::Abs($e), 1.0)
    return ([Math]::Abs($a - $e) -le ($Tolerance * $scale))
}

# ---------------------------------------------------------------------------
# SNAPSHOT IDENTITY - stricter than the analytical comparator
# ---------------------------------------------------------------------------
# Test-CalcValue answers "is this the value the oracle expected?", and for that
# question a blank and an empty string are the same absence and a tolerance is
# sometimes right. "Was this restored EXACTLY as it was?" is a different
# question, and the two must not share a comparator:
#
#   * a numeric 1 restored as the String "1" is a restoration failure, and
#     Test-CalcValue would have accepted neither - but a row-string comparison
#     would have seen "1" both times;
#   * a real Empty restored as "" is a restoration failure, and Test-CalcValue
#     treats the two as equivalent blanks by design.
#
# So snapshot identity gets its own rule: same TYPE CLASS, same value, no
# tolerance, no display-text conversion, and $null is never "".
function Test-Phase5ExactValue {
    param($Actual, $Expected)
    # EXACT CLR TYPE IDENTITY FIRST, then the value. Review round 4A: this used
    # to end in [double]$Actual -eq [double]$Expected, so a captured Int32 1
    # compared EQUAL to a restored Double 1 and the setter's normalisation was
    # invisible. Rule A - absence - is unchanged; rules B, C and D are now
    # consequences of one type gate rather than three separate probes.
    if ($null -eq $Expected) { return ($null -eq $Actual) }
    if ($null -eq $Actual) { return $false }
    if ($Actual.GetType().FullName -cne $Expected.GetType().FullName) { return $false }
    # The types are identical from here, so each comparison is between two
    # values of that one type.
    if ($Expected -is [string]) { return ([string]$Actual -ceq [string]$Expected) }
    if ($Expected -is [bool]) { return ([bool]$Actual -eq [bool]$Expected) }
    if ($Expected -is [double]) { return ([double]$Actual -eq [double]$Expected) }
    # A type outside the restoration set can still be READ from a cell, and a
    # snapshot must compare it faithfully rather than refuse it. Same type, so
    # this is a comparison of like with like.
    return ($Actual -eq $Expected)
}

function Format-Phase5Typed {
    param($Value)
    # DIAGNOSTICS ONLY. Nothing compares through this.
    if ($null -eq $Value) { return '<null>' }
    if ($Value -is [string]) { return "String'" + $Value + "'" }
    if ($Value -is [bool]) { return "Boolean:" + [string]$Value }
    return ($Value.GetType().Name + ':' +
            ([double]$Value).ToString('R', [System.Globalization.CultureInfo]::InvariantCulture))
}

function Format-CalcValue {
    param($Value)
    if ($null -eq $Value) { return '<blank>' }
    if ($Value -is [string]) {
        if ($Value.Length -eq 0) { return '<blank>' }
        return "'" + $Value + "'"
    }
    return ([double]$Value).ToString('R', [System.Globalization.CultureInfo]::InvariantCulture)
}

# ===========================================================================
# Applying an emitted fixture to the real workbook
# ===========================================================================
# The MODEL comes from phase5_cases.json; the ADDRESSES come from the manifest
# and the inspection projection. Nothing here decides what a model is.
#
# A FIXTURE MUST NOT BREAK THE INVARIANT IT IS TESTING - the Phase-4 rule, and it
# holds here too. Registers are emptied through the production delete endpoints
# and rows are keyed by the production add endpoints; only the ID counters are
# set directly, because the fixtures name CL-001 and R-001 and a permanent
# identifier is not re-issued by design.
#
# Calling those endpoints was never enough on its own. Run 5 called them and
# ignored what they said, and the fixture fabricated the orphan row this rule
# exists to forbid. What follows is that claim made checkable.
# ===========================================================================
# EVERY PRODUCTION MUTATION THE HARNESS MAKES IS CHECKED
# ===========================================================================
# RUNTIME RUN 5 ROOT. Set-Phase5Fixture invoked PCCM_AddCostLine and
# PCCM_AddRisk as `$Excel.Run(...) | Out-Null` and then wrote driver data into
# the row it ASSUMED the Add had just keyed. Whether the Add succeeded was never
# read. It had not:
#
#   Set-Phase5InflationProfileMaster rewrote the Config profile master
#   -> Inflation!tblInflation still held the PREVIOUS fixture's applied profiles
#   -> PCCM_AddCostLine -> RunDriverOperation -> ValidateStructure
#      -> CheckInflationProfiles found master and grid disagreeing
#   -> the Add raised, rolled its own ID allocation back, and recorded FAIL
#   -> the harness discarded that result and wrote description, quantity,
#      unit costs, currency, profile and distribution into row 1 anyway
#   -> row 1 held data and carried no key
#   -> PCCM_ApplyTimeline, several statements later, CORRECTLY refused with
#      [no_orphan_structural_data] tblCostLines row(s) 1 hold data but carry
#      no key.
#
# The refusal is right. The orphan was manufactured by the harness, and the
# harness reported it as a fixture-establishment failure fifty lines away from
# the operation that actually failed.
#
# TWO RULES COME OUT OF THAT, AND BOTH ARE STRUCTURAL HERE:
#
#   1. PCCM_AutomationResult IS THE AUTHORITY. That Excel.Run returned says only
#      that VBA did not raise across the COM boundary; the operation's own
#      verdict is the recorded result, and every production mutation below reads
#      it and requires OK|*.
#
#   2. THE RESULT IS CLEARED BEFORE THE OPERATION RUNS. gAutomationLastResult is
#      a single global that survives until something overwrites it, so an
#      endpoint that failed BEFORE reaching RecordResult would present the
#      PREVIOUS operation's OK| to anyone who read it afterwards - a fail-open
#      read of a stale success. PCCM_AutomationBegin calls ClearAutomation, so
#      arming immediately before the call makes the value that comes back this
#      operation's own or nothing at all. This is the accepted Phase-4 idiom;
#      Set-AppliedTimeline has used it since the Phase-4 matrix was written.
#
# Nothing here changes production behaviour. It reads production's own verdict
# instead of assuming it.
function Invoke-Phase5ProductionOperation {
    param($Excel, [string]$Operation, [string]$Stage, [string]$Argument = '',
          [switch]$WithArgument)
    $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
    if ($WithArgument) {
        $Excel.Run($Operation, $Argument) | Out-Null
    } else {
        $Excel.Run($Operation) | Out-Null
    }
    $result = [string]$Excel.Run('PCCM_AutomationResult')
    if ($result -notlike 'OK|*') {
        throw ('Gate-B harness failure at ' + $Stage + ': the production operation ' +
               $Operation + $(if ($WithArgument) { " '" + $Argument + "'" } else { '' }) +
               ' did not succeed. PCCM_AutomationResult returned ' + [char]39 +
               $result + [char]39 + '.')
    }
    return $result
}

# The workbook must be structurally coherent at both ends of every production
# mutation the fixture makes, and coherence is production's own judgement -
# PCCM_StructuralReport is ValidateStructure - never the harness's.
function Assert-Phase5StructurallyCoherent {
    param($Excel, [string]$Stage)
    $report = [string]$Excel.Run('PCCM_StructuralReport')
    if (-not [string]::IsNullOrWhiteSpace($report)) {
        throw ('Gate-B fixture establishment failed: the workbook is not structurally ' +
               'coherent ' + $Stage + '. PCCM_StructuralReport returned:' + [char]10 + $report)
    }
    return $report
}

# THE HARNESS-SIDE MIRROR OF modWorkbook.OrphanRows, and the detector for exactly
# what Run 5 manufactured: a register row holding data under a blank key.
#
# It exists so contamination is named WHERE IT IS CREATED. Production already
# refuses to mutate over an orphan - that is the refusal Run 5 recorded - but by
# then the operation being blamed is several statements away from the one that
# left the row behind. This is checked on the way in to Clear-Phase5Registers as
# well as on the way out, so a fixture that inherits a contaminated register from
# an earlier scenario fails naming the earlier scenario's damage rather than
# silently deleting it and carrying on.
#
# THE HARNESS MUST EXPOSE CONTAMINATION, NOT LAUNDER IT. There is deliberately no
# repair path here: an orphan row is not blanked, not adopted and not deleted.
#
# The predicate is production's, term for term: the key column's text is empty
# and some other column in the row is not. Get-TableBody reads Value2 and renders
# it as text, and IsNullOrWhiteSpace is Len(Trim$()) = 0, which is what
# modWorkbook.IsEmptyCell tests.
function Assert-Phase5NoUnkeyedRegisterData {
    param($Workbook, $Register, [string]$Stage)
    $body = @(Get-TableBody -Workbook $Workbook -SheetName $Register.sheet `
        -TableName $Register.table_name)
    $orphans = @()
    for ($row = 0; $row -lt $body.Count; $row++) {
        $cells = @($body[$row])
        if (-not [string]::IsNullOrWhiteSpace([string]$cells[0])) { continue }
        for ($column = 1; $column -lt $cells.Count; $column++) {
            if (-not [string]::IsNullOrWhiteSpace([string]$cells[$column])) {
                $orphans += [string]($row + 1)
                break
            }
        }
    }
    if ($orphans.Count -gt 0) {
        throw ('Gate-B fixture establishment failed: ' + [string]$Register.table_name +
               ' row(s) ' + ($orphans -join ', ') + ' hold data but carry no key ' +
               $Stage + '. Unkeyed structural data is contamination, not a fixture ' +
               'state: it is left exactly where it is so the run reports it rather ' +
               'than erasing the evidence.')
    }
}

# The Value2 of a defined name, with no [string] on the way out.
#
# Get-NamedValue is the accepted Phase-4 reader and it stringifies, which is
# right for every Phase-4 caller. It cannot tell a numeric counter of 0 from a
# TEXT counter of "0" - and production draws exactly that distinction:
# modDrivers.TryReadCounter refuses a counter that is not a whole number, so a
# reset that produced text would make the next Add refuse. The reset is proved
# with the same typed discipline the Run-4 correction established for cells.
function Get-Phase5TypedNamedValue {
    param($Workbook, [string]$DefinedName)
    $names = $null; $nm = $null; $rng = $null
    try {
        $names = $Workbook.Names
        $nm = $names.Item($DefinedName)
        $rng = $nm.RefersToRange
        return $rng.Value2
    } finally {
        if ($null -ne $rng)   { Release-Transient $rng   'Range(name)'; $rng   = $null }
        if ($null -ne $nm)    { Release-Transient $nm    'Name';        $nm    = $null }
        if ($null -ne $names) { Release-Transient $names 'Names';       $names = $null }
    }
}

function Clear-Phase5Registers {
    param($Excel, $Workbook, $Manifest)
    # FAIL CLOSED, AT EVERY STEP. This used to run the two delete endpoints with
    # their results piped to Out-Null and then write the counters, which meant a
    # register that refused to empty - because an earlier scenario left an orphan
    # row, because the counter was corrupt, because a delete raised - carried its
    # rows straight into the next fixture. Every scenario built on that fixture
    # would then be testing a model nobody wrote.
    foreach ($register in @($Manifest.registers)) {
        # ON THE WAY IN: contamination inherited from an earlier scenario is
        # named here, by the table and row that hold it, rather than surfacing
        # later as production's refusal to mutate over an orphan.
        Assert-Phase5NoUnkeyedRegisterData -Workbook $Workbook -Register $register `
            -Stage 'before the Gate-B fixture emptied the registers'

        $endpoint = 'PCCM_DeleteRiskById'
        if ([string]$register.key -eq 'cost_lines') { $endpoint = 'PCCM_DeleteCostLineById' }

        foreach ($id in @(Get-IdColumnValues -Workbook $Workbook -Info $register)) {
            $null = Invoke-Phase5ProductionOperation -Excel $Excel -Operation $endpoint `
                -Argument ([string]$id) -WithArgument `
                -Stage ('step A, emptying ' + [string]$register.table_name)
            # AND THE DELETE TOOK. A production Delete that recorded OK|cancelled
            # is a legitimate OK| and removes nothing, so the identifier is
            # proved gone rather than assumed gone.
            $remaining = @(Get-IdColumnValues -Workbook $Workbook -Info $register)
            if ($remaining -contains $id) {
                throw ('Gate-B fixture establishment failed at step A: ' + $endpoint +
                       ' reported success for ' + [string]$id + ' but ' +
                       [string]$register.table_name + ' still carries that identifier.')
            }
        }

        # POSTCONDITION: no keyed row survives, and no unkeyed row was created by
        # the emptying itself.
        $left = @(Get-IdColumnValues -Workbook $Workbook -Info $register)
        if ($left.Count -ne 0) {
            throw ('Gate-B fixture establishment failed at step A: ' +
                   [string]$register.table_name + ' still carries ' + [string]$left.Count +
                   ' identifier(s) after every one of them was deleted: ' +
                   ($left -join ', ') + '.')
        }
        Assert-Phase5NoUnkeyedRegisterData -Workbook $Workbook -Register $register `
            -Stage 'after the Gate-B fixture emptied the registers'
    }

    # THE COUNTERS ARE THE ONLY VALUES WRITTEN DIRECTLY, and only because the
    # emitted fixtures name CL-001 and R-001 while a permanent identifier is
    # never re-issued by design. The write is proved through the TYPED reader and
    # the strict comparator: modDrivers.TryReadCounter refuses a counter that is
    # not a whole number, so a reset that landed as text would make the very
    # first Add refuse - and Get-NamedValue would have reported "0" either way.
    foreach ($counter in @($Manifest.counters)) {
        $initial = [double]$counter.initial
        Set-NamedValue -Workbook $Workbook -DefinedName $counter.defined_name -Value $initial
        $readBack = Get-Phase5TypedNamedValue -Workbook $Workbook `
            -DefinedName $counter.defined_name
        if (-not (Test-Phase5ExactValue -Actual $readBack -Expected $initial)) {
            throw ('Gate-B fixture establishment failed at step A: the identity counter ' +
                   [string]$counter.defined_name + ' reads back as ' +
                   (Format-Phase5Typed $readBack) + ', not the numeric ' +
                   (Format-Phase5Typed $initial) + ' it was reset to.')
        }
    }
}

function Clear-Phase5GridBody {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$ColumnCount)
    # BLANKS the body. It does NOT delete rows.
    #
    # The Inflation grid carries reserved rows that Stage A builds and the
    # structural checks count. Deleting them to "clear" the grid would break the
    # very structure the fixture is about to calculate over - a fixture must not
    # break the invariant it is testing. A row whose key column is blank is
    # UNKEYED, which is exactly how the model reads "no profile here", and
    # Set-TableCell with $null clears contents rather than writing "" or 0.
    $rows = Get-TableRowCount -Workbook $Workbook -SheetName $SheetName -TableName $TableName
    for ($row = 1; $row -le $rows; $row++) {
        for ($column = 1; $column -le $ColumnCount; $column++) {
            Set-TableCell -Workbook $Workbook -SheetName $SheetName -TableName $TableName `
                -RowIndex $row -ColumnIndex $column -Value $null
        }
    }
}

function Clear-Phase5UserRows {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$KeepRows)
    # For a Setup/Config table whose first rows are a LOCKED seed: the seed stays
    # exactly as Stage A built it and the user rows above it are removed through
    # the accepted Phase-4 helper.
    $count = Get-TableRowCount -Workbook $Workbook -SheetName $SheetName -TableName $TableName
    for ($row = $count; $row -gt [Math]::Max($KeepRows, 1); $row--) {
        Remove-TableRow -Workbook $Workbook -SheetName $SheetName -TableName $TableName -RowIndex $row
    }
}

# ===========================================================================
# THE LOCKED FX SEED
# ===========================================================================
# `tblFXRates` row 1 is the reporting currency's own row, and Stage A builds it
# as a LOCKED seed. The Gate-B prerequisite matrix deliberately destroys it:
# PQ-10 REMOVES it (so the physical USD row shifts up into row 1), PQ-11
# duplicates it and PQ-12 rewrites its rate. `-KeepRows 1` therefore preserved
# whatever happened to be row 1 afterwards, and the next fixture inherited a
# shifted USD row or a reporting rate of 2 - deterministic cross-scenario
# contamination that would have made later scenarios refuse on the global
# reporting-currency invariant instead of the predicate they claim to test.
#
# THE SEED IS CAPTURED FROM THE REAL WORKBOOK, ONCE, BEFORE ANY PHASE-5
# MUTATION - the workbook that passed Stage-A verification, the Stage-B
# persistence checks and the Phase-4 functional matrix. It is NOT reconstructed
# from a literal and NOT taken from the emitted model: rebuilding it as "SAR, 1"
# would make the fixture manufacture the very invariant PQ-10 to PQ-12 exist to
# test, and if the BUILT seed is wrong the analytical calculation must still
# fail rather than be repaired into agreement.
$script:Phase5LockedFxSeed = $null

function Save-Phase5LockedFxSeed {
    param($Workbook, $Inspection)
    $fx = $Inspection.input_tables.fx_rates
    if ([int]$fx.locked_seed_rows -ne 1) {
        throw ("the FX table declares " + [string]$fx.locked_seed_rows +
               " locked seed rows; the Gate-B reset assumes exactly one")
    }
    # THE TYPED READER, not Get-TableBody.
    #
    # Get-TableBody stringifies, so a correctly built numeric rate of 1 was
    # captured as the String "1" - and an INCORRECTLY built text seed of "1" was
    # captured identically. The restoration then wrote ([double]$Seed.Rate),
    # silently converting a defective text seed into a number before the
    # analytical scenarios ran. That repairs the workbook into agreement with the
    # contract, which is exactly what the capture rule forbids.
    $body = @(Get-Phase5TypedTableBody -Workbook $Workbook -SheetName $fx.sheet `
        -TableName $fx.table_name)
    if ($body.Count -lt 1) {
        throw "the FX table has no body row, so there is no locked seed to capture"
    }
    $script:Phase5LockedFxSeed = [pscustomobject]@{
        Currency = $body[0][0]
        Rate     = $body[0][1]
    }
    return $script:Phase5LockedFxSeed
}

function Get-Phase5LockedFxSeed {
    if ($null -eq $script:Phase5LockedFxSeed) {
        throw ("the locked FX seed was never captured. Save-Phase5LockedFxSeed must run " +
               "on the untouched Stage-B workbook, before any Phase-5 mutation.")
    }
    return $script:Phase5LockedFxSeed
}

function Reset-Phase5FxTable {
    param($Workbook, $Inspection, $Seed)
    # SCENARIO-ISOLATION BASELINE STATE, not a production repair. Nothing here
    # changes production semantics; it undoes what the harness itself did to the
    # workbook it keeps reusing.
    $fx = $Inspection.input_tables.fx_rates
    $rows = Get-TableRowCount -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name
    if ($rows -lt 1) {
        Add-BlankTableRow -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name
        $rows = 1
    }
    # Everything after the seed row goes, whatever it is.
    for ($row = $rows; $row -gt 1; $row--) {
        Remove-TableRow -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name `
            -RowIndex $row
    }
    # Row 1 is REWRITTEN from the capture. It is not trusted to still be the seed:
    # after PQ-10 it is a shifted USD row, and after PQ-12 its rate is 2.
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
    # Read back through the TYPED reader and compare with the STRICT comparator,
    # because a restoration nobody checked is an assumption - and one checked
    # with the analytical comparator would accept a type change.
    $body = @(Get-Phase5TypedTableBody -Workbook $Workbook -SheetName $fx.sheet `
        -TableName $fx.table_name)
    if ((-not (Test-Phase5ExactValue -Actual $body[0][0] -Expected $Seed.Currency)) -or `
        (-not (Test-Phase5ExactValue -Actual $body[0][1] -Expected $Seed.Rate))) {
        throw ("the locked FX seed did not restore: row 1 is " +
               (Format-Phase5Typed $body[0][0]) + " / " + (Format-Phase5Typed $body[0][1]) +
               ", captured " + (Format-Phase5Typed $Seed.Currency) + " / " +
               (Format-Phase5Typed $Seed.Rate))
    }
}

# ONE ADD, FULLY CHECKED, AND ONLY THEN THE DATA.
#
# This is the whole of the Run-5 correction expressed as a procedure: the
# endpoint is invoked, production's own verdict is read, the row is proved to
# carry the permanent identifier the emitted fixture expects, and ONLY THEN does
# fixture data reach the worksheet. There is no path through this function that
# writes into a row production has not keyed, which is exactly the path that
# manufactured the orphan.
#
# Write-Phase5Driver is called from here and from nowhere else, so the check
# cannot be bypassed by a future caller reaching past it.
function Invoke-Phase5AddDriverAndRequireSuccess {
    param($Excel, $Workbook, $Register, [bool]$IsRisk, [int]$RowIndex, $Driver)
    $endpoint = 'PCCM_AddCostLine'
    if ($IsRisk) { $endpoint = 'PCCM_AddRisk' }
    $where = 'step F, ' + $endpoint + ' number ' + [string]$RowIndex

    # 1 and 2. Invoke, and read production's verdict for THIS invocation.
    $null = Invoke-Phase5ProductionOperation -Excel $Excel -Operation $endpoint -Stage $where

    # 3. The register grew by exactly this one keyed row. An Add that succeeded
    #    without keying a row, or that keyed two, is a different workbook from
    #    the one the fixture is describing.
    $ids = @(Get-IdColumnValues -Workbook $Workbook -Info $Register)
    if ($ids.Count -ne $RowIndex) {
        throw ('Gate-B fixture establishment failed at ' + $where + ': ' +
               [string]$Register.table_name + ' carries ' + [string]$ids.Count +
               ' keyed row(s) after ' + [string]$RowIndex + ' successful Add(s): ' +
               ($ids -join ', ') + '.')
    }

    # 4. THE ROW THE FIXTURE IS ABOUT TO WRITE INTO CARRIES A PERMANENT ID.
    #    The identifier is the register's FIRST declared column - the same fact
    #    Get-IdColumnValues reads as $row[0] and modConstants projects as
    #    REG_*_ID_COL - and the physical row is read directly, not through the
    #    blank-filtered identifier list.
    $body = @(Get-TableBody -Workbook $Workbook -SheetName $Register.sheet `
        -TableName $Register.table_name)
    if ($body.Count -lt $RowIndex) {
        throw ('Gate-B fixture establishment failed at ' + $where + ': ' +
               [string]$Register.table_name + ' has only ' + [string]$body.Count +
               ' body row(s), so row ' + [string]$RowIndex + ' does not exist.')
    }
    $issued = [string]@($body[$RowIndex - 1])[0]
    if ([string]::IsNullOrWhiteSpace($issued)) {
        throw ('Gate-B fixture establishment failed at ' + $where + ': row ' +
               [string]$RowIndex + ' of ' + [string]$Register.table_name +
               ' carries no permanent identifier. Writing fixture data into it ' +
               'would create exactly the unkeyed structural row production ' +
               'refuses to mutate over.')
    }

    # 5. AND IT IS THE IDENTIFIER THE EMITTED FIXTURE EXPECTS. The corpus names
    #    CL-001, CL-002, ... and R-001, R-002, ... in order, the counters were
    #    reset to their initial values at step A, and identifiers are issued in
    #    sequence - so the Nth Add must issue the model's Nth permanent_id.
    #
    #    Binary comparison. An identifier is an identity and 'cl-001' is not
    #    'CL-001'. This check also proves the identifier column is where the
    #    harness read it: a wrong column could not hold the expected value.
    $expected = [string]$Driver.permanent_id
    if ([string]::IsNullOrEmpty($expected)) {
        throw ('Gate-B fixture establishment failed at ' + $where +
               ': the emitted fixture driver declares no permanent_id, so the ' +
               'identifier production issued cannot be checked against anything.')
    }
    if ($issued -cne $expected) {
        throw ('Gate-B fixture establishment failed at ' + $where + ': row ' +
               [string]$RowIndex + ' of ' + [string]$Register.table_name +
               ' carries ' + [char]39 + $issued + [char]39 + ', but the emitted ' +
               'fixture expects ' + [char]39 + $expected + [char]39 + '.')
    }

    # 6. ONLY NOW.
    Write-Phase5Driver -Workbook $Workbook -Register $Register -RowIndex $RowIndex `
        -Driver $Driver -IsRisk $IsRisk
}

# ===========================================================================
# THE FIXTURE CHOREOGRAPHY
# ===========================================================================
# EVERY PRODUCTION MUTATION BEGINS AND ENDS STRUCTURALLY COHERENT. That is the
# ordering contract Run 5 broke, and the order below is the whole correction:
#
#   A  empty the registers and reset the identity counters   (checked, fail closed)
#   B  the Setup scalars
#   C  the FX table: the captured seed, then the fixture's rows
#   D  the Config profile master
#   E  PCCM_ApplyTimeline - HERE, with no driver added yet - then require OK|*
#      and a blank PCCM_StructuralReport
#   F  the drivers, one checked production Add each, data written only into a
#      row production has keyed
#   G  inflation rates and profiling weights, into the generated columns
#   H  prove the fixture ENDS coherent
#
# WHY E MOVED. Config!tblInflationProfiles is the profile master and
# modInflation.SyncProfileRows rebuilds Inflation!tblInflation from it during
# PCCM_ApplyTimeline. Step D therefore leaves master and grid deliberately
# DISAGREEING until an Apply reconciles them - and modStructuralCheck's
# CheckInflationProfiles reports exactly that disagreement. Every driver Add runs
# ValidateStructure on the way out, so an Add attempted between D and E is an Add
# attempted over a workbook that production is required to call incoherent. It
# fails, it rolls its own identifier allocation back, and the row it was supposed
# to key stays blank.
#
# Apply first makes the window empty: by the time the first Add runs, the grid
# has been rebuilt from the master and the two agree. Nothing about production
# changed to make this work - the operations are simply performed in an order
# where each one's preconditions actually hold.
#
# The Adds do not need a second Apply behind them: SyncRows preserves the year
# columns Apply generated and only adds the new driver's profiling row, and each
# Add revalidates the structure itself.
function Set-Phase5Fixture {
    param($Excel, $Workbook, $Manifest, $Inspection, $Model)

    # --- A. empty the registers and reset the identity counters -------------
    Clear-Phase5Registers -Excel $Excel -Workbook $Workbook -Manifest $Manifest

    # --- B. the Setup scalars ----------------------------------------------
    $inputs = $Inspection.inputs
    Set-NamedValue -Workbook $Workbook -DefinedName $inputs.base_year.defined_name `
        -Value ([double]$Model.timeline.base_year)
    Set-NamedValue -Workbook $Workbook -DefinedName $inputs.project_start_year.defined_name `
        -Value ([double]$Model.timeline.start_year)
    Set-NamedValue -Workbook $Workbook -DefinedName $inputs.duration_years.defined_name `
        -Value ([double]$Model.timeline.duration)
    Set-NamedValue -Workbook $Workbook -DefinedName $inputs.discount_rate.defined_name `
        -Value ([double]$Model.discount_rate)

    # --- C. FX: RESTORE THE CAPTURED SEED, then append the fixture's rows ---
    #
    # The reset comes FIRST and rewrites row 1 from the capture. It is not
    # enough to keep row 1: PQ-10 deletes the reporting row and shifts a foreign
    # one into its place, and PQ-12 leaves the reporting rate at 2.
    $fx = $Inspection.input_tables.fx_rates
    Reset-Phase5FxTable -Workbook $Workbook -Inspection $Inspection `
        -Seed (Get-Phase5LockedFxSeed)
    # The reporting currency's own row is the table's LOCKED seed row, and it
    # carries the value Stage A built: the model states SAR -> 1 as a global
    # invariant, and a fixture that wrote that value would be proving itself.
    $reporting = [string](Get-NamedValue -Workbook $Workbook `
        -DefinedName $inputs.reporting_currency.defined_name)
    $fxRow = 0
    foreach ($entry in @($Model.fx)) {
        if ([string]$entry.currency -eq $reporting) { continue }
        $fxRow++
        Add-BlankTableRow -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name
        $target = [int]$fx.locked_seed_rows + $fxRow
        Set-TableCell -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name `
            -RowIndex $target -ColumnIndex 1 -Value ([string]$entry.currency)
        if ($null -ne $entry.rate) {
            Set-TableCell -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name `
                -RowIndex $target -ColumnIndex 2 -Value ([double]$entry.rate)
        }
    }

    # --- D. the CONFIG PROFILE MASTER, which owns inflation profile rows ----
    #
    # Config!tblInflationProfiles is the Phase-4 source of truth, and
    # modInflation.SyncProfileRows REBUILDS Inflation!tblInflation from it during
    # PCCM_ApplyTimeline. A profile row planted straight into tblInflation is
    # therefore removed by the very Apply the fixture depends on - which is what
    # the first version of this loader did, leaving the rate writer searching for
    # a row production had just deleted.
    #
    # The master is synchronised here; the grid rows are created by production.
    # From this statement until step E the workbook is KNOWINGLY incoherent, and
    # no production mutation may be attempted inside that window.
    Set-Phase5InflationProfileMaster -Workbook $Workbook -Inspection $Inspection `
        -Profiles @($Model.inflation.PSObject.Properties.Name)

    # --- E. APPLY NOW, BEFORE ANY DRIVER IS ADDED ---------------------------
    #
    # A FIXTURE MUST PROVE ITS OWN PREREQUISITES BEFORE THE OPERATION UNDER TEST
    # RUNS. Both gates THROW: an unusable baseline must fail at fixture
    # establishment, loudly, and never masquerade as anything else.
    #
    # Production creates the structure here: SetYearColumns for the profiling and
    # inflation year bands, and SyncProfileRows for the inflation rows from the
    # Config master step D just rewrote. That second one is what closes the
    # incoherence window, and the structural report proves it closed.
    $applied = Invoke-Phase5ProductionOperation -Excel $Excel `
        -Operation 'PCCM_ApplyTimeline' -Stage 'step E, the structural baseline'
    $null = Assert-Phase5StructurallyCoherent -Excel $Excel `
        -Stage 'after PCCM_ApplyTimeline and before the first fixture driver was added'

    # --- F. the drivers, each Add proved before its data is written ---------
    $costReg = $null; $riskReg = $null
    foreach ($register in @($Manifest.registers)) {
        if ($register.key -eq 'cost_lines') { $costReg = $register }
        if ($register.key -eq 'risk_register') { $riskReg = $register }
    }
    $costIndex = 0
    foreach ($line in @($Model.cost_lines)) {
        $costIndex++
        Invoke-Phase5AddDriverAndRequireSuccess -Excel $Excel -Workbook $Workbook `
            -Register $costReg -IsRisk $false -RowIndex $costIndex -Driver $line
    }
    $riskIndex = 0
    foreach ($risk in @($Model.risks)) {
        $riskIndex++
        Invoke-Phase5AddDriverAndRequireSuccess -Excel $Excel -Workbook $Workbook `
            -Register $riskReg -IsRisk $true -RowIndex $riskIndex -Driver $risk
    }

    # --- G. rates and weights, into the generated columns -------------------
    Write-Phase5InflationRates -Workbook $Workbook -Manifest $Manifest -Model $Model
    Write-Phase5Weights -Workbook $Workbook -Manifest $Manifest -Model $Model

    # --- H. THE FIXTURE ENDS COHERENT ---------------------------------------
    #
    # Step G writes values into rows and columns production generated, so it can
    # only break the structure by writing somewhere production did not. Proving
    # it did not is one COM call, and the alternative is discovering it as a
    # wrong analytical answer in a scenario that has nothing to do with it.
    $null = Assert-Phase5StructurallyCoherent -Excel $Excel `
        -Stage 'at the end of Gate-B fixture establishment'

    return $applied
}

function Set-Phase5InflationProfileMaster {
    param($Workbook, $Inspection, $Profiles)
    # EXACTLY the distinct profile names the fixture model needs, in the Config
    # master. No rate is written here - the master carries identities only - and
    # the physical order is the emitted model's, which nothing downstream may
    # depend on: SyncProfileRows keys surviving rates by NAME, and the rate
    # writer looks its row up by name too.
    $master = $Inspection.input_tables.inflation_profiles
    $distinct = @()
    foreach ($name in @($Profiles)) {
        if ($distinct -notcontains [string]$name) { $distinct += [string]$name }
    }
    $rows = Get-TableRowCount -Workbook $Workbook -SheetName $master.sheet `
        -TableName $master.table_name
    # Blank every editable row first, so a profile left behind by an earlier
    # fixture cannot survive into this one.
    for ($row = 1; $row -le $rows; $row++) {
        Set-TableCell -Workbook $Workbook -SheetName $master.sheet `
            -TableName $master.table_name -RowIndex $row -ColumnIndex 1 -Value $null
    }
    $index = 0
    foreach ($name in $distinct) {
        $index++
        if ($index -gt (Get-TableRowCount -Workbook $Workbook -SheetName $master.sheet `
                -TableName $master.table_name)) {
            Add-BlankTableRow -Workbook $Workbook -SheetName $master.sheet `
                -TableName $master.table_name
        }
        # EXACT text, binary. A profile name is an identity and is never trimmed
        # or case-folded on its way in.
        Set-TableCell -Workbook $Workbook -SheetName $master.sheet `
            -TableName $master.table_name -RowIndex $index -ColumnIndex 1 -Value $name
    }
}

function Write-Phase5Driver {
    param($Workbook, $Register, [int]$RowIndex, $Driver, [bool]$IsRisk)
    # Column ORDINALS come from the manifest's own column list. The harness never
    # counts columns for itself.
    $columns = @($Register.columns)
    $set = {
        param([string]$Key, $Value)
        $ordinal = [array]::IndexOf($columns, $Key) + 1
        if ($ordinal -lt 1) { throw ("the register has no column '" + $Key + "'") }
        if ($null -eq $Value) { return }
        Set-TableCell -Workbook $Workbook -SheetName $Register.sheet `
            -TableName $Register.table_name -RowIndex $RowIndex -ColumnIndex $ordinal -Value $Value
    }
    if ($IsRisk) {
        & $set 'risk_name'          ('GateB ' + [string]$Driver.permanent_id)
        & $set 'probability'        ([double]$Driver.probability)
        & $set 'impact_min'         ([double]$Driver.min_value)
        if ($null -ne $Driver.most_likely) { & $set 'impact_most_likely' ([double]$Driver.most_likely) }
        & $set 'impact_max'         ([double]$Driver.max_value)
    } else {
        & $set 'description'        ('GateB ' + [string]$Driver.permanent_id)
        & $set 'quantity'           ([double]$Driver.quantity)
        & $set 'unit_cost_min'      ([double]$Driver.min_value)
        if ($null -ne $Driver.most_likely) { & $set 'unit_cost_most_likely' ([double]$Driver.most_likely) }
        & $set 'unit_cost_max'      ([double]$Driver.max_value)
    }
    & $set 'currency'          ([string]$Driver.currency)
    & $set 'inflation_profile' ([string]$Driver.inflation_profile)
    & $set 'distribution'      ([string]$Driver.distribution)
}

function Write-Phase5InflationRates {
    param($Workbook, $Manifest, $Model)
    # BOTH AXES ARE KEYED. The row is found by INFLATION PROFILE NAME and the
    # column by CALENDAR-YEAR HEADER.
    #
    # The row used to be an incremented counter, which assumed the model's
    # profile order equals the physical grid order. It does not: SyncProfileRows
    # rebuilds the grid in Config-master order, and nothing binds that to the
    # order the emitted model happens to list its profiles in. The column has
    # always been found by header - the first generated column is BaseYear + 1,
    # and assuming it is Start Year is the defect the Step-5 correction round
    # removed from production.
    $grid = $null
    foreach ($candidate in @($Manifest.grids)) { if ($candidate.key -eq 'inflation') { $grid = $candidate } }
    $headers = @(Get-TableColumnNames -Workbook $Workbook -SheetName $grid.sheet -TableName $grid.table_name)
    $fixed = @($grid.fixed_columns).Count

    # A surviving profile keeps its previous rates through SyncProfileRows, so
    # every rate cell is blanked before the fixture writes its own: a value left
    # by an earlier fixture must not be inherited as this one's assumption.
    $rowCount = Get-TableRowCount -Workbook $Workbook -SheetName $grid.sheet `
        -TableName $grid.table_name
    for ($row = 1; $row -le $rowCount; $row++) {
        for ($column = $fixed + 1; $column -le $headers.Count; $column++) {
            Set-TableCell -Workbook $Workbook -SheetName $grid.sheet `
                -TableName $grid.table_name -RowIndex $row -ColumnIndex $column -Value $null
        }
    }

    # RUN-6 ROOT. THE PROPERTY COLLECTION IS ENUMERATED; ITS `.Name` IS NOT
    # PROJECTED ACROSS IT.
    #
    # The inner loop used to read
    #
    #     foreach ($year in $rates.PSObject.Properties.Name)
    #
    # and `$rates` is legitimately an EMPTY object in 11 of the 28 emitted
    # models. `{}` is not malformed data: the rate map is empty exactly when the
    # timeline requires no inflation calendar year at all, which for the golden
    # plan case 1 - Base Year 2026, Start Year 2026, Duration 1 - is the correct
    # encoding of "there is nothing to inflate".
    #
    # `$collection.Name` is MEMBER ENUMERATION, and under Set-StrictMode -Version
    # 2.0 an empty collection cannot answer it: PowerShell raises
    #
    #     The property 'Name' cannot be found on this object.
    #
    # rather than returning zero names. Enumerating the PSPropertyInfo objects
    # themselves has no such edge: `.PSObject.Properties` is a real property of a
    # real PSObject whether or not it holds any members, and `foreach` over an
    # empty collection is zero iterations. Name and Value are then read from each
    # INDIVIDUAL property, which also removes the `$rates.$year` dynamic lookup.
    #
    # THE TWO EMPTINESSES ARE DIFFERENT AND BOTH ARE PRESERVED:
    #
    #   {}                 zero rate entries. The loop body never runs and no
    #                      cell is written. Plan cases 1, 2, 6, 7, 8, 16, 17,
    #                      18, 22, 25 and 30.
    #   {"2028": null}     calendar year 2028 IS an entry and its cell must be
    #                      BLANK. Plan case 14, the blank-required-rate refusal.
    #
    # Collapsing the second into the first would destroy case 14; collapsing the
    # first into an error is what Run 6 did.
    foreach ($profileProperty in $Model.inflation.PSObject.Properties) {
        $name = [string]$profileProperty.Name
        $rates = $profileProperty.Value
        $rowIndex = Find-GridRow -Workbook $Workbook -Grid $grid -Key $name
        foreach ($rateProperty in $rates.PSObject.Properties) {
            $year = [string]$rateProperty.Name
            $rateValue = $rateProperty.Value
            $ordinal = [array]::IndexOf($headers, $year) + 1
            if ($ordinal -lt 1) { throw ("no generated inflation column for calendar year " + $year) }
            # A NULL RATE IS A BLANK CELL, and the branch is BEFORE the cast.
            #
            # [double]$null is numeric ZERO in PowerShell. Casting first turned
            # plan case 14 - "blank required inflation rate" - into a rate of 0,
            # which is a VALID model. The refusal the case exists to prove could
            # never have fired, and the fixture would have quietly destroyed the
            # condition it was written to exercise.
            #
            # The value now comes off the SAME property object the year came
            # from, so the blank branch cannot be reached through a lookup that
            # missed.
            if ($null -eq $rateValue) {
                Set-TableCell -Workbook $Workbook -SheetName $grid.sheet `
                    -TableName $grid.table_name -RowIndex $rowIndex `
                    -ColumnIndex $ordinal -Value $null
            } else {
                Set-TableCell -Workbook $Workbook -SheetName $grid.sheet `
                    -TableName $grid.table_name -RowIndex $rowIndex `
                    -ColumnIndex $ordinal -Value ([double]$rateValue)
            }
        }
    }
}

function Write-Phase5Weights {
    param($Workbook, $Manifest, $Model)
    # Profiling weights are per PROJECT YEAR, in order, after the fixed columns.
    # The grid is synchronised by permanent ID, so the row is located by its own
    # identifier rather than by the order the fixture happened to add drivers in.
    foreach ($pair in @(
        @{ key = 'cost_profiling'; drivers = @($Model.cost_lines) },
        @{ key = 'risk_profiling'; drivers = @($Model.risks) })) {
        $grid = $null
        foreach ($candidate in @($Manifest.grids)) {
            if ($candidate.key -eq $pair.key) { $grid = $candidate }
        }
        $fixed = @($grid.fixed_columns).Count
        $body = @(Get-TableBody -Workbook $Workbook -SheetName $grid.sheet -TableName $grid.table_name)
        foreach ($driver in $pair.drivers) {
            $rowIndex = 0
            for ($r = 0; $r -lt $body.Count; $r++) {
                if ([string]$body[$r][0] -eq [string]$driver.permanent_id) { $rowIndex = $r + 1 }
            }
            if ($rowIndex -lt 1) {
                throw ("no profiling row for " + [string]$driver.permanent_id +
                       "; synchronisation did not key the row")
            }
            $offset = 0
            foreach ($weight in @($driver.profile_weights)) {
                $offset++
                # A NULL WEIGHT IS A BLANK CELL, for the same reason. Plan case 23
                # is "a profile summing to one hundred percent but containing a
                # BLANK"; [double]$null would have written 0 into that cell, the
                # profile would still have summed to 100%, and the case would have
                # asserted a refusal that production was never given a reason to
                # make.
                if ($null -eq $weight) {
                    Set-TableCell -Workbook $Workbook -SheetName $grid.sheet `
                        -TableName $grid.table_name -RowIndex $rowIndex `
                        -ColumnIndex ($fixed + $offset) -Value $null
                } else {
                    Set-TableCell -Workbook $Workbook -SheetName $grid.sheet `
                        -TableName $grid.table_name -RowIndex $rowIndex `
                        -ColumnIndex ($fixed + $offset) -Value ([double]$weight)
                }
            }
        }
    }
}

# ===========================================================================
# Gate-B workbook mutations
# ===========================================================================
# The plan section 18 prerequisite matrix is EMITTED, in
# `phase5_cases.json -> gate_b`. Most of its boundaries cannot be expressed as
# valid analytical models - "abc" is not a Discount Rate and a blank is not a
# Quantity - so the corpus describes WORKBOOK MUTATIONS: what to change, where,
# and what the refusal must say. This applies one.
#
# PowerShell holds no list of its own. Every predicate, target and expected
# detail token comes from the corpus, so a locked prerequisite cannot be dropped
# by editing this file.
function Invoke-Phase5Mutation {
    param($Excel, $Workbook, $Manifest, $Inspection, $Mutation)

    $registers = @{}
    foreach ($register in @($Manifest.registers)) { $registers[$register.key] = $register }
    $grids = @{}
    foreach ($grid in @($Manifest.grids)) { $grids[$grid.key] = $grid }
    $fx = $Inspection.input_tables.fx_rates

    switch ([string]$Mutation.kind) {
        'entered_structure' {
            # An ENTERED structural input. `apply_timeline` decides whether the
            # applied structure is refreshed: false is how STRUCTURE CHANGE
            # PENDING is produced, and it is the only mutation that leaves the
            # Phase-4 structural gate holding the refusal.
            $names = @{
                base_year  = 'nmBaseYear_Entered'
                start_year = 'nmStartYear_Entered'
                duration   = 'nmDuration_Entered'
            }
            Set-NamedValue -Workbook $Workbook `
                -DefinedName $Manifest.defined_names.($names[[string]$Mutation.target]) `
                -Value ([double]$Mutation.value)
            if ($Mutation.apply_timeline) {
                # THE APPLY IS A PRODUCTION MUTATION AND ITS RESULT IS READ.
                # This branch is where the mutation asks for the applied
                # structure to be REFRESHED, so an Apply that did not take means
                # the mutation did not take and the scenario below would be
                # asserting a predicate against the unmutated model.
                #
                # Requiring OK|* is not the clean-structure gate: the deformation
                # this applier exists to perform is still free to make the model
                # refuse at calculation time. It only requires that the operation
                # the corpus asked for actually happened.
                $null = Invoke-Phase5ProductionOperation -Excel $Excel `
                    -Operation 'PCCM_ApplyTimeline' `
                    -Stage ('mutation ' + [string]$Mutation.kind + '/' + [string]$Mutation.target)
            }
        }
        'named_number' {
            Set-NamedValue -Workbook $Workbook `
                -DefinedName $Inspection.inputs.([string]$Mutation.target).defined_name `
                -Value ([double]$Mutation.value)
        }
        'named_text' {
            Set-NamedValueText -Workbook $Workbook `
                -DefinedName $Inspection.inputs.([string]$Mutation.target).defined_name `
                -Text ([string]$Mutation.value)
        }
        'named_blank' {
            Clear-NamedValue -Workbook $Workbook `
                -DefinedName $Inspection.inputs.([string]$Mutation.target).defined_name
        }
        'register_cell' {
            $register = $registers[[string]$Mutation.register]
            $row = Find-RegisterRow -Workbook $Workbook -Register $register `
                -PermanentId ([string]$Mutation.permanent_id)
            $ordinal = [array]::IndexOf(@($register.columns), [string]$Mutation.column) + 1
            if ($ordinal -lt 1) { throw ("no register column " + [string]$Mutation.column) }
            Set-TableCell -Workbook $Workbook -SheetName $register.sheet `
                -TableName $register.table_name -RowIndex $row -ColumnIndex $ordinal `
                -Value (Get-MutationValue $Mutation)
        }
        'fx_row' {
            $repeat = 1
            if ($null -ne $Mutation.repeat) { $repeat = [int]$Mutation.repeat }
            for ($copy = 1; $copy -le $repeat; $copy++) {
                $row = 0
                if (-not $Mutation.append) {
                    $row = Find-FxRow -Workbook $Workbook -Table $fx `
                        -Currency ([string]$Mutation.currency)
                }
                if ($row -lt 1) {
                    $row = (Get-TableRowCount -Workbook $Workbook -SheetName $fx.sheet `
                        -TableName $fx.table_name) + 1
                    Add-BlankTableRow -Workbook $Workbook -SheetName $fx.sheet `
                        -TableName $fx.table_name
                    Set-TableCell -Workbook $Workbook -SheetName $fx.sheet `
                        -TableName $fx.table_name -RowIndex $row -ColumnIndex 1 `
                        -Value ([string]$Mutation.currency)
                }
                Set-TableCell -Workbook $Workbook -SheetName $fx.sheet `
                    -TableName $fx.table_name -RowIndex $row -ColumnIndex 2 `
                    -Value (Get-MutationValue $Mutation -Property 'rate')
            }
        }
        'fx_remove' {
            $row = Find-FxRow -Workbook $Workbook -Table $fx -Currency ([string]$Mutation.currency)
            if ($row -ge 1) {
                Remove-TableRow -Workbook $Workbook -SheetName $fx.sheet `
                    -TableName $fx.table_name -RowIndex $row
            }
        }
        'inflation_cell' {
            $grid = $grids['inflation']
            $row = Find-GridRow -Workbook $Workbook -Grid $grid -Key ([string]$Mutation.profile)
            $headers = @(Get-TableColumnNames -Workbook $Workbook -SheetName $grid.sheet `
                -TableName $grid.table_name)
            $ordinal = [array]::IndexOf($headers, [string]$Mutation.calendar_year) + 1
            if ($ordinal -lt 1) { throw ("no inflation column for " + [string]$Mutation.calendar_year) }
            Set-TableCell -Workbook $Workbook -SheetName $grid.sheet `
                -TableName $grid.table_name -RowIndex $row -ColumnIndex $ordinal `
                -Value (Get-MutationValue $Mutation)
        }
        'config_profile_add' {
            # THROUGH THE CONFIG MASTER AND PRODUCTION SYNC, never by adding a
            # row to tblInflation. A grid row with no Config entry breaks the
            # Phase-4 invariant, and ResolveModel then refuses at
            # ValidateStructure before referenced-only semantics are reached -
            # so a grid-only addition proves the structural gate, not the
            # no-block rule it claims to prove.
            $master = $Inspection.input_tables.inflation_profiles
            $row = (Get-TableRowCount -Workbook $Workbook -SheetName $master.sheet `
                -TableName $master.table_name) + 1
            Add-BlankTableRow -Workbook $Workbook -SheetName $master.sheet `
                -TableName $master.table_name
            Set-TableCell -Workbook $Workbook -SheetName $master.sheet `
                -TableName $master.table_name -RowIndex $row -ColumnIndex 1 `
                -Value ([string]$Mutation.profile)
            if ($Mutation.apply_timeline) {
                # SyncProfileRows creates the matching Inflation row, with BLANK
                # rates: a new profile arrives incomplete by construction, which
                # is exactly the condition under test.
                # Through the checked helper, so the result this reads is
                # THIS Apply's own: gAutomationLastResult survives until
                # something overwrites it, and an endpoint that failed before
                # reaching RecordResult would otherwise present the previous
                # operation's OK| as its own success.
                $null = Invoke-Phase5ProductionOperation -Excel $Excel `
                    -Operation 'PCCM_ApplyTimeline' `
                    -Stage 'mutation inflation_profile_add, applying the Config-master addition'
            }
            if ($Mutation.require_clean_structure) {
                # The workbook must still be STRUCTURALLY VALID: that is half the
                # claim. An unreferenced incomplete profile is legitimate
                # structure that the calculation ignores, not a fault.
                $report = [string]$Excel.Run('PCCM_StructuralReport')
                if (-not [string]::IsNullOrWhiteSpace($report)) {
                    throw ("the unreferenced profile left the structure invalid:" +
                           [char]10 + $report)
                }
            }
        }
        'profiling_cell' {
            $grid = $grids[[string]$Mutation.grid]
            $row = Find-GridRow -Workbook $Workbook -Grid $grid -Key ([string]$Mutation.permanent_id)
            $fixed = @($grid.fixed_columns).Count
            Set-TableCell -Workbook $Workbook -SheetName $grid.sheet `
                -TableName $grid.table_name -RowIndex $row `
                -ColumnIndex ($fixed + [int]$Mutation.project_year) `
                -Value (Get-MutationValue $Mutation)
        }
        default { throw ("unknown Gate-B mutation kind '" + [string]$Mutation.kind + "'") }
    }
}

function Get-MutationValue {
    param($Mutation, [string]$Property = 'value')
    # A NULL IN THE CORPUS IS A BLANK CELL, never zero and never "". The same
    # rule the fixture writers follow, applied to the mutation matrix: several
    # locked prerequisites ARE the blank, and casting first would write the very
    # value the predicate exists to refuse the absence of.
    $raw = $Mutation.$Property
    if ($null -eq $raw) { return $null }
    if ($raw -is [string]) { return [string]$raw }
    return [double]$raw
}

function Find-RegisterRow {
    param($Workbook, $Register, [string]$PermanentId)
    $body = @(Get-TableBody -Workbook $Workbook -SheetName $Register.sheet `
        -TableName $Register.table_name)
    for ($index = 0; $index -lt $body.Count; $index++) {
        if ([string]$body[$index][0] -eq $PermanentId) { return $index + 1 }
    }
    throw ("no register row carries the permanent ID " + $PermanentId)
}

function Find-GridRow {
    param($Workbook, $Grid, [string]$Key)
    $body = @(Get-TableBody -Workbook $Workbook -SheetName $Grid.sheet -TableName $Grid.table_name)
    for ($index = 0; $index -lt $body.Count; $index++) {
        if ([string]$body[$index][0] -eq $Key) { return $index + 1 }
    }
    throw ("no grid row carries the key " + $Key)
}

function Find-FxRow {
    param($Workbook, $Table, [string]$Currency)
    $body = @(Get-TableBody -Workbook $Workbook -SheetName $Table.sheet -TableName $Table.table_name)
    for ($index = 0; $index -lt $body.Count; $index++) {
        if ([string]$body[$index][0] -eq $Currency) { return $index + 1 }
    }
    return 0
}

function Clear-NamedValue {
    param($Workbook, [string]$DefinedName)
    $names = $null; $nm = $null; $rng = $null
    try {
        $names = $Workbook.Names
        $nm = $names.Item($DefinedName)
        $rng = $nm.RefersToRange
        # A GENUINE BLANK. Writing '' would leave a zero-length string, which is
        # a value the user entered, not the absence of one.
        $null = $rng.ClearContents()
    } finally {
        if ($null -ne $rng)   { Release-Transient $rng   'Range(name)'; $rng   = $null }
        if ($null -ne $nm)    { Release-Transient $nm    'Name';        $nm    = $null }
        if ($null -ne $names) { Release-Transient $names 'Names';       $names = $null }
    }
}

function Get-Phase5PlanRefusalTokens {
    param($Cases, [string]$PlanCaseId)
    # From the corpus, never from a list here. A predicate whose discriminator
    # was dropped upstream must fail rather than silently degrade to "some error
    # occurred".
    $tokens = $Cases.gate_b.plan_refusal_tokens.$PlanCaseId
    if ($null -eq $tokens) { return @() }
    return @($tokens)
}

function Add-Phase5DetailTokenChecks {
    param($List, [string]$Detail, $Tokens, [string]$Label)
    # THE DISCRIMINATOR. "some error occurred" is not evidence that the intended
    # predicate fired. Each token is a fragment of the accepted production
    # message that names the PREDICATE, so a harmless wording edit elsewhere in
    # the sentence does not break the proof but a refusal from the WRONG
    # predicate does.
    $null = Add-Check $List ($Label + ': the refusal detail is specific, not empty') `
        (-not [string]::IsNullOrWhiteSpace($Detail)) $Detail
    foreach ($token in @($Tokens)) {
        $null = Add-Check $List ($Label + ": the detail names the predicate ('" + $token + "')") `
            ($Detail -like ('*' + $token + '*')) ("detail: " + $Detail)
    }
}

# ===========================================================================
# Asserting an emitted `expected` block, in full
# ===========================================================================
# EVERY emitted expected value, not a handful of totals. The five analytical
# ListObjects, calc_totals and calc_state are all compared, and a row count that
# does not match the fixture is a failure in its own right rather than a reason
# to compare fewer rows.
function Add-Phase5AnalyticalChecks {
    param($List, $Workbook, $Inspection, $Case, $Tolerances)

    $expected = $Case.expected
    $label = 'case ' + [string]$Case.id
    $tolerance = [double]$Tolerances.identity_relative_coefficient

    # --- tblCalcYears: EVERY column, including Calendar Year ----------------
    # Calendar Year had no emitted expectation in the first submission and was
    # therefore never asserted. The corpus now emits `expected.calc_years`, from
    # the oracle's own AppliedTimeline.project_years(), so the column is compared
    # rather than derived here as `start_year + index - 1`.
    $years = @(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection -TableKey 'calc_years')
    $expectedYears = @($expected.calc_years)
    $null = Add-Check $List ($label + ': tblCalcYears has one row per applied project year') `
        ($years.Count -eq $expectedYears.Count) `
        ("rows " + $years.Count + ", expected " + $expectedYears.Count)
    $indexColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_years' -ColumnKey 'project_index'
    foreach ($wanted in $expectedYears) {
        $found = $null
        foreach ($row in $years) {
            if ([int]$row[$indexColumn] -eq [int]$wanted.project_index) { $found = $row }
        }
        if ($null -eq $found) {
            $null = Add-Check $List ($label + ': tblCalcYears row ' + [string]$wanted.project_index) $false 'no such row'
            continue
        }
        foreach ($field in $wanted.PSObject.Properties.Name) {
            $ordinal = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_years' -ColumnKey $field
            if ($ordinal -lt 0) {
                $null = Add-Check $List ($label + ': tblCalcYears has a column for ' + $field) $false
                continue
            }
            $null = Add-Check $List ($label + ': years ' + [string]$wanted.project_index + '.' + $field) `
                (Test-CalcValue -Actual $found[$ordinal] -Expected $wanted.$field -Tolerance $tolerance) `
                ("got " + (Format-CalcValue $found[$ordinal]) + ", expected " + (Format-CalcValue $wanted.$field))
        }
    }
    # The discount factors are ALSO stated as their own emitted mapping, and the
    # two emitted authorities must agree with the same workbook cells.
    $factorColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_years' -ColumnKey 'discount_factor'
    foreach ($row in $years) {
        $projectIndex = [string][int]$row[$indexColumn]
        $wanted = $expected.discount_factors.$projectIndex
        $null = Add-Check $List ($label + ': discount factor at project index ' + $projectIndex) `
            (Test-CalcValue -Actual $row[$factorColumn] -Expected $wanted -Tolerance $tolerance) `
            ("got " + (Format-CalcValue $row[$factorColumn]) + ", expected " + (Format-CalcValue $wanted))
    }

    # --- tblCalcInflationFactors -------------------------------------------
    # The Base-Year row carries a BLANK annual rate and a unit cumulative factor.
    # `annual_rate: null` in the fixture means BLANK, and Test-CalcValue refuses
    # a numeric zero in its place.
    $factors = @(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection -TableKey 'calc_inflation_factors')
    $expectedFactors = @($expected.inflation_factors)
    $null = Add-Check $List ($label + ': tblCalcInflationFactors row count') `
        ($factors.Count -eq $expectedFactors.Count) `
        ("rows " + $factors.Count + ", expected " + $expectedFactors.Count)
    $profileColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_inflation_factors' -ColumnKey 'inflation_profile'
    $yearColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_inflation_factors' -ColumnKey 'calendar_year'
    $rateColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_inflation_factors' -ColumnKey 'annual_rate'
    $cumulativeColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_inflation_factors' -ColumnKey 'cumulative_inflation_factor'
    foreach ($wanted in $expectedFactors) {
        $found = $null
        foreach ($row in $factors) {
            if (([string]$row[$profileColumn] -eq [string]$wanted.profile) -and `
                ([int]$row[$yearColumn] -eq [int]$wanted.calendar_year)) { $found = $row }
        }
        $key = $label + ': inflation ' + [string]$wanted.profile + ' ' + [string]$wanted.calendar_year
        if ($null -eq $found) {
            $null = Add-Check $List ($key + ' row exists') $false 'no such row'
            continue
        }
        $null = Add-Check $List ($key + ' annual rate') `
            (Test-CalcValue -Actual $found[$rateColumn] -Expected $wanted.annual_rate -Tolerance $tolerance) `
            ("got " + (Format-CalcValue $found[$rateColumn]) + ", expected " + (Format-CalcValue $wanted.annual_rate))
        $null = Add-Check $List ($key + ' cumulative factor') `
            (Test-CalcValue -Actual $found[$cumulativeColumn] -Expected $wanted.cumulative_factor -Tolerance $tolerance) `
            ("got " + (Format-CalcValue $found[$cumulativeColumn]) + ", expected " + (Format-CalcValue $wanted.cumulative_factor))
    }

    # --- tblCalcFX: REFERENCED currencies only, EVERY column ----------------
    # Referenced By had no emitted expectation either, and counting driver
    # references in PowerShell would have been the harness deriving the answer it
    # is supposed to check. `expected.resolved_fx_rows` now carries the count from
    # the model the oracle was given.
    $fxRows = @(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection -TableKey 'calc_fx')
    $expectedFx = @($expected.resolved_fx_rows)
    $null = Add-Check $List ($label + ': tblCalcFX carries the referenced currencies only') `
        ($fxRows.Count -eq $expectedFx.Count) `
        ("rows " + $fxRows.Count + ", expected " + $expectedFx.Count)
    $currencyColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_fx' -ColumnKey 'currency'
    foreach ($wanted in $expectedFx) {
        $found = $null
        foreach ($row in $fxRows) {
            if ([string]$row[$currencyColumn] -eq [string]$wanted.currency) { $found = $row }
        }
        if ($null -eq $found) {
            $null = Add-Check $List ($label + ': FX row for ' + [string]$wanted.currency) $false 'no such row'
            continue
        }
        foreach ($field in $wanted.PSObject.Properties.Name) {
            $ordinal = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_fx' -ColumnKey $field
            if ($ordinal -lt 0) {
                $null = Add-Check $List ($label + ': tblCalcFX has a column for ' + $field) $false
                continue
            }
            $null = Add-Check $List ($label + ': FX ' + [string]$wanted.currency + '.' + $field) `
                (Test-CalcValue -Actual $found[$ordinal] -Expected $wanted.$field -Tolerance $tolerance) `
                ("got " + (Format-CalcValue $found[$ordinal]) + ", expected " + (Format-CalcValue $wanted.$field))
        }
    }
    # The resolved-rate mapping is a second emitted authority over the same
    # cells, and it must agree.
    $rateColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_fx' -ColumnKey 'fx_to_sar'
    foreach ($currency in @($expected.resolved_fx.PSObject.Properties.Name)) {
        $found = $null
        foreach ($row in $fxRows) { if ([string]$row[$currencyColumn] -eq $currency) { $found = $row } }
        $null = Add-Check $List ($label + ': FX rate for ' + $currency) `
            (($null -ne $found) -and (Test-CalcValue -Actual $found[$rateColumn] `
                -Expected $expected.resolved_fx.$currency -Tolerance $tolerance))
    }

    # --- tblCalcDrivers: every emitted field of every driver ----------------
    $driverRows = @(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection -TableKey 'calc_drivers')
    $expectedDrivers = @($expected.drivers)
    $null = Add-Check $List ($label + ': tblCalcDrivers row count') `
        ($driverRows.Count -eq $expectedDrivers.Count) `
        ("rows " + $driverRows.Count + ", expected " + $expectedDrivers.Count)
    $idColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_drivers' -ColumnKey 'permanent_id'
    foreach ($wanted in $expectedDrivers) {
        $found = $null
        foreach ($row in $driverRows) {
            if ([string]$row[$idColumn] -eq [string]$wanted.permanent_id) { $found = $row }
        }
        if ($null -eq $found) {
            $null = Add-Check $List ($label + ': driver row ' + [string]$wanted.permanent_id) $false 'no such row'
            continue
        }
        # Driven from the FIXTURE's own field names, so a field added to the
        # corpus is asserted here without this file being edited - and a field
        # the emitted table does not carry is reported rather than skipped.
        foreach ($field in $wanted.PSObject.Properties.Name) {
            if ($field -eq 'weights') { continue }
            $ordinal = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_drivers' -ColumnKey $field
            if ($ordinal -lt 0) {
                $null = Add-Check $List `
                    ($label + ': tblCalcDrivers has a column for expected field ' + $field) $false
                continue
            }
            $null = Add-Check $List ($label + ': ' + [string]$wanted.permanent_id + '.' + $field) `
                (Test-CalcValue -Actual $found[$ordinal] -Expected $wanted.$field -Tolerance $tolerance) `
                ("got " + (Format-CalcValue $found[$ordinal]) + ", expected " + (Format-CalcValue $wanted.$field))
        }
    }

    # --- tblCalcAnnual ------------------------------------------------------
    $annualRows = @(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection -TableKey 'calc_annual')
    $expectedAnnual = @($expected.annual)
    $null = Add-Check $List ($label + ': tblCalcAnnual row count') `
        ($annualRows.Count -eq $expectedAnnual.Count) `
        ("rows " + $annualRows.Count + ", expected " + $expectedAnnual.Count)
    $annualIndexColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_annual' -ColumnKey 'project_index'
    foreach ($wanted in $expectedAnnual) {
        $found = $null
        foreach ($row in $annualRows) {
            if ([int]$row[$annualIndexColumn] -eq [int]$wanted.project_index) { $found = $row }
        }
        if ($null -eq $found) {
            $null = Add-Check $List ($label + ': annual row ' + [string]$wanted.project_index) $false 'no such row'
            continue
        }
        foreach ($field in $wanted.PSObject.Properties.Name) {
            $ordinal = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_annual' -ColumnKey $field
            if ($ordinal -lt 0) {
                $null = Add-Check $List ($label + ': tblCalcAnnual has a column for ' + $field) $false
                continue
            }
            $null = Add-Check $List ($label + ': annual ' + [string]$wanted.project_index + '.' + $field) `
                (Test-CalcValue -Actual $found[$ordinal] -Expected $wanted.$field -Tolerance $tolerance) `
                ("got " + (Format-CalcValue $found[$ordinal]) + ", expected " + (Format-CalcValue $wanted.$field))
        }
    }

    # --- calc_totals: all ten -----------------------------------------------
    foreach ($field in @($expected.totals.PSObject.Properties.Name)) {
        $actual = Get-CalcScalar -Workbook $Workbook -Inspection $Inspection `
            -Block 'calc_totals' -FieldKey $field
        $null = Add-Check $List ($label + ': calc_totals.' + $field) `
            (Test-CalcValue -Actual $actual -Expected $expected.totals.$field -Tolerance $tolerance) `
            ("got " + (Format-CalcValue $actual) + ", expected " + (Format-CalcValue $expected.totals.$field))
    }
}

function Add-Phase5SuccessStateChecks {
    param($List, $Excel, $Workbook, $Inspection, $Case, $Cases, [string]$Label)
    # THE SUCCESSFUL calc_state RECORD, cell by cell.
    #
    # The first submission asserted the four status ACCESSORS and left C13:C20
    # itself unexamined for a successful fixture, so a commit that wrote the right
    # answers to the wrong cells would have passed. Every non-time expectation
    # here comes from an emitted authority: the fingerprint version from
    # `fingerprint.constants.FP_VERSION`, the applied-timeline text from
    # `expected.applied_timeline`, and the stored digest from the API call whose
    # value was just asserted.
    $state = Get-CalcScalarBlock -Workbook $Workbook -Inspection $Inspection -Block 'calc_state'
    $stored = [string]$Excel.Run('PCCM_CalculationFingerprint')

    # C13 - a timestamp exists. Its VALUE is not predictable and is not asserted
    # equal to anything; the contract does not make the two stamps equal.
    $null = Add-Check $List ($Label + ': C13 last successful stamp is non-blank') `
        (-not (Test-CalcBlank -Actual $state['last_successful_stamp']))
    # C14 - exactly the digest the API just reported.
    $null = Add-Check $List ($Label + ': C14 is exactly the digest PCCM_CalculationFingerprint returned') `
        ([string]$state['last_successful_fingerprint'] -ceq $stored) `
        ("cell " + (Format-CalcValue $state['last_successful_fingerprint']) + ", API '" + $stored + "'")
    # C15 - the emitted fingerprint version, not a literal.
    $null = Add-Check $List ($Label + ': C15 is the emitted fingerprint version') `
        (Test-CalcValue -Actual $state['fingerprint_version'] `
            -Expected ([double]$Cases.fingerprint.constants.FP_VERSION)) `
        ("got " + (Format-CalcValue $state['fingerprint_version']) + `
         ", emitted " + [string]$Cases.fingerprint.constants.FP_VERSION)
    # C16 - the emitted applied-timeline text.
    $null = Add-Check $List ($Label + ': C16 is the emitted applied-timeline text') `
        ([string]$state['last_successful_applied_timeline'] -ceq [string]$Case.expected.applied_timeline) `
        ("got " + (Format-CalcValue $state['last_successful_applied_timeline']) + `
         ", emitted '" + [string]$Case.expected.applied_timeline + "'")
    # C17:C20 - the attempt axis of a SUCCESS.
    $null = Add-Check $List ($Label + ': C17 = SUCCESS') `
        ([string]$state['last_attempt_result'] -eq 'SUCCESS') `
        ("got " + (Format-CalcValue $state['last_attempt_result']))
    $null = Add-Check $List ($Label + ': C18 is BLANK on success, not an empty-looking zero') `
        (Test-CalcBlank -Actual $state['last_attempt_detail']) `
        ("got " + (Format-CalcValue $state['last_attempt_detail']))
    $null = Add-Check $List ($Label + ': C19 = CURRENT after a status evaluation') `
        ([string]$state['calculation_status'] -eq 'CURRENT') `
        ("got " + (Format-CalcValue $state['calculation_status']))
    $null = Add-Check $List ($Label + ': C20 status-evaluation timestamp is non-blank') `
        (-not (Test-CalcBlank -Actual $state['status_evaluated_at']))
}

# ===========================================================================
# Snapshot capture and comparison
# ===========================================================================
# FILE SCOPE, because three scenarios need it: the refusal proof, the rollback
# proof and the status matrix. Independent review found the refusal proof
# asserting the opposite of the rule, and factoring these upward is what lets all
# three make the SAME comparison instead of three different ones.
#
# THREE GROUPS, CAPTURED SEPARATELY, because they have three different fates:
#   C13:C16  the last successful record   - must be UNCHANGED
#   C17:C20  the attempt and status axis  - must CHANGE, as the row expects
#   C23:C32 + the five tables             - must be UNCHANGED
# Comparing all of C13:C20 as unchanged would assert that the refusal or the
# failure was never recorded, which is the opposite of the requirement.
$script:Phase5SuccessRecordFields = @('last_successful_stamp', 'last_successful_fingerprint',
                                      'fingerprint_version', 'last_successful_applied_timeline')
$script:Phase5AttemptFields = @('last_attempt_result', 'last_attempt_detail',
                                'calculation_status', 'status_evaluated_at')

function Get-Phase5SuccessRecordFields { return $script:Phase5SuccessRecordFields }
function Get-Phase5AttemptFields { return $script:Phase5AttemptFields }

function Get-Phase5Snapshot {
    param($Workbook, $Inspection)
    # THE ROWS ARE KEPT AS TYPED CELL ARRAYS.
    #
    # They used to be Format-CalcValue'd and joined into one String per row, so
    # the "exact" rollback proof was really a proof about display text: a numeric
    # 1 and the String "1" produced the same evidence, and a real Empty and an
    # empty String both collapsed to "<blank>". Formatting is now used only to
    # describe a failure, never to decide one.
    $tables = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($key in $Inspection.calc.tables.PSObject.Properties.Name) {
        $rows = @()
        foreach ($row in @(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection -TableKey $key)) {
            $rows += , @($row)
        }
        $tables.Add($key, $rows)
    }
    return [pscustomobject]@{
        State  = (Get-CalcScalarBlock -Workbook $Workbook -Inspection $Inspection -Block 'calc_state')
        Totals = (Get-CalcScalarBlock -Workbook $Workbook -Inspection $Inspection -Block 'calc_totals')
        Tables = $tables
    }
}

function Add-SnapshotUnchangedChecks {
    param($List, $Before, $After, [string]$Label, $SuccessFields)
    if ($null -eq $SuccessFields) { $SuccessFields = Get-Phase5SuccessRecordFields }
    # EVERY comparison below is Test-Phase5ExactValue, not the analytical
    # comparator: this is "restored exactly as it was", where a numeric 1 that
    # came back as "1" and a blank that came back as "" are both failures.
    #
    # C13:C16 exactly.
    foreach ($field in $SuccessFields) {
        $null = Add-Check $List ($Label + ': calc_state.' + $field + ' (C13:C16) is unchanged') `
            (Test-Phase5ExactValue -Actual $After.State[$field] -Expected $Before.State[$field]) `
            ("was " + (Format-Phase5Typed $Before.State[$field]) + `
             ", now " + (Format-Phase5Typed $After.State[$field]))
    }
    # C23:C32 exactly, INCLUDING blanks. A previously blank total that came back
    # as numeric zero - or as an empty String - would be a fabricated value, not
    # a preserved one.
    foreach ($field in $Before.Totals.Keys) {
        $null = Add-Check $List ($Label + ': calc_totals.' + $field + ' (C23:C32) is unchanged') `
            (Test-Phase5ExactValue -Actual $After.Totals[$field] -Expected $Before.Totals[$field]) `
            ("was " + (Format-Phase5Typed $Before.Totals[$field]) + `
             ", now " + (Format-Phase5Typed $After.Totals[$field]))
    }
    # All five analytical ListObjects: row count, then column count, then every
    # cell. Row count first, because a table that came back shorter would
    # otherwise compare only the rows that survived.
    foreach ($key in $Before.Tables.Keys) {
        $was = @($Before.Tables[$key]); $now = @($After.Tables[$key])
        $null = Add-Check $List ($Label + ': ' + $key + ' has its previous row count') `
            ($was.Count -eq $now.Count) ("was " + $was.Count + ", now " + $now.Count)
        $identical = ($was.Count -eq $now.Count)
        $firstDifference = ''
        if ($identical) {
            for ($i = 0; $i -lt $was.Count; $i++) {
                $wasRow = @($was[$i]); $nowRow = @($now[$i])
                if ($wasRow.Count -ne $nowRow.Count) {
                    $identical = $false
                    if ($firstDifference -eq '') {
                        $firstDifference = ('row ' + ($i + 1) + ' has ' + $nowRow.Count +
                                            ' cells, was ' + $wasRow.Count)
                    }
                    continue
                }
                for ($c = 0; $c -lt $wasRow.Count; $c++) {
                    if (-not (Test-Phase5ExactValue -Actual $nowRow[$c] -Expected $wasRow[$c])) {
                        $identical = $false
                        if ($firstDifference -eq '') {
                            $firstDifference = ('row ' + ($i + 1) + ' cell ' + ($c + 1) +
                                                ': was ' + (Format-Phase5Typed $wasRow[$c]) +
                                                ', now ' + (Format-Phase5Typed $nowRow[$c]))
                        }
                    }
                }
            }
        }
        $null = Add-Check $List ($Label + ': ' + $key + ' is the previous snapshot exactly') `
            $identical $firstDifference
    }
}

function Add-Phase5AttemptAxisChecks {
    param($List, $After, [string]$Label, [string]$ExpectedResult, [string]$ExpectedStatus)
    # THE OTHER GROUP, and it must have CHANGED. Asserting C17:C20 unchanged
    # alongside C13:C16 would assert that the attempt was never recorded at all.
    $null = Add-Check $List ($Label + ': C17 = ' + $ExpectedResult) `
        ([string]$After.State['last_attempt_result'] -eq $ExpectedResult) `
        ("got " + (Format-CalcValue $After.State['last_attempt_result']))
    $null = Add-Check $List ($Label + ': C18 carries a specific detail') `
        (-not [string]::IsNullOrWhiteSpace([string]$After.State['last_attempt_detail'])) `
        ([string]$After.State['last_attempt_detail'])
    $null = Add-Check $List ($Label + ': C19 = ' + $ExpectedStatus + ', a freshly derived status') `
        ([string]$After.State['calculation_status'] -eq $ExpectedStatus) `
        ("got " + (Format-CalcValue $After.State['calculation_status']))
    $null = Add-Check $List ($Label + ': C20 carries a fresh evaluation timestamp') `
        (-not (Test-CalcBlank -Actual $After.State['status_evaluated_at']))
}

# ===========================================================================
# The scenarios
# ===========================================================================
function Invoke-Phase5GateBScenarios {
    param(
        $Excel, $Workbook, $Manifest, $Inspection, $Cases,
        [string]$ScriptDir, [string]$TempRoot, $Results
    )

    $failpoints = Get-Phase5FailpointNames
    $ledger = Get-Phase5CoverageLedger
    $successRecordFields = Get-Phase5SuccessRecordFields
    $attemptFields = Get-Phase5AttemptFields

    # -------------------------------------------------------------------
    # PHASE-4 PREREQUISITE. Every case that CAN exist here: 0 FAIL, 0 SKIP.
    # -------------------------------------------------------------------
    # Gate-B acceptance requires the structural matrix intact BEFORE a Phase-5
    # result means anything. A Phase-5 pass on a workbook whose timeline
    # machinery is broken would be evidence of nothing.
    #
    # RUN-1 SEQUENCING DEFECT. This used to demand all 35 here, Y and Z included.
    # Y and Z are recorded after the automation session is torn down, and Phase 5
    # needs that session live - so the demand was unsatisfiable by construction:
    # 33 of 35, P5-ALL refused, then Y and Z passed moments later and the summary
    # printed 35 of 35. The threshold was never the problem. The matrix is still
    # 35 and 35/35 is still required; that demand now lives in P5-FIN, which runs
    # after Y and Z and which no run can pass without them.
    $required = Get-Phase4RequiredScenarioIds
    $deferred = Get-Phase4FinalizationScenarioIds
    $prerequisite = Get-Phase4PrerequisiteScenarioIds
    $list = New-Checklist
    $seen = @($Results | ForEach-Object { $_.Id })

    # THE PARTITION IS PROVED, NOT ASSERTED IN PROSE. Nothing may leave the
    # matrix by being called a lifecycle case: the two sets must be disjoint,
    # must cover all 35, and every deferred name must be a real matrix member.
    $overlap = @($prerequisite | Where-Object { $deferred -contains $_ })
    $strayDeferred = @($deferred | Where-Object { $required -notcontains $_ })
    $partitionOk = (($prerequisite.Count + $deferred.Count) -eq $required.Count) -and `
                   ($overlap.Count -eq 0) -and ($strayDeferred.Count -eq 0)
    $null = Add-Check $list 'the prerequisite and deferred sets partition the 35-case matrix' `
        $partitionOk `
        ("prerequisite " + $prerequisite.Count + " + deferred " + $deferred.Count +
         " vs matrix " + $required.Count + "; overlap: " + ($overlap -join ', ') +
         "; not in matrix: " + ($strayDeferred -join ', '))

    # AND THE DEFERRAL IS REAL. If a deferred case has already been recorded, it
    # was never a post-session case and must not be excluded from this gate.
    # This is the direct regression assertion for the Run-1 topology.
    $earlyDeferred = @($deferred | Where-Object { $seen -contains $_ })
    $null = Add-Check $list `
        ('the deferred lifecycle cases have not run yet, so deferring them is real: ' +
         ($deferred -join ', ')) `
        ($earlyDeferred.Count -eq 0) ("already recorded: " + ($earlyDeferred -join ', '))

    $missing = @()
    foreach ($id in $prerequisite) { if ($seen -notcontains $id) { $missing += $id } }
    $null = Add-Check $list `
        ('all ' + $prerequisite.Count + ' pre-Phase-5 Phase-4 scenarios reported a result') `
        ($missing.Count -eq 0) ("missing: " + ($missing -join ', '))
    $phase4 = @($Results | Where-Object { $prerequisite -contains $_.Id })
    $failed = @($phase4 | Where-Object { $_.Status -eq 'FAIL' })
    $skipped = @($phase4 | Where-Object { $_.Status -eq 'SKIP' })
    # Not one FAIL and not one SKIP is tolerated, exactly as before. Only the
    # membership of the set narrowed, and only to what can exist at this point.
    $null = Add-Check $list 'the Phase-4 matrix has 0 FAIL' ($failed.Count -eq 0) `
        (($failed | ForEach-Object { $_.Id }) -join ', ')
    $null = Add-Check $list 'the Phase-4 matrix has 0 SKIP' ($skipped.Count -eq 0) `
        (($skipped | ForEach-Object { $_.Id }) -join ', ')
    $passed = @($phase4 | Where-Object { $_.Status -eq 'PASS' })
    $null = Add-Check $list `
        ('the pre-Phase-5 Phase-4 matrix is ' + $prerequisite.Count + '/' +
         $prerequisite.Count + ' PASS') `
        ($passed.Count -eq $prerequisite.Count) `
        ("passed " + $passed.Count + " of " + $prerequisite.Count)
    $phase4Ok = Test-ChecklistOk $list
    Add-Phase5Result 'P5-P4' `
        ('Phase-4 prerequisite before Phase 5: ' + $prerequisite.Count + '/' +
         $prerequisite.Count + ' PASS, 0 FAIL, 0 SKIP (' + ($deferred -join ', ') +
         ' deferred to P5-FIN)') `
        $(if ($phase4Ok) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    if (-not $phase4Ok) {
        # A FAIL, never a SKIP. "Phase 5 was not attempted" must be as loud as
        # "Phase 5 failed", or a broken prerequisite reads as a quiet clean run.
        Add-Phase5Result 'P5-ALL' 'Phase-5 Gate-B scenarios' 'FAIL' `
            'not attempted: the Phase-4 structural matrix is not intact, so no Phase-5 result would mean anything'
        return
    }

    # -------------------------------------------------------------------
    # P5-CMP. THE WHOLE PROJECT COMPILES. NOT "an entry point answered."
    # -------------------------------------------------------------------
    # RUNTIME RUN 7. In one session of real Excel:
    #
    #   A1     PASS   PCCM_AutomationBegin is callable
    #   P5-M   PASS   fifteen modules present, and six API procedures reported
    #                 callable - under the evidence model P5-M had at the time.
    #                 One of those six, PCCM_Calculate, had never crossed
    #                 Application.Run; that borrowed claim was removed in the
    #                 review of ae52bdd. P5-M now proves six DECLARED and five
    #                 CALLABLE.
    #   P5-FIX FAIL   PCCM_Calculate -> HRESULT 0x800A9C68, and the VBE reported
    #                 "Compile error: Sub or Function not defined" on the call
    #                 to Contribute inside modCalcAnalytical.AccumulateTotals
    #
    # Contribute was declared, once, in that module. Its declaration carried
    # `ByRef scale As Double`, and `Scale` is a VBA statement keyword, so the
    # procedure never came into existence and every call to it was an undefined
    # symbol. VBA COMPILES ON DEMAND: a project answers an API call while a
    # procedure body nothing has reached yet still holds a fatal declaration.
    #
    # So callability is not compilation, and this scenario exists to make the
    # stronger claim separately, once, BEFORE anything relies on it.
    #
    # RUNTIME RUN 8. This scenario's first real Windows execution, 39 passed /
    # 2 failed, the second failure being P5-ALL's dependency gate. It settled
    # one question and opened another:
    #
    #   PASS   the VBE object model is reachable
    #   PASS   the Stage-B workbook exposes its VBProject
    #   PASS   the VBE reports an active VBProject
    #   PASS   both VBProjects name a file
    #   PASS   the active VBE project IS the Stage-B workbook project, by full
    #          path - both VBAProject, both C:\...\PCCM_stageB.xlsm
    #   PASS   the two VBProject names are recorded for diagnosis
    #   PASS   the VBE command bars are reachable
    #   FAIL   the Compile VBAProject command (ID 578) exists
    #
    # SETTLED: on a real owned Excel instance the ACTIVE project WAS the PCCM
    # Stage-B project, and the whole target-project identity chain above is now
    # backed by Windows evidence rather than by argument. It is frozen.
    #
    # OPENED: the run stopped at command discovery, so THE VBA COMPILER WAS
    # NEVER INVOKED. Run 8 licenses no verdict at all about whether the
    # production project compiles - not a good one and not a bad one. What it
    # proves is narrower and entirely about this harness: the FindControl call
    # as it was written returned no control for ID 578.
    #
    # RUNTIME RUN 9, AND THE MANUAL COMPILE THAT SETTLED IT. With the explicit
    # four-argument lookup, discovery worked. Run 9 got all the way through:
    #
    #   target VBProject acquired, ActiveVBProject acquired, active == target
    #   by FileName, CommandBars reachable, ID 578 found, control.Id == 578,
    #   control.Type == 1, Enabled before Execute == True, Execute() reached
    #
    # and then failed on one thing only: Enabled read True again in the
    # statement straight after Execute.
    #
    # THE RETAINED ARTIFACT WAS THEN COMPILED BY HAND. The same
    # PCCM_stageB.xlsm Run 9 left behind was opened, Debug > Compile VBAProject
    # was enabled, and invoking it once produced NO compile error, NO undefined
    # symbol, and left the command greyed out. So:
    #
    #   THE PRODUCTION VBA PROJECT COMPILES ON THE REAL TARGET ENVIRONMENT.
    #   The Run-7 reserved-identifier defect class is closed.
    #
    # WHAT THAT DOES AND DOES NOT SETTLE. It does not say whether Run 9's own
    # programmatic Execute finished after the harness looked - the manual
    # compile happened in a different Excel session, on a reopened file. The
    # honest conclusion is narrower and is the one this scenario now acts on:
    # RUN 9'S IMMEDIATE POST-EXECUTE OBSERVATION IS INSUFFICIENT. A
    # CommandBarControl's Enabled is cached UI state; reading it one statement
    # after Execute measures the harness's timing, not the compiler's outcome.
    # Settlement has to be observed in the SAME session, by reacquiring the
    # control, which is what the bounded poll below does.
    #
    # THE MECHANISM. The VBE exposes Compile VBAProject as a command bar control,
    # and it is addressed BY ID (578), never by caption: the caption is localised
    # and an English-only lookup would silently find nothing on a non-English
    # Excel and report success for a project that was never compiled.
    #
    # WHAT AN ENABLED CONTROL MEANS. The VBE enables Compile VBAProject when
    # there is something left to compile and disables it once the project is
    # fully compiled. So `Enabled = False` BEFORE the call means the project was
    # already compiled, and `Enabled = False` AFTER a successful Execute is the
    # positive evidence that the compilation completed. A control that is still
    # enabled afterwards is reported as a FAIL, not explained away.
    #
    # FAIL CLOSED, AND KEEP THE EVIDENCE. Every step is inside try/catch and any
    # throw is a FAIL carrying Excel's own message. Nothing here suppresses,
    # auto-answers or dismisses a compile-error dialog: DisplayAlerts is left
    # exactly as the accepted lifecycle set it, because a dialog that was
    # dismissed is a diagnostic that was destroyed. If VBProject access is
    # denied by Trust Center policy the scenario says so and FAILS rather than
    # assuming the project is fine.
    try {
        $list = New-Checklist
        $vbe = $null; $bars = $null; $control = $null
        $targetProject = $null; $activeProject = $null
        $targetIsActive = $false; $identity = 'not read'
        try {
            $vbe = $Excel.VBE
            $null = Add-Check $list 'the VBE object model is reachable' ($null -ne $vbe)

            # --- WHICH PROJECT IS THIS COMMAND GOING TO COMPILE? -------------
            #
            # REVIEW OF d21e1d7. Command 578 is a VBE command and it acts on the
            # VBE's ACTIVE project, not on a project the caller names. Reading
            # Enabled and calling Execute without binding that to the workbook
            # under test proves only that SOME active project compiled.
            #
            # A fresh owned Excel instance is not a guarantee of one project: an
            # add-in, a startup workbook or PERSONAL.XLSB each carry their own
            # VBProject, and Gate B may not assume the right one is active.
            #
            # NAME IS NOT IDENTITY. "VBAProject" is the default name every
            # project gets, so two projects routinely share it. The Stage-B
            # workbook is saved to a concrete .xlsm path before this runs, so
            # FileName is available and is the identity that distinguishes them.
            if ($null -ne $vbe) {
                $targetProject = $Workbook.VBProject
                $null = Add-Check $list 'the Stage-B workbook exposes its VBProject' `
                    ($null -ne $targetProject)
                $activeProject = $vbe.ActiveVBProject
                $null = Add-Check $list 'the VBE reports an active VBProject' `
                    ($null -ne $activeProject)
            }
            if (($null -ne $targetProject) -and ($null -ne $activeProject)) {
                # PLAIN DATA, read before anything is released.
                $targetName = [string]$targetProject.Name
                $activeName = [string]$activeProject.Name
                $targetFile = [string]$targetProject.FileName
                $activeFile = [string]$activeProject.FileName
                $identity = ('target: ' + $targetName + ' <' + $targetFile + '>; ' +
                             'active: ' + $activeName + ' <' + $activeFile + '>')
                Add-Note ('P5-CMP: VBProject identity - ' + $identity)

                # A project that has never been saved has no FileName, and an
                # empty string would compare equal to another empty string. Both
                # sides must actually name a file.
                $haveFiles = (-not [string]::IsNullOrWhiteSpace($targetFile)) -and `
                             (-not [string]::IsNullOrWhiteSpace($activeFile))
                $null = Add-Check $list `
                    'both VBProjects name a file, so identity is comparable at all' `
                    $haveFiles $identity

                # WINDOWS FILESYSTEM EQUIVALENCE ONLY. GetFullPath normalises
                # separators and relative segments; the comparison is
                # case-insensitive because NTFS paths are. No display caption is
                # involved and no substring match is accepted.
                $sameFile = $false
                if ($haveFiles) {
                    $targetFull = [System.IO.Path]::GetFullPath($targetFile)
                    $activeFull = [System.IO.Path]::GetFullPath($activeFile)
                    $sameFile = [string]::Equals($targetFull, $activeFull,
                        [System.StringComparison]::OrdinalIgnoreCase)
                    $identity = ('target: ' + $targetName + ' <' + $targetFull + '>; ' +
                                 'active: ' + $activeName + ' <' + $activeFull + '>')
                }
                $targetIsActive = $haveFiles -and $sameFile
                $null = Add-Check $list `
                    'the active VBE project IS the Stage-B workbook project (by file path)' `
                    $targetIsActive $identity
                # The names are recorded as CONTEXT, never as the identity test.
                $null = Add-Check $list `
                    'the two VBProject names are recorded for diagnosis' `
                    ($targetName.Length -gt 0) `
                    ('target name ' + $targetName + ', active name ' + $activeName +
                     ' (names are context; the file path above is the identity)')
            }

            # --- AND ONLY NOW THE COMMAND ------------------------------------
            #
            # THE GATE IS ON $targetIsActive, ON BOTH BRANCHES. Neither reading
            # Enabled nor calling Execute may happen against a project this
            # scenario has not identified, and a Compile command that is already
            # disabled is not evidence either: a disabled command over somebody
            # else's project says nothing about this one.
            #
            # FAIL CLOSED RATHER THAN ACTIVATE. Making the target project active
            # through the VBIDE model is possible, but it is UI manipulation this
            # round has no runtime evidence for, and compiling the wrong project
            # and reporting PASS is the failure mode being corrected. So a
            # mismatch is reported precisely and fails.
            if (-not $targetIsActive) {
                Add-Note ('P5-CMP: the VBE active project is NOT the PCCM workbook ' +
                          'project, so the Compile VBAProject command was NOT executed. ' +
                          'Compiling whatever happened to be active would prove nothing ' +
                          'about the project under test. ' + $identity)
            } else {
                # THE TWO CONSTANTS THE LOOKUP AND THE VERIFICATION SHARE, named
                # once and above every branch that reads them. PowerShell has no
                # Office type library to take msoControlButton from, so its value
                # is written down with the name it has in that enumeration.
                $msoControlButton = 1                          # MsoControlType
                $missing = [System.Reflection.Missing]::Value  # a truly omitted argument
                if ($null -ne $vbe) {
                    $bars = $vbe.CommandBars
                    $null = Add-Check $list 'the VBE command bars are reachable' ($null -ne $bars)
                }
                if ($null -ne $bars) {
                    # BY ID, AND THE CALL IS SPELLED OUT --------------------
                    #
                    # RUNTIME RUN 8. This lookup was `FindControl($null, 578)`
                    # and it returned nothing on real Windows, so the compiler
                    # was never invoked at all. CommandBars.FindControl takes
                    # (Type, Id, Tag, Visible) and every one of them is
                    # OPTIONAL. VBA writes `FindControl(ID:=578)` and omits the
                    # rest; PowerShell has no named-argument syntax for a COM
                    # method, and passing $null positionally is NOT the same as
                    # omitting an argument - the leading hypothesis for Run 8 is
                    # that $null was marshalled as a real Type criterion that
                    # matched no control. That remains a HYPOTHESIS until a
                    # Windows runtime confirms it.
                    #
                    # So: state Type explicitly, state Id explicitly, and omit
                    # Tag and Visible with the sentinel that actually means
                    # "omitted".
                    $control = $bars.FindControl($msoControlButton, 578, $missing, $missing)
                    $null = Add-Check $list 'the Compile VBAProject command (ID 578) exists' `
                        ($null -ne $control)

                    # DIAGNOSTIC ONLY, AND ONLY WHEN THE LOOKUP CAME BACK EMPTY.
                    # FindControls answers a different question with the same
                    # criteria: how many matching controls the collection API
                    # can see. It is an Add-Note, never an Add-Check, so it
                    # cannot move the verdict, and it NEVER yields a control to
                    # execute. If it throws, the throw is recorded and dropped.
                    if ($null -eq $control) {
                        $probeControls = $null
                        try {
                            $probeControls = $bars.FindControls($msoControlButton, 578, $missing, $missing)
                            if ($null -eq $probeControls) {
                                Add-Note ('P5-CMP: diagnostic - FindControls also returned ' +
                                          'nothing for Type 1 / Id 578.')
                            } else {
                                Add-Note ('P5-CMP: diagnostic - FindControls reported ' +
                                          [string]$probeControls.Count +
                                          ' control(s) for Type 1 / Id 578.')
                            }
                        } catch {
                            Add-Note ('P5-CMP: diagnostic - FindControls itself failed: ' +
                                      $_.Exception.Message)
                        } finally {
                            if ($null -ne $probeControls) {
                                Release-Transient $probeControls 'CommandBarControls(probe)'
                                $probeControls = $null
                            }
                        }
                    }
                }

                # --- IS THIS THE CONTROL WE ASKED FOR? ----------------------
                #
                # A non-null return is weaker evidence than it looks: it says
                # the collection handed something back, not that the something
                # is Compile VBAProject. Ask the control what it is, and let
                # its own answer decide. Caption is deliberately not read: it
                # is localised, and an English caption is not a fact about a
                # non-English Excel.
                $controlProved = $false
                if ($null -ne $control) {
                    $controlId = [int]$control.Id
                    $controlType = [int]$control.Type
                    $idOk = ($controlId -eq 578)
                    $typeOk = ($controlType -eq $msoControlButton)
                    $null = Add-Check $list `
                        'the control returned IS command Id 578' `
                        $idOk ('Id ' + [string]$controlId)
                    $null = Add-Check $list `
                        'the control returned IS an msoControlButton (Type 1)' `
                        $typeOk ('Type ' + [string]$controlType)
                    $controlProved = ($idOk -and $typeOk)
                }
                if ($controlProved) {
                    $before = [bool]$control.Enabled
                    Add-Note ('P5-CMP: Compile VBAProject enabled before the attempt: ' +
                              [string]$before)
                    $executeCount = 0
                    if ($before) {
                        # There is something to compile, so compile it. A throw
                        # here is the compile failure and is reported as one.
                        # ONCE. Nothing below invokes the command again.
                        $null = $control.Execute()
                        $executeCount = 1
                    }
                    $null = Add-Check $list `
                        'Compile VBAProject was executed at most once' `
                        ($executeCount -le 1) ('executions ' + [string]$executeCount)

                    # --- SETTLEMENT, NOT THE CACHED HANDLE ------------------
                    #
                    # RUNTIME RUN 9. Everything above passed - the right project
                    # was active, the exact Id-578 msoControlButton was found,
                    # Enabled read True, and Execute() was reached - and then
                    # the gate read `$control.Enabled` on the SAME handle in the
                    # next statement and saw True. It called that a failure.
                    #
                    # It was not one. The same retained Stage-B artifact was
                    # later opened by hand and Debug > Compile VBAProject
                    # completed with no error and went grey, so the production
                    # project is compile-clean. What the immediate read actually
                    # measured is unknown: a control's Enabled is a cached UI
                    # state, and one statement after Execute the VBE has not
                    # necessarily refreshed it.
                    #
                    # SO DROP THE HANDLE AND ASK AGAIN. The stale control is
                    # released here rather than at the end, because its Enabled
                    # is exactly the value that may not be trusted.
                    Release-Transient $control 'CommandBarControl'
                    $control = $null

                    # A BOUNDED POLL. At most five seconds, ~100 ms apart, and
                    # it stops the moment the command goes quiet. Every
                    # iteration REACQUIRES the control through the same explicit
                    # criteria and re-proves its Id and Type, so a settled
                    # reading is never taken from something that drifted into
                    # the collection; every acquired handle is released before
                    # the next iteration. Execute is NOT called again.
                    $settled = $false
                    $observations = 0
                    $lastEnabled = $true
                    $settleError = ''
                    $settleIdentityHeld = $true
                    $settleStarted = Get-Date
                    $settleDeadline = $settleStarted.AddSeconds(5)
                    while ((-not $settled) -and ((Get-Date) -lt $settleDeadline)) {
                        Start-Sleep -Milliseconds 100
                        $poll = $null
                        try {
                            $poll = $bars.FindControl($msoControlButton, 578, $missing, $missing)
                            if ($null -eq $poll) {
                                $settleError = 'the Compile VBAProject control could not be reacquired'
                                break
                            }
                            $pollId = [int]$poll.Id
                            $pollType = [int]$poll.Type
                            if (($pollId -ne 578) -or ($pollType -ne $msoControlButton)) {
                                $settleIdentityHeld = $false
                                $settleError = ('reacquired control Id ' + [string]$pollId +
                                                ' Type ' + [string]$pollType)
                                break
                            }
                            $lastEnabled = [bool]$poll.Enabled
                            $observations = $observations + 1
                            if (-not $lastEnabled) { $settled = $true }
                        } catch {
                            $settleError = ('reacquisition threw: ' + $_.Exception.Message)
                            break
                        } finally {
                            if ($null -ne $poll) {
                                Release-Transient $poll 'CommandBarControl(settle)'
                                $poll = $null
                            }
                        }
                    }
                    $settleMs = [int]((Get-Date) - $settleStarted).TotalMilliseconds
                    Add-Note ('P5-CMP: settlement - ' + [string]$observations +
                              ' observation(s) over ' + [string]$settleMs +
                              ' ms; last Enabled ' + [string]$lastEnabled +
                              $(if ($settleError) { '; ' + $settleError } else { '' }))

                    # THE OBSERVATION HAS TO HAVE HAPPENED, AND HAVE BEEN OF THE
                    # RIGHT CONTROL. A poll that never reacquired anything, or
                    # reacquired the wrong thing, is not evidence of settlement
                    # and fails here rather than being read as one.
                    $null = Add-Check $list `
                        'the compiled state was read by reacquiring the exact Id-578 control' `
                        (($observations -gt 0) -and $settleIdentityHeld -and
                         ($settleError.Length -eq 0)) `
                        ([string]$observations + ' observation(s) over ' + [string]$settleMs +
                         ' ms; ' + $(if ($settleError) { $settleError } else { 'no error' }))

                    # THE POSITIVE EVIDENCE. Either the target project was
                    # already fully compiled, or it was compiled just now and
                    # the command went quiet within the window. Both readings
                    # are about the target project, because the identity gate
                    # above is what let this branch run at all, and about the
                    # Compile command, because every observation re-proved the
                    # control's own Id and Type.
                    #
                    # AND IF IT NEVER SETTLES, SAY THAT AND ONLY THAT. Run 9 is
                    # the reason the detail is worded the way it is: a command
                    # still enabled at the deadline is an observation that was
                    # not established. It is not a compiler diagnostic, and the
                    # manual compile of that very artifact is why.
                    $null = Add-Check $list `
                        'the target PCCM VBProject reached the VBE compiled state' `
                        $settled `
                        ('Compile VBAProject did not settle to the disabled/compiled ' +
                         'state within the bounded observation window: enabled before ' +
                         [string]$before + ', still enabled after ' + [string]$observations +
                         ' observation(s) over ' + [string]$settleMs + ' ms. This is a ' +
                         'settlement observation that was not established, not a ' +
                         'compiler diagnostic. ' + $identity)
                }
            }
        } finally {
            if ($null -ne $control)       { Release-Transient $control       'CommandBarControl'; $control       = $null }
            if ($null -ne $bars)          { Release-Transient $bars          'CommandBars';       $bars          = $null }
            if ($null -ne $activeProject) { Release-Transient $activeProject 'VBProject(active)'; $activeProject = $null }
            if ($null -ne $targetProject) { Release-Transient $targetProject 'VBProject(target)'; $targetProject = $null }
            if ($null -ne $vbe)           { Release-Transient $vbe           'VBE';               $vbe           = $null }
        }
        $compileOk = Test-ChecklistOk $list
        Add-Phase5Result 'P5-CMP' `
            'Whole VBA project compiled through the VBE Compile command (ID 578)' `
            $(if ($compileOk) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        if (-not $compileOk) {
            # A FAIL, never a SKIP, and it gates like P5-FX and P5-FIX do.
            # Nothing below can mean anything while the whole-project compile is
            # unestablished, and Run 7 is the demonstration: nineteen scenarios
            # would have reported one compile defect as their own predicates
            # failing. Returning here leaves the caller's shutdown, Z, Y, P5-LDG
            # and P5-FIN untouched.
            #
            # WHAT THIS FAILURE DOES NOT SAY. A P5-CMP FAIL means the
            # prerequisite was not established, and Run 8 is why that
            # distinction is written down: the gate failed at command discovery
            # with the compiler never invoked. The CHECKLIST says which link
            # broke; the result line may not guess.
            Add-Phase5Result 'P5-ALL' 'Phase-5 Gate-B scenarios' 'FAIL' `
                ('not attempted: the whole-project VBA compile prerequisite was not ' +
                 'established, so no scenario below would be testing the model it ' +
                 'claims to test. See the P5-CMP checklist for the exact reason: it ' +
                 'may be VBE access, target-project identity, command discovery, or ' +
                 'a compiler diagnostic, and only the last of those is a statement ' +
                 'about the production project')
            return
        }
    } catch {
        Add-Phase5Result 'P5-CMP' 'Whole VBA project compile gate' 'FAIL' (Format-Phase5Err $_)
        Add-Phase5Result 'P5-ALL' 'Phase-5 Gate-B scenarios' 'FAIL' `
            'not attempted: the whole-project compile gate could not be completed'
        return
    }

    # -------------------------------------------------------------------
    # THE LOCKED FX SEED, CAPTURED ONCE
    # -------------------------------------------------------------------
    # Here, and nowhere later: the Phase-4 matrix has just been proved intact, so
    # this is the last moment the workbook is guaranteed untouched by Phase-5.
    # Every fixture below restores row 1 of tblFXRates from this capture.
    try {
        $list = New-Checklist
        $seed = Save-Phase5LockedFxSeed -Workbook $Workbook -Inspection $Inspection
        $null = Add-Check $list 'the locked FX seed was captured from the real Stage-B workbook' `
            (($null -ne $seed.Currency) -and ($null -ne $seed.Rate)) `
            ("captured " + (Format-Phase5Typed $seed.Currency) + " / " +
             (Format-Phase5Typed $seed.Rate))
        # The capture is REPORTED, not asserted against a literal. If the built
        # seed is wrong, the analytical scenarios below must fail; repairing it
        # here would hide a real build defect.
        # The TYPE is reported too, and is deliberately NOT asserted: whether the
        # built rate is numeric is the production build's claim to make, and the
        # analytical scenarios are what test it.
        Add-Note ("P5-FX: locked FX seed captured as " + (Format-Phase5Typed $seed.Currency) +
                  " / " + (Format-Phase5Typed $seed.Rate) +
                  " from the untouched Stage-B workbook")

        # THE RESTORATION PATH IS EXERCISED HERE, BEFORE ANYTHING DEPENDS ON IT.
        #
        # Runtime Run 4 failed eleven scenarios on one line inside this exact
        # path - Set-Phase5TypedCell, reached through Reset-Phase5FxTable - and
        # every one of them reported it as its own failure, forty lines into a
        # scenario that had never started. The path is a PREREQUISITE, so it is
        # proved like one.
        #
        # This restores the seed to ITS OWN CAPTURED VALUE, so the workbook is
        # left exactly as the capture found it; the write is a round trip, not a
        # mutation. The read-back is the strict typed comparator, so a String
        # that came back as a Double, or a Double as a String, fails here rather
        # than being discovered as a wrong analytical answer later.
        Reset-Phase5FxTable -Workbook $Workbook -Inspection $Inspection -Seed $seed
        $fx = $Inspection.input_tables.fx_rates
        $restored = @(Get-Phase5TypedTableBody -Workbook $Workbook -SheetName $fx.sheet `
            -TableName $fx.table_name)
        $null = Add-Check $list 'the seed round-trips through the real restoration path' `
            ($restored.Count -ge 1) ("body rows after restore: " + $restored.Count)
        if ($restored.Count -ge 1) {
            $null = Add-Check $list `
                'the restored currency is the captured value AND the captured type' `
                (Test-Phase5ExactValue -Actual $restored[0][0] -Expected $seed.Currency) `
                ("restored " + (Format-Phase5Typed $restored[0][0]) + ", captured " +
                 (Format-Phase5Typed $seed.Currency))
            $null = Add-Check $list `
                'the restored rate is the captured value AND the captured type' `
                (Test-Phase5ExactValue -Actual $restored[0][1] -Expected $seed.Rate) `
                ("restored " + (Format-Phase5Typed $restored[0][1]) + ", captured " +
                 (Format-Phase5Typed $seed.Rate))
        }

        $fxOk = Test-ChecklistOk $list
        Add-Phase5Result 'P5-FX' `
            'Locked FX seed captured, and its typed restoration path proved, before any Phase-5 mutation' `
            $(if ($fxOk) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        if (-not $fxOk) {
            # A FAIL, never a SKIP, and the same shape as the Phase-4 prerequisite
            # refusal: every fixture below restores this seed, so dozens of
            # analytical results built on a broken restoration would be noise.
            # Returning here leaves the caller's shutdown and P5-FIN untouched,
            # so the lifecycle evidence is still produced.
            Add-Phase5Result 'P5-ALL' 'Phase-5 Gate-B scenarios' 'FAIL' `
                ('not attempted: the locked FX seed does not survive its own restoration ' +
                 'path, so every fixture below would start from an unrestored workbook')
            return
        }
    } catch {
        Add-Phase5Result 'P5-FX' 'Locked FX seed capture and typed restoration' 'FAIL' (Format-Phase5Err $_)
        Add-Phase5Result 'P5-ALL' 'Phase-5 Gate-B scenarios' 'FAIL' `
            'not attempted: the locked FX seed could not be captured or restored'
        return
    }

    # -------------------------------------------------------------------
    # P5-M. The persisted project: modules by NAME, buttons, API procedures
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        # BY NAME AND BY TYPE, in both directions. A count alone would pass a
        # project that had gained a stray module and lost a real one; a
        # name-only scan over every component fails on the sheet and
        # ThisWorkbook documents Excel creates, which is what Run 2 hit.
        $expected = @($Manifest.vba.modules | ForEach-Object { [string]$_.name })
        $null = Add-Check $list 'the manifest declares 15 production modules' `
            ($expected.Count -eq 15) ("declared " + $expected.Count)
        $components = @(Get-Phase5VbComponentInventory -Workbook $Workbook)
        Add-Phase5ModuleInventoryChecks -List $list -Components $components `
            -ExpectedModules $expected -ExpectedSheetCount (@($Manifest.sheets).Count) `
            -Label 'saved project'

        # --- exactly five buttons, and not one of them calls PCCM_Calculate ---
        $declared = @($Manifest.buttons)
        $null = Add-Check $list 'the manifest declares exactly five buttons' ($declared.Count -eq 5) `
            ("declared " + $declared.Count)
        # RUN-2 ROOT. This used to count every Shape on every manifest sheet and
        # require the total to be five:
        #
        #   FAIL exactly five command buttons persist in the workbook -- found 6
        #
        # while all five declared buttons were present with the right OnAction
        # and NO shape called PCCM_Calculate. A Shape is not a command button: a
        # command button is a shape BOUND TO A MACRO. The count said six and did
        # not say what the sixth was, so the run could not be diagnosed from its
        # own evidence.
        #
        # Every shape is still enumerated and still judged. What changed is that
        # the inventory is now taken by NAME and OnAction, the command-button
        # requirement is applied to the shapes that actually command something,
        # and any shape that is not one of the five declared buttons must carry
        # NO PCCM_ macro at all. That is strictly stronger than the count it
        # replaces, and it names what it found.
        $shapeRecords = @()
        foreach ($sheetSpec in @($Manifest.sheets)) {
            $sheets = $null; $sheet = $null; $shapes = $null
            try {
                $sheets = $Workbook.Worksheets
                $sheet = $sheets.Item($sheetSpec.name)
                $shapes = $sheet.Shapes
                for ($i = 1; $i -le $shapes.Count; $i++) {
                    $shape = $null
                    try {
                        $shape = $shapes.Item($i)
                        $shapeRecords += [pscustomobject]@{
                            Sheet    = [string]$sheetSpec.name
                            Name     = [string]$shape.Name
                            OnAction = [string]$shape.OnAction
                        }
                    } finally {
                        if ($null -ne $shape) { Release-Transient $shape 'Shape'; $shape = $null }
                    }
                }
            } finally {
                if ($null -ne $shapes) { Release-Transient $shapes 'Shapes'; $shapes = $null }
                if ($null -ne $sheet) { Release-Transient $sheet 'Worksheet'; $sheet = $null }
                if ($null -ne $sheets) { Release-Transient $sheets 'Worksheets'; $sheets = $null }
            }
        }
        $onActions = @($shapeRecords | ForEach-Object { [string]$_.OnAction })
        $shapeInventory = (@($shapeRecords | ForEach-Object {
            [string]$_.Sheet + '!' + [string]$_.Name + ' -> ' +
            $(if ([string]::IsNullOrWhiteSpace([string]$_.OnAction)) { '<no macro>' }
              else { [string]$_.OnAction })
        }) -join ', ')
        $declaredNames = @($declared | ForEach-Object { [string]$_.shape_name })
        $commandButtons = @($shapeRecords | Where-Object {
            -not [string]::IsNullOrWhiteSpace([string]$_.OnAction) })

        # THE PROOF IS ON TRIPLES: (Sheet, ShapeName, OnAction).
        #
        # Three independent global sets are not enough, and the counterexample is
        # not hypothetical:
        #
        #     btnPCCMAddCostLine    -> PCCM_DeleteCostLine
        #     btnPCCMDeleteCostLine -> PCCM_AddCostLine
        #
        # All five shape names still exist. All five entry points still appear
        # somewhere in the OnAction list. Five shapes are still bound. Nothing
        # calls PCCM_Calculate. Every set-wise check passes and two real buttons
        # do the opposite of what they say. The manifest already carries the
        # whole identity - sheet, shape_name, entry_point - so the binding is
        # matched as a unit, on the sheet the manifest names.
        foreach ($button in $declared) {
            $wantSheet = [string]$button.sheet
            $wantName = [string]$button.shape_name
            $wantAction = [string]$button.entry_point

            # 1. EXACTLY ONE shape with that name on the declared sheet.
            $onSheet = @($shapeRecords | Where-Object {
                ([string]$_.Sheet -eq $wantSheet) -and ([string]$_.Name -eq $wantName) })
            $null = Add-Check $list `
                ('exactly one shape named ' + $wantName + ' exists on ' + $wantSheet) `
                ($onSheet.Count -eq 1) `
                ('found ' + $onSheet.Count + '; all shapes: ' + $shapeInventory)

            # 2. THAT shape's OnAction is the declared entry point. Not "the
            #    entry point exists somewhere" - this shape, this macro.
            $bound = $false
            $actual = '<no such shape>'
            if ($onSheet.Count -eq 1) {
                $actual = [string]$onSheet[0].OnAction
                if ([string]::IsNullOrWhiteSpace($actual)) { $actual = '<no macro>' }
                $bound = ($actual -ceq $wantAction)
            }
            $null = Add-Check $list `
                ('the button ' + $wantSheet + '!' + $wantName + ' calls ' + $wantAction) `
                $bound ('OnAction is ' + $actual)

            # 3. NO SECOND COPY of the declared shape name anywhere else. A
            #    duplicate on another sheet is a second command surface.
            $elsewhere = @($shapeRecords | Where-Object {
                ([string]$_.Name -eq $wantName) -and ([string]$_.Sheet -ne $wantSheet) })
            $null = Add-Check $list `
                ('no second shape named ' + $wantName + ' exists on any other sheet') `
                ($elsewhere.Count -eq 0) `
                ((@($elsewhere | ForEach-Object {
                    [string]$_.Sheet + '!' + [string]$_.Name + ' -> ' + [string]$_.OnAction }) -join ', '))
        }

        # 4. NO UNDECLARED MACRO-BOUND SHAPE. Judged as a triple too: a bound
        #    shape whose (sheet, name) is not a declared pair.
        $declaredPairs = @($declared | ForEach-Object {
            [string]$_.sheet + [char]31 + [string]$_.shape_name })
        $undeclaredBound = @($commandButtons | Where-Object {
            $declaredPairs -notcontains ([string]$_.Sheet + [char]31 + [string]$_.Name) })
        $null = Add-Check $list 'every macro-bound shape is one of the five declared buttons' `
            ($undeclaredBound.Count -eq 0) `
            ("undeclared: " + (@($undeclaredBound | ForEach-Object {
                [string]$_.Sheet + '!' + [string]$_.Name + ' -> ' + [string]$_.OnAction }) -join ', '))

        # AND NO UNDECLARED SHAPE MAY REACH THE PCCM SURFACE AT ALL, bound or not.
        $strayPccm = @($shapeRecords | Where-Object {
            ($declaredPairs -notcontains ([string]$_.Sheet + [char]31 + [string]$_.Name)) -and
            ([string]$_.OnAction -like 'PCCM_*') })
        $null = Add-Check $list 'no undeclared shape invokes a PCCM_ procedure' `
            ($strayPccm.Count -eq 0) `
            ((@($strayPccm | ForEach-Object {
                [string]$_.Sheet + '!' + [string]$_.Name + ' -> ' + [string]$_.OnAction }) -join ', '))

        # 6. EXACTLY THE FIVE DECLARED BINDINGS EXIST. The bound triples and the
        #    declared triples must be the same set - which closes the count from
        #    both ends without ever counting raw shapes.
        $declaredTriples = @($declared | ForEach-Object {
            [string]$_.sheet + [char]31 + [string]$_.shape_name + [char]31 + [string]$_.entry_point })
        $boundTriples = @($commandButtons | ForEach-Object {
            [string]$_.Sheet + [char]31 + [string]$_.Name + [char]31 + [string]$_.OnAction })
        $missingTriples = @($declaredTriples | Where-Object { $boundTriples -notcontains $_ })
        $extraTriples = @($boundTriples | Where-Object { $declaredTriples -notcontains $_ })
        $null = Add-Check $list `
            'exactly the five declared (sheet, shape, macro) bindings exist' `
            (($missingTriples.Count -eq 0) -and ($extraTriples.Count -eq 0) -and
             ($boundTriples.Count -eq $declaredTriples.Count)) `
            ('missing: ' + (($missingTriples -replace [string][char]31, '!') -join ', ') +
             '; unexpected: ' + (($extraTriples -replace [string][char]31, '!') -join ', ') +
             '; bound ' + $boundTriples.Count + ' of ' + $declaredTriples.Count)

        Add-Note ('P5-M: shape inventory across the ' + [string]@($Manifest.sheets).Count +
                  ' manifest sheets: ' + $shapeInventory)
        # THE SET-WISE PER-BUTTON CHECK IS GONE, deliberately. It asked whether
        # each declared entry point appeared ANYWHERE in the global OnAction
        # list, which is true even when two buttons have swapped macros - it
        # would have printed "ok the button btnPCCMAddCostLine calls
        # PCCM_AddCostLine" about a button that calls PCCM_DeleteCostLine. The
        # per-triple check above answers the same question about the actual
        # shape, so keeping the weaker one would only add a reassuring line that
        # can be wrong.

        # THE ONE THAT MATTERS: no shape may invoke the calculation endpoint.
        $null = Add-Check $list 'NO shape has OnAction = PCCM_Calculate' `
            ($onActions -notcontains 'PCCM_Calculate') (($onActions -join ', '))

        # --- api_procedures, consumed as api_procedures ----------------------
        # Deliberately NOT folded into entry_points: an entry point is bound to a
        # button and an API procedure is not, and the manifest is where the
        # harness learns the difference.
        $api = @($Manifest.vba.api_procedures)
        $entry = @($Manifest.vba.entry_points)
        $null = Add-Check $list 'the manifest projects exactly six API procedures' ($api.Count -eq 6) `
            ("projected " + $api.Count)
        $null = Add-Check $list 'no API procedure is also an entry point' `
            (@($api | Where-Object { $entry -contains $_ }).Count -eq 0)
        $null = Add-Check $list 'no API procedure is bound to a button' `
            (@($api | Where-Object { $onActions -contains $_ }).Count -eq 0)
        # THREE KINDS OF EVIDENCE, AND THEY ARE NOT INTERCHANGEABLE.
        #
        #   A  DECLARATION   the name exists in the persisted VBA project
        #   B  CALLABILITY   Application.Run reached it and it answered
        #   C  EXECUTION     it ran against a valid fixture and did its work
        #
        # This block used to report B for all six API procedures. It did not
        # have B for all six: PCCM_Calculate was never invoked here, and the
        # branch that skipped it set $callable = $true anyway and emitted "the
        # API procedure PCCM_Calculate is callable" as a PASS. The name had not
        # crossed Application.Run and no COM callability had been observed. That
        # is an expected future exercise counted as a present proof, which is the
        # exact overclaim Run 7 exists to have taught us to stop making.
        #
        # PCCM_Calculate IS NOT INVOKED HERE TO MAKE THE OLD LABEL TRUE EITHER.
        # It is stateful: it resolves the model, writes the _Calc block and
        # publishes a status, and running it against whatever the workbook
        # happens to hold at inventory time would establish a snapshot no
        # scenario asked for. Its first execution belongs on a valid fixture,
        # which is P5-FIX; P5-AN then drives it across the analytical corpus.
        #
        # So each of the six gets the evidence it actually has.
        $declared = @(Get-Phase5ProjectProcedureNames -Workbook $Workbook)
        $null = Add-Check $list 'the persisted project could be read for declared procedures' `
            ($declared.Count -gt 0) ("declared procedures found: " + $declared.Count)
        foreach ($name in $api) {
            # A. DECLARED, for all six, read out of the persisted project's own
            #    code rather than assumed from the manifest that names them.
            $null = Add-Check $list ('the API procedure ' + $name + ' is declared in the persisted project') `
                ($declared -contains $name) `
                ("declared: " + ((@($declared | Where-Object { $_ -like 'PCCM_*' })) -join ', '))

            if ($name -eq 'PCCM_Calculate') {
                # B is DEFERRED, and saying so is the whole correction. No
                # $callable flag is set, no callability check is emitted, and
                # nothing here is recorded as a pass on its behalf.
                Add-Note ('P5-M: PCCM_Calculate is declared; it is stateful, so its ' +
                          'runtime execution is deferred to P5-FIX, which is the first ' +
                          'valid-fixture PCCM_Calculate of the run. P5-M records no ' +
                          'callability evidence for it.')
                continue
            }

            # B. CALLABLE, for the five read-only procedures, and callable means
            #    Application.Run actually returned.
            $callable = $false; $detail = ''
            try {
                $probe = $Excel.Run($name)
                $callable = $true
                $detail = "returned '" + [string]$probe + "'"
            } catch { $detail = (Format-Phase5Err $_) }
            $null = Add-Check $list ('the API procedure ' + $name + ' is callable') $callable $detail
        }
        # AND NONE OF IT IS COMPILATION. Run 7 passed every check in this
        # scenario and then met a VBE compile error inside the analytical path:
        # VBA compiles on demand, so a project answers a call while an unreached
        # procedure body still holds a declaration the parser rejects. The
        # whole-project claim belongs to P5-CMP alone.
        Add-Phase5Result 'P5-M' 'Persisted project: 15 modules by name, 5 buttons, 6 API procedures' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Phase5Result 'P5-M' 'Persisted project inventory' 'FAIL' (Format-Phase5Err $_)
    }

    # -------------------------------------------------------------------
    # P5-EV. No change events. The status cell is last-evaluated, not live.
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $project = $null; $components = $null
        try {
            $project = $Workbook.VBProject
            $components = $project.VBComponents
            $offenders = @()
            $declaredHandlers = @()
            for ($i = 1; $i -le $components.Count; $i++) {
                $component = $null; $module = $null
                try {
                    $component = $components.Item($i)
                    $module = $component.CodeModule
                    if ($module.CountOfLines -gt 0) {
                        $raw = [string]$module.Lines(1, $module.CountOfLines)
                        # COMMENTARY IS NOT CODE, AND NEITHER IS A STRING PAYLOAD.
                        # Run-2 P5-EV flagged prose; a message text naming a
                        # forbidden construct would have been the same defect.
                        $code = Get-VbaExecutableCode -Code $raw
                        foreach ($forbidden in @($Manifest.vba.forbidden_constructs)) {
                            if ($code -match [regex]::Escape([string]$forbidden)) {
                                $offenders += ([string]$component.Name + ': ' + [string]$forbidden)
                            }
                        }
                        # AND THE TWO EVENT HANDLERS ARE ALSO CHECKED AS
                        # DECLARATIONS, against the same stripped code. The
                        # general scan above would catch a bare mention; this
                        # names a real handler as a real handler, which is the
                        # distinction the requirement is actually about.
                        foreach ($handler in 'Worksheet_Change', 'Workbook_SheetChange') {
                            if (Test-VbaProcedureDeclared -Code $code -ProcedureName $handler) {
                                $declaredHandlers += ([string]$component.Name + ': ' + $handler)
                            }
                        }
                    }
                } finally {
                    if ($null -ne $module) { Release-Transient $module 'CodeModule'; $module = $null }
                    if ($null -ne $component) { Release-Transient $component 'VBComponent'; $component = $null }
                }
            }
            $null = Add-Check $list `
                'no forbidden construct exists in the EXECUTABLE code of the real Stage-B project' `
                ($offenders.Count -eq 0) ($offenders -join '; ')
            $null = Add-Check $list `
                'no change-event procedure is DECLARED anywhere in the project' `
                ($declaredHandlers.Count -eq 0) ($declaredHandlers -join '; ')
            foreach ($handler in 'Worksheet_Change', 'Workbook_SheetChange') {
                $null = Add-Check $list ('the manifest forbids ' + $handler) `
                    (@($Manifest.vba.forbidden_constructs) -contains $handler)
            }
        } finally {
            if ($null -ne $components) { Release-Transient $components 'VBComponents'; $components = $null }
            if ($null -ne $project) { Release-Transient $project 'VBProject'; $project = $null }
        }
        Add-Phase5Result 'P5-EV' 'No change events: the status cell is last-evaluated, not live' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Phase5Result 'P5-EV' 'No change events' 'FAIL' (Format-Phase5Err $_)
    }

    # ===================================================================
    # THE TRANSIENT DIAGNOSTIC SECTION
    #
    # Imported HERE and nowhere earlier. P5-CMP has already driven the VBE's
    # Compile VBAProject command over the PRODUCTION project, so the proof that
    # the accepted project compiles is complete and unmasked before a test-only
    # module exists in the VBA project at all.
    #
    # IT USED TO SAY A1 DID THAT. Scenario A1 makes the first Application.Run of
    # the run, which proves the automation surface answers and nothing more.
    # Runtime Run 7 passed A1, passed P5-M as it then stood, and then met
    # a VBE compile error inside the analytical path - VBA compiles on demand, so
    # a project answers a call while an unreached procedure body still holds a
    # declaration the parser rejects. P5-CMP is the whole-project authority now,
    # and it is the only one.
    # ===================================================================
    $diagnosticName = 'modPhase5GateBDiagnostics'
    $diagnosticImported = $false
    try {
        $list = New-Checklist
        $source = Join-Path $ScriptDir 'phase5_gate_b_diagnostics.bas'
        $null = Add-Check $list 'the diagnostic source exists' (Test-Path -LiteralPath $source) $source
        $null = Add-Check $list 'the diagnostic module is NOT declared in the manifest' `
            (@($Manifest.vba.modules | ForEach-Object { [string]$_.name }) -notcontains $diagnosticName)

        $project = $null; $components = $null; $imported = $null
        try {
            $project = $Workbook.VBProject
            $components = $project.VBComponents
            $imported = $components.Import($source)
            $diagnosticImported = $true
            $null = Add-Check $list 'the diagnostic module imported into the disposable project' `
                ([string]$imported.Name -eq $diagnosticName) ("imported as " + [string]$imported.Name)
        } finally {
            if ($null -ne $imported) { Release-Transient $imported 'VBComponent(diagnostic)'; $imported = $null }
            if ($null -ne $components) { Release-Transient $components 'VBComponents'; $components = $null }
            if ($null -ne $project) { Release-Transient $project 'VBProject'; $project = $null }
        }

        $ping = [string]$Excel.Run('GBD_Ping')
        $null = Add-Check $list 'the diagnostic module is callable' ($ping -eq ('OK|' + $diagnosticName)) $ping
        Add-Phase5Result 'P5-D0' 'Transient diagnostic module imported AFTER the P5-CMP whole-project compile' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Phase5Result 'P5-D0' 'Transient diagnostic module import' 'FAIL' (Format-Phase5Err $_)
    }

    if ($diagnosticImported) {
        # ---------------------------------------------------------------
        # P5-D1. Canonical numeric encoding: the ten locked vectors
        # ---------------------------------------------------------------
        try {
            $list = New-Checklist
            $vectors = @($Cases.fingerprint.numeric_encodings.vectors)
            $null = Add-Check $list 'ten locked numeric vectors were emitted' ($vectors.Count -eq 10) `
                ("vectors " + $vectors.Count)
            foreach ($vector in $vectors) {
                $wanted = [string]$vector.expected
                $reply = [string]$Excel.Run('GBD_CanonicalNumber', [double]$vector.value, '.')
                $null = Add-Check $list ("canonical number '" + [string]$vector.label + "'") `
                    ($reply -eq ('OK|' + $wanted)) ("got " + $reply + ", expected OK|" + $wanted)
                # THE TWO EXTREMES ARE ALSO BUILT ON TARGET. If a COM Double round
                # trip disturbed MAX_DOUBLE or the minimum subnormal, the two
                # answers differ and the report says which one moved - the vector
                # is never skipped and never quietly weakened.
                if (@('MAX_DOUBLE', 'minimum subnormal') -contains [string]$vector.label) {
                    $built = [string]$Excel.Run('GBD_CanonicalNumberConstructed', [string]$vector.label, '.')
                    $null = Add-Check $list `
                        ("canonical number '" + [string]$vector.label + "' built on target, not marshalled") `
                        ($built -eq ('OK|' + $wanted)) ("got " + $built + ", expected OK|" + $wanted)
                    $null = Add-Check $list `
                        ("the marshalled and on-target '" + [string]$vector.label + "' agree") `
                        ($built -eq $reply) ("marshalled " + $reply + " / constructed " + $built)
                }
            }
            Add-Phase5Result 'P5-D1' 'Direct VBA: ten canonical numeric encodings (plan case 26)' `
                $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        } catch {
            Add-Phase5Result 'P5-D1' 'Direct VBA: canonical numeric encodings' 'FAIL' (Format-Phase5Err $_)
        }

        # ---------------------------------------------------------------
        # P5-DP. Canonical-Double PARITY across the binary64 domain
        # ---------------------------------------------------------------
        # Runtime Run 2 showed ten vectors were enough to EXPOSE a defective
        # canonical encoder and nowhere near enough to accept a corrected one:
        # Format$ was wrong on six of the ten and right on four. This drives the
        # emitted parity corpus - every named boundary, both neighbours of every
        # power of ten, and deterministic generated bit patterns.
        #
        # THE DOUBLE IS REBUILT FROM ITS BIT PATTERN, not parsed from a decimal
        # literal, so no JSON reader's last-bit behaviour can enter the proof.
        # THE EXPECTED TEXT IS THE CORPUS'S. Nothing here computes it.
        try {
            $list = New-Checklist
            $parity = $Cases.fingerprint.canonical_parity
            $vectors = @($parity.vectors)
            $null = Add-Check $list 'the canonical parity corpus was emitted' `
                ($vectors.Count -gt 2000) ("vectors " + $vectors.Count)
            $checked = 0
            $failures = @()
            foreach ($vector in $vectors) {
                $bits = [string]$vector.bits
                $probe = [BitConverter]::Int64BitsToDouble([Convert]::ToInt64($bits, 16))
                $reply = [string]$Excel.Run('GBD_CanonicalNumber', [double]$probe, '.')
                $checked++
                if ($reply -cne ('OK|' + [string]$vector.expected)) {
                    if ($failures.Count -lt 20) {
                        $failures += ('[' + $bits + '] ' + [string]$vector.label +
                                      ': got ' + $reply + ', expected OK|' +
                                      [string]$vector.expected)
                    }
                }
            }
            # ONE aggregate check, because 2432 ok lines would bury the evidence,
            # and up to twenty full discrepancies, because a bare count would
            # bury it just as effectively the other way.
            $null = Add-Check $list `
                ('every canonical parity vector matched on real VBA (' +
                 [string]$checked + ' vectors)') `
                ($failures.Count -eq 0) `
                ('first ' + [string]$failures.Count + ' of the failures: ' +
                 ($failures -join ' | '))
            $null = Add-Check $list 'every emitted parity vector was actually driven' `
                ($checked -eq $vectors.Count) `
                ('drove ' + $checked + ' of ' + $vectors.Count)

            # NEIGHBOUR / COLLISION PROOF. Three distinct Doubles must produce
            # three distinct canonical strings, or two different models could
            # fingerprint alike. This is why the contract is 17 digits.
            foreach ($triple in @($parity.neighbours)) {
                $texts = @()
                $ok = $true
                foreach ($member in @($triple.members)) {
                    $probe = [BitConverter]::Int64BitsToDouble(
                        [Convert]::ToInt64([string]$member.bits, 16))
                    $reply = [string]$Excel.Run('GBD_CanonicalNumber', [double]$probe, '.')
                    if ($reply -cne ('OK|' + [string]$member.expected)) { $ok = $false }
                    $texts += $reply
                }
                $distinct = @($texts | Select-Object -Unique)
                $null = Add-Check $list `
                    ('neighbours of ' + [string]$triple.label +
                     ' stay three distinct canonical strings') `
                    ($ok -and ($distinct.Count -eq 3)) `
                    (($texts -join ' | '))
            }
            Add-Phase5Result 'P5-DP' `
                ('Canonical-Double parity across the binary64 domain (' +
                 [string]$vectors.Count + ' vectors, plan case 26)') `
                $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        } catch {
            Add-Phase5Result 'P5-DP' 'Canonical-Double parity' 'FAIL' (Format-Phase5Err $_)
        }

        # ---------------------------------------------------------------
        # P5-D2. Decimal-separator INJECTION, on the same Windows host
        # ---------------------------------------------------------------
        # The runtime proof Gate A could not make. Both separators go into the
        # SAME accepted encoder as its own argument, on one host, in one run. No
        # regional setting is read or altered; Application.International is never
        # touched and UseSystemSeparators is never set.
        try {
            $list = New-Checklist
            $vectors = @($Cases.fingerprint.decimal_separator.vectors)
            $null = Add-Check $list 'the separator vector set was emitted' ($vectors.Count -ge 10) `
                ("vectors " + $vectors.Count)
            foreach ($vector in $vectors) {
                foreach ($pair in @(
                    @{ separator = '.'; expected = [string]$vector.point;  name = 'point' },
                    @{ separator = ','; expected = [string]$vector.comma;  name = 'comma' })) {
                    $reply = [string]$Excel.Run('GBD_CanonicalNumber', [double]$vector.value, $pair.separator)
                    $null = Add-Check $list `
                        ("separator '" + $pair.name + "' on vector '" + [string]$vector.label + "'") `
                        ($reply -eq ('OK|' + $pair.expected)) `
                        ("got " + $reply + ", expected OK|" + $pair.expected)
                }
                if (@('MAX_DOUBLE', 'minimum subnormal') -contains [string]$vector.label) {
                    foreach ($separator in '.', ',') {
                        $built = [string]$Excel.Run('GBD_CanonicalNumberConstructed', [string]$vector.label, $separator)
                        $wanted = [string]$vector.point
                        if ($separator -eq ',') { $wanted = [string]$vector.comma }
                        $null = Add-Check $list `
                            ("separator '" + $separator + "' on the on-target '" + [string]$vector.label + "'") `
                            ($built -eq ('OK|' + $wanted)) ("got " + $built)
                    }
                }
            }
            # The output must be IDENTICAL under both separators: the canonical
            # form is the model's, not the host's.
            $null = Add-Check $list 'the canonical form does not depend on the injected separator' `
                (@($vectors | Where-Object { [string]$_.point -cne [string]$_.comma }).Count -eq 0)
            Add-Phase5Result 'P5-D2' 'Direct VBA: decimal-separator injection, both separators (plan case 35)' `
                $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        } catch {
            Add-Phase5Result 'P5-D2' 'Direct VBA: decimal-separator injection' 'FAIL' (Format-Phase5Err $_)
        }

        # ---------------------------------------------------------------
        # P5-D3. The Double-only reducer: all four locked vectors
        # ---------------------------------------------------------------
        # PowerShell computes NOTHING here. It hands h, u and the modulus to the
        # accepted VBA reducer and compares the returned remainder against the
        # fixture. A PowerShell-side reduction would be testing PowerShell.
        try {
            $list = New-Checklist
            $vectors = @($Cases.fingerprint.reduction_vectors)
            $null = Add-Check $list 'all four locked reduction vectors were emitted' `
                ($vectors.Count -eq 4) ("vectors " + $vectors.Count)
            foreach ($vector in $vectors) {
                $reply = [string]$Excel.Run('GBD_ReduceDouble', [double]$vector.h, [double]$vector.u,
                                            [double]$vector.modulus)
                # The remainder comes back as canonical text, so the comparison is
                # exact and no PowerShell number formatting stands in the way.
                $wantedReply = [string]$Excel.Run('GBD_CanonicalNumber', [double]$vector.remainder, '.')
                $null = Add-Check $list `
                    ("reduction h=" + [string]$vector.h + " u=" + [string]$vector.u + " mod " + [string]$vector.modulus_name) `
                    ($reply -eq $wantedReply) ("got " + $reply + ", expected " + $wantedReply)
                $null = Add-Check $list `
                    ("the fixture's double-only remainder equals its exact remainder for " + [string]$vector.modulus_name) `
                    ([double]$vector.double_only_remainder -eq [double]$vector.remainder)
            }
            Add-Phase5Result 'P5-D3' 'Direct VBA: the four Double-only reductions (plan case 36)' `
                $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        } catch {
            Add-Phase5Result 'P5-D3' 'Direct VBA: the Double-only reducer' 'FAIL' (Format-Phase5Err $_)
        }

        # ---------------------------------------------------------------
        # P5-D4. UTF-16: signed AscW, unit counting, surrogates, prefixes
        # ---------------------------------------------------------------
        try {
            $list = New-Checklist
            $vectors = @($Cases.fingerprint.utf16_vectors.vectors)
            foreach ($vector in $vectors) {
                $units = (@($vector.code_units) -join ',')
                $key = [string]$vector.key

                # LENGTH IS IN CODE UNITS. A non-BMP character contributes two.
                $reply = [string]$Excel.Run('GBD_Utf16Length', $units)
                $null = Add-Check $list ($key + ': UTF-16 length counts code units') `
                    ($reply -eq ('OK|' + [string]$vector.utf16_length)) `
                    ("got " + $reply + ", expected OK|" + [string]$vector.utf16_length)
                if ($key -eq 'non_bmp') {
                    $null = Add-Check $list `
                        'the non-BMP character contributes TWO surrogate units, not one' `
                        (([int]$vector.utf16_length -eq 2) -and ([int]$vector.code_point_count -eq 1))
                }

                # AscW IS SIGNED. Every unit above U+7FFF must come back negative,
                # and the accepted normaliser must put it back.
                $position = 0
                foreach ($signed in @($vector.signed_ascw)) {
                    $position++
                    $raw = [string]$Excel.Run('GBD_RawAscW', $units, [long]$position)
                    $null = Add-Check $list ($key + ': raw AscW at position ' + $position + ' is signed') `
                        ($raw -eq ('OK|' + [string]$signed)) ("got " + $raw + ", expected OK|" + [string]$signed)
                    $normalised = [string]$Excel.Run('GBD_NormaliseCodeUnit', [long]$signed)
                    $wanted = [string]@($vector.code_units)[$position - 1]
                    $null = Add-Check $list ($key + ': the normaliser restores unit ' + $position) `
                        ($normalised -eq ('OK|' + $wanted)) ("got " + $normalised + ", expected OK|" + $wanted)
                }
                $null = Add-Check $list ($key + ': at least one unit above U+7FFF is exercised') `
                    (@($vector.signed_ascw | Where-Object { [int]$_ -lt 0 }).Count -ge 1) -Detail `
                    (@($vector.signed_ascw) -join ',')

                # THE COMPLETE CANONICAL FIELD, not just its prefix.
                #
                # The corpus emits `canonical_text_field` for every vector, so the
                # whole framed field is compared. Checking only "S<units>:" would
                # pass a field whose PAYLOAD was wrong - a mangled surrogate pair
                # with the right length is exactly the failure this vector exists
                # to catch.
                $field = [string]$Excel.Run('GBD_CanonicalTextField', $units)
                $expectedField = 'OK|' + [string]$vector.canonical_text_field
                $null = Add-Check $list ($key + ': the COMPLETE canonical text field matches the emitted one') `
                    ($field -ceq $expectedField) `
                    ("got " + $field + ", expected " + $expectedField)
                $null = Add-Check $list ($key + ': its length prefix is the UTF-16 unit count') `
                    ($field.StartsWith('OK|S' + [string]$vector.utf16_length + ':')) $field
            }
            Add-Phase5Result 'P5-D4' 'Direct VBA: UTF-16 signed AscW, unit counting and prefixes (plan case 26)' `
                $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        } catch {
            Add-Phase5Result 'P5-D4' 'Direct VBA: UTF-16 behaviour' 'FAIL' (Format-Phase5Err $_)
        }

        # ---------------------------------------------------------------
        # P5-D5. The complete locked reference stream
        # ---------------------------------------------------------------
        try {
            $list = New-Checklist
            $reference = $Cases.fingerprint.reference
            $stream = [string]$reference.stream

            # BOTH, and the count FIRST. A digest asserted on its own would agree
            # with itself over a stream that arrived truncated.
            $length = [string]$Excel.Run('GBD_StreamLength', $stream)
            $null = Add-Check $list 'the reference stream is the emitted code-unit count on real VBA' `
                ($length -eq ('OK|' + [string]$reference.code_units)) `
                ("got " + $length + ", expected OK|" + [string]$reference.code_units)
            $digest = [string]$Excel.Run('GBD_DigestStream', $stream)
            $null = Add-Check $list 'the reference digest matches the emitted digest on real VBA' `
                ($digest -eq ('OK|' + [string]$reference.digest)) `
                ("got " + $digest + ", expected OK|" + [string]$reference.digest)
            Add-Phase5Result 'P5-D5' 'Direct VBA: the complete reference stream, units and digest (plan case 26)' `
                $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        } catch {
            Add-Phase5Result 'P5-D5' 'Direct VBA: the reference stream' 'FAIL' (Format-Phase5Err $_)
        }

        # ---------------------------------------------------------------
        # P5-D6. Delimiter-hostile field content
        # ---------------------------------------------------------------
        # The probe values carry ':', NUL, the unit separator and a newline. Every
        # one is handed over as CODE UNITS, so nothing about them survives or dies
        # by the console encoding on the way in.
        try {
            $list = New-Checklist
            $probes = @($Cases.fingerprint.collision_probes)
            $null = Add-Check $list 'the collision probes were emitted' ($probes.Count -ge 8) `
                ("probes " + $probes.Count)
            $seenDigests = @()
            foreach ($probe in $probes) {
                $encoded = @()
                foreach ($value in @($probe.values)) {
                    $units = @()
                    foreach ($character in [char[]][string]$value) { $units += [int][char]$character }
                    $encoded += ($units -join ',')
                }
                $reply = [string]$Excel.Run('GBD_ProbeDigest', ($encoded -join ';'))
                $null = Add-Check $list ('probe digest for ' + ($encoded -join ' | ')) `
                    ($reply -eq ('OK|' + [string]$probe.digest)) `
                    ("got " + $reply + ", expected OK|" + [string]$probe.digest)
                $seenDigests += [string]$probe.digest
            }
            # The point of the probes: hostile content must not COLLIDE.
            $null = Add-Check $list 'every probe digest is distinct' `
                ((@($seenDigests | Select-Object -Unique)).Count -eq $seenDigests.Count)
            Add-Phase5Result 'P5-D6' 'Direct VBA: delimiter-hostile field content (plan case 27)' `
                $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        } catch {
            Add-Phase5Result 'P5-D6' 'Direct VBA: delimiter-hostile field content' 'FAIL' (Format-Phase5Err $_)
        }

        # ---------------------------------------------------------------
        # P5-D7. A naive overflow with a representable result
        # ---------------------------------------------------------------
        try {
            $list = New-Checklist
            $case = $null
            foreach ($candidate in @($Cases.plan_cases)) {
                if ([string]$candidate.id -eq '28') { $case = $candidate }
            }
            $null = Add-Check $list 'plan case 28 was emitted with its statistics vectors' `
                ($null -ne $case -and @($case.statistics).Count -ge 3)
            foreach ($vector in @($case.statistics)) {
                $points = @($vector.points)
                $third = 0.0
                if ($points.Count -ge 3) { $third = [double]$points[2] }
                $reply = [string]$Excel.Run('GBD_ConvexStatistic', [string]$vector.statistic,
                                            [double]$points[0], [double]$points[1], [double]$third)
                $wanted = [string]$Excel.Run('GBD_CanonicalNumber', [double]$vector.expected, '.')
                $null = Add-Check $list `
                    ([string]$vector.statistic + ' survives the naive sum overflow') `
                    ($reply -eq $wanted) ("got " + $reply + ", expected " + $wanted)
            }
            Add-Phase5Result 'P5-D7' 'Direct VBA: convex statistics at the overflow boundary (plan case 28)' `
                $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        } catch {
            Add-Phase5Result 'P5-D7' 'Direct VBA: convex statistics at the overflow boundary' 'FAIL' (Format-Phase5Err $_)
        }

        # ---------------------------------------------------------------
        # P5-DC. THE ACCEPTED CHECKER, CALLED DIRECTLY ON REAL VBA
        # ---------------------------------------------------------------
        # Some plan section 18 predicates are unreachable through the workbook
        # because an EARLIER accepted gate refuses first. Base Year after Start
        # Year is the clear case: modTimeline PREVALIDATES it and refuses the
        # Apply, so the workbook is left with entered <> applied and the next
        # PCCM_Calculate is refused by StructuralPrerequisites with STRUCTURE
        # CHANGE PENDING - a different predicate, in a different module, with a
        # different message. A workbook mutation claiming the modCalcCheck
        # predicate would have been exercising the Phase-4 gate instead.
        #
        # So the checker is called DIRECTLY, through the transient diagnostic
        # module, over a ResolvedModel built on target. Both the type and
        # modCalcCheck.CheckResolvedModel are already Public: nothing is reopened.
        try {
            $list = New-Checklist
            $directs = @($Cases.gate_b.direct_check_cases)
            $null = Add-Check $list 'the emitted direct-check matrix is present' `
                ($directs.Count -ge 1) ("cases " + $directs.Count)
            foreach ($entry in $directs) {
                $id = [string]$entry.id
                $arguments = $entry.arguments
                $reply = [string]$Excel.Run([string]$entry.procedure,
                    [long]$arguments.base_year, [long]$arguments.start_year,
                    [long]$arguments.duration, [double]$arguments.discount_rate)
                $null = Add-Check $list `
                    ($id + ' (' + [string]$entry.predicate + '): modCalcCheck.CheckResolvedModel REFUSED the model') `
                    ($reply -like 'OK|*') ("got " + $reply)
                Add-Phase5DetailTokenChecks -List $list `
                    -Detail ($reply -replace '^OK\|', '') -Tokens $entry.detail_tokens -Label $id

                # THE CONTROL. The same construction with the predicate NOT
                # violated must be ACCEPTED, or the refusal proves only that the
                # harness built a model the checker rejects for some other reason.
                $control = $entry.control_arguments
                $accepted = [string]$Excel.Run([string]$entry.control_procedure,
                    [long]$control.base_year, [long]$control.start_year,
                    [long]$control.duration, [double]$control.discount_rate)
                $null = Add-Check $list `
                    ($id + ': the control model with the predicate satisfied is ACCEPTED') `
                    ($accepted -like 'OK|*') ("got " + $accepted)
            }
            Add-Phase5Result 'P5-DC' `
                'Direct VBA: plan section 18 predicates the workbook cannot reach, through modCalcCheck' `
                $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        } catch {
            Add-Phase5Result 'P5-DC' 'Direct VBA: modCalcCheck predicates' 'FAIL' (Format-Phase5Err $_)
        }
    } else {
        foreach ($id in 'P5-D1', 'P5-D2', 'P5-D3', 'P5-D4', 'P5-D5', 'P5-D6', 'P5-D7',
                        'P5-DC') {
            Add-Phase5Result $id 'Direct VBA diagnostic vector' 'FAIL' `
                'the transient diagnostic module did not import, so no locked vector was exercised on real VBA'
        }
    }

    # -------------------------------------------------------------------
    # P5-D8. THE DIAGNOSTIC MODULE IS REMOVED AGAIN
    # -------------------------------------------------------------------
    # Evidence infrastructure, not product. The inventory must return to exactly
    # the 15 manifest modules BEFORE anything else is asserted about the project,
    # and no accepted workbook is ever saved with it installed.
    try {
        $list = New-Checklist
        $project = $null; $components = $null
        try {
            $project = $Workbook.VBProject
            $components = $project.VBComponents
            $target = $null
            try { $target = $components.Item($diagnosticName) } catch { $target = $null }
            if ($null -ne $target) {
                $components.Remove($target)
                Release-Transient $target 'VBComponent(diagnostic)'; $target = $null
            }
        } finally {
            if ($null -ne $components) { Release-Transient $components 'VBComponents'; $components = $null }
            if ($null -ne $project) { Release-Transient $project 'VBProject'; $project = $null }
        }

        # RUN-2 ROOT, same defect as P5-M: this compared every component name
        # against the manifest and reported "present 30 of 15" while the
        # diagnostic module was genuinely gone and all fifteen production
        # modules were genuinely there. The inventory is judged by TYPE now, by
        # the SAME helper P5-M uses, so the two cannot drift apart.
        #
        # The diagnostic module is a STANDARD module, so its absence is proved
        # against the standard-module partition - the partition it would have to
        # reappear in. A document component cannot mask it and cannot be
        # mistaken for it.
        $inventory = @(Get-Phase5VbComponentInventory -Workbook $Workbook)
        $standardNames = @($inventory |
            Where-Object { [int]$_.Type -eq $script:VbextComponentTypes.StdModule } |
            ForEach-Object { [string]$_.Name })
        $null = Add-Check $list 'the diagnostic module is absent from the standard modules' `
            ($standardNames -notcontains $diagnosticName) ($standardNames -join ', ')
        $null = Add-Check $list 'the diagnostic module is absent from the project entirely' `
            (@($inventory | Where-Object { [string]$_.Name -eq $diagnosticName }).Count -eq 0) `
            (Format-VbComponentList $inventory)
        $expected = @($Manifest.vba.modules | ForEach-Object { [string]$_.name })
        Add-Phase5ModuleInventoryChecks -List $list -Components $inventory `
            -ExpectedModules $expected -ExpectedSheetCount (@($Manifest.sheets).Count) `
            -Label 'after removal'
        # It must also be gone from the RUNTIME: a removed component whose
        # procedure still answers would mean the removal did not take.
        $stillCallable = $false
        try { $null = $Excel.Run('GBD_Ping'); $stillCallable = $true } catch { $stillCallable = $false }
        $null = Add-Check $list 'no diagnostic procedure is callable any more' (-not $stillCallable)
        Add-Phase5Result 'P5-D8' 'Transient diagnostic module removed; inventory back to 15' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Phase5Result 'P5-D8' 'Transient diagnostic module removal' 'FAIL' (Format-Phase5Err $_)
    }

    # -------------------------------------------------------------------
    # P5-FIX. THE FIXTURE PROVES ITSELF BEFORE ANY SCENARIO DEPENDS ON IT
    # -------------------------------------------------------------------
    # RUN-5 SEQUENCING LESSON, the same one P5-FX carries for the FX seed: a
    # prerequisite is proved like a prerequisite, once, at the point where it can
    # still be reported as itself.
    #
    # Run 5 discovered a fixture-establishment ordering defect only through the
    # scenarios that inherited it. Each of them reported the same sentence about
    # an orphan row in tblCostLines, none of them had reached the predicate it
    # exists to test, and the thing that was actually broken - the order in which
    # the fixture performs its production mutations - had no result of its own
    # anywhere in the ledger.
    #
    # This runs ONE fixture, the golden plan case, and asserts what fixture
    # establishment is supposed to have achieved: production keyed every driver
    # row, the identifiers are the emitted ones, the inflation grid matches the
    # Config master, and the workbook production hands back is structurally
    # coherent. A failure here is a HARNESS failure and says so.
    try {
        $list = New-Checklist
        $golden = $null
        foreach ($candidate in @($Cases.plan_cases)) {
            if ([string]$candidate.id -eq '1') { $golden = $candidate }
        }
        $null = Add-Check $list 'the golden fixture (plan case 1) was emitted' ($null -ne $golden)

        # THE REAL Set-Phase5Fixture, not a reimplementation of it. A self-proof
        # that exercised a copy would prove the copy.
        $applied = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
            -Inspection $Inspection -Model $golden.model
        $null = Add-Check $list 'fixture establishment reported OK from PCCM_ApplyTimeline' `
            ($applied -like 'OK|*') $applied

        # 1. THE STRUCTURE PRODUCTION HANDED BACK IS COHERENT, read again here
        #    rather than trusted from inside the fixture.
        $report = [string]$Excel.Run('PCCM_StructuralReport')
        $null = Add-Check $list 'PCCM_StructuralReport is blank after fixture establishment' `
            ([string]::IsNullOrWhiteSpace($report)) $report

        # 2. EVERY DRIVER ROW IS KEYED, WITH THE IDENTIFIER THE CORPUS NAMES.
        #    This is the exact condition Run 5 violated: row 1 of tblCostLines
        #    held data and carried no key.
        foreach ($pair in @(
            @{ key = 'cost_lines';    drivers = @($golden.model.cost_lines) },
            @{ key = 'risk_register'; drivers = @($golden.model.risks) })) {
            $register = $null
            foreach ($candidate in @($Manifest.registers)) {
                if ($candidate.key -eq $pair.key) { $register = $candidate }
            }
            $expected = @($pair.drivers | ForEach-Object { [string]$_.permanent_id })
            $ids = @(Get-IdColumnValues -Workbook $Workbook -Info $register)
            $null = Add-Check $list `
                ([string]$register.table_name + ' carries exactly the emitted identifiers, in order') `
                (($ids.Count -eq $expected.Count) -and `
                 ((($ids | ForEach-Object { [string]$_ }) -join '|') -ceq ($expected -join '|'))) `
                ("issued: " + ($ids -join ', ') + "; emitted: " + ($expected -join ', '))
            # AND NO ROW HOLDS DATA WITHOUT ONE. The harness-side mirror of
            # production's own orphan predicate, asserted rather than thrown, so
            # the checklist carries the evidence.
            $orphans = @()
            $body = @(Get-TableBody -Workbook $Workbook -SheetName $register.sheet `
                -TableName $register.table_name)
            for ($row = 0; $row -lt $body.Count; $row++) {
                $cells = @($body[$row])
                if (-not [string]::IsNullOrWhiteSpace([string]$cells[0])) { continue }
                for ($column = 1; $column -lt $cells.Count; $column++) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$cells[$column])) {
                        $orphans += [string]($row + 1)
                        break
                    }
                }
            }
            $null = Add-Check $list `
                ('no row of ' + [string]$register.table_name + ' holds data without a key') `
                ($orphans.Count -eq 0) ("unkeyed data in row(s): " + ($orphans -join ', '))
        }

        # 3. THE CONFIG MASTER AND THE INFLATION GRID AGREE. Their disagreement
        #    is the state step D creates and step E closes, and it is what made
        #    every Run-5 Add refuse. If the order ever regresses, this is the
        #    check that names the reason rather than the symptom.
        $master = $Inspection.input_tables.inflation_profiles
        $declared = @()
        foreach ($row in @(Get-TableBody -Workbook $Workbook -SheetName $master.sheet `
                -TableName $master.table_name)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$row[0])) { $declared += [string]$row[0] }
        }
        $grid = $null
        foreach ($candidate in @($Manifest.grids)) { if ($candidate.key -eq 'inflation') { $grid = $candidate } }
        $gridRows = @()
        foreach ($row in @(Get-TableBody -Workbook $Workbook -SheetName $grid.sheet `
                -TableName $grid.table_name)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$row[0])) { $gridRows += [string]$row[0] }
        }
        $emittedProfiles = @($golden.model.inflation.PSObject.Properties.Name |
            ForEach-Object { [string]$_ })
        $null = Add-Check $list 'the Config profile master holds exactly the emitted profiles' `
            ((@($declared | Sort-Object) -join '|') -ceq (@($emittedProfiles | Sort-Object) -join '|')) `
            ("master: " + ($declared -join ', ') + "; emitted: " + ($emittedProfiles -join ', '))
        $null = Add-Check $list `
            'SyncProfileRows rebuilt tblInflation to agree with the master, so no Add can be refused for it' `
            ((@($gridRows | Sort-Object) -join '|') -ceq (@($declared | Sort-Object) -join '|')) `
            ("grid: " + ($gridRows -join ', ') + "; master: " + ($declared -join ', '))

        # 4. AND THE BASELINE IS USABLE. A fixture that establishes a coherent
        #    but uncalculable workbook has not established anything.
        #
        #    THIS CHECK DOES NOT DIAGNOSE. Checks 1 to 3 are claims about what
        #    the harness did; this one can fail for a production reason instead,
        #    and the gate below is worded so that it reports what was observed
        #    rather than deciding whose fault it is. The attempt detail is
        #    carried into the checklist so the distinction can be made from the
        #    evidence.
        $Excel.Run('PCCM_Calculate') | Out-Null
        $null = Add-Check $list 'the self-proof fixture calculates successfully' `
            ([string]$Excel.Run('PCCM_CalculationAttemptResult') -eq 'SUCCESS') `
            ([string]$Excel.Run('PCCM_CalculationAttemptDetail'))

        $fixtureOk = Test-ChecklistOk $list
        Add-Phase5Result 'P5-FIX' `
            'Fixture establishment proves itself: production keyed every driver row before any data was written' `
            $(if ($fixtureOk) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        if (-not $fixtureOk) {
            # A FAIL, never a SKIP, and the same shape as P5-FX. Every scenario
            # below establishes a fixture; if establishment is broken, their
            # results would all be restatements of this one failure attributed to
            # predicates that were never reached. Returning here leaves the
            # caller's shutdown, Y, Z, P5-LDG and P5-FIN untouched, so the
            # lifecycle evidence is still produced.
            Add-Phase5Result 'P5-ALL' 'Phase-5 Gate-B scenarios' 'FAIL' `
                ('not attempted: the fixture self-proof did not hold, so the baseline ' +
                 'every scenario below builds on is unproven and each of them would ' +
                 'report this one failure as its own predicate failing. The P5-FIX ' +
                 'checklist carries which claim failed: checks 1 to 3 are claims about ' +
                 'the harness, check 4 can fail for a production reason instead')
            return
        }
    } catch {
        Add-Phase5Result 'P5-FIX' 'Fixture establishment self-proof' 'FAIL' (Format-Phase5Err $_)
        Add-Phase5Result 'P5-ALL' 'Phase-5 Gate-B scenarios' 'FAIL' `
            'not attempted: Gate-B fixture establishment failed on the golden plan case'
        return
    }

    # -------------------------------------------------------------------
    # P5-AN. Every analytical fixture, every emitted expected value
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $covered = @()
        foreach ($case in @($Cases.plan_cases)) {
            if ([string]$case.kind -ne 'analytical') { continue }
            $id = [string]$case.id
            $applied = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
                -Inspection $Inspection -Model $case.model
            $null = Add-Check $list ('case ' + $id + ': the fixture applied its timeline') `
                ($applied -like 'OK|*') $applied

            $Excel.Run('PCCM_Calculate') | Out-Null
            $attempt = [string]$Excel.Run('PCCM_CalculationAttemptResult')
            $detail = [string]$Excel.Run('PCCM_CalculationAttemptDetail')
            $status = [string]$Excel.Run('PCCM_CalculationStatus')
            $null = Add-Check $list ('case ' + $id + ': the attempt is SUCCESS') `
                ($attempt -eq 'SUCCESS') ("attempt '" + $attempt + "', detail '" + $detail + "'")
            $null = Add-Check $list ('case ' + $id + ': the attempt detail is blank on success') `
                ([string]::IsNullOrEmpty($detail)) $detail
            $null = Add-Check $list ('case ' + $id + ': the derived status is CURRENT') `
                ($status -eq 'CURRENT') $status
            $fingerprint = [string]$Excel.Run('PCCM_CalculationFingerprint')
            $current = [string]$Excel.Run('PCCM_CurrentInputFingerprint')
            $null = Add-Check $list ('case ' + $id + ': a stored fingerprint exists') `
                (-not [string]::IsNullOrEmpty($fingerprint))
            $null = Add-Check $list ('case ' + $id + ': the stored and current fingerprints agree') `
                ($fingerprint -ceq $current) ("stored " + $fingerprint + ", current " + $current)

            # THE GOLDEN END-TO-END PARITY. Case 1 IS the model the emitted
            # reference stream was built from, so the COMPLETE production path -
            # resolution, referenced factors, record construction, canonical
            # section ordering, BuildFingerprint - must land on the emitted digest.
            #
            # "stored equals current" alone would be satisfied by two identically
            # WRONG production fingerprints; this is the check that cannot be.
            # The direct primitive proof in P5-D5 is a different claim and both
            # are required.
            if ($id -eq '1') {
                $golden = [string]$Cases.fingerprint.reference.digest
                $null = Add-Check $list `
                    'GOLDEN: PCCM_CurrentInputFingerprint() on plan case 1 equals the emitted reference digest' `
                    ($current -ceq $golden) ("got " + $current + ", emitted " + $golden)
                $null = Add-Check $list `
                    'GOLDEN: PCCM_CalculationFingerprint() after the commit equals the emitted reference digest' `
                    ($fingerprint -ceq $golden) ("got " + $fingerprint + ", emitted " + $golden)
            }

            if ($attempt -eq 'SUCCESS') {
                Add-Phase5AnalyticalChecks -List $list -Workbook $Workbook -Inspection $Inspection `
                    -Case $case -Tolerances $Cases.tolerances
                Add-Phase5SuccessStateChecks -List $list -Excel $Excel -Workbook $Workbook `
                    -Inspection $Inspection -Case $case -Cases $Cases -Label ('case ' + $id)
            }
            $covered += $id
        }
        $null = Add-Check $list 'every analytical plan case was driven' ($covered.Count -eq 19) `
            ("covered " + $covered.Count + ": " + ($covered -join ', '))
        Add-Phase5Result 'P5-AN' `
            ('Analytical fixtures through PCCM_Calculate: ' + $covered.Count + ' cases, all emitted values') `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Phase5Result 'P5-AN' 'Analytical fixtures through PCCM_Calculate' 'FAIL' (Format-Phase5Err $_)
    }

    # -------------------------------------------------------------------
    # P5-RF. Every prerequisite refusal in the fixture corpus
    # -------------------------------------------------------------------
    # A REFUSAL PRESERVES THE PRIOR SUCCESSFUL SNAPSHOT. It does not empty the
    # calculation workspace.
    #
    # The first submission asserted the opposite: that every _Calc table held
    # zero populated rows after a refusal. That would have FAILED against correct
    # production behaviour. P5-AN runs first and leaves a successful snapshot;
    # Set-Phase5Fixture changes the INPUT model and never touches _Calc; and a
    # pre-write refusal is required to leave C13:C16, C23:C32 and all five tables
    # exactly as they were.
    #
    # "No partial result" means NO PARTIAL NEW SNAPSHOT SURVIVES - not "erase the
    # old successful one". The baseline is therefore established once and every
    # refusal is compared against it, which also proves that a run of successive
    # refusals never erodes the snapshot.
    try {
        $list = New-Checklist
        $baseline = $null
        foreach ($candidate in @($Cases.plan_cases)) {
            if ([string]$candidate.id -eq '3') { $baseline = $candidate }
        }
        $null = Add-Check $list 'the refusal baseline fixture (plan case 3) was emitted' `
            ($null -ne $baseline)
        $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
            -Inspection $Inspection -Model $baseline.model
        $Excel.Run('PCCM_Calculate') | Out-Null
        $null = Add-Check $list 'a successful baseline snapshot was established first' `
            ([string]$Excel.Run('PCCM_CalculationAttemptResult') -eq 'SUCCESS')
        $before = Get-Phase5Snapshot -Workbook $Workbook -Inspection $Inspection
        $baselineDigest = [string]$Excel.Run('PCCM_CalculationFingerprint')
        $null = Add-Check $list 'the baseline snapshot is not empty, so the comparison is not vacuous' `
            ((@($before.Tables['calc_drivers'])).Count -ge 1) `
            ("calc_drivers rows " + (@($before.Tables['calc_drivers'])).Count)

        $covered = @()
        foreach ($case in @($Cases.plan_cases)) {
            if ([string]$case.kind -ne 'refusal') { continue }
            $id = [string]$case.id
            # The INPUT model changes. The successful _Calc snapshot is left
            # exactly where it is, deliberately: clearing it to make an assertion
            # pass would be the harness proving itself.
            $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
                -Inspection $Inspection -Model $case.model

            $Excel.Run('PCCM_Calculate') | Out-Null
            $attempt = [string]$Excel.Run('PCCM_CalculationAttemptResult')
            $detail = [string]$Excel.Run('PCCM_CalculationAttemptDetail')
            $status = [string]$Excel.Run('PCCM_CalculationStatus')
            $null = Add-Check $list ('case ' + $id + ' (' + [string]$case.expected_refusal + '): REFUSED') `
                ($attempt -eq 'REFUSED') ("attempt '" + $attempt + "'")
            Add-Phase5DetailTokenChecks -List $list -Detail $detail `
                -Tokens (Get-Phase5PlanRefusalTokens -Cases $Cases -PlanCaseId ([string]$case.id)) `
                -Label ('case ' + $id)
            $null = Add-Check $list ('case ' + $id + ': the derived status is INVALID') `
                ($status -eq 'INVALID') $status

            $after = Get-Phase5Snapshot -Workbook $Workbook -Inspection $Inspection
            # GROUP ONE AND THREE: unchanged, exactly.
            Add-SnapshotUnchangedChecks -List $list -Before $before -After $after `
                -Label ('case ' + $id) -SuccessFields $successRecordFields
            # GROUP TWO: changed, as the refusal requires.
            Add-Phase5AttemptAxisChecks -List $list -After $after -Label ('case ' + $id) `
                -ExpectedResult 'REFUSED' -ExpectedStatus 'INVALID'
            $null = Add-Check $list ('case ' + $id + ': the stored digest is still the baseline one') `
                ([string]$Excel.Run('PCCM_CalculationFingerprint') -ceq $baselineDigest)
            # NO PARTIAL NEW SNAPSHOT: the current inputs are refused, so no
            # digest for them exists and none was published.
            $null = Add-Check $list ('case ' + $id + ': no current-input fingerprint exists for a refused model') `
                ([string]::IsNullOrEmpty([string]$Excel.Run('PCCM_CurrentInputFingerprint')))
            $covered += $id
        }
        $null = Add-Check $list 'every refusal plan case was driven' ($covered.Count -eq 9) `
            ("covered " + $covered.Count + ": " + ($covered -join ', '))
        Add-Phase5Result 'P5-RF' `
            ('Prerequisite refusals: ' + $covered.Count + ' cases; the prior successful snapshot survives each') `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Phase5Result 'P5-RF' 'Prerequisite refusals' 'FAIL' (Format-Phase5Err $_)
    }

    # -------------------------------------------------------------------
    # P5-PQ. THE COMPLETE plan section 18 PREREQUISITE MATRIX
    # -------------------------------------------------------------------
    # The nine refusal plan cases do not exhaust section 18. Base Year after
    # Start Year, STRUCTURE CHANGE PENDING, a duplicated referenced currency, a
    # non-numeric Probability, an unknown Distribution and a dozen more locked
    # predicates had no real-Windows scenario at all.
    #
    # The matrix is EMITTED, in phase5_cases.json -> gate_b.prerequisite_cases.
    # This scenario consumes it; it holds no list of its own, so a locked
    # predicate cannot be dropped by editing this file.
    try {
        $list = New-Checklist
        $prerequisites = @($Cases.gate_b.prerequisite_cases)
        $null = Add-Check $list 'the emitted prerequisite matrix is present' `
            ($prerequisites.Count -ge 1) ("cases " + $prerequisites.Count)

        # One successful baseline, and every prerequisite refusal is compared
        # against it: the prior snapshot must survive each, exactly as P5-RF
        # proves for the plan-case refusals.
        $planCase = @{}
        foreach ($candidate in @($Cases.plan_cases)) { $planCase[[string]$candidate.id] = $candidate }
        $covered = @()
        foreach ($entry in $prerequisites) {
            $id = [string]$entry.id
            $base = $planCase[[string]$entry.base_plan_case]
            if ($null -eq $base) {
                $null = Add-Check $list ($id + ': the base plan case exists') $false `
                    ("base_plan_case " + [string]$entry.base_plan_case)
                continue
            }
            $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
                -Inspection $Inspection -Model $base.model
            $Excel.Run('PCCM_Calculate') | Out-Null
            $null = Add-Check $list ($id + ': the unmutated base fixture calculates') `
                ([string]$Excel.Run('PCCM_CalculationAttemptResult') -eq 'SUCCESS') `
                ("base plan case " + [string]$entry.base_plan_case)
            $before = Get-Phase5Snapshot -Workbook $Workbook -Inspection $Inspection

            # THE MUTATION, from the corpus.
            Invoke-Phase5Mutation -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
                -Inspection $Inspection -Mutation $entry.mutation

            $Excel.Run('PCCM_Calculate') | Out-Null
            $attempt = [string]$Excel.Run('PCCM_CalculationAttemptResult')
            $detail = [string]$Excel.Run('PCCM_CalculationAttemptDetail')
            $status = [string]$Excel.Run('PCCM_CalculationStatus')
            $null = Add-Check $list ($id + ' (' + [string]$entry.predicate + '): attempt = ' + [string]$entry.expected_attempt) `
                ($attempt -eq [string]$entry.expected_attempt) ("got '" + $attempt + "'")
            $null = Add-Check $list ($id + ': status = ' + [string]$entry.expected_status) `
                ($status -eq [string]$entry.expected_status) ("got '" + $status + "'")
            Add-Phase5DetailTokenChecks -List $list -Detail $detail `
                -Tokens $entry.detail_tokens -Label $id
            if ($entry.snapshot_unchanged) {
                $after = Get-Phase5Snapshot -Workbook $Workbook -Inspection $Inspection
                Add-SnapshotUnchangedChecks -List $list -Before $before -After $after `
                    -Label $id -SuccessFields $successRecordFields
                Add-Phase5AttemptAxisChecks -List $list -After $after -Label $id `
                    -ExpectedResult ([string]$entry.expected_attempt) `
                    -ExpectedStatus ([string]$entry.expected_status)
            }
            $covered += $id
        }
        $null = Add-Check $list 'every emitted prerequisite case was driven' `
            ($covered.Count -eq $prerequisites.Count) `
            ("covered " + $covered.Count + " of " + $prerequisites.Count)
        Add-Phase5Result 'P5-PQ' `
            ('Plan section 18 prerequisite matrix: ' + $covered.Count + ' predicates, each with its own detail discriminator') `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Phase5Result 'P5-PQ' 'Plan section 18 prerequisite matrix' 'FAIL' (Format-Phase5Err $_)
    }

    # -------------------------------------------------------------------
    # P5-PN. THE REFERENCED-ONLY COMPLEMENT: does not block AND does not affect
    # -------------------------------------------------------------------
    # A harness that only proved refusals would accept a model that refused too
    # much. An assumption nobody references cannot block a valid model, and that
    # is a locked semantic in its own right.
    #
    # BUT "SUCCESS, CURRENT, blank detail, same digest" IS NOT THE WHOLE CLAIM.
    # A defect that excluded the unreferenced assumption from the FINGERPRINT
    # while accidentally consuming it in the CALCULATION would satisfy every one
    # of those and still publish wrong numbers. Referenced-only means the
    # assumption is outside the calculation model, not merely outside the digest.
    #
    # So every no-block scenario also re-asserts the COMPLETE analytical
    # workspace - all five tables and calc_totals - against the emitted expected
    # block of its own base plan case. No new corpus is needed: the base case
    # already carries it.
    try {
        $list = New-Checklist
        $noBlock = @($Cases.gate_b.no_block_cases)
        $null = Add-Check $list 'the emitted no-block matrix is present' `
            ($noBlock.Count -ge 1) ("cases " + $noBlock.Count)
        $planCase = @{}
        foreach ($candidate in @($Cases.plan_cases)) { $planCase[[string]$candidate.id] = $candidate }
        $covered = @()
        foreach ($entry in $noBlock) {
            $id = [string]$entry.id
            $base = $planCase[[string]$entry.base_plan_case]
            $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
                -Inspection $Inspection -Model $base.model
            $Excel.Run('PCCM_Calculate') | Out-Null
            $baseline = [string]$Excel.Run('PCCM_CalculationFingerprint')

            Invoke-Phase5Mutation -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
                -Inspection $Inspection -Mutation $entry.mutation

            $Excel.Run('PCCM_Calculate') | Out-Null
            $attempt = [string]$Excel.Run('PCCM_CalculationAttemptResult')
            $detail = [string]$Excel.Run('PCCM_CalculationAttemptDetail')
            $status = [string]$Excel.Run('PCCM_CalculationStatus')
            $null = Add-Check $list ($id + ' (' + [string]$entry.predicate + '): attempt = ' + [string]$entry.expected_attempt) `
                ($attempt -eq [string]$entry.expected_attempt) ("got '" + $attempt + "', detail '" + $detail + "'")
            $null = Add-Check $list ($id + ': status = ' + [string]$entry.expected_status) `
                ($status -eq [string]$entry.expected_status) ("got '" + $status + "'")
            $null = Add-Check $list ($id + ': the detail stays blank - nothing was refused') `
                ([string]::IsNullOrEmpty($detail)) $detail
            # The unreferenced assumption is not merely tolerated: it changes
            # NOTHING, so the digest is the one the clean model produced.
            $null = Add-Check $list ($id + ': the stored fingerprint is unchanged by the unreferenced row') `
                ([string]$Excel.Run('PCCM_CalculationFingerprint') -ceq $baseline)

            # AND THE NUMBERS ARE STILL THE BASE CASE'S. The whole analytical
            # workspace is compared against the emitted expected block of the
            # base plan case, so an assumption that leaked into the calculation
            # while staying out of the digest fails here rather than passing on
            # four green flags.
            if ($attempt -eq 'SUCCESS') {
                Add-Phase5AnalyticalChecks -List $list -Workbook $Workbook `
                    -Inspection $Inspection -Case $base -Tolerances $Cases.tolerances
                # The successful record too. The two timestamps are NOT required
                # to equal the first calculation's: a recalculation may refresh
                # them, and the contract does not say otherwise.
                Add-Phase5SuccessStateChecks -List $list -Excel $Excel -Workbook $Workbook `
                    -Inspection $Inspection -Case $base -Cases $Cases `
                    -Label ($id + ' recalculated')
            }
            $covered += $id
        }
        $null = Add-Check $list 'every emitted no-block case was driven' `
            ($covered.Count -eq $noBlock.Count) `
            ("covered " + $covered.Count + " of " + $noBlock.Count)
        Add-Phase5Result 'P5-PN' `
            ('Referenced-only no-block semantics: ' + $covered.Count + ' assumptions that must not block') `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Phase5Result 'P5-PN' 'Referenced-only no-block semantics' 'FAIL' (Format-Phase5Err $_)
    }

    # -------------------------------------------------------------------
    # P5-AR. THE DRIVER-AUDIT A/B/C/D RECONSTRUCTION
    # -------------------------------------------------------------------
    # A CROSS-CHECK BETWEEN TWO PARTS OF THE REAL WORKBOOK, not a second
    # comparison against the oracle. Every driver row and every headline total is
    # already asserted against the emitted oracle by Add-Phase5AnalyticalChecks;
    # this proves the published audit COLUMNS actually reconstruct the published
    # headline TOTALS, partitioned by driver kind:
    #
    #   A = sum of the Cost Line deterministic columns
    #   B = sum of the Cost Line uncertainty-mean-shift columns
    #   C = sum of the Cost Line mean-basis columns
    #   D = sum of the Risk expected-risk columns
    #
    # The mapping is EMITTED, in gate_b.audit_reconstruction.relationships, so
    # PowerShell never restates "column 18" in its own source. The fixture has
    # three Cost Lines and two Risks, so none of the four is trivial.
    try {
        $list = New-Checklist
        $audit = $Cases.gate_b.audit_reconstruction
        $null = Add-Check $list 'the emitted audit fixture is present' ($null -ne $audit)
        $null = Add-Check $list 'the audit fixture has more than one Cost Line' `
            ((@($audit.model.cost_lines)).Count -ge 2) `
            ("cost lines " + (@($audit.model.cost_lines)).Count)
        $null = Add-Check $list 'the audit fixture has at least one Risk' `
            ((@($audit.model.risks)).Count -ge 1) ("risks " + (@($audit.model.risks)).Count)

        $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
            -Inspection $Inspection -Model $audit.model
        $Excel.Run('PCCM_Calculate') | Out-Null
        $null = Add-Check $list 'the multi-driver audit fixture calculated' `
            ([string]$Excel.Run('PCCM_CalculationAttemptResult') -eq 'SUCCESS')

        # The published values still equal the oracle - the audit relationship is
        # an ADDITIONAL claim, not a replacement for value equality.
        Add-Phase5AnalyticalChecks -List $list -Workbook $Workbook -Inspection $Inspection `
            -Case $audit -Tolerances $Cases.tolerances

        $rows = @(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection -TableKey 'calc_drivers')
        $kindColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_drivers' -ColumnKey 'driver_kind'
        $null = Add-Check $list 'tblCalcDrivers carries rows of both kinds' `
            ((@($rows | Where-Object { [string]$_[$kindColumn] -eq 'Cost Line' }).Count -ge 2) -and
             (@($rows | Where-Object { [string]$_[$kindColumn] -eq 'Risk' }).Count -ge 1))

        foreach ($relationship in @($audit.relationships)) {
            $ordinal = Get-CalcTableColumnIndex -Inspection $Inspection `
                -TableKey 'calc_drivers' -ColumnKey ([string]$relationship.driver_column)
            $null = Add-Check $list `
                ([string]$relationship.headline + ': tblCalcDrivers publishes ' + [string]$relationship.driver_column) `
                ($ordinal -ge 0)
            if ($ordinal -lt 0) { continue }

            # PARTITIONED BY KIND, and a BLANK IS SKIPPED, never read as the
            # opposite kind's identity 1. An N/A cell means the column does not
            # apply to that row; folding a 1 into the sum would fabricate a
            # contribution the model never made.
            $sum = 0.0
            $contributors = 0
            $wrongKindPopulated = 0
            foreach ($row in $rows) {
                $isKind = ([string]$row[$kindColumn] -eq [string]$relationship.kind)
                $blank = Test-CalcBlank -Actual $row[$ordinal]
                if ($isKind) {
                    if (-not $blank) {
                        $sum = $sum + [double]$row[$ordinal]
                        $contributors++
                    }
                } elseif (-not $blank) {
                    $wrongKindPopulated++
                }
            }
            $null = Add-Check $list `
                ([string]$relationship.headline + ': the ' + [string]$relationship.kind + ' partition contributed rows') `
                ($contributors -ge 1) ("contributors " + $contributors)
            $null = Add-Check $list `
                ([string]$relationship.headline + ': the opposite kind publishes BLANK in ' + [string]$relationship.driver_column) `
                ($wrongKindPopulated -eq 0) ("populated on the wrong kind: " + $wrongKindPopulated)

            $headline = Get-CalcScalar -Workbook $Workbook -Inspection $Inspection `
                -Block 'calc_totals' -FieldKey ([string]$relationship.headline)
            # EXACT. NO TOLERANCE AT ALL.
            #
            # This is an AUDIT relationship - "does the audit table reconstruct
            # the headline the workbook publishes?" - not the I1-I4
            # reconciliation, which has its own production allowance. Both sides
            # are the same published Doubles summed in the same order, and the
            # emitted fixture was chosen because it reconstructs to the identical
            # IEEE Double. A relative epsilon here, even 1e-12, could hide a small
            # but real mismatch between two things that must agree bit for bit.
            #
            # Test-CalcValue's Tolerance defaults to 0.0; it is passed explicitly
            # so the intent is unmistakable in review.
            $null = Add-Check $list `
                ([string]$relationship.headline + ' = SUM(' + [string]$relationship.driver_column + ') over ' + [string]$relationship.kind + ' rows, EXACTLY') `
                (Test-CalcValue -Actual $headline -Expected $sum -Tolerance 0.0) `
                ("calc_totals " + (Format-CalcValue $headline) + ", reconstructed " + (Format-CalcValue $sum))
        }
        Add-Phase5Result 'P5-AR' `
            'Driver-audit reconstruction: A/B/C/D from the ACTUAL tblCalcDrivers columns to the ACTUAL calc_totals' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Phase5Result 'P5-AR' 'Driver-audit reconstruction' 'FAIL' (Format-Phase5Err $_)
    }

    # -------------------------------------------------------------------
    # P5-ID. The locked reconciliation identities I1, I2, I3a-c, I4a-c and I5
    # -------------------------------------------------------------------
    # THE PRODUCTION RECONCILIATION IS THE AUTHORITY, AND IT IS NOT REIMPLEMENTED
    # HERE.
    #
    # The first submission relabelled the identity set and, worse, decided each
    # identity with its own PowerShell tolerance of the shape
    #     max(|left|, |right|, floor) * coefficient
    # which is exactly the HEADLINE-BASED conditioning erratum C1 rejected. Plan
    # case 30 exists because headline conditioning can falsely fail a correct
    # cancellation-heavy calculation, so reintroducing it in PowerShell would
    # have made a rejected oracle the acceptance authority.
    #
    # What actually proves the identities on real Excel:
    #
    #   A  production `Reconcile` / `AllIdentitiesHold` runs INSIDE PCCM_Calculate
    #      and a commit is impossible unless they pass. A SUCCESS on case 30 is
    #      therefore production's own statement that the identities held under
    #      the accepted C1 allowance.
    #   B  every published analytical value equals the emitted Python oracle -
    #      asserted in full by Add-Phase5AnalyticalChecks.
    #   C  the annual Base, Risk and Total columns are each checked against their
    #      own emitted oracle values, nominal and PV, which is what I3a-c and
    #      I4a-c are statements about.
    #   D  the profile weight vector of each driver is asserted from the emitted
    #      corpus, and the refusal matrix proves invalid I5 variants are refused.
    #
    # The additive algebra below is kept because the MAPPING must be visible, but
    # each side is compared against the ORACLE's own value for that side. No new
    # PowerShell tolerance decides anything.
    try {
        $list = New-Checklist
        $tolerance = [double]$Cases.tolerances.identity_relative_coefficient
        $identityCases = @('3', '9', '30')
        $seen = @()
        foreach ($candidate in @($Cases.plan_cases)) {
            if ([string]$candidate.kind -ne 'analytical') { continue }
            $id = [string]$candidate.id
            if ($identityCases -notcontains $id) { continue }
            $seen += $id

            $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
                -Inspection $Inspection -Model $candidate.model
            $Excel.Run('PCCM_Calculate') | Out-Null
            $attempt = [string]$Excel.Run('PCCM_CalculationAttemptResult')
            # A. PRODUCTION ACCEPTED IT. This is the identity evidence: the commit
            #    is unreachable unless Reconcile and AllIdentitiesHold passed.
            $null = Add-Check $list `
                ('case ' + $id + ': PCCM_Calculate COMMITTED, so production Reconcile / AllIdentitiesHold passed') `
                ($attempt -eq 'SUCCESS') ("attempt '" + $attempt + "'")
            if ($attempt -ne 'SUCCESS') { continue }

            # B. Every published value equals the oracle. Same path as P5-AN, run
            #    again here so the identity evidence is self-contained.
            Add-Phase5AnalyticalChecks -List $list -Workbook $Workbook -Inspection $Inspection `
                -Case $candidate -Tolerances $Cases.tolerances

            $expected = $candidate.expected
            $totals = @{}
            foreach ($field in $expected.totals.PSObject.Properties.Name) {
                $totals[$field] = Get-CalcScalar -Workbook $Workbook -Inspection $Inspection `
                    -Block 'calc_totals' -FieldKey $field
            }
            $oracle = @{}
            foreach ($field in $expected.totals.PSObject.Properties.Name) {
                $oracle[$field] = [double]$expected.totals.$field
            }

            # --- I1  A + B = C, nominal AND present value -------------------
            # The workbook's C is compared against the ORACLE's C, and the oracle's
            # own A + B is shown alongside so the mapping is auditable. The
            # comparison authority is the oracle value, never a PowerShell sum.
            foreach ($pair in @(
                @{ name = 'I1 nominal'; left = 'c_nom'; a = 'a_nom'; b = 'b_nom' },
                @{ name = 'I1 present value'; left = 'c_pv'; a = 'a_pv'; b = 'b_pv' })) {
                $null = Add-Check $list ('case ' + $id + ' ' + $pair.name + ': ' + $pair.left + ' equals the oracle') `
                    (Test-CalcValue -Actual $totals[$pair.left] -Expected $oracle[$pair.left] -Tolerance $tolerance) `
                    ("got " + (Format-CalcValue $totals[$pair.left]) + ", oracle " + (Format-CalcValue $oracle[$pair.left]))
                $null = Add-Check $list ('case ' + $id + ' ' + $pair.name + ': the oracle maps ' + $pair.a + ' + ' + $pair.b + ' -> ' + $pair.left) `
                    ($null -ne $oracle[$pair.a] -and $null -ne $oracle[$pair.b])
            }
            # --- I2  C + D = E, nominal AND present value -------------------
            foreach ($pair in @(
                @{ name = 'I2 nominal'; left = 'e_nom'; c = 'c_nom'; d = 'd_nom' },
                @{ name = 'I2 present value'; left = 'e_pv'; c = 'c_pv'; d = 'd_pv' })) {
                $null = Add-Check $list ('case ' + $id + ' ' + $pair.name + ': ' + $pair.left + ' equals the oracle') `
                    (Test-CalcValue -Actual $totals[$pair.left] -Expected $oracle[$pair.left] -Tolerance $tolerance) `
                    ("got " + (Format-CalcValue $totals[$pair.left]) + ", oracle " + (Format-CalcValue $oracle[$pair.left]))
                $null = Add-Check $list ('case ' + $id + ' ' + $pair.name + ': the oracle maps ' + $pair.c + ' + ' + $pair.d + ' -> ' + $pair.left) `
                    ($null -ne $oracle[$pair.c] -and $null -ne $oracle[$pair.d])
            }

            # --- I3a/b/c and I4a/b/c  the annual columns --------------------
            # Each annual column is asserted against its OWN emitted oracle value,
            # row by row, and the headline it reconciles to is named. Base, Risk
            # and Total are asserted SEPARATELY - the first submission checked
            # only the Total column and called it I5.
            $annual = @(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection -TableKey 'calc_annual')
            $indexColumn = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey 'calc_annual' -ColumnKey 'project_index'
            foreach ($identity in @(
                @{ name = 'I3a'; column = 'base_cost_nominal';    headline = 'c_nom' },
                @{ name = 'I3b'; column = 'expected_risk_nominal'; headline = 'd_nom' },
                @{ name = 'I3c'; column = 'total_nominal';         headline = 'e_nom' },
                @{ name = 'I4a'; column = 'base_cost_pv';          headline = 'c_pv' },
                @{ name = 'I4b'; column = 'expected_risk_pv';      headline = 'd_pv' },
                @{ name = 'I4c'; column = 'total_pv';              headline = 'e_pv' })) {
                $ordinal = Get-CalcTableColumnIndex -Inspection $Inspection `
                    -TableKey 'calc_annual' -ColumnKey $identity.column
                $null = Add-Check $list `
                    ('case ' + $id + ' ' + $identity.name + ': tblCalcAnnual publishes ' + $identity.column) `
                    ($ordinal -ge 0)
                foreach ($wanted in @($expected.annual)) {
                    $found = $null
                    foreach ($row in $annual) {
                        if ([int]$row[$indexColumn] -eq [int]$wanted.project_index) { $found = $row }
                    }
                    $label = ('case ' + $id + ' ' + $identity.name + ': year ' +
                              [string]$wanted.project_index + ' ' + $identity.column)
                    if ($null -eq $found) {
                        $null = Add-Check $list ($label + ' row exists') $false 'no such row'
                        continue
                    }
                    $null = Add-Check $list ($label + ' equals the oracle') `
                        (Test-CalcValue -Actual $found[$ordinal] `
                            -Expected $wanted.($identity.column) -Tolerance $tolerance) `
                        ("got " + (Format-CalcValue $found[$ordinal]) + `
                         ", oracle " + (Format-CalcValue $wanted.($identity.column)))
                }
                $null = Add-Check $list `
                    ('case ' + $id + ' ' + $identity.name + ': the series reconciles to ' + $identity.headline) `
                    (Test-CalcValue -Actual $totals[$identity.headline] `
                        -Expected $oracle[$identity.headline] -Tolerance $tolerance) `
                    ("headline " + (Format-CalcValue $totals[$identity.headline]))
            }

            # --- I5  profile weights sum to 1 PER DRIVER --------------------
            # The vector is asserted against the emitted corpus, driver by driver,
            # in the profiling grid the fixture wrote it into. The SUM itself is
            # production's to enforce: cases 15 and 23 in P5-RF prove an invalid
            # vector is refused, which is the other half of I5.
            foreach ($pair in @(
                @{ key = 'cost_profiling'; drivers = @($candidate.model.cost_lines) },
                @{ key = 'risk_profiling'; drivers = @($candidate.model.risks) })) {
                $grid = $null
                foreach ($gridCandidate in @($Manifest.grids)) {
                    if ($gridCandidate.key -eq $pair.key) { $grid = $gridCandidate }
                }
                $fixed = @($grid.fixed_columns).Count
                # THE TYPED READER. These weights are compared with the
                # type-sensitive analytical comparator, so a stringifying reader
                # would have failed every one of them against a correct workbook -
                # the same defect the _Calc tables had.
                $body = @(Get-Phase5TypedTableBody -Workbook $Workbook -SheetName $grid.sheet `
                    -TableName $grid.table_name)
                foreach ($driver in $pair.drivers) {
                    $found = $null
                    foreach ($row in $body) {
                        if ([string]$row[0] -eq [string]$driver.permanent_id) { $found = $row }
                    }
                    $label = 'case ' + $id + ' I5: ' + [string]$driver.permanent_id
                    if ($null -eq $found) {
                        $null = Add-Check $list ($label + ' has a profiling row') $false
                        continue
                    }
                    $offset = 0
                    foreach ($weight in @($driver.profile_weights)) {
                        $offset++
                        $null = Add-Check $list ($label + ' weight ' + $offset + ' is the emitted one') `
                            (Test-CalcValue -Actual $found[$fixed + $offset - 1] -Expected $weight) `
                            ("got " + (Format-CalcValue $found[$fixed + $offset - 1]) + `
                             ", emitted " + (Format-CalcValue $weight))
                    }
                }
            }
            # The emitted driver record carries the same vector, so the two
            # authorities for I5 are asserted to agree.
            foreach ($wanted in @($expected.drivers)) {
                $model = $null
                foreach ($driver in @($candidate.model.cost_lines) + @($candidate.model.risks)) {
                    if ([string]$driver.permanent_id -eq [string]$wanted.permanent_id) { $model = $driver }
                }
                $null = Add-Check $list `
                    ('case ' + $id + ' I5: the emitted weights for ' + [string]$wanted.permanent_id + ' match the fixture') `
                    ((@($wanted.weights) -join ',') -eq (@($model.profile_weights) -join ','))
            }
        }
        $null = Add-Check $list 'the cancellation-heavy fixture (plan case 30) was among them' `
            ($seen -contains '30') ("covered " + ($seen -join ', '))
        Add-Phase5Result 'P5-ID' `
            'Reconciliation identities I1, I2, I3a-c, I4a-c, I5 - production Reconcile is the authority' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Phase5Result 'P5-ID' 'Reconciliation identities' 'FAIL' (Format-Phase5Err $_)
    }

    # ===================================================================
    # THE SIX-ROW STATUS MATRIX
    #
    # Every row asserts ALL FOUR accessors and, where applicable, the current
    # input fingerprint, plus the snapshot state the row requires. STATUS IS
    # NEVER DERIVED FROM ATTEMPT HISTORY: rows 5 and 6 exist precisely because
    # the two axes are allowed to disagree, and a harness that "tidied" that
    # disagreement would be asserting the defect.
    #
    # Every status read goes through PCCM_CalculationStatus FIRST. The status
    # cell is last-evaluated, not live: reading C19 without asking would report
    # whatever the previous scenario left there.
    # ===================================================================
    function Add-StatusRowChecks {
        param($List, $Excel, [string]$Row, [string]$ExpectedStatus, [string]$ExpectedAttempt,
              [string]$DetailRule, [string]$ExpectedFingerprint)
        $status = [string]$Excel.Run('PCCM_CalculationStatus')
        $attempt = [string]$Excel.Run('PCCM_CalculationAttemptResult')
        $detail = [string]$Excel.Run('PCCM_CalculationAttemptDetail')
        $stored = [string]$Excel.Run('PCCM_CalculationFingerprint')
        $current = [string]$Excel.Run('PCCM_CurrentInputFingerprint')
        $null = Add-Check $List ($Row + ': PCCM_CalculationStatus() = ' + $ExpectedStatus) `
            ($status -eq $ExpectedStatus) ("got '" + $status + "'")
        $null = Add-Check $List ($Row + ': PCCM_CalculationAttemptResult() = ' + $ExpectedAttempt) `
            ($attempt -eq $ExpectedAttempt) ("got '" + $attempt + "'")
        if ($DetailRule -eq 'blank') {
            $null = Add-Check $List ($Row + ': PCCM_CalculationAttemptDetail() is blank') `
                ([string]::IsNullOrEmpty($detail)) ("got '" + $detail + "'")
        } else {
            $null = Add-Check $List ($Row + ': PCCM_CalculationAttemptDetail() is specific') `
                (-not [string]::IsNullOrWhiteSpace($detail)) ("got '" + $detail + "'")
        }
        if ($ExpectedFingerprint -ne '') {
            $null = Add-Check $List ($Row + ': PCCM_CalculationFingerprint() is the expected snapshot digest') `
                ($stored -ceq $ExpectedFingerprint) ("got '" + $stored + "', expected '" + $ExpectedFingerprint + "'")
        }
        return [pscustomobject]@{
            Status = $status; Attempt = $attempt; Detail = $detail
            Stored = $stored; Current = $current
        }
    }

    # A single valid fixture underpins rows 1-6 and the staleness work below.
    $baseCase = $null
    foreach ($candidate in @($Cases.plan_cases)) {
        # Multi-year, compounded inflation, three profiling weights: the richest
        # analytical fixture the corpus emits, so a staleness edit has somewhere
        # meaningful to land.
        if ([string]$candidate.id -eq '3') { $baseCase = $candidate }
    }

    $establishedFingerprint = ''
    $establishedState = $null
    try {
        $list = New-Checklist
        $null = Add-Check $list 'the base fixture (plan case 3) was emitted' ($null -ne $baseCase)
        $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
            -Inspection $Inspection -Model $baseCase.model
        $Excel.Run('PCCM_Calculate') | Out-Null
        $row1 = Add-StatusRowChecks -List $list -Excel $Excel -Row 'row 1' `
            -ExpectedStatus 'CURRENT' -ExpectedAttempt 'SUCCESS' -DetailRule 'blank' -ExpectedFingerprint ''
        $establishedFingerprint = $row1.Stored
        $null = Add-Check $list 'row 1: a NEW snapshot was written (the stored digest is not empty)' `
            (-not [string]::IsNullOrEmpty($establishedFingerprint))
        $null = Add-Check $list 'row 1: the stored digest equals the current input digest' `
            ($row1.Stored -ceq $row1.Current)
        $establishedState = Get-CalcScalarBlock -Workbook $Workbook -Inspection $Inspection -Block 'calc_state'
        $null = Add-Check $list 'row 1: calc_state carries a fingerprint version' `
            (-not (Test-CalcBlank -Actual $establishedState['fingerprint_version']))
        Add-Phase5Result 'P5-S1' 'Status row 1: successful calculation, unchanged inputs' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Phase5Result 'P5-S1' 'Status row 1' 'FAIL' (Format-Phase5Err $_)
    }

    # -------------------------------------------------------------------
    # P5-S2 / P5-ST. A fingerprinted input changes, and NOTHING is calculated
    # -------------------------------------------------------------------
    # THE TRANSITION HAS AN ORACLE AT BOTH ENDS.
    #
    # The first submission made the model stale by exchanging two profiling
    # weights, which produces a model the corpus does not describe: after the
    # second Calculate there was no emitted expectation to compare against, and
    # the proof degenerated into an annual ROW COUNT. section 25.2 requires the affected
    # value to change TO THE ORACLE VALUE.
    #
    # Plan cases 3 and 19 have IDENTICAL applied structure - same timeline, same
    # FX, same inflation profile and rates, same driver, same weights - and differ
    # only in Discount Rate (0.10 -> -0.05). Changing that one ordinary
    # fingerprinted scalar turns the case-3 model into the case-19 model, so the
    # recalculated workbook can be asserted against case 19's OWN emitted expected
    # block, in full. No Apply Timeline is used to create staleness, and nothing
    # is hand-calculated in PowerShell.
    $targetCase = $null
    foreach ($candidate in @($Cases.plan_cases)) {
        if ([string]$candidate.id -eq '19') { $targetCase = $candidate }
    }
    try {
        $list = New-Checklist
        $null = Add-Check $list 'the staleness TARGET fixture (plan case 19) was emitted' `
            ($null -ne $targetCase)
        # The two fixtures really are the same structure apart from one scalar,
        # checked here rather than assumed, so a corpus change cannot silently
        # turn this into a two-variable transition.
        $sameStructure = (
            ((ConvertTo-Json $baseCase.model.timeline -Compress) -ceq
             (ConvertTo-Json $targetCase.model.timeline -Compress)) -and
            ((ConvertTo-Json $baseCase.model.fx -Compress) -ceq
             (ConvertTo-Json $targetCase.model.fx -Compress)) -and
            ((ConvertTo-Json $baseCase.model.inflation -Compress) -ceq
             (ConvertTo-Json $targetCase.model.inflation -Compress)) -and
            ((ConvertTo-Json $baseCase.model.cost_lines -Compress) -ceq
             (ConvertTo-Json $targetCase.model.cost_lines -Compress)) -and
            ((ConvertTo-Json $baseCase.model.risks -Compress) -ceq
             (ConvertTo-Json $targetCase.model.risks -Compress))
        )
        $null = Add-Check $list 'the source and target fixtures differ ONLY in Discount Rate' `
            ($sameStructure -and ([double]$baseCase.model.discount_rate -ne
                                  [double]$targetCase.model.discount_rate)) `
            ("source " + [string]$baseCase.model.discount_rate + `
             ", target " + [string]$targetCase.model.discount_rate)

        # THE WHOLE SUCCESSFUL SNAPSHOT, CAPTURED BEFORE THE EDIT.
        #
        # Row 2 previously proved only the accessor axis and C13:C16. A defect
        # where PCCM_CalculationStatus rewrote analytical outputs while merely
        # re-deriving the status would have passed it. C23:C32 and all five
        # analytical ListObjects are captured here and compared afterwards.
        $rowTwoBefore = Get-Phase5Snapshot -Workbook $Workbook -Inspection $Inspection

        # ONE ORDINARY FINGERPRINTED SCALAR. Not a timeline application.
        Set-NamedValue -Workbook $Workbook -DefinedName $Inspection.inputs.discount_rate.defined_name `
            -Value ([double]$targetCase.model.discount_rate)

        $row2 = Add-StatusRowChecks -List $list -Excel $Excel -Row 'row 2' `
            -ExpectedStatus 'STALE' -ExpectedAttempt 'SUCCESS' -DetailRule 'blank' `
            -ExpectedFingerprint $establishedFingerprint
        $null = Add-Check $list 'row 2: the STORED fingerprint is unchanged - no snapshot was written' `
            ($row2.Stored -ceq $establishedFingerprint)
        $null = Add-Check $list 'row 2: the CURRENT input fingerprint changed' `
            ($row2.Current -cne $establishedFingerprint) `
            ("stored " + $row2.Stored + ", current " + $row2.Current)

        # C13:C16, C23:C32 and the five tables, all unchanged. C17:C20 is NOT in
        # that comparison: PCCM_CalculationStatus deliberately refreshes C19 and
        # C20, and asserting the whole block unchanged would assert that asking
        # for the status did nothing.
        $rowTwoAfter = Get-Phase5Snapshot -Workbook $Workbook -Inspection $Inspection
        Add-SnapshotUnchangedChecks -List $list -Before $rowTwoBefore -After $rowTwoAfter `
            -Label 'row 2' -SuccessFields $successRecordFields
        $null = Add-Check $list 'row 2: C17 still records the previous SUCCESS' `
            ([string]$rowTwoAfter.State['last_attempt_result'] -eq 'SUCCESS') `
            ("got " + (Format-CalcValue $rowTwoAfter.State['last_attempt_result']))
        $null = Add-Check $list 'row 2: C18 is still blank' `
            (Test-CalcBlank -Actual $rowTwoAfter.State['last_attempt_detail'])
        $null = Add-Check $list 'row 2: C19 was re-derived to STALE by the status evaluation' `
            ([string]$rowTwoAfter.State['calculation_status'] -eq 'STALE') `
            ("got " + (Format-CalcValue $rowTwoAfter.State['calculation_status']))
        $null = Add-Check $list 'row 2: C20 carries a status-evaluation timestamp' `
            (-not (Test-CalcBlank -Actual $rowTwoAfter.State['status_evaluated_at']))
        Add-Phase5Result 'P5-S2' 'Status row 2: valid fingerprinted input changed, no Calculate' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)

        # --- the rest of the primary staleness sequence ---------------------
        $list = New-Checklist
        $Excel.Run('PCCM_Calculate') | Out-Null
        $attempt = [string]$Excel.Run('PCCM_CalculationAttemptResult')
        $status = [string]$Excel.Run('PCCM_CalculationStatus')
        $stored = [string]$Excel.Run('PCCM_CalculationFingerprint')
        $null = Add-Check $list 'recalculating returns the model to CURRENT' ($status -eq 'CURRENT') $status
        $null = Add-Check $list 'the attempt is SUCCESS' ($attempt -eq 'SUCCESS') $attempt
        $null = Add-Check $list 'the STORED fingerprint CHANGED - a new snapshot was written' `
            ($stored -cne $establishedFingerprint) `
            ("was " + $establishedFingerprint + ", now " + $stored)
        $null = Add-Check $list 'the stored digest now equals the current input digest' `
            ($stored -ceq [string]$Excel.Run('PCCM_CurrentInputFingerprint'))

        # THE ORACLE COMPARISON. Every published analytical value of the
        # recalculated model is asserted against plan case 19's emitted expected
        # block - the whole block, not a row count.
        if ($attempt -eq 'SUCCESS') {
            Add-Phase5AnalyticalChecks -List $list -Workbook $Workbook -Inspection $Inspection `
                -Case $targetCase -Tolerances $Cases.tolerances
            Add-Phase5SuccessStateChecks -List $list -Excel $Excel -Workbook $Workbook `
                -Inspection $Inspection -Case $targetCase -Cases $Cases -Label 'staleness target'
        }
        Add-Phase5Result 'P5-ST' `
            'Primary staleness sequence: case 3 -> Discount Rate -> case 19, verified against case 19s oracle' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)

    } catch {
        Add-Phase5Result 'P5-S2' 'Status row 2' 'FAIL' (Format-Phase5Err $_)
        Add-Phase5Result 'P5-ST' 'Primary staleness sequence' 'FAIL' (Format-Phase5Err $_)
    }

    # --- SETUP, NOT A SCENARIO ------------------------------------------
    # Restore the source fixture exactly, so the scenarios below start from the
    # model the corpus describes rather than from an edited one.
    #
    # THIS USED TO SIT INSIDE THE BLOCK ABOVE, after both scenarios had already
    # been recorded. When Runtime Run 4's typed-write defect made it throw, the
    # enclosing catch recorded P5-S2 and P5-ST a second time - two real results
    # followed by two spurious ones. A step that runs after a scenario has
    # finished is not part of that scenario, and it now owns its own failure.
    try {
        $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
            -Inspection $Inspection -Model $baseCase.model
        $Excel.Run('PCCM_Calculate') | Out-Null
        $establishedFingerprint = [string]$Excel.Run('PCCM_CalculationFingerprint')
        $establishedState = Get-CalcScalarBlock -Workbook $Workbook -Inspection $Inspection -Block 'calc_state'
    } catch {
        Add-Phase5Result 'P5-SU' `
            'Re-establishing the base fixture after the staleness sequence' 'FAIL' `
            (Format-Phase5Err $_)
    }

    # -------------------------------------------------------------------
    # P5-NS. NON-staleness: four INDEPENDENT changes, each restored
    # -------------------------------------------------------------------
    # THE FOUR EXCLUSIONS ARE STATED INDEPENDENTLY, SO THEY ARE PROVED
    # INDEPENDENTLY.
    #
    # The first submission accumulated each change while testing the next, so by
    # the fourth probe four edits were live at once and no single exclusion had
    # been isolated. Every probe now runs
    #     baseline CURRENT / SUCCESS / digest F
    #     -> change ONE excluded input
    #     -> assert CURRENT / SUCCESS / F
    #     -> restore exactly
    #     -> assert the baseline is back
    # before the next one begins.
    #
    # THE ROW-ORDER PROBE NEEDS MORE THAN ONE ROW. Plan case 3 has a single Cost
    # Line, and sorting a one-row table changes nothing: the Sort call was real
    # and the reorder evidence was not. It runs on plan case 30, which has three
    # Cost Lines, and the actual permanent-ID order is captured before and after
    # and required to have CHANGED.
    try {
        $list = New-Checklist
        $costReg = $null
        foreach ($register in @($Manifest.registers)) {
            if ($register.key -eq 'cost_lines') { $costReg = $register }
        }
        $descriptionOrdinal = [array]::IndexOf(@($costReg.columns), 'description') + 1

        # The probe: assert the excluded change moved nothing, against a digest
        # captured at the moment the probe began.
        $probe = {
            param([string]$Name, [string]$Digest)
            $status = [string]$Excel.Run('PCCM_CalculationStatus')
            $attempt = [string]$Excel.Run('PCCM_CalculationAttemptResult')
            $stored = [string]$Excel.Run('PCCM_CalculationFingerprint')
            $current = [string]$Excel.Run('PCCM_CurrentInputFingerprint')
            $null = Add-Check $list ($Name + ': status stays CURRENT') ($status -eq 'CURRENT') $status
            $null = Add-Check $list ($Name + ': attempt stays SUCCESS') ($attempt -eq 'SUCCESS') $attempt
            $null = Add-Check $list ($Name + ': the stored fingerprint is unchanged') `
                ($stored -ceq $Digest) ("got " + $stored)
            $null = Add-Check $list ($Name + ': the CURRENT input fingerprint is unchanged too') `
                ($current -ceq $Digest) ("got " + $current)
        }

        # ---- probe 1: Description ------------------------------------------
        $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
            -Inspection $Inspection -Model $baseCase.model
        $Excel.Run('PCCM_Calculate') | Out-Null
        $digest = [string]$Excel.Run('PCCM_CalculationFingerprint')
        $null = Add-Check $list 'probe 1 baseline: CURRENT / SUCCESS' `
            (([string]$Excel.Run('PCCM_CalculationStatus') -eq 'CURRENT') -and
             ([string]$Excel.Run('PCCM_CalculationAttemptResult') -eq 'SUCCESS'))
        $wasDescription = ''
        $body = @(Get-TableBody -Workbook $Workbook -SheetName $costReg.sheet -TableName $costReg.table_name)
        if ($body.Count -ge 1) { $wasDescription = [string]$body[0][$descriptionOrdinal - 1] }
        Set-TableCell -Workbook $Workbook -SheetName $costReg.sheet -TableName $costReg.table_name `
            -RowIndex 1 -ColumnIndex $descriptionOrdinal -Value 'a different description entirely'
        & $probe 'Description changed' $digest
        Set-TableCell -Workbook $Workbook -SheetName $costReg.sheet -TableName $costReg.table_name `
            -RowIndex 1 -ColumnIndex $descriptionOrdinal -Value $wasDescription
        & $probe 'Description restored' $digest

        # ---- probe 2: a REAL multi-row reorder ------------------------------
        $reorderCase = $null
        foreach ($candidate in @($Cases.plan_cases)) {
            if ([string]$candidate.id -eq '30') { $reorderCase = $candidate }
        }
        $null = Add-Check $list 'the reorder fixture (plan case 30) was emitted' ($null -ne $reorderCase)
        $null = Add-Check $list 'the reorder fixture has MORE THAN ONE Cost Line, so a sort can move rows' `
            ((@($reorderCase.model.cost_lines)).Count -ge 2) `
            ("cost lines " + (@($reorderCase.model.cost_lines)).Count)
        $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
            -Inspection $Inspection -Model $reorderCase.model
        $Excel.Run('PCCM_Calculate') | Out-Null
        $reorderDigest = [string]$Excel.Run('PCCM_CalculationFingerprint')
        $null = Add-Check $list 'probe 2 baseline: CURRENT / SUCCESS on the multi-row fixture' `
            (([string]$Excel.Run('PCCM_CalculationStatus') -eq 'CURRENT') -and
             ([string]$Excel.Run('PCCM_CalculationAttemptResult') -eq 'SUCCESS'))

        # NO CELL IS EDITED HERE. Write-Phase5Driver already gives every row a
        # deterministic, distinct Description - "GateB <PermanentId>" - so the
        # existing values are sufficient sort keys. The previous version rewrote
        # them to manufacture an ordering, which changed TWO non-fingerprinted
        # dimensions at once and stopped this being a row-order-only proof.
        $idsBefore = @(Get-IdColumnValues -Workbook $Workbook -Info $costReg)
        # Descending on the EXISTING descriptions, which reverses the ascending
        # order the fixture applier created.
        Invoke-TableSort -Workbook $Workbook -SheetName $costReg.sheet `
            -TableName $costReg.table_name -KeyColumnIndex $descriptionOrdinal -Order 2
        $idsAfter = @(Get-IdColumnValues -Workbook $Workbook -Info $costReg)
        # THE ORDER ACTUALLY CHANGED. "Sort was called" is not evidence.
        $null = Add-Check $list 'the physical permanent-ID order ACTUALLY changed' `
            (($idsBefore -join ',') -cne ($idsAfter -join ',')) `
            ("before " + ($idsBefore -join ',') + " / after " + ($idsAfter -join ','))
        $null = Add-Check $list 'the same identifiers are present, only reordered' `
            (((@($idsBefore | Sort-Object)) -join ',') -eq ((@($idsAfter | Sort-Object)) -join ',')) `
            ("before " + ($idsBefore -join ',') + " / after " + ($idsAfter -join ','))
        & $probe 'Cost Lines physically re-sorted (real ListObject sort, order changed)' $reorderDigest
        Invoke-TableSort -Workbook $Workbook -SheetName $costReg.sheet `
            -TableName $costReg.table_name -KeyColumnIndex $descriptionOrdinal -Order 1
        $idsRestored = @(Get-IdColumnValues -Workbook $Workbook -Info $costReg)
        $null = Add-Check $list 'the original physical order was restored' `
            (($idsRestored -join ',') -ceq ($idsBefore -join ',')) `
            ("restored " + ($idsRestored -join ','))
        & $probe 'Cost Lines re-sorted back' $reorderDigest

        # ---- probe 3: Selected Confidence Level ----------------------------
        $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
            -Inspection $Inspection -Model $baseCase.model
        $Excel.Run('PCCM_Calculate') | Out-Null
        $digest = [string]$Excel.Run('PCCM_CalculationFingerprint')
        $confidence = $Inspection.inputs.selected_confidence_level
        $wasConfidence = [string](Get-NamedValue -Workbook $Workbook -DefinedName $confidence.defined_name)
        Set-NamedValueText -Workbook $Workbook -DefinedName $confidence.defined_name -Text 'P90'
        & $probe 'Selected Confidence Level changed' $digest
        Set-NamedValueText -Workbook $Workbook -DefinedName $confidence.defined_name -Text $wasConfidence
        & $probe 'Selected Confidence Level restored' $digest

        # ---- probe 4: an UNREFERENCED FX assumption -------------------------
        # Referenced-only resolution means a currency no driver uses is never
        # consulted, so it cannot change the digest - and it must not.
        $fx = $Inspection.input_tables.fx_rates
        $fxRows = Get-TableRowCount -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name
        Add-BlankTableRow -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name
        Set-TableCell -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name `
            -RowIndex ($fxRows + 1) -ColumnIndex 1 -Value 'ZZZ'
        Set-TableCell -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name `
            -RowIndex ($fxRows + 1) -ColumnIndex 2 -Value ([double]3.75)
        & $probe 'an UNREFERENCED FX assumption added' $digest
        Remove-TableRow -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name `
            -RowIndex ($fxRows + 1)
        & $probe 'the unreferenced FX assumption removed' $digest

        # Leave the shared baseline exactly as the scenarios below expect it.
        $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
            -Inspection $Inspection -Model $baseCase.model
        $Excel.Run('PCCM_Calculate') | Out-Null
        $establishedFingerprint = [string]$Excel.Run('PCCM_CalculationFingerprint')
        $establishedState = Get-CalcScalarBlock -Workbook $Workbook -Inspection $Inspection -Block 'calc_state'

        Add-Phase5Result 'P5-NS' `
            'Non-staleness, four INDEPENDENT probes: description, real multi-row reorder, confidence level, unreferenced FX' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Phase5Result 'P5-NS' 'Non-staleness proofs' 'FAIL' (Format-Phase5Err $_)
    }

    # -------------------------------------------------------------------
    # P5-S3 / P5-S4 / P5-KP. An invalid input, before and after Calculate
    # -------------------------------------------------------------------
    $refusalDetail = ''
    $invalidWeight = $null
    try {
        $before = Get-Phase5Snapshot -Workbook $Workbook -Inspection $Inspection
        $grid = $null
        foreach ($candidate in @($Manifest.grids)) {
            if ($candidate.key -eq 'cost_profiling') { $grid = $candidate }
        }
        $fixed = @($grid.fixed_columns).Count
        $body = @(Get-TableBody -Workbook $Workbook -SheetName $grid.sheet -TableName $grid.table_name)
        $invalidWeight = $body[0][$fixed]

        # ROW 3: invalid current input, and NOTHING is calculated.
        $list = New-Checklist
        # The profile no longer sums to 100%, which the accepted checker refuses.
        Set-TableCell -Workbook $Workbook -SheetName $grid.sheet -TableName $grid.table_name `
            -RowIndex 1 -ColumnIndex ($fixed + 1) -Value ([double]0.99)
        $row3 = Add-StatusRowChecks -List $list -Excel $Excel -Row 'row 3' `
            -ExpectedStatus 'INVALID' -ExpectedAttempt 'SUCCESS' -DetailRule 'blank' `
            -ExpectedFingerprint $establishedFingerprint
        $null = Add-Check $list 'row 3: the current input fingerprint is blank while the model is invalid' `
            ([string]::IsNullOrEmpty($row3.Current)) ("got '" + $row3.Current + "'")
        $afterRow3 = Get-Phase5Snapshot -Workbook $Workbook -Inspection $Inspection
        Add-SnapshotUnchangedChecks -List $list -Before $before -After $afterRow3 -Label 'row 3' `
            -SuccessFields $successRecordFields
        Add-Phase5Result 'P5-S3' 'Status row 3: invalid current input, no Calculate attempted' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)

        # ROW 4 / P5-KP: the SAME invalid input, WITH Calculate.
        $list = New-Checklist
        $Excel.Run('PCCM_Calculate') | Out-Null
        $row4 = Add-StatusRowChecks -List $list -Excel $Excel -Row 'row 4' `
            -ExpectedStatus 'INVALID' -ExpectedAttempt 'REFUSED' -DetailRule 'specific' `
            -ExpectedFingerprint $establishedFingerprint
        $refusalDetail = $row4.Detail
        $afterRow4 = Get-Phase5Snapshot -Workbook $Workbook -Inspection $Inspection
        Add-SnapshotUnchangedChecks -List $list -Before $before -After $afterRow4 -Label 'row 4' `
            -SuccessFields $successRecordFields
        # AND THE OTHER GROUP CHANGED, as it must. C17:C20 is the refusal record;
        # asserting it unchanged would assert that the refusal was never written.
        $null = Add-Check $list 'row 4: C17 CHANGED to REFUSED' `
            ([string]$afterRow4.State['last_attempt_result'] -eq 'REFUSED') `
            ("was " + (Format-CalcValue $before.State['last_attempt_result']) + `
             ", now " + (Format-CalcValue $afterRow4.State['last_attempt_result']))
        $null = Add-Check $list 'row 4: C18 CHANGED to a specific refusal detail' `
            (-not [string]::IsNullOrWhiteSpace([string]$afterRow4.State['last_attempt_detail']))
        $null = Add-Check $list 'row 4: C19 is the freshly derived status' `
            ([string]$afterRow4.State['calculation_status'] -eq 'INVALID')
        $null = Add-Check $list 'row 4: C20 carries a status-evaluation timestamp' `
            (-not (Test-CalcBlank -Actual $afterRow4.State['status_evaluated_at']))
        Add-Phase5Result 'P5-S4' 'Status row 4: invalid current input + PCCM_Calculate' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)

        Add-Phase5Result 'P5-KP' 'Refusal preserves the prior successful snapshot (C13:C16, C23:C32, five tables)' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            'the two mutation groups are compared separately; see P5-S4'

        # ROW 5 / P5-RC: restore the input EXACTLY, and do NOT calculate.
        $list = New-Checklist
        Set-TableCell -Workbook $Workbook -SheetName $grid.sheet -TableName $grid.table_name `
            -RowIndex 1 -ColumnIndex ($fixed + 1) -Value ([double]$invalidWeight)
        $row5 = Add-StatusRowChecks -List $list -Excel $Excel -Row 'row 5' `
            -ExpectedStatus 'CURRENT' -ExpectedAttempt 'REFUSED' -DetailRule 'specific' `
            -ExpectedFingerprint $establishedFingerprint
        # THE DISAGREEMENT IS REQUIRED. The status axis says the current inputs
        # match the stored snapshot; the attempt axis still records the refusal
        # that really happened. Neither is corrected into the other.
        $null = Add-Check $list 'row 5: the refusal detail is STILL readable, unchanged' `
            ($row5.Detail -ceq $refusalDetail) ("was '" + $refusalDetail + "', now '" + $row5.Detail + "'")
        $null = Add-Check $list 'row 5: CURRENT and a historical REFUSED coexist by design' `
            (($row5.Status -eq 'CURRENT') -and ($row5.Attempt -eq 'REFUSED'))
        $null = Add-Check $list 'row 5: the current input fingerprint is the stored one again' `
            ($row5.Current -ceq $establishedFingerprint)
        $afterRow5 = Get-Phase5Snapshot -Workbook $Workbook -Inspection $Inspection
        Add-SnapshotUnchangedChecks -List $list -Before $before -After $afterRow5 -Label 'row 5' `
            -SuccessFields $successRecordFields
        Add-Phase5Result 'P5-S5' 'Status row 5: exact restoration of the prior input, no Calculate' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
        Add-Phase5Result 'P5-RC' 'Revert to CURRENT without calculating (plan case 32)' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            'CURRENT status with a historical REFUSED attempt; see P5-S5'
    } catch {
        foreach ($id in 'P5-S3', 'P5-S4', 'P5-S5', 'P5-KP', 'P5-RC') {
            Add-Phase5Result $id 'Refusal / revert sequence' 'FAIL' (Format-Phase5Err $_)
        }
    }

    # ===================================================================
    # P5-FA / P5-FC / P5-S6. ROLLBACK AT BOTH LOCKED FAILPOINT BOUNDARIES
    #
    # Through the accepted Phase-4 injection mechanism -
    # PCCM_AutomationBegin(confirm, failpointName) and FailPointCheck - and no
    # other. No second injection system is created, and no production source is
    # touched to make the failure happen.
    #
    # The two boundaries are genuinely different:
    #   Phase5AnalyticalWrite  fires AFTER analytical blocks have been mutated
    #                          and BEFORE the success commit
    #   Phase5SuccessCommit    fires at the FINAL C13:C20 assignment, inside
    #                          WriteSuccessCommit, one statement before
    #                          Range(CALC_STATE_VALUE_RANGE).Value2 = block
    # ===================================================================
    function Invoke-Phase5RollbackScenario {
        param($Excel, $Workbook, $Manifest, $Inspection, $Cases, $BaseCase,
              [string]$ScenarioId, [string]$Failpoint, [string]$Title,
              $SuccessFields, $AttemptFields)
        $list = New-Checklist
        try {
            # 1. A known-good snapshot to roll back TO.
            $Excel.Run('PCCM_AutomationEnd') | Out-Null
            $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
            $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
                -Inspection $Inspection -Model $BaseCase.model
            $Excel.Run('PCCM_Calculate') | Out-Null
            $null = Add-Check $list 'a successful snapshot was established first' `
                ([string]$Excel.Run('PCCM_CalculationAttemptResult') -eq 'SUCCESS')
            $storedBefore = [string]$Excel.Run('PCCM_CalculationFingerprint')
            $before = Get-Phase5Snapshot -Workbook $Workbook -Inspection $Inspection

            # 2. CHANGE A VALID FINGERPRINTED INPUT, so the model is genuinely
            #    STALE. The rolled-back state must then report STALE, never
            #    CURRENT: a FAILED attempt may not choose the derived status.
            $grid = $null
            foreach ($candidate in @($Manifest.grids)) {
                if ($candidate.key -eq 'cost_profiling') { $grid = $candidate }
            }
            $fixed = @($grid.fixed_columns).Count
            $weights = @(@($BaseCase.model.cost_lines)[0].profile_weights)
            Set-TableCell -Workbook $Workbook -SheetName $grid.sheet -TableName $grid.table_name `
                -RowIndex 1 -ColumnIndex ($fixed + 1) -Value ([double]$weights[1])
            Set-TableCell -Workbook $Workbook -SheetName $grid.sheet -TableName $grid.table_name `
                -RowIndex 1 -ColumnIndex ($fixed + 2) -Value ([double]$weights[0])
            $null = Add-Check $list 'the changed model is STALE before the injected failure' `
                ([string]$Excel.Run('PCCM_CalculationStatus') -eq 'STALE')

            # 2b. ESTABLISH A DELIBERATELY NON-DEFAULT CALLER STATE.
            #
            # Asserting EnableEvents/ScreenUpdating True and Calculation
            # Automatic afterwards proves only that the application happens to be
            # in convenient defaults - which it would be even if FinishOperation
            # restored nothing at all. The accepted Phase-4 scenario S already
            # rejected that pattern: it establishes unusual caller state and
            # proves EXACT restoration. Gate B does the same for the calculation
            # operation, with a StatusBar sentinel unique to this failpoint so a
            # value carried over from another scenario cannot satisfy it.
            $sentinel = 'PCCM Phase-5 rollback sentinel ' + $Failpoint
            $Excel.ScreenUpdating = $false
            $Excel.EnableEvents = $false
            $Excel.DisplayAlerts = $false
            $Excel.Calculation = -4135          # xlCalculationManual
            $Excel.StatusBar = $sentinel
            $callerState = [pscustomobject]@{
                ScreenUpdating = $Excel.ScreenUpdating
                EnableEvents   = $Excel.EnableEvents
                DisplayAlerts  = $Excel.DisplayAlerts
                Calculation    = $Excel.Calculation
                StatusBar      = $Excel.StatusBar
            }
            $null = Add-Check $list 'a NON-DEFAULT caller state was established before the operation' `
                (($callerState.ScreenUpdating -eq $false) -and ($callerState.EnableEvents -eq $false) -and
                 ($callerState.DisplayAlerts -eq $false) -and ([int]$callerState.Calculation -eq -4135) -and
                 ([string]$callerState.StatusBar -eq $sentinel)) `
                ("ScreenUpdating=" + [string]$callerState.ScreenUpdating +
                 " EnableEvents=" + [string]$callerState.EnableEvents +
                 " DisplayAlerts=" + [string]$callerState.DisplayAlerts +
                 " Calculation=" + [string]$callerState.Calculation +
                 " StatusBar='" + [string]$callerState.StatusBar + "'")

            # 3. ARM THE FAILPOINT through the accepted Phase-4 mechanism.
            $Excel.Run('PCCM_AutomationEnd') | Out-Null
            $Excel.Run('PCCM_AutomationBegin', $true, $Failpoint) | Out-Null
            $Excel.Run('PCCM_Calculate') | Out-Null
            $invocation = [string]$Excel.Run('PCCM_AutomationResult')
            $null = Add-Check $list ('the injected failure at ' + $Failpoint + ' was reported') `
                ($invocation -like 'FAIL|*') $invocation

            # 4. THE ATTEMPT AXIS.
            $attempt = [string]$Excel.Run('PCCM_CalculationAttemptResult')
            $detail = [string]$Excel.Run('PCCM_CalculationAttemptDetail')
            $null = Add-Check $list 'C17 = FAILED' ($attempt -eq 'FAILED') ("got '" + $attempt + "'")
            $null = Add-Check $list 'C18 carries a specific failure detail' `
                (-not [string]::IsNullOrWhiteSpace($detail)) $detail
            $after = Get-Phase5Snapshot -Workbook $Workbook -Inspection $Inspection
            $null = Add-Check $list 'C19 is a freshly DERIVED status, not the attempt result' `
                (@('NOT CALCULATED', 'CURRENT', 'STALE', 'INVALID') -contains `
                 [string]$after.State['calculation_status']) `
                ("got '" + [string]$after.State['calculation_status'] + "'")
            $null = Add-Check $list 'C19 is not FAILED: an attempt result may never be a status' `
                ([string]$after.State['calculation_status'] -ne 'FAILED')
            $null = Add-Check $list 'C20 carries a fresh evaluation timestamp' `
                (-not (Test-CalcBlank -Actual $after.State['status_evaluated_at']))

            # 5. THE DERIVED STATUS FOR A CHANGED VALID INPUT IS STALE.
            $status = [string]$Excel.Run('PCCM_CalculationStatus')
            $null = Add-Check $list 'PCCM_CalculationStatus() = STALE, not CURRENT' `
                ($status -eq 'STALE') ("got '" + $status + "'")

            # 6. FULL LOGICAL ROLLBACK: C13:C16, C23:C32 and all five tables are
            #    the previous successful snapshot EXACTLY.
            Add-SnapshotUnchangedChecks -List $list -Before $before -After $after `
                -Label 'rollback' -SuccessFields $SuccessFields
            $null = Add-Check $list 'the stored fingerprint is the previous successful one' `
                ([string]$Excel.Run('PCCM_CalculationFingerprint') -ceq $storedBefore)

            # 7. NO MIXED OLD/NEW ANALYTICAL STATE. The tables were compared row
            #    for row above; this states the claim in its own right so a
            #    reviewer sees it asserted rather than implied.
            $mixed = $false
            foreach ($key in $before.Tables.Keys) {
                $was = @($before.Tables[$key]); $now = @($after.Tables[$key])
                if ($was.Count -ne $now.Count) { $mixed = $true; continue }
                for ($i = 0; $i -lt $was.Count; $i++) {
                    $wasRow = @($was[$i]); $nowRow = @($now[$i])
                    if ($wasRow.Count -ne $nowRow.Count) { $mixed = $true; continue }
                    for ($c = 0; $c -lt $wasRow.Count; $c++) {
                        if (-not (Test-Phase5ExactValue -Actual $nowRow[$c] -Expected $wasRow[$c])) {
                            $mixed = $true
                        }
                    }
                }
            }
            $null = Add-Check $list 'no mixed old/new analytical state survived the rollback' (-not $mixed)

            # 8. THE CALLER'S OWN STATE IS RESTORED, EXACTLY.
            #
            # Compared against what was CAPTURED, not against Excel's defaults,
            # and asserted BEFORE the harness normalises anything for the
            # scenarios that follow. All five properties, every time, in both
            # failpoint scenarios - neither inherits the other's proof.
            $null = Add-Check $list 'ScreenUpdating was restored to the CAPTURED caller value' `
                ($Excel.ScreenUpdating -eq $callerState.ScreenUpdating) `
                ("captured " + [string]$callerState.ScreenUpdating + ", now " + [string]$Excel.ScreenUpdating)
            $null = Add-Check $list 'EnableEvents was restored to the CAPTURED caller value' `
                ($Excel.EnableEvents -eq $callerState.EnableEvents) `
                ("captured " + [string]$callerState.EnableEvents + ", now " + [string]$Excel.EnableEvents)
            $null = Add-Check $list 'DisplayAlerts was restored to the CAPTURED caller value' `
                ($Excel.DisplayAlerts -eq $callerState.DisplayAlerts) `
                ("captured " + [string]$callerState.DisplayAlerts + ", now " + [string]$Excel.DisplayAlerts)
            $null = Add-Check $list 'Calculation was restored to the CAPTURED caller value' `
                ([int]$Excel.Calculation -eq [int]$callerState.Calculation) `
                ("captured " + [string]$callerState.Calculation + ", now " + [string]$Excel.Calculation)
            $null = Add-Check $list 'StatusBar was restored to the CAPTURED sentinel' `
                ([string]$Excel.StatusBar -eq [string]$callerState.StatusBar) `
                ("captured '" + [string]$callerState.StatusBar + "', now '" + [string]$Excel.StatusBar + "'")
            $null = Add-Check $list 'the restored state is NOT merely Excel default state' `
                (([int]$callerState.Calculation -ne -4105) -and ($callerState.EnableEvents -eq $false)) `
                'the captured state must differ from the defaults, or the proof is vacuous'

            # 8b. ONLY NOW may the harness normalise for the scenarios that follow.
            $Excel.ScreenUpdating = $true
            $Excel.EnableEvents = $true
            $Excel.DisplayAlerts = $false
            $Excel.Calculation = -4105          # xlCalculationAutomatic
            $Excel.StatusBar = $false

            # 9. Disarm, and prove the model still calculates afterwards: a
            #    rollback that left the workbook unusable would not be a rollback.
            $Excel.Run('PCCM_AutomationEnd') | Out-Null
            $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
            $Excel.Run('PCCM_Calculate') | Out-Null
            $null = Add-Check $list 'the model calculates again once the failpoint is disarmed' `
                ([string]$Excel.Run('PCCM_CalculationAttemptResult') -eq 'SUCCESS')
        } catch {
            $null = Add-Check $list 'the rollback scenario ran to completion' $false (Format-Phase5Err $_)
        }
        Add-Result $ScenarioId $Title $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
            (Format-Checklist $list)
        return (Test-ChecklistOk $list)
    }

    $analyticalOk = Invoke-Phase5RollbackScenario -Excel $Excel -Workbook $Workbook `
        -Manifest $Manifest -Inspection $Inspection -Cases $Cases -BaseCase $baseCase `
        -ScenarioId 'P5-FA' -Failpoint $failpoints.AnalyticalWrite `
        -Title 'Rollback at the ANALYTICAL-WRITE boundary (plan case 33)' `
        -SuccessFields $successRecordFields -AttemptFields $attemptFields

    $commitOk = Invoke-Phase5RollbackScenario -Excel $Excel -Workbook $Workbook `
        -Manifest $Manifest -Inspection $Inspection -Cases $Cases -BaseCase $baseCase `
        -ScenarioId 'P5-FC' -Failpoint $failpoints.SuccessCommit `
        -Title 'Rollback at the C13:C20 COMMIT boundary (plan case 37)' `
        -SuccessFields $successRecordFields -AttemptFields $attemptFields

    # Status row 6 IS the injected-failure row, and it is recorded as its own
    # result so the six-row matrix is complete in the report rather than implied
    # by two rollback scenarios.
    Add-Phase5Result 'P5-S6' 'Status row 6: injected write failure on valid changed inputs' `
        $(if ($analyticalOk -and $commitOk) { 'PASS' } else { 'FAIL' }) `
        ('STALE / FAILED / specific detail / previous snapshot restored, at both locked boundaries: ' +
         $failpoints.AnalyticalWrite + ' (P5-FA) and ' + $failpoints.SuccessCommit + ' (P5-FC)')

    # -------------------------------------------------------------------
    # P5-AX. The two axes are read separately and never conflated
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $Excel.Run('PCCM_AutomationEnd') | Out-Null
        $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
        $null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
            -Inspection $Inspection -Model $baseCase.model
        $Excel.Run('PCCM_Calculate') | Out-Null

        # PCCM_Calculate publishes through the harness-aware Announce surface, so
        # a result is RECORDED for automation and no dialog blocks the run. A
        # MsgBox would have hung this call, not failed it - reaching the next
        # line at all is part of the evidence.
        $invocation = [string]$Excel.Run('PCCM_AutomationResult')
        $null = Add-Check $list 'PCCM_Calculate recorded an invocation result for automation' `
            (-not [string]::IsNullOrEmpty($invocation)) $invocation
        $null = Add-Check $list 'the invocation axis reports OK for a clean commit' `
            ($invocation -like 'OK|*') $invocation
        $null = Add-Check $list 'no dialog blocked automation (the call returned)' $true

        # THE TWO AXES ARE READ SEPARATELY. calc_state carries the calculation
        # attempt; PCCM_AutomationResult carries the invocation. They are allowed
        # to disagree - a committed calculation whose application cleanup later
        # failed reports SUCCESS on one axis and FAIL on the other, by design -
        # so the harness must never read one and report the other.
        $attempt = [string]$Excel.Run('PCCM_CalculationAttemptResult')
        $null = Add-Check $list 'the calculation attempt axis is read from calc_state' `
            ($attempt -eq 'SUCCESS') $attempt
        $null = Add-Check $list 'the two axes are distinct values, read through distinct endpoints' `
            ($invocation -ne $attempt) ("invocation '" + $invocation + "' / attempt '" + $attempt + "'")
        # The disagreement itself is NOT forced: nothing here makes application
        # state restoration fail, because the accepted harness has no safe way to
        # induce it. What is proved is that both axes are readable independently.
        Add-Note ('P5-AX: a committed-SUCCESS / cleanup-FAIL disagreement was not induced; ' +
                  'the accepted harness has no safe way to make FinishOperation fail, and ' +
                  'forcing it would prove the forcing. Both axes are read separately.')
        Add-Phase5Result 'P5-AX' 'Automation/invocation axis read separately from the calculation attempt' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Phase5Result 'P5-AX' 'Automation/invocation axis' 'FAIL' (Format-Phase5Err $_)
    }

    # -------------------------------------------------------------------
    # The coverage ledger, reported
    # -------------------------------------------------------------------
    $mapped = @()
    foreach ($id in $ledger.Keys) { $mapped += ($id + ' -> ' + (@($ledger[$id]) -join ', ')) }
    Add-Note ('Phase-5 coverage ledger (' + $ledger.Count + ' plan cases): ' + ($mapped -join ' | '))
}
