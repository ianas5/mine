#!/usr/bin/env python3
"""P10 W3 / PROBE MUTATION CONTROLS.

THE PROBE IS THE INSTRUMENT A PRODUCTION DECISION WILL REST ON, so a control
over it that cannot fail is worse than none: the next thing that happens after
this probe returns "BLOCKED" is that four accepted modules and an accepted
contract get edited.

Every control below breaks ONE thing - the `.Count` shape that ended Probe Run
1, a normalisation, a verdict branch, an evidence read, the strict mode - reruns
the WHOLE conformance battery against the damaged copy, and requires a NAMED
control among the refusers.

Nothing here writes to the repository. Damaged copies live in memory.

Runs standalone or under pytest.
"""
from __future__ import annotations

import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import pytest  # noqa: E402

import test_phase10_table_structure_under_protection as conformance  # noqa: E402


def _tests() -> list[str]:
    names = sorted(name for name in dir(conformance) if name.startswith("test_"))
    assert len(names) >= 30, names
    return names


def _run_battery() -> list[str]:
    refused: list[str] = []
    for name in _tests():
        try:
            getattr(conformance, name)()
        except BaseException:  # noqa: BLE001 - any refusal counts
            refused.append(name)
    return refused


def _probe_mutation(expected: str, before: str, after: str) -> None:
    """Damage the probe's source, and prove the anchor still matched.

    NOT AN AssertionError WHEN THE ANCHOR MISSES. A no-op mutation that raised
    one would be indistinguishable from the refusal it is meant to provoke.
    """
    original = conformance._probe()
    damaged = original.replace(before, after, 1)
    if damaged == original:
        raise RuntimeError(
            f"the mutation changed nothing: {before[:60]!r} is no longer in the probe")
    saved = dict(conformance._MEMO)
    conformance._MEMO["probe"] = damaged
    conformance._MEMO.pop("probe_code", None)
    try:
        refused = _run_battery()
    finally:
        conformance._MEMO.clear()
        conformance._MEMO.update(saved)
        conformance._MEMO.pop("probe_code", None)
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


# ===========================================================================
# A. THE PROBE RUN 1 DEFECT
# ===========================================================================
def test_01_restoring_the_count_over_a_com_collection_is_rejected() -> None:
    """THE EXACT SHAPE THAT ENDED PROBE RUN 1, put back."""
    _probe_mutation(
        "test_40",
        "        $sheets = $Workbook.Worksheets\n"
        "        foreach ($sheet in @($sheets)) {",
        "        $sheets = $Workbook.Worksheets\n"
        "        for ($index = 1; $index -le $sheets.Count; $index++) {\n"
        "            $sheet = $sheets.Item($index)")


def test_02_reading_count_off_a_list_columns_collection_is_rejected() -> None:
    """THE SAME SHAPE ONE LAYER DOWN. It was in the first draft too, and only
    the worksheet one happened to be reached first."""
    _probe_mutation(
        "test_40",
        "        $columns = Measure-ProbeCollection -Collection $columnCollection `\n"
        "            -Label 'ListColumn' -Where $where",
        "        $columns = [int]$columnCollection.Count")


def test_03_indexing_a_com_collection_by_position_is_rejected() -> None:
    _probe_mutation(
        "test_41",
        "    foreach ($item in @($Collection)) {",
        "    for ($index = 1; $index -le 10; $index++) {\n"
        "        $item = $Collection.Item($index)")


def test_04_counting_an_absent_collection_as_zero_is_rejected() -> None:
    """A FAKE ZERO WOULD BE A SHAPE CHANGE THE PROBE INVENTED, and shape change
    is what the verdict turns on."""
    _probe_mutation(
        "test_42",
        "    if ($null -eq $Collection) {\n"
        "        throw ($Where + ': expected a ' + $Label + ' collection and got nothing')\n"
        "    }",
        "    if ($null -eq $Collection) { return 0 }")


def test_05_accepting_a_workbook_with_no_sheets_is_rejected() -> None:
    _probe_mutation(
        "test_42",
        "        throw ($Where + ': the workbook enumerated no worksheets, so its protection ' +\n"
        "               'state could not be established')",
        "        $protectedNames = @()")


def test_06_turning_strict_mode_off_is_rejected() -> None:
    """THE FIX IS NOT TO STOP CHECKING. StrictMode is what turned a silent $null
    into a loud failure."""
    _probe_mutation("test_43", "Set-StrictMode -Version 2.0", "Set-StrictMode -Off")


def test_07_swallowing_a_probe_failure_is_rejected() -> None:
    _probe_mutation(
        "test_43",
        "        } catch {\n"
        "            $script:ProbePath = ''",
        "        } catch { }\n"
        "        if ($false) {\n"
        "            $script:ProbePath = ''")


def test_08_dropping_the_stage_cursor_is_rejected() -> None:
    """WITHOUT IT THE NEXT FAILURE SAYS ONLY WHAT THE EXCEPTION SAYS - which is
    what made Probe Run 1 take a round trip to diagnose."""
    _probe_mutation("test_44", "function Format-ProbeFailure {",
                    "function Build-ProbeFailureNote {")


def test_09_dropping_a_stage_boundary_is_rejected() -> None:
    _probe_mutation(
        "test_44",
        "    Set-ProbeStage -Stage 'verdict' -Action 'weighing the outcomes'",
        "")


def test_10_letting_the_diagnostics_touch_the_workbook_is_rejected() -> None:
    """A DIAGNOSTIC THAT CHANGED THE WORKBOOK WOULD CHANGE THE ANSWER."""
    _probe_mutation(
        "test_45",
        "    $script:ProbeCursor.Stage = $Stage",
        "    $script:ProbeCursor.Stage = $Stage\n"
        "    $null = $Workbook.Saved")


# ===========================================================================
# B. VERDICT DISCIPLINE
# ===========================================================================
def test_20_letting_a_probe_failure_become_a_verdict_is_rejected() -> None:
    """THE ONE THING THAT MUST NEVER HAPPEN. A probe that broke would otherwise
    be quoted as evidence about production."""
    _probe_mutation(
        "test_50",
        "    $verdictReason = ('the probe itself failed in stage ' + [string]$script:ProbeCursor.Stage +",
        "    $verdict = 'PRODUCTION IS BLOCKED BY PROTECTION'\n"
        "    $verdictReason = ('the probe itself failed in stage ' + [string]$script:ProbeCursor.Stage +")


def test_21_starting_from_a_conclusion_is_rejected() -> None:
    _probe_mutation(
        "test_50",
        "$verdict = 'INCONCLUSIVE'\n$verdictReason",
        "$verdict = 'PRODUCTION IS FINE UNDER PROTECTION'\n$verdictReason")


def test_22_calling_it_fine_without_a_structural_effect_is_rejected() -> None:
    """AN ANNOUNCEMENT IS NOT A STRUCTURAL OPERATION. A command that said OK and
    reshaped nothing has shown nothing about whether the operation is
    permitted."""
    _probe_mutation(
        "test_52",
        "        } elseif (@($structural).Count -lt 1) {",
        "        } elseif ($false) {")


def test_23_calling_it_fine_after_protection_was_lost_is_rejected() -> None:
    """IF PROTECTION WAS NOT IN FORCE, THE COMMANDS WERE NOT A TEST OF PROTECTED
    BEHAVIOUR."""
    _probe_mutation(
        "test_52",
        "        } elseif (@($lostProtection).Count -gt 0) {",
        "        } elseif ($false) {")


def test_24_calling_it_blocked_without_a_failed_endpoint_is_rejected() -> None:
    _probe_mutation(
        "test_51",
        "        if (@($notSucceeded).Count -gt 0) {",
        "        if ($true) {")


def test_25_conflating_a_raise_with_a_refusal_is_rejected() -> None:
    """A PRODUCTION REFUSAL AND AN EXCEL RUNTIME FAILURE ARE DIFFERENT FACTS."""
    _probe_mutation(
        "test_54",
        "    elseif (-not $announced) { $outcome = 'RAISED' }",
        "    elseif (-not $announced) { $outcome = 'REFUSED' }")


def test_26_letting_the_locked_cell_control_decide_the_question_is_rejected() -> None:
    """THE CONTROL PROVES A DIFFERENT CAPABILITY, and must not be able to answer
    this one."""
    _probe_mutation(
        "test_57",
        "        if (@($notSucceeded).Count -gt 0) {",
        "        if ($controlWorked -and (@($notSucceeded).Count -gt 0)) {")


def test_27_dropping_the_locked_cell_control_is_rejected() -> None:
    _probe_mutation(
        "test_57",
        "        Write-ProbeLine 'CONTROL - CAN CODE WRITE A VALUE TO A LOCKED CELL?'",
        "        Write-ProbeLine 'skipped'")


def test_28_an_inconclusive_run_exiting_zero_is_rejected() -> None:
    _probe_mutation(
        "test_56",
        "if ($verdict -eq 'INCONCLUSIVE') { exit 2 }",
        "if ($false) { exit 2 }")


# ===========================================================================
# C. THE EVIDENCE
# ===========================================================================
def test_30_dropping_the_before_protection_read_is_rejected() -> None:
    _probe_mutation(
        "test_53",
        "    $protectionBefore = Get-ProbeProtectionState -Workbook $Workbook "
        "-Where ('before ' + $Endpoint)",
        "    $protectionBefore = $null")


def test_31_dropping_the_shape_comparison_is_rejected() -> None:
    _probe_mutation(
        "test_53",
        "    $shapesAfter = Get-ProbeAllShapes -Resolution $Resolution",
        "    $shapesAfter = $shapesBefore")


def test_32_typing_the_watched_tables_into_the_probe_is_rejected() -> None:
    """A GRID THE MANIFEST GAINED AND THE PROBE DID NOT would be a structural
    effect nobody looked for."""
    _probe_mutation(
        "test_55",
        "    foreach ($grid in @($Manifest.grids)) {",
        "    foreach ($grid in @([pscustomobject]@{ key = 'x'; sheet = 'Inflation'; "
        "table_name = 'tblInflation' })) {")


def test_33_skipping_a_required_endpoint_is_rejected() -> None:
    _probe_mutation(
        "test_20",
        "            -Endpoint 'PCCM_Calculate' -Resolution $resolution",
        "            -Endpoint 'PCCM_RunSimulation' -Resolution $resolution")


def test_34_reaching_around_a_production_endpoint_is_rejected() -> None:
    """NO BYPASS. A probe that released protection to make the commands work
    would answer FINE every time, about a workbook nobody ships."""
    _probe_mutation(
        "test_55b",
        "    $protectionInForce = ([bool]",
        "    $null = $excel.Run('ProtectionRelease')\n"
        "    $protectionInForce = ([bool]")


def test_34b_performing_the_structural_operation_directly_is_rejected() -> None:
    _probe_mutation(
        "test_55b",
        "        $null = $outcomes.Add($timeline)",
        "        $null = $outcomes.Add($timeline)\n"
        "        $wb.Worksheets.Item('Inflation').ListObjects.Item(1).ListRows.Add()")


def test_35_saving_the_workbook_is_rejected() -> None:
    _probe_mutation(
        "test_24",
        "            try { $wb.Close($false); $rel.WorkbookClosed = $true }",
        "            try { $wb.Save(); $wb.Close($false); $rel.WorkbookClosed = $true }")


def test_36_killing_excel_instead_of_quitting_it_is_rejected() -> None:
    _probe_mutation(
        "test_24",
        "            try { $excel.Quit(); $rel.QuitCalled = $true }",
        "            try { Stop-Process -Name EXCEL -Force; $rel.QuitCalled = $true }")


def test_37_timing_something_in_the_probe_is_rejected() -> None:
    """IT IS NOT A BENCHMARK. Nothing it produces may be read as one."""
    _probe_mutation(
        "test_23",
        "    $raised = ''\n    $result = ''",
        "    $raised = ''\n    $result = ''\n"
        "    $watch = [System.Diagnostics.Stopwatch]::StartNew()")


# ===========================================================================
# D. THE PROBE RUN 2 DEFECTS
# ===========================================================================
def test_40_restoring_the_per_read_worksheet_lookup_is_rejected() -> None:
    """THE SHAPE THAT PRODUCED DISP_E_BADINDEX: acquire the Worksheets
    collection inside the shape read, look a sheet up in it, release it, and do
    it again five tables later."""
    _probe_mutation(
        "test_61",
        "    $lo = $Target.ListObject",
        "    $sheets = $Workbook.Worksheets\n"
        "    $ws = $sheets.Item($sheetName)\n"
        "    $lo = $Target.ListObject")


def test_41_falling_back_to_a_positional_sheet_index_is_rejected() -> None:
    """AN UNRESOLVABLE SHEET IS A PROBE FAILURE, never a substitution."""
    _probe_mutation(
        "test_62",
        "            throw ('the workbook has no worksheet named ' + [char]39 + $sheetName + [char]39 +",
        "            $ws = $sheets.Item(1)\n"
        "            $null = ('the workbook has no worksheet named ' + [char]39 + $sheetName + [char]39 +")


def test_42_swapping_the_sheet_and_table_identifiers_is_rejected() -> None:
    """NEITHER IS DERIVED FROM THE OTHER, and neither stands in for the other.
    Swapping them resolves a table name as a worksheet, which produces exactly
    the DISP_E_BADINDEX Run 3 reported."""
    _probe_mutation(
        "test_82",
        "        $sheetName = Get-ProbeScalarString -InputObject $entry -Name 'Sheet' -Where $where\n"
        "        $tableName = Get-ProbeScalarString -InputObject $entry -Name 'Table' -Where $where",
        "        $sheetName = Get-ProbeScalarString -InputObject $entry -Name 'Table' -Where $where\n"
        "        $tableName = Get-ProbeScalarString -InputObject $entry -Name 'Sheet' -Where $where")


def test_43_using_the_codename_to_find_the_sheet_is_rejected() -> None:
    """A CodeName IS NOT A TAB NAME. Conflating them is how a lookup starts
    failing on a workbook whose tabs were renamed - or never starts working."""
    _probe_mutation(
        "test_63",
        "            $ws = $sheets.Item($sheetName)",
        "            $ws = $sheets.Item($ws.CodeName)")


def test_44_claiming_the_endpoint_was_invoked_before_application_run_is_rejected() -> None:
    """PROBE RUN 2 SAID "invoking the production entry point" while it was still
    collecting pre-command evidence, and the endpoint was never reached."""
    _probe_mutation(
        "test_64",
        "    $invoked = $false\n\n    Set-ProbeStage",
        "    $invoked = $true\n\n    Set-ProbeStage")


def test_45_marking_the_endpoint_invoked_in_two_places_is_rejected() -> None:
    _probe_mutation(
        "test_54",
        "        $invoked = $true",
        "        $invoked = $true\n        $invoked = $true")


def test_46_reordering_the_endpoint_evidence_is_rejected() -> None:
    _probe_mutation(
        "test_65",
        "    $shapesAfter = Get-ProbeAllShapes -Resolution $Resolution\n"
        "    Set-ProbeStage -Stage 'endpoint' -Action 'reading protection after the command' -Endpoint $Endpoint\n"
        "    $protectionAfter = Get-ProbeProtectionState -Workbook $Workbook -Where ('after ' + $Endpoint)",
        "    $protectionAfter = Get-ProbeProtectionState -Workbook $Workbook -Where ('after ' + $Endpoint)\n"
        "    $shapesAfter = Get-ProbeAllShapes -Resolution $Resolution")


def test_47_putting_the_write_and_the_restore_back_in_one_try_is_rejected() -> None:
    """THE EXACT SHAPE THAT PRINTED SUCCEEDED AND REFUSED IN ONE RUN."""
    _probe_mutation(
        "test_66",
        "        try { $cell.Value2 = $original } catch { $restoreRaised = (Format-Err $_) }",
        "        try { $cell.Value2 = $original } catch { $writeRaised = (Format-Err $_) }")


def test_48_writing_to_a_cell_that_was_not_proved_locked_is_rejected() -> None:
    """PROBE RUN 2's TARGET WAS AN EDITABLE INPUT. A user can type in that cell,
    so writing to it says nothing about UserInterfaceOnly."""
    _probe_mutation(
        "test_67",
        "        if (-not $cell.Locked) {",
        "        if ($false) {")


def test_49_skipping_the_restoration_check_is_rejected() -> None:
    _probe_mutation(
        "test_68",
        "        if ($restored -ne [string]$original) {",
        "        if ($false) {")


def test_50_rewriting_a_cleanup_failure_as_a_refusal_is_rejected() -> None:
    """A FAILED RESTORE IS NOT "protection blocks value writes". It is an
    untrustworthy instrument, and it is reported as one."""
    _probe_mutation(
        "test_69",
        "            return [pscustomobject]@{ Result = 'INCONCLUSIVE'; "
        "Detail = 'the original value could not be restored'; Lines = @($lines) }",
        "            return [pscustomobject]@{ Result = 'REFUSED'; "
        "Detail = 'the original value could not be restored'; Lines = @($lines) }")


def test_51_letting_an_untrustworthy_control_proceed_is_rejected() -> None:
    _probe_mutation(
        "test_70",
        "        if ($controlResult -eq 'INCONCLUSIVE') {",
        "        if ($false) {")


def test_52_accepting_a_blank_control_original_is_rejected() -> None:
    """A BLANK ORIGINAL IS WHAT THREW IN RUN 2, and it also makes 'restored
    exactly' unverifiable."""
    _probe_mutation(
        "test_68",
        "        if ([string]::IsNullOrWhiteSpace([string]$original)) {",
        "        if ($false) {")


# ===========================================================================
# E. THE PROBE RUN 3 COLLAPSE
# ===========================================================================
def test_60_restoring_the_unary_comma_producer_is_rejected() -> None:
    """THE EXACT IDIOM THAT COLLAPSED FIVE RECORDS INTO ONE."""
    _probe_mutation(
        "test_80",
        "    return $watched\n}",
        "    return ,@($watched)\n}")


def test_61_the_same_idiom_in_the_shape_delta_is_rejected() -> None:
    """THE INSTANCE THAT WOULD HAVE BEEN WORSE THAN AN ABORT: an empty change
    list arriving as a one-element array makes StructuralEffect true for every
    endpoint, and FINE reachable with no shape change at all."""
    _probe_mutation(
        "test_80",
        "    return $changes\n}",
        "    return ,@($changes)\n}")


def test_62_dropping_the_watched_shape_assertion_is_rejected() -> None:
    _probe_mutation(
        "test_81",
        "    $null = Assert-ProbeWatchedShape -Watched $watched -Manifest $manifest",
        "    $null = $watched")


def test_63_hard_coding_the_record_count_is_rejected() -> None:
    """THE COUNT IS THE MANIFEST'S. A literal five would stop being true the day
    a grid was added, and would stop being a check at all."""
    _probe_mutation(
        "test_81",
        "    $expected = (@($Manifest.registers).Count + @($Manifest.grids).Count)",
        "    $expected = 5")


def test_64_string_coercion_hiding_an_array_is_rejected() -> None:
    """`[string]` MUST NEVER BE THE THING THAT DISCOVERS a property holds five
    values: it turns a structural error into a plausible-looking name."""
    _probe_mutation(
        "test_82",
        "        $sheetName = Get-ProbeScalarString -InputObject $entry -Name 'Sheet' -Where $where",
        "        $sheetName = [string]$entry.Sheet")


def test_65_casting_before_the_array_test_is_rejected() -> None:
    _probe_mutation(
        "test_82",
        "    if ($value -is [System.Array] -or $value -is [System.Collections.IEnumerable] -and",
        "    if ($false -and $value -is [System.Collections.IEnumerable] -and")


def test_66_accepting_a_record_that_is_itself_an_array_is_rejected() -> None:
    _probe_mutation(
        "test_83",
        "        if ($record -is [System.Array]) {",
        "        if ($false) {")


def test_67_dropping_the_uniqueness_checks_is_rejected() -> None:
    _probe_mutation(
        "test_81",
        "        if ($keys -contains $key) { throw ($where + \": the key '\" + $key + \"' is not unique\") }",
        "        if ($false) { }")


def test_68_printing_the_records_as_one_line_is_rejected() -> None:
    """FIVE RECORDS PRINTED AS ONE LINE is how the collapse went unnoticed."""
    _probe_mutation(
        "test_84",
        "    Write-ProbeLine ('    tab   = ' + [string]$entry.Sheet)",
        "    Write-ProbeLine ('    tab   = ' + [string]$entry.Sheet + [string]$entry.Table)")


def test_69_reporting_a_not_attempted_control_as_false_is_rejected() -> None:
    """RUN 3 PRINTED "permits code VALUE writes: False" after failing before the
    control ran. False reads as "blocking was observed", which nothing tested."""
    _probe_mutation(
        "test_85",
        "    Write-ProbeLine '  UserInterfaceOnly code-value-write capability: NOT TESTED'",
        "    Write-ProbeLine ('  UserInterfaceOnly permits code VALUE writes: ' + [string]$false)")


def test_70_starting_the_control_state_at_a_conclusion_is_rejected() -> None:
    _probe_mutation(
        "test_85",
        "$controlResult = 'NOT ATTEMPTED'",
        "$controlResult = 'REFUSED'")


def test_71_letting_the_control_state_decide_the_verdict_is_rejected() -> None:
    """THE REPORTING CORRECTION MUST NOT ALTER THE VERDICT LOGIC."""
    _probe_mutation(
        "test_86",
        "        if (@($notSucceeded).Count -gt 0) {",
        "        if (($controlResult -eq 'REFUSED') -or (@($notSucceeded).Count -gt 0)) {")


def test_72_skipping_the_resolved_set_recheck_is_rejected() -> None:
    _probe_mutation(
        "test_81",
        "    $null = Assert-ProbeWatchedShape -Watched $resolution.Targets -Manifest $manifest",
        "    $null = $resolution")


if __name__ == "__main__":
    raise SystemExit(pytest.main([__file__, "-q"]))
