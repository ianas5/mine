#!/usr/bin/env python3
"""P10 / PROBE RUN 4 MUTATION CONTROLS.

The Run-4 defect was a single missing cast, in a helper that had been copied
from an accepted one and quietly simplified. Every control that claims to refuse
that is fed the exact damage, including the two the authorisation names: the
faulty setter restored, and numeric values coerced to [string].

Each mutation breaks ONE thing, reruns the WHOLE battery against the damaged
copy, and requires a NAMED control among the refusers.

Nothing here writes to the repository. Damaged copies live in memory.

Runs standalone or under pytest.
"""
from __future__ import annotations

import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import test_phase10_probe_input_types as conformance  # noqa: E402


def _tests() -> list[str]:
    names = sorted(name for name in dir(conformance) if name.startswith("test_"))
    assert len(names) >= 20, names
    return names


def _run_battery() -> list[str]:
    refused: list[str] = []
    for name in _tests():
        try:
            getattr(conformance, name)()
        except BaseException:  # noqa: BLE001 - any refusal counts
            refused.append(name)
    return refused


def _mutate(expected: str, before: str, after: str) -> None:
    """Damage the probe, and prove the anchor still matched.

    NOT AN AssertionError WHEN THE ANCHOR MISSES. A no-op mutation that raised one
    would be indistinguishable from the refusal it is meant to provoke.
    """
    original = conformance._probe()
    damaged = original.replace(before, after, 1)
    if damaged == original:
        raise RuntimeError(
            f"the mutation changed nothing: {before[:70]!r} is no longer in the probe")
    saved = dict(conformance._MEMO)
    conformance._MEMO["probe"] = damaged
    conformance._MEMO.pop("probe_code", None)
    try:
        refused = _run_battery()
    finally:
        conformance._MEMO.clear()
        conformance._MEMO.update(saved)
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


# ===========================================================================
# A. THE RUN-4 DEFECT, PUT BACK
# ===========================================================================
def test_01_restoring_the_faulty_untyped_setter_is_rejected() -> None:
    """THE NAMED MUTATION. The exact statement that raised the
    InvalidCastException on Windows."""
    _mutate("test_0",
            "        if ($null -eq $Value) { $null = $rng.ClearContents() } else { $rng.Value2 = [double]$Value }",
            "        $rng.Value2 = $Value")


def test_02_dropping_only_the_double_cast_is_rejected() -> None:
    """THE SMALLER VERSION OF THE SAME MISTAKE, which is the one that would
    actually be made: the branch kept, the cast lost."""
    _mutate("test_0",
            "else { $rng.Value2 = [double]$Value }",
            "else { $rng.Value2 = $Value }")


def test_03_coercing_a_numeric_input_to_string_is_rejected() -> None:
    """THE NAMED MUTATION. It satisfies the COM binder and then makes
    modTimeline's TryReadDouble refuse the timeline - a REFUSED endpoint that
    reads exactly like protection blocking the work."""
    _mutate("test_0",
            "else { $rng.Value2 = [double]$Value }",
            "else { $rng.Value2 = [string]$Value }")


def test_04_stringifying_at_the_call_site_is_rejected() -> None:
    _mutate("test_0",
            "            -Value ([double]$entry.Value)",
            "            -Value ([string]$entry.Value)")


def test_05_dropping_the_clearcontents_branch_is_rejected() -> None:
    """REQUIRED CONTROL 5, and the defect that made Run 2 contradict itself."""
    _mutate("test_12",
            "        if ($null -eq $Value) { $null = $rng.ClearContents() } else { $rng.Value2 = [double]$Value }",
            "        $rng.Value2 = [double]$Value")


def test_06_deleting_the_recorded_reason_is_rejected() -> None:
    """A FIX WHOSE REASON IS NOT WRITTEN BESIDE IT IS ONE THE NEXT REWRITE
    DELETES - this defect has now been diagnosed twice in this repository."""
    _mutate("test_03",
            "phase5_gate_b_scenarios.ps1:922",
            "an earlier run")


# ===========================================================================
# B. THE COPY
# ===========================================================================
def test_10_drifting_from_the_accepted_definition_is_rejected() -> None:
    """REQUIRED CONTROL 4. A copy nothing proves is a copy nobody maintains."""
    _mutate("test_10",
            "        $rng = $nm.RefersToRange\n"
            "        if ($null -eq $Value) { $null = $rng.ClearContents() } else { $rng.Value2 = [double]$Value }",
            "        $rng = $nm.RefersToRange\n"
            "        if ($null -eq $Value) { $null = $rng.ClearContents() }\n"
            "        else { $rng.Value2 = [double]$Value }")


def test_11_collapsing_two_types_onto_one_assignment_site_is_rejected() -> None:
    """THE MECHANISM ITSELF. One line cannot carry two CLR types."""
    _mutate("test_14",
            "    elseif ($Value -is [string]) { $Cell.Value2 = [string]$Value }\n"
            "    elseif ($Value -is [double]) { $Cell.Value2 = [double]$Value }",
            "    elseif (($Value -is [string]) -or ($Value -is [double])) { $Cell.Value2 = $Value }")


def test_12_coercing_an_unsupported_type_instead_of_refusing_is_rejected() -> None:
    _mutate("test_13",
            "        throw ('the captured cell value is a ' + $Value.GetType().FullName +\n"
            "               ', which this probe will not write back by coercion')",
            "        $Cell.Value2 = [string]$Value")


# ===========================================================================
# C. THE INPUTS
# ===========================================================================
def test_20_changing_the_declared_timeline_values_is_rejected() -> None:
    """NOT A SECOND FIXTURE CONTRACT. The triple is the accepted one."""
    _mutate("test_2",
            "@{ Key = 'project_start_year'; Value = [double]2027; Endpoint = 'PCCM_ApplyTimeline' }",
            "@{ Key = 'project_start_year'; Value = [double]2030; Endpoint = 'PCCM_ApplyTimeline' }")


def test_21_hard_coding_a_defined_name_is_rejected() -> None:
    _mutate("test_22",
            "            DefinedName = (Get-ProbeScalarString -InputObject $spec -Name 'defined_name' `\n"
            "                               -Where ('the ' + $key + ' input'))",
            "            DefinedName = 'inpBaseYear'")


def test_22_setting_more_than_the_command_needs_is_rejected() -> None:
    """THE PROBE IS NOT A FIXTURE BUILDER."""
    _mutate("test_20",
            "        @{ Key = 'discount_rate';      Value = [double]0.05; Endpoint = 'PCCM_Calculate' }",
            "        @{ Key = 'discount_rate';      Value = [double]0.05; Endpoint = 'PCCM_Calculate' },\n"
            "        @{ Key = 'reporting_currency'; Value = [double]1;    Endpoint = 'PCCM_Calculate' }")


# ===========================================================================
# C2. THE CALCULATE STRUCTURAL EVIDENCE
# ===========================================================================
def test_23_going_back_to_judging_calculate_by_its_announcement_is_rejected() -> None:
    """THE RUN-6 WEAKNESS, put back."""
    _mutate("test_25",
            "            $calcBefore = @(Get-ProbeCalcShapes -Worksheet $calcSheet -CalcTables $calcTables)",
            "            $calcBefore = @()")


def test_23a_reading_the_shapes_only_after_the_call_is_rejected() -> None:
    _mutate("test_25",
            "            $calcAfter = @(Get-ProbeCalcShapes -Worksheet $calcSheet -CalcTables $calcTables)",
            "            $calcAfter = @($calcBefore)")


def test_23b_hard_coding_a_calc_table_name_is_rejected() -> None:
    _mutate("test_26",
            "            Table    = (Get-ProbeScalarString -InputObject $spec -Name 'table_name' `\n"
            "                            -Where ('the ' + [string]$property.Name + ' calc table'))",
            "            Table    = 'tblCalcYears'")


def test_23c_accepting_any_difference_as_proof_is_rejected() -> None:
    """"SOMETHING CHANGED" IS NOT EVIDENCE OF THE CONTRACTED WORK."""
    _mutate("test_27",
            "                if ([int]@($row)[0].Rows -ne $duration) {",
            "                if ($false) {")


def test_23d_dropping_the_row_rule_selection_is_rejected() -> None:
    _mutate("test_27",
            "        if ([string]$entry.RowRule -eq 'one row per applied project year') {",
            "        if ($true) {")


def test_23e_failing_a_refused_calculate_for_not_resizing_is_rejected() -> None:
    """A REFUSAL IS ENTITLED TO LEAVE THE TABLES ALONE."""
    _mutate("test_28",
            "        if ([string]$calculate.Outcome -eq 'SUCCEEDED') {",
            "        if ($true) {")


def test_23f_letting_a_contradicted_announcement_reach_fine_is_rejected() -> None:
    """NO FALSE PASS FROM THE ANNOUNCEMENT ALONE."""
    _mutate("test_29",
            "        } elseif ($calcStructuralProof -eq 'CONTRADICTED') {",
            "        } elseif ($false) {")


def test_23g_setting_a_verdict_from_the_contradicted_branch_is_rejected() -> None:
    _mutate("test_29",
            "            $verdictReason = ('Calculate announced success but the _Calc tables did not take ' +",
            "            $verdict = 'PRODUCTION IS FINE UNDER PROTECTION'\n"
            "            $verdictReason = ('Calculate announced success but the _Calc tables did not take ' +")


def test_23h_weakening_the_blocked_signature_is_rejected() -> None:
    """NO WEAKENING OF BLOCKED SEMANTICS."""
    _mutate("test_2a",
            "    if (-not [bool]$Outcome.Invoked) { return $false }",
            "    if ($false) { return $false }")


def test_23i_overclaiming_the_add_commands_is_rejected() -> None:
    _mutate("test_2b",
            "settles endpoint functionality under protection, NOT capacity expansion",
            "settles ListRows.Add capacity expansion under protection")


def test_23j_releasing_workbook_structure_from_the_probe_is_rejected() -> None:
    """RUN 6 PROVED IT IS NOT NEEDED: ApplyTimeline SUCCEEDED with structure
    protection True throughout."""
    _mutate("test_2c",
            "            $calcSheet = $resolution.Sheets.Item([string]@($calcTables)[0].Sheet)",
            "            $wb.Unprotect()\n"
            "            $calcSheet = $resolution.Sheets.Item([string]@($calcTables)[0].Sheet)")


# ===========================================================================
# D. READBACK AND INVOCATION DISCIPLINE
# ===========================================================================
def test_22a_dropping_the_discount_rate_prerequisite_is_rejected() -> None:
    """THE EXACT STATE RUN 6 REACHED: Calculate invoked and refused with
    "Discount Rate: the value is blank. A blank is not zero." """
    _mutate("test_20",
            "        @{ Key = 'discount_rate';      Value = [double]0.05; Endpoint = 'PCCM_Calculate' }\n",
            "")


def test_22b_stringifying_the_discount_rate_is_rejected() -> None:
    """modCalcResolve.IsRealNumber tests the VarType, so text would refuse."""
    _mutate("test_20",
            "@{ Key = 'discount_rate';      Value = [double]0.05; Endpoint = 'PCCM_Calculate' }",
            "@{ Key = 'discount_rate';      Value = '0.05'; Endpoint = 'PCCM_Calculate' }")


def test_22c_running_calculate_after_the_add_commands_is_rejected() -> None:
    """AN ADDED DRIVER CARRIES A PERMANENT ID AND NO OTHER FIELD, so Calculate
    would refuse on it and the probe would learn nothing about the _Calc resize."""
    _mutate("test_20b",
            "        $calculate = Invoke-ProbeEndpoint -Excel $excel -Workbook $wb `\n"
            "                -Endpoint 'PCCM_Calculate' -Resolution $resolution\n"
            "            $null = $outcomes.Add($calculate)\n",
            "")


def test_30_removing_the_readback_is_rejected() -> None:
    """REQUIRED CONTROL 6."""
    _mutate("test_3",
            "        $actual = Get-ProbeNamedValue -Workbook $Workbook -DefinedName $name",
            "        $actual = $expected")


def test_31_checking_the_value_without_the_type_is_rejected() -> None:
    """REQUIRED CONTROL 2. A stringified 2026 passes a value-only comparison."""
    _mutate("test_31",
            "        if ($actual -isnot [double]) {",
            "        if ($false) {")


def test_32_checking_the_value_before_the_type_is_rejected() -> None:
    _mutate("test_31",
            "        if ($actual -isnot [double]) {\n"
            "            $problems += ($name + ': reads back as ' + $actual.GetType().FullName +\n"
            "                          ', not System.Double - the numeric input was stringified')\n"
            "            continue\n"
            "        }\n"
            "        if ([double]$actual -ne $expected) {\n"
            "            $problems += ($name + ': reads back ' + [string]$actual +\n"
            "                          ', expected ' + [string]$expected)\n"
            "        }",
            "        if ([double]$actual -ne $expected) {\n"
            "            $problems += ($name + ': reads back ' + [string]$actual +\n"
            "                          ', expected ' + [string]$expected)\n"
            "        }\n"
            "        if ($actual -isnot [double]) {\n"
            "            $problems += ($name + ': reads back as ' + $actual.GetType().FullName +\n"
            "                          ', not System.Double - the numeric input was stringified')\n"
            "        }")


def test_33_stringifying_the_readback_reader_is_rejected() -> None:
    _mutate("test_32",
            "        return $range.Value2",
            "        return [string]$range.Value2")


def test_34_letting_a_failed_precondition_reach_the_endpoint_is_rejected() -> None:
    """REQUIRED CONTROL 7. This is the one that keeps a fixture failure from
    being recorded as a production endpoint result."""
    _mutate("test_33",
            "        if (@($inputProblems).Count -gt 0) {",
            "        if ($false) {")


def test_35_restoring_the_run_4_stage_wording_is_rejected() -> None:
    """SECTION 9. 'setting the timeline inputs' under stage 'endpoint' is what a
    reader could mistake for production having run."""
    _mutate("test_35",
            "            -Action 'SETTING ENDPOINT PRECONDITIONS: writing the declared inputs (production NOT invoked)' `",
            "            -Action 'setting the timeline inputs' `")


def test_36_marking_the_endpoint_invoked_early_is_rejected() -> None:
    _mutate("test_34",
            "        $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null",
            "        $invoked = $true\n"
            "        $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null")


# ===========================================================================
# E. THE RUN-4 EVIDENCE
# ===========================================================================
def test_40_weakening_the_locked_cell_control_is_rejected() -> None:
    """REQUIRED CONTROL 11. Run 4's one solid fact rests on this control."""
    _mutate("test_40",
            "        if (-not $cell.Locked) {",
            "        if ($false) {")


def test_41_stringifying_the_restoration_check_is_rejected() -> None:
    """A [string] COMPARISON ACCEPTS A DOUBLE RESTORED AS TEXT."""
    _mutate("test_40",
            "        if (-not (Test-ProbeExactValue -Actual $restored -Expected $original)) {",
            "        if ([string]$restored -ne [string]$original) {")


def test_42_projecting_the_control_onto_a_boolean_is_rejected() -> None:
    """REQUIRED CONTROL 12. NOT ATTEMPTED has no False, and the capability
    answer is not the verdict."""
    _mutate("test_41",
            "        $controlResult = [string]$control.Result",
            "        $controlWorked = ([string]$control.Result -eq 'SUCCEEDED')\n"
            "        $controlResult = [string]$control.Result")


def test_43_dropping_strict_mode_is_rejected() -> None:
    _mutate("test_43", "Set-StrictMode -Version 2.0", "Set-StrictMode -Off")


def test_44_dropping_the_transient_com_check_is_rejected() -> None:
    _mutate("test_44", "Get-TransientFailures", "Get-TransientFailuresRenamed")


def test_45_releasing_protection_from_the_probe_is_rejected() -> None:
    _mutate("test_45",
            "        $header = $target.ListObject.HeaderRowRange",
            "        $target.Worksheet.Unprotect()\n"
            "        $header = $target.ListObject.HeaderRowRange")


if __name__ == "__main__":
    import pytest
    raise SystemExit(pytest.main([__file__, "-q"]))
