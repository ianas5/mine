#!/usr/bin/env python3
"""P10 FINAL WINDOWS ACCEPTANCE - mutation controls for the important requirements.

Each mutation breaks ONE thing the authorisation named, reruns the WHOLE
runner-source battery against the damaged copy, and requires a NAMED control
among the refusers. Nothing here writes to the repository; damaged copies live
in memory.

Runs standalone or under pytest.
"""
from __future__ import annotations

import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import test_phase10_final_acceptance_source as conformance  # noqa: E402


def _tests() -> list[str]:
    names = sorted(name for name in dir(conformance) if name.startswith("test_"))
    assert len(names) >= 30, names
    return names


def _run_battery() -> list[str]:
    refused: list[str] = []
    for name in _tests():
        if name in ("test_73_every_reachable_command_resolves_exactly_once",
                    "test_74_no_variable_can_be_read_before_assignment"):
            continue  # read the file on disk through PowerShell; not part of an in-memory mutation
        try:
            getattr(conformance, name)()
        except BaseException:  # noqa: BLE001 - any refusal counts
            refused.append(name)
    return refused


def _mutate(expected: str, before: str, after: str) -> None:
    original = conformance._runner()
    damaged = original.replace(before, after, 1)
    if damaged == original:
        raise RuntimeError(
            f"the mutation changed nothing: {before[:70]!r} is no longer in the runner")
    saved = dict(conformance._MEMO)
    conformance._MEMO["runner"] = damaged
    conformance._MEMO.pop("runner_code", None)
    try:
        refused = _run_battery()
    finally:
        conformance._MEMO.clear()
        conformance._MEMO.update(saved)
    assert refused, "the mutation survived the whole battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def test_00_the_runner_as_written_passes_every_control() -> None:
    assert _run_battery() == []


def test_01_weakening_the_source_revision_comparison_is_refused() -> None:
    _mutate("test_22",
            "Add-FaCheck 'metadata.source-revision' ($observedRevision -ceq $revision.Expected)",
            "Add-FaCheck 'metadata.source-revision' ($observedRevision -ne '')")


def test_02_a_literal_module_count_is_refused() -> None:
    _mutate("test_30",
            "Add-FaCheck 'modules' ($moduleProblems.Count -eq 0)",
            "Add-FaCheck 'modules' ($observedModules.Count -eq 33)")


def test_03_a_performance_scale_request_is_refused() -> None:
    _mutate("test_06",
            "$acceptanceIterations = [int]$gateBCases.bounds.business_minimum_iterations",
            "$acceptanceIterations = 10000")


def test_04_dropping_the_reset_preservation_assertion_is_refused() -> None:
    _mutate("test_42",
            "    $null = Add-FaCheck 'reset.preserved' ($preservedAfter -ceq $preservedBefore) `\n",
            "    $null = Add-FaCheck 'reset.preserved' $true `\n")


def test_05_dropping_the_protection_assertion_after_a_refusal_is_refused() -> None:
    _mutate("test_62",
            "    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-refusal'\n",
            "")


def test_06_a_check_that_no_longer_fails_fast_is_refused() -> None:
    _mutate("test_64",
            "if ((-not $Ok) -and (-not $Continue)) {",
            "if ($false) {")


def test_07_reaching_for_excels_own_unprotect_is_refused() -> None:
    _mutate("test_12",
            "    $reply = [string]$Excel.Run('P10FW_Begin')\n",
            "    $reply = [string]$Excel.Run('P10FW_Begin')\n    $null = $Workbook.Worksheets.Item(1).Unprotect()\n")


def test_08_a_window_left_open_after_a_repair_precondition_is_refused() -> None:
    _mutate("test_05",
            "    try { Remove-TableRow -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex $missingRowIndex }\n"
            "    finally { $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.missing-row' }\n",
            "    Remove-TableRow -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex $missingRowIndex\n")


def test_09_ui_automation_is_refused() -> None:
    _mutate("test_13",
            "    $excel.Run('PCCM_AutomationEnd') | Out-Null\n",
            "    $excel.Run('PCCM_AutomationEnd') | Out-Null\n    [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')\n")


def test_10_a_drifted_helper_copy_is_refused() -> None:
    _mutate("test_03",
            "if ($null -eq $Value) { $null = $rng.ClearContents() } else { $rng.Value2 = [double]$Value }",
            "$rng.Value2 = $Value")


def test_11_a_state_word_typed_as_a_literal_is_refused() -> None:
    _mutate("test_60",
            "$statusStale         = [string]$p7.model_states.derived_status[2]",
            "$statusStale         = 'STALE'")


def test_12_expecting_the_live_untouched_state_to_read_not_calculated_is_refused() -> None:
    """THE RUN-1 DEFECT, PUT BACK."""
    _mutate("test_60b",
            "(($states0.Calculation -ceq $statusInvalid) -and ($states0.Simulation -ceq $statusInvalid) -and",
            "(($states0.Calculation -ceq $statusNotCalculated) -and ($states0.Simulation -ceq $statusInvalid) -and")


def test_13_reading_the_persisted_history_after_the_compile_check_is_refused() -> None:
    """THE COMPILE CHECK EVALUATES AND PERSISTS THE LIVE STATUS; a persisted read
    after it would read INVALID and the assertion would be wrong for a second reason."""
    original = conformance._runner()
    start = original.index("    # 6a. THE PERSISTED CALCULATION HISTORY")
    end = original.index("    # 3. THE ACCEPTED COMPILE CHECK")
    block = original[start:end]
    moved = original[:start] + original[end:]
    anchor = "    # 2b. SHEETS AND CODENAMES"
    moved = moved.replace(anchor, block + anchor, 1)
    _mutate("test_60b", original, moved)


def test_14_expecting_pass_after_calculate_at_the_business_minimum_is_refused() -> None:
    """THE RUN-2 DEFECT, PUT BACK: no advisory expected, so the derived overall is PASS."""
    _mutate("test_60e",
            "-Scenario 'modelcheck.calculated' `\n        -Expected @($advisoryExpected)\n",
            "-Scenario 'modelcheck.calculated' `\n        -Expected @()\n")


def test_15_dropping_the_calculation_error_from_the_invalid_checkpoint_is_refused() -> None:
    _mutate("test_60f",
            "-Expected @(@{ Id = [string]$calcErrorChecks[0].check_id; Severity = $severityError; Subject = $victimId },\n"
            "                    $advisoryExpected)\n",
            "-Expected @($advisoryExpected)\n")


def test_16_dropping_the_not_calculated_warning_after_reset_is_refused() -> None:
    _mutate("test_60g",
            "-Scenario 'modelcheck.after-reset' `\n        -Expected @($advisoryExpected, $notCalculatedExpected)\n",
            "-Scenario 'modelcheck.after-reset' `\n        -Expected @($advisoryExpected)\n")


def test_17_tolerating_an_unrelated_actionable_row_is_refused() -> None:
    _mutate("test_60h",
            "    if ($unmatched.Count -gt 0) { $problems += ('unexpected actionable row(s): ' + (Format-FaActionable $unmatched)) }\n",
            "")


if __name__ == "__main__":
    failures = 0
    for name in sorted(n for n in dir() if n.startswith("test_")):
        try:
            globals()[name]()
            print(f"PASS {name}")
        except BaseException as exc:  # noqa: BLE001
            failures += 1
            print(f"FAIL {name}: {exc}")
    sys.exit(1 if failures else 0)
