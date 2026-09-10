#!/usr/bin/env python3
"""P10-RP / STRUCTURAL WINDOW MUTATION CONTROLS.

A PROTECTION WINDOW THAT FAILS TO CLOSE HANDS THE USER AN UNPROTECTED WORKBOOK
AND CALLS IT SUCCESS. That is a worse outcome than the defect it fixes, so every
control claiming to refuse it is fed the exact damage.

Each mutation breaks ONE thing, reruns the WHOLE conformance battery against the
damaged copy, and requires a NAMED control among the refusers. A mutation that
survives means the control it belongs to is decoration.

Nothing here writes to the repository. Damaged copies live in memory.

Runs standalone or under pytest.
"""
from __future__ import annotations

import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import test_phase10_structural_window as conformance  # noqa: E402


def _tests() -> list:
    """Every conformance callable, with parametrised ones expanded."""
    out = []
    for name in sorted(n for n in dir(conformance) if n.startswith("test_")):
        fn = getattr(conformance, name)
        marks = getattr(fn, "pytestmark", [])
        params = [m for m in marks if m.name == "parametrize"]
        if not params:
            out.append((name, fn, ()))
            continue
        for case in params[0].args[1]:
            out.append((name, fn, tuple(case) if isinstance(case, tuple) else (case,)))
    assert len(out) >= 30, len(out)
    return out


def _run_battery() -> list[str]:
    refused: list[str] = []
    for name, fn, args in _tests():
        try:
            fn(*args)
        except BaseException:  # noqa: BLE001 - any refusal counts
            if name not in refused:
                refused.append(name)
    return refused


def _mutate(module: str, expected: str, before: str, after: str) -> None:
    """Damage one production module, and prove the anchor still matched.

    NOT AN AssertionError WHEN THE ANCHOR MISSES. A no-op mutation that raised one
    would be indistinguishable from the refusal it is meant to provoke.
    """
    original = conformance._src(module)
    damaged = original.replace(before, after, 1)
    if damaged == original:
        raise RuntimeError(
            f"the mutation changed nothing: {before[:70]!r} is no longer in {module}")
    saved = dict(conformance._MEMO)
    conformance._MEMO[module] = damaged
    conformance._MEMO.pop(module + "::code", None)
    try:
        refused = _run_battery()
    finally:
        conformance._MEMO.clear()
        conformance._MEMO.update(saved)
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def _mutate_evidence(expected: str, before: str, after: str) -> None:
    """Damage the run-evidence record, and prove the anchor still matched."""
    path = PCCM_ROOT / "docs" / "phase10_windows_run_evidence.md"
    original = path.read_text(encoding="utf-8")
    damaged = original.replace(before, after, 1)
    if damaged == original:
        raise RuntimeError(
            f"the mutation changed nothing: {before[:70]!r} is no longer in the evidence")
    try:
        path.write_text(damaged, encoding="utf-8")
        refused = _run_battery()
    finally:
        path.write_text(original, encoding="utf-8")
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


# ===========================================================================
# A. NOT REPROTECTING
# ===========================================================================
def test_01_no_reprotection_on_the_outermost_close_is_rejected() -> None:
    """THE NAMED MUTATION. The window opens, the command succeeds, and the
    workbook is handed back unprotected."""
    _mutate("modProtection.bas", "test_10",
            "    If Not ProtectionApply(detail) Then Exit Function",
            "    ProtectionEndStructural = True\n    Exit Function")


def test_02_reprotecting_without_userinterfaceonly_is_rejected() -> None:
    """THE NAMED MUTATION. A workbook protected against CODE looks protected and
    breaks every command on its next write."""
    _mutate("modProtection.bas", "test_10",
            "        sheet.Protect UserInterfaceOnly:=True, _",
            "        sheet.Protect _")


def test_03_not_closing_the_window_in_the_cleanup_is_rejected() -> None:
    """ON SUCCESS, ON REFUSAL, ON RUNTIME ERROR AND AFTER ROLLBACK AT ONCE:
    FinishOperation is the one path all of them take."""
    _mutate("modAppState.bas", "test_30",
            "    If Snapshot.Structural Then",
            "    If False Then")


def test_04_swallowing_a_restoration_failure_is_rejected() -> None:
    """THE NAMED MUTATION. The business mutation may have committed; that does
    not make an unprotected workbook an ordinary success."""
    _mutate("modAppState.bas", "test_33",
            '            problems = problems & "  PROTECTION WAS NOT RESTORED: " & _\n'
            "                       protectionProblem & vbCrLf",
            "            protectionProblem = vbNullString")


def test_05_reporting_success_without_verifying_it_is_rejected() -> None:
    _mutate("modProtection.bas", "test_11",
            "    If Not ProtectionIsApplied() Then",
            "    If False Then")


def test_06_leaving_a_half_open_window_unprotected_is_rejected() -> None:
    _mutate("modProtection.bas", "test_14",
            "    If Not ProtectionApply(reapply) Then",
            "    If False Then")


# ===========================================================================
# B. NESTING
# ===========================================================================
def test_10_a_nested_close_that_restores_too_early_is_rejected() -> None:
    """THE NAMED MUTATION. The outer operation is still running and still has a
    rollback ahead of it."""
    _mutate("modProtection.bas", "test_03",
            "    If mStructuralDepth > 0 Then\n"
            "        ' AN INNER CLOSE RESTORES NOTHING. The outer operation is still running\n"
            "        ' and may still have structural work and a rollback ahead of it.\n"
            "        ProtectionEndStructural = True\n"
            "        Exit Function\n"
            "    End If",
            "    If mStructuralDepth > 0 Then\n"
            "        ProtectionEndStructural = ProtectionApply(detail)\n"
            "        Exit Function\n"
            "    End If")


def test_11_a_nested_open_that_releases_again_is_rejected() -> None:
    _mutate("modProtection.bas", "test_03",
            "    If mStructuralDepth > 0 Then\n"
            "        mStructuralDepth = mStructuralDepth + 1",
            "    If mStructuralDepth > 0 Then\n"
            "        ThisWorkbook.Worksheets(1).Unprotect\n"
            "        mStructuralDepth = mStructuralDepth + 1")


def test_12_dropping_the_depth_counter_is_rejected() -> None:
    _mutate("modProtection.bas", "test_03",
            "Private mStructuralDepth As Long",
            "Private mStructuralHeld As Boolean")


def test_13_closing_more_often_than_opening_going_unguarded_is_rejected() -> None:
    _mutate("modProtection.bas", "test_04",
            "    If mStructuralDepth <= 0 Then",
            "    If False Then")


def test_14_a_second_close_decrementing_twice_is_rejected() -> None:
    """THE FLAG IS CLEARED BEFORE THE CLOSE IS ATTEMPTED for exactly this."""
    _mutate("modAppState.bas", "test_30",
            "    If Snapshot.Structural Then\n        Snapshot.Structural = False",
            "    If Snapshot.Structural Then")


# ===========================================================================
# C. WHO ASKS
# ===========================================================================
def test_20_apply_timeline_forgetting_to_request_the_window_is_rejected() -> None:
    """THE NAMED MUTATION, and the one that produced the Error 1004."""
    _mutate("modTimeline.bas", "test_20",
            "    modAppState.BeginStructuralOperation snapshot",
            "    modAppState.BeginOperation")


def test_21_calculate_forgetting_to_request_the_window_is_rejected() -> None:
    _mutate("modCalcReport.bas", "test_20",
            "    modAppState.BeginStructuralOperation state",
            "    modAppState.BeginOperation")


def test_22_repair_forgetting_to_request_the_window_is_rejected() -> None:
    _mutate("modRepair.bas", "test_20",
            "    modAppState.BeginStructuralOperation state",
            "    modAppState.BeginOperation")


def test_23_the_driver_commands_forgetting_to_request_it_is_rejected() -> None:
    _mutate("modDrivers.bas", "test_20",
            "    modAppState.BeginStructuralOperation snapshot",
            "    modAppState.BeginOperation")


def test_24_a_simulation_opening_the_window_is_rejected() -> None:
    """THE NAMED MUTATION. A stochastic run holds the workbook for minutes."""
    _mutate("modSimReport.bas", "test_21",
            "    modAppState.BeginOperation",
            "    modAppState.BeginStructuralOperation state")


def test_25_the_annual_run_opening_the_window_is_rejected() -> None:
    _mutate("modSimAnnualRun.bas", "test_21",
            "    modAppState.BeginOperation",
            "    modAppState.BeginStructuralOperation state")


def test_26_sensitivity_opening_the_window_is_rejected() -> None:
    _mutate("modSimPostReport.bas", "test_21",
            "    modAppState.BeginOperation",
            "    modAppState.BeginStructuralOperation state")


def test_27_reset_results_opening_the_window_is_rejected() -> None:
    _mutate("modReset.bas", "test_21",
            "    modAppState.BeginOperation",
            "    modAppState.BeginStructuralOperation state")


# ===========================================================================
# D. SCATTERING PROTECTION
# ===========================================================================
def test_30_unprotect_added_directly_to_modprofiling_is_rejected() -> None:
    """THE NAMED MUTATION, and the workaround this whole design exists to
    prevent: a helper that releases protection around its own writes and leaves
    the workbook open if it fails midway."""
    _mutate("modProfiling.bas", "test_01",
            "        Set added = target.ListColumns.Add",
            "        target.Parent.Unprotect\n"
            "        Set added = target.ListColumns.Add")


def test_31_unprotect_added_to_the_envelope_is_rejected() -> None:
    _mutate("modAppState.bas", "test_01",
            "Public Sub BeginStructuralOperation(ByRef Snapshot As AppStateSnapshot)\n"
            "    BeginOperation",
            "Public Sub BeginStructuralOperation(ByRef Snapshot As AppStateSnapshot)\n"
            "    BeginOperation\n"
            "    ThisWorkbook.Worksheets(1).Unprotect")


def test_32_unprotect_added_to_a_driver_command_is_rejected() -> None:
    _mutate("modDrivers.bas", "test_01",
            "        register.ListRows.Add",
            "        register.Parent.Unprotect\n        register.ListRows.Add")


def test_33_releasing_workbook_structure_in_the_window_is_rejected() -> None:
    """WIDER THAN ANY EVIDENCE ASKS FOR. The 1004 named the sheet."""
    _mutate("modProtection.bas", "test_12",
            "    On Error GoTo Failed\n    For Each sheet In ThisWorkbook.Worksheets\n"
            "        If sheet.ProtectContents Then sheet.Unprotect\n    Next sheet",
            "    On Error GoTo Failed\n    ThisWorkbook.Unprotect\n"
            "    For Each sheet In ThisWorkbook.Worksheets\n"
            "        If sheet.ProtectContents Then sheet.Unprotect\n    Next sheet")


def test_34_skipping_the_proof_that_the_release_took_is_rejected() -> None:
    _mutate("modProtection.bas", "test_13",
            "    Set sheet = Nothing\n    For Each sheet In ThisWorkbook.Worksheets\n"
            "        If sheet.ProtectContents Then\n"
            '            detail = "worksheet protection was not released from " & sheet.Name\n'
            "            GoTo Failed\n        End If\n    Next sheet",
            "    Set sheet = Nothing")


def test_35_attempting_structural_work_after_a_failed_open_is_rejected() -> None:
    _mutate("modAppState.bas", "test_15",
            "    If Not modProtection.ProtectionBeginStructural(detail) Then\n"
            "        Err.Raise vbObjectError + 5010,",
            "    If False Then\n"
            "        Err.Raise vbObjectError + 5010,")


# ===========================================================================
# E. WHAT MUST NOT HAVE MOVED
# ===========================================================================
def test_40_changing_the_command_feedback_is_rejected() -> None:
    """REQUIRED CONTROL 16. The reconciliation is a protection change and is not
    licence to touch how a command reports."""
    _mutate("modAppState.bas", "test_41",
            "        MsgBox body, vbExclamation, MSG_TITLE",
            "        MsgBox body, vbCritical, MSG_TITLE")


def test_41_losing_the_rollback_call_from_the_failure_path_is_rejected() -> None:
    """CAUGHT BY THE ORDERING CONTROL, WHICH IS THE HONEST OWNER OF THIS. The
    rollback FUNCTION is still defined - only the call is gone - so a marker
    sweep over the module cannot see it, and test_31 is what pins the call."""
    _mutate("modTimeline.bas", "test_31",
            "        restoreNote = TryRestoreTimeline(appliedBase, appliedStart, appliedDuration, _",
            "        restoreNote = vbNullString\n"
            "        Dim unusedRestore As String: unusedRestore = TryRestoreTimelineDisabled( _")


def test_42_closing_the_window_before_the_rollback_is_rejected() -> None:
    """THE ONE ORDERING THAT MUST NOT BE GOT WRONG."""
    _mutate("modCalcReport.bas", "test_31",
            "    result = RunCalculation(committed)",
            "    cleanup = modAppState.FinishOperation(state)\n"
            "    result = RunCalculation(committed)")


def test_43_a_command_closing_the_window_itself_is_rejected() -> None:
    _mutate("modRepair.bas", "test_32",
            "    cleanup = modAppState.FinishOperation(state)",
            "    Dim pd As String\n"
            "    If Not modProtection.ProtectionEndStructural(pd) Then pd = pd\n"
            "    cleanup = modAppState.FinishOperation(state)")


def test_44_introducing_a_password_is_rejected() -> None:
    _mutate("modProtection.bas", "test_44",
            "        sheet.Protect UserInterfaceOnly:=True, _",
            '        sheet.Protect Password:="pccm", UserInterfaceOnly:=True, _')


def test_45_a_public_unprotect_command_is_rejected() -> None:
    """NO USER-FACING UNPROTECTED MODE."""
    _mutate("modProtection.bas", "test_45",
            "Public Function ProtectionRelease(ByRef detail As String) As Boolean",
            "Public Sub PCCM_UnprotectWorkbook()\n"
            "    Dim d As String\n"
            "    If Not ProtectionRelease(d) Then d = d\n"
            "End Sub\n\n"
            "Public Function ProtectionRelease(ByRef detail As String) As Boolean")


def test_47_dropping_the_run_6_structural_effect_is_rejected() -> None:
    """AN ACCEPTANCE CLAIM WITHOUT THE OBSERVED SHAPES."""
    _mutate_evidence("test_52", "25×2 → 25×5", "a shape change")


def test_48_dropping_the_run_6_ordering_caveat_is_rejected() -> None:
    """THE PROBE RAN BEFORE THE FRESH STAGE-A BUILD, and saying so is the
    difference between evidence and an acceptance claim."""
    _mutate_evidence("test_52", "not** ideal acceptance evidence", "good evidence")


def test_49_rewriting_the_calculate_refusal_as_protection_is_rejected() -> None:
    _mutate_evidence("test_52",
                     "Discount Rate: the value is blank. A blank is not zero.",
                     "the sheet is protected.")


def test_50_dropping_the_structure_protected_evidence_is_rejected() -> None:
    _mutate_evidence("test_53", "privilege envelope is not widened",
                     "envelope may be revisited")


def test_51_claiming_delete_coverage_before_windows_proves_it_is_rejected() -> None:
    """THE OVERREACH THIS WHOLE ROUND EXISTS TO PREVENT. Nothing here is runtime
    evidence for the delete path; it is a probe that can now ask the question."""
    _mutate_evidence("test_52b",
                     "**The delete path is unproved on Windows.**",
                     "The delete path is proved and Benchmark Run 3 was a HARNESS defect only.")


def test_52_dropping_the_run_7_coverage_gap_is_rejected() -> None:
    """WITHOUT IT, RUN 7 READS AS A CLOSURE IT IS NOT."""
    _mutate_evidence("test_52a",
                     "remains PENDING exact delete-path runtime evidence",
                     "is settled")


def test_53_dropping_what_run_7_did_not_prove_is_rejected() -> None:
    _mutate_evidence("test_52a", "**Not proved:**", "**Also observed:**")


def test_54_rewriting_run_7_as_an_error_rather_than_a_gap_is_rejected() -> None:
    """RUN 7 WAS HONEST AGAINST THE CRITERIA OF ITS DAY. The gap is in the
    criteria, and saying otherwise rewrites a run's history."""
    _mutate_evidence("test_52a",
                     "coverage gap in the criteria",
                     "mistake by the probe")


def test_55_dropping_the_run_7_observed_calc_shapes_is_rejected() -> None:
    _mutate_evidence("test_52a", "1×8 → 3×8", "a larger table")


def test_46_deleting_the_recorded_evidence_is_rejected() -> None:
    """REQUIRED CONTROL 24. A reconciliation whose reason is not recorded is one
    the next rewrite undoes."""
    _mutate("modProtection.bas", "test_51",
            "' THE ORIGINAL ASSUMPTION, AND IT IS NOT BEING REWRITTEN. The accepted Phase-10",
            "' A NOTE ON PROTECTION. The accepted Phase-10")


if __name__ == "__main__":
    import pytest
    raise SystemExit(pytest.main([__file__, "-q"]))
