#!/usr/bin/env python3
"""PCCM Phase 6 Step-11 MUTATION CONTROLS for the modSimReport source battery.

A conformance test that cannot fail proves nothing. Every control damages one of
the three authorities this step touches - `modSimReport.bas`,
`modCalcReport.bas` or `spec/structure_contract.yaml` - reruns the WHOLE Step-11
battery against the damaged copy, and requires a NAMED detector among the
refusers.

Nothing here writes to the repository: damaged copies live in a temporary
directory and the conformance module is pointed at them for one control.

Runs standalone or under pytest.
"""

from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import test_phase6_sim_report_vba as conformance  # noqa: E402

_REPORT = conformance.REPORT_BAS.read_text(encoding="utf-8")
_CALC = conformance.CALC_REPORT_BAS.read_text(encoding="utf-8")
_STRUCTURE = (conformance.SPEC / "structure_contract.yaml").read_text(encoding="utf-8")


def _conformance_tests() -> list[str]:
    names = sorted(n for n in dir(conformance) if n.startswith("test_"))
    assert len(names) >= 45, names
    return names


def _run_battery() -> list[str]:
    refused = []
    for name in _conformance_tests():
        try:
            getattr(conformance, name)()
        except BaseException:  # noqa: BLE001 - any refusal counts
            refused.append(name)
    return refused


def _install(report: str | None = None, calc: str | None = None,
             structure: str | None = None):
    saved = (conformance.REPORT_BAS, conformance.CALC_REPORT_BAS, conformance.SPEC,
             dict(conformance._CACHE))
    conformance._CACHE.clear()
    temp = Path(tempfile.mkdtemp(prefix="pccm-step11-mutation-"))
    if report is not None:
        assert report != _REPORT, "the mutation changed nothing"
        target = temp / "modSimReport.bas"
        target.write_text(report, encoding="utf-8")
        conformance.REPORT_BAS = target
    if calc is not None:
        assert calc != _CALC, "the mutation changed nothing"
        target = temp / "modCalcReport.bas"
        target.write_text(calc, encoding="utf-8")
        conformance.CALC_REPORT_BAS = target
    if structure is not None:
        assert structure != _STRUCTURE, "the mutation changed nothing"
        spec_dir = temp / "spec"
        shutil.copytree(saved[2], spec_dir)
        (spec_dir / "structure_contract.yaml").write_text(structure, encoding="utf-8")
        conformance.SPEC = spec_dir

    def restore() -> None:
        conformance.REPORT_BAS = saved[0]
        conformance.CALC_REPORT_BAS = saved[1]
        conformance.SPEC = saved[2]
        conformance._CACHE.clear()
        conformance._CACHE.update(saved[3])

    return restore


def _control(expected: str, report: str | None = None, calc: str | None = None,
             structure: str | None = None) -> None:
    restore = _install(report, calc, structure)
    try:
        refused = _run_battery()
    finally:
        restore()
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def _swap(text: str, old: str, new: str, count: int = 1) -> str:
    assert text.count(old) == count, (old[:80], text.count(old))
    return text.replace(old, new)


def _after(text: str, anchor: str, inserted: str) -> str:
    return _swap(text, anchor, anchor + inserted)


def test_00_the_accepted_sources_pass_every_detector() -> None:
    restore = _install()
    try:
        refused = _run_battery()
    finally:
        restore()
    assert refused == [], refused


# ===========================================================================
# A. The transaction
# ===========================================================================
def test_01_the_candidate_writes_to_the_active_bank() -> None:
    damaged = _REPORT.replace("package.TargetBank = InactiveBank(package.ActiveBank)",
                              "package.TargetBank = package.ActiveBank")
    damaged = damaged.replace("SnapshotRange(package.TargetBank)",
                              "SnapshotRange(package.ActiveBank)")
    _control("test_25", report=damaged)


def test_02_the_first_success_targets_bank_b() -> None:
    damaged = _swap(
        _REPORT,
        "    If Len(active) = 0 Then\n        InactiveBank = SIM_BANK_A\n",
        "    If Len(active) = 0 Then\n        InactiveBank = SIM_BANK_B\n")
    _control("test_26", report=damaged)


def test_03_a_third_bank_is_introduced() -> None:
    damaged = _swap(
        _REPORT,
        "    If StrComp(active, SIM_BANK_B, vbBinaryCompare) = 0 Then\n"
        "        InactiveBank = SIM_BANK_A\n"
        "    End If\n",
        "    If StrComp(active, SIM_BANK_B, vbBinaryCompare) = 0 Then\n"
        '        InactiveBank = "C"\n'
        "    End If\n")
    _control("test_26", report=damaged)


def test_04_the_bank_is_switched_before_the_candidate_is_verified() -> None:
    damaged = _swap(
        _REPORT,
        "    If Not VerifyCandidateBank(package, snapshot, summary, contingency, detail) Then\n"
        "        Exit Function\n"
        "    End If\n", "")
    damaged = _swap(
        damaged,
        "    ' 19. The one final write. The active bank moves last, inside it.\n",
        "    ' 19.\n    If Not VerifyCandidateBank2(package, detail) Then Exit Function\n")
    _control("test_30", report=damaged)


def test_05_the_active_bank_is_written_separately_before_the_commit() -> None:
    damaged = _after(
        _REPORT,
        "    previous = SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2\n",
        "    SharedCell(SIM_IDENTITY_ROW_ACTIVE_BANK).Value2 = package.TargetBank\n")
    _control("test_31", report=damaged)


def test_06_the_final_commit_becomes_nine_writes() -> None:
    damaged = _swap(
        _REPORT,
        "    SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2 = block\n",
        "    SharedCell(SIM_IDENTITY_ROW_LAST_RUN_ID).Value2 = package.CandidateRunId\n"
        "    SharedCell(SIM_IDENTITY_ROW_ACTIVE_BANK).Value2 = package.TargetBank\n")
    _control("test_31", report=damaged)


def test_07_the_run_id_is_allocated_before_the_commit() -> None:
    damaged = _after(
        _REPORT,
        "    package.CandidateRunId = lastRunId + 1\n",
        "    SharedCell(SIM_IDENTITY_ROW_LAST_RUN_ID).Value2 = package.CandidateRunId\n")
    _control("test_17", report=damaged)


def test_08_run_id_exhaustion_consumes_the_nonce() -> None:
    """The headroom check moves AFTER the allocation, so an unrunnable run burns
    a sequence."""
    damaged = _swap(
        _REPORT,
        "    If lastRunId >= SIM_RUN_ID_MAXIMUM Then\n"
        '        detail = "simulation: the run identity counter is exhausted; no further run " & _\n'
        '                 "can be identified or committed"\n'
        "        Exit Function\n"
        "    End If\n"
        "    package.CandidateRunId = lastRunId + 1\n",
        "    package.CandidateRunId = lastRunId + 1\n")
    damaged = _after(
        damaged, "    package.NonceConsumed = True\n",
        "    If package.CandidateRunId > SIM_RUN_ID_MAXIMUM Then\n"
        '        detail = "simulation: the run identity counter is exhausted"\n'
        "        Exit Function\n"
        "    End If\n")
    _control("test_18", report=damaged)


def test_09_sampling_begins_before_the_nonce_is_persisted() -> None:
    damaged = _swap(
        _REPORT,
        "    SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 = package.ConsumedNonce + 1\n",
        "")
    _control("test_19", report=damaged)


def test_10_the_nonce_is_rolled_back_after_a_later_failure() -> None:
    damaged = _swap(
        _REPORT,
        "Private Function RecordFailure(ByRef package As SimRunPackage, _\n"
        "                               ByVal detail As String) As OperationResult\n",
        "Private Function RecordFailure(ByRef package As SimRunPackage, _\n"
        "                               ByVal detail As String) As OperationResult\n"
        "    If package.NonceConsumed Then\n"
        "        SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 = package.ConsumedNonce\n"
        "    End If\n")
    _control("test_20", report=damaged)


def test_11_the_failed_attempt_audit_loses_the_consumed_nonce() -> None:
    damaged = _swap(
        _REPORT,
        "    If package.NonceAllocated Then\n"
        "        block(5, 1) = package.ConsumedNonce\n"
        "    Else\n"
        "        block(5, 1) = vbNullString\n"
        "    End If\n",
        "    block(5, 1) = vbNullString\n")
    _control("test_20", report=damaged)


def test_12_a_mismatched_candidate_bank_still_commits() -> None:
    damaged = _swap(
        _REPORT,
        "    If Not VerifyCandidateBank(package, snapshot, summary, contingency, detail) Then\n"
        "        Exit Function\n"
        "    End If\n",
        "")
    _control("test_30", report=damaged)


def test_13_the_final_commit_verification_is_omitted() -> None:
    damaged = _swap(
        _REPORT,
        "    If SameBlock(SIM_FINAL_COMMIT_RANGE, block, 9, 1) Then\n"
        "        On Error GoTo 0\n"
        "        FinalCommit = True\n"
        "        Exit Function\n"
        "    End If\n",
        "    On Error GoTo 0\n    FinalCommit = True\n    Exit Function\n")
    _control("test_32", report=damaged)


def test_14_the_prior_commit_block_is_never_captured() -> None:
    damaged = _swap(
        _REPORT,
        "    previous = SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2\n", "")
    _control("test_31", report=damaged)


def test_15_the_restore_path_is_removed() -> None:
    damaged = _swap(
        _REPORT,
        "    SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2 = previous\n", "")
    _control("test_32", report=damaged)


def test_16_a_failed_restoration_is_reported_as_safe() -> None:
    damaged = _swap(
        _REPORT,
        '    detail = "simulation: the final commit did not complete (" & cause & _\n'
        '             ") AND the previous shared block could not be restored. The " & _\n'
        '             "publication selector cannot be guaranteed and requires recovery."\n',
        '    detail = "simulation: the final commit did not verify; the previous bank stands."\n')
    _control("test_44e", report=damaged)


def test_17_results_is_written_by_the_run() -> None:
    damaged = _after(
        _REPORT,
        "    committed = True\n",
        '    modWorkbook.Sh("Results").Range("D47").Value2 = package.RequestFingerprint\n')
    _control("test_40", report=damaged)


def test_18_the_active_bank_iteration_rows_are_cleared() -> None:
    damaged = _after(
        _REPORT,
        "    Dim offset As Long, rows As Long, index As Long\n\n    offset = 0\n",
        "    SimSheet.Range(IterationRange(package.ActiveBank, 0, 1)).Value2 = vbNullString\n")
    _control("test_25", report=damaged)


def test_19_the_iteration_bank_is_written_one_row_at_a_time() -> None:
    damaged = _swap(
        _REPORT,
        "        SimSheet.Range(IterationRange(package.TargetBank, offset, rows)).Value2 = block\n",
        "        For index = 0 To rows - 1\n"
        "            SimSheet.Cells(SIM_DATA_FIRST_ITERATION_ROW + offset + index, 2).Value2 = _\n"
        "                block(index + 1, 1)\n"
        "        Next index\n")
    _control("test_28", report=damaged)


def test_20_a_partial_bank_is_selected_on_failure() -> None:
    damaged = _swap(
        _REPORT,
        "    If Not PublishCandidate(package, detail) Then\n"
        "        RunSimulation = RecordFailure(package, detail)\n"
        "        Exit Function\n"
        "    End If\n",
        "    If Not PublishCandidate(package, detail) Then\n"
        "        SharedCell(SIM_IDENTITY_ROW_ACTIVE_BANK).Value2 = package.TargetBank\n"
        "        RunSimulation = RecordFailure(package, detail)\n"
        "        Exit Function\n"
        "    End If\n")
    _control("test_33", report=damaged)


# ===========================================================================
# B. Semantics
# ===========================================================================
def test_21_the_bridge_calls_the_calculation_endpoint() -> None:
    damaged = _after(
        _CALC,
        "    detail = vbNullString\n    ' The SAME accepted preparation the endpoint uses. Not a copy of it.\n",
        "    PCCM_Calculate\n")
    _control("test_10", calc=damaged)


def test_22_the_bridge_calls_the_status_endpoint() -> None:
    damaged = _after(
        _CALC,
        "    status = DeriveStatus(package, True)\n",
        "    status = PCCM_CalculationStatus()\n")
    _control("test_10", calc=damaged)


def test_23_the_bridge_rebuilds_the_driver_factors() -> None:
    damaged = _swap(
        _CALC,
        "    drivers = package.Drivers\n",
        "    If Not BuildDriverFactors(package, detail) Then Exit Function\n"
        "    drivers = package.Drivers\n")
    _control("test_11", calc=damaged)


def test_24_the_bridge_accepts_a_stale_calculation() -> None:
    damaged = _swap(
        _CALC,
        '    If StrComp(status, CALC_STATUS_CURRENT, vbBinaryCompare) <> 0 Then\n',
        '    If StrComp(status, CALC_STATUS_INVALID, vbBinaryCompare) = 0 Then\n')
    _control("test_09", calc=damaged)


def test_25_the_stored_phase5_fingerprint_is_used_as_the_prefix() -> None:
    damaged = _swap(
        _REPORT,
        "            package.AnalyticalFingerprint, package.Iterations, package.SeedMode, _\n"
        "            package.HasSuppliedSeed, package.SuppliedSeed, _\n"
        "            package.RequestFingerprint, detail) Then\n",
        "            modCalcReport.PCCM_CalculationFingerprint(), package.Iterations, _\n"
        "            package.SeedMode, package.HasSuppliedSeed, package.SuppliedSeed, _\n"
        "            package.RequestFingerprint, detail) Then\n")
    _control("test_12", report=damaged)


def test_26_the_run_reads_the_reporting_selector() -> None:
    damaged = _after(
        _REPORT,
        "    \' 2. The two simulation controls, strictly.\n"
        "    If Not ResolveIterations(package.Iterations, detail) Then Exit Function\n"
        "    If Not ResolveSeed(package, detail) Then Exit Function\n",
        "    If modWorkbook.IsEmptyCell(modWorkbook.NamedCell( _\n"
        '            "inpSelectedConfidenceLevel")) Then Exit Function\n')
    _control("test_13", report=damaged)


def test_27_the_selector_enters_the_request_fingerprint() -> None:
    damaged = _swap(
        _REPORT,
        "    package.SeedMode = SIM_SEED_MODE_AUTO\n        package.HasSuppliedSeed = False\n",
        "    package.SeedMode = SIM_SEED_MODE_AUTO & _\n"
        '            modWorkbook.TextOf(modWorkbook.NamedCell("inpSelectedConfidenceLevel"))\n'
        "        package.HasSuppliedSeed = False\n")
    _control("test_41", report=damaged)


def test_28_the_engine_is_bypassed() -> None:
    damaged = _swap(
        _REPORT,
        "    If Not modSimEngine.SimEngineRun(package.Drivers, package.DriverCount, _\n"
        "                                     package.EffectiveSeed, package.Iterations, _\n"
        "                                     package.TotalNominal, package.TotalPv, detail) Then\n"
        "        Exit Function\n"
        "    End If\n",
        "    ReDim package.TotalNominal(0 To package.Iterations - 1)\n"
        "    ReDim package.TotalPv(0 To package.Iterations - 1)\n")
    _control("test_16", report=damaged)


def test_29_a_statistic_is_duplicated_here() -> None:
    damaged = _after(
        _REPORT, "Option Explicit\n",
        "\nPrivate Function SimReportMean(ByRef values() As Double, _\n"
        "                               ByVal count As Long) As Double\n"
        "    Dim total As Double, index As Long\n"
        "    For index = 0 To count - 1\n"
        "        total = total + values(LBound(values) + index)\n"
        "    Next index\n"
        "    SimReportMean = total / CDbl(count)\n"
        "End Function\n")
    _control("test_07", report=damaged)


def test_30_only_the_selected_rung_gets_a_contingency() -> None:
    damaged = _swap(
        _REPORT,
        "    For index = 0 To SIM_QUANTILE_COUNT - 1\n"
        "        If Not modSimStats.SimStatsContingency( _\n",
        "    For index = 1 To 1\n"
        "        If Not modSimStats.SimStatsContingency( _\n")
    _control("test_23", report=damaged)


def test_31_the_contingency_is_subtracted_here() -> None:
    damaged = _swap(
        _REPORT,
        "        If Not modSimStats.SimStatsContingency( _\n"
        "                package.NominalLadder(LBound(package.NominalLadder) + index), _\n"
        "                package.BaseNominal, value, detail) Then\n"
        "            Exit Function\n"
        "        End If\n"
        "        package.NominalContingency(index) = value\n",
        "        package.NominalContingency(index) = _\n"
        "            package.NominalLadder(LBound(package.NominalLadder) + index) - _\n"
        "            package.BaseNominal\n")
    _control("test_23", report=damaged)


def test_32_the_result_digest_is_built_from_worksheet_data() -> None:
    damaged = _swap(
        _REPORT,
        "    If Not modSimFingerprint.SimFpResultDigest(package.TotalNominal, package.TotalPv, _\n",
        "    package.TotalNominal = SimSheet.Range(\"C34:C44\").Value2\n"
        "    If Not modSimFingerprint.SimFpResultDigest(package.TotalNominal, package.TotalPv, _\n")
    _control("test_24", report=damaged)


def test_33_a_ladder_is_mutated_after_describe() -> None:
    damaged = _after(
        _REPORT,
        "    If Not SameLadder(package, detail) Then Exit Function\n",
        "    package.PvLadder(LBound(package.PvLadder)) = package.BasePv\n")
    _control("test_22", report=damaged)


def test_34_the_two_ladders_are_never_compared() -> None:
    damaged = _swap(
        _REPORT, "    If Not SameLadder(package, detail) Then Exit Function\n", "")
    _control("test_22", report=damaged)


def test_35_the_effective_seed_enters_the_auto_request_identity() -> None:
    damaged = _swap(
        _REPORT,
        "            package.HasSuppliedSeed, package.SuppliedSeed, _\n"
        "            package.RequestFingerprint, detail) Then\n",
        "            True, package.EffectiveSeed, _\n"
        "            package.RequestFingerprint, detail) Then\n")
    _control("test_15", report=damaged)


def test_36_an_auto_zero_seed_sentinel_is_introduced() -> None:
    damaged = _swap(
        _REPORT,
        "        package.SeedMode = SIM_SEED_MODE_AUTO\n        package.HasSuppliedSeed = False\n",
        "        package.SeedMode = SIM_SEED_MODE_AUTO\n"
        "        package.HasSuppliedSeed = False\n        package.SuppliedSeed = 0\n")
    _control("test_15", report=damaged)


def test_37_the_retained_arrays_are_sorted_before_publication() -> None:
    damaged = _after(
        _REPORT, "Option Explicit\n",
        "\nPrivate Sub SortAscending(ByRef series() As Double, ByVal count As Long)\n"
        "    Dim outer As Long, inner As Long, held As Double\n"
        "    For outer = 1 To count - 1\n"
        "        held = series(outer)\n"
        "        inner = outer - 1\n"
        "        Do While inner >= 0\n"
        "            If series(inner) <= held Then Exit Do\n"
        "            series(inner + 1) = series(inner)\n"
        "            inner = inner - 1\n"
        "        Loop\n"
        "        series(inner + 1) = held\n"
        "    Next outer\n"
        "End Sub\n")
    _control("test_29", report=damaged)


def test_38_the_iteration_index_uses_the_physical_bound() -> None:
    damaged = _swap(
        _REPORT,
        "            block(index + 1, 1) = SIM_DIGEST_INDEX_ORIGIN + offset + index\n",
        "            block(index + 1, 1) = LBound(package.TotalNominal) + offset + index\n")
    _control("test_28", report=damaged)


def test_39_the_attempt_result_decides_the_status() -> None:
    damaged = _swap(
        _REPORT,
        "    If Not CurrentRequestFingerprint(fingerprint, detail) Then\n"
        "        DeriveSimStatus = SIM_STATE_INVALID\n"
        "        Exit Function\n"
        "    End If\n",
        "    If StrComp(SharedText(SIM_IDENTITY_ROW_LAST_ATTEMPT_RESULT), _\n"
        "               SIM_ATTEMPT_FAILED, vbBinaryCompare) = 0 Then\n"
        "        DeriveSimStatus = SIM_STATE_INVALID\n"
        "        Exit Function\n"
        "    End If\n"
        "    If Not CurrentRequestFingerprint(fingerprint, detail) Then\n"
        "        DeriveSimStatus = SIM_STATE_INVALID\n"
        "        Exit Function\n"
        "    End If\n")
    _control("test_36", report=damaged)


def test_40_a_no_success_workbook_gets_a_fourth_state() -> None:
    damaged = _swap(
        _REPORT,
        "    \' simulation that never ran.\n"
        "    If Len(active) = 0 Then Exit Function\n",
        "    \' simulation that never ran.\n"
        '    If Len(active) = 0 Then\n        DeriveSimStatus = "UNSELECTED"\n'
        "        Exit Function\n    End If\n")
    _control("test_36", report=damaged)


def test_41_the_current_request_path_allocates_a_nonce() -> None:
    damaged = _swap(
        _REPORT,
        "    If Not ResolveIterations(package.Iterations, detail) Then Exit Function\n"
        "    If Not ResolveSeed(package, detail) Then Exit Function\n"
        "    If Not modSimFingerprint.SimFpBuildRequestFingerprint( _\n",
        "    If Not ResolveIterations(package.Iterations, detail) Then Exit Function\n"
        "    If Not ResolveSeed(package, detail) Then Exit Function\n"
        "    If Not AllocateAutoNonce(package, detail) Then Exit Function\n"
        "    If Not modSimFingerprint.SimFpBuildRequestFingerprint( _\n")
    _control("test_37", report=damaged)


def test_42_a_stored_accessor_recomputes() -> None:
    damaged = _swap(
        _REPORT,
        "    PCCM_SimulationRequestFingerprint = _\n"
        "        ActiveSnapshotText(SIM_IDENTITY_ROW_REQUEST_FINGERPRINT)\n",
        "    Dim fingerprint As String, detail As String\n"
        "    If Not CurrentRequestFingerprint(fingerprint, detail) Then Exit Function\n"
        "    PCCM_SimulationRequestFingerprint = fingerprint\n")
    _control("test_38", report=damaged)


def test_43_a_read_accessor_writes() -> None:
    damaged = _swap(
        _REPORT,
        "Public Function PCCM_SimulationAttemptResult() As String\n",
        "Public Function PCCM_SimulationAttemptResult() As String\n"
        "    SharedCell(SIM_IDENTITY_ROW_LAST_ATTEMPT_DETAIL).Value2 = vbNullString\n")
    _control("test_39", report=damaged)


def test_44_the_invocation_envelope_is_removed() -> None:
    damaged = _swap(
        _REPORT, "    state = modAppState.CaptureAppState()\n    stateCaptured = True\n", "")
    _control("test_42", report=damaged)


def test_45_a_committed_run_is_rewritten_as_failed_by_a_cleanup_problem() -> None:
    damaged = _swap(
        _CALC if False else _REPORT,
        "    If committed Then\n"
        '        CleanupOutcome = modAppState.Failed("Run Simulation", _\n',
        "    If committed Then\n"
        "        SharedCell(SIM_IDENTITY_ROW_LAST_ATTEMPT_RESULT).Value2 = SIM_ATTEMPT_FAILED\n"
        '        CleanupOutcome = modAppState.Failed("Run Simulation", _\n')
    _control("test_43", report=damaged)


def test_46_a_named_failpoint_is_removed() -> None:
    damaged = _swap(
        _REPORT, "    modAppState.FailPointCheck FAILPOINT_SIM_AFTER_NONCE\n", "")
    _control("test_44", report=damaged)


# ===========================================================================
# C. D6-11 and the public surface
# ===========================================================================
def test_47_the_endpoint_construct_stays_globally_forbidden() -> None:
    damaged = _swap(
        _STRUCTURE,
        '    - construct: "RunSimulation"\n      allowed_in:\n        - "modSimReport"\n',
        '    - "RunSimulation"\n')
    _control("test_03", structure=damaged)


def test_48_the_endpoint_is_allowed_in_a_second_module() -> None:
    damaged = _swap(
        _STRUCTURE,
        '    - construct: "RunSimulation"\n      allowed_in:\n        - "modSimReport"\n',
        '    - construct: "RunSimulation"\n      allowed_in:\n'
        '        - "modSimReport"\n        - "modCalcReport"\n')
    _control("test_03", structure=damaged)


def test_49_the_endpoint_is_granted_to_the_wrong_owner() -> None:
    damaged = _swap(
        _STRUCTURE,
        '    - construct: "RunSimulation"\n      allowed_in:\n        - "modSimReport"\n',
        '    - construct: "RunSimulation"\n      allowed_in:\n        - "modSimEngine"\n')
    _control("test_03", structure=damaged)


def test_50_the_generator_scope_is_moved() -> None:
    damaged = _swap(
        _STRUCTURE,
        '    - construct: "MRG32k3a"\n      allowed_in:\n        - "modSimRng"\n',
        '    - construct: "MRG32k3a"\n      allowed_in:\n        - "modSimReport"\n')
    _control("test_03", structure=damaged)


def test_51_the_percentile_construct_is_scoped() -> None:
    damaged = _swap(
        _STRUCTURE,
        '    - "Percentile"\n',
        '    - construct: "Percentile"\n      allowed_in:\n        - "modSimReport"\n')
    _control("test_03", structure=damaged)


def test_52_an_extra_public_accessor_appears() -> None:
    damaged = _after(
        _REPORT,
        "Public Function PCCM_SimulationAttemptDetail() As String\n"
        "    PCCM_SimulationAttemptDetail = SharedText(SIM_IDENTITY_ROW_LAST_ATTEMPT_DETAIL)\n"
        "End Function\n",
        "\nPublic Function PCCM_SimulationRunId() As String\n"
        "    PCCM_SimulationRunId = ActiveSnapshotText(SIM_IDENTITY_ROW_RUN_ID)\n"
        "End Function\n")
    _control("test_04", report=damaged)


def test_53_a_settled_accessor_is_renamed() -> None:
    damaged = _REPORT.replace("PCCM_SimulationAttemptDetail",
                              "PCCM_SimulationLastDetail")
    _control("test_04", report=damaged)


def test_54_an_extra_public_helper_appears_in_the_reporter() -> None:
    damaged = _after(
        _REPORT, "Option Explicit\n",
        "\nPublic Function SimReportActiveBank() As String\n"
        '    SimReportActiveBank = SharedText(SIM_IDENTITY_ROW_ACTIVE_BANK)\n'
        "End Function\n")
    _control("test_04", report=damaged)


def test_55_an_extra_public_helper_appears_beside_the_bridge() -> None:
    damaged = _after(
        _CALC,
        "    CalcPrepareSimulationInputs = True\nEnd Function\n",
        "\nPublic Function CalcCurrentAnalyticalFingerprint() As String\n"
        "    Dim package As CalculationPackage, detail As String\n"
        "    If PrepareCurrentCalculation(package, detail) Then\n"
        "        CalcCurrentAnalyticalFingerprint = package.Fingerprint\n"
        "    End If\n"
        "End Function\n")
    _control("test_08", calc=damaged)


if __name__ == "__main__":  # pragma: no cover
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))


# ===========================================================================
# K. THE TRANSACTION FAILURE PATHS (Step-12 settlement)
#
# Every control here damages a recovery path, not an adjacency. The question
# each asks is "can a COM failure still reach the attempt axis, and can the
# prior published block still come back" - never "is this token near that one".
# ===========================================================================
def test_56_the_candidate_transaction_loses_its_error_envelope() -> None:
    damaged = _swap(_REPORT, "    On Error GoTo CandidateFailed\n", "")
    _control("test_44a", report=damaged)


def test_57_the_candidate_envelope_becomes_a_blanket_suppressor() -> None:
    damaged = _swap(_REPORT, "    On Error GoTo CandidateFailed\n",
                    "    On Error Resume Next\n")
    _control("test_44g", report=damaged)


def test_58_the_candidate_handler_reports_success() -> None:
    damaged = _swap(
        _REPORT,
        "CandidateFailed:\n    failure = Err.Description\n",
        "CandidateFailed:\n    PublishCandidate = True\n    failure = Err.Description\n")
    _control("test_44a", report=damaged)


def test_59_the_candidate_handler_erases_the_partial_bank() -> None:
    damaged = _swap(
        _REPORT,
        "CandidateFailed:\n    failure = Err.Description\n",
        "CandidateFailed:\n"
        "    SimSheet.Range(SnapshotRange(package.TargetBank)).ClearContents\n"
        "    failure = Err.Description\n")
    _control("test_44a", report=damaged)


def test_60_the_candidate_failpoint_moves_back_before_publication() -> None:
    damaged = _swap(
        _REPORT, "    modAppState.FailPointCheck FAILPOINT_SIM_CANDIDATE_BANK\n", "")
    damaged = _swap(
        damaged,
        "    package.TargetBank = InactiveBank(package.ActiveBank)\n",
        "    package.TargetBank = InactiveBank(package.ActiveBank)\n"
        "    modAppState.FailPointCheck FAILPOINT_SIM_CANDIDATE_BANK\n")
    _control("test_44", report=damaged)


def test_61_the_candidate_failpoint_moves_before_the_candidate_writes() -> None:
    damaged = _swap(
        _REPORT,
        "    modAppState.FailPointCheck FAILPOINT_SIM_CANDIDATE_BANK\n\n"
        "    If Not VerifyCandidateBank(", "\n    If Not VerifyCandidateBank(")
    damaged = _swap(
        damaged,
        "    SimSheet.Range(SnapshotRange(package.TargetBank)).Value2 = snapshot\n",
        "    modAppState.FailPointCheck FAILPOINT_SIM_CANDIDATE_BANK\n"
        "    SimSheet.Range(SnapshotRange(package.TargetBank)).Value2 = snapshot\n")
    _control("test_44", report=damaged)


def test_62_the_candidate_failpoint_moves_after_verification() -> None:
    damaged = _swap(
        _REPORT,
        "    modAppState.FailPointCheck FAILPOINT_SIM_CANDIDATE_BANK\n\n"
        "    If Not VerifyCandidateBank(package, snapshot, summary, contingency, detail) Then\n"
        "        Exit Function\n"
        "    End If\n",
        "    If Not VerifyCandidateBank(package, snapshot, summary, contingency, detail) Then\n"
        "        Exit Function\n"
        "    End If\n"
        "    modAppState.FailPointCheck FAILPOINT_SIM_CANDIDATE_BANK\n")
    _control("test_44", report=damaged)


def test_63_the_final_commit_loses_its_error_envelope() -> None:
    damaged = _swap(_REPORT, "    On Error GoTo CommitFailed\n", "")
    _control("test_44d", report=damaged)


def test_64_the_final_failpoint_moves_back_before_the_write() -> None:
    damaged = _swap(
        _REPORT,
        "    SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2 = block\n"
        "    modAppState.FailPointCheck FAILPOINT_SIM_FINAL_COMMIT\n",
        "    modAppState.FailPointCheck FAILPOINT_SIM_FINAL_COMMIT\n"
        "    SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2 = block\n")
    _control("test_31", report=damaged)


def test_65_the_final_failpoint_moves_after_verification() -> None:
    damaged = _swap(
        _REPORT,
        "    modAppState.FailPointCheck FAILPOINT_SIM_FINAL_COMMIT\n"
        "    If SameBlock(SIM_FINAL_COMMIT_RANGE, block, 9, 1) Then\n",
        "    If SameBlock(SIM_FINAL_COMMIT_RANGE, block, 9, 1) Then\n"
        "        modAppState.FailPointCheck FAILPOINT_SIM_FINAL_COMMIT\n")
    _control("test_44", report=damaged)


def test_66_a_raised_commit_write_exits_without_restoring() -> None:
    """The Step-11 defect itself, planted back: the handler just reports."""
    damaged = _swap(
        _REPORT,
        "CommitFailed:\n    cause = Err.Description\n    On Error GoTo 0\n",
        "CommitFailed:\n    cause = Err.Description\n    On Error GoTo 0\n"
        "    detail = \"simulation: the final commit raised: \" & cause\n"
        "    Exit Function\n")
    _control("test_44d", report=damaged)


def test_67_a_verification_mismatch_exits_without_restoring() -> None:
    damaged = _swap(
        _REPORT,
        "    cause = \"the committed block did not verify\"\n"
        "    GoTo RestorePrevious\n",
        "    cause = \"the committed block did not verify\"\n"
        "    detail = cause\n"
        "    Exit Function\n")
    _control("test_44d", report=damaged)


def test_68_the_verification_read_sits_outside_the_envelope() -> None:
    """A raised SameBlock must restore too, so it must be inside the handler."""
    damaged = _swap(
        _REPORT,
        "    On Error GoTo CommitFailed\n"
        "    SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2 = block\n"
        "    modAppState.FailPointCheck FAILPOINT_SIM_FINAL_COMMIT\n"
        "    If SameBlock(SIM_FINAL_COMMIT_RANGE, block, 9, 1) Then\n",
        "    On Error GoTo CommitFailed\n"
        "    SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2 = block\n"
        "    On Error GoTo 0\n"
        "    modAppState.FailPointCheck FAILPOINT_SIM_FINAL_COMMIT\n"
        "    If SameBlock(SIM_FINAL_COMMIT_RANGE, block, 9, 1) Then\n")
    _control("test_44d", report=damaged)


def test_69_the_prior_block_is_never_captured() -> None:
    damaged = _swap(
        _REPORT,
        "    On Error GoTo CaptureFailed\n"
        "    previous = SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2\n"
        "    On Error GoTo 0\n", "")
    _control("test_44c", report=damaged)


def test_70_the_capture_has_no_handler_of_its_own() -> None:
    damaged = _swap(_REPORT, "    On Error GoTo CaptureFailed\n", "")
    _control("test_44c", report=damaged)


def test_71_a_failed_capture_writes_an_unset_block_over_the_publication() -> None:
    # Anchored on the handler LABEL, so only the capture path is damaged: the
    # same statement opens several handlers and a blanket replace would mutate
    # all of them and prove less.
    at = _REPORT.index("CaptureFailed:")
    line = "    failure = Err.Description\n"
    cut = _REPORT.index(line, at)
    damaged = (_REPORT[:cut]
               + "    SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2 = previous\n"
               + _REPORT[cut:])
    assert damaged != _REPORT
    assert damaged.count("Range(SIM_FINAL_COMMIT_RANGE).Value2 = previous") == \
        _REPORT.count("Range(SIM_FINAL_COMMIT_RANGE).Value2 = previous") + 1
    _control("test_44c", report=damaged)


def test_72_the_restore_write_is_removed() -> None:
    damaged = _swap(
        _REPORT,
        "    SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2 = previous\n"
        "    If SameBlock(SIM_FINAL_COMMIT_RANGE, previous, 9, 1) Then\n",
        "    If SameBlock(SIM_FINAL_COMMIT_RANGE, previous, 9, 1) Then\n")
    _control("test_32", report=damaged)


def test_73_the_restore_is_never_verified() -> None:
    damaged = _swap(
        _REPORT,
        "    If SameBlock(SIM_FINAL_COMMIT_RANGE, previous, 9, 1) Then\n",
        "    If True Then\n")
    _control("test_44d", report=damaged)


def test_74_a_failed_restoration_claims_the_publication_is_safe() -> None:
    damaged = _swap(
        _REPORT,
        "RestoreFailed:\n"
        "    failure = Err.Description\n"
        "    On Error GoTo 0\n"
        "    detail = \"simulation: the final commit did not complete (\" & cause & _\n"
        "             \") AND the previous shared block could not be restored: \" & failure & _\n"
        "             \". The publication selector cannot be guaranteed and requires recovery.\"\n",
        "RestoreFailed:\n"
        "    failure = Err.Description\n"
        "    On Error GoTo 0\n"
        "    detail = \"simulation: the final commit did not complete; the previous \" & _\n"
        "             \"published bank remains authoritative\"\n")
    _control("test_44e", report=damaged)


def test_75_the_restore_itself_has_no_handler() -> None:
    damaged = _swap(_REPORT, "    On Error GoTo RestoreFailed\n", "")
    _control("test_44e", report=damaged)


def test_76_a_candidate_failure_bypasses_the_attempt_record() -> None:
    damaged = _swap(
        _REPORT,
        "    If Not PublishCandidate(package, detail) Then\n"
        "        RunSimulation = RecordFailure(package, detail)\n"
        "        Exit Function\n"
        "    End If\n",
        "    If Not PublishCandidate(package, detail) Then\n"
        "        Err.Raise vbObjectError + 1, , detail\n"
        "    End If\n")
    _control("test_44b", report=damaged)


def test_77_a_commit_failure_bypasses_the_attempt_record() -> None:
    damaged = _swap(
        _REPORT,
        "    If Not FinalCommit(package, detail) Then\n"
        "        RunSimulation = RecordFailure(package, detail)\n"
        "        Exit Function\n"
        "    End If\n",
        "    If Not FinalCommit(package, detail) Then\n"
        "        Err.Raise vbObjectError + 2, , detail\n"
        "    End If\n")
    _control("test_44f", report=damaged)


def test_78_the_attempt_record_writes_the_publication_rows() -> None:
    """A FAILED attempt must not reach D22 or D30 - those belong to the commit."""
    body = conformance._procedure("WriteAttemptBlock")
    anchor = body[body.index("    SimSheet.Range(AttemptRange())"):]
    anchor = anchor[: anchor.index("\n") + 1]
    damaged = _swap(
        _REPORT, anchor,
        "    SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2 = block\n" + anchor)
    _control("test_33", report=damaged)


# ===========================================================================
# L. THE AUTO-NONCE FAILURE PATH (Step-12 settlement, second finding)
#
# FailPointCheck RAISES. A naked call after the nonce was spent left the run
# through the invocation handler with no attempt record at all - the same class
# the candidate and commit stages carried, in the one stage that had already
# consumed something irreversible.
# ===========================================================================
_NONCE_FAILPOINT = "    modAppState.FailPointCheck FAILPOINT_SIM_AFTER_NONCE\n"


def test_79_the_after_nonce_failpoint_returns_to_run_simulation() -> None:
    """THE ORIGINAL DEFECT, planted back exactly as it shipped in 4df2af3."""
    damaged = _swap(_REPORT, _NONCE_FAILPOINT, "")
    damaged = _swap(
        damaged,
        "    If Not AllocateAutoNonce(package, detail) Then\n"
        "        RunSimulation = RecordRefusal(package, detail)\n"
        "        Exit Function\n"
        "    End If\n",
        "    If Not AllocateAutoNonce(package, detail) Then\n"
        "        RunSimulation = RecordRefusal(package, detail)\n"
        "        Exit Function\n"
        "    End If\n"
        "    modAppState.FailPointCheck FAILPOINT_SIM_AFTER_NONCE\n")
    _control("test_44", report=damaged)


def test_80_the_after_nonce_failpoint_fires_before_the_verification() -> None:
    damaged = _swap(_REPORT, _NONCE_FAILPOINT, "")
    damaged = _swap(
        damaged,
        "        SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 = package.ConsumedNonce + 1\n",
        "        SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 = package.ConsumedNonce + 1\n"
        "        modAppState.FailPointCheck FAILPOINT_SIM_AFTER_NONCE\n")
    _control("test_44", report=damaged)


def test_81_the_after_nonce_failpoint_fires_before_the_consumed_mark() -> None:
    damaged = _swap(_REPORT, _NONCE_FAILPOINT, "")
    damaged = _swap(
        damaged,
        "        package.NonceConsumed = True\n",
        "        modAppState.FailPointCheck FAILPOINT_SIM_AFTER_NONCE\n"
        "        package.NonceConsumed = True\n")
    _control("test_44", report=damaged)


def test_82_the_nonce_allocation_loses_its_error_envelope() -> None:
    damaged = _swap(_REPORT, "    On Error GoTo AllocationFailed\n", "")
    _control("test_44h", report=damaged)


def test_83_the_nonce_envelope_becomes_a_blanket_suppressor() -> None:
    damaged = _swap(_REPORT, "    On Error GoTo AllocationFailed\n",
                    "    On Error Resume Next\n")
    _control("test_44g", report=damaged)


def test_84_the_counter_write_sits_outside_the_envelope() -> None:
    """A raised write must not escape, so it must stay inside the handler."""
    damaged = _swap(
        _REPORT,
        "    On Error GoTo AllocationFailed\n\n"
        "    If package.HasSuppliedSeed Then\n",
        "    If package.HasSuppliedSeed Then\n")
    damaged = _swap(
        damaged,
        "        package.NonceConsumed = True\n"
        "    End If\n",
        "        package.NonceConsumed = True\n"
        "    End If\n"
        "    On Error GoTo AllocationFailed\n")
    _control("test_44h", report=damaged)


def test_85_the_nonce_handler_reports_success() -> None:
    damaged = _swap(
        _REPORT,
        "AllocationFailed:\n    failure = Err.Description\n",
        "AllocationFailed:\n    AllocateAutoNonce = True\n    failure = Err.Description\n")
    _control("test_44h", report=damaged)


def test_86_an_after_persist_failure_bypasses_the_attempt_recorder() -> None:
    damaged = _swap(
        _REPORT,
        "    If Not AllocateAutoNonce(package, detail) Then\n"
        "        RunSimulation = RecordRefusal(package, detail)\n"
        "        Exit Function\n"
        "    End If\n",
        "    If Not AllocateAutoNonce(package, detail) Then\n"
        "        Err.Raise vbObjectError + 3, , detail\n"
        "    End If\n")
    _control("test_44i", report=damaged)


def test_87_the_counter_is_rolled_back_on_failure() -> None:
    damaged = _swap(
        _REPORT,
        "AllocationFailed:\n    failure = Err.Description\n",
        "AllocationFailed:\n"
        "    SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 = package.ConsumedNonce - 1\n"
        "    failure = Err.Description\n")
    _control("test_44j", report=damaged)


def test_88_the_handler_claims_the_sequence_can_be_reused() -> None:
    """The raised-allocation detail collapses to one branch that says nothing."""
    body = conformance._procedure("AllocateAutoNonce")
    handler = body[body.index("AllocationFailed:"):]
    branch = handler[handler.index("    If package.NonceAllocated Then"):]
    branch = branch[: branch.index("    End If\n") + len("    End If\n")]
    damaged = _swap(
        _REPORT, branch,
        '    detail = "simulation: seed allocation did not complete: " & failure\n')
    _control("test_44h", report=damaged)


def test_89_a_fourth_failpoint_is_introduced() -> None:
    damaged = _swap(
        _REPORT,
        'Public Const FAILPOINT_SIM_FINAL_COMMIT As String = "Phase6FinalCommit"\n',
        'Public Const FAILPOINT_SIM_FINAL_COMMIT As String = "Phase6FinalCommit"\n'
        'Public Const FAILPOINT_SIM_EXTRA As String = "Phase6Extra"\n')
    damaged = _swap(
        damaged,
        "    package.Stamp = Now\n",
        "    modAppState.FailPointCheck FAILPOINT_SIM_EXTRA\n"
        "    package.Stamp = Now\n")
    _control("test_44k", report=damaged)


def test_90_sampling_begins_before_the_allocation_succeeds() -> None:
    damaged = _swap(
        _REPORT,
        "    If Not AllocateAutoNonce(package, detail) Then\n"
        "        RunSimulation = RecordRefusal(package, detail)\n"
        "        Exit Function\n"
        "    End If\n\n"
        "    ' 5-11. The accepted kernels, in order, entirely in memory.\n"
        "    If Not RunKernels(package, detail) Then\n",
        "    ' 5-11.\n"
        "    If Not RunKernels(package, detail) Then\n")
    _control("test_44i", report=damaged)


# ===========================================================================
# M. THE AUTO-ALLOCATION AUDIT IDENTITY (Step-12 settlement, third finding)
#
# `NonceConsumed` alone conflated "the counter advance is verified" with "this
# run has claimed a nonce". Every verification failure AFTER the counter write
# therefore blanked the effective seed and the AUTO nonce on the attempt record,
# leaving an advanced counter with no trace - the audit hole the retained
# authority forbids.
# ===========================================================================
_ALLOCATED = "        package.NonceAllocated = True\n"


def test_91_the_original_single_boolean_shape_is_rejected() -> None:
    """THE DEFECT AS IT SHIPPED IN e574fdb, restored verbatim in behaviour.

    One flag, set only after verification, and the attempt record gated on it.
    """
    damaged = _swap(_REPORT, _ALLOCATED, "")
    damaged = _swap(
        damaged,
        "        If package.HasSuppliedSeed Or package.NonceAllocated Then\n",
        "        If package.HasSuppliedSeed Or package.NonceConsumed Then\n")
    damaged = _swap(
        damaged,
        "    If package.NonceAllocated Then\n"
        "        block(5, 1) = package.ConsumedNonce\n",
        "    If package.NonceConsumed Then\n"
        "        block(5, 1) = package.ConsumedNonce\n")
    _control("test_44h3", report=damaged)


def test_92_the_allocation_is_claimed_only_after_the_write() -> None:
    """A raised assignment then loses the identity it may already have spent."""
    damaged = _swap(
        _REPORT,
        _ALLOCATED +
        "        SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 = package.ConsumedNonce + 1\n",
        "        SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 = package.ConsumedNonce + 1\n"
        + _ALLOCATED)
    _control("test_44h2", report=damaged)


def test_93_the_allocation_is_claimed_only_on_the_success_path() -> None:
    damaged = _swap(_REPORT, _ALLOCATED, "")
    damaged = _swap(
        damaged,
        "        package.NonceConsumed = True\n",
        "        package.NonceAllocated = True\n"
        "        package.NonceConsumed = True\n")
    _control("test_44h2", report=damaged)


def test_94_the_effective_seed_is_blanked_on_a_post_write_failure() -> None:
    damaged = _swap(
        _REPORT,
        "        If package.HasSuppliedSeed Or package.NonceAllocated Then\n",
        "        If package.HasSuppliedSeed Or package.NonceConsumed Then\n")
    _control("test_44h3", report=damaged)


def test_95_the_auto_nonce_is_blanked_on_a_post_write_failure() -> None:
    damaged = _swap(
        _REPORT,
        "    If package.NonceAllocated Then\n"
        "        block(5, 1) = package.ConsumedNonce\n",
        "    If package.NonceConsumed Then\n"
        "        block(5, 1) = package.ConsumedNonce\n")
    _control("test_44h3", report=damaged)


def test_96_a_post_write_read_failure_claims_nothing_was_written() -> None:
    body = conformance._procedure("AllocateAutoNonce")
    start = body.index('            detail = "simulation: the AUTO nonce advance was written')
    end = body.index("            Exit Function", start)
    damaged = _swap(
        _REPORT, body[start:end],
        '            detail = "simulation: the AUTO nonce advance did not persist"\n')
    _control("test_44h4", report=damaged)


def test_97_a_mismatch_claims_nothing_was_written() -> None:
    body = conformance._procedure("AllocateAutoNonce")
    start = body.index('            detail = "simulation: the AUTO nonce advance did not read back')
    end = body.index("            Exit Function", start)
    damaged = _swap(
        _REPORT, body[start:end],
        '            detail = "simulation: the AUTO nonce advance did not persist"\n')
    _control("test_44h4", report=damaged)


def test_98_the_published_snapshot_claims_an_unverified_nonce() -> None:
    """The opposite error: a PUBLISHED record may claim only proven consumption."""
    damaged = _swap(
        _REPORT,
        "    If package.NonceConsumed Then\n"
        "        built(8, 1) = package.ConsumedNonce\n",
        "    If package.NonceAllocated Then\n"
        "        built(8, 1) = package.ConsumedNonce\n")
    _control("test_44h3", report=damaged)


def test_99_the_commit_block_claims_an_unverified_nonce() -> None:
    damaged = _swap(
        _REPORT,
        "    If package.NonceConsumed Then\n"
        "        built(6, 1) = package.ConsumedNonce\n",
        "    If package.NonceAllocated Then\n"
        "        built(6, 1) = package.ConsumedNonce\n")
    _control("test_44h3", report=damaged)


def test_100_fixed_mode_writes_the_auto_counter() -> None:
    damaged = _swap(
        _REPORT,
        "    If package.HasSuppliedSeed Then\n"
        "        package.EffectiveSeed = package.SuppliedSeed\n",
        "    If package.HasSuppliedSeed Then\n"
        "        package.EffectiveSeed = package.SuppliedSeed\n"
        "        package.NonceAllocated = True\n"
        "        SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 = _\n"
        "            package.ConsumedNonce + 1\n")
    _control("test_44h2", report=damaged)


def test_101_sampling_begins_despite_a_failed_verification() -> None:
    damaged = _swap(
        _REPORT,
        "            Exit Function\n"
        "        End If\n"
        "        If stored <> package.ConsumedNonce + 1 Then\n",
        "        End If\n"
        "        If False Then\n")
    _control("test_44h4", report=damaged)


def test_102_the_audit_writer_grows_a_handler_without_settling_storage() -> None:
    """If storage IS settled, the guarantee must say so - not quietly change."""
    damaged = _swap(
        _REPORT,
        "    SimSheet.Range(AttemptRange()).Value2 = block\n",
        "    On Error Resume Next\n"
        "    SimSheet.Range(AttemptRange()).Value2 = block\n")
    _control("test_44g", report=damaged)
