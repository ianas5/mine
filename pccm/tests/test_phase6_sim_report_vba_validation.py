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
_NONCE = conformance.NONCE_BAS.read_text(encoding="utf-8")
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
             structure: str | None = None, nonce: str | None = None):
    saved = (conformance.REPORT_BAS, conformance.CALC_REPORT_BAS, conformance.SPEC,
             dict(conformance._CACHE), conformance.NONCE_BAS)
    conformance._CACHE.clear()
    temp = Path(tempfile.mkdtemp(prefix="pccm-step11-mutation-"))
    if nonce is not None:
        assert nonce != _NONCE, "the mutation changed nothing"
        target = temp / "modSimNonce.bas"
        target.write_text(nonce, encoding="utf-8")
        conformance.NONCE_BAS = target
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
        conformance.NONCE_BAS = saved[4]
        conformance._CACHE.clear()
        conformance._CACHE.update(saved[3])

    return restore


def _control(expected: str, report: str | None = None, calc: str | None = None,
             structure: str | None = None, nonce: str | None = None) -> None:
    restore = _install(report, calc, structure, nonce)
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


def test_08_run_id_exhaustion_is_checked_before_the_nonce_is_touched() -> None:
    """Moving the headroom check after allocation would burn a sequence on a run
    that could never have been committed."""
    damaged = _swap(
        _REPORT,
        "    If lastRunId >= SIM_RUN_ID_MAXIMUM Then\n"
        '        detail = "simulation: the run identity counter is exhausted; no further run " & _\n'
        '                 "can be identified or committed"\n'
        "        Exit Function\n"
        "    End If\n"
        "    package.CandidateRunId = lastRunId + 1\n",
        "    package.CandidateRunId = lastRunId + 1\n")
    _control("test_18", report=damaged)

def test_09_sampling_begins_before_the_nonce_is_persisted() -> None:
    damaged = _swap(
        _NONCE,
        "    If Not RunAllocationTransaction(autoNonce, state, detail) "
        "Then Exit Function\n", "")
    _control("test_44h2", nonce=damaged)


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
        "    If package.AutoIdentityKnown Then\n"
        "        block(5, 1) = package.ConsumedNonce\n"
        "    Else\n"
        "        block(5, 1) = vbNullString\n"
        "    End If\n",
        "    block(5, 1) = vbNullString\n")
    _control("test_44h3", report=damaged)


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
        _REPORT, "    modAppState.FailPointCheck FAILPOINT_SIM_CANDIDATE_BANK\n", "")
    _control("test_44", report=damaged)


def test_46b_the_owned_failpoint_is_removed_from_the_nonce_module() -> None:
    damaged = _swap(_NONCE, _INJECT, "")
    _control("test_44k", nonce=damaged)


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
# L. THE AUTO-NONCE TRANSACTION, AT ITS OWNER (Step-12 settlement)
#
# The transaction moved to modSimNonce, and with it every control below. The
# previous round's controls encoded a REJECTED definition - allocation claimed
# before the write - and are replaced rather than re-anchored: a control that
# pins the wrong authority is worse than no control.
# ===========================================================================
_INJECT = "    modAppState.FailPointCheck FAILPOINT_SIM_AFTER_NONCE\n"


def test_79_the_after_nonce_failpoint_returns_to_the_orchestrator() -> None:
    """A naked raising call in RunSimulation bypasses the attempt axis."""
    damaged = _swap(_NONCE, _INJECT, "")
    report = _swap(
        _REPORT,
        "    If Not AllocateAutoNonce(package, detail) Then\n",
        "    modAppState.FailPointCheck modSimNonce.FAILPOINT_SIM_AFTER_NONCE\n"
        "    If Not AllocateAutoNonce(package, detail) Then\n")
    _control("test_44", report=report, nonce=damaged)


def test_80_the_injection_fires_before_the_advance_is_established() -> None:
    damaged = _swap(_NONCE, _INJECT, "")
    damaged = _swap(
        damaged,
        "    If Not RunAllocationTransaction(autoNonce, state, detail) "
        "Then Exit Function\n",
        _INJECT +
        "    If Not RunAllocationTransaction(autoNonce, state, detail) "
        "Then Exit Function\n")
    _control("test_44", nonce=damaged)


def test_81_the_injection_fires_before_the_seed_is_derived() -> None:
    damaged = _swap(_NONCE, _INJECT, "")
    damaged = _swap(
        damaged,
        "    If Not ResolveNextNonce(autoNonce, state, detail) Then Exit Function\n",
        _INJECT +
        "    If Not ResolveNextNonce(autoNonce, state, detail) Then Exit Function\n")
    _control("test_44h2", nonce=damaged)


def test_82_the_nonce_entry_loses_its_error_envelope() -> None:
    damaged = _swap(_NONCE, "    On Error GoTo AllocationFailed\n", "")
    _control("test_44", nonce=damaged)


def test_83_the_nonce_envelope_becomes_a_blanket_suppressor() -> None:
    damaged = _swap(_NONCE, "    On Error GoTo AllocationFailed\n",
                    "    On Error Resume Next\n")
    _control("test_44g", nonce=damaged)


def test_84_the_counter_write_sits_outside_its_envelope() -> None:
    damaged = _swap(_NONCE, "    On Error GoTo StepRaised\n", "")
    _control("test_44h", nonce=damaged)


def test_85_the_nonce_handler_reports_success() -> None:
    damaged = _swap(
        _NONCE,
        "AllocationFailed:\n    failure = Err.Description\n",
        "AllocationFailed:\n    SimNonceAllocate = True\n    failure = Err.Description\n")
    _control("test_44g", nonce=damaged)


def test_86_an_allocation_failure_bypasses_the_attempt_recorder() -> None:
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
        _NONCE,
        "AllocationFailed:\n    failure = Err.Description\n",
        "AllocationFailed:\n"
        "    SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 = autoNonce - 1\n"
        "    failure = Err.Description\n")
    _control("test_44j", nonce=damaged)


def test_88_the_reconciliation_becomes_a_retry_loop() -> None:
    """One observation is the authority; a loop still would not decide it."""
    damaged = _swap(
        _NONCE,
        "    On Error GoTo ObservationRaised\n"
        "    If Not ReadPersistedNonce(stored, probe) Then\n",
        "    On Error GoTo ObservationRaised\n"
        "    Do While Not ReadPersistedNonce(stored, probe)\n"
        "    Loop\n"
        "    If Not ReadPersistedNonce(stored, probe) Then\n", count=1)
    _control("test_44h4", nonce=damaged)


def test_89_a_fourth_failpoint_is_introduced() -> None:
    damaged = _swap(
        _NONCE,
        'Public Const FAILPOINT_SIM_AFTER_NONCE As String = "Phase6AfterNoncePersisted"\n',
        'Public Const FAILPOINT_SIM_AFTER_NONCE As String = "Phase6AfterNoncePersisted"\n'
        'Public Const FAILPOINT_SIM_EXTRA As String = "Phase6Extra"\n')
    damaged = _swap(
        damaged, "    effectiveSeed = 0\n",
        "    modAppState.FailPointCheck FAILPOINT_SIM_EXTRA\n    effectiveSeed = 0\n")
    _control("test_44k", nonce=damaged)


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
# M. THE THREE OBSERVATIONS AND THE AUDIT IDENTITY BY STATE
# ===========================================================================
def test_91_the_rejected_pre_write_allocation_flag_returns() -> None:
    """The e574fdb/db85748 shape: allocation claimed before the write."""
    damaged = _swap(
        _NONCE, "    identityKnown = True\n",
        "    identityKnown = True\n    Dim NonceAllocated As Boolean\n"
        "    NonceAllocated = True\n")
    _control("test_44h2", nonce=damaged)


def test_92_the_consumed_state_is_set_without_an_observed_match() -> None:
    damaged = _swap(
        _NONCE,
        "    If stored = nonce + 1 Then\n        state = SIM_NONCE_STATE_CONSUMED\n",
        "    If True Then\n        state = SIM_NONCE_STATE_CONSUMED\n")
    _control("test_44h4", nonce=damaged)


def test_93_the_m_observation_is_classified_as_consumed() -> None:
    damaged = _swap(
        _NONCE,
        "    If stored = nonce Then\n"
        "        ' The advance did not land. Nothing was consumed, so the next run may\n"
        "        ' legitimately take this nonce again - and saying otherwise would be a\n"
        "        ' promise the source cannot keep.\n"
        "        state = SIM_NONCE_STATE_PRE_ALLOCATION\n",
        "    If stored = nonce Then\n"
        "        state = SIM_NONCE_STATE_CONSUMED\n")
    _control("test_44h4", nonce=damaged)


def test_94_the_pre_allocation_arm_promises_non_reuse() -> None:
    """The statement the source cannot make: m WILL be reissued."""
    damaged = _swap(
        _NONCE,
        '        detail = detail & ". Nonce " & CStr(nonce) & " was NOT consumed and no " & _\n'
        '                 "sampling was started; a retry may take it again."\n',
        '        detail = detail & ". Nonce " & CStr(nonce) & " will not be reused."\n')
    _control("test_44h4", nonce=damaged)


def test_95_an_impossible_counter_reading_is_normalised() -> None:
    body = conformance._procedure("Classify", conformance.NONCE_BAS)
    tail = body[body.index("    state = SIM_NONCE_STATE_RECOVERY"):]
    damaged = _swap(
        _NONCE, tail,
        "    state = SIM_NONCE_STATE_CONSUMED\n"
        "    detail = \"simulation: the counter was normalised\"\n")
    _control("test_44h4", nonce=damaged)


def test_96_the_indeterminate_case_claims_to_be_unconsumed() -> None:
    damaged = _swap(
        _NONCE,
        '    state = SIM_NONCE_STATE_INDETERMINATE\n',
        '    state = SIM_NONCE_STATE_PRE_ALLOCATION\n')
    _control("test_44h4", nonce=damaged)


def test_97_the_durable_token_is_never_written() -> None:
    """Every unsuccessful AUTO outcome collapses to the generic result."""
    damaged = _swap(
        _REPORT,
        "        RefusalResult = SIM_ATTEMPT_AUTO_NONCE_INDETERMINATE\n",
        "        RefusalResult = SIM_ATTEMPT_REFUSED\n", count=2)
    _control("test_44s", report=damaged)


def test_98_the_attempt_row_is_gated_on_verified_consumption() -> None:
    damaged = _swap(
        _REPORT,
        "        If package.HasSuppliedSeed Or package.AutoIdentityKnown Then\n",
        "        If package.HasSuppliedSeed Or package.NonceConsumed Then\n")
    _control("test_44h3", report=damaged)


def test_99_the_published_block_claims_an_unverified_nonce() -> None:
    damaged = _swap(
        _REPORT,
        "    If package.NonceConsumed Then\n        built(6, 1) = package.ConsumedNonce\n",
        "    If package.AutoIdentityKnown Then\n        built(6, 1) = package.ConsumedNonce\n")
    _control("test_44h3", report=damaged)


def test_100_fixed_mode_enters_the_auto_transaction() -> None:
    damaged = _swap(
        _NONCE,
        "        effectiveSeed = suppliedSeed\n",
        "        effectiveSeed = suppliedSeed\n"
        "        If Not ResolveNextNonce(autoNonce, state, detail) Then Exit Function\n")
    _control("test_44m", nonce=damaged)


def test_101_reconciliation_activates_on_the_attempt_row_again() -> None:
    """The REJECTED Option-3R carrier, restored.

    `Last Attempt Result` records the chronologically last attempt, so a FIXED
    run that has nothing to do with the pending AUTO transaction erases the
    lock. This is the shape the sidecar replaced.
    """
    damaged = _swap(
        _NONCE,
        "    If Not ReadPending(pending, hasPending, probe) Then\n",
        "    hasPending = (StrComp(SharedText(SIM_IDENTITY_ROW_LAST_ATTEMPT_RESULT), _\n"
        "                          SIM_ATTEMPT_AUTO_NONCE_INDETERMINATE, _\n"
        "                          vbBinaryCompare) = 0)\n"
        "    If Not ReadPending(pending, hasPending, probe) Then\n")
    _control("test_44h5", nonce=damaged)


def test_102_the_audit_writer_grows_a_blanket_suppressor() -> None:
    damaged = _swap(
        _REPORT,
        "    SimSheet.Range(AttemptRange()).Value2 = block\n",
        "    On Error Resume Next\n"
        "    SimSheet.Range(AttemptRange()).Value2 = block\n")
    _control("test_44g", report=damaged)


# ===========================================================================
# THE FOUR REJECTED SHAPES, RESTORED
#
# Each of these is the source as it actually stood when the static review
# rejected it. A control that cannot refuse the exact defect it was written for
# is not a control, so every one of them is required to fail.
# ===========================================================================
def test_103_the_first_verification_read_escapes_its_envelope() -> None:
    """BLOCKER 1, restored: `On Error GoTo 0` before the verification read.

    The counter write returns, the handler is disarmed, and then the read
    RAISES. With nothing armed in PersistAdvance the raise leaves through the
    outer allocation handler and no bounded reconciliation is ever taken - for
    a failure the contract names as `verification_read_raised`.
    """
    damaged = _swap(
        _NONCE,
        "    SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 = nonce + 1\n"
        "    If Not ReadPersistedNonce(stored, probe) Then\n",
        "    SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 = nonce + 1\n"
        "    On Error GoTo 0\n"
        "    If Not ReadPersistedNonce(stored, probe) Then\n")
    _control("test_44h", nonce=damaged)


def test_104_a_helper_false_verification_skips_the_reconciliation() -> None:
    """The other authorised cause, cut off from the same one observation."""
    damaged = _swap(
        _NONCE,
        "        PersistAdvance = Reconcile(nonce, \"the advance could not be verified (\" & _\n"
        "                                   probe & \")\", state, detail)\n",
        "        detail = \"simulation: the advance could not be verified\"\n")
    _control("test_44h", nonce=damaged)


def test_105_the_failure_arm_loses_the_effective_seed() -> None:
    """BLOCKER 2, restored: the seed copied on the success arm alone.

    Every refusal after the seed was derived then writes the attempted nonce
    beside a default `0` seed - an attempt row naming an identity nobody could
    reconstruct, in all three states the contract requires it in.
    """
    damaged = _swap(
        _REPORT,
        "    package.EffectiveSeed = seed\n"
        "    package.AutoIdentityKnown = identityKnown\n",
        "    If allocated Then package.EffectiveSeed = seed\n"
        "    package.AutoIdentityKnown = identityKnown\n")
    _control("test_44l", report=damaged)


def test_106_the_effective_seed_is_substituted_with_zero() -> None:
    damaged = _swap(_REPORT, "    package.EffectiveSeed = seed\n",
                    "    package.EffectiveSeed = 0\n")
    _control("test_44l", report=damaged)


def test_107_the_effective_seed_is_copied_only_when_consumed() -> None:
    damaged = _swap(
        _REPORT,
        "    package.EffectiveSeed = seed\n"
        "    package.AutoIdentityKnown = identityKnown\n",
        "    package.AutoIdentityKnown = identityKnown\n")
    damaged = _swap(
        damaged,
        "                                     vbBinaryCompare) = 0)\n",
        "                                     vbBinaryCompare) = 0)\n"
        "    If package.NonceConsumed Then package.EffectiveSeed = seed\n")
    _control("test_44l", report=damaged)


def test_108_the_nonce_is_retained_while_the_seed_is_dropped() -> None:
    damaged = _swap(_REPORT, "    package.EffectiveSeed = seed\n", "")
    _control("test_44l", report=damaged)


def test_109_the_failpoint_moves_below_the_fixed_branch() -> None:
    """BLOCKER 3, restored: FIXED falls through to `Phase6AfterNoncePersisted`.

    FIXED reads no counter, writes no counter and persists no advance, so the
    boundary the failpoint names does not exist on that path. Text order alone
    never saw this: the call really was below the branch - the branch simply
    did not leave.
    """
    damaged = _swap(
        _NONCE,
        "        On Error GoTo 0\n        SimNonceAllocate = True\n        Exit Function\n"
        "    End If\n",
        "    End If\n")
    _control("test_44m", nonce=damaged)


def test_110_the_recovery_state_is_recorded_as_a_generic_refusal() -> None:
    """BLOCKER 4's consequence, restored.

    `RECOVERY_REQUIRED` collapses to `REFUSED`, which says the run DECLINED to
    spend the nonce - the one claim this source cannot make when the counter
    reads a value that is neither the attempted nonce nor its advance.
    """
    damaged = _swap(
        _REPORT,
        "    ElseIf StrComp(package.NonceState, modSimNonce.SIM_NONCE_STATE_RECOVERY, _\n"
        "                   vbBinaryCompare) = 0 Then\n"
        "        RefusalResult = SIM_ATTEMPT_AUTO_NONCE_INDETERMINATE\n",
        "")
    _control("test_44s", report=damaged)


# ===========================================================================
# THE WRITE-AHEAD MARKER
# ===========================================================================
def test_111_the_counter_is_written_before_the_marker_is_laid() -> None:
    """The whole point of write-ahead, inverted."""
    damaged = _swap(
        _NONCE,
        "    If Not EstablishPending(nonce, detail) Then\n",
        "    If Not PersistAdvance(nonce, state, detail) Then Exit Function\n"
        "    If Not EstablishPending(nonce, detail) Then\n")
    damaged = _swap(
        damaged,
        "    If Not PersistAdvance(nonce, state, detail) Then Exit Function\n\n"
        "    ' A definite resolution clears the marker.",
        "\n    ' A definite resolution clears the marker.")
    _control("test_44n", nonce=damaged)


def test_112_a_failed_marker_write_falls_through_to_the_counter() -> None:
    """`counter_persist_forbidden_until_established`, ignored."""
    damaged = _swap(
        _NONCE,
        "        state = SIM_NONCE_STATE_PRE_ALLOCATION\n        Exit Function\n    End If\n"
        "    If Not PersistAdvance(nonce, state, detail) Then Exit Function\n",
        "        state = SIM_NONCE_STATE_PRE_ALLOCATION\n    End If\n"
        "    If Not PersistAdvance(nonce, state, detail) Then Exit Function\n")
    _control("test_44n", nonce=damaged)


def test_113_the_marker_is_written_but_never_verified() -> None:
    """A COM write that returns is not a write that landed."""
    damaged = _swap(
        _NONCE,
        "    If Not ReadPending(stored, present, probe) Then\n"
        "        detail = \"simulation: the pending AUTO nonce marker could not be verified (\" & _\n"
        "                 probe & \"). The counter was NOT touched and no nonce was consumed.\"\n"
        "        Exit Function\n"
        "    End If\n",
        "")
    damaged = _swap(
        damaged,
        "    If (Not present) Or stored <> nonce Then\n"
        "        detail = \"simulation: the pending AUTO nonce marker did not persist. The \" & _\n"
        "                 \"counter was NOT touched and no nonce was consumed.\"\n"
        "        Exit Function\n"
        "    End If\n",
        "")
    _control("test_44n", nonce=damaged)


def test_114_an_impossible_counter_reading_clears_the_marker() -> None:
    """`retained_on: counter_is_neither`, ignored - so the block lifts itself."""
    damaged = _swap(
        _NONCE,
        "    ' THE MARKER IS RETAINED. Future AUTO runs stay blocked until reconciled.\n"
        "    state = SIM_NONCE_STATE_RECOVERY\n",
        "    If Not ClearPending(detail) Then Exit Function\n"
        "    state = SIM_NONCE_STATE_RECOVERY\n")
    _control("test_44p", nonce=damaged)


def test_115_an_unreadable_counter_clears_the_marker() -> None:
    """`retained_on: counter_unreadable`, ignored."""
    damaged = _swap(
        _NONCE,
        "        state = SIM_NONCE_STATE_RECOVERY\n        Exit Function\n    End If\n\n"
        "    If hasPending Then\n",
        "        If Not ClearPending(probe) Then probe = probe\n"
        "        state = SIM_NONCE_STATE_RECOVERY\n        Exit Function\n    End If\n\n"
        "    If hasPending Then\n")
    _control("test_44p", nonce=damaged)


def test_116_the_indeterminate_case_clears_the_marker() -> None:
    """The state whose entire purpose is to leave the marker standing."""
    damaged = _swap(
        _NONCE,
        "    state = SIM_NONCE_STATE_INDETERMINATE\n",
        "    If Not ClearPending(detail) Then detail = detail\n"
        "    state = SIM_NONCE_STATE_INDETERMINATE\n")
    _control("test_44p", nonce=damaged)


def test_117_a_definite_resolution_leaves_the_marker_standing() -> None:
    """The opposite error: a resolved transaction blocks every later run."""
    damaged = _swap(
        _NONCE,
        "        state = SIM_NONCE_STATE_PRE_ALLOCATION\n"
        "        If Not ClearPending(detail) Then Exit Function\n",
        "        state = SIM_NONCE_STATE_PRE_ALLOCATION\n")
    _control("test_44p", nonce=damaged)


def test_118_the_clear_is_assumed_rather_than_verified() -> None:
    """`is_a_real_com_write: true`. An assumed clear is not a clear."""
    damaged = _swap(
        _NONCE,
        "    PendingCell.ClearContents\n    On Error GoTo 0\n",
        "    PendingCell.ClearContents\n    On Error GoTo 0\n"
        "    ClearPending = True\n    Exit Function\n")
    _control("test_44q", nonce=damaged)


def test_119_an_unresolved_cleanup_lets_the_run_sample() -> None:
    """`unresolved_cleanup_permits_sampling: false`, ignored."""
    damaged = _swap(
        _NONCE,
        "    If Not ClearPending(detail) Then\n"
        "        state = SIM_NONCE_STATE_RECOVERY\n"
        "        Exit Function\n"
        "    End If\n"
        "    RunAllocationTransaction = True\n",
        "    If Not ClearPending(detail) Then state = SIM_NONCE_STATE_RECOVERY\n"
        "    RunAllocationTransaction = True\n")
    _control("test_44q", nonce=damaged)


def test_120_a_failed_clear_rolls_the_counter_back() -> None:
    """Tidying up by moving the counter backwards is how a nonce gets replayed."""
    damaged = _swap(
        _NONCE,
        "    If present Then\n",
        "    If present Then\n"
        "        SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 = stored - 1\n")
    _control("test_44j", nonce=damaged)


def test_121_the_attempt_writer_can_reach_the_recovery_marker() -> None:
    """If the audit block could clear it, a FIXED attempt would erase the lock."""
    damaged = _swap(
        _REPORT,
        "    SimSheet.Range(AttemptRange()).Value2 = block\n",
        "    SimSheet.Range(SIM_PENDING_AUTO_NONCE_CELL).ClearContents\n"
        "    SimSheet.Range(AttemptRange()).Value2 = block\n")
    _control("test_44o", report=damaged)


def test_122_the_nonce_module_writes_the_attempt_row_itself() -> None:
    """Then the residual would be real again: safety back on audit storage."""
    damaged = _swap(
        _NONCE,
        "    state = SIM_NONCE_STATE_INDETERMINATE\n",
        "    modSimReport.WriteAttemptBlock\n"
        "    state = SIM_NONCE_STATE_INDETERMINATE\n")
    _control("test_44r", nonce=damaged)


def test_123_the_sidecar_coordinate_is_spelled_out_a_second_time() -> None:
    """Two independent literals are two authorities that agree only by luck."""
    damaged = _swap(
        _NONCE,
        "    Set PendingCell = modWorkbook.Sh(SIM_DATA_SHEET).Range(SIM_PENDING_AUTO_NONCE_CELL)\n",
        "    Set PendingCell = modWorkbook.Sh(SIM_DATA_SHEET).Range(\"F21\")\n")
    _control("test_44t", nonce=damaged)


def test_124_the_obsolete_allocation_language_is_left_standing() -> None:
    """A module that documents a rejected design still teaches it."""
    damaged = _swap(
        _REPORT,
        "        ' DIAGNOSTIC IDENTITY, not a consumption claim.",
        "        ' ALLOCATED, not CONSUMED. A verification failure leaves\n"
        "        ' allocation certain.\n"
        "        ' DIAGNOSTIC IDENTITY, not a consumption claim.")
    _control("test_44h3", report=damaged)


def test_125_a_raised_counter_read_escapes_into_the_entry_handler() -> None:
    """Then a standing marker gets recorded as an ordinary refusal.

    Selection reads the counter with no handler of its own armed. If the read
    is not decided where it happens, a raise leaves the classification at
    `NOT_APPLICABLE`, and the attempt row records a plain `REFUSED` - saying
    the run DECLINED to spend a nonce - while the sidecar still holds one.
    """
    damaged = _swap(
        _NONCE,
        "    On Error GoTo SharedReadRaised\n"
        "    raw = SharedCell(row).Value2\n"
        "    On Error GoTo 0\n",
        "    raw = SharedCell(row).Value2\n")
    _control("test_44g", nonce=damaged)


def test_126_a_raised_sidecar_read_escapes_into_the_entry_handler() -> None:
    """The same hole on the marker itself, which every caller depends on."""
    damaged = _swap(
        _NONCE,
        "    On Error GoTo ReadRaised\n"
        "    raw = PendingCell.Value2\n"
        "    blank = modWorkbook.IsEmptyCell(PendingCell)\n"
        "    On Error GoTo 0\n",
        "    raw = PendingCell.Value2\n"
        "    blank = modWorkbook.IsEmptyCell(PendingCell)\n")
    _control("test_44g", nonce=damaged)

