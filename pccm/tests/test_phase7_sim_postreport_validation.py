#!/usr/bin/env python3
"""PCCM Phase 7 Step-4 mutation controls for the sensitivity orchestration.

Each mutation is a plausible WRONG pipeline, applied to the real source, and
each must be refused by a named control. They protect executable and ownership
properties - what runs, in what order, over how many drivers, and what may be
written when. They do not police wording.

Runs standalone or under pytest.
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import test_phase7_sim_postreport as conformance  # noqa: E402

POST_BAS = PCCM_ROOT / "src" / "vba" / "modSimPostReport.bas"
_SOURCE = POST_BAS.read_text(encoding="utf-8")


def _swap(text: str, old: str, new: str, count: int = 1) -> str:
    assert text.count(old) == count, (old[:70], text.count(old))
    return text.replace(old, new, count)


def _refused(damaged: str, controls, reason: str) -> None:
    assert damaged != _SOURCE, f"{reason}: the mutation changed nothing"
    with tempfile.TemporaryDirectory(prefix="pccm-post-") as name:
        target = Path(name) / POST_BAS.name
        target.write_text(damaged, encoding="utf-8")
        saved = conformance.POST_BAS
        conformance.POST_BAS = target
        try:
            failures = []
            for control in controls:
                try:
                    control()
                except Exception as error:  # noqa: BLE001
                    failures.append(f"{control.__name__}: {error}")
        finally:
            conformance.POST_BAS = saved
    assert failures, f"{reason}: the mutation survived every named control"


# ===========================================================================
# A. THE PRECONDITION
# ===========================================================================
def test_01_a_stale_run_is_analysed_anyway() -> None:
    """The defect the precondition exists for: replaying the CURRENT model
    against a total published by a DIFFERENT one produces a table that is wrong
    in a way nothing on the sheet could reveal."""
    damaged = _swap(
        _SOURCE,
        "    If StrComp(status, SIM_STATE_CURRENT, vbBinaryCompare) <> 0 Then",
        "    If False Then")
    _refused(damaged, (conformance.test_15_only_a_current_run_may_be_analysed,),
             "a stale or invalid run was analysed")


def test_02_the_status_is_derived_here_instead_of_read() -> None:
    """A second derivation of run state is a second answer to the only question
    that matters before replay."""
    damaged = _swap(_SOURCE, "    status = modSimReport.PCCM_SimulationStatus()",
                    "    status = SharedText(SIM_IDENTITY_ROW_STATUS)")
    _refused(damaged, (conformance.test_11_each_step_is_delegated_to_its_accepted_owner,
                       conformance.test_15_only_a_current_run_may_be_analysed),
             "the simulation status was re-derived")


def test_03_the_absence_of_any_successful_run_is_accepted() -> None:
    damaged = _swap(_SOURCE, "    If Len(status) = 0 Then\n", "    If False Then\n")
    _refused(damaged, (conformance.test_15_only_a_current_run_may_be_analysed,),
             "sensitivity ran with no successful simulation to analyse")


# ===========================================================================
# B. THE PIPELINE
# ===========================================================================
def test_04_the_total_is_ranked_once_per_driver() -> None:
    """D sorts of one vector for one answer, and the contract says once."""
    damaged = _swap(
        _SOURCE,
        "        If Not modSimSensitivity.SimSensitivitySpearman( _\n"
        "                contributions, totalRanks, run.Iterations, rho, status, detail) Then",
        "        If Not modSimSensitivity.SimSensitivityMidRanks(totals, run.Iterations, _\n"
        "                totalRanks, detail) Then Exit Function\n"
        "        If Not modSimSensitivity.SimSensitivitySpearman( _\n"
        "                contributions, totalRanks, run.Iterations, rho, status, detail) Then")
    _refused(damaged, (conformance.test_06_the_total_is_ranked_once_and_reused_for_every_driver,),
             "the total was re-ranked inside the per-driver loop")


def test_05_only_the_first_driver_is_analysed() -> None:
    damaged = _swap(_SOURCE, "    For index = 0 To driverCount - 1\n"
                             "        If Not modSimEngine.SimEngineReplayDriver( _",
                    "    For index = 0 To 0\n"
                    "        If Not modSimEngine.SimEngineReplayDriver( _")
    _refused(damaged, (conformance.test_07_every_driver_is_processed_exactly_once,),
             "only the first driver was analysed")


def test_06_the_driver_is_chosen_by_supply_position() -> None:
    damaged = _swap(_SOURCE, "drivers(LBound(drivers) + index).PermanentId, contributions, detail",
                    "drivers(LBound(drivers)).PermanentId, contributions, detail")
    _refused(damaged, (conformance.test_07_every_driver_is_processed_exactly_once,),
             "every driver replayed the same permanent id")


def test_07_the_orchestrator_runs_its_own_simulation() -> None:
    """A second Monte Carlo would produce a different stochastic result and
    present it as the explanation of the published one."""
    damaged = _swap(
        _SOURCE, "    If Not ReadTotals(run, totals, detail) Then",
        "    If Not modSimEngine.SimEngineRun(drivers, driverCount, run.EffectiveSeed, _\n"
        "            run.Iterations, totals, totals, detail) Then Exit Function\n"
        "    If Not ReadTotals(run, totals, detail) Then")
    _refused(damaged, (conformance.test_12_it_starts_no_simulation_and_consumes_no_run_identity,),
             "the orchestrator ran its own simulation")


def test_08_the_mathematics_is_reimplemented_locally() -> None:
    damaged = _swap(_SOURCE, "        results(index).AbsRho = Abs(rho)",
                    "        results(index).AbsRho = Sqr(rho * rho)\n"
                    "        If SafeProduct(totals, 1, rho) Then rho = rho")
    _refused(damaged, (conformance.test_10_no_mathematics_is_reimplemented_here,
                       conformance.test_25_no_variance_share_or_squared_rho_is_produced),
             "arithmetic was reimplemented in the orchestrator")


# ===========================================================================
# C. MEMORY
# ===========================================================================
def test_09_every_driver_contribution_vector_is_retained() -> None:
    """The 240 MB matrix, arriving by the back door."""
    damaged = _swap(_SOURCE, "    Dim contributions() As Double",
                    "    Dim contributions() As Double\n    Dim kept() As Double")
    damaged = _swap(damaged, "    ReDim results(0 To driverCount - 1)",
                    "    ReDim results(0 To driverCount - 1)\n"
                    "    ReDim kept(0 To driverCount - 1, 0 To 1)")
    _refused(damaged, (conformance.test_08_one_contribution_vector_exists_at_a_time,),
             "a driver x iteration container was allocated")


# ===========================================================================
# D. PUBLICATION SAFETY
# ===========================================================================
def test_10_the_published_marker_is_written_first() -> None:
    """Then a failure part way through leaves a partial block wearing a current
    stamp - which is the whole defect the ordering prevents."""
    damaged = _swap(
        _SOURCE,
        "    StampCell(run.Bank, SIM_SENSITIVITY_STAMP_ROW_PUBLISHED).Value2 = vbNullString",
        "    StampCell(run.Bank, SIM_SENSITIVITY_STAMP_ROW_PUBLISHED).Value2 = "
        "SIM_SENSITIVITY_PUBLISHED")
    _refused(damaged, (conformance.test_17_the_published_marker_is_cleared_first_and_written_last,),
             "the published marker was set before the block was written")


def test_11_the_identity_is_stamped_after_the_marker() -> None:
    damaged = _swap(
        _SOURCE,
        "    StampCell(run.Bank, SIM_SENSITIVITY_STAMP_ROW_RUN_ID).Value2 = run.RunId\n",
        "")
    damaged = _swap(
        damaged,
        "    StampCell(run.Bank, SIM_SENSITIVITY_STAMP_ROW_PUBLISHED).Value2 = SIM_SENSITIVITY_PUBLISHED",
        "    StampCell(run.Bank, SIM_SENSITIVITY_STAMP_ROW_PUBLISHED).Value2 = SIM_SENSITIVITY_PUBLISHED\n"
        "    StampCell(run.Bank, SIM_SENSITIVITY_STAMP_ROW_RUN_ID).Value2 = run.RunId")
    _refused(damaged, (conformance.test_18_the_identity_is_written_before_the_marker,
                       conformance.test_17_the_published_marker_is_cleared_first_and_written_last),
             "the identity was stamped after the block was declared published")


def test_12_publication_happens_before_the_analysis_finishes() -> None:
    damaged = _swap(
        _SOURCE,
        "    If Not AnalyseDrivers(run, drivers, driverCount, totalRanks, results, detail) Then",
        "    If Not Publish(run, results, driverCount, order, eligibleCount, detail) Then\n"
        "        RunSensitivity = Refused(detail)\n        Exit Function\n    End If\n"
        "    If Not AnalyseDrivers(run, drivers, driverCount, totalRanks, results, detail) Then")
    _refused(damaged, (conformance.test_19_the_whole_result_is_built_before_anything_is_written,),
             "the block was published before the analysis had run")


def test_13_the_surplus_rows_of_a_larger_previous_result_survive() -> None:
    """20 drivers persisted, 7 now: overwriting the first 7 leaves 13 rows that
    are indistinguishable from the new result."""
    damaged = _swap(_SOURCE, "    ClearRecords run.Bank, first, last\n", "")
    _refused(damaged, (conformance.test_20_surplus_rows_from_a_larger_previous_result_are_cleared,),
             "the surplus rows of a longer previous result were left in place")


def test_14_only_the_new_rows_are_cleared() -> None:
    damaged = _swap(
        _SOURCE,
        "                     CStr(SIM_SENSITIVITY_FIRST_ROW + SIM_MAX_ITERATIONS - 1)).ClearContents",
        "                     CStr(SIM_SENSITIVITY_FIRST_ROW)).ClearContents")
    _refused(damaged, (conformance.test_20_surplus_rows_from_a_larger_previous_result_are_cleared,),
             "the clear stopped short of the surplus")


def test_15_the_record_count_stops_bounding_the_result() -> None:
    damaged = _swap(
        _SOURCE,
        "    StampCell(run.Bank, SIM_SENSITIVITY_STAMP_ROW_RECORD_COUNT).Value2 = driverCount",
        "    StampCell(run.Bank, SIM_SENSITIVITY_STAMP_ROW_RECORD_COUNT).Value2 = 1")
    _refused(damaged, (conformance.test_21_the_record_count_bounds_what_is_authoritative,),
             "the persisted count stopped describing the block")


def test_16_a_missing_record_is_not_noticed() -> None:
    damaged = _swap(_SOURCE, "    If slot <> driverCount Then", "    If False Then")
    _refused(damaged, (conformance.test_21_the_record_count_bounds_what_is_authoritative,),
             "a driver could vanish between analysis and publication")


# ===========================================================================
# E. THE RECORD AND THE UNDEFINED CASE
# ===========================================================================
def test_17_a_zero_variance_driver_is_published_as_rho_zero() -> None:
    damaged = _swap(
        _SOURCE,
        "        block(row, SIM_SENSITIVITY_OFFSET_RHO + 1) = vbNullString\n"
        "        block(row, SIM_SENSITIVITY_OFFSET_ABS_RHO + 1) = vbNullString",
        "        block(row, SIM_SENSITIVITY_OFFSET_RHO + 1) = 0#\n"
        "        block(row, SIM_SENSITIVITY_OFFSET_ABS_RHO + 1) = 0#")
    _refused(damaged, (conformance.test_23_a_zero_variance_record_carries_the_label_and_no_rho,),
             "an undefined correlation was published as rho = 0")


def test_18_the_no_variance_label_is_changed() -> None:
    damaged = _swap(_SOURCE, 'SENSITIVITY_NO_VARIANCE_LABEL As String = "n/a - no variance"',
                    'SENSITIVITY_NO_VARIANCE_LABEL As String = "0"')
    _refused(damaged, (conformance.test_23_a_zero_variance_record_carries_the_label_and_no_rho,),
             "the diagnostic label stopped saying no variance")


def test_19_a_contracted_field_is_dropped_from_the_record() -> None:
    damaged = _swap(
        _SOURCE,
        "    block(row, SIM_SENSITIVITY_OFFSET_DIRECTION + 1) = DirectionOf(record.Rho)\n", "")
    _refused(damaged, (conformance.test_22_every_contracted_field_is_written_and_no_other,),
             "a contracted field stopped being written")


def test_20_the_unranked_records_are_written_first() -> None:
    damaged = _swap(
        _SOURCE,
        "    slot = 0\n    For position = 0 To eligibleCount - 1",
        "    slot = 0\n    For index = 0 To driverCount - 1\n"
        "        If results(index).Status <> SIM_SENSITIVITY_DEFINED Then\n"
        "            FillRecord block, slot + 1, results(index), 0\n"
        "            slot = slot + 1\n        End If\n    Next index\n"
        "    For position = 0 To eligibleCount - 1")
    _refused(damaged, (conformance.test_24_the_ranked_records_come_first_and_carry_their_rank,),
             "the unranked records were written above the ranked ones")


def test_21_a_top_n_truncation_is_introduced() -> None:
    damaged = _swap(_SOURCE, "    For position = 0 To eligibleCount - 1",
                    "    If eligibleCount > 10 Then eligibleCount = 10\n"
                    "    For position = 0 To eligibleCount - 1")
    _refused(damaged, (conformance.test_26_no_top_n_truncation_and_no_subsampling,),
             "the ranked population was truncated")


# ===========================================================================
# F. THE SIMULATION IS NOT TOUCHED
# ===========================================================================
def test_22_sensitivity_runs_inside_the_simulation() -> None:
    """A successful run must stay successful even if the analysis of it fails."""
    report = (PCCM_ROOT / "src" / "vba" / "modSimReport.bas").read_text(encoding="utf-8")
    damaged_report = report.replace(
        "    result = RunSimulation(committed)",
        "    result = RunSimulation(committed)\n    modSimPostReport.PCCM_RunSensitivity", 1)
    assert damaged_report != report
    saved = (PCCM_ROOT / "src" / "vba" / "modSimReport.bas")
    original = saved.read_text(encoding="utf-8")
    try:
        saved.write_text(damaged_report, encoding="utf-8")
        refused = False
        try:
            conformance.test_13_the_simulation_does_not_run_sensitivity_for_you()
        except AssertionError:
            refused = True
    finally:
        saved.write_text(original, encoding="utf-8")
    assert refused, "sensitivity was wired into the simulation run unnoticed"


def test_23_publication_writes_an_iteration_row() -> None:
    damaged = _swap(
        _SOURCE,
        "    StampCell(run.Bank, SIM_SENSITIVITY_STAMP_ROW_ITERATIONS).Value2 = run.Iterations",
        "    SimSheet().Range(SIM_ITER_A_TOTAL_NOMINAL_COLUMN & \"34\").Value2 = 0#\n"
        "    StampCell(run.Bank, SIM_SENSITIVITY_STAMP_ROW_ITERATIONS).Value2 = run.Iterations")
    _refused(damaged, (conformance.test_14_it_writes_no_iteration_row_and_no_result_digest,),
             "publication wrote into the iteration records")


def test_24_a_refusal_clears_the_previous_result() -> None:
    """A refused attempt must not destroy sensitivity that belongs to an earlier
    successful run."""
    damaged = _swap(
        _SOURCE,
        "Private Function Refused(ByVal detail As String) As OperationResult\n"
        "    Refused.Ok = False",
        "Private Function Refused(ByVal detail As String) As OperationResult\n"
        "    SimSheet().Range(\"J34:Q100\").ClearContents\n"
        "    Refused.Ok = False")
    _refused(damaged, (conformance.test_16_a_refusal_writes_nothing_at_all,),
             "a refusal destroyed the previous result")
