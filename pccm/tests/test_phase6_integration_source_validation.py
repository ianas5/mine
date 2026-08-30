#!/usr/bin/env python3
"""PCCM Phase 6 Step-12 MUTATION CONTROLS for the integration source review.

A conformance test that cannot fail proves nothing. Every control damages one of
the authorities the integration review reads - a VBA module, the structure
contract, the emitted manifest, the Gate-B harness or the workbook manifest -
reruns the WHOLE Step-12 integration battery against the damaged copy, and
requires a NAMED detector among the refusers.

Nothing here writes to the repository: damaged copies live in a temporary
directory that is deleted on the way out, including on the exception path.

Runs standalone or under pytest.
"""

from __future__ import annotations

import re
import shutil
import sys
import tempfile
from contextlib import contextmanager
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import test_phase6_integration_source as conformance  # noqa: E402

_SRC = conformance.SRC_VBA
_SPEC = conformance.SPEC
_BUILD = conformance.BUILD


def _conformance_tests() -> list[str]:
    names = sorted(n for n in dir(conformance) if n.startswith("test_"))
    assert len(names) >= 25, names
    return names


def _run_battery() -> list[str]:
    refused = []
    for name in _conformance_tests():
        try:
            getattr(conformance, name)()
        except BaseException:  # noqa: BLE001 - any refusal counts
            refused.append(name)
    return refused


@contextmanager
def _installed(vba: dict[str, str] | None = None,
               spec: dict[str, str] | None = None,
               build: dict[str, str] | None = None):
    """Point the conformance module at damaged copies of whole trees.

    Whole trees, not single files: the review walks `src/vba` and reads several
    spec documents, so a damaged copy of one file has to sit beside undamaged
    copies of its siblings or the failure would be "the file vanished".
    """
    saved = (conformance.SRC_VBA, conformance.SPEC, conformance.BUILD,
             dict(conformance._CACHE))
    with tempfile.TemporaryDirectory(prefix="pccm-step12-mutation-") as name:
        temp = Path(name)
        conformance._CACHE.clear()
        try:
            for damaged, source, attribute in ((vba, saved[0], "SRC_VBA"),
                                               (spec, saved[1], "SPEC"),
                                               (build, saved[2], "BUILD")):
                if damaged is None:
                    continue
                target = temp / attribute.lower()
                shutil.copytree(source, target)
                for relative, text in damaged.items():
                    path = target / relative
                    assert path.is_file(), relative
                    assert path.read_text(encoding="utf-8") != text, (
                        f"the mutation changed nothing in {relative}"
                    )
                    path.write_text(text, encoding="utf-8")
                setattr(conformance, attribute, target)
            yield
        finally:
            conformance.SRC_VBA = saved[0]
            conformance.SPEC = saved[1]
            conformance.BUILD = saved[2]
            conformance._CACHE.clear()
            conformance._CACHE.update(saved[3])


def _control(expected: str, **damage) -> None:
    with _installed(**damage):
        refused = _run_battery()
    assert refused, "the mutation survived the whole integration battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def _read(root: Path, relative: str) -> str:
    return (root / relative).read_text(encoding="utf-8")


def _swap(text: str, old: str, new: str, count: int = 1) -> str:
    assert text.count(old) == count, (old[:80], text.count(old))
    return text.replace(old, new)


REPORT_BAS = "modSimReport.bas"
CALC_BAS = "modCalcReport.bas"
_REPORT = _read(_SRC, REPORT_BAS)
_CALC = _read(_SRC, CALC_BAS)
_STRUCTURE = _read(_SPEC, "structure_contract.yaml")
NONCE_BAS = "modSimNonce.bas"
_NONCE = _read(_SRC, NONCE_BAS)


def test_00_the_accepted_tree_passes_every_detector() -> None:
    with _installed():
        refused = _run_battery()
    assert refused == [], refused


# ===========================================================================
# 1. The request prefix
# ===========================================================================
def test_01_the_run_uses_a_stored_fingerprint_instead_of_the_bridge_output() -> None:
    damaged = _swap(
        _REPORT,
        "            package.AnalyticalFingerprint, package.Iterations, package.SeedMode, _\n"
        "            package.HasSuppliedSeed, package.SuppliedSeed, _\n"
        "            package.RequestFingerprint, detail) Then\n",
        "            ActiveSnapshotText(SIM_IDENTITY_ROW_REQUEST_FINGERPRINT), _\n"
        "            package.Iterations, package.SeedMode, _\n"
        "            package.HasSuppliedSeed, package.SuppliedSeed, _\n"
        "            package.RequestFingerprint, detail) Then\n")
    _control("test_07", vba={REPORT_BAS: damaged})


def test_02_the_request_prefix_becomes_the_effective_seed() -> None:
    damaged = _swap(
        _REPORT,
        "            package.AnalyticalFingerprint, package.Iterations, package.SeedMode, _\n"
        "            package.HasSuppliedSeed, package.SuppliedSeed, _\n"
        "            package.RequestFingerprint, detail) Then\n",
        "            CStr(package.EffectiveSeed), package.Iterations, package.SeedMode, _\n"
        "            package.HasSuppliedSeed, package.SuppliedSeed, _\n"
        "            package.RequestFingerprint, detail) Then\n")
    _control("test_07", vba={REPORT_BAS: damaged})


def test_03_the_analytical_fingerprint_is_reassigned_after_the_bridge() -> None:
    damaged = _swap(
        _REPORT,
        "    ' 3. Machine prerequisites, all of them, BEFORE anything is allocated.\n",
        "    package.AnalyticalFingerprint = ActiveSnapshotText( _\n"
        "        SIM_IDENTITY_ROW_REQUEST_FINGERPRINT)\n"
        "    ' 3. Machine prerequisites, all of them, BEFORE anything is allocated.\n")
    _control("test_07", vba={REPORT_BAS: damaged})


# ===========================================================================
# 2. Retained-array identity
# ===========================================================================
def test_04_the_digest_runs_over_a_rebuilt_array() -> None:
    damaged = _swap(
        _REPORT,
        "    If Not modSimFingerprint.SimFpResultDigest(package.TotalNominal, package.TotalPv, _\n",
        "    ReDim package.TotalNominal(0 To package.Iterations - 1)\n"
        "    If Not modSimFingerprint.SimFpResultDigest(package.TotalNominal, package.TotalPv, _\n")
    _control("test_09", vba={REPORT_BAS: damaged})


def test_05_the_publication_writes_a_different_carrier() -> None:
    damaged = _swap(
        _REPORT,
        "            block(index + 1, 2) = _\n"
        "                package.TotalNominal(LBound(package.TotalNominal) + offset + index)\n",
        "            block(index + 1, 2) = package.NominalLadder(0)\n")
    _control("test_09", vba={REPORT_BAS: damaged})


def test_06_the_arrays_are_sorted_before_they_are_used() -> None:
    damaged = _swap(
        _REPORT,
        "    ' 6-7. The statistics, over the SAME retained arrays that will be published.\n",
        "    SortAscending package.TotalNominal, package.Iterations\n"
        "    ' 6-7. The statistics, over the SAME retained arrays that will be published.\n")
    _control("test_10", vba={REPORT_BAS: damaged})


# ===========================================================================
# 3. Quantile provenance
# ===========================================================================
def test_07_a_quantile_carrier_is_mutated_after_describe() -> None:
    damaged = _swap(
        _REPORT,
        "    If Not SameLadder(package, detail) Then Exit Function\n",
        "    If Not SameLadder(package, detail) Then Exit Function\n"
        "    package.PvLadder(LBound(package.PvLadder)) = package.BasePv\n")
    _control("test_11", vba={REPORT_BAS: damaged})


def test_08_a_contingency_is_computed_by_subtraction_here() -> None:
    body = conformance._procedure("modSimReport", "BuildContingencies")
    match = re.search(r"^\s*package\.NominalContingency\([^)]*\)\s*=\s*value\s*$",
                      body, re.M)
    assert match, body
    damaged = _swap(
        _REPORT, match.group(0) + "\n",
        match.group(0).replace(
            "= value",
            "= package.NominalLadder(index) - package.BaseNominal") + "\n")
    _control("test_06", vba={REPORT_BAS: damaged})


# ===========================================================================
# 4. The reporting boundary
# ===========================================================================
def test_09_selected_confidence_level_enters_the_run() -> None:
    damaged = _swap(
        _REPORT,
        "    ' 2. The two simulation controls, strictly.\n",
        "    If modWorkbook.IsEmptyCell(modWorkbook.NamedCell( _\n"
        '            "inpSelectedConfidenceLevel")) Then Exit Function\n'
        "    ' 2. The two simulation controls, strictly.\n")
    _control("test_15", vba={REPORT_BAS: damaged})


def test_10_results_is_written_by_vba() -> None:
    damaged = _swap(
        _REPORT,
        "Private Function SimSheet() As Worksheet\n",
        "Private Sub PublishToResults(ByVal text As String)\n"
        '    modWorkbook.Sh("Results").Range("B3").Value2 = text\n'
        "End Sub\n\n"
        "Private Function SimSheet() As Worksheet\n")
    _control("test_14", vba={REPORT_BAS: damaged})


def test_11_the_active_bank_is_touched_during_candidate_publication() -> None:
    damaged = _swap(
        _REPORT,
        "    If Not WriteIterationBank(package, detail) Then Exit Function\n",
        "    SharedCell(SIM_IDENTITY_ROW_ACTIVE_BANK).Value2 = package.TargetBank\n"
        "    If Not WriteIterationBank(package, detail) Then Exit Function\n")
    _control("test_12", vba={REPORT_BAS: damaged})


# ===========================================================================
# 5. The Phase-5 bridge
# ===========================================================================
def test_12_the_bridge_calls_the_calculation_endpoint() -> None:
    damaged = _swap(
        _CALC,
        "    If Not PrepareCurrentCalculation(package, detail) Then Exit Function\n",
        "    PCCM_Calculate\n"
        "    If Not PrepareCurrentCalculation(package, detail) Then Exit Function\n")
    _control("test_04", vba={CALC_BAS: damaged})


def test_13_the_bridge_drops_the_current_gate() -> None:
    gate = ("    status = DeriveStatus(package, True)\n"
            "    If StrComp(status, CALC_STATUS_CURRENT, vbBinaryCompare) <> 0 Then\n"
            '        detail = "the calculation is " & status & _\n'
            '                 "; the simulation needs a CURRENT calculation"\n'
            "        Exit Function\n"
            "    End If\n")
    damaged = _swap(_CALC, gate, "")
    _control("test_02", vba={CALC_BAS: damaged})


def test_14_the_bridge_rebuilds_instead_of_projecting() -> None:
    damaged = _swap(
        _CALC,
        "    analyticalFingerprint = package.Fingerprint\n",
        "    analyticalFingerprint = BuildFingerprint(package)\n")
    _control("test_03", vba={CALC_BAS: damaged})


def test_15_the_bridge_writes_to_the_workbook() -> None:
    damaged = _swap(
        _CALC,
        "    CalcPrepareSimulationInputs = True\n",
        "    modWorkbook.NamedCell(NM_CALC_STATE).Value2 = status\n"
        "    CalcPrepareSimulationInputs = True\n")
    _control("test_04", vba={CALC_BAS: damaged})


def test_16_a_second_phase6_bridge_appears() -> None:
    damaged = _swap(
        _CALC,
        "Public Function CalcPrepareSimulationInputs(",
        "Public Function CalcPrepareSimulationTotals(ByRef total As Double) As Boolean\n"
        "    total = 0\n"
        "    CalcPrepareSimulationTotals = True\n"
        "End Function\n\n"
        "Public Function CalcPrepareSimulationInputs(")
    _control("test_25", vba={CALC_BAS: damaged})


# ===========================================================================
# 6. D6-11, repo-wide
# ===========================================================================
def test_17_run_simulation_appears_in_a_second_module() -> None:
    engine = _read(_SRC, "modSimEngine.bas")
    damaged = _swap(
        engine, "Option Explicit\n",
        "Option Explicit\n\n"
        "Private Function RunSimulationHelper() As Boolean\n"
        "    RunSimulationHelper = False\n"
        "End Function\n")
    _control("test_19", vba={"modSimEngine.bas": damaged})


def test_18_the_endpoint_is_granted_to_a_second_owner() -> None:
    damaged = _swap(
        _STRUCTURE,
        '    - construct: "RunSimulation"\n'
        '      allowed_in:\n        - "modSimReport"\n',
        '    - construct: "RunSimulation"\n'
        '      allowed_in:\n        - "modSimReport"\n        - "modSimEngine"\n')
    _control("test_19", spec={"structure_contract.yaml": damaged})


def test_19_the_algorithm_token_appears_in_the_reporter() -> None:
    damaged = _swap(
        _REPORT, "Option Explicit\n",
        "Option Explicit\n\n"
        "Private Const SIM_NOTE As String = \"x\"\n"
        "Private Function MRG32k3aNote() As String\n"
        "    MRG32k3aNote = SIM_NOTE\n"
        "End Function\n")
    _control("test_19", vba={REPORT_BAS: damaged})


def test_20_executable_percentile_appears() -> None:
    stats = _read(_SRC, "modSimStats.bas")
    damaged = _swap(
        stats, "Option Explicit\n",
        "Option Explicit\n\n"
        "Private Function PercentileOf(ByVal x As Double) As Double\n"
        "    PercentileOf = x\n"
        "End Function\n")
    _control("test_20", vba={"modSimStats.bas": damaged})


def test_21_the_scoped_rule_is_flattened() -> None:
    damaged = _swap(
        _STRUCTURE,
        '    - construct: "RunSimulation"\n'
        '      allowed_in:\n        - "modSimReport"\n',
        '    - "RunSimulation"\n')
    _control("test_19", spec={"structure_contract.yaml": damaged})


def test_22_a_wildcard_owner_is_granted() -> None:
    damaged = _swap(
        _STRUCTURE,
        '    - construct: "MRG32k3a"\n'
        '      allowed_in:\n        - "modSimRng"\n',
        '    - construct: "MRG32k3a"\n'
        '      allowed_in:\n        - "modSimRng"\n        - "*"\n')
    _control("test_19", spec={"structure_contract.yaml": damaged})


def test_23_percentile_is_granted_an_owner() -> None:
    damaged = _swap(
        _STRUCTURE, '    - "Percentile"\n',
        '    - construct: "Percentile"\n'
        '      allowed_in:\n        - "modSimStats"\n')
    _control("test_20", spec={"structure_contract.yaml": damaged})


def test_24_a_global_prohibition_is_scoped() -> None:
    damaged = _swap(
        _STRUCTURE, '    - "Randomize"\n',
        '    - construct: "Randomize"\n'
        '      allowed_in:\n        - "modSimRng"\n')
    _control("test_21", spec={"structure_contract.yaml": damaged})


# ===========================================================================
# 7. The public surfaces
# ===========================================================================
def test_25_an_eighth_phase6_endpoint_appears() -> None:
    damaged = _swap(
        _REPORT,
        "Public Function PCCM_SimulationStatus() As String\n",
        "Public Function PCCM_SimulationRunId() As String\n"
        "    PCCM_SimulationRunId = ActiveSnapshotText(SIM_IDENTITY_ROW_RUN_ID)\n"
        "End Function\n\n"
        "Public Function PCCM_SimulationStatus() As String\n")
    _control("test_17", vba={REPORT_BAS: damaged})


def test_26_a_phase6_endpoint_is_renamed() -> None:
    damaged = _swap(
        _REPORT,
        "Public Function PCCM_SimulationResultDigest() As String\n",
        "Public Function PCCM_SimulationDigest() As String\n")
    damaged = damaged.replace("    PCCM_SimulationResultDigest = ",
                              "    PCCM_SimulationDigest = ")
    _control("test_17", vba={REPORT_BAS: damaged})


def test_27_a_phase6_endpoint_becomes_a_button() -> None:
    damaged = _swap(
        _STRUCTURE, '    - "PCCM_ApplyTimeline"\n',
        '    - "PCCM_ApplyTimeline"\n    - "PCCM_RunSimulation"\n')
    _control("test_17", spec={"structure_contract.yaml": damaged})


def test_28_the_bridge_becomes_an_automation_endpoint() -> None:
    damaged = _swap(
        _STRUCTURE, '    - "PCCM_CurrentInputFingerprint"\n',
        '    - "PCCM_CurrentInputFingerprint"\n    - "CalcPrepareSimulationInputs"\n')
    _control("test_04", spec={"structure_contract.yaml": damaged})


# ===========================================================================
# 8. The frozen source and the frozen authority
# ===========================================================================
def test_29_an_accepted_kernel_moves_by_one_byte() -> None:
    engine = _read(_SRC, "modSimEngine.bas")
    damaged = engine.replace("Option Explicit\n", "Option Explicit\n\n", 1)
    _control("test_23", vba={"modSimEngine.bas": damaged})


def test_30_the_reporter_moves_by_one_byte() -> None:
    damaged = _REPORT.replace("Option Explicit\n", "Option Explicit\n\n", 1)
    _control("test_23", vba={REPORT_BAS: damaged})


def test_31_the_accepted_reporter_prefix_moves() -> None:
    damaged = _swap(_CALC, "Attribute VB_Name = \"modCalcReport\"\n",
                    "Attribute VB_Name = \"modCalcReport\"\n' moved\n")
    _control("test_25", vba={CALC_BAS: damaged})


def test_32_a_second_step_11_banner_appears() -> None:
    banner = ("' ==========================================================================\n"
              "' STEP 11 ADDITION - THE PHASE-6 PREPARATION BRIDGE\n")
    damaged = _CALC + "\n" + banner
    _control("test_25", vba={CALC_BAS: damaged})


def test_33_the_generated_projection_moves() -> None:
    module = _read(_BUILD / "vba", "modSimContract.bas")
    _control("test_24", build={"vba/modSimContract.bas": module + "\n"})


def test_34_the_manifest_loses_a_structured_rule() -> None:
    import json

    manifest = json.loads(_read(_BUILD, "stage_b_manifest.json"))
    manifest["vba"]["forbidden_construct_rules"] = [
        rule for rule in manifest["vba"]["forbidden_construct_rules"]
        if rule["construct"] != "RunSimulation"]
    _control("test_22", build={"stage_b_manifest.json": json.dumps(manifest, indent=2)})


def test_35_the_manifest_and_the_contract_disagree_on_the_inventory() -> None:
    import json

    manifest = json.loads(_read(_BUILD, "stage_b_manifest.json"))
    manifest["vba"]["modules"] = manifest["vba"]["modules"][:-1]
    _control("test_22", build={"stage_b_manifest.json": json.dumps(manifest, indent=2)})


# ===========================================================================
# 9. Nothing from Step 13 has arrived
# ===========================================================================
def test_36_a_phase7_module_is_declared_early() -> None:
    damaged = _swap(
        _STRUCTURE,
        '    - name: "modSimReport"\n',
        '    - name: "modSimSensitivity"\n'
        '      generated: false\n'
        '      responsibility: "Driver ranking."\n'
        '    - name: "modSimReport"\n')
    _control("test_22", spec={"structure_contract.yaml": damaged})


def test_37_a_phase7_module_appears_on_disk() -> None:
    with tempfile.TemporaryDirectory(prefix="pccm-step12-phase7-") as name:
        target = Path(name) / "vba"
        shutil.copytree(_SRC, target)
        (target / "modSimDashboard.bas").write_text(
            'Attribute VB_Name = "modSimDashboard"\nOption Explicit\n',
            encoding="utf-8")
        saved = conformance.SRC_VBA
        conformance._CACHE.clear()
        try:
            conformance.SRC_VBA = target
            refused = _run_battery()
        finally:
            conformance.SRC_VBA = saved
            conformance._CACHE.clear()
    assert refused, "an undeclared Phase-7 module survived the whole battery"
    assert any(n.startswith("test_27") for n in refused), refused


if __name__ == "__main__":
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))


# ===========================================================================
# 10. The cross-module failure-path guarantee (Step-12 settlement)
# ===========================================================================
def test_38_a_candidate_failure_is_re_raised_instead_of_recorded() -> None:
    """The Step-11 shape: an infrastructure failure skips the attempt axis."""
    damaged = _swap(
        _REPORT,
        "    If Not PublishCandidate(package, detail) Then\n"
        "        RunSimulation = RecordFailure(package, detail)\n"
        "        Exit Function\n"
        "    End If\n",
        "    If Not PublishCandidate(package, detail) Then\n"
        "        Err.Raise vbObjectError + 1, , detail\n"
        "    End If\n")
    _control("test_28", vba={REPORT_BAS: damaged})


def test_39_a_commit_failure_is_re_raised_instead_of_recorded() -> None:
    damaged = _swap(
        _REPORT,
        "    If Not FinalCommit(package, detail) Then\n"
        "        RunSimulation = RecordFailure(package, detail)\n"
        "        Exit Function\n"
        "    End If\n",
        "    If Not FinalCommit(package, detail) Then\n"
        "        Err.Raise vbObjectError + 2, , detail\n"
        "    End If\n")
    _control("test_28", vba={REPORT_BAS: damaged})


def test_40_the_candidate_envelope_is_removed() -> None:
    damaged = _swap(_REPORT, "    On Error GoTo CandidateFailed\n", "")
    _control("test_28", vba={REPORT_BAS: damaged})


def test_41_the_commit_envelope_is_removed() -> None:
    damaged = _swap(_REPORT, "    On Error GoTo CommitFailed\n", "")
    _control("test_28", vba={REPORT_BAS: damaged})


def test_42_a_phase6_module_suppresses_errors_wholesale() -> None:
    damaged = _swap(
        _REPORT, "    On Error GoTo CandidateFailed\n",
        "    On Error Resume Next\n")
    _control("test_28", vba={REPORT_BAS: damaged})


def test_43_the_consumed_nonce_is_rolled_back_on_failure() -> None:
    damaged = _swap(
        _NONCE,
        "AllocationFailed:\n    failure = Err.Description\n",
        "AllocationFailed:\n"
        "    SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 = autoNonce - 1\n"
        "    failure = Err.Description\n")
    _control("test_29", vba={NONCE_BAS: damaged})


def test_44_the_failed_candidate_bank_is_erased() -> None:
    damaged = _swap(
        _REPORT,
        "CandidateFailed:\n    failure = Err.Description\n",
        "CandidateFailed:\n"
        "    SimSheet.Range(SnapshotRange(package.TargetBank)).ClearContents\n"
        "    failure = Err.Description\n")
    _control("test_29", vba={REPORT_BAS: damaged})


def test_45_the_prior_block_is_captured_after_the_write() -> None:
    damaged = _swap(
        _REPORT,
        "    On Error GoTo CaptureFailed\n"
        "    previous = SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2\n"
        "    On Error GoTo 0\n\n"
        "    BuildCommitBlock package, block\n\n"
        "    ' B. THE COMMIT. From the assignment onward every exit restores.\n"
        "    On Error GoTo CommitFailed\n"
        "    SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2 = block\n",
        "    BuildCommitBlock package, block\n\n"
        "    On Error GoTo CommitFailed\n"
        "    SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2 = block\n"
        "    previous = SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2\n")
    _control("test_29", vba={REPORT_BAS: damaged})


def test_46_the_restore_write_is_removed_from_the_commit() -> None:
    damaged = _swap(
        _REPORT,
        "    SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2 = previous\n"
        "    If SameBlock(SIM_FINAL_COMMIT_RANGE, previous, 9, 1) Then\n",
        "    If SameBlock(SIM_FINAL_COMMIT_RANGE, previous, 9, 1) Then\n")
    _control("test_29", vba={REPORT_BAS: damaged})


def test_47_the_after_nonce_failpoint_raises_out_of_run_simulation() -> None:
    """A naked raising call in the orchestrator bypasses the attempt axis."""
    nonce = _swap(_NONCE,
                  "    modAppState.FailPointCheck FAILPOINT_SIM_AFTER_NONCE\n", "")
    report = _swap(
        _REPORT,
        "    If Not AllocateAutoNonce(package, detail) Then\n",
        "    modAppState.FailPointCheck modSimNonce.FAILPOINT_SIM_AFTER_NONCE\n"
        "    If Not AllocateAutoNonce(package, detail) Then\n")
    _control("test_28", vba={REPORT_BAS: report, NONCE_BAS: nonce})


def test_48_a_failpoint_fires_with_no_handler_armed() -> None:
    damaged = _swap(
        _REPORT,
        "    On Error GoTo CandidateFailed\n\n"
        "    BuildSnapshotBlock package, snapshot\n",
        "    BuildSnapshotBlock package, snapshot\n")
    damaged = _swap(
        damaged,
        "    If Not VerifyCandidateBank(package, snapshot, summary, contingency, detail) Then\n",
        "    On Error GoTo CandidateFailed\n"
        "    If Not VerifyCandidateBank(package, snapshot, summary, contingency, detail) Then\n")
    _control("test_28", vba={REPORT_BAS: damaged})


def test_49_the_nonce_transaction_loses_its_envelope() -> None:
    damaged = _swap(_NONCE, "    On Error GoTo AllocationFailed\n", "")
    _control("test_28", vba={NONCE_BAS: damaged})


def test_50_the_injection_precedes_the_verified_persistence() -> None:
    damaged = _swap(_NONCE,
                    "    modAppState.FailPointCheck FAILPOINT_SIM_AFTER_NONCE\n", "")
    damaged = _swap(
        damaged,
        "    If Not RunAllocationTransaction(autoNonce, allocationState, _\n",
        "    modAppState.FailPointCheck FAILPOINT_SIM_AFTER_NONCE\n"
        "    If Not RunAllocationTransaction(autoNonce, allocationState, _\n")
    _control("test_29", vba={NONCE_BAS: damaged})


def test_51_the_seed_is_derived_after_the_advance_is_persisted() -> None:
    """The contract's order is read < derive < persist, and it is load-bearing."""
    damaged = _swap(
        _NONCE,
        "    If Not modSimRng.SimRngAutoSeedFromNonce(autoNonce, seed, detail) Then\n"
        "        Exit Function\n"
        "    End If\n"
        "    effectiveSeed = seed\n"
        "    identityKnown = True\n",
        "")
    damaged = _swap(
        damaged,
        "    If Not RunAllocationTransaction(autoNonce, allocationState, _\n"
        "                                    recoveryRequired, detail) Then\n"
        "        Exit Function\n"
        "    End If\n",
        "    If Not RunAllocationTransaction(autoNonce, allocationState, _\n"
        "                                    recoveryRequired, detail) Then\n"
        "        Exit Function\n"
        "    End If\n"
        "    If Not modSimRng.SimRngAutoSeedFromNonce(autoNonce, seed, detail) Then\n"
        "        Exit Function\n"
        "    End If\n"
        "    effectiveSeed = seed\n"
        "    identityKnown = True\n")
    _control("test_29", vba={NONCE_BAS: damaged})


def test_52_the_attempt_record_loses_the_allocated_identity() -> None:
    damaged = _swap(
        _REPORT,
        "        If package.HasSuppliedSeed Or package.AutoIdentityKnown Then\n",
        "        If package.HasSuppliedSeed Or package.NonceConsumed Then\n")
    _control("test_29", vba={REPORT_BAS: damaged})


def test_53_the_nonce_module_writes_the_attempt_row() -> None:
    """Attempt-row persistence has exactly one owner, and it is the reporter."""
    damaged = _swap(
        _NONCE,
        "Private Function SharedCell(ByVal row As Long) As Range\n",
        "Private Sub WriteAttemptBlock(ByVal result As String)\n"
        "    SharedCell(SIM_IDENTITY_ROW_LAST_ATTEMPT_RESULT).Value2 = result\n"
        "End Sub\n\n"
        "Private Function SharedCell(ByVal row As Long) As Range\n")
    _control("test_30", vba={NONCE_BAS: damaged})


def test_54_the_audit_writer_gains_a_blanket_suppressor() -> None:
    damaged = _swap(
        _REPORT,
        "    SimSheet.Range(AttemptRange()).Value2 = block\n",
        "    On Error Resume Next\n"
        "    SimSheet.Range(AttemptRange()).Value2 = block\n")
    _control("test_28", vba={REPORT_BAS: damaged})


def test_55_the_nonce_module_gains_an_endpoint() -> None:
    damaged = _swap(
        _NONCE, "Public Function SimNonceAllocate(",
        "Public Function PCCM_SimulationNonce() As String\n"
        "    PCCM_SimulationNonce = vbNullString\n"
        "End Function\n\n"
        "Public Function SimNonceAllocate(")
    _control("test_30", vba={NONCE_BAS: damaged})


def test_56_the_nonce_module_depends_back_on_the_reporter() -> None:
    damaged = _swap(
        _NONCE,
        'Public Const FAILPOINT_SIM_AFTER_NONCE As String = "Phase6AfterNoncePersisted"\n',
        "Public Const FAILPOINT_SIM_AFTER_NONCE As String = _\n"
        "    modSimReport.FAILPOINT_SIM_CANDIDATE_BANK\n")
    _control("test_30", vba={NONCE_BAS: damaged})


def test_57_the_run_package_becomes_public() -> None:
    damaged = _swap(_REPORT, "Private Type SimRunPackage\n", "Public Type SimRunPackage\n")
    _control("test_30", vba={REPORT_BAS: damaged})


def test_58_the_sidecar_is_declared_on_top_of_the_bank_b_snapshot() -> None:
    """A coordinate is free or it is not; a comment saying so proves nothing."""
    contract = _read(conformance.SPEC, "sim_contract.yaml")
    damaged = _swap(contract, '    cell: "F21"\n    column: "F"\n    row: 21\n',
                    '    cell: "F20"\n    column: "F"\n    row: 20\n')
    _control("test_32", spec={"sim_contract.yaml": damaged})


def test_59_the_sidecar_is_declared_inside_the_iteration_table() -> None:
    contract = _read(conformance.SPEC, "sim_contract.yaml")
    damaged = _swap(contract, '    cell: "F21"\n    column: "F"\n    row: 21\n',
                    '    cell: "F40"\n    column: "F"\n    row: 40\n')
    _control("test_32", spec={"sim_contract.yaml": damaged})


def test_60_the_sidecar_is_moved_into_the_shared_commit_column() -> None:
    """Column D is where the final commit writes; a sidecar there is not durable."""
    contract = _read(conformance.SPEC, "sim_contract.yaml")
    damaged = _swap(contract, '    cell: "F21"\n    column: "F"\n    row: 21\n',
                    '    cell: "D21"\n    column: "D"\n    row: 21\n')
    _control("test_32", spec={"sim_contract.yaml": damaged})


def test_61_the_generated_coordinate_is_emitted_twice() -> None:
    """Two generated constants for one cell is two authorities."""
    generated = _read(conformance.BUILD, "vba/modSimContract.bas")
    damaged = _swap(
        generated,
        'Public Const SIM_PENDING_AUTO_NONCE_CELL As String = "F21"\n',
        'Public Const SIM_PENDING_AUTO_NONCE_CELL As String = "F21"\n'
        'Public Const SIM_PENDING_AUTO_NONCE_CELL_2 As String = "F21"\n')
    _control("test_34", build={"vba/modSimContract.bas": damaged})


def test_62_the_iteration_table_is_shifted_to_make_room() -> None:
    """The sidecar was chosen precisely so no row had to move."""
    contract = _read(conformance.SPEC, "sim_contract.yaml")
    damaged = _swap(contract, "    header_row: 33\n", "    header_row: 34\n")
    _control("test_33", spec={"sim_contract.yaml": damaged})


def test_63_the_registry_still_describes_the_rejected_carrier() -> None:
    """A registry entry that names a withdrawn authority teaches it."""
    structure = _read(conformance.SPEC, "structure_contract.yaml")
    damaged = _swap(
        structure,
        "Its durable recovery inputs are the pending marker and the nonce "
        "counter, never the mutable last-attempt audit axis.",
        "It interprets a prior AUTO_NONCE_INDETERMINATE attempt.")
    _control("test_30", spec={"structure_contract.yaml": damaged})



# ===========================================================================
# THE RUN-2 COMPILE DEFECT
# ===========================================================================
# `Argument not optional`, raised by the real VBE on the retained Run-2
# workbook. Five Phase-6 call sites passed three arguments to a four-argument
# helper, and VBA's compile-on-demand meant nothing had ever required those
# bodies to compile. The controls that existed proved the call was PRESENT.
def test_64_read_pending_drops_the_mandatory_result_argument() -> None:
    """The exact call the compiler stopped on."""
    damaged = _swap(
        _NONCE,
        "    If Not modWorkbook.IsWholeInRange(raw, CDbl(SIM_NONCE_FIRST_VALID), _\n"
        "                                      CDbl(SIM_NONCE_LAST_VALID), number) Then\n",
        "    If Not modWorkbook.IsWholeInRange(raw, CDbl(SIM_NONCE_FIRST_VALID), _\n"
        "                                      CDbl(SIM_NONCE_LAST_VALID)) Then\n")
    _control("test_35", vba={NONCE_BAS: damaged})


def test_65_read_shared_drops_the_mandatory_result_argument() -> None:
    """The second malformed call in the same module."""
    damaged = _swap(
        _NONCE,
        "    If Not modWorkbook.IsWholeInRange(raw, minValue, maxValue, number) Then\n",
        "    If Not modWorkbook.IsWholeInRange(raw, minValue, maxValue) Then\n")
    _control("test_35", vba={NONCE_BAS: damaged})


def test_66_the_iteration_control_drops_the_result_argument() -> None:
    """One of the three the Windows report did not know about."""
    damaged = _swap(
        _REPORT,
        "    If Not modWorkbook.IsWholeInRange(raw, CDbl(SIM_MIN_ITERATIONS), _\n"
        "                                      CDbl(SIM_MAX_ITERATIONS), value) Then\n",
        "    If Not modWorkbook.IsWholeInRange(raw, CDbl(SIM_MIN_ITERATIONS), _\n"
        "                                      CDbl(SIM_MAX_ITERATIONS)) Then\n")
    _control("test_35", vba={REPORT_BAS: damaged})


def test_67_the_seed_control_drops_the_result_argument() -> None:
    damaged = _swap(
        _REPORT,
        "    If Not modWorkbook.IsWholeInRange(raw, CDbl(SIM_SEED_MIN), CDbl(SIM_SEED_MAX), _\n"
        "                                      value) Then\n",
        "    If Not modWorkbook.IsWholeInRange(raw, CDbl(SIM_SEED_MIN), CDbl(SIM_SEED_MAX)) Then\n")
    _control("test_35", vba={REPORT_BAS: damaged})


def test_68_the_machine_long_reader_drops_the_result_argument() -> None:
    damaged = _swap(
        _REPORT,
        "    If Not modWorkbook.IsWholeInRange(raw, minValue, maxValue, number) Then\n",
        "    If Not modWorkbook.IsWholeInRange(raw, minValue, maxValue) Then\n")
    _control("test_35", vba={REPORT_BAS: damaged})


def test_69_the_detector_is_general_not_helper_specific() -> None:
    """A different helper, a different module, the same class of defect.

    If this only refused IsWholeInRange it would have closed five call sites and
    left the class open one compiler stop away.
    """
    damaged = _swap(
        _REPORT,
        "    If Not modWorkbook.TryReadDouble(raw, value) Then\n"
        "        detail = \"simulation: Monte Carlo Iterations is not a usable number\"\n",
        "    If Not modWorkbook.TryReadDouble(raw) Then\n"
        "        detail = \"simulation: Monte Carlo Iterations is not a usable number\"\n")
    _control("test_35", vba={REPORT_BAS: damaged})


def test_70_an_extra_argument_is_refused_too() -> None:
    """Too many is the same compile error class as too few."""
    damaged = _swap(
        _NONCE,
        "    If Not modWorkbook.IsWholeInRange(raw, minValue, maxValue, number) Then\n",
        "    If Not modWorkbook.IsWholeInRange(raw, minValue, maxValue, number, 0) Then\n")
    _control("test_35", vba={NONCE_BAS: damaged})


def test_71_the_helper_signature_is_widened_to_hide_the_defect() -> None:
    """Making the out-parameter Optional would make the defective calls legal
    again while changing what the helper means."""
    damaged = _swap(
        _read(_SRC, "modWorkbook.bas"),
        "Public Function IsWholeInRange(ByVal Value As Variant, ByVal MinValue As Double, _\n"
        "                               ByVal MaxValue As Double, ByRef Result As Double) As Boolean\n",
        "Public Function IsWholeInRange(ByVal Value As Variant, ByVal MinValue As Double, _\n"
        "                               ByVal MaxValue As Double, Optional ByRef Result As Double) As Boolean\n")
    _control("test_36", vba={"modWorkbook.bas": damaged})


# ===========================================================================
# The Run-4 publication verify defect
# ===========================================================================
_BLANK_RULE = """    If IsEmpty(wanted) Or (VarType(wanted) = vbString And Len(CStr(wanted)) = 0) Then
        SameCell = IsEmpty(written) Or _
                   (VarType(written) = vbString And Len(CStr(written)) = 0)
        Exit Function
    End If
    If IsEmpty(written) Then Exit Function
"""


def test_72_the_one_sided_blank_rule_is_restored() -> None:
    """The submitted shape, and Run 4 announced a restore failure over it.

    A captured block carries Empty where a built block carries vbNullString, so
    testing only `wanted` for the built spelling made Empty = Empty false.
    """
    damaged = _swap(_REPORT, _BLANK_RULE,
                    """    If IsEmpty(written) Then
        SameCell = (VarType(wanted) = vbString And Len(CStr(wanted)) = 0)
        Exit Function
    End If
""")
    _control("test_37", vba={REPORT_BAS: damaged})


def test_73_the_blank_written_guard_is_dropped() -> None:
    """Without it the predicate falls through to `IsNumeric(Empty)` - a coercion
    the project has decided not to rely on, and one Linux cannot settle."""
    damaged = _swap(_REPORT, "    If IsEmpty(written) Then Exit Function\n", "")
    _control("test_37", vba={REPORT_BAS: damaged})


def test_74_the_blank_branch_accepts_anything() -> None:
    """A verify that says yes is not a verify."""
    damaged = _swap(_REPORT,
                    """        SameCell = IsEmpty(written) Or _
                   (VarType(written) = vbString And Len(CStr(written)) = 0)
""",
                    "        SameCell = True\n")
    _control("test_37", vba={REPORT_BAS: damaged})


def test_75_the_candidate_blank_write_semantics_are_dropped() -> None:
    """A built block writes vbNullString and the cell reads back Empty; a rule
    that recognises only Empty would fail the CANDIDATE verify instead."""
    damaged = _swap(_REPORT,
                    """        SameCell = IsEmpty(written) Or _
                   (VarType(written) = vbString And Len(CStr(written)) = 0)
""",
                    "        SameCell = IsEmpty(written)\n")
    _control("test_37", vba={REPORT_BAS: damaged})


def test_76_the_predicate_stops_separating_the_two_banks() -> None:
    """The block-level control: a restore that came back naming the other bank."""
    damaged = _swap(_REPORT,
                    "        SameCell = (StrComp(CStr(written), CStr(wanted), vbBinaryCompare) = 0)\n",
                    "        SameCell = True\n")
    _control("test_38", vba={REPORT_BAS: damaged})
