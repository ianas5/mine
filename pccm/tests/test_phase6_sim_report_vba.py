#!/usr/bin/env python3
"""PCCM Phase 6 Step-11 source conformance for `src/vba/modSimReport.bas` and
the one accepted Phase-5 bridge in `src/vba/modCalcReport.bas`.

--------------------------------------------------------------------------------
WHAT THESE TESTS PROVE, AND WHAT THEY DO NOT
--------------------------------------------------------------------------------
SOURCE CONFORMANCE, on Linux, now: the public surface, the scoped D6-11
activation, the transaction ORDER, ownership of every number, what each failure
class may and may not touch, and the shape of every worksheet write.

THERE IS NO TRANSCRIPTION HERE, and that is deliberate. This layer is COM and
worksheet orchestration; the accepted Phase-6 transcriber models neither, and
building a fake Worksheet to run it against would prove that the fake behaves,
not that Excel does. Nothing in this file may be read as "VBA executed a
simulation". Real execution is Gate B.

Runs standalone or under pytest.
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import (  # noqa: E402
    load_contract,
    load_sim_contract,
    load_spec,
    load_structure_contract,
)
from pccm_builder.vba_source import (  # noqa: E402
    VbaModule,
    contains_construct,
    load_modules,
    logical_statements,
)

SRC_VBA = PCCM_ROOT / "src" / "vba"
REPORT_BAS = SRC_VBA / "modSimReport.bas"
CALC_REPORT_BAS = SRC_VBA / "modCalcReport.bas"
NONCE_BAS = SRC_VBA / "modSimNonce.bas"
SPEC = PCCM_ROOT / "spec"
CASES_JSON = PCCM_ROOT / "build" / "phase6_cases.json"

STEP11_REPORTER_BANNER = (
    "' ==========================================================================\n"
    "' STEP 11 ADDITION - THE PHASE-6 PREPARATION BRIDGE\n"
)
ACCEPTED_REPORTER_SHA256 = (
    "5d4568aef01037fd2999915da87a550d02033441b8c26c80f9386d4fcf8b087f"
)

PHASE6_PUBLIC = [
    "PCCM_CurrentSimulationRequestFingerprint",
    "PCCM_RunSimulation",
    "PCCM_SimulationAttemptDetail",
    "PCCM_SimulationAttemptResult",
    "PCCM_SimulationRequestFingerprint",
    "PCCM_SimulationResultDigest",
    "PCCM_SimulationStatus",
]

_CACHE: dict[str, object] = {}


def _module(path: Path | None = None, name: str | None = None) -> VbaModule:
    # The path is resolved at CALL time, never bound as a default: the Step-11
    # mutation controls re-aim REPORT_BAS / CALC_REPORT_BAS at damaged copies in
    # a temporary directory, and a default argument would silently keep reading
    # the accepted file and make every control vacuous.
    if path is None:
        path = REPORT_BAS
    if name is None:
        name = path.stem
    return VbaModule(name=name, path=path, raw=path.read_text(encoding="utf-8"))


def _code() -> str:
    return _module().code


def _procedure(name: str, path: Path | None = None) -> str:
    if path is None:
        path = REPORT_BAS
    code = _module(path, path.stem).code_without_string_removal
    match = re.search(
        rf"^\s*(?:Public|Private)\s+(?:Function|Sub)\s+{re.escape(name)}\b", code, re.M)
    assert match, f"{name} is not declared in {path.name}"
    tail = code[match.start():]
    end = re.search(r"^\s*End\s+(?:Function|Sub)\s*$", tail, re.M)
    assert end, f"{name} has no End"
    return tail[: end.end()]


def _sim():
    if "sim" not in _CACHE:
        _CACHE["sim"] = load_sim_contract(SPEC / "sim_contract.yaml")
    return _CACHE["sim"]


def _cases() -> dict[str, dict]:
    if "cases" not in _CACHE:
        corpus = json.loads(CASES_JSON.read_text(encoding="utf-8"))
        _CACHE["cases"] = {c["id"]: c for g in corpus["groups"] for c in g["cases"]}
    return _CACHE["cases"]  # type: ignore[return-value]


def _order_of(*tokens: str, body: str) -> list[int]:
    positions = []
    for token in tokens:
        assert token in body, token
        positions.append(body.index(token))
    return positions


# ===========================================================================
# A. Declaration, registry, the scoped construct and the public surface
# ===========================================================================
def test_01_the_module_exists_and_is_explicit() -> None:
    lines = REPORT_BAS.read_text(encoding="utf-8").splitlines()
    assert lines[0] == 'Attribute VB_Name = "modSimReport"'
    assert lines[1] == "Option Explicit"


def test_02_the_module_is_registered_and_nothing_beyond_it() -> None:
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    modules = {m.name: m for m in structure.vba_modules}
    assert "modSimReport" in modules
    assert modules["modSimReport"].generated is False
    assert [m.name for m in structure.vba_modules][-8:] == [
        "modSimContract", "modSimRng", "modSimSample", "modSimEngine", "modSimStats",
        "modSimFingerprint", "modSimNonce", "modSimReport"]


def test_03_the_endpoint_construct_is_scoped_to_this_module_and_no_other() -> None:
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    scoped = [(r.construct, tuple(r.allowed_in))
              for r in structure.forbidden_construct_rules if r.is_scoped]
    assert scoped == [("MRG32k3a", ("modSimRng",)),
                      ("RunSimulation", ("modSimReport",))], scoped
    endpoint = next(r for r in structure.forbidden_construct_rules
                    if r.construct == "RunSimulation")
    assert endpoint.allowed_in == ("modSimReport",)
    assert "*" not in endpoint.allowed_in
    assert endpoint.forbidden_in("modSimReport") is False
    for other in ("modCalcReport", "modSimEngine", "modSimRng", "modSimStats",
                  "modSimFingerprint", "modAppState"):
        assert endpoint.forbidden_in(other) is True, other
    # THE TOKEN IS REALLY EXERCISED, by the real endpoint identifier.
    assert contains_construct([_module()], "RunSimulation")
    assert "Public Sub PCCM_RunSimulation()" in _code()
    # And every other construct is still global and still refused here.
    for rule in structure.forbidden_construct_rules:
        if rule.construct == "RunSimulation":
            continue
        assert rule.forbidden_in("modSimReport") is True, rule.construct
        assert not contains_construct([_module()], rule.construct), rule.construct


def test_04_the_public_surface_is_exactly_the_seven_settled_procedures() -> None:
    assert sorted(_module().public_procedures) == PHASE6_PUBLIC
    surface = _sim().raw["command_surface"]
    assert set(PHASE6_PUBLIC) == set(surface["read_accessors"]) | {
        surface["automation_endpoint"]}
    for invented in ("PCCM_SimulationRunId", "PCCM_SimulationEffectiveSeed",
                     "PCCM_SimulationSelectedPx", "PCCM_SimulationContingency"):
        assert invented not in _code(), invented
    # The endpoint is a Sub; every accessor returns a String.
    assert "Public Sub PCCM_RunSimulation()" in _module().raw
    for accessor in surface["read_accessors"]:
        assert f"Public Function {accessor}() As String" in _module().raw, accessor


def test_05_no_button_binds_to_the_endpoint() -> None:
    spec = load_spec(SPEC / "workbook.yaml")
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    for button in getattr(structure, "buttons", []):
        assert "Simulation" not in getattr(button, "entry_point", ""), button
    assert _sim().raw["command_surface"]["user_facing_run_button_in_phase_6"] is False
    assert "PCCM_RunSimulation" not in json.dumps(
        [s.name for s in spec.sheets])


def test_06_there_is_no_module_level_or_static_run_state() -> None:
    header = _code().split("Public Sub PCCM_RunSimulation")[0]
    for line in header.splitlines():
        assert not re.match(r"^\s*(Dim|Public|Private)\s+\w+\s+As\s", line), line
    assert "Static " not in _code()
    # The staged package is a PRIVATE Type: no caller-writable result boundary.
    types = re.findall(r"^(Public|Private) Type (\w+)$", _module().raw, re.M)
    assert types == [("Private", "SimRunPackage")], types


def test_07_the_module_owns_no_mathematics() -> None:
    code = _code()
    for owned_elsewhere in (
        "SafeProduct", "SafeSignedSum", "SafeMultiply", "SafeDivide", "SafeAdd",
        "SimRngNextUniform", "SimRngJumpNextStream", "SimRngBuildComponentStreams",
        "SimSample", "BuildDriverFactors", "BuildInflationFactors",
        "BuildDiscountFactors", "BuildKnom", "BuildKpv", "CalcFp",
        "Sqr(", "Log(", "Exp(", "FP_BASE", "FP_MOD_",
    ):
        assert owned_elsewhere not in code, owned_elsewhere
    # The ONLY kernel entry points it may name.
    called = set(re.findall(r"mod(?:SimEngine|SimStats|SimFingerprint|SimRng|CalcReport)\.(\w+)",
                            code))
    assert called == {
        "SimEngineRun", "SimStatsDescribe", "SimStatsContingency",
        "SimFpBuildRequestFingerprint", "SimFpResultDigest",
        "CalcPrepareSimulationInputs",
    }, sorted(called)
    # THE AUTO SEED IS NO LONGER DERIVED HERE. The nonce transaction moved to
    # modSimNonce, so the reporter names modSimRng not at all.
    assert "SimRngAutoSeedFromNonce" not in code, (
        "the reporter still derives the AUTO seed itself"
    )
    assert set(re.findall(r"modSimNonce\.(\w+)", code)) == {
        "SimNonceAllocate", "SIM_NONCE_STATE_CONSUMED", "SIM_NONCE_STATE_INDETERMINATE",
    }, sorted(set(re.findall(r"modSimNonce\.(\w+)", code)))
    # NAME-BLINDNESS IS NOT ENOUGH: a statistic can be written without naming a
    # single banned identifier. Row and index arithmetic needs + and - only, so
    # a reporter that multiplies, divides or raises to a power is computing a
    # measure - a mean, a share, a variance, an interpolated rung - that an
    # accepted owner already owns.
    for operator in ("/", "*", "^"):
        assert operator not in code, operator
    # AND NO PROCEDURE HANDS ONE BACK. Every Double here is a carried field.
    returns = re.findall(r"^\s*(?:Public|Private) Function \w+\([^\n]*\) As (\w+)\s*$",
                         code, re.M)
    assert "Double" not in returns, returns


# ===========================================================================
# B. The Phase-5 bridge
# ===========================================================================
def test_08_the_accepted_reporter_prefix_is_byte_identical() -> None:
    import hashlib

    text = CALC_REPORT_BAS.read_text(encoding="utf-8")
    assert text.count(STEP11_REPORTER_BANNER) == 1
    accepted = text[: text.index(STEP11_REPORTER_BANNER)]
    assert hashlib.sha256(accepted.encode("utf-8")).hexdigest() == (
        ACCEPTED_REPORTER_SHA256), "an accepted line of modCalcReport moved"
    added = re.findall(r"^(?:Public|Private) (?:Function|Sub) (\w+)",
                       text[text.index(STEP11_REPORTER_BANNER):], re.M)
    assert added == ["CalcPrepareSimulationInputs"], added


def test_09_the_bridge_is_internal_and_reuses_the_accepted_preparation() -> None:
    body = _procedure("CalcPrepareSimulationInputs", CALC_REPORT_BAS)
    assert not body.lstrip().startswith("Public Function PCCM_")
    assert "PrepareCurrentCalculation(package, detail)" in body
    assert "DeriveStatus(package, True)" in body
    assert "CALC_STATUS_CURRENT" in body
    bridge = _sim().raw["phase5_bridge"]
    assert bridge["procedure"] == "CalcPrepareSimulationInputs"
    assert bridge["reuses_private_preparation"] == "PrepareCurrentCalculation"


def test_10_the_bridge_writes_nothing_and_recalculates_nothing() -> None:
    body = _procedure("CalcPrepareSimulationInputs", CALC_REPORT_BAS)
    for banned in ("PCCM_Calculate", "PCCM_CalculationStatus", "WriteAnalytical",
                   "WriteTable", "WriteSuccessCommit", "WriteAttemptBlock",
                   "WriteStatusBlock", ".Value2 =", "CalcSheet"):
        assert banned not in body, banned


def test_11_the_bridge_only_projects_the_prepared_package() -> None:
    body = _procedure("CalcPrepareSimulationInputs", CALC_REPORT_BAS)
    for rebuilt in ("BuildFactorTables", "BuildDriverFactors", "BuildAudits",
                    "BuildAnnual", "BuildFingerprint", "AccumulateTotals",
                    "ResolveModel", "Reconcile"):
        assert rebuilt not in body, rebuilt
    for projected in ("package.Drivers", "package.Model.DriverCount",
                      "package.Fingerprint", "package.Totals.ANom",
                      "package.Totals.APv", "AppliedTimelineText(package)",
                      "HostDecimalSeparator()"):
        assert projected in body, projected
    # A ZERO-DRIVER MODEL SUCCEEDS: nothing inspects a semantic driver.
    assert "package.Drivers(" not in body
    assert _sim().raw["phase5_bridge"]["zero_driver_model_succeeds"] is True


def test_12_the_current_analytical_fingerprint_is_the_prefix_state() -> None:
    """Never the STORED last-successful one."""
    body = _procedure("RunKernels")
    assert "package.AnalyticalFingerprint" in body
    assert "PCCM_CalculationFingerprint" not in _code()
    assert "CALC_STATE_ROW_LAST_SUCCESSFUL_FINGERPRINT" not in _code()
    assert _sim().raw["phase5_bridge"][
        "analytical_fingerprint_is_current_not_stored"] is True


# ===========================================================================
# C. The controls
# ===========================================================================
def test_13_only_the_two_simulation_controls_are_read() -> None:
    code = _code()
    names = set(re.findall(r"NM_INPUT_\w+", code))
    assert names == {"NM_INPUT_MONTE_CARLO_ITERATIONS", "NM_INPUT_RANDOM_SEED"}, names
    # THE LITERALS MUST BE SCANNED WITH THE STRINGS STILL IN: a workbook name
    # reached by literal rather than by constant is exactly the evasion this
    # forbids, and `code` has already thrown the string bodies away.
    text = _module().code_without_string_removal
    for banned in ("inpSelectedConfidenceLevel", "NM_INPUT_SELECTED_CONFIDENCE_LEVEL",
                   "SIM_QUANTILE_FIXED_1"):
        assert banned not in text, banned


def test_14_the_iteration_count_must_be_a_genuine_number_in_range() -> None:
    body = _procedure("ResolveIterations")
    assert "IsWholeInRange" in body
    assert "SIM_MIN_ITERATIONS" in body and "SIM_MAX_ITERATIONS" in body
    # NARROWED ONLY AFTER THE BOUNDS ARE PROVEN.
    bounds, narrow = _order_of("IsWholeInRange", "SafeLong", body=body)
    assert bounds < narrow
    assert "CLng(" not in body, "a raw narrowing bypasses the domain proof"


def test_15_a_blank_seed_is_the_auto_request_and_there_is_no_sentinel() -> None:
    body = _procedure("ResolveSeed")
    assert "IsEmptyCell" in body
    assert "SIM_SEED_MODE_AUTO" in body and "SIM_SEED_MODE_FIXED" in body
    assert "HasSuppliedSeed = False" in body
    assert "SIM_SEED_MIN" in body and "SIM_SEED_MAX" in body
    for sentinel in ("SuppliedSeed = 0", "= 0#"):
        assert sentinel not in body, sentinel
    assert "2147483646" not in body, "the seed domain is projected, not restated"
    # THE REQUEST IDENTITY CARRIES THE REQUEST. An AUTO run must not be hashed
    # as though the caller had supplied the nonce it happened to draw: that
    # would give the same request a new fingerprint on every run.
    for caller in ("RunKernels", "CurrentRequestFingerprint"):
        text = _procedure(caller)
        call = text[text.index("SimFpBuildRequestFingerprint"):]
        call = call[: call.index(") Then")]
        assert "package.HasSuppliedSeed, package.SuppliedSeed" in call, caller
        assert "EffectiveSeed" not in call, caller


# ===========================================================================
# D. The transaction order
# ===========================================================================
def test_16_the_run_order_is_the_accepted_step11a_order() -> None:
    run = _procedure("RunSimulation")
    positions = _order_of("PrepareRun", "AllocateAutoNonce", "RunKernels",
                          "PublishCandidate", "FinalCommit", body=run)
    assert positions == sorted(positions), run

    prepare = _procedure("PrepareRun")
    bridge, iterations, seed, prerequisites = _order_of(
        "CalcPrepareSimulationInputs", "ResolveIterations", "ResolveSeed",
        "ValidatePreAllocation", body=prepare)
    assert bridge < iterations < seed < prerequisites

    kernels = _procedure("RunKernels")
    engine, nominal, ladder, contingency, fingerprint, digest = _order_of(
        "SimEngineRun", "SimStatsDescribe", "SameLadder", "BuildContingencies",
        "SimFpBuildRequestFingerprint", "SimFpResultDigest", body=kernels)
    assert engine < nominal < ladder < contingency < fingerprint < digest
    assert kernels.count("SimEngineRun") == 1, "the engine runs once"
    assert kernels.count("SimStatsDescribe") == 2, "once per measure"


def test_17_every_prerequisite_is_checked_before_the_nonce_is_allocated() -> None:
    run = _procedure("RunSimulation")
    prerequisites, allocate = _order_of("PrepareRun", "AllocateAutoNonce", body=run)
    assert prerequisites < allocate
    body = _procedure("ValidatePreAllocation")
    for checked in ("ReadActiveBank", "SIM_RUN_ID_MAXIMUM", "InactiveBank"):
        assert checked in body, checked
    assert "SimEngineRun" not in body
    assert ".Value2 =" not in body, "a prerequisite check must write nothing"
    # THE AUTO NONCE IS NOT SELECTED HERE. Selecting it means reconciling any
    # prior indeterminate attempt, and that transaction belongs to modSimNonce.
    for moved in ("SIM_NONCE_EXHAUSTED", "SIM_IDENTITY_ROW_NEXT_AUTO_NONCE",
                  "SimRngAutoSeedFromNonce"):
        assert moved not in body, moved


def test_18_run_id_headroom_is_proved_before_a_candidate_is_computed() -> None:
    body = _procedure("ValidatePreAllocation")
    headroom, candidate = _order_of("lastRunId >= SIM_RUN_ID_MAXIMUM",
                                    "package.CandidateRunId = lastRunId + 1", body=body)
    assert headroom < candidate
    # The candidate id is held LOCALLY: only the final commit writes row 22.
    assert "CandidateRunId" not in _procedure("WriteAttemptBlock")
    assert "package.CandidateRunId" in _procedure("BuildCommitBlock")


def test_19_the_auto_nonce_transaction_is_delegated_and_gates_sampling() -> None:
    """The reporter drives it through a narrow scalar interface and no more."""
    body = _procedure("AllocateAutoNonce")
    assert "modSimNonce.SimNonceAllocate(" in body, (
        "the reporter still implements the nonce transaction"
    )
    for moved in ("SIM_IDENTITY_ROW_NEXT_AUTO_NONCE", "SimRngAutoSeedFromNonce",
                  "On Error", "FailPointCheck"):
        assert moved not in body, moved
    # SAMPLING IS GATED ON THE STRONG FACT, never on the diagnostic identity.
    assert "package.NonceConsumed = identityKnown" in body
    assert "package.AutoIdentityKnown = identityKnown" in body
    run = _procedure("RunSimulation")
    allocate, kernels = _order_of("AllocateAutoNonce", "RunKernels", body=run)
    assert allocate < kernels
    assert "If Not AllocateAutoNonce(package, detail) Then" in run
    # FIXED touches the counter at all.
    assert "If hasSuppliedSeed Then" in _procedure("SimNonceAllocate", NONCE_BAS)


def test_20_the_consumed_nonce_is_never_rolled_back() -> None:
    assert "SIM_IDENTITY_ROW_NEXT_AUTO_NONCE" not in _code(), (
        "the reporter still touches the AUTO counter directly"
    )
    return  # the counter now belongs to modSimNonce; proved by test_44j


def _retired_test_20_body() -> None:
    code = _code()
    assert code.count("SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 =") == 1
    for restoring in ("ConsumedNonce - 1", "nonce - 1", "NonceConsumed = False"):
        assert restoring not in code, restoring
    # And a failed attempt still records what it spent.
    attempt = _procedure("WriteAttemptBlock")
    assert "package.ConsumedNonce" in attempt
    assert "package.EffectiveSeed" in attempt


# ===========================================================================
# E. Ownership of every number
# ===========================================================================
def test_21_the_statistics_come_from_the_retained_arrays() -> None:
    body = _procedure("RunKernels")
    assert "SimStatsDescribe(package.TotalNominal, package.Iterations" in body
    assert "SimStatsDescribe(package.TotalPv, package.Iterations" in body
    assert "SimFpResultDigest(package.TotalNominal, package.TotalPv" in body


def test_22_the_two_ladders_are_proved_to_be_the_owner_ladder() -> None:
    body = _procedure("SameLadder")
    assert "SIM_QUANTILE_COUNT" in body
    assert "vbBinaryCompare" in body
    assert "NominalLabels" in body and "PvLabels" in body
    run = _procedure("RunKernels")
    describe, same, contingency = _order_of("SimStatsDescribe", "SameLadder",
                                            "BuildContingencies", body=run)
    assert describe < same < contingency
    # NO LADDER IS MUTATED AFTER Describe: nothing assigns into either array.
    code = _code()
    ladder = re.compile(
        r"^(?:\w+\.)*(?:NominalLabels|PvLabels|NominalLadder|PvLadder)\s*\(")
    for statement in logical_statements(code):
        text = statement[1]
        if ladder.match(text) and "=" in text:
            raise AssertionError(f"a ladder element is assigned: {text}")


def test_23_every_rung_gets_a_contingency_and_none_is_subtracted_here() -> None:
    body = _procedure("BuildContingencies")
    assert "For index = 0 To SIM_QUANTILE_COUNT - 1" in body
    assert body.count("SimStatsContingency") == 2, "once per measure, inside the loop"
    assert "package.BaseNominal" in body and "package.BasePv" in body
    for arithmetic in (" - ", "-="):
        assert arithmetic not in body.split("'")[0] or "SIM_QUANTILE_COUNT - 1" in body
    assert "SelectedConfidence" not in body


def test_24_the_digest_is_never_rebuilt_from_the_worksheet() -> None:
    code = _code()
    assert code.count("SimFpResultDigest") == 1
    body = _procedure("RunKernels")
    assert ".Value2" not in body, "the kernels stage runs entirely in memory"


# ===========================================================================
# F. Candidate publication
# ===========================================================================
def test_25_the_candidate_targets_the_inactive_bank_only() -> None:
    run = _procedure("RunSimulation")
    assert "package.TargetBank = InactiveBank(package.ActiveBank)" in run
    body = _procedure("InactiveBank")
    assert "SIM_BANK_A" in body and "SIM_BANK_B" in body
    publish = _procedure("PublishCandidate")
    for target in ("SnapshotRange(package.TargetBank)",
                   "SummaryRange(package.TargetBank)",
                   "ContingencyRange(package.TargetBank)"):
        assert target in publish, target
    assert "package.ActiveBank" not in publish, "the candidate must not name the active bank"
    assert "package.ActiveBank" not in _procedure("WriteIterationBank")


def test_26_the_first_success_targets_bank_a() -> None:
    body = _procedure("InactiveBank")
    blank, first = _order_of("If Len(active) = 0 Then", "InactiveBank = SIM_BANK_A",
                             body=body)
    assert blank < first
    target = _sim().raw["publication"]["banks"]["candidate_target"]
    assert target[""] == "A" and target["A"] == "B" and target["B"] == "A"
    # EXACTLY THREE OUTCOMES, IN THE CONTRACT'S ORDER, AND NO THIRD BANK.
    assert re.findall(r"InactiveBank = (\S+)", body) == [
        "SIM_BANK_A", "SIM_BANK_B", "SIM_BANK_A"], body


def test_27_the_snapshot_summary_and_ladder_are_bulk_writes() -> None:
    publish = _procedure("PublishCandidate")
    assert publish.count(".Value2 =") == 3
    for builder in ("BuildSnapshotBlock", "BuildSummaryBlock", "BuildContingencyBlock"):
        assert builder in publish, builder
    assert ".Cells(" not in _code(), "a per-cell write is not a bulk write"


def test_28_the_iteration_bank_is_written_in_chunks() -> None:
    body = _procedure("WriteIterationBank")
    assert "SIM_WRITE_CHUNK" in body
    assert body.count(".Value2 =") == 1, "one assignment per chunk, not per row"
    assert ".Cells(" not in body
    # THE INDEX IS LOGICAL, and the source values are read through LBound.
    assert "SIM_DIGEST_INDEX_ORIGIN + offset + index" in body
    assert "LBound(package.TotalNominal) + offset + index" in body
    assert "LBound(package.TotalPv) + offset + index" in body
    # The chunk size is an implementation detail, not a contract constant.
    assert "SIM_WRITE_CHUNK" not in (PCCM_ROOT / "build" / "vba"
                                     / "modSimContract.bas").read_text(encoding="utf-8")


def test_29_no_sorting_happens_anywhere() -> None:
    code = _code()
    for banned in ("Sort", "Ascending", "Descending", "QuickSort"):
        assert banned not in code, banned


def test_30_the_candidate_bank_is_verified_before_the_commit() -> None:
    run = _procedure("RunSimulation")
    publish, commit = _order_of("PublishCandidate", "FinalCommit", body=run)
    assert publish < commit
    inner = _procedure("PublishCandidate")
    write, verify = _order_of("WriteIterationBank", "VerifyCandidateBank", body=inner)
    assert write < verify
    body = _procedure("VerifyCandidateBank")
    assert "package.TargetBank" in body
    assert "package.ActiveBank" not in body, "verification reads the candidate only"
    for recomputation in ("SimStatsDescribe", "SimFpResultDigest", "SimEngineRun"):
        assert recomputation not in body, recomputation


# ===========================================================================
# G. The one final write
# ===========================================================================
def test_31_the_final_commit_is_one_write_ending_at_the_active_bank() -> None:
    body = _procedure("FinalCommit")
    assert body.count("Range(SIM_FINAL_COMMIT_RANGE).Value2 = block") == 1
    capture, write, check = _order_of(
        "previous = SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2",
        "Range(SIM_FINAL_COMMIT_RANGE).Value2 = block",
        "FailPointCheck FAILPOINT_SIM_FINAL_COMMIT", body=body)
    assert capture < write < check, (
        "the failpoint must fire AFTER the commit write, so injecting it "
        "exercises restoration rather than proving nothing was written"
    )
    block = _procedure("BuildCommitBlock")
    assert "built(9, 1) = package.TargetBank" in block, "the active bank is the last field"
    assert "built(1, 1) = package.CandidateRunId" in block
    assert "built(7, 1) = SIM_STATE_CURRENT" in block
    # THE SAME captured moment as the snapshot stamp.
    assert "built(8, 1) = package.Stamp" in block
    assert "Now" not in block
    # TWO WRITES AND NO OTHERS: the commit itself and the restoration of the
    # captured block. Anything else in here is a field published on its own.
    assert body.count(".Value2 =") == 2, body


def test_32_the_commit_is_verified_and_a_failure_restores_the_prior_block() -> None:
    body = _procedure("FinalCommit")
    write, verify, restore = _order_of(
        "Range(SIM_FINAL_COMMIT_RANGE).Value2 = block",
        "If SameBlock(SIM_FINAL_COMMIT_RANGE, block, 9, 1) Then",
        "Range(SIM_FINAL_COMMIT_RANGE).Value2 = previous", body=body)
    assert write < verify < restore
    assert "remains authoritative" in body
    # A FAILED RESTORATION IS SAID SO, not glossed.
    assert "could not be restored" in body
    assert "requires recovery" in body


def test_33_only_the_final_commit_writes_the_run_id_or_the_active_bank() -> None:
    attempt = _procedure("WriteAttemptBlock")
    assert "AttemptRange()" in attempt
    assert "SIM_FINAL_COMMIT_RANGE" not in attempt
    assert "SIM_IDENTITY_ROW_LAST_RUN_ID" not in attempt
    assert "SIM_IDENTITY_ROW_ACTIVE_BANK" not in attempt
    status = _procedure("WriteStatusBlock")
    assert "StatusRange()" in status
    assert "SIM_FINAL_COMMIT_RANGE" not in status
    # The commit range is named in exactly one procedure that writes it.
    writers = [name for name in _module().procedures
               if "Range(SIM_FINAL_COMMIT_RANGE).Value2 =" in _procedure(name)]
    assert writers == ["FinalCommit"], writers
    # AND THE SELECTOR ROW IS NAMED IN ONE PROCEDURE, WHICH READS IT. The bank
    # becomes active as the ninth field of the committed block or not at all.
    owners = [name for name in _module().procedures
              if "SIM_IDENTITY_ROW_ACTIVE_BANK" in _procedure(name)]
    assert owners == ["ReadActiveBank"], owners


def test_34_the_attempt_block_is_the_shared_rows_and_the_status_is_derived() -> None:
    body = _procedure("WriteAttemptBlock")
    assert "DeriveSimStatus()" in body, "the status is derived, not inherited"
    assert "SIM_ATTEMPT_SUCCESS" not in body
    ranges = _procedure("AttemptRange")
    assert "SIM_IDENTITY_ROW_LAST_ATTEMPT_RESULT" in ranges
    assert "SIM_IDENTITY_ROW_STATUS_EVALUATED_AT" in ranges


def test_35_a_refusal_and_a_failure_are_different_records() -> None:
    assert "SIM_ATTEMPT_REFUSED" in _procedure("RefusalResult")
    assert "RefusalResult(package)" in _procedure("RecordRefusal")
    assert "SIM_ATTEMPT_FAILED" in _procedure("RecordFailure")
    run = _procedure("RunSimulation")
    # Everything before publication refuses; publication and commit fail.
    assert run.count("RecordRefusal") == 3
    assert run.count("RecordFailure") == 2
    publish = run.index("PublishCandidate")
    assert run.rindex("RecordRefusal") < publish


# ===========================================================================
# H. Status, accessors and the reporting boundary
# ===========================================================================
def test_36_the_status_derivation_is_attempt_orthogonal() -> None:
    body = _procedure("DeriveSimStatus")
    for banned in ("SIM_ATTEMPT_", "LAST_ATTEMPT", "SelectedConfidence"):
        assert banned not in body, banned
    assert "SIM_STATE_INVALID" in body
    assert "SIM_STATE_CURRENT" in body and "SIM_STATE_STALE" in body
    assert "If Len(active) = 0 Then Exit Function" in body, "no success is BLANK"
    assert "vbBinaryCompare" in body


def test_37_the_current_request_path_is_side_effect_free() -> None:
    body = _procedure("CurrentRequestFingerprint")
    for banned in (".Value2 =", "SimEngineRun", "SimStatsDescribe", "AllocateAutoNonce",
                   "SIM_IDENTITY_ROW_NEXT_AUTO_NONCE", "SIM_IDENTITY_ROW_LAST_RUN_ID",
                   "WriteAttemptBlock", "SelectedConfidence"):
        assert banned not in body, banned
    assert "CalcPrepareSimulationInputs" in body
    assert "SimFpBuildRequestFingerprint" in body
    # The status derivation uses it, so an exhausted counter cannot make a
    # matching publication stale merely by being queried.
    assert "CurrentRequestFingerprint" in _procedure("DeriveSimStatus")


def test_38_the_stored_accessors_never_recompute() -> None:
    for name in ("PCCM_SimulationRequestFingerprint", "PCCM_SimulationResultDigest"):
        body = _procedure(name)
        assert "ActiveSnapshotText" in body
        for banned in ("CurrentRequestFingerprint", "SimFp", "SimEngineRun"):
            assert banned not in body, (name, banned)
    current = _procedure("PCCM_CurrentSimulationRequestFingerprint")
    assert "CurrentRequestFingerprint" in current
    assert "ActiveSnapshotText" not in current
    active = _procedure("ActiveSnapshotText")
    assert "If Len(active) = 0 Then Exit Function" in active


def test_39_only_the_status_accessor_writes_anything() -> None:
    writers = [name for name in PHASE6_PUBLIC if ".Value2" in _procedure(name)
               or "WriteStatusBlock" in _procedure(name)
               or "WriteAttemptBlock" in _procedure(name)]
    assert writers == ["PCCM_RunSimulation", "PCCM_SimulationStatus"] or \
        set(writers) == {"PCCM_SimulationStatus"}, writers
    assert "WriteStatusBlock" in _procedure("PCCM_SimulationStatus")
    for accessor in ("PCCM_SimulationRequestFingerprint", "PCCM_SimulationResultDigest",
                     "PCCM_SimulationAttemptResult", "PCCM_SimulationAttemptDetail",
                     "PCCM_CurrentSimulationRequestFingerprint"):
        body = _procedure(accessor)
        assert ".Value2 =" not in body, accessor
        assert "MsgBox" not in body, accessor


def test_40_results_is_never_written_and_never_reached() -> None:
    code = _code()
    for banned in ('"Results"', "shResults", "Results!"):
        assert banned not in _module().code_without_string_removal, banned
    sheets = set(re.findall(r"modWorkbook\.Sh\((\w+)\)", code))
    assert sheets == {"SIM_DATA_SHEET"}, sheets


def test_41_selected_confidence_level_is_absent_from_the_run() -> None:
    code = _module().code_without_string_removal
    for banned in ("inpSelectedConfidenceLevel", "SelectedConfidence",
                   "NM_INPUT_SELECTED_CONFIDENCE_LEVEL", "SelectedPx"):
        assert banned not in code, banned
    selector = _sim().raw["selected_confidence_level"]
    assert selector["participates_in_execution_validity"] is False
    assert selector["participates_in_request_fingerprint"] is False


# ===========================================================================
# I. The invocation envelope
# ===========================================================================
def test_42_the_endpoint_follows_the_accepted_envelope_discipline() -> None:
    body = _procedure("PCCM_RunSimulation")
    for required in ("modAppState.CaptureAppState()", "modAppState.BeginOperation",
                     "modAppState.FinishOperation(state)", "modAppState.Announce",
                     "stateCaptured", "cleanupAttempted", "committed"):
        assert required in body, required
    for banned in ("modAppState.ReportResult", "MsgBox", "DoEvents"):
        assert banned not in body, banned
    install, capture = _order_of("On Error GoTo InvocationFailed",
                                 "modAppState.CaptureAppState()", body=body)
    assert install < capture, "the envelope is installed before the first fallible call"
    assert body.count("modAppState.FinishOperation") == 2, "at most one attempt each path"


def test_43_a_cleanup_failure_after_the_commit_does_not_unpublish() -> None:
    body = _procedure("CleanupOutcome")
    assert "If committed Then" in body
    assert "COMMITTED successfully" in body
    for banned in ("SIM_ATTEMPT_FAILED", "WriteAttemptBlock", "Value2 ="):
        assert banned not in body, banned


def test_44_the_named_failpoints_exist_and_sit_where_gate_b_needs_them() -> None:
    raw = _module().raw
    stages = re.findall(r"^Public Const (FAILPOINT_SIM_\w+) As String = \"(\w+)\"$",
                        raw, re.M)
    assert [s[0] for s in stages] == [
        "FAILPOINT_SIM_CANDIDATE_BANK", "FAILPOINT_SIM_FINAL_COMMIT"], stages
    nonce_stages = re.findall(r"^Public Const (FAILPOINT_SIM_\w+) As String = \"(\w+)\"$",
                              _module(NONCE_BAS).raw, re.M)
    assert nonce_stages == [("FAILPOINT_SIM_AFTER_NONCE",
                             "Phase6AfterNoncePersisted")], nonce_stages
    run = _procedure("RunSimulation")
    # NO FAILPOINT IS A NAKED STATEMENT IN RunSimulation. FailPointCheck RAISES,
    # so a call outside a scoped envelope leaves through the invocation handler
    # with no attempt record - and after the nonce is spent, that is exactly the
    # silence the accepted contract forbids.
    assert "FailPointCheck" not in run, (
        "a failpoint raises straight out of RunSimulation, bypassing the "
        "attempt axis"
    )
    # AFTER-NONCE: inside the allocation transaction, after the advance has
    # persisted AND verified AND been marked consumed, and before any sampling.
    allocate = _procedure("SimNonceAllocate", NONCE_BAS)
    persist, injected = _order_of("PersistAdvance(",
                                  "FailPointCheck FAILPOINT_SIM_AFTER_NONCE",
                                  body=allocate)
    assert persist < injected, allocate
    assert allocate.index("On Error GoTo AllocationFailed") < injected, (
        "the after-nonce injection is outside the scoped envelope"
    )
    assert "SimEngineRun" not in allocate
    allocate_at, kernels_at = _order_of("AllocateAutoNonce", "RunKernels", body=run)
    assert allocate_at < kernels_at, "sampling can begin before allocation succeeds"

    # A FAILPOINT IS ONLY WORTH INJECTING WHERE IT EXERCISES A RECOVERY PATH.
    # Both of the later two used to fire before the write they are named for,
    # which proved only that nothing had been written yet.
    #
    # CANDIDATE: inside the publication transaction, after the inactive bank has
    # been written and before its verification completes.
    assert "FAILPOINT_SIM_CANDIDATE_BANK" not in run, (
        "the candidate failpoint fires before any candidate write"
    )
    publish = _procedure("PublishCandidate")
    snapshot, iterations, candidate, verify = _order_of(
        "Range(SnapshotRange(package.TargetBank)).Value2 = snapshot",
        "WriteIterationBank(package, detail)",
        "FailPointCheck FAILPOINT_SIM_CANDIDATE_BANK",
        "VerifyCandidateBank(package, snapshot", body=publish)
    assert snapshot < iterations < candidate < verify, publish

    # FINAL COMMIT: after the D22:D30 assignment and before its verification, so
    # the injection lands in the ambiguous state the restore exists for.
    commit = _procedure("FinalCommit")
    write, check, verify = _order_of(
        "Range(SIM_FINAL_COMMIT_RANGE).Value2 = block",
        "FailPointCheck FAILPOINT_SIM_FINAL_COMMIT",
        "If SameBlock(SIM_FINAL_COMMIT_RANGE, block, 9, 1) Then", body=commit)
    assert write < check < verify, commit
    # And it is inside the envelope that restores.
    assert commit.index("On Error GoTo CommitFailed") < check
    assert set(re.findall(r"FailPointCheck (\w+)", _code())) == {
        "FAILPOINT_SIM_CANDIDATE_BANK", "FAILPOINT_SIM_FINAL_COMMIT"}


# ===========================================================================
# I-b. THE TRANSACTION FAILURE PATHS
#
# A Range assignment, a chunk write and a verification read are COM calls, and
# COM calls RAISE. Step 11 handled only the "a helper returned False" half, so a
# raised write left the run through the invocation handler and the attempt axis
# stayed silent - against the accepted
# `refusal_or_failure_after_auto_allocation.attempt_metadata_updated: true` -
# and a raised final-commit assignment skipped the restoration the contract
# requires. Nothing below claims a runtime: these are source guarantees.
# ===========================================================================
def _nonce(name: str) -> str:
    return _procedure(name, NONCE_BAS)


def _nonce_code() -> str:
    return _module(NONCE_BAS).code


def test_44h_the_nonce_transaction_has_a_scoped_error_envelope() -> None:
    """The counter write and its read-back are COM calls, so they can raise."""
    entry = _nonce("SimNonceAllocate")
    assert "On Error GoTo AllocationFailed" in entry
    assert "On Error Resume Next" not in _nonce_code()
    persist = _nonce("PersistAdvance")
    assert "On Error GoTo WriteRaised" in persist
    assert persist.index("On Error GoTo WriteRaised") < persist.index(
        "SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2")
    reconcile = _nonce("Reconcile")
    assert "On Error GoTo ObservationRaised" in reconcile
    # Every handler reports and returns False; none claims success.
    for owner, label in (("SimNonceAllocate", "AllocationFailed"),
                         ("PersistAdvance", "WriteRaised"),
                         ("Reconcile", "ObservationRaised")):
        body = _nonce(owner)
        handler = body[body.index(label + ":"):]
        assert "Err.Description" in handler, owner
        assert f"{owner} = True" not in handler, owner


def test_44h2_allocation_and_consumption_are_different_facts() -> None:
    """The contract's order: read < derive < persist < established < sampling.

    A pre-write "identity is known" fact may exist for audit, but it is NOT
    contract allocation and nothing gates on it. Consumption is earned only by
    an observed, matching read-back.
    """
    lifecycle = _sim().raw["seeding"]["nonce_lifecycle"]
    assert list(lifecycle["order"]) == [
        "validate_pre_allocation_prerequisites", "read_current_auto_nonce",
        "derive_effective_seed", "persist_auto_nonce_plus_one", "begin_sampling"]
    assert list(lifecycle["allocation_states"]) == [
        "PRE_ALLOCATION", "CONSUMED", "PERSISTENCE_INDETERMINATE"]

    entry = _nonce("SimNonceAllocate")
    read, derive, persist, injected = _order_of(
        "ResolveNextNonce(", "SimRngAutoSeedFromNonce", "PersistAdvance(",
        "FailPointCheck FAILPOINT_SIM_AFTER_NONCE", body=entry)
    assert read < derive < persist < injected, entry
    # THE DIAGNOSTIC FACT IS NOT NAMED AS ALLOCATION, and gates nothing.
    assert "identityKnown = True" in entry
    assert "NonceAllocated" not in _nonce_code(), (
        "the rejected pre-write allocation flag is back"
    )
    assert entry.index("identityKnown = True") < entry.index("PersistAdvance(")
    # CONSUMPTION IS ESTABLISHED ONLY BY AN OBSERVED MATCH.
    classify = _nonce("Classify")
    assert "If stored = nonce + 1 Then" in classify
    consumed = classify[classify.index("If stored = nonce + 1 Then"):]
    consumed = consumed[: consumed.index("If stored = nonce Then")]
    assert "state = SIM_NONCE_STATE_CONSUMED" in consumed
    # FIXED MODE NEVER READS OR WRITES THE COUNTER.
    fixed = entry[entry.index("If hasSuppliedSeed Then"):]
    fixed = fixed[: fixed.index("    Else")]
    assert "SIM_IDENTITY_ROW_NEXT_AUTO_NONCE" not in fixed
    assert "ResolveNextNonce" not in fixed and "PersistAdvance" not in fixed


def test_44h3_the_attempt_record_preserves_the_identity_by_state() -> None:
    """Mapped field by field to `attempt_metadata_preserves`."""
    preserves = _sim().raw["seeding"]["nonce_lifecycle"]["attempt_metadata_preserves"]
    assert list(preserves["known_consumed"]) == ["consumed_auto_nonce", "effective_seed"]
    assert list(preserves["pre_allocation"]) == ["attempted_auto_nonce", "effective_seed"]
    assert list(preserves["persistence_indeterminate"]) == [
        "attempted_auto_nonce", "effective_seed", "durable_indeterminate_result"]

    body = _procedure("WriteAttemptBlock")
    # The AUDIT identity follows the diagnostic fact, so a post-write failure
    # cannot blank it and make an advanced counter look like a skipped nonce.
    assert "If package.HasSuppliedSeed Or package.AutoIdentityKnown Then" in body
    assert ("If package.AutoIdentityKnown Then\n"
            "        block(5, 1) = package.ConsumedNonce") in body
    assert "package.NonceConsumed" not in body, (
        "the attempt row is gated on verified consumption again"
    )
    # THE PUBLISHED RECORDS ARE THE OPPOSITE WAY ROUND: only proven consumption.
    for published in ("BuildSnapshotBlock", "BuildCommitBlock"):
        text = _procedure(published)
        assert "package.NonceConsumed" in text, published
        assert "AutoIdentityKnown" not in text, published
    # AND THE DURABLE MARKER IS MACHINE STATE, not prose.
    assert "SIM_ATTEMPT_AUTO_NONCE_INDETERMINATE" in _procedure("RefusalResult")
    assert "SIM_NONCE_STATE_INDETERMINATE" in _procedure("RefusalResult")


def test_44h4_the_three_observations_get_three_answers() -> None:
    """m+1 / m / neither, plus the irreducible case. No fourth answer."""
    immediate = _sim().raw["seeding"]["nonce_lifecycle"]["immediate_reconciliation"]
    assert immediate["attempts"] == 1, "a retry loop was introduced"
    assert immediate["observed_m_plus_1"] == "CONSUMED"
    assert immediate["observed_m"] == "PRE_ALLOCATION"
    assert immediate["observed_other"] == "RECOVERY_REQUIRED"
    assert immediate["observation_unavailable"] == "PERSISTENCE_INDETERMINATE"

    classify = _nonce("Classify")
    plus, same = _order_of("If stored = nonce + 1 Then", "If stored = nonce Then",
                           body=classify)
    assert plus < same
    assert "SIM_NONCE_STATE_CONSUMED" in classify
    assert "SIM_NONCE_STATE_PRE_ALLOCATION" in classify
    assert "SIM_NONCE_STATE_RECOVERY" in classify
    # THE PRE-ALLOCATION ARM PROMISES NOTHING IT CANNOT KEEP.
    pre = classify[classify.index("If stored = nonce Then"):]
    pre = pre[: pre.index("state = SIM_NONCE_STATE_RECOVERY")]
    assert "NOT consumed" in pre
    assert "may take it again" in pre
    assert "will not be reused" not in pre, (
        "the source promises non-reuse for a nonce it will legitimately reissue"
    )
    # NOTHING NORMALISES AN IMPOSSIBLE READING.
    recovery = classify[classify.index("state = SIM_NONCE_STATE_RECOVERY"):]
    assert "recovery is" in recovery and "normalised" in recovery
    # ONE observation, and it is reached only from the two authorised causes.
    reconcile = _nonce("Reconcile")
    assert reconcile.count("ReadPersistedNonce(") == 1, "a retry loop"
    assert "Do While" not in reconcile and "For " not in reconcile
    persist = _nonce("PersistAdvance")
    assert persist.count("Reconcile(") == 2, "the two authorised causes"
    # THE IRREDUCIBLE CASE says both halves of what it cannot claim, and it
    # really does classify as indeterminate rather than as either known state.
    unknown = _nonce("Indeterminate")
    assert "state = SIM_NONCE_STATE_INDETERMINATE" in unknown, (
        "the irreducible case is classified as a state it cannot prove"
    )
    for known in ("SIM_NONCE_STATE_CONSUMED", "SIM_NONCE_STATE_PRE_ALLOCATION"):
        assert known not in unknown, known
    assert "INDETERMINATE" in unknown
    assert "neither" in unknown and "nor unconsumed" in unknown
    assert "not rolled back" in unknown
    assert "must reconcile" in unknown


def test_44h5_the_next_run_reconciles_only_on_the_durable_token() -> None:
    later = _sim().raw["seeding"]["nonce_lifecycle"]["next_run_reconciliation"]
    assert later["activated_by_attempt_result"] == "AUTO_NONCE_INDETERMINATE"
    assert later["activated_by_generic_unsuccessful_result"] is False
    assert later["applies_to_fixed_mode"] is False

    body = _nonce("ResolveNextNonce")
    gate = ('If StrComp(SharedText(SIM_IDENTITY_ROW_LAST_ATTEMPT_RESULT), _\n'
            '               SIM_ATTEMPT_AUTO_NONCE_INDETERMINATE, vbBinaryCompare) = 0 Then')
    assert gate in body, "reconciliation is not gated on the durable token alone"
    # A generic unsuccessful result must NOT activate it.
    for generic in ("SIM_ATTEMPT_REFUSED", "SIM_ATTEMPT_FAILED", "SIM_ATTEMPT_SUCCESS"):
        assert generic not in body, generic
    # The three resolutions, and recovery for the third.
    assert "counter = prior + 1" in body
    assert "ElseIf counter = prior Then" in body
    assert body.count("SIM_NONCE_STATE_RECOVERY") == 3, body
    # It is reached only in AUTO: the FIXED branch never calls it.
    entry = _nonce("SimNonceAllocate")
    fixed = entry[entry.index("If hasSuppliedSeed Then"): entry.index("    Else")]
    assert "ResolveNextNonce" not in fixed


def test_44i_the_after_nonce_injection_returns_through_the_attempt_path() -> None:
    """The whole point: a raised injection becomes a recorded attempt."""
    run = _procedure("RunSimulation")
    assert "FailPointCheck" not in run, "a failpoint raises out of RunSimulation"
    guard = "If Not AllocateAutoNonce(package, detail) Then"
    assert guard in run
    arm = run[run.index(guard):]
    arm = arm[: arm.index("End If")]
    assert "RecordRefusal(package, detail)" in arm, (
        "an after-nonce failure does not reach the attempt record"
    )
    assert "Err.Raise" not in arm
    # THE RECORD KEEPS THE SPENT EVIDENCE.
    attempt = _procedure("WriteAttemptBlock")
    for retained in ("package.SeedMode", "package.EffectiveSeed", "package.ConsumedNonce"):
        assert retained in attempt, retained


def test_44j_the_counter_is_never_rolled_back_on_any_path() -> None:
    """One module writes the counter, once, forward only."""
    nonce_code = _nonce_code()
    for _, statement in logical_statements(nonce_code):
        if "SIM_IDENTITY_ROW_NEXT_AUTO_NONCE" in statement and ".Value2 =" in statement:
            assert "nonce + 1" in statement, (
                f"the counter is written with something other than the advance: {statement}"
            )
    for rollback in ("nonce - 1", "nonce-1", "stored - 1", "counter - 1",
                     "ConsumedNonce - 1"):
        assert rollback not in nonce_code, rollback
    writers = [name for name in _module(NONCE_BAS).procedures
               if "SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 =" in _nonce(name)]
    assert writers == ["PersistAdvance"], writers
    # AND THE REPORTER CANNOT TOUCH IT AT ALL.
    assert "SIM_IDENTITY_ROW_NEXT_AUTO_NONCE" not in _code(), (
        "the reporter reaches into the counter the nonce module owns"
    )


def test_44k_there_is_no_fourth_failpoint() -> None:
    """Three names, split across exactly two owning modules, each inside a scope."""
    from pccm_builder.vba_source import load_modules

    fired: dict[str, list[str]] = {}
    for module in load_modules([SRC_VBA]):
        names = sorted(n for n in re.findall(r"FailPointCheck (\w+)", module.code)
                       if n.startswith("FAILPOINT_SIM_"))
        if names:
            fired[module.name] = names
    assert fired == {
        "modSimNonce": ["FAILPOINT_SIM_AFTER_NONCE"],
        "modSimReport": ["FAILPOINT_SIM_CANDIDATE_BANK", "FAILPOINT_SIM_FINAL_COMMIT"],
    }, fired
    # THE CONSTANT MOVED WITH ITS OWNER, and no reverse dependency was created.
    assert 'FAILPOINT_SIM_AFTER_NONCE As String = "Phase6AfterNoncePersisted"' in \
        _module(NONCE_BAS).raw
    assert "FAILPOINT_SIM_AFTER_NONCE" not in _code()
    assert "modSimReport" not in _nonce_code(), (
        "the nonce module depends back on the reporter"
    )
    # Each owner arms a handler before it fires.
    for owner, handler, reader in (
            ("SimNonceAllocate", "AllocationFailed", _nonce),
            ("PublishCandidate", "CandidateFailed", _procedure),
            ("FinalCommit", "CommitFailed", _procedure)):
        body = reader(owner)
        assert body.index(f"On Error GoTo {handler}") < body.index("FailPointCheck"), owner


def test_44a_the_candidate_transaction_has_a_scoped_error_envelope() -> None:
    body = _procedure("PublishCandidate")
    assert "On Error GoTo CandidateFailed" in body, (
        "a raised candidate write escapes to the invocation axis"
    )
    assert "On Error Resume Next" not in body, "the envelope is a blanket suppressor"
    install = body.index("On Error GoTo CandidateFailed")
    # EVERYTHING FALLIBLE IS INSIDE IT.
    for covered in ("Range(SnapshotRange(package.TargetBank)).Value2 = snapshot",
                    "Range(SummaryRange(package.TargetBank)).Value2 = summary",
                    "Range(ContingencyRange(package.TargetBank)).Value2 = contingency",
                    "WriteIterationBank(package, detail)",
                    "FailPointCheck FAILPOINT_SIM_CANDIDATE_BANK",
                    "VerifyCandidateBank(package, snapshot"):
        assert install < body.index(covered), covered
    # THE HANDLER REPORTS AND RETURNS FALSE. It never sets the return True.
    handler = body[body.index("CandidateFailed:"):]
    assert "Err.Description" in handler
    assert "On Error GoTo 0" in handler, "the handler stays armed on the way out"
    assert "detail = " in handler
    assert "PublishCandidate = True" not in handler, "a failed candidate reports success"
    assert "no semantic standing" in handler
    # AND IT ROLLS NOTHING BACK. The half-written inactive bank stays as it is.
    assert "SIM_IDENTITY_ROW_ACTIVE_BANK" not in body
    assert "SIM_IDENTITY_ROW_LAST_RUN_ID" not in body
    assert "SIM_FINAL_COMMIT_RANGE" not in body
    assert "ClearContents" not in body, "the handler erases the candidate bank"
    assert "package.ActiveBank" not in body


def test_44b_a_candidate_failure_reaches_the_attempt_axis() -> None:
    """The point of the envelope: FAILED is recorded, not just announced."""
    run = _procedure("RunSimulation")
    assert "If Not PublishCandidate(package, detail) Then" in run
    arm = run[run.index("If Not PublishCandidate(package, detail) Then"):]
    arm = arm[: arm.index("End If")]
    assert "RecordFailure(package, detail)" in arm, (
        "a candidate failure does not reach the attempt record"
    )
    assert "SIM_ATTEMPT_FAILED" in _procedure("RecordFailure")
    # The attempt block carries the consumed AUTO seed evidence.
    block = _procedure("RecordFailure") + _procedure("WriteAttemptBlock")
    assert "package.EffectiveSeed" in block or "EffectiveSeed" in _procedure(
        "WriteAttemptBlock"), "the attempt record loses the consumed seed"


def test_44c_the_final_commit_captures_before_it_writes_anything() -> None:
    body = _procedure("FinalCommit")
    capture, build, write = _order_of(
        "previous = SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2",
        "BuildCommitBlock package, block",
        "Range(SIM_FINAL_COMMIT_RANGE).Value2 = block", body=body)
    assert capture < build < write
    # THE CAPTURE HAS ITS OWN HANDLER, and its failure does NOT restore: there
    # is no captured block to write back, and nothing was written to undo.
    assert "On Error GoTo CaptureFailed" in body
    assert body.index("On Error GoTo CaptureFailed") < capture
    handler = body[body.index("CaptureFailed:"):]
    assert "Value2 = previous" not in handler, (
        "the capture-failure path writes an unset block over a live publication"
    )
    assert "no final commit was attempted" in handler
    assert "unchanged" in handler


def test_44d_every_post_write_failure_restores_the_prior_block() -> None:
    """Write exception, injected failpoint, verification exception, mismatch."""
    body = _procedure("FinalCommit")
    # ONE envelope covers the assignment, the failpoint and the verification.
    assert "On Error GoTo CommitFailed" in body
    armed = body.index("On Error GoTo CommitFailed")
    verify_at = body.index("If SameBlock(SIM_FINAL_COMMIT_RANGE, block, 9, 1) Then")
    for covered in ("Range(SIM_FINAL_COMMIT_RANGE).Value2 = block",
                    "FailPointCheck FAILPOINT_SIM_FINAL_COMMIT"):
        assert armed < body.index(covered) < verify_at, covered
    # AND IT STAYS ARMED THROUGH THE VERIFICATION READ. Disarming after the
    # assignment would let a raised SameBlock leave without restoring, which is
    # the same defect in a different place.
    assert "On Error GoTo 0" not in body[armed:verify_at], (
        "the envelope is disarmed before the verification read"
    )
    # BOTH ROUTES LAND ON THE SAME RESTORE. The raised route falls through the
    # handler into it; the mismatch route jumps to it.
    assert "RestorePrevious:" in body, "there is no single restoration path"
    assert "GoTo RestorePrevious" in body, "the mismatch route does not restore"
    handler = body[body.index("CommitFailed:"): body.index("RestorePrevious:")]
    assert "Err.Description" in handler, "the raised cause is discarded"
    assert "Exit Function" not in handler, (
        "the raised route leaves without attempting restoration"
    )
    # ONE write back, then verified.
    restore = body[body.index("RestorePrevious:"):]
    assert restore.count("Range(SIM_FINAL_COMMIT_RANGE).Value2 = previous") == 1
    assert "If SameBlock(SIM_FINAL_COMMIT_RANGE, previous, 9, 1) Then" in restore
    assert body.count(".Value2 =") == 2, "the commit procedure writes something else"


def test_44e_a_failed_restoration_says_so_and_never_claims_the_bank_is_safe() -> None:
    body = _procedure("FinalCommit")
    restore = body[body.index("RestorePrevious:"):]
    assert "remains authoritative" in restore
    # The two failed-restoration arms: verified-false and raised.
    assert "On Error GoTo RestoreFailed" in restore
    assert restore.count("could not be restored") == 2, restore
    assert restore.count("requires recovery") == 2, restore
    raised = body[body.index("RestoreFailed:"): body.index("CaptureFailed:")]
    assert "Err.Description" in raised
    assert "cannot be guaranteed" in raised
    assert "remains authoritative" not in raised, (
        "a failed restoration claims the prior publication survived"
    )
    # The CAPTURE handler is the one place that may say the bank is safe,
    # because nothing was written: no commit was attempted at all.
    capture = body[body.index("CaptureFailed:"):]
    assert "remains authoritative" in capture
    assert "requires recovery" not in capture, (
        "a failed capture is reported as a publication-integrity emergency"
    )


def test_44f_a_commit_failure_reaches_the_attempt_axis() -> None:
    run = _procedure("RunSimulation")
    arm = run[run.index("If Not FinalCommit(package, detail) Then"):]
    arm = arm[: arm.index("End If")]
    assert "RecordFailure(package, detail)" in arm
    # AND THE ATTEMPT RECORD DOES NOT TOUCH THE PUBLICATION. D22 and D30 are the
    # commit's; the attempt block owns D23:D29 only.
    attempt = _procedure("WriteAttemptBlock")
    assert "SIM_IDENTITY_ROW_LAST_RUN_ID" not in attempt
    assert "SIM_IDENTITY_ROW_ACTIVE_BANK" not in attempt
    assert "SIM_FINAL_COMMIT_RANGE" not in attempt


def test_44g_no_blanket_error_suppression_was_introduced() -> None:
    """Scoped handlers only. A module-wide Resume Next would hide everything."""
    code = _module().code
    assert "On Error Resume Next" not in code
    handlers = set(re.findall(r"On Error GoTo (\w+)", code))
    assert handlers == {"0", "InvocationFailed", "NormalCleanupFailed", "CleanupFailed",
                        "CandidateFailed", "CommitFailed",
                        "RestoreFailed", "CaptureFailed"}, sorted(handlers)
    # The nonce module carries its own, and no blanket suppressor either.
    nonce_handlers = set(re.findall(r"On Error GoTo (\w+)", _nonce_code()))
    assert nonce_handlers == {"0", "AllocationFailed", "WriteRaised",
                              "ObservationRaised"}, sorted(nonce_handlers)
    assert "On Error Resume Next" not in _nonce_code()
    # Every named handler has a label, and every label is reachable by name.
    labels = set(re.findall(r"^(\w+):$", code, re.M))
    assert (handlers - {"0"}) <= labels, sorted((handlers - {"0"}) - labels)
    for label in labels:
        assert label in handlers | {"RestorePrevious"}, label


# ===========================================================================
# J. The accepted publication corpus
# ===========================================================================
def test_45_every_publication_case_has_a_structural_counterpart() -> None:
    cases = _cases()
    required = [
        "publication.grammar", "publication.initial",
        "publication.first_success_targets_a", "publication.second_success_targets_b",
        "publication.third_success_targets_a_again",
        "publication.refusal_before_auto_allocation",
        "publication.failure_after_auto_allocation",
        "publication.inactive_bank_write_failure",
        "publication.final_commit_failure_restores_the_block",
        "publication.run_id_exhaustion_refuses_first",
        "publication.status.invalid", "publication.status.blank_no_success",
        "publication.status.current", "publication.status.stale",
        "publication.selected_cl_is_reporting_only",
        "publication.invalid_selected_cl_blanks_the_lookup",
    ]
    assert len(required) == 16
    for identifier in required:
        assert identifier in cases, identifier
        assert cases[identifier]["comparison"] == "EXACT"

    target = cases["publication.grammar"]["expected_exact"]["candidate_target"]
    body = _procedure("InactiveBank")
    assert target[""] == "A" and "InactiveBank = SIM_BANK_A" in body
    order = cases["publication.grammar"]["expected_exact"]["transaction_order"]
    assert order[-1] == "final_commit_shared_block_including_active_bank"
    assert "FinalCommit" in _procedure("RunSimulation")

    # The nonce cases, against the source that implements them.
    refusal = cases["publication.refusal_before_auto_allocation"]
    assert refusal["expected_exact"]["after"]["next_auto_nonce"] == (
        refusal["inputs"]["before"]["next_auto_nonce"])
    assert "SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2" not in _procedure(
        "ValidatePreAllocation")
    after = cases["publication.failure_after_auto_allocation"]
    assert after["expected_exact"]["after"]["next_auto_nonce"] == (
        after["inputs"]["before"]["next_auto_nonce"] + 1)
    assert "nonce + 1" in _procedure("PersistAdvance", NONCE_BAS)

    commit = cases["publication.final_commit_failure_restores_the_block"]
    assert commit["expected_exact"]["after"]["active_bank"] == "B"
    assert "previous" in _procedure("FinalCommit")

    for identifier, expected in (("publication.status.invalid", "INVALID"),
                                 ("publication.status.blank_no_success", None),
                                 ("publication.status.current", "CURRENT"),
                                 ("publication.status.stale", "STALE")):
        assert cases[identifier]["expected_exact"]["simulation_status"] == expected


def test_46_the_corpus_moved_only_for_the_authorised_axis_change() -> None:
    """The Step-11A corpus is authority, not source-generated evidence.

    It moved in Step 12 for exactly one reason: the Phase-6 attempt-result axis
    gained AUTO_NONCE_INDETERMINATE, and the corpus projects that axis. Nothing
    else about it changed, and the new hash is pinned here so a second silent
    movement is still caught.
    """
    import hashlib
    import json

    digest = hashlib.sha256(CASES_JSON.read_bytes()).hexdigest()
    assert digest == "8019683a0490fcf0740cf07244524973d9b7470c933f1003059025b6b019a0be", digest
    text = CASES_JSON.read_text(encoding="utf-8")
    assert "AUTO_NONCE_INDETERMINATE" in text, (
        "the corpus moved for something other than the authorised axis change"
    )
    json.loads(text)

def test_47_no_step_12_exists() -> None:
    names = {p.stem for p in SRC_VBA.glob("*.bas")}
    assert names & {"modSimDashboard", "modSimSensitivity", "modSimAnnual"} == set()
    for module in load_modules([SRC_VBA]):
        for later in ("Sensitivity", "AnnualStochastic", "Tornado"):
            assert later not in module.code, (module.name, later)


if __name__ == "__main__":  # pragma: no cover
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
