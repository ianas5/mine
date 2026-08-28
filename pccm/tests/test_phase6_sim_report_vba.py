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
SETTLEMENT_MD = (
    PCCM_ROOT / "docs" / "phase6_step12_transaction_settlement.md"
)

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
    # THE RECOVERY TOKEN IS NOT REACHABLE FROM HERE AT ALL. It is an action, and
    # no reporter expression may derive a physical fact from it.
    assert "SIM_NONCE_STATE_RECOVERY" not in code, (
        "the reporter reads the recovery action as if it were an allocation state"
    )
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
    # SAMPLING IS GATED ON THE STRONG FACT, never on the diagnostic identity and
    # never on the call's own Boolean - which is True for a FIXED run that
    # consumed nothing at all.
    assert ("package.NonceConsumed = (StrComp(allocationState, _\n"
            "                                     modSimNonce.SIM_NONCE_STATE_CONSUMED, _\n"
            "                                     vbBinaryCompare) = 0)") in body
    assert "package.NonceConsumed = identityKnown" not in body, (
        "consumption is inferred from the diagnostic identity again"
    )
    assert "package.NonceConsumed = allocated" not in body
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
    persist, injected = _order_of("RunAllocationTransaction(",
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


def _statements(body: str) -> list[str]:
    """Logical VBA statements, continuations joined.

    Every structural control below works on these, never on raw text. A
    `Then` split across a line continuation reads as two physical lines and
    would make a block scanner mis-pair its `End If`.
    """
    return [text for _, text in logical_statements(body)]


def _branch(body: str, opener: str) -> list[str]:
    """The statements inside ONE `If ... Then` block, nesting-aware.

    Index arithmetic over raw text cannot tell a nested `End If` from the one
    that closes the branch, and a containment claim built on it proves nothing.
    That was the exact gap in the previous failpoint control: it compared two
    text positions and concluded a branch could not reach a call.
    """
    lines = _statements(body)
    start = next(i for i, text in enumerate(lines) if text.startswith(opener))
    depth = 0
    inside: list[str] = []
    for text in lines[start:]:
        opens = bool(re.match(r"^(If|ElseIf)\b.*\bThen$", text))
        if text == "End If":
            depth -= 1
            if depth == 0:
                return inside
        if depth == 1 and re.match(r"^(Else$|ElseIf\b)", text):
            return inside
        if depth >= 1:
            inside.append(text)
        if opens:
            depth += 1
    raise AssertionError(f"unterminated block: {opener}")


def _until_first(body: str, stop: str) -> list[str]:
    """Statements from the top of a body up to the first one containing `stop`."""
    lines = _statements(body)
    index = next(i for i, text in enumerate(lines) if stop in text)
    return lines[:index]


def test_44h_a_raised_verification_read_reaches_the_one_reconciliation() -> None:
    """The advance write AND its verification read are one envelope.

    A `Range.Value2` READ is a COM call too. Disarming the write handler before
    it - which is what this source used to do - let a raised read escape into
    the outer allocation handler and skip the bounded reconciliation the
    contract requires for a verification failure. The contract now names both
    causes, and the code must route both.
    """
    immediate = _sim().raw["seeding"]["nonce_lifecycle"]["immediate_reconciliation"]
    assert list(immediate["after"]) == [
        "counter_write_raised", "verification_read_failed", "verification_read_raised"]

    persist = _statements(_nonce("PersistAdvance"))
    write = next(i for i, t in enumerate(persist)
                 if "SIM_IDENTITY_ROW_NEXT_AUTO_NONCE" in t and ".Value2 =" in t)
    verify = next(i for i, t in enumerate(persist) if "ReadPersistedNonce(" in t)
    armed = next(i for i, t in enumerate(persist) if t.startswith("On Error GoTo ")
                 and not t.endswith(" 0"))
    assert armed < write < verify, persist
    # NOTHING DISARMS BETWEEN THE WRITE AND THE READ. This is the whole control:
    # the previous shape put `On Error GoTo 0` in exactly this gap.
    assert "On Error GoTo 0" not in persist[write:verify], (
        "the handler is disarmed before the first verification read, so a raised "
        "read escapes the bounded reconciliation"
    )
    # BOTH CAUSES REACH THE ONE OBSERVATION: the helper returning False, and the
    # handler that catches a raise.
    assert _nonce("PersistAdvance").count("Reconcile(") == 2
    handler = next(t for t in persist if re.fullmatch(r"\w+:", t))
    tail = persist[persist.index(handler):]
    assert any("Reconcile(" in t for t in tail), (
        "a raised advance or verification does not reach the reconciliation"
    )
    assert any("Err.Description" in t for t in tail)
    # AND THE OBSERVATION ITSELF IS BOUNDED AND GUARDED.
    reconcile = _nonce("Reconcile")
    assert reconcile.count("ReadPersistedNonce(") == 1, "a retry loop"
    assert "On Error GoTo ObservationRaised" in reconcile
    assert "Do While" not in reconcile and "For " not in reconcile
    assert "On Error Resume Next" not in _nonce_code()


def test_44h2_allocation_and_consumption_are_different_facts() -> None:
    """The settled order: read < derive < mark < persist < clear < sample.

    A pre-write "identity is known" fact may exist for audit, but it is NOT
    contract allocation and nothing gates on it. Consumption is earned only by
    an observed, matching read-back.
    """
    lifecycle = _sim().raw["seeding"]["nonce_lifecycle"]
    assert list(lifecycle["write_ahead_order"]) == [
        "read_current_auto_nonce", "derive_effective_seed",
        "establish_and_verify_pending_auto_nonce", "persist_auto_nonce_plus_one",
        "reconcile_and_clear_pending", "begin_sampling"]
    assert list(lifecycle["allocation_states"]) == [
        "PRE_ALLOCATION", "CONSUMED", "PERSISTENCE_INDETERMINATE"]

    entry = _nonce("SimNonceAllocate")
    read, derive, transaction, injected = _order_of(
        "ResolveNextNonce(", "SimRngAutoSeedFromNonce", "RunAllocationTransaction(",
        "FailPointCheck FAILPOINT_SIM_AFTER_NONCE", body=entry)
    assert read < derive < transaction < injected, entry
    # AND INSIDE THE TRANSACTION: mark, then advance, then clear.
    txn = _statements(_nonce("RunAllocationTransaction"))
    marked = next(i for i, t in enumerate(txn) if "EstablishPending(" in t)
    advanced = next(i for i, t in enumerate(txn) if "PersistAdvance(" in t)
    cleared = next(i for i, t in enumerate(txn) if "ClearPending(" in t)
    assert marked < advanced < cleared, txn
    # THE DIAGNOSTIC FACT IS NOT NAMED AS ALLOCATION, and gates nothing.
    assert "identityKnown = True" in entry
    assert "NonceAllocated" not in _nonce_code(), (
        "the rejected pre-write allocation flag is back"
    )
    assert entry.index("identityKnown = True") < entry.index("RunAllocationTransaction(")
    # CONSUMPTION IS ESTABLISHED ONLY BY AN OBSERVED MATCH.
    classify = _nonce("Classify")
    assert "If stored = nonce + 1 Then" in classify
    consumed = classify[classify.index("If stored = nonce + 1 Then"):]
    consumed = consumed[: consumed.index("If stored = nonce Then")]
    assert "allocationState = SIM_NONCE_STATE_CONSUMED" in consumed


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
    # THE OBSOLETE AUTHORITY LANGUAGE IS GONE. It claimed a post-write
    # verification failure left "allocation certain", which the three-state
    # model refuted; keeping it because it was once accepted would leave the
    # module documenting a rejected design.
    prose = _module(REPORT_BAS).raw
    assert prose.count("'") > 100, "the raw text lost its comments"
    for withdrawn in ("ALLOCATED, not CONSUMED", "allocation certain"):
        assert withdrawn not in prose, withdrawn
    # THE PUBLISHED RECORDS ARE THE OPPOSITE WAY ROUND: only proven consumption.
    for published in ("BuildSnapshotBlock", "BuildCommitBlock"):
        text = _procedure(published)
        assert "package.NonceConsumed" in text, published
        assert "AutoIdentityKnown" not in text, published


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
    assert "SIM_NONCE_STATE_INDETERMINATE" in classify
    # THE PRE-ALLOCATION ARM PROMISES NOTHING IT CANNOT KEEP.
    pre = classify[classify.index("If stored = nonce Then"):]
    pre = pre[: pre.index("allocationState = SIM_NONCE_STATE_INDETERMINATE")]
    assert "NOT consumed" in pre
    assert "may take it again" in pre
    assert "will not be reused" not in pre, (
        "the source promises non-reuse for a nonce it will legitimately reissue"
    )
    # NOTHING NORMALISES AN IMPOSSIBLE READING.
    recovery = classify[classify.index("allocationState = SIM_NONCE_STATE_INDETERMINATE"):]
    assert "recovery is" in recovery and "normalised" in recovery
    reconcile = _nonce("Reconcile")
    assert reconcile.count("ReadPersistedNonce(") == 1, "a retry loop"
    assert "Do While" not in reconcile and "For " not in reconcile
    assert "Do" not in _statements(reconcile), reconcile
    persist = _nonce("PersistAdvance")
    assert persist.count("Reconcile(") == 2, "the two authorised causes"
    # THE IRREDUCIBLE CASE says both halves of what it cannot claim, and it
    # really does classify as indeterminate rather than as either known state.
    unknown = _nonce("Indeterminate")
    assert "allocationState = SIM_NONCE_STATE_INDETERMINATE" in unknown, (
        "the irreducible case is classified as a state it cannot prove"
    )
    for known in ("SIM_NONCE_STATE_CONSUMED", "SIM_NONCE_STATE_PRE_ALLOCATION"):
        assert known not in unknown, known
    assert "INDETERMINATE" in unknown
    assert "neither" in unknown and "nor unconsumed" in unknown
    assert "not rolled back" in unknown
    assert "must reconcile" in unknown


def test_44h5_the_next_run_reconciles_on_the_sidecar_never_the_attempt_row() -> None:
    """The carrier decision, pinned in both authorities.

    `Last Attempt Result` records the CHRONOLOGICALLY LAST attempt. A durable
    AUTO lock has to survive attempts that have nothing to do with it - a FIXED
    run, in particular, which the contract correctly excuses from AUTO
    reconciliation but which still rewrites that row. One cell cannot be both.
    """
    later = _sim().raw["seeding"]["nonce_lifecycle"]["next_run_reconciliation"]
    assert later["activated_by_pending_auto_nonce_cell"] is True
    assert later["activated_by_attempt_result"] is False
    assert later["activated_by_generic_unsuccessful_result"] is False
    assert later["applies_to_fixed_mode"] is False

    body = _nonce("ResolveNextNonce")
    # THE SIDECAR IS READ FIRST, before the counter is even consulted.
    pending, counter = _order_of("ReadPending(", "ReadPersistedNonce(", body=body)
    assert pending < counter, body
    # THE ATTEMPT AXIS IS NOT AN INPUT TO THIS DECISION - anywhere in the module.
    code = _nonce_code()
    for attempt_axis in ("SIM_IDENTITY_ROW_LAST_ATTEMPT_RESULT",
                         "SIM_IDENTITY_ROW_LAST_ATTEMPT_AUTO_NONCE",
                         "SIM_ATTEMPT_AUTO_NONCE_INDETERMINATE", "SIM_ATTEMPT_REFUSED",
                         "SIM_ATTEMPT_FAILED", "SIM_ATTEMPT_SUCCESS"):
        assert attempt_axis not in code, attempt_axis
    # THE THREE RESOLUTIONS, AND RECOVERY FOR THE FOURTH.
    assert "If counter = pending + 1 Then" in body
    assert "ElseIf counter = pending Then" in body
    # PRIOR-MARKER REFUSALS RAISE THE ACTION AND NOTHING ELSE. No allocation
    # classification is produced here: this run has not selected an identity or
    # touched the counter, so it has no transition to classify.
    assert body.count("recoveryRequired = True") == 5, body
    for physical in ("SIM_NONCE_STATE_CONSUMED", "SIM_NONCE_STATE_PRE_ALLOCATION",
                     "SIM_NONCE_STATE_INDETERMINATE"):
        assert physical not in body, physical
    # It is reached only in AUTO: the FIXED branch never calls it.
    fixed = _branch(_nonce("SimNonceAllocate"), "If hasSuppliedSeed Then")
    assert not any("ResolveNextNonce" in t for t in fixed)


def _prose(path: Path) -> str:
    """Only the comment text of a module: raw minus the comment-stripped code.

    `VbaModule.code` blanks comment lines rather than deleting them, so the two
    align line for line and the difference is exactly the prose.
    """
    module = _module(path, path.stem)
    raw = module.raw.splitlines()
    stripped = module.code.splitlines()
    assert len(raw) == len(stripped), (len(raw), len(stripped))
    return "\n".join(r for r, c in zip(raw, stripped) if r.strip() and not c.strip())


def _flatten(text: str) -> str:
    """Comment prose as one line: VBA wraps sentences across `\' ` lines."""
    return " ".join(text.replace("'", " ").split())


def _within(flat: str, left: str, right: str, distance: int) -> str | None:
    """The first window where `left` and `right` sit within `distance` chars.

    Whole-file co-occurrence proves nothing - a module that truthfully mentions
    both words in unrelated paragraphs would fail it. Proximity is what
    distinguishes a claim from a coincidence.
    """
    start = 0
    while True:
        i = flat.find(left, start)
        if i < 0:
            return None
        window = flat[i: i + distance]
        if right in window:
            return window
        start = i + 1


def _doc_section(text: str, heading: str) -> str:
    """One markdown section: its heading line up to the next heading."""
    start = text.index(heading)
    end = text.index("\n#", start + len(heading))
    return text[start:end]


def _normative(section: str) -> str:
    """The section's ACTIVE prose, flattened: blockquote lines removed.

    A withdrawn claim is kept in this document on purpose - the history of a
    defect is part of the record - but it is quoted inside a `>` block that
    withdraws it. Only what stands OUTSIDE such a block is a claim the document
    still makes, and that is the only thing a wording detector may hold to
    account. Flattening is required too: markdown wraps a sentence across lines,
    so a phrase check against the raw text would miss half of them.
    """
    kept = [line for line in section.splitlines()
            if not line.lstrip().startswith(">")]
    return " ".join(" ".join(kept).split())


def _quoted(section: str) -> str:
    """The section's blockquoted record, flattened and stripped of `> `."""
    kept = [line for line in section.splitlines()
            if line.lstrip().startswith(">")]
    return " ".join(" ".join(kept).replace(">", " ").split())


def test_44u_the_prose_names_the_authority_the_code_actually_uses() -> None:
    """Documentation drift is a real defect, not a cosmetic one.

    A module whose executable source reconciles against F21 while its header
    still says it reads attempt rows teaches the rejected design to the next
    reader, and the next reader is the one who has to decide whether a change
    is safe. This checks the two against each other rather than checking that
    a comment exists.
    """
    raw = _module(NONCE_BAS).raw
    code = _nonce_code()
    prose = raw.replace(code, "")
    assert raw.count("'") > 100, "the raw text lost its comments"

    # THE EXECUTABLE FACT: F21 is read, the attempt axis is not.
    assert "SIM_PENDING_AUTO_NONCE_CELL" in code
    for attempt_axis in ("SIM_IDENTITY_ROW_LAST_ATTEMPT_RESULT",
                         "SIM_IDENTITY_ROW_LAST_ATTEMPT_AUTO_NONCE",
                         "SIM_ATTEMPT_AUTO_NONCE_INDETERMINATE"):
        assert attempt_axis not in code, attempt_axis

    # AND THE PROSE MUST NOT CONTRADICT IT. These are the exact claims the
    # rejected Option-3R carrier made; none may stand once F21 is the authority.
    for withdrawn in ("only READS the attempt rows",
                      "READS the attempt rows",
                      "carries the durable AUTO_NONCE_INDETERMINATE marker",
                      "durable AUTO_NONCE_INDETERMINATE"):
        assert withdrawn not in prose, withdrawn
    # THE POSITIVE CLAIM IS PRESENT AND SPECIFIC.
    assert "_SimData!F21" in prose
    assert "not the attempt row" in prose.lower()
    # THE SAME CHECK ON THE REPORTER, whose comments describe the same protocol.
    #
    # SEMANTIC, NOT ONE BRITTLE SENTENCE. A single full-sentence literal is
    # trivially evaded by rewording, which is how the stale RecordRefusal
    # paragraph survived an earlier pass of this control. Each rejected
    # authority below is expressed as a CO-OCCURRENCE within one comment block,
    # so a paraphrase that still teaches the rejected design is caught.
    reporter_prose = _prose(REPORT_BAS)
    assert reporter_prose.count("'") > 200, "the reporter prose was not extracted"
    flat = _flatten(reporter_prose)

    for withdrawn in ("ALLOCATED, not CONSUMED", "allocation certain"):
        assert withdrawn not in flat, withdrawn

    # THE FIVE REJECTED AUTHORITIES. Each is an AFFIRMATIVE phrase or a
    # proximity pair, chosen so the accepted prose - which denies every one of
    # them - cannot match. "it is not the recovery lock" does not contain "is
    # the recovery lock"; "it does not own the auto nonce lifecycle" does not
    # put "module owns" near "auto nonce lifecycle".
    for label, phrase in (
        ("the attempt result is its own durable result",
         "is its own durable result"),
        ("the token is durable recovery state", "the token is durable"),
        ("the next run reads the attempt result", "next run reads"),
        ("the next run reads it to reconcile", "reads to know it must reconcile"),
        ("Last Attempt Result is the recovery lock", "is the recovery lock"),
        ("the attempt result carries physical consumption",
         "carried by the attempt result"),
    ):
        assert phrase not in flat.lower(), (
            f"active reporter prose still teaches: {label}"
        )

    # PROXIMITY PAIRS, for claims a single phrase cannot pin.
    for label, left, right in (
        ("the reporter owns the AUTO nonce lifecycle",
         "module owns", "auto nonce lifecycle"),
        ("the attempt result is the durable recovery authority",
         "attempt result", "is the durable recovery"),
    ):
        near = _within(flat.lower(), left, right, 160)
        assert near is None, (
            f"active reporter prose still teaches: {label}\n...{near}..."
        )

    # POSITIVE ANCHORS. The prose must say, in substance, what is true.
    for required in ("modSimNonce",
                     "Pending AUTO Nonce",
                     "F21"):
        assert required in flat, required
    assert "durable recovery authority" in flat, (
        "the reporter prose never names where durable recovery authority lives"
    )
    assert "AUDIT" in flat or "audit" in flat
    assert "next run does not read it" in flat, (
        "the prose does not state that the next run ignores the attempt result"
    )
    assert "narrow scalar interface" in flat, (
        "the prose does not say the AUTO transaction is delegated"
    )


def test_44v_the_settlement_document_states_one_allocation_axis() -> None:
    """The settlement document is an ACTIVE authority, not an archive.

    Nothing read it before this round, so a sentence in it could contradict the
    source indefinitely - and one did: 9.4 still taught that a recovery action
    earns the fifth token, which is the conflated design 9.8 rejected. A reader
    deciding whether a change is safe reads this file, so a stale normative
    sentence here is a live defect.

    The distinction this control draws is STRUCTURAL, not editorial. History is
    supposed to survive: the document narrates the defect it fixed. What may not
    survive is a withdrawn claim standing as though it were still the rule, so
    active prose and blockquoted record are judged by different standards.
    """
    text = SETTLEMENT_MD.read_text(encoding="utf-8")

    section = _doc_section(text, "### 9.4 ")
    active = _normative(section)
    assert len(active) > 400, "9.4 lost its active prose"
    lower = active.lower()

    # 1. THE EXCLUSIVITY IS STATED, keyed on the FIELD the rule turns on. Plain
    #    "only" near "PERSISTENCE_INDETERMINATE" is not enough: the withdrawn
    #    paragraph happened to put those two within a sentence of each other for
    #    an unrelated reason, so the window must also name `allocationState`.
    exclusive = None
    start = 0
    while True:
        i = lower.find("only", start)
        if i < 0:
            break
        window = lower[i: i + 160]
        if "allocationstate" in window and "persistence_indeterminate" in window:
            exclusive = window
            break
        start = i + 1
    assert exclusive is not None, (
        "9.4 never says the fifth token is emitted only when this attempt's "
        "allocationState is PERSISTENCE_INDETERMINATE"
    )

    # 2. THE CONFLATION IS NOT STATED. Phrased affirmatively, so the settled
    #    wording - which denies it - cannot match.
    for withdrawn in ("unclassified states earn",
                      "both unclassified states",
                      "states earn it"):
        assert withdrawn not in lower, f"9.4 still teaches: {withdrawn}"

    # 3. EVERY SURVIVING ACTIVE MENTION OF THE RECOVERY ACTION DISCLAIMS IT.
    #    A paraphrase of the conflation still names RECOVERY_REQUIRED, so the
    #    check is not "is this exact sentence gone" but "does each mention say
    #    it is not an allocation classification".
    mentions = 0
    start = 0
    while True:
        i = lower.find("recovery_required", start)
        if i < 0:
            break
        mentions += 1
        window = lower[i: i + 140]
        assert any(marker in window for marker in
                   ("separate", "does not", "never", "not by itself")), (
            "9.4 mentions the recovery action without disclaiming it as an "
            f"allocation classification: ...{window}..."
        )
        start = i + 1
    assert mentions >= 1, "9.4 no longer distinguishes the two axes at all"

    # 4. AND THE WITHDRAWAL IS ON THE RECORD, not silently deleted. A reader who
    #    remembers the old rule has to be able to see that it was withdrawn.
    record = _quoted(section)
    assert "RECOVERY_REQUIRED" in record, (
        "9.4 drops the old rule instead of withdrawing it"
    )
    assert "withdrawn" in record.lower(), record[:200]

    # 5. THE ACTIVE PROSE NAMES THE DURABLE AUTHORITY.
    assert "F21" in active, "9.4 never names the durable recovery authority"

    # 6. THE SINGLE-AXIS TABLE IN 8.5 CARRIES ITS SUPERSESSION. It is accurate
    #    as history and wrong as a rule, so it may stand only under a note that
    #    points at the section which governs.
    table = _doc_section(text, "### 8.5 ")
    note = _quoted(table)
    assert "Superseded" in note, "the single-axis 8.5 table stands unqualified"
    assert "9.8" in note, note[:200]
    for required in ("RECOVERY_REQUIRED", "action"):
        assert required in note, required


def test_44t_the_sidecar_coordinate_has_exactly_one_authority() -> None:
    """One generated constant, one accessor, no second literal.

    A coordinate spelled independently in two production procedures is two
    authorities that agree only by luck, and the one that drifts is silent.
    """
    for name in (REPORT_BAS, NONCE_BAS):
        assert '"F21"' not in _module(name).code_without_string_removal, name.name
    users = [name for name in _module(NONCE_BAS).procedures
             if "SIM_PENDING_AUTO_NONCE_CELL" in _nonce(name)]
    assert users == ["PendingCell"], users
    assert "SIM_PENDING_AUTO_NONCE_CELL" not in _code(), (
        "the reporter addresses the recovery marker directly"
    )
    # AND EVERY SIDECAR TOUCH GOES THROUGH THAT ONE ACCESSOR.
    for name in _module(NONCE_BAS).procedures:
        if name == "PendingCell":
            continue
        assert "Sh(SIM_DATA_SHEET).Range(SIM_PENDING" not in _nonce(name), name


def test_44l_the_effective_seed_survives_every_post_derivation_failure() -> None:
    """A refusal that knows the nonce must know the seed that goes with it.

    `attempt_metadata_preserves` requires the effective seed in ALL THREE
    states, not only the consumed one. The previous shape copied it on the
    success arm alone, so every refusal after the seed was derived wrote the
    attempted nonce beside a default `0` seed - an attempt row naming an
    identity nobody could reconstruct.
    """
    body = _procedure("AllocateAutoNonce")
    lines = _statements(body)
    call = next(i for i, t in enumerate(lines) if "modSimNonce.SimNonceAllocate(" in t)
    ret = next(i for i, t in enumerate(lines) if t.startswith("AllocateAutoNonce ="))
    between = lines[call + 1: ret]
    # UNCONDITIONAL: no branch and no early exit stands between the call and the
    # copies, so there is no arm on which any of them can be skipped.
    for guard in ("Exit Function", "GoTo"):
        assert not any(guard in t for t in between), (
            f"{guard} sits between the call and the attempt-state copies: some "
            "failure arm reaches WriteAttemptBlock without them"
        )
    assert not any(re.match(r"^(If|ElseIf|Else|Select Case)\b", t) for t in between), (
        "the attempt-state copies are conditional on some arm"
    )
    copied = {t.split("=")[0].strip() for t in between if t.startswith("package.")}
    assert {"package.EffectiveSeed", "package.AutoIdentityKnown",
            "package.ConsumedNonce", "package.NonceState"} <= copied, sorted(copied)
    # AND THE OUT-PARAMETERS REALLY ARE ALL SET BEFORE ANY EXIT, which is what
    # makes an unconditional copy honest rather than a copy of stale locals.
    entry = _nonce("SimNonceAllocate")
    preamble = _until_first(entry, "On Error GoTo AllocationFailed")
    for out in ("effectiveSeed = 0", "autoNonce = 0", "identityKnown = False",
                "allocationState = SIM_NONCE_STATE_NOT_APPLICABLE",
                "recoveryRequired = False"):
        assert out in preamble, out
    # THE SEED IS NOT SUBSTITUTED, ZEROED, OR MADE CONDITIONAL ON CONSUMPTION.
    assert "package.EffectiveSeed = 0" not in body
    assert "package.EffectiveSeed = seed" in body
    assert body.index("package.EffectiveSeed") < body.index("package.NonceConsumed"), (
        "the seed copy sits after the consumption test, so it can be gated on it"
    )


def test_44m_the_after_nonce_failpoint_is_unreachable_in_fixed_mode() -> None:
    """`Phase6AfterNoncePersisted` names a persisted AUTO advance.

    FIXED reads no counter, writes no counter and persists no advance, so
    firing there would inject at a boundary that does not exist. Text order is
    not a containment proof: the branch has to LEAVE before the call.
    """
    assert _sim().raw["seeding"]["nonce_lifecycle"]["fixed_mode"][
        "executes_after_nonce_failpoint"] is False

    entry = _nonce("SimNonceAllocate")
    fixed = _branch(entry, "If hasSuppliedSeed Then")
    assert "FailPointCheck FAILPOINT_SIM_AFTER_NONCE" not in "\n".join(fixed), (
        "the failpoint sits inside the FIXED branch"
    )
    # THE BRANCH LEAVES. Without this, "not inside the branch" only means the
    # call is below the `End If` - which is precisely how FIXED reached it.
    assert fixed[-1] == "Exit Function", fixed
    assert not any(t.startswith("GoTo ") for t in fixed), fixed
    # FIXED TOUCHES NOTHING IN THE TRANSACTION.
    joined = "\n".join(fixed)
    for banned in ("SIM_IDENTITY_ROW_NEXT_AUTO_NONCE", "ResolveNextNonce",
                   "RunAllocationTransaction", "PendingCell", "ReadPending",
                   "ClearPending", "EstablishPending", "PersistAdvance"):
        assert banned not in joined, banned
    # AND THE AUTO PATH REACHES IT ONLY AFTER A CLEAN, CLEARED TRANSACTION.
    transaction, injected = _order_of("RunAllocationTransaction(",
                                      "FailPointCheck FAILPOINT_SIM_AFTER_NONCE",
                                      body=entry)
    assert transaction < injected
    guard = _statements(entry)
    call = next(i for i, t in enumerate(guard) if "RunAllocationTransaction(" in t)
    assert guard[call].startswith("If Not RunAllocationTransaction("), guard[call]
    arm = _branch(entry, "If Not RunAllocationTransaction(")
    assert any(t == "Exit Function" for t in arm), (
        "an unsuccessful transaction falls through to the failpoint and sampling"
    )


def test_44n_the_marker_is_established_before_the_counter_is_touched() -> None:
    """Write-ahead, and the counter untouched if the marker cannot be laid.

    Either physical outcome of the marker write is then safe: if it did not
    land the next AUTO run sees a blank sidecar and counter m; if it did, it
    sees pending m and counter m and resolves PRE_ALLOCATION. No nonce can
    vanish or replay because of uncertainty in THIS write.
    """
    cell = _sim().raw["sim_data"]["pending_auto_nonce"]
    assert cell["written_before_counter_persist"] is True
    assert cell["counter_persist_forbidden_until_established"] is True

    txn = _statements(_nonce("RunAllocationTransaction"))
    marked = next(i for i, t in enumerate(txn) if "EstablishPending(" in t)
    advanced = next(i for i, t in enumerate(txn) if "PersistAdvance(" in t)
    assert marked < advanced, txn
    assert txn[marked].startswith("If Not EstablishPending("), txn[marked]
    # THE GUARD REALLY LEAVES: a marker failure must not fall through to the
    # counter write.
    guarded = _branch("\n".join(txn), "If Not EstablishPending(")
    assert any(t == "Exit Function" for t in guarded), guarded
    assert not any("PersistAdvance" in t for t in guarded)
    # AND IT IS A DEFINITE OUTCOME, not an indeterminate one: nothing was
    # written to the counter, so the advance is KNOWN not to have persisted.
    assert any("SIM_NONCE_STATE_PRE_ALLOCATION" in t for t in guarded), guarded
    # ESTABLISHMENT IS WRITE PLUS VERIFY, not write alone.
    establish = _statements(_nonce("EstablishPending"))
    wrote = next(i for i, t in enumerate(establish) if "PendingCell.Value2 =" in t)
    verified = next(i for i, t in enumerate(establish) if "ReadPending(" in t)
    assert wrote < verified, establish
    assert any(t.startswith("On Error GoTo ") and not t.endswith(" 0")
               for t in establish[:wrote])
    # AND IT SAYS WHAT DID NOT HAPPEN, in words that survive a line continuation.
    text = _nonce("EstablishPending")
    assert text.count("counter was NOT touched and no nonce was consumed") == 3
    # THE COUNTER IS WRITTEN IN EXACTLY ONE PLACE, downstream of the marker.
    writers = [name for name in _module(NONCE_BAS).procedures
               if "SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 =" in _nonce(name)]
    assert writers == ["PersistAdvance"], writers


def test_44o_an_unrelated_later_attempt_cannot_destroy_the_recovery_state() -> None:
    """The decisive property, and the reason the sidecar exists.

    A FIXED run - successful or refused - legitimately rewrites the whole Last
    Attempt block. If that block carried the AUTO lock, the lock would be gone.
    The sidecar is in a different cell that no attempt writer can reach.
    """
    fixed = _sim().raw["seeding"]["nonce_lifecycle"]["fixed_mode"]
    assert fixed["may_proceed_while_pending_marker_exists"] is True
    assert fixed["may_overwrite_prior_auto_attempt_metadata"] is True
    assert fixed["writes_pending_auto_nonce"] is False
    assert fixed["clears_pending_auto_nonce"] is False
    assert _sim().raw["sim_data"]["pending_auto_nonce"]["survives_unrelated_attempts"] \
        is True

    # THE ATTEMPT WRITERS CANNOT NAME THE SIDECAR AT ALL. Both the SUCCESS path
    # and the unsuccessful path go through blocks the reporter builds, and none
    # of them addresses the pending cell.
    reporter = _module(REPORT_BAS).code
    assert "SIM_PENDING_AUTO_NONCE_CELL" not in reporter, (
        "the attempt writer can reach the recovery marker"
    )
    for writer in ("WriteAttemptBlock", "WriteStatusBlock", "BuildCommitBlock",
                   "BuildSnapshotBlock", "FinalCommit", "PublishCandidate"):
        assert "SIM_PENDING_AUTO_NONCE_CELL" not in _procedure(writer), writer
    # AND THE RANGES THEY DO WRITE EXCLUDE IT. The attempt range is a column-D
    # run of shared rows; the sidecar is in the bank-B column.
    ranges = _procedure("AttemptRange") + _procedure("StatusRange")
    assert "SIM_SHARED_VALUE_COLUMN" in ranges
    identity = _sim().raw["sim_data"]["run_identity"]
    assert _sim().raw["sim_data"]["pending_auto_nonce"]["column"] == \
        identity["bank_value_columns"]["B"]
    assert identity["bank_value_columns"]["B"] != identity["value_column"]


def test_44p_the_marker_is_cleared_only_on_a_definite_resolution() -> None:
    """m and m+1 clear it; an impossible or unreadable counter keeps it."""
    cell = _sim().raw["sim_data"]["pending_auto_nonce"]
    assert list(cell["cleared_on"]) == ["counter_equals_m", "counter_equals_m_plus_1"]
    assert list(cell["retained_on"]) == [
        "counter_is_neither", "counter_unreadable", "observation_unavailable"]

    # THE TWO DEFINITE ARMS CLEAR.
    classify = _nonce("Classify")
    consumed = classify[classify.index("If stored = nonce + 1 Then"):
                        classify.index("If stored = nonce Then")]
    pre = classify[classify.index("If stored = nonce Then"):
                   classify.index("allocationState = SIM_NONCE_STATE_INDETERMINATE")]
    assert "ClearPending(" in consumed, consumed
    assert "ClearPending(" in pre, pre
    # THE IMPOSSIBLE ARM DOES NOT, AND SAYS SO.
    definite = classify.index("If stored = nonce Then")
    tail = classify[definite:]
    closed = tail.index("\n    End If\n") + len("\n    End If\n")
    impossible = tail[closed:]
    assert "ClearPending" not in impossible, (
        "an impossible counter reading clears the marker that blocks reuse"
    )
    assert "marker is kept" in impossible
    # NEITHER DOES THE IRREDUCIBLE ONE.
    assert "ClearPending" not in _nonce("Indeterminate")
    assert "marker is retained" in _nonce("Indeterminate")
    # AND AN UNREADABLE COUNTER LEAVES IT STANDING TOO.
    resolve = _nonce("ResolveNextNonce")
    unreadable = _branch(resolve, "If Not ReadPersistedNonce(")
    assert not any("ClearPending" in t for t in unreadable), unreadable
    assert any(t == "recoveryRequired = True" for t in unreadable), unreadable
    assert any("is retained" in t for t in unreadable), unreadable
    assert any(t == "Exit Function" for t in unreadable), unreadable
    # THE NEXT-RUN ARMS AGREE WITH THE CONTRACT, ARM BY ARM.
    later = _sim().raw["seeding"]["nonce_lifecycle"]["next_run_reconciliation"]
    plus = _branch(resolve, "If counter = pending + 1 Then")
    same = _branch(resolve, "ElseIf counter = pending Then")
    assert later["counter_equals_m_plus_1"] == "CONSUMED"
    assert later["counter_equals_m"] == "PRE_ALLOCATION"
    for arm in (plus, same):
        assert any("ClearPending(" in t for t in arm), arm


def test_44q_a_cleanup_failure_stops_the_run_without_revising_the_observation() -> None:
    """The two axes, and the rule that keeps them apart.

    Clearing is a real COM write and a raised clear proves nothing, so THIS run
    must not sample while its own cleanup is unresolved. But by the time the
    clear is attempted the counter has already been OBSERVED, and a failure to
    clear a different cell cannot un-prove what that observation established.
    A cleanup failure therefore raises the recovery ACTION and leaves the
    physical classification exactly as it was found.
    """
    clear = _sim().raw["seeding"]["nonce_lifecycle"]["pending_clear"]
    assert clear["is_a_real_com_write"] is True
    assert clear["raised_clear_proves_marker_remains"] is False
    assert clear["unresolved_cleanup_permits_sampling"] is False
    assert clear["counter_rollback_on_clear_failure"] is False
    action = _sim().raw["seeding"]["nonce_lifecycle"]["recovery_action"]
    assert action["may_revise_allocation_state"] is False
    assert action["allocation_state_when_raised_by_cleanup"] == "unchanged"
    assert action["permits_sampling"] is False

    body = _statements(_nonce("ClearPending"))
    wrote = next(i for i, t in enumerate(body) if "PendingCell.ClearContents" in t)
    verified = next(i for i, t in enumerate(body) if "ReadPending(" in t)
    assert wrote < verified, body
    assert any(t.startswith("On Error GoTo ") and not t.endswith(" 0")
               for t in body[:wrote]), body
    # NO ARM RETURNS TRUE WITHOUT A VERIFIED-EMPTY READ-BACK.
    trues = [i for i, t in enumerate(body) if t == "ClearPending = True"]
    assert len(trues) == 1 and trues[0] > verified, body

    # THE RUN STOPS: the transaction guard leaves and never returns True.
    txn = _statements(_nonce("RunAllocationTransaction"))
    guard = next(t for t in txn if t.startswith("If Not ClearPending("))
    arm = _branch("\n".join(txn), guard)
    assert any(t == "Exit Function" for t in arm), arm
    assert txn[-2] == "RunAllocationTransaction = True", txn[-3:]

    # AND THE OBSERVATION SURVIVES. This is the whole control: every place a
    # clear failure is handled raises the ACTION and assigns NO allocation
    # state, so a proven CONSUMED stays CONSUMED and a proven PRE_ALLOCATION
    # stays PRE_ALLOCATION.
    assert arm == ["recoveryRequired = True", "Exit Function"], arm
    for owner in ("RunAllocationTransaction", "Classify"):
        for statement in _statements(_nonce(owner)):
            assert not statement.startswith("allocationState = SIM_NONCE_STATE_") or \
                "ClearPending" not in statement, statement
    classify = _nonce("Classify")
    for opener in ("If stored = nonce + 1 Then", "If stored = nonce Then"):
        block = _branch(classify, opener)
        assigned = [t for t in block if t.startswith("allocationState = ")]
        assert len(assigned) == 1, (opener, assigned)
        cleanup = _branch("\n".join(block), "If Not ClearPending(")
        assert cleanup == ["recoveryRequired = True", "Exit Function"], (opener, cleanup)
    # NOTHING ROLLS THE COUNTER BACK TO TIDY UP.
    assert "ClearContents" not in _nonce("PersistAdvance")
    for statement in _statements(_nonce_code()):
        if "SIM_IDENTITY_ROW_NEXT_AUTO_NONCE" in statement and ".Value2 =" in statement:
            assert "nonce + 1" in statement, statement


def test_44q2_the_recovery_action_is_not_an_allocation_state() -> None:
    """Two axes, and no expression that mixes them.

    `RECOVERY_REQUIRED` says the workbook must be reconciled before the next
    AUTO allocation. It says nothing about what physically happened to this
    attempt's counter, so it is not a member of `allocation_states` and no
    physical fact may be derived from it. Carrying it in the same scalar as the
    three classifications is what let a cleanup failure report a consumed nonce
    as unconsumed.
    """
    lifecycle = _sim().raw["seeding"]["nonce_lifecycle"]
    states = list(lifecycle["allocation_states"])
    action = lifecycle["recovery_action"]
    assert states == ["PRE_ALLOCATION", "CONSUMED", "PERSISTENCE_INDETERMINATE"]
    assert action["token"] == "RECOVERY_REQUIRED"
    assert action["token"] not in states, states
    assert action["is_an_allocation_state"] is False
    assert action["derives_physical_consumption"] is False
    assert action["carried_separately_from_allocation_state"] is True

    # THE INTERFACE CARRIES TWO SCALARS, and the recovery one is a Boolean, so
    # it cannot be compared against an allocation state even by accident.
    entry = _nonce("SimNonceAllocate")
    assert "ByRef allocationState As String" in entry, entry
    assert "ByRef recoveryRequired As Boolean" in entry, entry
    assert "ByRef state As String" not in _nonce_code(), (
        "the merged state parameter is back"
    )
    # NO CONSTANT NAMES THE ACTION ON THE ALLOCATION AXIS.
    declared = set(re.findall(r"Public Const (SIM_NONCE_STATE_\w+)",
                              _module(NONCE_BAS).raw))
    assert declared == {"SIM_NONCE_STATE_NOT_APPLICABLE", "SIM_NONCE_STATE_PRE_ALLOCATION",
                        "SIM_NONCE_STATE_CONSUMED", "SIM_NONCE_STATE_INDETERMINATE"}, \
        sorted(declared)
    assert "SIM_NONCE_STATE_RECOVERY" not in _module(NONCE_BAS).raw, (
        "the recovery action is declared as an allocation state again"
    )

    # AND THE REPORTER DERIVES CONSUMPTION FROM THE ALLOCATION AXIS ALONE.
    body = _procedure("AllocateAutoNonce")
    consumed = next(t for t in _statements(body)
                    if t.startswith("package.NonceConsumed ="))
    assert "SIM_NONCE_STATE_CONSUMED" in consumed, consumed
    assert "allocationState" in consumed, consumed
    for forbidden in ("recoveryRequired", "identityKnown", "allocated",
                      "SIM_NONCE_STATE_RECOVERY", "Or ", "And "):
        assert forbidden not in consumed, (forbidden, consumed)
    # THE TWO FACTS ARE STORED SEPARATELY, so neither can shadow the other.
    assert "package.NonceRecoveryRequired = recoveryRequired" in body
    assert "package.NonceState = allocationState" in body


def test_44q3_a_known_consumption_survives_every_later_failure() -> None:
    """The rule in one place: cleanup never erases an established fact.

    Once the counter has been observed at m+1 the nonce IS consumed. The run
    may still fail - and does - but `NonceConsumed` stays True through the
    clear, through the attempt write, and through anything else that follows.
    """
    classify = _nonce("Classify")
    consumed = _branch(classify, "If stored = nonce + 1 Then")
    # THE STATE IS SET FIRST, BEFORE ANY CLEANUP CAN FAIL.
    assigned = next(i for i, t in enumerate(consumed)
                    if t == "allocationState = SIM_NONCE_STATE_CONSUMED")
    cleanup = next(i for i, t in enumerate(consumed) if "ClearPending(" in t)
    assert assigned < cleanup, consumed
    # AND NOTHING AFTER IT REASSIGNS THE AXIS.
    assert not any(t.startswith("allocationState = ") for t in consumed[assigned + 1:]), \
        consumed
    # THE SAME FOR THE OTHER DEFINITE OBSERVATION.
    pre = _branch(classify, "If stored = nonce Then")
    assigned = next(i for i, t in enumerate(pre)
                    if t == "allocationState = SIM_NONCE_STATE_PRE_ALLOCATION")
    cleanup = next(i for i, t in enumerate(pre) if "ClearPending(" in t)
    assert assigned < cleanup, pre
    assert not any(t.startswith("allocationState = ") for t in pre[assigned + 1:]), pre
    # THE PUBLISHED RECORDS STILL REQUIRE THE STRONG STATE, so a consumed nonce
    # on an unsuccessful run is recorded for audit and published by nobody.
    for published in ("BuildSnapshotBlock", "BuildCommitBlock"):
        assert "package.NonceConsumed" in _procedure(published), published
    run = _procedure("RunSimulation")
    allocate = run.index("If Not AllocateAutoNonce(package, detail) Then")
    assert run.index("RunKernels") > allocate
    arm = _branch(run, "If Not AllocateAutoNonce(package, detail) Then")
    assert any("RecordRefusal" in t for t in arm), arm
    assert any(t == "Exit Function" for t in arm), arm


def test_44r_an_attempt_writer_failure_does_not_remove_the_recovery_authority() -> None:
    """The residual, as it now actually stands.

    It used to say that indeterminate persistence PLUS a failed attempt-row
    write lost the reconciliation authority. That was true only while the
    attempt row WAS the authority. The marker is established before the counter
    write and lives outside every block the reporter assembles, so a failed
    attempt write loses the audit line, not the safety property.
    """
    residual = _sim().raw["seeding"]["nonce_lifecycle"][
        "indeterminate_marker_storage_failure"]
    assert residual["reuse_prevention_depends_on_attempt_row"] is False
    assert residual["reuse_prevention_depends_on_pending_cell"] is True
    assert residual["attempt_row_failure_loses_audit_line_only"] is True
    assert residual["second_write_ahead_log_required"] is False

    # STRUCTURALLY: the marker is laid before the counter moves, and the module
    # that lays it never calls the attempt writer.
    assert "WriteAttemptBlock" not in _nonce_code(), (
        "the nonce module writes the attempt row"
    )
    assert "modSimReport" not in _nonce_code()
    txn = _statements(_nonce("RunAllocationTransaction"))
    assert next(i for i, t in enumerate(txn) if "EstablishPending(" in t) < \
        next(i for i, t in enumerate(txn) if "PersistAdvance(" in t)
    # AND THE OBSOLETE CLAIM IS NOT LEFT STANDING IN THE PROSE.
    for withdrawn in ("lost the reconciliation authority",
                      "has lost the reconciliation authority"):
        assert withdrawn not in _module(NONCE_BAS).code_without_string_removal, withdrawn


def test_44s_the_fifth_token_follows_the_allocation_axis_alone() -> None:
    """It records one thing: THIS attempt's transition could not be classified.

    Not "something needs reconciling". A recovery action layered on top of a
    known CONSUMED or PRE_ALLOCATION observation must not turn that known
    physical state into the indeterminate token, and a refusal taken while
    reconciling a PRIOR marker never began a transition to be indeterminate
    about. Both are ordinary unsuccessful attempts; F21 is what stops the next
    AUTO run, not this string.
    """
    token = _sim().raw["seeding"]["nonce_lifecycle"]["attempt_result_token"]
    assert token["token"] == "AUTO_NONCE_INDETERMINATE"
    assert token["audit_only"] is True
    assert token["is_the_durable_recovery_authority"] is False
    assert token["may_be_overwritten_by_a_later_attempt"] is True
    assert list(token["recorded_for"]) == ["PERSISTENCE_INDETERMINATE"]
    assert token["recorded_for_recovery_action"] is False
    assert token["other_unsuccessful_outcomes_record"] == "REFUSED"

    body = _procedure("RefusalResult")
    statements = _statements(body)
    # EXACTLY ONE TEST, AGAINST EXACTLY ONE ALLOCATION STATE.
    tests = [t for t in statements if "StrComp(" in t]
    assert len(tests) == 1, tests
    assert "SIM_NONCE_STATE_INDETERMINATE" in tests[0], tests[0]
    assert body.count("SIM_ATTEMPT_AUTO_NONCE_INDETERMINATE") == 1, body
    # THE RECOVERY AXIS IS NOT CONSULTED HERE AT ALL.
    for forbidden in ("NonceRecoveryRequired", "SIM_NONCE_STATE_RECOVERY",
                      "RECOVERY_REQUIRED"):
        assert forbidden not in body, forbidden
    # AND EVERYTHING ELSE FALLS THROUGH TO THE GENERIC RESULT.
    assert statements[-4:] == ["Else", "RefusalResult = SIM_ATTEMPT_REFUSED",
                               "End If", "End Function"], statements[-5:]

    # A PRIOR-MARKER REFUSAL LEAVES THE AXIS AT NOT_APPLICABLE, so it reaches
    # `RefusalResult` as an ordinary REFUSED - proved structurally, not by
    # reading the comment that says so.
    resolve = _nonce("ResolveNextNonce")
    assert not any(t.startswith("allocationState = ") for t in _statements(resolve)), (
        "reconciling a prior marker manufactures an allocation classification"
    )
    entry = _statements(_nonce("SimNonceAllocate"))
    default = next(i for i, t in enumerate(entry)
                   if t == "allocationState = SIM_NONCE_STATE_NOT_APPLICABLE")
    resolved = next(i for i, t in enumerate(entry) if "ResolveNextNonce(" in t)
    selected = next(i for i, t in enumerate(entry)
                    if t == "allocationState = SIM_NONCE_STATE_PRE_ALLOCATION")
    assert default < resolved < selected, entry
    assert entry[resolved].startswith("If Not ResolveNextNonce("), entry[resolved]
    assert "Exit Function" in entry[resolved], entry[resolved]

    # PHASE 5 IS UNTOUCHED. Its attempt axis stays four-valued; the fifth token
    # is a Phase-6 addition and appears in neither the Phase-5 contract nor its
    # loader.
    calc = (SPEC / "calc_contract.yaml").read_text(encoding="utf-8")
    assert "AUTO_NONCE_INDETERMINATE" not in calc
    loader = (PCCM_ROOT / "builder" / "pccm_builder" / "calc_loader.py").read_text(
        encoding="utf-8")
    assert "AUTO_NONCE_INDETERMINATE" not in loader
    assert list(_sim().raw["label_sets"]["attempt_result"]) == [
        "NONE", "SUCCESS", "REFUSED", "FAILED", "AUTO_NONCE_INDETERMINATE"]
    assert _sim().raw["sim_state"]["attempt_result_participates_in_derivation"] is False
    assert _sim().raw["sim_state"]["attempt_axis_is_orthogonal"] is True
    assert token["phase5_axis_unchanged"] is True

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
    assert nonce_handlers == {"0", "AllocationFailed", "MarkerFailed", "StepRaised",
                              "ObservationRaised", "ClearRaised", "ReadRaised",
                              "SharedReadRaised"}, sorted(nonce_handlers)
    # EVERY COM CALL IN THE TRANSACTION IS INSIDE ONE OF THEM. The marker write,
    # the counter advance and its verification read, the bounded observation,
    # the clear, and the sidecar read every caller depends on: a raise that
    # escapes into whichever handler happens to be armed further out is how the
    # first verification read used to skip its reconciliation.
    for owner, handler in (("SimNonceAllocate", "AllocationFailed"),
                           ("EstablishPending", "MarkerFailed"),
                           ("PersistAdvance", "StepRaised"),
                           ("Reconcile", "ObservationRaised"),
                           ("ClearPending", "ClearRaised"),
                           ("ReadPending", "ReadRaised"),
                           ("ReadShared", "SharedReadRaised")):
        body = _nonce(owner)
        assert f"On Error GoTo {handler}" in body, owner
        tail = body[body.index(f"{handler}:"):]
        assert "Err.Description" in tail, owner
        assert f"{owner} = True" not in tail, owner
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
