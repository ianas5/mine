#!/usr/bin/env python3
"""PCCM Phase 6 Step-12 INTEGRATION source review: the seams between modules.

--------------------------------------------------------------------------------
WHAT THIS LAYER IS FOR, AND WHAT IT IS NOT
--------------------------------------------------------------------------------
Steps 6 to 11 each proved ONE module against its own authority. Every one of
those suites still runs and nothing here duplicates them. What no per-step suite
could prove is the JOIN: that the value one module produces is the value the next
module consumes, that no second construction of a shared number exists anywhere
in the chain, and that the D6-11 scope holds across the whole repository rather
than inside a single file.

The accepted chain, which this reviews and does not redesign:

    CalcPrepareSimulationInputs -> modSimEngine -> modSimStats
        -> modSimFingerprint -> modSimReport -> inactive `_SimData` bank
        -> verified final D22:D30 active-bank commit

SOURCE CONFORMANCE, on Linux, now. NOTHING HERE EXECUTED VBA OR EXCEL. No claim
in this file may be read as "a simulation ran". Every runtime obligation is
enumerated in docs/phase6_step12.md and belongs to Step 13.

Runs standalone or under pytest.
"""

from __future__ import annotations

import hashlib
import re
import subprocess
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

from pccm_builder import load_sim_contract, load_structure_contract  # noqa: E402
from pccm_builder.vba_source import (  # noqa: E402
    VbaModule,
    contains_construct,
    load_modules,
    logical_statements,
    strip_comments,
)

SRC_VBA = PCCM_ROOT / "src" / "vba"
SPEC = PCCM_ROOT / "spec"
BUILD = PCCM_ROOT / "build"

REPORT = "modSimReport"
NONCE = "modSimNonce"
BRIDGE = "CalcPrepareSimulationInputs"

# The accepted Phase-6 public surface, settled by Step 11 and frozen here.
PHASE6_PUBLIC = (
    "PCCM_CurrentSimulationRequestFingerprint",
    "PCCM_RunSimulation",
    "PCCM_SimulationAttemptDetail",
    "PCCM_SimulationAttemptResult",
    "PCCM_SimulationRequestFingerprint",
    "PCCM_SimulationResultDigest",
    "PCCM_SimulationStatus",
)

# The accepted Phase-5 endpoints, settled long before Phase 6 existed.
PHASE5_ENDPOINTS = (
    "PCCM_Calculate",
    "PCCM_CalculationAttemptDetail",
    "PCCM_CalculationAttemptResult",
    "PCCM_CalculationFingerprint",
    "PCCM_CalculationStatus",
    "PCCM_CurrentInputFingerprint",
)

# Every hand-written module the accepted numerical/orchestration chain owns,
# with the hash the accepted work left it at. Named, never counted. Two entries
# have moved since Step 11 / Round 4A, both for a defect a Windows run found and
# both repointed in the same commit as the repair: modSimNonce and modSimReport
# for the Run-2 `IsWholeInRange` arity defect, and modSimReport again for the
# Run-4 `SameCell` blank-restore false negative.
FROZEN_SOURCE = {
    "modSimRng": "3d7c2cb365df03ccf73722f39b0c10e8964381e7cdd243732381dac7638257e3",
    "modSimSample": "5553198289bd98a7c84025868ac03c9f8ec95da3c01b23249c0da57d77901877",
    "modSimEngine": "f1283fe7d5d2ffcc5345dab9a00f68d3685b787563d104f50a886c5ed409abab",
    "modSimStats": "98bd21b227047d04e6847e554e027b339cf01dfb1112c1539a9e334966233be0",
    "modSimFingerprint": "9e6ad972fe59ead9e34c7d65b807dd0f2ca1cb1b29bfa71b377a4eb8f65cdfda",
    "modSimNonce": "b4e2d71ec2c73311f5f15c37d6ddffc35b06eae0b3335c4f99d937f08b28da00",
    "modSimReport": "55e383ee883e8470cb1ebfc7932c35c452dbc4860804e844051d7074aafc11d6",
    "modCalcFingerprint": "2efbb30c6f915c04b9c07adec07e25e11f4b5bd2b98e3efa818631dc510ce847",
    "modCalcReport": "8252b935b256b1abad9b26ca6b1d90c92c5e0d7566906308b191cd03dd6a71b3",
}

_CACHE: dict[str, object] = {}


# ---------------------------------------------------------------------------
# helpers - every one resolves its path at CALL time so the mutation controls
# can re-aim them at a damaged tree.
# ---------------------------------------------------------------------------
def _modules() -> dict[str, VbaModule]:
    return {module.name: module for module in load_modules([SRC_VBA])}


def _module(name: str) -> VbaModule:
    modules = _modules()
    assert name in modules, f"{name} is not in {SRC_VBA}"
    return modules[name]


def _procedure(module: str, name: str) -> str:
    code = _module(module).code_without_string_removal
    match = re.search(
        rf"^\s*(?:Public|Private)\s+(?:Function|Sub)\s+{re.escape(name)}\b", code, re.M)
    assert match, f"{name} is not declared in {module}"
    tail = code[match.start():]
    end = re.search(r"^\s*End\s+(?:Function|Sub)\s*$", tail, re.M)
    assert end, f"{name} has no End"
    return tail[: end.end()]


def _structure():
    return load_structure_contract(SPEC / "structure_contract.yaml")


def _sim():
    return load_sim_contract(SPEC / "sim_contract.yaml")


def _order(*tokens: str, body: str) -> list[int]:
    positions = []
    for token in tokens:
        assert token in body, f"missing: {token}"
        positions.append(body.index(token))
    return positions


# ===========================================================================
# A. The Phase-5 bridge is the ONLY seam into Phase 6
# ===========================================================================
def test_01_the_bridge_is_the_only_phase5_entry_into_phase6() -> None:
    """One door, and every Phase-6 module either uses it or does not know it."""
    # ONE BRIDGE, and now two callers of it. The claim was never "one caller":
    # it is that Phase 6 and later reach Phase 5 through this door and no other.
    # P7-4's orchestrator needs the resolved model to replay a driver, and
    # calling the accepted bridge is precisely how it avoids resolving one of
    # its own. Both callers are named, so a third cannot appear unremarked.
    callers = sorted(name for name, module in _modules().items()
                     if BRIDGE in module.code and name != "modCalcReport")
    assert callers == ["modSimPostReport", REPORT], callers
    # It is declared exactly once, in the accepted reporter.
    declarations = [name for name, module in _modules().items()
                    if re.search(rf"^(?:Public|Private) Function {BRIDGE}\b",
                                 module.code, re.M)]
    assert declarations == ["modCalcReport"], declarations
    # AND NO SECOND BRIDGE APPEARED. Nothing else in modCalcReport is reached
    # from Phase 6.
    report = _module(REPORT).code
    reached = set(re.findall(r"modCalcReport\.(\w+)", report))
    assert reached == {BRIDGE}, sorted(reached)


def test_02_the_bridge_reuses_the_accepted_preparation_and_gates_on_current() -> None:
    body = _procedure("modCalcReport", BRIDGE)
    assert "PrepareCurrentCalculation(package, detail)" in body, (
        "the bridge no longer reuses the accepted preparation"
    )
    assert "CALC_STATUS_CURRENT" in body, "the CURRENT gate is gone"
    prepare, status, gate = _order(
        "PrepareCurrentCalculation", "DeriveStatus", "CALC_STATUS_CURRENT", body=body)
    assert prepare < status < gate, body


def test_03_the_bridge_projects_and_never_rebuilds() -> None:
    body = _procedure("modCalcReport", BRIDGE)
    for rebuilt in ("BuildFactorTables", "BuildDriverFactors", "BuildAudits",
                    "BuildAnnual", "BuildFingerprint", "AccumulateTotals",
                    "ResolveModel", "Reconcile"):
        assert rebuilt not in body, f"the bridge recomputes {rebuilt}"
    # EVERY PROJECTED FIELD COMES OFF THE ONE PREPARED PACKAGE.
    for projected in ("drivers = package.Drivers",
                      "driverCount = package.Model.DriverCount",
                      "analyticalFingerprint = package.Fingerprint",
                      "deterministicBaseNominal = package.Totals.ANom",
                      "deterministicBasePv = package.Totals.APv"):
        assert projected in body, projected
    assert "AppliedTimelineText(package)" in body
    assert "HostDecimalSeparator()" in body


def test_04_the_bridge_writes_nothing_and_calls_no_endpoint() -> None:
    body = _procedure("modCalcReport", BRIDGE)
    assert ".Value2 =" not in body, "the bridge writes to the workbook"
    assert "ClearContents" not in body and "Calculate" not in body.replace(
        "PCCM_Calculate", "").replace("CalculationPackage", "").replace(
        "PrepareCurrentCalculation", "").replace(BRIDGE, ""), body
    for endpoint in ("PCCM_Calculate", "PCCM_CalculationStatus"):
        assert endpoint not in body, f"the bridge calls {endpoint}"
    # It is not itself an endpoint: no PCCM_ prefix, no button.
    assert not BRIDGE.startswith("PCCM_")
    structure = _structure()
    surface = set(structure.entry_points) | set(structure.api_procedures)
    assert BRIDGE not in surface, "the internal bridge became an automation endpoint"


# ===========================================================================
# B-D. Ownership: the reporter calls the kernels and implements none of them
# ===========================================================================
def test_05_the_reporter_owns_no_kernel_mathematics() -> None:
    code = _module(REPORT).code
    called = set(re.findall(
        r"mod(?:SimEngine|SimStats|SimFingerprint|SimRng|CalcReport)\.(\w+)", code))
    assert called == {
        "SimEngineRun", "SimStatsDescribe", "SimStatsContingency",
        "SimFpBuildRequestFingerprint", "SimFpResultDigest", BRIDGE,
    }, sorted(called)
    # The AUTO seed is derived inside the nonce transaction, by its owner.
    assert "SimRngAutoSeedFromNonce" in _module(NONCE).code
    # ENGINE. No contribution or distribution arithmetic.
    for engine_owned in ("SimSample", "BuildDriverFactors", "BuildInflationFactors",
                         "BuildDiscountFactors", "BuildKnom", "BuildKpv",
                         "SimRngNextUniform", "SimRngJumpNextStream",
                         "SimRngBuildComponentStreams", "Cheng", "MRG32k3a"):
        assert engine_owned not in code, engine_owned
    # STATISTICS. Reading `package.PvSummary.SampleStandardDeviation` is
    # CONSUMPTION - the field is what the owner handed back. What may not appear
    # is any procedure that could produce one.
    for stats_owned in ("SimStatsSortAscending", "SimStatsQuantileSorted",
                        "SimStatsUnitScale", "SimStatsProbabilityOf",
                        "SimStatsConstantValue", "Type7", "type_7",
                        "SafeSubtract", "n - 1"):
        assert stats_owned not in code, stats_owned
    # FINGERPRINT. No canonical encoding and no hash recurrence.
    for hash_owned in ("CalcFp", "FP_BASE", "FP_MOD_", "SimFpDigestRecord",
                       "SimFpRequestSuffix", "SimFpRetainedExtent"):
        assert hash_owned not in code, hash_owned
    # And no arithmetic operator that could express a measure at all.
    for operator in ("/", "*", "^"):
        assert operator not in code, operator


def test_06_the_contingency_is_never_computed_by_subtraction_here() -> None:
    body = _procedure(REPORT, "BuildContingencies")
    assert body.count("SimStatsContingency") == 2, "once per measure"
    assert "package.BaseNominal" in body and "package.BasePv" in body
    # The loop covers EVERY rung, not the selected one.
    assert "For index = 0 To SIM_QUANTILE_COUNT - 1" in body
    assert "SelectedConfidence" not in body
    # SUBTRACTION IS THE OWNER'S. The only procedure that fills either
    # contingency carrier is this one, and it fills it from the primitive's
    # out-parameter - there is no arithmetic between the base and the rung here.
    fillers = [name for name in _module(REPORT).procedures
               if re.search(r"package\.(?:Nominal|Pv)Contingency\(", _procedure(REPORT, name))
               and "=" in _procedure(REPORT, name)]
    assert set(fillers) <= {"BuildContingencies", "BuildContingencyBlock"}, fillers
    for _, statement in logical_statements(_module(REPORT).code):
        match = re.match(r"package\.(?:Nominal|Pv)Contingency\([^)]*\)\s*=\s*(.+)$",
                         statement)
        if match:
            assert match.group(1).strip() == "value", (
                f"a contingency element is computed rather than stored: {statement}"
            )
    assert "value, detail)" in body, "the primitive's out-parameter is not `value`"
    assert " - package.Base" not in body, "the base is subtracted here"


# ===========================================================================
# E. Request identity is the CURRENT analytical fingerprint
# ===========================================================================
def test_07_the_request_prefix_is_the_bridge_output_and_nothing_else() -> None:
    """The seam Step 11 could only half-prove: prefix in == prefix out."""
    for caller in ("RunKernels", "CurrentRequestFingerprint"):
        body = _procedure(REPORT, caller)
        call = body[body.index("SimFpBuildRequestFingerprint"):]
        call = call[: call.index(") Then")]
        # FIRST ARGUMENT is the field the bridge filled in. VBA continuations
        # split the argument list across lines, so they are folded away first.
        folded = re.sub(r"\s*_\s*\n\s*", " ", call)
        first = folded[folded.index("(") + 1:].split(",")[0].strip()
        assert first == "package.AnalyticalFingerprint", (caller, first)
        for forbidden in ("EffectiveSeed", "CandidateRunId", "Nonce",
                          "SelectedConfidence", "ActiveSnapshotText"):
            assert forbidden not in call, (caller, forbidden)
    # AND THAT FIELD IS ONLY EVER WRITTEN BY THE BRIDGE CALL.
    writers = []
    for _, statement in logical_statements(_module(REPORT).code):
        if re.match(r"package\.AnalyticalFingerprint\s*=", statement):
            writers.append(statement)
    assert writers == [], (
        "the analytical fingerprint is assigned outside the bridge call: "
        f"{writers}"
    )
    for caller in ("PrepareRun", "CurrentRequestFingerprint"):
        assert "package.AnalyticalFingerprint, detail" in _procedure(REPORT, caller) \
            or "package.AnalyticalFingerprint, _" in _procedure(REPORT, caller), caller


def test_08_no_stored_fingerprint_reaches_the_request_path() -> None:
    """A published snapshot is evidence of a past run, not a request."""
    for caller in ("PrepareRun", "RunKernels", "CurrentRequestFingerprint"):
        body = _procedure(REPORT, caller)
        assert "ActiveSnapshotText" not in body, (
            f"{caller} reads a published value into the request identity"
        )
        assert "SIM_IDENTITY_ROW_REQUEST_FINGERPRINT" not in body, caller
    # The stored fingerprint is read in exactly one place: the derivation that
    # compares it, and the accessor that reports it.
    readers = [name for name in _module(REPORT).procedures
               if "SIM_IDENTITY_ROW_REQUEST_FINGERPRINT" in _procedure(REPORT, name)]
    assert set(readers) <= {"DeriveSimStatus", "PCCM_SimulationRequestFingerprint",
                            "BuildSnapshotBlock"}, readers


# ===========================================================================
# F. Retained-array identity across stats, digest and publication
# ===========================================================================
def test_09_one_pair_of_arrays_feeds_statistics_digest_and_publication() -> None:
    kernels = _procedure(REPORT, "RunKernels")
    assert "SimEngineRun(package.Drivers, package.DriverCount" in kernels
    assert "package.TotalNominal, package.TotalPv, detail)" in kernels, (
        "the engine no longer fills the retained carriers directly"
    )
    assert "SimStatsDescribe(package.TotalNominal, package.Iterations" in kernels
    assert "SimStatsDescribe(package.TotalPv, package.Iterations" in kernels
    assert "SimFpResultDigest(package.TotalNominal, package.TotalPv" in kernels
    # PUBLICATION reads the same two fields.
    bank = _procedure(REPORT, "WriteIterationBank")
    assert "package.TotalNominal(LBound(package.TotalNominal)" in bank
    assert "package.TotalPv(LBound(package.TotalPv)" in bank
    # NOTHING REPLACES THEM. No assignment to either carrier anywhere.
    for _, statement in logical_statements(_module(REPORT).code):
        for carrier in ("package.TotalNominal", "package.TotalPv"):
            assert not re.match(rf"{re.escape(carrier)}\s*=", statement), statement
            assert not re.match(rf"ReDim\s+.*{re.escape(carrier)}", statement), statement


def test_10_nothing_sorts_or_reconstructs_a_retained_array() -> None:
    code = _module(REPORT).code
    for sorting in ("Sort", "sort", "Swap", "Reverse", "Shell", "Quick"):
        assert sorting not in code, f"the reporter reorders a retained array ({sorting})"
    # The engine is run exactly once, so there is no second population.
    assert _procedure(REPORT, "RunKernels").count("SimEngineRun") == 1
    assert code.count("SimEngineRun") == 1


# ===========================================================================
# G. Quantile provenance
# ===========================================================================
def test_11_every_published_rung_originates_in_describe() -> None:
    kernels = _procedure(REPORT, "RunKernels")
    describe, same, contingency, fingerprint, digest = _order(
        "SimStatsDescribe", "SameLadder", "BuildContingencies",
        "SimFpBuildRequestFingerprint", "SimFpResultDigest", body=kernels)
    assert describe < same < contingency < fingerprint < digest
    # The two ladders are proved to be one ladder before anything consumes them.
    same_body = _procedure(REPORT, "SameLadder")
    assert "SIM_QUANTILE_COUNT" in same_body and "vbBinaryCompare" in same_body
    # NO LADDER ELEMENT IS ASSIGNED ANYWHERE, qualified or not.
    ladder = re.compile(
        r"^(?:\w+\.)*(?:NominalLabels|PvLabels|NominalLadder|PvLadder)\s*\(")
    for _, statement in logical_statements(_module(REPORT).code):
        if ladder.match(statement) and "=" in statement:
            raise AssertionError(f"a ladder element is assigned: {statement}")
    # And the labels come from the owner, not from a second list here.
    assert "SIM_QUANTILE_LABEL" not in _module(REPORT).code_without_string_removal or \
        "modSimStats" in _module(REPORT).code


# ===========================================================================
# H. Dual-bank publication
# ===========================================================================
def test_12_the_candidate_never_touches_the_active_bank() -> None:
    run = _procedure(REPORT, "RunSimulation")
    assert "package.TargetBank = InactiveBank(package.ActiveBank)" in run
    for stage in ("PublishCandidate", "WriteIterationBank", "BuildSnapshotBlock",
                  "BuildSummaryBlock", "BuildContingencyBlock"):
        body = _procedure(REPORT, stage)
        assert "package.ActiveBank" not in body, f"{stage} names the active bank"
        assert "SIM_IDENTITY_ROW_ACTIVE_BANK" not in body, stage
    # THE SELECTOR HAS EXACTLY ONE NAMER, AND IT READS.
    owners = [name for name in _module(REPORT).procedures
              if "SIM_IDENTITY_ROW_ACTIVE_BANK" in _procedure(REPORT, name)]
    assert owners == ["ReadActiveBank"], owners
    # AND EXACTLY ONE PROCEDURE WRITES THE COMMIT RANGE.
    writers = [name for name in _module(REPORT).procedures
               if "Range(SIM_FINAL_COMMIT_RANGE).Value2 =" in _procedure(REPORT, name)]
    assert writers == ["FinalCommit"], writers
    commit = _procedure(REPORT, "FinalCommit")
    assert commit.count(".Value2 =") == 2, "the commit procedure writes something else"
    assert "built(9, 1) = package.TargetBank" in _procedure(REPORT, "BuildCommitBlock")


def test_13_the_candidate_is_verified_before_the_commit() -> None:
    run = _procedure(REPORT, "RunSimulation")
    publish, commit = _order("PublishCandidate", "FinalCommit", body=run)
    assert publish < commit
    assert "VerifyCandidateBank" in _procedure(REPORT, "PublishCandidate")
    assert "SameBlock(SIM_FINAL_COMMIT_RANGE, block, 9, 1)" in _procedure(
        REPORT, "FinalCommit")


# ===========================================================================
# I-J. The reporting boundary
# ===========================================================================
def test_14_no_vba_module_writes_or_resolves_results() -> None:
    for name, module in _modules().items():
        text = module.code_without_string_removal
        for banned in ('"Results"', "shResults", "Results!"):
            assert banned not in text, f"{name} names {banned}"
    sheets = set(re.findall(r"modWorkbook\.Sh\((\w+)\)", _module(REPORT).code))
    assert sheets == {"SIM_DATA_SHEET"}, sheets
    # Results is Stage-A presentation over `_SimData`, and the contract says so.
    presentation = _sim().raw["results_minimum"]["presentation"]
    assert presentation["written_by_the_run"] is False
    assert presentation["materialised_by_stage_a"] is True
    assert presentation["computes_statistics"] is False
    assert presentation["recomputes_quantiles"] is False
    assert presentation["contingency_by_subtraction_on_results"] is False


def test_15_selected_confidence_level_is_reporting_only() -> None:
    for name, module in _modules().items():
        text = module.code_without_string_removal
        for banned in ("inpSelectedConfidenceLevel", "SelectedConfidence",
                       "NM_INPUT_SELECTED_CONFIDENCE_LEVEL", "SelectedPx"):
            assert banned not in text, f"{name} reads the reporting selector"
    selector = _sim().raw["selected_confidence_level"]
    assert selector["participates_in_execution_validity"] is False
    assert selector["participates_in_request_fingerprint"] is False
    # It is absent from the run, the status derivation and AUTO allocation.
    for stage in ("RunSimulation", "PrepareRun", "AllocateAutoNonce", "RunKernels",
                  "DeriveSimStatus", "CurrentRequestFingerprint"):
        assert "Selected" not in _procedure(REPORT, stage), stage


# ===========================================================================
# K. The attempt axis and the status axis are orthogonal
# ===========================================================================
def test_16_the_attempt_result_cannot_decide_the_status() -> None:
    derive = _procedure(REPORT, "DeriveSimStatus")
    for banned in ("SIM_ATTEMPT_", "LAST_ATTEMPT", "SelectedConfidence"):
        assert banned not in derive, banned
    # The three states plus the absence-of-publication case.
    assert "SIM_STATE_CURRENT" in derive and "SIM_STATE_STALE" in derive
    assert "SIM_STATE_INVALID" in derive
    assert "If Len(active) = 0 Then Exit Function" in derive, (
        "a blank selector became a fourth state instead of an absence"
    )
    # And the attempt writer derives rather than inherits.
    attempt = _procedure(REPORT, "WriteAttemptBlock")
    assert "DeriveSimStatus()" in attempt
    assert "SIM_ATTEMPT_SUCCESS" not in attempt


# ===========================================================================
# L. The exact public surfaces, both phases
# ===========================================================================
def test_17_the_two_public_surfaces_are_exactly_the_accepted_ones() -> None:
    modules = _modules()
    assert tuple(sorted(modules[REPORT].public_procedures)) == PHASE6_PUBLIC, (
        sorted(modules[REPORT].public_procedures)
    )
    found = {name for module in modules.values()
             for name in module.public_procedures if name.startswith("PCCM_")}
    # THE PHASE-4 SURFACE IS THE CONTRACT'S, not a second list kept here: the
    # buttons plus the harness procedures, both declared. Deriving it means a
    # new endpoint cannot be absorbed by editing this file.
    import yaml

    structure = _structure()
    declared = yaml.safe_load(
        (SPEC / "structure_contract.yaml").read_text(encoding="utf-8"))["vba"]
    phase4 = set(declared["entry_points"]) | set(declared["harness_procedures"])
    assert set(structure.api_procedures) == set(PHASE5_ENDPOINTS), sorted(
        structure.api_procedures)
    # PHASE 7 has its own contract list, deliberately not `api_procedures`: the
    # Phase-5 Gate-B controls require every entry there to carry Windows
    # evidence, and a Phase-7 endpoint has none and must not appear to.
    phase7 = set(declared["phase7_api_procedures"])
    assert phase7 == {"PCCM_RunSensitivity"}, sorted(phase7)
    assert found == phase4 | set(PHASE5_ENDPOINTS) | set(PHASE6_PUBLIC) | phase7, sorted(
        found ^ (phase4 | set(PHASE5_ENDPOINTS) | set(PHASE6_PUBLIC) | phase7))
    assert not (phase4 & set(PHASE6_PUBLIC)), "a Phase-6 name entered the Phase-4 surface"
    # The reporter that owns Phase 5 gained exactly one non-endpoint Public name.
    extra = set(modules["modCalcReport"].public_procedures) - set(PHASE5_ENDPOINTS)
    assert extra == {BRIDGE}, sorted(extra)
    # NO PHASE-6 BUTTON.
    for endpoint in PHASE6_PUBLIC:
        assert endpoint not in set(structure.entry_points), endpoint


def test_18_no_invented_phase6_accessor_exists() -> None:
    code = "\n".join(module.code for module in _modules().values())
    for invented in ("PCCM_SimulationRunId", "PCCM_SimulationEffectiveSeed",
                     "PCCM_SimulationIterations", "PCCM_SimulationSeed",
                     "PCCM_SimulationLadder", "PCCM_SimulationContingency",
                     "PCCM_Dashboard"):
        assert invented not in code, invented
    # PCCM_RunSensitivity LEFT THIS LIST when P7-4 landed it, exactly as each
    # name leaves at the step that implements it. It is not invented now - it
    # is declared, owned by one module, and asserted to be there.
    owners = sorted(name for name, module in _modules().items()
                    if "PCCM_RunSensitivity" in module.public_procedures)
    assert owners == ["modSimPostReport"], owners


# ===========================================================================
# 4. D6-11, repo-wide, from the structured manifest authority
# ===========================================================================
def test_19_the_scoped_constructs_live_only_in_their_owners() -> None:
    structure = _structure()
    modules = _modules()
    scoped = [(rule.construct, tuple(rule.allowed_in))
              for rule in structure.forbidden_construct_rules if rule.is_scoped]
    assert scoped == [("MRG32k3a", ("modSimRng",)),
                      ("RunSimulation", (REPORT,))], scoped
    for construct, owners in scoped:
        assert len(owners) == 1, construct
        owner = owners[0]
        assert "*" not in owners, construct
        assert contains_construct([modules[owner]], construct), (
            f"the grant is vacuous: {owner} does not contain {construct}"
        )
        others = [module for name, module in modules.items() if name != owner]
        assert not contains_construct(others, construct), construct
        rule = next(r for r in structure.forbidden_construct_rules
                    if r.construct == construct)
        for name in modules:
            assert rule.forbidden_in(name) == (name != owner), (construct, name)


def test_20_percentile_is_in_no_executable_module() -> None:
    modules = list(_modules().values())
    rule = next(r for r in _structure().forbidden_construct_rules
                if r.construct == "Percentile")
    assert not rule.is_scoped, "Percentile was granted an owner"
    assert tuple(rule.allowed_in) == (), rule.allowed_in
    assert not contains_construct(modules, "Percentile")
    for name in _modules():
        assert rule.forbidden_in(name) is True, name


def test_21_the_global_prohibitions_are_still_global() -> None:
    structure = _structure()
    modules = list(_modules().values())
    for construct in ("Rnd(", "Randomize", "NPV", "Worksheet_Change",
                      "Workbook_SheetChange", "FinalReleaseComObject"):
        rule = next(r for r in structure.forbidden_construct_rules
                    if r.construct == construct)
        assert not rule.is_scoped, f"{construct} was scoped"
        assert not contains_construct(modules, construct), construct


def test_22_the_manifest_is_the_module_inventory_authority() -> None:
    """No count is hardcoded anywhere: the declared set IS the authority."""
    import json

    manifest = json.loads((BUILD / "stage_b_manifest.json").read_text(encoding="utf-8"))
    declared = [entry["name"] for entry in manifest["vba"]["modules"]]
    assert len(declared) == len(set(declared)), declared
    structure_names = [module.name for module in _structure().vba_modules]
    assert declared == structure_names, (declared, structure_names)
    # Every hand-written declared module has a file, and every file is declared.
    handwritten = {entry["name"] for entry in manifest["vba"]["modules"]
                   if entry.get("generated") is False}
    on_disk = {path.stem for path in SRC_VBA.glob("*.bas")}
    assert handwritten == on_disk, (sorted(handwritten ^ on_disk))
    # The structured rules survive into the manifest, unflattened.
    rules = {entry["construct"]: entry["allowed_in"]
             for entry in manifest["vba"]["forbidden_construct_rules"]}
    assert rules["MRG32k3a"] == ["modSimRng"]
    assert rules["RunSimulation"] == [REPORT]
    assert rules["Percentile"] == []
    assert set(rules) == set(manifest["vba"]["forbidden_constructs"]), (
        "a flattened entry has no structured rule behind it"
    )


# ===========================================================================
# 2/16. The accepted source is frozen
# ===========================================================================
# THE MODULES A LATER PHASE HAS REOPENED, and why each one was.
#
# The freeze records that Run 6 executed exactly these bytes. That is a
# HISTORICAL fact and it stays true; what cannot stay true forever is "and no
# byte will ever change again", because a later phase changing production code
# under its own authority is the normal way this project moves. Read against the
# current tree the freeze said the second thing.
#
# So a reopened module is verified where the fact actually lives - at the pinned
# Step-13 closure commit - and its current bytes are then asserted to DIFFER,
# which is the honest consequence: those bytes were never executed on Windows
# and are not runtime-proven until a Phase-7 run says so.
STEP13_CLOSURE_COMMIT = "85778b2854fee431a845499e5a2fe37f40e96610"
REOPENED_SINCE_CLOSURE = {
    "modSimEngine": "P7-3 extracted the shared per-driver contribution routine "
                    "and added deterministic per-driver replay; P7-5 added "
                    "annual block replay, which applies a supplied per-year "
                    "factor to the same sampled observation",
    "modCalcReport": "P7-5 copies the RESOLVED per-year inputs - FxRate, "
                     "Weights, Inflation - into DriverFactors at the one site "
                     "that already had all three in hand, so the annual layer "
                     "can regroup them. It resolves nothing new.",
    "modSimStats": "P7-5 exposed the type-7 ORDER-STATISTIC POSITION - which "
                   "source ordinals a percentile was interpolated between - and "
                   "extracted the h/lo/hi/f arithmetic into one owner shared by "
                   "the value and the position. No percentile value changed.",
}

# AND A REOPENED MODULE IS FROZEN AGAIN, to the bytes the phase that reopened it
# left behind. A module is always frozen to SOMETHING; what a later phase
# changes is which commit those bytes come from, never whether they are pinned.
# Dropping the pin instead would leave the one module under active change as the
# only one a stray byte could move unnoticed.
REOPENED_CURRENT = {
    "modSimEngine": "9c307eae451252a983fc2c36205759335e14642557c3c4ecf9ab4ee30ec3237e",
    "modSimStats": "0c75c0902980ced0a7ad4d59b985f07b7bc9978a0723f2d45fb98dff0da7c7c8",
    "modCalcReport": "c9f728f06bc7bc89eff5eb6e389d9fa305083e82af89b9e34f50340499364671",
}


def _blob_at_closure(name: str) -> str:
    blob = subprocess.run(
        ["git", "show", f"{STEP13_CLOSURE_COMMIT}:pccm/src/vba/{name}.bas"],
        cwd=PCCM_ROOT.parent, check=True, stdout=subprocess.PIPE).stdout
    return hashlib.sha256(blob).hexdigest()


def test_23_every_accepted_module_is_byte_identical() -> None:
    """Unchanged modules at HEAD; reopened ones at the commit that froze them."""
    assert set(REOPENED_CURRENT) == set(REOPENED_SINCE_CLOSURE), (
        "a reopened module has no current digest, so nothing pins its bytes")
    for name, expected in FROZEN_SOURCE.items():
        # A REOPENED MODULE IS STILL PINNED, to the digest of the phase that
        # reopened it rather than to the Run-6 one.
        expected = REOPENED_CURRENT.get(name, expected)
        path = SRC_VBA / f"{name}.bas"
        actual = hashlib.sha256(path.read_bytes()).hexdigest()
        assert actual == expected, f"{name}.bas moved: {actual}"


def test_23a_a_reopened_module_still_matches_the_bytes_run_6_executed() -> None:
    """The historical half, proved against the pinned commit rather than assumed.

    If this fails, the freeze digest was never the Run-6 digest and every claim
    resting on it is suspect - which is a different and much worse problem than
    a module having legitimately moved since.
    """
    assert REOPENED_SINCE_CLOSURE, "nothing is reopened; this control has no subject"
    for name, reason in REOPENED_SINCE_CLOSURE.items():
        assert name in FROZEN_SOURCE, name
        assert len(reason) > 40, f"{name} was reopened without a written reason"
        assert _blob_at_closure(name) == FROZEN_SOURCE[name], (
            f"{name}.bas at the Step-13 closure commit is not the frozen digest")


def test_23b_a_reopened_module_is_not_claimed_as_runtime_proven() -> None:
    """And the consequence is recorded, not left implicit.

    Run 6 ran the closure bytes. A module that has moved since carries bytes
    Windows has never executed, so listing it here is the statement that its
    runtime evidence is OUTSTANDING - not a note that it was edited.
    """
    for name in REOPENED_SINCE_CLOSURE:
        current = hashlib.sha256((SRC_VBA / f"{name}.bas").read_bytes()).hexdigest()
        assert current == REOPENED_CURRENT[name], f"{name}.bas moved: {current}"
        assert current != FROZEN_SOURCE[name], (
            f"{name}.bas is byte-identical to the frozen Run-6 source, so it was "
            "not reopened at all and must come off the reopened list")


# THE RUN-6 PROJECTION, and the current one. They are different artefacts now.
#
# `modSimContract` is GENERATED from `sim_contract.yaml`, so contracting the
# Phase-7 sensitivity block changed it. The identity Run 6 executed is a
# historical fact and is unchanged; the projection this tree builds is a
# different file that no Windows run has executed.
#
# AND IT MOVED AGAIN AFTER THE FIRST WINDOWS EVIDENCE. The P7-4 timing run of
# 0734a38 read the contracted sensitivity stamp and got summary-statistics
# labels: P7-1 had allocated the block at J-Q and S-Z, over two accepted
# Phase-6 blocks. The block is now at CC-CJ and CL-CS, so the projection of the
# corrected contract is a third artefact. 11e58482... was the colliding layout
# and must never be built again.
RUN6_GENERATED_IDENTITY = "daa4d27889c30eadb2ab892bcfa4e6f6bab8a137aae79a01a8d8f1e8e1c215ac"
PHASE7_GENERATED_IDENTITY = "2f4c6e4ad27d52aac67259f41b977818c858d7ba3082d50114ec45da05b55233"


def test_24_the_generated_authority_is_byte_identical() -> None:
    """Pinned, and pinned to the right thing.

    A generated artefact is always frozen to SOMETHING; what a later phase
    changes is which build those bytes come from. Both digests are asserted, so
    a stray change to either is still caught.
    """
    for path, expected in (
        (BUILD / "vba" / "modSimContract.bas", PHASE7_GENERATED_IDENTITY),
        (BUILD / "phase6_cases.json",
         "8019683a0490fcf0740cf07244524973d9b7470c933f1003059025b6b019a0be"),
    ):
        actual = hashlib.sha256(path.read_bytes()).hexdigest()
        assert actual == expected, f"{path.name} moved: {actual}"


def test_24a_the_run_6_projection_identity_is_still_what_the_harness_pins() -> None:
    """The historical half, read from the artefact that actually carries it.

    The Phase-6 Gate-B harness pins the identity Run 6 executed, and that file
    is accepted Run-6 evidence which P7-4 does not touch. Reading the pin from
    the harness rather than restating it here is what keeps this a statement
    about the run instead of a literal agreeing with another literal.
    """
    harness = (PCCM_ROOT / "bootstrap" / "windows" /
               "phase6_gate_b_scenarios.ps1").read_text(encoding="utf-8")
    assert f"'{RUN6_GENERATED_IDENTITY}'" in harness, (
        "the Gate-B harness no longer pins the Run-6 projection identity")


def test_24b_the_current_projection_is_not_the_one_run_6_executed() -> None:
    """AND THE CONSEQUENCE IS RECORDED, not left to be discovered on Windows.

    Contracting the Phase-7 sensitivity block regenerated `modSimContract`, so
    the projection this tree builds is not the module Run 6 ran. The Phase-6
    Gate-B harness pins the Run-6 identity and would REFUSE this one - which is
    the harness being right, not a defect in it. Phase-7 Windows acceptance has
    to settle which identity its own run checks against; nothing here may be
    read as that settlement having happened.
    """
    assert PHASE7_GENERATED_IDENTITY != RUN6_GENERATED_IDENTITY
    actual = hashlib.sha256((BUILD / "vba" / "modSimContract.bas").read_bytes()).hexdigest()
    assert actual == PHASE7_GENERATED_IDENTITY
    assert actual != RUN6_GENERATED_IDENTITY, (
        "the projection matches the Run-6 identity again; if that is intended, "
        "this control and the Phase-7 runtime note must both be revisited")


def test_25_the_accepted_reporter_prefix_is_still_byte_identical() -> None:
    """The Step-11 technique, re-proved at the integration layer.

    THE PREFIX DIGEST MOVED IN P7-5, and it is still a digest. P7-5 authorised
    the smallest extension of the resolution handoff, and BuildDriverFactors -
    which sits before this banner - now copies the resolved FxRate, Weights and
    Inflation into DriverFactors. What the control asserts is unchanged: the
    text before the bridge is pinned, and the bridge is still the only thing
    after it. Only which commit those bytes come from has moved.
    """
    banner = ("' ==========================================================================\n"
              "' STEP 11 ADDITION - THE PHASE-6 PREPARATION BRIDGE\n")
    text = (SRC_VBA / "modCalcReport.bas").read_text(encoding="utf-8")
    assert text.count(banner) == 1
    accepted = text[: text.index(banner)]
    assert hashlib.sha256(accepted.encode("utf-8")).hexdigest() == (
        "8d67d3f18b1ea8c4a8baba478f025d486f71afaa1ac31beac88d7b7ecfff80a9")
    after = re.findall(r"^(?:Public|Private) (?:Function|Sub) (\w+)",
                       text[text.index(banner):], re.M)
    assert after == [BRIDGE], after


# ===========================================================================
# The cross-module failure-path guarantee
# ===========================================================================
def test_28_every_transaction_failure_is_routed_to_the_attempt_recorder() -> None:
    """Every transaction-stage failure REACHES the attempt-recording path.

    The contract's `refusal_or_failure_after_auto_allocation` requires
    `attempt_metadata_updated: true`, and that is a statement about REAL COM
    failures, not only about a helper returning False. Each stage that touches
    the worksheet therefore carries its own scoped handler, and RunSimulation
    routes its False through RecordRefusal or RecordFailure.

    WHAT THIS DOES **NOT** CLAIM. It says a failure is ROUTED to the attempt
    writer. It does not say the attempt record can never fail to be STORED:
    `WriteAttemptBlock` ends in a single unguarded COM write, and if that write
    raises, the run leaves through the invocation axis with no attempt row. That
    is a distinct storage failure of the audit writer, it is deliberately not
    settled here - the contract requires no recovery state for it, and inventing
    one would be scope creep - and it is enumerated for Gate B instead. The
    distinction is the point: routing is what source can guarantee, storage is
    not.
    """
    run = _procedure(REPORT, "RunSimulation")
    # THE NONCE MODULE OWNS NO ENDPOINT AND WRITES NO ATTEMPT ROW.
    assert not [n for n in _module(NONCE).public_procedures if n.startswith("PCCM_")]
    assert "AttemptRange" not in _module(NONCE).code
    for banned in ("SIM_IDENTITY_ROW_LAST_ATTEMPT_DETAIL", "WriteAttemptBlock",
                   "SimEngineRun", "SimStatsDescribe", "SimFp", "SnapshotRange"):
        assert banned not in _module(NONCE).code, banned
    # Every staged call is tested, and every arm records something.
    for stage, recorder in (("PrepareRun", "RecordRefusal"),
                            ("AllocateAutoNonce", "RecordRefusal"),
                            ("RunKernels", "RecordRefusal"),
                            ("PublishCandidate", "RecordFailure"),
                            ("FinalCommit", "RecordFailure")):
        guard = f"If Not {stage}(package, detail) Then"
        assert guard in run, guard
        arm = run[run.index(guard):]
        arm = arm[: arm.index("End If")]
        assert recorder in arm, (stage, recorder)
        assert "Err.Raise" not in arm, (
            f"{stage} re-raises instead of recording an attempt"
        )
    # EVERY COM-FALLIBLE STAGE CARRIES ITS OWN ENVELOPE, in its owning module.
    assert "On Error GoTo AllocationFailed" in _procedure(NONCE, "SimNonceAllocate")
    assert "On Error GoTo MarkerFailed" in _procedure(NONCE, "EstablishPending")
    assert "On Error GoTo StepRaised" in _procedure(NONCE, "PersistAdvance")
    assert "On Error GoTo ObservationRaised" in _procedure(NONCE, "Reconcile")
    assert "On Error GoTo ClearRaised" in _procedure(NONCE, "ClearPending")
    assert "On Error GoTo ReadRaised" in _procedure(NONCE, "ReadPending")
    assert "On Error GoTo CandidateFailed" in _procedure(REPORT, "PublishCandidate")
    assert "On Error GoTo CommitFailed" in _procedure(REPORT, "FinalCommit")
    assert "On Error GoTo CaptureFailed" in _procedure(REPORT, "FinalCommit")
    assert "On Error GoTo RestoreFailed" in _procedure(REPORT, "FinalCommit")

    # AND NO RAISING CALL SITS OUTSIDE ONE. `FailPointCheck` RAISES by design -
    # `Err.Raise vbObjectError + 5001` in modAppState - so a naked call anywhere
    # in the orchestration leaves through the invocation handler and writes no
    # attempt record. That is the defect this guarantee exists to refuse, and
    # walking the `If Not <stage>` guards alone never saw it.
    check = _procedure("modAppState", "FailPointCheck")
    assert "Err.Raise" in check, (
        "FailPointCheck no longer raises, so this guarantee proves nothing"
    )
    assert "FailPointCheck" not in run, (
        "a failpoint raises straight out of RunSimulation"
    )
    for owner in (REPORT, NONCE):
        for name in _module(owner).procedures:
            body = _procedure(owner, name)
            if "FailPointCheck" not in body:
                continue
            armed = [m.start() for m in re.finditer(r"On Error GoTo (?!0\b)\w+", body)]
            assert armed, f"{owner}.{name} fires a failpoint with no handler armed"
            assert min(armed) < body.index("FailPointCheck"), f"{owner}.{name}"

    # THE LIMIT OF THE GUARANTEE, ASSERTED SO IT CANNOT DRIFT INTO A STRONGER
    # CLAIM. The audit writer's own COM write is unguarded, by design.
    attempt = _procedure(REPORT, "WriteAttemptBlock")
    assert "SimSheet.Range(AttemptRange()).Value2 = block" in attempt
    assert "On Error" not in attempt, (
        "WriteAttemptBlock grew an error handler; either the storage-failure "
        "case is now settled - in which case this guarantee must say so and be "
        "tested - or a blanket suppressor was introduced"
    )
    # AND NO PHASE-6 MODULE SUPPRESSES ERRORS WHOLESALE. (The Phase-4 modules
    # carry their own documented `On Error Resume Next` whitelist, policed by
    # test_phase4_stage_b_source.py; nothing here widens or narrows it.)
    for name in tuple(FROZEN_SOURCE) + (REPORT, NONCE):
        assert "On Error Resume Next" not in _module(name).code, name


def test_29_the_publication_contract_is_implemented_where_it_is_stated() -> None:
    """Each accepted failure-semantics clause maps to a source guarantee."""
    semantics = _sim().raw["publication"]["failure_semantics"]
    after = semantics["refusal_or_failure_after_auto_allocation"]
    assert after["next_auto_nonce_advanced"] is True
    assert after["active_bank_changed"] is False
    assert after["attempt_metadata_updated"] is True
    # NOTHING DECREMENTS THE COUNTER on any failure path, and the advance is
    # persisted AND verified AND marked before the accepted injection boundary,
    # so `next_auto_nonce_advanced: true` holds even for an injected failure.
    code = _module(NONCE).code
    assert "NEXT_AUTO_NONCE" in code
    for _, statement in logical_statements(code):
        if "SIM_IDENTITY_ROW_NEXT_AUTO_NONCE" in statement and "- 1" in statement:
            raise AssertionError(f"the consumed nonce is rolled back: {statement}")
    # THE CONTRACT'S ORDER, at its owner: read < derive < mark < persist <
    # clear < sample. The marker goes down BEFORE the counter is touched, so
    # `next_auto_nonce_advanced: true` is never claimed for a counter that was
    # written without a recoverable record of which nonce it was written for.
    entry = _procedure(NONCE, "SimNonceAllocate")
    read, derive, transaction, injected = _order(
        "ResolveNextNonce(", "SimRngAutoSeedFromNonce", "RunAllocationTransaction(",
        "FailPointCheck FAILPOINT_SIM_AFTER_NONCE", body=entry)
    assert read < derive < transaction < injected
    txn = _procedure(NONCE, "RunAllocationTransaction")
    marked, advanced, cleared = _order("EstablishPending(", "PersistAdvance(",
                                       "ClearPending(", body=txn)
    assert marked < advanced < cleared, txn
    assert "SimEngineRun" not in code, "sampling begins inside the allocation"
    # The attempt record keeps the identity by STATE, and consumption is the
    # stronger claim that only an observed match earns.
    lifecycle = _sim().raw["seeding"]["nonce_lifecycle"]
    assert list(lifecycle["attempt_metadata_preserves"]["persistence_indeterminate"]) == [
        "attempted_auto_nonce", "effective_seed", "durable_indeterminate_result"]
    attempt = _procedure(REPORT, "WriteAttemptBlock")
    assert "package.AutoIdentityKnown" in attempt
    assert "package.NonceConsumed" not in attempt, (
        "a post-write verification failure blanks the preserved identity"
    )

    inactive = semantics["inactive_bank_write_failure"]
    assert inactive["active_bank_changed"] is False
    assert inactive["prior_publication_remains_authoritative"] is True
    assert inactive["corrupted_candidate_has_semantic_standing"] is False
    publish = _procedure(REPORT, "PublishCandidate")
    assert "no semantic standing" in publish
    assert "ClearContents" not in publish, "the failed candidate is erased"

    commit = semantics["final_commit_failure"]
    assert commit["prior_block_restored"] is True
    assert commit["active_bank_changed"] is False
    final = _procedure(REPORT, "FinalCommit")
    assert final.count("Range(SIM_FINAL_COMMIT_RANGE).Value2 = previous") == 1
    assert "If SameBlock(SIM_FINAL_COMMIT_RANGE, previous, 9, 1) Then" in final
    # The capture happens before the write, as the contract states.
    layout = _sim().raw["publication"]["transaction"]
    assert layout["prior_final_commit_block_captured_before_write"] is True
    assert layout["final_commit_failure_restores_prior_block"] is True
    assert layout["final_commit_is_one_write"] is True
    capture, write = _order("previous = SimSheet.Range(SIM_FINAL_COMMIT_RANGE).Value2",
                            "Range(SIM_FINAL_COMMIT_RANGE).Value2 = block", body=final)
    assert capture < write


# ===========================================================================
# The modSimNonce responsibility split
# ===========================================================================
def test_30_the_nonce_responsibility_is_split_and_stays_split() -> None:
    """One coherent responsibility, one module, and a one-way dependency."""
    structure = _structure()
    modules = {m.name: m for m in structure.vba_modules}
    assert NONCE in modules and modules[NONCE].generated is False
    assert "AUTO nonce" in modules[NONCE].responsibility
    # The reporter no longer claims the lifecycle it delegated.
    assert "AUTO nonce lifecycle belongs to modSimNonce" in modules[REPORT].responsibility
    # AND THE REGISTERED RESPONSIBILITY NAMES THE AUTHORITY THE SOURCE ACTUALLY
    # USES. It used to say this module interprets a prior AUTO_NONCE_INDETERMINATE
    # attempt - the rejected carrier. A registry that still describes a withdrawn
    # design teaches it to the next reader.
    duty = modules[NONCE].responsibility
    assert "pending-AUTO-nonce marker" in duty, duty
    assert "never the mutable last-attempt audit axis" in duty, duty
    assert "AUTO_NONCE_INDETERMINATE" not in duty, duty

    nonce = _module(NONCE)
    # ONE PUBLIC ENTRY POINT plus its projected state vocabulary. No endpoint.
    assert "SimNonceAllocate" in nonce.public_procedures
    assert not [n for n in nonce.public_procedures if n.startswith("PCCM_")]
    for endpoint in PHASE6_PUBLIC:
        assert endpoint not in nonce.code, endpoint

    # IT OWNS NO SIMULATION WORK, and writes no attempt row.
    for banned in ("SimEngineRun", "SimStatsDescribe", "SimStatsContingency",
                   "SimFpBuildRequestFingerprint", "SimFpResultDigest",
                   "SnapshotRange", "SummaryRange", "ContingencyRange",
                   "IterationRange", "AttemptRange", "StatusRange",
                   "WriteAttemptBlock", "DeriveSimStatus", "CandidateRunId",
                   "SIM_FINAL_COMMIT_RANGE", "SIM_IDENTITY_ROW_ACTIVE_BANK"):
        assert banned not in nonce.code, banned
    # The only kernel it reaches is the seed derivation its responsibility names.
    assert set(re.findall(r"mod(?:SimRng|SimEngine|SimStats|SimFingerprint)\.(\w+)",
                          nonce.code)) == {"SimRngAutoSeedFromNonce"}

    # THE DEPENDENCY IS ONE-WAY. modSimReport drives modSimNonce and never the
    # reverse - not even to borrow a failpoint constant.
    assert "modSimNonce." in _module(REPORT).code
    assert "modSimReport" not in nonce.code, (
        "the nonce module depends back on the orchestrator"
    )

    # THE REPORTER NO LONGER IMPLEMENTS THE TRANSACTION.
    report = _module(REPORT).code
    for moved in ("SIM_IDENTITY_ROW_NEXT_AUTO_NONCE", "SimRngAutoSeedFromNonce",
                  "FAILPOINT_SIM_AFTER_NONCE", "SIM_NONCE_EXHAUSTED"):
        assert moved not in report, moved

    # THE RUN PACKAGE STAYS PRIVATE. The split is not paid for by exporting it.
    types = re.findall(r"^(Public|Private) Type (\w+)$", _module(REPORT).raw, re.M)
    assert types == [("Private", "SimRunPackage")], types
    assert "SimRunPackage" not in nonce.code, (
        "the run package crossed the module boundary"
    )
    # The interface is scalars, and every out-parameter is a primitive.
    signature = _procedure(NONCE, "SimNonceAllocate")
    signature = signature[: signature.index(") As Boolean")]
    for kind in ("As Boolean", "As Long", "As String"):
        assert kind in signature, kind
    for composite in ("As Variant", "As Object", "As Range", "()"):
        assert composite not in signature, composite


def test_31_both_phase6_orchestration_modules_stay_within_their_limits() -> None:
    """The split was made to satisfy the size control, not to dodge it."""
    import test_phase4_stage_b_source as sizes

    by_name = {m.name: m for m in sizes._handwritten_modules()}
    for name in (REPORT, NONCE):
        assert name in by_name, name
        raw, _, _, code = sizes._line_metrics(by_name[name])
        assert raw < sizes.PHASE5_RAW_LINE_LIMIT, (name, raw)
        assert code < sizes.PHASE5_CODE_LINE_LIMIT, (name, code)
    # And the ceilings themselves were not moved to make room.
    assert sizes.PHASE5_RAW_LINE_LIMIT == 1200
    assert sizes.PHASE5_CODE_LINE_LIMIT == 900
    assert sizes.PHASE4_RAW_LINE_LIMIT == 900


def test_32_the_pending_sidecar_is_a_genuinely_free_coordinate() -> None:
    """F21 collides with nothing the layout already owns.

    Comment text is not proof here; the coordinate is checked against the
    layout the builder actually emits. Column F is the bank-B value column and
    its snapshot ends at the last banked row; row 21 is a SHARED counter row,
    so it has no bank-B twin and nothing is displaced.
    """
    data = _sim().raw["sim_data"]
    cell = data["pending_auto_nonce"]
    identity = data["run_identity"]
    assert cell["cell"] == "F21"
    assert cell["column"] == identity["bank_value_columns"]["B"] == "F"
    row = int(cell["row"])

    fields = identity["fields"]
    snapshot = [f for f in fields if f.get("group") == "snapshot"]
    assert snapshot, "the snapshot group vanished, so this proves nothing"
    # 1. BELOW THE BANK-B SNAPSHOT, which is the only thing column F carries.
    assert row > max(int(f["row"]) for f in snapshot)
    # 2. NOT ANY BANKED CELL, in either bank.
    assert row not in {int(f["row"]) for f in snapshot}
    # 3. THE SHARED ROWS AT AND BELOW IT ARE COLUMN D ONLY, so the shared final
    #    commit cannot reach column F at all.
    shared = [f for f in fields if f.get("group") != "snapshot"]
    assert row in {int(f["row"]) for f in shared} or row > max(
        int(f["row"]) for f in snapshot)
    assert identity["value_column"] == "D" != cell["column"]
    # 4. ABOVE NOTHING: it is inside the identity block's own row span, so it
    #    steals no row from anything below.
    assert int(identity["first_row"]) <= row <= int(identity["last_row"])
    # 5. CLEAR OF THE ITERATION TABLE.
    records = data["iteration_records"]
    assert row < int(records["header_row"])
    # 6. THE ONLY OTHER COLUMN-F CONSUMER IS THE BANK-B ITERATION INDEX, and it
    #    starts below the header row this cell sits above.
    assert records["banks"]["B"]["iteration_index"] == "F"
    assert row < int(records["first_iteration_row"])
    # 7. THE OTHER BANKED BLOCKS ARE IN OTHER COLUMNS ENTIRELY, so the sidecar
    #    cannot be one of their cells whatever their rows are.
    for block in ("summary_statistics", "contingency_ladder"):
        columns = data[block]["bank_value_columns"]
        used = {c for bank in columns.values() for c in bank.values()}
        assert cell["column"] not in used, (block, sorted(used))

    # AND THE BUILT WORKBOOK AGREES: nothing was written there.
    from openpyxl import load_workbook

    book = load_workbook(BUILD / "PCCM_stageA.xlsx")
    sheet = book[data["sheet"]]
    assert sheet["F21"].value is None, sheet["F21"].value
    occupied = sorted(c.row for c in sheet["F"] if c.value is not None)
    assert 21 not in occupied, occupied


def test_33_the_sidecar_added_no_row_and_moved_no_ceiling() -> None:
    """A free cell was used precisely so nothing had to shift."""
    data = _sim().raw["sim_data"]
    records = data["iteration_records"]
    assert int(records["header_row"]) == 33
    assert int(records["first_iteration_row"]) == 34
    # THE TECHNICAL CEILING IS THE SAME NUMBER IT WAS, and it is still derived
    # from the layout rather than restated as a free literal.
    ceiling = _sim().raw["iterations"]["technical_ceiling"]
    assert int(ceiling["reserved_rows_h"]) == 33
    assert int(ceiling["max_iterations_representable"]) == 1048543
    assert int(ceiling["max_iterations_representable"]) == (
        int(ceiling["max_excel_rows"]) - int(ceiling["reserved_rows_h"]))
    # AND THE SIDECAR REALLY IS ABOVE ALL OF IT.
    assert int(data["pending_auto_nonce"]["row"]) < int(records["header_row"])


def test_34_the_sidecar_coordinate_is_written_down_exactly_once() -> None:
    """One authority, one generated constant, no second literal.

    A coordinate spelled independently in two production procedures is two
    authorities that agree only by luck.
    """
    generated = (BUILD / "vba" / "modSimContract.bas").read_text(encoding="utf-8")
    assert 'Public Const SIM_PENDING_AUTO_NONCE_CELL As String = "F21"' in generated
    assert generated.count("SIM_PENDING_AUTO_NONCE_CELL") == 1

    users = [name for name in _module(NONCE).procedures
             if "SIM_PENDING_AUTO_NONCE_CELL" in _procedure(NONCE, name)]
    assert users == ["PendingCell"], users
    # NO HANDWRITTEN MODULE SPELLS THE COORDINATE OUT.
    for name in tuple(FROZEN_SOURCE) + (REPORT, NONCE):
        code = _module(name).code_without_string_removal
        assert '"F21"' not in code, name
    # AND EVERY SIDECAR TOUCH GOES THROUGH THE ONE ACCESSOR.
    for name in _module(NONCE).procedures:
        if name == "PendingCell":
            continue
        body = _procedure(NONCE, name)
        assert "SIM_DATA_SHEET).Range(SIM_PENDING" not in body, name


# ===========================================================================
# 12. Nothing here claims a runtime
# ===========================================================================
def test_26_no_step_13_scenario_exists_yet() -> None:
    scenarios = (PCCM_ROOT / "bootstrap" / "windows"
                 / "phase5_gate_b_scenarios.ps1").read_text(encoding="utf-8")
    for premature in ("P5-SIM", "P6-RUN", "P6-SIM", "Add-Phase6Result",
                      "PCCM_RunSimulation"):
        assert premature not in scenarios, f"a Step-13 runtime scenario exists: {premature}"
    harness = (PCCM_ROOT / "bootstrap" / "windows"
               / "phase4_functional_test.ps1").read_text(encoding="utf-8")
    assert "PCCM_RunSimulation" not in harness, "the harness invokes the endpoint"


def test_27_no_phase7_module_exists() -> None:
    """THE LIST SHRINKS ONE PACKAGE AT A TIME, AND ONLY WHEN ONE LANDS.

    `modSimSensitivity` came off it in P7-2 and `modSimAnnual` comes off it now:
    the annual stochastic computation is the accepted P7-5 package and the
    module registry declares it. `modSimPostReport` is P7-4's. Everything still
    on the list is still premature - the dashboard, charts and reconciliation
    are Phase 8, and `modSimReplay` names a module that will never exist because
    replay lives in modSimEngine, which owns the generator. `modSimCorrelation`
    stays on it permanently: inter-driver correlation is out of scope, not
    deferred.
    """
    names = {path.stem for path in SRC_VBA.glob("*.bas")}
    assert "modSimSensitivity" in names, "the accepted P7-2 kernel is missing"
    assert "modSimAnnual" in names, "the accepted P7-5 annual module is missing"
    for premature in ("modSimDashboard", "modSimReconcile",
                      "modSimCorrelation", "modSimReplay", "modSimCharts"):
        assert premature not in names, premature
    declared = {module.name for module in _structure().vba_modules}
    assert declared >= names, sorted(names - declared)


if __name__ == "__main__":
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))


# ===========================================================================
# 17. EVERY CROSS-MODULE CALL IS WELL-FORMED
# ===========================================================================
# WHY THIS EXISTS, AND WHY IT IS GENERAL.
#
# Windows Runtime Run 2 reached Excel, and `Debug > Compile VBAProject` on the
# retained workbook produced a real compiler diagnostic:
#
#     Compile error: Argument not optional
#
# on `modWorkbook.IsWholeInRange(raw, CDbl(...), CDbl(...))` inside
# `modSimNonce.ReadPending`. The declaration takes FOUR arguments, the last a
# `ByRef Result As Double`, and five Phase-6 call sites passed three. VBA
# compiles on demand, so nothing before that point had ever required those
# procedure bodies to compile - and the whole Phase-5 and Phase-6 behavioural
# matrix was blocked behind it.
#
# THE CONTROLS THAT EXISTED PROVED THE CALL WAS PRESENT, never that it was
# well-formed - `assert "IsWholeInRange" in body`. Pinning the two corrected
# lines as text would close those two and leave the class open, so this walks
# EVERY qualified cross-module call in the hand-written production source and
# checks its argument count against the callee's own declaration.
_CALL_DECLARATION = re.compile(
    r"^\s*(?:Public|Private)\s+(?:Sub|Function)\s+(\w+)\s*\((.*?)\)\s*(?:As\s+\w+)?\s*$"
)


def _joined(module: VbaModule) -> str:
    """Comment- and string-stripped code with VBA line continuations joined.

    A call wrapped across `_` lines is one call, and counting its arguments per
    physical line would count some of them twice and some not at all.
    """
    return re.sub(r"\s+_\r?\n\s*", " ", module.code)


def _split_top_level(text: str) -> list[str]:
    if not text.strip():
        return []
    parts, depth, buf = [], 0, ""
    for char in text:
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        if char == "," and depth == 0:
            parts.append(buf)
            buf = ""
        else:
            buf += char
    parts.append(buf)
    return [part for part in parts if part.strip()]


def _declared_arity() -> dict[tuple[str, str], tuple[int, int | None]]:
    """(module, procedure) -> (required, maximum). `None` maximum means ParamArray."""
    found: dict[tuple[str, str], tuple[int, int | None]] = {}
    for name, module in _modules().items():
        for line in _joined(module).splitlines():
            match = _CALL_DECLARATION.match(line)
            if not match:
                continue
            params = _split_top_level(match.group(2))
            required = sum(
                1 for param in params
                if not re.match(r"\s*(Optional|ParamArray)\b", param, re.IGNORECASE)
            )
            variadic = any(
                re.match(r"\s*ParamArray\b", param, re.IGNORECASE) for param in params
            )
            found[(name, match.group(1))] = (required, None if variadic else len(params))
    return found


def _qualified_calls() -> list[tuple[str, str, str, int, str]]:
    """Every `modX.Proc(...)` call site: (caller, module, procedure, given, line)."""
    calls = []
    for caller, module in sorted(_modules().items()):
        for line in _joined(module).splitlines():
            if _CALL_DECLARATION.match(line):
                continue
            for target, procedure in re.findall(r"\b(mod\w+)\.(\w+)\s*\(", line):
                opened = re.search(rf"\b{target}\.{procedure}\s*\((.*)", line)
                depth, buf, args = 1, "", []
                for char in opened.group(1):
                    if char == "(":
                        depth += 1
                    elif char == ")":
                        depth -= 1
                        if depth == 0:
                            break
                    if depth == 1 and char == ",":
                        args.append(buf)
                        buf = ""
                    else:
                        buf += char
                if buf.strip():
                    args.append(buf)
                calls.append((caller, target, procedure,
                              len([a for a in args if a.strip()]), line.strip()))
    return calls


def test_35_every_qualified_cross_module_call_supplies_its_arguments() -> None:
    """`Argument not optional` is a COMPILE error, and VBA compiles on demand.

    A procedure body nothing has reached yet can hold a fatal call for as long
    as nothing reaches it, which is exactly how five malformed calls survived
    every static control up to Runtime Run 2.
    """
    declared = _declared_arity()
    assert len(declared) > 200, len(declared)
    calls = _qualified_calls()
    assert len(calls) > 300, len(calls)

    problems = []
    checked = 0
    for caller, target, procedure, given, line in calls:
        arity = declared.get((target, procedure))
        if arity is None:
            continue  # a generated module or a VBA/Excel member, not ours to type
        required, maximum = arity
        checked += 1
        if given < required or (maximum is not None and given > maximum):
            problems.append(
                f"{caller}: {target}.{procedure} declares {required}"
                + (f"..{maximum}" if maximum != required else "")
                + f" argument(s), {given} given -- {line[:70]}"
            )
    assert checked > 300, checked
    assert not problems, "malformed cross-module calls:\n  " + "\n  ".join(problems)


def test_36_the_range_check_helper_keeps_its_out_parameter() -> None:
    """The signature the five defective calls were written against.

    Pinned in full, and by SHAPE: the fourth parameter is what makes the helper
    a parse as well as a test, and widening it to Optional would silently make
    the defective calls legal again while changing what they mean.
    """
    declaration = re.search(
        r"Public Function IsWholeInRange\((.*?)\)\s*As Boolean",
        re.sub(r"\s+_\r?\n\s*", " ", _module("modWorkbook").code),
    )
    assert declaration, "modWorkbook no longer declares IsWholeInRange"
    params = [param.strip() for param in _split_top_level(declaration.group(1))]
    assert params == [
        "ByVal Value As Variant",
        "ByVal MinValue As Double",
        "ByVal MaxValue As Double",
        "ByRef Result As Double",
    ], params

    # AND EVERY PHASE-6 CALL SUPPLIES ALL FOUR, named here so the report is
    # about the Phase-6 surface rather than the whole repository.
    phase6 = [call for call in _qualified_calls()
              if call[0].startswith("modSim") and call[2] == "IsWholeInRange"]
    # SIX NOW. P7-4's orchestrator reads the published run identity through the
    # same helper, with all four arguments - which is the property this control
    # is about, and it is asserted for every caller below.
    assert len(phase6) == 6, [call[0] for call in phase6]
    for caller, _target, _procedure, given, line in phase6:
        assert given == 4, f"{caller}: {given} argument(s) -- {line[:70]}"


# ===========================================================================
# The publication verify predicate, EVALUATED as written
# ===========================================================================
# Run 4 announced "the previous shared block could not be restored" while the
# sheet showed a perfectly restored block. `SameCell` was the cause, and no
# control could see it because every control here reads STRUCTURE. So this one
# reads the predicate out of the module and RUNS it over the value pairs the
# publication transaction actually produces.
#
# It is not a VBA interpreter. It evaluates the small vocabulary this one
# function uses, over a model of the Variants Value2 returns, and it REFUSES
# rather than guesses if the function grows a construct outside that vocabulary.
class _VbaEmpty:
    """VBA `Empty` - what `Range.Value2` returns for a blank cell."""

    def __repr__(self) -> str:  # pragma: no cover - debugging aid
        return "Empty"


EMPTY = _VbaEmpty()
VB_EMPTY, VB_DOUBLE, VB_STRING = 0, 5, 8


def _vb_vartype(value: object) -> int:
    if value is EMPTY:
        return VB_EMPTY
    return VB_STRING if isinstance(value, str) else VB_DOUBLE


def _vb_cstr(value: object) -> str:
    if value is EMPTY:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _vb_len(value: object) -> int:
    return len(_vb_cstr(value))


def _vb_strcomp(left: str, right: str, _mode: int = 0) -> int:
    return 0 if left == right else (-1 if left < right else 1)


def _vb_is_numeric(value: object) -> bool:
    # DELIBERATELY UNDECIDED FOR Empty. Whether VBA calls a blank numeric is a
    # coercion only Windows can settle, and the predicate must not depend on it:
    # it has to decide blank BEFORE it reaches a numeric test. Raising here is
    # how this control refuses to guess.
    if value is EMPTY:
        raise AssertionError(
            "SameCell reached IsNumeric with a blank; the blank cases must be "
            "decided before any numeric coercion"
        )
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    try:
        float(value)
    except (TypeError, ValueError):
        return False
    return True


def _vb_cdbl(value: object) -> float:
    if value is EMPTY:
        return 0.0
    return float(value)


_VBA_NAMES = {
    "IsEmpty": "_vb_is_empty", "VarType": "_vb_vartype", "Len": "_vb_len",
    "CStr": "_vb_cstr", "StrComp": "_vb_strcomp", "IsNumeric": "_vb_is_numeric",
    "CDbl": "_vb_cdbl", "vbString": str(VB_STRING), "vbBinaryCompare": "0",
    "Not": "not", "And": "and", "Or": "or",
}
_VBA_EVAL_NAMESPACE = {
    "_vb_is_empty": lambda value: value is EMPTY,
    "_vb_vartype": _vb_vartype, "_vb_len": _vb_len, "_vb_cstr": _vb_cstr,
    "_vb_strcomp": _vb_strcomp, "_vb_is_numeric": _vb_is_numeric,
    "_vb_cdbl": _vb_cdbl,
}


def _vb_expression(text: str) -> str:
    """Translate one VBA expression, and REFUSE anything outside the vocabulary."""
    assert "<" not in text and ">" not in text, f"unsupported comparison: {text}"
    translated = re.sub(r"\b[A-Za-z_]\w*\b",
                        lambda m: _VBA_NAMES.get(m.group(0), m.group(0)), text)
    translated = re.sub(r"(?<![=!<>])=(?!=)", "==", translated)
    unknown = {name for name in re.findall(r"\b[A-Za-z_]\w*\b", translated)
               if name not in set(_VBA_EVAL_NAMESPACE) | {"written", "wanted",
                                                          "not", "and", "or"}}
    assert not unknown, f"SameCell uses a construct this control cannot evaluate: {unknown}"
    return translated


def _same_cell_program() -> list[tuple[str, str]]:
    """SameCell's statements, from the file, as (kind, payload) pairs."""
    source = strip_comments(_procedure(REPORT, "SameCell"))
    statements = [text for _, text in logical_statements(source)]
    assert re.match(r"\s*Private Function SameCell\b", statements[0]), statements[0]
    assert re.fullmatch(r"\s*End Function\s*", statements[-1]), statements[-1]
    program: list[tuple[str, str]] = []
    for statement in statements[1:-1]:
        text = statement.strip()
        guard = re.fullmatch(r"If\s+(.*?)\s+Then\s+Exit Function", text, re.I)
        block = re.fullmatch(r"If\s+(.*?)\s+Then", text, re.I)
        assign = re.fullmatch(r"SameCell\s*=\s*(.*)", text, re.I)
        if guard:
            program.append(("guard", guard.group(1)))
        elif block:
            program.append(("if", block.group(1)))
        elif assign:
            program.append(("set", assign.group(1)))
        elif re.fullmatch(r"Exit Function", text, re.I):
            program.append(("exit", ""))
        elif re.fullmatch(r"End If", text, re.I):
            program.append(("endif", ""))
        else:
            raise AssertionError(f"SameCell grew a statement this control cannot run: {text}")
    return program


def _same_cell(written: object, wanted: object) -> bool:
    """Run the predicate AS WRITTEN over one pair of Variants."""
    program = _same_cell_program()
    scope = dict(_VBA_EVAL_NAMESPACE, written=written, wanted=wanted)
    result, index, skipping = False, 0, False
    while index < len(program):
        kind, payload = program[index]
        index += 1
        if kind == "endif":
            skipping = False
            continue
        if skipping:
            continue
        if kind == "if":
            skipping = not eval(_vb_expression(payload), {"__builtins__": {}}, scope)
        elif kind == "guard":
            if eval(_vb_expression(payload), {"__builtins__": {}}, scope):
                return result
        elif kind == "set":
            result = bool(eval(_vb_expression(payload), {"__builtins__": {}}, scope))
        elif kind == "exit":
            return result
    return result


def test_37_the_publication_verify_accepts_a_restored_blank() -> None:
    """The Run-4 defect, as a truth table over the Variants Value2 returns.

    `BuildCommitBlock` writes `vbNullString` into the blank fields of a CANDIDATE
    block; `FinalCommit` captures the previous block with `Range.Value2`, which
    returns `Empty` for those same fields. Both blocks are verified with the same
    predicate, so it has to accept both spellings of blank - and reject a value
    that is not blank at all, including the zero `CDbl(Empty)` would produce.
    """
    # THE RESTORE PATH: a captured blank, written back, read back as Empty.
    assert _same_cell(EMPTY, EMPTY), (
        "a blank restored over a blank does not verify; FinalCommit would "
        "announce that the previous shared block could not be restored"
    )
    # THE CANDIDATE PATH, unchanged: a built vbNullString lands as a blank cell.
    assert _same_cell(EMPTY, ""), "the candidate blank-write semantics were lost"
    assert _same_cell("", EMPTY)
    assert _same_cell("", "")

    # AND BLANK IS STILL NOT ANYTHING ELSE. `_vb_is_numeric` refuses a blank, so
    # these also prove the predicate settles every blank case before it coerces.
    assert not _same_cell(0.0, EMPTY), (
        "a fabricated zero verifies against a captured blank: CDbl(Empty) is 0"
    )
    assert not _same_cell(EMPTY, 0.0)
    assert not _same_cell("text", EMPTY)
    assert not _same_cell(EMPTY, "text")
    assert not _same_cell(0.0, "")
    assert not _same_cell("", 0.0)

    # THE NON-BLANK COMPARISONS THE TRANSACTION ALSO DEPENDS ON.
    assert _same_cell("BankA", "BankA")
    assert not _same_cell("BankA", "BankB")
    assert _same_cell(1234.0, 1234.0)
    assert not _same_cell(1234.0, 1235.0)


def test_38_the_whole_captured_commit_block_verifies_after_a_restore() -> None:
    """Not one cell - the nine-field block FinalCommit actually restores.

    Rows 3 and 6 are the blank ones: `BuildCommitBlock` leaves the detail blank
    on success and the consumed nonce blank when no nonce was consumed, so a
    previously published block captured off the sheet carries Empty there.
    """
    captured = [46264.9, "SUCCESS", EMPTY, "AUTO", 8891.0, EMPTY, "CURRENT",
                46264.9, "BankA"]
    assert all(_same_cell(value, value) for value in captured), (
        "a captured block does not verify against itself, so SameBlock would "
        "report a restore failure for a restore that physically succeeded"
    )
    # AND A BLOCK THAT CAME BACK WRONG IS STILL REFUSED.
    damaged = list(captured)
    damaged[8] = "BankB"
    assert not all(_same_cell(new, old) for new, old in zip(damaged, captured))
