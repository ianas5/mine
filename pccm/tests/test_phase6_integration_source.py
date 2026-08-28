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
# with the hash Step 11 and Round 4A left it at. Named, never counted.
FROZEN_SOURCE = {
    "modSimRng": "3d7c2cb365df03ccf73722f39b0c10e8964381e7cdd243732381dac7638257e3",
    "modSimSample": "5553198289bd98a7c84025868ac03c9f8ec95da3c01b23249c0da57d77901877",
    "modSimEngine": "f1283fe7d5d2ffcc5345dab9a00f68d3685b787563d104f50a886c5ed409abab",
    "modSimStats": "98bd21b227047d04e6847e554e027b339cf01dfb1112c1539a9e334966233be0",
    "modSimFingerprint": "9e6ad972fe59ead9e34c7d65b807dd0f2ca1cb1b29bfa71b377a4eb8f65cdfda",
    "modSimNonce": "6e0ed05c90c09144fb1f7ecbab2eca03828ddb9b40c7711222db9084b1c83ef3",
    "modSimReport": "49b7602a51a4d5995f77182cb8c26aa53eceec4be5152477665ff9f4644d2b06",
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
    callers = [name for name, module in _modules().items()
               if BRIDGE in module.code and name != "modCalcReport"]
    assert callers == [REPORT], callers
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
    assert found == phase4 | set(PHASE5_ENDPOINTS) | set(PHASE6_PUBLIC), sorted(
        found ^ (phase4 | set(PHASE5_ENDPOINTS) | set(PHASE6_PUBLIC)))
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
                     "PCCM_RunSensitivity", "PCCM_Dashboard"):
        assert invented not in code, invented


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
def test_23_every_accepted_module_is_byte_identical() -> None:
    for name, expected in FROZEN_SOURCE.items():
        path = SRC_VBA / f"{name}.bas"
        actual = hashlib.sha256(path.read_bytes()).hexdigest()
        assert actual == expected, f"{name}.bas moved: {actual}"


def test_24_the_generated_authority_is_byte_identical() -> None:
    for path, expected in (
        (BUILD / "vba" / "modSimContract.bas",
         "daa4d27889c30eadb2ab892bcfa4e6f6bab8a137aae79a01a8d8f1e8e1c215ac"),
        (BUILD / "phase6_cases.json",
         "8019683a0490fcf0740cf07244524973d9b7470c933f1003059025b6b019a0be"),
    ):
        actual = hashlib.sha256(path.read_bytes()).hexdigest()
        assert actual == expected, f"{path.name} moved: {actual}"


def test_25_the_accepted_reporter_prefix_is_still_byte_identical() -> None:
    """The Step-11 technique, re-proved at the integration layer."""
    banner = ("' ==========================================================================\n"
              "' STEP 11 ADDITION - THE PHASE-6 PREPARATION BRIDGE\n")
    text = (SRC_VBA / "modCalcReport.bas").read_text(encoding="utf-8")
    assert text.count(banner) == 1
    accepted = text[: text.index(banner)]
    assert hashlib.sha256(accepted.encode("utf-8")).hexdigest() == (
        "5d4568aef01037fd2999915da87a550d02033441b8c26c80f9386d4fcf8b087f")
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
    names = {path.stem for path in SRC_VBA.glob("*.bas")}
    for premature in ("modSimSensitivity", "modSimDashboard", "modSimAnnual",
                      "modSimReconcile", "modSimCorrelation", "modSimReplay",
                      "modSimCharts"):
        assert premature not in names, premature
    declared = {module.name for module in _structure().vba_modules}
    assert declared >= names, sorted(names - declared)


if __name__ == "__main__":
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
