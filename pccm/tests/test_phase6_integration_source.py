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
    "modSimReport": "a48bddffc5c512ed30a0ab78c2cd802fea57031bb1ebbf5b09f0fed1a394f60b",
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
        "SimFpBuildRequestFingerprint", "SimFpResultDigest",
        "SimRngAutoSeedFromNonce", BRIDGE,
    }, sorted(called)
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
         "1d949be659d0afc3e18501a34b7d372bab3df575fc1a981cfd60dcf1f293a753"),
        (BUILD / "phase6_cases.json",
         "98f835375f5b8f548172c21ae6102b50fef7e6a001e196ece0741c987d78b6d1"),
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
def test_28_no_transaction_stage_escapes_to_the_invocation_axis() -> None:
    """After the AUTO nonce is spent, every failure owes an attempt record.

    The contract's `refusal_or_failure_after_auto_allocation` requires
    `attempt_metadata_updated: true`, and that is a statement about REAL COM
    failures, not only about a helper returning False. Each stage that touches
    the worksheet therefore carries its own scoped handler, and RunSimulation
    routes its False through RecordRefusal or RecordFailure.
    """
    run = _procedure(REPORT, "RunSimulation")
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
    # EVERY COM-FALLIBLE STAGE CARRIES ITS OWN ENVELOPE.
    assert "On Error GoTo AllocationFailed" in _procedure(REPORT, "AllocateAutoNonce")
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
    for name in _module(REPORT).procedures:
        body = _procedure(REPORT, name)
        if "FailPointCheck" not in body:
            continue
        armed = [m.start() for m in re.finditer(r"On Error GoTo (?!0\b)\w+", body)]
        assert armed, f"{name} fires a failpoint with no handler armed"
        assert min(armed) < body.index("FailPointCheck"), name
    # AND NO PHASE-6 MODULE SUPPRESSES ERRORS WHOLESALE. (The Phase-4 modules
    # carry their own documented `On Error Resume Next` whitelist, policed by
    # test_phase4_stage_b_source.py; nothing here widens or narrows it.)
    for name in tuple(FROZEN_SOURCE) + (REPORT,):
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
    code = _module(REPORT).code
    assert "NEXT_AUTO_NONCE" in code
    for _, statement in logical_statements(code):
        if "SIM_IDENTITY_ROW_NEXT_AUTO_NONCE" in statement and "- 1" in statement:
            raise AssertionError(f"the consumed nonce is rolled back: {statement}")
    allocate = _procedure(REPORT, "AllocateAutoNonce")
    write, verify, mark, injected = _order(
        "SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2",
        "ReadMachineLong", "package.NonceConsumed = True",
        "FailPointCheck FAILPOINT_SIM_AFTER_NONCE", body=allocate)
    assert write < verify < mark < injected
    assert "SimEngineRun" not in allocate, "sampling begins inside the allocation"

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
