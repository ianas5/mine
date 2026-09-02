#!/usr/bin/env python3
"""PCCM Phase 6 Step-8 conformance tests for `src/vba/modSimEngine.bas`.

The Monte Carlo iteration loop and canonical contribution accumulation over
already-resolved in-memory `DriverFactors`.

--------------------------------------------------------------------------------
WHAT THESE TESTS PROVE, AND WHAT THEY DO NOT
--------------------------------------------------------------------------------
SOURCE CONFORMANCE, on Linux, now. The module is read as text: its purity, its
public surface, the canonical execution order, the factor semantics, the
consumption contract, the preparation boundary and the transactional output -
and the arithmetic those statements describe, against the accepted Phase-6
Python oracle and the accepted Step-5 corpus.

VBA EXECUTION CONFORMANCE is NOT proved here and is deferred to Gate B on
Windows. No VBA runtime exists in this step. No test in this file may be read as
"VBA produced this total".

THE COMPARISON POLICY IS NOT STRENGTHENED. Each corpus case is checked under the
policy the corpus assigns it: EXACT where it says EXACT, the accepted
tolerance-bounded rule where it says TOLERANCE_BOUNDED, a distributional
statement where it says STATISTICAL, and a relation between two runs where it
says SAME_RUNTIME_ONLY.

TWO PRIMITIVES ARE BORROWED, NOT TRANSCRIBED. `SafeProduct` and `SafeSignedSum`
have accepted Phase-5 VBA bodies whose second tier uses an exact-arithmetic UDT
with dynamic limbs and scoped error handlers - constructs the source
transcriber does not model. Their accepted PYTHON counterparts in
`calc_numeric` are bound instead, and their real VBA SIGNATURES are read out of
`modCalcFactors.bas` so the ByRef/ByVal call convention stays the module's own
rather than something retyped here.

Runs standalone or under pytest.
"""

from __future__ import annotations

import json
import math
import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

from pccm_builder import (  # noqa: E402
    RngReference,
    load_calc_contract,
    result_digest,
    load_contract,
    load_sim_contract,
    load_structure_contract,
)
from pccm_builder.calc_cases import to_model, tolerances_from  # noqa: E402
from pccm_builder.sim_cases import GATE_B_PARITY_PLAN_CASES  # noqa: E402
from pccm_builder.calc_numeric import (  # noqa: E402
    CalculationRefusal,
    safe_product,
    safe_signed_sum,
)
from pccm_builder.calc_oracle import (  # noqa: E402
    AppliedTimeline,
    CalculationModel,
    CostDriver,
    FxRow,
    RiskDriver,
    calculate,
)
from pccm_builder.contract_loader import ContractError  # noqa: E402
from pccm_builder.sim_emit import render_sim_contract_module  # noqa: E402
from pccm_builder.sim_oracle import prepare_simulation, run_simulation  # noqa: E402
from pccm_builder.spec_loader import load_spec  # noqa: E402
from pccm_builder.vba_source import (  # noqa: E402
    VbaModule,
    contains_construct,
    logical_statements,
)

from phase6_vba_transcribe import _Ref, _val, build as _build_transcription  # noqa: E402

SRC_VBA = PCCM_ROOT / "src" / "vba"
SIM_ENGINE_BAS = SRC_VBA / "modSimEngine.bas"
SIM_RNG_BAS = SRC_VBA / "modSimRng.bas"
SIM_SAMPLE_BAS = SRC_VBA / "modSimSample.bas"
CALC_FACTORS_BAS = SRC_VBA / "modCalcFactors.bas"
SPEC = PCCM_ROOT / "spec"
CASES_JSON = PCCM_ROOT / "build" / "phase6_cases.json"
STEP0_RECORD = PCCM_ROOT / "docs" / "phase6_step0.md"

_CACHE: dict[str, object] = {}


# ---------------------------------------------------------------------------
# Source access
# ---------------------------------------------------------------------------
def _module() -> VbaModule:
    return VbaModule(
        name="modSimEngine", path=SIM_ENGINE_BAS,
        raw=SIM_ENGINE_BAS.read_text(encoding="utf-8"),
    )


def _code() -> str:
    return _module().code


def _procedure(name: str) -> str:
    code = _module().code_without_string_removal
    pattern = re.compile(
        rf"^\s*(?:Public|Private)\s+(?:Function|Sub)\s+{re.escape(name)}\b", re.M
    )
    match = pattern.search(code)
    assert match, f"{name} is not declared"
    tail = code[match.start():]
    end = re.search(r"^\s*End\s+(?:Function|Sub)\s*$", tail, re.M)
    assert end, f"{name} has no End"
    return tail[: end.end()]


def _loop_body() -> str:
    """The ITERATION LOOP alone - what the hot-loop rules govern."""
    body = _procedure("SimEngineRun")
    start = body.index("For iteration = 1 To iterations")
    end = body.index("Next iteration", start)
    return body[start:end]


def _contribution() -> str:
    """The shared per-driver contribution routine.

    P7-3 EXTRACTED IT OUT OF THE LOOP so the replay and the simulation reach one
    implementation. The properties the controls below assert did not change; the
    place they are written down did, and reading the loop alone would now be
    reading a body the arithmetic has left.
    """
    return _procedure("SimEngineContribution")


def _constants() -> dict:
    """The generated Phase-6 projection, the Phase-5 DIST kinds, and the engine's own.

    P7-3 ADDED THE THIRD SOURCE. `SIM_MEASURE_NOMINAL` and `SIM_MEASURE_PV` are
    declared by modSimEngine itself, so the transcription cannot resolve them
    unless this reads the module it is transcribing. They are read from the
    source rather than restated here: a harness that hardcoded the two words
    would still pass if the module swapped them.
    """
    if "consts" not in _CACHE:
        out: dict = {}
        rendered = render_sim_contract_module(
            load_spec(SPEC / "workbook.yaml"),
            load_sim_contract(SPEC / "sim_contract.yaml"),
            load_contract(SPEC / "input_contract.yaml"),
        )
        for text in (rendered, CALC_FACTORS_BAS.read_text(encoding="utf-8"),
                     SIM_ENGINE_BAS.read_text(encoding="utf-8")):
            for line in text.splitlines():
                match = re.match(r"^Public Const (\w+) As (\w+) = (.*)$", line)
                if not match:
                    continue
                name, kind, rest = match.groups()
                literal = rest.split("    '")[0].rstrip()
                out[name] = (literal[1:-1] if kind == "String"
                             else (float(literal) if kind == "Double" else int(literal)))
        _CACHE["consts"] = out
    return _CACHE["consts"]  # type: ignore[return-value]


def _const(name: str):
    return _constants()[name]


def _safe_product(factors, count, result) -> bool:
    try:
        result.v = safe_product([float(v) for v in factors[: int(_val(count))]], "engine")
    except (CalculationRefusal, ContractError):
        return False
    return True


def _safe_signed_sum(terms, count, result) -> bool:
    try:
        result.v = safe_signed_sum([float(v) for v in terms[: int(_val(count))]], "engine")
    except (CalculationRefusal, ContractError):
        return False
    return True


def _transcribe() -> dict:
    if "vba" not in _CACHE:
        _CACHE["vba"] = _build_transcription(
            {
                "modSimRng": SIM_RNG_BAS,
                "modSimSample": SIM_SAMPLE_BAS,
                "modSimEngine": SIM_ENGINE_BAS,
                "modCalcFactors": CALC_FACTORS_BAS,
            },
            _constants(),
            only={"modCalcFactors": {"IsUsableDouble"}},
            signature_only={"modCalcFactors": {"SafeProduct", "SafeSignedSum"}},
            extra={
                "MAX_DOUBLE": sys.float_info.max,
                "SafeProduct": _safe_product,
                "SafeSignedSum": _safe_signed_sum,
            },
        )
    return _CACHE["vba"]  # type: ignore[return-value]


# ---------------------------------------------------------------------------
# The accepted oracle, and fixtures built through it
# ---------------------------------------------------------------------------
DIST_OF = {"Triangular": 1, "Beta-PERT": 2, "Uniform": 3}


def _sim():
    if "sim" not in _CACHE:
        _CACHE["sim"] = load_sim_contract(SPEC / "sim_contract.yaml")
    return _CACHE["sim"]


def _inputs():
    if "inputs" not in _CACHE:
        _CACHE["inputs"] = load_contract(SPEC / "input_contract.yaml")
    return _CACHE["inputs"]


def _tolerances():
    if "tol" not in _CACHE:
        _CACHE["tol"] = tolerances_from(load_calc_contract(SPEC / "calc_contract.yaml"))
    return _CACHE["tol"]


def _ref() -> RngReference:
    if "ref" not in _CACHE:
        _CACHE["ref"] = RngReference.from_contracts(_sim(), _inputs())
    return _CACHE["ref"]  # type: ignore[return-value]


def _cost(permanent_id, distribution="Triangular", minimum=80.0, most_likely=100.0,
          maximum=150.0, quantity=1.0) -> CostDriver:
    return CostDriver(permanent_id, distribution, "SAR", "Standard",
                      minimum, most_likely, maximum, (1.0,), quantity=quantity)


def _risk(permanent_id, distribution="Triangular", minimum=100.0, most_likely=200.0,
          maximum=400.0, probability=0.5) -> RiskDriver:
    return RiskDriver(permanent_id, distribution, "SAR", "Standard",
                      minimum, most_likely, maximum, (1.0,), probability=probability)


def _model(costs=(), risks=()) -> CalculationModel:
    """One applied year, weight 1, no inflation, FX 1: Knom and Kpv are exactly 1."""
    return CalculationModel(
        timeline=AppliedTimeline(2026, 2026, 1), discount_rate=0.10,
        fx_rows=(FxRow("SAR", 1),), inflation_rates={"Standard": {}},
        cost_drivers=tuple(costs), risk_drivers=tuple(risks),
    )


def _discounted_model(costs=(), risks=()) -> CalculationModel:
    """Two applied years at weight 0.5: Knom stays 1 and Kpv does NOT.

    Every other fixture here collapses to Knom = Kpv = 1, which makes a
    nominal/PV confusion arithmetically invisible. This one exists so it is not.
    """
    return CalculationModel(
        timeline=AppliedTimeline(2026, 2026, 2), discount_rate=0.10,
        fx_rows=(FxRow("SAR", 1),), inflation_rates={"Standard": {2027: 0.0}},
        cost_drivers=tuple(costs), risk_drivers=tuple(risks),
    )


def _spread(permanent_id, kind, distribution="Uniform", minimum=10.0, most_likely=None,
            maximum=60.0, quantity=2.0, probability=0.4):
    """A driver whose profile spans the two applied years of `_discounted_model`."""
    if kind == "cost":
        return CostDriver(permanent_id, distribution, "SAR", "Standard",
                          minimum, most_likely, maximum, (0.5, 0.5), quantity=quantity)
    return RiskDriver(permanent_id, distribution, "SAR", "Standard",
                      minimum, most_likely, maximum, (0.5, 0.5), probability=probability)


def _factor_records(model: CalculationModel, order=None) -> list[dict]:
    """Resolved DriverFactors, in a chosen PHYSICAL order.

    This is the accepted Phase-5 factor boundary: Knom, Kpv, Quantity and
    Probability come from `calculate`, and the engine never re-derives them.
    """
    resolved = calculate(model, _tolerances())
    by_id: dict[tuple[bool, str], object] = {}
    for record in resolved.drivers:
        kind = str(getattr(record.driver_kind, "value", record.driver_kind)).upper()
        by_id[("RISK" in kind, record.permanent_id)] = record
    out: list[dict] = []
    for is_risk, drivers in ((False, model.cost_drivers), (True, model.risk_drivers)):
        for driver in drivers:
            record = by_id[(is_risk, driver.permanent_id)]
            out.append({
                "PermanentId": driver.permanent_id,
                "IsRisk": is_risk,
                "Knom": float(record.knom),
                "Kpv": float(record.kpv),
                "Quantity": 0.0 if is_risk else float(record.quantity),
                "Probability": float(record.probability) if is_risk else 0.0,
                "DistKind": DIST_OF[driver.distribution],
                "CentralBasis": "",
                "MinValue": float(driver.min_value),
                "MostLikely": 0.0 if driver.most_likely is None else float(driver.most_likely),
                "MaxValue": float(driver.max_value),
                "Central": 0.0,
                "MeanValue": 0.0,
            })
    return out if order is None else [out[i] for i in order]


def _run(records, seed=12345, iterations=1000):
    nominal, pv, detail = [], [], _Ref("")
    ok = _transcribe()["SimEngineRun"](
        records, _Ref(len(records)), _Ref(seed), _Ref(iterations), nominal, pv, detail)
    return ok, nominal, pv, detail.v


def _oracle(model: CalculationModel, seed=12345, iterations=1000):
    prepared, _ = prepare_simulation(
        _ref(), _sim(), _inputs(), model, _tolerances(),
        effective_seed=seed, iterations=iterations)
    return run_simulation(_ref(), prepared)


def _tolerance(subject: str) -> tuple[float, float | None]:
    rows = [line for line in STEP0_RECORD.read_text(encoding="utf-8").splitlines()
            if line.startswith(f"| {subject} ")]
    assert len(rows) == 1, (subject, len(rows))
    relative = re.search(r"rel\s*≤\s*([0-9.eE+-]+)", rows[0])
    absolute = re.search(r"abs\s*≤\s*([0-9.eE+-]+)", rows[0])
    assert relative, rows[0]
    return float(relative.group(1)), (float(absolute.group(1)) if absolute else None)


def _cases() -> dict[str, dict]:
    if "cases" not in _CACHE:
        corpus = json.loads(CASES_JSON.read_text(encoding="utf-8"))
        _CACHE["cases"] = {
            case["id"]: case
            for group in corpus["groups"] for case in group["cases"]
        }
    return _CACHE["cases"]  # type: ignore[return-value]


# ===========================================================================
# A. The module exists, is declared, and exposes one entry point
# ===========================================================================
def test_01_the_module_exists_and_is_explicit() -> None:
    raw = _module().raw
    assert raw.startswith('Attribute VB_Name = "modSimEngine"')
    assert re.search(r"^Option Explicit\s*$", raw, re.M)
    assert not re.search(r"^Option Base\b", raw, re.M)


def test_02_the_module_is_registered_and_nothing_beyond_it() -> None:
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    modules = {m.name: m for m in structure.vba_modules}
    assert "modSimEngine" in modules
    assert modules["modSimEngine"].generated is False
    # THE PHASE-6 BLOCK, CONTIGUOUS AND IN ORDER. This was written as the LAST
    # eight entries, which was the same claim while Phase 6 was the last phase.
    # "Nothing beyond it" has since been settled by P7-2 landing
    # modSimSensitivity under its own authority; what still matters, and is
    # still checked, is that the accepted block is intact and unreordered.
    names = [m.name for m in structure.vba_modules]
    block = ['modSimContract', 'modSimRng', 'modSimSample', 'modSimEngine', 'modSimStats', 'modSimFingerprint', 'modSimNonce', 'modSimReport']
    at = names.index(block[0])
    assert names[at:at + len(block)] == block, names[at:at + len(block)]
    # No new endpoint, and D6-11 is untouched.
    assert "SimEngineRun" not in set(structure.entry_points) | set(structure.api_procedures)
    scoped = [(r.construct, tuple(r.allowed_in))
              for r in structure.forbidden_construct_rules if r.is_scoped]
    assert scoped == [("MRG32k3a", ("modSimRng",)),
                      ("RunSimulation", ("modSimReport",))], scoped
    endpoint = next(r for r in structure.forbidden_construct_rules
                    if r.construct == "RunSimulation")
    # SCOPED SINCE STEP 11, to its owner and to nothing else. This module is not
    # that owner, and the token still may not appear here.
    assert endpoint.allowed_in == ("modSimReport",)
    assert endpoint.forbidden_in("modSimEngine") is True
    assert "RunSimulation" not in _code()


def test_03_there_is_exactly_one_public_entry_point() -> None:
    """TWO NOW, AND BOTH ARE NAMED. P7-3 added per-driver replay, which is a
    second way IN and deliberately not a second way to compute: it reaches the
    same preparation, the same sampler and the same contribution routine. The
    set is asserted exactly, so a third entry point cannot appear unremarked."""
    assert _module().public_procedures == ["SimEngineRun", "SimEngineReplayDriver"]
    private = set(_module().procedures) - {"SimEngineRun", "SimEngineReplayDriver"}
    assert private == {
        "SimEnginePrepare", "SimEngineClaim", "SimEngineAdopt",
        "SimEngineValidateFactor", "SimEngineSampleValue", "SimEngineContribution",
    }, sorted(private)
    # The prepared representation is PRIVATE, so it creates no second writable
    # trust boundary of the kind Step 7 had to harden.
    raw = _module().raw
    assert re.findall(r"^(Public|Private) Type (\w+)$", raw, re.M) == [
        ("Private", "SimEngineDriver")
    ]


def test_04_the_entry_point_takes_resolved_factors_and_no_object() -> None:
    signature = logical_statements(_procedure("SimEngineRun"))[0][1]
    assert "ByRef drivers() As DriverFactors" in signature
    assert "ByVal driverCount As Long" in signature
    assert "ByVal effectiveSeed As Long" in signature
    assert "ByVal iterations As Long" in signature
    assert "ByRef totalNominal() As Double" in signature
    assert "ByRef totalPv() As Double" in signature
    for banned in ("As Object", "As Variant", "As Range", "As Worksheet",
                   "As Workbook", "As ListObject", "ParamArray"):
        assert banned not in signature, banned


def test_05_the_prepared_driver_type_holds_no_worksheet_anchor() -> None:
    fields = _transcribe()["_types"]["SimEngineDriver"]
    names = [f for f, _ in fields]
    assert names == [
        "PermanentId", "IsRisk", "DistKind", "MinValue", "MostLikely", "MaxValue",
        "Quantity", "Probability", "Knom", "Kpv", "BetaShape", "HasBetaShape",
        "ValueInitialState", "ValueStreamIndex", "OccurrenceInitialState",
        "OccurrenceStreamIndex", "HasOccurrenceStream",
    ], names
    kinds = dict(fields)
    assert kinds["BetaShape"] == "SimSampleBetaShape"
    assert kinds["ValueInitialState"] == "SimRngState"
    for name in names:
        for anchor in ("Row", "Sheet", "Address", "Range", "Cell", "Fx", "Inflation",
                       "Profile", "Object"):
            assert anchor not in name, (name, anchor)


# ===========================================================================
# B. Purity, and no leakage in either direction
# ===========================================================================
def test_06_the_whole_module_is_worksheet_independent() -> None:
    """Not merely the loop: the ENTIRE module never reaches a workbook."""
    code = _code()
    for token in ("Range", "Cells", "Worksheet", "Worksheets", "Workbook",
                  "Workbooks", "ListObject", "Application", "ThisWorkbook",
                  "ActiveWorkbook", "ActiveSheet", "Names(", "Evaluate", "MsgBox",
                  "InputBox", "CreateObject", "GetObject", "Environ", "Shell",
                  "DoEvents", "Timer", "Now", "Date", "Sheets("):
        assert token not in code, token


def test_07_there_is_no_module_level_or_static_state() -> None:
    inside, module_level = False, []
    for _, text in logical_statements(_module().code_without_string_removal):
        if re.match(r"^(Public|Private)\s+(Function|Sub)\b", text):
            inside = True
        elif re.match(r"^End (Function|Sub)$", text):
            inside = False
        elif not inside and re.match(r"^(Public|Private|Dim|Global)\s+\w+\s+As\s", text):
            module_level.append(text)
    assert module_level == [], module_level
    assert not re.search(r"\bStatic\b", _code())


def test_08_no_rng_implementation_leaked_in() -> None:
    code = _code()
    assert not contains_construct([_module()], "MRG32k3a")
    for banned in ("Rnd", "Randomize", "SimRngNextUniform", "SimRngJumpNextStream",
                   "SimRngStreamInitialState", "SimRngAutoSeedFromNonce",
                   "SimRngValidateState"):
        assert not re.search(rf"\b{banned}\b", code), banned
    # Exactly two modSimRng entry points, and no generator constant at all.
    called = set(re.findall(r"\bSimRng\w+", code))
    assert called == {"SimRngStateFromFixedSeed", "SimRngBuildComponentStreams",
                      "SimRngState", "SimRngComponent"}, sorted(called)
    for constant in _constants():
        if constant.startswith(("SIM_RNG_", "SIM_JUMP_", "SIM_SEED_", "SIM_NONCE_",
                                "SIM_AUTO_")):
            assert constant not in code, constant


def test_09_no_sampler_implementation_leaked_in() -> None:
    code = _code()
    for banned in ("Cheng", "Sqr(", "Log(", "Exp(", "alpha", "Alpha", "SIM_CHENG",
                   "SIM_PERT", "u <", "< probability", "InverseCdf", "vbTextCompare"):
        assert banned not in code, banned
    # The engine SELECTS a family and calls the accepted sampler.
    called = set(re.findall(r"\bSimSample\w+", code))
    assert called == {"SimSampleUniform", "SimSampleTriangular",
                      "SimSamplePreparedBeta", "SimSamplePrepareBetaPert",
                      "SimSampleBernoulli", "SimSampleBetaShape"}, sorted(called)
    # IT DOES NOT IMPLEMENT THE OCCURRENCE DECISION ITSELF. Every assignment to
    # `occurred` must be the literal default a driver WITHOUT an occurrence
    # stream takes; the real decision arrives ByRef from SimSampleBernoulli and
    # is never written here.
    assignments = {line.split("=", 1)[1].strip()
                   for line in code.splitlines()
                   if re.match(r"\s*occurred\s*=", line)}
    assert assignments <= {"True"}, sorted(assignments)


def test_10_no_statistic_digest_or_publication_exists() -> None:
    code = _code()
    for token in ("Mean", "StdDev", "StandardDeviation", "Variance", "Percentile",
                  "Quantile", "Median", "Contingency", "Digest", "Fingerprint",
                  "RunId", "run_id", "_SimData", "SimData", "Results", "Publish",
                  "RunSimulation", "Sensitivity", "Sort", "Nonce", "AttemptState"):
        assert token not in code, token
    assert not [p for p in _module().public_procedures if p.startswith("PCCM_")]


def test_11_no_model_resolution_leaked_in() -> None:
    code = _code()
    for token in ("BuildKnom", "BuildKpv", "BuildInflationFactors",
                  "BuildDiscountFactors", "Resolve", "Timeline", "FxRate",
                  "Profile", "Inflation", "Discount", "Currency"):
        assert token not in code, token
    # Only the accepted numerical primitives are borrowed from Phase 5.
    borrowed = set(re.findall(r"\b(?:Safe\w+|IsUsableDouble|MAX_DOUBLE)\b", code))
    assert borrowed == {"SafeProduct", "SafeSignedSum", "IsUsableDouble"}, sorted(borrowed)


# ===========================================================================
# C. Preflight, seeding and the zero-driver model
# ===========================================================================
def test_12_the_iteration_preflight_runs_before_anything_else() -> None:
    body = _procedure("SimEngineRun")
    assert "If iterations < SIM_MIN_ITERATIONS Then" in body
    assert "If iterations > SIM_MAX_ITERATIONS Then" in body
    # ...and before allocation, seeding, stream construction and any draw.
    first_check = body.index("SIM_MIN_ITERATIONS")
    for later in ("SimEnginePrepare", "ReDim ", "SimSample"):
        assert first_check < body.index(later), later
    # No other cap of any kind.
    assert "10000" not in body and "100000" not in body
    for literal in re.findall(r"(?<![\w.])\d{4,}", body):
        raise AssertionError(f"an invented iteration limit appeared: {literal}")


def test_13_an_out_of_range_iteration_count_is_refused_untouched() -> None:
    records = _factor_records(_model([_cost("CL-001")]))
    for iterations, fragment in (
        (_const("SIM_MIN_ITERATIONS") - 1, "below the accepted minimum"),
        (0, "below the accepted minimum"),
        (-5, "below the accepted minimum"),
        (_const("SIM_MAX_ITERATIONS") + 1, "above the technical maximum"),
    ):
        nominal, pv, detail = ["untouched"], ["untouched"], _Ref("")
        ok = _transcribe()["SimEngineRun"](
            records, _Ref(len(records)), _Ref(12345), _Ref(iterations),
            nominal, pv, detail)
        assert ok is False, iterations
        assert fragment in detail.v, (iterations, detail.v)
        assert nominal == ["untouched"] and pv == ["untouched"]
    # The accepted minimum itself is legal.
    ok, nominal, pv, detail = _run(records, iterations=_const("SIM_MIN_ITERATIONS"))
    assert ok, detail
    assert len(nominal) == _const("SIM_MIN_ITERATIONS")


def test_14_the_effective_seed_is_expanded_through_the_generator() -> None:
    body = _procedure("SimEnginePrepare")
    assert "SimRngStateFromFixedSeed(effectiveSeed, baseState, detail)" in body
    assert body.index("SimRngStateFromFixedSeed") < body.index("SimRngBuildComponentStreams")
    records = _factor_records(_model([_cost("CL-001")]))
    for seed in (_const("SIM_SEED_MIN") - 1, _const("SIM_SEED_MAX") + 1, 0, -1):
        nominal, pv, detail = ["untouched"], ["untouched"], _Ref("")
        assert _transcribe()["SimEngineRun"](
            records, _Ref(len(records)), _Ref(seed), _Ref(1000),
            nominal, pv, detail) is False, seed
        assert "admissible domain" in detail.v
        assert nominal == ["untouched"] and pv == ["untouched"]


def test_15_a_zero_driver_model_runs_and_retains_all_zero_totals() -> None:
    """No accepted contract requires a driver, and the engine invents no minimum."""
    iterations = _const("SIM_MIN_ITERATIONS")
    nominal, pv, detail = [], [], _Ref("")
    assert _transcribe()["SimEngineRun"](
        [], _Ref(0), _Ref(12345), _Ref(iterations), nominal, pv, detail) is True, detail.v
    assert len(nominal) == len(pv) == iterations
    assert all(value == 0.0 for value in nominal)
    assert all(value == 0.0 for value in pv)
    # ...and the accepted oracle says the same for the same empty model.
    expected = _oracle(_model(), iterations=iterations)
    assert list(expected.total_nominal) == nominal
    assert list(expected.total_pv) == pv


def test_16_the_zero_driver_path_reads_no_driver_array_bound() -> None:
    prepare = _procedure("SimEnginePrepare")
    seeded = prepare.index("SimRngStateFromFixedSeed")
    empty = prepare[prepare.rindex("Else", 0, seeded):seeded]
    for token in ("LBound", "UBound", "drivers("):
        assert token not in empty, token
    run = _procedure("SimEngineRun")
    staged = run.index("ReDim stagedNominal")
    zero = run[run.rindex("Else", 0, staged):staged]
    for token in ("LBound", "UBound", "drivers(", "prepared("):
        assert token not in zero, token
    # The one-slot carrier modSimRng returns is never inspected as a component.
    guard = prepare[prepare.index("If driverCount = 0 Then"):
                    prepare.index("ReDim prepared(0 To driverCount - 1)")]
    assert "components(" not in guard, guard
    assert "ReDim prepared(0 To 0)" in guard


def test_17_a_zero_driver_model_still_validates_its_seed() -> None:
    nominal, pv, detail = ["untouched"], ["untouched"], _Ref("")
    assert _transcribe()["SimEngineRun"](
        [], _Ref(0), _Ref(0), _Ref(1000), nominal, pv, detail) is False
    assert "admissible domain" in detail.v
    assert nominal == ["untouched"] and pv == ["untouched"]


# ===========================================================================
# D. SOURCE ARITHMETIC CONFORMANCE against the accepted oracle and corpus
# ===========================================================================
_ENGINE_TOTALS = "F1 per-iteration no-Beta end-to-end totals"


def _agrees(model: CalculationModel, seed=12345, iterations=1000, order=None):
    """Run both, and hand back (vba_nominal, vba_pv, oracle_result)."""
    ok, nominal, pv, detail = _run(_factor_records(model, order), seed, iterations)
    assert ok, detail
    return nominal, pv, _oracle(model, seed, iterations)


def test_18_the_unit_interval_case_is_exact() -> None:
    """engine.exact_friendly.unit_interval - EXACT.

    Quantity 1, Knom 1, Kpv 1, so each retained total IS the drawn uniform and
    the arithmetic is exactly representable by construction.
    """
    case = _cases()["engine.exact_friendly.unit_interval"]
    assert case["comparison"] == "EXACT"
    expected = case["expected_exact"]
    nominal, pv, _ = _agrees(_model([_cost("CL-001", "Uniform", 0.0, None, 1.0, 1.0)]))
    assert len(nominal) == expected["total_nominal"]["count"] == 1000
    assert nominal[: len(expected["total_nominal"]["head"])] == \
        expected["total_nominal"]["head"]
    assert nominal[-len(expected["total_nominal"]["tail"]):] == \
        expected["total_nominal"]["tail"]
    assert len(set(nominal)) == expected["total_nominal"]["distinct_count"]
    assert pv[: len(expected["total_pv"]["head"])] == expected["total_pv"]["head"]
    assert pv[-len(expected["total_pv"]["tail"]):] == expected["total_pv"]["tail"]


def test_19_the_dyadic_mixed_case_is_exact() -> None:
    """engine.exact_friendly.dyadic_mixed - EXACT. Dyadic values, exact totals."""
    assert _cases()["engine.exact_friendly.dyadic_mixed"]["comparison"] == "EXACT"
    model = _model(
        [_cost("CL-001", "Triangular", 0.25, 0.25, 0.25, 4.0),
         _cost("CL-002", "Uniform", 0.5, None, 0.5, 2.0)],
        [_risk("R-001", "Triangular", 8.0, 8.0, 8.0, probability=1.0)],
    )
    nominal, pv, expected = _agrees(model)
    assert nominal == list(expected.total_nominal)
    assert pv == list(expected.total_pv)
    # Every driver is degenerate and the Risk is certain, so every iteration is
    # the same exact sum: 0.25*4 + 0.5*2 + 8 = 10, with no rounding anywhere.
    assert set(nominal) == {10.0}, sorted(set(nominal))[:4]
    assert set(pv) == {10.0}


def test_20_quantity_is_applied_exactly_once_and_the_total_is_linear() -> None:
    """engine.cost_line.quantity_applied_once - EXACT."""
    case = _cases()["engine.cost_line.quantity_applied_once"]
    assert case["comparison"] == "EXACT"
    unit_cost = case["inputs"]["unit_cost"]
    seen = []
    for row in case["expected_exact"]["rows"]:
        quantity = row["quantity"]
        model = _model([_cost("CL-001", "Uniform", unit_cost, None, unit_cost, quantity)])
        nominal, pv, expected = _agrees(model)
        assert set(nominal) == {row["total"]}, (quantity, sorted(set(nominal))[:3])
        assert set(pv) == {row["total"]}
        assert nominal == list(expected.total_nominal)
        seen.append((quantity, row["total"], row["applied_twice_would_be"]))
    # LINEAR, not quadratic: the corpus states what applying it twice would give.
    for quantity, total, twice in seen:
        assert total == unit_cost * quantity
        if quantity != 1.0:
            assert total != twice, quantity
    # And the SAMPLE support is unit cost, not total cost. The loop hands the
    # sampled unit cost to the shared routine, which is where Quantity is
    # applied - once.
    assert "unitCost" in _loop_body()
    assert "factors(1) = prepared.Quantity" in _contribution()


def test_21_d6_18b_severity_is_sampled_whatever_the_occurrence_decided() -> None:
    """engine.risk.d6_18b_unconditional_severity - EXACT.

    Two runs differing ONLY in Probability. The Bernoulli outcomes differ; the
    severity stream is untouched by that difference.
    """
    case = _cases()["engine.risk.d6_18b_unconditional_severity"]
    assert case["comparison"] == "EXACT"
    seed = case["inputs"]["seed"]
    iterations = case["inputs"]["iterations"]
    drawn: dict[float, list[float]] = {}
    for run in case["expected_exact"]["runs"]:
        probability = run["probability"]
        model = _model(
            [_cost("CL-001", "Uniform", 10.0, None, 10.0, 1.0)],
            [_risk("R-001", "Triangular", 100.0, 200.0, 400.0, probability=probability)],
        )
        vba = _transcribe()
        real = vba["SimSampleTriangular"]
        severities: list[float] = []

        def watched(state, a, m, b, sample, consumed, detail, _real=real, _out=severities):
            ok = _real(state, a, m, b, sample, consumed, detail)
            if ok:
                _out.append(sample.v)
            return ok

        vba["SimSampleTriangular"] = watched
        try:
            ok, nominal, pv, detail = _run(_factor_records(model), seed, iterations)
        finally:
            vba["SimSampleTriangular"] = real
        assert ok, detail
        assert len(severities) == run["severity_uniforms_consumed"] == iterations, probability
        occurrences = sum(1 for value in nominal if value != 10.0)
        assert occurrences == run["occurrences"], probability
        drawn[probability] = severities
        assert nominal == list(_oracle(model, seed, iterations).total_nominal)
    # PROBABILITY-ONLY COMPARABILITY. The severity sequence is a function of the
    # seed and the distribution alone: identical, draw for draw.
    keys = sorted(drawn)
    assert keys == [0.2, 0.8]
    assert drawn[keys[0]] == drawn[keys[1]], "the severity sequence moved with Probability"
    assert case["expected_exact"]["runs"][0]["occurrences"] != \
        case["expected_exact"]["runs"][1]["occurrences"]


def test_22_a_degenerate_severity_is_invoked_and_consumes_nothing() -> None:
    """engine.risk.degenerate_severity_zero_consumption - EXACT."""
    case = _cases()["engine.risk.degenerate_severity_zero_consumption"]
    assert case["comparison"] == "EXACT"
    expected = case["expected_exact"]
    model = _model([], [_risk("R-001", "Triangular", 90.0, 90.0, 90.0, probability=0.4)])
    vba = _transcribe()
    real = vba["SimSampleTriangular"]
    invocations = [0]

    def counted(*args, _real=real):
        invocations[0] += 1
        return _real(*args)

    vba["SimSampleTriangular"] = counted
    try:
        ok, nominal, pv, detail = _run(_factor_records(model), 12345, 1000)
    finally:
        vba["SimSampleTriangular"] = real
    assert ok, detail
    # INVOKED every iteration...
    assert invocations[0] == 1000
    # ...and the severity component consumed nothing, so its stream stood still.
    severity = next(c for c in expected["components"] if c["role"] == "severity")
    assert severity["uniforms_consumed"] == 0
    assert severity["initial_state"] == severity["final_state"]
    assert sorted(set(nominal)) == expected["distinct_totals"]
    assert nominal == list(_oracle(model, 12345, 1000).total_nominal)


def test_23_the_general_no_beta_case_is_tolerance_bounded() -> None:
    """engine.general.no_beta - TOLERANCE_BOUNDED, under the accepted bound."""
    assert _cases()["engine.general.no_beta"]["comparison"] == "TOLERANCE_BOUNDED"
    relative, absolute = _tolerance(_ENGINE_TOTALS)
    model = _model(
        [_cost("CL-001", "Triangular", 80.0, 100.0, 150.0, 3.0),
         _cost("CL-002", "Uniform", 40.0, None, 90.0, 1.5)],
        [_risk("R-001", "Triangular", 100.0, 200.0, 400.0, probability=0.3),
         _risk("R-002", "Uniform", 50.0, None, 250.0, probability=0.75)],
    )
    nominal, pv, expected = _agrees(model)
    for actual, want in zip(nominal, expected.total_nominal):
        scale = max(abs(actual), abs(want))
        assert abs(actual - want) <= max(relative * abs(want), absolute * scale)
    for actual, want in zip(pv, expected.total_pv):
        scale = max(abs(actual), abs(want))
        assert abs(actual - want) <= max(relative * abs(want), absolute * scale)


def test_24_the_with_beta_case_is_compared_distributionally() -> None:
    """engine.general.with_beta - STATISTICAL, and NOT strengthened to EXACT.

    A rejection sampler can legitimately desynchronise two implementations, so
    the corpus refuses sample-for-sample identity here and so does this test.
    """
    assert _cases()["engine.general.with_beta"]["comparison"] == "STATISTICAL"
    model = _model(
        [_cost("CL-001", "Beta-PERT", 0.0, 25.0, 100.0, 2.0)],
        [_risk("R-001", "Beta-PERT", 1.0, 2.0, 9.0, probability=0.4)],
    )
    nominal, pv, expected = _agrees(model)
    assert len(nominal) == len(expected.total_nominal) == 1000
    got = sorted(nominal)
    want = sorted(expected.total_nominal)
    for fraction in (0.05, 0.25, 0.5, 0.75, 0.95):
        position = int(fraction * (len(got) - 1))
        spread = want[-1] - want[0]
        assert abs(got[position] - want[position]) <= 0.02 * spread, fraction
    assert abs(sum(nominal) / len(nominal)
               - sum(expected.total_nominal) / len(expected.total_nominal)) \
        <= 0.02 * (want[-1] - want[0])
    assert all(math.isfinite(value) for value in nominal + pv)


def test_25_replay_and_seed_divergence_are_same_runtime_relations() -> None:
    """engine.replay.same_seed_identical / engine.seed.non_degenerate_divergence."""
    assert _cases()["engine.replay.same_seed_identical"]["comparison"] == "SAME_RUNTIME_ONLY"
    assert _cases()["engine.seed.non_degenerate_divergence"]["comparison"] == "SAME_RUNTIME_ONLY"
    model = _model([_cost("CL-001")], [_risk("R-001", probability=0.4)])
    records = _factor_records(model)
    first = _run(records, 12345, 1000)
    second = _run(records, 12345, 1000)
    assert first[0] and second[0]
    assert first[1] == second[1] and first[2] == second[2], "the same seed did not replay"
    other = _run(records, 54321, 1000)
    assert other[0]
    assert other[1] != first[1], "two different seeds produced the same run"


def test_26_a_fully_degenerate_fixture_is_seed_independent() -> None:
    """engine.seed.degenerate_equal_digest - EXACT, on the retained arrays only.

    The digest itself belongs to Step 10; what Step 8 can state is that the
    RETAINED TOTALS are identical for every seed, which is what the digest of
    that case is computed from.
    """
    assert _cases()["engine.seed.degenerate_equal_digest"]["comparison"] == "EXACT"
    model = _model(
        [_cost("CL-001", "Triangular", 40.0, 40.0, 40.0, 2.0)],
        [_risk("R-001", "Uniform", 15.0, None, 15.0, probability=1.0)],
    )
    records = _factor_records(model)
    runs = [_run(records, seed, 1000) for seed in (1, 12345, 2147483646)]
    for ok, nominal, pv, detail in runs:
        assert ok, detail
        assert set(nominal) == {95.0}, sorted(set(nominal))[:3]
    assert runs[0][1] == runs[1][1] == runs[2][1]
    assert runs[0][2] == runs[1][2] == runs[2][2]


def test_27_the_physical_register_order_reaches_nothing() -> None:
    """engine.row_order.invariant - SAME_RUNTIME_ONLY. Three permutations."""
    assert _cases()["engine.row_order.invariant"]["comparison"] == "SAME_RUNTIME_ONLY"
    model = _model(
        [_cost("CL-003"), _cost("CL-001", "Uniform", 10.0, None, 60.0, 2.0), _cost("CL-002")],
        [_risk("R-002", probability=0.3), _risk("R-001", "Uniform", 5.0, None, 25.0, 0.7)],
    )
    baseline = None
    for order in ([0, 1, 2, 3, 4], [4, 3, 2, 1, 0], [2, 0, 4, 1, 3], [1, 4, 0, 3, 2]):
        ok, nominal, pv, detail = _run(_factor_records(model, order), 12345, 1000)
        assert ok, (order, detail)
        if baseline is None:
            baseline = (nominal, pv)
        assert (nominal, pv) == baseline, order
    assert baseline[0] == list(_oracle(model, 12345, 1000).total_nominal)


def test_28_cl_1000_sorts_before_cl_999() -> None:
    """The ordering is ordinal, and nothing parses a numeric suffix."""
    model = _model([
        _cost("CL-999", "Uniform", 1.0, None, 1.0, 1.0),
        _cost("CL-1000", "Uniform", 2.0, None, 2.0, 1.0),
        _cost("CL-100", "Uniform", 4.0, None, 4.0, 1.0),
    ])
    nominal, _, expected = _agrees(model)
    assert nominal == list(expected.total_nominal)
    # The engine reads the canonical order OFF the component sequence, so there
    # is one collation implementation in the phase and it is modSimRng's.
    prepare = _procedure("SimEnginePrepare")
    assert "SimRngBuildComponentStreams" in prepare
    for token in ("Val(", "Right(", "Left(", "Mid(", "LCase", "UCase", "Trim",
                  "vbTextCompare", "InStr", "Replace(", "Split("):
        assert token not in _code(), token
    assert _code().count("vbBinaryCompare") == 1


# ===========================================================================
# E. Factor semantics - each kind reads only what it owns
# ===========================================================================
def _poisoned(records, permanent_id, field, value):
    out = [dict(record) for record in records]
    for record in out:
        if record["PermanentId"] == permanent_id:
            record[field] = value
    return out


def test_29_a_cost_lines_probability_is_never_read() -> None:
    model = _model([_cost("CL-001", "Uniform", 10.0, None, 60.0, 2.0)],
                   [_risk("R-001", probability=0.4)])
    records = _factor_records(model)
    baseline = _run(records)
    assert baseline[0], baseline[3]
    for poison in (0.5, -7.0, 12.0, float("nan"), float("inf")):
        ok, nominal, pv, detail = _run(_poisoned(records, "CL-001", "Probability", poison))
        assert ok, (poison, detail)
        assert (nominal, pv) == (baseline[1], baseline[2]), poison
    assert "prepared(index).Probability" not in _procedure("SimEngineRun").split(
        "For index = costCount To driverCount - 1")[0]


def test_30_a_risks_quantity_is_never_read() -> None:
    model = _model([_cost("CL-001", "Uniform", 10.0, None, 60.0, 2.0)],
                   [_risk("R-001", probability=0.4)])
    records = _factor_records(model)
    baseline = _run(records)
    assert baseline[0], baseline[3]
    for poison in (3.0, -1.0, float("nan"), float("inf")):
        ok, nominal, pv, detail = _run(_poisoned(records, "R-001", "Quantity", poison))
        assert ok, (poison, detail)
        assert (nominal, pv) == (baseline[1], baseline[2]), poison
    # The Risk arm of the loop carries no Quantity factor at all.
    risk_arm = _loop_body().split("For index = costCount To driverCount - 1")[1]
    assert "Quantity" not in risk_arm


def test_31_a_uniforms_most_likely_is_never_read() -> None:
    model = _model([_cost("CL-001", "Uniform", 10.0, None, 60.0, 2.0)])
    records = _factor_records(model)
    baseline = _run(records)
    assert baseline[0], baseline[3]
    for poison in (35.0, -1e300, 1e300, float("nan"), float("inf")):
        ok, nominal, pv, detail = _run(_poisoned(records, "CL-001", "MostLikely", poison))
        assert ok, (poison, detail)
        assert (nominal, pv) == (baseline[1], baseline[2]), poison
    # Not adopted, not validated, and SimSampleUniform has no parameter for it.
    adopt = _procedure("SimEngineAdopt")
    assert "If factor.DistKind <> DIST_UNIFORM Then" in adopt
    assert adopt.index("If factor.DistKind <> DIST_UNIFORM Then") < \
        adopt.index("target.MostLikely = factor.MostLikely")
    assert "prepared.MostLikely" not in _procedure("SimEngineSampleValue").split(
        "ElseIf")[0]


def test_32_the_central_and_mean_fields_are_not_inspected() -> None:
    """They belong to the analytical layer, and this engine is not it."""
    for token in ("Central", "MeanValue", "CentralBasis"):
        assert token not in _code(), token


# ===========================================================================
# F. Accumulation
# ===========================================================================
def test_33_nominal_and_pv_are_independent_accumulators() -> None:
    body = _loop_body()
    assert body.count("SafeSignedSum") == 2
    assert "SafeSignedSum(nominalTerm, driverCount, measured)" in body
    assert "SafeSignedSum(pvTerm, driverCount, measured)" in body
    # PV is never derived from the nominal term that was just computed.
    assert "pvTerm(index) = nominalTerm(index)" not in body
    for derived in ("nominalTerm(index) *", "stagedPv(iteration - 1) = stagedNominal",
                    "measured *", "* discount", "/ (1"):
        assert derived not in body, derived
    # Both use the same canonical driver order and the same logical count.
    assert body.count("prepared(index).Knom") == 2
    assert body.count("prepared(index).Kpv") == 2


def test_34_the_accepted_primitives_are_used_and_not_replaced() -> None:
    body = _loop_body()
    # FOUR CONTRIBUTIONS PER ITERATION, cost and risk, nominal and PV - and all
    # four now form their product in the one shared routine, which calls the
    # accepted primitive exactly once.
    assert body.count("SimEngineContribution(") == 4
    assert _contribution().count("SafeProduct(factors, count, term)") == 1
    assert "SafeProduct(" not in body, "a product is formed outside the shared routine"
    assert "SafeAccumulate" not in _code()
    # No naive running total, and no chained multiplication.
    for naive in ("= total +", "total = total", "measured = measured +",
                  "unitCost *", "severity *", "* prepared(index).Knom"):
        assert naive not in body, naive
    for naive in ("sample *", "* factor", "* prepared.Quantity"):
        assert naive not in _contribution(), naive


def test_35_the_accumulation_order_is_canonical_and_matters() -> None:
    """A constructed non-associative fixture, proved non-associative FIRST.

    Reversing the term order is only a detectable mutation on a fixture where
    binary64 addition is actually non-associative, so the fixture is validated
    through the independent Python primitive before the engine is asked about it.
    """
    canonical = [1.0, 1e16, -1e16, 1.0]
    # The transformation validated here is the one an order mutation performs:
    # the COST terms reversed, the Risk terms left where they are.
    disturbed = [-1e16, 1e16, 1.0, 1.0]
    assert safe_signed_sum(canonical, "fixture") != safe_signed_sum(disturbed, "fixture"), (
        "the fixture is associative, so an order mutation could not be detected"
    )
    assert safe_signed_sum(canonical, "fixture") == 1.0
    assert safe_signed_sum(disturbed, "fixture") == 2.0
    # The engine's own order is Cost Lines then Risks, ascending Permanent ID,
    # and it is read off the component sequence rather than re-derived. The
    # drivers are laid out so the CANONICAL term sequence is exactly the fixture
    # above, and they are handed over in a different physical order.
    model = _model(
        [_cost("CL-002", "Uniform", 1e16, None, 1e16, 1.0),
         _cost("CL-003", "Uniform", -1e16, None, -1e16, 1.0),
         _cost("CL-001", "Uniform", 1.0, None, 1.0, 1.0)],
        [_risk("R-001", "Uniform", 1.0, None, 1.0, probability=1.0)],
    )
    nominal, pv, expected = _agrees(model)
    assert nominal == list(expected.total_nominal)
    assert set(nominal) == {safe_signed_sum(canonical, "fixture")} == {1.0}
    assert set(pv) == {1.0}


def test_36_the_term_arrays_are_allocated_once_and_indexed_canonically() -> None:
    run = _procedure("SimEngineRun")
    body = _loop_body()
    assert "ReDim" not in body, "an allocation appeared inside the iteration loop"
    for name in ("nominalTerm", "pvTerm", "valueState", "occurrenceState",
                 "stagedNominal", "stagedPv"):
        assert f"ReDim {name}(" in run, name
        allocations = [m.start() for m in re.finditer(rf"ReDim {name}\(", run)]
        assert all(position < run.index("For iteration = 1 To iterations")
                   for position in allocations), name
    # The zero-driver carrier is sized but never written as a term.
    assert "SafeSignedSum(nominalTerm, driverCount, measured)" in body


def test_37_the_retained_output_is_exactly_two_iteration_arrays() -> None:
    run = _procedure("SimEngineRun")
    assert run.count("totalNominal = stagedNominal") == 1
    assert run.count("totalPv = stagedPv") == 1
    # Committed LAST, after the loop, and nothing else is retained.
    assert run.index("Next iteration") < run.index("totalNominal = stagedNominal")
    assert "stagedNominal(iteration - 1) = measured" in run
    assert "stagedPv(iteration - 1) = measured" in run
    # No per-driver-by-iteration matrix, no sample history.
    for retained in ("history", "History", "samples(", "perDriver", "matrix", "Matrix"):
        assert retained not in _code(), retained
    ok, nominal, pv, detail = _run(_factor_records(_model([_cost("CL-001")])), 12345, 1000)
    assert ok and len(nominal) == len(pv) == 1000
    # Zero-based, element i - 1 holding iteration i, as every Phase-6 array is.
    assert "ReDim stagedNominal(0 To iterations - 1)" in run


# ===========================================================================
# G. Preparation, the hot loop, and the Step-7 carry-forward
# ===========================================================================
def test_38_nothing_is_prepared_or_allocated_inside_the_loop() -> None:
    body = _loop_body()
    for banned in ("ReDim", "SimRngBuildComponentStreams", "SimRngStateFromFixedSeed",
                   "SimSamplePrepareBetaPert", "SimRngJumpNextStream", "StrComp",
                   "Sqr(", "Log(", "Exp(", "SimEngineClaim", "SimEngineAdopt",
                   "SimEngineValidateFactor", "SimEnginePrepare"):
        assert banned not in body, banned
    # ...and each of those DOES happen exactly once, before it.
    prepare = _procedure("SimEnginePrepare")
    for once in ("SimRngStateFromFixedSeed", "SimRngBuildComponentStreams",
                 "SimSamplePrepareBetaPert"):
        assert prepare.count(once) == 1, once


def test_39_a_beta_shape_comes_only_from_the_step_7_constructor() -> None:
    """The Step-7 carry-forward: the engine never assembles a shape itself."""
    code = _code()
    # Exactly one site produces one, and it is the accepted constructor.
    assert code.count("SimSamplePrepareBetaPert") == 1
    assert "SimSamplePrepareBetaPert(prepared(index).MinValue, prepared(index).MostLikely," in \
        re.sub(r"\s*_\s*\n\s*", " ", _procedure("SimEnginePrepare")).replace(
            "SimSamplePrepareBetaPert(prepared(index).MinValue, prepared(index).MostLikely, ",
            "SimSamplePrepareBetaPert(prepared(index).MinValue, prepared(index).MostLikely,")
    # NO FIELD OF A SHAPE IS EVER WRITTEN, here or anywhere in the module.
    assert not re.search(r"\.BetaShape\.\w+\s*=[^=]", code), "a Beta shape field was written"
    for field in ("Alpha", "Beta", "ChengA", "ChengB", "ChengAlpha", "ChengBeta",
                  "ChengGamma", "ChengDelta", "ChengK1", "ChengK2", "Degenerate",
                  "UseChengBB", "FirstParameterIsOrientedA", "Prepared"):
        assert f".BetaShape.{field}" not in code, field
    # The only assignment mentioning BetaShape at all is the HasBetaShape flag.
    writes = [statement for _, statement in logical_statements(code)
              if re.match(r"^\s*\w[\w.()]*BetaShape\w*\s*=[^=]", statement)]
    assert writes == ["prepared(index).HasBetaShape = True"], writes
    # And in the loop it is only READ, by the accepted prepared sampler.
    sampler = _procedure("SimEngineSampleValue")
    assert "SimSamplePreparedBeta(state, prepared.BetaShape, sample, consumed," in \
        re.sub(r"\s*_\s*\n\s*", " ", sampler)


def test_40_beta_preparation_happens_once_per_driver_not_per_iteration() -> None:
    vba = _transcribe()
    real = vba["SimSamplePrepareBetaPert"]
    calls = [0]

    def counted(*args, _real=real):
        calls[0] += 1
        return _real(*args)

    vba["SimSamplePrepareBetaPert"] = counted
    try:
        model = _model([_cost("CL-001", "Beta-PERT", 0.0, 25.0, 100.0, 2.0),
                        _cost("CL-002", "Uniform", 1.0, None, 2.0, 1.0)],
                       [_risk("R-001", "Beta-PERT", 1.0, 2.0, 9.0, probability=0.4)])
        ok, nominal, pv, detail = _run(_factor_records(model), 12345, 1000)
    finally:
        vba["SimSamplePrepareBetaPert"] = real
    assert ok, detail
    # TWO Beta drivers, 1000 iterations, and exactly two preparations.
    assert calls[0] == 2, calls[0]


def test_41_the_prepared_initial_states_are_never_mutated() -> None:
    """A run copies them; two runs of the same model replay identically."""
    run = _procedure("SimEngineRun")
    assert "valueState(index) = prepared(index).ValueInitialState" in run
    assert "occurrenceState(index) = prepared(index).OccurrenceInitialState" in run
    body = _loop_body()
    for mutation in ("prepared(index).ValueInitialState =",
                     "prepared(index).OccurrenceInitialState ="):
        assert mutation not in body, mutation
    assert "InitialState" not in body
    records = _factor_records(_model([_cost("CL-001")], [_risk("R-001")]))
    first = _run(records, 12345, 1000)
    second = _run(records, 12345, 1000)
    assert first[1] == second[1] and first[2] == second[2]


def test_42_every_component_identity_is_verified_before_use() -> None:
    claim = _procedure("SimEngineClaim")
    for check in ("If component.DriverKind <> wantKind Then",
                  "If component.Role <> wantRole Then",
                  "If component.StreamIndex <> SIM_STREAM_INDEX_ORIGIN + wantIndex Then",
                  "If Len(component.PermanentId) = 0 Then",
                  "If claimed(found) Then"):
        assert check in claim, check
    assert "two driver records claim one component identity" in claim
    assert "a component identity matches no driver record" in claim
    prepare = _procedure("SimEnginePrepare")
    for pair in ("If components(severity).PermanentId <> components(occurrence).PermanentId Then",
                 "If components(severity).DriverKind <> SIM_COMPONENT_3_DRIVER_KIND Then",
                 "If components(severity).Role <> SIM_COMPONENT_3_ROLE Then",
                 "If components(severity).StreamIndex <> SIM_STREAM_INDEX_ORIGIN + severity Then",
                 "If components(severity).StreamIndex = components(occurrence).StreamIndex Then"):
        assert pair in prepare, pair
    # occurrence THEN severity, adjacent.
    assert "occurrence = costCount + 2 * index" in prepare
    assert "severity = occurrence + 1" in prepare


# ===========================================================================
# H. Failure and transactional output
# ===========================================================================
def _refuse(records, fragment, seed=12345, iterations=1000):
    nominal, pv, detail = ["untouched"], ["untouched"], _Ref("")
    ok = _transcribe()["SimEngineRun"](
        records, _Ref(len(records)), _Ref(seed), _Ref(iterations), nominal, pv, detail)
    assert ok is False, (fragment, detail.v)
    assert fragment in detail.v, (fragment, detail.v)
    assert nominal == ["untouched"] and pv == ["untouched"], fragment
    return detail.v


def test_43_an_invalid_distribution_or_support_is_refused_before_any_draw() -> None:
    records = _factor_records(_model([_cost("CL-001")], [_risk("R-001")]))
    _refuse(_poisoned(records, "CL-001", "DistKind", 0), "unknown distribution")
    _refuse(_poisoned(records, "CL-001", "DistKind", 4), "unknown distribution")
    _refuse(_poisoned(records, "CL-001", "MostLikely", 999.0),
            "Min <= Most Likely <= Max")
    _refuse(_poisoned(records, "R-001", "MinValue", 1e9), "Min <= Most Likely <= Max")
    _refuse(_poisoned(records, "CL-001", "MinValue", float("nan")), "not a finite Double")
    _refuse(_poisoned(records, "CL-001", "Knom", float("inf")), "Knom is not a finite Double")
    _refuse(_poisoned(records, "CL-001", "Kpv", float("nan")), "Kpv is not a finite Double")
    _refuse(_poisoned(records, "CL-001", "Quantity", float("nan")),
            "Quantity is not a finite Double")
    _refuse(_poisoned(records, "CL-001", "PermanentId", ""), "blank permanent id")
    # A Uniform whose endpoints are the wrong way round is refused, not swapped.
    uniform = _factor_records(_model([_cost("CL-001", "Uniform", 10.0, None, 60.0, 1.0)]))
    _refuse(_poisoned(uniform, "CL-001", "MinValue", 90.0),
            "Min <= Max; the ordering is refused, not repaired")


def test_44_an_invalid_probability_is_refused_and_never_clamped() -> None:
    records = _factor_records(_model([_cost("CL-001")], [_risk("R-001")]))
    for poison in (-0.1, 1.1, float("nan"), float("inf")):
        detail = _refuse(_poisoned(records, "R-001", "Probability", poison), "Probability")
        assert "clamped" in detail or "finite" in detail, (poison, detail)
    for legal in (0.0, 1.0):
        ok, nominal, pv, refusal = _run(_poisoned(records, "R-001", "Probability", legal))
        assert ok, (legal, refusal)


def test_45_a_component_identity_mismatch_is_refused() -> None:
    """The mapping is verified, not assumed."""
    vba = _transcribe()
    real = vba["SimRngBuildComponentStreams"]

    def corrupt(field, value, only_role=None):
        def wrapped(costIds, costCount, riskIds, riskCount, baseState, components, detail):
            if not real(costIds, costCount, riskIds, riskCount, baseState, components, detail):
                return False
            for entry in components:
                if only_role is None or entry["Role"] == only_role:
                    entry[field] = value
                    break
            return True
        return wrapped

    records = _factor_records(_model([_cost("CL-001")], [_risk("R-001")]))
    for corruption, fragment in (
        (corrupt("PermanentId", "CL-999"), "matches no driver record"),
        (corrupt("Role", "severity"), "wrong role"),
        (corrupt("DriverKind", "risk"), "wrong driver kind"),
        (corrupt("StreamIndex", 7), "wrong stream index"),
        (corrupt("PermanentId", ""), "blank permanent id"),
    ):
        vba["SimRngBuildComponentStreams"] = corruption
        try:
            _refuse(records, fragment)
        finally:
            vba["SimRngBuildComponentStreams"] = real


def test_46_a_swapped_occurrence_severity_pair_is_refused() -> None:
    vba = _transcribe()
    real = vba["SimRngBuildComponentStreams"]

    def swapper(costIds, costCount, riskIds, riskCount, baseState, components, detail):
        if not real(costIds, costCount, riskIds, riskCount, baseState, components, detail):
            return False
        first = next(i for i, c in enumerate(components) if c["Role"] == "occurrence")
        components[first], components[first + 1] = components[first + 1], components[first]
        return True

    records = _factor_records(_model([_cost("CL-001")], [_risk("R-001")]))
    vba["SimRngBuildComponentStreams"] = swapper
    try:
        _refuse(records, "wrong role")
    finally:
        vba["SimRngBuildComponentStreams"] = real

    # ...and an occurrence borrowed from one Risk beside another Risk's severity.
    def crosser(costIds, costCount, riskIds, riskCount, baseState, components, detail):
        if not real(costIds, costCount, riskIds, riskCount, baseState, components, detail):
            return False
        first = next(i for i, c in enumerate(components) if c["Role"] == "occurrence")
        components[first + 1]["PermanentId"] = "R-002"
        return True

    records = _factor_records(_model([], [_risk("R-001"), _risk("R-002")]))
    vba["SimRngBuildComponentStreams"] = crosser
    try:
        _refuse(records, "belong to different risks")
    finally:
        vba["SimRngBuildComponentStreams"] = real


def test_47_a_refusal_mid_run_leaves_the_caller_arrays_untouched() -> None:
    """No partial success, at iteration 1, at 500 or at the last one."""
    vba = _transcribe()
    real = vba["SimSampleTriangular"]
    for failing_call in (1, 500, 1000):
        calls = [0]

        def failing(state, a, m, b, sample, consumed, detail, _n=failing_call, _real=real):
            calls[0] += 1
            if calls[0] == _n:
                detail.v = "sampler: injected failure"
                return False
            return _real(state, a, m, b, sample, consumed, detail)

        vba["SimSampleTriangular"] = failing
        try:
            records = _factor_records(_model([_cost("CL-001")]))
            nominal, pv, detail = ["untouched"], ["untouched"], _Ref("")
            ok = vba["SimEngineRun"](
                records, _Ref(len(records)), _Ref(12345), _Ref(1000), nominal, pv, detail)
        finally:
            vba["SimSampleTriangular"] = real
        assert ok is False, failing_call
        assert "injected failure" in detail.v
        assert f"iteration {failing_call}" in detail.v, (failing_call, detail.v)
        assert "CL-001" in detail.v
        assert nominal == ["untouched"] and pv == ["untouched"], failing_call


def test_48_a_numerical_refusal_names_the_driver_and_the_measure() -> None:
    huge = sys.float_info.max
    records = _factor_records(_model([_cost("CL-001", "Uniform", huge, None, huge, 1.0)]))
    poisoned = _poisoned(records, "CL-001", "Knom", huge)
    detail = _refuse(poisoned, "CL-001")
    assert "nominal contribution is not representable" in detail
    assert "iteration 1" in detail
    # ...and a total that cannot be represented names the iteration and measure.
    vba = _transcribe()
    real = vba["SafeSignedSum"]
    vba["SafeSignedSum"] = lambda terms, count, result: False
    try:
        detail = _refuse(_factor_records(_model([_cost("CL-001")])), "the nominal total")
    finally:
        vba["SafeSignedSum"] = real
    assert "iteration 1" in detail


def test_49_the_transcription_read_the_whole_module() -> None:
    vba = _transcribe()
    for name in _module().procedures:
        assert callable(vba[name]), name
    assert "SimEngineDriver" in vba["_types"]
    assert "DriverFactors" in vba["_types"]
    # The accepted modules it consumes are compiled from their own source.
    for name in ("SimRngStateFromFixedSeed", "SimRngBuildComponentStreams",
                 "SimSampleUniform", "SimSampleTriangular", "SimSamplePreparedBeta",
                 "SimSamplePrepareBetaPert", "SimSampleBernoulli", "IsUsableDouble"):
        assert callable(vba[name]), name


def test_50_bernoulli_runs_once_per_risk_per_iteration_and_first() -> None:
    """D6-18b ordering: the occurrence draw precedes the severity sampler."""
    vba = _transcribe()
    real_bernoulli = vba["SimSampleBernoulli"]
    real_triangular = vba["SimSampleTriangular"]
    trace: list[str] = []

    def watched_bernoulli(*args, _real=real_bernoulli):
        trace.append("occurrence")
        return _real(*args)

    def watched_triangular(*args, _real=real_triangular):
        trace.append("severity")
        return _real(*args)

    vba["SimSampleBernoulli"] = watched_bernoulli
    vba["SimSampleTriangular"] = watched_triangular
    try:
        model = _model([], [_risk("R-001", probability=0.0),
                            _risk("R-002", probability=1.0)])
        ok, nominal, pv, detail = _run(_factor_records(model), 12345, 1000)
    finally:
        vba["SimSampleBernoulli"] = real_bernoulli
        vba["SimSampleTriangular"] = real_triangular
    assert ok, detail
    # Two Risks, 1000 iterations: 2000 occurrence draws and 2000 severity draws,
    # strictly alternating, occurrence first - at p = 0 and at p = 1 alike.
    assert trace.count("occurrence") == 2000
    assert trace.count("severity") == 2000
    assert trace == ["occurrence", "severity"] * 2000
    # p = 0 never contributes, p = 1 always does, and neither skipped a draw.
    assert all(value > 0.0 for value in nominal)
    assert nominal == list(_oracle(model, 12345, 1000).total_nominal)
    # The source says the same: occurrence before severity, one call each.
    risk_arm = _loop_body().split("For index = costCount To driverCount - 1")[1]
    assert risk_arm.count("SimSampleBernoulli") == 1
    assert risk_arm.count("SimEngineSampleValue") == 1
    assert risk_arm.index("SimSampleBernoulli") < risk_arm.index("SimEngineSampleValue")
    # THE SEVERITY CALL SITS INSIDE NO CONDITIONAL AT ALL. P7-3 moved the
    # occurrence test out of the loop with the contribution arithmetic, so the
    # stronger statement is now also the simpler one: there is no `If occurred`
    # in the risk arm for the sampler to be inside of.
    assert "If occurred" not in risk_arm, (
        "the risk arm regained an occurrence conditional; the severity call may "
        "now sit inside it")


def test_51_a_uniform_driver_is_sampled_on_min_and_max_alone() -> None:
    sampler = _procedure("SimEngineSampleValue")
    uniform = sampler[sampler.index("If prepared.DistKind = DIST_UNIFORM Then"):
                      sampler.index("ElseIf prepared.DistKind = DIST_TRIANGULAR")]
    assert "SimSampleUniform(state, prepared.MinValue, prepared.MaxValue," in \
        re.sub(r"\s*_\s*\n\s*", " ", uniform)
    assert "MostLikely" not in uniform
    triangular = sampler[sampler.index("ElseIf prepared.DistKind = DIST_TRIANGULAR"):
                         sampler.index("ElseIf prepared.DistKind = DIST_BETA_PERT")]
    assert "prepared.MostLikely" in triangular
    # An unknown kind never reaches a sampler.
    assert "an unknown distribution reached the sampler" in sampler


def test_52_a_risk_contribution_carries_no_quantity_and_no_probability() -> None:
    risk_arm = _loop_body().split("For index = costCount To driverCount - 1")[1]
    # THE LOOP asks for each measure once, passing the occurrence decision in.
    assert risk_arm.count("SimEngineContribution(prepared(index), severity, occurred,") == 2
    assert "prepared(index).Knom, SIM_MEASURE_NOMINAL" in risk_arm
    assert "prepared(index).Kpv, SIM_MEASURE_PV" in risk_arm
    # THE ROUTINE gives a risk two factors and neither is Quantity or
    # Probability, and a risk that did not occur contributes zero.
    contribution = _contribution()
    risk_shape = contribution[contribution.index("If prepared.IsRisk Then"):
                              contribution.index("Else")]
    assert "factors(0) = sample" in risk_shape
    assert "factors(1) = factor" in risk_shape
    assert "count = 2" in risk_shape
    assert "If Not occurred Then" in risk_shape
    for banned in ("Quantity", "Probability", "factors(2)"):
        assert banned not in risk_shape, banned
    # Probability is spent on the Bernoulli draw and folded into nothing.
    assert risk_arm.count("prepared(index).Probability") == 1
    assert "Probability" not in contribution
    assert "Probability" not in _procedure("SimEngineAdopt").split("If factor.IsRisk Then")[0]


def test_53_a_cost_contribution_is_unit_cost_times_quantity_times_a_factor() -> None:
    """The shape, read in the shared routine, and the WIRING, read in the loop."""
    cost_arm = _loop_body().split("For index = costCount To driverCount - 1")[0]
    # THE LOOP: samples a unit cost, then asks for each measure once.
    assert "SimEngineSampleValue(prepared(index), valueState(index), unitCost" in cost_arm
    assert cost_arm.count("SimEngineContribution(prepared(index), unitCost, True,") == 2
    assert "prepared(index).Knom, SIM_MEASURE_NOMINAL" in cost_arm
    assert "prepared(index).Kpv, SIM_MEASURE_PV" in cost_arm
    for banned in ("Probability", "SimSampleBernoulli"):
        assert banned not in cost_arm, banned
    # THE ROUTINE: three factors, and Quantity applied exactly once in the one
    # place any contribution is formed.
    contribution = _contribution()
    assert "factors(0) = sample" in contribution
    assert "factors(1) = prepared.Quantity" in contribution
    assert "factors(2) = factor" in contribution
    assert contribution.count("prepared.Quantity") == 1
    assert "count = 3" in contribution


def test_54_nominal_and_pv_diverge_when_their_factors_do() -> None:
    """Every other fixture collapses to Knom = Kpv = 1. This one does not.

    With two applied years and a discount rate, `Kpv` is strictly below `Knom`,
    so a shared accumulator, a nominal term built on `Kpv`, or a PV term copied
    from the nominal one all become arithmetically visible - not merely
    textually detectable.
    """
    model = _discounted_model(
        [_spread("CL-001", "cost", quantity=2.0)],
        [_spread("R-001", "risk", "Triangular", 100.0, 200.0, 400.0, probability=0.4)],
    )
    records = _factor_records(model)
    knom = {record["Knom"] for record in records}
    kpv = {record["Kpv"] for record in records}
    assert knom == {1.0}
    assert kpv != knom and all(0.0 < value < 1.0 for value in kpv), kpv
    ok, nominal, pv, detail = _run(records, 12345, 1000)
    assert ok, detail
    expected = _oracle(model, 12345, 1000)
    assert nominal == list(expected.total_nominal)
    assert pv == list(expected.total_pv)
    assert nominal != pv, "the two measures are indistinguishable on this fixture"
    # PV is strictly below nominal here, term by term, and never derived from it.
    for a, b in zip(nominal, pv):
        assert b <= a, (a, b)
    assert any(b < a for a, b in zip(nominal, pv))


def test_55_the_written_algorithm_reproduces_the_oracle_on_every_parity_case() -> None:
    """Run 4's P6-ORA disagreement is NOT a disagreement between the algorithms.

    Every Gate-B parity case is run twice here at the corpus's own seed and
    iteration count: once through the accepted Python oracle, and once through
    the statements `modSimRng`, `modSimSample` and `modSimEngine` actually write
    down. The comparison is bit-for-bit over all 1000 retained iterations of
    BOTH measures, and then over the result digest, which is exact by contract.

    That matters for the open P6-ORA investigation. Real Excel produced a
    different digest for all four cases, and this control says the difference
    cannot be attributed to what either implementation writes down: it has to
    come from how the VBA RUNTIME executes those statements - evaluation
    precision, the intrinsics, coercion - which is the one thing a Linux
    transcription cannot settle and which this file has always disclaimed.

    It is pinned here so that a later numerical change cannot quietly move one
    side while the runtime question is still open.
    """
    from pccm_builder.calc_cases import CASES as PLAN_CASES
    from pccm_builder.sim_cases import GATE_B_ITERATIONS, GATE_B_SUPPLIED_SEED

    by_id = {case["id"]: case for case in PLAN_CASES}
    # (plan case id, sampling mechanism, is the golden case) - the corpus's
    # own list, so a corpus that adds a case adds it here too.
    parity = [by_id[entry[0]] for entry in GATE_B_PARITY_PLAN_CASES]
    assert len(parity) == len(GATE_B_PARITY_PLAN_CASES) >= 4, parity

    for plan in parity:
        model = to_model(plan["model"])
        expected = _oracle(model, seed=GATE_B_SUPPLIED_SEED,
                           iterations=GATE_B_ITERATIONS)
        ok, nominal, pv, detail = _run(_factor_records(model),
                                       seed=GATE_B_SUPPLIED_SEED,
                                       iterations=GATE_B_ITERATIONS)
        assert ok, (plan["id"], detail)
        assert len(nominal) == len(pv) == GATE_B_ITERATIONS
        divergent = [index for index in range(GATE_B_ITERATIONS)
                     if nominal[index] != expected.total_nominal[index]
                     or pv[index] != expected.total_pv[index]]
        assert not divergent, (
            f"plan case {plan['id']}: the written algorithm and the oracle "
            f"first disagree at retained iteration {divergent[0]}"
        )
        assert result_digest(expected.sim_method_version, nominal, pv) == \
            expected.result_digest, plan["id"]
