#!/usr/bin/env python3
"""PCCM Phase 6 Step-7 conformance tests for `src/vba/modSimSample.bas`.

The stochastic transforms above the accepted generator: Uniform, Triangular,
prepared Beta-PERT through the locked Cheng BB/BC formulation, and the Bernoulli
occurrence primitive.

--------------------------------------------------------------------------------
WHAT THESE TESTS PROVE, AND WHAT THEY DO NOT
--------------------------------------------------------------------------------
SOURCE CONFORMANCE, on Linux, now. The module is read as text: its purity, its
public surface, its use of the projected constants, the shape of every locked
formula and branch operator, the state-consumption contract, the preparation
boundary, and the arithmetic those formulas describe.

VBA EXECUTION CONFORMANCE is NOT proved here and is deferred to Gate B on
Windows. No VBA runtime exists in this step. Where a test evaluates the
arithmetic it evaluates a TRANSCRIPTION of the expressions read out of the
module - evidence about the algorithm, not about the interpreter. No test in
this file may be read as "VBA reproduced Cheng vector 24".

COMPARISON POLICY IS NOT STRENGTHENED HERE. `build/phase6_cases.json` classifies
transformed sampler outputs as TOLERANCE_BOUNDED, and they are compared under
the bound the accepted Step-0 evidence policy owns - NOT with `==`, however
close the transcription happens to land. What stays EXACT is what the corpus
says stays exact: draw counts, consumption, proposal attempts, RNG states,
dispatch, orientation and Bernoulli decisions.

Runs standalone or under pytest.
"""

from __future__ import annotations

import contextlib
import json
import math
import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

from pccm_builder import load_contract, load_sim_contract, load_structure_contract  # noqa: E402
from pccm_builder.sim_emit import render_sim_contract_module  # noqa: E402
from pccm_builder.spec_loader import load_spec  # noqa: E402
from pccm_builder.vba_source import (  # noqa: E402
    VbaModule,
    contains_construct,
    load_modules,
    logical_statements,
)

from phase6_vba_transcribe import _Ref, _copy, build as _build_transcription  # noqa: E402

SRC_VBA = PCCM_ROOT / "src" / "vba"
SIM_SAMPLE_BAS = SRC_VBA / "modSimSample.bas"
SIM_RNG_BAS = SRC_VBA / "modSimRng.bas"
CALC_FACTORS_BAS = SRC_VBA / "modCalcFactors.bas"
SPEC = PCCM_ROOT / "spec"
EVIDENCE = PCCM_ROOT / "evidence" / "phase6_step0" / "vectors"
CASES_JSON = PCCM_ROOT / "build" / "phase6_cases.json"
STEP0_RECORD = PCCM_ROOT / "docs" / "phase6_step0.md"

STATE_FIELDS = ("S10", "S11", "S12", "S20", "S21", "S22")

_CACHE: dict[str, object] = {}


# ---------------------------------------------------------------------------
# Source access
# ---------------------------------------------------------------------------
def _module() -> VbaModule:
    return VbaModule(
        name="modSimSample", path=SIM_SAMPLE_BAS,
        raw=SIM_SAMPLE_BAS.read_text(encoding="utf-8"),
    )


def _code() -> str:
    """The module with comments and string literals removed."""
    return _module().code


def _procedure(name: str) -> str:
    """One procedure body, from its declaration to its End."""
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


def _generated_constants() -> dict[str, tuple[str, str]]:
    """`name -> (type, literal)` from the generated projection, rendered fresh."""
    if "consts" not in _CACHE:
        text = render_sim_contract_module(
            load_spec(SPEC / "workbook.yaml"),
            load_sim_contract(SPEC / "sim_contract.yaml"),
            load_contract(SPEC / "input_contract.yaml"),
        )
        out: dict[str, tuple[str, str]] = {}
        for line in text.splitlines():
            match = re.match(r"^Public Const (\w+) As (\w+) = (.*)$", line)
            if not match:
                continue
            name, kind, rest = match.groups()
            out[name] = (kind, rest.split("    '")[0].rstrip())
        _CACHE["consts"] = out
    return _CACHE["consts"]  # type: ignore[return-value]


def _const(name: str):
    kind, literal = _generated_constants()[name]
    if kind == "String":
        return literal[1:-1].replace('""', '"')
    return float(literal) if kind == "Double" else int(literal)


def _evidence(name: str) -> dict:
    return json.loads((EVIDENCE / f"{name}.json").read_text(encoding="utf-8"))


def _sampler_cases() -> dict[str, dict]:
    if "cases" not in _CACHE:
        corpus = json.loads(CASES_JSON.read_text(encoding="utf-8"))
        group = next(g for g in corpus["groups"] if g["group"] == "C_sampler")
        _CACHE["cases"] = {case["id"]: case for case in group["cases"]}
    return _CACHE["cases"]  # type: ignore[return-value]


def _tolerance(subject: str) -> tuple[float, float | None]:
    """The bound the ACCEPTED Step-0 evidence policy states, read from its table.

    `build/phase6_cases.json` deliberately refuses to carry a tolerance value -
    a number there would be a new authority - and points at the policy that owns
    it. This reads that policy rather than restating it.
    """
    rows = [line for line in STEP0_RECORD.read_text(encoding="utf-8").splitlines()
            if line.startswith(f"| {subject} ")]
    assert len(rows) == 1, (subject, len(rows))
    relative = re.search(r"rel\s*≤\s*([0-9.eE+-]+)", rows[0])
    absolute = re.search(r"abs\s*≤\s*([0-9.eE+-]+)", rows[0])
    assert relative, rows[0]
    return float(relative.group(1)), (float(absolute.group(1)) if absolute else None)


def _within(actual: float, expected: float, subject: str, scale: float | None = None) -> bool:
    relative, absolute = _tolerance(subject)
    if abs(actual - expected) <= relative * abs(expected):
        return True
    if absolute is not None and scale is not None:
        return abs(actual - expected) <= absolute * scale
    return False


# ---------------------------------------------------------------------------
# THE TRANSCRIPTION
#
# `tests/phase6_vba_transcribe.py` compiles the statements of modSimSample.bas
# into Python and runs them, together with the accepted modSimRng it consumes
# and the one accepted Phase-5 predicate it borrows. Every expression it
# evaluates is read out of the .bas at test time, so a change to a locked
# formula changes the answer. It is not a VBA interpreter; see the banner above.
# ---------------------------------------------------------------------------
def _transcribe() -> dict:
    if "vba" not in _CACHE:
        _CACHE["vba"] = _build_transcription(
            {
                "modSimRng": SIM_RNG_BAS,
                "modSimSample": SIM_SAMPLE_BAS,
                # BORROWED, NOT REIMPLEMENTED. IsUsableDouble is the accepted
                # Phase-5 finite predicate; only that one procedure is compiled,
                # so nothing else of modCalcFactors enters this step's evidence.
                "modCalcFactors": CALC_FACTORS_BAS,
            },
            {name: _const(name) for name in _generated_constants()},
            only={"modCalcFactors": {"IsUsableDouble"}},
            extra={"MAX_DOUBLE": sys.float_info.max},
        )
    return _CACHE["vba"]  # type: ignore[return-value]


def _state(*words) -> dict:
    return dict(zip(STATE_FIELDS, (float(w) for w in words)))


def _seeded(seed: int) -> dict:
    return _state(*([seed] * 6))


def _words(state: dict) -> list[int]:
    return [int(state[f]) for f in STATE_FIELDS]


def _blank_shape() -> dict:
    return _transcribe()["_new"]("SimSampleBetaShape")


def _prepare(a: float, m: float, b: float) -> tuple[bool, dict, str]:
    shape, detail = _blank_shape(), _Ref("")
    ok = _transcribe()["SimSamplePrepareBetaPert"](_Ref(a), _Ref(m), _Ref(b), shape, detail)
    return ok, shape, detail.v


def _uniform(state: dict, a: float, b: float):
    sample, consumed, detail = _Ref(0.0), _Ref(0), _Ref("")
    ok = _transcribe()["SimSampleUniform"](state, _Ref(a), _Ref(b), sample, consumed, detail)
    return ok, sample.v, consumed.v, detail.v


def _triangular(state: dict, a: float, m: float, b: float):
    sample, consumed, detail = _Ref(0.0), _Ref(0), _Ref("")
    ok = _transcribe()["SimSampleTriangular"](
        state, _Ref(a), _Ref(m), _Ref(b), sample, consumed, detail)
    return ok, sample.v, consumed.v, detail.v


def _beta(state: dict, shape: dict):
    sample, consumed, attempts, detail = _Ref(0.0), _Ref(0), _Ref(0), _Ref("")
    ok = _transcribe()["SimSamplePreparedBeta"](
        state, shape, sample, consumed, attempts, detail)
    return ok, sample.v, consumed.v, attempts.v, detail.v


def _bernoulli(state: dict, probability: float):
    occurred, uniform, consumed, detail = _Ref(False), _Ref(0.0), _Ref(0), _Ref("")
    ok = _transcribe()["SimSampleBernoulli"](
        state, _Ref(probability), occurred, uniform, consumed, detail)
    return ok, occurred.v, uniform.v, consumed.v, detail.v


@contextlib.contextmanager
def _injected(uniforms):
    """Script the VALUE of each draw while leaving the draw itself real.

    The module has no injected-uniform entry point, and must not: a production
    path obtains randomness only from modSimRng. So the accepted generator still
    runs, still advances the state and still costs one draw - only the value the
    sampler sees is replaced. Consumption and state behaviour stay real; the
    transform under test stops depending on which uniform the stream happened to
    produce.
    """
    vba = _transcribe()
    real = vba["SimRngNextUniform"]
    queue = list(uniforms)

    def scripted(state, u, detail):
        if not real(state, u, detail):
            return False
        assert queue, "the sampler drew more uniforms than the script supplies"
        u.v = queue.pop(0)
        return True

    vba["SimRngNextUniform"] = scripted
    try:
        yield queue
    finally:
        vba["SimRngNextUniform"] = real


# ===========================================================================
# A. The module exists, is declared, and exposes exactly what it should
# ===========================================================================
def test_01_the_module_exists_and_is_explicit() -> None:
    raw = _module().raw
    assert raw.startswith('Attribute VB_Name = "modSimSample"')
    assert re.search(r"^Option Explicit\s*$", raw, re.M)
    assert not re.search(r"^Option Base\b", raw, re.M)


def test_02_the_module_is_registered_as_handwritten_source() -> None:
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    modules = {m.name: m for m in structure.vba_modules}
    assert "modSimSample" in modules
    assert modules["modSimSample"].generated is False
    assert [m.name for m in structure.vba_modules][-4:] == [
        "modSimContract", "modSimRng", "modSimSample", "modSimEngine"
    ]
    for absent in ("modSimStats", "modSimFingerprint", "modSimReport"):
        assert absent not in modules, absent


def test_03_the_public_surface_is_exactly_the_five_samplers() -> None:
    assert sorted(_module().public_procedures) == [
        "SimSampleBernoulli",
        "SimSamplePrepareBetaPert",
        "SimSamplePreparedBeta",
        "SimSampleTriangular",
        "SimSampleUniform",
    ]
    # Everything numerical is private: the shared predicates, the conditioning
    # scale, both Cheng bodies and the orientation rule.
    private = set(_module().procedures) - set(_module().public_procedures)
    assert private == {
        "SimSampleChengBB", "SimSampleChengBC", "SimSampleOrientedBeta",
        "SimSampleOrderedTriple", "SimSampleScale", "SimSampleShapeInFamily",
        "SimSampleMinOf", "SimSampleMaxOf",
        "SimSampleValidatePreparedBetaShape", "SimSampleTermsAreUnset",
    }, sorted(private)


def test_04_no_public_procedure_accepts_an_object_or_a_variant() -> None:
    for name in _module().public_procedures:
        signature = _procedure(name).split("\n")[0]
        signature = " ".join(logical_statements(_procedure(name))[0][1].split())
        for banned in ("As Object", "As Variant", "As Range", "As Worksheet",
                       "As Workbook", "As ListObject", "ParamArray"):
            assert banned not in signature, (name, banned)
        # And every parameter is explicitly typed.
        for raw in re.search(r"\((.*)\)", signature).group(1).split(","):
            assert " As " in raw, (name, raw)


def test_05_one_public_type_carries_the_prepared_shape() -> None:
    raw = _module().raw
    types = re.findall(r"^Public Type (\w+)$", raw, re.M)
    assert types == ["SimSampleBetaShape"], types
    fields = _transcribe()["_types"]["SimSampleBetaShape"]
    names = [f for f, _ in fields]
    assert names == [
        "MinValue", "MostLikely", "MaxValue", "Alpha", "Beta", "Degenerate",
        "UseChengBB", "ChengA", "ChengB", "ChengAlpha", "ChengBeta",
        "ChengGamma", "ChengDelta", "ChengK1", "ChengK2",
        "FirstParameterIsOrientedA", "Prepared",
    ], names
    kinds = dict(fields)
    # The dispatch is a Boolean, not an unowned magic string.
    assert kinds["UseChengBB"] == "Boolean"
    assert kinds["FirstParameterIsOrientedA"] == "Boolean"
    assert kinds["Degenerate"] == "Boolean"
    # And it carries NO RNG state, no worksheet object and no driver row.
    assert "SimRngState" not in dict(fields).values()
    assert not any("Row" in f or "Range" in f for f in names)


# ===========================================================================
# B. Purity, and no randomness of its own
# ===========================================================================
def test_06_the_module_never_reaches_a_workbook_or_the_environment() -> None:
    code = _code()
    for token in ("Worksheet", "Workbook", "ThisWorkbook", "ActiveWorkbook",
                  "ActiveSheet", "Range", "Cells", "ListObject", "Application",
                  "Evaluate", "MsgBox", "InputBox", "CreateObject", "GetObject",
                  "Environ", "Shell", "Open ", "Close ", "Print #", "Kill ",
                  "Names(", "DoEvents", "Timer", "Now", "Date", "Sheets("):
        assert token not in code, token


def test_07_there_is_no_module_level_or_static_state() -> None:
    """Two callers cannot interfere, and a run cannot depend on the last one."""
    inside, module_level = False, []
    for _, text in logical_statements(_module().code_without_string_removal):
        if re.match(r"^(Public|Private)\s+(Function|Sub)\b", text):
            inside = True
        elif re.match(r"^End (Function|Sub)$", text):
            inside = False
        elif not inside and re.match(r"^(Public|Private|Dim|Global)\s+\w+\s+As\s", text):
            module_level.append(text)
    assert module_level == [], module_level
    assert not re.search(r"\bStatic\b", _code()), "a Static local holds state between calls"


def test_08_all_randomness_comes_through_the_accepted_generator() -> None:
    code = _code()
    # The two accepted modSimRng entry points, and nothing else of it.
    called = set(re.findall(r"\bSimRng\w+", code))
    assert called == {"SimRngValidateState", "SimRngNextUniform", "SimRngState"}, sorted(called)
    for banned in ("Rnd", "Randomize", "SimRngJumpNextStream", "SimRngStreamInitialState",
                   "SimRngStateFromFixedSeed", "SimRngAutoSeedFromNonce",
                   "SimRngBuildComponentStreams"):
        assert not re.search(rf"\b{banned}\b", code), banned
    # No recurrence, no jump, no seeding: not one generator constant is read.
    for constant in _generated_constants():
        if constant.startswith(("SIM_RNG_", "SIM_JUMP_", "SIM_SEED_", "SIM_NONCE_",
                                "SIM_AUTO_", "SIM_STREAM_", "SIM_COMPONENT_")):
            assert constant not in code, constant


def test_09_the_d6_11_algorithm_token_is_absent_and_unneeded() -> None:
    """Step 7 required no new scoped exception, and took none."""
    assert not contains_construct([_module()], "MRG32k3a")
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    scoped = [r for r in structure.forbidden_construct_rules if r.is_scoped]
    assert [(r.construct, tuple(r.allowed_in)) for r in scoped] == [
        ("MRG32k3a", ("modSimRng",))
    ], scoped
    for rule in structure.forbidden_construct_rules:
        assert rule.forbidden_in("modSimSample") is True, rule.construct
        assert not contains_construct([_module()], rule.construct), rule.construct
    endpoint = next(r for r in structure.forbidden_construct_rules
                    if r.construct == "RunSimulation")
    assert not endpoint.is_scoped and endpoint.forbidden_in("modSimReport") is True


def test_10_the_cheng_formulation_lives_here_and_nowhere_else() -> None:
    """The generator backbone knows no distribution, and no Phase-5 module does.

    `TriangularMean` in modCalcAnalytical is the accepted Phase-5 ANALYTICAL
    mean, not a sampler, and Phase-5 prose legitimately says "a Bernoulli draw
    in Monte Carlo" about work that had not happened yet. What must be unique to
    this module is the rejection-sampler formulation itself.
    """
    for module in load_modules([SRC_VBA]):
        if module.name == "modSimSample":
            continue
        # The scan is over EXECUTABLE code - comments and string literals
        # stripped, the same discipline D6-11 enforcement uses. modSimEngine
        # says in prose that no Cheng arithmetic lives in it, and a rule that
        # forbade a module from naming what it refuses to contain would forbid
        # the clearest thing it can say.
        # `SimSamplePrepareBetaPert` is the accepted PUBLIC constructor, so a
        # module naming it is calling the sampler rather than copying it.
        # What must be unique here is the formulation: every Cheng term and
        # every projected Cheng/PERT constant.
        for token in ("Cheng", "SIM_CHENG", "SIM_PERT"):
            assert token not in module.code, (module.name, token)
    assert "Cheng" in _module().code


# ===========================================================================
# C. No simulation-engine leakage
# ===========================================================================
def test_11_no_engine_concept_leaked_into_the_sampler() -> None:
    code = _code()
    for token in ("Quantity", "Knom", "Kpv", "Contribution", "Contribute",
                  "Iteration", "_SimData", "SimData", "ResultDigest", "Digest",
                  "Percentile", "Quantile", "Mean", "StdDev", "StandardDeviation",
                  "Contingency", "RunId", "run_id", "Fingerprint", "Results",
                  "RunSimulation", "Severity", "Occurrence"):
        assert token not in code, token


def test_12_the_module_declares_no_externally_callable_macro() -> None:
    assert not [p for p in _module().public_procedures if p.startswith("PCCM_")]
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    import yaml
    declared = yaml.safe_load((SPEC / "structure_contract.yaml").read_text(
        encoding="utf-8"))["vba"]
    surface = (set(structure.entry_points) | set(declared["harness_procedures"])
               | set(structure.api_procedures))
    assert not (set(_module().public_procedures) & surface)


def test_13_bernoulli_is_a_primitive_and_not_d6_18_orchestration() -> None:
    body = _procedure("SimSampleBernoulli")
    # One decision, one draw, and no knowledge that a Risk has a severity.
    assert body.count("SimRngNextUniform") == 1
    assert "SimSamplePreparedBeta" not in body
    assert "SimSampleTriangular" not in body
    assert "SimSampleUniform" not in body
    # And nothing else in the module calls it: there is no pairing here.
    calls = [name for name in _module().procedures
             if name != "SimSampleBernoulli" and "SimSampleBernoulli" in _procedure(name)]
    assert calls == [], calls


def test_14_no_public_sampler_calls_another_public_sampler() -> None:
    """There is no one-shot Beta path that could become the hot loop."""
    public = set(_module().public_procedures)
    for name in public:
        body = _procedure(name)
        for other in public - {name}:
            assert not re.search(rf"\b{other}\s*\(", body), (name, other)
    # In particular: no public entry point both prepares and samples.
    assert "SimSamplePrepareBetaPert" not in _procedure("SimSamplePreparedBeta")


# ===========================================================================
# D. SOURCE ARITHMETIC CONFORMANCE - the transcription against the accepted
#    Step-0 vectors and the accepted Step-5 corpus. NOT a VBA execution result.
# ===========================================================================
_TRANSFORM = "individual Uniform / Triangular / PERT-rescale transformed samples"
_CHENG = "deterministic Cheng vector outputs"


def test_15_the_accepted_tolerance_policy_is_read_not_restated() -> None:
    """The corpus refuses to carry a tolerance; Step 0 owns it. This reads it."""
    corpus = json.loads(CASES_JSON.read_text(encoding="utf-8"))
    assert "THE TOLERANCE VALUE IS NOT STATED HERE" in \
        corpus["comparison_policies"]["TOLERANCE_BOUNDED"]
    assert _tolerance(_TRANSFORM) == (1e-12, 1e-12)
    assert _tolerance(_CHENG) == (1e-11, None)


def test_16_uniform_reproduces_the_accepted_convex_transform() -> None:
    case = _sampler_cases()["sampler.uniform.injected"]
    a, b = case["inputs"]["a"], case["inputs"]["b"]
    assert case["expected_exact"]["transform"] == "x = (1 - u) * a + u * b"
    assert case["expected_exact"]["uniforms_per_sample"] == 1
    for row in case["expected"]["rows"]:
        state = _seeded(12345)
        with _injected([row["u"]]):
            ok, value, consumed, detail = _uniform(state, a, b)
        assert ok, detail
        assert consumed == 1
        assert _within(value, row["value"], _TRANSFORM, scale=max(abs(a), abs(b))), (
            row, value
        )


def test_17_uniform_stays_finite_across_the_widest_authorised_span() -> None:
    case = _sampler_cases()["sampler.uniform.extreme_span"]
    a, b = case["inputs"]["a"], case["inputs"]["b"]
    u = case["inputs"]["injected_uniforms"][0]
    with _injected([u]):
        ok, value, consumed, detail = _uniform(_seeded(12345), a, b)
    assert ok, detail
    assert consumed == case["expected_exact"]["uniforms_per_sample"] == 1
    assert math.isfinite(value) is case["expected_exact"]["finite"] is True
    assert _within(value, case["expected"]["value"], _TRANSFORM, scale=max(abs(a), abs(b)))
    # And the naive form is exactly what the convex one exists to avoid.
    assert not math.isfinite(b - a), "the span no longer exercises the overflow"


def test_18_uniform_accepts_negative_and_crossing_zero_supports() -> None:
    """No positivity rule, and the transform stays inside its own support."""
    for a, b in ((-500.0, -100.0), (-250.0, 250.0), (-1e300, 1e300)):
        previous = None
        for u in (1e-9, 0.25, 0.5, 0.75, 1 - 1e-9):
            with _injected([u]):
                ok, value, consumed, detail = _uniform(_seeded(12345), a, b)
            assert ok, (a, b, u, detail)
            assert consumed == 1
            assert math.isfinite(value)
            assert a <= value <= b, (a, b, u, value)
            if previous is not None:
                assert value >= previous, "the transform is not monotone in u"
            previous = value


def test_19_a_degenerate_uniform_draws_nothing_whatever_most_likely_holds() -> None:
    """Most Likely is not a parameter, so it cannot make the driver live."""
    for key in ("sampler.uniform.degenerate.absent",
                "sampler.uniform.degenerate.populated_and_ignored"):
        case = _sampler_cases()[key]
        expected = case["expected_exact"]
        state = _state(*case["inputs"]["start_state"])
        ok, value, consumed, detail = _uniform(
            state, case["inputs"]["a"], case["inputs"]["b"])
        assert ok, detail
        assert value == expected["value"]
        assert consumed == expected["uniforms_consumed"] == 0
        assert _words(state) == expected["state_after"]
    # The two cases differ ONLY in a Most Likely the signature cannot accept.
    absent = _sampler_cases()["sampler.uniform.degenerate.absent"]["inputs"]
    populated = _sampler_cases()[
        "sampler.uniform.degenerate.populated_and_ignored"]["inputs"]
    assert absent["most_likely"] is None and populated["most_likely"] == 500.0
    assert (absent["a"], absent["b"]) == (populated["a"], populated["b"])


def test_20_triangular_reproduces_both_branches_and_the_branch_point() -> None:
    case = _sampler_cases()["sampler.triangular.injected_branches"]
    a, m, b = case["inputs"]["a"], case["inputs"]["m"], case["inputs"]["b"]
    assert case["inputs"]["comparison_operator"] == "u <= c takes the lower branch"
    scale = max(abs(a), abs(m), abs(b))
    for row in case["expected"]["rows"]:
        with _injected([row["u"]]):
            ok, value, consumed, detail = _triangular(_seeded(12345), a, m, b)
        assert ok, (row, detail)
        assert consumed == case["expected_exact"]["uniforms_per_sample"] == 1
        assert _within(value, row["value"], _TRANSFORM, scale=scale), (row, value)
    # At u == c exactly the answer is the mode; `<` would move it.
    with _injected([case["inputs"]["branch_point_c"]]):
        ok, value, _, detail = _triangular(_seeded(12345), a, m, b)
    assert ok, detail
    assert value == case["expected_exact"]["value_at_branch_point"] == m


def test_21_triangular_handles_a_mode_at_either_endpoint() -> None:
    for key in ("sampler.triangular.m_equals_a", "sampler.triangular.m_equals_b"):
        case = _sampler_cases()[key]
        a, m, b = case["inputs"]["a"], case["inputs"]["m"], case["inputs"]["b"]
        u = case["inputs"]["injected_uniforms"][0]
        with _injected([u]):
            ok, value, consumed, detail = _triangular(_seeded(12345), a, m, b)
        assert ok, (key, detail)
        assert consumed == 1
        assert _within(value, case["expected"]["value"], _TRANSFORM,
                       scale=max(abs(a), abs(m), abs(b))), (key, value)


def test_22_triangular_survives_a_crossing_zero_extreme_support() -> None:
    a, m, b = -1e300, 0.0, 1e300
    for u in (1e-9, 0.25, 0.5, 0.75, 1 - 1e-9):
        with _injected([u]):
            ok, value, consumed, detail = _triangular(_seeded(12345), a, m, b)
        assert ok, (u, detail)
        assert consumed == 1 and math.isfinite(value)
        assert a <= value <= b, (u, value)
    # The raw unconditioned product is exactly what the conditioning avoids.
    assert not math.isfinite((b - a) * (m - a) if (b - a) != 0 else 0.0)


def test_23_a_degenerate_triangular_draws_nothing() -> None:
    case = _sampler_cases()["sampler.triangular.degenerate"]
    state = _state(*case["inputs"]["start_state"])
    before = _words(state)
    ok, value, consumed, detail = _triangular(
        state, case["inputs"]["a"], case["inputs"]["m"], case["inputs"]["b"])
    assert ok, detail
    assert value == case["expected_exact"]["value"]
    assert consumed == case["expected_exact"]["uniforms_consumed"] == 0
    assert _words(state) == before


def test_24_the_pert_parameterisation_walks_the_accepted_ladder() -> None:
    """r = 0, .25, .5, .75, 1 -> the five accepted shapes, and their dispatch."""
    ladder = [
        (0.00, 1.0, 5.0, False),
        (0.25, 2.0, 4.0, True),
        (0.50, 3.0, 3.0, True),
        (0.75, 4.0, 2.0, True),
        (1.00, 5.0, 1.0, False),
    ]
    lam = _const("SIM_PERT_LAMBDA")
    for r, alpha, beta, is_bb in ladder:
        ok, shape, detail = _prepare(0.0, r, 1.0)
        assert ok, (r, detail)
        assert shape["Alpha"] == alpha and shape["Beta"] == beta, (r, shape["Alpha"])
        assert shape["Alpha"] == 1.0 + lam * r
        assert shape["UseChengBB"] is is_bb, r
        assert shape["Degenerate"] is False and shape["Prepared"] is True
        assert shape["Alpha"] + shape["Beta"] == _const("SIM_PERT_ALPHA_PLUS_BETA")
        assert _const("SIM_PERT_SHAPE_LOWER") <= shape["Alpha"] <= _const("SIM_PERT_SHAPE_UPPER")
        assert _const("SIM_PERT_SHAPE_LOWER") <= shape["Beta"] <= _const("SIM_PERT_SHAPE_UPPER")


def test_25_equality_belongs_to_bc_and_the_endpoints_reach_it_by_the_rule() -> None:
    for key, expected_alpha, expected_beta in (
        ("sampler.beta.endpoint_low", 1.0, 5.0),
        ("sampler.beta.endpoint_high", 5.0, 1.0),
    ):
        case = _sampler_cases()[key]
        ok, shape, detail = _prepare(
            case["inputs"]["a"], case["inputs"]["m"], case["inputs"]["b"])
        assert ok, (key, detail)
        assert case["expected_exact"]["dispatch"] == "BC"
        assert shape["UseChengBB"] is False, key
        assert shape["Alpha"] == case["expected_exact"]["alpha"] == expected_alpha
        assert shape["Beta"] == case["expected_exact"]["beta"] == expected_beta
        assert shape["Alpha"] + shape["Beta"] == case["expected_exact"]["alpha_plus_beta"]
    # min(alpha, beta) == 1 exactly, so the boundary is what decided it.
    ok, shape, _ = _prepare(0.0, 0.0, 100.0)
    assert min(shape["Alpha"], shape["Beta"]) == _const("SIM_PERT_SHAPE_LOWER")


def test_26_degeneracy_is_settled_before_the_shape_ratio_is_formed() -> None:
    case = _sampler_cases()["sampler.beta.degenerate"]
    assert case["expected_exact"]["shape_ratio_formed"] is False
    a = case["inputs"]["a"]
    ok, shape, detail = _prepare(a, case["inputs"]["m"], case["inputs"]["b"])
    assert ok, detail
    assert shape["Degenerate"] is True and shape["Prepared"] is True
    # Nothing downstream of the ratio was computed, so 0/0 never arose.
    assert shape["Alpha"] == 0.0 and shape["Beta"] == 0.0
    assert shape["ChengA"] == 0.0 and shape["ChengBeta"] == 0.0
    assert shape["ChengK1"] == 0.0 and shape["ChengK2"] == 0.0
    state = _state(*case["inputs"]["start_state"])
    before = _words(state)
    ok, value, consumed, attempts, detail = _beta(state, shape)
    assert ok, detail
    assert value == case["expected_exact"]["value"] == a
    assert consumed == case["expected_exact"]["uniforms_consumed"] == 0
    assert attempts == 0
    assert _words(state) == before


def test_27_preparation_draws_nothing_and_takes_no_rng_state() -> None:
    signature = logical_statements(_procedure("SimSamplePrepareBetaPert"))[0][1]
    assert "SimRngState" not in signature
    body = _procedure("SimSamplePrepareBetaPert")
    assert "SimRngNextUniform" not in body
    assert "SimRngValidateState" not in body


def _cheng_shape(alpha: float) -> tuple[dict, float]:
    """The (a, m, b) on [0, 1] whose PERT parameterisation is this alpha."""
    r = (alpha - 1.0) / _const("SIM_PERT_LAMBDA")
    ok, shape, detail = _prepare(0.0, r, 1.0)
    assert ok, detail
    assert shape["Alpha"] == alpha, (alpha, shape["Alpha"])
    return shape, r


def test_28_every_accepted_cheng_vector_family_is_reproduced() -> None:
    """All five retained families: value, attempts, consumption, state, orientation.

    The VALUES are compared under the accepted TOLERANCE_BOUNDED policy. The
    attempt counts, the uniforms consumed and every RNG state are EXACT, because
    that is what the corpus says stays exact even in a tolerance-bounded case.
    """
    vectors = _evidence("cheng_vectors")
    assert vectors["both_dispatches_covered"] == {"BB": 3, "BC": 2}
    assert len(vectors["cases"]) == 5
    per_attempt = _const("SIM_CHENG_UNIFORMS_PER_ATTEMPT")
    for case in vectors["cases"]:
        shape, _ = _cheng_shape(case["alpha"])
        assert shape["Beta"] == case["beta"], case["label"]
        assert shape["UseChengBB"] is (case["dispatch"] == "BB"), case["label"]
        state = _state(*case["initial_state"])
        assert case["exercises_at_least_one_retry"] is True
        assert case["exercises_immediate_acceptance"] is True
        cumulative = 0
        for spec in case["samples"]:
            ok, value, consumed, attempts, detail = _beta(state, shape)
            assert ok, (case["label"], spec["index"], detail)
            # EXACT: consumption, attempts and the state they left.
            assert attempts == spec["proposal_attempts_for_this_sample"], (
                case["label"], spec["index"])
            assert consumed == spec["uniforms_for_this_sample"] == per_attempt * attempts
            cumulative += consumed
            assert cumulative == spec["cumulative_uniforms"]
            assert _words(state) == spec["rng_state_after_sample"], (
                case["label"], spec["index"])
            # TOLERANCE_BOUNDED: the transformed value.
            assert _within(value, float(spec["accepted_sample"]), _CHENG), (
                case["label"], spec["index"], value, spec["accepted_sample"])
        assert _words(state) == case["final_state"], case["label"]
        assert max(s["proposal_attempts_for_this_sample"] for s in case["samples"]) \
            == case["max_attempts_in_case"]


def test_29_the_corpus_totals_agree_with_the_reproduced_consumption() -> None:
    for key in ("bb_interior", "bb_symmetric", "bb_near_boundary",
                "bc_alpha_1", "bc_beta_1"):
        case = _sampler_cases()[f"sampler.beta.cheng.{key}"]
        assert case["comparison"] == "TOLERANCE_BOUNDED"
        expected = case["expected_exact"]
        shape, _ = _cheng_shape(expected["alpha"])
        assert shape["Beta"] == expected["beta"]
        assert shape["UseChengBB"] is (expected["dispatch"] == "BB")
        state = _state(*expected["initial_state"])
        total_attempts, total_uniforms = 0, 0
        for row in case["expected"]["samples"]:
            ok, value, consumed, attempts, detail = _beta(state, shape)
            assert ok, (key, detail)
            total_attempts += attempts
            total_uniforms += consumed
            assert _within(value, row["value"], _CHENG), (key, row["index"], value)
        assert total_attempts == expected["total_proposal_attempts"], key
        assert total_uniforms == expected["total_uniforms"], key
        assert expected["uniforms_per_attempt"] == _const("SIM_CHENG_UNIFORMS_PER_ATTEMPT")
        assert _words(state) == expected["final_state"], key


def test_30_the_cheng_orientation_is_opposite_between_the_two_dispatches() -> None:
    for alpha, beta in ((2.0, 4.0), (4.0, 2.0), (3.0, 3.0)):
        shape, _ = _cheng_shape(alpha)
        assert shape["UseChengBB"] is True
        assert shape["ChengA"] == min(alpha, beta), (alpha, beta)
        assert shape["ChengB"] == max(alpha, beta), (alpha, beta)
        assert shape["FirstParameterIsOrientedA"] is (alpha == shape["ChengA"])
    for alpha, beta in ((1.0, 5.0), (5.0, 1.0)):
        shape, _ = _cheng_shape(alpha)
        assert shape["UseChengBB"] is False
        assert shape["ChengA"] == max(alpha, beta), (alpha, beta)
        assert shape["ChengB"] == min(alpha, beta), (alpha, beta)
        assert shape["FirstParameterIsOrientedA"] is (alpha == shape["ChengA"])


def test_31_the_prepared_cheng_terms_are_the_accepted_precomputation() -> None:
    """The per-driver terms, evaluated from the projected literals."""
    lit4 = _const("SIM_CHENG_BB_LITERAL_4")
    lit5 = _const("SIM_CHENG_BB_LITERAL_5")
    shape, _ = _cheng_shape(2.0)  # BB, a = 2, b = 4
    ca, cb = shape["ChengA"], shape["ChengB"]
    assert (ca, cb) == (2.0, 4.0)
    assert shape["ChengAlpha"] == ca + cb
    assert shape["ChengBeta"] == math.sqrt(
        (shape["ChengAlpha"] - lit4) / (lit4 * ca * cb - shape["ChengAlpha"]))
    assert shape["ChengGamma"] == ca + lit5 / shape["ChengBeta"]
    assert shape["ChengDelta"] == 0.0 and shape["ChengK1"] == 0.0

    shape, _ = _cheng_shape(1.0)  # BC, a = 5, b = 1
    ca, cb = shape["ChengA"], shape["ChengB"]
    assert (ca, cb) == (5.0, 1.0)
    assert shape["ChengAlpha"] == ca + cb
    assert shape["ChengBeta"] == 1.0 / cb
    assert shape["ChengDelta"] == 1.0 + ca - cb
    assert shape["ChengK1"] == shape["ChengDelta"] * (
        _const("SIM_CHENG_BC_LITERAL_1") + _const("SIM_CHENG_BC_LITERAL_2") * cb
    ) / (ca * shape["ChengBeta"] - _const("SIM_CHENG_BC_LITERAL_3"))
    assert shape["ChengK2"] == _const("SIM_CHENG_BC_LITERAL_5") + (
        _const("SIM_CHENG_BC_LITERAL_6")
        + _const("SIM_CHENG_BC_LITERAL_5") / shape["ChengDelta"]) * cb
    assert shape["ChengGamma"] == 0.0


def test_32_bernoulli_reproduces_the_accepted_decision_table() -> None:
    case = _sampler_cases()["sampler.bernoulli.decision_table"]
    assert case["comparison"] == "EXACT"
    assert case["inputs"]["comparison_operator"] == "strictly_less_than"
    assert case["inputs"]["rule"] == "occurred = u < probability"
    for row in case["expected_exact"]["rows"]:
        with _injected([row["u"]]):
            ok, occurred, uniform, consumed, detail = _bernoulli(
                _seeded(12345), row["probability"])
        assert ok, (row, detail)
        assert occurred is row["occurred"], row
        assert uniform == row["u"]
        assert consumed == 1
    # And the boundary the table exists to pin: u == p is False.
    boundary = [r for r in case["expected_exact"]["rows"] if r["u"] == r["probability"]]
    assert boundary, "the u == p boundary row disappeared from the corpus"
    for row in boundary:
        assert row["occurred"] is False, row


def test_33_bernoulli_consumes_exactly_one_uniform_at_both_extremes() -> None:
    case = _sampler_cases()["sampler.bernoulli.stream_consumption"]
    expected = case["expected_exact"]
    state = _state(*case["inputs"]["start_state"])
    ok, occurred, uniform, consumed, detail = _bernoulli(
        state, case["inputs"]["probability"])
    assert ok, detail
    assert uniform == expected["uniform"]
    assert occurred is expected["occurred"]
    assert consumed == expected["uniforms_consumed"] == 1
    assert _words(state) == expected["state_after"]
    # p = 0 and p = 1 still cost one draw each, and still decide by the rule.
    for probability, decision in ((0.0, False), (1.0, True)):
        state = _seeded(12345)
        before = _words(state)
        ok, occurred, uniform, consumed, detail = _bernoulli(state, probability)
        assert ok, detail
        assert consumed == 1, probability
        assert occurred is decision, probability
        assert _words(state) != before, "p = %r skipped the draw" % probability
        assert 0.0 < uniform < 1.0


def test_34_a_probability_outside_the_unit_interval_is_refused() -> None:
    case = _sampler_cases()["sampler.bernoulli.probability_refused"]
    assert case["expected_refusal"]["kind"] == "probability_domain"
    for probability in case["inputs"]["probabilities"]:
        state = _seeded(12345)
        before = _words(state)
        occurred, uniform, consumed, detail = _Ref(True), _Ref(-1.0), _Ref(-1), _Ref("")
        ok = _transcribe()["SimSampleBernoulli"](
            state, _Ref(probability), occurred, uniform, consumed, detail)
        assert ok is False, probability
        assert "outside [0, 1]" in detail.v and "not clamped" in detail.v
        assert _words(state) == before, "a refused probability still drew"
        assert (occurred.v, uniform.v, consumed.v) == (True, -1.0, -1)


# ===========================================================================
# E. FAILURE AND STATE ATOMICITY
# ===========================================================================
_BROKEN = (0.0, 0.0, 0.0, 1.0, 1.0, 1.0)   # an absorbing first component


def test_35_an_invalid_state_is_refused_even_on_a_degenerate_uniform() -> None:
    state = _state(*_BROKEN)
    before = _words(state)
    sample, consumed, detail = _Ref(-1.0), _Ref(-1), _Ref("")
    ok = _transcribe()["SimSampleUniform"](
        state, _Ref(100.0), _Ref(100.0), sample, consumed, detail)
    assert ok is False, "a degenerate distribution excused an invalid state"
    assert "all zero" in detail.v
    assert _words(state) == before
    assert (sample.v, consumed.v) == (-1.0, -1)


def test_36_an_invalid_state_is_refused_even_on_a_degenerate_triangular() -> None:
    state = _state(*_BROKEN)
    sample, consumed, detail = _Ref(-1.0), _Ref(-1), _Ref("")
    ok = _transcribe()["SimSampleTriangular"](
        state, _Ref(7.0), _Ref(7.0), _Ref(7.0), sample, consumed, detail)
    assert ok is False and "all zero" in detail.v
    assert (sample.v, consumed.v) == (-1.0, -1)


def test_37_an_invalid_state_is_refused_even_on_a_degenerate_beta() -> None:
    ok, shape, detail = _prepare(-2.0, -2.0, -2.0)
    assert ok and shape["Degenerate"] is True
    state = _state(*_BROKEN)
    sample, consumed, attempts, detail = _Ref(-1.0), _Ref(-1), _Ref(-1), _Ref("")
    ok = _transcribe()["SimSamplePreparedBeta"](
        state, shape, sample, consumed, attempts, detail)
    assert ok is False and "all zero" in detail.v
    assert (sample.v, consumed.v, attempts.v) == (-1.0, -1, -1)


def test_38_an_invalid_ordering_is_refused_before_any_draw() -> None:
    vba = _transcribe()
    for a, m, b in ((10.0, 5.0, 1.0), (0.0, 50.0, 10.0), (0.0, -1.0, 10.0)):
        state = _seeded(12345)
        before = _words(state)
        sample, consumed, detail = _Ref(-1.0), _Ref(-1), _Ref("")
        assert vba["SimSampleTriangular"](
            state, _Ref(a), _Ref(m), _Ref(b), sample, consumed, detail) is False
        assert "refused, not repaired" in detail.v
        assert _words(state) == before and (sample.v, consumed.v) == (-1.0, -1)
        shape, detail = _blank_shape(), _Ref("")
        assert vba["SimSamplePrepareBetaPert"](
            _Ref(a), _Ref(m), _Ref(b), shape, detail) is False
        assert "refused, not repaired" in detail.v
        assert shape["Prepared"] is False
    # Uniform refuses a > b and never swaps the endpoints.
    state = _seeded(12345)
    before = _words(state)
    sample, consumed, detail = _Ref(-1.0), _Ref(-1), _Ref("")
    assert vba["SimSampleUniform"](
        state, _Ref(100.0), _Ref(0.0), sample, consumed, detail) is False
    assert "never swapped" in detail.v or "refused" in detail.v
    assert _words(state) == before and (sample.v, consumed.v) == (-1.0, -1)


def test_39_an_unprepared_shape_cannot_be_sampled() -> None:
    state = _seeded(12345)
    before = _words(state)
    sample, consumed, attempts, detail = _Ref(-1.0), _Ref(-1), _Ref(-1), _Ref("")
    assert _transcribe()["SimSamplePreparedBeta"](
        state, _blank_shape(), sample, consumed, attempts, detail) is False
    assert "never prepared" in detail.v
    assert _words(state) == before


def test_40_a_rejected_cheng_proposal_advances_the_working_state() -> None:
    """No rewind. A retry continues from the state the rejection left."""
    vectors = _evidence("cheng_vectors")
    per_attempt = _const("SIM_CHENG_UNIFORMS_PER_ATTEMPT")
    seen_retry = False
    for case in vectors["cases"]:
        shape, _ = _cheng_shape(case["alpha"])
        state = _state(*case["initial_state"])
        for spec in case["samples"]:
            before = _copy(state)
            ok, value, consumed, attempts, detail = _beta(state, shape)
            assert ok, detail
            if attempts > 1:
                seen_retry = True
                # Every attempt cost its two uniforms, including the rejected
                # ones: a rewind would have committed fewer.
                assert consumed == per_attempt * attempts
                stepped = _copy(before)
                for _ in range(consumed):
                    u, d = _Ref(0.0), _Ref("")
                    assert _transcribe()["SimRngNextUniform"](stepped, u, d), d.v
                assert _words(stepped) == _words(state), (
                    "the committed state is not the one all consumed draws reached"
                )
    assert seen_retry, "no accepted vector exercised a retry"


def test_41_a_failing_cheng_sample_leaves_the_caller_untouched() -> None:
    """The working copy absorbs every rejection; the caller sees none of it."""
    vba = _transcribe()
    # A case whose FIRST sample genuinely retries, so the failure lands after
    # several rejections have already advanced the working copy.
    case = next(c for c in _evidence("cheng_vectors")["cases"]
                if c["samples"][0]["proposal_attempts_for_this_sample"] > 1)
    needed = case["samples"][0]["proposal_attempts_for_this_sample"]
    budget = _const("SIM_CHENG_UNIFORMS_PER_ATTEMPT") * (needed - 1)
    shape, _ = _cheng_shape(case["alpha"])
    state = _state(*case["initial_state"])
    before = _words(state)
    real = vba["SimRngNextUniform"]
    calls = [0]

    def failing(working, u, detail):
        calls[0] += 1
        if calls[0] <= budget:     # every rejection but the last one is real
            return real(working, u, detail)
        detail.v = "uniform: injected failure"
        return False

    vba["SimRngNextUniform"] = failing
    try:
        sample, consumed, attempts, detail = _Ref(-1.0), _Ref(-1), _Ref(-1), _Ref("")
        ok = vba["SimSamplePreparedBeta"](
            state, shape, sample, consumed, attempts, detail)
    finally:
        vba["SimRngNextUniform"] = real
    assert ok is False and detail.v == "uniform: injected failure"
    assert calls[0] == budget + 1, calls[0]
    assert _words(state) == before, "a failed Beta sample committed a partial walk"
    assert (sample.v, consumed.v, attempts.v) == (-1.0, -1, -1)


def test_42_an_accepted_cheng_result_commits_exactly_the_final_state() -> None:
    for case in _evidence("cheng_vectors")["cases"]:
        shape, _ = _cheng_shape(case["alpha"])
        state = _state(*case["initial_state"])
        for _ in case["samples"]:
            assert _beta(state, shape)[0]
        assert _words(state) == case["final_state"], case["label"]


def test_43_a_failing_uniform_or_triangular_draw_leaves_everything_untouched() -> None:
    vba = _transcribe()
    real = vba["SimRngNextUniform"]

    def failing(working, u, detail):
        detail.v = "uniform: injected failure"
        return False

    for name, args in (
        ("SimSampleUniform", (_Ref(0.0), _Ref(100.0))),
        ("SimSampleTriangular", (_Ref(0.0), _Ref(30.0), _Ref(100.0))),
    ):
        state = _seeded(12345)
        before = _words(state)
        sample, consumed, detail = _Ref(-1.0), _Ref(-1), _Ref("")
        vba["SimRngNextUniform"] = failing
        try:
            ok = vba[name](state, *args, sample, consumed, detail)
        finally:
            vba["SimRngNextUniform"] = real
        assert ok is False, name
        assert _words(state) == before, name
        assert (sample.v, consumed.v) == (-1.0, -1), name


def test_44_a_non_representable_rescale_is_refused_not_returned() -> None:
    """A silent inf would travel into a published total."""
    vba = _transcribe()
    state = _seeded(12345)
    before = _words(state)
    sample, consumed, detail = _Ref(-1.0), _Ref(-1), _Ref("")
    huge = sys.float_info.max
    with _injected([0.5]):
        # A support whose convex midpoint is finite still succeeds...
        assert vba["SimSampleUniform"](
            _seeded(12345), _Ref(-huge), _Ref(huge), sample, consumed, detail), detail.v
    assert sample.v == 0.0
    # ...while a genuinely unrepresentable endpoint is refused at the boundary.
    sample, consumed, detail = _Ref(-1.0), _Ref(-1), _Ref("")
    assert vba["SimSampleUniform"](
        state, _Ref(float("inf")), _Ref(huge), sample, consumed, detail) is False
    assert "not a finite Double" in detail.v
    assert _words(state) == before and (sample.v, consumed.v) == (-1.0, -1)


# ===========================================================================
# F. SHAPE CONFORMANCE - what the source is allowed to say
# ===========================================================================
def test_45_no_owned_value_is_restated_as_a_literal() -> None:
    """Every business number has an owner; a second copy is a second authority.

    The scan is over NUMERIC LITERALS, found with a lookbehind that refuses a
    digit preceded by a word character, so the `_3` of `SIM_CHENG_BC_LITERAL_3`
    is an identifier and not a number.
    """
    code = _code()
    literals = {m.group(0).rstrip("#")
                for m in re.finditer(r"(?<![\w.])\d+(?:\.\d+)?#?", code)}
    # ZERO AND ONE, and nothing else. Every shape parameter, every squeeze
    # literal, lambda and the two family bounds are read from the projection.
    assert literals == {"0", "1"}, sorted(literals)
    spelled = {float(v) for v in literals}
    for name, (kind, literal) in _generated_constants().items():
        if kind == "String":
            continue
        assert float(literal.rstrip("#")) not in spelled or float(literal) in (0.0, 1.0), (
            name, literal
        )
    # In particular the Cheng squeeze literals are never spelled here.
    for value in ("1.3862944", "2.609438", "0.0138889", "0.0416667",
                  "0.777778", "0.25", "0.5", "131072", "4", "5", "6"):
        assert value not in literals, value


def test_46_every_locked_cheng_literal_is_read_from_the_projection() -> None:
    bb, bc = _procedure("SimSampleChengBB"), _procedure("SimSampleChengBC")
    prepare = _procedure("SimSamplePrepareBetaPert")
    assert "SIM_CHENG_BB_LITERAL_1" in bb and "SIM_CHENG_BB_LITERAL_2" in bb
    assert "SIM_CHENG_BB_LITERAL_3" in bb
    assert "SIM_CHENG_BB_LITERAL_4" in prepare and "SIM_CHENG_BB_LITERAL_5" in prepare
    for index in (1, 2, 3, 5, 6):
        assert f"SIM_CHENG_BC_LITERAL_{index}" in prepare + bc, index
    assert "SIM_CHENG_BC_LITERAL_4" in bc
    assert "SIM_PERT_LAMBDA" in prepare
    # And they are never re-derived from a mathematical equivalent.
    code = _code()
    for derivation in ("Log(4", "Log(5", "Log(2", "1 / 72", "1/72", "3 / 72",
                       "7 / 9", "7/9", "Exp(1"):
        assert derivation not in code, derivation


def test_47_the_uniform_transform_is_the_convex_form() -> None:
    body = _procedure("SimSampleUniform")
    assert "candidate = (1# - u) * minValue + u * maxValue" in body
    assert "maxValue - minValue" not in body, "the naive difference reappeared"
    assert "(maxValue - minValue)" not in body
    # Most Likely is not even a parameter, so it cannot be read.
    signature = logical_statements(body)[0][1]
    assert "mostLikely" not in signature and "MostLikely" not in signature
    assert "ostLikely" not in body
    # Degeneracy is a = b, and only that.
    assert "If minValue = maxValue Then" in body
    assert body.count("uniformsConsumed = 0") == 1
    assert body.count("SimRngNextUniform") == 1


def test_48_the_triangular_branch_and_conditioning_are_exact() -> None:
    body = _procedure("SimSampleTriangular")
    assert "s = SimSampleScale(minValue, mostLikely, maxValue)" in body
    for statement in ("an = minValue / s", "mn = mostLikely / s", "bn = maxValue / s",
                      "span = bn - an", "c = (mn - an) / span"):
        assert statement in body, statement
    assert "If u <= c Then" in body, "the branch boundary moved"
    # AND THE RIGHT FORM IN THE RIGHT ARM. Both statements present in either
    # order would satisfy a containment check while sampling the mirror image.
    lower, upper = body.split("If u <= c Then", 1)[1].split("Else", 1)
    assert "conditioned = an + Sqr(u * span * (mn - an))" in lower, lower
    assert "conditioned = bn - Sqr" not in lower, lower
    assert "conditioned = bn - Sqr((1# - u) * span * (bn - mn))" in upper, upper
    assert "conditioned = an + Sqr" not in upper, upper
    assert "candidate = conditioned * s" in body
    # The unconditioned products never appear.
    for unsafe in ("(maxValue - minValue) * (mostLikely - minValue)",
                   "(mostLikely - minValue) * (maxValue - minValue)"):
        assert unsafe not in body, unsafe
    assert "If minValue = mostLikely And mostLikely = maxValue Then" in body
    assert body.count("SimRngNextUniform") == 1


def test_49_the_conditioning_scale_is_the_accepted_one() -> None:
    body = _procedure("SimSampleScale")
    assert "s = Abs(first)" in body
    assert "If Abs(second) > s Then s = Abs(second)" in body
    assert "If Abs(third) > s Then s = Abs(third)" in body
    assert "If s <= 0# Then s = 1#" in body
    vba = _transcribe()
    for values, expected in (((3.0, -7.0, 2.0), 7.0), ((0.0, 0.0, 0.0), 1.0),
                             ((-1e300, 1.0, 5.0), 1e300)):
        assert vba["SimSampleScale"](*[_Ref(v) for v in values]) == expected


def test_50_the_dispatch_boundary_gives_equality_to_bc() -> None:
    body = _procedure("SimSamplePrepareBetaPert")
    assert "If SimSampleMinOf(alpha0, beta0) > SIM_PERT_SHAPE_LOWER Then" in body
    assert ">= SIM_PERT_SHAPE_LOWER" not in body, "equality was given to BB"
    # Degeneracy is settled BEFORE the ratio, and the ratio is formed once.
    assert body.index("If minValue = mostLikely And mostLikely = maxValue Then") < \
        body.index("r = (mn - an) / span")
    assert body.count("r = (mn - an) / span") == 1
    assert "alpha0 = 1# + SIM_PERT_LAMBDA * r" in body
    assert "beta0 = 1# + SIM_PERT_LAMBDA * (1# - r)" in body


def test_51_the_cheng_orientation_statements_are_opposite() -> None:
    body = _procedure("SimSamplePrepareBetaPert")
    bb_at = body.index("candidate.UseChengBB = True")
    bc_at = body.index("candidate.UseChengBB = False")
    bb, bc = body[bb_at:bc_at], body[bc_at:]
    assert "ca = SimSampleMinOf(alpha0, beta0)" in bb
    assert "cb = SimSampleMaxOf(alpha0, beta0)" in bb
    assert "ca = SimSampleMaxOf(alpha0, beta0)" in bc
    assert "cb = SimSampleMinOf(alpha0, beta0)" in bc
    # The BB precomputation, and every term of it, is in the BB arm only.
    assert "Sqr(" in bb and "Sqr(" not in bc
    assert "ChengGamma" in bb and "ChengGamma" not in bc
    assert "ChengK1" in bc and "ChengK1" not in bb
    assert "ChengK2" in bc and "ChengK2" not in bb


def test_52_the_prepared_terms_are_never_recomputed_in_the_proposal_loop() -> None:
    """The hot loop reads the shape. It does not re-derive it."""
    for name in ("SimSampleChengBB", "SimSampleChengBC"):
        body = _procedure(name)
        for recomputation in ("Sqr(", "SimSampleMinOf", "SimSampleMaxOf",
                              "ChengK1 =", "ChengK2 =", "ChengBeta =",
                              "ChengGamma =", "ChengDelta =", "ChengAlpha =",
                              "SIM_PERT_LAMBDA", "SIM_PERT_SHAPE"):
            assert recomputation not in body, (name, recomputation)
        # Every Cheng term the loop uses is READ from the prepared shape.
        for term in re.findall(r"prepared\.(\w+)", body):
            assert term.startswith("Cheng") or term == "FirstParameterIsOrientedA", term
    # And the shape is only ever written during preparation.
    for name in _module().procedures:
        if name == "SimSamplePrepareBetaPert":
            continue
        assert not re.search(r"\bprepared\.\w+\s*=[^=]", _procedure(name)), name


def test_53_a_cheng_attempt_consumes_exactly_two_uniforms_and_never_rewinds() -> None:
    for name in ("SimSampleChengBB", "SimSampleChengBC"):
        body = _procedure(name)
        assert body.count("SimRngNextUniform") == _const("SIM_CHENG_UNIFORMS_PER_ATTEMPT")
        assert "If Not SimRngNextUniform(working, u1, detail) Then Exit Function" in body
        assert "If Not SimRngNextUniform(working, u2, detail) Then Exit Function" in body
        # One loop, entered once, with no state saved to restore.
        assert body.count("\n    Do\n") == 1
        assert "working = " not in body, "the loop rewinds the working state"
        assert "attempts = attempts + 1" in body
    consumption = _procedure("SimSamplePreparedBeta")
    assert "uniformsConsumed = SIM_CHENG_UNIFORMS_PER_ATTEMPT * attempts" in consumption
    assert "proposalAttempts = attempts" in consumption


def test_54_the_bb_acceptance_tests_are_in_the_accepted_order() -> None:
    body = _procedure("SimSampleChengBB")
    for statement in (
        "vlog = Log(u1 / (1# - u1))",
        "v = prepared.ChengBeta * vlog",
        "w = prepared.ChengA * Exp(v)",
        "z = u1 * u1 * u2",
        "rr = prepared.ChengGamma * v - SIM_CHENG_BB_LITERAL_1",
        "ss = prepared.ChengA + rr - w",
        "If ss + SIM_CHENG_BB_LITERAL_2 >= SIM_CHENG_BB_LITERAL_3 * z Then Exit Do",
        "t = Log(z)",
        "If ss >= t Then Exit Do",
    ):
        assert statement in body, statement
    assert ("If rr + prepared.ChengAlpha * Log(prepared.ChengAlpha / "
            "(prepared.ChengB + w)) >= t Then Exit Do") in body
    # Order: squeeze, then log test, then the full test.
    assert body.index("SIM_CHENG_BB_LITERAL_2 >=") < body.index("t = Log(z)")
    assert body.index("If ss >= t") < body.index("If rr + prepared.ChengAlpha")
    # The logit is the accepted single-expression form.
    assert "Log(u1) - Log(1# - u1)" not in body


def test_55_the_bc_operators_are_the_accepted_ones() -> None:
    body = _procedure("SimSampleChengBC")
    assert "If u1 < SIM_CHENG_BC_LITERAL_6 Then" in body
    assert ("If SIM_CHENG_BC_LITERAL_5 * u2 + z - y0 >= prepared.ChengK1 "
            "Then rejected = True") in body
    assert "y0 = u1 * u2" in body and "z = u1 * y0" in body
    assert "z = u1 * u1 * u2" in body
    assert "If z <= SIM_CHENG_BC_LITERAL_5 Then" in body
    assert "If z >= prepared.ChengK2 Then rejected = True" in body
    assert ("If prepared.ChengAlpha * (Log(prepared.ChengAlpha / (prepared.ChengB + w)) + v)"
            in re.sub(r"\s*_\s*\n\s*", " ", body))
    assert "- SIM_CHENG_BC_LITERAL_4 >= Log(z) Then Exit Do" in re.sub(
        r"\s*_\s*\n\s*", " ", body)
    # The z <= 0.25 arm accepts IMMEDIATELY, before the final test.
    arm = body[body.index("If z <= SIM_CHENG_BC_LITERAL_5 Then"):]
    assert arm[:arm.index("End If")].count("Exit Do") == 1
    assert "Log(u1) - Log(1# - u1)" not in body


def test_56_the_beta_rescale_is_convex_and_clips_nothing() -> None:
    body = _procedure("SimSamplePreparedBeta")
    assert "candidate = (1# - y) * prepared.MinValue + y * prepared.MaxValue" in body
    assert "MaxValue - prepared.MinValue" not in body
    assert "If Not (y > 0# And y < 1#) Then" in body
    for clip in ("If y > 1#", "If y < 0#", "If candidate > prepared.MaxValue",
                 "If candidate < prepared.MinValue"):
        assert clip not in body, clip
    assert "not representable" in body


def test_57_the_orientation_return_rule_is_shared_and_unmirrored() -> None:
    body = _procedure("SimSampleOrientedBeta")
    assert "If prepared.FirstParameterIsOrientedA Then" in body
    assert "denominator = prepared.ChengB + w" in body
    # AND THE RIGHT RETURN IN THE RIGHT ARM. Mirroring these is silent: the
    # sampler still returns a valid Beta variate, of the mirrored distribution.
    oriented, mirrored = body.split(
        "If prepared.FirstParameterIsOrientedA Then", 1)[1].split("Else", 1)
    assert "candidate = w / denominator" in oriented, oriented
    assert "prepared.ChengB / denominator" not in oriented, oriented
    assert "candidate = prepared.ChengB / denominator" in mirrored, mirrored
    assert "candidate = w / denominator" not in mirrored, mirrored
    # One rule, called from both dispatches, so they cannot drift.
    for name in ("SimSampleChengBB", "SimSampleChengBC"):
        assert "SimSampleOrientedBeta(prepared, w, y, detail)" in _procedure(name)


def test_58_bernoulli_is_strict_and_never_clamps() -> None:
    body = _procedure("SimSampleBernoulli")
    assert "occurred = (u < probability)" in body
    assert "<=" not in body.split("occurred =")[1].split("\n")[0]
    assert "If probability < 0# Or probability > 1# Then" in body
    assert "not clamped" in body
    for special in ("If probability = 0#", "If probability = 1#"):
        assert special not in body, special
    assert body.count("SimRngNextUniform") == 1
    assert "uniformsConsumed = 1" in body


def test_59_every_output_is_committed_after_its_last_check() -> None:
    for procedure, guard, commit in (
        ("SimSampleUniform", "IsUsableDouble(candidate)", "state = working"),
        ("SimSampleTriangular", "IsUsableDouble(candidate)", "state = working"),
        ("SimSamplePreparedBeta", "IsUsableDouble(candidate)", "state = working"),
        ("SimSampleBernoulli", "SimRngNextUniform(working, u, detail)", "state = working"),
        ("SimSamplePrepareBetaPert", "candidate.Prepared = True", "prepared = candidate"),
    ):
        body = _procedure(procedure)
        assert guard in body and commit in body, procedure
        assert body.rindex(guard) < body.rindex(commit), procedure
    # Every state-consuming sampler draws against a LOCAL copy.
    for procedure in ("SimSampleUniform", "SimSampleTriangular",
                      "SimSamplePreparedBeta", "SimSampleBernoulli"):
        body = _procedure(procedure)
        assert "working = state" in body
        assert body.index("working = state") < body.index("state = working")
        assert "Dim working As SimRngState" in body


def test_60_the_finite_predicate_is_borrowed_and_not_reimplemented() -> None:
    code = _code()
    assert "IsUsableDouble" in code
    assert "MAX_DOUBLE" not in code, "a second maximum-Double authority appeared"
    for reimplementation in ("1.79769", "308", "IsNumeric", "IsError", "VarType",
                             "WorksheetFunction", "CDec", "CCur"):
        assert reimplementation not in code, reimplementation


def test_61_the_transcription_read_the_whole_module() -> None:
    vba = _transcribe()
    for name in _module().procedures:
        assert callable(vba[name]), name
    assert "SimSampleBetaShape" in vba["_types"]
    # And the accepted generator it consumes is compiled from its own source.
    for name in ("SimRngValidateState", "SimRngNextUniform"):
        assert callable(vba[name]), name


# ===========================================================================
# G. THE PREPARED SHAPE IS CHECKED, NOT TRUSTED
#
# SimSampleBetaShape is a PUBLIC VBA UDT, so its fields are caller-writable and
# nothing makes the value immutable the way the frozen Python reference is.
# `Prepared = True` is therefore a CLAIM about provenance, and these tests prove
# it is not on its own an authority.
# ===========================================================================
def _forged(**fields) -> dict:
    """A prepared record built by hand, exactly as a caller could build one."""
    shape = _blank_shape()
    shape.update(fields)
    return shape


def _refuse(shape: dict, fragment: str, state: dict | None = None) -> str:
    """Sampling must refuse, and leave the caller and every output untouched."""
    state = state if state is not None else _seeded(12345)
    before = _words(state)
    sample, consumed, attempts, detail = _Ref(-1.0), _Ref(-1), _Ref(-1), _Ref("")
    ok = _transcribe()["SimSamplePreparedBeta"](
        state, shape, sample, consumed, attempts, detail)
    assert ok is False, (fragment, detail.v)
    assert fragment in detail.v, (fragment, detail.v)
    assert _words(state) == before, fragment
    assert (sample.v, consumed.v, attempts.v) == (-1.0, -1, -1), fragment
    return detail.v


def _honest(a: float, m: float, b: float) -> dict:
    ok, shape, detail = _prepare(a, m, b)
    assert ok, detail
    return shape


def test_62_a_provenance_flag_alone_is_not_validation_authority() -> None:
    """The reported defect: a hand-built degenerate record over a live support.

    `Prepared = True`, `Degenerate = True`, and a support that is not only
    non-degenerate but MISORDERED, used to return a zero-draw success carrying
    MinValue. It is now refused on the ordering, before the RNG state is looked
    at and before any draw.
    """
    forgery = _forged(Prepared=True, Degenerate=True,
                      MinValue=123.456, MostLikely=-999.0, MaxValue=-1000.0)
    _refuse(forgery, "refused, not repaired")
    # And it is refused even when the RNG state is ALSO invalid, so the refusal
    # cannot be coming from the state check.
    _refuse(forgery, "refused, not repaired", state=_state(*_BROKEN))
    # The validator runs first in the source, too.
    body = _procedure("SimSamplePreparedBeta")
    assert body.index("SimSampleValidatePreparedBetaShape") < \
        body.index("SimRngValidateState"), "the state is validated before the shape"
    assert body.index("SimSampleValidatePreparedBetaShape") < \
        body.index("If prepared.Degenerate Then")


def test_63_a_degeneracy_flag_must_agree_with_its_support_both_ways() -> None:
    # True over a live support: the zero-draw forgery.
    _refuse(_forged(Prepared=True, Degenerate=True,
                    MinValue=0.0, MostLikely=25.0, MaxValue=100.0),
            "degeneracy flag disagrees")
    # False over a dead support: a = m = b would reach the parameterisation.
    _refuse(_forged(Prepared=True, Degenerate=False,
                    MinValue=7.0, MostLikely=7.0, MaxValue=7.0),
            "degeneracy flag disagrees")


def test_64_a_non_finite_prepared_support_is_refused() -> None:
    for field, value in (("MinValue", float("nan")), ("MostLikely", float("inf")),
                         ("MaxValue", float("-inf")), ("MinValue", float("-inf"))):
        shape = _honest(0.0, 25.0, 100.0)
        shape[field] = value
        _refuse(shape, "not a finite Double")


def test_65_a_degenerate_record_may_carry_no_active_shape() -> None:
    for field, value in (("Alpha", 2.0), ("Beta", 4.0), ("UseChengBB", True),
                         ("FirstParameterIsOrientedA", True), ("ChengA", 2.0),
                         ("ChengB", 4.0), ("ChengAlpha", 6.0), ("ChengBeta", 0.5),
                         ("ChengGamma", 4.0), ("ChengDelta", 5.0),
                         ("ChengK1", 0.1), ("ChengK2", 0.8)):
        shape = _honest(-2.0, -2.0, -2.0)
        assert shape["Degenerate"] is True
        shape[field] = value
        _refuse(shape, "Beta-PERT:")
    # And the honest degenerate record still succeeds, at zero draws.
    shape = _honest(-2.0, -2.0, -2.0)
    state = _seeded(12345)
    before = _words(state)
    ok, value, consumed, attempts, detail = _beta(state, shape)
    assert ok, detail
    assert (value, consumed, attempts) == (-2.0, 0, 0)
    assert _words(state) == before


def _self_consistent(alpha: float, beta: float) -> dict:
    """A record consistent in EVERY structural way except the family bound.

    Built so the family check is the only thing standing between it and a
    sample: dispatch, orientation, the oriented sum and every active Cheng term
    are exactly what preparation would have written for this alpha and beta.
    """
    lower, upper = min(alpha, beta), max(alpha, beta)
    is_bb = lower > _const("SIM_PERT_SHAPE_LOWER")
    cheng_a, cheng_b = (lower, upper) if is_bb else (upper, lower)
    shape = _forged(Prepared=True, Degenerate=False,
                    MinValue=0.0, MostLikely=0.5, MaxValue=1.0,
                    Alpha=alpha, Beta=beta, UseChengBB=is_bb,
                    ChengA=cheng_a, ChengB=cheng_b, ChengAlpha=cheng_a + cheng_b,
                    FirstParameterIsOrientedA=(alpha == cheng_a))
    if is_bb:
        lit4, lit5 = _const("SIM_CHENG_BB_LITERAL_4"), _const("SIM_CHENG_BB_LITERAL_5")
        shape["ChengBeta"] = math.sqrt(
            (shape["ChengAlpha"] - lit4) / (lit4 * cheng_a * cheng_b - shape["ChengAlpha"]))
        shape["ChengGamma"] = cheng_a + lit5 / shape["ChengBeta"]
    else:
        shape["ChengBeta"] = 1.0 / cheng_b
        shape["ChengDelta"] = 1.0 + cheng_a - cheng_b
        shape["ChengK1"] = shape["ChengDelta"] * (
            _const("SIM_CHENG_BC_LITERAL_1") + _const("SIM_CHENG_BC_LITERAL_2") * cheng_b
        ) / (cheng_a * shape["ChengBeta"] - _const("SIM_CHENG_BC_LITERAL_3"))
        shape["ChengK2"] = _const("SIM_CHENG_BC_LITERAL_5") + (
            _const("SIM_CHENG_BC_LITERAL_6")
            + _const("SIM_CHENG_BC_LITERAL_5") / shape["ChengDelta"]) * cheng_b
    return shape


def test_66_a_shape_parameter_outside_the_accepted_family_is_refused() -> None:
    # Structurally consistent everywhere else, so ONLY the family bound refuses.
    for alpha, beta in ((5.5, 5.5), (0.5, 0.5)):
        _refuse(_self_consistent(alpha, beta), "left the accepted family")
    # And a shape parameter that is corrupted on its own is refused too.
    for field, value in (("Alpha", 0.5), ("Alpha", 5.5), ("Beta", 0.0),
                         ("Beta", 6.0), ("Alpha", float("nan"))):
        shape = _honest(0.0, 0.25, 1.0)
        shape[field] = value
        _refuse(shape, "Beta-PERT:")
    # The bound is the projected one, and the endpoints of it are legal.
    shape = _honest(0.0, 0.0, 100.0)
    assert min(shape["Alpha"], shape["Beta"]) == _const("SIM_PERT_SHAPE_LOWER")
    assert max(shape["Alpha"], shape["Beta"]) == _const("SIM_PERT_SHAPE_UPPER")
    assert _beta(_seeded(12345), shape)[0]
    # The control is not vacuous: the same construction INSIDE the family is
    # accepted by the validator, so the family bound is what did the refusing.
    inside = _self_consistent(2.0, 4.0)
    detail = _Ref("")
    assert _transcribe()["SimSampleValidatePreparedBetaShape"](inside, detail), detail.v


def test_67_a_dispatch_that_disagrees_with_the_shape_is_refused() -> None:
    bb = _honest(0.0, 0.25, 1.0)       # alpha 2, beta 4 -> BB
    assert bb["UseChengBB"] is True
    bb["UseChengBB"] = False
    _refuse(bb, "dispatch disagrees")
    bc = _honest(0.0, 0.0, 100.0)      # alpha 1, beta 5 -> BC, equality is BC
    assert bc["UseChengBB"] is False
    bc["UseChengBB"] = True
    _refuse(bc, "dispatch disagrees")


def test_68_a_swapped_cheng_orientation_is_refused() -> None:
    bb = _honest(0.0, 0.25, 1.0)
    assert (bb["ChengA"], bb["ChengB"]) == (2.0, 4.0)
    bb["ChengA"], bb["ChengB"] = bb["ChengB"], bb["ChengA"]
    _refuse(bb, "BB orientation is not min, max")
    bc = _honest(0.0, 0.0, 100.0)
    assert (bc["ChengA"], bc["ChengB"]) == (5.0, 1.0)
    bc["ChengA"], bc["ChengB"] = bc["ChengB"], bc["ChengA"]
    _refuse(bc, "BC orientation is not max, min")
    # And the oriented sum has to be the sum of the oriented pair.
    shape = _honest(0.0, 0.25, 1.0)
    shape["ChengAlpha"] = shape["ChengAlpha"] + 1.0
    _refuse(shape, "not the oriented sum")


def test_69_a_recorded_orientation_that_disagrees_is_refused() -> None:
    for a, m, b in ((0.0, 0.25, 1.0), (0.0, 0.75, 1.0), (0.0, 0.0, 100.0)):
        shape = _honest(a, m, b)
        shape["FirstParameterIsOrientedA"] = not shape["FirstParameterIsOrientedA"]
        _refuse(shape, "recorded orientation disagrees")


def test_70_a_malformed_active_cheng_term_is_refused() -> None:
    for field, value, fragment in (
        ("ChengBeta", float("nan"), "Cheng beta term is not a finite Double"),
        ("ChengBeta", 0.0, "Cheng beta term is not positive"),
        ("ChengBeta", -1.0, "Cheng beta term is not positive"),
        ("ChengGamma", float("inf"), "Cheng gamma term is not a finite Double"),
    ):
        shape = _honest(0.0, 0.25, 1.0)          # BB
        shape[field] = value
        _refuse(shape, fragment)
    for field, value, fragment in (
        ("ChengBeta", float("nan"), "Cheng beta term is not a finite Double"),
        ("ChengDelta", 0.0, "Cheng delta term is not positive"),
        ("ChengDelta", float("nan"), "Cheng delta term is not a finite Double"),
        ("ChengK1", float("inf"), "Cheng k1 term is not a finite Double"),
        ("ChengK2", float("nan"), "Cheng k2 term is not a finite Double"),
    ):
        shape = _honest(0.0, 0.0, 100.0)         # BC
        shape[field] = value
        _refuse(shape, fragment)


def test_71_a_record_may_not_carry_both_cheng_families_at_once() -> None:
    bb = _honest(0.0, 0.25, 1.0)
    for field in ("ChengDelta", "ChengK1", "ChengK2"):
        forged = dict(bb)
        forged[field] = 1.0
        _refuse(forged, "never writes is not at its default")
    bc = _honest(0.0, 0.0, 100.0)
    forged = dict(bc)
    forged["ChengGamma"] = 1.0
    _refuse(forged, "never writes is not at its default")


def test_72_a_valid_degenerate_shape_still_validates_the_rng_state() -> None:
    shape = _honest(-2.0, -2.0, -2.0)
    _refuse(shape, "all zero", state=_state(*_BROKEN))
    # ...and with a valid state it is an exact zero-draw success.
    state = _seeded(12345)
    before = _words(state)
    ok, value, consumed, attempts, detail = _beta(state, shape)
    assert ok, detail
    assert (value, consumed, attempts) == (-2.0, 0, 0)
    assert _words(state) == before


def test_73_the_hardening_moved_no_accepted_sample() -> None:
    """Every accepted vector, re-derived through the validated path.

    The validator is refusal hardening: a shape built by
    SimSamplePrepareBetaPert passes it, and nothing about the sample it produces
    changes - value, attempts, consumption, state, dispatch or orientation.
    """
    per_attempt = _const("SIM_CHENG_UNIFORMS_PER_ATTEMPT")
    for case in _evidence("cheng_vectors")["cases"]:
        shape, _ = _cheng_shape(case["alpha"])
        detail = _Ref("")
        assert _transcribe()["SimSampleValidatePreparedBetaShape"](shape, detail), detail.v
        assert shape["UseChengBB"] is (case["dispatch"] == "BB")
        state = _state(*case["initial_state"])
        for spec in case["samples"]:
            ok, value, consumed, attempts, refusal = _beta(state, shape)
            assert ok, (case["label"], refusal)
            assert attempts == spec["proposal_attempts_for_this_sample"]
            assert consumed == per_attempt * attempts
            assert _words(state) == spec["rng_state_after_sample"]
            assert _within(value, float(spec["accepted_sample"]), _CHENG)
        assert _words(state) == case["final_state"], case["label"]


def test_74_the_validator_recomputes_no_preparation_and_draws_nothing() -> None:
    body = _procedure("SimSampleValidatePreparedBetaShape")
    for recomputation in ("Sqr(", "Log(", "Exp(", "SIM_PERT_LAMBDA",
                          "SIM_CHENG_BB_LITERAL", "SIM_CHENG_BC_LITERAL",
                          "SimSampleScale", "SimSamplePrepareBetaPert"):
        assert recomputation not in body, recomputation
    # No division at all: k1, k2 and the BB square root stay per-driver work.
    assert "/" not in body
    # It draws nothing and it writes nothing.
    assert "SimRngNextUniform" not in body
    assert not re.search(r"\bprepared\.\w+\s*=[^=]", body), "the validator writes the shape"
    # And the sampler still never calls the constructor.
    assert "SimSamplePrepareBetaPert" not in _procedure("SimSamplePreparedBeta")
    # The Step-7 precomputation guarantee is untouched: the loops are unchanged.
    for name in ("SimSampleChengBB", "SimSampleChengBC"):
        assert "SimSampleValidatePreparedBetaShape" not in _procedure(name), name
