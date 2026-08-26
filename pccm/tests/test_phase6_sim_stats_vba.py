#!/usr/bin/env python3
"""PCCM Phase 6 Step-9 conformance tests for `src/vba/modSimStats.bas`.

Sort, moments, Hyndman-Fan type-7 quantiles, the accepted ladder and
selected-Px contingency. All scale-safe.

--------------------------------------------------------------------------------
WHAT THESE TESTS PROVE, AND WHAT THEY DO NOT
--------------------------------------------------------------------------------
SOURCE CONFORMANCE, on Linux, now: purity, the public surface, the exact source
formulas, the sorting semantics, the power-of-two scale selection, mean and
sample deviation, type 7, the ladder, contingency and the extreme-domain
behaviour - through the accepted Phase-6 source transcriber, against the
accepted Step-4 Python oracle and the accepted Step-5 corpus.

VBA EXECUTION CONFORMANCE is NOT proved and is deferred to Gate B on Windows.
No VBA runtime exists in this step. Nothing here may be read as "VBA executed a
near-MAX sample deviation".

FOUR PRIMITIVES ARE BORROWED, NOT TRANSCRIBED. `SafeSignedSum`, `SafeDivide`,
`SafeMultiply` and `SafeSubtract` have accepted Phase-5 VBA bodies using scoped
error handlers and an exact-arithmetic UDT the transcriber does not model. Their
accepted PYTHON counterparts in `calc_numeric` are bound instead, and their real
VBA SIGNATURES are read out of `modCalcFactors.bas` so the ByRef/ByVal call
convention stays the module's own.

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
    load_contract,
    load_sim_contract,
    load_structure_contract,
)
from pccm_builder import sim_stats as oracle  # noqa: E402
from pccm_builder.calc_numeric import (  # noqa: E402
    CalculationRefusal,
    safe_divide,
    safe_multiply,
    safe_signed_sum,
    safe_subtract,
)
from pccm_builder.contract_loader import ContractError  # noqa: E402
from pccm_builder.sim_emit import render_sim_contract_module  # noqa: E402
from pccm_builder.spec_loader import load_spec  # noqa: E402
from pccm_builder.vba_source import (  # noqa: E402
    VbaModule,
    contains_construct,
    load_modules,
    logical_statements,
)

from phase6_vba_transcribe import _Ref, _val, build as _build_transcription  # noqa: E402

SRC_VBA = PCCM_ROOT / "src" / "vba"
SIM_STATS_BAS = SRC_VBA / "modSimStats.bas"
CALC_FACTORS_BAS = SRC_VBA / "modCalcFactors.bas"
SPEC = PCCM_ROOT / "spec"
CASES_JSON = PCCM_ROOT / "build" / "phase6_cases.json"

_CACHE: dict[str, object] = {}


# ---------------------------------------------------------------------------
# Source access
# ---------------------------------------------------------------------------
def _module() -> VbaModule:
    return VbaModule(name="modSimStats", path=SIM_STATS_BAS,
                     raw=SIM_STATS_BAS.read_text(encoding="utf-8"))


def _code() -> str:
    return _module().code


def _procedure(name: str) -> str:
    code = _module().code_without_string_removal
    pattern = re.compile(
        rf"^\s*(?:Public|Private)\s+(?:Function|Sub)\s+{re.escape(name)}\b", re.M)
    match = pattern.search(code)
    assert match, f"{name} is not declared"
    tail = code[match.start():]
    end = re.search(r"^\s*End\s+(?:Function|Sub)\s*$", tail, re.M)
    assert end, f"{name} has no End"
    return tail[: end.end()]


def _constants() -> dict:
    if "consts" not in _CACHE:
        out: dict = {}
        rendered = render_sim_contract_module(
            load_spec(SPEC / "workbook.yaml"),
            load_sim_contract(SPEC / "sim_contract.yaml"),
            load_contract(SPEC / "input_contract.yaml"))
        for text in (rendered, CALC_FACTORS_BAS.read_text(encoding="utf-8")):
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


def _scalar_shim(function):
    def bound(*args):
        *operands, result = args
        try:
            result.v = function(*[float(_val(v)) for v in operands], "statistics")
        except (CalculationRefusal, ContractError):
            return False
        return True
    return bound


def _sequence_shim(function):
    def bound(sequence, count, result):
        try:
            result.v = function([float(v) for v in sequence[: int(_val(count))]], "statistics")
        except (CalculationRefusal, ContractError):
            return False
        return True
    return bound


# The ONE procedure of this module whose body the transcriber cannot execute:
# it reads a bound of an unproven carrier under a scoped error handler, and the
# engine models no `On Error`. Its SIGNATURE is read from the real source, its
# body is source-tested by `test_48`, and its shim reproduces the allocated arm
# (`UBound - LBound + 1`) so short and long carriers stay covered behaviourally.
# The genuinely never-sized VBA array - the arm that RAISES - is Gate-B work.
BORROWED_FROM_MODULE = {"SimStatsLadderExtent"}


def _declared_procedures() -> set[str]:
    return set(_module().procedures)


def _ladder_extent_shim(quantile_labels, quantile_values, label_extent, value_extent):
    """The allocated arm of `SimStatsLadderExtent`, and only that arm."""
    label_extent.v = len(quantile_labels)
    value_extent.v = len(quantile_values)
    return True


def _transcribe() -> dict:
    if "vba" not in _CACHE:
        _CACHE["vba"] = _build_transcription(
            {"modSimStats": SIM_STATS_BAS, "modCalcFactors": CALC_FACTORS_BAS},
            _constants(),
            only={"modCalcFactors": {"IsUsableDouble"},
                  "modSimStats": _declared_procedures() - BORROWED_FROM_MODULE},
            signature_only={"modCalcFactors": {
                "SafeSignedSum", "SafeDivide", "SafeMultiply", "SafeSubtract"},
                "modSimStats": set(BORROWED_FROM_MODULE)},
            extra={
                "MAX_DOUBLE": sys.float_info.max,
                "SafeSignedSum": _sequence_shim(safe_signed_sum),
                "SafeDivide": _scalar_shim(safe_divide),
                "SafeMultiply": _scalar_shim(safe_multiply),
                "SafeSubtract": _scalar_shim(safe_subtract),
                "SimStatsLadderExtent": _ladder_extent_shim,
            })
    return _CACHE["vba"]  # type: ignore[return-value]


# ---------------------------------------------------------------------------
# Calling conveniences
# ---------------------------------------------------------------------------
def _mean(values):
    result, detail = _Ref(0.0), _Ref("")
    ok = _transcribe()["SimStatsMean"](list(values), _Ref(len(values)), result, detail)
    return ok, result.v, detail.v


def _sd(values):
    result, detail = _Ref(0.0), _Ref("")
    ok = _transcribe()["SimStatsSampleStandardDeviation"](
        list(values), _Ref(len(values)), result, detail)
    return ok, result.v, detail.v


def _quantile(values, p):
    result, detail = _Ref(0.0), _Ref("")
    ok = _transcribe()["SimStatsQuantileType7"](
        list(values), _Ref(len(values)), _Ref(p), result, detail)
    return ok, result.v, detail.v


def _describe(values):
    summary = _transcribe()["_new"]("SimStatsMeasure")
    labels, ladder, detail = [], [], _Ref("")
    ok = _transcribe()["SimStatsDescribe"](
        list(values), _Ref(len(values)), summary, labels, ladder, detail)
    return ok, summary, labels, ladder, detail.v


def _selected(labels, ladder, label, count=None, result=None):
    """`count` defaults to the label carrier's own length; pass it explicitly to
    exercise a carrier whose physical length disagrees with the claimed one."""
    result = _Ref(0.0) if result is None else result
    detail = _Ref("")
    ok = _transcribe()["SimStatsSelectedQuantile"](
        labels, ladder, _Ref(len(labels) if count is None else count),
        _Ref(label), result, detail)
    return ok, result.v, detail.v


def _contingency(selected_total, base_estimate):
    result, detail = _Ref(0.0), _Ref("")
    ok = _transcribe()["SimStatsContingency"](
        _Ref(selected_total), _Ref(base_estimate), result, detail)
    return ok, result.v, detail.v


def _ladder_points():
    """The accepted (label, p) ladder, read from the projection."""
    return [(_const(f"SIM_QUANTILE_{i}"), float(_const(f"SIM_QUANTILE_{i}")[1:]) / 100.0)
            for i in range(1, _const("SIM_QUANTILE_COUNT") + 1)]


def _cases() -> dict[str, dict]:
    if "cases" not in _CACHE:
        corpus = json.loads(CASES_JSON.read_text(encoding="utf-8"))
        _CACHE["cases"] = {case["id"]: case
                           for group in corpus["groups"] for case in group["cases"]}
    return _CACHE["cases"]  # type: ignore[return-value]


# ===========================================================================
# A. Declaration, surface and purity
# ===========================================================================
def test_01_the_module_exists_and_is_explicit() -> None:
    raw = _module().raw
    assert raw.startswith('Attribute VB_Name = "modSimStats"')
    assert re.search(r"^Option Explicit\s*$", raw, re.M)
    assert not re.search(r"^Option Base\b", raw, re.M)


def test_02_the_module_is_registered_and_nothing_beyond_it() -> None:
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    modules = {m.name: m for m in structure.vba_modules}
    assert "modSimStats" in modules
    assert modules["modSimStats"].generated is False
    assert [m.name for m in structure.vba_modules][-5:] == [
        "modSimContract", "modSimRng", "modSimSample", "modSimEngine", "modSimStats"]
    for absent in ("modSimFingerprint", "modSimReport"):
        assert absent not in modules, absent
    assert not (set(_module().public_procedures)
                & (set(structure.entry_points) | set(structure.api_procedures)))
    # D6-11 is untouched, and Percentile is still global.
    scoped = [(r.construct, tuple(r.allowed_in))
              for r in structure.forbidden_construct_rules if r.is_scoped]
    assert scoped == [("MRG32k3a", ("modSimRng",))], scoped
    for construct in ("RunSimulation", "Percentile"):
        rule = next(r for r in structure.forbidden_construct_rules
                    if r.construct == construct)
        assert not rule.is_scoped, construct
        assert rule.forbidden_in("modSimStats") is True, construct


def test_03_the_public_surface_is_the_six_numerical_operations() -> None:
    assert sorted(_module().public_procedures) == [
        "SimStatsContingency",
        "SimStatsDescribe",
        "SimStatsMean",
        "SimStatsQuantileType7",
        "SimStatsSampleStandardDeviation",
        "SimStatsSelectedQuantile",
    ]
    private = set(_module().procedures) - set(_module().public_procedures)
    assert private == {
        "SimStatsConstantValue", "SimStatsLadderExtent", "SimStatsLadderLabel",
        "SimStatsProbabilityOf", "SimStatsQuantileSorted", "SimStatsSortAscending",
        "SimStatsSortedCopy", "SimStatsUnitScale", "SimStatsUsableProbability",
        "SimStatsUsableSequence", "SimStatsValidateLadder",
    }, sorted(private)
    raw = _module().raw
    assert re.findall(r"^(Public|Private) Type (\w+)$", raw, re.M) == [
        ("Public", "SimStatsMeasure")]


def test_04_the_summary_type_is_derived_reporting_output_only() -> None:
    fields = _transcribe()["_types"]["SimStatsMeasure"]
    assert [f for f, _ in fields] == [
        "Count", "Mean", "SampleStandardDeviation", "Minimum", "Maximum",
        "QuantileCount", "Described"], fields
    # It carries no retained sample and no simulation authority: every field is
    # a scalar, and none of them is a sequence, a state or an identity.
    assert all(kind in ("Long", "Double", "Boolean") for _, kind in fields), fields
    for name, _ in fields:
        for banned in ("Values", "Totals", "State", "Seed", "Digest", "Run"):
            assert banned not in name, (name, banned)


def test_05_the_module_never_reaches_a_workbook_or_the_environment() -> None:
    code = _code()
    for token in ("Range", "Cells", "Worksheet", "Worksheets", "Workbook",
                  "Workbooks", "ListObject", "Application", "WorksheetFunction",
                  "ThisWorkbook", "ActiveWorkbook", "ActiveSheet", "Names(",
                  "Evaluate", "MsgBox", "InputBox", "CreateObject", "GetObject",
                  "Environ", "Shell", "Open ", "Print #", "DoEvents", "Timer",
                  "Now", "Date", "Sheets("):
        assert token not in code, token
    for name in _module().public_procedures:
        signature = logical_statements(_procedure(name))[0][1]
        for banned in ("As Object", "As Variant", "As Range", "ParamArray"):
            assert banned not in signature, (name, banned)


def test_06_there_is_no_module_level_or_static_state() -> None:
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


def test_07_the_forbidden_word_never_appears_in_executable_code() -> None:
    """`Percentile` is globally forbidden and Step 9 took no exception."""
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    for rule in structure.forbidden_construct_rules:
        assert rule.forbidden_in("modSimStats") is True, rule.construct
        assert not contains_construct([_module()], rule.construct), rule.construct
    code = _code()
    for banned in ("Percentile", "PERCENTILE", "percentile", "PERCENTILE.INC",
                   "Quartile", "Median("):
        assert banned not in code, banned
    # The executable vocabulary is Quantile.
    assert "Quantile" in code
    assert all("Percentile" not in name for name in _module().procedures)


def test_08_the_module_knows_nothing_of_the_simulation_machinery() -> None:
    code = _code()
    for token in ("SimRng", "SimSample", "SimEngine", "DriverFactors", "Cheng",
                  "MRG32k3a", "Rnd", "Randomize", "_SimData", "SimData", "Results",
                  "Digest", "Fingerprint", "FNV", "FP_VERSION", "RunId", "run_id",
                  "Nonce", "RunSimulation", "Iteration", "Sensitivity"):
        assert token not in code, token
    # Only the accepted Phase-5 numerical primitives are borrowed.
    borrowed = set(re.findall(r"\b(?:Safe\w+|IsUsableDouble|MAX_DOUBLE)\b", code))
    assert borrowed == {"SafeSignedSum", "SafeDivide", "SafeMultiply",
                        "SafeSubtract", "IsUsableDouble"}, sorted(borrowed)


def test_09_no_other_module_owns_a_statistic() -> None:
    for module in load_modules([SRC_VBA]):
        if module.name == "modSimStats":
            continue
        for token in ("Quantile", "Percentile", "StandardDeviation", "Variance",
                      "SIM_QUANTILE", "SIM_STAT_", "Contingency"):
            assert token not in module.code, (module.name, token)


# ===========================================================================
# B. Sorting - copy only, once, and practical at 100,000
# ===========================================================================
def test_10_the_callers_array_is_never_reordered() -> None:
    original = [5.0, 1.0, 4.0, 2.0, 3.0, 2.0]
    snapshot = list(original)
    ok, value, detail = _quantile(original, 0.5)
    assert ok, detail
    assert original == snapshot, "the standalone quantile reordered the caller's array"
    ok, summary, labels, ladder, detail = _describe(original)
    assert ok, detail
    assert original == snapshot, "Describe reordered the caller's array"
    # And the source says so: the sort only ever sees a private copy.
    assert "SimStatsSortAscending(ordered, count, detail)" in _procedure("SimStatsSortedCopy")
    for name in _module().procedures:
        if name in ("SimStatsSortedCopy", "SimStatsSortAscending"):
            continue
        assert "SimStatsSortAscending" not in _procedure(name), name


def test_11_describe_sorts_exactly_once_for_eleven_quantiles() -> None:
    vba = _transcribe()
    real = vba["SimStatsSortAscending"]
    calls = [0]

    def counted(*args, _real=real):
        calls[0] += 1
        return _real(*args)

    vba["SimStatsSortAscending"] = counted
    try:
        ok, summary, labels, ladder, detail = _describe([3.0, 1.0, 2.0, 5.0, 4.0])
    finally:
        vba["SimStatsSortAscending"] = real
    assert ok, detail
    assert len(ladder) == _const("SIM_QUANTILE_COUNT") == 11
    assert calls[0] == 1, f"eleven quantiles cost {calls[0]} sorts"
    # The public standalone entry point is deliberately NOT used in the loop.
    body = _procedure("SimStatsDescribe")
    assert "SimStatsQuantileType7" not in body
    assert "SimStatsQuantileSorted(ordered, count, p, measured, detail)" in body
    assert body.count("SimStatsSortedCopy") == 1


def test_12_the_standalone_quantile_sorts_its_own_copy_once() -> None:
    vba = _transcribe()
    real = vba["SimStatsSortAscending"]
    calls = [0]

    def counted(*args, _real=real):
        calls[0] += 1
        return _real(*args)

    vba["SimStatsSortAscending"] = counted
    try:
        ok, value, detail = _quantile([3.0, 1.0, 2.0], 0.5)
    finally:
        vba["SimStatsSortAscending"] = real
    assert ok and value == 2.0, detail
    assert calls[0] == 1


def test_13_the_sort_is_a_bottom_up_merge_and_not_quadratic() -> None:
    body = _procedure("SimStatsSortAscending")
    assert "ReDim scratch(0 To count - 1)" in body
    assert body.count("ReDim") == 1, "the scratch buffer is allocated more than once"
    assert "runLength = runLength * 2" in body
    assert "Do While runLength < count" in body
    # THE TIE RULE: the left run wins, so equal Doubles keep their arrival order.
    assert "ElseIf series(fromLow) <= series(fromHigh) Then" in body
    # No quadratic pattern: a merge sort has no "shift everything right" inner
    # loop, and no comparison of an element against every earlier one.
    for quadratic in ("For probe = index - 1 To 0 Step -1", "Do While probe >= 0",
                      "series(probe + 1) = series(probe)"):
        assert quadratic not in body, quadratic
    # No library or COM sort of any kind. The module's own two Sort procedures
    # are the only things here that may carry the word.
    code = _code()
    for borrowed in ("ArrayList", "CreateObject", "System.Collections",
                     ".Sort", "Sort(", "SortByKey", "QuickSort"):
        assert borrowed not in code, borrowed
    assert set(re.findall(r"\bSimStats\w*Sort\w*\b", code)) == {
        "SimStatsSortAscending", "SimStatsSortedCopy", "SimStatsQuantileSorted"}


def test_14_the_sort_orders_correctly_at_every_shape() -> None:
    vba = _transcribe()
    import random

    generator = random.Random(20260826)
    shapes = [
        [], [1.0], [2.0, 1.0], [1.0, 2.0], [1.0, 1.0, 1.0],
        [3.0, 1.0, 2.0], list(map(float, range(10))), list(map(float, range(9, -1, -1))),
        [0.0, -0.0, 0.0, -0.0], [-1.7e308, 1.7e308, 0.0, 5e-324],
    ]
    shapes.append([generator.uniform(-1e6, 1e6) for _ in range(257)])
    shapes.append([float(generator.randint(0, 3)) for _ in range(300)])
    for values in shapes:
        if not values:
            continue
        ordered, detail = [], _Ref("")
        assert vba["SimStatsSortedCopy"](
            list(values), _Ref(len(values)), ordered, detail), detail.v
        assert ordered == sorted(values), values[:6]
        assert len(ordered) == len(values)


def test_15_the_sort_is_practical_at_the_design_target() -> None:
    """100,000 values. An O(n^2) sort would be five billion comparisons."""
    import random
    import time

    generator = random.Random(9)
    values = [generator.uniform(-1e6, 1e6) for _ in range(100000)]
    vba = _transcribe()
    ordered, detail = [], _Ref("")
    started = time.monotonic()
    assert vba["SimStatsSortedCopy"](values, _Ref(len(values)), ordered, detail), detail.v
    elapsed = time.monotonic() - started
    assert ordered == sorted(values)
    # The transcription is far slower than VBA would be; the point is the
    # ALGORITHM finishes at all at this size, which insertion sort would not.
    assert elapsed < 240.0, elapsed


# ===========================================================================
# C. Mean, scale and the constant-sample invariant
# ===========================================================================
_HAND_SAMPLES = (
    [2.0, 4.0, 4.0, 4.0, 5.0, 5.0, 7.0, 9.0],
    [1.0, 2.0, 3.0, 4.0],
    [-3.0, 3.0],
    [1.0, 1.0 + 2.0 ** -52],                      # one-ULP varying sample
    [-1.7e308, 1.7e308, 1.7e308, 1.7e308],        # opposite-sign near-MAX
    [-1e308, -1e308, 1e308],
    [5e-324, 1e-323, 1.5e-323],
)

_CONSTANT_SAMPLES = (0.1, 1.1, 1e100, 1.5e308, 5e-324, -12345.678, 0.0)


def test_16_the_mean_matches_the_accepted_oracle_on_every_hand_sample() -> None:
    for values in _HAND_SAMPLES:
        ok, value, detail = _mean(values)
        assert ok, (values[:3], detail)
        assert value == oracle.sample_mean(list(values)), values[:3]


def test_17_a_constant_sample_returns_its_own_value_exactly() -> None:
    """The invariant that stops [1.5e308]*1000 coming back as 1.4999...e308."""
    for repeated in _CONSTANT_SAMPLES:
        values = [repeated] * 1000
        ok, value, detail = _mean(values)
        assert ok, (repeated, detail)
        assert value == repeated, (repeated, value)
        assert value == oracle.sample_mean(values)
        ok, deviation, detail = _sd(values)
        assert ok, (repeated, detail)
        assert deviation == 0.0 and math.copysign(1.0, deviation) > 0.0, repeated
        assert deviation == oracle.sample_standard_deviation(values)
        ok, summary, labels, ladder, detail = _describe(values)
        assert ok, (repeated, detail)
        assert summary["Mean"] == repeated
        assert summary["SampleStandardDeviation"] == 0.0
        assert summary["Minimum"] == summary["Maximum"] == repeated
        assert set(ladder) == {repeated}, (repeated, sorted(set(ladder))[:3])
    # The shortcut fires BEFORE any accumulation, in the source too.
    body = _procedure("SimStatsMean")
    assert body.index("SimStatsConstantValue") < body.index("SimStatsUnitScale")
    body = _procedure("SimStatsSampleStandardDeviation")
    assert body.index("SimStatsConstantValue") < body.index("SimStatsUnitScale")


def test_18_the_naive_mean_is_what_the_invariant_prevents() -> None:
    """The control is not vacuous: the forbidden path gives a different answer."""
    values = [1.5e308] * 1000
    naive = 0.0
    for value in values:
        naive += value / len(values)
    assert naive != 1.5e308, "the naive path no longer differs; the fixture is stale"
    ok, value, detail = _mean(values)
    assert ok and value == 1.5e308, detail


def test_19_the_scale_is_the_largest_power_of_two_not_exceeding_the_sample() -> None:
    vba = _transcribe()

    def scale_of(values):
        result, detail = _Ref(0.0), _Ref("")
        ok = vba["SimStatsUnitScale"](list(values), _Ref(len(values)), result, detail)
        return ok, result.v, detail.v

    checks = [
        ([1.0], 1.0),
        ([1.5], 1.0),
        ([2.0], 2.0),
        ([3.9999], 2.0),
        ([-7.0, 1.0], 4.0),
        ([0.0, 0.0], 0.0),
        ([5e-324], 5e-324),
        ([1e-323], 1e-323),
        ([1.5e-323], 1e-323),
        ([1.5e308], math.ldexp(1.0, 1023)),
        ([sys.float_info.max], math.ldexp(1.0, 1023)),
        ([-1.7e308, 1.7e308], math.ldexp(1.0, 1023)),
        ([sys.float_info.min], sys.float_info.min),
    ]
    for values, expected in checks:
        ok, value, detail = scale_of(values)
        assert ok, (values, detail)
        assert value == expected, (values, value, expected)
        # ...and it agrees with the accepted oracle's own selection.
        assert value == oracle._scale_of([float(v) for v in values]), values
    # Every scale is a finite positive power of two not exceeding the magnitude.
    for values, expected in checks:
        if expected == 0.0:
            continue
        assert math.isfinite(expected) and expected > 0.0
        mantissa, _ = math.frexp(expected)
        assert mantissa == 0.5, expected
        assert expected <= max(abs(v) for v in values)
    # NO Log-BASED EXTRACTION AND NO 2 ^ exponent PATH.
    body = _procedure("SimStatsUnitScale")
    for banned in ("Log(", "Exp(", "^", "Fix(", "Int("):
        assert banned not in body, banned
    assert "candidate = candidate / 2#" in body
    assert "candidate = candidate * 2#" in body
    # The doubling test cannot overflow: it divides the magnitude, not the power.
    assert "Do While candidate <= largest / 2#" in body
    assert "candidate * 2# <= largest" not in body


def test_20_the_mean_uses_the_original_order_not_the_sorted_copy() -> None:
    """SafeSignedSum accumulates left to right, so order moves the last bits."""
    original = [1.0, 1e16, -1e16, 1.0, 1.0]
    ordered = sorted(original)
    assert safe_signed_sum(original, "fixture") != safe_signed_sum(ordered, "fixture"), (
        "the fixture is order-insensitive, so an order mutation could not be seen"
    )
    ok, value, detail = _mean(original)
    assert ok, detail
    assert value == oracle.sample_mean(original)
    assert value != oracle.sample_mean(ordered), "the mean was taken over sorted values"
    ok, summary, labels, ladder, detail = _describe(original)
    assert ok, detail
    assert summary["Mean"] == oracle.sample_mean(original)
    # And in the source: Describe hands `values`, never `ordered`, to the moments.
    body = _procedure("SimStatsDescribe")
    assert "SimStatsMean(values, count, measured, detail)" in body
    assert "SimStatsSampleStandardDeviation(values, count, measured, detail)" in body
    assert "SimStatsMean(ordered" not in body
    assert "SimStatsSampleStandardDeviation(ordered" not in body


# ===========================================================================
# D. Sample standard deviation
# ===========================================================================
def test_21_the_sample_deviation_matches_the_accepted_oracle() -> None:
    for values in _HAND_SAMPLES:
        ok, value, detail = _sd(values)
        assert ok, (values[:3], detail)
        assert value == oracle.sample_standard_deviation(list(values)), values[:3]


def test_22_the_divisor_is_n_minus_one_and_never_n() -> None:
    values = [2.0, 4.0, 4.0, 4.0, 5.0, 5.0, 7.0, 9.0]
    ok, value, detail = _sd(values)
    assert ok, detail
    # The textbook sample deviation of this sample is sqrt(32/7); the POPULATION
    # deviation is 2.0 exactly, so the two are far apart on this fixture.
    assert abs(value - math.sqrt(32.0 / 7.0)) < 1e-12, value
    assert abs(value - 2.0) > 1e-3, "the population divisor would give exactly 2"
    body = _procedure("SimStatsSampleStandardDeviation")
    assert "SafeDivide(residual, CDbl(count - 1), variance)" in body
    assert "CDbl(count), variance" not in body
    assert _const("SIM_STAT_SD_DIVISOR") == "n_minus_1"


def test_23_fewer_than_two_observations_refuses() -> None:
    for values in ([], [7.0]):
        result, detail = _Ref(-1.0), _Ref("")
        ok = _transcribe()["SimStatsSampleStandardDeviation"](
            list(values), _Ref(len(values)), result, detail)
        assert ok is False, values
        assert "at least two observations" in detail.v
        assert result.v == -1.0
    # The mean needs only one, and Describe needs two.
    ok, value, detail = _mean([7.0])
    assert ok and value == 7.0, detail
    summary = _transcribe()["_new"]("SimStatsMeasure")
    labels, ladder, detail = [], [], _Ref("")
    assert _transcribe()["SimStatsDescribe"](
        [7.0], _Ref(1), summary, labels, ladder, detail) is False
    assert "at least two observations" in detail.v


def test_24_neither_forbidden_deviation_path_exists() -> None:
    body = _procedure("SimStatsSampleStandardDeviation")
    # Two passes in the NORMALISED space, and the deviation is formed there.
    assert "deviation = scaled(index) - centre" in body
    assert "squares(index) = deviation * deviation" in body
    # No sum of squares of the ORIGINAL values, and no Welford recurrence.
    for banned in ("values(LBound(values) + index) * values(", "Welford", "delta",
                   "runningMean", "m2", "M2"):
        assert banned not in body, banned
    assert "SafeSignedSum(squares, count, residual)" in body
    # The load-bearing fixture: every forbidden path overflows, the accepted one
    # does not.
    values = [-1.7e308, 1.7e308, 1.7e308, 1.7e308]
    assert not math.isfinite(sum(v * v for v in values)), "the fixture is stale"
    naive_mean = 8.5e307
    assert not math.isfinite(values[0] - naive_mean), "the fixture is stale"
    ok, value, detail = _sd(values)
    assert ok, detail
    assert value == oracle.sample_standard_deviation(values)
    assert math.isfinite(value) and abs(value - 1.7e308) <= 1e-14 * 1.7e308, value


def test_25_a_varying_subnormal_dispersion_refuses_rather_than_reporting_zero() -> None:
    """The accepted Step-4 finding, kept.

    `[5e-324] * 999 + [1e-323]` genuinely varies, so the constant shortcut must
    not fire; and its true sample deviation is below the smallest positive
    Double. Reporting `0` would state that a sample which demonstrably varies has
    no dispersion at all.
    """
    values = [5e-324] * 999 + [1e-323]
    assert len(set(values)) == 2, "the fixture stopped varying"
    try:
        oracle.sample_standard_deviation(values)
    except Exception:  # noqa: BLE001 - the oracle's own refusal
        pass
    else:  # pragma: no cover
        raise AssertionError("the accepted oracle no longer refuses this sample")
    result, detail = _Ref(-1.0), _Ref("")
    ok = _transcribe()["SimStatsSampleStandardDeviation"](
        values, _Ref(len(values)), result, detail)
    assert ok is False, "a varying subnormal dispersion was answered"
    assert "rescale" in detail.v, detail.v
    assert result.v == -1.0, "a refused deviation still wrote its output"
    # ...while the genuinely constant subnormal sample is exactly zero.
    ok, value, detail = _sd([5e-324] * 1000)
    assert ok and value == 0.0, detail


def test_26_a_non_finite_observation_is_refused_not_skipped() -> None:
    for poison in (float("nan"), float("inf"), float("-inf")):
        values = [1.0, 2.0, poison, 4.0]
        for call in (_mean, _sd):
            ok, value, detail = call(values)
            assert ok is False, (poison, call.__name__)
            assert "not a finite Double" in detail
        ok, value, detail = _quantile(values, 0.5)
        assert ok is False and "not a finite Double" in detail
        ok, summary, labels, ladder, detail = _describe(values)
        assert ok is False and "not a finite Double" in detail


def test_27_an_empty_sequence_refuses_before_any_bound_is_read() -> None:
    """A zero-count caller may hand over an array that was never sized."""
    body = _procedure("SimStatsUsableSequence")
    assert body.index("If count = 0 Then") < body.index("LBound(values)")
    assert body.index("If count < 0 Then") < body.index("LBound(values)")
    for call in (_mean, _sd):
        ok, value, detail = call([])
        assert ok is False, call.__name__
    ok, value, detail = _quantile([], 0.5)
    assert ok is False and "empty" in detail
    # Each public entry point checks the count before it validates the sequence.
    for name, minimum in (("SimStatsMean", "count < 1"),
                          ("SimStatsSampleStandardDeviation", "count < 2"),
                          ("SimStatsQuantileType7", "count < 1"),
                          ("SimStatsDescribe", "count < 2")):
        body = _procedure(name)
        assert f"If {minimum} Then" in body, name
        assert body.index(minimum) < body.index("SimStatsUsableSequence"), name


# ===========================================================================
# E. Type-7 quantiles
# ===========================================================================
_QUANTILE_SHAPES = (
    [7.0],
    [1.0, 3.0],
    [1.0, 2.0, 4.0],
    [1.0, 2.0, 3.0, 4.0],
    [float(v) for v in range(1, 11)],
    [-1.7e308, 1.7e308],
    [5.0] * 7,
    [2.0, 4.0, 4.0, 4.0, 5.0, 5.0, 7.0, 9.0],
)
_PROBABILITIES = (0.0, 0.1, 0.25, 1.0 / 3.0, 0.5, 0.55, 0.7, 0.75, 0.9, 0.95, 1.0)


def test_28_type_7_matches_the_accepted_oracle_everywhere() -> None:
    assert _const("SIM_QUANTILE_METHOD") == "hyndman_fan_type_7"
    for values in _QUANTILE_SHAPES:
        for p in _PROBABILITIES:
            ok, value, detail = _quantile(values, p)
            assert ok, (values[:3], p, detail)
            assert value == oracle.percentile_type7(list(values), p), (values[:3], p)


def test_29_an_integral_h_returns_an_exact_order_statistic() -> None:
    values = [float(v) for v in range(1, 11)]
    ordered = sorted(values)
    for index in range(len(values)):
        p = index / (len(values) - 1)
        ok, value, detail = _quantile(values, p)
        assert ok, detail
        h = (len(values) - 1) * p
        if h == math.floor(h):
            assert value == ordered[index], (p, value)
    # p = 0 is the minimum and p = 1 the maximum, exactly, at every magnitude.
    for sample in ([5e-324, 1e-323, 1.5e-323], [-1.7e308, 0.0, 1.7e308]):
        assert _quantile(sample, 0.0)[1] == min(sample)
        assert _quantile(sample, 1.0)[1] == max(sample)
    body = _procedure("SimStatsQuantileSorted")
    assert "If fraction = 0# Then" in body
    assert body.index("If fraction = 0# Then") < body.index("ElseIf low = high Then")


def test_30_an_equal_bracket_returns_that_value_exactly() -> None:
    """0.7 * 0.1 + 0.3 * 0.1 is 0.10000000000000002; the ladder must not creep."""
    interpolated = 0.7 * 0.1 + 0.3 * 0.1
    assert interpolated != 0.1, "the fixture no longer demonstrates the drift"
    for value in (0.1, 1.1, 1e100, 1.5e308, 5e-324, -12345.678):
        ok, result, detail = _quantile([value] * 7, 0.55)
        assert ok, (value, detail)
        assert result == value, (value, result)
    assert "ElseIf low = high Then" in _procedure("SimStatsQuantileSorted")


def test_31_the_interpolation_is_the_convex_form() -> None:
    body = _procedure("SimStatsQuantileSorted")
    assert "candidate = (1# - fraction) * low + fraction * high" in body
    assert "high - low" not in body and "(high - low)" not in body
    # The point of the convex form: the difference does not exist here.
    assert not math.isfinite(1.7e308 - -1.7e308), "the fixture is stale"
    ok, value, detail = _quantile([-1.7e308, 1.7e308], 0.25)
    assert ok, detail
    assert math.isfinite(value)
    assert value == oracle.percentile_type7([-1.7e308, 1.7e308], 0.25)
    assert -1.7e308 <= value <= 1.7e308


def test_32_the_probability_domain_is_closed_and_refused_outside() -> None:
    body = _procedure("SimStatsUsableProbability")
    assert "If p < 0# Or p > 1# Then" in body
    for p in (-0.001, 1.001, -1.0, 2.0, float("nan"), float("inf")):
        result, detail = _Ref(-1.0), _Ref("")
        ok = _transcribe()["SimStatsQuantileType7"](
            [1.0, 2.0, 3.0], _Ref(3), _Ref(p), result, detail)
        assert ok is False, p
        assert result.v == -1.0
    # No nearest-rank method and no (n + 1) * p anywhere.
    quantile = _procedure("SimStatsQuantileSorted")
    assert "h = CDbl(count - 1) * p" in quantile
    for banned in ("count + 1", "(count + 1)", "Round(", "CInt(", "nearest"):
        assert banned not in quantile, banned


# ===========================================================================
# F. The ladder and the selected level
# ===========================================================================
def test_33_the_ladder_is_the_projection_in_the_projected_order() -> None:
    ok, summary, labels, ladder, detail = _describe([float(v) for v in range(1, 21)])
    assert ok, detail
    expected = [label for label, _ in _ladder_points()]
    assert labels == expected, labels
    assert expected == ["P10", "P50", "P55", "P60", "P65", "P70", "P75", "P80",
                        "P85", "P90", "P95"], expected
    assert len(ladder) == summary["QuantileCount"] == _const("SIM_QUANTILE_COUNT") == 11
    assert len(set(labels)) == len(labels), "the ladder repeats a label"
    # No second list: the module names only the projected constants.
    body = _procedure("SimStatsLadderLabel")
    for index in range(1, 12):
        assert f"SIM_QUANTILE_{index}" in body, index
    assert not re.search(r'"P\d', _module().raw), "a ladder label is spelled in the source"


def test_34_the_whole_ladder_matches_the_accepted_oracle() -> None:
    for values in ([float(v) for v in range(1, 21)],
                   [2.0, 4.0, 4.0, 4.0, 5.0, 5.0, 7.0, 9.0],
                   [-1.7e308, 0.0, 1.7e308, 5e-324],
                   [3.5] * 40):
        ok, summary, labels, ladder, detail = _describe(values)
        assert ok, (values[:3], detail)
        expected = oracle.describe(list(values), _ladder_points())
        assert summary["Count"] == expected.count
        assert summary["Mean"] == expected.mean
        assert summary["SampleStandardDeviation"] == expected.sample_standard_deviation
        assert summary["Minimum"] == expected.minimum
        assert summary["Maximum"] == expected.maximum
        assert dict(zip(labels, ladder)) == dict(expected.percentiles), values[:3]


def test_35_the_headline_levels_are_lookups_from_the_same_ladder() -> None:
    values = [float(v) for v in range(1, 41)]
    ok, summary, labels, ladder, detail = _describe(values)
    assert ok, detail
    for index in range(1, 5):
        headline = _const(f"SIM_QUANTILE_HEADLINE_{index}")
        assert headline in labels, headline
        position = labels.index(headline)
        assert ladder[position] == oracle.percentile_type7(
            values, float(headline[1:]) / 100.0), headline
    # There is no second quantile engine: Describe holds one ladder loop.
    body = _procedure("SimStatsDescribe")
    assert body.count("SimStatsQuantileSorted") == 1
    assert "HEADLINE" not in _code(), "a headline pass was added"


def test_36_the_fixed_level_is_reported_and_not_selectable() -> None:
    values = [float(v) for v in range(1, 21)]
    ok, summary, labels, ladder, detail = _describe(values)
    assert ok, detail
    fixed = _const("SIM_QUANTILE_FIXED_1")
    assert fixed == "P10"
    # PRESENT in the ladder...
    assert fixed in labels
    assert ladder[labels.index(fixed)] == oracle.percentile_type7(values, 0.10)
    # ...and refused as a selector.
    ok, value, detail = _selected(labels, ladder, fixed)
    assert ok is False
    assert "not selectable" in detail
    # Every other rung IS selectable, and returns the stored value.
    for position, label in enumerate(labels):
        if label == fixed:
            continue
        ok, value, detail = _selected(labels, ladder, label)
        assert ok, (label, detail)
        assert value == ladder[position], label


def test_37_an_unknown_confidence_level_is_refused() -> None:
    ok, summary, labels, ladder, detail = _describe([float(v) for v in range(1, 21)])
    assert ok, detail
    for label in ("P99", "P0", "p50", "P50 ", "", "MEAN", "UNSELECTED", "P100"):
        result, refusal = _Ref(-1.0), _Ref("")
        accepted = _transcribe()["SimStatsSelectedQuantile"](
            labels, ladder, _Ref(len(labels)), _Ref(label), result, refusal)
        assert accepted is False, label
        assert result.v == -1.0, label
    # And no UNSELECTED state was invented.
    assert "UNSELECTED" not in _code()
    # Selection reaches no simulation machinery at all.
    body = _procedure("SimStatsSelectedQuantile")
    for banned in ("SimStatsSortAscending", "SimStatsQuantileSorted", "SimStatsDescribe",
                   "SimStatsMean", "ReDim"):
        assert banned not in body, banned


# ===========================================================================
# G. Contingency
# ===========================================================================
def test_38_contingency_is_the_selected_total_minus_the_deterministic_base() -> None:
    assert _const("SIM_CONTINGENCY_FORMULA") == \
        "selected_px_total - deterministic_base_estimate_a"
    assert _const("SIM_CONTINGENCY_BASELINE") == "deterministic_base_estimate_a"
    for selected, base, expected in (
        (1250.0, 1000.0, 250.0),
        (1000.0, 1000.0, 0.0),
        (0.25, 0.125, 0.125),
        (-40.0, -100.0, 60.0),
    ):
        ok, value, detail = _contingency(selected, base)
        assert ok, (selected, base, detail)
        assert value == expected, (selected, base, value)
    body = _procedure("SimStatsContingency")
    assert "SafeSubtract(selectedTotal, baseEstimateA, candidate)" in body
    assert "selectedTotal - baseEstimateA" not in body
    # The baseline is A and nothing else.
    for banned in ("Mean", "Expected", "EMV", "Analytical"):
        assert banned not in body, banned


def test_39_a_negative_contingency_is_preserved_and_never_clamped() -> None:
    for selected, base, expected in ((900.0, 1000.0, -100.0),
                                     (-1e6, 1e6, -2e6),
                                     (0.0, 5e-324, -5e-324)):
        ok, value, detail = _contingency(selected, base)
        assert ok, (selected, base, detail)
        assert value == expected, (selected, base, value)
        assert value < 0.0
    body = _procedure("SimStatsContingency")
    for clamp in ("If candidate < 0#", "If result < 0#", "Then candidate = 0#"):
        assert clamp not in body, clamp


def test_40_the_extreme_contingency_domain_behaves_as_stated() -> None:
    # A. representable
    ok, value, detail = _contingency(1.0e308, -5.0e307)
    assert ok, detail
    assert value == 1.5e308
    assert math.isfinite(value)
    # B. unrepresentable - refused, not infinite, not clamped
    result, refusal = _Ref(-1.0), _Ref("")
    accepted = _transcribe()["SimStatsContingency"](
        _Ref(1.0e308), _Ref(-1.0e308), result, refusal)
    assert accepted is False
    assert "contingency" in refusal.v and "not representable" in refusal.v
    assert result.v == -1.0, "a refused contingency still wrote its output"
    assert not math.isfinite(1.0e308 - -1.0e308), "the fixture is stale"
    # A non-finite operand is refused at the boundary.
    for selected, base in ((float("inf"), 0.0), (0.0, float("nan"))):
        result, refusal = _Ref(-1.0), _Ref("")
        assert _transcribe()["SimStatsContingency"](
            _Ref(selected), _Ref(base), result, refusal) is False
        assert "not a finite Double" in refusal.v


def test_41_contingency_composes_with_the_selected_level() -> None:
    values = [float(v) for v in range(1, 101)]
    ok, summary, labels, ladder, detail = _describe(values)
    assert ok, detail
    base = 40.0
    for label in labels:
        if label == _const("SIM_QUANTILE_FIXED_1"):
            continue
        ok, selected, detail = _selected(labels, ladder, label)
        assert ok, (label, detail)
        ok, value, detail = _contingency(selected, base)
        assert ok, (label, detail)
        assert value == selected - base, label
    # Nominal and PV are described separately and never mixed: the module holds
    # one measure per call and knows no second one.
    assert "Nominal" not in _code() and "Pv" not in _code() and "PV" not in _code()


# ===========================================================================
# H. Transactional outputs
# ===========================================================================
def test_42_a_refused_description_publishes_nothing() -> None:
    vba = _transcribe()
    real = vba["SimStatsQuantileSorted"]
    calls = [0]

    def failing(ordered, count, p, result, detail, _real=real):
        calls[0] += 1
        if calls[0] == _const("SIM_QUANTILE_COUNT"):     # the ELEVENTH rung
            detail.v = "statistics: injected failure"
            return False
        return _real(ordered, count, p, result, detail)

    summary = vba["_new"]("SimStatsMeasure")
    before = dict(summary)
    labels, ladder, detail = ["untouched"], [-1.0], _Ref("")
    vba["SimStatsQuantileSorted"] = failing
    try:
        ok = vba["SimStatsDescribe"](
            [float(v) for v in range(1, 21)], _Ref(20), summary, labels, ladder, detail)
    finally:
        vba["SimStatsQuantileSorted"] = real
    assert ok is False
    assert calls[0] == _const("SIM_QUANTILE_COUNT")
    assert detail.v == "statistics: injected failure"
    assert summary == before, "a partial summary was published"
    assert labels == ["untouched"] and ladder == [-1.0], "a partial ladder was published"
    # The source commits all three together, after the loop.
    body = _procedure("SimStatsDescribe")
    for commit in ("summary = candidate", "quantileLabels = labels",
                   "quantileValues = ladder"):
        assert commit in body, commit
        assert body.index("Next index") < body.index(commit), commit


def test_43_every_refusal_leaves_its_scalar_output_alone() -> None:
    vba = _transcribe()
    for name, args in (
        ("SimStatsMean", ([], _Ref(0))),
        ("SimStatsSampleStandardDeviation", ([1.0], _Ref(1))),
        ("SimStatsQuantileType7", ([1.0], _Ref(1), _Ref(2.0))),
    ):
        result, detail = _Ref(-99.0), _Ref("")
        assert vba[name](*args, result, detail) is False, name
        assert result.v == -99.0, name
        assert detail.v != "", name


def test_44_a_refused_selection_leaves_its_output_alone() -> None:
    ok, summary, labels, ladder, detail = _describe([float(v) for v in range(1, 21)])
    assert ok, detail
    result, refusal = _Ref(-99.0), _Ref("")
    assert _transcribe()["SimStatsSelectedQuantile"](
        labels, ladder, _Ref(len(labels)), _Ref("P99"), result, refusal) is False
    assert result.v == -99.0
    # A ladder of the wrong length is refused before it is searched.
    result, refusal = _Ref(-99.0), _Ref("")
    assert _transcribe()["SimStatsSelectedQuantile"](
        labels[:3], ladder[:3], _Ref(3), _Ref("P50"), result, refusal) is False
    assert "not the accepted length" in refusal.v


def test_45_the_sorted_copy_is_never_published() -> None:
    """It exists during the calculation and is not a third retained result."""
    body = _procedure("SimStatsDescribe")
    assert "Dim ordered() As Double" in body
    # `ordered` is never assigned to any ByRef output of the procedure.
    for output in ("summary", "quantileLabels", "quantileValues"):
        assert f"{output} = ordered" not in body, output
    signature = logical_statements(body)[0][1]
    assert "ordered" not in signature
    # No public procedure hands a sorted sequence back.
    for name in _module().public_procedures:
        assert "ordered() As Double" not in logical_statements(_procedure(name))[0][1], name


# ===========================================================================
# I. The accepted corpus
# ===========================================================================
def test_46_the_corpus_moment_cases_are_honoured() -> None:
    cases = _cases()
    hand = cases["statistics.mean.hand_vector"]
    values = [float(v) for v in hand["inputs"]["values"]]
    assert hand["comparison"] == "EXACT"
    ok, value, detail = _mean(values)
    assert ok, detail
    assert value == hand["expected_exact"]["mean"] == 5.0

    deviation = cases["statistics.sd.hand_vector"]
    assert deviation["comparison"] == "TOLERANCE_BOUNDED"
    assert deviation["inputs"]["divisor"] == "n_minus_1"
    ok, value, detail = _sd([float(v) for v in deviation["inputs"]["values"]])
    assert ok, detail
    want = deviation["expected"]["sample_standard_deviation"]
    assert abs(value - want) <= 1e-12 * abs(want), (value, want)
    # ...and the population divisor would give a different, EXACT number.
    assert value != deviation["expected_exact"]["population_standard_deviation_would_be"]

    zero = cases["statistics.sd.zero_variance"]
    assert zero["comparison"] == "EXACT"
    values = [float(v) for v in zero["inputs"]["values"]]
    assert _mean(values)[1] == zero["expected_exact"]["mean"]
    assert _sd(values)[1] == zero["expected_exact"]["sample_standard_deviation"] == 0.0

    refused = cases["statistics.sd.refused_below_two"]
    assert refused["comparison"] == "EXACT"
    assert _sd([1.0])[0] is False


def test_46a_the_corpus_type_7_rows_are_reproduced() -> None:
    for key in ("n1", "n2", "n3", "n4", "n10"):
        case = _cases()[f"statistics.quantile.type7.{key}"]
        values = [float(v) for v in case["inputs"]["values"]]
        assert case["expected_exact"]["interpolation"] == "convex"
        assert case["expected_exact"]["sorting"] == "on copies only"
        for row in case["expected"]["rows"]:
            ok, value, detail = _quantile(values, row["p"])
            assert ok, (key, row["p"], detail)
            if row["comparison"] == "EXACT":
                assert value == row["value"], (key, row)
            else:
                assert abs(value - row["value"]) <= 1e-12 * max(abs(row["value"]), 1.0), (key, row)
            # The formula's own intermediate positions agree too.
            count = len(values)
            h = (count - 1) * row["p"]
            assert h == row["h"], (key, row)
            assert math.floor(h) == row["lo"], (key, row)
            assert min(row["lo"] + 1, count - 1) == row["hi"], (key, row)


def test_46b_the_corpus_ladder_and_constant_sample_rows_are_reproduced() -> None:
    ladder = _cases()["statistics.ladder.resolved"]["expected_exact"]
    labels = [label for label, _ in _ladder_points()]
    assert labels == ladder["ordered"]
    assert len(labels) == ladder["count"] == _const("SIM_QUANTILE_COUNT")
    assert ladder["fixed_non_selectable"] == [_const("SIM_QUANTILE_FIXED_1")]
    assert ladder["headline"] == [_const(f"SIM_QUANTILE_HEADLINE_{i}") for i in range(1, 5)]
    selectable = set(ladder["selectable"])
    assert selectable == set(labels) - set(ladder["fixed_non_selectable"])
    # Every accepted p is what the label parser reads out of the label.
    for point in ladder["points"]:
        assert dict(_ladder_points())[point["label"]] == point["p"], point

    constant = _cases()["statistics.constant_sample.zero_dispersion"]
    assert constant["comparison"] == "EXACT"
    for row in constant["expected_exact"]["rows"]:
        values = [row["value"]] * constant["inputs"]["count"]
        ok, summary, seen, rungs, detail = _describe(values)
        assert ok, (row["value"], detail)
        assert summary["Count"] == row["count"]
        assert summary["Mean"] == row["mean"]
        assert summary["SampleStandardDeviation"] == row["sample_standard_deviation"]
        assert summary["Minimum"] == row["minimum"]
        assert summary["Maximum"] == row["maximum"]
        assert (set(rungs) == {row["value"]}) is row["all_quantiles_equal_the_value"]


def test_46c_the_corpus_scale_safety_and_contingency_cases_are_reproduced() -> None:
    near = _cases()["statistics.scale_safety.near_maximum"]
    values = [float(v) for v in near["inputs"]["values"]]
    exact = near["expected_exact"]
    assert exact["naive_sum_overflows"] and exact["naive_sum_of_squares_overflows"]
    assert exact["unguarded_deviation_overflows"] and exact["finite"]
    ok, value, detail = _mean(values)
    assert ok, detail
    assert value == near["expected"]["mean"] == 8.5e307
    ok, value, detail = _sd(values)
    assert ok, detail
    assert value == near["expected"]["sample_standard_deviation"]

    unrepresentable = _cases()["statistics.scale_safety.unrepresentable_dispersion"]
    values = [5e-324] * 999 + [1e-323]
    assert len(values) == unrepresentable["inputs"]["count"]
    assert unrepresentable["expected_exact"]["mean_is_representable"] is True
    ok, value, detail = _mean(values)
    assert ok, detail
    assert value == unrepresentable["expected_exact"]["mean"]
    assert _sd(values)[0] is False
    assert "rescale" in unrepresentable["expected_refusal"]["stage"]

    negative = _cases()["contingency.negative_not_clamped"]["expected_exact"]
    ok, value, detail = _contingency(
        negative["selected_nominal"],
        _cases()["contingency.negative_not_clamped"]["inputs"]["deterministic_base_nominal"])
    assert ok, detail
    assert value == negative["contingency_nominal"] == -992.0
    assert negative["is_negative"] is True and negative["clamped"] is False

    unrep = _cases()["contingency.unrepresentable_subtraction"]
    neighbour = unrep["expected_exact"]["representable_neighbour"]
    ok, value, detail = _contingency(
        neighbour["selected_nominal"], neighbour["deterministic_base_nominal"])
    assert ok, detail
    assert value == neighbour["contingency_nominal"]
    assert _contingency(unrep["inputs"]["selected_nominal"],
                        unrep["inputs"]["deterministic_base_nominal"])[0] is False

    not_selectable = _cases()["contingency.p10_not_selectable"]
    assert not_selectable["inputs"]["selector"] == _const("SIM_QUANTILE_FIXED_1")
    assert not_selectable["inputs"]["reported"] is True
    ok, summary, seen, rungs, detail = _describe([float(v) for v in range(1, 21)])
    assert ok, detail
    assert _selected(seen, rungs, not_selectable["inputs"]["selector"])[0] is False


def test_47_the_transcription_read_the_whole_module() -> None:
    """Every procedure is COMPILED FROM SOURCE except the one guarded-bounds
    helper, whose body no `On Error`-free engine can execute. That one is named
    here, its real signature is read out of the module, and `test_48` proves its
    source shape - so nothing can hide in a body no test looks at."""
    vba = _transcribe()
    compiled = vba["_python_source"]
    assert BORROWED_FROM_MODULE == {"SimStatsLadderExtent"}, BORROWED_FROM_MODULE
    for name in _module().procedures:
        assert callable(vba[name]), name
        if name in BORROWED_FROM_MODULE:
            assert f"def {name}(" not in compiled, f"{name} claims to be borrowed"
        else:
            assert f"def {name}(" in compiled, f"{name} was not compiled from source"
    # The borrowed signature is the MODULE'S OWN declaration, not one retyped here.
    assert [(mode, pname, arr) for mode, pname, arr, _, _
            in vba["_procs"]["SimStatsLadderExtent"]] == [
        ("ByRef", "quantileLabels", True), ("ByRef", "quantileValues", True),
        ("ByRef", "labelExtent", False), ("ByRef", "valueExtent", False)]
    assert "SimStatsMeasure" in vba["_types"]
    for borrowed in ("IsUsableDouble", "SafeSignedSum", "SafeDivide",
                     "SafeMultiply", "SafeSubtract"):
        assert callable(vba[borrowed]), borrowed


# ===========================================================================
# J. Ladder integrity at the selection boundary
#
# The ladder arrays are ordinary caller-writable VBA arrays. `SimStatsDescribe`
# is their authoritative constructor; `SimStatsSelectedQuantile` proves their
# structure before reading one. What that CANNOT prove is that a finite value
# was not edited after Describe produced it - see `test_63`.
# ===========================================================================
def _genuine_ladder():
    ok, summary, labels, ladder, detail = _describe([float(v) for v in range(1, 21)])
    assert ok, detail
    return list(labels), list(ladder)


def _selectable_labels():
    return [label for label, _ in _ladder_points()
            if label != _const("SIM_QUANTILE_FIXED_1")]


def _refused(labels, ladder, label, count=None):
    """Refusal, with the caller's result Double left exactly as it was."""
    sentinel = _Ref(-987654.5)
    ok, _value, detail = _selected(labels, ladder, label, count=count, result=sentinel)
    assert ok is False, f"{label} was accepted on a malformed ladder"
    assert sentinel.v == -987654.5, "a refusal wrote to the selected output"
    assert detail, "a refusal carried no reason"
    return detail


def test_48_the_guarded_bounds_helper_is_scoped_and_reads_only_the_extents() -> None:
    # EXECUTABLE code, comments stripped: the module is allowed to say in prose
    # which construct it refuses to contain, exactly as Step 8 settled for Cheng.
    executable = _module().code_without_string_removal
    assert "On Error Resume Next" not in executable, "a broad suppression was introduced"
    # ONE scoped handler in the whole module, and it is this one.
    assert re.findall(r"On Error GoTo (\w+)", executable) == ["Unallocated", "0", "0"]
    body = _procedure("SimStatsLadderExtent")
    statements = [text for _, text in logical_statements(body)]
    assert statements == [
        'Private Function SimStatsLadderExtent(ByRef quantileLabels() As String, '
        'ByRef quantileValues() As Double, ByRef labelExtent As Long, '
        'ByRef valueExtent As Long) As Boolean',
        "On Error GoTo Unallocated",
        "labelExtent = UBound(quantileLabels) - LBound(quantileLabels) + 1",
        "valueExtent = UBound(quantileValues) - LBound(quantileValues) + 1",
        "On Error GoTo 0",
        "SimStatsLadderExtent = True",
        "Exit Function",
        "Unallocated:",
        "On Error GoTo 0",
        "SimStatsLadderExtent = False",
        "End Function",
    ], statements
    # NOTE: the arm that RAISES - a genuinely never-sized VBA array - has no
    # Linux execution proof and is deferred to Gate B. Its SHAPE is pinned above.


def test_49_a_genuine_ladder_returns_every_selectable_level_unchanged() -> None:
    labels, ladder = _genuine_ladder()
    for label in _selectable_labels():
        ok, value, detail = _selected(labels, ladder, label)
        assert ok, f"{label}: {detail}"
        assert value == ladder[labels.index(label)], label
    assert len(_selectable_labels()) == _const("SIM_QUANTILE_COUNT") - 1


def test_50_the_fixed_rung_is_still_refused_on_a_genuine_ladder() -> None:
    labels, ladder = _genuine_ladder()
    detail = _refused(labels, ladder, _const("SIM_QUANTILE_FIXED_1"))
    assert "not selectable" in detail, detail


def test_51_an_invented_rung_cannot_be_selected_by_inserting_it() -> None:
    """THE DEFECT THIS GROUP EXISTS FOR: membership in the caller's own array is
    not evidence that a label is an accepted confidence level."""
    labels, ladder = _genuine_ladder()
    forged, values = list(labels), list(ladder)
    forged[labels.index("P50")] = "P42"
    values[labels.index("P50")] = 4242.0
    detail = _refused(forged, values, "P42")
    assert "accepted projection" in detail, detail


def test_52_a_forged_rung_refuses_the_whole_ladder_not_just_that_rung() -> None:
    labels, ladder = _genuine_ladder()
    forged, values = list(labels), list(ladder)
    forged[labels.index("P85")] = "P42"
    values[labels.index("P85")] = 4242.0
    # P75 is untouched and would otherwise answer. The LADDER is malformed.
    _refused(forged, values, "P75")


def test_53_two_swapped_rungs_are_refused() -> None:
    labels, ladder = _genuine_ladder()
    forged, values = list(labels), list(ladder)
    low, high = labels.index("P55"), labels.index("P60")
    forged[low], forged[high] = forged[high], forged[low]
    values[low], values[high] = values[high], values[low]
    _refused(forged, values, "P55")
    _refused(forged, values, "P90")


def test_54_a_duplicated_rung_is_refused() -> None:
    labels, ladder = _genuine_ladder()
    forged, values = list(labels), list(ladder)
    forged[labels.index("P55")] = "P50"
    values[labels.index("P55")] = values[labels.index("P50")]
    _refused(forged, values, "P50")
    _refused(forged, values, "P90")


def test_55_a_missing_fixed_rung_is_refused() -> None:
    labels, ladder = _genuine_ladder()
    forged, values = list(labels), list(ladder)
    forged[labels.index("P10")] = "P05"
    _refused(forged, values, "P50")


def test_56_a_label_that_differs_only_in_whitespace_or_case_is_refused() -> None:
    labels, ladder = _genuine_ladder()
    for damaged in ("P50 ", " P50", "p50"):
        forged = list(labels)
        forged[labels.index("P50")] = damaged
        _refused(forged, ladder, "P50")
        _refused(forged, ladder, damaged)


def test_57_a_non_finite_value_at_the_selected_rung_is_refused() -> None:
    labels, ladder = _genuine_ladder()
    for poison in (float("nan"), float("inf"), float("-inf")):
        values = list(ladder)
        values[labels.index("P90")] = poison
        detail = _refused(labels, values, "P90")
        assert "finite Double" in detail, detail


def test_58_a_non_finite_value_at_another_rung_refuses_the_whole_ladder() -> None:
    labels, ladder = _genuine_ladder()
    values = list(ladder)
    values[labels.index("P95")] = float("inf")
    detail = _refused(labels, values, "P50")
    assert "finite Double" in detail, detail
    # NOT CLAMPED: no substitute value was invented for the poisoned rung.
    assert "0" != detail


def test_59_a_short_label_carrier_is_refused_without_a_subscript_error() -> None:
    labels, ladder = _genuine_ladder()
    detail = _refused(labels[:-1], ladder, "P50", count=_const("SIM_QUANTILE_COUNT"))
    assert "label carrier" in detail, detail
    # And a caller who also shortens the claimed count is refused on the length.
    detail = _refused(labels[:-1], ladder[:-1], "P50")
    assert "accepted length" in detail, detail


def test_60_a_short_value_carrier_is_refused() -> None:
    labels, ladder = _genuine_ladder()
    detail = _refused(labels, ladder[:-1], "P50", count=_const("SIM_QUANTILE_COUNT"))
    assert "value carrier" in detail, detail


def test_61_either_carrier_longer_than_the_ladder_is_refused() -> None:
    labels, ladder = _genuine_ladder()
    _refused(labels + ["P99"], ladder, "P50", count=_const("SIM_QUANTILE_COUNT"))
    _refused(labels, ladder + [1.0], "P50", count=_const("SIM_QUANTILE_COUNT"))
    # An empty carrier - the nearest thing this engine can model to a VBA array
    # that was never sized - refuses the same way rather than reading a bound.
    _refused([], [], "P50", count=_const("SIM_QUANTILE_COUNT"))


def test_62_the_structural_validation_sorts_nothing_and_computes_no_quantile() -> None:
    vba = _transcribe()
    watched = ("SimStatsSortAscending", "SimStatsSortedCopy", "SimStatsQuantileSorted",
               "SimStatsQuantileType7", "SimStatsMean",
               "SimStatsSampleStandardDeviation", "SimStatsDescribe")
    calls = {name: 0 for name in watched}
    real = {name: vba[name] for name in watched}

    def counted(name):
        def bound(*args):
            calls[name] += 1
            return real[name](*args)
        return bound

    labels, ladder = _genuine_ladder()
    for name in watched:
        vba[name] = counted(name)
    try:
        ok, value, detail = _selected(labels, ladder, "P80")
        assert ok, detail
        _refused(labels, ladder, _const("SIM_QUANTILE_FIXED_1"))
        forged = list(labels)
        forged[3] = "P42"
        _refused(forged, ladder, "P50")
    finally:
        for name in watched:
            vba[name] = real[name]
    assert calls == {name: 0 for name in watched}, calls
    # ...and the source says so: selection reaches no numerical machinery.
    for body in (_procedure("SimStatsSelectedQuantile"),
                 _procedure("SimStatsValidateLadder")):
        for banned in ("SimStatsSortAscending", "SimStatsSortedCopy",
                       "SimStatsQuantileSorted", "SimStatsQuantileType7",
                       "SimStatsMean", "SimStatsSampleStandardDeviation",
                       "SimStatsDescribe", "SimStatsUnitScale"):
            assert banned not in body, banned


def test_63_the_boundary_is_structural_and_the_module_claims_nothing_more() -> None:
    """A finite value edited from 100 to 101 after Describe is NOT detectable
    without a seal or a second quantile calculation. Step 9 takes neither, and
    the module must not pretend otherwise."""
    labels, ladder = _genuine_ladder()
    edited = list(ladder)
    edited[labels.index("P50")] = ladder[labels.index("P50")] + 1.0
    ok, value, detail = _selected(labels, edited, "P50")
    assert ok, detail
    assert value == edited[labels.index("P50")], "the module recomputed the rung"
    code = _code()
    for sealed in ("Checksum", "Digest", "Hash", "Seal", "Signature"):
        assert sealed not in code, sealed
