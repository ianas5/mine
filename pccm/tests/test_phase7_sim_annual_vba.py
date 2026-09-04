#!/usr/bin/env python3
"""P7-5 parity: the VBA order-statistic position against the Python reference.

`modSimStats.SimStatsQuantilePosition` is executed through the accepted Phase-6
source transcriber - the real `.bas` text, not a re-implementation - and every
answer is compared against `sim_annual.percentile_position`, which the P7-5
oracle baseline was established on.

WHAT THESE PROVE THAT A PYTHON-ONLY TEST CANNOT. The tie-break is contracted as
"lower original iteration index wins", and in the VBA it is not written anywhere:
it is a PROPERTY of the accepted stable merge, which keeps equal values in
arrival order. A test that only exercised the Python would say nothing about
whether the VBA merge is stable in the same way, and a merge that quietly became
unstable would change which iteration owns a position while every percentile
value stayed identical.
"""

from __future__ import annotations

import random
import re
import sys
from pathlib import Path

import pytest

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

from pccm_builder import (  # noqa: E402
    load_contract,
    load_sim_contract,
)
from pccm_builder import sim_annual as annual  # noqa: E402
from pccm_builder.calc_numeric import (  # noqa: E402
    safe_divide,
    safe_multiply,
    safe_signed_sum,
    safe_subtract,
)
from pccm_builder.sim_emit import render_sim_contract_module  # noqa: E402
from pccm_builder.sim_stats import percentile_type7  # noqa: E402
from pccm_builder.spec_loader import load_spec  # noqa: E402
from phase6_vba_transcribe import _Ref, build as _build_transcription  # noqa: E402

SPEC = PCCM_ROOT / "spec"
SRC_VBA = PCCM_ROOT / "src" / "vba"
SIM_STATS_BAS = SRC_VBA / "modSimStats.bas"
CALC_FACTORS_BAS = SRC_VBA / "modCalcFactors.bas"

_CACHE: dict = {}

# The procedures this suite executes, and the ones it only needs signatures for.
POSITION_PROCEDURES = {
    "SimStatsQuantilePosition",
    "SimStatsPositionOf",
    "SimStatsOrderedIndices",
    "SimStatsSortIndices",
    "SimStatsUsableSequence",
    "SimStatsUsableProbability",
}


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
                # THE ACCEPTED PARSER, not a second one. A projected constant
                # may carry a trailing comment ("4294967087.0    ' exceeds
                # Long"), and a parser that split literals differently from
                # test_phase6_sim_stats_vba.py would disagree with it about what
                # a constant is.
                literal = rest.split("    '")[0].rstrip()
                out[name] = (literal[1:-1] if kind == "String"
                             else (float(literal) if kind == "Double" else int(literal)))
        _CACHE["consts"] = out
    return _CACHE["consts"]


def _scalar_shim(function):
    def shim(a, b, result, where="shim"):
        try:
            result.v = function(_value(a), _value(b))
        except Exception:
            return False
        return True
    return shim


def _value(x):
    return x.v if isinstance(x, _Ref) else x


def _sequence_shim(function):
    def shim(values, count, result, *rest):
        try:
            result.v = function(list(values)[: int(_value(count))], "shim")
        except Exception:
            return False
        return True
    return shim


def _transcribe() -> dict:
    if "vba" not in _CACHE:
        _CACHE["vba"] = _build_transcription(
            {"modSimStats": SIM_STATS_BAS, "modCalcFactors": CALC_FACTORS_BAS},
            _constants(),
            only={"modCalcFactors": {"IsUsableDouble"},
                  "modSimStats": POSITION_PROCEDURES},
            signature_only={"modCalcFactors": {
                "SafeSignedSum", "SafeDivide", "SafeMultiply", "SafeSubtract"}},
            extra={
                "MAX_DOUBLE": sys.float_info.max,
                "SafeSignedSum": _sequence_shim(safe_signed_sum),
                "SafeDivide": _scalar_shim(safe_divide),
                "SafeMultiply": _scalar_shim(safe_multiply),
                "SafeSubtract": _scalar_shim(safe_subtract),
            })
    return _CACHE["vba"]


def _position(values, p):
    """The VBA answer. The transcriber models a UDT as a dict, so the
    fields are read by name rather than by attribute."""
    transcription = _transcribe()
    position = transcription["_new"]("SimStatsPosition")
    detail = _Ref("")
    ok = transcription["SimStatsQuantilePosition"](
        list(values), _Ref(len(values)), _Ref(p), position, detail)
    return ok, position, detail.v


# ---------------------------------------------------------------------------
# THE SAMPLES
# ---------------------------------------------------------------------------
# Deliberately including: strict ties, all-equal sequences, a single element,
# negative values, and a run long enough that the bottom-up merge performs more
# than one pass (a merge bug that only appears above one run length is exactly
# what a four-element sample would miss).
def _samples():
    rng = random.Random(1789)
    return {
        "single": [42.0],
        "pair": [2.0, 1.0],
        "distinct": [10.0, 30.0, 20.0, 40.0],
        "all_equal": [7.0] * 9,
        "ties": [50.0, 10.0, 50.0, 10.0, 30.0],
        "many_ties": [1.0, 2.0, 1.0, 2.0, 1.0, 2.0, 3.0, 1.0, 3.0, 2.0, 3.0],
        "negatives": [-5.0, 12.0, -5.0, 0.0, -100.0, 12.0],
        "long": [rng.uniform(-1e6, 1e6) for _ in range(257)],
        "long_with_ties": [float(rng.randint(0, 12)) for _ in range(200)],
    }


PROBABILITIES = [0.0, 0.01, 0.1, 0.25, 1.0 / 3.0, 0.5, 0.7, 0.9, 0.99, 1.0]


@pytest.mark.parametrize("name", sorted(_samples()))
@pytest.mark.parametrize("p", PROBABILITIES)
def test_01_the_vba_position_equals_the_reference(name, p) -> None:
    """Same lo, same hi, same f, same values - on every sample and every p."""
    values = _samples()[name]
    ok, position, detail = _position(values, p)
    assert ok, f"{name} p={p}: the VBA refused: {detail}"
    reference = annual.percentile_position(values, p)
    assert position["LoSource"] == reference.lo, f"{name} p={p}: lo"
    assert position["HiSource"] == reference.hi, f"{name} p={p}: hi"
    assert position["Fraction"] == reference.fraction, f"{name} p={p}: f"
    assert position["LoValue"] == reference.lo_value, f"{name} p={p}: lo value"
    assert position["HiValue"] == reference.hi_value, f"{name} p={p}: hi value"


@pytest.mark.parametrize("name", sorted(_samples()))
@pytest.mark.parametrize("p", PROBABILITIES)
def test_02_the_position_reproduces_the_published_percentile(name, p) -> None:
    """The convex blend of the two positioned values IS the percentile.

    This is what makes the exposure safe: it does not recompute the number, and
    the number it points at is the one `SimStatsQuantileType7` already
    published.
    """
    values = _samples()[name]
    ok, position, detail = _position(values, p)
    assert ok, detail
    f = position["Fraction"]
    if f == 0.0 or position["LoValue"] == position["HiValue"]:
        rebuilt = position["LoValue"]
    else:
        rebuilt = (1.0 - f) * position["LoValue"] + f * position["HiValue"]
    assert rebuilt == percentile_type7(values, p), f"{name} p={p}"


def test_03_equal_values_resolve_to_the_lower_original_iteration() -> None:
    """The contracted tie-break, observed in the VBA rather than in the spec.

    Every value in `all_equal` is identical, so every order statistic is a tie
    and the position must walk the ORIGINAL indices in ascending order. A merge
    that lost stability would still return the right VALUE here and the wrong
    ITERATION, which is precisely the failure this exists to catch.
    """
    values = [7.0] * 9
    seen = []
    for index in range(len(values)):
        p = index / (len(values) - 1)
        ok, position, detail = _position(values, p)
        assert ok, detail
        seen.append(position["LoSource"])
    assert seen == list(range(len(values))), (
        f"the tie-break did not follow original iteration order: {seen}"
    )


def test_04_the_tie_break_moves_no_percentile_value() -> None:
    """Proved, not asserted: the two candidates hold the same number.

    For every tie in every sample, the value at the chosen index equals the
    value at every other index holding that same value - so no choice among
    them could have produced a different percentile.
    """
    for name, values in _samples().items():
        for p in PROBABILITIES:
            ok, position, detail = _position(values, p)
            assert ok, detail
            published = percentile_type7(values, p)
            # The percentile computed from the sorted copy is unchanged by the
            # identity choice, because the identity choice ranges only over
            # equal values.
            assert percentile_type7(sorted(values), p) == published, f"{name} p={p}"
            assert values[position["LoSource"]] == position["LoValue"]
            assert values[position["HiSource"]] == position["HiValue"]


def test_05_the_caller_sequence_is_never_reordered() -> None:
    """The retained iteration arrays keep their original order - the result
    digest depends on it - so the position sorts a PERMUTATION, not the data."""
    for name, values in _samples().items():
        original = list(values)
        for p in (0.0, 0.5, 1.0):
            ok, _, detail = _position(values, p)
            assert ok, detail
        assert values == original, f"{name}: the VBA reordered its caller's array"


def test_06_an_unusable_sequence_or_probability_is_refused() -> None:
    ok, _, detail = _position([], 0.5)
    assert not ok and "empty" in detail
    for bad in (-0.01, 1.01):
        ok, _, detail = _position([1.0, 2.0], bad)
        assert not ok, f"p={bad} was accepted"
        assert detail, "a refusal must say why"


def test_07_the_merge_is_exercised_beyond_one_run_length() -> None:
    """A bottom-up merge with a defect in its second or later pass looks
    correct on four elements. These samples force at least eight passes."""
    values = _samples()["long"]
    assert len(values) > 128, "the long sample must exceed several merge passes"
    ok, position, detail = _position(values, 0.5)
    assert ok, detail
    reference = annual.percentile_position(values, 0.5)
    assert (position["LoSource"], position["HiSource"]) == (reference.lo, reference.hi)


def test_08_the_position_carries_no_percentile_arithmetic() -> None:
    """It reads a position; it does not interpolate one.

    Structural, over the procedure's own text: the convex blend belongs to
    `SimStatsQuantileSorted`, and a second copy of it here would be a second
    percentile implementation able to drift from the published one.
    """
    text = SIM_STATS_BAS.read_text(encoding="utf-8")
    start = text.index("Public Function SimStatsQuantilePosition")
    end = text.index("End Function", start)
    body = text[start:end]
    for forbidden in ("1# - ", "(1 -", "* low", "* high", "candidate"):
        assert forbidden not in body, (
            f"SimStatsQuantilePosition contains {forbidden!r}; it must expose a "
            "position, not compute a value"
        )
    assert "SimStatsPositionOf(count, p, lowIndex, highIndex)" in body, (
        "the position must come from the shared owner, so it cannot differ from "
        "the arithmetic the published value used"
    )
    # And the owner is genuinely shared: the quantile takes its position from
    # the same call, so there is exactly one type-7 position in the module.
    assert text.count("h = CDbl(count - 1) * p") == 1, (
        "the type-7 position arithmetic is written more than once"
    )
    assert text.count("SimStatsPositionOf(count, p, lowIndex, highIndex)") == 2, (
        "the shared owner must have both callers - the value and the position"
    )
