#!/usr/bin/env python3
"""P7-6 correction: a whole ladder of quantiles, sorting exactly once.

THE DEFECT THIS REMOVES. `SimAnnualLadder` took a year's eleven contracted rungs
by calling `SimStatsQuantileType7` eleven times, and that entry point sorts a
fresh copy of the series on every call. The annual distributions therefore cost
11 sorts of N per project year per measure - 22 * Y per invocation - where the
eleven rungs are taken over ONE series and one ordered copy serves all of them.

WHAT MUST NOT CHANGE, AND IS PROVED HERE RATHER THAN ASSERTED. Every value. The
batch service evaluates the SAME accepted rule - `SimStatsQuantileSorted`, whose
position is `SimStatsPositionOf`'s - against one ordered copy instead of eleven
identical ones, so parity with the standalone entry point is EXACT and is
required exactly, not to a tolerance.

NOTHING HERE IS RUNTIME EVIDENCE. No Windows run has executed this code.
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

from phase6_vba_transcribe import _Ref  # noqa: E402

# THE ONE TRANSCRIPTION. The persistence suite already compiles the real
# `.bas` text of modSimStats, modSimAnnual and the engine together; building a
# second transcription here would be a second statement about what the source
# is, free to drift from the first.
import test_phase7_annual_persistence as annual  # noqa: E402

SRC_VBA = PCCM_ROOT / "src" / "vba"
STATS_BAS = SRC_VBA / "modSimStats.bas"
ANNUAL_BAS = SRC_VBA / "modSimAnnual.bas"


def _vba():
    return annual._vba()


# ---------------------------------------------------------------------------
# THE SAMPLES
# ---------------------------------------------------------------------------
# Odd and even lengths, negatives, heavy ties, an all-equal series, a single
# element, and a run long enough that the bottom-up merge takes several passes -
# a batch that reused an ordered copy incorrectly would still look right on four
# elements.
def _samples() -> dict[str, list[float]]:
    rng = random.Random(20260905)
    return {
        "single": [42.0],
        "pair": [2.0, 1.0],
        "odd": [10.0, 30.0, 20.0, 40.0, 5.0],
        "even": [10.0, 30.0, 20.0, 40.0],
        "all_equal": [7.0] * 9,
        "heavy_ties": [float(rng.randint(0, 3)) for _ in range(101)],
        "negatives": [-5.0, 12.0, -5.0, 0.0, -100.0, 12.0],
        "mixed_magnitudes": [1e-9, -1e9, 0.0, 1e9, -1e-9, 3.5],
        "long": [rng.uniform(-1e6, 1e6) for _ in range(257)],
        "long_even": [rng.uniform(-1e3, 1e3) for _ in range(256)],
    }


LADDER = [0.10, 0.50, 0.55, 0.60, 0.65, 0.70, 0.75, 0.80, 0.85, 0.90, 0.95]
"""The eleven contracted rungs. Taken as literals here on purpose: this suite is
about the batch path returning what the accepted path returns, so it must not
read the ladder through the same projection the code under test reads it from."""


# THE CALLER'S OWN LIST GOES IN, NOT A COPY OF IT. The transcriber models a VBA
# array as the Python list it is handed, so passing `list(values)` would put a
# throwaway in front of the source and make an in-place sort invisible - which
# is exactly the defect test_07 exists to catch, and did not until this changed.
def _batch(values, probabilities):
    results, detail = [], _Ref("")
    ok = _vba()["SimStatsQuantileLadder"](
        values, _Ref(len(values)), list(probabilities),
        _Ref(len(probabilities)), results, detail)
    return ok, [float(v) for v in results], detail.v


def _single(values, p):
    result, detail = _Ref(0.0), _Ref("")
    ok = _vba()["SimStatsQuantileType7"](
        values, _Ref(len(values)), _Ref(p), result, detail)
    return ok, result.v, detail.v


# ===========================================================================
# A. VALUE PARITY
# ===========================================================================
@pytest.mark.parametrize("name", sorted(_samples()))
def test_01_the_batch_ladder_equals_eleven_independent_type7_calls(name) -> None:
    """EXACT equality, rung by rung. Not a tolerance: the batch path runs the
    same arithmetic on the same ordered values, so any difference at all is a
    different implementation rather than a rounding artefact."""
    values = _samples()[name]
    ok, batch, detail = _batch(values, LADDER)
    assert ok, f"{name}: the batch refused: {detail}"
    assert len(batch) == len(LADDER)
    for index, p in enumerate(LADDER):
        accepted_ok, accepted, accepted_detail = _single(values, p)
        assert accepted_ok, f"{name} p={p}: {accepted_detail}"
        assert batch[index] == accepted, f"{name} p={p}"


@pytest.mark.parametrize("name", sorted(_samples()))
def test_02_parity_holds_at_the_endpoints_and_off_ladder_probabilities(name) -> None:
    """The rungs the contract does not name are the same arithmetic, so a
    difference there would be a difference everywhere."""
    values = _samples()[name]
    probabilities = [0.0, 0.01, 1.0 / 3.0, 0.5, 2.0 / 3.0, 0.99, 1.0]
    ok, batch, detail = _batch(values, probabilities)
    assert ok, detail
    for index, p in enumerate(probabilities):
        assert batch[index] == _single(values, p)[1], f"{name} p={p}"


@pytest.mark.parametrize("name", sorted(_samples()))
def test_03_a_repeated_probability_returns_the_same_value_twice(name) -> None:
    """A ladder may name a rung more than once, and the second must not be
    served from a different ordered copy or a stale local."""
    values = _samples()[name]
    ok, batch, detail = _batch(values, [0.5, 0.9, 0.5, 0.1, 0.9])
    assert ok, detail
    assert batch[0] == batch[2] == _single(values, 0.5)[1], name
    assert batch[1] == batch[4] == _single(values, 0.9)[1], name


def test_04_the_order_of_the_requested_rungs_is_the_order_returned() -> None:
    """Descending, so a batch that quietly sorted its own probabilities - or
    reused a position from the previous rung - is visible."""
    values = _samples()["long"]
    descending = list(reversed(LADDER))
    ok, batch, detail = _batch(values, descending)
    assert ok, detail
    for index, p in enumerate(descending):
        assert batch[index] == _single(values, p)[1], p
    ascending = _batch(values, LADDER)[1]
    assert batch == list(reversed(ascending))


def test_05_an_empty_series_an_empty_ladder_and_a_bad_probability_are_refused() -> None:
    ok, _, detail = _batch([], LADDER)
    assert not ok and detail
    ok, _, detail = _batch([1.0, 2.0], [])
    assert not ok and "no rungs" in detail
    for bad in (-0.01, 1.01):
        ok, _, detail = _batch([1.0, 2.0], [0.5, bad])
        assert not ok, f"p={bad} was accepted"
        assert detail, "a refusal must say why"


def test_06_a_refused_rung_publishes_no_partial_ladder() -> None:
    """Transactional, like SimStatsDescribe: the ladder is committed only when
    every rung succeeded."""
    results, detail = [], _Ref("")
    ok = _vba()["SimStatsQuantileLadder"](
        [1.0, 2.0, 3.0], _Ref(3), [0.1, 0.5, 9.0], _Ref(3), results, detail)
    assert not ok
    assert results == [], f"a partial ladder was handed back: {results}"


# ===========================================================================
# B. SOURCE IMMUTABILITY
# ===========================================================================
@pytest.mark.parametrize("name", sorted(_samples()))
def test_07_the_callers_series_is_not_reordered_or_rewritten(name) -> None:
    """The retained iteration arrays keep their original order - the result
    digest depends on it - so the batch sorts a COPY."""
    values = _samples()[name]
    original = list(values)
    ok, _, detail = _batch(values, LADDER)
    assert ok, detail
    assert values == original, f"{name}: the batch mutated its caller's series"
    # And the standalone entry point, on the same terms - it is unchanged and
    # must stay that way.
    assert _single(values, 0.5)[0]
    assert values == original, f"{name}: the standalone entry point mutated it"


def test_08_the_batch_sorts_a_copy_and_says_so_in_the_source() -> None:
    body = _procedure("SimStatsQuantileLadder")
    assert "SimStatsSortedCopy(values, count, ordered, detail)" in body, (
        "the batch does not take an ordered COPY")
    assert re.search(r"^\s*values\(", body, re.M) is None, (
        "the batch assigns into the caller's series")


# ===========================================================================
# C. THE SORT-COUNT PROPERTY
# ===========================================================================
def _procedure(name: str, path: Path = STATS_BAS) -> str:
    text = path.read_text(encoding="utf-8")
    start = text.index(f"Function {name}(")
    return text[start:text.index("\nEnd Function", start)]


def _count_sorts(call) -> int:
    """How many ordered copies one call actually makes.

    COUNTED, NOT INFERRED FROM A NAME. A path that stopped calling
    SimStatsQuantileType7 but sorted eleven times some other way would satisfy a
    token check and fail this.
    """
    transcription = _vba()
    original = transcription["SimStatsSortedCopy"]
    seen = {"n": 0}

    def counting(*args):
        seen["n"] += 1
        return original(*args)

    transcription["SimStatsSortedCopy"] = counting
    try:
        call()
    finally:
        transcription["SimStatsSortedCopy"] = original
    return seen["n"]


def test_09_the_batch_makes_exactly_one_ordered_copy() -> None:
    values = _samples()["long"]
    assert _count_sorts(lambda: _batch(values, LADDER)) == 1
    # And the standalone entry point still makes one per call, unchanged.
    assert _count_sorts(lambda: _single(values, 0.5)) == 1
    assert _count_sorts(
        lambda: [_single(values, p) for p in LADDER]) == len(LADDER)


@pytest.mark.parametrize("years", [1, 4, 7])
def test_10_the_annual_ladder_costs_one_sort_per_year(years) -> None:
    """THE MATERIAL PROPERTY, measured on the real annual path.

    One ordered copy per year per measure. The count is taken by instrumenting
    the sort itself, so it holds however the rungs are reached.
    """
    iterations = 40
    rng = random.Random(7)
    column = [rng.uniform(0.0, 1000.0) for _ in range(iterations * years)]

    def run():
        ladder, detail = [], _Ref("")
        ok = _vba()["SimAnnualLadder"](
            list(column), _Ref(iterations), _Ref(years), list(LADDER),
            _Ref(len(LADDER)), ladder, detail)
        assert ok, detail.v
        return ladder

    assert _count_sorts(run) == years, (
        f"{years} project year(s) cost more than one ordered copy each")


def test_11_the_annual_ladder_reaches_no_per_rung_sorting_entry_point() -> None:
    """The structural half, and it is deliberately not the whole claim: test_10
    is what proves the cost, because a second sorting path under another name
    would pass this and fail that."""
    body = _procedure("SimAnnualLadder", ANNUAL_BAS)
    assert "SimStatsQuantileType7" not in body, (
        "the annual ladder still calls the per-rung sorting entry point")
    assert "SimStatsQuantileLadder" in body
    assert body.count("modSimStats.") == 1, (
        "the annual ladder reaches statistics more than once per year")
    # And modSimAnnual sorts nothing of its own.
    text = ANNUAL_BAS.read_text(encoding="utf-8")
    code = "\n".join(line for line in text.splitlines()
                     if not line.strip().startswith("'"))
    for banned in ("Sort", "ordered", "Swap", "merge"):
        assert banned not in code, f"modSimAnnual carries {banned!r}"


# ===========================================================================
# D. TYPE-7 REMAINS SINGLE-OWNED
# ===========================================================================
def test_12_the_position_and_the_interpolation_have_one_owner_each() -> None:
    text = STATS_BAS.read_text(encoding="utf-8")
    assert text.count("h = CDbl(count - 1) * p") == 1, (
        "the type-7 position arithmetic is written more than once")
    # The batch computes no position and no blend of its own.
    body = _procedure("SimStatsQuantileLadder")
    for forbidden in ("1# - ", "SimStatsPositionOf", "* low", "* high", "CDbl(count"):
        assert forbidden not in body, (
            f"SimStatsQuantileLadder contains {forbidden!r}; it must evaluate the "
            "accepted rule, not restate it")
    assert "SimStatsQuantileSorted(ordered, count," in body, (
        "the batch does not go through the accepted value owner")
    # BOTH ENTRY POINTS GO THROUGH THE SAME ONE.
    assert "SimStatsQuantileSorted(ordered, count," in _procedure("SimStatsQuantileType7")


def test_13_the_standalone_entry_point_is_unchanged_in_substance() -> None:
    """Existing callers must remain valid: same signature, same guards, same
    single-sort body."""
    body = _procedure("SimStatsQuantileType7")
    assert "ByRef values() As Double, ByVal count As Long" in body
    assert "ByVal p As Double, ByRef result As Double" in body
    assert "SimStatsSortedCopy(values, count, ordered, detail)" in body
    assert body.count("SimStatsSortedCopy") == 1
    assert "SimStatsQuantileLadder" not in body, (
        "the standalone entry point now depends on the batch; a change to the "
        "batch would then move accepted behaviour")


def test_14_equal_values_keep_the_accepted_constant_bracket_result() -> None:
    """An all-equal series reports its own value at every rung, with no creep -
    the invariant SimStatsQuantileSorted owns, reached through the batch."""
    ok, batch, detail = _batch([0.1] * 13, LADDER)
    assert ok, detail
    assert batch == [0.1] * len(LADDER)


def test_15_the_public_surface_gained_exactly_one_name() -> None:
    from pccm_builder.vba_source import VbaModule

    module = VbaModule(name="modSimStats", path=STATS_BAS,
                       raw=STATS_BAS.read_text(encoding="utf-8"))
    assert "SimStatsQuantileLadder" in module.public_procedures
    private = set(module.procedures) - set(module.public_procedures)
    for owned in ("SimStatsSortedCopy", "SimStatsSortAscending",
                  "SimStatsQuantileSorted", "SimStatsPositionOf"):
        assert owned in private, (
            f"{owned} was exposed; the sorting implementation stays inside the "
            "module that owns it")
