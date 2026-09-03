#!/usr/bin/env python3
"""PCCM Phase 7 Step-2 conformance tests for `src/vba/modSimSensitivity.bas`.

Mid-ranks, Spearman as Pearson-on-ranks, the undefined zero-variance status, and
the ordering of finished driver results.

--------------------------------------------------------------------------------
WHAT THESE TESTS PROVE, AND WHAT THEY DO NOT
--------------------------------------------------------------------------------
SOURCE CONFORMANCE, on Linux, now: purity, the public surface, the ownership
boundary, and the mathematics - through the accepted Phase-6 source transcriber,
against hand-derived vectors and an independent Python reference.

VBA EXECUTION CONFORMANCE is NOT proved and is deferred to Windows. No VBA
runtime exists here. Nothing in this file may be read as "VBA computed a rho".

THE REFERENCE IS NOT THE TRANSCRIPTION. Every numeric claim below is anchored
either to a vector derived by hand from the definition of a mid-rank, or to a
correlation computed in this file from first principles - not to the same code
under a different name. `sim_sensitivity` is the project's definition of these
semantics and is itself checked against those hand-derived answers.

Runs standalone or under pytest.
"""

from __future__ import annotations

import math
import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

from pccm_builder import sim_sensitivity as oracle  # noqa: E402
from pccm_builder.calc_numeric import (  # noqa: E402
    CalculationRefusal,
    safe_add,
    safe_divide,
    safe_multiply,
)
from pccm_builder.contract_loader import ContractError  # noqa: E402
from pccm_builder.structure_loader import load_structure_contract  # noqa: E402
from pccm_builder.vba_source import VbaModule, load_modules  # noqa: E402

from phase6_vba_transcribe import _Ref, _val, build as _build_transcription  # noqa: E402

SRC_VBA = PCCM_ROOT / "src" / "vba"
SENSITIVITY_BAS = SRC_VBA / "modSimSensitivity.bas"
CALC_FACTORS_BAS = SRC_VBA / "modCalcFactors.bas"
SPEC = PCCM_ROOT / "spec"

_CACHE: dict[str, object] = {}


# ---------------------------------------------------------------------------
# Source access and transcription
# ---------------------------------------------------------------------------
def _module() -> VbaModule:
    return VbaModule(name="modSimSensitivity", path=SENSITIVITY_BAS,
                     raw=SENSITIVITY_BAS.read_text(encoding="utf-8"))


def _code() -> str:
    return _module().code


def _procedure(name: str) -> str:
    """One procedure's source, string literals intact."""
    code = _module().code_without_string_removal
    match = re.search(
        rf"^\s*(?:Public|Private)\s+(?:Function|Sub)\s+{re.escape(name)}\b", code, re.M)
    assert match, f"{name} is not declared"
    tail = code[match.start():]
    end = re.search(r"^\s*End\s+(?:Function|Sub)\s*$", tail, re.M)
    assert end, f"{name} has no End"
    return tail[: end.end()]


def _scalar_shim(function):
    def bound(*args):
        *operands, result = args
        try:
            result.v = function(*[float(_val(v)) for v in operands], "sensitivity")
        except (CalculationRefusal, ContractError):
            return False
        return True
    return bound


def _module_constants() -> dict:
    """The module's own `Public Const` values, read out of its source.

    They are not restated here: a test that hardcoded 0 and 1 would still pass
    if the source swapped them, and `test_30` is what pins the literals.
    """
    out = {"MAX_DOUBLE": sys.float_info.max}
    for line in SENSITIVITY_BAS.read_text(encoding="utf-8").splitlines():
        match = re.match(r"^Public Const (\w+) As Long = (-?\d+)", line)
        if match:
            out[match.group(1)] = int(match.group(2))
    return out


def _transcribe() -> dict:
    if "vba" not in _CACHE:
        _CACHE["vba"] = _build_transcription(
            {"modSimSensitivity": SENSITIVITY_BAS, "modCalcFactors": CALC_FACTORS_BAS},
            _module_constants(),
            only={"modCalcFactors": {"IsUsableDouble"},
                  "modSimSensitivity": set(_module().procedures)},
            signature_only={"modCalcFactors": {
                "SafeAdd", "SafeDivide", "SafeMultiply"}},
            extra={
                "MAX_DOUBLE": sys.float_info.max,
                "SafeAdd": _scalar_shim(safe_add),
                "SafeDivide": _scalar_shim(safe_divide),
                "SafeMultiply": _scalar_shim(safe_multiply),
            })
    return _CACHE["vba"]  # type: ignore[return-value]


def _mid_ranks(values):
    ranks, detail = [], _Ref("")
    ok = _transcribe()["SimSensitivityMidRanks"](
        list(values), _Ref(len(values)), ranks, detail)
    return ok, [_val(r) for r in ranks], detail.v


def _correlate(driver_ranks, total_ranks):
    rho, status, detail = _Ref(0.0), _Ref(0), _Ref("")
    ok = _transcribe()["SimSensitivityRankCorrelation"](
        list(driver_ranks), list(total_ranks), _Ref(len(driver_ranks)),
        rho, status, detail)
    return ok, rho.v, status.v, detail.v


def _spearman(driver_values, total_ranks):
    rho, status, detail = _Ref(0.0), _Ref(0), _Ref("")
    ok = _transcribe()["SimSensitivitySpearman"](
        list(driver_values), list(total_ranks), _Ref(len(driver_values)),
        rho, status, detail)
    return ok, rho.v, status.v, detail.v


def _result(permanent_id, rho, status=oracle.SENSITIVITY_DEFINED):
    entry = _transcribe()["_new"]("SimSensitivityResult")
    entry["PermanentId"] = permanent_id
    entry["Rho"] = rho
    entry["AbsRho"] = abs(rho)
    entry["Status"] = status
    return entry


def _rank(results):
    order, eligible, detail = [], _Ref(0), _Ref("")
    ok = _transcribe()["SimSensitivityRank"](
        list(results), _Ref(len(results)), order, eligible, detail)
    return ok, [int(_val(o)) for o in order[: int(eligible.v)]], int(eligible.v), detail.v


# ---------------------------------------------------------------------------
# An INDEPENDENT correlation, written from the definition rather than reused
# ---------------------------------------------------------------------------
def _pearson_from_first_principles(x, y):
    """Textbook Pearson, deliberately not the oracle's implementation."""
    n = len(x)
    mx = sum(x) / n
    my = sum(y) / n
    sxy = sum((a - mx) * (b - my) for a, b in zip(x, y))
    sxx = sum((a - mx) ** 2 for a in x)
    syy = sum((b - my) ** 2 for b in y)
    return sxy / math.sqrt(sxx * syy)


# ===========================================================================
# A. THE HAND-DERIVED MID-RANK VECTORS
# ===========================================================================
# Every one of these is the definition applied by hand: a tie block spanning
# ordinal positions p..q takes (p + q) / 2.
HAND_DERIVED = (
    ([10.0, 20.0, 30.0], [1.0, 2.0, 3.0]),
    ([10.0, 10.0, 30.0], [1.5, 1.5, 3.0]),
    ([10.0, 20.0, 20.0, 20.0, 50.0], [1.0, 3.0, 3.0, 3.0, 5.0]),
    ([0.0, 0.0, 0.0, 1.0, 1.0], [2.0, 2.0, 2.0, 4.5, 4.5]),
)


def test_01_the_reference_reproduces_every_hand_derived_mid_rank() -> None:
    """The reference is checked before anything is checked against it."""
    for values, expected in HAND_DERIVED:
        assert oracle.mid_ranks(values) == expected, values


def test_02_the_vba_reproduces_every_hand_derived_mid_rank() -> None:
    for values, expected in HAND_DERIVED:
        ok, ranks, detail = _mid_ranks(values)
        assert ok, (values, detail)
        assert ranks == expected, (values, ranks, expected)


def test_03_a_tie_block_takes_the_midpoint_of_the_positions_it_spans() -> None:
    """Stated as the rule, not as four examples of it."""
    values = [5.0, 5.0, 5.0, 5.0, 9.0, 9.0, 1.0]
    ok, ranks, detail = _mid_ranks(values)
    assert ok, detail
    # sorted: 1 | 5 5 5 5 | 9 9   -> positions 1 | 2..5 | 6..7
    assert ranks == [3.5, 3.5, 3.5, 3.5, 6.5, 6.5, 1.0]
    assert ranks == oracle.mid_ranks(values)


def test_04_every_rank_vector_sums_to_the_triangular_number() -> None:
    """A property no example can express: mid-ranks are a permutation's worth of
    rank mass however the ties fall, so the sum is always n(n + 1) / 2."""
    for values in ([1.0], [2.0, 2.0], [3.0, 1.0, 2.0], [7.0] * 9,
                   [0.0, 0.0, 0.0, 1.0, 1.0], [1.0, 2.0, 2.0, 3.0, 3.0, 3.0]):
        ok, ranks, detail = _mid_ranks(values)
        assert ok, detail
        n = len(values)
        assert math.isclose(sum(ranks), n * (n + 1) / 2, rel_tol=0, abs_tol=1e-9), values


# ===========================================================================
# B-C. PERFECT MONOTONE ASSOCIATION
# ===========================================================================
def test_05_a_perfect_positive_monotone_relation_is_plus_one() -> None:
    total_ranks = oracle.mid_ranks([10.0, 20.0, 30.0, 40.0, 50.0])
    ok, rho, status, detail = _spearman([1.0, 2.0, 3.0, 4.0, 5.0], total_ranks)
    assert ok, detail
    assert status == oracle.SENSITIVITY_DEFINED
    assert rho == 1.0, rho


def test_06_a_perfect_negative_monotone_relation_is_minus_one() -> None:
    total_ranks = oracle.mid_ranks([50.0, 40.0, 30.0, 20.0, 10.0])
    ok, rho, status, detail = _spearman([1.0, 2.0, 3.0, 4.0, 5.0], total_ranks)
    assert ok, detail
    assert status == oracle.SENSITIVITY_DEFINED
    assert rho == -1.0, rho


def test_07_a_monotone_but_non_linear_relation_is_still_plus_one() -> None:
    """Spearman measures ORDER. A relation no straight line fits is still a
    perfect monotone association, and a Pearson taken on the raw values would
    report less than one here."""
    driver = [1.0, 2.0, 3.0, 4.0, 5.0]
    total = [1.0, 4.0, 9.0, 16.0, 25000.0]
    ok, rho, status, detail = _spearman(driver, oracle.mid_ranks(total))
    assert ok, detail
    assert rho == 1.0
    # A straight line fits this badly, and Pearson on the raw values says so.
    assert _pearson_from_first_principles(driver, total) < 0.75


# ===========================================================================
# D-E. TIES, AND THE RISK OCCURRENCE SHAPE
# ===========================================================================
def test_08_a_tied_driver_matches_an_independent_mid_rank_pearson() -> None:
    driver = [3.0, 1.0, 1.0, 7.0, 5.0, 5.0, 5.0, 2.0]
    total = [10.0, 40.0, 20.0, 80.0, 50.0, 60.0, 30.0, 70.0]
    total_ranks = oracle.mid_ranks(total)
    ok, rho, status, detail = _spearman(driver, total_ranks)
    assert ok, detail
    assert status == oracle.SENSITIVITY_DEFINED
    expected = _pearson_from_first_principles(oracle.mid_ranks(driver), total_ranks)
    assert math.isclose(rho, expected, rel_tol=1e-12, abs_tol=1e-12), (rho, expected)


def test_09_an_occurred_risk_leaves_a_large_tie_block() -> None:
    """The shape a Risk at roughly 20% occurrence actually produces: eight
    iterations contribute exactly zero and tie, two carry a severity.

    This is the case the no-ties shortcut gets WRONG, and it is the ordinary
    case rather than a corner. Nothing here samples anything - the observations
    are given, and only the mathematics after them is under test.
    """
    driver = [0.0] * 8 + [1.0, 2.0]
    total = [12.0, 31.0, 7.0, 55.0, 23.0, 44.0, 19.0, 38.0, 61.0, 90.0]
    ok, ranks, detail = _mid_ranks(driver)
    assert ok, detail
    # The zeros span ordinal positions 1..8, so every one of them ranks 4.5.
    assert ranks == [4.5] * 8 + [9.0, 10.0]
    total_ranks = oracle.mid_ranks(total)
    ok, rho, status, detail = _spearman(driver, total_ranks)
    assert ok, detail
    expected = _pearson_from_first_principles(ranks, total_ranks)
    assert math.isclose(rho, expected, rel_tol=1e-12, abs_tol=1e-12), (rho, expected)
    # AND THE FORBIDDEN SHORTCUT WOULD HAVE DISAGREED. Computed here only to
    # show the two are different numbers on this data.
    n = len(driver)
    shortcut = 1 - 6 * sum((a - b) ** 2 for a, b in zip(ranks, total_ranks)) / (n * (n * n - 1))
    assert not math.isclose(shortcut, rho, rel_tol=1e-9, abs_tol=1e-9), (shortcut, rho)


# ===========================================================================
# F-G. UNDEFINED IS NOT ZERO
# ===========================================================================
def test_10_a_constant_driver_is_undefined_and_not_zero_rho() -> None:
    total_ranks = oracle.mid_ranks([10.0, 20.0, 30.0, 40.0, 50.0])
    ok, rho, status, detail = _spearman([0.0] * 5, total_ranks)
    assert ok, detail
    assert status == oracle.SENSITIVITY_NO_VARIANCE
    assert oracle.spearman([0.0] * 5, total_ranks)[1] == oracle.SENSITIVITY_NO_VARIANCE


def test_11_a_constant_total_is_undefined_too() -> None:
    """The converse. A total that never moved is not a total every driver is
    uncorrelated with; it is a total nothing can be correlated against."""
    total_ranks = oracle.mid_ranks([7.0] * 5)
    ok, rho, status, detail = _spearman([1.0, 2.0, 3.0, 4.0, 5.0], total_ranks)
    assert ok, detail
    assert status == oracle.SENSITIVITY_NO_VARIANCE


def test_12_a_genuine_zero_correlation_is_defined_and_distinguishable() -> None:
    """THE CONTROL THAT MAKES THE STATUS WORTH HAVING.

    A driver that genuinely varies and genuinely has no monotone association
    with the total returns rho = 0 with status DEFINED. Reporting undefined as
    rho = 0 would make these two cases indistinguishable, and they are not the
    same fact.
    """
    driver = [1.0, 2.0, 3.0, 4.0]
    # Ranks 1,2,3,4 against 2,4,1,3: the centred products are
    # +0.75, -0.75, -0.75, +0.75, which cancel exactly.
    total = [20.0, 40.0, 10.0, 30.0]
    total_ranks = oracle.mid_ranks(total)
    ok, rho, status, detail = _spearman(driver, total_ranks)
    assert ok, detail
    assert status == oracle.SENSITIVITY_DEFINED
    assert rho == 0.0, rho
    # Same rho, different status, and the caller can tell them apart.
    _, undefined_rho, undefined_status, _ = _spearman([9.0] * 4, total_ranks)
    assert undefined_rho == 0.0 and undefined_status == oracle.SENSITIVITY_NO_VARIANCE
    assert status != undefined_status


# ===========================================================================
# H. RANKING
# ===========================================================================
def test_13_the_population_is_ordered_by_absolute_rho_descending() -> None:
    results = [_result("R-002", 0.10), _result("C-001", -0.90),
               _result("C-003", 0.50), _result("R-004", -0.30)]
    ok, order, eligible, detail = _rank(results)
    assert ok, detail
    assert eligible == 4
    assert [results[i]["PermanentId"] for i in order] == ["C-001", "C-003", "R-004", "R-002"]
    # SIGNED RHO SURVIVES the ordering that used its magnitude.
    assert [results[i]["Rho"] for i in order] == [-0.90, 0.50, -0.30, 0.10]


def test_14_equal_magnitude_is_broken_by_the_permanent_id() -> None:
    """Deterministic, and not by supply order: the same drivers handed over in
    the opposite order must produce the same report."""
    forward = [_result("C-010", 0.4), _result("C-002", -0.4), _result("R-001", 0.4)]
    backward = [_result("R-001", 0.4), _result("C-002", -0.4), _result("C-010", 0.4)]
    ok, order_f, _, detail = _rank(forward)
    assert ok, detail
    ok, order_b, _, detail = _rank(backward)
    assert ok, detail
    assert [forward[i]["PermanentId"] for i in order_f] == ["C-002", "C-010", "R-001"]
    assert [backward[i]["PermanentId"] for i in order_b] == ["C-002", "C-010", "R-001"]


def test_15_the_tie_break_is_ordinal_and_not_a_locale_text_compare() -> None:
    """`Z-001` precedes `a-001` by code unit and follows it in most text
    collations. An ordering that changes with the host is not an ordering."""
    results = [_result("a-001", 0.5), _result("Z-001", 0.5)]
    ok, order, _, detail = _rank(results)
    assert ok, detail
    assert [results[i]["PermanentId"] for i in order] == ["Z-001", "a-001"]
    assert oracle.rank_drivers([("a-001", 0.5, 0), ("Z-001", 0.5, 0)]) == [1, 0]


def test_16_a_zero_variance_driver_is_not_in_the_ranked_population() -> None:
    results = [_result("C-001", 0.8),
               _result("C-002", 0.0, oracle.SENSITIVITY_NO_VARIANCE),
               _result("R-001", -0.2)]
    ok, order, eligible, detail = _rank(results)
    assert ok, detail
    assert eligible == 2
    assert [results[i]["PermanentId"] for i in order] == ["C-001", "R-001"]
    # It is EXCLUDED, not deleted: the caller still holds it and can report it.
    assert results[1]["Status"] == oracle.SENSITIVITY_NO_VARIANCE


def test_17_nothing_is_truncated() -> None:
    """No Top-N. Choosing how many bars to draw is a chart decision."""
    results = [_result(f"C-{i:03d}", (100 - i) / 100.0) for i in range(40)]
    ok, order, eligible, detail = _rank(results)
    assert ok, detail
    assert eligible == 40
    assert len(order) == 40
    assert len(set(order)) == 40


def test_18_a_population_with_no_eligible_driver_is_empty_not_an_error() -> None:
    results = [_result("C-001", 0.0, oracle.SENSITIVITY_NO_VARIANCE)]
    ok, order, eligible, detail = _rank(results)
    assert ok, detail
    assert eligible == 0 and order == []


# ===========================================================================
# I. SOURCE IMMUTABILITY AND ITERATION ALIGNMENT
# ===========================================================================
def test_19_ranking_a_sequence_does_not_reorder_it() -> None:
    values = [30.0, 10.0, 20.0, 10.0]
    before = list(values)
    ok, ranks, detail = _mid_ranks(values)
    assert ok, detail
    assert values == before, "the caller's sequence was reordered"
    # AND OBSERVATION j KEPT ITS PLACE. The two 10.0s are at 1 and 3, and their
    # shared rank appears at 1 and 3.
    assert ranks == [4.0, 1.5, 3.0, 1.5]


def test_20_the_result_sequence_is_not_reordered_by_ranking() -> None:
    results = [_result("C-003", 0.1), _result("C-001", 0.9), _result("C-002", 0.5)]
    before = [r["PermanentId"] for r in results]
    ok, order, _, detail = _rank(results)
    assert ok, detail
    assert [r["PermanentId"] for r in results] == before, "the caller's results were reordered"
    assert order == [1, 2, 0], "the order is returned as indices into the caller's sequence"


def test_21_iteration_alignment_survives_a_shuffled_but_paired_input() -> None:
    """Pairing is by POSITION, and permuting both series the same way cannot
    change the correlation. Permuting only one must."""
    driver = [4.0, 1.0, 3.0, 2.0, 5.0]
    total = [40.0, 10.0, 30.0, 20.0, 50.0]
    ok, rho, _, detail = _spearman(driver, oracle.mid_ranks(total))
    assert ok, detail
    assert rho == 1.0
    permutation = [2, 0, 4, 1, 3]
    ok, same, _, detail = _spearman([driver[i] for i in permutation],
                                    oracle.mid_ranks([total[i] for i in permutation]))
    assert ok, detail
    assert same == rho
    ok, broken, _, detail = _spearman([driver[i] for i in permutation],
                                      oracle.mid_ranks(total))
    assert ok, detail
    assert broken != rho, "pairing survived a permutation of one series only"


# ===========================================================================
# J. THE TOTAL RANKS ARE COMPUTED ONCE
# ===========================================================================
def test_22_one_total_rank_vector_serves_every_driver_unchanged() -> None:
    """The reuse interface. Ranking the total again per driver would be D sorts
    of the same vector for the same answer, and the contract says once."""
    total = [50.0, 10.0, 40.0, 20.0, 30.0]
    total_ranks = oracle.mid_ranks(total)
    guard = list(total_ranks)
    drivers = {
        "C-001": [5.0, 1.0, 4.0, 2.0, 3.0],
        "C-002": [1.0, 5.0, 2.0, 4.0, 3.0],
        "R-001": [0.0, 0.0, 0.0, 7.0, 7.0],
    }
    seen = {}
    for name, values in drivers.items():
        ok, rho, status, detail = _spearman(values, total_ranks)
        assert ok, detail
        seen[name] = (rho, status)
        assert total_ranks == guard, f"{name} mutated the shared total ranks"
    assert seen["C-001"][0] == 1.0
    assert seen["C-002"][0] == -1.0
    assert seen["R-001"][1] == oracle.SENSITIVITY_DEFINED
    # And the shared vector is still usable afterwards.
    assert total_ranks == guard


def test_23_correlating_prebuilt_ranks_agrees_with_the_convenience_entry() -> None:
    driver = [3.0, 1.0, 1.0, 7.0]
    total_ranks = oracle.mid_ranks([10.0, 40.0, 20.0, 80.0])
    ok, direct, status_a, detail = _correlate(oracle.mid_ranks(driver), total_ranks)
    assert ok, detail
    ok, viaspearman, status_b, detail = _spearman(driver, total_ranks)
    assert ok, detail
    assert direct == viaspearman and status_a == status_b


# ===========================================================================
# K. REFUSALS
# ===========================================================================
def test_24_a_non_finite_observation_is_refused_rather_than_ranked() -> None:
    for bad in (float("nan"), float("inf"), float("-inf")):
        ok, _, detail = _mid_ranks([1.0, bad, 3.0])
        assert not ok, bad
        assert "finite Double" in detail, detail


def test_25_mismatched_lengths_and_short_series_are_refused() -> None:
    ok, _, _, detail = _correlate([1.0], [1.0])
    assert not ok and "at least two" in detail, detail
    ok, _, detail = _mid_ranks([])
    assert not ok, "an empty sequence was ranked"


# ===========================================================================
# L. OWNERSHIP - the module knows nothing it should not
# ===========================================================================
FORBIDDEN_OWNERSHIP = (
    "Worksheet", "Range(", "Cells(", "ListObject", "ThisWorkbook", "Application.",
    "_SimData", "MsgBox", "Announce", "run_id", "RunId", "EffectiveSeed",
    "AutoNonce", "SimRng", "SimSample", "SimEngine", "Sheets(", "Names(",
)


def test_26_the_module_reaches_nothing_outside_its_own_mathematics() -> None:
    code = _code()
    for token in FORBIDDEN_OWNERSHIP:
        assert token not in code, f"modSimSensitivity reaches {token!r}"


def test_27_no_no_ties_spearman_shortcut_appears_in_the_source() -> None:
    """Not as an optimisation, not as a fast path, not at all.

    STRUCTURAL, NOT A TOKEN SEARCH. Looking for `1 - 6` or `n * (n * n - 1)`
    catches one spelling and misses `1# - 6#` over `CDbl(count)`, which is the
    same wrong number. The shortcut has a PROPERTY instead: it is built from the
    per-observation DIFFERENCE of the two rank series, sum(d^2), and a centred
    Pearson never forms that difference at all. So the rule is that the two
    series are never differenced, and that rho leaves this procedure by exactly
    one route.
    """
    body = VbaModule(name="probe", path=SENSITIVITY_BAS,
                     raw=_procedure("SimSensitivityRankCorrelation")).code
    flat = " ".join(body.split())
    assert not re.search(r"driverRanks\s*\([^)]*\)\s*-\s*totalRanks", flat), (
        "the two rank series are differenced, which is the sum(d^2) shortcut")
    assert not re.search(r"totalRanks\s*\([^)]*\)\s*-\s*driverRanks", flat), (
        "the two rank series are differenced, which is the sum(d^2) shortcut")
    # ONE ROUTE OUT. `rho` is set to zero on entry and then only from the
    # centred quotient; any third assignment is a second way to reach an answer.
    assignments = re.findall(r"\brho\s*=\s*([^\n]+)", body)
    assert assignments == ["0#", "quotient"], assignments


def test_28_the_module_is_declared_and_within_the_ceiling() -> None:
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    declared = {m.name: m for m in structure.vba_modules}
    assert "modSimSensitivity" in declared, "the module is not registered"
    entry = declared["modSimSensitivity"]
    assert entry.generated is False
    for phrase in ("mid-ranks", "Pearson", "tie-break", "No driver replay"):
        assert phrase in entry.responsibility, phrase
    assert len(SENSITIVITY_BAS.read_text(encoding="utf-8").splitlines()) <= 1200


def test_29_the_public_surface_is_exactly_the_kernel() -> None:
    code = _code()
    public = set(re.findall(r"^Public Function (\w+)", code, re.M))
    assert public == {"SimSensitivityMidRanks", "SimSensitivityRankCorrelation",
                      "SimSensitivitySpearman", "SimSensitivityRank"}, sorted(public)
    assert "Public Sub" not in code, "a pure kernel exposes no Sub"


def test_30_the_status_constants_are_codes_and_not_presentation_strings() -> None:
    """The words a reader sees belong to the reporting layer. A mathematical
    routine that returns worksheet vocabulary has taken a decision that is not
    its to take, and P7-4 is where that decision belongs.

    READ WITH THE STRING LITERALS PRESENT. `VbaModule.code` removes them, so a
    control that used it could not see the very thing it is looking for.
    """
    executable = _module().code_without_string_removal
    for banned in ("n/a", "N/A", "no variance"):
        assert banned not in executable, f"a presentation string entered the kernel: {banned!r}"
    assert re.search(r"Public Const SIM_SENSITIVITY_DEFINED As Long = 0", executable)
    assert re.search(r"Public Const SIM_SENSITIVITY_NO_VARIANCE As Long = 1", executable)
    # STATUS IS A Long. A String constant here would be presentation by another
    # name, whatever it was called.
    assert not re.search(r"Public Const SIM_SENSITIVITY_\w+ As String", executable)


def test_31_the_correlation_bound_is_exactly_one() -> None:
    """The clamp is defensive - on well-conditioned data it never fires - so a
    behavioural test cannot see it widen. The bound is pinned at the source
    instead, which is where a widened clamp would be visible."""
    code = _code()
    assert "If quotient > 1# Then quotient = 1#" in code
    assert "If quotient < -1# Then quotient = -1#" in code
    assert not re.search(r"quotient\s*[<>]\s*-?[2-9]", code), "the clamp accepts |rho| > 1"


def test_32_the_identity_tie_break_never_compares_strings_directly() -> None:
    """VBA's `<` on String answers by Option Compare and host locale, and a
    transcription to Python cannot show the difference because Python's own
    string order happens to agree on ASCII. So the rule is enforced where it is
    visible: the comparison must be on code units, not on the strings.
    """
    body = _procedure("SimSensitivityIdPrecedes")
    assert "AscW(" in body, "the tie-break does not read code units"
    stripped = VbaModule(name="probe", path=SENSITIVITY_BAS, raw=body).code
    for shape in (r"Mid\$\([^)]*\)\s*[<>]", r"\bleft\s*[<>]\s*right",
                  r"\bright\s*[<>]\s*left", r"StrComp\("):
        assert not re.search(shape, stripped), f"a text comparison decides the order: {shape}"


def test_33_a_scenario_a_shaped_sort_agrees_at_ten_thousand_observations() -> None:
    """The scale the Windows run actually used, through the real source.

    The P7-4 Windows failure was inside this module's merge, at 10,000
    observations, and every existing vector here is small - the largest is a
    handful of values. Small vectors never reach the merge pass where the
    trailing block ends exactly at `count`, which is where the fault lived, so
    the SIZE is part of what this proves.

    Two shapes are exercised deliberately:

      descending tail   the last two observations in descending order, which is
                        what makes the first merge pass exhaust its right run
                        before its left one
      heavy ties        a Risk at 20% probability puts ~80% of a contribution
                        column on one value, so the tie blocks here are the
                        ordinary case rather than an edge one

    The comparison is against `sim_sensitivity.mid_ranks`, the definition.
    """
    total = 10000
    values = [float((index * 7919) % 4001) for index in range(total)]
    # Force the trailing pair to descend, and keep it distinct from the ties.
    values[-2] = 9_000_000.0
    values[-1] = 8_000_000.0
    assert values[-1] < values[-2]

    ok, ranks, detail = _mid_ranks(values)
    assert ok, detail
    assert ranks == oracle.mid_ranks(values), "the sort or the ranking disagrees"
    assert len(ranks) == total
    # A sanity anchor on the extremes, so a wholesale rank shift cannot pass by
    # matching a similarly shifted oracle.
    assert max(ranks) == float(total)
    assert min(ranks) >= 1.0

    ascending = sorted(values)
    assert ascending[-1] == 9_000_000.0 and ascending[-2] == 8_000_000.0

    # AND ONE DRIVER'S SPEARMAN OVER IT, which is the pairing the endpoint makes
    # for each of Scenario A's twenty drivers.
    contributions = [value * 3.0 + ((index % 5) - 2) for index, value in enumerate(values)]
    ok, rho, status, detail = _spearman(contributions, ranks)
    assert ok, detail
    expected_rho, expected_status = oracle.spearman(contributions, ranks)
    assert status == expected_status, detail
    assert abs(rho - expected_rho) <= 1e-12, (rho, expected_rho)
