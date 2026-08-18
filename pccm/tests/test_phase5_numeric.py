#!/usr/bin/env python3
"""PCCM Phase 5 Gate-A Step-2 tests: the safe-arithmetic numerical primitives.

These cover `builder/pccm_builder/calc_numeric.py` — the layer later VBA
`modCalcFactors` must reproduce.

PROOF SCOPE. Linux, Python, pure functions. NO VBA IS EXECUTED HERE and nothing
in this file proves anything about VBA runtime behaviour (plan §21.0). What it
proves is that the reference semantics are the ones the plan locks: stable forms
where a naive form would overflow, iterative factor series, and refusal rather
than fabrication when a value is not representable.

Runs standalone or under pytest.
"""

from __future__ import annotations

import ast
import itertools
import math
import sys
from fractions import Fraction
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

NUMERIC_PATH = PCCM_ROOT / "builder" / "pccm_builder" / "calc_numeric.py"

from pccm_builder.calc_numeric import (  # noqa: E402
    MAX_DOUBLE,
    MIN_NORMAL_DOUBLE,
    ModelInputRefusal,
    NumericalRangeRefusal,
    OracleInvariantError,
    beta_pert_mean,
    compound_inflation_factors,
    discount_factor_series,
    identity_allowance,
    is_usable_double,
    midpoint,
    safe_accumulate,
    safe_add,
    safe_divide,
    safe_multiply,
    exact_sum_of_products,
    safe_product,
    safe_signed_sum,
    safe_subtract,
    safe_sum,
    triangular_mean,
)


def _refuses(call, reason: str) -> str:
    try:
        call()
    except NumericalRangeRefusal as error:
        return str(error)
    except ModelInputRefusal as error:
        return str(error)
    raise AssertionError(f"{reason}: no refusal was raised")


# ---------------------------------------------------------------------------
# IsUsableDouble
# ---------------------------------------------------------------------------
def test_is_usable_double_accepts_ordinary_finite_values() -> None:
    for value in (0.0, -0.0, 1.0, -1.0, 1e-320, MAX_DOUBLE, -MAX_DOUBLE, 42, -7):
        assert is_usable_double(value), value


def test_is_usable_double_rejects_non_finite_and_non_numeric() -> None:
    for value in (float("nan"), float("inf"), float("-inf"), None, "1.0", True, False, [1]):
        assert not is_usable_double(value), value


def test_the_max_double_constant_is_the_ieee_754_maximum() -> None:
    assert MAX_DOUBLE == 1.7976931348623157e308
    assert math.isinf(MAX_DOUBLE * 1.0000000000000002)


# ---------------------------------------------------------------------------
# Safe primitives
# ---------------------------------------------------------------------------
def test_the_primitives_return_the_ordinary_result_when_it_is_representable() -> None:
    assert safe_add(2.0, 3.0) == 5.0
    assert safe_subtract(2.0, 3.0) == -1.0
    assert safe_multiply(2.0, 3.0) == 6.0
    assert safe_divide(3.0, 2.0) == 1.5
    assert safe_accumulate(10.0, 2.5) == 12.5
    assert safe_sum([1.0, 2.0, 3.0]) == 6.0


def test_overflow_is_refused_rather_than_becoming_infinity() -> None:
    """Python does not raise on float overflow; it produces `inf`. An `inf` that
    reached a total would be reported as a result, so it is refused here."""
    _refuses(lambda: safe_add(MAX_DOUBLE, MAX_DOUBLE), "overflowing addition")
    _refuses(lambda: safe_subtract(-MAX_DOUBLE, MAX_DOUBLE), "overflowing subtraction")
    _refuses(lambda: safe_multiply(1e308, 10.0), "overflowing multiplication")
    _refuses(lambda: safe_divide(1e308, 1e-10), "overflowing division")


def test_division_by_zero_is_refused_outright() -> None:
    message = _refuses(lambda: safe_divide(1.0, 0.0), "division by zero")
    assert "division by zero" in message


def test_a_product_of_non_zero_values_may_not_collapse_to_zero() -> None:
    """The silent failure mode: a contribution deleted with no error anywhere."""
    _refuses(lambda: safe_multiply(1e-200, 1e-200), "underflowing multiplication")


def test_a_quotient_of_a_non_zero_value_may_not_collapse_to_zero() -> None:
    _refuses(lambda: safe_divide(1e-200, 1e200), "underflowing division")


def test_an_exact_zero_operand_is_not_an_underflow() -> None:
    """Zero times something is legitimately zero; only a COLLAPSE is refused."""
    assert safe_multiply(0.0, 1e300) == 0.0
    assert safe_divide(0.0, 1e300) == 0.0


def test_non_finite_operands_are_refused_before_any_arithmetic() -> None:
    for bad in (float("nan"), float("inf"), float("-inf")):
        _refuses(lambda b=bad: safe_add(1.0, b), f"operand {bad}")
        _refuses(lambda b=bad: safe_multiply(1.0, b), f"operand {bad}")


def test_a_refusal_message_names_the_stage() -> None:
    message = _refuses(
        lambda: safe_multiply(1e308, 10.0, "Knom for CL-001"), "named multiplication"
    )
    assert "Knom for CL-001" in message


def test_accumulation_is_checked_at_the_term_that_fails() -> None:
    """Checking during accumulation, not at the end, is what names the culprit."""
    message = _refuses(
        lambda: safe_sum([1e308, 1e308, 1.0], "totals"), "overflowing accumulation"
    )
    assert "totals[1]" in message


# ---------------------------------------------------------------------------
# Stable products
# ---------------------------------------------------------------------------
def test_a_representable_product_survives_a_bad_multiplication_order() -> None:
    """THE LOCKED CASE: `1e308 * 10 * 0.01 = 1e307` is representable.

    Left to right, the first multiplication overflows. Refusing the whole
    calculation for that reason would be refusing an answer that exists.
    """
    assert 1e308 * 10 == math.inf                      # the naive order does overflow
    assert safe_product([1e308, 10.0, 0.01]) == 1e307


def test_a_genuinely_unrepresentable_product_is_still_refused() -> None:
    """No ordering rescues `1e308 * 10`, and none should."""
    _refuses(lambda: safe_product([1e308, 10.0]), "unrepresentable product")


def test_ordinary_products_are_evaluated_left_to_right_unchanged() -> None:
    """Tier 1 is the naive order, so no existing result moves."""
    for factors in ([2.0, 3.0, 4.0], [100.0, 10.0, 1.1085375], [0.3, 250.0, 1.0]):
        expected = 1.0
        for factor in factors:
            expected *= factor
        assert safe_product(factors) == expected, factors


def test_the_product_rescue_is_deterministic() -> None:
    """Same inputs, same answer, every time — the property VBA must reproduce."""
    factors = [1e308, 10.0, 0.01]
    assert safe_product(factors) == safe_product(list(factors)) == 1e307
    assert {safe_product([1e100, 0.5, 1e150, 5e-324, 1e-250]) for _ in range(20)} == {5e-324}


def test_a_zero_factor_makes_an_exactly_zero_product() -> None:
    assert safe_product([0.0, 1e300, 5.0]) == 0.0
    assert safe_product([1e-300, 0.0, 1e-300]) == 0.0


def test_product_signs_are_preserved_through_the_rescue() -> None:
    assert safe_product([-1e308, 10.0, 0.01]) == -1e307
    assert safe_product([-1e308, -10.0, 0.01]) == 1e307
    assert safe_product([-1e100, 0.5, 1e150, 5e-324, 1e-250]) == -5e-324


def test_an_empty_product_is_the_multiplicative_identity() -> None:
    assert safe_product([]) == 1.0


# ---------------------------------------------------------------------------
# Stable distribution statistics
# ---------------------------------------------------------------------------
def test_the_stable_forms_survive_inputs_the_naive_forms_cannot() -> None:
    """Plan §19.2, asserted against its own literals."""
    assert (1e308 + 1e308 + 1e308) == math.inf          # naive triangular numerator
    assert triangular_mean(1e308, 1e308, 1e308) == 1e308
    assert beta_pert_mean(1e308, 1e308, 1e308) == 1e308
    assert (1.5e308 + 1.5e308) == math.inf              # naive midpoint numerator
    assert midpoint(1.5e308, 1.5e308) == 1.5e308


def test_the_stable_forms_agree_with_exact_rational_arithmetic() -> None:
    """An INDEPENDENT oracle: `Fraction` is exact, and is used only in the test."""
    cases = [(80, 100, 150), (100, 200, 450), (0, 0, 0), (-50, -20, 10), (1, 2, 3)]
    for minimum, most_likely, maximum in cases:
        exact_tri = Fraction(minimum + most_likely + maximum, 3)
        exact_pert = Fraction(minimum + 4 * most_likely + maximum, 6)
        exact_mid = Fraction(minimum + maximum, 2)
        assert abs(Fraction(triangular_mean(minimum, most_likely, maximum)) - exact_tri) <= (
            abs(exact_tri) * Fraction(1, 10**14) + Fraction(1, 10**9)
        )
        assert abs(Fraction(beta_pert_mean(minimum, most_likely, maximum)) - exact_pert) <= (
            abs(exact_pert) * Fraction(1, 10**14) + Fraction(1, 10**9)
        )
        assert Fraction(midpoint(minimum, maximum)) == exact_mid


def test_the_mandated_stable_form_is_not_bit_identical_to_the_naive_form() -> None:
    """A REAL CONSEQUENCE OF THE LOCKED DESIGN, recorded rather than hidden.

    `(80 + 4*100 + 150)/6` is exactly `105.0`; the mandated
    `Min/6 + ML*(2/3) + Max/6` is `104.99999999999999` — one unit in the last
    place away. The stable form is required, so this deviation is accepted, and
    hand-derived literals must be compared with a Double-appropriate tolerance
    rather than by exact equality.
    """
    naive = (80 + 4 * 100 + 150) / 6
    stable = beta_pert_mean(80, 100, 150)
    assert naive == 105.0
    assert stable != naive
    assert abs(stable - 105.0) < 1e-13
    # Triangular happens to land exactly on this input; the difference is
    # input-dependent, not systematic.
    assert triangular_mean(80, 100, 150) == 110.0


# ---------------------------------------------------------------------------
# Inflation factors
# ---------------------------------------------------------------------------
def test_the_base_year_factor_is_exactly_one() -> None:
    factors = compound_inflation_factors(2026, 2026, {}, "Standard")
    assert factors == {2026: 1.0}


def test_inflation_compounds_by_calendar_year() -> None:
    rates = {2027: 0.05, 2028: 0.05, 2029: 0.05}
    factors = compound_inflation_factors(2026, 2029, rates, "Standard")
    assert factors[2026] == 1.0
    assert factors[2027] == 1.05
    assert factors[2028] == 1.1025
    assert abs(factors[2029] - 1.157625) < 1e-12


def test_zero_inflation_leaves_every_factor_at_one() -> None:
    factors = compound_inflation_factors(2026, 2030, {y: 0.0 for y in range(2027, 2031)}, "Flat")
    assert set(factors.values()) == {1.0}


def test_negative_but_valid_inflation_is_accepted() -> None:
    rates = {2027: -0.02, 2028: -0.02, 2029: -0.02}
    factors = compound_inflation_factors(2026, 2029, rates, "Deflation")
    assert factors[2027] == 0.98
    assert abs(factors[2028] - 0.9604) < 1e-12
    assert abs(factors[2029] - 0.941192) < 1e-12


def test_an_inflation_rate_of_minus_one_hundred_percent_is_refused() -> None:
    message = _refuses(
        lambda: compound_inflation_factors(2026, 2027, {2027: -1.0}, "Collapse"),
        "inflation rate of -100%",
    )
    assert "Collapse" in message and "2027" in message


def test_an_inflation_rate_below_minus_one_hundred_percent_is_refused() -> None:
    _refuses(
        lambda: compound_inflation_factors(2026, 2027, {2027: -1.5}, "Negative"),
        "inflation rate below -100%",
    )


def test_inflation_overflow_names_the_profile_and_the_calendar_year() -> None:
    rates = {2027: 1e300, 2028: 1e300, 2029: 1e300}
    message = _refuses(
        lambda: compound_inflation_factors(2026, 2029, rates, "Runaway"), "inflation overflow"
    )
    assert "Runaway" in message
    assert "2028" in message      # year 1 reaches 1e300; year 2 is the failure


# ---------------------------------------------------------------------------
# Discount factors
# ---------------------------------------------------------------------------
def test_project_year_one_is_period_zero() -> None:
    assert discount_factor_series(0.10, 3)[1] == 1.0


def test_discount_factors_are_built_iteratively() -> None:
    factors = discount_factor_series(0.10, 3)
    assert factors[2] == 1.0 / 1.1
    assert factors[3] == (1.0 / 1.1) / 1.1


def test_the_iterative_series_matches_the_locked_closed_form() -> None:
    """`1/(1+r)^(t-1)` is the SEMANTICS; iteration is the implementation path."""
    for rate in (0.0, 0.05, 0.10, -0.05, 0.5):
        factors = discount_factor_series(rate, 12)
        for index in range(1, 13):
            closed_form = 1.0 / (1.0 + rate) ** (index - 1)
            assert abs(factors[index] - closed_form) <= abs(closed_form) * 1e-12


def test_a_zero_discount_rate_leaves_every_factor_at_one() -> None:
    assert set(discount_factor_series(0.0, 5).values()) == {1.0}


def test_a_negative_discount_rate_above_minus_one_hundred_percent_is_accepted() -> None:
    factors = discount_factor_series(-0.05, 3)
    assert factors[1] == 1.0
    assert factors[2] == 1.0 / 0.95
    assert factors[3] == (1.0 / 0.95) / 0.95
    assert factors[3] > factors[2] > factors[1]     # PV exceeds nominal, correctly


def test_a_discount_rate_of_minus_one_hundred_percent_is_refused() -> None:
    _refuses(lambda: discount_factor_series(-1.0, 3), "discount rate of -100%")


def test_a_discount_rate_below_minus_one_hundred_percent_is_refused() -> None:
    _refuses(lambda: discount_factor_series(-1.5, 3), "discount rate below -100%")


def test_discount_underflow_refuses_at_the_exact_project_year(
) -> None:
    """Plan §19.3, case 29: `r = 1e10`.

    `disc(33)` is subnormal but non-zero; `disc(34)` is exactly zero, and a zero
    factor would silently delete a year's entire PV contribution.
    """
    factors = discount_factor_series(1e10, 33)
    assert factors[33] > 0.0
    assert factors[33] < 1e-300                     # subnormal territory
    message = _refuses(lambda: discount_factor_series(1e10, 34), "discount underflow")
    assert "project year 34" in message


# ---------------------------------------------------------------------------
# Cancellation-aware tolerance
# ---------------------------------------------------------------------------
def test_the_allowance_is_the_absolute_floor_for_a_tiny_model() -> None:
    assert identity_allowance([1.0, 2.0, 3.0], 1e-6, 1e-12) == 1e-6


def test_the_allowance_grows_with_the_magnitude_of_the_arithmetic() -> None:
    allowance = identity_allowance([1e9, -1e9, 0.0], 1e-6, 1e-12)
    assert allowance > 1e-6
    assert abs(allowance - 2e-3) < 1e-9


def test_the_scale_reflects_the_terms_not_the_net_result() -> None:
    """Cancellation is the whole point: two models with the same near-zero net but
    very different arithmetic must not get the same tolerance."""
    small = identity_allowance([1.0, -1.0, 0.0], 1e-6, 1e-12)
    large = identity_allowance([1e12, -1e12, 0.0], 1e-6, 1e-12)
    assert large > small


def test_the_allowance_survives_terms_whose_raw_sum_overflows() -> None:
    """Plan item: the raw conditioning sum may exceed Double while `1e-12 * scale`
    is perfectly representable. Distributing the coefficient is what keeps it."""
    terms = [1.5e308, 1.5e308, 1.5e308]
    assert sum(terms) == math.inf                    # the naive scale overflows
    allowance = identity_allowance(terms, 1e-6, 1e-12)
    assert math.isfinite(allowance)
    assert abs(allowance - 4.5e296) < 1e290


def test_the_allowance_is_a_maximum_not_a_sum() -> None:
    """The locked formula is

        max(absolute_floor, coefficient * max(scale_floor, sum |terms|))

    and the inner operation is a MAXIMUM. An earlier implementation added
    `coefficient * scale_floor` to the scaled sum, returning `1.000001e-6` for
    `[1e6]` where the contract says exactly `1e-6`. Small, but a tolerance may
    never be loosened by accident.
    """
    assert identity_allowance([1e6], 1e-6, 1e-12) == 1e-6
    assert identity_allowance([1e6], 1e-6, 1e-12) != 1e-6 + 1e-12
    assert identity_allowance([], 1e-6, 1e-12) == 1e-6
    assert identity_allowance([1e18], 1e-6, 1e-12) == 1e-12 * 1e18


def test_the_allowance_matches_the_locked_formula_computed_directly() -> None:
    """An independent evaluation of the same formula, term by term."""
    for terms in ([], [1.0], [1e6], [1e9, -1e9], [1e15, 2e15, 3e15], [-4.5, 0.0, 12.25]):
        scale = max(1.0, sum(abs(t) for t in terms))
        expected = max(1e-6, 1e-12 * scale)
        assert abs(identity_allowance(terms, 1e-6, 1e-12) - expected) <= expected * 1e-15, terms


def test_a_conditioning_term_too_small_to_scale_does_not_refuse_the_allowance() -> None:
    """CONDITIONING METADATA HAS ITS OWN UNDERFLOW POLICY.

    `1e-12 * 2e-312` rounds to exactly zero. Under the model-arithmetic rule that
    is a refusal — and it was, which rejected a model whose economic outputs are
    perfectly representable, purely because the tolerance bookkeeping could not
    hold a term far too small to affect the answer.

    The locked allowance for this input has no ambiguity at all:

        conditioning scale = max(1, 2e-312) = 1
        relative allowance = 1e-12
        final allowance    = max(1e-6, 1e-12) = 1e-6
    """
    assert 1e-12 * 2e-312 == 0.0                      # the scaled term really does vanish
    assert identity_allowance([2e-312], 1e-6, 1e-12, 1.0) == 1e-6
    assert identity_allowance([5e-324], 1e-6, 1e-12, 1.0) == 1e-6
    assert identity_allowance([2e-312, 1e-320, 4e-315], 1e-6, 1e-12, 1.0) == 1e-6


def test_a_vanishing_conditioning_term_cannot_move_a_scale_that_does_matter() -> None:
    """Mixing huge and vanishing terms: the dropped amount is far under one ulp."""
    big_only = identity_allowance([1e18], 1e-6, 1e-12, 1.0)
    with_tiny = identity_allowance([1e18, 2e-312, 5e-324], 1e-6, 1e-12, 1.0)
    assert with_tiny == big_only == 1e-12 * 1e18


def test_the_conditioning_exception_does_not_weaken_model_arithmetic() -> None:
    """The relaxed rule is scoped to conditioning metadata and nowhere else.

    An economic value or factor that collapses to zero is still refused, because
    there it would delete a real contribution with no error anywhere.
    """
    _refuses(lambda: safe_multiply(1e-200, 1e-200), "economic underflow")
    _refuses(lambda: safe_divide(1e-200, 1e200), "factor underflow")
    _refuses(lambda: safe_product([1e-200, 1e-200]), "product underflow")


def test_conditioning_overflow_is_still_refused() -> None:
    """A conditioning scale beyond Double makes the allowance itself
    unrepresentable, so any comparison against it would be meaningless."""
    _refuses(
        lambda: identity_allowance([1e308] * 40, 1e-6, 1.0, 1.0),
        "overflowing conditioning scale",
    )


def test_the_boundary_where_the_relative_allowance_meets_the_absolute_floor() -> None:
    """`1e-12 * scale` equals `1e-6` exactly at `scale = 1e6`."""
    assert identity_allowance([1e6 - 1.0], 1e-6, 1e-12) == 1e-6      # below: floor binds
    assert identity_allowance([1e6], 1e-6, 1e-12) == 1e-6            # at: they coincide
    above = identity_allowance([1e6 + 1e6], 1e-6, 1e-12)
    assert above > 1e-6                                              # beyond: relative binds
    assert abs(above - 2e-6) < 1e-18


def test_the_scale_floor_binds_only_below_unity() -> None:
    """`max(scale_floor, sum|terms|)` — for an empty or tiny model the floor of 1
    applies, and `1e-12 * 1` is still far under the absolute floor."""
    assert identity_allowance([], 1e-6, 1e-12, 1.0) == 1e-6
    assert identity_allowance([0.5], 1e-6, 1e-12, 1.0) == 1e-6
    # With a raised scale floor the relative term can overtake the absolute one.
    assert identity_allowance([], 1e-6, 1e-12, 1e9) == 1e-3


# ---------------------------------------------------------------------------
# Failure classification
# ---------------------------------------------------------------------------
def test_a_python_integer_too_large_for_a_double_is_not_usable() -> None:
    """`float(10**400)` RAISES `OverflowError` rather than returning `inf`.

    A predicate that raises is not a predicate, and a raw `OverflowError` escaping
    the kernel would bypass the entire failure contract. The pure oracle accepts
    plain Python numbers, so it has to answer "no" and let the caller turn that
    into a structured refusal.
    """
    huge = 10 ** 400
    try:
        float(huge)
    except OverflowError:
        pass
    else:
        raise AssertionError("this Python no longer overflows on the test value")

    assert is_usable_double(huge) is False
    assert is_usable_double(-huge) is False
    assert is_usable_double(10 ** 309) is False
    assert is_usable_double(10 ** 300) is True


def test_the_primitives_refuse_a_huge_integer_rather_than_raising() -> None:
    huge = 10 ** 400
    for call in (
        lambda: safe_add(1.0, huge),
        lambda: safe_multiply(1.0, huge),
        lambda: safe_divide(huge, 2.0),
        lambda: safe_product([2.0, huge]),
    ):
        _refuses(call, "huge integer operand")


def test_the_failure_hierarchy_separates_refusals_from_internal_defects() -> None:
    """A refusal is a statement about the model; an invariant failure is a
    statement about the implementation. They must never be confused."""
    from pccm_builder.calc_numeric import CalculationRefusal, OracleError

    assert issubclass(ModelInputRefusal, CalculationRefusal)
    assert issubclass(NumericalRangeRefusal, CalculationRefusal)
    assert issubclass(CalculationRefusal, OracleError)
    assert issubclass(OracleInvariantError, OracleError)
    assert not issubclass(OracleInvariantError, CalculationRefusal)


# ---------------------------------------------------------------------------
# Erratum C2 - a representable final Double is never refused for an
# intermediate that left the range
# ---------------------------------------------------------------------------
# A SIGNED SUM whose true total is representable must produce it, even when the
# canonical left-to-right partial sums are not. `MAX + MAX - MAX` is `MAX`, and
# refusing it refuses an answer that exists.
def test_the_four_locked_signed_sum_vectors() -> None:
    """The vectors the patch is specified against, asserted exactly."""
    assert safe_signed_sum([MAX_DOUBLE, MAX_DOUBLE, -MAX_DOUBLE]) == MAX_DOUBLE
    assert safe_signed_sum([MAX_DOUBLE, MAX_DOUBLE, -MAX_DOUBLE, -MAX_DOUBLE, 1.0]) == 1.0
    assert (
        safe_signed_sum([MAX_DOUBLE, MAX_DOUBLE, -MAX_DOUBLE, -MAX_DOUBLE, 5e-324]) == 5e-324
    )
    _refuses(
        lambda: safe_signed_sum([MAX_DOUBLE, MAX_DOUBLE]),
        "a sum with no cancellation and no representable total",
    )


def test_a_sum_that_already_succeeds_is_never_reordered() -> None:
    """TIER 1 IS THE CONTRACT for every ordinary model.

    The rescue exists to turn a refusal into an answer, not to improve answers.
    Wherever plain left-to-right accumulation produces a value, `safe_signed_sum`
    must return THAT value, bit for bit — including where left to right is the
    less accurate order. Anything else would move the numbers of models that work
    today, and would make canonical permanent-ID order stop determining the
    result.
    """
    corpus = [
        [0.1, 0.2, 0.3],
        [1e16, 1.0, -1e16],                      # left to right loses the 1.0
        [1e16, -1e16, 1.0],                      # the same terms, and it does not
        [MAX_DOUBLE, 1.0, -MAX_DOUBLE],          # absorbed, then cancelled to zero
        [5e-324, 1.0, -1.0],
        [-3.5, 2.25, 1.0],
        [1e300, 1e-300],
    ]
    for terms in corpus:
        expected = safe_sum(terms)
        actual = safe_signed_sum(terms)
        assert actual == expected and math.copysign(1.0, actual) == math.copysign(
            1.0, expected
        ), f"{terms}: tier 1 gave {expected!r} but safe_signed_sum gave {actual!r}"


def test_the_rescue_is_reached_only_when_the_canonical_order_produced_nothing() -> None:
    """`[1e16, 1.0, -1e16]` sums to `0.0` left to right and to `1.0` if reordered.

    Tier 1 succeeds, so `0.0` is the answer. A rescue that ran unconditionally —
    or that ran on any inaccuracy rather than only on overflow — would return
    `1.0` here and silently change every model that has ever been calculated.
    """
    assert safe_signed_sum([1e16, 1.0, -1e16]) == 0.0
    assert safe_sum([1e16, 1.0, -1e16]) == 0.0


def test_the_rescue_names_the_stage_and_is_deterministic() -> None:
    message = _refuses(
        lambda: safe_signed_sum([MAX_DOUBLE, MAX_DOUBLE], "headline totals"),
        "an unrepresentable signed total",
    )
    assert "headline totals" in message
    vector = [MAX_DOUBLE, MAX_DOUBLE, -MAX_DOUBLE, -MAX_DOUBLE, 5e-324]
    assert {safe_signed_sum(vector) for _ in range(20)} == {5e-324}


def test_a_signed_sum_names_the_term_that_failed() -> None:
    """Tier-1 messages still identify the driver or year, as §19.4 requires."""
    message = _refuses(
        lambda: safe_signed_sum(
            [MAX_DOUBLE, MAX_DOUBLE], "totals", ["driver 'CL-001'", "driver 'CL-002'"]
        ),
        "an unrepresentable signed total",
    )
    assert "totals" in message


def test_exact_cancellation_returns_positive_zero() -> None:
    """`+0.0`, not `-0.0`: the sign of a total that cancelled to nothing is not a
    fact about the model, and a negative zero would propagate into comparisons and
    audit columns for no reason."""
    total = safe_signed_sum([MAX_DOUBLE, -MAX_DOUBLE, MAX_DOUBLE, -MAX_DOUBLE])
    assert total == 0.0 and math.copysign(1.0, total) == 1.0


def test_a_negative_signed_total_survives_cancellation() -> None:
    assert safe_signed_sum([-MAX_DOUBLE, -MAX_DOUBLE, MAX_DOUBLE]) == -MAX_DOUBLE
    _refuses(
        lambda: safe_signed_sum([-MAX_DOUBLE, -MAX_DOUBLE]),
        "an unrepresentable negative signed total",
    )


def test_the_rescue_agrees_with_exact_rational_arithmetic_where_it_fires() -> None:
    """An INDEPENDENT oracle. `Fraction` is exact and is used only in the test —
    never in production, where IEEE-754 Double IS the semantic target."""
    vectors = [
        [MAX_DOUBLE, MAX_DOUBLE, -MAX_DOUBLE],
        [MAX_DOUBLE, MAX_DOUBLE, -MAX_DOUBLE, -MAX_DOUBLE, 1.0],
        [MAX_DOUBLE, MAX_DOUBLE, -MAX_DOUBLE, -MAX_DOUBLE, 5e-324],
        [MAX_DOUBLE, MAX_DOUBLE, MAX_DOUBLE, -MAX_DOUBLE, -MAX_DOUBLE],
        [1e308, 1e308, 1e308, -1e308, -1e308, -1e308],
        [MAX_DOUBLE, 1e308, -MAX_DOUBLE, -1e308, -7.5],
    ]
    for terms in vectors:
        exact = sum((Fraction(t) for t in terms), Fraction(0))
        assert Fraction(safe_signed_sum(terms)) == exact, f"{terms}: rescue disagrees"


def test_a_signed_sum_that_genuinely_exceeds_double_range_is_still_refused() -> None:
    """The rescue repairs an intermediate that left the range. It must not repair a
    RESULT that is genuinely outside it — that would be fabrication."""
    for terms in (
        [MAX_DOUBLE, MAX_DOUBLE],
        [MAX_DOUBLE, MAX_DOUBLE, MAX_DOUBLE, -MAX_DOUBLE],
        [MAX_DOUBLE, MAX_DOUBLE, -1.0],
        [-MAX_DOUBLE, -MAX_DOUBLE, MAX_DOUBLE, -MAX_DOUBLE],
    ):
        _refuses(lambda t=terms: safe_signed_sum(t), f"{terms} has no representable total")


def test_a_non_finite_term_is_refused_before_any_tier_runs() -> None:
    for bad in (math.inf, -math.inf, math.nan, 10**400):
        _refuses(
            lambda b=bad: safe_signed_sum([1.0, b, -1.0]),
            f"{bad!r} is not a usable Double",
        )


# A CONVEX COMBINATION is bracketed by its own points, so it always has a
# representable answer. Tier 0 returns a zero-uncertainty point exactly, tier 1 is
# the accepted stable form untouched, tier 2 rescues the rest.
DEGENERATE_POINTS = (
    0.0, 1.0, -1.0, 1e-300, -1e-300, 1e-320, 5e-324, -5e-324,
    1e308, MAX_DOUBLE, -MAX_DOUBLE, math.nextafter(MAX_DOUBLE, 0.0),
    0.1, -12345.6789,
)


def test_a_distribution_with_zero_uncertainty_returns_its_point_exactly() -> None:
    """§9, EXACTLY — no last-ulp drift anywhere in the usable Double range.

    `x/3 + x/3 + x/3 != x` for many subnormal `x`, and `x/2 + x/2` cannot even be
    formed for the smallest one, so the stable forms alone do not give this. A
    distribution with `Min == ML == Max` has no uncertainty at all, and answering
    anything but that number is wrong regardless of how small the error is.
    """
    for point in DEGENERATE_POINTS:
        assert triangular_mean(point, point, point) == point, f"triangular at {point!r}"
        assert beta_pert_mean(point, point, point) == point, f"Beta-PERT at {point!r}"
        assert midpoint(point, point) == point, f"midpoint at {point!r}"


def test_the_degenerate_invariant_is_needed_and_not_incidental() -> None:
    """Proof that tier 0 is doing work: at these points the stable form drifts."""
    assert 5e-324 / 3.0 + 5e-324 / 3.0 + 5e-324 / 3.0 != 5e-324
    assert 5e-324 / 2.0 == 0.0                        # the stable midpoint cannot form
    assert triangular_mean(5e-324, 5e-324, 5e-324) == 5e-324
    assert midpoint(5e-324, 5e-324) == 5e-324


def test_the_stable_forms_still_carry_a_non_degenerate_overflow() -> None:
    """Tier 1 is not made redundant by tier 0: with three DIFFERENT huge points the
    naive numerator still overflows and the stable form still answers."""
    assert (1e308 + 1.1e308 + 1.2e308) == math.inf
    assert triangular_mean(1e308, 1.1e308, 1.2e308) == 1.1e308
    assert beta_pert_mean(1e308, 1.1e308, 1.2e308) == 1.1e308


# The boundary corpus of §11. `Decimal`/`Fraction` appear here and nowhere in
# production.
BOUNDARY_CORPUS = (
    MAX_DOUBLE, math.nextafter(MAX_DOUBLE, 0.0), 1e308, 1.0, 1e-300, 1e-320, 5e-324, 0.0,
    -MAX_DOUBLE, -math.nextafter(MAX_DOUBLE, 0.0), -1e308, -1.0, -1e-300, -1e-320, -5e-324,
)

# (label, function, exact weights). The rescue is CORRECTLY ROUNDED for all
# three: it forms the exact numerator, divides once with a guard and a sticky
# remainder, and rounds once. The round-2 binade rescue was one ulp out on
# Beta-PERT; the exact form is not.
_STATISTICS = (
    ("triangular mean", triangular_mean, (Fraction(1, 3), Fraction(1, 3), Fraction(1, 3))),
    ("Beta-PERT mean", beta_pert_mean, (Fraction(1, 6), Fraction(2, 3), Fraction(1, 6))),
    ("midpoint", midpoint, (Fraction(1, 2), Fraction(1, 2))),
)


def _ulps_apart(a: float, b: float, limit: int = 64) -> int:
    if a == b:
        return 0
    low, high = (a, b) if a < b else (b, a)
    distance = 0
    while low < high and distance < limit:
        low = math.nextafter(low, math.inf)
        distance += 1
    return distance


def test_no_convex_statistic_is_refused_when_its_answer_is_representable() -> None:
    """THE CLASS THIS PATCH EXISTS TO CLOSE, swept over the whole boundary corpus.

    For every combination of boundary points, the exact rational statistic is
    computed with `Fraction` and rounded once to Double. If that rounding produces
    a usable non-zero Double, the oracle must produce a value; a refusal there is
    a refusal of an answer that exists.
    """
    for label, function, weights in _STATISTICS:
        for points in itertools.product(BOUNDARY_CORPUS, repeat=len(weights)):
            exact = sum(
                (weight * Fraction(point) for weight, point in zip(weights, points)),
                Fraction(0),
            )
            try:
                rounded = float(exact)
            except OverflowError:
                rounded = None
            representable = rounded is not None and abs(rounded) <= MAX_DOUBLE
            try:
                actual = function(*points)
            except NumericalRangeRefusal as error:
                assert not (representable and rounded != 0.0), (
                    f"{label}{points} refused ({error}) but rounds to {rounded!r}"
                )
                continue
            assert representable, f"{label}{points} returned {actual!r}, which is out of range"


def test_the_convex_statistics_are_correctly_rounded_on_the_boundary_corpus() -> None:
    """Stronger than "not refused": where the exact statistic has an exactly
    representable Double, that is the Double returned.

    `Fraction` is exact, so `float(exact)` is the correctly rounded answer. This
    is what pins the rescue's scale-back to a SINGLE rounding: halving one step at
    a time through the subnormal range rounds twice and lands a unit low, which is
    enough to turn `5e-324` into `0.0` and a value into a refusal.
    """
    rescued = 0
    for label, function, weights in _STATISTICS:
        for points in itertools.product(BOUNDARY_CORPUS, repeat=len(weights)):
            exact = sum(
                (weight * Fraction(point) for weight, point in zip(weights, points)),
                Fraction(0),
            )
            try:
                rounded = float(exact)
            except OverflowError:
                continue
            if abs(rounded) > MAX_DOUBLE or rounded == 0.0:
                continue
            if not _tier_two_fired(function, points):
                continue                    # tier 1's own rounding is accepted, not tested here
            rescued += 1
            try:
                actual = function(*points)
            except NumericalRangeRefusal:
                raise AssertionError(f"{label}{points} refused; exact rounds to {rounded!r}")
            assert actual == rounded, (
                f"{label}{points}: rescue gave {actual!r}, correctly rounded is {rounded!r}"
            )
    assert rescued > 1000, (
        f"only {rescued} corpus inputs actually reached the rescue; the sweep would pass "
        "vacuously"
    )


def _tier_two_fired(function, points) -> bool:
    """True when the accepted stable form could not produce a value on its own.

    Tier 1's own rounding is accepted and deliberately NOT compared against the
    exact value (see `test_the_mandated_stable_form_is_not_bit_identical...`);
    what must be correctly rounded is the rescue, so the check is scoped to the
    inputs that actually reach it.
    """
    from pccm_builder import calc_numeric

    stable_only = {
        triangular_mean: lambda v: (
            calc_numeric.safe_accumulate(
                calc_numeric.safe_accumulate(
                    calc_numeric.safe_divide(v[0], 3.0), calc_numeric.safe_divide(v[1], 3.0)
                ),
                calc_numeric.safe_divide(v[2], 3.0),
            )
        ),
        beta_pert_mean: lambda v: (
            calc_numeric.safe_accumulate(
                calc_numeric.safe_accumulate(
                    calc_numeric.safe_divide(v[0], 6.0),
                    calc_numeric.safe_multiply(v[1], 2.0 / 3.0),
                ),
                calc_numeric.safe_divide(v[2], 6.0),
            )
        ),
        midpoint: lambda v: calc_numeric.safe_accumulate(
            calc_numeric.safe_divide(v[0], 2.0), calc_numeric.safe_divide(v[1], 2.0)
        ),
    }[function]
    try:
        stable_only(points)
    except NumericalRangeRefusal:
        return True
    return False


def test_a_statistic_with_no_usable_non_zero_double_is_still_refused() -> None:
    """The other side of §10. `midpoint(5e-324, 0)` is `2.47e-324`, which is below
    half the smallest subnormal and rounds to zero: there is no usable non-zero
    Double for it, and the existing underflow contract refuses rather than
    silently deleting the value."""
    _refuses(lambda: midpoint(5e-324, 0.0), "a statistic that rounds to zero")
    _refuses(lambda: triangular_mean(5e-324, 5e-324, -5e-324), "a statistic that rounds to zero")
    assert float(Fraction(5e-324) / 2) == 0.0


def test_the_exact_rescues_use_only_vba_translatable_operations() -> None:
    """A STRUCTURAL guard on cross-language reproducibility (§10).

    The rescues must be expressible in VBA with `Double` arithmetic, `Fix` and
    counting loops. `math.frexp`, `math.ldexp`, `math.fsum`, `Decimal` and
    `Fraction` would each be correct in Python and untranslatable, so none may
    appear in the module at all — and neither may Python's arbitrary-precision
    `int`, which is why `_fix` is written with Double operations instead of
    `math.floor`.
    """
    source = NUMERIC_PATH.read_text(encoding="utf-8")
    tree = ast.parse(source)
    names: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Attribute):
            names.add(node.attr)
        elif isinstance(node, ast.Name):
            names.add(node.id)
    for banned in (
        "frexp", "ldexp", "fsum", "Decimal", "Fraction", "nextafter", "floor",
        "trunc", "getcontext", "as_integer_ratio",
    ):
        assert banned not in names, f"calc_numeric.py uses {banned}, which VBA has no form of"
    # `int` appears as an isinstance TYPE in `is_usable_double`, which is a type
    # test, not arithmetic. What must not appear is a CALL to it: converting a
    # Double to Python's arbitrary-precision integer would make the kernel exact
    # in a way VBA cannot reproduce.
    for node in ast.walk(tree):
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name):
            assert node.func.id not in ("int", "round", "divmod", "pow"), (
                f"calc_numeric.py calls {node.func.id}(), which has no Double-only meaning"
            )


def test_the_truncation_primitive_is_exact_and_matches_vba_fix() -> None:
    """`_fix` is the one place the exact kernel needs integer truncation, and it
    is written in Double operations so VBA `Fix` is a direct substitute."""
    from pccm_builder.calc_numeric import _fix

    for value in (0.0, 1.0, 1.5, 2.9999999999999996, 16777215.0, 16777216.5,
                  4503599627370495.5, 4503599627370496.0, 9007199254740991.0):
        assert _fix(value) == math.floor(value), value
    # exhaustive over a dense band that exercises the round-then-correct path
    for step in range(0, 20000):
        value = step / 7.0
        assert _fix(value) == math.floor(value), value


def test_the_minimum_normal_double_constant_is_the_ieee_754_value() -> None:
    assert MIN_NORMAL_DOUBLE == 2.0**-1022
    assert MIN_NORMAL_DOUBLE / 2.0 != 0.0                      # subnormals exist below it
    assert math.frexp(MIN_NORMAL_DOUBLE) == (0.5, -1021)


def test_the_exact_decomposition_round_trips_every_boundary_double() -> None:
    """`_decompose` is the entry point of both rescues: if it is not exact,
    nothing above it can be. Checked against `Fraction` — test-side only."""
    from pccm_builder.calc_numeric import _decompose

    for value in BOUNDARY_CORPUS + (0.1, 3.0, 2.0**-1022, 1.5e-323, -0.1):
        sign, mantissa, exponent = _decompose(value)
        if value == 0.0:
            assert (sign, mantissa, exponent) == (0, 0.0, 0)
            continue
        assert 4503599627370496.0 <= mantissa < 9007199254740992.0, value
        assert Fraction(sign) * Fraction(mantissa) * Fraction(2) ** exponent == Fraction(
            value
        ), value


# ---------------------------------------------------------------------------
# NEGATIVE CONTROLS - the tests above must fail if the mechanism is removed
# ---------------------------------------------------------------------------
def test_sabotaging_the_signed_sum_rescue_breaks_the_locked_vectors() -> None:
    """§13. Replace tier 2 with a plain left-to-right re-run and the cancellation
    vectors must refuse again. A guard that survives its own removal is not a
    guard."""
    vectors = (
        [MAX_DOUBLE, MAX_DOUBLE, -MAX_DOUBLE],
        [MAX_DOUBLE, MAX_DOUBLE, -MAX_DOUBLE, -MAX_DOUBLE, 1.0],
        [MAX_DOUBLE, MAX_DOUBLE, -MAX_DOUBLE, -MAX_DOUBLE, 5e-324],
    )
    for terms in vectors:
        assert safe_signed_sum(terms) is not None          # works with the rescue
        _refuses(lambda t=terms: safe_sum(t), f"{terms} without the rescue")


def test_sabotaging_the_degenerate_invariant_breaks_the_exactness_vectors() -> None:
    """§13. Without tier 0, the stable form drifts or refuses at these points."""
    saboteur = (5e-324, 1e-320, MAX_DOUBLE, -MAX_DOUBLE)
    for point in saboteur:
        assert triangular_mean(point, point, point) == point
    # The stable form alone, which is what tier 0 replaces:
    assert 5e-324 / 3.0 * 3.0 != 5e-324
    assert MAX_DOUBLE / 3.0 + MAX_DOUBLE / 3.0 + MAX_DOUBLE / 3.0 == math.inf
    assert -MAX_DOUBLE / 3.0 + -MAX_DOUBLE / 3.0 + -MAX_DOUBLE / 3.0 == -math.inf


def test_sabotaging_the_statistic_rescue_breaks_the_subnormal_statistics() -> None:
    """§13. Without the exact rescue, a subnormal Uniform has no midpoint at all."""
    assert midpoint(5e-324, 1e-323) == 1e-323
    _refuses(lambda: safe_divide(5e-324, 2.0), "the stable midpoint's own first step")
    assert float(Fraction(5e-324) / 2 + Fraction(1e-323) / 2) == 1e-323


# ---------------------------------------------------------------------------
# FAITHFUL RESCUE - the exact mathematical target, adversarially swept
# ---------------------------------------------------------------------------
# The round-2 rescues were heuristics: re-associating Doubles for sums, and
# reordering by magnitude for products. Both were shown to be wrong in ways no
# clean-cancellation vector could expose, so the criterion is now stated exactly
# and swept against an independent oracle.
#
# `Fraction.from_float` is EXACT and appears in TEST CODE ONLY. It is not
# production semantics and it is not what the model computes; it is independent
# Gate-A evidence about what the production algorithm should have produced.
def _exact_sum_oracle(terms) -> Fraction:
    return sum((Fraction(term) for term in terms), Fraction(0))


def _exact_product_oracle(factors) -> Fraction:
    exact = Fraction(1)
    for factor in factors:
        exact = exact * Fraction(factor)
    return exact


_EXACT_MAX = Fraction(MAX_DOUBLE)


def _classify(exact: Fraction):
    """(kind, target) for an exact value, using the EXACT range test of §11.

    `abs(exact) <= Fraction(MAX_DOUBLE)` is tested BEFORE `float(exact)`, because
    an exact value can exceed `MAX_DOUBLE` and still round down to it — Python's
    overflow threshold is half an ulp above `MAX_DOUBLE`, not at it. Using
    `math.isfinite(float(exact))` as the definition of "in range" would call that
    case representable and hide exactly the defect this sweep exists to find.
    """
    if abs(exact) > _EXACT_MAX:
        return "out-of-range", None
    target = float(exact)
    if exact != 0 and target == 0.0:
        return "collapsed", None
    return "value", target


def test_the_exact_range_test_is_not_the_same_as_float_overflow() -> None:
    """The half-ulp band §11 warns about, demonstrated rather than asserted.

    `MAX_DOUBLE + 0.5 ulp` is mathematically out of range yet `float()` of it is
    `MAX_DOUBLE`, so a classifier built on `float(exact)` would call it a value.
    """
    # ulp(MAX_DOUBLE) is 2**971, and Python's overflow threshold sits half an ulp
    # above MAX_DOUBLE, so anything in (MAX, MAX + 2**970) is mathematically out of
    # range while `float()` reports it as MAX_DOUBLE.
    beyond = _EXACT_MAX + Fraction(2) ** 969
    assert beyond > _EXACT_MAX
    assert float(beyond) == MAX_DOUBLE                # float() does NOT report overflow
    assert math.isfinite(float(beyond))
    assert _classify(beyond)[0] == "out-of-range"


class _Stream:
    """A fixed linear congruential generator.

    Deterministic and self-contained: the corpus does not depend on the stdlib
    RNG's implementation staying stable across Python versions, and there is no
    unseeded randomness anywhere in this file.
    """

    def __init__(self, seed: int) -> None:
        self._state = seed & 0xFFFFFFFFFFFFFFFF

    def next(self, bound: int) -> int:
        self._state = (self._state * 6364136223846793005 + 1442695040888963407) & (
            0xFFFFFFFFFFFFFFFF
        )
        return (self._state >> 33) % bound

    def pick(self, values):
        return values[self.next(len(values))]


_SUM_CORPUS = (
    MAX_DOUBLE, math.nextafter(MAX_DOUBLE, 0.0), 1e308, 1.78e308, 1.7e308, 1e292,
    1.0, 1e-292, 1e-308, 1e-320, 5e-324,
)
_SUM_CORPUS = _SUM_CORPUS + tuple(-value for value in _SUM_CORPUS)

_PRODUCT_CORPUS = (
    MAX_DOUBLE, 1e308, 1e250, 1e150, 1e100, 1e50, 10.0, 2.0, 1.0, 0.5, 0.1,
    1e-50, 1e-100, 1e-150, 1e-250, 1e-300, 1e-320, 5e-324,
)
_PRODUCT_CORPUS = _PRODUCT_CORPUS + tuple(-value for value in _PRODUCT_CORPUS)

_ADVERSARIAL_CASES = 10000


def _tier_one_sum(terms):
    """Plain canonical accumulation — the tier-1 path, reproduced in the test so
    a case can be classified as "rescued" without asking the module."""
    try:
        return safe_sum(terms), True
    except NumericalRangeRefusal:
        return None, False


def _tier_one_product(factors):
    if any(factor == 0.0 for factor in factors):
        return None, False
    try:
        result = 1.0
        for factor in factors:
            result = safe_multiply(result, factor)
        return result, True
    except NumericalRangeRefusal:
        return None, False


def test_the_signed_sum_rescue_is_faithful_over_ten_thousand_adversarial_cases() -> None:
    """§10. Lengths 2..20 from the boundary corpus, every rescue judged exactly.

    THE CASE THIS EXISTS FOR is rounding-residual loss:
    `[6e307, -8e307, -1.7e308, 6e307, 7e307, 6e307, -1e292]` sums to exactly
    `-1e292`, and a rescue that cancels the largest opposite-signed pair with one
    rounded subtraction throws away the residual and answers
    `-1.99792015476736e292`. Clean-cancellation vectors cannot see that; a sweep
    against exact rationals can.
    """
    stream = _Stream(0x50BE6EB0E26857)
    rescued = 0
    for _ in range(_ADVERSARIAL_CASES):
        length = 2 + stream.next(19)
        terms = [stream.pick(_SUM_CORPUS) for _ in range(length)]
        _, tier_one_worked = _tier_one_sum(terms)
        if tier_one_worked:
            continue                                   # tier 1 owns this case
        rescued += 1
        kind, target = _classify(_exact_sum_oracle(terms))
        try:
            actual = safe_signed_sum(terms)
        except NumericalRangeRefusal as error:
            assert kind != "value", f"{terms} refused ({error}) but is {target!r}"
            continue
        assert kind == "value", f"{terms} returned {actual!r} but is {kind}"
        assert actual == target, f"{terms}: got {actual!r}, correctly rounded is {target!r}"
    assert rescued > 5000, f"only {rescued} cases reached the rescue; the sweep is too easy"


def test_the_product_rescue_is_faithful_over_ten_thousand_adversarial_cases() -> None:
    """§10. Factor counts 2..6 from the product corpus, every rescue judged
    exactly — both the false acceptance and the false refusal directions."""
    stream = _Stream(0x2147483647002A)
    rescued = 0
    for _ in range(_ADVERSARIAL_CASES):
        count = 2 + stream.next(5)
        factors = [stream.pick(_PRODUCT_CORPUS) for _ in range(count)]
        _, tier_one_worked = _tier_one_product(factors)
        if tier_one_worked:
            continue
        rescued += 1
        kind, target = _classify(_exact_product_oracle(factors))
        try:
            actual = safe_product(factors)
        except NumericalRangeRefusal as error:
            assert kind != "value", f"{factors} refused ({error}) but is {target!r}"
            continue
        assert kind == "value", f"{factors} returned {actual!r} but is {kind}"
        assert actual == target, f"{factors}: got {actual!r}, correctly rounded is {target!r}"
    assert rescued > 3000, f"only {rescued} cases reached the rescue; the sweep is too easy"


def test_the_signed_sum_reproducers_of_the_round_three_review() -> None:
    """§1.1 and §1.3, asserted as literals rather than only inside a sweep."""
    residual = [6e307, -8e307, -1.7e308, 6e307, 7e307, 6e307, -1e292]
    assert _classify(_exact_sum_oracle(residual)) == ("value", -1e292)
    assert safe_signed_sum(residual) == -1e292
    _refuses(lambda: safe_sum(residual), "the canonical order does overflow here")

    beyond = [-8e307, -7e307, -1.78e308, 5e307, -1e292, 1e308, -MAX_DOUBLE, 1.78e308]
    exact = _exact_sum_oracle(beyond)
    excess = abs(exact) - _EXACT_MAX
    assert excess > 0, "the reproducer must be out of range, or it tests nothing"
    assert excess < Fraction(2) ** 971, (
        "and it must be within one ulp of MAX_DOUBLE, or a coarse check would catch it"
    )
    assert _classify(exact)[0] == "out-of-range"
    _refuses(lambda: safe_signed_sum(beyond), "a total that genuinely exceeds the range")


def test_the_product_reproducers_of_the_round_three_review() -> None:
    """§5.1 and §5.2."""
    beyond = [1e50, MAX_DOUBLE, 1e-150, 1e100]
    exact = _exact_product_oracle(beyond)
    excess = abs(exact) - _EXACT_MAX
    assert 0 < excess < Fraction(2) ** 971, "must be out of range by under one ulp"
    assert _classify(exact)[0] == "out-of-range"
    _refuses(lambda: safe_product(beyond), "a product that genuinely exceeds the range")

    representable = [1e100, 0.5, 1e150, 5e-324, 1e-250]
    assert _classify(_exact_product_oracle(representable)) == ("value", 5e-324)
    assert safe_product(representable) == 5e-324


# ---------------------------------------------------------------------------
# CONVEX ZERO - both directions
# ---------------------------------------------------------------------------
# §13. A statistic of exactly zero is a real answer; a statistic that is
# mathematically non-zero and collapses to zero is not. Tier 1 cannot tell them
# apart, so a non-degenerate zero is classified before it is returned.
_SUBNORMAL = 5e-324


def test_a_mathematically_exact_zero_statistic_is_returned_as_zero() -> None:
    """The over-refusal control: symmetric inputs whose statistic really is zero."""
    assert midpoint(-5.0, 5.0) == 0.0
    assert midpoint(-MAX_DOUBLE, MAX_DOUBLE) == 0.0
    assert midpoint(-_SUBNORMAL, _SUBNORMAL) == 0.0
    assert triangular_mean(-3.0, 0.0, 3.0) == 0.0
    assert triangular_mean(-4.0, 1.0, 3.0) == 0.0
    assert triangular_mean(-2 * _SUBNORMAL, _SUBNORMAL, _SUBNORMAL) == 0.0
    assert beta_pert_mean(-6.0, 0.0, 6.0) == 0.0
    assert beta_pert_mean(-2.0, 0.0, 2.0) == 0.0
    assert beta_pert_mean(-4 * _SUBNORMAL, _SUBNORMAL, 0.0) == 0.0


def test_a_non_zero_statistic_that_collapses_to_zero_is_refused_not_reported() -> None:
    """THE ROUND-3 §8 REPRODUCER. `midpoint(-20s, 19s)` with `s = 5e-324`.

    Tier 1 evaluates `-10s + fl(9.5s)` = `-10s + 10s` = `0` and raises nothing, so
    a rule that accepts any successful tier-1 result reports zero for a statistic
    whose exact value is `-0.5s`. That is the silent deletion §19.3 exists to
    prevent, and it is now a refusal.
    """
    assert -20 * _SUBNORMAL / 2.0 + 19 * _SUBNORMAL / 2.0 == 0.0      # tier 1 succeeds
    exact = Fraction(-20 * _SUBNORMAL) / 2 + Fraction(19 * _SUBNORMAL) / 2
    assert exact != 0 and float(exact) == 0.0
    _refuses(lambda: midpoint(-20 * _SUBNORMAL, 19 * _SUBNORMAL), "a collapsed midpoint")
    _refuses(lambda: triangular_mean(_SUBNORMAL, _SUBNORMAL, -_SUBNORMAL), "collapsed mean")
    _refuses(lambda: beta_pert_mean(_SUBNORMAL, -_SUBNORMAL, _SUBNORMAL), "collapsed mean")


def test_a_statistic_at_the_minimum_non_zero_subnormal_is_produced() -> None:
    """The other side again: an answer that only just exists must still appear."""
    assert midpoint(_SUBNORMAL, _SUBNORMAL) == _SUBNORMAL             # degenerate, tier 0
    assert midpoint(0.0, 2 * _SUBNORMAL) == _SUBNORMAL
    assert midpoint(-2 * _SUBNORMAL, 4 * _SUBNORMAL) == _SUBNORMAL
    assert triangular_mean(_SUBNORMAL, _SUBNORMAL, _SUBNORMAL) == _SUBNORMAL
    assert triangular_mean(0.0, _SUBNORMAL, 2 * _SUBNORMAL) == _SUBNORMAL
    assert beta_pert_mean(0.0, _SUBNORMAL, 2 * _SUBNORMAL) == _SUBNORMAL


def test_the_convex_zero_classification_sweeps_both_directions() -> None:
    """Every small multiple of the subnormal, judged exactly, in BOTH directions.

    Neither may happen:

      * a mathematically non-zero statistic returned as `0.0` — the §8 defect;
      * a representable non-zero statistic refused — the C2 defect.

    Documentation that claimed the second without testing the first is what let §8
    through, so both are counted and both must be zero.

    NOT asserted here: that a NON-ZERO tier-1 result matches the correctly rounded
    exact statistic. §7 locks tier 1's own rounding — `midpoint(-7s, 6s)` gives
    `-1s` where the exact statistic is `-0.5s` — and moving that would change
    ordinary calculations. Only the rescue path is held to exact rounding, which
    `test_the_convex_statistics_are_correctly_rounded_on_the_boundary_corpus` does.
    """
    scale = Fraction(_SUBNORMAL)
    false_zeros: list[tuple] = []
    false_refusals: list[tuple] = []

    for a, b in itertools.product(range(-8, 9), repeat=2):
        exact = (Fraction(a) + Fraction(b)) * scale / 2
        kind, target = _classify(exact)
        try:
            actual = midpoint(a * _SUBNORMAL, b * _SUBNORMAL)
        except NumericalRangeRefusal:
            if kind == "value" and target != 0.0:
                false_refusals.append(("midpoint", a, b, target))
            continue
        if actual == 0.0 and exact != 0:
            false_zeros.append(("midpoint", a, b, exact))

    for a, b, c in itertools.product(range(-4, 5), repeat=3):
        for label, function, exact in (
            ("triangular", triangular_mean, (Fraction(a) + Fraction(b) + Fraction(c)) * scale / 3),
            ("Beta-PERT", beta_pert_mean,
             (Fraction(a) + 4 * Fraction(b) + Fraction(c)) * scale / 6),
        ):
            kind, target = _classify(exact)
            try:
                actual = function(a * _SUBNORMAL, b * _SUBNORMAL, c * _SUBNORMAL)
            except NumericalRangeRefusal:
                if kind == "value" and target != 0.0:
                    false_refusals.append((label, a, b, c, target))
                continue
            if actual == 0.0 and exact != 0:
                false_zeros.append((label, a, b, c, exact))

    assert not false_zeros, f"non-zero statistics reported as zero: {false_zeros[:5]}"
    assert not false_refusals, f"representable statistics refused: {false_refusals[:5]}"


def test_the_zero_classification_actually_fires_on_this_corpus() -> None:
    """The sweep above would pass vacuously if nothing ever reached the
    classification, so count the cases where tier 1 produced a zero that had to be
    judged."""
    judged = 0
    collapsed = 0
    for a, b in itertools.product(range(-8, 9), repeat=2):
        if a == b:
            continue                                   # degenerate; tier 0 answers
        if a * _SUBNORMAL / 2.0 + b * _SUBNORMAL / 2.0 != 0.0:
            continue
        judged += 1
        exact = (Fraction(a) + Fraction(b)) * Fraction(_SUBNORMAL) / 2
        try:
            midpoint(a * _SUBNORMAL, b * _SUBNORMAL)
        except NumericalRangeRefusal:
            collapsed += 1
            assert exact != 0, (a, b)
    assert judged >= 20, f"only {judged} tier-1 zeros in the corpus"
    assert collapsed >= 8, f"only {collapsed} of them were non-zero statistics"


# ---------------------------------------------------------------------------
# COMPOUND EXPRESSIONS - a sum of products, exact end to end
# ---------------------------------------------------------------------------
# `exact_sum_of_products` is the one composition round 5 adds to the kernel. It
# exists because a NAMED Phase-5 output is often a sum of products whose
# individual terms are implementation detail:
#
#     Knom          = SUM_y ( FX * w_y * infl_y )
#     annual base   = SUM_i ( mean * qty * FX * w * infl )
#
# Neither `w_y * infl_y` nor one driver's annual contribution is published
# anywhere, so neither is a representability boundary. Carrying them exactly until
# the named value is reached is what stops a model being refused for its
# evaluation order.
_COMPOUND_CORPUS = (
    MAX_DOUBLE, math.nextafter(MAX_DOUBLE, 0.0), 1e308, 1e250, 1e150, 1e100,
    10.0, 2.0, 1.0, 0.5, 0.1, 1e-100, 1e-250, 1e-300, 1e-320, 5e-324,
)
_COMPOUND_CORPUS = _COMPOUND_CORPUS + tuple(-value for value in _COMPOUND_CORPUS)
_HALF_MIN_SUBNORMAL = Fraction(5e-324) / 2


def _exact_expression_oracle(groups) -> tuple[Fraction, list[Fraction]]:
    terms = [_exact_product_oracle(factors) for factors in groups]
    return sum(terms, Fraction(0)), terms


def _compound_case(stream: _Stream) -> list[list[float]]:
    """One deterministic expression.

    THREE SHAPES, because a purely random corpus of huge magnitudes is almost
    always out of range and would leave the interesting classes unexercised. Two of
    the three shapes are built as cancelling pairs — the same factors with one sign
    flipped — plus a residual, which is exactly how a real model produces terms far
    outside Double range whose sum is not.
    """
    def group() -> list[float]:
        return [stream.pick(_COMPOUND_CORPUS) for _ in range(2 + stream.next(5))]

    shape = stream.next(3)
    if shape == 0:
        return [group() for _ in range(2 + stream.next(19))]

    groups: list[list[float]] = []
    for base in [group() for _ in range(1 + stream.next(4))]:
        groups.append(list(base))
        mirror = list(base)
        mirror[0] = -mirror[0]
        groups.append(mirror)
    for _ in range(stream.next(3) + (1 if shape == 1 else 0)):
        groups.append(group())
    decorated = [(stream.next(1024), index, item) for index, item in enumerate(groups)]
    decorated.sort(key=lambda entry: (entry[0], entry[1]))
    return [item for _, _, item in decorated][:20]


def test_the_compound_rescue_is_faithful_over_ten_thousand_expressions() -> None:
    """§11. 2..20 summed terms, 2..6 factors each, judged against exact rationals.

    The generator is the same fixed LCG the other sweeps use, so the corpus is
    deterministic and does not depend on the stdlib RNG. Every case is a rescue
    case: `exact_sum_of_products` IS the rescue helper and is never the ordinary
    path.
    """
    stream = _Stream(0x5150BE6EB0E268)
    representable = wide_term = subnormal_term = out_of_range = 0

    for _ in range(_ADVERSARIAL_CASES):
        groups = _compound_case(stream)
        exact, terms = _exact_expression_oracle(groups)
        kind, target = _classify(exact)
        if kind == "value":
            representable += 1
            if any(abs(term) > _EXACT_MAX for term in terms):
                wide_term += 1
            if any(term != 0 and abs(term) < _HALF_MIN_SUBNORMAL for term in terms):
                subnormal_term += 1
        elif kind == "out-of-range":
            out_of_range += 1

        try:
            actual = exact_sum_of_products(groups, "expression")
        except NumericalRangeRefusal as error:
            assert kind != "value", f"{groups} refused ({error}) but is {target!r}"
            continue
        assert kind == "value", f"{groups} returned {actual!r} but is {kind}"
        assert actual == target, f"{groups}: got {actual!r}, correctly rounded is {target!r}"

    # The three classes §11 requires the corpus to contain, counted rather than
    # assumed. Without these floors the sweep could pass on easy cases alone.
    assert representable > 2000, representable
    assert wide_term > 500, (
        f"only {wide_term} cases had an individual product outside Double range with a "
        "representable sum — the class the rescue exists for"
    )
    assert subnormal_term > 500, (
        f"only {subnormal_term} cases had an individual product below the smallest "
        "Double with a representable sum"
    )
    assert out_of_range > 1000, out_of_range


def test_a_compound_expression_whose_own_value_is_out_of_range_is_refused() -> None:
    """§10.F, at the helper level. The rescue widens the arithmetic BETWEEN named
    values; it never widens a named value itself."""
    _refuses(
        lambda: exact_sum_of_products([[0.5, 2.0, MAX_DOUBLE], [0.5, 2.0, MAX_DOUBLE]], "Knom"),
        "a compound expression whose exact value exceeds MAX_DOUBLE",
    )
    # ... including by less than one ulp, which no staged evaluation could detect.
    beyond = [[MAX_DOUBLE], [1.0, 2.0**969]]
    exact, _ = _exact_expression_oracle(beyond)
    assert 0 < abs(exact) - _EXACT_MAX < Fraction(2) ** 971
    _refuses(lambda: exact_sum_of_products(beyond, "Knom"), "out of range by under an ulp")


def test_a_compound_expression_that_collapses_to_zero_is_refused() -> None:
    """The other classification, unchanged from the rest of the kernel."""
    exact, _ = _exact_expression_oracle([[5e-324, 0.5]])
    assert exact != 0 and float(exact) == 0.0
    _refuses(lambda: exact_sum_of_products([[5e-324, 0.5]], "annual"), "a collapsed expression")


def test_a_compound_expression_of_ordinary_numbers_is_the_ordinary_answer() -> None:
    """§12 at the helper level: the rescue is not a different arithmetic, it is the
    same arithmetic without the artificial intermediate boundaries."""
    assert exact_sum_of_products([[2.0, 3.0], [4.0, 5.0]], "t") == 26.0
    assert exact_sum_of_products([[1.0, 0.5], [1.0, 0.25]], "t") == 0.75
    assert exact_sum_of_products([], "t") == 0.0
    assert exact_sum_of_products([[1.0, 0.0, 1e308]], "t") == 0.0


def test_a_compound_expression_still_requires_usable_operands() -> None:
    """The widening is between named values, so every operand is still a Double."""
    for bad in (math.inf, math.nan, 10**400):
        _refuses(
            lambda b=bad: exact_sum_of_products([[1.0, b]], "t"),
            f"{bad!r} is not a usable Double",
        )


def test_the_conditioning_magnitude_of_an_unrepresentable_contribution() -> None:
    """§9. `1e-12 * 2 * MAX_DOUBLE` is finite even though `2 * MAX_DOUBLE` is not.

    The unscaled contribution must not have to become a Double just so its C1
    conditioning magnitude can be recorded.
    """
    from pccm_builder.calc_numeric import scaled_magnitude_of_product

    _refuses(lambda: safe_product([2.0, MAX_DOUBLE]), "the unscaled contribution")
    scaled = scaled_magnitude_of_product(0.0, (2.0, 1.0, 1.0, 1.0, MAX_DOUBLE), 1e-12, "t")
    assert scaled == float(Fraction(1e-12) * 2 * Fraction(MAX_DOUBLE)), scaled

    # The accepted conditioning-underflow rule is unchanged: a metadata term too
    # small to move an allowance floored at coefficient * 1 is dropped, not refused.
    assert scaled_magnitude_of_product(0.0, (5e-324, 0.5), 1e-12, "t") == 0.0

    # Overflow of the conditioning magnitude ITSELF is still refused: an allowance
    # outside Double range cannot be compared against.
    _refuses(
        lambda: scaled_magnitude_of_product(0.0, (MAX_DOUBLE, MAX_DOUBLE), 1e-12, "t"),
        "a conditioning magnitude outside Double range",
    )


# ---------------------------------------------------------------------------
# ROUND-3 NEGATIVE CONTROLS
# ---------------------------------------------------------------------------
def test_the_rounded_pair_cancellation_would_fail_the_residual_reproducer() -> None:
    """§17. The round-2 signed-sum rescue, reproduced here, gets §1.1 wrong.

    This is the sabotage written as a test rather than as an edit: if the module
    ever goes back to cancelling with one rounded subtraction, the number below is
    what it will produce, and the reproducer test above will catch it.
    """
    terms = [6e307, -8e307, -1.7e308, 6e307, 7e307, 6e307, -1e292]
    positives = sorted((abs(v), i) for i, v in enumerate(terms) if v > 0.0)
    negatives = sorted((abs(v), i) for i, v in enumerate(terms) if v < 0.0)
    while positives and negatives:
        p_magnitude, p_index = positives.pop()
        n_magnitude, n_index = negatives.pop()
        if p_magnitude == n_magnitude:
            continue
        if p_magnitude > n_magnitude:
            positives.append((p_magnitude - n_magnitude, p_index))
            positives.sort()
        else:
            negatives.append((n_magnitude - p_magnitude, n_index))
            negatives.sort()
    remaining = positives if positives else negatives
    total = 0.0
    for magnitude, _ in remaining:
        total = total + magnitude
    rounded_pair = total if positives else -total

    assert rounded_pair == -1.99792015476736e292, rounded_pair
    assert safe_signed_sum(terms) == -1e292
    assert rounded_pair != safe_signed_sum(terms), (
        "the sabotage must differ from the faithful answer, or it proves nothing"
    )


def test_the_magnitude_balanced_order_would_fail_both_product_reproducers() -> None:
    """§17. The round-2 product rescue, reproduced here, gets §5.1 and §5.2 wrong."""

    def balanced(factors):
        sign = -1.0 if sum(1 for value in factors if value < 0) % 2 else 1.0
        magnitudes = sorted(abs(value) for value in factors)
        low, high = 0, len(magnitudes) - 1
        result = 1.0
        while low <= high:
            if result >= 1.0:
                factor, low = magnitudes[low], low + 1
            else:
                factor, high = magnitudes[high], high - 1
            result = safe_multiply(result, factor)
        return result * sign

    assert balanced([1e50, MAX_DOUBLE, 1e-150, 1e100]) == MAX_DOUBLE
    _refuses(lambda: safe_product([1e50, MAX_DOUBLE, 1e-150, 1e100]), "an out-of-range product")

    _refuses(
        lambda: balanced([1e100, 0.5, 1e150, 5e-324, 1e-250]),
        "the balanced order cannot reach this product",
    )
    assert safe_product([1e100, 0.5, 1e150, 5e-324, 1e-250]) == 5e-324


def test_blindly_accepting_a_tier_one_zero_would_fail_the_convex_reproducer() -> None:
    """§17. The round-2 convex rule, reproduced here, reports zero for §8."""
    blind = -20 * _SUBNORMAL / 2.0 + 19 * _SUBNORMAL / 2.0
    assert blind == 0.0                                 # what tier 1 alone produces
    _refuses(lambda: midpoint(-20 * _SUBNORMAL, 19 * _SUBNORMAL), "a collapsed midpoint")


# ---------------------------------------------------------------------------
# runner
# ---------------------------------------------------------------------------
def _run_all() -> int:
    tests = sorted(
        (name, fn) for name, fn in globals().items()
        if name.startswith("test_") and callable(fn)
    )
    failures = 0
    print("PCCM Phase 5 Gate-A Step-2 numerical primitive tests")
    print("=" * 70)
    for name, fn in tests:
        try:
            fn()
        except AssertionError as error:
            failures += 1
            print(f"  [FAIL] {name}\n         {error}")
        except Exception as error:  # noqa: BLE001
            failures += 1
            print(f"  [ERROR] {name}\n          {type(error).__name__}: {error}")
        else:
            print(f"  [PASS] {name}")
    print("=" * 70)
    print(f"  {len(tests) - failures} passed, {failures} failed")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(_run_all())
