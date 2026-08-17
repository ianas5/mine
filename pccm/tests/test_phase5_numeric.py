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

import math
import sys
from fractions import Fraction
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder.calc_numeric import (  # noqa: E402
    MAX_DOUBLE,
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
    safe_product,
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


def test_the_balanced_order_is_deterministic() -> None:
    """Same inputs, same answer, every time — the property VBA must reproduce."""
    factors = [1e308, 10.0, 0.01]
    assert safe_product(factors) == safe_product(list(factors)) == 1e307


def test_a_zero_factor_makes_an_exactly_zero_product() -> None:
    assert safe_product([0.0, 1e300, 5.0]) == 0.0
    assert safe_product([1e-300, 0.0, 1e-300]) == 0.0


def test_product_signs_are_preserved_through_the_balanced_order() -> None:
    assert safe_product([-1e308, 10.0, 0.01]) == -1e307
    assert safe_product([-1e308, -10.0, 0.01]) == 1e307


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
