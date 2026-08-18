#!/usr/bin/env python3
"""PCCM Phase 5 Gate-A Step-2 tests: the pure analytical oracle.

Golden cases, refusal semantics, referenced-only resolution, reconciliation and
the architecture boundary of `builder/pccm_builder/calc_oracle.py`.

--------------------------------------------------------------------------------
HAND-DERIVED INDEPENDENCE
--------------------------------------------------------------------------------
Every expected value below is written as a literal derived by hand from
docs/phase5_plan.md §23, or by an independent exact-rational calculation in the
test itself. None is obtained by calling another function of `calc_oracle.py`:
the evidence chain is

    hand-derived literals -> Python oracle -> (later) phase5_cases.json -> (later) VBA

and a test that asks the oracle what it produces, then asserts it produces that,
proves nothing.

--------------------------------------------------------------------------------
WHY GOLDEN COMPARISONS USE A TOLERANCE, NOT EXACT EQUALITY
--------------------------------------------------------------------------------
The plan's hand-derived literals are exact MATHEMATICAL values; the oracle
computes in IEEE-754 Double, as the model must. Two effects put the two a unit in
the last place apart on some inputs:

  * ordinary Double rounding: `100 * 1.1085375` is `110.85374999999999`, because
    `1.1085375` has no exact binary representation;
  * the MANDATED stable forms: `Min/6 + ML*(2/3) + Max/6` gives
    `104.99999999999999` where the forbidden naive `(Min + 4*ML + Max)/6` gives
    exactly `105.0`.

Neither is a defect - §19.2 requires the stable form - so the comparison is a
relative tolerance of 1e-12, four orders tighter than the reconciliation
tolerance and about four orders looser than one ulp.

PROOF SCOPE: Linux, Python, pure functions. NO VBA IS EXECUTED HERE.

Runs standalone or under pytest.
"""

from __future__ import annotations

import ast
import dataclasses
import itertools
import subprocess
import sys
import textwrap
from fractions import Fraction
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder.calc_numeric import MAX_DOUBLE, safe_product  # noqa: E402
from pccm_builder.calc_oracle import (  # noqa: E402
    AppliedTimeline,
    CalculationModel,
    CostDriver,
    DriverKind,
    FxRow,
    ModelInputRefusal,
    NumericalRangeRefusal,
    OracleInvariantError,
    RiskDriver,
    Tolerances,
    assert_reconciled,
    calculate,
    reconcile,
)

ORACLE_PATH = PCCM_ROOT / "builder" / "pccm_builder" / "calc_oracle.py"
NUMERIC_PATH = PCCM_ROOT / "builder" / "pccm_builder" / "calc_numeric.py"

# The locked tolerance constants of spec/calc_contract.yaml, supplied at the
# adapter boundary. The pure numerical layer never reads a file.
TOL = Tolerances(
    profiling_sum_absolute=1e-9,
    identity_absolute_floor=1e-6,
    identity_relative_coefficient=1e-12,
    conditioning_scale_floor=1.0,
)

REL = 1e-12


def _close(actual: float, expected: float, rel: float = REL, note: str = "") -> None:
    allowance = abs(expected) * rel + 1e-9
    assert abs(actual - expected) <= allowance, (
        f"{note or 'value'}: got {actual!r}, expected {expected!r} "
        f"(difference {actual - expected!r} exceeds {allowance!r})"
    )


def _refuses(call, reason: str) -> str:
    try:
        call()
    except (ModelInputRefusal, NumericalRangeRefusal) as error:
        return str(error)
    raise AssertionError(f"{reason}: no refusal was raised")


# ---------------------------------------------------------------------------
# fixtures - the shared setup of plan §23
# ---------------------------------------------------------------------------
def _cost(
    permanent_id: str = "CL-001",
    distribution: str = "Triangular",
    currency: str = "SAR",
    profile: str = "Standard",
    minimum: object = 80,
    most_likely: object = 100,
    maximum: object = 150,
    weights: tuple[object, ...] = (1.0,),
    quantity: object = 10,
) -> CostDriver:
    return CostDriver(
        permanent_id, distribution, currency, profile,
        minimum, most_likely, maximum, weights, quantity=quantity,
    )


def _risk(
    permanent_id: str = "R-001",
    distribution: str = "Triangular",
    currency: str = "SAR",
    profile: str = "Standard",
    minimum: object = 100,
    most_likely: object = 200,
    maximum: object = 450,
    weights: tuple[object, ...] = (1.0,),
    probability: object = 0.3,
) -> RiskDriver:
    return RiskDriver(
        permanent_id, distribution, currency, profile,
        minimum, most_likely, maximum, weights, probability=probability,
    )


def _model(
    base: int = 2026,
    start: int = 2026,
    duration: int = 1,
    discount: object = 0.10,
    fx: tuple[FxRow, ...] = (FxRow("SAR", 1),),
    rates: dict[str, dict[int, object]] | None = None,
    costs: tuple[CostDriver, ...] = (),
    risks: tuple[RiskDriver, ...] = (),
) -> CalculationModel:
    return CalculationModel(
        timeline=AppliedTimeline(base, start, duration),
        discount_rate=discount,
        fx_rows=fx,
        inflation_rates=rates if rates is not None else {"Standard": {}},
        cost_drivers=costs,
        risk_drivers=risks,
    )


def _three_year(rate: object = 0.05) -> dict[str, dict[int, object]]:
    return {"Standard": {2027: rate, 2028: rate, 2029: rate}}


# ---------------------------------------------------------------------------
# CASES 1-13 - the hand-derived arithmetic
# ---------------------------------------------------------------------------
def test_case_01_sar_no_inflation_one_project_year() -> None:
    """Base 2026, Start 2026, Dur 1, r = 10%. infl = 1, disc = 1, Knom = Kpv = 1."""
    result = calculate(_model(costs=(_cost(),)), TOL)
    driver = result.drivers[0]
    _close(driver.knom, 1.0, note="Knom")
    _close(driver.kpv, 1.0, note="Kpv")
    _close(result.totals.a_nom, 1000.0, note="A_nom")
    _close(result.totals.c_nom, 1100.0, note="C_nom")
    _close(result.totals.b_nom, 100.0, note="B_nom")
    assert_reconciled(result, TOL)


def test_case_02_foreign_currency() -> None:
    """USD FX = 3.75, unit cost 100, Qty 4 -> Knom = 3.75, A = 1500 SAR."""
    result = calculate(
        _model(
            fx=(FxRow("SAR", 1), FxRow("USD", 3.75)),
            costs=(_cost(currency="USD", quantity=4),),
        ),
        TOL,
    )
    _close(result.drivers[0].knom, 3.75, note="Knom")
    _close(result.totals.a_nom, 1500.0, note="A_nom")


def test_case_03_multi_year_profiling_with_compounded_inflation() -> None:
    """Base 2026, Start 2027, Dur 3, rates 5%, profile 20/50/30.

    f = 1.05, 1.1025, 1.157625
    Knom = 0.21 + 0.55125 + 0.3472875 = 1.1085375
    A_nom = 1108.5375
    """
    result = calculate(
        _model(base=2026, start=2027, duration=3, rates=_three_year(),
               costs=(_cost(weights=(0.2, 0.5, 0.3)),)),
        TOL,
    )
    factors = {row.calendar_year: row.cumulative_factor for row in result.inflation_factors}
    _close(factors[2027], 1.05, note="infl 2027")
    _close(factors[2028], 1.1025, note="infl 2028")
    _close(factors[2029], 1.157625, note="infl 2029")
    _close(result.drivers[0].knom, 1.1085375, note="Knom")
    _close(result.totals.a_nom, 1108.5375, note="A_nom")


def test_case_04_present_value_across_multiple_years() -> None:
    """r = 10% on case 3.

    Kpv  = 0.21 + 0.501136363636 + 0.287014462810 = 0.998150826446
    A_pv = 998.150826446
    C_pv = 1097.965909091
    """
    result = calculate(
        _model(base=2026, start=2027, duration=3, rates=_three_year(),
               costs=(_cost(weights=(0.2, 0.5, 0.3)),)),
        TOL,
    )
    _close(result.drivers[0].kpv, 0.998150826446, rel=1e-11, note="Kpv")
    _close(result.totals.a_pv, 998.150826446, rel=1e-11, note="A_pv")
    _close(result.totals.c_pv, 1097.965909091, rel=1e-11, note="C_pv")


def test_case_05_triangular_deterministic_versus_mean() -> None:
    """80/100/150 -> central 100, mean 110. On case 3: C_nom 1219.39125, B_nom 110.85375."""
    result = calculate(
        _model(base=2026, start=2027, duration=3, rates=_three_year(),
               costs=(_cost(weights=(0.2, 0.5, 0.3)),)),
        TOL,
    )
    driver = result.drivers[0]
    _close(driver.central_value, 100.0, note="central")
    _close(driver.mean_value, 110.0, note="mean")
    _close(result.totals.c_nom, 1219.39125, note="C_nom")
    _close(result.totals.b_nom, 110.85375, note="B_nom")
    assert driver.central_basis == "ML"


def test_case_06_beta_pert_deterministic_versus_mean() -> None:
    """80/100/150, lambda 4 -> central 100, mean (80 + 400 + 150)/6 = 105."""
    result = calculate(_model(costs=(_cost(distribution="Beta-PERT"),)), TOL)
    driver = result.drivers[0]
    _close(driver.central_value, 100.0, note="central")
    _close(driver.mean_value, 105.0, note="mean")
    assert driver.central_basis == "ML"


def test_case_07_uniform_midpoint_equals_mean() -> None:
    """Min 80 / Max 150 -> central 115, mean 115, and B = 0 for this driver."""
    result = calculate(
        _model(costs=(_cost(distribution="Uniform", most_likely=None),)), TOL
    )
    driver = result.drivers[0]
    _close(driver.central_value, 115.0, note="central")
    _close(driver.mean_value, 115.0, note="mean")
    assert result.totals.b_nom == 0.0
    assert driver.central_basis == "Midpoint"


def test_case_08_risk_emv_with_probability_below_one() -> None:
    """P = 30%, severity 100/200/450 -> mean severity 250, D = 75."""
    result = calculate(_model(risks=(_risk(),)), TOL)
    _close(result.drivers[0].mean_value, 250.0, note="expected severity")
    _close(result.totals.d_nom, 75.0, note="D_nom")


def test_case_09_multi_year_risk_profile() -> None:
    """The case-8 risk on case-3 factors: D_nom 83.1403125, D_pv 74.8613119835."""
    result = calculate(
        _model(base=2026, start=2027, duration=3, rates=_three_year(),
               risks=(_risk(weights=(0.2, 0.5, 0.3)),)),
        TOL,
    )
    _close(result.totals.d_nom, 83.1403125, note="D_nom")
    _close(result.totals.d_pv, 74.8613119835, rel=1e-11, note="D_pv")


def test_case_10_base_year_equals_start_year() -> None:
    """Base 2027, Start 2027, Dur 2, rate 2028 = 5% -> infl(2027) = 1, infl(2028) = 1.05."""
    result = calculate(
        _model(base=2027, start=2027, duration=2, discount=0.0,
               rates={"Standard": {2028: 0.05}},
               costs=(_cost(weights=(0.5, 0.5)),)),
        TOL,
    )
    factors = {row.calendar_year: row.cumulative_factor for row in result.inflation_factors}
    assert factors[2027] == 1.0
    _close(factors[2028], 1.05, note="infl 2028")


def test_case_11_base_year_earlier_than_start_year() -> None:
    """Case 3: project year 1 (2027) already carries 1.05 from pre-project escalation."""
    result = calculate(
        _model(base=2026, start=2027, duration=3, rates=_three_year(),
               costs=(_cost(weights=(1.0, 0.0, 0.0)),)),
        TOL,
    )
    factors = {row.calendar_year: row.cumulative_factor for row in result.inflation_factors}
    _close(factors[2027], 1.05, note="project year 1 factor")
    _close(result.drivers[0].knom, 1.05, note="Knom is the year-1 factor alone")


def test_case_12_zero_inflation_leaves_knom_at_fx() -> None:
    result = calculate(
        _model(base=2026, start=2027, duration=3, discount=0.0,
               fx=(FxRow("SAR", 1), FxRow("USD", 3.75)),
               rates=_three_year(0.0),
               costs=(_cost(currency="USD", weights=(0.2, 0.5, 0.3)),)),
        TOL,
    )
    assert all(row.cumulative_factor == 1.0 for row in result.inflation_factors)
    _close(result.drivers[0].knom, 3.75, note="Knom = FX when inflation is flat")


def test_case_13_negative_but_valid_inflation() -> None:
    """rate -2%, three years -> 0.98, 0.9604, 0.941192 (D2)."""
    result = calculate(
        _model(base=2026, start=2027, duration=3, discount=0.0, rates=_three_year(-0.02),
               costs=(_cost(weights=(0.2, 0.5, 0.3)),)),
        TOL,
    )
    factors = {row.calendar_year: row.cumulative_factor for row in result.inflation_factors}
    _close(factors[2027], 0.98, note="2027")
    _close(factors[2028], 0.9604, note="2028")
    _close(factors[2029], 0.941192, note="2029")


# ---------------------------------------------------------------------------
# CASES 14-25 - refusals and the acceptances a naive implementation over-blocks
# ---------------------------------------------------------------------------
def test_case_14_blank_required_inflation_rate_is_refused_naming_profile_and_year() -> None:
    message = _refuses(
        lambda: calculate(
            _model(base=2026, start=2027, duration=3,
                   rates={"Standard": {2027: 0.05, 2028: None, 2029: 0.05}},
                   costs=(_cost(weights=(0.2, 0.5, 0.3)),)),
            TOL,
        ),
        "blank required inflation rate",
    )
    assert "Standard" in message
    assert "2028" in message


def test_case_15_profile_sum_not_one_hundred_percent_is_refused_naming_the_driver() -> None:
    message = _refuses(
        lambda: calculate(
            _model(base=2026, start=2027, duration=3, rates=_three_year(),
                   costs=(_cost(weights=(0.2, 0.5, 0.2)),)),
            TOL,
        ),
        "profile sum != 100%",
    )
    assert "CL-001" in message
    assert "0.8999999999999999" in message      # the offending sum is reported verbatim


def test_case_16_quantity_zero_is_refused() -> None:
    message = _refuses(
        lambda: calculate(_model(costs=(_cost(quantity=0),)), TOL), "Quantity = 0"
    )
    assert "CL-001" in message and "Quantity" in message


def test_case_17_negative_quantity_is_refused() -> None:
    _refuses(lambda: calculate(_model(costs=(_cost(quantity=-5),)), TOL), "Quantity < 0")


def test_case_18_discount_rate_of_minus_one_hundred_percent_is_refused() -> None:
    message = _refuses(
        lambda: calculate(_model(discount=-1.0, costs=(_cost(),)), TOL), "r = -100%"
    )
    assert "discount rate" in message


def test_case_19_negative_discount_rate_above_minus_one_hundred_percent_is_accepted() -> None:
    """r = -5% on case 3: disc = 1, 1/0.95, 1/0.9025, so A_pv > A_nom - correct,
    and exactly why `A_pv <= A_nom` is not a gate."""
    result = calculate(
        _model(base=2026, start=2027, duration=3, discount=-0.05, rates=_three_year(),
               costs=(_cost(weights=(0.2, 0.5, 0.3)),)),
        TOL,
    )
    _close(result.discount_factors[1], 1.0, note="disc 1")
    _close(result.discount_factors[2], 1.0 / 0.95, note="disc 2")
    _close(result.discount_factors[3], 1.0 / 0.9025, note="disc 3")
    assert result.totals.a_pv > result.totals.a_nom
    assert_reconciled(result, TOL)


def test_case_20_inflation_rate_of_minus_one_hundred_percent_is_refused() -> None:
    message = _refuses(
        lambda: calculate(
            _model(base=2026, start=2027, duration=3,
                   rates={"Standard": {2027: 0.05, 2028: -1.0, 2029: 0.05}},
                   costs=(_cost(weights=(0.2, 0.5, 0.3)),)),
            TOL,
        ),
        "inflation rate = -100%",
    )
    assert "2028" in message


def test_case_21_negative_but_valid_inflation_rate_is_accepted() -> None:
    result = calculate(
        _model(base=2026, start=2027, duration=3, discount=0.0, rates=_three_year(-0.02),
               costs=(_cost(weights=(0.2, 0.5, 0.3)),)),
        TOL,
    )
    assert result.totals.a_nom > 0.0
    assert_reconciled(result, TOL)


def test_case_22_uniform_with_a_populated_most_likely_is_accepted_and_ignores_it() -> None:
    """D1: Min 80 / ML 999 / Max 150 -> central = mean = 115, ML ignored."""
    result = calculate(
        _model(costs=(_cost(distribution="Uniform", most_likely=999),)), TOL
    )
    driver = result.drivers[0]
    _close(driver.central_value, 115.0, note="central")
    _close(driver.mean_value, 115.0, note="mean")


def test_case_23_a_hundred_percent_summing_profile_with_a_blank_is_refused() -> None:
    """D4: 50% / blank / 50% sums to 100% and is STILL refused."""
    message = _refuses(
        lambda: calculate(
            _model(base=2026, start=2027, duration=3, rates=_three_year(),
                   costs=(_cost(weights=(0.5, None, 0.5)),)),
            TOL,
        ),
        "blank profiling cell",
    )
    assert "CL-001" in message
    assert "project year 2" in message
    assert "blank" in message


def test_case_24_double_overflow_is_a_controlled_refusal() -> None:
    """Never an uncontrolled error, never a fabricated zero."""
    message = _refuses(
        lambda: calculate(
            _model(base=2026, start=2027, duration=3, discount=0.0,
                   rates=_three_year(1e300),
                   costs=(_cost(weights=(0.2, 0.5, 0.3)),)),
            TOL,
        ),
        "inflation overflow",
    )
    assert "Standard" in message

    extreme = _refuses(
        lambda: calculate(_model(costs=(_cost(minimum=1e308, most_likely=1e308,
                                              maximum=1e308, quantity=1e10),)), TOL),
        "extreme unit cost",
    )
    assert "CL-001" in extreme


def test_case_25_an_unreferenced_incomplete_fx_row_does_not_block() -> None:
    """A valid SAR-only model plus a duplicate, blank-rate EUR row referenced by
    nothing. Validating the whole Config universe would refuse a valid model."""
    result = calculate(
        _model(
            fx=(FxRow("SAR", 1), FxRow("EUR", None), FxRow("EUR", "n/a")),
            rates={"Standard": {}, "Unused": {2027: None}},
            costs=(_cost(),),
        ),
        TOL,
    )
    _close(result.totals.a_nom, 1000.0, note="A_nom")
    assert "EUR" not in result.resolved_fx


# ---------------------------------------------------------------------------
# CASES 28-31
# ---------------------------------------------------------------------------
def test_case_28_stable_means_accept_inputs_the_naive_form_cannot() -> None:
    for distribution, expected in (("Triangular", 1e308), ("Beta-PERT", 1e308)):
        result = calculate(
            _model(costs=(_cost(distribution=distribution, minimum=1e308,
                                most_likely=1e308, maximum=1e308, quantity=1),)),
            TOL,
        )
        _close(result.drivers[0].mean_value, expected, note=f"{distribution} mean")

    uniform = calculate(
        _model(costs=(_cost(distribution="Uniform", minimum=1.5e308,
                            most_likely=None, maximum=1.5e308, quantity=1),)),
        TOL,
    )
    _close(uniform.drivers[0].mean_value, 1.5e308, note="Uniform midpoint")


def test_case_29_discount_factor_underflow_refuses_at_project_year_34() -> None:
    message = _refuses(
        lambda: calculate(
            _model(base=2026, start=2026, duration=34, discount=1e10,
                   rates={"Standard": {y: 0.0 for y in range(2027, 2060)}},
                   costs=(_cost(weights=tuple([1.0] + [0.0] * 33)),)),
            TOL,
        ),
        "discount underflow",
    )
    assert "project year 34" in message

    # 33 years is the last representable duration for this rate.
    ok = calculate(
        _model(base=2026, start=2026, duration=33, discount=1e10,
               rates={"Standard": {y: 0.0 for y in range(2027, 2060)}},
               costs=(_cost(weights=tuple([1.0] + [0.0] * 32)),)),
        TOL,
    )
    assert ok.discount_factors[33] > 0.0


def test_case_30_cancellation_heavy_reconciliation_holds() -> None:
    """Large positive and negative contributions whose net is near zero.

    Cancellation ACROSS YEARS is where the plan's conditioning design does its
    work: `Sum_y |annual| + |headline|` keeps growing when the signed annual
    values cancel, so I3/I4 keep a tolerance proportional to the arithmetic
    actually performed rather than falling back to the 1e-6 floor.
    """
    costs = (
        _cost("CL-001", minimum=1e9, most_likely=1.1e9, maximum=1.2e9, weights=(1.0, 0.0),
              quantity=1),
        _cost("CL-002", minimum=-1.2e9, most_likely=-1.1e9, maximum=-1e9, weights=(0.0, 1.0),
              quantity=1),
    )
    result = calculate(
        _model(base=2026, start=2026, duration=2, discount=0.0,
               rates={"Standard": {2027: 0.0}}, costs=costs),
        TOL,
    )
    assert result.totals.c_nom == 0.0                       # the net really is zero
    assert abs(result.annual[0].base_cost_nominal) > 1e9     # the arithmetic was not
    assert abs(result.annual[1].base_cost_nominal) > 1e9

    checks = {check.name: check for check in reconcile(result, TOL)}
    assert checks["I3a nominal base"].allowance > TOL.identity_absolute_floor, (
        "the annual conditioning scale must reflect the terms, not the near-zero net"
    )
    assert_reconciled(result, TOL)


def test_erratum_c1_headline_cross_driver_cancellation_reconciles() -> None:
    """ERRATUM C1 regression — formerly a PINNED FALSE FAILURE.

    Step 2 originally reported this valid model as failing I1, because the locked
    conditioning scale was `max(1, |A| + |B| + |C|)` — the HEADLINE TOTALS, which
    are already-cancelled numbers. Two of the three drivers are exact mirrors, so
    A, B and C collapse to a few tens of SAR while the accumulation ran through
    partial sums of `1e17`, where one ulp is already 16 SAR.

    Review corrected the plan: the scale now sums the UNDERLYING PER-DRIVER
    contributions. The same model must now calculate and reconcile.
    """
    costs = (
        _cost("CL-001", minimum=0.0, most_likely=1e17, maximum=4e17, weights=(1.0,), quantity=1),
        _cost("CL-002", minimum=10.0, most_likely=30.0, maximum=110.0, weights=(1.0,),
              quantity=1),
        _cost("CL-003", minimum=-4e17, most_likely=-1e17, maximum=0.0, weights=(1.0,),
              quantity=1),
    )
    result = calculate(_model(discount=0.0, costs=costs), TOL)

    # The headline totals really have cancelled to almost nothing...
    assert abs(result.totals.a_nom) < 1e3 and abs(result.totals.c_nom) < 1e3
    checks = {check.name: check for check in reconcile(result, TOL)}
    i1 = checks["I1 nominal: A + B = C"]
    assert abs(i1.difference) > 1.0, "the rounding residue is real, not zero"

    # ...but the conditioning scale reflects the 1e17 arithmetic that produced them.
    assert i1.allowance > abs(i1.difference)
    assert i1.allowance > TOL.identity_absolute_floor
    assert i1.holds
    assert_reconciled(result, TOL)


def test_erratum_c1_annual_within_year_cancellation_reconciles() -> None:
    """ERRATUM C1 regression — the annual half, independently reproduced by review.

    The former annual scale summed `|annual aggregate_y|`, which is also already
    cancelled: within year 1 the `+1e16` and `-1e16` contributions annihilate, so
    the aggregate is `0` and the whole scale collapses to `1` — even though the
    annual arithmetic processed about `2e16`. I3a was reported as failing by 1 SAR.
    """
    costs = (
        _cost("CL-001", distribution="Uniform", minimum=1e16, most_likely=None, maximum=1e16,
              weights=(1.0, 0.0), quantity=1),
        _cost("CL-002", distribution="Uniform", minimum=1.0, most_likely=None, maximum=1.0,
              weights=(0.0, 1.0), quantity=1),
        _cost("CL-003", distribution="Uniform", minimum=-1e16, most_likely=None, maximum=-1e16,
              weights=(1.0, 0.0), quantity=1),
    )
    result = calculate(
        _model(base=2026, start=2026, duration=2, discount=0.0,
               rates={"Standard": {2027: 0.0}}, costs=costs),
        TOL,
    )
    assert result.totals.c_nom == 0.0
    assert [row.base_cost_nominal for row in result.annual] == [0.0, 1.0]

    checks = {check.name: check for check in reconcile(result, TOL)}
    i3a = checks["I3a nominal base"]
    assert abs(i3a.difference) == 1.0, "the 1 SAR residue is the reported one"
    assert i3a.allowance > 1.0, (
        "the annual scale must reflect the per-driver-per-year contributions, "
        "not the cancelled row aggregate"
    )
    assert i3a.holds
    assert_reconciled(result, TOL)


def test_erratum_c1_the_conditioning_scale_is_captured_during_accumulation() -> None:
    """The magnitudes describe the calculation that actually ran.

    They are captured alongside each contribution, so `reconcile` cannot be
    checking one calculation against another calculation's scale.
    """
    result = calculate(
        _model(base=2026, start=2027, duration=3, rates=_three_year(),
               costs=(_cost("CL-001", weights=(0.2, 0.5, 0.3)),),
               risks=(_risk("R-001", weights=(0.2, 0.5, 0.3)),)),
        TOL,
    )
    magnitudes = result.magnitudes
    assert magnitudes.relative_coefficient == TOL.identity_relative_coefficient
    # Every scaled magnitude is the coefficient times a positive contribution sum.
    for field_name in ("a_nom", "c_nom", "d_nom", "e_nom", "annual_base_nom",
                       "annual_risk_nom", "annual_total_nom"):
        assert getattr(magnitudes, field_name) > 0.0, field_name
    # E's magnitude covers cost AND risk contributions, so it exceeds either alone.
    assert magnitudes.e_nom > magnitudes.c_nom
    assert magnitudes.e_nom > magnitudes.d_nom
    assert magnitudes.annual_total_nom > magnitudes.annual_base_nom


def test_reconciliation_refuses_magnitudes_captured_at_a_different_coefficient() -> None:
    """A guard against exactly the mistake the capture design prevents."""
    result = calculate(_model(costs=(_cost(),)), TOL)
    other = Tolerances(
        profiling_sum_absolute=1e-9,
        identity_absolute_floor=1e-6,
        identity_relative_coefficient=1e-9,      # different from the capture
        conditioning_scale_floor=1.0,
    )
    try:
        reconcile(result, other)
    except OracleInvariantError as error:
        assert "coefficient" in str(error)
        return
    raise AssertionError("a mismatched conditioning coefficient was accepted")


def test_case_31_the_base_year_inflation_row_is_explicit() -> None:
    """Base 2026, Start 2028, Dur 3 -> rows 2026..2030, with 2026 blank at factor 1."""
    rates = {"Standard": {year: 0.05 for year in range(2027, 2031)}}
    result = calculate(
        _model(base=2026, start=2028, duration=3, rates=rates,
               costs=(_cost(weights=(0.2, 0.5, 0.3)),)),
        TOL,
    )
    rows = {row.calendar_year: row for row in result.inflation_factors}
    assert sorted(rows) == [2026, 2027, 2028, 2029, 2030]
    assert rows[2026].annual_rate is None, "the base-year rate is a model-controlled blank"
    assert rows[2026].cumulative_factor == 1.0
    _close(rows[2027].cumulative_factor, 1.05, note="pre-project 2027")
    _close(rows[2028].cumulative_factor, 1.1025, note="project year 1")


# ---------------------------------------------------------------------------
# REFERENCED-ONLY RESOLUTION
# ---------------------------------------------------------------------------
def test_a_referenced_bad_currency_is_refused() -> None:
    for bad in (None, "n/a", 0, -1):
        _refuses(
            lambda b=bad: calculate(
                _model(fx=(FxRow("SAR", 1), FxRow("USD", b)),
                       costs=(_cost(currency="USD"),)),
                TOL,
            ),
            f"referenced USD rate {bad!r}",
        )


def test_a_referenced_missing_currency_is_refused() -> None:
    message = _refuses(
        lambda: calculate(_model(costs=(_cost(currency="USD"),)), TOL), "missing USD"
    )
    assert "USD" in message


def test_a_referenced_duplicate_currency_is_refused() -> None:
    message = _refuses(
        lambda: calculate(
            _model(fx=(FxRow("SAR", 1), FxRow("USD", 3.75), FxRow("USD", 3.80)),
                   costs=(_cost(currency="USD"),)),
            TOL,
        ),
        "duplicate USD",
    )
    assert "USD" in message


def test_an_unreferenced_bad_currency_does_not_block() -> None:
    result = calculate(
        _model(fx=(FxRow("SAR", 1), FxRow("USD", -3.75), FxRow("USD", None)),
               costs=(_cost(currency="SAR"),)),
        TOL,
    )
    _close(result.totals.a_nom, 1000.0, note="A_nom")


def test_a_referenced_incomplete_inflation_profile_is_refused() -> None:
    message = _refuses(
        lambda: calculate(
            _model(base=2026, start=2027, duration=3,
                   rates={"Standard": {2027: 0.05, 2029: 0.05}},
                   costs=(_cost(weights=(0.2, 0.5, 0.3)),)),
            TOL,
        ),
        "missing required inflation year",
    )
    assert "2028" in message


def test_a_referenced_missing_inflation_profile_is_refused() -> None:
    message = _refuses(
        lambda: calculate(_model(rates={}, costs=(_cost(),)), TOL), "missing profile"
    )
    assert "Standard" in message


def test_an_unreferenced_incomplete_inflation_profile_does_not_block() -> None:
    result = calculate(
        _model(base=2026, start=2027, duration=3,
               rates={"Standard": {2027: 0.05, 2028: 0.05, 2029: 0.05},
                      "Neglected": {2027: None}},
               costs=(_cost(weights=(0.2, 0.5, 0.3)),)),
        TOL,
    )
    _close(result.totals.a_nom, 1108.5375, note="A_nom")


def test_the_sar_invariant_applies_even_with_no_foreign_currency() -> None:
    for fx in ((FxRow("SAR", 2),), (FxRow("SAR", None),), (), (FxRow("SAR", 1), FxRow("SAR", 1))):
        message = _refuses(
            lambda f=fx: calculate(_model(fx=f, costs=(_cost(),)), TOL),
            f"SAR invariant with fx={fx}",
        )
        assert "SAR" in message


def test_the_sar_invariant_applies_to_a_model_with_no_drivers_at_all() -> None:
    _refuses(lambda: calculate(_model(fx=(FxRow("SAR", 2),)), TOL), "empty model, SAR != 1")


# ---------------------------------------------------------------------------
# §22 MUST-HAVE TESTS
# ---------------------------------------------------------------------------
def test_a_populated_uniform_most_likely_has_no_influence_whatsoever() -> None:
    results = [
        calculate(_model(costs=(_cost(distribution="Uniform", most_likely=ml),)), TOL)
        for ml in (None, 0, 999, -1e6, 115)
    ]
    first = results[0].totals
    for other in results[1:]:
        assert other.totals == first


def test_a_numeric_zero_profile_weight_is_accepted() -> None:
    result = calculate(
        _model(base=2026, start=2027, duration=3, rates=_three_year(),
               costs=(_cost(weights=(0.5, 0.0, 0.5)),)),
        TOL,
    )
    assert result.drivers[0].weights == (0.5, 0.0, 0.5)
    assert_reconciled(result, TOL)


def test_a_blank_profile_weight_is_refused_and_is_not_zero() -> None:
    message = _refuses(
        lambda: calculate(
            _model(base=2026, start=2027, duration=3, rates=_three_year(),
                   costs=(_cost(weights=(0.5, None, 0.5)),)),
            TOL,
        ),
        "blank weight",
    )
    assert "blank" in message and "not zero" in message


def test_profile_weights_travel_with_the_permanent_id_not_the_row() -> None:
    """Reordering the driver sequence changes nothing at all.

    Retained, but no longer sufficient on its own: these values are tame enough
    that reordering the accumulation could not have changed the answer anyway.
    The adversarial fixture below is the real test.
    """
    a = _cost("CL-001", weights=(0.2, 0.5, 0.3), quantity=10)
    b = _cost("CL-002", weights=(0.6, 0.1, 0.3), quantity=4, maximum=200)
    forward = calculate(
        _model(base=2026, start=2027, duration=3, rates=_three_year(), costs=(a, b)), TOL
    )
    reversed_ = calculate(
        _model(base=2026, start=2027, duration=3, rates=_three_year(), costs=(b, a)), TOL
    )
    assert forward.totals == reversed_.totals
    assert {d.permanent_id: d.knom for d in forward.drivers} == {
        d.permanent_id: d.knom for d in reversed_.drivers
    }


# The adversarial fixture: three exactly-representable values whose ACCUMULATION
# ORDER decides the answer. `1e16 + 1 - 1e16` is `0`, because 1 is below the ulp
# of 1e16; `1e16 - 1e16 + 1` is `1`. Row order is excluded from the calculation
# fingerprint, so two workbooks with the SAME fingerprint must not be able to
# disagree about which of those is the answer.
_ADVERSARIAL = (
    _cost("CL-001", distribution="Uniform", minimum=1e16, most_likely=None, maximum=1e16,
          weights=(1.0,), quantity=1),
    _cost("CL-002", distribution="Uniform", minimum=1.0, most_likely=None, maximum=1.0,
          weights=(1.0,), quantity=1),
    _cost("CL-003", distribution="Uniform", minimum=-1e16, most_likely=None, maximum=-1e16,
          weights=(1.0,), quantity=1),
)


def test_the_adversarial_fixture_really_is_order_sensitive_when_summed_naively() -> None:
    """Guards the guard: if this fixture were not order-sensitive, the permutation
    test below would prove nothing."""
    assert (1e16 + 1.0) - 1e16 == 0.0
    assert (1e16 - 1e16) + 1.0 == 1.0


def test_row_order_cannot_change_any_result_under_all_six_permutations() -> None:
    """CANONICAL COMPUTATIONAL ORDER — the complete result, not just the totals.

    Before canonical ordering, order `(001, 002, 003)` gave `A = C = E = 0` and
    order `(001, 003, 002)` gave `A = C = E = 1` for these same three drivers with
    the same permanent IDs, values and profiles.
    """
    results = []
    for permutation in itertools.permutations(_ADVERSARIAL):
        results.append(calculate(_model(discount=0.0, costs=permutation), TOL))

    assert len(results) == 6
    first = results[0]
    for other in results[1:]:
        assert other.totals == first.totals
        assert other.annual == first.annual
        assert other.drivers == first.drivers
        assert other.inflation_factors == first.inflation_factors
        assert other.resolved_fx == first.resolved_fx
        assert other.discount_factors == first.discount_factors
        assert other.magnitudes == first.magnitudes
        assert other == first, "the complete CalculationResult must be identical"


def test_the_canonical_order_is_ascending_permanent_id() -> None:
    """Audit output order is canonical too, not the order rows arrived in."""
    for permutation in itertools.permutations(_ADVERSARIAL):
        result = calculate(_model(discount=0.0, costs=permutation), TOL)
        assert [d.permanent_id for d in result.drivers] == ["CL-001", "CL-002", "CL-003"]


def test_the_canonical_order_uses_utf16_ordinal_comparison() -> None:
    """The same comparison the fingerprint uses, so the two cannot drift apart."""
    from pccm_builder.calc_fingerprint import utf16_sort_key
    from pccm_builder.calc_oracle import canonical_order

    astral, private_use = "\U00010000", "\ue000"
    drivers = [_cost(private_use), _cost(astral)]
    assert [d.permanent_id for d in canonical_order(drivers)] == [astral, private_use]
    assert sorted([private_use, astral]) == [private_use, astral]      # Python disagrees
    assert sorted([private_use, astral], key=utf16_sort_key) == [astral, private_use]


def test_mixed_cost_and_risk_models_are_canonically_ordered_too() -> None:
    costs = (_cost("CL-003"), _cost("CL-001"), _cost("CL-002"))
    risks = (_risk("R-002"), _risk("R-001"))
    result = calculate(_model(costs=costs, risks=risks), TOL)
    assert [d.permanent_id for d in result.drivers] == [
        "CL-001", "CL-002", "CL-003", "R-001", "R-002"
    ]


def test_the_inflation_audit_order_does_not_follow_row_order() -> None:
    """Reference-set discovery is observable through the audit rows, so it is
    canonical rather than first-driver-wins."""
    rates = {"Alpha": {2027: 0.05}, "Zulu": {2027: 0.05}}
    forward = calculate(
        _model(base=2026, start=2027, duration=1, rates=rates,
               costs=(_cost("CL-001", profile="Zulu"), _cost("CL-002", profile="Alpha"))),
        TOL,
    )
    reversed_ = calculate(
        _model(base=2026, start=2027, duration=1, rates=rates,
               costs=(_cost("CL-002", profile="Alpha"), _cost("CL-001", profile="Zulu"))),
        TOL,
    )
    assert forward.inflation_factors == reversed_.inflation_factors
    assert [row.profile for row in forward.inflation_factors][0] == "Alpha"


def test_a_wrong_length_profile_is_refused() -> None:
    message = _refuses(
        lambda: calculate(
            _model(base=2026, start=2027, duration=3, rates=_three_year(),
                   costs=(_cost(weights=(0.5, 0.5)),)),
            TOL,
        ),
        "wrong profile length",
    )
    assert "CL-001" in message


def test_a_non_numeric_profile_cell_is_refused() -> None:
    _refuses(
        lambda: calculate(
            _model(base=2026, start=2027, duration=3, rates=_three_year(),
                   costs=(_cost(weights=(0.5, "half", 0.0)),)),
            TOL,
        ),
        "non-numeric weight",
    )


def test_quantity_scales_the_contribution_but_never_knom_or_kpv() -> None:
    """Quantity is a per-driver multiplier, not part of the escalation path."""
    base = calculate(_model(costs=(_cost(quantity=10),)), TOL)
    doubled = calculate(_model(costs=(_cost(quantity=20),)), TOL)
    assert doubled.drivers[0].knom == base.drivers[0].knom
    assert doubled.drivers[0].kpv == base.drivers[0].kpv
    _close(doubled.totals.a_nom, base.totals.a_nom * 2.0, note="A scales with Quantity")


def test_probability_scales_the_emv_but_never_knom_or_kpv() -> None:
    """Probability is replaced by a Bernoulli draw later, so folding it into Kpv
    would double-count it."""
    base = calculate(_model(risks=(_risk(probability=0.3),)), TOL)
    doubled = calculate(_model(risks=(_risk(probability=0.6),)), TOL)
    assert doubled.drivers[0].knom == base.drivers[0].knom
    assert doubled.drivers[0].kpv == base.drivers[0].kpv
    _close(doubled.totals.d_nom, base.totals.d_nom * 2.0, note="D scales with Probability")


def test_probability_zero_is_valid_and_gives_zero_emv() -> None:
    result = calculate(_model(risks=(_risk(probability=0),)), TOL)
    assert result.totals.d_nom == 0.0
    assert result.totals.d_pv == 0.0
    assert_reconciled(result, TOL)


def test_probability_one_is_valid() -> None:
    result = calculate(_model(risks=(_risk(probability=1),)), TOL)
    _close(result.totals.d_nom, 250.0, note="D at certainty is the expected severity")


def test_a_probability_outside_zero_to_one_is_refused() -> None:
    for bad in (-0.1, 1.1, None, "half"):
        _refuses(
            lambda b=bad: calculate(_model(risks=(_risk(probability=b),)), TOL),
            f"probability {bad!r}",
        )


def test_negative_but_ordered_cost_values_are_allowed() -> None:
    """No locked contract imposes positivity on Min / ML / Max, and inventing one
    would be inventing a business rule."""
    result = calculate(
        _model(costs=(_cost(minimum=-150, most_likely=-100, maximum=-80),)), TOL
    )
    assert result.totals.a_nom < 0.0
    assert_reconciled(result, TOL)


def test_negative_but_ordered_risk_impacts_are_allowed() -> None:
    result = calculate(
        _model(risks=(_risk(minimum=-450, most_likely=-200, maximum=-100),)), TOL
    )
    assert result.totals.d_nom < 0.0
    assert_reconciled(result, TOL)


def test_an_out_of_order_three_point_set_is_refused() -> None:
    for distribution in ("Triangular", "Beta-PERT"):
        _refuses(
            lambda d=distribution: calculate(
                _model(costs=(_cost(distribution=d, minimum=150, most_likely=100,
                                    maximum=80),)),
                TOL,
            ),
            f"{distribution} out of order",
        )


def test_a_uniform_with_min_above_max_is_refused() -> None:
    _refuses(
        lambda: calculate(
            _model(costs=(_cost(distribution="Uniform", minimum=150, most_likely=None,
                                maximum=80),)),
            TOL,
        ),
        "Uniform Min > Max",
    )


def test_an_invalid_distribution_is_refused() -> None:
    for bad in ("Normal", "triangular", "", "PERT"):
        message = _refuses(
            lambda b=bad: calculate(_model(costs=(_cost(distribution=b),)), TOL),
            f"distribution {bad!r}",
        )
        assert "CL-001" in message


def test_an_empty_driver_set_is_not_refused() -> None:
    """No accepted contract requires at least one driver, so none is invented."""
    result = calculate(_model(), TOL)
    assert result.totals.a_nom == 0.0
    assert result.totals.e_pv == 0.0
    assert result.drivers == ()
    assert len(result.annual) == 1
    assert_reconciled(result, TOL)


def test_a_python_integer_too_large_for_a_double_is_a_structured_refusal() -> None:
    """`float(10**400)` raises `OverflowError`, and a raw `OverflowError` escaping
    the oracle would bypass the whole failure contract.

    The pure oracle accepts plain Python numbers, so it must honour its own API:
    every one of these produces a `ModelInputRefusal` naming the subject, never a
    conversion error.
    """
    huge = 10 ** 400
    for label, model in (
        ("Quantity", _model(costs=(_cost(quantity=huge),))),
        ("FX rate", _model(fx=(FxRow("SAR", 1), FxRow("USD", huge)),
                           costs=(_cost(currency="USD"),))),
        ("discount rate", _model(discount=huge, costs=(_cost(),))),
        ("Min", _model(costs=(_cost(minimum=-huge),))),
        ("profile weight", _model(costs=(_cost(weights=(huge,)),))),
        ("inflation rate", _model(base=2026, start=2027, duration=1,
                                  rates={"Standard": {2027: huge}}, costs=(_cost(),))),
        ("Probability", _model(risks=(_risk(probability=huge),))),
    ):
        try:
            calculate(model, TOL)
        except OverflowError as error:      # noqa: PERF203 - the point of the test
            raise AssertionError(f"{label}: raw OverflowError escaped: {error}") from error
        except (ModelInputRefusal, NumericalRangeRefusal):
            continue
        raise AssertionError(f"{label}: a huge integer was silently accepted")


def test_a_base_year_after_the_start_year_is_refused() -> None:
    message = _refuses(
        lambda: calculate(_model(base=2030, start=2026, costs=(_cost(),)), TOL),
        "Base Year > Start Year",
    )
    assert "Base Year" in message


def test_a_blank_or_non_numeric_discount_rate_is_refused() -> None:
    for bad in (None, "ten percent"):
        _refuses(
            lambda b=bad: calculate(_model(discount=b, costs=(_cost(),)), TOL),
            f"discount rate {bad!r}",
        )


# ---------------------------------------------------------------------------
# ANNUAL SERIES AND RECONCILIATION
# ---------------------------------------------------------------------------
def test_the_annual_series_uses_the_mean_basis_not_the_deterministic_basis() -> None:
    """Locked: annual cash flow is mean-only. There is no deterministic series."""
    result = calculate(
        _model(base=2026, start=2027, duration=3, rates=_three_year(),
               costs=(_cost(weights=(0.2, 0.5, 0.3)),)),
        TOL,
    )
    annual_total = sum(row.base_cost_nominal for row in result.annual)
    _close(annual_total, result.totals.c_nom, note="annual base sums to C, not A")
    assert abs(annual_total - result.totals.a_nom) > 1.0


def test_each_annual_row_carries_its_own_calendar_year() -> None:
    result = calculate(
        _model(base=2026, start=2027, duration=3, rates=_three_year(),
               costs=(_cost(weights=(0.2, 0.5, 0.3)),)),
        TOL,
    )
    assert [(row.project_index, row.calendar_year) for row in result.annual] == [
        (1, 2027), (2, 2028), (3, 2029)
    ]


def test_all_reconciliation_identities_hold_on_a_mixed_model() -> None:
    result = calculate(
        _model(base=2026, start=2027, duration=3, rates=_three_year(),
               fx=(FxRow("SAR", 1), FxRow("USD", 3.75)),
               costs=(_cost("CL-001", weights=(0.2, 0.5, 0.3)),
                      _cost("CL-002", distribution="Beta-PERT", currency="USD",
                            weights=(0.1, 0.1, 0.8), quantity=3)),
               risks=(_risk("R-001", weights=(0.0, 0.5, 0.5)),
                      _risk("R-002", distribution="Uniform", most_likely=None,
                            weights=(1.0, 0.0, 0.0), probability=0.75))),
        TOL,
    )
    checks = reconcile(result, TOL)
    names = {check.name for check in checks}
    for expected in ("I1 nominal: A + B = C", "I2 PV: C + D = E", "I3c nominal total",
                     "I4a PV base", "I5 profile sum: R-002"):
        assert expected in names
    failing = [check.name for check in checks if not check.holds]
    assert not failing, failing
    assert_reconciled(result, TOL)


def test_reconciliation_failure_is_an_internal_invariant_error_not_a_refusal() -> None:
    """A user must never be told their model is invalid because the calculation
    disagreed with itself.

    The failure is simulated by corrupting a total AFTER a successful, accepted
    calculation - which is exactly the situation the class distinction exists for:
    the inputs were fine, the calculation ran, and two quantities that must agree
    no longer do.
    """
    result = calculate(_model(costs=(_cost(),)), TOL)
    corrupted = dataclasses.replace(
        result, totals=dataclasses.replace(result.totals, c_nom=result.totals.c_nom + 1000.0)
    )
    try:
        assert_reconciled(corrupted, TOL)
    except OracleInvariantError as error:
        assert not isinstance(error, (ModelInputRefusal, NumericalRangeRefusal))
        assert "reconciliation failed" in str(error)
        assert "I1 nominal" in str(error)
        return
    raise AssertionError("a failing identity did not raise an invariant error")


def test_the_driver_audit_columns_reconstruct_the_headline_measures() -> None:
    """Plain column sums over rows of one kind — the property the audit exists for."""
    result = calculate(
        _model(base=2026, start=2027, duration=3, rates=_three_year(),
               costs=(_cost("CL-001", weights=(0.2, 0.5, 0.3)),
                      _cost("CL-002", weights=(0.6, 0.2, 0.2), quantity=2)),
               risks=(_risk("R-001", weights=(0.3, 0.3, 0.4)),)),
        TOL,
    )
    costs = [d for d in result.drivers if d.driver_kind is DriverKind.COST_LINE]
    risks = [d for d in result.drivers if d.driver_kind is DriverKind.RISK]
    _close(sum(d.deterministic_nominal for d in costs), result.totals.a_nom, note="A")
    _close(sum(d.uncertainty_mean_shift_nominal for d in costs), result.totals.b_nom, note="B")
    _close(sum(d.mean_basis_nominal for d in costs), result.totals.c_nom, note="C")
    _close(sum(d.expected_risk_nominal for d in risks), result.totals.d_nom, note="D")


def test_no_audit_column_carries_two_meanings_by_driver_kind() -> None:
    result = calculate(_model(costs=(_cost(),), risks=(_risk(),)), TOL)
    cost = next(d for d in result.drivers if d.driver_kind is DriverKind.COST_LINE)
    risk = next(d for d in result.drivers if d.driver_kind is DriverKind.RISK)
    assert cost.probability is None and cost.expected_risk_nominal is None
    assert risk.quantity is None and risk.central_value is None
    assert risk.mean_basis_nominal is None and risk.deterministic_nominal is None


def test_an_independent_rational_derivation_agrees_with_the_oracle() -> None:
    """A SECOND independent check, exact rather than floating point.

    It supplements the accepted literals above; it does not replace them.
    """
    weights = [Fraction(1, 5), Fraction(1, 2), Fraction(3, 10)]
    rate = Fraction(1, 20)
    factors = [(1 + rate) ** k for k in (1, 2, 3)]
    knom = sum(w * f for w, f in zip(weights, factors))
    central = Fraction(100)
    quantity = Fraction(10)
    expected_a = float(central * quantity * knom)

    result = calculate(
        _model(base=2026, start=2027, duration=3, rates=_three_year(),
               costs=(_cost(weights=(0.2, 0.5, 0.3)),)),
        TOL,
    )
    _close(result.totals.a_nom, expected_a, note="A_nom against exact rationals")
    _close(result.drivers[0].knom, float(knom), note="Knom against exact rationals")


# ---------------------------------------------------------------------------
# CONDITIONING METADATA MUST NOT REFUSE A REPRESENTABLE MODEL
# ---------------------------------------------------------------------------
def test_a_subnormal_but_representable_model_calculates() -> None:
    """The scaled conditioning term `1e-12 * 2e-312` rounds to exactly zero.

    Under the model-arithmetic underflow rule that was a refusal, so this model —
    whose economic outputs are perfectly representable — was rejected because the
    TOLERANCE BOOKKEEPING could not hold a term far too small to affect the
    answer. The refusal even named `totals, driver 'CL-001': |A| nominal`.
    """
    result = calculate(
        _model(discount=0.0,
               costs=(_cost(distribution="Uniform", minimum=2e-312, most_likely=None,
                            maximum=2e-312, weights=(1.0,), quantity=1),)),
        TOL,
    )
    assert result.totals.a_nom > 0.0
    assert result.totals.a_nom < 1e-300
    _close(result.drivers[0].mean_value, 2e-312, note="Uniform midpoint")

    checks = {check.name: check for check in reconcile(result, TOL)}
    assert checks["I1 nominal: A + B = C"].allowance == TOL.identity_absolute_floor
    assert_reconciled(result, TOL)


def test_the_conditioning_exception_does_not_weaken_model_arithmetic_in_the_oracle() -> None:
    """Scoped to conditioning metadata: an economic collapse is still refused."""
    # A valid profile (sums to 1) and a positive Quantity, but the CONTRIBUTION
    # `1e-200 * 1e-200 * 1` collapses to exactly zero — a real economic value
    # deleted with no error anywhere, which stays a refusal.
    message = _refuses(
        lambda: calculate(
            _model(discount=0.0,
                   costs=(_cost(distribution="Uniform", minimum=1e-200, most_likely=None,
                                maximum=1e-200, weights=(1.0,), quantity=1e-200),)),
            TOL,
        ),
        "an economic contribution collapsing to zero",
    )
    assert "underflow" in message.lower()
    assert "CL-001" in message


# ---------------------------------------------------------------------------
# REFERENCE FIELDS ARE STRUCTURED REFUSALS, NEVER PYTHON ERRORS
# ---------------------------------------------------------------------------
_INVALID_IDENTIFIERS = (None, True, False, 123, 1.5, "", "   ", [], {})


def test_every_invalid_driver_reference_field_is_a_structured_refusal() -> None:
    """These reached the UTF-16 ordering helper and escaped as raw
    `AttributeError` / `TypeError` — a user-editable required field surfacing as a
    Python implementation error."""
    for field_name in ("currency", "inflation_profile", "distribution"):
        for bad in _INVALID_IDENTIFIERS:
            for build, kind in ((_cost, "cost line"), (_risk, "risk")):
                driver = dataclasses.replace(build(), **{field_name: bad})
                key = "costs" if kind == "cost line" else "risks"
                model = _model(**{key: (driver,)})
                try:
                    calculate(model, TOL)
                except ModelInputRefusal as error:
                    text = str(error)
                    assert driver.permanent_id in text, (field_name, bad, text)
                    assert kind in text, (field_name, bad, text)
                    continue
                except Exception as error:  # noqa: BLE001
                    raise AssertionError(
                        f"{kind} {field_name}={bad!r}: raw {type(error).__name__}: {error}"
                    ) from error
                raise AssertionError(f"{kind} {field_name}={bad!r}: silently accepted")


def test_the_refusal_names_the_field_and_the_offending_value() -> None:
    for field_name, label in (
        ("currency", "Currency"),
        ("inflation_profile", "Inflation Profile"),
        ("distribution", "Distribution"),
    ):
        message = _refuses(
            lambda f=field_name: calculate(
                _model(costs=(dataclasses.replace(_cost(), **{f: 123}),)), TOL
            ),
            f"{field_name} = 123",
        )
        assert label in message
        assert "123" in message
        assert "CL-001" in message


def test_a_blank_reference_field_says_so() -> None:
    for bad in (None, "", "   "):
        message = _refuses(
            lambda b=bad: calculate(_model(costs=(_cost(currency=b),)), TOL),
            f"currency {bad!r}",
        )
        assert "blank" in message


def test_a_non_blank_identifier_is_used_exactly_as_entered_never_repaired() -> None:
    """Config keys are exact. Silently trimming `" USD "` into `"USD"` to make a
    lookup succeed would be rewriting user data to invent an answer."""
    padded = " USD "
    message = _refuses(
        lambda: calculate(
            _model(fx=(FxRow("SAR", 1), FxRow("USD", 3.75)),
                   costs=(_cost(currency=padded),)),
            TOL,
        ),
        "a padded currency silently trimmed",
    )
    assert repr(padded) in message or padded in message
    assert "no row" in message

    # It resolves only if the FX table carries that exact key.
    exact = calculate(
        _model(fx=(FxRow("SAR", 1), FxRow(padded, 3.75)), costs=(_cost(currency=padded),)),
        TOL,
    )
    assert set(exact.resolved_fx) == {padded}


def test_a_non_text_permanent_id_is_refused_before_canonical_ordering() -> None:
    """Canonical ordering compares UTF-16 code units and so needs text.

    This is NOT Phase-4 permanent-ID structural validation — the `CL-`/`R-`
    prefixes, the pattern and the counter rules stay owned by Phase 4 and are not
    re-checked here. It only stops a non-text field escaping as an
    `AttributeError`.
    """
    for bad in (None, 123, True, "", "  "):
        message = _refuses(
            lambda b=bad: calculate(
                _model(costs=(dataclasses.replace(_cost(), permanent_id=b),)), TOL
            ),
            f"permanent id {bad!r}",
        )
        assert "Permanent ID" in message


# ---------------------------------------------------------------------------
# resolved_fx IS THE REFERENCED SET - the tblCalcFX row rule
# ---------------------------------------------------------------------------
_FX_TABLE = (FxRow("SAR", 1), FxRow("USD", 3.75), FxRow("AED", 0.98))


def _fx_model(*currencies: str) -> CalculationModel:
    costs = tuple(
        _cost(f"CL-{index:03d}", currency=currency)
        for index, currency in enumerate(currencies, start=1)
    )
    return _model(fx=_FX_TABLE, costs=costs)


def test_resolved_fx_contains_exactly_the_referenced_currencies() -> None:
    """`tblCalcFX` row rule: "one row per referenced currency".

    Being validated globally does not make a currency referenced. An earlier
    version seeded `{"SAR": 1.0}` unconditionally, so a USD-only model reported
    SAR as resolved and an empty model reported one resolved currency.
    """
    assert set(calculate(_model(fx=_FX_TABLE), TOL).resolved_fx) == set()
    assert set(calculate(_fx_model("SAR"), TOL).resolved_fx) == {"SAR"}
    assert set(calculate(_fx_model("USD"), TOL).resolved_fx) == {"USD"}
    assert set(calculate(_fx_model("USD", "SAR"), TOL).resolved_fx) == {"SAR", "USD"}
    assert set(calculate(_fx_model("USD", "AED", "SAR"), TOL).resolved_fx) == {
        "AED", "SAR", "USD"
    }


def test_resolved_fx_carries_the_right_rates() -> None:
    resolved = calculate(_fx_model("USD", "AED", "SAR"), TOL).resolved_fx
    assert resolved["SAR"] == 1.0
    assert resolved["USD"] == 3.75
    assert resolved["AED"] == 0.98


def test_resolved_fx_iterates_in_canonical_order() -> None:
    """Key ORDER, not only key equality: the SAR check must not disturb it."""
    result = calculate(_fx_model("USD", "AED", "SAR"), TOL)
    assert list(result.resolved_fx) == ["AED", "SAR", "USD"]

    # And the order does not follow the order the drivers arrived in.
    reversed_ = calculate(_fx_model("SAR", "AED", "USD"), TOL)
    assert list(reversed_.resolved_fx) == ["AED", "SAR", "USD"]


def test_the_sar_invariant_is_still_enforced_when_sar_is_unreferenced() -> None:
    """The global check is unchanged; only the RETURNED SET changed."""
    for bad_table in (
        (FxRow("SAR", 2), FxRow("USD", 3.75)),
        (FxRow("SAR", None), FxRow("USD", 3.75)),
        (FxRow("USD", 3.75),),
        (FxRow("SAR", 1), FxRow("SAR", 1), FxRow("USD", 3.75)),
    ):
        message = _refuses(
            lambda t=bad_table: calculate(
                _model(fx=t, costs=(_cost(currency="USD"),)), TOL
            ),
            f"SAR invariant with {bad_table}",
        )
        assert "SAR" in message


# ---------------------------------------------------------------------------
# ARCHITECTURE BOUNDARY - executable dependencies, not vocabulary
# ---------------------------------------------------------------------------
FORBIDDEN_IMPORTS = frozenset(
    {"openpyxl", "win32com", "pythoncom", "xlwings", "random", "secrets", "numpy", "scipy"}
)
FORBIDDEN_NAMES = frozenset(
    {
        "Workbook", "Worksheet", "Worksheets", "Range", "Cells", "ListObject", "ListObjects",
        "ThisWorkbook", "ActiveWorkbook", "Application", "Rnd", "Randomize", "percentile",
        "quantile",
    }
)


def _imported_modules(path: Path) -> set[str]:
    """Top-level module names of every EXECUTABLE import statement."""
    tree = ast.parse(path.read_text(encoding="utf-8"))
    modules: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            modules.update(alias.name.split(".")[0] for alias in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module and node.level == 0:
            modules.add(node.module.split(".")[0])
    return modules


def _referenced_identifiers(path: Path) -> set[str]:
    """Identifiers actually referenced in code.

    Walking the AST is what makes this a dependency test rather than a fragile
    word ban: comments never enter the tree at all, and docstrings are string
    constants rather than `Name` or `Attribute` nodes, so prose about later
    phases cannot trip it.
    """
    tree = ast.parse(path.read_text(encoding="utf-8"))
    names: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Name):
            names.add(node.id)
        elif isinstance(node, ast.Attribute):
            names.add(node.attr)
    return names


def test_the_oracle_imports_no_excel_com_or_random_dependency() -> None:
    for path in (ORACLE_PATH, NUMERIC_PATH):
        offending = _imported_modules(path) & FORBIDDEN_IMPORTS
        assert not offending, f"{path.name} imports {sorted(offending)}"


def test_the_oracle_references_no_workbook_or_simulation_identifier() -> None:
    for path in (ORACLE_PATH, NUMERIC_PATH):
        offending = _referenced_identifiers(path) & FORBIDDEN_NAMES
        assert not offending, f"{path.name} references {sorted(offending)}"


def test_prose_about_later_phases_does_not_trip_the_boundary_test() -> None:
    """Guards the guard: the modules DO discuss workbooks and Monte Carlo in
    comments and docstrings, and must be allowed to."""
    text = ORACLE_PATH.read_text(encoding="utf-8")
    assert "Monte Carlo" in text and "ListObject" in text
    assert not _referenced_identifiers(ORACLE_PATH) & FORBIDDEN_NAMES


def test_the_oracle_runs_with_no_excel_library_importable() -> None:
    """The strongest form: load the pure modules in a fresh interpreter, WITHOUT
    the `pccm_builder` package (whose `__init__` legitimately imports openpyxl),
    run a calculation, and assert no forbidden module was ever loaded.

    RUN WITH `-S`. Site initialisation can import third-party packages before any
    of our code executes - a `.pth` file, a sitecustomize hook, a vendored
    distribution - and review found an environment where `numpy`, `random` and
    `secrets` were already in `sys.modules` at interpreter start. Counting those
    against the oracle would conflate "present before the import" with "loaded
    because of the import", which is not what this test claims. Disabling site
    processing makes the claim exact and the result portable, and it is also the
    stronger statement: the two modules run on a bare interpreter.
    """
    script = textwrap.dedent(
        f"""
        import importlib.util, sys, types
        builder = {str(PCCM_ROOT / "builder" / "pccm_builder")!r}
        pkg = types.ModuleType("pccm5")
        pkg.__path__ = [builder]
        sys.modules["pccm5"] = pkg
        for name in ("calc_numeric", "calc_fingerprint", "calc_oracle"):
            spec = importlib.util.spec_from_file_location(
                f"pccm5.{{name}}", builder + "/" + name + ".py"
            )
            module = importlib.util.module_from_spec(spec)
            sys.modules[f"pccm5.{{name}}"] = module
            spec.loader.exec_module(module)
        oracle = sys.modules["pccm5.calc_oracle"]
        model = oracle.CalculationModel(
            timeline=oracle.AppliedTimeline(2026, 2026, 1),
            discount_rate=0.10,
            fx_rows=(oracle.FxRow("SAR", 1),),
            inflation_rates={{"Standard": {{}}}},
            cost_drivers=(oracle.CostDriver(
                "CL-001", "Triangular", "SAR", "Standard", 80, 100, 150, (1.0,), quantity=10
            ),),
        )
        tolerances = oracle.Tolerances(1e-9, 1e-6, 1e-12, 1.0)
        result = oracle.calculate(model, tolerances)
        assert result.totals.a_nom == 1000.0, result.totals.a_nom
        loaded = {{m for m in sys.modules if m.split(".")[0] in {sorted(FORBIDDEN_IMPORTS)!r}}}
        assert not loaded, "forbidden modules present: " + repr(sorted(loaded))
        print("OK")
        """
    )
    completed = subprocess.run(
        [sys.executable, "-S", "-c", script], capture_output=True, text=True, check=False
    )
    assert completed.returncode == 0, completed.stderr
    assert "OK" in completed.stdout


def test_e_is_accumulated_independently_and_not_derived_from_c_and_d() -> None:
    """A STRUCTURAL guard, because a behavioural one cannot see this.

    `E = C + D` produces the same number as independent accumulation for every
    ordinary model — that is precisely why I2 is worth checking — so no golden
    value can distinguish the two implementations. Substituting the derivation
    was verified to leave the entire suite passing.

    What the derivation destroys is the MEANING of I2: an identity computed by
    definition can never fail, so it stops being a check. The guard is therefore
    on the shape of `_accumulate_totals`: `e_nom` and `e_pv` must be summed over
    their OWN contribution lists, those lists must be filled from the drivers, and
    neither the sum nor the fill may draw on `c_*` or `d_*`.

    Erratum C2 changed the mechanism from a running `safe_accumulate` total to a
    contribution list summed by `safe_signed_sum`. The guard follows the mechanism
    and gains a second half: it is no longer enough that the SUM is independent,
    the LIST it sums must also be appended from the drivers rather than seeded
    from another measure's list.
    """
    tree = ast.parse(ORACLE_PATH.read_text(encoding="utf-8"))
    function = next(
        node for node in ast.walk(tree)
        if isinstance(node, ast.FunctionDef) and node.name == "_accumulate_totals"
    )
    forbidden = {
        "c_nom", "c_pv", "d_nom", "d_pv",
        "c_nom_terms", "c_pv_terms", "d_nom_terms", "d_pv_terms",
    }

    derived_from: set[str] = set()
    summed: set[str] = set()
    for node in ast.walk(function):
        if not isinstance(node, ast.Assign):
            continue
        targets = {t.id for t in node.targets if isinstance(t, ast.Name)}
        if not targets & {"e_nom", "e_pv", "e_nom_terms", "e_pv_terms"}:
            continue
        sources = {n.id for n in ast.walk(node.value) if isinstance(n, ast.Name)}
        derived_from |= sources & forbidden
        if "safe_signed_sum" in sources and sources & {"e_nom_terms", "e_pv_terms"}:
            summed |= targets & {"e_nom", "e_pv"}

    assert not derived_from, (
        f"E is assigned from {sorted(derived_from)}. `E = C + D` is a reconciliation "
        "identity, not the calculation path; deriving it makes I2 unfalsifiable."
    )
    assert summed == {"e_nom", "e_pv"}, (
        f"E must be summed by safe_signed_sum over its own term list; summed {summed}"
    )
    assert _appended_from(function, "e_nom_terms") == {"e_nom_terms"}, (
        "e_nom_terms must be appended to from the driver pass, not seeded elsewhere"
    )
    assert _appended_from(function, "e_pv_terms") == {"e_pv_terms"}, (
        "e_pv_terms must be appended to from the driver pass, not seeded elsewhere"
    )


def _appended_from(function: ast.FunctionDef, name: str) -> set[str]:
    """The term lists this function passes to `contribute` under `name`.

    `contribute(key, into, value, ...)` appends `value` to `into`, so a list that
    is never passed as `into` was never filled from a driver.
    """
    found: set[str] = set()
    for node in ast.walk(function):
        if not isinstance(node, ast.Call):
            continue
        if not (isinstance(node.func, ast.Name) and node.func.id == "contribute"):
            continue
        if len(node.args) >= 2 and isinstance(node.args[1], ast.Name):
            if node.args[1].id == name:
                found.add(name)
    return found


def test_b_is_accumulated_independently_and_not_derived_from_c_and_a() -> None:
    """The same rule for `B = C - A`, and for the same reason: it would make I1
    unfalsifiable."""
    tree = ast.parse(ORACLE_PATH.read_text(encoding="utf-8"))
    function = next(
        node for node in ast.walk(tree)
        if isinstance(node, ast.FunctionDef) and node.name == "_accumulate_totals"
    )
    forbidden = {
        "a_nom", "a_pv", "c_nom", "c_pv",
        "a_nom_terms", "a_pv_terms", "c_nom_terms", "c_pv_terms",
    }
    for node in ast.walk(function):
        if not isinstance(node, ast.Assign):
            continue
        targets = {t.id for t in node.targets if isinstance(t, ast.Name)}
        if not targets & {"b_nom", "b_pv", "b_nom_terms", "b_pv_terms"}:
            continue
        sources = {n.id for n in ast.walk(node.value) if isinstance(n, ast.Name)}
        assert not sources & forbidden, f"B is assigned from {sorted(sources & forbidden)}"
    assert _appended_from(function, "b_nom_terms") == {"b_nom_terms"}
    assert _appended_from(function, "b_pv_terms") == {"b_pv_terms"}


def test_the_uncertainty_shift_comes_from_the_driver_not_from_the_totals() -> None:
    """Per driver, B is `(mean - central) * Qty * K`, computed at the driver.

    Behavioural companion to the structural guards above: the audit column must
    hold the driver's own shift, so a totals-level derivation would leave it empty
    or wrong.
    """
    result = calculate(
        _model(costs=(_cost("CL-001", quantity=10), _cost("CL-002", quantity=3)),), TOL
    )
    for driver in result.drivers:
        expected_shift = (driver.mean_value - driver.central_value) * driver.quantity
        _close(driver.uncertainty_mean_shift_nominal, expected_shift * driver.knom,
               note=f"B for {driver.permanent_id}")


def test_the_oracle_writes_nothing_anywhere() -> None:
    """No file I/O of any kind: the pure layer never reads YAML or writes output."""
    forbidden = {"open", "write_text", "read_text", "safe_load", "dump", "mkdir"}
    assert not _referenced_identifiers(ORACLE_PATH) & forbidden
    assert not _referenced_identifiers(NUMERIC_PATH) & forbidden


def test_the_distribution_adapter_is_not_a_second_authority() -> None:
    """The master list stays in `input_contract.yaml`; this maps names to shapes.

    A name the adapter does not know is refused as an invalid distribution rather
    than silently accepted, so the adapter cannot quietly widen the offering.
    """
    from pccm_builder.calc_oracle import _DISTRIBUTION_ADAPTER, DistributionKind

    assert set(_DISTRIBUTION_ADAPTER) == {"Triangular", "Beta-PERT", "Uniform"}
    assert len(set(_DISTRIBUTION_ADAPTER.values())) == len(DistributionKind)
    _refuses(
        lambda: calculate(_model(costs=(_cost(distribution="Lognormal"),)), TOL),
        "a distribution the adapter does not know",
    )


def test_the_accepted_distribution_names_still_match_the_input_contract() -> None:
    """If the upstream master list ever changes, this fails loudly."""
    import yaml

    from pccm_builder.calc_oracle import _DISTRIBUTION_ADAPTER

    document = yaml.safe_load(
        (PCCM_ROOT / "spec" / "input_contract.yaml").read_text(encoding="utf-8")
    )
    table = next(t for t in document["config_tables"] if t["key"] == "distributions")
    upstream = set(table["values"])
    assert upstream == set(_DISTRIBUTION_ADAPTER), (
        f"upstream distributions {sorted(upstream)} disagree with the adapter "
        f"{sorted(_DISTRIBUTION_ADAPTER)}"
    )


# ---------------------------------------------------------------------------
# Erratum C2 - end to end
# ---------------------------------------------------------------------------
# The two models that were refused before the patch. Both are valid: every input
# is a usable Double, every rule is satisfied, and every headline the model asks
# for is representable. Only the ORDER of the intermediate additions made them
# fail, which is exactly what §19.2 says must not happen.
def _degenerate_cost(permanent_id: str, value: object, weights=(1.0,)) -> CostDriver:
    """A Uniform cost line with `Min == Max`, so its statistic is `value` exactly
    and the arithmetic under test is the SUM, not the distribution."""
    return CostDriver(
        permanent_id, "Uniform", "SAR", "Standard", value, None, value, weights, quantity=1
    )


def test_a_headline_total_that_cancels_to_a_representable_value_is_calculated() -> None:
    """REPRODUCER A. Three cost lines at `+MAX`, `+MAX`, `-MAX`.

    In canonical permanent-ID order the first addition is `MAX + MAX`, which has no
    Double. The total the model actually asks for is `MAX`, which does. Before the
    patch this refused with
    `NumericalRangeRefusal: totals, driver 'CL-002': A nom`.
    """
    model = _model(
        discount=0.0,
        costs=(
            _degenerate_cost("CL-001", MAX_DOUBLE),
            _degenerate_cost("CL-002", MAX_DOUBLE),
            _degenerate_cost("CL-003", -MAX_DOUBLE),
        ),
    )
    result = calculate(model, TOL)

    assert result.totals.a_nom == MAX_DOUBLE, f"A_nom is {result.totals.a_nom!r}"
    assert result.totals.c_nom == MAX_DOUBLE, f"C_nom is {result.totals.c_nom!r}"
    assert result.totals.e_nom == MAX_DOUBLE, f"E_nom is {result.totals.e_nom!r}"
    assert result.totals.a_pv == MAX_DOUBLE
    assert result.totals.b_nom == 0.0, "a degenerate Uniform has no mean shift"
    assert len(result.annual) == 1
    assert result.annual[0].base_cost_nominal == MAX_DOUBLE
    assert result.annual[0].total_nominal == MAX_DOUBLE
    assert_reconciled(result, TOL)


def test_a_profile_whose_weights_cancel_to_one_hundred_percent_is_accepted() -> None:
    """REPRODUCER B. A five-year profile of `[MAX, MAX, -MAX, -MAX, 1]`.

    The weights sum to exactly `1`, so the profile IS 100%. Validating that sum in
    project-year order overflows at year 2, and before the patch the model was
    refused with
    `NumericalRangeRefusal: profiling for driver 'CL-001', project year 2`.

    NO POSITIVITY RULE IS INVOLVED. A negative profile weight stays legal; what
    changed is that the sum being checked is computed so it can be checked at all.
    The same cancellation then has to survive `Knom`, `Kpv`, the annual series and
    the annual-to-headline reconciliation, so this one model exercises every sum
    §4 of the patch lists.
    """
    weights = (MAX_DOUBLE, MAX_DOUBLE, -MAX_DOUBLE, -MAX_DOUBLE, 1.0)
    model = _model(
        base=2026, start=2026, duration=5, discount=0.0,
        rates={"Standard": {year: 0.0 for year in range(2027, 2031)}},
        costs=(_degenerate_cost("CL-001", 1.0, weights),),
    )
    result = calculate(model, TOL)

    driver = result.drivers[0]
    assert driver.weights == weights, "the weights are used as supplied, not normalised"
    assert driver.knom == 1.0, f"Knom is {driver.knom!r}"
    assert driver.kpv == 1.0, f"Kpv is {driver.kpv!r}"
    assert result.totals.a_nom == 1.0
    assert result.totals.c_nom == 1.0
    # Each annual row is one driver's contribution, so the rows themselves are the
    # weights. The RECONCILIATION back to the headline is where they cancel.
    assert [row.base_cost_nominal for row in result.annual] == list(weights)
    assert_reconciled(result, TOL)


def test_the_cancelling_profile_is_still_refused_when_a_year_is_unrepresentable() -> None:
    """The patch repairs the SUM, not the model. Give the same profile a unit cost
    of 100 and project year 1 genuinely costs `100 * MAX`, which no Double holds -
    and that is refused, naming the year."""
    message = _refuses(
        lambda: calculate(
            _model(
                base=2026, start=2026, duration=5, discount=0.0,
                rates={"Standard": {year: 0.0 for year in range(2027, 2031)}},
                costs=(
                    _degenerate_cost(
                        "CL-001", 100.0, (MAX_DOUBLE, MAX_DOUBLE, -MAX_DOUBLE, -MAX_DOUBLE, 1.0)
                    ),
                ),
            ),
            TOL,
        ),
        "an annual contribution with no representable value",
    )
    assert "annual year 2026" in message and "CL-001" in message


def test_a_headline_total_that_genuinely_exceeds_double_range_is_still_refused() -> None:
    """Two cost lines at `+MAX` with nothing to cancel them. The total really is
    `2 * MAX`, and inventing a number for it would be worse than refusing."""
    message = _refuses(
        lambda: calculate(
            _model(
                discount=0.0,
                costs=(
                    _degenerate_cost("CL-001", MAX_DOUBLE),
                    _degenerate_cost("CL-002", MAX_DOUBLE),
                ),
            ),
            TOL,
        ),
        "a headline total with no representable value",
    )
    assert "totals" in message


def test_ordinary_models_are_bit_for_bit_unchanged_by_the_signed_sum() -> None:
    """THE REGRESSION THAT MATTERS MOST.

    Erratum C2 must move no number that already had one. Canonical order is still
    tier 1, so for every ordinary model the totals are exactly what a plain
    left-to-right accumulation of the same contributions in the same order gives.
    """
    result = calculate(
        _model(
            base=2026, start=2027, duration=3, discount=0.08,
            rates=_three_year(),
            fx=(FxRow("SAR", 1), FxRow("USD", 3.75)),
            costs=(
                _cost("CL-001", weights=(0.2, 0.5, 0.3)),
                _cost("CL-002", currency="USD", distribution="Beta-PERT",
                      weights=(0.5, 0.25, 0.25), quantity=3),
                _cost("CL-003", distribution="Uniform", most_likely=None,
                      weights=(0.1, 0.1, 0.8), quantity=7),
            ),
            risks=(
                _risk("R-001", weights=(0.0, 1.0, 0.0)),
                _risk("R-002", distribution="Uniform", most_likely=None,
                      weights=(0.34, 0.33, 0.33), probability=0.6),
            ),
        ),
        TOL,
    )
    costs = [d for d in result.drivers if d.driver_kind is DriverKind.COST_LINE]
    risks = [d for d in result.drivers if d.driver_kind is DriverKind.RISK]

    def left_to_right(values) -> float:
        total = 0.0
        for value in values:
            total = total + value
        return total

    assert result.totals.a_nom == left_to_right([d.deterministic_nominal for d in costs])
    assert result.totals.c_nom == left_to_right([d.mean_basis_nominal for d in costs])
    assert result.totals.c_pv == left_to_right([d.mean_basis_pv for d in costs])
    assert result.totals.d_nom == left_to_right([d.expected_risk_nominal for d in risks])
    assert result.totals.e_nom == left_to_right(
        [d.mean_basis_nominal for d in costs] + [d.expected_risk_nominal for d in risks]
    )
    assert_reconciled(result, TOL)


# ---------------------------------------------------------------------------
# Round 3 - the faithful rescue, end to end
# ---------------------------------------------------------------------------
# The seven-term reproducer as a real model. Every input is a usable Double,
# every rule is satisfied, and the headline the model asks for is representable —
# but the canonical order overflows on the way there, and a rescue that
# re-associates Doubles gets a number that is 100% wrong.
_RESIDUAL_TERMS = (6e307, -8e307, -1.7e308, 6e307, 7e307, 6e307, -1e292)
_RESIDUAL_TOTAL = -1e292


def _residual_model() -> CalculationModel:
    return _model(
        discount=0.0,
        costs=tuple(
            _degenerate_cost(f"CL-{index + 1:03d}", value)
            for index, value in enumerate(_RESIDUAL_TERMS)
        ),
    )


def test_a_headline_whose_cancellation_leaves_a_rounding_residual_is_exact() -> None:
    """REPRODUCER §1.2, end to end.

    `A_nom` is the exact signed sum of seven cost lines, `-1e292`. The round-2
    rescue answered `-1.99792015476736e292` because cancelling the two largest
    opposite-signed magnitudes with one rounded subtraction discarded the very
    residual that survives.
    """
    exact = sum((Fraction(term) for term in _RESIDUAL_TERMS), Fraction(0))
    assert float(exact) == _RESIDUAL_TOTAL, "the fixture's own arithmetic"

    result = calculate(_residual_model(), TOL)
    assert result.totals.a_nom == _RESIDUAL_TOTAL, f"A_nom is {result.totals.a_nom!r}"
    assert result.totals.c_nom == _RESIDUAL_TOTAL, f"C_nom is {result.totals.c_nom!r}"
    assert result.totals.e_nom == _RESIDUAL_TOTAL, f"E_nom is {result.totals.e_nom!r}"
    assert result.totals.a_pv == _RESIDUAL_TOTAL
    assert [row.base_cost_nominal for row in result.annual] == [_RESIDUAL_TOTAL]
    assert_reconciled(result, TOL)


def test_reconciliation_cannot_catch_a_consistently_wrong_rescue() -> None:
    """§12. THE POINT OF THIS TEST IS WHAT RECONCILIATION IS NOT.

    Substituting the round-2 rounded-pair cancellation makes A, C, E and the
    annual series all wrong — and all wrong in the SAME way, because they are the
    same algorithm applied to the same contributions. Every identity still holds,
    `assert_reconciled` still passes, and the model still reports a result.

    So reconciliation verifies consistency BETWEEN calculation paths. It is not an
    independent numerical-accuracy oracle, and it can never be one. Only a
    fixture that compares against the exact mathematical value catches this class,
    which is why the assertion above is on the calculated value and not on the
    identities.
    """
    def rounded_pair_sum(terms):
        """The round-2 tier 2, verbatim."""
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
        if not remaining:
            return 0.0
        total = 0.0
        for magnitude, _ in remaining:
            total = total + magnitude
        return total if positives else -total

    module = sys.modules[calculate.__module__]
    genuine = module.safe_signed_sum

    def rescue(terms, where="sum", labels=None):
        values = list(terms)
        try:
            total = 0.0
            for value in values:
                total = module.safe_accumulate(total, value, where)
            return total
        except NumericalRangeRefusal:
            return rounded_pair_sum(values)

    module.safe_signed_sum = rescue
    try:
        broken = calculate(_residual_model(), TOL)
        # The sabotage does NOT make the model fail, and it does NOT break any
        # identity: every path is wrong the same way.
        assert_reconciled(broken, TOL)
        assert broken.totals.a_nom == broken.totals.c_nom == broken.totals.e_nom
        assert broken.totals.a_nom != _RESIDUAL_TOTAL, (
            "the sabotage must change the answer, or this control proves nothing"
        )
        assert broken.totals.a_nom == -1.99792015476736e292, broken.totals.a_nom
    finally:
        module.safe_signed_sum = genuine

    assert calculate(_residual_model(), TOL).totals.a_nom == _RESIDUAL_TOTAL


def test_a_headline_total_beyond_range_by_less_than_one_ulp_is_refused() -> None:
    """REPRODUCER §1.3, end to end. The exact total exceeds `MAX_DOUBLE` by about
    half an ulp, so it rounds to `MAX_DOUBLE` — and returning that would be the
    fabricated value C2 forbids."""
    terms = (-8e307, -7e307, -1.78e308, 5e307, -1e292, 1e308, -MAX_DOUBLE, 1.78e308)
    exact = sum((Fraction(term) for term in terms), Fraction(0))
    assert abs(exact) > Fraction(MAX_DOUBLE)
    assert abs(exact) - Fraction(MAX_DOUBLE) < Fraction(2) ** 971

    message = _refuses(
        lambda: calculate(
            _model(
                discount=0.0,
                costs=tuple(
                    _degenerate_cost(f"CL-{index + 1:03d}", value)
                    for index, value in enumerate(terms)
                ),
            ),
            TOL,
        ),
        "a headline total outside Double range by under one ulp",
    )
    assert "totals" in message


# ---------------------------------------------------------------------------
# Round 5 - the materialization boundary
# ---------------------------------------------------------------------------
# A representability boundary sits at a NAMED, MATERIALIZED Phase-5 value, not at
# whatever subexpression the implementation happens to assign to a local variable.
# `Knom`, `Kpv`, every per-driver audit amount, each of the six annual columns and
# each headline total is published, so each must be a usable Double. `w * infl`,
# the pre-FX sum, and one driver's contribution to one annual row are published
# nowhere, so none of them is.
_HALF_MAX = MAX_DOUBLE / 2


def _profile_cost(
    permanent_id: str, value: float, weights: tuple[float, ...],
    currency: str = "SAR", profile: str = "P",
) -> CostDriver:
    return CostDriver(
        permanent_id, "Uniform", currency, profile, value, None, value, weights, quantity=1
    )


def test_knom_is_not_refused_for_a_pre_fx_intermediate() -> None:
    """REPRODUCER A. Profile `[2, -1]`, inflation factor `MAX_DOUBLE`, FX `0.5`.

    `Knom = 0.5 * (2*MAX - 1*MAX) = MAX/2`. The current staging forms `2 * MAX`
    before FX is applied, and that intermediate has no Double — but it is not a
    `_Calc` field. Every value the model publishes is representable.
    """
    model = _model(
        base=2025, start=2026, duration=2, discount=0.0,
        fx=(FxRow("SAR", 1), FxRow("X", 0.5)),
        rates={"P": {2026: MAX_DOUBLE, 2027: 0.0}},
        costs=(_profile_cost("CL-001", 1.0, (2.0, -1.0), currency="X"),),
    )
    result = calculate(model, TOL)
    driver = result.drivers[0]

    assert driver.knom == _HALF_MAX, f"Knom is {driver.knom!r}"
    assert driver.kpv == _HALF_MAX, f"Kpv is {driver.kpv!r}"
    assert [row.base_cost_nominal for row in result.annual] == [MAX_DOUBLE, -_HALF_MAX]
    assert [row.base_cost_pv for row in result.annual] == [MAX_DOUBLE, -_HALF_MAX]
    assert result.totals.a_nom == result.totals.a_pv == _HALF_MAX
    assert result.totals.c_nom == result.totals.c_pv == _HALF_MAX
    assert result.totals.e_nom == result.totals.e_pv == _HALF_MAX
    assert result.totals.b_nom == result.totals.b_pv == 0.0
    assert result.totals.d_nom == result.totals.d_pv == 0.0
    assert_reconciled(result, TOL)

    # NOT VACUOUS: the intermediate the old orchestration insisted on really has
    # no Double, so this model was refused before the patch.
    _refuses(lambda: safe_product([2.0, MAX_DOUBLE]), "the pre-FX intermediate")
    assert Fraction(0.5) * (2 * Fraction(MAX_DOUBLE) - Fraction(MAX_DOUBLE)) == Fraction(
        _HALF_MAX
    ), "the fixture's own arithmetic"


def test_an_annual_row_is_not_refused_for_a_per_driver_contribution() -> None:
    """REPRODUCER B. Two cost lines at `+2` and `-2`, inflation factor `MAX_DOUBLE`.

    Each driver's per-year contribution is `±2 * MAX_DOUBLE`, which no Double
    holds. `tblCalcAnnual` publishes the aggregate, and the aggregate is `0`.
    """
    weights = (1.0, -1.0, 1.0)
    model = _model(
        base=2025, start=2026, duration=3, discount=0.0,
        rates={"P": {2026: MAX_DOUBLE, 2027: 0.0, 2028: -0.5}},
        costs=(
            _profile_cost("CL-001", 2.0, weights),
            _profile_cost("CL-002", -2.0, weights),
        ),
    )
    result = calculate(model, TOL)

    for driver in result.drivers:
        assert driver.knom == _HALF_MAX, f"{driver.permanent_id} Knom is {driver.knom!r}"
        assert driver.kpv == _HALF_MAX
    # The per-driver audit rows ARE published, and they are representable.
    assert [d.mean_basis_nominal for d in result.drivers] == [MAX_DOUBLE, -MAX_DOUBLE]
    assert [d.deterministic_nominal for d in result.drivers] == [MAX_DOUBLE, -MAX_DOUBLE]

    assert [row.base_cost_nominal for row in result.annual] == [0.0, 0.0, 0.0]
    assert [row.base_cost_pv for row in result.annual] == [0.0, 0.0, 0.0]
    assert [row.total_nominal for row in result.annual] == [0.0, 0.0, 0.0]
    assert [row.total_pv for row in result.annual] == [0.0, 0.0, 0.0]
    assert result.totals.a_nom == result.totals.c_nom == result.totals.e_nom == 0.0
    assert result.totals.b_nom == result.totals.d_nom == 0.0
    assert_reconciled(result, TOL)

    # NOT VACUOUS: the internal year-2026 contributions really do cross the Double
    # boundary, in both directions.
    assert Fraction(2) * Fraction(MAX_DOUBLE) > Fraction(MAX_DOUBLE)
    _refuses(lambda: safe_product([2.0, 1.0, 1.0, 1.0, MAX_DOUBLE]), "the +2*MAX contribution")
    _refuses(lambda: safe_product([-2.0, 1.0, 1.0, 1.0, MAX_DOUBLE]), "the -2*MAX contribution")


def test_an_annual_row_is_not_refused_for_a_contribution_that_underflows() -> None:
    """REPRODUCER C. Two identical drivers, profile `[5e-324, 1]`, factor `0.5`.

    Each first-year contribution is `0.5 * 5e-324`, which has no non-zero Double.
    The published annual row is the sum of the two, which is exactly `5e-324` — the
    smallest Double there is.
    """
    subnormal = 5e-324
    weights = (subnormal, 1.0)
    model = _model(
        base=2025, start=2026, duration=2, discount=0.0,
        rates={"P": {2026: -0.5, 2027: 1.0}},
        costs=(
            _profile_cost("CL-001", 1.0, weights),
            _profile_cost("CL-002", 1.0, weights),
        ),
    )
    result = calculate(model, TOL)

    for driver in result.drivers:
        assert driver.knom == 1.0, f"{driver.permanent_id} Knom is {driver.knom!r}"
        assert driver.kpv == 1.0
    assert [row.base_cost_nominal for row in result.annual] == [subnormal, 2.0]
    assert [row.base_cost_pv for row in result.annual] == [subnormal, 2.0]
    assert result.totals.c_nom == result.totals.a_nom == result.totals.e_nom == 2.0
    assert_reconciled(result, TOL)

    # An INDEPENDENT exact oracle for the year-2026 aggregate: two contributions of
    # 0.5 * 5e-324 each, neither of which is a Double.
    contribution = Fraction(subnormal) * Fraction(0.5)
    assert contribution != 0 and float(contribution) == 0.0
    assert float(contribution + contribution) == subnormal
    _refuses(lambda: safe_product([1.0, 1.0, 1.0, subnormal, 0.5]), "one 0.5s contribution")


def test_the_annual_rescue_covers_risk_contributions_too() -> None:
    """REPRODUCER D. The same boundary on the Risk path, and mixed with Cost.

    Base Cost, Expected Risk and Total are three separately published columns, so
    the fixture makes the CONTRIBUTIONS unrepresentable in both the Cost and the
    Risk series while each published column is finite. The Risk products carry
    `probability` where the Cost products carry `quantity`, so this exercises a
    different factor list, not just a second copy of the cost path.
    """
    weights = (2.0, -1.0)                      # sums to 1
    model = _model(
        base=2025, start=2026, duration=2, discount=0.0,
        rates={"P": {2026: MAX_DOUBLE, 2027: 0.0}},
        costs=(
            _profile_cost("CL-001", 1.0, weights),
            _profile_cost("CL-002", -1.0, weights),
        ),
        risks=(
            RiskDriver("R-001", "Uniform", "SAR", "P", 2.0, None, 2.0, weights,
                       probability=0.5),
            RiskDriver("R-002", "Uniform", "SAR", "P", -2.0, None, -2.0, weights,
                       probability=0.5),
        ),
    )
    result = calculate(model, TOL)

    for driver in result.drivers:
        assert driver.knom == MAX_DOUBLE, f"{driver.permanent_id} Knom is {driver.knom!r}"

    # Published per-driver amounts, all representable.
    costs = [d for d in result.drivers if d.driver_kind is DriverKind.COST_LINE]
    risks = [d for d in result.drivers if d.driver_kind is DriverKind.RISK]
    assert [d.mean_basis_nominal for d in costs] == [MAX_DOUBLE, -MAX_DOUBLE]
    assert [d.expected_risk_nominal for d in risks] == [MAX_DOUBLE, -MAX_DOUBLE]

    # Published annual columns, each rescued independently.
    assert [row.base_cost_nominal for row in result.annual] == [0.0, 0.0]
    assert [row.expected_risk_nominal for row in result.annual] == [0.0, 0.0]
    assert [row.total_nominal for row in result.annual] == [0.0, 0.0]
    assert [row.base_cost_pv for row in result.annual] == [0.0, 0.0]
    assert [row.expected_risk_pv for row in result.annual] == [0.0, 0.0]
    assert [row.total_pv for row in result.annual] == [0.0, 0.0]
    assert result.totals.c_nom == result.totals.d_nom == result.totals.e_nom == 0.0
    assert_reconciled(result, TOL)

    # NOT VACUOUS in EITHER series: the year-2026 contributions have no Double on
    # the cost side or the risk side.
    _refuses(lambda: safe_product([1.0, 1.0, 1.0, 2.0, MAX_DOUBLE]), "the cost contribution")
    _refuses(lambda: safe_product([0.5, 2.0, 1.0, 2.0, MAX_DOUBLE]), "the risk contribution")


def test_a_published_driver_audit_value_outside_range_is_still_refused() -> None:
    """REPRODUCER E. THE LIMIT OF THE RESCUE.

    `tblCalcDrivers` publishes each driver's own amounts, so a driver whose
    Mean-Basis Nominal is mathematically `2 * MAX_DOUBLE` cannot be calculated —
    even though a second driver at `-2` would cancel it in A, C and E, and even
    though every annual row and every headline would be zero.

    Without this the rescue would drift into "arbitrary precision until the
    workbook total", which C2 explicitly forbids.
    """
    message = _refuses(
        lambda: calculate(
            _model(
                base=2025, start=2026, duration=1, discount=0.0,
                rates={"P": {2026: MAX_DOUBLE}},
                costs=(
                    _profile_cost("CL-001", 2.0, (1.0,)),
                    _profile_cost("CL-002", -2.0, (1.0,)),
                ),
            ),
            TOL,
        ),
        "a published driver audit amount outside Double range",
    )
    assert "CL-001" in message, message

    # And the headline it would have had is perfectly representable, which is
    # exactly why this boundary has to be stated rather than inferred.
    assert Fraction(2) * Fraction(MAX_DOUBLE) - 2 * Fraction(MAX_DOUBLE) == 0


def test_the_ordinary_factor_and_annual_pipelines_are_bit_for_bit_unchanged() -> None:
    """§12. THE NON-REGRESSION THAT MATTERS.

    The compound helper must not become the default path. For a model that works,
    `Knom`, `Kpv` and every annual column must be EXACTLY what the current staged
    sequence produces — per-year `safe_product`, canonical signed sum, then FX;
    and for PV, `nominal * discount` formed from the materialized nominal.

    Exact Double equality, not a tolerance.
    """
    rates = {"Standard": {2027: 0.05, 2028: 0.05, 2029: 0.05}}
    weights = (0.2, 0.5, 0.3)
    result = calculate(
        _model(
            base=2026, start=2027, duration=3, discount=0.08, rates=rates,
            fx=(FxRow("SAR", 1), FxRow("USD", 3.75)),
            costs=(
                _cost("CL-001", weights=weights),
                _cost("CL-002", currency="USD", distribution="Beta-PERT",
                      weights=(0.5, 0.25, 0.25), quantity=3),
            ),
            risks=(_risk("R-001", weights=(0.0, 1.0, 0.0)),),
        ),
        TOL,
    )

    inflation = {
        row.calendar_year: row.cumulative_factor
        for row in result.inflation_factors
        if row.profile == "Standard"
    }
    discounts = result.discount_factors
    driver = next(d for d in result.drivers if d.permanent_id == "CL-001")

    # Knom, staged exactly as the ordinary path stages it.
    staged = 0.0
    for offset, calendar_year in enumerate((2027, 2028, 2029)):
        staged = staged + weights[offset] * inflation[calendar_year]
    assert driver.knom == 1.0 * staged, (driver.knom, staged)

    # Each annual column, staged exactly as the ordinary path stages it: the
    # per-driver product left to right, then canonical accumulation, and for PV the
    # discount applied to the MATERIALIZED nominal rather than folded in.
    for offset, row in enumerate(result.annual):
        base_nominal = 0.0
        base_present = 0.0
        for cost in [d for d in result.drivers if d.driver_kind is DriverKind.COST_LINE]:
            contribution = (
                cost.mean_value * cost.quantity * cost.fx_to_sar
                * cost.weights[offset] * inflation[row.calendar_year]
            )
            base_nominal = base_nominal + contribution
            base_present = base_present + contribution * discounts[row.project_index]
        assert row.base_cost_nominal == base_nominal, (row.calendar_year, base_nominal)
        assert row.base_cost_pv == base_present, (row.calendar_year, base_present)

    assert_reconciled(result, TOL)


def test_the_compound_helper_is_not_reached_by_an_ordinary_model() -> None:
    """§12, structurally: if the rescue ever became the default, this fails.

    The helper is replaced by one that raises, and an ordinary model must still
    calculate — proving the staged path carried it end to end.
    """
    module = sys.modules[calculate.__module__]
    genuine = module.exact_sum_of_products

    def forbidden(groups, where):
        raise AssertionError(f"the compound rescue was reached for {where}")

    module.exact_sum_of_products = forbidden
    try:
        result = calculate(
            _model(
                base=2026, start=2027, duration=3, discount=0.08, rates=_three_year(),
                fx=(FxRow("SAR", 1), FxRow("USD", 3.75)),
                costs=(_cost("CL-001", weights=(0.2, 0.5, 0.3)),
                       _cost("CL-002", currency="USD", weights=(0.5, 0.25, 0.25))),
                risks=(_risk("R-001", weights=(0.34, 0.33, 0.33)),),
            ),
            TOL,
        )
        assert_reconciled(result, TOL)
    finally:
        module.exact_sum_of_products = genuine


# ---------------------------------------------------------------------------
# runner
# ---------------------------------------------------------------------------
def _run_all() -> int:
    tests = sorted(
        (name, fn) for name, fn in globals().items()
        if name.startswith("test_") and callable(fn)
    )
    failures = 0
    print("PCCM Phase 5 Gate-A Step-2 analytical oracle tests")
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
