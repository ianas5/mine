#!/usr/bin/env python3
"""P7-5: the annual stochastic computation, against an independent reference.

The identities here are the primary correctness anchor, not a supplement to it.
Every one of them is a RECONCILIATION - a number this module produced compared
against a number the accepted simulation produced by a different route - rather
than a comparison of the implementation with itself.

WHAT THE SCENARIOS DELIBERATELY CONTAIN. Multiple currencies, so FX is not 1;
two inflation profiles, so the per-year factors differ between drivers; a
non-zero discount rate, so nominal and PV cannot be confused; Risks whose
occurrence differs between iterations, so the zero path is exercised; and a
driver whose profile puts nothing in some years, so a year with no contribution
from a driver is a real case rather than a hypothetical.
"""

from __future__ import annotations

import random
import sys
from pathlib import Path

import pytest

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import sim_annual as annual  # noqa: E402
from pccm_builder.calc_numeric import safe_product, safe_signed_sum  # noqa: E402
from pccm_builder.sim_stats import percentile_type7  # noqa: E402

ABSOLUTE_FLOOR = 1e-6
RELATIVE_COEFFICIENT = 1e-12
SCALE_FLOOR = 1.0

DISCOUNT = (1.0, 1.0 / 1.05, 1.0 / 1.05 ** 2)
YEARS = 3


# ---------------------------------------------------------------------------
# THE ACCEPTED SCALAR FACTOR, formed the way modCalcFactors.BuildFactor forms it
# ---------------------------------------------------------------------------
# Tier 1: each w_y * infl_y (* disc_y), summed in project-year order, FX applied
# ONCE at the end and never distributed. That grouping is the whole reason the
# reconciliation below is a tolerance check rather than a bit comparison, so the
# reference reproduces it exactly rather than computing the factor some other
# way that happens to agree.
def _knom(fx: float, weights, inflation) -> float:
    terms = [safe_product([weights[y], inflation[y]]) for y in range(len(weights))]
    return safe_product([fx, safe_signed_sum(terms, "staged")])


def _kpv(fx: float, weights, inflation, discount) -> float:
    terms = [
        safe_product([weights[y], inflation[y], discount[y]]) for y in range(len(weights))
    ]
    return safe_product([fx, safe_signed_sum(terms, "staged")])


class _Driver:
    def __init__(self, permanent_id, fx, weights, inflation, is_risk=False):
        self.permanent_id = permanent_id
        self.is_risk = is_risk
        # The discount series is per PROJECT YEAR, so a driver spanning fewer
        # years takes the leading years of it. Handing a two-year driver a
        # three-year discount vector is a shape error the module refuses, and
        # the fixture must not manufacture one.
        discount = DISCOUNT[: len(weights)]
        self.knom = _knom(fx, weights, inflation)
        self.kpv = _kpv(fx, weights, inflation, discount)
        self.knom_y = annual.per_year_nominal_factors(fx, weights, inflation)
        self.kpv_y = annual.per_year_pv_factors(fx, weights, inflation, discount)


def _model():
    return [
        _Driver("CL-001", 1.00, (0.25, 0.50, 0.25), (1.0, 1.03, 1.0609)),
        _Driver("CL-002", 3.75, (0.10, 0.30, 0.60), (1.0, 1.06, 1.1236)),
        # A profile that puts NOTHING in year 3.
        _Driver("R-001", 1.00, (0.50, 0.50, 0.00), (1.0, 1.03, 1.0609), is_risk=True),
        # ...and one that puts nothing in year 1.
        _Driver("R-002", 3.75, (0.00, 0.40, 0.60), (1.0, 1.06, 1.1236), is_risk=True),
    ]


def _totals(observations, drivers, measure):
    scalar = [getattr(d, measure) for d in drivers]
    return [
        safe_signed_sum(
            [
                0.0 if row[i] == 0.0 else safe_product([row[i], scalar[i]])
                for i in range(len(scalar))
            ],
            "iteration total",
        )
        for row in observations
    ]


def _allowance(observations, drivers, measure_y, measure_s, iterations):
    annual_terms, total_terms = [], []
    for j in iterations:
        row = observations[j]
        for i, driver in enumerate(drivers):
            if row[i] == 0.0:
                continue
            total_terms.append(safe_product([row[i], getattr(driver, measure_s)]))
            for y in range(YEARS):
                annual_terms.append(safe_product([row[i], getattr(driver, measure_y)[y]]))
    return annual.reconciliation_allowance(
        annual_terms, total_terms, ABSOLUTE_FLOOR, RELATIVE_COEFFICIENT, SCALE_FLOOR
    )


def _observations(count, seed=11):
    rng = random.Random(seed)
    rows = []
    for index in range(count):
        rows.append([
            rng.uniform(100.0, 900.0),      # CL-001
            rng.uniform(20.0, 300.0),       # CL-002
            # Risks occur on some iterations and not others; a non-occurrence is
            # EXACTLY 0.0, which is what the engine records.
            0.0 if index % 3 == 0 else rng.uniform(1000.0, 9000.0),
            0.0 if index % 4 == 0 else rng.uniform(500.0, 4000.0),
        ])
    return rows


def _factors(drivers, measure_y):
    return [getattr(d, measure_y) for d in drivers]


# ===========================================================================
# A. THE PER-YEAR FACTOR IS A DECOMPOSITION
# ===========================================================================
@pytest.mark.parametrize("measure_y,measure_s", [("knom_y", "knom"), ("kpv_y", "kpv")])
def test_01_the_per_year_factors_sum_to_the_accepted_factor(measure_y, measure_s) -> None:
    """`sum_y Knom_y = Knom`, and the same for PV.

    Checked to the project's own allowance rather than to bit equality: `Knom`
    applies FX once to a summed staging and the per-year form applies it in each
    year, and floating-point multiplication does not distribute exactly over
    addition. The difference is reported when it exceeds what that grouping can
    explain.
    """
    for driver in _model():
        parts = list(getattr(driver, measure_y))
        whole = getattr(driver, measure_s)
        recombined = safe_signed_sum(parts, "factor recombination")
        allowance = annual.reconciliation_allowance(
            parts, [whole], ABSOLUTE_FLOOR, RELATIVE_COEFFICIENT, SCALE_FLOOR
        )
        assert abs(recombined - whole) <= allowance, (
            f"{driver.permanent_id}: sum_y {measure_y} = {recombined!r} but "
            f"{measure_s} = {whole!r}"
        )


def test_02_a_year_count_mismatch_is_refused_rather_than_padded() -> None:
    """A weight vector and an inflation vector of different lengths have no
    year to attach the difference to. Padding one would invent a year."""
    with pytest.raises(annual.AnnualError, match="inflation factor"):
        annual.per_year_nominal_factors(1.0, (0.5, 0.5), (1.0,))
    with pytest.raises(annual.AnnualError, match="discount factor"):
        annual.per_year_pv_factors(1.0, (0.5, 0.5), (1.0, 1.0), (1.0,))
    with pytest.raises(annual.AnnualError, match="no project years"):
        annual.per_year_nominal_factors(1.0, (), ())


# ===========================================================================
# B, C, D. PER-ITERATION RECONCILIATION
# ===========================================================================
def test_03_one_year_model_annual_equals_the_total_exactly() -> None:
    """With a single project year the decomposition is the identity, so this is
    the one case where bit equality IS available - and is therefore demanded."""
    driver = _Driver("CL-001", 3.75, (1.0,), (1.0,))
    assert driver.knom_y[0] == driver.knom
    for observation in (1000.0, -250.0, 0.5, 1e12):
        vector = annual.iteration_annual_vector([(observation, driver.knom_y)], range(1))
        total = safe_product([observation, driver.knom])
        assert vector[0] == total, f"{observation}: {vector[0]!r} != {total!r}"


@pytest.mark.parametrize("measure_y,measure_s", [("knom_y", "knom"), ("kpv_y", "kpv")])
def test_04_every_iteration_reconciles_for_cost_lines_and_risks(measure_y, measure_s) -> None:
    """`sum_y A_j(y) = Total(j)` for every iteration, nominal and PV.

    Mixed Cost Lines and Risks, with the Risks not occurring on some iterations,
    so both the multiplied path and the exact-zero path are exercised.
    """
    drivers = _model()
    observations = _observations(24)
    totals = _totals(observations, drivers, measure_s)
    factors = _factors(drivers, measure_y)
    for j, row in enumerate(observations):
        paired = [(row[d], factors[d]) for d in range(len(drivers))]
        vector = annual.iteration_annual_vector(paired, range(YEARS))
        recombined = safe_signed_sum(list(vector), "annual recombination")
        allowance = _allowance(observations, drivers, measure_y, measure_s, [j])
        assert abs(recombined - totals[j]) <= allowance, (
            f"iteration {j}: sum_y A(y) = {recombined!r}, total = {totals[j]!r}, "
            f"allowance = {allowance!r}"
        )


def test_05_a_risk_that_did_not_occur_contributes_zero_to_every_year() -> None:
    """Not "approximately zero", and not a product that happens to vanish."""
    drivers = _model()
    vector = annual.iteration_annual_vector(
        [(0.0, drivers[2].knom_y), (0.0, drivers[3].knom_y)], range(YEARS)
    )
    assert vector == (0.0, 0.0, 0.0)


def test_06_a_year_outside_the_drivers_profile_is_refused() -> None:
    driver = _Driver("CL-001", 1.0, (0.5, 0.5), (1.0, 1.0))
    with pytest.raises(annual.AnnualError, match="year 3 was asked for"):
        annual.iteration_annual_vector([(100.0, driver.knom_y)], range(3))


# ===========================================================================
# E, F. THE SELECTED-Px PROFILE
# ===========================================================================
@pytest.mark.parametrize("p", [0.5, 0.8, 0.123, 0.9])
def test_07_the_profile_reconciles_to_the_reported_percentile(p) -> None:
    """`sum_y Profile_Px(y) = Px`, with `f` strictly inside (0, 1).

    The profile is the convex blend of the two annual vectors belonging to the
    SAME order statistics that produced the total percentile, so its sum is the
    same convex blend of two iteration totals - which is the percentile.
    """
    drivers, observations = _model(), _observations(40)
    totals = _totals(observations, drivers, "knom")
    factors = _factors(drivers, "knom_y")
    position = annual.percentile_position(totals, p)
    assert 0.0 < position.fraction < 1.0, f"p={p} did not give an interior f"

    profile = annual.selected_px_profile(observations, factors, YEARS, totals, p)
    recombined = safe_signed_sum(list(profile), "profile recombination")
    reported = percentile_type7(totals, p)
    allowance = _allowance(
        observations, drivers, "knom_y", "knom", [position.lo, position.hi]
    )
    assert abs(recombined - reported) <= allowance, (
        f"p={p}: sum_y Profile = {recombined!r}, reported Px = {reported!r}"
    )


def test_08_the_blend_is_the_contracted_convex_form() -> None:
    """Term for term, `(1 - f) * A_lo(y) + f * A_hi(y)` - and not
    `A_lo(y) + f * (A_hi(y) - A_lo(y))`, which is the same real number and a
    different Double."""
    drivers, observations = _model(), _observations(40)
    totals = _totals(observations, drivers, "knom")
    factors = _factors(drivers, "knom_y")
    p = 0.8
    position = annual.percentile_position(totals, p)
    profile = annual.selected_px_profile(observations, factors, YEARS, totals, p)

    def vector_of(iteration):
        return annual.iteration_annual_vector(
            [(observations[iteration][d], factors[d]) for d in range(len(factors))],
            range(YEARS),
        )

    low, high = vector_of(position.lo), vector_of(position.hi)
    f = position.fraction
    for year in range(YEARS):
        expected = low[year] if low[year] == high[year] else (
            (1.0 - f) * low[year] + f * high[year]
        )
        assert profile[year] == expected, f"year {year + 1}"


def test_09_at_f_zero_the_profile_is_one_iterations_own_vector() -> None:
    """Exactly, and the neighbouring vector is not touched."""
    drivers, observations = _model(), _observations(40)
    totals = _totals(observations, drivers, "knom")
    factors = _factors(drivers, "knom_y")
    for p in (0.0, 1.0):
        position = annual.percentile_position(totals, p)
        assert position.fraction == 0.0
        profile = annual.selected_px_profile(observations, factors, YEARS, totals, p)
        own = annual.iteration_annual_vector(
            [(observations[position.lo][d], factors[d]) for d in range(len(factors))],
            range(YEARS),
        )
        assert profile == own, f"p={p}: the profile is not the order statistic's own vector"


def test_10_the_forbidden_definitions_are_not_what_this_produces() -> None:
    """Nearest-rank and per-year-percentile both produce DIFFERENT numbers here,
    so the scenario can tell them apart rather than merely asserting a name."""
    drivers, observations = _model(), _observations(40)
    totals = _totals(observations, drivers, "knom")
    factors = _factors(drivers, "knom_y")
    p = 0.8
    profile = annual.selected_px_profile(observations, factors, YEARS, totals, p)

    # Nearest rank: the single iteration whose total is closest to Px.
    reported = percentile_type7(totals, p)
    nearest = min(range(len(totals)), key=lambda j: abs(totals[j] - reported))
    nearest_vector = annual.iteration_annual_vector(
        [(observations[nearest][d], factors[d]) for d in range(len(factors))], range(YEARS)
    )
    assert profile != nearest_vector, "the profile coincides with nearest-rank"


# ===========================================================================
# G. THE LADDER IS NOT THE PROFILE
# ===========================================================================
def test_11_a_per_year_ladder_and_the_selected_profile_cannot_be_confused() -> None:
    """Two different objects, and this model makes the difference visible.

    A per-year percentile ladder takes the Px of each year INDEPENDENTLY, so it
    mixes different iterations in different years and does not sum to the total
    Px. The profile takes ONE pair of iterations and blends them. Constructing a
    model where the two agree would prove nothing; this one is built so they
    disagree, and so that the per-year ladder visibly fails the identity the
    profile satisfies.
    """
    drivers, observations = _model(), _observations(40)
    totals = _totals(observations, drivers, "knom")
    factors = _factors(drivers, "knom_y")
    p = 0.8

    ladder = annual.annual_ladder(observations, factors, YEARS, [p])
    per_year = tuple(rung[0] for rung in ladder)
    profile = annual.selected_px_profile(observations, factors, YEARS, totals, p)
    reported = percentile_type7(totals, p)

    assert per_year != profile, "the per-year ladder equals the profile in this model"
    # The profile reconciles; the per-year ladder does not, and by a margin far
    # outside any accumulation allowance.
    assert abs(safe_signed_sum(list(profile), "profile") - reported) <= 1e-6
    assert abs(sum(per_year) - reported) > 1.0, (
        "the per-year ladder happens to sum to the total percentile here, so "
        "this model cannot demonstrate that they are different objects"
    )


# ===========================================================================
# H. BLOCKING CHANGES MEMORY, NOT ANSWERS
# ===========================================================================
@pytest.mark.parametrize("block_width", [1, 2, 3, 5, 12, 40])
def test_12_the_blocked_ladder_equals_the_single_pass_ladder(block_width) -> None:
    drivers, observations = _model(), _observations(30)
    factors = _factors(drivers, "knom_y")
    probabilities = [0.1, 0.5, 0.8, 0.9]
    reference = annual.annual_ladder(
        observations, factors, YEARS, probabilities, block_width=YEARS
    )
    blocked = annual.annual_ladder(
        observations, factors, YEARS, probabilities, block_width=block_width
    )
    assert blocked == reference, f"block width {block_width} changed the ladder"


def test_13_a_block_width_below_one_year_is_refused() -> None:
    drivers, observations = _model(), _observations(4)
    with pytest.raises(annual.AnnualError, match="block width"):
        annual.annual_ladder(observations, _factors(drivers, "knom_y"), YEARS, [0.5],
                             block_width=0)


# ===========================================================================
# I. THE SOURCES ARE NOT MUTATED
# ===========================================================================
def test_14_no_annual_computation_reorders_or_rewrites_its_sources() -> None:
    """The retained iteration arrays keep their original order - the digest
    depends on it - and a statistic that sorted them in place would change a
    published identity as a side effect."""
    drivers, observations = _model(), _observations(30)
    factors = _factors(drivers, "knom_y")
    totals = _totals(observations, drivers, "knom")

    before_obs = [list(row) for row in observations]
    before_totals = list(totals)
    before_factors = [tuple(f) for f in factors]

    annual.annual_ladder(observations, factors, YEARS, [0.1, 0.9])
    annual.selected_px_profile(observations, factors, YEARS, totals, 0.8)
    annual.percentile_position(totals, 0.8)

    assert [list(row) for row in observations] == before_obs
    assert list(totals) == before_totals
    assert [tuple(f) for f in factors] == before_factors


# ===========================================================================
# THE TYPE-7 POSITION
# ===========================================================================
def test_15_the_position_names_the_iterations_the_percentile_came_from() -> None:
    values = [10.0, 30.0, 20.0, 40.0]
    position = annual.percentile_position(values, 0.5)
    assert (position.lo, position.hi) == (2, 1)
    assert position.lo_value == 20.0 and position.hi_value == 30.0
    assert position.fraction == 0.5
    assert percentile_type7(values, 0.5) == (
        (1.0 - position.fraction) * position.lo_value
        + position.fraction * position.hi_value
    )


def test_16_the_tie_break_is_deterministic_and_moves_no_percentile() -> None:
    """Equal totals with DIFFERENT annual shapes.

    Without a rule, which iteration owns an order statistic would depend on the
    sort. With it, the lower original iteration index owns it. The percentile
    value cannot move, because the two candidates hold the same value - and that
    is asserted here rather than argued.
    """
    values = [50.0, 10.0, 50.0, 10.0, 30.0]
    for p in (0.0, 0.25, 0.5, 0.75, 1.0):
        position = annual.percentile_position(values, p)
        assert values[position.lo] == position.lo_value
        assert values[position.hi] == position.hi_value
        # The reported percentile is unchanged by which duplicate was chosen.
        assert percentile_type7(values, p) == percentile_type7(sorted(values), p)
    # Ties resolve to the LOWER original index, every time.
    assert annual.percentile_position(values, 0.0).lo == 1   # 10.0 at index 1, not 3
    assert annual.percentile_position(values, 1.0).lo == 2   # 50.0 at index 2, not 0


def test_17_a_probability_outside_zero_to_one_is_refused() -> None:
    for bad in (-0.01, 1.01):
        with pytest.raises(annual.AnnualError, match="outside"):
            annual.percentile_position([1.0, 2.0], bad)
    with pytest.raises(annual.AnnualError, match="empty"):
        annual.percentile_position([], 0.5)


# ===========================================================================
# THE ALLOWANCE IS THE PROJECT'S OWN
# ===========================================================================
def test_18_the_allowance_is_a_maximum_and_conditions_on_contributions() -> None:
    """`max(floor, coefficient * max(scale_floor, scale))`, and the scale sums
    the CONTRIBUTIONS - ERRATUM C1 - so a year whose terms annihilate does not
    collapse the tolerance to the floor."""
    # Cancelling contributions: the aggregate is zero, the arithmetic is not.
    cancelling = [1e16, -1e16, 1.0]
    collapsed = annual.reconciliation_allowance(
        cancelling, [1.0], ABSOLUTE_FLOOR, RELATIVE_COEFFICIENT, SCALE_FLOOR
    )
    assert collapsed > ABSOLUTE_FLOOR, (
        "the allowance collapsed to the floor on a model that processed 2e16"
    )
    assert collapsed == pytest.approx(RELATIVE_COEFFICIENT * (2e16 + 2.0), rel=1e-12)
    # A tiny model stays on the floor.
    assert annual.reconciliation_allowance(
        [0.1], [0.1], ABSOLUTE_FLOOR, RELATIVE_COEFFICIENT, SCALE_FLOOR
    ) == ABSOLUTE_FLOOR
    # It is a MAXIMUM, never a sum: the floor is not added to the scaled scale.
    scaled = annual.reconciliation_allowance(
        [1e18], [0.0], ABSOLUTE_FLOOR, RELATIVE_COEFFICIENT, SCALE_FLOOR
    )
    assert scaled == RELATIVE_COEFFICIENT * 1e18
