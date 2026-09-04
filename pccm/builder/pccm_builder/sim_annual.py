#!/usr/bin/env python3
"""P7-5: the annual stochastic computation, as an independent reference.

WHAT THIS IS
------------
The Python definition of the two authoritative Phase-7 annual outputs:

  * the annual simulated distributions - a Type-7 percentile ladder per
    project year, across iterations;
  * the selected-Px annual profile - the convex Type-7 blend of the two annual
    vectors belonging to the SAME order statistics that produced the reported
    total Px.

They are different objects and this module never lets them touch. A per-year
percentile ladder does NOT sum to the total percentile and is not a profile;
the profile is not a ladder. `sim_contract.yaml` says both in as many words -
`annual_distributions.sums_to_total_percentile: false` and
`annual_distributions.is_a_selected_px_profile: false` - and the types here are
distinct so the two cannot be passed to one another's consumers.

WHAT THIS IS NOT
----------------
Not a second simulation. Nothing here samples, seeds, allocates a run id or
touches a nonce. Every annual value is a DECOMPOSITION of an accepted
simulation's own arithmetic: the same sampled observation, the same occurrence,
a different deployment factor.

Not a second inflation, FX or discount calculation either. The per-year factors
below are the per-year GROUPS that `modCalcFactors.BuildFactor` already forms on
its way to `Knom` and `Kpv` - the same `w_y * infl_y` terms, with FX applied.
The contract states the relationship as `sum_y Knom_y = Knom`, and this module
computes the parts of a sum whose whole is already accepted rather than
recomputing the whole.

THE ONE PLACE EXACTNESS CANNOT BE ASSUMED
-----------------------------------------
`Knom` is `FX * SUM_y (w_y * infl_y)`. `SUM_y Knom_y` is
`SUM_y (FX * w_y * infl_y)`. Those are the same real number and generally
DIFFERENT `Double`s: the multiplication by FX happens once in the first and
once per year in the second, and floating-point multiplication does not
distribute exactly over addition.

So the reconciliation identities are checked to the project's OWN tolerance -
`docs/phase5_plan.md` section 15, identities I3c/I4c, with the ERRATUM C1
conditioning scale computed on CONTRIBUTIONS rather than on aggregates - and not
to bit equality. Nothing is scaled, nudged or normalised to make a sum come out;
a difference outside the allowance is reported as the mismatch it is.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Sequence

from .calc_numeric import (
    is_usable_double,
    safe_product,
    safe_signed_sum,
)
from .sim_stats import SimStatsError

__all__ = [
    "AnnualError",
    "Type7Position",
    "percentile_position",
    "per_year_nominal_factors",
    "per_year_pv_factors",
    "iteration_annual_vector",
    "annual_ladder",
    "selected_px_profile",
    "reconciliation_allowance",
    "DEFAULT_BLOCK_WIDTH",
]


class AnnualError(SimStatsError):
    """An annual computation that cannot be performed as contracted."""


# THE BLOCK WIDTH.
#
# `retention.block_width_configurable` is true and the contract names no value,
# so this is the implementation's, chosen under the P7-5 instruction to take
# "the smallest straightforward bounded implementation" when the contract leaves
# it open. It is NOT a business input and is deliberately not a workbook cell:
# nothing a user can set may change an answer, and this changes none - only how
# many years are held in memory at once.
#
# 12 is one block for every project shorter than 13 years, which is every model
# this project has measured, so the ordinary case makes exactly one replay pass.
# Beyond that the passes grow as ceil(duration / 12) while peak memory stays at
# iterations * 12 doubles - about 960 KB at 10,000 iterations, against the
# 16 MB an unblocked 200-year model would hold.
DEFAULT_BLOCK_WIDTH = 12


# ===========================================================================
# 1. THE PER-YEAR FACTOR DECOMPOSITION
# ===========================================================================
def _factor_inputs(
    fx_rate: float,
    weights: Sequence[float],
    inflation: Sequence[float],
    discount: Sequence[float] | None,
    where: str,
) -> int:
    if not is_usable_double(fx_rate):
        raise AnnualError(f"{where}: the FX rate is not a usable Double")
    count = len(weights)
    if count < 1:
        raise AnnualError(f"{where}: a driver with no project years has no annual factors")
    if len(inflation) != count:
        raise AnnualError(
            f"{where}: {count} weight(s) but {len(inflation)} inflation factor(s); "
            "the per-year decomposition has no year to attach the difference to"
        )
    if discount is not None and len(discount) != count:
        raise AnnualError(
            f"{where}: {count} weight(s) but {len(discount)} discount factor(s)"
        )
    return count


def per_year_nominal_factors(
    fx_rate: float,
    weights: Sequence[float],
    inflation: Sequence[float],
    where: str = "annual nominal factor",
) -> tuple[float, ...]:
    """`Knom_y = FX * w_y * infl_y`, one per project year.

    The same three numbers `BuildFactor` multiplies together for year `y`, in
    the same order, through the same exact-product kernel. What differs is only
    that the sum is not taken: these are the terms of `Knom`, not `Knom`.
    """
    count = _factor_inputs(fx_rate, weights, inflation, None, where)
    return tuple(
        safe_product(
            [fx_rate, weights[year], inflation[year]],
            f"{where} project year {year + 1}",
        )
        for year in range(count)
    )


def per_year_pv_factors(
    fx_rate: float,
    weights: Sequence[float],
    inflation: Sequence[float],
    discount: Sequence[float],
    where: str = "annual PV factor",
) -> tuple[float, ...]:
    """`Kpv_y = FX * w_y * infl_y * disc_y`, one per project year."""
    count = _factor_inputs(fx_rate, weights, inflation, discount, where)
    return tuple(
        safe_product(
            [fx_rate, weights[year], inflation[year], discount[year]],
            f"{where} project year {year + 1}",
        )
        for year in range(count)
    )


# ===========================================================================
# 2. THE TYPE-7 POSITION, AND WHICH ITERATIONS OWN IT
# ===========================================================================
@dataclass(frozen=True)
class Type7Position:
    """Where a Type-7 percentile came from, in terms of SOURCE iterations.

    `lo` and `hi` are indices into the ORIGINAL, unsorted sequence - the
    iteration numbering the run itself used - because that is what an annual
    vector is keyed by. The sorted position is an implementation detail of the
    percentile and is not exposed.
    """

    lo: int
    hi: int
    fraction: float
    lo_value: float
    hi_value: float


def _stable_order(values: Sequence[float]) -> list[int]:
    """Ascending by value, and by ORIGINAL ITERATION INDEX where values tie.

    THE TIE-BREAK IS THE NARROW DETERMINISTIC RULE P7-5 IS PERMITTED TO ADD, and
    it is needed only because an annual PROFILE has to name a source iteration.
    The percentile itself never needed one: when two totals are equal it does
    not matter which the sort returns, because the VALUE is the same either way.
    That is why this cannot move a published number, and `test` proves it rather
    than asserting it.

    Two iterations with equal totals may still have entirely different annual
    SHAPES - the same total spread differently across years - so without a rule
    the profile would depend on sort implementation. With it, the lower original
    iteration index owns the order statistic.
    """
    return sorted(range(len(values)), key=lambda index: (values[index], index))


def percentile_position(
    values: Sequence[float], p: float, where: str = "annual percentile position"
) -> Type7Position:
    """The `lo`, `hi` and `f` of the SAME Type-7 position the percentile uses.

    `h = (n - 1) * p`, `lo = floor(h)`, `hi = min(lo + 1, n - 1)`, `f = h - lo`,
    computed exactly as `sim_stats.percentile_type7` and
    `modSimStats.SimStatsQuantileSorted` compute them. Nothing here recomputes
    the percentile VALUE; this answers only where it came from.
    """
    count = len(values)
    if count < 1:
        raise AnnualError(f"{where}: an empty sequence has no percentile position")
    if not isinstance(p, (int, float)) or isinstance(p, bool):
        raise AnnualError(f"{where}: the probability must be a number")
    p = float(p)
    if not 0.0 <= p <= 1.0:
        raise AnnualError(f"{where}: the probability {p!r} is outside [0, 1]")

    order = _stable_order(values)
    h = float(count - 1) * p
    low_index = int(h)
    if low_index < 0:
        low_index = 0
    if low_index > count - 1:
        low_index = count - 1
    high_index = min(low_index + 1, count - 1)
    fraction = h - float(low_index)
    return Type7Position(
        lo=order[low_index],
        hi=order[high_index],
        fraction=fraction,
        lo_value=float(values[order[low_index]]),
        hi_value=float(values[order[high_index]]),
    )


# ===========================================================================
# 3. ONE ITERATION'S ANNUAL VECTOR
# ===========================================================================
def iteration_annual_vector(
    contributions: Sequence[tuple[float, Sequence[float]]],
    years: Sequence[int],
    where: str = "annual vector",
) -> tuple[float, ...]:
    """`A_j(y) = sum_d observation_d * K_d_y`, for the years asked for.

    `contributions` is one entry per driver: the OBSERVATION the accepted
    simulation already made for iteration `j` - unit cost times quantity for a
    Cost Line, severity for a Risk that occurred, and exactly `0.0` for one that
    did not - paired with that driver's per-year factor vector.

    THE OBSERVATION IS AN INPUT, NEVER RECOMPUTED HERE. This module cannot
    sample and holds no generator: what it does is apply a different deployment
    factor to a number the run already produced, which is the whole of what the
    annual decomposition is.
    """
    out: list[float] = []
    for year in years:
        terms: list[float] = []
        for position, (observation, factors) in enumerate(contributions):
            if year < 0 or year >= len(factors):
                raise AnnualError(
                    f"{where}: driver {position} has {len(factors)} project year(s) "
                    f"and year {year + 1} was asked for"
                )
            if observation == 0.0:
                # A risk that did not occur contributes exactly zero to every
                # year. Forming 0 * K would be the same number but would also
                # refuse a factor that is unrepresentable in a year the driver
                # never reached.
                terms.append(0.0)
            else:
                terms.append(
                    safe_product(
                        [observation, factors[year]],
                        f"{where} driver {position} project year {year + 1}",
                    )
                )
        out.append(safe_signed_sum(terms, f"{where} project year {year + 1}"))
    return tuple(out)


# ===========================================================================
# 4. THE BLOCKED ANNUAL LADDER
# ===========================================================================
def _blocks(duration: int, block_width: int) -> list[range]:
    if duration < 1:
        raise AnnualError("an applied duration below one year has no annual output")
    if block_width < 1:
        raise AnnualError("the year block width must be at least one year")
    return [
        range(start, min(start + block_width, duration))
        for start in range(0, duration, block_width)
    ]


def annual_ladder(
    observations: Sequence[Sequence[float]],
    factors: Sequence[Sequence[float]],
    duration: int,
    probabilities: Sequence[float],
    block_width: int = DEFAULT_BLOCK_WIDTH,
    where: str = "annual ladder",
) -> tuple[tuple[float, ...], ...]:
    """A Type-7 percentile ladder per project year, computed a block at a time.

    `observations[j][d]` is driver `d`'s observation in iteration `j`;
    `factors[d][y]` is its per-year deployment factor.

    THE BLOCKING IS THE MEMORY ARCHITECTURE AND NOTHING ELSE. Each pass builds
    the iteration-by-year values for the years of ONE block, reduces them to
    that block's ladders, and discards them before the next block begins. Peak
    retention is `iterations * block_width` doubles rather than
    `iterations * duration`, and never `drivers * iterations * duration`.

    The ANSWER does not depend on the block width: a year's ladder is a function
    of that year's column alone, so any blocking that covers every year exactly
    once produces the same ladders. `test` proves that against a single-pass
    reference rather than assuming it.
    """
    from .sim_stats import percentile_type7

    iterations = len(observations)
    if iterations < 1:
        raise AnnualError(f"{where}: a run with no iterations has no annual distribution")
    ladders: list[tuple[float, ...]] = [() for _ in range(duration)]

    for block in _blocks(duration, block_width):
        years = list(block)
        # One block's worth of annual values, iteration by iteration.
        column: list[list[float]] = [[] for _ in years]
        for iteration in range(iterations):
            paired = [
                (observations[iteration][driver], factors[driver])
                for driver in range(len(factors))
            ]
            vector = iteration_annual_vector(paired, years, f"{where} iteration {iteration + 1}")
            for offset, value in enumerate(vector):
                column[offset].append(value)
        for offset, year in enumerate(years):
            ladders[year] = tuple(
                percentile_type7(column[offset], p, f"{where} project year {year + 1}")
                for p in probabilities
            )
        # THE BLOCK IS DISCARDED HERE. `column` goes out of scope with the
        # iteration of the loop, which is the entire retention claim.
        del column
    return tuple(ladders)


# ===========================================================================
# 5. THE SELECTED-Px ANNUAL PROFILE
# ===========================================================================
def selected_px_profile(
    observations: Sequence[Sequence[float]],
    factors: Sequence[Sequence[float]],
    duration: int,
    totals: Sequence[float],
    p: float,
    where: str = "selected Px annual profile",
) -> tuple[float, ...]:
    """`Profile_Px(y) = (1 - f) * A_lo(y) + f * A_hi(y)`.

    `lo`, `hi` and `f` are the SAME Type-7 position that produced the reported
    total Px - taken from `totals`, the accepted iteration totals - so the
    profile is the blend of the two annual vectors belonging to those two order
    statistics.

    FORBIDDEN AND NOT DONE HERE: nearest-rank substitution; picking whichever
    iteration is "closest" to Px; computing a percentile per year and calling
    the result a profile; scaling an arbitrary vector until it sums to Px. Only
    two annual vectors are ever built, and the blend is convex.
    """
    if len(totals) != len(observations):
        raise AnnualError(
            f"{where}: {len(totals)} total(s) for {len(observations)} iteration(s)"
        )
    position = percentile_position(totals, p, where)
    years = list(range(duration))

    def vector_of(iteration: int) -> tuple[float, ...]:
        paired = [
            (observations[iteration][driver], factors[driver])
            for driver in range(len(factors))
        ]
        return iteration_annual_vector(paired, years, f"{where} iteration {iteration + 1}")

    low = vector_of(position.lo)
    if position.fraction == 0.0:
        # AN INTEGRAL POSITION IS AN ORDER STATISTIC OUTRIGHT. Returning the
        # vector untouched rather than forming 1 * low + 0 * high is the same
        # rule `SimStatsQuantileSorted` applies to the value, and it keeps the
        # profile exactly equal to one iteration's own annual vector - which is
        # what `degenerates_to_single_iteration_when_f_is_zero` promises.
        return low
    high = vector_of(position.hi)
    out: list[float] = []
    for year in range(duration):
        if low[year] == high[year]:
            # The constant-bracket invariant, for the same reason the value
            # carries it: a convex combination of two equal numbers IS that
            # number, and 0.7 * 0.1 + 0.3 * 0.1 is not.
            out.append(low[year])
            continue
        blended = (1.0 - position.fraction) * low[year] + position.fraction * high[year]
        if not is_usable_double(blended):
            raise AnnualError(
                f"{where}: the blend for project year {year + 1} is not representable"
            )
        out.append(blended)
    return tuple(out)


# ===========================================================================
# 6. THE RECONCILIATION ALLOWANCE
# ===========================================================================
def reconciliation_allowance(
    annual_contributions: Sequence[float],
    total_contributions: Sequence[float],
    absolute_floor: float,
    relative_coefficient: float,
    scale_floor: float,
) -> float:
    """`docs/phase5_plan.md` section 15, identity I3c, with ERRATUM C1's scale.

        allowance = max(absolute_floor,
                        coefficient * max(scale_floor, conditioning_scale))

    THE SCALE IS ON CONTRIBUTIONS, NOT ON AGGREGATES, and that is the whole of
    ERRATUM C1: a year whose contributions annihilate has an aggregate of zero
    and may still have processed billions, so `sum_y |annual aggregate|` would
    collapse the tolerance exactly where floating-point error is largest.

    The inner operation is a MAXIMUM, not a sum. Adding the floor to the scaled
    scale would widen every allowance slightly, and a tolerance may not be
    loosened by accident.
    """
    scale = 0.0
    for value in annual_contributions:
        scale += abs(float(value))
    for value in total_contributions:
        scale += abs(float(value))
    if scale < scale_floor:
        scale = float(scale_floor)
    scaled = float(relative_coefficient) * scale
    allowance = scaled if scaled > absolute_floor else float(absolute_floor)
    if not is_usable_double(allowance):
        raise AnnualError("the reconciliation allowance is not representable")
    return allowance
