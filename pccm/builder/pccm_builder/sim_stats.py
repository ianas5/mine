#!/usr/bin/env python3
"""PCCM Phase 6 Step-4 statistics over retained iteration totals.

Pure numerical helpers. Nothing here knows about RNG state, drivers, contracts,
worksheets or Monte Carlo; it receives a finished sequence of Doubles and
returns statistics over them. That separation is why the mutation controls can
substitute a wrong statistic without touching the engine, and why the engine
cannot quietly acquire a statistic of its own.

--------------------------------------------------------------------------------
WHY EVERY ROUTINE IS SCALE-NORMALISED
--------------------------------------------------------------------------------
The accepted numerical domain is not narrowed by Phase 6: an iteration total may
legally sit near `Double` maximum, and totals of opposite sign near that
magnitude are legal together. Their MEAN is then perfectly representable while
their SUM is not, and their standard deviation is representable while both a
naive sum of squares and an unguarded `x - mean` are not.

So the sums are formed in a normalised space and rescaled exactly once:

    scale  = the largest power of two not exceeding max(|x|)
    scaled = x / scale                        (EXACT - see below)
    ...    accumulate, average, deviate ...
    result = (normalised result) * scale

A POWER-OF-TWO SCALE IS NOT A DETAIL. Dividing a Double by a power of two only
adjusts the exponent, so `x / scale` is exact for every value that stays in
range and is the correctly rounded quotient for one that does not. Dividing by
`max(|x|)` itself - the obvious choice - rounds every element and spends up to
an ulp of accuracy per value before any statistic has been computed. Both are
"scale aware"; only one costs nothing.

`safe_signed_sum` still performs the accumulation, so canonical left-to-right
order and its exact-rescue tier are unchanged; normalisation decides only the
SPACE the accepted primitive works in, never the order it works in.

WHAT NORMALISATION DOES NOT DO is make a long sum exact. Left-to-right
accumulation of `n` terms carries the usual `O(n * eps)` relative drift - about
`1e-13` at `n = 1000` - and nothing here re-associates or compensates to remove
it, because the accepted accumulation primitive is the one named by the
contract and a private summation algorithm here would be a second numerical
authority. Scale safety is a statement about RANGE: a statistic whose true value
is representable is produced rather than refused. It is not a claim of exactness.

A statistic whose true value has no `Double` refuses through the accepted
Phase-5 numerical hierarchy, naming the stage. Nothing here returns `inf`,
`-inf` or `NaN`, and nothing clips.

No NumPy. No `statistics` module. No worksheet function.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Mapping, Sequence

from .calc_numeric import (
    CalculationRefusal,
    NumericalRangeRefusal,
    is_usable_double,
    safe_divide,
    safe_multiply,
    safe_signed_sum,
)


class SimStatsError(CalculationRefusal):
    """A statistic that was asked for something it cannot answer.

    Derived from the accepted Phase-5 refusal hierarchy rather than invented
    alongside it: a caller already catching `CalculationRefusal` must not have to
    learn a second unrelated base to keep catching every refusal.
    """


@dataclass(frozen=True)
class MeasureStatistics:
    """Every statistic of one measure - nominal or PV - over one run."""

    count: int
    mean: float
    sample_standard_deviation: float
    minimum: float
    maximum: float
    percentiles: Mapping[str, float]

    def percentile(self, label: str) -> float:
        try:
            return self.percentiles[label]
        except KeyError:
            raise SimStatsError(
                f"percentile {label!r} was not computed for this run; the stored ladder is "
                f"{sorted(self.percentiles)}"
            ) from None


# ---------------------------------------------------------------------------
# input discipline
# ---------------------------------------------------------------------------
def _values(values: Sequence[float], where: str) -> tuple[float, ...]:
    """Every element must already be a usable Double. Nothing is coerced."""
    if isinstance(values, (str, bytes)) or not isinstance(values, Sequence):
        raise SimStatsError(f"{where}: expected a sequence of Doubles, got {type(values).__name__}")
    out: list[float] = []
    for index, value in enumerate(values):
        if not is_usable_double(value):
            raise SimStatsError(
                f"{where}[{index}]: {value!r} is not a usable Double. A statistic is never "
                "computed over a non-finite or non-numeric value."
            )
        out.append(float(value))
    return tuple(out)


def _scale_of(values: Sequence[float]) -> float:
    """The largest power of two not exceeding `max(|x|)`, or `0.0` for all zeros.

    `frexp` returns `x = m * 2**e` with `0.5 <= |m| < 1`, so `2**(e-1) <= |x|`.
    Taking the power BELOW the magnitude rather than above it is deliberate:
    `2**e` for `x` near `Double` maximum is `2**1024`, which does not exist.
    """
    largest = 0.0
    for value in values:
        magnitude = abs(value)
        if magnitude > largest:
            largest = magnitude
    if largest == 0.0:
        return 0.0
    _, exponent = math.frexp(largest)
    return math.ldexp(1.0, exponent - 1)


# ---------------------------------------------------------------------------
# mean
# ---------------------------------------------------------------------------
def sample_mean(values: Sequence[float], where: str = "sample mean") -> float:
    """The arithmetic mean, formed so that a representable mean is never lost.

    `sum(x) / n` is not implemented. `[-1.7e308, 1.7e308, 1.7e308, 1.7e308]` has a
    mean of `8.5e307`, and its running sum reaches `3.4e308` - which does not
    exist. Refusing that model would be refusing an answer it has.

    All-zero input short-circuits to `0.0`: there is no scale to normalise by,
    and the mean of zeros is zero exactly.
    """
    data = _values(values, where)
    count = len(data)
    if count == 0:
        raise SimStatsError(f"{where}: the mean of an empty sequence does not exist")

    scale = _scale_of(data)
    if scale == 0.0:
        return 0.0

    scaled = [value / scale for value in data]
    total = safe_signed_sum(scaled, f"{where}: normalised accumulation")
    normalised = safe_divide(total, float(count), f"{where}: normalised mean")
    return safe_multiply(normalised, scale, f"{where}: rescale")


# ---------------------------------------------------------------------------
# sample standard deviation
# ---------------------------------------------------------------------------
def sample_standard_deviation(
    values: Sequence[float], where: str = "sample standard deviation"
) -> float:
    """Divisor `n - 1`. Two passes, both in the normalised space.

    NEITHER FORBIDDEN PATH IS TAKEN. A naive `SUM x**2` overflows for any total
    beyond about `1.3e154` and reports `inf` for a spread that exists. An
    unguarded Welford `delta = x - mean` overflows too: for
    `[-1.7e308, 1.7e308, 1.7e308, 1.7e308]` the mean is `8.5e307` and the first
    deviation is `-2.55e308`, so the recurrence fails on its first step. In the
    normalised space every value lies in `[-2, 2]` and every deviation in
    `[-4, 4]`, so no intermediate can leave `Double` range at all, and this
    fixture returns `1.7e308`.

    `n < 2` REFUSES rather than inventing a value. A real run has `N >= 1000` by
    the business minimum, so the case only arises for a helper called directly,
    and a sample standard deviation of one observation is undefined, not zero.
    """
    data = _values(values, where)
    count = len(data)
    if count < 2:
        raise SimStatsError(
            f"{where}: a sample standard deviation needs at least two observations, got "
            f"{count}. The divisor is n - 1, which does not exist here. No value is invented."
        )

    scale = _scale_of(data)
    if scale == 0.0:
        # Every observation is exactly zero, so the spread is exactly zero. No
        # arithmetic can improve on that and normalising by zero is undefined.
        return 0.0

    scaled = [value / scale for value in data]
    total = safe_signed_sum(scaled, f"{where}: normalised accumulation")
    centre = safe_divide(total, float(count), f"{where}: normalised mean")

    squares = [(value - centre) * (value - centre) for value in scaled]
    residual = safe_signed_sum(squares, f"{where}: normalised sum of squared deviations")
    variance = safe_divide(residual, float(count - 1), f"{where}: normalised sample variance")
    if variance < 0.0:  # pragma: no cover - a sum of squares cannot be negative
        raise SimStatsError(f"{where}: the normalised sample variance is negative ({variance!r})")
    return safe_multiply(math.sqrt(variance), scale, f"{where}: rescale")


# ---------------------------------------------------------------------------
# extremes
# ---------------------------------------------------------------------------
def minimum(values: Sequence[float], where: str = "minimum") -> float:
    data = _values(values, where)
    if not data:
        raise SimStatsError(f"{where}: an empty sequence has no minimum")
    return min(data)


def maximum(values: Sequence[float], where: str = "maximum") -> float:
    data = _values(values, where)
    if not data:
        raise SimStatsError(f"{where}: an empty sequence has no maximum")
    return max(data)


# ---------------------------------------------------------------------------
# Type-7 percentile
# ---------------------------------------------------------------------------
def percentile_type7(
    values: Sequence[float], p: float, where: str = "percentile"
) -> float:
    """Hyndman-Fan type 7, exactly as the contract states it.

    ```
    h  = (n - 1) * p
    lo = floor(h)
    hi = min(lo + 1, n - 1)
    f  = h - lo
    Px = (1 - f) * x[lo] + f * x[hi]
    ```

    THE CONVEX FORM IS THE POINT. `x[lo] + f * (x[hi] - x[lo])` is the same
    number in exact arithmetic and a different one in `Double`: between
    `-1.7e308` and `1.7e308` the difference is `3.4e308`, which does not exist,
    while every convex combination of two representable endpoints is bracketed by
    them and therefore always exists.

    SORTING IS ON A COPY. The caller's sequence is never reordered - the retained
    iteration arrays keep their original order for the digest, and a statistic
    that mutated them would change the digest as a side effect.
    """
    data = _values(values, where)
    count = len(data)
    if count == 0:
        raise SimStatsError(f"{where}: an empty sequence has no percentiles")
    if isinstance(p, bool) or not isinstance(p, (int, float)):
        raise SimStatsError(f"{where}: p must be a number, got {p!r}")
    p = float(p)
    if not math.isfinite(p) or not 0.0 <= p <= 1.0:
        raise SimStatsError(f"{where}: p must lie in [0, 1], got {p!r}")

    ordered = sorted(data)
    h = (count - 1) * p
    lo = math.floor(h)
    hi = min(lo + 1, count - 1)
    f = h - lo
    low, high = ordered[lo], ordered[hi]
    if f == 0.0:
        # An integral h selects an order statistic outright. Returning it
        # untouched rather than forming 1.0 * low + 0.0 * high keeps p = 0 and
        # p = 1 exact at every magnitude, including subnormals.
        return low
    value = (1.0 - f) * low + f * high
    if not math.isfinite(value):  # pragma: no cover - convex of two finite endpoints
        raise NumericalRangeRefusal(
            f"{where}: the convex interpolation between {low!r} and {high!r} produced "
            f"{value!r} at the percentile stage"
        )
    return value


# ---------------------------------------------------------------------------
# the whole description of one measure
# ---------------------------------------------------------------------------
def describe(
    values: Sequence[float],
    points: Sequence[tuple[str, float]],
    where: str = "statistics",
) -> MeasureStatistics:
    """Every statistic of one measure, sorting exactly once.

    `points` is `(label, p)` pairs supplied by the caller. This module holds no
    ladder of its own: the selectable levels belong to `input_contract.yaml` and
    the fixed ones to `sim_contract.yaml`, and a copy here would be a third
    authority able to drift from both.
    """
    data = _values(values, where)
    if not data:
        raise SimStatsError(f"{where}: an empty run has no statistics")

    ordered = sorted(data)
    percentiles: dict[str, float] = {}
    for label, p in points:
        if label in percentiles:
            raise SimStatsError(f"{where}: percentile label {label!r} is requested twice")
        percentiles[label] = percentile_type7(ordered, p, f"{where}: {label}")

    return MeasureStatistics(
        count=len(data),
        mean=sample_mean(data, f"{where}: mean"),
        sample_standard_deviation=sample_standard_deviation(
            data, f"{where}: sample standard deviation"
        ),
        minimum=ordered[0],
        maximum=ordered[-1],
        percentiles=percentiles,
    )
