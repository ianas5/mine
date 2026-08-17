"""PCCM Phase-5 numerical primitives — the safe-arithmetic reference.

The lower of the oracle's two layers, and the reference for what later VBA
`modCalcFactors` must do. Everything here is a **pure function of plain numbers**:
no workbook, no Excel object model, no file I/O, no randomness, no state.

--------------------------------------------------------------------------------
WHY SAFE PRIMITIVES EXIST AT ALL
--------------------------------------------------------------------------------
A finiteness predicate applied *after* an operation cannot catch a failure that
prevents its own operand from existing. In VBA

    x = a * b

raises `Overflow` **before** `x` exists, so `IsUsableDouble(x)` never runs. The
primitives below therefore wrap each single arithmetic operation and convert a
representational failure into a controlled refusal that names the stage.

Python does not raise on float overflow — it produces `inf` — so these functions
check the result explicitly. The observable contract is what must match VBA, not
the mechanism:

    finite representable result        -> value returned
    genuinely unrepresentable result   -> NumericalRangeRefusal
    zero divisor                       -> NumericalRangeRefusal
    quotient/product underflows to
      exactly zero from non-zero
      operands                         -> NumericalRangeRefusal
    NaN or infinity                    -> never returned, never accepted

IEEE-754 `float` IS the Double semantic reference. Arbitrary-precision `Decimal`
is deliberately NOT used as the calculation engine: the model must behave the way
Excel and VBA behave, and computing something more accurate than the target would
hide exactly the representability failures this module exists to surface.
`Decimal` and `Fraction` are legitimate only in independent test oracles.
"""

from __future__ import annotations

import math
from typing import Iterable, Sequence

MAX_DOUBLE = 1.7976931348623157e308
"""The largest finite IEEE-754 double. Named because §19.5 states it as the bound
of `IsUsableDouble`, not because Python needs the constant."""


# ---------------------------------------------------------------------------
# Failure vocabulary
# ---------------------------------------------------------------------------
# Defined at the LOWER layer so the analytical layer can import it without a
# circular dependency. The numerical layer only ever raises
# `NumericalRangeRefusal`; the other classes are declared here so there is exactly
# one hierarchy rather than two that later have to be reconciled.
class OracleError(Exception):
    """Base of every failure the Phase-5 oracle can report."""


class CalculationRefusal(OracleError):
    """A refusal: this model cannot be calculated, and no result is produced.

    Everything under this class is destined for the future `PCCM_Calculate`
    REFUSED path. It says nothing is wrong with the implementation - the inputs,
    or their representability, make a correct answer impossible.
    """


class ModelInputRefusal(CalculationRefusal):
    """Invalid user or model input.

    A currency with no rate, a profile that does not sum to 1, a Quantity of zero,
    a missing inflation year. The user can fix these; the message must say which
    subject is at fault - permanent ID, currency, profile, calendar year or
    project year.
    """


class NumericalRangeRefusal(CalculationRefusal):
    """A representability limit, not a mistake by the user.

    An overflow, a division by zero, or a factor that collapsed to exactly zero.
    The inputs may be entirely reasonable and still have no representable answer.
    """


class OracleInvariantError(OracleError):
    """The implementation disagrees with itself.

    A reconciliation identity that fails is NOT a business-input refusal: the
    inputs were accepted, the calculation ran, and two independently accumulated
    quantities that must agree do not. That is a defect in the calculation, and it
    must never be reported to a user as though their model were invalid.
    """


# ---------------------------------------------------------------------------
# The predicate
# ---------------------------------------------------------------------------
def is_usable_double(value: object) -> bool:
    """`IsUsableDouble` — not NaN, not +/-infinity, within finite Double range.

    A backstop applied to every value before it leaves the kernel, not the
    mechanism that catches failures (§19.5).
    """
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return False
    number = float(value)
    if math.isnan(number) or math.isinf(number):
        return False
    return abs(number) <= MAX_DOUBLE


def _require_operands(where: str, *values: float) -> tuple[float, ...]:
    for value in values:
        if not is_usable_double(value):
            raise NumericalRangeRefusal(
                f"{where}: operand {value!r} is not a usable Double "
                "(NaN, infinite, or outside finite Double range)"
            )
    return tuple(float(v) for v in values)


def _require_result(where: str, result: float, operation: str) -> float:
    if not is_usable_double(result):
        raise NumericalRangeRefusal(
            f"{where}: {operation} produced a result outside finite Double range"
        )
    return result


# ---------------------------------------------------------------------------
# Safe arithmetic primitives
# ---------------------------------------------------------------------------
def safe_add(a: float, b: float, where: str = "addition") -> float:
    left, right = _require_operands(where, a, b)
    return _require_result(where, left + right, "addition")


def safe_subtract(a: float, b: float, where: str = "subtraction") -> float:
    left, right = _require_operands(where, a, b)
    return _require_result(where, left - right, "subtraction")


def safe_multiply(a: float, b: float, where: str = "multiplication") -> float:
    """Multiply, refusing overflow and refusing a silent collapse to zero.

    The underflow rule matters as much as the overflow rule: two non-zero operands
    whose product rounds to exactly zero would delete a contribution with no error
    anywhere, which is the failure mode §19.3 calls "the more dangerous case
    because it is silent".
    """
    left, right = _require_operands(where, a, b)
    result = left * right
    _require_result(where, result, "multiplication")
    if result == 0.0 and left != 0.0 and right != 0.0:
        raise NumericalRangeRefusal(
            f"{where}: multiplication of two non-zero values underflowed to exactly zero"
        )
    return result


def safe_divide(a: float, b: float, where: str = "division") -> float:
    """Divide, refusing a zero divisor outright rather than relying on an error."""
    numerator, divisor = _require_operands(where, a, b)
    if divisor == 0.0:
        raise NumericalRangeRefusal(f"{where}: division by zero")
    result = numerator / divisor
    _require_result(where, result, "division")
    if result == 0.0 and numerator != 0.0:
        raise NumericalRangeRefusal(
            f"{where}: quotient of a non-zero value underflowed to exactly zero"
        )
    return result


def safe_accumulate(accumulator: float, term: float, where: str = "accumulation") -> float:
    """Add one term to a running total, checked at THIS term.

    Checking during accumulation rather than at the end is what lets the caller
    name the driver, profile or year responsible, instead of reporting that a
    total came out infinite (§19.4).
    """
    return safe_add(accumulator, term, where)


def safe_sum(terms: Iterable[float], where: str = "sum") -> float:
    total = 0.0
    for index, term in enumerate(terms):
        total = safe_accumulate(total, term, f"{where}[{index}]")
    return total


# ---------------------------------------------------------------------------
# Stable products
# ---------------------------------------------------------------------------
def safe_product(factors: Sequence[float], where: str = "product") -> float:
    """A short product, evaluated so that a representable answer is not lost to a
    gratuitous intermediate overflow.

    TWO TIERS, deliberately:

    1. **Left to right first.** For every ordinary model this succeeds, and it is
       bit-for-bit what a naive implementation produces. Nothing about existing
       results changes.
    2. **Only if that fails**, re-evaluate in a magnitude-balanced order: start
       from `1.0` and repeatedly take the smallest remaining magnitude while the
       running product is `>= 1`, the largest otherwise. `1e308 * 10 * 0.01` then
       evaluates as `1e308 * 0.01 * 10 = 1e307` instead of overflowing at the
       first step on a product that is perfectly representable.

    The alternative to tier 2 is refusing a valid calculation, so it is only ever
    reached where the choice is "some answer" versus "no answer".

    NUMERICAL EDGE, STATED RATHER THAN HIDDEN: floating-point multiplication is
    commutative but NOT associative, so a reordered evaluation can differ from the
    left-to-right one in the last unit in the last place. Tier 2 therefore runs
    only when tier 1 has already failed to produce any value at all, and its order
    is fully deterministic (ascending magnitude, stable) so the same inputs always
    give the same answer. A later VBA implementation must reproduce BOTH tiers,
    in this order, to match.
    """
    if not factors:
        return 1.0
    values = _require_operands(where, *factors)

    if any(value == 0.0 for value in values):
        # An exact zero makes the product exactly zero; no ordering can change
        # that, and the underflow rule must not fire on a genuine zero input.
        negatives = sum(1 for value in values if math.copysign(1.0, value) < 0)
        return -0.0 if negatives % 2 else 0.0

    try:
        result = 1.0
        for index, value in enumerate(values):
            result = safe_multiply(result, value, f"{where}[{index}]")
        return result
    except NumericalRangeRefusal:
        pass  # tier 2

    sign = -1.0 if sum(1 for value in values if value < 0) % 2 else 1.0
    magnitudes = sorted(abs(value) for value in values)
    low, high = 0, len(magnitudes) - 1
    result = 1.0
    while low <= high:
        if result >= 1.0:
            factor, low = magnitudes[low], low + 1
        else:
            factor, high = magnitudes[high], high - 1
        result = safe_multiply(result, factor, f"{where} (magnitude-balanced)")
    return safe_multiply(result, sign, f"{where} (sign)")


# ---------------------------------------------------------------------------
# Stable distribution statistics
# ---------------------------------------------------------------------------
# Each division is applied BEFORE accumulation, so a numerator that would overflow
# is never formed. With Min = ML = Max = 1e308 the naive numerator is 3e308 and
# there is no result at all; the stable form returns exactly 1e308 (§19.2).
def triangular_mean(minimum: float, most_likely: float, maximum: float) -> float:
    """`Min/3 + ML/3 + Max/3` — never `(Min + ML + Max) / 3`."""
    where = "triangular mean"
    total = safe_accumulate(0.0, safe_divide(minimum, 3.0, where), where)
    total = safe_accumulate(total, safe_divide(most_likely, 3.0, where), where)
    return safe_accumulate(total, safe_divide(maximum, 3.0, where), where)


def beta_pert_mean(minimum: float, most_likely: float, maximum: float) -> float:
    """`Min/6 + ML*(2/3) + Max/6` — never `(Min + 4*ML + Max) / 6`.

    `ML * (2/3)` rather than `(4*ML)/6`: forming `4*ML` first is the same
    avoidable overflow the stable form exists to prevent.
    """
    where = "Beta-PERT mean"
    total = safe_accumulate(0.0, safe_divide(minimum, 6.0, where), where)
    total = safe_accumulate(total, safe_multiply(most_likely, 2.0 / 3.0, where), where)
    return safe_accumulate(total, safe_divide(maximum, 6.0, where), where)


def midpoint(minimum: float, maximum: float) -> float:
    """`Min/2 + Max/2` — never `(Min + Max) / 2`."""
    where = "midpoint"
    total = safe_accumulate(0.0, safe_divide(minimum, 2.0, where), where)
    return safe_accumulate(total, safe_divide(maximum, 2.0, where), where)


# ---------------------------------------------------------------------------
# Iterative factor series
# ---------------------------------------------------------------------------
def compound_inflation_factors(
    base_year: int,
    last_year: int,
    rate_for_year: dict[int, float],
    profile_name: str,
) -> dict[int, float]:
    """Cumulative inflation factors for `base_year .. last_year`, iteratively.

    ```
    infl(BaseYear) = 1
    infl(Y)        = infl(Y-1) * (1 + rate_Y)
    ```

    Never a power: `(1+r)**n` can overflow as an intermediate where the iteration
    detects the exact year that fails, and the year is what the refusal must name.

    The Base-Year factor of `1` is included in the result, so the audit can explain
    every project-year factor without re-deriving anything.
    """
    factors = {base_year: 1.0}
    running = 1.0
    for year in range(base_year + 1, last_year + 1):
        rate = rate_for_year[year]
        where = f"inflation profile {profile_name!r}, calendar year {year}"
        growth = safe_add(1.0, rate, where)
        if growth <= 0.0:
            raise ModelInputRefusal(
                f"{where}: 1 + rate is {growth!r}; an inflation rate of -100% or lower "
                "collapses the price base and is refused"
            )
        try:
            running = safe_multiply(running, growth, where)
        except NumericalRangeRefusal as failure:
            # Re-raised with the domain subject: the refusal must name the profile
            # and the calendar year, not merely "a multiplication".
            raise NumericalRangeRefusal(
                f"{where}: cumulative inflation factor is not representable ({failure})"
            ) from failure
        if running == 0.0:  # defensive backstop; safe_multiply already refuses this
            raise NumericalRangeRefusal(
                f"{where}: cumulative inflation factor underflowed to exactly zero"
            )
        factors[year] = running
    return factors


def discount_factor_series(discount_rate: float, duration: int) -> dict[int, float]:
    """Discount factors by PROJECT-YEAR INDEX, iteratively.

    ```
    disc(1) = 1                     project year 1 is period 0
    disc(t) = disc(t-1) / (1 + r)
    ```

    Mathematically `1 / (1+r)**(t-1)`, but the power is deliberately NOT the
    implementation path: it can overflow as an intermediate even where the
    reciprocal is perfectly representable, and it cannot say which year failed.

    A factor that reaches exactly zero is REFUSED, never accepted as "a very small
    number": a zero factor silently deletes an entire year's PV contribution.
    """
    growth = safe_add(1.0, discount_rate, "discount rate")
    if growth <= 0.0:
        raise ModelInputRefusal(
            f"discount rate: 1 + r is {growth!r}; a discount rate of -100% or lower is refused"
        )
    factors = {1: 1.0}
    running = 1.0
    for index in range(2, duration + 1):
        where = f"discount factor, project year {index}"
        try:
            running = safe_divide(running, growth, where)
        except NumericalRangeRefusal as failure:
            raise NumericalRangeRefusal(
                f"{where}: discount factor collapsed and is not usable ({failure}). A zero "
                "factor would silently delete this year's entire present-value contribution, "
                "so it is refused rather than accepted as 'a very small number'."
            ) from failure
        if running == 0.0:  # defensive backstop; safe_divide already refuses this
            raise NumericalRangeRefusal(
                f"{where}: discount factor underflowed to exactly zero"
            )
        factors[index] = running
    return factors


# ---------------------------------------------------------------------------
# Cancellation-aware tolerance
# ---------------------------------------------------------------------------
def identity_allowance(
    terms: Sequence[float],
    absolute_floor: float,
    relative_coefficient: float,
    scale_floor: float = 1.0,
) -> float:
    """`max(absolute_floor, coefficient * max(scale_floor, sum |terms|))`, stably.

    The conditioning scale must reflect the MAGNITUDE OF THE ARITHMETIC PERFORMED,
    not of its net result, or a model whose large positive and negative
    contributions cancel would collapse its own tolerance to the floor and report
    ordinary accumulation error as a bookkeeping mismatch.

    THE STABILITY POINT: the raw sum `|A| + |B| + |C|` can exceed Double while the
    tolerance `1e-12 * scale` is perfectly representable. The coefficient is
    therefore DISTRIBUTED OVER THE TERMS - `c*|A| + c*|B| + c*|C|` - which is the
    same number and never forms the overflowing intermediate. Neither the `1e-6`
    floor nor the `1e-12` coefficient is loosened to achieve that.
    """
    allowance = safe_multiply(relative_coefficient, scale_floor, "conditioning floor")
    for index, term in enumerate(terms):
        scaled = safe_multiply(
            relative_coefficient, abs(float(term)), f"conditioning term[{index}]"
        )
        allowance = safe_accumulate(allowance, scaled, f"conditioning allowance[{index}]")
    return max(absolute_floor, allowance)
