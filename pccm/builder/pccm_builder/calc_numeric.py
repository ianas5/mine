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
from typing import Callable, Iterable, Sequence

MAX_DOUBLE = 1.7976931348623157e308
"""The largest finite IEEE-754 double. Named because §19.5 states it as the bound
of `IsUsableDouble`, not because Python needs the constant."""

MIN_NORMAL_DOUBLE = 2.2250738585072014e-308
"""The smallest positive NORMAL double, `2**-1022`.

The boundary below which halving stops being exact, which is the one place the
binade rescue has to know about (see `_binade_rescue`). Nothing else in the
module needs it."""


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

    THE CONVERSION ITSELF CAN FAIL. Python's `int` is arbitrary-precision, so
    `float(10**400)` raises `OverflowError` rather than returning `inf` - and a
    predicate that raises is not a predicate. The pure oracle accepts plain Python
    numbers, so it must answer "no" for a value that cannot become a Double at
    all, and let the caller turn that into a structured refusal. Leaking a raw
    `OverflowError` would bypass the whole failure contract.
    """
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return False
    try:
        number = float(value)
    except (OverflowError, ValueError, TypeError):
        return False
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
    """Strict left-to-right accumulation, checked at every term.

    This is TIER 1 of `safe_signed_sum` and remains available on its own for
    sums that must not be rescued.
    """
    total = 0.0
    for index, term in enumerate(terms):
        total = safe_accumulate(total, term, f"{where}[{index}]")
    return total


def safe_signed_sum(
    terms: Sequence[float], where: str = "sum", labels: Sequence[str] | None = None
) -> float:
    """A signed sum whose FINAL value is what matters — two tiers.

    **Tier 1 — the canonical order, unchanged.** Accumulate left to right with
    `safe_accumulate`, naming the exact term that fails. If this succeeds, its
    result is returned bit-for-bit. **A sum that already works is never
    reordered**, so no ordinary model's numbers move, and canonical
    permanent-ID order remains what defines ordinary evaluation.

    **Tier 2 — cancellation rescue, only after tier 1 overflowed.** A partial sum
    can exceed Double while the true total is perfectly representable:
    `MAX + MAX - MAX` is `MAX`, but left to right the first addition is already
    infinite. Refusing that is refusing an answer that exists (§19.2).

    The rescue is deterministic and reproducible in VBA:

    1. every term is validated as a usable finite Double;
    2. terms are split into positive and negative MAGNITUDES, keeping the
       original canonical index as a tie-breaker, with exact zeros discarded;
    3. each bucket is ordered by magnitude;
    4. the largest opposite-signed magnitudes are repeatedly cancelled:
       equal magnitudes annihilate exactly, otherwise the residual `|p - n|` is
       re-inserted on the side that was larger. **This step cannot re-create the
       overflow being rescued**: it subtracts two non-negative magnitudes, and the
       residual never exceeds the larger operand;
    5. once one sign is exhausted, the remaining same-sign magnitudes are summed
       smallest-first. If THAT overflows, the true signed total genuinely exceeds
       Double range and a refusal is correct;
    6. the surviving sign is applied.

    Tier 2 is only reached where tier 1 produced no value at all, so the choice is
    between an answer and no answer. It is NOT invoked for underflow: a sum cannot
    underflow to zero except by genuine cancellation, which is a real result.

    `labels` name the terms in tier-1 messages, so a refusal can still say which
    driver or year was responsible.
    """
    values = _require_operands(where, *terms) if terms else ()

    def name(index: int) -> str:
        if labels is not None and index < len(labels):
            return f"{where}: {labels[index]}"
        return f"{where}[{index}]"

    try:
        total = 0.0
        for index, value in enumerate(values):
            total = safe_accumulate(total, value, name(index))
        return total
    except NumericalRangeRefusal:
        pass  # tier 2

    # --- tier 2: deterministic opposite-sign cancellation --------------------
    # `(magnitude, canonical index)` tuples sort by magnitude first and by the
    # ORIGINAL canonical position second, so equal magnitudes have one and only
    # one ordering. Nothing here depends on Python's sort being stable.
    positives = sorted((abs(v), i) for i, v in enumerate(values) if v > 0.0)
    negatives = sorted((abs(v), i) for i, v in enumerate(values) if v < 0.0)

    while positives and negatives:
        p_magnitude, p_index = positives.pop()          # largest positive
        n_magnitude, n_index = negatives.pop()          # largest negative
        if p_magnitude == n_magnitude:
            continue                                    # exact annihilation
        if p_magnitude > n_magnitude:
            positives.append((p_magnitude - n_magnitude, p_index))
            positives.sort()
        else:
            negatives.append((n_magnitude - p_magnitude, n_index))
            negatives.sort()

    remaining = positives if positives else negatives
    if not remaining:
        return 0.0                                      # everything annihilated
    sign = 1.0 if positives else -1.0
    total = 0.0
    for magnitude, index in remaining:                  # smallest magnitude first
        total = safe_accumulate(total, magnitude, f"{where} (cancellation rescue)[{index}]")
    return total if sign > 0.0 else -total


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
# Each of these is a CONVEX COMBINATION of its points: the weights are positive
# and sum to exactly 1, so the true value always lies between Min and Max. That
# is the whole justification for the machinery below — a statistic that is
# mathematically bracketed by two representable Doubles has a representable
# answer, and refusing to produce one is a defect, not caution (§19.2).
#
# THREE TIERS, in this order:
#
#   Tier 0 — degenerate invariant. If every point is the same number, the
#            distribution has zero uncertainty and the statistic IS that number.
#            Returned exactly, with no arithmetic at all, so no last-ulp drift is
#            possible anywhere in the Double range (including subnormals, where
#            `x/3 + x/3 + x/3 != x`).
#   Tier 1 — the accepted stable form, unchanged: divide each point by its
#            denominator BEFORE accumulating, so an oversized numerator is never
#            formed. Every ordinary model lands here and its bits do not move.
#   Tier 2 — binade rescue, reached ONLY when tier 1 raised. Scale every point by
#            one shared power of two, evaluate the same formula, scale the result
#            back by the same power of two.
#
# Why a power of two: multiplying or dividing a Double by 2 is exact in IEEE-754
# (it adjusts the exponent and leaves the significand alone) until the subnormal
# range, so the rescue introduces no error of its own. It also needs nothing more
# exotic than a counting loop, which is why it is reproducible in VBA — see
# `docs/phase5_gate_a_step2.md`. `frexp`/`ldexp` are deliberately NOT used.
def _degenerate_point(values: tuple[float, ...]) -> float | None:
    """The single point of a zero-uncertainty distribution, or `None`."""
    first = values[0]
    for value in values[1:]:
        if value != first:
            return None
    return first


def _binade_shift(biggest: float) -> int:
    """Halvings needed to bring `biggest` into `[1, 2)`; negative means doublings.

    A counting loop, not `frexp`, so the VBA translation is the same loop.
    """
    shifts = 0
    while biggest >= 2.0:
        biggest = biggest / 2.0
        shifts += 1
    while biggest < 1.0:
        biggest = biggest * 2.0
        shifts -= 1
    return shifts


def _binade_rescue(
    values: tuple[float, ...],
    where: str,
    formula: "Callable[[tuple[float, ...]], float]",
) -> float:
    """Evaluate a convex combination in a binade where it cannot overflow.

    The points are scaled so the LARGEST magnitude sits in `[1, 2)`. Every weight
    is at most 1 and they sum to 1, so the scaled statistic is bounded by 2 and the
    numerator `Min + 4*ML + Max` by 12: the intermediate arithmetic that defeated
    tier 1 cannot recur.

    That is why `formula` is the STRAIGHTFORWARD form — sum, then divide once —
    rather than the divide-first stable form. Dividing first exists only to keep a
    numerator inside Double range, which in this binade is not in question, and it
    costs real accuracy here: with points that nearly cancel, three separately
    rounded quotients differ from the exact result by far more than one rounding
    of their exact sum. It is also evaluated with plain operators rather than the
    refusing primitives, because a scaled point can be subnormal and
    `subnormal / 3` may round to zero even though its contribution is far below
    the last bit of the answer — refusing there would reintroduce the very defect
    being repaired.

    Scaling DOWN can flush a hugely smaller point to zero. That is not a lost
    contribution: a point more than 2^1074 times smaller than the largest cannot
    change any bit of a convex combination of them.

    Scaling the result BACK is where a genuine range failure is reported. If
    doubling overflows, the true statistic really does exceed Double range; if
    halving collapses a non-zero value to zero, the true statistic really has no
    usable non-zero Double. Both are correct refusals, and are the distinction
    §19.2 draws against an intermediate that merely stepped outside the range.
    """
    biggest = max(abs(value) for value in values)
    shifts = _binade_shift(biggest)

    scaled = list(values)
    for _ in range(abs(shifts)):
        if shifts > 0:
            scaled = [value / 2.0 for value in scaled]      # exact; may flush
        else:
            scaled = [value * 2.0 for value in scaled]      # exact; cannot overflow

    result = _require_result(where, formula(tuple(scaled)), "rescued convex combination")
    if result == 0.0:
        return 0.0

    if shifts > 0:
        for _ in range(shifts):
            result = safe_multiply(result, 2.0, f"{where} (binade rescue)")
        return result

    # Scaling back DOWN, where a naive repeated halving would round twice. Halving
    # is exact only while the value stays normal; each step taken inside the
    # subnormal range rounds again, and two roundings can land a bit below the
    # correctly-rounded answer -- far enough, at the bottom of the range, to turn a
    # representable statistic into a spurious "underflowed to zero".
    #
    # So: halve one step at a time while that is exact, then perform every
    # remaining step as ONE division, which rounds once. The single divisor is
    # always small: the exact loop cannot stop above `2**-1021`, so at most ~53
    # steps can remain before the true answer is below half the smallest subnormal.
    remaining = -shifts
    while remaining > 0 and abs(result) / 2.0 >= MIN_NORMAL_DOUBLE:
        result = result / 2.0                               # exact: still normal
        remaining -= 1
    if remaining > 0:
        if remaining > 1023:                                # cannot arise; see above
            raise NumericalRangeRefusal(
                f"{where} (binade rescue): the statistic underflowed to exactly zero"
            )
        divisor = 1.0
        for _ in range(remaining):
            divisor = divisor * 2.0                         # exact power of two
        result = result / divisor                           # ONE rounding
        if result == 0.0:
            raise NumericalRangeRefusal(
                f"{where} (binade rescue): the statistic underflowed to exactly zero"
            )
    return result


def _convex_mean(
    values: tuple[float, ...],
    where: str,
    stable: "Callable[[tuple[float, ...]], float]",
    formula: "Callable[[tuple[float, ...]], float]",
) -> float:
    """Tier 0 -> tier 1 -> tier 2, in that order. See the block comment above."""
    numbers = _require_operands(where, *values)
    point = _degenerate_point(numbers)
    if point is not None:
        return point
    try:
        return stable(numbers)
    except NumericalRangeRefusal:
        pass
    return _binade_rescue(numbers, where, formula)


def triangular_mean(minimum: float, most_likely: float, maximum: float) -> float:
    """`Min/3 + ML/3 + Max/3` — never `(Min + ML + Max) / 3`."""
    where = "triangular mean"

    def stable(v: tuple[float, ...]) -> float:
        total = safe_accumulate(0.0, safe_divide(v[0], 3.0, where), where)
        total = safe_accumulate(total, safe_divide(v[1], 3.0, where), where)
        return safe_accumulate(total, safe_divide(v[2], 3.0, where), where)

    def formula(v: tuple[float, ...]) -> float:
        return (v[0] + v[1] + v[2]) / 3.0

    return _convex_mean((minimum, most_likely, maximum), where, stable, formula)


def beta_pert_mean(minimum: float, most_likely: float, maximum: float) -> float:
    """`Min/6 + ML*(2/3) + Max/6` — never `(Min + 4*ML + Max) / 6`.

    `ML * (2/3)` rather than `(4*ML)/6`: forming `4*ML` first is the same
    avoidable overflow the stable form exists to prevent.
    """
    where = "Beta-PERT mean"

    def stable(v: tuple[float, ...]) -> float:
        total = safe_accumulate(0.0, safe_divide(v[0], 6.0, where), where)
        total = safe_accumulate(total, safe_multiply(v[1], 2.0 / 3.0, where), where)
        return safe_accumulate(total, safe_divide(v[2], 6.0, where), where)

    def formula(v: tuple[float, ...]) -> float:
        return (v[0] + 4.0 * v[1] + v[2]) / 6.0

    return _convex_mean((minimum, most_likely, maximum), where, stable, formula)


def midpoint(minimum: float, maximum: float) -> float:
    """`Min/2 + Max/2` — never `(Min + Max) / 2`."""
    where = "midpoint"

    def stable(v: tuple[float, ...]) -> float:
        total = safe_accumulate(0.0, safe_divide(v[0], 2.0, where), where)
        return safe_accumulate(total, safe_divide(v[1], 2.0, where), where)

    def formula(v: tuple[float, ...]) -> float:
        return (v[0] + v[1]) / 2.0

    return _convex_mean((minimum, maximum), where, stable, formula)


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
def scaled_magnitude(
    accumulator: float, term: float, relative_coefficient: float, where: str
) -> float:
    """Accumulate `coefficient * |term|` into a running conditioning magnitude.

    The coefficient is DISTRIBUTED OVER THE TERMS rather than applied to their
    sum. The two are the same number, but the raw sum of contributions can exceed
    Double while `1e-12 x sum` is perfectly representable, and only the
    distributed form avoids forming that intermediate.

    Called while the corresponding contribution is being accumulated, so the
    conditioning magnitude is a by-product of the calculation rather than a second
    pass that could disagree with it.

    --------------------------------------------------------------------------
    UNDERFLOW POLICY HERE IS DELIBERATELY DIFFERENT FROM MODEL ARITHMETIC
    --------------------------------------------------------------------------
    `safe_multiply` refuses a non-zero product that rounds to exactly zero,
    because an economic value or a factor that silently collapses would delete a
    real contribution with no error anywhere. That rule is right for MODEL
    arithmetic and wrong here.

    This is INTERNAL TOLERANCE-SCALING METADATA, and losing a scaled term to
    underflow cannot change the answer it feeds:

      * `coefficient * |term|` only underflows when `|term|` is below roughly
        `5e-312`;
      * the conditioning scale has a floor of `scale_floor` (locked at 1), so the
        relative allowance is at least `coefficient * 1 = 1e-12`;
      * a dropped term is therefore at most about `5e-324` against a quantity of
        at least `1e-12` - more than three hundred orders of magnitude below the
        value it would have to move. Even in a model mixing huge and tiny terms,
        the dropped amount is far under one ulp of the sum.

    Refusing here would reject a model whose economic outputs are perfectly
    representable, purely because the tolerance bookkeeping could not represent a
    term too small to matter. Overflow is still refused: a conditioning scale that
    exceeds Double makes the allowance itself unrepresentable, and any comparison
    against it would be meaningless.
    """
    operands = _require_operands(where, relative_coefficient, abs(float(term)))
    scaled = operands[0] * operands[1]
    _require_result(where, scaled, "conditioning scaling")
    return safe_accumulate(accumulator, scaled, where)


def allowance_from_scaled(
    scaled_terms: float,
    absolute_floor: float,
    relative_coefficient: float,
    scale_floor: float = 1.0,
) -> float:
    """The locked allowance, from an already-scaled conditioning magnitude.

    ```
    scaled_floor       = coefficient * scale_floor
    relative_allowance = max(scaled_floor, scaled_terms)
    allowance          = max(absolute_floor, relative_allowance)
    ```

    NOTE THE TWO `max` OPERATIONS. The locked formula is

        max(absolute_floor, coefficient * max(scale_floor, sum |terms|))

    and the inner `max` is a **maximum**, not an addition. An earlier
    implementation added `coefficient * scale_floor` to the scaled sum, which
    silently widened every allowance - `identity_allowance([1e6])` returned
    `1.000001e-6` where the contract says `1e-6`. Small, but this is a contract
    implementation and not a heuristic, and a tolerance may never be loosened by
    accident.
    """
    scaled_floor = safe_multiply(relative_coefficient, scale_floor, "conditioning floor")
    relative_allowance = max(scaled_floor, scaled_terms)
    return max(absolute_floor, relative_allowance)


def identity_allowance(
    terms: Sequence[float],
    absolute_floor: float,
    relative_coefficient: float,
    scale_floor: float = 1.0,
) -> float:
    """`max(absolute_floor, coefficient * max(scale_floor, sum |terms|))`, stably.

    Convenience wrapper over `scaled_magnitude` + `allowance_from_scaled` for
    callers that hold the raw terms. The analytical oracle does not: it captures
    the scaled magnitudes DURING accumulation, because the terms that matter are
    the underlying per-driver and per-year contributions, not the already-summed
    headline or annual aggregates they collapse into.
    """
    scaled = 0.0
    for index, term in enumerate(terms):
        scaled = scaled_magnitude(
            scaled, term, relative_coefficient, f"conditioning term[{index}]"
        )
    return allowance_from_scaled(scaled, absolute_floor, relative_coefficient, scale_floor)
