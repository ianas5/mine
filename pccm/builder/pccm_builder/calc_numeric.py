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

The boundary below which a Double carries fewer than 53 significand bits, which
is where the exact rescues in this module earn their keep. Exported so tests can
state that boundary by name rather than by literal."""


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


# ---------------------------------------------------------------------------
# Exact arithmetic on Double limbs — the faithful-rescue kernel
# ---------------------------------------------------------------------------
# WHY THIS EXISTS. A rescue that re-associates Double operations is a heuristic,
# not a proof. The round-2 rescues showed exactly how that fails:
#
#   * cancelling the largest opposite-signed pair with one rounded subtraction
#     `p - n` DISCARDS the rounding residual of that subtraction. When the large
#     terms then cancel, the discarded residual was the answer:
#     `[6e307, -8e307, -1.7e308, 6e307, 7e307, 6e307, -1e292]` sums to `-1e292`
#     and the rounded-pair rescue produced `-1.99792015476736e292` — a 100% error;
#   * a magnitude-balanced product order proves nothing about whether the exact
#     product is in range: it returned `MAX_DOUBLE` for a product that genuinely
#     exceeds `MAX_DOUBLE` by 0.887 ulp, and refused one whose exact value rounds
#     to `5e-324`.
#
# So the rescue paths do not re-associate Doubles at all. They compute the EXACT
# mathematical value of the already-converted IEEE-754 inputs in a wide
# fixed-point form, classify its range exactly, and round ONCE.
#
# EVERY OPERATION BELOW IS A DOUBLE OPERATION. Limbs are Doubles holding exact
# integers below `2**24`; intermediates stay below `2**49`, far inside the exact
# integer range of a Double. There is no `Decimal`, no `Fraction`, no Python
# arbitrary-precision integer, no `fsum`, no `frexp`/`ldexp`. Everything is
# addition, subtraction, multiplication, division by an exact power of two, and
# truncation — i.e. VBA `Double` arithmetic plus `Fix`. `docs/phase5_gate_a_step2.md`
# §18 specifies the translation.
#
# COST. The rescue is O(limbs) per term for sums and O(limbs**2) for products,
# on a path that only runs when the ordinary evaluation produced no value at all.
# Tier 1 remains a plain O(n) loop and is untouched.
_LIMB_BITS = 24
_LIMB_BASE = 16777216.0                      # 2**24
_TWO_52 = 4503599627370496.0
_TWO_53 = 9007199254740992.0
_MAX_SIGNIFICAND = 9007199254740991.0        # 2**53 - 1, the significand of MAX_DOUBLE
_MAX_EXPONENT = 971                          # MAX_DOUBLE = (2**53 - 1) * 2**971
_MIN_SUBNORMAL_EXPONENT = -1074
_GUARD_BITS = 64                             # extra bits before an inexact division

_POWERS_OF_TWO = (
    2.0**512, 2.0**256, 2.0**128, 2.0**64, 2.0**32,
    2.0**16, 2.0**8, 2.0**4, 2.0**2, 2.0**1,
)
_POWER_STEPS = (512, 256, 128, 64, 32, 16, 8, 4, 2, 1)
_SMALL_POWERS = tuple(float(1 << bit) for bit in range(_LIMB_BITS + 1))


def _fix(value: float) -> float:
    """Truncation toward zero — VBA `Fix` — for `0 <= value < 2**53`.

    Written with Double operations only, so nothing here depends on Python's
    arbitrary-precision `int`. Above `2**52` every Double is already an integer;
    below it, adding and subtracting `2**52` rounds to the nearest integer, and
    one correction turns that into truncation.
    """
    if value >= _TWO_52:
        return value
    rounded = (value + _TWO_52) - _TWO_52
    if rounded > value:
        rounded = rounded - 1.0
    return rounded


def _scale_by_power_of_two(value: float, exponent: int) -> float:
    """`value * 2**exponent`, applied in exact steps from a fixed power table.

    Scaling a Double by a power of two moves the exponent and leaves the
    significand alone, so every step is exact as long as the running value stays
    in range. Callers only use this where the final result is known to be
    representable and `|exponent| <= 1074`, which bounds every intermediate.
    """
    if exponent == 0 or value == 0.0:
        return value
    remaining = exponent if exponent > 0 else -exponent
    for power, step in zip(_POWERS_OF_TWO, _POWER_STEPS):
        while remaining >= step:
            value = value * power if exponent > 0 else value / power
            remaining -= step
    return value


def _decompose(value: float) -> tuple[int, float, int]:
    """`|value| = mantissa * 2**exponent` with `mantissa` an integer in
    `[2**52, 2**53)`; `(0, 0.0, 0)` for zero.

    A counting loop over a power-of-two table rather than `frexp`, because VBA has
    no `frexp` and every step here is an exact scaling. The first loop lifts
    subnormals into the normal range, which is exact — a subnormal scaled up by a
    power of two loses nothing.
    """
    if value == 0.0:
        return 0, 0.0, 0
    sign = 1 if value > 0.0 else -1
    magnitude = value if value > 0.0 else -value
    exponent = 0
    while magnitude < 1.0:
        magnitude = magnitude * _POWERS_OF_TWO[0]
        exponent -= _POWER_STEPS[0]
    for power, step in zip(_POWERS_OF_TWO, _POWER_STEPS):
        while magnitude >= power:
            magnitude = magnitude / power
            exponent += step
    return sign, magnitude * _TWO_52, exponent - 52


# --- unsigned magnitudes, base 2**24, every limb an integer in [0, 2**24) ---
def _big_new(count: int) -> list[float]:
    return [0.0] * count


def _big_add_at(limbs: list[float], index: int, amount: float) -> None:
    """`limbs += amount * 2**(24*index)`, with `amount` an integer below `2**47`."""
    carry = amount
    position = index
    while carry != 0.0:
        total = limbs[position] + carry
        carry = _fix(total / _LIMB_BASE)
        limbs[position] = total - carry * _LIMB_BASE
        position += 1


def _big_add_shifted(limbs: list[float], mantissa: float, offset: int) -> None:
    """`limbs += mantissa * 2**offset` for `offset >= 0`.

    The 53-bit mantissa is cut into three 24-bit pieces so that every piece,
    shifted by the sub-limb remainder, stays an exact integer below `2**47`.
    """
    index = offset // _LIMB_BITS
    scale = _SMALL_POWERS[offset - index * _LIMB_BITS]
    rest = mantissa
    for piece in range(3):
        quotient = _fix(rest / _LIMB_BASE)
        low = rest - quotient * _LIMB_BASE
        rest = quotient
        if low != 0.0:
            _big_add_at(limbs, index + piece, low * scale)
        if rest == 0.0:
            return


def _big_compare(left: list[float], right: list[float]) -> int:
    for index in range(len(left) - 1, -1, -1):
        if left[index] != right[index]:
            return 1 if left[index] > right[index] else -1
    return 0


def _big_subtract(left: list[float], right: list[float]) -> None:
    """`left -= right`; the caller has established `left >= right`."""
    borrow = 0.0
    for index in range(len(left)):
        value = left[index] - right[index] - borrow
        if value < 0.0:
            value = value + _LIMB_BASE
            borrow = 1.0
        else:
            borrow = 0.0
        left[index] = value


def _big_multiply(left: list[float], right: list[float]) -> list[float]:
    """Schoolbook product. Every intermediate stays below `2**49`."""
    result = _big_new(len(left) + len(right) + 1)
    for i in range(len(left)):
        if left[i] == 0.0:
            continue
        carry = 0.0
        for j in range(len(right)):
            total = result[i + j] + left[i] * right[j] + carry
            carry = _fix(total / _LIMB_BASE)
            result[i + j] = total - carry * _LIMB_BASE
        position = i + len(right)
        while carry != 0.0:
            total = result[position] + carry
            carry = _fix(total / _LIMB_BASE)
            result[position] = total - carry * _LIMB_BASE
            position += 1
    return result


def _big_divide_small(limbs: list[float], divisor: float) -> tuple[list[float], float]:
    """Exact `divmod` by a small integer. Used only for the 2, 3 and 6 of the
    convex statistics, where the numerator is a dyadic sum but the statistic is
    not."""
    quotient = _big_new(len(limbs))
    remainder = 0.0
    for index in range(len(limbs) - 1, -1, -1):
        current = remainder * _LIMB_BASE + limbs[index]
        share = _fix(current / divisor)
        quotient[index] = share
        remainder = current - share * divisor
    return quotient, remainder


def _big_top_bit(limbs: list[float]) -> int:
    """Index of the most significant set bit, or `-1` when the value is zero."""
    for index in range(len(limbs) - 1, -1, -1):
        value = limbs[index]
        if value != 0.0:
            bit = 0
            while value >= 2.0:
                value = _fix(value / 2.0)
                bit += 1
            return index * _LIMB_BITS + bit
    return -1


def _big_bit(limbs: list[float], position: int) -> bool:
    index = position // _LIMB_BITS
    if index >= len(limbs):
        return False
    shifted = _fix(limbs[index] / _SMALL_POWERS[position - index * _LIMB_BITS])
    return shifted - _fix(shifted / 2.0) * 2.0 != 0.0


def _big_any_below(limbs: list[float], position: int) -> bool:
    index = position // _LIMB_BITS
    offset = position - index * _LIMB_BITS
    for lower in range(min(index, len(limbs))):
        if limbs[lower] != 0.0:
            return True
    if offset == 0 or index >= len(limbs):
        return False
    scale = _SMALL_POWERS[offset]
    return limbs[index] - _fix(limbs[index] / scale) * scale != 0.0


def _big_high_part(limbs: list[float], drop: int) -> float:
    """`floor(value / 2**drop)`, which the caller guarantees is below `2**53`.

    At most four limbs can reach a 53-bit result, so the loop is bounded and every
    partial total is an exact integer.
    """
    index = drop // _LIMB_BITS
    if index >= len(limbs):
        return 0.0
    offset = drop - index * _LIMB_BITS
    total = _fix(limbs[index] / _SMALL_POWERS[offset])
    weight = _SMALL_POWERS[_LIMB_BITS - offset]
    for step in (1, 2, 3):
        position = index + step
        if position >= len(limbs):
            break
        if limbs[position] != 0.0:
            total = total + limbs[position] * weight
        weight = weight * _LIMB_BASE
    return total


def _big_whole(limbs: list[float]) -> float:
    """The whole magnitude, which the caller guarantees is below `2**53`."""
    total = 0.0
    weight = 1.0
    for index in range(min(len(limbs), 3)):
        if limbs[index] != 0.0:
            total = total + limbs[index] * weight
        weight = weight * _LIMB_BASE
    return total


def _round_exact(
    sign: int,
    limbs: list[float],
    shift: int,
    where: str,
    sticky_below: bool = False,
    underflow_to_zero: bool = False,
) -> float:
    """Round `sign * limbs * 2**shift` to the nearest Double, ties to even.

    THE RANGE TEST IS ON THE EXACT VALUE, not on the rounded one. A sum can exceed
    `MAX_DOUBLE` by less than half an ulp and still round to `MAX_DOUBLE`; that is
    an out-of-range result and it is refused, because C2 removes refusals of
    answers that exist and creates no fabricated ones. Likewise a non-zero exact
    value with no non-zero Double is refused rather than reported as zero.

    `sticky_below` carries the "there is more, below everything represented here"
    flag from an inexact division, so a value that only looks like a tie is not
    rounded as one.

    `underflow_to_zero` is for CONDITIONING METADATA ONLY. Model arithmetic must
    never take it: it turns the underflow refusal into `0.0`, which is the policy
    `scaled_magnitude` already documents for a tolerance term too small to move the
    allowance it feeds. Overflow is still refused under the flag, because a
    conditioning scale outside Double makes the allowance itself meaningless.
    """
    top = _big_top_bit(limbs)
    if top < 0:
        if sticky_below and not underflow_to_zero:
            raise NumericalRangeRefusal(
                f"{where}: underflowed — a non-zero result has no usable non-zero Double"
            )
        return 0.0

    exponent = top + shift
    if exponent > 1023:
        raise NumericalRangeRefusal(
            f"{where}: the exact result is outside finite Double range"
        )

    target = exponent - 52
    if target < _MIN_SUBNORMAL_EXPONENT:
        target = _MIN_SUBNORMAL_EXPONENT
    drop = target - shift

    if drop <= 0:
        # Fewer than 54 significant bits and no bit below 2**-1074: the value is
        # exactly a Double already, so there is nothing to round.
        quotient = _big_whole(limbs)
        scale = shift
        if sticky_below:
            raise NumericalRangeRefusal(
                f"{where}: the exact result needs more precision than a Double holds"
            )
    else:
        quotient = _big_high_part(limbs, drop)
        round_bit = _big_bit(limbs, drop - 1)
        sticky = sticky_below or _big_any_below(limbs, drop - 1)
        if target == _MAX_EXPONENT and (
            quotient > _MAX_SIGNIFICAND
            or (quotient == _MAX_SIGNIFICAND and (round_bit or sticky))
        ):
            raise NumericalRangeRefusal(
                f"{where}: the exact result is outside finite Double range"
            )
        odd = quotient - _fix(quotient / 2.0) * 2.0 != 0.0
        if round_bit and (sticky or odd):
            quotient = quotient + 1.0
        if quotient == 0.0:
            if underflow_to_zero:
                return 0.0
            raise NumericalRangeRefusal(
                f"{where}: underflowed — a non-zero result has no usable non-zero Double"
            )
        scale = target

    result = _scale_by_power_of_two(quotient, scale)
    return result if sign > 0 else -result


def _exact_sum(terms: Sequence[float]) -> tuple[int, list[float], int]:
    """`(sign, magnitude limbs, shift)` for the exact mathematical sum of Doubles.

    Every Double is an integer multiple of `2**smallest`, where `smallest` is the
    least of the terms' own exponents, so aligning them there is exact and the sum
    is an exact integer in that unit. Positive and negative magnitudes are
    accumulated separately and subtracted once, which needs no signed carries.
    """
    parts: list[tuple[int, float, int]] = []
    for value in terms:
        sign, mantissa, exponent = _decompose(value)
        if sign != 0:
            parts.append((sign, mantissa, exponent))
    if not parts:
        return 0, _big_new(1), 0

    smallest = parts[0][2]
    largest = parts[0][2]
    for _, _, exponent in parts:
        if exponent < smallest:
            smallest = exponent
        if exponent > largest:
            largest = exponent
    count = (largest - smallest) // _LIMB_BITS + 6

    positive = _big_new(count)
    negative = _big_new(count)
    for sign, mantissa, exponent in parts:
        _big_add_shifted(
            positive if sign > 0 else negative, mantissa, exponent - smallest
        )

    order = _big_compare(positive, negative)
    if order == 0:
        return 0, _big_new(1), 0
    if order > 0:
        _big_subtract(positive, negative)
        return 1, positive, smallest
    _big_subtract(negative, positive)
    return -1, negative, smallest


def _exact_product(factors: Sequence[float]) -> tuple[int, list[float], int]:
    """`(sign, magnitude limbs, shift)` for the exact mathematical product.

    The mantissas multiply as integers and the exponents add, so the product is
    exact regardless of how far outside Double range it lands — which is what lets
    the range classification be a fact rather than an artefact of evaluation order.
    """
    sign = 1
    shift = 0
    limbs = _big_new(3)
    limbs[0] = 1.0
    for value in factors:
        part_sign, mantissa, exponent = _decompose(value)
        if part_sign == 0:
            return 0, _big_new(1), 0
        sign = sign * part_sign
        shift = shift + exponent
        mantissa_limbs = _big_new(3)
        _big_add_shifted(mantissa_limbs, mantissa, 0)
        limbs = _big_multiply(limbs, mantissa_limbs)
    return sign, limbs, shift


def _big_add_big_shifted(target: list[float], limbs: list[float], offset: int) -> None:
    """`target += limbs * 2**offset` for `offset >= 0`. Limb by limb; each limb is
    an integer below `2**24`, well inside what `_big_add_shifted` accepts."""
    for index in range(len(limbs)):
        if limbs[index] != 0.0:
            _big_add_shifted(target, limbs[index], offset + index * _LIMB_BITS)


def _exact_sum_of_products(
    groups: Sequence[Sequence[float]],
) -> tuple[int, list[float], int]:
    """`(sign, magnitude limbs, shift)` for `SUM over groups of PRODUCT of factors`.

    THE ONE COMPOSITION THE KERNEL DID NOT HAVE, and the whole of what round 5
    adds to it. Each product is formed exactly — including products that have no
    Double of their own — and the products are then added exactly. Only the
    finished expression is range-classified and rounded, so an intermediate that
    steps outside Double range never becomes a boundary.

    That matters because a named Phase-5 output is often a sum of products whose
    individual terms are implementation detail:

        Knom = SUM_y ( FX * w_y * infl_y )

    is the same number as `FX * SUM_y (w_y * infl_y)`, but evaluating it in this
    form lets `w_y * infl_y` be wider than a Double while `Knom` is not.

    Nothing about the existing kernel changes: `_exact_product` builds each term
    and `_round_exact` finishes the job, exactly as they already do.
    """
    parts: list[tuple[int, list[float], int, int]] = []
    for factors in groups:
        sign, limbs, shift = _exact_product(factors)
        if sign == 0:
            continue
        top = _big_top_bit(limbs)
        if top < 0:
            continue
        parts.append((sign, limbs, shift, top))
    if not parts:
        return 0, _big_new(1), 0

    smallest = parts[0][2]
    highest = parts[0][2] + parts[0][3]
    for _, _, shift, top in parts:
        if shift < smallest:
            smallest = shift
        if shift + top > highest:
            highest = shift + top
    count = (highest - smallest) // _LIMB_BITS + 6

    positive = _big_new(count)
    negative = _big_new(count)
    for sign, limbs, shift, _ in parts:
        _big_add_big_shifted(
            positive if sign > 0 else negative, limbs, shift - smallest
        )

    order = _big_compare(positive, negative)
    if order == 0:
        return 0, _big_new(1), 0
    if order > 0:
        _big_subtract(positive, negative)
        return 1, positive, smallest
    _big_subtract(negative, positive)
    return -1, negative, smallest


def exact_sum_of_products(groups: Sequence[Sequence[float]], where: str) -> float:
    """`SUM over groups of PRODUCT of factors`, computed exactly and rounded once.

    THE COMPOUND-EXPRESSION RESCUE. Callers use it only after their ordinary
    staged evaluation has failed at an intermediate the model never publishes; see
    `docs/phase5_gate_a_step2.md` §20. The classification is the kernel's, so it is
    the same one every other rescue uses:

        |exact| > MAX_DOUBLE               -> NumericalRangeRefusal
        exact non-zero but rounding to 0   -> NumericalRangeRefusal
        otherwise                          -> the correctly rounded Double

    Every operand must already be a usable Double: this widens the arithmetic
    BETWEEN named values, never the values themselves.
    """
    for index, factors in enumerate(groups):
        _require_operands(f"{where}[{index}]", *factors)
    sign, limbs, shift = _exact_sum_of_products(groups)
    return _round_exact(sign, limbs, shift, where)


def _round_exact_quotient(
    sign: int, limbs: list[float], shift: int, divisor: float, where: str
) -> float:
    """Round `sign * limbs * 2**shift / divisor` to the nearest Double.

    Dividing by 3 or 6 leaves the dyadic world, so the numerator is first shifted
    left by `_GUARD_BITS` and the division remainder becomes a sticky flag. The
    quotient is then exact to far more bits than a Double holds, and a non-zero
    remainder means the true value is strictly above it — which is precisely what
    a sticky bit encodes, so the rounding, ties included, is still exact.
    """
    if divisor == 1.0:
        return _round_exact(sign, limbs, shift, where)
    guarded = _big_new(len(limbs) + _GUARD_BITS // _LIMB_BITS + 2)
    for index in range(len(limbs)):
        if limbs[index] != 0.0:
            _big_add_shifted(guarded, limbs[index], index * _LIMB_BITS + _GUARD_BITS)
    quotient, remainder = _big_divide_small(guarded, divisor)
    return _round_exact(
        sign, quotient, shift - _GUARD_BITS, where, sticky_below=remainder != 0.0
    )


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

    **Tier 2 — the exact sum, only after tier 1 overflowed.** A partial sum can
    exceed Double while the true total is perfectly representable: `MAX + MAX -
    MAX` is `MAX`, but left to right the first addition is already infinite.
    Refusing that is refusing an answer that exists (§19.2).

    Tier 2 does **not** re-associate Double additions. Re-association is a
    heuristic and it silently loses information: cancelling the largest
    opposite-signed pair with one rounded subtraction discards that subtraction's
    rounding residual, and when the large terms cancel the residual WAS the
    answer. `[6e307, -8e307, -1.7e308, 6e307, 7e307, 6e307, -1e292]` sums to
    exactly `-1e292`, and a rounded-pair rescue produces `-1.99792015476736e292`.

    Instead the terms are added in a wide exact fixed-point form (`_exact_sum`),
    the range of the exact total is classified against `MAX_DOUBLE` and against
    the smallest subnormal, and the result is rounded **once**, to nearest with
    ties to even. So:

        exact total representable   -> the correctly rounded Double
        |exact total| > MAX_DOUBLE  -> NumericalRangeRefusal
        exact total non-zero but
          rounding to zero          -> NumericalRangeRefusal
        exact total exactly zero    -> +0.0

    The range test is on the EXACT total, not on the rounded one: a sum can exceed
    `MAX_DOUBLE` by less than half an ulp and still round to it, and returning
    `MAX_DOUBLE` there would fabricate a value C2 explicitly forbids.

    Tier 2 is only reached where tier 1 produced no value at all. It is NOT
    invoked for underflow: a sum cannot underflow to zero except by genuine
    cancellation, which is a real result.

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

    # --- tier 2: the exact mathematical sum, rounded once --------------------
    sign, limbs, shift = _exact_sum(values)
    return _round_exact(sign, limbs, shift, f"{where} (faithful rescue)")


# ---------------------------------------------------------------------------
# Stable products
# ---------------------------------------------------------------------------
def safe_product(factors: Sequence[float], where: str = "product") -> float:
    """A short product, evaluated so that a representable answer is not lost to a
    gratuitous intermediate overflow or underflow.

    TWO TIERS, deliberately:

    1. **Left to right first.** For every ordinary model this succeeds, and it is
       bit-for-bit what a naive implementation produces. Nothing about existing
       results changes.
    2. **Only if that fails**, the EXACT product of the already-converted Doubles
       is formed (`_exact_product`: the mantissas multiply as integers and the
       exponents add), its range is classified exactly, and it is rounded once.

    A REORDERING IS NOT A PROOF, which is why tier 2 is not one. The round-2
    magnitude-balanced order returned `MAX_DOUBLE` for `[1e50, MAX_DOUBLE,
    1e-150, 1e100]`, whose exact product exceeds `MAX_DOUBLE` by 0.887 ulp, and
    refused `[1e100, 0.5, 1e150, 5e-324, 1e-250]`, whose exact product rounds to
    `5e-324`. Both are answered correctly by the exact form:

        exact product representable  -> the correctly rounded Double
        |exact product| > MAX_DOUBLE -> NumericalRangeRefusal
        exact product non-zero but
          rounding to zero           -> NumericalRangeRefusal

    The alternative to tier 2 is refusing a valid calculation, so it is only ever
    reached where the choice is "the right answer" versus "no answer". A later VBA
    implementation must reproduce BOTH tiers, in this order, to match.
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

    sign, limbs, shift = _exact_product(values)
    return _round_exact(sign, limbs, shift, f"{where} (faithful rescue)")


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
#   Tier 2 — the exact statistic, rounded once. Reached when tier 1 raised, AND
#            when tier 1 returned exactly zero.
#
# WHY A TIER-1 ZERO IS NOT AUTOMATICALLY ACCEPTED. `midpoint(-20s, 19s)` with
# `s = 5e-324` evaluates as `-10s + fl(9.5s)` = `-10s + 10s` = `0`, and tier 1
# raises nothing. The exact midpoint is `-0.5s`, which is NOT zero and has no
# usable non-zero Double. Returning `0` there reports a value the model does not
# have, which is the silent-underflow failure §19.3 exists to prevent. So a
# non-degenerate zero is classified exactly: a numerator that is mathematically
# zero gives `0.0`, and one that is not gives either the correctly rounded tiny
# Double or a controlled refusal. A NON-ZERO tier-1 result is never touched.
def _degenerate_point(values: tuple[float, ...]) -> float | None:
    """The single point of a zero-uncertainty distribution, or `None`."""
    first = values[0]
    for value in values[1:]:
        if value != first:
            return None
    return first


def _convex_mean(
    values: tuple[float, ...],
    where: str,
    stable: "Callable[[tuple[float, ...]], float]",
    numerator: "Callable[[tuple[float, ...]], tuple[float, ...]]",
    divisor: float,
) -> float:
    """Tier 0 -> tier 1 -> exact. See the block comment above.

    `numerator` returns the terms whose exact sum is the statistic's mathematical
    numerator, and `divisor` its denominator — `(Min + ML + Max) / 3`,
    `(Min + 4*ML + Max) / 6`, `(Min + Max) / 2`. The `4*ML` is supplied as four
    copies of `ML` rather than a multiplication, so the exact numerator can be
    formed even where `4*ML` itself has no Double.
    """
    numbers = _require_operands(where, *values)
    point = _degenerate_point(numbers)
    if point is not None:
        return point

    try:
        result = stable(numbers)
    except NumericalRangeRefusal:
        result = None
    if result is not None and result != 0.0:
        return result

    sign, limbs, shift = _exact_sum(numerator(numbers))
    return _round_exact_quotient(sign, limbs, shift, divisor, where)


def triangular_mean(minimum: float, most_likely: float, maximum: float) -> float:
    """`Min/3 + ML/3 + Max/3` — never `(Min + ML + Max) / 3`."""
    where = "triangular mean"

    def stable(v: tuple[float, ...]) -> float:
        total = safe_accumulate(0.0, safe_divide(v[0], 3.0, where), where)
        total = safe_accumulate(total, safe_divide(v[1], 3.0, where), where)
        return safe_accumulate(total, safe_divide(v[2], 3.0, where), where)

    def numerator(v: tuple[float, ...]) -> tuple[float, ...]:
        return (v[0], v[1], v[2])

    return _convex_mean(
        (minimum, most_likely, maximum), where, stable, numerator, 3.0
    )


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

    def numerator(v: tuple[float, ...]) -> tuple[float, ...]:
        return (v[0], v[1], v[1], v[1], v[1], v[2])

    return _convex_mean(
        (minimum, most_likely, maximum), where, stable, numerator, 6.0
    )


def midpoint(minimum: float, maximum: float) -> float:
    """`Min/2 + Max/2` — never `(Min + Max) / 2`."""
    where = "midpoint"

    def stable(v: tuple[float, ...]) -> float:
        total = safe_accumulate(0.0, safe_divide(v[0], 2.0, where), where)
        return safe_accumulate(total, safe_divide(v[1], 2.0, where), where)

    def numerator(v: tuple[float, ...]) -> tuple[float, ...]:
        return (v[0], v[1])

    return _convex_mean((minimum, maximum), where, stable, numerator, 2.0)


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


def scaled_magnitude_of_product(
    accumulator: float,
    factors: Sequence[float],
    relative_coefficient: float,
    where: str,
) -> float:
    """`scaled_magnitude` for a contribution that has NO Double of its own.

    Erratum C1 conditions reconciliation on the underlying contributions, and some
    of those contributions are non-materialized intermediates: a per-driver,
    per-year annual term can be `2 * MAX_DOUBLE` while the annual row it feeds is
    zero. The quantity actually needed is `coefficient * |contribution|`, and with
    the locked `1e-12` that is finite — so the unscaled contribution must not have
    to become a Double just so its metadata can be recorded.

    The coefficient is therefore folded into the SAME exact factor expression and
    the range classification happens once, at the end. The underflow policy is
    `scaled_magnitude`'s, unchanged and for the same reason: a tolerance term too
    small to represent cannot move an allowance floored at `coefficient * 1`.
    """
    magnitudes = _require_operands(where, relative_coefficient, *factors)
    group = (magnitudes[0],) + tuple(
        value if value >= 0.0 else -value for value in magnitudes[1:]
    )
    sign, limbs, shift = _exact_product(group)
    scaled = _round_exact(sign, limbs, shift, where, underflow_to_zero=True)
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
