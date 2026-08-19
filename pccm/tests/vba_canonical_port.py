"""A line-for-line port of the SHIPPED modCalcFingerprint canonical encoder.

WHY THIS FILE EXISTS. Linux has no VBA, so the corrected canonical Double
encoder cannot be executed here. It CAN be transcribed, routine for routine,
guard for guard, and then held against the Python oracle over the whole binary64
domain. That is what tests/test_phase5_canonical_number.py does.

WHAT THIS PROVES AND WHAT IT DOES NOT. It proves the ALGORITHM is correct - the
decomposition, the exact big-integer expansion, the round-half-even, the output
shape, the refusals. It does NOT prove that VBA executes this transcription
faithfully; that is Gate B's, against the emitted parity corpus, on Windows.

THIS FILE IS NOT AN AUTHORITY. It is never used to produce an expected value:
the expectations come from builder/pccm_builder/calc_fingerprint.py, and a test
that compared this port against itself would prove nothing. A source test pins
each routine here against the VBA it mirrors, so the two cannot drift.
"""
import math

FP_FRACTION_DIGITS = 16
FP_SIGNIFICANT_DIGITS = 17
FP_LIMB_BASE = 10000000.0
FP_LIMB_DIGITS = 7
FP_TWO_52 = 4503599627370496.0
FP_TWO_53 = 9007199254740992.0
FP_MANTISSA_DIGITS = 16
FP_MAX_LIMBS = 200
FP_DIGIT_TABLE = "0123456789"
MAX_DOUBLE = 1.7976931348623157e308


def is_usable_double(v):
    if not (v == v): return False
    if v > MAX_DOUBLE: return False
    if v < -MAX_DOUBLE: return False
    return True


def utf16_length(t):
    return sum(2 if ord(c) > 0xFFFF else 1 for c in t)


def decompose(value):
    scaled = abs(value); exponent = 0; guard = 0
    while scaled < FP_TWO_52:
        scaled = scaled * 2.0; exponent -= 1; guard += 1
        if guard > 1200: return None
    while scaled >= FP_TWO_53:
        scaled = scaled / 2.0; exponent += 1; guard += 1
        if guard > 2400: return None
    return scaled, exponent


def limbs_from_mantissa(mantissa):
    digits = [0] * (FP_MANTISSA_DIGITS + 1)
    remainder = mantissa
    power = 1000000000000000.0
    for place in range(1, FP_MANTISSA_DIGITS + 1):
        count = 0
        while remainder >= power:
            remainder = remainder - power; count += 1
            if count > 9: return None
        digits[place] = count
        power = power / 10.0
    if remainder != 0.0: return None
    limbs = [0.0] * (FP_MAX_LIMBS + 1)
    limb_count = 0
    place = FP_MANTISSA_DIGITS
    while place >= 1:
        first = place - FP_LIMB_DIGITS + 1
        if first < 1: first = 1
        value = 0.0
        for index in range(first, place + 1):
            value = value * 10.0 + digits[index]
        limbs[limb_count] = value; limb_count += 1
        place = first - 1
    while limb_count > 1:
        if limbs[limb_count - 1] != 0.0: break
        limb_count -= 1
    return limbs, limb_count


def integer_power(base, power):
    result = 1.0
    for _ in range(power): result = result * base
    return result


def multiply_small(limbs, limb_count, factor):
    carry = 0.0
    for index in range(limb_count):
        product = limbs[index] * factor + carry
        assert product < FP_TWO_53, "exact-integer ceiling exceeded"
        quotient = float(int(product / FP_LIMB_BASE))
        limbs[index] = product - quotient * FP_LIMB_BASE
        carry = quotient
    while carry > 0.0:
        if limb_count > FP_MAX_LIMBS: return None
        quotient = float(int(carry / FP_LIMB_BASE))
        limbs[limb_count] = carry - quotient * FP_LIMB_BASE
        limb_count += 1
        carry = quotient
    return limb_count


def multiply_power(limbs, limb_count, base, count, chunk):
    if count < 0: return None
    if chunk < 1: return None
    passes = 0
    remainder = count
    while remainder >= chunk:
        remainder = remainder - chunk
        passes += 1
    factor = integer_power(base, chunk)
    for _ in range(passes):
        limb_count = multiply_small(limbs, limb_count, factor)
        if limb_count is None: return None
    if remainder > 0:
        limb_count = multiply_small(limbs, limb_count, integer_power(base, remainder))
        if limb_count is None: return None
    return limb_count


def plain_digits(value):
    if value == 0.0: return "0"
    out = ""; remainder = value; power = 1000000.0
    while power >= 1.0:
        count = 0
        while remainder >= power:
            remainder = remainder - power; count += 1
        if len(out) > 0 or count > 0:
            out = out + FP_DIGIT_TABLE[count]
        power = power / 10.0
    return out


def limb_digits(limbs, limb_count):
    out = plain_digits(limbs[limb_count - 1])
    for index in range(limb_count - 2, -1, -1):
        part = plain_digits(limbs[index])
        out = out + "0" * (FP_LIMB_DIGITS - len(part)) + part
    return out


def digit_value(char):
    return FP_DIGIT_TABLE.find(char)


def has_non_zero(text):
    return any(c != "0" for c in text)


def increment_digits(digits):
    buffer = list(digits); carry = 1
    for index in range(len(buffer) - 1, -1, -1):
        value = digit_value(buffer[index])
        if value < 0: return None
        value += carry
        if value >= 10:
            value -= 10; carry = 1
        else:
            carry = 0
        buffer[index] = FP_DIGIT_TABLE[value]
        if carry == 0: break
    out = "".join(buffer)
    if carry == 1: out = "1" + out
    return out


def round_significant(all_digits, exp10):
    total = len(all_digits)
    if total <= FP_SIGNIFICANT_DIGITS:
        return all_digits + "0" * (FP_SIGNIFICANT_DIGITS - total), exp10
    head = all_digits[:FP_SIGNIFICANT_DIGITS]
    next_digit = all_digits[FP_SIGNIFICANT_DIGITS]
    tail = all_digits[FP_SIGNIFICANT_DIGITS + 1:]
    last_digit = digit_value(head[-1])
    if last_digit < 0: return None
    round_up = False
    if next_digit > "5":
        round_up = True
    elif next_digit == "5":
        if has_non_zero(tail):
            round_up = True
        else:
            round_up = last_digit in (1, 3, 5, 7, 9)
    if round_up:
        carried = increment_digits(head)
        if carried is None: return None
        if len(carried) > FP_SIGNIFICANT_DIGITS:
            head = carried[:FP_SIGNIFICANT_DIGITS]; exp10 += 1
        else:
            head = carried
    return head, exp10


def long_digits(value):
    if value == 0: return "0"
    if value < 0: return None
    out = ""; remainder = value; power = 10000
    tenth = {10000: 1000, 1000: 100, 100: 10, 10: 1}
    while power >= 1:
        count = 0
        while remainder >= power:
            remainder = remainder - power; count += 1
        if len(out) > 0 or count > 0:
            out = out + FP_DIGIT_TABLE[count]
        power = tenth.get(power, 0)
    return out


def exponent_text(exp10):
    if exp10 < 0:
        sign = "-"; magnitude = -exp10
    else:
        sign = "+"; magnitude = exp10
    digits = long_digits(magnitude)
    if len(digits) < 2: digits = "0" * (2 - len(digits)) + digits
    return sign + digits


def build_canonical(value):
    if value == 0.0:
        return "0." + "0" * FP_FRACTION_DIGITS + "E+00"
    sign = "-" if value < 0.0 else ""
    d = decompose(value)
    if d is None: return None
    mantissa, exponent = d
    lm = limbs_from_mantissa(mantissa)
    if lm is None: return None
    limbs, limb_count = lm
    if exponent >= 0:
        limb_count = multiply_power(limbs, limb_count, 2.0, exponent, 23); scale = 0
    else:
        limb_count = multiply_power(limbs, limb_count, 5.0, -exponent, 10); scale = exponent
    if limb_count is None: return None
    all_digits = limb_digits(limbs, limb_count)
    if len(all_digits) == 0: return None
    exp10 = len(all_digits) - 1 + scale
    r = round_significant(all_digits, exp10)
    if r is None: return None
    head, exp10 = r
    return sign + head[0] + "." + head[1:] + "E" + exponent_text(exp10)


def marker_index(text):
    """Port of CalcFpMarkerIndex - the structural post-condition."""
    first = 1
    if text[0:1] == "-": first = 2
    if not (len(text) >= first and text[first-1].isdigit()): return 0
    marker = first + 1
    if len(text) < marker + FP_FRACTION_DIGITS + 3: return 0
    for index in range(marker + 1, marker + FP_FRACTION_DIGITS + 1):
        if not text[index-1].isdigit(): return 0
    if text[marker + FP_FRACTION_DIGITS] != "E": return 0
    tail = marker + FP_FRACTION_DIGITS + 2
    if text[tail-1] not in "+-": return 0
    if len(text) - tail < 2: return 0
    for index in range(tail + 1, len(text) + 1):
        if not text[index-1].isdigit(): return 0
    return marker


def canonical_number(value, decimal_separator="."):
    """Port of CalcFpCanonicalNumber. Returns (ok, text)."""
    if utf16_length(decimal_separator) != 1: return False, ""
    if not is_usable_double(value): return False, ""
    number = value
    if number == 0.0: number = 0.0
    text = build_canonical(number)
    if text is None: return False, ""
    if marker_index(text) == 0: return False, ""
    return True, text
