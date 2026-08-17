"""PCCM Calculation Input Fingerprint - the Python reference implementation.

This module is the SINGLE SOURCE for the fingerprint mathematics. The hash base,
both moduli and both initial states are defined here and nowhere else: not in
`spec/calc_contract.yaml`, not in a second Python module, and not hand-copied into
VBA. The VBA implementation is projected from these values, and Gate B proves the
two agree on real Excel.

`spec/calc_contract.yaml` owns exactly one fingerprint value - the ALGORITHM
VERSION NUMBER - and it is passed in, never imported from here. Which encoding
produced a stored digest is a layout/audit fact; how the encoding works is
mathematics.

Scope, deliberately: FINGERPRINT SEMANTICS ONLY. No analytical cost, risk, FX,
inflation, discounting or reconciliation logic appears here, and none may be
added. The analytical oracle is a separate, later module.

--------------------------------------------------------------------------------
WHAT IS PROVEN WHERE
--------------------------------------------------------------------------------
Gate A (Linux, this module):  the reference semantics, asserted against fixed
                              literal vectors. NO VBA IS EXECUTED ON LINUX.
Gate B (Windows, real Excel):  actual VBA execution, actual `AscW` behaviour,
                              actual locale formatting, the actual Double-only
                              reducer, end-to-end parity.

`reduce_double_only` mirrors the locked VBA reduction so that Gate A can prove the
reference semantics are self-consistent. It is a Python mirror, not a VBA
execution, and no test may describe it as one.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from typing import Callable, Iterable, Iterator, Sequence

# ---------------------------------------------------------------------------
# Hash constants - LOCKED, and defined ONLY here
# ---------------------------------------------------------------------------
FP_BASE = 131
"""Polynomial base of the rolling hash."""

FP_MOD_1 = 2147483647
"""First modulus, 2**31 - 1, prime."""

FP_MOD_2 = 2147483629
"""Second modulus, prime. Two independent moduli make an accidental collision on
both accumulators simultaneously implausible for a change detector."""

FP_INIT_1 = 1
FP_INIT_2 = 1
"""Both accumulators start at 1, so a stream beginning with NUL is not absorbed."""

STREAM_TAG = "PCCM-FP"
"""First field of every canonical stream."""

SECTION_HEADER = "HEADER"
SECTION_COST = "COST"
SECTION_RISK = "RISK"

SECTION_ORDER = (SECTION_HEADER, SECTION_COST, SECTION_RISK)
"""Fixed, never sorted. Later phases append after these so the analytical subset
stays comparable across phases."""

TAG_TEXT = "S"
TAG_NUMBER = "N"
TAG_INTEGER = "I"

_SIGNIFICAND_DIGITS_AFTER_POINT = 16
"""17 significant digits: one before the point and sixteen after."""


class FingerprintError(ValueError):
    """Raised when a value cannot be canonically encoded.

    Fail loudly. A fingerprint that silently encodes a non-finite value as some
    placeholder would report two different models as unchanged.
    """


# ---------------------------------------------------------------------------
# UTF-16 code units
# ---------------------------------------------------------------------------
def utf16_code_units(text: str) -> Iterator[int]:
    """Yield the UTF-16 code units of `text`, in order, each in 0 .. 65535.

    The hash is defined over UTF-16 code units because that is what VBA's `AscW`
    returns when walking a string. A non-BMP character contributes TWO units - its
    surrogate pair - exactly as it would in VBA, and never one code point.
    """
    raw = text.encode("utf-16-le")
    for index in range(0, len(raw), 2):
        yield raw[index] | (raw[index + 1] << 8)


def utf16_length(text: str) -> int:
    """The number of UTF-16 code units in `text`.

    This is the length a length prefix carries, and it is NOT `len(text)`: Python
    counts code points, so a non-BMP character would be counted once here and
    twice by VBA. Using `len` would make the two implementations disagree on any
    astral character.
    """
    return len(text.encode("utf-16-le")) // 2


def normalise_code_unit(unit: int) -> int:
    """Normalise a possibly sign-extended code unit into 0 .. 65535.

    VBA's `AscW` returns a `Integer`, which is SIGNED 16-bit, so every code unit
    above `U+7FFF` comes back negative. The VBA implementation must add 65536 to a
    negative result; this function is the reference for that normalisation, and the
    `> U+7FFF` test vector is what proves it.
    """
    if unit < 0:
        unit += 65536
    if not 0 <= unit <= 65535:
        raise FingerprintError(f"code unit {unit} is outside 0 .. 65535 after normalisation")
    return unit


# ---------------------------------------------------------------------------
# Canonical numeric encoding
# ---------------------------------------------------------------------------
_HOST_FORM_RE = re.compile(
    r"""
    ^
    (?P<sign>-?)                 # optional leading minus
    (?P<lead>[0-9])              # exactly one digit before the decimal marker
    (?P<marker>.)                # the decimal marker, WHATEVER character it is
    (?P<frac>[0-9]{16})          # exactly sixteen fractional digits
    E                            # the exponent marker
    (?P<exp_sign>[+-])           # the exponent sign, always present
    (?P<exp>[0-9]{2,})           # at least two exponent digits
    $
    """,
    re.VERBOSE | re.DOTALL,
)
"""The accepted scientific form, matched STRUCTURALLY.

`marker` is `.` - any single character - and never the separator itself, so the
pattern still parses a host string whose separator is `E`, `+`, `-` or a digit.
The fixed sixteen-digit fraction is what keeps that unambiguous: the exponent
marker is the `E` that follows exactly sixteen digits, wherever else an `E` may
appear."""


def _decimal_marker_index(text: str) -> int:
    """The index of the mantissa decimal marker, by POSITION not by character.

    Searching for the separator is what makes a global replace unsafe: `E`, `+`,
    `-` and every digit already occur elsewhere in scientific notation, so a
    search-and-replace corrupts the exponent marker, the exponent sign or the
    mantissa itself. The position, by contrast, is fixed by the form: optional
    sign, one digit, then the marker.
    """
    match = _HOST_FORM_RE.match(text)
    if match is None:
        raise FingerprintError(
            f"not in the accepted scientific form (optional '-', one digit, decimal marker, "
            f"16 fractional digits, 'E', exponent sign, >=2 exponent digits): {text!r}"
        )
    return len(match.group("sign")) + 1


def apply_decimal_separator(text: str, decimal_separator: str) -> str:
    """Rewrite the mantissa decimal marker to `decimal_separator`, and nothing else.

    Models what a host formatter under that locale would have handed back. Only the
    single marker character is touched; the exponent marker, the exponent sign, the
    leading sign and every digit are left exactly where they are.
    """
    index = _decimal_marker_index(text)
    return text[:index] + decimal_separator + text[index + 1 :]


def normalise_decimal_separator(text: str, decimal_separator: str) -> str:
    """Rewrite the mantissa decimal marker back to `.`, and nothing else.

    The inverse of `apply_decimal_separator`, and the operation the VBA encoder
    must perform on whatever its host formatter produced.
    """
    index = _decimal_marker_index(text)
    if text[index] != decimal_separator:
        raise FingerprintError(
            f"expected the decimal marker at index {index} to be "
            f"{decimal_separator!r}, found {text[index]!r} in {text!r}"
        )
    return text[:index] + "." + text[index + 1 :]


def _check_separator(decimal_separator: str) -> None:
    if not isinstance(decimal_separator, str):
        raise FingerprintError(f"decimal separator must be a string, got {decimal_separator!r}")
    if utf16_length(decimal_separator) != 1:
        # Measured in UTF-16 code units, not Python code points: an astral
        # character is one code point but two units, and no locale decimal
        # separator is outside the BMP. Stating the constraint beats assuming it.
        raise FingerprintError(
            f"decimal separator must be exactly one UTF-16 code unit, got "
            f"{decimal_separator!r} ({utf16_length(decimal_separator)} units)"
        )


def canonical_number(value: float, decimal_separator: str = ".") -> str:
    """The canonical text of a Double.

    LOCKED form: 17 significant digits, a decimal point always present, an
    uppercase `E`, the exponent sign always present, and at least two exponent
    digits zero-padded. No thousands separator. Negative zero normalises to
    positive zero.

    `decimal_separator` is the separator the HOST FORMATTER would have produced.
    On a comma-locale Windows machine VBA's formatting emits `,`; the encoder must
    normalise it back to `.` before hashing, or the same model would fingerprint
    differently in Riyadh and in Berlin. Making the separator an explicit argument
    is what turns that into a testable pure function rather than an assumption
    about the machine the code happens to run on.

    THE NORMALISATION IS POSITIONAL, NOT TEXTUAL. An earlier version replaced every
    occurrence of the separator, which is only safe while the separator happens not
    to occur elsewhere in scientific notation. It does occur elsewhere for `E`, for
    `+`, for `-` and for every digit, and those inputs produced malformed output.
    The decimal marker's POSITION is fixed by the accepted form - optional sign,
    one digit, marker - so the marker is located by index and exactly one character
    is rewritten. Any single UTF-16 code unit is therefore a safe separator, and
    the exponent marker, exponent sign and digits are untouchable by construction.

    NOTE ON PROOF SCOPE: passing `.` and `,` - or any other separator - here proves
    the REFERENCE normalisation semantics. It does NOT prove VBA `Format`/`Str`
    runtime behaviour under a foreign locale; that proof is Gate B's (plan
    section 24.1).
    """
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        raise FingerprintError(f"not a numeric value: {value!r}")
    number = float(value)
    if math.isnan(number) or math.isinf(number):
        raise FingerprintError(f"cannot canonically encode a non-finite value: {number!r}")
    if number == 0.0:
        # Covers -0.0, which compares equal to 0.0 but formats with a sign.
        number = 0.0

    _check_separator(decimal_separator)
    text = f"{number:.{_SIGNIFICAND_DIGITS_AFTER_POINT}E}"

    # Model what the host formatter would have handed back, then normalise it.
    # Both steps are needed: normalising a string that never carried the foreign
    # separator would prove nothing.
    hosted = apply_decimal_separator(text, decimal_separator)
    return normalise_decimal_separator(hosted, decimal_separator)


# ---------------------------------------------------------------------------
# Field, record, section and stream encoding
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class Field:
    """One encoded field of the canonical stream: `<TAG><LEN>:<VALUE>`.

    `LEN` is the UTF-16 code-unit count of `VALUE`, which is what makes the
    encoding SELF-DELIMITING: a reader knows exactly how many units to consume, so
    no character in the content can be mistaken for a separator. That is the
    property the eight collision probes exist to demonstrate; a delimiter-joined
    encoding has no such property, whichever delimiter it picks.
    """

    tag: str
    value: str

    def encode(self) -> str:
        return f"{self.tag}{utf16_length(self.value)}:{self.value}"


def text_field(value: str) -> Field:
    """A text field. The value is hashed exactly as given - never trimmed, never
    case-folded, never re-encoded."""
    if not isinstance(value, str):
        raise FingerprintError(f"text field requires a string, got {value!r}")
    return Field(TAG_TEXT, value)


def number_field(value: float, decimal_separator: str = ".") -> Field:
    """A Double field, in the canonical numeric form."""
    return Field(TAG_NUMBER, canonical_number(value, decimal_separator))


def integer_field(value: int) -> Field:
    """A STREAM integer: a structural count or a version, not model data.

    Kept distinct from `number_field` so that a count of 1 and a Double of 1 can
    never encode identically. Structural shape and model magnitude are different
    kinds of fact.
    """
    if not isinstance(value, int) or isinstance(value, bool):
        raise FingerprintError(f"stream integer requires an int, got {value!r}")
    if value < 0:
        raise FingerprintError(f"stream integer must be non-negative, got {value}")
    return Field(TAG_INTEGER, str(value))


def encode_fields(fields: Iterable[Field]) -> str:
    return "".join(field.encode() for field in fields)


def encode_record(fields: Sequence[Field]) -> str:
    """`record ::= F_I(field_count) field*`

    The field count is part of the hashed stream, so a record that gained or lost
    a field cannot coincide with a differently-shaped one.
    """
    return integer_field(len(fields)).encode() + encode_fields(fields)


def encode_section(name: str, records: Sequence[Sequence[Field]]) -> str:
    """`section ::= F_S(name) F_I(record_count) record*`"""
    parts = [text_field(name).encode(), integer_field(len(records)).encode()]
    parts.extend(encode_record(record) for record in records)
    return "".join(parts)


# ---------------------------------------------------------------------------
# Driver records and their locked ordering
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class DriverRecord:
    """One Cost Line or Risk record.

    `permanent_id` is both the sort key and the record's first field; `fields`
    carries everything after it, already tagged.
    """

    permanent_id: str
    fields: tuple[Field, ...]

    def as_record(self) -> tuple[Field, ...]:
        return (text_field(self.permanent_id),) + tuple(self.fields)


def utf16_sort_key(text: str) -> tuple[int, ...]:
    """Ordinal UTF-16 code-unit sort key.

    Matches `Option Compare Binary` / `StrComp(..., vbBinaryCompare)`. Python's own
    string ordering compares CODE POINTS, which disagrees with UTF-16 ordering for
    any astral character - `U+10000` sorts above `U+E000` by code point but below
    it by code unit, because its high surrogate is `U+D800`. Sorting drivers by the
    raw string would therefore drift from VBA on exactly the inputs the UTF-16
    vectors exist to cover.
    """
    return tuple(utf16_code_units(text))


def sort_driver_records(records: Iterable[DriverRecord]) -> list[DriverRecord]:
    """Ascending by Permanent ID, ordinal on UTF-16 code units.

    Never by row and never by digest: row order is presentation, and ordering by
    digest would make the ordering depend on the thing being computed.
    """
    return sorted(records, key=lambda record: utf16_sort_key(record.permanent_id))


def build_canonical_stream(
    *,
    version: int,
    header_fields: Sequence[Field],
    cost_records: Sequence[DriverRecord] = (),
    risk_records: Sequence[DriverRecord] = (),
    extra_sections: Sequence[tuple[str, Sequence[Sequence[Field]]]] = (),
) -> str:
    """`stream ::= F_S("PCCM-FP") F_I(version) section*`

    Sections are emitted in the locked order HEADER, COST, RISK. `version` is the
    fingerprint algorithm version, supplied by `spec/calc_contract.yaml` - this
    module deliberately does not hold a copy of it.

    `extra_sections` exists so a later phase can append its own sections after
    RISK without reopening this function. It is empty in Phase 5.
    """
    if not isinstance(version, int) or isinstance(version, bool) or version < 1:
        raise FingerprintError(f"fingerprint version must be a positive integer, got {version!r}")

    parts = [text_field(STREAM_TAG).encode(), integer_field(version).encode()]
    parts.append(encode_section(SECTION_HEADER, [tuple(header_fields)]))
    parts.append(
        encode_section(
            SECTION_COST, [record.as_record() for record in sort_driver_records(cost_records)]
        )
    )
    parts.append(
        encode_section(
            SECTION_RISK, [record.as_record() for record in sort_driver_records(risk_records)]
        )
    )
    for name, records in extra_sections:
        parts.append(encode_section(name, [tuple(record) for record in records]))
    return "".join(parts)


# ---------------------------------------------------------------------------
# Modular reduction
# ---------------------------------------------------------------------------
def reduce_exact(accumulator: int, unit: int, modulus: int) -> int:
    """The mathematics, in exact integer arithmetic.

    Python has arbitrary-precision integers, so this IS the definition. It is the
    oracle the Double-only mirror is checked against.
    """
    return (accumulator * FP_BASE + unit) % modulus


def reduce_double_only(accumulator: int, unit: int, modulus: int) -> int:
    """A faithful mirror of the LOCKED VBA reduction - `Double` arithmetic only.

    VBA cannot express `reduce_exact`. Its `Mod` operator and its `\\` integer
    division both use an EFFECTIVE INTEGRAL TYPE OF LONG on floating-point
    operands, and the recurrence intermediate reaches `281,320,423,161` -
    approximately 131 times the signed-Long maximum. `x Mod m` would overflow or
    silently mis-reduce. The locked VBA form is therefore:

        x = h * FP_BASE + u
        q = Fix(x / modulus)
        r = x - q * modulus
        If r >= modulus Then r = r - modulus
        If r < 0        Then r = r + modulus

    Every step below is float arithmetic for that reason, even though Python would
    do it exactly in `int`. Using `int` here would silently repair the very
    imprecision the two corrections exist to absorb, and the mirror would stop
    testing what it claims to test.

    Why the corrections are enough: `x < 2**53` so `x` is an exact Double, and
    `x / modulus <= 131` carries a relative error of at most `2**-53`, an absolute
    error under `1.5e-14`. `Fix` can therefore be off by AT MOST ONE in either
    direction, and one correction in each direction absorbs exactly that.

    THIS IS A PYTHON MIRROR, NOT A VBA EXECUTION. It proves the reference
    semantics are self-consistent. Parity with real VBA arithmetic is Gate B's
    (plan section 24.1).
    """
    x = float(accumulator) * float(FP_BASE) + float(unit)
    q = float(math.trunc(x / float(modulus)))
    r = x - q * float(modulus)
    if r >= modulus:
        r -= float(modulus)
    if r < 0:
        r += float(modulus)
    return int(r)


Reducer = Callable[[int, int, int], int]


def digest_code_units(units: Iterable[int], reducer: Reducer = reduce_exact) -> str:
    """Run the two-modulus recurrence over already-normalised code units."""
    h1, h2 = FP_INIT_1, FP_INIT_2
    for unit in units:
        unit = normalise_code_unit(unit)
        h1 = reducer(h1, unit, FP_MOD_1)
        h2 = reducer(h2, unit, FP_MOD_2)
    return f"{h1:08X}{h2:08X}"


def fingerprint(stream: str, reducer: Reducer = reduce_exact) -> str:
    """The 16-character uppercase digest of a canonical stream.

    Tags, lengths, the colon and the values are ALL hashed - the entire stream,
    UTF-16 code unit for UTF-16 code unit, nothing excluded.
    """
    return digest_code_units(utf16_code_units(stream), reducer)


PROBE_SECTION_NAME = "X"
"""Section name used to frame the locked collision probes.

The probes measure whether FIELD CONTENT can be made ambiguous, so they are framed
as one single-record section rather than as a full stream: the HEADER/COST/RISK
grammar would add identical prefixes to all eight and obscure nothing. The name is
part of the locked vector definition - changing it changes every expected digest.
"""


def fingerprint_probe(values: Sequence[str], reducer: Reducer = reduce_exact) -> str:
    """Digest one collision probe: a single-record section of text fields.

    This is the exact framing under which the eight locked probe digests of the
    plan were computed.
    """
    record = [text_field(value) for value in values]
    return fingerprint(encode_section(PROBE_SECTION_NAME, [record]), reducer)
