#!/usr/bin/env python3
"""PCCM Phase 5 Gate-A Step-1 tests: the Calculation Input Fingerprint reference.

EVERY EXPECTED VALUE IN THIS FILE IS AN INDEPENDENT LITERAL, transcribed from
docs/phase5_plan.md Revision E. Nothing here is derived from the implementation
under test: a test that asks the code what it produces and then asserts it
produces that cannot fail.

--------------------------------------------------------------------------------
PROOF SCOPE - READ BEFORE STRENGTHENING A CLAIM HERE
--------------------------------------------------------------------------------
This suite runs on Linux. THERE IS NO VBA INTERPRETER, NO EXCEL AND NO `AscW`
HERE, so no test in this file proves anything about VBA runtime behaviour, and
none may be described as doing so (plan section 21.0, erratum E3).

What these tests DO prove:
  * the Python reference implementation matches the locked literal vectors;
  * the Python mirror of the locked VBA Double-only reducer agrees with exact
    integer arithmetic, so the locked reduction is self-consistent;
  * the canonical numeric encoder is a pure function of its arguments, including
    the decimal separator.

What they DO NOT prove, and what is reserved for Gate B on real Windows Excel
(plan section 24.1):
  * that VBA executes any of this;
  * that VBA `Format`/`Str` behave as assumed under a comma locale;
  * that `AscW` sign normalisation behaves as assumed on target;
  * that the real Double-only reducer reproduces these remainders.

Runs standalone or under pytest.
"""

from __future__ import annotations

import random
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import load_calc_contract  # noqa: E402
from pccm_builder.calc_fingerprint import (  # noqa: E402
    FP_BASE,
    FP_INIT_1,
    FP_INIT_2,
    FP_MOD_1,
    FP_MOD_2,
    SECTION_ORDER,
    STREAM_TAG,
    DriverRecord,
    FingerprintError,
    build_canonical_stream,
    canonical_number,
    encode_record,
    encode_section,
    fingerprint,
    fingerprint_probe,
    integer_field,
    normalise_code_unit,
    number_field,
    reduce_double_only,
    reduce_exact,
    sort_driver_records,
    text_field,
    utf16_code_units,
    utf16_length,
    utf16_sort_key,
)

CALC_CONTRACT_PATH = PCCM_ROOT / "spec" / "calc_contract.yaml"

US = "\u001f"
NUL = "\u0000"
LF = "\u000a"

# ---------------------------------------------------------------------------
# Locked literals - plan sections 11.3 to 11.7
# ---------------------------------------------------------------------------
REFERENCE_STREAM = (
    "S7:PCCM-FPI1:1S6:HEADERI1:1I1:4"
    "N22:2.0260000000000000E+03N22:2.0260000000000000E+03"
    "N22:1.0000000000000000E+00N22:1.0000000000000001E-01"
    "S4:COSTI1:1I1:9S6:CL-001S10:Triangular"
    "N22:1.0000000000000000E+01N22:8.0000000000000000E+01"
    "N22:1.5000000000000000E+02N22:1.0000000000000000E+02"
    "N22:1.0000000000000000E+00N22:1.0000000000000000E+00"
    "N22:1.0000000000000000E+00"
    "S4:RISKI1:0"
)
REFERENCE_CODE_UNITS = 366
REFERENCE_DIGEST = "50B6EB0E26857EA7"

NUMERIC_VECTORS = (
    (0.0, "0.0000000000000000E+00"),
    (-0.0, "0.0000000000000000E+00"),
    (1.0, "1.0000000000000000E+00"),
    (-1.0, "-1.0000000000000000E+00"),
    (0.1, "1.0000000000000001E-01"),
    (1e-20, "9.9999999999999995E-21"),
    (1e20, "1.0000000000000000E+20"),
    (0.1 + 0.2, "3.0000000000000004E-01"),
    (1.7976931348623157e308, "1.7976931348623157E+308"),
    (5e-324, "4.9406564584124654E-324"),
)

COLLISION_PROBES = (
    (("A:B", "C"), "041ACBD05C7BF72C"),
    (("A", "B:C"), "52704E9A542869CA"),
    (("AB", "C"), "0C8A057A0BE7EB51"),
    (("A", "B", "C"), "7674F1C35E639F98"),
    (("A" + US + "B", "C"), "7D26D4C95587DE0C"),
    (("A" + NUL + "B", "C"), "0821AFB0608291C8"),
    (("A" + LF + "B", "C"), "5B4CA2E133AD91A2"),
    (("A", US, "C"), "101504AC7803B226"),
)

# (modulus, h, u, x, q, r) - every x exceeds Long.MaxValue, which is the point.
REDUCTION_VECTORS = (
    (2147483647, 2147483646, 65535, 281320423161, 131, 65404),
    (2147483629, 2147483628, 65535, 281320420803, 131, 65404),
    (2147483647, 1234567890, 41, 161728393631, 75, 667120106),
    (2147483629, 1234567890, 41, 161728393631, 75, 667121456),
)

LONG_MAX = 2147483647


def _reference_stream(version: int) -> str:
    """Golden case 1: Base 2026, Start 2026, Duration 1, r = 0.10; one Cost Line."""
    header = [number_field(2026), number_field(2026), number_field(1), number_field(0.10)]
    cost = DriverRecord(
        "CL-001",
        (
            text_field("Triangular"),
            number_field(10),
            number_field(80),
            number_field(150),
            number_field(100),
            number_field(1),
            number_field(1),
            number_field(1),
        ),
    )
    return build_canonical_stream(version=version, header_fields=header, cost_records=[cost])


def _fp_version() -> int:
    """FP_VERSION is PROJECTED from the calculation contract, never hardcoded here."""
    return load_calc_contract(CALC_CONTRACT_PATH).fingerprint_version


# ---------------------------------------------------------------------------
# constants
# ---------------------------------------------------------------------------
def test_hash_constants_are_the_locked_literals() -> None:
    assert FP_BASE == 131
    assert FP_MOD_1 == 2147483647
    assert FP_MOD_2 == 2147483629
    assert FP_INIT_1 == 1
    assert FP_INIT_2 == 1
    assert STREAM_TAG == "PCCM-FP"
    assert SECTION_ORDER == ("HEADER", "COST", "RISK")


def test_both_moduli_are_prime_and_distinct() -> None:
    def is_prime(n: int) -> bool:
        if n < 2 or (n % 2 == 0 and n != 2):
            return False
        factor = 3
        while factor * factor <= n:
            if n % factor == 0:
                return False
            factor += 2
        return True

    assert FP_MOD_1 != FP_MOD_2
    assert is_prime(FP_MOD_1)
    assert is_prime(FP_MOD_2)


def test_accumulators_start_at_one_so_a_leading_nul_is_not_absorbed() -> None:
    """A zero initial state would make NUL-prefixed streams hash identically."""
    assert FP_INIT_1 == 1 and FP_INIT_2 == 1
    assert fingerprint(NUL + "A") != fingerprint("A")


def test_fingerprint_version_is_projected_from_the_calc_contract() -> None:
    """The ALGORITHM VERSION is a layout fact and belongs to the contract.

    The mathematics is not, and does not appear there. The fingerprint module must
    hold no copy of the version number.
    """
    import pccm_builder.calc_fingerprint as module

    assert _fp_version() == 1
    assert not hasattr(module, "FP_VERSION")


# ---------------------------------------------------------------------------
# UTF-16 semantics
# ---------------------------------------------------------------------------
def test_a_code_unit_above_u7fff_normalises_from_a_signed_ascw_result() -> None:
    """VBA `AscW` returns a SIGNED 16-bit Integer, so U+9AD8 comes back negative."""
    char = "高"
    assert ord(char) == 39640
    assert list(utf16_code_units(char)) == [39640]

    signed = 39640 - 65536
    assert signed == -25896
    assert normalise_code_unit(signed) == 39640
    assert normalise_code_unit(39640) == 39640


def test_a_non_bmp_character_contributes_two_surrogate_code_units() -> None:
    char = "\U0001f600"
    assert len(char) == 1                       # one Python code point
    assert utf16_length(char) == 2              # two UTF-16 code units
    assert list(utf16_code_units(char)) == [0xD83D, 0xDE00]


def test_length_prefixes_count_utf16_units_not_python_code_points() -> None:
    """`len()` would disagree with VBA on any astral character.

    "A" + emoji is 2 code points and 3 UTF-16 code units, so the prefix must be 3.
    """
    value = "A\U0001f600"
    assert len(value) == 2
    assert utf16_length(value) == 3
    assert text_field(value).encode() == "S3:" + value


def test_utf16_sort_key_disagrees_with_python_code_point_order() -> None:
    """Ordinal UTF-16 ordering is not Python's string ordering.

    U+10000's high surrogate is U+D800, which sorts BELOW U+E000, while its code
    point 0x10000 sorts above. Sorting Permanent IDs by the raw string would drift
    from `StrComp(..., vbBinaryCompare)` on exactly these inputs.
    """
    private_use = ""
    astral = "\U00010000"
    assert sorted([private_use, astral]) == [private_use, astral]
    assert sorted([private_use, astral], key=utf16_sort_key) == [astral, private_use]


def test_driver_records_sort_by_permanent_id_on_utf16_code_units() -> None:
    records = [
        DriverRecord("CL-010", ()),
        DriverRecord("CL-002", ()),
        DriverRecord("CL-001", ()),
    ]
    assert [r.permanent_id for r in sort_driver_records(records)] == [
        "CL-001",
        "CL-002",
        "CL-010",
    ]


def test_driver_sorting_uses_utf16_order_not_python_string_order() -> None:
    """The ASCII case above cannot tell the two orderings apart; this one can.

    Sorting the records by the raw string would put U+E000 first, matching Python's
    code-point order and disagreeing with `StrComp(..., vbBinaryCompare)`.
    """
    private_use = ""
    astral = "\U00010000"
    records = [DriverRecord(private_use, ()), DriverRecord(astral, ())]
    assert [r.permanent_id for r in sort_driver_records(records)] == [astral, private_use]
    assert sorted([private_use, astral]) == [private_use, astral]


def test_row_order_does_not_change_the_stream() -> None:
    """Presentation order is excluded: the same drivers reversed hash identically."""
    a = DriverRecord("CL-001", (number_field(1),))
    b = DriverRecord("CL-002", (number_field(2),))
    forward = build_canonical_stream(version=1, header_fields=[], cost_records=[a, b])
    reversed_ = build_canonical_stream(version=1, header_fields=[], cost_records=[b, a])
    assert forward == reversed_


# ---------------------------------------------------------------------------
# canonical numeric encoding
# ---------------------------------------------------------------------------
def test_the_ten_locked_numeric_encodings() -> None:
    for value, expected in NUMERIC_VECTORS:
        assert canonical_number(value) == expected, f"{value!r} encoded wrongly"


def test_negative_zero_normalises_to_positive_zero() -> None:
    assert canonical_number(-0.0) == canonical_number(0.0) == "0.0000000000000000E+00"


def test_canonical_numbers_are_separator_invariant() -> None:
    """LOCALE TEST - PYTHON REFERENCE ONLY.

    Gate A does not execute VBA. This asserts that the REFERENCE normalisation
    semantics are a pure function of the arguments: a host formatter that emitted
    `,` produces the same canonical text as one that emitted `.`.

    IT DOES NOT PROVE VBA `Format`/`Str` RUNTIME BEHAVIOUR under a comma locale.
    That proof is reserved for Windows Gate B (plan section 24.1).
    """
    for value, expected in NUMERIC_VECTORS:
        assert canonical_number(value, ".") == expected
        assert canonical_number(value, ",") == expected
        assert canonical_number(value, ",") == canonical_number(value, ".")
        assert "," not in canonical_number(value, ",")


def test_non_finite_values_are_refused_not_encoded() -> None:
    for bad in (float("nan"), float("inf"), float("-inf")):
        try:
            canonical_number(bad)
        except FingerprintError:
            continue
        raise AssertionError(f"{bad!r} was silently encoded")


def test_a_multi_character_separator_is_refused() -> None:
    try:
        canonical_number(1.0, "..")
    except FingerprintError:
        return
    raise AssertionError("a multi-character decimal separator was accepted")


# ---------------------------------------------------------------------------
# field, record and section grammar
# ---------------------------------------------------------------------------
def test_field_encoding_is_tag_length_colon_value() -> None:
    assert text_field("CL-001").encode() == "S6:CL-001"
    assert integer_field(0).encode() == "I1:0"
    assert integer_field(12).encode() == "I2:12"
    assert number_field(1).encode() == "N22:1.0000000000000000E+00"


def test_a_stream_integer_and_a_double_never_encode_identically() -> None:
    """Structural shape and model magnitude are different kinds of fact."""
    assert integer_field(1).encode() != number_field(1).encode()


def test_record_and_section_carry_their_own_counts() -> None:
    record = (text_field("A"), text_field("B"))
    assert encode_record(record) == "I1:2S1:AS1:B"
    assert encode_section("COST", [record]) == "S4:COSTI1:1I1:2S1:AS1:B"
    assert encode_section("RISK", []) == "S4:RISKI1:0"


def test_negative_stream_integers_are_refused() -> None:
    try:
        integer_field(-1)
    except FingerprintError:
        return
    raise AssertionError("a negative stream integer was accepted")


# ---------------------------------------------------------------------------
# the locked reference vector
# ---------------------------------------------------------------------------
def test_the_reference_stream_is_reproduced_exactly() -> None:
    assert _reference_stream(_fp_version()) == REFERENCE_STREAM


def test_the_reference_stream_is_366_utf16_code_units() -> None:
    assert utf16_length(REFERENCE_STREAM) == REFERENCE_CODE_UNITS
    assert utf16_length(_reference_stream(_fp_version())) == REFERENCE_CODE_UNITS


def test_the_reference_digest_is_the_locked_literal() -> None:
    """Gate A: the PYTHON reference produces this. Real VBA is proven at Gate B."""
    assert fingerprint(REFERENCE_STREAM) == REFERENCE_DIGEST
    assert fingerprint(_reference_stream(_fp_version())) == REFERENCE_DIGEST


def test_the_reference_digest_is_unchanged_under_the_double_only_reducer() -> None:
    assert fingerprint(REFERENCE_STREAM, reduce_double_only) == REFERENCE_DIGEST


def test_the_digest_is_sixteen_uppercase_hex_characters() -> None:
    digest = fingerprint(REFERENCE_STREAM)
    assert len(digest) == 16
    assert digest == digest.upper()
    assert all(ch in "0123456789ABCDEF" for ch in digest)


def test_every_code_unit_of_the_stream_is_hashed() -> None:
    """Tags, lengths and the colon are inside the hash, not framing around it."""
    for index in range(len(REFERENCE_STREAM)):
        swap = "~" if REFERENCE_STREAM[index] != "~" else "!"
        mutated = REFERENCE_STREAM[:index] + swap + REFERENCE_STREAM[index + 1 :]
        assert fingerprint(mutated) != REFERENCE_DIGEST, f"position {index} not hashed"


# ---------------------------------------------------------------------------
# collision probes
# ---------------------------------------------------------------------------
def test_the_eight_locked_probe_digests() -> None:
    for values, expected in COLLISION_PROBES:
        assert fingerprint_probe(values) == expected, f"{values!r} digested wrongly"


def test_the_eight_probe_digests_are_all_distinct() -> None:
    digests = [fingerprint_probe(values) for values, _ in COLLISION_PROBES]
    assert len(set(digests)) == 8


def test_the_probe_digests_are_unchanged_under_the_double_only_reducer() -> None:
    for values, expected in COLLISION_PROBES:
        assert fingerprint_probe(values, reduce_double_only) == expected


def test_delimiter_joins_collide_where_the_plan_says_they_do() -> None:
    """The CORRECTED collision analysis - plan section 11.7.

    Revision C claimed rows 1-2, 3-5 and 4-8 collide under a U+001F join. They do
    not. The real picture, by direct analysis of the flattened streams:

        U+001F join     4 <-> 5 only - ["A","B","C"] and ["A"+US+"B","C"] both
                        flatten to A US B US C
        colon join      1 <-> 2 and 1 <-> 4
        length-prefixed none

    This test asserts the corrected pairs, so the old wrong claim cannot regress
    into the suite.
    """

    def collisions(join) -> list[tuple[int, int]]:
        seen: dict[str, int] = {}
        found: list[tuple[int, int]] = []
        for number, (values, _) in enumerate(COLLISION_PROBES, start=1):
            flat = join(values)
            if flat in seen:
                found.append((seen[flat], number))
            else:
                seen[flat] = number
        return found

    assert collisions(lambda values: US.join(values)) == [(4, 5)]
    assert collisions(lambda values: ":".join(values)) == [(1, 2), (1, 4)]
    assert collisions(lambda values: "".join(text_field(v).encode() for v in values)) == []


def test_the_specific_u001f_collision_is_real() -> None:
    assert US.join(("A", "B", "C")) == US.join(("A" + US + "B", "C"))


def test_the_pairs_revision_c_named_do_not_actually_collide() -> None:
    """Guards the correction itself: 1-2, 3-5 and 4-8 are distinct under U+001F."""
    flat = [US.join(values) for values, _ in COLLISION_PROBES]
    for a, b in ((1, 2), (3, 5), (4, 8)):
        assert flat[a - 1] != flat[b - 1], f"probes {a} and {b} unexpectedly collide"


# ---------------------------------------------------------------------------
# the Double-only reduction
# ---------------------------------------------------------------------------
def test_the_four_locked_reduction_vectors() -> None:
    for modulus, h, u, x, q, r in REDUCTION_VECTORS:
        assert h * FP_BASE + u == x
        assert x // modulus == q
        assert x - q * modulus == r
        assert reduce_exact(h, u, modulus) == r
        assert reduce_double_only(h, u, modulus) == r


def test_every_reduction_vector_exceeds_the_signed_long_maximum() -> None:
    """This is why the vectors exist: a native VBA `Mod` fails on every one."""
    for _, _, _, x, _, _ in REDUCTION_VECTORS:
        assert x > LONG_MAX


def test_the_maximum_intermediate_arithmetic_is_stated_exactly() -> None:
    """Erratum E2. 281,320,423,161 is NOT equal to 131 x Long.MaxValue.

    The two exact statements, asserted here so the corrected prose cannot regress:
    """
    assert FP_BASE * LONG_MAX == 281320357757
    assert 281320423161 == FP_BASE * LONG_MAX + 65404
    assert 281320423161 != FP_BASE * LONG_MAX


def test_the_maximum_intermediate_stays_inside_exact_double_range() -> None:
    """What makes the Double-only reduction sound at all."""
    assert 281320423161 < 2**53
    assert FP_BASE * LONG_MAX < 2**53
    assert (FP_MOD_1 - 1) * FP_BASE + 65535 == 281320423161


def test_the_double_only_mirror_equals_exact_integer_arithmetic_on_boundaries() -> None:
    for modulus in (FP_MOD_1, FP_MOD_2):
        accumulators = (0, 1, 2, 3, modulus // 2, modulus - 3, modulus - 2, modulus - 1)
        units = (0, 1, 2, 127, 128, 32767, 32768, 65534, 65535)
        for h in accumulators:
            for u in units:
                assert reduce_double_only(h, u, modulus) == reduce_exact(h, u, modulus), (
                    h,
                    u,
                    modulus,
                )


def test_the_double_only_mirror_equals_exact_integer_arithmetic_on_a_random_sweep() -> None:
    """A Python mirror of the locked VBA reducer, not a VBA execution (E3)."""
    rng = random.Random(20260817)
    for _ in range(20000):
        modulus = rng.choice((FP_MOD_1, FP_MOD_2))
        h = rng.randrange(0, modulus)
        u = rng.randrange(0, 65536)
        assert reduce_double_only(h, u, modulus) == reduce_exact(h, u, modulus), (h, u, modulus)


def test_the_reducer_output_is_always_a_valid_next_accumulator() -> None:
    for modulus in (FP_MOD_1, FP_MOD_2):
        for h, u in ((0, 0), (modulus - 1, 65535), (1, 65535), (modulus - 1, 0)):
            r = reduce_double_only(h, u, modulus)
            assert 0 <= r < modulus


def test_both_corrections_recover_the_exact_remainder_when_fix_is_off_by_one() -> None:
    """The corrections are DEFENSIVE, and this is what proves they are right.

    On IEEE-754 doubles `Fix(x / m)` happens never to be off by one over this
    range, so no input can drive either correction branch - a behavioural test
    cannot reach them, and removing one would go unnoticed by black-box testing.

    The design still requires them, because the guarantee is an error bound, not
    an observation: `x / m <= 131` carries a relative error of at most 2**-53, so
    `Fix` MAY be off by at most one in either direction. This test forces exactly
    that, by offsetting `q` by -1 and by +1, and asserts a single correction in
    each direction recovers the exact remainder.
    """

    def reduce_with_offset(h: int, u: int, modulus: int, offset: int) -> int:
        x = float(h) * float(FP_BASE) + float(u)
        q = float(int(x / modulus)) + offset
        r = x - q * float(modulus)
        if r >= modulus:
            r -= float(modulus)
        if r < 0:
            r += float(modulus)
        return int(r)

    cases = [(h, u, m) for m, h, u, _, _, _ in REDUCTION_VECTORS]
    cases += [(0, 0, FP_MOD_1), (1, 65535, FP_MOD_2), (FP_MOD_1 - 1, 0, FP_MOD_1)]
    for h, u, modulus in cases:
        exact = reduce_exact(h, u, modulus)
        for offset in (-1, 0, 1):
            assert reduce_with_offset(h, u, modulus, offset) == exact, (h, u, modulus, offset)


def test_the_double_only_mirror_keeps_the_locked_shape() -> None:
    """A STATIC source check on the mirror - Gate A's own kind of proof.

    Because neither correction branch is behaviourally reachable, the only way to
    keep them from being "simplified away" is to assert the locked shape is still
    present, and that no exact-integer operator has crept into the mirror. A mirror
    that used `%` would silently stop mirroring anything.
    """
    import inspect

    import pccm_builder.calc_fingerprint as module

    source = inspect.getsource(module.reduce_double_only)
    body = source.split('"""')[-1]

    assert "math.trunc" in body, "the mirror must truncate toward zero, as VBA Fix does"
    assert "if r >= modulus:" in body, "the >= modulus correction is missing"
    assert "if r < 0:" in body, "the < 0 correction is missing"
    assert "%" not in body, "the mirror must not use exact integer modulo"
    assert "//" not in body, "the mirror must not use exact integer division"
    assert body.count("float(") >= 4, "every step of the mirror must be float arithmetic"


def test_a_long_typed_reduction_would_not_reproduce_the_digest() -> None:
    """The failure mode the locked reduction exists to prevent.

    A VBA `Mod` implementation wraps the intermediate into signed 32-bit range
    before reducing. Simulated here, it produces a different digest - so the
    reference digest is itself evidence that no `Mod` crept in.
    """

    def long_wrapped(accumulator: int, unit: int, modulus: int) -> int:
        x = accumulator * FP_BASE + unit
        wrapped = ((x + 2**31) % 2**32) - 2**31
        return wrapped % modulus

    assert fingerprint(REFERENCE_STREAM, long_wrapped) != REFERENCE_DIGEST


# ---------------------------------------------------------------------------
# runner
# ---------------------------------------------------------------------------
def _run_all() -> int:
    tests = sorted(
        (name, fn) for name, fn in globals().items()
        if name.startswith("test_") and callable(fn)
    )
    failures = 0
    print("PCCM Phase 5 Gate-A Step-1 fingerprint reference tests")
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
