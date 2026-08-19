#!/usr/bin/env python3
"""PCCM Gate B: the corrected canonical Double encoder, proved on Linux.

NO WINDOWS RUN HAPPENS HERE, AND NONE CAN. Every assertion is either a statement
about SOURCE TEXT or a run of `vba_canonical_port`, which is a line-for-line
transcription of the shipped VBA. The expectations are the Python oracle's.

WHAT RUNTIME RUN 2 ESTABLISHED. On real Excel,
`Format$(number, "0.0000000000000000E+00")` produced fifteen correct significant
digits followed by zero padding:

    0.1          got 1.0000000000000000E-01  want 1.0000000000000001E-01
    1e-20        got 1.0000000000000000E-20  want 9.9999999999999995E-21
    0.1 + 0.2    got 3.0000000000000000E-01  want 3.0000000000000004E-01
    MAX_DOUBLE   got 1.7976931348623200E+308 want 1.7976931348623157E+308

Runs standalone or under pytest.
"""

from __future__ import annotations

import json
import math
import re
import struct
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from pccm_builder import calc_fingerprint as fp          # noqa: E402
import vba_canonical_port as port                        # noqa: E402

SRC_VBA = PCCM_ROOT / "src" / "vba"
MODULE = SRC_VBA / "modCalcFingerprint.bas"
BUILD = PCCM_ROOT / "build" / "phase5_cases.json"

MAX_DOUBLE = 1.7976931348623157e308
MIN_NORMAL = 2.2250738585072014e-308
MIN_SUBNORMAL = 5e-324
MAX_SUBNORMAL = struct.unpack(">d", bytes.fromhex("000FFFFFFFFFFFFF"))[0]


def _bits(value: float) -> str:
    return f"{struct.unpack('>Q', struct.pack('>d', value))[0]:016X}"


def _from_bits(bits: str) -> float:
    return struct.unpack(">d", bytes.fromhex(bits))[0]


def _deterministic_doubles(count: int) -> list[float]:
    """A fixed LCG, so this corpus is reproducible from this file alone."""
    state = 0x2545F4914F6CDD1D
    mask = (1 << 64) - 1
    out: list[float] = []
    while len(out) < count:
        state = (state * 6364136223846793005 + 1442695040888963407) & mask
        value = struct.unpack(">d", struct.pack(">Q", state))[0]
        if math.isfinite(value):
            out.append(value)
    return out


# ===========================================================================
# 1. the contract is 17 digits, and it has to be
# ===========================================================================
def test_01_the_reference_encoder_is_seventeen_significant_digits() -> None:
    assert fp._SIGNIFICAND_DIGITS_AFTER_POINT == 16, "one digit before the point, sixteen after"
    text = fp.canonical_number(0.1)
    assert re.fullmatch(r"-?\d\.\d{16}E[+-]\d{2,}", text), text
    assert len(text.split("E")[0].lstrip("-").replace(".", "")) == 17


def test_02_fifteen_significant_digits_would_collide() -> None:
    """The reason the contract cannot be relaxed to suit a formatter.

    Two DISTINCT binary64 values must never share a canonical numeric field, or
    two different models would fingerprint alike. 17 digits is the shortest
    width that round-trips every binary64; 15 provably does not.
    """
    collisions = []
    for value in _deterministic_doubles(4000):
        above = math.nextafter(value, math.inf)
        if not math.isfinite(above) or above == value:
            continue
        if f"{value:.14E}" == f"{above:.14E}" and f"{value:.16E}" != f"{above:.16E}":
            collisions.append((value, above))
    assert collisions, "no 15-digit collision found, which would itself be suspicious"
    value, above = collisions[0]
    assert value != above
    assert f"{value:.14E}" == f"{above:.14E}"
    assert fp.canonical_number(value) != fp.canonical_number(above)
    # And it is not rare.
    assert len(collisions) > 100, len(collisions)


def test_03_seventeen_digits_separates_every_neighbour_pair() -> None:
    probes = [1.0, 0.1, 1e-20, 1e20, math.pi, MIN_NORMAL, MIN_SUBNORMAL,
              9007199254740992.0, -0.1, 1e100, 1e-100]
    for value in probes:
        below = math.nextafter(value, -math.inf)
        above = math.nextafter(value, math.inf)
        texts = {fp.canonical_number(below), fp.canonical_number(value),
                 fp.canonical_number(above)}
        assert len(texts) == 3, f"{value!r} collapses onto a neighbour: {texts}"


# ===========================================================================
# 2. the shipped algorithm, against the oracle
# ===========================================================================
def test_04_the_named_boundaries_match_the_oracle_exactly() -> None:
    named = {
        "+0": 0.0, "-0": -0.0, "1": 1.0, "-1": -1.0,
        "0.1": 0.1, "-0.1": -0.1, "0.1 + 0.2": 0.1 + 0.2,
        "1e-20": 1e-20, "-1e-20": -1e-20,
        "MAX_DOUBLE": MAX_DOUBLE, "-MAX_DOUBLE": -MAX_DOUBLE,
        "min positive normal": MIN_NORMAL, "max subnormal": MAX_SUBNORMAL,
        "min positive subnormal": MIN_SUBNORMAL, "-min positive subnormal": -MIN_SUBNORMAL,
        "0.5": 0.5, "1.5": 1.5, "2.5": 2.5,
    }
    for label, value in named.items():
        ok, got = port.canonical_number(value, ".")
        assert ok, f"{label} was refused"
        assert got == fp.canonical_number(value, "."), f"{label}: {got}"
    # The four Run-2 failures, by their observed and their correct text.
    assert port.canonical_number(0.1, ".")[1] == "1.0000000000000001E-01"
    assert port.canonical_number(1e-20, ".")[1] == "9.9999999999999995E-21"
    assert port.canonical_number(0.1 + 0.2, ".")[1] == "3.0000000000000004E-01"
    assert port.canonical_number(MAX_DOUBLE, ".")[1] == "1.7976931348623157E+308"
    assert port.canonical_number(MIN_SUBNORMAL, ".")[1] == "4.9406564584124654E-324"


def test_05_powers_of_ten_and_their_neighbours_match_the_oracle() -> None:
    checked = 0
    for exponent in range(-323, 309):
        value = float(f"1e{exponent}")
        if value == 0.0 or not math.isfinite(value):
            continue
        for probe in (value, math.nextafter(value, math.inf),
                      math.nextafter(value, -math.inf)):
            if not math.isfinite(probe) or probe == 0.0:
                continue
            ok, got = port.canonical_number(probe, ".")
            assert ok and got == fp.canonical_number(probe, "."), (
                f"1e{exponent} neighbourhood: {probe!r} -> {got}"
            )
            checked += 1
    assert checked > 1800, checked


def test_06_a_broad_deterministic_corpus_matches_the_oracle() -> None:
    values = _deterministic_doubles(20000)
    for value in values:
        ok, got = port.canonical_number(value, ".")
        assert ok, f"{value!r} was refused"
        assert got == fp.canonical_number(value, "."), f"{value!r} -> {got}"
    # The corpus really does span the domain it claims to.
    exponents = {int(fp.canonical_number(v).split("E")[1]) for v in values}
    assert min(exponents) < -250 and max(exponents) > 250, (min(exponents), max(exponents))
    assert any(v < 0 for v in values) and any(v > 0 for v in values)


def test_07_both_separators_produce_byte_identical_text() -> None:
    """Separator invariance is now structural, not repaired."""
    for value in [0.0, -0.0, 1.0, -1.0, 0.1, 1e-20, MAX_DOUBLE, MIN_SUBNORMAL,
                  0.1 + 0.2, -9.87e-5] + _deterministic_doubles(2000):
        point = port.canonical_number(value, ".")
        comma = port.canonical_number(value, ",")
        assert point == comma, f"{value!r}: {point} vs {comma}"
        assert point[1] == fp.canonical_number(value, ",")
        # No separator character other than "." ever appears.
        assert "," not in point[1]


def test_08_exact_ties_round_half_to_even() -> None:
    """A binary64's expansion terminates, so the tie case is reachable."""
    ties = []
    for value in _deterministic_doubles(30000):
        digits = f"{abs(value):.30E}".split("E")[0].replace(".", "")
        if len(digits) > 18 and digits[17] == "5" and set(digits[18:]) <= {"0"}:
            ties.append(value)
    assert ties, "no exact tie found in the corpus"
    for value in ties:
        ok, got = port.canonical_number(value, ".")
        assert ok and got == fp.canonical_number(value, "."), f"tie {value!r} -> {got}"
    # And the rule really is half-to-EVEN, not half-up.
    assert f"{1.5:.0E}" == "2E+00" and f"{2.5:.0E}" == "2E+00"


def test_09_non_finite_and_bad_separators_are_refused() -> None:
    for bad in (float("nan"), float("inf"), float("-inf")):
        ok, text = port.canonical_number(bad, ".")
        assert not ok and text == "", f"{bad!r} was encoded"
    for separator in ("", "..", "ab"):
        ok, _ = port.canonical_number(1.0, separator)
        assert not ok, f"separator {separator!r} was accepted"
    # A non-BMP separator is two UTF-16 units and is refused for that reason.
    ok, _ = port.canonical_number(1.0, "\U0001F600")
    assert not ok


def test_10_negative_zero_normalises_to_positive_zero() -> None:
    ok, got = port.canonical_number(-0.0, ".")
    assert ok and got == "0.0000000000000000E+00"
    assert got == port.canonical_number(0.0, ".")[1]
    assert got == fp.canonical_number(-0.0)
    assert math.copysign(1.0, -0.0) == -1.0, "the input really was negative zero"


# ===========================================================================
# 3. the emitted parity corpus
# ===========================================================================
def _corpus() -> dict:
    if not BUILD.is_file():
        import subprocess
        subprocess.run([sys.executable, str(PCCM_ROOT / "builder" / "build_stage_a.py")],
                       check=True, capture_output=True)
    return json.loads(BUILD.read_text(encoding="utf-8"))["fingerprint"]["canonical_parity"]


def test_11_the_emitted_parity_corpus_spans_the_domain() -> None:
    parity = _corpus()
    vectors = parity["vectors"]
    assert len(vectors) > 2000, len(vectors)
    assert len({v["bits"] for v in vectors}) == len(vectors), "duplicate bit patterns"
    exponents = {int(v["expected"].split("E")[1]) for v in vectors}
    assert min(exponents) <= -320 and max(exponents) >= 300
    assert any(v["expected"].startswith("-") for v in vectors)
    for name in ("MAX_DOUBLE", "minimum positive subnormal", "minimum positive normal",
                 "maximum subnormal", "0.1 + 0.2", "1e-20", "+0", "-0"):
        assert any(v["label"] == name for v in vectors), f"{name} is not in the corpus"
    assert len(parity["neighbours"]) >= 8


def test_12_the_shipped_algorithm_reproduces_every_emitted_expectation() -> None:
    parity = _corpus()
    for vector in parity["vectors"]:
        value = _from_bits(vector["bits"])
        ok, got = port.canonical_number(value, ".")
        assert ok, f"{vector['label']} was refused"
        assert got == vector["expected"], (
            f"{vector['label']} [{vector['bits']}]: got {got}, expected {vector['expected']}"
        )
    for triple in parity["neighbours"]:
        texts = []
        for member in triple["members"]:
            ok, got = port.canonical_number(_from_bits(member["bits"]), ".")
            assert ok and got == member["expected"], triple["label"]
            texts.append(got)
        assert len(set(texts)) == 3, f"{triple['label']} collapses: {texts}"


def test_13_the_emitted_expectations_come_from_the_oracle_not_from_powershell() -> None:
    """An expectation produced by the algorithm it tests proves nothing."""
    generator = (PCCM_ROOT / "builder" / "pccm_builder" / "calc_cases.py").read_text(encoding="utf-8")
    assert 'fp.canonical_number(value, ".")' in generator
    harness = (PCCM_ROOT / "bootstrap" / "windows" / "phase5_gate_b_scenarios.ps1").read_text(
        encoding="utf-8")
    block = harness[harness.index("P5-DP."):harness.index("Add-Result 'P5-DP'")]
    # The harness reads the expectation and reconstructs the Double. It does not
    # compute a canonical string of its own.
    assert "[string]$vector.expected" in block
    assert "[BitConverter]::Int64BitsToDouble" in block, (
        "the harness parses a decimal literal instead of the bit pattern"
    )
    for forbidden in ("ToString('E16')", 'ToString("E16")', "-f '{0:E16}'", "Format-"):
        assert forbidden not in block, f"the harness computes its own expectation ({forbidden})"


# ===========================================================================
# 4. the port really is the shipped VBA
# ===========================================================================
def test_14_every_ported_routine_exists_in_the_shipped_module() -> None:
    """The port is worthless if it has drifted from the source it mirrors."""
    vba = MODULE.read_text(encoding="utf-8")
    pairs = {
        "decompose": "CalcFpDecompose",
        "limbs_from_mantissa": "CalcFpLimbsFromMantissa",
        "multiply_power": "CalcFpMultiplyPower",
        "integer_power": "CalcFpIntegerPower",
        "multiply_small": "CalcFpMultiplySmall",
        "limb_digits": "CalcFpLimbDigits",
        "plain_digits": "CalcFpPlainDigits",
        "round_significant": "CalcFpRoundSignificant",
        "increment_digits": "CalcFpIncrementDigits",
        "exponent_text": "CalcFpExponentText",
        "long_digits": "CalcFpLongDigits",
        "build_canonical": "CalcFpBuildCanonical",
        "marker_index": "CalcFpMarkerIndex",
        "canonical_number": "CalcFpCanonicalNumber",
    }
    ported = Path(__file__).resolve().parent / "vba_canonical_port.py"
    text = ported.read_text(encoding="utf-8")
    for python_name, vba_name in pairs.items():
        assert f"def {python_name}(" in text, f"the port lost {python_name}"
        assert f"Function {vba_name}(" in vba, f"the module lost {vba_name}"
    # The constants must agree numerically with the VBA.
    for constant, value in (("FP_LIMB_BASE", "10000000#"), ("FP_LIMB_DIGITS", "7"),
                            ("FP_SIGNIFICANT_DIGITS", "17"), ("FP_FRACTION_DIGITS", "16"),
                            ("FP_MANTISSA_DIGITS", "16"), ("FP_MAX_LIMBS", "200"),
                            ("FP_TWO_52", "4503599627370496#"),
                            ("FP_TWO_53", "9007199254740992#")):
        assert f"Const {constant} As" in vba and f"= {value}" in vba, constant
    assert port.FP_LIMB_BASE == 10000000.0
    assert port.FP_SIGNIFICANT_DIGITS == 17
    assert port.FP_MAX_LIMBS == 200
    # The chunk exponents the module uses must be the ones the port uses.
    assert "2#, exponent, 23" in vba and "5#, -exponent, 10" in vba
    assert 2 ** 23 < 10 ** 7 and 5 ** 10 < 10 ** 7, "a chunk factor exceeds one limb"


def test_15_the_limb_arithmetic_never_leaves_the_exact_integer_range() -> None:
    """(10^7 - 1) * 10^7 + carry stays below 2^53, which is why it is exact."""
    assert (10 ** 7 - 1) * (10 ** 7 - 1) + 10 ** 7 < 2 ** 53
    # And the assertion is live inside the port, so a corpus run would trip it.
    ported = (Path(__file__).resolve().parent / "vba_canonical_port.py").read_text(encoding="utf-8")
    assert 'assert product < FP_TWO_53' in ported
    # The widest expansion the algorithm can produce still fits the limb array.
    widest_digits = len(str(9007199254740991 * 5 ** 1126))
    assert widest_digits < port.FP_MAX_LIMBS * port.FP_LIMB_DIGITS, widest_digits


# ===========================================================================
# 5. negative controls
# ===========================================================================
def test_nc_01_the_old_format_implementation_fails_the_locked_vectors() -> None:
    """Fifteen significant digits then padding: exactly what Run 2 reported."""
    def format_dollar(value: float) -> str:
        """What VBA's Format$ effectively produced on real Excel.

        Its numeric-to-text conversion carries about 15 significant decimal
        digits; the sixteen fractional PLACEHOLDERS beyond that are filled with
        zeros rather than with recovered digits. Modelled as: round correctly to
        15 significant digits, then pad the mantissa out to 17. Every one of the
        five Run-2 observations below falls out of that, which is what makes it
        a model of the defect rather than a guess at it.
        """
        if value == 0.0:
            return "0.0000000000000000E+00"
        mantissa, _, exponent = f"{value:.14E}".partition("E")
        sign = "-" if mantissa.startswith("-") else ""
        digits = mantissa.lstrip("-").replace(".", "").ljust(17, "0")
        return f"{sign}{digits[0]}.{digits[1:]}E{exponent}"

    for value, observed in ((0.1, "1.0000000000000000E-01"),
                            (1e-20, "1.0000000000000000E-20"),
                            (0.1 + 0.2, "3.0000000000000000E-01"),
                            (MAX_DOUBLE, "1.7976931348623200E+308"),
                            (MIN_SUBNORMAL, "4.9406564584124700E-324")):
        assert format_dollar(value) == observed, (value, format_dollar(value))
        assert format_dollar(value) != fp.canonical_number(value), (
            "the defective formatter must disagree with the contract"
        )
        assert port.canonical_number(value, ".")[1] == fp.canonical_number(value), (
            "the corrected encoder must agree with it"
        )
    # And it was right on four of the ten, which is why ten vectors were not enough.
    for value in (0.0, 1.0, -1.0, 1e20):
        assert format_dollar(value) == fp.canonical_number(value)


def test_nc_02_a_fifteen_digit_canonicaliser_loses_identity() -> None:
    """Reducing the contract to fit the formatter maps distinct models together."""
    lost = 0
    for value in _deterministic_doubles(4000):
        above = math.nextafter(value, math.inf)
        if not math.isfinite(above):
            continue
        if f"{value:.14E}" == f"{above:.14E}":
            lost += 1
    assert lost > 100, f"only {lost} neighbour pairs collapse at 15 digits"
    # The same pairs stay distinct at 17.
    for value in _deterministic_doubles(500):
        above = math.nextafter(value, math.inf)
        if math.isfinite(above) and above != value:
            assert fp.canonical_number(value) != fp.canonical_number(above)


def test_nc_03_a_locale_substitution_cannot_alter_the_canonical_text() -> None:
    """Not one character of the output can come from a separator argument."""
    for separator in (".", ",", ";", "'", "٫"):   # incl. the Arabic decimal separator
        for value in (0.1, -1e-20, MAX_DOUBLE, 0.0, 1234.5678):
            ok, got = port.canonical_number(value, separator)
            assert ok, (value, separator)
            assert got == port.canonical_number(value, ".")[1], (value, separator)
            assert separator not in got or separator == "."
    # The exponent sign, the E and every digit are equally untouchable.
    ok, got = port.canonical_number(-9.87e-5, ",")
    assert ok and got == fp.canonical_number(-9.87e-5, ",") == "-9.8700000000000000E-05"


def test_nc_04_a_truncating_rounder_would_be_caught() -> None:
    """Truncation is not rounding, and half-up is not half-even."""
    def truncate(value: float) -> str:
        digits = f"{abs(value):.30E}".split("E")[0].replace(".", "")[:17]
        exponent = f"{value:.16E}".split("E")[1]
        sign = "-" if value < 0 else ""
        return f"{sign}{digits[0]}.{digits[1:]}E{exponent}"

    disagreements = [v for v in _deterministic_doubles(3000)
                     if truncate(v) != fp.canonical_number(v)]
    assert disagreements, "truncation must differ from correct rounding somewhere"
    for value in disagreements[:20]:
        assert port.canonical_number(value, ".")[1] == fp.canonical_number(value)


def test_nc_05_a_non_finite_value_is_never_encoded() -> None:
    for bad in (float("nan"), float("inf"), float("-inf")):
        ok, _ = port.canonical_number(bad, ".")
        assert not ok
        try:
            fp.canonical_number(bad)
        except fp.FingerprintError:
            continue
        raise AssertionError(f"the oracle encoded {bad!r}")


if __name__ == "__main__":
    import pytest
    raise SystemExit(pytest.main([__file__, "-q"]))
