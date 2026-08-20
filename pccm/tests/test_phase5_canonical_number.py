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

import contextlib
import json
import math
import re
import struct
import sys
import tempfile
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from pccm_builder import calc_fingerprint as fp          # noqa: E402
import vba_canonical_port as port                        # noqa: E402

SRC_VBA = PCCM_ROOT / "src" / "vba"
MODULE = SRC_VBA / "modCalcFingerprint.bas"
# The repository build directory. Named ONLY so the regression below can
# plant a stale artifact there and prove the tests ignore it.
REPO_BUILD = PCCM_ROOT / "build"
REPO_CASES = REPO_BUILD / "phase5_cases.json"

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
_FRESH: dict = {}


def _fresh_cases() -> dict:
    """`phase5_cases.json` EMITTED HERE, into a test-owned temporary directory.

    THE DEFECT THIS REPLACES. This used to read `pccm/build/phase5_cases.json`
    and rebuild only when that file was ABSENT:

        if not BUILD.is_file():
            subprocess.run([... build_stage_a.py ...])
        return json.loads(BUILD.read_text(...))["fingerprint"]["canonical_parity"]

    which treats "a file exists" as "an artifact generated from THIS source". It
    is not. On a Windows checkout carrying `build/` from the previous commit,
    both parity tests failed with `KeyError: 'canonical_parity'` - not a
    production defect, not a runtime defect, just a test reading a stale file.
    A rebuild made the same suite pass, which is exactly the operator knowledge
    a test must never require.

    So nothing here consults `build/` at all. The artifact is emitted by the
    real builder - `emit_calc_artifacts`, the same entry point `build_stage_a.py`
    calls - into a fresh temporary directory, once per session. There is no
    second implementation of the corpus, and no repository state is read or
    written.

    This is the pattern `tests/test_phase5_gate_b_harness_source.py::_emitted`
    already used, and whose docstring already said "Never read from `build/`".
    """
    if _FRESH:
        return _FRESH["cases"]
    from pccm_builder import emit_calc_artifacts, load_calc_contract, load_spec

    directory = Path(tempfile.mkdtemp(prefix="pccm-canonical-"))
    spec = load_spec(PCCM_ROOT / "spec" / "workbook.yaml")
    calc = load_calc_contract(PCCM_ROOT / "spec" / "calc_contract.yaml")
    artifacts = emit_calc_artifacts(directory, spec, calc)
    assert artifacts.cases_path.is_file(), "the builder produced no phase5_cases.json"
    # An emitted artifact that returns early when a key is missing proves nothing;
    # it would pass loudest exactly when the emission is broken.
    document = json.loads(artifacts.cases_path.read_text(encoding="utf-8"))
    assert "fingerprint" in document, "the freshly emitted corpus has no fingerprint section"
    _FRESH.update(cases=document, directory=directory, path=artifacts.cases_path)
    return document


def _corpus() -> dict:
    document = _fresh_cases()
    assert "canonical_parity" in document["fingerprint"], (
        "the freshly emitted corpus carries no canonical_parity section"
    )
    return document["fingerprint"]["canonical_parity"]


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


# ===========================================================================
# 6. BUILD-ARTIFACT ISOLATION
# ===========================================================================
# The exact Windows observation this section closes. On a checkout carrying
# `pccm/build/phase5_cases.json` from the PREVIOUS commit, the first full-suite
# run after pulling the canonical-encoder correction reported:
#
#     1609 passed, 2 failed
#     test_11_the_emitted_parity_corpus_spans_the_domain            KeyError: 'canonical_parity'
#     test_12_the_shipped_algorithm_reproduces_every_emitted_...    KeyError: 'canonical_parity'
#
# Rebuilding Stage A by hand made the same suite report 1611 passed. So the
# production encoder was never implicated: a test was reading a stale file and
# only rebuilding when the file was ABSENT, which treats "a file exists" as
# "an artifact generated from this source".
#
# A suite that requires the operator to know it must rebuild first is not a
# suite that can be trusted after `git pull`.
def _previous_schema_document() -> str:
    """A VALID corpus in the previous schema: no `canonical_parity` section."""
    document = json.loads(json.dumps(_fresh_cases()))
    del document["fingerprint"]["canonical_parity"]
    assert "canonical_parity" not in document["fingerprint"]
    assert document["fingerprint"]["reference"]["digest"], "still a real corpus"
    return json.dumps(document, indent=2)


@contextlib.contextmanager
def _planted_repository_artifact(payload: str):
    """Plant `payload` at pccm/build/phase5_cases.json, then restore exactly.

    The session cache is cleared inside the block, so the helper under test has
    to go and fetch the corpus again while the stale file is in place. Without
    that, a previously cached corpus would make the test pass for the wrong
    reason.
    """
    REPO_BUILD.mkdir(parents=True, exist_ok=True)
    original = REPO_CASES.read_bytes() if REPO_CASES.is_file() else None
    cached = dict(_FRESH)
    REPO_CASES.write_text(payload, encoding="utf-8")
    _FRESH.clear()
    try:
        yield
    finally:
        _FRESH.clear()
        _FRESH.update(cached)
        if original is None:
            REPO_CASES.unlink()
        else:
            REPO_CASES.write_bytes(original)


def test_16_a_stale_repository_artifact_is_ignored_entirely() -> None:
    """The Windows observation, reproduced and then defeated."""
    stale = _previous_schema_document()
    assert json.loads(stale), "the planted artifact must be syntactically valid"
    with _planted_repository_artifact(stale):
        # 1. The stale file really is there, valid, and really lacks the section.
        planted = json.loads(REPO_CASES.read_text(encoding="utf-8"))
        assert "canonical_parity" not in planted["fingerprint"]

        # 2. The OLD helper would have raised exactly what Windows reported.
        try:
            planted["fingerprint"]["canonical_parity"]
        except KeyError as error:
            assert "canonical_parity" in str(error)
        else:
            raise AssertionError("the planted artifact does not reproduce the defect")

        # 3. The CURRENT helper obtains a fresh corpus regardless.
        parity = _corpus()
        assert len(parity["vectors"]) > 2000
        assert parity["neighbours"]

        # 4. And the two tests that failed on Windows both pass with it planted.
        test_11_the_emitted_parity_corpus_spans_the_domain()
        test_12_the_shipped_algorithm_reproduces_every_emitted_expectation()

    # The repository artifact is exactly as it was found.
    if REPO_CASES.is_file():
        assert "fingerprint" in json.loads(REPO_CASES.read_text(encoding="utf-8"))


def test_17_a_corrupt_repository_artifact_cannot_become_an_oracle() -> None:
    """Not merely stale: unreadable, truncated, or the wrong shape entirely."""
    for payload in ("{ this is not json",
                    "{}",
                    '{"fingerprint": {}}',
                    '{"fingerprint": {"canonical_parity": {"vectors": [], "neighbours": []}}}',
                    ""):
        with _planted_repository_artifact(payload):
            parity = _corpus()
            assert len(parity["vectors"]) > 2000, (
                f"a corrupt artifact reached the tests: {payload[:40]!r}"
            )
            # The tampered expectation in the fourth payload must not be adopted.
            assert parity["vectors"][0]["expected"] == fp.canonical_number(
                _from_bits(parity["vectors"][0]["bits"])
            )
            test_12_the_shipped_algorithm_reproduces_every_emitted_expectation()


def test_18_the_canonical_suite_never_reads_or_writes_the_repository_build() -> None:
    """Isolation, stated as a property of the source and of the behaviour."""
    text = Path(__file__).read_text(encoding="utf-8")
    # The corpus comes from the real emitter, into a temporary directory.
    assert "emit_calc_artifacts(directory, spec, calc)" in text
    assert 'tempfile.mkdtemp(prefix="pccm-canonical-")' in text
    # No path under the repository build directory is ever read as a corpus.
    # Judged on the two helpers' CODE: the docstring deliberately quotes the
    # defective form and names the CLI builder, and prose is not behaviour, so
    # the docstring is cut out rather than filtered line by line.
    body = text[text.index("def _fresh_cases"):text.index("def _previous_schema_document")]
    body = body[:body.index("def _corpus") + body[body.index("def _corpus"):].index("\n\n\n")]
    opening = body.index('"""')
    closing = body.index('"""', opening + 3) + 3
    code = body[:opening] + body[closing:]
    assert "emit_calc_artifacts" in code, "the slice did not capture the helper body"
    assert "REPO_CASES" not in code and "REPO_BUILD" not in code, (
        "the corpus helper still refers to the repository build directory"
    )
    assert "subprocess" not in code and "build_stage_a" not in code, (
        "the helper shells out to the CLI builder instead of emitting in process"
    )
    assert "is_file()" not in code.replace("artifacts.cases_path.is_file()", ""), (
        "the helper still branches on whether some other file happens to exist"
    )
    # And there is no second implementation of the corpus. Naming the emitter's
    # generators to assert they still exist is fine; DEFINING one here is not.
    #
    # The tokens are ASSEMBLED rather than written out, because a literal list of
    # them inside this file would match itself - the same reason test_39 in the
    # Gate-B suite assembles its forbidden list.
    define = "def " + "_parity_"
    for suffix in ("vectors", "neighbour_triples", "bit_stream"):
        assert define + suffix not in text, (
            f"the suite reimplements corpus generation ({define + suffix})"
        )
    constant = "_PARITY" + "_LCG_START"
    assert constant not in text, "the suite restates the emitter's generator constants"

    # BEHAVIOUR: emitting a fresh corpus leaves the repository build untouched.
    before = sorted(p.name for p in REPO_BUILD.iterdir()) if REPO_BUILD.is_dir() else None
    stamp = REPO_CASES.stat().st_mtime_ns if REPO_CASES.is_file() else None
    _FRESH.clear()
    parity = _corpus()
    assert len(parity["vectors"]) > 2000
    after = sorted(p.name for p in REPO_BUILD.iterdir()) if REPO_BUILD.is_dir() else None
    assert before == after, "the suite created or removed a repository build artifact"
    if stamp is not None:
        assert REPO_CASES.stat().st_mtime_ns == stamp, (
            "the suite rewrote the repository corpus"
        )
    # The temporary directory really is outside the repository.
    assert PCCM_ROOT not in _FRESH["directory"].parents, _FRESH["directory"]


def test_19_the_expectations_still_come_from_the_fingerprint_oracle() -> None:
    """Isolation must not have turned the port, or the corpus, into an authority."""
    generator = (PCCM_ROOT / "builder" / "pccm_builder" / "calc_cases.py").read_text(
        encoding="utf-8")
    assert 'fp.canonical_number(value, ".")' in generator
    parity = _corpus()
    # Every emitted expectation is reproducible from the oracle alone.
    for vector in parity["vectors"][::97]:
        value = _from_bits(vector["bits"])
        assert vector["expected"] == fp.canonical_number(value, "."), vector["label"]
    suite = Path(__file__).read_text(encoding="utf-8")
    assert "import vba_canonical_port as port" in suite
    # The port is compared AGAINST the oracle; it never supplies an expectation.
    assert "port.canonical_number" in suite and "fp.canonical_number" in suite
    ported = (Path(__file__).resolve().parent / "vba_canonical_port.py").read_text(
        encoding="utf-8")
    # The port must not IMPORT the oracle. Its docstring names it, to say that it
    # is not one, and prose is not behaviour.
    for forbidden in ("import calc_fingerprint", "from pccm_builder", "import pccm_builder"):
        assert forbidden not in ported, (
            f"the port reaches the oracle it is checked against ({forbidden})"
        )
    assert ported.count("import ") == 1 and "import math" in ported, (
        "the port gained a dependency"
    )


def test_20_the_isolation_fix_touches_no_production_file() -> None:
    """A test-plumbing correction has no business in shipped source."""
    raw = (SRC_VBA / "modCalcFingerprint.bas").read_text(encoding="utf-8")
    # EXECUTABLE lines only. The module quotes Format$ in a comment, to record
    # what it replaced and why, and prose is not behaviour.
    module = "\n".join(line for line in raw.splitlines()
                       if line.strip() and not line.strip().startswith("'"))
    # The corrected encoder is exactly as accepted: still generated, never formatted.
    assert "Format$" not in module and "FP_NUMBER_FORMAT" not in module
    assert "Private Function CalcFpBuildCanonical(" in module
    assert "Private Const FP_LIMB_BASE As Double = 10000000#" in module
    assert "CalcFpIsOddDigit(lastDigit)" in module
    # The emitter and the oracle are untouched by anything in this suite.
    generator = (PCCM_ROOT / "builder" / "pccm_builder" / "calc_cases.py").read_text(
        encoding="utf-8")
    assert "_parity_vectors()" in generator and "_parity_neighbour_triples()" in generator
    oracle = (PCCM_ROOT / "builder" / "pccm_builder" / "calc_fingerprint.py").read_text(
        encoding="utf-8")
    assert "_SIGNIFICAND_DIGITS_AFTER_POINT = 16" in oracle
    # And the P5-DP scenario is unchanged plumbing-wise: it reads BuildDir, which
    # the harness is GIVEN, not a path this suite decides.
    harness = (PCCM_ROOT / "bootstrap" / "windows" / "phase5_gate_b_scenarios.ps1").read_text(
        encoding="utf-8")
    assert "$parity = $Cases.fingerprint.canonical_parity" in harness
    assert "Join-Path $BuildDir 'phase5_cases.json'" in harness


# ===========================================================================
# 7. THE RUN-3 COMPILE-SAFE RENAMES CHANGED NO BEHAVIOUR
# ===========================================================================
def test_21_the_encoder_is_identical_modulo_the_compile_safe_renames() -> None:
    """`scale` -> `decimalScale` and `base` -> `powerBase`, and nothing else.

    Runtime Run 3's VBE refused `Dim scale As Long`. The repair is an identifier
    rename, and an identifier rename must be provably invisible: mapping the two
    new names back must reproduce the accepted executable text exactly.
    """
    import subprocess

    accepted = subprocess.run(
        ["git", "show", "2670ae8:pccm/src/vba/modCalcFingerprint.bas"],
        capture_output=True, text=True, cwd=str(PCCM_ROOT.parent),
    )
    if accepted.returncode != 0:
        # A reconstructed tree has no repository; the rest of this file already
        # proves the behaviour directly against the oracle.
        return

    def normalise(text: str) -> str:
        kept = []
        for line in text.splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("'"):
                continue
            stripped = re.sub(r"\s+", " ", stripped)
            stripped = re.sub(r"\bdecimalScale\b", "scale", stripped)
            stripped = re.sub(r"\bpowerBase\b", "base", stripped)
            kept.append(stripped)
        return "\n".join(kept)

    current = (SRC_VBA / "modCalcFingerprint.bas").read_text(encoding="utf-8")
    assert normalise(current) == normalise(accepted.stdout), (
        "the canonical encoder changed by more than the compile-safe renames"
    )
    # And the renames really did happen.
    assert "decimalScale" in current and "powerBase" in current
    executable = "\n".join(line for line in current.splitlines()
                           if line.strip() and not line.strip().startswith("'"))
    assert not re.search(r"\bscale\b", executable), "`scale` survives in executable text"
    assert not re.search(r"(?<![A-Za-z])base(?![A-Za-z])", executable), (
        "`base` survives in executable text"
    )


def test_22_the_parity_corpus_and_every_expectation_are_unchanged() -> None:
    """A compile fix must not move one canonical digit."""
    parity = _corpus()
    assert len(parity["vectors"]) > 2000
    # Every expectation still comes from the oracle, unchanged.
    for vector in parity["vectors"][::53]:
        value = _from_bits(vector["bits"])
        assert vector["expected"] == fp.canonical_number(value, "."), vector["label"]
        assert port.canonical_number(value, ".") == (True, vector["expected"])
    # The locked ten and the reference digest are untouched.
    document = _fresh_cases()["fingerprint"]
    assert document["reference"]["digest"] == "50B6EB0E26857EA7"
    assert document["reference"]["code_units"] == 366
    assert len(document["numeric_encodings"]["vectors"]) == 10
    for vector in document["numeric_encodings"]["vectors"]:
        assert port.canonical_number(vector["value"], ".") == (True, vector["expected"])


def test_23_max_double_is_built_and_is_the_true_maximum() -> None:
    """The second Run-3 blocker, and the semantics it had to preserve."""
    module = (SRC_VBA / "modCalcFactors.bas").read_text(encoding="utf-8")
    executable = "\n".join(line for line in module.splitlines()
                           if line.strip() and not line.strip().startswith("'"))
    assert "Public Const MAX_DOUBLE" not in executable, "MAX_DOUBLE is a Const again"
    assert "Public Function MAX_DOUBLE() As Double" in executable
    assert "1.7976931348623157E+308" not in executable, "the overflowing literal is back"
    assert "result = MAX_SIGNIFICAND" in executable
    assert "For doubling = 1 To MAX_EXPONENT" in executable

    # The construction, evaluated the way the VBA evaluates it.
    built = float(2 ** 53 - 1)
    for _ in range(971):
        built = built * 2.0
    assert built == sys.float_info.max, "the construction is not DBL_MAX"

    # SEMANTICS: IsUsableDouble compares against this bound, so the largest
    # representable Double must be USABLE and must encode as the locked vector.
    assert not (built > sys.float_info.max), "DBL_MAX would be refused as unusable"
    ok, text = port.canonical_number(built, ".")
    assert ok and text == "1.7976931348623157E+308"
    assert text == fp.canonical_number(sys.float_info.max)
    # A rounded-DOWN literal would have refused it, which is why none was used.
    rounded_down = float(f"{sys.float_info.max:.14E}".replace("2E+308", "1E+308"))
    assert rounded_down < sys.float_info.max
    assert sys.float_info.max > rounded_down, (
        "a rounded-down bound would refuse the largest representable Double"
    )

    # One coherent definition across the project.
    from pccm_builder import calc_numeric
    assert calc_numeric.MAX_DOUBLE == sys.float_info.max == built
    assert "MAX_DOUBLE = (2^53 - 1) * 2^971" in module, (
        "the module no longer states the identity it builds from"
    )
