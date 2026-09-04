#!/usr/bin/env python3
"""PCCM Phase 6 Step-10 conformance tests for `src/vba/modSimFingerprint.bas`
and the one canonical continuation helper added to `modCalcFingerprint.bas`.

--------------------------------------------------------------------------------
WHAT THESE TESTS PROVE, AND WHAT THEY DO NOT
--------------------------------------------------------------------------------
SOURCE CONFORMANCE, on Linux, now: purity, the public surface, the exact SIM
suffix bytes, the request fingerprint as a CONTINUATION of the analytical hash
state, the result-digest framing and order, and the arithmetic those statements
describe - through the accepted Phase-6 source transcriber, against the accepted
Step-10A authority and the accepted `phase6_cases.json` corpus.

VBA EXECUTION CONFORMANCE is NOT proved and is deferred to Gate B on Windows.
No VBA runtime exists in this step. Nothing here may be read as "VBA produced
this fingerprint".

TWO PROCEDURES ARE BORROWED, NOT TRANSCRIBED, each for a stated reason and each
with its REAL VBA SIGNATURE read out of the source:

  CalcFpCanonicalNumber   its second tier is the exact-integer limb machinery
                          rebuilt at Gate B Runtime Run 2 - dynamic limb arrays
                          and scoped rounding the transcriber does not model.
                          Its accepted PYTHON counterpart is bound instead.
  SimFpRetainedExtent     reads a bound of an unproven carrier under a SCOPED
                          error handler, and the engine models no `On Error`.
                          The shim reproduces the ALLOCATED arm; the arm that
                          RAISES is Gate-B work.

EVERYTHING ELSE IS COMPILED FROM THE REAL SOURCE, the whole hash core included:
the reduction, the code-unit normalisation, the digest loop, the hex conversion
and the new continuation.

Runs standalone or under pytest.
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

from pccm_builder import (  # noqa: E402
    load_contract,
    load_sim_contract,
    load_structure_contract,
)
from pccm_builder import calc_fingerprint as fp  # noqa: E402
from pccm_builder.calc_cases import reference_stream  # noqa: E402
from pccm_builder.sim_emit import render_sim_contract_module  # noqa: E402
from pccm_builder.spec_loader import load_spec  # noqa: E402
from pccm_builder.vba_source import VbaModule, load_modules, strip_comments  # noqa: E402

from phase6_vba_transcribe import _Ref, _val, build as _build_transcription  # noqa: E402

SRC_VBA = PCCM_ROOT / "src" / "vba"
SIM_FP_BAS = SRC_VBA / "modSimFingerprint.bas"
CALC_FP_BAS = SRC_VBA / "modCalcFingerprint.bas"
CALC_FACTORS_BAS = SRC_VBA / "modCalcFactors.bas"
CALC_CONTRACT_BAS = PCCM_ROOT / "build" / "vba" / "modCalcContract.bas"
SPEC = PCCM_ROOT / "spec"
CASES_JSON = PCCM_ROOT / "build" / "phase6_cases.json"

ANALYTICAL_DIGEST = "50B6EB0E26857EA7"
"""The accepted Phase-5 case-26 fingerprint. Step 10 does not recompute it."""

# The four Step-10A golden request vectors, as literals.
REQUEST_VECTORS = {
    "auto.1000": ((1000, "AUTO", False, 0),
                  "S3:SIMI1:1I1:4I4:1000S4:AUTOI1:1I1:1", "5EAB16E15C2ECE24"),
    "fixed.seed_1": ((1000, "FIXED", True, 1),
                     "S3:SIMI1:1I1:5I4:1000S5:FIXEDI1:1I1:1I1:1", "599C95E7274759B9"),
    "fixed.seed_max": ((1000, "FIXED", True, 2147483646),
                       "S3:SIMI1:1I1:5I4:1000S5:FIXEDI10:2147483646I1:1I1:1",
                       "0010FB954CC94B53"),
    "auto.1001": ((1001, "AUTO", False, 0),
                  "S3:SIMI1:1I1:4I4:1001S4:AUTOI1:1I1:1", "4777C8BC35F0FFEF"),
}

# The seven accepted E_digest vectors, as literals.
DIGEST_VECTORS = {
    "digest.base": "3181AF89642DE500",
    "digest.reversed_iteration_order": "4E0FEE211853E8F6",
    "digest.nominal_and_pv_swapped": "63A0E93074F0C2EA",
    "digest.one_iteration_dropped": "0CAC531732B88B2A",
    "digest.one_ulp_perturbation": "5DC1A76B56D75EF4",
    "digest.version_2": "7E8D58C46CCDD798",
    "digest.empty": "12ED977808313D71",
}

# The ONE procedure of each module the transcriber cannot execute.
BORROWED = {"modCalcFingerprint": {"CalcFpCanonicalNumber"},
            "modSimFingerprint": {"SimFpRetainedExtent"}}

# The accepted hash core, compiled from source. Nothing outside this set is
# needed, and the request/digest framing may call nothing else.
CALC_FP_COMPILED = {
    "CalcFpUtf16Length", "CalcFpNormaliseCodeUnit", "CalcFpCanonicalText",
    "CalcFpCanonicalInteger", "CalcFpNumberField", "CalcFpField", "CalcFpDigitsOf",
    "CalcFpReduceDouble", "CalcFpDigestStream", "CalcFpHex8", "CalcFpPowerOf16",
    "CalcFpContinueDigest", "CalcFpHexValue", "CalcFpHexDigitValue",
}

_CACHE: dict[str, object] = {}


# ---------------------------------------------------------------------------
# Source access
# ---------------------------------------------------------------------------
def _module(name: str = "modSimFingerprint") -> VbaModule:
    path = SIM_FP_BAS if name == "modSimFingerprint" else SRC_VBA / f"{name}.bas"
    return VbaModule(name=name, path=path, raw=path.read_text(encoding="utf-8"))


def _code(name: str = "modSimFingerprint") -> str:
    return _module(name).code


def _procedure(name: str, module: str = "modSimFingerprint") -> str:
    code = _module(module).code_without_string_removal
    match = re.search(
        rf"^\s*(?:Public|Private)\s+(?:Function|Sub)\s+{re.escape(name)}\b", code, re.M)
    assert match, f"{name} is not declared in {module}"
    tail = code[match.start():]
    end = re.search(r"^\s*End\s+(?:Function|Sub)\s*$", tail, re.M)
    assert end, f"{name} has no End"
    return tail[: end.end()]


def _constants() -> dict:
    if "consts" not in _CACHE:
        out: dict = {}
        rendered = render_sim_contract_module(
            load_spec(SPEC / "workbook.yaml"),
            load_sim_contract(SPEC / "sim_contract.yaml"),
            load_contract(SPEC / "input_contract.yaml"))
        for text in (rendered, CALC_CONTRACT_BAS.read_text(encoding="utf-8"),
                     CALC_FP_BAS.read_text(encoding="utf-8"),
                     CALC_FACTORS_BAS.read_text(encoding="utf-8")):
            for line in text.splitlines():
                match = re.match(r"^(?:Public|Private) Const (\w+) As (\w+) = (.*)$", line)
                if not match:
                    continue
                name, kind, rest = match.groups()
                literal = rest.split("    '")[0].rstrip()
                out[name] = (literal[1:-1] if kind == "String"
                             else (float(literal.rstrip("#")) if kind == "Double"
                                   else int(literal)))
        _CACHE["consts"] = out
    return _CACHE["consts"]  # type: ignore[return-value]


def _const(name: str):
    return _constants()[name]


def _canonical_number_shim(value, decimal_separator, result):
    """The accepted Python canonical encoder, bound to its real VBA signature."""
    try:
        result.v = fp.canonical_number(float(_val(value)), _val(decimal_separator))
    except fp.FingerprintError:
        return False
    return True


def _retained_extent_shim(total_nominal, total_pv, nominal_extent, pv_extent):
    """The ALLOCATED arm of `SimFpRetainedExtent`, and only that arm."""
    nominal_extent.v = len(total_nominal)
    pv_extent.v = len(total_pv)
    return True


def _transcribe() -> dict:
    if "vba" not in _CACHE:
        simfp = set(_module().procedures) - BORROWED["modSimFingerprint"]
        _CACHE["vba"] = _build_transcription(
            {"modSimFingerprint": SIM_FP_BAS,
             "modCalcFingerprint": CALC_FP_BAS,
             "modCalcFactors": CALC_FACTORS_BAS},
            _constants(),
            only={"modCalcFactors": {"IsUsableDouble"},
                  "modCalcFingerprint": CALC_FP_COMPILED,
                  "modSimFingerprint": simfp},
            signature_only={"modCalcFingerprint": set(BORROWED["modCalcFingerprint"]),
                            "modSimFingerprint": set(BORROWED["modSimFingerprint"])},
            extra={
                "MAX_DOUBLE": sys.float_info.max,
                "CalcFpCanonicalNumber": _canonical_number_shim,
                "SimFpRetainedExtent": _retained_extent_shim,
            })
    return _CACHE["vba"]  # type: ignore[return-value]


# ---------------------------------------------------------------------------
# Calling conveniences
# ---------------------------------------------------------------------------
def _request(prior, iterations, seed_mode, has_seed, seed=0, result=None):
    result = _Ref("") if result is None else result
    detail = _Ref("")
    ok = _transcribe()["SimFpBuildRequestFingerprint"](
        _Ref(prior), _Ref(iterations), _Ref(seed_mode), _Ref(has_seed), _Ref(seed),
        result, detail)
    return ok, result.v, detail.v


def _suffix(iterations, seed_mode, has_seed, seed=0):
    out, detail = _Ref(""), _Ref("")
    ok = _transcribe()["SimFpRequestSuffix"](
        _Ref(iterations), _Ref(seed_mode), _Ref(has_seed), _Ref(seed), out, detail)
    return ok, out.v, detail.v


def _digest(nominal, pv, count=None, separator=".", result=None):
    result = _Ref("") if result is None else result
    detail = _Ref("")
    ok = _transcribe()["SimFpResultDigest"](
        list(nominal), list(pv), _Ref(len(nominal) if count is None else count),
        _Ref(separator), result, detail)
    return ok, result.v, detail.v


def _versioned_digest(version, nominal, pv, count=None, separator="."):
    result, detail = _Ref(""), _Ref("")
    ok = _transcribe()["SimFpVersionedResultDigest"](
        _Ref(version), list(nominal), list(pv),
        _Ref(len(nominal) if count is None else count), _Ref(separator), result, detail)
    return ok, result.v, detail.v


def _continue(prior, suffix):
    result = _Ref("")
    ok = _transcribe()["CalcFpContinueDigest"](_Ref(prior), _Ref(suffix), result)
    return ok, result.v


def _digest_stream(stream):
    result = _Ref("")
    ok = _transcribe()["CalcFpDigestStream"](_Ref(stream), result)
    return ok, result.v


def _cases() -> dict[str, dict]:
    if "cases" not in _CACHE:
        corpus = json.loads(CASES_JSON.read_text(encoding="utf-8"))
        _CACHE["cases"] = {case["id"]: case
                           for group in corpus["groups"] for case in group["cases"]}
    return _CACHE["cases"]  # type: ignore[return-value]


# ===========================================================================
# A. Declaration, registry, surface and purity
# ===========================================================================
def test_01_the_module_exists_and_is_explicit() -> None:
    lines = SIM_FP_BAS.read_text(encoding="utf-8").splitlines()
    assert lines[0] == 'Attribute VB_Name = "modSimFingerprint"'
    assert lines[1] == "Option Explicit"


def test_02_the_module_is_registered_and_nothing_beyond_it() -> None:
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    modules = {m.name: m for m in structure.vba_modules}
    assert "modSimFingerprint" in modules
    assert modules["modSimFingerprint"].generated is False
    # THE PHASE-6 BLOCK, CONTIGUOUS AND IN ORDER. This was written as the LAST
    # eight entries, which was the same claim while Phase 6 was the last phase.
    # "Nothing beyond it" has since been settled by P7-2 landing
    # modSimSensitivity under its own authority; what still matters, and is
    # still checked, is that the accepted block is intact and unreordered.
    names = [m.name for m in structure.vba_modules]
    block = ['modSimContract', 'modSimRng', 'modSimSample', 'modSimEngine', 'modSimStats', 'modSimFingerprint', 'modSimNonce', 'modSimReport']
    at = names.index(block[0])
    assert names[at:at + len(block)] == block, names[at:at + len(block)]
    # No endpoint, and D6-11 is untouched.
    surface = set(structure.entry_points) | set(structure.api_procedures)
    assert not (set(_module().public_procedures) & surface)
    scoped = [(r.construct, tuple(r.allowed_in))
              for r in structure.forbidden_construct_rules if r.is_scoped]
    assert scoped == [("MRG32k3a", ("modSimRng",)),
                      ("RunSimulation", ("modSimReport",))], scoped
    for construct in ("Percentile",):
        rule = next(r for r in structure.forbidden_construct_rules
                    if r.construct == construct)
        assert not rule.is_scoped, construct
        assert rule.forbidden_in("modSimFingerprint") is True, construct


def test_03_the_public_surface_is_the_two_framing_operations() -> None:
    assert sorted(_module().public_procedures) == [
        "SimFpBuildRequestFingerprint", "SimFpResultDigest"]
    private = set(_module().procedures) - set(_module().public_procedures)
    assert private == {
        "SimFpDigestRecord", "SimFpRequestSuffix", "SimFpRetainedExtent",
        "SimFpValidateRequest", "SimFpVersionedResultDigest",
    }, sorted(private)
    assert not re.findall(r"^(?:Public|Private) Type (\w+)$",
                          _module().raw, re.M), "the module declares a Type"


def test_04_the_production_digest_surface_takes_no_version() -> None:
    """A caller who could choose a method version could produce a digest that
    claims a method the run did not use."""
    signature = _procedure("SimFpResultDigest").split(") As Boolean")[0]
    assert "methodVersion" not in signature, signature
    assert "version" not in signature.lower(), signature
    body = _procedure("SimFpResultDigest")
    assert "SIM_METHOD_VERSION" in body
    # The versioned helper exists and is PRIVATE.
    assert "SimFpVersionedResultDigest" not in _module().public_procedures
    assert "Private Function SimFpVersionedResultDigest" in _module().raw


def test_05_the_module_never_reaches_a_workbook_or_the_environment() -> None:
    code = _code()
    for banned in ("Range", "Cells", "Worksheet", "Workbook", "ListObject",
                   "Application", "ThisWorkbook", "ActiveWorkbook", "Names",
                   "Evaluate", "MsgBox", "Open ", "Close ", "Print #", "Date",
                   "Now", "Timer", "Environ", "Shell", "CreateObject", "GetObject"):
        assert banned not in code, banned
    for line in _module().raw.splitlines():
        assert not re.search(r"\bAs\s+Object\b", line), line
        assert not re.search(r"\bAs\s+Variant\b", line), line


def test_06_there_is_no_module_level_or_static_state() -> None:
    body = strip_comments(_module().raw)
    header = body.split("Public Function SimFpBuildRequestFingerprint")[0]
    for line in header.splitlines():
        assert not re.match(r"^\s*(Dim|Public|Private)\s+\w+\s+As\s", line), line
    assert "Static " not in body


def test_07_the_globally_forbidden_words_never_appear() -> None:
    code = _code()
    for banned in ("Percentile", "MRG32k3a", "RunSimulation", "Rnd(", "Randomize"):
        assert banned not in code, banned


def test_08_no_simulation_statistics_or_publication_leaks_in() -> None:
    # CASE-INSENSITIVE: the projected constants are spelled SIM_QUANTILE_*, and a
    # case-sensitive scan for "Quantile" would walk straight past one.
    code = _code().lower()
    for banned in ("simenginerun", "simrng", "simsample", "simstats", "knom", "kpv",
                   "quantity", "probability", "quantile", "contingency", "simdata",
                   "results", "runid", "run_id", "attemptresult", "pccm_"):
        assert banned not in code, banned


def test_09_it_owns_no_encoding_hash_or_hex_of_its_own() -> None:
    """modCalcFingerprint stays the ONLY owner of the canonical mathematics."""
    code = _code()
    for owned_elsewhere in ("FP_BASE", "FP_MOD_1", "FP_MOD_2", "FP_INIT_1", "FP_INIT_2",
                            "FP_HEX_DIGITS", "FP_HEX_WIDTH", "AscW", "Hex", "Format",
                            "0123456789ABCDEF", "131", "2147483647", "2147483629"):
        assert owned_elsewhere not in code, owned_elsewhere
    # No hand-assembled field, no hand-rolled canonical number. STRINGS KEPT:
    # `_code()` strips them, and a hand-rolled encoder is made of string literals.
    literals = _module().code_without_string_removal
    # OPEN-QUOTE forms: a hand-rolled exponent is written "E+00", not "E+".
    for hand_rolled in ('"S" &', '"I" &', '"N" &', '":" &', '"E+', '"E-', '"E"'):
        assert hand_rolled not in literals, hand_rolled
    # Every encoded field comes from the accepted framing authority.
    used = set(re.findall(r"modCalcFingerprint\.(\w+)", code))
    assert used == {"CalcFpCanonicalText", "CalcFpCanonicalInteger",
                    "CalcFpNumberField", "CalcFpDigestStream",
                    "CalcFpContinueDigest"}, sorted(used)


def test_10_there_is_no_second_hash_recurrence_here() -> None:
    code = _code()
    for recurrence in ("Fix(", " Mod ", "* 16#", "/ 16#", "65536", "32767"):
        assert recurrence not in code, recurrence
    assert "CalcFpReduceDouble" not in code
    assert "CalcFpNormaliseCodeUnit" not in code


def test_11_the_only_phase5_primitive_borrowed_directly_is_the_domain_predicate() -> None:
    code = _code()
    calls = set(re.findall(r"\b(IsUsableDouble|Safe\w+|Exact\w+)\b", code))
    assert calls == {"IsUsableDouble"}, sorted(calls)


def test_12_the_continuation_helper_is_public_and_owned_by_the_encoder() -> None:
    calc = _module("modCalcFingerprint")
    assert "CalcFpContinueDigest" in calc.public_procedures
    for private in ("CalcFpHexValue", "CalcFpHexDigitValue"):
        assert private in calc.procedures
        assert private not in calc.public_procedures, private
    # And it is the ONLY public procedure Step 10 added.
    added = set(calc.public_procedures) - {
        "CalcFpUtf16Length", "CalcFpNormaliseCodeUnit", "CalcFpCanonicalText",
        "CalcFpCanonicalNumber", "CalcFpCanonicalInteger", "CalcFpReduceDouble",
        "CalcFpDigestStream", "CalcFpBuildCostRecord", "CalcFpBuildRiskRecord",
        "CalcFpBuildFingerprint", "CalcFpNumberField"}
    assert added == {"CalcFpContinueDigest"}, sorted(added)


# ===========================================================================
# B. The canonical continuation primitive
# ===========================================================================
def test_13_continuing_from_a_digest_equals_digesting_the_joined_stream() -> None:
    """THE PROPERTY THE WHOLE DESIGN RESTS ON. The digest IS the accumulator
    pair, so there is no finalisation transform to undo."""
    prefixes = ["PCCM-FP", "S7:PCCM-FPI1:1", reference_stream(1),
                "A", "高", "\U0001F600ab"]
    suffixes = ["", "x", "S3:SIMI1:1", "高é", "\U0001F600",
                "\U00010000\U0010FFFF", "S4:AUTOI1:1I1:1"]
    for prefix in prefixes:
        ok, prior = _digest_stream(prefix)
        assert ok, prefix[:20]
        for suffix in suffixes:
            ok, continued = _continue(prior, suffix)
            assert ok, (prefix[:20], suffix)
            ok, whole = _digest_stream(prefix + suffix)
            assert ok
            assert continued == whole, (prefix[:20], repr(suffix))


def test_14_the_empty_suffix_returns_the_prior_digest_unchanged() -> None:
    ok, prior = _digest_stream(reference_stream(1))
    assert ok and prior == ANALYTICAL_DIGEST
    ok, continued = _continue(prior, "")
    assert ok and continued == prior


def test_15_non_bmp_text_is_consumed_as_code_units_not_code_points() -> None:
    """A surrogate pair contributes TWO units, exactly as AscW would see it."""
    text = "\U0001F600"
    assert len(text) == 1 and fp.utf16_length(text) == 2
    ok, prior = _digest_stream("PCCM-FP")
    assert ok
    ok, continued = _continue(prior, text)
    assert ok
    ok, whole = _digest_stream("PCCM-FP" + text)
    assert ok and continued == whole
    # And it differs from consuming the code point once.
    ok, single = _continue(prior, "�")
    assert ok and single != continued


def test_16_the_four_step10a_suffixes_continue_the_accepted_prefix() -> None:
    prefix = reference_stream(1)
    ok, prior = _digest_stream(prefix)
    assert ok and prior == ANALYTICAL_DIGEST
    for key, (_args, suffix, expected) in REQUEST_VECTORS.items():
        ok, continued = _continue(prior, suffix)
        assert ok, key
        assert continued == expected, key
        ok, whole = _digest_stream(prefix + suffix)
        assert ok and whole == expected, key


def test_17_a_digest_that_is_not_the_canonical_form_is_refused() -> None:
    for bad in ("", "50B6EB0E26857EA", "50B6EB0E26857EA77", "50b6eb0e26857ea7",
                "50B6EB0E26857ea7", " 50B6EB0E26857EA7", "50B6EB0E26857EA7 ",
                "+50B6EB0E26857EA", "0x50B6EB0E26857EA", "50B6EB0E26857EAG",
                "50B6EB0E-6857EA7", "50B6EB0E 6857EA7"):
        ok, _value = _continue(bad, "x")
        assert ok is False, repr(bad)


def test_18_a_state_at_or_above_its_own_modulus_is_refused() -> None:
    """h1 and h2 have DIFFERENT moduli, and each must be a residue of its own."""
    m1, m2 = _const("FP_MOD_1"), _const("FP_MOD_2")
    assert m1 != m2
    good = "%08X%08X" % (m1 - 1, m2 - 1)
    ok, _value = _continue(good, "x")
    assert ok, good
    for h1, h2 in ((m1, m2 - 1), (m1 - 1, m2), (m1, m2), (0xFFFFFFFF, 0)):
        ok, _value = _continue("%08X%08X" % (h1, h2), "x")
        assert ok is False, (h1, h2)
    # The value between the two moduli is legal for h1 and illegal for h2.
    assert m2 < m1
    ok, _value = _continue("%08X%08X" % (m2, 0), "x")
    assert ok, "a residue below FP_MOD_1 must be accepted as h1"
    ok, _value = _continue("%08X%08X" % (0, m2), "x")
    assert ok is False, "a value at FP_MOD_2 must be refused as h2"


def test_19_the_two_halves_are_not_interchangeable() -> None:
    ok, prior = _digest_stream("PCCM-FP")
    assert ok
    swapped = prior[8:] + prior[:8]
    ok, straight = _continue(prior, "x")
    assert ok
    ok, crossed = _continue(swapped, "x")
    if ok:
        assert crossed != straight, "swapping h1 and h2 changed nothing"


def test_20_the_continuation_reuses_the_accepted_primitives_and_adds_none() -> None:
    body = _procedure("CalcFpContinueDigest", "modCalcFingerprint")
    for reused in ("CalcFpUtf16Length", "CalcFpNormaliseCodeUnit",
                   "CalcFpReduceDouble", "CalcFpHex8", "FP_MOD_1", "FP_MOD_2"):
        assert reused in body, reused
    # No re-initialisation: continuing from FP_INIT_1 would ignore the prior.
    # Precise call forms: "Hex" is a substring of CalcFpHexValue and
    # FP_HEX_WIDTH, both of which this body is SUPPOSED to name.
    for forbidden in ("FP_INIT_1", "FP_INIT_2", "FP_BASE", "Hex$(", "Hex(",
                      "Format$(", "Format(", "Val(", "&H"):
        assert forbidden not in body, forbidden
    # The reduction is the accepted one, called - not re-derived.
    assert "* FP_BASE" not in body


def test_21_the_hex_decoder_is_ordinal_and_table_driven() -> None:
    body = _procedure("CalcFpHexDigitValue", "modCalcFingerprint")
    assert "FP_HEX_DIGITS" in body
    assert "vbBinaryCompare" in body, "a case-insensitive match would accept lowercase"
    for host in ("CLng(", "&H", "Hex$(", "Hex(", "Val(", "Asc(", "InStr("):
        assert host not in body, host
    value = _procedure("CalcFpHexValue", "modCalcFingerprint")
    assert "As Double" in value, "the accumulator must not narrow to a Long"
    assert "CLng" not in value


def test_22_the_accepted_digest_loop_was_not_touched() -> None:
    """Continuation is an ADDITION. `CalcFpDigestStream` still starts at the
    locked initial states and is byte-identical to its accepted body."""
    body = _procedure("CalcFpDigestStream", "modCalcFingerprint")
    assert "h1 = FP_INIT_1" in body and "h2 = FP_INIT_2" in body
    ok, value = _digest_stream("PCCM-FP")
    assert ok and value == "6551C6F365DA7F3F"


# ===========================================================================
# C. The request fingerprint
# ===========================================================================
def test_23_every_golden_request_vector_reproduces_exactly() -> None:
    for key, (args, suffix, expected) in REQUEST_VECTORS.items():
        ok, built, detail = _suffix(*args)
        assert ok, f"{key}: {detail}"
        assert built == suffix, key
        ok, value, detail = _request(ANALYTICAL_DIGEST, *args)
        assert ok, f"{key}: {detail}"
        assert value == expected, key


def test_24_the_corpus_suffixes_are_the_ones_the_source_builds() -> None:
    for key, (args, suffix, expected) in REQUEST_VECTORS.items():
        case = _cases()[f"request_fingerprint.{key}"]
        assert case["comparison"] == "EXACT"
        assert case["expected_exact"]["sim_suffix"] == suffix, key
        assert case["expected_exact"]["request_fingerprint"] == expected, key
        assert case["inputs"]["analytical_prefix_digest"] == ANALYTICAL_DIGEST
        ok, built, detail = _suffix(*args)
        assert ok and built == case["expected_exact"]["sim_suffix"], f"{key}: {detail}"


def test_25_auto_hashes_four_fields_and_fixed_hashes_five() -> None:
    ok, auto, _d = _suffix(1000, "AUTO", False)
    ok2, fixed, _d = _suffix(1000, "FIXED", True, 1)
    assert ok and ok2
    section = _const("SIM_REQUEST_SECTION")
    assert auto.startswith(f"S{len(section)}:{section}I1:1I1:4")
    assert fixed.startswith(f"S{len(section)}:{section}I1:1I1:5")
    assert _const("SIM_REQUEST_FIELD_COUNT_AUTO") == 4
    assert _const("SIM_REQUEST_FIELD_COUNT_FIXED") == 5
    assert _const("SIM_REQUEST_RECORD_COUNT") == 1


def test_26_the_auto_record_carries_no_seed_field_of_any_kind() -> None:
    ok, auto, _d = _suffix(1000, "AUTO", False, 12345)
    assert ok
    assert auto == "S3:SIMI1:1I1:4I4:1000S4:AUTOI1:1I1:1"
    # The ignored argument really is ignored, and no sentinel replaced it.
    for seed in (0, 1, 12345, 2147483646):
        ok, again, _d = _suffix(1000, "AUTO", False, seed)
        assert ok and again == auto, seed
    for sentinel in ("I1:0", "S0:"):
        assert sentinel not in auto, sentinel


def test_27_the_request_never_encodes_the_analytical_fingerprint() -> None:
    """It is prior HASH STATE. Encoding it would be a different stream."""
    body = _procedure("SimFpBuildRequestFingerprint")
    assert "CalcFpContinueDigest(analyticalFingerprint, suffix, candidate)" in body
    for encoded in ("CalcFpCanonicalText(analyticalFingerprint)",
                    "CalcFpCanonicalInteger(analyticalFingerprint",
                    "CalcFpNumberField(analyticalFingerprint",
                    '"S16:"'):
        assert encoded not in body, encoded
    suffix_body = _procedure("SimFpRequestSuffix")
    assert "analyticalFingerprint" not in suffix_body, (
        "the suffix builder cannot see the analytical fingerprint at all")
    # Behaviourally: hashing it as a field would be a DIFFERENT answer.
    ok, suffix, _d = _suffix(1000, "AUTO", False)
    assert ok
    as_a_field = fp.text_field(ANALYTICAL_DIGEST).encode() + suffix
    assert fp.fingerprint(as_a_field) != REQUEST_VECTORS["auto.1000"][2]


def test_28_a_new_stream_is_not_started_at_sim() -> None:
    ok, suffix, _d = _suffix(1000, "AUTO", False)
    assert ok
    for restarted in (fp.fingerprint(suffix),
                      fp.fingerprint(fp.text_field("PCCM-FP").encode() + suffix)):
        assert restarted != REQUEST_VECTORS["auto.1000"][2]
    literals = _module().code_without_string_removal
    assert "PCCM-FP" not in literals
    assert "FP_VERSION" not in literals


def test_29_the_iteration_bounds_are_the_projected_ones() -> None:
    low, high = _const("SIM_MIN_ITERATIONS"), _const("SIM_MAX_ITERATIONS")
    for bad in (low - 1, 0, -1, high + 1):
        ok, _value, detail = _request(ANALYTICAL_DIGEST, bad, "AUTO", False)
        assert ok is False, bad
        assert detail
    for good in (low, high):
        ok, _value, detail = _request(ANALYTICAL_DIGEST, good, "AUTO", False)
        assert ok, f"{good}: {detail}"
    body = _procedure("SimFpValidateRequest")
    assert "SIM_MIN_ITERATIONS" in body and "SIM_MAX_ITERATIONS" in body
    assert not re.search(r"\b\d{4,}\b", body), "an iteration bound is spelled out"


def test_30_the_seed_mode_is_matched_ordinally() -> None:
    for bad in ("auto", "Auto", "AUTO ", " AUTO", "fixed", "RANDOM", ""):
        ok, _value, detail = _request(ANALYTICAL_DIGEST, 1000, bad, False)
        assert ok is False, repr(bad)
    assert "vbBinaryCompare" in _procedure("SimFpValidateRequest")


def test_31_the_flag_and_the_mode_must_agree() -> None:
    ok, _value, detail = _request(ANALYTICAL_DIGEST, 1000, "AUTO", True, 1)
    assert ok is False and "AUTO" in detail
    ok, _value, detail = _request(ANALYTICAL_DIGEST, 1000, "FIXED", False, 1)
    assert ok is False and "FIXED" in detail


def test_32_the_fixed_seed_domain_is_the_input_contracts() -> None:
    low, high = _const("SIM_SEED_MIN"), _const("SIM_SEED_MAX")
    assert (low, high) == (1, 2147483646)
    for bad in (low - 1, high + 1, -1):
        ok, _value, detail = _request(ANALYTICAL_DIGEST, 1000, "FIXED", True, bad)
        assert ok is False, bad
    for good in (low, high):
        ok, _value, detail = _request(ANALYTICAL_DIGEST, 1000, "FIXED", True, good)
        assert ok, f"{good}: {detail}"
    body = _procedure("SimFpValidateRequest")
    assert "SIM_SEED_MIN" in body and "SIM_SEED_MAX" in body
    assert "2147483646" not in body


def test_33_an_invalid_analytical_fingerprint_refuses_transactionally() -> None:
    sentinel = _Ref("UNTOUCHED")
    ok, _value, detail = _request("50b6eb0e26857ea7", 1000, "AUTO", False,
                                  result=sentinel)
    assert ok is False and detail
    assert sentinel.v == "UNTOUCHED", "a refusal wrote to the caller's output"


def test_34_every_request_refusal_leaves_the_output_alone() -> None:
    for args in ((999, "AUTO", False, 0), (1000, "auto", False, 0),
                 (1000, "AUTO", True, 1), (1000, "FIXED", False, 0),
                 (1000, "FIXED", True, 0), (1000, "FIXED", True, 2147483647)):
        sentinel = _Ref("UNTOUCHED")
        ok, _value, detail = _request(ANALYTICAL_DIGEST, *args, result=sentinel)
        assert ok is False, args
        assert sentinel.v == "UNTOUCHED", args
        assert detail.startswith("request fingerprint:"), detail


# ===========================================================================
# D. The result digest
# ===========================================================================
def _digest_case(identifier):
    case = _cases()[identifier]
    return (case["inputs"]["total_nominal"], case["inputs"]["total_pv"],
            case["inputs"]["sim_method_version"], case["expected_exact"])


def test_35_every_accepted_digest_vector_reproduces_exactly() -> None:
    for identifier, expected in DIGEST_VECTORS.items():
        nominal, pv, version, block = _digest_case(identifier)
        assert block["result_digest"] == expected, identifier
        if version == _const("SIM_METHOD_VERSION"):
            ok, value, detail = _digest(nominal, pv)
        else:
            ok, value, detail = _versioned_digest(version, nominal, pv)
        assert ok, f"{identifier}: {detail}"
        assert value == expected, identifier


def test_36_the_canonical_stream_matches_the_corpus_literal() -> None:
    """The small retained vectors, byte for byte - not merely the same digest."""
    for identifier in DIGEST_VECTORS:
        nominal, pv, version, block = _digest_case(identifier)
        stream = block["canonical_stream"]
        assert fp.fingerprint(stream) == block["result_digest"], identifier
        # The framing prefix the source builds is the head of that stream.
        head = (fp.text_field(_const("SIM_DIGEST_STREAM_TAG")).encode()
                + fp.integer_field(version).encode()
                + fp.text_field(_const("SIM_DIGEST_SECTION")).encode()
                + fp.integer_field(len(nominal)).encode())
        assert stream.startswith(head), identifier


def test_37_the_empty_framing_vector_reads_no_bound() -> None:
    ok, value, detail = _digest([], [], count=0)
    assert ok, detail
    assert value == DIGEST_VECTORS["digest.empty"]
    body = _procedure("SimFpVersionedResultDigest")
    assert "If sampleCount > 0 Then" in body
    guarded = body.index("If sampleCount > 0 Then")
    assert body.index("SimFpRetainedExtent") > guarded, (
        "a bound is read before the zero-count guard")


def test_38_the_iteration_index_is_logical_and_one_based() -> None:
    assert _const("SIM_DIGEST_INDEX_ORIGIN") == 1
    nominal, pv, _v, block = _digest_case("digest.base")
    stream = block["canonical_stream"]
    for position in range(1, len(nominal) + 1):
        assert fp.integer_field(position).encode() in stream, position
    assert fp.integer_field(0).encode() not in stream.split("S6:RESULT")[1][:8]
    body = _procedure("SimFpDigestRecord")
    assert "SIM_DIGEST_INDEX_ORIGIN + offset" in body
    assert "LBound" not in body, "the record must not see a physical bound"
    # ...and the CALLER hands it the LOGICAL offset. Passing the physical index
    # is invisible in the numbers whenever LBound happens to be zero, so only a
    # source detector can see it.
    caller = _procedure("SimFpVersionedResultDigest")
    assert "SimFpDigestRecord(offset, " in caller, caller


def test_39_the_physical_lbound_does_not_change_the_digest() -> None:
    """Step-8 arrays are zero-based; a one-based carrier must digest the same."""
    nominal, pv, _v, block = _digest_case("digest.base")
    ok, zero_based, detail = _digest(nominal, pv)
    assert ok, detail
    body = _procedure("SimFpVersionedResultDigest")
    assert "totalNominal(LBound(totalNominal) + offset)" in body
    assert "totalPv(LBound(totalPv) + offset)" in body
    assert zero_based == block["result_digest"]


def test_40_the_retained_order_is_used_and_never_sorted() -> None:
    base = DIGEST_VECTORS["digest.base"]
    reversed_digest = DIGEST_VECTORS["digest.reversed_iteration_order"]
    assert base != reversed_digest
    nominal, pv, _v, _b = _digest_case("digest.base")
    rnominal, rpv, _v, _b = _digest_case("digest.reversed_iteration_order")
    ok, a, _d = _digest(nominal, pv)
    ok2, b, _d = _digest(rnominal, rpv)
    assert ok and ok2 and a != b
    assert sorted(nominal) == sorted(rnominal), "the fixtures are the same multiset"
    for banned in ("Sort", "Order", "Ascending", "Descending"):
        assert banned not in _code(), banned


def test_41_the_two_measures_are_not_interchangeable() -> None:
    nominal, pv, _v, _b = _digest_case("digest.base")
    ok, straight, _d = _digest(nominal, pv)
    ok2, swapped, _d = _digest(pv, nominal)
    assert ok and ok2
    assert straight == DIGEST_VECTORS["digest.base"]
    assert swapped == DIGEST_VECTORS["digest.nominal_and_pv_swapped"]
    assert straight != swapped


def test_42_a_dropped_record_and_one_ulp_are_both_visible() -> None:
    nominal, pv, _v, _b = _digest_case("digest.base")
    ok, whole, _d = _digest(nominal, pv)
    ok2, short, _d = _digest(nominal[:-1], pv[:-1])
    assert ok and ok2
    assert short == DIGEST_VECTORS["digest.one_iteration_dropped"] != whole
    import math

    perturbed = [math.nextafter(nominal[0], math.inf)] + list(nominal[1:])
    ok3, moved, _d = _digest(perturbed, pv)
    assert ok3 and moved == DIGEST_VECTORS["digest.one_ulp_perturbation"] != whole


def test_43_the_same_arrays_always_produce_the_same_digest() -> None:
    nominal, pv, _v, _b = _digest_case("digest.base")
    first = _digest(nominal, pv)
    second = _digest(list(nominal), list(pv))
    assert first[1] == second[1] == DIGEST_VECTORS["digest.base"]


def test_44_a_non_finite_retained_total_is_refused_not_skipped() -> None:
    nominal, pv, _v, _b = _digest_case("digest.base")
    for poison in (float("nan"), float("inf"), float("-inf")):
        for position in (0, len(nominal) - 1):
            damaged = list(nominal)
            damaged[position] = poison
            sentinel = _Ref("UNTOUCHED")
            ok, _value, detail = _digest(damaged, pv, result=sentinel)
            assert ok is False, (poison, position)
            assert sentinel.v == "UNTOUCHED"
            assert "finite Double" in detail, detail
            damaged = list(pv)
            damaged[position] = poison
            ok, _value, detail = _digest(nominal, damaged)
            assert ok is False, (poison, position)


def test_45_a_malformed_carrier_refuses_without_a_subscript_error() -> None:
    nominal, pv, _v, _b = _digest_case("digest.base")
    count = len(nominal)
    for bad in ((nominal[:-1], pv, count), (nominal, pv[:-1], count),
                (nominal + [1.0], pv, count), (nominal, pv + [1.0], count),
                ([], [], count)):
        sentinel = _Ref("UNTOUCHED")
        ok, _value, detail = _digest(bad[0], bad[1], count=bad[2], result=sentinel)
        assert ok is False, bad[2]
        assert sentinel.v == "UNTOUCHED"
        assert "carrier" in detail, detail
    ok, _value, detail = _digest(nominal, pv, count=-1)
    assert ok is False and "negative" in detail


def test_46_the_digest_is_streamed_and_never_concatenated_whole() -> None:
    """The canonical text alive at any moment is ONE record."""
    vba = _transcribe()
    seen: list[int] = []
    real_stream, real_continue = vba["CalcFpDigestStream"], vba["CalcFpContinueDigest"]

    def watched_stream(text, result):
        seen.append(len(_val(text)))
        return real_stream(text, result)

    def watched_continue(prior, suffix, result):
        seen.append(len(_val(suffix)))
        return real_continue(prior, suffix, result)

    nominal = [float(v) for v in range(1, 501)]
    pv = [float(v) / 2.0 for v in range(1, 501)]
    vba["CalcFpDigestStream"] = watched_stream
    vba["CalcFpContinueDigest"] = watched_continue
    try:
        ok, value, detail = _digest(nominal, pv)
    finally:
        vba["CalcFpDigestStream"] = real_stream
        vba["CalcFpContinueDigest"] = real_continue
    assert ok, detail
    assert len(value) == 16
    # 500 records folded one at a time, and nothing longer than one record was
    # ever handed to the hash. A whole-stream build would be tens of thousands.
    assert len(seen) == 501, len(seen)
    assert max(seen) < 200, max(seen)
    body = _procedure("SimFpVersionedResultDigest")
    assert body.count("CalcFpDigestStream") == 1
    assert "CalcFpContinueDigest(running, record, folded)" in body


def test_47_the_result_is_committed_last() -> None:
    body = _procedure("SimFpVersionedResultDigest")
    commit = body.rindex("result = running")
    for stage in ("CalcFpDigestStream", "SimFpDigestRecord", "CalcFpContinueDigest"):
        assert body.index(stage) < commit, stage
    assert body.count("result = ") == 1
    # A failure part way through publishes nothing.
    nominal = [1.0, 2.0, float("inf"), 4.0]
    sentinel = _Ref("UNTOUCHED")
    ok, _value, detail = _digest(nominal, [1.0, 2.0, 3.0, 4.0], result=sentinel)
    assert ok is False and sentinel.v == "UNTOUCHED"
    assert "iteration 3" in detail, detail


def test_48_the_version_is_the_projected_one_and_is_hashed() -> None:
    nominal, pv, _v, _b = _digest_case("digest.base")
    ok, at_one, _d = _versioned_digest(1, nominal, pv)
    ok2, at_two, _d = _versioned_digest(2, nominal, pv)
    ok3, production, _d = _digest(nominal, pv)
    assert ok and ok2 and ok3
    assert at_one != at_two
    assert production == at_one == DIGEST_VECTORS["digest.base"]
    assert at_two == DIGEST_VECTORS["digest.version_2"]
    assert _const("SIM_METHOD_VERSION") == 1
    ok, _value, detail = _versioned_digest(0, nominal, pv)
    assert ok is False and "positive" in detail


# ===========================================================================
# E. The accepted Phase-5 encoder did not move
# ===========================================================================
def test_49_the_accepted_prefix_of_the_encoder_is_byte_identical() -> None:
    import hashlib

    banner = ("' ==========================================================================\n"
              "' STEP 10 ADDITION - THE CANONICAL DIGEST CONTINUATION\n")
    text = CALC_FP_BAS.read_text(encoding="utf-8")
    assert text.count(banner) == 1
    accepted = text[: text.index(banner)]
    assert hashlib.sha256(accepted.encode("utf-8")).hexdigest() == (
        "39e80b9ef9252a9822cd57c8ae441b67571ca3725b3d78124bd6af2ddccc4744"), (
        "an accepted line of modCalcFingerprint moved")
    # Everything after the banner declares ONLY the three named additions.
    added = re.findall(r"^(?:Public|Private) Function (\w+)",
                       text[text.index(banner):], re.M)
    assert added == ["CalcFpContinueDigest", "CalcFpHexValue",
                     "CalcFpHexDigitValue"], added


def test_50_every_locked_phase5_fingerprint_vector_still_holds() -> None:
    assert _digest_stream("PCCM-FP")[1] == "6551C6F365DA7F3F"
    assert fp.fingerprint_probe(["A", "B"]) == "42E49DC715F06970"
    assert fp.fingerprint_probe(["AB", ""]) == "7558FD9248656EAD"
    assert fp.canonical_number(1 / 3) == "3.3333333333333331E-01"
    stream = reference_stream(1)
    assert fp.utf16_length(stream) == 366
    ok, value = _digest_stream(stream)
    assert ok and value == ANALYTICAL_DIGEST


def test_51_every_phase5_case_fingerprint_is_unchanged() -> None:
    document = json.loads((PCCM_ROOT / "build" / "phase5_cases.json")
                          .read_text(encoding="utf-8"))
    block = document["fingerprint"]
    assert block["reference"]["digest"] == ANALYTICAL_DIGEST
    assert block["reference"]["code_units"] == 366
    ok, value = _digest_stream(block["reference"]["stream"])
    assert ok and value == block["reference"]["digest"]
    for probe in block["collision_probes"]:
        assert fp.fingerprint_probe(probe["values"]) == probe["digest"], probe["values"]
    for vector in block["numeric_encodings"]["vectors"]:
        assert fp.canonical_number(vector["value"]) == vector["expected"], vector["label"]


def test_52_the_accepted_field_encoders_still_produce_the_accepted_bytes() -> None:
    vba = _transcribe()
    assert vba["CalcFpCanonicalText"](_Ref("SIM")) == "S3:SIM"
    result = _Ref("")
    assert vba["CalcFpCanonicalInteger"](_Ref(1000), result)
    assert result.v == "I4:1000"
    assert vba["CalcFpCanonicalInteger"](_Ref(0), result) and result.v == "I1:0"
    assert vba["CalcFpCanonicalInteger"](_Ref(-1), result) is False
    assert vba["CalcFpNumberField"](_Ref(1.0), _Ref("."), result)
    assert result.v == "N22:1.0000000000000000E+00"


# ===========================================================================
# F. Scope
# ===========================================================================
def test_53_the_orchestration_layer_arrived_and_nothing_beyond_it() -> None:
    """Step 11 added the reporting module; nothing past it exists."""
    names = {path.name for path in SRC_VBA.glob("*.bas")}
    assert "modSimReport.bas" in names
    # EXECUTABLE code: a module is allowed to say in prose which later-step
    # concepts it refuses to contain - the discipline Step 8 settled for Cheng.
    #
    # The endpoint, the reporting module's own name and the machine sheet belong
    # to modSimReport and to NOBODY ELSE. The scope of this test is unchanged;
    # only its owner exists now.
    for module in load_modules([SRC_VBA]):
        if module.name == "modSimReport":
            continue
        if module.name in ("modSimPostReport", "modSimAnnualStore"):
            # P7-4's orchestrator and P7-6's store READ the published block
            # through the accepted accessors and the contracted coordinates.
            # Neither publishes a simulation and neither owns a run identity -
            # which is asserted where that claim belongs, in the Phase-7 battery
            # - so what they must not have is the ENDPOINT, not the words.
            assert "PCCM_RunSimulation" not in module.code
            continue
        for banned in ("PCCM_RunSimulation", "SimReport", "_SimData"):
            assert banned not in module.code, f"{module.name} carries {banned}"
    report = next(m for m in load_modules([SRC_VBA]) if m.name == "modSimReport")
    assert "PCCM_RunSimulation" in report.code
    # And nothing beyond it EXCEPT what a later phase has landed under its own
    # authority. `modSimReport` is still the last Phase-6 module; P7-2's pure
    # kernel is not a Phase-6 module and is excluded by name rather than by
    # widening the set this control is about.
    assert not [p for p in SRC_VBA.glob("*.bas") if p.stem.startswith("modSim")
                and p.stem not in ("modSimSensitivity", "modSimPostReport", "modSimAnnual",
                                   "modSimAnnualRun", "modSimAnnualStore",
                                   "modSimContract", "modSimRng", "modSimSample",
                                   "modSimEngine", "modSimStats", "modSimFingerprint",
                                   "modSimNonce", "modSimReport")]


def test_54_no_other_module_frames_a_phase6_stream() -> None:
    for module in load_modules([SRC_VBA]):
        if module.name in ("modSimFingerprint", "modCalcFingerprint"):
            continue
        banned = ["SIM_DIGEST_STREAM_TAG", "SIM_DIGEST_SECTION", "SIM_DIGEST_FIELD_",
                  "SIM_REQUEST_", "CalcFpContinueDigest", "CalcFpDigestStream"]
        if module.name != "modCalcReport":
            # The accepted Phase-5 reporter frames its OWN header scalars through
            # the accepted N-field authority; that is Step-7 authorised and is
            # not a Phase-6 stream.
            banned.extend(["CalcFpCanonicalText", "CalcFpCanonicalInteger",
                           "CalcFpNumberField"])
        for token in banned:
            assert token not in module.code, f"{module.name} carries {token}"
    # The ORCHESTRATION layer legitimately reads ONE projected identity - the
    # digest's index origin - because Step-8 element k is iteration
    # origin + k wherever it is written. Reading an index origin is not framing
    # a stream, and every framing constant above is still refused to it.
    report = next(m for m in load_modules([SRC_VBA]) if m.name == "modSimReport")
    assert set(re.findall(r"SIM_DIGEST_\w+", report.code)) == {
        "SIM_DIGEST_INDEX_ORIGIN"}


def test_55_the_transcription_read_the_whole_module() -> None:
    vba = _transcribe()
    compiled = vba["_python_source"]
    assert BORROWED["modSimFingerprint"] == {"SimFpRetainedExtent"}
    for name in _module().procedures:
        assert callable(vba[name]), name
        if name in BORROWED["modSimFingerprint"]:
            assert f"def {name}(" not in compiled, name
        else:
            assert f"def {name}(" in compiled, f"{name} was not compiled from source"
    # The whole accepted hash core is compiled from source too.
    for name in CALC_FP_COMPILED:
        assert f"def {name}(" in compiled, f"{name} was not compiled from source"
    assert "def CalcFpCanonicalNumber(" not in compiled
    assert [(mode, pname) for mode, pname, _a, _t, _k
            in vba["_procs"]["SimFpRetainedExtent"]] == [
        ("ByRef", "totalNominal"), ("ByRef", "totalPv"),
        ("ByRef", "nominalExtent"), ("ByRef", "pvExtent")]


def test_56_the_guarded_extent_helper_is_scoped_and_reads_only_the_extents() -> None:
    executable = _module().code_without_string_removal
    assert "On Error Resume Next" not in executable
    assert re.findall(r"On Error GoTo (\w+)", executable) == ["Unallocated", "0", "0"]
    from pccm_builder.vba_source import logical_statements

    statements = [text for _n, text in
                  logical_statements(_procedure("SimFpRetainedExtent"))]
    assert statements == [
        "Private Function SimFpRetainedExtent(ByRef totalNominal() As Double, "
        "ByRef totalPv() As Double, ByRef nominalExtent As Long, "
        "ByRef pvExtent As Long) As Boolean",
        "On Error GoTo Unallocated",
        "nominalExtent = UBound(totalNominal) - LBound(totalNominal) + 1",
        "pvExtent = UBound(totalPv) - LBound(totalPv) + 1",
        "On Error GoTo 0",
        "SimFpRetainedExtent = True",
        "Exit Function",
        "Unallocated:",
        "On Error GoTo 0",
        "SimFpRetainedExtent = False",
        "End Function",
    ], statements
    # NOTE: the arm that RAISES - a genuinely never-sized VBA array - has no
    # Linux execution proof and is deferred to Gate B. Its SHAPE is pinned here.


if __name__ == "__main__":  # pragma: no cover
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
