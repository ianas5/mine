#!/usr/bin/env python3
"""PCCM Phase 6 Step-10A - the request-fingerprint golden vectors.

WHAT THIS FILE IS FOR
--------------------------------------------------------------------------------
The emitter reads its shape from `sim_contract.yaml` and hands the bytes to the
accepted `calc_fingerprint` hash. A test that called the emitter twice and
compared the answers would prove nothing at all, so EVERY expected suffix, code
unit count and digest below is a LITERAL. If the grammar moves, these fail.

THE ANALYTICAL PREFIX IS NOT REGENERATED. It is the accepted Phase-5 case-26
reference stream, digesting to `50B6EB0E26857EA7`, taken from
`calc_cases.reference_stream` unchanged. Its HEADER/COST/RISK bytes are not
re-encoded here and are not hashed as a field.

THERE IS NO SECOND HASH IMPLEMENTATION. `calc_fingerprint.py` owns the
mathematics; this file supplies a stream and pins the answer.

NO VBA EXISTS FOR THIS YET. `modSimFingerprint` is not authorised, so nothing
here may be read as "VBA produced this fingerprint". This is byte-grammar
authority and test data only.

Runs standalone or under pytest.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import (  # noqa: E402
    load_calc_contract,
    load_contract,
    load_sim_contract,
)
from pccm_builder import calc_fingerprint as fp  # noqa: E402
from pccm_builder.calc_cases import reference_stream  # noqa: E402
from pccm_builder.sim_cases import (  # noqa: E402
    request_fingerprint,
    request_fingerprint_stream,
    request_sim_section,
)
from pccm_builder.sim_oracle import SimOracleError  # noqa: E402

SPEC = PCCM_ROOT / "spec"
CASES_JSON = PCCM_ROOT / "build" / "phase6_cases.json"

ANALYTICAL_DIGEST = "50B6EB0E26857EA7"
"""The accepted Phase-5 case-26 fingerprint. Unchanged by this step."""

# ---------------------------------------------------------------------------
# THE GOLDEN VECTORS, as literals
# ---------------------------------------------------------------------------
AUTO_1000_SUFFIX = "S3:SIMI1:1I1:4I4:1000S4:AUTOI1:1I1:1"
AUTO_1000_DIGEST = "5EAB16E15C2ECE24"

FIXED_SEED_1_SUFFIX = "S3:SIMI1:1I1:5I4:1000S5:FIXEDI1:1I1:1I1:1"
FIXED_SEED_1_DIGEST = "599C95E7274759B9"

FIXED_SEED_MAX_SUFFIX = "S3:SIMI1:1I1:5I4:1000S5:FIXEDI10:2147483646I1:1I1:1"
FIXED_SEED_MAX_DIGEST = "0010FB954CC94B53"

AUTO_1001_SUFFIX = "S3:SIMI1:1I1:4I4:1001S4:AUTOI1:1I1:1"
AUTO_1001_DIGEST = "4777C8BC35F0FFEF"

_CACHE: dict[str, object] = {}


def _sim():
    if "sim" not in _CACHE:
        _CACHE["sim"] = load_sim_contract(SPEC / "sim_contract.yaml")
    return _CACHE["sim"]


def _calc():
    if "calc" not in _CACHE:
        _CACHE["calc"] = load_calc_contract(SPEC / "calc_contract.yaml")
    return _CACHE["calc"]


def _inputs():
    if "inputs" not in _CACHE:
        _CACHE["inputs"] = load_contract(SPEC / "input_contract.yaml")
    return _CACHE["inputs"]


def _cases() -> dict[str, dict]:
    if "cases" not in _CACHE:
        corpus = json.loads(CASES_JSON.read_text(encoding="utf-8"))
        _CACHE["cases"] = {case["id"]: case
                           for group in corpus["groups"] for case in group["cases"]}
    return _CACHE["cases"]  # type: ignore[return-value]


def _suffix(iterations: int, seed_mode: str, supplied_seed=None) -> str:
    return request_sim_section(_sim(), iterations, seed_mode, supplied_seed)


def _digest(iterations: int, seed_mode: str, supplied_seed=None) -> str:
    return request_fingerprint(_sim(), _calc(), iterations, seed_mode, supplied_seed)


# ===========================================================================
# A. The prefix is the accepted analytical stream, untouched
# ===========================================================================
def test_01_the_analytical_prefix_is_the_accepted_case_26_stream() -> None:
    prefix = reference_stream(_calc().fingerprint_version)
    assert fp.fingerprint(prefix) == ANALYTICAL_DIGEST
    assert fp.utf16_length(prefix) == 366
    assert prefix.startswith('S7:PCCM-FPI1:1S6:HEADER')
    assert prefix.endswith("S4:RISKI1:0")


def test_02_the_request_stream_is_the_prefix_plus_the_extension() -> None:
    prefix = reference_stream(_calc().fingerprint_version)
    for iterations, mode, seed in ((1000, "AUTO", None), (1000, "FIXED", 1)):
        stream = request_fingerprint_stream(_sim(), _calc(), iterations, mode, seed)
        assert stream.startswith(prefix), "the analytical bytes were re-encoded"
        assert stream[: len(prefix)] == prefix
        assert stream[len(prefix):] == _suffix(iterations, mode, seed)


def test_03_the_extension_carries_no_stream_tag_of_its_own() -> None:
    for suffix in (AUTO_1000_SUFFIX, FIXED_SEED_1_SUFFIX,
                   FIXED_SEED_MAX_SUFFIX, AUTO_1001_SUFFIX):
        assert "PCCM-FP" not in suffix
        assert suffix.startswith("S3:SIM")


# ===========================================================================
# B. The four golden vectors, pinned as literals
# ===========================================================================
def test_04_auto_at_one_thousand_iterations() -> None:
    assert _suffix(1000, "AUTO") == AUTO_1000_SUFFIX
    assert fp.utf16_length(AUTO_1000_SUFFIX) == 36
    assert _digest(1000, "AUTO") == AUTO_1000_DIGEST


def test_05_fixed_at_the_lowest_accepted_seed() -> None:
    assert _suffix(1000, "FIXED", 1) == FIXED_SEED_1_SUFFIX
    assert _digest(1000, "FIXED", 1) == FIXED_SEED_1_DIGEST


def test_06_fixed_at_the_highest_accepted_seed() -> None:
    from pccm_builder.sim_rng import _seed_domain

    minimum, maximum = _seed_domain(_inputs())
    assert (minimum, maximum) == (1, 2147483646), "the seed domain moved"
    assert _suffix(1000, "FIXED", maximum) == FIXED_SEED_MAX_SUFFIX
    assert "I10:2147483646" in FIXED_SEED_MAX_SUFFIX
    assert _digest(1000, "FIXED", maximum) == FIXED_SEED_MAX_DIGEST


def test_07_one_more_iteration_is_a_different_request() -> None:
    assert _suffix(1001, "AUTO") == AUTO_1001_SUFFIX
    assert _digest(1001, "AUTO") == AUTO_1001_DIGEST


def test_08_the_four_vectors_are_pairwise_distinct() -> None:
    digests = [AUTO_1000_DIGEST, FIXED_SEED_1_DIGEST,
               FIXED_SEED_MAX_DIGEST, AUTO_1001_DIGEST]
    assert len(set(digests)) == 4, digests
    # The three relations the authority names explicitly.
    assert AUTO_1000_DIGEST != FIXED_SEED_1_DIGEST
    assert FIXED_SEED_1_DIGEST != FIXED_SEED_MAX_DIGEST
    assert AUTO_1000_DIGEST != AUTO_1001_DIGEST
    # And the streams differ too, so the distinctness is not a hash accident.
    assert len({AUTO_1000_SUFFIX, FIXED_SEED_1_SUFFIX,
                FIXED_SEED_MAX_SUFFIX, AUTO_1001_SUFFIX}) == 4


# ===========================================================================
# C. The shape rules, exercised
# ===========================================================================
def test_09_auto_hashes_four_fields_and_fixed_hashes_five() -> None:
    assert _suffix(1000, "AUTO").startswith("S3:SIMI1:1I1:4")
    assert _suffix(1000, "FIXED", 1).startswith("S3:SIMI1:1I1:5")
    # I1:1 immediately after the section name IS the record count.
    for suffix in (AUTO_1000_SUFFIX, FIXED_SEED_1_SUFFIX):
        assert suffix[: len("S3:SIM") + len("I1:1")] == "S3:SIMI1:1"


def test_10_an_auto_seed_is_absent_and_no_sentinel_takes_its_place() -> None:
    prefix = reference_stream(_calc().fingerprint_version)
    auto = _suffix(1000, "AUTO")
    assert auto == AUTO_1000_SUFFIX

    # Each rejected representation, written out as the stream it WOULD have
    # produced. Every one is a different suffix and a different fingerprint, so
    # "absent" is a substantive choice and not a wording preference.
    rejected = {
        "F_I(0) sentinel":
            'S3:SIMI1:1I1:5I4:1000S4:AUTO' + fp.integer_field(0).encode() + "I1:1I1:1",
        "blank F_S":
            'S3:SIMI1:1I1:5I4:1000S4:AUTO' + fp.text_field("").encode() + "I1:1I1:1",
        "previous effective seed":
            'S3:SIMI1:1I1:5I4:1000S4:AUTO' + fp.integer_field(12345).encode() + "I1:1I1:1",
    }
    for label, suffix in rejected.items():
        assert suffix != auto, label
        assert fp.fingerprint(prefix + suffix) != AUTO_1000_DIGEST, label
        # ...and each one claims FIVE fields, which is the FIXED shape.
        assert suffix.startswith("S3:SIMI1:1I1:5"), label
    assert auto.startswith("S3:SIMI1:1I1:4"), "AUTO must hash four fields"


def test_11_passing_a_seed_to_auto_is_refused_rather_than_dropped() -> None:
    try:
        request_sim_section(_sim(), 1000, "AUTO", 12345)
    except SimOracleError:
        return
    raise AssertionError("AUTO silently ignored a supplied seed")


def test_12_fixed_without_a_seed_is_refused() -> None:
    try:
        request_sim_section(_sim(), 1000, "FIXED", None)
    except SimOracleError:
        return
    raise AssertionError("FIXED was built without its supplied seed")


def test_13_an_unknown_seed_mode_is_refused() -> None:
    for mode in ("auto", "Auto", "RANDOM", ""):
        try:
            request_sim_section(_sim(), 1000, mode)
        except SimOracleError:
            continue
        raise AssertionError(f"{mode!r} was accepted as a seed mode")


def test_14_the_versions_hashed_are_the_contract_s_own() -> None:
    sim = _sim()
    assert sim.rng_version == 1 and sim.sim_method_version == 1, (
        "a version bumped; the golden literals above must be re-pinned deliberately"
    )
    assert _suffix(1000, "AUTO").endswith("I1:1I1:1")


# ===========================================================================
# D. The emitted corpus carries the same literals
# ===========================================================================
def test_15_the_corpus_pins_the_same_four_vectors() -> None:
    expected = {
        "request_fingerprint.auto.1000": (AUTO_1000_SUFFIX, AUTO_1000_DIGEST, 1000),
        "request_fingerprint.fixed.seed_1": (FIXED_SEED_1_SUFFIX, FIXED_SEED_1_DIGEST, 1000),
        "request_fingerprint.fixed.seed_max": (
            FIXED_SEED_MAX_SUFFIX, FIXED_SEED_MAX_DIGEST, 1000),
        "request_fingerprint.auto.1001": (AUTO_1001_SUFFIX, AUTO_1001_DIGEST, 1001),
    }
    cases = _cases()
    for identifier, (suffix, digest, iterations) in expected.items():
        case = cases[identifier]
        assert case["comparison"] == "EXACT", identifier
        assert case["layer"] == "I_request_fingerprint", identifier
        assert case["inputs"]["iterations"] == iterations, identifier
        assert case["inputs"]["analytical_prefix_digest"] == ANALYTICAL_DIGEST
        assert case["expected_exact"]["sim_suffix"] == suffix, identifier
        assert case["expected_exact"]["request_fingerprint"] == digest, identifier
        assert case["expected_exact"]["sim_suffix_code_units"] == fp.utf16_length(suffix)


def test_16_the_corpus_carries_the_grammar_authority_case() -> None:
    case = _cases()["request_fingerprint.grammar"]
    block = case["expected_exact"]
    assert case["comparison"] == "EXACT"
    assert block["section_order"] == ["HEADER", "COST", "RISK", "SIM"]
    assert block["analytical_prefix"] == ["HEADER", "COST", "RISK"]
    assert block["section_name"] == "SIM"
    assert block["record_count"] == 1
    assert block["fields"] == ["iterations", "seed_mode", "supplied_seed",
                               "rng_version", "sim_method_version"]
    assert block["field_types"] == ["F_I", "F_S", "F_I", "F_I", "F_I"]
    assert block["encoded_field_names"] is False
    assert block["auto_field_count"] == 4
    assert block["fixed_field_count"] == 5
    assert "supplied_seed" not in block["auto_fields"]
    assert "supplied_seed" in block["fixed_fields"]
    assert block["auto_supplied_seed_representation"] == "absent"
    assert block["stream_tag_repeated_in_extension"] is False
    assert block["analytical_fingerprint_hashed_as_a_field"] is False
    assert block["excluded_fields"] == ["effective_seed", "auto_nonce", "run_id",
                                        "selected_confidence_level"]
    assert set(block["grammar"]) == {"section", "auto_record", "fixed_record"}


def test_17_the_corpus_group_is_exactly_these_five_cases() -> None:
    corpus = json.loads(CASES_JSON.read_text(encoding="utf-8"))
    group = [g for g in corpus["groups"] if g["group"] == "I_request_fingerprint"]
    assert len(group) == 1, "the request-fingerprint group is missing or duplicated"
    ids = [case["id"] for case in group[0]["cases"]]
    assert ids == [
        "request_fingerprint.grammar",
        "request_fingerprint.auto.1000",
        "request_fingerprint.fixed.seed_1",
        "request_fingerprint.fixed.seed_max",
        "request_fingerprint.auto.1001",
    ], ids


# ===========================================================================
# E. Scope: no implementation exists yet
# ===========================================================================
def test_18_the_implementation_arrived_in_step_10_and_nothing_beyond_it() -> None:
    """Step 10A closed the grammar with NO implementation. Step 10 then built it.

    This test moved with that authorisation rather than being deleted: it still
    says exactly what may exist, and `modSimReport` and the endpoint still may
    not.
    """
    src = PCCM_ROOT / "src" / "vba"
    names = {path.name for path in src.glob("*.bas")}
    assert "modSimFingerprint.bas" in names
    assert "modSimReport.bas" in names, "Step 11 built the orchestration layer"
    from pccm_builder.vba_source import load_modules

    for module in load_modules([src]):
        if module.name == "modSimPostReport":
            # P7-4's orchestrator READS the published status through the
            # accepted accessor, which necessarily names its owner. What it must
            # not carry is the ENDPOINT - the ability to start a run - and that
            # is what stays asserted here.
            assert "PCCM_RunSimulation" not in module.code
            continue
        if module.name != "modSimReport":
            for banned in ("PCCM_RunSimulation", "SimReport"):
                assert banned not in module.code, f"{module.name} carries {banned}"
        # THE FRAMING IS STILL OWNED BY ONE MODULE. modSimReport CALLS the two
        # public entry points; no third module may name either.
        if module.name not in ("modSimFingerprint", "modSimReport"):
            for owned in ("SimFpBuildRequestFingerprint", "SimFpResultDigest"):
                assert owned not in module.code, f"{module.name} carries {owned}"
    report = next(m for m in load_modules([src]) if m.name == "modSimReport")
    for private in ("SimFpRequestSuffix", "SimFpVersionedResultDigest",
                    "SimFpDigestRecord", "SimFpRetainedExtent"):
        assert private not in report.code, private


def test_19_the_accepted_phase5_encoder_was_not_reopened() -> None:
    """No continuation API, no second hash, no new encoder."""
    text = (PCCM_ROOT / "builder" / "pccm_builder" / "calc_fingerprint.py").read_text(
        encoding="utf-8")
    for invented in ("def continue_stream", "def append_section", "SIM_FP_VERSION",
                     "request_fingerprint"):
        assert invented not in text, invented
    vba = (PCCM_ROOT / "src" / "vba" / "modCalcFingerprint.bas").read_text(encoding="utf-8")
    for invented in ('"SIM"', "SIM_REQUEST_", "SimRequest", "seed_mode"):
        assert invented not in vba, f"the Phase-5 VBA encoder gained {invented}"


if __name__ == "__main__":  # pragma: no cover
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
