#!/usr/bin/env python3
"""PCCM Phase 6 Step-10 MUTATION CONTROLS for the fingerprint conformance battery.

A conformance test that cannot fail proves nothing. Every control damages one of
the two sources - `modSimFingerprint.bas` or `modCalcFingerprint.bas` - reruns
the WHOLE Step-10 battery against the damaged copy, and requires a NAMED
detector among the refusers.

Nothing here writes to the repository: the damaged copies live in a temporary
directory and the conformance module is pointed at them for one control.

Runs standalone or under pytest.
"""

from __future__ import annotations

import signal
import sys
import tempfile
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import test_phase6_sim_fingerprint_vba as conformance  # noqa: E402

_SIM = conformance.SIM_FP_BAS.read_text(encoding="utf-8")
_CALC = conformance.CALC_FP_BAS.read_text(encoding="utf-8")

# The whole accepted battery runs in about one second. A budget of one minute is
# a fiftyfold margin, so only a genuinely non-terminating mutation trips it.
_TEST_BUDGET_SECONDS = 60


class _Timeout(Exception):
    pass


def _conformance_tests() -> list[str]:
    names = sorted(n for n in dir(conformance) if n.startswith("test_"))
    assert len(names) >= 50, names
    return names


def _run_battery() -> list[str]:
    refused = []

    def alarm(signum, frame):  # pragma: no cover - only fires under a mutation
        raise _Timeout("the detector did not terminate")

    previous = signal.signal(signal.SIGALRM, alarm)
    try:
        for name in _conformance_tests():
            signal.setitimer(signal.ITIMER_REAL, _TEST_BUDGET_SECONDS)
            try:
                getattr(conformance, name)()
            except BaseException:  # noqa: BLE001 - any refusal counts
                refused.append(name)
            finally:
                signal.setitimer(signal.ITIMER_REAL, 0)
    finally:
        signal.signal(signal.SIGALRM, previous)
    return refused


def _install(sim: str | None = None, calc: str | None = None):
    saved = (conformance.SIM_FP_BAS, conformance.CALC_FP_BAS, dict(conformance._CACHE))
    conformance._CACHE.clear()
    temp = Path(tempfile.mkdtemp(prefix="pccm-step10-mutation-"))
    if sim is not None:
        assert sim != _SIM, "the mutation changed nothing"
        target = temp / "modSimFingerprint.bas"
        target.write_text(sim, encoding="utf-8")
        conformance.SIM_FP_BAS = target
    if calc is not None:
        assert calc != _CALC, "the mutation changed nothing"
        target = temp / "modCalcFingerprint.bas"
        target.write_text(calc, encoding="utf-8")
        conformance.CALC_FP_BAS = target

    def restore() -> None:
        conformance.SIM_FP_BAS = saved[0]
        conformance.CALC_FP_BAS = saved[1]
        conformance._CACHE.clear()
        conformance._CACHE.update(saved[2])

    return restore


def _control(expected: str, sim: str | None = None, calc: str | None = None) -> None:
    restore = _install(sim, calc)
    try:
        refused = _run_battery()
    finally:
        restore()
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def _swap(text: str, old: str, new: str, count: int = 1) -> str:
    assert text.count(old) == count, (old[:80], text.count(old))
    return text.replace(old, new)


def _after(text: str, anchor: str, inserted: str) -> str:
    return _swap(text, anchor, anchor + inserted)


# ===========================================================================
# The battery must pass on the ACCEPTED sources.
# ===========================================================================
def test_00_the_accepted_sources_pass_every_detector() -> None:
    restore = _install()
    try:
        refused = _run_battery()
    finally:
        restore()
    assert refused == [], refused


# ===========================================================================
# A. The canonical continuation primitive
# ===========================================================================
def test_01_the_hash_is_reinitialised_before_the_suffix() -> None:
    damaged = _swap(
        _CALC,
        "    If h1 < 0# Or h1 >= FP_MOD_1 Then Exit Function\n",
        "    h1 = FP_INIT_1\n    h2 = FP_INIT_2\n"
        "    If h1 < 0# Or h1 >= FP_MOD_1 Then Exit Function\n")
    _control("test_13", calc=damaged)


def test_02_the_two_accumulator_halves_are_swapped_on_output() -> None:
    damaged = _swap(
        _CALC,
        "    result = CalcFpHex8(h1) & CalcFpHex8(h2)\n"
        "    CalcFpContinueDigest = True\n",
        "    result = CalcFpHex8(h2) & CalcFpHex8(h1)\n"
        "    CalcFpContinueDigest = True\n")
    _control("test_13", calc=damaged)


def test_03_one_half_is_decoded_from_the_wrong_digits() -> None:
    damaged = _swap(
        _CALC,
        "    If Not CalcFpHexValue(Mid$(priorDigest, FP_HEX_WIDTH + 1, FP_HEX_WIDTH), h2) "
        "Then Exit Function\n",
        "    If Not CalcFpHexValue(Mid$(priorDigest, 1, FP_HEX_WIDTH), h2) "
        "Then Exit Function\n")
    _control("test_13", calc=damaged)


def test_04_the_code_unit_normalisation_is_dropped() -> None:
    """AscW returns a SIGNED Integer, so everything above U+7FFF comes back
    negative. Without the normalisation those units are wrong or refused."""
    damaged = _swap(
        _CALC,
        "        unit = CalcFpNormaliseCodeUnit(AscW(Mid$(suffix, index, 1)))\n"
        "        If unit < 0 Then Exit Function\n",
        "        unit = AscW(Mid$(suffix, index, 1))\n")
    _control("test_13", calc=damaged)


def test_05_the_suffix_is_not_consumed_in_full() -> None:
    """One code unit short. Every unit of the suffix is part of the stream.

    NOT CONTROLLED HERE, and deliberately: substituting `Asc` for `AscW` is
    invisible to this harness, because the transcriber models a VBA String as
    the UTF-16 sequence it is and has no ANSI code page to narrow through.
    `Asc` versus `AscW` is a REAL-VBA distinction and is listed as deferred
    Gate-B work rather than claimed here.
    """
    damaged = _swap(
        _CALC,
        "    For index = 1 To CalcFpUtf16Length(suffix)\n",
        "    For index = 1 To CalcFpUtf16Length(suffix) - 1\n")
    _control("test_13", calc=damaged)


def test_06_a_lowercase_digest_is_silently_normalised() -> None:
    damaged = _swap(
        _CALC,
        'Private Const FP_HEX_DIGITS As String = "0123456789ABCDEF"',
        'Private Const FP_HEX_DIGITS As String = "0123456789ABCDEFabcdef"')
    damaged = _swap(
        damaged,
        "            CalcFpHexDigitValue = index - 1\n",
        "            If index > FP_HEX_WIDTH + FP_HEX_WIDTH Then\n"
        "                CalcFpHexDigitValue = index - FP_HEX_WIDTH - FP_HEX_WIDTH - 1\n"
        "            Else\n"
        "                CalcFpHexDigitValue = index - 1\n"
        "            End If\n")
    _control("test_17", calc=damaged)


def test_07_a_state_at_its_own_modulus_is_accepted() -> None:
    damaged = _swap(
        _CALC,
        "    If h2 < 0# Or h2 >= FP_MOD_2 Then Exit Function\n",
        "    If h2 < 0# Then Exit Function\n")
    _control("test_18", calc=damaged)


def test_08_the_digest_length_is_no_longer_checked() -> None:
    damaged = _swap(
        _CALC,
        "    If CalcFpUtf16Length(priorDigest) <> FP_HEX_WIDTH + FP_HEX_WIDTH "
        "Then Exit Function\n",
        "")
    _control("test_17", calc=damaged)


def test_09_the_accepted_digest_loop_starts_from_a_supplied_state() -> None:
    """CalcFpDigestStream must keep starting at the LOCKED initial states."""
    damaged = _swap(
        _CALC,
        "    h1 = FP_INIT_1\n    h2 = FP_INIT_2\n",
        "    h1 = FP_INIT_1 + 1\n    h2 = FP_INIT_2\n")
    _control("test_22", calc=damaged)


def test_10_a_second_hash_recurrence_is_added_to_the_framing_module() -> None:
    damaged = _after(
        _SIM, "Option Explicit\n",
        "\nPrivate Function SimFpFold(ByVal h As Double, ByVal u As Double) As Double\n"
        "    Dim x As Double, q As Double\n"
        "    x = h * 131# + u\n"
        "    q = Fix(x / 2147483647#)\n"
        "    SimFpFold = x - q * 2147483647#\n"
        "End Function\n")
    _control("test_10", sim=damaged)


# ===========================================================================
# B. The request fingerprint
# ===========================================================================
def test_11_the_analytical_digest_is_encoded_as_a_text_field() -> None:
    """The defect the whole continuation design exists to prevent."""
    damaged = _swap(
        _SIM,
        "    If Not modCalcFingerprint.CalcFpContinueDigest(analyticalFingerprint, "
        "suffix, candidate) Then\n",
        "    suffix = modCalcFingerprint.CalcFpCanonicalText(analyticalFingerprint) "
        "& suffix\n"
        "    If Not modCalcFingerprint.CalcFpDigestStream(suffix, candidate) Then\n")
    _control("test_23", sim=damaged)


def test_12_a_new_stream_is_started_at_the_sim_section() -> None:
    damaged = _swap(
        _SIM,
        "    If Not modCalcFingerprint.CalcFpContinueDigest(analyticalFingerprint, "
        "suffix, candidate) Then\n",
        "    If Not modCalcFingerprint.CalcFpDigestStream(suffix, candidate) Then\n")
    _control("test_23", sim=damaged)


def test_13_the_extension_gains_a_stream_tag_of_its_own() -> None:
    damaged = _swap(
        _SIM,
        '    built = modCalcFingerprint.CalcFpCanonicalText(SIM_REQUEST_SECTION)\n',
        '    built = modCalcFingerprint.CalcFpCanonicalText("PCCM-FP")\n'
        '    built = built & modCalcFingerprint.CalcFpCanonicalText(SIM_REQUEST_SECTION)\n')
    _control("test_28", sim=damaged)


def test_14_auto_claims_the_fixed_field_count() -> None:
    damaged = _swap(
        _SIM,
        "        If Not modCalcFingerprint.CalcFpCanonicalInteger(SIM_REQUEST_FIELD_COUNT_AUTO, _\n"
        "                                                         encoded) Then\n",
        "        If Not modCalcFingerprint.CalcFpCanonicalInteger(SIM_REQUEST_FIELD_COUNT_FIXED, _\n"
        "                                                         encoded) Then\n")
    _control("test_23", sim=damaged)


def test_15_auto_emits_a_zero_seed_sentinel() -> None:
    damaged = _swap(
        _SIM,
        "    If hasSuppliedSeed Then\n"
        "        If Not modCalcFingerprint.CalcFpCanonicalInteger(suppliedSeed, encoded) Then\n"
        '            detail = "request fingerprint: the supplied seed is not encodable"\n'
        "            Exit Function\n"
        "        End If\n"
        "        built = built & encoded\n"
        "    End If\n",
        "    If hasSuppliedSeed Then\n"
        "        If Not modCalcFingerprint.CalcFpCanonicalInteger(suppliedSeed, encoded) Then\n"
        '            detail = "request fingerprint: the supplied seed is not encodable"\n'
        "            Exit Function\n"
        "        End If\n"
        "    Else\n"
        "        If Not modCalcFingerprint.CalcFpCanonicalInteger(0, encoded) Then\n"
        "            Exit Function\n"
        "        End If\n"
        "    End If\n"
        "    built = built & encoded\n")
    _control("test_26", sim=damaged)


def test_16_fixed_omits_its_supplied_seed() -> None:
    damaged = _swap(
        _SIM,
        "    If hasSuppliedSeed Then\n"
        "        If Not modCalcFingerprint.CalcFpCanonicalInteger(suppliedSeed, encoded) Then\n"
        '            detail = "request fingerprint: the supplied seed is not encodable"\n'
        "            Exit Function\n"
        "        End If\n"
        "        built = built & encoded\n"
        "    End If\n",
        "")
    _control("test_23", sim=damaged)


def test_17_the_iteration_count_is_encoded_as_a_double() -> None:
    damaged = _swap(
        _SIM,
        "    If Not modCalcFingerprint.CalcFpCanonicalInteger(iterations, encoded) Then\n",
        '    If Not modCalcFingerprint.CalcFpNumberField(CDbl(iterations), ".", encoded) Then\n')
    _control("test_23", sim=damaged)


def test_18_the_seed_mode_is_encoded_as_an_integer() -> None:
    damaged = _swap(
        _SIM,
        "    built = built & modCalcFingerprint.CalcFpCanonicalText(seedMode)\n",
        "    If Not modCalcFingerprint.CalcFpCanonicalInteger(SIM_REQUEST_RECORD_COUNT, _\n"
        "                                                     encoded) Then\n"
        "        Exit Function\n"
        "    End If\n"
        "    built = built & encoded\n")
    _control("test_23", sim=damaged)


def test_19_the_supplied_seed_is_encoded_as_a_double() -> None:
    damaged = _swap(
        _SIM,
        "        If Not modCalcFingerprint.CalcFpCanonicalInteger(suppliedSeed, encoded) Then\n",
        '        If Not modCalcFingerprint.CalcFpNumberField(CDbl(suppliedSeed), ".", '
        "encoded) Then\n")
    _control("test_23", sim=damaged)


def test_20_a_version_is_encoded_as_text() -> None:
    damaged = _swap(
        _SIM,
        "    If Not modCalcFingerprint.CalcFpCanonicalInteger(SIM_RNG_VERSION, encoded) Then\n"
        '        detail = "request fingerprint: the generator version is not encodable"\n'
        "        Exit Function\n"
        "    End If\n"
        "    built = built & encoded\n",
        "    built = built & modCalcFingerprint.CalcFpCanonicalText(CStr(SIM_RNG_VERSION))\n")
    _control("test_23", sim=damaged)


def test_21_the_field_order_is_moved() -> None:
    """The versions before the seed. Every field is present; the ORDER is not."""
    damaged = _swap(
        _SIM,
        "    If hasSuppliedSeed Then\n"
        "        If Not modCalcFingerprint.CalcFpCanonicalInteger(suppliedSeed, encoded) Then\n"
        '            detail = "request fingerprint: the supplied seed is not encodable"\n'
        "            Exit Function\n"
        "        End If\n"
        "        built = built & encoded\n"
        "    End If\n\n"
        "    If Not modCalcFingerprint.CalcFpCanonicalInteger(SIM_RNG_VERSION, encoded) Then\n"
        '        detail = "request fingerprint: the generator version is not encodable"\n'
        "        Exit Function\n"
        "    End If\n"
        "    built = built & encoded\n",
        "    If Not modCalcFingerprint.CalcFpCanonicalInteger(SIM_RNG_VERSION, encoded) Then\n"
        '        detail = "request fingerprint: the generator version is not encodable"\n'
        "        Exit Function\n"
        "    End If\n"
        "    built = built & encoded\n\n"
        "    If hasSuppliedSeed Then\n"
        "        If Not modCalcFingerprint.CalcFpCanonicalInteger(suppliedSeed, encoded) Then\n"
        '            detail = "request fingerprint: the supplied seed is not encodable"\n'
        "            Exit Function\n"
        "        End If\n"
        "        built = built & encoded\n"
        "    End If\n")
    _control("test_23", sim=damaged)


def test_22_the_semantic_field_names_are_encoded_into_the_record() -> None:
    damaged = _swap(
        _SIM,
        "    If Not modCalcFingerprint.CalcFpCanonicalInteger(iterations, encoded) Then\n",
        '    built = built & modCalcFingerprint.CalcFpCanonicalText("iterations")\n'
        "    If Not modCalcFingerprint.CalcFpCanonicalInteger(iterations, encoded) Then\n")
    _control("test_23", sim=damaged)


def _extra_request_field(emitted: str) -> str:
    return _swap(
        _SIM,
        "    If Not modCalcFingerprint.CalcFpCanonicalInteger(SIM_METHOD_VERSION, encoded) Then\n",
        emitted +
        "    If Not modCalcFingerprint.CalcFpCanonicalInteger(SIM_METHOD_VERSION, encoded) Then\n")


def test_23_the_effective_seed_enters_the_sim_record() -> None:
    _control("test_23", sim=_extra_request_field(
        "    If Not modCalcFingerprint.CalcFpCanonicalInteger(suppliedSeed, encoded) Then\n"
        "        Exit Function\n"
        "    End If\n"
        "    built = built & encoded\n"))


def test_24_the_auto_nonce_enters_the_sim_record() -> None:
    _control("test_23", sim=_extra_request_field(
        "    If Not modCalcFingerprint.CalcFpCanonicalInteger(SIM_AUTO_NONCE_INITIAL, _\n"
        "                                                     encoded) Then\n"
        "        Exit Function\n"
        "    End If\n"
        "    built = built & encoded\n"))


def test_25_the_run_id_enters_the_sim_record() -> None:
    _control("test_23", sim=_extra_request_field(
        "    If Not modCalcFingerprint.CalcFpCanonicalInteger(SIM_RUN_ID_FIRST, encoded) Then\n"
        "        Exit Function\n"
        "    End If\n"
        "    built = built & encoded\n"))


def test_26_the_selected_confidence_level_enters_the_sim_record() -> None:
    """Also caught statically: this module may not name a quantile at all."""
    _control("test_08", sim=_extra_request_field(
        "    built = built & modCalcFingerprint.CalcFpCanonicalText(SIM_QUANTILE_HEADLINE_3)\n"))


def test_27_the_result_is_written_before_the_request_is_validated() -> None:
    damaged = _swap(
        _SIM,
        "    detail = vbNullString\n"
        "    If Not SimFpValidateRequest(iterations, seedMode, hasSuppliedSeed, "
        "suppliedSeed, detail) Then\n",
        "    detail = vbNullString\n"
        "    result = analyticalFingerprint\n"
        "    If Not SimFpValidateRequest(iterations, seedMode, hasSuppliedSeed, "
        "suppliedSeed, detail) Then\n")
    _control("test_34", sim=damaged)


def test_28_the_seed_mode_is_matched_case_insensitively() -> None:
    """A mode the accepted grammar never spells would reach the stream as itself."""
    damaged = _swap(
        _SIM,
        "    If Not isAuto And Not isFixed Then\n"
        '        detail = "request fingerprint: an unknown seed mode"\n'
        "        Exit Function\n"
        "    End If\n",
        "    If Not isAuto And Not isFixed Then\n"
        "        isAuto = True\n"
        "    End If\n")
    _control("test_30", sim=damaged)


def test_29_the_iteration_bounds_are_not_checked() -> None:
    damaged = _swap(
        _SIM,
        "    If iterations < SIM_MIN_ITERATIONS Then\n"
        '        detail = "request fingerprint: fewer iterations than the business minimum"\n'
        "        Exit Function\n"
        "    End If\n",
        "")
    _control("test_29", sim=damaged)


def test_30_the_seed_domain_is_not_checked() -> None:
    damaged = _swap(
        _SIM,
        "        If suppliedSeed < SIM_SEED_MIN Or suppliedSeed > SIM_SEED_MAX Then\n"
        '            detail = "request fingerprint: the supplied seed is outside its '
        'accepted domain"\n'
        "            Exit Function\n"
        "        End If\n",
        "")
    _control("test_32", sim=damaged)


def test_31_the_flag_and_the_mode_no_longer_have_to_agree() -> None:
    damaged = _swap(
        _SIM,
        "    If isAuto And hasSuppliedSeed Then\n"
        '        detail = "request fingerprint: an AUTO request carries no supplied seed"\n'
        "        Exit Function\n"
        "    End If\n",
        "")
    _control("test_31", sim=damaged)


# ===========================================================================
# C. The result digest
# ===========================================================================
def test_32_the_result_stream_tag_is_changed() -> None:
    damaged = _swap(
        _SIM,
        "    prefix = modCalcFingerprint.CalcFpCanonicalText(SIM_DIGEST_STREAM_TAG)\n",
        "    prefix = modCalcFingerprint.CalcFpCanonicalText(SIM_REQUEST_SECTION)\n")
    _control("test_35", sim=damaged)


def test_33_the_version_field_is_omitted() -> None:
    damaged = _swap(
        _SIM,
        "    If Not modCalcFingerprint.CalcFpCanonicalInteger(methodVersion, encoded) Then\n"
        '        detail = "result digest: the method version is not encodable"\n'
        "        Exit Function\n"
        "    End If\n"
        "    prefix = prefix & encoded\n",
        "")
    _control("test_35", sim=damaged)


def test_34_the_section_name_is_changed() -> None:
    damaged = _swap(
        _SIM,
        "    prefix = prefix & modCalcFingerprint.CalcFpCanonicalText(SIM_DIGEST_SECTION)\n",
        "    prefix = prefix & modCalcFingerprint.CalcFpCanonicalText(SIM_REQUEST_SECTION)\n")
    _control("test_35", sim=damaged)


def test_35_the_record_count_is_omitted() -> None:
    damaged = _swap(
        _SIM,
        "    If Not modCalcFingerprint.CalcFpCanonicalInteger(sampleCount, encoded) Then\n"
        '        detail = "result digest: the record count is not encodable"\n'
        "        Exit Function\n"
        "    End If\n"
        "    prefix = prefix & encoded\n",
        "")
    _control("test_35", sim=damaged)


def test_36_the_record_field_count_is_wrong() -> None:
    damaged = _swap(
        _SIM,
        "    If Not modCalcFingerprint.CalcFpCanonicalInteger(SIM_DIGEST_FIELD_COUNT, "
        "encoded) Then\n",
        "    If Not modCalcFingerprint.CalcFpCanonicalInteger(SIM_REQUEST_FIELD_COUNT_AUTO, "
        "encoded) Then\n")
    _control("test_35", sim=damaged)


def test_37_the_iteration_index_becomes_zero_based() -> None:
    damaged = _SIM.replace("SIM_DIGEST_INDEX_ORIGIN + offset", "offset")
    assert damaged != _SIM
    _control("test_35", sim=damaged)


def test_38_the_physical_array_index_is_used_as_the_iteration_identity() -> None:
    """Invisible in the numbers whenever LBound is zero. Only source sees it."""
    damaged = _swap(
        _SIM,
        "        If Not SimFpDigestRecord(offset, totalNominal(LBound(totalNominal) + offset), _\n",
        "        If Not SimFpDigestRecord(LBound(totalNominal) + offset, _\n"
        "                                 totalNominal(LBound(totalNominal) + offset), _\n")
    _control("test_38", sim=damaged)


def test_39_the_two_measures_are_swapped() -> None:
    damaged = _swap(
        _SIM,
        "        If Not SimFpDigestRecord(offset, totalNominal(LBound(totalNominal) + offset), _\n"
        "                                 totalPv(LBound(totalPv) + offset), decimalSeparator, _\n",
        "        If Not SimFpDigestRecord(offset, totalPv(LBound(totalPv) + offset), _\n"
        "                                 totalNominal(LBound(totalNominal) + offset), "
        "decimalSeparator, _\n")
    _control("test_35", sim=damaged)


def test_40_a_total_is_encoded_as_text_instead_of_a_number() -> None:
    damaged = _swap(
        _SIM,
        "    If Not modCalcFingerprint.CalcFpNumberField(nominal, decimalSeparator, "
        "encoded) Then\n"
        '        detail = "result digest: the retained nominal total is not canonically '
        'encodable"\n'
        "        Exit Function\n"
        "    End If\n"
        "    built = built & encoded\n",
        "    built = built & modCalcFingerprint.CalcFpCanonicalText(CStr(nominal))\n")
    _control("test_35", sim=damaged)


def test_41_the_retained_order_is_reversed() -> None:
    damaged = _swap(
        _SIM,
        "    For offset = 0 To sampleCount - 1\n",
        "    For offset = sampleCount - 1 To 0 Step -1\n")
    _control("test_35", sim=damaged)


def test_42_a_sort_is_introduced_before_the_digest() -> None:
    damaged = _after(
        _SIM, "Option Explicit\n",
        "\nPrivate Function SimFpSortAscending(ByRef series() As Double, _\n"
        "                                    ByVal count As Long) As Boolean\n"
        "    Dim outer As Long, inner As Long, held As Double\n"
        "    For outer = 1 To count - 1\n"
        "        held = series(outer)\n"
        "        inner = outer - 1\n"
        "        Do While inner >= 0\n"
        "            If series(inner) <= held Then Exit Do\n"
        "            series(inner + 1) = series(inner)\n"
        "            inner = inner - 1\n"
        "        Loop\n"
        "        series(inner + 1) = held\n"
        "    Next outer\n"
        "    SimFpSortAscending = True\n"
        "End Function\n")
    _control("test_40", sim=damaged)


def test_43_one_retained_iteration_is_dropped() -> None:
    damaged = _swap(
        _SIM,
        "    For offset = 0 To sampleCount - 1\n",
        "    For offset = 0 To sampleCount - 2\n")
    _control("test_35", sim=damaged)


def test_44_the_whole_canonical_stream_is_concatenated_before_hashing() -> None:
    damaged = _swap(
        _SIM,
        "    If Not modCalcFingerprint.CalcFpDigestStream(prefix, running) Then\n"
        '        detail = "result digest: the framing prefix could not be digested"\n'
        "        Exit Function\n"
        "    End If\n",
        "    running = prefix\n")
    damaged = _swap(
        damaged,
        "        If Not modCalcFingerprint.CalcFpContinueDigest(running, record, folded) Then\n"
        '            detail = "result digest: the running digest could not be continued '
        'at iteration " & _\n'
        "                     CStr(SIM_DIGEST_INDEX_ORIGIN + offset)\n"
        "            Exit Function\n"
        "        End If\n"
        "        running = folded\n"
        "    Next offset\n",
        "        running = running & record\n"
        "    Next offset\n"
        "    If Not modCalcFingerprint.CalcFpDigestStream(running, folded) Then\n"
        "        Exit Function\n"
        "    End If\n"
        "    running = folded\n")
    _control("test_46", sim=damaged)


def test_45_a_caller_selected_method_version_is_exposed() -> None:
    damaged = _swap(
        _SIM,
        "Public Function SimFpResultDigest(ByRef totalNominal() As Double, "
        "ByRef totalPv() As Double, _\n"
        "                                  ByVal sampleCount As Long, _\n",
        "Public Function SimFpResultDigest(ByRef totalNominal() As Double, "
        "ByRef totalPv() As Double, _\n"
        "                                  ByVal sampleCount As Long, _\n"
        "                                  ByVal methodVersion As Long, _\n")
    damaged = _swap(
        damaged,
        "    SimFpResultDigest = SimFpVersionedResultDigest(SIM_METHOD_VERSION, "
        "totalNominal, totalPv, _\n",
        "    SimFpResultDigest = SimFpVersionedResultDigest(methodVersion, "
        "totalNominal, totalPv, _\n")
    _control("test_04", sim=damaged)


def test_46_a_canonical_number_formatter_is_duplicated() -> None:
    damaged = _after(
        _SIM, "Option Explicit\n",
        "\nPrivate Function SimFpNumberText(ByVal value As Double) As String\n"
        '    SimFpNumberText = CStr(value) & "E+00"\n'
        "End Function\n")
    _control("test_09", sim=damaged)


def test_47_a_non_finite_retained_total_is_skipped() -> None:
    damaged = _swap(
        _SIM,
        "    If Not IsUsableDouble(nominal) Then\n"
        '        detail = "result digest: the retained nominal total at iteration " & _\n'
        '                 CStr(SIM_DIGEST_INDEX_ORIGIN + offset) & " is not a finite Double"\n'
        "        Exit Function\n"
        "    End If\n",
        "    If Not IsUsableDouble(nominal) Then\n"
        "        nominal = 0#\n"
        "    End If\n")
    _control("test_44", sim=damaged)


def test_48_a_partial_digest_is_published_during_the_loop() -> None:
    damaged = _swap(
        _SIM,
        "        running = folded\n"
        "    Next offset\n",
        "        running = folded\n"
        "        result = running\n"
        "    Next offset\n")
    _control("test_47", sim=damaged)


def test_49_a_bound_is_read_before_the_zero_count_guard() -> None:
    damaged = _swap(
        _SIM,
        "    If sampleCount > 0 Then\n"
        "        If Not SimFpRetainedExtent(totalNominal, totalPv, nominalExtent, "
        "pvExtent) Then\n"
        '            detail = "result digest: the retained carrier is not allocated"\n'
        "            Exit Function\n"
        "        End If\n",
        "    If Not SimFpRetainedExtent(totalNominal, totalPv, nominalExtent, "
        "pvExtent) Then\n"
        '        detail = "result digest: the retained carrier is not allocated"\n'
        "        Exit Function\n"
        "    End If\n"
        "    If sampleCount > 0 Then\n")
    _control("test_37", sim=damaged)


def test_50_the_carrier_length_checks_are_removed() -> None:
    damaged = _swap(
        _SIM,
        "        If nominalExtent <> sampleCount Then\n"
        '            detail = "result digest: the retained nominal carrier is not the '
        'claimed length"\n'
        "            Exit Function\n"
        "        End If\n"
        "        If pvExtent <> sampleCount Then\n"
        '            detail = "result digest: the retained PV carrier is not the '
        'claimed length"\n'
        "            Exit Function\n"
        "        End If\n",
        "")
    _control("test_45", sim=damaged)


if __name__ == "__main__":  # pragma: no cover
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
