#!/usr/bin/env python3
"""PCCM Phase 6 Step-7 MUTATION CONTROLS for the modSimSample conformance battery.

A conformance test that cannot fail proves nothing. Every control here damages
the accepted source, reruns the WHOLE Step-7 conformance battery against the
damaged copy, and requires at least one real detector to refuse it. The name of
a refusing detector is asserted too, so a control cannot quietly degrade into
"something, somewhere, went red".

Nothing here writes to the repository: the damaged copy lives in a temporary
directory and the conformance module is pointed at it for one control.

Runs standalone or under pytest.
"""

from __future__ import annotations

import re
import signal
import sys
import tempfile
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import test_phase6_sim_sample_vba as conformance  # noqa: E402

_ORIGINAL = conformance.SIM_SAMPLE_BAS.read_text(encoding="utf-8")
_TEST_BUDGET_SECONDS = 5


class _Timeout(Exception):
    pass


def _conformance_tests() -> list[str]:
    names = sorted(n for n in dir(conformance) if n.startswith("test_"))
    assert len(names) >= 55, names
    return names


def _run_battery() -> list[str]:
    """Every conformance test that REFUSES the currently installed source."""
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


def _install(source: str | None):
    saved = (conformance.SIM_SAMPLE_BAS, dict(conformance._CACHE))
    conformance._CACHE.clear()
    if source is not None:
        assert source != _ORIGINAL, "the mutation changed nothing"
        temp = Path(tempfile.mkdtemp(prefix="pccm-step7-mutation-"))
        target = temp / "modSimSample.bas"
        target.write_text(source, encoding="utf-8")
        conformance.SIM_SAMPLE_BAS = target

    def restore() -> None:
        conformance.SIM_SAMPLE_BAS = saved[0]
        conformance._CACHE.clear()
        conformance._CACHE.update(saved[1])

    return restore


def _control(expected: str, source: str) -> None:
    """Assert the battery refuses the damage, `expected` among the refusers."""
    restore = _install(source)
    try:
        refused = _run_battery()
    finally:
        restore()
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def _swap(text: str, old: str, new: str, count: int = 1) -> str:
    assert text.count(old) == count, (old[:70], text.count(old))
    return text.replace(old, new)


def _after(text: str, anchor: str, inserted: str) -> str:
    return _swap(text, anchor, anchor + inserted)


# ===========================================================================
# The battery must pass on the ACCEPTED source, or every control below is
# measuring noise.
# ===========================================================================
def test_00_the_accepted_source_passes_every_detector() -> None:
    restore = _install(None)
    try:
        refused = _run_battery()
    finally:
        restore()
    assert refused == [], refused


# ===========================================================================
# 1-4. Uniform
# ===========================================================================
_UNIFORM_SIGNATURE = """Public Function SimSampleUniform(ByRef state As SimRngState, _
                                 ByVal minValue As Double, ByVal maxValue As Double, _
                                 ByRef sample As Double, ByRef uniformsConsumed As Long, _
                                 ByRef detail As String) As Boolean"""
_UNIFORM_WITH_ML = """Public Function SimSampleUniform(ByRef state As SimRngState, _
                                 ByVal minValue As Double, ByVal mostLikely As Double, _
                                 ByVal maxValue As Double, _
                                 ByRef sample As Double, ByRef uniformsConsumed As Long, _
                                 ByRef detail As String) As Boolean"""


def test_01_uniform_reads_most_likely() -> None:
    damaged = _swap(_ORIGINAL, _UNIFORM_SIGNATURE, _UNIFORM_WITH_ML)
    damaged = _swap(
        damaged,
        "    candidate = (1# - u) * minValue + u * maxValue\n",
        "    candidate = (1# - u) * minValue + u * maxValue + 0# * mostLikely\n")
    _control("test_47", damaged)


def test_02_uniform_degeneracy_changed_to_the_three_way_predicate() -> None:
    damaged = _swap(_ORIGINAL, _UNIFORM_SIGNATURE, _UNIFORM_WITH_ML)
    damaged = _swap(
        damaged,
        "    If minValue = maxValue Then\n        sample = minValue",
        "    If minValue = maxValue And minValue = mostLikely Then\n        sample = minValue")
    _control("test_47", damaged)


def test_03_uniform_uses_the_naive_difference() -> None:
    _control("test_47", _swap(
        _ORIGINAL,
        "    candidate = (1# - u) * minValue + u * maxValue\n",
        "    candidate = minValue + u * (maxValue - minValue)\n"))


def test_04_a_degenerate_distribution_consumes_a_draw() -> None:
    _control("test_19", _swap(
        _ORIGINAL,
        "    If minValue = maxValue Then\n"
        "        sample = minValue\n"
        "        uniformsConsumed = 0\n",
        "    If minValue = maxValue Then\n"
        "        working = state\n"
        "        If Not SimRngNextUniform(working, u, detail) Then Exit Function\n"
        "        state = working\n"
        "        sample = minValue\n"
        "        uniformsConsumed = 1\n"))


# ===========================================================================
# 5-8. Triangular and ordering
# ===========================================================================
def test_05_the_triangular_branch_boundary_narrowed() -> None:
    _control("test_48", _swap(_ORIGINAL, "    If u <= c Then\n", "    If u < c Then\n"))


def test_06_the_triangular_branches_swapped() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        conditioned = an + Sqr(u * span * (mn - an))\n"
        "    Else\n"
        "        conditioned = bn - Sqr((1# - u) * span * (bn - mn))\n",
        "        conditioned = bn - Sqr((1# - u) * span * (bn - mn))\n"
        "    Else\n"
        "        conditioned = an + Sqr(u * span * (mn - an))\n")
    _control("test_48", damaged)


def test_07_the_triangular_conditioning_removed() -> None:
    _control("test_48", _swap(
        _ORIGINAL,
        "    s = SimSampleScale(minValue, mostLikely, maxValue)\n"
        "    an = minValue / s\n"
        "    mn = mostLikely / s\n"
        "    bn = maxValue / s\n"
        "    span = bn - an\n"
        "    If span <= 0# Then\n"
        '        detail = "Triangular: the conditioned support has no width"\n',
        "    s = 1#\n"
        "    an = minValue / s\n"
        "    mn = mostLikely / s\n"
        "    bn = maxValue / s\n"
        "    span = bn - an\n"
        "    If span <= 0# Then\n"
        '        detail = "Triangular: the conditioned support has no width"\n'))


def test_08_an_invalid_ordering_is_silently_repaired() -> None:
    _control("test_38", _swap(
        _ORIGINAL,
        "    If minValue > maxValue Then\n"
        '        detail = "Uniform: Min exceeds Max; the ordering is refused, not repaired"\n'
        "        Exit Function\n"
        "    End If\n",
        "    If minValue > maxValue Then\n"
        "        candidate = minValue\n"
        "        minValue = maxValue\n"
        "        maxValue = candidate\n"
        "    End If\n"))


# ===========================================================================
# 9-12. Beta-PERT preparation
# ===========================================================================
def test_09_the_pert_lambda_is_hardcoded_independently() -> None:
    _control("test_45", _swap(
        _ORIGINAL,
        "    alpha0 = 1# + SIM_PERT_LAMBDA * r\n",
        "    alpha0 = 1# + 4# * r\n"))


def test_09a_the_pert_lambda_is_changed() -> None:
    _control("test_24", _swap(
        _ORIGINAL,
        "    beta0 = 1# + SIM_PERT_LAMBDA * (1# - r)\n",
        "    beta0 = 1# + SIM_PERT_SHAPE_UPPER * (1# - r)\n"))


def test_10_the_dispatch_boundary_gives_equality_to_bb() -> None:
    _control("test_50", _swap(
        _ORIGINAL,
        "    If SimSampleMinOf(alpha0, beta0) > SIM_PERT_SHAPE_LOWER Then\n",
        "    If SimSampleMinOf(alpha0, beta0) >= SIM_PERT_SHAPE_LOWER Then\n"))


def test_11_the_bb_orientation_is_reversed() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        ca = SimSampleMinOf(alpha0, beta0)\n"
        "        cb = SimSampleMaxOf(alpha0, beta0)\n",
        "        ca = SimSampleMaxOf(alpha0, beta0)\n"
        "        cb = SimSampleMinOf(alpha0, beta0)\n")
    _control("test_51", damaged)


def test_12_the_bc_orientation_is_reversed() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        ca = SimSampleMaxOf(alpha0, beta0)\n"
        "        cb = SimSampleMinOf(alpha0, beta0)\n",
        "        ca = SimSampleMinOf(alpha0, beta0)\n"
        "        cb = SimSampleMaxOf(alpha0, beta0)\n")
    _control("test_51", damaged)


# ===========================================================================
# 13-20. The Cheng formulation
# ===========================================================================
def test_13_a_locked_bb_literal_is_replaced_by_its_mathematical_equivalent() -> None:
    _control("test_46", _swap(
        _ORIGINAL,
        "        rr = prepared.ChengGamma * v - SIM_CHENG_BB_LITERAL_1\n",
        "        rr = prepared.ChengGamma * v - Log(4#)\n"))


def test_14_a_locked_bc_decimal_is_replaced_by_its_rational_equivalent() -> None:
    _control("test_46", _swap(
        _ORIGINAL,
        "        denominator = ca * candidate.ChengBeta - SIM_CHENG_BC_LITERAL_3\n",
        "        denominator = ca * candidate.ChengBeta - 7# / 9#\n"))


def test_15_the_logit_is_split_into_two_logarithms() -> None:
    damaged = _ORIGINAL.replace(
        "vlog = Log(u1 / (1# - u1))", "vlog = Log(u1) - Log(1# - u1)")
    assert damaged != _ORIGINAL
    _control("test_54", damaged)


def test_16_the_bb_acceptance_operator_is_tightened() -> None:
    _control("test_54", _swap(
        _ORIGINAL,
        "        If ss + SIM_CHENG_BB_LITERAL_2 >= SIM_CHENG_BB_LITERAL_3 * z Then Exit Do\n",
        "        If ss + SIM_CHENG_BB_LITERAL_2 > SIM_CHENG_BB_LITERAL_3 * z Then Exit Do\n"))


def test_16a_the_bb_acceptance_tests_are_reordered() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        If ss + SIM_CHENG_BB_LITERAL_2 >= SIM_CHENG_BB_LITERAL_3 * z Then Exit Do\n"
        "        t = Log(z)\n"
        "        If ss >= t Then Exit Do\n",
        "        t = Log(z)\n"
        "        If ss >= t Then Exit Do\n"
        "        If ss + SIM_CHENG_BB_LITERAL_2 >= SIM_CHENG_BB_LITERAL_3 * z Then Exit Do\n")
    _control("test_54", damaged)


def test_17_the_bc_first_branch_operator_is_loosened() -> None:
    _control("test_55", _swap(
        _ORIGINAL,
        "        If u1 < SIM_CHENG_BC_LITERAL_6 Then\n",
        "        If u1 <= SIM_CHENG_BC_LITERAL_6 Then\n"))


def test_18_the_bc_immediate_acceptance_operator_is_tightened() -> None:
    _control("test_55", _swap(
        _ORIGINAL,
        "            If z <= SIM_CHENG_BC_LITERAL_5 Then\n",
        "            If z < SIM_CHENG_BC_LITERAL_5 Then\n"))


def test_19_the_bc_k2_rejection_operator_is_changed() -> None:
    _control("test_55", _swap(
        _ORIGINAL,
        "            If z >= prepared.ChengK2 Then rejected = True\n",
        "            If z > prepared.ChengK2 Then rejected = True\n"))


def test_20_the_final_bc_acceptance_operator_is_changed() -> None:
    _control("test_55", _swap(
        _ORIGINAL,
        "               - SIM_CHENG_BC_LITERAL_4 >= Log(z) Then Exit Do\n",
        "               - SIM_CHENG_BC_LITERAL_4 > Log(z) Then Exit Do\n"))


# ===========================================================================
# 21-26. Consumption, precomputation and the rescale
# ===========================================================================
def test_21_a_rejected_proposal_rewinds_the_rng_state() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    Dim u1 As Double, u2 As Double\n"
        "    Dim vlog As Double, v As Double, w As Double, z As Double\n"
        "    Dim rr As Double, ss As Double, t As Double\n",
        "    Dim u1 As Double, u2 As Double\n"
        "    Dim vlog As Double, v As Double, w As Double, z As Double\n"
        "    Dim rr As Double, ss As Double, t As Double\n"
        "    Dim saved As SimRngState\n")
    damaged = _swap(
        damaged,
        "    attempts = 0\n    Do\n        attempts = attempts + 1\n"
        "        If Not SimRngNextUniform(working, u1, detail) Then Exit Function\n"
        "        If Not SimRngNextUniform(working, u2, detail) Then Exit Function\n"
        "\n        vlog = Log(u1 / (1# - u1))\n"
        "        v = prepared.ChengBeta * vlog\n",
        "    attempts = 0\n    Do\n        attempts = attempts + 1\n"
        "        saved = working\n"
        "        If Not SimRngNextUniform(working, u1, detail) Then Exit Function\n"
        "        If Not SimRngNextUniform(working, u2, detail) Then Exit Function\n"
        "\n        vlog = Log(u1 / (1# - u1))\n"
        "        v = prepared.ChengBeta * vlog\n")
    damaged = _swap(
        damaged,
        "        If rr + prepared.ChengAlpha * Log(prepared.ChengAlpha / "
        "(prepared.ChengB + w)) >= t Then Exit Do\n    Loop\n",
        "        If rr + prepared.ChengAlpha * Log(prepared.ChengAlpha / "
        "(prepared.ChengB + w)) >= t Then Exit Do\n"
        "        working = saved\n    Loop\n")
    _control("test_53", damaged)


def test_22_a_cheng_attempt_consumes_one_uniform_instead_of_two() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        If Not SimRngNextUniform(working, u1, detail) Then Exit Function\n"
        "        If Not SimRngNextUniform(working, u2, detail) Then Exit Function\n"
        "\n        vlog = Log(u1 / (1# - u1))\n"
        "        v = prepared.ChengBeta * vlog\n"
        "        w = prepared.ChengA * Exp(v)\n"
        "        z = u1 * u1 * u2\n",
        "        If Not SimRngNextUniform(working, u1, detail) Then Exit Function\n"
        "        u2 = u1\n"
        "\n        vlog = Log(u1 / (1# - u1))\n"
        "        v = prepared.ChengBeta * vlog\n"
        "        w = prepared.ChengA * Exp(v)\n"
        "        z = u1 * u1 * u2\n")
    _control("test_53", damaged)


def test_23_the_prepared_constants_are_recomputed_in_the_proposal_loop() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    attempts = 0\n    Do\n        attempts = attempts + 1\n"
        "        If Not SimRngNextUniform(working, u1, detail) Then Exit Function\n"
        "        If Not SimRngNextUniform(working, u2, detail) Then Exit Function\n"
        "\n        vlog = Log(u1 / (1# - u1))\n"
        "        v = prepared.ChengBeta * vlog\n",
        "    attempts = 0\n    Do\n        attempts = attempts + 1\n"
        "        prepared.ChengBeta = Sqr((prepared.ChengAlpha - SIM_CHENG_BB_LITERAL_4) / _\n"
        "                                 (SIM_CHENG_BB_LITERAL_4 * prepared.ChengA * _\n"
        "                                  prepared.ChengB - prepared.ChengAlpha))\n"
        "        If Not SimRngNextUniform(working, u1, detail) Then Exit Function\n"
        "        If Not SimRngNextUniform(working, u2, detail) Then Exit Function\n"
        "\n        vlog = Log(u1 / (1# - u1))\n"
        "        v = prepared.ChengBeta * vlog\n")
    _control("test_52", damaged)


def test_24_the_beta_orientation_return_is_mirrored() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If prepared.FirstParameterIsOrientedA Then\n"
        "        candidate = w / denominator\n"
        "    Else\n"
        "        candidate = prepared.ChengB / denominator\n",
        "    If prepared.FirstParameterIsOrientedA Then\n"
        "        candidate = prepared.ChengB / denominator\n"
        "    Else\n"
        "        candidate = w / denominator\n")
    _control("test_57", damaged)


def test_25_an_unsafe_beta_rescale_is_introduced() -> None:
    _control("test_56", _swap(
        _ORIGINAL,
        "    candidate = (1# - y) * prepared.MinValue + y * prepared.MaxValue\n",
        "    candidate = prepared.MinValue + y * (prepared.MaxValue - prepared.MinValue)\n"))


def test_26_the_beta_variate_is_clipped() -> None:
    _control("test_56", _swap(
        _ORIGINAL,
        "    If Not (y > 0# And y < 1#) Then\n",
        "    If y > 1# Then y = 1#\n"
        "    If y < 0# Then y = 0#\n"
        "    If Not (y > 0# And y < 1#) Then\n"))


# ===========================================================================
# 27-29. Bernoulli
# ===========================================================================
def test_27_the_bernoulli_comparison_is_loosened() -> None:
    _control("test_58", _swap(
        _ORIGINAL,
        "    occurred = (u < probability)\n",
        "    occurred = (u <= probability)\n"))


def test_28_the_probability_extremes_are_special_cased_to_zero_draws() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    working = state\n"
        "    If Not SimRngNextUniform(working, u, detail) Then Exit Function\n"
        "\n    occurred = (u < probability)\n",
        "    If probability = 0# Then\n"
        "        occurred = False\n"
        "        uniform = 0#\n"
        "        uniformsConsumed = 0\n"
        "        SimSampleBernoulli = True\n"
        "        Exit Function\n"
        "    End If\n"
        "    working = state\n"
        "    If Not SimRngNextUniform(working, u, detail) Then Exit Function\n"
        "\n    occurred = (u < probability)\n")
    _control("test_58", damaged)


def test_29_the_probability_is_clamped_instead_of_refused() -> None:
    _control("test_34", _swap(
        _ORIGINAL,
        "    If probability < 0# Or probability > 1# Then\n"
        '        detail = "Bernoulli: Probability is outside [0, 1]; it is refused, not clamped"\n'
        "        Exit Function\n"
        "    End If\n",
        "    If probability < 0# Then probability = 0#\n"
        "    If probability > 1# Then probability = 1#\n"))


# ===========================================================================
# 30-34. Purity, provenance and commit order
# ===========================================================================
def test_30_a_direct_rnd_is_introduced() -> None:
    _control("test_08", _swap(
        _ORIGINAL,
        "    If Not SimRngNextUniform(working, u, detail) Then Exit Function\n"
        "\n    occurred = (u < probability)\n",
        "    u = Rnd()\n"
        "\n    occurred = (u < probability)\n"))


def test_31_a_generator_constant_is_read_directly() -> None:
    _control("test_08", _after(
        _ORIGINAL,
        "    candidate = (1# - u) * minValue + u * maxValue\n",
        "    If u * SIM_RNG_M1 < 0# Then candidate = 0#\n"))


def test_31a_the_algorithm_token_is_introduced() -> None:
    _control("test_09", _after(
        _ORIGINAL,
        "Option Explicit\n",
        "\nPrivate Function MRG32k3aInline() As Double\n"
        "    MRG32k3aInline = 0#\nEnd Function\n"))


def test_32_a_worksheet_reference_is_introduced() -> None:
    _control("test_06", _after(
        _ORIGINAL,
        "Private Function SimSampleScale(ByVal first As Double, ByVal second As Double, _\n"
        "                                ByVal third As Double) As Double\n",
        "    Dim probe As Worksheet\n"))


def test_33_module_level_mutable_sampler_state_is_introduced() -> None:
    _control("test_07", _after(
        _ORIGINAL,
        "Option Explicit\n",
        "\nPrivate mLastSample As Double\n"))


def test_33a_a_static_local_is_introduced() -> None:
    _control("test_07", _swap(
        _ORIGINAL,
        "    Dim working As SimRngState\n    Dim u As Double, candidate As Double\n",
        "    Dim working As SimRngState\n    Static u As Double\n    Dim candidate As Double\n"))


def test_34_the_caller_state_is_committed_before_the_sample_succeeds() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    candidate = (1# - u) * minValue + u * maxValue\n"
        "    If Not IsUsableDouble(candidate) Then\n"
        '        detail = "Uniform: the convex rescale is not representable as a finite Double"\n'
        "        Exit Function\n"
        "    End If\n"
        "\n    sample = candidate\n"
        "    uniformsConsumed = 1\n"
        "    state = working\n",
        "    state = working\n"
        "    candidate = (1# - u) * minValue + u * maxValue\n"
        "    If Not IsUsableDouble(candidate) Then\n"
        '        detail = "Uniform: the convex rescale is not representable as a finite Double"\n'
        "        Exit Function\n"
        "    End If\n"
        "\n    sample = candidate\n"
        "    uniformsConsumed = 1\n")
    _control("test_59", damaged)


# ===========================================================================
# 35-41. The prepared-shape validator
# ===========================================================================
def test_35_the_validator_call_is_removed() -> None:
    _control("test_62", _swap(
        _ORIGINAL,
        "    If Not SimSampleValidatePreparedBetaShape(prepared, detail) Then Exit Function\n",
        "    If Not prepared.Prepared Then\n"
        '        detail = "Beta-PERT: the shape was never prepared"\n'
        "        Exit Function\n"
        "    End If\n"))


def test_36_the_validator_runs_after_the_rng_state_instead_of_before() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If Not SimSampleValidatePreparedBetaShape(prepared, detail) Then Exit Function\n"
        "\n    ' Then the incoming state, before any path can return, degenerate included.\n"
        "    If Not SimRngValidateState(state, detail) Then Exit Function\n",
        "    If Not SimRngValidateState(state, detail) Then Exit Function\n"
        "    If Not SimSampleValidatePreparedBetaShape(prepared, detail) Then Exit Function\n")
    _control("test_62", damaged)


def test_37_the_prepared_raw_order_validation_is_removed() -> None:
    _control("test_62", _swap(
        _ORIGINAL,
        '    If Not SimSampleOrderedTriple(lowValue, modeValue, highValue, "Beta-PERT", detail) '
        "Then Exit Function\n",
        ""))


def test_38_the_degeneracy_consistency_check_is_removed() -> None:
    _control("test_63", _swap(
        _ORIGINAL,
        "    If prepared.Degenerate <> degenerate Then\n"
        '        detail = "Beta-PERT: the degeneracy flag disagrees with the support"\n'
        "        Exit Function\n"
        "    End If\n",
        ""))


def test_38a_the_degeneracy_check_is_narrowed_to_one_direction() -> None:
    """`Degenerate = True` over a live support is the zero-draw forgery."""
    damaged = _swap(
        _ORIGINAL,
        "    If prepared.Degenerate <> degenerate Then\n",
        "    If degenerate And Not prepared.Degenerate Then\n")
    _control("test_63", damaged)


def test_39_the_dispatch_validation_is_removed() -> None:
    _control("test_67", _swap(
        _ORIGINAL,
        "    If prepared.UseChengBB <> isBB Then\n"
        '        detail = "Beta-PERT: the dispatch disagrees with the shape parameters"\n'
        "        Exit Function\n"
        "    End If\n",
        "    isBB = prepared.UseChengBB\n"))


def test_40_the_orientation_validation_is_removed() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If isBB Then\n"
        "        If chengAValue <> lower Or chengBValue <> upper Then\n"
        '            detail = "Beta-PERT: the BB orientation is not min, max"\n'
        "            Exit Function\n"
        "        End If\n"
        "    Else\n"
        "        If chengAValue <> upper Or chengBValue <> lower Then\n"
        '            detail = "Beta-PERT: the BC orientation is not max, min"\n'
        "            Exit Function\n"
        "        End If\n"
        "    End If\n",
        "")
    _control("test_68", damaged)


def test_40a_the_recorded_orientation_check_is_removed() -> None:
    _control("test_69", _swap(
        _ORIGINAL,
        "    If prepared.FirstParameterIsOrientedA <> orientedA Then\n"
        '        detail = "Beta-PERT: the recorded orientation disagrees with the shape"\n'
        "        Exit Function\n"
        "    End If\n",
        ""))


def test_41_the_active_cheng_term_validation_is_removed() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If Not IsUsableDouble(prepared.ChengBeta) Then\n"
        '        detail = "Beta-PERT: the Cheng beta term is not a finite Double"\n'
        "        Exit Function\n"
        "    End If\n"
        "    If prepared.ChengBeta <= 0# Then\n"
        '        detail = "Beta-PERT: the Cheng beta term is not positive"\n'
        "        Exit Function\n"
        "    End If\n",
        "")
    _control("test_70", damaged)


def test_41a_the_inactive_term_defaults_are_no_longer_pinned() -> None:
    damaged = _ORIGINAL.replace(
        "        If Not SimSampleTermsAreUnset(prepared.ChengDelta, prepared.ChengK1, _\n"
        "                                      prepared.ChengK2, detail) Then Exit Function\n", "")
    damaged = damaged.replace(
        "        If Not SimSampleTermsAreUnset(prepared.ChengGamma, 0#, 0#, detail) "
        "Then Exit Function\n", "")
    assert damaged != _ORIGINAL
    _control("test_71", damaged)


def test_41b_the_degenerate_record_field_pinning_is_removed() -> None:
    damaged = _ORIGINAL.replace(
        "        If prepared.Alpha <> 0# Or prepared.Beta <> 0# Then\n"
        '            detail = "Beta-PERT: a degenerate shape carries a parameterisation"\n'
        "            Exit Function\n"
        "        End If\n"
        "        If prepared.UseChengBB Or prepared.FirstParameterIsOrientedA Then\n"
        '            detail = "Beta-PERT: a degenerate shape carries a dispatch or an orientation"\n'
        "            Exit Function\n"
        "        End If\n"
        "        If Not SimSampleTermsAreUnset(prepared.ChengA, prepared.ChengB, _\n"
        "                                      prepared.ChengAlpha, detail) Then Exit Function\n"
        "        If Not SimSampleTermsAreUnset(prepared.ChengBeta, prepared.ChengGamma, _\n"
        "                                      prepared.ChengDelta, detail) Then Exit Function\n"
        "        If Not SimSampleTermsAreUnset(prepared.ChengK1, prepared.ChengK2, 0#, detail) "
        "Then Exit Function\n", "")
    assert damaged != _ORIGINAL
    _control("test_65", damaged)


def test_41c_the_shape_family_bound_is_removed() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If Not SimSampleShapeInFamily(alphaValue, detail) Then Exit Function\n"
        "    If Not SimSampleShapeInFamily(betaValue, detail) Then Exit Function\n",
        "")
    _control("test_66", damaged)



if __name__ == "__main__":  # pragma: no cover
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
