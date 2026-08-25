#!/usr/bin/env python3
"""PCCM Phase 6 Step-6 MUTATION CONTROLS for the modSimRng conformance battery.

A conformance test that cannot fail proves nothing. Every control here damages
the accepted source (or the accepted D6-11 authority), reruns the WHOLE Step-6
conformance battery against the damaged copy, and requires at least one real
detector to refuse it. The names of the refusing detectors are asserted too, so
a control cannot quietly degrade into "something, somewhere, went red".

Nothing in this file writes to the repository: the damaged copies live in a
temporary directory and the conformance module is pointed at them for the
duration of one control.

Runs standalone or under pytest.
"""

from __future__ import annotations

import re
import shutil
import signal
import sys
import tempfile
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import test_phase6_sim_rng_vba as conformance  # noqa: E402

_ORIGINAL_BAS = conformance.SIM_RNG_BAS.read_text(encoding="utf-8")
_TEST_BUDGET_SECONDS = 5


class _Timeout(Exception):
    pass


def _conformance_tests() -> list[str]:
    names = sorted(n for n in dir(conformance) if n.startswith("test_"))
    assert len(names) >= 45, names
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


def _install(source: str | None = None, spec: Path | None = None,
             vba_dir: Path | None = None):
    """Point the conformance module at damaged inputs. Returns a restore callable."""
    saved = (conformance.SIM_RNG_BAS, conformance.SPEC, conformance.SRC_VBA,
             dict(conformance._CACHE))
    conformance._CACHE.clear()
    if source is not None:
        assert source != _ORIGINAL_BAS, "the mutation changed nothing"
        temp = Path(tempfile.mkdtemp(prefix="pccm-step6-mutation-"))
        target = temp / "modSimRng.bas"
        target.write_text(source, encoding="utf-8")
        conformance.SIM_RNG_BAS = target
    if spec is not None:
        conformance.SPEC = spec
    if vba_dir is not None:
        conformance.SRC_VBA = vba_dir

    def restore() -> None:
        (conformance.SIM_RNG_BAS, conformance.SPEC, conformance.SRC_VBA,
         restored) = saved
        conformance._CACHE.clear()
        conformance._CACHE.update(restored)

    return restore


def _control(expected: str, *, source: str | None = None, spec: Path | None = None,
             vba_dir: Path | None = None) -> None:
    """Assert the battery refuses the damage, and that `expected` is among the refusers."""
    restore = _install(source=source, spec=spec, vba_dir=vba_dir)
    try:
        refused = _run_battery()
    finally:
        restore()
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def _swap(text: str, old: str, new: str, count: int = 1) -> str:
    assert text.count(old) == count, (old, text.count(old))
    return text.replace(old, new)


def _drop(text: str, pattern: str) -> str:
    damaged, hits = re.subn(pattern, "", text, flags=re.M | re.S)
    assert hits == 1, (pattern, hits)
    return damaged


def _spec_copy(mutate) -> Path:
    temp = Path(tempfile.mkdtemp(prefix="pccm-step6-spec-"))
    shutil.copytree(PCCM_ROOT / "spec", temp / "spec")
    contract = temp / "spec" / "structure_contract.yaml"
    original = contract.read_text(encoding="utf-8")
    damaged = mutate(original)
    assert damaged != original, "the authority mutation changed nothing"
    contract.write_text(damaged, encoding="utf-8")
    return temp / "spec"


# ===========================================================================
# The battery itself must pass on the ACCEPTED source, or every control below
# is measuring noise.
# ===========================================================================
def test_00_the_accepted_source_passes_every_detector() -> None:
    restore = _install()
    try:
        refused = _run_battery()
    finally:
        restore()
    assert refused == [], refused


# ===========================================================================
# 1-3. the recurrence constants
# ===========================================================================
def test_01_a12_changed() -> None:
    _control("test_37", source=_swap(
        _ORIGINAL_BAS,
        "signed = SIM_RNG_A12 * state.S11",
        "signed = SIM_RNG_A21 * state.S11"))


def test_02_a13n_changed() -> None:
    _control("test_37", source=_swap(
        _ORIGINAL_BAS,
        "- SIM_RNG_A13N * state.S10",
        "- SIM_RNG_A23N * state.S10"))


def test_03_m1_and_m2_swapped() -> None:
    damaged = _swap(_ORIGINAL_BAS,
                    "p1 = SimRngReduce(signed, SIM_RNG_M1)",
                    "p1 = SimRngReduce(signed, SIM_RNG_M2)")
    damaged = _swap(damaged,
                    "p2 = SimRngReduce(signed, SIM_RNG_M2)",
                    "p2 = SimRngReduce(signed, SIM_RNG_M1)")
    _control("test_37", source=damaged)


# ===========================================================================
# 4-6. the combination and the reduction
# ===========================================================================
def test_04_the_combination_boundary_narrowed() -> None:
    _control("test_38", source=_swap(_ORIGINAL_BAS, "If p1 <= p2 Then", "If p1 < p2 Then"))


def test_05_the_fix_reduction_replaced_by_vba_mod() -> None:
    damaged = _swap(_ORIGINAL_BAS,
                    "    k = Fix(p / m)\n    r = p - k * m\n",
                    "    k = 0#\n    r = p Mod m\n")
    _control("test_34", source=damaged)


def test_06_the_negative_remainder_correction_deleted() -> None:
    _control("test_35", source=_swap(
        _ORIGINAL_BAS, "    If r < 0# Then r = r + m\n", ""))


def test_06a_the_multiply_negative_correction_deleted() -> None:
    _control("test_35", source=_swap(
        _ORIGINAL_BAS, "    If v < 0# Then v = v + m\n", ""))


# ===========================================================================
# 7-9. seeding
# ===========================================================================
def test_07_a_seed_mixer_introduced() -> None:
    _control("test_14", source=_swap(
        _ORIGINAL_BAS,
        "    candidate.S12 = CDbl(seed)\n",
        "    candidate.S12 = CDbl(seed) + 1#\n"))


def test_08_auto_changed_from_modular_power_to_linear_stepping() -> None:
    damaged = _swap(
        _ORIGINAL_BAS,
        "        half = Fix(remaining / 2#)\n"
        "        If remaining - half * 2# = 1# Then\n"
        "            result = SimRngMultModM(result, baseValue, 0#, CDbl(SIM_AUTO_MODULUS))\n"
        "        End If\n"
        "        baseValue = SimRngMultModM(baseValue, baseValue, 0#, CDbl(SIM_AUTO_MODULUS))\n"
        "        remaining = half\n",
        "        half = 0#\n"
        "        result = SimRngMultModM(result, baseValue, 0#, CDbl(SIM_AUTO_MODULUS))\n"
        "        remaining = remaining - 1#\n")
    _control("test_46", source=damaged)


def test_09_unsafe_multiplication_in_the_modular_exponent() -> None:
    damaged = _swap(
        _ORIGINAL_BAS,
        "            result = SimRngMultModM(result, baseValue, 0#, CDbl(SIM_AUTO_MODULUS))",
        "            result = SimRngReduce(result * baseValue, CDbl(SIM_AUTO_MODULUS))")
    damaged = _swap(
        damaged,
        "        baseValue = SimRngMultModM(baseValue, baseValue, 0#, CDbl(SIM_AUTO_MODULUS))",
        "        baseValue = SimRngReduce(baseValue * baseValue, CDbl(SIM_AUTO_MODULUS))")
    _control("test_16", source=damaged)


# ===========================================================================
# 10-14. the jump
# ===========================================================================
def test_10_the_jump_safe_multiply_replaced_by_a_naive_product() -> None:
    damaged = _swap(
        _ORIGINAL_BAS,
        "    acc = SimRngMultModM(c0, vector(0), 0#, m)\n"
        "    acc = SimRngMultModM(c1, vector(1), acc, m)\n"
        "    acc = SimRngMultModM(c2, vector(2), acc, m)\n",
        "    acc = SimRngReduce(c0 * vector(0) + c1 * vector(1) + c2 * vector(2), m)\n")
    _control("test_39", source=damaged)


def test_11_the_input_reversal_removed() -> None:
    damaged = _swap(_ORIGINAL_BAS, "    inFirst(0) = state.S12", "    inFirst(0) = state.S10")
    damaged = _swap(damaged, "    inFirst(2) = state.S10", "    inFirst(2) = state.S12")
    _control("test_39", source=damaged)


def test_12_the_output_reversal_removed() -> None:
    damaged = _swap(_ORIGINAL_BAS, "    candidate.S10 = outFirst(2)",
                    "    candidate.S10 = outFirst(0)")
    damaged = _swap(damaged, "    candidate.S12 = outFirst(0)",
                    "    candidate.S12 = outFirst(2)")
    _control("test_39", source=damaged)


def test_13_a_matrix_element_transposed() -> None:
    damaged = _swap(_ORIGINAL_BAS, "SIM_JUMP_A1_R1C2", "SIM_JUMP_A1_R2C1")
    _control("test_39", source=damaged)


def test_14_a_precomputed_stream_table_introduced() -> None:
    damaged = _swap(
        _ORIGINAL_BAS,
        "    current = baseState\n    For index = 1 To k\n",
        "    current = baseState\n    If k > 400 Then k = 400\n    For index = 1 To k\n")
    _control("test_21", source=damaged)


# ===========================================================================
# 15-18. component assignment
# ===========================================================================
_SORT_BLOCK = (
    r"    For index = 1 To count - 1\n"
    r"        moving = order\(index\)\n.*?"
    r"        order\(probe \+ 1\) = moving\n"
    r"    Next index\n"
)


def test_15_the_physical_input_order_used_directly() -> None:
    _control("test_25", source=_drop(_ORIGINAL_BAS, _SORT_BLOCK))


def test_16_a_numeric_permanent_id_suffix_sort_introduced() -> None:
    damaged = _swap(
        _ORIGINAL_BAS,
        "            If StrComp(ids(LBound(ids) + order(probe)), ids(LBound(ids) + moving), _\n"
        "                       vbBinaryCompare) <= 0 Then Exit Do\n",
        "            If Val(Mid(ids(LBound(ids) + order(probe)), 4)) <= _\n"
        "               Val(Mid(ids(LBound(ids) + moving), 4)) Then Exit Do\n")
    _control("test_41", source=damaged)


def test_17_risk_occurrence_and_severity_not_interleaved() -> None:
    damaged = _swap(
        _ORIGINAL_BAS,
        "    For index = 0 To riskCount - 1\n"
        "        built(slot).DriverKind = SIM_COMPONENT_2_DRIVER_KIND\n"
        "        built(slot).PermanentId = riskIds(LBound(riskIds) + riskOrder(index))\n"
        "        built(slot).Role = SIM_COMPONENT_2_ROLE\n"
        "        slot = slot + 1\n"
        "        built(slot).DriverKind = SIM_COMPONENT_3_DRIVER_KIND\n"
        "        built(slot).PermanentId = riskIds(LBound(riskIds) + riskOrder(index))\n"
        "        built(slot).Role = SIM_COMPONENT_3_ROLE\n"
        "        slot = slot + 1\n"
        "    Next index\n",
        "    For index = 0 To riskCount - 1\n"
        "        built(slot).DriverKind = SIM_COMPONENT_2_DRIVER_KIND\n"
        "        built(slot).PermanentId = riskIds(LBound(riskIds) + riskOrder(index))\n"
        "        built(slot).Role = SIM_COMPONENT_2_ROLE\n"
        "        slot = slot + 1\n"
        "    Next index\n"
        "    For index = 0 To riskCount - 1\n"
        "        built(slot).DriverKind = SIM_COMPONENT_3_DRIVER_KIND\n"
        "        built(slot).PermanentId = riskIds(LBound(riskIds) + riskOrder(index))\n"
        "        built(slot).Role = SIM_COMPONENT_3_ROLE\n"
        "        slot = slot + 1\n"
        "    Next index\n")
    _control("test_24", source=damaged)


def test_18_a_duplicate_component_silently_accepted() -> None:
    damaged = _drop(
        _ORIGINAL_BAS,
        r"    For index = 1 To count - 1\n"
        r"        If StrComp\(ids\(LBound\(ids\) \+ order\(index - 1\)\).*?"
        r"        End If\n    Next index\n")
    _control("test_28", source=damaged)


# ===========================================================================
# 19-21. purity
# ===========================================================================
def test_19_a_worksheet_reference_introduced() -> None:
    _control("test_05", source=_swap(
        _ORIGINAL_BAS,
        "Private Function SimRngNorm() As Double\n",
        "Private Function SimRngNorm() As Double\n    Dim probe As Worksheet\n"))


def test_20_global_mutable_generator_state_introduced() -> None:
    _control("test_06", source=_swap(
        _ORIGINAL_BAS,
        "Private Function SimRngNorm() As Double\n",
        "Private mGeneratorState As SimRngState\n\n"
        "Private Function SimRngNorm() As Double\n"))


def test_20a_a_static_local_introduced() -> None:
    _control("test_06", source=_swap(
        _ORIGINAL_BAS,
        "    Dim k As Double, r As Double\n",
        "    Static k As Double\n    Dim r As Double\n"))


def test_21_rnd_introduced() -> None:
    _control("test_05", source=_swap(
        _ORIGINAL_BAS,
        "        candidate = (p1 - p2) * SimRngNorm()",
        "        candidate = Rnd()"))


# ===========================================================================
# 22-24. the D6-11 authority itself
# ===========================================================================
def test_22_the_algorithm_token_placed_in_another_module() -> None:
    temp = Path(tempfile.mkdtemp(prefix="pccm-step6-vba-"))
    for module in sorted((PCCM_ROOT / "src" / "vba").glob("*.bas")):
        shutil.copy(module, temp / module.name)
    smuggled = temp / "modCalcFactors.bas"
    assert smuggled.exists(), "the module the token was smuggled into is gone"
    smuggled.write_text(
        smuggled.read_text(encoding="utf-8").replace(
            "Option Explicit",
            "Option Explicit\n\nPrivate Function MRG32k3aSmuggled() As Double\n"
            "    MRG32k3aSmuggled = 0#\nEnd Function", 1),
        encoding="utf-8")
    _control("test_09", vba_dir=temp)


def test_23_the_allowed_in_grant_widened() -> None:
    def widen(text: str) -> str:
        return _swap(text,
                     '      allowed_in:\n        - "modSimRng"\n',
                     '      allowed_in:\n        - "modSimRng"\n        - "modCalcFactors"\n')

    _control("test_10", spec=_spec_copy(widen))


def test_24_run_simulation_scoped_prematurely() -> None:
    def scope(text: str) -> str:
        return _swap(text,
                     '    - "RunSimulation"\n',
                     '    - construct: "RunSimulation"\n'
                     '      allowed_in:\n        - "modSimRng"\n')

    _control("test_10", spec=_spec_copy(scope))


if __name__ == "__main__":  # pragma: no cover
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
