#!/usr/bin/env python3
"""PCCM Phase 6 Step-8 MUTATION CONTROLS for the modSimEngine conformance battery.

A conformance test that cannot fail proves nothing. Every control here damages
the accepted source, reruns the WHOLE Step-8 conformance battery against the
damaged copy, and requires a NAMED detector among the refusers - so a control
cannot quietly degrade into "something, somewhere, went red".

Nothing here writes to the repository: the damaged copy lives in a temporary
directory and the conformance module is pointed at it for one control.

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

import test_phase6_sim_engine_vba as conformance  # noqa: E402

_ORIGINAL = conformance.SIM_ENGINE_BAS.read_text(encoding="utf-8")
_TEST_BUDGET_SECONDS = 30


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


def _install(source: str | None):
    saved = (conformance.SIM_ENGINE_BAS, dict(conformance._CACHE))
    conformance._CACHE.clear()
    if source is not None:
        assert source != _ORIGINAL, "the mutation changed nothing"
        temp = Path(tempfile.mkdtemp(prefix="pccm-step8-mutation-"))
        target = temp / "modSimEngine.bas"
        target.write_text(source, encoding="utf-8")
        conformance.SIM_ENGINE_BAS = target

    def restore() -> None:
        conformance.SIM_ENGINE_BAS = saved[0]
        conformance._CACHE.clear()
        conformance._CACHE.update(saved[1])

    return restore


def _control(expected: str, source: str) -> None:
    restore = _install(source)
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
# The battery must pass on the ACCEPTED source.
# ===========================================================================
def test_00_the_accepted_source_passes_every_detector() -> None:
    restore = _install(None)
    try:
        refused = _run_battery()
    finally:
        restore()
    assert refused == [], refused


# ===========================================================================
# 1-6. Order, identity and stream mapping
# ===========================================================================
_CANONICAL_COST_LOOP = """    For index = 0 To costCount - 1
        If Not SimEngineClaim(drivers, driverCount, components(index), False, _
                              SIM_COMPONENT_1_DRIVER_KIND, SIM_COMPONENT_1_ROLE, index, _
                              claimed, source, detail) Then Exit Function
        If Not SimEngineAdopt(drivers(LBound(drivers) + source), prepared(index), detail) Then Exit Function
"""


def test_01_the_physical_input_order_is_used_as_execution_order() -> None:
    """Adopt each driver where the caller put it, ignoring the canonical order."""
    damaged = _swap(
        _ORIGINAL, _CANONICAL_COST_LOOP,
        """    For index = 0 To costCount - 1
        If Not SimEngineClaim(drivers, driverCount, components(index), False, _
                              SIM_COMPONENT_1_DRIVER_KIND, SIM_COMPONENT_1_ROLE, index, _
                              claimed, source, detail) Then Exit Function
        If Not SimEngineAdopt(drivers(LBound(drivers) + index), prepared(index), detail) Then Exit Function
""")
    _control("test_27", damaged)


def test_02_risks_are_accumulated_before_cost_lines() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        For index = 0 To costCount - 1\n            ' THE SAMPLE IS UNIT COST.",
        "        For index = costCount To driverCount - 1\n            ' THE SAMPLE IS UNIT COST.")
    damaged = _swap(
        damaged,
        "        For index = costCount To driverCount - 1\n"
        "            ' D6-18b. THE OCCURRENCE DRAW COMES FIRST",
        "        For index = 0 To costCount - 1\n"
        "            ' D6-18b. THE OCCURRENCE DRAW COMES FIRST")
    _control("test_18", damaged)


def test_03_a_numeric_permanent_id_suffix_sort_is_introduced() -> None:
    damaged = _swap(
        _ORIGINAL,
        "            If StrComp(drivers(LBound(drivers) + index).PermanentId, component.PermanentId, _\n"
        "                       vbBinaryCompare) = 0 Then\n",
        "            If Val(Mid(drivers(LBound(drivers) + index).PermanentId, 4)) = _\n"
        "               Val(Mid(component.PermanentId, 4)) Then\n")
    _control("test_28", damaged)


def test_04_the_component_mapping_trusts_a_mismatched_identity() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If found < 0 Then\n"
        '        detail = "engine: a component identity matches no driver record"\n'
        "        Exit Function\n"
        "    End If\n",
        "    If found < 0 Then\n"
        "        found = 0\n"
        "    End If\n")
    _control("test_45", damaged)


def test_05_the_occurrence_and_severity_roles_are_no_longer_checked() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If component.Role <> wantRole Then\n"
        '        detail = "engine: a component has the wrong role"\n'
        "        Exit Function\n"
        "    End If\n",
        "")
    damaged = _swap(
        damaged,
        "        If components(severity).Role <> SIM_COMPONENT_3_ROLE Then\n"
        '            detail = "engine: a risk severity component has the wrong role"\n'
        "            Exit Function\n"
        "        End If\n",
        "")
    _control("test_46", damaged)


def test_06_a_duplicate_or_aliased_stream_is_accepted() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If claimed(found) Then\n"
        '        detail = "engine: one driver record was claimed by two components"\n'
        "        Exit Function\n"
        "    End If\n",
        "")
    damaged = _swap(
        damaged,
        "        If components(severity).StreamIndex = components(occurrence).StreamIndex Then\n"
        '            detail = "engine: a risk shares one stream between occurrence and severity"\n'
        "            Exit Function\n"
        "        End If\n",
        "")
    damaged = _swap(
        damaged,
        "    If component.StreamIndex <> SIM_STREAM_INDEX_ORIGIN + wantIndex Then\n"
        '        detail = "engine: a component has the wrong stream index"\n'
        "        Exit Function\n"
        "    End If\n",
        "")
    _control("test_42", damaged)


# ===========================================================================
# 7-8. The Step-7 prepared-shape carry-forward
# ===========================================================================
_PREPARE_BETA = """            If Not SimSamplePrepareBetaPert(prepared(index).MinValue, prepared(index).MostLikely, _
                                            prepared(index).MaxValue, prepared(index).BetaShape, _
                                            detail) Then Exit Function
"""


def test_07_a_beta_shape_is_assembled_by_hand() -> None:
    damaged = _swap(
        _ORIGINAL, _PREPARE_BETA,
        """            prepared(index).BetaShape.MinValue = prepared(index).MinValue
            prepared(index).BetaShape.MostLikely = prepared(index).MostLikely
            prepared(index).BetaShape.MaxValue = prepared(index).MaxValue
            prepared(index).BetaShape.Alpha = SIM_PERT_SHAPE_LOWER
            prepared(index).BetaShape.Beta = SIM_PERT_SHAPE_UPPER
            prepared(index).BetaShape.Prepared = True
""")
    _control("test_39", damaged)


def test_08_a_beta_shape_field_is_mutated_after_preparation() -> None:
    damaged = _after(
        _ORIGINAL, _PREPARE_BETA,
        "            prepared(index).BetaShape.Alpha = prepared(index).BetaShape.Beta\n")
    _control("test_39", damaged)


# ===========================================================================
# 9-11. Factor semantics
# ===========================================================================
def test_09_a_uniforms_most_likely_is_read() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If factor.DistKind <> DIST_UNIFORM Then\n"
        "        target.MostLikely = factor.MostLikely\n"
        "    End If\n",
        "    target.MostLikely = factor.MostLikely\n")
    damaged = _swap(
        damaged,
        "        ok = SimSampleUniform(state, prepared.MinValue, prepared.MaxValue, _\n"
        "                              sample, consumed, detail)\n",
        "        ok = SimSampleTriangular(state, prepared.MinValue, prepared.MostLikely, _\n"
        "                                 prepared.MaxValue, sample, consumed, detail)\n")
    _control("test_31", damaged)


def test_10_a_cost_lines_probability_is_read() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If factor.IsRisk Then\n"
        "        target.Probability = factor.Probability\n"
        "    Else\n"
        "        target.Quantity = factor.Quantity\n"
        "    End If\n",
        "    target.Probability = factor.Probability\n"
        "    target.Quantity = factor.Quantity\n")
    damaged = _swap(
        damaged,
        "        factors(1) = prepared.Quantity\n",
        "        factors(1) = prepared.Quantity * (1# + prepared.Probability)\n")
    _control("test_29", damaged)


def test_11_a_risks_quantity_is_read() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If factor.IsRisk Then\n"
        "        target.Probability = factor.Probability\n"
        "    Else\n"
        "        target.Quantity = factor.Quantity\n"
        "    End If\n",
        "    target.Probability = factor.Probability\n"
        "    target.Quantity = factor.Quantity\n")
    damaged = _swap(
        damaged,
        "        factors(0) = sample\n        factors(1) = factor\n        count = 2\n",
        "        factors(0) = sample * prepared.Quantity\n        factors(1) = factor\n"
        "        count = 2\n")
    _control("test_30", damaged)


# ===========================================================================
# 12-18. Contribution arithmetic
# ===========================================================================
def test_12_quantity_is_omitted() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        factors(1) = prepared.Quantity\n",
        "        factors(1) = 1#\n")
    _control("test_20", damaged)


def test_13_quantity_is_applied_twice() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        factors(0) = sample\n        factors(1) = prepared.Quantity\n",
        "        factors(0) = sample * prepared.Quantity\n"
        "        factors(1) = prepared.Quantity\n")
    _control("test_20", damaged)


def test_14_the_total_cost_is_sampled_instead_of_the_unit_cost() -> None:
    """The distribution support becomes total cost, so Quantity vanishes."""
    damaged = _swap(
        _ORIGINAL,
        "    target.MinValue = factor.MinValue\n"
        "    target.MaxValue = factor.MaxValue\n",
        "    target.MinValue = factor.MinValue * factor.Quantity\n"
        "    target.MaxValue = factor.MaxValue * factor.Quantity\n")
    damaged = _swap(
        damaged,
        "        factors(1) = prepared.Quantity\n",
        "        factors(1) = 1#\n")
    _control("test_20", damaged)


def test_15_the_cost_nominal_term_uses_kpv() -> None:
    damaged = _swap(
        _ORIGINAL,
        "                                         prepared(index).Knom, SIM_MEASURE_NOMINAL, _\n",
        "                                         prepared(index).Kpv, SIM_MEASURE_NOMINAL, _\n", 2)
    _control("test_33", damaged)


def test_16_the_cost_pv_term_is_derived_from_the_nominal_one() -> None:
    damaged = _swap(
        _ORIGINAL,
        "            If Not SimEngineContribution(prepared(index), unitCost, True, _\n"
        "                                         prepared(index).Kpv, SIM_MEASURE_PV, _\n"
        "                                         iteration, term, detail) Then Exit Function\n"
        "            pvTerm(index) = term\n",
        "            pvTerm(index) = nominalTerm(index)\n")
    _control("test_33", damaged)


def test_17_a_risk_contribution_gains_a_quantity_factor() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        factors(1) = factor\n        count = 2\n",
        "        factors(1) = factor\n        factors(2) = 2#\n        count = 3\n")
    _control("test_52", damaged)


def test_18_probability_is_folded_into_the_contribution() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        factors(0) = sample\n        factors(1) = factor\n",
        "        factors(0) = sample * prepared.Probability\n        factors(1) = factor\n")
    _control("test_21", damaged)


# ===========================================================================
# 19-22. D6-18b
# ===========================================================================
_SEVERITY_CALL = """            If Not SimEngineSampleValue(prepared(index), valueState(index), severity, _
                                        iteration, detail) Then Exit Function
"""


def test_19_the_severity_sampler_is_skipped_when_the_risk_did_not_occur() -> None:
    damaged = _swap(
        _ORIGINAL, _SEVERITY_CALL,
        """            If occurred Then
                If Not SimEngineSampleValue(prepared(index), valueState(index), severity, _
                                            iteration, detail) Then Exit Function
            End If
""")
    _control("test_50", damaged)


def test_20_the_severity_sampler_is_skipped_at_probability_zero() -> None:
    damaged = _swap(
        _ORIGINAL, _SEVERITY_CALL,
        """            If prepared(index).Probability > 0# Then
                If Not SimEngineSampleValue(prepared(index), valueState(index), severity, _
                                            iteration, detail) Then Exit Function
            End If
""")
    _control("test_50", damaged)


_BERNOULLI_CALL = """            If Not SimSampleBernoulli(occurrenceState(index), prepared(index).Probability, _
                                      occurred, drawn, consumed, detail) Then
                detail = "engine: iteration " & CStr(iteration) & ", risk " & _
                         prepared(index).PermanentId & ": " & detail
                Exit Function
            End If
"""


def test_21_the_occurrence_draw_is_taken_after_the_severity_sample() -> None:
    """Unobservable in the numbers - the two draws are on different streams -
    so only a source detector can see it, which is why one exists."""
    damaged = _swap(_ORIGINAL, _BERNOULLI_CALL, "\x00MOVED\x00")
    damaged = _swap(
        damaged,
        "            If Not SimEngineSampleValue(prepared(index), valueState(index), severity, _\n"
        "                                        iteration, detail) Then Exit Function\n",
        "            If Not SimEngineSampleValue(prepared(index), valueState(index), severity, _\n"
        "                                        iteration, detail) Then Exit Function\n"
        + _BERNOULLI_CALL)
    damaged = damaged.replace("\x00MOVED\x00", "")
    _control("test_50", damaged)


def test_22_the_engine_implements_the_occurrence_comparison_itself() -> None:
    damaged = _swap(
        _ORIGINAL,
        """            If Not SimSampleBernoulli(occurrenceState(index), prepared(index).Probability, _
                                      occurred, drawn, consumed, detail) Then
""",
        """            occurred = (drawn < prepared(index).Probability)
            If Not SimSampleBernoulli(occurrenceState(index), prepared(index).Probability, _
                                      occurred, drawn, consumed, detail) Then
""")
    _control("test_09", damaged)


# ===========================================================================
# 23-26. Accumulation
# ===========================================================================
def test_23_safe_product_is_replaced_by_chained_multiplication() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If Not SafeProduct(factors, count, term) Then\n",
        "    term = factors(0) * factors(1)\n    If False Then\n")
    _control("test_34", damaged)


def test_24_safe_signed_sum_is_replaced_by_a_running_addition() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        If Not SafeSignedSum(nominalTerm, driverCount, measured) Then\n"
        '            detail = "engine: iteration " & CStr(iteration) & ": the nominal total is not representable"\n'
        "            Exit Function\n"
        "        End If\n"
        "        stagedNominal(iteration - 1) = measured\n",
        "        measured = 0#\n"
        "        For index = 0 To driverCount - 1\n"
        "            measured = measured + nominalTerm(index)\n"
        "        Next index\n"
        "        stagedNominal(iteration - 1) = measured\n")
    _control("test_33", damaged)


def test_25_nominal_and_pv_share_one_accumulator() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        stagedPv(iteration - 1) = measured\n",
        "        stagedPv(iteration - 1) = stagedNominal(iteration - 1)\n")
    _control("test_33", damaged)


def test_26_the_accumulation_order_is_reversed() -> None:
    """Only detectable on the constructed non-associative fixture."""
    damaged = _swap(
        _ORIGINAL,
        "            nominalTerm(index) = term\n"
        "            ' PV IS AN INDEPENDENT ACCUMULATOR.",
        "            nominalTerm(costCount - 1 - index) = term\n"
        "            ' PV IS AN INDEPENDENT ACCUMULATOR.")
    damaged = _swap(
        damaged,
        "            pvTerm(index) = term\n"
        "        Next index\n\n"
        "        For index = costCount To driverCount - 1",
        "            pvTerm(costCount - 1 - index) = term\n"
        "        Next index\n\n"
        "        For index = costCount To driverCount - 1")
    _control("test_35", damaged)


# ===========================================================================
# 27-29. The hot loop
# ===========================================================================
def test_27_a_redim_is_introduced_inside_the_iteration_loop() -> None:
    # ANCHORED ON THE CANONICAL LOOP. P7-3's replay opens with the same line, so
    # the bare `For iteration` no longer names one place - and a ReDim planted
    # in the replay would not be the defect this control is about.
    damaged = _swap(
        _ORIGINAL,
        "    For iteration = 1 To iterations\n        For index = 0 To costCount - 1\n",
        "    For iteration = 1 To iterations\n        ReDim nominalTerm(0 To 2)\n"
        "        For index = 0 To costCount - 1\n")
    _control("test_38", damaged)


def test_28_beta_preparation_is_introduced_inside_the_loop() -> None:
    damaged = _swap(
        _ORIGINAL,
        "            If Not SimEngineSampleValue(prepared(index), valueState(index), unitCost, _\n"
        "                                        iteration, detail) Then Exit Function\n",
        "            If prepared(index).DistKind = DIST_BETA_PERT Then\n"
        "                If Not SimSamplePrepareBetaPert(prepared(index).MinValue, _\n"
        "                                                prepared(index).MostLikely, _\n"
        "                                                prepared(index).MaxValue, _\n"
        "                                                prepared(index).BetaShape, detail) Then Exit Function\n"
        "            End If\n"
        "            If Not SimEngineSampleValue(prepared(index), valueState(index), unitCost, _\n"
        "                                        iteration, detail) Then Exit Function\n")
    _control("test_38", damaged)


def test_29_stream_construction_is_introduced_inside_the_loop() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        If Not SafeSignedSum(nominalTerm, driverCount, measured) Then\n",
        "        If Not SimRngBuildComponentStreams(costIds, costCount, riskIds, riskCount, _\n"
        "                                           baseState, components, detail) Then Exit Function\n"
        "        If Not SafeSignedSum(nominalTerm, driverCount, measured) Then\n")
    _control("test_38", damaged)


# ===========================================================================
# 30-33. Purity and provenance
# ===========================================================================
def test_30_a_direct_generator_draw_is_introduced() -> None:
    damaged = _swap(
        _ORIGINAL,
        "            If Not SimEngineSampleValue(prepared(index), valueState(index), unitCost, _\n"
        "                                        iteration, detail) Then Exit Function\n",
        "            If Not SimRngNextUniform(valueState(index), unitCost, detail) Then Exit Function\n")
    _control("test_08", damaged)


def test_31_a_generator_constant_is_read_directly() -> None:
    damaged = _after(
        _ORIGINAL,
        "        factors(2) = factor\n",
        "        If sample > SIM_RNG_M1 Then factors(2) = 0#\n")
    _control("test_08", damaged)


def test_31a_the_algorithm_token_is_introduced() -> None:
    damaged = _after(
        _ORIGINAL, "Option Explicit\n",
        "\nPrivate Function MRG32k3aInline() As Double\n"
        "    MRG32k3aInline = 0#\nEnd Function\n")
    _control("test_08", damaged)


def test_32_a_worksheet_reference_is_introduced() -> None:
    # ANCHORED ON THE ENTRY POINT'S OWN DECLARATIONS. P7-3 gave the replay the
    # same local, so the bare declaration no longer names one place.
    damaged = _after(
        _ORIGINAL,
        "    Dim prepared() As SimEngineDriver\n    Dim valueState() As SimRngState",
        "\n    Dim probe As Worksheet")
    _control("test_06", damaged)


def test_33_module_level_mutable_state_is_introduced() -> None:
    _control("test_07", _after(
        _ORIGINAL, "Option Explicit\n", "\nPrivate mLastTotal As Double\n"))


def test_33a_a_static_local_is_introduced() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    Dim iteration As Long, index As Long\n",
        "    Static iteration As Long\n    Dim index As Long\n")
    _control("test_07", damaged)


# ===========================================================================
# 34-38. Output, the empty model, and premature scope
# ===========================================================================
def test_34_a_partial_output_is_committed_during_the_loop() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        stagedPv(iteration - 1) = measured\n"
        "    Next iteration\n",
        "        stagedPv(iteration - 1) = measured\n"
        "        totalNominal = stagedNominal\n"
        "        totalPv = stagedPv\n"
        "    Next iteration\n")
    _control("test_47", damaged)


def test_35_a_zero_driver_model_is_refused() -> None:
    for invented in (
        '    If driverCount = 0 Then\n'
        '        detail = "engine: the model declares no driver"\n'
        "        Exit Function\n"
        "    End If\n",
        "    If driverCount < 1 Then Exit Function\n",
    ):
        damaged = _after(
            _ORIGINAL,
            "    If driverCount < 0 Then\n"
            '        detail = "engine: a negative driver count"\n'
            "        Exit Function\n"
            "    End If\n",
            invented)
        _control("test_15", damaged)


def test_36_the_zero_driver_carrier_slot_is_treated_as_a_real_component() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        ReDim prepared(0 To 0)\n"
        "        SimEnginePrepare = True\n",
        "        ReDim prepared(0 To 0)\n"
        "        prepared(0).PermanentId = components(0).PermanentId\n"
        "        prepared(0).ValueInitialState = components(0).InitialState\n"
        "        SimEnginePrepare = True\n")
    _control("test_16", damaged)


def test_37_a_driver_array_bound_is_read_on_the_zero_driver_path() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        ReDim costIds(0 To 0)\n"
        "        ReDim riskIds(0 To 0)\n"
        "        ReDim claimed(0 To 0)\n",
        "        ReDim costIds(0 To 0)\n"
        "        ReDim riskIds(0 To 0)\n"
        "        ReDim claimed(0 To 0)\n"
        "        If UBound(drivers) >= LBound(drivers) Then costIds(0) = vbNullString\n")
    _control("test_16", damaged)


def test_38_a_statistic_is_implemented_prematurely() -> None:
    damaged = _after(
        _ORIGINAL,
        "Option Explicit\n",
        "\nPublic Function SimEngineMean(ByRef totals() As Double, ByVal count As Long, _\n"
        "                              ByRef result As Double) As Boolean\n"
        "    Dim total As Double\n"
        "    If Not SafeSignedSum(totals, count, total) Then Exit Function\n"
        "    result = total / CDbl(count)\n"
        "    SimEngineMean = True\n"
        "End Function\n")
    _control("test_10", damaged)


def test_38a_a_result_digest_is_implemented_prematurely() -> None:
    damaged = _after(
        _ORIGINAL,
        "Option Explicit\n",
        "\nPrivate Function SimEngineResultDigest(ByRef totals() As Double, _\n"
        "                                       ByVal count As Long) As String\n"
        "    SimEngineResultDigest = CStr(count)\n"
        "End Function\n")
    _control("test_10", damaged)


if __name__ == "__main__":  # pragma: no cover
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
