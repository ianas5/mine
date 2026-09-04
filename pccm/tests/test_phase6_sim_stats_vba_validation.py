#!/usr/bin/env python3
"""PCCM Phase 6 Step-9 MUTATION CONTROLS for the modSimStats conformance battery.

A conformance test that cannot fail proves nothing. Every control damages the
accepted source, reruns the WHOLE Step-9 battery against the damaged copy, and
requires a NAMED detector among the refusers.

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

import test_phase6_sim_stats_vba as conformance  # noqa: E402

_ORIGINAL = conformance.SIM_STATS_BAS.read_text(encoding="utf-8")
# The whole accepted battery runs in under two seconds, the 100,000-value
# practicality test included. A budget of one minute is a hundredfold of the
# slowest accepted test, so only a genuinely non-terminating mutation trips it.
_TEST_BUDGET_SECONDS = 60


class _Timeout(Exception):
    pass


def _conformance_tests() -> list[str]:
    names = sorted(n for n in dir(conformance) if n.startswith("test_"))
    assert len(names) >= 45, names
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
    saved = (conformance.SIM_STATS_BAS, dict(conformance._CACHE))
    conformance._CACHE.clear()
    if source is not None:
        assert source != _ORIGINAL, "the mutation changed nothing"
        temp = Path(tempfile.mkdtemp(prefix="pccm-step9-mutation-"))
        target = temp / "modSimStats.bas"
        target.write_text(source, encoding="utf-8")
        conformance.SIM_STATS_BAS = target

    def restore() -> None:
        conformance.SIM_STATS_BAS = saved[0]
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
# 1-3. The exactness shortcuts
# ===========================================================================
_MEAN_SHORTCUT = """    If SimStatsConstantValue(values, count, constantValue) Then
        result = constantValue
        SimStatsMean = True
        Exit Function
    End If
"""
_SD_SHORTCUT = """    If SimStatsConstantValue(values, count, constantValue) Then
        result = 0#
        SimStatsSampleStandardDeviation = True
        Exit Function
    End If
"""


def test_01_the_constant_mean_shortcut_is_removed() -> None:
    _control("test_17", _swap(_ORIGINAL, _MEAN_SHORTCUT, ""))


def test_02_the_constant_deviation_shortcut_is_removed() -> None:
    _control("test_17", _swap(_ORIGINAL, _SD_SHORTCUT, ""))


def test_03_the_equal_bracket_quantile_shortcut_is_removed() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    ElseIf low = high Then\n",
        "    ElseIf low <> low Then\n")
    _control("test_30", damaged)


# ===========================================================================
# 4-7. The mean
# ===========================================================================
def test_04_the_mean_uses_a_raw_sum_over_n() -> None:
    damaged = _swap(_ORIGINAL, _MEAN_SHORTCUT, "")
    damaged = _swap(
        damaged,
        "    ReDim scaled(0 To count - 1)\n"
        "    For index = 0 To count - 1\n"
        "        scaled(index) = values(LBound(values) + index) / unitScale\n"
        "    Next index\n"
        "\n    If Not SafeSignedSum(scaled, count, total) Then\n"
        '        detail = "statistics: the normalised accumulation for the mean is not representable"\n'
        "        Exit Function\n"
        "    End If\n"
        "    If Not SafeDivide(total, CDbl(count), normalised) Then\n"
        '        detail = "statistics: the normalised mean is not representable"\n'
        "        Exit Function\n"
        "    End If\n"
        "    If Not SafeMultiply(normalised, unitScale, candidate) Then\n"
        '        detail = "statistics: the mean rescale is not representable"\n'
        "        Exit Function\n"
        "    End If\n",
        "    total = 0#\n"
        "    For index = 0 To count - 1\n"
        "        total = total + values(LBound(values) + index)\n"
        "    Next index\n"
        "    candidate = total / CDbl(count)\n")
    _control("test_16", damaged)


def test_05_the_scale_is_the_magnitude_rather_than_a_power_of_two() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    candidate = 1#\n"
        "    If candidate > largest Then\n"
        "        Do While candidate > largest\n"
        "            candidate = candidate / 2#\n"
        "        Loop\n"
        "    Else\n"
        "        Do While candidate <= largest / 2#\n"
        "            candidate = candidate * 2#\n"
        "        Loop\n"
        "    End If\n",
        "    candidate = largest\n")
    _control("test_19", damaged)


def test_06_the_scale_is_the_next_power_of_two_above_the_magnitude() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        Do While candidate <= largest / 2#\n"
        "            candidate = candidate * 2#\n"
        "        Loop\n",
        "        Do While candidate < largest\n"
        "            candidate = candidate * 2#\n"
        "        Loop\n")
    _control("test_19", damaged)


def test_07_the_moments_are_taken_over_the_sorted_copy() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If Not SimStatsMean(values, count, measured, detail) Then Exit Function\n"
        "    candidate.Mean = measured\n"
        "    If Not SimStatsSampleStandardDeviation(values, count, measured, detail) Then Exit Function\n",
        "    If Not SimStatsMean(ordered, count, measured, detail) Then Exit Function\n"
        "    candidate.Mean = measured\n"
        "    If Not SimStatsSampleStandardDeviation(ordered, count, measured, detail) Then Exit Function\n")
    _control("test_20", damaged)


# ===========================================================================
# 8-11. The sample deviation
# ===========================================================================
def test_08_the_deviation_uses_the_population_divisor() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If Not SafeDivide(residual, CDbl(count - 1), variance) Then\n",
        "    If Not SafeDivide(residual, CDbl(count), variance) Then\n")
    _control("test_22", damaged)


def test_09_the_deviation_uses_a_sum_of_squares() -> None:
    damaged = _swap(_ORIGINAL, _SD_SHORTCUT, "")
    damaged = _swap(
        damaged,
        "    ReDim squares(0 To count - 1)\n"
        "    For index = 0 To count - 1\n"
        "        deviation = scaled(index) - centre\n"
        "        squares(index) = deviation * deviation\n"
        "    Next index\n",
        "    ReDim squares(0 To count - 1)\n"
        "    For index = 0 To count - 1\n"
        "        deviation = values(LBound(values) + index)\n"
        "        squares(index) = deviation * deviation\n"
        "    Next index\n")
    _control("test_21", damaged)


def test_10_the_deviation_uses_an_unguarded_original_scale_deviation() -> None:
    damaged = _swap(_ORIGINAL, _SD_SHORTCUT, "")
    damaged = _swap(
        damaged,
        "        deviation = scaled(index) - centre\n",
        "        deviation = values(LBound(values) + index) - centre * unitScale\n")
    _control("test_21", damaged)


def test_11_an_unrepresentable_varying_dispersion_is_returned_as_zero() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If Not SafeMultiply(Sqr(variance), unitScale, candidate) Then\n"
        '        detail = "statistics: the sample standard deviation rescale is not representable"\n'
        "        Exit Function\n"
        "    End If\n",
        "    If Not SafeMultiply(Sqr(variance), unitScale, candidate) Then\n"
        "        candidate = 0#\n"
        "    End If\n")
    _control("test_25", damaged)


# ===========================================================================
# 12-15. Type 7
# ===========================================================================
def test_12_the_type_7_position_formula_is_changed() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    h = CDbl(count - 1) * p\n",
        "    h = CDbl(count + 1) * p\n")
    _control("test_28", damaged)


def test_13_a_nearest_rank_method_is_substituted() -> None:
    # The interpolation fraction is now returned by the shared position owner
    # rather than assigned inside the quantile, so the mutation is applied where
    # the arithmetic actually lives. Forcing it to zero is still exactly a
    # nearest-rank substitution: every percentile would collapse onto the lower
    # order statistic.
    damaged = _swap(
        _ORIGINAL,
        "    SimStatsPositionOf = h - CDbl(lowIndex)\n",
        "    SimStatsPositionOf = 0#\n")
    _control("test_28", damaged)


def test_14_the_integral_order_statistic_is_formed_by_interpolation() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If fraction = 0# Then\n"
        "        ' An integral h selects an order statistic outright. Returning it\n"
        "        ' untouched rather than forming 1 * low + 0 * high keeps p = 0 and p = 1\n"
        "        ' exact at every magnitude, including subnormals.\n"
        "        candidate = low\n"
        "    ElseIf low = high Then\n",
        "    If fraction < 0# Then\n"
        "        candidate = low\n"
        "    ElseIf low = high Then\n")
    _control("test_29", damaged)


def test_15_the_convex_interpolation_is_replaced_by_the_difference_form() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        candidate = (1# - fraction) * low + fraction * high\n",
        "        candidate = low + fraction * (high - low)\n")
    _control("test_31", damaged)


# ===========================================================================
# 16-17. Sorting
# ===========================================================================
def test_16_the_callers_array_is_sorted_in_place() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    ReDim ordered(0 To count - 1)\n"
        "    For index = 0 To count - 1\n"
        "        ordered(index) = values(LBound(values) + index)\n"
        "    Next index\n"
        "    If Not SimStatsSortAscending(ordered, count, detail) Then Exit Function\n",
        "    If Not SimStatsSortAscending(values, count, detail) Then Exit Function\n"
        "    ReDim ordered(0 To count - 1)\n"
        "    For index = 0 To count - 1\n"
        "        ordered(index) = values(LBound(values) + index)\n"
        "    Next index\n")
    _control("test_10", damaged)


def test_17_describe_sorts_separately_for_every_quantile() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        If Not SimStatsQuantileSorted(ordered, count, p, measured, detail) Then Exit Function\n",
        "        If Not SimStatsSortedCopy(values, count, ordered, detail) Then Exit Function\n"
        "        If Not SimStatsQuantileSorted(ordered, count, p, measured, detail) Then Exit Function\n")
    _control("test_11", damaged)


def test_17a_a_quadratic_insertion_sort_is_substituted() -> None:
    """A CORRECT stable insertion sort. It orders every shape properly, so only
    the source shape and the 100,000-value practicality test can tell it from
    the accepted bottom-up merge."""
    opening = _ORIGINAL.index("    ReDim scratch(0 To count - 1)\n")
    closing = _ORIGINAL.index(
        "\n    SimStatsSortAscending = True\nEnd Function", opening)
    damaged = _swap(
        _ORIGINAL, "    Dim scratch() As Double\n", "    Dim held As Double\n")
    damaged = _swap(
        damaged,
        _ORIGINAL[opening:closing],
        "    For fromLow = 1 To count - 1\n"
        "        held = series(fromLow)\n"
        "        fromHigh = fromLow - 1\n"
        "        Do While fromHigh >= 0\n"
        "            If series(fromHigh) <= held Then Exit Do\n"
        "            series(fromHigh + 1) = series(fromHigh)\n"
        "            fromHigh = fromHigh - 1\n"
        "        Loop\n"
        "        series(fromHigh + 1) = held\n"
        "    Next fromLow")
    _control("test_13", damaged)


# ===========================================================================
# 18-22. The ladder and the selected level
# ===========================================================================
def test_18_the_ladder_hardcodes_its_own_labels() -> None:
    damaged = _swap(_ORIGINAL, "        label = SIM_QUANTILE_3\n", '        label = "P55"\n')
    _control("test_33", damaged)


def test_19_one_selectable_rung_is_omitted() -> None:
    damaged = _swap(_ORIGINAL, "        label = SIM_QUANTILE_9\n",
                    "        label = SIM_QUANTILE_8\n")
    _control("test_33", damaged)


def test_20_the_fixed_rung_is_omitted_from_the_ladder() -> None:
    damaged = _swap(_ORIGINAL, "        label = SIM_QUANTILE_1\n",
                    "        label = SIM_QUANTILE_2\n")
    _control("test_33", damaged)


def test_21_the_fixed_rung_is_accepted_as_a_selector() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If StrComp(selectedLabel, SIM_QUANTILE_FIXED_1, vbBinaryCompare) = 0 Then\n"
        '        detail = "statistics: " & SIM_QUANTILE_FIXED_1 & " is reported and fixed; it is not selectable"\n'
        "        Exit Function\n"
        "    End If\n",
        "")
    _control("test_36", damaged)


def test_22_an_unknown_confidence_level_is_silently_accepted() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If found < 0 Then\n"
        '        detail = "statistics: an unknown confidence level"\n'
        "        Exit Function\n"
        "    End If\n",
        "    If found < 0 Then\n"
        "        found = 0\n"
        "    End If\n")
    _control("test_37", damaged)


# ===========================================================================
# 23-27. Contingency and transactional output
# ===========================================================================
def test_23_contingency_uses_raw_subtraction() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If Not SafeSubtract(selectedTotal, baseEstimateA, candidate) Then\n"
        '        detail = "contingency: the selected total minus the deterministic base estimate A " & _\n'
        '                 "is not representable"\n'
        "        Exit Function\n"
        "    End If\n",
        "    candidate = selectedTotal - baseEstimateA\n")
    _control("test_38", damaged)


def test_24_a_negative_contingency_is_clamped_to_zero() -> None:
    damaged = _after(
        _ORIGINAL,
        "    If Not SafeSubtract(selectedTotal, baseEstimateA, candidate) Then\n"
        '        detail = "contingency: the selected total minus the deterministic base estimate A " & _\n'
        '                 "is not representable"\n'
        "        Exit Function\n"
        "    End If\n",
        "    If candidate < 0# Then candidate = 0#\n")
    _control("test_39", damaged)


def test_25_a_wrong_contingency_baseline_is_introduced() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If Not SafeSubtract(selectedTotal, baseEstimateA, candidate) Then\n",
        "    If Not SafeSubtract(selectedTotal, baseEstimateA * 2#, candidate) Then\n")
    _control("test_38", damaged)


def test_26_the_selected_value_is_taken_from_the_wrong_rung() -> None:
    """Nominal and PV are described separately; mixing rungs is the same defect."""
    damaged = _swap(
        _ORIGINAL,
        "    result = quantileValues(LBound(quantileValues) + found)\n",
        "    result = quantileValues(LBound(quantileValues))\n")
    _control("test_36", damaged)


def test_27_a_partial_ladder_is_published_before_every_value_succeeds() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        labels(index) = label\n"
        "        ladder(index) = measured\n"
        "    Next index\n",
        "        labels(index) = label\n"
        "        ladder(index) = measured\n"
        "        quantileLabels = labels\n"
        "        quantileValues = ladder\n"
        "    Next index\n")
    _control("test_42", damaged)


# ===========================================================================
# 28-32. Purity, provenance and boundary
# ===========================================================================
def test_28_a_worksheet_statistic_is_introduced() -> None:
    damaged = _swap(
        _ORIGINAL, _MEAN_SHORTCUT,
        "    If Application.WorksheetFunction.Count(values) = count Then\n"
        "        result = constantValue\n"
        "        SimStatsMean = True\n"
        "        Exit Function\n"
        "    End If\n")
    _control("test_05", damaged)


def test_29_the_globally_forbidden_token_is_introduced() -> None:
    damaged = _after(
        _ORIGINAL, "Option Explicit\n",
        "\nPrivate Function SimStatsPercentileAlias(ByVal p As Double) As Double\n"
        "    SimStatsPercentileAlias = p\nEnd Function\n")
    _control("test_07", damaged)


def test_30_a_result_digest_is_introduced_prematurely() -> None:
    damaged = _after(
        _ORIGINAL, "Option Explicit\n",
        "\nPrivate Function SimStatsResultDigest(ByVal count As Long) As String\n"
        "    SimStatsResultDigest = CStr(count)\nEnd Function\n")
    _control("test_08", damaged)


def test_31_module_level_mutable_state_is_introduced() -> None:
    _control("test_06", _after(
        _ORIGINAL, "Option Explicit\n", "\nPrivate mLastMean As Double\n"))


def test_31a_a_static_local_is_introduced() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    Dim constantValue As Double, index As Long\n\n    detail = vbNullString\n"
        '    If count < 1 Then\n',
        "    Static constantValue As Double\n    Dim index As Long\n\n    detail = vbNullString\n"
        '    If count < 1 Then\n')
    _control("test_06", damaged)


def test_32_an_empty_sequence_reads_a_bound_before_refusing() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If count = 0 Then\n"
        '        detail = "statistics: an empty sequence has no " & where\n'
        "        Exit Function\n"
        "    End If\n"
        "    For index = 0 To count - 1\n",
        "    If LBound(values) > UBound(values) Then Exit Function\n"
        "    If count = 0 Then\n"
        '        detail = "statistics: an empty sequence has no " & where\n'
        "        Exit Function\n"
        "    End If\n"
        "    For index = 0 To count - 1\n")
    _control("test_27", damaged)


# ===========================================================================
# 33-38. The ladder-integrity boundary (Step-9 final hardening)
# ===========================================================================
_VALIDATE_CALL = """    If Not SimStatsValidateLadder(quantileLabels, quantileValues, quantileCount, detail) Then
        Exit Function
    End If
"""
_LABEL_CHECK = """        If StrComp(quantileLabels(LBound(quantileLabels) + index), expectedLabel, _
                   vbBinaryCompare) <> 0 Then
            detail = "statistics: the ladder is not the accepted projection at " & expectedLabel
            Exit Function
        End If
"""
_FINITE_CHECK = """        If Not IsUsableDouble(quantileValues(LBound(quantileValues) + index)) Then
            detail = "statistics: the ladder carries a value at " & expectedLabel & _
                     " that is not a finite Double"
            Exit Function
        End If
"""
_EXTENT_CHECKS = """    If Not SimStatsLadderExtent(quantileLabels, quantileValues, labelExtent, valueExtent) Then
        detail = "statistics: the ladder carrier is not allocated"
        Exit Function
    End If
    If labelExtent <> quantileCount Then
        detail = "statistics: the ladder label carrier is not the accepted length"
        Exit Function
    End If
    If valueExtent <> quantileCount Then
        detail = "statistics: the ladder value carrier is not the accepted length"
        Exit Function
    End If
"""


def test_33_the_ladder_is_not_validated_before_it_is_searched() -> None:
    """The f44af49 boundary exactly: membership in the caller's array decides."""
    _control("test_51", _swap(_ORIGINAL, _VALIDATE_CALL, ""))


def test_34_the_owner_label_authority_is_removed() -> None:
    _control("test_51", _swap(_ORIGINAL, _LABEL_CHECK, ""))


def test_35_the_labels_are_checked_for_membership_and_not_for_position() -> None:
    """Every accepted rung is present, but the ORDER is no longer proved."""
    damaged = _swap(
        _ORIGINAL, _LABEL_CHECK,
        "        valueExtent = 0\n"
        "        For labelExtent = 0 To quantileCount - 1\n"
        "            If StrComp(quantileLabels(LBound(quantileLabels) + labelExtent), "
        "expectedLabel, _\n"
        "                       vbBinaryCompare) = 0 Then\n"
        "                valueExtent = 1\n"
        "            End If\n"
        "        Next labelExtent\n"
        "        If valueExtent = 0 Then\n"
        '            detail = "statistics: the ladder is not the accepted projection at " '
        "& expectedLabel\n"
        "            Exit Function\n"
        "        End If\n")
    _control("test_53", damaged)


def test_36_a_case_insensitive_label_comparison_is_substituted() -> None:
    damaged = _swap(
        _ORIGINAL,
        "        If StrComp(quantileLabels(LBound(quantileLabels) + index), expectedLabel, _\n"
        "                   vbBinaryCompare) <> 0 Then\n",
        "        If StrComp(quantileLabels(LBound(quantileLabels) + index), expectedLabel, _\n"
        "                   vbTextCompare) <> 0 Then\n")
    _control("test_56", damaged)


def test_37_the_ladder_value_finiteness_check_is_removed() -> None:
    _control("test_57", _swap(_ORIGINAL, _FINITE_CHECK, ""))


def test_38_the_physical_carrier_length_checks_are_removed() -> None:
    _control("test_59", _swap(_ORIGINAL, _EXTENT_CHECKS, ""))


def test_38a_only_the_label_carrier_length_is_checked() -> None:
    damaged = _swap(
        _ORIGINAL,
        "    If valueExtent <> quantileCount Then\n"
        '        detail = "statistics: the ladder value carrier is not the accepted length"\n'
        "        Exit Function\n"
        "    End If\n",
        "")
    _control("test_60", damaged)


def test_38b_the_scoped_bounds_handler_becomes_a_broad_suppression() -> None:
    damaged = _swap(
        _ORIGINAL, "    On Error GoTo Unallocated\n", "    On Error Resume Next\n")
    _control("test_48", damaged)


def test_38c_the_selected_value_is_read_before_the_ladder_is_proved() -> None:
    """Validation that runs but no longer GATES is validation in name only."""
    damaged = _swap(
        _ORIGINAL,
        "    If Not SimStatsValidateLadder(quantileLabels, quantileValues, quantileCount, "
        "detail) Then\n        Exit Function\n    End If\n",
        "    If Not SimStatsValidateLadder(quantileLabels, quantileValues, quantileCount, "
        "detail) Then\n        detail = vbNullString\n    End If\n")
    _control("test_51", damaged)


if __name__ == "__main__":  # pragma: no cover
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
