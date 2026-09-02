#!/usr/bin/env python3
"""PCCM Phase 7 Step-2 mutation controls for `modSimSensitivity`.

Each mutation is a plausible WRONG IMPLEMENTATION of the sensitivity kernel,
applied to the real source, and each must be refused by a named control. These
protect the mathematics and the ownership boundary. They do not police wording:
a comment can be improved without a control failing, and a comment cannot be the
thing that makes the module correct.

Runs standalone or under pytest.
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import test_phase7_sim_sensitivity as conformance  # noqa: E402

SENSITIVITY_BAS = PCCM_ROOT / "src" / "vba" / "modSimSensitivity.bas"
_SOURCE = SENSITIVITY_BAS.read_text(encoding="utf-8")


def _swap(text: str, old: str, new: str, count: int = 1) -> str:
    assert text.count(old) == count, (old[:70], text.count(old))
    return text.replace(old, new, count)


def _refused(damaged: str, controls, reason: str) -> None:
    """Install a damaged module and require a named control to refuse it."""
    assert damaged != _SOURCE, f"{reason}: the mutation changed nothing"
    with tempfile.TemporaryDirectory(prefix="pccm-sensitivity-") as name:
        target = Path(name) / SENSITIVITY_BAS.name
        target.write_text(damaged, encoding="utf-8")
        saved_path, saved_cache = conformance.SENSITIVITY_BAS, dict(conformance._CACHE)
        conformance.SENSITIVITY_BAS = target
        conformance._CACHE.clear()
        try:
            failures = []
            for control in controls:
                try:
                    control()
                except (AssertionError, Exception) as error:  # noqa: BLE001
                    failures.append(f"{control.__name__}: {error}")
        finally:
            conformance.SENSITIVITY_BAS = saved_path
            conformance._CACHE.clear()
            conformance._CACHE.update(saved_cache)
    assert failures, f"{reason}: the mutation survived every named control"


# ===========================================================================
# A. THE MATHEMATICS
# ===========================================================================
def test_01_ordinal_ranks_replace_mid_ranks() -> None:
    """The single most consequential wrong answer: give each tied observation a
    distinct rank in encounter order. Untied data is unaffected, so only a tied
    vector can catch it."""
    damaged = _swap(
        _SOURCE,
        "        high = SimSensitivityUpperBound(ordered, count, values(LBound(values) + index))",
        "        high = SimSensitivityLowerBound(ordered, count, values(LBound(values) + index))")
    _refused(damaged, (conformance.test_02_the_vba_reproduces_every_hand_derived_mid_rank,
                       conformance.test_03_a_tie_block_takes_the_midpoint_of_the_positions_it_spans,
                       conformance.test_04_every_rank_vector_sums_to_the_triangular_number),
             "ordinal ranks replaced mid-ranks")


def test_02_the_tie_block_takes_its_first_position_instead_of_the_midpoint() -> None:
    damaged = _swap(_SOURCE, "        ranks(index) = average",
                    "        ranks(index) = first")
    _refused(damaged, (conformance.test_02_the_vba_reproduces_every_hand_derived_mid_rank,
                       conformance.test_04_every_rank_vector_sums_to_the_triangular_number),
             "the tie block took its first position")


def test_03_the_ranks_become_zero_based() -> None:
    """A 1-based convention is not decoration: it is what makes the tie-block
    midpoint (p + q) / 2 come out as the accepted vectors."""
    damaged = _swap(_SOURCE, "        first = CDbl(low + 1)", "        first = CDbl(low)")
    _refused(damaged, (conformance.test_02_the_vba_reproduces_every_hand_derived_mid_rank,
                       conformance.test_04_every_rank_vector_sums_to_the_triangular_number),
             "the ranks became zero-based")


def test_04_the_correlation_loses_the_centring() -> None:
    damaged = _swap(_SOURCE, "        dx = driverRanks(LBound(driverRanks) + index) - meanX",
                    "        dx = driverRanks(LBound(driverRanks) + index)")
    _refused(damaged, (conformance.test_05_a_perfect_positive_monotone_relation_is_plus_one,
                       conformance.test_06_a_perfect_negative_monotone_relation_is_minus_one,
                       conformance.test_08_a_tied_driver_matches_an_independent_mid_rank_pearson),
             "the correlation stopped centring the driver ranks")


def test_05_the_sign_of_the_correlation_is_dropped() -> None:
    """A tornado that cannot say whether a driver pushes the total up or down
    is not a sensitivity analysis."""
    damaged = _swap(_SOURCE, "    rho = quotient", "    rho = Abs(quotient)")
    _refused(damaged, (conformance.test_06_a_perfect_negative_monotone_relation_is_minus_one,
                       conformance.test_08_a_tied_driver_matches_an_independent_mid_rank_pearson),
             "the correlation lost its sign")


def test_06_the_denominator_uses_one_series_only() -> None:
    damaged = _swap(_SOURCE, "    If Not SafeMultiply(sxx, syy, product) Then",
                    "    If Not SafeMultiply(sxx, sxx, product) Then")
    _refused(damaged, (conformance.test_05_a_perfect_positive_monotone_relation_is_plus_one,
                       conformance.test_08_a_tied_driver_matches_an_independent_mid_rank_pearson),
             "the correlation denominator dropped a series")


def test_07_the_clamp_hides_an_impossible_correlation() -> None:
    """THE CLAMP IS DEFENSIVE, AND THAT IS WHY IT NEEDS A SOURCE CONTROL.

    On well-conditioned data the quotient never leaves [-1, 1], so widening the
    clamp changes no answer any behavioural test can reach. A control that ran
    vectors and found nothing would be reporting that the mutation is harmless
    rather than that it was caught. The bound is pinned where it is visible.
    """
    damaged = _swap(_SOURCE, "    If quotient > 1# Then quotient = 1#",
                    "    If quotient > 2# Then quotient = 2#")
    damaged = _swap(damaged, "    If quotient < -1# Then quotient = -1#",
                    "    If quotient < -2# Then quotient = -2#")
    _refused(damaged, (conformance.test_31_the_correlation_bound_is_exactly_one,),
             "the correlation clamp was widened past one")


# ===========================================================================
# B. UNDEFINED IS NOT ZERO
# ===========================================================================
def test_08_zero_variance_is_reported_as_a_defined_zero_rho() -> None:
    """The defect the status exists to prevent. rho = 0 with status DEFINED says
    a monotone association was looked for and not found; a constant column
    offers nothing to look for."""
    damaged = _swap(
        _SOURCE,
        "    If sxx = 0# Or syy = 0# Then\n        status = SIM_SENSITIVITY_NO_VARIANCE",
        "    If sxx = 0# Or syy = 0# Then\n        status = SIM_SENSITIVITY_DEFINED")
    _refused(damaged, (conformance.test_10_a_constant_driver_is_undefined_and_not_zero_rho,
                       conformance.test_11_a_constant_total_is_undefined_too,
                       conformance.test_12_a_genuine_zero_correlation_is_defined_and_distinguishable),
             "zero variance was reported as a defined zero rho")


def test_09_only_the_driver_is_checked_for_variance() -> None:
    """A constant TOTAL is undefined too, and a control that only looked at the
    driver would miss it."""
    # BOTH GUARDS GO. Removing only the first one is survivable, and correctly
    # so: the denominator check below it catches a constant total on its own.
    # A mutation that the code legitimately withstands proves nothing about the
    # control, so the damage here removes the defence in depth as well.
    damaged = _swap(_SOURCE, "    If sxx = 0# Or syy = 0# Then",
                    "    If sxx = 0# Then")
    damaged = _swap(damaged, "    If denominator = 0# Then",
                    "    If False Then")
    _refused(damaged, (conformance.test_11_a_constant_total_is_undefined_too,),
             "a constant total was accepted as a defined correlation")


# ===========================================================================
# C. RANKING AND THE POPULATION
# ===========================================================================
def test_10_the_order_becomes_signed_rho_instead_of_magnitude() -> None:
    """A driver at rho = -0.9 is the strongest in the model, not the weakest."""
    damaged = _swap(_SOURCE, "    leftAbs = results(LBound(results) + left).AbsRho",
                    "    leftAbs = results(LBound(results) + left).Rho")
    damaged = _swap(damaged, "    rightAbs = results(LBound(results) + right).AbsRho",
                    "    rightAbs = results(LBound(results) + right).Rho")
    _refused(damaged, (conformance.test_13_the_population_is_ordered_by_absolute_rho_descending,),
             "the ranking ordered on signed rho")


def test_11_the_order_is_reversed() -> None:
    damaged = _swap(_SOURCE, "    If leftAbs > rightAbs Then\n        SimSensitivityPrecedes = True",
                    "    If leftAbs < rightAbs Then\n        SimSensitivityPrecedes = True")
    _refused(damaged, (conformance.test_13_the_population_is_ordered_by_absolute_rho_descending,),
             "the ranking ran ascending")


def test_12_the_equal_magnitude_tie_break_is_dropped() -> None:
    """Without it the report depends on the order the drivers were supplied in,
    so the same model yields two different tables."""
    damaged = _swap(
        _SOURCE,
        "    SimSensitivityPrecedes = SimSensitivityIdPrecedes( _",
        "    SimSensitivityPrecedes = False\n    If True Then Exit Function\n    SimSensitivityPrecedes = SimSensitivityIdPrecedes( _")
    _refused(damaged, (conformance.test_14_equal_magnitude_is_broken_by_the_permanent_id,),
             "the equal-magnitude tie-break was dropped")


def test_13_the_tie_break_becomes_a_text_comparison() -> None:
    """VBA's `<` on String answers by Option Compare and host locale. An
    ordering that changes with the machine is not an ordering."""
    damaged = _swap(
        _SOURCE,
        "        If leftUnit < rightUnit Then\n            SimSensitivityIdPrecedes = True\n            Exit Function\n        End If",
        "        If Mid$(left, index, 1) < Mid$(right, index, 1) Then\n            SimSensitivityIdPrecedes = True\n            Exit Function\n        End If")
    # NOT A BEHAVIOURAL CONTROL, and it cannot be one: the transcriber runs the
    # comparison in Python, whose own string order agrees with code units on
    # ASCII, so the wrong implementation returns the right answer here and the
    # wrong answer on a Windows host with a different Option Compare.
    _refused(damaged, (conformance.test_32_the_identity_tie_break_never_compares_strings_directly,),
             "the tie-break became a text comparison")


def test_14_a_zero_variance_driver_enters_the_ranked_population() -> None:
    damaged = _swap(
        _SOURCE,
        "        If results(LBound(results) + index).Status = SIM_SENSITIVITY_DEFINED Then",
        "        If True Then")
    _refused(damaged, (conformance.test_16_a_zero_variance_driver_is_not_in_the_ranked_population,),
             "a zero-variance driver entered the ranked population")


def test_15_the_population_is_truncated_to_a_top_n() -> None:
    """Top-N is a chart decision and charts are Phase 8. Truncating here would
    discard data the later phase is entitled to choose from."""
    damaged = _swap(_SOURCE, "    ReDim order(0 To eligibleCount - 1)",
                    "    If eligibleCount > 10 Then eligibleCount = 10\n    ReDim order(0 To eligibleCount - 1)")
    _refused(damaged, (conformance.test_17_nothing_is_truncated,),
             "the ranked population was truncated to a top N")


# ===========================================================================
# D. SOURCE IMMUTABILITY AND PAIRING
# ===========================================================================
def test_16_the_caller_sequence_is_sorted_in_place() -> None:
    """The contract says sorted copies only. Sorting the caller's array would
    also destroy the iteration identity every later pairing depends on."""
    damaged = _swap(
        _SOURCE,
        "    If Not SimSensitivitySortedCopy(values, count, ordered, detail) Then Exit Function",
        "    If Not SimSensitivitySortAscending(values, count, detail) Then Exit Function\n"
        "    If Not SimSensitivitySortedCopy(values, count, ordered, detail) Then Exit Function")
    _refused(damaged, (conformance.test_19_ranking_a_sequence_does_not_reorder_it,
                       conformance.test_21_iteration_alignment_survives_a_shuffled_but_paired_input),
             "the caller's sequence was sorted in place")


def test_17_the_ranks_come_back_in_sorted_order() -> None:
    """The rank of observation j must sit at position j. A rank vector in sorted
    order pairs contribution j with a total it does not belong to."""
    damaged = _swap(_SOURCE, "        ranks(index) = average",
                    "        ranks(index) = CDbl(index + 1)")
    _refused(damaged, (conformance.test_02_the_vba_reproduces_every_hand_derived_mid_rank,
                       conformance.test_19_ranking_a_sequence_does_not_reorder_it),
             "the ranks came back in sorted order")


def test_18_the_shared_total_rank_vector_is_written_to() -> None:
    """The reuse interface only works if a driver cannot damage the vector the
    next driver will use."""
    damaged = _swap(
        _SOURCE,
        "        dy = totalRanks(LBound(totalRanks) + index) - meanY",
        "        totalRanks(LBound(totalRanks) + index) = 0#\n"
        "        dy = totalRanks(LBound(totalRanks) + index) - meanY")
    _refused(damaged, (conformance.test_22_one_total_rank_vector_serves_every_driver_unchanged,),
             "a driver wrote to the shared total-rank vector")


# ===========================================================================
# E. REFUSALS AND OWNERSHIP
# ===========================================================================
def test_19_a_non_finite_observation_is_ranked_instead_of_refused() -> None:
    damaged = _swap(
        _SOURCE,
        "        If Not IsUsableDouble(values(LBound(values) + index)) Then",
        "        If False Then")
    _refused(damaged, (conformance.test_24_a_non_finite_observation_is_refused_rather_than_ranked,),
             "a NaN was ranked instead of refused")


def test_20_a_single_observation_is_correlated() -> None:
    damaged = _swap(_SOURCE, "    If count < 2 Then\n        detail = \"sensitivity: a correlation needs at least two observations\"",
                    "    If count < 1 Then\n        detail = \"sensitivity: a correlation needs at least two observations\"")
    _refused(damaged, (conformance.test_25_mismatched_lengths_and_short_series_are_refused,),
             "a single observation was correlated")


def test_21_the_no_ties_shortcut_is_smuggled_in_as_a_fast_path() -> None:
    """Forbidden outright, including as a fast path taken only when the data
    happens to be untied - which is precisely where it looks harmless."""
    # REAL CODE, not a comment. `VbaModule.code` strips comments, and rightly:
    # a comment cannot compute a correlation, so a control that failed on one
    # would be policing prose.
    damaged = _swap(
        _SOURCE,
        "    If Not SafeMultiply(sxx, syy, product) Then",
        "    If sxy = 0# Then\n"
        "        rho = 1# - 6# * sxy / (CDbl(count) * (CDbl(count) * CDbl(count) - 1#))\n"
        "        SimSensitivityRankCorrelation = True\n"
        "        Exit Function\n"
        "    End If\n"
        "    If Not SafeMultiply(sxx, syy, product) Then")
    _refused(damaged, (conformance.test_27_no_no_ties_spearman_shortcut_appears_in_the_source,),
             "the no-ties shortcut appeared in the source")


def test_22_the_kernel_reaches_the_workbook() -> None:
    damaged = _swap(_SOURCE, "    detail = vbNullString\n    rho = 0#",
                    "    MsgBox \"rho\"\n    detail = vbNullString\n    rho = 0#")
    _refused(damaged, (conformance.test_26_the_module_reaches_nothing_outside_its_own_mathematics,),
             "the kernel reached the workbook")


def test_23_the_kernel_reaches_the_rng() -> None:
    """Replay is P7-3 and reads component streams. A pure kernel that acquired
    one would have taken ownership it was deliberately not given."""
    damaged = _swap(
        _SOURCE,
        "Public Function SimSensitivitySpearman(",
        "Public Function SimSensitivityReplay(ByRef state As SimRngState) As Boolean\n"
        "End Function\n\n"
        "Public Function SimSensitivitySpearman(")
    _refused(damaged, (conformance.test_26_the_module_reaches_nothing_outside_its_own_mathematics,
                       conformance.test_29_the_public_surface_is_exactly_the_kernel),
             "the kernel acquired an RNG dependency")


def test_24_the_status_becomes_a_presentation_string() -> None:
    """Presentation belongs to P7-4. A mathematical routine that returns
    worksheet vocabulary has taken a decision that is not its to take."""
    damaged = _swap(_SOURCE, "Public Const SIM_SENSITIVITY_NO_VARIANCE As Long = 1",
                    "Public Const SIM_SENSITIVITY_NO_VARIANCE As Long = 1\n"
                    "Public Const SIM_SENSITIVITY_NO_VARIANCE_LABEL As String = \"n/a - no variance\"")
    _refused(damaged, (conformance.test_30_the_status_constants_are_codes_and_not_presentation_strings,),
             "a presentation string entered the kernel")
