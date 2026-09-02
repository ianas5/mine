#!/usr/bin/env python3
"""PCCM Phase 7 Step-3 mutation controls for per-driver replay.

Each mutation is a plausible WRONG replay, applied to the real `modSimEngine`
source, and each must be refused by a named control. These protect executable
and ownership properties: the accepted RNG semantics, the single contribution
owner, iteration identity, and the observational boundary. They do not police
wording.

Runs standalone or under pytest.
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import test_phase6_sim_engine_vba as engine  # noqa: E402
import test_phase7_sim_replay as conformance  # noqa: E402

SIM_ENGINE_BAS = PCCM_ROOT / "src" / "vba" / "modSimEngine.bas"
_SOURCE = SIM_ENGINE_BAS.read_text(encoding="utf-8")


def _swap(text: str, old: str, new: str, count: int = 1) -> str:
    assert text.count(old) == count, (old[:70], text.count(old))
    return text.replace(old, new, count)


def _refused(damaged: str, controls, reason: str) -> None:
    assert damaged != _SOURCE, f"{reason}: the mutation changed nothing"
    with tempfile.TemporaryDirectory(prefix="pccm-replay-") as name:
        target = Path(name) / SIM_ENGINE_BAS.name
        target.write_text(damaged, encoding="utf-8")
        saved = (engine.SIM_ENGINE_BAS, conformance.SIM_ENGINE_BAS,
                 dict(engine._CACHE))
        engine.SIM_ENGINE_BAS = target
        conformance.SIM_ENGINE_BAS = target
        engine._CACHE.clear()
        try:
            failures = []
            for control in controls:
                try:
                    control()
                except Exception as error:  # noqa: BLE001
                    failures.append(f"{control.__name__}: {error}")
        finally:
            engine.SIM_ENGINE_BAS, conformance.SIM_ENGINE_BAS = saved[0], saved[1]
            engine._CACHE.clear()
            engine._CACHE.update(saved[2])
    assert failures, f"{reason}: the mutation survived every named control"


# ===========================================================================
# A. THE ACCEPTED RNG SEMANTICS
# ===========================================================================
def test_01_the_severity_sampler_becomes_conditional() -> None:
    """THE MUTATION THIS PACKAGE EXISTS TO REFUSE. Skipping the severity draw
    when a Risk did not occur is faster, looks harmless, and produces a
    different severity sequence from the run the replay claims to explain."""
    damaged = _swap(
        _SOURCE,
        "        ' UNCONDITIONAL, on both kinds. See D6-18b above.\n"
        "        If Not SimEngineSampleValue(prepared(target), valueState, sample, _\n"
        "                                    iteration, detail) Then Exit Function",
        "        If occurred Then\n"
        "        If Not SimEngineSampleValue(prepared(target), valueState, sample, _\n"
        "                                    iteration, detail) Then Exit Function\n"
        "        End If")
    _refused(damaged, (conformance.test_01_every_single_driver_case_replays_bit_for_bit,
                       conformance.test_04_severity_advances_on_non_occurrence_iterations,
                       conformance.test_05_probability_changes_which_iterations_occur_and_not_the_severities),
             "the severity sampler became conditional")


def test_02_the_occurrence_draw_is_taken_after_the_severity() -> None:
    """Order matters: the two streams are separate, but the accepted loop draws
    occurrence first, and a replay that reversed it would still have to agree."""
    damaged = _swap(
        _SOURCE,
        "        occurred = True\n        If prepared(target).HasOccurrenceStream Then",
        "        occurred = True\n        If False Then")
    _refused(damaged, (conformance.test_01_every_single_driver_case_replays_bit_for_bit,
                       conformance.test_03_a_risk_at_probability_one_contributes_at_every_iteration),
             "the occurrence draw was skipped")


def test_03_the_replay_starts_from_the_wrong_component_state() -> None:
    damaged = _swap(_SOURCE, "    valueState = prepared(target).ValueInitialState",
                    "    valueState = prepared(target).OccurrenceInitialState")
    _refused(damaged, (conformance.test_01_every_single_driver_case_replays_bit_for_bit,
                       conformance.test_06_replaying_every_driver_rebuilds_the_published_total),
             "the replay started from the wrong stream")


def test_04_the_occurrence_stream_is_reinitialised_every_iteration() -> None:
    """A stream re-seeded inside the loop repeats its first draw forever."""
    damaged = _swap(
        _SOURCE,
        "        occurred = True\n        If prepared(target).HasOccurrenceStream Then",
        "        occurred = True\n        occurrenceState = prepared(target).OccurrenceInitialState\n"
        "        If prepared(target).HasOccurrenceStream Then")
    _refused(damaged, (conformance.test_01_every_single_driver_case_replays_bit_for_bit,),
             "the occurrence stream was reinitialised inside the loop")


def test_05_the_value_stream_is_reinitialised_every_iteration() -> None:
    damaged = _swap(
        _SOURCE,
        "        ' UNCONDITIONAL, on both kinds. See D6-18b above.",
        "        valueState = prepared(target).ValueInitialState\n"
        "        ' UNCONDITIONAL, on both kinds. See D6-18b above.")
    _refused(damaged, (conformance.test_01_every_single_driver_case_replays_bit_for_bit,
                       conformance.test_09_a_shorter_replay_is_a_prefix_of_a_longer_one),
             "the value stream was reinitialised inside the loop")


# ===========================================================================
# B. THE SINGLE CONTRIBUTION OWNER
# ===========================================================================
def test_06_the_replay_carries_its_own_contribution_expression() -> None:
    """A second copy of the arithmetic is the defect the extraction removes. It
    agrees today, which is exactly why nothing would notice it drifting."""
    damaged = _swap(
        _SOURCE,
        "        If Not SimEngineContribution(prepared(target), sample, occurred, _\n"
        "                                     prepared(target).Knom, SIM_MEASURE_NOMINAL, _\n"
        "                                     iteration, term, detail) Then Exit Function",
        "        term = 0#\n"
        "        If occurred Then\n"
        "            factorsLocal(0) = sample\n"
        "            factorsLocal(1) = prepared(target).Quantity\n"
        "            factorsLocal(2) = prepared(target).Knom\n"
        "            If Not SafeProduct(factorsLocal, 3, term) Then Exit Function\n"
        "        End If")
    damaged = _swap(damaged, "    Dim occurred As Boolean\n\n    detail = vbNullString\n    If iterations < 1 Then",
                    "    Dim occurred As Boolean\n    Dim factorsLocal(0 To 2) As Double\n\n"
                    "    detail = vbNullString\n    If iterations < 1 Then")
    _refused(damaged, (conformance.test_17_the_contribution_arithmetic_has_exactly_one_owner,
                       conformance.test_18_both_the_simulation_and_the_replay_call_it),
             "the replay grew its own contribution expression")


def test_07_a_risk_gains_a_quantity_factor() -> None:
    damaged = _swap(_SOURCE, "        factors(0) = sample\n        factors(1) = factor\n        count = 2",
                    "        factors(0) = sample\n        factors(1) = factor\n"
                    "        factors(2) = prepared.Quantity\n        count = 3")
    # A risk's Quantity is zero, so the extra factor zeroes every occurrence -
    # visible on any fixture with a risk that occurs.
    _refused(damaged, (conformance.test_01b_replay_is_bit_exact_where_nominal_and_pv_differ,
                       conformance.test_03_a_risk_at_probability_one_contributes_at_every_iteration),
             "a risk gained a Quantity factor")


def test_08_probability_is_folded_into_the_risk_contribution() -> None:
    """It was already spent on the Bernoulli draw; applying it again would
    discount every occurrence twice."""
    damaged = _swap(_SOURCE, "        factors(1) = factor\n        count = 2",
                    "        factors(1) = factor * prepared.Probability\n        count = 2")
    _refused(damaged, (conformance.test_01_every_single_driver_case_replays_bit_for_bit,
                       conformance.test_19_the_shared_routine_keeps_the_accepted_shapes),
             "Probability was folded into the contribution")


def test_09_a_non_occurring_risk_contributes_its_severity() -> None:
    damaged = _swap(_SOURCE, "        If Not occurred Then\n            SimEngineContribution = True\n            Exit Function\n        End If",
                    "        If False Then\n            SimEngineContribution = True\n            Exit Function\n        End If")
    _refused(damaged, (conformance.test_01_every_single_driver_case_replays_bit_for_bit,
                       conformance.test_02_a_risk_at_probability_zero_contributes_zero_at_every_iteration),
             "a risk that did not occur contributed anyway")


def test_10_the_nominal_measure_is_built_from_the_pv_factor() -> None:
    damaged = _swap(
        _SOURCE,
        "        If Not SimEngineContribution(prepared(target), sample, occurred, _\n"
        "                                     prepared(target).Knom, SIM_MEASURE_NOMINAL, _",
        "        If Not SimEngineContribution(prepared(target), sample, occurred, _\n"
        "                                     prepared(target).Kpv, SIM_MEASURE_NOMINAL, _")
    # ONLY THE DISCOUNTED FIXTURE CAN SEE THIS. Everywhere else Knom = Kpv = 1
    # and the substitution is arithmetically invisible, so naming test_01 here
    # would be crediting a control with a refusal it cannot make.
    _refused(damaged, (conformance.test_01b_replay_is_bit_exact_where_nominal_and_pv_differ,),
             "the replay emitted a PV contribution as nominal")


# ===========================================================================
# C. ITERATION IDENTITY AND ISOLATION
# ===========================================================================
def test_11_the_non_occurrence_zeros_are_compacted_away() -> None:
    """A vector of only the occurring iterations cannot be paired with TotalNom
    by position, which is the one thing P7-4 needs from it."""
    damaged = _swap(
        _SOURCE,
        "        contributions(iteration - 1) = term",
        "        If term <> 0# Then contributions(iteration - 1) = term")
    damaged = _swap(damaged, "    ReDim contributions(0 To iterations - 1)",
                    "    ReDim contributions(0 To iterations - 1)\n"
                    "    For iteration = 1 To iterations\n"
                    "        contributions(iteration - 1) = -1#\n"
                    "    Next iteration")
    _refused(damaged, (conformance.test_01_every_single_driver_case_replays_bit_for_bit,
                       conformance.test_02_a_risk_at_probability_zero_contributes_zero_at_every_iteration),
             "the non-occurrence observations were replaced")


def test_12_the_driver_is_found_by_supply_position_rather_than_identity() -> None:
    """Worksheet row order reaches the engine as supply order, and identity that
    moves when an unrelated driver is added is not identity."""
    damaged = _swap(
        _SOURCE,
        "        If target < 0 And prepared(index).PermanentId = permanentId Then\n"
        "            target = index\n        End If",
        "        If target < 0 And index = 0 Then\n            target = index\n        End If")
    _refused(damaged, (conformance.test_10_supply_order_is_not_driver_identity,
                       conformance.test_11_a_canonically_later_driver_does_not_disturb_an_earlier_one,
                       conformance.test_13_replay_refuses_what_it_cannot_honour),
             "the driver was located by position rather than by permanent id")


def test_13_an_unknown_driver_is_replayed_as_the_first_one() -> None:
    damaged = _swap(
        _SOURCE,
        "    If target < 0 Then\n        detail = \"engine: replay was asked for driver \" & permanentId & _\n"
        "                 \", which is not in this model\"\n        Exit Function\n    End If",
        "    If target < 0 Then\n        target = 0\n    End If")
    _refused(damaged, (conformance.test_13_replay_refuses_what_it_cannot_honour,),
             "an unknown driver silently replayed a different one")


# ===========================================================================
# D. THE NORMAL SIMULATION IS UNCHANGED BY THE EXTRACTION
# ===========================================================================
def test_14_the_simulation_stops_reaching_the_shared_routine() -> None:
    """If only the replay used it, the extraction would have created the second
    implementation it was meant to remove."""
    damaged = _swap(
        _SOURCE,
        "            If Not SimEngineContribution(prepared(index), unitCost, True, _\n"
        "                                         prepared(index).Knom, SIM_MEASURE_NOMINAL, _\n"
        "                                         iteration, term, detail) Then Exit Function",
        "            term = unitCost * prepared(index).Quantity * prepared(index).Knom")
    _refused(damaged, (conformance.test_18_both_the_simulation_and_the_replay_call_it,
                       conformance.test_17_the_contribution_arithmetic_has_exactly_one_owner),
             "the simulation stopped calling the shared routine")


def test_15_the_pv_accumulator_is_derived_from_the_nominal_term() -> None:
    damaged = _swap(
        _SOURCE,
        "            If Not SimEngineContribution(prepared(index), unitCost, True, _\n"
        "                                         prepared(index).Kpv, SIM_MEASURE_PV, _\n"
        "                                         iteration, term, detail) Then Exit Function\n"
        "            pvTerm(index) = term",
        "            pvTerm(index) = nominalTerm(index) * prepared(index).Kpv / prepared(index).Knom")
    _refused(damaged, (conformance.test_20_pv_is_not_derived_from_the_nominal_term,
                       conformance.test_18_both_the_simulation_and_the_replay_call_it),
             "PV became a discounted nominal term")
