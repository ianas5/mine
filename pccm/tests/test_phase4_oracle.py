#!/usr/bin/env python3
"""PCCM Phase 4: structural-semantics tests over the pure Python oracle.

These test the TRANSFORMATION RULES, exhaustively and without Excel. They do NOT
test the VBA, and passing here is not evidence that the runtime is correct: only
a clean Windows functional run proves that. What they do prove is that the rules
the VBA must implement are themselves coherent, and they pin the expected values
the Windows harness compares the real workbook against.

Runs standalone or under pytest.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import load_structure_contract  # noqa: E402
from pccm_builder.scenarios import build_scenarios  # noqa: E402
from pccm_builder.stage_b_emit import oracle_limits  # noqa: E402
from pccm_builder.structure_oracle import (  # noqa: E402
    DestructiveImpact,
    Timeline,
    TimelineChange,
    allocate_id,
    apply_change,
    assess,
    highest_sequence,
    prevalidate,
    remap_inflation,
    remap_profiling,
    removed_inflation_values,
    removed_profiling_values,
    sync_rows,
)

STRUCTURE_PATH = PCCM_ROOT / "spec" / "structure_contract.yaml"


def _limits():
    return oracle_limits(load_structure_contract(STRUCTURE_PATH))


# ===========================================================================
# timeline arithmetic
# ===========================================================================
def test_01_last_project_year_and_project_years() -> None:
    t = Timeline(2026, 2028, 3)
    assert t.last_project_year == 2030
    assert t.project_years == [2028, 2029, 2030]


def test_02_inflation_span_starts_the_year_after_the_base_year() -> None:
    t = Timeline(2026, 2028, 3)
    assert t.inflation_years == [2027, 2028, 2029, 2030]


def test_03_degenerate_span_is_legitimately_empty() -> None:
    """Base = Start and Duration = 1 requires no escalation assumption at all."""
    t = Timeline(2035, 2035, 1)
    assert t.inflation_years == []
    assert t.has_empty_inflation_span is True
    assert t.is_complete is True, "an empty span is a valid model, not an incomplete one"


def test_04_blank_timeline_produces_no_years() -> None:
    t = Timeline(None, None, None)
    assert t.is_blank and not t.is_complete
    assert t.project_years == [] and t.inflation_years == []
    assert t.last_project_year is None


# ===========================================================================
# prevalidation
# ===========================================================================
def test_05_a_valid_triple_passes() -> None:
    assert prevalidate(Timeline(2026, 2028, 3), _limits()) == []


def test_06_blank_values_are_rejected() -> None:
    problems = prevalidate(Timeline(None, 2028, 3), _limits())
    assert any("Base Year" in p for p in problems)


def test_07_non_integer_values_are_rejected() -> None:
    problems = prevalidate(Timeline(2026.5, 2028, 3), _limits())
    assert any("whole number" in p for p in problems)


def test_08_duration_below_one_is_rejected() -> None:
    assert any("at least 1" in p for p in prevalidate(Timeline(2026, 2028, 0), _limits()))


def test_09_base_year_after_start_year_is_rejected() -> None:
    problems = prevalidate(Timeline(2030, 2028, 3), _limits())
    assert any("later than" in p for p in problems)


def test_10_years_outside_the_supported_window_are_rejected() -> None:
    limits = _limits()
    assert prevalidate(Timeline(1899, 2028, 3), limits)
    assert prevalidate(Timeline(2026, 2201, 3), limits)


def test_11_last_project_year_beyond_the_boundary_is_rejected() -> None:
    limits = _limits()
    problems = prevalidate(Timeline(2026, 2199, 10), limits)
    assert any("structural year boundary" in p for p in problems)


def test_12_twenty_five_years_is_not_a_cap() -> None:
    """25 years is an Architecture benchmark target, never a business maximum."""
    limits = _limits()
    for duration in (26, 40, 75, 120):
        assert prevalidate(Timeline(2026, 2030, duration), limits) == [], (
            f"a {duration}-year project must be a legitimate model"
        )


def test_13_the_generation_guard_is_the_year_window_width() -> None:
    limits = _limits()
    assert limits.max_generated_year_columns == limits.max_year - limits.min_year + 1


def test_14_prevalidation_reports_every_problem_at_once() -> None:
    problems = prevalidate(Timeline(1800, 1700, 0), _limits())
    assert len(problems) >= 3, f"expected several failures, got {problems}"


# ===========================================================================
# profiling: anchored by project-year index
# ===========================================================================
def test_15_growth_appends_zero_and_preserves_every_existing_value() -> None:
    rows = {"CL-001": [0.2, 0.5, 0.3]}
    grown = remap_profiling(rows, 5)
    assert grown["CL-001"] == [0.2, 0.5, 0.3, 0.0, 0.0]
    assert sum(grown["CL-001"]) == 1.0, "a row that totalled 100% still totals 100%"


def test_16_growth_never_redistributes_or_normalises() -> None:
    rows = {"CL-001": [0.6, 0.6]}   # deliberately not 100%
    assert remap_profiling(rows, 4)["CL-001"] == [0.6, 0.6, 0.0, 0.0]


def test_17_shrink_truncates_from_the_tail_only() -> None:
    rows = {"CL-001": [0.2, 0.5, 0.3]}
    assert remap_profiling(rows, 2)["CL-001"] == [0.2, 0.5]


def test_18_a_start_year_shift_moves_no_profiling_value() -> None:
    """Same duration, different calendar labels: positions are untouched."""
    rows = {"CL-001": [0.2, 0.5, 0.3]}
    before = Timeline(2026, 2028, 3)
    after = Timeline(2026, 2031, 3)
    change = TimelineChange(old=before, new=after)
    assert change.headers_relabelled is True
    assert change.removed_project_year_indices == []
    reshaped, _ = apply_change(change, {"cost": rows}, {})
    assert reshaped["cost"]["CL-001"] == [0.2, 0.5, 0.3]
    assert before.project_years != after.project_years, "only the headers change"


def test_19_only_non_zero_removals_count_as_destruction() -> None:
    rows = {"CL-001": [0.5, 0.5, 0.0, 0.0]}
    assert removed_profiling_values(rows, 2) == []
    rows = {"CL-001": [0.5, 0.3, 0.2, 0.0]}
    assert removed_profiling_values(rows, 2) == [("CL-001", 3, 0.2)]


def test_20_destructive_indices_are_one_based_and_match_the_user_view() -> None:
    rows = {"CL-001": [0.1, 0.2, 0.3, 0.4]}
    losses = removed_profiling_values(rows, 1)
    assert [index for _key, index, _value in losses] == [2, 3, 4]


# ===========================================================================
# inflation: anchored by calendar year
# ===========================================================================
def test_21_rates_survive_by_calendar_year_not_by_column_index() -> None:
    rates = {"Local": {2027: 0.03, 2028: 0.04, 2029: 0.05}}
    remapped = remap_inflation(rates, [2029, 2030, 2031])
    assert remapped["Local"][2029] == 0.05, "2029 keeps its rate although its index changed"
    assert remapped["Local"][2030] is None and remapped["Local"][2031] is None


def test_22_newly_required_years_arrive_blank_never_zero() -> None:
    rates = {"Local": {2027: 0.03}}
    remapped = remap_inflation(rates, [2026, 2027, 2028])
    assert remapped["Local"][2026] is None, "a rate the user never entered must not be invented"
    assert remapped["Local"][2028] is None
    assert 0 not in remapped["Local"].values()


def test_23_only_non_blank_rates_leaving_the_span_count_as_destruction() -> None:
    rates = {"Local": {2027: None, 2028: 0.04}}
    assert removed_inflation_values(rates, [2028]) == []
    assert removed_inflation_values(rates, [2029]) == [("Local", 2028, 0.04)]


def test_24_an_empty_span_removes_every_rate() -> None:
    rates = {"Local": {2027: 0.03, 2028: 0.04}}
    assert len(removed_inflation_values(rates, [])) == 2
    assert remap_inflation(rates, []) == {"Local": {}}


def test_25_base_year_moved_earlier_adds_blank_years_at_the_early_end() -> None:
    old = Timeline(2026, 2030, 3)
    new = Timeline(2024, 2030, 3)
    change = TimelineChange(old=old, new=new)
    assert change.added_inflation_years == [2025, 2026]
    assert change.removed_inflation_years == []


def test_26_base_year_moved_later_removes_years_at_the_early_end() -> None:
    old = Timeline(2024, 2030, 3)
    new = Timeline(2029, 2030, 3)
    change = TimelineChange(old=old, new=new)
    assert change.removed_inflation_years == [2025, 2026, 2027, 2028, 2029]
    assert change.added_inflation_years == []


def test_27_base_year_change_does_not_touch_profiling() -> None:
    rows = {"CL-001": [0.4, 0.6]}
    change = TimelineChange(old=Timeline(2026, 2030, 2), new=Timeline(2022, 2030, 2))
    assert change.removed_project_year_indices == []
    reshaped, _ = apply_change(change, {"cost": rows}, {})
    assert reshaped["cost"]["CL-001"] == [0.4, 0.6]


# ===========================================================================
# one combined transition
# ===========================================================================
def test_28_a_three_way_change_is_assessed_as_one_delta() -> None:
    old = Timeline(2026, 2028, 4)
    new = Timeline(2031, 2033, 4)
    change = TimelineChange(old=old, new=new)
    # Duration is unchanged, so nothing is removed from profiling even though both
    # the base year and the start year moved a long way.
    assert change.removed_project_year_indices == []
    assert change.removed_inflation_years == [2027, 2028, 2029, 2030, 2031]


def test_29_sequential_partial_transitions_would_over_report_destruction() -> None:
    """Why the delta is computed once: doing it in steps invents losses.

    Shrinking to 2 and then growing back to 4 destroys two project years. The
    combined old -> new delta, with duration unchanged at 4, destroys nothing.
    """
    rows = {"CL-001": [0.1, 0.2, 0.3, 0.4]}
    combined = TimelineChange(old=Timeline(2026, 2028, 4), new=Timeline(2031, 2033, 4))
    combined_impact = assess(combined, {"cost": rows}, {})
    assert combined_impact.profiling_loss_count == 0

    partial = TimelineChange(old=Timeline(2026, 2028, 4), new=Timeline(2026, 2028, 2))
    partial_impact = assess(partial, {"cost": rows}, {})
    assert partial_impact.profiling_loss_count == 2


def test_30_impact_is_not_destructive_when_only_zeros_and_blanks_are_removed() -> None:
    change = TimelineChange(old=Timeline(2026, 2028, 4), new=Timeline(2026, 2028, 2))
    impact = assess(change, {"cost": {"CL-001": [0.5, 0.5, 0.0, 0.0]}}, {"Local": {2029: None}})
    assert impact.is_destructive is False
    assert impact.removed_project_year_indices == [3, 4]


def test_31_impact_reports_counts_and_representative_ids() -> None:
    rows = {f"CL-{n:03d}": [0.5, 0.5] for n in range(1, 9)}
    change = TimelineChange(old=Timeline(2026, 2028, 2), new=Timeline(2026, 2028, 1))
    impact = assess(change, {"cost": rows}, {})
    assert impact.is_destructive is True
    assert impact.profiling_loss_count == 8
    assert len(impact.affected_ids()) == 5, "the prompt shows a representative sample"


def test_32_impact_is_computed_before_anything_moves() -> None:
    """assess() is pure: it cannot mutate the grids it inspects."""
    rows = {"CL-001": [0.5, 0.5]}
    snapshot = {k: list(v) for k, v in rows.items()}
    change = TimelineChange(old=Timeline(2026, 2028, 2), new=Timeline(2026, 2028, 1))
    result = assess(change, {"cost": rows}, {})
    assert isinstance(result, DestructiveImpact)
    assert rows == snapshot, "cancelling needs no rollback because nothing has moved"


# ===========================================================================
# permanent identity
# ===========================================================================
def test_33_ids_are_allocated_from_the_counter_alone() -> None:
    counter = 0
    issued = []
    for _ in range(3):
        new_id, counter = allocate_id(counter, "CL-", 3)
        issued.append(new_id)
    assert issued == ["CL-001", "CL-002", "CL-003"]
    assert counter == 3


def test_34_deletion_never_causes_reuse() -> None:
    """Add, add, add, delete CL-002, add -> CL-004. Never CL-002 again."""
    counter = 0
    for _ in range(3):
        _, counter = allocate_id(counter, "CL-", 3)
    # A deletion removes a row. It does not touch the counter.
    surviving = ["CL-001", "CL-003"]
    next_id, counter = allocate_id(counter, "CL-", 3)
    assert next_id == "CL-004"
    assert next_id not in surviving and next_id != "CL-002"


def test_35_pad_width_is_a_floor_not_a_ceiling() -> None:
    assert allocate_id(998, "CL-", 3)[0] == "CL-999"
    assert allocate_id(999, "CL-", 3)[0] == "CL-1000"
    assert allocate_id(99999, "CL-", 3)[0] == "CL-100000"


def test_36_cost_and_risk_sequences_are_independent() -> None:
    cost_counter, risk_counter = 0, 0
    for _ in range(4):
        _, cost_counter = allocate_id(cost_counter, "CL-", 3)
    first_risk, risk_counter = allocate_id(risk_counter, "R-", 3)
    assert first_risk == "R-001", "the risk sequence is untouched by four cost allocations"


def test_37_highest_sequence_ignores_foreign_and_blank_ids() -> None:
    assert highest_sequence(["CL-001", "CL-017", "", None, "R-999", "junk"], "CL-") == 17
    assert highest_sequence([], "CL-") == 0


def test_38_a_counter_behind_its_highest_issued_id_would_reuse_one() -> None:
    """The revalidation rule: counter >= highest issued, never counter == row count."""
    surviving = ["CL-001", "CL-004"]
    highest = highest_sequence(surviving, "CL-")
    assert highest == 4
    assert len(surviving) == 2, "a correct counter is deliberately ahead of the row count"
    reused, _ = allocate_id(len(surviving), "CL-", 3)
    assert reused == "CL-003", "counting rows instead of reading the counter reissues an ID"


# ===========================================================================
# row synchronisation
# ===========================================================================
def test_39_sync_follows_permanent_identity_not_row_order() -> None:
    existing = {"CL-001": "a", "CL-002": "b"}
    ordered, added, removed = sync_rows(["CL-002", "CL-001"], existing)
    assert ordered == ["CL-002", "CL-001"]
    assert added == [] and removed == [], "reordering transfers nothing between rows"


def test_40_sync_reports_additions_and_removals() -> None:
    ordered, added, removed = sync_rows(["CL-001", "CL-003"], {"CL-001": "a", "CL-002": "b"})
    assert ordered == ["CL-001", "CL-003"]
    assert added == ["CL-003"] and removed == ["CL-002"]


def test_41_blank_ids_are_not_rows() -> None:
    ordered, _added, _removed = sync_rows(["CL-001", "", None, "CL-002"], {})
    assert ordered == ["CL-001", "CL-002"]


def test_42_duplicate_identifiers_are_a_hard_error() -> None:
    try:
        sync_rows(["CL-001", "CL-001"], {})
    except ValueError as error:
        assert "duplicate" in str(error)
        return
    raise AssertionError("duplicate permanent identifiers were silently accepted")


# ===========================================================================
# the emitted fixture must still be the oracle's own output
# ===========================================================================
def test_43_emitted_scenarios_match_the_oracle() -> None:
    """The Windows harness reads this fixture instead of hard-coding years.

    If it could drift from the oracle, the harness would be asserting numbers no
    Linux test had ever checked.
    """
    emitted = PCCM_ROOT / "build" / "phase4_scenarios.json"
    if not emitted.is_file():
        return  # build/ is git-ignored; the build test covers regeneration
    expected = build_scenarios(_limits())
    actual = json.loads(emitted.read_text(encoding="utf-8"))
    assert actual == json.loads(json.dumps(expected)), (
        "build/phase4_scenarios.json has drifted from the structural oracle; rebuild"
    )


def test_44_scenario_matrix_covers_every_required_transition() -> None:
    steps = build_scenarios(_limits())["steps"]
    keys = [s["key"] for s in steps]
    for required in (
        "D_first_apply",
        "E_duration_increase",
        "F_start_year_shift",
        "G_duration_decrease_cancel",
        "G_duration_decrease_accept",
        "H_base_year_earlier",
        "H_base_year_later",
        "I_combined_change",
        "J_degenerate_span",
    ):
        assert required in keys, f"the functional matrix is missing {required}"


def test_45_the_cancelled_step_expects_the_previous_shape() -> None:
    steps = {s["key"]: s for s in build_scenarios(_limits())["steps"]}
    cancelled = steps["G_duration_decrease_cancel"]
    accepted = steps["G_duration_decrease_accept"]
    assert cancelled["confirm"] is False
    assert cancelled["expect"]["applied"]["duration"] == 5, (
        "a cancelled shrink must leave the previous duration in place"
    )
    assert accepted["expect"]["applied"]["duration"] == 3


def test_46_the_degenerate_step_expects_no_inflation_years() -> None:
    steps = {s["key"]: s for s in build_scenarios(_limits())["steps"]}
    degenerate = steps["J_degenerate_span"]
    assert degenerate["expect"]["inflation_years"] == []
    assert degenerate["expect"]["inflation_span_is_empty"] is True


def _run_all() -> int:
    tests = sorted(
        (name, fn) for name, fn in globals().items()
        if name.startswith("test_") and callable(fn)
    )
    failures = 0
    print("PCCM Phase 4 structural oracle tests")
    print("=" * 70)
    for name, fn in tests:
        try:
            fn()
        except AssertionError as error:
            failures += 1
            print(f"  [FAIL] {name}\n         {error}")
        except Exception as error:  # noqa: BLE001
            failures += 1
            print(f"  [ERROR] {name}\n          {type(error).__name__}: {error}")
        else:
            print(f"  [PASS] {name}")
    print("=" * 70)
    print(f"  {len(tests) - failures} passed, {failures} failed")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(_run_all())
