#!/usr/bin/env python3
"""PCCM Phase 4 negative tests.

A malformed structure contract, or one that disagrees with the other three
specifications, must fail loudly rather than produce a workbook whose structural
runtime is quietly wrong.

Runs standalone or under pytest.
"""

from __future__ import annotations

import copy
import sys
import tempfile
from pathlib import Path
from typing import Any, Callable

import yaml

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import (  # noqa: E402
    StructureContractError,
    build_workbook,
    load_contract,
    load_driver_contract,
    load_spec,
    load_structure_contract,
)
from pccm_builder.structure_loader import validate_structure_against  # noqa: E402

SPEC_PATH = PCCM_ROOT / "spec" / "workbook.yaml"
CONTRACT_PATH = PCCM_ROOT / "spec" / "input_contract.yaml"
DRIVERS_PATH = PCCM_ROOT / "spec" / "driver_contract.yaml"
STRUCTURE_PATH = PCCM_ROOT / "spec" / "structure_contract.yaml"


def _base() -> dict[str, Any]:
    with STRUCTURE_PATH.open("r", encoding="utf-8") as handle:
        return yaml.safe_load(handle)


def _write(data: dict[str, Any], tmp: str, name: str = "broken.yaml") -> Path:
    path = Path(tmp) / name
    with path.open("w", encoding="utf-8") as handle:
        yaml.safe_dump(data, handle, sort_keys=False)
    return path


def _rejected(mutate: Callable[[dict[str, Any]], None], reason: str) -> None:
    """The mutated contract must fail at load time."""
    data = copy.deepcopy(_base())
    mutate(data)
    with tempfile.TemporaryDirectory(prefix="pccm-badstructure-") as tmp:
        path = _write(data, tmp)
        try:
            load_structure_contract(path)
        except StructureContractError:
            return
        except Exception as error:  # noqa: BLE001
            raise AssertionError(
                f"{reason}: raised {type(error).__name__} instead of StructureContractError"
            ) from error
    raise AssertionError(f"{reason}: an invalid structure contract was silently accepted")


def _rejected_cross(mutate: Callable[[dict[str, Any]], None], reason: str) -> None:
    """The mutated contract must fail cross-contract validation."""
    data = copy.deepcopy(_base())
    mutate(data)
    with tempfile.TemporaryDirectory(prefix="pccm-badstructure-") as tmp:
        structure = load_structure_contract(_write(data, tmp))
        try:
            validate_structure_against(
                structure, load_contract(CONTRACT_PATH), load_driver_contract(DRIVERS_PATH)
            )
        except StructureContractError:
            return
    raise AssertionError(f"{reason}: silently accepted")


def _field(data: dict[str, Any], group: str, key: str) -> dict[str, Any]:
    return next(f for f in data["timeline"][group] if f["key"] == key)


def _counter(data: dict[str, Any], key: str) -> dict[str, Any]:
    return next(c for c in data["identity"]["counters"] if c["key"] == key)


# ===========================================================================
# limits
# ===========================================================================
def test_rejects_a_generation_guard_other_than_the_locked_two_hundred() -> None:
    """Architecture Lock Revision B: generated column count > 200 = ERROR.

    301 is the specific wrong value this guard previously held, produced by deriving
    it from the calendar-year window. It is rejected explicitly.
    """
    for wrong in (301, 25, 199, 201, 1000):
        _rejected(
            lambda d, value=wrong: d["limits"].__setitem__("max_generated_year_columns", value),
            f"a project-year generation guard of {wrong}",
        )


def test_the_calendar_window_is_not_forced_to_match_the_column_guard() -> None:
    """The two protections are independent, and widening one must not require the other."""
    structure = load_structure_contract(STRUCTURE_PATH)
    assert structure.limits.max_generated_year_columns == 200
    assert structure.limits.max_year - structure.limits.min_year + 1 == 301
    assert structure.limits.max_generated_year_columns != 301


def test_rejects_an_inverted_year_window() -> None:
    _rejected(
        lambda d: d["limits"].__setitem__("min_year", 2500),
        "min_year above max_year",
    )


def test_rejects_a_year_window_wider_than_the_excel_grid() -> None:
    def widen(d: dict[str, Any]) -> None:
        d["limits"]["min_year"] = 1
        d["limits"]["max_year"] = 20000
        d["limits"]["max_generated_year_columns"] = 20000

    _rejected(widen, "more generated year columns than Excel has columns")


# ===========================================================================
# entered vs applied
# ===========================================================================
def test_rejects_an_alias_pointing_at_an_unknown_input() -> None:
    _rejected_cross(
        lambda d: d["timeline"]["entered_aliases"][0].__setitem__("input_key", "nope"),
        "an entered alias over an input the contract does not declare",
    )


def test_rejects_an_alias_over_a_model_controlled_input() -> None:
    _rejected_cross(
        lambda d: d["timeline"]["entered_aliases"][0].__setitem__(
            "input_key", "reporting_currency"
        ),
        "an entered alias over a model-controlled input",
    )


def test_rejects_an_alias_over_a_non_integer_input() -> None:
    _rejected_cross(
        lambda d: d["timeline"]["entered_aliases"][0].__setitem__("input_key", "project_name"),
        "an entered alias over a text input",
    )


def test_rejects_a_structural_name_that_shadows_an_accepted_input_name() -> None:
    """Two guards catch this, and the naming convention catches it first.

    The nm* convention makes a collision with inp*, lst* or tbl* impossible by
    construction, so the load-time check fires before the cross-contract one. The
    cross-contract guard remains as the backstop for the day another contract
    introduces an nm* name of its own.
    """
    _rejected(
        lambda d: _field(d, "applied", "applied_base_year").__setitem__(
            "defined_name", "inpBaseYear"
        ),
        "a structural name shadowing an accepted inp* name",
    )


def test_rejects_a_structural_name_outside_the_nm_convention() -> None:
    _rejected(
        lambda d: _field(d, "applied", "applied_base_year").__setitem__(
            "defined_name", "AppliedBaseYear"
        ),
        "a structural name that does not use the nm prefix",
    )


def test_rejects_duplicate_structural_names() -> None:
    _rejected(
        lambda d: _field(d, "applied", "applied_start_year").__setitem__(
            "defined_name", "nmBaseYear_Applied"
        ),
        "two structural fields sharing a defined name",
    )


def test_rejects_a_formula_on_an_applied_state_cell() -> None:
    """The applied triple is model-written state; a formula would make it derived."""
    _rejected(
        lambda d: _field(d, "applied", "applied_duration").__setitem__(
            "formula", "=nmDuration_Entered"
        ),
        "a formula on an applied-timeline cell",
    )


def test_rejects_a_derived_field_without_a_formula() -> None:
    _rejected(
        lambda d: _field(d, "derived", "applied_last_year").__setitem__("formula", None),
        "a derived structural field with no formula",
    )


def test_rejects_a_derived_formula_referencing_an_unknown_name() -> None:
    _rejected(
        lambda d: _field(d, "derived", "applied_last_year").__setitem__(
            "formula", '=IF(nmNotAThing="","",1)'
        ),
        "a derived formula over an undeclared structural name",
    )


def test_rejects_an_applied_cell_landing_on_the_phase2_setup_area() -> None:
    _rejected_cross(
        lambda d: _field(d, "applied", "applied_base_year").__setitem__("cell", "C9"),
        "an applied cell overwriting an accepted Phase-2 input",
    )


# ===========================================================================
# structural state indicator
# ===========================================================================
def test_rejects_a_state_formula_that_ignores_an_entered_value() -> None:
    def drop(d: dict[str, Any]) -> None:
        d["structural_state"]["formula"] = (
            '=IF(AND(nmBaseYear_Entered=nmBaseYear_Applied,'
            'nmStartYear_Entered=nmStartYear_Applied),'
            '"Timeline current","STRUCTURE CHANGE PENDING")'
        )

    _rejected(drop, "a pending flag that never compares the duration")


def test_rejects_a_state_label_that_does_not_appear_in_its_formula() -> None:
    _rejected(
        lambda d: d["structural_state"]["labels"].__setitem__("pending", "CHANGES PENDING"),
        "a declared state label the formula never emits",
    )


def test_rejects_a_grid_message_formula_that_omits_the_declared_text() -> None:
    _rejected(
        lambda d: d["state_messages"].__setitem__(
            "not_applied", "Something else entirely."
        ),
        "a state message the grid formula never emits",
    )


def test_rejects_a_missing_empty_span_message() -> None:
    _rejected(
        lambda d: d["state_messages"].pop("inflation_empty_span"),
        "no explanatory message for the legitimate empty inflation span",
    )


# ===========================================================================
# identity
# ===========================================================================
def test_rejects_counters_sharing_an_id_prefix() -> None:
    _rejected(
        lambda d: _counter(d, "risk").__setitem__("prefix", "CL-"),
        "two registers issuing identifiers from the same prefix",
    )


def test_rejects_two_counters_claiming_one_register() -> None:
    _rejected(
        lambda d: _counter(d, "risk").__setitem__("driver_register", "cost_lines"),
        "two counters claiming the same driver register",
    )


def test_rejects_a_counter_for_an_unknown_register() -> None:
    _rejected_cross(
        lambda d: _counter(d, "risk").__setitem__("driver_register", "nope"),
        "a counter over a register the driver contract does not declare",
    )


def test_rejects_a_pattern_that_caps_the_id_range() -> None:
    """Pad width is a display floor. A pattern that rejects CL-1000 is a cap."""
    _rejected(
        lambda d: _counter(d, "cost_line").__setitem__("pattern", "^CL-[0-9]{3}$"),
        "an ID pattern that imposes an artificial maximum",
    )


def test_rejects_a_pattern_that_its_own_prefix_fails() -> None:
    _rejected(
        lambda d: _counter(d, "cost_line").__setitem__("pattern", "^XX-[0-9]{3,}$"),
        "a counter whose allocations do not match its own pattern",
    )


def test_rejects_a_prefix_without_a_separator() -> None:
    _rejected(
        lambda d: _counter(d, "cost_line").__setitem__("prefix", "CL"),
        "a prefix that cannot be separated from its sequence",
    )


def test_rejects_a_negative_initial_counter() -> None:
    _rejected(
        lambda d: _counter(d, "risk").__setitem__("initial", -1),
        "a negative starting counter",
    )


# ===========================================================================
# grids
# ===========================================================================
def test_rejects_a_profiling_grid_whose_capacity_leaves_its_register() -> None:
    _rejected_cross(
        lambda d: d["grids"]["cost_profiling"].__setitem__("reserved_rows", 12),
        "a profiling grid reserving fewer rows than its driver register",
    )


def test_rejects_an_inflation_grid_whose_capacity_leaves_the_profile_master() -> None:
    _rejected_cross(
        lambda d: d["grids"]["inflation"].__setitem__("reserved_rows", 3),
        "an inflation grid reserving fewer rows than the Config profile master",
    )


def test_rejects_a_profiling_grid_over_an_unknown_register() -> None:
    _rejected_cross(
        lambda d: d["grids"]["cost_profiling"].__setitem__("driver_register", "nope"),
        "a profiling grid mirroring a register that does not exist",
    )


def test_rejects_a_trace_column_over_an_unknown_driver_column() -> None:
    _rejected_cross(
        lambda d: d["grids"]["cost_profiling"]["fixed_columns"][1].__setitem__(
            "source_driver_column", "nope"
        ),
        "a trace column over a driver column that does not exist",
    )


def test_rejects_a_profiling_grid_not_keyed_on_the_permanent_identifier() -> None:
    _rejected(
        lambda d: d["grids"]["cost_profiling"].__setitem__("key_driver_column", "description"),
        "a profiling grid keyed on something other than the permanent ID",
    )


def test_rejects_a_profiling_column_that_originates_a_driver_attribute() -> None:
    _rejected(
        lambda d: d["grids"]["risk_profiling"]["fixed_columns"][1].pop("source_driver_column"),
        "a profiling column that is not a trace copy",
    )


def test_rejects_a_zero_seeded_inflation_year() -> None:
    """A blank rate and a 0% rate are different assumptions and must stay so."""
    _rejected(
        lambda d: d["grids"]["inflation"]["year_column"].__setitem__("initial_value", 0),
        "a newly required inflation year fabricated as zero",
    )


def test_rejects_a_blank_seeded_profiling_year() -> None:
    _rejected(
        lambda d: d["grids"]["cost_profiling"]["year_column"].__setitem__("initial_value", None),
        "a new project-year profiling cell left blank instead of 0%",
    )


def test_rejects_a_grid_table_name_already_used_by_another_contract() -> None:
    _rejected_cross(
        lambda d: d["grids"]["cost_profiling"].__setitem__("table_name", "tblCostLines"),
        "a grid table name colliding with a driver register",
    )


def test_rejects_two_grids_on_one_sheet() -> None:
    _rejected(
        lambda d: d["grids"]["risk_profiling"].__setitem__("sheet", "Cost Profiling"),
        "two structural grids sharing a sheet",
    )


def test_rejects_a_grid_overlapping_its_own_heading() -> None:
    _rejected(
        lambda d: d["grids"]["cost_profiling"].__setitem__("section_row", 14),
        "a section heading inside the grid body",
    )


def test_rejects_zero_reserved_rows() -> None:
    _rejected(
        lambda d: d["grids"]["inflation"].__setitem__("reserved_rows", 0),
        "a grid with no reserved capacity",
    )


def test_rejects_a_grid_beyond_the_excel_row_limit() -> None:
    _rejected(
        lambda d: d["grids"]["cost_profiling"].__setitem__("header_row", 1048570),
        "a grid whose reserved rows run past the last Excel row",
    )


def test_rejects_a_malformed_grid_table_name() -> None:
    _rejected(
        lambda d: d["grids"]["inflation"].__setitem__("table_name", "Inflation"),
        "a grid table name not matching tbl<PascalCase>",
    )


# ===========================================================================
# buttons and VBA
# ===========================================================================
def test_rejects_a_button_bound_to_an_undeclared_entry_point() -> None:
    _rejected(
        lambda d: d["buttons"]["definitions"][0].__setitem__("entry_point", "PCCM_Nope"),
        "a button bound to a macro the contract does not declare",
    )


def test_rejects_an_entry_point_with_no_button() -> None:
    _rejected(
        lambda d: d["vba"]["entry_points"].append("PCCM_Orphan"),
        "a declared entry point that no button invokes",
    )


def test_rejects_two_buttons_sharing_a_macro() -> None:
    _rejected(
        lambda d: d["buttons"]["definitions"][1].__setitem__(
            "entry_point", d["buttons"]["definitions"][0]["entry_point"]
        ),
        "two buttons bound to one macro",
    )


def test_rejects_duplicate_button_shape_names() -> None:
    _rejected(
        lambda d: d["buttons"]["definitions"][1].__setitem__(
            "shape_name", d["buttons"]["definitions"][0]["shape_name"]
        ),
        "two buttons sharing a shape name",
    )


def test_rejects_a_malformed_entry_point_name() -> None:
    def rename(d: dict[str, Any]) -> None:
        d["vba"]["entry_points"][0] = "ApplyTimeline"
        d["buttons"]["definitions"][0]["entry_point"] = "ApplyTimeline"

    _rejected(rename, "an entry point outside the PCCM_ convention")


def test_rejects_a_module_claiming_to_be_generated_that_the_builder_does_not_emit() -> None:
    """The surviving half of "a second generated module is refused".

    The builder now emits two: modConstants and modCalcContract. What must still be
    refused is a module claiming to be generated when nothing generates it - the
    contract and the builder would then disagree about what the build produces,
    and the disagreement would only surface on Windows.
    """
    _rejected(
        lambda d: d["vba"]["modules"][1].__setitem__("generated", True),
        "a hand-written module marked generated",
    )


def test_rejects_dropping_the_generated_flag_from_the_primary_generated_module() -> None:
    """`vba.generated_module` must still be declared generated.

    This is the other half of the original assertion, and it is the half that
    matters most: modConstants is emitted whatever the contract says, so a contract
    that stopped calling it generated would invite a hand-written copy.
    """
    _rejected(
        lambda d: d["vba"]["modules"][0].__setitem__("generated", False),
        "the primary generated module no longer declared generated",
    )


def test_rejects_a_duplicate_module_name() -> None:
    """Also carried forward: two entries for one module is still refused."""
    _rejected(
        lambda d: d["vba"]["modules"][1].__setitem__("name", d["vba"]["modules"][2]["name"]),
        "duplicate VBA module names",
    )


def test_the_generated_module_inventory_is_exactly_what_the_builder_emits() -> None:
    """The current generated set, stated once and asserted in both directions."""
    from pccm_builder.structure_loader import GENERATED_MODULES

    assert sorted(GENERATED_MODULES) == ["modCalcContract", "modConstants"]
    declared = [m["name"] for m in _base()["vba"]["modules"] if m.get("generated")]
    assert sorted(declared) == sorted(GENERATED_MODULES)


def test_rejects_a_malformed_module_name() -> None:
    _rejected(
        lambda d: d["vba"]["modules"][1].__setitem__("name", "TimelineStuff"),
        "a module outside the mod<PascalCase> convention",
    )


def test_rejects_an_empty_forbidden_construct_list() -> None:
    _rejected(
        lambda d: d["vba"].__setitem__("forbidden_constructs", []),
        "no guard against later-phase functionality leaking into Phase 4",
    )


def test_rejects_a_button_on_an_unknown_sheet() -> None:
    data = copy.deepcopy(_base())
    data["buttons"]["definitions"][0]["sheet"] = "Nowhere"
    with tempfile.TemporaryDirectory(prefix="pccm-badstructure-") as tmp:
        structure = load_structure_contract(_write(data, tmp))
        try:
            build_workbook(
                load_spec(SPEC_PATH),
                load_contract(CONTRACT_PATH),
                load_driver_contract(DRIVERS_PATH),
                structure,
            )
        except RuntimeError:
            return
    raise AssertionError("a button on an unknown sheet was accepted")


# ===========================================================================
# manifest agreement
# ===========================================================================
def test_rejects_a_grid_on_a_sheet_the_manifest_does_not_mark_structural() -> None:
    data = copy.deepcopy(_base())
    data["grids"]["inflation"]["sheet"] = "Methodology"
    with tempfile.TemporaryDirectory(prefix="pccm-badstructure-") as tmp:
        structure = load_structure_contract(_write(data, tmp))
        try:
            build_workbook(
                load_spec(SPEC_PATH),
                load_contract(CONTRACT_PATH),
                load_driver_contract(DRIVERS_PATH),
                structure,
            )
        except RuntimeError as error:
            assert "structure-bodied sheets" in str(error)
            return
    raise AssertionError("manifest and structure contract were allowed to disagree")


def test_rejects_an_applied_block_on_the_wrong_setup_sheet() -> None:
    data = copy.deepcopy(_base())
    data["timeline"]["applied_block"]["sheet"] = "Config"
    data["structural_state"]["sheet"] = "Config"
    with tempfile.TemporaryDirectory(prefix="pccm-badstructure-") as tmp:
        structure = load_structure_contract(_write(data, tmp))
        try:
            build_workbook(
                load_spec(SPEC_PATH),
                load_contract(CONTRACT_PATH),
                load_driver_contract(DRIVERS_PATH),
                structure,
            )
        except RuntimeError as error:
            assert "Setup sheet" in str(error)
            return
    raise AssertionError("the applied timeline was allowed onto the wrong sheet")


def test_rejects_a_missing_structure_contract_file() -> None:
    try:
        load_structure_contract(PCCM_ROOT / "spec" / "does_not_exist.yaml")
    except StructureContractError:
        return
    raise AssertionError("a missing structure contract was silently accepted")


def test_the_real_structure_contract_loads_and_agrees_with_the_others() -> None:
    structure = load_structure_contract(STRUCTURE_PATH)
    assert len(structure.all_grids) == 3
    assert len(structure.counters) == 2
    assert len(structure.buttons) == 5
    validate_structure_against(
        structure, load_contract(CONTRACT_PATH), load_driver_contract(DRIVERS_PATH)
    )


def _run_all() -> int:
    tests = sorted(
        (name, fn) for name, fn in globals().items()
        if name.startswith("test_") and callable(fn)
    )
    failures = 0
    print("PCCM Phase 4 structure contract negative tests")
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
