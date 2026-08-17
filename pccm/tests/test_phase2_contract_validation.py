#!/usr/bin/env python3
"""PCCM Phase 2 negative tests: a malformed input contract must fail loudly.

Each case mutates a copy of the real contract into an invalid state and asserts
that loading it raises ContractError. The builder must never repair a bad
specification silently.

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

from pccm_builder import ContractError, load_contract  # noqa: E402

CONTRACT_PATH = PCCM_ROOT / "spec" / "input_contract.yaml"


def _base() -> dict[str, Any]:
    with CONTRACT_PATH.open("r", encoding="utf-8") as handle:
        return yaml.safe_load(handle)


def _rejected(mutate: Callable[[dict[str, Any]], None], reason: str) -> None:
    data = copy.deepcopy(_base())
    mutate(data)
    with tempfile.TemporaryDirectory(prefix="pccm-badcontract-") as tmp:
        path = Path(tmp) / "broken.yaml"
        with path.open("w", encoding="utf-8") as handle:
            yaml.safe_dump(data, handle, sort_keys=False)
        try:
            load_contract(path)
        except ContractError:
            return
        except Exception as error:  # noqa: BLE001
            raise AssertionError(
                f"{reason}: raised {type(error).__name__} instead of ContractError"
            ) from error
    raise AssertionError(f"{reason}: invalid contract was silently accepted")


def test_rejects_duplicate_defined_name() -> None:
    _rejected(
        lambda d: d["inputs"]["base_year"].__setitem__("defined_name", "inpProjectName"),
        "duplicate input defined name",
    )


def test_rejects_defined_name_collision_between_input_and_list() -> None:
    _rejected(
        lambda d: d["config_tables"][0].__setitem__("defined_name", "lstShared")
        or d["config_tables"][1].__setitem__("defined_name", "lstShared"),
        "duplicate list defined name",
    )


def test_rejects_duplicate_table_name() -> None:
    _rejected(
        lambda d: d["config_tables"][1].__setitem__("table_name", d["config_tables"][0]["table_name"]),
        "duplicate table name",
    )


def test_rejects_malformed_defined_name() -> None:
    _rejected(
        lambda d: d["inputs"]["base_year"].__setitem__("defined_name", "baseYear"),
        "defined name not matching the inp<PascalCase> convention",
    )


def test_rejects_malformed_table_name() -> None:
    _rejected(
        lambda d: d["config_tables"][0].__setitem__("table_name", "Categories"),
        "table name not matching the tbl<PascalCase> convention",
    )


def test_rejects_invalid_cell_reference() -> None:
    _rejected(
        lambda d: d["inputs"]["base_year"].__setitem__("cell", "C0"),
        "invalid cell reference",
    )


def test_rejects_cell_collision() -> None:
    _rejected(
        lambda d: d["inputs"]["base_year"].__setitem__(
            "cell", d["inputs"]["project_name"]["cell"]
        ),
        "two inputs claiming the same cell",
    )


def test_rejects_input_overlapping_a_table() -> None:
    _rejected(
        lambda d: d["inputs"]["base_year"].__setitem__("cell", "B28"),
        "an input cell inside the FX table body",
    )


def test_rejects_unknown_validation_source() -> None:
    _rejected(
        lambda d: d["inputs"]["selected_confidence_level"]["validation"].__setitem__(
            "source", "lstDoesNotExist"
        ),
        "validation referencing an undeclared list",
    )


def test_rejects_unknown_validation_kind() -> None:
    _rejected(
        lambda d: d["inputs"]["duration_years"]["validation"].__setitem__("kind", "magic"),
        "unknown validation kind",
    )


def test_rejects_input_not_placed_in_a_section() -> None:
    _rejected(
        lambda d: d["setup_layout"]["sections"][0]["inputs"].remove("base_year"),
        "a Setup input absent from every section",
    )


def test_rejects_section_referencing_unknown_input() -> None:
    _rejected(
        lambda d: d["setup_layout"]["sections"][0]["inputs"].append("no_such_input"),
        "section referencing an unknown input",
    )


def test_rejects_section_referencing_unknown_table() -> None:
    _rejected(
        lambda d: d["setup_layout"]["sections"][3].__setitem__("table", "no_such_table"),
        "section referencing an unknown table",
    )


def test_rejects_locked_list_without_values() -> None:
    _rejected(
        lambda d: d["config_tables"][4].__setitem__("values", []),
        "a locked constant list with no values",
    )


def test_rejects_locked_list_sized_wrong() -> None:
    _rejected(
        lambda d: d["config_tables"][4].__setitem__("data_rows", 9),
        "a locked list whose data_rows do not match its value count",
    )


def test_rejects_more_seed_values_than_rows() -> None:
    _rejected(
        lambda d: d["config_tables"][1].__setitem__("data_rows", 0),
        "more seeded values than data rows",
    )


def test_rejects_model_controlled_input_without_default() -> None:
    _rejected(
        lambda d: d["inputs"]["reporting_currency"].__setitem__("default", None),
        "a model-controlled input with no locked default",
    )


def test_rejects_seed_row_of_wrong_width() -> None:
    _rejected(
        lambda d: d["tables"]["fx_rates"].__setitem__("seed_rows", [["SAR"]]),
        "a seed row narrower than the table",
    )


# ===========================================================================
# SAR identity invariants
# ===========================================================================
def test_rejects_removing_the_sar_currency_identity() -> None:
    _rejected(
        lambda d: d["config_tables"][1].__setitem__("values", ["USD"]),
        "currency master identity changed away from SAR",
    )


def test_rejects_unlocking_the_sar_currency_identity() -> None:
    _rejected(
        lambda d: d["config_tables"][1].__setitem__("locked_seed_rows", 0),
        "SAR currency identity left user-owned",
    )


def test_rejects_changing_the_fx_identity_rate() -> None:
    _rejected(
        lambda d: d["tables"]["fx_rates"].__setitem__("seed_rows", [["SAR", 2]]),
        "SAR FX identity rate changed away from 1",
    )


def test_rejects_removing_the_fx_identity_row() -> None:
    _rejected(
        lambda d: d["tables"]["fx_rates"].__setitem__("seed_rows", []),
        "SAR FX identity row removed",
    )


def test_rejects_unlocking_the_fx_identity_row() -> None:
    _rejected(
        lambda d: d["tables"]["fx_rates"].__setitem__("locked_seed_rows", 0),
        "SAR FX identity left user-owned",
    )


def test_rejects_identity_not_beginning_with_reporting_currency() -> None:
    _rejected(
        lambda d: d["model_invariants"]["locked_identities"][1].__setitem__("values", ["USD", 1]),
        "declared identity not beginning with the reporting currency",
    )


def test_rejects_identity_targeting_unknown_table() -> None:
    _rejected(
        lambda d: d["model_invariants"]["locked_identities"][0].__setitem__("table", "tblNope"),
        "identity targeting an unknown table",
    )


def test_rejects_editable_reporting_currency_input() -> None:
    _rejected(
        lambda d: d["inputs"]["reporting_currency"].__setitem__("editable", True),
        "reporting currency made user-editable",
    )


def test_rejects_locking_more_rows_than_are_seeded() -> None:
    _rejected(
        lambda d: d["config_tables"][1].__setitem__("locked_seed_rows", 3),
        "locking rows that carry no model-declared value",
    )


def test_rejects_locked_seed_rows_on_a_wholly_locked_table() -> None:
    _rejected(
        lambda d: d["config_tables"][4].__setitem__("locked_seed_rows", 1),
        "locked_seed_rows on an already wholly locked table",
    )


# ===========================================================================
# Excel address bounds. A reference is not valid merely because it parses.
# ===========================================================================
def test_rejects_column_beyond_xfd() -> None:
    _rejected(
        lambda d: d["inputs"]["base_year"].__setitem__("cell", "XFE1"),
        "column XFE is beyond the Excel maximum XFD",
    )


def test_rejects_row_beyond_excel_maximum() -> None:
    _rejected(
        lambda d: d["inputs"]["base_year"].__setitem__("cell", "A1048577"),
        "row 1048577 is beyond the Excel maximum 1048576",
    )


def test_rejects_label_cell_beyond_bounds() -> None:
    _rejected(
        lambda d: d["inputs"]["base_year"].__setitem__("label_cell", "A1048577"),
        "label cell beyond the Excel maximum row",
    )


def test_rejects_two_column_table_starting_at_xfd() -> None:
    _rejected(
        lambda d: d["tables"]["fx_rates"].__setitem__("first_column", "XFD"),
        "a two-column table starting at XFD reaches XFE",
    )


def test_rejects_table_extending_beyond_last_row() -> None:
    _rejected(
        lambda d: (
            d["tables"]["fx_rates"].__setitem__("header_row", 1048570),
            d["tables"]["fx_rates"].__setitem__("data_rows", 20),
        ),
        "header_row + data_rows exceeds the Excel maximum row",
    )


def test_rejects_table_header_row_beyond_bounds() -> None:
    _rejected(
        lambda d: d["config_tables"][0].__setitem__("header_row", 1048577),
        "table header row beyond the Excel maximum",
    )


def test_rejects_section_row_beyond_bounds() -> None:
    _rejected(
        lambda d: d["setup_layout"]["sections"][0].__setitem__("row", 1048577),
        "Setup section row beyond the Excel maximum",
    )


def test_rejects_note_row_beyond_bounds() -> None:
    _rejected(
        lambda d: d["config_tables"][0].__setitem__("note_row", 2000000),
        "Config note row beyond the Excel maximum",
    )


def test_rejects_convention_row_beyond_bounds() -> None:
    _rejected(
        lambda d: d["setup_layout"]["sections"][3].__setitem__("convention_row", 1048577),
        "Setup convention row beyond the Excel maximum",
    )


def test_rejects_intro_row_beyond_bounds() -> None:
    _rejected(
        lambda d: d["setup_layout"]["intro"].__setitem__("row", 1048577),
        "Setup intro row beyond the Excel maximum",
    )


def test_rejects_config_intro_row_beyond_bounds() -> None:
    _rejected(
        lambda d: d["config_layout"]["intro"].__setitem__("row", 9999999),
        "Config intro row beyond the Excel maximum",
    )


def test_bound_validators_accept_the_exact_limits() -> None:
    """XFD and row 1048576 are inside the grid and must NOT be rejected."""
    from pccm_builder.contract_loader import check_cell, check_column, check_row
    assert check_column("XFD", "test") == "XFD"
    assert check_row(1_048_576, "test") == 1_048_576
    assert check_cell("XFD1048576", "test") == "XFD1048576"
    for bad in ("XFE1", "A1048577"):
        try:
            check_cell(bad, "test")
        except ContractError:
            continue
        raise AssertionError(f"{bad} was accepted")


def test_rejects_missing_contract_file() -> None:
    try:
        load_contract(PCCM_ROOT / "spec" / "does_not_exist.yaml")
    except ContractError:
        return
    raise AssertionError("a missing contract was silently accepted")


def test_valid_contract_loads() -> None:
    contract = load_contract(CONTRACT_PATH)
    assert len(contract.inputs) == 9
    assert len(contract.config_tables) == 6
    assert contract.tables["fx_rates"].sheet == "Setup"


def _run_all() -> int:
    tests = sorted(
        (name, fn) for name, fn in globals().items()
        if name.startswith("test_") and callable(fn)
    )
    failures = 0
    print("PCCM Phase 2 input contract validation tests")
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
