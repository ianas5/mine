#!/usr/bin/env python3
"""PCCM Phase 3 negative tests.

Two groups:
  * a malformed driver contract must fail loudly as DriverContractError
  * the preflight cross-specification consistency guards must fail loudly

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
    ContractError,
    DriverContractError,
    build_workbook,
    load_contract,
    load_driver_contract,
    load_spec,
)
from pccm_builder.driver_loader import validate_against_input_contract  # noqa: E402

SPEC_PATH = PCCM_ROOT / "spec" / "workbook.yaml"
CONTRACT_PATH = PCCM_ROOT / "spec" / "input_contract.yaml"
DRIVERS_PATH = PCCM_ROOT / "spec" / "driver_contract.yaml"


def _base() -> dict[str, Any]:
    with DRIVERS_PATH.open("r", encoding="utf-8") as handle:
        return yaml.safe_load(handle)


def _write(data: dict[str, Any], tmp: str, name: str = "broken.yaml") -> Path:
    path = Path(tmp) / name
    with path.open("w", encoding="utf-8") as handle:
        yaml.safe_dump(data, handle, sort_keys=False)
    return path


def _rejected(mutate: Callable[[dict[str, Any]], None], reason: str) -> None:
    data = copy.deepcopy(_base())
    mutate(data)
    with tempfile.TemporaryDirectory(prefix="pccm-baddriver-") as tmp:
        path = _write(data, tmp)
        try:
            load_driver_contract(path)
        except DriverContractError:
            return
        except Exception as error:  # noqa: BLE001
            raise AssertionError(
                f"{reason}: raised {type(error).__name__} instead of DriverContractError"
            ) from error
    raise AssertionError(f"{reason}: invalid driver contract was silently accepted")


def _cost() -> str:
    return "cost_lines"


def _index_of(data: dict[str, Any], register: str, header: str) -> int:
    return next(
        i for i, c in enumerate(data["registers"][register]["columns"]) if c["header"] == header
    )


# ===========================================================================
# schema integrity
# ===========================================================================
def test_rejects_missing_required_column() -> None:
    _rejected(
        lambda d: d["registers"]["cost_lines"]["columns"].pop(
            _index_of(d, "cost_lines", "Quantity")
        ),
        "a locked column removed from the schema",
    )


def test_rejects_reordered_locked_schema() -> None:
    def swap(d: dict[str, Any]) -> None:
        cols = d["registers"]["cost_lines"]["columns"]
        cols[1], cols[2] = cols[2], cols[1]

    _rejected(swap, "locked column order changed")


def test_rejects_renamed_locked_header() -> None:
    _rejected(
        lambda d: d["registers"]["cost_lines"]["columns"][4].__setitem__("header", "Qty"),
        "a locked header renamed",
    )


def test_rejects_duplicate_semantic_key() -> None:
    _rejected(
        lambda d: d["registers"]["cost_lines"]["columns"][2].__setitem__("key", "category"),
        "duplicate column key",
    )


def test_rejects_duplicate_header() -> None:
    def dupe(d: dict[str, Any]) -> None:
        cols = d["registers"]["cost_lines"]["columns"]
        cols[2]["header"] = cols[1]["header"]
        d["locked_schema"]["cost_lines"][2] = d["locked_schema"]["cost_lines"][1]

    _rejected(dupe, "duplicate header")


def test_rejects_duplicate_table_name() -> None:
    _rejected(
        lambda d: d["registers"]["risk_register"].__setitem__("table_name", "tblCostLines"),
        "duplicate driver table name",
    )


def test_rejects_malformed_table_name() -> None:
    _rejected(
        lambda d: d["registers"]["cost_lines"].__setitem__("table_name", "CostLines"),
        "table name not matching tbl<PascalCase>",
    )


# ===========================================================================
# forbidden columns
# ===========================================================================
def test_rejects_included_risk_column() -> None:
    def add_included(d: dict[str, Any]) -> None:
        d["registers"]["risk_register"]["columns"].append(
            {
                "key": "included", "header": "Included", "type": "text",
                "editable": True, "required": True, "number_format": "@", "width": 10,
            }
        )
        d["locked_schema"]["risk_register"].append("Included")

    _rejected(add_included, "an Included column on the risk register")


def test_rejects_cost_total_user_input_column() -> None:
    def add_total(d: dict[str, Any]) -> None:
        d["registers"]["cost_lines"]["columns"].append(
            {
                "key": "total_cost", "header": "Total Cost", "type": "decimal",
                "editable": True, "required": True, "number_format": "#,##0.00", "width": 16,
            }
        )
        d["locked_schema"]["cost_lines"].append("Total Cost")

    _rejected(add_total, "a user-input Total Cost column")


# ===========================================================================
# ownership
# ===========================================================================
def test_rejects_editable_id_column() -> None:
    _rejected(
        lambda d: d["registers"]["cost_lines"]["columns"][0].__setitem__("editable", True),
        "a user-editable identity column",
    )


def test_rejects_model_controlled_user_column_without_validation() -> None:
    """Ownership drift must fail even where there is no validation to notice it by.

    Description declares no validation of its own, so a rule keyed on 'not editable
    AND has validation' would let it silently leave the user's control.
    """
    _rejected(
        lambda d: d["registers"]["cost_lines"]["columns"][
            _index_of(d, "cost_lines", "Description")
        ].__setitem__("editable", False),
        "a user cost-line field demoted to model-controlled",
    )


def test_rejects_model_controlled_quantity_column() -> None:
    _rejected(
        lambda d: d["registers"]["cost_lines"]["columns"][
            _index_of(d, "cost_lines", "Quantity")
        ].__setitem__("editable", False),
        "Quantity demoted to model-controlled",
    )


def test_rejects_model_controlled_risk_owner_column() -> None:
    _rejected(
        lambda d: d["registers"]["risk_register"]["columns"][
            _index_of(d, "risk_register", "Risk Owner")
        ].__setitem__("editable", False),
        "Risk Owner demoted to model-controlled",
    )


def test_rejects_model_controlled_three_point_parameter() -> None:
    _rejected(
        lambda d: d["registers"]["risk_register"]["columns"][
            _index_of(d, "risk_register", "Impact Most Likely")
        ].__setitem__("editable", False),
        "a three-point parameter demoted to model-controlled",
    )


def test_rejects_validation_on_the_id_column() -> None:
    _rejected(
        lambda d: d["registers"]["risk_register"]["columns"][0].__setitem__(
            "validation", {"kind": "list", "source": "lstCategories"}
        ),
        "data validation on a model-controlled identity column",
    )


# ===========================================================================
# validation sources
# ===========================================================================
def test_rejects_unknown_validation_kind() -> None:
    _rejected(
        lambda d: d["registers"]["cost_lines"]["columns"][1]["validation"].__setitem__(
            "kind", "magic"
        ),
        "unknown validation kind",
    )


def test_rejects_validation_source_missing_from_input_contract() -> None:
    """Cross-contract: a driver list source must exist in the input contract."""
    data = copy.deepcopy(_base())
    index = _index_of(data, "cost_lines", "Category")
    data["registers"]["cost_lines"]["columns"][index]["validation"]["source"] = "lstNope"
    with tempfile.TemporaryDirectory(prefix="pccm-baddriver-") as tmp:
        drivers = load_driver_contract(_write(data, tmp))
        try:
            validate_against_input_contract(drivers, load_contract(CONTRACT_PATH))
        except DriverContractError:
            return
    raise AssertionError("an unknown list source was silently accepted")


# ===========================================================================
# bounds, sheets, capacity
# ===========================================================================
def test_rejects_unknown_sheet() -> None:
    data = copy.deepcopy(_base())
    data["registers"]["cost_lines"]["sheet"] = "Nowhere"
    with tempfile.TemporaryDirectory(prefix="pccm-baddriver-") as tmp:
        drivers = load_driver_contract(_write(data, tmp))
        try:
            build_workbook(load_spec(SPEC_PATH), load_contract(CONTRACT_PATH), drivers)
        except RuntimeError:
            return
    raise AssertionError("a register targeting an unknown sheet was accepted")


def test_rejects_register_on_an_input_contract_sheet() -> None:
    data = copy.deepcopy(_base())
    data["registers"]["cost_lines"]["sheet"] = "Setup"
    with tempfile.TemporaryDirectory(prefix="pccm-baddriver-") as tmp:
        drivers = load_driver_contract(_write(data, tmp))
        try:
            validate_against_input_contract(drivers, load_contract(CONTRACT_PATH))
        except DriverContractError:
            return
    raise AssertionError("a register on a Setup/Config sheet was accepted")


def test_rejects_zero_reserved_rows() -> None:
    _rejected(
        lambda d: d["registers"]["cost_lines"].__setitem__("reserved_rows", 0),
        "zero reserved rows",
    )


def test_rejects_negative_reserved_rows() -> None:
    _rejected(
        lambda d: d["registers"]["risk_register"].__setitem__("reserved_rows", -5),
        "negative reserved rows",
    )


def test_rejects_column_beyond_xfd() -> None:
    _rejected(
        lambda d: d["registers"]["cost_lines"].__setitem__("first_column", "XFD"),
        "an 11-column table starting at XFD",
    )


def test_rejects_table_beyond_last_row() -> None:
    _rejected(
        lambda d: (
            d["registers"]["cost_lines"].__setitem__("header_row", 1048570),
            d["registers"]["cost_lines"].__setitem__("reserved_rows", 20),
        ),
        "header_row + reserved_rows beyond the Excel maximum row",
    )


def test_rejects_header_row_beyond_bounds() -> None:
    _rejected(
        lambda d: d["registers"]["cost_lines"].__setitem__("header_row", 1048577),
        "header row beyond the Excel maximum",
    )


def test_rejects_table_overlapping_its_own_section() -> None:
    _rejected(
        lambda d: d["registers"]["cost_lines"].__setitem__("section_row", 12),
        "a section heading inside the table body",
    )


def test_rejects_conditional_format_on_unknown_column() -> None:
    _rejected(
        lambda d: d["registers"]["cost_lines"]["conditional_formatting"][0].__setitem__(
            "target_column", "no_such_column"
        ),
        "conditional formatting referencing an unknown column",
    )


def test_rejects_missing_driver_contract_file() -> None:
    try:
        load_driver_contract(PCCM_ROOT / "spec" / "does_not_exist.yaml")
    except DriverContractError:
        return
    raise AssertionError("a missing driver contract was silently accepted")


# ===========================================================================
# preflight cross-specification guards
# ===========================================================================
def test_rejects_reporting_currency_drift_between_specs() -> None:
    """workbook.yaml and input_contract.yaml must not disagree about SAR."""
    with SPEC_PATH.open(encoding="utf-8") as handle:
        spec_data = yaml.safe_load(handle)
    spec_data["model"]["reporting_currency"] = "USD"
    with tempfile.TemporaryDirectory(prefix="pccm-drift-") as tmp:
        path = Path(tmp) / "workbook.yaml"
        with path.open("w", encoding="utf-8") as handle:
            yaml.safe_dump(spec_data, handle, sort_keys=False)
        spec = load_spec(path)
        try:
            build_workbook(spec, load_contract(CONTRACT_PATH), load_driver_contract(DRIVERS_PATH))
        except RuntimeError as error:
            assert "reporting currency" in str(error).lower()
            return
    raise AssertionError("reporting-currency drift between specifications was accepted")


def test_rejects_reporting_currency_input_pointing_elsewhere() -> None:
    """The invariant is checked against the NAMED semantic input."""
    with CONTRACT_PATH.open(encoding="utf-8") as handle:
        data = yaml.safe_load(handle)
    data["model_invariants"]["reporting_currency_input"] = "project_name"
    with tempfile.TemporaryDirectory(prefix="pccm-drift-") as tmp:
        path = Path(tmp) / "input_contract.yaml"
        with path.open("w", encoding="utf-8") as handle:
            yaml.safe_dump(data, handle, sort_keys=False)
        try:
            load_contract(path)
        except ContractError:
            return
    raise AssertionError("a reporting-currency input pointing at the wrong key was accepted")


def test_rejects_reporting_currency_defined_name_drift() -> None:
    with CONTRACT_PATH.open(encoding="utf-8") as handle:
        data = yaml.safe_load(handle)
    data["model_invariants"]["reporting_currency_defined_name"] = "inpSomethingElse"
    with tempfile.TemporaryDirectory(prefix="pccm-drift-") as tmp:
        path = Path(tmp) / "input_contract.yaml"
        with path.open("w", encoding="utf-8") as handle:
            yaml.safe_dump(data, handle, sort_keys=False)
        try:
            load_contract(path)
        except ContractError:
            return
    raise AssertionError("a drifted reporting-currency defined name was accepted")


def test_valid_driver_contract_loads() -> None:
    drivers = load_driver_contract(DRIVERS_PATH)
    assert len(drivers.all_registers) == 2
    assert drivers.registers["cost_lines"].table_name == "tblCostLines"
    assert drivers.registers["risk_register"].table_name == "tblRiskRegister"
    validate_against_input_contract(drivers, load_contract(CONTRACT_PATH))


def _run_all() -> int:
    tests = sorted(
        (name, fn) for name, fn in globals().items()
        if name.startswith("test_") and callable(fn)
    )
    failures = 0
    print("PCCM Phase 3 driver contract + preflight negative tests")
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
