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
