#!/usr/bin/env python3
"""PCCM Phase 3: regression tests for the post-build verifier itself.

The gate must not be able to report PASS while a model-controlled cell is covered
by data validation. Comparing range strings could not make that promise: a check
that asserts ``'B12:B36' not in {...}`` is satisfied by 'B12', by 'B12:B20', by
'B20:B36' and by any multi-area sqref that happens to include one of them, every
one of which covers protected identity cells.

Each test therefore builds the real workbook, injects one specific offending
validation, saves it, and runs the ACTUAL post-build verification path --
``verify_workbook`` -- asserting that it fails and names the right check. A test
of the helper alone would prove nothing about the gate.

Runs standalone or under pytest.
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path
from typing import Callable

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from openpyxl.worksheet.datavalidation import DataValidation  # noqa: E402

from pccm_builder import (  # noqa: E402
    build_workbook,
    load_contract,
    load_driver_contract,
    load_structure_contract,
    load_spec,
    verify_workbook,
)
from pccm_builder.verify import data_validation_intersects  # noqa: E402

SPEC_PATH = PCCM_ROOT / "spec" / "workbook.yaml"
CONTRACT_PATH = PCCM_ROOT / "spec" / "input_contract.yaml"
DRIVERS_PATH = PCCM_ROOT / "spec" / "driver_contract.yaml"
STRUCTURE_PATH = PCCM_ROOT / "spec" / "structure_contract.yaml"


# ---------------------------------------------------------------------------
def _specs():
    return (
        load_spec(SPEC_PATH),
        load_contract(CONTRACT_PATH),
        load_driver_contract(DRIVERS_PATH),
        load_structure_contract(STRUCTURE_PATH),
    )


def _inject(worksheet, *refs: str) -> None:
    """Attach one list validation covering *refs* -- several refs give a multi-area sqref."""
    dv = DataValidation(type="list", formula1='"A,B"', allow_blank=True)
    worksheet.add_data_validation(dv)
    for ref in refs:
        dv.add(ref)


def _verify_with(mutate: Callable | None = None):
    """Build, optionally corrupt, save, then run the real post-build verifier."""
    spec, contract, drivers, structure = _specs()
    workbook, _ = build_workbook(spec, contract, drivers, structure)
    if mutate is not None:
        mutate(workbook)
    with tempfile.TemporaryDirectory(prefix="pccm-verifier-") as tmp:
        path = Path(tmp) / "stage_a.xlsx"
        workbook.save(path)
        workbook.close()
        return verify_workbook(path, spec, contract, drivers, structure)


def _failure_mentioning(result, *fragments: str) -> str:
    for failure in result.failures:
        if all(fragment.lower() in failure.lower() for fragment in fragments):
            return failure
    raise AssertionError(
        f"the verifier reported no failure mentioning {fragments}. "
        f"failures: {result.failures or '(none -- it reported PASS)'}"
    )


def _cell_of(table, column_index: int, row_offset: int = 0) -> str:
    return f"{table.column_letter(column_index)}{table.first_data_row + row_offset}"


def _identity_letter(register) -> str:
    return register.column_letter(0)


# ---------------------------------------------------------------------------
# baseline
# ---------------------------------------------------------------------------
def test_01_clean_build_passes_verification() -> None:
    """Without an injected fault the gate must pass, or nothing below means anything."""
    result = _verify_with()
    assert result.ok, f"clean build failed verification: {result.failures}"


# ---------------------------------------------------------------------------
# locked Setup / Config identity rows
# ---------------------------------------------------------------------------
def test_02_validation_on_the_locked_currency_cell_is_caught() -> None:
    _, contract, _, _ = _specs()
    table = contract.table_by_name("tblCurrencies")
    target = _cell_of(table, 0)  # the locked SAR cell

    result = _verify_with(lambda wb: _inject(wb[table.sheet], target))
    failure = _failure_mentioning(result, "tblCurrencies", "locked identity")
    assert target in failure, failure


def test_03_validation_on_the_locked_fx_identity_cell_is_caught() -> None:
    _, contract, _, _ = _specs()
    table = contract.table_by_name("tblFXRates")
    target = _cell_of(table, 1)  # the '1' of the locked SAR|1 identity

    result = _verify_with(lambda wb: _inject(wb[table.sheet], target))
    failure = _failure_mentioning(result, "tblFXRates", "locked identity")
    assert target in failure, failure


def test_04_validation_spanning_locked_and_user_rows_is_caught() -> None:
    """A range that starts in the locked row and runs into user rows still fails."""
    _, contract, _, _ = _specs()
    table = contract.table_by_name("tblFXRates")
    target = f"{_cell_of(table, 0)}:{table.column_letter(0)}{table.last_data_row}"

    result = _verify_with(lambda wb: _inject(wb[table.sheet], target))
    _failure_mentioning(result, "tblFXRates", "locked identity")


# ---------------------------------------------------------------------------
# driver identity columns
# ---------------------------------------------------------------------------
def test_05_validation_on_a_single_cost_line_id_cell_is_caught() -> None:
    """The exact case an equality test on the full range string cannot see."""
    _, _, drivers, _ = _specs()
    register = drivers.registers["cost_lines"]
    target = f"{_identity_letter(register)}{register.first_data_row}"   # B12

    result = _verify_with(lambda wb: _inject(wb[register.sheet], target))
    failure = _failure_mentioning(result, "tblCostLines", "identity column")
    assert target in failure, failure


def test_06_validation_on_a_partial_cost_line_id_range_is_caught() -> None:
    _, _, drivers, _ = _specs()
    register = drivers.registers["cost_lines"]
    letter = _identity_letter(register)
    letter_range = f"{letter}{register.first_data_row}:{letter}{register.first_data_row + 8}"

    result = _verify_with(lambda wb: _inject(wb[register.sheet], letter_range))
    failure = _failure_mentioning(result, "tblCostLines", "identity column")
    assert letter_range in failure, failure


def test_07_validation_on_a_trailing_risk_id_range_is_caught() -> None:
    """Overlap at the tail of the column, not the head."""
    _, _, drivers, _ = _specs()
    register = drivers.registers["risk_register"]
    letter = _identity_letter(register)
    target = f"{letter}{register.last_data_row - 4}:{letter}{register.last_data_row}"

    result = _verify_with(lambda wb: _inject(wb[register.sheet], target))
    failure = _failure_mentioning(result, "tblRiskRegister", "identity column")
    assert target in failure, failure


def test_08_multi_area_sqref_touching_the_id_column_is_caught() -> None:
    """One offending area inside an otherwise innocent multi-area validation."""
    _, _, drivers, _ = _specs()
    register = drivers.registers["risk_register"]
    letter = _identity_letter(register)
    offending = f"{letter}{register.first_data_row + 3}:{letter}{register.first_data_row + 5}"

    result = _verify_with(lambda wb: _inject(wb[register.sheet], "P4", offending, "P6"))
    failure = _failure_mentioning(result, "tblRiskRegister", "identity column")
    assert offending in failure, failure


def test_09_validation_crossing_the_id_column_horizontally_is_caught() -> None:
    """A wide range that merely passes through the ID column still overlaps it."""
    _, _, drivers, _ = _specs()
    register = drivers.registers["cost_lines"]
    row = register.first_data_row + 2
    target = f"{register.column_letter(0)}{row}:{register.column_letter(3)}{row}"

    result = _verify_with(lambda wb: _inject(wb[register.sheet], target))
    _failure_mentioning(result, "tblCostLines", "identity column")


# ---------------------------------------------------------------------------
# no false positives
# ---------------------------------------------------------------------------
def test_10_unrelated_validation_does_not_fail_the_gate() -> None:
    """A rule outside every protected range must leave the whole gate green."""
    _, _, drivers, _ = _specs()
    register = drivers.registers["cost_lines"]
    outside = f"{register.column_letter(len(register.columns) + 2)}{register.last_data_row + 5}"

    result = _verify_with(lambda wb: _inject(wb[register.sheet], outside))
    assert result.ok, (
        f"validation at {outside}, outside every protected range, was reported as a "
        f"breach: {result.failures}"
    )


def test_11_the_real_user_validations_are_not_false_positives() -> None:
    """The build's own validations sit beside protected ranges and must not trip it."""
    result = _verify_with()
    protected = [p for p in result.passed if "locked identity" in p or "identity column" in p]
    assert len(protected) >= 4, f"expected the protection checks to run, saw {protected}"
    assert result.ok


# ---------------------------------------------------------------------------
# helper semantics (in addition to, never instead of, the gate tests above)
# ---------------------------------------------------------------------------
def test_12_helper_handles_cells_ranges_and_multi_area_sqrefs() -> None:
    spec, contract, drivers, structure = _specs()
    workbook, _ = build_workbook(spec, contract, drivers, structure)
    try:
        worksheet = workbook[drivers.registers["cost_lines"].sheet]
        _inject(worksheet, "P4", "R10:T14")

        assert data_validation_intersects(worksheet, "P4")          # single cell
        assert data_validation_intersects(worksheet, "S12")         # inside a range
        assert data_validation_intersects(worksheet, "S1:S1000")    # crossing a range
        assert data_validation_intersects(worksheet, "R10:T14")     # exact match
        assert data_validation_intersects(worksheet, "T14:V20")     # corner touch
        assert not data_validation_intersects(worksheet, "P5")      # adjacent cell
        assert not data_validation_intersects(worksheet, "U10:V14")  # beyond the right edge
        assert not data_validation_intersects(worksheet, "R15:T20")  # below the bottom edge
    finally:
        workbook.close()


def _run_all() -> int:
    tests = sorted(
        (name, fn) for name, fn in globals().items()
        if name.startswith("test_") and callable(fn)
    )
    failures = 0
    print("PCCM Phase 3 post-build verifier intersection regression tests")
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
