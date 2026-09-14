#!/usr/bin/env python3
"""P10-R4 RUNTIME YEAR-CELL LOCK STATE - mutation controls.

Each mutation is applied IN MEMORY through the conformance module's memo and
must be refused by the named control; the sources on disk are never touched.
"""
from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import test_phase10_runtime_lock_state as conformance  # noqa: E402


def _tests() -> list[str]:
    names = sorted(name for name in dir(conformance) if name.startswith("test_"))
    assert len(names) >= 10, names
    return names


def _run_battery() -> list[str]:
    refused: list[str] = []
    for name in _tests():
        test = getattr(conformance, name)
        params = getattr(test, "pytestmark", [])
        try:
            marks = [m for m in params if m.name == "parametrize"]
            if marks:
                argnames, values = marks[0].args
                for value in values:
                    test(*value)
            else:
                test()
        except BaseException:  # noqa: BLE001 - any refusal counts
            refused.append(name)
    return refused


def _mutate(expected: str, name: str, before: str, after: str) -> None:
    original = conformance._src(name)
    damaged = original.replace(before, after, 1)
    if damaged == original:
        raise RuntimeError(f"the mutation changed nothing: {before[:70]!r} is no longer in {name}")
    saved = dict(conformance._MEMO)
    conformance._MEMO[name] = damaged
    try:
        refused = _run_battery()
    finally:
        conformance._MEMO.clear()
        conformance._MEMO.update(saved)
    assert refused, "the mutation survived the whole battery"
    assert any(n.startswith(expected) for n in refused), (expected, refused)


def test_00_the_sources_as_written_pass_every_control() -> None:
    assert _run_battery() == []


def test_01_a_fill_only_paint_is_refused() -> None:
    _mutate("test_01", "modWorkbook.bas", "            CellIn(Target, r, c).Locked = False\n", "")


def test_02_unlocking_only_keyed_rows_is_refused() -> None:
    _mutate("test_01", "modWorkbook.bas",
            "            CellIn(Target, r, c).Locked = False\n            If keyed Then\n                CellIn(Target, r, c).Interior.Color = FILL_INPUT\n",
            "            If keyed Then\n                CellIn(Target, r, c).Locked = False\n                CellIn(Target, r, c).Interior.Color = FILL_INPUT\n")


def test_03_locking_unkeyed_runtime_year_cells_is_refused() -> None:
    _mutate("test_01", "modWorkbook.bas",
            "            Else\n                CellIn(Target, r, c).Interior.Color = FILL_LOCKED\n",
            "            Else\n                CellIn(Target, r, c).Locked = True\n                CellIn(Target, r, c).Interior.Color = FILL_LOCKED\n")


def test_04_a_repair_only_lock_correction_is_refused() -> None:
    _mutate("test_07", "modRepair.bas",
            "    modProfiling.SyncRows plan.Kind\n",
            "    modProfiling.SyncRows plan.Kind\n    modProfiling.ProfilingTable(plan.Kind).DataBodyRange.Locked = False\n")


def test_05_a_profiling_only_correction_that_leaves_inflation_exposed_is_refused() -> None:
    _mutate("test_04", "modInflation.bas",
            "    modWorkbook.PaintYearCells target, fixedCols + 1, YearCount, 1\n", "")


def test_06_changing_the_projection_instead_of_the_runtime_is_refused() -> None:
    _mutate("test_08", "protection.py",
            "for row in range(grid.first_data_row, grid.last_data_row + 1)]",
            "for row in range(grid.first_data_row, grid.first_data_row + 1)]")


def test_07_unlocking_the_fixed_columns_is_refused() -> None:
    _mutate("test_03", "modWorkbook.bas",
            "        For c = FirstYearColumn To FirstYearColumn + YearCount - 1\n            CellIn(Target, r, c).Locked = False\n",
            "        For c = 1 To FirstYearColumn + YearCount - 1\n            CellIn(Target, r, c).Locked = False\n")


def test_08_the_runner_assigning_locked_itself_is_refused() -> None:
    _mutate("test_11", "runner",
            "        $state = Get-FaCellLockState -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex $probe.Row -ColumnIndex $probe.Column\n",
            "        $wb.Worksheets.Item($gridSheet).Range('D13').Locked = $false\n        $state = Get-FaCellLockState -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex $probe.Row -ColumnIndex $probe.Column\n")


def test_09_weakening_worksheet_protection_is_refused() -> None:
    _mutate("test_08", "modProtection.bas",
            "        sheet.Protect UserInterfaceOnly:=True, _\n",
            "        sheet.Protect UserInterfaceOnly:=True, AllowFormattingCells:=True, _\n")


def test_10_a_row_materialised_after_the_paint_is_refused() -> None:
    """THE UNCOVERED PATH: a row sync that adds rows and no longer repaints."""
    _mutate("test_04", "modProfiling.bas",
            "    modWorkbook.PaintYearCells target, fixedCols + 1, yearCols, 1\n", "")


def test_11_a_restore_that_drops_the_lock_state_is_refused() -> None:
    _mutate("test_06", "modWorkbook.bas",
            "            Target.DataBodyRange.Cells(r, c).Locked = CBool(Snapshot.Locks(r, c))\n", "")


def test_12_a_second_production_module_changed_alongside_is_refused() -> None:
    _mutate("test_09", "modWorkbook.bas",
            "Public Function IsErrorText(ByVal Text As String) As Boolean\n",
            "Public Function IsErrorText(ByVal Text As String) As Boolean\n    ' drifted\n")


if __name__ == "__main__":
    failures = 0
    for name in sorted(n for n in dir() if n.startswith("test_")):
        try:
            globals()[name]()
            print(f"PASS {name}")
        except BaseException as exc:  # noqa: BLE001
            failures += 1
            print(f"FAIL {name}: {exc}")
    sys.exit(1 if failures else 0)
