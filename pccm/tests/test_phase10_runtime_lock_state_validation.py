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


# ---------------------------------------------------------------------------
# THE DECLARATIONS: each way of faking one is refused
# ---------------------------------------------------------------------------
def _with_record(damaged: str):
    """The Phase-7 record swapped in memory for the record module AND for the
    P10-R4 suite, so both controls read the same damaged text."""
    import test_phase7_closure_record as record
    return record, damaged


def test_13_a_missing_phase_7_declaration_is_refused_by_the_record_control() -> None:
    import test_phase7_closure_record as record
    original = record._text()
    damaged = "\n".join(line for line in original.splitlines()
                        if not line.startswith("| `pccm/src/vba/modWorkbook.bas`")) + "\n"
    saved = record._text
    record._text = lambda: damaged
    try:
        try:
            record.test_10_the_implementation_authority_is_the_last_commit_touching_src_or_spec()
        except AssertionError as exc:
            assert "modWorkbook.bas" in str(exc)
        else:
            raise AssertionError("the record control accepted an undeclared modWorkbook change")
    finally:
        record._text = saved
    _mutate("test_13", "phase7_closure.md", "| `pccm/src/vba/modWorkbook.bas` |", "| `pccm/src/vba/modWorkbook.bas.old` |")


def test_14_a_rewritten_phase_7_authority_is_refused() -> None:
    _mutate("test_13", "phase7_closure.md", "`79d4c3e` remains the implementation baseline", "`548799f` remains the implementation baseline")


def test_15_rewritten_historical_windows_evidence_is_refused() -> None:
    """A HISTORICAL COUNT OUTSIDE THE §1.1 TABLE MAY NOT MOVE."""
    original = conformance._record_text()
    line = next(l for l in original.splitlines() if l.startswith("| Tests collected |"))
    assert "### 1.1" in original and original.index(line) > original.index("### 1.1")
    _mutate("test_13", "phase7_closure.md", line, line.replace("4,663", "4,000"))


def test_16_a_wrong_layer_commit_in_the_phase_8_declaration_is_refused() -> None:
    import test_phase8_charts as charts
    reason, removals = charts.DECLARED_PRODUCTION_CORRECTIONS["pccm/src/vba/modWorkbook.bas"]
    saved = dict(charts.DECLARED_PRODUCTION_CORRECTIONS)
    charts.DECLARED_PRODUCTION_CORRECTIONS["pccm/src/vba/modWorkbook.bas"] = (reason.replace("548799f", "5e0df9b"), removals)
    try:
        try:
            conformance.test_14_the_phase_8_declared_corrections_admit_only_the_named_layer_with_its_removal_and_reversal()
        except AssertionError:
            pass
        else:
            raise AssertionError("a wrong layer commit survived")
    finally:
        charts.DECLARED_PRODUCTION_CORRECTIONS.clear()
        charts.DECLARED_PRODUCTION_CORRECTIONS.update(saved)


def test_17_a_missing_phase_8_declaration_is_refused_by_the_shared_control() -> None:
    import pytest as _pytest
    import test_phase8_charts as charts
    saved = dict(charts.DECLARED_PRODUCTION_CORRECTIONS)
    del charts.DECLARED_PRODUCTION_CORRECTIONS["pccm/src/vba/modWorkbook.bas"]
    try:
        with _pytest.raises(AssertionError, match="without being declared"):
            charts._declared_production_changes(charts._git, charts.P81_ACCEPTANCE)
    finally:
        charts.DECLARED_PRODUCTION_CORRECTIONS.clear()
        charts.DECLARED_PRODUCTION_CORRECTIONS.update(saved)


def test_18_a_wildcard_or_a_second_file_admitted_under_the_layer_is_refused() -> None:
    import pytest as _pytest
    import test_phase8_charts as charts
    saved = dict(charts.DECLARED_PRODUCTION_CORRECTIONS)
    for key in ("pccm/src/vba/*", "pccm/src/vba/modProfiling.bas"):
        charts.DECLARED_PRODUCTION_CORRECTIONS[key] = charts.DECLARED_PRODUCTION_CORRECTIONS["pccm/src/vba/modWorkbook.bas"]
        try:
            with _pytest.raises(AssertionError):
                conformance.test_14_the_phase_8_declared_corrections_admit_only_the_named_layer_with_its_removal_and_reversal()
            with _pytest.raises(AssertionError):
                charts.test_94_the_declared_production_rule_passes_on_the_real_repository()
        finally:
            charts.DECLARED_PRODUCTION_CORRECTIONS.clear()
            charts.DECLARED_PRODUCTION_CORRECTIONS.update(saved)


def test_19_a_declaration_without_its_mechanical_reversal_is_refused_everywhere() -> None:
    import pytest as _pytest
    import vba_runtime_lock_state as layer
    import test_phase8_charts as charts
    import test_phase9_model_check as model_check
    saved = dict(layer._HUNKS)
    layer._HUNKS.clear()
    try:
        with _pytest.raises(AssertionError):
            conformance.test_09_the_correction_reverses_exactly_to_the_candidate_and_nothing_else_moved()
        with _pytest.raises(AssertionError):
            charts.test_94_the_declared_production_rule_passes_on_the_real_repository()
        with _pytest.raises(AssertionError):
            model_check.test_46_every_production_change_is_declared_and_is_only_plumbing()
    finally:
        layer._HUNKS.clear()
        layer._HUNKS.update(saved)


def test_20_a_second_production_file_admitted_under_the_phase_9_declaration_is_refused() -> None:
    import pytest as _pytest
    import vba_runtime_lock_state as layer
    import test_phase9_model_check as model_check
    saved = dict(layer.DECLARED_RUNTIME_LOCK_STATE_CHANGES)
    layer.DECLARED_RUNTIME_LOCK_STATE_CHANGES["modProfiling.bas"] = "smuggled"
    try:
        with _pytest.raises(AssertionError):
            model_check.test_46_every_production_change_is_declared_and_is_only_plumbing()
    finally:
        layer.DECLARED_RUNTIME_LOCK_STATE_CHANGES.clear()
        layer.DECLARED_RUNTIME_LOCK_STATE_CHANGES.update(saved)


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
