#!/usr/bin/env python3
"""P10-R4 RUNTIME YEAR-CELL LOCK STATE - the source controls.

WHAT WINDOWS PROVED. Final acceptance run 10 at 5e0df9b, with protection ON
and no window open, read Locked=True on a keyed year cell from the first Apply
Timeline (D13), on a regrown keyed year cell (G13) and on a regrown unkeyed
reserved year cell (G16), while the protection projection declared all three
unlocked and the permanent-id cell (B13) read Locked=True, declared locked.
Stage A marks the year-column worksheet addresses unlocked over every reserved
body row, but the year columns do not exist at Stage A: the runtime
materialises them with ListColumns.Add, and Excel materialises those cells
locked.

THE ACCEPTED LOCK CONTRACT is the projection's: every project-year body cell of
the three grids unlocked, keyed or not; fixed columns and headers locked. The
fill stays a separate concern and still follows the key.

THE OWNER. modWorkbook.PaintYearCells is the shared post-materialisation
treatment every profiling and inflation year column already receives, so it
sets Locked = False on every year body cell it traverses. SnapshotTable /
RestoreTable carry the lock state with the fill, so a rollback that rebuilds a
column or a row puts back what the user could type in before.

Runs standalone or under pytest.
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path

import pytest

PCCM_ROOT = Path(__file__).resolve().parent.parent
REPO_ROOT = PCCM_ROOT.parent
SRC = PCCM_ROOT / "src" / "vba"
RUNNER = PCCM_ROOT / "bootstrap" / "windows" / "phase10_final_acceptance.ps1"
PROJECTION = PCCM_ROOT / "build" / "phase10_protection_inspection.json"
CANDIDATE = "5e0df9b"

_MEMO: dict[str, str] = {}


def _src(name: str) -> str:
    """A production, runner or record text through a memo, so the validation
    battery can mutate it in memory."""
    if name not in _MEMO:
        path = RUNNER if name == "runner" else (
            PCCM_ROOT / "builder" / "pccm_builder" / "protection.py" if name == "protection.py" else
            PCCM_ROOT / "spec" / "structure_contract.yaml" if name == "structure_contract.yaml" else
            PCCM_ROOT / "docs" / "phase7_closure.md" if name == "phase7_closure.md" else
            SRC / name)
        _MEMO[name] = path.read_text(encoding="utf-8")
    return _MEMO[name]


def _code(name: str) -> str:
    return "\n".join(line for line in _src(name).splitlines() if not line.strip().startswith("'"))


def _procedure(name: str, header: str) -> str:
    code = _code(name)
    start = code.index(header)
    end = code.index("End Sub" if header.lstrip().startswith(("Public Sub", "Private Sub")) else "End Function", start)
    return code[start:end]


def _git(*args: str) -> str:
    return subprocess.run(["git", *args], cwd=REPO_ROOT, capture_output=True, text=True, check=True).stdout


def _statements(text: str) -> list[str]:
    return [line.strip() for line in text.splitlines() if line.strip()]


# ===========================================================================
# A. THE OWNER
# ===========================================================================
def test_01_every_painted_year_cell_is_unlocked_before_the_fill_and_not_by_key() -> None:
    paint = _procedure("modWorkbook.bas", "Public Sub PaintYearCells")
    statements = _statements(paint)
    loop = statements.index("For c = FirstYearColumn To FirstYearColumn + YearCount - 1")
    assert statements[loop + 1] == "CellIn(Target, r, c).Locked = False"
    assert statements[loop + 2] == "If keyed Then"
    assert paint.count(".Locked = False") == 1 and ".Locked = True" not in paint
    # not inside the keyed branch, not conditional on anything
    unlocked_at = paint.index("CellIn(Target, r, c).Locked = False")
    assert paint.rfind("If keyed Then", 0, unlocked_at) == -1
    assert "keyed" not in paint[paint.rfind("\n", 0, unlocked_at): unlocked_at]


def test_02_the_fill_treatment_still_follows_the_key_exactly_as_accepted() -> None:
    paint = _procedure("modWorkbook.bas", "Public Sub PaintYearCells")
    statements = _statements(paint)
    branch = statements.index("If keyed Then")
    assert statements[branch: branch + 5] == ["If keyed Then",
                                              "CellIn(Target, r, c).Interior.Color = FILL_INPUT",
                                              "Else",
                                              "CellIn(Target, r, c).Interior.Color = FILL_LOCKED",
                                              "End If"]
    then = _git("show", f"{CANDIDATE}:pccm/src/vba/modWorkbook.bas")
    then_paint = then[then.index("Public Sub PaintYearCells"): then.index("End Sub", then.index("Public Sub PaintYearCells"))]
    then_statements = [line.strip() for line in then_paint.splitlines() if line.strip() and not line.strip().startswith("'")]
    assert [s for s in statements if s != "CellIn(Target, r, c).Locked = False"] == then_statements


def test_03_only_the_year_body_cells_are_touched_never_fixed_columns_or_headers() -> None:
    paint = _procedure("modWorkbook.bas", "Public Sub PaintYearCells")
    assert "For c = FirstYearColumn To FirstYearColumn + YearCount - 1" in paint
    assert "For r = 1 To rowCount" in paint and "rowCount = BodyRowCount(Target)" in paint
    assert "HeaderRowRange" not in paint and "ListColumns" not in paint
    assert "KeyColumn" in paint and "CellIn(Target, r, KeyColumn)).Locked" not in paint
    writes = re.findall(r"CellIn\(Target, r, (\w+)\)\.(?:Locked|Interior)", paint)
    assert set(writes) == {"c"}, writes


# ===========================================================================
# B. EVERY RUNTIME PATH
# ===========================================================================
PAINT_SITES = (
    ("modProfiling.bas", "Public Sub SetYearColumns", "modWorkbook.PaintYearCells target, fixedCols + 1, NewCount, 1"),
    ("modProfiling.bas", "Public Sub SyncRows", "modWorkbook.PaintYearCells target, fixedCols + 1, yearCols, 1"),
    ("modInflation.bas", "Public Sub SetYearColumns", "modWorkbook.PaintYearCells target, fixedCols + 1, YearCount, 1"),
    ("modInflation.bas", "Public Sub SyncProfileRows", "modWorkbook.PaintYearCells target, fixedCols + 1, yearCols, 1"),
)


@pytest.mark.parametrize("module,header,call", PAINT_SITES)
def test_04_every_materialising_owner_ends_with_the_paint_after_its_adds(module: str, header: str, call: str) -> None:
    """Columns are materialised by SetYearColumns, rows by the row sync; each
    procedure's LAST statement is the shared paint over every body row."""
    body = _procedure(module, header)
    statements = _statements(body)
    assert statements[-1] == call, statements[-1]
    for add in ("ListColumns.Add", "ListRows.Add"):
        for match in re.finditer(re.escape(add), body):
            assert match.start() < body.index(call), (module, header, add)


def test_05_no_grid_row_or_column_is_materialised_outside_a_painting_owner_or_the_restore() -> None:
    covered = {(module, header) for module, header, _call in PAINT_SITES}
    for module in ("modProfiling.bas", "modInflation.bas"):
        code = _code(module)
        for match in re.finditer(r"ListColumns\.Add|ListRows\.Add", code):
            start = code.rfind("\nPublic Sub ", 0, match.start())
            start2 = code.rfind("\nPrivate Sub ", 0, match.start())
            start = max(start, start2)
            header = code[start + 1: code.index("(", start)]
            assert (module, header) in covered, (module, header)
    # the only other materialising site is the snapshot restore, which carries the lock state
    workbook = _code("modWorkbook.bas")
    sites = [m.start() for m in re.finditer(r"ListColumns\.Add|ListRows\.Add", workbook)]
    restore = workbook.index("Public Sub RestoreTable")
    restore_end = workbook.index("End Sub", restore)
    assert sites and all(restore < site < restore_end for site in sites), sites
    # Repair and Timeline reach the grids only through those owners
    for module in ("modRepair.bas", "modTimeline.bas", "modDrivers.bas"):
        code = _code(module)
        assert "ListColumns.Add" not in code
        assert ".Locked" not in code, module
    assert "modProfiling.SetYearColumns" in _code("modRepair.bas") and "modProfiling.SyncRows" in _code("modRepair.bas")
    assert "modProfiling.SetYearColumns" in _code("modTimeline.bas") and "modInflation.SetYearColumns" in _code("modTimeline.bas")


def test_06_the_snapshot_carries_the_lock_state_and_the_restore_puts_it_back() -> None:
    code = _code("modWorkbook.bas")
    kind = code[code.index("Public Type TableSnapshot"): code.index("End Type")]
    assert "Locks()       As Variant" in kind and "Fills()       As Variant" in kind
    snapshot = _procedure("modWorkbook.bas", "Public Function SnapshotTable")
    assert "ReDim s.Locks(1 To IIf(s.RowCount < 1, 1, s.RowCount), 1 To s.ColumnCount)" in snapshot
    assert "s.Locks(r, c) = Target.DataBodyRange.Cells(r, c).Locked" in snapshot
    assert snapshot.index("s.Fills(r, c) = ") < snapshot.index("s.Locks(r, c) = ")
    restore = _procedure("modWorkbook.bas", "Public Sub RestoreTable")
    statements = _statements(restore)
    fill = statements.index("Target.DataBodyRange.Cells(r, c).Interior.Color = Snapshot.Fills(r, c)")
    assert statements[fill + 1] == "Target.DataBodyRange.Cells(r, c).Locked = CBool(Snapshot.Locks(r, c))"
    assert restore.index("ListRows.Add") < restore.index("Snapshot.Locks(r, c)")
    assert restore.index("ListColumns.Add") < restore.index("Snapshot.Locks(r, c)")


# ===========================================================================
# C. ONE OWNER, NOTHING ELSE MOVED
# ===========================================================================
def test_07_the_lock_state_is_written_by_the_shared_owner_alone() -> None:
    writers = {}
    for path in sorted(SRC.glob("*.bas")) + [SRC / "ThisWorkbook.vba"]:
        code = "\n".join(line for line in _src(path.name).splitlines() if not line.strip().startswith("'"))
        count = len(re.findall(r"\.Locked\s*=", code))
        if count:
            writers[path.name] = count
    assert writers == {"modWorkbook.bas": 2}, writers
    workbook = _code("modWorkbook.bas")
    for header in ("Public Sub PaintYearCells", "Public Sub RestoreTable"):
        assert ".Locked =" in _procedure("modWorkbook.bas", header), header


def test_08_the_projection_and_the_protection_policy_are_untouched() -> None:
    """RUNTIME MOVES TO MATCH THE PROJECTION, never the other way."""
    for path in ("pccm/builder/pccm_builder/protection.py", "pccm/spec/structure_contract.yaml",
                 "pccm/src/vba/modProtection.bas"):
        name = path.rsplit("/", 1)[1]
        assert _src(name) == _git("show", f"{CANDIDATE}:{path}"), path
    builder = _src("protection.py")
    assert "for row in range(grid.first_data_row, grid.last_data_row + 1)]" in builder
    assert "return int(structure.limits.max_generated_year_columns)" in builder
    protection = _code("modProtection.bas")
    assert "sheet.Protect UserInterfaceOnly:=True, _" in protection


def test_09_the_correction_reverses_exactly_to_the_candidate_and_nothing_else_moved() -> None:
    sys.path.insert(0, str(PCCM_ROOT / "tests"))
    from vba_runtime_lock_state import (ACCEPTED_BEFORE_RUNTIME_LOCK_STATE, DECLARED_RUNTIME_LOCK_STATE_CHANGES,
                                        strip_runtime_lock_state)
    assert ACCEPTED_BEFORE_RUNTIME_LOCK_STATE == CANDIDATE
    assert set(DECLARED_RUNTIME_LOCK_STATE_CHANGES) == {"modWorkbook.bas"}
    current = _src("modWorkbook.bas")
    then = _git("show", f"{CANDIDATE}:pccm/src/vba/modWorkbook.bas")
    assert strip_runtime_lock_state("modWorkbook.bas", current) == then
    assert current != then, "the correction is absent"
    changed = _git("diff", "--name-only", CANDIDATE, "--", "pccm/src", "pccm/spec", "pccm/builder").split()
    assert changed == ["pccm/src/vba/modWorkbook.bas"], changed


# ===========================================================================
# D. THE PROJECTION AND THE RUNNER AGREE WITH THE OWNER
# ===========================================================================
@pytest.mark.skipif(not PROJECTION.is_file(), reason="Stage A has not been built into pccm/build")
def test_10_the_projection_declares_every_reserved_year_cell_unlocked_and_the_fixed_columns_locked() -> None:
    import yaml
    structure = yaml.safe_load(_src("structure_contract.yaml"))
    projection = json.loads(PROJECTION.read_text(encoding="utf-8"))
    span = int(structure["limits"]["max_generated_year_columns"])
    for key, sheet in (("cost_profiling", "Cost Profiling"), ("risk_profiling", "Risk Profiling"), ("inflation", "Inflation")):
        grid = structure["grids"][key]
        entry = next(e for e in projection["sheets"] if e["sheet"] == sheet)
        rows = int(grid["reserved_rows"])
        assert len(entry["unlocked"]) == rows * span, (sheet, len(entry["unlocked"]))
        first_row = min(int("".join(ch for ch in a if ch.isdigit())) for a in entry["unlocked"])
        last_row = max(int("".join(ch for ch in a if ch.isdigit())) for a in entry["unlocked"])
        assert last_row - first_row + 1 == rows, sheet
        fixed = len(grid["fixed_columns"])
        first_letter = min("".join(ch for ch in a if ch.isalpha()) for a in entry["unlocked"] if a.endswith(str(first_row)))
        assert first_letter != "B" and fixed >= 1


def test_11_the_windows_check_expects_the_projection_rule_and_never_assigns_locked() -> None:
    runner = "\n".join(line for line in _src("runner").splitlines() if not line.strip().startswith("#"))
    assert "Add-FaCheck 'repair.width-growth.lock-state' ($lockProblems.Count -eq 0)" in runner
    for label in ("regrown keyed weight", "existing keyed weight", "regrown unkeyed reserved weight"):
        assert re.search(r"Label = '" + re.escape(label) + r"';[^\n]*ExpectLocked = \$false", runner), label
    assert re.search(r"Label = 'permanent id';[^\n]*ExpectLocked = \$true", runner)
    assert "$declared = ($declaredUnlocked -ccontains $state.Address)" in runner
    assert not re.search(r"\.Locked\s*=[^=]", runner)


# ===========================================================================
# E. THE HISTORICAL EVIDENCE CONTROLS DECLARE THE LAYER, BY NAME AND BY REVERSAL
# ===========================================================================
LAYER_COMMIT = "548799f"
LAYER_SUBJECT = "Phase 10: runtime year cells are unlocked by the shared paint owner"


def _record_text() -> str:
    return _src("phase7_closure.md")


def test_12_the_layer_commit_is_the_one_the_subject_resolves_to_and_touches_only_the_owner() -> None:
    found = _git("log", "--format=%h", f"--grep={LAYER_SUBJECT}", "--fixed-strings", f"{CANDIDATE}..HEAD").split()
    assert found == [LAYER_COMMIT], found
    assert _git("show", "--name-only", "--format=", LAYER_COMMIT).split() and \
        [p for p in _git("show", "--name-only", "--format=", LAYER_COMMIT).split() if p.startswith("pccm/src/")] == \
        ["pccm/src/vba/modWorkbook.bas"]
    assert _git("diff", "--name-only", CANDIDATE, LAYER_COMMIT, "--", "pccm/src", "pccm/spec", "pccm/builder").split() == \
        ["pccm/src/vba/modWorkbook.bas"]


def test_13_the_phase_7_closure_record_declares_the_layer_and_keeps_both_authorities() -> None:
    text = _record_text()
    block = text.split("### 1.1")[1].split("\n---")[0]
    rows = [line for line in block.splitlines() if line.startswith("| `pccm/src/vba/modWorkbook.bas`")]
    assert len(rows) == 1, rows
    row = rows[0]
    for fact in (LAYER_SUBJECT, "P10-R4", "final Windows acceptance run 10", "`548799f`", "`ListColumns.Add`",
                 "`PaintYearCells`", "`SnapshotTable` / `RestoreTable`", "keyed or unkeyed",
                 "No Phase-7 algorithm, numerical contract, simulation contract or acceptance authority changed",
                 "`tests/vba_runtime_lock_state.py`", "Phase 7 remains CLOSED", "`79d4c3e`", "`ad78988`"):
        assert fact in row, fact
    assert f"--grep='{LAYER_SUBJECT}'" in block
    assert "`79d4c3e` remains the implementation baseline" in text
    assert "`ad78988` remains the evidence head" in text
    assert "does **not** reopen Phase 7" in text
    # THE HISTORY IS NOT REWRITTEN: outside the §1.1 table the record is byte-identical to the layer commit's.
    then = _git("show", f"{LAYER_COMMIT}:pccm/docs/phase7_closure.md")

    def outside(record: str) -> str:
        head, rest = record.split("### 1.1", 1)
        return head + rest.split("\n---", 1)[1]

    assert outside(text) == outside(then)


def test_14_the_phase_8_declared_corrections_admit_only_the_named_layer_with_its_removal_and_reversal() -> None:
    sys.path.insert(0, str(PCCM_ROOT / "tests"))
    import test_phase8_charts as charts
    declared = charts.DECLARED_PRODUCTION_CORRECTIONS
    assert "pccm/src/vba/modWorkbook.bas" in declared
    reason, removals = declared["pccm/src/vba/modWorkbook.bas"]
    assert LAYER_COMMIT in reason and "P10-R4" in reason and "Locked = False" in reason
    assert removals == ("' Keyed-ness decides the treatment:",)
    for key in declared:
        assert key.startswith("pccm/src/vba/") and key.endswith(".bas") and not any(ch in key for ch in "*?[")
        assert (PCCM_ROOT.parent / key).is_file(), key
    charts._declared_production_changes(charts._git, charts.P81_ACCEPTANCE)
    charts._declared_production_changes(charts._git, charts.P82_ACCEPTANCE)


def test_15_the_phase_9_control_declares_the_layer_from_the_reversal_authority_alone() -> None:
    sys.path.insert(0, str(PCCM_ROOT / "tests"))
    import test_phase9_model_check as model_check
    source = (PCCM_ROOT / "tests" / "test_phase9_model_check.py").read_text(encoding="utf-8")
    body = source[source.index("def test_46_every_production_change_is_declared_and_is_only_plumbing"):]
    body = body[: body.index("\ndef ")]
    assert 'phase10_r4 = {f"pccm/src/vba/{name}" for name in DECLARED_RUNTIME_LOCK_STATE_CHANGES}' in body
    assert 'assert phase10_r4 == {"pccm/src/vba/modWorkbook.bas"}, phase10_r4' in body
    assert "assert strip_runtime_lock_state(name, current) == accepted" in body
    assert "assert accepted == at_head" in body and 'assert current != accepted' in body
    model_check.test_46_every_production_change_is_declared_and_is_only_plumbing()


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
