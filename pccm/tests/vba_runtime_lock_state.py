#!/usr/bin/env python3
"""P10-R4 CORRECTION: take the runtime year-cell lock state back out of modWorkbook.

WHY A REVERSAL. Final acceptance run 10 at 5e0df9b proved on Windows, with
protection on, that the year cells the runtime materialises through
ListColumns.Add read Locked=True - a keyed year cell from the first Apply
Timeline, a regrown keyed year cell and a regrown unkeyed reserved year cell
alike - while the protection projection declares every one of them unlocked.
Stage A marks those worksheet addresses unlocked, but the year columns do not
exist at Stage A. The correction lives in modWorkbook.bas only: PaintYearCells,
the shared post-materialisation treatment every profiling and inflation year
column already receives, sets Locked = False on every year body cell it
traverses; and SnapshotTable / RestoreTable carry the lock state with the fill,
so a rollback that rebuilds a column or a row puts back what the user could
type in before. Every control that pins production to an accepted tree keeps
its claim the way the earlier layers established: taking the declared change
back out must reproduce the accepted bytes exactly.

THE FRAGMENTS ARE GENERATED FROM THE REAL DIFF against 5e0df9b, the candidate
that executed run 10, and applying them all must reproduce that file byte for
byte. modWorkbook.bas is LF; the fragments carry LF.
"""
from __future__ import annotations

# The tree the reversal must reproduce: the candidate final acceptance run 10 executed.
ACCEPTED_BEFORE_RUNTIME_LOCK_STATE = "5e0df9b"

DECLARED_RUNTIME_LOCK_STATE_CHANGES = {
    "modWorkbook.bas":
        "PaintYearCells sets Locked = False on every year body cell it traverses, "
        "keyed or not, before the unchanged fill treatment, and SnapshotTable / "
        "RestoreTable capture and restore the Locked state of every body cell "
        "beside its fill (P10-R4, final acceptance run 10)",
}

# (current fragment, accepted fragment), in file order.
_HUNKS: dict[str, tuple[tuple[str, str], ...]] = {
    "modWorkbook.bas": (('    Fills()       As Variant\n    Locks()       As Variant\n    NumberFormats() As Variant\n',
  '    Fills()       As Variant\n    NumberFormats() As Variant\n'),
 ("'\n' Keyed-ness decides the FILL:\n'\n", "'\n' Keyed-ness decides the treatment:\n'\n"),
 ("'                         and anything typed there becomes orphan data.\n"
  "'\n"
  "' THE LOCK STATE IS NOT THE FILL, AND IT DOES NOT FOLLOW THE KEY. The accepted\n"
  "' protection policy - the builder's grid_year_columns rule and the projection it\n"
  "' emits - unlocks the year columns over EVERY reserved body row, keyed or not,\n"
  "' and keeps the fixed columns locked; a row keyed later must already be a cell\n"
  "' the user can type in. Stage A marks those worksheet cells unlocked, but the\n"
  "' year columns themselves do not exist at Stage A: they are materialised here at\n"
  "' runtime by ListColumns.Add, and final acceptance run 10 proved on Windows that\n"
  "' the cells Excel materialises that way read Locked=True - a keyed year cell from\n"
  "' the first Apply Timeline, a regrown keyed year cell and a regrown unkeyed\n"
  "' reserved year cell alike, every one of them declared unlocked. So the one\n"
  "' treatment every runtime year cell receives is applied here, explicitly, by the\n"
  "' owner that already paints it: Locked = False on every year body cell this\n"
  "' procedure traverses, and nothing outside that range. Headers and fixed columns\n"
  "' are never touched by this procedure and keep what Stage A gave them.\n"
  'Public Sub PaintYearCells(ByVal Target As ListObject, ByVal FirstYearColumn As Long, _\n',
  "'                         and anything typed there becomes orphan data.\n"
  'Public Sub PaintYearCells(ByVal Target As ListObject, ByVal FirstYearColumn As Long, _\n'),
 ('        For c = FirstYearColumn To FirstYearColumn + YearCount - 1\n'
  '            CellIn(Target, r, c).Locked = False\n'
  '            If keyed Then\n',
  '        For c = FirstYearColumn To FirstYearColumn + YearCount - 1\n'
  '            If keyed Then\n'),
 ('    ReDim s.Fills(1 To IIf(s.RowCount < 1, 1, s.RowCount), 1 To s.ColumnCount)\n'
  '    ReDim s.Locks(1 To IIf(s.RowCount < 1, 1, s.RowCount), 1 To s.ColumnCount)\n'
  '\n',
  '    ReDim s.Fills(1 To IIf(s.RowCount < 1, 1, s.RowCount), 1 To s.ColumnCount)\n\n'),
 ('            s.Fills(r, c) = Target.DataBodyRange.Cells(r, c).Interior.Color\n'
  "            ' And the lock state with the fill: a rollback that rebuilds a column or\n"
  "            ' a row through ListColumns.Add / ListRows.Add gets cells Excel\n"
  "            ' materialises locked, so what the user could type in before the\n"
  "            ' operation is captured here and put back by RestoreTable.\n"
  '            s.Locks(r, c) = Target.DataBodyRange.Cells(r, c).Locked\n'
  '        Next c\n',
  '            s.Fills(r, c) = Target.DataBodyRange.Cells(r, c).Interior.Color\n        Next c\n'),
 ('            Target.DataBodyRange.Cells(r, c).Interior.Color = Snapshot.Fills(r, c)\n'
  '            Target.DataBodyRange.Cells(r, c).Locked = CBool(Snapshot.Locks(r, c))\n'
  '        Next c\n',
  '            Target.DataBodyRange.Cells(r, c).Interior.Color = Snapshot.Fills(r, c)\n'
  '        Next c\n')),
}


def strip_runtime_lock_state(module_name: str, text: str) -> str:
    """`text` with the P10-R4 correction removed - all or none, exactly as the
    earlier layers: none of the fragments present returns the text unchanged,
    all of them present takes the layer off exactly, some of them raises."""
    hunks = _HUNKS.get(module_name, ())
    if not hunks:
        return text
    matched = []
    for current, accepted in hunks:
        for pair in ((current, accepted),
                     (current.replace("\r\n", "\n"), accepted.replace("\r\n", "\n"))):
            if text.count(pair[0]) == 1:
                matched.append(pair)
                break
        else:
            matched.append(None)
    if all(pair is None for pair in matched):
        return text
    if any(pair is None for pair in matched):
        missing = [hunks[i][0].splitlines()[0] for i, pair in enumerate(matched) if pair is None]
        raise AssertionError(
            f"{module_name}: the declared runtime lock-state correction is partially present - "
            f"{len(hunks) - len(missing)} of {len(hunks)} fragments match, so the reversal "
            f"cannot be exact. A fragment moved or something rode along inside it:\n  {missing[0]!r}")
    for current, accepted in matched:  # type: ignore[misc]
        text = text.replace(current, accepted, 1)
    return text


def runtime_lock_state_touches(module_name: str) -> bool:
    return module_name in _HUNKS
