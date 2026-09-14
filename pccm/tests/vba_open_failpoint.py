#!/usr/bin/env python3
"""P10-R3 CLOSURE: take the Workbook_Open failpoint back out of ThisWorkbook.

WHY A REVERSAL. Phase-10 section 3 makes Workbook_Open failure safety a required
Windows scenario, and no runner technique can make the real apply fail
deterministically without a dialog or a workbook that cannot be put right. The
narrow closure round after 889b6b5 therefore gave the handler the project's
ordinary failpoint - one dormant modAppState.FailPointCheck after a successful
apply, armed only through the accepted automation seam - and read Err before the
release owner's own On Error clears it. Both changes live in ThisWorkbook.vba
only. Every control that pins production to an accepted tree keeps its claim the
way P10-RP and P10-2C established: taking the declared change back out must
reproduce the accepted bytes exactly, so the ONLY thing that moved is the
declared closure.

THE FRAGMENTS ARE GENERATED FROM THE REAL DIFF against 889b6b5, the starting
authority of the closure round, and applying them all must reproduce that file
byte for byte. ThisWorkbook.vba is LF; the fragments carry LF.
"""
from __future__ import annotations

# The tree the reversal must reproduce: the closure round's starting authority.
ACCEPTED_BEFORE_OPEN_FAILPOINT = "889b6b5"

DECLARED_OPEN_FAILPOINT_CHANGES = {
    "ThisWorkbook.vba":
        "Workbook_Open carries one dormant modAppState.FailPointCheck, named by a "
        "private constant, placed after a successful ProtectionApply so an injected "
        "failure exercises the whole failure path from a fully protected workbook; "
        "and the failure path reads Err.Description before calling the release "
        "owner, whose On Error statement clears Err (P10-R3 closure)",
}

# (current fragment, accepted fragment), in file order.
_HUNKS: dict[str, tuple[tuple[str, str], ...]] = {
    "ThisWorkbook.vba": (("' in the one place no static control globs.\n"
  "'\n"
  "' THE ONE FAILPOINT, AND WHY IT SITS AFTER THE APPLY. The settled contract calls\n"
  "' Workbook_Open failure safety a required Windows scenario, and nothing outside\n"
  "' this handler can make the real apply fail deterministically without a dialog\n"
  "' or a workbook that cannot be put right afterwards. So the handler carries the\n"
  "' project's ordinary failpoint, through the same owner every other command uses:\n"
  "' modAppState.FailPointCheck exits at once unless the accepted automation seam\n"
  "' has been begun with exactly this stage name, and no user command, button or\n"
  "' open ever begins it - a workbook opened by a person runs this handler with the\n"
  "' seam dormant and the check costs one comparison. It is placed AFTER a\n"
  "' successful apply so that what the injected failure proves is the whole of the\n"
  "' failure path from a fully protected workbook: the release, the application\n"
  "' state put back, the record, and the dialog withheld under automation.\n"
  'Private Const FAILPOINT_WORKBOOK_OPEN As String = "Phase10WorkbookOpen"\n'
  '\n',
  "' in the one place no static control globs.\n\n"),
 ('    If Not modProtection.ProtectionApply(detail) Then GoTo Failed\n'
  '    modAppState.FailPointCheck FAILPOINT_WORKBOOK_OPEN\n'
  '\n',
  '    If Not modProtection.ProtectionApply(detail) Then GoTo Failed\n\n'),
 ('Failed:\n'
  "    ' THE ERROR IS READ BEFORE ANYTHING ELSE RUNS. A runtime error arrives here\n"
  "    ' with detail empty and its description in Err, and Err is cleared by the\n"
  "    ' next On Error statement executed - which is the first line of the release\n"
  "    ' owner called below. Read after that call, the description would be gone\n"
  "    ' and the record would name nothing.\n"
  '    If Len(detail) = 0 Then detail = Err.Description\n'
  "    ' A HALF-PROTECTED WORKBOOK IS THE ONE OUTCOME TO AVOID. If protection could\n",
  "Failed:\n    ' A HALF-PROTECTED WORKBOOK IS THE ONE OUTCOME TO AVOID. If protection could\n"),
 ('    End If\n\n', '    End If\n    If Len(detail) = 0 Then detail = Err.Description\n\n')),
}


def strip_open_failpoint(module_name: str, text: str) -> str:
    """`text` with the P10-R3 closure removed.

    ALL OR NONE, exactly as strip_repair_reconstruction: a module the closure
    never touched, or a text that predates it, carries NONE of the fragments and
    comes back unchanged; a text carrying ALL of them has the layer taken off
    exactly; a text carrying SOME of them raises, because "unchanged" must never
    be the answer to a partial match.
    """
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
            f"{module_name}: the declared Workbook_Open failpoint closure is partially present - "
            f"{len(hunks) - len(missing)} of {len(hunks)} fragments match, so the reversal "
            f"cannot be exact. A fragment moved or something rode along inside it:\n  {missing[0]!r}")
    for current, accepted in matched:  # type: ignore[misc]
        text = text.replace(current, accepted, 1)
    return text


def open_failpoint_touches(module_name: str) -> bool:
    return module_name in _HUNKS
