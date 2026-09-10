#!/usr/bin/env python3
"""P10-RP: take the structural protection window back out of production VBA.

WHY A REVERSAL RATHER THAN A NEW PIN. Windows disproved the accepted Phase-10
assumption that UserInterfaceOnly:=True lets every command mutate a ListObject
on a protected sheet: PCCM_ApplyTimeline was invoked and refused with

    Error 1004: Table features aren't available because the sheet is protected.

Six production modules therefore had to change, and every control that pinned
them to the accepted 58b2394 tree was going to be false. The two honest ways to
settle that are:

    move the pin to the new bytes, and lose the claim that nothing else moved;
    or remove the declared change mechanically and require the accepted bytes back.

The second is stronger, and it is the settlement P9-2B and P10-2B already
established here. A pin that moves says "these are the bytes now". A reversal
says "the ONLY thing that changed is the declared reconciliation, because taking
it out reproduces the accepted file exactly" - which is the claim the control
exists to make.

WHY FRAGMENTS RATHER THAN A SUFFIX. The P10-2B plumbing was APPENDED, so its
reversal could take a prefix. This change is INSERTED - a declaration in a
declaration section, a field in a Type, an envelope call in place of the one
beside it - so the reversal is a list of exact (current, accepted) fragments per
module. Each fragment carries enough context to appear exactly once in each
file, and applying them all must reproduce the accepted bytes byte-for-byte. A
line that rode along inside a hunk changes the fragment and fails to match; a
line that rode along anywhere else survives the reversal and fails the digest.

LINE ENDINGS ARE PART OF THE FRAGMENT. modCalcReport.bas is CRLF and every other
module is LF, so the fragments are stored with the endings their own file uses.
Rewriting one of these files through a text round-trip silently converts it, and
that is a real defect this batch made and a control caught.
"""
from __future__ import annotations

# The commit the reversal must reproduce: the accepted tree this batch started
# from, before the runtime protection reconciliation.
ACCEPTED_BEFORE_RECONCILIATION = "58b2394"

# Every production module the reconciliation touches, and why it had to move.
DECLARED_STRUCTURAL_WINDOW_CHANGES = {
    "modProtection.bas":
        "the sole protection owner gains the depth-safe structural window and the "
        "record of what Windows disproved",
    "modAppState.bas":
        "the operation envelope gains BeginStructuralOperation and closes the window "
        "in FinishOperation on every exit path",
    "modTimeline.bas": "PCCM_ApplyTimeline declares itself structural",
    "modDrivers.bas": "Add/Delete Cost Line and Risk declare themselves structural",
    "modCalcReport.bas": "PCCM_Calculate declares itself structural",
    "modRepair.bas": "PCCM_RepairProfiling declares itself structural",
}

# (current fragment, accepted fragment) per module, generated from the real diff
# and verified to reproduce the accepted bytes exactly.
_HUNKS: dict[str, tuple[tuple[str, str], ...]] = {
    'modProtection.bas': (
        ('\' USERINTERFACEONLY, AND WHY IT MUST BE RE-APPLIED. `UserInterfaceOnly:=True`\n\' protects a sheet against the user while leaving it writable by code. Excel does\n\' NOT persist that flag across a file close: a reopened workbook comes back\n\' protected against code as well. Re-applying it on open is the whole reason\n\' ThisWorkbook has an event.\n\'\n\' ==========================================================================\n\' PHASE-10 RUNTIME RECONCILIATION - WHAT WINDOWS DISPROVED\n\' ==========================================================================\n\' THE ORIGINAL ASSUMPTION, AND IT IS NOT BEING REWRITTEN. The accepted Phase-10\n\' contract (6ab8f6a) held that UserInterfaceOnly:=True let every accepted command\n\' go on writing exactly as it did before protection existed, so no command would\n\' ever need to unprotect anything. That was a reasonable reading of what\n\' UserInterfaceOnly means, it was recorded honestly, and it was WRONG about one\n\' specific capability. It stands as historical evidence.\n\'\n\' WHAT THE PROTECTION PROBE ACTUALLY OBSERVED, on Windows 11 / Excel 64-bit, with\n\' 14 of 14 worksheets protected, workbook structure protected and UIOnly applied:\n\'\n\'   * a code-driven VALUE write to a proved-locked cell SUCCEEDED, read back, and\n\'     restored exactly, with protection still in force afterwards;\n\'   * the real PCCM_ApplyTimeline endpoint was invoked and REFUSED:\n\'\n\'       Error 1004: Table features aren\'t available because the sheet is\n\'       protected.\n\'\n\'     with protection 14/14 before and after and NO structural effect.\n\'\n\' SO UserInterfaceOnly SEPARATES TWO CAPABILITIES, and only now is that written\n\' down as fact rather than assumption:\n\'\n\'   ordinary code-driven cell VALUE writes .............. PERMITTED\n\'   ListObject STRUCTURAL mutations while protected ..... NOT PERMITTED\n\'\n\' THE RECONCILED RULE. The old absolute "no command may call Unprotect" is\n\' replaced by a narrower one that this module alone can honour:\n\'\n\'   Only modProtection may temporarily release WORKSHEET protection, inside the\n\'   contracted structural-operation envelope, and it must restore the full\n\'   accepted protection state before the operation returns.\n\'\n\' WORKBOOK STRUCTURE PROTECTION IS NOT RELEASED. The 1004 named the SHEET, and\n\' workbook-structure protection governs adding, deleting and renaming WORKSHEETS\n\' rather than table columns. Releasing it would be a wider change than any\n\' evidence asks for, so the window does not touch it and a control refuses one\n\' that does. If a future Windows run proves a specific need, that is its own\n\' reconciliation.\n\'\n\' AND NOTHING HERE BECOMES A USER-FACING CAPABILITY. There is still no password,\n\' still no interactive unprotected mode, and no button opens this window. From\n\' the user\'s side the workbook is protected before a command, protected after it,\n\' and the release exists only for the instants a command is executing code.\n\n\' ---------------------------------------------------------------------------\n\' THE STRUCTURAL WINDOW\'S DEPTH\n\' ---------------------------------------------------------------------------\n\' IN THE DECLARATION SECTION, WHERE VBA REQUIRES IT. Everything before the first\n\' executable procedure is the declaration section; a Private further down is a\n\' "Compile error: Variable not defined" that no Linux control would see and that\n\' ended Gate-B run 3 once already.\n\'\n\' DEPTH-COUNTED, AND THE COUNT LIVES HERE. Nothing else may hold it: a caller\n\' that tracked its own depth would be a second protection authority, and the two\n\' would disagree the first time an operation exited by a path someone forgot.\n\'\n\' The outermost Begin releases worksheet protection once. Inner Begins only\n\' count. The outermost End re-applies the whole accepted policy - every sheet,\n\' UserInterfaceOnly:=True, structure - and PROVES it took before saying so.\nPrivate mStructuralDepth As Long\n',
         "' USERINTERFACEONLY, AND WHY IT MUST BE RE-APPLIED. `UserInterfaceOnly:=True`\n' protects a sheet against the user while leaving it writable by code, which is\n' what lets every accepted command go on writing exactly as it did before\n' protection existed. Excel does NOT persist that flag across a file close: a\n' reopened workbook comes back protected against code as well. Re-applying it on\n' open is the whole reason ThisWorkbook has an event, and it is why no command\n' needs to unprotect anything.\n"),
        ('\n\' ==========================================================================\n\' THE STRUCTURAL WINDOW\n\' ==========================================================================\n\' How many opens are outstanding. Reported so a control, a diagnostic or a caller\n\' can see the window\'s state without inferring it from a worksheet.\nPublic Function ProtectionStructuralDepth() As Long\n    ProtectionStructuralDepth = mStructuralDepth\nEnd Function\n\nPublic Function ProtectionInStructuralWindow() As Boolean\n    ProtectionInStructuralWindow = (mStructuralDepth > 0)\nEnd Function\n\n\' OPEN. Returns False and a reason if the release could not be established; the\n\' caller must then NOT attempt the structural work.\nPublic Function ProtectionBeginStructural(ByRef detail As String) As Boolean\n    Dim sheet As Worksheet\n    detail = vbNullString\n\n    \' NESTED: count only. Re-releasing sheets that are already released would be\n    \' harmless, but re-protecting them at an inner End would not be, and one of\n    \' those follows from the other.\n    If mStructuralDepth > 0 Then\n        mStructuralDepth = mStructuralDepth + 1\n        ProtectionBeginStructural = True\n        Exit Function\n    End If\n\n    On Error GoTo Failed\n    For Each sheet In ThisWorkbook.Worksheets\n        If sheet.ProtectContents Then sheet.Unprotect\n    Next sheet\n\n    \' PROVED, NOT ASSUMED. A command that began structural work on a sheet that\n    \' was still protected would fail with the same 1004 this window exists to\n    \' remove, and would look like the defect rather than a failed release.\n    Set sheet = Nothing\n    For Each sheet In ThisWorkbook.Worksheets\n        If sheet.ProtectContents Then\n            detail = "worksheet protection was not released from " & sheet.Name\n            GoTo Failed\n        End If\n    Next sheet\n\n    mStructuralDepth = 1\n    ProtectionBeginStructural = True\n    Exit Function\n\nFailed:\n    \' A HALF-OPEN WINDOW IS THE WORST OUTCOME THERE IS: some sheets released, no\n    \' depth recorded, and therefore nobody who will close it. Protection is\n    \' re-applied here and now, and the failure is still reported.\n    If Len(detail) = 0 Then\n        detail = "the structural protection window could not be opened"\n        If Not sheet Is Nothing Then detail = detail & " at " & sheet.Name\n        detail = detail & ": " & Err.Description\n    End If\n    mStructuralDepth = 0\n    Dim reapply As String\n    If Not ProtectionApply(reapply) Then\n        detail = detail & "; protection could not be re-applied either: " & reapply\n    End If\nEnd Function\n\n\' CLOSE. Returns False and a reason if the accepted protection state could not be\n\' restored. That is a delivery failure and the caller reports it as one: the\n\' business mutation may well have committed, and an unprotected workbook handed\n\' back as an ordinary success is exactly what must not happen.\nPublic Function ProtectionEndStructural(ByRef detail As String) As Boolean\n    detail = vbNullString\n\n    If mStructuralDepth <= 0 Then\n        detail = "the structural protection window was closed more often than it was opened"\n        mStructuralDepth = 0\n        Exit Function\n    End If\n\n    mStructuralDepth = mStructuralDepth - 1\n    If mStructuralDepth > 0 Then\n        \' AN INNER CLOSE RESTORES NOTHING. The outer operation is still running\n        \' and may still have structural work and a rollback ahead of it.\n        ProtectionEndStructural = True\n        Exit Function\n    End If\n\n    \' OUTERMOST. The whole accepted policy comes back - every sheet,\n    \' UserInterfaceOnly:=True, and the structure flag - through the same one\n    \' function that establishes it on open. There is no second policy here.\n    If Not ProtectionApply(detail) Then Exit Function\n\n    \' AND THE OWNER MUST AGREE. ProtectionApply reporting True is what it TRIED;\n    \' ProtectionIsApplied is what the workbook now SAYS.\n    If Not ProtectionIsApplied() Then\n        detail = "protection was re-applied but the workbook does not report itself protected"\n        Exit Function\n    End If\n\n    ProtectionEndStructural = True\nEnd Function\n\n\' THE MAINTENANCE PATH. It exists so a maintainer is never forced to reach for a',
         "\n' THE MAINTENANCE PATH. It exists so a maintainer is never forced to reach for a"),
    ),
    'modAppState.bas': (
        ("    Captured        As Boolean\n    ' PHASE-10 RUNTIME RECONCILIATION. True only when THIS operation opened the\n    ' structural protection window. It rides in the snapshot rather than in a\n    ' module-level counter here so that open and close are paired per operation\n    ' by the same value: FinishOperation cannot close a window this operation did\n    ' not open, and cannot close one twice. modProtection still owns the depth.\n    Structural      As Boolean\n    ScreenUpdating  As Boolean",
         '    Captured        As Boolean\n    ScreenUpdating  As Boolean'),
        ('    Dim s As AppStateSnapshot\n    s.Structural = False\n    s.ScreenUpdating = Application.ScreenUpdating',
         '    Dim s As AppStateSnapshot\n    s.ScreenUpdating = Application.ScreenUpdating'),
        ("' a workbook picks one up from elsewhere.\n'\n' THIS OPENS NO PROTECTION WINDOW, and that is the point. A command that does not\n' perform ListObject structural work has no business leaving worksheets\n' unprotected - least of all a stochastic run that holds the workbook for\n' minutes. Those commands keep calling this, unchanged.\nPublic Sub BeginOperation()",
         "' a workbook picks one up from elsewhere.\nPublic Sub BeginOperation()"),
        ('    Application.Calculation = xlCalculationManual\nEnd Sub\n\n\' THE STRUCTURAL ENVELOPE, AND THE ONLY WAY TO ASK FOR ONE.\n\'\n\' Windows proved that UserInterfaceOnly:=True does NOT permit ListObject\n\' structural mutation on a protected sheet: PCCM_ApplyTimeline was invoked and\n\' refused with "Error 1004: Table features aren\'t available because the sheet is\n\' protected." A command whose forward work OR whose transactional rollback can\n\' add or delete a ListRow or a ListColumn calls this instead of BeginOperation.\n\'\n\' IT IS A DECLARATION, NOT A CONVENIENCE. Each of the seven structural commands\n\' names itself here by calling this; the four non-structural ones are visibly\n\' still on BeginOperation. There is no default and no flag to forget.\n\'\n\' A FAILED OPEN RAISES rather than returning. Every caller already has an error\n\' handler that rolls back and routes through FinishOperation, and attempting\n\' structural work through a window that did not open would produce the very 1004\n\' this exists to prevent - reported as a mysterious command failure.\nPublic Sub BeginStructuralOperation(ByRef Snapshot As AppStateSnapshot)\n    BeginOperation\n\n    Dim detail As String\n    If Not modProtection.ProtectionBeginStructural(detail) Then\n        Err.Raise vbObjectError + 5010, "modAppState.BeginStructuralOperation", _\n                  "The structural protection window could not be opened, so no " & _\n                  "structural work was attempted: " & detail\n    End If\n    Snapshot.Structural = True\nEnd Sub',
         '    Application.Calculation = xlCalculationManual\nEnd Sub'),
        ('    Dim recalcProblem As String\n\n    \' THE WINDOW CLOSES FIRST, AND ON EVERY PATH THAT REACHES HERE - success,\n    \' refusal, validation error, runtime error, injected failure, and after a\n    \' rollback. Every command\'s rollback runs BEFORE its cleanup, so the window is\n    \' still open while RestoreTable rebuilds a table, which is the one ordering\n    \' that must not be got wrong.\n    \'\n    \' THE FLAG IS CLEARED BEFORE THE CLOSE IS ATTEMPTED, so a path that reached\n    \' cleanup twice cannot decrement modProtection\'s depth twice.\n    If Snapshot.Structural Then\n        Snapshot.Structural = False\n        Dim protectionProblem As String\n        If Not modProtection.ProtectionEndStructural(protectionProblem) Then\n            \' A DELIVERY FAILURE, NOT A FOOTNOTE. This lands in the same string\n            \' every caller already treats as an unsafe-cleanup failure, so a\n            \' command whose protection could not be restored can no longer be\n            \' announced as an ordinary success.\n            problems = problems & "  PROTECTION WAS NOT RESTORED: " & _\n                       protectionProblem & vbCrLf\n        End If\n    End If\n',
         '    Dim recalcProblem As String\n'),
    ),
    'modTimeline.bas': (
        ("    On Error GoTo Failure\n    ' STRUCTURAL, AND IT SAYS SO. Apply / Update Timeline adds the project-year ListColumns to both profiling\n    ' grids and the inflation grid, and its rollback rebuilds all three tables.\n    ' Windows proved a protected sheet refuses all of that with Error 1004, so\n    ' this command opens the protection window that FinishOperation closes on\n    ' every exit path - including after the rollback below.\n    modAppState.BeginStructuralOperation snapshot\n",
         '    On Error GoTo Failure\n    modAppState.BeginOperation\n'),
    ),
    'modDrivers.bas': (
        ("    On Error GoTo Failure\n    ' STRUCTURAL, AND IT SAYS SO. Add and Delete Cost Line / Risk add and delete register ListRows past the\n    ' reserved block, SyncRows reshapes the profiling grid on every one of them,\n    ' and the rollback rebuilds both tables.\n    ' Windows proved a protected sheet refuses all of that with Error 1004, so\n    ' this command opens the protection window that FinishOperation closes on\n    ' every exit path - including after the rollback below.\n    modAppState.BeginStructuralOperation snapshot\n",
         '    On Error GoTo Failure\n    modAppState.BeginOperation\n'),
    ),
    'modCalcReport.bas': (
        ("    stateCaptured = True\r\n    ' STRUCTURAL: ResizeBody adds and deletes _Calc ListRows and the rollback\r\n    ' rebuilds five tables, all of which a protected sheet refuses with 1004.\r\n    ' FinishOperation closes the window on every path, rollback included.\r\n    modAppState.BeginStructuralOperation state\r\n    result = RunCalculation(committed)",
         '    stateCaptured = True\r\n    modAppState.BeginOperation\r\n    result = RunCalculation(committed)'),
    ),
    'modRepair.bas': (
        ("    stateCaptured = True\n    ' STRUCTURAL, AND IT SAYS SO. Repair Profiling rewrites the project-year ListColumns and re-syncs the\n    ' profiling rows, and its rollback rebuilds both grids.\n    ' Windows proved a protected sheet refuses all of that with Error 1004, so\n    ' this command opens the protection window that FinishOperation closes on\n    ' every exit path - including after the rollback below.\n    modAppState.BeginStructuralOperation state\n    result = RepairProfiling()",
         '    stateCaptured = True\n    modAppState.BeginOperation\n    result = RepairProfiling()'),
    ),
}


def strip_structural_window(module_name: str, text: str) -> str:
    """`text` with this batch's declared reconciliation removed.

    A module the reconciliation never touched comes back unchanged, so the same
    reversal can be applied to every owner without asking first which ones moved.

    A fragment that no longer matches raises rather than silently returning the
    text unchanged: "this file was not touched" is the one answer that would turn
    every control built on this into a control that passes over anything.
    """
    for current, accepted in _HUNKS.get(module_name, ()):
        # THE SAME FRAGMENT IN WHICHEVER ENDING THE CALLER IS HOLDING. Some
        # controls read these files as bytes and some through read_text, which
        # performs universal-newline translation - so a CRLF module arrives here
        # with LF endings and the stored fragment would match nothing. Matching
        # nothing must never mean "unchanged", so the variants are tried
        # explicitly and a genuine miss still raises.
        for pair in ((current, accepted),
                     (current.replace("\r\n", "\n"), accepted.replace("\r\n", "\n"))):
            if text.count(pair[0]) == 1:
                text = text.replace(pair[0], pair[1], 1)
                break
        else:
            found = text.count(current)
            raise AssertionError(
                f"{module_name}: the declared structural-window fragment appears "
                f"{found} times, so the reversal cannot be exact. Either the change "
                f"moved or something rode along inside it:\n  {current.splitlines()[0]!r}")
    return text


def structural_window_touches(module_name: str) -> bool:
    return module_name in _HUNKS
