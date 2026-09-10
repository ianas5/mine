Attribute VB_Name = "modProtection"
Option Explicit

' ==========================================================================
' PROTECTION - THE ONE OWNER
' ==========================================================================
' WHAT THIS IS FOR. A delivered workbook is handed to someone who did not build
' it. Protection stops them overwriting a formula, a header, a permanent id or a
' published result by accident - by clicking the wrong cell, by pasting a block
' one row too high, by dragging a fill handle. That is the whole goal.
'
' IT IS NOT A SECURITY CONTROL AND IS NEVER DESCRIBED AS ONE. There is no
' password here and there is no password anywhere: Excel sheet protection is
' trivially removable and a password kept in source would be worse than none,
' because it would claim something that is not true.
'
' WHY THE POLICY IS NOT IN THIS FILE. Which cells a user may type in is decided
' by the contracts that already declare them - `editable: true` on an input, on
' a register column, on a config table, and the year columns of a grid. The
' builder resolves those declarations and writes the locked/unlocked state into
' the workbook itself. So this module applies a UNIFORM action and holds no list:
' every sheet is protected, and what differs between them is already in the file.
' There is nothing here to drift from the schema, because there is nothing here
' that describes the schema.
'
' USERINTERFACEONLY, AND WHY IT MUST BE RE-APPLIED. `UserInterfaceOnly:=True`
' protects a sheet against the user while leaving it writable by code. Excel does
' NOT persist that flag across a file close: a reopened workbook comes back
' protected against code as well. Re-applying it on open is the whole reason
' ThisWorkbook has an event.
'
' ==========================================================================
' PHASE-10 RUNTIME RECONCILIATION - WHAT WINDOWS DISPROVED
' ==========================================================================
' THE ORIGINAL ASSUMPTION, AND IT IS NOT BEING REWRITTEN. The accepted Phase-10
' contract (6ab8f6a) held that UserInterfaceOnly:=True let every accepted command
' go on writing exactly as it did before protection existed, so no command would
' ever need to unprotect anything. That was a reasonable reading of what
' UserInterfaceOnly means, it was recorded honestly, and it was WRONG about one
' specific capability. It stands as historical evidence.
'
' WHAT THE PROTECTION PROBE ACTUALLY OBSERVED, on Windows 11 / Excel 64-bit, with
' 14 of 14 worksheets protected, workbook structure protected and UIOnly applied:
'
'   * a code-driven VALUE write to a proved-locked cell SUCCEEDED, read back, and
'     restored exactly, with protection still in force afterwards;
'   * the real PCCM_ApplyTimeline endpoint was invoked and REFUSED:
'
'       Error 1004: Table features aren't available because the sheet is
'       protected.
'
'     with protection 14/14 before and after and NO structural effect.
'
' SO UserInterfaceOnly SEPARATES TWO CAPABILITIES, and only now is that written
' down as fact rather than assumption:
'
'   ordinary code-driven cell VALUE writes .............. PERMITTED
'   ListObject STRUCTURAL mutations while protected ..... NOT PERMITTED
'
' THE RECONCILED RULE. The old absolute "no command may call Unprotect" is
' replaced by a narrower one that this module alone can honour:
'
'   Only modProtection may temporarily release WORKSHEET protection, inside the
'   contracted structural-operation envelope, and it must restore the full
'   accepted protection state before the operation returns.
'
' WORKBOOK STRUCTURE PROTECTION IS NOT RELEASED. The 1004 named the SHEET, and
' workbook-structure protection governs adding, deleting and renaming WORKSHEETS
' rather than table columns. Releasing it would be a wider change than any
' evidence asks for, so the window does not touch it and a control refuses one
' that does. If a future Windows run proves a specific need, that is its own
' reconciliation.
'
' AND NOTHING HERE BECOMES A USER-FACING CAPABILITY. There is still no password,
' still no interactive unprotected mode, and no button opens this window. From
' the user's side the workbook is protected before a command, protected after it,
' and the release exists only for the instants a command is executing code.

' ---------------------------------------------------------------------------
' THE STRUCTURAL WINDOW'S DEPTH
' ---------------------------------------------------------------------------
' IN THE DECLARATION SECTION, WHERE VBA REQUIRES IT. Everything before the first
' executable procedure is the declaration section; a Private further down is a
' "Compile error: Variable not defined" that no Linux control would see and that
' ended Gate-B run 3 once already.
'
' DEPTH-COUNTED, AND THE COUNT LIVES HERE. Nothing else may hold it: a caller
' that tracked its own depth would be a second protection authority, and the two
' would disagree the first time an operation exited by a path someone forgot.
'
' The outermost Begin releases worksheet protection once. Inner Begins only
' count. The outermost End re-applies the whole accepted policy - every sheet,
' UserInterfaceOnly:=True, structure - and PROVES it took before saying so.
Private mStructuralDepth As Long

' The single fact this module reports about itself. A caller that wants to know
' whether protection is in force asks here rather than testing a sheet.
Public Function ProtectionIsApplied() As Boolean
    Dim sheet As Worksheet
    For Each sheet In ThisWorkbook.Worksheets
        If Not sheet.ProtectContents Then Exit Function
    Next sheet
    ProtectionIsApplied = ThisWorkbook.ProtectStructure
End Function

' APPLY, AND SAY WHETHER IT TOOK. Every sheet, then the structure. A sheet that
' is already protected is re-protected anyway: the UserInterfaceOnly flag is the
' point of the call, and a sheet that came back from a file close protected
' WITHOUT it would look identical to one that has it.
'
' NO PASSWORD ARGUMENT. Not an empty one, not a constant, not a variable.
Public Function ProtectionApply(ByRef detail As String) As Boolean
    Dim sheet As Worksheet
    detail = vbNullString
    On Error GoTo Failed
    For Each sheet In ThisWorkbook.Worksheets
        ' Unprotect first so the flag is re-established rather than assumed. A
        ' sheet protected without UserInterfaceOnly cannot be upgraded in place.
        If sheet.ProtectContents Then sheet.Unprotect
        sheet.Protect UserInterfaceOnly:=True, _
                      DrawingObjects:=True, Contents:=True, Scenarios:=False, _
                      AllowFormattingCells:=False
    Next sheet
    If Not ThisWorkbook.ProtectStructure Then
        ThisWorkbook.Protect Structure:=True, Windows:=False
    End If
    ProtectionApply = True
    Exit Function
Failed:
    ' THE SHEET IS NAMED, because "protection failed" is not a diagnosis. A
    ' caller that cannot say which sheet refused cannot tell a locked file from
    ' a renamed one.
    detail = "protection could not be applied"
    If Not sheet Is Nothing Then detail = detail & " to " & sheet.Name
    detail = detail & ": " & Err.Description
End Function

' ==========================================================================
' THE STRUCTURAL WINDOW
' ==========================================================================
' How many opens are outstanding. Reported so a control, a diagnostic or a caller
' can see the window's state without inferring it from a worksheet.
Public Function ProtectionStructuralDepth() As Long
    ProtectionStructuralDepth = mStructuralDepth
End Function

Public Function ProtectionInStructuralWindow() As Boolean
    ProtectionInStructuralWindow = (mStructuralDepth > 0)
End Function

' OPEN. Returns False and a reason if the release could not be established; the
' caller must then NOT attempt the structural work.
Public Function ProtectionBeginStructural(ByRef detail As String) As Boolean
    Dim sheet As Worksheet
    detail = vbNullString

    ' NESTED: count only. Re-releasing sheets that are already released would be
    ' harmless, but re-protecting them at an inner End would not be, and one of
    ' those follows from the other.
    If mStructuralDepth > 0 Then
        mStructuralDepth = mStructuralDepth + 1
        ProtectionBeginStructural = True
        Exit Function
    End If

    On Error GoTo Failed
    For Each sheet In ThisWorkbook.Worksheets
        If sheet.ProtectContents Then sheet.Unprotect
    Next sheet

    ' PROVED, NOT ASSUMED. A command that began structural work on a sheet that
    ' was still protected would fail with the same 1004 this window exists to
    ' remove, and would look like the defect rather than a failed release.
    Set sheet = Nothing
    For Each sheet In ThisWorkbook.Worksheets
        If sheet.ProtectContents Then
            detail = "worksheet protection was not released from " & sheet.Name
            GoTo Failed
        End If
    Next sheet

    mStructuralDepth = 1
    ProtectionBeginStructural = True
    Exit Function

Failed:
    ' A HALF-OPEN WINDOW IS THE WORST OUTCOME THERE IS: some sheets released, no
    ' depth recorded, and therefore nobody who will close it. Protection is
    ' re-applied here and now, and the failure is still reported.
    If Len(detail) = 0 Then
        detail = "the structural protection window could not be opened"
        If Not sheet Is Nothing Then detail = detail & " at " & sheet.Name
        detail = detail & ": " & Err.Description
    End If
    mStructuralDepth = 0
    Dim reapply As String
    If Not ProtectionApply(reapply) Then
        detail = detail & "; protection could not be re-applied either: " & reapply
    End If
End Function

' CLOSE. Returns False and a reason if the accepted protection state could not be
' restored. That is a delivery failure and the caller reports it as one: the
' business mutation may well have committed, and an unprotected workbook handed
' back as an ordinary success is exactly what must not happen.
Public Function ProtectionEndStructural(ByRef detail As String) As Boolean
    detail = vbNullString

    If mStructuralDepth <= 0 Then
        detail = "the structural protection window was closed more often than it was opened"
        mStructuralDepth = 0
        Exit Function
    End If

    mStructuralDepth = mStructuralDepth - 1
    If mStructuralDepth > 0 Then
        ' AN INNER CLOSE RESTORES NOTHING. The outer operation is still running
        ' and may still have structural work and a rollback ahead of it.
        ProtectionEndStructural = True
        Exit Function
    End If

    ' OUTERMOST. The whole accepted policy comes back - every sheet,
    ' UserInterfaceOnly:=True, and the structure flag - through the same one
    ' function that establishes it on open. There is no second policy here.
    If Not ProtectionApply(detail) Then Exit Function

    ' AND THE OWNER MUST AGREE. ProtectionApply reporting True is what it TRIED;
    ' ProtectionIsApplied is what the workbook now SAYS.
    If Not ProtectionIsApplied() Then
        detail = "protection was re-applied but the workbook does not report itself protected"
        Exit Function
    End If

    ProtectionEndStructural = True
End Function

' THE MAINTENANCE PATH. It exists so a maintainer is never forced to reach for a
' password that does not exist, and it is deliberately not bound to a button:
' releasing protection is a maintenance act, not a user command.
Public Function ProtectionRelease(ByRef detail As String) As Boolean
    Dim sheet As Worksheet
    detail = vbNullString
    On Error GoTo Failed
    If ThisWorkbook.ProtectStructure Then ThisWorkbook.Unprotect
    For Each sheet In ThisWorkbook.Worksheets
        If sheet.ProtectContents Then sheet.Unprotect
    Next sheet
    ProtectionRelease = True
    Exit Function
Failed:
    detail = "protection could not be released"
    If Not sheet Is Nothing Then detail = detail & " from " & sheet.Name
    detail = detail & ": " & Err.Description
End Function
