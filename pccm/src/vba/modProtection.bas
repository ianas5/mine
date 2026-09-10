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
' protects a sheet against the user while leaving it writable by code, which is
' what lets every accepted command go on writing exactly as it did before
' protection existed. Excel does NOT persist that flag across a file close: a
' reopened workbook comes back protected against code as well. Re-applying it on
' open is the whole reason ThisWorkbook has an event, and it is why no command
' needs to unprotect anything.

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
