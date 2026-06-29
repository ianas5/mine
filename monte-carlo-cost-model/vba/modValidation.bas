Attribute VB_Name = "modValidation"
'==============================================================================
' modValidation
' Reads the Checks sheet (which holds all the live validation formulas) and
' returns True only if every check passes. Keeping the logic in-sheet means the
' same checks work with or without macros.
'==============================================================================
Option Explicit

Public Function ValidateInputs() As Boolean
    Dim ws As Worksheet, c As Range, anyFail As Boolean
    On Error GoTo Fail
    Set ws = ThisWorkbook.Sheets("Checks")
    anyFail = False
    ' Column B holds PASS / FAIL results; scan the used range.
    For Each c In ws.Range("B5:B100").Cells
        If UCase$(CStr(c.Value)) = "FAIL" Then anyFail = True
    Next c
    ValidateInputs = Not anyFail
    Exit Function
Fail:
    ValidateInputs = False
End Function

' Optional: jump the user to the first failing check.
Public Sub GoToFirstFailure()
    Dim ws As Worksheet, c As Range
    Set ws = ThisWorkbook.Sheets("Checks")
    For Each c In ws.Range("B5:B100").Cells
        If UCase$(CStr(c.Value)) = "FAIL" Then
            Application.Goto c, True
            Exit Sub
        End If
    Next c
    MsgBox "All checks pass.", vbInformation
End Sub
