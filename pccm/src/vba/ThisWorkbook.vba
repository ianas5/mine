Option Explicit

' ==========================================================================
' ThisWorkbook - THE PROJECT'S ONLY DOCUMENT MODULE
' ==========================================================================
' WHY THIS FILE IS NOT A .bas. ThisWorkbook is a DOCUMENT module: Excel creates
' it with the workbook and it cannot be imported like a standard module. The
' bootstrap writes this text into the existing component instead, which is also
' why the extension is deliberately not .bas - the module inventory globs .bas
' and this is not a module in that inventory.
'
' WHAT THIS EVENT IS FOR, AND IT IS ONE THING. Sheet protection applied with
' UserInterfaceOnly:=True does not survive a file close. A reopened workbook
' comes back protected against CODE as well as against the user, and every
' accepted command would start failing on its first write. Re-applying it on
' open is the only reason this event exists.
'
' WHAT IT DOES NOT DO. It does not calculate, simulate, run sensitivity or the
' annual step, derive or inspect model state, read or write a publication, touch
' a fingerprint, repair anything, or reset anything. It shows no dialog when it
' succeeds - an operator opening a workbook has asked for a workbook, not for a
' message - and it makes no decision of its own.
'
' AND IT HOLDS NO POLICY. Which sheets are protected and which cells are
' unlocked is modProtection's question and the contracts' answer. This is a
' four-line delegation on purpose: policy in a document module would be policy
' in the one place no static control globs.

Private Sub Workbook_Open()
    Dim detail As String
    Dim previousUpdating As Boolean
    Dim restored As Boolean

    ' STATE IS CAPTURED BEFORE ANYTHING IS CHANGED and restored on every path,
    ' including the error path. An open handler that leaves ScreenUpdating off
    ' leaves Excel looking frozen, and the user has no command to run to fix it.
    On Error GoTo Failed
    previousUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    restored = False

    If Not modProtection.ProtectionApply(detail) Then GoTo Failed

    Application.ScreenUpdating = previousUpdating
    restored = True
    Exit Sub

Failed:
    ' A HALF-PROTECTED WORKBOOK IS THE ONE OUTCOME TO AVOID. If protection could
    ' not be established, it is released rather than left partly applied: a
    ' workbook the user can edit is recoverable, and one where some sheets are
    ' protected against code is not.
    If Not restored Then Application.ScreenUpdating = previousUpdating
    Dim releaseDetail As String
    If Not modProtection.ProtectionRelease(releaseDetail) Then
        detail = detail & " (and it could not be released again: " & releaseDetail & ")"
    End If
    If Len(detail) = 0 Then detail = Err.Description

    ' REPORTED THE WAY EVERY OTHER COMMAND REPORTS. The accepted idiom is that
    ' the CALLER asks whether automation is active and ReportFailure stays
    ' unconditional - modDrivers and modTimeline both guard it exactly here, and
    ' this follows them rather than changing the owner they share. A harness
    ' opening this workbook would otherwise deadlock on a modal dialog nobody is
    ' there to dismiss, which is the one failure an open handler must not have.
    modAppState.RecordResult "Workbook_Open: " & detail
    If Not modAppState.gAutomationActive Then
        modAppState.ReportFailure "Opening the workbook", detail, _
            "The workbook is usable and unprotected. Protection will be " & _
            "attempted again the next time it is opened."
    End If
End Sub
