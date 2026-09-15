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
'
' THE ONE FAILPOINT, AND WHY IT SITS AFTER THE APPLY. The settled contract calls
' Workbook_Open failure safety a required Windows scenario, and nothing outside
' this handler can make the real apply fail deterministically without a dialog
' or a workbook that cannot be put right afterwards. So the handler carries the
' project's ordinary failpoint, through the same owner every other command uses:
' modAppState.FailPointCheck exits at once unless the accepted automation seam
' has been begun with exactly this stage name, and no user command, button or
' open ever begins it - a workbook opened by a person runs this handler with the
' seam dormant and the check costs one comparison. It is placed AFTER a
' successful apply so that what the injected failure proves is the whole of the
' failure path from a fully protected workbook: the release, the application
' state put back, the record, and the dialog withheld under automation.
Private Const FAILPOINT_WORKBOOK_OPEN As String = "Phase10WorkbookOpen"

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
    ' THE DASHBOARD CHARTS' CATEGORY AXES, BOUND FOR WHAT IS PUBLISHED. It runs
    ' here because this is the first moment a person can see the Dashboard, and
    ' AFTER the apply above because binding a chart needs the accepted
    ' structural window, which protection has to be in force to open. A binding
    ' that could not be applied is a presentation fault, not a failed open: it
    ' is recorded and the workbook opens.
    Dim chartDetail As String
    If Not modChartPresentation.ChartPresentationApplyCategories(chartDetail) Then
        modAppState.RecordResult "Workbook_Open: " & chartDetail
    End If
    modAppState.FailPointCheck FAILPOINT_WORKBOOK_OPEN

    Application.ScreenUpdating = previousUpdating
    restored = True
    Exit Sub

Failed:
    ' THE ERROR IS READ BEFORE ANYTHING ELSE RUNS. A runtime error arrives here
    ' with detail empty and its description in Err, and Err is cleared by the
    ' next On Error statement executed - which is the first line of the release
    ' owner called below. Read after that call, the description would be gone
    ' and the record would name nothing.
    If Len(detail) = 0 Then detail = Err.Description
    ' A HALF-PROTECTED WORKBOOK IS THE ONE OUTCOME TO AVOID. If protection could
    ' not be established, it is released rather than left partly applied: a
    ' workbook the user can edit is recoverable, and one where some sheets are
    ' protected against code is not.
    If Not restored Then Application.ScreenUpdating = previousUpdating
    Dim releaseDetail As String
    If Not modProtection.ProtectionRelease(releaseDetail) Then
        detail = detail & " (and it could not be released again: " & releaseDetail & ")"
    End If

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
