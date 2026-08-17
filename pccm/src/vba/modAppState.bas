Attribute VB_Name = "modAppState"
Option Explicit

' ===========================================================================
' PCCM - Application state, failure reporting and user confirmation
' ===========================================================================
' Phase 4. Structural runtime only. No cost, risk, escalation, FX, NPV, EMV,
' simulation, RNG, Model Check, sensitivity or results logic belongs here.
'
' Two responsibilities:
'   1. Save and restore Excel application state around a structural operation,
'      so a failure never leaves the user with ScreenUpdating off.
'   2. Present confirmations and failures in one consistent voice.
'
' There is deliberately no transaction framework, no undo journal and no
' workbook snapshot system here. Logical restoration of the specific blocks an
' operation touches is the requirement; see modTimeline.RestorePoint.
' ===========================================================================

Public Type AppStateSnapshot
    Captured        As Boolean
    ScreenUpdating  As Boolean
    EnableEvents    As Boolean
    DisplayAlerts   As Boolean
    Calculation     As XlCalculation
    StatusBar       As Variant
End Type

Public Type OperationResult
    Ok          As Boolean
    Message     As String
    Detail      As String
End Type

Private Const MSG_TITLE As String = "PCCM"

' ---------------------------------------------------------------------------
' Windows functional-harness state.
' ---------------------------------------------------------------------------
' The harness must be able to exercise the destructive-confirmation path without
' a human clicking a dialog, and to inject a controlled mid-operation failure to
' prove logical restore. Both are inert unless PCCM_AutomationBegin explicitly
' enables automation, both are reset by ClearAutomation, and neither can be
' reached from the UI.
'
' THESE BELONG IN THE DECLARATION SECTION, and that is not a style preference.
' VBA has no "module-level statement anywhere in the file": everything before the
' first executable procedure is the declaration section, and everything after it
' is procedure bodies. These five sat after ConfirmDestructiveChange, so under
' Option Explicit the compiler reached PCCM_AutomationBegin with gAutomationActive
' undefined and stopped with "Compile error: Variable not defined". That is what
' ended Gate-B run 3, on the harness's very first VBA call.
'
' Public, deliberately: modTimeline and modDrivers reference
' modAppState.gAutomationActive by qualified name.
Public gAutomationActive          As Boolean
Public gAutomationConfirmReply    As Boolean
Public gAutomationLastPrompt      As String
Public gAutomationFailAfterStage  As String
Public gAutomationLastResult      As String

' ---------------------------------------------------------------------------
' Application state
' ---------------------------------------------------------------------------
Public Function CaptureAppState() As AppStateSnapshot
    Dim s As AppStateSnapshot
    s.ScreenUpdating = Application.ScreenUpdating
    s.EnableEvents = Application.EnableEvents
    s.DisplayAlerts = Application.DisplayAlerts
    s.Calculation = Application.Calculation
    s.StatusBar = Application.StatusBar
    s.Captured = True
    CaptureAppState = s
End Function

' Quietens the UI for the duration of a structural operation. Events are
' suppressed because structural runtime must not trigger anything: PCCM installs
' no input Worksheet_Change handler, and this guarantees that stays true even if
' a workbook picks one up from elsewhere.
Public Sub BeginOperation()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
End Sub

' Restores every captured property and REPORTS whether all of them succeeded.
'
' Returns "" on complete success, otherwise a description of every property that
' could not be restored. It attempts them ALL rather than stopping at the first
' failure: leaving ScreenUpdating off because DisplayAlerts failed first would be
' the worst possible outcome.
'
' A successful structural mutation followed by a failed application-state
' restoration is NOT a successful operation, and the caller surfaces it as an
' unsafe-cleanup failure.
Public Function RestoreAppState(ByRef Snapshot As AppStateSnapshot) As String
    If Not Snapshot.Captured Then
        RestoreAppState = "no application state was captured; nothing could be restored"
        Exit Function
    End If

    Dim failures As String
    failures = failures & TryRestoreCalculation(Snapshot.Calculation)
    failures = failures & TryRestoreDisplayAlerts(Snapshot.DisplayAlerts)
    failures = failures & TryRestoreEnableEvents(Snapshot.EnableEvents)
    failures = failures & TryRestoreScreenUpdating(Snapshot.ScreenUpdating)
    failures = failures & TryRestoreStatusBar(Snapshot.StatusBar)
    RestoreAppState = failures
End Function

' One property per function, each with its own handler, so a failure is attributed
' precisely and never prevents the remaining properties from being restored.
Private Function TryRestoreCalculation(ByVal Value As XlCalculation) As String
    On Error GoTo Failed
    Application.Calculation = Value
    Exit Function
Failed:
    TryRestoreCalculation = "    Calculation: " & Err.Description & vbCrLf
End Function

Private Function TryRestoreDisplayAlerts(ByVal Value As Boolean) As String
    On Error GoTo Failed
    Application.DisplayAlerts = Value
    Exit Function
Failed:
    TryRestoreDisplayAlerts = "    DisplayAlerts: " & Err.Description & vbCrLf
End Function

Private Function TryRestoreEnableEvents(ByVal Value As Boolean) As String
    On Error GoTo Failed
    Application.EnableEvents = Value
    Exit Function
Failed:
    TryRestoreEnableEvents = "    EnableEvents: " & Err.Description & vbCrLf
End Function

Private Function TryRestoreScreenUpdating(ByVal Value As Boolean) As String
    On Error GoTo Failed
    Application.ScreenUpdating = Value
    Exit Function
Failed:
    TryRestoreScreenUpdating = "    ScreenUpdating: " & Err.Description & vbCrLf
End Function

Private Function TryRestoreStatusBar(ByVal Value As Variant) As String
    On Error GoTo Failed
    Application.StatusBar = Value
    Exit Function
Failed:
    TryRestoreStatusBar = "    StatusBar: " & Err.Description & vbCrLf
End Function

' Recalculates the structural-state formulas after the applied triple changes,
' while calculation is still manual. Returns "" on success, otherwise a
' description.
'
' A recalculation failure is NOT swallowed. On the success path the caller turns it
' into a controlled failure; on the cleanup path the caller appends it to the
' restore report WITHOUT hiding the original error that caused the rollback.
Public Function RecalculateStructuralState() As String
    On Error GoTo Failed
    ThisWorkbook.Worksheets(SH_SETUP).Calculate
    ThisWorkbook.Worksheets(SH_COST_PROFILING).Calculate
    ThisWorkbook.Worksheets(SH_RISK_PROFILING).Calculate
    ThisWorkbook.Worksheets(SH_INFLATION).Calculate
    Exit Function
Failed:
    RecalculateStructuralState = "structural-state recalculation failed. Error " & _
                                 Err.Number & ": " & Err.Description
End Function

' The shared end-of-operation cleanup. Recalculates, restores application state, and
' returns "" only when BOTH succeeded. Every command routes its cleanup through here
' so no path can quietly skip either half.
Public Function FinishOperation(ByRef Snapshot As AppStateSnapshot) As String
    Dim problems As String
    Dim recalcProblem As String

    recalcProblem = RecalculateStructuralState()
    If Len(recalcProblem) > 0 Then
        problems = problems & "  " & recalcProblem & vbCrLf
    End If

    Dim restoreProblem As String
    restoreProblem = RestoreAppState(Snapshot)
    If Len(restoreProblem) > 0 Then
        problems = problems & "  application state could not be fully restored:" & vbCrLf & _
                   restoreProblem
    End If

    FinishOperation = problems
End Function

' ---------------------------------------------------------------------------
' Results
' ---------------------------------------------------------------------------
Public Function Succeeded(ByVal Message As String) As OperationResult
    Dim r As OperationResult
    r.Ok = True
    r.Message = Message
    Succeeded = r
End Function

Public Function Failed(ByVal Message As String, ByVal Detail As String) As OperationResult
    Dim r As OperationResult
    r.Ok = False
    r.Message = Message
    r.Detail = Detail
    Failed = r
End Function

Public Sub ReportResult(ByRef Result As OperationResult)
    If Result.Ok Then
        If Len(Result.Message) > 0 Then
            MsgBox Result.Message, vbInformation, MSG_TITLE
        End If
    Else
        Dim body As String
        body = Result.Message
        If Len(Result.Detail) > 0 Then
            body = body & vbCrLf & vbCrLf & Result.Detail
        End If
        MsgBox body, vbExclamation, MSG_TITLE
    End If
End Sub

' A structural failure is never silent and is never converted into a partial
' success. The caller has already restored the blocks it touched.
Public Sub ReportFailure(ByVal Operation As String, ByVal Reason As String, _
                         ByVal RestoreNote As String)
    Dim body As String
    body = Operation & " FAILED." & vbCrLf & vbCrLf & _
           Reason & vbCrLf & vbCrLf & RestoreNote
    MsgBox body, vbCritical, MSG_TITLE
End Sub

' ---------------------------------------------------------------------------
' Confirmation
' ---------------------------------------------------------------------------
Public Function ConfirmStructuralChange(ByVal Summary As String) As Boolean
    ConfirmStructuralChange = _
        (MsgBox(Summary & vbCrLf & vbCrLf & "Continue?", _
                vbOKCancel + vbQuestion, MSG_TITLE) = vbOK)
End Function

' Destructive confirmation is shown BEFORE anything is modified, which is why a
' cancellation needs no rollback: nothing has moved yet.
Public Function ConfirmDestructiveChange(ByVal Summary As String) As Boolean
    ConfirmDestructiveChange = _
        (MsgBox(Summary & vbCrLf & vbCrLf & _
                "This data will be PERMANENTLY DELETED and cannot be recovered." & vbCrLf & _
                "Continue?", _
                vbOKCancel + vbExclamation, MSG_TITLE) = vbCancel) = False
End Function

' ---------------------------------------------------------------------------
' Automation hooks for the Windows functional harness.
' ---------------------------------------------------------------------------
' The five gAutomation* variables these procedures use are declared at the top of
' this module, in the declaration section, where VBA requires them.
Public Sub ClearAutomation()
    gAutomationActive = False
    gAutomationConfirmReply = False
    gAutomationLastPrompt = vbNullString
    gAutomationFailAfterStage = vbNullString
    gAutomationLastResult = vbNullString
End Sub

Public Function AskConfirm(ByVal Summary As String, ByVal Destructive As Boolean) As Boolean
    gAutomationLastPrompt = Summary
    If gAutomationActive Then
        AskConfirm = gAutomationConfirmReply
        Exit Function
    End If
    If Destructive Then
        AskConfirm = ConfirmDestructiveChange(Summary)
    Else
        AskConfirm = ConfirmStructuralChange(Summary)
    End If
End Function

' Raises a controlled error when the harness has armed this stage. Real runs never
' arm it, so this costs one string comparison and changes nothing.
Public Sub FailPointCheck(ByVal StageName As String)
    If Not gAutomationActive Then Exit Sub
    If Len(gAutomationFailAfterStage) = 0 Then Exit Sub
    If StrComp(gAutomationFailAfterStage, StageName, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 5001, "modAppState.FailPointCheck", _
                  "Injected structural failure after stage '" & StageName & "'."
    End If
End Sub

Public Sub RecordResult(ByVal Text As String)
    gAutomationLastResult = Text
End Sub

' --- harness-callable surface ----------------------------------------------
' Named PCCM_ so the static tests can enumerate every externally callable
' procedure. All of these are inert in normal use: nothing in the workbook calls
' PCCM_AutomationBegin, so gAutomationActive stays False and every dialog behaves
' exactly as a user sees it.
Public Sub PCCM_AutomationBegin(ByVal ConfirmReply As Boolean, ByVal FailAfterStage As String)
    ClearAutomation
    gAutomationActive = True
    gAutomationConfirmReply = ConfirmReply
    gAutomationFailAfterStage = FailAfterStage
End Sub

Public Sub PCCM_AutomationEnd()
    ClearAutomation
End Sub

Public Function PCCM_AutomationResult() As String
    PCCM_AutomationResult = gAutomationLastResult
End Function

Public Function PCCM_AutomationPrompt() As String
    PCCM_AutomationPrompt = gAutomationLastPrompt
End Function

Public Sub Announce(ByRef Result As OperationResult)
    If Result.Ok Then
        RecordResult "OK|" & Result.Message
    Else
        RecordResult "FAIL|" & Result.Message & "|" & Result.Detail
    End If
    If Not gAutomationActive Then ReportResult Result
End Sub
