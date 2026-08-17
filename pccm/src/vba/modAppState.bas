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

Public Sub RestoreAppState(ByRef Snapshot As AppStateSnapshot)
    If Not Snapshot.Captured Then Exit Sub
    On Error Resume Next
    Application.Calculation = Snapshot.Calculation
    Application.DisplayAlerts = Snapshot.DisplayAlerts
    Application.EnableEvents = Snapshot.EnableEvents
    Application.ScreenUpdating = Snapshot.ScreenUpdating
    Application.StatusBar = Snapshot.StatusBar
    On Error GoTo 0
End Sub

' Forces a full recalculation of the structural-state formulas after the applied
' triple changes, while calculation is still manual.
Public Sub RecalculateStructuralState()
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_SETUP).Calculate
    ThisWorkbook.Worksheets(SH_COST_PROFILING).Calculate
    ThisWorkbook.Worksheets(SH_RISK_PROFILING).Calculate
    ThisWorkbook.Worksheets(SH_INFLATION).Calculate
    On Error GoTo 0
End Sub

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
' The harness must be able to exercise the destructive-confirmation path without
' a human clicking a dialog, and to inject a controlled mid-operation failure to
' prove logical restore. Both are inert unless explicitly set, both are reset by
' ClearAutomation, and neither can be reached from the UI.
Public gAutomationActive          As Boolean
Public gAutomationConfirmReply    As Boolean
Public gAutomationLastPrompt      As String
Public gAutomationFailAfterStage  As String
Public gAutomationLastResult      As String

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
