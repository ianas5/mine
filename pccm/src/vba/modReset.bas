Attribute VB_Name = "modReset"
Option Explicit

' ==========================================================================
' RESET RESULTS - THE ONE DESTRUCTIVE COMMAND THAT PRODUCES NOTHING
' ==========================================================================
' WHAT IT IS FOR. A workbook accumulates published answers: a deterministic
' calculation, a distribution in one of two banks, the annual decomposition of
' that distribution, and the sensitivity ranking that explains it. A user who
' wants to start again - because the model has moved on, because the results
' belong to a scenario they no longer want on the sheet, because they are about
' to hand the file to somebody else - has had no way to say so. Deleting the
' results by hand is exactly the accident protection exists to prevent.
'
' WHAT IT DOES NOT DO. It runs nothing, resolves nothing, publishes nothing and
' repairs nothing. It is not "recalculate": there is no output afterwards, and
' the way back is to press Calculate again.
'
' -------------------------------------------------------------------------
' THIS MODULE HOLDS NO GEOMETRY, AND THAT IS THE DESIGN
' -------------------------------------------------------------------------
' Not one bank letter, stamp row, block column, table name or fingerprint
' address appears below. Four owners already know where their publication
' lives, and each of them gained exactly two procedures - a clear and a restore
' - that speak in their own terms:
'
'   modCalcReport      the five analytical tables, the totals block, C13:C20
'   modSimReport       both banks, both summaries, both ladders, the selector
'   modSimAnnualStore  both annual blocks and their identity stamps
'   modSimPostReport   both sensitivity record blocks and their stamps
'
' What travels between them and this module is an OPAQUE undo carrier. This
' module never indexes into one, never inspects one and never builds one. A
' second opinion about where a publication lives is the one failure this design
' does not have.
'
' -------------------------------------------------------------------------
' WHAT IS PRESERVED, AND WHY IT IS NOT A LIST HERE EITHER
' -------------------------------------------------------------------------
' Every input - Setup, Config, Cost Lines, Risk Register, Inflation and both
' profiling grids - is preserved because NOTHING BELOW TOUCHES IT. There is no
' input-preservation routine here to get wrong: the clear scope is what the four
' owners write, and no owner of a publication is an owner of an input.
'
' The identity counters are preserved the same way. The permanent-id counters
' belong to modDrivers, the AUTO nonce and its pending marker to modSimNonce,
' and the run-id counter sits one row above the simulation publication record -
' deliberately outside the range that record clears. A discarded run must never
' be re-creatable by a future one, so a reset that renumbered anything would
' break the anti-replay guarantee it has no business touching.
'
' -------------------------------------------------------------------------
' NO STATE WORD IS WRITTEN
' -------------------------------------------------------------------------
' Nothing here decides that the calculation is NOT CALCULATED or that the annual
' output is NOT PRODUCED. Those readings are DERIVED by the owners that already
' own them, from inputs and from the absence of a publication. Clearing the
' publication is the whole of the action; the words follow from it. A reset that
' wrote them would be a second opinion about state.
' ==========================================================================

' The stages the Windows harness may fail this command after. They sit BETWEEN
' owners, because "after the calculation clear and before the simulation clear"
' is the boundary the acceptance scenarios name; no owner carries a failpoint of
' its own, and none of these is reachable unless automation is explicitly begun.
Public Const FAILPOINT_RESET_CALCULATION As String = "Phase10ResetCalculation"
Public Const FAILPOINT_RESET_SIMULATION As String = "Phase10ResetSimulation"
Public Const FAILPOINT_RESET_ANNUAL As String = "Phase10ResetAnnual"
Public Const FAILPOINT_RESET_SENSITIVITY As String = "Phase10ResetSensitivity"

' THE ONE PLACE THIS SENTENCE EXISTS. It is not repeated in a spec, in a second
' module or in a test fixture; anything that needs to know what the command says
' reads it from here.
Private Const RESET_SUCCEEDED As String = _
    "Results reset. Model inputs and identity counters were preserved."

' ==========================================================================
' THE ENDPOINT
' ==========================================================================
Public Sub PCCM_ResetResults()
    ' THE SAME ENVELOPE THE ACCEPTED ENDPOINTS USE, and for the same reason: a
    ' failure must not leave ScreenUpdating off or Calculation manual. There is
    ' no `committed` distinction here - a reset either completed or was rolled
    ' back, and neither outcome leaves anything for a cleanup problem to strand.
    Dim state As AppStateSnapshot, result As OperationResult
    Dim stateCaptured As Boolean, cleanupAttempted As Boolean
    Dim cleanup As String, failure As String

    On Error GoTo InvocationFailed
    state = modAppState.CaptureAppState()
    stateCaptured = True
    modAppState.BeginOperation
    result = ResetResults()
    On Error GoTo 0

    On Error GoTo NormalCleanupFailed
    cleanupAttempted = True
    cleanup = modAppState.FinishOperation(state)
    On Error GoTo 0
    stateCaptured = False

    If Len(cleanup) > 0 Then
        result = modAppState.Failed("Reset Results", _
            "Application state could not be restored: " & cleanup)
    End If
    modAppState.Announce result
    Exit Sub

NormalCleanupFailed:
    failure = Err.Description
    On Error GoTo 0
    modAppState.Announce modAppState.Failed("Reset Results", _
        "Application state could not be restored: " & failure)
    Exit Sub

InvocationFailed:
    failure = Err.Description
    On Error GoTo CleanupFailed
    If stateCaptured And Not cleanupAttempted Then
        cleanupAttempted = True
        cleanup = modAppState.FinishOperation(state)
        stateCaptured = False
        If Len(cleanup) > 0 Then failure = failure & vbCrLf & cleanup
    End If
    On Error GoTo 0
    modAppState.Announce modAppState.Failed("Reset Results", failure)
    Exit Sub

CleanupFailed:
    On Error GoTo 0
    modAppState.Announce modAppState.Failed("Reset Results", failure & vbCrLf & _
        "Application state could not be restored after the failure.")
End Sub

' ==========================================================================
' THE COMMAND
' ==========================================================================
Private Function ResetResults() As OperationResult
    Dim calcUndo As Variant, simUndo As Variant
    Dim annualUndo As Variant, sensitivityUndo As Variant
    Dim detail As String

    ' THE CONFIRMATION COMES BEFORE ANYTHING MOVES, which is why a cancellation
    ' needs no rollback and is not a failure. An empty success message is the
    ' accepted way to say "the user chose not to": the reporting owner shows no
    ' dialog for one, and the automation record still carries the outcome.
    If Not modAppState.AskConfirm(ConfirmationSummary(), True) Then
        ResetResults = modAppState.Succeeded(vbNullString)
        Exit Function
    End If

    If ClearEveryPublication(calcUndo, simUndo, annualUndo, sensitivityUndo, detail) Then
        ResetResults = modAppState.Succeeded(RESET_SUCCEEDED)
        Exit Function
    End If

    ResetResults = RollbackAndReport(calcUndo, simUndo, annualUndo, sensitivityUndo, detail)
End Function

' ONE ORDER, FOUR OWNERS, AND A CONTROLLED REFUSAL OR A RAISED ERROR ENDS IT THE
' SAME WAY. Every owner reports its own failure in its own words; nothing here
' interprets a message or infers an outcome from one.
Private Function ClearEveryPublication(ByRef calcUndo As Variant, ByRef simUndo As Variant, _
                                       ByRef annualUndo As Variant, _
                                       ByRef sensitivityUndo As Variant, _
                                       ByRef detail As String) As Boolean
    Dim failure As String

    On Error GoTo StepRaised
    If Not modCalcReport.CalcReportClearPublication(calcUndo, detail) Then Exit Function
    modAppState.FailPointCheck FAILPOINT_RESET_CALCULATION
    If Not modSimReport.SimReportClearPublication(simUndo, detail) Then Exit Function
    modAppState.FailPointCheck FAILPOINT_RESET_SIMULATION
    If Not modSimAnnualStore.SimAnnualStoreClearPublication(annualUndo, detail) Then Exit Function
    modAppState.FailPointCheck FAILPOINT_RESET_ANNUAL
    If Not modSimPostReport.SimPostReportClearPublication(sensitivityUndo, detail) Then Exit Function
    modAppState.FailPointCheck FAILPOINT_RESET_SENSITIVITY
    On Error GoTo 0

    ClearEveryPublication = True
    Exit Function

StepRaised:
    failure = Err.Description
    On Error GoTo 0
    detail = failure
End Function

' REVERSE ORDER, AND EVERY OWNER IS ASKED. An owner whose clear never ran holds
' an empty undo carrier and restores nothing, so no bookkeeping here has to
' decide which owners to skip - a decision that could be wrong in exactly the
' situation where being wrong costs the most.
Private Function RollbackAndReport(ByRef calcUndo As Variant, ByRef simUndo As Variant, _
                                   ByRef annualUndo As Variant, _
                                   ByRef sensitivityUndo As Variant, _
                                   ByVal detail As String) As OperationResult
    Dim note As String, problems As String

    On Error GoTo RollbackRaised
    If Not modSimPostReport.SimPostReportRestorePublication(sensitivityUndo, note) Then _
        problems = Appended(problems, note)
    If Not modSimAnnualStore.SimAnnualStoreRestorePublication(annualUndo, note) Then _
        problems = Appended(problems, note)
    If Not modSimReport.SimReportRestorePublication(simUndo, note) Then _
        problems = Appended(problems, note)
    If Not modCalcReport.CalcReportRestorePublication(calcUndo, note) Then _
        problems = Appended(problems, note)
    On Error GoTo 0

    If Len(problems) = 0 Then
        RollbackAndReport = modAppState.Failed("Reset Results", detail & vbCrLf & _
            "Every publication this command had cleared was put back, and no input, " & _
            "counter or identity was touched at any point.")
        Exit Function
    End If

    ' SAID PLAINLY, NEVER GLOSSED. A restore that did not complete is the one
    ' outcome a user has to act on, and both diagnostics are kept.
    RollbackAndReport = modAppState.Failed("Reset Results", detail & vbCrLf & _
        "The reset failed AND the previous publications could not be fully restored:" & _
        vbCrLf & problems)
    Exit Function

RollbackRaised:
    note = Err.Description
    On Error GoTo 0
    RollbackAndReport = modAppState.Failed("Reset Results", detail & vbCrLf & _
        "The reset failed AND the restore itself could not be completed: " & note)
End Function

Private Function Appended(ByVal existing As String, ByVal note As String) As String
    If Len(existing) = 0 Then
        Appended = note
    Else
        Appended = existing & vbCrLf & note
    End If
End Function

' WHAT THE USER IS AGREEING TO, IN THE TERMS THEY SEE ON THE SHEET. The
' reporting owner adds the permanence warning and the question; this says only
' what goes and what stays.
Private Function ConfirmationSummary() As String
    ConfirmationSummary = _
        "Reset Results clears every published result in this workbook:" & vbCrLf & _
        "    the deterministic calculation and its fingerprint" & vbCrLf & _
        "    both simulation publication banks" & vbCrLf & _
        "    the annual distributions and the selected-Px profile" & vbCrLf & _
        "    the sensitivity ranking" & vbCrLf & vbCrLf & _
        "Nothing you have typed is affected. Setup, Config, the cost and risk " & _
        "registers, inflation and both profiling grids are left exactly as they " & _
        "are, and the permanent-id, run-id and nonce counters keep their values " & _
        "so no identifier is ever reissued." & vbCrLf & vbCrLf & _
        "Reset produces no results. Press Calculate to build the model again."
End Function
