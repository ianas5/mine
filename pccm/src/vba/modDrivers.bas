Attribute VB_Name = "modDrivers"
Option Explicit

' ===========================================================================
' PCCM - Permanent identity and driver row operations
' ===========================================================================
' Phase 4. Structure only. Nothing here reads a unit cost, an impact, a
' probability or a distribution; those are user inputs the register owns and no
' calculation consumes them yet.
'
' Permanent identity rules, all of them load-bearing:
'
'   * An identifier comes from a PERSISTENT COUNTER on _Calc and from nowhere
'     else. Never from a row number, never from a row count.
'   * Deletion NEVER decrements the counter, so an identifier is never reused:
'     Add, Add, Delete CL-002, Add yields CL-003 and not CL-002.
'   * Pad width is a minimum display width and nothing more: CL-999 is followed by
'     CL-1000, so there is no three-digit cap. The VBA implementation does have a
'     representational sequence ceiling (ID_COUNTER_MAX); reaching it stops NEW
'     allocation and is not a limit on how many drivers the model may currently hold,
'     nor a reason to call the existing structure invalid.
'   * Reordering the register moves rows, not identity: an ID travels with its
'     row data because it lives in that row.
' ===========================================================================

Private Const KIND_COST As String = "cost"
Private Const KIND_RISK As String = "risk"

' ---------------------------------------------------------------------------
' Identity
' ---------------------------------------------------------------------------
Public Function CounterName(ByVal Kind As String) As String
    If Kind = KIND_COST Then CounterName = NM_COUNTER_COST_LINE Else CounterName = NM_COUNTER_RISK
End Function

Public Function IdPrefix(ByVal Kind As String) As String
    If Kind = KIND_COST Then IdPrefix = ID_PREFIX_COST_LINE Else IdPrefix = ID_PREFIX_RISK
End Function

Public Function IdPad(ByVal Kind As String) As Long
    If Kind = KIND_COST Then IdPad = ID_PAD_COST_LINE Else IdPad = ID_PAD_RISK
End Function

Public Function FormatId(ByVal Kind As String, ByVal Sequence As Long) As String
    Dim digits As String
    digits = CStr(Sequence)
    Do While Len(digits) < IdPad(Kind)
        digits = "0" & digits
    Loop
    FormatId = IdPrefix(Kind) & digits
End Function

' The persistent counter is the model's HISTORICAL MEMORY of what has been issued.
' A corrupt counter must never silently become zero.
'
' Falling back to 0 was unsafe in a specific and reachable way:
'
'   CL-001 issued -> CL-001 deleted -> no IDs remain in the register
'   -> counter corrupted to blank or text
'   -> ReadCounter returns 0, HighestIssued from current rows is also 0
'   -> validation passes, and the next Add reissues CL-001.
'
' Current rows cannot testify about deleted history, so an invalid counter is now an
' explicit failure rather than a value.
Public Function TryReadCounter(ByVal Kind As String, ByRef Value As Long) As Boolean
    Value = 0
    Dim raw As Variant, d As Double
    raw = modWorkbook.ReadValue(CounterName(Kind))
    If IsEmpty(raw) Then Exit Function
    If Not modWorkbook.IsWholeInRange(raw, 0, ID_COUNTER_MAX, d) Then Exit Function
    Value = modWorkbook.SafeLong(d)
    TryReadCounter = True
End Function

' There is deliberately NO ReadCounter()-style accessor that maps invalid state to
' zero. One existed, and its only caller was the operation snapshot, where it did
' exactly the damage the counter-integrity protection was added to prevent:
'
'   R-001 issued -> every Risk deleted -> counter corrupted to text
'   -> the snapshot read it through the lossy accessor as 0
'   -> AllocateId correctly refused
'   -> rollback wrote the snapshotted 0 back
'   -> the corrupt historical counter became a valid zero, and R-001 could be reissued.
'
' Callers that ALLOCATE use TryReadCounter and refuse invalid state. Callers that
' SNAPSHOT use RawCounter and restore the value byte for byte.
Public Function RawCounter(ByVal Kind As String) As Variant
    RawCounter = modWorkbook.ReadValue(CounterName(Kind))
End Function

' Advances the counter and returns the identifier it just issued.
'
' Refuses outright when the stored counter is invalid, and refuses CLEANLY at the
' representational ceiling rather than overflowing on counter + 1. The ceiling is an
' implementation limit of the VBA Long, not a business maximum on identifiers.
Public Function AllocateId(ByVal Kind As String) As String
    Dim current As Long
    If Not TryReadCounter(Kind, current) Then
        Err.Raise vbObjectError + 5020, "modDrivers.AllocateId", _
                  "The persistent " & KindLabel(Kind) & " ID counter at " & _
                  CounterName(Kind) & " is missing, blank or not a whole number within " & _
                  "0-" & ID_COUNTER_MAX & ". It is the model's record of every identifier " & _
                  "ever issued, including deleted ones, so allocation cannot continue " & _
                  "without it: guessing would risk reissuing an identifier that has " & _
                  "already been used. Restore the counter and try again."
    End If
    If current >= ID_COUNTER_MAX Then
        Err.Raise vbObjectError + 5021, "modDrivers.AllocateId", _
                  "The " & KindLabel(Kind) & " ID counter has reached " & ID_COUNTER_MAX & _
                  ", the largest sequence this implementation can represent. No " & _
                  "identifier was allocated. This is a representation ceiling, not a " & _
                  "limit on how many drivers the model may hold."
    End If

    Dim nextSequence As Long
    nextSequence = current + 1
    ' Persisted immediately so a later failure cannot hand the same number out twice.
    modWorkbook.WriteValue CounterName(Kind), nextSequence
    AllocateId = FormatId(Kind, nextSequence)
End Function

' The largest sequence present in the CURRENT rows. Zero when there are none, which
' says nothing about history -- that is what the persistent counter is for.
'
' Unrepresentable tails are reported through Unrepresentable rather than skipped: an
' ID such as CL-99999999999 is corrupt, and silently ignoring it would let it sit in
' the register unnoticed.
'
' Unrepresentable is REQUIRED, not Optional. `Optional ByRef Unrepresentable As Long`
' is a compile error in VBA: a typed Optional parameter must declare a default, and
' only a Variant Optional may be omitted without one. Every caller supplies the
' argument anyway, so the optional semantics bought nothing and cost a build.
Public Function HighestIssued(ByVal Kind As String, _
                              ByRef Unrepresentable As Long) As Long
    Dim register As ListObject
    Dim r As Long, rowCount As Long, idCol As Long, best As Long
    Dim prefix As String, idText As String, tail As String
    Dim seq As Double

    Set register = RegisterTable(Kind)
    idCol = IdColumn(Kind)
    rowCount = modWorkbook.BodyRowCount(register)
    prefix = IdPrefix(Kind)
    Unrepresentable = 0

    For r = 1 To rowCount
        idText = modWorkbook.TextOf(modWorkbook.CellIn(register, r, idCol))
        If Len(idText) > Len(prefix) Then
            If StrComp(Left$(idText, Len(prefix)), prefix, vbTextCompare) = 0 Then
                tail = Mid$(idText, Len(prefix) + 1)
                ' Bounded before conversion: a pasted identifier such as CL-99999999999
                ' must not overflow the scan that exists to detect exactly that corruption.
                If modWorkbook.IsWholeInRange(tail, 0, ID_COUNTER_MAX, seq) Then
                    If modWorkbook.SafeLong(seq) > best Then best = modWorkbook.SafeLong(seq)
                ElseIf IsAllDigitsText(tail) Then
                    Unrepresentable = Unrepresentable + 1
                End If
            End If
        End If
    Next r
    HighestIssued = best
End Function

Public Function IsAllDigitsText(ByVal Text As String) As Boolean
    Dim i As Long
    If Len(Text) = 0 Then Exit Function
    For i = 1 To Len(Text)
        If Mid$(Text, i, 1) < "0" Or Mid$(Text, i, 1) > "9" Then Exit Function
    Next i
    IsAllDigitsText = True
End Function

' ---------------------------------------------------------------------------
' Register geometry
' ---------------------------------------------------------------------------
Public Function RegisterTable(ByVal Kind As String) As ListObject
    If Kind = KIND_COST Then
        Set RegisterTable = modWorkbook.Lo(SH_COST_LINES, TBL_COST_LINES)
    Else
        Set RegisterTable = modWorkbook.Lo(SH_RISK_REGISTER, TBL_RISK_REGISTER)
    End If
End Function

Public Function IdColumn(ByVal Kind As String) As Long
    If Kind = KIND_COST Then IdColumn = COL_COST_LINES_COST_LINE_ID Else IdColumn = COL_RISK_REGISTER_RISK_ID
End Function

Private Function TraceColumn(ByVal Kind As String) As Long
    If Kind = KIND_COST Then TraceColumn = COL_COST_LINES_DESCRIPTION Else TraceColumn = COL_RISK_REGISTER_RISK_NAME
End Function

Private Function FirstEditableColumn(ByVal Kind As String) As Long
    ' The identity column is column 1 and is model-controlled, so the first field
    ' the user actually owns is column 2.
    FirstEditableColumn = 2
End Function

Private Function KindLabel(ByVal Kind As String) As String
    If Kind = KIND_COST Then KindLabel = "Cost Line" Else KindLabel = "Risk"
End Function

' ---------------------------------------------------------------------------
' Row search
' ---------------------------------------------------------------------------
Public Function RowOfId(ByVal Kind As String, ByVal PermanentId As String) As Long
    Dim register As ListObject
    Dim r As Long, rowCount As Long, idCol As Long
    Set register = RegisterTable(Kind)
    idCol = IdColumn(Kind)
    rowCount = modWorkbook.BodyRowCount(register)
    For r = 1 To rowCount
        If StrComp(modWorkbook.TextOf(modWorkbook.CellIn(register, r, idCol)), PermanentId, vbTextCompare) = 0 Then
            RowOfId = r
            Exit Function
        End If
    Next r
End Function

' The first reserved row that is genuinely blank across EVERY column. A row that
' already holds user content but carries no ID is an orphan, not free space.
Private Function FirstFreeRow(ByVal Kind As String, ByRef OrphanRow As Long) As Long
    Dim register As ListObject
    Dim r As Long, c As Long, rowCount As Long, colCount As Long
    Dim idCol As Long, hasContent As Boolean

    Set register = RegisterTable(Kind)
    idCol = IdColumn(Kind)
    rowCount = modWorkbook.BodyRowCount(register)
    colCount = register.ListColumns.Count
    OrphanRow = 0

    ' The WHOLE register is scanned before a row is chosen. Returning at the first
    ' blank row meant an orphan at row 1 was recorded and then ignored because row 2
    ' happened to be free -- the operation proceeded and the orphan survived,
    ' unowned and invisible to every keyed assessment.
    Dim firstBlank As Long
    For r = 1 To rowCount
        If Len(modWorkbook.TextOf(modWorkbook.CellIn(register, r, idCol))) = 0 Then
            hasContent = False
            For c = 1 To colCount
                If c <> idCol Then
                    If Not modWorkbook.IsEmptyCell(modWorkbook.CellIn(register, r, c)) Then
                        hasContent = True
                        Exit For
                    End If
                End If
            Next c
            If hasContent Then
                ' Do not silently take ownership of somebody's typing.
                If OrphanRow = 0 Then OrphanRow = r
            ElseIf firstBlank = 0 Then
                firstBlank = r
            End If
        End If
    Next r

    ' An orphan anywhere blocks the operation, even when a blank row is available.
    If OrphanRow > 0 Then Exit Function
    FirstFreeRow = firstBlank
End Function

' ---------------------------------------------------------------------------
' Add
' ---------------------------------------------------------------------------
Public Function AddDriver(ByVal Kind As String) As OperationResult
    Dim register As ListObject
    Dim targetRow As Long, orphanRow As Long
    Dim newId As String

    Set register = RegisterTable(Kind)
    targetRow = FirstFreeRow(Kind, orphanRow)

    If orphanRow > 0 Then
        AddDriver = modAppState.Failed( _
            "Add " & KindLabel(Kind) & " could not continue.", _
            "Row " & orphanRow & " of " & register.Name & " already contains data but has " & _
            "no permanent identifier. Adopting it would attach a model-controlled ID to " & _
            "content the model never issued one for, so nothing has been changed. Clear " & _
            "that row, or move its content into a properly added row, and try again.")
        Exit Function
    End If

    modAppState.FailPointCheck "add.before_allocate"

    newId = AllocateId(Kind)

    modAppState.FailPointCheck "add.after_allocate"

    If targetRow = 0 Then
        ' No reserved capacity left: grow the table. Reserved rows were only ever
        ' initial capacity, never a business maximum.
        register.ListRows.Add
        targetRow = modWorkbook.BodyRowCount(register)
    End If

    modWorkbook.CellIn(register, targetRow, IdColumn(Kind)).Value = newId

    modAppState.FailPointCheck "add.after_write_id"

    modProfiling.SyncRows Kind

    modAppState.FailPointCheck "add.after_sync"

    ' Put the cursor on the first field the user actually owns. This is the only
    ' cosmetic step in the operation, and it is the only place a suppressed error is
    ' acceptable: failing to move the selection must not fail an otherwise complete
    ' structural change. Nothing that alters data is inside this guard.
    On Error Resume Next
    register.Parent.Activate
    modWorkbook.CellIn(register, targetRow, FirstEditableColumn(Kind)).Select
    On Error GoTo 0

    AddDriver = modAppState.Succeeded(KindLabel(Kind) & " " & newId & " added.")
End Function

' ---------------------------------------------------------------------------
' Delete
' ---------------------------------------------------------------------------
' The permanent ID is the operation key from the first line to the last. The row
' number is resolved once for the prompt and resolved AGAIN after confirmation,
' so a row that moved in between cannot cause the wrong driver to be deleted.
Public Function DeleteDriver(ByVal Kind As String, ByVal PermanentId As String) As OperationResult
    Dim register As ListObject
    Dim rowIndex As Long
    Dim traceText As String, summary As String

    Set register = RegisterTable(Kind)

    If Len(PermanentId) = 0 Then
        DeleteDriver = modAppState.Failed( _
            "Delete " & KindLabel(Kind) & " could not continue.", _
            "Select a row that has a permanent identifier. A reserved blank row is not " & _
            "a driver and there is nothing to delete.")
        Exit Function
    End If

    rowIndex = RowOfId(Kind, PermanentId)
    If rowIndex = 0 Then
        DeleteDriver = modAppState.Failed( _
            "Delete " & KindLabel(Kind) & " could not continue.", _
            "No row in " & register.Name & " carries the identifier " & PermanentId & ".")
        Exit Function
    End If

    traceText = modWorkbook.TextOf(modWorkbook.CellIn(register, rowIndex, TraceColumn(Kind)))
    summary = "Delete " & KindLabel(Kind) & " " & PermanentId & "?" & vbCrLf & vbCrLf & _
              "  Identifier      : " & PermanentId & vbCrLf & _
              "  " & IIf(Kind = KIND_COST, "Description     : ", "Risk Name       : ") & _
              IIf(Len(traceText) = 0, "(blank)", traceText) & vbCrLf & _
              "  Profiling row   : the matching row will also be removed" & vbCrLf & vbCrLf & _
              "The identifier " & PermanentId & " will NOT be reused by a future add."

    If Not modAppState.AskConfirm(summary, True) Then
        DeleteDriver = modAppState.Succeeded(vbNullString)
        Exit Function
    End If

    ' Re-resolve by identity. Never trust the earlier row number after a prompt.
    rowIndex = RowOfId(Kind, PermanentId)
    If rowIndex = 0 Then
        DeleteDriver = modAppState.Failed( _
            "Delete " & KindLabel(Kind) & " could not continue.", _
            "Identifier " & PermanentId & " is no longer present. Nothing was changed.")
        Exit Function
    End If

    modAppState.FailPointCheck "delete.before_remove"

    If modWorkbook.BodyRowCount(register) > 1 Then
        register.ListRows(rowIndex).Delete
    Else
        register.DataBodyRange.Rows(rowIndex).ClearContents
    End If

    modAppState.FailPointCheck "delete.after_remove"

    modProfiling.RemoveRow Kind, PermanentId
    modProfiling.SyncRows Kind

    ' The counter is deliberately NOT decremented.
    DeleteDriver = modAppState.Succeeded( _
        KindLabel(Kind) & " " & PermanentId & " deleted. The identifier will not be reused.")
End Function

' ---------------------------------------------------------------------------
' Selection
' ---------------------------------------------------------------------------
' The permanent ID of the row the user has selected in the register, or "".
Public Function SelectedId(ByVal Kind As String) As String
    Dim register As ListObject
    Dim body As Range, hit As Range
    Dim rowIndex As Long

    Set register = RegisterTable(Kind)
    If register.DataBodyRange Is Nothing Then Exit Function
    ' Intersect raises when the selection is on another sheet; that simply means "no
    ' driver row is selected", which the next line handles. No data is touched here.
    On Error Resume Next
    Set body = Application.Intersect(Selection, register.DataBodyRange)
    On Error GoTo 0
    If body Is Nothing Then Exit Function

    Set hit = body.Cells(1, 1)
    rowIndex = hit.Row - register.DataBodyRange.Row + 1
    If rowIndex < 1 Or rowIndex > modWorkbook.BodyRowCount(register) Then Exit Function
    SelectedId = modWorkbook.TextOf(modWorkbook.CellIn(register, rowIndex, IdColumn(Kind)))
End Function

' ---------------------------------------------------------------------------
' Entry points bound to the command buttons
' ---------------------------------------------------------------------------
Public Sub PCCM_AddCostLine()
    RunDriverOperation KIND_COST, True, vbNullString
End Sub

Public Sub PCCM_AddRisk()
    RunDriverOperation KIND_RISK, True, vbNullString
End Sub

Public Sub PCCM_DeleteCostLine()
    RunDeleteCommand KIND_COST
End Sub

Public Sub PCCM_DeleteRisk()
    RunDeleteCommand KIND_RISK
End Sub

' Selection resolution happens INSIDE a protected shell, never as an argument
' expression evaluated before the command is entered.
'
' Writing `RunDriverOperation Kind, False, SelectedId(Kind)` looked equivalent but is
' not: VBA evaluates SelectedId BEFORE entering RunDriverOperation, so a failure
' while resolving the register, its DataBodyRange, the selection intersection or the
' selected identifier escaped the shared command handler entirely and surfaced as a
' raw VBA dialog.
Public Sub RunDeleteCommand(ByVal Kind As String)
    Dim permanentId As String
    On Error GoTo ResolveFailed
    permanentId = SelectedId(Kind)
    On Error GoTo 0

    RunDriverOperation Kind, False, permanentId
    Exit Sub

ResolveFailed:
    Dim reason As String
    reason = "Error " & Err.Number & ": " & Err.Description
    modAppState.RecordResult "FAIL|" & reason
    If Not modAppState.gAutomationActive Then
        modAppState.ReportFailure "Delete " & KindLabel(Kind), reason, _
            "The failure occurred while working out which row is selected, before the " & _
            "operation started. Nothing has been changed."
    End If
End Sub

' --- harness-callable surface ----------------------------------------------
' Deletion by explicit identifier rather than by worksheet selection. The button
' path resolves the selection to an identifier and then behaves identically; this
' pair simply removes the selection step so an automated run is deterministic.
Public Sub PCCM_DeleteCostLineById(ByVal PermanentId As String)
    RunDriverOperation KIND_COST, False, PermanentId
End Sub

Public Sub PCCM_DeleteRiskById(ByVal PermanentId As String)
    RunDriverOperation KIND_RISK, False, PermanentId
End Sub

' Shared driver-operation shell: capture app state, snapshot only the blocks this
' operation can modify, run, revalidate, and restore logically on failure.
Public Sub RunDriverOperation(ByVal Kind As String, ByVal IsAdd As Boolean, _
                              ByVal PermanentId As String)
    Dim snapshot As AppStateSnapshot
    Dim result As OperationResult
    Dim registerBefore As TableSnapshot, profilingBefore As TableSnapshot
    ' RAW, not Long. A corrupt counter must come back corrupt: converting it to a
    ' number here would let the rollback launder corruption into a valid zero.
    Dim counterBefore As Variant
    Dim problems As String
    Dim captured As Boolean
    Dim stateCaptured As Boolean

    ' The handler is installed BEFORE the first fallible operation, and capturing
    ' application state is itself fallible. stateCaptured records whether a snapshot
    ' actually exists, so no cleanup path can claim to restore one that does not.
    On Error GoTo AssessmentFailure
    snapshot = modAppState.CaptureAppState()
    stateCaptured = True

    ' The unkeyed-data gate, before any mutation. Nothing has changed yet, so this
    ' path needs no rollback.
    problems = modStructuralCheck.PreMutationCheck()
    If Len(problems) > 0 Then
        modAppState.Announce modAppState.Failed( _
            IIf(IsAdd, "Add ", "Delete ") & KindLabel(Kind) & _
            " was refused. Nothing has been changed.", problems)
        Exit Sub
    End If

    On Error GoTo Failure
    modAppState.BeginOperation

    registerBefore = modWorkbook.SnapshotTable(RegisterTable(Kind))
    profilingBefore = modWorkbook.SnapshotTable(modProfiling.ProfilingTable(Kind))
    counterBefore = RawCounter(Kind)
    captured = True

    If IsAdd Then
        result = AddDriver(Kind)
    Else
        result = DeleteDriver(Kind, PermanentId)
    End If

    If result.Ok Then
        problems = modStructuralCheck.ValidateStructure()
        If Len(problems) > 0 Then
            Err.Raise vbObjectError + 5002, "modDrivers.RunDriverOperation", _
                      "Structural revalidation failed:" & vbCrLf & problems
        End If
    End If

    Dim cleanup As String
    cleanup = FinishIfCaptured(snapshot, stateCaptured)
    If Len(cleanup) > 0 Then
        modAppState.Announce modAppState.Failed( _
            IIf(IsAdd, "Add ", "Delete ") & KindLabel(Kind) & _
            " completed, but the workbook was NOT left in a safe state.", _
            "Cleanup did not complete:" & vbCrLf & cleanup)
        Exit Sub
    End If

    modAppState.Announce result
    Exit Sub

AssessmentFailure:
    Dim assessReason As String
    assessReason = "Error " & Err.Number & ": " & Err.Description
    Dim assessCleanup As String
    assessCleanup = FinishIfCaptured(snapshot, stateCaptured)
    modAppState.RecordResult "FAIL|" & assessReason
    If Not modAppState.gAutomationActive Then
        modAppState.ReportFailure IIf(IsAdd, "Add ", "Delete ") & KindLabel(Kind), assessReason, _
            "The failure occurred before any change was made, so nothing needed to be " & _
            "rolled back." & _
            IIf(Len(assessCleanup) > 0, vbCrLf & vbCrLf & "Cleanup also reported:" & _
                vbCrLf & assessCleanup, "")
    End If
    Exit Sub

Failure:
    Dim reason As String, restoreNote As String
    reason = "Error " & Err.Number & ": " & Err.Description

    If captured Then
        restoreNote = TryRestoreDriver(Kind, registerBefore, profilingBefore, counterBefore)
    Else
        restoreNote = "Nothing had been modified when the failure occurred."
    End If

    Dim failureCleanup As String
    failureCleanup = FinishIfCaptured(snapshot, stateCaptured)
    If Len(failureCleanup) > 0 Then
        restoreNote = restoreNote & vbCrLf & vbCrLf & _
                      "Cleanup ALSO reported problems:" & vbCrLf & failureCleanup
    End If

    modAppState.RecordResult "FAIL|" & reason & "|" & restoreNote
    If Not modAppState.gAutomationActive Then
        modAppState.ReportFailure IIf(IsAdd, "Add ", "Delete ") & KindLabel(Kind), reason, restoreNote
    End If
End Sub

' Cleanup, but only when a snapshot genuinely exists. If CaptureAppState itself
' failed there is nothing to restore, and saying otherwise would be a false claim.
Private Function FinishIfCaptured(ByRef Snapshot As AppStateSnapshot, _
                                  ByVal StateCaptured As Boolean) As String
    If Not StateCaptured Then
        FinishIfCaptured = "  application state was never captured, so nothing could be " & _
                           "restored. Check Excel's calculation mode, alerts, events and " & _
                           "screen updating before continuing." & vbCrLf
        Exit Function
    End If
    FinishIfCaptured = modAppState.FinishOperation(Snapshot)
End Function

' Restores exactly the blocks an Add or Delete may modify, and REPORTS the outcome.
' The register and profiling snapshots carry row count as well as column count, so a
' failed Add that had already grown the table cannot leave an extra row behind and a
' failed Delete cannot leave a missing one.
'
' The counter is restored ONLY here, because this operation failed and therefore
' issued no surviving identifier. A successful Delete never restores it.
Private Function TryRestoreDriver(ByVal Kind As String, _
                                  ByRef RegisterBefore As TableSnapshot, _
                                  ByRef ProfilingBefore As TableSnapshot, _
                                  ByVal CounterBefore As Variant) As String
    On Error GoTo RestoreFailed

    modWorkbook.RestoreTable RegisterTable(Kind), RegisterBefore
    modWorkbook.RestoreTable modProfiling.ProfilingTable(Kind), ProfilingBefore
    ' Exactly as it was: Empty stays Empty, text stays that text, a number stays that
    ' number. Restoring anything else would destroy the corruption the operation
    ' refused to act on, and re-enable identifier reuse.
    modWorkbook.WriteValue CounterName(Kind), CounterBefore

    TryRestoreDriver = "The register, its profiling grid and the ID counter have been " & _
                       "restored to their state from before this operation, including " & _
                       "row and column counts. No identifier issued by the failed " & _
                       "operation survives, and no partial change remains."
    Exit Function

RestoreFailed:
    TryRestoreDriver = "RESTORE INCOMPLETE. Error " & Err.Number & ": " & Err.Description & _
                       vbCrLf & "The workbook may hold a partial structural change. " & _
                       "Do not continue; close without saving and reopen the last saved copy."
End Function
