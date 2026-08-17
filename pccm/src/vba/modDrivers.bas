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
'   * Pad width is a minimum display width. CL-999 is followed by CL-1000; there
'     is no artificial ID maximum anywhere.
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

Public Function ReadCounter(ByVal Kind As String) As Long
    ReadCounter = modWorkbook.ReadLong(CounterName(Kind), 0)
End Function

' Advances the counter and returns the identifier it just issued. The counter is
' persisted immediately so a later failure cannot hand the same number out twice.
Public Function AllocateId(ByVal Kind As String) As String
    Dim nextSequence As Long
    nextSequence = ReadCounter(Kind) + 1
    modWorkbook.WriteValue CounterName(Kind), nextSequence
    AllocateId = FormatId(Kind, nextSequence)
End Function

Public Function HighestIssued(ByVal Kind As String) As Long
    Dim register As ListObject
    Dim r As Long, rowCount As Long, idCol As Long, best As Long
    Dim prefix As String, idText As String, tail As String

    Set register = RegisterTable(Kind)
    idCol = IdColumn(Kind)
    rowCount = modWorkbook.BodyRowCount(register)
    prefix = IdPrefix(Kind)

    For r = 1 To rowCount
        idText = modWorkbook.TextOf(modWorkbook.CellIn(register, r, idCol))
        If Len(idText) > Len(prefix) Then
            If StrComp(Left$(idText, Len(prefix)), prefix, vbTextCompare) = 0 Then
                tail = Mid$(idText, Len(prefix) + 1)
                If IsNumeric(tail) Then
                    If CLng(tail) > best Then best = CLng(tail)
                End If
            End If
        End If
    Next r
    HighestIssued = best
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
            Else
                FirstFreeRow = r
                Exit Function
            End If
        End If
    Next r
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

    If targetRow = 0 And orphanRow > 0 Then
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

    ' Put the cursor on the first field the user actually owns.
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
    RunDriverOperation KIND_COST, False, SelectedId(KIND_COST)
End Sub

Public Sub PCCM_DeleteRisk()
    RunDriverOperation KIND_RISK, False, SelectedId(KIND_RISK)
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
    Dim registerBefore As Variant, profilingBefore As Variant
    Dim counterBefore As Long
    Dim problems As String
    Dim captured As Boolean

    snapshot = modAppState.CaptureAppState()
    On Error GoTo Failure
    modAppState.BeginOperation

    registerBefore = modWorkbook.SnapshotTable(RegisterTable(Kind))
    profilingBefore = modWorkbook.SnapshotTable(modProfiling.ProfilingTable(Kind))
    counterBefore = ReadCounter(Kind)
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

    modAppState.RecalculateStructuralState
    modAppState.RestoreAppState snapshot
    modAppState.Announce result
    Exit Sub

Failure:
    Dim reason As String
    reason = "Error " & Err.Number & ": " & Err.Description
    If captured Then
        On Error Resume Next
        modWorkbook.RestoreTable RegisterTable(Kind), registerBefore
        modWorkbook.RestoreTable modProfiling.ProfilingTable(Kind), profilingBefore
        ' The counter is restored ONLY because this operation failed and issued no
        ' surviving identifier. A successful delete never restores it.
        modWorkbook.WriteValue CounterName(Kind), counterBefore
        On Error GoTo 0
    End If
    modAppState.RecalculateStructuralState
    modAppState.RestoreAppState snapshot
    modAppState.RecordResult "FAIL|" & reason
    If Not modAppState.gAutomationActive Then
        modAppState.ReportFailure IIf(IsAdd, "Add ", "Delete ") & KindLabel(Kind), reason, _
            "The register, its profiling grid and the ID counter have been restored to " & _
            "their values from before this operation. No partial change remains."
    End If
End Sub
