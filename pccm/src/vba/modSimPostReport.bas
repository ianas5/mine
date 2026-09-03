Attribute VB_Name = "modSimPostReport"
Option Explicit

' ==========================================================================
' PCCM Phase 7 - sensitivity orchestration, persistence and materialisation.
'
' It runs the pipeline and owns none of the mathematics in it:
'
'   modSimReport      the simulation status and the published identity
'   modCalcReport     the ONE Phase-5 bridge that resolves DriverFactors
'   modSimEngine      per-driver replay and the shared contribution routine
'   modSimSensitivity mid-ranks, Spearman, the undefined case and the ordering
'
' No RNG, no sampler, no contribution arithmetic, no rank arithmetic, no
' fingerprint and no digest is computed here. This module decides WHEN the
' analysis may run, in WHAT order, and WHERE the answer is written.
'
' ----------------------------------------------------------------------------
' IT IS NOT PART OF A SIMULATION RUN
' ----------------------------------------------------------------------------
' PCCM_RunSensitivity is a separate endpoint on purpose. A successful
' simulation must stay successful even if the analysis of it later fails; at
' full N the analysis can cost about as much as the run it explains; and the
' user must be able to repeat it without producing a new stochastic result. So
' it allocates no run id, advances no AUTO nonce, touches no pending-nonce
' marker, writes no attempt row, and leaves the iteration records and the
' result digest exactly as it found them.
'
' ----------------------------------------------------------------------------
' WHAT IT REFUSES, AND WHY REFUSING IS THE POINT
' ----------------------------------------------------------------------------
' Sensitivity explains ONE run. Replaying the CURRENT model against a total
' vector published by a DIFFERENT model would produce a table that is wrong in
' a way nothing on the sheet could reveal, so the analysis runs only while the
' simulation status is CURRENT and refuses otherwise. A refusal is an attempt
' outcome, not a state: it changes no simulation state and destroys no
' sensitivity that belongs to an earlier successful run.
'
' ----------------------------------------------------------------------------
' MEMORY
' ----------------------------------------------------------------------------
' O(N), never O(D x N). The total's ranks are computed ONCE and reused; one
' driver's contribution vector exists at a time and is replaced on the next
' driver; what accumulates is one result record per driver. The D x N matrix
' this architecture refuses is never built, in memory or on a sheet.
' ==========================================================================

Private Const SENSITIVITY_NO_VARIANCE_LABEL As String = "n/a - no variance"
Private Const SENSITIVITY_RANKED_LABEL As String = "ranked"
Private Const SENSITIVITY_DIRECTION_UP As String = "+"
Private Const SENSITIVITY_DIRECTION_DOWN As String = "-"
Private Const SENSITIVITY_COST_TYPE As String = "Cost Line"
Private Const SENSITIVITY_RISK_TYPE As String = "Risk"

' The published identity of the run being explained, read once and carried.
Private Type SensitivityRun
    Bank As String
    RunId As Long
    EffectiveSeed As Long
    RequestFingerprint As String
    ResultDigest As String
    Iterations As Long
End Type

' ==========================================================================
' THE ENDPOINT
' ==========================================================================
Public Sub PCCM_RunSensitivity()
    ' THE SAME ENVELOPE THE ACCEPTED ENDPOINTS USE, and for the same reason: a
    ' failure must not leave ScreenUpdating off or Calculation manual. There is
    ' no `committed` distinction here - sensitivity commits nothing that a
    ' cleanup problem could strand, because it changes no simulation state.
    Dim state As AppStateSnapshot, result As OperationResult
    Dim stateCaptured As Boolean, cleanupAttempted As Boolean
    Dim cleanup As String, failure As String

    On Error GoTo InvocationFailed
    state = modAppState.CaptureAppState()
    stateCaptured = True
    modAppState.BeginOperation
    result = RunSensitivity()
    On Error GoTo 0

    On Error GoTo NormalCleanupFailed
    cleanupAttempted = True
    cleanup = modAppState.FinishOperation(state)
    On Error GoTo 0
    stateCaptured = False

    If Len(cleanup) > 0 Then
        result = modAppState.Failed("Run Sensitivity", _
            "Application state could not be restored: " & cleanup)
    End If
    modAppState.Announce result
    Exit Sub

NormalCleanupFailed:
    failure = Err.Description
    On Error GoTo 0
    modAppState.Announce modAppState.Failed("Run Sensitivity", _
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
    modAppState.Announce modAppState.Failed("Run Sensitivity", failure)
    Exit Sub

CleanupFailed:
    On Error GoTo 0
    modAppState.Announce modAppState.Failed("Run Sensitivity", failure & vbCrLf & _
        "Application state could not be restored after the failure.")
End Sub

' ==========================================================================
' THE PIPELINE
'
' Every step must succeed before ANYTHING is written. The whole result is
' built in memory, the block is cleared, the records and identity are written,
' and only then is the block marked published - so a failure at driver 287
' cannot leave rows 1 to 286 looking like a complete current answer.
' ==========================================================================
Private Function RunSensitivity() As OperationResult
    Dim run As SensitivityRun
    Dim drivers() As DriverFactors
    Dim results() As SimSensitivityResult
    Dim order() As Long
    Dim totals() As Double, totalRanks() As Double
    Dim driverCount As Long, eligibleCount As Long
    Dim detail As String

    If Not RequireCurrentRun(run, detail) Then
        RunSensitivity = Refused(detail)
        Exit Function
    End If
    If Not ResolveDrivers(drivers, driverCount, detail) Then
        RunSensitivity = Refused(detail)
        Exit Function
    End If
    If Not ReadTotals(run, totals, detail) Then
        RunSensitivity = Refused(detail)
        Exit Function
    End If

    ' THE TOTAL'S RANKS, ONCE. Ranking it again per driver would be D sorts of
    ' one vector for one answer.
    If Not modSimSensitivity.SimSensitivityMidRanks(totals, run.Iterations, _
                                                    totalRanks, detail) Then
        RunSensitivity = Refused(detail)
        Exit Function
    End If

    If Not AnalyseDrivers(run, drivers, driverCount, totalRanks, results, detail) Then
        RunSensitivity = Refused(detail)
        Exit Function
    End If
    If Not modSimSensitivity.SimSensitivityRank(results, driverCount, order, _
                                                eligibleCount, detail) Then
        RunSensitivity = Refused(detail)
        Exit Function
    End If

    ' NOTHING HAS BEEN WRITTEN UNTIL HERE.
    If Not Publish(run, results, driverCount, order, eligibleCount, detail) Then
        RunSensitivity = Refused(detail)
        Exit Function
    End If

    RunSensitivity.Ok = True
    RunSensitivity.Message = "Sensitivity complete: " & CStr(eligibleCount) & _
                             " ranked of " & CStr(driverCount) & " drivers, over " & _
                             CStr(run.Iterations) & " iterations."
End Function

' ==========================================================================
' THE PRECONDITION
'
' CURRENT, and bound to the published identity. The status is derived by
' modSimReport and is not re-derived here: a second derivation of run state is
' a second answer to the only question that matters before replay.
' ==========================================================================
Private Function RequireCurrentRun(ByRef run As SensitivityRun, _
                                   ByRef detail As String) As Boolean
    Dim status As String

    status = modSimReport.PCCM_SimulationStatus()
    If Len(status) = 0 Then
        detail = "sensitivity: no successful simulation has been published, so " & _
                 "there is nothing to analyse"
        Exit Function
    End If
    If StrComp(status, SIM_STATE_CURRENT, vbBinaryCompare) <> 0 Then
        detail = "sensitivity: the simulation is " & status & ". Sensitivity " & _
                 "explains one run, and replaying the current model against a " & _
                 "total published by a different one would be wrong in a way the " & _
                 "sheet could not show. Run the simulation again first."
        Exit Function
    End If

    If Not ReadRunIdentity(run, detail) Then Exit Function
    RequireCurrentRun = True
End Function

Private Function ReadRunIdentity(ByRef run As SensitivityRun, _
                                 ByRef detail As String) As Boolean
    Dim bank As String

    bank = SharedText(SIM_IDENTITY_ROW_ACTIVE_BANK)
    If StrComp(bank, SIM_BANK_A, vbBinaryCompare) <> 0 And _
       StrComp(bank, SIM_BANK_B, vbBinaryCompare) <> 0 Then
        detail = "sensitivity: the active publication bank is not readable"
        Exit Function
    End If
    run.Bank = bank
    run.RequestFingerprint = modSimReport.PCCM_SimulationRequestFingerprint()
    run.ResultDigest = modSimReport.PCCM_SimulationResultDigest()
    If Len(run.RequestFingerprint) = 0 Or Len(run.ResultDigest) = 0 Then
        detail = "sensitivity: the published run carries no identity to bind to"
        Exit Function
    End If
    If Not ReadSnapshotLong(bank, SIM_IDENTITY_ROW_RUN_ID, run.RunId, detail) Then Exit Function
    If Not ReadSnapshotLong(bank, SIM_IDENTITY_ROW_EFFECTIVE_SEED, _
                            run.EffectiveSeed, detail) Then Exit Function
    If Not ReadSnapshotLong(bank, SIM_IDENTITY_ROW_ITERATIONS_RUN, _
                            run.Iterations, detail) Then Exit Function
    If run.Iterations < 1 Then
        detail = "sensitivity: the published run records no iterations"
        Exit Function
    End If
    ReadRunIdentity = True
End Function

' ==========================================================================
' THE RESOLVED MODEL - through the ONE accepted Phase-5 bridge
' ==========================================================================
Private Function ResolveDrivers(ByRef drivers() As DriverFactors, _
                                ByRef driverCount As Long, _
                                ByRef detail As String) As Boolean
    Dim fingerprint As String, timeline As String, separator As String
    Dim baseNominal As Double, basePv As Double

    If Not modCalcReport.CalcPrepareSimulationInputs( _
            drivers, driverCount, fingerprint, baseNominal, basePv, _
            timeline, separator, detail) Then
        Exit Function
    End If
    If driverCount < 1 Then
        detail = "sensitivity: the model resolves no drivers to analyse"
        Exit Function
    End If
    ResolveDrivers = True
End Function

' ==========================================================================
' THE PERSISTED TOTALS, in original iteration order
' ==========================================================================
Private Function ReadTotals(ByRef run As SensitivityRun, ByRef totals() As Double, _
                            ByRef detail As String) As Boolean
    Dim block As Variant
    Dim index As Long
    Dim address As String
    Dim value As Double

    address = IterationTotalColumn(run.Bank) & CStr(SIM_DATA_FIRST_ITERATION_ROW) & ":" & _
              IterationTotalColumn(run.Bank) & _
              CStr(SIM_DATA_FIRST_ITERATION_ROW + run.Iterations - 1)
    block = SimSheet().Range(address).Value2
    ReDim totals(0 To run.Iterations - 1)
    For index = 1 To run.Iterations
        If Not modWorkbook.TryReadDouble(block(index, 1), value) Then
            detail = "sensitivity: iteration " & CStr(index) & " of the published " & _
                     "total is not a usable number"
            Exit Function
        End If
        totals(index - 1) = value
    Next index
    ReadTotals = True
End Function

' ==========================================================================
' ONE DRIVER AT A TIME
'
' `contributions` is REPLACED on each pass, never accumulated. That is the
' whole memory design: the matrix is not built, so it cannot be retained.
' ==========================================================================
Private Function AnalyseDrivers(ByRef run As SensitivityRun, _
                                ByRef drivers() As DriverFactors, _
                                ByVal driverCount As Long, _
                                ByRef totalRanks() As Double, _
                                ByRef results() As SimSensitivityResult, _
                                ByRef detail As String) As Boolean
    Dim contributions() As Double
    Dim index As Long
    Dim rho As Double
    Dim status As Long

    ReDim results(0 To driverCount - 1)
    For index = 0 To driverCount - 1
        If Not modSimEngine.SimEngineReplayDriver( _
                drivers, driverCount, run.EffectiveSeed, run.Iterations, _
                drivers(LBound(drivers) + index).PermanentId, contributions, detail) Then
            Exit Function
        End If
        If Not modSimSensitivity.SimSensitivitySpearman( _
                contributions, totalRanks, run.Iterations, rho, status, detail) Then
            Exit Function
        End If
        results(index).PermanentId = drivers(LBound(drivers) + index).PermanentId
        results(index).Rho = rho
        results(index).AbsRho = Abs(rho)
        results(index).Status = status
    Next index
    AnalyseDrivers = True
End Function

' ==========================================================================
' PUBLICATION
'
' Clear, write the records, write the identity, and mark it published LAST. A
' block that fails part way through carries no published marker, so nothing
' partial can be read as a current answer.
' ==========================================================================
Private Function Publish(ByRef run As SensitivityRun, _
                         ByRef results() As SimSensitivityResult, _
                         ByVal driverCount As Long, ByRef order() As Long, _
                         ByVal eligibleCount As Long, ByRef detail As String) As Boolean
    Dim block As Variant
    Dim index As Long, slot As Long, position As Long
    Dim first As String, last As String

    first = SensitivityFirstColumn(run.Bank)
    last = SensitivityLastColumn(run.Bank)

    ' THE MARKER GOES FIRST, and it goes BLANK. From here until the last write
    ' the block is explicitly not published, so an interruption anywhere in
    ' between leaves it unreadable rather than half-true.
    StampCell(run.Bank, SIM_SENSITIVITY_STAMP_ROW_PUBLISHED).Value2 = vbNullString

    ' SURPLUS ROWS GO TOO. A later model can have fewer drivers than this bank
    ' already holds, and overwriting only the first n would leave the remainder
    ' visible and indistinguishable from the new result.
    ClearRecords run.Bank, first, last

    ReDim block(1 To driverCount, 1 To SIM_SENSITIVITY_FIELD_COUNT)
    ' RANKED FIRST, in the order the kernel returned, then everything that has
    ' no rho to rank - still reported, never ranked.
    slot = 0
    For position = 0 To eligibleCount - 1
        index = order(LBound(order) + position)
        FillRecord block, slot + 1, results(index), position + 1
        slot = slot + 1
    Next position
    For index = 0 To driverCount - 1
        If results(index).Status <> SIM_SENSITIVITY_DEFINED Then
            FillRecord block, slot + 1, results(index), 0
            slot = slot + 1
        End If
    Next index
    If slot <> driverCount Then
        detail = "sensitivity: " & CStr(slot) & " records were built for " & _
                 CStr(driverCount) & " drivers"
        Exit Function
    End If

    SimSheet().Range(first & CStr(SIM_SENSITIVITY_FIRST_ROW) & ":" & last & _
                     CStr(SIM_SENSITIVITY_FIRST_ROW + driverCount - 1)).Value2 = block

    ' THE IDENTITY THIS TABLE BELONGS TO. Without it a table produced for one
    ' run could be read as the answer for another.
    StampCell(run.Bank, SIM_SENSITIVITY_STAMP_ROW_RUN_ID).Value2 = run.RunId
    StampCell(run.Bank, SIM_SENSITIVITY_STAMP_ROW_EFFECTIVE_SEED).Value2 = run.EffectiveSeed
    StampCell(run.Bank, SIM_SENSITIVITY_STAMP_ROW_REQUEST_FINGERPRINT).Value2 = _
        run.RequestFingerprint
    StampCell(run.Bank, SIM_SENSITIVITY_STAMP_ROW_RESULT_DIGEST).Value2 = run.ResultDigest
    StampCell(run.Bank, SIM_SENSITIVITY_STAMP_ROW_ITERATIONS).Value2 = run.Iterations
    StampCell(run.Bank, SIM_SENSITIVITY_STAMP_ROW_RECORD_COUNT).Value2 = driverCount

    ' AND ONLY NOW.
    StampCell(run.Bank, SIM_SENSITIVITY_STAMP_ROW_PUBLISHED).Value2 = SIM_SENSITIVITY_PUBLISHED
    Publish = True
End Function

Private Sub FillRecord(ByRef block As Variant, ByVal row As Long, _
                       ByRef record As SimSensitivityResult, ByVal rank As Long)
    block(row, SIM_SENSITIVITY_OFFSET_DRIVER_ID + 1) = record.PermanentId
    block(row, SIM_SENSITIVITY_OFFSET_DRIVER_TYPE + 1) = DriverTypeOf(record.PermanentId)
    block(row, SIM_SENSITIVITY_OFFSET_DRIVER_NAME + 1) = DriverNameOf(record.PermanentId)
    If record.Status = SIM_SENSITIVITY_DEFINED Then
        block(row, SIM_SENSITIVITY_OFFSET_RHO + 1) = record.Rho
        block(row, SIM_SENSITIVITY_OFFSET_ABS_RHO + 1) = record.AbsRho
        block(row, SIM_SENSITIVITY_OFFSET_RANK + 1) = rank
        block(row, SIM_SENSITIVITY_OFFSET_DIRECTION + 1) = DirectionOf(record.Rho)
        block(row, SIM_SENSITIVITY_OFFSET_STATUS + 1) = SENSITIVITY_RANKED_LABEL
    Else
        ' NO RHO, NO RANK, NO DIRECTION. A constant column offers no monotone
        ' relationship to find, and printing 0 would say one was looked for.
        block(row, SIM_SENSITIVITY_OFFSET_RHO + 1) = vbNullString
        block(row, SIM_SENSITIVITY_OFFSET_ABS_RHO + 1) = vbNullString
        block(row, SIM_SENSITIVITY_OFFSET_RANK + 1) = vbNullString
        block(row, SIM_SENSITIVITY_OFFSET_DIRECTION + 1) = vbNullString
        block(row, SIM_SENSITIVITY_OFFSET_STATUS + 1) = SENSITIVITY_NO_VARIANCE_LABEL
    End If
End Sub

Private Function DirectionOf(ByVal rho As Double) As String
    If rho < 0# Then
        DirectionOf = SENSITIVITY_DIRECTION_DOWN
    Else
        DirectionOf = SENSITIVITY_DIRECTION_UP
    End If
End Function

Private Sub ClearRecords(ByVal bank As String, ByVal first As String, ByVal last As String)
    SimSheet().Range(first & CStr(SIM_SENSITIVITY_FIRST_ROW) & ":" & last & _
                     CStr(SIM_SENSITIVITY_FIRST_ROW + SIM_MAX_ITERATIONS - 1)).ClearContents
End Sub

' ==========================================================================
' Reading helpers - no state derivation, no arithmetic
' ==========================================================================
Private Function DriverTypeOf(ByVal permanentId As String) As String
    If modDrivers.RowOfId(modProfiling.RiskKind(), permanentId) > 0 Then
        DriverTypeOf = SENSITIVITY_RISK_TYPE
    Else
        DriverTypeOf = SENSITIVITY_COST_TYPE
    End If
End Function

Private Function DriverNameOf(ByVal permanentId As String) As String
    Dim kind As String, table As ListObject
    Dim row As Long, column As Long

    kind = modProfiling.RiskKind()
    row = modDrivers.RowOfId(kind, permanentId)
    column = COL_RISK_REGISTER_DESCRIPTION
    If row < 1 Then
        kind = modProfiling.CostKind()
        row = modDrivers.RowOfId(kind, permanentId)
        column = COL_COST_LINES_DESCRIPTION
    End If
    If row < 1 Then Exit Function
    Set table = modDrivers.RegisterTable(kind)
    DriverNameOf = modWorkbook.TextOf(modWorkbook.CellIn(table, row, column))
End Function

Private Function ReadSnapshotLong(ByVal bank As String, ByVal row As Long, _
                                  ByRef value As Long, ByRef detail As String) As Boolean
    Dim raw As Variant, measured As Double

    raw = SimSheet().Range(SnapshotColumn(bank) & CStr(row)).Value2
    If Not modWorkbook.IsWholeInRange(raw, 0#, CDbl(SIM_MAX_ITERATIONS), measured) Then
        detail = "sensitivity: the published run identity at row " & CStr(row) & _
                 " is not a whole number"
        Exit Function
    End If
    value = modWorkbook.SafeLong(measured)
    ReadSnapshotLong = True
End Function

Private Function Refused(ByVal detail As String) As OperationResult
    Refused.Ok = False
    Refused.Message = "Sensitivity was not produced."
    Refused.Detail = detail
End Function

Private Function SimSheet() As Worksheet
    Set SimSheet = modWorkbook.Sh(SIM_DATA_SHEET)
End Function

Private Function SharedText(ByVal row As Long) As String
    SharedText = modWorkbook.TextOf(SimSheet().Range(SIM_SHARED_VALUE_COLUMN & CStr(row)))
End Function

Private Function SnapshotColumn(ByVal bank As String) As String
    If StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Then
        SnapshotColumn = SIM_SNAPSHOT_COLUMN_A
    Else
        SnapshotColumn = SIM_SNAPSHOT_COLUMN_B
    End If
End Function

Private Function IterationTotalColumn(ByVal bank As String) As String
    If StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Then
        IterationTotalColumn = SIM_ITER_A_TOTAL_NOMINAL_COLUMN
    Else
        IterationTotalColumn = SIM_ITER_B_TOTAL_NOMINAL_COLUMN
    End If
End Function

Private Function SensitivityFirstColumn(ByVal bank As String) As String
    If StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Then
        SensitivityFirstColumn = SIM_SENSITIVITY_A_FIRST_COLUMN
    Else
        SensitivityFirstColumn = SIM_SENSITIVITY_B_FIRST_COLUMN
    End If
End Function

Private Function SensitivityLastColumn(ByVal bank As String) As String
    If StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Then
        SensitivityLastColumn = SIM_SENSITIVITY_A_LAST_COLUMN
    Else
        SensitivityLastColumn = SIM_SENSITIVITY_B_LAST_COLUMN
    End If
End Function

Private Function StampCell(ByVal bank As String, ByVal row As Long) As Range
    Dim column As String
    If StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Then
        column = SIM_SENSITIVITY_STAMP_COLUMN_A
    Else
        column = SIM_SENSITIVITY_STAMP_COLUMN_B
    End If
    Set StampCell = SimSheet().Range(column & CStr(row))
End Function
