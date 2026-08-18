Attribute VB_Name = "modCalcReport"
Option Explicit

' ==========================================================================
' modCalcReport - the Phase-5 presentation and orchestration layer.
'
' It runs the whole calculation IN MEMORY, and only then writes. The workbook
' is the RECORD of a calculation, never an input back into one: no analytical
' value is ever read back out of _Calc to produce another analytical value.
'
' --------------------------------------------------------------------------
' ONE PREPARATION PATH, THREE CONSUMERS
' --------------------------------------------------------------------------
' PCCM_Calculate, PCCM_CalculationStatus and PCCM_CurrentInputFingerprint all
' go through PrepareCurrentCalculation. There is exactly one definition of
' "valid current inputs", so a state the write path would refuse can never be
' reported CURRENT because a partial digest happened to be constructible.
'
' --------------------------------------------------------------------------
' TWO AXES THAT NEVER MIX
' --------------------------------------------------------------------------
'   DERIVED STATUS   what do CURRENT inputs say about the stored successful
'                    snapshot?  NOT CALCULATED / CURRENT / STALE / INVALID
'   ATTEMPT RESULT   what happened the last time PCCM_Calculate was explicitly
'                    attempted?  NONE / SUCCESS / REFUSED / FAILED
'
' REFUSED and FAILED are never statuses, and the historical attempt never
' decides the status. A refusal last week does not make today's matching
' fingerprint stale.
'
' --------------------------------------------------------------------------
' CONTROLLED REFUSAL VERSUS INTERNAL FAILURE
' --------------------------------------------------------------------------
'   REFUSED  the accepted machinery declined BEFORE any analytical write - a
'            structural prerequisite, a numerical prerequisite, a range
'            refusal, an identity that did not hold, a digest that could not be
'            built. Nothing analytical is touched.
'   FAILED   something went wrong DURING the write. The previous successful
'            snapshot is restored and the attempt is recorded.
'
' An unexpected error is never quietly downgraded to a refusal to keep going.
'
' NOTHING HERE HAS BEEN EXECUTED. It is source, submitted for review.
' ==========================================================================

' The audit vocabulary for the Driver Kind column. The calculation contract
' states it as the column's units, "Cost Line / Risk"; it is a display label,
' not a coordinate, and every anchor and schema below comes from
' modCalcContract.
Private Const DRIVER_KIND_COST As String = "Cost Line"
Private Const DRIVER_KIND_RISK As String = "Risk"

' Failpoint stages, for the Phase-4 injection mechanism the Gate-B harness
' already knows how to arm. Two, deliberately: one where analytical state is
' half-written, and one around the commit itself.
Public Const FAILPOINT_ANALYTICAL_WRITE As String = "Phase5AnalyticalWrite"
Public Const FAILPOINT_SUCCESS_COMMIT As String = "Phase5SuccessCommit"

' Everything the calculation produces, before anything is written.
Private Type CalculationPackage
    Model As ResolvedModel
    Drivers() As DriverFactors
    Audits() As DriverAudit
    Years() As YearFactors
    Totals As AnalyticalTotals
    Annual() As AnnualRow
    Magnitudes As ReconciliationMagnitudes
    Checks() As IdentityCheck
    ' InflationFactors(profileIndex, offset) over BaseYear .. LastYear.
    InflationFactors() As Double
    InflationSpan As Long
    ' ReferencedBy(currencyIndex): how many resolved drivers use that currency.
    ReferencedBy() As Long
    Fingerprint As String
End Type

' The five analytical tables plus the two scalar blocks, as they were before
' the transaction began.
Private Type CalculationSnapshot
    Years As TableSnapshot
    Inflation As TableSnapshot
    Fx As TableSnapshot
    Drivers As TableSnapshot
    Annual As TableSnapshot
    Totals As Variant
    State As Variant
    Captured As Boolean
End Type

' ==========================================================================
' THE SIX PHASE-5 ENDPOINTS
' ==========================================================================
Public Sub PCCM_Calculate()
    ' The only endpoint that writes. Application state is captured and restored
    ' through the accepted Phase-4 discipline, whatever the outcome.
    Dim state As AppStateSnapshot, result As OperationResult
    Dim cleanup As String
    state = modAppState.CaptureAppState()
    modAppState.BeginOperation
    result = RunCalculation()
    cleanup = modAppState.FinishOperation(state)
    If Len(cleanup) > 0 Then
        result = modAppState.Failed("Calculate", cleanup)
    End If
    modAppState.ReportResult result
End Sub

Public Function PCCM_CalculationStatus() As String
    ' Re-evaluates the status and writes ONLY C19:C20. It touches no analytical
    ' block and no part of the last-success record: status is last-evaluated,
    ' not live, and asking for it is not a calculation.
    Dim package As CalculationPackage, detail As String
    Dim prepared As Boolean, status As String
    prepared = PrepareCurrentCalculation(package, detail)
    status = DeriveStatus(package, prepared)
    WriteStatusBlock status
    PCCM_CalculationStatus = status
End Function

Public Function PCCM_CalculationFingerprint() As String
    ' The STORED digest of the last successful calculation, read as it stands.
    ' It is never recomputed: recomputing would answer a different question,
    ' and the empty string is the honest answer when no success exists.
    PCCM_CalculationFingerprint = StoredText(CALC_STATE_ROW_LAST_SUCCESSFUL_FINGERPRINT)
End Function

Public Function PCCM_CurrentInputFingerprint() As String
    ' The digest of the CURRENT inputs, through the same preparation path the
    ' write uses. Empty when the current inputs would be refused - not a
    ' sentinel digest, because a sentinel would eventually be compared.
    Dim package As CalculationPackage, detail As String
    If PrepareCurrentCalculation(package, detail) Then
        PCCM_CurrentInputFingerprint = package.Fingerprint
    End If
End Function

Public Function PCCM_CalculationAttemptResult() As String
    PCCM_CalculationAttemptResult = StoredText(CALC_STATE_ROW_LAST_ATTEMPT_RESULT)
End Function

Public Function PCCM_CalculationAttemptDetail() As String
    PCCM_CalculationAttemptDetail = StoredText(CALC_STATE_ROW_LAST_ATTEMPT_DETAIL)
End Function

' ==========================================================================
' The transaction
' ==========================================================================
Private Function RunCalculation() As OperationResult
    Dim package As CalculationPackage, snapshot As CalculationSnapshot
    Dim detail As String, committed As Boolean

    ' EVERYTHING IN MEMORY FIRST. A refusal here has touched no analytical
    ' block, so the previous successful snapshot stands untouched and only the
    ' attempt/status metadata moves.
    If Not PrepareCurrentCalculation(package, detail) Then
        RecordRefusal detail
        RunCalculation = modAppState.Failed("Calculate", detail)
        Exit Function
    End If

    CaptureSnapshot snapshot

    On Error GoTo TransactionFailed
    WriteAnalytical package
    modAppState.FailPointCheck FAILPOINT_ANALYTICAL_WRITE
    If Not VerifyAnalytical(package) Then
        Err.Raise vbObjectError + 5101, "modCalcReport.RunCalculation", _
                  "the analytical snapshot did not verify against the prepared result"
    End If
    modAppState.FailPointCheck FAILPOINT_SUCCESS_COMMIT
    ' THE COMMIT IS INSIDE THE TRANSACTION. The assignment can fail and its
    ' verification can fail, so both sit inside the rollback envelope; SUCCESS
    ' is published only once the analytical snapshot has been verified.
    WriteSuccessCommit package
    If Not VerifySuccessCommit(package) Then
        Err.Raise vbObjectError + 5102, "modCalcReport.RunCalculation", _
                  "the success commit did not verify"
    End If
    committed = True
    On Error GoTo 0
    RunCalculation = modAppState.Succeeded("Calculation committed.")
    Exit Function

TransactionFailed:
    detail = Err.Description
    On Error GoTo 0
    If committed Then
        ' Unreachable by construction, and stated anyway: once committed, no
        ' later mutation may turn the calculation into a failure.
        RunCalculation = modAppState.Succeeded("Calculation committed.")
        Exit Function
    End If
    ' ROLLBACK FIRST, metadata second. The first observable moment after a
    ' failure must be the previous successful snapshot, exactly.
    RestoreSnapshot snapshot
    RecordFailure detail
    RunCalculation = modAppState.Failed("Calculate", detail)
End Function

' ==========================================================================
' Preparation - the ONE definition of a valid current calculation
' ==========================================================================
Private Function PrepareCurrentCalculation(ByRef package As CalculationPackage, _
                                           ByRef detail As String) As Boolean
    Dim blank As CalculationPackage
    package = blank
    detail = vbNullString
    If Not modCalcResolve.ResolveModel(package.Model, detail) Then Exit Function
    If Not modCalcCheck.CheckResolvedModel(package.Model, detail) Then Exit Function
    If Not BuildFactorTables(package, detail) Then Exit Function
    If Not BuildDriverFactors(package, detail) Then Exit Function
    If Not BuildAudits(package, detail) Then Exit Function
    If Not modCalcAnalytical.AccumulateTotals(package.Audits, package.Model.DriverCount, _
                                              package.Totals, package.Magnitudes, _
                                              detail) Then Exit Function
    If Not BuildAnnual(package, detail) Then Exit Function
    If Not modCalcAnalytical.Reconcile(package.Totals, package.Annual, package.Drivers, _
                                       package.Model.DriverCount, package.Model.Weights, _
                                       package.Magnitudes, package.Checks, detail) Then
        Exit Function
    End If
    If Not modCalcAnalytical.AllIdentitiesHold(package.Checks) Then
        detail = "reconciliation: at least one identity does not hold within its allowance"
        Exit Function
    End If
    CountCurrencyReferences package
    If Not BuildFingerprint(package, detail) Then Exit Function
    If Len(package.Fingerprint) = 0 Then
        detail = "the calculation fingerprint could not be constructed"
        Exit Function
    End If
    PrepareCurrentCalculation = True
End Function

Private Function BuildFactorTables(ByRef package As CalculationPackage, _
                                   ByRef detail As String) As Boolean
    ' The inflation factor for every referenced profile over the whole audited
    ' span, and the discount factor for every applied project year. Both come
    ' from the accepted builders; nothing is compounded here.
    Dim profile As Long, offset As Long, index As Long
    Dim series() As Double, discount() As Double

    package.InflationSpan = package.Model.Timeline.LastYear - _
                            package.Model.Timeline.BaseYear + 1
    If package.Model.ProfileCount > 0 Then
        ReDim package.InflationFactors(0 To package.Model.ProfileCount - 1, _
                                       0 To package.InflationSpan - 1)
        For profile = 0 To package.Model.ProfileCount - 1
            If Not modCalcResolve.ResolveInflationFactors( _
                    package.Model.InflationRates, profile, _
                    package.Model.RequiredYearCount, package.Model.Timeline, _
                    series, detail) Then Exit Function
            For offset = 0 To package.InflationSpan - 1
                package.InflationFactors(profile, offset) = series(offset)
            Next offset
        Next profile
    End If

    If Not modCalcFactors.BuildDiscountFactors(package.Model.Timeline.DiscountRate, _
                                               package.Model.Timeline.Duration, _
                                               discount, detail) Then
        detail = "discount factors: " & detail
        Exit Function
    End If
    ReDim package.Years(0 To package.Model.Timeline.Duration - 1)
    For index = 0 To package.Model.Timeline.Duration - 1
        package.Years(index).ProjectIndex = package.Model.ProjectIndexes(index)
        package.Years(index).CalendarYear = package.Model.CalendarYears(index)
        package.Years(index).DiscountF = discount(index)
    Next index
    BuildFactorTables = True
End Function

Private Function BuildDriverFactors(ByRef package As CalculationPackage, _
                                    ByRef detail As String) As Boolean
    ' One DriverFactors per resolved driver, with Knom and Kpv from the accepted
    ' builders.
    '
    ' THE CARRY CONVENTION LIVES HERE AND NOWHERE ELSE. A cost line carries its
    ' Quantity and a Probability of 1; a risk carries a Quantity of 1 and its
    ' Probability. Those identities are in-memory calculation semantics: neither
    ' is a user input, and neither is ever written into the audit, where the
    ' inapplicable field is BLANK.
    Dim index As Long, offset As Long, profile As Long
    Dim inflation() As Double, weights() As Double

    If package.Model.DriverCount = 0 Then
        BuildDriverFactors = True
        Exit Function
    End If
    ReDim package.Drivers(0 To package.Model.DriverCount - 1)
    ReDim inflation(0 To package.Model.Timeline.Duration - 1)
    ReDim weights(0 To package.Model.Timeline.Duration - 1)

    For index = 0 To package.Model.DriverCount - 1
        With package.Drivers(index)
            .PermanentId = package.Model.Drivers(index).PermanentId
            .IsRisk = package.Model.Drivers(index).IsRisk
            .DistKind = package.Model.Drivers(index).DistKind
            .MinValue = package.Model.Drivers(index).MinValue
            .MostLikely = package.Model.Drivers(index).MostLikely
            .MaxValue = package.Model.Drivers(index).MaxValue
            If .IsRisk Then
                .Quantity = 1#
                .Probability = package.Model.Drivers(index).Probability
            Else
                .Quantity = package.Model.Drivers(index).Quantity
                .Probability = 1#
            End If
        End With
        profile = ProfileIndexOf(package.Model, index)
        If profile < 0 Then
            detail = "driver " & package.Model.Drivers(index).PermanentId & _
                     ": the inflation profile is not in the resolved reference set"
            Exit Function
        End If
        For offset = 0 To package.Model.Timeline.Duration - 1
            inflation(offset) = InflationFor(package, profile, _
                                             package.Model.CalendarYears(offset))
            weights(offset) = package.Model.Weights(index, offset)
        Next offset
        If Not modCalcFactors.BuildKnom(package.Model.DriverFxRates(index), weights, _
                                        inflation, package.Drivers(index).Knom, _
                                        detail) Then
            detail = "driver " & package.Model.Drivers(index).PermanentId & " Knom: " & detail
            Exit Function
        End If
        If Not modCalcFactors.BuildKpv(package.Model.DriverFxRates(index), weights, _
                                       inflation, DiscountVector(package), _
                                       package.Drivers(index).Kpv, detail) Then
            detail = "driver " & package.Model.Drivers(index).PermanentId & " Kpv: " & detail
            Exit Function
        End If
    Next index
    BuildDriverFactors = True
End Function

Private Function BuildAudits(ByRef package As CalculationPackage, _
                             ByRef detail As String) As Boolean
    Dim index As Long
    If package.Model.DriverCount = 0 Then
        BuildAudits = True
        Exit Function
    End If
    ReDim package.Audits(0 To package.Model.DriverCount - 1)
    For index = 0 To package.Model.DriverCount - 1
        ' BuildDriverAudit derives Central and MeanValue and every published
        ' per-driver amount. Nothing is recomputed here.
        If Not modCalcAnalytical.BuildDriverAudit(package.Drivers(index), _
                                                  package.Audits(index), detail) Then
            detail = "driver " & package.Drivers(index).PermanentId & ": " & detail
            Exit Function
        End If
        package.Audits(index).PermanentId = package.Drivers(index).PermanentId
    Next index
    BuildAudits = True
End Function

Private Function BuildAnnual(ByRef package As CalculationPackage, _
                             ByRef detail As String) As Boolean
    Dim inflation() As Double, index As Long, offset As Long, profile As Long
    If package.Model.DriverCount > 0 Then
        ReDim inflation(0 To package.Model.DriverCount - 1, _
                        0 To package.Model.Timeline.Duration - 1)
        For index = 0 To package.Model.DriverCount - 1
            profile = ProfileIndexOf(package.Model, index)
            For offset = 0 To package.Model.Timeline.Duration - 1
                inflation(index, offset) = InflationFor(package, profile, _
                                                        package.Model.CalendarYears(offset))
            Next offset
        Next index
    End If
    BuildAnnual = modCalcAnalytical.BuildAnnualSeries( _
        package.Drivers, package.Model.DriverCount, package.Model.DriverFxRates, _
        package.Model.Weights, inflation, package.Years, package.Annual, _
        package.Magnitudes, detail)
End Function

Private Function BuildFingerprint(ByRef package As CalculationPackage, _
                                  ByRef detail As String) As Boolean
    ' The accepted schema, through the accepted encoder. The separator is taken
    ' from the SAME formatter the encoder uses, so the two cannot disagree, and
    ' no Excel object is passed into modCalcFingerprint.
    Dim header(0 To 3) As String, separator As String
    Dim costIds() As String, costRecords() As String, costCount As Long
    Dim riskIds() As String, riskRecords() As String, riskCount As Long
    Dim index As Long, record As String

    separator = HostDecimalSeparator()
    If Not NumberField(package.Model.Timeline.BaseYear, separator, header(0), detail) Then Exit Function
    If Not NumberField(package.Model.Timeline.StartYear, separator, header(1), detail) Then Exit Function
    If Not NumberField(package.Model.Timeline.Duration, separator, header(2), detail) Then Exit Function
    If Not NumberField(package.Model.Timeline.DiscountRate, separator, header(3), detail) Then Exit Function

    ReDim costIds(0 To 0): ReDim costRecords(0 To 0)
    ReDim riskIds(0 To 0): ReDim riskRecords(0 To 0)
    If package.Model.DriverCount > 0 Then
        ReDim costIds(0 To package.Model.DriverCount - 1)
        ReDim costRecords(0 To package.Model.DriverCount - 1)
        ReDim riskIds(0 To package.Model.DriverCount - 1)
        ReDim riskRecords(0 To package.Model.DriverCount - 1)
    End If
    For index = 0 To package.Model.DriverCount - 1
        If Not DriverRecord(package, index, separator, record, detail) Then Exit Function
        If package.Model.Drivers(index).IsRisk Then
            riskIds(riskCount) = package.Model.Drivers(index).PermanentId
            riskRecords(riskCount) = record
            riskCount = riskCount + 1
        Else
            costIds(costCount) = package.Model.Drivers(index).PermanentId
            costRecords(costCount) = record
            costCount = costCount + 1
        End If
    Next index

    If Not modCalcFingerprint.CalcFpBuildFingerprint(header, 4, costIds, costRecords, _
                                                     costCount, riskIds, riskRecords, _
                                                     riskCount, package.Fingerprint) Then
        detail = "the calculation fingerprint could not be constructed"
        Exit Function
    End If
    BuildFingerprint = True
End Function

Private Function DriverRecord(ByRef package As CalculationPackage, ByVal index As Long, _
                              ByVal separator As String, ByRef record As String, _
                              ByRef detail As String) As Boolean
    ' The opposite kind's identity is NOT fingerprinted: a cost record carries
    ' Quantity and a risk record carries Probability, and neither carries the
    ' other.
    Dim inflation() As Double, weights() As Double, offset As Long, profile As Long
    Dim ok As Boolean
    profile = ProfileIndexOf(package.Model, index)
    ReDim inflation(0 To package.Model.Timeline.Duration - 1)
    ReDim weights(0 To package.Model.Timeline.Duration - 1)
    For offset = 0 To package.Model.Timeline.Duration - 1
        inflation(offset) = InflationFor(package, profile, package.Model.CalendarYears(offset))
        weights(offset) = package.Model.Weights(index, offset)
    Next offset
    With package.Model.Drivers(index)
        If .IsRisk Then
            ok = modCalcFingerprint.CalcFpBuildRiskRecord( _
                .PermanentId, .Distribution, .Probability, .MinValue, .MaxValue, _
                .MostLikely, .HasMostLikely, package.Model.DriverFxRates(index), _
                inflation, weights, separator, record)
        Else
            ok = modCalcFingerprint.CalcFpBuildCostRecord( _
                .PermanentId, .Distribution, .Quantity, .MinValue, .MaxValue, _
                .MostLikely, .HasMostLikely, package.Model.DriverFxRates(index), _
                inflation, weights, separator, record)
        End If
        If Not ok Then
            detail = "driver " & .PermanentId & ": the fingerprint record could not be encoded"
            Exit Function
        End If
    End With
    DriverRecord = True
End Function

' ==========================================================================
' Status
' ==========================================================================
Private Function DeriveStatus(ByRef package As CalculationPackage, _
                              ByVal prepared As Boolean) As String
    ' THE CURRENT STATE DECIDES, and the historical attempt never does.
    '
    ' An empty digest is never compared. Two blanks are not a match: they are
    ' two absences, and reporting CURRENT from them would claim a calculation
    ' that never happened.
    Dim stored As String
    If Not prepared Then
        DeriveStatus = CALC_STATUS_INVALID
        Exit Function
    End If
    stored = StoredText(CALC_STATE_ROW_LAST_SUCCESSFUL_FINGERPRINT)
    If Len(stored) = 0 Then
        DeriveStatus = CALC_STATUS_NOT_CALCULATED
    ElseIf StrComp(package.Fingerprint, stored, vbBinaryCompare) = 0 Then
        DeriveStatus = CALC_STATUS_CURRENT
    Else
        DeriveStatus = CALC_STATUS_STALE
    End If
End Function

Private Function CurrentStatus() As String
    ' Freshly derived, from a fresh preparation. Used after a rollback, where
    ' the question is what the CURRENT inputs say about the RESTORED snapshot -
    ' never "the attempt failed, so the status is failed".
    Dim package As CalculationPackage, detail As String
    CurrentStatus = DeriveStatus(package, PrepareCurrentCalculation(package, detail))
End Function

' ==========================================================================
' Writing
' ==========================================================================
Private Sub WriteAnalytical(ByRef package As CalculationPackage)
    WriteTable TBL_CALC_YEARS, YearsBlock(package)
    WriteTable TBL_CALC_INFLATION_FACTORS, InflationBlock(package)
    WriteTable TBL_CALC_FX, FxBlock(package)
    WriteTable TBL_CALC_DRIVERS, DriversBlock(package)
    WriteTable TBL_CALC_ANNUAL, AnnualBlock(package)
    CalcSheet.Range(CALC_TOTALS_VALUE_RANGE).Value2 = TotalsBlock(package)
End Sub

Private Sub WriteTable(ByVal tableName As String, ByRef block As Variant)
    ' One resize, one block assignment. Never a cell at a time, and never a
    ' header or a number format: those are build-owned.
    '
    ' A ZERO-ROW TABLE KEEPS ITS PHYSICAL PLACEHOLDER AND CLEARS IT. The
    ' placeholder is not a semantic record, and no record is fabricated to fill
    ' it.
    Dim target As ListObject, rows As Long
    Set target = modWorkbook.Lo(CALC_SHEET, tableName)
    rows = SemanticRowCount(block)
    ResizeBody target, rows
    If rows = 0 Then
        target.DataBodyRange.ClearContents
    Else
        target.DataBodyRange.Value2 = block
    End If
End Sub

Private Sub ResizeBody(ByVal target As ListObject, ByVal rows As Long)
    ' Excel keeps at least one physical body row. Growing appends; shrinking
    ' deletes down to one and then clears it.
    Dim wanted As Long
    wanted = rows
    If wanted < 1 Then wanted = 1
    Do While modWorkbook.BodyRowCount(target) > wanted
        target.ListRows(modWorkbook.BodyRowCount(target)).Delete
    Loop
    Do While modWorkbook.BodyRowCount(target) < wanted
        target.ListRows.Add
    Loop
End Sub

Private Sub WriteSuccessCommit(ByRef package As CalculationPackage)
    ' ONE 8x1 assignment. Not four writes that could half-succeed and leave a
    ' fingerprint with no stamp, or a stamp with no version.
    Dim block(1 To 8, 1 To 1) As Variant
    block(1, 1) = Now
    block(2, 1) = package.Fingerprint
    block(3, 1) = FP_VERSION
    block(4, 1) = AppliedTimelineText(package)
    block(5, 1) = CALC_ATTEMPT_SUCCESS
    block(6, 1) = vbNullString
    block(7, 1) = CALC_STATUS_CURRENT
    block(8, 1) = Now
    CalcSheet.Range(CALC_STATE_VALUE_RANGE).Value2 = block
End Sub

Private Sub RecordRefusal(ByVal detail As String)
    ' A refusal touches C17:C20 and nothing else. C13:C16 - the last successful
    ' record - and every analytical block stand exactly as they were.
    WriteAttemptBlock CALC_ATTEMPT_REFUSED, detail, CALC_STATUS_INVALID
End Sub

Private Sub RecordFailure(ByVal detail As String)
    ' Only ever called AFTER a successful rollback. The status is derived afresh
    ' against the RESTORED snapshot: FAILED is an attempt result and never a
    ' status.
    WriteAttemptBlock CALC_ATTEMPT_FAILED, detail, CurrentStatus()
End Sub

Private Sub WriteAttemptBlock(ByVal result As String, ByVal detail As String, _
                              ByVal status As String)
    Dim block(1 To 4, 1 To 1) As Variant
    block(1, 1) = result
    block(2, 1) = detail
    block(3, 1) = status
    block(4, 1) = Now
    CalcSheet.Range(AttemptRange()).Value2 = block
End Sub

Private Sub WriteStatusBlock(ByVal status As String)
    Dim block(1 To 2, 1 To 1) As Variant
    block(1, 1) = status
    block(2, 1) = Now
    CalcSheet.Range(StatusRange()).Value2 = block
End Sub

' ==========================================================================
' Verification - a write is not proven by the absence of an error
' ==========================================================================
Private Function VerifyAnalytical(ByRef package As CalculationPackage) As Boolean
    If Not VerifyTable(TBL_CALC_YEARS, YearsBlock(package)) Then Exit Function
    If Not VerifyTable(TBL_CALC_INFLATION_FACTORS, InflationBlock(package)) Then Exit Function
    If Not VerifyTable(TBL_CALC_FX, FxBlock(package)) Then Exit Function
    If Not VerifyTable(TBL_CALC_DRIVERS, DriversBlock(package)) Then Exit Function
    If Not VerifyTable(TBL_CALC_ANNUAL, AnnualBlock(package)) Then Exit Function
    VerifyAnalytical = VerifyRange(CALC_TOTALS_VALUE_RANGE, TotalsBlock(package), 10)
End Function

Private Function VerifyTable(ByVal tableName As String, ByRef block As Variant) As Boolean
    Dim target As ListObject, rows As Long, columns As Long, r As Long, c As Long
    Set target = modWorkbook.Lo(CALC_SHEET, tableName)
    rows = SemanticRowCount(block)
    If modWorkbook.BodyRowCount(target) <> IIf(rows < 1, 1, rows) Then Exit Function
    If rows = 0 Then
        ' The placeholder must be BLANK. A blank verifies as blank, never as zero.
        For c = 1 To target.ListColumns.Count
            If Not IsEmpty(target.DataBodyRange.Cells(1, c).Value) Then Exit Function
        Next c
        VerifyTable = True
        Exit Function
    End If
    columns = UBound(block, 2)
    For r = 1 To rows
        For c = 1 To columns
            If Not SameCell(target.DataBodyRange.Cells(r, c).Value, block(r, c)) Then Exit Function
        Next c
    Next r
    VerifyTable = True
End Function

Private Function VerifyRange(ByVal address As String, ByRef block As Variant, _
                             ByVal rows As Long) As Boolean
    Dim index As Long
    For index = 1 To rows
        If Not SameCell(CalcSheet.Range(address).Cells(index, 1).Value, _
                        block(index, 1)) Then Exit Function
    Next index
    VerifyRange = True
End Function

Private Function VerifySuccessCommit(ByRef package As CalculationPackage) As Boolean
    ' The fingerprint, the version, the applied timeline, the attempt result and
    ' the status are the parts that must be exactly what was intended. The two
    ' timestamps are not compared against a recomputed Now.
    If StoredText(CALC_STATE_ROW_LAST_SUCCESSFUL_FINGERPRINT) <> package.Fingerprint Then Exit Function
    If StoredText(CALC_STATE_ROW_LAST_SUCCESSFUL_APPLIED_TIMELINE) <> _
       AppliedTimelineText(package) Then Exit Function
    If StoredText(CALC_STATE_ROW_LAST_ATTEMPT_RESULT) <> CALC_ATTEMPT_SUCCESS Then Exit Function
    If Len(StoredText(CALC_STATE_ROW_LAST_ATTEMPT_DETAIL)) <> 0 Then Exit Function
    If StoredText(CALC_STATE_ROW_CALCULATION_STATUS) <> CALC_STATUS_CURRENT Then Exit Function
    If StoredValue(CALC_STATE_ROW_FINGERPRINT_VERSION) <> FP_VERSION Then Exit Function
    If Len(StoredText(CALC_STATE_ROW_LAST_SUCCESSFUL_STAMP)) = 0 Then Exit Function
    VerifySuccessCommit = True
End Function

Private Function SameCell(ByVal written As Variant, ByVal wanted As Variant) As Boolean
    ' Blank matches only blank. A fabricated zero in an N/A field would not
    ' verify, which is the point.
    If IsEmpty(wanted) Or (VarType(wanted) = vbString And Len(wanted) = 0) Then
        SameCell = IsEmpty(written) Or (VarType(written) = vbString And Len(written) = 0)
        Exit Function
    End If
    If IsEmpty(written) Then Exit Function
    If VarType(wanted) = vbString Then
        SameCell = (VarType(written) = vbString) And (StrComp(CStr(written), CStr(wanted), _
                                                              vbBinaryCompare) = 0)
        Exit Function
    End If
    If VarType(written) = vbString Then Exit Function
    SameCell = (CDbl(written) = CDbl(wanted))
End Function

' ==========================================================================
' The analytical blocks
'
' Each returns a 1-based 2-D Variant sized to its SEMANTIC row count, or Empty
' where that count is zero. Variant is unavoidable here and is deliberate: a
' cell must be able to hold a number, a text label or a BLANK, and blank is a
' value this schema uses to mean "does not apply".
' ==========================================================================
Private Function YearsBlock(ByRef package As CalculationPackage) As Variant
    Dim block() As Variant, index As Long, rows As Long
    rows = package.Model.Timeline.Duration
    If rows < 1 Then Exit Function
    ReDim block(1 To rows, 1 To TBL_CALC_YEARS_COLUMN_COUNT)
    For index = 1 To rows
        block(index, COL_CALC_YEARS_PROJECT_INDEX) = package.Years(index - 1).ProjectIndex
        block(index, COL_CALC_YEARS_CALENDAR_YEAR) = package.Years(index - 1).CalendarYear
        block(index, COL_CALC_YEARS_DISCOUNT_FACTOR) = package.Years(index - 1).DiscountF
    Next index
    YearsBlock = block
End Function

Private Function InflationBlock(ByRef package As CalculationPackage) As Variant
    ' One row per referenced profile per calendar year, Base Year first.
    '
    ' THE BASE YEAR HAS NO ANNUAL RATE. Its Annual Rate cell is BLANK - a
    ' model-controlled blank, not a fabricated zero - and its cumulative factor
    ' is 1. Where Base Year < Start Year the pre-project compounding years stay
    ' visible, which is the whole point of auditing the span rather than the
    ' project years.
    Dim block() As Variant, profile As Long, offset As Long, row As Long, rows As Long
    rows = package.Model.ProfileCount * package.InflationSpan
    If rows < 1 Then Exit Function
    ReDim block(1 To rows, 1 To TBL_CALC_INFLATION_FACTORS_COLUMN_COUNT)
    For profile = 0 To package.Model.ProfileCount - 1
        For offset = 0 To package.InflationSpan - 1
            row = row + 1
            block(row, COL_CALC_INFLATION_FACTORS_INFLATION_PROFILE) = _
                package.Model.Profiles(profile)
            block(row, COL_CALC_INFLATION_FACTORS_CALENDAR_YEAR) = _
                package.Model.Timeline.BaseYear + offset
            If offset = 0 Then
                block(row, COL_CALC_INFLATION_FACTORS_ANNUAL_RATE) = Empty
            Else
                block(row, COL_CALC_INFLATION_FACTORS_ANNUAL_RATE) = _
                    package.Model.InflationRates(profile, offset - 1)
            End If
            block(row, COL_CALC_INFLATION_FACTORS_CUMULATIVE_INFLATION_FACTOR) = _
                package.InflationFactors(profile, offset)
        Next offset
    Next profile
    InflationBlock = block
End Function

Private Function FxBlock(ByRef package As CalculationPackage) As Variant
    ' REFERENCED currencies only. The global reporting-currency invariant does
    ' not create an audit row: being validated is not being referenced, and an
    ' empty driver set produces no rows at all.
    Dim block() As Variant, index As Long, rows As Long
    rows = package.Model.CurrencyCount
    If rows < 1 Then Exit Function
    ReDim block(1 To rows, 1 To TBL_CALC_FX_COLUMN_COUNT)
    For index = 1 To rows
        block(index, COL_CALC_FX_CURRENCY) = package.Model.Currencies(index - 1)
        block(index, COL_CALC_FX_FX_TO_SAR) = package.Model.CurrencyRates(index - 1)
        block(index, COL_CALC_FX_REFERENCED_BY) = package.ReferencedBy(index - 1)
    Next index
    FxBlock = block
End Function

Private Function DriversBlock(ByRef package As CalculationPackage) As Variant
    ' AN INAPPLICABLE FIELD IS BLANK. Never the in-memory identity 1, and never
    ' zero: a Quantity of 1 shown against a risk would read as a real entry, and
    ' a zero would read as a real amount.
    Dim block() As Variant, index As Long, rows As Long, row As Long
    rows = package.Model.DriverCount
    If rows < 1 Then Exit Function
    ReDim block(1 To rows, 1 To TBL_CALC_DRIVERS_COLUMN_COUNT)
    For index = 0 To rows - 1
        row = index + 1
        block(row, COL_CALC_DRIVERS_PERMANENT_ID) = package.Model.Drivers(index).PermanentId
        block(row, COL_CALC_DRIVERS_DISTRIBUTION) = package.Model.Drivers(index).Distribution
        block(row, COL_CALC_DRIVERS_CURRENCY) = package.Model.Drivers(index).Currency
        block(row, COL_CALC_DRIVERS_FX_TO_SAR) = package.Model.DriverFxRates(index)
        block(row, COL_CALC_DRIVERS_INFLATION_PROFILE) = _
            package.Model.Drivers(index).InflationProfile
        block(row, COL_CALC_DRIVERS_MEAN_VALUE) = package.Audits(index).MeanValue
        block(row, COL_CALC_DRIVERS_KNOM) = package.Audits(index).Knom
        block(row, COL_CALC_DRIVERS_KPV) = package.Audits(index).Kpv
        If package.Model.Drivers(index).IsRisk Then
            block(row, COL_CALC_DRIVERS_DRIVER_KIND) = DRIVER_KIND_RISK
            block(row, COL_CALC_DRIVERS_CENTRAL_BASIS) = Empty
            block(row, COL_CALC_DRIVERS_QUANTITY) = Empty
            block(row, COL_CALC_DRIVERS_PROBABILITY) = package.Model.Drivers(index).Probability
            block(row, COL_CALC_DRIVERS_CENTRAL_VALUE) = Empty
            block(row, COL_CALC_DRIVERS_DETERMINISTIC_NOMINAL) = Empty
            block(row, COL_CALC_DRIVERS_DETERMINISTIC_PV) = Empty
            block(row, COL_CALC_DRIVERS_MEAN_BASIS_NOMINAL) = Empty
            block(row, COL_CALC_DRIVERS_MEAN_BASIS_PV) = Empty
            block(row, COL_CALC_DRIVERS_UNCERTAINTY_MEAN_SHIFT_NOMINAL) = Empty
            block(row, COL_CALC_DRIVERS_UNCERTAINTY_MEAN_SHIFT_PV) = Empty
            block(row, COL_CALC_DRIVERS_EXPECTED_RISK_NOMINAL) = _
                package.Audits(index).ExpectedRiskNominal
            block(row, COL_CALC_DRIVERS_EXPECTED_RISK_PV) = package.Audits(index).ExpectedRiskPv
        Else
            block(row, COL_CALC_DRIVERS_DRIVER_KIND) = DRIVER_KIND_COST
            block(row, COL_CALC_DRIVERS_CENTRAL_BASIS) = package.Audits(index).CentralBasis
            block(row, COL_CALC_DRIVERS_QUANTITY) = package.Model.Drivers(index).Quantity
            block(row, COL_CALC_DRIVERS_PROBABILITY) = Empty
            block(row, COL_CALC_DRIVERS_CENTRAL_VALUE) = package.Audits(index).Central
            block(row, COL_CALC_DRIVERS_DETERMINISTIC_NOMINAL) = _
                package.Audits(index).DeterministicNominal
            block(row, COL_CALC_DRIVERS_DETERMINISTIC_PV) = package.Audits(index).DeterministicPv
            block(row, COL_CALC_DRIVERS_MEAN_BASIS_NOMINAL) = _
                package.Audits(index).MeanBasisNominal
            block(row, COL_CALC_DRIVERS_MEAN_BASIS_PV) = package.Audits(index).MeanBasisPv
            block(row, COL_CALC_DRIVERS_UNCERTAINTY_MEAN_SHIFT_NOMINAL) = _
                package.Audits(index).ShiftNominal
            block(row, COL_CALC_DRIVERS_UNCERTAINTY_MEAN_SHIFT_PV) = _
                package.Audits(index).ShiftPv
            block(row, COL_CALC_DRIVERS_EXPECTED_RISK_NOMINAL) = Empty
            block(row, COL_CALC_DRIVERS_EXPECTED_RISK_PV) = Empty
        End If
    Next index
    DriversBlock = block
End Function

Private Function AnnualBlock(ByRef package As CalculationPackage) As Variant
    Dim block() As Variant, index As Long, rows As Long
    rows = package.Model.Timeline.Duration
    If rows < 1 Then Exit Function
    ReDim block(1 To rows, 1 To TBL_CALC_ANNUAL_COLUMN_COUNT)
    For index = 1 To rows
        With package.Annual(index - 1)
            block(index, COL_CALC_ANNUAL_PROJECT_INDEX) = .ProjectIndex
            block(index, COL_CALC_ANNUAL_CALENDAR_YEAR) = .CalendarYear
            block(index, COL_CALC_ANNUAL_BASE_COST_NOMINAL) = .BaseCostNominal
            block(index, COL_CALC_ANNUAL_EXPECTED_RISK_NOMINAL) = .ExpectedRiskNominal
            block(index, COL_CALC_ANNUAL_TOTAL_NOMINAL) = .TotalNominal
            block(index, COL_CALC_ANNUAL_BASE_COST_PV) = .BaseCostPv
            block(index, COL_CALC_ANNUAL_EXPECTED_RISK_PV) = .ExpectedRiskPv
            block(index, COL_CALC_ANNUAL_TOTAL_PV) = .TotalPv
        End With
    Next index
    AnnualBlock = block
End Function

Private Function TotalsBlock(ByRef package As CalculationPackage) As Variant
    ' The ten totals already accumulated in memory. The audit table is never
    ' summed to produce them: the worksheet records the calculation and is never
    ' an input back into one.
    Dim block(1 To 10, 1 To 1) As Variant
    block(1, 1) = package.Totals.ANom
    block(2, 1) = package.Totals.APv
    block(3, 1) = package.Totals.BNom
    block(4, 1) = package.Totals.BPv
    block(5, 1) = package.Totals.CNom
    block(6, 1) = package.Totals.CPv
    block(7, 1) = package.Totals.DNom
    block(8, 1) = package.Totals.DPv
    block(9, 1) = package.Totals.ENom
    block(10, 1) = package.Totals.EPv
    TotalsBlock = block
End Function

' ==========================================================================
' Snapshot and rollback
' ==========================================================================
Private Sub CaptureSnapshot(ByRef snapshot As CalculationSnapshot)
    snapshot.Years = modWorkbook.SnapshotTable(modWorkbook.Lo(CALC_SHEET, TBL_CALC_YEARS))
    snapshot.Inflation = modWorkbook.SnapshotTable( _
        modWorkbook.Lo(CALC_SHEET, TBL_CALC_INFLATION_FACTORS))
    snapshot.Fx = modWorkbook.SnapshotTable(modWorkbook.Lo(CALC_SHEET, TBL_CALC_FX))
    snapshot.Drivers = modWorkbook.SnapshotTable(modWorkbook.Lo(CALC_SHEET, TBL_CALC_DRIVERS))
    snapshot.Annual = modWorkbook.SnapshotTable(modWorkbook.Lo(CALC_SHEET, TBL_CALC_ANNUAL))
    ' VALUES only. Labels, notes and number formats are build-owned and are
    ' neither captured nor rewritten.
    snapshot.Totals = CalcSheet.Range(CALC_TOTALS_VALUE_RANGE).Value2
    snapshot.State = CalcSheet.Range(CALC_STATE_VALUE_RANGE).Value2
    snapshot.Captured = True
End Sub

Private Sub RestoreSnapshot(ByRef snapshot As CalculationSnapshot)
    If Not snapshot.Captured Then
        Err.Raise vbObjectError + 5103, "modCalcReport.RestoreSnapshot", _
                  "no snapshot was captured; refusing to restore"
    End If
    modWorkbook.RestoreTable modWorkbook.Lo(CALC_SHEET, TBL_CALC_YEARS), snapshot.Years
    modWorkbook.RestoreTable modWorkbook.Lo(CALC_SHEET, TBL_CALC_INFLATION_FACTORS), _
                             snapshot.Inflation
    modWorkbook.RestoreTable modWorkbook.Lo(CALC_SHEET, TBL_CALC_FX), snapshot.Fx
    modWorkbook.RestoreTable modWorkbook.Lo(CALC_SHEET, TBL_CALC_DRIVERS), snapshot.Drivers
    modWorkbook.RestoreTable modWorkbook.Lo(CALC_SHEET, TBL_CALC_ANNUAL), snapshot.Annual
    CalcSheet.Range(CALC_TOTALS_VALUE_RANGE).Value2 = snapshot.Totals
    CalcSheet.Range(CALC_STATE_VALUE_RANGE).Value2 = snapshot.State
End Sub

' ==========================================================================
' Small helpers
' ==========================================================================
Private Function CalcSheet() As Worksheet
    Set CalcSheet = modWorkbook.Sh(CALC_SHEET)
End Function

Private Function StoredText(ByVal row As Long) As String
    StoredText = modWorkbook.TextOf(StateCell(row))
End Function

Private Function StoredValue(ByVal row As Long) As Double
    Dim value As Double
    If modWorkbook.TryReadDouble(StateCell(row).Value, value) Then StoredValue = value
End Function

Private Function StateCell(ByVal row As Long) As Range
    Set StateCell = CalcSheet.Range(CALC_STATE_VALUE_COLUMN & CStr(row))
End Function

Private Function AttemptRange() As String
    AttemptRange = CALC_STATE_VALUE_COLUMN & CStr(CALC_STATE_ROW_LAST_ATTEMPT_RESULT) & _
                   ":" & CALC_STATE_VALUE_COLUMN & CStr(CALC_STATE_ROW_STATUS_EVALUATED_AT)
End Function

Private Function StatusRange() As String
    StatusRange = CALC_STATE_VALUE_COLUMN & CStr(CALC_STATE_ROW_CALCULATION_STATUS) & _
                  ":" & CALC_STATE_VALUE_COLUMN & CStr(CALC_STATE_ROW_STATUS_EVALUATED_AT)
End Function

Private Function AppliedTimelineText(ByRef package As CalculationPackage) As String
    AppliedTimelineText = CStr(package.Model.Timeline.BaseYear) & "/" & _
                          CStr(package.Model.Timeline.StartYear) & "/" & _
                          CStr(package.Model.Timeline.Duration)
End Function

Private Function SemanticRowCount(ByRef block As Variant) As Long
    If IsEmpty(block) Then Exit Function
    SemanticRowCount = UBound(block, 1)
End Function

Private Function ProfileIndexOf(ByRef model As ResolvedModel, ByVal index As Long) As Long
    Dim probe As Long
    ProfileIndexOf = -1
    For probe = 0 To model.ProfileCount - 1
        If StrComp(model.Profiles(probe), model.Drivers(index).InflationProfile, _
                   vbBinaryCompare) = 0 Then
            ProfileIndexOf = probe
            Exit Function
        End If
    Next probe
End Function

Private Function InflationFor(ByRef package As CalculationPackage, ByVal profile As Long, _
                              ByVal calendarYear As Long) As Double
    ' Calendar-year anchored, exactly as the resolver anchored the rates.
    InflationFor = package.InflationFactors(profile, calendarYear - _
                                            package.Model.Timeline.BaseYear)
End Function

Private Function DiscountVector(ByRef package As CalculationPackage) As Double()
    Dim out() As Double, index As Long
    ReDim out(0 To package.Model.Timeline.Duration - 1)
    For index = 0 To package.Model.Timeline.Duration - 1
        out(index) = package.Years(index).DiscountF
    Next index
    DiscountVector = out
End Function

Private Function HostDecimalSeparator() As String
    ' Taken from the SAME formatter the encoder uses, so the separator it is
    ' told about cannot disagree with the one it will see. No Application
    ' setting is consulted, and no Excel object reaches modCalcFingerprint.
    HostDecimalSeparator = Mid$(Format$(0#, "0.0"), 2, 1)
End Function

Private Function NumberField(ByVal value As Double, ByVal separator As String, _
                             ByRef field As String, ByRef detail As String) As Boolean
    If Not modCalcFingerprint.CalcFpCanonicalNumber(value, separator, field) Then
        detail = "a header value could not be canonically encoded"
        Exit Function
    End If
    field = modCalcFingerprint.CalcFpCanonicalText(field)
    NumberField = True
End Function

Private Sub CountCurrencyReferences(ByRef package As CalculationPackage)
    Dim currency As Long, driver As Long
    If package.Model.CurrencyCount < 1 Then Exit Sub
    ReDim package.ReferencedBy(0 To package.Model.CurrencyCount - 1)
    For currency = 0 To package.Model.CurrencyCount - 1
        For driver = 0 To package.Model.DriverCount - 1
            If StrComp(package.Model.Drivers(driver).Currency, _
                       package.Model.Currencies(currency), vbBinaryCompare) = 0 Then
                package.ReferencedBy(currency) = package.ReferencedBy(currency) + 1
            End If
        Next driver
    Next currency
End Sub
