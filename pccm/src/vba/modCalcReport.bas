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
    ' THE WHOLE INVOCATION IS INSIDE AN ENVELOPE, installed before the first
    ' fallible operation. Application state is captured, changed, and must be
    ' restored however this returns - a VBA error that bypassed the normal path
    ' would leave EnableEvents, Calculation mode and ScreenUpdating dirty, which
    ' is the one outcome worse than a failed calculation.
    '
    ' The result is published through Announce, NOT ReportResult: Announce
    ' records the outcome for automation and shows a dialog only when automation
    ' is inactive. Gate B drives this endpoint through that mechanism, and a
    ' direct dialog would block it.
    '
    ' CLEANUP IS ATTEMPTED AT MOST ONCE, AND WHICHEVER CONTEXT ATTEMPTS IT DOES
    ' SO INSIDE AN ENVELOPE. FinishOperation returns a diagnostic String for a
    ' restoration it could not complete, but it is an Excel call and can also
    ' RAISE - and a raise on the NORMAL path would escape this endpoint with no
    ' announcement, no recorded automation result, and stateCaptured still True.
    '
    ' Two Boolean facts carry the distinction, not statement position:
    '   stateCaptured    - a snapshot exists and something still owes a restore
    '   cleanupAttempted - the one permitted attempt has been spent, raise or not
    Dim state As AppStateSnapshot, result As OperationResult
    Dim stateCaptured As Boolean, cleanupAttempted As Boolean, committed As Boolean
    Dim cleanup As String, failure As String

    On Error GoTo InvocationFailed
    state = modAppState.CaptureAppState()
    stateCaptured = True
    modAppState.BeginOperation
    result = RunCalculation(committed)
    On Error GoTo 0

    On Error GoTo NormalCleanupFailed
    cleanupAttempted = True
    cleanup = modAppState.FinishOperation(state)
    On Error GoTo 0
    stateCaptured = False

    If Len(cleanup) > 0 Then result = CleanupOutcome(result, committed, cleanup)
    modAppState.Announce result
    Exit Sub

NormalCleanupFailed:
    ' THE NORMAL CLEANUP RAISED. The attempt is spent and is NOT retried: a
    ' restoration that failed by raising is not made more likely by running it
    ' again, and running it again is how an endpoint restores twice.
    '
    ' Nothing analytical is touched and calc_state is not rewritten, so a
    ' committed calculation stays committed and C17 keeps saying SUCCESS. The
    ' problem is reported on the invocation axis through the SAME outcome rule
    ' the returned-diagnostic path uses.
    failure = Err.Description
    On Error GoTo 0
    modAppState.Announce CleanupOutcome(result, committed, _
        "Application state could not be restored: " & failure)
    Exit Sub

InvocationFailed:
    failure = Err.Description
    On Error GoTo CleanupFailed
    ' Reached only from BEFORE the normal cleanup, so the attempt is still
    ' available - but that is asserted from state rather than assumed from where
    ' this label sits.
    If stateCaptured And Not cleanupAttempted Then
        cleanupAttempted = True
        cleanup = modAppState.FinishOperation(state)
        stateCaptured = False
        If Len(cleanup) > 0 Then failure = failure & vbCrLf & cleanup
    End If
    On Error GoTo 0
    modAppState.Announce modAppState.Failed("Calculate", failure)
    Exit Sub

CleanupFailed:
    ' The recovery cleanup attempt itself raised. It is not retried either, and
    ' the original failure must still reach the caller rather than being replaced
    ' by the failure of the attempt to recover from it.
    On Error GoTo 0
    modAppState.Announce modAppState.Failed("Calculate", failure & vbCrLf & _
        "Application state could not be restored after the failure.")
End Sub

Private Function CleanupOutcome(ByRef result As OperationResult, ByVal committed As Boolean, _
                                ByVal cleanup As String) As OperationResult
    ' A CLEANUP PROBLEM AFTER THE COMMIT IS AN INVOCATION FAILURE, NOT A FAILED
    ' CALCULATION.
    '
    ' Once the transaction commits, C17 says SUCCESS and that is committed
    ' workbook truth. Re-reading the attempt as FAILED while the workbook still
    ' says SUCCESS would publish two contradictory answers, so nothing here
    ' rewrites calc_state, rolls anything back or touches an analytical block.
    ' The cleanup problem is surfaced on the invocation axis instead, where it
    ' belongs, and the message says plainly that the calculation committed.
    If committed Then
        CleanupOutcome = modAppState.Failed("Calculate", _
            "The calculation COMMITTED successfully and the calculation state " & _
            "records it. Application state could not be fully restored " & _
            "afterwards:" & vbCrLf & cleanup)
        Exit Function
    End If
    ' Before the commit there is no committed truth to contradict, so the two
    ' diagnostics are simply combined.
    If result.Ok Then
        CleanupOutcome = modAppState.Failed("Calculate", cleanup)
    Else
        CleanupOutcome = modAppState.Failed("Calculate", result.Detail & vbCrLf & cleanup)
    End If
End Function

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
Private Function RunCalculation(ByRef committed As Boolean) As OperationResult
    Dim package As CalculationPackage, snapshot As CalculationSnapshot
    Dim successBlock As Variant
    Dim detail As String, prepared As Boolean

    ' PREPARATION AND SNAPSHOT SIT IN THEIR OWN ENVELOPE. A CONTROLLED refusal
    ' from the accepted machinery is REFUSED; an unexpected runtime error in the
    ' same region is FAILED. Downgrading a runtime fault to a refusal to keep
    ' going would report a model problem the user does not have.
    On Error GoTo PreWriteFailed
    prepared = PrepareCurrentCalculation(package, detail)
    If prepared Then CaptureSnapshot snapshot
    On Error GoTo 0

    If Not prepared Then
        ' Nothing analytical has been touched, so nothing is rolled back.
        RunCalculation = RecordRefusal(detail)
        Exit Function
    End If

    On Error GoTo TransactionFailed
    WriteAnalytical package
    modAppState.FailPointCheck FAILPOINT_ANALYTICAL_WRITE
    If Not VerifyAnalytical(package) Then
        Err.Raise vbObjectError + 5101, "modCalcReport.RunCalculation", _
                  "the analytical snapshot did not verify against the prepared result"
    End If
    ' BUILT ONCE, WRITTEN ONCE, VERIFIED AGAINST THE SAME BLOCK. Both timestamps
    ' are captured into it here; verification never generates a second Now,
    ' which would compare against a value the commit never contained.
    BuildSuccessBlock package, successBlock
    WriteSuccessCommit successBlock
    If Not VerifySuccessCommit(successBlock) Then
        Err.Raise vbObjectError + 5102, "modCalcReport.RunCalculation", _
                  "the success commit did not verify"
    End If
    committed = True
    On Error GoTo 0
    RunCalculation = modAppState.Succeeded("Calculation committed.")
    Exit Function

PreWriteFailed:
    ' An UNEXPECTED failure before any analytical mutation. FAILED, never
    ' REFUSED - and no rollback, because nothing was written.
    detail = Err.Description
    On Error GoTo 0
    RunCalculation = RecordFailureWithoutRollback(detail)
    Exit Function

TransactionFailed:
    detail = Err.Description
    On Error GoTo 0
    RunCalculation = RollbackAndRecord(snapshot, detail)
End Function

Private Function RecordRefusal(ByVal detail As String) As OperationResult
    ' A refusal touches C17:C20 and nothing else. C13:C16 - the last successful
    ' record - and every analytical block stand exactly as they were.
    Dim note As String
    On Error GoTo BookkeepingFailed
    WriteAttemptBlock CALC_ATTEMPT_REFUSED, detail, CALC_STATUS_INVALID
    On Error GoTo 0
    RecordRefusal = modAppState.Failed("Calculate", detail)
    Exit Function
BookkeepingFailed:
    note = Err.Description
    On Error GoTo 0
    RecordRefusal = modAppState.Failed("Calculate", detail & vbCrLf & _
        "The refusal could not be recorded in the calculation state: " & note)
End Function

Private Function RecordFailureWithoutRollback(ByVal detail As String) As OperationResult
    ' No analytical state was mutated, so there is nothing to restore - but the
    ' attempt is still FAILED and is still recorded if the workbook permits it.
    Dim note As String
    On Error GoTo BookkeepingFailed
    WriteAttemptBlock CALC_ATTEMPT_FAILED, detail, CurrentStatus()
    On Error GoTo 0
    RecordFailureWithoutRollback = modAppState.Failed("Calculate", detail)
    Exit Function
BookkeepingFailed:
    note = Err.Description
    On Error GoTo 0
    RecordFailureWithoutRollback = modAppState.Failed("Calculate", detail & vbCrLf & _
        "The failed attempt could not be recorded in the calculation state: " & note)
End Function

Private Function RollbackAndRecord(ByRef snapshot As CalculationSnapshot, _
                                   ByVal detail As String) As OperationResult
    ' ROLLBACK FIRST, metadata second, and the metadata is written ONLY because
    ' the rollback succeeded.
    Dim note As String
    On Error GoTo RollbackFailed
    RestoreSnapshot snapshot
    On Error GoTo 0

    On Error GoTo BookkeepingFailed
    WriteAttemptBlock CALC_ATTEMPT_FAILED, detail, CurrentStatus()
    On Error GoTo 0
    RollbackAndRecord = modAppState.Failed("Calculate", detail)
    Exit Function

RollbackFailed:
    ' The restoration itself failed. NO failed-attempt metadata is written: that
    ' record asserts "the previous snapshot stands", and asserting it here would
    ' be a claim nobody has established. Both diagnostics are preserved.
    note = Err.Description
    On Error GoTo 0
    RollbackAndRecord = modAppState.Failed("Calculate", _
        "The calculation failed AND the previous snapshot could not be fully " & _
        "restored." & vbCrLf & "Original failure: " & detail & vbCrLf & _
        "Restore failure: " & note)
    Exit Function

BookkeepingFailed:
    ' The rollback succeeded and is NOT undone. Only its record failed, and the
    ' previous successful snapshot remains authoritative.
    note = Err.Description
    On Error GoTo 0
    RollbackAndRecord = modAppState.Failed("Calculate", detail & vbCrLf & _
        "The previous successful snapshot was restored, but the failed attempt " & _
        "could not be recorded: " & note)
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

    ' THE FOUR HEADER SCALARS ARE NUMBER FIELDS. They go through the accepted
    ' N-field encoder, not through the text one: tagging a number as text would
    ' change what the digest covers, and the framing authority is
    ' modCalcFingerprint's in either case.
    separator = HostDecimalSeparator()
    If Not modCalcFingerprint.CalcFpNumberField(package.Model.Timeline.BaseYear, _
            separator, header(0)) Then GoTo HeaderFailed
    If Not modCalcFingerprint.CalcFpNumberField(package.Model.Timeline.StartYear, _
            separator, header(1)) Then GoTo HeaderFailed
    If Not modCalcFingerprint.CalcFpNumberField(package.Model.Timeline.Duration, _
            separator, header(2)) Then GoTo HeaderFailed
    If Not modCalcFingerprint.CalcFpNumberField(package.Model.Timeline.DiscountRate, _
            separator, header(3)) Then GoTo HeaderFailed

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
    Exit Function
HeaderFailed:
    detail = "a header value could not be canonically encoded"
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

Private Sub BuildSuccessBlock(ByRef package As CalculationPackage, ByRef block As Variant)
    ' ONE 8x1 block, built once. Both timestamps are the SAME captured moment:
    ' the commit and the status evaluation it publishes happen together, and
    ' capturing them here is what lets verification compare against the exact
    ' values that were written.
    Dim built(1 To 8, 1 To 1) As Variant
    Dim stamp As Date
    stamp = Now
    built(1, 1) = stamp
    built(2, 1) = package.Fingerprint
    built(3, 1) = FP_VERSION
    built(4, 1) = AppliedTimelineText(package)
    built(5, 1) = CALC_ATTEMPT_SUCCESS
    built(6, 1) = vbNullString
    built(7, 1) = CALC_STATUS_CURRENT
    built(8, 1) = stamp
    block = built
End Sub

Private Sub WriteSuccessCommit(ByRef block As Variant)
    ' ONE assignment. Not four writes that could half-succeed and leave a
    ' fingerprint with no stamp, or a stamp with no version.
    '
    ' THE COMMIT FAILPOINT SITS HERE, IMMEDIATELY BEFORE THE ASSIGNMENT, because
    ' the boundary Gate B has to exercise is the final C13:C20 write itself. A
    ' hook several statements upstream proves a failure during commit
    ' PREPARATION - the analytical snapshot is written and verified either way,
    ' but the block has not been built and nothing has been published, so the
    ' rollback it exercises is not the rollback from the commit boundary.
    ' Nothing fallible stands between the check and the write.
    modAppState.FailPointCheck FAILPOINT_SUCCESS_COMMIT
    CalcSheet.Range(CALC_STATE_VALUE_RANGE).Value2 = block
End Sub

Private Function VerifySuccessCommit(ByRef block As Variant) As Boolean
    ' ALL EIGHT CELLS, against the SAME block that was written - including C20,
    ' and including both timestamps as the values the commit actually carried.
    ' A verifier that regenerated Now, or that checked a stamp only for being
    ' non-blank, would be proving something other than "this commit landed".
    VerifySuccessCommit = VerifyRange(CALC_STATE_VALUE_RANGE, block, 8)
End Function

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
    '
    ' WHICH FIELDS ARE INAPPLICABLE IS THE CONTRACT'S DECISION, not this
    ' function's. calc_contract declares applies_to per column, and Central
    ' Basis is declared for BOTH kinds - it labels the distribution, not the
    ' deterministic value. Publishing it blank for a Risk was a defect, found
    ' by Runtime Run 10 as case 9 / R-001.central_basis actual BLANK against an
    ' expected 'ML'.
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
            block(row, COL_CALC_DRIVERS_CENTRAL_BASIS) = package.Audits(index).CentralBasis
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

Private Sub CountCurrencyReferences(ByRef package As CalculationPackage)
    Dim currencyIndex As Long, driver As Long
    If package.Model.CurrencyCount < 1 Then Exit Sub
    ReDim package.ReferencedBy(0 To package.Model.CurrencyCount - 1)
    For currencyIndex = 0 To package.Model.CurrencyCount - 1
        For driver = 0 To package.Model.DriverCount - 1
            If StrComp(package.Model.Drivers(driver).Currency, _
                       package.Model.Currencies(currencyIndex), vbBinaryCompare) = 0 Then
                package.ReferencedBy(currencyIndex) = package.ReferencedBy(currencyIndex) + 1
            End If
        Next driver
    Next currencyIndex
End Sub
' ==========================================================================
' STEP 11 ADDITION - THE PHASE-6 PREPARATION BRIDGE
'
' EVERYTHING ABOVE THIS BANNER IS THE ACCEPTED PHASE-5 REPORTER, BYTE FOR BYTE.
' The accepted digest gates hash the text before this line and still require the
' accepted literals, so "and nothing else" keeps its full meaning.
'
' WHY THIS EXISTS. Phase 6 needs the resolved DriverFactors, the CURRENT
' analytical fingerprint, the deterministic base estimate A and the applied
' timeline. All four already exist inside PrepareCurrentCalculation, and a
' second construction of any of them is a second answer that drifts the first
' time either side changes. So the phase reaches THIS preparation instead of
' rebuilding it.
'
' IT IS AN INTERNAL CROSS-MODULE API. It is Public because modSimReport must
' reach it; it is not an automation endpoint, its name does not begin with
' PCCM_, and no button binds to it.
'
' IT WRITES NOTHING. No `_Calc` table, no calc_state cell, no attempt metadata
' and no status. Asking Phase 5 for its current package is not a calculation,
' and a read that silently recalculated would make a simulation depend on when
' it was asked.
' ==========================================================================
Public Function CalcPrepareSimulationInputs(ByRef drivers() As DriverFactors, _
                                            ByRef driverCount As Long, _
                                            ByRef analyticalFingerprint As String, _
                                            ByRef deterministicBaseNominal As Double, _
                                            ByRef deterministicBasePv As Double, _
                                            ByRef appliedTimeline As String, _
                                            ByRef decimalSeparator As String, _
                                            ByRef detail As String) As Boolean
    Dim package As CalculationPackage
    Dim status As String

    detail = vbNullString
    ' The SAME accepted preparation the endpoint uses. Not a copy of it.
    If Not PrepareCurrentCalculation(package, detail) Then Exit Function

    ' D6-14: Phase 6 runs on a CURRENT Phase 5 or it does not run. This does not
    ' call PCCM_Calculate and does not repair anything - a simulation that
    ' silently recalculated its own inputs would publish a distribution nobody
    ' asked for.
    status = DeriveStatus(package, True)
    If StrComp(status, CALC_STATUS_CURRENT, vbBinaryCompare) <> 0 Then
        detail = "the calculation is " & status & _
                 "; the simulation needs a CURRENT calculation"
        Exit Function
    End If

    ' PROJECTION ONLY. Every value below is read off the package that was just
    ' prepared; not one is recomputed here, and no factor, FX rate, inflation
    ' factor, profile weight, Knom or Kpv is constructed in this procedure.
    '
    ' A ZERO-DRIVER MODEL SUCCEEDS. The carrier is handed over with its logical
    ' count, and no semantic driver is inspected - the accepted zero-count
    ' convention, unchanged.
    drivers = package.Drivers
    driverCount = package.Model.DriverCount
    analyticalFingerprint = package.Fingerprint
    deterministicBaseNominal = package.Totals.ANom
    deterministicBasePv = package.Totals.APv
    appliedTimeline = AppliedTimelineText(package)
    decimalSeparator = HostDecimalSeparator()

    CalcPrepareSimulationInputs = True
End Function
