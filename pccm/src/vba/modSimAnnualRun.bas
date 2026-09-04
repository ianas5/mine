Attribute VB_Name = "modSimAnnualRun"
Option Explicit

' ==========================================================================
' PCCM Phase 7 - THE ANNUAL STOCHASTIC ENDPOINT.
'
' It produces TWO different objects from one replay, and they are not
' interchangeable:
'
'   THE ANNUAL DISTRIBUTIONS  the accepted eleven-rung ladder taken per project
'                             year, across iterations. A property of the RUN.
'                             No selector enters it. It does NOT sum to the
'                             total percentile and it is NOT a profile.
'   THE SELECTED-Px PROFILE   the annual vectors of the two iterations owning
'                             the type-7 lo and hi positions of the selected
'                             Px total, blended by the same f. A property of
'                             the run AND of the reporting selector.
'
' WHY THAT DISTINCTION IS LOAD-BEARING. Selected Confidence Level is a REPORTING
' selector: moving it does not invalidate the simulation, does not consume a run
' and does not require a re-simulation. So it must not retire the ladders
' either - they would be marked stale for a reason that has nothing to do with
' them. The profile is the opposite: it is the blend at ONE resolved Px, it
' stays valid for that Px, and it stops being the profile anyone is asking for
' the moment the selector resolves elsewhere. It is never relabelled.
'
' That is why the stamp records the Px, and why the handoff carries two states.
'
' --------------------------------------------------------------------------
' WHAT IT IS NOT
' --------------------------------------------------------------------------
' NOT A SIMULATION. It allocates no run id, advances no AUTO nonce, touches no
' pending-nonce marker, writes no attempt row, writes no iteration record and
' leaves the result digest exactly as it found it. A successful simulation must
' stay successful even if the annual step later fails, and the user must be able
' to repeat the annual step without producing a new stochastic result.
'
' NOT A SECOND MONTE CARLO. The replay reproduces the accepted run's own draws
' from its own effective seed, through the engine that owns the generator and
' the contribution rule.
'
' IT RECOMPUTES NO INFLATION, NO FX AND NO DISCOUNT FACTOR. The per-year factors
' are the TERMS of the sums Phase 5 already formed, built by modSimAnnual from
' the resolved inputs DriverFactors carries. The project-year axis is read back
' from the CURRENT Phase-5 year table rather than rebuilt.
'
' IT NAMES NO CELL. Every read and write belongs to modSimAnnualStore; this
' module spells no sheet, no column and no row.
'
' --------------------------------------------------------------------------
' MEMORY
' --------------------------------------------------------------------------
' Never iterations x years. One BLOCK of years exists at a time - iterations x
' SIM_ANNUAL_BLOCK_WIDTH doubles - and is replaced by the next block. What
' accumulates is the answer: eleven rungs per year per measure, plus one profile
' value per year per measure. No iteration-level annual value is retained, in
' memory or on a sheet.
' ==========================================================================

Private Const ANNUAL_OPERATION As String = "Run Annual Stochastic"

' ==========================================================================
' THE ENDPOINT
' ==========================================================================
Public Sub PCCM_RunAnnualStochastic()
    ' THE SAME ENVELOPE THE ACCEPTED ENDPOINTS USE, and for the same reason: a
    ' failure must not leave ScreenUpdating off or Calculation manual. There is
    ' no `committed` distinction here - the annual step commits nothing that a
    ' cleanup problem could strand, because it changes no simulation state.
    Dim state As AppStateSnapshot, result As OperationResult
    Dim stateCaptured As Boolean, cleanupAttempted As Boolean
    Dim cleanup As String, failure As String

    On Error GoTo InvocationFailed
    state = modAppState.CaptureAppState()
    stateCaptured = True
    modAppState.BeginOperation
    result = RunAnnual()
    On Error GoTo 0

    On Error GoTo NormalCleanupFailed
    cleanupAttempted = True
    cleanup = modAppState.FinishOperation(state)
    On Error GoTo 0
    stateCaptured = False

    If Len(cleanup) > 0 Then
        result = modAppState.Failed(ANNUAL_OPERATION, _
            "Application state could not be restored: " & cleanup)
    End If
    modAppState.Announce result
    Exit Sub

NormalCleanupFailed:
    failure = Err.Description
    On Error GoTo 0
    modAppState.Announce modAppState.Failed(ANNUAL_OPERATION, _
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
    modAppState.Announce modAppState.Failed(ANNUAL_OPERATION, failure)
    Exit Sub

CleanupFailed:
    On Error GoTo 0
    modAppState.Announce modAppState.Failed(ANNUAL_OPERATION, failure & vbCrLf & _
        "Application state could not be restored after the failure.")
End Sub

' ==========================================================================
' THE PIPELINE
'
' Every step must succeed before ANYTHING is written. Both measures are produced
' in full and both are reconciled before the block is cleared - so a refusal in
' the PV pass cannot leave the nominal ladders on the sheet looking like a
' complete current answer.
' ==========================================================================
Private Function RunAnnual() As OperationResult
    Dim run As SimAnnualIdentity
    Dim drivers() As DriverFactors
    Dim projectIndex() As Long, calendarYear() As Long, discount() As Double
    Dim nominalFactors() As Double, pvFactors() As Double
    Dim nominalTotals() As Double, pvTotals() As Double
    Dim nominalAt As SimStatsPosition, pvAt As SimStatsPosition
    Dim nominalLadder() As Double, pvLadder() As Double
    Dim nominalProfile() As Double, pvProfile() As Double
    Dim probabilities() As Double
    Dim flat() As Double
    Dim driverCount As Long, fields As Long
    Dim detail As String

    If Not modSimAnnualStore.SimAnnualStoreCurrentRun(run, detail) Then
        RunAnnual = Refused(detail)
        Exit Function
    End If
    If Not ResolveSelectedPx(run, detail) Then
        RunAnnual = Refused(detail)
        Exit Function
    End If
    If Not ResolveDrivers(drivers, driverCount, detail) Then
        RunAnnual = Refused(detail)
        Exit Function
    End If
    If Not modSimAnnualStore.SimAnnualStoreYearAxis(projectIndex, calendarYear, _
                                                    discount, run.YearCount, detail) Then
        RunAnnual = Refused(detail)
        Exit Function
    End If
    If Not CrossCheckYearCount(drivers, driverCount, run.YearCount, detail) Then
        RunAnnual = Refused(detail)
        Exit Function
    End If

    ' THE FACTORS, AND THE PROOF THAT REGROUPING THEM CHANGED NOTHING. Each
    ' driver's per-year terms must sum back to the scalar Phase 5 built, to the
    ' project's own identity allowance - never to bit equality, because FX
    ' applied once to a summed staging and FX applied in every year are the same
    ' real number and may be different Doubles.
    If Not BuildYearFactors(drivers, driverCount, run.YearCount, discount, False, _
                            nominalFactors, detail) Then
        RunAnnual = Refused(detail)
        Exit Function
    End If
    If Not BuildYearFactors(drivers, driverCount, run.YearCount, discount, True, _
                            pvFactors, detail) Then
        RunAnnual = Refused(detail)
        Exit Function
    End If

    ' THE RUNGS ARE THE ACCEPTED LADDER'S. Not a ladder of this module's own:
    ' the annual distributions are the SAME eleven probabilities the summary
    ' publishes, taken per project year.
    If Not modSimStats.SimStatsLadderProbabilities(probabilities, detail) Then
        RunAnnual = Refused("annual: " & detail)
        Exit Function
    End If

    ' THE POSITIONS COME FROM THE PUBLISHED TOTALS. lo, hi and f are the SAME
    ' type-7 position that produced the reported total Px; nothing here looks
    ' for an iteration close to Px, and nothing recomputes the percentile.
    If Not modSimAnnualStore.SimAnnualStoreTotals(run, SIM_MEASURE_NOMINAL, _
                                                  nominalTotals, detail) Then
        RunAnnual = Refused(detail)
        Exit Function
    End If
    If Not modSimAnnualStore.SimAnnualStoreTotals(run, SIM_MEASURE_PV, pvTotals, _
                                                  detail) Then
        RunAnnual = Refused(detail)
        Exit Function
    End If
    If Not modSimStats.SimStatsQuantilePosition(nominalTotals, run.Iterations, _
            run.SelectedProbability, nominalAt, detail) Then
        RunAnnual = Refused("annual, nominal: " & detail)
        Exit Function
    End If
    If Not modSimStats.SimStatsQuantilePosition(pvTotals, run.Iterations, _
            run.SelectedProbability, pvAt, detail) Then
        RunAnnual = Refused("annual, PV: " & detail)
        Exit Function
    End If

    If Not ProduceAnnual(run, drivers, driverCount, nominalFactors, SIM_MEASURE_NOMINAL, _
                         nominalAt, probabilities, nominalLadder, nominalProfile, _
                         detail) Then
        RunAnnual = Refused(detail)
        Exit Function
    End If
    If Not ProduceAnnual(run, drivers, driverCount, pvFactors, SIM_MEASURE_PV, _
                         pvAt, probabilities, pvLadder, pvProfile, detail) Then
        RunAnnual = Refused(detail)
        Exit Function
    End If

    ' THE PROFILE'S OWN IDENTITY, CHECKED. Every iteration's annual vector sums
    ' to that iteration's total, so the convex blend of two of them sums to the
    ' convex blend of their totals - which IS the selected Px total. A profile
    ' that failed this would be a shape not belonging to the number the sheet
    ' reports. The LADDERS are deliberately not checked this way: a ladder does
    ' not sum to the total percentile and never did.
    If Not ReconcileProfile(nominalProfile, run, nominalTotals, "nominal", detail) Then
        RunAnnual = Refused(detail)
        Exit Function
    End If
    If Not ReconcileProfile(pvProfile, run, pvTotals, "PV", detail) Then
        RunAnnual = Refused(detail)
        Exit Function
    End If

    If Not modSimAnnualStore.SimAnnualStoreFlatten(run, projectIndex, calendarYear, _
            nominalLadder, pvLadder, nominalProfile, pvProfile, flat, fields, _
            detail) Then
        RunAnnual = Refused(detail)
        Exit Function
    End If

    ' NOTHING HAS BEEN WRITTEN UNTIL HERE.
    If Not modSimAnnualStore.SimAnnualStorePublish(run, flat, fields, detail) Then
        RunAnnual = Refused(detail)
        Exit Function
    End If

    RunAnnual.Ok = True
    RunAnnual.Message = "Annual stochastic complete: " & CStr(run.YearCount) & _
                        " project year(s) over " & CStr(run.Iterations) & _
                        " iterations. The selected-Px profile is " & _
                        run.SelectedLabel & "."
End Function

' ==========================================================================
' THE REPORTING SELECTOR
'
' RESOLVED, NEVER INVENTED. The label is the accepted input cell's, read by the
' store; which labels are selectable and what probability a label spells belong
' to modSimStats and its one projected ladder. This module spells no confidence
' level and decodes no label.
'
' A SELECTOR THAT CANNOT BE RESOLVED REFUSES THE WHOLE STEP, rather than
' publishing ladders with an unstamped profile beside them - an unstamped
' profile is exactly the object the stamp exists to prevent.
' ==========================================================================
Private Function ResolveSelectedPx(ByRef run As SimAnnualIdentity, _
                                   ByRef detail As String) As Boolean
    Dim label As String
    Dim p As Double

    If Not modSimAnnualStore.SimAnnualStoreSelector(label, detail) Then Exit Function
    If Not modSimStats.SimStatsSelectedProbability(label, p, detail) Then
        detail = "annual: the selected confidence level cannot be resolved: " & detail
        Exit Function
    End If
    run.SelectedLabel = label
    run.SelectedProbability = p
    ResolveSelectedPx = True
End Function

' ==========================================================================
' THE RESOLVED MODEL - through the ONE accepted Phase-5 bridge
'
' The bridge refuses unless Phase 5 is CURRENT, which is what makes the year
' axis the store reads back the axis THIS model resolves to.
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
        detail = "annual: the model resolves no drivers to decompose"
        Exit Function
    End If
    ResolveDrivers = True
End Function

' THE AXIS AND THE MODEL MUST AGREE ABOUT HOW MANY YEARS THERE ARE. Both come
' from ONE current calculation, so a disagreement does not get reconciled - it
' means the two were not produced together, and every per-year factor built from
' them would be paired with the wrong year.
Private Function CrossCheckYearCount(ByRef drivers() As DriverFactors, _
                                     ByVal driverCount As Long, _
                                     ByVal yearCount As Long, _
                                     ByRef detail As String) As Boolean
    Dim index As Long, carried As Long

    For index = 0 To driverCount - 1
        carried = UBound(drivers(LBound(drivers) + index).Weights) - _
                  LBound(drivers(LBound(drivers) + index).Weights) + 1
        If carried <> yearCount Then
            detail = "annual: driver " & drivers(LBound(drivers) + index).PermanentId & _
                     " carries " & CStr(carried) & " project year(s) and the " & _
                     "published year axis carries " & CStr(yearCount)
            Exit Function
        End If
    Next index
    CrossCheckYearCount = True
End Function

' ==========================================================================
' THE PER-YEAR FACTORS, AND THE PROOF THAT REGROUPING CHANGED NOTHING
'
' `factors` is flat, driver-major, stride yearCount - the caller's SUPPLY order,
' which the engine maps to canonical order BY PERMANENT ID and never by
' position.
'
' Each driver's terms are summed back and compared against the scalar Phase 5
' built, to the project's own I3c/I4c allowance. NOTHING IS SCALED, NUDGED OR
' NORMALISED to make that comparison pass: a driver outside the allowance is a
' refusal, and the refusal names the driver.
' ==========================================================================
Private Function BuildYearFactors(ByRef drivers() As DriverFactors, _
                                  ByVal driverCount As Long, ByVal yearCount As Long, _
                                  ByRef discount() As Double, _
                                  ByVal withDiscount As Boolean, _
                                  ByRef factors() As Double, _
                                  ByRef detail As String) As Boolean
    Dim index As Long, year As Long
    Dim terms() As Double
    Dim against As Double
    Dim what As String

    ReDim factors(0 To driverCount * yearCount - 1)
    For index = 0 To driverCount - 1
        If Not modSimAnnual.SimAnnualFactors( _
                drivers(LBound(drivers) + index).FxRate, _
                drivers(LBound(drivers) + index).Weights, _
                drivers(LBound(drivers) + index).Inflation, _
                discount, withDiscount, terms, detail) Then
            detail = "annual: driver " & drivers(LBound(drivers) + index).PermanentId & _
                     ": " & detail
            Exit Function
        End If
        If withDiscount Then
            against = drivers(LBound(drivers) + index).Kpv
            what = "Kpv"
        Else
            against = drivers(LBound(drivers) + index).Knom
            what = "Knom"
        End If
        If Not ReconcileTerms(terms, yearCount, against, "annual: driver " & _
                drivers(LBound(drivers) + index).PermanentId & " " & what, _
                detail) Then
            Exit Function
        End If
        For year = 0 To yearCount - 1
            factors(index * yearCount + year) = terms(LBound(terms) + year)
        Next year
    Next index
    BuildYearFactors = True
End Function

' ==========================================================================
' SUM OF TERMS = AGGREGATE, to the accepted allowance
'
' THE CONDITIONING SCALE SUMS THE CONTRIBUTIONS, NOT THE AGGREGATES. That is
' ERRATUM C1 and it matters here: the terms can cancel, and a scale built from
' the aggregate alone would collapse to nothing exactly where the cancellation
' is largest. The floor is applied by IdentityAllowance as a MAXIMUM, never as
' an addition, and both the coefficient and the floor are the project's own.
' ==========================================================================
Private Function ReconcileTerms(ByRef terms() As Double, ByVal count As Long, _
                                ByVal against As Double, ByVal what As String, _
                                ByRef detail As String) As Boolean
    Dim summed As Double, difference As Double, allowance As Double
    Dim scale As Double
    Dim index As Long

    If Not modCalcFactors.SafeSignedSum(terms, count, summed) Then
        detail = what & ": the per-year terms do not sum to a representable number"
        Exit Function
    End If
    scale = 0#
    For index = 0 To count - 1
        If Not modCalcFactors.ConditioningScaledMagnitude(scale, _
                terms(LBound(terms) + index), TOL_IDENTITY_RELATIVE_COEFFICIENT) Then
            detail = what & ": the reconciliation scale is not representable"
            Exit Function
        End If
    Next index
    If Not modCalcFactors.ConditioningScaledMagnitude(scale, against, _
            TOL_IDENTITY_RELATIVE_COEFFICIENT) Then
        detail = what & ": the reconciliation scale is not representable"
        Exit Function
    End If
    If Not modCalcFactors.IdentityAllowance(scale, TOL_IDENTITY_ABSOLUTE_FLOOR, _
            TOL_IDENTITY_RELATIVE_COEFFICIENT, TOL_CONDITIONING_SCALE_FLOOR, _
            allowance) Then
        detail = what & ": the reconciliation allowance is not representable"
        Exit Function
    End If
    If Not modCalcFactors.SafeSubtract(summed, against, difference) Then
        detail = what & ": the reconciliation difference is not representable"
        Exit Function
    End If
    If Abs(difference) > allowance Then
        detail = what & ": the per-year terms sum to " & CStr(summed) & " against " & _
                 CStr(against) & ", outside the allowance of " & CStr(allowance)
        Exit Function
    End If
    ReconcileTerms = True
End Function

' ==========================================================================
' ONE MEASURE, ONE BLOCK OF YEARS AT A TIME
'
' The block is the ONLY iteration-level annual object that ever exists, and it
' is replaced on the next pass. The eleven rungs and the two profile values for
' those years are lifted out of it and the block is dropped.
'
' THE PROFILE'S TWO ITERATIONS ARE THE POSITION'S, NOT A SEARCH'S. LoSource and
' HiSource are original iteration indices from the type-7 position over the
' published totals, and the blend uses that position's own f.
' ==========================================================================
Private Function ProduceAnnual(ByRef run As SimAnnualIdentity, _
                               ByRef drivers() As DriverFactors, _
                               ByVal driverCount As Long, ByRef factors() As Double, _
                               ByVal measure As String, ByRef at As SimStatsPosition, _
                               ByRef probabilities() As Double, ByRef ladder() As Double, _
                               ByRef profile() As Double, ByRef detail As String) As Boolean
    Dim column() As Double, blockLadder() As Double, blockProfile() As Double
    Dim low() As Double, high() As Double
    Dim blocks As Long, block As Long, firstYear As Long, blockYears As Long
    Dim year As Long, rung As Long

    If run.YearCount < 1 Then
        detail = "annual, " & measure & ": the project covers no year to decompose"
        Exit Function
    End If
    ReDim ladder(0 To run.YearCount * SIM_ANNUAL_QUANTILE_COUNT - 1)
    ReDim profile(0 To run.YearCount - 1)
    blocks = modSimAnnual.SimAnnualBlockCount(run.YearCount, SIM_ANNUAL_BLOCK_WIDTH)
    If blocks < 1 Then
        detail = "annual, " & measure & ": the year blocking produced no pass"
        Exit Function
    End If

    For block = 0 To blocks - 1
        firstYear = block * SIM_ANNUAL_BLOCK_WIDTH
        blockYears = run.YearCount - firstYear
        If blockYears > SIM_ANNUAL_BLOCK_WIDTH Then blockYears = SIM_ANNUAL_BLOCK_WIDTH
        If Not modSimEngine.SimEngineReplayAnnualBlock(drivers, driverCount, _
                run.EffectiveSeed, run.Iterations, factors, run.YearCount, _
                firstYear, blockYears, measure, column, detail) Then
            detail = "annual, " & measure & ": " & detail
            Exit Function
        End If
        If Not modSimAnnual.SimAnnualLadder(column, run.Iterations, blockYears, probabilities, _
                                            SIM_ANNUAL_QUANTILE_COUNT, blockLadder, _
                                            detail) Then
            detail = "annual, " & measure & ": " & detail
            Exit Function
        End If

        ' THE TWO ORDER STATISTICS' OWN ANNUAL VECTORS, lifted out of the block.
        ReDim low(0 To blockYears - 1)
        ReDim high(0 To blockYears - 1)
        For year = 0 To blockYears - 1
            low(year) = column(at.LoSource * blockYears + year)
            high(year) = column(at.HiSource * blockYears + year)
        Next year
        If Not modSimAnnual.SimAnnualProfile(low, high, blockYears, at.Fraction, _
                                             blockProfile, detail) Then
            detail = "annual, " & measure & ": " & detail
            Exit Function
        End If

        For year = 0 To blockYears - 1
            profile(firstYear + year) = blockProfile(year)
            For rung = 0 To SIM_ANNUAL_QUANTILE_COUNT - 1
                ladder((firstYear + year) * SIM_ANNUAL_QUANTILE_COUNT + rung) = _
                    blockLadder(year * SIM_ANNUAL_QUANTILE_COUNT + rung)
            Next rung
        Next year
    Next block
    ProduceAnnual = True
End Function

' THE PROFILE SUMS TO THE SELECTED Px TOTAL, to the same allowance the factors
' were reconciled to. The comparison is against the ACCEPTED owner's percentile
' over the published totals, never against a number recomputed here.
Private Function ReconcileProfile(ByRef profile() As Double, _
                                  ByRef run As SimAnnualIdentity, _
                                  ByRef totals() As Double, ByVal what As String, _
                                  ByRef detail As String) As Boolean
    Dim px As Double

    If Not modSimStats.SimStatsQuantileType7(totals, run.Iterations, _
            run.SelectedProbability, px, detail) Then
        detail = "annual, " & what & ": " & detail
        Exit Function
    End If
    ReconcileProfile = ReconcileTerms(profile, run.YearCount, px, _
        "annual: the " & what & " " & run.SelectedLabel & " profile", detail)
End Function

Private Function Refused(ByVal detail As String) As OperationResult
    Refused.Ok = False
    Refused.Message = "The annual stochastic answer was not produced."
    Refused.Detail = detail
End Function
