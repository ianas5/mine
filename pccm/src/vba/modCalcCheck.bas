Attribute VB_Name = "modCalcCheck"
Option Explicit

' ==========================================================================
' modCalcCheck - the Phase-5 NUMERICAL prerequisite checker.
'
' It reports. It refuses. IT NEVER REPAIRS.
'
' Nothing here writes a corrected value back into the resolved model, clamps a
' probability, forces a quantity, normalises a profile, replaces an ordering or
' adjusts a year. A prerequisite checker that quietly fixed its input would be
' calculating from a model the user never entered, and the refusal it suppressed
' is the only thing that would have told them so.
'
' --------------------------------------------------------------------------
' THREE OWNERS, AND THIS IS THE THIRD
' --------------------------------------------------------------------------
'   Phase 4          the STRUCTURAL prerequisites - the structural state and
'                    ValidateStructure(): ID patterns, duplicates, orphans,
'                    profiling and inflation grid shape, counters. Already
'                    invoked by modCalcResolve before anything is resolved.
'   Step 5           RESOLUTION - reference sets, referenced-only FX and
'                    inflation, exact identifiers, strict numeric typing,
'                    blank versus zero, Permanent-ID profiling.
'   Step 6, HERE     the NUMERICAL prerequisites over the already-resolved
'                    model.
'
' None of the first two is re-implemented here. A second copy of a rule is a
' second authority, and two authorities disagree eventually.
'
' --------------------------------------------------------------------------
' IT CHECKS THE RESOLVED MODEL, NOT THE WORKBOOK AGAIN
' --------------------------------------------------------------------------
' Every value this module needs is already in ResolvedModel. Re-reading a cell
' to obtain a value Step 5 has resolved would create a second resolution
' authority and break the locked pipeline: resolve everything into memory,
' validate everything in memory, calculate everything in memory. So there is no
' worksheet access here at all - no Range, no ListObject, no defined name.
'
' --------------------------------------------------------------------------
' AND IT OWNS NO ARITHMETIC
' --------------------------------------------------------------------------
' The profiling sum goes through modCalcFactors.SafeSignedSum, and the
' difference from 100% through SafeSubtract. This module defines no numerical
' kernel of its own and holds no tolerance of its own: the tolerance is the
' generated TOL_PROFILING_SUM_ABSOLUTE, whose authority is the calculation
' contract.
' ==========================================================================

' The target a profile's weights must sum to. Not a tolerance and not a
' magnitude - it is 100%, expressed once so the comparison and the diagnostic
' cannot drift apart.
Private Const PROFILE_SUM_TARGET As Double = 1#

' ==========================================================================
' The entry point
' ==========================================================================
Public Function CheckResolvedModel(ByRef model As ResolvedModel, _
                                   ByRef detail As String) As Boolean
    ' Model-level predicates first, then per-driver ones. A model with no
    ' drivers still has a timeline and a discount rate, and both must hold.
    Dim index As Long
    detail = vbNullString
    If Not CheckTimeline(model.Timeline, detail) Then Exit Function
    If Not CheckDiscountRate(model.Timeline, detail) Then Exit Function

    ' AN EMPTY DRIVER SET IS VALID. It produces zero ordering checks, zero
    ' scalar checks and zero profiling checks - not a refusal. The count is
    ' tested before any bound of Drivers or Weights is read, because a VBA array
    ' cannot represent a zero-element set and an unallocated one raises on
    ' LBound.
    If model.DriverCount < 0 Then
        detail = "the resolved driver count is negative"
        Exit Function
    End If
    If model.DriverCount = 0 Then
        CheckResolvedModel = True
        Exit Function
    End If

    For index = 0 To model.DriverCount - 1
        If Not CheckDriver(model.Drivers(LBound(model.Drivers) + index), detail) Then
            Exit Function
        End If
        If Not CheckProfileSum(model, index, detail) Then Exit Function
    Next index
    CheckResolvedModel = True
End Function

' ==========================================================================
' Model-level predicates
' ==========================================================================
Private Function CheckTimeline(ByRef timeline As ResolvedTimeline, _
                               ByRef detail As String) As Boolean
    ' The price Base Year may equal the project Start Year, and may precede it -
    ' pre-project years then participate in inflation compounding. It may not
    ' FOLLOW it: escalating a cost from a base that has not happened yet is not
    ' a model, it is a sign error.
    '
    ' Phase 4 may also report this relationship structurally. That does not
    ' remove the Phase-5 predicate: two consumers at different boundaries is
    ' fine, and a numerical layer that assumed someone else had checked would be
    ' assuming rather than checking.
    If timeline.BaseYear > timeline.StartYear Then
        detail = "applied timeline: Base Year " & CStr(timeline.BaseYear) & _
                 " is later than Start Year " & CStr(timeline.StartYear) & _
                 ". The price base year cannot postdate the project start year."
        Exit Function
    End If
    CheckTimeline = True
End Function

Private Function CheckDiscountRate(ByRef timeline As ResolvedTimeline, _
                                   ByRef detail As String) As Boolean
    ' D3: 1 + r > 0. A discount rate of -100% or lower makes the discount
    ' divisor zero or negative and no present-value factor exists. Step 5 has
    ' already proven the value is a usable Double, so for a finite Double the
    ' condition is exactly r <= -1 and no arithmetic is needed to test it.
    '
    ' The rate is NOT clamped, NOT defaulted to zero and NOT replaced with an
    ' identity factor. BuildDiscountFactors stays in modCalcFactors and stays
    ' unweakened; it is simply never reached with a rate that cannot produce a
    ' factor.
    If timeline.DiscountRate <= -1# Then
        detail = "discount rate: the rate gives 1 + r <= 0; a discount rate of " & _
                 "-100% or lower is refused"
        Exit Function
    End If
    CheckDiscountRate = True
End Function

' ==========================================================================
' Per-driver predicates
' ==========================================================================
Private Function CheckDriver(ByRef driver As ResolvedDriver, _
                             ByRef detail As String) As Boolean
    If Not CheckOrdering(driver, detail) Then Exit Function
    If driver.IsRisk Then
        If Not CheckProbability(driver, detail) Then Exit Function
    Else
        If Not CheckQuantity(driver, detail) Then Exit Function
    End If
    CheckDriver = True
End Function

Private Function CheckOrdering(ByRef driver As ResolvedDriver, _
                               ByRef detail As String) As Boolean
    ' The three-point ordering, on the DistKind Step 5 already resolved. The
    ' distribution NAME is not mapped a second time.
    '
    ' NO POSITIVITY RULE IS INVENTED. A correctly ordered set of negative values
    ' is a valid distribution - a credit, a saving, a transfer out - and no
    ' accepted contract says otherwise. Only the ordering is checked.
    '
    ' Nothing here calls a mean or a central-value function. An ordering check
    ' that computed a statistic would be doing the calculation early, and would
    ' refuse for a representability reason that has nothing to do with ordering.
    Select Case driver.DistKind
    Case DIST_UNIFORM
        ' D1: Uniform is a TWO-POINT distribution. A populated Most Likely is
        ' accepted and IGNORED - the cell may hold a leftover from another
        ' choice of distribution, and refusing it would block a valid model.
        ' MostLikely is deliberately not read on this path.
        If driver.MinValue > driver.MaxValue Then
            detail = OrderingFailure(driver, "requires Min <= Max")
            Exit Function
        End If
    Case DIST_TRIANGULAR, DIST_BETA_PERT
        If driver.MinValue > driver.MostLikely Or driver.MostLikely > driver.MaxValue Then
            detail = OrderingFailure(driver, "requires Min <= Most Likely <= Max")
            Exit Function
        End If
    Case Else
        ' Unreachable through the resolver, which refuses an unmapped
        ' distribution. Stated anyway, because a silently unchecked driver is
        ' worse than a refusal nobody expected to see.
        detail = DriverLabel(driver) & ": distribution kind is not recognised"
        Exit Function
    End Select
    CheckOrdering = True
End Function

Private Function CheckQuantity(ByRef driver As ResolvedDriver, _
                               ByRef detail As String) As Boolean
    ' A cost line's Quantity must be strictly positive: a zero or negative
    ' quantity has no meaning as a multiplier of a unit cost.
    '
    ' Probability is NOT examined here. The `Probability = 1 for cost lines`
    ' convention belongs to the in-memory DriverFactors carry type that the
    ' calculation and the simulation share; it is not a user input and there is
    ' nothing for a user to have got wrong.
    If driver.Quantity <= 0# Then
        detail = DriverLabel(driver) & ": Quantity must be strictly positive"
        Exit Function
    End If
    CheckQuantity = True
End Function

Private Function CheckProbability(ByRef driver As ResolvedDriver, _
                                  ByRef detail As String) As Boolean
    ' A risk's Probability is a fraction in the CLOSED interval [0, 1]. Both
    ' boundaries are valid: a risk that certainly happens and a risk that
    ' certainly does not are both expressible models.
    '
    ' Quantity is NOT examined here, for the mirror of the reason above.
    If driver.Probability < 0# Or driver.Probability > 1# Then
        detail = DriverLabel(driver) & ": Probability must be a fraction in [0, 1]"
        Exit Function
    End If
    CheckProbability = True
End Function

' ==========================================================================
' The profiling sum - the rule Step 5 deliberately deferred
'
' Step 5 resolved the cells: exact, blank distinguished from zero, attached by
' Permanent ID, in applied project-year order. What it did not do is add them
' up, because the sum needs the accepted signed-sum authority and the tolerance
' belongs with the calculation contract. Both live here.
' ==========================================================================
Private Function CheckProfileSum(ByRef model As ResolvedModel, ByVal index As Long, _
                                 ByRef detail As String) As Boolean
    Dim weights() As Double, offset As Long, count As Long
    Dim total As Double, difference As Double

    count = model.Timeline.Duration
    If count > 0 Then
        ReDim weights(0 To count - 1)
        For offset = 0 To count - 1
            weights(offset) = model.Weights(LBound(model.Weights, 1) + index, _
                                            LBound(model.Weights, 2) + offset)
        Next offset
    End If

    ' SIGNED, and through the accepted primitive. A profile may legitimately
    ' contain negative weights, so the sum must not be refused merely because
    ' two large opposite-signed weights are adjacent in project-year order
    ' (erratum C2). A hand-written accumulation loop here would be a second
    ' summation rule with none of that behaviour.
    If Not modCalcFactors.SafeSignedSum(weights, count, total) Then
        detail = DriverLabel(model.Drivers(LBound(model.Drivers) + index)) & _
                 ": the profiling weights cannot be summed within Double range"
        Exit Function
    End If

    ' The difference goes through the safe primitive too. Both operands are
    ' usable Doubles so this cannot fail today, but a checker that reported a
    ' tolerance breach from an uncontrolled subtraction would be reporting
    ' nonsense, and the guard costs nothing.
    If Not modCalcFactors.SafeSubtract(total, PROFILE_SUM_TARGET, difference) Then
        detail = DriverLabel(model.Drivers(LBound(model.Drivers) + index)) & _
                 ": the profiling weight sum cannot be compared against 100%"
        Exit Function
    End If

    If Abs(difference) > TOL_PROFILING_SUM_ABSOLUTE Then
        detail = DriverLabel(model.Drivers(LBound(model.Drivers) + index)) & _
                 ": the profiling weights sum to " & CStr(total) & _
                 ", which is not " & CStr(PROFILE_SUM_TARGET) & " within " & _
                 CStr(TOL_PROFILING_SUM_ABSOLUTE)
        Exit Function
    End If
    CheckProfileSum = True
End Function

' ==========================================================================
' Diagnostics
'
' A refusal a user cannot act on is barely better than a crash, so every
' message names the driver it is about. The label is built once so the scope
' cannot drift between one check and the next.
' ==========================================================================
Private Function DriverLabel(ByRef driver As ResolvedDriver) As String
    If driver.IsRisk Then
        DriverLabel = "risk " & driver.PermanentId
    Else
        DriverLabel = "cost line " & driver.PermanentId
    End If
End Function

Private Function OrderingFailure(ByRef driver As ResolvedDriver, _
                                 ByVal rule As String) As String
    OrderingFailure = DriverLabel(driver) & ": " & driver.Distribution & " " & rule
End Function
