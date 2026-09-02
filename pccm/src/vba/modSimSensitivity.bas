Attribute VB_Name = "modSimSensitivity"
Option Explicit

' ==========================================================================
' PCCM Phase 7 - sensitivity mathematics. WORKSHEET-FREE BY CONSTRUCTION.
'
' Mid-ranks, Spearman as Pearson-on-ranks, the undefined case, and the ordering
' of finished driver results. That is the whole of this module.
'
' builder/pccm_builder/sim_sensitivity.py is the single definition of these
' semantics; this module is their VBA implementation.
'
' IT KNOWS NOTHING ELSE. No RNG, no component stream, no sampler, no replay, no
' iteration engine, no request fingerprint, no result digest, no run_id, no AUTO
' nonce, no _SimData, no Results, no Sensitivity sheet, no run state and no
' workbook. It receives finished sequences of Doubles and returns statistics
' over them. That separation is why a wrong statistic can be substituted here
' without touching the engine, and why the engine cannot quietly acquire a
' correlation of its own.
'
' WHAT IT DELIBERATELY DOES NOT DO. It does not reconstruct a driver's
' contribution column - that is replay, and replay reads component streams. By
' the time a sequence arrives here the observations already exist.
'
' ----------------------------------------------------------------------------
' WHY THE SCALE NORMALISATION OF modSimStats IS NOT REPEATED
' ----------------------------------------------------------------------------
' modSimStats normalises because an iteration total may legally sit near Double
' maximum. RANKS CANNOT. A mid-rank lies in [1, n], and n is bounded by the
' accepted technical ceiling of 1048543 iterations, so the largest quantity
' formed below is a centred sum of squares under n^3 / 4 - about 2.9E17, which
' is seventeen orders below the Double maximum. Copying machinery that guards
' against nothing would only obscure the guards that are real.
'
' The CONTRIBUTIONS are unbounded and are checked for finiteness on the way in.
' Only ranks reach the arithmetic.
'
' ----------------------------------------------------------------------------
' WHY THE NO-TIES SHORTCUT IS ABSENT
' ----------------------------------------------------------------------------
' 1 - 6 * SUM(d^2) / (n * (n^2 - 1)) is not a faster route to the same number.
' On tied data it is a different and wrong number, and ties are the ordinary
' case here: a Risk at Probability 0.2 puts roughly 80% of its contribution
' column on one tied value. The shortcut appears nowhere, including as a fast
' path, and sim_contract.yaml forbids it.
'
' ----------------------------------------------------------------------------
' WHY THIS MODULE CARRIES ITS OWN SORT
' ----------------------------------------------------------------------------
' modSimStats has an accepted ascending sort, and it is Private. Making it
' Public would change the bytes of a module Run 6 executed and that the freeze
' pins by blob, which would cost the runtime authority of an accepted Phase-6
' module to save a few lines here. The sort below is a primitive, not a
' semantic: it is the same merge, on this module's own private copy.
' ==========================================================================

' The status of one driver's correlation. NOT a presentation string: the words
' a reader sees belong to the reporting layer, and a mathematical routine that
' returns worksheet vocabulary has taken a decision that is not its to take.
Public Const SIM_SENSITIVITY_DEFINED As Long = 0
Public Const SIM_SENSITIVITY_NO_VARIANCE As Long = 1

' One driver's finished result, ready to be ordered. `Rho` is meaningful only
' when `Status` is SIM_SENSITIVITY_DEFINED.
Public Type SimSensitivityResult
    PermanentId As String
    Rho As Double
    AbsRho As Double
    Status As Long
End Type

' ==========================================================================
' MID-RANKS - average ranks, 1-based, in the ORIGINAL order
'
' A tie block occupying ordinal positions p..q takes (p + q) / 2, so
' [10, 20, 20, 20, 50] ranks as [1, 3, 3, 3, 5].
'
' THE CALLER'S SEQUENCE IS NEVER REORDERED. A private sorted copy is taken and
' each original value is located in it by binary search, so observation j keeps
' its rank at position j. That is not tidiness: the correlation pairs
' contribution j with total j, and a rank vector that had been permuted would
' silently pair neither.
'
' TIES ARE EXACT EQUALITY, with no epsilon. Two contributions a single ulp apart
' are two different outcomes of the model, and grouping them would be this
' module inventing a numerical policy the project has not settled.
' ==========================================================================
Public Function SimSensitivityMidRanks(ByRef values() As Double, ByVal count As Long, _
                                       ByRef ranks() As Double, _
                                       ByRef detail As String) As Boolean
    Dim ordered() As Double
    Dim index As Long, low As Long, high As Long
    Dim first As Double, last As Double, total As Double, average As Double

    detail = vbNullString
    If Not SimSensitivityUsableSequence(values, count, "mid-rank", detail) Then Exit Function
    If Not SimSensitivitySortedCopy(values, count, ordered, detail) Then Exit Function

    ReDim ranks(0 To count - 1)
    For index = 0 To count - 1
        low = SimSensitivityLowerBound(ordered, count, values(LBound(values) + index))
        high = SimSensitivityUpperBound(ordered, count, values(LBound(values) + index))
        If low < 0 Or high < low Then
            detail = "sensitivity: a value is absent from its own sorted copy"
            Exit Function
        End If
        ' 1-based ordinals, and the midpoint of the block they span.
        first = CDbl(low + 1)
        last = CDbl(high + 1)
        If Not SafeAdd(first, last, total) Then
            detail = "sensitivity: a tie block midpoint is not representable"
            Exit Function
        End If
        If Not SafeDivide(total, 2#, average) Then
            detail = "sensitivity: a tie block midpoint is not representable"
            Exit Function
        End If
        ranks(index) = average
    Next index
    SimSensitivityMidRanks = True
End Function

' ==========================================================================
' PEARSON OVER TWO ALREADY-RANKED SERIES
'
' Taking ranks rather than raw observations is what lets the total's ranks be
' computed ONCE and reused for every driver. Ranking the total again per driver
' would be D sorts of the same vector for the same answer.
'
' ZERO VARIANCE IS UNDEFINED, NOT ZERO. A constant series has no dispersion to
' associate with anything, so there is no monotone relationship to find - and
' rho = 0 would assert that one was looked for and not found. The caller is told
' through Status, and `rho` is left at zero only because a Double must hold
' something; Status is what may be read.
' ==========================================================================
Public Function SimSensitivityRankCorrelation(ByRef driverRanks() As Double, _
                                              ByRef totalRanks() As Double, _
                                              ByVal count As Long, ByRef rho As Double, _
                                              ByRef status As Long, _
                                              ByRef detail As String) As Boolean
    Dim meanX As Double, meanY As Double
    Dim sxx As Double, syy As Double, sxy As Double
    Dim dx As Double, dy As Double, product As Double, denominator As Double
    Dim quotient As Double, index As Long

    detail = vbNullString
    rho = 0#
    status = SIM_SENSITIVITY_DEFINED
    If count < 2 Then
        detail = "sensitivity: a correlation needs at least two observations"
        Exit Function
    End If
    If Not SimSensitivityUsableSequence(driverRanks, count, "driver rank", detail) Then Exit Function
    If Not SimSensitivityUsableSequence(totalRanks, count, "total rank", detail) Then Exit Function

    If Not SimSensitivityRankMean(driverRanks, count, meanX, detail) Then Exit Function
    If Not SimSensitivityRankMean(totalRanks, count, meanY, detail) Then Exit Function

    sxx = 0#
    syy = 0#
    sxy = 0#
    For index = 0 To count - 1
        dx = driverRanks(LBound(driverRanks) + index) - meanX
        dy = totalRanks(LBound(totalRanks) + index) - meanY
        sxx = sxx + dx * dx
        syy = syy + dy * dy
        sxy = sxy + dx * dy
    Next index

    If sxx = 0# Or syy = 0# Then
        status = SIM_SENSITIVITY_NO_VARIANCE
        SimSensitivityRankCorrelation = True
        Exit Function
    End If

    If Not SafeMultiply(sxx, syy, product) Then
        detail = "sensitivity: the correlation denominator is not representable"
        Exit Function
    End If
    denominator = Sqr(product)
    If denominator = 0# Then
        status = SIM_SENSITIVITY_NO_VARIANCE
        SimSensitivityRankCorrelation = True
        Exit Function
    End If
    If Not SafeDivide(sxy, denominator, quotient) Then
        detail = "sensitivity: the correlation is not representable"
        Exit Function
    End If
    ' Rounding can carry a perfect monotone association a hair outside [-1, 1].
    ' Stating the bound is honest; publishing an impossible correlation is not.
    If quotient > 1# Then quotient = 1#
    If quotient < -1# Then quotient = -1#
    rho = quotient
    SimSensitivityRankCorrelation = True
End Function

' ==========================================================================
' ONE DRIVER'S SPEARMAN against a PRECOMPUTED total-rank vector
' ==========================================================================
Public Function SimSensitivitySpearman(ByRef driverValues() As Double, _
                                       ByRef totalRanks() As Double, _
                                       ByVal count As Long, ByRef rho As Double, _
                                       ByRef status As Long, _
                                       ByRef detail As String) As Boolean
    Dim driverRanks() As Double
    detail = vbNullString
    If Not SimSensitivityMidRanks(driverValues, count, driverRanks, detail) Then Exit Function
    If Not SimSensitivityRankCorrelation(driverRanks, totalRanks, count, rho, status, detail) Then
        Exit Function
    End If
    SimSensitivitySpearman = True
End Function

' ==========================================================================
' THE ORDER OF THE RANKED TABLE
'
' Returns INDICES into `results`, not a reordered copy, so the caller's sequence
' is untouched and every entry stays findable where it arrived.
'
' |rho| descending. A driver with no variance is not ranked at all - it has no
' rho to rank - and stays reportable through its Status.
'
' THE TIE-BREAK IS THE PERMANENT ID, ascending, compared as ordinal UTF-16 code
' units: the comparison this project already uses wherever driver order must be
' reproducible. Without it, two drivers at equal |rho| would emerge in whatever
' order the sort happened to produce, which is a different report from the same
' numbers. Worksheet row position cannot serve here - a driver's row moves when
' an unrelated driver is added, and identity that moves is not identity.
'
' NO TOP-N. Every eligible driver stays in the population; choosing how many to
' draw is a chart decision and belongs to the presentation phase.
' ==========================================================================
Public Function SimSensitivityRank(ByRef results() As SimSensitivityResult, _
                                   ByVal resultCount As Long, ByRef order() As Long, _
                                   ByRef eligibleCount As Long, _
                                   ByRef detail As String) As Boolean
    Dim eligible() As Long
    Dim index As Long, slot As Long, scan As Long, chosen As Long, candidate As Long

    detail = vbNullString
    eligibleCount = 0
    If resultCount < 0 Then
        detail = "sensitivity: a negative result count"
        Exit Function
    End If
    ReDim order(0 To 0)
    If resultCount = 0 Then
        SimSensitivityRank = True
        Exit Function
    End If

    ReDim eligible(0 To resultCount - 1)
    For index = 0 To resultCount - 1
        If results(LBound(results) + index).Status = SIM_SENSITIVITY_DEFINED Then
            eligible(eligibleCount) = index
            eligibleCount = eligibleCount + 1
        End If
    Next index
    If eligibleCount = 0 Then
        SimSensitivityRank = True
        Exit Function
    End If

    ' SELECTION, and it is total. Every comparison below decides on |rho| and
    ' then on the permanent id, so no two distinct drivers can compare equal and
    ' the result does not depend on the order they were supplied in.
    ReDim order(0 To eligibleCount - 1)
    For slot = 0 To eligibleCount - 1
        chosen = slot
        For scan = slot + 1 To eligibleCount - 1
            candidate = scan
            If SimSensitivityPrecedes(results, eligible(candidate), eligible(chosen)) Then
                chosen = candidate
            End If
        Next scan
        index = eligible(slot)
        eligible(slot) = eligible(chosen)
        eligible(chosen) = index
        order(slot) = eligible(slot)
    Next slot
    SimSensitivityRank = True
End Function

' ==========================================================================
' Private - ordering, the sorted copy, the bounds, and the checks
' ==========================================================================
Private Function SimSensitivityPrecedes(ByRef results() As SimSensitivityResult, _
                                        ByVal left As Long, ByVal right As Long) As Boolean
    Dim leftAbs As Double, rightAbs As Double
    leftAbs = results(LBound(results) + left).AbsRho
    rightAbs = results(LBound(results) + right).AbsRho
    If leftAbs > rightAbs Then
        SimSensitivityPrecedes = True
        Exit Function
    End If
    If leftAbs < rightAbs Then Exit Function
    ' Equal magnitude: the permanent id decides, ordinally and ascending.
    SimSensitivityPrecedes = SimSensitivityIdPrecedes( _
        results(LBound(results) + left).PermanentId, _
        results(LBound(results) + right).PermanentId)
End Function

Private Function SimSensitivityIdPrecedes(ByVal left As String, ByVal right As String) As Boolean
    ' ORDINAL UTF-16 CODE UNITS, one at a time. VBA's `<` on String is a TEXT
    ' comparison whose answer depends on the module's Option Compare and on the
    ' host locale, and an ordering that changes with the machine it runs on is
    ' not an ordering.
    Dim index As Long, shortest As Long
    Dim leftUnit As Long, rightUnit As Long
    shortest = Len(left)
    If Len(right) < shortest Then shortest = Len(right)
    For index = 1 To shortest
        leftUnit = AscW(Mid$(left, index, 1))
        rightUnit = AscW(Mid$(right, index, 1))
        ' AscW returns a SIGNED 16-bit value, so every code unit above U+7FFF
        ' arrives negative and would sort before every ASCII character. Adding
        ' the modulus restores the unsigned code unit the ordering is defined on.
        If leftUnit < 0 Then leftUnit = leftUnit + 65536
        If rightUnit < 0 Then rightUnit = rightUnit + 65536
        If leftUnit < rightUnit Then
            SimSensitivityIdPrecedes = True
            Exit Function
        End If
        If leftUnit > rightUnit Then Exit Function
    Next index
    SimSensitivityIdPrecedes = (Len(left) < Len(right))
End Function

Private Function SimSensitivityRankMean(ByRef ranks() As Double, ByVal count As Long, _
                                        ByRef result As Double, ByRef detail As String) As Boolean
    ' A PLAIN ACCUMULATION, and it is safe for the reason stated in the header:
    ' ranks lie in [1, n], so the sum cannot exceed n * (n + 1) / 2, about
    ' 5.5E11 at the accepted ceiling.
    Dim total As Double, index As Long
    total = 0#
    For index = 0 To count - 1
        total = total + ranks(LBound(ranks) + index)
    Next index
    If Not SafeDivide(total, CDbl(count), result) Then
        detail = "sensitivity: the mean rank is not representable"
        Exit Function
    End If
    SimSensitivityRankMean = True
End Function

Private Function SimSensitivityLowerBound(ByRef ordered() As Double, ByVal count As Long, _
                                          ByVal value As Double) As Long
    ' The FIRST position holding `value`, or -1. Exact equality throughout.
    '
    ' `Fix` is floor here because the loop only runs while low <= high, so
    ' (high - low) is never negative and truncation toward zero and toward
    ' minus infinity are the same operation. It is the project's accepted
    ' truncation primitive, and the midpoint is formed as low + half the span
    ' rather than (low + high) / 2 so no intermediate can overflow Long.
    Dim low As Long, high As Long, middle As Long, found As Long
    low = 0
    high = count - 1
    found = -1
    Do While low <= high
        middle = low + CLng(Fix((high - low) / 2))
        If ordered(middle) < value Then
            low = middle + 1
        ElseIf ordered(middle) > value Then
            high = middle - 1
        Else
            found = middle
            high = middle - 1
        End If
    Loop
    SimSensitivityLowerBound = found
End Function

Private Function SimSensitivityUpperBound(ByRef ordered() As Double, ByVal count As Long, _
                                          ByVal value As Double) As Long
    ' The LAST position holding `value`, or -1.
    Dim low As Long, high As Long, middle As Long, found As Long
    low = 0
    high = count - 1
    found = -1
    Do While low <= high
        middle = low + CLng(Fix((high - low) / 2))
        If ordered(middle) < value Then
            low = middle + 1
        ElseIf ordered(middle) > value Then
            high = middle - 1
        Else
            found = middle
            low = middle + 1
        End If
    Loop
    SimSensitivityUpperBound = found
End Function

Private Function SimSensitivitySortedCopy(ByRef values() As Double, ByVal count As Long, _
                                          ByRef ordered() As Double, _
                                          ByRef detail As String) As Boolean
    Dim index As Long
    If count < 1 Then
        detail = "sensitivity: an empty sequence cannot be ordered"
        Exit Function
    End If
    ReDim ordered(0 To count - 1)
    For index = 0 To count - 1
        ordered(index) = values(LBound(values) + index)
    Next index
    If Not SimSensitivitySortAscending(ordered, count, detail) Then Exit Function
    SimSensitivitySortedCopy = True
End Function

Private Function SimSensitivitySortAscending(ByRef series() As Double, ByVal count As Long, _
                                             ByRef detail As String) As Boolean
    ' Bottom-up merge. No recursion, and one scratch buffer for the whole sort.
    Dim scratch() As Double
    Dim runLength As Long, lowEnd As Long, midPoint As Long, highEnd As Long
    Dim fromLow As Long, fromHigh As Long, target As Long

    detail = vbNullString
    If count < 2 Then
        SimSensitivitySortAscending = True
        Exit Function
    End If
    ReDim scratch(0 To count - 1)

    runLength = 1
    Do While runLength < count
        lowEnd = 0
        Do While lowEnd < count
            midPoint = lowEnd + runLength
            If midPoint > count Then midPoint = count
            highEnd = midPoint + runLength
            If highEnd > count Then highEnd = count
            fromLow = lowEnd
            fromHigh = midPoint
            For target = lowEnd To highEnd - 1
                If fromLow < midPoint And _
                   (fromHigh >= highEnd Or Not (series(fromHigh) < series(fromLow))) Then
                    scratch(target) = series(fromLow)
                    fromLow = fromLow + 1
                Else
                    scratch(target) = series(fromHigh)
                    fromHigh = fromHigh + 1
                End If
            Next target
            lowEnd = lowEnd + 2 * runLength
        Loop
        For target = 0 To count - 1
            series(target) = scratch(target)
        Next target
        runLength = runLength * 2
    Loop
    SimSensitivitySortAscending = True
End Function

Private Function SimSensitivityUsableSequence(ByRef values() As Double, ByVal count As Long, _
                                              ByVal where As String, _
                                              ByRef detail As String) As Boolean
    ' The project's existing convention, not a new one: a NaN or an infinity is
    ' refused rather than ranked, because neither has a position in an order.
    Dim index As Long
    If count < 0 Then
        detail = "sensitivity: a negative observation count"
        Exit Function
    End If
    If count = 0 Then
        detail = "sensitivity: an empty sequence has no " & where
        Exit Function
    End If
    For index = 0 To count - 1
        If Not IsUsableDouble(values(LBound(values) + index)) Then
            detail = "sensitivity: the " & where & " sequence carries a value that is not a " & _
                     "finite Double"
            Exit Function
        End If
    Next index
    SimSensitivityUsableSequence = True
End Function
