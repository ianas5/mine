Attribute VB_Name = "modSimStats"
Option Explicit

' ==========================================================================
' PCCM Phase 6 - simulation statistics. WORKSHEET-FREE BY CONSTRUCTION.
'
' Sort, moments, Hyndman-Fan type-7 quantiles, the full accepted ladder and
' selected-Px contingency. All scale-safe. That is the whole of this module.
'
' builder/pccm_builder/sim_stats.py is the single definition of these semantics;
' this module is their VBA implementation.
'
' IT KNOWS NOTHING ELSE. No RNG, no sampler, no iteration engine, no driver, no
' request fingerprint, no result digest, no run_id, no _SimData, no Results, no
' run state and no workbook. It receives a finished sequence of Doubles and
' returns statistics over it. That separation is why a wrong statistic can be
' substituted here without touching the engine, and why the engine cannot
' quietly acquire a statistic of its own.
'
' THE FORBIDDEN WORD. `Percentile` is globally forbidden in PCCM VBA and Step 9
' takes no exception to that. The executable vocabulary here is QUANTILE
' throughout, and there is no WorksheetFunction call of any kind.
'
' ------------------------------------------------------------------------
' WHY EVERY MOMENT IS SCALE-NORMALISED
' ------------------------------------------------------------------------
' The accepted numerical domain is not narrowed by Phase 6: an iteration total
' may legally sit near Double maximum, and totals of opposite sign near that
' magnitude are legal together. Their MEAN is then perfectly representable while
' their SUM is not, and their standard deviation is representable while both a
' naive sum of squares and an unguarded x - mean are not.
'
' So the sums are formed in a normalised space and rescaled exactly once:
'
'     unitScale = the largest power of two NOT EXCEEDING max(|x|)
'     scaled    = x / unitScale                       (EXACT)
'     ...       accumulate, average, deviate ...
'     result    = (normalised result) * unitScale
'
' A POWER-OF-TWO SCALE IS NOT A DETAIL. Dividing a Double by a power of two only
' adjusts the exponent, so the quotient is exact for every value that stays in
' range. Dividing by max(|x|) itself - the obvious choice - rounds every element
' and spends up to an ulp per value before any statistic exists. Both are "scale
' aware"; only one costs nothing. And the power BELOW the magnitude is taken
' rather than the one above it, because the one above a value near Double
' maximum is 2^1024, which does not exist.
'
' SafeSignedSum still performs the accumulation, so canonical left-to-right
' order and its exact-rescue tier are unchanged; normalisation decides only the
' SPACE the accepted primitive works in, never the order it works in.
'
' ------------------------------------------------------------------------
' THE CONSTANT-SAMPLE INVARIANT
' ------------------------------------------------------------------------
' Left-to-right drift is acceptable on a sample that genuinely varies. It is NOT
' acceptable on one that does not.
'
' If every retained observation is the same Double then the distribution has no
' dispersion at all: the mean IS that Double and the sample standard deviation
' IS exactly zero. Rediscovering that by accumulating the value a thousand times
' produces 1.4999999999999677E308 where the answer is 1.5E308, and a standard
' deviation of 3.2E294 where the answer is 0 - PCCM would be reporting
' stochastic dispersion for a distribution that has none.
'
' So each moment tests for a constant sample AFTER validation and BEFORE any
' accumulation. The quantile ladder carries the same invariant in its
' equal-bracket rule. NO NON-DEGENERATE STATISTIC MOVES: the shortcut can only
' fire where the answer is already known exactly.
'
' ------------------------------------------------------------------------
' ORDER
' ------------------------------------------------------------------------
' THE CALLER'S RETAINED ARRAY IS NEVER REORDERED. Quantiles work on a private
' sorted COPY, made once per description; the mean and the sample deviation work
' on the ORIGINAL order, because SafeSignedSum accumulates left to right and
' handing it a reordered sequence would move the last bits of a valid statistic.
'
' Step 8's retained arrays are zero-based with element 0 holding iteration 1.
' Nothing here creates or reinterprets iteration identity: every bound is read
' through LBound and a logical count, and Step 10 will encode the iteration
' index against the original retained order.
' ==========================================================================

' Every statistic of one measure - nominal or PV - over one run.
'
' DERIVED REPORTING OUTPUT ONLY. It carries no retained sample, no RNG state and
' no simulation authority, and nothing downstream may treat it as an input that
' could re-authorise a run. The ladder travels beside it in two parallel arrays
' rather than inside it, so no fixed bound here can drift from
' SIM_QUANTILE_COUNT.
Public Type SimStatsMeasure
    Count As Long
    Mean As Double
    SampleStandardDeviation As Double
    Minimum As Double
    Maximum As Double
    QuantileCount As Long
    Described As Boolean
End Type

' ==========================================================================
' The sample mean
'
' `sum(x) / n` is not implemented. [-1.7E308, 1.7E308, 1.7E308, 1.7E308] has a
' mean of 8.5E307 and a running sum of 3.4E308, which does not exist. Refusing
' that model would be refusing an answer it has.
'
' The accumulation uses the ORIGINAL order it was given.
' ==========================================================================
Public Function SimStatsMean(ByRef values() As Double, ByVal count As Long, _
                             ByRef result As Double, ByRef detail As String) As Boolean
    Dim scaled() As Double
    Dim unitScale As Double, total As Double, normalised As Double, candidate As Double
    Dim constantValue As Double, index As Long

    detail = vbNullString
    If count < 1 Then
        detail = "statistics: the mean of an empty sequence does not exist"
        Exit Function
    End If
    If Not SimStatsUsableSequence(values, count, "mean", detail) Then Exit Function

    ' EXACT, and checked before any accumulation.
    If SimStatsConstantValue(values, count, constantValue) Then
        result = constantValue
        SimStatsMean = True
        Exit Function
    End If

    If Not SimStatsUnitScale(values, count, unitScale, detail) Then Exit Function
    If unitScale = 0# Then
        ' An all-zero sample is constant and was answered above; this is the
        ' guard, not a path.
        detail = "statistics: a varying sample produced no scale"
        Exit Function
    End If

    ReDim scaled(0 To count - 1)
    For index = 0 To count - 1
        scaled(index) = values(LBound(values) + index) / unitScale
    Next index

    If Not SafeSignedSum(scaled, count, total) Then
        detail = "statistics: the normalised accumulation for the mean is not representable"
        Exit Function
    End If
    If Not SafeDivide(total, CDbl(count), normalised) Then
        detail = "statistics: the normalised mean is not representable"
        Exit Function
    End If
    If Not SafeMultiply(normalised, unitScale, candidate) Then
        detail = "statistics: the mean rescale is not representable"
        Exit Function
    End If

    result = candidate
    SimStatsMean = True
End Function

' ==========================================================================
' The SAMPLE standard deviation - divisor n - 1, never n
'
' Two passes, both in the normalised space. NEITHER FORBIDDEN PATH IS TAKEN. A
' naive SUM x^2 overflows for any total beyond about 1.3E154 and reports
' infinity for a spread that exists. An unguarded Welford delta = x - mean
' overflows too: for [-1.7E308, 1.7E308, 1.7E308, 1.7E308] the mean is 8.5E307
' and the first deviation is -2.55E308, so the recurrence fails on its first
' step. In the normalised space every value lies in [-2, 2] and every deviation
' in [-4, 4], so no intermediate can leave Double range at all.
'
' n < 2 REFUSES rather than inventing a value: the divisor is n - 1, which does
' not exist there, and the deviation of one observation is undefined, not zero.
'
' A VARYING SAMPLE WHOSE TRUE DEVIATION HAS NO DOUBLE ALSO REFUSES. Returning
' zero would state that a sample which demonstrably varies has no dispersion.
' ==========================================================================
Public Function SimStatsSampleStandardDeviation(ByRef values() As Double, ByVal count As Long, _
                                                ByRef result As Double, _
                                                ByRef detail As String) As Boolean
    Dim scaled() As Double, squares() As Double
    Dim unitScale As Double, total As Double, centre As Double
    Dim residual As Double, variance As Double, deviation As Double, candidate As Double
    Dim constantValue As Double, index As Long

    detail = vbNullString
    If count < 2 Then
        detail = "statistics: a sample standard deviation needs at least two observations"
        Exit Function
    End If
    If Not SimStatsUsableSequence(values, count, "sample standard deviation", detail) Then Exit Function

    ' A sample with one distinct value has no dispersion. EXACTLY +0.
    If SimStatsConstantValue(values, count, constantValue) Then
        result = 0#
        SimStatsSampleStandardDeviation = True
        Exit Function
    End If

    If Not SimStatsUnitScale(values, count, unitScale, detail) Then Exit Function
    If unitScale = 0# Then
        detail = "statistics: a varying sample produced no scale"
        Exit Function
    End If

    ReDim scaled(0 To count - 1)
    For index = 0 To count - 1
        scaled(index) = values(LBound(values) + index) / unitScale
    Next index

    If Not SafeSignedSum(scaled, count, total) Then
        detail = "statistics: the normalised accumulation for the deviation is not representable"
        Exit Function
    End If
    If Not SafeDivide(total, CDbl(count), centre) Then
        detail = "statistics: the normalised centre is not representable"
        Exit Function
    End If

    ReDim squares(0 To count - 1)
    For index = 0 To count - 1
        deviation = scaled(index) - centre
        squares(index) = deviation * deviation
    Next index

    If Not SafeSignedSum(squares, count, residual) Then
        detail = "statistics: the normalised sum of squared deviations is not representable"
        Exit Function
    End If
    If Not SafeDivide(residual, CDbl(count - 1), variance) Then
        detail = "statistics: the normalised sample variance is not representable"
        Exit Function
    End If
    If variance < 0# Then
        detail = "statistics: the normalised sample variance is negative"
        Exit Function
    End If

    If Not SafeMultiply(Sqr(variance), unitScale, candidate) Then
        detail = "statistics: the sample standard deviation rescale is not representable"
        Exit Function
    End If

    result = candidate
    SimStatsSampleStandardDeviation = True
End Function

' ==========================================================================
' Hyndman-Fan type 7 - the standalone entry point
'
' Sorts its own private copy, once. The caller's sequence is never reordered.
' ==========================================================================
Public Function SimStatsQuantileType7(ByRef values() As Double, ByVal count As Long, _
                                      ByVal p As Double, ByRef result As Double, _
                                      ByRef detail As String) As Boolean
    Dim ordered() As Double
    Dim candidate As Double

    detail = vbNullString
    If count < 1 Then
        detail = "statistics: an empty sequence has no quantile"
        Exit Function
    End If
    If Not SimStatsUsableProbability(p, detail) Then Exit Function
    If Not SimStatsUsableSequence(values, count, "quantile", detail) Then Exit Function
    If Not SimStatsSortedCopy(values, count, ordered, detail) Then Exit Function
    If Not SimStatsQuantileSorted(ordered, count, p, candidate, detail) Then Exit Function

    result = candidate
    SimStatsQuantileType7 = True
End Function

' ==========================================================================
' Every statistic of one measure, SORTING EXACTLY ONCE
'
' The ordered copy is formed here and reused for all eleven ladder values; the
' standalone quantile entry point is deliberately NOT called in the loop,
' because it would sort again for each label. The mean and the sample deviation
' are computed over the ORIGINAL order.
'
' TRANSACTIONAL. The summary, the labels and the ladder are built into locals
' and committed together. A ladder whose tenth value succeeded and whose
' eleventh refused is never published.
' ==========================================================================
Public Function SimStatsDescribe(ByRef values() As Double, ByVal count As Long, _
                                 ByRef summary As SimStatsMeasure, _
                                 ByRef quantileLabels() As String, _
                                 ByRef quantileValues() As Double, _
                                 ByRef detail As String) As Boolean
    Dim ordered() As Double
    Dim labels() As String, ladder() As Double
    Dim candidate As SimStatsMeasure
    Dim label As String
    Dim index As Long
    Dim p As Double, measured As Double

    detail = vbNullString
    ' Enough observations for EVERY statistic it reports, so the strictest of
    ' them decides: the sample deviation needs two.
    If count < 2 Then
        detail = "statistics: a description needs at least two observations"
        Exit Function
    End If
    If Not SimStatsUsableSequence(values, count, "statistics", detail) Then Exit Function

    If Not SimStatsSortedCopy(values, count, ordered, detail) Then Exit Function

    ReDim labels(0 To SIM_QUANTILE_COUNT - 1)
    ReDim ladder(0 To SIM_QUANTILE_COUNT - 1)
    For index = 0 To SIM_QUANTILE_COUNT - 1
        If Not SimStatsLadderLabel(index, label, detail) Then Exit Function
        If Not SimStatsProbabilityOf(label, p, detail) Then Exit Function
        If Not SimStatsQuantileSorted(ordered, count, p, measured, detail) Then Exit Function
        labels(index) = label
        ladder(index) = measured
    Next index

    ' ORIGINAL ORDER, not the sorted copy that happens to be lying there.
    If Not SimStatsMean(values, count, measured, detail) Then Exit Function
    candidate.Mean = measured
    If Not SimStatsSampleStandardDeviation(values, count, measured, detail) Then Exit Function
    candidate.SampleStandardDeviation = measured

    candidate.Count = count
    candidate.Minimum = ordered(LBound(ordered))
    candidate.Maximum = ordered(LBound(ordered) + count - 1)
    candidate.QuantileCount = SIM_QUANTILE_COUNT
    candidate.Described = True

    summary = candidate
    quantileLabels = labels
    quantileValues = ladder
    SimStatsDescribe = True
End Function

' ==========================================================================
' The selected confidence level - A LOOKUP, NOT A RECOMPUTATION
'
' Changing Selected CL is reporting-only: no simulation, no RNG, no draw, no
' re-sort and no second quantile engine. The headline levels in modSimContract
' are presentation identity and are read out of this same ladder.
'
' THE FIXED LEVEL IS NOT SELECTABLE. It is computed, stored and reported like
' every other rung, and refused as a selector.
' ==========================================================================
Public Function SimStatsSelectedQuantile(ByRef quantileLabels() As String, _
                                         ByRef quantileValues() As Double, _
                                         ByVal quantileCount As Long, _
                                         ByVal selectedLabel As String, _
                                         ByRef result As Double, _
                                         ByRef detail As String) As Boolean
    Dim index As Long, found As Long

    detail = vbNullString
    If quantileCount <> SIM_QUANTILE_COUNT Then
        detail = "statistics: the ladder is not the accepted length"
        Exit Function
    End If
    If Len(selectedLabel) = 0 Then
        detail = "statistics: a blank confidence level"
        Exit Function
    End If
    If StrComp(selectedLabel, SIM_QUANTILE_FIXED_1, vbBinaryCompare) = 0 Then
        detail = "statistics: " & SIM_QUANTILE_FIXED_1 & " is reported and fixed; it is not selectable"
        Exit Function
    End If

    found = -1
    For index = 0 To quantileCount - 1
        If StrComp(quantileLabels(LBound(quantileLabels) + index), selectedLabel, _
                   vbBinaryCompare) = 0 Then
            If found >= 0 Then
                detail = "statistics: the ladder carries a duplicate label"
                Exit Function
            End If
            found = index
        End If
    Next index
    If found < 0 Then
        detail = "statistics: an unknown confidence level"
        Exit Function
    End If

    result = quantileValues(LBound(quantileValues) + found)
    SimStatsSelectedQuantile = True
End Function

' ==========================================================================
' Contingency
'
'     contingency = selected Px total - deterministic base estimate A
'
' The baseline is the Phase-5 DETERMINISTIC BASE ESTIMATE A and nothing else -
' not the simulation mean, not the analytical expected total, not A + EMV. It
' arrives as an explicit scalar; this module derives it from nothing.
'
' A NEGATIVE CONTINGENCY IS LEGAL and is preserved. Clamping it to zero would
' hide a selected level below the deterministic base, which is a real and
' reportable outcome.
' ==========================================================================
Public Function SimStatsContingency(ByVal selectedTotal As Double, _
                                    ByVal baseEstimateA As Double, _
                                    ByRef result As Double, ByRef detail As String) As Boolean
    Dim candidate As Double

    detail = vbNullString
    If Not IsUsableDouble(selectedTotal) Then
        detail = "contingency: the selected total is not a finite Double"
        Exit Function
    End If
    If Not IsUsableDouble(baseEstimateA) Then
        detail = "contingency: the deterministic base estimate A is not a finite Double"
        Exit Function
    End If
    If Not SafeSubtract(selectedTotal, baseEstimateA, candidate) Then
        detail = "contingency: the selected total minus the deterministic base estimate A " & _
                 "is not representable"
        Exit Function
    End If

    result = candidate
    SimStatsContingency = True
End Function

' ==========================================================================
' Shared predicates and machinery
' ==========================================================================

' Every semantically consumed element must be a usable Double. THE COUNT IS
' CHECKED FIRST: an empty sequence is refused before any bound of the carrier is
' read, because a zero-count caller may hand over an array that was never sized.
Private Function SimStatsUsableSequence(ByRef values() As Double, ByVal count As Long, _
                                        ByVal where As String, ByRef detail As String) As Boolean
    Dim index As Long
    If count < 0 Then
        detail = "statistics: a negative observation count"
        Exit Function
    End If
    If count = 0 Then
        detail = "statistics: an empty sequence has no " & where
        Exit Function
    End If
    For index = 0 To count - 1
        If Not IsUsableDouble(values(LBound(values) + index)) Then
            detail = "statistics: the " & where & " sequence carries a value that is not a " & _
                     "finite Double"
            Exit Function
        End If
    Next index
    SimStatsUsableSequence = True
End Function

Private Function SimStatsUsableProbability(ByVal p As Double, ByRef detail As String) As Boolean
    If Not IsUsableDouble(p) Then
        detail = "statistics: the quantile probability is not a finite Double"
        Exit Function
    End If
    If p < 0# Or p > 1# Then
        detail = "statistics: the quantile probability is outside [0, 1]"
        Exit Function
    End If
    SimStatsUsableProbability = True
End Function

' True when every observation is the SAME Double. The test is on the Double, not
' on an approximate closeness: two values differing in the last bit are a real
' (if tiny) dispersion and must not be flattened.
Private Function SimStatsConstantValue(ByRef values() As Double, ByVal count As Long, _
                                       ByRef constantValue As Double) As Boolean
    Dim first As Double, index As Long
    first = values(LBound(values))
    For index = 1 To count - 1
        If values(LBound(values) + index) <> first Then Exit Function
    Next index
    constantValue = first
    SimStatsConstantValue = True
End Function

' The largest power of two NOT EXCEEDING max(|x|), or 0 for an all-zero sample.
'
' Built by exact halving and doubling rather than 2 ^ exponent: the power above a
' value near Double maximum is 2^1024, which raises before the scale exists. The
' doubling test is written as `candidate <= largest / 2` for the same reason -
' `candidate * 2 <= largest` overflows on its last comparison. Halving is exact
' all the way into the subnormal range, so the minimum positive Double gets
' itself as its scale.
Private Function SimStatsUnitScale(ByRef values() As Double, ByVal count As Long, _
                                   ByRef unitScale As Double, ByRef detail As String) As Boolean
    Dim index As Long, magnitude As Double, largest As Double, candidate As Double
    largest = 0#
    For index = 0 To count - 1
        magnitude = Abs(values(LBound(values) + index))
        If magnitude > largest Then largest = magnitude
    Next index
    If largest = 0# Then
        unitScale = 0#
        SimStatsUnitScale = True
        Exit Function
    End If

    candidate = 1#
    If candidate > largest Then
        Do While candidate > largest
            candidate = candidate / 2#
        Loop
    Else
        Do While candidate <= largest / 2#
            candidate = candidate * 2#
        Loop
    End If

    If Not IsUsableDouble(candidate) Then
        detail = "statistics: the normalisation scale is not a finite Double"
        Exit Function
    End If
    If candidate <= 0# Then
        detail = "statistics: the normalisation scale is not positive"
        Exit Function
    End If
    If candidate > largest Then
        detail = "statistics: the normalisation scale exceeds the sample magnitude"
        Exit Function
    End If

    unitScale = candidate
    SimStatsUnitScale = True
End Function

' A private copy, sorted once. THE CALLER'S ARRAY IS NOT TOUCHED.
Private Function SimStatsSortedCopy(ByRef values() As Double, ByVal count As Long, _
                                    ByRef ordered() As Double, ByRef detail As String) As Boolean
    Dim index As Long
    If count < 1 Then
        detail = "statistics: an empty sequence cannot be ordered"
        Exit Function
    End If
    ReDim ordered(0 To count - 1)
    For index = 0 To count - 1
        ordered(index) = values(LBound(values) + index)
    Next index
    If Not SimStatsSortAscending(ordered, count, detail) Then Exit Function
    SimStatsSortedCopy = True
End Function

' ==========================================================================
' BOTTOM-UP STABLE MERGE SORT - O(n log n)
'
' The design target is 100,000 retained totals. An insertion or bubble sort is
' O(n^2) and would take about five billion comparisons there, so neither is
' used. Bottom-up merging needs no recursion, allocates its scratch buffer once
' before any merging begins, and never allocates inside the elemental loop.
'
' THE TIE RULE IS `<=`, so the left run wins and the order of equal Doubles is
' the order they arrived in. Nothing exposed depends on which of two equal
' Doubles is returned - they are equal - but a deterministic rule is what makes
' two runs of the same data agree bit for bit without argument.
'
' No library, no COM, no worksheet sort.
' ==========================================================================
Private Function SimStatsSortAscending(ByRef series() As Double, ByVal count As Long, _
                                       ByRef detail As String) As Boolean
    Dim scratch() As Double
    Dim runLength As Long, lowEnd As Long, midPoint As Long, highEnd As Long
    Dim fromLow As Long, fromHigh As Long, target As Long

    If count < 2 Then
        SimStatsSortAscending = True
        Exit Function
    End If
    ReDim scratch(0 To count - 1)

    runLength = 1
    Do While runLength < count
        lowEnd = 0
        Do While lowEnd < count
            midPoint = lowEnd + runLength
            If midPoint > count Then midPoint = count
            highEnd = lowEnd + 2 * runLength
            If highEnd > count Then highEnd = count
            If midPoint < highEnd Then
                fromLow = lowEnd
                fromHigh = midPoint
                target = lowEnd
                Do While target < highEnd
                    If fromLow >= midPoint Then
                        scratch(target) = series(fromHigh)
                        fromHigh = fromHigh + 1
                    ElseIf fromHigh >= highEnd Then
                        scratch(target) = series(fromLow)
                        fromLow = fromLow + 1
                    ElseIf series(fromLow) <= series(fromHigh) Then
                        scratch(target) = series(fromLow)
                        fromLow = fromLow + 1
                    Else
                        scratch(target) = series(fromHigh)
                        fromHigh = fromHigh + 1
                    End If
                    target = target + 1
                Loop
                For target = lowEnd To highEnd - 1
                    series(target) = scratch(target)
                Next target
            End If
            lowEnd = lowEnd + 2 * runLength
        Loop
        runLength = runLength * 2
    Loop

    SimStatsSortAscending = True
End Function

' ==========================================================================
' Type 7 over an ALREADY SORTED sequence. SORTS NOTHING.
'
'     h  = (n - 1) * p
'     lo = floor(h)
'     hi = min(lo + 1, n - 1)
'     f  = h - lo
'     Px = (1 - f) * x[lo] + f * x[hi]
'
' The positions are zero-relative in the sorted logical sample, independent of
' the physical LBound of the carrier.
'
' `Fix` is `floor` here and only here: h is (n - 1) * p with n >= 1 and p in
' [0, 1], so h is never negative and the two agree on every value this can see.
'
' THE CONVEX FORM IS THE POINT. x[lo] + f * (x[hi] - x[lo]) is the same number in
' exact arithmetic and a different one in Double: between -1.7E308 and 1.7E308
' the difference is 3.4E308, which does not exist, while every convex
' combination of two representable endpoints is bracketed by them and therefore
' always exists.
' ==========================================================================
Private Function SimStatsQuantileSorted(ByRef ordered() As Double, ByVal count As Long, _
                                        ByVal p As Double, ByRef result As Double, _
                                        ByRef detail As String) As Boolean
    Dim h As Double, fraction As Double
    Dim lowIndex As Long, highIndex As Long
    Dim low As Double, high As Double, candidate As Double

    h = CDbl(count - 1) * p
    lowIndex = CLng(Fix(h))
    If lowIndex < 0 Then lowIndex = 0
    If lowIndex > count - 1 Then lowIndex = count - 1
    highIndex = lowIndex + 1
    If highIndex > count - 1 Then highIndex = count - 1
    fraction = h - CDbl(lowIndex)
    low = ordered(LBound(ordered) + lowIndex)
    high = ordered(LBound(ordered) + highIndex)

    If fraction = 0# Then
        ' An integral h selects an order statistic outright. Returning it
        ' untouched rather than forming 1 * low + 0 * high keeps p = 0 and p = 1
        ' exact at every magnitude, including subnormals.
        candidate = low
    ElseIf low = high Then
        ' THE CONSTANT-BRACKET INVARIANT. A convex combination of two equal
        ' numbers IS that number, but 0.7 * 0.1 + 0.3 * 0.1 is
        ' 0.10000000000000002. A run with no dispersion cannot report a ladder
        ' that creeps.
        candidate = low
    Else
        candidate = (1# - fraction) * low + fraction * high
    End If

    If Not IsUsableDouble(candidate) Then
        detail = "statistics: the quantile interpolation is not representable"
        Exit Function
    End If
    result = candidate
    SimStatsQuantileSorted = True
End Function

' The accepted ladder, in the accepted order, read from the projection. There is
' no second list here: a copy would be a third authority able to drift from both
' sim_contract.yaml and input_contract.yaml.
Private Function SimStatsLadderLabel(ByVal position As Long, ByRef label As String, _
                                     ByRef detail As String) As Boolean
    If position = 0 Then
        label = SIM_QUANTILE_1
    ElseIf position = 1 Then
        label = SIM_QUANTILE_2
    ElseIf position = 2 Then
        label = SIM_QUANTILE_3
    ElseIf position = 3 Then
        label = SIM_QUANTILE_4
    ElseIf position = 4 Then
        label = SIM_QUANTILE_5
    ElseIf position = 5 Then
        label = SIM_QUANTILE_6
    ElseIf position = 6 Then
        label = SIM_QUANTILE_7
    ElseIf position = 7 Then
        label = SIM_QUANTILE_8
    ElseIf position = 8 Then
        label = SIM_QUANTILE_9
    ElseIf position = 9 Then
        label = SIM_QUANTILE_10
    ElseIf position = 10 Then
        label = SIM_QUANTILE_11
    Else
        detail = "statistics: the ladder has no rung at that position"
        Exit Function
    End If
    SimStatsLadderLabel = True
End Function

' `P<number>` decoded numerically, by its own accepted form.
'
' THIS IS NOT THE PERMANENT-ID RULE. Those identifiers are ordinal strings and
' their numeric suffix is explicitly never read; a ladder label is DEFINED by
' the number it spells, so it is read as one - digit by digit, with no Val, no
' locale and no library.
Private Function SimStatsProbabilityOf(ByVal label As String, ByRef p As Double, _
                                       ByRef detail As String) As Boolean
    Dim index As Long, code As Long, digits As Long, magnitude As Double
    If Len(label) < 2 Then
        detail = "statistics: an unreadable quantile label"
        Exit Function
    End If
    If Mid(label, 1, 1) <> "P" Then
        detail = "statistics: a quantile label must begin with P"
        Exit Function
    End If
    magnitude = 0#
    digits = 0
    For index = 2 To Len(label)
        code = Asc(Mid(label, index, 1))
        If code < 48 Or code > 57 Then
            detail = "statistics: a quantile label carries a character that is not a digit"
            Exit Function
        End If
        magnitude = magnitude * 10# + CDbl(code - 48)
        digits = digits + 1
    Next index
    If digits = 0 Then
        detail = "statistics: a quantile label carries no digits"
        Exit Function
    End If
    If magnitude < 0# Or magnitude > 100# Then
        detail = "statistics: a quantile label is outside [P0, P100]"
        Exit Function
    End If
    p = magnitude / 100#
    SimStatsProbabilityOf = True
End Function
