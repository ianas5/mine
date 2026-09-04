Attribute VB_Name = "modSimAnnual"
Option Explicit

' ==========================================================================
' PCCM Phase 7 - the annual stochastic MATHEMATICS.
'
' Two authoritative outputs, and they are different objects:
'
'   THE ANNUAL DISTRIBUTIONS   a type-7 percentile ladder per project year,
'                              taken across iterations. It does NOT sum to the
'                              total percentile and it is NOT a profile.
'   THE SELECTED-Px PROFILE    the convex blend of the two annual vectors
'                              belonging to the SAME order statistics that
'                              produced the reported total Px.
'
' `annual_distributions.sums_to_total_percentile` is false and
' `annual_distributions.is_a_selected_px_profile` is false. Nothing here lets
' one stand in for the other: they are produced by different procedures from
' different inputs and neither is derived from the other.
'
' --------------------------------------------------------------------------
' WHAT THIS MODULE DOES NOT OWN
' --------------------------------------------------------------------------
' No RNG. No sampling. No stream assignment. No application state, run id,
' nonce or request identity. No workbook, worksheet, range or persistence. No
' presentation. No sensitivity. It is given numbers and returns numbers.
'
' Replay belongs to modSimEngine, which owns the generator and the contribution
' rule; the percentile belongs to modSimStats, which owns type 7. This module
' owns the ANNUAL arithmetic on top of both: the per-year regrouping of an
' accepted factor, the ladder assembly and the profile blend.
'
' IT DOES NOT BUILD AN ITERATION'S ANNUAL VECTOR. That is replay - the same
' sample, the same occurrence, a different factor - and replay belongs to
' modSimEngine, which owns the generator and the contribution rule. A second
' way to produce an annual vector would be a second contribution rule able to
' disagree with the accepted one.
'
' --------------------------------------------------------------------------
' IT RECOMPUTES NO INFLATION, NO FX AND NO DISCOUNT
' --------------------------------------------------------------------------
' `Knom` is `FX * SUM_y (w_y * infl_y)`. The per-year factor is the TERM of
' that sum, `FX * w_y * infl_y`, formed from the SAME resolved inputs Phase 5
' already carried into DriverFactors - its FxRate, Weights and Inflation. No
' rate is looked up here, no profile is resolved, no discount series is built.
'
' SUM_y Knom_y AND Knom ARE THE SAME REAL NUMBER AND MAY DIFFER AS DOUBLES,
' because the first applies FX once to a summed staging and the second applies
' it in every year, and floating-point multiplication does not distribute
' exactly over addition. The reconciliation identities are therefore checked to
' the project's own I3c/I4c allowance and never to bit equality - and nothing
' here scales, nudges or normalises a result to make a sum come out.
' ==========================================================================

' The year-block width. `retention.block_width_configurable` is true and the
' contract names no value, so this is the implementation's - NOT a business
' input, NOT a workbook cell, and not something a user can set. It changes how
' many years are held at once and changes no answer.
'
' One block covers every project shorter than 13 years. Beyond that the passes
' grow as ceil(duration / 12) while peak retention stays iterations * 12
' doubles.
Public Const SIM_ANNUAL_BLOCK_WIDTH As Long = 12

' ==========================================================================
' THE PER-YEAR FACTOR
' ==========================================================================
' Knom_y = FX * w_y * infl_y                     (withDiscount = False)
' Kpv_y  = FX * w_y * infl_y * disc_y            (withDiscount = True)
'
' The same factors, in the same order, through the same exact-product kernel
' modCalcFactors uses on its way to the scalar. What differs is only that the
' sum is not taken.
' IT TAKES ARRAYS, NOT A DriverFactors. This module names no Phase-5 type and
' knows nothing of drivers, registers or resolution: it is handed a rate and
' three series and returns their per-year products. The caller unpacks what it
' owns, which is also what keeps modSimAnnual a pure numerical module in the
' way modSimStats is one.
Public Function SimAnnualFactors(ByVal fxRate As Double, ByRef weights() As Double, _
                                 ByRef inflation() As Double, _
                                 ByRef discount() As Double, _
                                 ByVal withDiscount As Boolean, _
                                 ByRef result() As Double, _
                                 ByRef detail As String) As Boolean
    Dim years As Long, index As Long
    Dim group() As Double, groupWidth As Long
    Dim term As Double

    detail = vbNullString
    years = UBound(weights) - LBound(weights) + 1
    If years < 1 Then
        detail = "annual: no project year to decompose a factor over"
        Exit Function
    End If
    If UBound(inflation) - LBound(inflation) + 1 <> years Then
        detail = "annual: a weight for every project year but not an inflation factor"
        Exit Function
    End If
    If withDiscount Then
        If UBound(discount) - LBound(discount) + 1 < years Then
            detail = "annual: the discount series is shorter than the project years"
            Exit Function
        End If
        groupWidth = 4
    Else
        groupWidth = 3
    End If

    ReDim result(0 To years - 1)
    ReDim group(0 To groupWidth - 1)
    For index = 0 To years - 1
        group(0) = fxRate
        group(1) = weights(LBound(weights) + index)
        group(2) = inflation(LBound(inflation) + index)
        If withDiscount Then group(3) = discount(LBound(discount) + index)
        ' A SCALAR LOCAL, then the element. An array element handed straight
        ' to a ByRef out-parameter is a construct the accepted source
        ' transcriber cannot model, and a line only the compiler can read is a
        ' line no Linux test can check.
        If Not modCalcFactors.SafeProduct(group, groupWidth, term) Then
            detail = "annual: project year " & CStr(index + 1) & _
                     " factor is not representable"
            Exit Function
        End If
        result(index) = term
    Next index
    SimAnnualFactors = True
End Function

' ==========================================================================
' THE ANNUAL LADDER FOR ONE BLOCK OF YEARS
' ==========================================================================
' Given the block's iteration-by-year values - column-major, one column per
' year - produce that block's type-7 ladder, year by year.
'
' THE PERCENTILE IS modSimStats'. There is no second type-7 here: a year's
' ladder is SimStatsQuantileType7 over that year's column, which is an ordinary
' sequence of iteration values.
Public Function SimAnnualLadder(ByRef column() As Double, ByVal iterations As Long, _
                                ByVal yearCount As Long, ByRef probabilities() As Double, _
                                ByVal probabilityCount As Long, _
                                ByRef ladder() As Double, ByRef detail As String) As Boolean
    Dim year As Long, rung As Long, iteration As Long
    Dim series() As Double, value As Double

    detail = vbNullString
    If iterations < 1 Then
        detail = "annual: a run with no iterations has no annual distribution"
        Exit Function
    End If
    If probabilityCount < 1 Then
        detail = "annual: an empty ladder has no rungs"
        Exit Function
    End If

    ReDim ladder(0 To yearCount * probabilityCount - 1)
    ReDim series(0 To iterations - 1)
    For year = 0 To yearCount - 1
        For iteration = 0 To iterations - 1
            series(iteration) = column(iteration * yearCount + year)
        Next iteration
        For rung = 0 To probabilityCount - 1
            If Not modSimStats.SimStatsQuantileType7(series, iterations, _
                    probabilities(LBound(probabilities) + rung), value, detail) Then
                detail = "annual: project year " & CStr(year + 1) & ": " & detail
                Exit Function
            End If
            ladder(year * probabilityCount + rung) = value
        Next rung
    Next year
    SimAnnualLadder = True
End Function

' ==========================================================================
' THE SELECTED-Px ANNUAL PROFILE
' ==========================================================================
' Profile_Px(y) = (1 - f) * A_lo(y) + f * A_hi(y)
'
' `lo`, `hi` and `f` are the SAME type-7 position that produced the reported
' total Px, obtained from modSimStats.SimStatsQuantilePosition over the
' accepted iteration totals. This module chooses no iteration of its own.
'
' FORBIDDEN, AND ABSENT: nearest-rank substitution; picking whichever iteration
' is closest to Px; computing a percentile per year and calling that a profile;
' scaling a vector until it sums to Px. Two annual vectors are supplied and the
' blend of them is convex.
Public Function SimAnnualProfile(ByRef lowVector() As Double, _
                                 ByRef highVector() As Double, _
                                 ByVal yearCount As Long, ByVal fraction As Double, _
                                 ByRef result() As Double, _
                                 ByRef detail As String) As Boolean
    Dim year As Long
    Dim low As Double, high As Double, blended As Double

    detail = vbNullString
    If yearCount < 1 Then
        detail = "annual: a profile needs at least one project year"
        Exit Function
    End If
    If fraction < 0# Or fraction > 1# Then
        detail = "annual: the interpolation fraction is outside [0, 1]"
        Exit Function
    End If

    ReDim result(0 To yearCount - 1)
    For year = 0 To yearCount - 1
        low = lowVector(LBound(lowVector) + year)
        If fraction = 0# Then
            ' AN INTEGRAL POSITION IS AN ORDER STATISTIC OUTRIGHT. Returning the
            ' low vector untouched rather than forming 1 * low + 0 * high is the
            ' rule SimStatsQuantileSorted applies to the value, and it is what
            ' makes the profile exactly one iteration's own annual vector.
            result(year) = low
        Else
            high = highVector(LBound(highVector) + year)
            If low = high Then
                ' THE CONSTANT-BRACKET INVARIANT. A convex combination of two
                ' equal numbers IS that number, and 0.7 * 0.1 + 0.3 * 0.1 is
                ' 0.10000000000000002.
                result(year) = low
            Else
                blended = (1# - fraction) * low + fraction * high
                If Not modCalcFactors.IsUsableDouble(blended) Then
                    detail = "annual: the profile blend for project year " & _
                             CStr(year + 1) & " is not representable"
                    Exit Function
                End If
                result(year) = blended
            End If
        End If
    Next year
    SimAnnualProfile = True
End Function

' ==========================================================================
' HOW MANY BLOCKS A DURATION NEEDS
'
' ceil(duration / width), computed without `/` on Longs so no rounding rule is
' inherited from a division.
' ==========================================================================
Public Function SimAnnualBlockCount(ByVal duration As Long, ByVal width As Long) As Long
    Dim blocks As Long
    If duration < 1 Or width < 1 Then
        SimAnnualBlockCount = 0
        Exit Function
    End If
    ' `Fix` of a Double division, not the `\` operator: the accepted source
    ' transcriber does not model `\`, and a construct only the compiler can
    ' read is a construct no Linux test can check.
    blocks = CLng(Fix(CDbl(duration) / CDbl(width)))
    If blocks * width < duration Then blocks = blocks + 1
    SimAnnualBlockCount = blocks
End Function
