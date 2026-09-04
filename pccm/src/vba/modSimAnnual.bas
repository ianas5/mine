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
' accepted factor, the block accumulation, the ladder assembly and the profile
' blend.
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
Public Function SimAnnualFactors(ByRef driver As DriverFactors, _
                                 ByRef discount() As Double, _
                                 ByVal withDiscount As Boolean, _
                                 ByRef result() As Double, _
                                 ByRef detail As String) As Boolean
    Dim years As Long, index As Long
    Dim group() As Double, groupWidth As Long

    detail = vbNullString
    years = UBound(driver.Weights) - LBound(driver.Weights) + 1
    If years < 1 Then
        detail = "annual: driver " & driver.PermanentId & " carries no project year"
        Exit Function
    End If
    If UBound(driver.Inflation) - LBound(driver.Inflation) + 1 <> years Then
        detail = "annual: driver " & driver.PermanentId & " has a weight for every " & _
                 "project year but not an inflation factor"
        Exit Function
    End If
    If withDiscount Then
        If UBound(discount) - LBound(discount) + 1 < years Then
            detail = "annual: the discount series is shorter than driver " & _
                     driver.PermanentId & "'s project years"
            Exit Function
        End If
        groupWidth = 4
    Else
        groupWidth = 3
    End If

    ReDim result(0 To years - 1)
    ReDim group(0 To groupWidth - 1)
    For index = 0 To years - 1
        group(0) = driver.FxRate
        group(1) = driver.Weights(LBound(driver.Weights) + index)
        group(2) = driver.Inflation(LBound(driver.Inflation) + index)
        If withDiscount Then group(3) = discount(LBound(discount) + index)
        If Not modCalcFactors.SafeProduct(group, groupWidth, result(index)) Then
            detail = "annual: driver " & driver.PermanentId & " project year " & _
                     CStr(index + 1) & " factor is not representable"
            Exit Function
        End If
    Next index
    SimAnnualFactors = True
End Function

' ==========================================================================
' ONE ITERATION'S ANNUAL VECTOR, over the years of one block
' ==========================================================================
' A_j(y) = SUM_d observation_d_j * K_d_y
'
' THE OBSERVATION IS AN INPUT. It is the number the accepted run already
' produced for driver d in iteration j - unit cost times quantity for a Cost
' Line, severity for a Risk that occurred, and exactly zero for one that did
' not. This module cannot sample and holds no generator; what it does is apply
' a different deployment factor to a number that already exists.
Public Function SimAnnualVector(ByRef observations() As Double, _
                                ByRef factors() As Double, _
                                ByVal driverCount As Long, ByVal yearCount As Long, _
                                ByVal firstYear As Long, ByRef result() As Double, _
                                ByRef detail As String) As Boolean
    Dim year As Long, driver As Long
    Dim terms() As Double, observation As Double, term As Double

    detail = vbNullString
    If driverCount < 1 Then
        detail = "annual: a model with no drivers has no annual vector"
        Exit Function
    End If
    If yearCount < 1 Then
        detail = "annual: a block with no project years has no annual vector"
        Exit Function
    End If

    ReDim result(0 To yearCount - 1)
    ReDim terms(0 To driverCount - 1)
    For year = 0 To yearCount - 1
        For driver = 0 To driverCount - 1
            observation = observations(LBound(observations) + driver)
            If observation = 0# Then
                ' A Risk that did not occur contributes EXACTLY zero to every
                ' year. Forming 0 * K would be the same number and would also
                ' refuse a factor the driver never reached.
                terms(driver) = 0#
            Else
                If Not modCalcFactors.SafeMultiply(observation, _
                        factors(driver * yearCount + year), term) Then
                    detail = "annual: driver " & CStr(driver + 1) & " project year " & _
                             CStr(firstYear + year + 1) & " contribution is not representable"
                    Exit Function
                End If
                terms(driver) = term
            End If
        Next driver
        If Not modCalcFactors.SafeSignedSum(terms, driverCount, result(year)) Then
            detail = "annual: project year " & CStr(firstYear + year + 1) & _
                     " total is not representable"
            Exit Function
        End If
    Next year
    SimAnnualVector = True
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
