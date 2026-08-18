Attribute VB_Name = "modCalcFactors"
Option Explicit

' ==========================================================================
' PCCM Phase 5 - numerical kernel. WORKSHEET-FREE BY CONSTRUCTION.
'
' Every procedure is a pure function of Doubles, Longs and typed arrays, so the
' same arguments give the same answers with no workbook open. The Python oracle
' in builder/pccm_builder/calc_numeric.py is the single definition of these
' semantics; this module is their VBA implementation.
'
' THE OBSERVABLE CONTRACT, not the mechanism, is what must match:
'   representable result -> True, result set;  unrepresentable -> False, result
'   untouched;  divide by zero -> False;  a non-zero product or quotient that
'   collapses to exactly zero -> False. A failure never becomes a zero and never
'   escapes as a raw VBA runtime error.
'
' TWO TIERS EVERYWHERE. The ordinary evaluation runs first and, when it
' succeeds, its result is returned bit for bit. Only a failed ordinary
' evaluation reaches the exact kernel, which computes the exact mathematical
' value of the already-converted Doubles, classifies its range exactly and
' rounds once.
' ==========================================================================

Public Const MAX_DOUBLE As Double = 1.79769313486232E+308
Private Const MIN_NORMAL_DOUBLE As Double = 2.2250738585072E-308

' Base 2^24 keeps every limb product below 2^48 and every running limb-plus-
' carry expression below 2^49, comfortably inside the 53-bit exact integer
' range of a Double. Nothing here relies on extended precision.
Private Const LIMB_BITS As Long = 24
Private Const LIMB_BASE As Double = 16777216#
Private Const TWO_52 As Double = 4503599627370500#
Private Const MAX_SIGNIFICAND As Double = 9007199254740990#
Private Const MAX_EXPONENT As Long = 971
Private Const MIN_SUBNORMAL_EXPONENT As Long = -1074
Private Const GUARD_BITS As Long = 64

' The internal shape vocabulary. Deliberately NOT the accepted display names:
' these say what the mathematics does, and the master list of distribution names
' stays with input_contract.yaml and the resolver that reads it.
Public Const DIST_TRIANGULAR As Long = 1
Public Const DIST_BETA_PERT As Long = 2
Public Const DIST_UNIFORM As Long = 3

' --------------------------------------------------------------------------
' The two carry types Phase 6 reuses. Row numbers and addresses are
' deliberately absent: identity here is the permanent ID, never a position.
' --------------------------------------------------------------------------
Public Type DriverFactors
    PermanentId   As String
    IsRisk        As Boolean
    Knom          As Double
    Kpv           As Double
    Quantity      As Double
    Probability   As Double
    DistKind      As Long
    CentralBasis  As String
    MinValue      As Double
    MostLikely    As Double
    MaxValue      As Double
    Central       As Double
    MeanValue     As Double
End Type

Public Type YearFactors
    ProjectIndex  As Long
    CalendarYear  As Long
    DiscountF     As Double
End Type

' sign is -1, 0 or +1; limbs(0) is least significant; value is
' sign * SUM limbs(i) * 2^(24i) * 2^shift.
Private Type ExactNumber
    Sign  As Long
    Shift As Long
    Count As Long
    Limbs() As Double
End Type

' ==========================================================================
' The predicate
' ==========================================================================
Public Function IsUsableDouble(ByVal value As Double) As Boolean
    ' A NaN fails every comparison including equality with itself, which is the
    ' only portable test available here. An infinity fails the range test.
    If Not (value = value) Then Exit Function
    If value > MAX_DOUBLE Then Exit Function
    If value < -MAX_DOUBLE Then Exit Function
    IsUsableDouble = True
End Function

' ==========================================================================
' Safe primitives. These four are the ONLY procedures in this module that
' install an error handler, and each scopes it to a single expression.
' ==========================================================================
Public Function SafeAdd(ByVal a As Double, ByVal b As Double, _
                        ByRef result As Double) As Boolean
    Dim tmp As Double
    If Not IsUsableDouble(a) Then Exit Function
    If Not IsUsableDouble(b) Then Exit Function
    On Error GoTo ArithmeticFailure
    tmp = a + b
    On Error GoTo 0
    If Not IsUsableDouble(tmp) Then Exit Function
    result = tmp
    SafeAdd = True
    Exit Function
ArithmeticFailure:
    On Error GoTo 0
    SafeAdd = False
End Function

Public Function SafeSubtract(ByVal a As Double, ByVal b As Double, _
                             ByRef result As Double) As Boolean
    Dim tmp As Double
    If Not IsUsableDouble(a) Then Exit Function
    If Not IsUsableDouble(b) Then Exit Function
    On Error GoTo ArithmeticFailure
    tmp = a - b
    On Error GoTo 0
    If Not IsUsableDouble(tmp) Then Exit Function
    result = tmp
    SafeSubtract = True
    Exit Function
ArithmeticFailure:
    On Error GoTo 0
    SafeSubtract = False
End Function

Public Function SafeMultiply(ByVal a As Double, ByVal b As Double, _
                             ByRef result As Double) As Boolean
    ' The underflow rule matters as much as the overflow rule: two non-zero
    ' operands whose product rounds to exactly zero would delete a real
    ' contribution with no error anywhere - the silent failure.
    Dim tmp As Double
    If Not IsUsableDouble(a) Then Exit Function
    If Not IsUsableDouble(b) Then Exit Function
    On Error GoTo ArithmeticFailure
    tmp = a * b
    On Error GoTo 0
    If Not IsUsableDouble(tmp) Then Exit Function
    If tmp = 0# And a <> 0# And b <> 0# Then Exit Function
    result = tmp
    SafeMultiply = True
    Exit Function
ArithmeticFailure:
    On Error GoTo 0
    SafeMultiply = False
End Function

Public Function SafeDivide(ByVal a As Double, ByVal b As Double, _
                           ByRef result As Double) As Boolean
    Dim tmp As Double
    If Not IsUsableDouble(a) Then Exit Function
    If Not IsUsableDouble(b) Then Exit Function
    If b = 0# Then Exit Function
    On Error GoTo ArithmeticFailure
    tmp = a / b
    On Error GoTo 0
    If Not IsUsableDouble(tmp) Then Exit Function
    If tmp = 0# And a <> 0# Then Exit Function
    result = tmp
    SafeDivide = True
    Exit Function
ArithmeticFailure:
    On Error GoTo 0
    SafeDivide = False
End Function

Public Function SafeAccumulate(ByRef accumulator As Double, _
                               ByVal term As Double) As Boolean
    ' Checked at THIS term, so a caller can name the driver or year responsible
    ' instead of reporting that a total came out infinite. It installs no
    ' handler of its own; SafeAdd owns that.
    Dim tmp As Double
    If Not SafeAdd(accumulator, term, tmp) Then Exit Function
    accumulator = tmp
    SafeAccumulate = True
End Function

' ==========================================================================
' EXACT KERNEL - the accepted Step-2 algorithm, translated.
'
' A value is (sign, base-2^24 limbs, binary shift). Every operation below is
' Double addition, subtraction, multiplication, division by an exact power of
' two, comparison or Fix. No Currency, no Decimal, no native Mod, no integer
' division, no WorksheetFunction, and no conversion of a wide significand to
' Long.
' ==========================================================================
Private Function FixNonNegative(ByVal value As Double) As Double
    ' Truncation toward zero for 0 <= value < 2^53, written so nothing depends
    ' on an integer type wider than Double.
    FixNonNegative = Fix(value)
End Function

Private Function PowerOfTwo(ByVal exponent As Long) As Double
    ' 2^exponent for 0 <= exponent <= 512, by doubling. One definition, no
    ' transcribed table to drift.
    Dim result As Double, index As Long
    result = 1#
    For index = 1 To exponent
        result = result * 2#
    Next index
    PowerOfTwo = result
End Function

Private Function ScaleByPowerOfTwo(ByVal value As Double, _
                                   ByVal exponent As Long) As Double
    ' Applied in exact steps of at most 2^512. Scaling a Double by a power of
    ' two moves the exponent and leaves the significand alone, so each step is
    ' exact while the running value stays in range - which callers guarantee.
    Dim remaining As Long, chunk As Long
    remaining = exponent
    Do While remaining <> 0
        If remaining > 512 Then
            chunk = 512
        ElseIf remaining < -512 Then
            chunk = -512
        Else
            chunk = remaining
        End If
        If chunk >= 0 Then
            value = value * PowerOfTwo(chunk)
        Else
            value = value / PowerOfTwo(-chunk)
        End If
        remaining = remaining - chunk
    Loop
    ScaleByPowerOfTwo = value
End Function

Private Function DecomposeDouble(ByVal value As Double, ByRef sign As Long, _
                                 ByRef mantissa As Double, _
                                 ByRef exponent As Long) As Boolean
    ' |value| = mantissa * 2^exponent with mantissa an integer in
    ' [2^52, 2^53). A counting loop over powers of two, not frexp: every step
    ' is an exact scaling, and lifting a subnormal into the normal range loses
    ' nothing.
    Dim magnitude As Double, power As Double, stepBits As Long
    If value = 0# Then
        sign = 0: mantissa = 0#: exponent = 0
        DecomposeDouble = True
        Exit Function
    End If
    If Not IsUsableDouble(value) Then Exit Function
    If value > 0# Then sign = 1 Else sign = -1
    magnitude = value
    If magnitude < 0# Then magnitude = -magnitude
    exponent = 0
    Do While magnitude < 1#
        magnitude = magnitude * PowerOfTwo(512)
        exponent = exponent - 512
    Loop
    stepBits = 512
    Do While stepBits >= 1
        power = PowerOfTwo(stepBits)
        Do While magnitude >= power
            magnitude = magnitude / power
            exponent = exponent + stepBits
        Loop
        stepBits = Fix(stepBits / 2)
    Loop
    mantissa = magnitude * TWO_52
    exponent = exponent - 52
    DecomposeDouble = True
End Function

Private Sub ExactInit(ByRef value As ExactNumber, ByVal limbCount As Long)
    If limbCount < 1 Then limbCount = 1
    value.Sign = 0
    value.Shift = 0
    value.Count = limbCount
    ReDim value.Limbs(0 To limbCount - 1)
End Sub

Private Sub ExactAddAt(ByRef value As ExactNumber, ByVal index As Long, _
                       ByVal amount As Double)
    ' value += amount * 2^(24*index); amount is an exact integer below 2^47.
    Dim carry As Double, total As Double
    carry = amount
    Do While carry <> 0#
        If index > value.Count - 1 Then Exit Sub
        total = value.Limbs(index) + carry
        carry = FixNonNegative(total / LIMB_BASE)
        value.Limbs(index) = total - carry * LIMB_BASE
        index = index + 1
    Loop
End Sub

Private Sub ExactAddShifted(ByRef value As ExactNumber, ByVal mantissa As Double, _
                            ByVal offsetBits As Long)
    ' value += mantissa * 2^offsetBits, offsetBits >= 0. The 53-bit mantissa is
    ' cut into three 24-bit pieces so each piece, shifted by the sub-limb
    ' remainder, stays an exact integer below 2^47.
    Dim index As Long, scale As Double, rest As Double
    Dim quotient As Double, low As Double, piece As Long
    index = Fix(offsetBits / LIMB_BITS)
    scale = PowerOfTwo(offsetBits - index * LIMB_BITS)
    rest = mantissa
    For piece = 0 To 2
        quotient = FixNonNegative(rest / LIMB_BASE)
        low = rest - quotient * LIMB_BASE
        rest = quotient
        If low <> 0# Then ExactAddAt value, index + piece, low * scale
        If rest = 0# Then Exit For
    Next piece
End Sub

Private Function ExactCompareMagnitude(ByRef leftSide As ExactNumber, _
                                       ByRef rightSide As ExactNumber) As Long
    Dim index As Long
    For index = leftSide.Count - 1 To 0 Step -1
        If leftSide.Limbs(index) <> rightSide.Limbs(index) Then
            If leftSide.Limbs(index) > rightSide.Limbs(index) Then
                ExactCompareMagnitude = 1
            Else
                ExactCompareMagnitude = -1
            End If
            Exit Function
        End If
    Next index
End Function

Private Sub ExactSubtractMagnitude(ByRef leftSide As ExactNumber, _
                                   ByRef rightSide As ExactNumber)
    ' leftSide -= rightSide; the caller has established leftSide >= rightSide.
    Dim index As Long, borrow As Double, value As Double
    For index = 0 To leftSide.Count - 1
        value = leftSide.Limbs(index) - rightSide.Limbs(index) - borrow
        If value < 0# Then
            value = value + LIMB_BASE
            borrow = 1#
        Else
            borrow = 0#
        End If
        leftSide.Limbs(index) = value
    Next index
End Sub

Private Function ExactIsZero(ByRef value As ExactNumber) As Boolean
    Dim index As Long
    For index = 0 To value.Count - 1
        If value.Limbs(index) <> 0# Then Exit Function
    Next index
    ExactIsZero = True
End Function

Private Sub ExactMultiply(ByRef leftSide As ExactNumber, ByRef rightSide As ExactNumber, _
                          ByRef product As ExactNumber)
    ' Schoolbook. Every intermediate stays below 2^49.
    Dim i As Long, j As Long, position As Long
    Dim carry As Double, total As Double
    ExactInit product, leftSide.Count + rightSide.Count + 1
    product.Sign = leftSide.Sign * rightSide.Sign
    product.Shift = leftSide.Shift + rightSide.Shift
    For i = 0 To leftSide.Count - 1
        If leftSide.Limbs(i) <> 0# Then
            carry = 0#
            For j = 0 To rightSide.Count - 1
                total = product.Limbs(i + j) + leftSide.Limbs(i) * rightSide.Limbs(j) + carry
                carry = FixNonNegative(total / LIMB_BASE)
                product.Limbs(i + j) = total - carry * LIMB_BASE
            Next j
            position = i + rightSide.Count
            Do While carry <> 0#
                total = product.Limbs(position) + carry
                carry = FixNonNegative(total / LIMB_BASE)
                product.Limbs(position) = total - carry * LIMB_BASE
                position = position + 1
            Loop
        End If
    Next i
End Sub

Private Function ExactDivideSmall(ByRef value As ExactNumber, ByVal divisor As Double, _
                                  ByRef quotient As ExactNumber) As Double
    ' Exact divmod by 2, 3 or 6. Returns the remainder, which the caller folds
    ' in as sticky information.
    Dim index As Long, current As Double, share As Double, remainder As Double
    ExactInit quotient, value.Count
    quotient.Sign = value.Sign
    quotient.Shift = value.Shift
    For index = value.Count - 1 To 0 Step -1
        current = remainder * LIMB_BASE + value.Limbs(index)
        share = FixNonNegative(current / divisor)
        quotient.Limbs(index) = share
        remainder = current - share * divisor
    Next index
    ExactDivideSmall = remainder
End Function

Private Function ExactTopBit(ByRef value As ExactNumber) As Long
    ' Index of the most significant set bit, or -1 when the value is zero.
    Dim index As Long, bit As Long, limb As Double
    For index = value.Count - 1 To 0 Step -1
        limb = value.Limbs(index)
        If limb <> 0# Then
            bit = 0
            Do While limb >= 2#
                limb = FixNonNegative(limb / 2#)
                bit = bit + 1
            Loop
            ExactTopBit = index * LIMB_BITS + bit
            Exit Function
        End If
    Next index
    ExactTopBit = -1
End Function

Private Function ExactBit(ByRef value As ExactNumber, ByVal position As Long) As Boolean
    Dim index As Long, shifted As Double
    index = Fix(position / LIMB_BITS)
    If index > value.Count - 1 Then Exit Function
    shifted = FixNonNegative(value.Limbs(index) / PowerOfTwo(position - index * LIMB_BITS))
    ExactBit = (shifted - 2# * FixNonNegative(shifted / 2#)) <> 0#
End Function

Private Function ExactAnyBelow(ByRef value As ExactNumber, ByVal position As Long) As Boolean
    Dim index As Long, offset As Long, lower As Long, scale As Double
    index = Fix(position / LIMB_BITS)
    offset = position - index * LIMB_BITS
    For lower = 0 To index - 1
        If lower <= value.Count - 1 Then
            If value.Limbs(lower) <> 0# Then
                ExactAnyBelow = True
                Exit Function
            End If
        End If
    Next lower
    If offset = 0 Or index > value.Count - 1 Then Exit Function
    scale = PowerOfTwo(offset)
    ExactAnyBelow = (value.Limbs(index) - FixNonNegative(value.Limbs(index) / scale) * scale) <> 0#
End Function

Private Function ExactHighPart(ByRef value As ExactNumber, ByVal drop As Long) As Double
    ' floor(value / 2^drop), which the caller guarantees is below 2^53. At most
    ' four limbs can reach a 53-bit result, so the loop is bounded.
    Dim index As Long, offset As Long, total As Double, weight As Double, stepBits As Long
    index = Fix(drop / LIMB_BITS)
    If index > value.Count - 1 Then Exit Function
    offset = drop - index * LIMB_BITS
    total = FixNonNegative(value.Limbs(index) / PowerOfTwo(offset))
    weight = PowerOfTwo(LIMB_BITS - offset)
    For stepBits = 1 To 3
        If index + stepBits > value.Count - 1 Then Exit For
        If value.Limbs(index + stepBits) <> 0# Then
            total = total + value.Limbs(index + stepBits) * weight
        End If
        weight = weight * LIMB_BASE
    Next stepBits
    ExactHighPart = total
End Function

Private Function RoundExact(ByRef value As ExactNumber, ByRef result As Double, _
                            ByVal stickyBelow As Boolean, _
                            ByVal underflowToZero As Boolean) As Boolean
    ' Round to nearest, ties to even, ONCE.
    '
    ' THE RANGE TEST IS ON THE EXACT VALUE, not on the rounded one. A total can
    ' exceed MAX_DOUBLE by less than half an ulp and still round to it; the
    ' q = 971 branch below refuses that rather than fabricating MAX_DOUBLE.
    Dim top As Long, exponent As Long, target As Long, drop As Long
    Dim quotient As Double, scale As Long
    Dim roundBit As Boolean, sticky As Boolean, odd As Boolean
    top = ExactTopBit(value)
    If top < 0 Then
        If stickyBelow And Not underflowToZero Then Exit Function
        result = 0#
        RoundExact = True
        Exit Function
    End If
    exponent = top + value.Shift
    If exponent > 1023 Then Exit Function
    target = exponent - 52
    If target < MIN_SUBNORMAL_EXPONENT Then target = MIN_SUBNORMAL_EXPONENT
    drop = target - value.Shift
    If drop <= 0 Then
        If stickyBelow Then Exit Function
        quotient = ExactHighPart(value, 0)
        scale = value.Shift
    Else
        quotient = ExactHighPart(value, drop)
        roundBit = ExactBit(value, drop - 1)
        sticky = stickyBelow
        If Not sticky Then sticky = ExactAnyBelow(value, drop - 1)
        If target = MAX_EXPONENT Then
            If quotient > MAX_SIGNIFICAND Then Exit Function
            If quotient = MAX_SIGNIFICAND And (roundBit Or sticky) Then Exit Function
        End If
        odd = (quotient - 2# * FixNonNegative(quotient / 2#)) <> 0#
        If roundBit And (sticky Or odd) Then quotient = quotient + 1#
        If quotient = 0# Then
            If Not underflowToZero Then Exit Function
            result = 0#
            RoundExact = True
            Exit Function
        End If
        scale = target
    End If
    quotient = ScaleByPowerOfTwo(quotient, scale)
    If Not IsUsableDouble(quotient) Then Exit Function
    If value.Sign < 0 Then quotient = -quotient
    result = quotient
    RoundExact = True
End Function

Private Function ExactSumOf(ByRef terms() As Double, ByRef total As ExactNumber) As Boolean
    ' Every Double is an exact integer multiple of 2^smallest, where smallest is
    ' the least of the terms' own exponents, so aligning there is exact. Positive
    ' and negative magnitudes accumulate separately and are subtracted once,
    ' which needs no signed carries and no ordering rule at all.
    Dim index As Long, sign As Long, mantissa As Double, exponent As Long
    Dim smallest As Long, largest As Long, seen As Boolean, count As Long
    Dim positive As ExactNumber, negative As ExactNumber
    For index = LBound(terms) To UBound(terms)
        If Not DecomposeDouble(terms(index), sign, mantissa, exponent) Then Exit Function
        If sign <> 0 Then
            If Not seen Then
                smallest = exponent: largest = exponent: seen = True
            Else
                If exponent < smallest Then smallest = exponent
                If exponent > largest Then largest = exponent
            End If
        End If
    Next index
    If Not seen Then
        ExactInit total, 1
        ExactSumOf = True
        Exit Function
    End If
    count = Fix((largest - smallest) / LIMB_BITS) + 6
    ExactInit positive, count
    ExactInit negative, count
    For index = LBound(terms) To UBound(terms)
        If DecomposeDouble(terms(index), sign, mantissa, exponent) Then
            If sign > 0 Then
                ExactAddShifted positive, mantissa, exponent - smallest
            ElseIf sign < 0 Then
                ExactAddShifted negative, mantissa, exponent - smallest
            End If
        End If
    Next index
    CombineMagnitudes positive, negative, smallest, total
    ExactSumOf = True
End Function

Private Sub CombineMagnitudes(ByRef positive As ExactNumber, ByRef negative As ExactNumber, _
                              ByVal smallest As Long, ByRef total As ExactNumber)
    ' One comparison and one subtraction. Accumulating the two signs separately
    ' is what makes this representation order-independent - there is no
    ' tie-breaking rule to specify and none to get wrong.
    Dim order As Long
    order = ExactCompareMagnitude(positive, negative)
    If order = 0 Then
        ExactInit total, 1
        Exit Sub
    End If
    If order > 0 Then
        ExactSubtractMagnitude positive, negative
        total = positive
        total.Sign = 1
    Else
        ExactSubtractMagnitude negative, positive
        total = negative
        total.Sign = -1
    End If
    total.Shift = smallest
End Sub

Private Function ExactProductOf(ByRef factors() As Double, _
                                ByRef product As ExactNumber) As Boolean
    ' The mantissas multiply as integers and the exponents add, so the product
    ' is exact however far outside Double range it lands - which is what makes
    ' the range classification a fact rather than an artefact of evaluation
    ' order.
    Dim index As Long, sign As Long, mantissa As Double, exponent As Long
    Dim factor As ExactNumber, running As ExactNumber, scratch As ExactNumber
    ExactInit running, 3
    running.Sign = 1
    running.Limbs(0) = 1#
    For index = LBound(factors) To UBound(factors)
        If Not DecomposeDouble(factors(index), sign, mantissa, exponent) Then Exit Function
        If sign = 0 Then
            ExactInit product, 1
            ExactProductOf = True
            Exit Function
        End If
        ExactInit factor, 3
        factor.Sign = sign
        factor.Shift = exponent
        ExactAddShifted factor, mantissa, 0
        ExactMultiply running, factor, scratch
        running = scratch
    Next index
    product = running
    ExactProductOf = True
End Function

' ==========================================================================
' Public exact-rescue surface
' ==========================================================================
Public Function SafeSignedSum(ByRef terms() As Double, ByRef result As Double) As Boolean
    ' TIER 1 - the canonical supplied order, unchanged. If it produces a value
    ' that value is returned bit for bit; a sum that already works is NEVER
    ' reordered, so canonical permanent-ID order still decides the answer.
    '
    ' TIER 2 - the exact sum, reached only when tier 1 overflowed. It does not
    ' re-associate Double additions: re-association discards the rounding
    ' residual of each intermediate subtraction, and once the large terms cancel
    ' that residual WAS the answer.
    Dim index As Long, total As Double, ok As Boolean
    Dim exact As ExactNumber
    total = 0#
    ok = True
    For index = LBound(terms) To UBound(terms)
        If Not SafeAccumulate(total, terms(index)) Then
            ok = False
            Exit For
        End If
    Next index
    If ok Then
        result = total
        SafeSignedSum = True
        Exit Function
    End If
    If Not ExactSumOf(terms, exact) Then Exit Function
    SafeSignedSum = RoundExact(exact, result, False, False)
End Function

Public Function SafeProduct(ByRef factors As Variant, ByRef result As Double) As Boolean
    ' Tier 1 is left to right. Tier 2 is the EXACT product, not a reordering: a
    ' magnitude-balanced order proves nothing about whether the exact product is
    ' in range, and was shown to accept one that exceeds MAX_DOUBLE while
    ' refusing one that rounds to 5e-324.
    Dim values() As Double
    Dim index As Long, running As Double, ok As Boolean, anyZero As Boolean
    Dim negatives As Long, exact As ExactNumber
    values = AsDoubleArray(factors)
    If UBound(values) < LBound(values) Then
        result = 1#
        SafeProduct = True
        Exit Function
    End If
    For index = LBound(values) To UBound(values)
        If Not IsUsableDouble(values(index)) Then Exit Function
        If values(index) = 0# Then anyZero = True
        If values(index) < 0# Then negatives = negatives + 1
    Next index
    If anyZero Then
        ' An exact zero makes the product exactly zero; no ordering changes that
        ' and the underflow rule must not fire on a genuine zero input.
        If negatives - 2 * Fix(negatives / 2) = 1 Then result = -0# Else result = 0#
        SafeProduct = True
        Exit Function
    End If
    running = 1#
    ok = True
    For index = LBound(values) To UBound(values)
        If Not SafeMultiply(running, values(index), running) Then
            ok = False
            Exit For
        End If
    Next index
    If ok Then
        result = running
        SafeProduct = True
        Exit Function
    End If
    If Not ExactProductOf(values, exact) Then Exit Function
    SafeProduct = RoundExact(exact, result, False, False)
End Function

Public Function ExactSumOfProducts(ByRef groups As Variant, _
                                   ByRef result As Double) As Boolean
    ' SUM over groups of PRODUCT of factors, formed exactly - including a
    ' product that has no Double of its own - and rounded once at the end.
    '
    ' This is the composition the materialization rule needs. Knom is
    ' SUM_y (FX * w_y * infl_y), the same number as FX * SUM_y (w_y * infl_y),
    ' but in this form w_y * infl_y may be wider than a Double while Knom is
    ' not. A per-driver annual contribution is the same story.
    Dim index As Long, sign As Long, smallest As Long, largest As Long
    Dim seen As Boolean, count As Long, top As Long
    Dim term As ExactNumber, positive As ExactNumber, negative As ExactNumber
    Dim total As ExactNumber
    Dim shifts() As Long, tops() As Long
    ReDim shifts(0 To UBound(groups) - LBound(groups))
    ReDim tops(0 To UBound(groups) - LBound(groups))
    For index = LBound(groups) To UBound(groups)
        If Not ExactProductOf(AsDoubleArray(groups(index)), term) Then Exit Function
        top = ExactTopBit(term)
        shifts(index - LBound(groups)) = term.Shift
        tops(index - LBound(groups)) = top
        If top >= 0 Then
            If Not seen Then
                smallest = term.Shift: largest = term.Shift + top: seen = True
            Else
                If term.Shift < smallest Then smallest = term.Shift
                If term.Shift + top > largest Then largest = term.Shift + top
            End If
        End If
    Next index
    If Not seen Then
        result = 0#
        ExactSumOfProducts = True
        Exit Function
    End If
    count = Fix((largest - smallest) / LIMB_BITS) + 6
    ExactInit positive, count
    ExactInit negative, count
    For index = LBound(groups) To UBound(groups)
        If ExactProductOf(AsDoubleArray(groups(index)), term) Then
            If tops(index - LBound(groups)) >= 0 Then
                sign = term.Sign
                AddExactShifted positive, negative, sign, term, term.Shift - smallest
            End If
        End If
    Next index
    CombineMagnitudes positive, negative, smallest, total
    If total.Sign = 0 Then
        result = 0#
        ExactSumOfProducts = True
        Exit Function
    End If
    ExactSumOfProducts = RoundExact(total, result, False, False)
End Function

Private Sub AddExactShifted(ByRef positive As ExactNumber, ByRef negative As ExactNumber, _
                            ByVal sign As Long, ByRef term As ExactNumber, _
                            ByVal offsetBits As Long)
    Dim index As Long
    For index = 0 To term.Count - 1
        If term.Limbs(index) <> 0# Then
            If sign > 0 Then
                ExactAddShifted positive, term.Limbs(index), offsetBits + index * LIMB_BITS
            Else
                ExactAddShifted negative, term.Limbs(index), offsetBits + index * LIMB_BITS
            End If
        End If
    Next index
End Sub

Public Function ExactQuotientOfSum(ByRef terms() As Double, ByVal divisor As Double, _
                                   ByRef result As Double) As Boolean
    ' Round (SUM terms) / divisor for divisor in {2, 3, 6}. The numerator is
    ' dyadic but the quotient is not, so the numerator is shifted left by
    ' GUARD_BITS and the division remainder becomes a sticky flag: a non-zero
    ' remainder means the true value is strictly above the quotient, which is
    ' exactly what a sticky bit encodes, so ties still resolve correctly.
    Dim exact As ExactNumber, guarded As ExactNumber, quotient As ExactNumber
    Dim index As Long, remainder As Double
    If Not ExactSumOf(terms, exact) Then Exit Function
    If divisor = 1# Then
        ExactQuotientOfSum = RoundExact(exact, result, False, False)
        Exit Function
    End If
    ExactInit guarded, exact.Count + Fix(GUARD_BITS / LIMB_BITS) + 2
    guarded.Sign = exact.Sign
    guarded.Shift = exact.Shift - GUARD_BITS
    For index = 0 To exact.Count - 1
        If exact.Limbs(index) <> 0# Then
            ExactAddShifted guarded, exact.Limbs(index), index * LIMB_BITS + GUARD_BITS
        End If
    Next index
    remainder = ExactDivideSmall(guarded, divisor, quotient)
    ExactQuotientOfSum = RoundExact(quotient, result, remainder <> 0#, False)
End Function

Private Function AsDoubleArray(ByRef source As Variant) As Double()
    Dim out() As Double, index As Long
    ReDim out(LBound(source) To UBound(source))
    For index = LBound(source) To UBound(source)
        out(index) = CDbl(source(index))
    Next index
    AsDoubleArray = out
End Function

' ==========================================================================
' Iterative factor series. Never a power: (1+r)^(t-1) can overflow as an
' intermediate where the reciprocal is representable, and it cannot say which
' year failed.
' ==========================================================================
Public Function BuildInflationFactors(ByVal baseYear As Long, ByVal lastYear As Long, _
                                      ByRef rates() As Double, ByRef factors() As Double, _
                                      ByRef detail As String) As Boolean
    Dim calendar As Long, running As Double, growth As Double
    detail = vbNullString
    If lastYear < baseYear Then Exit Function
    ReDim factors(0 To lastYear - baseYear)
    factors(0) = 1#
    running = 1#
    For calendar = baseYear + 1 To lastYear
        If Not SafeAdd(1#, rates(calendar - baseYear - 1), growth) Then
            detail = "inflation rate " & CStr(calendar)
            Exit Function
        End If
        If growth <= 0# Then
            detail = "inflation rate " & CStr(calendar)
            Exit Function
        End If
        If Not SafeMultiply(running, growth, running) Then
            detail = "inflation factor " & CStr(calendar)
            Exit Function
        End If
        If running <= 0# Then
            detail = "inflation factor " & CStr(calendar)
            Exit Function
        End If
        factors(calendar - baseYear) = running
    Next calendar
    BuildInflationFactors = True
End Function

Public Function BuildDiscountFactors(ByVal discountRate As Double, ByVal duration As Long, _
                                     ByRef factors() As Double, _
                                     ByRef detail As String) As Boolean
    Dim index As Long, divisor As Double, running As Double
    detail = vbNullString
    If duration < 1 Then Exit Function
    If Not SafeAdd(1#, discountRate, divisor) Then
        detail = "discount rate"
        Exit Function
    End If
    If divisor <= 0# Then
        detail = "discount rate"
        Exit Function
    End If
    ReDim factors(0 To duration - 1)
    factors(0) = 1#
    running = 1#
    For index = 2 To duration
        If Not SafeDivide(running, divisor, running) Then
            detail = "discount factor project year " & CStr(index)
            Exit Function
        End If
        If running <= 0# Then
            detail = "discount factor project year " & CStr(index)
            Exit Function
        End If
        factors(index - 1) = running
    Next index
    BuildDiscountFactors = True
End Function

' ==========================================================================
' Knom / Kpv - the C2 materialization boundary.
'
' Quantity and Probability are deliberately absent. Probability is replaced by
' a Bernoulli draw in Monte Carlo and must not be folded in here; Quantity is a
' per-driver multiplier, not a factor of the escalation path.
' ==========================================================================
Public Function BuildKnom(ByVal fxRate As Double, ByRef weights() As Double, _
                          ByRef inflation() As Double, ByRef result As Double, _
                          ByRef detail As String) As Boolean
    BuildKnom = BuildFactor(fxRate, weights, inflation, inflation, False, result, detail)
End Function

Public Function BuildKpv(ByVal fxRate As Double, ByRef weights() As Double, _
                         ByRef inflation() As Double, ByRef discount() As Double, _
                         ByRef result As Double, ByRef detail As String) As Boolean
    BuildKpv = BuildFactor(fxRate, weights, inflation, discount, True, result, detail)
End Function

Private Function BuildFactor(ByVal fxRate As Double, ByRef weights() As Double, _
                             ByRef inflation() As Double, ByRef discount() As Double, _
                             ByVal withDiscount As Boolean, ByRef result As Double, _
                             ByRef detail As String) As Boolean
    ' TIER 1 is the ordinary staging: form each w_y * infl_y (and * disc_y),
    ' sum in project-year order, then apply FX. If that COMPLETE pipeline
    ' succeeds the result is returned bit for bit, and FX is not distributed.
    '
    ' TIER 2 distributes FX into one exact expression. Neither w_y * infl_y nor
    ' the pre-FX sum is a published value, so neither is a representability
    ' boundary; Knom is published, so it is.
    Dim index As Long, count As Long, ok As Boolean
    Dim terms() As Double, staged As Double, scaled As Double
    Dim groups() As Variant, group() As Double
    detail = vbNullString
    count = UBound(weights) - LBound(weights) + 1
    If count < 1 Then Exit Function
    ReDim terms(0 To count - 1)
    ok = True
    For index = 0 To count - 1
        ReDim group(0 To 1 + IIf(withDiscount, 1, 0))
        group(0) = weights(LBound(weights) + index)
        group(1) = inflation(LBound(inflation) + index)
        If withDiscount Then group(2) = discount(LBound(discount) + index)
        If Not SafeProduct(group, terms(index)) Then
            ok = False
            detail = "project year " & CStr(index + 1)
            Exit For
        End If
    Next index
    If ok Then
        If SafeSignedSum(terms, staged) Then
            If SafeMultiply(fxRate, staged, scaled) Then
                result = scaled
                detail = vbNullString
                BuildFactor = True
                Exit Function
            End If
        End If
    End If
    ReDim groups(0 To count - 1)
    For index = 0 To count - 1
        ReDim group(0 To 2 + IIf(withDiscount, 1, 0))
        group(0) = fxRate
        group(1) = weights(LBound(weights) + index)
        group(2) = inflation(LBound(inflation) + index)
        If withDiscount Then group(3) = discount(LBound(discount) + index)
        groups(index) = group
    Next index
    detail = "compound factor expression"
    If ExactSumOfProducts(groups, result) Then
        detail = vbNullString
        BuildFactor = True
    End If
End Function

' ==========================================================================
' C1 conditioning magnitudes.
'
' The coefficient is DISTRIBUTED OVER THE TERMS rather than applied to their
' sum: the two are the same number, but the raw sum of contributions can exceed
' Double while coefficient * sum is perfectly representable.
' ==========================================================================
Public Function ConditioningScaledMagnitude(ByRef accumulator As Double, _
                                            ByVal term As Double, _
                                            ByVal coefficient As Double) As Boolean
    ' Underflow policy here is deliberately different from model arithmetic: a
    ' scaled term below roughly 5e-312 cannot move an allowance floored at
    ' coefficient * 1, so losing it changes no answer. Overflow is still a
    ' failure, because an allowance outside Double range cannot be compared
    ' against.
    Dim magnitude As Double, scaled As Double
    If Not IsUsableDouble(term) Then Exit Function
    magnitude = term
    If magnitude < 0# Then magnitude = -magnitude
    If Not SafeMultiply(coefficient, magnitude, scaled) Then
        ' Only an underflow to zero is tolerated; overflow is still a failure.
        scaled = coefficient * magnitude
        If Not IsUsableDouble(scaled) Then Exit Function
        If scaled <> 0# Then Exit Function
    End If
    ConditioningScaledMagnitude = SafeAccumulate(accumulator, scaled)
End Function

Public Function ConditioningScaledProduct(ByRef accumulator As Double, _
                                          ByRef factors As Variant, _
                                          ByVal coefficient As Double) As Boolean
    ' The same magnitude for a contribution that has NO Double of its own. The
    ' quantity C1 needs is coefficient * |contribution|, which is finite even
    ' where the contribution is 2 * MAX_DOUBLE, so the coefficient is folded
    ' into the SAME exact factor expression rather than forcing the unscaled
    ' contribution into a Double first.
    Dim values() As Double, group() As Double, groups() As Variant
    Dim index As Long, scaled As Double
    values = AsDoubleArray(factors)
    ReDim group(0 To UBound(values) - LBound(values) + 1)
    group(0) = coefficient
    For index = LBound(values) To UBound(values)
        If values(index) < 0# Then
            group(index - LBound(values) + 1) = -values(index)
        Else
            group(index - LBound(values) + 1) = values(index)
        End If
    Next index
    ReDim groups(0 To 0)
    groups(0) = group
    If Not ExactSumOfProducts(groups, scaled) Then
        ' Only an underflow of the metadata term is tolerated; see above.
        scaled = 0#
    End If
    ConditioningScaledProduct = SafeAccumulate(accumulator, scaled)
End Function

Public Function IdentityAllowance(ByVal scale As Double, ByVal absoluteFloor As Double, _
                                  ByVal coefficient As Double, ByVal scaleFloor As Double, _
                                  ByRef result As Double) As Boolean
    ' max(absoluteFloor, coefficient * max(scaleFloor, sum |terms|)).
    '
    ' NOTE THE TWO MAXIMA. The inner one is a MAXIMUM and not an addition:
    ' adding coefficient * scaleFloor to the scaled sum silently widens every
    ' allowance. `scale` already carries the coefficient distributed over the
    ' terms, so the floor is scaled once, here, and compared in the same units.
    Dim scaledFloor As Double, relative As Double
    If Not IsUsableDouble(scale) Then Exit Function
    If Not SafeMultiply(coefficient, scaleFloor, scaledFloor) Then Exit Function
    relative = scaledFloor
    If scale > relative Then relative = scale
    result = absoluteFloor
    If relative > result Then result = relative
    IdentityAllowance = True
End Function
