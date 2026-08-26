Attribute VB_Name = "modCalcFingerprint"
Option Explicit

' ==========================================================================
' modCalcFingerprint - the Calculation Input Fingerprint, and nothing else.
'
' Canonical encoding of the inputs, the double-modulus polynomial hash over
' UTF-16 code units, and the sixteen-character digest. No analytical cost, risk,
' FX, inflation, discounting or reconciliation logic appears here, and none may
' be added.
'
' THE HASH MATHEMATICS IS NOT RESTATED HERE. The base, both moduli, both initial
' states, the stream tag, the section names and the field tags all arrive as
' generated FP_* constants from modCalcContract, whose authority is the Python
' reference implementation. A modulus written as a literal in this module would
' be a second authority that could drift silently and would only be discovered
' on Windows.
'
' --------------------------------------------------------------------------
' THE GRAMMAR
' --------------------------------------------------------------------------
'   field   ::= <TAG> <LEN> ":" <VALUE>      LEN in UTF-16 code units
'   record  ::= F_I(field_count) field*
'   section ::= F_S(name) F_I(record_count) record*
'   stream  ::= F_S("PCCM-FP") F_I(version) section*
'
' The length prefix is what makes the encoding SELF-DELIMITING: a reader knows
' exactly how many code units to consume, so no character inside a value can be
' mistaken for a separator. A delimiter-joined encoding has no such property,
' whichever delimiter it picks.
'
' --------------------------------------------------------------------------
' WHY DOUBLE-ONLY ARITHMETIC
' --------------------------------------------------------------------------
' VBA's Mod operator and its \ integer division both use an effective integral
' type of Long on floating-point operands, and the recurrence intermediate
' reaches 281,320,423,161 - roughly 131 times the signed-Long maximum. Neither
' operator can express this recurrence, so the reduction is done entirely in
' Double with two corrections; see CalcFpReduceDouble.
'
' NOTHING IN THIS MODULE HAS BEEN EXECUTED. It is source, submitted for review.
' Parity with the reference implementation is proven on Windows at Gate B.
' ==========================================================================

' The locked canonical form: one digit, a decimal point, sixteen fractional
' digits (17 significant digits in total), an uppercase E, an always-present
' exponent sign and at least two exponent digits. The digit count is the
' encoding authority's, mirrored here because it is not expressible as a
' projected constant.
'
' THERE IS NO FORMAT STRING ANY MORE, and that is the point. See
' CalcFpCanonicalNumber.
Private Const FP_FRACTION_DIGITS As Long = 16
Private Const FP_SIGNIFICANT_DIGITS As Long = 17

' The exact-integer scratch arithmetic below. A limb holds seven decimal digits,
' so a limb-by-factor product stays under 10^14 and every intermediate remains
' an EXACT Double integer (the exact-integer ceiling is 2^53, about 9.007e15).
Private Const FP_LIMB_BASE As Double = 10000000#
Private Const FP_LIMB_DIGITS As Long = 7
' The decomposition bounds, as BIT WIDTHS rather than as decimal spellings.
' These were sixteen-significant-digit Double Consts, and this module's own
' account of Runtime Run 3 is that VBA's literal conversion cannot be assumed to
' preserve more than about fifteen. CalcFpDecompose builds them with
' CalcFpIntegerPower, which doubles from 1# and is exact.
Private Const FP_MANTISSA_BITS As Long = 52
Private Const FP_SIGNIFICAND_BITS As Long = 53
Private Const FP_MANTISSA_DIGITS As Long = 16
' M * 5^1126 is at most 804 decimal digits, so 115 limbs; 200 is headroom, and
' CalcFpMultiplySmall refuses rather than overruns if it is ever reached.
Private Const FP_MAX_LIMBS As Long = 200
Private Const FP_DIGIT_TABLE As String = "0123456789"
Private Const FP_FIELD_SEPARATOR As String = ":"
Private Const FP_HEX_DIGITS As String = "0123456789ABCDEF"
Private Const FP_HEX_WIDTH As Long = 8

' ==========================================================================
' UTF-16 code units
' ==========================================================================
Public Function CalcFpUtf16Length(ByVal text As String) As Long
    ' A VBA String IS UTF-16, so Len already counts code units and a non-BMP
    ' character is already counted twice, as its surrogate pair. That is exactly
    ' the length the prefix carries. The Python reference cannot use its own len
    ' for this, because Python counts code points; the asymmetry is real and is
    ' why the length rule is stated rather than assumed.
    CalcFpUtf16Length = Len(text)
End Function

Public Function CalcFpNormaliseCodeUnit(ByVal unit As Long) As Long
    ' AscW returns an Integer, which is SIGNED 16-bit, so every code unit above
    ' U+7FFF comes back negative. Adding 65536 is the whole normalisation, and
    ' the U+9AD8 vector is what proves it matters.
    Dim normalised As Long
    normalised = unit
    If normalised < 0 Then normalised = normalised + 65536
    If normalised < 0 Or normalised > 65535 Then
        CalcFpNormaliseCodeUnit = -1
        Exit Function
    End If
    CalcFpNormaliseCodeUnit = normalised
End Function

' ==========================================================================
' Canonical field encoding
' ==========================================================================
Public Function CalcFpCanonicalText(ByVal value As String) As String
    ' A text field is hashed exactly as given: never trimmed, never case-folded,
    ' never re-encoded.
    CalcFpCanonicalText = CalcFpField(FP_TAG_TEXT, value)
End Function

Public Function CalcFpCanonicalInteger(ByVal value As Long, ByRef result As String) As Boolean
    ' A STREAM integer: a structural count or a version, never model data. Kept
    ' distinct from a number field so a count of 1 and a Double of 1 can never
    ' encode identically - structural shape and model magnitude are different
    ' kinds of fact.
    If value < 0 Then Exit Function
    result = CalcFpField(FP_TAG_INTEGER, CalcFpDigitsOf(value))
    CalcFpCanonicalInteger = True
End Function

Public Function CalcFpCanonicalNumber(ByVal value As Double, ByVal decimalSeparator As String, _
                                      ByRef result As String) As Boolean
    ' The canonical TEXT of a Double: 17 significant digits, a decimal point
    ' always present, an uppercase E, the exponent sign always present, at least
    ' two exponent digits, no thousands separator, and negative zero normalised
    ' to positive zero.
    '
    ' --------------------------------------------------------------------
    ' WHY THIS IS NOT Format$ ANY MORE
    ' --------------------------------------------------------------------
    ' It was, and Gate B Runtime Run 2 proved on real Excel that it could not
    ' meet the contract. Format$(number, "0.0000000000000000E+00") supplies
    ' sixteen fractional PLACEHOLDERS, but VBA's numeric-to-text conversion
    ' carries only about 15 significant decimal digits; the placeholders beyond
    ' that are filled with ZEROS rather than with recovered digits. The observed
    ' failures were exactly that shape:
    '
    '   0.1          got 1.0000000000000000E-01  want 1.0000000000000001E-01
    '   1e-20        got 1.0000000000000000E-20  want 9.9999999999999995E-21
    '   0.1 + 0.2    got 3.0000000000000000E-01  want 3.0000000000000004E-01
    '   MAX_DOUBLE   got 1.7976931348623200E+308 want 1.7976931348623157E+308
    '
    ' Every one is 15 correct significant digits followed by padding. Adding
    ' more placeholders cannot help: the digits were never produced.
    '
    ' A binary64 is exactly M * 2^E with M an integer, so its decimal expansion
    ' is FINITE and can be computed exactly with integer arithmetic. That is
    ' what happens below - the text is GENERATED, digit by digit, never
    ' formatted. Nothing is asked of the host's number-to-text conversion, so
    ' nothing depends on its precision.
    '
    ' --------------------------------------------------------------------
    ' WHY THE SEPARATOR IS STILL A PARAMETER
    ' --------------------------------------------------------------------
    ' It is the locked public interface, it is what the resolver reports about
    ' its environment, and it is what Gate B injects. It is still validated as
    ' exactly one UTF-16 code unit, so an unusable separator is still refused.
    '
    ' What changed is that it can no longer AFFECT the output. The old
    ' implementation had to locate the host formatter's marker and rewrite it,
    ' because the host chose that character. This implementation emits the
    ' marker itself, so separator invariance - the locked acceptance case that
    ' injects both "." and "," and requires byte-identical output - is now true
    ' by construction rather than by repair.
    Dim text As String, number As Double
    If CalcFpUtf16Length(decimalSeparator) <> 1 Then Exit Function
    If Not IsUsableDouble(value) Then Exit Function
    number = value
    ' Normalises negative zero, which compares equal to zero but carries a sign.
    If number = 0# Then number = 0#
    If Not CalcFpBuildCanonical(number, text) Then Exit Function
    ' THE ACCEPTED STRUCTURAL VALIDATION IS KEPT, now as a post-condition on
    ' this module's own output rather than on a foreign formatter's. A generator
    ' that ever produced the wrong shape refuses instead of emitting it.
    If CalcFpMarkerIndex(text) = 0 Then Exit Function
    result = text
    CalcFpCanonicalNumber = True
End Function

Private Function CalcFpBuildCanonical(ByVal value As Double, ByRef text As String) As Boolean
    ' value = M * 2^E exactly, with M an integer in [2^52, 2^53).
    '
    '   E >= 0 : the value is the integer M * 2^E, decimal point at the end.
    '   E <  0 : M / 2^-E = M * 5^-E / 10^-E, so the digits are those of the
    '            integer M * 5^-E with the point -E places from the right.
    '
    ' Either way the whole expansion is computed exactly, then rounded ONCE to
    ' 17 significant digits.
    Dim limbs() As Double, limbCount As Long
    Dim mantissa As Double, exponent As Long
    Dim allDigits As String, head As String, sign As String
    ' decimalScale, NOT scale. Runtime Run 3's VBE stopped here with a Syntax
    ' error and highlighted `scale`. Scale is a Visual Basic statement keyword,
    ' and this is the one place in the project where it stands as the token
    ' immediately after Dim - the seven pre-existing `scale` locals in
    ' modCalcFactors and modCalcAnalytical all sit later in their Dim lists and
    ' all compiled in Runtime Run 2. Whatever the parser's precise rule, the
    ' identifier is not worth the argument.
    Dim decimalScale As Long, exp10 As Long

    If value = 0# Then
        text = "0." & String$(FP_FRACTION_DIGITS, "0") & "E+00"
        CalcFpBuildCanonical = True
        Exit Function
    End If
    If value < 0# Then sign = "-"

    If Not CalcFpDecompose(value, mantissa, exponent) Then Exit Function
    If Not CalcFpLimbsFromMantissa(mantissa, limbs, limbCount) Then Exit Function

    If exponent >= 0 Then
        ' 2^23 is 8388608, under one limb, so each pass consumes 23 powers.
        If Not CalcFpMultiplyPower(limbs, limbCount, 2#, exponent, 23) Then Exit Function
        decimalScale = 0
    Else
        ' 5^10 is 9765625, under one limb, so each pass consumes ten powers.
        If Not CalcFpMultiplyPower(limbs, limbCount, 5#, -exponent, 10) Then Exit Function
        decimalScale = exponent
    End If

    allDigits = CalcFpLimbDigits(limbs, limbCount)
    If Len(allDigits) = 0 Then Exit Function
    exp10 = Len(allDigits) - 1 + decimalScale
    If Not CalcFpRoundSignificant(allDigits, head, exp10) Then Exit Function

    text = sign & Left$(head, 1) & "." & Mid$(head, 2) & "E" & CalcFpExponentText(exp10)
    CalcFpBuildCanonical = True
End Function

Private Function CalcFpDecompose(ByVal value As Double, ByRef mantissa As Double, _
                                 ByRef exponent As Long) As Boolean
    ' EXACT, because the only operations are multiplication and division by two.
    ' Scaling up stops below 2^53 so it cannot overflow; scaling down stops at
    ' or above 2^52 so it cannot underflow. Subnormals are covered: the smallest
    ' positive Double, 2^-1074, normalises to M = 2^52 with E = -1126.
    Dim scaled As Double, guard As Long
    Dim lowerBound As Double, upperBound As Double
    ' Built once per call, by exact doubling. No decimal literal is trusted.
    lowerBound = CalcFpIntegerPower(2#, FP_MANTISSA_BITS)
    upperBound = CalcFpIntegerPower(2#, FP_SIGNIFICAND_BITS)
    scaled = Abs(value)
    exponent = 0
    Do While scaled < lowerBound
        scaled = scaled * 2#
        exponent = exponent - 1
        guard = guard + 1
        If guard > 1200 Then Exit Function
    Loop
    Do While scaled >= upperBound
        scaled = scaled / 2#
        exponent = exponent + 1
        guard = guard + 1
        If guard > 2400 Then Exit Function
    Loop
    mantissa = scaled
    CalcFpDecompose = True
End Function

Private Function CalcFpLimbsFromMantissa(ByVal mantissa As Double, ByRef limbs() As Double, _
                                         ByRef limbCount As Long) As Boolean
    ' M is an integer below 2^53, so it has at most sixteen decimal digits and
    ' every subtraction here is between exact integers - no division, no
    ' rounding, no dependence on how the host converts numbers to text.
    Dim digits(1 To FP_MANTISSA_DIGITS) As Long
    Dim power As Double, remainder As Double, value As Double
    Dim place As Long, count As Long, first As Long, index As Long

    remainder = mantissa
    ' 10^15, built by exact multiplication rather than spelled as a sixteen-digit
    ' literal. Every intermediate 10^k for k <= 15 is exactly representable, so
    ' each step is exact.
    power = CalcFpIntegerPower(10#, FP_MANTISSA_DIGITS - 1)
    For place = 1 To FP_MANTISSA_DIGITS
        count = 0
        Do While remainder >= power
            remainder = remainder - power
            count = count + 1
            If count > 9 Then Exit Function
        Loop
        digits(place) = count
        power = power / 10#
    Next place
    If remainder <> 0# Then Exit Function

    ReDim limbs(0 To FP_MAX_LIMBS)
    limbCount = 0
    place = FP_MANTISSA_DIGITS
    Do While place >= 1
        first = place - FP_LIMB_DIGITS + 1
        If first < 1 Then first = 1
        value = 0#
        For index = first To place
            value = value * 10# + digits(index)
        Next index
        limbs(limbCount) = value
        limbCount = limbCount + 1
        place = first - 1
    Loop
    Do While limbCount > 1
        If limbs(limbCount - 1) <> 0# Then Exit Do
        limbCount = limbCount - 1
    Loop
    CalcFpLimbsFromMantissa = True
End Function

Private Function CalcFpMultiplyPower(ByRef limbs() As Double, ByRef limbCount As Long, _
                                     ByVal powerBase As Double, ByVal count As Long, _
                                     ByVal chunk As Long) As Boolean
    ' Chunked so the work is proportional to count/chunk passes rather than to
    ' count. The chunk factor is itself under one limb, which is what keeps
    ' every product inside the exact-integer range.
    ' NO \ AND NO Mod ANYWHERE IN THIS MODULE, including here where both
    ' operands are Longs and either would have been correct. The accepted
    ' invariant is that this module contains neither operator at all, so that a
    ' reader never has to decide which occurrences are safe; the loop below is
    ' bounded by count/chunk, at most 113 iterations.
    Dim passes As Long, remainder As Long, factor As Double, index As Long
    If count < 0 Then Exit Function
    If chunk < 1 Then Exit Function
    passes = 0
    remainder = count
    Do While remainder >= chunk
        remainder = remainder - chunk
        passes = passes + 1
    Loop
    factor = CalcFpIntegerPower(powerBase, chunk)
    For index = 1 To passes
        If Not CalcFpMultiplySmall(limbs, limbCount, factor) Then Exit Function
    Next index
    If remainder > 0 Then
        If Not CalcFpMultiplySmall(limbs, limbCount, CalcFpIntegerPower(powerBase, remainder)) Then Exit Function
    End If
    CalcFpMultiplyPower = True
End Function

Private Function CalcFpIntegerPower(ByVal powerBase As Double, ByVal power As Long) As Double
    Dim result As Double, index As Long
    result = 1#
    For index = 1 To power
        result = result * powerBase
    Next index
    CalcFpIntegerPower = result
End Function

Private Function CalcFpMultiplySmall(ByRef limbs() As Double, ByRef limbCount As Long, _
                                     ByVal factor As Double) As Boolean
    ' limb < 10^7 and factor < 10^7, so limb * factor < 10^14 and the running
    ' carry stays under 10^7 - the whole product is an exact Double integer.
    Dim index As Long, carry As Double, product As Double, quotient As Double
    carry = 0#
    For index = 0 To limbCount - 1
        product = limbs(index) * factor + carry
        quotient = Int(product / FP_LIMB_BASE)
        limbs(index) = product - quotient * FP_LIMB_BASE
        carry = quotient
    Next index
    Do While carry > 0#
        If limbCount > UBound(limbs) Then Exit Function
        quotient = Int(carry / FP_LIMB_BASE)
        limbs(limbCount) = carry - quotient * FP_LIMB_BASE
        limbCount = limbCount + 1
        carry = quotient
    Loop
    CalcFpMultiplySmall = True
End Function

Private Function CalcFpLimbDigits(ByRef limbs() As Double, ByVal limbCount As Long) As String
    ' The most significant limb keeps its natural width; every other limb is
    ' zero-padded to the full limb width, or its leading zeros would vanish.
    Dim out As String, part As String, index As Long
    out = CalcFpPlainDigits(limbs(limbCount - 1))
    For index = limbCount - 2 To 0 Step -1
        part = CalcFpPlainDigits(limbs(index))
        out = out & String$(FP_LIMB_DIGITS - Len(part), "0") & part
    Next index
    CalcFpLimbDigits = out
End Function

Private Function CalcFpPlainDigits(ByVal value As Double) As String
    ' NO CStr AND NO Format. The canonical text must not depend on any locale,
    ' and a digit is selected from a literal table rather than converted.
    Dim out As String, remainder As Double, power As Double, count As Long
    If value = 0# Then
        CalcFpPlainDigits = "0"
        Exit Function
    End If
    remainder = value
    power = 1000000#
    Do While power >= 1#
        count = 0
        Do While remainder >= power
            remainder = remainder - power
            count = count + 1
        Loop
        If Len(out) > 0 Or count > 0 Then out = out & Mid$(FP_DIGIT_TABLE, count + 1, 1)
        power = power / 10#
    Loop
    CalcFpPlainDigits = out
End Function

Private Function CalcFpRoundSignificant(ByVal allDigits As String, ByRef head As String, _
                                        ByRef exp10 As Long) As Boolean
    ' ROUND HALF TO EVEN, matching the reference encoder. The tie case is not
    ' theoretical: a binary64's exact expansion terminates, so the eighteenth
    ' significant digit really can be a 5 with nothing after it.
    Dim total As Long, nextDigit As String, tail As String, carried As String
    Dim roundUp As Boolean, lastDigit As Long
    total = Len(allDigits)
    If total <= FP_SIGNIFICANT_DIGITS Then
        head = allDigits & String$(FP_SIGNIFICANT_DIGITS - total, "0")
        CalcFpRoundSignificant = True
        Exit Function
    End If
    head = Left$(allDigits, FP_SIGNIFICANT_DIGITS)
    nextDigit = Mid$(allDigits, FP_SIGNIFICANT_DIGITS + 1, 1)
    tail = Mid$(allDigits, FP_SIGNIFICANT_DIGITS + 2)
    lastDigit = CalcFpDigitValue(Right$(head, 1))
    If lastDigit < 0 Then Exit Function
    If nextDigit > "5" Then
        roundUp = True
    ElseIf nextDigit = "5" Then
        If CalcFpHasNonZero(tail) Then
            roundUp = True
        Else
            roundUp = CalcFpIsOddDigit(lastDigit)
        End If
    End If
    If roundUp Then
        If Not CalcFpIncrementDigits(head, carried) Then Exit Function
        If Len(carried) > FP_SIGNIFICANT_DIGITS Then
            ' 999...9 carried into 1000...0: one more decimal place.
            head = Left$(carried, FP_SIGNIFICANT_DIGITS)
            exp10 = exp10 + 1
        Else
            head = carried
        End If
    End If
    CalcFpRoundSignificant = True
End Function

Private Function CalcFpHasNonZero(ByVal text As String) As Boolean
    Dim index As Long
    For index = 1 To Len(text)
        If Mid$(text, index, 1) <> "0" Then
            CalcFpHasNonZero = True
            Exit Function
        End If
    Next index
End Function

Private Function CalcFpIncrementDigits(ByVal digits As String, ByRef out As String) As Boolean
    ' String arithmetic, because seventeen digits do not fit any VBA integer.
    Dim buffer As String, index As Long, value As Long, carry As Long
    buffer = digits
    carry = 1
    For index = Len(buffer) To 1 Step -1
        value = CalcFpDigitValue(Mid$(buffer, index, 1))
        If value < 0 Then Exit Function
        value = value + carry
        If value >= 10 Then
            value = value - 10
            carry = 1
        Else
            carry = 0
        End If
        Mid$(buffer, index, 1) = Mid$(FP_DIGIT_TABLE, value + 1, 1)
        If carry = 0 Then Exit For
    Next index
    If carry = 1 Then buffer = "1" & buffer
    out = buffer
    CalcFpIncrementDigits = True
End Function

Private Function CalcFpDigitValue(ByVal char As String) As Long
    Dim index As Long
    CalcFpDigitValue = -1
    For index = 1 To 10
        If Mid$(FP_DIGIT_TABLE, index, 1) = char Then
            CalcFpDigitValue = index - 1
            Exit Function
        End If
    Next index
End Function

Private Function CalcFpExponentText(ByVal exp10 As Long) As String
    Dim sign As String, magnitude As Long, digits As String
    If exp10 < 0 Then
        sign = "-"
        magnitude = -exp10
    Else
        sign = "+"
        magnitude = exp10
    End If
    digits = CalcFpLongDigits(magnitude)
    If Len(digits) < 2 Then digits = String$(2 - Len(digits), "0") & digits
    CalcFpExponentText = sign & digits
End Function

Private Function CalcFpIsOddDigit(ByVal digit As Long) As Boolean
    ' Stated as the ten cases rather than as digit Mod 2, because this module
    ' contains no Mod and no \ - see CalcFpMultiplyPower.
    Select Case digit
        Case 1, 3, 5, 7, 9
            CalcFpIsOddDigit = True
    End Select
End Function

Private Function CalcFpLongDigits(ByVal value As Long) As String
    ' Again no CStr: the exponent digits are as locale-sensitive as any other.
    ' The exponent magnitude never exceeds 1200, so four places are enough and
    ' the subtraction loop runs at most nine times per place.
    Dim out As String, remainder As Long, power As Long, count As Long
    If value = 0 Then
        CalcFpLongDigits = "0"
        Exit Function
    End If
    If value < 0 Then Exit Function
    remainder = value
    power = 10000
    Do While power >= 1
        count = 0
        Do While remainder >= power
            remainder = remainder - power
            count = count + 1
        Loop
        If Len(out) > 0 Or count > 0 Then out = out & Mid$(FP_DIGIT_TABLE, count + 1, 1)
        power = CalcFpTenthOf(power)
    Loop
    CalcFpLongDigits = out
End Function

Private Function CalcFpTenthOf(ByVal power As Long) As Long
    ' The five exact places a Long exponent magnitude can occupy.
    Select Case power
        Case 10000
            CalcFpTenthOf = 1000
        Case 1000
            CalcFpTenthOf = 100
        Case 100
            CalcFpTenthOf = 10
        Case 10
            CalcFpTenthOf = 1
        Case Else
            CalcFpTenthOf = 0
    End Select
End Function

Public Function CalcFpNumberField(ByVal value As Double, ByVal decimalSeparator As String, _
                                 ByRef result As String) As Boolean
    ' The N field around a canonical number. Separate from CalcFpCanonicalNumber
    ' because the numeric encoding is the only one that can fail, and Gate B
    ' compares the canonical TEXT of the ten locked numeric vectors rather than
    ' their framed fields.
    '
    ' PUBLIC since Step 7. The orchestration layer frames the four header
    ' scalars - Base Year, Start Year, Duration and Discount Rate - and they are
    ' NUMBER fields. It must reach the accepted framing authority rather than
    ' assemble an N field of its own, and it must not reach CalcFpCanonicalText,
    ' which would tag a number as text and change what the digest covers. The
    ' body below is unchanged.
    Dim text As String
    If Not CalcFpCanonicalNumber(value, decimalSeparator, text) Then Exit Function
    result = CalcFpField(FP_TAG_NUMBER, text)
    CalcFpNumberField = True
End Function

Private Function CalcFpField(ByVal tag As String, ByVal value As String) As String
    CalcFpField = tag & CalcFpDigitsOf(CalcFpUtf16Length(value)) & FP_FIELD_SEPARATOR & value
End Function

Private Function CalcFpDigitsOf(ByVal value As Long) As String
    ' Plain decimal digits. CStr of a Long carries no separator and no sign for a
    ' non-negative value, which is the whole requirement.
    CalcFpDigitsOf = CStr(value)
End Function

Private Function CalcFpMarkerIndex(ByVal text As String) As Long
    ' The 1-based index of the mantissa decimal marker, located by POSITION and
    ' validated structurally: optional "-", exactly one digit, the marker,
    ' sixteen fractional digits, "E", an exponent sign, at least two exponent
    ' digits. Returns 0 if the host formatter produced anything else.
    Dim first As Long, marker As Long, index As Long, tail As Long
    first = 1
    If Mid$(text, 1, 1) = "-" Then first = 2
    If Not CalcFpIsDigit(Mid$(text, first, 1)) Then Exit Function
    marker = first + 1
    If Len(text) < marker + FP_FRACTION_DIGITS + 3 Then Exit Function
    For index = marker + 1 To marker + FP_FRACTION_DIGITS
        If Not CalcFpIsDigit(Mid$(text, index, 1)) Then Exit Function
    Next index
    If Mid$(text, marker + FP_FRACTION_DIGITS + 1, 1) <> "E" Then Exit Function
    tail = marker + FP_FRACTION_DIGITS + 2
    If Mid$(text, tail, 1) <> "+" And Mid$(text, tail, 1) <> "-" Then Exit Function
    If Len(text) - tail < 2 Then Exit Function
    For index = tail + 1 To Len(text)
        If Not CalcFpIsDigit(Mid$(text, index, 1)) Then Exit Function
    Next index
    CalcFpMarkerIndex = marker
End Function

Private Function CalcFpIsDigit(ByVal char As String) As Boolean
    If Len(char) <> 1 Then Exit Function
    CalcFpIsDigit = (char >= "0" And char <= "9")
End Function

' ==========================================================================
' Modular reduction - Double arithmetic only
' ==========================================================================
Public Function CalcFpReduceDouble(ByVal h As Double, ByVal u As Double, _
                                   ByVal modulus As Double) As Double
    ' The LOCKED reduction. Neither Mod nor \ appears, because both would apply
    ' an effective Long to an intermediate that reaches roughly 2.8e11.
    '
    ' WHY TWO CORRECTIONS ARE ENOUGH: x < 2^53, so x is an exact Double, and
    ' x / modulus <= 131 carries a relative error of at most 2^-53 - an absolute
    ' error under 1.5e-14. Fix can therefore be wrong by AT MOST ONE in either
    ' direction, and one correction in each direction absorbs exactly that.
    '
    ' Neither x nor q * modulus is ever narrowed to a Long: both exceed the
    ' signed-Long range by design, and converting either is the bug this form
    ' exists to avoid.
    Dim x As Double, q As Double, r As Double
    x = h * FP_BASE + u
    q = Fix(x / modulus)
    r = x - q * modulus
    If r >= modulus Then r = r - modulus
    If r < 0# Then r = r + modulus
    CalcFpReduceDouble = r
End Function

Public Function CalcFpDigestStream(ByVal stream As String, ByRef result As String) As Boolean
    ' Tags, lengths, the colon and the values are ALL hashed - the entire stream,
    ' UTF-16 code unit for UTF-16 code unit, nothing excluded. Both accumulators
    ' start at 1 so a stream beginning with NUL is not absorbed.
    Dim index As Long, unit As Long, h1 As Double, h2 As Double
    h1 = FP_INIT_1
    h2 = FP_INIT_2
    For index = 1 To CalcFpUtf16Length(stream)
        unit = CalcFpNormaliseCodeUnit(AscW(Mid$(stream, index, 1)))
        If unit < 0 Then Exit Function
        h1 = CalcFpReduceDouble(h1, unit, FP_MOD_1)
        h2 = CalcFpReduceDouble(h2, unit, FP_MOD_2)
    Next index
    result = CalcFpHex8(h1) & CalcFpHex8(h2)
    CalcFpDigestStream = True
End Function

Private Function CalcFpHex8(ByVal value As Double) As String
    ' Eight uppercase hex digits, produced by repeated Double division rather
    ' than by Hex$: the accumulator is a Double throughout, and handing it to a
    ' function that narrows to a Long would reintroduce exactly the conversion
    ' the reducer avoids. Only the single digit, which is 0 to 15, is narrowed.
    Dim digits As String, index As Long, place As Double, digit As Double
    Dim remaining As Double
    remaining = value
    For index = FP_HEX_WIDTH - 1 To 0 Step -1
        place = CalcFpPowerOf16(index)
        digit = Fix(remaining / place)
        remaining = remaining - digit * place
        digits = digits & Mid$(FP_HEX_DIGITS, CLng(digit) + 1, 1)
    Next index
    CalcFpHex8 = digits
End Function

Private Function CalcFpPowerOf16(ByVal exponent As Long) As Double
    Dim index As Long, running As Double
    running = 1#
    For index = 1 To exponent
        running = running * 16#
    Next index
    CalcFpPowerOf16 = running
End Function

' ==========================================================================
' Driver records
'
' FIELD ORDER IS LOCKED, and the two kinds are NOT the same shape:
'
'   COST:  S(PermanentId) S(Distribution) N(Quantity)    N(Min) N(Max)
'          [ N(MostLikely) ] N(FxToSar) N(inflation)* N(weight)*
'
'   RISK:  S(PermanentId) S(Distribution) N(Probability) N(Min) N(Max)
'          [ N(MostLikely) ] N(FxToSar) N(inflation)* N(weight)*
'
' THE OPPOSITE KIND'S MULTIPLICATIVE IDENTITY IS NOT FINGERPRINTED. A cost line
' encodes Quantity and no Probability; a risk encodes Probability and no
' Quantity. The `Quantity = 1 for risks, Probability = 1 for cost lines`
' convention belongs to the in-memory DriverFactors carry type that the
' calculation and the simulation share - it is a calculation convenience, not
' the fingerprint schema, and writing an identity 1 into the stream would put a
' field in the record that the locked grammar does not have.
'
' THE RESOLVED INFLATION-FACTOR VECTOR IS PART OF THE RECORD. It is what makes
' the fingerprint detect a changed inflation profile: without it a model whose
' referenced rates moved would hash identically to the model before the move,
' and a stale result would present itself as current. The vector is RESOLVED -
' the cumulative factor for each applied project year, not the profile name and
' not the annual rates - because that is what the calculation actually consumed.
'
' Both vectors are encoded in project-year order, inflation first and then the
' profiling weights. This encoder hashes exactly the vectors it is handed;
' whether their lengths match Applied Duration is the later resolver and check
' layer's question, not a pure encoder's.
' ==========================================================================
Public Function CalcFpBuildCostRecord(ByVal permanentId As String, ByVal distribution As String, _
                                      ByVal quantity As Double, ByVal minValue As Double, _
                                      ByVal maxValue As Double, ByVal mostLikely As Double, _
                                      ByVal includeMostLikely As Boolean, _
                                      ByVal fxToSar As Double, _
                                      ByRef inflationFactors() As Double, _
                                      ByRef weights() As Double, _
                                      ByVal decimalSeparator As String, _
                                      ByRef record As String) As Boolean
    ' `quantity` is the cost line's own Quantity, and no Probability field is
    ' emitted at all.
    CalcFpBuildCostRecord = CalcFpBuildDriverRecord(permanentId, distribution, quantity, _
                                                    minValue, maxValue, mostLikely, _
                                                    includeMostLikely, fxToSar, _
                                                    inflationFactors, weights, _
                                                    decimalSeparator, record)
End Function

Public Function CalcFpBuildRiskRecord(ByVal permanentId As String, ByVal distribution As String, _
                                      ByVal probability As Double, ByVal minValue As Double, _
                                      ByVal maxValue As Double, ByVal mostLikely As Double, _
                                      ByVal includeMostLikely As Boolean, _
                                      ByVal fxToSar As Double, _
                                      ByRef inflationFactors() As Double, _
                                      ByRef weights() As Double, _
                                      ByVal decimalSeparator As String, _
                                      ByRef record As String) As Boolean
    ' `probability` is the risk's own Probability, and no Quantity field is
    ' emitted at all.
    CalcFpBuildRiskRecord = CalcFpBuildDriverRecord(permanentId, distribution, probability, _
                                                    minValue, maxValue, mostLikely, _
                                                    includeMostLikely, fxToSar, _
                                                    inflationFactors, weights, _
                                                    decimalSeparator, record)
End Function

Private Function CalcFpBuildDriverRecord(ByVal permanentId As String, _
                                         ByVal distribution As String, _
                                         ByVal kindScalar As Double, ByVal minValue As Double, _
                                         ByVal maxValue As Double, ByVal mostLikely As Double, _
                                         ByVal includeMostLikely As Boolean, _
                                         ByVal fxToSar As Double, _
                                         ByRef inflationFactors() As Double, _
                                         ByRef weights() As Double, _
                                         ByVal decimalSeparator As String, _
                                         ByRef record As String) As Boolean
    ' ONE kind-specific scalar, in ONE position: Quantity for a cost line,
    ' Probability for a risk. The shared builder never sees both, so it cannot
    ' emit an identity for the one that does not apply.
    '
    ' THE CAPACITY IS COMPUTED, NOT GUESSED. Six fixed fields - Permanent ID,
    ' Distribution, the kind-specific scalar, Min, Max and FX - plus the two
    ' vectors, plus ONE MORE when Most Likely is present. Folding the optional
    ' field into a constant is exactly how a record that emits nine fields came
    ' to be given eight slots.
    Dim fields() As String, count As Long, index As Long, fieldCount As Long
    Dim inflationCount As Long, weightCount As Long
    inflationCount = UBound(inflationFactors) - LBound(inflationFactors) + 1
    weightCount = UBound(weights) - LBound(weights) + 1
    If inflationCount < 1 Or weightCount < 1 Then Exit Function
    fieldCount = 6 + inflationCount + weightCount
    If includeMostLikely Then fieldCount = fieldCount + 1
    ReDim fields(0 To fieldCount - 1)
    fields(0) = CalcFpCanonicalText(permanentId)
    fields(1) = CalcFpCanonicalText(distribution)
    count = 2
    If Not CalcFpNumberField(kindScalar, decimalSeparator, fields(count)) Then Exit Function
    count = count + 1
    If Not CalcFpNumberField(minValue, decimalSeparator, fields(count)) Then Exit Function
    count = count + 1
    If Not CalcFpNumberField(maxValue, decimalSeparator, fields(count)) Then Exit Function
    count = count + 1
    If includeMostLikely Then
        If Not CalcFpNumberField(mostLikely, decimalSeparator, fields(count)) Then Exit Function
        count = count + 1
    End If
    If Not CalcFpNumberField(fxToSar, decimalSeparator, fields(count)) Then Exit Function
    count = count + 1
    For index = 0 To inflationCount - 1
        If Not CalcFpNumberField(inflationFactors(LBound(inflationFactors) + index), _
                                 decimalSeparator, fields(count)) Then Exit Function
        count = count + 1
    Next index
    For index = 0 To weightCount - 1
        If Not CalcFpNumberField(weights(LBound(weights) + index), decimalSeparator, _
                                 fields(count)) Then Exit Function
        count = count + 1
    Next index
    ' The emitted count must equal the capacity the schema asked for. This turns
    ' a future field-order or schema edit into a controlled failure instead of
    ' another silent buffer mismatch, and `count` - never the array size -
    ' remains the encoded field count.
    If count <> fieldCount Then Exit Function
    CalcFpBuildDriverRecord = CalcFpEncodeRecord(fields, count, record)
End Function

Private Function CalcFpEncodeRecord(ByRef fields() As String, ByVal count As Long, _
                                    ByRef record As String) As Boolean
    ' The field count is part of the hashed stream, so a record that gained or
    ' lost a field cannot coincide with a differently-shaped one.
    Dim index As Long, prefix As String, body As String
    If Not CalcFpCanonicalInteger(count, prefix) Then Exit Function
    For index = 0 To count - 1
        body = body & fields(index)
    Next index
    record = prefix & body
    CalcFpEncodeRecord = True
End Function

Private Function CalcFpEncodeSection(ByVal sectionName As String, ByRef records() As String, _
                                     ByVal count As Long, ByRef section As String) As Boolean
    Dim index As Long, prefix As String, body As String
    If Not CalcFpCanonicalInteger(count, prefix) Then Exit Function
    For index = 0 To count - 1
        body = body & records(index)
    Next index
    section = CalcFpCanonicalText(sectionName) & prefix & body
    CalcFpEncodeSection = True
End Function

' ==========================================================================
' The stream and the digest
' ==========================================================================
Public Function CalcFpBuildFingerprint(ByRef headerFields() As String, _
                                       ByVal headerCount As Long, ByRef costIds() As String, _
                                       ByRef costRecords() As String, ByVal costCount As Long, _
                                       ByRef riskIds() As String, ByRef riskRecords() As String, _
                                       ByVal riskCount As Long, ByRef result As String) As Boolean
    ' THE ALGORITHM VERSION IS NOT A PARAMETER. FP_VERSION is projected from
    ' spec/calc_contract.yaml into modCalcContract, and letting a caller choose
    ' it would let a resolver silently select a different encoding - which is
    ' the one thing a version stamp exists to make impossible.
    CalcFpBuildFingerprint = CalcFpBuildVersionedFingerprint(FP_VERSION, headerFields, _
                                                             headerCount, costIds, costRecords, _
                                                             costCount, riskIds, riskRecords, _
                                                             riskCount, result)
End Function

Private Function CalcFpBuildVersionedFingerprint(ByVal version As Long, _
                                                 ByRef headerFields() As String, _
                                                 ByVal headerCount As Long, _
                                                 ByRef costIds() As String, _
                                                 ByRef costRecords() As String, _
                                                 ByVal costCount As Long, _
                                                 ByRef riskIds() As String, _
                                                 ByRef riskRecords() As String, _
                                                 ByVal riskCount As Long, _
                                                 ByRef result As String) As Boolean
    ' PRIVATE, and deliberately so. The version is injectable here only because
    ' the stream grammar has to be expressible for a version other than the
    ' current one - a future migration compares two encodings of the same
    ' inputs. Production has exactly one entry point, above, and it reads the
    ' generated constant.
    ' stream ::= F_S("PCCM-FP") F_I(version) section*
    '
    ' Sections are emitted in the locked order HEADER, COST, RISK - fixed, never
    ' sorted - so a later phase can append its own sections after RISK and leave
    ' the analytical subset comparable across phases.
    Dim stream As String, part As String
    Dim headerRecord As String, header(0 To 0) As String
    Dim sortedCost() As String, sortedRisk() As String
    If version < 1 Then Exit Function
    If Not CalcFpCanonicalInteger(version, part) Then Exit Function
    stream = CalcFpCanonicalText(FP_STREAM_TAG) & part

    If Not CalcFpEncodeRecord(headerFields, headerCount, headerRecord) Then Exit Function
    header(0) = headerRecord
    If Not CalcFpEncodeSection(FP_SECTION_1, header, 1, part) Then Exit Function
    stream = stream & part

    ' Records are ordered by ascending Permanent ID on UTF-16 code units - never
    ' by row, because row order is presentation, and never by digest, because
    ' that would make the ordering depend on the thing being computed.
    If Not CalcFpSortedRecords(costIds, costRecords, costCount, sortedCost) Then Exit Function
    If Not CalcFpEncodeSection(FP_SECTION_2, sortedCost, costCount, part) Then Exit Function
    stream = stream & part

    If Not CalcFpSortedRecords(riskIds, riskRecords, riskCount, sortedRisk) Then Exit Function
    If Not CalcFpEncodeSection(FP_SECTION_3, sortedRisk, riskCount, part) Then Exit Function
    stream = stream & part

    CalcFpBuildVersionedFingerprint = CalcFpDigestStream(stream, result)
End Function

Private Function CalcFpSortedRecords(ByRef ids() As String, ByRef records() As String, _
                                     ByVal count As Long, ByRef ordered() As String) As Boolean
    ' Insertion sort on a private index permutation. The caller's arrays are not
    ' reordered: the fingerprint must not have a side effect on the data it is
    ' asked to describe.
    Dim order() As Long, index As Long, probe As Long, moving As Long
    If count < 0 Then Exit Function
    If count = 0 Then
        ReDim ordered(0 To 0)
        CalcFpSortedRecords = True
        Exit Function
    End If
    ReDim order(0 To count - 1)
    For index = 0 To count - 1
        order(index) = index
    Next index
    For index = 1 To count - 1
        moving = order(index)
        probe = index - 1
        Do While probe >= 0
            If StrComp(ids(LBound(ids) + order(probe)), ids(LBound(ids) + moving), _
                       vbBinaryCompare) <= 0 Then Exit Do
            order(probe + 1) = order(probe)
            probe = probe - 1
        Loop
        order(probe + 1) = moving
    Next index
    ReDim ordered(0 To count - 1)
    For index = 0 To count - 1
        ordered(index) = records(LBound(records) + order(index))
    Next index
    CalcFpSortedRecords = True
End Function
' ==========================================================================
' STEP 10 ADDITION - THE CANONICAL DIGEST CONTINUATION
'
' EVERYTHING ABOVE THIS BANNER IS THE ACCEPTED PHASE-5 MODULE, BYTE FOR BYTE.
' The accepted digest gates hash the text before this line and still require
' the accepted literals, so "and nothing else" keeps its full meaning: not one
' accepted line may move to make room for what follows.
'
' WHY THIS EXISTS. The Phase-6 request fingerprint is the analytical
' HEADER/COST/RISK stream followed by a SIM section. `CalcFpBuildFingerprint`
' returns the analytical DIGEST and does not expose the stream that produced
' it, and the accepted Step-10A authority forbids hashing that digest as a
' field. The alternatives were all worse: rebuilding the analytical grammar in
' a second module, or writing a second polynomial hash loop.
'
' The digest IS the final pair of accumulator states - eight hex digits of h1
' then eight of h2, with no finalisation transform - so continuing from those
' states is not an approximation of appending to the stream. It IS appending to
' the stream:
'
'     ContinueDigest(DigestStream(prefix), suffix) = DigestStream(prefix & suffix)
'
' THIS CHANGES NO FINGERPRINT AUTHORITY. It is an implementation technique that
' lets Phase 6 reach the accepted hash instead of copying it. The base, the two
' moduli, the reduction, the code-unit normalisation and the hex conversion are
' the same accepted procedures above; nothing here restates one.
' ==========================================================================
Public Function CalcFpContinueDigest(ByVal priorDigest As String, ByVal suffix As String, _
                                     ByRef result As String) As Boolean
    Dim h1 As Double, h2 As Double
    Dim index As Long, unit As Long

    If CalcFpUtf16Length(priorDigest) <> FP_HEX_WIDTH + FP_HEX_WIDTH Then Exit Function
    If Not CalcFpHexValue(Mid$(priorDigest, 1, FP_HEX_WIDTH), h1) Then Exit Function
    If Not CalcFpHexValue(Mid$(priorDigest, FP_HEX_WIDTH + 1, FP_HEX_WIDTH), h2) Then Exit Function
    ' A pair of accumulator states, not an opaque label: each must be a residue
    ' of ITS OWN modulus, and h1 and h2 have different moduli. Accepting a value
    ' at or above the modulus would continue from a state the hash can never
    ' reach, and swapping the halves would silently produce a valid-looking
    ' digest for the wrong stream.
    If h1 < 0# Or h1 >= FP_MOD_1 Then Exit Function
    If h2 < 0# Or h2 >= FP_MOD_2 Then Exit Function

    ' The identical loop CalcFpDigestStream runs, from a supplied state rather
    ' than from FP_INIT_1 / FP_INIT_2. There is no second recurrence here.
    For index = 1 To CalcFpUtf16Length(suffix)
        unit = CalcFpNormaliseCodeUnit(AscW(Mid$(suffix, index, 1)))
        If unit < 0 Then Exit Function
        h1 = CalcFpReduceDouble(h1, unit, FP_MOD_1)
        h2 = CalcFpReduceDouble(h2, unit, FP_MOD_2)
    Next index

    result = CalcFpHex8(h1) & CalcFpHex8(h2)
    CalcFpContinueDigest = True
End Function

Private Function CalcFpHexValue(ByVal text As String, ByRef value As Double) As Boolean
    ' Eight uppercase hex digits to a Double. The accumulator stays a Double all
    ' the way: eight digits reach 4294967295, which is outside signed-Long range,
    ' and narrowing on the way would be exactly the conversion CalcFpReduceDouble
    ' and CalcFpHex8 exist to avoid.
    Dim index As Long, digit As Long, running As Double
    If CalcFpUtf16Length(text) <> FP_HEX_WIDTH Then Exit Function
    running = 0#
    For index = 1 To FP_HEX_WIDTH
        digit = CalcFpHexDigitValue(Mid$(text, index, 1))
        If digit < 0 Then Exit Function
        running = running * 16# + CDbl(digit)
    Next index
    value = running
    CalcFpHexValue = True
End Function

Private Function CalcFpHexDigitValue(ByVal char As String) As Long
    ' By POSITION in the accepted uppercase table, compared ORDINALLY.
    '
    ' Not CLng("&H" & text) and not any host hex parser: those are host- and
    ' locale-sensitive, and would silently accept lowercase, surrounding
    ' whitespace, a sign or an 0x prefix. A digest is an identity, so a
    ' representation this module never produces must not be accepted back.
    Dim index As Long
    CalcFpHexDigitValue = -1
    If Len(char) <> 1 Then Exit Function
    For index = 1 To CalcFpUtf16Length(FP_HEX_DIGITS)
        If StrComp(Mid$(FP_HEX_DIGITS, index, 1), char, vbBinaryCompare) = 0 Then
            CalcFpHexDigitValue = index - 1
            Exit Function
        End If
    Next index
End Function
