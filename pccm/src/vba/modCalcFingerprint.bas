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

' The host format that yields the locked canonical form: one digit, the host's
' decimal marker, sixteen fractional digits (17 significant digits in total),
' an uppercase E, an always-present exponent sign and at least two exponent
' digits. The digit count is the encoding authority's, mirrored here because a
' format string is not expressible as a projected constant.
Private Const FP_NUMBER_FORMAT As String = "0.0000000000000000E+00"
Private Const FP_FRACTION_DIGITS As Long = 16
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
    ' THE NORMALISATION IS POSITIONAL, AND THE MARKER IS FOUND IN THE OUTPUT.
    ' CalcFpMarkerIndex validates the scientific-notation shape and returns the
    ' index of the single mantissa marker, WHATEVER character the host formatter
    ' put there. That position is then rewritten to "." unconditionally.
    '
    ' Replacing every occurrence of a separator instead would only be safe while
    ' the separator happens not to occur elsewhere in scientific notation - and
    ' it does occur elsewhere for "E", for "+", for "-" and for every digit. By
    ' locating the marker positionally, exactly one character is rewritten and
    ' the exponent marker, the exponent sign and every digit are untouchable by
    ' construction.
    '
    ' WHY THE HOST MARKER IS NOT COMPARED AGAINST decimalSeparator. The locked
    ' acceptance case injects BOTH "." and "," on ONE host and requires the two
    ' outputs to be byte-identical. A gate demanding that the host's own marker
    ' equal the supplied separator makes that pair unsatisfiable on any single
    ' machine: whichever separator the formatter emits, the other injection is
    ' refused. The encoder does not need to trust the supplied value to discover
    ' the marker, so it does not.
    '
    ' The separator stays a parameter: it is the locked public interface, it is
    ' what the later resolver reports about its environment, and it is what Gate
    ' B injects. It is validated as one UTF-16 code unit, and no Excel object,
    ' Application setting or worksheet state is consulted to obtain it.
    Dim text As String, marker As Long, number As Double
    If CalcFpUtf16Length(decimalSeparator) <> 1 Then Exit Function
    If Not IsUsableDouble(value) Then Exit Function
    number = value
    If number = 0# Then number = 0#
    text = Format$(number, FP_NUMBER_FORMAT)
    marker = CalcFpMarkerIndex(text)
    If marker = 0 Then Exit Function
    result = Left$(text, marker - 1) & "." & Mid$(text, marker + 1)
    CalcFpCanonicalNumber = True
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

Private Function CalcFpEncodeSection(ByVal name As String, ByRef records() As String, _
                                     ByVal count As Long, ByRef section As String) As Boolean
    Dim index As Long, prefix As String, body As String
    If Not CalcFpCanonicalInteger(count, prefix) Then Exit Function
    For index = 0 To count - 1
        body = body & records(index)
    Next index
    section = CalcFpCanonicalText(name) & prefix & body
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
