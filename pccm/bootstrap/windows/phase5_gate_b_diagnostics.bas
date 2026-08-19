Attribute VB_Name = "modPhase5GateBDiagnostics"
Option Explicit

' ==========================================================================
' PCCM Phase 5 Gate B - TRANSIENT DIAGNOSTIC MODULE
'
' THIS IS NOT A PRODUCTION MODULE, AND IT MUST NEVER BECOME ONE.
'
'   * it is NOT listed in stage_b_manifest.json
'   * it is NOT imported by build_stage_b.ps1
'   * it creates no button and declares no PCCM_ endpoint
'   * it is imported into the DISPOSABLE functional-test workbook only, and
'     only AFTER scenario A1 has proved the production VBA project compiles
'   * it is removed again as soon as the direct-vector section finishes, and
'     the module inventory is re-asserted at exactly 15 before anything else
'   * no accepted workbook is ever saved with it installed
'
' WHY IT EXISTS. Gate A could prove what the fingerprint source SAYS. It could
' not run it. The locked vectors of plan section 24.1 - the ten canonical numeric
' encodings, both decimal separators, the four Double-only reductions, the
' UTF-16 set and the 366-unit reference stream - have to be evaluated by REAL
' VBA on REAL Excel, and the only way to reach a Function from Application.Run
' is through a Public Sub or Function in a standard module. The production
' modules deliberately expose six PCCM_ endpoints and nothing else, so the
' vectors are reached through a transient wrapper rather than by widening the
' accepted production surface.
'
' WHAT IT MAY AND MAY NOT DO. Every procedure here is a THIN WRAPPER over an
' already-Public accepted helper. It computes no canonical form of its own, no
' digest of its own and no statistic of its own: it marshals arguments in,
' calls the accepted authority, and marshals the answer out as a String that
' Application.Run can return. Nothing here is an oracle - every EXPECTED value
' comes from build/phase5_cases.json on the PowerShell side.
'
' RETURN CONVENTION. Every function returns a String:
'
'       "OK|<value>"    the accepted helper returned True / succeeded
'       "FAIL|<why>"    the accepted helper returned False, or an argument was
'                       rejected before it was called
'
' A Boolean-returning helper that fails must never be reported as an empty
' string: a blank compares equal to a blank, and a vector that silently
' produced nothing would pass a careless comparison.
' ==========================================================================

Private Const GBD_OK As String = "OK|"
Private Const GBD_FAIL As String = "FAIL|"

' ==========================================================================
' Identity
' ==========================================================================
Public Function GBD_Ping() As String
    ' Proves the transient module is present and callable, and gives the harness
    ' a single name to assert the REMOVAL of afterwards.
    GBD_Ping = GBD_OK & "modPhase5GateBDiagnostics"
End Function

' ==========================================================================
' 8.1 / 8.2  Canonical numeric encoding, and separator injection
' ==========================================================================
Public Function GBD_CanonicalNumber(ByVal value As Double, _
                                    ByVal decimalSeparator As String) As String
    ' modCalcFingerprint.CalcFpCanonicalNumber, unmodified. The separator is
    ' INJECTED as the accepted encoder's own argument - no regional setting is
    ' read or changed, Application.International is never touched, and
    ' UseSystemSeparators is never set.
    Dim text As String
    If modCalcFingerprint.CalcFpCanonicalNumber(value, decimalSeparator, text) Then
        GBD_CanonicalNumber = GBD_OK & text
    Else
        GBD_CanonicalNumber = GBD_FAIL & "CalcFpCanonicalNumber returned False"
    End If
End Function

Public Function GBD_CanonicalNumberConstructed(ByVal label As String, _
                                               ByVal decimalSeparator As String) As String
    ' THE TWO EXTREME VECTORS, BUILT ON TARGET.
    '
    ' MAX_DOUBLE and the minimum subnormal are the two values a COM Double
    ' marshalling round trip is most likely to disturb. This entry point removes
    ' the marshalling from the question entirely by constructing the value inside
    ' VBA and then calling the SAME accepted encoder. The harness runs both this
    ' and GBD_CanonicalNumber for those two labels and requires both to equal the
    ' fixture, so a marshalling fault is reported as itself rather than as an
    ' encoder defect.
    '
    ' The construction is exact, not approximate:
    '   minimum subnormal - 1 halved 1074 times. Every intermediate is a power of
    '                       two and therefore exact, including the subnormal tail;
    '                       the last halving lands on 2^-1074 with no rounding.
    '   MAX_DOUBLE        - taken from the accepted kernel constant, not retyped.
    Dim value As Double
    Select Case label
        Case "MAX_DOUBLE"
            value = MAX_DOUBLE
        Case "minimum subnormal"
            value = GBD_MinSubnormal()
        Case Else
            GBD_CanonicalNumberConstructed = GBD_FAIL & "no on-target construction for " & label
            Exit Function
    End Select
    GBD_CanonicalNumberConstructed = GBD_CanonicalNumber(value, decimalSeparator)
End Function

Private Function GBD_MinSubnormal() As Double
    Dim value As Double, halving As Long
    value = 1#
    For halving = 1 To 1074
        value = value / 2#
    Next halving
    GBD_MinSubnormal = value
End Function

Public Function GBD_ConstructedValueText(ByVal label As String) As String
    ' The constructed value's own canonical text, so the harness can show WHICH
    ' value was built when a construction and a marshalled argument disagree.
    Dim value As Double
    Select Case label
        Case "MAX_DOUBLE": value = MAX_DOUBLE
        Case "minimum subnormal": value = GBD_MinSubnormal()
        Case Else
            GBD_ConstructedValueText = GBD_FAIL & "unknown label"
            Exit Function
    End Select
    GBD_ConstructedValueText = GBD_CanonicalNumber(value, ".")
End Function

' ==========================================================================
' 8.3  The Double-only reducer
' ==========================================================================
Public Function GBD_ReduceDouble(ByVal h As Double, ByVal u As Double, _
                                 ByVal modulus As Double) As String
    ' modCalcFingerprint.CalcFpReduceDouble, on REAL VBA Double arithmetic. The
    ' remainder is returned as its canonical text through the accepted encoder,
    ' so no PowerShell number formatting stands between the result and the
    ' comparison - and PowerShell never computes a reduction of its own.
    Dim remainder As Double, text As String
    remainder = modCalcFingerprint.CalcFpReduceDouble(h, u, modulus)
    If modCalcFingerprint.CalcFpCanonicalNumber(remainder, ".", text) Then
        GBD_ReduceDouble = GBD_OK & text
    Else
        GBD_ReduceDouble = GBD_FAIL & "the remainder has no canonical form"
    End If
End Function

' ==========================================================================
' 8.4  UTF-16 behaviour
' ==========================================================================
Public Function GBD_TextFromUnits(ByVal unitList As String) As String
    ' Builds a String from a comma-separated list of UTF-16 CODE UNITS, exactly
    ' as phase5_cases.json states them. Nothing about the text is assumed on the
    ' PowerShell side: a surrogate pair is two units in the fixture and two
    ' ChrW$ calls here, so the vector cannot be reshaped by a string literal or
    ' by the console encoding on the way in.
    Dim parts() As String, index As Long, unit As Long, built As String
    If Len(unitList) = 0 Then
        GBD_TextFromUnits = GBD_FAIL & "empty unit list"
        Exit Function
    End If
    parts = Split(unitList, ",")
    For index = LBound(parts) To UBound(parts)
        unit = CLng(Trim$(parts(index)))
        If unit < 0 Or unit > 65535 Then
            GBD_TextFromUnits = GBD_FAIL & "code unit out of range: " & CStr(unit)
            Exit Function
        End If
        built = built & ChrW$(GBD_SignedUnit(unit))
    Next index
    GBD_TextFromUnits = GBD_OK & built
End Function

Private Function GBD_SignedUnit(ByVal unit As Long) As Long
    ' ChrW$ takes a signed 16-bit argument, so a unit above U+7FFF has to be
    ' handed over in its negative form. This is the INVERSE of the normalisation
    ' CalcFpNormaliseCodeUnit performs on the way back, and proving the round
    ' trip is part of what the UTF-16 vectors are for.
    If unit > 32767 Then
        GBD_SignedUnit = unit - 65536
    Else
        GBD_SignedUnit = unit
    End If
End Function

Public Function GBD_Utf16Length(ByVal unitList As String) As String
    ' modCalcFingerprint.CalcFpUtf16Length over the reconstructed text: the count
    ' of CODE UNITS, which is two for a non-BMP character.
    Dim text As String
    text = GBD_Unwrap(GBD_TextFromUnits(unitList))
    If Len(text) = 0 And Len(unitList) > 0 Then
        GBD_Utf16Length = GBD_FAIL & "the text could not be reconstructed"
        Exit Function
    End If
    GBD_Utf16Length = GBD_OK & CStr(modCalcFingerprint.CalcFpUtf16Length(text))
End Function

Public Function GBD_RawAscW(ByVal unitList As String, ByVal position As Long) As String
    ' The RAW AscW result at one position - Integer, and therefore SIGNED. The
    ' fixture states the signed value for every unit, and a unit above U+7FFF
    ' must come back negative here or the normalisation below is proving nothing.
    Dim text As String
    text = GBD_Unwrap(GBD_TextFromUnits(unitList))
    If position < 1 Or position > Len(text) Then
        GBD_RawAscW = GBD_FAIL & "position out of range"
        Exit Function
    End If
    GBD_RawAscW = GBD_OK & CStr(CLng(AscW(Mid$(text, position, 1))))
End Function

Public Function GBD_NormaliseCodeUnit(ByVal rawUnit As Long) As String
    ' modCalcFingerprint.CalcFpNormaliseCodeUnit: the signed AscW result turned
    ' back into 0..65535. Given the fixture's signed_ascw, this must return the
    ' fixture's code_units.
    Dim normalised As Long
    normalised = modCalcFingerprint.CalcFpNormaliseCodeUnit(rawUnit)
    If normalised < 0 Then
        GBD_NormaliseCodeUnit = GBD_FAIL & "the unit was rejected as out of range"
        Exit Function
    End If
    GBD_NormaliseCodeUnit = GBD_OK & CStr(normalised)
End Function

Public Function GBD_CanonicalTextField(ByVal unitList As String) As String
    ' modCalcFingerprint.CalcFpCanonicalText. The length prefix must be the
    ' UTF-16 UNIT count, so "S2:" precedes a single non-BMP character and "S3:"
    ' precedes "A" plus that character.
    Dim text As String
    text = GBD_Unwrap(GBD_TextFromUnits(unitList))
    GBD_CanonicalTextField = GBD_OK & modCalcFingerprint.CalcFpCanonicalText(text)
End Function

' ==========================================================================
' 8.5  The complete reference stream
' ==========================================================================
Public Function GBD_StreamLength(ByVal stream As String) As String
    GBD_StreamLength = GBD_OK & CStr(modCalcFingerprint.CalcFpUtf16Length(stream))
End Function

Public Function GBD_DigestStream(ByVal stream As String) As String
    ' modCalcFingerprint.CalcFpDigestStream over the whole locked stream. The
    ' harness supplies the stream from phase5_cases.json and asserts both the
    ' 366-unit count and the digest; asserting the digest alone would pass over
    ' a stream that arrived truncated.
    Dim digest As String
    If modCalcFingerprint.CalcFpDigestStream(stream, digest) Then
        GBD_DigestStream = GBD_OK & digest
    Else
        GBD_DigestStream = GBD_FAIL & "CalcFpDigestStream returned False"
    End If
End Function

Public Function GBD_ProbeDigest(ByVal unitLists As String) As String
    ' THE DELIMITER-HOSTILE COLLISION PROBES (plan case 27).
    '
    ' Each probe is one single-record section of text fields, framed as
    '     section ::= S(name) I(record_count) record
    '     record  ::= I(field_count) field*
    ' with the probe section name "X".
    '
    ' The FIELDS and the DIGEST are the accepted public authorities -
    ' CalcFpCanonicalText, CalcFpCanonicalInteger and CalcFpDigestStream. Only
    ' the two-line concatenation above is restated here, because
    ' CalcFpEncodeSection and CalcFpEncodeRecord are Private and Gate B may not
    ' reopen production visibility for a diagnostic. The expected digests come
    ' from phase5_cases.json, so a wrong framing here fails loudly rather than
    ' agreeing with itself.
    '
    ' `unitLists` is a ';'-separated list of ','-separated code-unit lists, one
    ' per field. The probe values contain ':', NUL, unit separator and newline
    ' precisely because those are the characters that would break a naive
    ' encoding, so they are never passed as literal text.
    Dim fields() As String, index As Long
    Dim recordBody As String, fieldCount As String, recordCount As String
    Dim stream As String, digest As String, value As String

    If Len(unitLists) = 0 Then
        GBD_ProbeDigest = GBD_FAIL & "empty probe"
        Exit Function
    End If
    fields = Split(unitLists, ";")

    For index = LBound(fields) To UBound(fields)
        value = GBD_Unwrap(GBD_TextFromUnits(fields(index)))
        recordBody = recordBody & modCalcFingerprint.CalcFpCanonicalText(value)
    Next index

    If Not modCalcFingerprint.CalcFpCanonicalInteger( _
            UBound(fields) - LBound(fields) + 1, fieldCount) Then
        GBD_ProbeDigest = GBD_FAIL & "the field count has no canonical form"
        Exit Function
    End If
    If Not modCalcFingerprint.CalcFpCanonicalInteger(1, recordCount) Then
        GBD_ProbeDigest = GBD_FAIL & "the record count has no canonical form"
        Exit Function
    End If

    stream = modCalcFingerprint.CalcFpCanonicalText("X") & recordCount & _
             fieldCount & recordBody
    If modCalcFingerprint.CalcFpDigestStream(stream, digest) Then
        GBD_ProbeDigest = GBD_OK & digest
    Else
        GBD_ProbeDigest = GBD_FAIL & "CalcFpDigestStream returned False"
    End If
End Function

' ==========================================================================
' Plan case 28 - a naive overflow with a representable result
' ==========================================================================
Public Function GBD_ConvexStatistic(ByVal statistic As String, ByVal a As Double, _
                                    ByVal b As Double, ByVal c As Double) As String
    ' modCalcAnalytical's accepted convex statistics, on REAL VBA arithmetic.
    ' The point of the vector is that (a + b + c) overflows a Double while the
    ' mean does not, so the exact rescue tier has to carry it. Nothing is
    ' recomputed here; the expected results come from phase5_cases.json.
    Dim result As Double, text As String, ok As Boolean
    Select Case statistic
        Case "triangular_mean": ok = modCalcAnalytical.TriangularMean(a, b, c, result)
        Case "beta_pert_mean":  ok = modCalcAnalytical.PertMean(a, b, c, result)
        Case "midpoint":        ok = modCalcAnalytical.UniformMean(a, b, result)
        Case Else
            GBD_ConvexStatistic = GBD_FAIL & "unknown statistic " & statistic
            Exit Function
    End Select
    If Not ok Then
        GBD_ConvexStatistic = GBD_FAIL & statistic & " returned False"
        Exit Function
    End If
    If modCalcFingerprint.CalcFpCanonicalNumber(result, ".", text) Then
        GBD_ConvexStatistic = GBD_OK & text
    Else
        GBD_ConvexStatistic = GBD_FAIL & "the result has no canonical form"
    End If
End Function

' ==========================================================================
' The Phase-5 numerical checker, called directly on real VBA
' ==========================================================================
Public Function GBD_CheckBaseAfterStart(ByVal baseYear As Long, ByVal startYear As Long, _
                                        ByVal duration As Long, _
                                        ByVal discountRate As Double) As String
    ' PLAN SECTION 18: "Base Year later than Start Year".
    '
    ' WHY THIS EXISTS RATHER THAN A WORKBOOK MUTATION. Entering Base > Start and
    ' calling PCCM_ApplyTimeline does NOT reach this predicate: modTimeline
    ' PREVALIDATES the relationship and refuses the Apply without changing the
    ' applied timeline, so the workbook is left with entered <> applied and the
    ' next PCCM_Calculate is refused by StructuralPrerequisites with STRUCTURE
    ' CHANGE PENDING - a different predicate, with a different message, in a
    ' different module. Independent review found the Gate-B mutation claiming
    ' the modCalcCheck predicate while actually exercising the Phase-4 gate.
    '
    ' So the checker is called DIRECTLY, on a ResolvedModel built here. Both the
    ' type and modCalcCheck.CheckResolvedModel are already Public: no production
    ' visibility is reopened, and no production source is touched.
    '
    ' DriverCount is 0 deliberately. A model with no drivers still has a timeline
    ' and a discount rate, and both must hold - which is exactly the model-level
    ' predicate under test, with nothing else able to refuse first.
    Dim model As ResolvedModel
    Dim detail As String
    model.Timeline.BaseYear = baseYear
    model.Timeline.StartYear = startYear
    model.Timeline.Duration = duration
    model.Timeline.LastYear = startYear + duration - 1
    model.Timeline.DiscountRate = discountRate
    model.DriverCount = 0
    If modCalcCheck.CheckResolvedModel(model, detail) Then
        GBD_CheckBaseAfterStart = GBD_FAIL & "CheckResolvedModel ACCEPTED the model"
        Exit Function
    End If
    GBD_CheckBaseAfterStart = GBD_OK & detail
End Function

Public Function GBD_CheckTimelineAccepted(ByVal baseYear As Long, ByVal startYear As Long, _
                                          ByVal duration As Long, _
                                          ByVal discountRate As Double) As String
    ' THE CONTROL FOR THE ABOVE. The same construction with Base <= Start must be
    ' ACCEPTED, or the refusal proves only that the harness built a model
    ' modCalcCheck rejects for some other reason.
    Dim model As ResolvedModel
    Dim detail As String
    model.Timeline.BaseYear = baseYear
    model.Timeline.StartYear = startYear
    model.Timeline.Duration = duration
    model.Timeline.LastYear = startYear + duration - 1
    model.Timeline.DiscountRate = discountRate
    model.DriverCount = 0
    If modCalcCheck.CheckResolvedModel(model, detail) Then
        GBD_CheckTimelineAccepted = GBD_OK & "accepted"
        Exit Function
    End If
    GBD_CheckTimelineAccepted = GBD_FAIL & detail
End Function

' ==========================================================================
' Shared
' ==========================================================================
Private Function GBD_Unwrap(ByVal reply As String) As String
    ' The value out of an "OK|..." reply. A "FAIL|..." reply unwraps to the empty
    ' string, which every caller above treats as a failure rather than as data.
    If Left$(reply, Len(GBD_OK)) = GBD_OK Then
        GBD_Unwrap = Mid$(reply, Len(GBD_OK) + 1)
    End If
End Function
