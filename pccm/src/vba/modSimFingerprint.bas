Attribute VB_Name = "modSimFingerprint"
Option Explicit

' ==========================================================================
' modSimFingerprint - the Phase-6 request fingerprint and the iteration-result
' digest, and nothing else.
'
' This module is FRAMING AND ORDER. It decides which fields go into a stream,
' in which order, with which tags, and how many records there are. It does not
' decide how a field is encoded, how a code unit is normalised, how the hash
' recurs, how the reduction works or how a digest is spelled in hexadecimal:
' every one of those belongs to the accepted modCalcFingerprint and is reached
' through its public surface. THERE IS NO SECOND HASH IN THIS MODULE.
'
' --------------------------------------------------------------------------
' THE REQUEST FINGERPRINT IS A CONTINUATION, NOT A COMPOSITION
' --------------------------------------------------------------------------
' The accepted Step-10A authority states the request stream as
'
'     PCCM-FP  FP_VERSION  HEADER  COST  RISK  |  SIM
'
' and forbids hashing the analytical DIGEST as a field. The analytical digest
' IS the pair of accumulator states that stream reached - eight hex digits of
' h1 then eight of h2, with no finalisation transform - so continuing the hash
' from it is byte-for-byte the same as appending to the stream itself.
'
' `analyticalFingerprint` therefore arrives here as HASH STATE. It is never
' encoded, never framed as a text field and never appears in the stream. The
' controls in the validation suite plant both mistakes and require refusal.
'
' WHAT THIS MODULE DOES NOT PROVE: that the analytical fingerprint it is handed
' is CURRENT. A syntactically valid digest is not evidence that the model has
' not moved since it was computed. Step 11 must recompute the analytical
' fingerprint from the current resolved inputs and pass THAT, never a stored
' last-successful value, unless the current one has first been proved identical.
'
' --------------------------------------------------------------------------
' THE RESULT DIGEST IS STREAMED
' --------------------------------------------------------------------------
' A run may retain 1,048,543 iterations. Concatenating that whole canonical
' RESULT stream into one VBA String before hashing it would be tens of
' megabytes of String for no reason. Instead the small framing prefix is
' digested once and each record is folded in as it is built, so the canonical
' text alive at any moment is ONE record.
'
' --------------------------------------------------------------------------
' SCOPE
' --------------------------------------------------------------------------
' No worksheet, no workbook, no Application, no environment, no file, no clock.
' No RNG, no sampler, no statistic, no simulation loop over drivers, no
' contingency, no _SimData, no Results, no run_id, no state or attempt
' derivation and no endpoint. It hashes values it is handed, and nothing more.
'
' NOTHING IN THIS MODULE HAS BEEN EXECUTED. It is source, submitted for review.
' Parity with the reference implementation is proven on Windows at Gate B.
' ==========================================================================

' ==========================================================================
' The request fingerprint
' ==========================================================================
Public Function SimFpBuildRequestFingerprint(ByVal analyticalFingerprint As String, _
                                             ByVal iterations As Long, _
                                             ByVal seedMode As String, _
                                             ByVal hasSuppliedSeed As Boolean, _
                                             ByVal suppliedSeed As Long, _
                                             ByRef result As String, _
                                             ByRef detail As String) As Boolean
    Dim suffix As String, candidate As String

    detail = vbNullString
    If Not SimFpValidateRequest(iterations, seedMode, hasSuppliedSeed, suppliedSeed, detail) Then
        Exit Function
    End If
    If Not SimFpRequestSuffix(iterations, seedMode, hasSuppliedSeed, suppliedSeed, _
                              suffix, detail) Then
        Exit Function
    End If
    ' THE ANALYTICAL FINGERPRINT IS THE PRIOR HASH STATE. It is handed to the
    ' accepted continuation, never to a field encoder.
    If Not modCalcFingerprint.CalcFpContinueDigest(analyticalFingerprint, suffix, candidate) Then
        detail = "request fingerprint: the analytical fingerprint is not a canonical digest"
        Exit Function
    End If

    result = candidate
    SimFpBuildRequestFingerprint = True
End Function

' The SIM extension, exactly as Step-10A locked it. Every count, name and
' version below is a projected constant; not one is spelled here.
Private Function SimFpRequestSuffix(ByVal iterations As Long, ByVal seedMode As String, _
                                    ByVal hasSuppliedSeed As Boolean, _
                                    ByVal suppliedSeed As Long, _
                                    ByRef suffix As String, ByRef detail As String) As Boolean
    Dim built As String, encoded As String

    built = modCalcFingerprint.CalcFpCanonicalText(SIM_REQUEST_SECTION)
    If Not modCalcFingerprint.CalcFpCanonicalInteger(SIM_REQUEST_RECORD_COUNT, encoded) Then
        detail = "request fingerprint: the record count is not encodable"
        Exit Function
    End If
    built = built & encoded

    If hasSuppliedSeed Then
        If Not modCalcFingerprint.CalcFpCanonicalInteger(SIM_REQUEST_FIELD_COUNT_FIXED, _
                                                         encoded) Then
            detail = "request fingerprint: the FIXED field count is not encodable"
            Exit Function
        End If
    Else
        If Not modCalcFingerprint.CalcFpCanonicalInteger(SIM_REQUEST_FIELD_COUNT_AUTO, _
                                                         encoded) Then
            detail = "request fingerprint: the AUTO field count is not encodable"
            Exit Function
        End If
    End If
    built = built & encoded

    If Not modCalcFingerprint.CalcFpCanonicalInteger(iterations, encoded) Then
        detail = "request fingerprint: the iteration count is not encodable"
        Exit Function
    End If
    built = built & encoded
    built = built & modCalcFingerprint.CalcFpCanonicalText(seedMode)

    ' AUTO EMITS NO SEED FIELD AT ALL. Not a zero, not a blank, not a null and
    ' not the previous effective seed: the field does not exist, which is why
    ' two AUTO runs of the same question share one request fingerprint.
    If hasSuppliedSeed Then
        If Not modCalcFingerprint.CalcFpCanonicalInteger(suppliedSeed, encoded) Then
            detail = "request fingerprint: the supplied seed is not encodable"
            Exit Function
        End If
        built = built & encoded
    End If

    If Not modCalcFingerprint.CalcFpCanonicalInteger(SIM_RNG_VERSION, encoded) Then
        detail = "request fingerprint: the generator version is not encodable"
        Exit Function
    End If
    built = built & encoded
    If Not modCalcFingerprint.CalcFpCanonicalInteger(SIM_METHOD_VERSION, encoded) Then
        detail = "request fingerprint: the method version is not encodable"
        Exit Function
    End If
    built = built & encoded

    suffix = built
    SimFpRequestSuffix = True
End Function

Private Function SimFpValidateRequest(ByVal iterations As Long, ByVal seedMode As String, _
                                      ByVal hasSuppliedSeed As Boolean, _
                                      ByVal suppliedSeed As Long, _
                                      ByRef detail As String) As Boolean
    Dim isAuto As Boolean, isFixed As Boolean

    If iterations < SIM_MIN_ITERATIONS Then
        detail = "request fingerprint: fewer iterations than the business minimum"
        Exit Function
    End If
    If iterations > SIM_MAX_ITERATIONS Then
        detail = "request fingerprint: more iterations than the technical ceiling"
        Exit Function
    End If

    ' ORDINAL comparison against the projected labels. A case-insensitive match
    ' would let "auto" and "Auto" reach the stream as themselves and hash to
    ' something the accepted grammar never produces.
    isAuto = False
    isFixed = False
    If StrComp(seedMode, SIM_SEED_MODE_AUTO, vbBinaryCompare) = 0 Then
        isAuto = True
    End If
    If StrComp(seedMode, SIM_SEED_MODE_FIXED, vbBinaryCompare) = 0 Then
        isFixed = True
    End If
    If Not isAuto And Not isFixed Then
        detail = "request fingerprint: an unknown seed mode"
        Exit Function
    End If

    ' The flag and the mode must AGREE. Either one alone would let a FIXED
    ' request omit its seed, or an AUTO request smuggle one in.
    If isAuto And hasSuppliedSeed Then
        detail = "request fingerprint: an AUTO request carries no supplied seed"
        Exit Function
    End If
    If isFixed And Not hasSuppliedSeed Then
        detail = "request fingerprint: a FIXED request needs its supplied seed"
        Exit Function
    End If
    If isFixed Then
        ' The domain is the input contract's, read through the projection and
        ' not restated here.
        If suppliedSeed < SIM_SEED_MIN Or suppliedSeed > SIM_SEED_MAX Then
            detail = "request fingerprint: the supplied seed is outside its accepted domain"
            Exit Function
        End If
    End If

    SimFpValidateRequest = True
End Function

' ==========================================================================
' The result digest
'
' The PRODUCTION surface takes no version. SIM_METHOD_VERSION is the version,
' and a caller that could choose another could produce a digest that claims a
' method the run did not use.
' ==========================================================================
Public Function SimFpResultDigest(ByRef totalNominal() As Double, ByRef totalPv() As Double, _
                                  ByVal sampleCount As Long, _
                                  ByVal decimalSeparator As String, _
                                  ByRef result As String, ByRef detail As String) As Boolean
    SimFpResultDigest = SimFpVersionedResultDigest(SIM_METHOD_VERSION, totalNominal, totalPv, _
                                                   sampleCount, decimalSeparator, result, detail)
End Function

' The version travels as a parameter ONLY here, and this procedure is Private.
' The accepted corpus retains a version-2 framing vector, and a vector that no
' test can reach is not a vector.
Private Function SimFpVersionedResultDigest(ByVal methodVersion As Long, _
                                            ByRef totalNominal() As Double, _
                                            ByRef totalPv() As Double, _
                                            ByVal sampleCount As Long, _
                                            ByVal decimalSeparator As String, _
                                            ByRef result As String, _
                                            ByRef detail As String) As Boolean
    Dim prefix As String, encoded As String, record As String
    Dim running As String, folded As String
    Dim nominalExtent As Long, pvExtent As Long
    Dim offset As Long

    detail = vbNullString
    If sampleCount < 0 Then
        detail = "result digest: a negative retained sample count"
        Exit Function
    End If
    If methodVersion < 1 Then
        detail = "result digest: the method version must be positive"
        Exit Function
    End If

    ' A ZERO-COUNT CARRIER IS NEVER INSPECTED. VBA has no zero-length dynamic
    ' array, so a caller with nothing retained may hand over an array that was
    ' never sized; the accepted empty framing vector is exactly this case.
    If sampleCount > 0 Then
        If Not SimFpRetainedExtent(totalNominal, totalPv, nominalExtent, pvExtent) Then
            detail = "result digest: the retained carrier is not allocated"
            Exit Function
        End If
        If nominalExtent <> sampleCount Then
            detail = "result digest: the retained nominal carrier is not the claimed length"
            Exit Function
        End If
        If pvExtent <> sampleCount Then
            detail = "result digest: the retained PV carrier is not the claimed length"
            Exit Function
        End If
    End If

    prefix = modCalcFingerprint.CalcFpCanonicalText(SIM_DIGEST_STREAM_TAG)
    If Not modCalcFingerprint.CalcFpCanonicalInteger(methodVersion, encoded) Then
        detail = "result digest: the method version is not encodable"
        Exit Function
    End If
    prefix = prefix & encoded
    prefix = prefix & modCalcFingerprint.CalcFpCanonicalText(SIM_DIGEST_SECTION)
    If Not modCalcFingerprint.CalcFpCanonicalInteger(sampleCount, encoded) Then
        detail = "result digest: the record count is not encodable"
        Exit Function
    End If
    prefix = prefix & encoded

    If Not modCalcFingerprint.CalcFpDigestStream(prefix, running) Then
        detail = "result digest: the framing prefix could not be digested"
        Exit Function
    End If

    ' ONE RECORD OF CANONICAL TEXT ALIVE AT A TIME. The retained arrays are read
    ' in their ORIGINAL order and are never sorted: the digest identifies the
    ' retained SEQUENCE, and sorting would make two runs that produced the same
    ' multiset in different orders indistinguishable.
    For offset = 0 To sampleCount - 1
        If Not SimFpDigestRecord(offset, totalNominal(LBound(totalNominal) + offset), _
                                 totalPv(LBound(totalPv) + offset), decimalSeparator, _
                                 record, detail) Then
            Exit Function
        End If
        If Not modCalcFingerprint.CalcFpContinueDigest(running, record, folded) Then
            detail = "result digest: the running digest could not be continued at iteration " & _
                     CStr(SIM_DIGEST_INDEX_ORIGIN + offset)
            Exit Function
        End If
        running = folded
    Next offset

    ' COMMIT LAST. A failure at record 50,000 leaves the caller's output exactly
    ' as it was; the working digest above is local and is discarded with it.
    result = running
    SimFpVersionedResultDigest = True
End Function

' One record: F_I(field count) F_I(iteration index) F_N(nominal) F_N(PV).
'
' THE INDEX IS LOGICAL. Step 8 retains iteration i at physical element i - 1, and
' the digest's index origin is 1, so element `offset` is iteration
' `SIM_DIGEST_INDEX_ORIGIN + offset` whatever the carrier's physical LBound
' happens to be. Encoding a physical index would make the digest depend on how
' the array was declared.
Private Function SimFpDigestRecord(ByVal offset As Long, ByVal nominal As Double, _
                                   ByVal pv As Double, ByVal decimalSeparator As String, _
                                   ByRef record As String, ByRef detail As String) As Boolean
    Dim built As String, encoded As String

    If Not IsUsableDouble(nominal) Then
        detail = "result digest: the retained nominal total at iteration " & _
                 CStr(SIM_DIGEST_INDEX_ORIGIN + offset) & " is not a finite Double"
        Exit Function
    End If
    If Not IsUsableDouble(pv) Then
        detail = "result digest: the retained PV total at iteration " & _
                 CStr(SIM_DIGEST_INDEX_ORIGIN + offset) & " is not a finite Double"
        Exit Function
    End If

    If Not modCalcFingerprint.CalcFpCanonicalInteger(SIM_DIGEST_FIELD_COUNT, encoded) Then
        detail = "result digest: the field count is not encodable"
        Exit Function
    End If
    built = encoded
    If Not modCalcFingerprint.CalcFpCanonicalInteger(SIM_DIGEST_INDEX_ORIGIN + offset, _
                                                     encoded) Then
        detail = "result digest: the iteration index is not encodable"
        Exit Function
    End If
    built = built & encoded
    If Not modCalcFingerprint.CalcFpNumberField(nominal, decimalSeparator, encoded) Then
        detail = "result digest: the retained nominal total is not canonically encodable"
        Exit Function
    End If
    built = built & encoded
    If Not modCalcFingerprint.CalcFpNumberField(pv, decimalSeparator, encoded) Then
        detail = "result digest: the retained PV total is not canonically encodable"
        Exit Function
    End If
    built = built & encoded

    record = built
    SimFpDigestRecord = True
End Function

' The physical extent of both retained carriers, read under a SCOPED error
' handler - the same discipline modCalcFactors uses for its arithmetic and
' modSimStats uses for the ladder carrier. `LBound` on an array that was never
' sized raises 9, so the one place that reads a bound of an unproven carrier is
' this procedure. There is no `On Error Resume Next` here or anywhere in this
' module.
Private Function SimFpRetainedExtent(ByRef totalNominal() As Double, _
                                     ByRef totalPv() As Double, _
                                     ByRef nominalExtent As Long, _
                                     ByRef pvExtent As Long) As Boolean
    On Error GoTo Unallocated
    nominalExtent = UBound(totalNominal) - LBound(totalNominal) + 1
    pvExtent = UBound(totalPv) - LBound(totalPv) + 1
    On Error GoTo 0
    SimFpRetainedExtent = True
    Exit Function
Unallocated:
    On Error GoTo 0
    SimFpRetainedExtent = False
End Function
