Attribute VB_Name = "modSimNonce"
Option Explicit

' ==========================================================================
' modSimNonce - THE PHASE-6 AUTO NONCE TRANSACTION AND ITS RECOVERY PROTOCOL
' ==========================================================================
' WHY THIS MODULE EXISTS. The AUTO nonce is the only thing a Phase-6 run spends
' that it cannot take back. Everything else a failed run touches - a candidate
' bank, an attempt row, a status cell - is either unpublished or rewritable. A
' consumed nonce is not: reissuing it would let a later AUTO run replay a
' sequence, which is the single property the nonce exists to provide.
'
' THE THREE STATES, from seeding.nonce_lifecycle.allocation_states:
'
'   PRE_ALLOCATION            the advance is known NOT to have persisted. The
'                             nonce was not consumed, nothing sampled, and a
'                             retry may legitimately attempt the same nonce.
'                             That is not reuse - nothing ever used it.
'   CONSUMED                  the advance is known to have persisted. Sampling
'                             may begin, and no retry may take this nonce.
'   PERSISTENCE_INDETERMINATE a write was attempted and the durable state can
'                             be neither confirmed nor refuted. It must NOT be
'                             called allocated and must NOT be called
'                             unconsumed. The attempt carries the durable
'                             AUTO_NONCE_INDETERMINATE marker and the next run
'                             reconciles before allocating.
'
' A `Range.Value2` that RAISES is not proof that Excel wrote nothing. That is
' the whole reason the third state exists: the binary model forced an unknown
' outcome to be mislabelled as one of the two known ones, and a run-local flag
' cannot carry the doubt across invocations.
'
' WHAT THIS MODULE DOES NOT DO. No sampling, no statistics, no fingerprint, no
' publication, no run-ID allocation, no status derivation, and it never writes
' the Last Attempt block - modSimReport owns attempt-row persistence. This
' module only READS the attempt rows it needs to reconcile a prior run.
' ==========================================================================
Public Const FAILPOINT_SIM_AFTER_NONCE As String = "Phase6AfterNoncePersisted"

' The classification an allocation attempt ended in. Returned as a scalar so no
' caller-writable state crosses the module boundary in either direction.
Public Const SIM_NONCE_STATE_NOT_APPLICABLE As String = "NOT_APPLICABLE"
Public Const SIM_NONCE_STATE_PRE_ALLOCATION As String = "PRE_ALLOCATION"
Public Const SIM_NONCE_STATE_CONSUMED As String = "CONSUMED"
Public Const SIM_NONCE_STATE_INDETERMINATE As String = "PERSISTENCE_INDETERMINATE"
Public Const SIM_NONCE_STATE_RECOVERY As String = "RECOVERY_REQUIRED"

' ==========================================================================
' THE ONE ENTRY POINT
' ==========================================================================
Public Function SimNonceAllocate(ByVal hasSuppliedSeed As Boolean, _
                                 ByVal suppliedSeed As Long, _
                                 ByRef effectiveSeed As Long, _
                                 ByRef autoNonce As Long, _
                                 ByRef identityKnown As Boolean, _
                                 ByRef state As String, _
                                 ByRef detail As String) As Boolean
    ' THE NARROW INTERFACE. Scalars only, ByRef out-parameters only: the caller's
    ' run package stays Private to modSimReport and no shared mutable context
    ' object exists in either direction.
    '
    '   effectiveSeed  the seed the engine will use, when this returns True
    '   autoNonce      the AUTO identity, meaningful only when identityKnown
    '   identityKnown  an AUTO nonce was selected; it may appear in the attempt
    '                  row for audit. NOT a claim that it was consumed.
    '   state          one of the SIM_NONCE_STATE_* classifications
    Dim seed As Long, failure As String

    effectiveSeed = 0
    autoNonce = 0
    identityKnown = False
    state = SIM_NONCE_STATE_NOT_APPLICABLE
    detail = vbNullString

    On Error GoTo AllocationFailed

    If hasSuppliedSeed Then
        ' FIXED: no counter is read, no counter is written, and the AUTO
        ' reconciliation protocol is never entered.
        effectiveSeed = suppliedSeed
    Else
        If Not ResolveNextNonce(autoNonce, state, detail) Then Exit Function
        If Not modSimRng.SimRngAutoSeedFromNonce(autoNonce, seed, detail) Then
            Exit Function
        End If
        effectiveSeed = seed
        identityKnown = True
        If Not PersistAdvance(autoNonce, state, detail) Then Exit Function
    End If

    ' THE ACCEPTED INJECTION BOUNDARY. Reached only once the advance is KNOWN
    ' consumed (or the run is FIXED), and always before any sampling, because
    ' modSimReport runs the kernels only if this returns True. Injecting here
    ' therefore proves what its name says: the nonce is spent, nothing was
    ' sampled, and the run must still record an unsuccessful attempt.
    modAppState.FailPointCheck FAILPOINT_SIM_AFTER_NONCE

    On Error GoTo 0
    SimNonceAllocate = True
    Exit Function

AllocationFailed:
    failure = Err.Description
    On Error GoTo 0
    detail = "simulation: seed allocation did not complete: " & failure & _
             ". No sampling was started."
End Function

' ==========================================================================
' Selection, including reconciliation of a prior indeterminate attempt
' ==========================================================================
Private Function ResolveNextNonce(ByRef nonce As Long, ByRef state As String, _
                                  ByRef detail As String) As Boolean
    Dim counter As Long, prior As Long, probe As String

    If Not ReadPersistedNonce(counter, probe) Then
        detail = "simulation: the AUTO nonce counter is not readable - " & probe
        state = SIM_NONCE_STATE_RECOVERY
        Exit Function
    End If

    ' ONLY the durable machine token activates reconciliation. A generic REFUSED
    ' or FAILED attempt may follow a conclusively persisted advance, and reading
    ' the two alike would lose exactly the distinction the token exists to keep.
    If StrComp(SharedText(SIM_IDENTITY_ROW_LAST_ATTEMPT_RESULT), _
               SIM_ATTEMPT_AUTO_NONCE_INDETERMINATE, vbBinaryCompare) = 0 Then
        If Not ReadPriorNonce(prior, probe) Then
            detail = "simulation: the previous attempt is recorded as " & _
                     SIM_ATTEMPT_AUTO_NONCE_INDETERMINATE & " but the nonce it " & _
                     "attempted cannot be read (" & probe & "), so the sequence " & _
                     "cannot be reconciled. Recovery is required."
            state = SIM_NONCE_STATE_RECOVERY
            Exit Function
        End If
        If counter = prior + 1 Then
            ' Resolved: the prior advance DID persist. Allocate from the counter.
        ElseIf counter = prior Then
            ' Resolved: the prior advance did NOT persist. The prior nonce was
            ' never consumed, so allocating it now is not reuse.
        Else
            detail = "simulation: the previous attempt is recorded as " & _
                     SIM_ATTEMPT_AUTO_NONCE_INDETERMINATE & " for nonce " & _
                     CStr(prior) & ", but the counter reads " & CStr(counter) & _
                     ", which is neither that nonce nor its advance. The " & _
                     "sequence is inconsistent and recovery is required."
            state = SIM_NONCE_STATE_RECOVERY
            Exit Function
        End If
    End If

    If counter >= SIM_NONCE_EXHAUSTED Then
        detail = "simulation: the AUTO nonce sequence is exhausted"
        Exit Function
    End If
    nonce = counter
    ResolveNextNonce = True
End Function

' ==========================================================================
' The advance, its verification, and one bounded reconciliation
' ==========================================================================
Private Function PersistAdvance(ByVal nonce As Long, ByRef state As String, _
                                ByRef detail As String) As Boolean
    Dim stored As Long, probe As String, failure As String

    On Error GoTo WriteRaised
    SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 = nonce + 1
    On Error GoTo 0

    If Not ReadPersistedNonce(stored, probe) Then
        PersistAdvance = Reconcile(nonce, "the advance could not be verified (" & _
                                   probe & ")", state, detail)
        Exit Function
    End If
    PersistAdvance = Classify(nonce, stored, vbNullString, state, detail)
    Exit Function

WriteRaised:
    failure = Err.Description
    On Error GoTo 0
    PersistAdvance = Reconcile(nonce, "the counter write raised (" & failure & ")", _
                               state, detail)
End Function

Private Function Reconcile(ByVal nonce As Long, ByVal cause As String, _
                           ByRef state As String, ByRef detail As String) As Boolean
    ' EXACTLY ONE observation, after the original error has been captured and
    ' the handler disarmed. No retry loop: a loop would turn one ambiguous COM
    ' failure into an unbounded series and still not decide it.
    Dim stored As Long, probe As String, failure As String

    On Error GoTo ObservationRaised
    If Not ReadPersistedNonce(stored, probe) Then
        On Error GoTo 0
        Reconcile = Indeterminate(nonce, cause, "it could not be observed (" & _
                                  probe & ")", state, detail)
        Exit Function
    End If
    On Error GoTo 0
    Reconcile = Classify(nonce, stored, cause, state, detail)
    Exit Function

ObservationRaised:
    failure = Err.Description
    On Error GoTo 0
    Reconcile = Indeterminate(nonce, cause, "observing it raised (" & failure & ")", _
                              state, detail)
End Function

Private Function Classify(ByVal nonce As Long, ByVal stored As Long, _
                          ByVal cause As String, ByRef state As String, _
                          ByRef detail As String) As Boolean
    ' THREE OBSERVATIONS, THREE ANSWERS, and no fourth. Nothing here normalises,
    ' decrements, or skips forward to make an inconvenient reading go away.
    If stored = nonce + 1 Then
        state = SIM_NONCE_STATE_CONSUMED
        If Len(cause) = 0 Then
            Classify = True
            Exit Function
        End If
        ' The advance is confirmed, so the nonce IS spent - but the operation
        ' that discovered it failed, so the run does not continue.
        detail = "simulation: " & cause & ", but the advance is confirmed " & _
                 "persisted. Nonce " & CStr(nonce) & " is CONSUMED and " & _
                 "will not be reused; no sampling was started."
        Exit Function
    End If

    If stored = nonce Then
        ' The advance did not land. Nothing was consumed, so the next run may
        ' legitimately take this nonce again - and saying otherwise would be a
        ' promise the source cannot keep.
        state = SIM_NONCE_STATE_PRE_ALLOCATION
        detail = "simulation: the AUTO nonce advance did not persist"
        If Len(cause) > 0 Then detail = detail & " (" & cause & ")"
        detail = detail & ". Nonce " & CStr(nonce) & " was NOT consumed and no " & _
                 "sampling was started; a retry may take it again."
        Exit Function
    End If

    state = SIM_NONCE_STATE_RECOVERY
    detail = "simulation: the AUTO nonce counter reads " & CStr(stored) & _
             ", which is neither the attempted nonce " & CStr(nonce) & " nor " & _
             "its advance. The sequence is inconsistent and recovery is " & _
             "required; nothing has been normalised and no sampling was started."
End Function

Private Function Indeterminate(ByVal nonce As Long, ByVal cause As String, _
                               ByVal why As String, ByRef state As String, _
                               ByRef detail As String) As Boolean
    ' THE IRREDUCIBLE CASE. Neither confirmed nor refuted, and said so.
    state = SIM_NONCE_STATE_INDETERMINATE
    detail = "simulation: " & cause & " and " & why & ". Whether nonce " & _
             CStr(nonce) & " was consumed is INDETERMINATE: it is neither " & _
             "allocated nor unconsumed, the counter is not rolled back, no " & _
             "sampling was started, and the next run must reconcile this " & _
             "attempt before allocating."
End Function

' ==========================================================================
' The minimum shared-cell access this responsibility needs. READS ONLY,
' except the counter advance itself, which is this module's own transaction.
' ==========================================================================
' NOT named ReadCounter: that name is banned repo-wide by the Phase-4 control
' test_24f, because the historical accessor of that name mapped missing, blank
' or non-numeric state to 0. This one refuses instead of converting - but the
' ban is on the accessor, not on one call site, and it is right to be.
Private Function ReadPersistedNonce(ByRef value As Long, ByRef detail As String) As Boolean
    ReadPersistedNonce = ReadShared(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE, _
                             CDbl(SIM_NONCE_FIRST_VALID), _
                             CDbl(SIM_NONCE_EXHAUSTED), value, detail)
End Function

Private Function ReadPriorNonce(ByRef value As Long, ByRef detail As String) As Boolean
    ReadPriorNonce = ReadShared(SIM_IDENTITY_ROW_LAST_ATTEMPT_AUTO_NONCE, _
                                CDbl(SIM_NONCE_FIRST_VALID), _
                                CDbl(SIM_NONCE_LAST_VALID), value, detail)
End Function

Private Function ReadShared(ByVal row As Long, ByVal minValue As Double, _
                            ByVal maxValue As Double, ByRef value As Long, _
                            ByRef detail As String) As Boolean
    Dim raw As Variant, number As Double
    raw = SharedCell(row).Value2
    If Not modWorkbook.IsWholeInRange(raw, minValue, maxValue) Then
        detail = "the stored value is not a whole number in its accepted range"
        Exit Function
    End If
    If Not modWorkbook.TryReadDouble(raw, number) Then
        detail = "the stored value is not a usable number"
        Exit Function
    End If
    value = modWorkbook.SafeLong(number)
    ReadShared = True
End Function

Private Function SharedCell(ByVal row As Long) As Range
    Set SharedCell = modWorkbook.Sh(SIM_DATA_SHEET).Range( _
        SIM_SHARED_VALUE_COLUMN & CStr(row))
End Function

Private Function SharedText(ByVal row As Long) As String
    SharedText = modWorkbook.TextOf(SharedCell(row))
End Function
