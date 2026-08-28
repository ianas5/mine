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
' THE THREE ALLOCATION STATES, from seeding.nonce_lifecycle.allocation_states.
' They classify one thing: what is PHYSICALLY known about this attempt's counter
' transition.
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
'                             unconsumed.
'
' A `Range.Value2` that RAISES is not proof that Excel wrote nothing. That is
' the whole reason the third state exists: the binary model forced an unknown
' outcome to be mislabelled as one of the two known ones, and a run-local flag
' cannot carry the doubt across invocations.
'
' RECOVERY_REQUIRED IS NOT ONE OF THEM. It is an ACTION - the workbook must be
' reconciled before the next AUTO allocation - and it travels as its own
' Boolean, never as a value of the allocation axis. A cleanup problem is not a
' persistence problem, and folding the two into one scalar let a proven CONSUMED
' be overwritten and then reported as unconsumed.
'
' THE DURABLE RECOVERY AUTHORITY IS THE SIDECAR: _SimData!F21, the Pending AUTO
' Nonce, established before the counter is touched. It is NOT the attempt row
' and NOT the AUTO_NONCE_INDETERMINATE result token - that token is an audit
' record of one attempt and any later attempt, FIXED included, may overwrite it.
'
' WHAT THIS MODULE DOES NOT DO. No sampling, no statistics, no fingerprint, no
' publication, no run-ID allocation, no status derivation, and it never writes
' OR READS the Last Attempt block - modSimReport owns attempt-row persistence,
' and this module's durable recovery inputs are the Pending AUTO Nonce and the
' Next AUTO Nonce counter, nothing else.
' ==========================================================================
Public Const FAILPOINT_SIM_AFTER_NONCE As String = "Phase6AfterNoncePersisted"

' THE ALLOCATION AXIS. What is PHYSICALLY known about this attempt's counter
' transition, and nothing else. Returned as a scalar so no caller-writable state
' crosses the module boundary in either direction.
'
' `NOT_APPLICABLE` is the FIXED-mode answer: there was no AUTO transition to
' classify. The other three are the contract's allocation_states.
Public Const SIM_NONCE_STATE_NOT_APPLICABLE As String = "NOT_APPLICABLE"
Public Const SIM_NONCE_STATE_PRE_ALLOCATION As String = "PRE_ALLOCATION"
Public Const SIM_NONCE_STATE_CONSUMED As String = "CONSUMED"
Public Const SIM_NONCE_STATE_INDETERMINATE As String = "PERSISTENCE_INDETERMINATE"

' THE RECOVERY AXIS, SEPARATELY. `RECOVERY_REQUIRED` is an ACTION, not a
' physical classification, and it travels as its own Boolean out-parameter -
' never as a value of `allocationState`.
'
' WHY THE SEPARATION IS LOAD-BEARING. A cleanup problem is not a persistence
' problem. If the counter was verified at m+1 the nonce IS consumed; a later
' failure to clear the sidecar cannot un-prove that, and folding the two axes
' into one string let exactly that happen - a proven CONSUMED became
' RECOVERY_REQUIRED, and the reporter, deriving consumption by comparing that
' one string, then reported the nonce as NOT consumed.

' ==========================================================================
' THE ONE ENTRY POINT
' ==========================================================================
Public Function SimNonceAllocate(ByVal hasSuppliedSeed As Boolean, _
                                 ByVal suppliedSeed As Long, _
                                 ByRef effectiveSeed As Long, _
                                 ByRef autoNonce As Long, _
                                 ByRef identityKnown As Boolean, _
                                 ByRef allocationState As String, _
                                 ByRef recoveryRequired As Boolean, _
                                 ByRef detail As String) As Boolean
    ' THE NARROW INTERFACE. Scalars only, ByRef out-parameters only: the caller's
    ' run package stays Private to modSimReport and no shared mutable context
    ' object exists in either direction.
    '
    ' EVERY OUT-PARAMETER IS SET BEFORE ANY EXIT, success or failure. The caller
    ' copies them on both arms, so an attempt that got as far as deriving a seed
    ' records that seed even though the run failed.
    Dim seed As Long, failure As String

    effectiveSeed = 0
    autoNonce = 0
    identityKnown = False
    allocationState = SIM_NONCE_STATE_NOT_APPLICABLE
    recoveryRequired = False
    detail = vbNullString

    On Error GoTo AllocationFailed

    If hasSuppliedSeed Then
        ' FIXED: no counter, no sidecar, no reconciliation - and no failpoint.
        ' Phase6AfterNoncePersisted names a persisted AUTO advance; firing it
        ' here would inject at a boundary that does not exist.
        effectiveSeed = suppliedSeed
        On Error GoTo 0
        SimNonceAllocate = True
        Exit Function
    End If

    ' RECONCILING A PRIOR MARKER IS NOT AN ALLOCATION. A refusal here happens
    ' BEFORE this run selects an identity, so it leaves the allocation axis at
    ' NOT_APPLICABLE and raises only the recovery flag: nothing of THIS attempt
    ' became indeterminate, and the durable lock is F21, not the attempt row.
    If Not ResolveNextNonce(autoNonce, recoveryRequired, detail) Then Exit Function

    ' THE STANDING CLASSIFICATION FROM HERE ON IS PRE_ALLOCATION. A nonce has
    ' been selected but nothing has been written, so the advance is KNOWN not to
    ' have persisted. Leaving it NOT_APPLICABLE would let a refusal between here
    ' and the counter write report the FIXED-mode classification for an AUTO
    ' attempt that really does have an attempted nonce and an effective seed.
    allocationState = SIM_NONCE_STATE_PRE_ALLOCATION

    If Not modSimRng.SimRngAutoSeedFromNonce(autoNonce, seed, detail) Then
        Exit Function
    End If
    effectiveSeed = seed
    identityKnown = True

    If Not RunAllocationTransaction(autoNonce, allocationState, _
                                    recoveryRequired, detail) Then
        Exit Function
    End If

    ' THE ACCEPTED INJECTION BOUNDARY, AUTO ONLY. Reached once the advance is
    ' KNOWN consumed and the pending marker is cleared, and always before any
    ' sampling, because modSimReport runs the kernels only if this returns True.
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
' Selection, reconciling any pending transaction first
' ==========================================================================
Private Function ResolveNextNonce(ByRef nonce As Long, _
                                  ByRef recoveryRequired As Boolean, _
                                  ByRef detail As String) As Boolean
    ' THE DURABLE AUTHORITY IS THE SIDECAR, never the attempt row. The attempt
    ' row records the LAST attempt and is rewritten by any later attempt -
    ' including a FIXED one with nothing to do with this transaction - so it
    ' cannot carry a lock that has to outlive it.
    '
    ' NO ALLOCATION CLASSIFICATION IS PRODUCED HERE. Every refusal below happens
    ' BEFORE this run selects an identity or touches the counter, so THIS
    ' attempt has no counter transition to classify. It raises the recovery
    ' ACTION and nothing more; manufacturing an indeterminate persistence
    ' outcome from a prior run's unreconciled state would be a claim about a
    ' transition that never began.
    Dim counter As Long, pending As Long, hasPending As Boolean, probe As String

    If Not ReadPending(pending, hasPending, probe) Then
        detail = "simulation: the pending AUTO nonce marker is not readable - " & _
                 probe & ". Recovery is required."
        recoveryRequired = True
        Exit Function
    End If

    If Not ReadPersistedNonce(counter, probe) Then
        detail = "simulation: the AUTO nonce counter is not readable - " & probe
        If hasPending Then
            detail = detail & ". Pending nonce " & CStr(pending) & " is retained " & _
                     "and recovery is required."
        End If
        recoveryRequired = True
        Exit Function
    End If

    If hasPending Then
        If counter = pending + 1 Then
            ' Resolved: the prior advance persisted. That nonce was consumed.
            If Not ClearPending(detail) Then
                recoveryRequired = True
                Exit Function
            End If
        ElseIf counter = pending Then
            ' Resolved: the prior advance never persisted, so the pending nonce
            ' was never consumed and this run may take it.
            If Not ClearPending(detail) Then
                recoveryRequired = True
                Exit Function
            End If
        Else
            ' The sidecar is RETAINED, so every future AUTO run stays blocked
            ' until someone reconciles it. Nothing is normalised or skipped.
            detail = "simulation: AUTO nonce " & CStr(pending) & " is pending but " & _
                     "the counter reads " & CStr(counter) & ", which is neither " & _
                     "that nonce nor its advance. The pending marker is kept, no " & _
                     "allocation is made and recovery is required."
            recoveryRequired = True
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
' The write-ahead transaction: mark, advance, reconcile, clear
' ==========================================================================
Private Function RunAllocationTransaction(ByVal nonce As Long, _
                                          ByRef allocationState As String, _
                                          ByRef recoveryRequired As Boolean, _
                                          ByRef detail As String) As Boolean
    ' THE MARKER GOES FIRST, and the counter is not touched until it is durably
    ' established. Either outcome of the marker write is then safe: if it did
    ' not land the next run sees a blank sidecar and counter m; if it did land
    ' the next run sees pending m and counter m and resolves PRE_ALLOCATION.
    ' No nonce can vanish or replay because of uncertainty in THIS write.
    If Not EstablishPending(nonce, detail) Then
        ' A DEFINITE OUTCOME, not an indeterminate one. The counter was never
        ' touched, so the advance is known not to have persisted whichever way
        ' the marker write physically went.
        allocationState = SIM_NONCE_STATE_PRE_ALLOCATION
        Exit Function
    End If
    If Not PersistAdvance(nonce, allocationState, recoveryRequired, detail) Then
        Exit Function
    End If

    ' CLEANUP IS THE OTHER AXIS. By here the counter has been OBSERVED and
    ' `allocationState` records what that observation proved. A failed clear
    ' cannot un-prove it: if the counter read m+1 the nonce IS consumed, and
    ' saying otherwise because a different cell would not clear is a claim about
    ' the counter that no observation supports.
    '
    ' So the clear failure raises the recovery ACTION and leaves the physical
    ' classification exactly as it was found. The run is still unsuccessful and
    ' still samples nothing - `RunAllocationTransaction` returns False - and
    ' either physical outcome of the clear leaves the next run safe, because a
    ' marker that survives costs one more reconciliation and nothing worse.
    If Not ClearPending(detail) Then
        recoveryRequired = True
        Exit Function
    End If
    RunAllocationTransaction = True
End Function

Private Function EstablishPending(ByVal nonce As Long, ByRef detail As String) As Boolean
    Dim stored As Long, present As Boolean, probe As String, failure As String

    On Error GoTo MarkerFailed
    PendingCell.Value2 = nonce
    On Error GoTo 0

    If Not ReadPending(stored, present, probe) Then
        detail = "simulation: the pending AUTO nonce marker could not be verified (" & _
                 probe & "). The counter was NOT touched and no nonce was consumed."
        Exit Function
    End If
    If (Not present) Or stored <> nonce Then
        detail = "simulation: the pending AUTO nonce marker did not persist. The " & _
                 "counter was NOT touched and no nonce was consumed."
        Exit Function
    End If
    EstablishPending = True
    Exit Function

MarkerFailed:
    failure = Err.Description
    On Error GoTo 0
    detail = "simulation: the pending AUTO nonce marker could not be written (" & _
             failure & "). The counter was NOT touched and no nonce was consumed."
End Function

Private Function PersistAdvance(ByVal nonce As Long, _
                                ByRef allocationState As String, _
                                ByRef recoveryRequired As Boolean, _
                                ByRef detail As String) As Boolean
    ' THE WHOLE STEP IS INSIDE ONE ENVELOPE. The verification READ is a COM call
    ' too: disarming after the write would let a raised read escape without the
    ' bounded reconciliation the contract requires for a verification failure.
    Dim stored As Long, probe As String, failure As String

    On Error GoTo StepRaised
    SharedCell(SIM_IDENTITY_ROW_NEXT_AUTO_NONCE).Value2 = nonce + 1
    If Not ReadPersistedNonce(stored, probe) Then
        On Error GoTo 0
        PersistAdvance = Reconcile(nonce, "the advance could not be verified (" & _
                                   probe & ")", allocationState, recoveryRequired, _
                                   detail)
        Exit Function
    End If
    On Error GoTo 0
    PersistAdvance = Classify(nonce, stored, vbNullString, allocationState, _
                              recoveryRequired, detail)
    Exit Function

StepRaised:
    failure = Err.Description
    On Error GoTo 0
    PersistAdvance = Reconcile(nonce, "the advance raised (" & failure & ")", _
                               allocationState, recoveryRequired, detail)
End Function

Private Function Reconcile(ByVal nonce As Long, ByVal cause As String, _
                           ByRef allocationState As String, _
                           ByRef recoveryRequired As Boolean, _
                           ByRef detail As String) As Boolean
    ' EXACTLY ONE observation, after the original error has been captured and
    ' the handler disarmed. No retry loop: a loop would turn one ambiguous COM
    ' failure into an unbounded series and still not decide it.
    Dim stored As Long, probe As String, failure As String

    On Error GoTo ObservationRaised
    If Not ReadPersistedNonce(stored, probe) Then
        On Error GoTo 0
        Reconcile = Indeterminate(nonce, cause, "it could not be observed (" & _
                                  probe & ")", allocationState, detail)
        Exit Function
    End If
    On Error GoTo 0
    Reconcile = Classify(nonce, stored, cause, allocationState, recoveryRequired, _
                         detail)
    Exit Function

ObservationRaised:
    failure = Err.Description
    On Error GoTo 0
    Reconcile = Indeterminate(nonce, cause, "observing it raised (" & failure & ")", _
                              allocationState, detail)
End Function

Private Function Classify(ByVal nonce As Long, ByVal stored As Long, _
                          ByVal cause As String, ByRef allocationState As String, _
                          ByRef recoveryRequired As Boolean, _
                          ByRef detail As String) As Boolean
    ' THREE OBSERVATIONS, THREE ANSWERS, and no fourth. Nothing here normalises,
    ' decrements, or skips forward to make an inconvenient reading go away.
    '
    ' THE OBSERVATION DECIDES `allocationState`, and once decided nothing below
    ' revises it. Cleanup trouble downstream raises `recoveryRequired` instead.
    If stored = nonce + 1 Then
        allocationState = SIM_NONCE_STATE_CONSUMED
        If Len(cause) = 0 Then
            Classify = True
            Exit Function
        End If
        ' The advance is confirmed, so the nonce IS spent - but the operation
        ' that discovered it failed, so the run does not continue. The pending
        ' marker is cleared because the outcome is now definite; if that clear
        ' cannot be completed the nonce is STILL consumed and still says so.
        If Not ClearPending(detail) Then
            recoveryRequired = True
            Exit Function
        End If
        detail = "simulation: " & cause & ", but the advance is confirmed " & _
                 "persisted. Nonce " & CStr(nonce) & " is CONSUMED and " & _
                 "will not be reused; no sampling was started."
        Exit Function
    End If

    If stored = nonce Then
        ' The advance did not land. Nothing was consumed, so the next run may
        ' legitimately take this nonce again - and saying otherwise would be a
        ' promise the source cannot keep.
        allocationState = SIM_NONCE_STATE_PRE_ALLOCATION
        If Not ClearPending(detail) Then
            recoveryRequired = True
            Exit Function
        End If
        detail = "simulation: the AUTO nonce advance did not persist"
        If Len(cause) > 0 Then detail = detail & " (" & cause & ")"
        detail = detail & ". Nonce " & CStr(nonce) & " was NOT consumed and no " & _
                 "sampling was started; a retry may take it again."
        Exit Function
    End If

    ' AN IMPOSSIBLE READING, ON BOTH AXES AT ONCE. The counter is neither the
    ' attempted nonce nor its advance, so this transition cannot be classified
    ' under the accepted model - that is PERSISTENCE_INDETERMINATE, the honest
    ' physical answer - and separately the workbook needs reconciling before any
    ' further AUTO allocation, which is the recovery ACTION. Conflating the two
    ' would either hide the unreadable counter or claim a fourth physical state.
    ' THE MARKER IS RETAINED. Future AUTO runs stay blocked until reconciled.
    allocationState = SIM_NONCE_STATE_INDETERMINATE
    recoveryRequired = True
    detail = "simulation: the AUTO nonce counter reads " & CStr(stored) & _
             ", which is neither the attempted nonce " & CStr(nonce) & " nor " & _
             "its advance. Whether it was consumed cannot be classified, the " & _
             "pending marker is kept, nothing has been normalised and no " & _
             "sampling was started; recovery is required."
End Function

Private Function Indeterminate(ByVal nonce As Long, ByVal cause As String, _
                               ByVal why As String, _
                               ByRef allocationState As String, _
                               ByRef detail As String) As Boolean
    ' THE IRREDUCIBLE CASE. Neither confirmed nor refuted, and said so. The
    ' pending marker is RETAINED - it, not the attempt row, is what makes the
    ' next AUTO run reconcile before allocating.
    allocationState = SIM_NONCE_STATE_INDETERMINATE
    detail = "simulation: " & cause & " and " & why & ". Whether nonce " & _
             CStr(nonce) & " was consumed is INDETERMINATE: it is neither " & _
             "allocated nor unconsumed, the counter is not rolled back, no " & _
             "sampling was started, and the pending marker is retained so the " & _
             "next run must reconcile before allocating."
End Function

Private Function ClearPending(ByRef detail As String) As Boolean
    ' A REAL COM WRITE, not an assumption. If it cannot be completed the
    ' transaction is not clean and this run must not sample; either physical
    ' outcome still leaves the next run safe, because a marker that survives
    ' simply causes one more reconciliation.
    Dim present As Boolean, stored As Long, probe As String, failure As String

    On Error GoTo ClearRaised
    PendingCell.ClearContents
    On Error GoTo 0

    If Not ReadPending(stored, present, probe) Then
        detail = "simulation: the pending AUTO nonce marker could not be read back " & _
                 "after clearing (" & probe & "); no sampling was started."
        Exit Function
    End If
    If present Then
        detail = "simulation: the pending AUTO nonce marker did not clear; no " & _
                 "sampling was started."
        Exit Function
    End If
    ClearPending = True
    Exit Function

ClearRaised:
    failure = Err.Description
    On Error GoTo 0
    detail = "simulation: clearing the pending AUTO nonce marker raised (" & _
             failure & "); no sampling was started."
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

Private Function PendingCell() As Range
    Set PendingCell = modWorkbook.Sh(SIM_DATA_SHEET).Range(SIM_PENDING_AUTO_NONCE_CELL)
End Function

Private Function ReadPending(ByRef value As Long, ByRef present As Boolean, _
                             ByRef detail As String) As Boolean
    ' BLANK IS A VALID, EXPECTED STATE - it means no transaction needs recovery -
    ' so it is reported through `present`, not as a read failure.
    Dim raw As Variant, number As Double, blank As Boolean, failure As String

    value = 0
    present = False

    ' READING THE SIDECAR IS ITSELF A COM CALL. Every caller - selection,
    ' establishment verification and clear read-back - needs a decided answer,
    ' so a raise is converted here into a named refusal rather than being left
    ' to escape into whichever handler happens to be armed further out.
    On Error GoTo ReadRaised
    raw = PendingCell.Value2
    blank = modWorkbook.IsEmptyCell(PendingCell)
    On Error GoTo 0

    If blank Then
        ReadPending = True
        Exit Function
    End If
    If Not modWorkbook.IsWholeInRange(raw, CDbl(SIM_NONCE_FIRST_VALID), _
                                      CDbl(SIM_NONCE_LAST_VALID)) Then
        detail = "the pending marker is not a whole nonce in its accepted range"
        Exit Function
    End If
    If Not modWorkbook.TryReadDouble(raw, number) Then
        detail = "the pending marker is not a usable number"
        Exit Function
    End If
    value = modWorkbook.SafeLong(number)
    present = True
    ReadPending = True
    Exit Function

ReadRaised:
    failure = Err.Description
    On Error GoTo 0
    detail = "reading the pending marker raised (" & failure & ")"
End Function

Private Function ReadShared(ByVal row As Long, ByVal minValue As Double, _
                            ByVal maxValue As Double, ByRef value As Long, _
                            ByRef detail As String) As Boolean
    ' THE COUNTER READ IS A COM CALL, and it is decided HERE. Selection reads
    ' the counter with no handler of its own armed; letting a raise escape to
    ' the entry handler would leave the classification at NOT_APPLICABLE and the
    ' attempt row would record a plain REFUSED - saying the run declined to
    ' spend a nonce - while a pending marker was actually standing.
    Dim raw As Variant, number As Double, failure As String

    On Error GoTo SharedReadRaised
    raw = SharedCell(row).Value2
    On Error GoTo 0

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
    Exit Function

SharedReadRaised:
    failure = Err.Description
    On Error GoTo 0
    detail = "reading the stored value raised (" & failure & ")"
End Function

Private Function SharedCell(ByVal row As Long) As Range
    Set SharedCell = modWorkbook.Sh(SIM_DATA_SHEET).Range( _
        SIM_SHARED_VALUE_COLUMN & CStr(row))
End Function

