Attribute VB_Name = "modSimRng"
Option Explicit

' ==========================================================================
' PCCM Phase 6 - the generator backbone. WORKSHEET-FREE BY CONSTRUCTION.
'
' State validation, FIXED seeding, the AUTO nonce mapping, one recurrence step,
' the canonical 2^127 stream jump and canonical component-stream assignment.
' Every procedure is a pure function of Doubles, Longs, Strings and typed
' arrays: the same arguments give the same answers with no workbook open.
'
' builder/pccm_builder/sim_rng.py is the single definition of these semantics;
' this module is their VBA implementation. Every operational constant comes from
' modSimContract, which projects spec/sim_contract.yaml. No literal here
' restates a value that already has an owner.
'
' WHAT IS DELIBERATELY ABSENT. There is no Uniform, Triangular, Beta-PERT or
' Bernoulli sampler, no iteration loop, no contribution arithmetic, no
' statistic, no quantile, no contingency, no digest, no _SimData, no Results, no
' run state and no user command. This module produces RAW uniforms and stream
' identities; everything downstream begins in a later step.
'
' NO GLOBAL MUTABLE STATE. There is no module-level generator, no Static local
' and no hidden singleton. Every operational value is an explicit parameter or a
' returned typed value, so two callers cannot interfere and a run cannot depend
' on what happened before it.
'
' FAILURE SEMANTICS follow the accepted pure-kernel pattern: a public operation
' returns False and names the stage in `detail`, and a failing state-changing
' call leaves the caller's state exactly as it found it. Results are computed
' into locals, the candidate is validated, and only then is the output
' committed. Nothing here displays anything.
' ==========================================================================

' --------------------------------------------------------------------------
' THE ALGORITHM NAME. `MRG32k3aStep` below is the one place in the whole
' repository where this token may appear in executable code: D6-11 scopes it to
' this module and to no other. It is a real private procedure name rather than a
' comment, so the scoped rule is genuinely exercised rather than vacuously
' satisfied by a token the source scanners strip.
' --------------------------------------------------------------------------

' The six-word state, OLDEST-FIRST, exactly as the contract orders it:
'   [s10, s11, s12, s20, s21, s22]
'
' Double, not Long. The words reach 4294967086, which overflows a signed Long,
' and every value they take part in is an exact integer well inside 2^53.
Public Type SimRngState
    S10 As Double
    S11 As Double
    S12 As Double
    S20 As Double
    S21 As Double
    S22 As Double
End Type

' One consumer of one stream. No worksheet row number is retained: a component
' is identified by what it IS, not by where it was entered.
Public Type SimRngComponent
    DriverKind As String
    PermanentId As String
    Role As String
    StreamIndex As Long
    InitialState As SimRngState
End Type

' ==========================================================================
' THE NORMALISATION CONSTANT IS CONSTRUCTED, NOT SPELLED.
'
' modSimContract projects SIM_RNG_NORM as 2.328306549295728E-10, which is the
' accepted Double and needs SIXTEEN significant digits to name.
'
' VBA converts a numeric literal at about fifteen. Phase-5 Gate-B Runtime Run 3
' proved this from the other direction, on MAX_DOUBLE, and modCalcFactors
' records it: a literal that needs more precision than the parser keeps becomes
' a DIFFERENT Double. Here the fifteen-digit form is 2.32830654929573E-10, which
' is four ulp away - every uniform drawn through it would be wrong in the last
' bits, and every downstream vector would miss.
'
' The accepted value is exactly 1 / (m1 + 1) in binary64. m1 is 4294967087, ten
' digits, so SIM_RNG_M1 parses exactly; adding one is exact; and IEEE division
' is correctly rounded. The result is bit-for-bit the accepted normalisation.
'
' This is the same discipline modCalcFactors applies to MAX_DOUBLE: build the
' value from constants that DO survive the parser rather than trust a spelling
' the parser cannot keep. SIM_RNG_NORM remains the authority - a Python test
' asserts the projected literal equals 1 / (SIM_RNG_M1 + 1) exactly, so this
' construction is bound to the projection and cannot drift from it.
'
' It is a Function rather than a Const because a Const initialiser cannot
' compute.
' ==========================================================================
Private Function SimRngNorm() As Double
    SimRngNorm = 1# / (SIM_RNG_M1 + 1#)
End Function

' ==========================================================================
' State validation
'
' Checked at the public boundary rather than assumed. An out-of-range or
' all-zero component produces a stream that looks fine and is not the accepted
' generator, and the all-zero case is ABSORBING: the recurrence can never leave
' it. Nothing here normalises or repairs; an inadmissible state is refused.
' ==========================================================================
Public Function SimRngValidateState(ByRef state As SimRngState, _
                                    ByRef detail As String) As Boolean
    detail = vbNullString
    If Not SimRngValidTriple(state.S10, state.S11, state.S12, SIM_RNG_M1, _
                             "first", detail) Then Exit Function
    If Not SimRngValidTriple(state.S20, state.S21, state.S22, SIM_RNG_M2, _
                             "second", detail) Then Exit Function
    SimRngValidateState = True
End Function

Private Function SimRngValidTriple(ByVal w0 As Double, ByVal w1 As Double, _
                                   ByVal w2 As Double, ByVal m As Double, _
                                   ByVal label As String, _
                                   ByRef detail As String) As Boolean
    If Not SimRngValidWord(w0, m, label, 0, detail) Then Exit Function
    If Not SimRngValidWord(w1, m, label, 1, detail) Then Exit Function
    If Not SimRngValidWord(w2, m, label, 2, detail) Then Exit Function
    If w0 = 0# And w1 = 0# And w2 = 0# Then
        detail = "state: the " & label & " component is all zero, which is absorbing"
        Exit Function
    End If
    SimRngValidTriple = True
End Function

Private Function SimRngValidWord(ByVal word As Double, ByVal m As Double, _
                                 ByVal label As String, ByVal position As Long, _
                                 ByRef detail As String) As Boolean
    ' A fractional word is refused as loudly as an out-of-range one: the
    ' recurrence is integer arithmetic carried in Doubles, and a fraction would
    ' propagate silently through every later step.
    If word <> Fix(word) Then
        detail = "state: " & label & " word " & CStr(position) & " is not an integer"
        Exit Function
    End If
    If word < 0# Then
        detail = "state: " & label & " word " & CStr(position) & " is negative"
        Exit Function
    End If
    If word >= m Then
        detail = "state: " & label & " word " & CStr(position) & " is not below its modulus"
        Exit Function
    End If
    SimRngValidWord = True
End Function

' ==========================================================================
' Modular reduction - NOT VBA `Mod`
'
' VBA's `Mod` coerces its operands to an integer type, so it cannot be used on
' values that exceed Long. The accepted reduction is the Fix-based form, and the
' negative-remainder correction is not optional: the recurrence forms a SIGNED
' difference, and the mathematics requires the non-negative residue.
'
' Fix truncates toward zero. Int floors, and Round rounds; either would give a
' different k for a negative p and therefore a different residue.
' ==========================================================================
Private Function SimRngReduce(ByVal p As Double, ByVal m As Double) As Double
    Dim k As Double, r As Double
    k = Fix(p / m)
    r = p - k * m
    If r < 0# Then r = r + m
    SimRngReduce = r
End Function

' ==========================================================================
' Safe modular multiplication - the locked MultModM decomposition
'
' Computes (a * s + c) mod m exactly, for operands whose naive product would
' leave the exactly-representable integer range of a Double. The split constant
' is SIM_JUMP_DECOMPOSITION_H, projected from the contract; it is not spelled
' here, because a second copy of 2^17 is a second authority.
'
' One primitive, used by BOTH the jump matrix product and the AUTO modular
' exponentiation. A second modular multiply for AUTO would be a second chance to
' get it wrong.
' ==========================================================================
Private Function SimRngMultModM(ByVal a As Double, ByVal s As Double, _
                                ByVal c As Double, ByVal m As Double) As Double
    Dim a1 As Double, v As Double, rest As Double
    rest = a
    a1 = Fix(rest / SIM_JUMP_DECOMPOSITION_H)
    If a1 <> 0# Then
        rest = rest - a1 * SIM_JUMP_DECOMPOSITION_H
        v = a1 * s
        v = v - Fix(v / m) * m
        v = v * SIM_JUMP_DECOMPOSITION_H + rest * s + c
        v = v - Fix(v / m) * m
    Else
        v = rest * s + c
        v = v - Fix(v / m) * m
    End If
    If v < 0# Then v = v + m
    SimRngMultModM = v
End Function

' ==========================================================================
' FIXED seeding
'
' D6-05(a): the scalar repeated into all six words. No mixer, no hash, no
' expansion. The admissible domain belongs to input_contract.yaml and reaches
' this module only as SIM_SEED_MIN / SIM_SEED_MAX.
' ==========================================================================
Public Function SimRngStateFromFixedSeed(ByVal seed As Long, _
                                         ByRef state As SimRngState, _
                                         ByRef detail As String) As Boolean
    Dim candidate As SimRngState
    detail = vbNullString
    If seed < SIM_SEED_MIN Or seed > SIM_SEED_MAX Then
        detail = "fixed seed: outside the admissible domain"
        Exit Function
    End If
    candidate.S10 = CDbl(seed)
    candidate.S11 = CDbl(seed)
    candidate.S12 = CDbl(seed)
    candidate.S20 = CDbl(seed)
    candidate.S21 = CDbl(seed)
    candidate.S22 = CDbl(seed)
    If Not SimRngValidateState(candidate, detail) Then Exit Function
    state = candidate
    SimRngStateFromFixedSeed = True
End Function

' ==========================================================================
' AUTO nonce -> effective seed
'
' D6-03(b): effective_seed = multiplier ^ nonce mod modulus, a modular POWER
' evaluated by square-and-multiply in O(log nonce).
'
' Stepping the cycle `nonce` times gives the same answer and is NOT the
' authority: at a nonce near the period it is unusable, and stating the mapping
' as a power is what tells an implementation to square and multiply. Floating
' exponentiation is not an option either - the intermediate values leave exact
' range immediately.
'
' PURE. This persists nothing, advances no stored counter and allocates no
' attempt metadata. The transactional nonce lifecycle belongs to a later module.
' ==========================================================================
Public Function SimRngAutoSeedFromNonce(ByVal nonce As Long, ByRef seed As Long, _
                                        ByRef detail As String) As Boolean
    Dim baseValue As Double, result As Double, remaining As Double, half As Double
    detail = vbNullString
    If nonce < SIM_NONCE_FIRST_VALID Then
        detail = "auto nonce: below the first valid allocation"
        Exit Function
    End If
    If nonce >= SIM_NONCE_EXHAUSTED Then
        ' Never wrap. Reissuing the seed for nonce 0 silently is the one outcome
        ' the lifecycle exists to prevent.
        detail = "auto nonce: exhausted"
        Exit Function
    End If

    result = 1#
    baseValue = SimRngReduce(CDbl(SIM_AUTO_MULTIPLIER), CDbl(SIM_AUTO_MODULUS))
    remaining = CDbl(nonce)
    Do While remaining > 0#
        half = Fix(remaining / 2#)
        If remaining - half * 2# = 1# Then
            result = SimRngMultModM(result, baseValue, 0#, CDbl(SIM_AUTO_MODULUS))
        End If
        baseValue = SimRngMultModM(baseValue, baseValue, 0#, CDbl(SIM_AUTO_MODULUS))
        remaining = half
    Loop

    If result <> Fix(result) Or result < 0# Or result >= CDbl(SIM_AUTO_MODULUS) Then
        detail = "auto nonce: the modular power left its residue class"
        Exit Function
    End If
    seed = CLng(result)
    SimRngAutoSeedFromNonce = True
End Function

' ==========================================================================
' One recurrence step
'
' The scoped algorithm name lives here. The products are already proven to stay
' inside the exactly-representable integer range by the accepted Step-0
' evidence, so the base recurrence uses the plain Fix reduction rather than the
' decomposition; the decomposition exists for the jump and the AUTO power, whose
' operands do not have that guarantee.
' ==========================================================================
Private Function MRG32k3aStep(ByRef state As SimRngState, ByRef advanced As SimRngState, _
                              ByRef p1 As Double, ByRef p2 As Double, _
                              ByRef detail As String) As Boolean
    Dim signed As Double
    If Not SimRngValidateState(state, detail) Then Exit Function

    signed = SIM_RNG_A12 * state.S11 - SIM_RNG_A13N * state.S10
    p1 = SimRngReduce(signed, SIM_RNG_M1)

    signed = SIM_RNG_A21 * state.S22 - SIM_RNG_A23N * state.S20
    p2 = SimRngReduce(signed, SIM_RNG_M2)

    ' Advance oldest-first: each component drops its oldest word and gains the
    ' new one at the newest end.
    advanced.S10 = state.S11
    advanced.S11 = state.S12
    advanced.S12 = p1
    advanced.S20 = state.S21
    advanced.S21 = state.S22
    advanced.S22 = p2

    If Not SimRngValidateState(advanced, detail) Then Exit Function
    MRG32k3aStep = True
End Function

' ==========================================================================
' One uniform
'
' The state is mutated in place for hot-loop efficiency, but ONLY after the
' candidate has been validated and the uniform has been shown to lie strictly
' inside (0, 1). A failure leaves the caller's state exactly as it was.
'
' The comparison is `p1 <= p2`, and the boundary is authoritative: `<` would
' produce a different uniform whenever the two residues are equal.
' ==========================================================================
Public Function SimRngNextUniform(ByRef state As SimRngState, ByRef u As Double, _
                                  ByRef detail As String) As Boolean
    Dim advanced As SimRngState
    Dim p1 As Double, p2 As Double, candidate As Double
    detail = vbNullString
    If Not MRG32k3aStep(state, advanced, p1, p2, detail) Then Exit Function

    If p1 <= p2 Then
        candidate = (p1 - p2 + SIM_RNG_M1) * SimRngNorm()
    Else
        candidate = (p1 - p2) * SimRngNorm()
    End If

    If Not (candidate > 0# And candidate < 1#) Then
        ' Both endpoints are excluded by construction, so this cannot happen for
        ' a valid state: it means a constant or the state is wrong.
        detail = "uniform: the combination left the open interval (0, 1)"
        Exit Function
    End If

    u = candidate
    state = advanced
    SimRngNextUniform = True
End Function

' ==========================================================================
' The canonical 2^127 stream jump
'
' ORIENTATION, stated because getting it wrong is silent. PCCM stores state
' OLDEST-FIRST; the jump matrices operate on NEWEST-FIRST triples. Each triple
' is therefore reversed on the way in and reversed back on the way out. A
' transpose is a different matrix, and a dropped reversal is a different vector:
' both produce a plausible stream that is not the canonical one.
'
' Every term goes through the safe modular multiply. The naive matrix product
' leaves exact Double range by three orders of magnitude, so a plain
' row-dot-vector would be silently wrong rather than loudly wrong.
'
' No matrix literal appears here - the elements are the projected constants. No
' substreams: 2^76 is not used in Phase 6.
' ==========================================================================
Public Function SimRngJumpNextStream(ByRef state As SimRngState, _
                                     ByRef jumped As SimRngState, _
                                     ByRef detail As String) As Boolean
    Dim candidate As SimRngState
    Dim inFirst(0 To 2) As Double, inSecond(0 To 2) As Double
    Dim outFirst(0 To 2) As Double, outSecond(0 To 2) As Double
    detail = vbNullString
    If Not SimRngValidateState(state, detail) Then Exit Function

    ' Reverse in: oldest-first [s10, s11, s12] becomes newest-first.
    inFirst(0) = state.S12
    inFirst(1) = state.S11
    inFirst(2) = state.S10
    inSecond(0) = state.S22
    inSecond(1) = state.S21
    inSecond(2) = state.S20

    outFirst(0) = SimRngJumpRow(SIM_JUMP_A1_R1C1, SIM_JUMP_A1_R1C2, SIM_JUMP_A1_R1C3, _
                                inFirst, SIM_RNG_M1)
    outFirst(1) = SimRngJumpRow(SIM_JUMP_A1_R2C1, SIM_JUMP_A1_R2C2, SIM_JUMP_A1_R2C3, _
                                inFirst, SIM_RNG_M1)
    outFirst(2) = SimRngJumpRow(SIM_JUMP_A1_R3C1, SIM_JUMP_A1_R3C2, SIM_JUMP_A1_R3C3, _
                                inFirst, SIM_RNG_M1)

    outSecond(0) = SimRngJumpRow(SIM_JUMP_A2_R1C1, SIM_JUMP_A2_R1C2, SIM_JUMP_A2_R1C3, _
                                 inSecond, SIM_RNG_M2)
    outSecond(1) = SimRngJumpRow(SIM_JUMP_A2_R2C1, SIM_JUMP_A2_R2C2, SIM_JUMP_A2_R2C3, _
                                 inSecond, SIM_RNG_M2)
    outSecond(2) = SimRngJumpRow(SIM_JUMP_A2_R3C1, SIM_JUMP_A2_R3C2, SIM_JUMP_A2_R3C3, _
                                 inSecond, SIM_RNG_M2)

    ' Reverse out: newest-first result becomes oldest-first state.
    candidate.S10 = outFirst(2)
    candidate.S11 = outFirst(1)
    candidate.S12 = outFirst(0)
    candidate.S20 = outSecond(2)
    candidate.S21 = outSecond(1)
    candidate.S22 = outSecond(0)

    If Not SimRngValidateState(candidate, detail) Then Exit Function
    jumped = candidate
    SimRngJumpNextStream = True
End Function

Private Function SimRngJumpRow(ByVal c0 As Double, ByVal c1 As Double, ByVal c2 As Double, _
                               ByRef vector() As Double, ByVal m As Double) As Double
    Dim acc As Double
    acc = SimRngMultModM(c0, vector(0), 0#, m)
    acc = SimRngMultModM(c1, vector(1), acc, m)
    acc = SimRngMultModM(c2, vector(2), acc, m)
    SimRngJumpRow = acc
End Function

' ==========================================================================
' Stream k
'
' Stream 0 is the base state; stream k is the base advanced by k canonical
' jumps, computed by walking the ladder. There is no lookup table and no
' maximum of 400: 400 is the design-target component count, not a contract cap,
' and the accepted vectors include stream 401 precisely so a table masquerading
' as an algorithm cannot pass.
' ==========================================================================
Public Function SimRngStreamInitialState(ByRef baseState As SimRngState, ByVal k As Long, _
                                         ByRef state As SimRngState, _
                                         ByRef detail As String) As Boolean
    Dim current As SimRngState, stepped As SimRngState
    Dim index As Long
    detail = vbNullString
    If k < 0 Then
        detail = "stream index: negative"
        Exit Function
    End If
    If Not SimRngValidateState(baseState, detail) Then Exit Function
    current = baseState
    For index = 1 To k
        If Not SimRngJumpNextStream(current, stepped, detail) Then Exit Function
        current = stepped
    Next index
    state = current
    SimRngStreamInitialState = True
End Function

' ==========================================================================
' Canonical component stream assignment - D6-16 family A
'
' One Cost Line takes one value stream. One Risk takes two: occurrence and
' severity. The total is C + 2R.
'
' THE ORDER IS NOT THREE GLOBAL BLOCKS. Kind and role are SEPARATE sort keys, so
' the Risks interleave per driver:
'
'   ... R-099 occurrence, R-099 severity, R-100 occurrence, R-100 severity
'
' and not "every occurrence, then every severity". Collapsing the two axes into
' one is the mistake this comment exists to prevent.
'
' The caller's arrays are never reordered. Ordering runs on a private index
' permutation, exactly as the accepted Phase-5 fingerprint ordering does.
' ==========================================================================
Public Function SimRngBuildComponentStreams(ByRef costIds() As String, ByVal costCount As Long, _
                                            ByRef riskIds() As String, ByVal riskCount As Long, _
                                            ByRef baseState As SimRngState, _
                                            ByRef components() As SimRngComponent, _
                                            ByRef detail As String) As Boolean
    Dim costOrder() As Long, riskOrder() As Long
    Dim built() As SimRngComponent
    Dim total As Long, index As Long, slot As Long
    Dim current As SimRngState, stepped As SimRngState

    detail = vbNullString
    If costCount < 0 Or riskCount < 0 Then
        detail = "components: a negative driver count"
        Exit Function
    End If
    total = costCount + 2 * riskCount

    ' Validated before any path can return, the EMPTY one included. An empty
    ' component set is not permission to accept a state the recurrence cannot
    ' legally be in.
    If Not SimRngValidateState(baseState, detail) Then Exit Function

    ' ==================================================================
    ' ZERO DRIVERS IS A LEGAL MODEL.
    '
    ' No accepted contract requires a Cost Line or a Risk to exist, the
    ' accepted Phase-5 source tests pin that an empty driver set is NOT
    ' refused once the model-level prerequisites resolve, and Phase 6
    ' introduced no minimum of its own. sim_rng.py agrees: components_for
    ' ((), ()) is the empty tuple and component_stream_states validates the
    ' base state, produces nothing and jumps nowhere.
    '
    ' An earlier draft of this module refused here. That was an INVENTED
    ' business prerequisite, and it put the VBA at odds with the accepted
    ' Python reference. There is no minimum-driver authority, so there is no
    ' minimum here - not `total >= 1`, not "at least one Cost Line", not
    ' "at least one Risk".
    '
    ' NEITHER DRIVER ARRAY IS TOUCHED on this path. No LBound is read, no
    ' ordering runs, no identity is inspected: with no component there is
    ' nothing to order, and reading a bound off an array the caller may
    ' never have sized would be work invented to service the empty case.
    '
    ' THE CARRIER FOLLOWS THE ACCEPTED PHASE-5 ZERO-COUNT CONVENTION, the
    ' one CalcFpSortedRecords uses and SimRngOrderIds already uses above:
    ' VBA has no zero-length dynamic array, so the output is sized to one
    ' slot and THE LOGICAL COUNT - costCount + 2 * riskCount, which the
    ' caller supplied - decides whether any element may be inspected. At
    ' zero, no element is semantically present.
    '
    ' The slot is left at its Type defaults and NOTHING is written into it.
    ' It is not a component: its PermanentId is the empty string, which is
    ' exactly what SimRngOrderIds refuses, so it cannot be mistaken for one.
    ' It is assigned rather than left alone so a caller cannot retain an
    ' earlier, longer result and read it as this answer.
    ' ==================================================================
    If total = 0 Then
        ReDim built(0 To 0)
        components = built
        SimRngBuildComponentStreams = True
        Exit Function
    End If

    If Not SimRngOrderIds(costIds, costCount, "cost line", costOrder, detail) Then Exit Function
    If Not SimRngOrderIds(riskIds, riskCount, "risk", riskOrder, detail) Then Exit Function

    ReDim built(0 To total - 1)
    slot = 0
    For index = 0 To costCount - 1
        built(slot).DriverKind = SIM_COMPONENT_1_DRIVER_KIND
        built(slot).PermanentId = costIds(LBound(costIds) + costOrder(index))
        built(slot).Role = SIM_COMPONENT_1_ROLE
        slot = slot + 1
    Next index
    For index = 0 To riskCount - 1
        built(slot).DriverKind = SIM_COMPONENT_2_DRIVER_KIND
        built(slot).PermanentId = riskIds(LBound(riskIds) + riskOrder(index))
        built(slot).Role = SIM_COMPONENT_2_ROLE
        slot = slot + 1
        built(slot).DriverKind = SIM_COMPONENT_3_DRIVER_KIND
        built(slot).PermanentId = riskIds(LBound(riskIds) + riskOrder(index))
        built(slot).Role = SIM_COMPONENT_3_ROLE
        slot = slot + 1
    Next index

    ' Walk the jump ladder ONCE. Recomputing each stream k from the base would
    ' be the same states in O(N^2) jumps instead of O(N).
    current = baseState
    For index = 0 To total - 1
        built(index).StreamIndex = SIM_STREAM_INDEX_ORIGIN + index
        built(index).InitialState = current
        If index < total - 1 Then
            If Not SimRngJumpNextStream(current, stepped, detail) Then Exit Function
            current = stepped
        End If
    Next index

    components = built
    SimRngBuildComponentStreams = True
End Function

' Insertion sort on a private index permutation, ordinal on UTF-16 code units.
'
' vbBinaryCompare is the accepted Phase-5 ordering discipline and is not
' negotiable here: a locale collation, a case fold, a trim or a numeric-suffix
' reading would all put CL-999 before CL-1000, and the accepted order is
' lexical, so CL-1000 comes first.
Private Function SimRngOrderIds(ByRef ids() As String, ByVal count As Long, _
                                ByVal label As String, ByRef order() As Long, _
                                ByRef detail As String) As Boolean
    Dim index As Long, probe As Long, moving As Long
    If count = 0 Then
        ReDim order(0 To 0)
        SimRngOrderIds = True
        Exit Function
    End If
    ReDim order(0 To count - 1)
    For index = 0 To count - 1
        If Len(ids(LBound(ids) + index)) = 0 Then
            detail = "components: a " & label & " has a blank permanent id"
            Exit Function
        End If
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
    ' A duplicate identity is REFUSED, never silently deduplicated: two drivers
    ' claiming one identity would quietly share a stream.
    For index = 1 To count - 1
        If StrComp(ids(LBound(ids) + order(index - 1)), ids(LBound(ids) + order(index)), _
                   vbBinaryCompare) = 0 Then
            detail = "components: duplicate " & label & " permanent id"
            Exit Function
        End If
    Next index
    SimRngOrderIds = True
End Function
