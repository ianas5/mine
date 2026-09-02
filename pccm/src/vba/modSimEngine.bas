Attribute VB_Name = "modSimEngine"
Option Explicit

' ==========================================================================
' PCCM Phase 6 - the Monte Carlo iteration loop. WORKSHEET-FREE BY CONSTRUCTION.
'
' One pure kernel: take already-resolved in-memory DriverFactors, prepare the
' per-driver simulation representation ONCE, walk the canonical iteration loop,
' and retain exactly two iteration-ordered arrays - one nominal total and one PV
' total per iteration.
'
' builder/pccm_builder/sim_oracle.py is the single definition of these
' semantics; this module is their VBA implementation.
'
' WHAT IS DELIBERATELY ABSENT. No statistic, no mean, no standard deviation, no
' percentile, no contingency, no result digest, no request fingerprint, no
' run_id, no AUTO nonce allocation, no _SimData, no Results, no run state, no
' attempt state, no user command, no workbook publication, no sensitivity and no
' annual stochastic output. The two retained arrays are the INPUT to the later
' statistics and reporting steps; nothing here interprets them.
'
' NOR IS THERE ANY MODEL RESOLUTION. The engine does not touch a worksheet,
' does not read a register, does not build Knom or Kpv, does not derive a
' timeline and does not recompute FX, inflation, profiles or discounting. Knom
' and Kpv arrive ALREADY COLLAPSED by the accepted Phase-5 factor boundary. The
' worksheet handoff belongs to the later reporting module.
'
' NOR ANY RANDOMNESS OF ITS OWN. Seeding and stream identity come from the
' accepted modSimRng public surface; every draw during the loop goes through
' modSimSample. There is no recurrence here, no jump, no Rnd, no Randomize, no
' D6-11 algorithm token, and the engine does not implement Bernoulli, an inverse
' CDF or any Cheng arithmetic - it selects the family and calls the sampler.
'
' NO GLOBAL MUTABLE STATE. No module-level variable, no Static local, no hidden
' generator. Two calls with the same DriverFactors, the same effective seed and
' the same iteration count replay the identical run.
'
' THE PREPARED REPRESENTATION IS PRIVATE. Step 7 proved that a Public VBA UDT is
' caller-writable and had to grow a structural validator for its prepared Beta
' shape. The engine declines to create a second such boundary: its prepared
' driver Type is Private, it is built inside the one public call, and it never
' leaves it.
'
' PREPARE ONCE, THEN RUN. Identifier collection, component-stream construction,
' identity verification, Beta-PERT preparation and every allocation happen
' before the first iteration. Inside the loop there is no ReDim, no stream
' construction, no jump, no seed expansion, no sorting and no shape preparation.
'
' NO PARTIAL SUCCESS. The retained totals are staged locally and committed to
' the caller only after the last iteration succeeds. A refusal at iteration 1,
' at iteration 500 or at the last one leaves the caller's arrays exactly as it
' found them.
'
' RETAINED-ARRAY INDEXING IS ZERO-BASED, element `i - 1` holding iteration `i`.
' The Step-8 authorisation permits one-based indexing only where it agrees with
' the existing Phase-6 authority, and it does not: every array modSimRng and
' modSimSample declare is `0 To n - 1`, so a one-based retained array would be
' the only exception in the phase and a caller mixing LBound conventions across
' the three modules is a hazard worth refusing.
' ==========================================================================

' The two measures, named once. They are the `measure` argument of the shared
' contribution routine and appear in its refusal text; naming them here keeps
' the two spellings from drifting apart in three call sites.
Public Const SIM_MEASURE_NOMINAL As String = "nominal"
Public Const SIM_MEASURE_PV As String = "PV"

' One driver, prepared for the run. PRIVATE, and it never escapes the call.
'
' It holds no worksheet row, no sheet name, no cell address, no ListObject, no
' Range, no Object, no FX table, no inflation vector and no profile vector.
' Quantity is meaningful only for a Cost Line and Probability only for a Risk;
' each is read only by the kind that owns it.
Private Type SimEngineDriver
    PermanentId As String
    IsRisk As Boolean
    DistKind As Long
    MinValue As Double
    MostLikely As Double
    MaxValue As Double
    Quantity As Double
    Probability As Double
    Knom As Double
    Kpv As Double
    BetaShape As SimSampleBetaShape
    HasBetaShape As Boolean
    ValueInitialState As SimRngState
    ValueStreamIndex As Long
    OccurrenceInitialState As SimRngState
    OccurrenceStreamIndex As Long
    HasOccurrenceStream As Boolean
End Type

' ==========================================================================
' THE ONE PUBLIC ENTRY POINT
'
' `drivers()` is the accepted Phase-5 DriverFactors carrier and `driverCount` is
' its LOGICAL length. The array may be unallocated when the count is zero, so no
' bound of it is read on that path.
'
' `totalNominal()` and `totalPv()` are the only outputs besides the refusal
' detail. They are the semantic retained result of Step 8 and nothing here
' sorts, summarises or digests them.
' ==========================================================================
Public Function SimEngineRun(ByRef drivers() As DriverFactors, ByVal driverCount As Long, _
                             ByVal effectiveSeed As Long, ByVal iterations As Long, _
                             ByRef totalNominal() As Double, ByRef totalPv() As Double, _
                             ByRef detail As String) As Boolean
    Dim prepared() As SimEngineDriver
    Dim valueState() As SimRngState, occurrenceState() As SimRngState
    Dim nominalTerm() As Double, pvTerm() As Double
    Dim stagedNominal() As Double, stagedPv() As Double
    Dim costCount As Long, riskCount As Long
    Dim iteration As Long, index As Long
    Dim unitCost As Double, severity As Double, term As Double
    Dim occurred As Boolean, drawn As Double, consumed As Long
    Dim measured As Double

    detail = vbNullString

    ' ---- ITERATION PREFLIGHT, FIRST. Before any allocation, before the seed is
    ' expanded, before a stream exists and before a single draw. `iterations` is
    ' a Long, so there is no non-whole case to reject at this boundary. The
    ' minimum is the business one and the maximum is the technical
    ' representability ceiling; there is no third limit, and 100,000 is a
    ' performance target rather than a cap. ----
    If iterations < SIM_MIN_ITERATIONS Then
        detail = "engine: the iteration count is below the accepted minimum"
        Exit Function
    End If
    If iterations > SIM_MAX_ITERATIONS Then
        detail = "engine: the iteration count is above the technical maximum"
        Exit Function
    End If
    If driverCount < 0 Then
        detail = "engine: a negative driver count"
        Exit Function
    End If

    If Not SimEnginePrepare(drivers, driverCount, effectiveSeed, prepared, _
                            costCount, riskCount, detail) Then Exit Function

    ' ---- ALLOCATION, ONCE, OUTSIDE THE LOOP ----
    ' The working states are COPIES. The prepared initial states are never
    ' mutated, so the same prepared model replays identically.
    If driverCount > 0 Then
        ReDim valueState(0 To driverCount - 1)
        ReDim occurrenceState(0 To driverCount - 1)
        ReDim nominalTerm(0 To driverCount - 1)
        ReDim pvTerm(0 To driverCount - 1)
        For index = 0 To driverCount - 1
            valueState(index) = prepared(index).ValueInitialState
            If prepared(index).HasOccurrenceStream Then
                occurrenceState(index) = prepared(index).OccurrenceInitialState
            End If
        Next index
    Else
        ' The accepted zero-count carrier convention: one physical slot, and the
        ' LOGICAL count - zero - is what says no element is present. No term is
        ' ever written into it and SafeSignedSum is called with a count of zero.
        ReDim valueState(0 To 0)
        ReDim occurrenceState(0 To 0)
        ReDim nominalTerm(0 To 0)
        ReDim pvTerm(0 To 0)
    End If
    ReDim stagedNominal(0 To iterations - 1)
    ReDim stagedPv(0 To iterations - 1)

    ' ==================================================================
    ' THE CANONICAL ITERATION LOOP
    '
    ' Cost Lines first, then Risks, each already in ordinal Permanent-ID order
    ' because `prepared` was built from the canonical component sequence. The
    ' physical order of the caller's DriverFactors array reaches nothing.
    ' ==================================================================
    For iteration = 1 To iterations
        For index = 0 To costCount - 1
            ' THE SAMPLE IS UNIT COST. Quantity is deterministic, sits outside
            ' the distribution and is applied exactly once, here.
            If Not SimEngineSampleValue(prepared(index), valueState(index), unitCost, _
                                        iteration, detail) Then Exit Function
            If Not SimEngineContribution(prepared(index), unitCost, True, _
                                         prepared(index).Knom, SIM_MEASURE_NOMINAL, _
                                         iteration, term, detail) Then Exit Function
            nominalTerm(index) = term
            ' PV IS AN INDEPENDENT ACCUMULATOR. It is built from the same unit
            ' cost and the same Quantity against Kpv, never by discounting the
            ' nominal term that was just computed.
            If Not SimEngineContribution(prepared(index), unitCost, True, _
                                         prepared(index).Kpv, SIM_MEASURE_PV, _
                                         iteration, term, detail) Then Exit Function
            pvTerm(index) = term
        Next index

        For index = costCount To driverCount - 1
            ' D6-18b. THE OCCURRENCE DRAW COMES FIRST, exactly once per Risk per
            ' iteration, at every Probability including 0 and 1.
            If Not SimSampleBernoulli(occurrenceState(index), prepared(index).Probability, _
                                      occurred, drawn, consumed, detail) Then
                detail = "engine: iteration " & CStr(iteration) & ", risk " & _
                         prepared(index).PermanentId & ": " & detail
                Exit Function
            End If
            ' ...AND THE SEVERITY SAMPLER IS THEN INVOKED UNCONDITIONALLY.
            '
            ' Consumption is a property of the DISTRIBUTION, not of the
            ' occurrence. Sampling only when the Risk occurred would make every
            ' later draw on that stream depend on the occurrence decisions
            ' before it, so two runs differing only in one Probability would
            ' produce unrelated severity sequences and could not be compared.
            ' Under D6-18b the severity sequence is a function of the seed and
            ' the distribution alone. A degenerate severity is still invoked and
            ' still consumes nothing.
            If Not SimEngineSampleValue(prepared(index), valueState(index), severity, _
                                        iteration, detail) Then Exit Function
            ' A RISK HAS NO QUANTITY, and Probability was spent on the Bernoulli
            ' draw and appears in no factor. When the Risk did not occur the
            ' severity that was sampled is discarded - it was still drawn - and
            ' the contribution is zero on both measures.
            If Not SimEngineContribution(prepared(index), severity, occurred, _
                                         prepared(index).Knom, SIM_MEASURE_NOMINAL, _
                                         iteration, term, detail) Then Exit Function
            nominalTerm(index) = term
            If Not SimEngineContribution(prepared(index), severity, occurred, _
                                         prepared(index).Kpv, SIM_MEASURE_PV, _
                                         iteration, term, detail) Then Exit Function
            pvTerm(index) = term
        Next index

        ' THE ACCEPTED SIGNED SUM, not a running total. A canonical sequence can
        ' overflow at an intermediate point while the final signed sum remains
        ' representable, and SafeSignedSum reaches that answer without
        ' re-associating a sum that already worked.
        If Not SafeSignedSum(nominalTerm, driverCount, measured) Then
            detail = "engine: iteration " & CStr(iteration) & ": the nominal total is not representable"
            Exit Function
        End If
        stagedNominal(iteration - 1) = measured
        If Not SafeSignedSum(pvTerm, driverCount, measured) Then
            detail = "engine: iteration " & CStr(iteration) & ": the PV total is not representable"
            Exit Function
        End If
        stagedPv(iteration - 1) = measured
    Next iteration

    ' COMMITTED LAST, and only here.
    totalNominal = stagedNominal
    totalPv = stagedPv
    SimEngineRun = True
End Function

' ==========================================================================
' THE ONE PER-DRIVER CONTRIBUTION, and there is exactly one of it
'
' Both the canonical iteration loop above and the replay below reach a driver's
' contribution through here. That is the whole point of the routine existing:
' sensitivity replays a driver in order to explain the total it already
' published, so a second expression of this arithmetic - however carefully
' copied - would let the explanation and the number drift apart while both
' looked right.
'
' THE SHAPE IS THE ACCEPTED ONE, unchanged. A Cost Line multiplies its sampled
' unit cost by the deterministic Quantity and the deployment factor. A Risk has
' no Quantity, and Probability was already spent on the Bernoulli draw and
' appears in no factor here. A Risk that did not occur contributes exactly zero
' on both measures - the severity it drew is discarded, and it was still drawn.
'
' `factor` is Knom or Kpv. PV IS NOT DERIVED FROM NOMINAL: the caller passes the
' other factor and gets an independent product, never a discounted total.
' ==========================================================================
Private Function SimEngineContribution(ByRef prepared As SimEngineDriver, _
                                       ByVal sample As Double, ByVal occurred As Boolean, _
                                       ByVal factor As Double, ByVal measure As String, _
                                       ByVal iteration As Long, ByRef term As Double, _
                                       ByRef detail As String) As Boolean
    Dim factors(0 To 2) As Double
    Dim count As Long
    Dim kind As String

    term = 0#
    If prepared.IsRisk Then
        If Not occurred Then
            SimEngineContribution = True
            Exit Function
        End If
        factors(0) = sample
        factors(1) = factor
        count = 2
        kind = "risk"
    Else
        factors(0) = sample
        factors(1) = prepared.Quantity
        factors(2) = factor
        count = 3
        kind = "cost line"
    End If
    If Not SafeProduct(factors, count, term) Then
        detail = "engine: iteration " & CStr(iteration) & ", " & kind & " " & _
                 prepared.PermanentId & ": the " & measure & _
                 " contribution is not representable"
        Exit Function
    End If
    SimEngineContribution = True
End Function

' ==========================================================================
' PER-DRIVER REPLAY - Phase 7
'
' Reconstructs ONE driver's nominal contribution sequence for the SAME accepted
' run: same resolved model, same effective seed, same iteration count, same
' canonical component-stream assignment. It is OBSERVATIONAL. It allocates no
' run identity, consumes no AUTO nonce, writes nothing and changes no published
' number; the successful run is the same successful run before and after.
'
' WHY IT PREPARES EVERY DRIVER AND REPLAYS ONE. A component's stream index comes
' from its position in the canonical sequence of ALL components, so the target's
' own initial state cannot be derived without building that sequence. Preparing
' derives initial states by jump-ahead; it draws nothing. Only the target's
' streams are then advanced, and no unrelated stream is consumed.
'
' SEQUENTIAL, NEVER A SEEK. A rejection sampler consumes a variable number of
' uniforms per draw, so iteration j is reached by advancing to it. The withdrawn
' direct-seek claim stays withdrawn.
'
' D6-18b IS REPRODUCED EXACTLY. For a Risk the occurrence draw comes first, once
' per iteration at every Probability including 0 and 1, and the severity sampler
' is then invoked UNCONDITIONALLY. Skipping the severity draw on a
' non-occurrence would be faster and would produce a different severity sequence
' from the run this replay claims to explain.
'
' ITERATION IDENTITY IS THE OUTPUT'S WHOLE VALUE. `contributions(j - 1)` is that
' driver's contribution in accepted iteration j. Nothing is sorted, compacted,
' renumbered or dropped - a zero from a Risk that did not occur is an
' observation, not an absence - because P7-4 pairs this vector positionally with
' the persisted TotalNom.
' ==========================================================================
Public Function SimEngineReplayDriver(ByRef drivers() As DriverFactors, _
                                      ByVal driverCount As Long, _
                                      ByVal effectiveSeed As Long, ByVal iterations As Long, _
                                      ByVal permanentId As String, _
                                      ByRef contributions() As Double, _
                                      ByRef detail As String) As Boolean
    Dim prepared() As SimEngineDriver
    Dim valueState As SimRngState, occurrenceState As SimRngState
    Dim costCount As Long, riskCount As Long
    Dim index As Long, target As Long, iteration As Long
    Dim sample As Double, term As Double, drawn As Double
    Dim consumed As Long
    Dim occurred As Boolean

    detail = vbNullString
    If iterations < 1 Then
        detail = "engine: replay needs at least one iteration"
        Exit Function
    End If
    If Not SimEnginePrepare(drivers, driverCount, effectiveSeed, prepared, _
                            costCount, riskCount, detail) Then Exit Function

    ' THE DRIVER IS FOUND BY PERMANENT ID, never by a row or a supply position.
    ' A FULL SCAN, and the first match wins. Permanent ids are unique, so this
    ' is the match; scanning past it costs nothing and keeps the loop shape the
    ' rest of this module uses.
    target = -1
    For index = 0 To costCount + riskCount - 1
        If target < 0 And prepared(index).PermanentId = permanentId Then
            target = index
        End If
    Next index
    If target < 0 Then
        detail = "engine: replay was asked for driver " & permanentId & _
                 ", which is not in this model"
        Exit Function
    End If

    ReDim contributions(0 To iterations - 1)
    valueState = prepared(target).ValueInitialState
    If prepared(target).HasOccurrenceStream Then
        occurrenceState = prepared(target).OccurrenceInitialState
    End If

    For iteration = 1 To iterations
        occurred = True
        If prepared(target).HasOccurrenceStream Then
            If Not SimSampleBernoulli(occurrenceState, prepared(target).Probability, _
                                      occurred, drawn, consumed, detail) Then
                detail = "engine: iteration " & CStr(iteration) & ", risk " & _
                         prepared(target).PermanentId & ": " & detail
                Exit Function
            End If
        End If
        ' UNCONDITIONAL, on both kinds. See D6-18b above.
        If Not SimEngineSampleValue(prepared(target), valueState, sample, _
                                    iteration, detail) Then Exit Function
        If Not SimEngineContribution(prepared(target), sample, occurred, _
                                     prepared(target).Knom, SIM_MEASURE_NOMINAL, _
                                     iteration, term, detail) Then Exit Function
        contributions(iteration - 1) = term
    Next iteration
    SimEngineReplayDriver = True
End Function

' ==========================================================================
' PREPARATION - everything that must happen exactly once
'
' Validates the engine-relevant factor domain, collects the two identifier
' lists, expands the effective seed through the accepted generator, builds the
' canonical component streams, VERIFIES every component identity, and prepares
' each Beta-PERT shape through the Step-7 constructor.
'
' `prepared` comes back in canonical execution order: every Cost Line in ordinal
' Permanent-ID order, then every Risk in ordinal Permanent-ID order. That order
' is not derived here - it is READ OFF the component sequence modSimRng
' produced, so there is exactly one collation implementation in the phase.
' ==========================================================================
Private Function SimEnginePrepare(ByRef drivers() As DriverFactors, ByVal driverCount As Long, _
                                  ByVal effectiveSeed As Long, _
                                  ByRef prepared() As SimEngineDriver, _
                                  ByRef costCount As Long, ByRef riskCount As Long, _
                                  ByRef detail As String) As Boolean
    Dim costIds() As String, riskIds() As String
    Dim claimed() As Boolean
    Dim components() As SimRngComponent
    Dim baseState As SimRngState
    Dim index As Long, source As Long, slot As Long, occurrence As Long, severity As Long

    costCount = 0
    riskCount = 0

    ' ---- the engine-relevant factor domain, and the two identifier lists ----
    If driverCount > 0 Then
        ReDim costIds(0 To driverCount - 1)
        ReDim riskIds(0 To driverCount - 1)
        ReDim claimed(0 To driverCount - 1)
        For index = 0 To driverCount - 1
            If Not SimEngineValidateFactor(drivers(LBound(drivers) + index), detail) Then Exit Function
            If drivers(LBound(drivers) + index).IsRisk Then
                riskIds(riskCount) = drivers(LBound(drivers) + index).PermanentId
                riskCount = riskCount + 1
            Else
                costIds(costCount) = drivers(LBound(drivers) + index).PermanentId
                costCount = costCount + 1
            End If
        Next index
    Else
        ' NO BOUND OF `drivers` IS READ HERE. A zero-driver model is legal and
        ' the caller's array may be unallocated, so nothing inspects it.
        ReDim costIds(0 To 0)
        ReDim riskIds(0 To 0)
        ReDim claimed(0 To 0)
    End If

    ' ---- the effective seed, expanded through the accepted generator ----
    ' Step 8 receives an already-selected seed. It does not decide FIXED versus
    ' AUTO, does not allocate a nonce and persists no counter.
    If Not SimRngStateFromFixedSeed(effectiveSeed, baseState, detail) Then Exit Function

    ' ---- the canonical component streams, built once ----
    If Not SimRngBuildComponentStreams(costIds, costCount, riskIds, riskCount, _
                                       baseState, components, detail) Then Exit Function

    If driverCount = 0 Then
        ' The logical component count is zero, so the one-slot carrier modSimRng
        ' returned is NOT inspected. There is no driver to prepare.
        ReDim prepared(0 To 0)
        SimEnginePrepare = True
        Exit Function
    End If

    ReDim prepared(0 To driverCount - 1)
    For index = 0 To costCount - 1
        If Not SimEngineClaim(drivers, driverCount, components(index), False, _
                              SIM_COMPONENT_1_DRIVER_KIND, SIM_COMPONENT_1_ROLE, index, _
                              claimed, source, detail) Then Exit Function
        If Not SimEngineAdopt(drivers(LBound(drivers) + source), prepared(index), detail) Then Exit Function
        prepared(index).ValueStreamIndex = components(index).StreamIndex
        prepared(index).ValueInitialState = components(index).InitialState
    Next index

    For index = 0 To riskCount - 1
        occurrence = costCount + 2 * index
        severity = occurrence + 1
        slot = costCount + index
        If Not SimEngineClaim(drivers, driverCount, components(occurrence), True, _
                              SIM_COMPONENT_2_DRIVER_KIND, SIM_COMPONENT_2_ROLE, occurrence, _
                              claimed, source, detail) Then Exit Function
        ' THE PAIR MUST BE THE SAME RISK, ADJACENT, AND IN THIS ORDER. An
        ' occurrence borrowed from one Risk and a severity from another would
        ' still run, and would sample a distribution nobody declared.
        If components(severity).PermanentId <> components(occurrence).PermanentId Then
            detail = "engine: the occurrence and severity streams belong to different risks"
            Exit Function
        End If
        If components(severity).DriverKind <> SIM_COMPONENT_3_DRIVER_KIND Then
            detail = "engine: a risk severity component has the wrong driver kind"
            Exit Function
        End If
        If components(severity).Role <> SIM_COMPONENT_3_ROLE Then
            detail = "engine: a risk severity component has the wrong role"
            Exit Function
        End If
        If components(severity).StreamIndex <> SIM_STREAM_INDEX_ORIGIN + severity Then
            detail = "engine: a risk severity component has the wrong stream index"
            Exit Function
        End If
        If components(severity).StreamIndex = components(occurrence).StreamIndex Then
            detail = "engine: a risk shares one stream between occurrence and severity"
            Exit Function
        End If
        If Not SimEngineAdopt(drivers(LBound(drivers) + source), prepared(slot), detail) Then Exit Function
        prepared(slot).ValueStreamIndex = components(severity).StreamIndex
        prepared(slot).ValueInitialState = components(severity).InitialState
        prepared(slot).OccurrenceStreamIndex = components(occurrence).StreamIndex
        prepared(slot).OccurrenceInitialState = components(occurrence).InitialState
        prepared(slot).HasOccurrenceStream = True
    Next index

    ' ---- Beta-PERT shapes, prepared ONCE, by the Step-7 constructor ----
    ' The engine never assembles a SimSampleBetaShape itself and never writes one
    ' of its fields. This is the only place a shape is produced, and after it
    ' nothing mutates one.
    For index = 0 To driverCount - 1
        If prepared(index).DistKind = DIST_BETA_PERT Then
            If Not SimSamplePrepareBetaPert(prepared(index).MinValue, prepared(index).MostLikely, _
                                            prepared(index).MaxValue, prepared(index).BetaShape, _
                                            detail) Then Exit Function
            prepared(index).HasBetaShape = True
        End If
    Next index

    SimEnginePrepare = True
End Function

' One component, verified against what it claims to be, and matched back to
' exactly one unclaimed DriverFactors record.
Private Function SimEngineClaim(ByRef drivers() As DriverFactors, ByVal driverCount As Long, _
                                ByRef component As SimRngComponent, ByVal wantRisk As Boolean, _
                                ByVal wantKind As String, ByVal wantRole As String, _
                                ByVal wantIndex As Long, ByRef claimed() As Boolean, _
                                ByRef source As Long, ByRef detail As String) As Boolean
    Dim index As Long, found As Long
    If component.DriverKind <> wantKind Then
        detail = "engine: a component has the wrong driver kind"
        Exit Function
    End If
    If component.Role <> wantRole Then
        detail = "engine: a component has the wrong role"
        Exit Function
    End If
    If component.StreamIndex <> SIM_STREAM_INDEX_ORIGIN + wantIndex Then
        detail = "engine: a component has the wrong stream index"
        Exit Function
    End If
    If Len(component.PermanentId) = 0 Then
        detail = "engine: a component carries a blank permanent id"
        Exit Function
    End If
    found = -1
    For index = 0 To driverCount - 1
        If drivers(LBound(drivers) + index).IsRisk = wantRisk Then
            If StrComp(drivers(LBound(drivers) + index).PermanentId, component.PermanentId, _
                       vbBinaryCompare) = 0 Then
                If found >= 0 Then
                    detail = "engine: two driver records claim one component identity"
                    Exit Function
                End If
                found = index
            End If
        End If
    Next index
    If found < 0 Then
        detail = "engine: a component identity matches no driver record"
        Exit Function
    End If
    If claimed(found) Then
        detail = "engine: one driver record was claimed by two components"
        Exit Function
    End If
    claimed(found) = True
    source = found
    SimEngineClaim = True
End Function

' Copy the engine-relevant factors across. ONLY what the driver's own kind owns:
' a Cost Line's Probability and a Risk's Quantity are never read, so neither can
' change preparation, refusal, consumption or output.
Private Function SimEngineAdopt(ByRef factor As DriverFactors, ByRef target As SimEngineDriver, _
                                ByRef detail As String) As Boolean
    target.PermanentId = factor.PermanentId
    target.IsRisk = factor.IsRisk
    target.DistKind = factor.DistKind
    target.MinValue = factor.MinValue
    target.MaxValue = factor.MaxValue
    ' MOST LIKELY IS NOT READ FOR A UNIFORM. Accepted Phase-5 D1 ignores it
    ' numerically for that family, so whatever it happens to hold - including a
    ' value no other family would accept - reaches nothing.
    If factor.DistKind <> DIST_UNIFORM Then
        target.MostLikely = factor.MostLikely
    End If
    target.Knom = factor.Knom
    target.Kpv = factor.Kpv
    If factor.IsRisk Then
        target.Probability = factor.Probability
    Else
        target.Quantity = factor.Quantity
    End If
    SimEngineAdopt = True
End Function

' The engine-relevant factor domain, and no more. No repair, no clamp, no
' endpoint swap, no coercion to another family, and no inspection of the
' Central, MeanValue or CentralBasis fields the analytical layer owns.
Private Function SimEngineValidateFactor(ByRef factor As DriverFactors, _
                                         ByRef detail As String) As Boolean
    If Len(factor.PermanentId) = 0 Then
        detail = "engine: a driver has a blank permanent id"
        Exit Function
    End If
    If factor.DistKind <> DIST_UNIFORM And factor.DistKind <> DIST_TRIANGULAR _
       And factor.DistKind <> DIST_BETA_PERT Then
        detail = "engine: driver " & factor.PermanentId & " has an unknown distribution"
        Exit Function
    End If
    If Not IsUsableDouble(factor.MinValue) Then
        detail = "engine: driver " & factor.PermanentId & ": Min is not a finite Double"
        Exit Function
    End If
    If Not IsUsableDouble(factor.MaxValue) Then
        detail = "engine: driver " & factor.PermanentId & ": Max is not a finite Double"
        Exit Function
    End If
    If factor.DistKind = DIST_UNIFORM Then
        If factor.MinValue > factor.MaxValue Then
            detail = "engine: driver " & factor.PermanentId & _
                     " requires Min <= Max; the ordering is refused, not repaired"
            Exit Function
        End If
    Else
        If Not IsUsableDouble(factor.MostLikely) Then
            detail = "engine: driver " & factor.PermanentId & ": Most Likely is not a finite Double"
            Exit Function
        End If
        If factor.MinValue > factor.MostLikely Or factor.MostLikely > factor.MaxValue Then
            detail = "engine: driver " & factor.PermanentId & _
                     " requires Min <= Most Likely <= Max; the ordering is refused, not repaired"
            Exit Function
        End If
    End If
    If Not IsUsableDouble(factor.Knom) Then
        detail = "engine: driver " & factor.PermanentId & ": Knom is not a finite Double"
        Exit Function
    End If
    If Not IsUsableDouble(factor.Kpv) Then
        detail = "engine: driver " & factor.PermanentId & ": Kpv is not a finite Double"
        Exit Function
    End If
    If factor.IsRisk Then
        If Not IsUsableDouble(factor.Probability) Then
            detail = "engine: risk " & factor.PermanentId & ": Probability is not a finite Double"
            Exit Function
        End If
        If factor.Probability < 0# Or factor.Probability > 1# Then
            detail = "engine: risk " & factor.PermanentId & _
                     ": Probability is outside [0, 1]; it is refused, not clamped"
            Exit Function
        End If
    Else
        If Not IsUsableDouble(factor.Quantity) Then
            detail = "engine: cost line " & factor.PermanentId & ": Quantity is not a finite Double"
            Exit Function
        End If
    End If
    SimEngineValidateFactor = True
End Function

' ==========================================================================
' ONE DRAW FROM A DRIVER'S OWN DISTRIBUTION
'
' The engine selects the family and calls the accepted sampler. No inverse CDF,
' no Cheng, no shape parameterisation and no probability comparison lives here.
'
' Uniform receives ONLY Min and Max: Most Likely is not passed, and there is no
' parameter on SimSampleUniform that could receive it.
'
' Beta-PERT uses the shape prepared once per driver by SimSamplePrepareBetaPert.
' Nothing about it is recomputed per iteration.
' ==========================================================================
Private Function SimEngineSampleValue(ByRef prepared As SimEngineDriver, _
                                      ByRef state As SimRngState, ByRef sample As Double, _
                                      ByVal iteration As Long, ByRef detail As String) As Boolean
    Dim consumed As Long, attempts As Long, ok As Boolean
    ok = False
    If prepared.DistKind = DIST_UNIFORM Then
        ok = SimSampleUniform(state, prepared.MinValue, prepared.MaxValue, _
                              sample, consumed, detail)
    ElseIf prepared.DistKind = DIST_TRIANGULAR Then
        ok = SimSampleTriangular(state, prepared.MinValue, prepared.MostLikely, _
                                 prepared.MaxValue, sample, consumed, detail)
    ElseIf prepared.DistKind = DIST_BETA_PERT Then
        If Not prepared.HasBetaShape Then
            detail = "engine: a Beta-PERT driver has no prepared shape"
            Exit Function
        End If
        ok = SimSamplePreparedBeta(state, prepared.BetaShape, sample, consumed, _
                                   attempts, detail)
    Else
        detail = "engine: an unknown distribution reached the sampler"
        Exit Function
    End If
    If Not ok Then
        detail = "engine: iteration " & CStr(iteration) & ", driver " & _
                 prepared.PermanentId & ": " & detail
        Exit Function
    End If
    SimEngineSampleValue = True
End Function
