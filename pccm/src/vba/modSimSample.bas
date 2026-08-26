Attribute VB_Name = "modSimSample"
Option Explicit

' ==========================================================================
' PCCM Phase 6 - the stochastic transforms. WORKSHEET-FREE BY CONSTRUCTION.
'
' Uniform, Triangular, prepared Beta-PERT through the locked Cheng BB/BC
' formulation, and the Bernoulli occurrence primitive. That is the whole of this
' module.
'
' builder/pccm_builder/sim_sample.py is the single definition of these
' semantics; this module is their VBA implementation. Every operational constant
' comes from modSimContract, which projects spec/sim_contract.yaml. No literal
' here restates a value that already has an owner.
'
' WHAT IS DELIBERATELY ABSENT. There is no iteration loop, no Cost Line or Risk
' contribution, no Quantity, no Knom, no Kpv, no retained array, no statistic,
' no quantile, no contingency, no request fingerprint, no result digest, no
' _SimData, no Results, no run state, no run counter and no user command. This
' module turns raw uniforms into samples and decisions; the engine that calls it
' begins in a later step.
'
' NOR IS THERE ANY RANDOMNESS OF ITS OWN. Every uniform comes from the accepted
' modSimRng public surface - SimRngValidateState and SimRngNextUniform. There is
' no recurrence here, no jump, no seeding, no Rnd and no Randomize, and the
' D6-11 algorithm token does not appear: this module is NOT granted it and does
' not need it.
'
' NO GLOBAL MUTABLE STATE. No module-level generator, no Static local, no hidden
' singleton. Every operational value is an explicit parameter or a returned
' typed value.
'
' CONSUMPTION IS PART OF THE CONTRACT. Every state-consuming sampler reports how
' many uniforms it consumed and, for the rejection sampler, how many proposal
' attempts it made. Under acceptance/rejection the count is not a fixed property
' of the call - it depends on the values drawn - and a reproduction has to match
' it exactly, not merely match the value.
'
'     Uniform, non-degenerate       1 uniform
'     Triangular, non-degenerate    1 uniform
'     Beta-PERT, non-degenerate     SIM_CHENG_UNIFORMS_PER_ATTEMPT x attempts
'     Bernoulli, ANY valid p        1 uniform
'     ANY degenerate distribution   0 uniforms, state unchanged
'
' FAILURE SEMANTICS follow the accepted pure-kernel pattern used by modSimRng: a
' public operation returns False and names the family and stage in `detail`, and
' a failing call leaves the caller's state, sample and counts exactly as it
' found them. Every state-consuming sampler draws against a LOCAL working copy
' of the caller's state and commits that copy last.
'
' A REJECTED CHENG PROPOSAL IS NEVER REWOUND. It advances the local working
' state, and the retry continues from there. A later failure still leaves the
' CALLER untouched, because the advancing happened on the local copy.
' ==========================================================================

' The per-DRIVER prepared Beta-PERT shape.
'
' A simulation samples one driver many thousands of times, and none of this
' changes between iterations. Preparing it once is not an optimisation detail:
' recomputing a square root and the Cheng constants inside the proposal loop
' would be work the shape already settled, per iteration, per driver.
'
' It holds NO RNG state, NO worksheet object, NO driver row and no mutable
' singleton. Preparing a shape draws ZERO uniforms.
'
' UseChengBB is a Boolean, not a dispatch string: the choice is binary, and an
' unowned magic string would be a second spelling of something the rule already
' decides.
Public Type SimSampleBetaShape
    MinValue As Double
    MostLikely As Double
    MaxValue As Double
    Alpha As Double
    Beta As Double
    Degenerate As Boolean
    UseChengBB As Boolean
    ChengA As Double
    ChengB As Double
    ChengAlpha As Double
    ChengBeta As Double
    ChengGamma As Double
    ChengDelta As Double
    ChengK1 As Double
    ChengK2 As Double
    FirstParameterIsOrientedA As Boolean
    Prepared As Boolean
End Type

' ==========================================================================
' Uniform
'
' MOST LIKELY IS NOT A PARAMETER. Accepted Phase-5 D1 ignores it numerically for
' this family, so it is absent from the signature entirely rather than accepted
' and quietly dropped - an argument that cannot be passed cannot be read.
'
' DEGENERACY IS a = b, and only that. A shared a = m = b predicate would make a
' degenerate Uniform depend on the very input the family ignores: a Uniform with
' Min = Max and an unrelated populated Most Likely would enter the sampler,
' consume a uniform, and shift every later draw on that component.
' ==========================================================================
Public Function SimSampleUniform(ByRef state As SimRngState, _
                                 ByVal minValue As Double, ByVal maxValue As Double, _
                                 ByRef sample As Double, ByRef uniformsConsumed As Long, _
                                 ByRef detail As String) As Boolean
    Dim working As SimRngState
    Dim u As Double, candidate As Double
    detail = vbNullString

    If Not IsUsableDouble(minValue) Then
        detail = "Uniform: Min is not a finite Double"
        Exit Function
    End If
    If Not IsUsableDouble(maxValue) Then
        detail = "Uniform: Max is not a finite Double"
        Exit Function
    End If
    If minValue > maxValue Then
        detail = "Uniform: Min exceeds Max; the ordering is refused, not repaired"
        Exit Function
    End If

    ' VALIDATED BEFORE ANY PATH CAN RETURN, the degenerate one included. Zero
    ' consumption is a property of the DISTRIBUTION; it is not a licence to
    ' accept a state the recurrence cannot legally be in.
    If Not SimRngValidateState(state, detail) Then Exit Function

    If minValue = maxValue Then
        sample = minValue
        uniformsConsumed = 0
        SimSampleUniform = True
        Exit Function
    End If

    working = state
    If Not SimRngNextUniform(working, u, detail) Then Exit Function

    ' THE STABLE CONVEX FORM, and it is not interchangeable with a + u*(b - a):
    ' for a = -MAX_DOUBLE, b = +MAX_DOUBLE the difference overflows while every
    ' convex result is representable, and a legal support may not be lost to a
    ' naive intermediate.
    candidate = (1# - u) * minValue + u * maxValue
    If Not IsUsableDouble(candidate) Then
        detail = "Uniform: the convex rescale is not representable as a finite Double"
        Exit Function
    End If

    sample = candidate
    uniformsConsumed = 1
    state = working
    SimSampleUniform = True
End Function

' ==========================================================================
' Triangular
'
' The inverse CDF, evaluated in a CONDITIONED space and rescaled afterwards.
' The raw (b - a)(m - a) product overflows for supports near Double maximum long
' before the answer does, so the arithmetic is done on a/s, m/s, b/s.
'
' THE BRANCH BOUNDARY IS `u <= c`. At u = c exactly the result is the mode;
' `<` moves it. m = a gives c = 0 and every draw takes the upper branch, m = b
' gives c = 1 and every draw takes the lower one - by the arithmetic, not by an
' endpoint special case.
' ==========================================================================
Public Function SimSampleTriangular(ByRef state As SimRngState, _
                                    ByVal minValue As Double, ByVal mostLikely As Double, _
                                    ByVal maxValue As Double, ByRef sample As Double, _
                                    ByRef uniformsConsumed As Long, _
                                    ByRef detail As String) As Boolean
    Dim working As SimRngState
    Dim u As Double, s As Double, an As Double, mn As Double, bn As Double
    Dim span As Double, c As Double, conditioned As Double, candidate As Double
    detail = vbNullString

    If Not SimSampleOrderedTriple(minValue, mostLikely, maxValue, "Triangular", detail) Then Exit Function
    If Not SimRngValidateState(state, detail) Then Exit Function

    If minValue = mostLikely And mostLikely = maxValue Then
        sample = minValue
        uniformsConsumed = 0
        SimSampleTriangular = True
        Exit Function
    End If

    ' Conditioned BEFORE the draw, so a refusal costs nothing. The support is
    ' ordered and not degenerate here, so the span is positive by construction;
    ' the guard exists because "by construction" is a claim, not a check.
    s = SimSampleScale(minValue, mostLikely, maxValue)
    an = minValue / s
    mn = mostLikely / s
    bn = maxValue / s
    span = bn - an
    If span <= 0# Then
        detail = "Triangular: the conditioned support has no width"
        Exit Function
    End If
    c = (mn - an) / span

    working = state
    If Not SimRngNextUniform(working, u, detail) Then Exit Function

    If u <= c Then
        conditioned = an + Sqr(u * span * (mn - an))
    Else
        conditioned = bn - Sqr((1# - u) * span * (bn - mn))
    End If

    candidate = conditioned * s
    If Not IsUsableDouble(candidate) Then
        detail = "Triangular: the conditioned rescale is not representable as a finite Double"
        Exit Function
    End If

    sample = candidate
    uniformsConsumed = 1
    state = working
    SimSampleTriangular = True
End Function

' ==========================================================================
' Beta-PERT preparation - ONCE PER DRIVER
'
' DEGENERACY IS DETECTED BEFORE r IS FORMED. a = m = b would make (m-a)/(b-a)
' an evaluated 0/0; the order of these two steps is the reason it never is.
'
' The shape family is the accepted one: alpha = 1 + lambda*r,
' beta = 1 + lambda*(1 - r), with lambda projected as SIM_PERT_LAMBDA. Both lie
' in [SIM_PERT_SHAPE_LOWER, SIM_PERT_SHAPE_UPPER] and that IS checked here.
'
' The companion identity alpha + beta = SIM_PERT_ALPHA_PLUS_BETA is NOT checked
' at runtime, and deliberately so: it is exact in real arithmetic but not in
' binary64 for every r - 1 + 4r and 1 + 4(1 - r) can each round, leaving a sum
' one ulp from six. Gating on it would refuse correct shapes, and a tolerance
' inside a sampler is exactly what the numerical authority forbids. The identity
' is evidence, asserted against the accepted corpus by the Step-7 tests.
'
' DISPATCH: min(alpha, beta) > 1 is BB, and EQUALITY BELONGS TO BC. m = a gives
' alpha = 1 and m = b gives beta = 1, so both endpoints reach BC by the rule
' rather than by a special case bolted on afterwards.
'
' PREPARATION DRAWS ZERO UNIFORMS and takes no RNG state.
' ==========================================================================
Public Function SimSamplePrepareBetaPert(ByVal minValue As Double, ByVal mostLikely As Double, _
                                         ByVal maxValue As Double, _
                                         ByRef prepared As SimSampleBetaShape, _
                                         ByRef detail As String) As Boolean
    Dim candidate As SimSampleBetaShape
    Dim s As Double, an As Double, mn As Double, bn As Double, span As Double, r As Double
    Dim alpha0 As Double, beta0 As Double, ca As Double, cb As Double, denominator As Double
    detail = vbNullString

    If Not SimSampleOrderedTriple(minValue, mostLikely, maxValue, "Beta-PERT", detail) Then Exit Function

    candidate.MinValue = minValue
    candidate.MostLikely = mostLikely
    candidate.MaxValue = maxValue

    If minValue = mostLikely And mostLikely = maxValue Then
        candidate.Degenerate = True
        candidate.Prepared = True
        prepared = candidate
        SimSamplePrepareBetaPert = True
        Exit Function
    End If

    s = SimSampleScale(minValue, mostLikely, maxValue)
    an = minValue / s
    mn = mostLikely / s
    bn = maxValue / s
    span = bn - an
    If span <= 0# Then
        detail = "Beta-PERT: the conditioned support has no width"
        Exit Function
    End If

    r = (mn - an) / span
    alpha0 = 1# + SIM_PERT_LAMBDA * r
    beta0 = 1# + SIM_PERT_LAMBDA * (1# - r)
    If Not SimSampleShapeInFamily(alpha0, detail) Then Exit Function
    If Not SimSampleShapeInFamily(beta0, detail) Then Exit Function
    candidate.Alpha = alpha0
    candidate.Beta = beta0

    If SimSampleMinOf(alpha0, beta0) > SIM_PERT_SHAPE_LOWER Then
        ' ---- BB. Orientation is min, max. ----
        candidate.UseChengBB = True
        ca = SimSampleMinOf(alpha0, beta0)
        cb = SimSampleMaxOf(alpha0, beta0)
        candidate.ChengA = ca
        candidate.ChengB = cb
        candidate.ChengAlpha = ca + cb
        denominator = SIM_CHENG_BB_LITERAL_4 * ca * cb - candidate.ChengAlpha
        If denominator <= 0# Then
            detail = "Beta-PERT: the BB beta denominator is not positive"
            Exit Function
        End If
        candidate.ChengBeta = Sqr((candidate.ChengAlpha - SIM_CHENG_BB_LITERAL_4) / denominator)
        If candidate.ChengBeta <= 0# Then
            detail = "Beta-PERT: the BB beta term is not positive"
            Exit Function
        End If
        candidate.ChengGamma = ca + SIM_CHENG_BB_LITERAL_5 / candidate.ChengBeta
    Else
        ' ---- BC. Orientation is max, min - the OPPOSITE of BB. Inverting it is
        ' silent: the sampler still returns a valid Beta variate, just of the
        ' mirrored distribution. ----
        candidate.UseChengBB = False
        ca = SimSampleMaxOf(alpha0, beta0)
        cb = SimSampleMinOf(alpha0, beta0)
        If cb <= 0# Then
            detail = "Beta-PERT: the BC oriented b is not positive"
            Exit Function
        End If
        candidate.ChengA = ca
        candidate.ChengB = cb
        candidate.ChengAlpha = ca + cb
        candidate.ChengBeta = 1# / cb
        candidate.ChengDelta = 1# + ca - cb
        If candidate.ChengDelta <= 0# Then
            detail = "Beta-PERT: the BC delta term is not positive"
            Exit Function
        End If
        denominator = ca * candidate.ChengBeta - SIM_CHENG_BC_LITERAL_3
        If denominator = 0# Then
            detail = "Beta-PERT: the BC k1 denominator is zero"
            Exit Function
        End If
        candidate.ChengK1 = candidate.ChengDelta * _
                            (SIM_CHENG_BC_LITERAL_1 + SIM_CHENG_BC_LITERAL_2 * cb) / denominator
        candidate.ChengK2 = SIM_CHENG_BC_LITERAL_5 + _
                            (SIM_CHENG_BC_LITERAL_6 + SIM_CHENG_BC_LITERAL_5 / candidate.ChengDelta) * cb
    End If

    ' WHICH ORIENTED PARAMETER THE ORIGINAL alpha BECAME. The accepted return
    ' rule reads this, and mirroring it returns the mirrored distribution.
    If alpha0 = candidate.ChengA Then
        candidate.FirstParameterIsOrientedA = True
    Else
        candidate.FirstParameterIsOrientedA = False
    End If
    candidate.Prepared = True
    prepared = candidate
    SimSamplePrepareBetaPert = True
End Function

' ==========================================================================
' Sample a PREPARED shape
'
' Nothing here recomputes a shape constant. The square root, the orientation,
' the dispatch and k1/k2 were settled once by SimSamplePrepareBetaPert; this is
' the per-iteration path and it reads them.
' ==========================================================================
Public Function SimSamplePreparedBeta(ByRef state As SimRngState, _
                                      ByRef prepared As SimSampleBetaShape, _
                                      ByRef sample As Double, ByRef uniformsConsumed As Long, _
                                      ByRef proposalAttempts As Long, _
                                      ByRef detail As String) As Boolean
    Dim working As SimRngState
    Dim y As Double, attempts As Long, candidate As Double
    detail = vbNullString

    If Not prepared.Prepared Then
        detail = "Beta-PERT: the shape was never prepared"
        Exit Function
    End If

    ' Validated before any path can return, the degenerate one included.
    If Not SimRngValidateState(state, detail) Then Exit Function

    If prepared.Degenerate Then
        sample = prepared.MinValue
        uniformsConsumed = 0
        proposalAttempts = 0
        SimSamplePreparedBeta = True
        Exit Function
    End If

    ' THE WORKING COPY IS THE WHOLE OF THE ATOMICITY GUARANTEE. Every rejected
    ' proposal advances it and nothing rewinds it, yet a refusal below still
    ' leaves the caller exactly as it was.
    working = state
    If prepared.UseChengBB Then
        If Not SimSampleChengBB(working, prepared, y, attempts, detail) Then Exit Function
    Else
        If Not SimSampleChengBC(working, prepared, y, attempts, detail) Then Exit Function
    End If

    If Not (y > 0# And y < 1#) Then
        detail = "Beta-PERT: the Cheng stage left the open interval (0, 1)"
        Exit Function
    End If

    ' The stable convex form again, and no clipping of y or of the sample.
    candidate = (1# - y) * prepared.MinValue + y * prepared.MaxValue
    If Not IsUsableDouble(candidate) Then
        detail = "Beta-PERT: the convex rescale is not representable as a finite Double"
        Exit Function
    End If

    sample = candidate
    uniformsConsumed = SIM_CHENG_UNIFORMS_PER_ATTEMPT * attempts
    proposalAttempts = attempts
    state = working
    SimSamplePreparedBeta = True
End Function

' ==========================================================================
' Cheng BB
'
' EVERY ATTEMPT CONSUMES EXACTLY TWO UNIFORMS, u1 then u2, and a rejection
' consumes them just as an acceptance does. The three acceptance tests are
' applied IN THIS ORDER: the cheap squeeze first, then the log test, then the
' full test. Reordering them changes which proposals are accepted, and therefore
' the consumption count and every draw after it.
'
' Log(u1 / (1 - u1)) is the accepted form. Log(u1) - Log(1 - u1) is the same
' function of a real number and a different function of a Double.
'
' The exponent cannot overflow for the accepted shape family: a raw uniform lies
' strictly inside (0, 1) at a distance of about 2.3E-10 from each end, so
' |Log(u1/(1-u1))| stays below 23, and the BB beta term lies in (0.57, 1). The
' result is still checked rather than assumed.
' ==========================================================================
Private Function SimSampleChengBB(ByRef working As SimRngState, _
                                  ByRef prepared As SimSampleBetaShape, _
                                  ByRef y As Double, ByRef attempts As Long, _
                                  ByRef detail As String) As Boolean
    Dim u1 As Double, u2 As Double
    Dim vlog As Double, v As Double, w As Double, z As Double
    Dim rr As Double, ss As Double, t As Double

    attempts = 0
    Do
        attempts = attempts + 1
        If Not SimRngNextUniform(working, u1, detail) Then Exit Function
        If Not SimRngNextUniform(working, u2, detail) Then Exit Function

        vlog = Log(u1 / (1# - u1))
        v = prepared.ChengBeta * vlog
        w = prepared.ChengA * Exp(v)
        z = u1 * u1 * u2
        rr = prepared.ChengGamma * v - SIM_CHENG_BB_LITERAL_1
        ss = prepared.ChengA + rr - w

        If ss + SIM_CHENG_BB_LITERAL_2 >= SIM_CHENG_BB_LITERAL_3 * z Then Exit Do
        t = Log(z)
        If ss >= t Then Exit Do
        If rr + prepared.ChengAlpha * Log(prepared.ChengAlpha / (prepared.ChengB + w)) >= t Then Exit Do
    Loop

    If Not SimSampleOrientedBeta(prepared, w, y, detail) Then Exit Function
    SimSampleChengBB = True
End Function

' ==========================================================================
' Cheng BC
'
' Same two-uniform attempt, same no-rewind rule. The operators are authority:
' u1 < 0.5, z <= 0.25, z >= k2, and a final >=. Loosening or tightening any of
' them moves the acceptance region.
'
' The z <= 0.25 arm accepts IMMEDIATELY, without the final test.
' ==========================================================================
Private Function SimSampleChengBC(ByRef working As SimRngState, _
                                  ByRef prepared As SimSampleBetaShape, _
                                  ByRef y As Double, ByRef attempts As Long, _
                                  ByRef detail As String) As Boolean
    Dim u1 As Double, u2 As Double
    Dim vlog As Double, v As Double, w As Double, z As Double, y0 As Double
    Dim rejected As Boolean

    attempts = 0
    Do
        attempts = attempts + 1
        If Not SimRngNextUniform(working, u1, detail) Then Exit Function
        If Not SimRngNextUniform(working, u2, detail) Then Exit Function

        rejected = False
        If u1 < SIM_CHENG_BC_LITERAL_6 Then
            y0 = u1 * u2
            z = u1 * y0
            If SIM_CHENG_BC_LITERAL_5 * u2 + z - y0 >= prepared.ChengK1 Then rejected = True
        Else
            z = u1 * u1 * u2
            If z <= SIM_CHENG_BC_LITERAL_5 Then
                vlog = Log(u1 / (1# - u1))
                v = prepared.ChengBeta * vlog
                w = prepared.ChengA * Exp(v)
                Exit Do
            End If
            If z >= prepared.ChengK2 Then rejected = True
        End If

        If Not rejected Then
            vlog = Log(u1 / (1# - u1))
            v = prepared.ChengBeta * vlog
            w = prepared.ChengA * Exp(v)
            If prepared.ChengAlpha * (Log(prepared.ChengAlpha / (prepared.ChengB + w)) + v) _
               - SIM_CHENG_BC_LITERAL_4 >= Log(z) Then Exit Do
        End If
    Loop

    If Not SimSampleOrientedBeta(prepared, w, y, detail) Then Exit Function
    SimSampleChengBC = True
End Function

' The accepted return rule, shared so BB and BC cannot drift apart. Which branch
' applies is decided by the ORIENTATION recorded at preparation, not by the
' dispatch: mirroring this returns a valid Beta variate of the wrong shape.
Private Function SimSampleOrientedBeta(ByRef prepared As SimSampleBetaShape, ByVal w As Double, _
                                       ByRef y As Double, ByRef detail As String) As Boolean
    Dim denominator As Double, candidate As Double
    If Not IsUsableDouble(w) Then
        detail = "Beta-PERT: the Cheng proposal is not representable as a finite Double"
        Exit Function
    End If
    denominator = prepared.ChengB + w
    If denominator <= 0# Then
        detail = "Beta-PERT: the Cheng return denominator is not positive"
        Exit Function
    End If
    If prepared.FirstParameterIsOrientedA Then
        candidate = w / denominator
    Else
        candidate = prepared.ChengB / denominator
    End If
    If Not IsUsableDouble(candidate) Then
        detail = "Beta-PERT: the Cheng return is not representable as a finite Double"
        Exit Function
    End If
    y = candidate
    SimSampleOrientedBeta = True
End Function

' ==========================================================================
' Bernoulli occurrence - A PRIMITIVE, NOT D6-18 ORCHESTRATION
'
' One decision. It is not paired with severity here, and nothing in this module
' knows that a Risk has a severity at all.
'
' EXACTLY ONE UNIFORM FOR EVERY VALID PROBABILITY, p = 0 and p = 1 included.
' Skipping the draw at either end would save nothing and would desynchronise the
' component stream against every other run of the same model.
'
' occurred = u < Probability, STRICTLY. Because a raw uniform is strictly inside
' (0, 1), strictness is what makes p = 0 never occur and p = 1 always occur -
' exactly, with no special case anywhere. At u = Probability the answer is False.
' ==========================================================================
Public Function SimSampleBernoulli(ByRef state As SimRngState, ByVal probability As Double, _
                                   ByRef occurred As Boolean, ByRef uniform As Double, _
                                   ByRef uniformsConsumed As Long, _
                                   ByRef detail As String) As Boolean
    Dim working As SimRngState
    Dim u As Double
    detail = vbNullString

    If Not IsUsableDouble(probability) Then
        detail = "Bernoulli: Probability is not a finite Double"
        Exit Function
    End If
    If probability < 0# Or probability > 1# Then
        detail = "Bernoulli: Probability is outside [0, 1]; it is refused, not clamped"
        Exit Function
    End If
    If Not SimRngValidateState(state, detail) Then Exit Function

    working = state
    If Not SimRngNextUniform(working, u, detail) Then Exit Function

    occurred = (u < probability)
    uniform = u
    uniformsConsumed = 1
    state = working
    SimSampleBernoulli = True
End Function

' ==========================================================================
' Shared predicates
' ==========================================================================

' Finite, and ordered Min <= Most Likely <= Max. REFUSED, never repaired:
' silently swapping endpoints turns a data-entry error into a plausible
' distribution nobody asked for. No positivity rule - negative and crossing-zero
' supports are legal.
Private Function SimSampleOrderedTriple(ByVal minValue As Double, ByVal mostLikely As Double, _
                                        ByVal maxValue As Double, ByVal family As String, _
                                        ByRef detail As String) As Boolean
    If Not IsUsableDouble(minValue) Then
        detail = family & ": Min is not a finite Double"
        Exit Function
    End If
    If Not IsUsableDouble(mostLikely) Then
        detail = family & ": Most Likely is not a finite Double"
        Exit Function
    End If
    If Not IsUsableDouble(maxValue) Then
        detail = family & ": Max is not a finite Double"
        Exit Function
    End If
    If minValue > mostLikely Or mostLikely > maxValue Then
        detail = family & ": requires Min <= Most Likely <= Max; the ordering is refused, not repaired"
        Exit Function
    End If
    SimSampleOrderedTriple = True
End Function

' The accepted conditioning scale: s = max(|a|, |m|, |b|), never zero. Working on
' a/s, m/s, b/s is what keeps the shape arithmetic finite for endpoints near
' Double maximum, where the naive products overflow long before the answer does.
Private Function SimSampleScale(ByVal first As Double, ByVal second As Double, _
                                ByVal third As Double) As Double
    Dim s As Double
    s = Abs(first)
    If Abs(second) > s Then s = Abs(second)
    If Abs(third) > s Then s = Abs(third)
    If s <= 0# Then s = 1#
    SimSampleScale = s
End Function

Private Function SimSampleShapeInFamily(ByVal shapeValue As Double, ByRef detail As String) As Boolean
    If Not IsUsableDouble(shapeValue) Then
        detail = "Beta-PERT: a shape parameter is not a finite Double"
        Exit Function
    End If
    If shapeValue < SIM_PERT_SHAPE_LOWER Or shapeValue > SIM_PERT_SHAPE_UPPER Then
        detail = "Beta-PERT: a shape parameter left the accepted family"
        Exit Function
    End If
    SimSampleShapeInFamily = True
End Function

Private Function SimSampleMinOf(ByVal first As Double, ByVal second As Double) As Double
    If first <= second Then
        SimSampleMinOf = first
    Else
        SimSampleMinOf = second
    End If
End Function

Private Function SimSampleMaxOf(ByVal first As Double, ByVal second As Double) As Double
    If first >= second Then
        SimSampleMaxOf = first
    Else
        SimSampleMaxOf = second
    End If
End Function
