Attribute VB_Name = "modCalcAnalytical"
Option Explicit

' ==========================================================================
' modCalcAnalytical - the analytical calculation, and nothing else.
'
' Distribution statistics, the per-driver audit amounts, the five headline
' measures A/B/C/D/E, the six annual series, and the reconciliation identities
' I1-I5. Every number here is produced by the primitives in modCalcFactors, so
' the range rules, the two-tier rescues and the materialization boundary are
' decided in exactly one place.
'
' WHAT THIS MODULE DOES NOT OWN. It never touches a worksheet, never resolves
' Config, never encodes a fingerprint, never writes anything back and knows
' nothing about the user interface or Monte Carlo. It is handed resolved
' numbers and it hands back resolved numbers; a later resolver layer is what
' reads the workbook and what decides when to call it.
'
' TOLERANCES ARE NOT RESTATED HERE. Every tolerance comes from the generated
' TOL_* constants in modCalcContract, whose authority is spec/calc_contract.yaml.
' A tolerance literal written into this module would be a second authority.
' ==========================================================================

' The audit labels for the deterministic central basis. An auditor must never
' have to infer which basis a row used, so the basis is recorded as a word
' rather than left implicit in the distribution name.
Public Const CENTRAL_BASIS_ML As String = "ML"
Public Const CENTRAL_BASIS_MIDPOINT As String = "Midpoint"

' Which stable staged form a convex statistic uses. These are internal to the
' convex-mean helper; the distribution vocabulary itself is DIST_* in
' modCalcFactors.
Private Const CONVEX_THIRDS As Long = 1
Private Const CONVEX_PERT As Long = 2
Private Const CONVEX_HALVES As Long = 3

' The five factors of one per-driver, per-year annual contribution, held
' unmultiplied so the exact rescue can still form the series.
Private Const ANNUAL_FACTOR_COUNT As Long = 5

' --------------------------------------------------------------------------
' Carry types
' --------------------------------------------------------------------------
' One per-driver audit record.
'
' Fields that do not apply to a kind are NOT reused and NOT zeroed to mean
' something: IsRisk decides which half of the record is meaningful. A cost line
' owns Central, Deterministic*, MeanBasis* and Shift*; a risk owns
' ExpectedRisk*. Reading the other half is a caller error, exactly as reading a
' blank _Calc cell would be.
Public Type DriverAudit
    PermanentId As String
    IsRisk As Boolean
    CentralBasis As String
    Central As Double
    MeanValue As Double
    Knom As Double
    Kpv As Double
    DeterministicNominal As Double
    DeterministicPv As Double
    MeanBasisNominal As Double
    MeanBasisPv As Double
    ShiftNominal As Double
    ShiftPv As Double
    ExpectedRiskNominal As Double
    ExpectedRiskPv As Double
End Type

Public Type AnalyticalTotals
    ANom As Double
    APv As Double
    BNom As Double
    BPv As Double
    CNom As Double
    CPv As Double
    DNom As Double
    DPv As Double
    ENom As Double
    EPv As Double
End Type

Public Type AnnualRow
    ProjectIndex As Long
    CalendarYear As Long
    BaseCostNominal As Double
    ExpectedRiskNominal As Double
    TotalNominal As Double
    BaseCostPv As Double
    ExpectedRiskPv As Double
    TotalPv As Double
End Type

' Conditioning magnitudes, captured WHILE the contributions are accumulated.
'
' Erratum C1: every scale sums the scaled absolute magnitudes of the per-driver
' and per-driver-per-year contributions, never the headline totals and never the
' annual row aggregates. Both of those are already-cancelled numbers, and
' conditioning on them would size the tolerance by what survived rather than by
' what happened. Every field already carries the relative coefficient, so the
' raw sum of contributions - which can exceed Double where the tolerance cannot
' - is never formed.
'
' This is internal reconciliation metadata. It is not an audit column and is
' written nowhere.
Public Type ReconciliationMagnitudes
    RelativeCoefficient As Double
    ANom As Double
    APv As Double
    BNom As Double
    BPv As Double
    CNom As Double
    CPv As Double
    DNom As Double
    DPv As Double
    ENom As Double
    EPv As Double
    AnnualBaseNom As Double
    AnnualBasePv As Double
    AnnualRiskNom As Double
    AnnualRiskPv As Double
    AnnualTotalNom As Double
    AnnualTotalPv As Double
End Type

Public Type IdentityCheck
    Label As String
    LeftValue As Double
    RightValue As Double
    Difference As Double
    Allowance As Double
    Holds As Boolean
End Type

' ==========================================================================
' Distribution statistics.
'
' Every one is the STABLE form - Min/3 + ML/3 + Max/3, never (Min+ML+Max)/3 -
' and every one has three tiers:
'
'   TIER 0  a distribution with zero uncertainty returns its single point
'           exactly. No last-ulp drift is acceptable where there is nothing to
'           be uncertain about.
'   TIER 1  the ordinary staged Double path. If it produces a non-zero value
'           that value is returned bit for bit.
'   EXACT   the exact numerator, divided once. Reached when tier 1 overflowed,
'           or when it produced zero - a zero can be a true zero, but it can
'           also be an underflow hiding a small non-zero answer, and only the
'           exact numerator can tell the two apart.
' ==========================================================================
Public Function TriangularMean(ByVal minimum As Double, ByVal mostLikely As Double, _
                               ByVal maximum As Double, ByRef result As Double) As Boolean
    Dim values(0 To 2) As Double, numerator(0 To 2) As Double
    Dim staged As Double, stagedOk As Boolean
    values(0) = minimum
    values(1) = mostLikely
    values(2) = maximum
    If Not UsableOperands(values) Then Exit Function
    If DegeneratePoint(values, result) Then
        TriangularMean = True
        Exit Function
    End If
    stagedOk = StableConvex(values, CONVEX_THIRDS, staged)
    numerator(0) = minimum
    numerator(1) = mostLikely
    numerator(2) = maximum
    TriangularMean = ConvexFinish(stagedOk, staged, numerator, 3#, result)
End Function

Public Function PertMean(ByVal minimum As Double, ByVal mostLikely As Double, _
                         ByVal maximum As Double, ByRef result As Double) As Boolean
    ' The exact numerator is Min + 4*ML + Max, supplied as FOUR COPIES of ML
    ' rather than a multiplication: 4*ML can exceed Double where the mean does
    ' not, and forming it first is the avoidable overflow the stable form exists
    ' to prevent.
    Dim values(0 To 2) As Double, numerator(0 To 5) As Double
    Dim staged As Double, stagedOk As Boolean
    values(0) = minimum
    values(1) = mostLikely
    values(2) = maximum
    If Not UsableOperands(values) Then Exit Function
    If DegeneratePoint(values, result) Then
        PertMean = True
        Exit Function
    End If
    stagedOk = StableConvex(values, CONVEX_PERT, staged)
    numerator(0) = minimum
    numerator(1) = mostLikely
    numerator(2) = mostLikely
    numerator(3) = mostLikely
    numerator(4) = mostLikely
    numerator(5) = maximum
    PertMean = ConvexFinish(stagedOk, staged, numerator, 6#, result)
End Function

Public Function UniformMean(ByVal minimum As Double, ByVal maximum As Double, _
                            ByRef result As Double) As Boolean
    Dim values(0 To 1) As Double, numerator(0 To 1) As Double
    Dim staged As Double, stagedOk As Boolean
    values(0) = minimum
    values(1) = maximum
    If Not UsableOperands(values) Then Exit Function
    If DegeneratePoint(values, result) Then
        UniformMean = True
        Exit Function
    End If
    stagedOk = StableConvex(values, CONVEX_HALVES, staged)
    numerator(0) = minimum
    numerator(1) = maximum
    UniformMean = ConvexFinish(stagedOk, staged, numerator, 2#, result)
End Function

Public Function DistributionMean(ByVal distKind As Long, ByVal minimum As Double, _
                                 ByVal mostLikely As Double, ByVal maximum As Double, _
                                 ByRef result As Double) As Boolean
    ' The distribution mean. Uniform is a TWO-POINT distribution: a populated
    ' Most Likely is accepted and ignored (D1), because the cell may hold a
    ' leftover value from another choice of distribution and refusing it would
    ' block a valid model.
    Select Case distKind
    Case DIST_TRIANGULAR
        DistributionMean = TriangularMean(minimum, mostLikely, maximum, result)
    Case DIST_BETA_PERT
        DistributionMean = PertMean(minimum, mostLikely, maximum, result)
    Case DIST_UNIFORM
        DistributionMean = UniformMean(minimum, maximum, result)
    End Select
End Function

Public Function DeterministicCentral(ByVal distKind As Long, ByVal minimum As Double, _
                                     ByVal mostLikely As Double, ByVal maximum As Double, _
                                     ByRef result As Double, ByRef basis As String) As Boolean
    ' The DETERMINISTIC central value - risks excluded, and never called a mean.
    ' Uniform has no Most Likely, so its central value is the midpoint; every
    ' other accepted distribution centres on Most Likely.
    If distKind = DIST_UNIFORM Then
        basis = CENTRAL_BASIS_MIDPOINT
        DeterministicCentral = UniformMean(minimum, maximum, result)
        Exit Function
    End If
    basis = CENTRAL_BASIS_ML
    If Not IsUsableDouble(mostLikely) Then Exit Function
    result = mostLikely
    DeterministicCentral = True
End Function

Public Function ExpectedRisk(ByVal probability As Double, ByVal severity As Double, _
                             ByVal factor As Double, ByRef result As Double) As Boolean
    ' Probability * severity * escalation factor, as ONE product: staging it as
    ' (probability * severity) first would refuse a risk whose expected value is
    ' representable because an intermediate was not.
    Dim group(0 To 2) As Double
    group(0) = probability
    group(1) = severity
    group(2) = factor
    ExpectedRisk = SafeProduct(group, result)
End Function

Private Function UsableOperands(ByRef values() As Double) As Boolean
    Dim index As Long
    For index = LBound(values) To UBound(values)
        If Not IsUsableDouble(values(index)) Then Exit Function
    Next index
    UsableOperands = True
End Function

Private Function DegeneratePoint(ByRef values() As Double, ByRef point As Double) As Boolean
    ' A distribution with zero uncertainty is its single point, exactly.
    Dim index As Long
    For index = LBound(values) + 1 To UBound(values)
        If values(index) <> values(LBound(values)) Then Exit Function
    Next index
    point = values(LBound(values))
    DegeneratePoint = True
End Function

Private Function StableConvex(ByRef values() As Double, ByVal mode As Long, _
                              ByRef total As Double) As Boolean
    ' TIER 1. Each weight is applied to its own operand before anything is
    ' added, so the sum of the weighted operands cannot overflow where the
    ' statistic itself does not.
    Dim term As Double
    total = 0#
    Select Case mode
    Case CONVEX_THIRDS
        If Not DivideInto(total, values(0), 3#) Then Exit Function
        If Not DivideInto(total, values(1), 3#) Then Exit Function
        If Not DivideInto(total, values(2), 3#) Then Exit Function
    Case CONVEX_PERT
        If Not DivideInto(total, values(0), 6#) Then Exit Function
        If Not SafeMultiply(values(1), 2# / 3#, term) Then Exit Function
        If Not SafeAccumulate(total, term) Then Exit Function
        If Not DivideInto(total, values(2), 6#) Then Exit Function
    Case CONVEX_HALVES
        If Not DivideInto(total, values(0), 2#) Then Exit Function
        If Not DivideInto(total, values(1), 2#) Then Exit Function
    Case Else
        Exit Function
    End Select
    StableConvex = True
End Function

Private Function DivideInto(ByRef total As Double, ByVal value As Double, _
                            ByVal divisor As Double) As Boolean
    Dim term As Double
    If Not SafeDivide(value, divisor, term) Then Exit Function
    DivideInto = SafeAccumulate(total, term)
End Function

Private Function ConvexFinish(ByVal stagedOk As Boolean, ByVal staged As Double, _
                              ByRef numerator() As Double, ByVal divisor As Double, _
                              ByRef result As Double) As Boolean
    ' A staged zero is not trusted. The midpoint of -20 and 19 subnormals is
    ' -0.5 of a subnormal, which every staged evaluation rounds to zero; the
    ' exact numerator knows the difference between that and a true zero.
    If stagedOk Then
        If staged <> 0# Then
            result = staged
            ConvexFinish = True
            Exit Function
        End If
    End If
    ConvexFinish = ExactQuotientOfSum(numerator, divisor, result)
End Function

' ==========================================================================
' Canonical computational order.
'
' Ascending Permanent ID on UTF-16 code units, cost lines first and then risks,
' each group sorted independently. Floating-point addition is not associative,
' so the sequence in which contributions are accumulated is part of the answer
' and not an implementation detail. Row order is deliberately excluded from the
' calculation fingerprint, so two workbooks with the same fingerprint must not
' be able to produce different totals because someone sorted a table.
'
' This is NOT magnitude ordering. Magnitude never decides the order drivers are
' processed in.
' ==========================================================================
Public Function CanonicalOrder(ByRef ids() As String, ByRef isRisk() As Boolean, _
                               ByRef order() As Long) As Boolean
    Dim count As Long, first As Long, index As Long, slot As Long
    Dim probe As Long, moving As Long
    first = LBound(ids)
    count = UBound(ids) - first + 1
    If count < 1 Then Exit Function
    ReDim order(0 To count - 1)
    slot = 0
    For index = 0 To count - 1
        If Not isRisk(LBound(isRisk) + index) Then
            order(slot) = index
            slot = slot + 1
        End If
    Next index
    For index = 0 To count - 1
        If isRisk(LBound(isRisk) + index) Then
            order(slot) = index
            slot = slot + 1
        End If
    Next index
    ' Insertion sort inside each contiguous kind block. Stable and in place; the
    ' block boundary is never crossed, so a risk can never sort ahead of a cost
    ' line.
    For slot = 1 To count - 1
        moving = order(slot)
        probe = slot - 1
        Do While probe >= 0
            If isRisk(LBound(isRisk) + order(probe)) <> isRisk(LBound(isRisk) + moving) Then Exit Do
            If StrComp(ids(first + order(probe)), ids(first + moving), vbBinaryCompare) <= 0 Then Exit Do
            order(probe + 1) = order(probe)
            probe = probe - 1
        Loop
        order(probe + 1) = moving
    Next slot
    CanonicalOrder = True
End Function

Private Function AuditOrder(ByRef audits() As DriverAudit, ByRef order() As Long) As Boolean
    Dim ids() As String, risky() As Boolean, index As Long, count As Long
    count = UBound(audits) - LBound(audits) + 1
    If count < 1 Then Exit Function
    ReDim ids(0 To count - 1)
    ReDim risky(0 To count - 1)
    For index = 0 To count - 1
        ids(index) = audits(LBound(audits) + index).PermanentId
        risky(index) = audits(LBound(audits) + index).IsRisk
    Next index
    AuditOrder = CanonicalOrder(ids, risky, order)
End Function

Private Function DriverOrder(ByRef drivers() As DriverFactors, ByRef order() As Long) As Boolean
    Dim ids() As String, risky() As Boolean, index As Long, count As Long
    count = UBound(drivers) - LBound(drivers) + 1
    If count < 1 Then Exit Function
    ReDim ids(0 To count - 1)
    ReDim risky(0 To count - 1)
    For index = 0 To count - 1
        ids(index) = drivers(LBound(drivers) + index).PermanentId
        risky(index) = drivers(LBound(drivers) + index).IsRisk
    Next index
    DriverOrder = CanonicalOrder(ids, risky, order)
End Function

' ==========================================================================
' The per-driver audit.
'
' Every published per-driver amount is a range boundary in its own right: a
' representable A must not be refused because C overflowed, so each is formed
' as ONE product of its own factors rather than staged through a shared
' intermediate.
'
' Quantity and Probability enter HERE and only here. They are deliberately
' absent from Knom and Kpv - Probability is replaced by a Bernoulli draw in
' Monte Carlo, and Quantity is a per-driver multiplier rather than a factor of
' the escalation path - so folding either into the factors would double-count
' it at this step.
' ==========================================================================
Public Function BuildDriverAudit(ByRef driver As DriverFactors, ByRef audit As DriverAudit, _
                                 ByRef detail As String) As Boolean
    Dim mean As Double, central As Double, shift As Double, basis As String
    detail = vbNullString
    audit.PermanentId = driver.PermanentId
    audit.IsRisk = driver.IsRisk
    audit.Knom = driver.Knom
    audit.Kpv = driver.Kpv

    If Not DistributionMean(driver.DistKind, driver.MinValue, driver.MostLikely, _
                            driver.MaxValue, mean) Then
        detail = "distribution mean"
        Exit Function
    End If
    driver.MeanValue = mean
    audit.MeanValue = mean

    If driver.IsRisk Then
        ' A risk has no deterministic basis at all: the deterministic measures
        ' are defined over cost lines, and a risk contributes only its expected
        ' value. Central and CentralBasis stay unset rather than being given a
        ' zero that would read as a real central value.
        If Not ExpectedRisk(driver.Probability, mean, driver.Knom, _
                            audit.ExpectedRiskNominal) Then
            detail = "expected risk nominal"
            Exit Function
        End If
        If Not ExpectedRisk(driver.Probability, mean, driver.Kpv, audit.ExpectedRiskPv) Then
            detail = "expected risk PV"
            Exit Function
        End If
        BuildDriverAudit = True
        Exit Function
    End If

    If Not DeterministicCentral(driver.DistKind, driver.MinValue, driver.MostLikely, _
                                driver.MaxValue, central, basis) Then
        detail = "deterministic central value"
        Exit Function
    End If
    driver.Central = central
    driver.CentralBasis = basis
    audit.Central = central
    audit.CentralBasis = basis

    If Not SafeSubtract(mean, central, shift) Then
        detail = "uncertainty mean shift"
        Exit Function
    End If

    If Not TripleProduct(central, driver.Quantity, driver.Knom, audit.DeterministicNominal) Then
        detail = "A nominal"
        Exit Function
    End If
    If Not TripleProduct(central, driver.Quantity, driver.Kpv, audit.DeterministicPv) Then
        detail = "A PV"
        Exit Function
    End If
    If Not TripleProduct(shift, driver.Quantity, driver.Knom, audit.ShiftNominal) Then
        detail = "B nominal"
        Exit Function
    End If
    If Not TripleProduct(shift, driver.Quantity, driver.Kpv, audit.ShiftPv) Then
        detail = "B PV"
        Exit Function
    End If
    If Not TripleProduct(mean, driver.Quantity, driver.Knom, audit.MeanBasisNominal) Then
        detail = "C nominal"
        Exit Function
    End If
    If Not TripleProduct(mean, driver.Quantity, driver.Kpv, audit.MeanBasisPv) Then
        detail = "C PV"
        Exit Function
    End If
    BuildDriverAudit = True
End Function

Private Function TripleProduct(ByVal a As Double, ByVal b As Double, ByVal c As Double, _
                               ByRef result As Double) As Boolean
    Dim group(0 To 2) As Double
    group(0) = a
    group(1) = b
    group(2) = c
    TripleProduct = SafeProduct(group, result)
End Function

' ==========================================================================
' A / B / C / D / E - FIVE INDEPENDENT PASSES.
'
' B is NOT computed as C - A, and E is NOT computed as C + D. Those are the
' reconciliation identities, and an identity computed by definition checks
' nothing. Each measure keeps its own named contribution list and is summed
' over that list alone; I1 and I2 are real checks precisely because the two
' journeys to the same number never share a step.
'
' Each contribution is captured into its list, and its scaled absolute
' magnitude into the conditioning scale, at the moment it is produced - so the
' scale measures the arithmetic THAT WAS PERFORMED rather than the total that
' survived it (erratum C1).
' ==========================================================================
Public Function AccumulateTotals(ByRef audits() As DriverAudit, ByRef totals As AnalyticalTotals, _
                                 ByRef magnitudes As ReconciliationMagnitudes, _
                                 ByRef detail As String) As Boolean
    Dim order() As Long, index As Long, slot As Long, who As String
    Dim count As Long, costs As Long, risks As Long, coefficient As Double
    Dim aNomTerms() As Double, aPvTerms() As Double
    Dim bNomTerms() As Double, bPvTerms() As Double
    Dim cNomTerms() As Double, cPvTerms() As Double
    Dim dNomTerms() As Double, dPvTerms() As Double
    Dim eNomTerms() As Double, ePvTerms() As Double
    Dim costSlot As Long, riskSlot As Long

    detail = vbNullString
    coefficient = TOL_IDENTITY_RELATIVE_COEFFICIENT
    ClearMagnitudes magnitudes, coefficient
    count = UBound(audits) - LBound(audits) + 1
    If count < 1 Then
        ClearTotals totals
        AccumulateTotals = True
        Exit Function
    End If
    If Not AuditOrder(audits, order) Then
        detail = "canonical driver order"
        Exit Function
    End If
    For index = 0 To count - 1
        If audits(LBound(audits) + index).IsRisk Then risks = risks + 1 Else costs = costs + 1
    Next index
    If costs > 0 Then
        ReDim aNomTerms(0 To costs - 1)
        ReDim aPvTerms(0 To costs - 1)
        ReDim bNomTerms(0 To costs - 1)
        ReDim bPvTerms(0 To costs - 1)
        ReDim cNomTerms(0 To costs - 1)
        ReDim cPvTerms(0 To costs - 1)
    End If
    If risks > 0 Then
        ReDim dNomTerms(0 To risks - 1)
        ReDim dPvTerms(0 To risks - 1)
    End If
    ReDim eNomTerms(0 To count - 1)
    ReDim ePvTerms(0 To count - 1)

    For slot = 0 To count - 1
        index = LBound(audits) + order(slot)
        who = audits(index).PermanentId
        If audits(index).IsRisk Then
            If Not Contribute(dNomTerms, riskSlot, audits(index).ExpectedRiskNominal, _
                              magnitudes.DNom, coefficient, "D nominal", who, detail) Then Exit Function
            If Not Contribute(dPvTerms, riskSlot, audits(index).ExpectedRiskPv, _
                              magnitudes.DPv, coefficient, "D PV", who, detail) Then Exit Function
            riskSlot = riskSlot + 1
        Else
            If Not Contribute(aNomTerms, costSlot, audits(index).DeterministicNominal, _
                              magnitudes.ANom, coefficient, "A nominal", who, detail) Then Exit Function
            If Not Contribute(aPvTerms, costSlot, audits(index).DeterministicPv, _
                              magnitudes.APv, coefficient, "A PV", who, detail) Then Exit Function
            If Not Contribute(bNomTerms, costSlot, audits(index).ShiftNominal, _
                              magnitudes.BNom, coefficient, "B nominal", who, detail) Then Exit Function
            If Not Contribute(bPvTerms, costSlot, audits(index).ShiftPv, _
                              magnitudes.BPv, coefficient, "B PV", who, detail) Then Exit Function
            If Not Contribute(cNomTerms, costSlot, audits(index).MeanBasisNominal, _
                              magnitudes.CNom, coefficient, "C nominal", who, detail) Then Exit Function
            If Not Contribute(cPvTerms, costSlot, audits(index).MeanBasisPv, _
                              magnitudes.CPv, coefficient, "C PV", who, detail) Then Exit Function
            costSlot = costSlot + 1
        End If
    Next slot

    ' E gets its OWN pass over the same contributions. Two independent journeys
    ' to the same number are what I2 tests; deriving E from C and D here would
    ' turn I2 into a tautology.
    For slot = 0 To count - 1
        index = LBound(audits) + order(slot)
        who = audits(index).PermanentId
        If audits(index).IsRisk Then
            If Not Contribute(eNomTerms, slot, audits(index).ExpectedRiskNominal, _
                              magnitudes.ENom, coefficient, "E nominal", who, detail) Then Exit Function
            If Not Contribute(ePvTerms, slot, audits(index).ExpectedRiskPv, _
                              magnitudes.EPv, coefficient, "E PV", who, detail) Then Exit Function
        Else
            If Not Contribute(eNomTerms, slot, audits(index).MeanBasisNominal, _
                              magnitudes.ENom, coefficient, "E nominal", who, detail) Then Exit Function
            If Not Contribute(ePvTerms, slot, audits(index).MeanBasisPv, _
                              magnitudes.EPv, coefficient, "E PV", who, detail) Then Exit Function
        End If
    Next slot

    If Not SumMeasure(aNomTerms, costs, totals.ANom, "A nominal", detail) Then Exit Function
    If Not SumMeasure(aPvTerms, costs, totals.APv, "A PV", detail) Then Exit Function
    If Not SumMeasure(bNomTerms, costs, totals.BNom, "B nominal", detail) Then Exit Function
    If Not SumMeasure(bPvTerms, costs, totals.BPv, "B PV", detail) Then Exit Function
    If Not SumMeasure(cNomTerms, costs, totals.CNom, "C nominal", detail) Then Exit Function
    If Not SumMeasure(cPvTerms, costs, totals.CPv, "C PV", detail) Then Exit Function
    If Not SumMeasure(dNomTerms, risks, totals.DNom, "D nominal", detail) Then Exit Function
    If Not SumMeasure(dPvTerms, risks, totals.DPv, "D PV", detail) Then Exit Function
    If Not SumMeasure(eNomTerms, count, totals.ENom, "E nominal", detail) Then Exit Function
    If Not SumMeasure(ePvTerms, count, totals.EPv, "E PV", detail) Then Exit Function
    AccumulateTotals = True
End Function

Private Function Contribute(ByRef terms() As Double, ByVal slot As Long, ByVal value As Double, _
                            ByRef scale As Double, ByVal coefficient As Double, _
                            ByVal measure As String, ByVal who As String, _
                            ByRef detail As String) As Boolean
    ' One contribution: into its own measure's list, and into that measure's own
    ' conditioning scale. A failure names the measure and the driver, because a
    ' total that came out unusable is not a diagnosis.
    terms(slot) = value
    If ConditioningScaledMagnitude(scale, value, coefficient) Then
        Contribute = True
        Exit Function
    End If
    detail = measure & ", driver " & who
End Function

Private Function SumMeasure(ByRef terms() As Double, ByVal count As Long, ByRef result As Double, _
                            ByVal measure As String, ByRef detail As String) As Boolean
    If SumSeries(terms, count, result) Then
        SumMeasure = True
        Exit Function
    End If
    detail = measure & " total"
End Function

Private Function SumSeries(ByRef terms() As Double, ByVal count As Long, _
                           ByRef result As Double) As Boolean
    ' An empty measure is zero, not a failure: a model with no risks has D = 0.
    If count < 1 Then
        result = 0#
        SumSeries = True
        Exit Function
    End If
    SumSeries = SafeSignedSum(terms, result)
End Function

Private Sub ClearTotals(ByRef totals As AnalyticalTotals)
    Dim blank As AnalyticalTotals
    totals = blank
End Sub

Private Sub ClearMagnitudes(ByRef magnitudes As ReconciliationMagnitudes, _
                            ByVal coefficient As Double)
    Dim blank As ReconciliationMagnitudes
    magnitudes = blank
    magnitudes.RelativeCoefficient = coefficient
End Sub

' ==========================================================================
' The six annual series.
'
' Per applied project year, on the MEAN basis: the locked Results requirement
' is that annual cash flow is mean-only, so the deterministic basis has no
' annual series at all.
'
' WHAT IS PUBLISHED AND WHAT IS NOT. The annual table publishes six aggregates
' per year and never a per-driver, per-year contribution. So each of the six is
' a representability boundary IN ITS OWN RIGHT - a representable Total does not
' rescue an unrepresentable Base - while an individual contribution is not a
' boundary at all.
'
' Each series therefore has two tiers. Tier 1 forms every contribution as a
' Double and accumulates in canonical driver order, with PV formed as
' nominal * discount; if that complete pipeline succeeds its result is returned
' bit for bit. Only on failure is the series re-formed as one exact sum of
' exact products, with the discount factor INSIDE each PV product so a nominal
' contribution that has no Double never blocks the PV series that does.
'
' Two cost lines at +2 and -2 against an inflation factor of MAX_DOUBLE have
' contributions of +2*MAX_DOUBLE and -2*MAX_DOUBLE and an annual row of 0.
' Refusing that model would be refusing it for its evaluation order (C2).
' ==========================================================================
Public Function BuildAnnualSeries(ByRef drivers() As DriverFactors, ByRef fxRate() As Double, _
                                  ByRef weights() As Double, ByRef inflation() As Double, _
                                  ByRef years() As YearFactors, ByRef rows() As AnnualRow, _
                                  ByRef magnitudes As ReconciliationMagnitudes, _
                                  ByRef detail As String) As Boolean
    Dim order() As Long, count As Long, costs As Long, risks As Long
    Dim yearCount As Long, offset As Long, slot As Long, index As Long
    Dim coefficient As Double, discount As Double
    Dim factorsOf() As Double, group() As Double
    Dim nominal() As Double, present() As Double
    Dim hasNominal() As Boolean, hasPresent() As Boolean

    detail = vbNullString
    coefficient = TOL_IDENTITY_RELATIVE_COEFFICIENT
    ClearMagnitudes magnitudes, coefficient
    yearCount = UBound(years) - LBound(years) + 1
    If yearCount < 1 Then
        detail = "applied timeline"
        Exit Function
    End If
    count = UBound(drivers) - LBound(drivers) + 1
    ReDim rows(0 To yearCount - 1)
    If count < 1 Then
        For offset = 0 To yearCount - 1
            rows(offset).ProjectIndex = years(LBound(years) + offset).ProjectIndex
            rows(offset).CalendarYear = years(LBound(years) + offset).CalendarYear
        Next offset
        BuildAnnualSeries = True
        Exit Function
    End If
    If Not DriverOrder(drivers, order) Then
        detail = "canonical driver order"
        Exit Function
    End If
    For index = 0 To count - 1
        If drivers(LBound(drivers) + index).IsRisk Then risks = risks + 1 Else costs = costs + 1
    Next index

    ReDim factorsOf(0 To count - 1, 0 To ANNUAL_FACTOR_COUNT - 1)
    ReDim nominal(0 To count - 1)
    ReDim present(0 To count - 1)
    ReDim hasNominal(0 To count - 1)
    ReDim hasPresent(0 To count - 1)

    For offset = 0 To yearCount - 1
        discount = years(LBound(years) + offset).DiscountF
        rows(offset).ProjectIndex = years(LBound(years) + offset).ProjectIndex
        rows(offset).CalendarYear = years(LBound(years) + offset).CalendarYear

        ' The FACTORS of every contribution, in canonical order, held before any
        ' of them is multiplied out. Cost lines carry Quantity and risks carry
        ' Probability, which is the only structural difference between the two
        ' contributions.
        For slot = 0 To count - 1
            index = LBound(drivers) + order(slot)
            factorsOf(slot, 0) = drivers(index).MeanValue
            If drivers(index).IsRisk Then
                factorsOf(slot, 1) = drivers(index).Probability
            Else
                factorsOf(slot, 1) = drivers(index).Quantity
            End If
            factorsOf(slot, 2) = fxRate(LBound(fxRate) + order(slot))
            factorsOf(slot, 3) = weights(LBound(weights, 1) + order(slot), LBound(weights, 2) + offset)
            factorsOf(slot, 4) = inflation(LBound(inflation, 1) + order(slot), _
                                           LBound(inflation, 2) + offset)
        Next slot

        ' TIER 1 staging, contribution by contribution. A contribution with no
        ' Double of its own is recorded as absent rather than refused; the
        ' series it feeds decides whether that matters.
        For slot = 0 To count - 1
            group = GroupOf(factorsOf, slot, False, discount)
            hasNominal(slot) = SafeProduct(group, nominal(slot))
            hasPresent(slot) = False
            If hasNominal(slot) Then
                hasPresent(slot) = SafeMultiply(nominal(slot), discount, present(slot))
            End If
        Next slot

        ' C1 conditioning, PER DRIVER PER YEAR and before any row aggregate
        ' exists. Where the contribution has no Double, the coefficient is
        ' folded into the same exact factor expression instead - the annual
        ' aggregate can be 1 where the contributions that produced it were 1e16
        ' apart, and conditioning on the aggregate would report ordinary Double
        ' rounding as a bookkeeping mismatch.
        For slot = 0 To count - 1
            If Not RecordAnnual(factorsOf, slot, False, discount, nominal(slot), _
                                hasNominal(slot), magnitudes, coefficient, _
                                drivers(LBound(drivers) + order(slot)).IsRisk, False) Then
                detail = "conditioning magnitude, nominal, year " & CStr(rows(offset).CalendarYear)
                Exit Function
            End If
            If Not RecordAnnual(factorsOf, slot, True, discount, present(slot), _
                                hasPresent(slot), magnitudes, coefficient, _
                                drivers(LBound(drivers) + order(slot)).IsRisk, True) Then
                detail = "conditioning magnitude, PV, year " & CStr(rows(offset).CalendarYear)
                Exit Function
            End If
        Next slot

        If Not AnnualSeries(nominal, hasNominal, factorsOf, 0, costs, False, discount, _
                            rows(offset).BaseCostNominal) Then
            detail = "annual base nominal, year " & CStr(rows(offset).CalendarYear)
            Exit Function
        End If
        If Not AnnualSeries(nominal, hasNominal, factorsOf, costs, risks, False, discount, _
                            rows(offset).ExpectedRiskNominal) Then
            detail = "annual risk nominal, year " & CStr(rows(offset).CalendarYear)
            Exit Function
        End If
        ' The annual TOTAL is summed over its own contiguous list of
        ' contributions rather than added from the two series above it, so I3c
        ' and I4c stay real checks.
        If Not AnnualSeries(nominal, hasNominal, factorsOf, 0, count, False, discount, _
                            rows(offset).TotalNominal) Then
            detail = "annual total nominal, year " & CStr(rows(offset).CalendarYear)
            Exit Function
        End If
        If Not AnnualSeries(present, hasPresent, factorsOf, 0, costs, True, discount, _
                            rows(offset).BaseCostPv) Then
            detail = "annual base PV, year " & CStr(rows(offset).CalendarYear)
            Exit Function
        End If
        If Not AnnualSeries(present, hasPresent, factorsOf, costs, risks, True, discount, _
                            rows(offset).ExpectedRiskPv) Then
            detail = "annual risk PV, year " & CStr(rows(offset).CalendarYear)
            Exit Function
        End If
        If Not AnnualSeries(present, hasPresent, factorsOf, 0, count, True, discount, _
                            rows(offset).TotalPv) Then
            detail = "annual total PV, year " & CStr(rows(offset).CalendarYear)
            Exit Function
        End If
    Next offset
    BuildAnnualSeries = True
End Function

Private Function GroupOf(ByRef factorsOf() As Double, ByVal slot As Long, _
                         ByVal discounted As Boolean, ByVal discount As Double) As Double()
    Dim out() As Double, position As Long
    If discounted Then
        ReDim out(0 To ANNUAL_FACTOR_COUNT)
    Else
        ReDim out(0 To ANNUAL_FACTOR_COUNT - 1)
    End If
    For position = 0 To ANNUAL_FACTOR_COUNT - 1
        out(position) = factorsOf(slot, position)
    Next position
    If discounted Then out(ANNUAL_FACTOR_COUNT) = discount
    GroupOf = out
End Function

Private Function RecordAnnual(ByRef factorsOf() As Double, ByVal slot As Long, _
                              ByVal discounted As Boolean, ByVal discount As Double, _
                              ByVal value As Double, ByVal present As Boolean, _
                              ByRef magnitudes As ReconciliationMagnitudes, _
                              ByVal coefficient As Double, ByVal isRisk As Boolean, _
                              ByVal isPv As Boolean) As Boolean
    ' Each contribution conditions TWO series: its own kind and the total. The
    ' six annual scales are six separate boundaries, so they are accumulated
    ' separately even though every term is shared with the total.
    Dim group() As Double, ok As Boolean
    group = GroupOf(factorsOf, slot, discounted, discount)
    If isRisk Then
        If isPv Then
            ok = ScaleOne(magnitudes.AnnualRiskPv, group, value, present, coefficient)
            If ok Then ok = ScaleOne(magnitudes.AnnualTotalPv, group, value, present, coefficient)
        Else
            ok = ScaleOne(magnitudes.AnnualRiskNom, group, value, present, coefficient)
            If ok Then ok = ScaleOne(magnitudes.AnnualTotalNom, group, value, present, coefficient)
        End If
    Else
        If isPv Then
            ok = ScaleOne(magnitudes.AnnualBasePv, group, value, present, coefficient)
            If ok Then ok = ScaleOne(magnitudes.AnnualTotalPv, group, value, present, coefficient)
        Else
            ok = ScaleOne(magnitudes.AnnualBaseNom, group, value, present, coefficient)
            If ok Then ok = ScaleOne(magnitudes.AnnualTotalNom, group, value, present, coefficient)
        End If
    End If
    RecordAnnual = ok
End Function

Private Function ScaleOne(ByRef scale As Double, ByRef group() As Double, ByVal value As Double, _
                          ByVal present As Boolean, ByVal coefficient As Double) As Boolean
    If present Then
        ScaleOne = ConditioningScaledMagnitude(scale, value, coefficient)
    Else
        ScaleOne = ConditioningScaledProduct(scale, group, coefficient)
    End If
End Function

Private Function AnnualSeries(ByRef values() As Double, ByRef present() As Boolean, _
                              ByRef factorsOf() As Double, ByVal first As Long, _
                              ByVal count As Long, ByVal discounted As Boolean, _
                              ByVal discount As Double, ByRef result As Double) As Boolean
    Dim terms() As Double, expression() As Variant, group() As Double
    Dim index As Long, complete As Boolean, staged As Double
    If count < 1 Then
        result = 0#
        AnnualSeries = True
        Exit Function
    End If
    complete = True
    For index = 0 To count - 1
        If Not present(first + index) Then
            complete = False
            Exit For
        End If
    Next index
    If complete Then
        ReDim terms(0 To count - 1)
        For index = 0 To count - 1
            terms(index) = values(first + index)
        Next index
        If SafeSignedSum(terms, staged) Then
            result = staged
            AnnualSeries = True
            Exit Function
        End If
    End If
    ReDim expression(0 To count - 1)
    For index = 0 To count - 1
        group = GroupOf(factorsOf, first + index, discounted, discount)
        expression(index) = group
    Next index
    AnnualSeries = ExactSumOfProducts(expression, result)
End Function

' ==========================================================================
' Reconciliation - I1 to I5.
'
' Every allowance is conditioned on the UNDERLYING CONTRIBUTIONS captured during
' accumulation, never on the headline totals and never on the annual row
' aggregates (erratum C1). The magnitudes were captured by the same pass that
' produced the numbers being checked, so they cannot describe a different
' calculation.
'
' All arithmetic here goes through the safe primitives. A reconciliation that
' quietly produced an out-of-range value and then compared it would be
' reporting nonsense.
' ==========================================================================
Public Function Reconcile(ByRef totals As AnalyticalTotals, ByRef rows() As AnnualRow, _
                          ByRef drivers() As DriverFactors, ByRef weights() As Double, _
                          ByRef magnitudes As ReconciliationMagnitudes, _
                          ByRef checks() As IdentityCheck, ByRef detail As String) As Boolean
    Dim order() As Long, count As Long, yearCount As Long
    Dim index As Long, slot As Long, position As Long
    Dim series() As Double, total As Double, scale As Double
    Dim labels(0 To 5) As String
    Dim headline(0 To 5) As Double, annualScale(0 To 5) As Double, headScale(0 To 5) As Double
    Dim seriesValue() As Double

    detail = vbNullString
    If magnitudes.RelativeCoefficient <> TOL_IDENTITY_RELATIVE_COEFFICIENT Then
        ' The scales would describe a different tolerance from the one asked
        ' for, which is a defect in the calculation and not a model refusal.
        detail = "conditioning magnitudes captured at a different coefficient"
        Exit Function
    End If
    count = UBound(drivers) - LBound(drivers) + 1
    yearCount = UBound(rows) - LBound(rows) + 1
    If count < 1 Then
        detail = "no drivers"
        Exit Function
    End If
    ReDim checks(0 To 9 + count)

    If Not TotalIdentity(checks(0), "I1 nominal: A + B = C", totals.ANom, totals.BNom, _
                         totals.CNom, magnitudes.ANom, magnitudes.BNom, magnitudes.CNom, _
                         detail) Then Exit Function
    If Not TotalIdentity(checks(1), "I1 PV: A + B = C", totals.APv, totals.BPv, _
                         totals.CPv, magnitudes.APv, magnitudes.BPv, magnitudes.CPv, _
                         detail) Then Exit Function
    If Not TotalIdentity(checks(2), "I2 nominal: C + D = E", totals.CNom, totals.DNom, _
                         totals.ENom, magnitudes.CNom, magnitudes.DNom, magnitudes.ENom, _
                         detail) Then Exit Function
    If Not TotalIdentity(checks(3), "I2 PV: C + D = E", totals.CPv, totals.DPv, _
                         totals.EPv, magnitudes.CPv, magnitudes.DPv, magnitudes.EPv, _
                         detail) Then Exit Function

    ' I3 and I4: each annual series against the headline it must reproduce.
    ' SIX SEPARATE IDENTITIES, each with its own conditioning scale - a
    ' representable Total series does not vouch for the Base series beside it.
    labels(0) = "I3a nominal base": labels(1) = "I3b nominal risk"
    labels(2) = "I3c nominal total": labels(3) = "I4a PV base"
    labels(4) = "I4b PV risk": labels(5) = "I4c PV total"
    headline(0) = totals.CNom: headline(1) = totals.DNom: headline(2) = totals.ENom
    headline(3) = totals.CPv: headline(4) = totals.DPv: headline(5) = totals.EPv
    headScale(0) = magnitudes.CNom: headScale(1) = magnitudes.DNom
    headScale(2) = magnitudes.ENom: headScale(3) = magnitudes.CPv
    headScale(4) = magnitudes.DPv: headScale(5) = magnitudes.EPv
    annualScale(0) = magnitudes.AnnualBaseNom: annualScale(1) = magnitudes.AnnualRiskNom
    annualScale(2) = magnitudes.AnnualTotalNom: annualScale(3) = magnitudes.AnnualBasePv
    annualScale(4) = magnitudes.AnnualRiskPv: annualScale(5) = magnitudes.AnnualTotalPv

    If yearCount > 0 Then
        ReDim seriesValue(0 To 5, 0 To yearCount - 1)
        For index = 0 To yearCount - 1
            seriesValue(0, index) = rows(LBound(rows) + index).BaseCostNominal
            seriesValue(1, index) = rows(LBound(rows) + index).ExpectedRiskNominal
            seriesValue(2, index) = rows(LBound(rows) + index).TotalNominal
            seriesValue(3, index) = rows(LBound(rows) + index).BaseCostPv
            seriesValue(4, index) = rows(LBound(rows) + index).ExpectedRiskPv
            seriesValue(5, index) = rows(LBound(rows) + index).TotalPv
        Next index
    End If
    For position = 0 To 5
        If yearCount > 0 Then
            ReDim series(0 To yearCount - 1)
            For index = 0 To yearCount - 1
                series(index) = seriesValue(position, index)
            Next index
        End If
        ' SIGNED: annual rows carry either sign, and the reconciliation must not
        ' refuse a model whose headline it is about to confirm (erratum C2).
        If Not SumSeries(series, yearCount, total) Then
            detail = labels(position)
            Exit Function
        End If
        If Not Pair(annualScale(position), 0#, headScale(position), scale) Then
            detail = labels(position) & " scale"
            Exit Function
        End If
        If Not Identity(checks(4 + position), labels(position), total, headline(position), _
                        scale) Then
            detail = labels(position)
            Exit Function
        End If
    Next position

    ' I5: each driver's profile weights must sum to 1. Its allowance is the
    ' profiling tolerance - an absolute tolerance on a normalised sum, which is
    ' conditioned on nothing.
    If Not DriverOrder(drivers, order) Then
        detail = "canonical driver order"
        Exit Function
    End If
    For slot = 0 To count - 1
        index = LBound(drivers) + order(slot)
        ReDim series(0 To UBound(weights, 2) - LBound(weights, 2))
        For position = 0 To UBound(series)
            series(position) = weights(LBound(weights, 1) + order(slot), _
                                       LBound(weights, 2) + position)
        Next position
        If Not SafeSignedSum(series, total) Then
            detail = "I5 profile sum, driver " & drivers(index).PermanentId
            Exit Function
        End If
        checks(10 + slot).Label = "I5 profile sum: " & drivers(index).PermanentId
        checks(10 + slot).LeftValue = total
        checks(10 + slot).RightValue = 1#
        If Not SafeSubtract(total, 1#, checks(10 + slot).Difference) Then
            detail = "I5 profile sum, driver " & drivers(index).PermanentId
            Exit Function
        End If
        checks(10 + slot).Allowance = TOL_PROFILING_SUM_ABSOLUTE
        checks(10 + slot).Holds = Abs(checks(10 + slot).Difference) <= checks(10 + slot).Allowance
    Next slot
    Reconcile = True
End Function

Public Function AllIdentitiesHold(ByRef checks() As IdentityCheck) As Boolean
    Dim index As Long
    For index = LBound(checks) To UBound(checks)
        If Not checks(index).Holds Then Exit Function
    Next index
    AllIdentitiesHold = True
End Function

Private Function TotalIdentity(ByRef check As IdentityCheck, ByVal label As String, _
                               ByVal firstTerm As Double, ByVal secondTerm As Double, _
                               ByVal against As Double, ByVal firstScale As Double, _
                               ByVal secondScale As Double, ByVal againstScale As Double, _
                               ByRef detail As String) As Boolean
    Dim summed As Double, scale As Double
    detail = label
    If Not SafeAdd(firstTerm, secondTerm, summed) Then Exit Function
    If Not Pair(firstScale, secondScale, againstScale, scale) Then Exit Function
    If Not Identity(check, label, summed, against, scale) Then Exit Function
    detail = vbNullString
    TotalIdentity = True
End Function

Private Function Identity(ByRef check As IdentityCheck, ByVal label As String, _
                          ByVal leftValue As Double, ByVal rightValue As Double, _
                          ByVal scale As Double) As Boolean
    Dim allowance As Double, difference As Double
    If Not IdentityAllowance(scale, TOL_IDENTITY_ABSOLUTE_FLOOR, _
                             TOL_IDENTITY_RELATIVE_COEFFICIENT, _
                             TOL_CONDITIONING_SCALE_FLOOR, allowance) Then Exit Function
    If Not SafeSubtract(leftValue, rightValue, difference) Then Exit Function
    check.Label = label
    check.LeftValue = leftValue
    check.RightValue = rightValue
    check.Difference = difference
    check.Allowance = allowance
    check.Holds = Abs(difference) <= allowance
    Identity = True
End Function

Private Function Pair(ByVal firstScale As Double, ByVal secondScale As Double, _
                      ByVal thirdScale As Double, ByRef scale As Double) As Boolean
    ' The already-scaled magnitudes of the measures an identity accumulated.
    ' They arrive pre-multiplied by the coefficient, so this is a plain sum and
    ' the coefficient is never applied twice.
    scale = 0#
    If Not SafeAccumulate(scale, firstScale) Then Exit Function
    If Not SafeAccumulate(scale, secondScale) Then Exit Function
    Pair = SafeAccumulate(scale, thirdScale)
End Function
