Attribute VB_Name = "modSimAnnualStore"
Option Explicit

' ==========================================================================
' PCCM Phase 7 - WHERE THE ANNUAL ANSWER LIVES.
'
' Every cell the annual step reads or writes is addressed here, and nowhere
' else. The producer beside it - modSimAnnualRun - owns the pipeline and names
' no column, no row and no sheet; this module owns the addresses and computes
' none of the numbers that go into them.
'
' It also answers Phase 8's question, because the answer is a property of what
' is STORED: which run the persisted block belongs to, how many years it covers,
' and which confidence level its profile is the profile for.
'
' --------------------------------------------------------------------------
' EVERY ADDRESS COMES FROM THE PROJECTED CONTRACT
' --------------------------------------------------------------------------
' Not one column letter or row number is typed into this module. They arrive as
' SIM_ANNUAL_* constants projected from sim_contract.yaml, and the block's own
' width is DERIVED from them rather than counted out - which is exactly the
' discipline whose absence left the sensitivity availability formula pointing at
' the statistics band after the sensitivity block moved.
'
' --------------------------------------------------------------------------
' IT OWNS NO MATHEMATICS AND NO STATE DERIVATION
' --------------------------------------------------------------------------
' No RNG, no sampler, no percentile, no factor, no reconciliation. It does not
' derive the SIMULATION's state either: that belongs to modSimReport and is
' asked for, never re-derived. What it does derive is which of the annual
' STATES the stored block is in, and those rules are separated from the cells
' that feed them so every branch can be exercised without a workbook.
' ==========================================================================

' The published identity an annual answer belongs to, plus the two things the
' sensitivity block has no equivalent of: how many project years the answer
' covers, and WHICH confidence level its profile is the profile for.
Public Type SimAnnualIdentity
    Bank As String
    RunId As Long
    EffectiveSeed As Long
    RequestFingerprint As String
    ResultDigest As String
    Iterations As Long
    YearCount As Long
    SelectedLabel As String
    SelectedProbability As Double
End Type

' ==========================================================================
' THE PRECONDITION - THE CURRENT RUN, AND NOTHING ELSE
'
' The status is derived by modSimReport and is not re-derived here. Producing
' annual answers for the current model against totals published by a different
' one would be wrong in a way nothing on the sheet could reveal.
'
' A REFUSAL IS AN ATTEMPT OUTCOME, NOT A STATE. Nothing is written, no
' simulation state changes, and an annual answer belonging to an earlier
' successful run is left exactly where it is.
' ==========================================================================
Public Function SimAnnualStoreCurrentRun(ByRef run As SimAnnualIdentity, _
                                         ByRef detail As String) As Boolean
    Dim status As String

    status = modSimReport.PCCM_SimulationStatus()
    If Len(status) = 0 Then
        detail = "annual: no successful simulation has been published, so there " & _
                 "is nothing to decompose"
        Exit Function
    End If
    If StrComp(status, SIM_STATE_CURRENT, vbBinaryCompare) <> 0 Then
        detail = "annual: the simulation is " & status & ". The annual answer " & _
                 "belongs to one run, and producing it for the current model " & _
                 "against totals published by a different one would be wrong in " & _
                 "a way the sheet could not show. Run the simulation again first."
        Exit Function
    End If
    SimAnnualStoreCurrentRun = SimAnnualStoreIdentity(run, detail)
End Function

Public Function SimAnnualStoreIdentity(ByRef run As SimAnnualIdentity, _
                                       ByRef detail As String) As Boolean
    Dim bank As String

    bank = SharedText(SIM_IDENTITY_ROW_ACTIVE_BANK)
    If Not IsBank(bank) Then
        detail = "annual: the active publication bank is not readable"
        Exit Function
    End If
    run.Bank = bank
    run.RequestFingerprint = modSimReport.PCCM_SimulationRequestFingerprint()
    run.ResultDigest = modSimReport.PCCM_SimulationResultDigest()
    If Len(run.RequestFingerprint) = 0 Or Len(run.ResultDigest) = 0 Then
        detail = "annual: the published run carries no identity to bind to"
        Exit Function
    End If
    ' EACH QUANTITY THROUGH ITS OWN DOMAIN. A run id, a seed and an iteration
    ' count are three different quantities with three different ranges, and
    ' reading all three through the iteration ceiling is the P7-4 defect that
    ' refused every AUTO run from the third one onward.
    If Not SnapshotLong(bank, SIM_IDENTITY_ROW_RUN_ID, CDbl(SIM_RUN_ID_FIRST), _
                        CDbl(SIM_RUN_ID_MAXIMUM), "run id", _
                        run.RunId, detail) Then Exit Function
    If Not SnapshotLong(bank, SIM_IDENTITY_ROW_EFFECTIVE_SEED, CDbl(SIM_SEED_MIN), _
                        CDbl(SIM_SEED_MAX), "effective seed", _
                        run.EffectiveSeed, detail) Then Exit Function
    If Not SnapshotLong(bank, SIM_IDENTITY_ROW_ITERATIONS_RUN, CDbl(SIM_MIN_ITERATIONS), _
                        CDbl(SIM_MAX_ITERATIONS), "iteration count", _
                        run.Iterations, detail) Then Exit Function
    SimAnnualStoreIdentity = True
End Function

' ==========================================================================
' THE REPORTING SELECTOR, AS TEXT
'
' READ, NEVER INTERPRETED. Which labels are selectable, and what probability a
' label spells, belong to modSimStats and its one projected ladder. This module
' hands over the cell's text and decides nothing about it.
' ==========================================================================
Public Function SimAnnualStoreSelector(ByRef label As String, _
                                       ByRef detail As String) As Boolean
    If Not modWorkbook.NameExists(NM_INPUT_SELECTED_CONFIDENCE_LEVEL) Then
        detail = "annual: the Selected Confidence Level input is not defined"
        Exit Function
    End If
    label = modWorkbook.TextOf(modWorkbook.NamedCell(NM_INPUT_SELECTED_CONFIDENCE_LEVEL))
    SimAnnualStoreSelector = True
End Function

' ==========================================================================
' THE PROJECT-YEAR AXIS - READ BACK, NOT REBUILT
'
' The project index, the calendar year and the discount factor are Phase-5
' output, and the annual step runs only on a CURRENT Phase 5, so the published
' year table IS the axis this model resolves to.
'
' REBUILDING IT WOULD BE A SECOND CONSTRUCTION. No timeline is resolved here, no
' discount rate is consulted and no discount series is built: the discount
' factor Phase 7 uses is the one Kpv was built from, because it is literally the
' same number read back. This is the discipline that makes sensitivity read the
' published iteration totals rather than replaying the run to recover them.
' ==========================================================================
Public Function SimAnnualStoreYearAxis(ByRef projectIndex() As Long, _
                                       ByRef calendarYear() As Long, _
                                       ByRef discount() As Double, _
                                       ByRef yearCount As Long, _
                                       ByRef detail As String) As Boolean
    Dim table As ListObject
    Dim index As Long
    Dim measured As Double

    If Not modWorkbook.LoExists(CALC_SHEET, TBL_CALC_YEARS) Then
        detail = "annual: the analytical project-year table is not present"
        Exit Function
    End If
    Set table = modWorkbook.Lo(CALC_SHEET, TBL_CALC_YEARS)
    yearCount = modWorkbook.BodyRowCount(table)
    If yearCount < 1 Then
        detail = "annual: the analytical project-year table carries no year"
        Exit Function
    End If
    If yearCount > LIMIT_MAX_YEAR_COLUMNS Then
        detail = "annual: the analytical project-year table carries " & CStr(yearCount) & _
                 " years, beyond the structural maximum of " & CStr(LIMIT_MAX_YEAR_COLUMNS)
        Exit Function
    End If

    ReDim projectIndex(0 To yearCount - 1)
    ReDim calendarYear(0 To yearCount - 1)
    ReDim discount(0 To yearCount - 1)
    For index = 1 To yearCount
        If Not modWorkbook.IsWholeInRange( _
                modWorkbook.CellIn(table, index, COL_CALC_YEARS_PROJECT_INDEX).Value2, _
                1#, CDbl(LIMIT_MAX_YEAR_COLUMNS), measured) Then
            detail = "annual: project year " & CStr(index) & " carries no usable index"
            Exit Function
        End If
        projectIndex(index - 1) = modWorkbook.SafeLong(measured)
        If Not modWorkbook.IsWholeInRange( _
                modWorkbook.CellIn(table, index, COL_CALC_YEARS_CALENDAR_YEAR).Value2, _
                CDbl(LIMIT_MIN_YEAR), CDbl(LIMIT_MAX_YEAR), measured) Then
            detail = "annual: project year " & CStr(index) & _
                     " carries no usable calendar year"
            Exit Function
        End If
        calendarYear(index - 1) = modWorkbook.SafeLong(measured)
        If Not modWorkbook.TryReadDouble( _
                modWorkbook.CellIn(table, index, COL_CALC_YEARS_DISCOUNT_FACTOR).Value2, _
                measured) Then
            detail = "annual: project year " & CStr(index) & _
                     " carries no usable discount factor"
            Exit Function
        End If
        discount(index - 1) = measured
    Next index
    SimAnnualStoreYearAxis = True
End Function

' ==========================================================================
' THE PERSISTED TOTALS OF ONE MEASURE, in original iteration order
'
' The order matters and is never disturbed: the result digest depends on it, and
' so does which iteration owns an order statistic.
' ==========================================================================
Public Function SimAnnualStoreTotals(ByRef run As SimAnnualIdentity, _
                                     ByVal measure As String, ByRef totals() As Double, _
                                     ByRef detail As String) As Boolean
    Dim block As Variant
    Dim column As String
    Dim index As Long
    Dim value As Double

    If Not TotalColumn(run.Bank, measure, column, detail) Then Exit Function
    block = AnnualSheet().Range(column & CStr(SIM_DATA_FIRST_ITERATION_ROW) & ":" & _
                                column & CStr(SIM_DATA_FIRST_ITERATION_ROW + _
                                              run.Iterations - 1)).Value2
    ReDim totals(0 To run.Iterations - 1)
    For index = 1 To run.Iterations
        If Not modWorkbook.TryReadDouble(block(index, 1), value) Then
            detail = "annual: iteration " & CStr(index) & " of the published " & _
                     measure & " total is not a usable number"
            Exit Function
        End If
        totals(index - 1) = value
    Next index
    SimAnnualStoreTotals = True
End Function

' ==========================================================================
' THE RECORD BLOCK, FLATTENED THROUGH THE CONTRACT'S OWN OFFSETS
'
' Row-major, `fields` values per project year. EVERY offset is derived from the
' projected columns; none is counted out by hand. A contract move therefore
' reaches this writer without anyone retyping a letter.
' ==========================================================================
Public Function SimAnnualStoreFlatten(ByRef run As SimAnnualIdentity, _
                                      ByRef projectIndex() As Long, _
                                      ByRef calendarYear() As Long, _
                                      ByRef nominalLadder() As Double, _
                                      ByRef pvLadder() As Double, _
                                      ByRef nominalProfile() As Double, _
                                      ByRef pvProfile() As Double, _
                                      ByRef flat() As Double, ByRef fields As Long, _
                                      ByRef detail As String) As Boolean
    Dim indexAt As Long, yearAt As Long, nominalAt As Long, pvAt As Long
    Dim nominalProfileAt As Long, pvProfileAt As Long
    Dim year As Long, rung As Long, origin As Long

    indexAt = OffsetOf(run.Bank, FirstColumn(run.Bank))
    yearAt = OffsetOf(run.Bank, CalendarYearColumn(run.Bank))
    nominalAt = OffsetOf(run.Bank, NominalFirstColumn(run.Bank))
    pvAt = OffsetOf(run.Bank, PvFirstColumn(run.Bank))
    nominalProfileAt = OffsetOf(run.Bank, NominalProfileColumn(run.Bank))
    pvProfileAt = OffsetOf(run.Bank, PvProfileColumn(run.Bank))
    fields = OffsetOf(run.Bank, LastColumn(run.Bank)) + 1
    If Not LayoutIsSound(indexAt, yearAt, nominalAt, pvAt, nominalProfileAt, _
                         pvProfileAt, fields, detail) Then Exit Function

    ReDim flat(0 To run.YearCount * fields - 1)
    For year = 0 To run.YearCount - 1
        origin = year * fields
        flat(origin + indexAt) = CDbl(projectIndex(LBound(projectIndex) + year))
        flat(origin + yearAt) = CDbl(calendarYear(LBound(calendarYear) + year))
        flat(origin + nominalProfileAt) = nominalProfile(LBound(nominalProfile) + year)
        flat(origin + pvProfileAt) = pvProfile(LBound(pvProfile) + year)
        For rung = 0 To SIM_ANNUAL_QUANTILE_COUNT - 1
            flat(origin + nominalAt + rung) = _
                nominalLadder(LBound(nominalLadder) + year * SIM_ANNUAL_QUANTILE_COUNT + rung)
            flat(origin + pvAt + rung) = _
                pvLadder(LBound(pvLadder) + year * SIM_ANNUAL_QUANTILE_COUNT + rung)
        Next rung
    Next year
    SimAnnualStoreFlatten = True
End Function

' EVERY CONTRACTED FIELD INSIDE THE BLOCK, NO TWO ON TOP OF EACH OTHER, AND NO
' COLUMN LEFT UNWRITTEN.
'
' This is not a restatement of the layout - it is a refusal to write through a
' layout that has stopped making sense. A contract edit that overlapped the PV
' ladder with the nominal one would otherwise publish one measure twice and the
' other never, with every value on the sheet looking entirely plausible.
Private Function LayoutIsSound(ByVal indexAt As Long, ByVal yearAt As Long, _
                               ByVal nominalAt As Long, ByVal pvAt As Long, _
                               ByVal nominalProfileAt As Long, ByVal pvProfileAt As Long, _
                               ByVal fields As Long, ByRef detail As String) As Boolean
    Dim used() As Long
    Dim slot As Long, rung As Long, needed As Long

    ' Two index columns, a ladder per measure, a profile value per measure.
    needed = 2 + 2 * SIM_ANNUAL_QUANTILE_COUNT + 2
    If fields <> needed Then
        detail = "annual: the record block spans " & CStr(fields) & _
                 " column(s) and the contracted fields need " & CStr(needed)
        Exit Function
    End If
    ' AND NO COLUMN IS LEFT UNWRITTEN - which follows and is not scanned for.
    ' The count above is exactly the number of claims made below, every claim is
    ' proved to be inside the block, and no two may share a slot. Twenty-six
    ' distinct in-range claims over twenty-six columns leave none over. A loop
    ' looking for a gap here could never find one, and code no input can reach is
    ' code no test can check.
    ReDim used(0 To fields - 1)
    If Not Claim(used, fields, indexAt, "the project index", detail) Then Exit Function
    If Not Claim(used, fields, yearAt, "the calendar year", detail) Then Exit Function
    If Not Claim(used, fields, nominalProfileAt, "the nominal profile", detail) Then Exit Function
    If Not Claim(used, fields, pvProfileAt, "the PV profile", detail) Then Exit Function
    For rung = 0 To SIM_ANNUAL_QUANTILE_COUNT - 1
        slot = nominalAt + rung
        If Not Claim(used, fields, slot, "a nominal ladder rung", detail) Then Exit Function
        slot = pvAt + rung
        If Not Claim(used, fields, slot, "a PV ladder rung", detail) Then Exit Function
    Next rung
    LayoutIsSound = True
End Function

Private Function Claim(ByRef used() As Long, ByVal fields As Long, ByVal slot As Long, _
                       ByVal what As String, ByRef detail As String) As Boolean
    If slot < 0 Or slot >= fields Then
        detail = "annual: " & what & " is at column offset " & CStr(slot) & _
                 ", outside the record block"
        Exit Function
    End If
    If used(slot) <> 0 Then
        detail = "annual: " & what & " and another contracted field both write " & _
                 "column offset " & CStr(slot)
        Exit Function
    End If
    used(slot) = 1
    Claim = True
End Function

' ==========================================================================
' PUBLICATION
'
' Clear the marker, clear the block, write the records, write the stamp, and
' mark it published LAST. Between the first write and the last the block is
' explicitly NOT published, so an interruption anywhere in between leaves it
' unreadable rather than half-true.
'
' THE CLEAR SPANS THE WHOLE STRUCTURAL MAXIMUM, not the rows about to be
' written. A twenty-year run followed by a four-year one would otherwise leave
' years five to twenty on the sheet, indistinguishable from the new answer. The
' stamped year count is the second guard: it says where the answer stops, and
' the last non-blank row never does.
'
' IT TOUCHES ONE BANK. Every address is the run's own bank's, so the other
' bank's annual block - which may still be the answer for a different run - is
' not read, not cleared and not written.
' ==========================================================================
Public Function SimAnnualStorePublish(ByRef run As SimAnnualIdentity, _
                                      ByRef flat() As Double, ByVal fields As Long, _
                                      ByRef detail As String) As Boolean
    Dim block As Variant
    Dim year As Long, slot As Long
    Dim first As String, last As String

    If run.YearCount < 1 Then
        detail = "annual: there is no project year to publish"
        Exit Function
    End If
    first = FirstColumn(run.Bank)
    last = LastColumn(run.Bank)

    StampCell(run.Bank, SIM_ANNUAL_STAMP_ROW_PUBLISHED).Value2 = vbNullString
    AnnualSheet().Range(first & CStr(SIM_ANNUAL_FIRST_ROW) & ":" & last & _
                        CStr(SIM_ANNUAL_FIRST_ROW + LIMIT_MAX_YEAR_COLUMNS - 1)).ClearContents

    ReDim block(1 To run.YearCount, 1 To fields)
    For year = 1 To run.YearCount
        For slot = 1 To fields
            block(year, slot) = flat(LBound(flat) + (year - 1) * fields + slot - 1)
        Next slot
    Next year
    AnnualSheet().Range(first & CStr(SIM_ANNUAL_FIRST_ROW) & ":" & last & _
                        CStr(SIM_ANNUAL_FIRST_ROW + run.YearCount - 1)).Value2 = block

    StampCell(run.Bank, SIM_ANNUAL_STAMP_ROW_RUN_ID).Value2 = run.RunId
    StampCell(run.Bank, SIM_ANNUAL_STAMP_ROW_EFFECTIVE_SEED).Value2 = run.EffectiveSeed
    StampCell(run.Bank, SIM_ANNUAL_STAMP_ROW_REQUEST_FINGERPRINT).Value2 = _
        run.RequestFingerprint
    StampCell(run.Bank, SIM_ANNUAL_STAMP_ROW_RESULT_DIGEST).Value2 = run.ResultDigest
    StampCell(run.Bank, SIM_ANNUAL_STAMP_ROW_ITERATIONS).Value2 = run.Iterations
    StampCell(run.Bank, SIM_ANNUAL_STAMP_ROW_YEAR_COUNT).Value2 = run.YearCount
    ' WHICH Px THE PROFILE IS THE PROFILE FOR. Both halves, because both must
    ' agree with the authoritative resolution before a reader may call the
    ' profile current, and a stamp that disagrees with itself says nothing.
    StampCell(run.Bank, SIM_ANNUAL_STAMP_ROW_SELECTED_PX_LABEL).Value2 = run.SelectedLabel
    StampCell(run.Bank, SIM_ANNUAL_STAMP_ROW_SELECTED_PX_PROBABILITY).Value2 = _
        run.SelectedProbability

    ' AND ONLY NOW.
    StampCell(run.Bank, SIM_ANNUAL_STAMP_ROW_PUBLISHED).Value2 = SIM_ANNUAL_PUBLISHED
    SimAnnualStorePublish = True
End Function

' ==========================================================================
' ==========================================================================
' THE PHASE-8 HANDOFF
' ==========================================================================
' ==========================================================================
' FIVE SITUATIONS, AND A SINGLE BOOLEAN CANNOT CARRY THEM:
'
'   the annual distributions are current for the current successful run
'   the selected-Px profile is current for the selector as it stands NOW
'   a profile exists and belongs to a DIFFERENT confidence level
'   no annual output was produced for the current successful run
'   the annual output on the sheet belongs to an earlier run
'
' So there are two state accessors, because there are two products with two
' different currentness rules, plus the stamped Px itself - which is what lets a
' reader NAME the level a profile belongs to instead of assuming it is the one
' selected now.
'
' MOVING THE SELECTOR DOES NOT RETIRE THE LADDERS. Every rung of every year is
' taken across all iterations and no selector enters it, so a moved selector
' cannot make a ladder wrong and must not be allowed to make it stale. It does
' retire the profile, which is the blend at ONE resolved Px - and the profile is
' never relabelled as the newly selected level.
'
' THESE ARE READS. They run nothing, write nothing, consume no run identity and
' derive no simulation state.
' ==========================================================================
Public Function PCCM_AnnualDistributionState() As String
    Dim run As SimAnnualIdentity
    Dim detail As String
    Dim bank As String

    bank = SharedText(SIM_IDENTITY_ROW_ACTIVE_BANK)
    If Not IsBank(bank) Then
        PCCM_AnnualDistributionState = SIM_ANNUAL_STATE_NOT_PRODUCED
        Exit Function
    End If
    PCCM_AnnualDistributionState = DistributionStateOf( _
        StampText(bank, SIM_ANNUAL_STAMP_ROW_PUBLISHED), _
        SimAnnualStoreCurrentRun(run, detail), StampBelongsTo(bank, run))
End Function

Public Function PCCM_AnnualProfileState() As String
    Dim run As SimAnnualIdentity
    Dim detail As String
    Dim bank As String, stamped As String, selector As String, distribution As String
    Dim stampedProbability As Double, ownProbability As Double, resolvedProbability As Double
    Dim consistent As Boolean, resolved As Boolean

    bank = SharedText(SIM_IDENTITY_ROW_ACTIVE_BANK)
    If Not IsBank(bank) Then
        PCCM_AnnualProfileState = SIM_ANNUAL_STATE_NOT_PRODUCED
        Exit Function
    End If
    distribution = DistributionStateOf( _
        StampText(bank, SIM_ANNUAL_STAMP_ROW_PUBLISHED), _
        SimAnnualStoreCurrentRun(run, detail), StampBelongsTo(bank, run))

    stamped = StampText(bank, SIM_ANNUAL_STAMP_ROW_SELECTED_PX_LABEL)
    stampedProbability = StampNumber(bank, SIM_ANNUAL_STAMP_ROW_SELECTED_PX_PROBABILITY)
    ' THE STAMP AGAINST ITSELF. A label and a probability that disagree describe
    ' no confidence level at all, and a reader must not be told the profile is
    ' the selected one on the strength of whichever half happens to match.
    consistent = modSimStats.SimStatsSelectedProbability(stamped, ownProbability, detail)
    If consistent Then consistent = (ownProbability = stampedProbability)

    resolved = SimAnnualStoreSelector(selector, detail)
    If resolved Then
        resolved = modSimStats.SimStatsSelectedProbability(selector, resolvedProbability, _
                                                           detail)
    End If

    PCCM_AnnualProfileState = ProfileStateOf(distribution, consistent, resolved, stamped, _
                                             stampedProbability, selector, _
                                             resolvedProbability)
End Function

' THE Px THE PERSISTED PROFILE IS THE PROFILE FOR - the stamp's own label,
' whatever the selector says now. Blank when no profile is published.
Public Function PCCM_AnnualProfilePx() As String
    Dim bank As String

    bank = SharedText(SIM_IDENTITY_ROW_ACTIVE_BANK)
    If Not IsBank(bank) Then Exit Function
    If Not IsPublished(bank) Then Exit Function
    PCCM_AnnualProfilePx = StampText(bank, SIM_ANNUAL_STAMP_ROW_SELECTED_PX_LABEL)
End Function

' HOW MANY YEARS THE PERSISTED ANSWER COVERS. The stamped count, never the last
' non-blank row: a shorter answer written over a longer one leaves nothing
' behind, and the count is what says so.
Public Function PCCM_AnnualYearCount() As Long
    Dim bank As String
    Dim measured As Double

    bank = SharedText(SIM_IDENTITY_ROW_ACTIVE_BANK)
    If Not IsBank(bank) Then Exit Function
    If Not IsPublished(bank) Then Exit Function
    If Not modWorkbook.IsWholeInRange( _
            StampCell(bank, SIM_ANNUAL_STAMP_ROW_YEAR_COUNT).Value2, _
            1#, CDbl(LIMIT_MAX_YEAR_COLUMNS), measured) Then Exit Function
    PCCM_AnnualYearCount = modWorkbook.SafeLong(measured)
End Function

' ==========================================================================
' THE STATE RULES - no worksheet, no arithmetic, no derivation
'
' They are separated from the cells that feed them so that every branch of the
' settlement can be exercised directly, over this source rather than over a
' description of it.
' ==========================================================================
Private Function DistributionStateOf(ByVal published As String, _
                                     ByVal simulationCurrent As Boolean, _
                                     ByVal identityMatches As Boolean) As String
    ' NOTHING PUBLISHED IS NOT THE SAME AS PUBLISHED-BUT-OLD, and a reader has
    ' to act differently on them: one asks for a first run of the annual step,
    ' the other warns that what is on the sheet is somebody else's answer.
    If StrComp(published, SIM_ANNUAL_PUBLISHED, vbBinaryCompare) <> 0 Then
        DistributionStateOf = SIM_ANNUAL_STATE_NOT_PRODUCED
        Exit Function
    End If
    If simulationCurrent And identityMatches Then
        DistributionStateOf = SIM_ANNUAL_STATE_CURRENT
        Exit Function
    End If
    DistributionStateOf = SIM_ANNUAL_STATE_HISTORICAL
End Function

Private Function ProfileStateOf(ByVal distribution As String, _
                                ByVal stampConsistent As Boolean, _
                                ByVal selectorResolved As Boolean, _
                                ByVal stampedLabel As String, _
                                ByVal stampedProbability As Double, _
                                ByVal resolvedLabel As String, _
                                ByVal resolvedProbability As Double) As String
    ' THE LADDERS' STATE IS THE FLOOR. A profile cannot be current for a run
    ' whose distributions are not, and it cannot exist for a run that produced
    ' nothing. Every other outcome is inherited unchanged, which is what keeps
    ' the two answers from being collapsed into one.
    If StrComp(distribution, SIM_ANNUAL_STATE_CURRENT, vbBinaryCompare) <> 0 Then
        ProfileStateOf = distribution
        Exit Function
    End If
    ' AND ONLY NOW DOES THE SELECTOR SPEAK - for the profile alone. The ladders
    ' above have already been called current and the selector never entered them.
    If Not stampConsistent Then
        ProfileStateOf = SIM_ANNUAL_STATE_OTHER_PX
        Exit Function
    End If
    If Not selectorResolved Then
        ProfileStateOf = SIM_ANNUAL_STATE_OTHER_PX
        Exit Function
    End If
    ' BOTH STAMPED FIELDS MUST AGREE with the authoritative resolution. Half a
    ' match is not a match.
    If StrComp(stampedLabel, resolvedLabel, vbBinaryCompare) = 0 And _
       stampedProbability = resolvedProbability Then
        ProfileStateOf = SIM_ANNUAL_STATE_CURRENT
        Exit Function
    End If
    ' NEVER RELABELLED. The profile stays the profile for the Px it was computed
    ' at; what changed is only that nobody is asking for that Px right now.
    ProfileStateOf = SIM_ANNUAL_STATE_OTHER_PX
End Function

Private Function StampBelongsTo(ByVal bank As String, _
                                ByRef run As SimAnnualIdentity) As Boolean
    Dim runId As Double, seed As Double, iterations As Double

    If Not modWorkbook.IsWholeInRange(StampCell(bank, SIM_ANNUAL_STAMP_ROW_RUN_ID).Value2, _
            CDbl(SIM_RUN_ID_FIRST), CDbl(SIM_RUN_ID_MAXIMUM), runId) Then Exit Function
    If Not modWorkbook.IsWholeInRange( _
            StampCell(bank, SIM_ANNUAL_STAMP_ROW_EFFECTIVE_SEED).Value2, _
            CDbl(SIM_SEED_MIN), CDbl(SIM_SEED_MAX), seed) Then Exit Function
    If Not modWorkbook.IsWholeInRange( _
            StampCell(bank, SIM_ANNUAL_STAMP_ROW_ITERATIONS).Value2, _
            CDbl(SIM_MIN_ITERATIONS), CDbl(SIM_MAX_ITERATIONS), iterations) Then Exit Function
    ' ALL FIVE, NOT THE RUN ID ALONE. The run id says which attempt; the
    ' fingerprint says which request and the digest says which answer.
    StampBelongsTo = IdentityMatches( _
        modWorkbook.SafeLong(runId), run.RunId, modWorkbook.SafeLong(seed), _
        run.EffectiveSeed, StampText(bank, SIM_ANNUAL_STAMP_ROW_REQUEST_FINGERPRINT), _
        run.RequestFingerprint, StampText(bank, SIM_ANNUAL_STAMP_ROW_RESULT_DIGEST), _
        run.ResultDigest, modWorkbook.SafeLong(iterations), run.Iterations)
End Function

Private Function IdentityMatches(ByVal stampedRunId As Long, ByVal runId As Long, _
                                 ByVal stampedSeed As Long, ByVal seed As Long, _
                                 ByVal stampedFingerprint As String, _
                                 ByVal fingerprint As String, _
                                 ByVal stampedDigest As String, ByVal digest As String, _
                                 ByVal stampedIterations As Long, _
                                 ByVal iterations As Long) As Boolean
    ' A BLANK IS NOT A MATCH. Two empty fingerprints are equal strings and say
    ' nothing about whether the block belongs to this run.
    If Len(fingerprint) = 0 Or Len(digest) = 0 Then Exit Function
    If stampedRunId <> runId Then Exit Function
    If stampedSeed <> seed Then Exit Function
    If StrComp(stampedFingerprint, fingerprint, vbBinaryCompare) <> 0 Then Exit Function
    If StrComp(stampedDigest, digest, vbBinaryCompare) <> 0 Then Exit Function
    If stampedIterations <> iterations Then Exit Function
    IdentityMatches = True
End Function

' ==========================================================================
' ADDRESSING - every letter from the projected contract, none typed here
' ==========================================================================
Private Function AnnualSheet() As Worksheet
    Set AnnualSheet = modWorkbook.Sh(SIM_DATA_SHEET)
End Function

Private Function SharedText(ByVal row As Long) As String
    SharedText = modWorkbook.TextOf(AnnualSheet().Range(SIM_SHARED_VALUE_COLUMN & CStr(row)))
End Function

Private Function IsBank(ByVal bank As String) As Boolean
    IsBank = (StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Or _
              StrComp(bank, SIM_BANK_B, vbBinaryCompare) = 0)
End Function

Private Function IsPublished(ByVal bank As String) As Boolean
    IsPublished = (StrComp(StampText(bank, SIM_ANNUAL_STAMP_ROW_PUBLISHED), _
                           SIM_ANNUAL_PUBLISHED, vbBinaryCompare) = 0)
End Function

Private Function StampCell(ByVal bank As String, ByVal row As Long) As Range
    Dim column As String
    If StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Then
        column = SIM_ANNUAL_STAMP_COLUMN_A
    Else
        column = SIM_ANNUAL_STAMP_COLUMN_B
    End If
    Set StampCell = AnnualSheet().Range(column & CStr(row))
End Function

Private Function StampText(ByVal bank As String, ByVal row As Long) As String
    StampText = modWorkbook.TextOf(StampCell(bank, row))
End Function

Private Function StampNumber(ByVal bank As String, ByVal row As Long) As Double
    Dim value As Double
    If modWorkbook.TryReadDouble(StampCell(bank, row).Value2, value) Then
        StampNumber = value
    End If
End Function

Private Function FirstColumn(ByVal bank As String) As String
    If StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Then
        FirstColumn = SIM_ANNUAL_A_PROJECT_INDEX_COLUMN
    Else
        FirstColumn = SIM_ANNUAL_B_PROJECT_INDEX_COLUMN
    End If
End Function

Private Function CalendarYearColumn(ByVal bank As String) As String
    If StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Then
        CalendarYearColumn = SIM_ANNUAL_A_CALENDAR_YEAR_COLUMN
    Else
        CalendarYearColumn = SIM_ANNUAL_B_CALENDAR_YEAR_COLUMN
    End If
End Function

Private Function NominalFirstColumn(ByVal bank As String) As String
    If StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Then
        NominalFirstColumn = SIM_ANNUAL_A_NOMINAL_FIRST_COLUMN
    Else
        NominalFirstColumn = SIM_ANNUAL_B_NOMINAL_FIRST_COLUMN
    End If
End Function

Private Function PvFirstColumn(ByVal bank As String) As String
    If StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Then
        PvFirstColumn = SIM_ANNUAL_A_PV_FIRST_COLUMN
    Else
        PvFirstColumn = SIM_ANNUAL_B_PV_FIRST_COLUMN
    End If
End Function

Private Function NominalProfileColumn(ByVal bank As String) As String
    If StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Then
        NominalProfileColumn = SIM_ANNUAL_A_NOMINAL_PROFILE_COLUMN
    Else
        NominalProfileColumn = SIM_ANNUAL_B_NOMINAL_PROFILE_COLUMN
    End If
End Function

Private Function PvProfileColumn(ByVal bank As String) As String
    If StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Then
        PvProfileColumn = SIM_ANNUAL_A_PV_PROFILE_COLUMN
    Else
        PvProfileColumn = SIM_ANNUAL_B_PV_PROFILE_COLUMN
    End If
End Function

' THE BLOCK'S LAST COLUMN IS THE PV PROFILE'S, because the PV profile is the
' last contracted field. Derived, not written out, so a contract move cannot
' leave the clear range short of the block it is meant to clear.
Private Function LastColumn(ByVal bank As String) As String
    LastColumn = PvProfileColumn(bank)
End Function

' The distance from the block's first column, in columns. Excel resolves the
' letters, so this module does no letter arithmetic of its own and cannot
' disagree with the sheet about where a column is.
Private Function OffsetOf(ByVal bank As String, ByVal column As String) As Long
    OffsetOf = AnnualSheet().Range(column & CStr(SIM_ANNUAL_HEADER_ROW)).Column - _
               AnnualSheet().Range(FirstColumn(bank) & CStr(SIM_ANNUAL_HEADER_ROW)).Column
End Function

Private Function TotalColumn(ByVal bank As String, ByVal measure As String, _
                             ByRef column As String, ByRef detail As String) As Boolean
    If StrComp(measure, SIM_MEASURE_NOMINAL, vbBinaryCompare) = 0 Then
        If StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Then
            column = SIM_ITER_A_TOTAL_NOMINAL_COLUMN
        Else
            column = SIM_ITER_B_TOTAL_NOMINAL_COLUMN
        End If
    ElseIf StrComp(measure, SIM_MEASURE_PV, vbBinaryCompare) = 0 Then
        If StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Then
            column = SIM_ITER_A_TOTAL_PV_COLUMN
        Else
            column = SIM_ITER_B_TOTAL_PV_COLUMN
        End If
    Else
        detail = "annual: the published totals were asked for in an unknown measure"
        Exit Function
    End If
    TotalColumn = True
End Function

' ONE SNAPSHOT FIELD, READ THROUGH THE DOMAIN THE CALLER NAMES. The bounds are
' parameters because the fields this reads have different domains, and a reader
' carrying one ceiling of its own would apply it to all of them.
Private Function SnapshotLong(ByVal bank As String, ByVal row As Long, _
                              ByVal minimum As Double, ByVal maximum As Double, _
                              ByVal what As String, ByRef value As Long, _
                              ByRef detail As String) As Boolean
    Dim raw As Variant, measured As Double
    Dim column As String

    If StrComp(bank, SIM_BANK_A, vbBinaryCompare) = 0 Then
        column = SIM_SNAPSHOT_COLUMN_A
    Else
        column = SIM_SNAPSHOT_COLUMN_B
    End If
    raw = AnnualSheet().Range(column & CStr(row)).Value2
    If Not modWorkbook.IsWholeInRange(raw, minimum, maximum, measured) Then
        detail = "annual: the published " & what & " is not a whole number in " & _
                 CStr(minimum) & " to " & CStr(maximum)
        Exit Function
    End If
    value = modWorkbook.SafeLong(measured)
    SnapshotLong = True
End Function
