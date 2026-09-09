Attribute VB_Name = "modResultsState"
Option Explicit

' ==========================================================================
' PCCM Phase 8 - THE RESULTS STATE ADAPTER, AND NOTHING ELSE.
'
' Four functions. Each one calls a Phase-7 handoff accessor and returns what it
' says. There is no state rule in this module: no marker comparison, no stamp
' identity test, no selector arm, no CURRENT, no HISTORICAL, no OTHER Px and no
' NOT PRODUCED. If you are looking for the rule, it is in modSimAnnualStore,
' where it has always been.
'
' --------------------------------------------------------------------------
' WHY IT EXISTS AT ALL
' --------------------------------------------------------------------------
' The first Phase-8 Results sheet rebuilt the annual state in worksheet
' formulas. It could not be right, and it was wrong in two specific ways: it
' read the PERSISTED simulation status, so after an ordinary model change the
' banner kept saying CURRENT until some later operation happened to re-evaluate
' it; and it implemented the selector arm of the profile rule while the accepted
' accessor also weighs whether the stamp agrees with itself. A presentation
' layer that owns half a semantic owns a semantic that will diverge.
'
' So the sheet asks the owner. What stopped it asking directly is that a
' worksheet function is only re-run when Excel thinks its inputs moved, and
' these accessors take no arguments: a cell holding one would answer once and
' then keep answering with the same word forever, which is the first defect
' again with better provenance.
'
' --------------------------------------------------------------------------
' VOLATILE, DELIBERATELY, AND NOT ONE FUNCTION MORE
' --------------------------------------------------------------------------
' `Application.Volatile` tells Excel to re-run the function on every calculation
' cycle, whatever moved. That is exactly the dependency these four have: the
' annual state depends on the model, the request, the selector and the stored
' result, which is very nearly the whole workbook. Enumerating that as a range
' argument would be a second, partial copy of the request fingerprint's scope -
' the same mistake in a different place - and it would go stale the moment a
' driver row was added.
'
' THE COST IS BOUNDED AND MEASURED IN THESE FOUR CELLS. Each call reads a
' handful of `_SimData` cells and returns a string or a count. Nothing here
' simulates, replays, samples, sorts or reconciles, and NOTHING HEAVY IS MADE
' VOLATILE: the calculation and the simulation stay exactly as un-volatile as
' they were.
'
' --------------------------------------------------------------------------
' THE ONE RUNTIME UNKNOWN, STATED WHERE IT LIVES
' --------------------------------------------------------------------------
' Two of the four accessors reach `modSimReport.PCCM_SimulationStatus`, which
' re-derives the status AND PERSISTS the two derived rows. Excel does not let a
' function called from a cell change the workbook. The ordinary behaviour is
' that such a write is ignored and the function returns normally - which is all
' this needs, because those two rows are the operations' to maintain and not the
' display's.
'
' NO LINUX TEST CAN SETTLE THAT, so it is not assumed silently. Every adapter
' fails LOUD rather than wrong: if the accessor raises for any reason, the cell
' shows an error instead of a state word. A wrong word would be indistinguishable
' from a right one; an error is not. It is the first thing the Phase-8 Windows
' acceptance has to look at.
' ==========================================================================

Public Function PCCM_ResultsAnnualDistributionState() As Variant
    On Error GoTo Unavailable
    Application.Volatile True
    PCCM_ResultsAnnualDistributionState = _
        modSimAnnualStore.PCCM_AnnualDistributionState()
    Exit Function
Unavailable:
    PCCM_ResultsAnnualDistributionState = CVErr(xlErrValue)
End Function

Public Function PCCM_ResultsAnnualProfileState() As Variant
    On Error GoTo Unavailable
    Application.Volatile True
    PCCM_ResultsAnnualProfileState = modSimAnnualStore.PCCM_AnnualProfileState()
    Exit Function
Unavailable:
    PCCM_ResultsAnnualProfileState = CVErr(xlErrValue)
End Function

' THE Px THE PERSISTED PROFILE IS THE PROFILE FOR - the stamp's own label,
' whatever the selector says now. Volatile for the same reason as the two above:
' the stamp moves when an annual run commits, and no cell reference would tell
' Excel that.
Public Function PCCM_ResultsAnnualProfilePx() As Variant
    On Error GoTo Unavailable
    Application.Volatile True
    PCCM_ResultsAnnualProfilePx = modSimAnnualStore.PCCM_AnnualProfilePx()
    Exit Function
Unavailable:
    PCCM_ResultsAnnualProfilePx = CVErr(xlErrValue)
End Function

Public Function PCCM_ResultsAnnualYearCount() As Variant
    On Error GoTo Unavailable
    Application.Volatile True
    PCCM_ResultsAnnualYearCount = modSimAnnualStore.PCCM_AnnualYearCount()
    Exit Function
Unavailable:
    PCCM_ResultsAnnualYearCount = CVErr(xlErrValue)
End Function

' ==========================================================================
' THE FIFTH ADAPTER, AND IT IS NOT AN ANNUAL ONE
' ==========================================================================
' P8-3 DISCOVERED THE NEED FROM ITS OWN CHART LAYER. The histogram plots the
' published SIMULATION distribution, and the only live state Results exposed
' was the four annual lines above. Two things were wrong with borrowing one:
'
'   THE ANNUAL STATE ANSWERS ABOUT THE ANNUAL PRODUCT. It reads NOT PRODUCED
'   whenever the annual step has not run - which says nothing about a histogram
'   whose data is present and current.
'
'   AND THE PERSISTED ROW IS NOT AN ANSWER AT ALL. `_SimData` D28 is the
'   LAST EVALUATED status; P8-1 proved live that it can still read CURRENT
'   after a request change. A chart qualified by it would look current for a
'   run the model has moved past.
'
' SO THE SEMANTIC OWNER IS ASKED DIRECTLY, through the pure evaluator the P8-1
' read-only correction introduced. `SimReportDerivedStatus` returns exactly what
' `DeriveSimStatus` says and writes nothing; `PCCM_SimulationStatus` - which
' derives the same answer AND persists it - is not reachable from here, because
' a worksheet cell may not change the workbook.
'
' THE WORD IS PASSED THROUGH UNTRANSLATED. STALE stays STALE and INVALID stays
' INVALID: they are different facts about a model, and a presentation layer that
' folded both into HISTORICAL would be inventing a vocabulary to make a chart
' caption tidier. Phase 8 owns no state word.
Public Function PCCM_ResultsSimulationState() As Variant
    On Error GoTo Unavailable
    Application.Volatile True
    PCCM_ResultsSimulationState = modSimReport.SimReportDerivedStatus()
    Exit Function
Unavailable:
    PCCM_ResultsSimulationState = CVErr(xlErrValue)
End Function

' ==========================================================================
' THE SIXTH ADAPTER - THE LIVE CALCULATION STATE, FOR MODEL CHECK
' ==========================================================================
' PHASE 9 NEEDS A LIVE CALCULATION STATE and there was none a cell could ask
' for. The two reasons are the two this module already exists for:
'
'   THE PERSISTED ROW IS NOT AN ANSWER. _Calc C19 is the LAST EVALUATED status.
'   After an ordinary input change it goes on reading CURRENT until something
'   happens to re-evaluate it, so a Model Check built on it would report a
'   model as current after the model had moved - which is the whole failure the
'   sheet exists to catch.
'
'   AND THE OWNER'S OWN ENTRY POINT WRITES. PCCM_CalculationStatus derives the
'   status and then persists it; a worksheet cell may not change the workbook.
'
' SO THE SEMANTIC OWNER IS ASKED THROUGH ITS PURE HALF, exactly as the
' simulation adapter above asks modSimReport.SimReportDerivedStatus. No rule
' moves, no word is translated, and NOT CALCULATED / CURRENT / STALE / INVALID
' arrive on the sheet spelled the way modCalcReport spells them.
'
' VOLATILE FOR THE REASON THE OTHERS ARE. It takes no arguments, and its true
' dependency is every input the calculation fingerprint covers. Enumerating that
' as a range argument would be a second, partial copy of the fingerprint's scope.
'
' AND IT FAILS LOUD. A wrong state word is indistinguishable from a right one;
' an error is not.
Public Function PCCM_ModelCheckCalculationState() As Variant
    Dim detail As String, subject As String
    On Error GoTo Unavailable
    Application.Volatile True
    ' THE DETAIL IS TAKEN AND DROPPED HERE ON PURPOSE. It is required because a
    ' typed Optional with no default does not compile and both callers want it;
    ' the row that reports WHY is the adapter below, not this one.
    PCCM_ModelCheckCalculationState = modCalcReport.CalcReportDerivedStatus(detail, subject)
    Exit Function
Unavailable:
    PCCM_ModelCheckCalculationState = CVErr(xlErrValue)
End Function

' ==========================================================================
' THE SEVENTH ADAPTER - THE LIVE REFUSAL DETAIL, FOR MODEL CHECK
' ==========================================================================
' MODEL CHECK HAD NO LIVE REASON, ONLY A PERSISTED ONE. When the current inputs
' cannot form a calculation, the sheet could say THAT it was invalid but not
' WHY: the only text naming the fault was PCCM_CalculationAttemptDetail, which
' is the LAST ATTEMPT - history, blank until somebody has pressed Calculate, and
' stale the moment the model moves past it. An actionable error whose reason
' came from there would be describing a workbook that no longer exists.
'
' SO THE OWNER IS ASKED FOR THE TEXT IT ALREADY PRODUCED. modCalcReport's
' preparation builds one refusal sentence on the way to deciding the status;
' CalcReportDerivedStatus now hands that sentence back instead of discarding it.
' NOTHING IS RE-DERIVED AND NOTHING IS PARSED: the detail is returned exactly as
' the owner wrote it, and the status it comes with is not this function's to
' report - the sixth adapter above already reports it.
'
' AND IT IS EMPTY WHEN THERE IS NOTHING TO SAY. The preparation clears the
' detail before it starts and assigns it only on a failing branch, so a valid
' current model yields the empty string rather than a stale sentence.
'
' VOLATILE AND LOUD, for the reasons every adapter above is.
Public Function PCCM_ModelCheckRefusalDetail() As Variant
    Dim detail As String, subject As String, ignored As String
    On Error GoTo Unavailable
    Application.Volatile True
    ignored = modCalcReport.CalcReportDerivedStatus(detail, subject)
    PCCM_ModelCheckRefusalDetail = detail
    Exit Function
Unavailable:
    PCCM_ModelCheckRefusalDetail = CVErr(xlErrValue)
End Function

' ==========================================================================
' THE EIGHTH ADAPTER - THE LIVE REFUSAL SUBJECT, FOR MODEL CHECK
' ==========================================================================
' THE SAME PREPARATION, THE OTHER OUTPUT. P9-2A gave Model Check the sentence a
' refusal is written in; this gives it the PERMANENT ID that refusal is about,
' as a value rather than as words inside one. The two come out of one call to
' one preparation: there is no second traversal, no second validation and no
' lookup of any kind here.
'
' NOTHING IS PARSED. The id is the one modCalcCheck and modCalcResolve already
' hold when they refuse - the same id their sentences are built from - handed
' out through a ByRef rather than recovered from prose afterwards.
'
' AND IT IS BLANK WHEN NO DRIVER IS AT FAULT. A missing register, an unusable
' discount rate, an inflation grid that is not there: those are model-wide, no
' permanent id exists for them, and this returns the empty string rather than
' the last driver anybody happened to look at. The owners clear it on success
' for exactly that reason.
'
' NO CACHE, NO STATIC, NO MODULE STATE. Every call re-asks the owner; nothing
' here remembers the previous answer.
Public Function PCCM_ModelCheckRefusalSubject() As Variant
    Dim detail As String, subject As String, ignored As String
    On Error GoTo Unavailable
    Application.Volatile True
    ignored = modCalcReport.CalcReportDerivedStatus(detail, subject)
    PCCM_ModelCheckRefusalSubject = subject
    Exit Function
Unavailable:
    PCCM_ModelCheckRefusalSubject = CVErr(xlErrValue)
End Function
