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
