Attribute VB_Name = "modTimeline"
Option Explicit

' ===========================================================================
' PCCM - Apply / Update Timeline
' ===========================================================================
' Phase 4. The single structural synchronisation pathway. There is deliberately
' no second structural state machine: profiling row sync, inflation profile sync
' and year-column generation all run from here, because splitting them would let
' the workbook sit in a state where one grid had been synchronised and another
' had not.
'
' The operation is ordered so that it can never expose a half-applied state:
'
'   1. PREVALIDATE the entered triple. Nothing is touched until it passes, and a
'      rejection leaves the applied triple exactly as it was.
'   2. ASSESS what the combined old -> new delta would destroy. One old
'      configuration, one new configuration, one delta - never three sequential
'      partial transitions.
'   3. CONFIRM before modifying. A cancellation therefore needs no rollback,
'      because nothing has moved.
'   4. SNAPSHOT the blocks this operation may modify.
'   5. APPLY, REVALIDATE, and on any failure RESTORE those blocks logically.
' ===========================================================================

Private Type TimelineTriple
    HasBase     As Boolean
    HasStart    As Boolean
    HasDuration As Boolean
    BaseYear    As Long
    StartYear   As Long
    Duration    As Long
End Type

' ---------------------------------------------------------------------------
' Reading the two timelines
' ---------------------------------------------------------------------------
Private Function ReadTriple(ByVal BaseName As String, ByVal StartName As String, _
                            ByVal DurationName As String) As TimelineTriple
    Dim t As TimelineTriple
    Dim v As Variant

    v = modWorkbook.ReadValue(BaseName)
    If modWorkbook.IsWholeNumber(v) Then
        t.HasBase = True
        t.BaseYear = CLng(v)
    End If
    v = modWorkbook.ReadValue(StartName)
    If modWorkbook.IsWholeNumber(v) Then
        t.HasStart = True
        t.StartYear = CLng(v)
    End If
    v = modWorkbook.ReadValue(DurationName)
    If modWorkbook.IsWholeNumber(v) Then
        t.HasDuration = True
        t.Duration = CLng(v)
    End If
    ReadTriple = t
End Function

Private Function Entered() As TimelineTriple
    Entered = ReadTriple(NM_BASE_YEAR_ENTERED, NM_PROJECT_START_YEAR_ENTERED, NM_DURATION_YEARS_ENTERED)
End Function

Private Function Applied() As TimelineTriple
    Applied = ReadTriple(NM_APPLIED_BASE_YEAR, NM_APPLIED_START_YEAR, NM_APPLIED_DURATION)
End Function

Private Function IsComplete(ByRef T As TimelineTriple) As Boolean
    IsComplete = T.HasBase And T.HasStart And T.HasDuration
End Function

Private Function LastProjectYear(ByRef T As TimelineTriple) As Long
    LastProjectYear = T.StartYear + T.Duration - 1
End Function

Private Function InflationFirstYear(ByRef T As TimelineTriple) As Long
    InflationFirstYear = T.BaseYear + 1
End Function

' Legitimately zero when base year = start year and duration = 1.
Private Function InflationYearCount(ByRef T As TimelineTriple) As Long
    If Not IsComplete(T) Then Exit Function
    Dim firstYear As Long, lastYear As Long
    firstYear = InflationFirstYear(T)
    lastYear = LastProjectYear(T)
    If firstYear > lastYear Then Exit Function
    InflationYearCount = lastYear - firstYear + 1
End Function

Private Function DescribeTriple(ByRef T As TimelineTriple) As String
    If Not IsComplete(T) Then
        DescribeTriple = "(not applied)"
    Else
        DescribeTriple = "Base " & T.BaseYear & " / Start " & T.StartYear & _
                         " / Duration " & T.Duration & _
                         " (project years " & T.StartYear & "-" & LastProjectYear(T) & ")"
    End If
End Function

' ---------------------------------------------------------------------------
' 1. Prevalidation
' ---------------------------------------------------------------------------
' Returns every reason the entered triple may not be applied, so the user is not
' sent back one failure at a time. An empty string means it may be applied.
'
' Note what is NOT here: no 25-year cap. Twenty-five years is an Architecture
' benchmark target, not a business maximum, and a 40-year project is a legitimate
' model. The only width guard is the structural generation limit, which is the
' width of the supported calendar-year window.
Public Function PrevalidateEntered() As String
    Dim t As TimelineTriple
    Dim problems As String
    t = Entered()

    If Not t.HasBase Then problems = problems & "  - Base Year must be a whole number." & vbCrLf
    If Not t.HasStart Then problems = problems & "  - Project Start Year must be a whole number." & vbCrLf
    If Not t.HasDuration Then problems = problems & "  - Project Duration (Years) must be a whole number." & vbCrLf
    If Len(problems) > 0 Then
        PrevalidateEntered = problems
        Exit Function
    End If

    If t.Duration < 1 Then
        problems = problems & "  - Project Duration must be at least 1 year, not " & t.Duration & "." & vbCrLf
    End If
    If t.BaseYear < LIMIT_MIN_YEAR Or t.BaseYear > LIMIT_MAX_YEAR Then
        problems = problems & "  - Base Year " & t.BaseYear & " is outside the supported range " & _
                   LIMIT_MIN_YEAR & "-" & LIMIT_MAX_YEAR & "." & vbCrLf
    End If
    If t.StartYear < LIMIT_MIN_YEAR Or t.StartYear > LIMIT_MAX_YEAR Then
        problems = problems & "  - Project Start Year " & t.StartYear & " is outside the supported range " & _
                   LIMIT_MIN_YEAR & "-" & LIMIT_MAX_YEAR & "." & vbCrLf
    End If
    If t.BaseYear > t.StartYear Then
        problems = problems & "  - Base Year " & t.BaseYear & " is later than Project Start Year " & _
                   t.StartYear & ". Costs cannot be priced after the project begins." & vbCrLf
    End If
    If t.Duration >= 1 Then
        If LastProjectYear(t) > LIMIT_MAX_YEAR Then
            problems = problems & "  - Last Project Year would be " & LastProjectYear(t) & _
                       ", beyond the supported structural year boundary " & LIMIT_MAX_YEAR & "." & vbCrLf
        End If
        If t.Duration > LIMIT_MAX_YEAR_COLUMNS Then
            problems = problems & "  - Project Duration " & t.Duration & " would generate " & t.Duration & _
                       " project-year columns, beyond the structural protection limit of " & _
                       LIMIT_MAX_YEAR_COLUMNS & "." & vbCrLf
        End If
    End If

    PrevalidateEntered = problems
End Function

' ---------------------------------------------------------------------------
' 2. Assessment of the combined delta
' ---------------------------------------------------------------------------
Private Function BuildSummary(ByRef oldT As TimelineTriple, ByRef newT As TimelineTriple, _
                              ByRef IsDestructive As Boolean) As String
    Dim summary As String
    Dim removedProfileCells As Long, removedRates As Long
    Dim costIds() As String, riskIds() As String, profiles() As String
    Dim costCount As Long, riskCount As Long, profileCount As Long
    Dim newInflFirst As Long, newInflCount As Long
    Dim removedYearsText As String
    Dim i As Long

    newInflCount = InflationYearCount(newT)
    If newInflCount > 0 Then newInflFirst = InflationFirstYear(newT)

    ' Profiling: only a duration reduction removes cells, and only from the tail.
    If IsComplete(oldT) Then
        If newT.Duration < oldT.Duration Then
            For i = newT.Duration + 1 To oldT.Duration
                If Len(removedYearsText) > 0 Then removedYearsText = removedYearsText & ", "
                removedYearsText = removedYearsText & "project year " & i & _
                                   " (" & (oldT.StartYear + i - 1) & ")"
            Next i
        End If
    End If

    removedProfileCells = _
        modProfiling.CountNonZeroBeyond(modProfiling.CostKind(), newT.Duration, costIds, costCount) + _
        modProfiling.CountNonZeroBeyond(modProfiling.RiskKind(), newT.Duration, riskIds, riskCount)

    ' Inflation: rates leave the span whenever a year leaves it, from either end.
    removedRates = modInflation.CountRatesLeavingSpan(newInflFirst, newInflCount, profiles, profileCount)

    IsDestructive = (removedProfileCells > 0) Or (removedRates > 0)

    summary = "Apply / Update Timeline" & vbCrLf & vbCrLf & _
              "  Current applied : " & DescribeTriple(oldT) & vbCrLf & _
              "  New applied     : " & DescribeTriple(newT) & vbCrLf

    If newInflCount > 0 Then
        summary = summary & "  Inflation years : " & newInflFirst & "-" & _
                  (newInflFirst + newInflCount - 1) & vbCrLf
    Else
        summary = summary & "  Inflation years : none required (" & MSG_INFLATION_EMPTY_SPAN & ")" & vbCrLf
    End If

    If Len(removedYearsText) > 0 Then
        summary = summary & vbCrLf & "  Project years removed : " & removedYearsText & vbCrLf
    End If

    If IsDestructive Then
        summary = summary & vbCrLf & "  DATA THAT WILL BE PERMANENTLY DELETED" & vbCrLf
        If removedProfileCells > 0 Then
            summary = summary & "    non-zero profiling percentages : " & removedProfileCells & vbCrLf
            If costCount > 0 Then
                summary = summary & "    affected cost lines            : " & _
                          modWorkbook.JoinLimited(costIds, costCount, 5, ", ") & vbCrLf
            End If
            If riskCount > 0 Then
                summary = summary & "    affected risks                 : " & _
                          modWorkbook.JoinLimited(riskIds, riskCount, 5, ", ") & vbCrLf
            End If
        End If
        If removedRates > 0 Then
            summary = summary & "    inflation rates leaving span   : " & removedRates & vbCrLf
            If profileCount > 0 Then
                summary = summary & "    affected profiles              : " & _
                          modWorkbook.JoinLimited(profiles, profileCount, 5, ", ") & vbCrLf
            End If
        End If
    End If

    BuildSummary = summary
End Function

' ---------------------------------------------------------------------------
' Entry point
' ---------------------------------------------------------------------------
Public Sub PCCM_ApplyTimeline()
    Dim snapshot As AppStateSnapshot
    Dim oldT As TimelineTriple, newT As TimelineTriple
    Dim problems As String, summary As String
    Dim isDestructive As Boolean
    Dim costBefore As Variant, riskBefore As Variant, inflationBefore As Variant
    Dim appliedBase As Variant, appliedStart As Variant, appliedDuration As Variant
    Dim captured As Boolean
    Dim result As OperationResult

    snapshot = modAppState.CaptureAppState()

    ' --- 1. prevalidate, before anything at all is touched -------------------
    problems = PrevalidateEntered()
    If Len(problems) > 0 Then
        result = modAppState.Failed( _
            "Apply / Update Timeline was rejected. The applied timeline is unchanged.", _
            "The entered timeline is not valid:" & vbCrLf & problems)
        modAppState.Announce result
        Exit Sub
    End If

    oldT = Applied()
    newT = Entered()

    ' --- 2 and 3. assess the combined delta, then confirm --------------------
    summary = BuildSummary(oldT, newT, isDestructive)
    If Not modAppState.AskConfirm(summary, isDestructive) Then
        modAppState.RecordResult "OK|cancelled"
        Exit Sub
    End If

    On Error GoTo Failure
    modAppState.BeginOperation

    ' --- 4. snapshot exactly the blocks this operation may modify ------------
    appliedBase = modWorkbook.ReadValue(NM_APPLIED_BASE_YEAR)
    appliedStart = modWorkbook.ReadValue(NM_APPLIED_START_YEAR)
    appliedDuration = modWorkbook.ReadValue(NM_APPLIED_DURATION)
    costBefore = modWorkbook.SnapshotTable(modProfiling.ProfilingTable(modProfiling.CostKind()))
    riskBefore = modWorkbook.SnapshotTable(modProfiling.ProfilingTable(modProfiling.RiskKind()))
    inflationBefore = modWorkbook.SnapshotTable(modWorkbook.Lo(SH_INFLATION, TBL_INFLATION))
    captured = True

    ' --- 5. apply one coherent transition -----------------------------------
    modWorkbook.WriteValue NM_APPLIED_BASE_YEAR, newT.BaseYear
    modWorkbook.WriteValue NM_APPLIED_START_YEAR, newT.StartYear
    modWorkbook.WriteValue NM_APPLIED_DURATION, newT.Duration

    modAppState.FailPointCheck "apply.after_triple"

    modProfiling.SetYearColumns modProfiling.CostKind(), newT.StartYear, newT.Duration
    modProfiling.SetYearColumns modProfiling.RiskKind(), newT.StartYear, newT.Duration

    modAppState.FailPointCheck "apply.after_profiling_columns"

    modProfiling.SyncRows modProfiling.CostKind()
    modProfiling.SyncRows modProfiling.RiskKind()

    modAppState.FailPointCheck "apply.after_profiling_rows"

    Dim inflCount As Long, inflFirst As Long
    inflCount = InflationYearCount(newT)
    If inflCount > 0 Then inflFirst = InflationFirstYear(newT)
    modInflation.SetYearColumns inflFirst, inflCount
    modInflation.SyncProfileRows

    modAppState.FailPointCheck "apply.after_inflation"

    ' --- structural revalidation gates the operation ------------------------
    problems = modStructuralCheck.ValidateStructure()
    If Len(problems) > 0 Then
        Err.Raise vbObjectError + 5003, "modTimeline.PCCM_ApplyTimeline", _
                  "Structural revalidation failed:" & vbCrLf & problems
    End If

    modAppState.RecalculateStructuralState
    modAppState.RestoreAppState snapshot
    modAppState.Announce modAppState.Succeeded( _
        "Timeline applied: " & DescribeTriple(newT) & ".")
    Exit Sub

Failure:
    Dim reason As String
    reason = "Error " & Err.Number & ": " & Err.Description
    If captured Then
        On Error Resume Next
        modWorkbook.WriteValue NM_APPLIED_BASE_YEAR, appliedBase
        modWorkbook.WriteValue NM_APPLIED_START_YEAR, appliedStart
        modWorkbook.WriteValue NM_APPLIED_DURATION, appliedDuration
        modWorkbook.RestoreTable modProfiling.ProfilingTable(modProfiling.CostKind()), costBefore
        modWorkbook.RestoreTable modProfiling.ProfilingTable(modProfiling.RiskKind()), riskBefore
        modWorkbook.RestoreTable modWorkbook.Lo(SH_INFLATION, TBL_INFLATION), inflationBefore
        On Error GoTo 0
    End If
    modAppState.RecalculateStructuralState
    modAppState.RestoreAppState snapshot
    modAppState.RecordResult "FAIL|" & reason
    If Not modAppState.gAutomationActive Then
        modAppState.ReportFailure "Apply / Update Timeline", reason, _
            "The applied timeline, both profiling grids and the inflation grid have been " & _
            "restored to their values from before this operation. No partial change remains."
    End If
End Sub
