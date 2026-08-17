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

' Values are held as Double, not Long. An entered cell is user-controlled and can be
' changed by paste even where Data Validation exists, so a value such as 1E10 must be
' REJECTED, not converted. Nothing here calls CLng on an entered value until the
' relevant bound has been proved.
'
' HasX means "present and a whole number". InRangeX means "additionally inside the
' bound that governs it", and only that flag licenses arithmetic on the value.
Private Type TimelineTriple
    HasBase       As Boolean
    HasStart      As Boolean
    HasDuration   As Boolean
    BaseInRange   As Boolean
    StartInRange  As Boolean
    DurationInRange As Boolean
    BaseYear      As Double
    StartYear     As Double
    Duration      As Double
End Type

' ---------------------------------------------------------------------------
' Reading the two timelines
' ---------------------------------------------------------------------------
Private Function ReadTriple(ByVal BaseName As String, ByVal StartName As String, _
                            ByVal DurationName As String) As TimelineTriple
    Dim t As TimelineTriple
    Dim v As Variant
    Dim d As Double

    v = modWorkbook.ReadValue(BaseName)
    If modWorkbook.TryReadDouble(v, d) Then
        If d = Int(d) Then
            t.HasBase = True
            t.BaseYear = d
            t.BaseInRange = (d >= LIMIT_MIN_YEAR And d <= LIMIT_MAX_YEAR)
        End If
    End If

    v = modWorkbook.ReadValue(StartName)
    If modWorkbook.TryReadDouble(v, d) Then
        If d = Int(d) Then
            t.HasStart = True
            t.StartYear = d
            t.StartInRange = (d >= LIMIT_MIN_YEAR And d <= LIMIT_MAX_YEAR)
        End If
    End If

    v = modWorkbook.ReadValue(DurationName)
    If modWorkbook.TryReadDouble(v, d) Then
        If d = Int(d) Then
            t.HasDuration = True
            t.Duration = d
            t.DurationInRange = (d >= 1 And d <= LIMIT_MAX_YEAR_COLUMNS)
        End If
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

' True only when every value is present, whole AND inside the bound that governs it,
' so arithmetic on the triple cannot overflow.
Private Function IsUsable(ByRef T As TimelineTriple) As Boolean
    IsUsable = IsComplete(T) And T.BaseInRange And T.StartInRange And T.DurationInRange
End Function

' Safe only when IsUsable holds: 2200 + 200 - 1 is far inside a Long.
Private Function LastProjectYear(ByRef T As TimelineTriple) As Long
    LastProjectYear = modWorkbook.SafeLong(T.StartYear + T.Duration - 1)
End Function

Private Function InflationFirstYear(ByRef T As TimelineTriple) As Long
    InflationFirstYear = modWorkbook.SafeLong(T.BaseYear) + 1
End Function

' Legitimately zero when base year = start year and duration = 1.
Private Function InflationYearCount(ByRef T As TimelineTriple) As Long
    If Not IsUsable(T) Then Exit Function
    Dim firstYear As Long, lastYear As Long
    firstYear = InflationFirstYear(T)
    lastYear = LastProjectYear(T)
    If firstYear > lastYear Then Exit Function
    InflationYearCount = lastYear - firstYear + 1
End Function

Private Function DescribeTriple(ByRef T As TimelineTriple) As String
    If Not IsUsable(T) Then
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
' model. The structural width guard is LIMIT_MAX_YEAR_COLUMNS, the Architecture Lock
' Revision B protection on generated PROJECT-YEAR columns ("Generated column count
' > 200 = ERROR"). It is independent of the calendar-year window, which separately
' bounds Base Year, Start Year, Last Project Year and therefore the inflation span.
'
' ORDER IS LOAD-BEARING. Each value is bounded BEFORE any arithmetic that depends on
' it. Duration is checked against the column guard before Last Project Year is
' calculated from it, so an entered 1E10 produces a controlled message rather than a
' VBA Overflow inside the check that was supposed to reject it.
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

    ' --- duration, bounded before it is used in any calculation --------------
    If t.Duration < 1 Then
        problems = problems & "  - Project Duration must be at least 1 year, not " & _
                   Format$(t.Duration, "0") & "." & vbCrLf
    ElseIf Not t.DurationInRange Then
        problems = problems & "  - Project Duration " & Format$(t.Duration, "0") & _
                   " would generate " & Format$(t.Duration, "0") & _
                   " project-year columns, beyond the structural protection limit of " & _
                   LIMIT_MAX_YEAR_COLUMNS & "." & vbCrLf
    End If

    ' --- years, bounded before any comparison or sum uses them ---------------
    If Not t.BaseInRange Then
        problems = problems & "  - Base Year " & Format$(t.BaseYear, "0") & _
                   " is outside the supported range " & LIMIT_MIN_YEAR & "-" & _
                   LIMIT_MAX_YEAR & "." & vbCrLf
    End If
    If Not t.StartInRange Then
        problems = problems & "  - Project Start Year " & Format$(t.StartYear, "0") & _
                   " is outside the supported range " & LIMIT_MIN_YEAR & "-" & _
                   LIMIT_MAX_YEAR & "." & vbCrLf
    End If

    If t.BaseInRange And t.StartInRange Then
        If t.BaseYear > t.StartYear Then
            problems = problems & "  - Base Year " & Format$(t.BaseYear, "0") & _
                       " is later than Project Start Year " & Format$(t.StartYear, "0") & _
                       ". Costs cannot be priced after the project begins." & vbCrLf
        End If
    End If

    ' --- only now is the derived last project year safe to compute -----------
    If t.DurationInRange And t.StartInRange Then
        If LastProjectYear(t) > LIMIT_MAX_YEAR Then
            problems = problems & "  - Last Project Year would be " & LastProjectYear(t) & _
                       ", beyond the supported structural year boundary " & _
                       LIMIT_MAX_YEAR & "." & vbCrLf
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
    Dim costIds() As String, riskIds() As String
    Dim profiles() As String, goneProfiles() As String
    Dim costCount As Long, riskCount As Long
    Dim profileCount As Long, goneCount As Long
    Dim newInflFirst As Long, newInflCount As Long
    Dim removedYearsText As String
    Dim i As Long

    newInflCount = InflationYearCount(newT)
    If newInflCount > 0 Then newInflFirst = InflationFirstYear(newT)

    ' Profiling: only a duration reduction removes cells, and only from the tail.
    If IsUsable(oldT) Then
        If newT.Duration < oldT.Duration Then
            For i = modWorkbook.SafeLong(newT.Duration) + 1 To modWorkbook.SafeLong(oldT.Duration)
                If Len(removedYearsText) > 0 Then removedYearsText = removedYearsText & ", "
                removedYearsText = removedYearsText & "project year " & i & _
                                   " (" & (modWorkbook.SafeLong(oldT.StartYear) + i - 1) & ")"
            Next i
        End If
    End If

    removedProfileCells = _
        modProfiling.CountDataBeyond(modProfiling.CostKind(), modWorkbook.SafeLong(newT.Duration), costIds, costCount) + _
        modProfiling.CountDataBeyond(modProfiling.RiskKind(), modWorkbook.SafeLong(newT.Duration), riskIds, riskCount)

    ' Inflation: assessed for BOTH loss mechanisms in one pass, so a rate lost
    ' because its profile left the Config master is reported even when the timeline
    ' itself is unchanged, and a cell whose profile and year both disappear is
    ' counted once.
    removedRates = modInflation.CountRateLosses(newInflFirst, newInflCount, _
                                                profiles, profileCount, _
                                                goneProfiles, goneCount)

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

    ' A removed profile is reported even when it held no rates: the row is going,
    ' and the user should be told so before it goes.
    If goneCount > 0 Then
        summary = summary & vbCrLf & "  Inflation profiles removed from Config : " & _
                  modWorkbook.JoinLimited(goneProfiles, goneCount, 5, ", ") & vbCrLf
    End If

    If IsDestructive Then
        summary = summary & vbCrLf & "  DATA THAT WILL BE PERMANENTLY DELETED" & vbCrLf
        If removedProfileCells > 0 Then
            summary = summary & "    profiling cells holding data   : " & removedProfileCells & vbCrLf
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
            summary = summary & "    inflation rates lost           : " & removedRates & vbCrLf
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
    Dim costBefore As TableSnapshot, riskBefore As TableSnapshot
    Dim inflationBefore As TableSnapshot
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
    ' Prevalidation has already proved each value is whole and inside its bound, so
    ' these conversions cannot overflow.
    modWorkbook.WriteValue NM_APPLIED_BASE_YEAR, modWorkbook.SafeLong(newT.BaseYear)
    modWorkbook.WriteValue NM_APPLIED_START_YEAR, modWorkbook.SafeLong(newT.StartYear)
    modWorkbook.WriteValue NM_APPLIED_DURATION, modWorkbook.SafeLong(newT.Duration)

    modAppState.FailPointCheck "apply.after_triple"

    modProfiling.SetYearColumns modProfiling.CostKind(), modWorkbook.SafeLong(newT.StartYear), _
                                modWorkbook.SafeLong(newT.Duration)
    modProfiling.SetYearColumns modProfiling.RiskKind(), modWorkbook.SafeLong(newT.StartYear), _
                                modWorkbook.SafeLong(newT.Duration)

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
    Dim reason As String, restoreNote As String
    reason = "Error " & Err.Number & ": " & Err.Description

    If captured Then
        restoreNote = TryRestoreTimeline(appliedBase, appliedStart, appliedDuration, _
                                         costBefore, riskBefore, inflationBefore)
    Else
        restoreNote = "Nothing had been modified when the failure occurred."
    End If

    modAppState.RecalculateStructuralState
    modAppState.RestoreAppState snapshot
    modAppState.RecordResult "FAIL|" & reason & "|" & restoreNote
    If Not modAppState.gAutomationActive Then
        modAppState.ReportFailure "Apply / Update Timeline", reason, restoreNote
    End If
End Sub

' Restores exactly the blocks Apply may modify, and REPORTS whether it succeeded.
' Deliberately not On Error Resume Next: a restore that could not complete is the
' most important thing the user could be told, and burying it would leave them
' believing the workbook was clean when it was not.
Private Function TryRestoreTimeline(ByVal AppliedBase As Variant, _
                                    ByVal AppliedStart As Variant, _
                                    ByVal AppliedDuration As Variant, _
                                    ByRef CostBefore As TableSnapshot, _
                                    ByRef RiskBefore As TableSnapshot, _
                                    ByRef InflationBefore As TableSnapshot) As String
    On Error GoTo RestoreFailed

    modWorkbook.WriteValue NM_APPLIED_BASE_YEAR, AppliedBase
    modWorkbook.WriteValue NM_APPLIED_START_YEAR, AppliedStart
    modWorkbook.WriteValue NM_APPLIED_DURATION, AppliedDuration
    modWorkbook.RestoreTable modProfiling.ProfilingTable(modProfiling.CostKind()), CostBefore
    modWorkbook.RestoreTable modProfiling.ProfilingTable(modProfiling.RiskKind()), RiskBefore
    modWorkbook.RestoreTable modWorkbook.Lo(SH_INFLATION, TBL_INFLATION), InflationBefore

    TryRestoreTimeline = "The applied timeline, both profiling grids and the inflation " & _
                         "grid have been restored to their state from before this " & _
                         "operation, including row count, column count, number formats " & _
                         "and column widths. No partial change remains."
    Exit Function

RestoreFailed:
    TryRestoreTimeline = "RESTORE INCOMPLETE. Error " & Err.Number & ": " & Err.Description & _
                         vbCrLf & "The workbook may hold a partial structural change. " & _
                         "Do not continue; close without saving and reopen the last saved copy."
End Function
