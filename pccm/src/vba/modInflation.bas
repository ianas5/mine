Attribute VB_Name = "modInflation"
Option Explicit

' ===========================================================================
' PCCM - Inflation grid structure
' ===========================================================================
' Phase 4. Structure only. No escalation factor is computed, no rate is
' interpolated and no missing rate is filled in.
'
' The governing invariant differs from profiling and the difference is the whole
' point of this module:
'
'   ANCHOR     An inflation rate is anchored by CALENDAR YEAR. A rate entered
'              against 2029 stays attached to 2029 for as long as 2029 is inside
'              the required span, whatever column index it ends up in. A rate is
'              NEVER shifted because its column position changed.
'
'   OWNERSHIP  A row belongs to a PROFILE NAME from the Config master, not to a
'              worksheet row.
'
'   BLANKS     A newly required year arrives BLANK. Seeding 0% would fabricate an
'              escalation assumption the user never made and would hide a missing
'              rate from the Model Check phase that is meant to block on it.
' ===========================================================================

' ---------------------------------------------------------------------------
' Required span
' ---------------------------------------------------------------------------
' Applied base year + 1 through applied last project year. Legitimately EMPTY
' when base year = start year and duration = 1: there is no year between the
' price base and the end of the project, so no escalation assumption is required.
Public Function RequiredFirstYear() As Variant
    Dim baseYear As Variant
    baseYear = modWorkbook.ReadValue(NM_APPLIED_BASE_YEAR)
    If IsEmpty(baseYear) Then
        RequiredFirstYear = Empty
    Else
        RequiredFirstYear = CLng(baseYear) + 1
    End If
End Function

Public Function RequiredLastYear() As Variant
    Dim startYear As Variant, duration As Variant
    startYear = modWorkbook.ReadValue(NM_APPLIED_START_YEAR)
    duration = modWorkbook.ReadValue(NM_APPLIED_DURATION)
    If IsEmpty(startYear) Or IsEmpty(duration) Then
        RequiredLastYear = Empty
    Else
        RequiredLastYear = CLng(startYear) + CLng(duration) - 1
    End If
End Function

Public Function RequiredYearCount() As Long
    Dim firstYear As Variant, lastYear As Variant
    firstYear = RequiredFirstYear()
    lastYear = RequiredLastYear()
    If IsEmpty(firstYear) Or IsEmpty(lastYear) Then Exit Function
    If CLng(firstYear) > CLng(lastYear) Then Exit Function
    RequiredYearCount = CLng(lastYear) - CLng(firstYear) + 1
End Function

Public Function HasEmptySpan() As Boolean
    Dim firstYear As Variant, lastYear As Variant
    firstYear = RequiredFirstYear()
    lastYear = RequiredLastYear()
    If IsEmpty(firstYear) Or IsEmpty(lastYear) Then Exit Function
    HasEmptySpan = (CLng(firstYear) > CLng(lastYear))
End Function

' ---------------------------------------------------------------------------
' Year columns, remapped by calendar year
' ---------------------------------------------------------------------------
Public Sub SetYearColumns(ByVal FirstYear As Variant, ByVal YearCount As Long)
    Dim target As ListObject
    Dim fixedCols As Long, rowCount As Long
    Dim r As Long, c As Long, y As Long

    Set target = modWorkbook.Lo(SH_INFLATION, TBL_INFLATION)
    fixedCols = GRID_INFLATION_FIXED_COLS
    rowCount = modWorkbook.BodyRowCount(target)

    ' Capture every existing rate against its calendar year BEFORE the shape
    ' changes. This is what makes the remap positional-independent.
    Dim held As Object
    Set held = CreateObject("Scripting.Dictionary")
    Dim existingCols As Long
    existingCols = target.ListColumns.Count - fixedCols
    For r = 1 To rowCount
        Dim profileName As String
        profileName = modWorkbook.TextOf(modWorkbook.CellIn(target, r, 1))
        If Len(profileName) > 0 Then
            For c = 1 To existingCols
                Dim headerText As String
                headerText = target.ListColumns(fixedCols + c).Name
                If IsNumeric(headerText) Then
                    Dim cell As Range
                    Set cell = modWorkbook.CellIn(target, r, fixedCols + c)
                    If Not modWorkbook.IsEmptyCell(cell) Then
                        held(profileName & "|" & CStr(CLng(headerText))) = cell.Value
                    End If
                End If
            Next c
        End If
    Next r

    ' Reshape. A zero-width span is a legitimate outcome, so the table simply keeps
    ' its fixed columns; it is never left malformed and never given a stub column.
    Do While target.ListColumns.Count > fixedCols + YearCount
        target.ListColumns(target.ListColumns.Count).Delete
    Loop
    Do While target.ListColumns.Count < fixedCols + YearCount
        Dim added As ListColumn
        Set added = target.ListColumns.Add
        added.DataBodyRange.NumberFormat = GRID_INFLATION_YEAR_FORMAT
        added.Range.ColumnWidth = GRID_INFLATION_YEAR_WIDTH
    Loop

    ' Relabel, then re-place every surviving rate by its calendar year. Any year
    ' with no captured rate stays blank - including newly required years.
    For y = 1 To YearCount
        target.ListColumns(fixedCols + y).Name = CStr(CLng(FirstYear) + y - 1)
        target.HeaderRowRange.Cells(1, fixedCols + y).NumberFormat = GRID_INFLATION_HEADER_FORMAT
    Next y

    For r = 1 To rowCount
        Dim rowName As String
        rowName = modWorkbook.TextOf(modWorkbook.CellIn(target, r, 1))
        For y = 1 To YearCount
            Dim slot As Range
            Set slot = modWorkbook.CellIn(target, r, fixedCols + y)
            slot.ClearContents
            If Len(rowName) > 0 Then
                Dim lookupKey As String
                lookupKey = rowName & "|" & CStr(CLng(FirstYear) + y - 1)
                If held.Exists(lookupKey) Then slot.Value = held(lookupKey)
            End If
        Next y
    Next r
End Sub

Public Sub ClearYearColumns()
    SetYearColumns 0, 0
End Sub

' ---------------------------------------------------------------------------
' Destructive assessment - runs BEFORE anything is modified
' ---------------------------------------------------------------------------
' Counts the non-blank rates that would leave the required span. A blank leaving
' the span destroys nothing and is deliberately not counted.
Public Function CountRatesLeavingSpan(ByVal NewFirstYear As Variant, _
                                      ByVal NewYearCount As Long, _
                                      ByRef AffectedProfiles() As String, _
                                      ByRef AffectedCount As Long) As Long
    Dim target As ListObject
    Dim fixedCols As Long, rowCount As Long, existingCols As Long
    Dim r As Long, c As Long, hits As Long
    Dim seen As Object

    Set target = modWorkbook.Lo(SH_INFLATION, TBL_INFLATION)
    fixedCols = GRID_INFLATION_FIXED_COLS
    rowCount = modWorkbook.BodyRowCount(target)
    existingCols = target.ListColumns.Count - fixedCols
    Set seen = CreateObject("Scripting.Dictionary")

    ReDim AffectedProfiles(1 To IIf(rowCount < 1, 1, rowCount))
    AffectedCount = 0

    Dim newLast As Long, newFirst As Long
    If NewYearCount > 0 Then
        newFirst = CLng(NewFirstYear)
        newLast = newFirst + NewYearCount - 1
    End If

    For r = 1 To rowCount
        Dim profileName As String
        profileName = modWorkbook.TextOf(modWorkbook.CellIn(target, r, 1))
        If Len(profileName) > 0 Then
            For c = 1 To existingCols
                Dim headerText As String
                headerText = target.ListColumns(fixedCols + c).Name
                If IsNumeric(headerText) Then
                    Dim yearValue As Long
                    yearValue = CLng(headerText)
                    Dim survives As Boolean
                    survives = (NewYearCount > 0) And (yearValue >= newFirst) And (yearValue <= newLast)
                    If Not survives Then
                        If Not modWorkbook.IsEmptyCell(modWorkbook.CellIn(target, r, fixedCols + c)) Then
                            hits = hits + 1
                            If Not seen.Exists(profileName) Then
                                seen.Add profileName, True
                                AffectedCount = AffectedCount + 1
                                AffectedProfiles(AffectedCount) = profileName
                            End If
                        End If
                    End If
                End If
            Next c
        End If
    Next r

    CountRatesLeavingSpan = hits
End Function

' Counts non-blank rates that would be lost because a profile name has been
' removed from the Config master. Handled as destructive for the same reason.
Public Function CountRatesForRemovedProfiles(ByRef AffectedProfiles() As String, _
                                             ByRef AffectedCount As Long) As Long
    Dim target As ListObject
    Dim fixedCols As Long, rowCount As Long, r As Long, c As Long, hits As Long
    Dim wanted As Object

    Set target = modWorkbook.Lo(SH_INFLATION, TBL_INFLATION)
    fixedCols = GRID_INFLATION_FIXED_COLS
    rowCount = modWorkbook.BodyRowCount(target)
    Set wanted = ProfileNameSet()

    ReDim AffectedProfiles(1 To IIf(rowCount < 1, 1, rowCount))
    AffectedCount = 0

    For r = 1 To rowCount
        Dim profileName As String
        profileName = modWorkbook.TextOf(modWorkbook.CellIn(target, r, 1))
        If Len(profileName) > 0 Then
            If Not wanted.Exists(LCase$(profileName)) Then
                For c = fixedCols + 1 To target.ListColumns.Count
                    If Not modWorkbook.IsEmptyCell(modWorkbook.CellIn(target, r, c)) Then
                        hits = hits + 1
                    End If
                Next c
                AffectedCount = AffectedCount + 1
                AffectedProfiles(AffectedCount) = profileName
            End If
        End If
    Next r

    CountRatesForRemovedProfiles = hits
End Function

' ---------------------------------------------------------------------------
' Row synchronisation - keyed by profile name
' ---------------------------------------------------------------------------
Public Function ProfileNameSet() As Object
    Dim master As ListObject
    Dim r As Long, rowCount As Long
    Dim names As Object
    Set names = CreateObject("Scripting.Dictionary")
    Set master = modWorkbook.Lo(SH_CONFIG, TBL_INFLATION_PROFILES)
    rowCount = modWorkbook.BodyRowCount(master)
    For r = 1 To rowCount
        Dim profileName As String
        profileName = modWorkbook.TextOf(modWorkbook.CellIn(master, r, 1))
        If Len(profileName) > 0 Then
            If Not names.Exists(LCase$(profileName)) Then names.Add LCase$(profileName), profileName
        End If
    Next r
    Set ProfileNameSet = names
End Function

' Rebuilds the inflation rows against the Config profile master. Surviving rows
' keep their rates by NAME, not by position; new profiles arrive with blank rates.
Public Sub SyncProfileRows()
    Dim target As ListObject, master As ListObject
    Dim fixedCols As Long, yearCols As Long
    Dim r As Long, c As Long, rowCount As Long

    Set target = modWorkbook.Lo(SH_INFLATION, TBL_INFLATION)
    Set master = modWorkbook.Lo(SH_CONFIG, TBL_INFLATION_PROFILES)
    fixedCols = GRID_INFLATION_FIXED_COLS
    yearCols = target.ListColumns.Count - fixedCols
    rowCount = modWorkbook.BodyRowCount(target)

    Dim held As Object
    Set held = CreateObject("Scripting.Dictionary")
    For r = 1 To rowCount
        Dim rowName As String
        rowName = modWorkbook.TextOf(modWorkbook.CellIn(target, r, 1))
        If Len(rowName) > 0 Then
            For c = 1 To yearCols
                Dim cell As Range
                Set cell = modWorkbook.CellIn(target, r, fixedCols + c)
                If Not modWorkbook.IsEmptyCell(cell) Then
                    held(LCase$(rowName) & "|" & target.ListColumns(fixedCols + c).Name) = cell.Value
                End If
            Next c
        End If
    Next r

    Dim writeRow As Long, masterRows As Long
    masterRows = modWorkbook.BodyRowCount(master)
    writeRow = 0
    For r = 1 To masterRows
        Dim profileName As String
        profileName = modWorkbook.TextOf(modWorkbook.CellIn(master, r, 1))
        If Len(profileName) > 0 Then
            writeRow = writeRow + 1
            If writeRow > modWorkbook.BodyRowCount(target) Then target.ListRows.Add
            modWorkbook.CellIn(target, writeRow, 1).Value = profileName
            For c = 1 To yearCols
                Dim slot As Range
                Set slot = modWorkbook.CellIn(target, writeRow, fixedCols + c)
                Dim lookupKey As String
                lookupKey = LCase$(profileName) & "|" & target.ListColumns(fixedCols + c).Name
                If held.Exists(lookupKey) Then
                    slot.Value = held(lookupKey)
                Else
                    ' A new profile, or a newly required year: BLANK, never zero.
                    slot.ClearContents
                End If
            Next c
        End If
    Next r

    rowCount = modWorkbook.BodyRowCount(target)
    For r = writeRow + 1 To rowCount
        For c = 1 To target.ListColumns.Count
            modWorkbook.CellIn(target, r, c).ClearContents
        Next c
    Next r
End Sub

' ---------------------------------------------------------------------------
' Inspection
' ---------------------------------------------------------------------------
Public Function YearHeaders() As String
    Dim target As ListObject
    Dim i As Long, out As String
    Set target = modWorkbook.Lo(SH_INFLATION, TBL_INFLATION)
    For i = GRID_INFLATION_FIXED_COLS + 1 To target.ListColumns.Count
        If Len(out) > 0 Then out = out & ","
        out = out & target.ListColumns(i).Name
    Next i
    YearHeaders = out
End Function

Public Function ProfileList() As String
    Dim target As ListObject
    Dim r As Long, rowCount As Long, out As String, rowName As String
    Set target = modWorkbook.Lo(SH_INFLATION, TBL_INFLATION)
    rowCount = modWorkbook.BodyRowCount(target)
    For r = 1 To rowCount
        rowName = modWorkbook.TextOf(modWorkbook.CellIn(target, r, 1))
        If Len(rowName) > 0 Then
            If Len(out) > 0 Then out = out & ","
            out = out & rowName
        End If
    Next r
    ProfileList = out
End Function

Public Function RateFor(ByVal ProfileName As String, ByVal CalendarYear As Long) As Variant
    Dim target As ListObject
    Dim fixedCols As Long, r As Long, c As Long, rowCount As Long
    Set target = modWorkbook.Lo(SH_INFLATION, TBL_INFLATION)
    fixedCols = GRID_INFLATION_FIXED_COLS
    rowCount = modWorkbook.BodyRowCount(target)
    RateFor = Empty
    For r = 1 To rowCount
        If StrComp(modWorkbook.TextOf(modWorkbook.CellIn(target, r, 1)), ProfileName, vbTextCompare) = 0 Then
            For c = fixedCols + 1 To target.ListColumns.Count
                If StrComp(target.ListColumns(c).Name, CStr(CalendarYear), vbTextCompare) = 0 Then
                    If Not modWorkbook.IsEmptyCell(modWorkbook.CellIn(target, r, c)) Then
                        RateFor = modWorkbook.CellIn(target, r, c).Value
                    End If
                    Exit Function
                End If
            Next c
            Exit Function
        End If
    Next r
End Function

Public Sub SetRateFor(ByVal ProfileName As String, ByVal CalendarYear As Long, _
                      ByVal NewValue As Variant)
    Dim target As ListObject
    Dim fixedCols As Long, r As Long, c As Long, rowCount As Long
    Set target = modWorkbook.Lo(SH_INFLATION, TBL_INFLATION)
    fixedCols = GRID_INFLATION_FIXED_COLS
    rowCount = modWorkbook.BodyRowCount(target)
    For r = 1 To rowCount
        If StrComp(modWorkbook.TextOf(modWorkbook.CellIn(target, r, 1)), ProfileName, vbTextCompare) = 0 Then
            For c = fixedCols + 1 To target.ListColumns.Count
                If StrComp(target.ListColumns(c).Name, CStr(CalendarYear), vbTextCompare) = 0 Then
                    modWorkbook.CellIn(target, r, c).Value = NewValue
                    Exit Sub
                End If
            Next c
            Exit Sub
        End If
    Next r
End Sub
