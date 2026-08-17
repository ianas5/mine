Attribute VB_Name = "modProfiling"
Option Explicit

' ===========================================================================
' PCCM - Cost and Risk profiling grid structure
' ===========================================================================
' Phase 4. Structure only. This module allocates no percentage, redistributes
' nothing, normalises nothing and computes no cost. The requirement that a row
' totals 100% is a Model Check rule and is deliberately absent here.
'
' Two invariants govern everything below:
'
'   ANCHOR      A profiling percentage is anchored by PROJECT-YEAR INDEX, not by
'               calendar year. Changing the start year relabels the headers and
'               moves no value. Only a duration change adds or removes cells,
'               and only at the tail.
'
'   OWNERSHIP   A profiling row belongs to a PERMANENT ID, not to a worksheet
'               row. Reordering the driver register reorders the grid but
'               transfers no data between rows.
' ===========================================================================

Private Const KIND_COST As String = "cost"
Private Const KIND_RISK As String = "risk"

' ---------------------------------------------------------------------------
' Geometry
' ---------------------------------------------------------------------------
Public Function ProfilingTable(ByVal Kind As String) As ListObject
    If Kind = KIND_COST Then
        Set ProfilingTable = modWorkbook.Lo(SH_COST_PROFILING, TBL_COST_PROFILING)
    Else
        Set ProfilingTable = modWorkbook.Lo(SH_RISK_PROFILING, TBL_RISK_PROFILING)
    End If
End Function

Public Function FixedColumnCount(ByVal Kind As String) As Long
    If Kind = KIND_COST Then
        FixedColumnCount = GRID_COST_PROFILING_FIXED_COLS
    Else
        FixedColumnCount = GRID_RISK_PROFILING_FIXED_COLS
    End If
End Function

Public Function YearColumnCount(ByVal Kind As String) As Long
    YearColumnCount = ProfilingTable(Kind).ListColumns.Count - FixedColumnCount(Kind)
End Function

Public Function CostKind() As String
    CostKind = KIND_COST
End Function

Public Function RiskKind() As String
    RiskKind = KIND_RISK
End Function

' ---------------------------------------------------------------------------
' Year columns
' ---------------------------------------------------------------------------
' Reshapes a profiling grid to exactly NewCount project-year columns labelled from
' StartYear. Existing cells keep their positions; growth appends cells initialised
' to PROFILE_INITIAL_VALUE, so a row that totalled 100% still totals 100%.
Public Sub SetYearColumns(ByVal Kind As String, ByVal StartYear As Variant, _
                          ByVal NewCount As Long)
    Dim target As ListObject
    Dim fixedCols As Long, i As Long, r As Long, rowCount As Long
    Dim numberFormat As String, headerFormat As String, colWidth As Long

    Set target = ProfilingTable(Kind)
    fixedCols = FixedColumnCount(Kind)
    rowCount = modWorkbook.BodyRowCount(target)

    If Kind = KIND_COST Then
        numberFormat = GRID_COST_PROFILING_YEAR_FORMAT
        headerFormat = GRID_COST_PROFILING_HEADER_FORMAT
        colWidth = GRID_COST_PROFILING_YEAR_WIDTH
    Else
        numberFormat = GRID_RISK_PROFILING_YEAR_FORMAT
        headerFormat = GRID_RISK_PROFILING_HEADER_FORMAT
        colWidth = GRID_RISK_PROFILING_YEAR_WIDTH
    End If

    ' Shrink from the tail: the removed project years are always the last ones.
    Do While target.ListColumns.Count > fixedCols + NewCount
        target.ListColumns(target.ListColumns.Count).Delete
    Loop

    ' Grow at the tail, seeding each new project year at 0%.
    Do While target.ListColumns.Count < fixedCols + NewCount
        Dim added As ListColumn
        Set added = target.ListColumns.Add
        added.DataBodyRange.NumberFormat = numberFormat
        added.Range.ColumnWidth = colWidth
        For r = 1 To rowCount
            ' Only an identified row gets a profile cell; a reserved blank row is
            ' not a driver and must stay empty.
            If Len(modWorkbook.TextOf(modWorkbook.CellIn(target, r, 1))) > 0 Then
                added.DataBodyRange.Cells(r, 1).Value = PROFILE_INITIAL_VALUE
            End If
        Next r
    Loop

    ' Relabel. Headers are display only; no value moves because of a relabel.
    For i = 1 To NewCount
        target.ListColumns(fixedCols + i).Name = CStr(CLng(StartYear) + i - 1)
        target.HeaderRowRange.Cells(1, fixedCols + i).NumberFormat = headerFormat
    Next i
End Sub

' Removes every project-year column. Used when the applied timeline is cleared.
Public Sub ClearYearColumns(ByVal Kind As String)
    SetYearColumns Kind, 0, 0
End Sub

' ---------------------------------------------------------------------------
' Destructive assessment - runs BEFORE anything is modified
' ---------------------------------------------------------------------------
' Counts the profiling cells a shrink to NewCount would destroy and collects
' representative permanent IDs.
'
' A cell counts as a loss when modWorkbook.IsDataCell holds: blank and numeric zero
' destroy nothing, and EVERYTHING else does. Counting only numeric non-zero cells
' was too narrow -- a percentage pasted as text, or an error value, would have been
' deleted by a duration reduction with no destructive warning at all.
Public Function CountDataBeyond(ByVal Kind As String, ByVal NewCount As Long, _
                                ByRef AffectedIds() As String, _
                                ByRef AffectedCount As Long) As Long
    Dim target As ListObject
    Dim fixedCols As Long, existing As Long, r As Long, c As Long
    Dim rowCount As Long, hits As Long
    Dim idText As String
    Dim seen As Object

    Set target = ProfilingTable(Kind)
    fixedCols = FixedColumnCount(Kind)
    existing = target.ListColumns.Count - fixedCols
    rowCount = modWorkbook.BodyRowCount(target)
    Set seen = CreateObject("Scripting.Dictionary")

    ReDim AffectedIds(1 To IIf(rowCount < 1, 1, rowCount))
    AffectedCount = 0
    hits = 0

    For r = 1 To rowCount
        idText = modWorkbook.TextOf(modWorkbook.CellIn(target, r, 1))
        If Len(idText) > 0 Then
            For c = NewCount + 1 To existing
                Dim cell As Range
                Set cell = modWorkbook.CellIn(target, r, fixedCols + c)
                If modWorkbook.IsDataCell(cell) Then
                    hits = hits + 1
                    If Not seen.Exists(idText) Then
                        seen.Add idText, True
                        AffectedCount = AffectedCount + 1
                        AffectedIds(AffectedCount) = idText
                    End If
                End If
            Next c
        End If
    Next r

    CountDataBeyond = hits
End Function

' ---------------------------------------------------------------------------
' Row synchronisation - keyed by permanent ID
' ---------------------------------------------------------------------------
' Rebuilds the grid's rows so that, for every non-blank driver ID, exactly one
' profiling row exists and carries that ID. Percentages follow their ID: a row
' that survives keeps its values whatever position it now occupies.
Public Sub SyncRows(ByVal Kind As String)
    Dim target As ListObject, register As ListObject
    Dim fixedCols As Long, yearCols As Long
    Dim idCol As Long, traceCol As Long
    Dim regRowCount As Long, gridRowCount As Long
    Dim r As Long, c As Long

    Set target = ProfilingTable(Kind)
    fixedCols = FixedColumnCount(Kind)
    yearCols = target.ListColumns.Count - fixedCols

    If Kind = KIND_COST Then
        Set register = modWorkbook.Lo(SH_COST_LINES, TBL_COST_LINES)
        idCol = COL_COST_LINES_COST_LINE_ID
        traceCol = COL_COST_LINES_DESCRIPTION
    Else
        Set register = modWorkbook.Lo(SH_RISK_REGISTER, TBL_RISK_REGISTER)
        idCol = COL_RISK_REGISTER_RISK_ID
        traceCol = COL_RISK_REGISTER_RISK_NAME
    End If

    regRowCount = modWorkbook.BodyRowCount(register)
    gridRowCount = modWorkbook.BodyRowCount(target)

    ' Snapshot the existing percentages against their permanent IDs, so ownership
    ' survives any reordering of either table.
    Dim held As Object
    Set held = CreateObject("Scripting.Dictionary")
    For r = 1 To gridRowCount
        Dim key As String
        key = modWorkbook.TextOf(modWorkbook.CellIn(target, r, 1))
        If Len(key) > 0 Then
            If Not held.Exists(key) Then
                Dim values() As Variant
                ReDim values(1 To IIf(yearCols < 1, 1, yearCols))
                For c = 1 To yearCols
                    Dim source As Range
                    Set source = modWorkbook.CellIn(target, r, fixedCols + c)
                    If modWorkbook.IsEmptyCell(source) Then
                        values(c) = Empty
                    Else
                        values(c) = source.Value
                    End If
                Next c
                held.Add key, values
            End If
        End If
    Next r

    ' Rewrite the grid in register order.
    Dim writeRow As Long
    writeRow = 0
    For r = 1 To regRowCount
        Dim driverId As String
        driverId = modWorkbook.TextOf(modWorkbook.CellIn(register, r, idCol))
        If Len(driverId) > 0 Then
            writeRow = writeRow + 1
            If writeRow > modWorkbook.BodyRowCount(target) Then
                target.ListRows.Add
            End If
            modWorkbook.CellIn(target, writeRow, 1).Value = driverId
            modWorkbook.CellIn(target, writeRow, 2).Value = _
                modWorkbook.TextOf(modWorkbook.CellIn(register, r, traceCol))
            For c = 1 To yearCols
                Dim cell As Range
                Set cell = modWorkbook.CellIn(target, writeRow, fixedCols + c)
                If held.Exists(driverId) Then
                    ' An EXISTING (permanent ID, project-year index) value is
                    ' preserved exactly -- including blank. Filling a blank with 0%
                    ' here would repair invalid user data inside a structural
                    ' operation and hide it from the Model Check phase whose job is
                    ' to report it. Only a genuinely new driver, or a genuinely new
                    ' project-year column, starts at 0%.
                    Dim kept As Variant
                    kept = held(driverId)
                    If IsEmpty(kept(c)) Then
                        cell.ClearContents
                    Else
                        cell.Value = kept(c)
                    End If
                Else
                    ' A newly identified driver starts at 0% in every project year.
                    cell.Value = PROFILE_INITIAL_VALUE
                End If
            Next c
        End If
    Next r

    ' Clear the tail: rows below the last identified driver hold no profile.
    gridRowCount = modWorkbook.BodyRowCount(target)
    For r = writeRow + 1 To gridRowCount
        For c = 1 To target.ListColumns.Count
            modWorkbook.CellIn(target, r, c).ClearContents
        Next c
    Next r
End Sub

' Removes the profiling row owned by one permanent ID. Identity is the operation
' key throughout: no row number is ever inferred after the user has confirmed.
Public Sub RemoveRow(ByVal Kind As String, ByVal PermanentId As String)
    Dim target As ListObject
    Dim r As Long, rowCount As Long
    Set target = ProfilingTable(Kind)
    rowCount = modWorkbook.BodyRowCount(target)
    For r = rowCount To 1 Step -1
        If StrComp(modWorkbook.TextOf(modWorkbook.CellIn(target, r, 1)), PermanentId, vbTextCompare) = 0 Then
            If rowCount > 1 Then
                target.ListRows(r).Delete
            Else
                target.DataBodyRange.Rows(r).ClearContents
            End If
            Exit Sub
        End If
    Next r
End Sub

' ---------------------------------------------------------------------------
' Inspection, used by modStructuralCheck and by the Windows functional harness
' ---------------------------------------------------------------------------
Public Function IdList(ByVal Kind As String) As String
    Dim target As ListObject
    Dim r As Long, rowCount As Long, out As String, idText As String
    Set target = ProfilingTable(Kind)
    rowCount = modWorkbook.BodyRowCount(target)
    For r = 1 To rowCount
        idText = modWorkbook.TextOf(modWorkbook.CellIn(target, r, 1))
        If Len(idText) > 0 Then
            If Len(out) > 0 Then out = out & ","
            out = out & idText
        End If
    Next r
    IdList = out
End Function

Public Function YearHeaders(ByVal Kind As String) As String
    Dim target As ListObject
    Dim fixedCols As Long, i As Long, out As String
    Set target = ProfilingTable(Kind)
    fixedCols = FixedColumnCount(Kind)
    For i = fixedCols + 1 To target.ListColumns.Count
        If Len(out) > 0 Then out = out & ","
        out = out & target.ListColumns(i).Name
    Next i
    YearHeaders = out
End Function

Public Function ValueFor(ByVal Kind As String, ByVal PermanentId As String, _
                         ByVal ProjectYearIndex As Long) As Variant
    Dim target As ListObject
    Dim fixedCols As Long, r As Long, rowCount As Long
    Set target = ProfilingTable(Kind)
    fixedCols = FixedColumnCount(Kind)
    rowCount = modWorkbook.BodyRowCount(target)
    For r = 1 To rowCount
        If StrComp(modWorkbook.TextOf(modWorkbook.CellIn(target, r, 1)), PermanentId, vbTextCompare) = 0 Then
            ValueFor = modWorkbook.CellIn(target, r, fixedCols + ProjectYearIndex).Value
            Exit Function
        End If
    Next r
    ValueFor = Empty
End Function

Public Sub SetValueFor(ByVal Kind As String, ByVal PermanentId As String, _
                       ByVal ProjectYearIndex As Long, ByVal NewValue As Variant)
    Dim target As ListObject
    Dim fixedCols As Long, r As Long, rowCount As Long
    Set target = ProfilingTable(Kind)
    fixedCols = FixedColumnCount(Kind)
    rowCount = modWorkbook.BodyRowCount(target)
    For r = 1 To rowCount
        If StrComp(modWorkbook.TextOf(modWorkbook.CellIn(target, r, 1)), PermanentId, vbTextCompare) = 0 Then
            modWorkbook.CellIn(target, r, fixedCols + ProjectYearIndex).Value = NewValue
            Exit Sub
        End If
    Next r
End Sub
