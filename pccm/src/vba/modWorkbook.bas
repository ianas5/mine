Attribute VB_Name = "modWorkbook"
Option Explicit

' ===========================================================================
' PCCM - Workbook access primitives
' ===========================================================================
' Phase 4. Plumbing only: sheets, defined names, list objects, blank handling.
' This module holds NO structural policy and NO business rule. If a decision is
' being made, it belongs in modTimeline, modProfiling, modInflation or modDrivers.
'
' Every address comes from modConstants, which is generated from the structure
' contract. Nothing here hardcodes a sheet name, a cell reference or a table name.
' ===========================================================================

' ---------------------------------------------------------------------------
' Sheets and tables
' ---------------------------------------------------------------------------
Public Function Sh(ByVal SheetName As String) As Worksheet
    Set Sh = ThisWorkbook.Worksheets(SheetName)
End Function

Public Function Lo(ByVal SheetName As String, ByVal TableName As String) As ListObject
    Set Lo = Sh(SheetName).ListObjects(TableName)
End Function

Public Function LoExists(ByVal SheetName As String, ByVal TableName As String) As Boolean
    Dim probe As ListObject
    On Error Resume Next
    Set probe = Sh(SheetName).ListObjects(TableName)
    On Error GoTo 0
    LoExists = Not probe Is Nothing
End Function

' Row count of the table's reserved body. A structural grid always has a data body
' because Stage A materialises it with reserved rows.
Public Function BodyRowCount(ByVal Target As ListObject) As Long
    If Target.DataBodyRange Is Nothing Then
        BodyRowCount = 0
    Else
        BodyRowCount = Target.DataBodyRange.Rows.Count
    End If
End Function

Public Function CellIn(ByVal Target As ListObject, ByVal RowIndex As Long, _
                       ByVal ColumnIndex As Long) As Range
    Set CellIn = Target.DataBodyRange.Cells(RowIndex, ColumnIndex)
End Function

' ---------------------------------------------------------------------------
' Defined names
' ---------------------------------------------------------------------------
Public Function NameExists(ByVal DefinedName As String) As Boolean
    Dim probe As Name
    On Error Resume Next
    Set probe = ThisWorkbook.Names(DefinedName)
    On Error GoTo 0
    NameExists = Not probe Is Nothing
End Function

Public Function NamedCell(ByVal DefinedName As String) As Range
    Set NamedCell = ThisWorkbook.Names(DefinedName).RefersToRange
End Function

' Empty when the cell is blank. Callers must distinguish "blank" from "zero":
' a blank inflation rate and a zero inflation rate are different assumptions.
Public Function ReadValue(ByVal DefinedName As String) As Variant
    Dim c As Range
    Set c = NamedCell(DefinedName)
    If IsEmptyCell(c) Then
        ReadValue = Empty
    Else
        ReadValue = c.Value
    End If
End Function

' Bounded read. A corrupt or oversized stored value yields the fallback rather than
' overflowing; the caller decides what an out-of-range value means.
Public Function ReadLongInRange(ByVal DefinedName As String, ByVal MinValue As Double, _
                                ByVal MaxValue As Double, ByVal Fallback As Long) As Long
    Dim d As Double
    If IsWholeInRange(ReadValue(DefinedName), MinValue, MaxValue, d) Then
        ReadLongInRange = SafeLong(d)
    Else
        ReadLongInRange = Fallback
    End If
End Function

Public Sub WriteValue(ByVal DefinedName As String, ByVal NewValue As Variant)
    Dim c As Range
    Set c = NamedCell(DefinedName)
    If IsEmpty(NewValue) Then
        c.ClearContents
    Else
        c.Value = NewValue
    End If
End Sub

' ---------------------------------------------------------------------------
' Blank and type handling
' ---------------------------------------------------------------------------
Public Function IsEmptyCell(ByVal Target As Range) As Boolean
    If IsError(Target.Value) Then Exit Function
    IsEmptyCell = (Len(Trim$(CStr(Target.Value & ""))) = 0)
End Function

' True when losing this cell would lose real user data.
'
'   blank         -> no loss; the user entered nothing
'   numeric zero  -> no loss; 0% is what a new project-year cell is created with
'   anything else -> LOSS, including text and error values
'
' Counting only numeric non-zero cells was too narrow: a percentage accidentally
' pasted as text would have been deleted by a duration reduction with no warning.
Public Function IsDataCell(ByVal Target As Range) As Boolean
    If IsError(Target.Value) Then
        IsDataCell = True
        Exit Function
    End If
    If IsEmptyCell(Target) Then Exit Function
    If IsNumeric(Target.Value) Then
        Dim d As Double
        If TryReadDouble(Target.Value, d) Then
            IsDataCell = (d <> 0)
            Exit Function
        End If
    End If
    IsDataCell = True
End Function

' Timeline cells are user-controlled and can be changed by paste even where Data
' Validation exists, so an entered value may be a string, an error, or a number far
' outside anything a Long can hold. Every check below therefore works in Double and
' NOTHING here calls CLng: converting an untrusted value before proving it fits is
' exactly how a paste of 1E10 turns a clean rejection into an unhandled Overflow.
Public Function TryReadDouble(ByVal Value As Variant, ByRef Result As Double) As Boolean
    Result = 0
    If IsEmpty(Value) Then Exit Function
    If IsNull(Value) Then Exit Function
    If IsError(Value) Then Exit Function
    If IsObject(Value) Then Exit Function
    If VarType(Value) = vbBoolean Then Exit Function
    If VarType(Value) = vbString Then
        If Len(Trim$(CStr(Value))) = 0 Then Exit Function
    End If
    If Not IsNumeric(Value) Then Exit Function

    On Error GoTo NotADouble
    Result = CDbl(Value)
    On Error GoTo 0
    TryReadDouble = True
    Exit Function

NotADouble:
    Result = 0
    TryReadDouble = False
End Function

Public Function IsWholeNumber(ByVal Value As Variant) As Boolean
    Dim d As Double
    If Not TryReadDouble(Value, d) Then Exit Function
    IsWholeNumber = (d = Int(d))
End Function

' True only when the value is a whole number AND lies inside [MinValue, MaxValue].
' Callers use this to prove a value is representable and admissible BEFORE any
' arithmetic depends on it and before any conversion to Long.
Public Function IsWholeInRange(ByVal Value As Variant, ByVal MinValue As Double, _
                               ByVal MaxValue As Double, ByRef Result As Double) As Boolean
    Result = 0
    Dim d As Double
    If Not TryReadDouble(Value, d) Then Exit Function
    If d <> Int(d) Then Exit Function
    If d < MinValue Then Exit Function
    If d > MaxValue Then Exit Function
    Result = d
    IsWholeInRange = True
End Function

' Conversion is safe ONLY after IsWholeInRange has proved the bound. The guard is
' kept here as well so a future caller cannot skip the proof silently.
Public Function SafeLong(ByVal Value As Double) As Long
    If Value < -2147483648# Or Value > 2147483647# Then
        Err.Raise vbObjectError + 5010, "modWorkbook.SafeLong", _
                  "Value " & Value & " is outside the range a Long can represent. " & _
                  "It must be bounded before conversion."
    End If
    SafeLong = CLng(Value)
End Function

Public Function TextOf(ByVal Target As Range) As String
    TextOf = Trim$(CStr(Target.Value & ""))
End Function

' ---------------------------------------------------------------------------
' Snapshots for logical restore
' ---------------------------------------------------------------------------
' A snapshot is plain data: one table's shape, presentation and contents. It is NOT
' a workbook copy, an undo journal or a transaction log. It exists so one operation
' can put back exactly the block it was about to change.
'
' Shape means BOTH dimensions. Restoring only the column count was not enough: a
' failed Add that had already grown the table left an extra row behind, a failed
' Delete left a missing one, and a profiling synchronisation that changed row shape
' survived the rollback.
'
' Presentation is captured too, because a rolled-back year column that came back
' without its number format and width would leave the workbook technically correct
' and practically unusable.
Public Type TableSnapshot
    Captured      As Boolean
    RowCount      As Long
    ColumnCount   As Long
    Headers()     As Variant
    Values()      As Variant
    NumberFormats() As Variant
    ColumnWidths()  As Variant
End Type

Public Function SnapshotTable(ByVal Target As ListObject) As TableSnapshot
    Dim s As TableSnapshot
    Dim r As Long, c As Long

    s.ColumnCount = Target.ListColumns.Count
    s.RowCount = BodyRowCount(Target)

    ReDim s.Headers(1 To s.ColumnCount)
    ReDim s.NumberFormats(1 To s.ColumnCount)
    ReDim s.ColumnWidths(1 To s.ColumnCount)
    ReDim s.Values(1 To IIf(s.RowCount < 1, 1, s.RowCount), 1 To s.ColumnCount)

    For c = 1 To s.ColumnCount
        s.Headers(c) = Target.HeaderRowRange.Cells(1, c).Value
        s.ColumnWidths(c) = Target.ListColumns(c).Range.ColumnWidth
        If s.RowCount > 0 Then
            s.NumberFormats(c) = Target.DataBodyRange.Cells(1, c).NumberFormat
        Else
            s.NumberFormats(c) = Target.HeaderRowRange.Cells(1, c).NumberFormat
        End If
    Next c

    For r = 1 To s.RowCount
        For c = 1 To s.ColumnCount
            s.Values(r, c) = Target.DataBodyRange.Cells(r, c).Value
        Next c
    Next r

    s.Captured = True
    SnapshotTable = s
End Function

' Restores a snapshot taken by SnapshotTable: shape first, then presentation, then
' contents. Deliberately NOT wrapped in On Error Resume Next -- a restore that
' cannot complete is a fault the caller must see, not one to bury.
Public Sub RestoreTable(ByVal Target As ListObject, ByRef Snapshot As TableSnapshot)
    If Not Snapshot.Captured Then
        Err.Raise vbObjectError + 5011, "modWorkbook.RestoreTable", _
                  "No snapshot was captured for " & Target.Name & "; refusing to restore."
    End If

    Dim r As Long, c As Long

    ' --- columns ---------------------------------------------------------
    Do While Target.ListColumns.Count > Snapshot.ColumnCount
        Target.ListColumns(Target.ListColumns.Count).Delete
    Loop
    Do While Target.ListColumns.Count < Snapshot.ColumnCount
        Target.ListColumns.Add
    Loop

    ' --- rows ------------------------------------------------------------
    ' An Excel ListObject cannot hold zero data rows while it has a DataBodyRange,
    ' and deleting the last one removes the body entirely. That case is handled
    ' explicitly rather than left to chance: the final row is cleared, not deleted.
    Do While BodyRowCount(Target) > Snapshot.RowCount
        If BodyRowCount(Target) = 1 Then
            Target.DataBodyRange.Rows(1).ClearContents
            Exit Do
        End If
        Target.ListRows(BodyRowCount(Target)).Delete
    Loop
    Do While BodyRowCount(Target) < Snapshot.RowCount
        Target.ListRows.Add
    Loop

    If BodyRowCount(Target) <> Snapshot.RowCount Then
        Err.Raise vbObjectError + 5012, "modWorkbook.RestoreTable", _
                  Target.Name & " restored to " & BodyRowCount(Target) & _
                  " row(s); the snapshot recorded " & Snapshot.RowCount & "."
    End If
    If Target.ListColumns.Count <> Snapshot.ColumnCount Then
        Err.Raise vbObjectError + 5013, "modWorkbook.RestoreTable", _
                  Target.Name & " restored to " & Target.ListColumns.Count & _
                  " column(s); the snapshot recorded " & Snapshot.ColumnCount & "."
    End If

    ' --- presentation, before contents so formats apply to the values ----
    For c = 1 To Snapshot.ColumnCount
        Target.HeaderRowRange.Cells(1, c).Value = Snapshot.Headers(c)
        Target.ListColumns(c).Range.ColumnWidth = Snapshot.ColumnWidths(c)
        If Snapshot.RowCount > 0 Then
            Target.ListColumns(c).DataBodyRange.NumberFormat = Snapshot.NumberFormats(c)
        End If
    Next c

    ' --- contents --------------------------------------------------------
    For r = 1 To Snapshot.RowCount
        For c = 1 To Snapshot.ColumnCount
            If IsEmpty(Snapshot.Values(r, c)) Then
                Target.DataBodyRange.Cells(r, c).ClearContents
            Else
                Target.DataBodyRange.Cells(r, c).Value = Snapshot.Values(r, c)
            End If
        Next c
    Next r
End Sub

' ---------------------------------------------------------------------------
' Small utilities
' ---------------------------------------------------------------------------
Public Function JoinLimited(ByRef Items() As String, ByVal Count As Long, _
                            ByVal Limit As Long, ByVal Separator As String) As String
    Dim i As Long, out As String
    For i = 1 To Count
        If i > Limit Then
            out = out & Separator & "..."
            Exit For
        End If
        If Len(out) > 0 Then out = out & Separator
        out = out & Items(i)
    Next i
    JoinLimited = out
End Function
