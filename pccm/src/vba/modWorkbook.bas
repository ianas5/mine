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

Public Function ReadLong(ByVal DefinedName As String, ByVal Fallback As Long) As Long
    Dim v As Variant
    v = ReadValue(DefinedName)
    If IsWholeNumber(v) Then
        ReadLong = CLng(v)
    Else
        ReadLong = Fallback
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
    IsEmptyCell = (Len(Trim$(CStr(Target.Value & ""))) = 0)
End Function

Public Function IsWholeNumber(ByVal Value As Variant) As Boolean
    If IsEmpty(Value) Then Exit Function
    If IsNull(Value) Then Exit Function
    If VarType(Value) = vbString Then
        If Len(Trim$(CStr(Value))) = 0 Then Exit Function
        If Not IsNumeric(Value) Then Exit Function
    ElseIf Not IsNumeric(Value) Then
        Exit Function
    End If
    Dim d As Double
    d = CDbl(Value)
    IsWholeNumber = (d = Int(d))
End Function

Public Function TextOf(ByVal Target As Range) As String
    TextOf = Trim$(CStr(Target.Value & ""))
End Function

' ---------------------------------------------------------------------------
' Snapshots for logical restore
' ---------------------------------------------------------------------------
' A snapshot is plain data: the values of one table's body plus its header row.
' It is NOT a workbook copy, an undo journal or a transaction log. It exists so
' one operation can put back exactly the blocks it was about to change.
Public Function SnapshotTable(ByVal Target As ListObject) As Variant
    Dim rowCount As Long, colCount As Long
    colCount = Target.ListColumns.Count
    rowCount = BodyRowCount(Target)

    Dim data() As Variant
    ReDim data(0 To rowCount, 1 To colCount)

    Dim c As Long, r As Long
    For c = 1 To colCount
        data(0, c) = Target.HeaderRowRange.Cells(1, c).Value
    Next c
    For r = 1 To rowCount
        For c = 1 To colCount
            data(r, c) = Target.DataBodyRange.Cells(r, c).Value
        Next c
    Next r
    SnapshotTable = data
End Function

' Restores a snapshot taken by SnapshotTable, rebuilding the column count first so
' a partially added or removed year column cannot survive the restore.
Public Sub RestoreTable(ByVal Target As ListObject, ByVal Snapshot As Variant)
    Dim wantCols As Long, wantRows As Long
    wantCols = UBound(Snapshot, 2)
    wantRows = UBound(Snapshot, 1)

    Do While Target.ListColumns.Count > wantCols
        Target.ListColumns(Target.ListColumns.Count).Delete
    Loop
    Do While Target.ListColumns.Count < wantCols
        Target.ListColumns.Add
    Loop

    Dim c As Long, r As Long
    For c = 1 To wantCols
        Target.HeaderRowRange.Cells(1, c).Value = Snapshot(0, c)
    Next c
    For r = 1 To wantRows
        For c = 1 To wantCols
            If IsEmpty(Snapshot(r, c)) Then
                Target.DataBodyRange.Cells(r, c).ClearContents
            Else
                Target.DataBodyRange.Cells(r, c).Value = Snapshot(r, c)
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
