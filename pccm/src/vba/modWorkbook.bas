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
' Logical-rollback snapshot type
' ---------------------------------------------------------------------------
' Declared HERE, not beside the procedures that use it. Everything before the
' first executable procedure is the module's declaration section; a Type after
' that point is a compile error under Option Explicit, the same defect that
' ended Gate-B run 3 in modAppState.
'
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
    Fills()       As Variant
    NumberFormats() As Variant
    ColumnWidths()  As Variant
End Type

' ---------------------------------------------------------------------------
' Collision-safe header renaming
' ---------------------------------------------------------------------------
' Excel requires ListObject column names to be UNIQUE, and it does NOT report a
' collision -- it silently disambiguates by appending a digit. A single sequential
' pass therefore corrupts any OVERLAPPING rename, which is precisely the common
' case here: shifting the project start year from 2028 to 2030 asks
'
'     2028 2029 2030 2031 2032   ->   2030 2031 2032 2033 2034
'
' and renaming column 1 to "2030" while column 3 is still "2030" produced the
' corruption the target machine reported at Gate-B run 4: 20272, 20282, 20292,
' 20302. Excel had appended a "2" to each collided name.
'
' The two-pass rename below is the ONLY header-renaming path in the model. Every
' caller goes through it: modProfiling.SetYearColumns, modInflation.SetYearColumns
' and RestoreTable. Rollback needs it as much as the forward path does -- a
' restore that corrupts the headers it is putting back is not a rollback.
Private Const HEADER_TEMP_PREFIX As String = "PCCM_TMP_HDR_"

' Renames Count columns starting at FirstColumn to exactly FinalNames, which must
' be a 1-based String array. Raises rather than accepting an approximate result.
Public Sub SetHeaderBlock(ByVal Target As ListObject, ByVal FirstColumn As Long, _
                          ByRef FinalNames() As String)
    Dim total As Long, i As Long, j As Long, suffix As Long
    Dim temps() As String
    Dim candidate As String

    total = UBound(FinalNames) - LBound(FinalNames) + 1
    If total <= 0 Then Exit Sub
    If FirstColumn < 1 Or FirstColumn + total - 1 > Target.ListColumns.Count Then
        Err.Raise vbObjectError + 5030, "modWorkbook.SetHeaderBlock", _
                  "Header block " & FirstColumn & ".." & (FirstColumn + total - 1) & _
                  " lies outside " & Target.Name & ", which has " & _
                  Target.ListColumns.Count & " column(s)."
    End If

    ' --- 2. the desired names must be usable -----------------------------
    For i = 1 To total
        If Len(Trim$(FinalNames(i))) = 0 Then
            Err.Raise vbObjectError + 5031, "modWorkbook.SetHeaderBlock", _
                      "Header " & i & " of " & Target.Name & " would be blank."
        End If
        For j = i + 1 To total
            If StrComp(FinalNames(i), FinalNames(j), vbTextCompare) = 0 Then
                Err.Raise vbObjectError + 5032, "modWorkbook.SetHeaderBlock", _
                          "Headers " & i & " and " & j & " of " & Target.Name & _
                          " would both be '" & FinalNames(i) & "'. Excel requires " & _
                          "unique column names, so this rename can never succeed."
            End If
        Next j
    Next i

    ' A name already held by a column OUTSIDE the block is equally impossible, and
    ' this operation is not allowed to rename that column to make room.
    For i = 1 To Target.ListColumns.Count
        If i < FirstColumn Or i > FirstColumn + total - 1 Then
            For j = 1 To total
                If StrComp(Target.ListColumns(i).Name, FinalNames(j), vbTextCompare) = 0 Then
                    Err.Raise vbObjectError + 5033, "modWorkbook.SetHeaderBlock", _
                              "Column " & i & " of " & Target.Name & " is already " & _
                              "named '" & FinalNames(j) & "', which is outside the " & _
                              "block being renamed."
                End If
            Next j
        End If
    Next i

    ' --- 3. deterministic temporary names --------------------------------
    ' Not random. Each candidate is checked against every CURRENT column name,
    ' every DESIRED name and every temporary name already chosen, and gains an
    ' incrementing suffix until it is unique on all three counts.
    ReDim temps(1 To total)
    For i = 1 To total
        suffix = 0
        Do
            candidate = HEADER_TEMP_PREFIX & CStr(FirstColumn + i - 1)
            If suffix > 0 Then candidate = candidate & "_" & CStr(suffix)
            suffix = suffix + 1
        Loop While HeaderNameInUse(Target, candidate) _
                Or NameInArray(FinalNames, candidate, total) _
                Or NameInArray(temps, candidate, i - 1)
        temps(i) = candidate
    Next i

    ' --- 4. first pass: vacate every name in the block --------------------
    For i = 1 To total
        Target.ListColumns(FirstColumn + i - 1).Name = temps(i)
    Next i

    ' --- 5. second pass: no desired name can collide now ------------------
    For i = 1 To total
        Target.ListColumns(FirstColumn + i - 1).Name = FinalNames(i)
    Next i

    ' --- 6/7. verify EXACTLY, never trust auto-disambiguation -------------
    For i = 1 To total
        If StrComp(Target.ListColumns(FirstColumn + i - 1).Name, FinalNames(i), _
                   vbBinaryCompare) <> 0 Then
            Err.Raise vbObjectError + 5034, "modWorkbook.SetHeaderBlock", _
                      "Column " & (FirstColumn + i - 1) & " of " & Target.Name & _
                      " is named '" & Target.ListColumns(FirstColumn + i - 1).Name & _
                      "' after the rename; '" & FinalNames(i) & "' was requested. " & _
                      "Excel altered the name rather than accepting it."
        End If
    Next i
End Sub

Private Function HeaderNameInUse(ByVal Target As ListObject, ByVal Candidate As String) As Boolean
    Dim i As Long
    For i = 1 To Target.ListColumns.Count
        If StrComp(Target.ListColumns(i).Name, Candidate, vbTextCompare) = 0 Then
            HeaderNameInUse = True
            Exit Function
        End If
    Next i
End Function

Private Function NameInArray(ByRef Names() As String, ByVal Candidate As String, _
                             ByVal UpTo As Long) As Boolean
    Dim i As Long
    For i = 1 To UpTo
        If StrComp(Names(i), Candidate, vbTextCompare) = 0 Then
            NameInArray = True
            Exit Function
        End If
    Next i
End Function

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

' Error-safe. A pasted Excel error value in a key or name cell would otherwise raise
' a Type mismatch inside CStr, crashing the very code that is trying to INSPECT the
' corruption. An error cell yields a deterministic, non-blank marker instead, so it
' can never be mistaken for a blank key and never disappears silently.
Public Function TextOf(ByVal Target As Range) As String
    If IsError(Target.Value) Then
        TextOf = ERROR_CELL_MARKER
        Exit Function
    End If
    On Error GoTo Unreadable
    TextOf = Trim$(CStr(Target.Value & ""))
    Exit Function
Unreadable:
    TextOf = ERROR_CELL_MARKER
End Function

Public Function IsErrorText(ByVal Text As String) As Boolean
    IsErrorText = (Text = ERROR_CELL_MARKER)
End Function

' ---------------------------------------------------------------------------
' Input-language treatment for RUNTIME-generated cells
' ---------------------------------------------------------------------------
' A year column created at runtime sits beside model-controlled fixed columns, and
' relying on Excel table-format propagation to give it the editable-input fill is
' not deterministic enough. The fill is applied explicitly, from the contract-emitted
' FILL_INPUT / FILL_LOCKED constants, so a generated editable region is never
' visually ambiguous.
'
' Keyed-ness decides the treatment:
'
'   row HAS a key      -> editable input. The user owns these percentages/rates.
'   row has NO key     -> model-controlled. An unkeyed reserved row must not invite
'                         input, because the model cannot own what it cannot key,
'                         and anything typed there becomes orphan data.
Public Sub PaintYearCells(ByVal Target As ListObject, ByVal FirstYearColumn As Long, _
                          ByVal YearCount As Long, ByVal KeyColumn As Long)
    If YearCount < 1 Then Exit Sub
    Dim r As Long, c As Long, rowCount As Long
    Dim keyed As Boolean
    rowCount = BodyRowCount(Target)
    For r = 1 To rowCount
        keyed = (Len(TextOf(CellIn(Target, r, KeyColumn))) > 0)
        For c = FirstYearColumn To FirstYearColumn + YearCount - 1
            If keyed Then
                CellIn(Target, r, c).Interior.Color = FILL_INPUT
            Else
                CellIn(Target, r, c).Interior.Color = FILL_LOCKED
            End If
        Next c
    Next r
End Sub

' ---------------------------------------------------------------------------
' Orphan structural data
' ---------------------------------------------------------------------------
' INVARIANT: no unkeyed structural row may contain owned data.
'
' A structural row is owned by its key -- a permanent ID, or an inflation profile
' name. Synchronisation rebuilds rows from those keys and clears the tail, so a row
' whose key is blank but whose other cells are not is data the next structural
' operation would silently erase. It is also invisible to every destructive
' assessment, because those are keyed too.
'
' This is a STRUCTURAL fault, not a Model Check business rule, and it is REPORTED,
' never repaired.
Public Function OrphanRows(ByVal Target As ListObject, ByVal KeyColumn As Long, _
                           ByRef Rows() As Long) As Long
    Dim r As Long, c As Long, rowCount As Long, colCount As Long, found As Long
    rowCount = BodyRowCount(Target)
    colCount = Target.ListColumns.Count
    ReDim Rows(1 To IIf(rowCount < 1, 1, rowCount))

    For r = 1 To rowCount
        If Len(TextOf(CellIn(Target, r, KeyColumn))) = 0 Then
            For c = 1 To colCount
                If c <> KeyColumn Then
                    If Not IsEmptyCell(CellIn(Target, r, c)) Then
                        found = found + 1
                        Rows(found) = r
                        Exit For
                    End If
                End If
            Next c
        End If
    Next r
    OrphanRows = found
End Function

Public Function DescribeOrphans(ByVal TableName As String, ByRef Rows() As Long, _
                                ByVal Count As Long) As String
    If Count = 0 Then Exit Function
    Dim i As Long, list As String
    For i = 1 To Count
        If i > 5 Then
            list = list & ", ..."
            Exit For
        End If
        If Len(list) > 0 Then list = list & ", "
        list = list & CStr(Rows(i))
    Next i
    DescribeOrphans = TableName & " row(s) " & list & _
                      " hold data but carry no key. Structural synchronisation would " & _
                      "erase that data, and no destructive assessment can see it because " & _
                      "every assessment is keyed."
End Function

' ---------------------------------------------------------------------------
' Snapshots for logical restore
' ---------------------------------------------------------------------------
' The TableSnapshot type these procedures use is declared at the top of this
' module, in the declaration section, where VBA requires it.
Public Function SnapshotTable(ByVal Target As ListObject) As TableSnapshot
    Dim s As TableSnapshot
    Dim r As Long, c As Long

    s.ColumnCount = Target.ListColumns.Count
    s.RowCount = BodyRowCount(Target)

    ReDim s.Headers(1 To s.ColumnCount)
    ReDim s.NumberFormats(1 To s.ColumnCount)
    ReDim s.ColumnWidths(1 To s.ColumnCount)
    ReDim s.Values(1 To IIf(s.RowCount < 1, 1, s.RowCount), 1 To s.ColumnCount)
    ReDim s.Fills(1 To IIf(s.RowCount < 1, 1, s.RowCount), 1 To s.ColumnCount)

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
            ' Per cell, not per column: within one year column an identified row is
            ' editable while an unkeyed reserved row is model-controlled.
            s.Fills(r, c) = Target.DataBodyRange.Cells(r, c).Interior.Color
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

    ' --- headers, collision-safely -----------------------------------------
    ' Writing header cells one at a time is the same collision-unsafe rename as a
    ' sequential ListColumn.Name loop: restoring 2030..2034 back to 2028..2032
    ' overlaps, and Excel would silently disambiguate. A rollback that corrupts the
    ' headers it is putting back is not a rollback, so it goes through the same
    ' two-pass primitive the forward path uses.
    If Snapshot.ColumnCount > 0 Then
        Dim restoredHeaders() As String
        ReDim restoredHeaders(1 To Snapshot.ColumnCount)
        For c = 1 To Snapshot.ColumnCount
            restoredHeaders(c) = CStr(Snapshot.Headers(c))
        Next c
        SetHeaderBlock Target, 1, restoredHeaders
    End If

    ' --- presentation, before contents so formats apply to the values ----
    For c = 1 To Snapshot.ColumnCount
        Target.ListColumns(c).Range.ColumnWidth = Snapshot.ColumnWidths(c)
        If Snapshot.RowCount > 0 Then
            Target.ListColumns(c).DataBodyRange.NumberFormat = Snapshot.NumberFormats(c)
        End If
    Next c

    ' --- contents and per-cell treatment ---------------------------------
    For r = 1 To Snapshot.RowCount
        For c = 1 To Snapshot.ColumnCount
            If IsEmpty(Snapshot.Values(r, c)) Then
                Target.DataBodyRange.Cells(r, c).ClearContents
            Else
                Target.DataBodyRange.Cells(r, c).Value = Snapshot.Values(r, c)
            End If
            ' A recreated row must come back with its input language intact, or the
            ' workbook is restored in name only: the user could not tell which cells
            ' they still own.
            Target.DataBodyRange.Cells(r, c).Interior.Color = Snapshot.Fills(r, c)
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
