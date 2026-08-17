Attribute VB_Name = "modStructuralCheck"
Option Explicit

' ===========================================================================
' PCCM - Phase-4 structural revalidation
' ===========================================================================
' This is NOT the Model Check engine. It has no severities, no register, no
' overall status and no simulation gate; those belong to a later phase. It checks
' one thing: that a structural operation left the workbook structurally coherent.
'
' Two rules govern its behaviour:
'
'   * It REPORTS. It never repairs. Silently fixing a structural inconsistency
'     would hide the defect that produced it.
'   * A failure makes the macro operation FAIL, and the caller restores the
'     blocks it touched.
'
' It deliberately checks nothing about business validity: a profiling row that
' does not total 100%, a missing inflation rate and a blank unit cost are all
' Model Check concerns and all pass here.
' ===========================================================================

Public Function ValidateStructure() As String
    Dim problems As String
    problems = problems & CheckAppliedTriple()
    problems = problems & CheckProfilingShape(modProfiling.CostKind(), TBL_COST_PROFILING)
    problems = problems & CheckProfilingShape(modProfiling.RiskKind(), TBL_RISK_PROFILING)
    problems = problems & CheckProfilingIdentity(modProfiling.CostKind(), "cost", TBL_COST_PROFILING)
    problems = problems & CheckProfilingIdentity(modProfiling.RiskKind(), "risk", TBL_RISK_PROFILING)
    problems = problems & CheckIdPatterns("cost", ID_PREFIX_COST_LINE, TBL_COST_LINES)
    problems = problems & CheckIdPatterns("risk", ID_PREFIX_RISK, TBL_RISK_REGISTER)
    problems = problems & CheckCounters()
    problems = problems & CheckInflationHeaders()
    problems = problems & CheckInflationProfiles()
    ValidateStructure = problems
End Function

Private Function Fault(ByVal CheckKey As String, ByVal Text As String) As String
    Fault = "  [" & CheckKey & "] " & Text & vbCrLf
End Function

' ---------------------------------------------------------------------------
Private Function CheckAppliedTriple() As String
    Dim baseYear As Variant, startYear As Variant, duration As Variant
    Dim present As Long, problems As String

    baseYear = modWorkbook.ReadValue(NM_APPLIED_BASE_YEAR)
    startYear = modWorkbook.ReadValue(NM_APPLIED_START_YEAR)
    duration = modWorkbook.ReadValue(NM_APPLIED_DURATION)

    If Not IsEmpty(baseYear) Then present = present + 1
    If Not IsEmpty(startYear) Then present = present + 1
    If Not IsEmpty(duration) Then present = present + 1

    ' Wholly blank or wholly populated. A partially applied triple is exactly the
    ' half-applied state this phase exists to make impossible.
    If present <> 0 And present <> 3 Then
        problems = problems & Fault(CHK_APPLIED_TRIPLE_CONSISTENT, _
            "the applied timeline is partially populated (" & present & " of 3 values).")
        CheckAppliedTriple = problems
        Exit Function
    End If
    If present = 0 Then Exit Function

    If Not modWorkbook.IsWholeNumber(baseYear) Or Not modWorkbook.IsWholeNumber(startYear) _
       Or Not modWorkbook.IsWholeNumber(duration) Then
        problems = problems & Fault(CHK_APPLIED_TRIPLE_CONSISTENT, _
            "an applied timeline value is not a whole number.")
        CheckAppliedTriple = problems
        Exit Function
    End If

    If CLng(duration) < 1 Then
        problems = problems & Fault(CHK_APPLIED_TRIPLE_CONSISTENT, _
            "applied duration " & duration & " is below 1.")
    End If
    If CLng(baseYear) > CLng(startYear) Then
        problems = problems & Fault(CHK_APPLIED_TRIPLE_CONSISTENT, _
            "applied base year " & baseYear & " is later than applied start year " & startYear & ".")
    End If
    If CLng(baseYear) < LIMIT_MIN_YEAR Or CLng(startYear) > LIMIT_MAX_YEAR Then
        problems = problems & Fault(CHK_APPLIED_TRIPLE_CONSISTENT, _
            "an applied year is outside the supported window " & LIMIT_MIN_YEAR & "-" & LIMIT_MAX_YEAR & ".")
    End If

    CheckAppliedTriple = problems
End Function

' ---------------------------------------------------------------------------
Private Function CheckProfilingShape(ByVal Kind As String, ByVal TableName As String) As String
    Dim duration As Variant, startYear As Variant
    Dim expected As Long, actual As Long, i As Long
    Dim problems As String
    Dim target As ListObject

    duration = modWorkbook.ReadValue(NM_APPLIED_DURATION)
    startYear = modWorkbook.ReadValue(NM_APPLIED_START_YEAR)
    If IsEmpty(duration) Then expected = 0 Else expected = CLng(duration)

    Set target = modProfiling.ProfilingTable(Kind)
    actual = modProfiling.YearColumnCount(Kind)

    If actual <> expected Then
        problems = problems & Fault(CHK_PROFILING_COLUMN_COUNT, _
            TableName & " carries " & actual & " project-year column(s); the applied duration is " & _
            expected & ".")
        CheckProfilingShape = problems
        Exit Function
    End If

    If expected > 0 Then
        Dim fixedCols As Long
        fixedCols = modProfiling.FixedColumnCount(Kind)
        For i = 1 To expected
            Dim headerText As String
            headerText = target.ListColumns(fixedCols + i).Name
            If Not IsNumeric(headerText) Then
                problems = problems & Fault(CHK_PROFILING_YEAR_HEADERS, _
                    TableName & " project-year header " & i & " is '" & headerText & "', not a year.")
            ElseIf CLng(headerText) <> CLng(startYear) + i - 1 Then
                problems = problems & Fault(CHK_PROFILING_YEAR_HEADERS, _
                    TableName & " project-year header " & i & " is " & headerText & _
                    "; expected " & (CLng(startYear) + i - 1) & ".")
            End If
        Next i
    End If

    ' Header row intact and the table still consistent with its rendered extent.
    If target.HeaderRowRange Is Nothing Then
        problems = problems & Fault(CHK_GRID_SHAPE, TableName & " has lost its header row.")
    ElseIf target.ListColumns.Count <> modProfiling.FixedColumnCount(Kind) + expected Then
        problems = problems & Fault(CHK_GRID_SHAPE, _
            TableName & " column count is inconsistent with its fixed and generated columns.")
    End If

    CheckProfilingShape = problems
End Function

' ---------------------------------------------------------------------------
' One profiling row per identified driver, and no more. Matched by permanent ID,
' never by row position, so a reordered register is not an inconsistency.
Private Function CheckProfilingIdentity(ByVal Kind As String, ByVal DriverKind As String, _
                                        ByVal TableName As String) As String
    Dim register As ListObject, grid As ListObject
    Dim r As Long, rowCount As Long, idCol As Long
    Dim driverIds As Object, gridIds As Object
    Dim problems As String
    Dim idText As String

    Set register = modDrivers.RegisterTable(DriverKind)
    Set grid = modProfiling.ProfilingTable(Kind)
    idCol = modDrivers.IdColumn(DriverKind)
    Set driverIds = CreateObject("Scripting.Dictionary")
    Set gridIds = CreateObject("Scripting.Dictionary")

    rowCount = modWorkbook.BodyRowCount(register)
    For r = 1 To rowCount
        idText = modWorkbook.TextOf(modWorkbook.CellIn(register, r, idCol))
        If Len(idText) > 0 Then
            If driverIds.Exists(idText) Then
                problems = problems & Fault(CHK_NO_DUPLICATE_IDS, _
                    register.Name & " contains the identifier " & idText & " more than once.")
            Else
                driverIds.Add idText, True
            End If
        End If
    Next r

    rowCount = modWorkbook.BodyRowCount(grid)
    For r = 1 To rowCount
        idText = modWorkbook.TextOf(modWorkbook.CellIn(grid, r, 1))
        If Len(idText) > 0 Then
            If gridIds.Exists(idText) Then
                problems = problems & Fault(CHK_NO_DUPLICATE_IDS, _
                    TableName & " contains the identifier " & idText & " more than once.")
            Else
                gridIds.Add idText, True
            End If
        End If
    Next r

    Dim key As Variant
    For Each key In driverIds.Keys
        If Not gridIds.Exists(key) Then
            problems = problems & Fault(IIf(DriverKind = "cost", CHK_COST_PROFILING_IDS_MATCH, CHK_RISK_PROFILING_IDS_MATCH), _
                "driver " & key & " has no row in " & TableName & ".")
        End If
    Next key
    For Each key In gridIds.Keys
        If Not driverIds.Exists(key) Then
            problems = problems & Fault(IIf(DriverKind = "cost", CHK_COST_PROFILING_IDS_MATCH, CHK_RISK_PROFILING_IDS_MATCH), _
                TableName & " has a row for " & key & ", which is not an identified driver.")
        End If
    Next key

    CheckProfilingIdentity = problems
End Function

' ---------------------------------------------------------------------------
Private Function CheckIdPatterns(ByVal DriverKind As String, ByVal Prefix As String, _
                                 ByVal TableName As String) As String
    Dim register As ListObject
    Dim r As Long, rowCount As Long, idCol As Long
    Dim idText As String, tail As String
    Dim problems As String

    Set register = modDrivers.RegisterTable(DriverKind)
    idCol = modDrivers.IdColumn(DriverKind)
    rowCount = modWorkbook.BodyRowCount(register)

    For r = 1 To rowCount
        idText = modWorkbook.TextOf(modWorkbook.CellIn(register, r, idCol))
        If Len(idText) > 0 Then
            If Len(idText) <= Len(Prefix) Then
                problems = problems & Fault(CHK_ID_PATTERN, _
                    TableName & " row " & r & " holds '" & idText & "', which is not a " & Prefix & " identifier.")
            ElseIf StrComp(Left$(idText, Len(Prefix)), Prefix, vbBinaryCompare) <> 0 Then
                problems = problems & Fault(CHK_ID_PATTERN, _
                    TableName & " row " & r & " holds '" & idText & "'; the required prefix is '" & Prefix & "'.")
            Else
                tail = Mid$(idText, Len(Prefix) + 1)
                If Not IsAllDigits(tail) Then
                    problems = problems & Fault(CHK_ID_PATTERN, _
                        TableName & " row " & r & " holds '" & idText & "'; the sequence part must be digits.")
                End If
            End If
        End If
    Next r

    CheckIdPatterns = problems
End Function

Private Function IsAllDigits(ByVal Text As String) As Boolean
    Dim i As Long
    If Len(Text) = 0 Then Exit Function
    For i = 1 To Len(Text)
        If Mid$(Text, i, 1) < "0" Or Mid$(Text, i, 1) > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function

' ---------------------------------------------------------------------------
' A counter that has fallen below an identifier it already issued would reissue
' that identifier. The test is 'at least the highest issued', never 'equals the
' number of rows': deletion leaves the counter deliberately ahead of the count.
Private Function CheckCounters() As String
    Dim problems As String
    Dim kinds As Variant, i As Long
    kinds = Array("cost", "risk")
    For i = LBound(kinds) To UBound(kinds)
        Dim counterValue As Long, highest As Long
        counterValue = modDrivers.ReadCounter(CStr(kinds(i)))
        highest = modDrivers.HighestIssued(CStr(kinds(i)))
        If counterValue < highest Then
            problems = problems & Fault(CHK_COUNTERS_NOT_BEHIND, _
                "the " & kinds(i) & " counter is " & counterValue & _
                " but identifier number " & highest & " has already been issued; the next " & _
                "allocation would reuse an identifier.")
        End If
    Next i
    CheckCounters = problems
End Function

' ---------------------------------------------------------------------------
Private Function CheckInflationHeaders() As String
    Dim target As ListObject
    Dim fixedCols As Long, expected As Long, actual As Long, i As Long
    Dim firstYear As Variant
    Dim problems As String

    Set target = modWorkbook.Lo(SH_INFLATION, TBL_INFLATION)
    fixedCols = GRID_INFLATION_FIXED_COLS
    expected = modInflation.RequiredYearCount()
    actual = target.ListColumns.Count - fixedCols

    If actual <> expected Then
        problems = problems & Fault(CHK_INFLATION_YEAR_HEADERS, _
            TBL_INFLATION & " carries " & actual & " calendar-year column(s); the applied span " & _
            "requires " & expected & ".")
        CheckInflationHeaders = problems
        Exit Function
    End If

    If expected > 0 Then
        firstYear = modInflation.RequiredFirstYear()
        For i = 1 To expected
            Dim headerText As String
            headerText = target.ListColumns(fixedCols + i).Name
            If Not IsNumeric(headerText) Then
                problems = problems & Fault(CHK_INFLATION_YEAR_HEADERS, _
                    TBL_INFLATION & " year header " & i & " is '" & headerText & "', not a year.")
            ElseIf CLng(headerText) <> CLng(firstYear) + i - 1 Then
                problems = problems & Fault(CHK_INFLATION_YEAR_HEADERS, _
                    TBL_INFLATION & " year header " & i & " is " & headerText & "; expected " & _
                    (CLng(firstYear) + i - 1) & ".")
            End If
        Next i
    End If

    If target.HeaderRowRange Is Nothing Then
        problems = problems & Fault(CHK_GRID_SHAPE, TBL_INFLATION & " has lost its header row.")
    End If

    CheckInflationHeaders = problems
End Function

' ---------------------------------------------------------------------------
Private Function CheckInflationProfiles() As String
    Dim target As ListObject
    Dim r As Long, rowCount As Long
    Dim wanted As Object, seen As Object
    Dim problems As String
    Dim rowName As String

    Set target = modWorkbook.Lo(SH_INFLATION, TBL_INFLATION)
    Set wanted = modInflation.ProfileNameSet()
    Set seen = CreateObject("Scripting.Dictionary")
    rowCount = modWorkbook.BodyRowCount(target)

    For r = 1 To rowCount
        rowName = modWorkbook.TextOf(modWorkbook.CellIn(target, r, 1))
        If Len(rowName) > 0 Then
            If seen.Exists(LCase$(rowName)) Then
                problems = problems & Fault(CHK_INFLATION_PROFILE_ROWS, _
                    TBL_INFLATION & " contains the profile '" & rowName & "' more than once.")
            Else
                seen.Add LCase$(rowName), True
            End If
            If Not wanted.Exists(LCase$(rowName)) Then
                problems = problems & Fault(CHK_INFLATION_PROFILE_ROWS, _
                    TBL_INFLATION & " has a row for '" & rowName & "', which is not in the Config " & _
                    "profile master.")
            End If
        End If
    Next r

    Dim key As Variant
    For Each key In wanted.Keys
        If Not seen.Exists(key) Then
            problems = problems & Fault(CHK_INFLATION_PROFILE_ROWS, _
                "Config profile '" & wanted(key) & "' has no row in " & TBL_INFLATION & ".")
        End If
    Next key

    CheckInflationProfiles = problems
End Function

' ---------------------------------------------------------------------------
' Callable from the Windows functional harness. Returns "" when the workbook is
' structurally coherent, otherwise the accumulated fault list.
Public Function PCCM_StructuralReport() As String
    PCCM_StructuralReport = ValidateStructure()
End Function
