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
    problems = problems & CheckOrphanRows()
    problems = problems & CheckInflationHeaders()
    problems = problems & CheckInflationProfiles()
    ValidateStructure = problems
End Function

Private Function Fault(ByVal CheckKey As String, ByVal Text As String) As String
    Fault = "  [" & CheckKey & "] " & Text & vbCrLf
End Function

' ---------------------------------------------------------------------------
' Every structural invariant is checked INDEPENDENTLY, so a workbook that violates
' several is told about all of them. Nothing here performs arithmetic on a value it
' has not already bounded: inspecting corruption must not itself overflow.
Private Function CheckAppliedTriple() As String
    Dim rawBase As Variant, rawStart As Variant, rawDuration As Variant
    Dim baseYear As Double, startYear As Double, duration As Double
    Dim baseOk As Boolean, startOk As Boolean, durationOk As Boolean
    Dim present As Long, problems As String

    rawBase = modWorkbook.ReadValue(NM_APPLIED_BASE_YEAR)
    rawStart = modWorkbook.ReadValue(NM_APPLIED_START_YEAR)
    rawDuration = modWorkbook.ReadValue(NM_APPLIED_DURATION)

    If Not IsEmpty(rawBase) Then present = present + 1
    If Not IsEmpty(rawStart) Then present = present + 1
    If Not IsEmpty(rawDuration) Then present = present + 1

    ' Wholly blank or wholly populated. A partially applied triple is exactly the
    ' half-applied state this phase exists to make impossible.
    If present = 0 Then Exit Function
    If present <> 3 Then
        CheckAppliedTriple = Fault(CHK_APPLIED_TRIPLE_CONSISTENT, _
            "the applied timeline is partially populated (" & present & " of 3 values).")
        Exit Function
    End If

    ' --- whole-number check, independently per value ------------------------
    If Not modWorkbook.IsWholeNumber(rawBase) Then
        problems = problems & Fault(CHK_APPLIED_TRIPLE_CONSISTENT, _
            "applied Base Year is not a whole number.")
    End If
    If Not modWorkbook.IsWholeNumber(rawStart) Then
        problems = problems & Fault(CHK_APPLIED_TRIPLE_CONSISTENT, _
            "applied Start Year is not a whole number.")
    End If
    If Not modWorkbook.IsWholeNumber(rawDuration) Then
        problems = problems & Fault(CHK_APPLIED_TRIPLE_CONSISTENT, _
            "applied Duration is not a whole number.")
    End If

    ' --- range check, independently per bound -------------------------------
    baseOk = modWorkbook.IsWholeInRange(rawBase, LIMIT_MIN_YEAR, LIMIT_MAX_YEAR, baseYear)
    startOk = modWorkbook.IsWholeInRange(rawStart, LIMIT_MIN_YEAR, LIMIT_MAX_YEAR, startYear)
    durationOk = modWorkbook.IsWholeInRange(rawDuration, 1, LIMIT_MAX_YEAR_COLUMNS, duration)

    If Not baseOk Then
        If modWorkbook.IsWholeNumber(rawBase) Then
            problems = problems & Fault(CHK_APPLIED_TRIPLE_CONSISTENT, _
                "applied Base Year is outside the supported calendar-year window " & _
                LIMIT_MIN_YEAR & "-" & LIMIT_MAX_YEAR & ".")
        End If
    End If
    If Not startOk Then
        If modWorkbook.IsWholeNumber(rawStart) Then
            problems = problems & Fault(CHK_APPLIED_TRIPLE_CONSISTENT, _
                "applied Start Year is outside the supported calendar-year window " & _
                LIMIT_MIN_YEAR & "-" & LIMIT_MAX_YEAR & ".")
        End If
    End If
    If Not durationOk Then
        If modWorkbook.IsWholeNumber(rawDuration) Then
            Dim rawDurationValue As Double
            If modWorkbook.TryReadDouble(rawDuration, rawDurationValue) Then
                If rawDurationValue < 1 Then
                    problems = problems & Fault(CHK_APPLIED_TRIPLE_CONSISTENT, _
                        "applied Duration is below 1.")
                Else
                    problems = problems & Fault(CHK_APPLIED_TRIPLE_CONSISTENT, _
                        "applied Duration exceeds the structural protection limit of " & _
                        LIMIT_MAX_YEAR_COLUMNS & " generated project-year columns.")
                End If
            End If
        End If
    End If

    ' --- relationships, only between values already proven bounded ----------
    If baseOk And startOk Then
        If baseYear > startYear Then
            problems = problems & Fault(CHK_APPLIED_TRIPLE_CONSISTENT, _
                "applied Base Year " & modWorkbook.SafeLong(baseYear) & _
                " is later than applied Start Year " & modWorkbook.SafeLong(startYear) & ".")
        End If
    End If
    If startOk And durationOk Then
        Dim lastYear As Long
        lastYear = modWorkbook.SafeLong(startYear + duration - 1)
        If lastYear > LIMIT_MAX_YEAR Then
            problems = problems & Fault(CHK_APPLIED_TRIPLE_CONSISTENT, _
                "applied Last Project Year " & lastYear & _
                " is beyond the supported structural year boundary " & LIMIT_MAX_YEAR & ".")
        End If
    End If

    CheckAppliedTriple = problems
End Function


' ---------------------------------------------------------------------------
Private Function CheckProfilingShape(ByVal Kind As String, ByVal TableName As String) As String
    Dim startYear As Long
    Dim expected As Long, actual As Long, i As Long
    Dim problems As String
    Dim target As ListObject

    ' Bounded reads. CheckAppliedTriple has already reported a corrupt triple; this
    ' check must still inspect the grid shape without overflowing on the same value.
    expected = modWorkbook.ReadLongInRange(NM_APPLIED_DURATION, 1, LIMIT_MAX_YEAR_COLUMNS, 0)
    startYear = modWorkbook.ReadLongInRange(NM_APPLIED_START_YEAR, LIMIT_MIN_YEAR, LIMIT_MAX_YEAR, 0)

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
            Dim headerYear As Double
            headerText = target.ListColumns(fixedCols + i).Name
            If Not modWorkbook.IsWholeInRange(headerText, LIMIT_MIN_YEAR, LIMIT_MAX_YEAR, headerYear) Then
                problems = problems & Fault(CHK_PROFILING_YEAR_HEADERS, _
                    TableName & " project-year header " & i & " is '" & headerText & _
                    "', not a supported calendar year.")
            ElseIf modWorkbook.SafeLong(headerYear) <> startYear + i - 1 Then
                problems = problems & Fault(CHK_PROFILING_YEAR_HEADERS, _
                    TableName & " project-year header " & i & " is " & headerText & _
                    "; expected " & (startYear + i - 1) & ".")
            End If
        Next i
    End If

    ' Header row intact. The column count is already covered by the year-column check
    ' above, so it is deliberately NOT reported a second time here.
    If target.HeaderRowRange Is Nothing Then
        problems = problems & Fault(CHK_GRID_SHAPE, TableName & " has lost its header row.")
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
    Dim r As Long, rowCount As Long, idCol As Long, padWidth As Long
    Dim idText As String, tail As String
    Dim problems As String

    Set register = modDrivers.RegisterTable(DriverKind)
    idCol = modDrivers.IdColumn(DriverKind)
    padWidth = modDrivers.IdPad(DriverKind)
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
                ElseIf Len(tail) < padWidth Then
                    ' Prefix plus digits is not sufficient. The contract fixes a minimum
                    ' display width, so CL-001 and CL-1000 are both valid identifiers
                    ' while CL-1 is not: it could only have been produced by something
                    ' other than the allocator.
                    problems = problems & Fault(CHK_ID_PATTERN, _
                        TableName & " row " & r & " holds '" & idText & "'; the sequence must " & _
                        "be at least " & padWidth & " digits wide (" & Prefix & _
                        String$(padWidth, "0") & " or longer).")
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
' The counter is the model's historical memory. Two independent faults:
'
'   integrity      the stored value is missing, blank, non-whole, negative or beyond
'                  the representable domain. Reported even when the register holds
'                  ZERO identifiers, because current rows cannot testify about deleted
'                  history -- that is exactly the case where a silent fallback to 0
'                  would let the next Add reissue a deleted identifier.
'   not behind     a valid counter that has fallen below an identifier it already
'                  issued would reissue that identifier next time.
'
' A counter exactly AT the representation ceiling is neither. It is valid, exhausted
' state: allocation stops, and the workbook stays structurally sound.
Private Function CheckCounters() As String
    Dim problems As String
    Dim kinds As Variant, i As Long
    kinds = Array("cost", "risk")

    For i = LBound(kinds) To UBound(kinds)
        Dim kind As String
        kind = CStr(kinds(i))

        Dim counterValue As Long
        If Not modDrivers.TryReadCounter(kind, counterValue) Then
            problems = problems & Fault(CHK_COUNTER_INTEGRITY, _
                "the " & kind & " ID counter at " & modDrivers.CounterName(kind) & _
                " is missing, blank, not a whole number, or beyond the representable " & _
                "range 0-" & ID_COUNTER_MAX & ". It records every identifier ever " & _
                "issued, including deleted ones, so allocation cannot safely continue.")
        Else
            Dim highest As Long, unrepresentable As Long
            highest = modDrivers.HighestIssued(kind, unrepresentable)
            If counterValue < highest Then
                problems = problems & Fault(CHK_COUNTERS_NOT_BEHIND, _
                    "the " & kind & " counter is " & counterValue & _
                    " but identifier number " & highest & " has already been issued; the " & _
                    "next allocation would reuse an identifier.")
            End If
            ' A counter sitting exactly at the representation ceiling is VALID state,
            ' not corruption. It means no further identifier can be ALLOCATED -- which
            ' AllocateId refuses cleanly -- and says nothing about whether the existing
            ' structure is coherent. Reporting it as a fault here would fail structural
            ' revalidation, and therefore roll back Apply and Delete, blocking every
            ' unrelated structural operation merely because the sequence is exhausted.
            If unrepresentable > 0 Then
                problems = problems & Fault(CHK_ID_PATTERN, _
                    unrepresentable & " " & kind & " identifier(s) carry a sequence beyond " & _
                    "the representable range 0-" & ID_COUNTER_MAX & " and are corrupt.")
            End If
        End If
    Next i

    CheckCounters = problems
End Function

' ---------------------------------------------------------------------------
' INVARIANT: no unkeyed structural row may hold owned data.
'
' Synchronisation rebuilds rows from their keys and clears the tail, so a row whose
' key is blank but whose other cells are not is data the next structural operation
' would silently erase. It is also invisible to every destructive assessment,
' because those are keyed too.
Private Function CheckOrphanRows() As String
    Dim problems As String
    problems = problems & OrphanFault(modDrivers.RegisterTable("cost"), _
                                      modDrivers.IdColumn("cost"), TBL_COST_LINES)
    problems = problems & OrphanFault(modDrivers.RegisterTable("risk"), _
                                      modDrivers.IdColumn("risk"), TBL_RISK_REGISTER)
    problems = problems & OrphanFault(modProfiling.ProfilingTable(modProfiling.CostKind()), _
                                      1, TBL_COST_PROFILING)
    problems = problems & OrphanFault(modProfiling.ProfilingTable(modProfiling.RiskKind()), _
                                      1, TBL_RISK_PROFILING)
    problems = problems & OrphanFault(modWorkbook.Lo(SH_INFLATION, TBL_INFLATION), _
                                      1, TBL_INFLATION)
    CheckOrphanRows = problems
End Function

Private Function OrphanFault(ByVal Target As ListObject, ByVal KeyColumn As Long, _
                             ByVal TableName As String) As String
    Dim rows() As Long, count As Long
    count = modWorkbook.OrphanRows(Target, KeyColumn, rows)
    If count = 0 Then Exit Function
    OrphanFault = Fault(CHK_NO_ORPHAN_STRUCTURAL_DATA, _
                        modWorkbook.DescribeOrphans(TableName, rows, count))
End Function

' ---------------------------------------------------------------------------
' The focused PRE-MUTATION safety gate.
'
' Deliberately NOT a full ValidateStructure: running that before Apply would block
' the intended "a Config profile was removed, Apply will synchronise it away"
' workflow, which is a legitimate operation the destructive prompt already covers.
'
' This targets only corruption that a structural operation would SILENTLY ERASE,
' which no confirmation could ever have warned about because it is unkeyed.
Public Function PreMutationCheck() As String
    Dim problems As String
    problems = CheckOrphanRows()
    If Len(problems) > 0 Then
        PreMutationCheck = "Unkeyed structural data was found. The operation was " & _
                           "refused because synchronisation would have deleted it " & _
                           "without warning:" & vbCrLf & problems & _
                           "Give each of those rows its key, or clear them, then try again."
    End If
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
            Dim headerYear As Double
            headerText = target.ListColumns(fixedCols + i).Name
            If Not modWorkbook.IsWholeInRange(headerText, LIMIT_MIN_YEAR, LIMIT_MAX_YEAR, headerYear) Then
                problems = problems & Fault(CHK_INFLATION_YEAR_HEADERS, _
                    TBL_INFLATION & " year header " & i & " is '" & headerText & _
                    "', not a supported calendar year.")
            ElseIf modWorkbook.SafeLong(headerYear) <> CLng(firstYear) + i - 1 Then
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
