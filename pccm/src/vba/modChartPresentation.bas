Attribute VB_Name = "modChartPresentation"
Option Explicit

' ==========================================================================
' THE DASHBOARD CHART PRESENTATION OWNER
'
' WHAT IT IS FOR. The two year-based Dashboard charts - the cumulative cost
' profile and the annual cash flow - must plot the years a run actually
' produced, not the two hundred rows the annual bridge reserves. Their VALUES
' come from workbook-scoped names that cut the bridge at the stamped year
' count, and real Excel keeps those. Their CATEGORIES could not be bound the
' same way: three Windows runs proved that a name in a series' category slot
' comes back blank, whatever the markup around it, and the workbook then
' numbers the axis 1, 2, 3. Excel binds a category from a RANGE, at runtime,
' through Series.XValues - which Windows also proved directly, by assigning
' one and reading the SERIES formula back. That is what this module does.
'
' WHAT IT IS NOT. It computes nothing. It reads no result, derives no state,
' owns no year count and publishes nothing. It resizes a range that already
' exists to a length that is already published, and assigns it to the charts
' that already plot that data. Every number it uses comes from the workbook:
'
'   CHART_CATEGORY_WINDOW  the reserved calendar-year column of the annual
'                          chart bridge, as the builder defined it
'   CHART_YEAR_COUNT       the Results annual Years Covered cell - the SAME
'                          cell the bridge's own guard reads, so there is one
'                          year-count authority and this is not a second one
'
' WHICH SERIES IT TOUCHES, AND HOW IT KNOWS. Not by chart title and not by
' position: a series is a year-chart series when its VALUES are one of the
' applied-year names, which is a fact about the series itself. The histogram
' and the tornado plot ordinary ranges, carry no such name, and are never
' touched here - their categories are static by contract and Windows has
' proved that wiring since P8-3.
'
' NOTHING IS FABRICATED. With no published annual result the year count is
' blank, and the categories are bound to the FIRST reserved row alone. The
' bridge still answers NA() for that row's value and a blank for its label, so
' the chart draws no point and prints no caption - which is what an empty
' Dashboard should look like, and is what it looked like before.
'
' PROTECTION. Assigning XValues is a change to a protected sheet, and Windows
' proved an external client cannot make it while the Dashboard is protected.
' This does not invent a way around that: it opens the ACCEPTED structural
' window through modProtection, exactly as every other structural operation in
' this workbook does, and closes it again on every path including failure.
' Nothing here protects, unprotects or touches workbook structure itself.
' ==========================================================================

' The three names the builder writes for this owner, and the one prefix that
' says a series plots the applied years. They are the manifest's, projected
' into the workbook as defined names; a control asserts these spellings against
' the manifest that produces them.
Private Const CHART_CATEGORY_WINDOW As String = "chartAnnual_category_window"
Private Const CHART_YEAR_COUNT As String = "chartAnnual_year_count"
Private Const CHART_APPLIED_PREFIX As String = "chartAnnual_"

' ==========================================================================
' THE ONE ENTRY POINT. Called at the presentation boundaries where the annual
' publication can have changed - the workbook opening, an annual run that
' published, and a reset that cleared - and nowhere inside a calculation or a
' simulation, which produce no chart and would pay for this on every iteration.
'
' IT REPORTS, IT DOES NOT ANNOUNCE. A presentation binding that could not be
' applied is not a failed command: the caller decides what to do with it, and
' every current caller records it rather than interrupting the person.
' ==========================================================================
Public Function ChartPresentationApplyCategories(ByRef detail As String) As Boolean
    Dim categories As Range
    Dim windowOpen As Boolean
    Dim windowDetail As String
    Dim bound As Long

    detail = vbNullString
    On Error GoTo Failed

    If Not CategoryRange(categories, detail) Then Exit Function

    ' THE ACCEPTED STRUCTURAL WINDOW, AND NOTHING ELSE. A chart on a protected
    ' sheet is not writable; this is the one mechanism this workbook has for
    ' that, and it is closed again below whatever happens.
    If Not modProtection.ProtectionBeginStructural(windowDetail) Then
        detail = "the chart categories could not be bound: " & windowDetail
        Exit Function
    End If
    windowOpen = True

    bound = BindAppliedSeries(categories, detail)
    If bound < 0 Then GoTo Failed

    windowOpen = False
    If Not modProtection.ProtectionEndStructural(windowDetail) Then
        detail = "the chart categories were bound but protection was not restored: " & _
                 windowDetail
        Exit Function
    End If

    ChartPresentationApplyCategories = True
    Exit Function

Failed:
    If Len(detail) = 0 Then detail = "the chart categories could not be bound: " & Err.Description
    ' THE WINDOW IS CLOSED ON THE FAILURE PATH TOO. A half-open window leaves
    ' the workbook unprotected with nobody left to protect it, which is worse
    ' than the binding this function came here to apply.
    If windowOpen Then
        If Not modProtection.ProtectionEndStructural(windowDetail) Then
            detail = detail & " (and protection was not restored: " & windowDetail & ")"
        End If
    End If
End Function

' ==========================================================================
' THE RANGE THE CATEGORIES ARE BOUND TO: the reserved calendar-year column,
' resized to the published year count.
'
' THE BOUNDS ARE THE WINDOW'S OWN. A count larger than the reserved column is
' cut to it - the rows beyond it were never written and hold nothing to label -
' and a count that is blank, zero, negative or not a number leaves exactly one
' row, whose label the bridge already answers blank. No year is invented and no
' point is created either way.
' ==========================================================================
Private Function CategoryRange(ByRef categories As Range, ByRef detail As String) As Boolean
    Dim reserved As Range
    Dim countCell As Range
    Dim reported As Variant
    Dim years As Long

    On Error GoTo Failed
    If Not NamedRange(CHART_CATEGORY_WINDOW, reserved, detail) Then Exit Function
    If Not NamedRange(CHART_YEAR_COUNT, countCell, detail) Then Exit Function

    years = 1
    reported = countCell.Cells(1, 1).Value2
    If IsNumeric(reported) Then
        If CDbl(reported) >= 1 Then
            If CDbl(reported) >= CDbl(reserved.Rows.Count) Then
                years = reserved.Rows.Count
            Else
                years = CLng(Int(CDbl(reported)))
            End If
        End If
    End If

    Set categories = reserved.Resize(years, 1)
    CategoryRange = True
    Exit Function
Failed:
    detail = "the chart category range could not be read: " & Err.Description
End Function

Private Function NamedRange(ByVal DefinedName As String, ByRef Target As Range, _
                            ByRef detail As String) As Boolean
    On Error GoTo Failed
    Set Target = ThisWorkbook.Names(DefinedName).RefersToRange
    NamedRange = True
    Exit Function
Failed:
    detail = "the defined name " & DefinedName & " does not resolve to a range"
End Function

' ==========================================================================
' EVERY SERIES THAT PLOTS THE APPLIED YEARS, BOUND TO THE SAME CATEGORIES.
' Returns how many were bound, or -1 with a detail if one refused.
'
' A SERIES IS IDENTIFIED BY WHAT IT PLOTS. Its values formula names one of the
' applied-year names or it does not; no title is read, no position is assumed,
' and a chart that plots ordinary ranges is passed over untouched.
' ==========================================================================
Private Function BindAppliedSeries(ByVal categories As Range, ByRef detail As String) As Long
    Dim sheet As Worksheet
    Dim plot As ChartObject
    Dim item As Series
    Dim bound As Long

    On Error GoTo Failed
    Set sheet = ThisWorkbook.Worksheets(SH_DASHBOARD)
    For Each plot In sheet.ChartObjects
        For Each item In plot.Chart.SeriesCollection
            If PlotsAppliedYears(item) Then
                item.XValues = categories
                bound = bound + 1
            End If
        Next item
    Next plot
    BindAppliedSeries = bound
    Exit Function
Failed:
    detail = "a Dashboard chart refused its category binding: " & Err.Description
    BindAppliedSeries = -1
End Function

Private Function PlotsAppliedYears(ByVal item As Series) As Boolean
    ' A SERIES THAT CANNOT EVEN BE ASKED WHAT IT PLOTS IS NOT ONE OF OURS, and
    ' saying so is not the same as ignoring the error: the answer is False, the
    ' series is left alone, and the caller's own handler still owns every
    ' failure that matters.
    Dim formula As String
    On Error GoTo Unreadable
    formula = item.formula
    On Error GoTo 0
    PlotsAppliedYears = (InStr(1, formula, CHART_APPLIED_PREFIX, vbBinaryCompare) > 0)
    Exit Function
Unreadable:
    PlotsAppliedYears = False
End Function
