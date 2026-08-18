Attribute VB_Name = "modCalcResolve"
Option Explicit

' ==========================================================================
' modCalcResolve - the worksheet-aware resolution layer, and the ONLY one.
'
' It reads the workbook and hands back plain typed data. After it returns, the
' numerical kernel needs no workbook at all:
'
'   modCalcResolve      worksheet access: YES
'   modCalcFactors      worksheet access: no
'   modCalcAnalytical   worksheet access: no
'   modCalcFingerprint  worksheet access: no
'
' No Range, Worksheet, ListObject or Object crosses this boundary, and a
' workbook ROW NUMBER is never a driver identity - the Permanent ID is.
'
' --------------------------------------------------------------------------
' WHAT THIS MODULE DOES NOT DO
' --------------------------------------------------------------------------
' It computes no distribution mean, no deterministic central value, no
' inflation or discount compounding of its own, no Knom or Kpv, no A/B/C/D/E,
' no fingerprint and no reconciliation. Where a resolved factor must be
' materialised it CALLS the accepted function in modCalcFactors. A second
' implementation of an accepted formula is the defect this rule exists to
' prevent.
'
' It also does not write anything, anywhere. There is no _Calc write-back in
' this step and no user-facing message: status and detail are returned across
' the boundary and later orchestration owns what a user is told.
'
' --------------------------------------------------------------------------
' THE ORDERING RULE
' --------------------------------------------------------------------------
' The reference sets are built FIRST, from the identified drivers, and only
' then are FX and inflation consulted. That order IS the referenced-only rule:
' a Config assumption for a currency or a profile nobody uses cannot block a
' valid model, because resolution never asks about it.
'
' --------------------------------------------------------------------------
' IDENTIFIERS ARE EXACT
' --------------------------------------------------------------------------
' A Permanent ID, a Currency, an Inflation Profile and a Distribution are used
' EXACTLY as entered. " USD " is not "USD". Nothing here trims a key into
' another key, case-folds one to make a lookup succeed, repairs whitespace or
' substitutes a default. If the exact key does not resolve that is a controlled
' failure, not permission to rewrite what the user typed.
'
' The STRUCTURAL validity of a Permanent ID - the CL-/R- prefixes, the pattern,
' the counter rules - remains Phase 4's and is not re-checked here. Resolution
' needs only enough text semantics to look up, order and reference.
' ==========================================================================

' Which register a driver came from. Kind is carried explicitly rather than
' inferred from the ID prefix, which is Phase-4 business the resolver does not
' reopen.
Private Const KIND_COST As Long = 1
Private Const KIND_RISK As Long = 2

' --------------------------------------------------------------------------
' The plain carriers handed to the numerical layer
' --------------------------------------------------------------------------
Public Type ResolvedTimeline
    BaseYear As Long
    StartYear As Long
    Duration As Long
    LastYear As Long
    DiscountRate As Double
End Type

' One driver, as read. These are RESOLVED values, not yet CHECKED values:
' modCalcCheck is the next step and owns the numerical prerequisites. What is
' guaranteed here is that every field could be read as the type the model needs.
Public Type ResolvedDriver
    PermanentId As String
    IsRisk As Boolean
    Currency As String
    InflationProfile As String
    Distribution As String
    DistKind As Long
    Quantity As Double
    Probability As Double
    MinValue As Double
    MostLikely As Double
    MaxValue As Double
    HasMostLikely As Boolean
End Type

' The whole resolved model. Every array is plain and typed; the parallel arrays
' are indexed together and their lengths are carried explicitly, because a VBA
' array cannot represent a zero-element set and the empty model is valid.
Public Type ResolvedModel
    Timeline As ResolvedTimeline
    Drivers() As ResolvedDriver
    DriverCount As Long
    Currencies() As String
    CurrencyRates() As Double
    CurrencyCount As Long
    Profiles() As String
    ProfileCount As Long
    ' InflationRates(profileIndex, yearOffset) over BaseYear+1 .. LastYear.
    InflationRates() As Double
    RequiredYearCount As Long
    ' Weights(driverIndex, projectYearOffset), in project-year order.
    Weights() As Double
    ' The applied project-year / calendar-year mapping.
    ProjectIndexes() As Long
    CalendarYears() As Long
    ' Per-driver resolved FX, parallel to Drivers.
    DriverFxRates() As Double
End Type

' ==========================================================================
' The entry point. The ORDER OF THE CALLS BELOW IS THE SEMANTIC RULE.
' ==========================================================================
Public Function ResolveModel(ByRef model As ResolvedModel, ByRef detail As String) As Boolean
    detail = vbNullString
    If Not ResolveAppliedTimeline(model.Timeline, detail) Then Exit Function
    If Not ResolveProjectYears(model.Timeline, model.ProjectIndexes, _
                               model.CalendarYears, detail) Then Exit Function

    ' STEP 1 - identify the drivers. Nothing about FX or inflation is consulted
    ' until this has succeeded.
    If Not ResolveDrivers(model.Drivers, model.DriverCount, detail) Then Exit Function

    ' STEP 2 - derive the reference sets FROM those drivers.
    If Not ReferencedCurrencies(model.Drivers, model.DriverCount, _
                                model.Currencies, model.CurrencyCount, detail) Then Exit Function
    If Not ReferencedProfiles(model.Drivers, model.DriverCount, _
                              model.Profiles, model.ProfileCount, detail) Then Exit Function

    ' STEP 3 - and only now look at the assumptions, for those references only.
    If Not ResolveFxRates(model.Currencies, model.CurrencyCount, _
                          model.CurrencyRates, detail) Then Exit Function
    If Not ResolveInflationRates(model.Profiles, model.ProfileCount, model.Timeline, _
                                 model.InflationRates, model.RequiredYearCount, _
                                 detail) Then Exit Function

    If Not AttachDriverFx(model, detail) Then Exit Function
    If Not ResolveProfileWeights(model.Drivers, model.DriverCount, model.Timeline, _
                                 model.Weights, detail) Then Exit Function
    ResolveModel = True
End Function

' ==========================================================================
' Applied structural state
'
' THE APPLIED VALUES, NEVER THE ENTERED ONES. An entered Duration that has not
' been applied has not generated its project-year columns, so calculating from
' it would calculate against a shape the workbook does not have.
' ==========================================================================
Public Function ResolveAppliedTimeline(ByRef timeline As ResolvedTimeline, _
                                       ByRef detail As String) As Boolean
    Dim value As Double
    If Not AppliedWholeNumber(NM_APPLIED_BASE_YEAR, LIMIT_MIN_YEAR, LIMIT_MAX_YEAR, _
                              value, "applied Base Year", detail) Then Exit Function
    timeline.BaseYear = modWorkbook.SafeLong(value)
    If Not AppliedWholeNumber(NM_APPLIED_START_YEAR, LIMIT_MIN_YEAR, LIMIT_MAX_YEAR, _
                              value, "applied Start Year", detail) Then Exit Function
    timeline.StartYear = modWorkbook.SafeLong(value)
    If Not AppliedWholeNumber(NM_APPLIED_DURATION, 1, LIMIT_MAX_YEAR_COLUMNS, _
                              value, "applied Duration", detail) Then Exit Function
    timeline.Duration = modWorkbook.SafeLong(value)
    If Not AppliedWholeNumber(NM_APPLIED_LAST_YEAR, LIMIT_MIN_YEAR, LIMIT_MAX_YEAR, _
                              value, "applied Last Project Year", detail) Then Exit Function
    timeline.LastYear = modWorkbook.SafeLong(value)

    ' The discount rate is an ordinary Setup input, required and not defaulted.
    ' A blank one is an unmade assumption, not zero.
    If Not NumericNamedCell(NM_INPUT_DISCOUNT_RATE, timeline.DiscountRate, _
                            "Discount Rate", detail) Then Exit Function
    ResolveAppliedTimeline = True
End Function

Public Function ResolveProjectYears(ByRef timeline As ResolvedTimeline, _
                                    ByRef projectIndexes() As Long, _
                                    ByRef calendarYears() As Long, _
                                    ByRef detail As String) As Boolean
    ' The applied project-year / calendar-year mapping. Project year 1 is the
    ' Start Year; this is structure, not calculation.
    Dim index As Long
    If timeline.Duration < 1 Then
        detail = "applied Duration is not a project-year span"
        Exit Function
    End If
    ReDim projectIndexes(0 To timeline.Duration - 1)
    ReDim calendarYears(0 To timeline.Duration - 1)
    For index = 0 To timeline.Duration - 1
        projectIndexes(index) = index + 1
        calendarYears(index) = timeline.StartYear + index
    Next index
    ResolveProjectYears = True
End Function

' ==========================================================================
' Driver identification
' ==========================================================================
Public Function ResolveDrivers(ByRef drivers() As ResolvedDriver, ByRef driverCount As Long, _
                               ByRef detail As String) As Boolean
    ' AN EMPTY DRIVER SET IS VALID. A workbook with no Cost Lines and no Risks
    ' resolves to zero drivers and an empty reference set; no minimum-driver
    ' rule is invented here, because no accepted contract states one.
    Dim capacity As Long
    driverCount = 0
    capacity = RegisterRowCapacity(REG_COST_LINES_SHEET, TBL_COST_LINES) + _
               RegisterRowCapacity(REG_RISK_REGISTER_SHEET, TBL_RISK_REGISTER)
    If capacity < 1 Then
        ResolveDrivers = True
        Exit Function
    End If
    ReDim drivers(0 To capacity - 1)
    If Not ReadRegister(KIND_COST, drivers, driverCount, detail) Then Exit Function
    If Not ReadRegister(KIND_RISK, drivers, driverCount, detail) Then Exit Function
    ResolveDrivers = True
End Function

Private Function RegisterRowCapacity(ByVal sheetName As String, _
                                     ByVal tableName As String) As Long
    If Not modWorkbook.LoExists(sheetName, tableName) Then Exit Function
    RegisterRowCapacity = modWorkbook.BodyRowCount(modWorkbook.Lo(sheetName, tableName))
End Function

Private Function ReadRegister(ByVal kind As Long, ByRef drivers() As ResolvedDriver, _
                              ByRef driverCount As Long, ByRef detail As String) As Boolean
    Dim table As ListObject, rowIndex As Long, rows As Long
    Dim sheetName As String, tableName As String, label As String
    Dim idColumn As Long, slot As Long

    If kind = KIND_COST Then
        sheetName = REG_COST_LINES_SHEET: tableName = TBL_COST_LINES
        idColumn = COL_COST_LINES_COST_LINE_ID: label = "cost line"
    Else
        sheetName = REG_RISK_REGISTER_SHEET: tableName = TBL_RISK_REGISTER
        idColumn = COL_RISK_REGISTER_RISK_ID: label = "risk"
    End If
    If Not modWorkbook.LoExists(sheetName, tableName) Then
        detail = "register " & tableName & " is missing"
        Exit Function
    End If
    Set table = modWorkbook.Lo(sheetName, tableName)
    rows = modWorkbook.BodyRowCount(table)

    For rowIndex = 1 To rows
        ' A row is a driver when its key column carries something. Row PRESENCE
        ' is Phase 4's own definition and is reused rather than restated; the
        ' key ITSELF is then read raw, because presence and identity are
        ' different questions and only one of them may trim.
        If Len(modWorkbook.TextOf(modWorkbook.CellIn(table, rowIndex, idColumn))) > 0 Then
            slot = driverCount
            If Not ReadDriverRow(kind, table, rowIndex, drivers(slot), label, detail) Then
                Exit Function
            End If
            driverCount = driverCount + 1
        End If
    Next rowIndex
    ReadRegister = True
End Function

Private Function ReadDriverRow(ByVal kind As Long, ByVal table As ListObject, _
                               ByVal rowIndex As Long, ByRef driver As ResolvedDriver, _
                               ByVal label As String, ByRef detail As String) As Boolean
    Dim where As String
    Dim idColumn As Long, currencyColumn As Long, profileColumn As Long
    Dim distributionColumn As Long, scalarColumn As Long
    Dim minColumn As Long, likelyColumn As Long, maxColumn As Long

    If kind = KIND_COST Then
        idColumn = COL_COST_LINES_COST_LINE_ID
        currencyColumn = COL_COST_LINES_CURRENCY
        profileColumn = COL_COST_LINES_INFLATION_PROFILE
        distributionColumn = COL_COST_LINES_DISTRIBUTION
        scalarColumn = COL_COST_LINES_QUANTITY
        minColumn = COL_COST_LINES_UNIT_COST_MIN
        likelyColumn = COL_COST_LINES_UNIT_COST_MOST_LIKELY
        maxColumn = COL_COST_LINES_UNIT_COST_MAX
    Else
        idColumn = COL_RISK_REGISTER_RISK_ID
        currencyColumn = COL_RISK_REGISTER_CURRENCY
        profileColumn = COL_RISK_REGISTER_INFLATION_PROFILE
        distributionColumn = COL_RISK_REGISTER_DISTRIBUTION
        scalarColumn = COL_RISK_REGISTER_PROBABILITY
        minColumn = COL_RISK_REGISTER_IMPACT_MIN
        likelyColumn = COL_RISK_REGISTER_IMPACT_MOST_LIKELY
        maxColumn = COL_RISK_REGISTER_IMPACT_MAX
    End If

    driver.IsRisk = (kind = KIND_RISK)
    where = label & " row " & CStr(rowIndex)
    If Not ExactIdentifier(table, rowIndex, idColumn, driver.PermanentId, _
                           where & ": Permanent ID", detail) Then Exit Function
    where = label & " " & driver.PermanentId
    If Not ExactIdentifier(table, rowIndex, currencyColumn, driver.Currency, _
                           where & ": Currency", detail) Then Exit Function
    If Not ExactIdentifier(table, rowIndex, profileColumn, driver.InflationProfile, _
                           where & ": Inflation Profile", detail) Then Exit Function
    If Not ExactIdentifier(table, rowIndex, distributionColumn, driver.Distribution, _
                           where & ": Distribution", detail) Then Exit Function
    driver.DistKind = DistributionKindOf(driver.Distribution)
    If driver.DistKind = 0 Then
        detail = where & ": Distribution " & driver.Distribution & " is not an accepted distribution"
        Exit Function
    End If

    ' Quantity belongs to a cost line and Probability to a risk. The one that
    ' does not apply is left unset rather than given a 1 that would read as a
    ' real entry; the numerical carry convention is applied later, by the layer
    ' that owns it.
    If kind = KIND_COST Then
        If Not NumericCell(table, rowIndex, scalarColumn, driver.Quantity, _
                           where & ": Quantity", detail) Then Exit Function
    Else
        If Not NumericCell(table, rowIndex, scalarColumn, driver.Probability, _
                           where & ": Probability", detail) Then Exit Function
    End If
    If Not NumericCell(table, rowIndex, minColumn, driver.MinValue, _
                       where & ": Min", detail) Then Exit Function
    If Not NumericCell(table, rowIndex, maxColumn, driver.MaxValue, _
                       where & ": Max", detail) Then Exit Function

    ' D1: Uniform is a two-point distribution. A populated Most Likely is
    ' accepted and IGNORED - the cell may hold a leftover from another choice of
    ' distribution, and refusing it would block a valid model.
    driver.HasMostLikely = (driver.DistKind <> DIST_UNIFORM)
    If driver.HasMostLikely Then
        If Not NumericCell(table, rowIndex, likelyColumn, driver.MostLikely, _
                           where & ": Most Likely", detail) Then Exit Function
    End If
    ReadDriverRow = True
End Function

Private Function DistributionKindOf(ByVal name As String) As Long
    ' ADAPTER, NOT AUTHORITY. The master list of distribution names is owned by
    ' the input contract and reaches VBA as the projected DISTRIBUTION_NAME_*
    ' constants; this says only which internal shape each accepted name selects.
    ' A name absent from the list is refused, never mapped to a default.
    Select Case name
    Case DISTRIBUTION_NAME_1
        DistributionKindOf = DIST_TRIANGULAR
    Case DISTRIBUTION_NAME_2
        DistributionKindOf = DIST_BETA_PERT
    Case DISTRIBUTION_NAME_3
        DistributionKindOf = DIST_UNIFORM
    End Select
End Function

' ==========================================================================
' Reference sets - built from the DRIVERS, before any assumption is consulted
' ==========================================================================
Public Function ReferencedCurrencies(ByRef drivers() As ResolvedDriver, _
                                     ByVal driverCount As Long, ByRef names() As String, _
                                     ByRef nameCount As Long, ByRef detail As String) As Boolean
    Dim raw() As String, index As Long
    detail = vbNullString
    nameCount = 0
    If driverCount < 0 Then Exit Function
    If driverCount = 0 Then
        ReferencedCurrencies = True
        Exit Function
    End If
    ReDim raw(0 To driverCount - 1)
    For index = 0 To driverCount - 1
        raw(index) = drivers(LBound(drivers) + index).Currency
    Next index
    ReferencedCurrencies = DistinctSorted(raw, driverCount, names, nameCount)
End Function

Public Function ReferencedProfiles(ByRef drivers() As ResolvedDriver, _
                                   ByVal driverCount As Long, ByRef names() As String, _
                                   ByRef nameCount As Long, ByRef detail As String) As Boolean
    Dim raw() As String, index As Long
    detail = vbNullString
    nameCount = 0
    If driverCount < 0 Then Exit Function
    If driverCount = 0 Then
        ReferencedProfiles = True
        Exit Function
    End If
    ReDim raw(0 To driverCount - 1)
    For index = 0 To driverCount - 1
        raw(index) = drivers(LBound(drivers) + index).InflationProfile
    Next index
    ReferencedProfiles = DistinctSorted(raw, driverCount, names, nameCount)
End Function

Private Function DistinctSorted(ByRef raw() As String, ByVal count As Long, _
                                ByRef names() As String, ByRef nameCount As Long) As Boolean
    ' Distinct, then ascending on UTF-16 code units. The order is OBSERVABLE -
    ' it decides the order of the later audit rows - so it is pinned by the same
    ' binary comparison the fingerprint and the canonical driver order use,
    ' rather than left to whichever driver happened to be first in the sheet.
    Dim index As Long, probe As Long, moving As String, seen As Boolean
    nameCount = 0
    If count < 1 Then Exit Function
    ReDim names(0 To count - 1)
    For index = 0 To count - 1
        seen = False
        For probe = 0 To nameCount - 1
            If StrComp(names(probe), raw(index), vbBinaryCompare) = 0 Then
                seen = True
                Exit For
            End If
        Next probe
        If Not seen Then
            names(nameCount) = raw(index)
            nameCount = nameCount + 1
        End If
    Next index
    For index = 1 To nameCount - 1
        moving = names(index)
        probe = index - 1
        Do While probe >= 0
            If StrComp(names(probe), moving, vbBinaryCompare) <= 0 Then Exit Do
            names(probe + 1) = names(probe)
            probe = probe - 1
        Loop
        names(probe + 1) = moving
    Next index
    DistinctSorted = True
End Function

' ==========================================================================
' FX - referenced currencies only, after a global reporting-currency invariant
'
' TWO DISTINCT QUESTIONS, deliberately not conflated:
'
'   Is the reporting currency sound?  A GLOBAL INVARIANT. It must appear
'     exactly once and must equal 1, in every model - including one that
'     references no currency at all and one with no drivers.
'   Which currencies does this model resolve?  Only those a driver references.
'
' Being validated globally does not make a currency referenced. The resolved
' set is NOT seeded with the reporting currency: a USD-only model resolves USD
' and nothing else, and an empty model resolves nothing.
' ==========================================================================
Public Function ResolveFxRates(ByRef names() As String, ByVal nameCount As Long, _
                               ByRef rates() As Double, ByRef detail As String) As Boolean
    Dim table As ListObject, index As Long, matches As Long, row As Long
    Dim rate As Double, key As String
    detail = vbNullString
    If nameCount < 0 Then Exit Function
    If Not modWorkbook.LoExists(SH_SETUP, TBL_FX_RATES) Then
        detail = "the FX table " & TBL_FX_RATES & " is missing"
        Exit Function
    End If
    Set table = modWorkbook.Lo(SH_SETUP, TBL_FX_RATES)

    If Not ReportingCurrencyInvariant(table, detail) Then Exit Function
    If nameCount = 0 Then
        ResolveFxRates = True
        Exit Function
    End If

    ReDim rates(0 To nameCount - 1)
    For index = 0 To nameCount - 1
        key = names(LBound(names) + index)
        If StrComp(key, REPORTING_CURRENCY, vbBinaryCompare) = 0 Then
            ' Already proven above, and proven to be exactly the identity rate.
            rates(index) = REPORTING_CURRENCY_RATE
        Else
            matches = MatchingFxRows(table, key, row)
            If matches <> 1 Then
                detail = "FX: referenced currency " & key & " matches " & _
                         CStr(matches) & " rows; exactly one rate must resolve"
                Exit Function
            End If
            If Not NumericCell(table, row, COL_FX_RATES_FX_TO_SAR, rate, _
                               "FX rate for referenced currency " & key, detail) Then Exit Function
            If rate <= 0# Then
                detail = "FX: referenced currency " & key & " has a rate that is not " & _
                         "strictly positive"
                Exit Function
            End If
            rates(index) = rate
        End If
    Next index
    ResolveFxRates = True
End Function

Private Function ReportingCurrencyInvariant(ByVal table As ListObject, _
                                            ByRef detail As String) As Boolean
    Dim matches As Long, row As Long, rate As Double
    matches = MatchingFxRows(table, REPORTING_CURRENCY, row)
    If matches <> 1 Then
        detail = "FX: the reporting currency " & REPORTING_CURRENCY & " must appear " & _
                 "exactly once, found " & CStr(matches) & _
                 ". This is a global invariant and applies whether or not any driver " & _
                 "references it."
        Exit Function
    End If
    If Not NumericCell(table, row, COL_FX_RATES_FX_TO_SAR, rate, _
                       "FX rate for " & REPORTING_CURRENCY, detail) Then Exit Function
    If rate <> REPORTING_CURRENCY_RATE Then
        detail = "FX: the reporting currency " & REPORTING_CURRENCY & " must resolve to " & _
                 "exactly " & CStr(REPORTING_CURRENCY_RATE)
        Exit Function
    End If
    ReportingCurrencyInvariant = True
End Function

Private Function MatchingFxRows(ByVal table As ListObject, ByVal key As String, _
                                ByRef firstMatch As Long) As Long
    ' Counts EVERY match, so a duplicate is reported rather than resolved by
    ' taking the first row. An unreferenced currency is never asked about, so a
    ' bad row for one cannot block the model.
    Dim rowIndex As Long, found As Long, text As String
    firstMatch = 0
    For rowIndex = 1 To modWorkbook.BodyRowCount(table)
        If RawCellText(table, rowIndex, COL_FX_RATES_CURRENCY, text) Then
            If StrComp(text, key, vbBinaryCompare) = 0 Then
                found = found + 1
                If firstMatch = 0 Then firstMatch = rowIndex
            End If
        End If
    Next rowIndex
    MatchingFxRows = found
End Function

' ==========================================================================
' Inflation - referenced profiles only, over BaseYear+1 .. LastProjectYear
'
' CALENDAR-YEAR ANCHORED. The required span is selected by calendar year, so a
' Start Year shift selects the rates for the new years rather than moving the
' old values positionally. A missing required year is refused and is never
' treated as zero: the grid is seeded blank precisely so an assumption the user
' never made cannot be fabricated as 0%.
'
' An incomplete or invalid UNREFERENCED profile is never consulted.
' ==========================================================================
Public Function ResolveInflationRates(ByRef names() As String, ByVal nameCount As Long, _
                                      ByRef timeline As ResolvedTimeline, _
                                      ByRef rates() As Double, ByRef yearCount As Long, _
                                      ByRef detail As String) As Boolean
    Dim table As ListObject, index As Long, offset As Long
    Dim row As Long, column As Long, year As Long, rate As Double, key As String
    detail = vbNullString
    yearCount = timeline.LastYear - timeline.BaseYear
    If yearCount < 0 Then
        detail = "applied Last Project Year precedes the applied Base Year"
        Exit Function
    End If
    If nameCount < 0 Then Exit Function
    If nameCount = 0 Or yearCount = 0 Then
        ' No referenced profile, or Base Year = Last Year so no annual rate is
        ' required at all. Both are legitimate.
        ResolveInflationRates = True
        Exit Function
    End If
    If Not modWorkbook.LoExists(GRID_INFLATION_SHEET, TBL_INFLATION) Then
        detail = "the inflation grid " & TBL_INFLATION & " is missing"
        Exit Function
    End If
    Set table = modWorkbook.Lo(GRID_INFLATION_SHEET, TBL_INFLATION)

    ReDim rates(0 To nameCount - 1, 0 To yearCount - 1)
    For index = 0 To nameCount - 1
        key = names(LBound(names) + index)
        row = MatchingGridRow(table, GCOL_INFLATION_PROFILE_NAME, key)
        If row = 0 Then
            detail = "inflation: profile " & key & " is referenced by at least one driver " & _
                     "but is not present in the inflation table"
            Exit Function
        End If
        For offset = 0 To yearCount - 1
            year = timeline.BaseYear + 1 + offset
            column = YearColumn(table, GRID_INFLATION_FIXED_COLS, year)
            If column = 0 Then
                detail = "inflation profile " & key & ", calendar year " & CStr(year) & _
                         ": the required year column is missing"
                Exit Function
            End If
            If Not NumericCell(table, row, column, rate, _
                               "inflation profile " & key & ", calendar year " & CStr(year), _
                               detail) Then Exit Function
            rates(index, offset) = rate
        Next offset
    Next index
    ResolveInflationRates = True
End Function

Public Function ResolveInflationFactors(ByRef rates() As Double, ByVal profileIndex As Long, _
                                        ByVal yearCount As Long, _
                                        ByRef timeline As ResolvedTimeline, _
                                        ByRef factors() As Double, _
                                        ByRef detail As String) As Boolean
    ' The cumulative factors for one referenced profile, produced by the
    ' ACCEPTED function in modCalcFactors. No compounding happens here: a second
    ' implementation of an accepted formula is exactly what this layer must not
    ' contain.
    Dim series() As Double, offset As Long
    If yearCount > 0 Then
        ReDim series(0 To yearCount - 1)
        For offset = 0 To yearCount - 1
            series(offset) = rates(profileIndex, offset)
        Next offset
    Else
        ReDim series(0 To 0)
    End If
    ResolveInflationFactors = modCalcFactors.BuildInflationFactors( _
        timeline.BaseYear, timeline.LastYear, series, factors, detail)
End Function

' ==========================================================================
' Profiling - BY PERMANENT ID, never by row position
'
' A driver reorder must not attach another driver's weights, so the grid row is
' found by matching the Permanent ID and never by walking the two tables in
' parallel. Only the project-year columns belonging to the APPLIED timeline are
' read.
'
' A numeric 0 weight is legitimate - a driver may genuinely spend nothing in a
' year. A BLANK weight is not zero: it is an unmade assumption, and it is
' refused rather than fabricated.
' ==========================================================================
Public Function ResolveProfileWeights(ByRef drivers() As ResolvedDriver, _
                                      ByVal driverCount As Long, _
                                      ByRef timeline As ResolvedTimeline, _
                                      ByRef weights() As Double, _
                                      ByRef detail As String) As Boolean
    Dim costGrid As ListObject, riskGrid As ListObject, grid As ListObject
    Dim index As Long, offset As Long, row As Long, column As Long
    Dim keyColumn As Long, fixedCols As Long, weight As Double
    detail = vbNullString
    If driverCount < 0 Then Exit Function
    If driverCount = 0 Then
        ResolveProfileWeights = True
        Exit Function
    End If
    If timeline.Duration < 1 Then
        detail = "applied Duration is not a project-year span"
        Exit Function
    End If
    If Not modWorkbook.LoExists(GRID_COST_PROFILING_SHEET, TBL_COST_PROFILING) Then
        detail = "the profiling grid " & TBL_COST_PROFILING & " is missing"
        Exit Function
    End If
    If Not modWorkbook.LoExists(GRID_RISK_PROFILING_SHEET, TBL_RISK_PROFILING) Then
        detail = "the profiling grid " & TBL_RISK_PROFILING & " is missing"
        Exit Function
    End If
    Set costGrid = modWorkbook.Lo(GRID_COST_PROFILING_SHEET, TBL_COST_PROFILING)
    Set riskGrid = modWorkbook.Lo(GRID_RISK_PROFILING_SHEET, TBL_RISK_PROFILING)

    ReDim weights(0 To driverCount - 1, 0 To timeline.Duration - 1)
    For index = 0 To driverCount - 1
        If drivers(LBound(drivers) + index).IsRisk Then
            Set grid = riskGrid
            keyColumn = GCOL_RISK_PROFILING_RISK_ID
            fixedCols = GRID_RISK_PROFILING_FIXED_COLS
        Else
            Set grid = costGrid
            keyColumn = GCOL_COST_PROFILING_COST_LINE_ID
            fixedCols = GRID_COST_PROFILING_FIXED_COLS
        End If
        row = MatchingGridRow(grid, keyColumn, drivers(LBound(drivers) + index).PermanentId)
        If row = 0 Then
            detail = "profiling for driver " & drivers(LBound(drivers) + index).PermanentId & _
                     ": no grid row carries that Permanent ID"
            Exit Function
        End If
        For offset = 0 To timeline.Duration - 1
            column = YearColumn(grid, fixedCols, offset + 1)
            If column = 0 Then
                detail = "profiling for driver " & _
                         drivers(LBound(drivers) + index).PermanentId & _
                         ", project year " & CStr(offset + 1) & _
                         ": the applied project-year column is missing"
                Exit Function
            End If
            If Not NumericCell(grid, row, column, weight, _
                               "profiling for driver " & _
                               drivers(LBound(drivers) + index).PermanentId & _
                               ", project year " & CStr(offset + 1), detail) Then Exit Function
            weights(index, offset) = weight
        Next offset
    Next index
    ResolveProfileWeights = True
End Function

Private Function AttachDriverFx(ByRef model As ResolvedModel, ByRef detail As String) As Boolean
    ' Each driver's own rate, taken from the already-resolved referenced set.
    ' Nothing is looked up in the workbook a second time.
    Dim index As Long, probe As Long, found As Boolean
    If model.DriverCount = 0 Then
        AttachDriverFx = True
        Exit Function
    End If
    ReDim model.DriverFxRates(0 To model.DriverCount - 1)
    For index = 0 To model.DriverCount - 1
        found = False
        For probe = 0 To model.CurrencyCount - 1
            If StrComp(model.Currencies(probe), model.Drivers(index).Currency, _
                       vbBinaryCompare) = 0 Then
                model.DriverFxRates(index) = model.CurrencyRates(probe)
                found = True
                Exit For
            End If
        Next probe
        If Not found Then
            detail = "driver " & model.Drivers(index).PermanentId & _
                     ": currency is not in the resolved reference set"
            Exit Function
        End If
    Next index
    AttachDriverFx = True
End Function

' ==========================================================================
' Cell reading - exact, typed, and never repaired
' ==========================================================================
Private Function RawCellText(ByVal table As ListObject, ByVal rowIndex As Long, _
                             ByVal columnIndex As Long, ByRef text As String) As Boolean
    ' The cell's text EXACTLY as entered. Deliberately not modWorkbook.TextOf,
    ' which trims: trimming is right for deciding whether a row is populated and
    ' wrong for the key that a lookup will compare, because " USD " and "USD"
    ' are different keys and rewriting one into the other invents an answer.
    Dim cell As Range
    Set cell = modWorkbook.CellIn(table, rowIndex, columnIndex)
    If IsError(cell.Value) Then Exit Function
    If IsEmpty(cell.Value) Then Exit Function
    If IsObject(cell.Value) Then Exit Function
    text = CStr(cell.Value)
    RawCellText = True
End Function

Private Function ExactIdentifier(ByVal table As ListObject, ByVal rowIndex As Long, _
                                 ByVal columnIndex As Long, ByRef value As String, _
                                 ByVal where As String, ByRef detail As String) As Boolean
    Dim text As String
    If Not RawCellText(table, rowIndex, columnIndex, text) Then
        detail = where & ": the value is blank or unreadable"
        Exit Function
    End If
    If Len(Trim$(text)) = 0 Then
        ' Whitespace only carries no identity. It is refused rather than trimmed
        ' into an empty key that would then fail a lookup for the wrong reason.
        detail = where & ": the value is blank or whitespace only"
        Exit Function
    End If
    value = text
    ExactIdentifier = True
End Function

Private Function NumericCell(ByVal table As ListObject, ByVal rowIndex As Long, _
                             ByVal columnIndex As Long, ByRef value As Double, _
                             ByVal where As String, ByRef detail As String) As Boolean
    ' A blank is NOT zero. It is an assumption the user never made, and the
    ' whole point of seeding the inflation grid blank is that it cannot be
    ' fabricated as 0%.
    Dim cell As Range
    Set cell = modWorkbook.CellIn(table, rowIndex, columnIndex)
    If modWorkbook.IsEmptyCell(cell) Then
        detail = where & ": the value is blank. A blank is not zero."
        Exit Function
    End If
    If Not modWorkbook.TryReadDouble(cell.Value, value) Then
        detail = where & ": the value is not numeric"
        Exit Function
    End If
    If Not modCalcFactors.IsUsableDouble(value) Then
        detail = where & ": the value is not a usable Double"
        Exit Function
    End If
    NumericCell = True
End Function

Private Function NumericNamedCell(ByVal definedName As String, ByRef value As Double, _
                                  ByVal where As String, ByRef detail As String) As Boolean
    Dim raw As Variant
    If Not modWorkbook.NameExists(definedName) Then
        detail = where & ": the defined name " & definedName & " does not exist"
        Exit Function
    End If
    raw = modWorkbook.ReadValue(definedName)
    If IsEmpty(raw) Then
        detail = where & ": the value is blank. A blank is not zero."
        Exit Function
    End If
    If Not modWorkbook.TryReadDouble(raw, value) Then
        detail = where & ": the value is not numeric"
        Exit Function
    End If
    If Not modCalcFactors.IsUsableDouble(value) Then
        detail = where & ": the value is not a usable Double"
        Exit Function
    End If
    NumericNamedCell = True
End Function

Private Function AppliedWholeNumber(ByVal definedName As String, ByVal minValue As Double, _
                                    ByVal maxValue As Double, ByRef value As Double, _
                                    ByVal where As String, ByRef detail As String) As Boolean
    If Not modWorkbook.NameExists(definedName) Then
        detail = where & ": the defined name " & definedName & " does not exist"
        Exit Function
    End If
    If Not modWorkbook.IsWholeInRange(modWorkbook.ReadValue(definedName), minValue, _
                                      maxValue, value) Then
        detail = where & ": the applied value is missing or out of range"
        Exit Function
    End If
    AppliedWholeNumber = True
End Function

Private Function MatchingGridRow(ByVal table As ListObject, ByVal keyColumn As Long, _
                                 ByVal key As String) As Long
    ' The FIRST exact match, and 0 when there is none. Comparison is binary and
    ' the key is used exactly as it was read.
    Dim rowIndex As Long, text As String
    For rowIndex = 1 To modWorkbook.BodyRowCount(table)
        If RawCellText(table, rowIndex, keyColumn, text) Then
            If StrComp(text, key, vbBinaryCompare) = 0 Then
                MatchingGridRow = rowIndex
                Exit Function
            End If
        End If
    Next rowIndex
End Function

Private Function YearColumn(ByVal table As ListObject, ByVal fixedColumns As Long, _
                            ByVal headerValue As Long) As Long
    ' The generated year column carrying this header, or 0. Located by its
    ' HEADER rather than by arithmetic on a column index, so a grid that has not
    ' been regenerated for the applied timeline reports a missing column instead
    ' of silently reading the wrong year.
    Dim columnIndex As Long, header As Range, value As Double
    For columnIndex = fixedColumns + 1 To table.ListColumns.Count
        Set header = table.HeaderRowRange.Cells(1, columnIndex)
        If modWorkbook.TryReadDouble(header.Value, value) Then
            If value = CDbl(headerValue) Then
                YearColumn = columnIndex
                Exit Function
            End If
        End If
    Next columnIndex
End Function
