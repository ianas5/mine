Attribute VB_Name = "modEngine"
'==============================================================================
' modEngine — the real Monte Carlo engine (used by MonteCarloCostModel_VBA.xlsx)
'
' Runs the whole simulation in memory, reading the input TABLES so everything is
' dynamic: any number of cost lines / risks (just add rows), any number of
' iterations or years (Setup cells). It then writes the Results, Sensitivity,
' Dashboard KPIs and chart data.
'
' Requires: modDistributions (SampleByName / SampleTriangular / SamplePert /
'           SampleNormal). Assign the Dashboard "Run Simulation" button to
'           RunSimulation.
'
' Layout constants below mirror generate_vba_model.py. If you restructure the
' sheets, update these.
'==============================================================================
Option Explicit

' ---- Setup cells
Private Const SET_YEARS As String = "C8"
Private Const SET_ITERS As String = "C9"
Private Const SET_CONF  As String = "C10"
Private Const SET_DISC  As String = "C11"
Private Const SET_CONFP As String = "C16"
Private Const MAXY      As Long = 30      ' FY columns built into the profiling tables

Public Sub RunSimulation()
    Dim t0 As Double: t0 = Timer
    On Error GoTo Fail

    If Not AllChecksPass() Then
        If MsgBox("Some checks on the Checks sheet are FAIL. Run anyway?", _
                  vbQuestion + vbYesNo, "Monte Carlo") = vbNo Then
            Sheets("Checks").Activate: Exit Sub
        End If
    End If

    Dim ws As Worksheet: Set ws = Sheets("Setup")
    Dim nYears As Long:  nYears = CLng(ws.Range(SET_YEARS).Value)
    Dim nIter As Long:   nIter = CLng(ws.Range(SET_ITERS).Value)
    Dim disc As Double:  disc = CDbl(ws.Range(SET_DISC).Value)
    Dim confP As Double: confP = CDbl(ws.Range(SET_CONFP).Value)
    If nYears < 1 Then nYears = 1
    If nIter < 1 Then nIter = 1

    SyncTables   ' make sure every cost line / risk has a matching profiling row

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Monte Carlo: reading inputs..."

    '---- inflation factors (FY1..nYears) from the Inflation sheet (col D = factor)
    Dim fac() As Double: ReDim fac(1 To nYears)
    Dim infS As Worksheet: Set infS = Sheets("Inflation")
    Dim y As Long
    For y = 1 To nYears: fac(y) = CDbl(infS.Cells(3 + y, 4).Value): Next y

    '---- cost lines + their profiles
    Dim clT As ListObject: Set clT = Sheets("Cost Lines").ListObjects("tbl_CostLines")
    Dim cpT As ListObject: Set cpT = Sheets("Cost Profiling").ListObjects("tbl_CostProfile")
    Dim nL As Long: nL = clT.DataBodyRange.Rows.Count
    Dim lMin() As Double, lML() As Double, lMax() As Double, lDist() As String, lOn() As Boolean, lName() As String
    Dim lProf() As Double
    ReDim lMin(1 To nL): ReDim lML(1 To nL): ReDim lMax(1 To nL)
    ReDim lDist(1 To nL): ReDim lOn(1 To nL): ReDim lName(1 To nL): ReDim lProf(1 To nL, 1 To nYears)
    Dim i As Long
    For i = 1 To nL
        lName(i) = CStr(clT.DataBodyRange.Cells(i, 2).Value)
        lMin(i) = Num(clT.DataBodyRange.Cells(i, 10).Value)   ' Min Total
        lML(i) = Num(clT.DataBodyRange.Cells(i, 11).Value)    ' ML Total
        lMax(i) = Num(clT.DataBodyRange.Cells(i, 12).Value)   ' Max Total
        lDist(i) = CStr(clT.DataBodyRange.Cells(i, 13).Value)
        lOn(i) = (UCase$(Trim$(CStr(clT.DataBodyRange.Cells(i, 14).Value))) = "YES")
        For y = 1 To nYears
            lProf(i, y) = ProfileFrac(cpT, i, y, nYears)
        Next y
    Next i

    '---- risks + their profiles
    Dim rrT As ListObject: Set rrT = Sheets("Risk Register").ListObjects("tbl_RiskRegister")
    Dim rpT As ListObject: Set rpT = Sheets("Risk Profiling").ListObjects("tbl_RiskProfile")
    Dim nR As Long: nR = rrT.DataBodyRange.Rows.Count
    Dim rProb() As Double, rMin() As Double, rML() As Double, rMax() As Double
    Dim rDist() As String, rOn() As Boolean, rName() As String, rProf() As Double
    ReDim rProb(1 To nR): ReDim rMin(1 To nR): ReDim rML(1 To nR): ReDim rMax(1 To nR)
    ReDim rDist(1 To nR): ReDim rOn(1 To nR): ReDim rName(1 To nR): ReDim rProf(1 To nR, 1 To nYears)
    For i = 1 To nR
        rName(i) = CStr(rrT.DataBodyRange.Cells(i, 2).Value)
        rProb(i) = Num(rrT.DataBodyRange.Cells(i, 4).Value)
        rMin(i) = Num(rrT.DataBodyRange.Cells(i, 5).Value)
        rML(i) = Num(rrT.DataBodyRange.Cells(i, 6).Value)
        rMax(i) = Num(rrT.DataBodyRange.Cells(i, 7).Value)
        rDist(i) = CStr(rrT.DataBodyRange.Cells(i, 8).Value)
        rOn(i) = (UCase$(Trim$(CStr(rrT.DataBodyRange.Cells(i, 10).Value))) = "YES")
        For y = 1 To nYears
            rProf(i, y) = ProfileFrac(rpT, i, y, nYears)
        Next y
    Next i

    '---- base (deterministic) = sum of ML totals of included lines
    Dim baseML As Double
    For i = 1 To nL: If lOn(i) Then baseML = baseML + lML(i)
    Next i

    '---- simulate
    Application.StatusBar = "Monte Carlo: running " & Format(nIter, "#,##0") & " iterations..."
    Dim total() As Double, npvA() As Double, baseNom() As Double, riskNom() As Double
    Dim yearT() As Double, lSamp() As Double, rSamp() As Double
    ReDim total(1 To nIter): ReDim npvA(1 To nIter)
    ReDim baseNom(1 To nIter): ReDim riskNom(1 To nIter)
    ReDim yearT(1 To nYears, 1 To nIter)
    ReDim lSamp(1 To nL, 1 To nIter): ReDim rSamp(1 To nR, 1 To nIter)

    Dim it As Long, s As Double, yb() As Double, yr() As Double, tot As Double, np As Double, bn As Double, rn As Double
    ReDim yb(1 To nYears): ReDim yr(1 To nYears)
    Randomize 12345
    For it = 1 To nIter
        For y = 1 To nYears: yb(y) = 0: yr(y) = 0: Next y
        bn = 0: rn = 0
        For i = 1 To nL
            If lOn(i) Then
                s = SampleByName(lDist(i), lMin(i), lML(i), lMax(i))
            Else
                s = 0
            End If
            lSamp(i, it) = s: bn = bn + s
            For y = 1 To nYears: yb(y) = yb(y) + s * lProf(i, y): Next y
        Next i
        For i = 1 To nR
            s = 0
            If rOn(i) Then
                If Rnd() <= rProb(i) Then s = SampleByName(rDist(i), rMin(i), rML(i), rMax(i))
            End If
            rSamp(i, it) = s: rn = rn + s
            For y = 1 To nYears: yr(y) = yr(y) + s * rProf(i, y): Next y
        Next i
        tot = 0: np = 0
        For y = 1 To nYears
            Dim ytv As Double: ytv = (yb(y) + yr(y)) * fac(y)
            yearT(y, it) = ytv
            tot = tot + ytv
            np = np + ytv / (1 + disc) ^ y
        Next y
        total(it) = tot: npvA(it) = np: baseNom(it) = bn: riskNom(it) = rn
    Next it

    '---- percentiles + write outputs
    Application.StatusBar = "Monte Carlo: writing results..."
    WriteResults total, npvA, baseNom, riskNom, yearT, nYears, nIter, baseML, confP, disc
    WriteSensitivity lName, rName, lSamp, rSamp, lOn, rOn, nL, nR, nIter
    WriteCharts total, yearT, nYears, nIter
    Sheets("Checks").Calculate

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Sheets("Dashboard").Activate
    MsgBox "Simulation complete." & vbCrLf & _
           Format(nIter, "#,##0") & " iterations, " & nL & " cost lines, " & nR & _
           " risks, " & nYears & " years." & vbCrLf & "Elapsed: " & Format(Timer - t0, "0.0") & " s", _
           vbInformation, "Monte Carlo Cost Model"
    Exit Sub
Fail:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "RunSimulation error: " & Err.Description, vbCritical
End Sub

' Tidy the input sheets to show only the active number of years (cosmetic; the
' engine reads "Number of years" regardless). Assign to the "Apply Years" button.
Public Sub ApplyYears()
    Dim n As Long: n = CLng(Sheets("Setup").Range(SET_YEARS).Value)
    Dim cp As Worksheet, rp As Worksheet, infS As Worksheet, y As Long, hide As Boolean
    Set cp = Sheets("Cost Profiling"): Set rp = Sheets("Risk Profiling"): Set infS = Sheets("Inflation")
    Application.ScreenUpdating = False
    For y = 1 To 30
        hide = (y > n)
        cp.Columns(2 + y).Hidden = hide       ' FY columns start at C (col 3)
        rp.Columns(2 + y).Hidden = hide
        infS.Rows(3 + y).Hidden = hide        ' FY rows start at row 4
    Next y
    Application.ScreenUpdating = True
End Sub

' Keep the profiling tables in step with the register tables: one profiling row
' per cost line / risk, name/ID linked, new rows defaulted to an even split over
' the active years. Runs automatically before each simulation; can also be run
' on its own (Alt+F8 ▸ SyncTables) after you add cost lines or risks.
Public Sub SyncTables()
    Dim nA As Long: nA = CLng(Sheets("Setup").Range(SET_YEARS).Value)
    If nA < 1 Then nA = 1
    If nA > MAXY Then nA = MAXY
    Application.ScreenUpdating = False
    SyncOne Sheets("Cost Lines").ListObjects("tbl_CostLines"), _
            Sheets("Cost Profiling").ListObjects("tbl_CostProfile"), nA
    SyncOne Sheets("Risk Register").ListObjects("tbl_RiskRegister"), _
            Sheets("Risk Profiling").ListObjects("tbl_RiskProfile"), nA
    Application.ScreenUpdating = True
End Sub

' Add a fully-formatted cost line (inherits the table's formatting, dropdowns and
' the Min/ML/Max Total formulas) and create its matching profiling row.
Public Sub AddCostLine()
    Dim t As ListObject, r As ListRow, rr As Long, n As Long
    On Error GoTo Fail
    Set t = Sheets("Cost Lines").ListObjects("tbl_CostLines")
    Set r = t.ListRows.Add
    n = t.ListRows.Count: rr = r.Range.row
    With r.Range
        .Cells(1, 1).Value = n & ".0"
        .Cells(1, 2).Value = "New Cost Item"
        .Cells(1, 5).Value = "LS"
        .Cells(1, 6).Value = 1
        If .Cells(1, 10).Formula = "" Then .Cells(1, 10).Formula = "=F" & rr & "*G" & rr
        If .Cells(1, 11).Formula = "" Then .Cells(1, 11).Formula = "=F" & rr & "*H" & rr
        If .Cells(1, 12).Formula = "" Then .Cells(1, 12).Formula = "=F" & rr & "*I" & rr
        .Cells(1, 13).Value = "Triangular"
        .Cells(1, 14).Value = "Yes"
    End With
    SyncTables
    Sheets("Cost Lines").Activate: Application.Goto r.Range.Cells(1, 7), False
    MsgBox "New cost line added (fully formatted)." & vbCrLf & _
           "Enter Min / Most Likely / Max unit cost. A matching Cost Profiling row was created.", _
           vbInformation, "Add Cost Line"
    Exit Sub
Fail: MsgBox "AddCostLine error: " & Err.Description, vbCritical
End Sub

' Add a fully-formatted risk and create its matching risk-profiling row.
Public Sub AddRisk()
    Dim t As ListObject, r As ListRow, n As Long
    On Error GoTo Fail
    Set t = Sheets("Risk Register").ListObjects("tbl_RiskRegister")
    Set r = t.ListRows.Add
    n = t.ListRows.Count
    With r.Range
        .Cells(1, 1).Value = "R" & n
        .Cells(1, 2).Value = "New Risk"
        .Cells(1, 4).Value = 0.3
        .Cells(1, 8).Value = "Triangular"
        .Cells(1, 10).Value = "Yes"
    End With
    SyncTables
    Sheets("Risk Register").Activate: Application.Goto r.Range.Cells(1, 5), False
    MsgBox "New risk added (fully formatted)." & vbCrLf & _
           "Enter Probability and Min / Most Likely / Max impact. A matching Risk Profiling row was created.", _
           vbInformation, "Add Risk"
    Exit Sub
Fail: MsgBox "AddRisk error: " & Err.Description, vbCritical
End Sub

Private Sub SyncOne(regTbl As ListObject, profTbl As ListObject, ByVal nActive As Long)
    Dim nReg As Long: nReg = regTbl.DataBodyRange.Rows.Count
    Dim nProf As Long
    On Error Resume Next
    nProf = profTbl.DataBodyRange.Rows.Count
    On Error GoTo 0
    Do While nProf < nReg: profTbl.ListRows.Add: nProf = nProf + 1: Loop
    Do While nProf > nReg: profTbl.ListRows(nProf).Delete: nProf = nProf - 1: Loop

    Dim regSheet As String: regSheet = regTbl.Parent.Name
    Dim ncols As Long: ncols = profTbl.ListColumns.Count        ' WBS/ID, Name, FY1..FYMAXY, Total%
    Dim i As Long, y As Long, sm As Double
    For i = 1 To nReg
        profTbl.DataBodyRange.Cells(i, 1).Formula = _
            "='" & regSheet & "'!" & regTbl.DataBodyRange.Cells(i, 1).Address(False, False)
        profTbl.DataBodyRange.Cells(i, 2).Formula = _
            "='" & regSheet & "'!" & regTbl.DataBodyRange.Cells(i, 2).Address(False, False)
        sm = 0
        For y = 1 To MAXY: sm = sm + Num(profTbl.DataBodyRange.Cells(i, 2 + y).Value): Next y
        If sm = 0 Then                                          ' new/blank row -> even split
            For y = 1 To MAXY
                profTbl.DataBodyRange.Cells(i, 2 + y).Value = IIf(y <= nActive, 1# / nActive, 0)
            Next y
        End If
        profTbl.DataBodyRange.Cells(i, ncols).Formula = "=SUM(" & _
            profTbl.DataBodyRange.Cells(i, 3).Address(False, False) & ":" & _
            profTbl.DataBodyRange.Cells(i, 2 + MAXY).Address(False, False) & ")"
    Next i
End Sub

'------------------------------------------------------------------ helpers
Private Function ProfileFrac(tbl As ListObject, ByVal row As Long, ByVal y As Long, ByVal nYears As Long) As Double
    ' fraction for year y of profiling table row `row`; even split if missing
    On Error GoTo Even
    If row <= tbl.DataBodyRange.Rows.Count Then
        ProfileFrac = Num(tbl.DataBodyRange.Cells(row, 2 + y).Value)   ' cols: 1=ID,2=Name,3=FY1...
        Exit Function
    End If
Even:
    ProfileFrac = 1# / nYears
End Function

Private Function Num(ByVal v As Variant) As Double
    If IsNumeric(v) Then Num = CDbl(v) Else Num = 0#
End Function

Private Function Pctl(arr() As Double, ByVal p As Double) As Double
    Pctl = Application.WorksheetFunction.Percentile_Inc(arr, p)
End Function

Private Function ColPctl(mat() As Double, ByVal rowIdx As Long, ByVal n As Long, ByVal p As Double) As Double
    Dim tmp() As Double, j As Long: ReDim tmp(1 To n)
    For j = 1 To n: tmp(j) = mat(rowIdx, j): Next j
    ColPctl = Application.WorksheetFunction.Percentile_Inc(tmp, p)
End Function

Private Function Avg(arr() As Double) As Double
    Avg = Application.WorksheetFunction.Average(arr)
End Function

Private Function AllChecksPass() As Boolean
    Dim c As Range: AllChecksPass = True
    For Each c In Sheets("Checks").Range("B5:B100").Cells
        If UCase$(CStr(c.Value)) = "FAIL" Then AllChecksPass = False
    Next c
End Function

'------------------------------------------------------------------ writers
Private Sub WriteResults(total() As Double, npvA() As Double, baseNom() As Double, riskNom() As Double, _
                         yearT() As Double, ByVal nYears As Long, ByVal nIter As Long, _
                         ByVal baseML As Double, ByVal confP As Double, ByVal disc As Double)
    Dim rs As Worksheet: Set rs = Sheets("Results")
    ' A. percentiles  (B6..B13)
    Dim ps, i As Long
    ps = Array(0.1, 0.3, 0.5, 0.6, 0.7, 0.8, 0.9, 0.95)
    For i = 0 To 7: rs.Cells(6 + i, 2).Value = Pctl(total, ps(i)): Next i
    ' B. contingency (rows 17..20: P50,P70,P80,P90)
    Dim cp, k As Long: cp = Array(0.5, 0.7, 0.8, 0.9)
    For i = 0 To 3
        Dim tc As Double: tc = Pctl(total, cp(i))
        rs.Cells(17 + i, 2).Value = tc
        rs.Cells(17 + i, 3).Value = baseML
        rs.Cells(17 + i, 4).Value = tc - baseML
        rs.Cells(17 + i, 5).Value = IIf(baseML <> 0, (tc - baseML) / baseML, 0)
    Next i
    ' C. annual cash flow (rows 24..) per year, columns P50,P70,P80,P90
    Dim startY As Long: startY = CLng(Sheets("Setup").Range("C7").Value)
    Dim y As Long
    For y = 1 To nYears
        rs.Cells(23 + y, 1).Value = "FY" & y & " (" & (startY + y - 1) & ")"
        rs.Cells(23 + y, 2).Value = ColPctl(yearT, y, nIter, 0.5)
        rs.Cells(23 + y, 3).Value = ColPctl(yearT, y, nIter, 0.7)
        rs.Cells(23 + y, 4).Value = ColPctl(yearT, y, nIter, 0.8)
        rs.Cells(23 + y, 5).Value = ColPctl(yearT, y, nIter, 0.9)
    Next y
    rs.Range("A" & (24 + nYears) & ":E53").ClearContents   ' clear leftover rows
    ' D. NPV (rows 58..61)
    For i = 0 To 3
        rs.Cells(58 + i, 2).Value = Pctl(total, cp(i))
        rs.Cells(58 + i, 3).Value = Pctl(npvA, cp(i))
    Next i
    ' Dashboard KPIs
    Dim db As Worksheet: Set db = Sheets("Dashboard")
    db.Range("A12").Value = baseML
    db.Range("C12").Value = Avg(baseNom)
    db.Range("E12").Value = Avg(riskNom)
    db.Range("G12").Value = Avg(total)
    db.Range("A15").Value = Pctl(total, 0.5)
    db.Range("C15").Value = Pctl(total, 0.7)
    db.Range("E15").Value = Pctl(total, 0.8)
    db.Range("G15").Value = Pctl(total, 0.9)
    db.Range("A18").Value = Pctl(total, confP) - baseML
    db.Range("C18").Value = Pctl(total, confP)
    db.Range("E18").Value = Pctl(npvA, confP)
    db.Range("G18").Value = Pctl(total, 0.95)
End Sub

Private Sub WriteSensitivity(lName() As String, rName() As String, lSamp() As Double, rSamp() As Double, _
                             lOn() As Boolean, rOn() As Boolean, ByVal nL As Long, ByVal nR As Long, ByVal nIter As Long)
    Dim sn As Worksheet: Set sn = Sheets("Sensitivity")
    sn.Range("A5:E40").ClearContents
    Dim row As Long: row = 5
    Dim i As Long, p10 As Double, p90 As Double
    For i = 1 To nL
        If lOn(i) Then
            p10 = ColPctl(lSamp, i, nIter, 0.1): p90 = ColPctl(lSamp, i, nIter, 0.9)
            sn.Cells(row, 1).Value = lName(i): sn.Cells(row, 2).Value = "Cost"
            sn.Cells(row, 3).Value = p10: sn.Cells(row, 4).Value = p90: sn.Cells(row, 5).Value = p90 - p10
            row = row + 1
        End If
    Next i
    For i = 1 To nR
        If rOn(i) Then
            p10 = ColPctl(rSamp, i, nIter, 0.1): p90 = ColPctl(rSamp, i, nIter, 0.9)
            sn.Cells(row, 1).Value = rName(i): sn.Cells(row, 2).Value = "Risk"
            sn.Cells(row, 3).Value = p10: sn.Cells(row, 4).Value = p90: sn.Cells(row, 5).Value = p90 - p10
            row = row + 1
        End If
    Next i
End Sub

Private Sub WriteCharts(total() As Double, yearT() As Double, ByVal nYears As Long, ByVal nIter As Long)
    Dim cd As Worksheet: Set cd = Sheets("ChartData")
    Dim lo As Double, hi As Double, i As Long
    lo = Application.WorksheetFunction.Min(total)
    hi = Application.WorksheetFunction.Max(total)
    Dim NB As Long: NB = 30
    Dim w As Double: w = (hi - lo) / NB: If w = 0 Then w = 1
    Dim cnt() As Long: ReDim cnt(1 To NB)
    Dim b As Long
    For i = 1 To nIter
        b = Int((total(i) - lo) / w) + 1
        If b < 1 Then b = 1
        If b > NB Then b = NB
        cnt(b) = cnt(b) + 1
    Next i
    For b = 1 To NB
        cd.Cells(1 + b, 1).Value = lo + (b - 0.5) * w     ' center
        cd.Cells(1 + b, 2).Value = cnt(b)                 ' count
    Next b
    ' S-curve (21 points)
    Dim NP As Long: NP = 21
    Dim k As Long
    For k = 0 To NP - 1
        cd.Cells(2 + k, 4).Value = Pctl(total, k / (NP - 1))
        cd.Cells(2 + k, 5).Value = k / (NP - 1)
    Next k
    ' cumulative cash flow by year (P50 / P80)
    Dim startY As Long: startY = CLng(Sheets("Setup").Range("C7").Value)
    Dim c50 As Double, c80 As Double, y As Long
    c50 = 0: c80 = 0
    For y = 1 To nYears
        c50 = c50 + ColPctl(yearT, y, nIter, 0.5)
        c80 = c80 + ColPctl(yearT, y, nIter, 0.8)
        cd.Cells(1 + y, 7).Value = startY + y - 1
        cd.Cells(1 + y, 8).Value = c50
        cd.Cells(1 + y, 9).Value = c80
    Next y
    cd.Range("G" & (2 + nYears) & ":I31").ClearContents
End Sub
