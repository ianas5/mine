Attribute VB_Name = "modSetup"
'==============================================================================
' modSetup  —  in-Excel duration control (FLEX build only)
'
' The FLEX workbook is built for MAXY (30) years with the unused years pre-hidden.
' ApplySettings reads "Number of years" from the Setup sheet and shows exactly
' that many FY columns/rows across Cost Profiling, Risk Profiling, Inflation,
' Results and the Engine — so you change the duration without regenerating.
'
'   • Start year  : handled by the Setup "Start year" cell (labels follow it).
'   • Currency    : handled by the Setup "Currency" cell (no symbol is baked in).
'   • Duration    : run ApplySettings after editing "Number of years".
'
' Layout constants below MUST match the generator (generate_advanced_model.py,
' FLEX mode). They are verified against the shipped file:
'   Cost/Risk Profiling : first FY column = C (3)  -> year y at column 2 + y
'   Inflation           : first FY row    = 4      -> year y at row    3 + y
'   Results cash flow    : first row       = 24     -> year y at row    23 + y
'   Engine year columns  : first YT column = L (12) -> year y at column 11 + y
'   Setup "Number of years" : cell C8
'==============================================================================
Option Explicit

Public Const MAXY As Long = 30
Public Const NYEARS_CELL As String = "C8"

Public Sub ApplySettings()
    Dim active As Long, y As Long, hide As Boolean
    Dim cp As Worksheet, rp As Worksheet, inf As Worksheet, rs As Worksheet, en As Worksheet
    On Error GoTo Fail

    active = CLng(Val(ThisWorkbook.Sheets("Setup").Range(NYEARS_CELL).Value))
    If active < 1 Then active = 1
    If active > MAXY Then
        MsgBox "This workbook is built for up to " & MAXY & " years. " & _
               "For a longer horizon, regenerate with the Python script.", vbExclamation
        active = MAXY
    End If

    Application.ScreenUpdating = False
    Set cp = ThisWorkbook.Sheets("Cost Profiling")
    Set rp = ThisWorkbook.Sheets("Risk Profiling")
    Set inf = ThisWorkbook.Sheets("Inflation")
    Set rs = ThisWorkbook.Sheets("Results")
    Set en = ThisWorkbook.Sheets("Engine")

    For y = 1 To MAXY
        hide = (y > active)
        cp.Columns(2 + y).Hidden = hide        ' Cost Profiling FY column
        rp.Columns(2 + y).Hidden = hide        ' Risk Profiling FY column
        inf.Rows(3 + y).Hidden = hide          ' Inflation FY row
        rs.Rows(23 + y).Hidden = hide          ' Results cash-flow row
        en.Columns(11 + y).Hidden = hide       ' Engine year column (calc)
    Next y

    Application.CalculateFull
    ThisWorkbook.Sheets("Checks").Calculate
    Application.ScreenUpdating = True

    MsgBox "Model set to " & active & " active year(s)." & vbCrLf & vbCrLf & _
           "Check that each Cost/Risk Profile row still sums to 100% across the " & _
           "active years — the Checks sheet flags any that don't. Newly switched-on " & _
           "years start at 0%, so edit their profile cells.", _
           vbInformation, "Apply Settings"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "ApplySettings error: " & Err.Description, vbCritical
End Sub

' Convenience: bump the active years up/down by one and re-apply.
Public Sub YearsPlus():  ChangeYears 1:  End Sub
Public Sub YearsMinus(): ChangeYears -1: End Sub

Private Sub ChangeYears(ByVal delta As Long)
    Dim c As Range
    Set c = ThisWorkbook.Sheets("Setup").Range(NYEARS_CELL)
    c.Value = Application.WorksheetFunction.Max(1, _
              Application.WorksheetFunction.Min(MAXY, CLng(Val(c.Value)) + delta))
    ApplySettings
End Sub
