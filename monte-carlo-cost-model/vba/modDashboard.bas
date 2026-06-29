Attribute VB_Name = "modDashboard"
'==============================================================================
' modDashboard
' Refreshes the Dashboard after a run: recalculates dependent sheets, stamps a
' "last run" timestamp, and (optionally) re-points charts. Because the Dashboard
' KPIs are live formulas, this mostly just forces calc + a visual refresh.
'==============================================================================
Option Explicit

Public Sub RefreshDashboard()
    Dim db As Worksheet
    On Error Resume Next
    Set db = ThisWorkbook.Sheets("Dashboard")

    ThisWorkbook.Sheets("Results").Calculate
    ThisWorkbook.Sheets("Sensitivity").Calculate
    ThisWorkbook.Sheets("ChartData").Calculate
    db.Calculate

    ' Stamp last-run time into a named cell if present (create the name to use).
    On Error Resume Next
    db.Range("LastRun").Value = "Last run: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    On Error GoTo 0

    db.Activate
    Application.Goto db.Range("A1"), True
End Sub

' Recalc helper used by the run routine.
Public Sub CalculatePercentiles()
    ThisWorkbook.Sheets("Results").Calculate
End Sub
