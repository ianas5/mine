Attribute VB_Name = "modSimulation"
'==============================================================================
' modSimulation
' Drives the Monte Carlo run. The Engine sheet is formula-based and volatile,
' so "running" the simulation = forcing a full recalculation, then refreshing
' the Results/Dashboard. This module also offers an optional "value snapshot"
' mode that freezes the current iteration outcomes so they stop changing.
'==============================================================================
Option Explicit

Public Const SHEET_ENGINE As String = "Engine"
Public Const SHEET_RESULTS As String = "Results"
Public Const SHEET_DASH As String = "Dashboard"
Public Const SHEET_CHECKS As String = "Checks"

' Main entry point — wired to the "Run Simulation" button on the Dashboard.
Public Sub RunMonteCarlo()
    Dim t As Double
    On Error GoTo Fail

    If Not ValidateInputs() Then
        MsgBox "Validation failed. See the Checks sheet — fix the red items first.", _
               vbExclamation, "Monte Carlo Cost Model"
        ThisWorkbook.Sheets(SHEET_CHECKS).Activate
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic
    t = Timer

    ' Force several volatile recalculations so RAND() re-rolls cleanly.
    Application.CalculateFull
    DoEvents
    Application.Calculate

    CalculatePercentiles            ' ensures Results are fresh (formulas already live)
    RefreshDashboard

    Application.ScreenUpdating = True
    MsgBox "Simulation complete." & vbCrLf & _
           "Iterations: " & ThisWorkbook.Sheets("Setup").Range("C9").Value & vbCrLf & _
           "Elapsed: " & Format(Timer - t, "0.0") & " s", _
           vbInformation, "Monte Carlo Cost Model"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "RunMonteCarlo error: " & Err.Description, vbCritical
End Sub

' Re-roll only (no message box) — handy for quick iteration.
Public Sub Reroll()
    Application.CalculateFull
End Sub
