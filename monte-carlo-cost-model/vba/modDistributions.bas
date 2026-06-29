Attribute VB_Name = "modDistributions"
'==============================================================================
' modDistributions
' Sampling functions for 3-point estimates. These mirror the in-sheet formulas
' and can be used if you prefer to build a VBA-array engine instead of the
' formula engine (faster for very large iteration counts).
'==============================================================================
Option Explicit

' Triangular distribution via inverse-transform sampling.
Public Function SampleTriangular(ByVal minVal As Double, ByVal modeVal As Double, _
                                 ByVal maxVal As Double) As Double
    Dim u As Double, fc As Double
    If maxVal <= minVal Then SampleTriangular = minVal: Exit Function
    u = Rnd()
    fc = (modeVal - minVal) / (maxVal - minVal)
    If u < fc Then
        SampleTriangular = minVal + Sqr(u * (maxVal - minVal) * (modeVal - minVal))
    Else
        SampleTriangular = maxVal - Sqr((1 - u) * (maxVal - minVal) * (maxVal - modeVal))
    End If
End Function

' Normal sample using a 3-point heuristic: mean = mode, sd = (max - min) / 6.
Public Function SampleNormal(ByVal minVal As Double, ByVal modeVal As Double, _
                             ByVal maxVal As Double) As Double
    Dim sd As Double
    sd = (maxVal - minVal) / 6#
    SampleNormal = Application.WorksheetFunction.Norm_Inv(Rnd(), modeVal, sd)
    If SampleNormal < 0 Then SampleNormal = 0
End Function

' PERT (scaled Beta) sample.
Public Function SamplePert(ByVal minVal As Double, ByVal modeVal As Double, _
                           ByVal maxVal As Double) As Double
    Dim a As Double, b As Double
    If maxVal <= minVal Then SamplePert = minVal: Exit Function
    a = 1 + 4 * (modeVal - minVal) / (maxVal - minVal)
    b = 1 + 4 * (maxVal - modeVal) / (maxVal - minVal)
    SamplePert = Application.WorksheetFunction.Beta_Inv(Rnd(), a, b, minVal, maxVal)
End Function

' Dispatch by distribution name.
Public Function SampleByName(ByVal dist As String, ByVal mn As Double, _
                             ByVal ml As Double, ByVal mx As Double) As Double
    Select Case LCase$(Trim$(dist))
        Case "pert":   SampleByName = SamplePert(mn, ml, mx)
        Case "normal": SampleByName = SampleNormal(mn, ml, mx)
        Case Else:     SampleByName = SampleTriangular(mn, ml, mx)  ' triangular default
    End Select
End Function
