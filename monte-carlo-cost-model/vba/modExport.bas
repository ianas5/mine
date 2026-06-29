Attribute VB_Name = "modExport"
'==============================================================================
' modExport
' Exports a management-ready report. Saves the Dashboard (and Results) to a PDF
' next to the workbook. No external add-ins required.
'==============================================================================
Option Explicit

Public Sub ExportReport()
    Dim path As String, fname As String
    On Error GoTo Fail

    path = ThisWorkbook.path
    If path = "" Then
        MsgBox "Please save the workbook first, then export.", vbExclamation
        Exit Sub
    End If
    fname = path & Application.PathSeparator & _
            "CostModel_Report_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"

    ' Export Dashboard + Results as a single PDF.
    ThisWorkbook.Sheets(Array("Dashboard", "Results")).Select
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=fname, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
    ThisWorkbook.Sheets("Dashboard").Select

    MsgBox "Report exported:" & vbCrLf & fname, vbInformation, "Export"
    Exit Sub
Fail:
    MsgBox "ExportReport error: " & Err.Description, vbCritical
End Sub
