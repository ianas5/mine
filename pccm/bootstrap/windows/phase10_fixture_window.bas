Attribute VB_Name = "modPhase10FixtureWindow"
Option Explicit

' ==========================================================================
' PCCM Phase 10 Step 4B - TRANSIENT FIXTURE-WINDOW SHIM
'
' TEST-ONLY. HARNESS-OWNED. NOT PRODUCTION.
'
' This module lives in `bootstrap/windows`, is never declared in
' `stage_b_manifest.json`, and is imported into the DISPOSABLE copy of the
' workbook that `phase10_benchmark.ps1` builds for one measurement session. It
' is the same mechanism, in the same place, as `phase5_gate_b_diagnostics.bas`,
' which Gate B has used since Phase 5.
'
' ==========================================================================
' WHY IT EXISTS: Application.Run CANNOT CARRY A ByRef
' ==========================================================================
' PERF-SMALL at 7077608 aborted before a single timed run:
'
'   statement : $null = $cell.ClearContents()
'   message   : The cell or chart you're trying to change is on a protected
'               sheet. To make a change, unprotect the sheet.
'
' An external COM caller cannot clear a cell on a protected sheet, so the
' benchmark needs worksheet protection released around its fixture writes.
'
' THE AUTHORITY FOR THAT ALREADY EXISTS AND IS NOT REIMPLEMENTED HERE.
' `modProtection.ProtectionBeginStructural` and `ProtectionEndStructural` are
' the accepted, depth-counted, single-owner window. What the harness cannot do
' is CALL them: both take `ByRef detail As String`, and every production and
' diagnostic procedure any accepted harness has ever invoked through
' `Application.Run` takes ByVal parameters or none. There is no precedent in
' this repository for marshalling a ByRef out-parameter across that boundary,
' and the failure mode - a lost or mangled diagnostic on the one call whose
' failure must abort the run - is not one to discover on Windows.
'
' So this module OWNS the String. It declares a local, passes it ByRef inside
' VBA where ByRef means what it says, and returns the result as a String - the
' shape every accepted `Application.Run` call in this tree already uses.
'
' ==========================================================================
' WHAT IT IS NOT
' ==========================================================================
' It holds NO protection policy. There is no `.Protect` and no `.Unprotect` in
' this file: a second implementation of the policy is exactly what the
' single-owner rule forbids, and a shim that decided anything would be one.
' Every decision is `modProtection`'s and every state it reports is read back
' out of the workbook.
'
' It does NOT call `modProtection.ProtectionRelease`. That is the maintenance
' path and it releases WORKBOOK STRUCTURE protection - `ThisWorkbook.Unprotect`
' - which no evidence asks for and which the structural window deliberately
' never touches. Structure protection stays applied throughout.
' ==========================================================================

' PROVES THE MODULE IMPORTED AND THE PROJECT COMPILES FAR ENOUGH TO CALL IT.
Public Function P10FW_Ping() As String
    P10FW_Ping = "OK|modPhase10FixtureWindow"
End Function

' THE WORKBOOK'S PROTECTION STATE, READ AND NOT DECIDED.
'
' The VERDICT is `modProtection.ProtectionIsApplied` - production's own single
' fact about itself, which is every worksheet protected AND the structure
' protected. The counts beside it are diagnostics, so a harness refusal can say
' "11 of 14" rather than "protection is wrong".
'
' READ IN VBA, ON PURPOSE. Phase-10 probe Run 8 failed reading
' `Worksheet.ProtectContents` across COM: PowerShell had no type information for
' the object and StrictMode turned a missing member into a terminating
' PropertyNotFoundException. Inside VBA the property is bound at compile time and
' that class of defect cannot occur.
Public Function P10FW_State() As String
    Dim sheet As Worksheet
    Dim total As Long, guarded As Long
    On Error GoTo Failed
    For Each sheet In ThisWorkbook.Worksheets
        total = total + 1
        If sheet.ProtectContents Then guarded = guarded + 1
    Next sheet
    P10FW_State = "OK|applied=" & CStr(modProtection.ProtectionIsApplied()) & _
                  "|depth=" & CStr(modProtection.ProtectionStructuralDepth()) & _
                  "|structure=" & CStr(ThisWorkbook.ProtectStructure) & _
                  "|sheets=" & CStr(total) & _
                  "|protected=" & CStr(guarded)
    Exit Function
Failed:
    P10FW_State = "FAIL|the protection state could not be read: " & Err.Description
End Function

' OPEN. Forwards, and returns what the owner said.
'
' A FAILED OPEN IS REPORTED, NEVER RETRIED. `ProtectionBeginStructural` already
' re-applies protection itself if it could not establish a clean release, so a
' FAIL here means the workbook is protected and the fixture must not be built.
Public Function P10FW_Begin() As String
    Dim detail As String
    If modProtection.ProtectionBeginStructural(detail) Then
        P10FW_Begin = "OK|depth=" & CStr(modProtection.ProtectionStructuralDepth())
    Else
        P10FW_Begin = "FAIL|" & detail
    End If
End Function

' CLOSE. The outermost close routes through `modProtection.ProtectionApply` -
' the same function `Workbook_Open` uses - and then requires
' `ProtectionIsApplied` before it will report success. That is the accepted
' architecture's own definition of "restored", including the
' UserInterfaceOnly:=True flag, which Excel exposes no property to read back.
Public Function P10FW_End() As String
    Dim detail As String
    If modProtection.ProtectionEndStructural(detail) Then
        P10FW_End = "OK|depth=" & CStr(modProtection.ProtectionStructuralDepth())
    Else
        P10FW_End = "FAIL|" & detail
    End If
End Function
