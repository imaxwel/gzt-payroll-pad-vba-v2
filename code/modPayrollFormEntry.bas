Attribute VB_Name = "modPayrollFormEntry"
'==============================================================================
' Module: modPayrollFormEntry
' Purpose: Public entry macros to show the payroll control form.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Sub: ShowPayrollForm
' Purpose: Show the payroll control form.
'------------------------------------------------------------------------------
Public Sub ShowPayrollForm()
    On Error GoTo ErrHandler
    frmPayrollMain.Show vbModeless
    Exit Sub
ErrHandler:
    LogError "modPayrollFormEntry", "ShowPayrollForm", Err.Number, Err.Description
    MsgBox "Failed to show payroll form: " & Err.Description, vbCritical, "HK Payroll Automation"
End Sub

'------------------------------------------------------------------------------
' Sub: startformMain
' Purpose: Alias for PAD compatibility.
'------------------------------------------------------------------------------
Public Sub startformMain()
    ShowPayrollForm
End Sub

