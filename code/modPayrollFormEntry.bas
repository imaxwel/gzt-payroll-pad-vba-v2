Attribute VB_Name = "modPayrollFormEntry"
'==============================================================================
' Module: modPayrollFormEntry
' Purpose: Public entry macros to show the payroll control form.
' Note: If frmPayrollMain does not exist, run CreatePayrollForm macro first
'       (from modSetupForm module)
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Sub: ShowPayrollForm
' Purpose: Show the payroll control form.
'------------------------------------------------------------------------------
Public Sub ShowPayrollForm()
    On Error GoTo ErrHandler
    
    ' Check if form exists
    Dim frm As Object
    On Error Resume Next
    Set frm = UserForms.Add("frmPayrollMain")
    On Error GoTo ErrHandler
    
    If frm Is Nothing Then
        MsgBox "frmPayrollMain not found." & vbCrLf & vbCrLf & _
               "Please run CreatePayrollForm macro first to create the form." & vbCrLf & _
               "(Macros > modSetupForm.CreatePayrollForm)", _
               vbExclamation, "HK Payroll Automation"
        Exit Sub
    End If
    
    frm.Show vbModeless
    Exit Sub
    
ErrHandler:
    LogError "modPayrollFormEntry", "ShowPayrollForm", Err.Number, Err.Description
    MsgBox "Failed to show payroll form: " & Err.Description & vbCrLf & vbCrLf & _
           "If the form does not exist, run CreatePayrollForm macro first.", _
           vbCritical, "HK Payroll Automation"
End Sub

'------------------------------------------------------------------------------
' Sub: startformMain
' Purpose: Alias for PAD compatibility.
'------------------------------------------------------------------------------
Public Sub startformMain()
    ShowPayrollForm
End Sub

