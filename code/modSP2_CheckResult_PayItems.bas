Attribute VB_Name = "modSP2_CheckResult_PayItems"
'==============================================================================
' Module: modSP2_CheckResult_PayItems
' Purpose: Subprocess 2 - Pay Items Check columns
' Description: Validates Base Pay, Leave Payments, EAO Adjustments
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Sub: SP2_Check_PayItems
' Purpose: Populate pay items Check columns
'------------------------------------------------------------------------------
Public Sub SP2_Check_PayItems(valWb As Workbook, weinIndex As Object)
    Dim ws As Worksheet
    
    On Error GoTo ErrHandler
    
    Set ws = valWb.Worksheets("Check Result")
    
    ' Load EAO data
    LoadEAOData
    
    ' Process each WEIN
    Dim wein As Variant
    Dim row As Long
    
    For Each wein In weinIndex.Keys
        row = weinIndex(wein)
        
        ' Write Check values
        WriteBasePayCheck ws, row, CStr(wein)
        WriteLeavePaymentChecks ws, row, CStr(wein)
        WriteEAOChecks ws, row, CStr(wein)
    Next wein
    
    LogInfo "modSP2_CheckResult_PayItems", "SP2_Check_PayItems", "Pay items checks completed"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_PayItems", "SP2_Check_PayItems", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: WriteBasePayCheck
' Purpose: Write Base Pay Check columns
'------------------------------------------------------------------------------
Private Sub WriteBasePayCheck(ws As Worksheet, row As Long, wein As String)
    Dim col As Long
    Dim rec As tEAORecord
    
    On Error Resume Next
    
    rec = GetEAORecord(wein)
    
    ' Base Pay 60001000 Check
    col = FindColumnByHeader(ws.Rows(4), "Base Pay 60001000 Check")
    If col > 0 Then
        ' Formula: Actual working days * Monthly Salary / calendar days
        ' This would need actual working days from another source
        ' Placeholder implementation
    End If
End Sub

'------------------------------------------------------------------------------
' Sub: WriteLeavePaymentChecks
' Purpose: Write leave payment Check columns
'------------------------------------------------------------------------------
Private Sub WriteLeavePaymentChecks(ws As Worksheet, row As Long, wein As String)
    Dim col As Long
    
    On Error Resume Next
    
    ' Maternity Leave Payment Check
    col = FindColumnByHeader(ws.Rows(4), "Maternity Leave Payment Check")
    If col > 0 Then
        ws.Cells(row, col).Value = CalcMaternityLeavePayment(wein)
    End If
    
    ' Sick Leave Payment Check
    col = FindColumnByHeader(ws.Rows(4), "Sick Leave Payment Check")
    If col > 0 Then
        ws.Cells(row, col).Value = CalcSickLeavePayment(wein)
    End If
    
    ' PPTO Payment Check
    col = FindColumnByHeader(ws.Rows(4), "PPTO Payment Check")
    If col > 0 Then
        ws.Cells(row, col).Value = CalcPPTOPayment(wein)
    End If
    
    ' No Pay Leave Deduction Check
    col = FindColumnByHeader(ws.Rows(4), "No Pay Leave Deduction Check")
    If col > 0 Then
        ws.Cells(row, col).Value = CalcNoPayLeaveDeduction(wein)
    End If
End Sub

'------------------------------------------------------------------------------
' Sub: WriteEAOChecks
' Purpose: Write EAO adjustment Check columns
'------------------------------------------------------------------------------
Private Sub WriteEAOChecks(ws As Worksheet, row As Long, wein As String)
    Dim col As Long
    
    On Error Resume Next
    
    ' Total EAO Adj Check
    col = FindColumnByHeader(ws.Rows(4), "Total EAO Adj Check")
    If col > 0 Then
        ws.Cells(row, col).Value = CalcTotalEAOAdj(wein)
    End If
End Sub
