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
    
    ' Write PPTO EAO Rate from 额外表
    WritePPTOEAORateCheck ws, weinIndex
    
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
    
    On Error Resume Next
    
    ' Base Pay 60001000 Check
    col = GetCheckColIndex("Base Pay 60001000")
    If col > 0 Then
        ' Formula: Actual working days * Monthly Salary / calendar days
        ' This would need actual working days from another source
        ' Placeholder implementation
    End If
    
    ' Base Pay(Temp) 60101000 Check
    col = GetCheckColIndex("Base Pay(Temp) 60101000")
    If col > 0 Then
        ' Placeholder implementation for temp employees
    End If
    
    ' Salary Adj 60001000 Check
    col = GetCheckColIndex("Salary Adj 60001000")
    If col > 0 Then
        ' Placeholder implementation
    End If
    
    ' Transport Allowance 60409960 Check
    col = GetCheckColIndex("Transport Allowance 60409960")
    If col > 0 Then
        ' Placeholder implementation
    End If
    
    ' Transport Allowance Adj 60409960 Check
    col = GetCheckColIndex("Transport Allowance Adj 60409960")
    If col > 0 Then
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
    col = GetCheckColIndex("Maternity Leave Payment 60001000")
    If col > 0 Then
        ws.Cells(row, col).Value = CalcMaternityLeavePayment(wein)
    End If
    
    ' Sick Leave Payment Check
    col = GetCheckColIndex("Sick Leave Payment 60001000")
    If col > 0 Then
        ws.Cells(row, col).Value = CalcSickLeavePayment(wein)
    End If
    
    ' PPTO Payment Check
    col = GetCheckColIndex("Paid Parental Time Off (PPTO) payment")
    If col > 0 Then
        ws.Cells(row, col).Value = CalcPPTOPayment(wein)
    End If
    
    ' No Pay Leave Deduction Check
    col = GetCheckColIndex("No Pay Leave Deduction 60001000")
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
    col = GetCheckColIndex("Total EAO Adj 60409960")
    If col > 0 Then
        ws.Cells(row, col).Value = CalcTotalEAOAdj(wein)
    End If
End Sub


'------------------------------------------------------------------------------
' Sub: WritePPTOEAORateCheck
' Purpose: Write PPTO EAO Rate input Check column from 额外表
'------------------------------------------------------------------------------
Private Sub WritePPTOEAORateCheck(ws As Worksheet, weinIndex As Object)
    Dim extraWb As Workbook
    Dim srcWs As Worksheet
    Dim lastRow As Long, i As Long
    Dim headers As Object
    Dim wein As String
    Dim row As Long, col As Long
    Dim pptoRate As Double
    Dim headerRow As Long, keyCol As Long
    
    On Error GoTo ErrHandler
    
    col = GetCheckColIndex("PPTO EAO Rate input")
    If col = 0 Then Exit Sub
    
    Set extraWb = OpenExtraTableWorkbook()
    If extraWb Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set srcWs = extraWb.Worksheets("特殊奖金")
    On Error GoTo ErrHandler
    
    If srcWs Is Nothing Then Exit Sub
    
    ' Detect header row and build header index
    headerRow = FindHeaderRowSafe(srcWs, "WEIN,WIN", 1, 50)
    Set headers = BuildHeaderIndex(srcWs, headerRow)

    keyCol = GetColumnFromHeaders(headers, "WEIN,WIN")
    If keyCol = 0 Then keyCol = 1
    lastRow = srcWs.Cells(srcWs.Rows.count, keyCol).End(xlUp).Row

    For i = headerRow + 1 To lastRow
        ' Get WEIN
        wein = GetPPTOCellVal(srcWs, i, headers, "WEIN")
        If wein = "" Then wein = GetPPTOCellVal(srcWs, i, headers, "WIN")
        
        If wein <> "" And weinIndex.exists(wein) Then
            row = weinIndex(wein)
            pptoRate = ToDouble(GetPPTOCellVal(srcWs, i, headers, "PPTO EAO RATE INPUT"))
            If pptoRate > 0 Then
                ws.Cells(row, col).Value = pptoRate
            End If
        End If
    Next i
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_PayItems", "WritePPTOEAORateCheck", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Helper: GetPPTOCellVal
'------------------------------------------------------------------------------
Private Function GetPPTOCellVal(ws As Worksheet, row As Long, headers As Object, headerName As String) As String
    Dim col As Long
    GetPPTOCellVal = ""
    
    If headers.exists(UCase(headerName)) Then
        col = headers(UCase(headerName))
        GetPPTOCellVal = Trim(CStr(Nz(ws.Cells(row, col).Value, "")))
    End If
End Function

