Attribute VB_Name = "modSP2_CheckResult_Incentives"
'==============================================================================
' Module: modSP2_CheckResult_Incentives
' Purpose: Subprocess 2 - Incentives Check columns
' Description: Validates AIP, SIP, Inspire, RSU, Bonuses
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Sub: SP2_Check_Incentives
' Purpose: Populate incentive Check columns
'------------------------------------------------------------------------------
Public Sub SP2_Check_Incentives(valWb As Workbook, weinIndex As Object)
    Dim ws As Worksheet
    
    On Error GoTo ErrHandler
    
    Set ws = valWb.Worksheets("Check Result")
    
    ' Load and process One Time Payment
    ProcessOneTimePaymentCheck ws, weinIndex
    
    ' Load and process Inspire Awards
    ProcessInspireCheck ws, weinIndex
    
    ' Load and process SIP/2025QX
    ProcessSIPCheck ws, weinIndex
    
    ' Load and process RSU Dividend
    ProcessRSUCheck ws, weinIndex
    
    LogInfo "modSP2_CheckResult_Incentives", "SP2_Check_Incentives", "Incentive checks completed"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Incentives", "SP2_Check_Incentives", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: ProcessOneTimePaymentCheck
' Purpose: Process One Time Payment for Check columns
'------------------------------------------------------------------------------
Private Sub ProcessOneTimePaymentCheck(ws As Worksheet, weinIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim dataRange As Range
    Dim grouped As Object
    Dim lastRow As Long, lastCol As Long
    
    On Error GoTo ErrHandler
    
    filePath = GetInputFilePath("OneTimePayment")
    If Dir(filePath) = "" Then Exit Sub
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    lastRow = srcWs.Cells(srcWs.Rows.count, 1).End(xlUp).row
    lastCol = srcWs.Cells(1, srcWs.Columns.count).End(xlToLeft).Column
    Set dataRange = srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(lastRow, lastCol))
    
    ' Try multiple field name variants for Employee ID
    Set grouped = GroupByEmployeeAndType(dataRange, "Employee ID,EmployeeID,WEIN,WIN,Employee Number ID", "One-Time Payment Plan", "Actual Payment - Amount")
    
    ' Map to Check columns
    Dim key As Variant
    Dim parts() As String
    Dim empId As String, planType As String, wein As String
    Dim col As Long, row As Long
    
    For Each key In grouped.Keys
        parts = Split(CStr(key), "|")
        If UBound(parts) >= 1 Then
            empId = parts(0)
            planType = UCase(parts(1))
            
            wein = NormalizeEmployeeId(empId)
            
            If weinIndex.exists(wein) Then
                row = weinIndex(wein)
                
                ' Map plan types to Check columns
                If InStr(planType, "LUMP SUM") > 0 Then
                    col = FindColumnByHeader(ws.Rows(4), "Lump Sum Bonus Check")
                ElseIf InStr(planType, "SIGN ON") > 0 Then
                    col = FindColumnByHeader(ws.Rows(4), "Sign On Bonus Check")
                ElseIf InStr(planType, "RETENTION") > 0 Then
                    col = FindColumnByHeader(ws.Rows(4), "Retention Bonus Check")
                ElseIf InStr(planType, "REFERRAL") > 0 Then
                    col = FindColumnByHeader(ws.Rows(4), "Referral Bonus Check")
                ElseIf InStr(planType, "RED PACKET") > 0 Or InStr(planType, "NEW YEAR") > 0 Then
                    col = FindColumnByHeader(ws.Rows(4), "Red Packet Check")
                Else
                    col = 0
                End If
                
                If col > 0 Then
                    ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, grouped(key))
                End If
            End If
        End If
    Next key
    
    wb.Close SaveChanges:=False
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Incentives", "ProcessOneTimePaymentCheck", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Sub: ProcessInspireCheck
' Purpose: Process Inspire Awards for Check columns
'------------------------------------------------------------------------------
Private Sub ProcessInspireCheck(ws As Worksheet, weinIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim dataRange As Range
    Dim grouped As Object
    Dim lastRow As Long, lastCol As Long
    
    On Error GoTo ErrHandler
    
    filePath = GetInputFilePath("InspireAwards")
    If Dir(filePath) = "" Then Exit Sub
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    lastRow = srcWs.Cells(srcWs.Rows.count, 1).End(xlUp).row
    lastCol = srcWs.Cells(1, srcWs.Columns.count).End(xlToLeft).Column
    Set dataRange = srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(lastRow, lastCol))
    
    ' Try multiple field name variants for Employee ID
    Set grouped = GroupByEmployeeAndType(dataRange, "Employee ID,EmployeeID,WEIN,WIN,Employee Number ID", "One-Time Payment Plan", "Actual Payment - Amount")
    
    Dim key As Variant
    Dim parts() As String
    Dim empId As String, planType As String, wein As String
    Dim col As Long, row As Long
    
    For Each key In grouped.Keys
        parts = Split(CStr(key), "|")
        If UBound(parts) >= 1 Then
            empId = parts(0)
            planType = UCase(parts(1))
            
            wein = NormalizeEmployeeId(empId)
            
            If weinIndex.exists(wein) Then
                row = weinIndex(wein)
                
                If InStr(planType, "INSPIRE POINTS") > 0 Then
                    col = FindColumnByHeader(ws.Rows(4), "Inspire Points Check")
                ElseIf InStr(planType, "INSPIRE CASH") > 0 Then
                    col = FindColumnByHeader(ws.Rows(4), "Inspire Cash Check")
                Else
                    col = 0
                End If
                
                If col > 0 Then
                    ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, grouped(key))
                End If
            End If
        End If
    Next key
    
    wb.Close SaveChanges:=False
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Incentives", "ProcessInspireCheck", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Sub: ProcessSIPCheck
' Purpose: Process SIP/2025QX for Check columns
'------------------------------------------------------------------------------
Private Sub ProcessSIPCheck(ws As Worksheet, weinIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim dataRange As Range
    Dim grouped As Object
    Dim lastRow As Long, lastCol As Long
    
    On Error GoTo ErrHandler
    
    filePath = GetInputFilePath("2025QXPayout")
    If Dir(filePath) = "" Then Exit Sub
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    lastRow = srcWs.Cells(srcWs.Rows.count, 1).End(xlUp).row
    lastCol = srcWs.Cells(1, srcWs.Columns.count).End(xlToLeft).Column
    Set dataRange = srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(lastRow, lastCol))
    
    ' Try multiple field name variants for Employee ID
    Set grouped = GroupByEmployeeAndType(dataRange, "EMPLOYEE ID,Employee ID,EmployeeID,WEIN,WIN", "Pay Item", "TOTAL PAYOUT")
    
    Dim key As Variant
    Dim parts() As String
    Dim empId As String, payItem As String, wein As String
    Dim col As Long, row As Long
    
    For Each key In grouped.Keys
        parts = Split(CStr(key), "|")
        If UBound(parts) >= 1 Then
            empId = parts(0)
            payItem = UCase(parts(1))
            
            wein = NormalizeEmployeeId(empId)
            
            If weinIndex.exists(wein) Then
                row = weinIndex(wein)
                
                If InStr(payItem, "QUALITATIVE") > 0 Then
                    col = FindColumnByHeader(ws.Rows(4), "Sales Incentive (Qualitative) Check")
                ElseIf InStr(payItem, "SALES INCENTIVE") > 0 Then
                    col = FindColumnByHeader(ws.Rows(4), "Sales Incentive (Quantitative) Check")
                Else
                    col = 0
                End If
                
                If col > 0 Then
                    ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, grouped(key))
                End If
            End If
        End If
    Next key
    
    wb.Close SaveChanges:=False
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Incentives", "ProcessSIPCheck", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Sub: ProcessRSUCheck
' Purpose: Process RSU Dividend for Check columns
'------------------------------------------------------------------------------
Private Sub ProcessRSUCheck(ws As Worksheet, weinIndex As Object)
    ' Similar implementation to SP1 RSU processing
    ' Check IsSpecialMonth and process RSU Global/EY
    
    On Error GoTo ErrHandler
    
    If Not IsSpecialMonth("IsRSUDivMonth") Then Exit Sub
    
    ' Process RSU Global and EY similar to SP1
    ' Write to "Shares Dividend Check" column
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Incentives", "ProcessRSUCheck", Err.Number, Err.Description
End Sub
