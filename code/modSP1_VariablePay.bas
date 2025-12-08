Attribute VB_Name = "modSP1_VariablePay"
'==============================================================================
' Module: modSP1_VariablePay
' Purpose: Subprocess 1 - VariablePay sheet population
' Description: Handles variable pay items from multiple data sources
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Sub: SP1_PopulateVariablePay
' Purpose: Main routine to populate VariablePay sheet
'------------------------------------------------------------------------------
Public Sub SP1_PopulateVariablePay(flexWb As Workbook)
    Dim ws As Worksheet
    Dim empIndex As Object
    
    On Error GoTo ErrHandler
    
    LogInfo "modSP1_VariablePay", "SP1_PopulateVariablePay", "Starting VariablePay population"
    
    Set ws = flexWb.Worksheets("VariablePay")
    
    ' Build employee index (try multiple field name variants)
    Set empIndex = BuildEmployeeIndex(ws, "Employee Code,EmployeeCode,Employee Reference,EmployeeNumber,Employee Number")
    
    ' Load EAO data
    LoadEAOData
    
    ' Process each data source
    ProcessOneTimePayment ws, empIndex
    ProcessInspireAwards ws, empIndex
    ProcessSIPQIP ws, empIndex
    ProcessRSUDividend ws, empIndex
    ProcessAIPPayouts ws, empIndex
    ProcessFlexClaim ws, empIndex
    ProcessMerckPayrollSummary ws, empIndex
    ProcessExtraTable ws, empIndex
    
    LogInfo "modSP1_VariablePay", "SP1_PopulateVariablePay", "VariablePay population completed"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_VariablePay", "SP1_PopulateVariablePay", Err.Number, Err.Description
End Sub


'------------------------------------------------------------------------------
' Sub: ProcessOneTimePayment
' Purpose: Process One Time Payment report
'------------------------------------------------------------------------------
Private Sub ProcessOneTimePayment(ws As Worksheet, empIndex As Object)
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
    
    ' Group by Employee ID and One-Time Payment Plan (try multiple field name variants)
    Set grouped = GroupByEmployeeAndType(dataRange, "Employee ID,EmployeeID,WEIN,WIN,Employee Number ID", "One-Time Payment Plan", "Actual Payment - Amount")
    
    ' Map to VariablePay columns
    Dim planMapping As Object
    Set planMapping = CreateObject("Scripting.Dictionary")
    planMapping.Add "LUMP SUM MERIT", "Lump Sum Bonus"
    planMapping.Add "SIGN ON BONUS", "Sign On Bonus"
    planMapping.Add "RETENTION BONUS", "Retention Bonus"
    planMapping.Add "REFERRAL PAYMENT", "Referral Bonus"
    planMapping.Add "MANAGER OF THE YEAR AWARD", "Manager of the Year Award"
    planMapping.Add "MD AWARD", "MD Award"
    planMapping.Add "EMPLOYEE AWARD", "Employee Award"
    planMapping.Add "NEW YEAR'S ALLOWANCE", "Red Packet"
    planMapping.Add "RED PACKET", "Red Packet"
    planMapping.Add "CASH AWARD", "Other Allowance"
    planMapping.Add "SIP TO AIP TRANSITION", "Other Allowance"
    
    ' Write values
    Dim key As Variant
    Dim parts() As String
    Dim empId As String, planType As String, wein As String
    Dim targetCol As String, col As Long, row As Long
    
    For Each key In grouped.Keys
        parts = Split(CStr(key), "|")
        If UBound(parts) >= 1 Then
            empId = parts(0)
            planType = UCase(parts(1))
            
            ' Skip Inspire (handled separately)
            If InStr(planType, "INSPIRE") > 0 Then GoTo NextKey
            
            wein = NormalizeEmployeeId(empId)
            
            ' Find target column
            targetCol = ""
            Dim pm As Variant
            For Each pm In planMapping.Keys
                If InStr(planType, pm) > 0 Then
                    targetCol = planMapping(pm)
                    Exit For
                End If
            Next pm
            
            If targetCol <> "" Then
                col = FindColumnByHeader(ws.Rows(1), targetCol)
                If col > 0 Then
                    row = GetOrAddRow(ws, wein, empIndex)
                    If row > 0 Then
                        ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, grouped(key))
                    End If
                End If
            End If
        End If
NextKey:
    Next key
    
    wb.Close SaveChanges:=False
    LogInfo "modSP1_VariablePay", "ProcessOneTimePayment", "Processed One Time Payment"
    Exit Sub
    
ErrHandler:
    LogError "modSP1_VariablePay", "ProcessOneTimePayment", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Sub: ProcessInspireAwards
' Purpose: Process Inspire Awards payroll report
'------------------------------------------------------------------------------
Private Sub ProcessInspireAwards(ws As Worksheet, empIndex As Object)
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
            
            row = GetOrAddRow(ws, wein, empIndex)
            If row > 0 Then
                If InStr(planType, "INSPIRE POINTS") > 0 Then
                    col = FindColumnByHeader(ws.Rows(1), "Inspire Points")
                    If col > 0 Then ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, grouped(key))
                ElseIf InStr(planType, "INSPIRE CASH") > 0 Then
                    col = FindColumnByHeader(ws.Rows(1), "Inspire Cash")
                    If col > 0 Then ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, grouped(key))
                End If
            End If
        End If
    Next key
    
    wb.Close SaveChanges:=False
    LogInfo "modSP1_VariablePay", "ProcessInspireAwards", "Processed Inspire Awards"
    Exit Sub
    
ErrHandler:
    LogError "modSP1_VariablePay", "ProcessInspireAwards", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub


'------------------------------------------------------------------------------
' Sub: ProcessSIPQIP
' Purpose: Process SIP QIP report
'------------------------------------------------------------------------------
Private Sub ProcessSIPQIP(ws As Worksheet, empIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim dataRange As Range
    Dim grouped As Object
    Dim lastRow As Long, lastCol As Long
    
    On Error GoTo ErrHandler
    
    filePath = GetInputFilePath("SIPQIP")
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
            
            row = GetOrAddRow(ws, wein, empIndex)
            If row > 0 Then
                If InStr(payItem, "QUALITATIVE") > 0 Then
                    col = FindColumnByHeader(ws.Rows(1), "Sales Incentive (Qualitative)")
                    If col > 0 Then ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, grouped(key))
                ElseIf InStr(payItem, "SALES INCENTIVE") > 0 Then
                    col = FindColumnByHeader(ws.Rows(1), "Sales Incentive (Quantitative)")
                    If col > 0 Then ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, grouped(key))
                End If
            End If
        End If
    Next key
    
    wb.Close SaveChanges:=False
    LogInfo "modSP1_VariablePay", "ProcessSIPQIP", "Processed SIP QIP"
    Exit Sub
    
ErrHandler:
    LogError "modSP1_VariablePay", "ProcessSIPQIP", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Sub: ProcessRSUDividend
' Purpose: Process RSU Dividend reports (Global and EY)
'------------------------------------------------------------------------------
Private Sub ProcessRSUDividend(ws As Worksheet, empIndex As Object)
    On Error GoTo ErrHandler
    
    ' Check if this is RSU month
    If Not IsSpecialMonth("IsRSUDivMonth") Then
        LogInfo "modSP1_VariablePay", "ProcessRSUDividend", "Not RSU month, skipping"
        Exit Sub
    End If
    
    ' Process RSU Global
    ProcessRSUGlobal ws, empIndex
    
    ' Process RSU EY
    ProcessRSUEY ws, empIndex
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_VariablePay", "ProcessRSUDividend", Err.Number, Err.Description
End Sub

Private Sub ProcessRSUGlobal(ws As Worksheet, empIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim lastRow As Long, i As Long
    Dim empRef As String, wein As String
    Dim grossAmt As Double, fxRate As Double
    Dim col As Long, row As Long
    Dim empRefCol As Long, amtCol As Long
    
    On Error GoTo ErrHandler
    
    filePath = GetInputFilePath("RSUGlobal")
    If Dir(filePath) = "" Then Exit Sub
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    fxRate = GetExchangeRate("RSU_Global")
    
    ' Try multiple field name variants for Employee Reference
    empRefCol = FindColumnByHeader(srcWs.Rows(1), "Employee Reference,EmployeeNumber,Employee Number,Employee ID,EmployeeID")
    amtCol = FindColumnByHeader(srcWs.Rows(1), "Gross Award Amount to be Paid")
    
    If empRefCol = 0 Or amtCol = 0 Then
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    col = FindColumnByHeader(ws.Rows(1), "Shares Dividend")
    If col = 0 Then
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    lastRow = srcWs.Cells(srcWs.Rows.count, empRefCol).End(xlUp).row
    
    For i = 2 To lastRow
        empRef = Trim(CStr(Nz(srcWs.Cells(i, empRefCol).Value, "")))
        grossAmt = ToDouble(srcWs.Cells(i, amtCol).Value)
        
        If empRef <> "" And grossAmt <> 0 Then
            wein = NormalizeEmployeeId(empRef)
            
            row = GetOrAddRow(ws, wein, empIndex)
            If row > 0 Then
                ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, grossAmt * fxRate)
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    LogInfo "modSP1_VariablePay", "ProcessRSUGlobal", "Processed RSU Global"
    Exit Sub
    
ErrHandler:
    LogError "modSP1_VariablePay", "ProcessRSUGlobal", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

Private Sub ProcessRSUEY(ws As Worksheet, empIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim lastRow As Long, i As Long
    Dim empNum As String, wein As String
    Dim divAmt As Double, fxRate As Double
    Dim col As Long, row As Long
    Dim empNumCol As Long, amtCol As Long
    
    On Error GoTo ErrHandler
    
    filePath = GetInputFilePath("RSUEY")
    If Dir(filePath) = "" Then Exit Sub
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    fxRate = GetExchangeRate("RSU_EY")
    
    ' Try multiple field name variants for EmployeeNumber
    empNumCol = FindColumnByHeader(srcWs.Rows(1), "EmployeeNumber,Employee Number,Employee ID,EmployeeID,Employee Reference")
    amtCol = FindColumnByHeader(srcWs.Rows(1), "Dividend To Pay")
    
    If empNumCol = 0 Or amtCol = 0 Then
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    col = FindColumnByHeader(ws.Rows(1), "Shares Dividend")
    If col = 0 Then
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    lastRow = srcWs.Cells(srcWs.Rows.count, empNumCol).End(xlUp).row
    
    For i = 2 To lastRow
        empNum = Trim(CStr(Nz(srcWs.Cells(i, empNumCol).Value, "")))
        divAmt = ToDouble(srcWs.Cells(i, amtCol).Value)
        
        If empNum <> "" And divAmt <> 0 Then
            wein = NormalizeEmployeeId(empNum)
            
            row = GetOrAddRow(ws, wein, empIndex)
            If row > 0 Then
                ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, divAmt * fxRate)
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    LogInfo "modSP1_VariablePay", "ProcessRSUEY", "Processed RSU EY"
    Exit Sub
    
ErrHandler:
    LogError "modSP1_VariablePay", "ProcessRSUEY", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub


'------------------------------------------------------------------------------
' Sub: ProcessAIPPayouts
' Purpose: Process AIP Payouts Payroll Report
'------------------------------------------------------------------------------
Private Sub ProcessAIPPayouts(ws As Worksheet, empIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim lastRow As Long, i As Long
    Dim wein As String
    Dim bonusAmt As Double
    Dim col As Long, row As Long
    Dim weinCol As Long, amtCol As Long
    
    On Error GoTo ErrHandler
    
    ' Check if this is AIP month
    If Not IsSpecialMonth("IsAIPMonth") Then
        LogInfo "modSP1_VariablePay", "ProcessAIPPayouts", "Not AIP month, skipping"
        Exit Sub
    End If
    
    filePath = GetInputFilePath("AIPPayouts")
    If Dir(filePath) = "" Then Exit Sub
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    ' Try multiple field name variants for WEIN
    weinCol = FindColumnByHeader(srcWs.Rows(1), "WIN,WEIN,WEINEmployee ID,Employee CodeWIN,Employee ID,EmployeeID")
    amtCol = FindColumnByHeader(srcWs.Rows(1), "Bonus Amount")
    
    If weinCol = 0 Or amtCol = 0 Then
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    col = FindColumnByHeader(ws.Rows(1), "Annual Incentive")
    If col = 0 Then
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    lastRow = srcWs.Cells(srcWs.Rows.count, weinCol).End(xlUp).row
    
    For i = 2 To lastRow
        wein = Trim(CStr(Nz(srcWs.Cells(i, weinCol).Value, "")))
        bonusAmt = ToDouble(srcWs.Cells(i, amtCol).Value)
        
        If wein <> "" And bonusAmt <> 0 Then
            row = GetOrAddRow(ws, wein, empIndex)
            If row > 0 Then
                ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, bonusAmt)
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    LogInfo "modSP1_VariablePay", "ProcessAIPPayouts", "Processed AIP Payouts"
    Exit Sub
    
ErrHandler:
    LogError "modSP1_VariablePay", "ProcessAIPPayouts", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Sub: ProcessFlexClaim
' Purpose: Process MSD HK Flex Claim Summary Report
'------------------------------------------------------------------------------
Private Sub ProcessFlexClaim(ws As Worksheet, empIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim lastRow As Long, i As Long
    Dim empNumId As String, wein As String
    Dim transAmt As Double, claimStatus As String
    Dim col As Long, row As Long
    Dim empNumCol As Long, amtCol As Long, statusCol As Long
    
    On Error GoTo ErrHandler
    
    ' Check if this is Flex Benefit month
    If Not IsSpecialMonth("IsFlexBenefitMonth") Then
        LogInfo "modSP1_VariablePay", "ProcessFlexClaim", "Not Flex Benefit month, skipping"
        Exit Sub
    End If
    
    filePath = GetInputFilePath("FlexClaim")
    If Dir(filePath) = "" Then Exit Sub
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    ' Try multiple field name variants for Employee Number ID
    empNumCol = FindColumnByHeader(srcWs.Rows(1), "Employee Number ID,EmployeeNumber,Employee Number,Employee ID,EmployeeID")
    amtCol = FindColumnByHeader(srcWs.Rows(1), "Transacted Amount")
    statusCol = FindColumnByHeader(srcWs.Rows(1), "Claim Status")
    
    If empNumCol = 0 Or amtCol = 0 Then
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    col = FindColumnByHeader(ws.Rows(1), "Flexible benefits")
    If col = 0 Then
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    lastRow = srcWs.Cells(srcWs.Rows.count, empNumCol).End(xlUp).row
    
    For i = 2 To lastRow
        ' Filter by Approved status
        If statusCol > 0 Then
            claimStatus = UCase(Trim(CStr(Nz(srcWs.Cells(i, statusCol).Value, ""))))
            If claimStatus <> "APPROVED" Then GoTo NextRow
        End If
        
        empNumId = Trim(CStr(Nz(srcWs.Cells(i, empNumCol).Value, "")))
        transAmt = ToDouble(srcWs.Cells(i, amtCol).Value)
        
        If empNumId <> "" And transAmt <> 0 Then
            wein = NormalizeEmployeeId(empNumId)
            
            row = GetOrAddRow(ws, wein, empIndex)
            If row > 0 Then
                ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, transAmt)
            End If
        End If
NextRow:
    Next i
    
    wb.Close SaveChanges:=False
    LogInfo "modSP1_VariablePay", "ProcessFlexClaim", "Processed Flex Claim"
    Exit Sub
    
ErrHandler:
    LogError "modSP1_VariablePay", "ProcessFlexClaim", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Sub: ProcessMerckPayrollSummary
' Purpose: Process Merck Payroll Summary Report for IA Pay Split
'------------------------------------------------------------------------------
Private Sub ProcessMerckPayrollSummary(ws As Worksheet, empIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim lastRow As Long, i As Long
    Dim empId As String, wein As String
    Dim netPay As Double, mpfEEMC As Double, mpfEEVC As Double
    Dim iaPaySplit As Double, mpfRI As Double, mpfVCRI As Double
    Dim row As Long
    Dim empIdCol As Long, netPayCol As Long, mpfRICol As Long
    Dim mpfVCRICol As Long, mpfEEMCCol As Long, mpfEEVCCol As Long
    Dim colIAPaySplit As Long, colMPFRI As Long, colMPFVCRI As Long
    
    On Error GoTo ErrHandler
    
    filePath = GetInputFilePath("MerckPayroll")
    If Dir(filePath) = "" Then Exit Sub
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    ' Find columns (try multiple field name variants)
    empIdCol = FindColumnByHeader(srcWs.Rows(1), "Employee ID,EmployeeID,WEIN,WIN,Employee Number ID")
    netPayCol = FindColumnByHeader(srcWs.Rows(1), "Net Pay (include EAO & leave payment)")
    mpfRICol = FindColumnByHeader(srcWs.Rows(1), "MPF Relevant Income")
    mpfVCRICol = FindColumnByHeader(srcWs.Rows(1), "MPF VC Relevant Income")
    mpfEEMCCol = FindColumnByHeader(srcWs.Rows(1), "MPF EE MC")
    mpfEEVCCol = FindColumnByHeader(srcWs.Rows(1), "MPF EE VC")
    
    If empIdCol = 0 Then
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Find target columns
    colIAPaySplit = FindColumnByHeader(ws.Rows(1), "IA Pay Split")
    colMPFRI = FindColumnByHeader(ws.Rows(1), "MPF Relevant Income Rewrite")
    colMPFVCRI = FindColumnByHeader(ws.Rows(1), "MPF VC Relevant Income Rewrite")
    
    lastRow = srcWs.Cells(srcWs.Rows.count, empIdCol).End(xlUp).row
    
    For i = 2 To lastRow
        empId = Trim(CStr(Nz(srcWs.Cells(i, empIdCol).Value, "")))
        
        If empId <> "" Then
            wein = NormalizeEmployeeId(empId)
            
            row = GetOrAddRow(ws, wein, empIndex)
            If row > 0 Then
                ' Calculate IA Pay Split
                If netPayCol > 0 And mpfEEMCCol > 0 And mpfEEVCCol > 0 Then
                    netPay = ToDouble(srcWs.Cells(i, netPayCol).Value)
                    mpfEEMC = ToDouble(srcWs.Cells(i, mpfEEMCCol).Value)
                    mpfEEVC = ToDouble(srcWs.Cells(i, mpfEEVCCol).Value)
                    iaPaySplit = RoundAmount2(netPay + mpfEEMC + mpfEEVC)
                    
                    If colIAPaySplit > 0 Then
                        ws.Cells(row, colIAPaySplit).Value = iaPaySplit
                    End If
                End If
                
                ' MPF Relevant Income Rewrite
                If mpfRICol > 0 And colMPFRI > 0 Then
                    mpfRI = ToDouble(srcWs.Cells(i, mpfRICol).Value)
                    ws.Cells(row, colMPFRI).Value = RoundAmount2(mpfRI)
                End If
                
                ' MPF VC Relevant Income Rewrite
                If mpfVCRICol > 0 And colMPFVCRI > 0 Then
                    mpfVCRI = ToDouble(srcWs.Cells(i, mpfVCRICol).Value)
                    ws.Cells(row, colMPFVCRI).Value = RoundAmount2(mpfVCRI)
                End If
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    LogInfo "modSP1_VariablePay", "ProcessMerckPayrollSummary", "Processed Merck Payroll Summary"
    Exit Sub
    
ErrHandler:
    LogError "modSP1_VariablePay", "ProcessMerckPayrollSummary", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Sub: ProcessExtraTable
' Purpose: Process 额外表 for PPTO EAO Rate and special bonuses
'------------------------------------------------------------------------------
Private Sub ProcessExtraTable(ws As Worksheet, empIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim lastRow As Long, i As Long
    Dim wein As String
    Dim pptoRate As Double, flexBenefit As Double
    Dim row As Long
    Dim weinCol As Long, pptoCol As Long, flexCol As Long
    Dim colPPTORate As Long, colFlexBenefit As Long
    
    On Error GoTo ErrHandler
    
    Set wb = OpenExtraTableWorkbook()
    If wb Is Nothing Then Exit Sub
    
    ' Process [需要每月维护] sheet for PPTO EAO Rate
    On Error Resume Next
    Set srcWs = wb.Worksheets("需要每月维护")
    On Error GoTo ErrHandler
    
    If Not srcWs Is Nothing Then
        ' Try multiple field name variants for WEIN
        weinCol = FindColumnByHeader(srcWs.Rows(1), "WEIN,WIN,WEINEmployee ID,Employee CodeWIN,Employee ID,EmployeeID")
        pptoCol = FindColumnByHeader(srcWs.Rows(1), "PPTO EAO Rate input")
        
        colPPTORate = FindColumnByHeader(ws.Rows(1), "PPTO EAO Rate input")
        
        If weinCol > 0 And pptoCol > 0 And colPPTORate > 0 Then
            lastRow = srcWs.Cells(srcWs.Rows.count, weinCol).End(xlUp).row
            
            For i = 2 To lastRow
                wein = Trim(CStr(Nz(srcWs.Cells(i, weinCol).Value, "")))
                pptoRate = ToDouble(srcWs.Cells(i, pptoCol).Value)
                
                If wein <> "" Then
                    row = GetOrAddRow(ws, wein, empIndex)
                    If row > 0 Then
                        ws.Cells(row, colPPTORate).Value = RoundAmount2(pptoRate)
                    End If
                End If
            Next i
        End If
    End If
    
    ' Process [特殊奖金] sheet for Flexible benefits
    On Error Resume Next
    Set srcWs = wb.Worksheets("特殊奖金")
    On Error GoTo ErrHandler
    
    If Not srcWs Is Nothing Then
        ' Try multiple field name variants for WEIN
        weinCol = FindColumnByHeader(srcWs.Rows(1), "WEIN,WIN,WEINEmployee ID,Employee CodeWIN,Employee ID,EmployeeID")
        flexCol = FindColumnByHeader(srcWs.Rows(1), "Flexible benefits")
        
        colFlexBenefit = FindColumnByHeader(ws.Rows(1), "Flexible benefits")
        
        If weinCol > 0 And flexCol > 0 And colFlexBenefit > 0 Then
            lastRow = srcWs.Cells(srcWs.Rows.count, weinCol).End(xlUp).row
            
            For i = 2 To lastRow
                wein = Trim(CStr(Nz(srcWs.Cells(i, weinCol).Value, "")))
                flexBenefit = ToDouble(srcWs.Cells(i, flexCol).Value)
                
                If wein <> "" And flexBenefit <> 0 Then
                    row = GetOrAddRow(ws, wein, empIndex)
                    If row > 0 Then
                        ws.Cells(row, colFlexBenefit).Value = SafeAdd2(ws.Cells(row, colFlexBenefit).Value, flexBenefit)
                    End If
                End If
            Next i
        End If
    End If
    
    LogInfo "modSP1_VariablePay", "ProcessExtraTable", "Processed Extra Table"
    Exit Sub
    
ErrHandler:
    LogError "modSP1_VariablePay", "ProcessExtraTable", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Helper: GetOrAddRow
'------------------------------------------------------------------------------
Private Function GetOrAddRow(ws As Worksheet, wein As String, empIndex As Object) As Long
    Dim empCode As String
    Dim empCodeCol As Long
    Dim newRow As Long
    
    ' Try WEIN first
    If empIndex.exists(wein) Then
        GetOrAddRow = empIndex(wein)
        Exit Function
    End If
    
    ' Try Employee Code
    empCode = EmpCodeFromWein(wein)
    If empCode <> "" And empIndex.exists(empCode) Then
        GetOrAddRow = empIndex(empCode)
        Exit Function
    End If
    
    ' Add new row - try multiple field name variants
    empCodeCol = FindColumnByHeader(ws.Rows(1), "Employee Code,EmployeeCode,Employee Reference,EmployeeNumber,Employee Number")
    If empCodeCol = 0 Then empCodeCol = 1
    
    newRow = ws.Cells(ws.Rows.count, empCodeCol).End(xlUp).row + 1
    
    If empCode <> "" Then
        ws.Cells(newRow, empCodeCol).Value = empCode
        empIndex.Add empCode, newRow
    Else
        ws.Cells(newRow, empCodeCol).Value = wein
        empIndex.Add wein, newRow
    End If
    
    GetOrAddRow = newRow
End Function

