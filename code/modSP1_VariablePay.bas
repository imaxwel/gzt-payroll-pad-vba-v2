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
' Filter Logic:
'   1. Exclude records where One-Time Payment Plan is "Inspire Points Value" or "Inspire Cash Award"
'   2. Keep only records where:
'      - Completed On > Previous Cutoff AND Completed On <= Current Cutoff
'      - Scheduled Payment Date is in the previous month
'   3. Group by Employee ID + One-Time Payment Plan, sum Actual Payment - Amount
'------------------------------------------------------------------------------
Private Sub ProcessOneTimePayment(ws As Worksheet, empIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim dataRange As Range
    Dim grouped As Object
    Dim lastRow As Long, lastCol As Long
    Dim excludeTypes As Variant
    
    On Error GoTo ErrHandler
    
    filePath = GetInputFilePath("OneTimePayment")
    If Dir(filePath) = "" Then Exit Sub
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    lastRow = srcWs.Cells(srcWs.Rows.count, 1).End(xlUp).row
    lastCol = srcWs.Cells(1, srcWs.Columns.count).End(xlToLeft).Column
    Set dataRange = srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(lastRow, lastCol))
    
    ' Types to exclude (Inspire handled separately)
    excludeTypes = Array("Inspire Points Value", "Inspire Cash Award")
    
    ' Group by Employee ID and One-Time Payment Plan with date filtering
    ' Filter: Completed On between (PreviousCutoff, CurrentCutoff]
    '         Scheduled Payment Date in previous month
    Set grouped = GroupByEmployeeAndTypeWithDateFilter( _
        dataRange, _
        "Employee ID,EmployeeID,WEIN,WIN,Employee Number ID", _
        "One-Time Payment Plan", _
        "Actual Payment - Amount", _
        "Completed On,CompletedOn,Completed Date", _
        "Scheduled Payment Date,ScheduledPaymentDate,Scheduled Pay Date", _
        G.Payroll.PreviousCutoff, _
        G.Payroll.currentCutoff, _
        G.Payroll.prevMonthStart, _
        G.Payroll.prevMonthEnd, _
        excludeTypes)
    
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
    ' Note: RED PACKET type is handled separately in Check Result only (not VariablePay)
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
            
            ' Note: Inspire types already excluded by GroupByEmployeeAndTypeWithDateFilter
            
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
                If InStr(payItem, "QUALITATIVE INCENTIVE PLAN") > 0 Then
                    col = FindColumnByHeader(ws.Rows(1), "Sales Incentive (Qualitative)")
                    If col > 0 Then ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, grouped(key))
                ElseIf InStr(payItem, "SALES INCENTIVE PLAN") > 0 Then
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
' Note: RSU Global is processed in May only, RSU EY is processed in June only
'------------------------------------------------------------------------------
Private Sub ProcessRSUDividend(ws As Worksheet, empIndex As Object)
    Dim currentMonth As Integer
    
    On Error GoTo ErrHandler
    
    ' Get current payroll month
    currentMonth = Month(G.Payroll.monthStart)
    
    ' May: Process RSU Global only
    If currentMonth = 5 Then
        LogInfo "modSP1_VariablePay", "ProcessRSUDividend", "May - Processing RSU Global"
        ProcessRSUGlobal ws, empIndex
        Exit Sub
    End If
    
    ' June: Process RSU EY only
    If currentMonth = 6 Then
        LogInfo "modSP1_VariablePay", "ProcessRSUDividend", "June - Processing RSU EY"
        ProcessRSUEY ws, empIndex
        Exit Sub
    End If
    
    LogInfo "modSP1_VariablePay", "ProcessRSUDividend", "Not RSU month (May/June), skipping"
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
            wein = NormalizeEmployeeId(wein)
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
    
    ' Try to get "Flex Claim Summary" sheet by name, fallback to first sheet
    On Error Resume Next
    Set srcWs = wb.Worksheets("Flex Claim Summary")
    On Error GoTo ErrHandler
    If srcWs Is Nothing Then
        Set srcWs = wb.Worksheets(1)
        LogWarning "modSP1_VariablePay", "ProcessFlexClaim", _
            "Sheet 'Flex Claim Summary' not found, using first sheet"
    End If
    
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
' Note: Each employee has a separate sheet named "Merck Payroll Summary Report--xxx"
'       where xxx is the Employee ID
'------------------------------------------------------------------------------
Private Sub ProcessMerckPayrollSummary(ws As Worksheet, empIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim sheetName As String
    Dim empIdFromSheet As String, empIdFromCell As String, wein As String
    Dim netPay As Double, mpfEEMC As Double, mpfEEVC As Double
    Dim iaPaySplit As Double, mpfRI As Double, mpfVCRI As Double
    Dim row As Long
    Dim colIAPaySplit As Long, colMPFRI As Long, colMPFVCRI As Long
    Dim processedCount As Long
    
    On Error GoTo ErrHandler
    
    filePath = GetInputFilePath("MerckPayroll")
    If Dir(filePath) = "" Then Exit Sub
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    
    ' Find target columns in VariablePay sheet
    colIAPaySplit = FindColumnByHeader(ws.Rows(1), "IA Pay Split")
    colMPFRI = FindColumnByHeader(ws.Rows(1), "MPF Relevant Income Rewrite")
    colMPFVCRI = FindColumnByHeader(ws.Rows(1), "MPF VC Relevant Income Rewrite")
    
    processedCount = 0
    
    ' Iterate through all sheets looking for employee report sheets
    For Each srcWs In wb.Worksheets
        sheetName = srcWs.Name
        
        ' Check if sheet name matches pattern "Merck Payroll Summary Report--xxx"
        empIdFromSheet = ExtractEmployeeIdFromSheetName(sheetName)
        If empIdFromSheet = "" Then GoTo NextSheet
        
        ' Validate Employee ID from "Flexi form:" cell
        empIdFromCell = FindFlexiFormEmployeeId(srcWs)
        If empIdFromCell <> "" And empIdFromCell <> empIdFromSheet Then
            LogWarning "modSP1_VariablePay", "ProcessMerckPayrollSummary", _
                "Employee ID mismatch: Sheet name has '" & empIdFromSheet & _
                "' but Flexi form cell has '" & empIdFromCell & "'. Using sheet name value."
        End If
        
        wein = NormalizeEmployeeId(empIdFromSheet)
        
        row = GetOrAddRow(ws, wein, empIndex)
        If row > 0 Then
            ' Extract values from the sheet using adaptive header search
            netPay = FindMerckPayrollValue(srcWs, "Net Pay (include EAO & leave payment)")
            mpfEEMC = FindMerckPayrollValue(srcWs, "MPF EE MC")
            mpfEEVC = FindMerckPayrollValue(srcWs, "MPF EE VC")
            mpfRI = FindMerckPayrollValue(srcWs, "MPF Relevant Income")
            mpfVCRI = FindMerckPayrollValue(srcWs, "MPF VC Relevant Income")
            
            ' Calculate IA Pay Split = Net Pay + MPF EE MC + MPF EE VC
            iaPaySplit = RoundAmount2(netPay + mpfEEMC + mpfEEVC)
            
            ' Write values to VariablePay sheet
            If colIAPaySplit > 0 Then
                ws.Cells(row, colIAPaySplit).Value = iaPaySplit
            End If
            
            If colMPFRI > 0 Then
                ws.Cells(row, colMPFRI).Value = RoundAmount2(mpfRI)
            End If
            
            If colMPFVCRI > 0 Then
                ws.Cells(row, colMPFVCRI).Value = RoundAmount2(mpfVCRI)
            End If
            
            processedCount = processedCount + 1
        End If
        
NextSheet:
    Next srcWs
    
    wb.Close SaveChanges:=False
    LogInfo "modSP1_VariablePay", "ProcessMerckPayrollSummary", _
        "Processed Merck Payroll Summary: " & processedCount & " employees"
    Exit Sub
    
ErrHandler:
    LogError "modSP1_VariablePay", "ProcessMerckPayrollSummary", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Function: ExtractEmployeeIdFromSheetName
' Purpose: Extract Employee ID from sheet name pattern "Merck Payroll Summary Report--xxx"
' Returns: Employee ID string or empty string if pattern not matched
'------------------------------------------------------------------------------
Private Function ExtractEmployeeIdFromSheetName(sheetName As String) As String
    Dim prefix As String
    Dim pos As Long
    
    ExtractEmployeeIdFromSheetName = ""
    
    ' Look for the separator "����" (Chinese em dash) or "--" (double hyphen)
    pos = InStr(sheetName, "����")
    If pos > 0 Then
        ExtractEmployeeIdFromSheetName = Trim(Mid(sheetName, pos + 2))
        Exit Function
    End If
    
    pos = InStr(sheetName, "--")
    If pos > 0 Then
        ExtractEmployeeIdFromSheetName = Trim(Mid(sheetName, pos + 2))
        Exit Function
    End If
    
    ' Also try single em dash "��"
    pos = InStr(sheetName, "��")
    If pos > 0 Then
        ExtractEmployeeIdFromSheetName = Trim(Mid(sheetName, pos + 1))
        Exit Function
    End If
End Function

'------------------------------------------------------------------------------
' Function: FindFlexiFormEmployeeId
' Purpose: Find Employee ID from "Flexi form:" label in the sheet
' Returns: Employee ID string or empty string if not found
'------------------------------------------------------------------------------
Private Function FindFlexiFormEmployeeId(srcWs As Worksheet) As String
    Dim cell As Range
    Dim searchRange As Range
    Dim lastRow As Long, lastCol As Long
    
    FindFlexiFormEmployeeId = ""
    
    On Error Resume Next
    lastRow = srcWs.Cells(srcWs.Rows.count, 1).End(xlUp).row
    lastCol = srcWs.Cells(1, srcWs.Columns.count).End(xlToLeft).Column
    If lastRow < 1 Then lastRow = 100
    If lastCol < 1 Then lastCol = 20
    
    Set searchRange = srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(lastRow, lastCol))
    
    ' Search for "Flexi form" label
    Set cell = searchRange.Find(What:="Flexi form", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    
    If Not cell Is Nothing Then
        ' Employee ID is in the cell to the right of the label
        FindFlexiFormEmployeeId = Trim(CStr(Nz(cell.offset(0, 1).Value, "")))
    End If
    
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Function: FindMerckPayrollValue
' Purpose: Find value by searching for header keyword and reading the value below it
' Parameters:
'   srcWs - Source worksheet
'   headerKeyword - Header text to search for
' Returns: Double value found below the header, or 0 if not found
'------------------------------------------------------------------------------
Private Function FindMerckPayrollValue(srcWs As Worksheet, headerKeyword As String) As Double
    Dim cell As Range
    Dim searchRange As Range
    Dim lastRow As Long, lastCol As Long
    Dim valueCell As Range
    
    FindMerckPayrollValue = 0
    
    On Error Resume Next
    lastRow = srcWs.Cells(srcWs.Rows.count, 1).End(xlUp).row
    lastCol = srcWs.Cells(1, srcWs.Columns.count).End(xlToLeft).Column
    If lastRow < 1 Then lastRow = 100
    If lastCol < 1 Then lastCol = 20
    
    Set searchRange = srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(lastRow, lastCol))
    
    ' Search for header keyword
    Set cell = searchRange.Find(What:=headerKeyword, LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False)
    
    If Not cell Is Nothing Then
        ' Value is in the cell directly below the header
        Set valueCell = cell.offset(1, 0)
        FindMerckPayrollValue = ToDouble(valueCell.Value)
    End If
    
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Sub: ProcessExtraTable
' Purpose: Process Additional table for PPTO EAO Rate input and Flexible benefits
' Note: Both PPTO EAO Rate input and Flexible benefits come from [���⽱��] sheet
'       Header row is auto-detected (not necessarily row 1)
'------------------------------------------------------------------------------
Private Sub ProcessExtraTable(ws As Worksheet, empIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim headerRow As Long, lastRow As Long, i As Long
    Dim wein As String
    Dim pptoRate As Double, flexBenefit As Double
    Dim row As Long
    Dim weinCol As Long, pptoCol As Long, flexCol As Long
    Dim colPPTORate As Long, colFlexBenefit As Long
    
    On Error GoTo ErrHandler
    
    Set wb = OpenExtraTableWorkbook()
    If wb Is Nothing Then Exit Sub
    
    ' Process [���⽱��] sheet for PPTO EAO Rate input and Flexible benefits
    On Error Resume Next
    Set srcWs = wb.Worksheets("���⽱��")
    On Error GoTo ErrHandler
    
    If srcWs Is Nothing Then
        LogWarning "modSP1_VariablePay", "ProcessExtraTable", "Sheet [���⽱��] not found in Extra Table"
        Exit Sub
    End If
    
    ' Auto-detect header row by searching for WEIN keyword
    headerRow = FindHeaderRow(srcWs, "WEIN,WIN,Employee ID,EmployeeID")
    If headerRow = 0 Then
        LogWarning "modSP1_VariablePay", "ProcessExtraTable", "Header row with WEIN not found in [���⽱��] sheet"
        Exit Sub
    End If
    
    ' Find columns in the detected header row
    weinCol = FindColumnByHeader(srcWs.Rows(headerRow), "WEIN,WIN,WEINEmployee ID,Employee CodeWIN,Employee ID,EmployeeID")
    pptoCol = FindColumnByHeader(srcWs.Rows(headerRow), "PPTO EAO Rate input")
    flexCol = FindColumnByHeader(srcWs.Rows(headerRow), "Flexible benefits")
    
    colPPTORate = FindColumnByHeader(ws.Rows(1), "PPTO EAO Rate input")
    colFlexBenefit = FindColumnByHeader(ws.Rows(1), "Flexible benefits")
    
    If weinCol = 0 Then
        LogWarning "modSP1_VariablePay", "ProcessExtraTable", "WEIN column not found in [���⽱��] sheet"
        Exit Sub
    End If
    
    lastRow = srcWs.Cells(srcWs.Rows.count, weinCol).End(xlUp).row
    
    ' Data starts from row after header
    For i = headerRow + 1 To lastRow
        wein = Trim(CStr(Nz(srcWs.Cells(i, weinCol).Value, "")))
        
        If wein <> "" Then
            row = GetOrAddRow(ws, wein, empIndex)
            If row > 0 Then
                ' Write PPTO EAO Rate input
                If pptoCol > 0 And colPPTORate > 0 Then
                    pptoRate = ToDouble(srcWs.Cells(i, pptoCol).Value)
                    If pptoRate <> 0 Then
                        ws.Cells(row, colPPTORate).Value = RoundAmount2(pptoRate)
                    End If
                End If
                
                ' Write Flexible benefits
                If flexCol > 0 And colFlexBenefit > 0 Then
                    flexBenefit = ToDouble(srcWs.Cells(i, flexCol).Value)
                    If flexBenefit <> 0 Then
                        ws.Cells(row, colFlexBenefit).Value = SafeAdd2(ws.Cells(row, colFlexBenefit).Value, flexBenefit)
                    End If
                End If
            End If
        End If
    Next i
    
    LogInfo "modSP1_VariablePay", "ProcessExtraTable", "Processed Extra Table [���⽱��] sheet (header at row " & headerRow & ")"
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



