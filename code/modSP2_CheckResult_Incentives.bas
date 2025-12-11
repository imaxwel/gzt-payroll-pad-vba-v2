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
    
    ' Load and process Special Bonuses from ?????
    ProcessSpecialBonusCheck ws, weinIndex
    
    ' Load and process IA Pay Split from Merck Payroll Summary
    ProcessIAPaySplitCheck ws, weinIndex
    
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
    Dim headerRow As Long
    
    On Error GoTo ErrHandler
    
    ' Use new path service
    filePath = GetInputFilePathAuto("OneTimePayment", poCurrentMonth)
    If Not FileExistsSafe(filePath) Then
        LogWarning "modSP2_CheckResult_Incentives", "ProcessOneTimePaymentCheck", _
            "One Time Payment file does not exist (optional): " & filePath
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    ' Auto-detect header row by searching for "Employee ID" keyword
    headerRow = FindHeaderRow(srcWs, "Employee ID,EmployeeID,WEIN", 20)
    If headerRow = 0 Then headerRow = 1  ' Default to row 1 if not found
    
    lastRow = srcWs.Cells(srcWs.Rows.count, 1).End(xlUp).row
    lastCol = srcWs.Cells(headerRow, srcWs.Columns.count).End(xlToLeft).Column
    Set dataRange = srcWs.Range(srcWs.Cells(headerRow, 1), srcWs.Cells(lastRow, lastCol))
    
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
                
                ' Map plan types to Check columns using template
                If InStr(planType, "LUMP SUM") > 0 Then
                    col = GetCheckColIndex("Lump Sum Bonus 60409960")
                ElseIf InStr(planType, "SIGN ON") > 0 Then
                    col = GetCheckColIndex("Sign On Bonus 60409960")
                ElseIf InStr(planType, "RETENTION") > 0 Then
                    col = GetCheckColIndex("Retention Bonus 60409960")
                ElseIf InStr(planType, "REFERRAL") > 0 Then
                    col = GetCheckColIndex("Referral Bonus 69001000")
                ElseIf InStr(planType, "RED PACKET") > 0 Or InStr(planType, "NEW YEAR") > 0 Then
                    col = GetCheckColIndex("Red Packet 69001000")
                ElseIf InStr(planType, "ANNUAL INCENTIVE") > 0 Or InStr(planType, "AIP") > 0 Then
                    col = GetCheckColIndex("Annual Incentive 60201000")
                ElseIf InStr(planType, "YEAR END") > 0 Then
                    col = GetCheckColIndex("Year End Bonus 60208000")
                ElseIf InStr(planType, "MANAGER OF THE YEAR") > 0 Then
                    col = GetCheckColIndex("Manager of the Year Award 60208000")
                ElseIf InStr(planType, "MD AWARD") > 0 Then
                    col = GetCheckColIndex("MD Award 60208000")
                ElseIf InStr(planType, "OTHER BONUS") > 0 Then
                    col = GetCheckColIndex("Other Bonus 99999999")
                Else
                    col = 0
                End If
                
                If col > 0 Then
                    ws.Cells(row, col).value = SafeAdd2(ws.Cells(row, col).value, grouped(key))
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
    Dim headerRow As Long
    
    On Error GoTo ErrHandler
    
    ' Use new path service
    filePath = GetInputFilePathAuto("InspireAwards", poCurrentMonth)
    If Not FileExistsSafe(filePath) Then
        LogWarning "modSP2_CheckResult_Incentives", "ProcessInspireCheck", _
            "Inspire Awards file does not exist (optional): " & filePath
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    ' Auto-detect header row by searching for "Employee ID" keyword
    headerRow = FindHeaderRow(srcWs, "Employee ID,EmployeeID,WEIN", 20)
    If headerRow = 0 Then
        LogWarning "modSP2_CheckResult_Incentives", "ProcessInspireCheck", _
            "Could not find header row in Inspire Awards file"
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    lastRow = srcWs.Cells(srcWs.Rows.count, 1).End(xlUp).row
    lastCol = srcWs.Cells(headerRow, srcWs.Columns.count).End(xlToLeft).Column
    Set dataRange = srcWs.Range(srcWs.Cells(headerRow, 1), srcWs.Cells(lastRow, lastCol))
    
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
                    col = GetCheckColIndex("InspirePoints")
                ElseIf InStr(planType, "INSPIRE CASH") > 0 Then
                    col = GetCheckColIndex("Inspire Cash 60702000")
                Else
                    col = 0
                End If
                
                If col > 0 Then
                    ws.Cells(row, col).value = SafeAdd2(ws.Cells(row, col).value, grouped(key))
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
    
    ' Use new path service (quarterly file)
    filePath = GetInputFilePathAuto("QXPayout", poCurrentMonth)
    If Not FileExistsSafe(filePath) Then
        LogWarning "modSP2_CheckResult_Incentives", "ProcessSIPCheck", _
            "QX Payout file does not exist (optional): " & filePath
        Exit Sub
    End If
    
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
                
                If InStr(payItem, "QUALITATIVE INCENTIVE PLAN") > 0 Then
                    col = GetCheckColIndex("Sales Incentive (Qualitative) 21201000")
                ElseIf InStr(payItem, "SALES INCENTIVE PLAN") > 0 Then
                    col = GetCheckColIndex("Sales Incentive (Quantitative)   21201000")
                Else
                    col = 0
                End If
                
                If col > 0 Then
                    ws.Cells(row, col).value = SafeAdd2(ws.Cells(row, col).value, grouped(key))
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
' Note: RSU Global is processed in May only, RSU EY is processed in June only
'------------------------------------------------------------------------------
Private Sub ProcessRSUCheck(ws As Worksheet, weinIndex As Object)
    Dim currentMonth As Integer
    
    On Error GoTo ErrHandler
    
    ' Get current payroll month
    currentMonth = Month(G.Payroll.monthStart)
    
    ' May: Process RSU Global only
    If currentMonth = 5 Then
        LogInfo "modSP2_CheckResult_Incentives", "ProcessRSUCheck", "May - Processing RSU Global Check"
        ProcessRSUGlobalCheck ws, weinIndex
        Exit Sub
    End If
    
    ' June: Process RSU EY only
    If currentMonth = 6 Then
        LogInfo "modSP2_CheckResult_Incentives", "ProcessRSUCheck", "June - Processing RSU EY Check"
        ProcessRSUEYCheck ws, weinIndex
        Exit Sub
    End If
    
    LogInfo "modSP2_CheckResult_Incentives", "ProcessRSUCheck", "Not RSU month (May/June), skipping"
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Incentives", "ProcessRSUCheck", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: ProcessRSUGlobalCheck
' Purpose: Process RSU Global Dividend for Check columns (May only)
' Source: RSU Dividend global report - Employee Reference, Gross Award Amount to be Paid
' Formula: Gross Award Amount to be Paid * Exchange rate
'------------------------------------------------------------------------------
Private Sub ProcessRSUGlobalCheck(ws As Worksheet, weinIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim lastRow As Long, i As Long, headerRow As Long
    Dim empRef As String, wein As String
    Dim grossAmt As Double, fxRate As Double, calcValue As Double
    Dim col As Long, row As Long
    Dim empRefCol As Long, amtCol As Long
    Dim empValues As Object ' Dictionary to aggregate duplicate Employee References
    
    On Error GoTo ErrHandler
    
    ' Get Check column index
    col = GetCheckColIndex("Shares Dividend 60204001")
    If col = 0 Then
        LogWarning "modSP2_CheckResult_Incentives", "ProcessRSUGlobalCheck", _
            "Shares Dividend 60204001 Check column not found"
        Exit Sub
    End If
    
    ' Get file path
    filePath = GetInputFilePathAuto("RSUGlobal", poCurrentMonth)
    If Not FileExistsSafe(filePath) Then
        LogWarning "modSP2_CheckResult_Incentives", "ProcessRSUGlobalCheck", _
            "RSU Global file does not exist: " & filePath
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    ' Get exchange rate from config
    fxRate = GetExchangeRate("RSU_Global")
    
    ' Detect header row and columns (header may not be on first row)
    headerRow = FindHeaderRowSafe(srcWs, "Employee Reference,EmployeeNumber,Employee Number,Employee ID,EmployeeID", 1, 50)
    empRefCol = FindColumnByHeader(srcWs.Rows(headerRow), "Employee Reference,EmployeeNumber,Employee Number,Employee ID,EmployeeID")
    amtCol = FindColumnByHeader(srcWs.Rows(headerRow), "Gross Award Amount to be Paid")
    
    If empRefCol = 0 Or amtCol = 0 Then
        LogWarning "modSP2_CheckResult_Incentives", "ProcessRSUGlobalCheck", _
            "Required columns not found in RSU Global file"
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Aggregate values by Employee Reference (handle duplicates)
    Set empValues = CreateObject("Scripting.Dictionary")
    lastRow = srcWs.Cells(srcWs.Rows.count, empRefCol).End(xlUp).row
    
    For i = headerRow + 1 To lastRow
        empRef = Trim(CStr(Nz(srcWs.Cells(i, empRefCol).value, "")))
        grossAmt = ToDouble(srcWs.Cells(i, amtCol).value)
        
        If empRef <> "" And grossAmt <> 0 Then
            calcValue = grossAmt * fxRate
            wein = NormalizeEmployeeId(empRef)
            
            If empValues.exists(wein) Then
                empValues(wein) = empValues(wein) + calcValue
            Else
                empValues.Add wein, calcValue
            End If
        End If
    Next i
    
    ' Write aggregated values to Check column
    Dim key As Variant
    For Each key In empValues.Keys
        If weinIndex.exists(key) Then
            row = weinIndex(key)
            ws.Cells(row, col).value = SafeAdd2(ws.Cells(row, col).value, empValues(key))
        End If
    Next key
    
    wb.Close SaveChanges:=False
    LogInfo "modSP2_CheckResult_Incentives", "ProcessRSUGlobalCheck", _
        "Processed RSU Global Check: " & empValues.count & " employees"
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Incentives", "ProcessRSUGlobalCheck", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Sub: ProcessRSUEYCheck
' Purpose: Process RSU EY Dividend for Check columns (June only)
' Source: RSU Dividend EY report - EmployeeNumber, Dividend To Pay
' Formula: Dividend To Pay * Exchange rate
'------------------------------------------------------------------------------
Private Sub ProcessRSUEYCheck(ws As Worksheet, weinIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim lastRow As Long, i As Long, headerRow As Long
    Dim empNum As String, wein As String
    Dim divAmt As Double, fxRate As Double, calcValue As Double
    Dim col As Long, row As Long
    Dim empNumCol As Long, amtCol As Long
    Dim empValues As Object ' Dictionary to aggregate duplicate Employee Numbers
    
    On Error GoTo ErrHandler
    
    ' Get Check column index
    col = GetCheckColIndex("Shares Dividend 60204001")
    If col = 0 Then
        LogWarning "modSP2_CheckResult_Incentives", "ProcessRSUEYCheck", _
            "Shares Dividend 60204001 Check column not found"
        Exit Sub
    End If
    
    ' Get file path
    filePath = GetInputFilePathAuto("RSUEY", poCurrentMonth)
    If Not FileExistsSafe(filePath) Then
        LogWarning "modSP2_CheckResult_Incentives", "ProcessRSUEYCheck", _
            "RSU EY file does not exist: " & filePath
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    ' Get exchange rate from config
    fxRate = GetExchangeRate("RSU_EY")
    
    ' Detect header row and columns (header may not be on first row)
    headerRow = FindHeaderRowSafe(srcWs, "EmployeeNumber,Employee Number,Employee ID,EmployeeID,Employee Reference", 1, 50)
    empNumCol = FindColumnByHeader(srcWs.Rows(headerRow), "EmployeeNumber,Employee Number,Employee ID,EmployeeID,Employee Reference")
    amtCol = FindColumnByHeader(srcWs.Rows(headerRow), "Dividend To Pay")
    
    If empNumCol = 0 Or amtCol = 0 Then
        LogWarning "modSP2_CheckResult_Incentives", "ProcessRSUEYCheck", _
            "Required columns not found in RSU EY file"
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Aggregate values by Employee Number (handle duplicates)
    Set empValues = CreateObject("Scripting.Dictionary")
    lastRow = srcWs.Cells(srcWs.Rows.count, empNumCol).End(xlUp).row
    
    For i = headerRow + 1 To lastRow
        empNum = Trim(CStr(Nz(srcWs.Cells(i, empNumCol).value, "")))
        divAmt = ToDouble(srcWs.Cells(i, amtCol).value)
        
        If empNum <> "" And divAmt <> 0 Then
            calcValue = divAmt * fxRate
            wein = NormalizeEmployeeId(empNum)
            
            If empValues.exists(wein) Then
                empValues(wein) = empValues(wein) + calcValue
            Else
                empValues.Add wein, calcValue
            End If
        End If
    Next i
    
    ' Write aggregated values to Check column
    Dim key As Variant
    For Each key In empValues.Keys
        If weinIndex.exists(key) Then
            row = weinIndex(key)
            ws.Cells(row, col).value = SafeAdd2(ws.Cells(row, col).value, empValues(key))
        End If
    Next key
    
    wb.Close SaveChanges:=False
    LogInfo "modSP2_CheckResult_Incentives", "ProcessRSUEYCheck", _
        "Processed RSU EY Check: " & empValues.count & " employees"
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Incentives", "ProcessRSUEYCheck", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub


'------------------------------------------------------------------------------
' Sub: ProcessSpecialBonusCheck
' Purpose: Process special bonuses from ?????.[ÌØÊâ½±½ð] for Check columns
'------------------------------------------------------------------------------
Private Sub ProcessSpecialBonusCheck(ws As Worksheet, weinIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim lastRow As Long, i As Long
    Dim headers As Object
    Dim wein As String
    Dim row As Long, col As Long
    Dim headerRow As Long, keyCol As Long
    
    On Error GoTo ErrHandler
    
    Set wb = OpenExtraTableWorkbook()
    If wb Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set srcWs = wb.Worksheets("ÌØÊâ½±½ð")
    On Error GoTo ErrHandler
    
    If srcWs Is Nothing Then Exit Sub
    
    ' Detect header row and build header index
    headerRow = FindHeaderRowSafe(srcWs, "WEIN,WIN", 1, 50)
    Set headers = BuildHeaderIndex(srcWs, headerRow)
    
    keyCol = GetColumnFromHeaders(headers, "WEIN,WIN")
    If keyCol = 0 Then keyCol = 1
    lastRow = srcWs.Cells(srcWs.Rows.count, keyCol).End(xlUp).row
    
    For i = headerRow + 1 To lastRow
        ' Get WEIN
        wein = GetCellValFromHeaders(srcWs, i, headers, "WEIN")
        If wein = "" Then wein = GetCellValFromHeaders(srcWs, i, headers, "WIN")
        
        If wein <> "" And weinIndex.exists(wein) Then
            row = weinIndex(wein)
            
            ' Flexible benefits Check
            col = GetCheckColIndex("Flexible benefits")
            If col > 0 Then
                ws.Cells(row, col).value = ToDouble(GetCellValFromHeaders(srcWs, i, headers, "FLEXIBLE BENEFITS"))
            End If
            
            ' Other Allowance Check
            col = GetCheckColIndex("Other Allowance 60409960")
            If col > 0 Then
                ws.Cells(row, col).value = ToDouble(GetCellValFromHeaders(srcWs, i, headers, "OTHER ALLOWANCE"))
            End If
            
            ' Other Bonus Check
            col = GetCheckColIndex("Other Bonus 99999999")
            If col > 0 Then
                ws.Cells(row, col).value = SafeAdd2(ws.Cells(row, col).value, _
                    ToDouble(GetCellValFromHeaders(srcWs, i, headers, "OTHER BONUS")))
            End If
            
            ' Other Rewards Check
            col = GetCheckColIndex("Other Rewards 99999999")
            If col > 0 Then
                ws.Cells(row, col).value = ToDouble(GetCellValFromHeaders(srcWs, i, headers, "OTHER REWARDS"))
            End If
        End If
    Next i
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Incentives", "ProcessSpecialBonusCheck", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Helper: GetCellValFromHeaders
'------------------------------------------------------------------------------
Private Function GetCellValFromHeaders(ws As Worksheet, row As Long, headers As Object, headerName As String) As String
    Dim col As Long
    GetCellValFromHeaders = ""
    
    If headers.exists(UCase(headerName)) Then
        col = headers(UCase(headerName))
        GetCellValFromHeaders = Trim(CStr(Nz(ws.Cells(row, col).value, "")))
    End If
End Function


'------------------------------------------------------------------------------
' Sub: ProcessIAPaySplitCheck
' Purpose: Process IA Pay Split Check from Merck Payroll Summary Report
' Formula: IA Pay Split = Net Pay (include EAO & leave payment) + MPF EE MC + MPF EE VC
' Note: Each employee has a separate sheet named "Merck Payroll Summary Report--xxx"
'       where xxx is the Employee ID
'------------------------------------------------------------------------------
Private Sub ProcessIAPaySplitCheck(ws As Worksheet, weinIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim sheetName As String
    Dim empIdFromSheet As String, empIdFromCell As String, wein As String
    Dim row As Long, col As Long
    Dim netPay As Double, mpfEEMC As Double, mpfEEVC As Double
    Dim iaPaySplit As Double
    Dim processedCount As Long
    
    On Error GoTo ErrHandler
    
    col = GetCheckColIndex("IA Pay Split")
    If col = 0 Then Exit Sub
    
    ' Use new path service
    filePath = GetInputFilePathAuto("MerckPayroll", poCurrentMonth)
    If Not FileExistsSafe(filePath) Then
        LogWarning "modSP2_CheckResult_Incentives", "ProcessIAPaySplitCheck", _
            "Merck Payroll Summary file does not exist (optional): " & filePath
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    
    processedCount = 0
    
    ' Iterate through all sheets looking for employee report sheets
    For Each srcWs In wb.Worksheets
        sheetName = srcWs.Name
        
        ' Check if sheet name matches pattern "Merck Payroll Summary Report--xxx"
        empIdFromSheet = ExtractEmpIdFromSheetName(sheetName)
        If empIdFromSheet = "" Then GoTo NextSheet
        
        ' Validate Employee ID from "Flexi form:" cell
        empIdFromCell = FindFlexiFormEmpId(srcWs)
        If empIdFromCell <> "" And empIdFromCell <> empIdFromSheet Then
            LogWarning "modSP2_CheckResult_Incentives", "ProcessIAPaySplitCheck", _
                "Employee ID mismatch: Sheet name has '" & empIdFromSheet & _
                "' but Flexi form cell has '" & empIdFromCell & "'. Using sheet name value."
        End If
        
        wein = NormalizeEmployeeId(empIdFromSheet)
        
        If weinIndex.exists(wein) Then
            row = weinIndex(wein)
            
            ' Extract values from the sheet using adaptive header search
            netPay = FindMerckValue(srcWs, "Net Pay (include EAO & leave payment)")
            mpfEEMC = FindMerckValue(srcWs, "MPF EE MC")
            mpfEEVC = FindMerckValue(srcWs, "MPF EE VC")
            
            ' Calculate IA Pay Split = Net Pay + MPF EE MC + MPF EE VC
            iaPaySplit = netPay + mpfEEMC + mpfEEVC
            
            If iaPaySplit <> 0 Then
                ws.Cells(row, col).value = RoundAmount2(iaPaySplit)
            End If
            
            processedCount = processedCount + 1
        End If
        
NextSheet:
    Next srcWs
    
    wb.Close SaveChanges:=False
    LogInfo "modSP2_CheckResult_Incentives", "ProcessIAPaySplitCheck", _
        "Processed IA Pay Split Check: " & processedCount & " employees"
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Incentives", "ProcessIAPaySplitCheck", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Function: ExtractEmpIdFromSheetName
' Purpose: Extract Employee ID from sheet name pattern "Merck Payroll Summary Report--xxx"
' Returns: Employee ID string or empty string if pattern not matched
'------------------------------------------------------------------------------
Private Function ExtractEmpIdFromSheetName(sheetName As String) As String
    Dim pos As Long
    
    ExtractEmpIdFromSheetName = ""
    
    ' Look for the separator "????" (Chinese em dash) or "--" (double hyphen)
    pos = InStr(sheetName, "????")
    If pos > 0 Then
        ExtractEmpIdFromSheetName = Trim(Mid(sheetName, pos + 2))
        Exit Function
    End If
    
    pos = InStr(sheetName, "--")
    If pos > 0 Then
        ExtractEmpIdFromSheetName = Trim(Mid(sheetName, pos + 2))
        Exit Function
    End If
    
    ' Also try single em dash "??"
    pos = InStr(sheetName, "??")
    If pos > 0 Then
        ExtractEmpIdFromSheetName = Trim(Mid(sheetName, pos + 1))
        Exit Function
    End If
End Function

'------------------------------------------------------------------------------
' Function: FindFlexiFormEmpId
' Purpose: Find Employee ID from "Flexi form:" label in the sheet
' Returns: Employee ID string or empty string if not found
'------------------------------------------------------------------------------
Private Function FindFlexiFormEmpId(srcWs As Worksheet) As String
    Dim cell As Range
    Dim searchRange As Range
    Dim lastRow As Long, lastCol As Long
    
    FindFlexiFormEmpId = ""
    
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
        FindFlexiFormEmpId = Trim(CStr(Nz(cell.offset(0, 1).value, "")))
    End If
    
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Function: FindMerckValue
' Purpose: Find value by searching for header keyword and reading the value below it
' Parameters:
'   srcWs - Source worksheet
'   headerKeyword - Header text to search for
' Returns: Double value found below the header, or 0 if not found
'------------------------------------------------------------------------------
Private Function FindMerckValue(srcWs As Worksheet, headerKeyword As String) As Double
    Dim cell As Range
    Dim searchRange As Range
    Dim lastRow As Long, lastCol As Long
    Dim valueCell As Range
    
    FindMerckValue = 0
    
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
        FindMerckValue = ToDouble(valueCell.value)
    End If
    
    On Error GoTo 0
End Function



