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
    
    ' Load and process Special Bonuses from 额外表
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
                ElseIf InStr(planType, "OTHER BONUS") > 0 Then
                    col = GetCheckColIndex("Other Bonus 99999999")
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
    
    ' Use new path service
    filePath = GetInputFilePathAuto("InspireAwards", poCurrentMonth)
    If Not FileExistsSafe(filePath) Then
        LogWarning "modSP2_CheckResult_Incentives", "ProcessInspireCheck", _
            "Inspire Awards file does not exist (optional): " & filePath
        Exit Sub
    End If
    
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
                    col = GetCheckColIndex("InspirePoints")
                ElseIf InStr(planType, "INSPIRE CASH") > 0 Then
                    col = GetCheckColIndex("Inspire Cash 60702000")
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
    ' Write to "Shares Dividend 60204001 Check" column
    Dim col As Long
    col = GetCheckColIndex("Shares Dividend 60204001")
    
    ' Placeholder: actual RSU processing logic would go here
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Incentives", "ProcessRSUCheck", Err.Number, Err.Description
End Sub


'------------------------------------------------------------------------------
' Sub: ProcessSpecialBonusCheck
' Purpose: Process special bonuses from 额外表.[特殊奖金] for Check columns
'------------------------------------------------------------------------------
Private Sub ProcessSpecialBonusCheck(ws As Worksheet, weinIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim lastRow As Long, i As Long
    Dim headers As Object
    Dim wein As String
    Dim row As Long, col As Long
    
    On Error GoTo ErrHandler
    
    Set wb = OpenExtraTableWorkbook()
    If wb Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set srcWs = wb.Worksheets("特殊奖金")
    On Error GoTo ErrHandler
    
    If srcWs Is Nothing Then Exit Sub
    
    ' Build header index
    Set headers = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To srcWs.Cells(1, srcWs.Columns.count).End(xlToLeft).Column
        headers(UCase(Trim(CStr(srcWs.Cells(1, c).Value)))) = c
    Next c
    
    lastRow = srcWs.Cells(srcWs.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        ' Get WEIN
        wein = GetCellValFromHeaders(srcWs, i, headers, "WEIN")
        If wein = "" Then wein = GetCellValFromHeaders(srcWs, i, headers, "WIN")
        
        If wein <> "" And weinIndex.exists(wein) Then
            row = weinIndex(wein)
            
            ' Flexible benefits Check
            col = GetCheckColIndex("Flexible benefits")
            If col > 0 Then
                ws.Cells(row, col).Value = ToDouble(GetCellValFromHeaders(srcWs, i, headers, "FLEXIBLE BENEFITS"))
            End If
            
            ' Other Allowance Check
            col = GetCheckColIndex("Other Allowance 60409960")
            If col > 0 Then
                ws.Cells(row, col).Value = ToDouble(GetCellValFromHeaders(srcWs, i, headers, "OTHER ALLOWANCE"))
            End If
            
            ' Other Bonus Check
            col = GetCheckColIndex("Other Bonus 99999999")
            If col > 0 Then
                ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, _
                    ToDouble(GetCellValFromHeaders(srcWs, i, headers, "OTHER BONUS")))
            End If
            
            ' Other Rewards Check
            col = GetCheckColIndex("Other Rewards 99999999")
            If col > 0 Then
                ws.Cells(row, col).Value = ToDouble(GetCellValFromHeaders(srcWs, i, headers, "OTHER REWARDS"))
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
        GetCellValFromHeaders = Trim(CStr(Nz(ws.Cells(row, col).Value, "")))
    End If
End Function


'------------------------------------------------------------------------------
' Sub: ProcessIAPaySplitCheck
' Purpose: Process IA Pay Split Check from Merck Payroll Summary Report
' Formula: IA Pay Split = Net Pay (include EAO & leave payment) + MPF EE MC + MPF EE VC
' Note: Each employee has a separate sheet named "Merck Payroll Summary Report——xxx"
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
        
        ' Check if sheet name matches pattern "Merck Payroll Summary Report——xxx"
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
                ws.Cells(row, col).Value = RoundAmount2(iaPaySplit)
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
' Purpose: Extract Employee ID from sheet name pattern "Merck Payroll Summary Report——xxx"
' Returns: Employee ID string or empty string if pattern not matched
'------------------------------------------------------------------------------
Private Function ExtractEmpIdFromSheetName(sheetName As String) As String
    Dim pos As Long
    
    ExtractEmpIdFromSheetName = ""
    
    ' Look for the separator "——" (Chinese em dash) or "--" (double hyphen)
    pos = InStr(sheetName, "——")
    If pos > 0 Then
        ExtractEmpIdFromSheetName = Trim(Mid(sheetName, pos + 2))
        Exit Function
    End If
    
    pos = InStr(sheetName, "--")
    If pos > 0 Then
        ExtractEmpIdFromSheetName = Trim(Mid(sheetName, pos + 2))
        Exit Function
    End If
    
    ' Also try single em dash "—"
    pos = InStr(sheetName, "—")
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
        FindFlexiFormEmpId = Trim(CStr(Nz(cell.Offset(0, 1).Value, "")))
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
        Set valueCell = cell.Offset(1, 0)
        FindMerckValue = ToDouble(valueCell.Value)
    End If
    
    On Error GoTo 0
End Function
