Attribute VB_Name = "modSP2_CheckResult_PayItems"
'==============================================================================
' Module: modSP2_CheckResult_PayItems
' Purpose: Subprocess 2 - Pay Items Check columns
' Description: Validates Base Pay, Leave Payments, EAO Adjustments
'==============================================================================
Option Explicit

' Workforce/allowance/leave caches for pay-item checks
Private mPayWorkforce As Object          ' Dictionary: WEIN -> record
Private mTransportAllowance As Object    ' Dictionary: WEIN -> transport allowance amount
Private mPrevLeaveKeys As Object         ' Dictionary of unique leave keys from previous month
Private mNewLeaveRecords As Collection   ' Collection of new approved leave records for current month
Private mMatDaysPrev As Object           ' Dictionary: WEIN -> maternity days in prev month
Private mPatDaysPrev As Object           ' Dictionary: WEIN -> paternity days in prev month
Private mUnpaidDaysCurr As Object        ' Dictionary: WEIN -> unpaid leave days in current month

' Leave record indices
Private Const LR_WEIN As Long = 0
Private Const LR_LEAVETYPE As Long = 1
Private Const LR_FROMDATE As Long = 2
Private Const LR_TODATE As Long = 3
Private Const LR_APPLYDATE As Long = 4
Private Const LR_APPROVALDATE As Long = 5
Private Const LR_TOTALDAYS As Long = 6
Private Const LR_UNIQUEKEY As Long = 7

'------------------------------------------------------------------------------
' Sub: SP2_Check_PayItems
' Purpose: Populate pay items Check columns
'------------------------------------------------------------------------------
Public Sub SP2_Check_PayItems(valWb As Workbook, weinIndex As Object)
    Dim ws As Worksheet
    Dim wein As Variant
    Dim row As Long
    
    On Error GoTo ErrHandler
    
    Set ws = valWb.Worksheets("Check Result")
    
    ' Load dependent data
    LoadEAOData
    LoadPayWorkforceData
    LoadTransportAllowanceData
    LoadLeaveDataForPayItems
    
    ' Process each WEIN
    For Each wein In weinIndex.Keys
        row = weinIndex(wein)
        
        WriteBasePayCheck ws, row, CStr(wein)
        WriteLeavePaymentChecks ws, row, CStr(wein)
        WriteEAOChecks ws, row, CStr(wein)
    Next wein
    
    ' Write PPTO EAO Rate from extra table
    WritePPTOEAORateCheck ws, weinIndex
    
    LogInfo "modSP2_CheckResult_PayItems", "SP2_Check_PayItems", "Pay items checks completed"
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_PayItems", "SP2_Check_PayItems", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: WriteBasePayCheck
' Purpose: Write Base Pay related Check columns
'------------------------------------------------------------------------------
Private Sub WriteBasePayCheck(ws As Worksheet, row As Long, wein As String)
    Dim col As Long
    Dim rec As Object
    Dim monthlySalary As Double
    Dim empType As String
    Dim actualWork As Double
    Dim leaveDays As Double
    Dim transportAmt As Double
    Dim noPayDays As Double
    
    On Error Resume Next
    
    Set rec = GetPayEmpRecord(wein)
    If rec Is Nothing Then Exit Sub
    
    monthlySalary = Nz(rec("MonthlySalary"), 0)
    empType = UCase(CStr(Nz(rec("EmployeeType"), "")))
    actualWork = GetActualWorkingDay(ws, row)
    If actualWork = 0 Then actualWork = 1
    
    ' Base Pay for regular employees
    If monthlySalary <> 0 And (InStr(empType, "INTERN") = 0 And InStr(empType, "CO-OP") = 0) Then
        col = GetCheckColIndex("Base Pay 60001000")
        If col > 0 Then
            ws.Cells(row, col).Value = RoundAmount2(actualWork * monthlySalary)
        End If
    End If
    
    ' Base Pay (Temp) for interns/co-ops
    If monthlySalary <> 0 And (InStr(empType, "INTERN") > 0 Or InStr(empType, "CO-OP") > 0) Then
        col = GetCheckColIndex("Base Pay(Temp) 60101000")
        If col > 0 Then
            ws.Cells(row, col).Value = RoundAmount2(actualWork * monthlySalary)
        End If
    End If
    
    ' Salary Adj & Transport Allowance Adj (maternity + paternity leave days in previous month)
    leaveDays = GetLeaveDaysForPrevMonth(wein)
    If leaveDays <> 0 And monthlySalary <> 0 Then
        col = GetCheckColIndex("Salary Adj 60001000")
        If col > 0 Then ws.Cells(row, col).Value = CalcSalaryAdj(monthlySalary, leaveDays)
        
        col = GetCheckColIndex("Transport Allowance Adj 60409960")
        If col > 0 Then ws.Cells(row, col).Value = CalcSalaryAdj(monthlySalary, leaveDays)
    End If
    
    ' Transport Allowance with unpaid leave deduction (current month)
    transportAmt = GetTransportAllowance(wein)
    noPayDays = GetNoPayLeaveDaysCurrent(wein)
    col = GetCheckColIndex("Transport Allowance 60409960")
    If col > 0 And transportAmt <> 0 Then
        ws.Cells(row, col).Value = CalcTransportAllowance(transportAmt, noPayDays)
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
' Purpose: Write PPTO EAO Rate input Check column from extra table
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
    lastRow = srcWs.Cells(srcWs.Rows.Count, keyCol).End(xlUp).row

    For i = headerRow + 1 To lastRow
        ' Get WEIN
        wein = GetPPTOCellVal(srcWs, i, headers, "WEIN")
        If wein = "" Then wein = GetPPTOCellVal(srcWs, i, headers, "WIN")
        
        If wein <> "" And weinIndex.Exists(wein) Then
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
    
    If headers.Exists(UCase(headerName)) Then
        col = headers(UCase(headerName))
        GetPPTOCellVal = Trim(CStr(Nz(ws.Cells(row, col).Value, "")))
    End If
End Function

'------------------------------------------------------------------------------
' Data loading helpers
'------------------------------------------------------------------------------
Private Sub LoadPayWorkforceData()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filePath As String
    Dim headers As Object
    Dim headerRow As Long, lastRow As Long
    Dim empId As String, wein As String
    Dim rec As Object
    Dim hireDateVal As Variant
    Dim empType As String
    
    On Error GoTo ErrHandler
    
    Set mPayWorkforce = CreateObject("Scripting.Dictionary")
    
    filePath = GetInputFilePathAuto("WorkforceDetail", poCurrentMonth)
    If Not FileExistsSafe(filePath) Then
        LogError "modSP2_CheckResult_PayItems", "LoadPayWorkforceData", 0, _
            "Workforce Detail file does not exist: " & filePath
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Worksheets(1)
    
    headerRow = FindHeaderRowSafe(ws, "EMPLOYEE ID,EMPLOYEEID", 1, 50)
    Set headers = BuildHeaderIndex(ws, headerRow)
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    Dim i As Long
    For i = headerRow + 1 To lastRow
        empId = Trim(CStr(Nz(GetCellFromHeaders(ws, i, headers, "EMPLOYEE ID,EMPLOYEEID,EMPLOYEE NUMBER ID"), "")))
        wein = Trim(CStr(Nz(GetCellFromHeaders(ws, i, headers, "WEIN,WIN"), "")))
        If wein = "" Then wein = empId
        wein = NormalizeEmployeeId(wein)
        If wein <> "" Then
            Set rec = CreateObject("Scripting.Dictionary")
            rec("MonthlySalary") = RoundMonthlySalary(GetCellFromHeaders(ws, i, headers, "MONTHLY SALARY"))
            hireDateVal = GetCellFromHeaders(ws, i, headers, "LAST HIRE DATE")
            If IsDate(hireDateVal) Then rec("LastHireDate") = CDate(hireDateVal) Else rec("LastHireDate") = 0
            empType = Trim(CStr(Nz(GetCellFromHeaders(ws, i, headers, "EMPLOYEE TYPE"), "")))
            rec("EmployeeType") = empType
            
            If Not mPayWorkforce.Exists(wein) Then
                mPayWorkforce.Add wein, rec
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_PayItems", "LoadPayWorkforceData", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

Private Sub LoadTransportAllowanceData()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filePath As String
    Dim headers As Object
    Dim headerRow As Long, lastRow As Long
    Dim empId As String
    Dim compPlan As String
    Dim amt As Double
    
    On Error GoTo ErrHandler
    
    Set mTransportAllowance = CreateObject("Scripting.Dictionary")
    
    filePath = GetInputFilePathAuto("AllowancePlan", poCurrentMonth)
    If Not FileExistsSafe(filePath) Then
        LogWarning "modSP2_CheckResult_PayItems", "LoadTransportAllowanceData", _
            "Allowance plan file not found (optional): " & filePath
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Worksheets(1)
    
    headerRow = FindHeaderRowSafe(ws, "EMPLOYEE ID,EMPLOYEEID,EMPLOYEE NUMBER ID", 1, 50)
    Set headers = BuildHeaderIndex(ws, headerRow)
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    Dim i As Long
    For i = headerRow + 1 To lastRow
        compPlan = UCase(Trim(CStr(Nz(GetCellFromHeaders(ws, i, headers, "COMPENSATION PLAN"), ""))))
        If InStr(compPlan, "TRANSPORT") > 0 Then
            empId = NormalizeEmployeeId(Trim(CStr(Nz(GetCellFromHeaders(ws, i, headers, "EMPLOYEE ID,EMPLOYEEID,EMPLOYEE NUMBER ID"), ""))))
            amt = ToDouble(GetCellFromHeaders(ws, i, headers, "AMOUNT"))
            If empId <> "" Then
                If mTransportAllowance.Exists(empId) Then
                    mTransportAllowance(empId) = mTransportAllowance(empId) + amt
                Else
                    mTransportAllowance.Add empId, amt
                End If
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_PayItems", "LoadTransportAllowanceData", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

Private Sub LoadLeaveDataForPayItems()
    Set mMatDaysPrev = CreateObject("Scripting.Dictionary")
    Set mPatDaysPrev = CreateObject("Scripting.Dictionary")
    Set mUnpaidDaysCurr = CreateObject("Scripting.Dictionary")
    
    Set mPrevLeaveKeys = LoadLeaveKeyDict(poPreviousMonth)
    Set mNewLeaveRecords = LoadNewLeaveRecords(poCurrentMonth, mPrevLeaveKeys)
    
    Dim prevMonthKey As String
    Dim currMonthKey As String
    prevMonthKey = Format(G.Payroll.prevMonthStart, "yyyymm")
    currMonthKey = Format(G.Payroll.monthStart, "yyyymm")
    
    Dim v As Variant
    Dim rec As Variant
    Dim wein As String
    Dim leaveType As String
    Dim recHireDate As Date
    Dim weeksOfService As Double
    Dim splitDays As Object
    Dim monthKey As Variant
    Dim daysValue As Double
    
    For Each v In mNewLeaveRecords
        rec = v
        wein = NormalizeEmployeeId(CStr(rec(LR_WEIN)))
        leaveType = UCase(CStr(rec(LR_LEAVETYPE)))
        If wein = "" Then GoTo NextRec
        
        Set splitDays = SplitLeaveDaysByMonth(CDate(rec(LR_FROMDATE)), CDate(rec(LR_TODATE)), ToDouble(rec(LR_TOTALDAYS)))
        
        ' Maternity leave requires 40 weeks of service before start
        If InStr(leaveType, "MATERNITY") > 0 Then
            recHireDate = GetHireDate(wein)
            If recHireDate > 0 Then
                weeksOfService = (CDate(rec(LR_FROMDATE)) - recHireDate) / 7
                If weeksOfService < 40 Then GoTo NextRec
            End If
            For Each monthKey In splitDays.Keys
                If CStr(monthKey) = prevMonthKey Then
                    daysValue = splitDays(monthKey)
                    AddToDict mMatDaysPrev, wein, daysValue
                End If
            Next monthKey
        ElseIf InStr(leaveType, "PATERNITY") > 0 Then
            For Each monthKey In splitDays.Keys
                If CStr(monthKey) = prevMonthKey Then
                    daysValue = splitDays(monthKey)
                    AddToDict mPatDaysPrev, wein, daysValue
                End If
            Next monthKey
        ElseIf InStr(leaveType, "UNPAID") > 0 Then
            For Each monthKey In splitDays.Keys
                If CStr(monthKey) = currMonthKey Then
                    daysValue = splitDays(monthKey)
                    AddToDict mUnpaidDaysCurr, wein, daysValue
                End If
            Next monthKey
        End If
NextRec:
    Next v
End Sub

Private Function LoadLeaveKeyDict(offset As ePeriodOffset) As Object
    Dim dict As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filePath As String
    Dim headers As Object
    Dim headerRow As Long, lastRow As Long
    Dim i As Long
    Dim recStatus As String
    Dim wein As String
    Dim fromDate As Variant, toDate As Variant, applyDate As Variant, approvalDate As Variant
    Dim key As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    filePath = GetInputFilePathAuto("EmployeeLeave", offset)
    If Not FileExistsSafe(filePath) Then
        Set LoadLeaveKeyDict = dict
        Exit Function
    End If
    
    On Error GoTo ErrHandler
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Worksheets(1)
    
    headerRow = FindHeaderRowSafe(ws, _
        "WIN,WEIN,WEINEmployee ID,EMPLOYEE CODEWIN,EMPLOYEE CODE,EMPLOYEECODE,EMPLOYEE REFERENCE,EMPLOYEE NUMBER,EMPLOYEE NUMBER ID,EMPLOYEE ID,STATUS", _
        1, 50)
    Set headers = BuildHeaderIndex(ws, headerRow)
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    For i = headerRow + 1 To lastRow
        recStatus = UCase(CStr(Nz(GetCellFromHeaders(ws, i, headers, "STATUS"), "")))
        If recStatus = "APPROVED" Then
            wein = NormalizeEmployeeId(CStr(Nz(GetCellFromHeaders(ws, i, headers, "WIN,WEIN,EMPLOYEE ID,EMPLOYEE NUMBER ID,EMPLOYEECODE,EMPLOYEE NUMBER"), "")))
            fromDate = GetCellFromHeaders(ws, i, headers, "FROM_DATE")
            toDate = GetCellFromHeaders(ws, i, headers, "TO_DATE")
            applyDate = GetCellFromHeaders(ws, i, headers, "APPLY_DATE")
            approvalDate = GetCellFromHeaders(ws, i, headers, "APPROVAL_DATE")
            If wein <> "" And IsDate(fromDate) And IsDate(toDate) Then
                key = BuildLeaveKey(wein, CDate(fromDate), CDate(toDate), applyDate, approvalDate)
                If Not dict.Exists(key) Then dict.Add key, True
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    Set LoadLeaveKeyDict = dict
    Exit Function
    
ErrHandler:
    LogError "modSP2_CheckResult_PayItems", "LoadLeaveKeyDict", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set LoadLeaveKeyDict = dict
End Function

Private Function LoadNewLeaveRecords(offset As ePeriodOffset, prevKeys As Object) As Collection
    Dim col As New Collection
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filePath As String
    Dim headers As Object
    Dim headerRow As Long, lastRow As Long
    Dim recStatus As String
    Dim i As Long
    Dim wein As String
    Dim fromDate As Variant, toDate As Variant, applyDate As Variant, approvalDate As Variant
    Dim totalDays As Double
    Dim uniqueKey As String
    
    filePath = GetInputFilePathAuto("EmployeeLeave", offset)
    If Not FileExistsSafe(filePath) Then
        Set LoadNewLeaveRecords = col
        Exit Function
    End If
    
    On Error GoTo ErrHandler
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Worksheets(1)
    
    headerRow = FindHeaderRowSafe(ws, _
        "WIN,WEIN,WEINEmployee ID,EMPLOYEE CODEWIN,EMPLOYEE CODE,EMPLOYEECODE,EMPLOYEE REFERENCE,EMPLOYEE NUMBER,EMPLOYEE NUMBER ID,EMPLOYEE ID,STATUS", _
        1, 50)
    Set headers = BuildHeaderIndex(ws, headerRow)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    
    For i = headerRow + 1 To lastRow
        recStatus = UCase(CStr(Nz(GetCellFromHeaders(ws, i, headers, "STATUS"), "")))
        If recStatus = "APPROVED" Then
            wein = NormalizeEmployeeId(CStr(Nz(GetCellFromHeaders(ws, i, headers, "WIN,WEIN,EMPLOYEE ID,EMPLOYEE NUMBER ID,EMPLOYEECODE,EMPLOYEE NUMBER"), "")))
            fromDate = GetCellFromHeaders(ws, i, headers, "FROM_DATE")
            toDate = GetCellFromHeaders(ws, i, headers, "TO_DATE")
            applyDate = GetCellFromHeaders(ws, i, headers, "APPLY_DATE")
            approvalDate = GetCellFromHeaders(ws, i, headers, "APPROVAL_DATE")
            totalDays = ToDouble(GetCellFromHeaders(ws, i, headers, "TOTAL_DAYS"))
            If wein <> "" And IsDate(fromDate) And IsDate(toDate) Then
                uniqueKey = BuildLeaveKey(wein, CDate(fromDate), CDate(toDate), applyDate, approvalDate)
                If prevKeys Is Nothing Or Not prevKeys.Exists(uniqueKey) Then
                    Dim rec(0 To 7) As Variant
                    rec(LR_WEIN) = wein
                    rec(LR_LEAVETYPE) = Nz(GetCellFromHeaders(ws, i, headers, "LEAVE TYPE"), "")
                    rec(LR_FROMDATE) = CDate(fromDate)
                    rec(LR_TODATE) = CDate(toDate)
                    If IsDate(applyDate) Then rec(LR_APPLYDATE) = CDate(applyDate) Else rec(LR_APPLYDATE) = #1/1/1900#
                    If IsDate(approvalDate) Then rec(LR_APPROVALDATE) = CDate(approvalDate) Else rec(LR_APPROVALDATE) = #1/1/1900#
                    rec(LR_TOTALDAYS) = totalDays
                    rec(LR_UNIQUEKEY) = uniqueKey
                    col.Add rec
                End If
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    Set LoadNewLeaveRecords = col
    Exit Function
    
ErrHandler:
    LogError "modSP2_CheckResult_PayItems", "LoadNewLeaveRecords", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set LoadNewLeaveRecords = col
End Function

'------------------------------------------------------------------------------
' Calculation helpers
'------------------------------------------------------------------------------
Private Function GetPayEmpRecord(wein As String) As Object
    If mPayWorkforce Is Nothing Then
        Set GetPayEmpRecord = Nothing
        Exit Function
    End If
    If mPayWorkforce.Exists(wein) Then
        Set GetPayEmpRecord = mPayWorkforce(wein)
    Else
        Set GetPayEmpRecord = Nothing
    End If
End Function

Private Function GetHireDate(wein As String) As Date
    Dim rec As Object
    Set rec = GetPayEmpRecord(wein)
    If Not rec Is Nothing Then
        If IsDate(Nz(rec("LastHireDate"), 0)) Then
            GetHireDate = CDate(rec("LastHireDate"))
            Exit Function
        End If
    End If
    GetHireDate = 0
End Function

Private Function GetTransportAllowance(wein As String) As Double
    If Not mTransportAllowance Is Nothing Then
        If mTransportAllowance.Exists(wein) Then
            GetTransportAllowance = mTransportAllowance(wein)
            Exit Function
        End If
    End If
    GetTransportAllowance = 0
End Function

Private Function GetLeaveDaysForPrevMonth(wein As String) As Double
    GetLeaveDaysForPrevMonth = Nz(GetDictVal(mMatDaysPrev, wein), 0) + Nz(GetDictVal(mPatDaysPrev, wein), 0)
End Function

Private Function GetNoPayLeaveDaysCurrent(wein As String) As Double
    GetNoPayLeaveDaysCurrent = Nz(GetDictVal(mUnpaidDaysCurr, wein), 0)
End Function

Private Function GetDictVal(dict As Object, key As String) As Variant
    If dict Is Nothing Then
        GetDictVal = 0
    ElseIf dict.Exists(key) Then
        GetDictVal = dict(key)
    Else
        GetDictVal = 0
    End If
End Function

Private Function GetActualWorkingDay(ws As Worksheet, row As Long) As Double
    Dim colActual As Long
    colActual = FindColumnByHeader(ws.Rows(4), "Actual working day,Actual working days")
    If colActual > 0 Then
        GetActualWorkingDay = ToDouble(ws.Cells(row, colActual).Value)
    Else
        GetActualWorkingDay = 1
    End If
End Function

Private Function CalcSalaryAdj(monthlySalary As Double, leaveDays As Double) As Double
    Dim calendarDaysPrev As Long
    calendarDaysPrev = DateDiff("d", G.Payroll.prevMonthStart, G.Payroll.prevMonthEnd) + 1
    CalcSalaryAdj = SafeMultiply2(SafeDivide2(monthlySalary, calendarDaysPrev), leaveDays)
End Function

Private Function CalcTransportAllowance(amount As Double, noPayDays As Double) As Double
    Dim calendarDaysCurr As Long
    calendarDaysCurr = DateDiff("d", G.Payroll.monthStart, G.Payroll.monthEnd) + 1
    CalcTransportAllowance = RoundAmount2(amount - SafeMultiply2(SafeDivide2(amount, calendarDaysCurr), noPayDays))
End Function

Private Function BuildLeaveKey(wein As String, fromDate As Date, toDate As Date, applyDate As Variant, approvalDate As Variant) As String
    Dim applyStr As String, approvalStr As String
    If IsDate(applyDate) Then
        applyStr = Format(CDate(applyDate), "yyyymmdd")
    Else
        applyStr = "00000000"
    End If
    If IsDate(approvalDate) Then
        approvalStr = Format(CDate(approvalDate), "yyyymmdd")
    Else
        approvalStr = "00000000"
    End If
    BuildLeaveKey = wein & "|" & Format(fromDate, "yyyymmdd") & "|" & Format(toDate, "yyyymmdd") & "|" & applyStr & "|" & approvalStr
End Function

Private Function SplitLeaveDaysByMonth(fromDate As Date, toDate As Date, totalDays As Double) As Object
    Dim dict As Object
    Dim segmentStart As Date, segmentEnd As Date
    Dim totalSpan As Double
    Dim daysInSeg As Double
    Dim alloc As Double
    
    Set dict = CreateObject("Scripting.Dictionary")
    totalSpan = DateDiff("d", fromDate, toDate) + 1
    If totalSpan <= 0 Then
        Set SplitLeaveDaysByMonth = dict
        Exit Function
    End If
    
    segmentStart = fromDate
    Do While segmentStart <= toDate
        segmentEnd = DateSerial(Year(segmentStart), Month(segmentStart) + 1, 0)
        If segmentEnd > toDate Then segmentEnd = toDate
        daysInSeg = DateDiff("d", segmentStart, segmentEnd) + 1
        alloc = totalDays * (daysInSeg / totalSpan)
        AddToDict dict, Format(segmentStart, "yyyymm"), alloc
        segmentStart = segmentEnd + 1
    Loop
    
    Set SplitLeaveDaysByMonth = dict
End Function

Private Sub AddToDict(dict As Object, key As String, value As Double)
    If dict.Exists(key) Then
        dict(key) = Nz(dict(key), 0) + value
    Else
        dict.Add key, value
    End If
End Sub

Private Function GetCellFromHeaders(ws As Worksheet, rowNum As Long, headers As Object, possibleNames As String) As Variant
    Dim col As Long
    col = GetColumnFromHeaders(headers, possibleNames)
    If col > 0 Then
        GetCellFromHeaders = ws.Cells(rowNum, col).Value
    Else
        GetCellFromHeaders = ""
    End If
End Function
