Attribute VB_Name = "modSP1_Attendance"
'==============================================================================
' Module: modSP1_Attendance
' Purpose: Subprocess 1 - Attendance sheet population
' Description: Handles leave processing and attendance data population
'==============================================================================
Option Explicit

' Leave record structure - stored as array indices for Collection compatibility
' Index: 0=WEIN, 1=EmployeeCode, 2=LeaveType, 3=FromDate, 4=ToDate,
'        5=ApplyDate, 6=ApprovalDate, 7=Status, 8=TotalDays, 9=UniqueKey
Private Const LR_WEIN As Long = 0
Private Const LR_EMPCODE As Long = 1
Private Const LR_LEAVETYPE As Long = 2
Private Const LR_FROMDATE As Long = 3
Private Const LR_TODATE As Long = 4
Private Const LR_APPLYDATE As Long = 5
Private Const LR_APPROVALDATE As Long = 6
Private Const LR_STATUS As Long = 7
Private Const LR_TOTALDAYS As Long = 8
Private Const LR_UNIQUEKEY As Long = 9

' Leave history for tracking processed records
Private mLeaveHistory As Object ' Dictionary of processed leave keys

' Workforce data cache for Last Hire Date lookup
Private mWorkforceHireDates As Object ' Dictionary of WEIN/EmpId -> Last Hire Date

' Maternity Report excluded records
Private mMaternityExcludedRecords As Collection ' Collection of excluded leave records

'------------------------------------------------------------------------------
' Sub: SP1_PopulateAttendance
' Purpose: Main routine to populate Attendance sheet with leave data
'------------------------------------------------------------------------------
Public Sub SP1_PopulateAttendance(flexWb As Workbook)
    Dim ws As Worksheet
    Dim leaveRecords As Collection
    Dim empIndex As Object
    
    On Error GoTo ErrHandler
    
    LogInfo "modSP1_Attendance", "SP1_PopulateAttendance", "Starting attendance population"
    
    Set ws = flexWb.Worksheets("Attendance")
    
    ' Build employee index for the Attendance sheet (try multiple field name variants)
    Set empIndex = BuildEmployeeIndex(ws, "Employee Code,EmployeeCode,Employee Reference,EmployeeNumber,Employee Number")
    
    ' Load leave transactions
    Set leaveRecords = LoadLeaveTransactions()
    
    If leaveRecords.count = 0 Then
        LogWarning "modSP1_Attendance", "SP1_PopulateAttendance", "No leave records found"
        Exit Sub
    End If
    
    ' Process each leave type
    ProcessAnnualLeave ws, leaveRecords, empIndex
    ProcessSickLeave ws, leaveRecords, empIndex
    ProcessUnpaidLeave ws, leaveRecords, empIndex
    ProcessPPTO ws, leaveRecords, empIndex
    ProcessMaternityLeave ws, leaveRecords, empIndex
    ProcessPaternityLeave ws, leaveRecords, empIndex
    
    LogInfo "modSP1_Attendance", "SP1_PopulateAttendance", "Attendance population completed"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Attendance", "SP1_PopulateAttendance", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Function: LoadLeaveTransactions
' Purpose: Load and filter leave transactions from input file
' Returns: Collection of tLeaveRecord
'------------------------------------------------------------------------------
Private Function LoadLeaveTransactions() As Collection
    Dim col As New Collection
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filePath As String
    Dim lastRow As Long, i As Long
    Dim rec(0 To 9) As Variant  ' Array to store leave record
    Dim headers As Object
    Dim uniqueKey As String
    Dim recStatus As String
    
    On Error GoTo ErrHandler
    
    filePath = GetInputFilePath("EmployeeLeave")
    
    If Dir(filePath) = "" Then
        LogWarning "modSP1_Attendance", "LoadLeaveTransactions", "File not found: " & filePath
        Set LoadLeaveTransactions = col
        Exit Function
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Worksheets(1)
    
    ' Build header index
    Set headers = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        headers(UCase(Trim(CStr(ws.Cells(1, c).Value)))) = c
    Next c
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ' Initialize leave history
    If mLeaveHistory Is Nothing Then
        Set mLeaveHistory = CreateObject("Scripting.Dictionary")
    End If
    
    For i = 2 To lastRow
        ' Only process approved records
        recStatus = GetCellValue(ws, i, headers, "STATUS")
        If UCase(recStatus) = "APPROVED" Then
            ' Try multiple WEIN field name variants
            rec(LR_WEIN) = GetCellValue(ws, i, headers, "WIN")
            If rec(LR_WEIN) = "" Then rec(LR_WEIN) = GetCellValue(ws, i, headers, "WEIN")
            If rec(LR_WEIN) = "" Then rec(LR_WEIN) = GetCellValue(ws, i, headers, "WEINEMPLOYEE ID")
            If rec(LR_WEIN) = "" Then rec(LR_WEIN) = GetCellValue(ws, i, headers, "EMPLOYEE CODEWIN")
            
            ' Try multiple Employee Code field name variants
            rec(LR_EMPCODE) = GetCellValue(ws, i, headers, "EMPLOYEE CODE")
            If rec(LR_EMPCODE) = "" Then rec(LR_EMPCODE) = GetCellValue(ws, i, headers, "EMPLOYEECODE")
            If rec(LR_EMPCODE) = "" Then rec(LR_EMPCODE) = GetCellValue(ws, i, headers, "EMPLOYEE REFERENCE")
            If rec(LR_EMPCODE) = "" Then rec(LR_EMPCODE) = GetCellValue(ws, i, headers, "EMPLOYEENUMBER")
            If rec(LR_EMPCODE) = "" Then rec(LR_EMPCODE) = GetCellValue(ws, i, headers, "EMPLOYEE NUMBER")
            If rec(LR_EMPCODE) = "" Then rec(LR_EMPCODE) = GetCellValue(ws, i, headers, "EMPLOYEE NUMBER ID")
            rec(LR_LEAVETYPE) = GetCellValue(ws, i, headers, "LEAVE TYPE")
            rec(LR_STATUS) = recStatus
            
            ' Parse dates
            Dim fromDateVal As Variant, toDateVal As Variant
            Dim applyDateVal As Variant, approvalDateVal As Variant
            
            fromDateVal = GetCellValue(ws, i, headers, "FROM_DATE")
            toDateVal = GetCellValue(ws, i, headers, "TO_DATE")
            applyDateVal = GetCellValue(ws, i, headers, "APPLY_DATE")
            approvalDateVal = GetCellValue(ws, i, headers, "APPROVAL_DATE")
            
            ' Skip records with missing required dates (FROM_DATE and TO_DATE are mandatory)
            If IsEmpty(fromDateVal) Or IsNull(fromDateVal) Or Trim(CStr(fromDateVal)) = "" Then
                GoTo NextRow
            End If
            If IsEmpty(toDateVal) Or IsNull(toDateVal) Or Trim(CStr(toDateVal)) = "" Then
                GoTo NextRow
            End If
            
            On Error Resume Next
            rec(LR_FROMDATE) = CDate(fromDateVal)
            rec(LR_TODATE) = CDate(toDateVal)
            If Not IsEmpty(applyDateVal) And Not IsNull(applyDateVal) And Trim(CStr(applyDateVal)) <> "" Then
                rec(LR_APPLYDATE) = CDate(applyDateVal)
            Else
                rec(LR_APPLYDATE) = #1/1/1900#
            End If
            If Not IsEmpty(approvalDateVal) And Not IsNull(approvalDateVal) And Trim(CStr(approvalDateVal)) <> "" Then
                rec(LR_APPROVALDATE) = CDate(approvalDateVal)
            Else
                rec(LR_APPROVALDATE) = #1/1/1900#
            End If
            On Error GoTo ErrHandler
            
            ' Validate dates were parsed successfully
            If rec(LR_FROMDATE) = 0 Or rec(LR_TODATE) = 0 Then
                GoTo NextRow
            End If
            
            rec(LR_TOTALDAYS) = ToDouble(GetCellValue(ws, i, headers, "TOTAL_DAYS"))
            
            ' Build unique key
            uniqueKey = rec(LR_WEIN) & "|" & Format(rec(LR_FROMDATE), "YYYYMMDD") & "|" & _
                       Format(rec(LR_TODATE), "YYYYMMDD") & "|" & Format(rec(LR_APPLYDATE), "YYYYMMDD") & "|" & _
                       Format(rec(LR_APPROVALDATE), "YYYYMMDD")
            rec(LR_UNIQUEKEY) = uniqueKey
            
            ' Only add if not already processed
            If Not mLeaveHistory.exists(uniqueKey) Then
                col.Add rec
            End If
NextRow:
        End If
    Next i
    
    wb.Close SaveChanges:=False
    
    LogInfo "modSP1_Attendance", "LoadLeaveTransactions", "Loaded " & col.count & " new leave records"
    
    Set LoadLeaveTransactions = col
    Exit Function
    
ErrHandler:
    LogError "modSP1_Attendance", "LoadLeaveTransactions", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set LoadLeaveTransactions = col
End Function

'------------------------------------------------------------------------------
' Sub: ProcessAnnualLeave
' Purpose: Process Annual Leave records
'------------------------------------------------------------------------------
Private Sub ProcessAnnualLeave(ws As Worksheet, leaveRecords As Collection, empIndex As Object)
    Dim rec As Variant
    Dim v As Variant
    Dim row As Long
    Dim currentMonthDays As Double, prevMonthDays As Double, olderDays As Double
    Dim empDays As Object ' Dictionary to aggregate by employee
    Dim recWein As String, recUniqueKey As String
    Dim prevYM As String
    Dim arr As Variant
    
    On Error GoTo ErrHandler
    
    Set empDays = CreateObject("Scripting.Dictionary")
    prevYM = Format(G.Payroll.prevMonthStart, "YYYYMM")
    
    ' Process each annual leave record
    For Each v In leaveRecords
        rec = v
        
        If UCase(CStr(rec(LR_LEAVETYPE))) Like "*ANNUAL*" Then
            ' Calculate business days by month directly
            CalcBusinessDaysByMonth CDate(rec(LR_FROMDATE)), CDate(rec(LR_TODATE)), _
                                    G.Payroll.payrollMonth, prevYM, _
                                    currentMonthDays, prevMonthDays, olderDays
            
            ' Aggregate by employee
            recWein = CStr(rec(LR_WEIN))
            If Not empDays.exists(recWein) Then
                empDays.Add recWein, Array(0#, 0#, 0#)
            End If
            
            arr = empDays(recWein)
            arr(0) = arr(0) + currentMonthDays
            arr(1) = arr(1) + prevMonthDays
            arr(2) = arr(2) + olderDays
            empDays(recWein) = arr
            
            ' Mark as processed
            recUniqueKey = CStr(rec(LR_UNIQUEKEY))
            mLeaveHistory.Add recUniqueKey, True
        End If
    Next v
    
    ' Write to Attendance sheet
    Dim wein As Variant
    Dim colCurrent As Long, colPrev As Long, colDeduction As Long, colDeductionLast As Long
    
    colCurrent = FindColumnByHeader(ws.Rows(1), "Days_AnnualLeave")
    colPrev = FindColumnByHeader(ws.Rows(1), "Days_AnnualLeave_LastMonth")
    colDeduction = FindColumnByHeader(ws.Rows(1), "Days_AnnualLeaveForDeduction")
    colDeductionLast = FindColumnByHeader(ws.Rows(1), "Days_AnnualLeaveForDeduction_LastMonth")
    
    For Each wein In empDays.Keys
        row = GetOrAddEmployeeRow(ws, CStr(wein), empIndex)
        If row > 0 Then
            arr = empDays(wein)
            ' Current month days
            If colCurrent > 0 Then ws.Cells(row, colCurrent).Value = RoundAmount2(arr(0))
            ' Previous month days
            If colPrev > 0 Then ws.Cells(row, colPrev).Value = RoundAmount2(arr(1))
            ' Current month deduction (same as current month days per requirement)
            If colDeduction > 0 Then ws.Cells(row, colDeduction).Value = RoundAmount2(arr(0))
            ' Previous month deduction (same as previous month days per requirement)
            If colDeductionLast > 0 Then ws.Cells(row, colDeductionLast).Value = RoundAmount2(arr(1))
        End If
    Next wein
    
    ' Write Annual Leave EAO Adj_Input to VariablePay for older periods (before previous month)
    WriteAnnualLeaveEAOAdj empDays
    
    LogInfo "modSP1_Attendance", "ProcessAnnualLeave", "Processed " & empDays.count & " employees"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Attendance", "ProcessAnnualLeave", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: ProcessSickLeave
' Purpose: Process Sick Leave records (requires 4+ consecutive business days)
' Output columns:
'   - Days_SickLeave (current month)
'   - Days_SickLeaveForDeduction (current month)
'   - Days_SickLeave_LastMonth (previous month)
'   - Days_SickLeaveForDeduction_LastMonth (previous month)
'   - Sick Leave EAO Adj_Input (for older periods, written to VariablePay)
'------------------------------------------------------------------------------
Private Sub ProcessSickLeave(ws As Worksheet, leaveRecords As Collection, empIndex As Object)
    Dim rec As Variant
    Dim v As Variant
    Dim row As Long
    Dim empDays As Object
    Dim recWein As String, recUniqueKey As String
    Dim currentDays As Double, prevDays As Double, olderDays As Double
    Dim prevYM As String
    Dim arr As Variant
    
    On Error GoTo ErrHandler
    
    Set empDays = CreateObject("Scripting.Dictionary")
    prevYM = Format(G.Payroll.prevMonthStart, "YYYYMM")
    
    For Each v In leaveRecords
        rec = v
        
        If UCase(CStr(rec(LR_LEAVETYPE))) Like "*SICK*" Then
            ' Check for 4 consecutive business days requirement
            If HasFourConsecutiveBusinessDays(CDate(rec(LR_FROMDATE)), CDate(rec(LR_TODATE))) Then
                ' Calculate calendar days by month
                CalcDaysByMonth CDate(rec(LR_FROMDATE)), CDate(rec(LR_TODATE)), _
                                G.Payroll.payrollMonth, prevYM, _
                                currentDays, prevDays, olderDays
                
                ' Aggregate by employee: arr(0)=currentDays, arr(1)=prevDays, arr(2)=olderDays
                recWein = CStr(rec(LR_WEIN))
                If Not empDays.exists(recWein) Then
                    empDays.Add recWein, Array(0#, 0#, 0#)
                End If
                
                arr = empDays(recWein)
                arr(0) = arr(0) + currentDays
                arr(1) = arr(1) + prevDays
                arr(2) = arr(2) + olderDays
                empDays(recWein) = arr
            End If
            
            recUniqueKey = CStr(rec(LR_UNIQUEKEY))
            If Not mLeaveHistory.exists(recUniqueKey) Then
                mLeaveHistory.Add recUniqueKey, True
            End If
        End If
    Next v
    
    ' Write to Attendance sheet
    Dim wein As Variant
    Dim colCurrent As Long, colPrev As Long
    Dim colDeduction As Long, colDeductionLast As Long
    
    colCurrent = FindColumnByHeader(ws.Rows(1), "Days_SickLeave")
    colPrev = FindColumnByHeader(ws.Rows(1), "Days_SickLeave_LastMonth")
    colDeduction = FindColumnByHeader(ws.Rows(1), "Days_SickLeaveForDeduction")
    colDeductionLast = FindColumnByHeader(ws.Rows(1), "Days_SickLeaveForDeduction_LastMonth")
    
    For Each wein In empDays.Keys
        row = GetOrAddEmployeeRow(ws, CStr(wein), empIndex)
        If row > 0 Then
            arr = empDays(wein)
            ' Current month days
            If colCurrent > 0 Then ws.Cells(row, colCurrent).Value = RoundAmount2(arr(0))
            ' Current month deduction (same as current month days)
            If colDeduction > 0 Then ws.Cells(row, colDeduction).Value = RoundAmount2(arr(0))
            ' Previous month days
            If colPrev > 0 Then ws.Cells(row, colPrev).Value = RoundAmount2(arr(1))
            ' Previous month deduction (same as previous month days)
            If colDeductionLast > 0 Then ws.Cells(row, colDeductionLast).Value = RoundAmount2(arr(1))
        End If
    Next wein
    
    ' Write Sick Leave EAO Adj_Input to VariablePay for older periods (before previous month)
    WriteSickLeaveEAOAdj empDays
    
    LogInfo "modSP1_Attendance", "ProcessSickLeave", "Processed " & empDays.count & " employees"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Attendance", "ProcessSickLeave", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: ProcessUnpaidLeave
' Purpose: Process Unpaid/No Pay Leave records
' Output columns:
'   - Days_NoPayLeave (current month)
'   - Days_NoPayLeave_LastMonth (previous month)
'   - Annual Leave EAO Adj_Input (for older periods, written to VariablePay)
'------------------------------------------------------------------------------
Private Sub ProcessUnpaidLeave(ws As Worksheet, leaveRecords As Collection, empIndex As Object)
    Dim rec As Variant
    Dim v As Variant
    Dim row As Long
    Dim empDays As Object
    Dim recWein As String, recUniqueKey As String, recLeaveType As String
    Dim currentDays As Double, prevDays As Double, olderDays As Double
    Dim prevYM As String
    Dim arr As Variant
    
    On Error GoTo ErrHandler
    
    Set empDays = CreateObject("Scripting.Dictionary")
    prevYM = Format(G.Payroll.prevMonthStart, "YYYYMM")
    
    For Each v In leaveRecords
        rec = v
        recLeaveType = UCase(CStr(rec(LR_LEAVETYPE)))
        
        If recLeaveType Like "*UNPAID*" Or recLeaveType Like "*NO PAY*" Then
            ' Calculate calendar days by month
            CalcDaysByMonth CDate(rec(LR_FROMDATE)), CDate(rec(LR_TODATE)), _
                            G.Payroll.payrollMonth, prevYM, _
                            currentDays, prevDays, olderDays
            
            ' Aggregate by employee: arr(0)=currentDays, arr(1)=prevDays, arr(2)=olderDays
            recWein = CStr(rec(LR_WEIN))
            If Not empDays.exists(recWein) Then
                empDays.Add recWein, Array(0#, 0#, 0#)
            End If
            
            arr = empDays(recWein)
            arr(0) = arr(0) + currentDays
            arr(1) = arr(1) + prevDays
            arr(2) = arr(2) + olderDays
            empDays(recWein) = arr
            
            recUniqueKey = CStr(rec(LR_UNIQUEKEY))
            If Not mLeaveHistory.exists(recUniqueKey) Then
                mLeaveHistory.Add recUniqueKey, True
            End If
        End If
    Next v
    
    ' Write to Attendance sheet
    Dim wein As Variant
    Dim colCurrent As Long, colPrev As Long
    
    colCurrent = FindColumnByHeader(ws.Rows(1), "Days_NoPayLeave")
    colPrev = FindColumnByHeader(ws.Rows(1), "Days_NoPayLeave_LastMonth")
    
    For Each wein In empDays.Keys
        row = GetOrAddEmployeeRow(ws, CStr(wein), empIndex)
        If row > 0 Then
            arr = empDays(wein)
            If colCurrent > 0 Then ws.Cells(row, colCurrent).Value = RoundAmount2(arr(0))
            If colPrev > 0 Then ws.Cells(row, colPrev).Value = RoundAmount2(arr(1))
        End If
    Next wein
    
    ' Write No Pay Leave Deduction to VariablePay for older periods (before previous month)
    WriteNoPayLeaveDeduction empDays
    
    LogInfo "modSP1_Attendance", "ProcessUnpaidLeave", "Processed " & empDays.count & " employees"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Attendance", "ProcessUnpaidLeave", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: ProcessPPTO
' Purpose: Process Paid Parental Time Off records
' Output columns:
'   - Days_Paid Parental Time Off (current month)
'   - Days_Paid Parental Time Off ForDeduction (current month)
'   - Days_Paid Parental Time Off_LastMonth (previous month)
'   - Days_Paid Parental Time Off ForDeduction_LastMonth (previous month)
'   - PPTO EAO Adj_Input (for older periods, written to VariablePay)
'------------------------------------------------------------------------------
Private Sub ProcessPPTO(ws As Worksheet, leaveRecords As Collection, empIndex As Object)
    Dim rec As Variant
    Dim v As Variant
    Dim row As Long
    Dim empDays As Object
    Dim recWein As String, recUniqueKey As String, recLeaveType As String
    Dim currentDays As Double, prevDays As Double, olderDays As Double
    Dim prevYM As String
    Dim arr As Variant
    
    On Error GoTo ErrHandler
    
    Set empDays = CreateObject("Scripting.Dictionary")
    prevYM = Format(G.Payroll.prevMonthStart, "YYYYMM")
    
    For Each v In leaveRecords
        rec = v
        recLeaveType = UCase(CStr(rec(LR_LEAVETYPE)))
        
        If recLeaveType Like "*PPTO*" Or recLeaveType Like "*PARENTAL TIME OFF*" Then
            ' Calculate calendar days by month
            CalcDaysByMonth CDate(rec(LR_FROMDATE)), CDate(rec(LR_TODATE)), _
                            G.Payroll.payrollMonth, prevYM, _
                            currentDays, prevDays, olderDays
            
            recWein = CStr(rec(LR_WEIN))
            If Not empDays.exists(recWein) Then
                empDays.Add recWein, Array(0#, 0#, 0#)
            End If
            
            ' Aggregate by employee: arr(0)=currentDays, arr(1)=prevDays, arr(2)=olderDays
            arr = empDays(recWein)
            arr(0) = arr(0) + currentDays
            arr(1) = arr(1) + prevDays
            arr(2) = arr(2) + olderDays
            empDays(recWein) = arr
            
            recUniqueKey = CStr(rec(LR_UNIQUEKEY))
            If Not mLeaveHistory.exists(recUniqueKey) Then
                mLeaveHistory.Add recUniqueKey, True
            End If
        End If
    Next v
    
    ' Write to Attendance sheet
    Dim wein As Variant
    Dim colCurrent As Long, colPrev As Long
    Dim colDeduction As Long, colDeductionLast As Long
    
    colCurrent = FindColumnByHeader(ws.Rows(1), "Days_Paid Parental Time Off")
    colPrev = FindColumnByHeader(ws.Rows(1), "Days_Paid Parental Time Off_LastMonth")
    colDeduction = FindColumnByHeader(ws.Rows(1), "Days_Paid Parental Time Off ForDeduction")
    colDeductionLast = FindColumnByHeader(ws.Rows(1), "Days_Paid Parental Time Off ForDeduction_LastMonth")
    
    For Each wein In empDays.Keys
        row = GetOrAddEmployeeRow(ws, CStr(wein), empIndex)
        If row > 0 Then
            arr = empDays(wein)
            ' Current month days
            If colCurrent > 0 Then ws.Cells(row, colCurrent).Value = RoundAmount2(arr(0))
            ' Current month deduction (same as current month days per requirement)
            If colDeduction > 0 Then ws.Cells(row, colDeduction).Value = RoundAmount2(arr(0))
            ' Previous month days
            If colPrev > 0 Then ws.Cells(row, colPrev).Value = RoundAmount2(arr(1))
            ' Previous month deduction (same as previous month days per requirement)
            If colDeductionLast > 0 Then ws.Cells(row, colDeductionLast).Value = RoundAmount2(arr(1))
        End If
    Next wein
    
    ' Write PPTO EAO Adj_Input to VariablePay for older periods (before previous month)
    WritePPTOEAOAdj empDays
    
    LogInfo "modSP1_Attendance", "ProcessPPTO", "Processed " & empDays.count & " employees"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Attendance", "ProcessPPTO", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: ProcessMaternityLeave
' Purpose: Process Maternity Leave records
' Output columns:
'   - Days_MaternityLeave (current month)
'   - Days_MaternityLeaveForDeduction (current month)
'   - Days_MaternityLeave_LastMonth (previous month)
'   - Days_MaternityLeaveForDeduction_LastMonth (previous month)
'   - Maternity Leave EAO Adj_Input (for older periods, written to VariablePay)
' Note: Requires 40 weeks service before maternity leave start date
'       If <40 weeks, record is excluded and written to Maternity Report
'------------------------------------------------------------------------------
Private Sub ProcessMaternityLeave(ws As Worksheet, leaveRecords As Collection, empIndex As Object)
    Dim rec As Variant
    Dim v As Variant
    Dim row As Long
    Dim empDays As Object
    Dim recWein As String, recUniqueKey As String
    Dim currentDays As Double, prevDays As Double, olderDays As Double
    Dim prevYM As String
    Dim arr As Variant
    Dim lastHireDate As Date
    Dim leaveStartDate As Date
    Dim weeksOfService As Double
    Dim excludedCount As Long
    
    On Error GoTo ErrHandler
    
    Set empDays = CreateObject("Scripting.Dictionary")
    Set mMaternityExcludedRecords = New Collection
    prevYM = Format(G.Payroll.prevMonthStart, "YYYYMM")
    excludedCount = 0
    
    ' Ensure Workforce Hire Dates are loaded
    If mWorkforceHireDates Is Nothing Then
        LoadWorkforceHireDates
    End If
    
    For Each v In leaveRecords
        rec = v
        
        If UCase(CStr(rec(LR_LEAVETYPE))) Like "*MATERNITY*" Then
            recWein = CStr(rec(LR_WEIN))
            leaveStartDate = CDate(rec(LR_FROMDATE))
            
            ' Check 40 weeks service requirement
            lastHireDate = GetEmployeeLastHireDate(recWein)
            
            If lastHireDate > 0 Then
                ' Calculate weeks of service before maternity leave start
                weeksOfService = (leaveStartDate - lastHireDate) / 7
                
                If weeksOfService < 40 Then
                    ' Employee has less than 40 weeks service - exclude from payroll
                    mMaternityExcludedRecords.Add rec
                    excludedCount = excludedCount + 1
                    
                    ' Mark as processed but don't add to empDays
                    recUniqueKey = CStr(rec(LR_UNIQUEKEY))
                    If Not mLeaveHistory.exists(recUniqueKey) Then
                        mLeaveHistory.Add recUniqueKey, True
                    End If
                    
                    LogInfo "modSP1_Attendance", "ProcessMaternityLeave", _
                        "Excluded WEIN " & recWein & " - only " & Format(weeksOfService, "0.0") & " weeks service (< 40 weeks required)"
                    
                    GoTo NextRecord
                End If
            End If
            
            ' Employee has 40+ weeks service - process normally
            ' Calculate calendar days by month
            CalcDaysByMonth leaveStartDate, CDate(rec(LR_TODATE)), _
                            G.Payroll.payrollMonth, prevYM, _
                            currentDays, prevDays, olderDays
            
            ' Aggregate by employee: arr(0)=currentDays, arr(1)=prevDays, arr(2)=olderDays
            If Not empDays.exists(recWein) Then
                empDays.Add recWein, Array(0#, 0#, 0#)
            End If
            
            arr = empDays(recWein)
            arr(0) = arr(0) + currentDays
            arr(1) = arr(1) + prevDays
            arr(2) = arr(2) + olderDays
            empDays(recWein) = arr
            
            recUniqueKey = CStr(rec(LR_UNIQUEKEY))
            If Not mLeaveHistory.exists(recUniqueKey) Then
                mLeaveHistory.Add recUniqueKey, True
            End If
        End If
NextRecord:
    Next v
    
    ' Write to Attendance sheet
    Dim wein As Variant
    Dim colCurrent As Long, colPrev As Long
    Dim colDeduction As Long, colDeductionLast As Long
    
    colCurrent = FindColumnByHeader(ws.Rows(1), "Days_MaternityLeave")
    colPrev = FindColumnByHeader(ws.Rows(1), "Days_MaternityLeave_LastMonth")
    colDeduction = FindColumnByHeader(ws.Rows(1), "Days_MaternityLeaveForDeduction")
    colDeductionLast = FindColumnByHeader(ws.Rows(1), "Days_MaternityLeaveForDeduction_LastMonth")
    
    For Each wein In empDays.Keys
        row = GetOrAddEmployeeRow(ws, CStr(wein), empIndex)
        If row > 0 Then
            arr = empDays(wein)
            ' Current month days
            If colCurrent > 0 Then ws.Cells(row, colCurrent).Value = RoundAmount2(arr(0))
            ' Current month deduction (same as current month days)
            If colDeduction > 0 Then ws.Cells(row, colDeduction).Value = RoundAmount2(arr(0))
            ' Previous month days
            If colPrev > 0 Then ws.Cells(row, colPrev).Value = RoundAmount2(arr(1))
            ' Previous month deduction (same as previous month days)
            If colDeductionLast > 0 Then ws.Cells(row, colDeductionLast).Value = RoundAmount2(arr(1))
        End If
    Next wein
    
    ' Write Maternity Leave EAO Adj_Input to VariablePay for older periods (before previous month)
    WriteMaternityLeaveEAOAdj empDays
    
    ' Output Maternity Report for excluded records
    If excludedCount > 0 Then
        OutputMaternityReport
    End If
    
    LogInfo "modSP1_Attendance", "ProcessMaternityLeave", _
        "Processed " & empDays.count & " employees, excluded " & excludedCount & " records (< 40 weeks service)"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Attendance", "ProcessMaternityLeave", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: ProcessPaternityLeave
' Purpose: Process Paternity Leave records
' Output columns:
'   - Days_PaternityLeave (current month)
'   - Days_PaternityLeaveForDeduction (current month)
'   - Days_PaternityLeave_LastMonth (previous month)
'   - Days_PaternityLeaveForDeduction_LastMonth (previous month)
'   - Paternity Leave EAO Adj_Input (for older periods, written to VariablePay)
'------------------------------------------------------------------------------
Private Sub ProcessPaternityLeave(ws As Worksheet, leaveRecords As Collection, empIndex As Object)
    Dim rec As Variant
    Dim v As Variant
    Dim row As Long
    Dim empDays As Object
    Dim recWein As String, recUniqueKey As String
    Dim currentDays As Double, prevDays As Double, olderDays As Double
    Dim prevYM As String
    Dim arr As Variant
    
    On Error GoTo ErrHandler
    
    Set empDays = CreateObject("Scripting.Dictionary")
    prevYM = Format(G.Payroll.prevMonthStart, "YYYYMM")
    
    For Each v In leaveRecords
        rec = v
        
        If UCase(CStr(rec(LR_LEAVETYPE))) Like "*PATERNITY*" Then
            ' Calculate calendar days by month
            CalcDaysByMonth CDate(rec(LR_FROMDATE)), CDate(rec(LR_TODATE)), _
                            G.Payroll.payrollMonth, prevYM, _
                            currentDays, prevDays, olderDays
            
            ' Aggregate by employee: arr(0)=currentDays, arr(1)=prevDays, arr(2)=olderDays
            recWein = CStr(rec(LR_WEIN))
            If Not empDays.exists(recWein) Then
                empDays.Add recWein, Array(0#, 0#, 0#)
            End If
            
            arr = empDays(recWein)
            arr(0) = arr(0) + currentDays
            arr(1) = arr(1) + prevDays
            arr(2) = arr(2) + olderDays
            empDays(recWein) = arr
            
            recUniqueKey = CStr(rec(LR_UNIQUEKEY))
            If Not mLeaveHistory.exists(recUniqueKey) Then
                mLeaveHistory.Add recUniqueKey, True
            End If
        End If
    Next v
    
    ' Write to Attendance sheet
    Dim wein As Variant
    Dim colCurrent As Long, colPrev As Long
    Dim colDeduction As Long, colDeductionLast As Long
    
    colCurrent = FindColumnByHeader(ws.Rows(1), "Days_PaternityLeave")
    colPrev = FindColumnByHeader(ws.Rows(1), "Days_PaternityLeave_LastMonth")
    colDeduction = FindColumnByHeader(ws.Rows(1), "Days_PaternityLeaveForDeduction")
    colDeductionLast = FindColumnByHeader(ws.Rows(1), "Days_PaternityLeaveForDeduction_LastMonth")
    
    For Each wein In empDays.Keys
        row = GetOrAddEmployeeRow(ws, CStr(wein), empIndex)
        If row > 0 Then
            arr = empDays(wein)
            ' Current month days
            If colCurrent > 0 Then ws.Cells(row, colCurrent).Value = RoundAmount2(arr(0))
            ' Current month deduction (same as current month days)
            If colDeduction > 0 Then ws.Cells(row, colDeduction).Value = RoundAmount2(arr(0))
            ' Previous month days
            If colPrev > 0 Then ws.Cells(row, colPrev).Value = RoundAmount2(arr(1))
            ' Previous month deduction (same as previous month days)
            If colDeductionLast > 0 Then ws.Cells(row, colDeductionLast).Value = RoundAmount2(arr(1))
        End If
    Next wein
    
    ' Write Paternity Leave EAO Adj_Input to VariablePay for older periods (before previous month)
    WritePaternityLeaveEAOAdj empDays
    
    LogInfo "modSP1_Attendance", "ProcessPaternityLeave", "Processed " & empDays.count & " employees"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Attendance", "ProcessPaternityLeave", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Helper Functions
'------------------------------------------------------------------------------

Private Function GetCellValue(ws As Worksheet, row As Long, headers As Object, headerName As String) As String
    Dim col As Long
    GetCellValue = ""
    
    If headers.exists(UCase(headerName)) Then
        col = headers(UCase(headerName))
        GetCellValue = Trim(CStr(Nz(ws.Cells(row, col).Value, "")))
    End If
End Function

Private Function GetOrAddEmployeeRow(ws As Worksheet, wein As String, empIndex As Object) As Long
    Dim empCodeCol As Long
    Dim newRow As Long
    
    ' Try to find by WEIN in index
    If empIndex.exists(wein) Then
        GetOrAddEmployeeRow = empIndex(wein)
        Exit Function
    End If
    
    ' Try Employee Code
    Dim empCode As String
    empCode = EmpCodeFromWein(wein)
    If empCode <> "" And empIndex.exists(empCode) Then
        GetOrAddEmployeeRow = empIndex(empCode)
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
    
    GetOrAddEmployeeRow = newRow
End Function

'------------------------------------------------------------------------------
' Sub: WriteAnnualLeaveEAOAdj
' Purpose: Write Annual Leave EAO Adj_Input to VariablePay for older periods
' Description: For Annual Leave records dated before the previous month,
'              calculate EAO adjustment using formula:
'              (AverageDayWage_12Month - DailySalary) * TOTAL_DAYS
'              and write to VariablePay.Annual Leave EAO Adj_Input
'------------------------------------------------------------------------------
Private Sub WriteAnnualLeaveEAOAdj(empDays As Object)
    Dim ws As Worksheet
    Dim empIndex As Object
    Dim wein As Variant
    Dim arr As Variant
    Dim olderDays As Double
    Dim eaoAdj As Double
    Dim col As Long, row As Long
    
    On Error GoTo ErrHandler
    
    ' Get VariablePay sheet from the flexi output workbook
    On Error Resume Next
    Set ws = G.FlexiOutputWb.Worksheets("VariablePay")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then
        LogWarning "modSP1_Attendance", "WriteAnnualLeaveEAOAdj", "VariablePay sheet not found"
        Exit Sub
    End If
    
    ' Build employee index for VariablePay sheet
    Set empIndex = BuildEmployeeIndex(ws, "Employee Code,EmployeeCode,Employee Reference,EmployeeNumber,Employee Number")
    
    ' Find target column
    col = FindColumnByHeader(ws.Rows(1), "Annual Leave EAO Adj_Input")
    
    If col = 0 Then
        LogWarning "modSP1_Attendance", "WriteAnnualLeaveEAOAdj", "Annual Leave EAO Adj_Input column not found"
        Exit Sub
    End If
    
    ' Process each employee with older period days
    For Each wein In empDays.Keys
        arr = empDays(wein)
        olderDays = arr(2)  ' Days from periods before previous month
        
        If olderDays > 0 Then
            ' Calculate EAO adjustment: (AverageDayWage_12Month - DailySalary) * TOTAL_DAYS
            eaoAdj = CalcAnnualLeaveEAOAdj(CStr(wein), olderDays)
            
            If eaoAdj > 0 Then
                row = GetOrAddEmployeeRow(ws, CStr(wein), empIndex)
                If row > 0 Then
                    ' Add to existing value (in case there are multiple entries)
                    ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, eaoAdj)
                End If
            End If
        End If
    Next wein
    
    LogInfo "modSP1_Attendance", "WriteAnnualLeaveEAOAdj", "Processed Annual Leave EAO adjustments"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Attendance", "WriteAnnualLeaveEAOAdj", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: WriteSickLeaveEAOAdj
' Purpose: Write Sick Leave EAO Adj_Input to VariablePay for older periods
' Description: For Sick Leave records dated before the previous month,
'              calculate EAO adjustment using formula:
'              (DayWage_Maternity/Paternity/Sick Leave - DailySalary) * Days_SickLeave
'              and write to VariablePay.Sick Leave EAO Adj_Input
'------------------------------------------------------------------------------
Private Sub WriteSickLeaveEAOAdj(empDays As Object)
    Dim ws As Worksheet
    Dim empIndex As Object
    Dim wein As Variant
    Dim arr As Variant
    Dim olderDays As Double
    Dim eaoAdj As Double
    Dim col As Long, row As Long
    
    On Error GoTo ErrHandler
    
    ' Get VariablePay sheet from the flexi output workbook
    On Error Resume Next
    Set ws = G.FlexiOutputWb.Worksheets("VariablePay")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then
        LogWarning "modSP1_Attendance", "WriteSickLeaveEAOAdj", "VariablePay sheet not found"
        Exit Sub
    End If
    
    ' Build employee index for VariablePay sheet
    Set empIndex = BuildEmployeeIndex(ws, "Employee Code,EmployeeCode,Employee Reference,EmployeeNumber,Employee Number")
    
    ' Find target column
    col = FindColumnByHeader(ws.Rows(1), "Sick Leave EAO Adj_Input")
    
    If col = 0 Then
        LogWarning "modSP1_Attendance", "WriteSickLeaveEAOAdj", "Sick Leave EAO Adj_Input column not found"
        Exit Sub
    End If
    
    ' Process each employee with older period days
    For Each wein In empDays.Keys
        arr = empDays(wein)
        olderDays = arr(2)  ' Days from periods before previous month
        
        If olderDays > 0 Then
            ' Calculate EAO adjustment: (DayWage_Maternity/Paternity/Sick Leave - DailySalary) * Days_SickLeave
            eaoAdj = CalcSickLeaveEAOAdj(CStr(wein), olderDays)
            
            If eaoAdj <> 0 Then
                row = GetOrAddEmployeeRow(ws, CStr(wein), empIndex)
                If row > 0 Then
                    ' Add to existing value (in case there are multiple entries)
                    ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, eaoAdj)
                End If
            End If
        End If
    Next wein
    
    LogInfo "modSP1_Attendance", "WriteSickLeaveEAOAdj", "Processed Sick Leave EAO adjustments"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Attendance", "WriteSickLeaveEAOAdj", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: WriteNoPayLeaveDeduction
' Purpose: Write No Pay Leave Deduction to VariablePay for older periods
' Description: For Unpaid Leave records dated before the previous month,
'              calculate deduction using formula:
'              NoPayLeaveCalculationBase * Days_NoPayLeave
'              and write to VariablePay.Annual Leave EAO Adj_Input
'              (if cell non-empty, add to existing value)
'------------------------------------------------------------------------------
Private Sub WriteNoPayLeaveDeduction(empDays As Object)
    Dim ws As Worksheet
    Dim empIndex As Object
    Dim wein As Variant
    Dim arr As Variant
    Dim olderDays As Double
    Dim deduction As Double
    Dim col As Long, row As Long
    
    On Error GoTo ErrHandler
    
    ' Get VariablePay sheet from the flexi output workbook
    On Error Resume Next
    Set ws = G.FlexiOutputWb.Worksheets("VariablePay")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then
        LogWarning "modSP1_Attendance", "WriteNoPayLeaveDeduction", "VariablePay sheet not found"
        Exit Sub
    End If
    
    ' Build employee index for VariablePay sheet
    Set empIndex = BuildEmployeeIndex(ws, "Employee Code,EmployeeCode,Employee Reference,EmployeeNumber,Employee Number")
    
    ' Find target column (per requirement: No Pay Leave Deduction -> Annual Leave EAO Adj_Input)
    col = FindColumnByHeader(ws.Rows(1), "Annual Leave EAO Adj_Input")
    
    If col = 0 Then
        LogWarning "modSP1_Attendance", "WriteNoPayLeaveDeduction", "Annual Leave EAO Adj_Input column not found"
        Exit Sub
    End If
    
    ' Process each employee with older period days
    For Each wein In empDays.Keys
        arr = empDays(wein)
        olderDays = arr(2)  ' Days from periods before previous month
        
        If olderDays > 0 Then
            ' Calculate No Pay Leave Deduction: NoPayLeaveCalculationBase * Days_NoPayLeave
            deduction = CalcNoPayLeaveDeductionForOlderPeriod(CStr(wein), olderDays)
            
            If deduction <> 0 Then
                row = GetOrAddEmployeeRow(ws, CStr(wein), empIndex)
                If row > 0 Then
                    ' Add to existing value (if cell non-empty, add, not overwrite)
                    ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, deduction)
                End If
            End If
        End If
    Next wein
    
    LogInfo "modSP1_Attendance", "WriteNoPayLeaveDeduction", "Processed No Pay Leave Deduction adjustments"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Attendance", "WriteNoPayLeaveDeduction", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: WritePPTOEAOAdj
' Purpose: Write PPTO EAO Adj_Input to VariablePay for older periods
' Description: For PPTO records dated before the previous month,
'              calculate EAO adjustment using formula:
'              (Max(DailySalary, AverageDayWage_12Month*80%) - DailySalary) * Days_PPTO
'              and write to VariablePay.PPTO EAO Adj_Input
'------------------------------------------------------------------------------
Private Sub WritePPTOEAOAdj(empDays As Object)
    Dim ws As Worksheet
    Dim empIndex As Object
    Dim wein As Variant
    Dim arr As Variant
    Dim olderDays As Double
    Dim eaoAdj As Double
    Dim col As Long, row As Long
    
    On Error GoTo ErrHandler
    
    ' Get VariablePay sheet from the flexi output workbook
    On Error Resume Next
    Set ws = G.FlexiOutputWb.Worksheets("VariablePay")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then
        LogWarning "modSP1_Attendance", "WritePPTOEAOAdj", "VariablePay sheet not found"
        Exit Sub
    End If
    
    ' Build employee index for VariablePay sheet
    Set empIndex = BuildEmployeeIndex(ws, "Employee Code,EmployeeCode,Employee Reference,EmployeeNumber,Employee Number")
    
    ' Find target column
    col = FindColumnByHeader(ws.Rows(1), "PPTO EAO Adj_Input")
    
    If col = 0 Then
        ' Try alternative column name
        col = FindColumnByHeader(ws.Rows(1), "Adjustment of Parental Paid Time Off (PPTO) payment")
    End If
    
    If col = 0 Then
        LogWarning "modSP1_Attendance", "WritePPTOEAOAdj", "PPTO EAO Adj_Input column not found"
        Exit Sub
    End If
    
    ' Process each employee with older period days
    For Each wein In empDays.Keys
        arr = empDays(wein)
        olderDays = arr(2)  ' Days from periods before previous month
        
        If olderDays > 0 Then
            ' Calculate PPTO EAO adjustment:
            ' (Max(DailySalary, AverageDayWage_12Month*80%) - DailySalary) * Days_PPTO
            eaoAdj = CalcPPTOEAOAdj(CStr(wein), olderDays)
            
            If eaoAdj <> 0 Then
                row = GetOrAddEmployeeRow(ws, CStr(wein), empIndex)
                If row > 0 Then
                    ' Add to existing value (in case there are multiple entries)
                    ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, eaoAdj)
                End If
            End If
        End If
    Next wein
    
    LogInfo "modSP1_Attendance", "WritePPTOEAOAdj", "Processed PPTO EAO adjustments"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Attendance", "WritePPTOEAOAdj", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: WriteMaternityLeaveEAOAdj
' Purpose: Write Maternity Leave EAO Adj_Input to VariablePay for older periods
' Description: For Maternity Leave records dated before the previous month,
'              calculate EAO adjustment using formula:
'              (DayWage_Maternity/Paternity/Sick Leave - DailySalary) * Days_MaternityLeave
'              and write to VariablePay.Maternity Leave EAO Adj_Input
'------------------------------------------------------------------------------
Private Sub WriteMaternityLeaveEAOAdj(empDays As Object)
    Dim ws As Worksheet
    Dim empIndex As Object
    Dim wein As Variant
    Dim arr As Variant
    Dim olderDays As Double
    Dim eaoAdj As Double
    Dim col As Long, row As Long
    
    On Error GoTo ErrHandler
    
    ' Get VariablePay sheet from the flexi output workbook
    On Error Resume Next
    Set ws = G.FlexiOutputWb.Worksheets("VariablePay")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then
        LogWarning "modSP1_Attendance", "WriteMaternityLeaveEAOAdj", "VariablePay sheet not found"
        Exit Sub
    End If
    
    ' Build employee index for VariablePay sheet
    Set empIndex = BuildEmployeeIndex(ws, "Employee Code,EmployeeCode,Employee Reference,EmployeeNumber,Employee Number")
    
    ' Find target column
    col = FindColumnByHeader(ws.Rows(1), "Maternity Leave EAO Adj_Input")
    
    If col = 0 Then
        LogWarning "modSP1_Attendance", "WriteMaternityLeaveEAOAdj", "Maternity Leave EAO Adj_Input column not found"
        Exit Sub
    End If
    
    ' Process each employee with older period days
    For Each wein In empDays.Keys
        arr = empDays(wein)
        olderDays = arr(2)  ' Days from periods before previous month
        
        If olderDays > 0 Then
            ' Calculate Maternity Leave EAO adjustment:
            ' (DayWage_Maternity/Paternity/Sick Leave - DailySalary) * Days_MaternityLeave
            eaoAdj = CalcMaternityLeaveEAOAdj(CStr(wein), olderDays)
            
            If eaoAdj <> 0 Then
                row = GetOrAddEmployeeRow(ws, CStr(wein), empIndex)
                If row > 0 Then
                    ' Add to existing value (in case there are multiple entries)
                    ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, eaoAdj)
                End If
            End If
        End If
    Next wein
    
    LogInfo "modSP1_Attendance", "WriteMaternityLeaveEAOAdj", "Processed Maternity Leave EAO adjustments"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Attendance", "WriteMaternityLeaveEAOAdj", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: WritePaternityLeaveEAOAdj
' Purpose: Write Paternity Leave EAO Adj_Input to VariablePay for older periods
' Description: For Paternity Leave records dated before the previous month,
'              calculate EAO adjustment using formula:
'              (DayWage_Maternity/Paternity/Sick Leave - DailySalary) * Days_PaternityLeave
'              and write to VariablePay.Paternity Leave EAO Adj_Input
'------------------------------------------------------------------------------
Private Sub WritePaternityLeaveEAOAdj(empDays As Object)
    Dim ws As Worksheet
    Dim empIndex As Object
    Dim wein As Variant
    Dim arr As Variant
    Dim olderDays As Double
    Dim eaoAdj As Double
    Dim col As Long, row As Long
    
    On Error GoTo ErrHandler
    
    ' Get VariablePay sheet from the flexi output workbook
    On Error Resume Next
    Set ws = G.FlexiOutputWb.Worksheets("VariablePay")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then
        LogWarning "modSP1_Attendance", "WritePaternityLeaveEAOAdj", "VariablePay sheet not found"
        Exit Sub
    End If
    
    ' Build employee index for VariablePay sheet
    Set empIndex = BuildEmployeeIndex(ws, "Employee Code,EmployeeCode,Employee Reference,EmployeeNumber,Employee Number")
    
    ' Find target column
    col = FindColumnByHeader(ws.Rows(1), "Paternity Leave EAO Adj_Input")
    
    If col = 0 Then
        ' Try alternative column name
        col = FindColumnByHeader(ws.Rows(1), "Paternity Leave payment adjustment")
    End If
    
    If col = 0 Then
        LogWarning "modSP1_Attendance", "WritePaternityLeaveEAOAdj", "Paternity Leave EAO Adj_Input column not found"
        Exit Sub
    End If
    
    ' Process each employee with older period days
    For Each wein In empDays.Keys
        arr = empDays(wein)
        olderDays = arr(2)  ' Days from periods before previous month
        
        If olderDays > 0 Then
            ' Calculate Paternity Leave EAO adjustment:
            ' (DayWage_Maternity/Paternity/Sick Leave - DailySalary) * Days_PaternityLeave
            eaoAdj = CalcPaternityLeaveEAOAdj(CStr(wein), olderDays)
            
            If eaoAdj <> 0 Then
                row = GetOrAddEmployeeRow(ws, CStr(wein), empIndex)
                If row > 0 Then
                    ' Add to existing value (in case there are multiple entries)
                    ws.Cells(row, col).Value = SafeAdd2(ws.Cells(row, col).Value, eaoAdj)
                End If
            End If
        End If
    Next wein
    
    LogInfo "modSP1_Attendance", "WritePaternityLeaveEAOAdj", "Processed Paternity Leave EAO adjustments"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Attendance", "WritePaternityLeaveEAOAdj", Err.Number, Err.Description
End Sub


'------------------------------------------------------------------------------
' Sub: LoadWorkforceHireDates
' Purpose: Load Last Hire Date from Workforce Detail for 40-week service check
'------------------------------------------------------------------------------
Private Sub LoadWorkforceHireDates()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filePath As String
    Dim lastRow As Long, lastCol As Long, i As Long, c As Long
    Dim headers As Object
    Dim empId As String, wein As String
    Dim hireDate As Variant
    Dim headerRow As Long
    Dim cellVal As String
    
    On Error GoTo ErrHandler
    
    Set mWorkforceHireDates = CreateObject("Scripting.Dictionary")
    
    filePath = GetInputFilePath("WorkforceDetail")
    
    If Dir(filePath) = "" Then
        LogWarning "modSP1_Attendance", "LoadWorkforceHireDates", _
            "Workforce Detail file not found: " & filePath
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Worksheets(1)
    
    ' Find header row (search for "Employee ID" in first 50 rows)
    headerRow = 0
    For i = 1 To 50
        For c = 1 To 200
            On Error Resume Next
            cellVal = UCase(Trim(CStr(ws.Cells(i, c).Value)))
            On Error GoTo ErrHandler
            If cellVal = "EMPLOYEE ID" Or cellVal = "WIN" Or cellVal = "WEIN" Then
                headerRow = i
                Exit For
            End If
        Next c
        If headerRow > 0 Then Exit For
    Next i
    
    If headerRow = 0 Then headerRow = 1
    
    ' Build header index
    Set headers = CreateObject("Scripting.Dictionary")
    lastCol = ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column
    
    For c = 1 To lastCol
        cellVal = UCase(Trim(CStr(Nz(ws.Cells(headerRow, c).Value, ""))))
        If cellVal <> "" And Not headers.exists(cellVal) Then
            headers(cellVal) = c
        End If
    Next c
    
    ' Find Employee ID column for determining last row
    Dim empIdCol As Long
    empIdCol = 0
    If headers.exists("EMPLOYEE ID") Then empIdCol = headers("EMPLOYEE ID")
    If empIdCol = 0 And headers.exists("WIN") Then empIdCol = headers("WIN")
    If empIdCol = 0 And headers.exists("WEIN") Then empIdCol = headers("WEIN")
    
    If empIdCol = 0 Then
        LogWarning "modSP1_Attendance", "LoadWorkforceHireDates", _
            "Cannot find Employee ID column in Workforce Detail"
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, empIdCol).End(xlUp).row
    
    ' Load data
    For i = headerRow + 1 To lastRow
        ' Get Employee ID / WEIN
        empId = ""
        wein = ""
        
        If headers.exists("EMPLOYEE ID") Then
            empId = Trim(CStr(Nz(ws.Cells(i, headers("EMPLOYEE ID")).Value, "")))
        End If
        If headers.exists("WIN") Then
            wein = Trim(CStr(Nz(ws.Cells(i, headers("WIN")).Value, "")))
        End If
        If wein = "" And headers.exists("WEIN") Then
            wein = Trim(CStr(Nz(ws.Cells(i, headers("WEIN")).Value, "")))
        End If
        
        ' Get Last Hire Date
        hireDate = Empty
        If headers.exists("LAST HIRE DATE") Then
            On Error Resume Next
            hireDate = ws.Cells(i, headers("LAST HIRE DATE")).Value
            On Error GoTo ErrHandler
        End If
        
        ' Store by both WEIN and Employee ID for flexible lookup
        If IsDate(hireDate) Then
            If wein <> "" And Not mWorkforceHireDates.exists(wein) Then
                mWorkforceHireDates.Add wein, CDate(hireDate)
            End If
            If empId <> "" And Not mWorkforceHireDates.exists(empId) Then
                mWorkforceHireDates.Add empId, CDate(hireDate)
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    
    LogInfo "modSP1_Attendance", "LoadWorkforceHireDates", _
        "Loaded " & mWorkforceHireDates.count & " hire date records"
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Attendance", "LoadWorkforceHireDates", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Function: GetEmployeeLastHireDate
' Purpose: Get Last Hire Date for an employee
' Parameters:
'   empIdOrWein - Employee ID or WEIN
' Returns: Last Hire Date or 0 if not found
'------------------------------------------------------------------------------
Private Function GetEmployeeLastHireDate(empIdOrWein As String) As Date
    Dim empCode As String
    Dim empId As String
    
    GetEmployeeLastHireDate = 0
    
    If mWorkforceHireDates Is Nothing Then
        LoadWorkforceHireDates
    End If
    
    If mWorkforceHireDates Is Nothing Then Exit Function
    
    ' Try direct lookup
    If mWorkforceHireDates.exists(empIdOrWein) Then
        GetEmployeeLastHireDate = mWorkforceHireDates(empIdOrWein)
        Exit Function
    End If
    
    ' Try Employee Code -> WEIN mapping
    empCode = EmpCodeFromWein(empIdOrWein)
    If empCode <> "" And mWorkforceHireDates.exists(empCode) Then
        GetEmployeeLastHireDate = mWorkforceHireDates(empCode)
        Exit Function
    End If
    
    ' Try WEIN -> Employee ID mapping
    empId = EmpIdFromWein(empIdOrWein)
    If empId <> "" And mWorkforceHireDates.exists(empId) Then
        GetEmployeeLastHireDate = mWorkforceHireDates(empId)
        Exit Function
    End If
End Function

'------------------------------------------------------------------------------
' Sub: OutputMaternityReport
' Purpose: Output Maternity Report for excluded records (< 40 weeks service)
' Description: Creates a new workbook with the same header structure as
'              Employee_Leave_Transactions_Report, containing excluded records
'------------------------------------------------------------------------------
Private Sub OutputMaternityReport()
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim rptWb As Workbook
    Dim rptWs As Worksheet
    Dim filePath As String
    Dim outputPath As String
    Dim lastCol As Long, c As Long
    Dim rec As Variant
    Dim v As Variant
    Dim rptRow As Long
    
    On Error GoTo ErrHandler
    
    If mMaternityExcludedRecords Is Nothing Then Exit Sub
    If mMaternityExcludedRecords.count = 0 Then Exit Sub
    
    ' Open source file to get header structure
    filePath = GetInputFilePath("EmployeeLeave")
    
    If Dir(filePath) = "" Then
        LogWarning "modSP1_Attendance", "OutputMaternityReport", _
            "Employee Leave file not found, cannot create Maternity Report"
        Exit Sub
    End If
    
    Set srcWb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = srcWb.Worksheets(1)
    
    ' Create new workbook for Maternity Report
    Set rptWb = Workbooks.Add
    Set rptWs = rptWb.Worksheets(1)
    rptWs.Name = "Maternity Report"
    
    ' Copy header row
    lastCol = srcWs.Cells(1, srcWs.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        rptWs.Cells(1, c).Value = srcWs.Cells(1, c).Value
    Next c
    
    ' Add additional columns for service info
    rptWs.Cells(1, lastCol + 1).Value = "Last Hire Date"
    rptWs.Cells(1, lastCol + 2).Value = "Weeks of Service"
    rptWs.Cells(1, lastCol + 3).Value = "Exclusion Reason"
    
    ' Build header index for source file
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    For c = 1 To lastCol
        Dim hdrName As String
        hdrName = UCase(Trim(CStr(Nz(srcWs.Cells(1, c).Value, ""))))
        If hdrName <> "" And Not headers.exists(hdrName) Then
            headers(hdrName) = c
        End If
    Next c
    
    srcWb.Close SaveChanges:=False
    Set srcWb = Nothing
    
    ' Write excluded records
    rptRow = 2
    For Each v In mMaternityExcludedRecords
        rec = v
        
        Dim recWein As String
        Dim lastHireDate As Date
        Dim leaveStartDate As Date
        Dim weeksOfService As Double
        
        recWein = CStr(rec(LR_WEIN))
        leaveStartDate = CDate(rec(LR_FROMDATE))
        lastHireDate = GetEmployeeLastHireDate(recWein)
        
        If lastHireDate > 0 Then
            weeksOfService = (leaveStartDate - lastHireDate) / 7
        Else
            weeksOfService = 0
        End If
        
        ' Write record data
        If headers.exists("WIN") Then rptWs.Cells(rptRow, headers("WIN")).Value = rec(LR_WEIN)
        If headers.exists("WEIN") Then rptWs.Cells(rptRow, headers("WEIN")).Value = rec(LR_WEIN)
        If headers.exists("EMPLOYEE CODE") Then rptWs.Cells(rptRow, headers("EMPLOYEE CODE")).Value = rec(LR_EMPCODE)
        If headers.exists("LEAVE TYPE") Then rptWs.Cells(rptRow, headers("LEAVE TYPE")).Value = rec(LR_LEAVETYPE)
        If headers.exists("FROM_DATE") Then rptWs.Cells(rptRow, headers("FROM_DATE")).Value = rec(LR_FROMDATE)
        If headers.exists("TO_DATE") Then rptWs.Cells(rptRow, headers("TO_DATE")).Value = rec(LR_TODATE)
        If headers.exists("APPLY_DATE") Then rptWs.Cells(rptRow, headers("APPLY_DATE")).Value = rec(LR_APPLYDATE)
        If headers.exists("APPROVAL_DATE") Then rptWs.Cells(rptRow, headers("APPROVAL_DATE")).Value = rec(LR_APPROVALDATE)
        If headers.exists("STATUS") Then rptWs.Cells(rptRow, headers("STATUS")).Value = rec(LR_STATUS)
        If headers.exists("TOTAL_DAYS") Then rptWs.Cells(rptRow, headers("TOTAL_DAYS")).Value = rec(LR_TOTALDAYS)
        
        ' Write additional service info
        rptWs.Cells(rptRow, lastCol + 1).Value = lastHireDate
        rptWs.Cells(rptRow, lastCol + 2).Value = Round(weeksOfService, 1)
        rptWs.Cells(rptRow, lastCol + 3).Value = "Less than 40 weeks service before maternity leave"
        
        rptRow = rptRow + 1
    Next v
    
    ' Format and save
    rptWs.Columns.AutoFit
    
    ' Build output path
    outputPath = G.RunParams.OutputFolder
    If Right(outputPath, 1) <> "\" Then outputPath = outputPath & "\"
    outputPath = outputPath & "Maternity Report_" & Format(G.Payroll.payrollMonth, "YYYYMM") & ".xlsx"
    
    ' Save workbook
    On Error Resume Next
    rptWb.SaveAs outputPath, xlOpenXMLWorkbook
    If Err.Number <> 0 Then
        LogError "modSP1_Attendance", "OutputMaternityReport", Err.Number, _
            "Failed to save Maternity Report: " & Err.Description
        Err.Clear
    Else
        LogInfo "modSP1_Attendance", "OutputMaternityReport", _
            "Maternity Report saved: " & outputPath & " (" & (rptRow - 2) & " records)"
    End If
    On Error GoTo ErrHandler
    
    rptWb.Close SaveChanges:=False
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Attendance", "OutputMaternityReport", Err.Number, Err.Description
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    If Not rptWb Is Nothing Then rptWb.Close SaveChanges:=False
End Sub
