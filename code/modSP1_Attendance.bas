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
            On Error Resume Next
            rec(LR_FROMDATE) = CDate(GetCellValue(ws, i, headers, "FROM_DATE"))
            rec(LR_TODATE) = CDate(GetCellValue(ws, i, headers, "TO_DATE"))
            rec(LR_APPLYDATE) = CDate(GetCellValue(ws, i, headers, "APPLY_DATE"))
            rec(LR_APPROVALDATE) = CDate(GetCellValue(ws, i, headers, "APPROVAL_DATE"))
            On Error GoTo ErrHandler
            
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
    prevYM = Format(G.Payroll.PrevMonthStart, "YYYYMM")
    
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
    prevYM = Format(G.Payroll.PrevMonthStart, "YYYYMM")
    
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
    prevYM = Format(G.Payroll.PrevMonthStart, "YYYYMM")
    
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
    prevYM = Format(G.Payroll.PrevMonthStart, "YYYYMM")
    
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
'------------------------------------------------------------------------------
Private Sub ProcessMaternityLeave(ws As Worksheet, leaveRecords As Collection, empIndex As Object)
    Dim rec As Variant
    Dim v As Variant
    Dim row As Long
    Dim empDays As Object
    Dim recWein As String, recUniqueKey As String
    
    On Error GoTo ErrHandler
    
    Set empDays = CreateObject("Scripting.Dictionary")
    
    For Each v In leaveRecords
        rec = v
        
        If UCase(CStr(rec(LR_LEAVETYPE))) Like "*MATERNITY*" Then
            ' TODO: Check 40 weeks service requirement
            
            recWein = CStr(rec(LR_WEIN))
            If Not empDays.exists(recWein) Then
                empDays.Add recWein, 0#
            End If
            empDays(recWein) = empDays(recWein) + CDbl(rec(LR_TOTALDAYS))
            
            recUniqueKey = CStr(rec(LR_UNIQUEKEY))
            mLeaveHistory.Add recUniqueKey, True
        End If
    Next v
    
    ' Write to sheet
    Dim wein As Variant
    Dim colDays As Long
    
    colDays = FindColumnByHeader(ws.Rows(1), "Days_MaternityLeave")
    
    For Each wein In empDays.Keys
        row = GetOrAddEmployeeRow(ws, CStr(wein), empIndex)
        If row > 0 And colDays > 0 Then
            ws.Cells(row, colDays).Value = RoundAmount2(empDays(wein))
        End If
    Next wein
    
    LogInfo "modSP1_Attendance", "ProcessMaternityLeave", "Processed " & empDays.count & " employees"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Attendance", "ProcessMaternityLeave", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: ProcessPaternityLeave
' Purpose: Process Paternity Leave records
'------------------------------------------------------------------------------
Private Sub ProcessPaternityLeave(ws As Worksheet, leaveRecords As Collection, empIndex As Object)
    Dim rec As Variant
    Dim v As Variant
    Dim row As Long
    Dim empDays As Object
    Dim recWein As String, recUniqueKey As String
    
    On Error GoTo ErrHandler
    
    Set empDays = CreateObject("Scripting.Dictionary")
    
    For Each v In leaveRecords
        rec = v
        
        If UCase(CStr(rec(LR_LEAVETYPE))) Like "*PATERNITY*" Then
            recWein = CStr(rec(LR_WEIN))
            If Not empDays.exists(recWein) Then
                empDays.Add recWein, 0#
            End If
            empDays(recWein) = empDays(recWein) + CDbl(rec(LR_TOTALDAYS))
            
            recUniqueKey = CStr(rec(LR_UNIQUEKEY))
            mLeaveHistory.Add recUniqueKey, True
        End If
    Next v
    
    ' Write to sheet
    Dim wein As Variant
    Dim colDays As Long
    
    colDays = FindColumnByHeader(ws.Rows(1), "Days_PaternityLeave")
    
    For Each wein In empDays.Keys
        row = GetOrAddEmployeeRow(ws, CStr(wein), empIndex)
        If row > 0 And colDays > 0 Then
            ws.Cells(row, colDays).Value = RoundAmount2(empDays(wein))
        End If
    Next wein
    
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
