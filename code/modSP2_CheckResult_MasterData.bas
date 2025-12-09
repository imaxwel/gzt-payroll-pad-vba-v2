Attribute VB_Name = "modSP2_CheckResult_MasterData"
'==============================================================================
' Module: modSP2_CheckResult_MasterData
' Purpose: Subprocess 2 - Master Data Check columns
' Description: Validates Legal Name, Hire Date, Org Info, Base Pay, Transport
'==============================================================================
Option Explicit

' Workforce Detail cache
Private mWorkforceData As Object ' Dictionary of Employee ID -> record

'------------------------------------------------------------------------------
' Sub: SP2_Check_MasterData
' Purpose: Populate master data Check columns
'------------------------------------------------------------------------------
Public Sub SP2_Check_MasterData(valWb As Workbook, weinIndex As Object)
    Dim ws As Worksheet
    
    On Error GoTo ErrHandler
    
    Set ws = valWb.Worksheets("Check Result")
    
    ' Load Workforce Detail data
    LoadWorkforceData
    
    ' Load Allowance Plan data
    Dim allowanceData As Object
    Set allowanceData = LoadAllowanceData()
    
    ' Load Termination data
    Dim termData As Object
    Set termData = LoadTerminationData()
    
    ' Process each WEIN
    Dim wein As Variant
    Dim row As Long
    Dim empId As String
    Dim weinStr As String
    
    For Each wein In weinIndex.Keys
        row = weinIndex(wein)
        weinStr = CStr(wein)
        empId = EmpIdFromWein(weinStr)
        
        ' If empId mapping not found, use WEIN directly as Employee ID
        ' (In Workforce Detail, Employee ID is often the same as WEIN)
        If empId = "" Then empId = weinStr
        
        ' Write Check values
        WriteNameCheck ws, row, empId, weinStr
        WriteDateChecks ws, row, empId, termData
        WriteOrgChecks ws, row, empId
        WritePayChecks ws, row, empId, allowanceData
    Next wein
    
    LogInfo "modSP2_CheckResult_MasterData", "SP2_Check_MasterData", "Master data checks completed"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_MasterData", "SP2_Check_MasterData", Err.Number, Err.Description
End Sub


'------------------------------------------------------------------------------
' Sub: LoadWorkforceData
' Purpose: Load Workforce Detail into memory
'------------------------------------------------------------------------------
Private Sub LoadWorkforceData()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filePath As String
    Dim lastRow As Long, i As Long
    Dim headers As Object
    Dim empId As String
    Dim rec As Object
    
    On Error GoTo ErrHandler
    
    Set mWorkforceData = CreateObject("Scripting.Dictionary")
    
    ' Use new path service
    filePath = GetInputFilePathAuto("WorkforceDetail", poCurrentMonth)
    If Not FileExistsSafe(filePath) Then
        LogError "modSP2_CheckResult_MasterData", "LoadWorkforceData", 0, _
            "Workforce Detail file does not exist: " & filePath
        Exit Sub
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
    
    For i = 2 To lastRow
        ' Try multiple field name variants for Employee ID
        empId = GetCellVal(ws, i, headers, "EMPLOYEE ID")
        If empId = "" Then empId = GetCellVal(ws, i, headers, "EMPLOYEEID")
        If empId = "" Then empId = GetCellVal(ws, i, headers, "EMPLOYEE NUMBER ID")
        
        If empId <> "" Then
            Set rec = CreateObject("Scripting.Dictionary")
            rec("EmployeeID") = empId
            ' Try multiple field name variants for WEIN
            rec("WEIN") = GetCellVal(ws, i, headers, "WEIN")
            If rec("WEIN") = "" Then rec("WEIN") = GetCellVal(ws, i, headers, "WIN")
            If rec("WEIN") = "" Then rec("WEIN") = GetCellVal(ws, i, headers, "WEINEMPLOYEE ID")
            ' Directly extract Legal Full Name from Workforce Detail (not concatenated)
            rec("LegalFullName") = GetCellVal(ws, i, headers, "LEGAL FULL NAME")
            rec("LastHireDate") = GetCellVal(ws, i, headers, "LAST HIRE DATE")
            rec("BusinessDepartment") = GetCellVal(ws, i, headers, "BUSINESS DEPARTMENT")
            rec("PositionTitle") = GetCellVal(ws, i, headers, "POSITION TITLE")
            rec("CostCenterID") = GetCellVal(ws, i, headers, "COST CENTER - ID")
            rec("MonthlySalary") = RoundMonthlySalary(GetCellVal(ws, i, headers, "MONTHLY SALARY"))
            rec("EmployeeType") = GetCellVal(ws, i, headers, "EMPLOYEE TYPE")
            
            If Not mWorkforceData.exists(empId) Then
                mWorkforceData.Add empId, rec
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    
    LogInfo "modSP2_CheckResult_MasterData", "LoadWorkforceData", "Loaded " & mWorkforceData.count & " records"
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_MasterData", "LoadWorkforceData", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Function: LoadAllowanceData
' Purpose: Load Allowance Plan data
'------------------------------------------------------------------------------
Private Function LoadAllowanceData() As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filePath As String
    Dim lastRow As Long, i As Long
    Dim headers As Object
    Dim empId As String, compPlan As String
    Dim amt As Double
    Dim dict As Object
    
    On Error GoTo ErrHandler
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Use new path service
    filePath = GetInputFilePathAuto("AllowancePlan", poCurrentMonth)
    If Not FileExistsSafe(filePath) Then
        LogInfo "modSP2_CheckResult_MasterData", "LoadAllowanceData", _
            "Allowance Plan file does not exist (optional): " & filePath
        Set LoadAllowanceData = dict
        Exit Function
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Worksheets(1)
    
    Set headers = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        headers(UCase(Trim(CStr(ws.Cells(1, c).Value)))) = c
    Next c
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        compPlan = UCase(GetCellVal(ws, i, headers, "COMPENSATION PLAN"))
        
        If InStr(compPlan, "TRANSPORTATION") > 0 Then
            ' Try multiple field name variants for Employee ID
            empId = GetCellVal(ws, i, headers, "EMPLOYEE ID")
            If empId = "" Then empId = GetCellVal(ws, i, headers, "EMPLOYEEID")
            If empId = "" Then empId = GetCellVal(ws, i, headers, "EMPLOYEE NUMBER ID")
            amt = ToDouble(GetCellVal(ws, i, headers, "AMOUNT"))
            
            If empId <> "" Then
                If dict.exists(empId) Then
                    dict(empId) = dict(empId) + amt
                Else
                    dict.Add empId, amt
                End If
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    
    Set LoadAllowanceData = dict
    Exit Function
    
ErrHandler:
    LogError "modSP2_CheckResult_MasterData", "LoadAllowanceData", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set LoadAllowanceData = CreateObject("Scripting.Dictionary")
End Function

'------------------------------------------------------------------------------
' Function: LoadTerminationData
' Purpose: Load Termination data
'------------------------------------------------------------------------------
Private Function LoadTerminationData() As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filePath As String
    Dim lastRow As Long, i As Long
    Dim headers As Object
    Dim empCode As String, wein As String
    Dim termDate As String
    Dim dict As Object
    
    On Error GoTo ErrHandler
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Use new path service
    filePath = GetInputFilePathAuto("Termination", poCurrentMonth)
    If Not FileExistsSafe(filePath) Then
        LogError "modSP2_CheckResult_MasterData", "LoadTerminationData", 0, _
            "Termination file does not exist: " & filePath
        Set LoadTerminationData = dict
        Exit Function
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Worksheets(1)
    
    Set headers = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        headers(UCase(Trim(CStr(ws.Cells(1, c).Value)))) = c
    Next c
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        ' Try multiple field name variants for Employee Code
        empCode = GetCellVal(ws, i, headers, "EMPLOYEE CODE")
        If empCode = "" Then empCode = GetCellVal(ws, i, headers, "EMPLOYEECODE")
        If empCode = "" Then empCode = GetCellVal(ws, i, headers, "EMPLOYEE REFERENCE")
        If empCode = "" Then empCode = GetCellVal(ws, i, headers, "EMPLOYEENUMBER")
        termDate = GetCellVal(ws, i, headers, "TERMINATION DATE")
        
        If empCode <> "" Then
            wein = NormalizeEmployeeId(empCode)
            
            If Not dict.exists(wein) Then
                dict.Add wein, termDate
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    
    Set LoadTerminationData = dict
    Exit Function
    
ErrHandler:
    LogError "modSP2_CheckResult_MasterData", "LoadTerminationData", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set LoadTerminationData = CreateObject("Scripting.Dictionary")
End Function

'------------------------------------------------------------------------------
' Sub: WriteNameCheck
' Purpose: Write Legal Full Name and Check columns
' Logic:
'   - "Legal Full Name" column: Concatenate Legal First Name & " " & Legal Last Name
'     from Check Result sheet itself (HK Payroll Validation Output)
'   - "Legal Full Name Check" column: Directly use Legal Full Name from
'     Workforce Detail - Payroll-AP (mapped by Employee ID to WEIN)
'   - Diff column is computed later by SP2_ComputeDiff
'------------------------------------------------------------------------------
Private Sub WriteNameCheck(ws As Worksheet, row As Long, empId As String, wein As String)
    Dim colFirstName As Long
    Dim colLastName As Long
    Dim colFullName As Long
    Dim colCheck As Long
    Dim firstName As String
    Dim lastName As String
    Dim fullName As String
    Dim checkValue As String
    Dim lookupKey As String
    
    On Error Resume Next
    
    ' Find Legal First Name and Legal Last Name columns from Payroll Report
    colFirstName = FindColumnByHeader(ws.Rows(4), "Legal First Name")
    colLastName = FindColumnByHeader(ws.Rows(4), "Legal Last Name")
    
    ' Find the inserted Legal Full Name and Check columns
    ' These are inserted right after Legal First Name column
    colFullName = FindColumnByHeader(ws.Rows(4), "Legal Full Name")
    colCheck = FindColumnByHeader(ws.Rows(4), "Legal Full Name Check")
    
    ' Step 1: Populate "Legal Full Name" column by concatenating
    ' Legal First Name & " " & Legal Last Name from Check Result sheet
    If colFullName > 0 And colFirstName > 0 And colLastName > 0 Then
        firstName = Trim(CStr(Nz(ws.Cells(row, colFirstName).Value, "")))
        lastName = Trim(CStr(Nz(ws.Cells(row, colLastName).Value, "")))
        fullName = Trim(firstName & " " & lastName)
        ws.Cells(row, colFullName).Value = fullName
    End If
    
    ' Step 2: Populate "Legal Full Name Check" column with Legal Full Name
    ' directly from Workforce Detail - Payroll-AP
    ' Note: In Workforce Detail, Employee ID is the same as WEIN in Payroll Report
    ' Try empId first, then fall back to wein as the lookup key
    If colCheck > 0 Then
        Dim rec As Object
        lookupKey = ""
        
        ' Try empId first
        If empId <> "" And mWorkforceData.exists(empId) Then
            lookupKey = empId
        ' Fall back to wein (WEIN in Payroll Report = Employee ID in Workforce Detail)
        ElseIf wein <> "" And mWorkforceData.exists(wein) Then
            lookupKey = wein
        End If
        
        If lookupKey <> "" Then
            Set rec = mWorkforceData(lookupKey)
            checkValue = Trim(CStr(Nz(rec("LegalFullName"), "")))
            ws.Cells(row, colCheck).Value = checkValue
        End If
    End If
    
    ' Note: Diff column will be computed by SP2_ComputeDiff module
End Sub

'------------------------------------------------------------------------------
' Sub: WriteDateChecks
' Purpose: Write date-related Check columns
'------------------------------------------------------------------------------
Private Sub WriteDateChecks(ws As Worksheet, row As Long, empId As String, termData As Object)
    Dim col As Long
    Dim wein As String
    
    On Error Resume Next
    
    ' Last Hired Date Check
    col = GetCheckColIndex("Last Hired Date")
    If col > 0 And mWorkforceData.exists(empId) Then
        ws.Cells(row, col).Value = mWorkforceData(empId)("LastHireDate")
    End If
    
    ' Last Employment Date Check (Termination Date)
    col = GetCheckColIndex("Last Employment Date")
    If col > 0 Then
        wein = NormalizeEmployeeId(empId)
        If wein <> "" And termData.exists(wein) Then
            ws.Cells(row, col).Value = termData(wein)
        End If
    End If
End Sub

'------------------------------------------------------------------------------
' Sub: WriteOrgChecks
' Purpose: Write organization-related Check columns
'------------------------------------------------------------------------------
Private Sub WriteOrgChecks(ws As Worksheet, row As Long, empId As String)
    Dim col As Long
    
    On Error Resume Next
    
    If Not mWorkforceData.exists(empId) Then Exit Sub
    
    Dim rec As Object
    Set rec = mWorkforceData(empId)
    
    ' Business Department Check
    col = GetCheckColIndex("Business Department")
    If col > 0 Then ws.Cells(row, col).Value = rec("BusinessDepartment")
    
    ' Position Title Check
    col = GetCheckColIndex("Position Title")
    If col > 0 Then ws.Cells(row, col).Value = rec("PositionTitle")
    
    ' Cost Center Code Check
    col = GetCheckColIndex("Cost Center Code")
    If col > 0 Then ws.Cells(row, col).Value = rec("CostCenterID")
End Sub

'------------------------------------------------------------------------------
' Sub: WritePayChecks
' Purpose: Write pay-related Check columns
'------------------------------------------------------------------------------
Private Sub WritePayChecks(ws As Worksheet, row As Long, empId As String, allowanceData As Object)
    Dim col As Long
    Dim monthlySalary As Double
    Dim empType As String
    
    On Error Resume Next
    
    If Not mWorkforceData.exists(empId) Then Exit Sub
    
    Dim rec As Object
    Set rec = mWorkforceData(empId)
    
    monthlySalary = rec("MonthlySalary")
    empType = UCase(rec("EmployeeType"))
    
    ' Monthly Base Pay Check (Regular employees)
    col = GetCheckColIndex("Monthly Base Pay")
    If col > 0 Then
        If InStr(empType, "REGULAR") > 0 Then
            ws.Cells(row, col).Value = monthlySalary
        End If
    End If
    
    ' Monthly Base Pay (Temp) Check (Intern/Co-ops)
    col = GetCheckColIndex("Monthly Base Pay(Temp)")
    If col > 0 Then
        If InStr(empType, "INTERN") > 0 Or InStr(empType, "CO-OP") > 0 Then
            ws.Cells(row, col).Value = monthlySalary
        End If
    End If
    
    ' Monthly Transport Allowance Check
    col = GetCheckColIndex("Monthly Transport Allowance")
    If col > 0 Then
        If allowanceData.exists(empId) Then
            ws.Cells(row, col).Value = RoundAmount2(allowanceData(empId))
        End If
    End If
End Sub

'------------------------------------------------------------------------------
' Helper: GetCellVal
'------------------------------------------------------------------------------
Private Function GetCellVal(ws As Worksheet, row As Long, headers As Object, headerName As String) As String
    Dim col As Long
    GetCellVal = ""
    
    If headers.exists(UCase(headerName)) Then
        col = headers(UCase(headerName))
        GetCellVal = Trim(CStr(Nz(ws.Cells(row, col).Value, "")))
    End If
End Function
