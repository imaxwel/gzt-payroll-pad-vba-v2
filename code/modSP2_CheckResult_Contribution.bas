Attribute VB_Name = "modSP2_CheckResult_Contribution"
'==============================================================================
' Module: modSP2_CheckResult_Contribution
' Purpose: Subprocess 2 - Contribution (MPF/ORSO) Check columns
' Description: Validates MPF and ORSO calculations
'==============================================================================
Option Explicit

' MPF/ORSO parameters from 额外表
Private mMPFParams As Object
' Goods & Services Differential from 额外表[特殊奖金]
Private mGoodsServicesDiff As Object
' Workforce Monthly Salary cache for ORSO Relevant Income
Private mOrsoWorkforce As Object

'------------------------------------------------------------------------------
' Sub: SP2_Check_Contribution
' Purpose: Populate contribution Check columns
'------------------------------------------------------------------------------
Public Sub SP2_Check_Contribution(valWb As Workbook, weinIndex As Object)
    Dim ws As Worksheet
    
    On Error GoTo ErrHandler
    
    Set ws = valWb.Worksheets("Check Result")
    
    ' Load MPF/ORSO parameters
    LoadMPFParams
    ' Load Goods & Services Differential from 额外表
    LoadGoodsServicesDiff
    ' Load workforce data for ORSO Relevant Income
    LoadWorkforceForOrso
    
    ' Process each WEIN
    Dim wein As Variant
    Dim row As Long
    
    For Each wein In weinIndex.Keys
        row = weinIndex(wein)
        
        ' Write Check values
        WriteMPFChecks ws, row, CStr(wein)
        WriteORSOChecks ws, row, CStr(wein)
    Next wein
    
    ' Write Optional Medical Check
    WriteOptionalMedicalCheck ws, weinIndex
    
    LogInfo "modSP2_CheckResult_Contribution", "SP2_Check_Contribution", "Contribution checks completed"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Contribution", "SP2_Check_Contribution", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: LoadMPFParams
' Purpose: Load MPF/ORSO parameters from 额外表
'------------------------------------------------------------------------------
Private Sub LoadMPFParams()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim headers As Object
    Dim wein As String
    Dim rec As Object
    Dim headerRow As Long, keyCol As Long
    
    On Error GoTo ErrHandler
    
    Set mMPFParams = CreateObject("Scripting.Dictionary")
    
    Set wb = OpenExtraTableWorkbook()
    If wb Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set ws = wb.Worksheets("MPF&ORSO")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then Exit Sub
    
    ' Detect header row and build header index
    headerRow = FindHeaderRowSafe(ws, "WEIN,WIN,WEINEmployee ID,EMPLOYEE CODEWIN,EMPLOYEE ID,EMPLOYEEID", 1, 50)
    Set headers = BuildHeaderIndex(ws, headerRow)
    
    keyCol = GetColumnFromHeaders(headers, "WEIN,WIN,WEINEmployee ID,EMPLOYEE CODEWIN,EMPLOYEE ID,EMPLOYEEID")
    If keyCol = 0 Then keyCol = 1
    lastRow = ws.Cells(ws.Rows.count, keyCol).End(xlUp).row
    
    For i = headerRow + 1 To lastRow
        ' Try multiple field name variants for WEIN
        wein = GetCellVal(ws, i, headers, "WEIN")
        If wein = "" Then wein = GetCellVal(ws, i, headers, "WIN")
        If wein = "" Then wein = GetCellVal(ws, i, headers, "WEINEMPLOYEE ID")
        If wein = "" Then wein = GetCellVal(ws, i, headers, "EMPLOYEE CODEWIN")
        If wein = "" Then wein = GetCellVal(ws, i, headers, "EMPLOYEE ID")
        If wein = "" Then wein = GetCellVal(ws, i, headers, "EMPLOYEEID")
        
        If wein <> "" Then
            Set rec = CreateObject("Scripting.Dictionary")
            rec("MPF_EE_VC_Pct") = ToDouble(GetCellVal(ws, i, headers, "MPF EE VC %"))
            rec("MPF_ER_VC_Pct") = ToDouble(GetCellVal(ws, i, headers, "MPF ER VC %"))
            rec("ORSO_Pct") = ToDouble(GetCellVal(ws, i, headers, "ORSO %"))
            rec("ORSO_ER_Adj") = ToDouble(GetCellVal(ws, i, headers, "ORSO ER Adj"))
            rec("ORSO_ER_Pct") = ToDouble(GetCellVal(ws, i, headers, "Percent Of ORSO ER"))
            rec("ORSO_EE_Pct") = ToDouble(GetCellVal(ws, i, headers, "Percent Of ORSO EE"))
            
            If Not mMPFParams.exists(wein) Then
                mMPFParams.Add wein, rec
            End If
        End If
    Next i
    
    LogInfo "modSP2_CheckResult_Contribution", "LoadMPFParams", "Loaded " & mMPFParams.count & " MPF params"
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Contribution", "LoadMPFParams", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: LoadGoodsServicesDiff
' Purpose: Load Goods & Services Differential from 额外表[特殊奖金]
'------------------------------------------------------------------------------
Private Sub LoadGoodsServicesDiff()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim headers As Object
    Dim headerRow As Long, lastRow As Long, i As Long
    Dim keyCol As Long
    Dim wein As String
    Dim amt As Double
    
    On Error GoTo ErrHandler
    
    Set mGoodsServicesDiff = CreateObject("Scripting.Dictionary")
    
    Set wb = OpenExtraTableWorkbook()
    If wb Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set ws = wb.Worksheets("特殊奖金")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then Exit Sub
    
    headerRow = FindHeaderRowSafe(ws, "WEIN,WIN,WEINEmployee ID,EMPLOYEE CODEWIN,EMPLOYEE ID,EMPLOYEEID", 1, 50)
    Set headers = BuildHeaderIndex(ws, headerRow)
    
    keyCol = GetColumnFromHeaders(headers, "WEIN,WIN,WEINEmployee ID,EMPLOYEE CODEWIN,EMPLOYEE ID,EMPLOYEEID")
    If keyCol = 0 Then keyCol = 1
    lastRow = ws.Cells(ws.Rows.count, keyCol).End(xlUp).row
    
    For i = headerRow + 1 To lastRow
        wein = GetCellVal(ws, i, headers, "WEIN")
        If wein = "" Then wein = GetCellVal(ws, i, headers, "WIN")
        If wein = "" Then wein = GetCellVal(ws, i, headers, "WEINEMPLOYEE ID")
        If wein = "" Then wein = GetCellVal(ws, i, headers, "EMPLOYEE CODEWIN")
        If wein = "" Then wein = GetCellVal(ws, i, headers, "EMPLOYEE ID")
        If wein = "" Then wein = GetCellVal(ws, i, headers, "EMPLOYEEID")
        
        If wein <> "" Then
            amt = ToDouble(GetCellVal(ws, i, headers, "Goods & Services Differential"))
            If Not mGoodsServicesDiff.exists(wein) Then
                mGoodsServicesDiff.Add wein, amt
            Else
                mGoodsServicesDiff(wein) = amt
            End If
        End If
    Next i
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Contribution", "LoadGoodsServicesDiff", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: LoadWorkforceForOrso
' Purpose: Load Monthly Salary (rounded) from Workforce Detail for ORSO Relevant Income
'------------------------------------------------------------------------------
Private Sub LoadWorkforceForOrso()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim headers As Object
    Dim headerRow As Long, lastRow As Long, i As Long
    Dim wein As String, empId As String
    Dim salary As Double
    
    On Error GoTo ErrHandler
    
    Set mOrsoWorkforce = CreateObject("Scripting.Dictionary")
    
    Dim filePath As String
    filePath = GetInputFilePathAuto("WorkforceDetail", poCurrentMonth)
    If Not FileExistsSafe(filePath) Then
        LogError "modSP2_CheckResult_Contribution", "LoadWorkforceForOrso", 0, _
            "Workforce Detail file does not exist: " & filePath
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Worksheets(1)
    
    headerRow = FindHeaderRowSafe(ws, "EMPLOYEE ID,EMPLOYEEID,EMPLOYEE NUMBER ID,WEIN,WIN", 1, 50)
    Set headers = BuildHeaderIndex(ws, headerRow)
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = headerRow + 1 To lastRow
        wein = GetCellVal(ws, i, headers, "WEIN")
        If wein = "" Then wein = GetCellVal(ws, i, headers, "WIN")
        If wein = "" Then wein = GetCellVal(ws, i, headers, "EMPLOYEE ID")
        If wein = "" Then wein = GetCellVal(ws, i, headers, "EMPLOYEEID")
        If wein = "" Then wein = GetCellVal(ws, i, headers, "EMPLOYEE NUMBER ID")
        
        empId = NormalizeEmployeeId(wein)
        If empId <> "" Then
            salary = RoundMonthlySalary(GetCellVal(ws, i, headers, "MONTHLY SALARY"))
            If Not mOrsoWorkforce.exists(empId) Then
                mOrsoWorkforce.Add empId, salary
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Contribution", "LoadWorkforceForOrso", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub
'------------------------------------------------------------------------------
' Sub: WriteMPFChecks
' Purpose: Write MPF Check columns
'------------------------------------------------------------------------------
Private Sub WriteMPFChecks(ws As Worksheet, row As Long, wein As String)
    Dim col As Long
    Dim mpfRI As Double, mpfVCRI As Double
    Dim mpfEEMC As Double, mpfERMC As Double
    Dim mpfEEVC As Double, mpfERVC As Double
    Dim mpfEEVCPct As Double, mpfERVCPct As Double
    
    On Error Resume Next
    
    ' Compute and write MPF Relevant Income (and VC) Check
    mpfRI = ComputeAndWriteMPFRelevantIncome(ws, row, wein)
    mpfVCRI = ComputeAndWriteMPFVCRelevantIncome(ws, row)
    
    ' Get percentages from params
    If mMPFParams.exists(wein) Then
        mpfEEVCPct = mMPFParams(wein)("MPF_EE_VC_Pct")
        mpfERVCPct = mMPFParams(wein)("MPF_ER_VC_Pct")
    End If
    
    ' Write MPF EE VC Percentage Check
    col = GetCheckColIndex("MPF EE VC Percentage")
    If col > 0 Then
        ws.Cells(row, col).value = mpfEEVCPct
    End If
    
    ' Write MPF ER VC Percentage Check
    col = GetCheckColIndex("MPF ER VC Percentage")
    If col > 0 Then
        ws.Cells(row, col).value = mpfERVCPct
    End If
    
    ' MPF EE MC Check = MIN(MPF Relevant Income * 5%, 1500)
    col = GetCheckColIndex("MPF EE MC 21251000")
    If col > 0 Then
        mpfEEMC = WorksheetFunction.Min(mpfRI * 0.05, 1500)
        ws.Cells(row, col).value = RoundAmount2(mpfEEMC)
    End If
    
    ' MPF ER MC Check = MIN(MPF Relevant Income * 5%, 1500)
    col = GetCheckColIndex("MPF ER MC 60801000")
    If col > 0 Then
        mpfERMC = WorksheetFunction.Min(mpfRI * 0.05, 1500)
        ws.Cells(row, col).value = RoundAmount2(mpfERMC)
    End If
    
    ' MPF EE VC Check = MPF VC Relevant Income * MPF EE VC %
    col = GetCheckColIndex("MPF EE VC 21251000")
    If col > 0 Then
        mpfEEVC = mpfVCRI * mpfEEVCPct
        ws.Cells(row, col).value = RoundAmount2(mpfEEVC)
    End If
    
    ' MPF ER VC Check = MAX(0, ROUND(MPF VC Relevant Income * MPF ER VC %, 2) - MPF ER MC)
    col = GetCheckColIndex("MPF ER VC 60801000")
    If col > 0 Then
        mpfERVC = RoundAmount2(mpfVCRI * mpfERVCPct) - mpfERMC
        If mpfERVC < 0 Then mpfERVC = 0
        ws.Cells(row, col).value = RoundAmount2(mpfERVC)
    End If
End Sub

'------------------------------------------------------------------------------
' Helper: Compute and write MPF Relevant Income Check (and VC Relevant Income)
'------------------------------------------------------------------------------
Private Function ComputeAndWriteMPFRelevantIncome(ws As Worksheet, row As Long, wein As String) As Double
    Const HEADER_ROW As Long = 4
    
    Dim total As Double
    Dim goodsVal As Double
    Dim col As Long
    
    ' Goods & Services Differential: prefer 额外表 value, fallback to Payroll Report
    goodsVal = GetGoodsServicesVal(wein)
    If goodsVal = 0 Then
        goodsVal = GetBenchmarkVal(ws, row, HEADER_ROW, "Goods & Services Differential,Goods & Services Differential 60601000")
    End If
    total = total + goodsVal
    
    ' Variance / incentive pay items
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Gratuity Bonus 60208000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Lump Sum Bonus 60409960")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Sign On Bonus 60409960")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Retention Bonus 60409960")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Referral Bonus 69001000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Annual Incentive 60201000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Sales Incentive (Quantitative)   21201000,Sales Incentive (Quantitative) 21201000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Sales Incentive (Qualitative) 21201000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "AP President Club 60208000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Manager of the Year Award 60208000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "MD Award 60208000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Employee Award 60208000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Inspire Points (Gross Up) 60701000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Inspire Cash 60702000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Year End Bonus 60208000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Shares Dividend 60204001")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Red Packet 69001000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Other Allowance 60409960")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Other Bonus 99999999")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Other Rewards 99999999")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Education Benefits 99999999")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Pensions 99999999")
    
    ' Attendance-related pay
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Maternity Leave Payment 60001000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Paternity Leave Payment 60001000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Sick Leave Payment 60001000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Salary Adj 60001000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Salary Adj (Temp) 60101000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Transport Allowance Adj 60409960")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Housing Allowance Adj 60409960")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Rental Reimbursement Adj 60001000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Total EAO Adj 60409960")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "No Pay Leave Deduction 60001000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Overtime Payment 99999999")
    
    ' Fixed pay
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Base Pay 60001000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Base Pay(Temp) 60101000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Housing Allowance 60409960")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Transport Allowance 60409960")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Rental Reimbursement 60001000")
    
    ' Termination-related pay
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Payroll_HK_UntakenAnnualLeavePayment,Untaken Annual Leave Payment 60409960")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Payroll_HK_OvertakenAnnualLeaveDeduct,Overtaken Annual Leave Deduct 60001000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Payroll_HK_BackPay,Back Pay 99999999")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Payroll_HK_Gratuities,Gratuities 99999999")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Payroll_HK_TerminationPayment,Termination Payment 99999999")
    
    ComputeAndWriteMPFRelevantIncome = RoundAmount2(total)
    
    ' Write Check columns
    col = GetCheckColIndex("MPF Relevant Income")
    If col > 0 Then ws.Cells(row, col).value = ComputeAndWriteMPFRelevantIncome
End Function

'------------------------------------------------------------------------------
' Helper: Compute and write MPF VC Relevant Income Check
'------------------------------------------------------------------------------
Private Function ComputeAndWriteMPFVCRelevantIncome(ws As Worksheet, row As Long) As Double
    Const HEADER_ROW As Long = 4
    
    Dim total As Double
    Dim col As Long
    
    ' Positive components
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Base Pay 60001000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Base Pay(Temp) 60101000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Rental Reimbursement 60001000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Rental Reimbursement Adj 60001000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Total EAO Adj 60409960")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Salary Adj 60001000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Salary Adj (Temp) 60101000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Maternity Leave Payment 60001000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Paternity Leave Payment 60001000")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Paid Parental Time Off (PPTO) payment")
    total = total + GetBenchmarkVal(ws, row, HEADER_ROW, "Sick Leave Payment 60001000")
    
    ' Negative component
    total = total - GetBenchmarkVal(ws, row, HEADER_ROW, "No Pay Leave Deduction 60001000")
    
    ComputeAndWriteMPFVCRelevantIncome = RoundAmount2(total)
    
    col = GetCheckColIndex("MPF VC Relevant Income")
    If col > 0 Then ws.Cells(row, col).value = ComputeAndWriteMPFVCRelevantIncome
End Function

'------------------------------------------------------------------------------
' Helper: Get benchmark value from Payroll Report row by header (comma variants)
'------------------------------------------------------------------------------
Private Function GetBenchmarkVal(ws As Worksheet, row As Long, headerRow As Long, headerVariants As String) As Double
    Dim col As Long
    col = FindColumnByHeader(ws.Rows(headerRow), headerVariants)
    If col > 0 Then
        GetBenchmarkVal = ToDouble(ws.Cells(row, col).value)
    Else
        GetBenchmarkVal = 0
    End If
End Function

'------------------------------------------------------------------------------
' Helper: Get Goods & Services Differential value for a WEIN from 额外表
'------------------------------------------------------------------------------
Private Function GetGoodsServicesVal(wein As String) As Double
    On Error Resume Next
    If Not mGoodsServicesDiff Is Nothing Then
        If mGoodsServicesDiff.exists(wein) Then
            GetGoodsServicesVal = mGoodsServicesDiff(wein)
        Else
            GetGoodsServicesVal = 0
        End If
    End If
End Function

'------------------------------------------------------------------------------
' Helper: Compute and write ORSO Relevant Income Check (Monthly Salary)
'------------------------------------------------------------------------------
Private Function ComputeAndWriteORSORelevantIncome(ws As Worksheet, row As Long, wein As String) As Double
    Dim base As Double
    Dim col As Long
    
    base = GetOrsoMonthlySalary(wein)
    ComputeAndWriteORSORelevantIncome = base
    
    col = GetCheckColIndex("ORSO Relevant Income")
    If col > 0 Then ws.Cells(row, col).value = base
End Function

'------------------------------------------------------------------------------
' Helper: Get Monthly Salary for ORSO base by WEIN
'------------------------------------------------------------------------------
Private Function GetOrsoMonthlySalary(wein As String) As Double
    On Error Resume Next
    If Not mOrsoWorkforce Is Nothing Then
        If mOrsoWorkforce.exists(wein) Then
            GetOrsoMonthlySalary = mOrsoWorkforce(wein)
            Exit Function
        End If
    End If
    GetOrsoMonthlySalary = 0
End Function

'------------------------------------------------------------------------------
' Sub: WriteORSOChecks
' Purpose: Write ORSO Check columns
'------------------------------------------------------------------------------
Private Sub WriteORSOChecks(ws As Worksheet, row As Long, wein As String)
    Dim col As Long
    Dim orsoRI As Double
    Dim orsoEE As Double, orsoER As Double
    Dim orsoERAdj As Double, orsoERPct As Double, orsoEEPct As Double
    
    On Error Resume Next
    
    ' Compute and write ORSO Relevant Income (Monthly Salary from Workforce Detail)
    orsoRI = ComputeAndWriteORSORelevantIncome(ws, row, wein)
    
    ' Get parameters
    If mMPFParams.exists(wein) Then
        orsoERAdj = mMPFParams(wein)("ORSO_ER_Adj")
        orsoERPct = mMPFParams(wein)("ORSO_ER_Pct")
        orsoEEPct = mMPFParams(wein)("ORSO_EE_Pct")
    End If
    
    ' Percent Of ORSO EE Check
    col = GetCheckColIndex("Percent Of ORSO EE")
    If col > 0 Then
        ws.Cells(row, col).value = orsoEEPct
    End If
    
    ' Percent Of ORSO ER Check
    col = GetCheckColIndex("Percent Of ORSO ER")
    If col > 0 Then
        ws.Cells(row, col).value = orsoERPct
    End If
    
    ' ORSO EE Check = ORSO Relevant Income * 5%
    col = GetCheckColIndex("ORSO EE 60801000")
    If col > 0 Then
        orsoEE = orsoRI * 0.05
        ws.Cells(row, col).value = RoundAmount2(orsoEE)
    End If
    
    ' ORSO ER Check = ORSO Relevant Income * Percent Of ORSO ER
    col = GetCheckColIndex("ORSO ER 60801000")
    If col > 0 Then
        orsoER = orsoRI * orsoERPct
        ws.Cells(row, col).value = RoundAmount2(orsoER)
    End If
    
    ' ORSO ER Adj Check (from 额外表)
    col = GetCheckColIndex("ORSO ER Adj")
    If col > 0 Then
        ws.Cells(row, col).value = RoundAmount2(orsoERAdj)
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
        GetCellVal = Trim(CStr(Nz(ws.Cells(row, col).value, "")))
    End If
End Function


'------------------------------------------------------------------------------
' Sub: WriteOptionalMedicalCheck
' Purpose: Write Optional Group Upgrade Check column from Optional medical plan
'------------------------------------------------------------------------------
Private Sub WriteOptionalMedicalCheck(ws As Worksheet, weinIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim lastRow As Long, i As Long
    Dim headers As Object
    Dim empId As String, wein As String
    Dim row As Long, col As Long
    Dim amount As Double
    Dim headerRow As Long, keyCol As Long
    
    On Error GoTo ErrHandler
    
    col = GetCheckColIndex("Optional Group Upgrade 21351000")
    If col = 0 Then Exit Sub
    
    ' Use new path service
    filePath = GetInputFilePathAuto("OptionalMedical", poCurrentMonth)
    If Not FileExistsSafe(filePath) Then
        LogWarning "modSP2_CheckResult_Contribution", "WriteOptionalMedicalCheck", _
            "Optional medical plan file does not exist (optional): " & filePath
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    ' Detect header row and build header index
    headerRow = FindHeaderRowSafe(srcWs, "EMPLOYEE ID,EMPLOYEEID,WEIN", 1, 50)
    Set headers = BuildHeaderIndex(srcWs, headerRow)
    
    keyCol = GetColumnFromHeaders(headers, "EMPLOYEE ID,EMPLOYEEID,WEIN")
    If keyCol = 0 Then keyCol = 1
    lastRow = srcWs.Cells(srcWs.Rows.count, keyCol).End(xlUp).row
    
    For i = headerRow + 1 To lastRow
        ' Get Employee ID
        empId = GetContribCellVal(srcWs, i, headers, "EMPLOYEE ID")
        If empId = "" Then empId = GetContribCellVal(srcWs, i, headers, "EMPLOYEEID")
        If empId = "" Then empId = GetContribCellVal(srcWs, i, headers, "WEIN")
        
        If empId <> "" Then
            wein = NormalizeEmployeeId(empId)
            
            If weinIndex.exists(wein) Then
                row = weinIndex(wein)
                amount = ToDouble(GetContribCellVal(srcWs, i, headers, "AMOUNT"))
                If amount > 0 Then
                    ws.Cells(row, col).value = SafeAdd2(ws.Cells(row, col).value, amount)
                End If
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Contribution", "WriteOptionalMedicalCheck", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Helper: GetContribCellVal
'------------------------------------------------------------------------------
Private Function GetContribCellVal(ws As Worksheet, row As Long, headers As Object, headerName As String) As String
    Dim col As Long
    GetContribCellVal = ""
    
    If headers.exists(UCase(headerName)) Then
        col = headers(UCase(headerName))
        GetContribCellVal = Trim(CStr(Nz(ws.Cells(row, col).value, "")))
    End If
End Function
