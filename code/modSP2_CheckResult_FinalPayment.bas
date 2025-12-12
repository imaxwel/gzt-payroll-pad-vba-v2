Attribute VB_Name = "modSP2_CheckResult_FinalPayment"
'==============================================================================
' Module: modSP2_CheckResult_FinalPayment
' Purpose: Subprocess 2 - Final Payment Check columns
' Description: Validates Severance, Long Service, PIL, Gratuities, Back Pay
'==============================================================================
Option Explicit

' Final payment parameters from 额外表
Private mFinalPayParams As Object

' Workforce cache for year end bonus
Private mYearEndWorkforce As Object ' Dictionary: WEIN -> record(monthlySalary, lastHireDate)
' Termination WEIN set
Private mTerminationWeins As Object

'------------------------------------------------------------------------------
' Sub: SP2_Check_FinalPayment
' Purpose: Populate final payment Check columns
'------------------------------------------------------------------------------
Public Sub SP2_Check_FinalPayment(valWb As Workbook, weinIndex As Object)
    Dim ws As Worksheet
    
    On Error GoTo ErrHandler
    
    Set ws = valWb.Worksheets("Check Result")
    
    ' EAO data needed for Untaken Annual Leave Payment
    LoadEAOData
    
    ' Load Final Payment parameters
    LoadFinalPayParams
    
    ' Load Workforce Detail for Year End Bonus
    LoadYearEndWorkforceData
    
    ' Load termination WEINs for current month
    LoadTerminationWeins
    
    ' Process each WEIN
    Dim wein As Variant
    Dim row As Long
    
    For Each wein In weinIndex.Keys
        row = weinIndex(wein)
        
        ' Write Check values
        WriteSeveranceLongServiceCheck ws, row, CStr(wein)
        WritePILCheck ws, row, CStr(wein)
        WriteGratuitiesBackPayCheck ws, row, CStr(wein)
        WriteYearEndBonusCheck ws, row, CStr(wein)
        WriteUntakenALPaymentCheck ws, row, CStr(wein)
    Next wein
    
    LogInfo "modSP2_CheckResult_FinalPayment", "SP2_Check_FinalPayment", "Final payment checks completed"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_FinalPayment", "SP2_Check_FinalPayment", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: LoadFinalPayParams
' Purpose: Load Final Payment parameters from 额外表
'------------------------------------------------------------------------------
Private Sub LoadFinalPayParams()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim headers As Object
    Dim wein As String
    Dim rec As Object
    Dim headerRow As Long, keyCol As Long
    
    On Error GoTo ErrHandler
    
    Set mFinalPayParams = CreateObject("Scripting.Dictionary")
    
    Set wb = OpenExtraTableWorkbook()
    If wb Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set ws = wb.Worksheets("Final payment")
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
            rec("PolicyType") = GetCellVal(ws, i, headers, "MSD_or_Statutory")
            rec("TerminationType") = GetCellVal(ws, i, headers, "TerminationType")
            rec("PILIndicator") = GetCellVal(ws, i, headers, "PILIndicator")
            rec("NoticeGivenDate") = GetCellVal(ws, i, headers, "NoticeGivenDate")
            rec("NoticePeriod") = ToDouble(GetCellVal(ws, i, headers, "NoticePeriod"))
            rec("Gratuities") = ToDouble(GetCellVal(ws, i, headers, "Gratuities"))
            rec("BackPay") = ToDouble(GetCellVal(ws, i, headers, "BackPay"))
            
            If Not mFinalPayParams.exists(wein) Then
                mFinalPayParams.Add wein, rec
            End If
        End If
    Next i
    
    LogInfo "modSP2_CheckResult_FinalPayment", "LoadFinalPayParams", "Loaded " & mFinalPayParams.count & " final pay params"
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_FinalPayment", "LoadFinalPayParams", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: WriteSeveranceLongServiceCheck
' Purpose: Write Severance/Long Service Payment Check
'------------------------------------------------------------------------------
Private Sub WriteSeveranceLongServiceCheck(ws As Worksheet, row As Long, wein As String)
    Dim col As Long
    Dim monthlySalary As Double
    Dim lastHireDate As Date, termDate As Date
    Dim yos As Double
    Dim policyType As String, termType As String
    Dim payment As Double
    
    On Error Resume Next
    
    If Not mFinalPayParams.exists(wein) Then Exit Sub
    
    Dim params As Object
    Set params = mFinalPayParams(wein)
    
    policyType = UCase(params("PolicyType"))
    termType = UCase(params("TerminationType"))
    
    ' Get Monthly Salary (prefer workforce cache, fallback to Check Result)
    monthlySalary = GetWorkforceMonthlySalary(wein)
    If monthlySalary = 0 Then
        Dim colSalary As Long
        colSalary = FindColumnByHeader(ws.Rows(4), "Monthly Base Pay Check")
        If colSalary > 0 Then monthlySalary = ToDouble(ws.Cells(row, colSalary).value)
        If monthlySalary = 0 Then
            colSalary = FindColumnByHeader(ws.Rows(4), "Monthly Base Pay(Temp) Check")
            If colSalary > 0 Then monthlySalary = ToDouble(ws.Cells(row, colSalary).value)
        End If
    End If
    If monthlySalary = 0 Then Exit Sub
    
    ' Get dates from workforce and termination data
    lastHireDate = GetWorkforceHireDate(wein)
    termDate = GetTerminationDate(wein)
    If lastHireDate = 0 Or termDate = 0 Then Exit Sub
    
    yos = RoundAmount2((termDate - lastHireDate) / 365)
    
    ' Determine Severance vs Long Service
    If IsRedundancy(termType) Or yos < 5 Then
        payment = CalcSeverancePayment(policyType, monthlySalary, yos)
        col = GetCheckColIndex("Severance Payment 60404000")
        If col > 0 Then ws.Cells(row, col).value = payment
    Else
        payment = CalcSeverancePayment(policyType, monthlySalary, yos)
        col = GetCheckColIndex("Long Service Payment 60409960")
        If col > 0 Then ws.Cells(row, col).value = payment
    End If
End Sub

'------------------------------------------------------------------------------
' Sub: WritePILCheck
' Purpose: Write Payment in Lieu of Notice Check
'------------------------------------------------------------------------------
Private Sub WritePILCheck(ws As Worksheet, row As Long, wein As String)
    Dim col As Long
    Dim monthlyBase As Double, monthlyTemp As Double
    Dim transportAmt As Double
    Dim totalWage12 As Double
    Dim noticeDays As Double
    Dim noticeGiven As Date, termDate As Date, hireDate As Date
    Dim serviceDays As Double
    Dim pilEetoErDays As Double, pilErtoEeDays As Double
    Dim basePerDay As Double, eaoPerDay As Double
    Dim params As Object
    
    On Error Resume Next
    
    If Not mFinalPayParams.exists(wein) Then Exit Sub
    Set params = mFinalPayParams(wein)
    
    ' Base pay from Check Result
    monthlyBase = ToDouble(GetValueFromCheck(ws, row, "Monthly Base Pay"))
    monthlyTemp = ToDouble(GetValueFromCheck(ws, row, "Monthly Base Pay(Temp)"))
    
    ' Transport allowance
    transportAmt = GetTransportAllowanceAmount(wein)
    
    ' TotalWage_12Month from EAO
    totalWage12 = GetEAOTotalWage(wein)
    
    noticeDays = ToDouble(Nz(params("NoticePeriod"), 0))
    noticeGiven = ParseDateSafe(params("NoticeGivenDate"))
    termDate = GetTerminationDate(wein)
    hireDate = GetWorkforceHireDate(wein)
    
    If noticeDays <= 0 Or termDate = 0 Then Exit Sub
    
    pilEetoErDays = noticeDays - (termDate - noticeGiven)
    pilErtoEeDays = noticeDays - (termDate - noticeGiven)
    
    ' Service days up to end of previous month of notice given date if needed
    If noticeDays = 14 Or hireDate > 0 Then
        Dim serviceEnd As Date
        serviceEnd = DateSerial(Year(noticeGiven), Month(noticeGiven), 0)
        serviceDays = serviceEnd - hireDate + 1
        If serviceDays < 0 Then serviceDays = 0
    Else
        serviceDays = noticeDays * 30
    End If
    
    ' Wage per day calculations
    basePerDay = SafeDivide2(monthlyBase + monthlyTemp + transportAmt, noticeDays)
    If serviceDays > 0 Then
        eaoPerDay = SafeDivide2(totalWage12, serviceDays)
    Else
        eaoPerDay = SafeDivide2(totalWage12, noticeDays)
    End If
    
    ' PIL EE to ER: PIL Days * Min(basePerDay, eaoPerDay)
    col = GetCheckColIndex("PIL EE to ER 60001000")
    If col > 0 And pilEetoErDays > 0 Then
        ws.Cells(row, col).value = RoundAmount2(pilEetoErDays * WorksheetFunction.Min(basePerDay, eaoPerDay))
    End If
    
    ' PIL ER to EE: PIL Days * Max(basePerDay, eaoPerDay)
    col = GetCheckColIndex("PIL ER to EE 60001000")
    If col > 0 And pilErtoEeDays > 0 Then
        ws.Cells(row, col).value = RoundAmount2(pilErtoEeDays * WorksheetFunction.Max(basePerDay, eaoPerDay))
    End If
End Sub

'------------------------------------------------------------------------------
' Sub: WriteGratuitiesBackPayCheck
' Purpose: Write Gratuities and Back Pay Check
'------------------------------------------------------------------------------
Private Sub WriteGratuitiesBackPayCheck(ws As Worksheet, row As Long, wein As String)
    Dim col As Long
    
    On Error Resume Next
    
    If Not mFinalPayParams.exists(wein) Then Exit Sub
    
    Dim params As Object
    Set params = mFinalPayParams(wein)
    
    ' Gratuities Check
    col = GetCheckColIndex("Gratuities 99999999")
    If col > 0 Then
        ws.Cells(row, col).value = RoundAmount2(params("Gratuities"))
    End If
    
    ' Back Pay Check
    col = GetCheckColIndex("Back Pay 99999999")
    If col > 0 Then
        ws.Cells(row, col).value = RoundAmount2(params("BackPay"))
    End If
End Sub

'------------------------------------------------------------------------------
' Sub: WriteYearEndBonusCheck
' Purpose: Write Year End Bonus Check
'------------------------------------------------------------------------------
Private Sub WriteYearEndBonusCheck(ws As Worksheet, row As Long, wein As String)
    Dim col As Long
    Dim monthlySalary As Double
    Dim hireDate As Date
    Dim payment As Double
    
    On Error Resume Next
    
    ' Year End Bonus Check
    col = GetCheckColIndex("Year End Bonus 60208000")
    If col = 0 Then Exit Sub
    
    ' Get Monthly Salary from workforce cache; fallback to Check Result
    monthlySalary = GetWorkforceMonthlySalary(wein)
    If monthlySalary = 0 Then
        Dim colSalary As Long
        colSalary = GetCheckColIndex("Monthly Base Pay")
        If colSalary > 0 Then monthlySalary = ToDouble(ws.Cells(row, colSalary).value)
        If monthlySalary = 0 Then
            colSalary = GetCheckColIndex("Monthly Base Pay(Temp)")
            If colSalary > 0 Then monthlySalary = ToDouble(ws.Cells(row, colSalary).value)
        End If
    End If
    
    If monthlySalary = 0 Then Exit Sub
    
    ' If termination in current month, pay monthly salary
    If Not mTerminationWeins Is Nothing Then
        If mTerminationWeins.exists(wein) Then
            ws.Cells(row, col).value = RoundAmount2(monthlySalary)
            Exit Sub
        End If
    End If
    
    ' December rule based on service period within current year
    If Month(G.Payroll.monthEnd) = 12 Then
        hireDate = GetWorkforceHireDate(wein)
        payment = CalcYearEndBonus(monthlySalary, hireDate)
        If payment <> 0 Then ws.Cells(row, col).value = payment
    End If
End Sub

'------------------------------------------------------------------------------
' Function: CalcYearEndBonus
' Purpose: Calculate Year End Bonus based on service period in current year
' Formula:
'   If service period < full year: Monthly Salary / AnnualDays * ServiceDays
'   Else: Monthly Salary
'------------------------------------------------------------------------------
Private Function CalcYearEndBonus(monthlySalary As Double, hireDate As Date) As Double
    Dim yearStart As Date, yearEnd As Date
    Dim serviceStart As Date
    Dim serviceDays As Long
    Dim annualDays As Long
    
    yearEnd = G.Payroll.monthEnd
    yearStart = DateSerial(Year(yearEnd), 1, 1)
    serviceStart = yearStart
    If hireDate > 0 And hireDate > yearStart Then serviceStart = hireDate
    
    annualDays = DateDiff("d", yearStart, DateSerial(Year(yearEnd), 12, 31)) + 1
    serviceDays = DateDiff("d", serviceStart, yearEnd) + 1
    If serviceDays >= annualDays Then
        CalcYearEndBonus = RoundAmount2(monthlySalary)
    Else
        CalcYearEndBonus = RoundAmount2(SafeDivide2(monthlySalary, annualDays) * serviceDays)
    End If
End Function

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
' Data loaders for Year End Bonus
'------------------------------------------------------------------------------
Private Sub LoadYearEndWorkforceData()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filePath As String
    Dim headers As Object
    Dim headerRow As Long, lastRow As Long
    Dim wein As String
    Dim monthlySalary As Double
    Dim hireDate As Variant
    Dim i As Long
    
    On Error GoTo ErrHandler
    
    Set mYearEndWorkforce = CreateObject("Scripting.Dictionary")
    
    filePath = GetInputFilePathAuto("WorkforceDetail", poCurrentMonth)
    If Not FileExistsSafe(filePath) Then
        LogError "modSP2_CheckResult_FinalPayment", "LoadYearEndWorkforceData", 0, _
            "Workforce Detail file does not exist: " & filePath
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Worksheets(1)
    
    headerRow = FindHeaderRowSafe(ws, "EMPLOYEE ID,EMPLOYEEID", 1, 50)
    Set headers = BuildHeaderIndex(ws, headerRow)
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = headerRow + 1 To lastRow
        wein = NormalizeEmployeeId(Trim(CStr(Nz(GetCellVal(ws, i, headers, "WEIN"), ""))))
        If wein = "" Then wein = NormalizeEmployeeId(Trim(CStr(Nz(GetCellVal(ws, i, headers, "WIN"), ""))))
        If wein = "" Then wein = NormalizeEmployeeId(Trim(CStr(Nz(GetCellVal(ws, i, headers, "EMPLOYEE ID"), ""))))
        If wein <> "" Then
            monthlySalary = RoundMonthlySalary(GetCellVal(ws, i, headers, "MONTHLY SALARY"))
            hireDate = GetCellVal(ws, i, headers, "LAST HIRE DATE")
            Dim rec As Object
            Set rec = CreateObject("Scripting.Dictionary")
            rec("MonthlySalary") = monthlySalary
            If IsDate(hireDate) Then
                rec("HireDate") = CDate(hireDate)
            Else
                rec("HireDate") = 0
            End If
            If Not mYearEndWorkforce.exists(wein) Then
                mYearEndWorkforce.Add wein, rec
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_FinalPayment", "LoadYearEndWorkforceData", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

Private Sub LoadTerminationWeins()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filePath As String
    Dim headers As Object
    Dim headerRow As Long, lastRow As Long
    Dim empCode As String
    Dim wein As String
    Dim termDate As Variant
    Dim i As Long
    
    On Error GoTo ErrHandler
    
    Set mTerminationWeins = CreateObject("Scripting.Dictionary")
    
    filePath = GetInputFilePathAuto("Termination", poCurrentMonth)
    If Not FileExistsSafe(filePath) Then
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Worksheets(1)
    
    headerRow = FindHeaderRowSafe(ws, "EMPLOYEE CODE,EMPLOYEECODE,EMPLOYEE REFERENCE,EMPLOYEENUMBER,EMPLOYEE NUMBER", 1, 50)
    Set headers = BuildHeaderIndex(ws, headerRow)
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = headerRow + 1 To lastRow
        empCode = GetCellVal(ws, i, headers, "EMPLOYEE CODE")
        If empCode = "" Then empCode = GetCellVal(ws, i, headers, "EMPLOYEECODE")
        If empCode = "" Then empCode = GetCellVal(ws, i, headers, "EMPLOYEE REFERENCE")
        If empCode = "" Then empCode = GetCellVal(ws, i, headers, "EMPLOYEENUMBER")
        If empCode = "" Then empCode = GetCellVal(ws, i, headers, "EMPLOYEE NUMBER")
        termDate = GetCellVal(ws, i, headers, "TERMINATION DATE")
        If empCode <> "" Then
            wein = NormalizeEmployeeId(empCode)
            If wein <> "" Then
                If Not mTerminationWeins.exists(wein) Then mTerminationWeins.Add wein, True
                If Not mTerminationWeins.exists(wein & "|DATE") Then mTerminationWeins.Add wein & "|DATE", termDate
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_FinalPayment", "LoadTerminationWeins", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

Private Function GetWorkforceMonthlySalary(wein As String) As Double
    If Not mYearEndWorkforce Is Nothing Then
        If mYearEndWorkforce.exists(wein) Then
            GetWorkforceMonthlySalary = Nz(mYearEndWorkforce(wein)("MonthlySalary"), 0)
            Exit Function
        End If
    End If
    GetWorkforceMonthlySalary = 0
End Function

Private Function GetWorkforceHireDate(wein As String) As Date
    If Not mYearEndWorkforce Is Nothing Then
        If mYearEndWorkforce.exists(wein) Then
            If IsDate(Nz(mYearEndWorkforce(wein)("HireDate"), 0)) Then
                GetWorkforceHireDate = CDate(mYearEndWorkforce(wein)("HireDate"))
                Exit Function
            End If
        End If
    End If
    GetWorkforceHireDate = 0
End Function


'------------------------------------------------------------------------------
' Sub: WriteUntakenALPaymentCheck
' Purpose: Write Untaken Annual Leave Payment Check
' Formula: MAX(Monthly Salary / 22, AverageDayWage_12Month) * Untaken Annual Leave Days
'------------------------------------------------------------------------------
Private Sub WriteUntakenALPaymentCheck(ws As Worksheet, row As Long, wein As String)
    Dim col As Long
    Dim monthlySalary As Double
    Dim payment As Double
    Dim colSalary As Long
    
    On Error Resume Next
    
    col = GetCheckColIndex("Untaken Annual Leave Payment 60409960")
    If col = 0 Then Exit Sub
    
    ' Get Monthly Salary from Check Result (regular or temp)
    colSalary = GetCheckColIndex("Monthly Base Pay")
    If colSalary > 0 Then
        monthlySalary = ToDouble(ws.Cells(row, colSalary).value)
    End If
    If monthlySalary = 0 Then
        colSalary = GetCheckColIndex("Monthly Base Pay(Temp)")
        If colSalary > 0 Then monthlySalary = ToDouble(ws.Cells(row, colSalary).value)
    End If
    
    If monthlySalary = 0 Then Exit Sub
    
    ' Calculate payment from EAO summary
    payment = CalcUntakenAnnualLeavePayment(wein, monthlySalary)
    If payment <> 0 Then
        ws.Cells(row, col).value = payment
    End If
End Sub

'------------------------------------------------------------------------------
' Helper: IsRedundancy
'------------------------------------------------------------------------------
Private Function IsRedundancy(termType As String) As Boolean
    IsRedundancy = (InStr(UCase(termType), "REDUND") > 0)
End Function

'------------------------------------------------------------------------------
' Helper: CalcSeverancePayment
'------------------------------------------------------------------------------
Private Function CalcSeverancePayment(policyType As String, monthlySalary As Double, yos As Double) As Double
    Dim msdVal As Double, statVal As Double
    
    ' MSD policy: Base Pay * Min(24, YOS)
    msdVal = RoundAmount2(monthlySalary * WorksheetFunction.Min(24, yos))
    
    ' Statutory: Min(Min(Monthly Salary * 2/3, 15000) * YOS, 390000)
    statVal = WorksheetFunction.Min(WorksheetFunction.Min(monthlySalary * 2 / 3, 15000) * yos, 390000)
    statVal = RoundAmount2(statVal)
    
    If InStr(policyType, "MSD") > 0 Then
        CalcSeverancePayment = msdVal
    Else
        CalcSeverancePayment = statVal
    End If
End Function

'------------------------------------------------------------------------------
' Helper: GetTerminationDate
'------------------------------------------------------------------------------
Private Function GetTerminationDate(wein As String) As Date
    If Not mTerminationWeins Is Nothing Then
        If mTerminationWeins.exists(wein & "|DATE") Then
            If IsDate(mTerminationWeins(wein & "|DATE")) Then
                GetTerminationDate = CDate(mTerminationWeins(wein & "|DATE"))
                Exit Function
            End If
        End If
    End If
    GetTerminationDate = 0
End Function

'------------------------------------------------------------------------------
' Helper: GetValueFromCheck
'------------------------------------------------------------------------------
Private Function GetValueFromCheck(ws As Worksheet, row As Long, headerName As String) As Variant
    Dim col As Long
    col = FindColumnByHeader(ws.Rows(4), headerName & " Check")
    If col = 0 Then col = FindColumnByHeader(ws.Rows(4), headerName)
    If col > 0 Then
        GetValueFromCheck = ws.Cells(row, col).value
    Else
        GetValueFromCheck = 0
    End If
End Function

'------------------------------------------------------------------------------
' Helper: GetTransportAllowanceAmount
'------------------------------------------------------------------------------
Private Function GetTransportAllowanceAmount(wein As String) As Double
    Dim wb As Workbook, ws As Worksheet
    Dim filePath As String
    Dim headers As Object
    Dim headerRow As Long, lastRow As Long
    Dim empId As String
    Dim compPlan As String
    Dim amt As Double
    Dim i As Long
    
    On Error GoTo ErrHandler
    
    filePath = GetInputFilePathAuto("AllowancePlan", poCurrentMonth)
    If Not FileExistsSafe(filePath) Then
        GetTransportAllowanceAmount = 0
        Exit Function
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Worksheets(1)
    
    headerRow = FindHeaderRowSafe(ws, "EMPLOYEE ID,EMPLOYEEID,EMPLOYEE NUMBER ID,WEIN", 1, 50)
    Set headers = BuildHeaderIndex(ws, headerRow)
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = headerRow + 1 To lastRow
        compPlan = UCase(Trim(CStr(Nz(GetCellVal(ws, i, headers, "COMPENSATION PLAN"), ""))))
        If InStr(compPlan, "TRANSPORT") > 0 Then
            empId = NormalizeEmployeeId(Trim(CStr(Nz(GetCellVal(ws, i, headers, "EMPLOYEE ID"), ""))))
            If empId = "" Then empId = NormalizeEmployeeId(Trim(CStr(Nz(GetCellVal(ws, i, headers, "EMPLOYEEID"), ""))))
            If empId = "" Then empId = NormalizeEmployeeId(Trim(CStr(Nz(GetCellVal(ws, i, headers, "EMPLOYEE NUMBER ID"), ""))))
            If empId = "" Then empId = NormalizeEmployeeId(Trim(CStr(Nz(GetCellVal(ws, i, headers, "WEIN"), ""))))
            If empId = wein Then
                amt = ToDouble(GetCellVal(ws, i, headers, "AMOUNT"))
                If amt <> 0 Then
                    GetTransportAllowanceAmount = amt
                    Exit For
                End If
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    Exit Function
    
ErrHandler:
    GetTransportAllowanceAmount = 0
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Function

'------------------------------------------------------------------------------
' Helper: GetEAOTotalWage
'------------------------------------------------------------------------------
Private Function GetEAOTotalWage(wein As String) As Double
    ' EAO cache loaded earlier
    Dim rec As Variant
    rec = GetEAORecord(wein)
    If IsArray(rec) Then
        On Error Resume Next
        GetEAOTotalWage = ToDouble(rec(14))
        On Error GoTo 0
    Else
        GetEAOTotalWage = 0
    End If
End Function

'------------------------------------------------------------------------------
' Helper: ParseDateSafe
'------------------------------------------------------------------------------
Private Function ParseDateSafe(v As Variant) As Date
    If IsDate(v) Then
        ParseDateSafe = CDate(v)
    Else
        ParseDateSafe = 0
    End If
End Function
