Attribute VB_Name = "modSP2_CheckResult_FinalPayment"
'==============================================================================
' Module: modSP2_CheckResult_FinalPayment
' Purpose: Subprocess 2 - Final Payment Check columns
' Description: Validates Severance, Long Service, PIL, Gratuities, Back Pay
'==============================================================================
Option Explicit

' Final payment parameters from 额外表
Private mFinalPayParams As Object

'------------------------------------------------------------------------------
' Sub: SP2_Check_FinalPayment
' Purpose: Populate final payment Check columns
'------------------------------------------------------------------------------
Public Sub SP2_Check_FinalPayment(valWb As Workbook, weinIndex As Object)
    Dim ws As Worksheet
    
    On Error GoTo ErrHandler
    
    Set ws = valWb.Worksheets("Check Result")
    
    ' Load Final Payment parameters
    LoadFinalPayParams
    
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
    
    On Error GoTo ErrHandler
    
    Set mFinalPayParams = CreateObject("Scripting.Dictionary")
    
    Set wb = OpenExtraTableWorkbook()
    If wb Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set ws = wb.Worksheets("Final payment")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then Exit Sub
    
    Set headers = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        headers(UCase(Trim(CStr(ws.Cells(1, c).Value)))) = c
    Next c
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        wein = GetCellVal(ws, i, headers, "WEIN")
        If wein = "" Then wein = GetCellVal(ws, i, headers, "WIN")
        
        If wein <> "" Then
            Set rec = CreateObject("Scripting.Dictionary")
            rec("PolicyType") = GetCellVal(ws, i, headers, "MSD_or_Statutory")
            rec("TerminationType") = GetCellVal(ws, i, headers, "TerminationType")
            rec("PILIndicator") = GetCellVal(ws, i, headers, "PILIndicator")
            rec("NoticeGivenDate") = GetCellVal(ws, i, headers, "NoticeGivenDate")
            rec("NoticePeriod") = ToDouble(GetCellVal(ws, i, headers, "NoticePeriod"))
            rec("Gratuities") = ToDouble(GetCellVal(ws, i, headers, "Gratuities"))
            rec("BackPay") = ToDouble(GetCellVal(ws, i, headers, "BackPay"))
            
            If Not mFinalPayParams.Exists(wein) Then
                mFinalPayParams.Add wein, rec
            End If
        End If
    Next i
    
    LogInfo "modSP2_CheckResult_FinalPayment", "LoadFinalPayParams", "Loaded " & mFinalPayParams.Count & " final pay params"
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
    
    If Not mFinalPayParams.Exists(wein) Then Exit Sub
    
    Dim params As Object
    Set params = mFinalPayParams(wein)
    
    policyType = UCase(params("PolicyType"))
    termType = UCase(params("TerminationType"))
    
    ' Get Monthly Salary (from Check Result)
    Dim colSalary As Long
    colSalary = FindColumnByHeader(ws.Rows(4), "Monthly Base Pay Check")
    If colSalary > 0 Then
        monthlySalary = ToDouble(ws.Cells(row, colSalary).Value)
    End If
    
    ' Get dates (would need to be loaded from Workforce Detail and Termination)
    ' Placeholder: YOS calculation
    ' yos = (termDate - lastHireDate) / 365
    
    ' Determine Severance vs Long Service
    ' If redundancy or YOS < 5 -> Severance
    ' Else -> Long Service
    
    ' MSD Policy: Base Pay * Min(24, YOS)
    ' Statutory: Min(Min(Monthly Salary * 2/3, 15000) * YOS, 390000)
    
    ' Severance Payment Check
    col = FindColumnByHeader(ws.Rows(4), "Severance Payment Check")
    If col > 0 Then
        ' Placeholder calculation
        ' ws.Cells(row, col).Value = RoundAmount2(payment)
    End If
    
    ' Long Service Payment Check
    col = FindColumnByHeader(ws.Rows(4), "Long Service Payment Check")
    If col > 0 Then
        ' Placeholder calculation
        ' ws.Cells(row, col).Value = RoundAmount2(payment)
    End If
End Sub

'------------------------------------------------------------------------------
' Sub: WritePILCheck
' Purpose: Write Payment in Lieu of Notice Check
'------------------------------------------------------------------------------
Private Sub WritePILCheck(ws As Worksheet, row As Long, wein As String)
    Dim col As Long
    
    On Error Resume Next
    
    If Not mFinalPayParams.Exists(wein) Then Exit Sub
    
    Dim params As Object
    Set params = mFinalPayParams(wein)
    
    ' PIL EE to ER Check
    col = FindColumnByHeader(ws.Rows(4), "PIL EE to ER Check")
    If col > 0 Then
        ' Placeholder: Calculate based on notice period and dates
    End If
    
    ' PIL ER to EE Check
    col = FindColumnByHeader(ws.Rows(4), "PIL ER to EE Check")
    If col > 0 Then
        ' Placeholder: Calculate based on notice period and dates
    End If
End Sub

'------------------------------------------------------------------------------
' Sub: WriteGratuitiesBackPayCheck
' Purpose: Write Gratuities and Back Pay Check
'------------------------------------------------------------------------------
Private Sub WriteGratuitiesBackPayCheck(ws As Worksheet, row As Long, wein As String)
    Dim col As Long
    
    On Error Resume Next
    
    If Not mFinalPayParams.Exists(wein) Then Exit Sub
    
    Dim params As Object
    Set params = mFinalPayParams(wein)
    
    ' Gratuities Check
    col = FindColumnByHeader(ws.Rows(4), "Gratuities Check")
    If col > 0 Then
        ws.Cells(row, col).Value = RoundAmount2(params("Gratuities"))
    End If
    
    ' Back Pay Check
    col = FindColumnByHeader(ws.Rows(4), "Back Pay Check")
    If col > 0 Then
        ws.Cells(row, col).Value = RoundAmount2(params("BackPay"))
    End If
End Sub

'------------------------------------------------------------------------------
' Sub: WriteYearEndBonusCheck
' Purpose: Write Year End Bonus Check
'------------------------------------------------------------------------------
Private Sub WriteYearEndBonusCheck(ws As Worksheet, row As Long, wein As String)
    Dim col As Long
    Dim monthlySalary As Double
    
    On Error Resume Next
    
    ' Year End Bonus Check
    col = FindColumnByHeader(ws.Rows(4), "Year End Bonus Check")
    If col = 0 Then Exit Sub
    
    ' Get Monthly Salary
    Dim colSalary As Long
    colSalary = FindColumnByHeader(ws.Rows(4), "Monthly Base Pay Check")
    If colSalary > 0 Then
        monthlySalary = ToDouble(ws.Cells(row, colSalary).Value)
    End If
    
    ' For December or termination cases
    ' If service < 1 year: Monthly Salary / Annual Period * Service Period
    ' Else: Monthly Salary
    
    ' Placeholder implementation
    If Month(G.Payroll.MonthEnd) = 12 Then
        ws.Cells(row, col).Value = RoundAmount2(monthlySalary)
    End If
End Sub

'------------------------------------------------------------------------------
' Helper: GetCellVal
'------------------------------------------------------------------------------
Private Function GetCellVal(ws As Worksheet, row As Long, headers As Object, headerName As String) As String
    Dim col As Long
    GetCellVal = ""
    
    If headers.Exists(UCase(headerName)) Then
        col = headers(UCase(headerName))
        GetCellVal = Trim(CStr(Nz(ws.Cells(row, col).Value, "")))
    End If
End Function
