Attribute VB_Name = "modSP2_CheckResult_Contribution"
'==============================================================================
' Module: modSP2_CheckResult_Contribution
' Purpose: Subprocess 2 - Contribution (MPF/ORSO) Check columns
' Description: Validates MPF and ORSO calculations
'==============================================================================
Option Explicit

' MPF/ORSO parameters from 额外表
Private mMPFParams As Object

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
    
    ' Process each WEIN
    Dim wein As Variant
    Dim row As Long
    
    For Each wein In weinIndex.Keys
        row = weinIndex(wein)
        
        ' Write Check values
        WriteMPFChecks ws, row, CStr(wein)
        WriteORSOChecks ws, row, CStr(wein)
    Next wein
    
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
    
    On Error GoTo ErrHandler
    
    Set mMPFParams = CreateObject("Scripting.Dictionary")
    
    Set wb = OpenExtraTableWorkbook()
    If wb Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set ws = wb.Worksheets("MPF&ORSO")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then Exit Sub
    
    Set headers = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        headers(UCase(Trim(CStr(ws.Cells(1, c).Value)))) = c
    Next c
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
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
    
    ' Get MPF Relevant Income from Check Result (already calculated)
    Dim colRI As Long
    colRI = FindColumnByHeader(ws.Rows(4), "MPF Relevant Income Check")
    If colRI > 0 Then
        mpfRI = ToDouble(ws.Cells(row, colRI).Value)
    End If
    
    ' Get MPF VC Relevant Income
    Dim colVCRI As Long
    colVCRI = FindColumnByHeader(ws.Rows(4), "MPF VC Relevant Income Check")
    If colVCRI > 0 Then
        mpfVCRI = ToDouble(ws.Cells(row, colVCRI).Value)
    End If
    
    ' Get percentages from params
    If mMPFParams.exists(wein) Then
        mpfEEVCPct = mMPFParams(wein)("MPF_EE_VC_Pct")
        mpfERVCPct = mMPFParams(wein)("MPF_ER_VC_Pct")
    End If
    
    ' MPF EE MC Check = MIN(MPF Relevant Income * 5%, 1500)
    col = FindColumnByHeader(ws.Rows(4), "MPF EE MC Check")
    If col > 0 Then
        mpfEEMC = WorksheetFunction.Min(mpfRI * 0.05, 1500)
        ws.Cells(row, col).Value = RoundAmount2(mpfEEMC)
    End If
    
    ' MPF ER MC Check = MIN(MPF Relevant Income * 5%, 1500)
    col = FindColumnByHeader(ws.Rows(4), "MPF ER MC Check")
    If col > 0 Then
        mpfERMC = WorksheetFunction.Min(mpfRI * 0.05, 1500)
        ws.Cells(row, col).Value = RoundAmount2(mpfERMC)
    End If
    
    ' MPF EE VC Check = MPF VC Relevant Income * MPF EE VC %
    col = FindColumnByHeader(ws.Rows(4), "MPF EE VC Check")
    If col > 0 Then
        mpfEEVC = mpfVCRI * mpfEEVCPct
        ws.Cells(row, col).Value = RoundAmount2(mpfEEVC)
    End If
    
    ' MPF ER VC Check = MAX(0, ROUND(MPF VC Relevant Income * MPF ER VC %, 2) - MPF ER MC)
    col = FindColumnByHeader(ws.Rows(4), "MPF ER VC Check")
    If col > 0 Then
        mpfERVC = RoundAmount2(mpfVCRI * mpfERVCPct) - mpfERMC
        If mpfERVC < 0 Then mpfERVC = 0
        ws.Cells(row, col).Value = RoundAmount2(mpfERVC)
    End If
End Sub

'------------------------------------------------------------------------------
' Sub: WriteORSOChecks
' Purpose: Write ORSO Check columns
'------------------------------------------------------------------------------
Private Sub WriteORSOChecks(ws As Worksheet, row As Long, wein As String)
    Dim col As Long
    Dim orsoRI As Double
    Dim orsoEE As Double, orsoER As Double
    Dim orsoERAdj As Double, orsoERPct As Double
    
    On Error Resume Next
    
    ' Get ORSO Relevant Income (Monthly Salary from Workforce Detail)
    Dim colRI As Long
    colRI = FindColumnByHeader(ws.Rows(4), "ORSO Relevant Income Check")
    If colRI > 0 Then
        orsoRI = ToDouble(ws.Cells(row, colRI).Value)
    End If
    
    ' Get parameters
    If mMPFParams.exists(wein) Then
        orsoERAdj = mMPFParams(wein)("ORSO_ER_Adj")
        orsoERPct = mMPFParams(wein)("ORSO_ER_Pct")
    End If
    
    ' ORSO EE Check = ORSO Relevant Income * 5%
    col = FindColumnByHeader(ws.Rows(4), "ORSO EE Check")
    If col > 0 Then
        orsoEE = orsoRI * 0.05
        ws.Cells(row, col).Value = RoundAmount2(orsoEE)
    End If
    
    ' ORSO ER Check = ORSO Relevant Income * Percent Of ORSO ER
    col = FindColumnByHeader(ws.Rows(4), "ORSO ER Check")
    If col > 0 Then
        orsoER = orsoRI * orsoERPct
        ws.Cells(row, col).Value = RoundAmount2(orsoER)
    End If
    
    ' ORSO ER Adj Check (from 额外表)
    col = FindColumnByHeader(ws.Rows(4), "ORSO ER Adj Check")
    If col > 0 Then
        ws.Cells(row, col).Value = RoundAmount2(orsoERAdj)
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
