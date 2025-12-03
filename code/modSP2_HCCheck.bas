Attribute VB_Name = "modSP2_HCCheck"
'==============================================================================
' Module: modSP2_HCCheck
' Purpose: Subprocess 2 - HC Check sheet
' Description: Headcount validation and cross-check
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Sub: SP2_BuildHCCheck
' Purpose: Build the HC Check sheet
'------------------------------------------------------------------------------
Public Sub SP2_BuildHCCheck(valWb As Workbook)
    Dim ws As Worksheet
    
    On Error GoTo ErrHandler
    
    Set ws = valWb.Worksheets("HC Check")
    
    ' Build header
    BuildHCCheckHeader ws
    
    ' Calculate headcounts
    CalculatePayrollHC ws
    CalculateTerminatedHC ws
    CalculateNewHireHC ws
    CalculateHCFormula ws
    
    ' Create pivot table for Hire Status
    CreateHireStatusPivot valWb
    
    LogInfo "modSP2_HCCheck", "SP2_BuildHCCheck", "HC Check sheet built"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_HCCheck", "SP2_BuildHCCheck", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: BuildHCCheckHeader
' Purpose: Build the HC Check sheet header
'------------------------------------------------------------------------------
Private Sub BuildHCCheckHeader(ws As Worksheet)
    On Error Resume Next
    
    ws.Cells(1, 1).Value = "HK Payroll HC Check"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    
    ws.Cells(2, 1).Value = "Payroll Month: " & G.Payroll.PayrollMonth
    
    ' Headers
    ws.Cells(4, 1).Value = "Category"
    ws.Cells(4, 2).Value = "Previous Month"
    ws.Cells(4, 3).Value = "Current Month"
    ws.Cells(4, 4).Value = "Check"
    ws.Range("A4:D4").Font.Bold = True
    ws.Range("A4:D4").Interior.Color = RGB(200, 200, 200)
    
    ' Row labels
    ws.Cells(5, 1).Value = "Payroll Active HC"
    ws.Cells(6, 1).Value = "Terminated HC (Included)"
    ws.Cells(7, 1).Value = "Terminated HC (OC)"
    ws.Cells(8, 1).Value = "New Hire HC"
    ws.Cells(9, 1).Value = "Previous Month Terminated HC (额外表)"
    ws.Cells(10, 1).Value = ""
    ws.Cells(11, 1).Value = "Calculated HC"
    ws.Cells(11, 1).Font.Bold = True
    
    ws.Columns("A:D").AutoFit
End Sub

'------------------------------------------------------------------------------
' Sub: CalculatePayrollHC
' Purpose: Calculate Payroll Active HC from Payroll Report
'------------------------------------------------------------------------------
Private Sub CalculatePayrollHC(ws As Worksheet)
    Dim checkWs As Worksheet
    Dim lastRow As Long
    Dim hireStatusCol As Long
    Dim activeCount As Long
    Dim i As Long
    Dim hireStatus As String
    
    On Error GoTo ErrHandler
    
    Set checkWs = ws.Parent.Worksheets("Check Result")
    
    lastRow = checkWs.Cells(checkWs.Rows.Count, 1).End(xlUp).Row
    hireStatusCol = FindColumnByHeader(checkWs.Rows(4), "Hire Status")
    
    If hireStatusCol = 0 Then Exit Sub
    
    activeCount = 0
    For i = 5 To lastRow
        hireStatus = UCase(Trim(CStr(Nz(checkWs.Cells(i, hireStatusCol).Value, ""))))
        If hireStatus = "ACTIVE" Then
            activeCount = activeCount + 1
        End If
    Next i
    
    ' Write to Current Month column
    ws.Cells(5, 3).Value = activeCount
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_HCCheck", "CalculatePayrollHC", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: CalculateTerminatedHC
' Purpose: Calculate Terminated HC from Termination flexiform
'------------------------------------------------------------------------------
Private Sub CalculateTerminatedHC(ws As Worksheet)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim lastRow As Long, i As Long
    Dim headers As Object
    Dim termDate As Date
    Dim payDate As Date
    Dim includedCount As Long, ocCount As Long
    
    On Error GoTo ErrHandler
    
    filePath = GetInputFilePath("Termination")
    If Dir(filePath) = "" Then Exit Sub
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    Set headers = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column
        headers(UCase(Trim(CStr(srcWs.Cells(1, c).Value)))) = c
    Next c
    
    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    payDate = G.Payroll.PayDate
    
    includedCount = 0
    ocCount = 0
    
    For i = 2 To lastRow
        Dim termDateStr As String
        termDateStr = GetCellVal(srcWs, i, headers, "TERMINATION DATE")
        
        If termDateStr <> "" Then
            On Error Resume Next
            termDate = CDate(termDateStr)
            On Error GoTo ErrHandler
            
            ' Rule: If Termination Date + 7 > Pay Date -> Included, else OC
            If termDate + 7 > payDate Then
                includedCount = includedCount + 1
            Else
                ocCount = ocCount + 1
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    
    ' Write to Current Month column
    ws.Cells(6, 3).Value = includedCount
    ws.Cells(7, 3).Value = ocCount
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_HCCheck", "CalculateTerminatedHC", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Sub: CalculateNewHireHC
' Purpose: Calculate New Hire HC from NewHire flexiform
'------------------------------------------------------------------------------
Private Sub CalculateNewHireHC(ws As Worksheet)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim lastRow As Long
    Dim newHireCount As Long
    
    On Error GoTo ErrHandler
    
    filePath = GetInputFilePath("NewHire")
    If Dir(filePath) = "" Then Exit Sub
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    newHireCount = lastRow - 1 ' Exclude header
    
    wb.Close SaveChanges:=False
    
    ' Write to Current Month column
    ws.Cells(8, 3).Value = newHireCount
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_HCCheck", "CalculateNewHireHC", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Sub: CalculateHCFormula
' Purpose: Calculate HC check formula
' Formula: LastMonthPayrollHC - LastMonthTerminatedIncluded - CurrentMonthTerminatedOC + CurrentMonthNewHC
'------------------------------------------------------------------------------
Private Sub CalculateHCFormula(ws As Worksheet)
    Dim prevActiveHC As Double
    Dim prevTermIncluded As Double
    Dim currTermOC As Double
    Dim currNewHire As Double
    Dim calculatedHC As Double
    Dim actualHC As Double
    
    On Error Resume Next
    
    prevActiveHC = ToDouble(ws.Cells(5, 2).Value)
    prevTermIncluded = ToDouble(ws.Cells(6, 2).Value)
    currTermOC = ToDouble(ws.Cells(7, 3).Value)
    currNewHire = ToDouble(ws.Cells(8, 3).Value)
    actualHC = ToDouble(ws.Cells(5, 3).Value)
    
    calculatedHC = prevActiveHC - prevTermIncluded - currTermOC + currNewHire
    
    ws.Cells(11, 3).Value = calculatedHC
    
    ' Check column
    If calculatedHC = actualHC Then
        ws.Cells(11, 4).Value = "MATCH"
        ws.Cells(11, 4).Interior.Color = RGB(200, 255, 200)
    Else
        ws.Cells(11, 4).Value = "MISMATCH"
        ws.Cells(11, 4).Interior.Color = RGB(255, 200, 200)
    End If
End Sub

'------------------------------------------------------------------------------
' Sub: CreateHireStatusPivot
' Purpose: Create pivot table for Hire Status
'------------------------------------------------------------------------------
Private Sub CreateHireStatusPivot(valWb As Workbook)
    ' Pivot table creation would require more complex implementation
    ' Placeholder for now
    
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = valWb.Worksheets("HC Check")
    
    ws.Cells(14, 1).Value = "Hire Status Summary"
    ws.Cells(14, 1).Font.Bold = True
    
    ' Note: Full pivot table implementation would go here
    ' For simplicity, we'll just add a note
    ws.Cells(15, 1).Value = "(Pivot table to be created manually or via additional code)"
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
