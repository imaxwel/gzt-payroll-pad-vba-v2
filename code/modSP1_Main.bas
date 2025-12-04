Attribute VB_Name = "modSP1_Main"
'==============================================================================
' Module: modSP1_Main
' Purpose: Subprocess 1 main orchestration
' Description: Creates Flexi form output workbook with all required sheets
'==============================================================================
Option Explicit

' Output workbook reference
Private mFlexWb As Workbook

'------------------------------------------------------------------------------
' Sub: SP1_Execute
' Purpose: Main execution routine for Subprocess 1
'------------------------------------------------------------------------------
Public Sub SP1_Execute()
    On Error GoTo ErrHandler
    
    EnsureInitialised
    
    LogInfo "modSP1_Main", "SP1_Execute", "Starting Subprocess 1 execution"
    
    ' 1. Create Flexi form output workbook
    Set mFlexWb = CreateFlexiOutputWorkbook()
    
    ' 2. Load raw flexiform data into relevant sheets
    SP1_LoadFlexiformData mFlexWb
    
    ' 3. Populate Attendance sheet (leave days & adjustments)
    SP1_PopulateAttendance mFlexWb
    
    ' 4. Populate VariablePay sheet (variable pay, EAO inputs, etc.)
    SP1_PopulateVariablePay mFlexWb
    
    ' 5. Final formatting & save
    SP1_FinalizeFlexOutput mFlexWb
    
    LogInfo "modSP1_Main", "SP1_Execute", "Subprocess 1 execution completed"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Main", "SP1_Execute", Err.Number, Err.Description
    Err.Raise Err.Number, "SP1_Execute", Err.Description
End Sub

'------------------------------------------------------------------------------
' Function: CreateFlexiOutputWorkbook
' Purpose: Create the Flexi form output workbook with required sheets
' Returns: Workbook object
'------------------------------------------------------------------------------
Private Function CreateFlexiOutputWorkbook() As Workbook
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fileName As String
    Dim filePath As String
    Dim sheetNames As Variant
    Dim i As Long
    
    On Error GoTo ErrHandler
    
    ' Build filename with date
    fileName = "Flexi form out put " & Format(G.Payroll.payDate, "yyyymmdd") & ".xlsx"
    filePath = G.RunParams.OutputFolder & fileName
    
    LogInfo "modSP1_Main", "CreateFlexiOutputWorkbook", "Creating: " & filePath
    
    ' Create new workbook
    Set wb = Workbooks.Add
    
    ' Define sheet names
    sheetNames = Array("NewHire", "InformationChange", "SalaryChange", "Termination", "Attendance", "VariablePay")
    
    ' Rename first sheet and add others
    wb.Worksheets(1).Name = sheetNames(0)
    
    For i = 1 To UBound(sheetNames)
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.Name = sheetNames(i)
    Next i
    
    ' Delete any extra default sheets
    Application.DisplayAlerts = False
    Do While wb.Worksheets.count > UBound(sheetNames) + 1
        wb.Worksheets(wb.Worksheets.count).Delete
    Loop
    Application.DisplayAlerts = True
    
    ' Save workbook
    wb.SaveAs filePath, xlOpenXMLWorkbook
    
    Set CreateFlexiOutputWorkbook = wb
    Exit Function
    
ErrHandler:
    LogError "modSP1_Main", "CreateFlexiOutputWorkbook", Err.Number, Err.Description
    Err.Raise Err.Number, "CreateFlexiOutputWorkbook", Err.Description
End Function

'------------------------------------------------------------------------------
' Sub: SP1_LoadFlexiformData
' Purpose: Load raw flexiform data into output workbook sheets
'------------------------------------------------------------------------------
Public Sub SP1_LoadFlexiformData(flexWb As Workbook)
    On Error GoTo ErrHandler
    
    LogInfo "modSP1_Main", "SP1_LoadFlexiformData", "Loading flexiform data"
    
    ' Copy NewHire data
    CopyFlexiformSheet "NewHire", flexWb.Worksheets("NewHire")
    
    ' Copy DataChange -> InformationChange
    CopyFlexiformSheet "DataChange", flexWb.Worksheets("InformationChange")
    
    ' Copy Comp -> SalaryChange
    CopyFlexiformSheet "Comp", flexWb.Worksheets("SalaryChange")
    
    ' Copy Termination data
    CopyFlexiformSheet "Termination", flexWb.Worksheets("Termination")
    
    ' Copy Attendance data
    CopyFlexiformSheet "Attendance", flexWb.Worksheets("Attendance")
    
    ' Copy Variable -> VariablePay base data
    CopyFlexiformSheet "Variable", flexWb.Worksheets("VariablePay")
    
    ' Add additional columns to VariablePay
    AddVariablePayColumns flexWb.Worksheets("VariablePay")
    
    LogInfo "modSP1_Main", "SP1_LoadFlexiformData", "Flexiform data loaded"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Main", "SP1_LoadFlexiformData", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: CopyFlexiformSheet
' Purpose: Copy data from a flexiform template to output sheet
'------------------------------------------------------------------------------
Private Sub CopyFlexiformSheet(logicalName As String, destWs As Worksheet)
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim lastRow As Long, lastCol As Long
    Dim srcRange As Range
    
    On Error GoTo ErrHandler
    
    filePath = GetInputFilePath(logicalName)
    
    If Dir(filePath) = "" Then
        LogWarning "modSP1_Main", "CopyFlexiformSheet", "File not found: " & filePath
        Exit Sub
    End If
    
    Set srcWb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = srcWb.Worksheets(1)
    
    ' Find data range
    lastRow = srcWs.Cells(srcWs.Rows.count, 1).End(xlUp).row
    lastCol = srcWs.Cells(1, srcWs.Columns.count).End(xlToLeft).Column
    
    If lastRow >= 1 And lastCol >= 1 Then
        Set srcRange = srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(lastRow, lastCol))
        srcRange.Copy destWs.Cells(1, 1)
    End If
    
    srcWb.Close SaveChanges:=False
    
    LogInfo "modSP1_Main", "CopyFlexiformSheet", "Copied " & logicalName & ": " & lastRow & " rows"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Main", "CopyFlexiformSheet", Err.Number, Err.Description
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Sub: AddVariablePayColumns
' Purpose: Add additional columns to VariablePay sheet
'------------------------------------------------------------------------------
Private Sub AddVariablePayColumns(ws As Worksheet)
    Dim lastCol As Long
    Dim newHeaders As Variant
    Dim i As Long
    Dim insertCol As Long
    
    On Error GoTo ErrHandler
    
    ' Find the column after "Adjustment of Parental Paid Time Off (PPTO) payment"
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    ' Find insertion point (after PPTO column or at end)
    insertCol = FindColumnByHeader(ws.Rows(1), "Adjustment of Parental Paid Time Off (PPTO) payment")
    If insertCol = 0 Then
        insertCol = lastCol + 1
    Else
        insertCol = insertCol + 1
    End If
    
    ' New columns to add
    newHeaders = Array( _
        "IA Pay Split", _
        "MPF Relevant Income Rewrite", _
        "MPF VC Relevant Income Rewrite", _
        "Paternity Leave payment adjustment", _
        "PPTO EAO Rate input", _
        "Flexible benefits")
    
    ' Insert columns
    For i = 0 To UBound(newHeaders)
        ws.Columns(insertCol + i).Insert Shift:=xlToRight
        ws.Cells(1, insertCol + i).Value = newHeaders(i)
    Next i
    
    LogInfo "modSP1_Main", "AddVariablePayColumns", "Added " & (UBound(newHeaders) + 1) & " columns to VariablePay"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Main", "AddVariablePayColumns", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: SP1_FinalizeFlexOutput
' Purpose: Apply final formatting and save the output workbook
'------------------------------------------------------------------------------
Public Sub SP1_FinalizeFlexOutput(flexWb As Workbook)
    Dim ws As Worksheet
    
    On Error GoTo ErrHandler
    
    LogInfo "modSP1_Main", "SP1_FinalizeFlexOutput", "Finalizing output"
    
    ' Apply formatting to each sheet
    For Each ws In flexWb.Worksheets
        ApplyStandardFormatting ws
    Next ws
    
    ' Add run summary sheet
    AddRunSummary flexWb
    
    ' Save workbook
    flexWb.Save
    
    LogInfo "modSP1_Main", "SP1_FinalizeFlexOutput", "Output saved: " & flexWb.fullName
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Main", "SP1_FinalizeFlexOutput", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: AddRunSummary
' Purpose: Add a summary sheet with run information
'------------------------------------------------------------------------------
Private Sub AddRunSummary(wb As Workbook)
    Dim ws As Worksheet
    Dim r As Long
    Dim sheetWs As Worksheet
    
    On Error GoTo ErrHandler
    
    ' Add summary sheet at the beginning
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    ws.Name = "RunSummary"
    
    r = 1
    ws.Cells(r, 1).Value = "HK Payroll Subprocess 1 - Run Summary"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 14
    
    r = r + 2
    ws.Cells(r, 1).Value = "Payroll Month:"
    ws.Cells(r, 2).Value = G.Payroll.payrollMonth
    
    r = r + 1
    ws.Cells(r, 1).Value = "Run Date:"
    ws.Cells(r, 2).Value = Format(G.RunParams.RunDate, "yyyy-mm-dd")
    
    r = r + 1
    ws.Cells(r, 1).Value = "Pay Date:"
    ws.Cells(r, 2).Value = Format(G.Payroll.payDate, "yyyy-mm-dd")
    
    r = r + 1
    ws.Cells(r, 1).Value = "Cutoff Date:"
    ws.Cells(r, 2).Value = Format(G.Payroll.CurrentCutoff, "yyyy-mm-dd")
    
    r = r + 2
    ws.Cells(r, 1).Value = "Sheet"
    ws.Cells(r, 2).Value = "Record Count"
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Font.Bold = True
    
    ' Count records in each sheet
    For Each sheetWs In wb.Worksheets
        If sheetWs.Name <> "RunSummary" Then
            r = r + 1
            ws.Cells(r, 1).Value = sheetWs.Name
            ws.Cells(r, 2).Value = sheetWs.Cells(sheetWs.Rows.count, 1).End(xlUp).row - 1
        End If
    Next sheetWs
    
    ws.Columns("A:B").AutoFit
    
    Exit Sub
    
ErrHandler:
    LogError "modSP1_Main", "AddRunSummary", Err.Number, Err.Description
End Sub
