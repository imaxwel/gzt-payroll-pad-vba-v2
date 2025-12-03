Attribute VB_Name = "modSP2_CheckResult_BenefitsTax"
'==============================================================================
' Module: modSP2_CheckResult_BenefitsTax
' Purpose: Subprocess 2 - Benefits for Tax Check columns
' Description: Validates Inspire Points Gross-up
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Sub: SP2_Check_BenefitsTax
' Purpose: Populate benefits for tax Check columns
'------------------------------------------------------------------------------
Public Sub SP2_Check_BenefitsTax(valWb As Workbook, weinIndex As Object)
    Dim ws As Worksheet
    
    On Error GoTo ErrHandler
    
    Set ws = valWb.Worksheets("Check Result")
    
    ' Process Inspire Points Gross-up
    ProcessInspireGrossUp ws, weinIndex
    
    LogInfo "modSP2_CheckResult_BenefitsTax", "SP2_Check_BenefitsTax", "Benefits for tax checks completed"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_BenefitsTax", "SP2_Check_BenefitsTax", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: ProcessInspireGrossUp
' Purpose: Calculate Inspire Points Gross-up
' Formula: ROUNDUP(Actual Payment - Amount / (1 - 0.17) * 0.17, 0)
'------------------------------------------------------------------------------
Private Sub ProcessInspireGrossUp(ws As Worksheet, weinIndex As Object)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim dataRange As Range
    Dim grouped As Object
    Dim lastRow As Long, lastCol As Long
    
    On Error GoTo ErrHandler
    
    filePath = GetInputFilePath("InspireAwards")
    If Dir(filePath) = "" Then Exit Sub
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    lastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column
    Set dataRange = srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(lastRow, lastCol))
    
    ' Filter for Inspire Points Value only
    Set grouped = GroupByEmployeeAndTypeFiltered(dataRange, "Employee ID", "One-Time Payment Plan", _
        "Actual Payment - Amount", "One-Time Payment Plan", Array("Inspire Points Value"))
    
    Dim key As Variant
    Dim parts() As String
    Dim empId As String, wein As String
    Dim col As Long, row As Long
    Dim inspireAmt As Double, grossUp As Double
    
    col = FindColumnByHeader(ws.Rows(4), "Inspire Points (Gross Up) 60701000 Check")
    If col = 0 Then
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    For Each key In grouped.Keys
        parts = Split(CStr(key), "|")
        If UBound(parts) >= 0 Then
            empId = parts(0)
            
            wein = WeinFromEmpId(empId)
            If wein = "" Then wein = empId
            
            If weinIndex.Exists(wein) Then
                row = weinIndex(wein)
                
                inspireAmt = grouped(key)
                
                ' Gross-up formula: ROUNDUP(Amount / (1 - 0.17) * 0.17, 0)
                If inspireAmt > 0 Then
                    grossUp = RoundUpInteger(inspireAmt / (1 - 0.17) * 0.17)
                    ws.Cells(row, col).Value = grossUp
                End If
            End If
        End If
    Next key
    
    wb.Close SaveChanges:=False
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_BenefitsTax", "ProcessInspireGrossUp", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub
