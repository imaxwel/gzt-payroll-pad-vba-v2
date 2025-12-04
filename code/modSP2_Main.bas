Attribute VB_Name = "modSP2_Main"
'==============================================================================
' Module: modSP2_Main
' Purpose: Subprocess 2 main orchestration
' Description: Creates HK Payroll Validation Output workbook
'==============================================================================
Option Explicit

' Output workbook reference
Private mValWb As Workbook
' WEIN index for Check Result sheet
Private mWeinIndex As Object

'------------------------------------------------------------------------------
' Sub: SP2_Execute
' Purpose: Main execution routine for Subprocess 2
'------------------------------------------------------------------------------
Public Sub SP2_Execute()
    On Error GoTo ErrHandler
    
    EnsureInitialised
    
    LogInfo "modSP2_Main", "SP2_Execute", "Starting Subprocess 2 execution"
    
    ' 1. Create Validation Output workbook
    Set mValWb = CreateValidationOutputWorkbook()
    
    ' 2. Build Check Result benchmark and WEIN index
    BuildBenchmarkAndIndex mValWb
    
    ' 3. Populate Check columns by group
    RunMasterDataChecks mValWb
    RunPayItemChecks mValWb
    RunIncentiveChecks mValWb
    RunFinalPaymentChecks mValWb
    RunContributionChecks mValWb
    RunBenefitsForTaxChecks mValWb
    
    ' 4. Compute Diff columns
    ComputeDiffs mValWb
    
    ' 5. Build HC Check sheet
    BuildHCCheck mValWb
    
    ' 6. Final formatting and save
    FinalizeValidationOutput mValWb
    
    LogInfo "modSP2_Main", "SP2_Execute", "Subprocess 2 execution completed"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_Main", "SP2_Execute", Err.Number, Err.Description
    Err.Raise Err.Number, "SP2_Execute", Err.Description
End Sub


'------------------------------------------------------------------------------
' Function: CreateValidationOutputWorkbook
' Purpose: Create the HK Payroll Validation Output workbook
' Returns: Workbook object
'------------------------------------------------------------------------------
Private Function CreateValidationOutputWorkbook() As Workbook
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fileName As String
    Dim filePath As String
    
    On Error GoTo ErrHandler
    
    ' Build filename with date
    fileName = "HK Payroll Validation Output " & Format(G.RunParams.RunDate, "YYYYMMDD") & ".xlsx"
    filePath = G.RunParams.OutputFolder & fileName
    
    LogInfo "modSP2_Main", "CreateValidationOutputWorkbook", "Creating: " & filePath
    
    ' Create new workbook
    Set wb = Workbooks.Add
    
    ' Rename first sheet to Check Result
    wb.Worksheets(1).Name = "Check Result"
    
    ' Add HC Check sheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    ws.Name = "HC Check"
    
    ' Delete any extra default sheets
    Application.DisplayAlerts = False
    Do While wb.Worksheets.count > 2
        wb.Worksheets(wb.Worksheets.count).Delete
    Loop
    Application.DisplayAlerts = True
    
    ' Save workbook
    wb.SaveAs filePath, xlOpenXMLWorkbook
    
    Set CreateValidationOutputWorkbook = wb
    Exit Function
    
ErrHandler:
    LogError "modSP2_Main", "CreateValidationOutputWorkbook", Err.Number, Err.Description
    Err.Raise Err.Number, "CreateValidationOutputWorkbook", Err.Description
End Function

'------------------------------------------------------------------------------
' Sub: BuildBenchmarkAndIndex
' Purpose: Copy Payroll Report to Check Result and build WEIN index
'------------------------------------------------------------------------------
Private Sub BuildBenchmarkAndIndex(valWb As Workbook)
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim destWs As Worksheet
    Dim filePath As String
    Dim lastRow As Long, lastCol As Long
    Dim srcRange As Range
    Dim weinCol As Long
    Dim i As Long
    
    On Error GoTo ErrHandler
    
    LogInfo "modSP2_Main", "BuildBenchmarkAndIndex", "Building benchmark from Payroll Report"
    
    Set destWs = valWb.Worksheets("Check Result")
    Set mWeinIndex = CreateObject("Scripting.Dictionary")
    
    ' Open Payroll Report
    filePath = GetInputFilePath("PayrollReport")
    
    If Dir(filePath) = "" Then
        LogError "modSP2_Main", "BuildBenchmarkAndIndex", 0, "Payroll Report not found: " & filePath
        Exit Sub
    End If
    
    Set srcWb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = srcWb.Worksheets(1)
    
    ' Find data range
    lastRow = srcWs.Cells(srcWs.Rows.count, 1).End(xlUp).row
    lastCol = srcWs.Cells(1, srcWs.Columns.count).End(xlToLeft).Column
    
    ' Copy to Check Result (starting at row 4 to leave room for summary)
    Set srcRange = srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(lastRow, lastCol))
    srcRange.Copy destWs.Cells(4, 1)
    
    ' Build WEIN index (try multiple field name variants)
    weinCol = FindColumnByHeader(destWs.Rows(4), "WEIN,WIN,WEINEmployee ID,Employee CodeWIN,Employee ID,EmployeeID")
    
    If weinCol > 0 Then
        For i = 5 To 4 + lastRow - 1
            Dim wein As String
            wein = Trim(CStr(Nz(destWs.Cells(i, weinCol).Value, "")))
            If wein <> "" And Not mWeinIndex.Exists(wein) Then
                mWeinIndex.Add wein, i
            End If
        Next i
    End If
    
    srcWb.Close SaveChanges:=False
    
    ' Add summary header row
    destWs.Cells(1, 1).Value = "HK Payroll Validation - Check Result"
    destWs.Cells(1, 1).Font.Bold = True
    destWs.Cells(1, 1).Font.Size = 14
    
    destWs.Cells(2, 1).Value = "Payroll Month: " & G.Payroll.payrollMonth
    destWs.Cells(2, 2).Value = "Run Date: " & Format(G.RunParams.RunDate, "yyyy-mm-dd")
    
    ' Row 3 will be used for FALSE counts
    destWs.Cells(3, 1).Value = "FALSE Count:"
    destWs.Cells(3, 1).Font.Bold = True
    
    LogInfo "modSP2_Main", "BuildBenchmarkAndIndex", "Built index with " & mWeinIndex.count & " WEINs"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_Main", "BuildBenchmarkAndIndex", Err.Number, Err.Description
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Sub: RunMasterDataChecks
' Purpose: Run master data validation checks
'------------------------------------------------------------------------------
Private Sub RunMasterDataChecks(valWb As Workbook)
    On Error GoTo ErrHandler
    
    LogInfo "modSP2_Main", "RunMasterDataChecks", "Running master data checks"
    
    ' Call SP2_CheckResult_MasterData module
    SP2_Check_MasterData valWb, mWeinIndex
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_Main", "RunMasterDataChecks", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: RunPayItemChecks
' Purpose: Run pay item validation checks
'------------------------------------------------------------------------------
Private Sub RunPayItemChecks(valWb As Workbook)
    On Error GoTo ErrHandler
    
    LogInfo "modSP2_Main", "RunPayItemChecks", "Running pay item checks"
    
    ' Call SP2_CheckResult_PayItems module
    SP2_Check_PayItems valWb, mWeinIndex
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_Main", "RunPayItemChecks", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: RunIncentiveChecks
' Purpose: Run incentive validation checks
'------------------------------------------------------------------------------
Private Sub RunIncentiveChecks(valWb As Workbook)
    On Error GoTo ErrHandler
    
    LogInfo "modSP2_Main", "RunIncentiveChecks", "Running incentive checks"
    
    ' Call SP2_CheckResult_Incentives module
    SP2_Check_Incentives valWb, mWeinIndex
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_Main", "RunIncentiveChecks", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: RunFinalPaymentChecks
' Purpose: Run final payment validation checks
'------------------------------------------------------------------------------
Private Sub RunFinalPaymentChecks(valWb As Workbook)
    On Error GoTo ErrHandler
    
    LogInfo "modSP2_Main", "RunFinalPaymentChecks", "Running final payment checks"
    
    ' Call SP2_CheckResult_FinalPayment module
    SP2_Check_FinalPayment valWb, mWeinIndex
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_Main", "RunFinalPaymentChecks", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: RunContributionChecks
' Purpose: Run contribution (MPF/ORSO) validation checks
'------------------------------------------------------------------------------
Private Sub RunContributionChecks(valWb As Workbook)
    On Error GoTo ErrHandler
    
    LogInfo "modSP2_Main", "RunContributionChecks", "Running contribution checks"
    
    ' Call SP2_CheckResult_Contribution module
    SP2_Check_Contribution valWb, mWeinIndex
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_Main", "RunContributionChecks", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: RunBenefitsForTaxChecks
' Purpose: Run benefits for tax validation checks
'------------------------------------------------------------------------------
Private Sub RunBenefitsForTaxChecks(valWb As Workbook)
    On Error GoTo ErrHandler
    
    LogInfo "modSP2_Main", "RunBenefitsForTaxChecks", "Running benefits for tax checks"
    
    ' Call SP2_CheckResult_BenefitsTax module
    SP2_Check_BenefitsTax valWb, mWeinIndex
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_Main", "RunBenefitsForTaxChecks", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: ComputeDiffs
' Purpose: Compute Diff columns (TRUE/FALSE comparison)
'------------------------------------------------------------------------------
Private Sub ComputeDiffs(valWb As Workbook)
    On Error GoTo ErrHandler
    
    LogInfo "modSP2_Main", "ComputeDiffs", "Computing diff columns"
    
    ' Call SP2_CheckResult_Diff module
    SP2_ComputeDiff valWb, mWeinIndex
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_Main", "ComputeDiffs", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: BuildHCCheck
' Purpose: Build HC Check sheet
'------------------------------------------------------------------------------
Private Sub BuildHCCheck(valWb As Workbook)
    On Error GoTo ErrHandler
    
    LogInfo "modSP2_Main", "BuildHCCheck", "Building HC Check sheet"
    
    ' Call SP2_HCCheck module
    SP2_BuildHCCheck valWb
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_Main", "BuildHCCheck", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: FinalizeValidationOutput
' Purpose: Apply final formatting and save
'------------------------------------------------------------------------------
Private Sub FinalizeValidationOutput(valWb As Workbook)
    Dim ws As Worksheet
    
    On Error GoTo ErrHandler
    
    LogInfo "modSP2_Main", "FinalizeValidationOutput", "Finalizing output"
    
    ' Apply formatting to Check Result
    Set ws = valWb.Worksheets("Check Result")
    ApplyStandardFormatting ws, 4
    
    ' Apply conditional formatting to Diff columns
    ApplyDiffFormatting ws
    
    ' Apply formatting to HC Check
    Set ws = valWb.Worksheets("HC Check")
    ApplyStandardFormatting ws
    
    ' Save workbook
    valWb.Save
    
    LogInfo "modSP2_Main", "FinalizeValidationOutput", "Output saved: " & valWb.fullName
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_Main", "FinalizeValidationOutput", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: ApplyDiffFormatting
' Purpose: Apply formatting to Diff columns
'------------------------------------------------------------------------------
Private Sub ApplyDiffFormatting(ws As Worksheet)
    Dim lastCol As Long
    Dim col As Long
    Dim headerValue As String
    Dim firstDiffCol As Long, lastDiffCol As Long
    Dim lastRow As Long
    
    On Error GoTo ErrHandler
    
    lastCol = ws.Cells(4, ws.Columns.count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    firstDiffCol = 0
    lastDiffCol = 0
    
    ' Find Diff columns (columns ending with "Diff")
    For col = 1 To lastCol
        headerValue = Trim(CStr(Nz(ws.Cells(4, col).Value, "")))
        If Right(UCase(headerValue), 4) = "DIFF" Then
            If firstDiffCol = 0 Then firstDiffCol = col
            lastDiffCol = col
            
            ' Apply conditional formatting
            ApplyConditionalFormatting ws.Range(ws.Cells(5, col), ws.Cells(lastRow, col))
        End If
    Next col
    
    ' Summarize Diff columns
    If firstDiffCol > 0 And lastDiffCol > 0 Then
        SummarizeDiffColumns ws, 3, 5, lastRow, firstDiffCol, lastDiffCol
    End If
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_Main", "ApplyDiffFormatting", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Function: GetWeinIndex
' Purpose: Get the WEIN index dictionary (for use by other modules)
'------------------------------------------------------------------------------
Public Function GetWeinIndex() As Object
    Set GetWeinIndex = mWeinIndex
End Function
