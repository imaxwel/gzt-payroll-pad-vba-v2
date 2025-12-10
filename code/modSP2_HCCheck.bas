Attribute VB_Name = "modSP2_HCCheck"
'==============================================================================
' Module: modSP2_HCCheck
' Purpose: Subprocess 2 - HC Check sheet
' Description: Headcount validation and cross-check
'              支持当月/上月文件读取，使用 modPathService 分层路径服务
'              按需求说明书结构输出：PivotTable区域 + HC明细表
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
    
    ' Build header and structure
    BuildHCCheckHeader ws
    
    ' Create pivot table for Hire Status (Row 3-5)
    CreateHireStatusPivot valWb, ws
    
    ' Build HC detail table (Row 10-15)
    BuildHCDetailTable ws
    
    ' Calculate headcounts - 当月数据
    CalculatePayrollHC ws, poCurrentMonth
    CalculateTerminatedHC ws, poCurrentMonth
    CalculateNewHireHC ws, poCurrentMonth
    CalculateExtraTableHC ws, poCurrentMonth
    
    ' Calculate headcounts - 上月数据
    CalculatePayrollHC ws, poPreviousMonth
    CalculateTerminatedHC ws, poPreviousMonth
    CalculateNewHireHC ws, poPreviousMonth
    CalculateExtraTableHC ws, poPreviousMonth
    
    ' Calculate HC formula and Check column
    CalculateHCFormula ws
    
    ' Apply formatting
    ApplyHCCheckFormatting ws
    
    LogInfo "modSP2_HCCheck", "SP2_BuildHCCheck", "HC Check sheet built"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_HCCheck", "SP2_BuildHCCheck", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: BuildHCCheckHeader
' Purpose: Build the HC Check sheet header (Row 1)
'------------------------------------------------------------------------------
Private Sub BuildHCCheckHeader(ws As Worksheet)
    On Error Resume Next
    
    ' Row 1: Payroll Month
    ws.Cells(1, 1).Value = "Payroll Month"
    ws.Cells(1, 2).Value = G.Payroll.payrollMonth
End Sub

'------------------------------------------------------------------------------
' Sub: BuildHCDetailTable
' Purpose: Build the HC detail table structure (Row 10-15)
'------------------------------------------------------------------------------
Private Sub BuildHCDetailTable(ws As Worksheet)
    Dim prevMonthName As String
    Dim currMonthName As String
    
    On Error Resume Next
    
    ' Get month names dynamically
    prevMonthName = GetMonthShortName(G.Payroll.prevMonthStart) & "(Previous Month)"
    currMonthName = GetMonthShortName(G.Payroll.monthStart) & "(Current Month)"
    
    ' Row 10: Column headers
    ws.Cells(10, 1).Value = ""
    ws.Cells(10, 2).Value = prevMonthName
    ws.Cells(10, 3).Value = currMonthName
    ws.Range("A10:C10").Font.Bold = True
    
    ' Row 11-15: Row labels
    ws.Cells(11, 1).Value = "Payroll HC"
    ws.Cells(12, 1).Value = "Current Month Terminated HC(included)"
    ws.Cells(13, 1).Value = "Current Month Terminated HC(OC)"
    ws.Cells(14, 1).Value = "Previous Month Terminated HC(included)"
    ws.Cells(15, 1).Value = "Current Month New HC"
    
    ' Initialize values to 0
    Dim r As Long
    For r = 11 To 15
        ws.Cells(r, 2).Value = 0
        ws.Cells(r, 3).Value = 0
    Next r
End Sub

'------------------------------------------------------------------------------
' Function: GetMonthShortName
' Purpose: Get short month name (e.g., "Nov", "Dec") from a date
'------------------------------------------------------------------------------
Private Function GetMonthShortName(d As Date) As String
    GetMonthShortName = Format(d, "mmm")
End Function

'------------------------------------------------------------------------------
' Sub: CreateHireStatusPivot
' Purpose: Create pivot table for Hire Status (Row 3-5)
'          Using current month's Payroll Report as data source
'------------------------------------------------------------------------------
Private Sub CreateHireStatusPivot(valWb As Workbook, ws As Worksheet)
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim lastRow As Long, i As Long
    Dim hireStatusCol As Long, weinCol As Long
    Dim statusCounts As Object
    Dim hireStatus As String
    Dim activeCount As Long
    Dim grandTotal As Long
    Dim calculatedCheck As Long
    
    On Error GoTo ErrHandler
    
    ' Row 3: Headers for pivot-like table
    ws.Cells(3, 1).Value = "Row Labels"
    ws.Cells(3, 2).Value = "Count of WEIN"
    ws.Cells(3, 3).Value = "Check"
    ws.Range("A3:C3").Font.Bold = True
    
    ' Get current month Payroll Report
    filePath = GetInputFilePathAuto("PayrollReport", poCurrentMonth)
    
    If Not FileExistsSafe(filePath) Then
        LogError "modSP2_HCCheck", "CreateHireStatusPivot", 0, "File not found: " & filePath
        ws.Cells(4, 1).Value = "Active"
        ws.Cells(4, 2).Value = 0
        ws.Cells(5, 1).Value = "Grand Total"
        ws.Cells(5, 2).Value = 0
        Exit Sub
    End If
    
    Set srcWb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = srcWb.Worksheets(1)
    
    lastRow = srcWs.Cells(srcWs.Rows.count, 1).End(xlUp).row
    hireStatusCol = FindColumnByHeader(srcWs.Rows(1), "Hire Status")
    weinCol = FindColumnByHeader(srcWs.Rows(1), "WEIN")
    
    If hireStatusCol = 0 Then
        srcWb.Close SaveChanges:=False
        ws.Cells(4, 1).Value = "Active"
        ws.Cells(4, 2).Value = 0
        ws.Cells(5, 1).Value = "Grand Total"
        ws.Cells(5, 2).Value = 0
        Exit Sub
    End If
    
    ' Count by Hire Status
    Set statusCounts = CreateObject("Scripting.Dictionary")
    grandTotal = 0
    
    For i = 2 To lastRow
        hireStatus = Trim(CStr(Nz(srcWs.Cells(i, hireStatusCol).Value, "")))
        If hireStatus <> "" Then
            If Not statusCounts.exists(hireStatus) Then
                statusCounts(hireStatus) = 0
            End If
            statusCounts(hireStatus) = statusCounts(hireStatus) + 1
            grandTotal = grandTotal + 1
        End If
    Next i
    
    srcWb.Close SaveChanges:=False
    
    ' Write Active count (Row 4)
    activeCount = 0
    If statusCounts.exists("Active") Then activeCount = statusCounts("Active")
    
    ws.Cells(4, 1).Value = "Active"
    ws.Cells(4, 2).Value = activeCount
    
    ' Row 5: Grand Total
    ws.Cells(5, 1).Value = "Grand Total"
    ws.Cells(5, 2).Value = grandTotal
    ws.Range("A5:B5").Font.Bold = True
    ws.Range("A5:B5").Interior.Color = RGB(255, 255, 200)  ' Light yellow for Grand Total
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_HCCheck", "CreateHireStatusPivot", Err.Number, Err.Description
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Sub: CalculatePayrollHC
' Purpose: Calculate Payroll HC from Payroll Report (Active status count)
' Parameters:
'   ws - HC Check worksheet
'   offset - 期间偏移量 (poCurrentMonth 或 poPreviousMonth)
' Writes to Row 11
'------------------------------------------------------------------------------
Private Sub CalculatePayrollHC(ws As Worksheet, offset As ePeriodOffset)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim lastRow As Long, i As Long
    Dim hireStatusCol As Long
    Dim activeCount As Long
    Dim hireStatus As String
    Dim targetCol As Long
    
    On Error GoTo ErrHandler
    
    ' 确定写入列: 上月=2, 当月=3
    targetCol = IIf(offset = poPreviousMonth, 2, 3)
    
    ' 使用新路径服务获取文件路径
    filePath = GetInputFilePathAuto("PayrollReport", offset)
    
    LogInfo "modSP2_HCCheck", "CalculatePayrollHC", _
        "Reading Payroll Report (" & GetPeriodDescription(offset) & "): " & filePath
    
    If Not FileExistsSafe(filePath) Then
        LogError "modSP2_HCCheck", "CalculatePayrollHC", 0, _
            "File not found: " & filePath
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    lastRow = srcWs.Cells(srcWs.Rows.count, 1).End(xlUp).row
    hireStatusCol = FindColumnByHeader(srcWs.Rows(1), "Hire Status")
    
    If hireStatusCol = 0 Then
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    activeCount = 0
    For i = 2 To lastRow
        hireStatus = UCase(Trim(CStr(Nz(srcWs.Cells(i, hireStatusCol).Value, ""))))
        If hireStatus = "ACTIVE" Then
            activeCount = activeCount + 1
        End If
    Next i
    
    wb.Close SaveChanges:=False
    
    ' Write to Row 11 (Payroll HC)
    ws.Cells(11, targetCol).Value = activeCount
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_HCCheck", "CalculatePayrollHC", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Sub: CalculateTerminatedHC
' Purpose: Calculate Terminated HC from Termination flexiform
' Parameters:
'   ws - HC Check worksheet
'   offset - 期间偏移量 (poCurrentMonth 或 poPreviousMonth)
' Writes to Row 12 (included) and Row 13 (OC)
'------------------------------------------------------------------------------
Private Sub CalculateTerminatedHC(ws As Worksheet, offset As ePeriodOffset)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim lastRow As Long, i As Long
    Dim headers As Object
    Dim termDate As Date
    Dim payDate As Date
    Dim includedCount As Long, ocCount As Long
    Dim targetCol As Long
    
    On Error GoTo ErrHandler
    
    ' 确定写入列: 上月=2, 当月=3
    targetCol = IIf(offset = poPreviousMonth, 2, 3)
    
    ' 使用新路径服务获取文件路径
    filePath = GetInputFilePathAuto("Termination", offset)
    
    LogInfo "modSP2_HCCheck", "CalculateTerminatedHC", _
        "Reading Termination (" & GetPeriodDescription(offset) & "): " & filePath
    
    If Not FileExistsSafe(filePath) Then
        LogError "modSP2_HCCheck", "CalculateTerminatedHC", 0, _
            "File not found: " & filePath
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    Set headers = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To srcWs.Cells(1, srcWs.Columns.count).End(xlToLeft).Column
        headers(UCase(Trim(CStr(srcWs.Cells(1, c).Value)))) = c
    Next c
    
    lastRow = srcWs.Cells(srcWs.Rows.count, 1).End(xlUp).row
    payDate = G.Payroll.payDate
    
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
    
    ' Write to Row 12 (Current Month Terminated HC included) and Row 13 (OC)
    ws.Cells(12, targetCol).Value = includedCount
    ws.Cells(13, targetCol).Value = ocCount
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_HCCheck", "CalculateTerminatedHC", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Sub: CalculateNewHireHC
' Purpose: Calculate New Hire HC from NewHire flexiform
' Parameters:
'   ws - HC Check worksheet
'   offset - 期间偏移量 (poCurrentMonth 或 poPreviousMonth)
' Writes to Row 15
'------------------------------------------------------------------------------
Private Sub CalculateNewHireHC(ws As Worksheet, offset As ePeriodOffset)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim lastRow As Long
    Dim newHireCount As Long
    Dim targetCol As Long
    
    On Error GoTo ErrHandler
    
    ' 确定写入列: 上月=2, 当月=3
    targetCol = IIf(offset = poPreviousMonth, 2, 3)
    
    ' 使用新路径服务获取文件路径
    filePath = GetInputFilePathAuto("NewHire", offset)
    
    LogInfo "modSP2_HCCheck", "CalculateNewHireHC", _
        "Reading NewHire (" & GetPeriodDescription(offset) & "): " & filePath
    
    If Not FileExistsSafe(filePath) Then
        LogError "modSP2_HCCheck", "CalculateNewHireHC", 0, _
            "File not found: " & filePath
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcWs = wb.Worksheets(1)
    
    lastRow = srcWs.Cells(srcWs.Rows.count, 1).End(xlUp).row
    newHireCount = lastRow - 1 ' Exclude header
    If newHireCount < 0 Then newHireCount = 0
    
    wb.Close SaveChanges:=False
    
    ' Write to Row 15 (Current Month New HC)
    ws.Cells(15, targetCol).Value = newHireCount
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_HCCheck", "CalculateNewHireHC", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Sub: CalculateExtraTableHC
' Purpose: Calculate HC from 额外表 (Extra Table) - Previous Month Terminated HC
' Parameters:
'   ws - HC Check worksheet
'   offset - 期间偏移量 (poCurrentMonth 或 poPreviousMonth)
' Writes to Row 14
'------------------------------------------------------------------------------
Private Sub CalculateExtraTableHC(ws As Worksheet, offset As ePeriodOffset)
    Dim wb As Workbook
    Dim srcWs As Worksheet
    Dim filePath As String
    Dim lastRow As Long
    Dim extraCount As Long
    Dim targetCol As Long
    Dim needClose As Boolean
    
    On Error GoTo ErrHandler
    
    ' 确定写入列: 上月=2, 当月=3
    targetCol = IIf(offset = poPreviousMonth, 2, 3)
    
    ' 使用新路径服务获取文件路径
    filePath = GetInputFilePathAuto("ExtraTable", offset)
    
    LogInfo "modSP2_HCCheck", "CalculateExtraTableHC", _
        "Reading ExtraTable (" & GetPeriodDescription(offset) & "): " & filePath
    
    If Not FileExistsSafe(filePath) Then
        LogError "modSP2_HCCheck", "CalculateExtraTableHC", 0, _
            "File not found: " & filePath
        ws.Cells(14, targetCol).Value = 0
        Exit Sub
    End If
    
    ' 检查文件是否已经打开，避免重复打开导致错误
    needClose = False
    On Error Resume Next
    Set wb = Workbooks(Dir(filePath))
    On Error GoTo ErrHandler
    
    If wb Is Nothing Then
        Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=0)
        needClose = True
    End If
    
    ' 检查工作簿是否成功打开
    If wb Is Nothing Then
        LogError "modSP2_HCCheck", "CalculateExtraTableHC", 0, _
            "Failed to open workbook: " & filePath
        ws.Cells(14, targetCol).Value = 0
        Exit Sub
    End If
    
    ' 检查工作簿是否有工作表
    If wb.Worksheets.count = 0 Then
        LogError "modSP2_HCCheck", "CalculateExtraTableHC", 0, _
            "Workbook has no worksheets: " & filePath
        If needClose Then wb.Close SaveChanges:=False
        ws.Cells(14, targetCol).Value = 0
        Exit Sub
    End If
    
    ' Try to find "Previous Month Terminated HC" sheet, fallback to first sheet
    On Error Resume Next
    Set srcWs = wb.Worksheets("Previous Month Terminated HC")
    If srcWs Is Nothing Then
        Set srcWs = wb.Worksheets(1)
    End If
    On Error GoTo ErrHandler
    
    ' 检查工作表对象是否有效
    If srcWs Is Nothing Then
        LogError "modSP2_HCCheck", "CalculateExtraTableHC", 0, _
            "Failed to get worksheet from: " & filePath
        If needClose Then wb.Close SaveChanges:=False
        ws.Cells(14, targetCol).Value = 0
        Exit Sub
    End If
    
    lastRow = srcWs.Cells(srcWs.Rows.count, 1).End(xlUp).row
    extraCount = lastRow - 1 ' Exclude header
    If extraCount < 0 Then extraCount = 0
    
    ' 只关闭我们自己打开的工作簿
    If needClose And Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    
    ' Write to Row 14 (Previous Month Terminated HC included)
    ws.Cells(14, targetCol).Value = extraCount
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_HCCheck", "CalculateExtraTableHC", Err.Number, Err.Description
    On Error Resume Next
    If needClose And Not wb Is Nothing Then wb.Close SaveChanges:=False
    ws.Cells(14, targetCol).Value = 0
End Sub

'------------------------------------------------------------------------------
' Sub: CalculateHCFormula
' Purpose: Calculate HC check formula and write to Check column
' Formula: Last Month's Payroll HC - Last Month's Current Month Terminated HC (included)
'          - Current Month's Current Month Terminated HC (OC) + Current Month's Current Month New HC
' Check column is placed at Row 4, Column C (next to Active row in pivot area)
'------------------------------------------------------------------------------
Private Sub CalculateHCFormula(ws As Worksheet)
    Dim prevPayrollHC As Double
    Dim prevTermIncluded As Double
    Dim currTermOC As Double
    Dim currNewHire As Double
    Dim calculatedHC As Double
    Dim actualActiveHC As Double
    
    On Error Resume Next
    
    ' Read values from HC detail table (Row 11-15)
    prevPayrollHC = ToDouble(ws.Cells(11, 2).Value)      ' Previous Month Payroll HC
    prevTermIncluded = ToDouble(ws.Cells(12, 2).Value)   ' Previous Month - Current Month Terminated HC (included)
    currTermOC = ToDouble(ws.Cells(13, 3).Value)         ' Current Month - Current Month Terminated HC (OC)
    currNewHire = ToDouble(ws.Cells(15, 3).Value)        ' Current Month - Current Month New HC
    
    ' Actual Active HC from pivot table (Row 4, Column B)
    actualActiveHC = ToDouble(ws.Cells(4, 2).Value)
    
    ' Calculate: prevPayrollHC - prevTermIncluded - currTermOC + currNewHire
    calculatedHC = prevPayrollHC - prevTermIncluded - currTermOC + currNewHire
    
    ' Write calculated value to Check column (Row 4, Column C)
    ws.Cells(4, 3).Value = calculatedHC
End Sub

'------------------------------------------------------------------------------
' Sub: ApplyHCCheckFormatting
' Purpose: Apply formatting to HC Check sheet
'------------------------------------------------------------------------------
Private Sub ApplyHCCheckFormatting(ws As Worksheet)
    On Error Resume Next
    
    ' Format pivot table area (Row 3-5)
    ws.Range("A3:C3").Interior.Color = RGB(200, 200, 200)  ' Header gray
    
    ' Format Grand Total row
    ws.Range("A5:B5").Font.Bold = True
    ws.Range("A5:B5").Interior.Color = RGB(255, 255, 200)  ' Light yellow
    
    ' Format HC detail table header (Row 10)
    ws.Range("A10:C10").Font.Bold = True
    ws.Range("A10:C10").Interior.Color = RGB(0, 0, 0)      ' Black background
    ws.Range("A10:C10").Font.Color = RGB(255, 255, 255)    ' White text
    
    ' Highlight current month column (Column C) with yellow for data rows
    ws.Range("C11:C15").Interior.Color = RGB(255, 255, 0)  ' Yellow
    
    ' Auto-fit columns
    ws.Columns("A:C").AutoFit
    
    ' Set minimum column width for readability
    If ws.Columns("A").ColumnWidth < 35 Then ws.Columns("A").ColumnWidth = 35
    If ws.Columns("B").ColumnWidth < 20 Then ws.Columns("B").ColumnWidth = 20
    If ws.Columns("C").ColumnWidth < 20 Then ws.Columns("C").ColumnWidth = 20
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
