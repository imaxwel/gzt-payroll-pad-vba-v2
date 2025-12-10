Attribute VB_Name = "modConfigService"
'==============================================================================
' Module: modConfigService
' Purpose: Configuration and parameter reading services
' Description: Handles reading from config.xlsx and Additional table.xlsx
'==============================================================================
Option Explicit

Private Const CONFIG_FILE As String = "config.xlsx"
Private Const ADDITIONAL_TABLE_FILE As String = "Additional table.xlsx"

'------------------------------------------------------------------------------
' Function: LoadRunParamsFromWorkbook
' Purpose: Load run parameters from the Runtime sheet
' Returns: tRunParams structure with all parameters
'------------------------------------------------------------------------------
Public Function LoadRunParamsFromWorkbook() As tRunParams
    Dim p As tRunParams
    
    On Error GoTo ErrHandler
    
    With ThisWorkbook.Worksheets("Runtime")
        p.InputFolder = Trim(CStr(.Range("InputFolder").Value))
        p.OutputFolder = Trim(CStr(.Range("OutputFolder").Value))
        p.ConfigFolder = Trim(CStr(.Range("ConfigFolder").Value))
        p.payrollMonth = Trim(CStr(.Range("PayrollMonth").Value))
        p.RunDate = CDate(.Range("RunDate").Value)
        p.LogFolder = Trim(CStr(.Range("LogFolder").Value))
    End With
    
    ' Ensure folders end with backslash
    If Right(p.InputFolder, 1) <> "\" Then p.InputFolder = p.InputFolder & "\"
    If Right(p.OutputFolder, 1) <> "\" Then p.OutputFolder = p.OutputFolder & "\"
    If Right(p.ConfigFolder, 1) <> "\" Then p.ConfigFolder = p.ConfigFolder & "\"
    If Right(p.LogFolder, 1) <> "\" Then p.LogFolder = p.LogFolder & "\"
    
    LoadRunParamsFromWorkbook = p
    Exit Function
    
ErrHandler:
    LogError "modConfigService", "LoadRunParamsFromWorkbook", Err.Number, Err.Description
    Err.Raise Err.Number, "LoadRunParamsFromWorkbook", Err.Description
End Function

'------------------------------------------------------------------------------
' Function: GetPayrollContext
' Purpose: Get payroll calendar context for a given payroll month
' Parameters:
'   payrollMonth - Payroll month in "YYYYMM" format
' Returns: tPayrollContext structure
'------------------------------------------------------------------------------
Public Function GetPayrollContext(payrollMonth As String) As tPayrollContext
    Dim ctx As tPayrollContext
    Dim yr As Integer, mo As Integer
    
    On Error GoTo ErrHandler
    
    ' Parse payroll month
    yr = CInt(Left(payrollMonth, 4))
    mo = CInt(Right(payrollMonth, 2))
    
    ctx.payrollMonth = payrollMonth
    ctx.monthStart = DateSerial(yr, mo, 1)
    ctx.monthEnd = DateSerial(yr, mo + 1, 0)
    ctx.PrevMonthStart = DateSerial(yr, mo - 1, 1)
    ctx.PrevMonthEnd = DateSerial(yr, mo, 0)
    
    ' Calculate calendar days
    ctx.CalendarDaysCurrentMonth = Day(ctx.monthEnd)
    ctx.CalendarDaysPrevMonth = Day(ctx.PrevMonthEnd)
    
    ' Try to load cutoff and pay dates from config
    Dim configWb As Workbook
    Set configWb = OpenConfigWorkbook()
    
    If Not configWb Is Nothing Then
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = configWb.Worksheets("PayrollSchedule")
        On Error GoTo ErrHandler
        
        If Not ws Is Nothing Then
            Dim lastRow As Long, i As Long
            lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
            
            For i = 2 To lastRow
                If Trim(CStr(ws.Cells(i, 1).Value)) = payrollMonth Then
                    ctx.CurrentCutoff = CDate(ws.Cells(i, 2).Value)
                    ctx.payDate = CDate(ws.Cells(i, 3).Value)
                    Exit For
                End If
            Next i
            
            ' Get previous month cutoff
            Dim prevMonth As String
            prevMonth = Format(DateAdd("m", -1, ctx.monthStart), "YYYYMM")
            For i = 2 To lastRow
                If Trim(CStr(ws.Cells(i, 1).Value)) = prevMonth Then
                    ctx.PreviousCutoff = CDate(ws.Cells(i, 2).Value)
                    Exit For
                End If
            Next i
        End If
    End If
    
    ' Default values if not found in config
    If ctx.payDate = 0 Then ctx.payDate = ctx.monthEnd
    If ctx.CurrentCutoff = 0 Then ctx.CurrentCutoff = ctx.monthEnd
    If ctx.PreviousCutoff = 0 Then ctx.PreviousCutoff = ctx.PrevMonthEnd
    
    GetPayrollContext = ctx
    Exit Function
    
ErrHandler:
    LogError "modConfigService", "GetPayrollContext", Err.Number, Err.Description
    ' Return partial context with defaults
    ctx.payDate = ctx.monthEnd
    ctx.CurrentCutoff = ctx.monthEnd
    ctx.PreviousCutoff = ctx.PrevMonthEnd
    GetPayrollContext = ctx
End Function

'------------------------------------------------------------------------------
' Function: OpenConfigWorkbook
' Purpose: Open the config.xlsx workbook
' Returns: Workbook object or Nothing if not found
'------------------------------------------------------------------------------
Public Function OpenConfigWorkbook() As Workbook
    Dim filePath As String
    
    On Error GoTo ErrHandler
    
    If Not G.configWb Is Nothing Then
        Set OpenConfigWorkbook = G.configWb
        Exit Function
    End If
    
    filePath = G.RunParams.ConfigFolder & CONFIG_FILE
    
    If Dir(filePath) <> "" Then
        Set G.configWb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
        Set OpenConfigWorkbook = G.configWb
    End If
    
    Exit Function
    
ErrHandler:
    LogError "modConfigService", "OpenConfigWorkbook", Err.Number, Err.Description
    Set OpenConfigWorkbook = Nothing
End Function

'------------------------------------------------------------------------------
' Function: OpenExtraTableWorkbook
' Purpose: Open the Additional table.xlsx workbook
' Returns: Workbook object or Nothing if not found
'------------------------------------------------------------------------------
Public Function OpenExtraTableWorkbook() As Workbook
    Dim filePath As String
    
    On Error GoTo ErrHandler
    
    If Not G.ExtraTableWb Is Nothing Then
        Set OpenExtraTableWorkbook = G.ExtraTableWb
        Exit Function
    End If
    
    filePath = G.RunParams.InputFolder & ADDITIONAL_TABLE_FILE
    
    If Dir(filePath) <> "" Then
        Set G.ExtraTableWb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
        Set OpenExtraTableWorkbook = G.ExtraTableWb
    End If
    
    Exit Function
    
ErrHandler:
    LogError "modConfigService", "OpenExtraTableWorkbook", Err.Number, Err.Description
    Set OpenExtraTableWorkbook = Nothing
End Function

'------------------------------------------------------------------------------
' Function: GetExchangeRate
' Purpose: Get exchange rate from config
' Parameters:
'   rateName - Name of the rate (e.g., "RSU_Global", "RSU_EY")
' Returns: Exchange rate as Double
'------------------------------------------------------------------------------
Public Function GetExchangeRate(rateName As String) As Double
    Dim configWb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    
    On Error GoTo ErrHandler
    
    GetExchangeRate = 1# ' Default
    
    Set configWb = OpenConfigWorkbook()
    If configWb Is Nothing Then Exit Function
    
    On Error Resume Next
    Set ws = configWb.Worksheets("ExchangeRates")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then Exit Function
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If UCase(Trim(CStr(ws.Cells(i, 1).Value))) = UCase(rateName) Then
            GetExchangeRate = CDbl(ws.Cells(i, 2).Value)
            Exit For
        End If
    Next i
    
    Exit Function
    
ErrHandler:
    LogError "modConfigService", "GetExchangeRate", Err.Number, Err.Description
    GetExchangeRate = 1#
End Function

'------------------------------------------------------------------------------
' Function: IsSpecialMonth
' Purpose: Check if current payroll month has a special flag
' Parameters:
'   flagName - Name of the flag (e.g., "IsAIPMonth", "IsRSUDivMonth")
' Returns: Boolean
'------------------------------------------------------------------------------
Public Function IsSpecialMonth(flagName As String) As Boolean
    Dim configWb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, col As Long
    Dim headerRow As Range
    
    On Error GoTo ErrHandler
    
    IsSpecialMonth = False
    
    Set configWb = OpenConfigWorkbook()
    If configWb Is Nothing Then Exit Function
    
    On Error Resume Next
    Set ws = configWb.Worksheets("PayrollSchedule")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then Exit Function
    
    ' Find column by header name
    Set headerRow = ws.Rows(1)
    col = 0
    For i = 1 To ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        If UCase(Trim(CStr(ws.Cells(1, i).Value))) = UCase(flagName) Then
            col = i
            Exit For
        End If
    Next i
    
    If col = 0 Then Exit Function
    
    ' Find row for current payroll month
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If Trim(CStr(ws.Cells(i, 1).Value)) = G.Payroll.payrollMonth Then
            IsSpecialMonth = CBool(ws.Cells(i, col).Value)
            Exit For
        End If
    Next i
    
    Exit Function
    
ErrHandler:
    LogError "modConfigService", "IsSpecialMonth", Err.Number, Err.Description
    IsSpecialMonth = False
End Function

'------------------------------------------------------------------------------
' Function: GetInputFilePath
' Purpose: Get full path for an input file by logical name
' Parameters:
'   logicalName - Logical name of the file
' Returns: Full file path
'------------------------------------------------------------------------------
Public Function GetInputFilePath(logicalName As String) As String
    Dim fileName As String
    
    Select Case UCase(logicalName)
        Case "NEWHIRE"
            fileName = "1263 ADP flexiform template_HK_NewHire.xlsx"
        Case "TERMINATION"
            fileName = "1263 ADP flexiform template_HK_Termination.xlsx"
        Case "DATACHANGE"
            fileName = "1263 ADP flexiform template_HK_DataChange.xlsx"
        Case "COMP"
            fileName = "1263 ADP flexiform template_HK_Comp.xlsx"
        Case "ATTENDANCE"
            fileName = "1263 ADP flexiform template_HK_Attendance.xlsx"
        Case "VARIABLE"
            fileName = "1263 ADP flexiform template_HK_Variable.xlsx"
        Case "ONETIMEPAYMENT"
            fileName = "One time payment report.xlsx"
        Case "INSPIREWARDS"
            fileName = "Inspire Awards payroll report.xlsx"
        Case "EMPLOYEELEAVE"
            fileName = "Employee_Leave_Transactions_Report.xlsx"
        Case "EAOSUMMARY"
            fileName = "EAO Summary Report_YYYYMM.xlsx"
        Case "WORKFORCEDETAIL"
            fileName = "Workforce Detail - Payroll-AP.xlsx"
        Case "MERCKPAYROLL"
            fileName = "Merck Payroll Summary Report����xxx.xlsx"
        Case "SIPQIP"
            fileName = "SIP QIP.xlsx"
        Case "FLEXCLAIM"
            fileName = "MSD HK Flex_Claim_Summary_Report.xlsx"
        Case "RSUGLOBAL"
            fileName = "RSU Dividend global report.xlsx"
        Case "RSUEY"
            fileName = "Dividend EY report.xlsx"
        Case "AIPPAYOUTS"
            fileName = "AIP Payouts Payroll Report.xlsx"
        Case "EXTRATABLE"
            fileName = "Additional table.xlsx"
        Case "PAYROLLREPORT"
            fileName = "Payroll Report.xlsx"
        Case "ALLOWANCEPLAN"
            fileName = "Allowance plan report.xlsx"
        Case "2025QXPAYOUT"
            fileName = "2025QX Payout Summary.xlsx"
        Case "OPTIONALMEDICAL"
            fileName = "Optional medical plan enrollment form.xlsx"
        Case Else
            fileName = logicalName
    End Select
    
    GetInputFilePath = G.RunParams.InputFolder & fileName
End Function


'------------------------------------------------------------------------------
' Sub: EnsureFolderExists
' Purpose: Create folder if it doesn't exist (supports nested paths)
' Parameters:
'   folderPath - Full path to the folder
'------------------------------------------------------------------------------
Public Sub EnsureFolderExists(folderPath As String)
    Dim fso As Object
    
    On Error GoTo ErrHandler
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then
        ' CreateFolder can handle nested paths
        fso.CreateFolder folderPath
        LogInfo "modConfigService", "EnsureFolderExists", "Created folder: " & folderPath
    End If
    
    Set fso = Nothing
    Exit Sub
    
ErrHandler:
    ' If CreateFolder fails for nested path, try building path recursively
    On Error Resume Next
    Dim parentPath As String
    parentPath = fso.GetParentFolderName(folderPath)
    
    If Len(parentPath) > 0 And Not fso.FolderExists(parentPath) Then
        EnsureFolderExists parentPath
    End If
    
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
    Set fso = Nothing
End Sub

