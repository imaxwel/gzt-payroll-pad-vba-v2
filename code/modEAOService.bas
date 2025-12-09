Attribute VB_Name = "modEAOService"
'==============================================================================
' Module: modEAOService
' Purpose: EAO (Employment Allowance Ordinance) calculation services
' Description: Handles EAO-related calculations for leave payments and adjustments
'==============================================================================
Option Explicit

' EAO record field indices (for array-based storage)
Private Const EAO_WEIN As Long = 0
Private Const EAO_AVG_DAY_WAGE As Long = 1
Private Const EAO_DAILY_SALARY As Long = 2
Private Const EAO_DAY_WAGE_MAT As Long = 3
Private Const EAO_DAYS_AL As Long = 4
Private Const EAO_DAYS_AL_LAST As Long = 5
Private Const EAO_DAYS_STAT As Long = 6
Private Const EAO_DAYS_MAT As Long = 7
Private Const EAO_DAYS_PAT As Long = 8
Private Const EAO_DAYS_SICK As Long = 9
Private Const EAO_DAYS_PPTO As Long = 10
Private Const EAO_DAYS_NPL As Long = 11
Private Const EAO_DAYS_NPL_LAST As Long = 12
Private Const EAO_NPL_BASE As Long = 13
Private Const EAO_TOTAL_WAGE As Long = 14
Private Const EAO_UNTAKEN_AL As Long = 15

' Cache for EAO data
Private mEAOCache As Object ' Dictionary of WEIN -> Array(0 to 15)

'------------------------------------------------------------------------------
' Sub: LoadEAOData
' Purpose: Load EAO Summary data into cache
'------------------------------------------------------------------------------
Public Sub LoadEAOData()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filePath As String
    Dim lastRow As Long, i As Long
    Dim rec(0 To 15) As Variant
    Dim headers As Object
    Dim wein As String
    
    On Error GoTo ErrHandler
    
    Set mEAOCache = CreateObject("Scripting.Dictionary")
    
    filePath = GetInputFilePath("EAOSummary")
    
    If Dir(filePath) = "" Then
        LogWarning "modEAOService", "LoadEAOData", "EAO Summary file not found: " & filePath
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Worksheets(1)
    
    ' Build header index
    Set headers = BuildHeaderIndex(ws.Rows(1))
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        ' Try multiple field name variants for WEIN
        wein = GetCellValueByHeader(ws, i, headers, "WEIN")
        If wein = "" Then wein = GetCellValueByHeader(ws, i, headers, "WIN")
        If wein = "" Then wein = GetCellValueByHeader(ws, i, headers, "WEINEmployee ID")
        If wein = "" Then wein = GetCellValueByHeader(ws, i, headers, "Employee CodeWIN")
        If wein = "" Then wein = GetCellValueByHeader(ws, i, headers, "Employee ID")
        If wein = "" Then wein = GetCellValueByHeader(ws, i, headers, "EmployeeID")
        
        If wein <> "" Then
            rec(EAO_WEIN) = wein
            rec(EAO_AVG_DAY_WAGE) = ToDouble(GetCellValueByHeader(ws, i, headers, "AverageDayWage_12Month"))
            rec(EAO_DAILY_SALARY) = ToDouble(GetCellValueByHeader(ws, i, headers, "DailySalary"))
            rec(EAO_DAY_WAGE_MAT) = ToDouble(GetCellValueByHeader(ws, i, headers, "DayWage_Maternity/Paternity/Sick Leave"))
            rec(EAO_DAYS_AL) = ToDouble(GetCellValueByHeader(ws, i, headers, "Days_AnnualLeave"))
            rec(EAO_DAYS_AL_LAST) = ToDouble(GetCellValueByHeader(ws, i, headers, "Days_AnnualLeave_LastMonth"))
            rec(EAO_DAYS_STAT) = ToDouble(GetCellValueByHeader(ws, i, headers, "Days_StatutoryHolidays"))
            rec(EAO_DAYS_MAT) = ToDouble(GetCellValueByHeader(ws, i, headers, "Days_MaternityLeave"))
            rec(EAO_DAYS_PAT) = ToDouble(GetCellValueByHeader(ws, i, headers, "Days_PaternityLeave"))
            rec(EAO_DAYS_SICK) = ToDouble(GetCellValueByHeader(ws, i, headers, "Days_SickLeave"))
            rec(EAO_DAYS_PPTO) = ToDouble(GetCellValueByHeader(ws, i, headers, "Days_Paid Parental Time Off"))
            rec(EAO_DAYS_NPL) = ToDouble(GetCellValueByHeader(ws, i, headers, "Days_NoPayLeave"))
            rec(EAO_DAYS_NPL_LAST) = ToDouble(GetCellValueByHeader(ws, i, headers, "Days_NoPayLeave_LastMonth"))
            rec(EAO_NPL_BASE) = ToDouble(GetCellValueByHeader(ws, i, headers, "NoPayLeaveCalculationBase"))
            rec(EAO_TOTAL_WAGE) = ToDouble(GetCellValueByHeader(ws, i, headers, "TotalWage_12Month"))
            rec(EAO_UNTAKEN_AL) = ToDouble(GetCellValueByHeader(ws, i, headers, "UntakenAnnualLeaveDays"))
            
            If Not mEAOCache.exists(wein) Then
                mEAOCache.Add wein, rec
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    
    LogInfo "modEAOService", "LoadEAOData", "Loaded " & mEAOCache.count & " EAO records"
    Exit Sub
    
ErrHandler:
    LogError "modEAOService", "LoadEAOData", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Function: GetEAORecord
' Purpose: Get EAO record for a WEIN
' Parameters:
'   wein - WEIN to look up
' Returns: Variant array (empty array if not found)
'------------------------------------------------------------------------------
Private Function GetEAORecord(wein As String) As Variant
    Dim emptyRec(0 To 15) As Variant
    Dim i As Long
    
    ' Initialize empty record
    For i = 0 To 15
        emptyRec(i) = 0
    Next i
    emptyRec(EAO_WEIN) = ""
    
    If mEAOCache Is Nothing Then
        LoadEAOData
    End If
    
    If mEAOCache Is Nothing Then
        GetEAORecord = emptyRec
        Exit Function
    End If
    
    If mEAOCache.exists(wein) Then
        GetEAORecord = mEAOCache(wein)
    Else
        GetEAORecord = emptyRec
    End If
End Function

'------------------------------------------------------------------------------
' Function: CalcTotalEAOAdj
' Purpose: Calculate Total EAO Adjustment
' Parameters:
'   wein - WEIN
' Returns: Total EAO Adjustment amount
' Formula: (AverageDayWage_12Month - DailySalary) * (Days_AnnualLeave_LastMonth + Days_AnnualLeave + Days_StatutoryHolidays)
'------------------------------------------------------------------------------
Public Function CalcTotalEAOAdj(wein As String) As Double
    Dim rec As Variant
    Dim days As Double
    
    rec = GetEAORecord(wein)
    
    days = rec(EAO_DAYS_AL_LAST) + rec(EAO_DAYS_AL) + rec(EAO_DAYS_STAT)
    CalcTotalEAOAdj = RoundAmount2((rec(EAO_AVG_DAY_WAGE) - rec(EAO_DAILY_SALARY)) * days)
End Function

'------------------------------------------------------------------------------
' Function: CalcAnnualLeaveEAOAdj
' Purpose: Calculate Annual Leave EAO Adjustment for older periods
' Parameters:
'   wein - WEIN
'   totalDays - Total days from periods before previous month
' Returns: EAO adjustment amount
' Formula: (AverageDayWage_12Month - DailySalary) * totalDays
'------------------------------------------------------------------------------
Public Function CalcAnnualLeaveEAOAdj(wein As String, TotalDays As Double) As Double
    Dim rec As Variant
    
    rec = GetEAORecord(wein)
    CalcAnnualLeaveEAOAdj = RoundAmount2((rec(EAO_AVG_DAY_WAGE) - rec(EAO_DAILY_SALARY)) * TotalDays)
End Function

'------------------------------------------------------------------------------
' Function: CalcSickLeaveEAOAdj
' Purpose: Calculate Sick Leave EAO Adjustment
' Parameters:
'   wein - WEIN
'   days - Sick leave days
' Returns: EAO adjustment amount
' Formula: (DayWage_Maternity/Paternity/Sick Leave - DailySalary) * Days_SickLeave
'------------------------------------------------------------------------------
Public Function CalcSickLeaveEAOAdj(wein As String, days As Double) As Double
    Dim rec As Variant
    
    rec = GetEAORecord(wein)
    CalcSickLeaveEAOAdj = RoundAmount2((rec(EAO_DAY_WAGE_MAT) - rec(EAO_DAILY_SALARY)) * days)
End Function

'------------------------------------------------------------------------------
' Function: CalcNoPayLeaveDeduction
' Purpose: Calculate No Pay Leave Deduction
' Parameters:
'   wein - WEIN
' Returns: Deduction amount
' Formula: NoPayLeaveCalculationBase * (Days_NoPayLeave + Days_NoPayLeave_LastMonth)
'------------------------------------------------------------------------------
Public Function CalcNoPayLeaveDeduction(wein As String) As Double
    Dim rec As Variant
    
    rec = GetEAORecord(wein)
    CalcNoPayLeaveDeduction = RoundAmount2(rec(EAO_NPL_BASE) * (rec(EAO_DAYS_NPL) + rec(EAO_DAYS_NPL_LAST)))
End Function

'------------------------------------------------------------------------------
' Function: CalcMaternityLeavePayment
' Purpose: Calculate Maternity Leave Payment
' Parameters:
'   wein - WEIN
' Returns: Payment amount
' Formula: DayWage_Maternity/Paternity/Sick Leave * Days_MaternityLeave
'------------------------------------------------------------------------------
Public Function CalcMaternityLeavePayment(wein As String) As Double
    Dim rec As Variant
    
    rec = GetEAORecord(wein)
    CalcMaternityLeavePayment = RoundAmount2(rec(EAO_DAY_WAGE_MAT) * rec(EAO_DAYS_MAT))
End Function

'------------------------------------------------------------------------------
' Function: CalcPaternityLeavePayment
' Purpose: Calculate Paternity Leave Payment
' Parameters:
'   wein - WEIN
' Returns: Payment amount
'------------------------------------------------------------------------------
Public Function CalcPaternityLeavePayment(wein As String) As Double
    Dim rec As Variant
    
    rec = GetEAORecord(wein)
    CalcPaternityLeavePayment = RoundAmount2(rec(EAO_DAY_WAGE_MAT) * rec(EAO_DAYS_PAT))
End Function

'------------------------------------------------------------------------------
' Function: CalcSickLeavePayment
' Purpose: Calculate Sick Leave Payment
' Parameters:
'   wein - WEIN
' Returns: Payment amount
' Formula: DayWage_Maternity/Paternity/Sick Leave * Days_SickLeave
'------------------------------------------------------------------------------
Public Function CalcSickLeavePayment(wein As String) As Double
    Dim rec As Variant
    
    rec = GetEAORecord(wein)
    CalcSickLeavePayment = RoundAmount2(rec(EAO_DAY_WAGE_MAT) * rec(EAO_DAYS_SICK))
End Function

'------------------------------------------------------------------------------
' Function: CalcPPTOPayment
' Purpose: Calculate PPTO (Paid Parental Time Off) Payment
' Parameters:
'   wein - WEIN
' Returns: Payment amount
' Formula: Max(DailySalary, AverageDayWage_12Month * 80%) * Days_PPTO
'------------------------------------------------------------------------------
Public Function CalcPPTOPayment(wein As String) As Double
    Dim rec As Variant
    Dim dayRate As Double
    
    rec = GetEAORecord(wein)
    dayRate = WorksheetFunction.Max(rec(EAO_DAILY_SALARY), rec(EAO_AVG_DAY_WAGE) * 0.8)
    CalcPPTOPayment = RoundAmount2(dayRate * rec(EAO_DAYS_PPTO))
End Function

'------------------------------------------------------------------------------
' Function: CalcUntakenAnnualLeavePayment
' Purpose: Calculate Untaken Annual Leave Payment
' Parameters:
'   wein - WEIN
'   monthlySalary - Monthly salary (already rounded to integer)
' Returns: Payment amount
' Formula: Max(MonthlySalary/22, AverageDayWage_12Month) * UntakenAnnualLeaveDays
'------------------------------------------------------------------------------
Public Function CalcUntakenAnnualLeavePayment(wein As String, monthlySalary As Double) As Double
    Dim rec As Variant
    Dim dayRate As Double
    
    rec = GetEAORecord(wein)
    dayRate = WorksheetFunction.Max(monthlySalary / 22, rec(EAO_AVG_DAY_WAGE))
    CalcUntakenAnnualLeavePayment = RoundAmount2(dayRate * rec(EAO_UNTAKEN_AL))
End Function

'------------------------------------------------------------------------------
' Helper Functions
'------------------------------------------------------------------------------

Private Function BuildHeaderIndex(headerRow As Range) As Object
    Dim dict As Object
    Dim i As Long
    Dim headerName As String
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To headerRow.Columns.count
        headerName = Trim(CStr(Nz(headerRow.Cells(1, i).Value, "")))
        If headerName <> "" And Not dict.exists(headerName) Then
            dict.Add headerName, i
        End If
    Next i
    
    Set BuildHeaderIndex = dict
End Function

Private Function GetCellValueByHeader(ws As Worksheet, rowNum As Long, headers As Object, headerName As String) As String
    Dim col As Long
    
    GetCellValueByHeader = ""
    
    If headers.exists(headerName) Then
        col = headers(headerName)
        GetCellValueByHeader = Trim(CStr(Nz(ws.Cells(rowNum, col).Value, "")))
    End If
End Function

'------------------------------------------------------------------------------
' Sub: ClearEAOCache
' Purpose: Clear the EAO cache
'------------------------------------------------------------------------------
Public Sub ClearEAOCache()
    Set mEAOCache = Nothing
End Sub
