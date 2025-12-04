Attribute VB_Name = "modEAOService"
'==============================================================================
' Module: modEAOService
' Purpose: EAO (Employment Allowance Ordinance) calculation services
' Description: Handles EAO-related calculations for leave payments and adjustments
'==============================================================================
Option Explicit

' EAO record structure
Public Type tEAORecord
    wein As String
    AverageDayWage_12Month As Double
    DailySalary As Double
    DayWage_MaternityPaternity As Double
    Days_AnnualLeave As Double
    Days_AnnualLeave_LastMonth As Double
    Days_StatutoryHolidays As Double
    Days_MaternityLeave As Double
    Days_PaternityLeave As Double
    Days_SickLeave As Double
    Days_PPTO As Double
    Days_NoPayLeave As Double
    Days_NoPayLeave_LastMonth As Double
    NoPayLeaveCalculationBase As Double
    TotalWage_12Month As Double
    UntakenAnnualLeaveDays As Double
End Type

' Cache for EAO data
Private mEAOCache As Object ' Dictionary of WEIN -> tEAORecord

'------------------------------------------------------------------------------
' Sub: LoadEAOData
' Purpose: Load EAO Summary data into cache
'------------------------------------------------------------------------------
Public Sub LoadEAOData()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim filePath As String
    Dim lastRow As Long, i As Long
    Dim rec As tEAORecord
    Dim headers As Object
    
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
        rec.wein = GetCellValueByHeader(ws, i, headers, "WEIN")
        If rec.wein = "" Then rec.wein = GetCellValueByHeader(ws, i, headers, "WIN")
        If rec.wein = "" Then rec.wein = GetCellValueByHeader(ws, i, headers, "WEINEmployee ID")
        If rec.wein = "" Then rec.wein = GetCellValueByHeader(ws, i, headers, "Employee CodeWIN")
        If rec.wein = "" Then rec.wein = GetCellValueByHeader(ws, i, headers, "Employee ID")
        If rec.wein = "" Then rec.wein = GetCellValueByHeader(ws, i, headers, "EmployeeID")
        
        If rec.wein <> "" Then
            rec.AverageDayWage_12Month = ToDouble(GetCellValueByHeader(ws, i, headers, "AverageDayWage_12Month"))
            rec.DailySalary = ToDouble(GetCellValueByHeader(ws, i, headers, "DailySalary"))
            rec.DayWage_MaternityPaternity = ToDouble(GetCellValueByHeader(ws, i, headers, "DayWage_Maternity/Paternity/Sick Leave"))
            rec.Days_AnnualLeave = ToDouble(GetCellValueByHeader(ws, i, headers, "Days_AnnualLeave"))
            rec.Days_AnnualLeave_LastMonth = ToDouble(GetCellValueByHeader(ws, i, headers, "Days_AnnualLeave_LastMonth"))
            rec.Days_StatutoryHolidays = ToDouble(GetCellValueByHeader(ws, i, headers, "Days_StatutoryHolidays"))
            rec.Days_MaternityLeave = ToDouble(GetCellValueByHeader(ws, i, headers, "Days_MaternityLeave"))
            rec.Days_PaternityLeave = ToDouble(GetCellValueByHeader(ws, i, headers, "Days_PaternityLeave"))
            rec.Days_SickLeave = ToDouble(GetCellValueByHeader(ws, i, headers, "Days_SickLeave"))
            rec.Days_PPTO = ToDouble(GetCellValueByHeader(ws, i, headers, "Days_Paid Parental Time Off"))
            rec.Days_NoPayLeave = ToDouble(GetCellValueByHeader(ws, i, headers, "Days_NoPayLeave"))
            rec.Days_NoPayLeave_LastMonth = ToDouble(GetCellValueByHeader(ws, i, headers, "Days_NoPayLeave_LastMonth"))
            rec.NoPayLeaveCalculationBase = ToDouble(GetCellValueByHeader(ws, i, headers, "NoPayLeaveCalculationBase"))
            rec.TotalWage_12Month = ToDouble(GetCellValueByHeader(ws, i, headers, "TotalWage_12Month"))
            rec.UntakenAnnualLeaveDays = ToDouble(GetCellValueByHeader(ws, i, headers, "UntakenAnnualLeaveDays"))
            
            If Not mEAOCache.Exists(rec.wein) Then
                mEAOCache.Add rec.wein, rec
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
' Returns: tEAORecord (empty if not found)
'------------------------------------------------------------------------------
Public Function GetEAORecord(wein As String) As tEAORecord
    Dim emptyRec As tEAORecord
    
    If mEAOCache Is Nothing Then
        LoadEAOData
    End If
    
    If mEAOCache Is Nothing Then
        GetEAORecord = emptyRec
        Exit Function
    End If
    
    If mEAOCache.Exists(wein) Then
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
    Dim rec As tEAORecord
    Dim days As Double
    
    rec = GetEAORecord(wein)
    
    days = rec.Days_AnnualLeave_LastMonth + rec.Days_AnnualLeave + rec.Days_StatutoryHolidays
    CalcTotalEAOAdj = RoundAmount2((rec.AverageDayWage_12Month - rec.DailySalary) * days)
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
    Dim rec As tEAORecord
    
    rec = GetEAORecord(wein)
    CalcAnnualLeaveEAOAdj = RoundAmount2((rec.AverageDayWage_12Month - rec.DailySalary) * TotalDays)
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
    Dim rec As tEAORecord
    
    rec = GetEAORecord(wein)
    CalcSickLeaveEAOAdj = RoundAmount2((rec.DayWage_MaternityPaternity - rec.DailySalary) * days)
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
    Dim rec As tEAORecord
    
    rec = GetEAORecord(wein)
    CalcNoPayLeaveDeduction = RoundAmount2(rec.NoPayLeaveCalculationBase * (rec.Days_NoPayLeave + rec.Days_NoPayLeave_LastMonth))
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
    Dim rec As tEAORecord
    
    rec = GetEAORecord(wein)
    CalcMaternityLeavePayment = RoundAmount2(rec.DayWage_MaternityPaternity * rec.Days_MaternityLeave)
End Function

'------------------------------------------------------------------------------
' Function: CalcPaternityLeavePayment
' Purpose: Calculate Paternity Leave Payment
' Parameters:
'   wein - WEIN
' Returns: Payment amount
'------------------------------------------------------------------------------
Public Function CalcPaternityLeavePayment(wein As String) As Double
    Dim rec As tEAORecord
    
    rec = GetEAORecord(wein)
    CalcPaternityLeavePayment = RoundAmount2(rec.DayWage_MaternityPaternity * rec.Days_PaternityLeave)
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
    Dim rec As tEAORecord
    
    rec = GetEAORecord(wein)
    CalcSickLeavePayment = RoundAmount2(rec.DayWage_MaternityPaternity * rec.Days_SickLeave)
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
    Dim rec As tEAORecord
    Dim dayRate As Double
    
    rec = GetEAORecord(wein)
    dayRate = WorksheetFunction.Max(rec.DailySalary, rec.AverageDayWage_12Month * 0.8)
    CalcPPTOPayment = RoundAmount2(dayRate * rec.Days_PPTO)
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
    Dim rec As tEAORecord
    Dim dayRate As Double
    
    rec = GetEAORecord(wein)
    dayRate = WorksheetFunction.Max(monthlySalary / 22, rec.AverageDayWage_12Month)
    CalcUntakenAnnualLeavePayment = RoundAmount2(dayRate * rec.UntakenAnnualLeaveDays)
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
        If headerName <> "" And Not dict.Exists(headerName) Then
            dict.Add headerName, i
        End If
    Next i
    
    Set BuildHeaderIndex = dict
End Function

Private Function GetCellValueByHeader(ws As Worksheet, rowNum As Long, headers As Object, headerName As String) As String
    Dim col As Long
    
    GetCellValueByHeader = ""
    
    If headers.Exists(headerName) Then
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
