Attribute VB_Name = "modCalendarService"
'==============================================================================
' Module: modCalendarService
' Purpose: Date and calendar services for HK payroll processing
' Description: Handles business days, HK public holidays, and cross-month splitting
'==============================================================================
Option Explicit

' Cache for HK public holidays
Private mHolidayCache As Object ' Scripting.Dictionary

'------------------------------------------------------------------------------
' Function: IsWeekend
' Purpose: Check if a date falls on weekend (Saturday or Sunday)
' Parameters:
'   d - Date to check
' Returns: True if weekend, False otherwise
'------------------------------------------------------------------------------
Public Function IsWeekend(d As Date) As Boolean
    IsWeekend = (Weekday(d, vbMonday) > 5)
End Function

'------------------------------------------------------------------------------
' Function: IsHKPublicHoliday
' Purpose: Check if a date is a Hong Kong public holiday
' Parameters:
'   d - Date to check
' Returns: True if HK public holiday, False otherwise
'------------------------------------------------------------------------------
Public Function IsHKPublicHoliday(d As Date) As Boolean
    On Error GoTo ErrHandler
    
    ' Initialize cache if needed
    If mHolidayCache Is Nothing Then
        LoadHolidayCache
    End If
    
    If mHolidayCache Is Nothing Then
        IsHKPublicHoliday = False
        Exit Function
    End If
    
    IsHKPublicHoliday = mHolidayCache.exists(CLng(d))
    Exit Function
    
ErrHandler:
    IsHKPublicHoliday = False
End Function

'------------------------------------------------------------------------------
' Sub: LoadHolidayCache
' Purpose: Load HK public holidays from config into cache
'------------------------------------------------------------------------------
Private Sub LoadHolidayCache()
    Dim configWb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim holidayDate As Date
    
    On Error GoTo ErrHandler
    
    Set mHolidayCache = CreateObject("Scripting.Dictionary")
    
    Set configWb = OpenConfigWorkbook()
    If configWb Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set ws = configWb.Worksheets("Calendar")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    For i = 2 To lastRow
        If IsDate(ws.Cells(i, 1).value) Then
            ' Check if IsHKHoliday column (column 2) is True
            If CBool(Nz(ws.Cells(i, 2).value, False)) Then
                holidayDate = CDate(ws.Cells(i, 1).value)
                If Not mHolidayCache.exists(CLng(holidayDate)) Then
                    mHolidayCache.Add CLng(holidayDate), True
                End If
            End If
        End If
    Next i
    
    Exit Sub
    
ErrHandler:
    Set mHolidayCache = CreateObject("Scripting.Dictionary")
End Sub

'------------------------------------------------------------------------------
' Function: IsBusinessDay
' Purpose: Check if a date is a business day (not weekend, not HK holiday)
' Parameters:
'   d - Date to check
' Returns: True if business day, False otherwise
'------------------------------------------------------------------------------
Public Function IsBusinessDay(d As Date) As Boolean
    IsBusinessDay = Not IsWeekend(d) And Not IsHKPublicHoliday(d)
End Function

'------------------------------------------------------------------------------
' Function: CountBusinessDays
' Purpose: Count business days between two dates (inclusive)
' Parameters:
'   startDate - Start date
'   endDate - End date
' Returns: Number of business days
'------------------------------------------------------------------------------
Public Function CountBusinessDays(startDate As Date, endDate As Date) As Long
    Dim d As Date
    Dim count As Long
    
    count = 0
    For d = startDate To endDate
        If IsBusinessDay(d) Then
            count = count + 1
        End If
    Next d
    
    CountBusinessDays = count
End Function

'------------------------------------------------------------------------------
' Function: CountCalendarDays
' Purpose: Count calendar days between two dates (inclusive)
' Parameters:
'   startDate - Start date
'   endDate - End date
' Returns: Number of calendar days
'------------------------------------------------------------------------------
Public Function CountCalendarDays(startDate As Date, endDate As Date) As Long
    CountCalendarDays = endDate - startDate + 1
End Function

'------------------------------------------------------------------------------
' Sub: SplitByCalendarMonth
' Purpose: Split a date range into segments by calendar month
' Parameters:
'   startDate - Start date of the range
'   endDate - End date of the range
'   spans - Collection to receive tDateSpan objects (ByRef)
'------------------------------------------------------------------------------
Public Sub SplitByCalendarMonth(startDate As Date, endDate As Date, ByRef spans As Collection)
    Dim currentStart As Date
    Dim currentEnd As Date
    Dim span As tDateSpan
    
    Set spans = New Collection
    
    If startDate > endDate Then Exit Sub
    
    currentStart = startDate
    
    Do While currentStart <= endDate
        ' End of current month
        currentEnd = DateSerial(Year(currentStart), Month(currentStart) + 1, 0)
        
        ' Don't go past the end date
        If currentEnd > endDate Then currentEnd = endDate
        
        ' Create span
        span.startDate = currentStart
        span.endDate = currentEnd
        span.YearMonth = Format(currentStart, "yyyymm")
        span.days = CountCalendarDays(currentStart, currentEnd)
        
        spans.Add span
        
        ' Move to first day of next month
        currentStart = DateSerial(Year(currentEnd), Month(currentEnd) + 1, 1)
    Loop
End Sub

'------------------------------------------------------------------------------
' Sub: SplitByCalendarMonthWithBusinessDays
' Purpose: Split a date range by calendar month, counting business days
' Parameters:
'   startDate - Start date of the range
'   endDate - End date of the range
'   spans - Collection to receive tDateSpan objects (ByRef)
' Note: span.Days will contain business days count
'------------------------------------------------------------------------------
Public Sub SplitByCalendarMonthWithBusinessDays(startDate As Date, endDate As Date, ByRef spans As Collection)
    Dim currentStart As Date
    Dim currentEnd As Date
    Dim span As tDateSpan
    
    Set spans = New Collection
    
    If startDate > endDate Then Exit Sub
    
    currentStart = startDate
    
    Do While currentStart <= endDate
        ' End of current month
        currentEnd = DateSerial(Year(currentStart), Month(currentStart) + 1, 0)
        
        ' Don't go past the end date
        If currentEnd > endDate Then currentEnd = endDate
        
        ' Create span with business days
        span.startDate = currentStart
        span.endDate = currentEnd
        span.YearMonth = Format(currentStart, "yyyymm")
        span.days = CountBusinessDays(currentStart, currentEnd)
        
        spans.Add span
        
        ' Move to first day of next month
        currentStart = DateSerial(Year(currentEnd), Month(currentEnd) + 1, 1)
    Loop
End Sub

'------------------------------------------------------------------------------
' Function: HasFourConsecutiveBusinessDays
' Purpose: Check if a date range contains at least 4 consecutive business days
' Parameters:
'   startDate - Start date
'   endDate - End date
' Returns: True if 4+ consecutive business days exist
' Note: Used for Sick Leave eligibility rule
'------------------------------------------------------------------------------
Public Function HasFourConsecutiveBusinessDays(startDate As Date, endDate As Date) As Boolean
    Dim d As Date
    Dim consecutiveCount As Long
    
    consecutiveCount = 0
    
    For d = startDate To endDate
        If IsBusinessDay(d) Then
            consecutiveCount = consecutiveCount + 1
            If consecutiveCount >= 4 Then
                HasFourConsecutiveBusinessDays = True
                Exit Function
            End If
        Else
            consecutiveCount = 0
        End If
    Next d
    
    HasFourConsecutiveBusinessDays = False
End Function

'------------------------------------------------------------------------------
' Function: GetMonthYearString
' Purpose: Get YYYYMM string from a date
' Parameters:
'   d - Date
' Returns: String in "YYYYMM" format
'------------------------------------------------------------------------------
Public Function GetMonthYearString(d As Date) As String
    GetMonthYearString = Format(d, "yyyymm")
End Function

'------------------------------------------------------------------------------
' Function: GetBusinessDaysInMonth
' Purpose: Get total business days in a given month
' Parameters:
'   yr - Year
'   mo - Month
' Returns: Number of business days
'------------------------------------------------------------------------------
Public Function GetBusinessDaysInMonth(yr As Integer, mo As Integer) As Long
    Dim monthStart As Date
    Dim monthEnd As Date
    
    monthStart = DateSerial(yr, mo, 1)
    monthEnd = DateSerial(yr, mo + 1, 0)
    
    GetBusinessDaysInMonth = CountBusinessDays(monthStart, monthEnd)
End Function

'------------------------------------------------------------------------------
' Function: GetNextBusinessDay
' Purpose: Get the next business day from a given date
' Parameters:
'   d - Starting date
' Returns: Next business day
'------------------------------------------------------------------------------
Public Function GetNextBusinessDay(d As Date) As Date
    Dim nextDay As Date
    nextDay = d + 1
    
    Do While Not IsBusinessDay(nextDay)
        nextDay = nextDay + 1
    Loop
    
    GetNextBusinessDay = nextDay
End Function

'------------------------------------------------------------------------------
' Function: AdjustToBusinessDay
' Purpose: Adjust a date to the nearest business day (forward)
' Parameters:
'   d - Date to adjust
' Returns: Same date if business day, otherwise next business day
'------------------------------------------------------------------------------
Public Function AdjustToBusinessDay(d As Date) As Date
    If IsBusinessDay(d) Then
        AdjustToBusinessDay = d
    Else
        AdjustToBusinessDay = GetNextBusinessDay(d)
    End If
End Function

'------------------------------------------------------------------------------
' Sub: ClearHolidayCache
' Purpose: Clear the holiday cache (call when config changes)
'------------------------------------------------------------------------------
Public Sub ClearHolidayCache()
    Set mHolidayCache = Nothing
End Sub

'------------------------------------------------------------------------------
' Sub: CalcDaysByMonth
' Purpose: Calculate days split by month for a date range (calendar days)
' Parameters:
'   startDate - Start date of the range
'   endDate - End date of the range
'   targetYM - Target year-month string "YYYYMM"
'   prevYM - Previous year-month string "YYYYMM"
'   currentDays - (ByRef) Days in target month
'   prevDays - (ByRef) Days in previous month
'   olderDays - (ByRef) Days in older months
'------------------------------------------------------------------------------
Public Sub CalcDaysByMonth(startDate As Date, endDate As Date, _
                           targetYM As String, prevYM As String, _
                           ByRef currentDays As Double, ByRef prevDays As Double, _
                           ByRef olderDays As Double)
    Dim currentStart As Date, currentEnd As Date
    Dim spanYM As String
    Dim spanDays As Double
    
    currentDays = 0
    prevDays = 0
    olderDays = 0
    
    If startDate > endDate Then Exit Sub
    
    currentStart = startDate
    
    Do While currentStart <= endDate
        currentEnd = DateSerial(Year(currentStart), Month(currentStart) + 1, 0)
        If currentEnd > endDate Then currentEnd = endDate
        
        spanYM = Format(currentStart, "YYYYMM")
        spanDays = CountCalendarDays(currentStart, currentEnd)
        
        If spanYM = targetYM Then
            currentDays = currentDays + spanDays
        ElseIf spanYM = prevYM Then
            prevDays = prevDays + spanDays
        Else
            olderDays = olderDays + spanDays
        End If
        
        currentStart = DateSerial(Year(currentEnd), Month(currentEnd) + 1, 1)
    Loop
End Sub

'------------------------------------------------------------------------------
' Sub: CalcBusinessDaysByMonth
' Purpose: Calculate business days split by month for a date range
' Parameters:
'   startDate - Start date of the range
'   endDate - End date of the range
'   targetYM - Target year-month string "YYYYMM"
'   prevYM - Previous year-month string "YYYYMM"
'   currentDays - (ByRef) Business days in target month
'   prevDays - (ByRef) Business days in previous month
'   olderDays - (ByRef) Business days in older months
'------------------------------------------------------------------------------
Public Sub CalcBusinessDaysByMonth(startDate As Date, endDate As Date, _
                                   targetYM As String, prevYM As String, _
                                   ByRef currentDays As Double, ByRef prevDays As Double, _
                                   ByRef olderDays As Double)
    Dim currentStart As Date, currentEnd As Date
    Dim spanYM As String
    Dim spanDays As Double
    
    currentDays = 0
    prevDays = 0
    olderDays = 0
    
    If startDate > endDate Then Exit Sub
    
    currentStart = startDate
    
    Do While currentStart <= endDate
        currentEnd = DateSerial(Year(currentStart), Month(currentStart) + 1, 0)
        If currentEnd > endDate Then currentEnd = endDate
        
        spanYM = Format(currentStart, "YYYYMM")
        spanDays = CountBusinessDays(currentStart, currentEnd)
        
        If spanYM = targetYM Then
            currentDays = currentDays + spanDays
        ElseIf spanYM = prevYM Then
            prevDays = prevDays + spanDays
        Else
            olderDays = olderDays + spanDays
        End If
        
        currentStart = DateSerial(Year(currentEnd), Month(currentEnd) + 1, 1)
    Loop
End Sub
