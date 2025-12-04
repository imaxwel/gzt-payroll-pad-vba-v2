Attribute VB_Name = "modAggregationService"
'==============================================================================
' Module: modAggregationService
' Purpose: Grouping and aggregation services
' Description: Provides reusable functions for grouping data by employee and type,
'              ensuring "only one amount per employee per plan type" rule
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Function: GroupByEmployeeAndType
' Purpose: Group data by employee and type, summing amounts
' Parameters:
'   dataRange - Range containing data (including header row)
'   employeeColName - Name of the employee ID column
'   typeColName - Name of the type/plan column
'   amountColName - Name of the amount column
' Returns: Scripting.Dictionary with key = "employeeID|typeValue", value = summed amount
' Note: Uses header-based column detection, not hard-coded indices
'------------------------------------------------------------------------------
Public Function GroupByEmployeeAndType( _
    dataRange As Range, _
    employeeColName As String, _
    typeColName As String, _
    amountColName As String _
) As Object
    
    Dim dict As Object
    Dim empCol As Long, typeCol As Long, amtCol As Long
    Dim headerRow As Range
    Dim i As Long, lastRow As Long
    Dim empId As String, typeVal As String, key As String
    Dim amt As Double
    
    On Error GoTo ErrHandler
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    If dataRange Is Nothing Then
        Set GroupByEmployeeAndType = dict
        Exit Function
    End If
    
    ' Find column indices by header names
    Set headerRow = dataRange.Rows(1)
    empCol = FindColumnByHeader(headerRow, employeeColName)
    typeCol = FindColumnByHeader(headerRow, typeColName)
    amtCol = FindColumnByHeader(headerRow, amountColName)
    
    If empCol = 0 Or typeCol = 0 Or amtCol = 0 Then
        LogError "modAggregationService", "GroupByEmployeeAndType", 0, _
            "Column not found: Emp=" & employeeColName & ", Type=" & typeColName & ", Amt=" & amountColName
        Set GroupByEmployeeAndType = dict
        Exit Function
    End If
    
    ' Process data rows
    lastRow = dataRange.Rows.count
    
    For i = 2 To lastRow
        empId = Trim(CStr(Nz(dataRange.Cells(i, empCol).Value, "")))
        typeVal = Trim(CStr(Nz(dataRange.Cells(i, typeCol).Value, "")))
        amt = ToDouble(dataRange.Cells(i, amtCol).Value)
        
        If empId <> "" Then
            key = empId & "|" & typeVal
            
            If dict.Exists(key) Then
                dict(key) = dict(key) + amt
            Else
                dict.Add key, amt
            End If
        End If
    Next i
    
    ' Round all final values to 2 decimals
    Dim k As Variant
    For Each k In dict.Keys
        dict(k) = RoundAmount2(dict(k))
    Next k
    
    Set GroupByEmployeeAndType = dict
    Exit Function
    
ErrHandler:
    LogError "modAggregationService", "GroupByEmployeeAndType", Err.Number, Err.Description
    Set GroupByEmployeeAndType = CreateObject("Scripting.Dictionary")
End Function

'------------------------------------------------------------------------------
' Function: SumPerEmployee
' Purpose: Sum amounts per employee (without type grouping)
' Parameters:
'   dataRange - Range containing data (including header row)
'   employeeColName - Name of the employee ID column
'   amountColName - Name of the amount column
' Returns: Scripting.Dictionary with key = employeeID, value = summed amount
'------------------------------------------------------------------------------
Public Function SumPerEmployee( _
    dataRange As Range, _
    employeeColName As String, _
    amountColName As String _
) As Object
    
    Dim dict As Object
    Dim empCol As Long, amtCol As Long
    Dim headerRow As Range
    Dim i As Long, lastRow As Long
    Dim empId As String
    Dim amt As Double
    
    On Error GoTo ErrHandler
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    If dataRange Is Nothing Then
        Set SumPerEmployee = dict
        Exit Function
    End If
    
    ' Find column indices by header names
    Set headerRow = dataRange.Rows(1)
    empCol = FindColumnByHeader(headerRow, employeeColName)
    amtCol = FindColumnByHeader(headerRow, amountColName)
    
    If empCol = 0 Or amtCol = 0 Then
        LogError "modAggregationService", "SumPerEmployee", 0, _
            "Column not found: Emp=" & employeeColName & ", Amt=" & amountColName
        Set SumPerEmployee = dict
        Exit Function
    End If
    
    ' Process data rows
    lastRow = dataRange.Rows.count
    
    For i = 2 To lastRow
        empId = Trim(CStr(Nz(dataRange.Cells(i, empCol).Value, "")))
        amt = ToDouble(dataRange.Cells(i, amtCol).Value)
        
        If empId <> "" Then
            If dict.Exists(empId) Then
                dict(empId) = dict(empId) + amt
            Else
                dict.Add empId, amt
            End If
        End If
    Next i
    
    ' Round all final values to 2 decimals
    Dim k As Variant
    For Each k In dict.Keys
        dict(k) = RoundAmount2(dict(k))
    Next k
    
    Set SumPerEmployee = dict
    Exit Function
    
ErrHandler:
    LogError "modAggregationService", "SumPerEmployee", Err.Number, Err.Description
    Set SumPerEmployee = CreateObject("Scripting.Dictionary")
End Function

'------------------------------------------------------------------------------
' Function: GroupByEmployeeAndTypeFiltered
' Purpose: Group data with additional filter condition
' Parameters:
'   dataRange - Range containing data (including header row)
'   employeeColName - Name of the employee ID column
'   typeColName - Name of the type/plan column
'   amountColName - Name of the amount column
'   filterColName - Name of the column to filter on
'   filterValues - Array of values to include (case-insensitive)
' Returns: Scripting.Dictionary with key = "employeeID|typeValue", value = summed amount
'------------------------------------------------------------------------------
Public Function GroupByEmployeeAndTypeFiltered( _
    dataRange As Range, _
    employeeColName As String, _
    typeColName As String, _
    amountColName As String, _
    filterColName As String, _
    filterValues As Variant _
) As Object
    
    Dim dict As Object
    Dim empCol As Long, typeCol As Long, amtCol As Long, filterCol As Long
    Dim headerRow As Range
    Dim i As Long, lastRow As Long
    Dim empId As String, typeVal As String, key As String, filterVal As String
    Dim amt As Double
    Dim includeRow As Boolean
    Dim fv As Variant
    
    On Error GoTo ErrHandler
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    If dataRange Is Nothing Then
        Set GroupByEmployeeAndTypeFiltered = dict
        Exit Function
    End If
    
    ' Find column indices by header names
    Set headerRow = dataRange.Rows(1)
    empCol = FindColumnByHeader(headerRow, employeeColName)
    typeCol = FindColumnByHeader(headerRow, typeColName)
    amtCol = FindColumnByHeader(headerRow, amountColName)
    filterCol = FindColumnByHeader(headerRow, filterColName)
    
    If empCol = 0 Or typeCol = 0 Or amtCol = 0 Or filterCol = 0 Then
        Set GroupByEmployeeAndTypeFiltered = dict
        Exit Function
    End If
    
    ' Process data rows
    lastRow = dataRange.Rows.count
    
    For i = 2 To lastRow
        filterVal = UCase(Trim(CStr(Nz(dataRange.Cells(i, filterCol).Value, ""))))
        
        ' Check if row should be included
        includeRow = False
        If IsArray(filterValues) Then
            For Each fv In filterValues
                If filterVal = UCase(Trim(CStr(fv))) Then
                    includeRow = True
                    Exit For
                End If
            Next fv
        Else
            includeRow = (filterVal = UCase(Trim(CStr(filterValues))))
        End If
        
        If includeRow Then
            empId = Trim(CStr(Nz(dataRange.Cells(i, empCol).Value, "")))
            typeVal = Trim(CStr(Nz(dataRange.Cells(i, typeCol).Value, "")))
            amt = ToDouble(dataRange.Cells(i, amtCol).Value)
            
            If empId <> "" Then
                key = empId & "|" & typeVal
                
                If dict.Exists(key) Then
                    dict(key) = dict(key) + amt
                Else
                    dict.Add key, amt
                End If
            End If
        End If
    Next i
    
    ' Round all final values to 2 decimals
    Dim k As Variant
    For Each k In dict.Keys
        dict(k) = RoundAmount2(dict(k))
    Next k
    
    Set GroupByEmployeeAndTypeFiltered = dict
    Exit Function
    
ErrHandler:
    LogError "modAggregationService", "GroupByEmployeeAndTypeFiltered", Err.Number, Err.Description
    Set GroupByEmployeeAndTypeFiltered = CreateObject("Scripting.Dictionary")
End Function

'------------------------------------------------------------------------------
' Function: FindColumnByHeader
' Purpose: Find column index by header name (case-insensitive)
' Parameters:
'   headerRow - Range containing header row
'   headerName - Name to search for (can be comma-separated list of variants)
' Returns: Column index (1-based) or 0 if not found
'------------------------------------------------------------------------------
Public Function FindColumnByHeader(headerRow As Range, headerName As String) As Long
    Dim i As Long
    Dim cellValue As String
    Dim searchName As String
    Dim variants() As String
    Dim v As Long
    
    FindColumnByHeader = 0
    
    ' Support comma-separated variant names
    variants = Split(headerName, ",")
    
    For i = 1 To headerRow.Columns.count
        cellValue = UCase(Trim(CStr(Nz(headerRow.Cells(1, i).Value, ""))))
        
        For v = LBound(variants) To UBound(variants)
            searchName = UCase(Trim(variants(v)))
            If cellValue = searchName Then
                FindColumnByHeader = i
                Exit Function
            End If
        Next v
    Next i
End Function

'------------------------------------------------------------------------------
' Function: FindEmployeeColumn
' Purpose: Find any employee ID column using all known variants
' Parameters:
'   headerRow - Range containing header row
' Returns: Column index (1-based) or 0 if not found
'------------------------------------------------------------------------------
Public Function FindEmployeeColumn(headerRow As Range) As Long
    FindEmployeeColumn = FindColumnByHeader(headerRow, _
        "WEIN,WIN,Employee ID,EmployeeID,Employee Code,EmployeeCode," & _
        "Employee Number ID,Employee Reference,EmployeeNumber,Employee Number," & _
        "WEINEmployee ID,Employee CodeWIN")
End Function

'------------------------------------------------------------------------------
' Function: GetValueFromGroupedDict
' Purpose: Get value from grouped dictionary by employee and type
' Parameters:
'   dict - Dictionary from GroupByEmployeeAndType
'   employeeId - Employee ID
'   typeValue - Type/plan value
' Returns: Amount or 0 if not found
'------------------------------------------------------------------------------
Public Function GetValueFromGroupedDict( _
    dict As Object, _
    employeeId As String, _
    typeValue As String _
) As Double
    
    Dim key As String
    
    GetValueFromGroupedDict = 0
    
    If dict Is Nothing Then Exit Function
    
    key = employeeId & "|" & typeValue
    
    If dict.Exists(key) Then
        GetValueFromGroupedDict = dict(key)
    End If
End Function

'------------------------------------------------------------------------------
' Function: BuildEmployeeIndex
' Purpose: Build a dictionary mapping employee ID to row number
' Parameters:
'   ws - Worksheet containing data
'   employeeColName - Name of the employee ID column
'   headerRow - Row number of header (default 1)
' Returns: Dictionary with key = employeeID, value = row number
'------------------------------------------------------------------------------
Public Function BuildEmployeeIndex( _
    ws As Worksheet, _
    employeeColName As String, _
    Optional headerRow As Long = 1 _
) As Object
    
    Dim dict As Object
    Dim empCol As Long
    Dim i As Long, lastRow As Long
    Dim empId As String
    
    On Error GoTo ErrHandler
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Find employee column
    empCol = FindColumnByHeader(ws.Rows(headerRow), employeeColName)
    
    If empCol = 0 Then
        Set BuildEmployeeIndex = dict
        Exit Function
    End If
    
    lastRow = ws.Cells(ws.Rows.count, empCol).End(xlUp).row
    
    For i = headerRow + 1 To lastRow
        empId = Trim(CStr(Nz(ws.Cells(i, empCol).Value, "")))
        If empId <> "" And Not dict.Exists(empId) Then
            dict.Add empId, i
        End If
    Next i
    
    Set BuildEmployeeIndex = dict
    Exit Function
    
ErrHandler:
    LogError "modAggregationService", "BuildEmployeeIndex", Err.Number, Err.Description
    Set BuildEmployeeIndex = CreateObject("Scripting.Dictionary")
End Function
