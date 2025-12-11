Attribute VB_Name = "modEmployeeMappingService"
'==============================================================================
' Modulex: modEmployeeMappingService
' Purpose: Employeer ID mapping services
' Description: Handles WEIN <-> Employee ID <-> Employee Code mappings
' Note: Different systems use different field names for the same employee ID:
'       WEIN, WIN, Employee ID, Employee Code, Employee Number ID,
'       Employee Reference, EmployeeNumber, etc.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Employee ID Field Name Variants
' These are all equivalent employee identifier fields across different systems
'------------------------------------------------------------------------------
Public Const EMP_ID_VARIANTS As String = "Employee ID,EmployeeID,Employee Number ID,EMPLOYEE ID"
Public Const WEIN_VARIANTS As String = "WEIN,WIN,WEINEmployee ID,Employee CodeWIN"
Public Const EMP_CODE_VARIANTS As String = "Employee Code,EmployeeCode,Employee Reference,EmployeeNumber,Employee Number"
Public Const ALL_EMP_ID_VARIANTS As String = "WEIN,WIN,Employee ID,EmployeeID,Employee Code,EmployeeCode,Employee Number ID,Employee Reference,EmployeeNumber,Employee Number,WEINEmployee ID,Employee CodeWIN"

'------------------------------------------------------------------------------
' Sub: BuildEmployeeMappings
' Purpose: Build all employee mapping dictionaries from Workforce Detail
'------------------------------------------------------------------------------
Public Sub BuildEmployeeMappings()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim empId As String, wein As String, empCode As String
    Dim filePath As String
    Dim empIdCol As Long, weinCol As Long, empCodeCol As Long
    
    On Error GoTo ErrHandler
    
    ' Initialize dictionaries
    Set G.DictWeinToEmpId = CreateObject("Scripting.Dictionary")
    Set G.DictEmpIdToWein = CreateObject("Scripting.Dictionary")
    Set G.DictEmpCodeToWein = CreateObject("Scripting.Dictionary")
    Set G.DictWeinToEmpCode = CreateObject("Scripting.Dictionary")
    
    ' Open Workforce Detail
    filePath = GetInputFilePath("WorkforceDetail")
    
    If Dir(filePath) = "" Then
        LogError "modEmployeeMappingService", "BuildEmployeeMappings", 0, _
            "Workforce Detail file not found: " & filePath
        Exit Sub
    End If
    
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set ws = wb.Worksheets(1)
    
    ' Find the header row (look for "Employee ID" in first 20 rows)
    Dim headerRow As Long
    headerRow = FindHeaderRow(ws, "Employee ID")
    If headerRow = 0 Then headerRow = 1 ' Default to row 1 if not found
    
    ' Find columns using variant names
    empIdCol = FindEmployeeIdColumn(ws.Rows(headerRow), EMP_ID_VARIANTS)
    weinCol = FindEmployeeIdColumn(ws.Rows(headerRow), WEIN_VARIANTS)
    empCodeCol = FindEmployeeIdColumn(ws.Rows(headerRow), EMP_CODE_VARIANTS)
    
    If empIdCol = 0 And weinCol = 0 Then
        LogError "modEmployeeMappingService", "BuildEmployeeMappings", 0, _
            "Required columns not found in Workforce Detail (searched in row " & headerRow & ")"
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, IIf(empIdCol > 0, empIdCol, weinCol)).End(xlUp).row
    
    For i = headerRow + 1 To lastRow
        If empIdCol > 0 Then empId = Trim(CStr(Nz(ws.Cells(i, empIdCol).value, "")))
        If weinCol > 0 Then wein = Trim(CStr(Nz(ws.Cells(i, weinCol).value, "")))
        If empCodeCol > 0 Then empCode = Trim(CStr(Nz(ws.Cells(i, empCodeCol).value, "")))
        
        ' Build WEIN <-> Employee ID mappings
        If wein <> "" And empId <> "" Then
            If Not G.DictWeinToEmpId.exists(wein) Then
                G.DictWeinToEmpId.Add wein, empId
            End If
            If Not G.DictEmpIdToWein.exists(empId) Then
                G.DictEmpIdToWein.Add empId, wein
            End If
        End If
        
        ' Build Employee Code <-> WEIN mappings
        If empCode <> "" And wein <> "" Then
            If Not G.DictEmpCodeToWein.exists(empCode) Then
                G.DictEmpCodeToWein.Add empCode, wein
            End If
            If Not G.DictWeinToEmpCode.exists(wein) Then
                G.DictWeinToEmpCode.Add wein, empCode
            End If
        End If
    Next i
    
    wb.Close SaveChanges:=False
    Debug.Print TypeName(G.Payroll.payDate), G.Payroll.payDate
    LogInfo "modEmployeeMappingService", "BuildEmployeeMappings", _
        "Loaded " & G.DictWeinToEmpId.count & " WEIN mappings"
    
    Exit Sub
    
ErrHandler:
    LogError "modEmployeeMappingService", "BuildEmployeeMappings", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Function: WeinFromEmpId
' Purpose: Get WEIN from Employee ID
' Parameters:
'   empId - Employee ID
' Returns: WEIN or empty string if not found
'------------------------------------------------------------------------------
Public Function WeinFromEmpId(empId As String) As String
    WeinFromEmpId = ""
    
    If G.DictEmpIdToWein Is Nothing Then Exit Function
    
    If G.DictEmpIdToWein.exists(empId) Then
        WeinFromEmpId = G.DictEmpIdToWein(empId)
    End If
End Function

'------------------------------------------------------------------------------
' Function: EmpIdFromWein
' Purpose: Get Employee ID from WEIN
' Parameters:
'   wein - WEIN
' Returns: Employee ID or empty string if not found
'------------------------------------------------------------------------------
Public Function EmpIdFromWein(wein As String) As String
    EmpIdFromWein = ""
    
    If G.DictWeinToEmpId Is Nothing Then Exit Function
    
    If G.DictWeinToEmpId.exists(wein) Then
        EmpIdFromWein = G.DictWeinToEmpId(wein)
    End If
End Function

'------------------------------------------------------------------------------
' Function: WeinFromEmpCode
' Purpose: Get WEIN from Employee Code
' Parameters:
'   empCode - Employee Code
' Returns: WEIN or empty string if not found
'------------------------------------------------------------------------------
Public Function WeinFromEmpCode(empCode As String) As String
    WeinFromEmpCode = ""
    
    If G.DictEmpCodeToWein Is Nothing Then Exit Function
    
    If G.DictEmpCodeToWein.exists(empCode) Then
        WeinFromEmpCode = G.DictEmpCodeToWein(empCode)
    End If
End Function

'------------------------------------------------------------------------------
' Function: EmpCodeFromWein
' Purpose: Get Employee Code from WEIN
' Parameters:
'   wein - WEIN
' Returns: Employee Code or empty string if not found
'------------------------------------------------------------------------------
Public Function EmpCodeFromWein(wein As String) As String
    EmpCodeFromWein = ""
    
    If G.DictWeinToEmpCode Is Nothing Then Exit Function
    
    If G.DictWeinToEmpCode.exists(wein) Then
        EmpCodeFromWein = G.DictWeinToEmpCode(wein)
    End If
End Function

'------------------------------------------------------------------------------
' Function: MapOrAppendByWein
' Purpose: Find row for WEIN in worksheet, or append new row if not found
' Parameters:
'   ws - Worksheet
'   wein - WEIN to find/add
'   weinColName - Name of WEIN column
'   weinIndex - Dictionary of WEIN -> row (will be updated if new row added)
' Returns: Row number for the WEIN
'------------------------------------------------------------------------------
Public Function MapOrAppendByWein( _
    ws As Worksheet, _
    wein As String, _
    weinColName As String, _
    ByRef weinIndex As Object _
) As Long
    
    Dim weinCol As Long
    Dim newRow As Long
    
    On Error GoTo ErrHandler
    
    ' Check if already in index
    If weinIndex.exists(wein) Then
        MapOrAppendByWein = weinIndex(wein)
        Exit Function
    End If
    
    ' Find WEIN column
    weinCol = FindColumnByHeader(ws.Rows(1), weinColName)
    If weinCol = 0 Then
        MapOrAppendByWein = 0
        Exit Function
    End If
    
    ' Append new row
    newRow = ws.Cells(ws.Rows.count, weinCol).End(xlUp).row + 1
    ws.Cells(newRow, weinCol).value = wein
    
    ' Update index
    weinIndex.Add wein, newRow
    
    MapOrAppendByWein = newRow
    Exit Function
    
ErrHandler:
    LogError "modEmployeeMappingService", "MapOrAppendByWein", Err.Number, Err.Description
    MapOrAppendByWein = 0
End Function

'------------------------------------------------------------------------------
' Function: GetAllWeins
' Purpose: Get collection of all known WEINs
' Returns: Collection of WEIN strings
'------------------------------------------------------------------------------
Public Function GetAllWeins() As Collection
    Dim col As New Collection
    Dim k As Variant
    
    If Not G.DictWeinToEmpId Is Nothing Then
        For Each k In G.DictWeinToEmpId.Keys
            col.Add CStr(k)
        Next k
    End If
    
    Set GetAllWeins = col
End Function

'------------------------------------------------------------------------------
' Function: GetAllEmployeeIds
' Purpose: Get collection of all known Employee IDs
' Returns: Collection of Employee ID strings
'------------------------------------------------------------------------------
Public Function GetAllEmployeeIds() As Collection
    Dim col As New Collection
    Dim k As Variant
    
    If Not G.DictEmpIdToWein Is Nothing Then
        For Each k In G.DictEmpIdToWein.Keys
            col.Add CStr(k)
        Next k
    End If
    
    Set GetAllEmployeeIds = col
End Function

'------------------------------------------------------------------------------
' Function: FindEmployeeIdColumn
' Purpose: Find employee ID column by trying multiple variant names
' Parameters:
'   headerRow - Range containing header row
'   variants - Comma-separated list of possible column names
' Returns: Column index (1-based) or 0 if not found
'------------------------------------------------------------------------------
Public Function FindEmployeeIdColumn(headerRow As Range, variants As String) As Long
    Dim variantArr() As String
    Dim i As Long
    Dim col As Long
    
    FindEmployeeIdColumn = 0
    variantArr = Split(variants, ",")
    
    For i = LBound(variantArr) To UBound(variantArr)
        col = FindColumnByHeader(headerRow, Trim(variantArr(i)))
        If col > 0 Then
            FindEmployeeIdColumn = col
            Exit Function
        End If
    Next i
End Function

'------------------------------------------------------------------------------
' Function: FindAnyEmployeeIdColumn
' Purpose: Find any employee ID column by trying all known variants
' Parameters:
'   headerRow - Range containing header row
' Returns: Column index (1-based) or 0 if not found
'------------------------------------------------------------------------------
Public Function FindAnyEmployeeIdColumn(headerRow As Range) As Long
    FindAnyEmployeeIdColumn = FindEmployeeIdColumn(headerRow, ALL_EMP_ID_VARIANTS)
End Function

'------------------------------------------------------------------------------
' Function: NormalizeEmployeeId
' Purpose: Convert any employee ID to WEIN (the canonical form)
' Parameters:
'   empIdValue - Employee ID value from any system
' Returns: WEIN or original value if no mapping found
'------------------------------------------------------------------------------
Public Function NormalizeEmployeeId(empIdValue As String) As String
    Dim result As String
    
    result = Trim(empIdValue)
    If result = "" Then
        NormalizeEmployeeId = ""
        Exit Function
    End If
    
    ' Try to convert to WEIN using all available mappings
    ' First check if it's already a WEIN
    If Not G.DictWeinToEmpId Is Nothing Then
        If G.DictWeinToEmpId.exists(result) Then
            NormalizeEmployeeId = result
            Exit Function
        End If
    End If
    
    ' Try Employee ID -> WEIN
    If Not G.DictEmpIdToWein Is Nothing Then
        If G.DictEmpIdToWein.exists(result) Then
            NormalizeEmployeeId = G.DictEmpIdToWein(result)
            Exit Function
        End If
    End If
    
    ' Try Employee Code -> WEIN
    If Not G.DictEmpCodeToWein Is Nothing Then
        If G.DictEmpCodeToWein.exists(result) Then
            NormalizeEmployeeId = G.DictEmpCodeToWein(result)
            Exit Function
        End If
    End If
    
    ' Return original value if no mapping found
    NormalizeEmployeeId = result
End Function

'------------------------------------------------------------------------------
' Function: GetEmployeeIdVariants
' Purpose: Get array of all employee ID field name variants
' Returns: Array of variant names
'------------------------------------------------------------------------------
Public Function GetEmployeeIdVariants() As Variant
    GetEmployeeIdVariants = Split(ALL_EMP_ID_VARIANTS, ",")
End Function

'------------------------------------------------------------------------------
' Function: GetWeinVariants
' Purpose: Get array of WEIN field name variants
' Returns: Array of variant names
'------------------------------------------------------------------------------
Public Function GetWeinVariants() As Variant
    GetWeinVariants = Split(WEIN_VARIANTS, ",")
End Function

'------------------------------------------------------------------------------
' Function: GetEmpCodeVariants
' Purpose: Get array of Employee Code field name variants
' Returns: Array of variant names
'------------------------------------------------------------------------------
Public Function GetEmpCodeVariants() As Variant
    GetEmpCodeVariants = Split(EMP_CODE_VARIANTS, ",")
End Function

'------------------------------------------------------------------------------
' Function: FindHeaderRow
' Purpose: Find the row containing the header by searching for a key column name
' Parameters:
'   ws - Worksheet to search
'   keyColumnName - Column name to search for (e.g., "Employee ID")
' Returns: Row number (1-based) or 0 if not found
'------------------------------------------------------------------------------
Private Function FindHeaderRow(ws As Worksheet, keyColumnName As String) As Long
    Dim i As Long, j As Long
    Dim cellValue As String
    
    FindHeaderRow = 0
    
    ' Search first 20 rows
    For i = 1 To 20
        ' Search across columns
        For j = 1 To 100
            On Error Resume Next
            cellValue = Trim(CStr(ws.Cells(i, j).value))
            On Error GoTo 0
            
            If UCase(cellValue) = UCase(keyColumnName) Then
                FindHeaderRow = i
                Exit Function
            End If
        Next j
    Next i
End Function


