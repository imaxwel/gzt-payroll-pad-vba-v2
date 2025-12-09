Attribute VB_Name = "modSP2_CheckResultTemplate"
'==============================================================================
' Module: modSP2_CheckResultTemplate
' Purpose: Subprocess 2 - Check Result Template Structure
' Description: Defines the 63 fields with Check and Diff columns according to
'              HK payroll validation output template format
'==============================================================================
Option Explicit

' Template column structure type
Public Type tCheckResultColumn
    benchmarkName As String      ' Original column name from Payroll Report
    HasCheck As Boolean          ' Whether this field has a Check column
    HasDiff As Boolean           ' Whether this field has a Diff column
    CheckColIndex As Long        ' Column index for Check (runtime)
    DiffColIndex As Long         ' Column index for Diff (runtime)
    BenchmarkColIndex As Long    ' Column index for Benchmark (runtime)
End Type

' Module-level template definition
Private mTemplateFields() As tCheckResultColumn
Private mTemplateInitialized As Boolean

'------------------------------------------------------------------------------
' Sub: InitializeTemplate
' Purpose: Initialize the 63-field template structure based on HK payroll
'          validation output template format
'------------------------------------------------------------------------------
Public Sub InitializeTemplate()
    Dim idx As Long
    
    If mTemplateInitialized Then Exit Sub
    
    ' Define all 63 fields with Check/Diff requirements
    ' Based on the template header provided by user
    ReDim mTemplateFields(1 To 63)
    idx = 0
    
    ' Field 1: WEIN - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "WEIN"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 2: Legal full name - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Legal full name"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 3: Last Hired Date - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Last Hired Date"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 4: Last Employment Date - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Last Employment Date"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 5: Business Department - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Business Department"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 6: Position Title - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Position Title"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 7: Cost Center Code - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Cost Center Code"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 8: Monthly Base Pay - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Monthly Base Pay"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 9: Base Pay 60001000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Base Pay 60001000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 10: Salary Adj 60001000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Salary Adj 60001000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 11: Monthly Base Pay(Temp) - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Monthly Base Pay(Temp)"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 12: Base Pay(Temp) 60101000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Base Pay(Temp) 60101000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 13: Monthly Transport Allowance - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Monthly Transport Allowance"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 14: Transport Allowance 60409960 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Transport Allowance 60409960"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 15: Transport Allowance Adj 60409960 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Transport Allowance Adj 60409960"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 16: Total EAO Adj 60409960 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Total EAO Adj 60409960"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 17: Maternity Leave Payment 60001000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Maternity Leave Payment 60001000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 18: Paid Parental Time Off (PPTO) payment - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Paid Parental Time Off (PPTO) payment"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 19: Sick Leave Payment 60001000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Sick Leave Payment 60001000"
        .HasCheck = True
        .HasDiff = True
    End With

    ' Field 20: Annual Incentive 60201000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Annual Incentive 60201000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 21: Sales Incentive (Quantitative) 21201000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Sales Incentive (Quantitative)   21201000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 22: Sales Incentive (Qualitative) 21201000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Sales Incentive (Qualitative) 21201000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 23: Inspire Cash 60702000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Inspire Cash 60702000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 24: Inspire Points (Gross Up) 60701000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Inspire Points (Gross Up) 60701000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 25: Shares Dividend 60204001 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Shares Dividend 60204001"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 26: Red Packet 69001000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Red Packet 69001000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 27: Year End Bonus 60208000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Year End Bonus 60208000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 28: Lump Sum Bonus 60409960 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Lump Sum Bonus 60409960"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 29: Referral Bonus 69001000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Referral Bonus 69001000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 30: Sign On Bonus 60409960 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Sign On Bonus 60409960"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 31: Retention Bonus 60409960 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Retention Bonus 60409960"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 32: Other Bonus 99999999 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Other Bonus 99999999"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 33: Manager of the Year Award 60208000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Manager of the Year Award 60208000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 34: MD Award 60208000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "MD Award 60208000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 35: Other Rewards 99999999 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Other Rewards 99999999"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 36: Other Allowance 60409960 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Other Allowance 60409960"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 37: Flexible benefits - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Flexible benefits"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 38: Untaken Annual Leave Payment 60409960 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Untaken Annual Leave Payment 60409960"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 39: PIL ER to EE 60001000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "PIL ER to EE 60001000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 40: Back Pay 99999999 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Back Pay 99999999"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 41: Gratuities 99999999 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Gratuities 99999999"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 42: Severance Payment 60404000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Severance Payment 60404000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 43: Long Service Payment 60409960 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Long Service Payment 60409960"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 44: IA Pay Split - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "IA Pay Split"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 45: MPF Relevant Income - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "MPF Relevant Income"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 46: MPF VC Relevant Income - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "MPF VC Relevant Income"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 47: ORSO Relevant Income - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "ORSO Relevant Income"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 48: Optional Group Upgrade 21351000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Optional Group Upgrade 21351000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 49: No Pay Leave Deduction 60001000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "No Pay Leave Deduction 60001000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 50: PIL EE to ER 60001000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "PIL EE to ER 60001000"
        .HasCheck = True
        .HasDiff = True
    End With

    ' Field 51: MPF EE VC Percentage - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "MPF EE VC Percentage"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 52: MPF ER VC Percentage - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "MPF ER VC Percentage"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 53: Percent Of ORSO EE - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Percent Of ORSO EE"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 54: Percent Of ORSO ER - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "Percent Of ORSO ER"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 55: ORSO EE 60801000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "ORSO EE 60801000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 56: ORSO ER Adj - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "ORSO ER Adj"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 57: ORSO ER 60801000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "ORSO ER 60801000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 58: MPF EE MC 21251000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "MPF EE MC 21251000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 59: MPF EE VC 21251000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "MPF EE VC 21251000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 60: MPF ER MC 60801000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "MPF ER MC 60801000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 61: MPF ER VC 60801000 - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "MPF ER VC 60801000"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 62: InspirePoints - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "InspirePoints"
        .HasCheck = True
        .HasDiff = True
    End With
    
    ' Field 63: PPTO EAO Rate input - has Check and Diff
    idx = idx + 1
    With mTemplateFields(idx)
        .benchmarkName = "PPTO EAO Rate input"
        .HasCheck = True
        .HasDiff = True
    End With
    
    mTemplateInitialized = True
    
    LogInfo "modSP2_CheckResultTemplate", "InitializeTemplate", _
        "Template initialized with " & idx & " fields"
End Sub

'------------------------------------------------------------------------------
' Function: GetTemplateFieldCount
' Purpose: Get the number of template fields
' Returns: Number of fields (63)
'------------------------------------------------------------------------------
Public Function GetTemplateFieldCount() As Long
    If Not mTemplateInitialized Then InitializeTemplate
    GetTemplateFieldCount = UBound(mTemplateFields)
End Function

'------------------------------------------------------------------------------
' Function: GetTemplateField
' Purpose: Get a template field by index
' Parameters:
'   idx - Field index (1-based)
' Returns: tCheckResultColumn structure
'------------------------------------------------------------------------------
Public Function GetTemplateField(idx As Long) As tCheckResultColumn
    If Not mTemplateInitialized Then InitializeTemplate
    If idx >= 1 And idx <= UBound(mTemplateFields) Then
        GetTemplateField = mTemplateFields(idx)
    End If
End Function

'------------------------------------------------------------------------------
' Function: FindTemplateFieldByName
' Purpose: Find a template field by benchmark name
' Parameters:
'   benchmarkName - Name of the benchmark column
' Returns: Field index (1-based) or 0 if not found
'------------------------------------------------------------------------------
Public Function FindTemplateFieldByName(benchmarkName As String) As Long
    Dim i As Long
    Dim searchName As String
    
    If Not mTemplateInitialized Then InitializeTemplate
    
    searchName = UCase(Trim(benchmarkName))
    
    For i = 1 To UBound(mTemplateFields)
        If UCase(Trim(mTemplateFields(i).benchmarkName)) = searchName Then
            FindTemplateFieldByName = i
            Exit Function
        End If
    Next i
    
    FindTemplateFieldByName = 0
End Function

'------------------------------------------------------------------------------
' Sub: UpdateTemplateColumnIndices
' Purpose: Update column indices after building the Check Result structure
' Parameters:
'   ws - Check Result worksheet
'   headerRow - Row number of the header
'------------------------------------------------------------------------------
Public Sub UpdateTemplateColumnIndices(ws As Worksheet, headerRow As Long)
    Dim i As Long
    Dim lastCol As Long
    Dim col As Long
    Dim headerValue As String
    Dim benchmarkName As String
    
    If Not mTemplateInitialized Then InitializeTemplate
    
    lastCol = ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column
    
    ' Reset all indices
    For i = 1 To UBound(mTemplateFields)
        mTemplateFields(i).BenchmarkColIndex = 0
        mTemplateFields(i).CheckColIndex = 0
        mTemplateFields(i).DiffColIndex = 0
    Next i
    
    ' Scan header row and update indices
    For col = 1 To lastCol
        headerValue = Trim(CStr(Nz(ws.Cells(headerRow, col).Value, "")))
        
        ' Check if this is a Check column
        If Right(UCase(headerValue), 5) = "CHECK" Then
            benchmarkName = Left(headerValue, Len(headerValue) - 6) ' Remove " Check"
            i = FindTemplateFieldByName(benchmarkName)
            If i > 0 Then
                mTemplateFields(i).CheckColIndex = col
            End If
        ' Check if this is a Diff column
        ElseIf Right(UCase(headerValue), 4) = "DIFF" Then
            benchmarkName = Left(headerValue, Len(headerValue) - 5) ' Remove " Diff"
            i = FindTemplateFieldByName(benchmarkName)
            If i > 0 Then
                mTemplateFields(i).DiffColIndex = col
            End If
        Else
            ' This is a benchmark column
            i = FindTemplateFieldByName(headerValue)
            If i > 0 Then
                mTemplateFields(i).BenchmarkColIndex = col
            End If
        End If
    Next col
End Sub

'------------------------------------------------------------------------------
' Function: GetBenchmarkColIndex
' Purpose: Get the benchmark column index for a field
' Parameters:
'   fieldName - Name of the field
' Returns: Column index or 0 if not found
'------------------------------------------------------------------------------
Public Function GetBenchmarkColIndex(fieldName As String) As Long
    Dim i As Long
    i = FindTemplateFieldByName(fieldName)
    If i > 0 Then
        GetBenchmarkColIndex = mTemplateFields(i).BenchmarkColIndex
    Else
        GetBenchmarkColIndex = 0
    End If
End Function

'------------------------------------------------------------------------------
' Function: GetCheckColIndex
' Purpose: Get the Check column index for a field
' Parameters:
'   fieldName - Name of the field
' Returns: Column index or 0 if not found
'------------------------------------------------------------------------------
Public Function GetCheckColIndex(fieldName As String) As Long
    Dim i As Long
    i = FindTemplateFieldByName(fieldName)
    If i > 0 Then
        GetCheckColIndex = mTemplateFields(i).CheckColIndex
    Else
        GetCheckColIndex = 0
    End If
End Function

'------------------------------------------------------------------------------
' Function: GetDiffColIndex
' Purpose: Get the Diff column index for a field
' Parameters:
'   fieldName - Name of the field
' Returns: Column index or 0 if not found
'------------------------------------------------------------------------------
Public Function GetDiffColIndex(fieldName As String) As Long
    Dim i As Long
    i = FindTemplateFieldByName(fieldName)
    If i > 0 Then
        GetDiffColIndex = mTemplateFields(i).DiffColIndex
    Else
        GetDiffColIndex = 0
    End If
End Function

'------------------------------------------------------------------------------
' Function: GetAllTemplateFields
' Purpose: Get all template fields as an array
' Returns: Array of tCheckResultColumn
'------------------------------------------------------------------------------
Public Function GetAllTemplateFields() As Variant
    If Not mTemplateInitialized Then InitializeTemplate
    GetAllTemplateFields = mTemplateFields
End Function
