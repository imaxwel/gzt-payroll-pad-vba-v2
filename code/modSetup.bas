Attribute VB_Name = "modSetup"
'==============================================================================
' Module: modSetup
' Purpose: Setup and initialization utilities
' Description: Creates Runtime sheet and named ranges for initial setup
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Sub: SetupRuntimeSheet
' Purpose: Create and configure the Runtime sheet with named ranges
'------------------------------------------------------------------------------
Public Sub SetupRuntimeSheet()
    Dim ws As Worksheet
    Dim wb As Workbook
    
    On Error GoTo ErrHandler
    
    Set wb = ThisWorkbook
    
    ' Check if Runtime sheet exists
    On Error Resume Next
    Set ws = wb.Worksheets("Runtime")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then
        ' Create Runtime sheet
        Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        ws.Name = "Runtime"
    End If
    
    ' Clear existing content
    ws.Cells.Clear
    
    ' Setup headers and labels
    ws.Cells(1, 1).Value = "HK Payroll Automation - Runtime Configuration"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 1).Font.Size = 14
    
    ' Parameter labels
    ws.Cells(2, 1).Value = "InputFolder"
    ws.Cells(3, 1).Value = "OutputFolder"
    ws.Cells(4, 1).Value = "ConfigFolder"
    ws.Cells(5, 1).Value = "PayrollMonth"
    ws.Cells(6, 1).Value = "RunDate"
    ws.Cells(7, 1).Value = "LogFolder"
    ws.Cells(8, 1).Value = "SP_Status"
    ws.Cells(9, 1).Value = "SP_Message"
    
    ' Default values
    ws.Cells(2, 2).Value = "C:\HK_Payroll\input\"
    ws.Cells(3, 2).Value = "C:\HK_Payroll\output\"
    ws.Cells(4, 2).Value = "C:\HK_Payroll\config\"
    ws.Cells(5, 2).Value = Format(Date, "YYYYMM")
    ws.Cells(6, 2).Value = Date
    ws.Cells(7, 2).Value = "C:\HK_Payroll\log\"
    ws.Cells(8, 2).Value = ""
    ws.Cells(9, 2).Value = ""
    
    ' Create named ranges
    CreateNamedRangeIfNotExists wb, "InputFolder", ws.Range("B2")
    CreateNamedRangeIfNotExists wb, "OutputFolder", ws.Range("B3")
    CreateNamedRangeIfNotExists wb, "ConfigFolder", ws.Range("B4")
    CreateNamedRangeIfNotExists wb, "PayrollMonth", ws.Range("B5")
    CreateNamedRangeIfNotExists wb, "RunDate", ws.Range("B6")
    CreateNamedRangeIfNotExists wb, "LogFolder", ws.Range("B7")
    CreateNamedRangeIfNotExists wb, "SP_Status", ws.Range("B8")
    CreateNamedRangeIfNotExists wb, "SP_Message", ws.Range("B9")
    
    ' Format
    ws.Range("A2:A9").Font.Bold = True
    ws.Range("B2:B9").Interior.Color = RGB(255, 255, 200)
    ws.Columns("A:B").AutoFit
    
    ' Add instructions
    ws.Cells(11, 1).Value = "Instructions:"
    ws.Cells(11, 1).Font.Bold = True
    ws.Cells(12, 1).Value = "1. Update the paths above to match your environment"
    ws.Cells(13, 1).Value = "2. Set PayrollMonth to the target month (YYYYMM format)"
    ws.Cells(14, 1).Value = "3. Set RunDate to today's date"
    ws.Cells(15, 1).Value = "4. Run 'Run_Subprocess1' or 'Run_Subprocess2' macro"
    
    MsgBox "Runtime sheet created successfully!" & vbCrLf & vbCrLf & _
           "Please update the configuration values before running.", _
           vbInformation, "Setup Complete"
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error setting up Runtime sheet: " & Err.Description, vbCritical, "Setup Error"
End Sub

'------------------------------------------------------------------------------
' Sub: CreateNamedRangeIfNotExists
' Purpose: Create a named range if it doesn't already exist
'------------------------------------------------------------------------------
Private Sub CreateNamedRangeIfNotExists(wb As Workbook, rangeName As String, rng As Range)
    On Error Resume Next
    wb.Names(rangeName).Delete
    On Error GoTo 0
    
    wb.Names.Add Name:=rangeName, RefersTo:=rng
End Sub

'------------------------------------------------------------------------------
' Sub: SetupLogSheet
' Purpose: Create the Log sheet
'------------------------------------------------------------------------------
Public Sub SetupLogSheet()
    Dim ws As Worksheet
    Dim wb As Workbook
    
    On Error GoTo ErrHandler
    
    Set wb = ThisWorkbook
    
    ' Check if Log sheet exists
    On Error Resume Next
    Set ws = wb.Worksheets("Log")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "Log"
    End If
    
    ' Setup headers
    ws.Cells(1, 1).Value = "Timestamp"
    ws.Cells(1, 2).Value = "Level"
    ws.Cells(1, 3).Value = "Message"
    ws.Rows(1).Font.Bold = True
    
    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 100
    
    MsgBox "Log sheet created successfully!", vbInformation, "Setup Complete"
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error setting up Log sheet: " & Err.Description, vbCritical, "Setup Error"
End Sub

'------------------------------------------------------------------------------
' Sub: SetupAll
' Purpose: Run all setup routines
'------------------------------------------------------------------------------
Public Sub SetupAll()
    SetupRuntimeSheet
    SetupLogSheet
    
    MsgBox "All setup completed!" & vbCrLf & vbCrLf & _
           "Next steps:" & vbCrLf & _
           "1. Update Runtime sheet with your paths" & vbCrLf & _
           "2. Place input files in the input folder" & vbCrLf & _
           "3. Create config.xlsx in the config folder" & vbCrLf & _
           "4. Run 'Run_Subprocess1' to start", _
           vbInformation, "Setup Complete"
End Sub

'------------------------------------------------------------------------------
' Sub: ValidateSetup
' Purpose: Validate that all required setup is complete
'------------------------------------------------------------------------------
Public Sub ValidateSetup()
    Dim errors As String
    Dim ws As Worksheet
    
    errors = ""
    
    ' Check Runtime sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Runtime")
    On Error GoTo 0
    
    If ws Is Nothing Then
        errors = errors & "- Runtime sheet not found" & vbCrLf
    Else
        ' Check named ranges
        If Range("InputFolder").Value = "" Then errors = errors & "- InputFolder is empty" & vbCrLf
        If Range("OutputFolder").Value = "" Then errors = errors & "- OutputFolder is empty" & vbCrLf
        If Range("ConfigFolder").Value = "" Then errors = errors & "- ConfigFolder is empty" & vbCrLf
        If Range("PayrollMonth").Value = "" Then errors = errors & "- PayrollMonth is empty" & vbCrLf
        If Range("RunDate").Value = "" Then errors = errors & "- RunDate is empty" & vbCrLf
        
        ' Check folder existence
        If Dir(Range("InputFolder").Value, vbDirectory) = "" Then
            errors = errors & "- Input folder does not exist" & vbCrLf
        End If
        If Dir(Range("OutputFolder").Value, vbDirectory) = "" Then
            errors = errors & "- Output folder does not exist" & vbCrLf
        End If
        If Dir(Range("ConfigFolder").Value, vbDirectory) = "" Then
            errors = errors & "- Config folder does not exist" & vbCrLf
        End If
    End If
    
    If errors = "" Then
        MsgBox "Setup validation passed!" & vbCrLf & vbCrLf & _
               "You can now run the automation.", _
               vbInformation, "Validation Passed"
    Else
        MsgBox "Setup validation failed:" & vbCrLf & vbCrLf & errors, _
               vbExclamation, "Validation Failed"
    End If
End Sub
