Attribute VB_Name = "modEntryPoints"
'==============================================================================
' Module: modEntryPoints
' Purpose: Entry point macros for PAD (Power Automate Desktop)
' Description: Contains Run_Subprocess1 and Run_Subprocess2 entry macros
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Sub: Run_Subprocess1
' Purpose: Main entry point for Subprocess 1 (Flexi form output generation)
' Called by: PAD
'------------------------------------------------------------------------------
Public Sub Run_Subprocess1()
    Dim p As tRunParams
    
    On Error GoTo ErrHandler
    
    ' Clear previous log
    ClearLogSheet
    
    LogInfo "modEntryPoints", "Run_Subprocess1", "=== Starting Subprocess 1 ==="
    
    ' 1) Fetch parameters from Runtime sheet
    p = LoadRunParamsFromWorkbook()
    
    LogInfo "modEntryPoints", "Run_Subprocess1", _
        "Parameters: PayrollMonth=" & p.payrollMonth & ", InputFolder=" & p.InputFolder
    
    ' 2) Initialize global context
    InitAppContext p
    
    ' 3) Execute Subprocess 1 logic
    SP1_Execute
    
    ' 4) Set success status for PAD
    SetStatus "OK", "Subprocess 1 completed successfully"
    
    LogInfo "modEntryPoints", "Run_Subprocess1", "=== Subprocess 1 Completed Successfully ==="
    
    MsgBox "Subprocess 1 completed successfully!", vbInformation, "HK Payroll Automation"
    Exit Sub
    
ErrHandler:
    LogError "modEntryPoints", "Run_Subprocess1", Err.Number, Err.Description
    SetStatus "ERROR", Err.Description
    MsgBox "Subprocess 1 failed: " & Err.Description, vbCritical, "HK Payroll Automation"
End Sub

'------------------------------------------------------------------------------
' Sub: Run_Subprocess2
' Purpose: Main entry point for Subprocess 2 (Validation output generation)
' Called by: PAD
'------------------------------------------------------------------------------
Public Sub Run_Subprocess2()
    Dim p As tRunParams
    
    On Error GoTo ErrHandler
    
    ' Clear previous log
    ClearLogSheet
    
    LogInfo "modEntryPoints", "Run_Subprocess2", "=== Starting Subprocess 2 ==="
    
    ' 1) Fetch parameters from Runtime sheet
    p = LoadRunParamsFromWorkbook()
    
    LogInfo "modEntryPoints", "Run_Subprocess2", _
        "Parameters: PayrollMonth=" & p.payrollMonth & ", InputFolder=" & p.InputFolder
    
    ' 2) Initialize global context
    InitAppContext p
    
    ' 3) Execute Subprocess 2 logic
    SP2_Execute
    
    ' 4) Set success status for PAD
    SetStatus "OK", "Subprocess 2 completed successfully"
    
    LogInfo "modEntryPoints", "Run_Subprocess2", "=== Subprocess 2 Completed Successfully ==="
    
    MsgBox "Subprocess 2 completed successfully!", vbInformation, "HK Payroll Automation"
    Exit Sub
    
ErrHandler:
    LogError "modEntryPoints", "Run_Subprocess2", Err.Number, Err.Description
    SetStatus "ERROR", Err.Description
    MsgBox "Subprocess 2 failed: " & Err.Description, vbCritical, "HK Payroll Automation"
End Sub

'------------------------------------------------------------------------------
' Sub: Run_Both
' Purpose: Run both Subprocess 1 and Subprocess 2 sequentially
' Called by: PAD or manual execution
'------------------------------------------------------------------------------
Public Sub Run_Both()
    On Error GoTo ErrHandler
    
    ' Run Subprocess 1
    Run_Subprocess1
    
    ' Check if Subprocess 1 succeeded
    If ThisWorkbook.Worksheets("Runtime").Range("SP_Status").value = "OK" Then
        ' Run Subprocess 2
        Run_Subprocess2
    Else
        LogWarning "modEntryPoints", "Run_Both", "Subprocess 1 failed, skipping Subprocess 2"
    End If
    
    Exit Sub
    
ErrHandler:
    LogError "modEntryPoints", "Run_Both", Err.Number, Err.Description
    SetStatus "ERROR", Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: TestConfiguration
' Purpose: Test configuration loading without running full process
'------------------------------------------------------------------------------
Public Sub TestConfiguration()
    Dim p As tRunParams
    
    On Error GoTo ErrHandler
    
    p = LoadRunParamsFromWorkbook()
    
    MsgBox "Configuration loaded successfully!" & vbCrLf & vbCrLf & _
           "Input Folder: " & p.InputFolder & vbCrLf & _
           "Output Folder: " & p.OutputFolder & vbCrLf & _
           "Config Folder: " & p.ConfigFolder & vbCrLf & _
           "Payroll Month: " & p.payrollMonth & vbCrLf & _
           "Run Date: " & Format(p.RunDate, "yyyy-mm-dd"), _
           vbInformation, "Configuration Test"
    Exit Sub
    
ErrHandler:
    MsgBox "Configuration test failed: " & Err.Description, vbCritical, "Configuration Test"
End Sub
