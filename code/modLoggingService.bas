Attribute VB_Name = "modLoggingService"
'==============================================================================
' Module: modLoggingService
' Purpose: Centralized logging services
' Description: Handles logging to worksheet and/or text file
'==============================================================================
Option Explicit

Private Const LOG_SHEET_NAME As String = "Log"
Private mLogFileHandle As Integer
Private mLogFilePath As String

'------------------------------------------------------------------------------
' Sub: LogError
' Purpose: Log an error message
' Parameters:
'   moduleName - Name of the module where error occurred
'   procName - Name of the procedure where error occurred
'   errNum - Error number
'   errDesc - Error description
'------------------------------------------------------------------------------
Public Sub LogError(moduleName As String, procName As String, errNum As Long, errDesc As String)
    Dim msg As String
    msg = "ERROR [" & moduleName & "." & procName & "] #" & errNum & ": " & errDesc
    WriteLog "ERROR", msg
End Sub

'------------------------------------------------------------------------------
' Sub: LogWarning
' Purpose: Log a warning message
' Parameters:
'   moduleName - Name of the module
'   procName - Name of the procedure
'   message - Warning message
'------------------------------------------------------------------------------
Public Sub LogWarning(moduleName As String, procName As String, message As String)
    Dim msg As String
    msg = "WARNING [" & moduleName & "." & procName & "] " & message
    WriteLog "WARNING", msg
End Sub

'------------------------------------------------------------------------------
' Sub: LogInfo
' Purpose: Log an informational message
' Parameters:
'   moduleName - Name of the module
'   procName - Name of the procedure
'   message - Info message
'------------------------------------------------------------------------------
Public Sub LogInfo(moduleName As String, procName As String, message As String)
    Dim msg As String
    msg = "INFO [" & moduleName & "." & procName & "] " & message
    WriteLog "INFO", msg
End Sub

'------------------------------------------------------------------------------
' Sub: WriteLog
' Purpose: Write a log entry to sheet and/or file
' Parameters:
'   logLevel - Log level (ERROR, WARNING, INFO)
'   message - Log message
'------------------------------------------------------------------------------
Private Sub WriteLog(logLevel As String, message As String)
    Dim timestamp As String
    timestamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    ' Write to log sheet
    WriteToLogSheet timestamp, logLevel, message
    
    ' Write to log file if configured
    WriteToLogFile timestamp, logLevel, message
    
    ' Also output to Immediate window for debugging
    Debug.Print timestamp & " " & logLevel & " " & message
End Sub

'------------------------------------------------------------------------------
' Sub: WriteToLogSheet
' Purpose: Write log entry to the Log worksheet
'------------------------------------------------------------------------------
Private Sub WriteToLogSheet(timestamp As String, logLevel As String, message As String)
    Dim ws As Worksheet
    Dim nextRow As Long
    
    On Error Resume Next
    
    ' Get or create log sheet
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = LOG_SHEET_NAME
        
        ' Add headers
        ws.Cells(1, 1).Value = "Timestamp"
        ws.Cells(1, 2).Value = "Level"
        ws.Cells(1, 3).Value = "Message"
        ws.Rows(1).Font.Bold = True
    End If
    
    ' Find next row
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Write log entry
    ws.Cells(nextRow, 1).Value = timestamp
    ws.Cells(nextRow, 2).Value = logLevel
    ws.Cells(nextRow, 3).Value = message
    
    ' Color code by level
    Select Case logLevel
        Case "ERROR"
            ws.Cells(nextRow, 2).Interior.Color = RGB(255, 200, 200)
        Case "WARNING"
            ws.Cells(nextRow, 2).Interior.Color = RGB(255, 255, 200)
        Case "INFO"
            ws.Cells(nextRow, 2).Interior.Color = RGB(200, 255, 200)
    End Select
    
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Sub: WriteToLogFile
' Purpose: Write log entry to text file
'------------------------------------------------------------------------------
Private Sub WriteToLogFile(timestamp As String, logLevel As String, message As String)
    On Error Resume Next
    
    ' Only write if log folder is configured
    If G.RunParams.LogFolder = "" Then Exit Sub
    
    ' Initialize log file if needed
    If mLogFilePath = "" Then
        mLogFilePath = G.RunParams.LogFolder & "PayrollAutomation_" & _
            Format(Now, "YYYYMMDD_HHNNSS") & ".log"
    End If
    
    ' Append to log file
    mLogFileHandle = FreeFile
    Open mLogFilePath For Append As #mLogFileHandle
    Print #mLogFileHandle, timestamp & vbTab & logLevel & vbTab & message
    Close #mLogFileHandle
    
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Sub: ClearLogSheet
' Purpose: Clear all entries from the log sheet
'------------------------------------------------------------------------------
Public Sub ClearLogSheet()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    
    If Not ws Is Nothing Then
        If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row > 1 Then
            ws.Range(ws.Rows(2), ws.Rows(ws.Cells(ws.Rows.Count, 1).End(xlUp).Row)).Delete
        End If
    End If
    
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Sub: ExportLogToCSV
' Purpose: Export log sheet to CSV file
' Parameters:
'   filePath - Full path for the CSV file
'------------------------------------------------------------------------------
Public Sub ExportLogToCSV(filePath As String)
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim fileNum As Integer
    Dim line As String
    
    On Error GoTo ErrHandler
    
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    If ws Is Nothing Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = 3 ' Timestamp, Level, Message
    
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    
    For i = 1 To lastRow
        line = ""
        For j = 1 To lastCol
            If j > 1 Then line = line & ","
            line = line & """" & Replace(CStr(ws.Cells(i, j).Value), """", """""") & """"
        Next j
        Print #fileNum, line
    Next i
    
    Close #fileNum
    Exit Sub
    
ErrHandler:
    On Error Resume Next
    Close #fileNum
End Sub

'------------------------------------------------------------------------------
' Function: GetLogFilePath
' Purpose: Get the current log file path
' Returns: Log file path or empty string
'------------------------------------------------------------------------------
Public Function GetLogFilePath() As String
    GetLogFilePath = mLogFilePath
End Function

'------------------------------------------------------------------------------
' Sub: SetStatus
' Purpose: Set status for PAD to read
' Parameters:
'   status - Status string (e.g., "OK", "ERROR")
'   message - Optional message
'------------------------------------------------------------------------------
Public Sub SetStatus(status As String, Optional message As String = "")
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Runtime")
    
    If Not ws Is Nothing Then
        ws.Range("SP_Status").Value = status
        If message <> "" Then
            ws.Range("SP_Message").Value = message
        End If
    End If
    
    On Error GoTo 0
End Sub
