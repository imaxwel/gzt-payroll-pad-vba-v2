Attribute VB_Name = "modPathService"
'==============================================================================
' Module: modPathService
' Purpose: Layered directory path service - supports organizing input files by year/period type (Month/Quarter/Adhoc)
' Description: Provides unified file path resolution interface, supports current/previous month switching,
'              cross-month, cross-quarter, cross-year logic independently layered for easy directory structure adjustment
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Enum: ePeriodOffset
' Purpose: Period offset enumeration
'------------------------------------------------------------------------------
Public Enum ePeriodOffset
    poCurrentMonth = 0      ' Current month
    poPreviousMonth = -1    ' Previous month
    poNextMonth = 1         ' Next month (reserved)
End Enum

'------------------------------------------------------------------------------
' Enum: ePeriodType
' Purpose: Period type enumeration
'------------------------------------------------------------------------------
Public Enum ePeriodType
    ptMonth = 1             ' Monthly file
    ptQuarter = 2           ' Quarterly file
    ptAdhoc = 3             ' One-time/temporary file
End Enum

'------------------------------------------------------------------------------
' Type: tPeriodInfo
' Purpose: Period information structure
'------------------------------------------------------------------------------
Public Type tPeriodInfo
    Year As Integer         ' Year (e.g., 2025)
    Month As Integer        ' Month (1-12)
    Quarter As Integer      ' Quarter (1-4)
    YearMonth As String     ' "YYYYMM" format
    YearQuarter As String   ' "YYYYQX" format
End Type

' Module-level cache - base paths
Private mBaseInputPath As String
Private mBaseOutputPath As String
Private mIsPathInitialized As Boolean


'==============================================================================
' Layer 1: Base path initialization
'==============================================================================

'------------------------------------------------------------------------------
' Sub: InitPathService
' Purpose: Initialize path service, set base paths
' Parameters:
'   baseInputPath - Input file root directory (e.g., "C:\Payroll_HK\Input\")
'   baseOutputPath - Output file root directory (e.g., "C:\Payroll_HK\Output\")
'------------------------------------------------------------------------------
Public Sub InitPathService(baseInputPath As String, baseOutputPath As String)
    mBaseInputPath = EnsureTrailingSlash(baseInputPath)
    mBaseOutputPath = EnsureTrailingSlash(baseOutputPath)
    mIsPathInitialized = True
    
    LogInfo "modPathService", "InitPathService", _
        "Path service initialized. Input: " & mBaseInputPath & ", Output: " & mBaseOutputPath
End Sub

'------------------------------------------------------------------------------
' Sub: InitPathServiceFromContext
' Purpose: Initialize path service from global context
'------------------------------------------------------------------------------
Public Sub InitPathServiceFromContext()
    If G.IsInitialised Then
        ' Assume RunParams.InputFolder is the root directory of new structure
        ' If using old structure, adjustment is needed
        InitPathService G.RunParams.InputFolder, G.RunParams.OutputFolder
    Else
        Err.Raise vbObjectError + 1001, "InitPathServiceFromContext", _
            "Global context not initialized"
    End If
End Sub

'==============================================================================
' Layer 2: Period calculation logic (cross-month/cross-quarter/cross-year)
'==============================================================================

'------------------------------------------------------------------------------
' Function: GetPeriodInfo
' Purpose: Calculate target period information based on base period and offset
' Parameters:
'   baseYearMonth - Base year-month "YYYYMM"
'   offset - Period offset (ePeriodOffset)
' Returns: tPeriodInfo structure
'------------------------------------------------------------------------------
Public Function GetPeriodInfo(baseYearMonth As String, offset As ePeriodOffset) As tPeriodInfo
    Dim info As tPeriodInfo
    Dim baseYear As Integer, baseMonth As Integer
    Dim targetDate As Date
    
    ' Parse base year-month
    baseYear = CInt(Left(baseYearMonth, 4))
    baseMonth = CInt(Right(baseYearMonth, 2))
    
    ' Calculate target date (use DateAdd to automatically handle cross-year)
    targetDate = DateAdd("m", CLng(offset), DateSerial(baseYear, baseMonth, 1))
    
    ' Populate period information
    info.Year = Year(targetDate)
    info.Month = Month(targetDate)
    info.Quarter = GetQuarterFromMonth(info.Month)
    info.YearMonth = Format(targetDate, "YYYYMM")
    info.YearQuarter = CStr(info.Year) & "Q" & CStr(info.Quarter)
    
    GetPeriodInfo = info
End Function


'------------------------------------------------------------------------------
' Function: GetQuarterFromMonth
' Purpose: Get quarter from month
'------------------------------------------------------------------------------
Private Function GetQuarterFromMonth(mo As Integer) As Integer
    GetQuarterFromMonth = ((mo - 1) \ 3) + 1
End Function

'------------------------------------------------------------------------------
' Function: GetCurrentPeriodInfo
' Purpose: Get period information for current payroll month
'------------------------------------------------------------------------------
Public Function GetCurrentPeriodInfo() As tPeriodInfo
    GetCurrentPeriodInfo = GetPeriodInfo(G.Payroll.payrollMonth, poCurrentMonth)
End Function

'------------------------------------------------------------------------------
' Function: GetPreviousPeriodInfo
' Purpose: Get period information for previous month
'------------------------------------------------------------------------------
Public Function GetPreviousPeriodInfo() As tPeriodInfo
    GetPreviousPeriodInfo = GetPeriodInfo(G.Payroll.payrollMonth, poPreviousMonth)
End Function

'==============================================================================
' Layer 3: Directory path construction
'==============================================================================

'------------------------------------------------------------------------------
' Function: BuildMonthlyInputPath
' Purpose: Build monthly input file directory path
' Parameters:
'   periodInfo - Period information
' Returns: Complete directory path (e.g., "...\2025\Month\202501\")
'------------------------------------------------------------------------------
Public Function BuildMonthlyInputPath(periodInfo As tPeriodInfo) As String
    Dim path As String
    
    path = mBaseInputPath
    path = path & CStr(periodInfo.Year) & "\"
    path = path & "Month\"
    path = path & periodInfo.YearMonth & "\"
    
    BuildMonthlyInputPath = path
End Function

'------------------------------------------------------------------------------
' Function: BuildQuarterlyInputPath
' Purpose: Build quarterly input file directory path
' Parameters:
'   periodInfo - Period information
' Returns: Complete directory path (e.g., "...\2025\Quarter\")
'------------------------------------------------------------------------------
Public Function BuildQuarterlyInputPath(periodInfo As tPeriodInfo) As String
    Dim path As String
    
    path = mBaseInputPath
    path = path & CStr(periodInfo.Year) & "\"
    path = path & "Quarter\"
    
    BuildQuarterlyInputPath = path
End Function

'------------------------------------------------------------------------------
' Function: BuildAdhocInputPath
' Purpose: Build adhoc/one-time input file directory path
' Parameters:
'   periodInfo - Period information
' Returns: Complete directory path (e.g., "...\2025\Adhoc\")
'------------------------------------------------------------------------------
Public Function BuildAdhocInputPath(periodInfo As tPeriodInfo) As String
    Dim path As String
    
    path = mBaseInputPath
    path = path & CStr(periodInfo.Year) & "\"
    path = path & "Adhoc\"
    
    BuildAdhocInputPath = path
End Function


'------------------------------------------------------------------------------
' Function: BuildOutputPath
' Purpose: Build output file directory path
' Parameters:
'   periodInfo - Period information
' Returns: Complete directory path (e.g., "...\Output\2025\")
'------------------------------------------------------------------------------
Public Function BuildOutputPath(periodInfo As tPeriodInfo) As String
    Dim path As String
    
    path = mBaseOutputPath
    path = path & CStr(periodInfo.Year) & "\"
    
    BuildOutputPath = path
End Function

'==============================================================================
' Layer 4: File name mapping (logical name -> physical file name)
'==============================================================================

'------------------------------------------------------------------------------
' Function: GetPhysicalFileName
' Purpose: Get physical file name from logical name
' Parameters:
'   logicalName - Logical file name
'   periodInfo - Period information (for dynamic file names)
' Returns: Physical file name
'------------------------------------------------------------------------------
Public Function GetPhysicalFileName(logicalName As String, periodInfo As tPeriodInfo) As String
    Dim fileName As String
    
    Select Case UCase(logicalName)
        ' === Monthly files ===
        Case "PAYROLLREPORT"
            fileName = "Payroll Report.xlsx"
        Case "WORKFORCEDETAIL"
            fileName = "Workforce Detail - Payroll-AP.xlsx"
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
        Case "EMPLOYEELEAVE"
            fileName = "Employee_Leave_Transactions_Report.xlsx"
        Case "ONETIMEPAYMENT"
            fileName = "One time payment report.xlsx"
        Case "INSPIREWARDS", "INSPIREAWARDS"
            fileName = "Inspire Awards payroll report.xlsx"
        Case "EAOSUMMARY"
            fileName = "EAO Summary Report_YYYYMM.xlsx"
        Case "MERCKPAYROLL"
            fileName = "Merck Payroll Summary Report����xxx.xlsx"
        Case "SIPQIP"
            fileName = "SIP QIP.xlsx"
        Case "FLEXCLAIM"
            fileName = "MSD HK Flex_Claim_Summary_Report.xlsx"
        Case "RSUGLOBAL"
            fileName = "RSU Dividend global report.xlsx"
        Case "RSUEY"
            fileName = "RSU Dividend EY report.xlsx"
        Case "DIVIDENDEY"
            fileName = "Dividend EY report.xlsx"
        Case "AIPPAYOUTS"
            fileName = "AIP Payouts Payroll Report.xlsx"
        Case "EXTRATABLE"
            fileName = "Extra.xlsx"
        Case "ALLOWANCEPLAN"
            fileName = "Allowance plan report.xlsx"
            
        ' === Quarterly files - file name contains quarter info ===
        Case "QXPAYOUT"
            fileName = CStr(periodInfo.Year) & "QX Payout Summary.xlsx"
            
        ' === Adhoc files ===
        Case "OPTIONALMEDICAL"
            fileName = "Optional medical plan enrollment form.xlsx"
        Case "SPECIALBONUS"
            fileName = "Special_Bonus_List_" & CStr(periodInfo.Year) & ".xlsx"
            
        Case Else
            fileName = logicalName
    End Select
    
    GetPhysicalFileName = fileName
End Function


'------------------------------------------------------------------------------
' Function: GetFilePeriodType
' Purpose: Get file period type from logical name
' Parameters:
'   logicalName - Logical file name
' Returns: ePeriodType
'------------------------------------------------------------------------------
Public Function GetFilePeriodType(logicalName As String) As ePeriodType
    Select Case UCase(logicalName)
        ' Quarterly files
        Case "QXPAYOUT"
            GetFilePeriodType = ptQuarter
        ' Adhoc files
        Case "OPTIONALMEDICAL", "SPECIALBONUS"
            GetFilePeriodType = ptAdhoc
        ' Default to monthly files
        Case Else
            GetFilePeriodType = ptMonth
    End Select
End Function

'==============================================================================
' Layer 5: Unified file path interface (main API exposed externally)
'==============================================================================

'------------------------------------------------------------------------------
' Function: GetInputFilePathEx
' Purpose: Get complete input file path (supports period offset)
' Parameters:
'   logicalName - Logical file name
'   offset - Period offset (default current month)
' Returns: Complete file path
' Example:
'   GetInputFilePathEx("PayrollReport", poCurrentMonth)  -> Current month Payroll Report
'   GetInputFilePathEx("PayrollReport", poPreviousMonth) -> Previous month Payroll Report
'   GetInputFilePathEx("Termination", poPreviousMonth)   -> Previous month Termination
'------------------------------------------------------------------------------
Public Function GetInputFilePathEx(logicalName As String, _
                                   Optional offset As ePeriodOffset = poCurrentMonth) As String
    Dim periodInfo As tPeriodInfo
    Dim basePath As String
    Dim fileName As String
    Dim periodType As ePeriodType
    
    ' Ensure path service is initialized
    EnsurePathInitialized
    
    ' Get target period information
    periodInfo = GetPeriodInfo(G.Payroll.payrollMonth, offset)
    
    ' Get file period type
    periodType = GetFilePeriodType(logicalName)
    
    ' Build base path based on period type
    Select Case periodType
        Case ptMonth
            basePath = BuildMonthlyInputPath(periodInfo)
        Case ptQuarter
            basePath = BuildQuarterlyInputPath(periodInfo)
        Case ptAdhoc
            basePath = BuildAdhocInputPath(periodInfo)
    End Select
    
    ' Get physical file name
    fileName = GetPhysicalFileName(logicalName, periodInfo)
    
    GetInputFilePathEx = basePath & fileName
End Function

'------------------------------------------------------------------------------
' Function: GetCurrentMonthFilePath
' Purpose: Get current month input file path (convenience method)
'------------------------------------------------------------------------------
Public Function GetCurrentMonthFilePath(logicalName As String) As String
    GetCurrentMonthFilePath = GetInputFilePathEx(logicalName, poCurrentMonth)
End Function

'------------------------------------------------------------------------------
' Function: GetPreviousMonthFilePath
' Purpose: Get previous month input file path (convenience method)
'------------------------------------------------------------------------------
Public Function GetPreviousMonthFilePath(logicalName As String) As String
    GetPreviousMonthFilePath = GetInputFilePathEx(logicalName, poPreviousMonth)
End Function


'------------------------------------------------------------------------------
' Function: GetOutputFilePath
' Purpose: Get complete output file path
' Parameters:
'   fileName - Output file name
' Returns: Complete file path
'------------------------------------------------------------------------------
Public Function GetOutputFilePath(fileName As String) As String
    Dim periodInfo As tPeriodInfo
    
    EnsurePathInitialized
    
    periodInfo = GetCurrentPeriodInfo()
    GetOutputFilePath = BuildOutputPath(periodInfo) & fileName
End Function

'==============================================================================
' Layer 6: Compatibility layer - supports legacy directory structure (optional)
'==============================================================================

'------------------------------------------------------------------------------
' Function: GetInputFilePathLegacy
' Purpose: Get file path compatible with legacy directory structure
' Note: Use this method when directory structure has not been migrated
'------------------------------------------------------------------------------
Public Function GetInputFilePathLegacy(logicalName As String) As String
    ' Directly call the original GetInputFilePath function
    GetInputFilePathLegacy = GetInputFilePath(logicalName)
End Function

'------------------------------------------------------------------------------
' Function: GetInputFilePathAuto
' Purpose: Auto-detect directory structure and return correct path
' Parameters:
'   logicalName - Logical file name
'   offset - Period offset
' Returns: Complete file path (prioritize new structure, fallback to legacy)
'------------------------------------------------------------------------------
Public Function GetInputFilePathAuto(logicalName As String, _
                                     Optional offset As ePeriodOffset = poCurrentMonth) As String
    Dim newPath As String
    Dim legacyPath As String
    
    ' Try new structure path
    On Error Resume Next
    newPath = GetInputFilePathEx(logicalName, offset)
    On Error GoTo 0
    
    ' Check if file exists at new path
    If Len(newPath) > 0 And FileExistsSafe(newPath) Then
        GetInputFilePathAuto = newPath
        Exit Function
    End If
    
    ' Fallback to legacy structure (only valid for current month)
    If offset = poCurrentMonth Then
        legacyPath = GetInputFilePathLegacy(logicalName)
        If FileExistsSafe(legacyPath) Then
            GetInputFilePathAuto = legacyPath
            Exit Function
        End If
    End If
    
    ' Return new structure path (even if file does not exist, let caller handle)
    GetInputFilePathAuto = newPath
End Function

'==============================================================================
' Helper functions
'==============================================================================

'------------------------------------------------------------------------------
' Sub: EnsurePathInitialized
' Purpose: Ensure path service is initialized
'------------------------------------------------------------------------------
Private Sub EnsurePathInitialized()
    If Not mIsPathInitialized Then
        ' Try to initialize from global context
        If G.IsInitialised Then
            InitPathServiceFromContext
        Else
            Err.Raise vbObjectError + 1002, "EnsurePathInitialized", _
                "Path service not initialized. Call InitPathService first."
        End If
    End If
End Sub

'------------------------------------------------------------------------------
' Function: EnsureTrailingSlash
' Purpose: Ensure path ends with backslash
'------------------------------------------------------------------------------
Private Function EnsureTrailingSlash(path As String) As String
    If Len(path) > 0 And Right(path, 1) <> "\" Then
        EnsureTrailingSlash = path & "\"
    Else
        EnsureTrailingSlash = path
    End If
End Function


'------------------------------------------------------------------------------
' Sub: EnsureInputFolderExists
' Purpose: Ensure input directory exists
'------------------------------------------------------------------------------
Public Sub EnsureInputFolderExists(logicalName As String, _
                                   Optional offset As ePeriodOffset = poCurrentMonth)
    Dim periodInfo As tPeriodInfo
    Dim folderPath As String
    Dim periodType As ePeriodType
    
    EnsurePathInitialized
    
    periodInfo = GetPeriodInfo(G.Payroll.payrollMonth, offset)
    periodType = GetFilePeriodType(logicalName)
    
    Select Case periodType
        Case ptMonth
            folderPath = BuildMonthlyInputPath(periodInfo)
        Case ptQuarter
            folderPath = BuildQuarterlyInputPath(periodInfo)
        Case ptAdhoc
            folderPath = BuildAdhocInputPath(periodInfo)
    End Select
    
    EnsureFolderExists folderPath
End Sub

'------------------------------------------------------------------------------
' Sub: EnsureOutputFolderExists
' Purpose: Ensure output directory exists
'------------------------------------------------------------------------------
Public Sub EnsureOutputFolderExists()
    Dim periodInfo As tPeriodInfo
    Dim folderPath As String
    
    EnsurePathInitialized
    
    periodInfo = GetCurrentPeriodInfo()
    folderPath = BuildOutputPath(periodInfo)
    
    EnsureFolderExists folderPath
End Sub

'------------------------------------------------------------------------------
' Function: FileExistsForPeriod
' Purpose: Check if file exists for specified period
'------------------------------------------------------------------------------
Public Function FileExistsForPeriod(logicalName As String, _
                                    Optional offset As ePeriodOffset = poCurrentMonth) As Boolean
    Dim filePath As String
    
    filePath = GetInputFilePathEx(logicalName, offset)
    FileExistsForPeriod = FileExistsSafe(filePath)
End Function

'------------------------------------------------------------------------------
' Function: FileExistsSafe
' Purpose: Safely check if file exists (handles paths with spaces and special chars)
' Note: Uses FileSystemObject for more reliable file existence check
'------------------------------------------------------------------------------
Public Function FileExistsSafe(filePath As String) As Boolean
    Dim fso As Object
    
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then
        ' Fallback to Dir if FSO not available
        Err.Clear
        FileExistsSafe = (Len(Dir(filePath)) > 0)
        Exit Function
    End If
    On Error GoTo 0
    
    FileExistsSafe = fso.FileExists(filePath)
    Set fso = Nothing
End Function

'------------------------------------------------------------------------------
' Function: GetPeriodDescription
' Purpose: Get period description text (for logging and UI)
'------------------------------------------------------------------------------
Public Function GetPeriodDescription(offset As ePeriodOffset) As String
    Select Case offset
        Case poCurrentMonth
            GetPeriodDescription = "CurrentMonth"
        Case poPreviousMonth
            GetPeriodDescription = "PreviousMonth"
        Case poNextMonth
            GetPeriodDescription = "NextMonth"
        Case Else
            GetPeriodDescription = "UnknownPeriod"
    End Select
End Function

'------------------------------------------------------------------------------
' Sub: LogPathInfo
' Purpose: Log path information (for debugging)
'------------------------------------------------------------------------------
Public Sub LogPathInfo(logicalName As String, offset As ePeriodOffset)
    Dim filePath As String
    Dim periodDesc As String
    Dim exists As String
    
    filePath = GetInputFilePathEx(logicalName, offset)
    periodDesc = GetPeriodDescription(offset)
    exists = IIf(FileExistsSafe(filePath), "Exists", "NotFound")
    
    LogInfo "modPathService", "LogPathInfo", _
        logicalName & " (" & periodDesc & "): " & filePath & " [" & exists & "]"
End Sub


