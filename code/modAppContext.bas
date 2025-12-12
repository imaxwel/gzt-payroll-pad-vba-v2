Attribute VB_Name = "modAppContext"
'==============================================================================
' Module: modAppContext
' Purpose: Global application context and type definitions
' Description: Defines shared types and the single global context variable
'              used by both Subprocess 1 and Subprocess 2
'==============================================================================
Option Explicit

' Run parameters passed from PAD or config
Public Type tRunParams
    inputFolder As String
    OutputFolder As String
    ConfigFolder As String
    payrollMonth As String      ' "YYYYMM" format
    RunDate As Date
    LogFolder As String
End Type

' Payroll calendar context derived from config
Public Type tPayrollContext
    payrollMonth As String      ' "YYYYMM"
    monthStart As Date          ' First day of payroll month
    monthEnd As Date            ' Last day of payroll month
    prevMonthStart As Date      ' First day of previous month
    prevMonthEnd As Date        ' Last day of previous month
    payDate As Date             ' Pay date for this month
    PreviousCutoff As Date      ' Previous month cutoff date
    currentCutoff As Date       ' Current month cutoff date
    CalendarDaysCurrentMonth As Long
    CalendarDaysPrevMonth As Long
End Type

' Date span for cross-month splitting
Public Type tDateSpan
    startDate As Date
    endDate As Date
    YearMonth As String         ' "YYYYMM"
    days As Double              ' Number of days (semantics depend on leave type)
End Type

' Application-wide shared state
Public Type tAppContext
    RunParams As tRunParams
    Payroll As tPayrollContext
    ' Common mappings (Scripting.Dictionary objects)
    DictWeinToEmpId As Object
    DictEmpIdToWein As Object
    DictEmpCodeToWein As Object
    DictWeinToEmpCode As Object
    ' Config workbook reference
    configWb As Workbook
    ExtraTableWb As Workbook    ' 额外表
    FlexiOutputWb As Workbook   ' Flexi output workbook for SP1
    ' Status flags
    IsInitialised As Boolean
End Type

' The SINGLE global variable for application context
Public G As tAppContext

'------------------------------------------------------------------------------
' Sub: InitAppContext
' Purpose: Initialize the global application context
' Parameters:
'   p - Run parameters structure
'------------------------------------------------------------------------------
Public Sub InitAppContext(p As tRunParams)
    On Error GoTo ErrHandler
    
    ' Always reset first
    ResetAppContext
    
    ' Store run parameters
    G.RunParams = p
    
    ' Load payroll calendar from config
    G.Payroll = GetPayrollContext(p.payrollMonth)
    
    ' Build shared employee mappings
    BuildEmployeeMappings
    
    G.IsInitialised = True
    
    LogInfo "modAppContext", "InitAppContext", "Application context initialized for payroll month: " & p.payrollMonth
    Exit Sub
    
ErrHandler:
    LogError "modAppContext", "InitAppContext", Err.Number, Err.Description
    Err.Raise Err.Number, "InitAppContext", Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: ResetAppContext
' Purpose: Reset/clear the global application context
'------------------------------------------------------------------------------
Public Sub ResetAppContext()
    G.IsInitialised = False
    Set G.DictWeinToEmpId = Nothing
    Set G.DictEmpIdToWein = Nothing
    Set G.DictEmpCodeToWein = Nothing
    Set G.DictWeinToEmpCode = Nothing
    
    ' Close config workbooks if open
    On Error Resume Next
    If Not G.configWb Is Nothing Then
        G.configWb.Close SaveChanges:=False
        Set G.configWb = Nothing
    End If
    If Not G.ExtraTableWb Is Nothing Then
        G.ExtraTableWb.Close SaveChanges:=False
        Set G.ExtraTableWb = Nothing
    End If
    If Not G.FlexiOutputWb Is Nothing Then
        G.FlexiOutputWb.Close SaveChanges:=False
        Set G.FlexiOutputWb = Nothing
    End If
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Function: EnsureInitialised
' Purpose: Check if context is initialized, raise error if not
'------------------------------------------------------------------------------
Public Sub EnsureInitialised()
    If Not G.IsInitialised Then
        Err.Raise vbObjectError + 1000, "EnsureInitialised", _
            "Application context not initialized. Call InitAppContext first."
    End If
End Sub
