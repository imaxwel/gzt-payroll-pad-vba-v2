VERSION 5.00
Begin VB.UserForm frmPayrollMain 
   Caption         =   "HK Payroll Automation"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11000
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.ListBox lstInputFiles
      ColumnCount    =   6
      Left           =   120
      Top            =   480
      Width          =   10400
      Height         =   4200
      MultiSelect    =   0
   End
   Begin MSForms.CommandButton btnRefresh
      Caption        =   "Refresh FilePaths"
      Left           =   120
      Top            =   4920
      Width          =   1500
      Height         =   360
   End
   Begin MSForms.CommandButton btnRunInput
      Caption        =   "Run Payroll Input"
      Left           =   2100
      Top            =   4920
      Width          =   1800
      Height         =   360
   End
   Begin MSForms.CommandButton btnRunValidation
      Caption        =   "Run Payroll Validation"
      Left           =   4200
      Top            =   4920
      Width          =   2200
      Height         =   360
   End
   Begin MSForms.ComboBox cmbMonth
      Left           =   8700
      Top            =   4920
      Width          =   1200
      Height         =   300
      Style          =   2  'Dropdown List
   End
   Begin MSForms.TextBox txtYear
      Left           =   8700
      Top            =   5280
      Width          =   1200
      Height         =   300
   End
   Begin MSForms.Label lblHeaderName
      Caption        =   "Name"
      Left           =   120
      Top            =   240
      Width          =   1600
      Height         =   240
      FontBold       =   True
   End
   Begin MSForms.Label lblHeaderKeyword
      Caption        =   "Keyword"
      Left           =   1800
      Top            =   240
      Width          =   1200
      Height         =   240
      FontBold       =   True
   End
   Begin MSForms.Label lblHeaderFilePath
      Caption        =   "FilePath"
      Left           =   3120
      Top            =   240
      Width          =   5200
      Height         =   240
      FontBold       =   True
   End
   Begin MSForms.Label lblHeaderFunction
      Caption        =   "Function"
      Left           =   8400
      Top            =   240
      Width          =   900
      Height         =   240
      FontBold       =   True
   End
   Begin MSForms.Label lblHeaderRun
      Caption        =   "Run"
      Left           =   9420
      Top            =   240
      Width          =   600
      Height         =   240
      FontBold       =   True
   End
   Begin MSForms.Label lblMonth
      Caption        =   "Month"
      Left           =   7920
      Top            =   4920
      Width          =   720
      Height         =   240
   End
   Begin MSForms.Label lblYear
      Caption        =   "Year"
      Left           =   7920
      Top            =   5280
      Width          =   720
      Height         =   240
   End
   Begin MSForms.Label lblStatus
      Caption        =   ""
      Left           =   120
      Top            =   5400
      Width          =   7600
      Height         =   360
      ForeColor      =   255
   End
End
Attribute VB_Name = "frmPayrollMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mItems As Collection
Private mIsRefreshed As Boolean

Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler
    
    mIsRefreshed = False
    InitPeriodControls
    LoadAndDisplayConfig
    
    Exit Sub
ErrHandler:
    LogError "frmPayrollMain", "UserForm_Initialize", Err.Number, Err.Description
    MsgBox "Failed to initialize form: " & Err.Description, vbCritical, "HK Payroll Automation"
End Sub

Private Sub btnRefresh_Click()
    On Error GoTo ErrHandler
    
    Dim configPath As String
    configPath = GetDefaultConfigPath()
    
    If mItems Is Nothing Then
        Set mItems = LoadInputFilesConfig(configPath)
    End If
    
    Dim inputFolder As String
    inputFolder = GetDefaultInputFolder()
    
    Dim payrollMonth As String
    payrollMonth = GetSelectedMonthString()
    
    ResolveInputFilePaths mItems, inputFolder, payrollMonth
    WriteBackFilePathsToConfig mItems, configPath
    PopulateListBox
    UpdateStatusLabel
    
    mIsRefreshed = True
    Exit Sub
ErrHandler:
    LogError "frmPayrollMain", "btnRefresh_Click", Err.Number, Err.Description
    MsgBox "Refresh failed: " & Err.Description, vbCritical, "HK Payroll Automation"
End Sub

Private Sub btnRunInput_Click()
    RunWithScope "PROCESS"
End Sub

Private Sub btnRunValidation_Click()
    RunWithScope "VALIDATION"
End Sub

Private Sub RunWithScope(scopeName As String)
    On Error GoTo ErrHandler
    
    If Not mIsRefreshed Then
        MsgBox "Please click Refresh FilePaths before running.", vbExclamation, "HK Payroll Automation"
        Exit Sub
    End If
    
    If HasBlockingErrors(scopeName) Then
        MsgBox "Mandatory files are missing or not unique. Please fix and refresh.", vbCritical, "HK Payroll Automation"
        Exit Sub
    End If
    
    Dim payrollMonth As String
    payrollMonth = GetSelectedMonthString()
    
    ' Update Runtime parameters for subprocess macros
    With ThisWorkbook.Worksheets("Runtime")
        .Range("PayrollMonth").Value = payrollMonth
        .Range("RunDate").Value = Date
    End With
    
    If scopeName = "PROCESS" Then
        LogInfo "frmPayrollMain", "RunWithScope", "Running Subprocess 1"
        Run_Subprocess1
    Else
        LogInfo "frmPayrollMain", "RunWithScope", "Running Subprocess 2"
        Run_Subprocess2
    End If
    
    Exit Sub
ErrHandler:
    LogError "frmPayrollMain", "RunWithScope", Err.Number, Err.Description
    MsgBox "Run failed: " & Err.Description, vbCritical, "HK Payroll Automation"
End Sub

Private Sub InitPeriodControls()
    Dim i As Long
    cmbMonth.Clear
    For i = 1 To 12
        cmbMonth.AddItem MonthName(i, False)
    Next i
    cmbMonth.ListIndex = Month(Date) - 1
    txtYear.Value = Year(Date)
End Sub

Private Sub LoadAndDisplayConfig()
    Dim configPath As String
    configPath = GetDefaultConfigPath()
    Set mItems = LoadInputFilesConfig(configPath)
    PopulateListBox
    UpdateStatusLabel
End Sub

Private Sub PopulateListBox()
    Dim item As Object
    Dim rowIndex As Long
    
    lstInputFiles.Clear
    lstInputFiles.ColumnCount = 6
    lstInputFiles.ColumnWidths = "160 pt;100 pt;420 pt;70 pt;50 pt;0 pt"
    
    For Each item In mItems
        Dim displayName As String
        displayName = CStr(item("Name"))
        
        Select Case CLng(item("Status"))
            Case fsMissingMandatory
                displayName = "[MISSING] " & displayName
            Case fsNotUnique
                displayName = "[NOT UNIQUE] " & displayName
        End Select
        
        lstInputFiles.AddItem displayName
        rowIndex = lstInputFiles.ListCount - 1
        lstInputFiles.List(rowIndex, 1) = CStr(item("Keyword"))
        lstInputFiles.List(rowIndex, 2) = CStr(item("FilePath"))
        lstInputFiles.List(rowIndex, 3) = CStr(item("Function"))
        lstInputFiles.List(rowIndex, 4) = CStr(item("Run"))
        lstInputFiles.List(rowIndex, 5) = CStr(item("Status"))
    Next item
End Sub

Private Sub UpdateStatusLabel()
    Dim missingCount As Long, notUniqueCount As Long
    Dim item As Object
    
    missingCount = 0
    notUniqueCount = 0
    
    If Not mItems Is Nothing Then
        For Each item In mItems
            If CLng(item("Status")) = fsMissingMandatory Then missingCount = missingCount + 1
            If CLng(item("Status")) = fsNotUnique Then notUniqueCount = notUniqueCount + 1
        Next item
    End If
    
    If missingCount > 0 Or notUniqueCount > 0 Then
        lblStatus.Caption = "Issues found: Missing mandatory=" & missingCount & _
            ", Not unique=" & notUniqueCount & ". Please fix and refresh."
        lblStatus.ForeColor = RGB(255, 0, 0)
    Else
        lblStatus.Caption = "Ready. No blocking issues."
        lblStatus.ForeColor = RGB(0, 128, 0)
    End If
End Sub

Private Function HasBlockingErrors(scopeName As String) As Boolean
    Dim item As Object
    Dim fn As String
    
    HasBlockingErrors = False
    
    For Each item In mItems
        If UCase(CStr(item("Run"))) = "YES" Then
            fn = UCase(CStr(item("Function")))
            If scopeName = "PROCESS" Then
                If fn = "PROCESS" Or fn = "BOTH" Then
                    If CLng(item("Status")) = fsMissingMandatory Or CLng(item("Status")) = fsNotUnique Then
                        HasBlockingErrors = True
                        Exit Function
                    End If
                End If
            ElseIf scopeName = "VALIDATION" Then
                If fn = "VALIDATION" Or fn = "BOTH" Then
                    If CLng(item("Status")) = fsMissingMandatory Or CLng(item("Status")) = fsNotUnique Then
                        HasBlockingErrors = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next item
End Function

Private Function GetSelectedMonthString() As String
    Dim yearVal As Long, monthVal As Long
    yearVal = CLng(Val(txtYear.Value))
    monthVal = CLng(cmbMonth.ListIndex + 1)
    GetSelectedMonthString = GetSelectedPayrollMonth(yearVal, monthVal)
End Function

