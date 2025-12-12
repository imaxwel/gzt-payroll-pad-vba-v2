VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPayrollMain 
   Caption         =   "HK Payroll Automation"
   ClientHeight    =   9210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18390
   OleObjectBlob   =   "frmPayrollMain.frx":0000
   StartUpPosition =   1  '所有者中心
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
      MsgBox "Failed to initialize form: " & Err.Description, vbCritical
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
      MsgBox "Refresh failed: " & Err.Description, vbCritical
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
          MsgBox "Please click Refresh FilePaths before running.", vbExclamation
          Exit Sub
      End If

      Dim blockingDetails As String
      blockingDetails = GetBlockingErrorDetails(scopeName)
      If blockingDetails <> "" Then
          MsgBox "Mandatory files are missing or not unique:" & blockingDetails, vbCritical
          Exit Sub
      End If

      Dim payrollMonth As String
      payrollMonth = GetSelectedMonthString()
      With ThisWorkbook.Worksheets("Runtime")
          .Range("PayrollMonth").value = payrollMonth
          .Range("RunDate").value = Date
      End With

      If scopeName = "PROCESS" Then
          Run_Subprocess1
      Else
          Run_Subprocess2
      End If

      Exit Sub
ErrHandler:
      LogError "frmPayrollMain", "RunWithScope", Err.Number, Err.Description
      MsgBox "Run failed: " & Err.Description, vbCritical
  End Sub

  Private Sub InitPeriodControls()
      Dim i As Long
      cmbMonth.Clear
      For i = 1 To 12
          cmbMonth.AddItem Format(i, "00") & " - " & MonthName(i, True)
      Next i
      cmbMonth.ListIndex = Month(Date) - 1
      txtYear.value = CStr(Year(Date))
  End Sub

  Private Sub LoadAndDisplayConfig()
      On Error GoTo ErrHandler
      Dim configPath As String
      configPath = GetDefaultConfigPath()
      Set mItems = LoadInputFilesConfig(configPath)
      PopulateListBox
      UpdateStatusLabel
      Exit Sub
ErrHandler:
      LogError "frmPayrollMain", "LoadAndDisplayConfig", Err.Number, Err.Description
  End Sub

  Private Sub PopulateListBox()
      On Error GoTo ErrHandler
      Dim item As Object, rowIndex As Long, displayName As String
      lstInputFiles.Clear
      If mItems Is Nothing Then Exit Sub

      For Each item In mItems
          displayName = CStr(item("Name"))
          Select Case CLng(item("Status"))
              Case fsMissingMandatory: displayName = "[MISSING] " & displayName
              Case fsNotUnique: displayName = "[NOT UNIQUE] " & displayName
          End Select

          lstInputFiles.AddItem displayName
          rowIndex = lstInputFiles.ListCount - 1
          lstInputFiles.List(rowIndex, 1) = CStr(item("Keyword"))
          lstInputFiles.List(rowIndex, 2) = CStr(item("FilePath"))
          lstInputFiles.List(rowIndex, 3) = CStr(item("Function"))
          lstInputFiles.List(rowIndex, 4) = CStr(item("Run"))
          lstInputFiles.List(rowIndex, 5) = CStr(item("Status"))
      Next item

      Exit Sub
ErrHandler:
      LogError "frmPayrollMain", "PopulateListBox", Err.Number, Err.Description
  End Sub

  Private Sub UpdateStatusLabel()
      Dim missingCount As Long, notUniqueCount As Long, item As Object

      If Not mItems Is Nothing Then
          For Each item In mItems
              If CLng(item("Status")) = fsMissingMandatory Then missingCount = missingCount + 1
              If CLng(item("Status")) = fsNotUnique Then notUniqueCount = notUniqueCount + 1
          Next item
      End If

      If missingCount > 0 Or notUniqueCount > 0 Then
          lblStatus.Caption = "Issues: Missing=" & missingCount & ", Not unique=" & notUniqueCount
          lblStatus.ForeColor = RGB(255, 0, 0)
      Else
          lblStatus.Caption = "Ready. No blocking issues."
          lblStatus.ForeColor = RGB(0, 128, 0)
      End If
  End Sub

  Private Function HasBlockingErrors(scopeName As String) As Boolean
      Dim item As Object, fn As String, st As Long
      HasBlockingErrors = False
      If mItems Is Nothing Then Exit Function

      For Each item In mItems
          If UCase(CStr(item("Run"))) = "YES" Then
              fn = UCase(CStr(item("Function")))
              st = CLng(item("Status"))

              If scopeName = "PROCESS" And (fn = "PROCESS" Or fn = "BOTH") Then
                  If st = fsMissingMandatory Or st = fsNotUnique Then HasBlockingErrors = True: Exit Function
              ElseIf scopeName = "VALIDATION" And (fn = "VALIDATION" Or fn = "BOTH") Then
                  If st = fsMissingMandatory Or st = fsNotUnique Then HasBlockingErrors = True: Exit Function
              End If
          End If
      Next item
  End Function

  Private Function GetBlockingErrorDetails(scopeName As String) As String
      Dim item As Object, fn As String, st As Long, details As String, inScope As Boolean
      details = ""

      If mItems Is Nothing Then
          GetBlockingErrorDetails = ""
          Exit Function
      End If

      For Each item In mItems
          If UCase(CStr(item("Run"))) = "YES" Then
              fn = UCase(CStr(item("Function")))
              st = CLng(item("Status"))
              inScope = False

              If scopeName = "PROCESS" And (fn = "PROCESS" Or fn = "BOTH") Then inScope = True
              If scopeName = "VALIDATION" And (fn = "VALIDATION" Or fn = "BOTH") Then inScope = True

              If inScope Then
                  If st = fsMissingMandatory Then
                      details = details & vbCrLf & " - Missing: " & CStr(item("Name"))
                  ElseIf st = fsNotUnique Then
                      details = details & vbCrLf & " - Not unique: " & CStr(item("Name"))
                  End If
              End If
          End If
      Next item

      GetBlockingErrorDetails = details
  End Function

  Private Function GetSelectedMonthString() As String
      Dim yearVal As Long, monthVal As Long
      On Error Resume Next
      yearVal = CLng(Val(txtYear.value))
      monthVal = CLng(cmbMonth.ListIndex + 1)
      GetSelectedMonthString = GetSelectedPayrollMonth(yearVal, monthVal)
  End Function


