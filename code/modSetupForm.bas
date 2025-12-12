Attribute VB_Name = "modSetupForm"
'==============================================================================
' Module: modSetupForm
' Purpose: Programmatically create the frmPayrollMain UserForm
' Usage: Run CreatePayrollForm macro once to create the form
' Note: Requires "Trust access to the VBA project object model" enabled
'       (File > Options > Trust Center > Trust Center Settings > Macro Settings)
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Sub: CreatePayrollForm
' Purpose: Create frmPayrollMain UserForm programmatically
'------------------------------------------------------------------------------
Public Sub CreatePayrollForm()
    On Error GoTo ErrHandler
    
    Dim vbProj As Object
    Dim vbComp As Object
    Dim frm As Object
    Dim ctrl As Object
    Dim codeModule As Object
    
    Set vbProj = ThisWorkbook.VBProject
    
    ' Check if form already exists
    On Error Resume Next
    Set vbComp = vbProj.VBComponents("frmPayrollMain")
    On Error GoTo ErrHandler
    
    If Not vbComp Is Nothing Then
        If MsgBox("frmPayrollMain already exists. Delete and recreate?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
        vbProj.VBComponents.Remove vbComp
    End If
    
    ' Create new UserForm
    Set vbComp = vbProj.VBComponents.Add(3) ' vbext_ct_MSForm = 3
    vbComp.Name = "frmPayrollMain"
    
    Set frm = vbComp.Designer
    
    ' Set form properties
    frm.Caption = "HK Payroll Automation"
    frm.Width = 660
    frm.Height = 360
    
    ' Add ListBox
    Set ctrl = frm.Controls.Add("Forms.ListBox.1", "lstInputFiles")
    With ctrl
        .Left = 12
        .Top = 18
        .Width = 624
        .Height = 270
        .ColumnCount = 6
        .ColumnWidths = "160;100;300;70;50;0"
        .ColumnHeads = True
    End With
    
    ' Add Refresh button
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnRefresh")
    With ctrl
        .Left = 12
        .Top = 294
        .Width = 90
        .Height = 24
        .Caption = "Refresh FilePaths"
    End With
    
    ' Add Run Input button
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnRunInput")
    With ctrl
        .Left = 108
        .Top = 294
        .Width = 108
        .Height = 24
        .Caption = "Run Payroll Input"
    End With
    
    ' Add Run Validation button
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "btnRunValidation")
    With ctrl
        .Left = 222
        .Top = 294
        .Width = 132
        .Height = 24
        .Caption = "Run Payroll Validation"
    End With
    
    ' Add Month label and combo
    AddLabel frm, "lblMonth", "Month:", 420, 297, 36, 15, False
    Set ctrl = frm.Controls.Add("Forms.ComboBox.1", "cmbMonth")
    With ctrl
        .Left = 456
        .Top = 294
        .Width = 72
        .Height = 18
        .Style = 2 ' fmStyleDropDownList
    End With
    
    ' Add Year label and textbox
    AddLabel frm, "lblYear", "Year:", 534, 297, 30, 15, False
    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtYear")
    With ctrl
        .Left = 564
        .Top = 294
        .Width = 48
        .Height = 18
    End With
    
    ' Add Status label
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblStatus")
    With ctrl
        .Left = 12
        .Top = 324
        .Width = 456
        .Height = 18
        .Caption = ""
        .ForeColor = RGB(255, 0, 0)
    End With
    
    ' Add form code
    Set codeModule = vbComp.codeModule
    AddFormCode codeModule
    
    MsgBox "frmPayrollMain created successfully!" & vbCrLf & vbCrLf & _
           "You can now run ShowPayrollForm or startformMain macro.", _
           vbInformation, "Setup Complete"
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error creating form: " & Err.Description & vbCrLf & vbCrLf & _
           "Make sure 'Trust access to the VBA project object model' is enabled:" & vbCrLf & _
           "File > Options > Trust Center > Trust Center Settings > Macro Settings", _
           vbCritical, "Setup Error"
End Sub

Private Sub AddLabel(frm As Object, ctrlName As String, captionText As String, _
                     l As Single, t As Single, w As Single, h As Single, isBold As Boolean)
    Dim ctrl As Object
    Set ctrl = frm.Controls.Add("Forms.Label.1", ctrlName)
    With ctrl
        .Left = l
        .Top = t
        .Width = w
        .Height = h
        .Caption = captionText
        If isBold Then .Font.Bold = True
    End With
End Sub

Private Sub AddFormCode(codeModule As Object)
    Dim code As String
    
    ' Clear existing code
    If codeModule.CountOfLines > 0 Then
        codeModule.DeleteLines 1, codeModule.CountOfLines
    End If
    
    code = "Option Explicit" & vbCrLf & vbCrLf
    code = code & "Private mItems As Collection" & vbCrLf
    code = code & "Private mIsRefreshed As Boolean" & vbCrLf & vbCrLf
    
    ' UserForm_Initialize
    code = code & "Private Sub UserForm_Initialize()" & vbCrLf
    code = code & "    On Error GoTo ErrHandler" & vbCrLf
    code = code & "    mIsRefreshed = False" & vbCrLf
    code = code & "    ConfigureInputFilesTable" & vbCrLf
    code = code & "    InitPeriodControls" & vbCrLf
    code = code & "    LoadAndDisplayConfig" & vbCrLf
    code = code & "    Exit Sub" & vbCrLf
    code = code & "ErrHandler:" & vbCrLf
    code = code & "    LogError ""frmPayrollMain"", ""UserForm_Initialize"", Err.Number, Err.Description" & vbCrLf
    code = code & "    MsgBox ""Failed to initialize form: "" & Err.Description, vbCritical" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    codeModule.AddFromString code
    
    ' Add remaining code in parts to avoid string length issues
    AddFormCodePart2 codeModule
    AddFormCodePart3 codeModule
    AddFormCodePart4 codeModule
End Sub

Private Sub AddFormCodePart2(codeModule As Object)
    Dim code As String
    
    ' btnRefresh_Click
    code = "Private Sub btnRefresh_Click()" & vbCrLf
    code = code & "    On Error GoTo ErrHandler" & vbCrLf
    code = code & "    Dim configPath As String" & vbCrLf
    code = code & "    configPath = GetDefaultConfigPath()" & vbCrLf
    code = code & "    If mItems Is Nothing Then" & vbCrLf
    code = code & "        Set mItems = LoadInputFilesConfig(configPath)" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    Dim inputFolder As String" & vbCrLf
    code = code & "    inputFolder = GetDefaultInputFolder()" & vbCrLf
    code = code & "    Dim payrollMonth As String" & vbCrLf
    code = code & "    payrollMonth = GetSelectedMonthString()" & vbCrLf
    code = code & "    ResolveInputFilePaths mItems, inputFolder, payrollMonth" & vbCrLf
    code = code & "    WriteBackFilePathsToConfig mItems, configPath" & vbCrLf
    code = code & "    PopulateListBox" & vbCrLf
    code = code & "    UpdateStatusLabel" & vbCrLf
    code = code & "    mIsRefreshed = True" & vbCrLf
    code = code & "    Exit Sub" & vbCrLf
    code = code & "ErrHandler:" & vbCrLf
    code = code & "    LogError ""frmPayrollMain"", ""btnRefresh_Click"", Err.Number, Err.Description" & vbCrLf
    code = code & "    MsgBox ""Refresh failed: "" & Err.Description, vbCritical" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' Button click handlers
    code = code & "Private Sub btnRunInput_Click()" & vbCrLf
    code = code & "    RunWithScope ""PROCESS""" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    code = code & "Private Sub btnRunValidation_Click()" & vbCrLf
    code = code & "    RunWithScope ""VALIDATION""" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    codeModule.AddFromString code
End Sub

Private Sub AddFormCodePart3(codeModule As Object)
    Dim code As String
    
    ' RunWithScope
    code = "Private Sub RunWithScope(scopeName As String)" & vbCrLf
    code = code & "    On Error GoTo ErrHandler" & vbCrLf
    code = code & "    If Not mIsRefreshed Then" & vbCrLf
    code = code & "        MsgBox ""Please click Refresh FilePaths before running."", vbExclamation" & vbCrLf
    code = code & "        Exit Sub" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    Dim blockingDetails As String" & vbCrLf
    code = code & "    blockingDetails = GetBlockingErrorDetails(scopeName)" & vbCrLf
    code = code & "    If blockingDetails <> """" Then" & vbCrLf
    code = code & "        MsgBox ""Mandatory files are missing or not unique:"" & blockingDetails, vbCritical" & vbCrLf
    code = code & "        Exit Sub" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    Dim payrollMonth As String" & vbCrLf
    code = code & "    payrollMonth = GetSelectedMonthString()" & vbCrLf
    code = code & "    With ThisWorkbook.Worksheets(""Runtime"")" & vbCrLf
    code = code & "        .Range(""PayrollMonth"").Value = payrollMonth" & vbCrLf
    code = code & "        .Range(""RunDate"").Value = Date" & vbCrLf
    code = code & "    End With" & vbCrLf
    code = code & "    If scopeName = ""PROCESS"" Then" & vbCrLf
    code = code & "        Run_Subprocess1" & vbCrLf
    code = code & "    Else" & vbCrLf
    code = code & "        Run_Subprocess2" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    Exit Sub" & vbCrLf
    code = code & "ErrHandler:" & vbCrLf
    code = code & "    LogError ""frmPayrollMain"", ""RunWithScope"", Err.Number, Err.Description" & vbCrLf
    code = code & "    MsgBox ""Run failed: "" & Err.Description, vbCritical" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' InitPeriodControls
    code = code & "Private Sub InitPeriodControls()" & vbCrLf
    code = code & "    Dim i As Long" & vbCrLf
    code = code & "    cmbMonth.Clear" & vbCrLf
    code = code & "    For i = 1 To 12" & vbCrLf
    code = code & "        cmbMonth.AddItem Format(i, ""00"") & "" - "" & MonthName(i, True)" & vbCrLf
    code = code & "    Next i" & vbCrLf
    code = code & "    cmbMonth.ListIndex = Month(Date) - 1" & vbCrLf
    code = code & "    txtYear.Value = CStr(Year(Date))" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' LoadAndDisplayConfig
    code = code & "Private Sub LoadAndDisplayConfig()" & vbCrLf
    code = code & "    On Error GoTo ErrHandler" & vbCrLf
    code = code & "    Dim configPath As String" & vbCrLf
    code = code & "    configPath = GetDefaultConfigPath()" & vbCrLf
    code = code & "    Set mItems = LoadInputFilesConfig(configPath)" & vbCrLf
    code = code & "    PopulateListBox" & vbCrLf
    code = code & "    UpdateStatusLabel" & vbCrLf
    code = code & "    Exit Sub" & vbCrLf
    code = code & "ErrHandler:" & vbCrLf
    code = code & "    LogError ""frmPayrollMain"", ""LoadAndDisplayConfig"", Err.Number, Err.Description" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    codeModule.AddFromString code
End Sub

Private Sub AddFormCodePart4(codeModule As Object)
    Dim code As String
    
    ' ConfigureInputFilesTable
    code = "Private Sub ConfigureInputFilesTable()" & vbCrLf
    code = code & "    HideInputFilesHeaderLabels" & vbCrLf & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    With lstInputFiles" & vbCrLf
    code = code & "        .Top = 18" & vbCrLf
    code = code & "        .Height = 270" & vbCrLf
    code = code & "        .ColumnCount = 6" & vbCrLf
    code = code & "        .ColumnWidths = ""160;100;300;70;50;0""" & vbCrLf
    code = code & "    End With" & vbCrLf
    code = code & "    On Error GoTo 0" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    ' HideInputFilesHeaderLabels
    code = code & "Private Sub HideInputFilesHeaderLabels()" & vbCrLf
    code = code & "    Dim ctrlName As Variant" & vbCrLf
    code = code & "    For Each ctrlName In Array( _" & vbCrLf
    code = code & "        ""lblHeaderName"", _" & vbCrLf
    code = code & "        ""lblHeaderKeyword"", _" & vbCrLf
    code = code & "        ""lblHeaderFilePath"", _" & vbCrLf
    code = code & "        ""lblHeaderFunction"", _" & vbCrLf
    code = code & "        ""lblHeaderRun"" _" & vbCrLf
    code = code & "    )" & vbCrLf
    code = code & "        On Error Resume Next" & vbCrLf
    code = code & "        Me.Controls(CStr(ctrlName)).Visible = False" & vbCrLf
    code = code & "        On Error GoTo 0" & vbCrLf
    code = code & "    Next ctrlName" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf

    ' PopulateListBox
    code = code & "Private Sub PopulateListBox()" & vbCrLf
    code = code & "    On Error GoTo ErrHandler" & vbCrLf
    code = code & "    Dim ws As Worksheet" & vbCrLf
    code = code & "    Set ws = ThisWorkbook.Worksheets(""Runtime"")" & vbCrLf & vbCrLf
    code = code & "    Dim totalRows As Long" & vbCrLf
    code = code & "    If mItems Is Nothing Then" & vbCrLf
    code = code & "        totalRows = 2" & vbCrLf
    code = code & "    Else" & vbCrLf
    code = code & "        totalRows = mItems.Count + 1" & vbCrLf
    code = code & "        If totalRows < 2 Then totalRows = 2" & vbCrLf
    code = code & "    End If" & vbCrLf & vbCrLf
    code = code & "    Dim dataArr() As Variant" & vbCrLf
    code = code & "    ReDim dataArr(1 To totalRows, 1 To 6)" & vbCrLf & vbCrLf
    code = code & "    dataArr(1, 1) = ""Name""" & vbCrLf
    code = code & "    dataArr(1, 2) = ""Keyword""" & vbCrLf
    code = code & "    dataArr(1, 3) = ""FilePath""" & vbCrLf
    code = code & "    dataArr(1, 4) = ""Function""" & vbCrLf
    code = code & "    dataArr(1, 5) = ""Run""" & vbCrLf
    code = code & "    dataArr(1, 6) = ""Status""" & vbCrLf & vbCrLf
    code = code & "    Dim item As Object, displayName As String, writeRow As Long" & vbCrLf
    code = code & "    If Not mItems Is Nothing Then" & vbCrLf
    code = code & "        writeRow = 2" & vbCrLf
    code = code & "        For Each item In mItems" & vbCrLf
    code = code & "            displayName = CStr(item(""Name""))" & vbCrLf
    code = code & "            Select Case CLng(item(""Status""))" & vbCrLf
    code = code & "                Case fsMissingMandatory: displayName = ""[MISSING] "" & displayName" & vbCrLf
    code = code & "                Case fsNotUnique: displayName = ""[NOT UNIQUE] "" & displayName" & vbCrLf
    code = code & "            End Select" & vbCrLf & vbCrLf
    code = code & "            dataArr(writeRow, 1) = displayName" & vbCrLf
    code = code & "            dataArr(writeRow, 2) = CStr(item(""Keyword""))" & vbCrLf
    code = code & "            dataArr(writeRow, 3) = CStr(item(""FilePath""))" & vbCrLf
    code = code & "            dataArr(writeRow, 4) = CStr(item(""Function""))" & vbCrLf
    code = code & "            dataArr(writeRow, 5) = CStr(item(""Run""))" & vbCrLf
    code = code & "            dataArr(writeRow, 6) = CStr(item(""Status""))" & vbCrLf
    code = code & "            writeRow = writeRow + 1" & vbCrLf
    code = code & "        Next item" & vbCrLf
    code = code & "    End If" & vbCrLf & vbCrLf
    code = code & "    ws.Range(""AA1"").Resize(totalRows, 6).Value = dataArr" & vbCrLf
    code = code & "    ws.Range(""AA1"").Resize(1, 6).Font.Bold = True" & vbCrLf & vbCrLf
    code = code & "    With lstInputFiles" & vbCrLf
    code = code & "        .RowSource = """"" & vbCrLf
    code = code & "        .ColumnCount = 6" & vbCrLf
    code = code & "        .ColumnWidths = ""160;100;300;70;50;0""" & vbCrLf
    code = code & "        .ColumnHeads = True" & vbCrLf
    code = code & "        .RowSource = ""'"" & ws.Name & ""'!AA2:AF"" & totalRows" & vbCrLf
    code = code & "    End With" & vbCrLf
    code = code & "    Exit Sub" & vbCrLf
    code = code & "ErrHandler:" & vbCrLf
    code = code & "    LogError ""frmPayrollMain"", ""PopulateListBox"", Err.Number, Err.Description" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' UpdateStatusLabel
    code = code & "Private Sub UpdateStatusLabel()" & vbCrLf
    code = code & "    Dim missingCount As Long, notUniqueCount As Long, item As Object" & vbCrLf
    code = code & "    If Not mItems Is Nothing Then" & vbCrLf
    code = code & "        For Each item In mItems" & vbCrLf
    code = code & "            If CLng(item(""Status"")) = fsMissingMandatory Then missingCount = missingCount + 1" & vbCrLf
    code = code & "            If CLng(item(""Status"")) = fsNotUnique Then notUniqueCount = notUniqueCount + 1" & vbCrLf
    code = code & "        Next item" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    If missingCount > 0 Or notUniqueCount > 0 Then" & vbCrLf
    code = code & "        lblStatus.Caption = ""Issues: Missing="" & missingCount & "", Not unique="" & notUniqueCount" & vbCrLf
    code = code & "        lblStatus.ForeColor = RGB(255, 0, 0)" & vbCrLf
    code = code & "    Else" & vbCrLf
    code = code & "        lblStatus.Caption = ""Ready. No blocking issues.""" & vbCrLf
    code = code & "        lblStatus.ForeColor = RGB(0, 128, 0)" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' HasBlockingErrors
    code = code & "Private Function HasBlockingErrors(scopeName As String) As Boolean" & vbCrLf
    code = code & "    Dim item As Object, fn As String, st As Long" & vbCrLf
    code = code & "    HasBlockingErrors = False" & vbCrLf
    code = code & "    If mItems Is Nothing Then Exit Function" & vbCrLf
    code = code & "    For Each item In mItems" & vbCrLf
    code = code & "        If UCase(CStr(item(""Run""))) = ""YES"" Then" & vbCrLf
    code = code & "            fn = UCase(CStr(item(""Function"")))" & vbCrLf
    code = code & "            st = CLng(item(""Status""))" & vbCrLf
    code = code & "            If scopeName = ""PROCESS"" And (fn = ""PROCESS"" Or fn = ""BOTH"") Then" & vbCrLf
    code = code & "                If st = fsMissingMandatory Or st = fsNotUnique Then HasBlockingErrors = True: Exit Function" & vbCrLf
    code = code & "            ElseIf scopeName = ""VALIDATION"" And (fn = ""VALIDATION"" Or fn = ""BOTH"") Then" & vbCrLf
    code = code & "                If st = fsMissingMandatory Or st = fsNotUnique Then HasBlockingErrors = True: Exit Function" & vbCrLf
    code = code & "            End If" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next item" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf
    
    ' GetBlockingErrorDetails
    code = code & "Private Function GetBlockingErrorDetails(scopeName As String) As String" & vbCrLf
    code = code & "    Dim item As Object, fn As String, st As Long, details As String, inScope As Boolean" & vbCrLf
    code = code & "    details = """"" & vbCrLf
    code = code & "    If mItems Is Nothing Then" & vbCrLf
    code = code & "        GetBlockingErrorDetails = """"" & vbCrLf
    code = code & "        Exit Function" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    For Each item In mItems" & vbCrLf
    code = code & "        If UCase(CStr(item(""Run""))) = ""YES"" Then" & vbCrLf
    code = code & "            fn = UCase(CStr(item(""Function"")))" & vbCrLf
    code = code & "            st = CLng(item(""Status""))" & vbCrLf
    code = code & "            inScope = False" & vbCrLf
    code = code & "            If scopeName = ""PROCESS"" And (fn = ""PROCESS"" Or fn = ""BOTH"") Then inScope = True" & vbCrLf
    code = code & "            If scopeName = ""VALIDATION"" And (fn = ""VALIDATION"" Or fn = ""BOTH"") Then inScope = True" & vbCrLf
    code = code & "            If inScope Then" & vbCrLf
    code = code & "                If st = fsMissingMandatory Then" & vbCrLf
    code = code & "                    details = details & vbCrLf & "" - Missing: "" & CStr(item(""Name""))" & vbCrLf
    code = code & "                ElseIf st = fsNotUnique Then" & vbCrLf
    code = code & "                    details = details & vbCrLf & "" - Not unique: "" & CStr(item(""Name""))" & vbCrLf
    code = code & "                End If" & vbCrLf
    code = code & "            End If" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next item" & vbCrLf
    code = code & "    GetBlockingErrorDetails = details" & vbCrLf
    code = code & "End Function" & vbCrLf & vbCrLf
    
    ' GetSelectedMonthString
    code = code & "Private Function GetSelectedMonthString() As String" & vbCrLf
    code = code & "    Dim yearVal As Long, monthVal As Long" & vbCrLf
    code = code & "    On Error Resume Next" & vbCrLf
    code = code & "    yearVal = CLng(Val(txtYear.Value))" & vbCrLf
    code = code & "    monthVal = CLng(cmbMonth.ListIndex + 1)" & vbCrLf
    code = code & "    GetSelectedMonthString = GetSelectedPayrollMonth(yearVal, monthVal)" & vbCrLf
    code = code & "End Function" & vbCrLf
    
    codeModule.AddFromString code
End Sub
