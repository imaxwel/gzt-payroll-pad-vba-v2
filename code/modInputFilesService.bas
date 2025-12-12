Attribute VB_Name = "modInputFilesService"
'==============================================================================
' Module: modInputFilesService
' Purpose: Load Input Files configuration, resolve runtime file paths,
'          and write back resolved paths to config.xlsx.
' Notes:
'   - All objects are late-bound (CreateObject).
'   - No Chinese characters in new code/comments.
'==============================================================================
Option Explicit

Public Enum eFileStatus
    fsOk = 0
    fsMissingMandatory = 1
    fsNotUnique = 2
    fsMissingOptional = 3
End Enum

'------------------------------------------------------------------------------
' Function: GetDefaultConfigPath
' Purpose: Resolve config.xlsx path from Runtime.ConfigFolder, fallback to
'          workbook-relative config\config.xlsx.
'------------------------------------------------------------------------------
Public Function GetDefaultConfigPath() As String
    Dim p As tRunParams
    Dim configPath As String
    
    On Error Resume Next
    p = LoadRunParamsFromWorkbook()
    On Error GoTo 0
    
    If Trim(p.ConfigFolder) <> "" Then
        configPath = EnsureTrailingBackslash(p.ConfigFolder) & "config.xlsx"
    Else
        configPath = ThisWorkbook.path & "\config\config.xlsx"
    End If
    
    GetDefaultConfigPath = configPath
End Function

'------------------------------------------------------------------------------
' Function: GetDefaultInputFolder
' Purpose: Read InputFolder from Runtime sheet.
'------------------------------------------------------------------------------
Public Function GetDefaultInputFolder() As String
    Dim p As tRunParams
    On Error Resume Next
    p = LoadRunParamsFromWorkbook()
    On Error GoTo 0
    GetDefaultInputFolder = EnsureTrailingBackslash(p.inputFolder)
End Function

'------------------------------------------------------------------------------
' Function: LoadInputFilesConfig
' Purpose: Read Input Files sheet into a Collection of Dictionary items.
' Each item contains:
'   Name, Keyword, FilePath, Function, Run, Status, Matches,
'   ConfigRow, FilePathCol
'------------------------------------------------------------------------------
Public Function LoadInputFilesConfig(configPath As String) As Collection
    Dim items As Collection
    Dim wb As Workbook, ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim colName As Long, colKeyword As Long, colFilePath As Long, colFunction As Long, colRun As Long
    Dim r As Long
    
    Set items = New Collection
    
    On Error GoTo ErrHandler
    
    If Dir(configPath) = "" Then
        LogWarning "modInputFilesService", "LoadInputFilesConfig", _
            "Config workbook not found: " & configPath
        Set LoadInputFilesConfig = items
        Exit Function
    End If
    
    Set wb = Workbooks.Open(configPath, ReadOnly:=True, UpdateLinks:=False)
    On Error Resume Next
    Set ws = wb.Worksheets("Input Files")
    On Error GoTo ErrHandler
    
    If ws Is Nothing Then
        LogWarning "modInputFilesService", "LoadInputFilesConfig", _
            "Sheet 'Input Files' not found in config.xlsx"
        wb.Close SaveChanges:=False
        Set LoadInputFilesConfig = items
        Exit Function
    End If
    
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    colName = FindHeaderColumn(ws, "Name", lastCol)
    colKeyword = FindHeaderColumn(ws, "Keyword", lastCol)
    colFilePath = FindHeaderColumn(ws, "FilePath", lastCol)
    colFunction = FindHeaderColumn(ws, "Function", lastCol)
    colRun = FindHeaderColumn(ws, "Run", lastCol)
    
    If colName = 0 Or colKeyword = 0 Or colFilePath = 0 Or colFunction = 0 Or colRun = 0 Then
        LogWarning "modInputFilesService", "LoadInputFilesConfig", _
            "Required headers missing in 'Input Files' sheet"
        wb.Close SaveChanges:=False
        Set LoadInputFilesConfig = items
        Exit Function
    End If
    
    lastRow = ws.Cells(ws.Rows.count, colName).End(xlUp).row
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, colName).value)) = "" Then
            ' Stop at first blank Name row
            Exit For
        End If
        
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d("Name") = Trim(CStr(ws.Cells(r, colName).value))
        d("Keyword") = Trim(CStr(ws.Cells(r, colKeyword).value))
        d("FilePath") = Trim(CStr(ws.Cells(r, colFilePath).value))
        d("Function") = Trim(CStr(ws.Cells(r, colFunction).value))
        d("Run") = UCase(Trim(CStr(ws.Cells(r, colRun).value)))
        d("Status") = fsOk
        Set d("Matches") = New Collection
        d("ConfigRow") = r
        d("FilePathCol") = colFilePath
        
        items.Add d
    Next r
    
    wb.Close SaveChanges:=False
    Set LoadInputFilesConfig = items
    Exit Function
    
ErrHandler:
    LogError "modInputFilesService", "LoadInputFilesConfig", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set LoadInputFilesConfig = items
End Function

'------------------------------------------------------------------------------
' Sub: ResolveInputFilePaths
' Purpose: Scan InputFolder root first, then current payroll month folder,
'          update each item with Matches, FilePath and Status.
'------------------------------------------------------------------------------
Public Sub ResolveInputFilePaths(items As Collection, inputFolder As String, payrollMonth As String)
    Dim baseFolder As String
    Dim monthFolder As String
    Dim yearPart As String
    Dim item As Object
    
    On Error GoTo ErrHandler
    
    baseFolder = EnsureTrailingBackslash(inputFolder)
    yearPart = Left(payrollMonth, 4)
    monthFolder = EnsureTrailingBackslash(baseFolder & yearPart & "\Month\" & payrollMonth)
    
    For Each item In items
        Dim keyword As String
        keyword = CStr(item("Keyword"))
        
        Dim matches As Collection
        Set matches = FindMatches(baseFolder, keyword)
        If matches.count = 0 Then
            Set matches = FindMatches(monthFolder, keyword)
        End If
        
        Set item("Matches") = matches
        item("FilePath") = JoinMatches(matches)
        
        Dim allowMulti As Boolean
        allowMulti = IsMultiAllowed(item)
        
        If matches.count = 0 Then
            If UCase(CStr(item("Run"))) = "YES" Then
                item("Status") = fsMissingMandatory
            Else
                item("Status") = fsMissingOptional
            End If
        ElseIf matches.count > 1 And Not allowMulti Then
            item("Status") = fsNotUnique
        Else
            item("Status") = fsOk
        End If
    Next item
    
    Exit Sub
    
ErrHandler:
    LogError "modInputFilesService", "ResolveInputFilePaths", Err.Number, Err.Description
    Err.Raise Err.Number, "ResolveInputFilePaths", Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: WriteBackFilePathsToConfig
' Purpose: Update FilePath column in Input Files sheet based on resolved items.
'------------------------------------------------------------------------------
Public Sub WriteBackFilePathsToConfig(items As Collection, configPath As String)
    Dim wb As Workbook, ws As Worksheet
    Dim item As Object
    Dim lastCol As Long
    
    On Error GoTo ErrHandler
    
    If Dir(configPath) = "" Then Exit Sub
    
    Set wb = Workbooks.Open(configPath, ReadOnly:=False, UpdateLinks:=False)
    If wb.ReadOnly Then
        LogWarning "modInputFilesService", "WriteBackFilePathsToConfig", _
            "Config workbook opened read-only, cannot write back: " & configPath
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    On Error Resume Next
    Set ws = wb.Worksheets("Input Files")
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    For Each item In items
        Dim r As Long, c As Long
        r = CLng(item("ConfigRow"))
        c = CLng(item("FilePathCol"))
        ws.Cells(r, c).value = CStr(item("FilePath"))
        ApplyStatusHighlight ws, r, lastCol, CLng(item("Status"))
    Next item
    
    wb.Close SaveChanges:=True
    Exit Sub
    
ErrHandler:
    LogError "modInputFilesService", "WriteBackFilePathsToConfig", Err.Number, Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

'------------------------------------------------------------------------------
' Sub: ApplyStatusHighlight
' Purpose: Highlight a config row based on resolved file status.
'------------------------------------------------------------------------------
Private Sub ApplyStatusHighlight(ws As Worksheet, rowIndex As Long, lastCol As Long, statusVal As Long)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, lastCol))
    
    Select Case statusVal
        Case fsMissingMandatory
            rng.Interior.Color = RGB(255, 200, 200)
        Case fsNotUnique
            rng.Interior.Color = RGB(255, 255, 200)
        Case Else
            rng.Interior.ColorIndex = xlColorIndexNone
    End Select
End Sub

'------------------------------------------------------------------------------
' Function: GetSelectedPayrollMonth
' Purpose: Convert Year + MonthIndex into "YYYYMM".
'------------------------------------------------------------------------------
Public Function GetSelectedPayrollMonth(yearValue As Long, monthIndex As Long) As String
    GetSelectedPayrollMonth = CStr(yearValue) & Format(monthIndex, "00")
End Function

'------------------------------------------------------------------------------
' Function: EnsureTrailingBackslash
'------------------------------------------------------------------------------
Private Function EnsureTrailingBackslash(path As String) As String
    path = Trim(path)
    If path = "" Then
        EnsureTrailingBackslash = ""
    ElseIf Right(path, 1) <> "\" Then
        EnsureTrailingBackslash = path & "\"
    Else
        EnsureTrailingBackslash = path
    End If
End Function

'------------------------------------------------------------------------------
' Function: FindHeaderColumn
'------------------------------------------------------------------------------
Private Function FindHeaderColumn(ws As Worksheet, headerName As String, lastCol As Long) As Long
    Dim i As Long
    For i = 1 To lastCol
        If UCase(Trim(CStr(ws.Cells(1, i).value))) = UCase(headerName) Then
            FindHeaderColumn = i
            Exit Function
        End If
    Next i
    FindHeaderColumn = 0
End Function

'------------------------------------------------------------------------------
' Function: FindMatches
' Purpose: Enumerate *.xls* in a folder (non-recursive) and match by keyword.
'------------------------------------------------------------------------------
Private Function FindMatches(folder As String, keyword As String) As Collection
    Dim matches As Collection
    Dim f As String
    
    Set matches = New Collection
    
    If folder = "" Then
        Set FindMatches = matches
        Exit Function
    End If
    
    If Dir(folder, vbDirectory) = "" Then
        Set FindMatches = matches
        Exit Function
    End If
    
    folder = EnsureTrailingBackslash(folder)
    f = Dir(folder & "*.xls*")
    Do While f <> ""
        If Left$(f, 2) <> "~$" Then
            If IsKeywordMatch(f, keyword) Then
                matches.Add folder & f
            End If
        End If
        f = Dir()
    Loop
    
    Set FindMatches = matches
End Function

'------------------------------------------------------------------------------
' Function: IsKeywordMatch
' Purpose: Case-insensitive contains match after normalization.
'------------------------------------------------------------------------------
Private Function IsKeywordMatch(fileName As String, keyword As String) As Boolean
    Dim k As String, f As String
    k = NormalizeText(StripPlaceholders(keyword))
    f = NormalizeText(RemoveExtension(fileName))
    IsKeywordMatch = (k <> "" And InStr(1, f, k, vbTextCompare) > 0)
End Function

'------------------------------------------------------------------------------
' Function: StripPlaceholders
' Purpose: Remove dynamic placeholder tokens like XXX or YYYYMM.
'------------------------------------------------------------------------------
Private Function StripPlaceholders(keyword As String) As String
    Dim k As String
    k = UCase(keyword)
    k = Replace(k, "_XXX", "")
    k = Replace(k, "XXX", "")
    k = Replace(k, "_YYYYMM", "")
    k = Replace(k, "YYYYMM", "")
    k = Replace(k, "_YYYYQX", "")
    k = Replace(k, "YYYYQX", "")
    StripPlaceholders = k
End Function

'------------------------------------------------------------------------------
' Function: NormalizeText
' Purpose: Keep only A-Z and 0-9 for matching.
'------------------------------------------------------------------------------
Private Function NormalizeText(text As String) As String
    Dim i As Long, ch As String, out As String
    text = UCase(text)
    out = ""
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            out = out & ch
        End If
    Next i
    NormalizeText = out
End Function

'------------------------------------------------------------------------------
' Function: RemoveExtension
'------------------------------------------------------------------------------
Private Function RemoveExtension(fileName As String) As String
    Dim pos As Long
    pos = InStrRev(fileName, ".")
    If pos > 0 Then
        RemoveExtension = Left$(fileName, pos - 1)
    Else
        RemoveExtension = fileName
    End If
End Function

'------------------------------------------------------------------------------
' Function: JoinMatches
'------------------------------------------------------------------------------
Private Function JoinMatches(matches As Collection) As String
    Dim i As Long, s As String
    s = ""
    For i = 1 To matches.count
        If s <> "" Then s = s & "; "
        s = s & CStr(matches(i))
    Next i
    JoinMatches = s
End Function

'------------------------------------------------------------------------------
' Function: IsMultiAllowed
' Purpose: Only Merck Payroll Summary Report allows multiple files.
'------------------------------------------------------------------------------
Private Function IsMultiAllowed(item As Object) As Boolean
    Dim nameVal As String, keyVal As String
    nameVal = UCase(CStr(item("Name")))
    keyVal = UCase(CStr(item("Keyword")))
    IsMultiAllowed = (InStr(1, nameVal, "MERCK PAYROLL SUMMARY", vbTextCompare) > 0) _
        Or (InStr(1, keyVal, "MERCK_PAYROLL_SUMMARY", vbTextCompare) > 0)
End Function
