Attribute VB_Name = "modSP2_CheckResult_Diff"
'==============================================================================
' Module: modSP2_CheckResult_Diff
' Purpose: Subprocess 2 - Diff column computation
' Description: Computes TRUE/FALSE comparison between Benchmark and Check
'              for all 63 fields according to HK payroll validation output template
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Sub: SP2_ComputeDiff
' Purpose: Compute all Diff columns for the 63 template fields
'------------------------------------------------------------------------------
Public Sub SP2_ComputeDiff(valWb As Workbook, weinIndex As Object)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fieldCount As Long
    Dim field As tCheckResultColumn
    Dim benchmarkCol As Long, checkCol As Long, diffCol As Long
    Dim processedCount As Long
    
    On Error GoTo ErrHandler
    
    Set ws = valWb.Worksheets("Check Result")
    
    ' Get data range
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow < 5 Then lastRow = 5 ' At least one data row
    
    ' Update template column indices
    UpdateTemplateColumnIndices ws, 4
    
    ' Process each template field
    fieldCount = GetTemplateFieldCount()
    processedCount = 0
    
    For i = 1 To fieldCount
        field = GetTemplateField(i)
        
        If field.HasDiff Then
            benchmarkCol = GetBenchmarkColIndex(field.BenchmarkName)
            checkCol = GetCheckColIndex(field.BenchmarkName)
            diffCol = GetDiffColIndex(field.BenchmarkName)
            
            If benchmarkCol > 0 And checkCol > 0 And diffCol > 0 Then
                ComputeDiffColumn ws, benchmarkCol, checkCol, diffCol, 5, lastRow, field.BenchmarkName
                processedCount = processedCount + 1
            Else
                LogInfo "modSP2_CheckResult_Diff", "SP2_ComputeDiff", _
                    "Skipping field (columns not found): " & field.BenchmarkName & _
                    " (Bench=" & benchmarkCol & ", Check=" & checkCol & ", Diff=" & diffCol & ")"
            End If
        End If
    Next i
    
    ' Compute FALSE counts and highlight in row 3
    ComputeFalseCountsAndHighlight ws, 4, 3, 5, lastRow
    
    LogInfo "modSP2_CheckResult_Diff", "SP2_ComputeDiff", _
        "Diff computation completed for " & processedCount & " fields"
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Diff", "SP2_ComputeDiff", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: ComputeFalseCountsAndHighlight
' Purpose: Count FALSE values in each Diff column and highlight if > 0
' Parameters:
'   ws - Check Result worksheet
'   headerRow - Row number of the header
'   countRow - Row number to write FALSE counts
'   firstDataRow - First row of data
'   lastDataRow - Last row of data
'------------------------------------------------------------------------------
Private Sub ComputeFalseCountsAndHighlight(ws As Worksheet, headerRow As Long, _
    countRow As Long, firstDataRow As Long, lastDataRow As Long)
    
    Dim lastCol As Long
    Dim col As Long
    Dim headerValue As String
    Dim falseCount As Long
    Dim rng As Range
    
    On Error GoTo ErrHandler
    
    lastCol = ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        headerValue = Trim(CStr(Nz(ws.Cells(headerRow, col).Value, "")))
        
        ' Check if this is a Diff column
        If Right(UCase(headerValue), 4) = "DIFF" Then
            ' Count FALSE values
            Set rng = ws.Range(ws.Cells(firstDataRow, col), ws.Cells(lastDataRow, col))
            falseCount = WorksheetFunction.CountIf(rng, "FALSE")
            
            ' Write count to count row
            With ws.Cells(countRow, col)
                .Value = falseCount
                
                ' Highlight red if FALSE count > 0
                If falseCount > 0 Then
                    .Interior.Color = vbRed
                    .Font.Color = vbWhite
                    .Font.Bold = True
                Else
                    .Interior.ColorIndex = xlColorIndexNone
                    .Font.Color = vbBlack
                    .Font.Bold = False
                End If
            End With
        End If
    Next col
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Diff", "ComputeFalseCountsAndHighlight", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: ComputeDiffColumn
' Purpose: Compute a single Diff column
'------------------------------------------------------------------------------
Private Sub ComputeDiffColumn(ws As Worksheet, benchmarkCol As Long, checkCol As Long, _
    diffCol As Long, firstRow As Long, lastRow As Long, fieldName As String)
    
    Dim row As Long
    Dim benchVal As Variant, checkVal As Variant
    Dim diffResult As Boolean
    
    On Error GoTo ErrHandler
    
    For row = firstRow To lastRow
        benchVal = ws.Cells(row, benchmarkCol).Value
        checkVal = ws.Cells(row, checkCol).Value
        
        ' Compute diff
        diffResult = CompareCellValues(benchVal, checkVal, fieldName)
        
        ' Write result
        If diffResult Then
            ws.Cells(row, diffCol).Value = "TRUE"
        Else
            ws.Cells(row, diffCol).Value = "FALSE"
        End If
    Next row
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Diff", "ComputeDiffColumn", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Function: CompareCellValues
' Purpose: Compare two cell values with appropriate logic
' Returns: TRUE if values match, FALSE if different
'------------------------------------------------------------------------------
Private Function CompareCellValues(benchVal As Variant, checkVal As Variant, fieldName As String) As Boolean
    Dim result As Boolean
    
    On Error Resume Next
    
    ' Both blank = TRUE
    If IsEmpty(benchVal) And IsEmpty(checkVal) Then
        CompareCellValues = True
        Exit Function
    End If
    
    If (IsEmpty(benchVal) Or benchVal = "") And (IsEmpty(checkVal) Or checkVal = "") Then
        CompareCellValues = True
        Exit Function
    End If
    
    ' Special rule for Last Hired Date: if either is before 2025-01-01, return TRUE
    If InStr(UCase(fieldName), "LAST HIRE") > 0 Or InStr(UCase(fieldName), "LAST HIRED") > 0 Then
        If IsDate(benchVal) Then
            If CDate(benchVal) < DateSerial(2025, 1, 1) Then
                CompareCellValues = True
                Exit Function
            End If
        End If
        If IsDate(checkVal) Then
            If CDate(checkVal) < DateSerial(2025, 1, 1) Then
                CompareCellValues = True
                Exit Function
            End If
        End If
    End If
    
    ' Date comparison
    If IsDate(benchVal) And IsDate(checkVal) Then
        result = (CLng(CDate(benchVal)) = CLng(CDate(checkVal)))
        CompareCellValues = result
        Exit Function
    End If
    
    ' Numeric comparison (with small tolerance for floating point)
    If IsNumeric(benchVal) And IsNumeric(checkVal) Then
        Dim diff As Double
        diff = Abs(CDbl(benchVal) - CDbl(checkVal))
        result = (diff < 0.01) ' Allow small tolerance
        CompareCellValues = result
        Exit Function
    End If
    
    ' Text comparison (case-insensitive, trimmed)
    Dim benchStr As String, checkStr As String
    benchStr = UCase(Trim(CStr(Nz(benchVal, ""))))
    checkStr = UCase(Trim(CStr(Nz(checkVal, ""))))
    
    result = (benchStr = checkStr)
    CompareCellValues = result
    
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' Sub: AddDiffColumns
' Purpose: Add Diff columns next to Check columns (if not already present)
'------------------------------------------------------------------------------
Public Sub AddDiffColumns(ws As Worksheet)
    Dim lastCol As Long
    Dim col As Long
    Dim headerValue As String
    Dim baseName As String
    Dim diffColName As String
    Dim diffCol As Long
    
    On Error GoTo ErrHandler
    
    lastCol = ws.Cells(4, ws.Columns.count).End(xlToLeft).Column
    
    ' Process from right to left to avoid column shift issues
    For col = lastCol To 1 Step -1
        headerValue = Trim(CStr(Nz(ws.Cells(4, col).Value, "")))
        
        ' Check if this is a Check column
        If Right(UCase(headerValue), 5) = "CHECK" Then
            baseName = Left(headerValue, Len(headerValue) - 6) ' Remove " Check"
            diffColName = baseName & " Diff"
            
            ' Check if Diff column already exists
            diffCol = FindColumnByHeader(ws.Rows(4), diffColName)
            
            If diffCol = 0 Then
                ' Insert Diff column after Check column
                ws.Columns(col + 1).Insert Shift:=xlToRight
                ws.Cells(4, col + 1).Value = diffColName
            End If
        End If
    Next col
    
    Exit Sub
    
ErrHandler:
    LogError "modSP2_CheckResult_Diff", "AddDiffColumns", Err.Number, Err.Description
End Sub
