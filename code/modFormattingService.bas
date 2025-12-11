Attribute VB_Name = "modFormattingService"
'==============================================================================
' Module: modFormattingService
' Purpose: Formatting services for output workbooks
' Description: Handles Diff summary formatting and general worksheet formatting
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Sub: SummarizeDiffColumns
' Purpose: Count FALSE values in Diff columns and highlight headers
' Parameters:
'   ws - Worksheet containing Diff columns
'   headerRow - Row number of the header
'   firstDataRow - First row of data
'   lastDataRow - Last row of data
'   firstDiffCol - First Diff column index
'   lastDiffCol - Last Diff column index
' Note: Writes FALSE count to header row and highlights red if count > 0
'------------------------------------------------------------------------------
Public Sub SummarizeDiffColumns( _
    ws As Worksheet, _
    headerRow As Long, _
    firstDataRow As Long, _
    lastDataRow As Long, _
    firstDiffCol As Long, _
    lastDiffCol As Long)
    
    Dim col As Long
    Dim falseCount As Long
    Dim rng As Range
    
    On Error GoTo ErrHandler
    
    For col = firstDiffCol To lastDiffCol
        Set rng = ws.Range(ws.Cells(firstDataRow, col), ws.Cells(lastDataRow, col))
        falseCount = WorksheetFunction.CountIf(rng, "FALSE")
        
        With ws.Cells(headerRow, col)
            .value = falseCount
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
    Next col
    
    Exit Sub
    
ErrHandler:
    LogError "modFormattingService", "SummarizeDiffColumns", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: ApplyStandardFormatting
' Purpose: Apply standard formatting to a worksheet
' Parameters:
'   ws - Worksheet to format
'   headerRow - Row number of the header (default 1)
'------------------------------------------------------------------------------
Public Sub ApplyStandardFormatting(ws As Worksheet, Optional headerRow As Long = 1)
    Dim lastRow As Long, lastCol As Long
    
    On Error GoTo ErrHandler
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    lastCol = ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column
    
    If lastRow < headerRow Or lastCol < 1 Then Exit Sub
    
    With ws
        ' Format header row
        With .Range(.Cells(headerRow, 1), .Cells(headerRow, lastCol))
            .Font.Bold = True
            .Interior.Color = RGB(200, 200, 200)
        End With
        
        ' Freeze panes
        .Activate
        .Cells(headerRow + 1, 1).Select
        ActiveWindow.FreezePanes = True
        
        ' Autofit columns (with max width limit)
        .Columns.AutoFit
        Dim col As Long
        For col = 1 To lastCol
            If .Columns(col).ColumnWidth > 50 Then
                .Columns(col).ColumnWidth = 50
            End If
        Next col
    End With
    
    Exit Sub
    
ErrHandler:
    LogError "modFormattingService", "ApplyStandardFormatting", Err.Number, Err.Description
End Sub

'------------------------------------------------------------------------------
' Sub: FormatAsNumber
' Purpose: Format a range as number with specified decimal places
' Parameters:
'   rng - Range to format
'   decimals - Number of decimal places (default 2)
'------------------------------------------------------------------------------
Public Sub FormatAsNumber(rng As Range, Optional decimals As Long = 2)
    On Error Resume Next
    
    Select Case decimals
        Case 0
            rng.NumberFormat = "#,##0"
        Case 1
            rng.NumberFormat = "#,##0.0"
        Case 2
            rng.NumberFormat = "#,##0.00"
        Case Else
            rng.NumberFormat = "#,##0." & String(decimals, "0")
    End Select
    
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Sub: FormatAsDate
' Purpose: Format a range as date
' Parameters:
'   rng - Range to format
'   dateFormat - Date format string (default "yyyy/mm/dd")
'------------------------------------------------------------------------------
Public Sub FormatAsDate(rng As Range, Optional dateFormat As String = "yyyy/mm/dd")
    On Error Resume Next
    rng.NumberFormat = dateFormat
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Sub: FormatAsCurrency
' Purpose: Format a range as currency (HKD)
' Parameters:
'   rng - Range to format
'------------------------------------------------------------------------------
Public Sub FormatAsCurrency(rng As Range)
    On Error Resume Next
    rng.NumberFormat = "#,##0.00"
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Sub: HighlightCell
' Purpose: Highlight a cell with specified color
' Parameters:
'   cell - Cell to highlight
'   colorType - Color type ("red", "yellow", "green", "none")
'------------------------------------------------------------------------------
Public Sub HighlightCell(cell As Range, colorType As String)
    On Error Resume Next
    
    Select Case LCase(colorType)
        Case "red"
            cell.Interior.Color = RGB(255, 200, 200)
        Case "yellow"
            cell.Interior.Color = RGB(255, 255, 200)
        Case "green"
            cell.Interior.Color = RGB(200, 255, 200)
        Case "none"
            cell.Interior.ColorIndex = xlColorIndexNone
        Case Else
            cell.Interior.ColorIndex = xlColorIndexNone
    End Select
    
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Sub: AddBorders
' Purpose: Add borders to a range
' Parameters:
'   rng - Range to add borders to
'------------------------------------------------------------------------------
Public Sub AddBorders(rng As Range)
    On Error Resume Next
    
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Sub: CreateNamedRange
' Purpose: Create or update a named range
' Parameters:
'   wb - Workbook
'   rangeName - Name for the range
'   rng - Range to name
'------------------------------------------------------------------------------
Public Sub CreateNamedRange(wb As Workbook, rangeName As String, rng As Range)
    On Error Resume Next
    
    ' Delete existing name if exists
    wb.Names(rangeName).Delete
    
    ' Create new named range
    wb.Names.Add Name:=rangeName, RefersTo:=rng
    
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Sub: SetColumnWidths
' Purpose: Set specific column widths
' Parameters:
'   ws - Worksheet
'   colWidths - Array of column widths (index 0 = column A)
'------------------------------------------------------------------------------
Public Sub SetColumnWidths(ws As Worksheet, colWidths As Variant)
    Dim i As Long
    
    On Error Resume Next
    
    If IsArray(colWidths) Then
        For i = LBound(colWidths) To UBound(colWidths)
            ws.Columns(i + 1).ColumnWidth = colWidths(i)
        Next i
    End If
    
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Sub: ApplyConditionalFormatting
' Purpose: Apply conditional formatting for TRUE/FALSE values
' Parameters:
'   rng - Range to format
'------------------------------------------------------------------------------
Public Sub ApplyConditionalFormatting(rng As Range)
    On Error Resume Next
    
    ' Clear existing conditional formatting
    rng.FormatConditions.Delete
    
    ' Add formatting for FALSE
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="FALSE")
        .Interior.Color = RGB(255, 200, 200)
        .Font.Color = RGB(156, 0, 6)
    End With
    
    ' Add formatting for TRUE
    With rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="TRUE")
        .Interior.Color = RGB(200, 255, 200)
        .Font.Color = RGB(0, 97, 0)
    End With
    
    On Error GoTo 0
End Sub
