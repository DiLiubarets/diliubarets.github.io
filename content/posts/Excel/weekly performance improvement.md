```vb
Sub AdjustGeneralReport()
    Dim ws As Worksheet
    Dim rng As Range
    Dim deleteRange As Range
    Dim ROICol As Range
    Dim epicLinkCol As Range
    Dim lastRow As Long
    Dim wpColLetter As String
    Dim tbl As ListObject
    
    Set ws = ThisWorkbook.Sheets("general_report")
    
    ' AutoFit columns
    ws.Columns.AutoFit
    
    ' Remove Picture if exists
    On Error Resume Next
    ws.Shapes("Picture 1").Delete
    On Error GoTo 0
    
    With ws
        .Cells.UnMerge
        .Cells.Borders.LineStyle = xlNone ' Clear all borders
        .Rows("1:3").Delete ' Delete first three rows
        
        ' Format the first row
        With .Rows(1)
            .Interior.Color = RGB(64, 64, 64)
            .Font.Color = RGB(255, 255, 255)
        End With
        
        ' Delete rows containing "Generated at" using AutoFilter (more efficient)
        DeleteRowsContaining ws, "Generated at"
        
        ' Convert range to a table
        Set rng = .UsedRange
        If Not rng Is Nothing Then
            ' Check if a table already exists
            If .ListObjects.Count = 0 Then
                Set tbl = .ListObjects.Add(xlSrcRange, rng, , xlYes)
                tbl.Name = "JiraData_WeeklyPerformance_Table"
                tbl.TableStyle = "TableStyleLight8"
            Else
                Set tbl = .ListObjects(1) ' Use existing table
            End If
        End If
    End With
    
    ' Delete specific columns using a separate module function
    Module1.DeleteColumnsByNames
    
    ' Highlight non-numeric values in "ROI (Hours)" column
    HighlightNonNumericValues ws, "ROI (Hours)"
    
    ' Insert WP column and apply XLOOKUP formula
    InsertWPColumn ws
    
    ' Call additional functions from Module1
    Module1.Insert_TypeOFWork
    Module1.InsertEVColumn
    Module1.InsertETCColumn
    Module1.InsertACWP1_2Column
    Module1.InsertColumnACWP3_4Formula
    
    ' Call external function for finding differences
    FindDifferences.FindUniqueDifferencesAndAddDetails
    
    MsgBox "Jira general_report adjusted"
End Sub
```
####  **Delete Rows Containing a Specific Text (Using AutoFilter)**
```vb
Sub DeleteRowsContaining(ws As Worksheet, searchText As String)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    With ws
        .UsedRange.AutoFilter Field:=1, Criteria1:="*" & searchText & "*"
        On Error Resume Next
        .Rows("2:" & lastRow).SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        .AutoFilterMode = False
    End With
End Sub
```
#### **Highlight Non-Numeric Values in a Column**
```vb
Sub HighlightNonNumericValues(ws As Worksheet, colName As String)
    Dim col As Range, cell As Range
    Dim lastRow As Long
    
    Set col = ws.Rows(1).Find(What:=colName, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not col Is Nothing Then
        lastRow = ws.Cells(ws.Rows.Count, col.Column).End(xlUp).Row
        For Each cell In ws.Range(col.Offset(1, 0), ws.Cells(lastRow, col.Column))
            If Not IsNumeric(cell.Value) Then
                cell.Interior.Color = RGB(139, 0, 0) ' Dark red
            End If
        Next cell
    Else
        MsgBox "Column '" & colName & "' not found!", vbExclamation
    End If
End Sub
```
####  **Insert WP Column and Apply XLOOKUP Formula**
```vb
Sub InsertWPColumn(ws As Worksheet)
    Dim epicLinkCol As Range
    Dim lastRow As Long
    Dim wpColLetter As String
    
    Set epicLinkCol = ws.Rows(1).Find(What:="Epic Link", LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not epicLinkCol Is Nothing Then
        ' Insert new column
        epicLinkCol.Offset(0, 1).EntireColumn.Insert Shift:=xlToRight
        ws.Cells(1, epicLinkCol.Column + 1).Value = "WP"
        
        ' Get last row
        lastRow = ws.Cells(ws.Rows.Count, epicLinkCol.Column).End(xlUp).Row
        wpColLetter = Split(epicLinkCol.Offset(0, 1).Address, "$")(1)
        
        ' Apply XLOOKUP formula
        ws.Range(wpColLetter & "2:" & wpColLetter & lastRow).Formula = _
            "=XLOOKUP([@[Epic Link]],working!D:D,working!B:B,""NOT FOUND"")"
        
        ' Apply conditional formatting to highlight "NOT FOUND"
        With ws.Range(wpColLetter & "2:" & wpColLetter & lastRow)
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""NOT FOUND"""
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
            With .FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = RGB(255, 0, 0) ' Red
            End With
        End With
    Else
        MsgBox "Column 'Epic Link' not found!", vbExclamation
    End If
End Sub
```