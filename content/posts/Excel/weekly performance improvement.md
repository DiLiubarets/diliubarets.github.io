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
#### **CreatePivot_Week1_2**
```vb
Sub CreatePivot_Week1_2()

    Dim wsData As Worksheet, wsPivot As Worksheet
    Dim pivotCache As PivotCache, pivotTable As PivotTable
    Dim pivotRange As Range, pivotDestination As Range
    Dim lastColumn As Long, lastRow As Long
    Dim cell As Range, field As PivotField, pf As PivotField
    Dim totalValue As Double

    ' Set source data
    Set wsData = ThisWorkbook.Worksheets("general_report")
    Set pivotRange = wsData.Range("A1").CurrentRegion

    ' Delete existing "SP#" sheet if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("SP#").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Create new Pivot Table sheet
    Set wsPivot = ThisWorkbook.Worksheets.Add
    wsPivot.Name = "SP#"

    ' Set pivot table destination
    Set pivotDestination = wsPivot.Range("B10")

    ' Create Pivot Cache & Pivot Table
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotDestination, TableName:="MyPivotTable")

    ' Add fields to Pivot Table
    With pivotTable
        .PivotFields("WP").Orientation = xlRowField
        .PivotFields("Epic Link").Orientation = xlRowField
        .PivotFields("ETC").Orientation = xlDataField
        .PivotFields("EV").Orientation = xlDataField
        .PivotFields("Type of Work").Orientation = xlPageField
    End With

    ' Apply filter to "Type of Work"
    pivotTable.PivotFields("Type of Work").CurrentPage = "Discrete"

    ' Set "ACWP1/2" to show the maximum value
    With pivotTable.PivotFields("ACWP1/2")
        .Orientation = xlDataField
        .Function = xlMax
    End With

    ' Add calculated fields
    With pivotTable.CalculatedFields
        .Add "EV,%", "= 'EV' / 'ETC'"
        .Add "AC/ETC", "= MAX('ACWP1/2') / 'ETC'"
    End With

    ' Add calculated fields to Pivot Table
    With pivotTable
        .PivotFields("EV,%").Orientation = xlDataField
        .PivotFields("AC/ETC").Orientation = xlDataField
    End With

    ' Rename field headers
    For Each field In pivotTable.DataFields
        field.Caption = Replace(field.Caption, "Sum of ", "")
        field.Caption = Replace(field.Caption, "Max of ", "")
    Next field

    ' Set number format
    'pivotTable.DataFields("EV,%").NumberFormat = "0.00%"
    'pivotTable.DataFields("AC/ETC").NumberFormat = "0.00%"
    Dim df As PivotField
	
	' Loop through Pivot Table Data Fields to find and format "EV,%" and "AC/ETC"
	For Each df In pivotTable.DataFields
	    Select Case df.Name
	        Case "EV,%"
	            df.NumberFormat = "0.00%"
	        Case "AC/ETC"
	            df.NumberFormat = "0.00%"
	    End Select
	Next df

    ' Pivot Table formatting
    With pivotTable
        .RowAxisLayout xlTabularRow
        .TableStyle2 = "PivotStyleLight15"
        .DisplayFieldCaptions = False
        .ColumnGrand = False
        .RowGrand = False
    End With

    ' Disable subtotals
    For Each pf In pivotTable.RowFields
        pf.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    Next pf

    ' Find last column
    lastColumn = FindLastColumn(wsPivot, 13)

    ' Insert "Total" row
    With wsPivot.Cells(9, 2)
        .Value = "Total"
        .Font.Bold = True
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Add total formulas
    For Each cell In wsPivot.Range(wsPivot.Cells(9, 3), wsPivot.Cells(9, lastColumn))
        lastRow = FindLastRow(wsPivot, cell.Column)
        Dim columnHeader As String
        columnHeader = wsPivot.Cells(10, cell.Column).Value

        Select Case columnHeader
            Case "EV,%", "AC/ETC"
                cell.Value = ""
            Case Else
                cell.Formula = "=SUM(" & wsPivot.Cells(11, cell.Column).Address & ":" & wsPivot.Cells(lastRow, cell.Column).Address & ")"
                If IsNumeric(cell.Value) Then
                    totalValue = cell.Value
                    If totalValue = 0 Then cell.Value = ""
                End If
        End Select

        ' Format total row
        cell.Font.Bold = True
        cell.Interior.Color = RGB(0, 0, 0)
        cell.Font.Color = RGB(255, 255, 255)
    Next cell

    ' AutoFit columns
    wsPivot.Cells.EntireColumn.AutoFit

    ' Run Macro3
    Macro3

    ' Run Macro6
    Macro6

    ' Success message
    MsgBox "Pivot Table for Week 1-2 created successfully!", vbInformation

End Sub

' Function to find the last row in a column
Function FindLastRow(ws As Worksheet, col As Long) As Long
    FindLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' Function to find the last column in a row
Function FindLastColumn(ws As Worksheet, row As Long) As Long
    FindLastColumn = ws.Cells(row, ws.Columns.Count).End(xlToLeft).Column
End Function

Sub Macro3()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("SP#")

    ' Copy and paste formats
    ws.Range("K13").Copy
    ws.Range("K13").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ws.Range("H13:I13").Copy
    ws.Range("L13:M13").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Remove background color from N13
    ws.Range("N13").Interior.Pattern = xlNone

    ' Hide rows 10-11
    ws.Rows("10:11").Hidden = True

    ' Insert blank rows above row 8
    ws.Rows("8:8").Insert Shift:=xlDown
    ws.Rows("8:8").Insert Shift:=xlDown

    ' Insert new column A
    ws.Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    ' Insert formula in L16
    ws.Range("L16").FormulaR1C1 = "=IFERROR(RC[-2]/RC[-4],""NO ETC"")"

    ' Autofill formula down
    Dim lastRow As Long
    lastRow = FindLastRow(ws, 12) ' Column L
    ws.Range("L16").AutoFill Destination:=ws.Range("L16:L" & lastRow)

    Application.CutCopyMode = False
End Sub

Sub Macro6()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("SP#")

    ' Insert three columns at the beginning
    ws.Columns("A:A").Insert Shift:=xlToRight
    ws.Columns("A:A").Insert Shift:=xlToRight
    ws.Columns("A:A").Insert Shift:=xlToRight

    ' Insert three rows above row 8
    ws.Rows("8:8").Insert Shift:=xlDown
    ws.Rows("8:8").Insert Shift:=xlDown
    ws.Rows("8:8").Insert Shift:=xlDown

End Sub
```


```vb
Sub CreatePivotWAR_report()
    Dim wsData As Worksheet, wsPivot As Worksheet
    Dim pivotCache As PivotCache, pivotTable As PivotTable
    Dim pivotRange As Range, pivotDestination As Range
    Dim monthName As String
    Dim lastRow As Long, lastCol As Long
    Dim cell As Range, field As PivotField, pf As PivotField
    Dim totalValue As Double
    Dim dataField As PivotField

    ' Get current month name
    monthName = Format(Date, "mmm")

    ' Set wsData to the sheet named "WAR <monthName>"
    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets("WAR " & monthName)
    On Error GoTo 0

    If wsData Is Nothing Then
        MsgBox "Sheet 'WAR " & monthName & "' does not exist.", vbExclamation
        Exit Sub
    End If

    ' Find last row and last column dynamically
    lastRow = wsData.Cells(Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, Columns.Count).End(xlToLeft).Column

    ' Define pivot range dynamically
    Set pivotRange = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))

    ' Ensure the "SP#" sheet exists
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("SP#")
    On Error GoTo 0

    If wsPivot Is Nothing Then
        MsgBox "Sheet 'SP#' does not exist.", vbExclamation
        Exit Sub
    End If

    ' Set pivot table destination
    Set pivotDestination = wsPivot.Range("A12")

    ' Delete existing PivotTable if it exists
    On Error Resume Next
    wsPivot.PivotTables("WAR_Pivot").TableRange2.Clear
    On Error GoTo 0

    ' Create Pivot Cache & PivotTable
    On Error Resume Next
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)
    If Err.Number <> 0 Then
        MsgBox "Error creating PivotCache: " & Err.Description, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    On Error Resume Next
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotDestination, TableName:="WAR_Pivot")
    If Err.Number <> 0 Then
        MsgBox "Error creating PivotTable: " & Err.Description, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    ' Add fields to PivotTable dynamically
    With pivotTable
        ' Add "WP" as Row Field if it exists
        If FieldExists(wsData, "WP") Then .PivotFields("WP").Orientation = xlRowField

        ' Add "Type of Work" as Page Field (Filter) if it exists
        If FieldExists(wsData, "Type of Work") Then .PivotFields("Type of Work").Orientation = xlPageField

        ' Loop through columns to add numeric fields as Data Fields
        For Each cell In wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, lastCol))
            If IsNumeric(wsData.Cells(2, cell.Column).Value) Then
                .PivotFields(cell.Value).Orientation = xlDataField
            End If
        Next cell
    End With

    ' Apply filter to "Type of Work" field if it exists
    On Error Resume Next
    pivotTable.PivotFields("Type of Work").CurrentPage = "Discrete"
    On Error GoTo 0

    ' Format PivotTable
    With pivotTable
        .RowAxisLayout xlTabularRow
        .TableStyle2 = "PivotStyleMedium15"
        .DisplayFieldCaptions = False
        .ColumnGrand = True
        .RowGrand = False
    End With

    ' Update field captions
    For Each field In pivotTable.DataFields
        field.Caption = Replace(field.Caption, "Sum of ", "")
        field.Caption = Replace(field.Caption, "Max of ", "")
    Next field

    ' Disable subtotals for Row Fields
    For Each pf In pivotTable.RowFields
        pf.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    Next pf

    ' AutoFit all columns
    wsPivot.Cells.EntireColumn.AutoFit

    ' Success message
    MsgBox "Pivot Table created successfully in 'SP#'!", vbInformation

End Sub

' Function to check if a field exists in the dataset
Function FieldExists(ws As Worksheet, fieldName As String) As Boolean
    Dim cell As Range
    FieldExists = False
    For Each cell In ws.Range("A1:Z1") ' Adjust range if needed
        If cell.Value = fieldName Then
            FieldExists = True
            Exit Function
        End If
    Next cell
End Function
```
```vb
Sub CreatePivotWAR_report()
    Dim wsData As Worksheet, wsPivot As Worksheet
    Dim pivotCache As PivotCache, pivotTable As PivotTable
    Dim pivotRange As Range, pivotDestination As Range
    Dim monthName As String
    Dim lastRow As Long, lastCol As Long
    Dim field As PivotField, pf As PivotField

    ' Get current month name
    monthName = Format(Date, "mmm")

    ' Set wsData to the sheet named "WAR <monthName>"
    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets("WAR " & monthName)
    On Error GoTo 0

    If wsData Is Nothing Then
        MsgBox "Sheet 'WAR " & monthName & "' does not exist.", vbExclamation
        Exit Sub
    End If

    ' Find last row and last column dynamically
    lastRow = wsData.Cells(Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, Columns.Count).End(xlToLeft).Column

    ' Define pivot range dynamically
    Set pivotRange = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))

    ' Ensure the "SP#" sheet exists
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("SP#")
    On Error GoTo 0

    If wsPivot Is Nothing Then
        MsgBox "Sheet 'SP#' does not exist.", vbExclamation
        Exit Sub
    End If

    ' Set pivot table destination
    Set pivotDestination = wsPivot.Range("A12")

    ' Delete existing PivotTable if it exists
    On Error Resume Next
    wsPivot.PivotTables("WAR_Pivot").TableRange2.Clear
    On Error GoTo 0

    ' Create Pivot Cache & PivotTable
    On Error Resume Next
    Set pivotCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=pivotRange)
    If Err.Number <> 0 Then
        MsgBox "Error creating PivotCache: " & Err.Description, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    On Error Resume Next
    Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotDestination, TableName:="WAR_Pivot")
    If Err.Number <> 0 Then
        MsgBox "Error creating PivotTable: " & Err.Description, vbExclamation
        Exit Sub
    End If
    On Error GoTo 0

    ' Add only the required fields if they exist
    With pivotTable
        If FieldExists(wsData, "WP") Then .PivotFields("WP").Orientation = xlRowField
        If FieldExists(wsData, "AC week1-2") Then .PivotFields("AC week1-2").Orientation = xlDataField
        If FieldExists(wsData, "AC week3-4") Then .PivotFields("AC week3-4").Orientation = xlDataField
        If FieldExists(wsData, "Type of Work") Then .PivotFields("Type of Work").Orientation = xlPageField
    End With

    ' Apply filter to "Type of Work" field if it exists
    On Error Resume Next
    pivotTable.PivotFields("Type of Work").CurrentPage = "Discrete"
    On Error GoTo 0

    ' Format PivotTable
    With pivotTable
        .RowAxisLayout xlTabularRow
        .TableStyle2 = "PivotStyleMedium15"
        .DisplayFieldCaptions = False
        .ColumnGrand = True
        .RowGrand = False
    End With

    ' Update field captions
    For Each field In pivotTable.DataFields
        field.Caption = Replace(field.Caption, "Sum of ", "")
        field.Caption = Replace(field.Caption, "Max of ", "")
    Next field

    ' Disable subtotals for Row Fields
    For Each pf In pivotTable.RowFields
        pf.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    Next pf

    ' AutoFit all columns
    wsPivot.Cells.EntireColumn.AutoFit

    ' Success message
    MsgBox "Pivot Table created successfully in 'SP#'!", vbInformation

End Sub

' Function to check if a field exists in the dataset
Function FieldExists(ws As Worksheet, fieldName As String) As Boolean
    Dim cell As Range
    FieldExists = False
    For Each cell In ws.Range("A1:Z1") ' Adjust range if needed
        If cell.Value = fieldName Then
            FieldExists = True
            Exit Function
        End If
    Next cell
End Function
```
#### **copy Pivot**
```vb
Sub AddFormulaAndConvertToTable()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim tableRange As Range
    Dim tbl As ListObject
    Dim tblName As String

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("SP#")

    ' Find the last used row in column E dynamically
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    ' Find the last used column dynamically (starting from column T)
    lastCol = ws.Cells(11, ws.Columns.Count).End(xlToLeft).Column

    ' Ensure lastCol is at least column T (20th column)
    If lastCol < 20 Then lastCol = 29 ' Column AC (29th column)

    ' Apply formula dynamically in column T
    Dim i As Long
    For i = 11 To lastRow
        ws.Cells(i, "T").Formula = "=IF(ISBLANK(E" & i & "),"""",E" & i & ")"
    Next i

    ' Define the range to autofill dynamically
    Dim formulaRange As Range, fillRange As Range
    Set formulaRange = ws.Range("T11:T" & lastRow)
    Set fillRange = ws.Range("T11", ws.Cells(11, lastCol)) ' Expands to the last column

    ' Fill across dynamically
    formulaRange.Copy
    fillRange.PasteSpecial Paste:=xlPasteFormulas

    ' Fill down dynamically
    fillRange.AutoFill Destination:=ws.Range("T11", ws.Cells(lastRow, lastCol)), Type:=xlFillDefault

    ' Clean up clipboard
    Application.CutCopyMode = False

    ' Define the table range
    Set tableRange = ws.Range("T11", ws.Cells(lastRow, lastCol))

    ' Delete existing table if it exists
    On Error Resume Next
    Set tbl = ws.ListObjects("SP_Table")
    If Not tbl Is Nothing Then tbl.Delete
    On Error GoTo 0

    ' Create a new table
    Set tbl = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    tbl.Name = "SP_Table"
    tbl.TableStyle = "TableStyleMedium9" ' Change style if needed

    ' Notify user
    MsgBox "Formula added dynamically and converted to a table (SP_Table)!", vbInformation

End Sub
```
```vb
Sub AddFormulaAndConvertToTable_NoHeader()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim tableRange As Range
    Dim tbl As ListObject
    Dim tblName As String

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("SP#")

    ' Find the last used row in column E dynamically
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    ' Find the last used column dynamically (starting from column T)
    lastCol = ws.Cells(11, ws.Columns.Count).End(xlToLeft).Column

    ' Ensure lastCol is at least column T (20th column)
    If lastCol < 20 Then lastCol = 29 ' Column AC (29th column)

    ' Apply formula dynamically in column T
    Dim i As Long
    For i = 11 To lastRow
        ws.Cells(i, "T").Formula = "=IF(ISBLANK(E" & i & "),"""",E" & i & ")"
    Next i

    ' Define the range to autofill dynamically
    Dim formulaRange As Range, fillRange As Range
    Set formulaRange = ws.Range("T11:T" & lastRow)
    Set fillRange = ws.Range("T11", ws.Cells(11, lastCol)) ' Expands to the last column

    ' Fill across dynamically
    formulaRange.Copy
    fillRange.PasteSpecial Paste:=xlPasteFormulas

    ' Fill down dynamically
    fillRange.AutoFill Destination:=ws.Range("T11", ws.Cells(lastRow, lastCol)), Type:=xlFillDefault

    ' Clean up clipboard
    Application.CutCopyMode = False

    ' Define the table range (without headers)
    Set tableRange = ws.Range("T11", ws.Cells(lastRow, lastCol))

    ' Delete existing table if it exists
    On Error Resume Next
    Set tbl = ws.ListObjects("SP_Table")
    If Not tbl Is Nothing Then tbl.Delete
    On Error GoTo 0

    ' Create a new table WITHOUT HEADERS
    Set tbl = ws.ListObjects.Add(xlSrcRange, tableRange, , xlNo)
    tbl.Name = "SP_Table"
    tbl.TableStyle = "TableStyleMedium9" ' Change style if needed

    ' Notify user
    MsgBox "Formula added dynamically and converted to a table (SP_Table) **without headers**!", vbInformation

End Sub
```