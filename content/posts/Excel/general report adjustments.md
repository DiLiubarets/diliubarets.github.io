Combine sheet with col 
```vb
Sub CombineSheets_withCol()

    Dim ws As Worksheet

    Dim combinedWs As Worksheet

    Dim rng As Range

    Dim deleteRange As Range

    Dim sprintCol As Range

    Dim statusCol As Range

    Dim lastRow As Long

    Dim cell As Range

    Dim sprintColLetter As String

    Dim newColLetter As String

    Dim columnsToDelete As Variant

    Dim colName As Variant

    Dim found As Range

    Dim nextRow As Long

    Dim spName As String

    Dim tbl As ListObject

    Dim firstSheet As Boolean

    Dim statusColNum As Integer

    Dim originalName As String

    Dim counter As Integer

    ' Define the columns to delete

    columnsToDelete = Array("Issue Links", "Fix Version/s", "ROI($)", "Updated", "Sprint History", "Sprint commitment", _

                            "Project Status (Date / Comments)", "Last Issue Comment", "Description", "Solution")

        ' Define new columns to add in CombinedData

    newHeaders = Array("Original Estima hrs", "Remaining hrs", "Time Spent hrs", "Original Estima SP", _

                       "Remaining SP", "Time Spent SP", "EV, SP", "WP")

    ' Create a new sheet for the combined data

    On Error Resume Next

    Set combinedWs = ThisWorkbook.Sheets("CombinedData")

    On Error GoTo 0

    If combinedWs Is Nothing Then

        Set combinedWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))

        combinedWs.Name = "CombinedData"

    Else

        combinedWs.Cells.Clear

    End If

    nextRow = 1

    firstSheet = True ' Flag to track the first sheet

    For Each ws In ThisWorkbook.Worksheets

        If ws.Name <> "Instructions" And ws.Name <> "CombinedData" Then

            With ws

                On Error Resume Next

                .Shapes.Range(Array("Picture 1")).Delete

                On Error GoTo 0

                .Cells.UnMerge

                .Cells.Borders.LineStyle = xlNone ' Clear all borders

                .Rows("1:3").Delete ' Delete the first three rows

                ' Format the first row

                .Rows(1).Interior.Color = RGB(64, 64, 64)

                .Rows(1).Font.Color = RGB(255, 255, 255)

                ' Find and delete rows containing "Generated at"

                Set rng = ws.UsedRange.Find(What:="Generated at", LookIn:=xlValues, LookAt:=xlPart, _

                                            SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)

                If Not rng Is Nothing Then

                    Set deleteRange = rng

                    Do

                        If deleteRange Is Nothing Then

                            Set deleteRange = rng

                        Else

                            Set deleteRange = Union(deleteRange, rng)

                        End If

                        Set rng = ws.UsedRange.FindNext(rng)

                    Loop While Not rng Is Nothing And rng.Address <> deleteRange.Cells(1, 1).Address

                    deleteRange.EntireRow.Delete

                End If

            End With

            ' Find the column with the header "Sprint"

            Set sprintCol = ws.Rows(1).Find(What:="Sprint", LookIn:=xlValues, LookAt:=xlWhole)

            If Not sprintCol Is Nothing Then

                sprintColLetter = Split(sprintCol.Address, "$")(1)

                ' Insert a new column after the Sprint column

                sprintCol.Offset(0, 1).EntireColumn.Insert Shift:=xlToRight

                newColLetter = Split(sprintCol.Offset(0, 1).Address, "$")(1)

                ws.Cells(1, sprintCol.Column + 1).Value = "SP#"

                ' Find the last row in the Sprint column

                lastRow = ws.Cells(ws.Rows.Count, sprintCol.Column).End(xlUp).Row

                ' Apply the formula to extract sprint number

                For Each cell In ws.Range(newColLetter & "2:" & newColLetter & lastRow)

                    cell.Formula = "=MID(" & sprintColLetter & cell.Row & ", FIND(""_"", " & sprintColLetter & cell.Row & ") + 1, " & _

                                   "FIND(""_"", " & sprintColLetter & cell.Row & ", FIND(""_"", " & sprintColLetter & cell.Row & ") + 1) - " & _

                                   "FIND(""_"", " & sprintColLetter & cell.Row & ") - 1)"

                Next cell

                ' Convert formulas to values

                For Each cell In ws.Range(newColLetter & "2:" & newColLetter & lastRow)

                    cell.Value = cell.Value

                Next cell

                ' Rename the sheet based on the value in the "SP#" column

                On Error Resume Next

                spName = Trim(CStr(ws.Cells(2, sprintCol.Column + 1).Value))

                If Err.Number <> 0 Or spName = "" Then spName = "Unknown" ' Assign default name if error occurs

                On Error GoTo 0

                If spName <> "" Then

                    originalName = spName

                    counter = 1

                    ' Ensure unique sheet name

                    Do While SheetExists(spName)

                        spName = originalName & "_" & counter

                        counter = counter + 1

                    Loop

                    ws.Name = spName

                Else

                    ws.Name = "Not Found"

                End If

            End If

            ' Find the "Status" column

            Set statusCol = ws.Rows(1).Find(What:="Status", LookIn:=xlValues, LookAt:=xlWhole)

            If Not statusCol Is Nothing Then

                statusColNum = statusCol.Column

                lastRow = ws.Cells(ws.Rows.Count, statusColNum).End(xlUp).Row

                ' Loop from bottom to top to delete rows with "Open"

                For i = lastRow To 2 Step -1

                    If Trim(LCase(ws.Cells(i, statusColNum).Value)) = "open" Then

                        ws.Rows(i).Delete

                    End If

                Next i

            End If

            ' Create a table

            Set rng = ws.UsedRange

            Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)

            tbl.Name = "JiraData_Table"

            tbl.TableStyle = "TableStyleLight8"

            ' Apply formatting to the first row

            With ws.Rows(1)

                .Interior.Color = RGB(0, 0, 0)

                .Font.Color = RGB(255, 255, 255)

                .Font.Size = 11

                .Font.Name = "Arial"

            End With

            ' Delete specified columns

            For Each colName In columnsToDelete

                Set found = ws.Rows(1).Find(What:=colName, LookIn:=xlValues, LookAt:=xlWhole)

                If Not found Is Nothing Then ws.Columns(found.Column).Delete

            Next colName

            ' AutoFit all columns and rows

            ws.Cells.EntireColumn.AutoFit

            ws.Cells.EntireRow.AutoFit

            ' Copy data to the combined sheet

            If firstSheet Then

                ws.UsedRange.Copy Destination:=combinedWs.Cells(nextRow, 1)

                firstSheet = False

            Else

                ws.UsedRange.Offset(1, 0).Resize(ws.UsedRange.Rows.Count - 1, ws.UsedRange.Columns.Count).Copy _

                    Destination:=combinedWs.Cells(nextRow, 1)

            End If

            nextRow = combinedWs.Cells(combinedWs.Rows.Count, 1).End(xlUp).Row + 1

        End If

    Next ws

  ' Add new columns in CombinedData

    lastCol = combinedWs.Cells(1, combinedWs.Columns.Count).End(xlToLeft).Column

    For i = 0 To UBound(newHeaders)

        combinedWs.Cells(1, lastCol + i + 1).Value = newHeaders(i)

    Next i

    ' Find relevant columns

    Set origEstCol = combinedWs.Rows(1).Find(What:="Original Estimate", LookIn:=xlValues, LookAt:=xlWhole)

    Set origEstHrsCol = combinedWs.Rows(1).Find(What:="Original Estima hrs", LookIn:=xlValues, LookAt:=xlWhole)

    Set remEstCol = combinedWs.Rows(1).Find(What:="Remaining Estimate", LookIn:=xlValues, LookAt:=xlWhole)

    Set remHrsCol = combinedWs.Rows(1).Find(What:="Remaining hrs", LookIn:=xlValues, LookAt:=xlWhole)

    Set timeSpentCol = combinedWs.Rows(1).Find(What:="Time Spent", LookIn:=xlValues, LookAt:=xlWhole)

    Set timeSpentHrsCol = combinedWs.Rows(1).Find(What:="Time Spent hrs", LookIn:=xlValues, LookAt:=xlWhole)

    ' Apply formulas if columns are found

    lastRow = combinedWs.Cells(combinedWs.Rows.Count, 1).End(xlUp).Row

    If Not origEstCol Is Nothing And Not origEstHrsCol Is Nothing Then

        origEstHrsColNum = origEstHrsCol.Column

        For i = 2 To lastRow

            combinedWs.Cells(i, origEstHrsColNum).Formula = "=" & combinedWs.Cells(i, origEstCol.Column).Address(False, False) & "/3600"

        Next i

    End If

    If Not remEstCol Is Nothing And Not remHrsCol Is Nothing Then

        remHrsColNum = remHrsCol.Column

        For i = 2 To lastRow

            combinedWs.Cells(i, remHrsColNum).Formula = "=" & combinedWs.Cells(i, remEstCol.Column).Address(False, False) & "/3600"

        Next i

    End If

    If Not timeSpentCol Is Nothing And Not timeSpentHrsCol Is Nothing Then

        timeSpentHrsColNum = timeSpentHrsCol.Column

        For i = 2 To lastRow

            combinedWs.Cells(i, timeSpentHrsColNum).Formula = "=" & combinedWs.Cells(i, timeSpentCol.Column).Address(False, False) & "/3600"

        Next i

    End If

    ' AutoFit all columns and rows in the combined sheet

    combinedWs.Cells.EntireColumn.AutoFit

    combinedWs.Cells.EntireRow.AutoFit

    MsgBox "All sheets adjusted, renamed, cleaned, and combined without 'Open' rows. Everything is AutoFit!"

End Sub

' Function to check if a sheet with a given name exists

Function SheetExists(sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = ThisWorkbook.Sheets(sheetName)

    On Error GoTo 0

    SheetExists = Not ws Is Nothing

End Function
```

with extra formulas
```vb
Sub CombineSheets_withCol()
    Dim ws As Worksheet
    Dim combinedWs As Worksheet
    Dim rng As Range
    Dim deleteRange As Range
    Dim sprintCol As Range
    Dim statusCol As Range
    Dim lastRow As Long
    Dim cell As Range
    Dim sprintColLetter As String
    Dim newColLetter As String
    Dim columnsToDelete As Variant
    Dim colName As Variant
    Dim found As Range
    Dim nextRow As Long
    Dim spName As String
    Dim tbl As ListObject
    Dim firstSheet As Boolean
    Dim statusColNum As Integer
    Dim originalName As String
    Dim counter As Integer
    Dim newHeaders As Variant
    Dim lastCol As Integer
    Dim i As Integer

    ' Define the columns to delete
    columnsToDelete = Array("Issue Links", "Fix Version/s", "ROI($)", "Updated", "Sprint History", "Sprint commitment", _
                            "Project Status (Date / Comments)", "Last Issue Comment", "Description", "Solution")

    ' Define new columns to add in CombinedData
    newHeaders = Array("Original Estima hrs", "Remaining hrs", "Time Spent hrs", "Original Estima SP", _
                       "Remaining SP", "Time Spent SP", "EV, SP", "WP")

    ' Create a new sheet for the combined data
    On Error Resume Next
    Set combinedWs = ThisWorkbook.Sheets("CombinedData")
    On Error GoTo 0

    If combinedWs Is Nothing Then
        Set combinedWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        combinedWs.Name = "CombinedData"
    Else
        combinedWs.Cells.Clear
    End If

    nextRow = 1
    firstSheet = True ' Flag to track the first sheet

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Instructions" And ws.Name <> "CombinedData" Then
            With ws
                On Error Resume Next
                .Shapes.Range(Array("Picture 1")).Delete
                On Error GoTo 0

                .Cells.UnMerge
                .Cells.Borders.LineStyle = xlNone ' Clear all borders
                .Rows("1:3").Delete ' Delete the first three rows

                ' Format the first row
                .Rows(1).Interior.Color = RGB(64, 64, 64)
                .Rows(1).Font.Color = RGB(255, 255, 255)

                ' Find and delete rows containing "Generated at"
                Set rng = ws.UsedRange.Find(What:="Generated at", LookIn:=xlValues, LookAt:=xlPart, _
                                            SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
                If Not rng Is Nothing Then
                    Set deleteRange = rng
                    Do
                        If deleteRange Is Nothing Then
                            Set deleteRange = rng
                        Else
                            Set deleteRange = Union(deleteRange, rng)
                        End If
                        Set rng = ws.UsedRange.FindNext(rng)
                    Loop While Not rng Is Nothing And rng.Address <> deleteRange.Cells(1, 1).Address
                    deleteRange.EntireRow.Delete
                End If
            End With

            ' Find the column with the header "Sprint"
            Set sprintCol = ws.Rows(1).Find(What:="Sprint", LookIn:=xlValues, LookAt:=xlWhole)

            If Not sprintCol Is Nothing Then
                sprintColLetter = Split(sprintCol.Address, "$")(1)

                ' Insert a new column after the Sprint column
                sprintCol.Offset(0, 1).EntireColumn.Insert Shift:=xlToRight
                newColLetter = Split(sprintCol.Offset(0, 1).Address, "$")(1)
                ws.Cells(1, sprintCol.Column + 1).Value = "SP#"

                ' Find the last row in the Sprint column
                lastRow = ws.Cells(ws.Rows.Count, sprintCol.Column).End(xlUp).Row

                ' Apply the formula to extract sprint number
                For Each cell In ws.Range(newColLetter & "2:" & newColLetter & lastRow)
                    cell.Formula = "=MID(" & sprintColLetter & cell.Row & ", FIND(""_"", " & sprintColLetter & cell.Row & ") + 1, " & _
                                   "FIND(""_"", " & sprintColLetter & cell.Row & ", FIND(""_"", " & sprintColLetter & cell.Row & ") + 1) - " & _
                                   "FIND(""_"", " & sprintColLetter & cell.Row & ") - 1)"
                Next cell

                ' Convert formulas to values
                For Each cell In ws.Range(newColLetter & "2:" & newColLetter & lastRow)
                    cell.Value = cell.Value
                Next cell

                ' Rename the sheet based on the value in the "SP#" column
                On Error Resume Next
                spName = Trim(CStr(ws.Cells(2, sprintCol.Column + 1).Value))
                If Err.Number <> 0 Or spName = "" Then spName = "Unknown"
                On Error GoTo 0

                If spName <> "" Then
                    originalName = spName
                    counter = 1

                    ' Ensure unique sheet name
                    Do While SheetExists(spName)
                        spName = originalName & "_" & counter
                        counter = counter + 1
                    Loop

                    ws.Name = spName
                Else
                    ws.Name = "Not Found"
                End If
            End If

            ' Find the "Status" column
            Set statusCol = ws.Rows(1).Find(What:="Status", LookIn:=xlValues, LookAt:=xlWhole)

            If Not statusCol Is Nothing Then
                statusColNum = statusCol.Column
                lastRow = ws.Cells(ws.Rows.Count, statusColNum).End(xlUp).Row

                ' Loop from bottom to top to delete rows with "Open"
                For i = lastRow To 2 Step -1
                    If Trim(LCase(ws.Cells(i, statusColNum).Value)) = "open" Then
                        ws.Rows(i).Delete
                    End If
                Next i
            End If

            ' Create a table
            Set rng = ws.UsedRange
            Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
            tbl.Name = "JiraData_Table"
            tbl.TableStyle = "TableStyleLight8"

            ' Apply formatting to the first row
            With ws.Rows(1)
                .Interior.Color = RGB(0, 0, 0)
                .Font.Color = RGB(255, 255, 255)
                .Font.Size = 11
                .Font.Name = "Arial"
            End With

            ' Delete specified columns
            For Each colName In columnsToDelete
                Set found = ws.Rows(1).Find(What:=colName, LookIn:=xlValues, LookAt:=xlWhole)
                If Not found Is Nothing Then ws.Columns(found.Column).Delete
            Next colName

            ' AutoFit all columns and rows
            ws.Cells.EntireColumn.AutoFit
            ws.Cells.EntireRow.AutoFit

            ' Copy data to the combined sheet
            If firstSheet Then
                ws.UsedRange.Copy Destination:=combinedWs.Cells(nextRow, 1)
                firstSheet = False
            Else
                ws.UsedRange.Offset(1, 0).Resize(ws.UsedRange.Rows.Count - 1, ws.UsedRange.Columns.Count).Copy _
                    Destination:=combinedWs.Cells(nextRow, 1)
            End If

            nextRow = combinedWs.Cells(combinedWs.Rows.Count, 1).End(xlUp).Row + 1
        End If
    Next ws

    ' AutoFit all columns and rows in the combined sheet
    combinedWs.Cells.EntireColumn.AutoFit
    combinedWs.Cells.EntireRow.AutoFit

    MsgBox "All sheets adjusted, renamed, cleaned, and combined without 'Open' rows. Everything is AutoFit!"
End Sub

' Function to check if a sheet with a given name exists
Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
```

```vb
Private Sub Workbook_NewSheet(ByVal Sh As Object)
    Dim ws As Worksheet
    Dim delayTime As Double
    Dim newSheetName As String
    
    ' Check if the new sheet name starts with "Details"
    If Left(Sh.Name, 7) = "Details" Then
        Set ws = Sh
        
        ' Rename the sheet to a custom name (change as needed)
        newSheetName = "DrillDown_Data_" & Format(Now, "hhmmss")
        On Error Resume Next
        ws.Name = newSheetName
        On Error GoTo 0
        
        ' Apply formatting (modify as needed)
        With ws.Cells
            .Font.Name = "Arial"
            .Font.Size = 10
            .Interior.ColorIndex = xlNone ' Remove background color
        End With
        
        ' Set delay time (e.g., delete after 10 seconds)
        delayTime = Now + TimeValue("00:00:10")
        
        ' Schedule deletion of the sheet
        Application.OnTime delayTime, "'DeleteSheet """ & newSheetName & """'"
    End If
End Sub

Public Sub DeleteSheet(sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    
    ' Delete the sheet if it still exists
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Sub
```